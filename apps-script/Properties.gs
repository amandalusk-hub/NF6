/**
 * Properties.gs — Mike's rental properties (Dorado, Condado, future) tracker.
 *
 * Goal: replace Amanda's manual monthly Excel report with an automated pipeline:
 *   1. PROPERTIES sheet holds one row per property with its config (linked
 *      Google Calendar for occupancy, linked Asset/Liability rows for value +
 *      mortgage, Alma management fee %, comma-list of Plaid Account IDs whose
 *      transactions belong to this property).
 *   2. PROPERTY_RULES sheet categorizes each Plaid transaction on those
 *      accounts into report line items (HOA Dues, Dorado Club Fees, Internet,
 *      Utilities-Electricity, Maintenance-Exterminator/AC/Cleaning, Rental
 *      Income from Alma, etc.). Same pattern as TLMND_RULES.
 *   3. PROPERTY_MANUAL sheet lets Amanda enter one-off overrides that the
 *      rules engine doesn't cover — e.g. insurance if it isn't in the feed,
 *      or reclassifying a guest deposit as a reimbursement instead of income.
 *   4. Google Calendar events are the reservation log — no manual reservation
 *      entry needed. Event title conventions distinguish revenue-generating
 *      bookings (Alma / direct guest) from non-revenue stays (family / owner).
 *   5. Monthly report data is stitched together on demand from all of the
 *      above and rendered on the Properties tab + emailed as a PDF on the
 *      20th of each month (covering the prior full month — by then Alma has
 *      settled all rents and every expense has cleared).
 *
 * This file (commit 1): data model + sheet lifecycle + Dorado seed + basic
 * CRUD. UI, rules engine, calendar integration, PDF, and email trigger land
 * in follow-on commits.
 */

// ── Sheet schemas ──────────────────────────────────────────────────────────

var PROPERTIES_HEADERS = [
  'ID',
  'Name',                     // e.g. "Dorado - 1405 Plantation Vlg"
  'Address',
  'City',
  'Country',
  'Currency',
  'Manager',                  // e.g. "Alma" — the booking / property manager
  'Manager Fee %',            // e.g. 20 — used to gross up net deposits back
                              // into full rental amounts on the report.
  'Google Calendar ID',       // The calendar whose events ARE the reservations.
                              // Format: xxxxxxx@group.calendar.google.com or
                              // 'primary' for the account's default calendar.
  'Linked Asset ID',          // The Assets sheet row for this property. Its
                              // value gets displayed on the Properties tab
                              // (and unchanged — property value is set on the
                              // Assets tab, not here).
  'Linked Liability ID',      // The Liabilities row for the property's mortgage.
                              // Its balance gets synced DOWN each month as the
                              // mortgage principal is paid.
  'Plaid Account IDs',        // Comma-separated Plaid Account IDs (or account
                              // name substrings, if IDs aren't known) whose
                              // TLMND_TRANSACTIONS rows belong to this
                              // property's report. e.g.
                              //   "oriental...1234,ath...5678,fidelity...9007"
  'Active',                   // Yes / No — inactive properties are hidden.
  'Notes',
  'Date Added',
  'Last Updated'
];

var PROPERTY_RULES_HEADERS = [
  'ID',
  'Property ID',              // Which property this rule scores against.
  'Priority',                 // Lower = evaluated first. Ties broken by rule ID.
  'Match Account',            // Substring match against the transaction's Account
                              // field. Empty = any account (for this property).
  'Match Name',               // Substring match against transaction Name/description.
                              // Empty = any name. Both Account + Name substrings
                              // are lowercased before comparison.
  'Match Amount',             // Exact match on the amount. e.g. "-1254.88" for
                              // Dorado HOA. Use pipe for multiple: "-140|-160"
                              // for Lourdes cleaning (either amount). Empty =
                              // any amount.
  'Direction',                // 'in' (positive only) | 'out' (negative only) |
                              // 'any'. Filters before Match Amount check so a
                              // negative Match Amount doesn't accidentally
                              // match a positive of the same magnitude.
  'Category',                 // Top-level report bucket: 'Revenue', 'Direct Cost',
                              // 'Operating Expense', 'Debt Service',
                              // 'Reimbursement' (not counted as income).
  'Subcategory',              // Report line item: 'Rental Income (Alma)',
                              // 'HOA Dues', 'Dorado Club Fees', 'Internet',
                              // 'Utilities - Electricity',
                              // 'Maintenance - Exterminator', etc. Matches the
                              // labels in Amanda's Excel report exactly so the
                              // PDF layout can group by this field.
  'Auto-Applied Count',       // Bump every time the rule matches something —
                              // helps spot dead rules and hot ones.
  'Last Applied',
  'Notes',
  'Active',                   // Yes / No — soft-disable without deleting.
  'Date Added'
];

var PROPERTY_MANUAL_HEADERS = [
  'ID',
  'Property ID',
  'Month',                    // 'YYYY-MM'. All manual entries roll into their
                              // month's report.
  'Date',                     // Optional specific date within the month.
  'Amount',
  'Category',                 // Same category vocabulary as PROPERTY_RULES.
  'Subcategory',
  'Notes',
  'Entered By',
  'Entered At'
];

// A separate sheet lets Amanda override a specific transaction's auto-derived
// category — e.g. a guest deposit that the rules called Rental Income but is
// actually a membership reimbursement. Keyed by Plaid transaction ID so it
// survives every re-run of the rules engine.
var PROPERTY_TXN_OVERRIDES_HEADERS = [
  'Transaction ID',           // Plaid txn ID from TLMND_TRANSACTIONS
  'Property ID',
  'Category',
  'Subcategory',
  'Notes',
  'Entered By',
  'Entered At'
];

// Recurring entries — fixed-amount monthly expenses (or revenue) that don't
// come through Plaid. Use this for things like Dorado's $21,668.72 mortgage
// payment where the amount is always the same and the account isn't
// connected via Plaid. The monthly report builder auto-includes a virtual
// transaction for every month in [Start Month, End Month] that falls within
// the report's month.
var PROPERTY_RECURRING_HEADERS = [
  'ID',
  'Property ID',
  'Label',                    // Human-readable: 'Dorado Mortgage (Oriental)'
  'Amount',                   // signed: + = inflow, − = outflow
  'Category',                 // Same vocab as PROPERTY_RULES
  'Subcategory',
  'Day of Month',             // 1-28 (clamped to end-of-month if needed)
  'Start Month',              // 'YYYY-MM' — first month this applies
  'End Month',                // 'YYYY-MM' — last month, blank = indefinite
  'Active',                   // Yes / No
  'Notes',
  'Date Added',
  'Last Updated'
];


// ── Sheet lifecycle ────────────────────────────────────────────────────────

function ensurePropertiesSheets_() {
  _ensurePropertySheet_('PROPERTIES', PROPERTIES_HEADERS);
  _ensurePropertySheet_('PROPERTY_RULES', PROPERTY_RULES_HEADERS);
  _ensurePropertySheet_('PROPERTY_MANUAL', PROPERTY_MANUAL_HEADERS);
  _ensurePropertySheet_('PROPERTY_TXN_OVERRIDES', PROPERTY_TXN_OVERRIDES_HEADERS);
  _ensurePropertySheet_('PROPERTY_RECURRING', PROPERTY_RECURRING_HEADERS);
}

function _ensurePropertySheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold').setBackground('#14263d').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
    return sheet;
  }
  // Auto-add any missing columns to survive schema drift across code updates.
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var missing = headers.filter(function(h) { return existing.indexOf(h) < 0; });
  if (missing.length) {
    var startCol = existing.length + 1;
    sheet.getRange(1, startCol, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold').setBackground('#14263d').setFontColor('#ffffff');
  }
  return sheet;
}


// ── Header-name row read/write helpers ─────────────────────────────────────
// Same pattern as Loans.gs — reading + writing by header NAME insulates us
// from column position drift, which cost us a day on the Loans launch.

function _getPropertySheetRows_(sheetName, headers) {
  ensurePropertiesSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), headers.length)).getValues();
  var hdr = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h) {
      var idx = hdr.indexOf(h);
      obj[h] = idx >= 0 ? row[idx] : '';
    });
    return obj;
  }).filter(function(o) { return o.ID; });
}

function _writePropertyRow_(sheetName, headers, obj) {
  ensurePropertiesSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var lastCol = Math.max(sheet.getLastColumn(), headers.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = hdr.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.appendRow(row);
  // Google Sheets auto-converts anything that LOOKS like a date (e.g.
  // "2024-01" for a month key) to a Date object. For known text-only fields
  // on this schema, force the newly-written cell to text format so it stays
  // a string. The read side still handles Dates defensively (_mkKey) but
  // this prevents the drift for future writes.
  var textColumns = ['Start Month', 'End Month', 'Month', 'Plaid Account IDs'];
  var newRowNum = sheet.getLastRow();
  textColumns.forEach(function(colName) {
    var idx = hdr.indexOf(colName);
    if (idx >= 0 && obj[colName] !== undefined && obj[colName] !== '') {
      var cell = sheet.getRange(newRowNum, idx + 1);
      cell.setNumberFormat('@');
      cell.setValue(String(obj[colName]));
    }
  });
}

function _updatePropertyRow_(sheetName, headers, id, patch) {
  ensurePropertiesSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet.getLastRow() < 2) return false;
  var lastCol = Math.max(sheet.getLastColumn(), headers.length);
  var data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  var hdr = data[0];
  var iId = hdr.indexOf('ID');
  if (iId < 0) return false;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iId]) !== String(id)) continue;
    Object.keys(patch).forEach(function(k) {
      var idx = hdr.indexOf(k);
      if (idx >= 0) sheet.getRange(r + 1, idx + 1).setValue(patch[k]);
    });
    return true;
  }
  return false;
}


// ── Web-callable CRUD ──────────────────────────────────────────────────────

// Read: list every active property. Used by the Properties tab card view and
// by the monthly report generator.
function getProperties() {
  var rows = _getPropertySheetRows_('PROPERTIES', PROPERTIES_HEADERS);
  return rows.filter(function(p) {
    var a = String(p['Active'] || '').toLowerCase();
    return a !== 'no' && a !== 'false' && a !== '0';
  });
}

function getProperty(id) {
  var rows = _getPropertySheetRows_('PROPERTIES', PROPERTIES_HEADERS);
  return rows.find(function(p) { return String(p['ID']) === String(id); }) || null;
}

function addProperty(data) {
  _requireEditor_();
  var id = 'prop_' + Utilities.getUuid().substring(0, 8);
  var now = new Date();
  _writePropertyRow_('PROPERTIES', PROPERTIES_HEADERS, {
    'ID': id,
    'Name':                String(data.name || '').trim(),
    'Address':             String(data.address || '').trim(),
    'City':                String(data.city || '').trim(),
    'Country':             String(data.country || '').trim(),
    'Currency':            String(data.currency || 'USD'),
    'Manager':             String(data.manager || '').trim(),
    'Manager Fee %':       Number(data.managerFeePct) || 0,
    'Google Calendar ID':  String(data.googleCalendarId || '').trim(),
    'Linked Asset ID':     String(data.linkedAssetId || '').trim(),
    'Linked Liability ID': String(data.linkedLiabilityId || '').trim(),
    'Plaid Account IDs':   String(data.plaidAccountIds || '').trim(),
    'Active':              data.active === false ? 'No' : 'Yes',
    'Notes':               String(data.notes || '').trim(),
    'Date Added':          now,
    'Last Updated':        now
  });
  _logAudit_('addProperty', 'property', id, data.name, 'Added property: ' + data.name);
  return { success: true, id: id };
}

function updateProperty(id, patch) {
  _requireEditor_();
  var writePatch = {};
  var map = {
    name: 'Name', address: 'Address', city: 'City', country: 'Country',
    currency: 'Currency', manager: 'Manager', managerFeePct: 'Manager Fee %',
    googleCalendarId: 'Google Calendar ID', linkedAssetId: 'Linked Asset ID',
    linkedLiabilityId: 'Linked Liability ID', plaidAccountIds: 'Plaid Account IDs',
    notes: 'Notes'
  };
  Object.keys(map).forEach(function(k) {
    if (patch[k] !== undefined) writePatch[map[k]] = patch[k];
  });
  if (patch.active !== undefined) writePatch['Active'] = patch.active ? 'Yes' : 'No';
  writePatch['Last Updated'] = new Date();
  var ok = _updatePropertyRow_('PROPERTIES', PROPERTIES_HEADERS, id, writePatch);
  _logAudit_('updateProperty', 'property', id, patch.name || '', 'Updated property');
  return { success: ok };
}

function getPropertyRules(propertyId) {
  var rows = _getPropertySheetRows_('PROPERTY_RULES', PROPERTY_RULES_HEADERS);
  return rows.filter(function(r) {
    if (propertyId && String(r['Property ID']) !== String(propertyId)) return false;
    var a = String(r['Active'] || '').toLowerCase();
    return a !== 'no' && a !== 'false' && a !== '0';
  }).sort(function(a, b) {
    return (Number(a['Priority']) || 999) - (Number(b['Priority']) || 999);
  });
}

function addPropertyRule(data) {
  _requireEditor_();
  var id = 'prule_' + Utilities.getUuid().substring(0, 8);
  _writePropertyRow_('PROPERTY_RULES', PROPERTY_RULES_HEADERS, {
    'ID':               id,
    'Property ID':      String(data.propertyId || ''),
    'Priority':         Number(data.priority) || 100,
    'Match Account':    String(data.matchAccount || '').trim(),
    'Match Name':       String(data.matchName || '').trim(),
    'Match Amount':     String(data.matchAmount || '').trim(),
    'Direction':        String(data.direction || 'any'),
    'Category':         String(data.category || ''),
    'Subcategory':      String(data.subcategory || ''),
    'Auto-Applied Count': 0,
    'Last Applied':     '',
    'Notes':            String(data.notes || '').trim(),
    'Active':           data.active === false ? 'No' : 'Yes',
    'Date Added':       new Date()
  });
  return { success: true, id: id };
}


// ── Seed: Dorado ───────────────────────────────────────────────────────────
// Menu-callable one-shot: creates the Dorado property row if it doesn't exist
// yet, wires in the calendar ID Amanda provided, and installs the rules that
// carry over from her Excel report + email description. Idempotent — running
// it twice won't create duplicates. Linked Asset ID + Linked Liability ID are
// left blank on the first seed; Amanda picks them in the Property modal's
// dropdowns (same UX as the Loans Linked Asset picker) so we never expose raw
// sheet IDs.

var DORADO_CALENDAR_ID = 'c_fd70b80d64a3985cdd035b5451d575377a5e63c9dafa5f86bfb6e878a276644a@group.calendar.google.com';

function seedDoradoProperty() {
  _requireEditor_();
  ensurePropertiesSheets_();

  // 1. Property row — create only if not already there (match by name substring).
  var existing = getProperties().find(function(p) {
    return /dorado/i.test(String(p['Name'] || ''));
  });
  var propId;
  if (existing) {
    propId = existing['ID'];
    // Refresh calendar ID + manager info in case Amanda edited by hand.
    updateProperty(propId, {
      googleCalendarId: DORADO_CALENDAR_ID,
      manager: 'Alma',
      managerFeePct: 20,
      city: 'Dorado',
      country: 'Puerto Rico',
      currency: 'USD'
    });
  } else {
    var r = addProperty({
      name:              'Dorado - 1405 Plantation Vlg',
      address:           '1405 Plantation Vlg, Dorado, PR 00646',
      city:              'Dorado',
      country:           'Puerto Rico',
      currency:          'USD',
      manager:           'Alma',
      managerFeePct:     20,
      googleCalendarId:  DORADO_CALENDAR_ID,
      linkedAssetId:     '',   // Amanda picks in the modal
      linkedLiabilityId: '',
      plaidAccountIds:   '',   // will be filled once we identify the Plaid
                               // accounts (Oriental Bank, ATH, Fidelity 9007)
      active:            true,
      notes:             'Short-term rental managed by Alma. Family + Mike + friends also stay. Report cadence: monthly on the 20th, covering prior full month.'
    });
    propId = r.id;
  }

  // 2. Rules — install the known patterns. Deduped by (Property ID + Match
  //    Account + Match Amount + Match Name + Subcategory) so re-running is a
  //    no-op after the first seed.
  var existingRules = getPropertyRules(propId);
  var ruleKey = function(r) {
    return [r.matchAccount||'', r.matchName||'', r.matchAmount||'', r.subcategory||''].join('|').toLowerCase();
  };
  var have = {};
  existingRules.forEach(function(r) {
    have[ruleKey({
      matchAccount: r['Match Account'], matchName: r['Match Name'],
      matchAmount:  r['Match Amount'],  subcategory: r['Subcategory']
    })] = true;
  });

  var DORADO_RULES = [
    // ── EXCLUSIONS (highest priority — evaluated first) ─────────────────
    // Oriental auto-moves money between Mike's Savings and Checking
    // whenever a card charge needs funds. These show up as paired entries
    // on both accounts with the same amount — if we let them fall through
    // to Needs Review (or worse, let them match an exact-amount rule like
    // HOA $1,254.88), the report double-counts. Catch them FIRST and tag
    // as Internal Transfer, which gets ignored from report totals.
    { priority: 1, matchAccount:'oriental', matchName:'automatic transfer', matchAmount:'', direction:'any',
      category:'Internal Transfer', subcategory:'Savings ↔ Checking',
      notes:'Oriental auto-shuffle between Mike\'s own accounts. Ignored from totals.' },
    // Mastercard authorizations are temporary holds that get reversed when
    // the real purchase posts. Same story — ignore them so they don\'t
    // appear as pending income or expense in the report.
    { priority: 1, matchAccount:'oriental', matchName:'mc authorization', matchAmount:'', direction:'any',
      category:'Internal Transfer', subcategory:'MC Auth (reversed by MC PURCHASE)',
      notes:'Pending card hold — reverses when actual purchase posts.' },

    // ── REVENUE ─────────────────────────────────────────────────────────
    // Rental income from Alma — arrives via Oriental Bank as a positive
    // deposit. The manager fee % on the property row is used to gross the
    // deposit up to the pre-fee rental amount on the report.
    { priority: 10, matchAccount:'oriental', matchName:'alma',   matchAmount:'', direction:'in',
      category:'Revenue', subcategory:'Rental Income (Alma)',
      notes:'Net deposit (after 20% Alma fee). Gross up on report.' },

    // NO auto-rule for direct guest wires. A "WIRE IN" deposit from a named
    // guest to account 6179 could be: rental payment for a current stay,
    // deposit for a future stay, membership fees being prepaid, spending
    // money the guest asks us to hold, etc. Only Amanda can tell which.
    // Those wires stay in Needs Review and she categorizes each manually
    // via a PROPERTY_TXN_OVERRIDES entry — the paired $15 service charge
    // stays unmatched too since its destination (Direct Cost vs
    // Reimbursement cost) depends on what the wire itself was for.

    // HOA — fixed monthly amount, matches by exact amount.
    { priority: 20, matchAccount:'oriental', matchName:'', matchAmount:'-1254.88', direction:'out',
      category:'Operating Expense', subcategory:'HOA Dues' },

    // Dorado Beach Resort statement — bundles Club Dues ($1,112.22) +
    // Mike's food charges at the club. Payment amount varies month-to-month
    // depending on food charges. Rule: hide the full Plaid payment from the
    // report ("Internal Transfer" = ignored), and let the PROPERTY_RECURRING
    // entry supply the clean $1,112.22 Dorado Club Fees every month. Food
    // charges stay off the Dorado report entirely (Mike personal).
    { priority: 15, matchAccount:'6179', matchName:'dbr dorado', matchAmount:'', direction:'out',
      category:'Internal Transfer', subcategory:'DBR Dorado statement (food + dues bundled; dues covered by recurring)',
      notes:'Hide the raw payment — PROPERTY_RECURRING adds the clean $1,112.22 dues line.' },

    // Utilities — by biller name on Oriental.
    { priority: 30, matchAccount:'oriental', matchName:'claro', matchAmount:'', direction:'out',
      category:'Operating Expense', subcategory:'Internet' },
    { priority: 30, matchAccount:'oriental', matchName:'luma',  matchAmount:'', direction:'out',
      category:'Operating Expense', subcategory:'Utilities - Electricity' },

    // ATH mobile — Oriental's A2A mobile payment to vendors. The ATH
    // identifier is in the transaction NAME ("A2A PMT DEBIT|...ATH MOVIL
    // PHONE SAN JUAN"), NOT the account label. Each vendor's amount is a
    // standard round number, so we identify by (name = "ath movil") + amount.
    { priority: 40, matchAccount:'0451', matchName:'ath movil', matchAmount:'-60.00',  direction:'out',
      category:'Operating Expense', subcategory:'Maintenance - Exterminator' },
    { priority: 40, matchAccount:'0451', matchName:'ath movil', matchAmount:'-295.00', direction:'out',
      category:'Operating Expense', subcategory:'Maintenance - AC Quarterly' },
    { priority: 40, matchAccount:'0451', matchName:'ath movil', matchAmount:'-140.00|-160.00', direction:'out',
      category:'Direct Cost', subcategory:'Cleaning (Lourdes)' },

    // Mortgage — paid out of Fidelity 9007 (not currently in Plaid, so this
    // rule won't fire; the PROPERTY_RECURRING entry handles the mortgage
    // line until Amanda connects Fidelity 9007 via Plaid).
    { priority: 50, matchAccount:'9007', matchName:'', matchAmount:'-21668.72', direction:'out',
      category:'Debt Service', subcategory:'Mortgage' }
  ];

  var added = 0;
  DORADO_RULES.forEach(function(rule) {
    var key = ruleKey(rule);
    if (have[key]) return;
    addPropertyRule(Object.assign({ propertyId: propId }, rule));
    added++;
  });

  return {
    success: true,
    propertyId: propId,
    rulesInstalled: added,
    rulesAlreadyPresent: DORADO_RULES.length - added,
    calendarId: DORADO_CALENDAR_ID,
    nextSteps: [
      'Open the Properties tab',
      'Open the Dorado card and pick the Linked Asset (1405 Plantation Vlg) and Linked Liability (mortgage) from the dropdowns',
      'Add the Plaid Account IDs / substrings once we identify Oriental, ATH, and Fidelity 9007 in the feed'
    ]
  };
}

// ── Categorization engine ──────────────────────────────────────────────────
// Given a Plaid transaction and a set of PROPERTY_RULES, find the highest-
// priority rule that matches and return its category/subcategory. If no rule
// matches, the transaction lands in the "Needs Review" queue in the UI so
// Amanda can either categorize it once (via override) or add a rule so all
// future ones self-categorize.

// Amount matcher: rule field can be
//   ""              → matches any amount
//   "-1254.88"      → exact (within 1 cent, to absorb rounding)
//   "-140|-160"     → any of these values
// Returns true if the transaction amount matches the rule.
function _matchAmount_(txAmount, ruleAmount) {
  var s = String(ruleAmount || '').trim();
  if (!s) return true;
  var candidates = s.split('|').map(function(v){ return Number(String(v).trim()); }).filter(function(n){ return !isNaN(n); });
  if (!candidates.length) return true;
  return candidates.some(function(c) { return Math.abs(txAmount - c) < 0.01; });
}

// Substring matcher: rule field can be
//   ""            → matches any
//   "oriental"    → substring match, case-insensitive
// Multiple pipe-separated substrings are all treated as OR.
function _matchSubstring_(txValue, ruleValue) {
  var s = String(ruleValue || '').toLowerCase().trim();
  if (!s) return true;
  var candidates = s.split('|').map(function(v){ return v.trim(); }).filter(Boolean);
  var haystack = String(txValue || '').toLowerCase();
  return candidates.some(function(c) { return haystack.indexOf(c) >= 0; });
}

// Direction filter — 'in' requires positive, 'out' requires negative, 'any'
// or unset accepts either. Runs BEFORE amount check so a rule tagged -1254
// won't accidentally match a +1254 refund.
function _matchDirection_(txAmount, ruleDirection) {
  var d = String(ruleDirection || 'any').toLowerCase();
  if (d === 'in')  return txAmount > 0;
  if (d === 'out') return txAmount < 0;
  return true;
}

// Try each rule in priority order (already sorted by getPropertyRules).
// Returns { category, subcategory, ruleId } or null if nothing matched.
function _categorizeTxn_(txn, rules) {
  for (var i = 0; i < rules.length; i++) {
    var r = rules[i];
    if (!_matchSubstring_(txn.account, r['Match Account'])) continue;
    if (!_matchSubstring_(txn.name,    r['Match Name']))    continue;
    if (!_matchDirection_(txn.amount,  r['Direction']))     continue;
    if (!_matchAmount_(txn.amount,     r['Match Amount']))  continue;
    return {
      category:    String(r['Category'] || ''),
      subcategory: String(r['Subcategory'] || ''),
      ruleId:      r['ID']
    };
  }
  return null;
}

// Read every transaction belonging to this property. Scans BOTH TLMND_
// TRANSACTIONS (the user-editable TLMND feed) AND PLAID_TRANSACTIONS (the
// full Plaid feed with every connected account), filtering to rows whose
// Account column contains ANY of the substrings listed in the property's
// "Plaid Account IDs" field. Dedupes across the two sheets by Transaction
// ID so the same txn in both sources isn't counted twice.
function _getPropertyTransactions_(property, startDate, endDate) {
  var accountFilters = String(property['Plaid Account IDs'] || '')
    .split(',').map(function(s){ return s.trim().toLowerCase(); }).filter(Boolean);
  if (!accountFilters.length) return [];

  var seen = {};
  var out  = [];
  var startMs = startDate ? startDate.getTime() : -Infinity;
  var endMs   = endDate   ? endDate.getTime()   :  Infinity;

  ['TLMND_TRANSACTIONS', 'PLAID_TRANSACTIONS'].forEach(function(sheetName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var iId      = headers.indexOf('Transaction ID');
    var iDate    = headers.indexOf('Date');
    var iAccount = headers.indexOf('Account');
    var iName    = headers.indexOf('Name');
    var iMerch   = headers.indexOf('Merchant');
    var iAmount  = headers.indexOf('Amount USD');
    var iPending = headers.indexOf('Pending');
    if (iId < 0 || iDate < 0 || iAccount < 0 || iName < 0 || iAmount < 0) return;

    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    rows.forEach(function(r) {
      var acct = String(r[iAccount] || '').toLowerCase();
      if (!accountFilters.some(function(f) { return acct.indexOf(f) >= 0; })) return;
      var d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]);
      if (isNaN(d.getTime())) return;
      var t = d.getTime();
      if (t < startMs || t > endMs) return;
      if (iPending >= 0 && String(r[iPending] || '').toLowerCase() === 'yes') return;
      var id = String(r[iId] || '');
      if (!id || seen[id]) return;
      seen[id] = true;
      out.push({
        id:       id,
        date:     d,
        dateIso:  Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd'),
        account:  String(r[iAccount] || ''),
        name:     String(r[iName] || ''),
        merchant: iMerch >= 0 ? String(r[iMerch] || '') : '',
        amount:   Number(r[iAmount] || 0)
      });
    });
  });
  return out;
}

// Per-transaction overrides let Amanda promote a specific txn to a different
// category (e.g. reclassify one guest deposit from Rental Income to
// Reimbursement) without adding a rule that would blanket-affect similar txns.
function _getPropertyTxnOverrides_(propertyId) {
  var rows = _getPropertySheetRows_('PROPERTY_TXN_OVERRIDES', PROPERTY_TXN_OVERRIDES_HEADERS);
  var map = {};
  rows.forEach(function(r) {
    if (String(r['Property ID']) !== String(propertyId)) return;
    map[String(r['Transaction ID'])] = {
      category:    String(r['Category'] || ''),
      subcategory: String(r['Subcategory'] || '')
    };
  });
  return map;
}

// Web-callable: set a per-transaction override. Overwrites any prior override
// for the same (txn, property) pair. Amanda calls this from the "Needs Review"
// queue in the UI when she wants a one-off classification without a rule.
function setPropertyTxnOverride(txnId, propertyId, category, subcategory, notes) {
  _requireEditor_();
  ensurePropertiesSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PROPERTY_TXN_OVERRIDES');
  var lastCol = Math.max(sheet.getLastColumn(), PROPERTY_TXN_OVERRIDES_HEADERS.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var iTxn  = hdr.indexOf('Transaction ID');
  var iProp = hdr.indexOf('Property ID');
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    for (var r = 0; r < data.length; r++) {
      if (String(data[r][iTxn]) === String(txnId) && String(data[r][iProp]) === String(propertyId)) {
        var iCat = hdr.indexOf('Category'), iSub = hdr.indexOf('Subcategory');
        var iNotes = hdr.indexOf('Notes'), iEnt = hdr.indexOf('Entered At');
        sheet.getRange(r + 2, iCat + 1).setValue(category || '');
        sheet.getRange(r + 2, iSub + 1).setValue(subcategory || '');
        if (iNotes >= 0) sheet.getRange(r + 2, iNotes + 1).setValue(notes || '');
        if (iEnt >= 0)   sheet.getRange(r + 2, iEnt + 1).setValue(new Date());
        return { success: true, updated: true };
      }
    }
  }
  _writePropertyRow_('PROPERTY_TXN_OVERRIDES', PROPERTY_TXN_OVERRIDES_HEADERS, {
    'Transaction ID': String(txnId),
    'Property ID':    String(propertyId),
    'Category':       category || '',
    'Subcategory':    subcategory || '',
    'Notes':          notes || '',
    'Entered By':     _currentUserEmail_(),
    'Entered At':     new Date()
  });
  return { success: true, updated: false };
}


// ── Google Calendar reader ─────────────────────────────────────────────────
// Read the property's linked Google Calendar to figure out (a) upcoming
// reservations for the "next 60 days" widget, and (b) nights booked for
// the occupancy % on the monthly report. Event source is inferred from
// the event's title so Alma / direct-guest bookings (revenue) are
// distinguished from family / owner stays (occupancy only, no revenue).

// Case-insensitive keywords used to tag an event's source. Amanda can add
// aliases here or in the event titles — keep this list broad so a naming
// slip doesn't wreck the classification.
var _PROPERTY_EVENT_SOURCES = [
  { source: 'owner',  keywords: ['owner', 'mike', 'dr mike', 'dr. mike',
                                 'michael', 'dr michael', 'dr. michael',
                                 'michael nguyen', 'nguyen'] },
  { source: 'family', keywords: ['family', 'brother', 'sister', 'parents',
                                'mom', 'dad', 'kids', 'personal'] },
  { source: 'friend', keywords: ['friend', 'friends', 'guest of mike'] },
  { source: 'alma',   keywords: ['alma', 'airbnb', 'vrbo', 'booking.com', 'expedia'] },
  { source: 'direct', keywords: ['direct'] }
];

function _classifyEventSource_(title) {
  var t = String(title || '').toLowerCase();
  for (var i = 0; i < _PROPERTY_EVENT_SOURCES.length; i++) {
    var s = _PROPERTY_EVENT_SOURCES[i];
    for (var j = 0; j < s.keywords.length; j++) {
      if (t.indexOf(s.keywords[j]) >= 0) return s.source;
    }
  }
  return 'unknown';   // untagged event — Amanda can leave it as-is or tweak
                      // the title to include a source keyword.
}

// Return every event on the calendar between startDate (inclusive) and
// endDate (exclusive), FILTERED to just the ones for this property.
// Amanda's shared calendar carries events for every property (Dorado,
// Condado, Paris, Quebrada Arriba) with the property name as the first
// word of the event title — so a Dorado event looks like
// "Dorado - Dr Michael Nguyen". We filter to events whose title starts
// with the property's name prefix (first word of the PROPERTIES.Name
// field) so a Dorado property report doesn't pull in Condado stays.
function getPropertyReservations(propertyId, startIso, endIso) {
  var prop = getProperty(propertyId);
  if (!prop) return { error: 'Property not found: ' + propertyId, events: [] };
  var calId = String(prop['Google Calendar ID'] || '').trim();
  if (!calId) return { error: 'No Google Calendar ID set on property.', events: [] };

  // Prefix = first word of the property Name, lowercased. e.g. a property
  // named "Dorado - 1405 Plantation Vlg" gets prefix "dorado" and only
  // events like "Dorado - ..." in the shared calendar count toward its
  // report.
  var prefix = String(prop['Name'] || '').trim().split(/[\s\-]+/)[0].toLowerCase();

  var start = startIso ? new Date(startIso) : new Date();
  var end   = endIso   ? new Date(endIso)   : new Date(start.getTime() + 60 * 86400000);

  var cal;
  try { cal = CalendarApp.getCalendarById(calId); }
  catch(e) { return { error: 'CalendarApp.getCalendarById failed: ' + e.message, events: [] }; }
  if (!cal) return { error: 'Calendar not accessible. Make sure ' + _currentUserEmail_() + ' has at least read access.', events: [] };

  var events;
  try { events = cal.getEvents(start, end); }
  catch(e) { return { error: 'cal.getEvents failed: ' + e.message, events: [] }; }

  // Keep only events whose title starts with (or contains near the start)
  // the property's name prefix. Skip the rest.
  if (prefix) {
    events = events.filter(function(e) {
      var t = String(e.getTitle() || '').toLowerCase().trim();
      return t.indexOf(prefix) === 0 || t.indexOf(prefix + ' ') >= 0 || t.indexOf(prefix + '-') >= 0;
    });
  }

  var out = events.map(function(e) {
    var s = e.getStartTime();
    var f = e.getEndTime();
    var nights = Math.max(0, Math.round((f.getTime() - s.getTime()) / 86400000));
    if (e.isAllDayEvent()) {
      // getAllDayStartDate is timezone-safe; getStartTime for all-day events
      // may skew by a day depending on script timezone.
      s = e.getAllDayStartDate();
      f = e.getAllDayEndDate();
      nights = Math.max(0, Math.round((f.getTime() - s.getTime()) / 86400000));
    }
    return {
      id:         e.getId(),
      title:      e.getTitle(),
      start:      s.toISOString(),
      end:        f.toISOString(),
      startIso:   Utilities.formatDate(s, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      endIso:     Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      nights:     nights,
      source:     _classifyEventSource_(e.getTitle()),
      description: (e.getDescription() || '').substring(0, 500)
    };
  });
  return { propertyId: propertyId, events: out };
}

// Compute occupancy stats for a given month by clipping every reservation
// that overlaps the month to that month's window. Nights are summed with
// zero double-counting when reservations don't overlap; if two events DO
// overlap (shouldn't happen on a real calendar, but…) it still won't count
// the same day twice — we build a set of booked date strings.
function _computeMonthOccupancy_(events, year, month) {
  var monthStart = new Date(Date.UTC(year, month - 1, 1));
  var monthEnd   = new Date(Date.UTC(year, month, 1));   // exclusive
  var daysInMonth = new Date(year, month, 0).getDate();

  // Only paying-guest events count toward occupancy, nights, and reservation
  // count — Amanda's rule: "when it has Mike then it's not a reservation."
  // Owner / family / friend stays are the property being used but not
  // generating revenue, and she doesn't track them as reservations. If she
  // ever wants to show owner usage separately, we can add an ownerNights
  // field — but for now, owner + family + friend are completely excluded.
  function isCounted(src) {
    return src === 'alma' || src === 'direct' || src === 'unknown';
  }

  var bookedDates = {};
  var reservations = 0;
  events.forEach(function(e) {
    if (!isCounted(e.source)) return;   // owner / family / friend — skip
    var s = new Date(e.start);
    var f = new Date(e.end);
    if (f <= monthStart || s >= monthEnd) return;
    reservations++;
    var cur = s > monthStart ? new Date(s) : new Date(monthStart);
    var stop = f < monthEnd ? new Date(f) : new Date(monthEnd);
    while (cur < stop) {
      var iso = Utilities.formatDate(cur, 'UTC', 'yyyy-MM-dd');
      bookedDates[iso] = true;
      cur.setUTCDate(cur.getUTCDate() + 1);
    }
  });
  var nightsBooked = Object.keys(bookedDates).length;
  return {
    nightsBooked:  nightsBooked,
    revenueNights: nightsBooked,   // same thing now that owner is excluded
    reservations:  reservations,
    daysInMonth:   daysInMonth,
    occupancyPct:  daysInMonth > 0 ? Math.round(nightsBooked / daysInMonth * 100) : 0
  };
}


// ── Monthly report builder ─────────────────────────────────────────────────
// Web-callable — for a property + year + month, returns the FULL data
// structure the Properties tab (and later, the PDF) render from. Stitches
// together bank transactions (categorized via rules + overrides), manual
// entries, calendar reservations, and the linked liability's mortgage
// balance. Never writes anywhere — safe to call repeatedly from the UI.

function getPropertyMonthlyReport(propertyId, year, month) {
  var prop = getProperty(propertyId);
  if (!prop) return { error: 'Property not found: ' + propertyId };

  year  = Number(year)  || new Date().getFullYear();
  month = Number(month) || (new Date().getMonth() + 1);
  var monthStart = new Date(Date.UTC(year, month - 1, 1));
  var monthEnd   = new Date(Date.UTC(year, month, 1));   // exclusive
  var monthKey   = Utilities.formatDate(monthStart, 'UTC', 'yyyy-MM');

  // 1. Bank transactions ↔ rules.
  var rules = getPropertyRules(propertyId);
  var txns  = _getPropertyTransactions_(prop, monthStart, new Date(monthEnd.getTime() - 1));
  var overrides = _getPropertyTxnOverrides_(propertyId);
  var categorized = [];
  var needsReview = [];
  txns.forEach(function(t) {
    var ovr = overrides[t.id];
    if (ovr && (ovr.category || ovr.subcategory)) {
      categorized.push(Object.assign({}, t, { category: ovr.category, subcategory: ovr.subcategory, source: 'override' }));
      return;
    }
    var hit = _categorizeTxn_(t, rules);
    if (hit) {
      categorized.push(Object.assign({}, t, { category: hit.category, subcategory: hit.subcategory, ruleId: hit.ruleId, source: 'rule' }));
    } else {
      needsReview.push(t);
    }
  });

  // 2a. Recurring entries that fall in this month (static monthly expenses
  //     like the Dorado mortgage — amount is always the same, account isn't
  //     in Plaid, so we materialize a virtual transaction).
  // Normalize Start/End Month cells to 'YYYY-MM' strings regardless of
  // whether Google Sheets stored them as text or auto-converted to Date.
  function _mkKey(v) {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, 'UTC', 'yyyy-MM');
    return String(v).trim();
  }
  var recurrings = _getPropertySheetRows_('PROPERTY_RECURRING', PROPERTY_RECURRING_HEADERS)
    .filter(function(r) {
      if (String(r['Property ID']) !== String(propertyId)) return false;
      var active = String(r['Active'] || 'Yes').toLowerCase();
      if (active === 'no' || active === 'false') return false;
      var sm = _mkKey(r['Start Month']);
      var em = _mkKey(r['End Month']);
      if (sm && sm > monthKey) return false;
      if (em && em < monthKey) return false;
      return true;
    })
    .map(function(r) {
      var day = Math.min(28, Math.max(1, Number(r['Day of Month']) || 1));
      var d = new Date(Date.UTC(year, month - 1, day));
      return {
        id:          'recurring:' + r['ID'] + ':' + monthKey,
        date:        d,
        dateIso:     Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd'),
        account:     '(recurring)',
        name:        String(r['Label'] || ''),
        amount:      Number(r['Amount']) || 0,
        category:    String(r['Category'] || ''),
        subcategory: String(r['Subcategory'] || ''),
        source:      'recurring'
      };
    });

  // 2b. Manual one-off entries for this month.
  var manuals = _getPropertySheetRows_('PROPERTY_MANUAL', PROPERTY_MANUAL_HEADERS)
    .filter(function(m) { return String(m['Property ID']) === String(propertyId) && String(m['Month']) === monthKey; })
    .map(function(m) {
      return {
        id:          'manual:' + m['ID'],
        date:        m['Date'] instanceof Date ? m['Date'] : (m['Date'] ? new Date(m['Date']) : monthStart),
        account:     '(manual)',
        name:        String(m['Notes'] || ''),
        amount:      Number(m['Amount']) || 0,
        category:    String(m['Category'] || ''),
        subcategory: String(m['Subcategory'] || ''),
        source:      'manual'
      };
    });

  var lines = categorized.concat(recurrings).concat(manuals);

  // 3. Aggregate by category → subcategory. Preserves order of appearance
  //    within a category so the PDF renders line items in the same order
  //    Amanda's Excel does. Signed amounts are kept negative for out; the
  //    UI + PDF layer decides display formatting.
  var buckets = {};
  var subOrder = {};   // per-category insertion order
  lines.forEach(function(l) {
    var c = l.category || 'Uncategorized';
    var s = l.subcategory || '(no subcategory)';
    buckets[c] = buckets[c] || {};
    buckets[c][s] = buckets[c][s] || { total: 0, transactions: [] };
    buckets[c][s].total += l.amount;
    buckets[c][s].transactions.push({
      id: l.id, date: l.dateIso || (l.date && l.date.toISOString ? l.date.toISOString().substring(0,10) : ''),
      account: l.account, name: l.name, amount: l.amount, source: l.source
    });
    subOrder[c] = subOrder[c] || [];
    if (subOrder[c].indexOf(s) < 0) subOrder[c].push(s);
  });

  // 4. Rental income gross-up.
  //    Amanda receives NET (after Alma's fee %) but her Excel reports GROSS.
  //    Compute the implied Alma fee and expose both numbers.
  var feePct = Number(prop['Manager Fee %']) || 0;
  var almaBucket = buckets['Revenue'] && buckets['Revenue']['Rental Income (Alma)'];
  var netRoomRevenue = almaBucket ? almaBucket.total : 0;
  var grossRentalAmount = feePct > 0 && feePct < 100
    ? netRoomRevenue / (1 - feePct / 100)
    : netRoomRevenue;
  var almaFee = grossRentalAmount - netRoomRevenue;   // negative sign convention

  // Non-Alma revenue rows (direct guest, other) count at face value.
  var otherRevenue = 0;
  if (buckets['Revenue']) {
    Object.keys(buckets['Revenue']).forEach(function(sub) {
      if (sub !== 'Rental Income (Alma)') otherRevenue += buckets['Revenue'][sub].total;
    });
  }
  var totalRentAmount = grossRentalAmount + otherRevenue;

  // 5. Top-line totals (match the PDF's 4 tiles + Net Cash Flow).
  var directCostsTotal = 0;
  if (buckets['Direct Cost']) Object.keys(buckets['Direct Cost']).forEach(function(s) { directCostsTotal += buckets['Direct Cost'][s].total; });
  directCostsTotal += -almaFee;  // Alma fee is a direct cost too (already
                                  // reflected in the net deposit, but shown
                                  // explicitly on the report).

  var opexTotal = 0;
  if (buckets['Operating Expense']) Object.keys(buckets['Operating Expense']).forEach(function(s) { opexTotal += buckets['Operating Expense'][s].total; });

  var debtServiceTotal = 0;
  if (buckets['Debt Service']) Object.keys(buckets['Debt Service']).forEach(function(s) { debtServiceTotal += buckets['Debt Service'][s].total; });

  // Rents Collected (per Excel) = gross rental amount (top-line, pre-fees).
  var rentsCollected = totalRentAmount;
  var totalOpExOnly  = opexTotal + directCostsTotal;   // Amanda's "Total Operating Expenses" tile
  var netOperatingIncome = rentsCollected + totalOpExOnly;   // signs cancel: rents + (−costs)
  var netCashFlow    = netOperatingIncome + debtServiceTotal;

  // 6. Calendar / occupancy for the month.
  var occupancy = { nightsBooked: 0, revenueNights: 0, reservations: 0, daysInMonth: 0, occupancyPct: 0 };
  var res = getPropertyReservations(propertyId, monthStart.toISOString(), monthEnd.toISOString());
  if (res.events) {
    occupancy = _computeMonthOccupancy_(res.events, year, month);
  }
  var adr = occupancy.revenueNights > 0 ? (grossRentalAmount / occupancy.revenueNights) : 0;

  return {
    property:   prop,
    year:       year,
    month:      month,
    monthKey:   monthKey,
    monthLabel: Utilities.formatDate(monthStart, 'UTC', 'MMMM yyyy'),
    tiles: {
      rentsCollected:       rentsCollected,
      totalOperatingExpense: totalOpExOnly,
      netOperatingIncome:    netOperatingIncome,
      debtService:           debtServiceTotal,
      netCashFlow:           netCashFlow
    },
    operational: {
      occupancyPct:  occupancy.occupancyPct,
      reservations:  occupancy.reservations,
      nightsBooked:  occupancy.nightsBooked,
      averageDailyRate: adr
    },
    revenue: {
      grossRentalAmount:  grossRentalAmount,
      netRoomRevenue:     netRoomRevenue,
      otherRevenue:       otherRevenue
    },
    directCosts: {
      almaFee: almaFee,   // will be negative
      subcategories: buckets['Direct Cost'] || {}
    },
    operatingExpenses: buckets['Operating Expense'] || {},
    debtService:       buckets['Debt Service']       || {},
    reimbursements:    buckets['Reimbursement']      || {},
    uncategorized:     buckets['Uncategorized']      || {},
    needsReview:       needsReview,
    reservations:      (res.events || []).filter(function(e) {
      var s = new Date(e.start);
      var f = new Date(e.end);
      return f > monthStart && s < monthEnd;
    }),
    calendarError:     res.error || null
  };
}

// Web-callable — next N days of upcoming reservations, for the Properties
// tab "Upcoming Bookings" widget above the monthly report.
function getPropertyUpcoming(propertyId, days) {
  var d = Number(days) || 60;
  var start = new Date();
  var end   = new Date(start.getTime() + d * 86400000);
  return getPropertyReservations(propertyId, start.toISOString(), end.toISOString());
}


// ── Menu diagnostics ───────────────────────────────────────────────────────

// Menu-callable — wipe every rule attached to the Dorado property, then
// re-run seedDoradoProperty to install the current canonical rule set.
// Use after I update the DORADO_RULES in code so Amanda doesn't have to
// hand-reconcile existing rows. Idempotent + safe to re-run.
function resetDoradoRules() {
  _requireEditor_();
  ensurePropertiesSheets_();
  var dorado = getProperties().find(function(p) { return /dorado/i.test(String(p['Name']||'')); });
  if (!dorado) {
    try { SpreadsheetApp.getUi().alert('No Dorado property found. Run Seed Dorado first.'); } catch(e) {}
    return;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PROPERTY_RULES');
  if (!sheet || sheet.getLastRow() < 2) {
    // No rules yet — just seed.
    seedDoradoProperty();
    return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), PROPERTY_RULES_HEADERS.length);
  var data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  var hdr = data[0];
  var iProp = hdr.indexOf('Property ID');
  if (iProp < 0) return;
  var deleted = 0;
  // Iterate bottom-up so row numbers stay valid after deletes.
  for (var r = data.length - 1; r >= 1; r--) {
    if (String(data[r][iProp]) === String(dorado['ID'])) {
      sheet.deleteRow(r + 1);
      deleted++;
    }
  }
  // Re-install from the current DORADO_RULES in seedDoradoProperty.
  var seed = seedDoradoProperty();
  var installed = seed && seed.rulesInstalled || 0;
  var msg = 'Deleted ' + deleted + ' old Dorado rule(s).\nInstalled ' + installed + ' current rule(s).\n\nNext: "Debug: Dorado This Month" to see the categorized report.';
  try { SpreadsheetApp.getUi().alert('Dorado Rules Reset', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
  Logger.log(msg);
}


// Menu-callable: pick ANY month to debug (YYYY-MM). Useful when the current
// month is sparse (first few days) and Amanda wants to see a full-month
// example.
function debugDoradoPickMonth() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Debug: Dorado for Specific Month',
    'Enter month as YYYY-MM (e.g. 2026-09 for September 2026):',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var raw = String(r.getResponseText() || '').trim();
  var m = raw.match(/^(\d{4})-(\d{1,2})$/);
  if (!m) { ui.alert('Format must be YYYY-MM (e.g. 2026-09).'); return; }
  var year = Number(m[1]), month = Number(m[2]);
  if (month < 1 || month > 12) { ui.alert('Month must be 1-12.'); return; }
  _runDoradoDebugForMonth(year, month);
}

// Menu-callable: run the categorization engine for the month Amanda would
// actually REPORT ON right now, which is the LAST completed month (the
// monthly PDF goes out on the 20th covering the prior month). Current month
// is usually too sparse to be useful mid-month.
function debugDoradoThisMonth() {
  var now = new Date();
  // Last-completed month: back up one month from now, then take year + month.
  var y = now.getFullYear();
  var m = now.getMonth();          // 0-indexed. Previous month = this value.
  if (m === 0) { y--; m = 12; }    // January → December of prior year
  _runDoradoDebugForMonth(y, m);
}

function _runDoradoDebugForMonth(year, month) {
  var props = getProperties();
  var dorado = props.find(function(p) { return /dorado/i.test(String(p['Name'] || '')); });
  if (!dorado) { SpreadsheetApp.getUi().alert('Dorado property not found. Run Seed Dorado first.'); return; }
  var report = getPropertyMonthlyReport(dorado['ID'], year, month);
  var lines = [
    'Dorado — ' + report.monthLabel,
    '',
    'TILES',
    '  Rents Collected:       $' + report.tiles.rentsCollected.toFixed(2),
    '  Total Operating Exp:   $' + report.tiles.totalOperatingExpense.toFixed(2),
    '  Net Operating Income:  $' + report.tiles.netOperatingIncome.toFixed(2),
    '  Debt Service:          $' + report.tiles.debtService.toFixed(2),
    '  Net Cash Flow:         $' + report.tiles.netCashFlow.toFixed(2),
    '',
    'OPERATIONAL',
    '  Occupancy: ' + report.operational.occupancyPct + '%',
    '  Reservations: ' + report.operational.reservations,
    '  Nights booked: ' + report.operational.nightsBooked,
    '  ADR: $' + report.operational.averageDailyRate.toFixed(2),
    '',
    'REVENUE',
    '  Gross Rental Amount:  $' + report.revenue.grossRentalAmount.toFixed(2),
    '  Net Room Revenue:     $' + report.revenue.netRoomRevenue.toFixed(2),
    '',
    'DIRECT COSTS'
  ];
  // Direct costs: show the Alma fee (if any) first, then the sub-rules.
  if (report.directCosts && report.directCosts.almaFee) {
    lines.push('  Alma Fees (20%):  $' + report.directCosts.almaFee.toFixed(2) + '  (implied)');
  }
  if (report.directCosts && report.directCosts.subcategories) {
    Object.keys(report.directCosts.subcategories).forEach(function(sub) {
      lines.push('  ' + sub + ':  $' + report.directCosts.subcategories[sub].total.toFixed(2) +
                 '  (' + report.directCosts.subcategories[sub].transactions.length + ' txn)');
    });
  }
  if (!report.directCosts || (!report.directCosts.almaFee && (!report.directCosts.subcategories || !Object.keys(report.directCosts.subcategories).length))) {
    lines.push('  (none)');
  }
  lines.push('');
  lines.push('OPERATING EXPENSES');
  Object.keys(report.operatingExpenses).forEach(function(sub) {
    lines.push('  ' + sub + ':  $' + report.operatingExpenses[sub].total.toFixed(2) +
               '  (' + report.operatingExpenses[sub].transactions.length + ' txn)');
  });
  if (!Object.keys(report.operatingExpenses).length) lines.push('  (none)');
  lines.push('');
  lines.push('DEBT SERVICE');
  Object.keys(report.debtService).forEach(function(sub) {
    lines.push('  ' + sub + ':  $' + report.debtService[sub].total.toFixed(2));
  });
  lines.push('');
  lines.push('NEEDS REVIEW: ' + report.needsReview.length + ' uncategorized transaction(s)');
  report.needsReview.slice(0, 8).forEach(function(t) {
    lines.push('  ' + t.dateIso + '  ' + t.account + '  $' + t.amount.toFixed(2) + '  ' + t.name.substring(0, 40));
  });
  if (report.needsReview.length > 8) lines.push('  … +' + (report.needsReview.length - 8) + ' more');
  lines.push('');
  lines.push('CALENDAR: ' + (report.calendarError ? ('⚠ ' + report.calendarError) : (report.reservations.length + ' reservation(s) this month')));
  report.reservations.forEach(function(r) {
    lines.push('  ' + r.startIso + ' → ' + r.endIso + '  [' + r.source + ']  ' + r.title);
  });

  _propertyDebugOutput_('Dorado This Month', lines.join('\n'));
}

// Prints debug output to whatever channel is available — the sheet UI if we
// were invoked from a menu click, otherwise the execution log (for when
// Amanda has to Run from the Apps Script editor to trigger a reauth). Either
// way she sees the report; neither invocation throws "Cannot call getUi from
// this context".
function _propertyDebugOutput_(title, body) {
  Logger.log('══════ ' + title + ' ══════\n' + body);
  try {
    SpreadsheetApp.getUi().alert(title, body, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // No spreadsheet UI (running from Apps Script editor). The output is in
    // the execution log; nothing else to do. Rethrowing would hide it.
    Logger.log('(No spreadsheet UI — output above is in this execution log.)');
  }
}


// Menu-callable — triggers the Google Calendar OAuth consent prompt by
// calling CalendarApp WITHOUT a try/catch. This is the only reliable way
// to force Apps Script to re-prompt for the calendar.readonly scope after
// it's been added to appsscript.json: the normal reservation reader
// catches auth errors so the dashboard doesn't die on them, which means
// Apps Script never sees the uncaught error and never asks the user.
//
// Amanda runs this once from the Apps Script editor's Run button — it
// shows the "Review permissions" dialog, she clicks Allow for the
// calendar scope, and from then on CalendarApp works everywhere else.
// Deep diagnostic: show the Plaid Account IDs filter on the Dorado property,
// which accounts matched, and the last 90 days of Dorado-filtered transactions
// so we can see what expenses are actually posting + spot things the
// categorization rules don't yet match.
function debugDoradoTransactionPull() {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch(e) { ui = null; }
  var dorado = getProperties().find(function(p) { return /dorado/i.test(String(p['Name']||'')); });
  if (!dorado) {
    var m0 = 'No Dorado property found. Run "Properties → Seed Dorado" first.';
    Logger.log(m0); if (ui) ui.alert(m0); return;
  }
  var filterRaw = String(dorado['Plaid Account IDs'] || '').trim();
  var filters = filterRaw.split(',').map(function(s){return s.trim().toLowerCase();}).filter(Boolean);

  // Last 90 days.
  var now = new Date();
  var cutoff = new Date(now.getTime() - 90 * 86400000);

  var lines = [
    'DORADO TRANSACTION DIAGNOSTIC — last 90 days',
    '',
    'Dorado property:',
    '  Name: ' + dorado['Name'],
    '  ID: ' + dorado['ID'],
    '  Plaid Account IDs filter: "' + filterRaw + '"',
    '  → substrings matched (case-insensitive): [' + filters.join(', ') + ']',
    ''
  ];
  if (!filters.length) {
    lines.push('⚠ NO FILTERS SET. Run "Properties → Wire Dorado Plaid Accounts (6179 + 0451)" first.');
    _propertyDebugOutput_('Dorado Txn Pull', lines.join('\n'));
    return;
  }

  // Scan both sheets, group matches by account.
  var perAcctAll = {};
  var perAcct90  = {};
  var samples    = [];
  ['TLMND_TRANSACTIONS','PLAID_TRANSACTIONS'].forEach(function(sheetName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var iAcct = hdr.indexOf('Account');
    var iDate = hdr.indexOf('Date');
    var iName = hdr.indexOf('Name');
    var iAmt  = hdr.indexOf('Amount USD');
    if (iAcct < 0) return;
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, hdr.length).getValues();
    rows.forEach(function(r) {
      var acct = String(r[iAcct] || '');
      var acctLc = acct.toLowerCase();
      if (!filters.some(function(f) { return acctLc.indexOf(f) >= 0; })) return;
      perAcctAll[acct] = (perAcctAll[acct] || 0) + 1;
      var d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]);
      if (isNaN(d.getTime())) return;
      if (d >= cutoff) {
        perAcct90[acct] = (perAcct90[acct] || 0) + 1;
        samples.push({ date: d, acct: acct, name: String(r[iName]||''), amount: Number(r[iAmt]||0), sheet: sheetName });
      }
    });
  });

  lines.push('ACCOUNTS MATCHED');
  var allAccts = Object.keys(perAcctAll).sort();
  if (!allAccts.length) {
    lines.push('  (none — the filter substrings "' + filterRaw + '" don\'t match any Account value)');
    _propertyDebugOutput_('Dorado Txn Pull', lines.join('\n'));
    return;
  }
  allAccts.forEach(function(a) {
    lines.push('  • ' + a + '  →  ' + perAcctAll[a] + ' total, ' + (perAcct90[a] || 0) + ' in last 90 days');
  });

  lines.push('');
  lines.push('LAST 90 DAYS OF TRANSACTIONS (' + samples.length + ' total, showing up to 40 date-desc)');
  samples.sort(function(a, b) { return b.date - a.date; });
  samples.slice(0, 40).forEach(function(t) {
    var dStr = Utilities.formatDate(t.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var amtStr = (t.amount >= 0 ? '+' : '') + '$' + Math.abs(t.amount).toFixed(2);
    var acctTag = /dorado/i.test(t.acct) ? '[DOR]' : /savings/i.test(t.acct) ? '[SAV]' : '[CHK]';
    lines.push('  ' + dStr + '  ' + acctTag + '  ' + amtStr + '  ' + t.name.substring(0, 55));
  });
  if (!samples.length) {
    lines.push('  (nothing posted on any of the matched accounts in the last 90 days)');
  }

  _propertyDebugOutput_('Dorado Txn Pull', lines.join('\n'));
}


function grantCalendarAccess() {
  // No try/catch on purpose — propagate so Apps Script prompts for the scope.
  var cals = CalendarApp.getAllCalendars();
  Logger.log('Calendar access granted. Found ' + cals.length + ' calendars.');
  try {
    SpreadsheetApp.getUi().alert('Calendar Access Granted',
      'Found ' + cals.length + ' calendars available.\n\nYou can now run "Debug: Dorado This Month" and calendar reservations will load.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) { /* no UI (running from editor); output above is in exec log */ }
}


// Menu-callable — scans BOTH TLMND_TRANSACTIONS and PLAID_TRANSACTIONS to
// list every distinct Account value with its txn count, most recent date,
// and a few sample transaction names. Used to figure out what substrings to
// paste into a property's "Plaid Account IDs" field (e.g. does the Oriental
// account show up as "Oriental Bank ···4321" or "ORIENTAL - Checking"?).
function debugListPlaidAccounts() {
  var acctMap = {};
  var sheetsScanned = [];
  ['TLMND_TRANSACTIONS', 'PLAID_TRANSACTIONS'].forEach(function(sheetName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    sheetsScanned.push(sheetName + ' (' + (sheet.getLastRow() - 1) + ' rows)');
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var iAcct = headers.indexOf('Account');
    var iDate = headers.indexOf('Date');
    var iName = headers.indexOf('Name');
    if (iAcct < 0) return;
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    rows.forEach(function(r) {
      var a = String(r[iAcct] || '').trim();
      if (!a) return;
      if (!acctMap[a]) acctMap[a] = { count: 0, latest: null, samples: [], sources: {} };
      acctMap[a].count++;
      acctMap[a].sources[sheetName] = true;
      var d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]);
      if (!isNaN(d.getTime()) && (!acctMap[a].latest || d > acctMap[a].latest)) acctMap[a].latest = d;
      if (acctMap[a].samples.length < 3) acctMap[a].samples.push(String(r[iName] || '').substring(0, 45));
    });
  });
  if (!sheetsScanned.length) {
    SpreadsheetApp.getUi().alert('Both TLMND_TRANSACTIONS and PLAID_TRANSACTIONS are empty. Run "All Transactions → Sync ALL Plaid Transactions" first.');
    return;
  }
  var lines = ['PLAID ACCOUNTS FOUND (' + Object.keys(acctMap).length + ')', '',
               'Scanned: ' + sheetsScanned.join(', '), ''];
  Object.keys(acctMap).sort().forEach(function(a) {
    var m = acctMap[a];
    var srcs = Object.keys(m.sources).join(' + ');
    lines.push('  • "' + a + '"  [' + srcs + ']');
    lines.push('      ' + m.count + ' txn(s) · latest: ' + (m.latest ? Utilities.formatDate(m.latest, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '—'));
    m.samples.forEach(function(s) { lines.push('      ex: ' + s); });
    lines.push('');
  });
  lines.push('To attach an account to a property:');
  lines.push('  1. Open the PROPERTIES sheet');
  lines.push('  2. In the row for that property, edit the "Plaid Account IDs" column');
  lines.push('  3. Comma-separate substrings that appear in the account names above');
  lines.push('     (e.g. "oriental,ath,9007" catches all three)');
  _propertyDebugOutput_('Plaid Accounts', lines.join('\n'));
}


// Menu-callable — wire Dorado's PROPERTIES row to the right Plaid account
// substrings so transactions start flowing. Based on Amanda's PLAID_TRANSACTIONS
// debug:
//   Oriental - MN - Dorado PH ···6179       → Dorado operating account
//   Oriental - MN Personal - Checking ···0451 → Mike's ATH / personal (cleaner,
//                                               exterminator, AC pass-through)
//   Fidelity ···9007                        → not in Plaid; mortgage handled
//                                               by the PROPERTY_RECURRING entry
// Idempotent — safe to re-run; always sets the field to the canonical value.
function wireDoradoPlaidAccounts() {
  _requireEditor_();
  ensurePropertiesSheets_();
  var dorado = getProperties().find(function(p) { return /dorado/i.test(String(p['Name']||'')); });
  if (!dorado) {
    try { SpreadsheetApp.getUi().alert('No Dorado property found. Run "Properties → Seed Dorado" first.'); } catch(e) {}
    return;
  }

  // Writing "6179,0451" via setValue causes Google Sheets to parse it as a
  // NUMBER (comma = thousands separator) and store 61790451 — then no
  // account substring matches. Fix: force the cell to text format first,
  // then write. Also includes both Checking AND Savings sub-accounts that
  // Amanda confirmed both end in 0451.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PROPERTIES');
  var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var iId  = hdr.indexOf('ID');
  var iPai = hdr.indexOf('Plaid Account IDs');
  var iUpd = hdr.indexOf('Last Updated');
  if (iId < 0 || iPai < 0) {
    try { SpreadsheetApp.getUi().alert('PROPERTIES sheet missing ID or Plaid Account IDs column.'); } catch(e) {}
    return;
  }
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, hdr.length).getValues();
  var newVal = '6179,0451';
  var found = false;
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][iId]) !== String(dorado['ID'])) continue;
    var cell = sheet.getRange(r + 2, iPai + 1);
    cell.setNumberFormat('@');   // force text — prevents comma→number parsing
    cell.setValue(newVal);
    if (iUpd >= 0) sheet.getRange(r + 2, iUpd + 1).setValue(new Date());
    found = true;
    break;
  }
  if (!found) {
    try { SpreadsheetApp.getUi().alert('Dorado row not found on PROPERTIES sheet.'); } catch(e) {}
    return;
  }
  _logAudit_('wireDoradoPlaid', 'property', dorado['ID'], 'Dorado',
             'Set Plaid Account IDs = ' + newVal);

  try {
    SpreadsheetApp.getUi().alert('Dorado Plaid Accounts Wired',
      'Set Dorado "Plaid Account IDs" = ' + newVal + '\n' +
      '(cell forced to text format so Sheets doesn\'t mangle the comma)\n\n' +
      'This filters PLAID_TRANSACTIONS + TLMND_TRANSACTIONS to the Dorado-relevant rows:\n' +
      '  • Oriental - MN - Dorado PH ···6179 (operating account)\n' +
      '  • Oriental - MN Personal - Checking ···0451 (ATH pass-through for Lourdes / exterminator / AC)\n' +
      '  • Oriental - MN Personal - Savings ···0451 (same mask — captured too)\n\n' +
      'Fidelity 9007 (mortgage) stays handled by the PROPERTY_RECURRING entry.\n\n' +
      'Next: "Properties → Debug: Dorado This Month" to see the live report numbers.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}


// Menu-callable — install the Dorado mortgage as a monthly recurring entry.
// $21,668.72 due the 1st of every month (per the Oriental mortgage statement
// Amanda shared). Amount is fixed so no Plaid matching needed — the monthly
// report auto-includes it for every month going forward.
// Idempotent: skips if an existing recurring row already matches.
// Menu-callable — install the Dorado Beach Club dues ($1,112.22/mo) as a
// recurring entry. The raw Plaid payment to DBR Dorado varies month-to-month
// (dues + occasional food charges) so we don't rely on auto-match. Instead
// this recurring entry supplies the clean dues amount every month and the
// matching Plaid rule hides the raw payment as "Internal Transfer".
function seedDoradoClubDuesRecurring() {
  _requireEditor_();
  ensurePropertiesSheets_();
  var dorado = getProperties().find(function(p) { return /dorado/i.test(String(p['Name']||'')); });
  if (!dorado) {
    try { SpreadsheetApp.getUi().alert('No Dorado property found. Run "Properties → Seed Dorado" first.'); } catch(e) {}
    return;
  }
  var existing = _getPropertySheetRows_('PROPERTY_RECURRING', PROPERTY_RECURRING_HEADERS)
    .find(function(r) {
      return String(r['Property ID']) === String(dorado['ID']) &&
             /dorado club|club dues/i.test(String(r['Label']||''));
    });
  if (existing) {
    try { SpreadsheetApp.getUi().alert('Dorado Club Dues recurring entry already exists (' + existing['ID'] + '). No changes made.'); } catch(e) {}
    return;
  }
  var id = 'rec_' + Utilities.getUuid().substring(0, 8);
  _writePropertyRow_('PROPERTY_RECURRING', PROPERTY_RECURRING_HEADERS, {
    'ID':            id,
    'Property ID':   dorado['ID'],
    'Label':         'Dorado Club Dues',
    'Amount':        -1112.22,          // negative = outflow
    'Category':      'Operating Expense',
    'Subcategory':   'Dorado Club Fees',
    'Day of Month':  25,
    'Start Month':   '2024-01',
    'End Month':     '',
    'Active':        'Yes',
    'Notes':         'Fixed monthly dues. Raw DBR Dorado Plaid payment is hidden (varies with Mike\'s food charges); this clean $1,112.22 line is the real dues expense.',
    'Date Added':    new Date(),
    'Last Updated':  new Date()
  });
  _logAudit_('seedDoradoClubDues', 'property', dorado['ID'], 'Dorado',
             'Added $1,112.22 monthly recurring (Dorado Club Dues).');
  try {
    SpreadsheetApp.getUi().alert('Dorado Club Dues Installed',
      '$1,112.22 outflow added as a monthly recurring entry for Dorado on the 25th of each month, starting 2024-01.\n\n' +
      'The Dorado monthly report will now auto-include this as a $1,112.22 Operating Expense → Dorado Club Fees line, regardless of what Mike spent on food at the club.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}


function seedDoradoMortgageRecurring() {
  _requireEditor_();
  ensurePropertiesSheets_();
  var dorado = getProperties().find(function(p) { return /dorado/i.test(String(p['Name']||'')); });
  if (!dorado) {
    try { SpreadsheetApp.getUi().alert('No Dorado property found. Run "Properties → Seed Dorado" first.'); } catch(e) {}
    return;
  }
  var existing = _getPropertySheetRows_('PROPERTY_RECURRING', PROPERTY_RECURRING_HEADERS)
    .find(function(r) {
      return String(r['Property ID']) === String(dorado['ID']) &&
             /mortgage/i.test(String(r['Label']||''));
    });
  if (existing) {
    try { SpreadsheetApp.getUi().alert('Dorado mortgage recurring entry already exists (' + existing['ID'] + '). No changes made.'); } catch(e) {}
    return;
  }
  var id = 'rec_' + Utilities.getUuid().substring(0, 8);
  _writePropertyRow_('PROPERTY_RECURRING', PROPERTY_RECURRING_HEADERS, {
    'ID':            id,
    'Property ID':   dorado['ID'],
    'Label':         'Dorado Mortgage (Oriental)',
    'Amount':        -21668.72,                // negative = outflow
    'Category':      'Debt Service',
    'Subcategory':   'Mortgage',
    'Day of Month':  1,
    'Start Month':   '2024-01',                // adjust in the sheet if needed
    'End Month':     '',
    'Active':        'Yes',
    'Notes':         'Static recurring entry. Fidelity ···9007 is not in Plaid; this covers the mortgage line until we connect that account.',
    'Date Added':    new Date(),
    'Last Updated':  new Date()
  });
  _logAudit_('seedDoradoMortgage', 'property', dorado['ID'], 'Dorado',
             'Added $21,668.72 monthly recurring (Dorado Mortgage).');
  try {
    SpreadsheetApp.getUi().alert('Dorado Mortgage Installed',
      '$21,668.72 outflow added as a monthly recurring entry for Dorado.\n\n' +
      'Day of month: 1st\nStart month: 2024-01\nEnd month: (indefinite)\n\n' +
      'The monthly Dorado report will auto-include this as a Debt Service → Mortgage line, starting Jan 2024, every month going forward. Edit the PROPERTY_RECURRING sheet directly if you need to change the amount, date, or end it.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}


// Menu-callable — dumps every property + rule to a text alert. Handy after
// seeding to confirm everything landed. Read-only.
function debugProperties() {
  var props = _getPropertySheetRows_('PROPERTIES', PROPERTIES_HEADERS);
  var rules = _getPropertySheetRows_('PROPERTY_RULES', PROPERTY_RULES_HEADERS);
  var lines = ['PROPERTIES (' + props.length + ')', ''];
  props.forEach(function(p) {
    lines.push('  • ' + p['Name']);
    lines.push('      ID: ' + p['ID'] + '  ·  Active: ' + p['Active']);
    lines.push('      Calendar: ' + (p['Google Calendar ID'] || '(none)'));
    lines.push('      Linked Asset: ' + (p['Linked Asset ID'] || '(none)'));
    lines.push('      Linked Liability: ' + (p['Linked Liability ID'] || '(none)'));
    lines.push('      Plaid Accounts: ' + (p['Plaid Account IDs'] || '(none)'));
    var pr = rules.filter(function(r){return r['Property ID']===p['ID'];});
    lines.push('      Rules: ' + pr.length);
    pr.forEach(function(r){
      lines.push('        [' + r['Priority'] + '] ' + r['Subcategory'] +
                 '  ← acct~"' + (r['Match Account']||'*') + '" name~"' + (r['Match Name']||'*') +
                 '" amt="' + (r['Match Amount']||'*') + '" ' + r['Direction']);
    });
    lines.push('');
  });
  _propertyDebugOutput_('Properties Diagnostic', lines.join('\n'));
}
