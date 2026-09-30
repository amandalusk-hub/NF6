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


// ── Sheet lifecycle ────────────────────────────────────────────────────────

function ensurePropertiesSheets_() {
  _ensurePropertySheet_('PROPERTIES', PROPERTIES_HEADERS);
  _ensurePropertySheet_('PROPERTY_RULES', PROPERTY_RULES_HEADERS);
  _ensurePropertySheet_('PROPERTY_MANUAL', PROPERTY_MANUAL_HEADERS);
  _ensurePropertySheet_('PROPERTY_TXN_OVERRIDES', PROPERTY_TXN_OVERRIDES_HEADERS);
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
    // Rental income from Alma — arrives via Oriental Bank as a positive
    // deposit. The manager fee % on the property row is used to gross the
    // deposit up to the pre-fee rental amount on the report.
    { priority: 10, matchAccount:'oriental', matchName:'alma',   matchAmount:'', direction:'in',
      category:'Revenue', subcategory:'Rental Income (Alma)',
      notes:'Net deposit (after 20% Alma fee). Gross up on report.' },

    // Recurring HOA + club fees — identified by exact amount on Oriental.
    { priority: 20, matchAccount:'oriental', matchName:'', matchAmount:'-1254.88', direction:'out',
      category:'Operating Expense', subcategory:'HOA Dues' },
    { priority: 20, matchAccount:'oriental', matchName:'', matchAmount:'-1112.22', direction:'out',
      category:'Operating Expense', subcategory:'Dorado Club Fees' },

    // Utilities — by biller name.
    { priority: 30, matchAccount:'oriental', matchName:'claro', matchAmount:'', direction:'out',
      category:'Operating Expense', subcategory:'Internet' },
    { priority: 30, matchAccount:'oriental', matchName:'luma',  matchAmount:'', direction:'out',
      category:'Operating Expense', subcategory:'Utilities - Electricity' },

    // ATH mobile — payee not exposed by Plaid, so we match by amount.
    // Amount-only rules are scoped to the ATH account so a $60 payment from
    // Oriental doesn't get mis-tagged as the exterminator.
    { priority: 40, matchAccount:'ath', matchName:'', matchAmount:'-60.00',  direction:'out',
      category:'Operating Expense', subcategory:'Maintenance - Exterminator' },
    { priority: 40, matchAccount:'ath', matchName:'', matchAmount:'-295.00', direction:'out',
      category:'Operating Expense', subcategory:'Maintenance - AC Quarterly' },
    { priority: 40, matchAccount:'ath', matchName:'', matchAmount:'-140.00|-160.00', direction:'out',
      category:'Direct Cost', subcategory:'Cleaning (Lourdes)' },

    // Mortgage — paid out of Mike's Fidelity account ending 9007.
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
  SpreadsheetApp.getUi().alert('Properties Diagnostic', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
