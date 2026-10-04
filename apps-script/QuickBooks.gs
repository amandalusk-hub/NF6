/**
 * QuickBooks.gs — true double-entry bookkeeping module.
 *
 * Goal: an exact in-app replica of Mike's QuickBooks files, one entity at a
 * time. Pass 1 lives here for NF USA CA LLC ("NF CA"), seeded from the three
 * reports Amanda exported from QBO:
 *   • Account List           → QB_COA
 *   • Trial Balance 12/31/25 → opening JE in QB_GL_ENTRIES (source=opening)
 *   • Balance Sheet 12/31/25 → sanity check only
 *
 * Data model (per entity, same shape clones for the next entity):
 *   QB_COA              — the chart of accounts. One row per account.
 *   QB_GL_ENTRIES       — the general ledger. One row per JE LINE. Entries
 *                         are groups of rows sharing the same Entry ID.
 *                         Σ debits = Σ credits per Entry ID (hard-enforced
 *                         on manual post; opening JE is self-balancing).
 *   QB_RULES            — Plaid auto-coding rules (same shape as
 *                         PROPERTY_RULES — matchAccount/matchName/matchAmount
 *                         → DR account + CR account). Future.
 *   QB_RECONCILIATIONS  — bank-rec headers (one per statement). Each GL line
 *                         carries a reconciliation ref back to this. Future.
 *
 * Everything flows out of QB_GL_ENTRIES. Trial Balance = sum (Debit−Credit)
 * grouped by account. Balance Sheet = TB pivoted by COA type for Asset /
 * Liability / Equity. Income Statement = Revenue − Expense within a window.
 * General Ledger = all lines against one account within a window.
 */

// ── Sheet schemas ─────────────────────────────────────────────────────────

var QB_COA_HEADERS = [
  'ID',                    // qbcoa-NFCA-001 (stable, used as foreign key)
  'Entity',                // 'NF CA' etc.
  'Account Number',        // Optional — QB NFCA doesn't use them; leave blank
  'Account Name',          // 'Chase - 2086', 'Rental Income', etc.
  'Parent Account Name',   // '' or e.g. 'Capital TLMND' for sub-accounts
  'Type',                  // Bank | Accounts receivable (A/R) | Other Current Assets | Fixed Assets | Other Assets |
                           // Accounts payable (A/P) | Credit Card | Other Current Liabilities | Long Term Liabilities |
                           // Equity | Income | Cost of Goods Sold | Expenses | Other Income | Other Expense
  'Detail Type',           // QB's finer-grained subtype (free text per QB)
  'Normal Balance',        // 'Debit' or 'Credit' — derived from Type on seed
  'Active',                // 'Yes' | 'No'
  'Description',           // Free-text, from the QB account description
  'Plaid Account IDs',     // Comma-sep Plaid account IDs for Bank/CC accounts
  'Date Added',
  'Last Updated'
];

var QB_GL_ENTRIES_HEADERS = [
  'Entry ID',              // je-001234 — groups 2+ rows into one journal entry
  'Line #',                // 1..N within the entry
  'Date',                  // Transaction date (not post date)
  'Entity',
  'Memo',                  // Entry-level memo (same for every line of one JE)
  'Account Name',          // Joins to QB_COA.Account Name
  'Debit',                 // Signed +ve; blank for a credit-only line
  'Credit',                // Signed +ve; blank for a debit-only line
  'Source',                // opening | manual | plaid | movement | recurring | depreciation
  'Source Ref',            // Plaid txn ID / Movement wire ID / opening-balance / ...
  'Line Memo',             // Optional per-line note
  'Reconciled Date',       // Set when bank-reconciled against a statement
  'Reconciliation Ref',    // Points to a QB_RECONCILIATIONS row
  'Created By',
  'Created At'
];

var QB_RULES_HEADERS = [
  'ID', 'Entity', 'Priority',
  'Match Plaid Account',   // last-4 or full ID
  'Match Name Contains',
  'Match Amount',          // '-250.00' or '-250.00|-500.00' (any of)
  'Direction',             // in | out | both
  'DR Account',            // name from QB_COA
  'CR Account',
  'Memo Template',
  'Active',
  'Date Added', 'Last Updated'
];

var QB_RECONCILIATIONS_HEADERS = [
  'ID', 'Entity', 'Account Name', 'Statement Date',
  'Statement Beginning Balance', 'Statement Ending Balance',
  'Reconciled On', 'Reconciled By', 'Notes'
];


// ── Sheet lifecycle ───────────────────────────────────────────────────────

function ensureQBSheets_() {
  _ensureQBSheet_('QB_COA', QB_COA_HEADERS);
  _ensureQBSheet_('QB_GL_ENTRIES', QB_GL_ENTRIES_HEADERS);
  _ensureQBSheet_('QB_RULES', QB_RULES_HEADERS);
  _ensureQBSheet_('QB_RECONCILIATIONS', QB_RECONCILIATIONS_HEADERS);
}

function _ensureQBSheet_(name, headers) {
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
  // Schema drift: add any missing columns at the end.
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


// ── Row read/write helpers ────────────────────────────────────────────────

function _getQBRows_(sheetName, headers) {
  ensureQBSheets_();
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
  }).filter(function(o) {
    // Filter blanks: require at least one key field. COA uses ID; GL uses Entry ID.
    return o.ID || o['Entry ID'];
  });
}

function _writeQBRow_(sheetName, headers, obj) {
  ensureQBSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var lastCol = Math.max(sheet.getLastColumn(), headers.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = hdr.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.appendRow(row);
  // Keep Plaid Account IDs as text so commas-separated last-4s don't get
  // auto-converted to numbers (same bug we hit in Properties).
  var textCols = ['Plaid Account IDs', 'Account Number'];
  var newRow = sheet.getLastRow();
  textCols.forEach(function(c) {
    var idx = hdr.indexOf(c);
    if (idx >= 0 && obj[c] !== undefined && obj[c] !== '') {
      var cell = sheet.getRange(newRow, idx + 1);
      cell.setNumberFormat('@');
      cell.setValue(String(obj[c]));
    }
  });
}


// ── Type → Normal Balance derivation ──────────────────────────────────────
// Mirrors QB's convention: Asset + Expense are debit-normal; everything else
// is credit-normal. Draws / Distributions are Equity type but behave like
// debits — those we tag explicitly in the seed since the account NAME tells
// us, not the TYPE.

var _QB_ASSET_TYPES = {
  'Bank': true, 'Accounts receivable (A/R)': true, 'Other Current Assets': true,
  'Fixed Assets': true, 'Other Assets': true
};
var _QB_LIABILITY_TYPES = {
  'Accounts payable (A/P)': true, 'Credit Card': true,
  'Other Current Liabilities': true, 'Long Term Liabilities': true
};
var _QB_EQUITY_TYPES = { 'Equity': true };
var _QB_REVENUE_TYPES = { 'Income': true, 'Other Income': true };
var _QB_EXPENSE_TYPES = { 'Cost of Goods Sold': true, 'Expenses': true, 'Other Expense': true };

function _qbClassifyType_(type) {
  if (_QB_ASSET_TYPES[type])     return 'Asset';
  if (_QB_LIABILITY_TYPES[type]) return 'Liability';
  if (_QB_EQUITY_TYPES[type])    return 'Equity';
  if (_QB_REVENUE_TYPES[type])   return 'Revenue';
  if (_QB_EXPENSE_TYPES[type])   return 'Expense';
  return 'Other';
}

function _qbNormalBalance_(type, accountName) {
  var cls = _qbClassifyType_(type);
  if (cls === 'Asset')   return 'Debit';
  if (cls === 'Expense') return 'Debit';
  // z-Accumulated Depreciation is a contra-asset — normal balance CREDIT.
  // Caller passes the account name so we can special-case these.
  if (/accumulated depreciation/i.test(accountName)) return 'Credit';
  if (cls === 'Liability') return 'Credit';
  if (cls === 'Revenue')   return 'Credit';
  // Equity: default credit, but Draws + Distributions are debit-natured.
  if (cls === 'Equity') {
    if (/distribution|draw/i.test(accountName)) return 'Debit';
    return 'Credit';
  }
  return 'Debit';
}


// ── NF CA chart of accounts seed (from Account List PDF, Oct 4 2026) ──────
// EXACT replica of QB. Account hierarchy, types, detail types, descriptions
// all mirrored. Opening balances come separately from the Trial Balance.

var NFCA_COA = [
  // Assets
  { name:'Chase - 2086',                     type:'Bank',                        detail:'Checking',             desc:'' },
  { name:'Accounts Receivable',              type:'Accounts receivable (A/R)',   detail:'Accounts Receivable (A/R)', desc:'Unpaid or unapplied customer invoices and credits' },
  { name:'Uncategorized Asset',              type:'Other Current Assets',        detail:'Other Current Assets', desc:'' },
  { name:'Uncategorized Asset-1',            type:'Other Current Assets',        detail:'Other Current Assets', desc:'' },
  { name:'Undeposited Funds',                type:'Other Current Assets',        detail:'Undeposited Funds',    desc:'' },
  { name:'Furniture and Equipment',          type:'Fixed Assets',                detail:'Other fixed assets',   desc:'Furniture and equipment with useful life exceeding one year' },
  { name:'Investment -5330 Carroll Canyon',  type:'Fixed Assets',                detail:'Other fixed assets',   desc:'' },
  { name:'Medical Equipment',                type:'Fixed Assets',                detail:'Other fixed assets',   desc:'Equipment used in diagnostic and therapeutic procedures, exam tables, etc.' },
  { name:'z-Accumulated Depreciation',       type:'Fixed Assets',                detail:'Other fixed assets',   desc:'Accumulated depreciation on equipment, buildings and improvements' },
  { name:'Security Deposits Asset',          type:'Other Assets',                detail:'Other Long-term Assets', desc:'Deposits and other returnable funds held by other entities (Rent, Utilities, etc.)' },

  // Liabilities
  { name:'Accounts Payable',                 type:'Accounts payable (A/P)',      detail:'Accounts Payable (A/P)', desc:'Unpaid or unapplied vendor bills or credits' },
  { name:'Payroll Liabilities',              type:'Other Current Liabilities',   detail:'Payroll Tax Payable',   desc:'Unpaid payroll liabilities. Amounts withheld or accrued, but not yet paid' },

  // Equity
  { name:'Capital TLMND',                    type:'Equity',                      detail:"Owner's Equity",        desc:'' },
  { name:'Capital TLMND:5330 Carroll Canyon',type:'Equity', parent:'Capital TLMND', detail:"Owner's Equity",    desc:'' },
  { name:'Distributions',                    type:'Equity',                      detail:"Owner's Equity",        desc:'' },
  { name:'Member 1 Draws',                   type:'Equity',                      detail:"Owner's Equity",        desc:'Monies taken out of the business by member 1' },
  { name:'Member 1 Equity',                  type:'Equity',                      detail:"Owner's Equity",        desc:'Equity for member 1' },
  { name:'Member 2 Draws',                   type:'Equity',                      detail:"Owner's Equity",        desc:'Monies taken out of the business by member 2' },
  { name:'Member 2 Equity',                  type:'Equity',                      detail:"Owner's Equity",        desc:'Equity for member 2' },
  { name:'Opening Balance Equity',           type:'Equity',                      detail:'Opening Balance Equity',desc:'Opening balances during setup post to this account. The balance of this account should be zero after' },
  { name:'Retained Earnings',                type:'Equity',                      detail:'Retained Earnings',     desc:'Undistributed earnings of the business' },

  // Income
  { name:'Billable Expense Income',          type:'Income',                      detail:'Sales of Product Income', desc:'' },
  { name:'Capitation Fees',                  type:'Income',                      detail:'Sales of Product Income', desc:'Fees received from healthcare management organizations to cover patient care contracts' },
  { name:'Fee for Service Income',           type:'Income',                      detail:'Sales of Product Income', desc:'Receipts for patient fees, directly related to services (non-capitated fees)' },
  { name:'Markup',                           type:'Income',                      detail:'Sales of Product Income', desc:'' },
  { name:'Nonmedical Income',                type:'Income',                      detail:'Sales of Product Income', desc:'Records fees and other nonmedical income' },
  { name:'Refunds',                          type:'Income',                      detail:'Sales of Product Income', desc:'Refunds paid to patients' },
  { name:'Rental Income',                    type:'Income',                      detail:'Sales of Product Income', desc:'' },
  { name:'Sales',                            type:'Income',                      detail:'Sales of Product Income', desc:'' },
  { name:'Uncategorized Income',             type:'Income',                      detail:'Sales of Product Income', desc:'' },

  // Expenses
  { name:'Advertising and Promotion',        type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Advertising, marketing, graphic design, and other promotional expenses' },
  { name:'Automobile Expense',               type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Fuel, oil, repairs, and other automobile maintenance for business autos' },
  { name:'Bank Service Charges',             type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Bank account service fees, bad check charges and other bank fees' },
  { name:'Business Licenses and Permits',    type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Business licenses, permits, and other business-related fees' },
  { name:'Computer and Internet Expenses',   type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Computer supplies, off-the-shelf software, online fees, and other computer or internet related expen' },
  { name:'Continuing Education',             type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Seminars, educational expenses and employee development, not including travel' },
  { name:'Depreciation Expense',             type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Depreciation on equipment, buildings and improvements' },
  { name:'Dues and Subscriptions',           type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Subscriptions and membership dues for civic, service, professional, trade organizations' },
  { name:'HOA Fees',                         type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'' },
  { name:'Insurance Expense',                type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Insurance expenses' },
  { name:'Interest Expense',                 type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Interest payments on business loans, credit card balances, or other business debt' },
  { name:'Janitorial Expense',               type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Janitorial expenses and cleaning supplies' },
  { name:'Laboratory Fees',                  type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Charges from outside laboratories' },
  { name:'Meals and Entertainment',          type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Business meals and entertainment expenses, including travel-related meals (may have limited deductib' },
  { name:'Medical Records and Supplies',     type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Filing supplies for medical records' },
  { name:'Office Supplies',                  type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Office supplies expense' },
  { name:'Payroll Expenses',                 type:'Expenses', detail:'Payroll Expenses',                 desc:'Payroll expenses' },
  { name:'Professional Fees',                type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Payments to accounting professionals and attorneys for accounting or legal services' },
  { name:'Real Estate Property Taxes',       type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'' },
  { name:'Reference Materials',              type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Coding books, anatomical charts and models, etc.' },
  { name:'Rent Expense',                     type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Rent paid for company offices or other structures used in the business' },
  { name:'Repairs and Maintenance',          type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Incidental repairs and maintenance of business assets that do not add to the value or appreciably pr' },
  { name:'Small Medical Equipment',          type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Purchases of small instruments and equipment not classified as fixed assets' },
  { name:'Telephone Expense',                type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Telephone and long distance charges, faxing, and other fees Not equipment purchases' },
  { name:'Travel Expense',                   type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Business-related travel expenses including airline tickets, taxi fares, hotel and other travel expen' },
  { name:'Uncategorized Expense',            type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'' },
  { name:'Uniforms',                         type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Uniforms for employees and contractors' },
  { name:'Utilities',                        type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Water, electricity, garbage, and other basic utilities expenses' },
  { name:'Vaccines and Medicines',           type:'Expenses', detail:'Other Miscellaneous Service Cost', desc:'Vaccines, medicines, and other drugs' },

  // Other Expense
  { name:'Ask My Accountant',                type:'Other Expense', detail:'Other Miscellaneous Expense', desc:'Transactions to be discussed with accountant, consultant, or tax preparer' },
  { name:'Reconciliation Discrepancies',     type:'Other Expense', detail:'Other Miscellaneous Expense', desc:'' }
];

// Opening balances from the Trial Balance 12/31/2025. Signed by DR/CR as QB
// shows them. Posting this as a single balancing JE gives every account its
// opening balance. Opening Balance Equity is the plug when a TB would
// otherwise not balance — Amanda's TB already balances perfectly ($587,654.75
// on both sides), so OBE stays $0 (QB's goal).
var NFCA_OPENING_TB = {
  asOfDate: '2025-12-31',
  lines: [
    { account:'Chase - 2086',                    debit: 6258.00,     credit: 0 },
    { account:'Investment -5330 Carroll Canyon', debit: 340000.00,   credit: 0 },
    { account:'z-Accumulated Depreciation',      debit: 0,           credit: 8718.00 },
    { account:'Capital TLMND',                   debit: 0,           credit: 342000.00 },
    { account:'Distributions',                   debit: 241396.75,   credit: 0 },
    { account:'Retained Earnings',               debit: 0,           credit: 185840.75 },
    { account:'Rental Income',                   debit: 0,           credit: 51096.00 }
  ]
};


// ── Public API: seed NF CA ────────────────────────────────────────────────
// Idempotent — safe to re-run. If the COA already has rows for NF CA, we
// skip the COA load. If the opening JE already exists (we tag it with a
// stable Entry ID), we skip it too.

function seedNFCA() {
  _requireEditor_();
  ensureQBSheets_();
  var entity = 'NF CA';
  var now = new Date();
  var user = _currentUserEmail_();

  // 1. COA — skip if any NF CA row already present.
  var existingCOA = _getQBRows_('QB_COA', QB_COA_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity; });
  var seededCOA = 0;
  if (existingCOA.length === 0) {
    NFCA_COA.forEach(function(a, i) {
      var id = 'qbcoa-NFCA-' + _pad3_(i + 1);
      _writeQBRow_('QB_COA', QB_COA_HEADERS, {
        'ID':                 id,
        'Entity':             entity,
        'Account Number':     '',
        'Account Name':       a.name,
        'Parent Account Name': a.parent || '',
        'Type':               a.type,
        'Detail Type':        a.detail || '',
        'Normal Balance':     _qbNormalBalance_(a.type, a.name),
        'Active':             'Yes',
        'Description':        a.desc || '',
        'Plaid Account IDs':  '',
        'Date Added':         now,
        'Last Updated':       now
      });
      seededCOA++;
    });
  }

  // 2. Opening JE — stable Entry ID so re-runs skip it.
  var openingEntryId = 'je-NFCA-opening-2025-12-31';
  var existingOpening = _getQBRows_('QB_GL_ENTRIES', QB_GL_ENTRIES_HEADERS)
    .some(function(r) { return String(r['Entry ID']) === openingEntryId; });
  var seededGL = 0;
  if (!existingOpening) {
    var dt = new Date(NFCA_OPENING_TB.asOfDate + 'T00:00:00Z');
    NFCA_OPENING_TB.lines.forEach(function(ln, i) {
      _writeQBRow_('QB_GL_ENTRIES', QB_GL_ENTRIES_HEADERS, {
        'Entry ID':         openingEntryId,
        'Line #':           i + 1,
        'Date':             dt,
        'Entity':           entity,
        'Memo':             'Opening balances as of ' + NFCA_OPENING_TB.asOfDate + ' (from QB Trial Balance)',
        'Account Name':     ln.account,
        'Debit':            ln.debit  || '',
        'Credit':           ln.credit || '',
        'Source':           'opening',
        'Source Ref':       'tb-' + NFCA_OPENING_TB.asOfDate,
        'Line Memo':        '',
        'Reconciled Date':  '',
        'Reconciliation Ref':'',
        'Created By':       user,
        'Created At':       now
      });
      seededGL++;
    });
  }

  var msg = 'NF CA seed complete:\n' +
            '  COA accounts: ' + seededCOA + (existingCOA.length ? ' (skipped — ' + existingCOA.length + ' already there)' : '') + '\n' +
            '  Opening JE lines: ' + seededGL + (existingOpening ? ' (skipped — opening JE already posted)' : '');
  try { SpreadsheetApp.getUi().alert('Seed NF CA', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
  Logger.log(msg);
  return { entity: entity, coaSeeded: seededCOA, glSeeded: seededGL };
}

function _pad3_(n) {
  var s = String(n);
  while (s.length < 3) s = '0' + s;
  return s;
}


// ── Web-callable reads ────────────────────────────────────────────────────

// Chart of Accounts for an entity. Returns rows in COA display order
// (asset→liability→equity→revenue→expense) with the current running balance
// for each account computed from GL entries.
function getQBChartOfAccounts(entity) {
  entity = String(entity || 'NF CA');
  var coa = _getQBRows_('QB_COA', QB_COA_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity; });
  var balances = _computeAccountBalances_(entity, null, null);   // all time → "current"
  var typeOrder = {
    'Bank': 10, 'Accounts receivable (A/R)': 20, 'Other Current Assets': 30,
    'Fixed Assets': 40, 'Other Assets': 50,
    'Accounts payable (A/P)': 60, 'Credit Card': 70, 'Other Current Liabilities': 80, 'Long Term Liabilities': 90,
    'Equity': 100,
    'Income': 110, 'Other Income': 120,
    'Cost of Goods Sold': 130, 'Expenses': 140, 'Other Expense': 150
  };
  var rows = coa.map(function(a) {
    var bal = balances[a['Account Name']] || 0;
    // Display sign: for debit-normal accounts, +ve means debit side; for
    // credit-normal, +ve means credit side. Match QB display convention —
    // positive balances on their natural side read as positive.
    var display = a['Normal Balance'] === 'Credit' ? -bal : bal;
    return {
      id:            a['ID'],
      accountNumber: a['Account Number'],
      name:          a['Account Name'],
      parent:        a['Parent Account Name'],
      type:          a['Type'],
      detailType:    a['Detail Type'],
      normalBalance: a['Normal Balance'],
      classification:_qbClassifyType_(a['Type']),
      active:        String(a['Active'] || 'Yes').toLowerCase() === 'yes',
      description:   a['Description'],
      plaidAccountIds: a['Plaid Account IDs'],
      balance:       display,
      typeSort:      typeOrder[a['Type']] || 999
    };
  });
  rows.sort(function(a, b) {
    if (a.typeSort !== b.typeSort) return a.typeSort - b.typeSort;
    return String(a.name).localeCompare(String(b.name));
  });
  return JSON.parse(JSON.stringify({ entity: entity, accounts: rows }));
}

// Trial balance as-of a date (or all time if null). Returns one row per
// account with a non-zero running balance, split into Debit / Credit columns
// so Σ Debit must equal Σ Credit.
function getQBTrialBalance(entity, asOfDate) {
  entity = String(entity || 'NF CA');
  var coa = _getQBRows_('QB_COA', QB_COA_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity; });
  var acctByName = {};
  coa.forEach(function(a) { acctByName[a['Account Name']] = a; });

  var balances = _computeAccountBalances_(entity, null, asOfDate);
  var typeOrder = {
    'Bank': 10, 'Accounts receivable (A/R)': 20, 'Other Current Assets': 30,
    'Fixed Assets': 40, 'Other Assets': 50,
    'Accounts payable (A/P)': 60, 'Credit Card': 70, 'Other Current Liabilities': 80, 'Long Term Liabilities': 90,
    'Equity': 100,
    'Income': 110, 'Other Income': 120,
    'Cost of Goods Sold': 130, 'Expenses': 140, 'Other Expense': 150
  };

  var rows = [];
  var totalDebit = 0, totalCredit = 0;
  Object.keys(balances).forEach(function(acct) {
    var bal = balances[acct];
    if (!bal || Math.abs(bal) < 0.005) return;   // skip $0 lines
    // bal is a SIGNED net (debits positive, credits negative). Split into
    // DR / CR columns matching how QB prints a TB.
    var dr = bal > 0 ? bal : 0;
    var cr = bal < 0 ? -bal : 0;
    totalDebit += dr; totalCredit += cr;
    var a = acctByName[acct] || {};
    rows.push({
      account:    acct,
      type:       a['Type'] || '',
      typeSort:   typeOrder[a['Type']] || 999,
      debit:      dr,
      credit:     cr
    });
  });
  rows.sort(function(a, b) {
    if (a.typeSort !== b.typeSort) return a.typeSort - b.typeSort;
    return String(a.account).localeCompare(String(b.account));
  });
  return JSON.parse(JSON.stringify({
    entity: entity,
    asOfDate: asOfDate || _todayIso_(),
    rows: rows,
    totalDebit: totalDebit,
    totalCredit: totalCredit,
    balanced: Math.abs(totalDebit - totalCredit) < 0.01
  }));
}

// Sum (Debit − Credit) per account from GL entries within [startIso, endIso].
// startIso = null → from inception. endIso = null → through today.
// Returns a map { 'Chase - 2086': 4680.14, 'Rental Income': -51096.00, ... }
// where a POSITIVE value means net debit, NEGATIVE means net credit.
function _computeAccountBalances_(entity, startIso, endIso) {
  var rows = _getQBRows_('QB_GL_ENTRIES', QB_GL_ENTRIES_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity; });
  var start = startIso ? new Date(startIso) : null;
  var end   = endIso   ? _endOfDay_(endIso) : null;
  var out = {};
  rows.forEach(function(r) {
    var d = r.Date instanceof Date ? r.Date : (r.Date ? new Date(r.Date) : null);
    if (!d) return;
    if (start && d < start) return;
    if (end && d > end) return;
    var acct = String(r['Account Name']);
    if (!acct) return;
    var dr = Number(r.Debit)  || 0;
    var cr = Number(r.Credit) || 0;
    out[acct] = (out[acct] || 0) + (dr - cr);
  });
  return out;
}

function _endOfDay_(iso) {
  var d = new Date(iso);
  d.setHours(23, 59, 59, 999);
  return d;
}

function _todayIso_() {
  var d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}


// ── Menu-callable debug ───────────────────────────────────────────────────

function debugQBNFCA() {
  var coa = getQBChartOfAccounts('NF CA');
  var tb  = getQBTrialBalance('NF CA', '2025-12-31');
  var lines = [
    'NF CA — QuickBooks Debug',
    '',
    'COA: ' + coa.accounts.length + ' accounts',
    '',
    'TRIAL BALANCE as of 2025-12-31',
    '  Total Debit:  $' + tb.totalDebit.toFixed(2),
    '  Total Credit: $' + tb.totalCredit.toFixed(2),
    '  Balanced:     ' + (tb.balanced ? '✓ YES' : '✗ NO'),
    ''
  ];
  tb.rows.forEach(function(r) {
    lines.push('  ' + _rpad_(r.account, 36) +
               (r.debit  ? '  DR $' + r.debit.toFixed(2)  : '              ') +
               (r.credit ? '  CR $' + r.credit.toFixed(2) : ''));
  });
  var out = lines.join('\n');
  try { SpreadsheetApp.getUi().alert('NF CA Debug', out, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
  Logger.log(out);
}

function _rpad_(s, n) {
  s = String(s);
  while (s.length < n) s += ' ';
  return s;
}
