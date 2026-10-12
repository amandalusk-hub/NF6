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

// Balance Sheet as-of a date. Pivots the trial balance into Asset / Liability
// / Equity sections. Computes a "Net Income" equity line = Σ(Revenue − Expense)
// within the fiscal year ending on asOfDate (same convention QBO uses). On
// 12/31 of the fiscal year, Net Income rolls into Retained Earnings via a
// closing JE (not automated here — matches how QB works).
function getQBBalanceSheet(entity, asOfDate) {
  entity = String(entity || 'NF CA');
  var coa = _getQBRows_('QB_COA', QB_COA_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity; });
  var acctByName = {};
  coa.forEach(function(a) { acctByName[a['Account Name']] = a; });

  var balances = _computeAccountBalances_(entity, null, asOfDate);

  // Group accounts by Asset/Liability/Equity classification. For BS display,
  // flip the sign on credit-normal accounts so a Liability with CR $5000
  // shows as $5,000 (not -$5,000).
  var sections = {
    Asset:     { subGroups: {}, total: 0 },
    Liability: { subGroups: {}, total: 0 },
    Equity:    { subGroups: {}, total: 0 }
  };
  var netIncome = 0;

  Object.keys(balances).forEach(function(acct) {
    var bal = balances[acct];
    var a = acctByName[acct];
    if (!a) return;
    var cls = _qbClassifyType_(a['Type']);
    if (cls === 'Revenue')  { netIncome -= bal; return; }   // credit-normal: more CR = more income
    if (cls === 'Expense')  { netIncome -= bal; return; }   // debit-normal: more DR = more expense (subtracts)
    if (cls !== 'Asset' && cls !== 'Liability' && cls !== 'Equity') return;
    if (!bal || Math.abs(bal) < 0.005) return;
    var sec = sections[cls];
    var sub = a['Type'];
    sec.subGroups[sub] = sec.subGroups[sub] || { total: 0, accounts: [] };
    // BS display convention: Assets show debit-positive; Liabilities and
    // Equity show credit-positive. That means debit-balanced equity accounts
    // (Draws, Distributions) show as NEGATIVE in the Equity section — same
    // as how QB prints "Distributions -241,396.75".
    var display = (cls === 'Asset') ? bal : -bal;
    sec.subGroups[sub].accounts.push({ name: acct, value: display });
    sec.subGroups[sub].total += display;
    sec.total += display;
  });
  // Net Income: Σ (Revenue CR − Expense DR). For Revenue, bal is NEGATIVE
  // (credit side), so -bal is a positive income number. For Expense, bal is
  // POSITIVE (debit side); we subtract it from net income. Combined:
  //   netIncome -= balRevenue  (-(-x) = +x)
  //   netIncome -= balExpense  (subtracts expense)
  // That gives netIncome = Σ revenues (as positive) − Σ expenses (as positive).

  // Add Net Income as an Equity line.
  if (Math.abs(netIncome) >= 0.005) {
    sections.Equity.subGroups['(Net Income)'] = { total: netIncome, accounts: [{ name: 'Net Income', value: netIncome }] };
    sections.Equity.total += netIncome;
  }

  // Sort sub-groups and accounts for display stability.
  function sortSection(sec) {
    var keys = Object.keys(sec.subGroups).sort();
    sec.ordered = keys.map(function(k) {
      sec.subGroups[k].accounts.sort(function(a, b) { return a.name.localeCompare(b.name); });
      return { label: k, total: sec.subGroups[k].total, accounts: sec.subGroups[k].accounts };
    });
  }
  sortSection(sections.Asset);
  sortSection(sections.Liability);
  sortSection(sections.Equity);

  return JSON.parse(JSON.stringify({
    entity: entity,
    asOfDate: asOfDate || _todayIso_(),
    assets: sections.Asset,
    liabilities: sections.Liability,
    equity: sections.Equity,
    netIncome: netIncome,
    balanced: Math.abs(sections.Asset.total - (sections.Liability.total + sections.Equity.total)) < 0.01
  }));
}

// Income Statement (P&L) for a period. Positive income values; expenses are
// subtracted. Net Income = Σ income − Σ expense.
function getQBIncomeStatement(entity, startIso, endIso) {
  entity = String(entity || 'NF CA');
  var coa = _getQBRows_('QB_COA', QB_COA_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity; });
  var acctByName = {};
  coa.forEach(function(a) { acctByName[a['Account Name']] = a; });

  var balances = _computeAccountBalances_(entity, startIso, endIso);

  var income = { subGroups: {}, total: 0 };
  var expense = { subGroups: {}, total: 0 };

  Object.keys(balances).forEach(function(acct) {
    var bal = balances[acct];
    if (!bal || Math.abs(bal) < 0.005) return;
    var a = acctByName[acct];
    if (!a) return;
    var cls = _qbClassifyType_(a['Type']);
    if (cls === 'Revenue') {
      // Credit-normal: display as positive = -bal
      var v = -bal;
      income.subGroups[a['Type']] = income.subGroups[a['Type']] || { total: 0, accounts: [] };
      income.subGroups[a['Type']].accounts.push({ name: acct, value: v });
      income.subGroups[a['Type']].total += v;
      income.total += v;
    } else if (cls === 'Expense') {
      // Debit-normal: display as positive = bal
      expense.subGroups[a['Type']] = expense.subGroups[a['Type']] || { total: 0, accounts: [] };
      expense.subGroups[a['Type']].accounts.push({ name: acct, value: bal });
      expense.subGroups[a['Type']].total += bal;
      expense.total += bal;
    }
  });

  function sortSection(sec) {
    var keys = Object.keys(sec.subGroups).sort();
    sec.ordered = keys.map(function(k) {
      sec.subGroups[k].accounts.sort(function(a, b) { return a.name.localeCompare(b.name); });
      return { label: k, total: sec.subGroups[k].total, accounts: sec.subGroups[k].accounts };
    });
  }
  sortSection(income);
  sortSection(expense);

  return JSON.parse(JSON.stringify({
    entity: entity,
    startDate: startIso || '',
    endDate: endIso || _todayIso_(),
    income: income,
    expense: expense,
    netIncome: income.total - expense.total
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


// ── Plaid sync for a QB entity ────────────────────────────────────────────
// For every Bank/Credit-Card COA account linked to a Plaid account ID, pull
// every Plaid txn since the entity's go-live date, apply QB_RULES, and post a
// double-entry GL row for each txn.
//
// Posting shape per txn:
//   Cash INFLOW  (+amount)  →  DR <bank account>  /  CR <income/other account>
//   Cash OUTFLOW (−amount)  →  DR <expense/other account>  /  CR <bank account>
//
// If no rule matches, the other leg defaults to "Ask My Accountant" (QB's own
// convention — same account Amanda sees in QB when a bank txn hasn't been
// categorized). Those appear in the Needs Review inbox.
//
// Idempotency: each GL row's Source Ref carries the Plaid Transaction ID. On
// re-run, txns already posted are skipped. This makes sync safe to run on a
// cron or on-demand from the UI.

var NFCA_GO_LIVE = '2025-12-31';
// Opening balance was posted AS OF 2025-12-31 23:59. Any Plaid txn posted
// on or before 2025-12-31 is implicitly covered by the opening JE — so sync
// starts strictly AFTER that.
var NFCA_SYNC_FROM = '2026-01-01';

// Web-callable (and menu-callable). Returns { bankAccounts, postedByAccount,
// skipped, needsReview } for the UI.
function syncQBEntityFromPlaid(entity) {
  _requireEditor_();
  ensureQBSheets_();
  entity = String(entity || 'NF CA');
  var now = new Date();
  var user = _currentUserEmail_();

  // 1. Find every bank/CC COA account for this entity that has a Plaid link.
  var coa = _getQBRows_('QB_COA', QB_COA_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity; });
  var bankAccounts = coa.filter(function(a) {
    var t = String(a['Type'] || '');
    return (t === 'Bank' || t === 'Credit Card') && String(a['Plaid Account IDs'] || '').trim();
  });
  if (bankAccounts.length === 0) {
    return { entity: entity, bankAccounts: 0, posted: 0, skipped: 0, needsReview: 0,
             error: 'No bank COA accounts have a Plaid Account ID yet. Run "Wire NF CA Plaid Accounts" first.' };
  }

  // 2. Load PLAID_TRANSACTIONS once (as objects).
  var ptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PLAID_TRANSACTIONS');
  if (!ptSheet || ptSheet.getLastRow() < 2) {
    return { entity: entity, bankAccounts: bankAccounts.length, posted: 0, skipped: 0, needsReview: 0,
             error: 'PLAID_TRANSACTIONS is empty. Run "Sync ALL Plaid Transactions" first.' };
  }
  var ptHdr = ptSheet.getRange(1, 1, 1, ptSheet.getLastColumn()).getValues()[0];
  var ptData = ptSheet.getRange(2, 1, ptSheet.getLastRow() - 1, ptHdr.length).getValues();
  var iId   = ptHdr.indexOf('Transaction ID');
  var iDate = ptHdr.indexOf('Date');
  var iAcct = ptHdr.indexOf('Account');
  var iAcctId = ptHdr.indexOf('Account ID');
  var iName = ptHdr.indexOf('Name');
  var iAmt  = ptHdr.indexOf('Amount USD');
  var iPending = ptHdr.indexOf('Pending');

  var fromDate = new Date(NFCA_SYNC_FROM + 'T00:00:00Z');

  // 3. Already-posted Plaid txn IDs for this entity, so we skip them.
  var gl = _getQBRows_('QB_GL_ENTRIES', QB_GL_ENTRIES_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity && String(r.Source) === 'plaid'; });
  var alreadyPosted = {};
  gl.forEach(function(r) {
    var ref = String(r['Source Ref'] || '').trim();
    if (ref) alreadyPosted[ref] = true;
  });

  // 4. Load rules for this entity, sorted by priority ascending.
  var rules = _getQBRows_('QB_RULES', QB_RULES_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity && String(r.Active || 'Yes').toLowerCase() === 'yes'; })
    .sort(function(a, b) { return (Number(a.Priority) || 999) - (Number(b.Priority) || 999); });

  // 5. For each bank account, find its Plaid txns and post double-entry rows.
  var posted = 0, skipped = 0, needsReview = 0;
  var postedByAccount = {};
  bankAccounts.forEach(function(bank) {
    var acctKey = bank['Account Name'];
    var plaidIds = String(bank['Plaid Account IDs'] || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    postedByAccount[acctKey] = { posted: 0, skipped: 0, needsReview: 0 };

    for (var r = 0; r < ptData.length; r++) {
      var row = ptData[r];
      var txnId = String(row[iId] || '');
      if (!txnId) continue;

      // Match by Plaid Account ID (preferred) OR by last-4 in the Account string.
      var pAcctId = String(row[iAcctId] || '');
      var pAcctStr = String(row[iAcct] || '');
      var matched = false;
      for (var p = 0; p < plaidIds.length; p++) {
        var pid = plaidIds[p];
        if (!pid) continue;
        if (pAcctId === pid) { matched = true; break; }
        // last-4 fallback: pid like '2086' matches "···2086" or "(2086)" in label.
        if (/^\d{3,4}$/.test(pid) && pAcctStr.indexOf(pid) >= 0) { matched = true; break; }
      }
      if (!matched) continue;

      // Date filter: strictly after go-live.
      var d = row[iDate] instanceof Date ? row[iDate] : (row[iDate] ? new Date(row[iDate]) : null);
      if (!d || d < fromDate) continue;
      // Skip pending — they'll re-post with a different ID when they clear.
      if (String(row[iPending] || '').toLowerCase() === 'yes') continue;

      if (alreadyPosted[txnId]) { skipped++; postedByAccount[acctKey].skipped++; continue; }

      var amt = Number(row[iAmt]) || 0;
      if (!amt) { skipped++; postedByAccount[acctKey].skipped++; continue; }
      var name = String(row[iName] || '');

      // Rule hit: returns { category: 'DR' | 'CR' | 'both', drAccount, crAccount, memo }.
      var hit = _qbCategorizePlaidTxn_({ amount: amt, name: name, account: pAcctStr, accountId: pAcctId }, rules);
      var otherAccount, isReview = false;
      // Only use the two-sided path when the rule names TWO DIFFERENT accounts.
      // (addQBRule stores the SAME category on both legs for a direction='both'
      // rule — in that case fall through to single-sided so the bank side gets
      // resolved automatically from the txn sign.)
      if (hit && hit.drAccount && hit.crAccount && hit.drAccount !== hit.crAccount) {
        _postJE_(entity, d, hit.memo || name, [
          { account: hit.drAccount, debit: Math.abs(amt), credit: 0 },
          { account: hit.crAccount, debit: 0, credit: Math.abs(amt) }
        ], 'plaid', txnId, user, now);
        posted++; postedByAccount[acctKey].posted++;
        continue;
      }
      // Single-sided or no rule: the bank side is determined by sign; the
      // other side is either the rule's single account OR Ask My Accountant.
      otherAccount = hit && (hit.drAccount || hit.crAccount) ? (hit.drAccount || hit.crAccount) : 'Ask My Accountant';
      if (otherAccount === 'Ask My Accountant') { needsReview++; postedByAccount[acctKey].needsReview++; isReview = true; }

      var lines;
      if (amt > 0) {
        // Cash in: DR Bank / CR Other
        lines = [
          { account: acctKey,      debit: amt, credit: 0 },
          { account: otherAccount, debit: 0,   credit: amt }
        ];
      } else {
        // Cash out: DR Other / CR Bank
        lines = [
          { account: otherAccount, debit: Math.abs(amt), credit: 0 },
          { account: acctKey,      debit: 0,            credit: Math.abs(amt) }
        ];
      }
      _postJE_(entity, d, name, lines, 'plaid', txnId, user, now);
      posted++; postedByAccount[acctKey].posted++;
    }
  });

  return {
    entity: entity,
    bankAccounts: bankAccounts.length,
    posted: posted,
    skipped: skipped,
    needsReview: needsReview,
    postedByAccount: postedByAccount
  };
}

// Apply rules in priority order to a Plaid txn and return the first hit.
// Rule shape: { Match Plaid Account, Match Name Contains, Match Amount,
//               Direction, DR Account, CR Account, Memo Template }.
function _qbCategorizePlaidTxn_(txn, rules) {
  for (var i = 0; i < rules.length; i++) {
    var r = rules[i];
    var mAcct = String(r['Match Plaid Account'] || '').toLowerCase().trim();
    var mName = String(r['Match Name Contains'] || '').toLowerCase().trim();
    var mAmt  = String(r['Match Amount'] || '').trim();
    var dir   = String(r['Direction'] || 'both').toLowerCase();

    if (dir === 'in'  && txn.amount < 0) continue;
    if (dir === 'out' && txn.amount > 0) continue;
    if (mAcct && String(txn.account || '').toLowerCase().indexOf(mAcct) < 0 && String(txn.accountId || '') !== mAcct) continue;
    if (mName && String(txn.name || '').toLowerCase().indexOf(mName) < 0) continue;
    if (mAmt) {
      var want = mAmt.split('|').map(function(s) { return Number(s.trim()); });
      var hit = want.some(function(v) { return Math.abs(v - txn.amount) < 0.005; });
      if (!hit) continue;
    }
    return {
      drAccount: String(r['DR Account'] || '').trim(),
      crAccount: String(r['CR Account'] || '').trim(),
      memo: String(r['Memo Template'] || '').trim()
    };
  }
  return null;
}

// Post a journal entry — assigns an Entry ID, writes each line, validates
// Σ debit = Σ credit before writing (throws on imbalance). Returns the Entry
// ID. All lines of one entry share the same date, memo, source, and source
// ref so later queries (GL, reconciliation) can group them.
function _postJE_(entity, date, memo, lines, source, sourceRef, user, now) {
  var dr = 0, cr = 0;
  lines.forEach(function(l) {
    dr += Number(l.debit)  || 0;
    cr += Number(l.credit) || 0;
  });
  if (Math.abs(dr - cr) > 0.005) {
    throw new Error('JE out of balance: DR ' + dr.toFixed(2) + ' vs CR ' + cr.toFixed(2) + ' (' + memo + ')');
  }
  var entryId = 'je-' + Utilities.getUuid().substring(0, 12);
  lines.forEach(function(l, i) {
    _writeQBRow_('QB_GL_ENTRIES', QB_GL_ENTRIES_HEADERS, {
      'Entry ID':         entryId,
      'Line #':           i + 1,
      'Date':             date,
      'Entity':           entity,
      'Memo':             memo || '',
      'Account Name':     l.account,
      'Debit':            l.debit  || '',
      'Credit':           l.credit || '',
      'Source':           source,
      'Source Ref':       sourceRef || '',
      'Line Memo':        l.memo || '',
      'Reconciled Date':  '',
      'Reconciliation Ref': '',
      'Created By':       user,
      'Created At':       now
    });
  });
  return entryId;
}


// ── Plaid wiring helpers ──────────────────────────────────────────────────
// Attach a Plaid Account ID (or last-4) to a COA bank/CC account so sync
// knows which Plaid txns belong to this entity.

function setQBAccountPlaidIds(entity, accountName, plaidIdsCsv) {
  _requireEditor_();
  ensureQBSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QB_COA');
  var lastCol = Math.max(sheet.getLastColumn(), QB_COA_HEADERS.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var iEnt = hdr.indexOf('Entity');
  var iName = hdr.indexOf('Account Name');
  var iPid = hdr.indexOf('Plaid Account IDs');
  var iUpd = hdr.indexOf('Last Updated');
  if (iEnt < 0 || iName < 0 || iPid < 0) throw new Error('QB_COA headers missing');
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][iEnt]) !== String(entity)) continue;
    if (String(data[r][iName]) !== String(accountName)) continue;
    var cell = sheet.getRange(r + 2, iPid + 1);
    cell.setNumberFormat('@');
    cell.setValue(String(plaidIdsCsv || ''));
    if (iUpd >= 0) sheet.getRange(r + 2, iUpd + 1).setValue(new Date());
    return { success: true };
  }
  throw new Error('COA account not found: ' + entity + ' / ' + accountName);
}

// Menu-callable: auto-detect Chase 2086 in PLAID_TRANSACTIONS and wire it to
// NF CA's "Chase - 2086" COA account. If multiple candidates match the mask,
// list them and let Amanda pick.
function wireNFCAPlaidAccounts() {
  _requireEditor_();
  ensureQBSheets_();
  var ptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PLAID_TRANSACTIONS');
  if (!ptSheet || ptSheet.getLastRow() < 2) {
    try { SpreadsheetApp.getUi().alert('PLAID_TRANSACTIONS is empty. Run "Sync ALL Plaid Transactions" first.'); } catch(e) {}
    return;
  }
  var hdr = ptSheet.getRange(1, 1, 1, ptSheet.getLastColumn()).getValues()[0];
  var iAcct = hdr.indexOf('Account');
  var iAcctId = hdr.indexOf('Account ID');
  var data = ptSheet.getRange(2, 1, ptSheet.getLastRow() - 1, hdr.length).getValues();
  var distinct = {};
  for (var r = 0; r < data.length; r++) {
    var label = String(data[r][iAcct] || '');
    var id = String(data[r][iAcctId] || '');
    if (!label || !id) continue;
    distinct[id] = distinct[id] || { label: label, id: id, count: 0 };
    distinct[id].count++;
  }
  // Find Chase 2086 by mask substring. Prefer accounts with 'chase' in the label.
  var candidates = Object.keys(distinct).map(function(k) { return distinct[k]; })
    .filter(function(a) { return a.label.indexOf('2086') >= 0; })
    .sort(function(a, b) {
      var ac = /chase/i.test(a.label) ? 0 : 1;
      var bc = /chase/i.test(b.label) ? 0 : 1;
      if (ac !== bc) return ac - bc;
      return b.count - a.count;
    });
  var ui = SpreadsheetApp.getUi();
  if (candidates.length === 0) {
    ui.alert('No Plaid account found with "2086" in its label.\n\nAll accounts in PLAID_TRANSACTIONS:\n' +
             Object.keys(distinct).map(function(k) { return '  ' + distinct[k].label; }).join('\n'));
    return;
  }
  if (candidates.length > 1) {
    var msg = 'Multiple candidates for Chase 2086:\n\n' +
              candidates.map(function(c, i) { return (i + 1) + '. ' + c.label + '  (' + c.count + ' txns)\n   ID: ' + c.id; }).join('\n\n') +
              '\n\nUsing the first one (' + candidates[0].label + '). Edit QB_COA → Chase - 2086 → Plaid Account IDs manually if that is wrong.';
    ui.alert('Chase 2086 — Multiple Matches', msg, ui.ButtonSet.OK);
  }
  var pick = candidates[0];
  setQBAccountPlaidIds('NF CA', 'Chase - 2086', pick.id);
  ui.alert('Wired Chase 2086', 'NF CA → Chase - 2086 is now linked to:\n' + pick.label + '\n(ID: ' + pick.id + ')\n' + pick.count + ' Plaid txns on this account.\n\nNext: run "QuickBooks → Sync NF CA from Plaid" to post them.', ui.ButtonSet.OK);
}

// Daily-cron entry point: iterate every QB-active entity (anything that has
// at least one row in QB_COA), sync Plaid, then re-apply active rules.
// Called from Code.gs dailySync_. Non-fatal — exceptions logged, not thrown.
function dailySyncQBEntities_() {
  ensureQBSheets_();
  var coa = _getQBRows_('QB_COA', QB_COA_HEADERS);
  var seen = {};
  coa.forEach(function(r) { if (r.Entity) seen[r.Entity] = true; });
  var entities = Object.keys(seen);
  if (!entities.length) {
    Logger.log('dailySyncQB: no QB entities seeded yet, skipping');
    return;
  }
  entities.forEach(function(entity) {
    try {
      var sync = syncQBEntityFromPlaid(entity);
      Logger.log('dailySyncQB ' + entity + ' sync: ' + JSON.stringify(sync));
    } catch (e) {
      Logger.log('dailySyncQB ' + entity + ' sync FAILED: ' + e.message);
    }
    // After sync, re-sweep every Ask My Accountant txn through active rules —
    // catches cases where a rule was added yesterday and today's new txns
    // (or any older ones that stayed in Ask My Accountant) should now auto-code.
    try {
      var r = reapplyAllQBRules(entity);
      if (r.updated) Logger.log('dailySyncQB ' + entity + ' re-applied rules: ' + r.updated);
    } catch (e) {
      Logger.log('dailySyncQB ' + entity + ' reapply FAILED: ' + e.message);
    }
  });
}

// Menu-callable: run the sync for NF CA + show a summary dialog.
function syncQBNFCAFromPlaidMenu() {
  var result = syncQBEntityFromPlaid('NF CA');
  var lines = [
    'NF CA Plaid Sync',
    '',
    'Bank accounts linked: ' + result.bankAccounts
  ];
  if (result.error) lines.push('', '⚠ ' + result.error);
  else {
    lines.push('Posted new GL entries: ' + result.posted);
    lines.push('Needs review (→ Ask My Accountant): ' + result.needsReview);
    lines.push('Skipped (already posted): ' + result.skipped);
    if (result.postedByAccount) {
      lines.push('');
      Object.keys(result.postedByAccount).forEach(function(acct) {
        var pb = result.postedByAccount[acct];
        lines.push('  ' + acct + ': ' + pb.posted + ' posted · ' + pb.needsReview + ' review · ' + pb.skipped + ' skipped');
      });
    }
  }
  try { SpreadsheetApp.getUi().alert('NF CA Plaid Sync', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
  Logger.log(lines.join('\n'));
}


// ── Banking view (QB "For Review" style) ──────────────────────────────────
// Returns the live Plaid feed for every bank account linked to this entity,
// with each txn tagged as either POSTED (and to which account) or UNPOSTED.
// This is what the Banking tab renders — Amanda classifies inline from here
// instead of running a batch sync. (Batch sync still exists for mass-ops.)

function getQBBankingFeed(entity) {
  entity = String(entity || 'NF CA');
  ensureQBSheets_();

  // 1. Linked bank COA accounts for this entity.
  var coa = _getQBRows_('QB_COA', QB_COA_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity; });
  var bankAccounts = coa.filter(function(a) {
    var t = String(a['Type'] || '');
    return (t === 'Bank' || t === 'Credit Card') && String(a['Plaid Account IDs'] || '').trim();
  });
  if (bankAccounts.length === 0) {
    return { entity: entity, bankAccounts: [], txns: [], needsWiring: true };
  }

  // Build a lookup { plaidAccountId or last4 → COA bank account name }.
  var plaidIdToBank = {};
  var last4ToBank = {};
  bankAccounts.forEach(function(a) {
    String(a['Plaid Account IDs']).split(',').forEach(function(idRaw) {
      var id = idRaw.trim();
      if (!id) return;
      if (/^\d{3,4}$/.test(id)) last4ToBank[id] = a['Account Name'];
      else plaidIdToBank[id] = a['Account Name'];
    });
  });

  // 2. Load PLAID_TRANSACTIONS, filter to these accounts, sort newest-first.
  var ptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PLAID_TRANSACTIONS');
  if (!ptSheet || ptSheet.getLastRow() < 2) {
    return { entity: entity, bankAccounts: bankAccounts.map(_qbBankCard_), txns: [] };
  }
  var ptHdr = ptSheet.getRange(1, 1, 1, ptSheet.getLastColumn()).getValues()[0];
  var ptData = ptSheet.getRange(2, 1, ptSheet.getLastRow() - 1, ptHdr.length).getValues();
  var iId   = ptHdr.indexOf('Transaction ID');
  var iDate = ptHdr.indexOf('Date');
  var iAcct = ptHdr.indexOf('Account');
  var iAcctId = ptHdr.indexOf('Account ID');
  var iName = ptHdr.indexOf('Name');
  var iMerch= ptHdr.indexOf('Merchant');
  var iAmt  = ptHdr.indexOf('Amount USD');
  var iPending = ptHdr.indexOf('Pending');
  var iCategory = ptHdr.indexOf('Plaid Category');
  var fromDate = new Date(NFCA_SYNC_FROM + 'T00:00:00Z');

  var txns = [];
  for (var r = 0; r < ptData.length; r++) {
    var row = ptData[r];
    var txnId = String(row[iId] || '');
    if (!txnId) continue;
    var pAcctId = String(row[iAcctId] || '');
    var pAcctStr = String(row[iAcct] || '');
    var bankName = plaidIdToBank[pAcctId];
    if (!bankName) {
      // last-4 fallback
      Object.keys(last4ToBank).forEach(function(k) {
        if (!bankName && pAcctStr.indexOf(k) >= 0) bankName = last4ToBank[k];
      });
    }
    if (!bankName) continue;
    var d = row[iDate] instanceof Date ? row[iDate] : (row[iDate] ? new Date(row[iDate]) : null);
    if (!d || d < fromDate) continue;
    txns.push({
      plaidId:   txnId,
      date:      Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      name:      String(row[iName] || ''),
      merchant:  iMerch >= 0 ? String(row[iMerch] || '') : '',
      amount:    Number(row[iAmt]) || 0,
      pending:   String(row[iPending] || '').toLowerCase() === 'yes',
      plaidCategory: iCategory >= 0 ? String(row[iCategory] || '') : '',
      bankAccount: bankName
    });
  }
  txns.sort(function(a, b) { return b.date.localeCompare(a.date); });

  // 3. Already-posted GL entries for these Plaid txns (source=plaid).
  var gl = _getQBRows_('QB_GL_ENTRIES', QB_GL_ENTRIES_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity && String(r.Source) === 'plaid'; });
  // Group by Plaid txn ID (source ref). For each, record { entryId, category,
  // bankAccount, memo } where category = the non-bank leg's account.
  var postedByPlaidId = {};
  gl.forEach(function(r) {
    var ref = String(r['Source Ref'] || '');
    if (!ref) return;
    postedByPlaidId[ref] = postedByPlaidId[ref] || { entryId: r['Entry ID'], lines: [] };
    postedByPlaidId[ref].lines.push({
      account: r['Account Name'],
      debit:   Number(r.Debit)  || 0,
      credit:  Number(r.Credit) || 0,
      memo:    r['Memo']
    });
  });

  // Build the per-plaid-id → bank + category map, and a history map of
  // {merchant hint → most common category} for the suggestion engine.
  var historyCounts = {};   // hint → { category: count }
  txns.forEach(function(t) {
    var p = postedByPlaidId[t.plaidId];
    if (!p) { t.posted = false; return; }
    t.posted = true;
    t.entryId = p.entryId;
    var other = p.lines.find(function(l) { return l.account !== t.bankAccount; });
    t.category = other ? other.account : '';
    t.memo     = (p.lines[0] && p.lines[0].memo) || '';
    t.needsReview = t.category === 'Ask My Accountant';
    if (!t.needsReview && t.category) {
      var hint = _qbExtractMerchantHint_(t.name).toLowerCase();
      if (hint) {
        historyCounts[hint] = historyCounts[hint] || {};
        historyCounts[hint][t.category] = (historyCounts[hint][t.category] || 0) + 1;
      }
    }
  });
  var historyMap = {};
  Object.keys(historyCounts).forEach(function(h) {
    var best = null, bestN = 0;
    Object.keys(historyCounts[h]).forEach(function(cat) {
      if (historyCounts[h][cat] > bestN) { best = cat; bestN = historyCounts[h][cat]; }
    });
    if (best) historyMap[h] = best;
  });

  // Load active rules once for the suggestion engine.
  var rules = _getQBRows_('QB_RULES', QB_RULES_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity && String(r.Active || 'Yes').toLowerCase() === 'yes'; })
    .sort(function(a, b) { return (Number(a.Priority) || 999) - (Number(b.Priority) || 999); });

  // Tag each txn with its extracted merchant hint + category suggestion.
  // (The hint is also the default rule-matcher text in the create-rule modal.)
  txns.forEach(function(t) {
    t.merchantHint = _qbExtractMerchantHint_(t.name);
    if (!t.posted || t.needsReview) {
      var sug = _qbSuggestCategory_(t, rules, historyMap);
      if (sug) t.suggestion = sug;
    }
  });

  return JSON.parse(JSON.stringify({
    entity: entity,
    bankAccounts: bankAccounts.map(_qbBankCard_),
    txns: txns
  }));
}

function _qbBankCard_(a) {
  return {
    name: a['Account Name'],
    type: a['Type'],
    plaidAccountIds: a['Plaid Account IDs']
  };
}

// Post ONE Plaid txn as a 2-leg JE. Idempotent via Source Ref — if already
// posted, returns the existing entry id without duplicating.
function postQBPlaidTxn(entity, plaidTxnId, categoryAccount, memoOverride) {
  _requireEditor_();
  ensureQBSheets_();
  entity = String(entity || 'NF CA');
  if (!plaidTxnId) throw new Error('plaidTxnId required');
  if (!categoryAccount) throw new Error('categoryAccount required');

  // Skip if already posted.
  var existing = _getQBRows_('QB_GL_ENTRIES', QB_GL_ENTRIES_HEADERS)
    .filter(function(r) {
      return String(r.Entity) === entity && String(r.Source) === 'plaid' && String(r['Source Ref']) === String(plaidTxnId);
    });
  if (existing.length) {
    return { success: true, alreadyPosted: true, entryId: existing[0]['Entry ID'] };
  }

  // Look up the Plaid txn.
  var ptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PLAID_TRANSACTIONS');
  if (!ptSheet || ptSheet.getLastRow() < 2) throw new Error('PLAID_TRANSACTIONS is empty');
  var ptHdr = ptSheet.getRange(1, 1, 1, ptSheet.getLastColumn()).getValues()[0];
  var ptData = ptSheet.getRange(2, 1, ptSheet.getLastRow() - 1, ptHdr.length).getValues();
  var iId = ptHdr.indexOf('Transaction ID');
  var iDate = ptHdr.indexOf('Date');
  var iAcct = ptHdr.indexOf('Account');
  var iAcctId = ptHdr.indexOf('Account ID');
  var iName = ptHdr.indexOf('Name');
  var iAmt = ptHdr.indexOf('Amount USD');
  var txn = null;
  for (var r = 0; r < ptData.length; r++) {
    if (String(ptData[r][iId]) === String(plaidTxnId)) {
      txn = {
        date:      ptData[r][iDate],
        account:   String(ptData[r][iAcct]),
        accountId: String(ptData[r][iAcctId]),
        name:      String(ptData[r][iName]),
        amount:    Number(ptData[r][iAmt]) || 0
      };
      break;
    }
  }
  if (!txn) throw new Error('Plaid txn not found: ' + plaidTxnId);

  // Resolve the bank COA account for this entity.
  var coa = _getQBRows_('QB_COA', QB_COA_HEADERS)
    .filter(function(a) { return String(a.Entity) === entity; });
  var bank = coa.find(function(a) {
    var t = String(a['Type'] || '');
    if (t !== 'Bank' && t !== 'Credit Card') return false;
    var ids = String(a['Plaid Account IDs'] || '').split(',').map(function(s) { return s.trim(); });
    if (ids.indexOf(txn.accountId) >= 0) return true;
    // last-4 fallback
    return ids.some(function(id) { return /^\d{3,4}$/.test(id) && txn.account.indexOf(id) >= 0; });
  });
  if (!bank) throw new Error('No bank COA account linked to this Plaid txn for ' + entity);

  var dt = txn.date instanceof Date ? txn.date : new Date(txn.date);
  var amt = txn.amount;
  var lines;
  if (amt > 0) {
    lines = [
      { account: bank['Account Name'], debit: amt, credit: 0 },
      { account: categoryAccount,      debit: 0,   credit: amt }
    ];
  } else {
    lines = [
      { account: categoryAccount,      debit: Math.abs(amt), credit: 0 },
      { account: bank['Account Name'], debit: 0,             credit: Math.abs(amt) }
    ];
  }
  var entryId = _postJE_(entity, dt, memoOverride || txn.name, lines, 'plaid', plaidTxnId, _currentUserEmail_(), new Date());
  return { success: true, entryId: entryId };
}

// Reclassify an already-posted Plaid txn to a different category. Rewrites
// the non-bank leg of the existing JE in place (keeps the Entry ID + audit).
function reclassifyQBPlaidTxn(entity, plaidTxnId, newCategoryAccount) {
  _requireEditor_();
  ensureQBSheets_();
  entity = String(entity || 'NF CA');
  if (!plaidTxnId || !newCategoryAccount) throw new Error('plaidTxnId and newCategoryAccount required');

  // Find bank COA account to identify which leg to leave alone.
  var coa = _getQBRows_('QB_COA', QB_COA_HEADERS)
    .filter(function(a) { return String(a.Entity) === entity; });
  var bankNames = {};
  coa.forEach(function(a) {
    var t = String(a['Type'] || '');
    if (t === 'Bank' || t === 'Credit Card') bankNames[a['Account Name']] = true;
  });

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QB_GL_ENTRIES');
  var lastCol = Math.max(sheet.getLastColumn(), QB_GL_ENTRIES_HEADERS.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var iEnt = hdr.indexOf('Entity');
  var iSrc = hdr.indexOf('Source');
  var iRef = hdr.indexOf('Source Ref');
  var iAcc = hdr.indexOf('Account Name');
  var iUpd = hdr.indexOf('Created At');   // repurpose as "last updated" for simplicity
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  var updated = 0;
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][iEnt]) !== entity) continue;
    if (String(data[r][iSrc]) !== 'plaid') continue;
    if (String(data[r][iRef]) !== String(plaidTxnId)) continue;
    var acct = String(data[r][iAcc]);
    if (bankNames[acct]) continue;   // leave the bank leg alone
    sheet.getRange(r + 2, iAcc + 1).setValue(newCategoryAccount);
    if (iUpd >= 0) sheet.getRange(r + 2, iUpd + 1).setValue(new Date());
    updated++;
  }
  if (!updated) throw new Error('No non-bank GL leg found for Plaid txn ' + plaidTxnId);
  return { success: true, updated: updated };
}

// Extract the merchant-ish key from a Plaid txn name, used as the default
// rule matcher (QBO's "when Description contains X" style). Handles the
// common Chase ACH formats first; falls back to the first few significant
// words for anything else. This is suggestion fodder — Amanda can always
// edit the matcher before saving a rule.
function _qbExtractMerchantHint_(plaidName) {
  var s = String(plaidName || '').trim();
  if (!s) return '';
  // ACH credit/debit: "ORIG CO NAME:ELLISON MEDICAL ORIG ID:... DESC..."
  var m = s.match(/ORIG\s+CO\s+NAME:\s*([^\s][^]*?)\s+(?:ORIG\s+ID|DESC|ENTRY|CO\s+ID|$)/i);
  if (m && m[1]) return m[1].trim().replace(/\s+/g, ' ');
  // Zelle / Venmo / Cash App
  m = s.match(/(?:ZELLE|VENMO|CASH\s*APP|PAYPAL)\s+(?:TO|FROM|PAYMENT)[\s:-]+([A-Z][A-Za-z0-9 .&'-]{2,40})/i);
  if (m && m[1]) return m[1].trim().replace(/\s+/g, ' ');
  // Wire In/Out: "WIRE OUT TO <NAME>" or "FEDWIRE CREDIT ... FROM <NAME>"
  m = s.match(/WIRE\s+(?:IN|OUT|CREDIT|DEBIT)[\s:-]*(?:TO|FROM)?\s+([A-Z][A-Za-z0-9 .&'-]{2,40})/i);
  if (m && m[1]) return m[1].trim().replace(/\s+/g, ' ');
  // Online transfers: unique by the destination account tail (e.g. CHK ...2001).
  // Keep the bank-type prefix (CHK, SAV, etc.) in the hint so the matcher
  // finds it as a literal substring in Plaid's actual description.
  m = s.match(/online\s+transfer\s+to\s+([a-z]+)\s+(?:\.{3}|\.\.\.)?(\d{3,4})/i);
  if (m) return 'Online Transfer to ' + m[1].toUpperCase() + ' ...' + m[2];
  m = s.match(/online\s+transfer\s+to\s+(?:\.{3}|\.\.\.)?(\d{3,4})/i);
  if (m) return 'Online Transfer to ...' + m[1];
  m = s.match(/online\s+transfer\s+from\s+([a-z]+)\s+(?:\.{3}|\.\.\.)?(\d{3,4})/i);
  if (m) return 'Online Transfer from ' + m[1].toUpperCase() + ' ...' + m[2];
  m = s.match(/online\s+transfer\s+from\s+(?:\.{3}|\.\.\.)?(\d{3,4})/i);
  if (m) return 'Online Transfer from ...' + m[1];
  // Debit card: "PURCHASE AUTH ... <MERCHANT> <CITY> <STATE>"
  m = s.match(/(?:PURCHASE|DEBIT\s+CARD)\s+(?:AUTH)?[\s:#\d-]+([A-Z][A-Za-z0-9 .&'-]{3,40})/i);
  if (m && m[1]) return m[1].trim().replace(/\s+/g, ' ');
  // Default: first 3 significant words, strip long numeric runs
  var cleaned = s.replace(/\b\d{5,}\b/g, '').replace(/\s+/g, ' ').trim();
  return cleaned.split(/\s+/).slice(0, 4).join(' ');
}

// Does this Plaid txn match this rule? Mirrors _qbCategorizePlaidTxn_'s
// filter logic but takes a single rule — exposed for the suggestion engine.
//
// Name matching: multi-token (word-order) so a rule like "Online Transfer
// to ...2001" matches "Online Transfer to CHK ...2001 transaction#:..." —
// each whitespace-separated token must appear in order with any text in
// between (wildcard). Single-token patterns behave like a plain substring
// contains, same as before.
function _qbRuleMatches_(txn, rule) {
  var mAcct = String(rule['Match Plaid Account'] || '').toLowerCase().trim();
  var mName = String(rule['Match Name Contains'] || '').trim();
  var mAmt  = String(rule['Match Amount'] || '').trim();
  var dir   = String(rule['Direction'] || 'both').toLowerCase();
  if (dir === 'in'  && txn.amount < 0) return false;
  if (dir === 'out' && txn.amount > 0) return false;
  if (mAcct && String(txn.account || '').toLowerCase().indexOf(mAcct) < 0 && String(txn.accountId || '') !== mAcct) return false;
  if (mName) {
    var tokens = mName.split(/\s+/).filter(Boolean);
    var hay = String(txn.name || '');
    var pos = 0;
    for (var i = 0; i < tokens.length; i++) {
      var t = tokens[i];
      var found = hay.toLowerCase().indexOf(t.toLowerCase(), pos);
      if (found < 0) return false;
      pos = found + t.length;
    }
  }
  if (mAmt) {
    var want = mAmt.split('|').map(function(s) { return Number(s.trim()); });
    var hit = want.some(function(v) { return Math.abs(v - txn.amount) < 0.005; });
    if (!hit) return false;
  }
  return true;
}

// Re-apply ALL active rules for this entity to every posted-to-Ask-My-Accountant
// txn. Useful when a rule has been updated or when the matcher logic changes
// (and existing posts should retro-code). Returns { updated, rulesSeen }.
function reapplyAllQBRules(entity) {
  _requireEditor_();
  ensureQBSheets_();
  entity = String(entity || 'NF CA');
  var rules = _getQBRows_('QB_RULES', QB_RULES_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity && String(r.Active || 'Yes').toLowerCase() === 'yes'; })
    .sort(function(a, b) { return (Number(a.Priority) || 999) - (Number(b.Priority) || 999); });
  if (!rules.length) return { entity: entity, updated: 0, rulesSeen: 0 };
  var feed = getQBBankingFeed(entity);
  var updated = 0;
  (feed.txns || []).forEach(function(t) {
    if (!t.posted || !t.needsReview) return;   // only sweep Ask My Accountant
    for (var i = 0; i < rules.length; i++) {
      if (!_qbRuleMatches_(t, rules[i])) continue;
      // Pick the right side of the rule based on txn direction.
      var dr = String(rules[i]['DR Account'] || '').trim();
      var cr = String(rules[i]['CR Account'] || '').trim();
      var cat = t.amount > 0 ? (cr || dr) : (dr || cr);
      if (!cat) continue;
      try {
        reclassifyQBPlaidTxn(entity, t.plaidId, cat);
        updated++;
      } catch (e) {
        Logger.log('reapply skipped ' + t.plaidId + ': ' + e.message);
      }
      break;   // first matching rule wins
    }
  });
  return { entity: entity, updated: updated, rulesSeen: rules.length };
}

// Suggest a category for a single Plaid txn by scanning rules + prior
// categorizations. Returns { category, source: 'rule'|'history', hint, ruleId }.
function _qbSuggestCategory_(txn, rules, historyMap) {
  // 1. Rule hit wins
  for (var i = 0; i < rules.length; i++) {
    if (_qbRuleMatches_(txn, rules[i])) {
      var dr = String(rules[i]['DR Account'] || '').trim();
      var cr = String(rules[i]['CR Account'] || '').trim();
      // Choose the non-bank leg from the rule (if two-sided).
      // For simplicity: inflow → the CR account is the suggestion; outflow → DR.
      var cat = txn.amount > 0 ? cr : dr;
      if (!cat) cat = dr || cr;
      return { category: cat, source: 'rule', hint: rules[i]['Match Name Contains'], ruleId: rules[i].ID };
    }
  }
  // 2. History: did a prior txn with the same merchant hint get categorized?
  if (historyMap) {
    var hint = _qbExtractMerchantHint_(txn.name).toLowerCase();
    if (hint && historyMap[hint] && historyMap[hint] !== 'Ask My Accountant') {
      return { category: historyMap[hint], source: 'history', hint: hint, ruleId: '' };
    }
  }
  return null;
}

// Public: create a bank rule, then optionally auto-post every unposted Plaid
// txn that matches it. Returns { ruleId, applied } where applied is the
// number of txns that got posted as a result.
function addQBRule(entity, matcherConfig, applyToUnposted) {
  _requireEditor_();
  ensureQBSheets_();
  entity = String(entity || 'NF CA');
  matcherConfig = matcherConfig || {};
  var cat = String(matcherConfig.category || '').trim();
  if (!cat) throw new Error('category is required');
  var dir = String(matcherConfig.direction || 'both').toLowerCase();
  var isIn = dir === 'in' || (dir === 'both' && Number(matcherConfig.sampleAmount) > 0);

  // Pick DR/CR from category + inferred direction. For an INflow, bank is DR,
  // category is CR. For an OUTflow, category is DR, bank is CR. Store both
  // directions in the rule so the matcher doesn't have to guess at sync time.
  // (The sync picks the right leg based on sign.) We store the category on
  // BOTH legs so the single-sided fallback in syncQBEntityFromPlaid picks
  // the category regardless of sign — then direction filter narrows which
  // txns the rule applies to.
  var drAcct, crAcct;
  if (dir === 'in') { drAcct = ''; crAcct = cat; }
  else if (dir === 'out') { drAcct = cat; crAcct = ''; }
  else {
    // 'both': store category on both sides. Sync picks the opposite side of
    // whichever direction the actual txn is.
    drAcct = cat; crAcct = cat;
  }

  // Priority: later rules match last → put each new rule at the back of its
  // priority band so earlier user-created rules keep precedence by order.
  var existingRules = _getQBRows_('QB_RULES', QB_RULES_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity; });
  var nextPriority = 100 + existingRules.length;

  var now = new Date();
  var ruleId = 'qbrule-' + Utilities.getUuid().substring(0, 10);
  _writeQBRow_('QB_RULES', QB_RULES_HEADERS, {
    'ID':                   ruleId,
    'Entity':               entity,
    'Priority':             nextPriority,
    'Match Plaid Account':  String(matcherConfig.matchPlaidAccount || ''),
    'Match Name Contains':  String(matcherConfig.matchName || ''),
    'Match Amount':         String(matcherConfig.matchAmount || ''),
    'Direction':            dir,
    'DR Account':           drAcct,
    'CR Account':           crAcct,
    'Memo Template':        String(matcherConfig.memoTemplate || ''),
    'Active':               'Yes',
    'Date Added':           now,
    'Last Updated':         now
  });

  // Optionally sweep every matching Plaid txn and apply the rule's category.
  // Three cases to handle:
  //   1. Unposted         → post a fresh JE with the category.
  //   2. Posted + needs review (Ask My Accountant) → reclassify in place.
  //   3. Posted to a real category → skip (don't overwrite a prior decision).
  // Case 2 is the common one here because the Banking view auto-syncs on
  // open — every txn ends up posted to Ask My Accountant before Amanda
  // ever touches it.
  var applied = 0;
  if (applyToUnposted) {
    var feed = getQBBankingFeed(entity);
    var rule = {
      'Match Plaid Account': matcherConfig.matchPlaidAccount || '',
      'Match Name Contains': matcherConfig.matchName || '',
      'Match Amount':        matcherConfig.matchAmount || '',
      'Direction':           dir,
      'DR Account':          drAcct,
      'CR Account':          crAcct
    };
    (feed.txns || []).forEach(function(t) {
      if (!_qbRuleMatches_(t, rule)) return;
      try {
        if (!t.posted) {
          postQBPlaidTxn(entity, t.plaidId, cat, '');
          applied++;
        } else if (t.needsReview) {
          reclassifyQBPlaidTxn(entity, t.plaidId, cat);
          applied++;
        }
        // Case 3: posted + already-categorized → leave it alone.
      } catch (e) {
        Logger.log('Rule apply skipped ' + t.plaidId + ': ' + e.message);
      }
    });
  }

  return { success: true, ruleId: ruleId, applied: applied };
}

// Public: list all bank rules for an entity, with the Plaid-side matcher
// summarized for the UI.
function getQBRules(entity) {
  entity = String(entity || 'NF CA');
  var rules = _getQBRows_('QB_RULES', QB_RULES_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity; })
    .sort(function(a, b) { return (Number(a.Priority) || 999) - (Number(b.Priority) || 999); });
  return JSON.parse(JSON.stringify(rules));
}

// Public: delete a rule. Doesn't un-post anything already classified via it.
function deleteQBRule(ruleId) {
  _requireEditor_();
  ensureQBSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QB_RULES');
  if (!sheet || sheet.getLastRow() < 2) return { success: false };
  var lastCol = Math.max(sheet.getLastColumn(), QB_RULES_HEADERS.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var iId = hdr.indexOf('ID');
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  for (var r = data.length - 1; r >= 0; r--) {
    if (String(data[r][iId]) === String(ruleId)) {
      sheet.deleteRow(r + 2);
      return { success: true };
    }
  }
  return { success: false };
}


// Return a flat list of COA account options for the inline category picker,
// grouped and sorted sensibly. Excludes bank/CC accounts by default (those
// are handled via Transfer, not Category — can enable later).
function getQBCategoryOptions(entity) {
  entity = String(entity || 'NF CA');
  var coa = _getQBRows_('QB_COA', QB_COA_HEADERS)
    .filter(function(r) { return String(r.Entity) === entity && String(r.Active || 'Yes').toLowerCase() === 'yes'; });
  var groups = { Income: [], Expense: [], Equity: [], Asset: [], Liability: [], Other: [] };
  coa.forEach(function(a) {
    var cls = _qbClassifyType_(a['Type']);
    (groups[cls] || groups.Other).push({ name: a['Account Name'], type: a['Type'], detailType: a['Detail Type'] });
  });
  Object.keys(groups).forEach(function(k) {
    groups[k].sort(function(a, b) { return a.name.localeCompare(b.name); });
  });
  return JSON.parse(JSON.stringify({ entity: entity, groups: groups }));
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
