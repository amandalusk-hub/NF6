/**
 * Code.gs — Family Office Wealth Tracker
 * Google Apps Script — single-file backend + web app
 *
 * SCRIPT PROPERTIES TO SET (Extensions > Apps Script > Project Settings > Script Properties):
 *   SPREADSHEET_ID      — ID of the backing Google Sheet (optional if running bound to sheet)
 *   RENTCAST_API_KEY    — from rentcast.io (free tier: 50 req/month)
 *   EXCHANGERATE_API_KEY— from exchangerate-api.com (optional, free tier works without key)
 *   PLAID_CLIENT_ID     — your Plaid client ID
 *   PLAID_SECRET        — your Plaid secret
 *   PLAID_ENV           — sandbox | production
 */

// ── Constants ─────────────────────────────────────────────────────────────────

var CATEGORIES = [
  'Real Estate - United States',
  'Real Estate - Colombia',
  'Real Estate - Puerto Rico',
  'Real Estate - Dominican Republic',
  'Real Estate - Europe',
  'Cash - Personal',
  'Cash - Business',
  'Private Equity',
  'Public Equity (Dividends)',
  'Public Equity (Growth)',
  'Loans Receivable',
  'Automobile',
  'Art/Jewelry/Other',
  'Crypto',
  'VIP Medical Group',
  'Insurance',
  'Other'
];

var CURRENCIES = ['USD','EUR','GBP','COP','BRL','MXN','CAD','JPY','CHF','AUD','DOP','PYG'];

var SUPPORTED_CURRENCIES = {
  'AED':'UAE Dirham','ARS':'Argentine Peso','AUD':'Australian Dollar',
  'BRL':'Brazilian Real','CAD':'Canadian Dollar','CHF':'Swiss Franc',
  'CLP':'Chilean Peso','CNY':'Chinese Yuan','COP':'Colombian Peso',
  'DOP':'Dominican Peso','EUR':'Euro','GBP':'British Pound',
  'HKD':'Hong Kong Dollar','INR':'Indian Rupee','JPY':'Japanese Yen',
  'KRW':'South Korean Won','MXN':'Mexican Peso','MYR':'Malaysian Ringgit',
  'NOK':'Norwegian Krone','NZD':'New Zealand Dollar','PEN':'Peruvian Sol',
  'PHP':'Philippine Peso','PLN':'Polish Zloty','PYG':'Paraguayan Guaraní',
  'SAR':'Saudi Riyal','SEK':'Swedish Krona','SGD':'Singapore Dollar',
  'THB':'Thai Baht','TRY':'Turkish Lira','TWD':'Taiwan Dollar',
  'USD':'US Dollar','VND':'Vietnamese Dong','ZAR':'South African Rand'
};

var COL = {
  ASSETS:      ['ID','Name','Category','Entity','Currency','Local Value','USD Rate','USD Value','My Share %','My Share USD','Date Added','Last Updated','Notes','Plaid Account ID','Address','Cost Basis','Details'],
  LIABILITIES: ['ID','Name','Type','Currency','Amount','USD Value','Date Added','Last Updated','Notes','Location','Details'],
  ENTITIES:    ['Name','Type','Jurisdiction','Ownership %','Notes','Tax ID','Date Created','Purpose','Trust Structure','Operating Agreement','EIN Document','Owners'],
  FX:          ['Currency','Rate to USD','Last Fetched'],
  NW_SNAPSHOTS: ['Date','Month Key','Type','Name','Category','USD Value'],
  HISTORY:     ['Date','Asset Name','Old Value USD','New Value USD','Delta USD','Currency','Notes'],
  SNAPSHOTS:   ['Date','Month Key','Asset Name','Category','Currency','My Share USD'],
  ASSET_DETAILS: [
    // ── Core ──────────────────────────────────────────────────────────────
    'Asset ID','Asset Name','Status','Updated By','Description','Drive Folder',
    'Project Leader',
    // ── Real Estate ───────────────────────────────────────────────────────
    'Occupancy','Location','Type','Sqft',
    'Purchase Price','Purchase Date','Closing Costs','Permits',
    'Revenue','OpEx','Property Tax','Insurance','HOA','Maintenance','Utilities',
    'Loan Info','Financial Notes',
    // ── Contacts ──────────────────────────────────────────────────────────
    'Contact 1 Type','Contact 1 Name','Contact 2 Type','Contact 2 Name',
    'Contact 3 Type','Contact 3 Name','Contact 4 Type','Contact 4 Name',
    // ── Loans Receivable ──────────────────────────────────────────────────
    'Borrower Name','Borrower Contact','Original Amount','Outstanding Balance',
    'Interest Rate','Loan Status','Loan Date','Due Date','Loan Terms',
    'Payment Schedule','Received To Date','Collateral','Drive Link','Attorney','Loan Notes',
    // ── Cash / Bank ───────────────────────────────────────────────────────
    'Cash Bank','Cash Account Type','Cash Account Number','Cash Interest Rate',
    // ── Public Equity ─────────────────────────────────────────────────────
    'Custodian / Manager','Equity Account Number','Shares / Units',
    'Avg Cost Per Share','Equity Notes',
    // ── Private Equity ────────────────────────────────────────────────────
    'PE Manager','PE Tax Treatment','PE Year Invested','PE Target Exit Year',
    'PE Year Sold','PE Year Written Off','PE Initial Investment','PE Ownership %',
    'PE Maturity Date','PE Return Rate','PE Capital Calls','PE Distributions','PE Notes'
  ],
  LIABILITY_DETAILS: ['Liability ID','Liability Name','Bank / Lender','Account Number','Interest Rate','Loan Type','Original Amount','Current Balance','Start Date','Maturity Date','Loan Term','Months Remaining','Monthly Payment','Principal','Interest Payment','Escrow','Property Tax','Insurance','HOA','Loan Officer','Attorney / Title','Insurance Agent','Other Contacts','Notes'],
  ORG_CHART: ['ID','Name','Parents','Node Type','Tax ID','Jurisdiction','Date Created','Ownership','Color','Text Color','Notes','Structure','X','Y']
};

// Maps JS field names ↔ Asset Details sheet column names
var ASSET_DET_MAP = [
  // Core
  ['status','Status'],['updatedBy','Updated By'],['description','Description'],
  ['folder','Drive Folder'],['projectLeader','Project Leader'],
  // Real Estate
  ['occupancy','Occupancy'],['location','Location'],['type','Type'],['sqft','Sqft'],
  ['purchasePrice','Purchase Price'],['purchaseDate','Purchase Date'],
  ['closingCosts','Closing Costs'],['permits','Permits'],['revenue','Revenue'],['opex','OpEx'],
  ['taxes','Property Tax'],['insurance','Insurance'],['hoa','HOA'],['maintenance','Maintenance'],
  ['utilities','Utilities'],['loan','Loan Info'],['finNotes','Financial Notes'],
  // Loans Receivable
  ['borrower','Borrower Name'],['borrowerContact','Borrower Contact'],
  ['loanOriginal','Original Amount'],['loanBalance','Outstanding Balance'],
  ['loanRate','Interest Rate'],['loanStatus','Loan Status'],['loanDate','Loan Date'],
  ['loanDue','Due Date'],['loanTerms','Loan Terms'],['loanPayment','Payment Schedule'],
  ['loanReceived','Received To Date'],['loanCollateral','Collateral'],
  ['loanDrive','Drive Link'],['loanAttorney','Attorney'],['loanNotes','Loan Notes'],
  // Cash / Bank
  ['cashBank','Cash Bank'],['cashAcctType','Cash Account Type'],
  ['cashAcctNum','Cash Account Number'],['cashRate','Cash Interest Rate'],
  // Public Equity
  ['eqBroker','Custodian / Manager'],['eqAcct','Equity Account Number'],
  ['eqShares','Shares / Units'],['eqCost','Avg Cost Per Share'],['eqNotes','Equity Notes'],
  // Private Equity
  ['peManager','PE Manager'],['peTax','PE Tax Treatment'],
  ['peYearIn','PE Year Invested'],['peYearExit','PE Target Exit Year'],
  ['peYearSold','PE Year Sold'],['peYearWo','PE Year Written Off'],
  ['peInitial','PE Initial Investment'],['peOwnership','PE Ownership %'],
  ['peMaturity','PE Maturity Date'],['peRate','PE Return Rate'],
  ['capitalCalls','PE Capital Calls'],['distributions','PE Distributions'],
  ['peNotes','PE Notes']
];

// Maps JS field names ↔ Liability Details sheet column names
var LIAB_DET_MAP = [
  ['bank','Bank / Lender'],['account','Account Number'],['rate','Interest Rate'],
  ['loanType','Loan Type'],['original','Original Amount'],['balance','Current Balance'],
  ['startDate','Start Date'],['maturity','Maturity Date'],['term','Loan Term'],
  ['remaining','Months Remaining'],['payment','Monthly Payment'],['principal','Principal'],
  ['interestPmt','Interest Payment'],['escrow','Escrow'],['tax','Property Tax'],
  ['insurance','Insurance'],['hoa','HOA'],['officer','Loan Officer'],
  ['attorney','Attorney / Title'],['insAgent','Insurance Agent'],
  ['contacts','Other Contacts'],['notes','Notes']
];

// ── Menu ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tracker')
    .addItem('Check Setup', 'deploymentReadinessCheck')
    .addSeparator()
    .addItem('Refresh FX Rates', 'fetchExchangeRates')
    .addItem('Refresh US Property Values', 'refreshPropertyValues')
    .addItem('Lookup Single Property', 'lookupSingleProperty')
    .addItem('Sync Plaid Accounts', 'syncPlaidAccounts')
    .addSeparator()
    .addItem('Connect Bank Account', 'openPlaidLink')
    .addItem('Configure Plaid Credentials', 'setPlaidCredentials')
    .addItem('Remove Plaid Connection', 'removePlaidConnection')
    .addSeparator()
    .addItem('Install Daily Trigger', 'installTriggers')
    .addSeparator()
    .addItem('Fix Sheet Headers (run once)', 'fixSheetHeaders')
    .addSeparator()
    .addItem('Seed Org Chart Structure (run once)', 'seedOrgChart')
    .addSeparator()
    .addItem('Refresh Balances Sheet', 'generateBalancesSheet')
    .addSeparator()
    .addItem('Take Net Worth Snapshot (1st of month)', 'takeNWSnapshot')
    .addItem('Refresh Net Worth History Sheet', 'generateNetWorthHistorySheet')
    .addSeparator()
    .addItem('Setup Database Structure', 'setupDatabase')
    .addItem('Reset Asset Details Schema', 'resetSchema')
    .addToUi();
}

// ── Deployment Readiness Check ────────────────────────────────────────────────

function deploymentReadinessCheck() {
  var props  = PropertiesService.getScriptProperties();
  var checks = [];
  var ok     = true;

  function check(label, passed, detail) {
    checks.push((passed ? '✅' : '❌') + ' ' + label + (detail ? ': ' + detail : ''));
    if (!passed) ok = false;
  }

  // 1. Spreadsheet access
  try {
    var ss = getSpreadsheet_();
    check('Spreadsheet', true, ss.getName());
  } catch(e) {
    check('Spreadsheet', false, 'Cannot open — set SPREADSHEET_ID or run from Sheet');
  }

  // 2. Required sheets
  try {
    ensureSheets_();
    check('Sheets (Assets / Liabilities / Entities / FX Rates / History)', true, 'all present');
  } catch(e) {
    check('Sheets', false, e.message);
  }

  // 3. FX API
  try {
    var fxResp = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/USD', { muteHttpExceptions: true });
    check('FX API (open.er-api.com)', fxResp.getResponseCode() === 200, 'HTTP ' + fxResp.getResponseCode());
  } catch(e) {
    check('FX API', false, e.message);
  }

  // 4. Rentcast API key
  var rentcastKey = props.getProperty('RENTCAST_API_KEY');
  check('RENTCAST_API_KEY', !!rentcastKey, rentcastKey ? 'set' : 'missing — US property valuations will not work');

  // 5. Plaid credentials
  var plaidClientId = props.getProperty('PLAID_CLIENT_ID');
  var plaidSecret   = props.getProperty('PLAID_SECRET');
  var plaidEnv      = props.getProperty('PLAID_ENV');
  check('PLAID_CLIENT_ID', !!plaidClientId, plaidClientId ? 'set' : 'missing — bank sync will not work');
  check('PLAID_SECRET',    !!plaidSecret,   plaidSecret   ? 'set' : 'missing — bank sync will not work');
  check('PLAID_ENV',       !!plaidEnv,      plaidEnv      ? plaidEnv : 'missing (sandbox / production)');

  // 6. Plaid connectivity
  if (plaidClientId && plaidSecret && plaidEnv) {
    try {
      var plaidResult = getPlaidLinkToken();
      check('Plaid API connectivity', plaidResult.success, plaidResult.success ? 'OK' : plaidResult.error);
    } catch(e) {
      check('Plaid API connectivity', false, e.message);
    }
  } else {
    checks.push('⚠️  Plaid API connectivity: skipped (credentials not set)');
  }

  // 7. Daily trigger
  var triggers   = ScriptApp.getProjectTriggers();
  var hasTrigger = triggers.some(function(t) { return t.getHandlerFunction() === 'dailySync_'; });
  check('Daily sync trigger', hasTrigger, hasTrigger ? 'installed' : 'not installed — run "Install Daily Trigger"');

  // 8. Web app (informational)
  checks.push('ℹ️  Web app: deploy via Deploy > New deployment to get a shareable URL');

  var summary = (ok ? '✅ All required checks passed.' : '⚠️  Some checks failed — see details below.') +
    '\n\n' + checks.join('\n');

  SpreadsheetApp.getUi().alert('Deployment Readiness Check', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  return { ok: ok, checks: checks };
}

// ── Web App Entry Point ───────────────────────────────────────────────────────

function doGet() {
  var email = Session.getActiveUser().getEmail();
  if (!email || !email.toLowerCase().endsWith('@nf6capital.com')) {
    var display = email ? email : 'not signed in';
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
      '<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#f0f2f5;display:flex;align-items:center;justify-content:center;min-height:100vh}' +
      '.box{background:#fff;border-radius:10px;box-shadow:0 4px 20px rgba(0,0,0,.12);padding:48px 40px;max-width:400px;width:100%;text-align:center}' +
      '.logo{font-size:28px;font-weight:700;color:#0d2137;margin-bottom:8px}' +
      '.sub{font-size:14px;color:#666;margin-bottom:32px}' +
      '.icon{font-size:48px;margin-bottom:20px}' +
      'h2{font-size:20px;color:#0d2137;margin-bottom:10px}' +
      'p{font-size:14px;color:#666;line-height:1.6;margin-bottom:8px}' +
      '.account{background:#fce8e6;border-radius:6px;padding:10px 14px;font-size:13px;color:#c5221f;margin-top:20px}' +
      '</style></head><body>' +
      '<div class="box"><div class="logo">NF6</div><div class="sub">Family Office Wealth Tracker</div>' +
      '<div class="icon">🔒</div>' +
      '<h2>Access Restricted</h2>' +
      '<p>This application is only available to NF6 Capital team members.</p>' +
      '<p>Please sign in with your <strong>@nf6capital.com</strong> account.</p>' +
      '<div class="account">Currently signed in as: <strong>' + display + '</strong></div>' +
      '</div></body></html>'
    ).setTitle('Access Restricted');
  }
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Family Office — Wealth Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Sheet Bootstrapping ───────────────────────────────────────────────────────

var _sheetsReady = false;
function ensureSheets_() {
  if (_sheetsReady) return;
  // Use script cache so the full sheet scan is skipped across executions
  // (GAS globals reset per execution, but CacheService persists for up to 6h)
  var cache = CacheService.getScriptCache();
  if (cache.get('sheets_ready') === '1') { _sheetsReady = true; return; }

  var ss = getSpreadsheet_();
  Object.keys(COL).forEach(function(key) {
    var name    = sheetName_(key);
    var headers = COL[key];
    if (Array.isArray(headers) && headers[0] && headers[0].constructor === Array) {
      // COL.ASSET_DETAILS is an array of arrays after the schema change — flatten
      headers = headers.reduce(function(a, b) { return a.concat(b); }, []);
    }
    var sheet   = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 220);
      cache.remove('sheets_ready'); // new sheet added — invalidate cache
    } else {
      var lastCol  = sheet.getLastColumn();
      var existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
      var missing  = headers.filter(function(h) { return existing.indexOf(h) === -1; });
      if (missing.length) {
        var startCol = lastCol + 1;
        sheet.getRange(1, startCol, 1, missing.length).setValues([missing])
          .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
        cache.remove('sheets_ready'); // schema changed — invalidate
      }
    }
  });
  cache.put('sheets_ready', '1', 21600); // valid for 6 hours
  _sheetsReady = true;
}

function sheetName_(key) {
  return {
    ASSETS: 'Assets', LIABILITIES: 'Liabilities', ENTITIES: 'Entities',
    FX: 'FX Rates', HISTORY: 'History', SNAPSHOTS: 'Snapshots',
    ASSET_DETAILS: 'Asset Details', LIABILITY_DETAILS: 'Liability Details',
    ORG_CHART: 'Org Chart', NW_SNAPSHOTS: 'NW Snapshots'
  }[key];
}

function getSpreadsheet_() {
  if (_ss) return _ss;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  _ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  return _ss;
}

var _ss = null;

function getSheet_(key) {
  ensureSheets_();
  return getSpreadsheet_().getSheetByName(sheetName_(key));
}

function sheetToObjects_(key) {
  var sheet = getSheet_(key);
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row, i) {
    var obj = { _row: i + 2 };
    headers.forEach(function(h, j) { obj[h] = row[j]; });
    return obj;
  });
}

// ── Main Data Fetch ───────────────────────────────────────────────────────────

function getFullData() {
  ensureSheets_();

  // Read ASSETS once and reuse for both ID assignment and data return
  var assetSheet  = getSheet_('ASSETS');
  var assetData   = assetSheet.getDataRange().getValues();
  var assetHeader = assetData[0] || [];

  // Auto-assign IDs to any rows missing one — batch all writes
  var missingIdCells = [];
  var uuidMap = {};
  for (var i = 1; i < assetData.length; i++) {
    if (!assetData[i][0]) {
      var newId = Utilities.getUuid();
      assetData[i][0] = newId;
      uuidMap[i] = newId;
    }
  }
  Object.keys(uuidMap).forEach(function(rowIdx) {
    assetSheet.getRange(Number(rowIdx) + 1, 1).setValue(uuidMap[rowIdx]);
  });

  function rowsToObjects(headers, rows) {
    return rows.slice(1).map(function(row) {
      var obj = {};
      headers.forEach(function(h, j) { obj[h] = row[j]; });
      return obj;
    });
  }

  function clean(arr) {
    return arr.map(function(obj) {
      var out = {};
      Object.keys(obj).forEach(function(k) {
        out[k] = obj[k] instanceof Date ? obj[k].toISOString() : obj[k];
      });
      return out;
    });
  }

  // Build asset objects from the already-read data (no second sheet read)
  var assets = clean(rowsToObjects(assetHeader, assetData));

  return {
    assets:      assets,
    liabilities: clean(sheetToObjects_('LIABILITIES')),
    entities:    clean(sheetToObjects_('ENTITIES')),
    fxRates:     clean(sheetToObjects_('FX')),
    // history omitted — fetched on demand via getAssetHistory() when detail panel opens
    snapshots:   getSnapshotTrend(),
    categories:  CATEGORIES,
    currencies:  CURRENCIES,
    _debug: {
      assetHeaders:  assetHeader,
      assetRowCount: Math.max(assetData.length - 1, 0)
    }
  };
}

// Fetch history for a single asset — called lazily when detail panel opens
function getAssetHistory(assetName) {
  return sheetToObjects_('HISTORY')
    .filter(function(h) { return h['Asset Name'] === assetName; })
    .map(function(h) {
      var out = {};
      Object.keys(h).forEach(function(k) { out[k] = h[k] instanceof Date ? h[k].toISOString() : h[k]; });
      return out;
    });
}

// ── Bulk Details Save (used by recovery to persist patched data back to sheet) ─
function bulkSaveDetails(idToDetailsMap) {
  var sheet    = getSheet_('ASSETS');
  var allRows  = sheet.getDataRange().getValues();
  var headers  = allRows[0];
  var detCol   = headers.indexOf('Details') + 1;
  if (detCol < 1) return { error: 'Details column not found' };
  var saved = 0;
  for (var i = 1; i < allRows.length; i++) {
    var id = String(allRows[i][0] || '');
    if (idToDetailsMap.hasOwnProperty(id)) {
      sheet.getRange(i + 1, detCol).setValue(idToDetailsMap[id]);
      saved++;
    }
  }
  return { saved: saved };
}

// ── Details Recovery ──────────────────────────────────────────────────────────
function getAllDetailsForRecovery() {
  var sheet = getSheet_('ASSETS');
  var data  = sheet.getDataRange().getValues();
  var h     = data[0];
  var detIdx  = h.indexOf('Details');
  var idIdx   = h.indexOf('ID');
  var nameIdx = h.indexOf('Name');
  var catIdx  = h.indexOf('Category');
  if (detIdx < 0) return { error: 'Details column not found' };
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][nameIdx]) continue;
    var raw = String(data[i][detIdx] || '');
    var det = {};
    try { det = JSON.parse(raw); } catch(e) {}
    // Determine how much real data is actually in the Details blob
    var meaningfulFields = Object.keys(det).filter(function(k) {
      var v = det[k];
      if (!v) return false;
      if (Array.isArray(v)) return v.length > 0 && v.some(function(item) {
        return item && typeof item === 'object' && Object.values(item).some(function(x){ return x && x !== 0; });
      });
      return String(v).trim().length > 0;
    });
    result.push({
      id:              String(data[i][idIdx] || ''),
      name:            String(data[i][nameIdx] || ''),
      category:        String(data[i][catIdx] || ''),
      detailsLength:   raw.length,
      meaningfulCount: meaningfulFields.length,
      meaningfulFields: meaningfulFields.slice(0, 10), // first 10 for diagnosis
      details:         det
    });
  }
  return result;
}

// Fetch all history — called lazily when History tab is opened
function getHistoryData() {
  return sheetToObjects_('HISTORY').map(function(h) {
    var out = {};
    Object.keys(h).forEach(function(k) { out[k] = h[k] instanceof Date ? h[k].toISOString() : h[k]; });
    return out;
  });
}

// ── FX Rates ──────────────────────────────────────────────────────────────────

function fetchExchangeRates() {
  try {
    var resp = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/USD', { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return { success: false, error: 'HTTP ' + resp.getResponseCode() };

    var data  = JSON.parse(resp.getContentText());
    var sheet = getSheet_('FX');
    var now   = new Date();
    var rows  = [];

    CURRENCIES.forEach(function(code) {
      var rate = code === 'USD' ? 1 : (data.rates[code] ? 1 / data.rates[code] : null);
      if (rate !== null) rows.push([code, rate, now]);
    });

    var last = sheet.getLastRow();
    if (last > 1) sheet.getRange(2, 1, last - 1, 3).clearContent();
    if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);

    var map = {};
    rows.forEach(function(r) { map[r[0]] = r[1]; });
    PropertiesService.getScriptProperties().setProperty('FX_CACHE', JSON.stringify({ rates: map, fetched: now.toISOString() }));

    return { success: true, rates: map };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getFxRate_(currency) {
  if (!currency || currency === 'USD') return 1;
  try {
    var cached = PropertiesService.getScriptProperties().getProperty('FX_CACHE');
    if (cached) {
      var obj    = JSON.parse(cached);
      var ageHrs = (Date.now() - new Date(obj.fetched).getTime()) / 3600000;
      if (ageHrs < 4 && obj.rates[currency]) return obj.rates[currency];
    }
  } catch(e) {}
  var result = fetchExchangeRates();
  return (result.success && result.rates[currency]) ? result.rates[currency] : 1;
}

/**
 * Get today's exchange rate for a currency to USD.
 * Use directly in sheet cells: =FX_RATE("COP")
 * @param {string} currencyCode - ISO 4217 code (e.g. "COP", "EUR")
 * @return {number} Rate: 1 unit of currencyCode = X USD
 * @customfunction
 */
function FX_RATE(currencyCode) {
  if (!currencyCode) return '';
  currencyCode = currencyCode.toString().trim().toUpperCase();
  if (currencyCode === 'USD') return 1;

  var cache    = CacheService.getScriptCache();
  var cacheKey = 'fx_' + currencyCode + '_USD';
  var cached   = cache.get(cacheKey);
  if (cached) return Number(cached);

  try {
    var apiKey  = PropertiesService.getScriptProperties().getProperty('EXCHANGERATE_API_KEY');
    var baseUrl = apiKey
      ? 'https://v6.exchangerate-api.com/v6/' + apiKey + '/latest/' + currencyCode
      : 'https://open.er-api.com/v6/latest/' + currencyCode;

    var resp = UrlFetchApp.fetch(baseUrl, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return 'Error: HTTP ' + resp.getResponseCode();

    var data = JSON.parse(resp.getContentText());
    if (!data.rates || !data.rates['USD']) return 'Error: no rate';

    var rate = data.rates['USD'];
    cache.put(cacheKey, rate.toString(), 14400); // cache 4 hours
    return rate;
  } catch(e) {
    return 'Error: ' + e.message;
  }
}

/**
 * Convert a foreign currency amount to USD.
 * Use directly in sheet cells: =TO_USD(1000000, "COP")
 * @param {number} amount
 * @param {string} currencyCode - ISO 4217 code
 * @return {number} USD equivalent
 * @customfunction
 */
function TO_USD(amount, currencyCode) {
  if (!amount || !currencyCode) return '';
  currencyCode = currencyCode.toString().trim().toUpperCase();
  if (currencyCode === 'USD') return Number(amount);
  var rate = FX_RATE(currencyCode);
  if (typeof rate !== 'number') return rate;
  return Number(amount) * rate;
}

// ── Assets CRUD ───────────────────────────────────────────────────────────────

function addAsset(data) {
  var sheet    = getSheet_('ASSETS');
  var id       = Utilities.getUuid();
  var now      = new Date();
  var fxRate   = getFxRate_(data.currency || 'USD');
  var localVal = Number(data.localValue) || 0;
  var usdVal   = localVal * fxRate;
  var sharePct = data.mySharePct !== undefined ? Number(data.mySharePct) : 100;
  var shareUsd = usdVal * sharePct / 100;

  var costBasis  = Number(data.costBasis) || 0;
  var nameToSave = data.name || '';
  sheet.appendRow([
    id, nameToSave, data.category || '', data.entity || '',
    data.currency || 'USD', localVal, fxRate, usdVal,
    sharePct, shareUsd, now, now, data.notes || '', '', data.address || '', costBasis,
    data.details || ''
  ]);
  return { success: true, id: id, savedName: nameToSave };
}

function updateAsset(data) {
  var sheet = getSheet_('ASSETS');
  var rows  = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] !== data.id) continue;

    var oldUsd   = Number(rows[i][7]) || 0;
    var currency = data.currency || rows[i][4];
    var fxRate   = getFxRate_(currency);
    var localVal = data.localValue !== undefined ? Number(data.localValue) : Number(rows[i][5]);
    var usdVal   = localVal * fxRate;
    var sharePct = data.mySharePct !== undefined ? Number(data.mySharePct) : Number(rows[i][8]);
    var shareUsd = usdVal * sharePct / 100;
    var now      = new Date();

    var updates = [
      [2,  data.name     !== undefined ? data.name     : rows[i][1]],
      [3,  data.category !== undefined ? data.category : rows[i][2]],
      [4,  data.entity   !== undefined ? data.entity   : rows[i][3]],
      [5,  currency],
      [6,  localVal],
      [7,  fxRate],
      [8,  usdVal],
      [9,  sharePct],
      [10, shareUsd],
      [12, now],
      [13, data.notes   !== undefined ? data.notes   : rows[i][12]],
      [15, data.address  !== undefined ? data.address  : rows[i][14]],
      [16, data.costBasis !== undefined ? Number(data.costBasis) : (Number(rows[i][15]) || 0)],
      [17, data.details   !== undefined ? data.details           : (rows[i][16] || '')]
    ];
    var newRow = rows[i].slice();
    updates.forEach(function(u) { newRow[u[0] - 1] = u[1]; });
    sheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);

    if (Math.abs(usdVal - oldUsd) > 0.01) {
      logHistory_(data.name || rows[i][1], oldUsd, usdVal, currency, data.notes || 'Manual update');
    }
    return { success: true };
  }
  return { success: false, error: 'Asset not found' };
}

function saveAssetDetails(id, detailsJson) {
  var det = {};
  try { det = JSON.parse(detailsJson || '{}'); } catch(e) {}

  // 1. Write JSON blob to ASSETS.Details column (fast read path)
  var assetsSheet  = getSheet_('ASSETS');
  var assetsData   = assetsSheet.getDataRange().getValues();
  var headers      = assetsData[0];
  var detailsCol   = headers.indexOf('Details') + 1;
  var lastUpdCol   = headers.indexOf('Last Updated') + 1;
  var categoryCol  = headers.indexOf('Category') + 1;
  var shareCol     = headers.indexOf('My Share %') + 1;
  var localValCol  = headers.indexOf('Local Value') + 1;
  var usdValCol    = headers.indexOf('USD Value') + 1;
  var shareUsdCol  = headers.indexOf('My Share USD') + 1;
  var assetName  = '';
  var savedAsset = false;
  for (var i = 1; i < assetsData.length; i++) {
    if (String(assetsData[i][0]) !== String(id)) continue;
    assetName = String(assetsData[i][1] || '');
    var newRow = assetsData[i].slice();
    if (detailsCol > 0) newRow[detailsCol - 1] = detailsJson || '';
    if (lastUpdCol > 0) newRow[lastUpdCol  - 1] = new Date();
    assetsSheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
    savedAsset = true;
    break;
  }

  if (!savedAsset) {
    return { success: false, error: 'Asset ID not found in sheet (id=' + id + '). Data was NOT saved.' };
  }

  // 2. Upsert row in Asset Details sheet (flat columns, spreadsheet-editable)
  var detSheet   = getSheet_('ASSET_DETAILS');
  var detData    = detSheet.getDataRange().getValues();
  var detHeaders = detData[0];
  var contacts   = det.contacts || [];

  // Build a column-name → value lookup
  // Arrays (capitalCalls, distributions) are stored as JSON strings in flat columns
  var colLookup = { 'Asset ID': id, 'Asset Name': assetName };
  ASSET_DET_MAP.forEach(function(m) {
    var val = det[m[0]];
    colLookup[m[1]] = Array.isArray(val) ? JSON.stringify(val) : (val || '');
  });
  for (var ci = 1; ci <= 4; ci++) {
    var c = contacts[ci - 1] || {};
    colLookup['Contact ' + ci + ' Type'] = c.type || '';
    colLookup['Contact ' + ci + ' Name'] = c.name || '';
  }
  var rowData = detHeaders.map(function(h) { return colLookup[h] !== undefined ? colLookup[h] : ''; });

  try {
    var existingRow = -1;
    for (var j = 1; j < detData.length; j++) {
      if (String(detData[j][0]) === String(id)) { existingRow = j + 1; break; }
    }
    if (existingRow > 0) {
      detSheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      detSheet.appendRow(rowData);
    }
  } catch(e) {
    Logger.log('ASSET_DETAILS flat sheet update failed (non-fatal): ' + e.message);
  }

  return { success: true };
}

// Combined save: updates core fields + details in one GAS call (half the round-trips).
// coreData mirrors the updateAsset() payload; detailsJson is the JSON string for Details.
function saveFullAsset(coreData, id, detailsJson) {
  var det = {};
  try { det = JSON.parse(detailsJson || '{}'); } catch(e) {}

  var sheet      = getSheet_('ASSETS');
  var allRows    = sheet.getDataRange().getValues();
  var headers    = allRows[0];
  var detailsCol = headers.indexOf('Details') + 1;
  var assetName  = '';
  var savedRow   = false;

  // Declare outside loop so they're accessible for the return statement
  var shareUsd = 0, localVal = 0, ownershipPct = 0;

  for (var i = 1; i < allRows.length; i++) {
    if (String(allRows[i][0]) !== String(id)) continue;

    assetName    = String(allRows[i][1] || '');
    var oldUsd   = Number(allRows[i][7]) || 0;
    var currency = (coreData && coreData.currency) || allRows[i][4];
    var fxRate   = getFxRate_(currency);
    localVal     = (coreData && coreData.localValue !== undefined) ? Number(coreData.localValue) : Number(allRows[i][5]);
    var usdVal   = localVal * fxRate;
    // Ownership % is informational only — My Share USD = local value directly (no multiplication)
    ownershipPct = (coreData && coreData.ownershipPct !== undefined) ? Number(coreData.ownershipPct) : Number(allRows[i][8]);
    shareUsd     = usdVal; // always equals Mike's entered value × FX, no % applied
    var newRow   = allRows[i].slice();

    // Core field updates
    if (coreData) {
      if (coreData.name      !== undefined) newRow[1]  = coreData.name;
      if (coreData.category  !== undefined) newRow[2]  = coreData.category;
      if (coreData.entity    !== undefined) newRow[3]  = coreData.entity;
      newRow[4]  = currency;
      newRow[5]  = localVal;
      newRow[6]  = fxRate;
      newRow[7]  = usdVal;
      newRow[8]  = ownershipPct; // stored for reference only
      newRow[9]  = shareUsd;
      if (coreData.notes     !== undefined) newRow[12] = coreData.notes;
      if (coreData.address   !== undefined) newRow[14] = coreData.address;
      if (coreData.costBasis !== undefined) newRow[15] = Number(coreData.costBasis);
    }
    // Details JSON + last updated
    if (detailsCol > 0) newRow[detailsCol - 1] = detailsJson || '';
    newRow[11] = new Date(); // Last Updated

    sheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
    savedRow = true;

    if (coreData && Math.abs(usdVal - oldUsd) > 0.01) {
      logHistory_(newRow[1], oldUsd, usdVal, currency, (coreData && coreData.notes) || 'Updated');
    }
    break;
  }

  if (!savedRow) {
    return { success: false, error: 'Asset ID not found in sheet (id=' + id + '). Data was NOT saved.' };
  }

  // Update flat ASSET_DETAILS sheet (for spreadsheet viewing)
  try {
    var detSheet   = getSheet_('ASSET_DETAILS');
    var detData    = detSheet.getDataRange().getValues();
    var detHeaders = detData[0];
    var contacts   = det.contacts || [];
    var colLookup  = { 'Asset ID': id, 'Asset Name': assetName };
    ASSET_DET_MAP.forEach(function(m) {
      var val = det[m[0]];
      colLookup[m[1]] = Array.isArray(val) ? JSON.stringify(val) : (val || '');
    });
    for (var ci = 1; ci <= 4; ci++) {
      var c = contacts[ci - 1] || {};
      colLookup['Contact ' + ci + ' Type'] = c.type || '';
      colLookup['Contact ' + ci + ' Name'] = c.name || '';
    }
    var rowData    = detHeaders.map(function(h) { return colLookup[h] !== undefined ? colLookup[h] : ''; });
    var existingRow = -1;
    for (var j = 1; j < detData.length; j++) {
      if (String(detData[j][0]) === String(id)) { existingRow = j + 1; break; }
    }
    if (existingRow > 0) {
      detSheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      detSheet.appendRow(rowData);
    }
  } catch(e) {
    Logger.log('ASSET_DETAILS flat sheet update failed (non-fatal): ' + e.message);
  }

  // Return the updated values so frontend can sync
  return {
    success:      true,
    myShareUsd:   shareUsd,
    localValue:   localVal,
    ownershipPct: ownershipPct
  };
}

// Fast auto-save: reads only the header row + ID column, writes just the Details cell.
// Called by the 2-second auto-save timer — much faster than saveFullAsset (~300ms vs 1-3s).
function saveAssetDetailsOnly(id, detailsJson) {
  var sheet   = getSheet_('ASSETS');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: 'No asset rows' };
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var detCol  = headers.indexOf('Details') + 1;
  var updCol  = headers.indexOf('Last Updated') + 1;
  if (detCol < 1) return { success: false, error: 'No Details column in ASSETS sheet' };
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sheet.getRange(i + 2, detCol).setValue(detailsJson || '');
      if (updCol > 0) sheet.getRange(i + 2, updCol).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false, error: 'Asset not found: ' + id };
}

function deleteAsset(id) {
  var sheet = getSheet_('ASSETS');
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, error: 'Not found' };
}

// ── Entities CRUD ─────────────────────────────────────────────────────────────

function addEntity(data) {
  var sheet   = getSheet_('ENTITIES');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = headers.map(function(h) {
    switch (h) {
      case 'Name':                return data.name                 || '';
      case 'Type':                return data.type                 || '';
      case 'Jurisdiction':        return data.jurisdiction         || '';
      case 'Ownership %':         return data.ownershipPct !== undefined ? Number(data.ownershipPct) : 100;
      case 'Notes':               return data.notes                || '';
      case 'Tax ID':              return data.taxId                || '';
      case 'Date Created':        return data.dateCreated          || '';
      case 'Purpose':             return data.purpose              || '';
      case 'Trust Structure':     return data.trustStructure       || '';
      case 'Operating Agreement': return data.operatingAgreement   || '';
      case 'EIN Document':        return data.einDocument          || '';
      case 'Owners':              return data.ownersJson           || '';
      default:                    return '';
    }
  });
  sheet.appendRow(row);
  return { success: true };
}

function updateEntity(data) {
  var sheet   = getSheet_('ENTITIES');
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0];
  function ci(name) { return headers.indexOf(name); }
  function cur(name) { var j = ci(name); return j >= 0 ? rows[i][j] : ''; }
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] !== data.originalName) continue;
    var newRow = rows[i].slice();
    // Always-present legacy columns
    newRow[0] = data.name             !== undefined ? data.name                          : newRow[0];
    newRow[1] = data.type             !== undefined ? data.type                          : newRow[1];
    newRow[2] = data.jurisdiction     !== undefined ? data.jurisdiction                  : newRow[2];
    newRow[3] = data.ownershipPct     !== undefined ? Number(data.ownershipPct)          : newRow[3];
    newRow[4] = data.notes            !== undefined ? data.notes                         : newRow[4];
    // New extended columns (only written if column exists)
    if (ci('Tax ID')              >= 0) newRow[ci('Tax ID')]              = data.taxId              !== undefined ? data.taxId              : cur('Tax ID');
    if (ci('Date Created')        >= 0) newRow[ci('Date Created')]        = data.dateCreated        !== undefined ? data.dateCreated        : cur('Date Created');
    if (ci('Purpose')             >= 0) newRow[ci('Purpose')]             = data.purpose            !== undefined ? data.purpose            : cur('Purpose');
    if (ci('Trust Structure')     >= 0) newRow[ci('Trust Structure')]     = data.trustStructure     !== undefined ? data.trustStructure     : cur('Trust Structure');
    if (ci('Operating Agreement') >= 0) newRow[ci('Operating Agreement')] = data.operatingAgreement !== undefined ? data.operatingAgreement : cur('Operating Agreement');
    if (ci('EIN Document')        >= 0) newRow[ci('EIN Document')]        = data.einDocument        !== undefined ? data.einDocument        : cur('EIN Document');
    if (ci('Owners')              >= 0) newRow[ci('Owners')]              = data.ownersJson         !== undefined ? data.ownersJson         : cur('Owners');
    sheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
    return { success: true };
  }
  return { success: false, error: 'Not found' };
}

function deleteEntity(name) {
  var sheet = getSheet_('ENTITIES');
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === name) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, error: 'Not found' };
}

// ── Liabilities CRUD ──────────────────────────────────────────────────────────

function addLiability(data) {
  var sheet  = getSheet_('LIABILITIES');
  var id     = Utilities.getUuid();
  var now    = new Date();
  var fxRate = getFxRate_(data.currency || 'USD');
  var amount = Number(data.amount) || 0;
  sheet.appendRow([id, data.name || '', data.type || '', data.currency || 'USD', amount, amount * fxRate, now, now, data.notes || '', data.location || '', data.details || '']);
  return { success: true, id: id };
}

function updateLiability(data) {
  var sheet   = getSheet_('LIABILITIES');
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0] || [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] !== data.id) continue;
    var fxRate = getFxRate_(data.currency || rows[i][3]);
    var amount = data.amount !== undefined ? Number(data.amount) : Number(rows[i][4]);
    sheet.getRange(i + 1, 2).setValue(data.name     !== undefined ? data.name     : rows[i][1]);
    sheet.getRange(i + 1, 3).setValue(data.type     !== undefined ? data.type     : rows[i][2]);
    sheet.getRange(i + 1, 4).setValue(data.currency !== undefined ? data.currency : rows[i][3]);
    sheet.getRange(i + 1, 5).setValue(amount);
    sheet.getRange(i + 1, 6).setValue(amount * fxRate);
    sheet.getRange(i + 1, 8).setValue(new Date());
    sheet.getRange(i + 1, 9).setValue(data.notes    !== undefined ? data.notes    : rows[i][8]);
    var locCol = headers.indexOf('Location') + 1;
    if (locCol > 0) sheet.getRange(i + 1, locCol).setValue(data.location !== undefined ? data.location : (rows[i][locCol - 1] || ''));
    // NOTE: Details column is intentionally NOT written here — it is managed
    // exclusively by saveLiabilityDetails(). Writing it here would cause a
    // race condition that overwrites detail data with a stale value.
    return { success: true, usdValue: amount * fxRate, amount: amount };
  }
  return { success: false, error: 'Liability not found' };
}

function saveLiabilityDetails(id, detailsJson) {
  var det = {};
  try { det = JSON.parse(detailsJson || '{}'); } catch(e) {}

  // 1. Write JSON blob to LIABILITIES.Details column
  var liabSheet  = getSheet_('LIABILITIES');
  var liabData   = liabSheet.getDataRange().getValues();
  var headers    = liabData[0];
  var detailsCol = headers.indexOf('Details') + 1;
  var lastUpdCol = headers.indexOf('Last Updated') + 1;
  var amtCol     = headers.indexOf('Amount') + 1;
  var usdCol     = headers.indexOf('USD Value') + 1;
  var currCol    = headers.indexOf('Currency') + 1;

  // Add Details column if it's missing (schema may have been updated after sheet was created)
  if (detailsCol === 0) {
    detailsCol = liabSheet.getLastColumn() + 1;
    liabSheet.getRange(1, detailsCol).setValue('Details')
      .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
    CacheService.getScriptCache().remove('sheets_ready');
  }

  // Parse balance from details (stored as "$1,234" string from accounting format)
  var balanceNum = 0;
  if (det.balance) {
    balanceNum = parseFloat(String(det.balance).replace(/[$,\s]/g, '')) || 0;
  }

  var liabName  = '';
  var currency  = 'USD';
  var savedLiab = false;
  for (var i = 1; i < liabData.length; i++) {
    if (String(liabData[i][0]) !== String(id)) continue;
    liabName = String(liabData[i][1] || '');
    currency = (currCol > 0 ? liabData[i][currCol - 1] : '') || 'USD';
    liabSheet.getRange(i + 1, detailsCol).setValue(detailsJson || '');
    if (lastUpdCol > 0) liabSheet.getRange(i + 1, lastUpdCol).setValue(new Date());
    // Sync current balance → Amount + USD Value so the dashboard reflects it
    if (balanceNum > 0 && amtCol > 0 && usdCol > 0) {
      var fxRate = getFxRate_(currency);
      liabSheet.getRange(i + 1, amtCol).setValue(balanceNum);
      liabSheet.getRange(i + 1, usdCol).setValue(balanceNum * fxRate);
    }
    savedLiab = true;
    break;
  }

  if (!savedLiab) {
    return { success: false, error: 'Liability ID not found in sheet (id=' + id + '). Data was NOT saved.' };
  }

  // 2. Upsert row in Liability Details sheet (flat columns)
  try {
    var detSheet   = getSheet_('LIABILITY_DETAILS');
    var detData    = detSheet.getDataRange().getValues();
    var detHeaders = detData[0] || [];
    if (detHeaders.length > 0 && detHeaders[0] !== '') {
      var colLookup  = { 'Liability ID': id, 'Liability Name': liabName };
      LIAB_DET_MAP.forEach(function(m) { colLookup[m[1]] = det[m[0]] || ''; });
      var rowData = detHeaders.map(function(h) { return colLookup[h] !== undefined ? colLookup[h] : ''; });
      var existingRow = -1;
      for (var j = 1; j < detData.length; j++) {
        if (String(detData[j][0]) === String(id)) { existingRow = j + 1; break; }
      }
      if (existingRow > 0) {
        detSheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
      } else {
        detSheet.appendRow(rowData);
      }
    }
  } catch(e) {
    Logger.log('LIABILITY_DETAILS write failed: ' + e.message);
  }

  var fxRateForReturn = getFxRate_(currency);
  return { success: true, balance: balanceNum, usdValue: balanceNum * fxRateForReturn, currency: currency };
}

function deleteLiability(id) {
  var sheet = getSheet_('LIABILITIES');
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, error: 'Not found' };
}

// ── Org Chart Seed ────────────────────────────────────────────────────────────

function seedOrgChart() {
  ensureSheets_();
  var sheet = getSheet_('ORG_CHART');
  var existing = sheet.getLastRow();
  if (existing > 1) {
    var ui = SpreadsheetApp.getUi();
    var resp = ui.alert('Org Chart already has data (' + (existing - 1) + ' nodes). This will ADD new nodes without deleting existing ones. Continue?', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  // Helper: build row aligned to headers
  var headers = COL.ORG_CHART;
  function makeRow(obj) {
    return headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
  }

  // Pre-assign stable IDs so parent references work
  var ID = {
    MICHAEL:     Utilities.getUuid(),
    MICHELLE:    Utilities.getUuid(),
    DAVID:       Utilities.getUuid(),
    NANCY:       Utilities.getUuid(),
    MN_REV:      Utilities.getUuid(),  // 2013 MN Family Revocable Trust
    MN_IRREV:    Utilities.getUuid(),  // MN Family Trust Irrevocable
    NF6_MGMT:    Utilities.getUuid(),  // NF6 Joint MGMT LLC
    NF6_HOLD:    Utilities.getUuid(),  // NF6 Family Holding LP
    BPMGMT:      Utilities.getUuid(),  // BPMGMT LLC
    TIGER:       Utilities.getUuid(),  // NF5 Tiger Capital LLC
    BP_LP:       Utilities.getUuid(),  // Blue Panda Family LP
    YM_PR:       Utilities.getUuid(),  // YM PR Investment Group LLC
    NF6_V8:      Utilities.getUuid(),  // NF6 Venture 8 LLC
    NF5_CAP:     Utilities.getUuid(),  // NF5 Capital LLC
    TLMND:       Utilities.getUuid(),  // TLMND LLC
    NF_MOE:      Utilities.getUuid(),  // NF MOE CO SA S
    MEJ_DR:      Utilities.getUuid(),  // MEJ DR INC
    NF_EUR:      Utilities.getUuid(),  // NF Europe Holding S
    PARIS_SCI:   Utilities.getUuid(),  // Paris Thacko SCI
    NF5_SPAIN:   Utilities.getUuid(),  // NF5 Spain Holdings SL
    NF_US_CA:    Utilities.getUuid(),  // NF U S CA LLC
    CLIFTON:     Utilities.getUuid(),  // Clifton MCI LLC
    NF_US_TX:    Utilities.getUuid(),  // NF U S TX LLC
    // NF6 Venture 8 — US investments
    SOCIODOC:    Utilities.getUuid(),
    SOLARIS:     Utilities.getUuid(),
    NAT_NUTRA:   Utilities.getUuid(),
    LEXI_SKIN:   Utilities.getUuid(),
    MICIDA:      Utilities.getUuid(),
    TALKSELF:    Utilities.getUuid(),
    FS_NYC:      Utilities.getUuid(),
    SAN_ORL:     Utilities.getUuid(),
    SAN_RES:     Utilities.getUuid(),
    // NF6 Venture 8 — PR investments
    FIT_PR:      Utilities.getUuid(),
    SUN_HARBOR:  Utilities.getUuid(),
    SAN_D_EAST:  Utilities.getUuid(),
    SAN_D_WEST:  Utilities.getUuid(),
    // TLMND US investments
    NW_WOOD:     Utilities.getUuid(),
    SANU:        Utilities.getUuid(),
    KUHLMAN:     Utilities.getUuid(),
    // MEJ DR / Colombia
    FARANDA:     Utilities.getUuid(),
    LAND_AM:     Utilities.getUuid(),
    PARAISO:     Utilities.getUuid(),
    RF_HOLD:     Utilities.getUuid(),
    COLCIENO_A:  Utilities.getUuid(),
    COLCIENO_2:  Utilities.getUuid(),
    LAURELES:    Utilities.getUuid(),
    SEATLON:     Utilities.getUuid(),
    MANILA:      Utilities.getUuid(),
    LOAO:        Utilities.getUuid(),
    // Europe
    DGUN:        Utilities.getUuid(),
    SANGLIMA:    Utilities.getUuid(),
    PLAZA_COL:   Utilities.getUuid(),
    // 2026 Trust (post structure only)
    TRUST_2026:  Utilities.getUuid()
  };

  // Nodes: [ID-key, Name, Parents(array of ID-keys), NodeType, TaxID, Jurisdiction, DateCreated, Ownership, Color, Structure, X, Y]
  var NODES = [
    // ── Individuals ──────────────────────────────────────────────────────────
    ['MICHAEL',  'Michael Nguyen',   [],                       'Individual',         '',             'Puerto Rico', '',           'DOB: 7/19/1969\nType: 1099 Individual\nPhone: 216-313-4759', '#1a4f7a', 'both',    1150, 50],
    ['MICHELLE', 'Michelle Lam',     [],                       'Individual',         '',             'USA',         '',           '',                                                            '#1a4f7a', 'both',    100,  50],
    ['DAVID',    'David Nguyen',     [],                       'Individual',         '',             'USA',         '',           '',                                                            '#1a4f7a', 'both',    310,  50],
    ['NANCY',    'Nancy Nguyen',     [],                       'Individual',         '',             'USA',         '',           '',                                                            '#1a4f7a', 'both',    520,  50],

    // ── Level 1 ───────────────────────────────────────────────────────────────
    ['NF6_MGMT', 'NF6 Joint MGMT LLC',            ['MICHELLE','DAVID','NANCY'], 'LLC',              '83-4156504',  'USA',         '10/30/2022', '33.34% Michelle Lam\n33.33% Nancy Nguyen\n33.33% David Nguyen', '#0e4d5c', 'both', 300, 220],
    ['MN_REV',   '2013 MN Family Revocable Trust', ['MICHAEL'],                 'Trust (Revocable)', '',           'Puerto Rico', '',           'Grantor & Trustee: Michael Nguyen',                             '#2d6a4f', 'both', 730, 220],
    ['MN_IRREV', 'MN Family Trust - Irrevocable',  ['MICHAEL'],                 'Trust (Irrevocable)','',          'Puerto Rico', '',           'Grantor: Michael Nguyen\nTrustee: Nancy Nguyen',                '#2d6a4f', 'both', 1150, 220],

    // ── Level 2 ───────────────────────────────────────────────────────────────
    ['NF6_HOLD', 'NF6 Family Holding LP',  ['NF6_MGMT','MN_REV'], 'LP (Limited Partnership)', '', 'USA', '10/10/2022', 'General Partner: NF6 Joint MGMT LLC 1%\nLimited Partner: 2013 MN Family Revocable Trust 99%', '#0e4d5c', 'both', 300, 390],
    ['BPMGMT',   'BPMGMT LLC',             ['NANCY','MN_REV'],    'LLC',                      '82-3131738', 'USA', '10/01/2018', '49% Nancy Nguyen\n51% 2013 MN Family Revocable Trust', '#0e4d5c', 'both', 800, 390],
    ['YM_PR',    'YM PR Investment Group LLC', ['MN_IRREV'],       'LLC',                      '', 'Puerto Rico', '', 'Ownership: 50%', '#0e4d5c', 'both', 1150, 390],

    // ── Level 3 ───────────────────────────────────────────────────────────────
    ['TIGER',  'NF5 Tiger Capital LLC',  ['NF6_HOLD'], 'LLC',                    '', 'USA',          '10/15/2022', 'Ownership: 100% NF6 Family Holding LP\nType: Single Member Disregarded', '#0e4d5c', 'both', 100, 560],
    ['BP_LP',  'Blue Panda Family LP',  ['BPMGMT'],   'LP (Limited Partnership)','83-3832234', 'Puerto Rico', '09/15/2019', 'Owner/GP: BPMGMT LLC\nDate Est. 2019', '#0e4d5c', 'both', 800, 560],

    // ── Level 4 — Blue Panda children ────────────────────────────────────────
    ['NF6_V8',  'NF6 Venture 8 LLC',  ['BP_LP'], 'LLC',            '', 'USA', '06/09/2018', 'Type: Single Member Disregarded\nPurpose: Business and Real Estate Investment Holding', '#1e6e3b', 'both', 420, 730],
    ['NF5_CAP', 'NF5 Capital LLC',    ['BP_LP'], 'LLC',            '', 'USA', '',           'Type: Single Member Disregarded\nReturn Reported on Blue Panda LP\nPurpose: Cash Investment Holding\nOwnership: BP Family 100%', '#1e6e3b', 'both', 630, 730],
    ['TLMND',   'TLMND LLC',          ['BP_LP'], 'LLC',            '83-2303718', 'USA', '10/15/2016', 'Type: S Corporation 1120\nOwnership: BP Family 12%\nPurpose: RE & Foreign Investment', '#1e6e3b', 'both', 840, 730],
    ['NF_MOE',  'NF MOE CO SA S',     ['BP_LP'], 'Other',          '', '',    '',           'Purpose: Business and RE Investment Holding', '#1e3a5c', 'both', 1080, 730],

    // ── Level 5 ───────────────────────────────────────────────────────────────
    ['MEJ_DR',   'MEJ DR INC',          ['TLMND'],  'Corporation',    '', 'Dominican Republic', '', 'Purpose: Business and RE Investment Holding', '#1e6e3b', 'both', 840, 900],
    ['NF_EUR',   'NF Europe Holding S', ['NF_MOE'], 'Holding Company','', 'Europe',   '',       'Purpose: Business and RE Investment Holding', '#1e3a5c', 'both', 1200, 900],

    // ── Level 6 — Europe ─────────────────────────────────────────────────────
    ['PARIS_SCI', 'Paris Thacko SCI',     ['NF_EUR'], 'Other', '', 'France', '', 'Purpose: Business and RE Investment Holding', '#7a1f1f', 'both', 1080, 1070],
    ['NF5_SPAIN', 'NF5 Spain Holdings SL',['NF_EUR'], 'Other', '', 'Spain',  '', 'Purpose: Business and RE Investment Holding', '#7a1f1f', 'both', 1310, 1070],

    // ── TLMND US subsidiaries ─────────────────────────────────────────────────
    ['NF_US_CA', 'NF U S CA LLC',   ['TLMND'], 'LLC', '', 'USA - California', '', '6330 Cameo Canyon Rd\nSan Diego Clinic\nOwnership: 100%', '#7a1f1f', 'both', 540, 900],
    ['CLIFTON',  'Clifton MCI LLC', ['TLMND'], 'LLC', '', 'USA',              '', 'Clifton Clinic\nOwnership: 50%',                            '#7a1f1f', 'both', 690, 900],
    ['NF_US_TX', 'NF U S TX LLC',   ['TLMND'], 'LLC', '', 'USA - Texas',      '', 'Purpose: Business and RE Investment Holding',               '#0e4d5c', 'both', 380, 1070],
    ['NW_WOOD',  'NW Woodland Park ASC', ['TLMND'], 'Other', '', 'USA', '', 'Source: MN Asset Card', '#7a1f1f', 'both', 230, 900],
    ['SANU',     'SA NU Beauty LLC',     ['TLMND'], 'LLC',   '', 'USA', '', 'Source: MN Asset Card', '#1e6e3b', 'both', 80,  900],
    ['KUHLMAN',  '709 Kuhlman',          ['NF_US_TX'], 'Real Estate', '', 'USA - Texas', '', 'Ownership: 100%', '#7a1f1f', 'both', 380, 1240],

    // ── MEJ DR — DR/Colombia investments ────────────────────────────────────
    ['FARANDA',  'Faranda Beach House',             ['MEJ_DR'], 'Real Estate', '', 'Dominican Republic', '', '', '#7a1f1f', 'both', 680,  1070],
    ['LAND_AM',  'Land America Property Investments',['MEJ_DR'],'Investment Category','','Colombia',   '', '', '#8b6914', 'both', 900,  1070],
    ['PARAISO',  'Paraiso Beach',                   ['LAND_AM'],'Real Estate', '', 'Colombia',  '', 'Ownership: 15%', '#7a1f1f', 'both', 900, 1240],
    ['RF_HOLD',  'RF Holdings SAS',                 ['NF_MOE'], 'Corporation', '', 'Colombia',  '', 'Source: MN Asset Card', '#1e6e3b', 'both', 1080, 1070],
    ['COLCIENO_A','Colcieno Associates LLC',        ['NF_MOE'], 'LLC',         '', 'Colombia',  '', 'Ownership: 16%\nSource: MN Asset Card', '#0e4d5c', 'both', 1240, 1070],
    ['COLCIENO_2','Colcieno 2 SAS',                ['NF_MOE'], 'Corporation', '', 'Colombia',  '', 'Ownership: 30%', '#1e3a5c', 'both', 1430, 1070],
    ['LAURELES', 'Laureles Factory',                ['NF_MOE'], 'Other',       '', 'Colombia',  '', '', '#b5521a', 'both', 1080, 1240],
    ['SEATLON',  'Seatlon Hotel',                   ['NF_MOE'], 'Other',       '', 'Colombia',  '', 'Ownership: 5%', '#b5521a', 'both', 1240, 1240],
    ['MANILA',   'Manila Holdings',                 ['NF_MOE'], 'Other',       '', 'Colombia',  '', 'AKU Landing Hotel\nOwnership: 71%', '#b5521a', 'both', 1400, 1240],
    ['LOAO',     'LOAO - Restaurant',               ['NF_MOE'], 'Other',       '', 'Colombia',  '', 'Ownership: 11%', '#b5521a', 'both', 1560, 1240],

    // ── Europe property investments ───────────────────────────────────────────
    ['DGUN',     "D'Gun DC Montreuillo", ['PARIS_SCI'], 'Real Estate', '', 'France', '', 'Ownership: 100%', '#7a1f1f', 'both', 980,  1240],
    ['SANGLIMA', 'Sanglima',             ['PARIS_SCI'], 'Real Estate', '', 'France', '', 'Ownership: 100%', '#7a1f1f', 'both', 1080, 1240],
    ['PLAZA_COL','Plaza Colmar',         ['PARIS_SCI'], 'Real Estate', '', 'France', '', 'Ownership: 100%', '#7a1f1f', 'both', 1180, 1240],

    // ── NF6 Venture 8 — US Business Investments ──────────────────────────────
    ['SOCIODOC', 'Sociodoc',             ['NF6_V8'], 'Investment Category', '', 'USA', '', 'Source: MN Asset Card', '#1e6e3b', 'both', -200, 900],
    ['SOLARIS',  'Solaris FL Holding',   ['NF6_V8'], 'LLC',                 '', 'USA - Florida', '', 'Source: MN Asset Card', '#1e6e3b', 'both', -10,  900],
    ['NAT_NUTRA','Natural Nutra',        ['NF6_V8'], 'Other',               '', 'USA', '', 'Source: MN Asset Card', '#1e6e3b', 'both', 180, 900],
    ['LEXI_SKIN','Lexi Skin',            ['NF6_V8'], 'Other',               '', 'USA', '', 'Source: MN Asset Card', '#1e6e3b', 'both', -200, 1070],
    ['MICIDA',   'Micida Capital Partners',['NF6_V8'],'Investment Category', '', 'USA', '', '', '#5b2c87', 'both', -10,  1070],
    ['TALKSELF', 'Talkself Florida LLC', ['NF6_V8'], 'LLC',                 '', 'USA - Florida', '', '', '#1e6e3b', 'both', 180, 1070],
    ['FS_NYC',   'FS NYC Chelsea',       ['NF6_V8'], 'Other',               '', 'USA - New York', '', 'Source: MN Asset Card', '#1e6e3b', 'both', -200, 1240],
    ['SAN_ORL',  'Sanchezon at Orlando', ['NF6_V8'], 'Other',               '', 'USA - Florida',  '', '', '#b5521a', 'both', -10, 1240],
    ['SAN_RES',  'Sanchezon Residential',['NF6_V8'], 'Other',               '', 'USA', '', '', '#b5521a', 'both', 180, 1240],

    // ── NF6 Venture 8 — PR Business Investments ──────────────────────────────
    ['FIT_PR',    'FIT Investments PR',   ['NF6_V8'], 'Investment Category', '', 'Puerto Rico', '', 'Source: MN Asset Card', '#5b2c87', 'both', -380, 900],
    ['SUN_HARBOR','Sun Harbor Capital',   ['NF6_V8'], 'Other',               '', 'Puerto Rico', '', 'Source: MN Asset Card', '#5b2c87', 'both', -380, 1070],
    ['SAN_D_EAST','Sanchezon Dorado East',['NF6_V8'], 'Real Estate',         '', 'Puerto Rico', '', '', '#7a1f1f', 'both', -380, 1240],
    ['SAN_D_WEST','Sanchezon Dorado West',['NF6_V8'], 'Real Estate',         '', 'Puerto Rico', '', '', '#7a1f1f', 'both', -380, 1410],

    // ── NF5 Tiger Capital — Investments ──────────────────────────────────────
    // (Brokerage accounts - listed here as structural nodes)

    // ── POST STRUCTURE — NF6 Family 2026 Trust ───────────────────────────────
    ['TRUST_2026','NF6 Family 2026 Trust',['MICHAEL'], 'Trust (Irrevocable)', '', 'USA', '2026', 'Grantor: Michael Nguyen\n(New structure — in formation)', '#2d6a4f', 'post', 850, 220]
  ];

  var rows = NODES.map(function(n) {
    var key       = n[0];
    var id        = ID[key];
    var parentIds = n[2].map(function(pk) { return ID[pk]; }).join(',');
    return makeRow({
      'ID':           id,
      'Name':         n[1],
      'Parents':      parentIds,
      'Node Type':    n[3],
      'Tax ID':       n[4],
      'Jurisdiction': n[5],
      'Date Created': n[6],
      'Ownership':    n[7],
      'Color':        n[8],
      'Text Color':   '#ffffff',
      'Notes':        '',
      'Structure':    n[9],
      'X':            n[10],
      'Y':            n[11]
    });
  });

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

  SpreadsheetApp.getUi().alert('✅ Org chart seeded with ' + rows.length + ' entities.\n\nOpen the web app and go to the Org Chart tab. Click "Fit View" to see everything.\n\nNote: The 2026 Trust node is set to "Post" view — switch to that view to see it.');
}

// ── Org Chart ─────────────────────────────────────────────────────────────────

function getOrgChart() {
  ensureSheets_();
  var sheet = getSheet_('ORG_CHART');
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, j) { obj[h] = row[j] instanceof Date ? row[j].toISOString() : row[j]; });
    // Parse Parents from comma-separated string to array
    obj.parents = obj['Parents'] ? String(obj['Parents']).split(',').map(function(s){ return s.trim(); }).filter(Boolean) : [];
    return obj;
  });
}

function saveOrgNode(nodeJson) {
  var node  = JSON.parse(nodeJson);
  var sheet = getSheet_('ORG_CHART');
  var data  = sheet.getDataRange().getValues();
  var headers = data[0];

  var parentsStr = Array.isArray(node.parents) ? node.parents.join(',') : (node.parents || '');
  var rowMap = {
    'ID':          node.id || Utilities.getUuid(),
    'Name':        node.name || '',
    'Parents':     parentsStr,
    'Node Type':   node.nodeType || '',
    'Tax ID':      node.taxId || '',
    'Jurisdiction':node.jurisdiction || '',
    'Date Created':node.dateCreated || '',
    'Ownership':   node.ownership || '',
    'Color':       node.color || '#1a5c6b',
    'Text Color':  node.textColor || '#ffffff',
    'Notes':       node.notes || '',
    'Structure':   node.structure || 'both',
    'X':           node.x || 100,
    'Y':           node.y || 100
  };
  var rowData = headers.map(function(h) { return rowMap[h] !== undefined ? rowMap[h] : ''; });

  // Update existing row or append
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(rowMap['ID'])) {
      sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
      return { success: true, id: rowMap['ID'] };
    }
  }
  sheet.appendRow(rowData);
  return { success: true, id: rowMap['ID'] };
}

function deleteOrgNode(id) {
  var sheet = getSheet_('ORG_CHART');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Node not found' };
}

function saveOrgPositions(positionsJson) {
  // positionsJson: [{id, x, y}, ...]
  var positions = JSON.parse(positionsJson);
  var sheet = getSheet_('ORG_CHART');
  var data  = sheet.getDataRange().getValues();
  var headers = data[0];
  var xCol = headers.indexOf('X') + 1;
  var yCol = headers.indexOf('Y') + 1;
  if (xCol < 1 || yCol < 1) return { success: false, error: 'X/Y columns missing' };

  var posMap = {};
  positions.forEach(function(p) { posMap[String(p.id)] = p; });

  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0]);
    if (posMap[id]) {
      sheet.getRange(i + 1, xCol).setValue(posMap[id].x);
      sheet.getRange(i + 1, yCol).setValue(posMap[id].y);
    }
  }
  return { success: true };
}

// ── Property Valuation (Rentcast) ─────────────────────────────────────────────

/**
 * Refresh Rentcast AVM values for all US Real Estate assets.
 * Assets with Category = "Real Estate" and a US address in Notes
 * (format: "address: 123 Main St, City, TX 77001") are auto-updated.
 */
function refreshPropertyValues() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('RENTCAST_API_KEY');
  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      'Rentcast API key not set.\n\n' +
      'Go to: Extensions > Apps Script > Project Settings > Script Properties\n' +
      'Add: RENTCAST_API_KEY = your key from rentcast.io'
    );
    return { success: false, error: 'No API key' };
  }

  var sheet   = getSheet_('ASSETS');
  var rows    = sheet.getDataRange().getValues();
  var updated = 0;
  var errors  = [];

  for (var i = 1; i < rows.length; i++) {
    var category = String(rows[i][2] || '').trim();
    var currency = String(rows[i][4] || '').trim();
    var address  = String(rows[i][14] || '').trim();

    if (category !== 'Real Estate' || currency !== 'USD' || !address) continue;

    if (!/\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY|DC)\b/i.test(address)) continue;

    var result = getRentcastEstimate_(address, apiKey);

    if (!result.success) {
      errors.push(rows[i][1] + ': ' + result.error);
      Utilities.sleep(500);
      continue;
    }

    var sheetRow = i + 1;
    var oldUsd   = Number(rows[i][7]) || 0;
    var newUsd   = result.value;
    var sharePct = Number(rows[i][8]) || 100;

    sheet.getRange(sheetRow, 6).setValue(newUsd);
    sheet.getRange(sheetRow, 7).setValue(1);
    sheet.getRange(sheetRow, 8).setValue(newUsd);
    sheet.getRange(sheetRow, 10).setValue(newUsd * sharePct / 100);
    sheet.getRange(sheetRow, 12).setValue(new Date());

    var existingNotes = String(rows[i][12] || '');
    var rangeNote = 'Rentcast ' + formatDate_(new Date()) +
                    ': $' + formatNumber_(result.lowValue) + '–$' + formatNumber_(result.highValue);
    var newNotes = existingNotes.replace(/Rentcast [^\|]*/g, rangeNote);
    if (newNotes === existingNotes) newNotes = existingNotes ? existingNotes + ' | ' + rangeNote : rangeNote;
    sheet.getRange(sheetRow, 13).setValue(newNotes);

    if (Math.abs(newUsd - oldUsd) > 0.01) {
      logHistory_(rows[i][1], oldUsd, newUsd, 'USD', 'Auto-updated via Rentcast');
    }

    updated++;
    Utilities.sleep(600);
  }

  var msg = updated + ' US propert' + (updated === 1 ? 'y' : 'ies') + ' updated.';
  if (errors.length) msg += '\n\nNot updated:\n' + errors.join('\n');
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Rentcast Update', 10);
  return { success: true, updated: updated, errors: errors };
}

function lookupSingleProperty() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Property Lookup', 'Enter full US address (e.g. 123 Main St, Houston, TX 77001):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var address = resp.getResponseText().trim();
  if (!address) return;
  var result = getPropertyValue(address);
  if (result.success) {
    ui.alert('Property Estimate',
      address + '\n\n' +
      'Value:  $' + formatNumber_(result.value) + '\n' +
      'Range:  $' + formatNumber_(result.lowValue) + ' – $' + formatNumber_(result.highValue),
      ui.ButtonSet.OK);
  } else {
    ui.alert('Could not get estimate: ' + result.error);
  }
}

function getPropertyValue(address) {
  if (!address) return { success: false, error: 'No address provided' };
  var props = PropertiesService.getScriptProperties();

  // API Ninjas — free tier: 3,000 req/month (api-ninjas.com)
  var ninjasKey = props.getProperty('API_NINJAS_KEY');
  if (ninjasKey) return getApiNinjasEstimate_(address, ninjasKey);

  // Rentcast fallback
  var rentcastKey = props.getProperty('RENTCAST_API_KEY');
  if (rentcastKey) return getRentcastEstimate_(address, rentcastKey);

  return { success: false, error: 'No API key set. Add API_NINJAS_KEY in Apps Script > Project Settings > Script Properties (get free key at api-ninjas.com)' };
}

function getApiNinjasEstimate_(address, apiKey) {
  try {
    var url  = 'https://api.api-ninjas.com/v1/houseprice?address=' + encodeURIComponent(address);
    var resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'X-Api-Key': apiKey },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code === 401 || code === 403) return { success: false, error: 'Invalid API Ninjas key — check API_NINJAS_KEY in Script Properties' };
    if (code === 429) return { success: false, error: 'API Ninjas rate limit hit' };
    if (code !== 200) return { success: false, error: 'API Ninjas HTTP ' + code + ': ' + resp.getContentText().substring(0, 120) };
    var raw  = resp.getContentText();
    var data = JSON.parse(raw);
    // Response may be an array or a single object
    var item = Array.isArray(data) ? data[0] : data;
    if (!item) return { success: false, error: 'No data returned for this address' };
    var value = item.price || item.value || item.estimated_value || item.zestimate || null;
    if (!value) return { success: false, error: 'No price field in response — raw: ' + raw.substring(0, 200) };
    value = Number(value);
    return {
      success:   true,
      value:     Math.round(value),
      lowValue:  Math.round(item.price_low  || item.low  || value * 0.95),
      highValue: Math.round(item.price_high || item.high || value * 1.05)
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getRentcastEstimate_(address, apiKey) {
  try {
    var url  = 'https://api.rentcast.io/v1/avm/value?address=' + encodeURIComponent(address);
    var resp = UrlFetchApp.fetch(url, {
      method: 'GET', headers: { 'X-Api-Key': apiKey }, muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code === 401) return { success: false, error: 'Rentcast API key inactive — subscription required at rentcast.io' };
    if (code === 404) return { success: false, error: 'Address not found in Rentcast database' };
    if (code === 429) return { success: false, error: 'Rentcast rate limit hit (50 requests/month on free tier)' };
    if (code !== 200) return { success: false, error: 'Rentcast HTTP ' + code };
    var data  = JSON.parse(resp.getContentText());
    var value = data.price || data.value || data.priceRangeMid || null;
    if (!value) return { success: false, error: 'No valuation returned' };
    return {
      success:   true,
      value:     Math.round(value),
      lowValue:  Math.round(data.priceLow  || value * 0.95),
      highValue: Math.round(data.priceHigh || value * 1.05)
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function formatNumber_(n) {
  if (!n) return '0';
  return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatDate_(d) {
  return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
}

// ── History ───────────────────────────────────────────────────────────────────

function logHistory_(assetName, oldVal, newVal, currency, notes) {
  getSheet_('HISTORY').appendRow([
    new Date(), assetName, oldVal, newVal, newVal - oldVal, currency || 'USD', notes || ''
  ]);
}

function getHistory(filters) {
  var data = sheetToObjects_('HISTORY').map(function(obj) {
    var out = {};
    Object.keys(obj).forEach(function(k) {
      out[k] = obj[k] instanceof Date ? obj[k].toISOString() : obj[k];
    });
    return out;
  });
  if (!filters) return data;
  if (filters.assetName) {
    data = data.filter(function(r) {
      return (r['Asset Name'] || '').toLowerCase().indexOf(filters.assetName.toLowerCase()) !== -1;
    });
  }
  return data;
}

// ── Monthly Snapshots ─────────────────────────────────────────────────────────

function takeMonthlySnapshot() {
  var now      = new Date();
  var monthKey = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');

  // Deduplicate — only one snapshot per calendar month
  var snapSheet = getSheet_('SNAPSHOTS');
  var existing  = snapSheet.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (existing[i][1] === monthKey) {
      return { success: false, alreadyDone: true, msg: 'Snapshot already taken for ' + monthKey };
    }
  }

  var assets = sheetToObjects_('ASSETS');
  if (!assets.length) return { success: false, alreadyDone: false, msg: 'No assets to snapshot' };

  var rows = assets.map(function(a) {
    return [now, monthKey, a['Name'] || '', a['Category'] || '', a['Currency'] || 'USD', Number(a['My Share USD']) || 0];
  });
  var startRow = snapSheet.getLastRow() + 1;
  snapSheet.getRange(startRow, 1, rows.length, 6).setValues(rows);
  // Force Month Key column (col 2) to plain text so GAS won't auto-convert "YYYY-MM" to a date
  snapSheet.getRange(startRow, 2, rows.length, 1).setNumberFormat('@');
  return { success: true, alreadyDone: false, count: rows.length, monthKey: monthKey };
}

function normalizeMonthKey_(mk) {
  // Google Sheets may auto-convert "2026-04" strings to Date objects
  if (mk instanceof Date) {
    return mk.getFullYear() + '-' + String(mk.getMonth() + 1).padStart(2, '0');
  }
  return String(mk || '').trim();
}

function getSnapshotTrend() {
  var data = sheetToObjects_('SNAPSHOTS');
  var byMonth = {};
  data.forEach(function(row) {
    var mk = normalizeMonthKey_(row['Month Key']);
    if (!mk) return;
    if (!byMonth[mk]) byMonth[mk] = 0;
    byMonth[mk] += Number(row['My Share USD']) || 0;
  });
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return Object.keys(byMonth).sort().map(function(mk) {
    var parts = mk.split('-');
    return { monthKey: mk, label: months[parseInt(parts[1]) - 1] + ' ' + parts[0], total: byMonth[mk] };
  });
}

function getSnapshotMatrix() {
  var monthSet  = {};  // mk -> 'nw' (primary) or 'old' (fallback)
  var assetData = {};  // name -> { category, mk -> value }
  var liabData  = {};  // name -> { type, mk -> value }
  var liabByMonth = {}; // mk -> total USD

  // ── PRIMARY: NW_SNAPSHOTS (complete — assets + liabilities together) ──────
  try {
    var nwSheet = getNWSnapshotsSheet_();
    var nwRaw   = nwSheet.getDataRange().getValues();
    for (var i = 1; i < nwRaw.length; i++) {
      var r = nwRaw[i];
      var mk = r[1];
      if (mk instanceof Date) mk = mk.getFullYear() + '-' + String(mk.getMonth() + 1).padStart(2, '0');
      mk = String(mk).trim();
      if (!mk) continue;
      var recType = String(r[2]).trim();
      var name    = String(r[3]).trim();
      var cat     = String(r[4]).trim() || 'Other';
      var val     = Number(r[5]) || 0;
      monthSet[mk] = 'nw'; // mark month as having full NW_SNAPSHOTS data
      if (recType === 'ASSET') {
        if (!assetData[name]) assetData[name] = { category: cat };
        assetData[name][mk] = val;
      } else if (recType === 'LIABILITY') {
        liabByMonth[mk] = (liabByMonth[mk] || 0) + val;
        if (!liabData[name]) liabData[name] = { type: cat };
        liabData[name][mk] = val;
      }
    }
  } catch(e) { /* NW_SNAPSHOTS empty — will fall back below */ }

  var months = Object.keys(monthSet).sort().slice(-12);
  if (!months.length) return { months: [], assetTotals: [], liabTotals: [], netWorthTotals: [], assets: [], liabilities: [] };

  var assetNames = Object.keys(assetData);
  var assets = assetNames.map(function(name) {
    return { name: name, category: assetData[name].category,
             values: months.map(function(m) { return assetData[name][m] || 0; }) };
  });
  assets.sort(function(a, b) {
    if (a.category !== b.category) return a.category.localeCompare(b.category);
    return (b.values[b.values.length - 1] || 0) - (a.values[a.values.length - 1] || 0);
  });

  var assetTotals = months.map(function(m) {
    return assetNames.reduce(function(s, n) { return s + (assetData[n][m] || 0); }, 0);
  });

  // Liability totals: use NW_SNAPSHOTS per-month where available; current total otherwise
  var currentLiabTotal = sheetToObjects_('LIABILITIES')
    .reduce(function(s, l) { return s + (Number(l['USD Value']) || 0); }, 0);
  var liabTotals     = months.map(function(m) { return liabByMonth[m] !== undefined ? liabByMonth[m] : currentLiabTotal; });
  var netWorthTotals = assetTotals.map(function(a, i) { return a - liabTotals[i]; });

  var liabilities = Object.keys(liabData).map(function(name) {
    return { name: name, type: liabData[name].type,
             values: months.map(function(m) { return liabData[name][m] || 0; }) };
  }).sort(function(a, b) {
    if (a.type !== b.type) return a.type.localeCompare(b.type);
    return (b.values[b.values.length - 1] || 0) - (a.values[a.values.length - 1] || 0);
  });

  return { months: months, assetTotals: assetTotals, liabTotals: liabTotals,
           netWorthTotals: netWorthTotals, assets: assets, liabilities: liabilities };
}

// ── Plaid Integration ─────────────────────────────────────────────────────────

function getPlaidConfig_() {
  var p = PropertiesService.getScriptProperties();
  return {
    clientId: p.getProperty('PLAID_CLIENT_ID') || '',
    secret:   p.getProperty('PLAID_SECRET')    || '',
    env:      p.getProperty('PLAID_ENV')        || 'sandbox'
  };
}

function getPlaidBaseUrl_(env) {
  return 'https://' + (env === 'production' ? 'production' : env) + '.plaid.com';
}

function getPlaidLinkToken() {
  var cfg = getPlaidConfig_();
  try {
    var resp = UrlFetchApp.fetch(getPlaidBaseUrl_(cfg.env) + '/link/token/create', {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({
        client_id:     cfg.clientId,
        secret:        cfg.secret,
        client_name:   'MNW Family Office',
        country_codes: ['US'],
        language:      'en',
        user:          { client_user_id: 'mnw-family-office' },
        products:      ['transactions']
      }),
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    if (data.link_token) return { success: true, linkToken: data.link_token, env: cfg.env };
    return { success: false, error: data.error_message || JSON.stringify(data) };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function exchangePlaidToken(publicToken) {
  var cfg = getPlaidConfig_();
  try {
    var resp = UrlFetchApp.fetch(getPlaidBaseUrl_(cfg.env) + '/item/public_token/exchange', {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ client_id: cfg.clientId, secret: cfg.secret, public_token: publicToken }),
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    if (!data.access_token) return { success: false, error: data.error_message };
    var p      = PropertiesService.getScriptProperties();
    var tokens = JSON.parse(p.getProperty('PLAID_TOKENS') || '[]');
    if (tokens.indexOf(data.access_token) === -1) tokens.push(data.access_token);
    p.setProperty('PLAID_TOKENS', JSON.stringify(tokens));
    return { success: true, accessToken: data.access_token };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function syncPlaidAccounts() {
  var cfg     = getPlaidConfig_();
  var p       = PropertiesService.getScriptProperties();
  var tokens  = JSON.parse(p.getProperty('PLAID_TOKENS') || '[]');
  var instMap = JSON.parse(p.getProperty('PLAID_INSTITUTIONS') || '{}');
  if (!tokens.length) return { success: false, error: 'No Plaid accounts connected. Use Connect Bank first.' };

  // ── Step 1: fire all Plaid API requests in parallel via fetchAll ───────────
  var requests = tokens.map(function(token) {
    return {
      url: getPlaidBaseUrl_(cfg.env) + '/accounts/balance/get',
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ client_id: cfg.clientId, secret: cfg.secret, access_token: token }),
      muteHttpExceptions: true
    };
  });
  var responses = UrlFetchApp.fetchAll(requests);

  // Collect all accounts across all institutions
  var allAccounts = [];
  tokens.forEach(function(token, idx) {
    var institution = instMap[token] || '';
    try {
      var data = JSON.parse(responses[idx].getContentText());
      if (!data.accounts) return;
      data.accounts.forEach(function(acct) {
        allAccounts.push({
          balance:  (acct.balances.current != null ? acct.balances.current : acct.balances.available) || 0,
          name:     (institution ? institution + ' - ' : '') + (acct.name || 'Account') + ' \u00b7\u00b7\u00b7' + (acct.mask || ''),
          acctId:   acct.account_id
        });
      });
    } catch(e) {
      console.error('Plaid parse error for token ' + idx + ':', e);
    }
  });

  // ── Step 2: read sheet ONCE, build lookup maps ────────────────────────────
  var sheet = getSheet_('ASSETS');
  var rows  = sheet.getDataRange().getValues();
  var now   = new Date();

  // Map: plaidId → row index (0-based, data rows start at 1)
  var byPlaidId = {};
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][13]) byPlaidId[rows[i][13]] = i;
  }

  // ── Step 3: match and collect updates ─────────────────────────────────────
  var updates   = [];   // {rowIdx, acctName, balance, acctId, oldUsd}
  var newAccts  = [];   // accounts with no matching row

  allAccounts.forEach(function(acct) {
    var matchRow = byPlaidId[acct.acctId];

    if (matchRow === undefined) {
      // Name+category match
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][1] === acct.name && String(rows[i][2]).startsWith('Cash')) {
          matchRow = i; break;
        }
      }
    }

    if (matchRow !== undefined) {
      updates.push({ rowIdx: matchRow, acctName: acct.name, balance: acct.balance,
                     acctId: acct.acctId, oldUsd: Number(rows[matchRow][7]) || 0 });
    } else {
      newAccts.push(acct);
    }
  });

  // ── Step 4: apply all updates with one setValues call per row ─────────────
  updates.forEach(function(u) {
    var r = u.rowIdx + 1;   // 1-based sheet row
    // Columns: 2=Name, 6=LocalVal, 7=FXRate, 8=USD Value, 10=My Share USD, 12=Last Updated, 14=PlaidID
    sheet.getRange(r, 2).setValue(u.acctName);
    sheet.getRange(r, 6, 1, 3).setValues([[u.balance, 1, u.balance]]);   // cols 6,7,8
    sheet.getRange(r, 10).setValue(u.balance);
    sheet.getRange(r, 12).setValue(now);
    sheet.getRange(r, 14).setValue(u.acctId);
    if (Math.abs(u.balance - u.oldUsd) > 0.01) logHistory_(u.acctName, u.oldUsd, u.balance, 'USD', 'Plaid sync');
  });

  // ── Step 5: add new accounts (append rows) ────────────────────────────────
  newAccts.forEach(function(acct) {
    addAsset({ name: acct.name, category: 'Cash', currency: 'USD',
               localValue: acct.balance, mySharePct: 100, notes: 'Plaid: ' + acct.acctId });
    // Tag the newly appended row with the Plaid ID
    var newRowCount = sheet.getLastRow();
    sheet.getRange(newRowCount, 14).setValue(acct.acctId);
  });

  return { success: true, synced: updates.length + newAccts.length };
}

// ── Plaid UI Helpers ──────────────────────────────────────────────────────────

function setPlaidCredentials() {
  var ui       = SpreadsheetApp.getUi();
  var clientId = ui.prompt('Plaid Setup', 'Enter your Plaid Client ID:', ui.ButtonSet.OK_CANCEL);
  if (clientId.getSelectedButton() !== ui.Button.OK) return;
  var secret   = ui.prompt('Plaid Setup', 'Enter your Plaid Secret:', ui.ButtonSet.OK_CANCEL);
  if (secret.getSelectedButton() !== ui.Button.OK) return;
  var env      = ui.prompt('Plaid Setup', 'Environment (sandbox / production):', ui.ButtonSet.OK_CANCEL);
  if (env.getSelectedButton() !== ui.Button.OK) return;
  var props = PropertiesService.getScriptProperties();
  props.setProperty('PLAID_CLIENT_ID', clientId.getResponseText().trim());
  props.setProperty('PLAID_SECRET',    secret.getResponseText().trim());
  props.setProperty('PLAID_ENV',       env.getResponseText().trim() || 'sandbox');
  ui.alert('Plaid credentials saved. You can now connect bank accounts.');
}

function openPlaidLink() {
  var cfg = getPlaidConfig_();
  if (!cfg.clientId || !cfg.secret) {
    var ui   = SpreadsheetApp.getUi();
    var resp = ui.alert('Plaid Not Configured', 'Plaid credentials are not set. Would you like to set them now?', ui.ButtonSet.YES_NO);
    if (resp === ui.Button.YES) setPlaidCredentials();
    return;
  }
  var html = HtmlService.createHtmlOutputFromFile('PlaidLink')
    .setTitle('Connect Bank Account')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function handlePlaidSuccess(publicToken, institutionName) {
  try {
    var exchResult = exchangePlaidToken(publicToken);
    if (!exchResult.success) return { success: false, message: exchResult.error };
    if (institutionName && exchResult.accessToken) {
      var p       = PropertiesService.getScriptProperties();
      var instMap = JSON.parse(p.getProperty('PLAID_INSTITUTIONS') || '{}');
      instMap[exchResult.accessToken] = institutionName;
      p.setProperty('PLAID_INSTITUTIONS', JSON.stringify(instMap));
    }
    var syncResult = syncPlaidAccounts();
    return { success: true, message: 'Bank connected! ' + (syncResult.synced || 0) + ' account(s) synced to Assets tab.' };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function getPlaidConnections() {
  var p       = PropertiesService.getScriptProperties();
  var tokens  = JSON.parse(p.getProperty('PLAID_TOKENS') || '[]');
  var instMap = JSON.parse(p.getProperty('PLAID_INSTITUTIONS') || '{}');
  return tokens.map(function(token, i) {
    return { index: i, name: instMap[token] || '', tokenHint: '···' + token.slice(-4) };
  });
}

function setPlaidInstitutionName(index, name) {
  var p       = PropertiesService.getScriptProperties();
  var tokens  = JSON.parse(p.getProperty('PLAID_TOKENS') || '[]');
  var instMap = JSON.parse(p.getProperty('PLAID_INSTITUTIONS') || '{}');
  if (index < 0 || index >= tokens.length) return { success: false, error: 'Invalid index' };
  instMap[tokens[index]] = name;
  p.setProperty('PLAID_INSTITUTIONS', JSON.stringify(instMap));
  return { success: true };
}

function removePlaidConnection() {
  var ui     = SpreadsheetApp.getUi();
  var props  = PropertiesService.getScriptProperties();
  var tokens = JSON.parse(props.getProperty('PLAID_TOKENS') || '[]');
  if (!tokens.length) { ui.alert('No Plaid connections to remove.'); return; }
  var resp = ui.alert('Remove Plaid Connections',
    'This will disconnect all ' + tokens.length + ' bank connection(s). Continue?',
    ui.ButtonSet.YES_NO);
  if (resp === ui.Button.YES) {
    props.deleteProperty('PLAID_TOKENS');
    ui.alert('All Plaid connections removed.');
  }
}

function fixSheetHeaders() {
  var ss = getSpreadsheet_();
  Object.keys(COL).forEach(function(key) {
    var sheetName = sheetName_(key);
    var sheet     = ss.getSheetByName(sheetName);
    if (!sheet) return;
    var headers = COL[key];
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#0d2137')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('All sheet headers fixed!', 'Done', 5);
  return { success: true };
}

function debugSheetHeaders() {
  var sheet = getSheet_('ASSETS');
  var data  = sheet.getDataRange().getValues();
  return {
    headers:  data[0]  || [],
    firstRow: data[1]  || [],
    rowCount: data.length - 1
  };
}

// ── Triggers ──────────────────────────────────────────────────────────────────

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailySync_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailySync_').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Daily sync scheduled for 7 AM', 'Trigger Installed', 5);
  return { success: true };
}

function dailySync_() {
  fetchExchangeRates();
  syncPlaidAccounts();
  refreshPropertyValues();
  if (new Date().getDate() === 1) {
    takeMonthlySnapshot();   // asset-only snapshot for the in-app trend chart
    takeNWSnapshot();        // assets + liabilities for the Tiller-style NW History sheet
    generateNetWorthHistorySheet();
  }
  generateBalancesSheet();
}

// ── Balances Sheet (Tiller-style Net Worth view) ──────────────────────────────

function generateBalancesSheet() {
  var ss        = getSpreadsheet_();
  var SHEET_NAME = 'Balances';

  // ── get or create the Balances sheet ──────────────────────────────────────
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) {
    sheet.clearContents();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // ── pull live data ─────────────────────────────────────────────────────────
  var assets      = sheetToObjects_('ASSETS');
  var liabilities = sheetToObjects_('LIABILITIES');

  // Compute USD values
  assets.forEach(function(a) {
    a._usd = parseFloat(a['My Share USD']) || 0;
    a._lastUpdated = a['Last Updated'] ? new Date(a['Last Updated']) : null;
  });
  liabilities.forEach(function(l) {
    l._usd = parseFloat(l['USD Value']) || 0;
    l._lastUpdated = l['Last Updated'] ? new Date(l['Last Updated']) : null;
  });

  // Group assets by Category
  var assetCats = {};
  CATEGORIES.forEach(function(c) { assetCats[c] = []; });
  assets.forEach(function(a) {
    var cat = a['Category'] || 'Other';
    if (!assetCats[cat]) assetCats[cat] = [];
    assetCats[cat].push(a);
  });
  // Only keep categories that have items
  var usedAssetCats = CATEGORIES.filter(function(c) { return assetCats[c] && assetCats[c].length > 0; });

  // Group liabilities by Type
  var liabTypes = {};
  liabilities.forEach(function(l) {
    var t = l['Type'] || 'Other';
    if (!liabTypes[t]) liabTypes[t] = [];
    liabTypes[t].push(l);
  });
  var usedLiabTypes = Object.keys(liabTypes).sort();

  // Totals
  var totalAssets = assets.reduce(function(s, a) { return s + a._usd; }, 0);
  var totalLiabs  = liabilities.reduce(function(s, l) { return s + l._usd; }, 0);
  var netWorth    = totalAssets - totalLiabs;

  // Helper: days ago string
  function daysAgo(d) {
    if (!d || !(d instanceof Date) || isNaN(d.getTime())) return '';
    var days = Math.floor((Date.now() - d.getTime()) / 86400000);
    if (days === 0) return 'today';
    if (days === 1) return '1 day ago';
    return days + ' days ago';
  }

  // Helper: latest date in a group
  function latestDate(items) {
    var best = null;
    items.forEach(function(item) {
      if (item._lastUpdated && (!best || item._lastUpdated > best)) best = item._lastUpdated;
    });
    return best;
  }

  // Helper: format currency
  function fmt(v) {
    if (v == null || isNaN(v)) return '$0';
    return '$' + Math.round(v).toLocaleString();
  }

  // ── build row data ─────────────────────────────────────────────────────────
  // We'll write rows into parallel arrays (left side = assets, right side = liabilities)
  // Each entry: { type: 'title'|'net_worth'|'section_header'|'cat_header'|'item'|'total'|'blank', ... }

  var assetRows  = [];  // left side rows
  var liabRows   = [];  // right side rows

  // Asset section header
  assetRows.push({ type: 'section_header', label: 'ASSETS', total: totalAssets });

  usedAssetCats.forEach(function(cat) {
    var items    = assetCats[cat];
    var catTotal = items.reduce(function(s, a) { return s + a._usd; }, 0);
    var latest   = latestDate(items);
    assetRows.push({ type: 'cat_header', label: cat, updated: daysAgo(latest), total: catTotal });
    items.forEach(function(a) {
      assetRows.push({ type: 'item', label: a['Name'] || '', updated: daysAgo(a._lastUpdated), value: a._usd });
    });
  });

  // Liability section header
  liabRows.push({ type: 'section_header', label: 'LIABILITIES', total: totalLiabs });

  usedLiabTypes.forEach(function(t) {
    var items    = liabTypes[t];
    var typeTotal = items.reduce(function(s, l) { return s + l._usd; }, 0);
    var latest    = latestDate(items);
    liabRows.push({ type: 'cat_header', label: t, updated: daysAgo(latest), total: typeTotal });
    items.forEach(function(l) {
      liabRows.push({ type: 'item', label: l['Name'] || '', updated: daysAgo(l._lastUpdated), value: l._usd });
    });
  });

  // ── write to sheet ─────────────────────────────────────────────────────────
  // Column layout (1-indexed):
  // A(1): Asset name     — wide
  // B(2): (name cont.)
  // C(3): (name cont.)
  // D(4): days ago       — right-aligned
  // E(5): spacer
  // F(6): value          — right-aligned
  // G(7): gap
  // H(8): Liab name
  // I(9): (name cont.)
  // J(10): (name cont.)
  // K(11): days ago
  // L(12): spacer
  // M(13): value

  var TOTAL_COLS = 13;
  var NOW        = new Date();
  var dateStr    = Utilities.formatDate(NOW, Session.getScriptTimeZone(), 'MMMM d, yyyy');

  // Ensure at least TOTAL_COLS columns exist
  if (sheet.getMaxColumns() < TOTAL_COLS) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), TOTAL_COLS - sheet.getMaxColumns());
  }

  // Row 1: Title
  sheet.getRange(1, 1, 1, TOTAL_COLS).merge()
    .setValue('Net Worth')
    .setFontSize(18).setFontWeight('bold').setFontColor('#000000')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBackground('#ffffff');

  // Row 2: Date
  sheet.getRange(2, 1, 1, TOTAL_COLS).merge()
    .setValue(dateStr)
    .setFontSize(10).setFontColor('#666666')
    .setHorizontalAlignment('left')
    .setBackground('#ffffff');

  // Row 3: blank spacer
  sheet.getRange(3, 1, 1, TOTAL_COLS).setBackground('#ffffff');

  // Row 4: NET WORTH banner
  sheet.getRange(4, 1, 1, 6).merge()
    .setValue('NET WORTH')
    .setFontSize(12).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBackground('#1A7341');
  sheet.getRange(4, 7, 1, 7).merge()
    .setValue(netWorth < 0 ? '-' + fmt(Math.abs(netWorth)) : fmt(netWorth))
    .setFontSize(12).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('right').setVerticalAlignment('middle')
    .setBackground('#1A7341');

  // Row 5: blank spacer
  sheet.getRange(5, 1, 1, TOTAL_COLS).setBackground('#ffffff');

  // Data starts at row 6
  var maxRows = Math.max(assetRows.length, liabRows.length);

  for (var i = 0; i < maxRows; i++) {
    var r    = 6 + i;
    var aRow = assetRows[i];
    var lRow = liabRows[i];

    // ── Left: assets ────────────────────────────────────────────────────────
    if (aRow) {
      if (aRow.type === 'section_header') {
        // Dark header: col A-C merged = "ASSETS", D=Updated, F=total
        sheet.getRange(r, 1, 1, 3).merge()
          .setValue('ASSETS')
          .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
          .setBackground('#1B3A5C').setHorizontalAlignment('left');
        sheet.getRange(r, 4).setValue('Updated')
          .setFontSize(9).setFontColor('#ffffff').setFontWeight('bold')
          .setBackground('#1B3A5C').setHorizontalAlignment('right');
        sheet.getRange(r, 5).setBackground('#1B3A5C');
        sheet.getRange(r, 6)
          .setValue(fmt(aRow.total))
          .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
          .setBackground('#1B3A5C').setHorizontalAlignment('right');

      } else if (aRow.type === 'cat_header') {
        sheet.getRange(r, 1, 1, 3).merge()
          .setValue(aRow.label)
          .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
          .setBackground('#2E6DA4').setHorizontalAlignment('left');
        sheet.getRange(r, 4).setValue(aRow.updated || '')
          .setFontSize(8).setFontColor('#cce0f5')
          .setBackground('#2E6DA4').setHorizontalAlignment('right');
        sheet.getRange(r, 5).setBackground('#2E6DA4');
        sheet.getRange(r, 6).setValue(fmt(aRow.total))
          .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
          .setBackground('#2E6DA4').setHorizontalAlignment('right');

      } else if (aRow.type === 'item') {
        var bg = (i % 2 === 0) ? '#f5f8fc' : '#ffffff';
        sheet.getRange(r, 1, 1, 3).merge()
          .setValue('  ' + aRow.label)
          .setFontSize(9).setFontColor('#1a1a1a')
          .setBackground(bg).setHorizontalAlignment('left');
        sheet.getRange(r, 4).setValue(aRow.updated || '')
          .setFontSize(8).setFontColor('#888888')
          .setBackground(bg).setHorizontalAlignment('right');
        sheet.getRange(r, 5).setBackground(bg);
        sheet.getRange(r, 6).setValue(fmt(aRow.value))
          .setFontSize(9).setFontColor('#1a1a1a')
          .setBackground(bg).setHorizontalAlignment('right');
      }
    } else {
      // fill blank left cells
      sheet.getRange(r, 1, 1, 6).setBackground('#ffffff');
    }

    // Gap col G
    sheet.getRange(r, 7).setBackground('#ffffff');

    // ── Right: liabilities ──────────────────────────────────────────────────
    if (lRow) {
      if (lRow.type === 'section_header') {
        sheet.getRange(r, 8, 1, 3).merge()
          .setValue('LIABILITIES')
          .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
          .setBackground('#5B1A1A').setHorizontalAlignment('left');
        sheet.getRange(r, 11).setValue('Updated')
          .setFontSize(9).setFontColor('#ffffff').setFontWeight('bold')
          .setBackground('#5B1A1A').setHorizontalAlignment('right');
        sheet.getRange(r, 12).setBackground('#5B1A1A');
        sheet.getRange(r, 13).setValue(fmt(lRow.total))
          .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
          .setBackground('#5B1A1A').setHorizontalAlignment('right');

      } else if (lRow.type === 'cat_header') {
        sheet.getRange(r, 8, 1, 3).merge()
          .setValue(lRow.label)
          .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
          .setBackground('#A33030').setHorizontalAlignment('left');
        sheet.getRange(r, 11).setValue(lRow.updated || '')
          .setFontSize(8).setFontColor('#f5cccc')
          .setBackground('#A33030').setHorizontalAlignment('right');
        sheet.getRange(r, 12).setBackground('#A33030');
        sheet.getRange(r, 13).setValue(fmt(lRow.total))
          .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
          .setBackground('#A33030').setHorizontalAlignment('right');

      } else if (lRow.type === 'item') {
        var lbg = (i % 2 === 0) ? '#fdf5f5' : '#ffffff';
        sheet.getRange(r, 8, 1, 3).merge()
          .setValue('  ' + lRow.label)
          .setFontSize(9).setFontColor('#1a1a1a')
          .setBackground(lbg).setHorizontalAlignment('left');
        sheet.getRange(r, 11).setValue(lRow.updated || '')
          .setFontSize(8).setFontColor('#888888')
          .setBackground(lbg).setHorizontalAlignment('right');
        sheet.getRange(r, 12).setBackground(lbg);
        sheet.getRange(r, 13).setValue(fmt(lRow.value))
          .setFontSize(9).setFontColor('#1a1a1a')
          .setBackground(lbg).setHorizontalAlignment('right');
      }
    } else {
      sheet.getRange(r, 8, 1, 6).setBackground('#ffffff');
    }
  }

  // ── Column widths ──────────────────────────────────────────────────────────
  sheet.setColumnWidth(1,  140);  // A - asset name (part 1)
  sheet.setColumnWidth(2,  100);  // B - name (cont.)
  sheet.setColumnWidth(3,   80);  // C - name (cont.)
  sheet.setColumnWidth(4,  100);  // D - days ago
  sheet.setColumnWidth(5,    8);  // E - spacer
  sheet.setColumnWidth(6,  110);  // F - value
  sheet.setColumnWidth(7,   20);  // G - gap
  sheet.setColumnWidth(8,  140);  // H - liab name (part 1)
  sheet.setColumnWidth(9,  100);  // I - name (cont.)
  sheet.setColumnWidth(10,  80);  // J - name (cont.)
  sheet.setColumnWidth(11, 100);  // K - days ago
  sheet.setColumnWidth(12,   8);  // L - spacer
  sheet.setColumnWidth(13, 110);  // M - value

  // ── Row heights ────────────────────────────────────────────────────────────
  sheet.setRowHeight(1, 40);
  sheet.setRowHeight(2, 22);
  sheet.setRowHeight(3, 10);
  sheet.setRowHeight(4, 36);
  sheet.setRowHeight(5, 10);
  for (var ri = 6; ri < 6 + maxRows; ri++) {
    sheet.setRowHeight(ri, 22);
  }

  // ── Hide gridlines & freeze top rows ──────────────────────────────────────
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(5);

  // ── Activate the sheet ────────────────────────────────────────────────────
  ss.setActiveSheet(sheet);
  ss.toast('Balances sheet refreshed!', 'Done', 4);

  return { success: true };
}

// ── Net Worth Snapshot (monthly, assets + liabilities) ────────────────────────
// Stores one row per asset/liability with a month key so we can pivot into the
// Tiller-style Net Worth History sheet.  Safe to call daily — deduplicates by
// month key so only one snapshot per calendar month is ever stored.

// Returns the NW Snapshots sheet, creating it directly if it doesn't exist.
// We bypass getSheet_() / ensureSheets_() entirely because the script cache
// may still hold 'sheets_ready=1' from before this sheet was introduced,
// and getSheet_() returns null when the sheet hasn't been created yet.
function getNWSnapshotsSheet_() {
  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName('NW Snapshots');
  if (!sheet) {
    sheet = ss.insertSheet('NW Snapshots');
    var hdrs = ['Date', 'Month Key', 'Type', 'Name', 'Category', 'USD Value'];
    sheet.getRange(1, 1, 1, hdrs.length).setValues([hdrs])
      .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 220);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 100);
    // Bust the sheet-ready cache so ensureSheets_() re-registers on next run
    CacheService.getScriptCache().remove('sheets_ready');
  }
  return sheet;
}

function takeNWSnapshot(force) {
  var now      = new Date();
  var monthKey = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');

  var snapSheet = getNWSnapshotsSheet_();
  var existing  = snapSheet.getDataRange().getValues();

  // Find and delete existing rows for this month (if force=true, overwrite; otherwise skip)
  var rowsToDelete = [];
  for (var i = existing.length - 1; i >= 1; i--) {
    var mk = existing[i][1];
    if (mk instanceof Date) mk = mk.getFullYear() + '-' + String(mk.getMonth() + 1).padStart(2, '0');
    if (String(mk).trim() === monthKey) {
      if (!force) return { success: false, alreadyDone: true, monthKey: monthKey };
      rowsToDelete.push(i + 1); // 1-based sheet row
    }
  }
  // Delete in reverse order so indices stay valid
  rowsToDelete.sort(function(a,b){return b-a;}).forEach(function(r){ snapSheet.deleteRow(r); });

  var assets = sheetToObjects_('ASSETS');
  var liabs  = sheetToObjects_('LIABILITIES');
  var rows   = [];

  assets.forEach(function(a) {
    rows.push([now, monthKey, 'ASSET', a['Name'] || '', a['Category'] || 'Other', Number(a['My Share USD']) || 0]);
  });
  liabs.forEach(function(l) {
    rows.push([now, monthKey, 'LIABILITY', l['Name'] || '', l['Type'] || 'Other', Number(l['USD Value']) || 0]);
  });

  if (!rows.length) return { success: false, msg: 'No data to snapshot' };

  var startRow = snapSheet.getLastRow() + 1;
  snapSheet.getRange(startRow, 1, rows.length, 6).setValues(rows);
  // Prevent GAS auto-converting "YYYY-MM" to a Date
  snapSheet.getRange(startRow, 2, rows.length, 1).setNumberFormat('@');

  return { success: true, count: rows.length, monthKey: monthKey };
}

// Called from the frontend so the user can trigger both steps in one click.
// Always force-overwrites the current month so clicking the button always refreshes.
function snapshotAndRefreshNWHistory() {
  var snap = takeNWSnapshot(true);  // force=true: delete existing month data and retake
  var gen  = generateNetWorthHistorySheet();
  return { snapshot: snap, sheet: gen };
}

// ── Net Worth History Sheet (Tiller-style pivot) ──────────────────────────────
// Reads NW_SNAPSHOTS and writes a pivot where:
//   Column A  = row labels (NET WORTH, category names, asset names, …)
//   Column B+ = one column per month (oldest → newest, up to 24 months)

function generateNetWorthHistorySheet() {
  var ss = getSpreadsheet_();
  var SHEET_NAME = 'Net Worth History';

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) {
    sheet.clearContents();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // ── Read raw snapshot data ───────────────────────────────────────────────
  var snapSheet = getNWSnapshotsSheet_();
  var raw = snapSheet.getDataRange().getValues();

  if (raw.length < 2) {
    sheet.getRange(1, 1).setValue(
      'No Net Worth snapshot data yet.\n' +
      'Click "Snapshot Net Worth" in the web app Balance History tab, ' +
      'or run Tracker → Take Net Worth Snapshot (1st of month).'
    );
    return { success: false, msg: 'No data' };
  }

  // Parse: Date(0) | Month Key(1) | Type(2) | Name(3) | Category(4) | USD Value(5)
  var byMonth = {}; // mk -> { assets: {name->{cat,val}}, liabs: {name->{type,val}} }

  for (var ri = 1; ri < raw.length; ri++) {
    var row = raw[ri];
    var mk  = row[1];
    if (mk instanceof Date) mk = mk.getFullYear() + '-' + String(mk.getMonth() + 1).padStart(2, '0');
    mk = String(mk).trim();
    if (!mk) continue;

    var recType = String(row[2]).trim();
    var name    = String(row[3]).trim();
    var cat     = String(row[4]).trim() || 'Other';
    var val     = Number(row[5]) || 0;

    if (!byMonth[mk]) byMonth[mk] = { assets: {}, liabs: {} };
    if (recType === 'ASSET')     byMonth[mk].assets[name] = { cat: cat, val: val };
    else if (recType === 'LIABILITY') byMonth[mk].liabs[name] = { type: cat, val: val };
  }

  // Sort months, keep last 24
  var allMonths = Object.keys(byMonth).sort();
  var months    = allMonths.slice(-24);
  if (!months.length) {
    sheet.getRange(1, 1).setValue('No valid snapshot data found.');
    return { success: false };
  }

  var numMonths = months.length;
  var numCols   = 1 + numMonths; // label col + one per month

  // ── Month label helper ───────────────────────────────────────────────────
  var MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  function fmtMK(mk) {
    var p = mk.split('-'); return MN[parseInt(p[1]) - 1] + ' ' + p[0];
  }

  // ── Build row/asset structure from the LATEST month's roster ────────────
  var latest = byMonth[months[months.length - 1]] || { assets: {}, liabs: {} };

  var assetsByCat = {};
  CATEGORIES.forEach(function(c) { assetsByCat[c] = []; });
  Object.keys(latest.assets).forEach(function(n) {
    var c = latest.assets[n].cat || 'Other';
    if (!assetsByCat[c]) assetsByCat[c] = [];
    assetsByCat[c].push(n);
  });
  var usedCats = CATEGORIES.filter(function(c) { return assetsByCat[c] && assetsByCat[c].length; });

  var liabsByType = {};
  Object.keys(latest.liabs).forEach(function(n) {
    var t = latest.liabs[n].type || 'Other';
    if (!liabsByType[t]) liabsByType[t] = [];
    liabsByType[t].push(n);
  });
  var usedLiabTypes = Object.keys(liabsByType).sort();

  // ── Aggregation helpers ──────────────────────────────────────────────────
  function assetTotal(mk) {
    if (!byMonth[mk]) return 0;
    return Object.keys(byMonth[mk].assets).reduce(function(s, n) { return s + byMonth[mk].assets[n].val; }, 0);
  }
  function liabTotal(mk) {
    if (!byMonth[mk]) return 0;
    return Object.keys(byMonth[mk].liabs).reduce(function(s, n) { return s + byMonth[mk].liabs[n].val; }, 0);
  }
  function nwTotal(mk) { return assetTotal(mk) - liabTotal(mk); }

  function catTotalF(mk, cat) {
    if (!byMonth[mk]) return null;
    var t = 0, found = false;
    Object.keys(byMonth[mk].assets).forEach(function(n) {
      if (byMonth[mk].assets[n].cat === cat) { t += byMonth[mk].assets[n].val; found = true; }
    });
    return found ? t : null;
  }
  function assetValF(mk, name) {
    return (byMonth[mk] && byMonth[mk].assets[name] != null) ? byMonth[mk].assets[name].val : null;
  }
  function liabTypeTotalF(mk, type) {
    if (!byMonth[mk]) return null;
    var t = 0, found = false;
    Object.keys(byMonth[mk].liabs).forEach(function(n) {
      if (byMonth[mk].liabs[n].type === type) { t += byMonth[mk].liabs[n].val; found = true; }
    });
    return found ? t : null;
  }
  function liabValF(mk, name) {
    return (byMonth[mk] && byMonth[mk].liabs[name] != null) ? byMonth[mk].liabs[name].val : null;
  }

  // ── Define row list ──────────────────────────────────────────────────────
  // t = type string; lbl = label text; vals = array[numMonths] of numbers|null
  var rowDefs = [];

  rowDefs.push({ t: 'header' });

  rowDefs.push({ t: 'net_worth', lbl: 'NET WORTH',
    vals: months.map(nwTotal) });

  rowDefs.push({ t: 'pct_change', lbl: '% Change',
    vals: months.map(function(mk, i) {
      if (i === 0) return null;
      var prev = nwTotal(months[i - 1]), curr = nwTotal(mk);
      return prev ? (curr - prev) / Math.abs(prev) : null;
    })
  });

  rowDefs.push({ t: 'blank' });

  rowDefs.push({ t: 'asset_hdr', lbl: 'ASSET',
    vals: months.map(assetTotal) });

  usedCats.forEach(function(cat) {
    rowDefs.push({ t: 'cat_hdr', lbl: cat.toUpperCase(),
      vals: months.map(function(mk) { return catTotalF(mk, cat); }) });
    assetsByCat[cat].forEach(function(name) {
      rowDefs.push({ t: 'item', lbl: '  ' + name,
        vals: months.map(function(mk) { return assetValF(mk, name); }) });
    });
  });

  rowDefs.push({ t: 'blank' });

  if (usedLiabTypes.length) {
    rowDefs.push({ t: 'liab_hdr', lbl: 'LIABILITIES',
      vals: months.map(liabTotal) });
    usedLiabTypes.forEach(function(type) {
      rowDefs.push({ t: 'liab_type', lbl: type.toUpperCase(),
        vals: months.map(function(mk) { return liabTypeTotalF(mk, type); }) });
      liabsByType[type].forEach(function(name) {
        rowDefs.push({ t: 'liab_item', lbl: '  ' + name,
          vals: months.map(function(mk) { return liabValF(mk, name); }) });
      });
    });
  }

  var numRows = rowDefs.length;

  // ── Build 2D values array and write in ONE call ───────────────────────────
  var allVals = rowDefs.map(function(row) {
    var arr = [];
    for (var c = 0; c < numCols; c++) arr.push('');
    if (row.t === 'header') {
      months.forEach(function(mk, ci) { arr[1 + ci] = fmtMK(mk); });
    } else if (row.t !== 'blank') {
      arr[0] = row.lbl || '';
      (row.vals || []).forEach(function(v, ci) {
        arr[1 + ci] = (v !== null && v !== undefined) ? v : '';
      });
    }
    return arr;
  });

  // ── Chart area offset ────────────────────────────────────────────────────
  // The top CHART_OFFSET rows are reserved for the line chart; data starts below.
  var CHART_OFFSET = 17; // rows of blank space above the data table for the chart
  var totalRows    = CHART_OFFSET + numRows;

  // Resize sheet to accommodate chart rows + data rows
  if (sheet.getMaxColumns() < numCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), numCols - sheet.getMaxColumns());
  if (sheet.getMaxRows() < totalRows)  sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());

  // Write data starting at row CHART_OFFSET + 1
  sheet.getRange(CHART_OFFSET + 1, 1, numRows, numCols).setValues(allVals);

  // ── Apply formatting (rows shifted by CHART_OFFSET) ───────────────────────
  var STYLE = {
    header:     { bg: '#1B3A5C', fg: '#ffffff', bold: true,  sz: 9  },
    net_worth:  { bg: '#1A7341', fg: '#ffffff', bold: true,  sz: 11 },
    pct_change: { bg: '#f2f6fa', fg: '#555555', bold: false, sz: 9  },
    blank:      { bg: '#ffffff', fg: '#ffffff', bold: false, sz: 9  },
    asset_hdr:  { bg: '#1B3A5C', fg: '#ffffff', bold: true,  sz: 10 },
    cat_hdr:    { bg: '#2E6DA4', fg: '#ffffff', bold: true,  sz: 9  },
    item:       { bg: null,      fg: '#1a1a1a', bold: false, sz: 9  },
    liab_hdr:   { bg: '#5B1A1A', fg: '#ffffff', bold: true,  sz: 10 },
    liab_type:  { bg: '#A33030', fg: '#ffffff', bold: true,  sz: 9  },
    liab_item:  { bg: null,      fg: '#c5221f', bold: false, sz: 9  }
  };
  var itemBgIdx = 0, liabItemBgIdx = 0;
  var numFmt    = '$#,##0';

  rowDefs.forEach(function(row, ri) {
    var r  = ri + 1 + CHART_OFFSET; // ← shifted down by chart offset
    var st = STYLE[row.t] || STYLE.blank;
    var bg = st.bg;

    if (row.t === 'item')      bg = (itemBgIdx++    % 2 === 0) ? '#f5f8fc' : '#ffffff';
    if (row.t === 'liab_item') bg = (liabItemBgIdx++ % 2 === 0) ? '#fcf5f5' : '#ffffff';

    var rng = sheet.getRange(r, 1, 1, numCols);
    rng.setBackground(bg).setFontColor(st.fg).setFontWeight(st.bold ? 'bold' : 'normal')
       .setFontSize(st.sz).setVerticalAlignment('middle');

    sheet.getRange(r, 1).setHorizontalAlignment('left');
    if (numMonths > 0) sheet.getRange(r, 2, 1, numMonths).setHorizontalAlignment('right');

    if (row.t !== 'header' && row.t !== 'blank' && row.t !== 'pct_change' && numMonths > 0) {
      sheet.getRange(r, 2, 1, numMonths).setNumberFormat(numFmt);
    }

    if (row.t === 'pct_change' && row.vals) {
      row.vals.forEach(function(v, ci) {
        var cell = sheet.getRange(r, 2 + ci);
        if (v === null || v === undefined || v === '') { cell.setBackground(st.bg); return; }
        cell.setNumberFormat('0.0%').setFontWeight('bold')
            .setFontColor(v < 0 ? '#c5221f' : '#1a7341');
      });
    }
  });

  // ── Title row (row 1) ────────────────────────────────────────────────────
  var currentNW = nwTotal(months[months.length - 1]);
  sheet.getRange(1, 1).setValue('Net Worth Over Time')
    .setFontSize(14).setFontWeight('bold').setFontColor('#0d2137')
    .setBackground('#ffffff').setVerticalAlignment('middle');
  sheet.getRange(1, numCols).setValue(currentNW)
    .setFontSize(14).setFontWeight('bold').setFontColor('#0d2137')
    .setHorizontalAlignment('right').setNumberFormat('$#,##0')
    .setBackground('#ffffff').setVerticalAlignment('middle');
  // Fill chart area with white
  sheet.getRange(1, 1, CHART_OFFSET, numCols).setBackground('#ffffff');

  // ── Row heights ───────────────────────────────────────────────────────────
  sheet.setRowHeight(1, 30); // title
  for (var ci = 2; ci <= CHART_OFFSET; ci++) sheet.setRowHeight(ci, 18); // chart area rows
  sheet.setRowHeight(CHART_OFFSET + 1, 26); // month header
  sheet.setRowHeight(CHART_OFFSET + 2, 32); // NET WORTH
  sheet.setRowHeight(CHART_OFFSET + 3, 22); // % Change
  for (var ri2 = CHART_OFFSET + 4; ri2 <= totalRows; ri2++) sheet.setRowHeight(ri2, 20);

  // ── Column widths ─────────────────────────────────────────────────────────
  sheet.setColumnWidth(1, 235);
  for (var ci2 = 0; ci2 < numMonths; ci2++) sheet.setColumnWidth(2 + ci2, 105);

  // ── Freeze: lock the month header row + label column ────────────────────
  sheet.setFrozenRows(CHART_OFFSET + 1);
  sheet.setFrozenColumns(1);
  sheet.setHiddenGridlines(true);

  // ── Line chart ────────────────────────────────────────────────────────────
  // Write chart series data in off-screen columns (to the right of the visible data)
  // so the chart has contiguous data to read from.
  var chartDataCol = numCols + 3;
  var chartSeries  = [
    [''].concat(months.map(fmtMK)),           // row 0: x-axis labels
    ['Net Worth'].concat(months.map(nwTotal)), // row 1
    ['Assets'].concat(months.map(assetTotal)), // row 2
    ['Liabilities'].concat(months.map(liabTotal)) // row 3
  ];
  sheet.getRange(1, chartDataCol, 4, numMonths + 1).setValues(chartSeries);

  // Remove any existing chart before inserting a new one
  sheet.getCharts().forEach(function(c) { sheet.removeChart(c); });

  var chartWidthPx  = Math.min(235 + numMonths * 105, 1100);
  var chartHeightPx = (CHART_OFFSET - 1) * 18; // match reserved rows

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, chartDataCol, 4, numMonths + 1))
    .setTransposeRowsAndColumns(true) // rows become series, first row = x-axis labels
    .setPosition(2, 1, 0, 0)         // anchor top-left at row 2, col 1
    .setOption('title', '')
    .setOption('legend', { position: 'right' })
    .setOption('series', {
      0: { color: '#1e8e3e', lineWidth: 2 },  // Net Worth — green
      1: { color: '#1a73e8', lineWidth: 2 },  // Assets — blue
      2: { color: '#c5221f', lineWidth: 2 }   // Liabilities — red
    })
    .setOption('vAxis', {
      format: '$#,##0,,"M"',
      textStyle: { fontSize: 9, color: '#555555' },
      gridlines: { color: '#e8eef4' }
    })
    .setOption('hAxis', {
      textStyle: { fontSize: 9, color: '#555555' },
      gridlines: { color: 'transparent' }
    })
    .setOption('backgroundColor', { fill: '#f8fafd' })
    .setOption('chartArea', { left: 75, top: 15, right: 130, bottom: 35 })
    .setOption('height', chartHeightPx)
    .setOption('width', chartWidthPx)
    .build();

  sheet.insertChart(chart);

  ss.toast('Net Worth History refreshed — ' + numMonths + ' months shown', 'Done', 4);
  return { success: true, months: numMonths, rows: numRows };
}

// ── Database Setup ────────────────────────────────────────────────────────────

function setupDatabase() {
  var ss = getSpreadsheet_();
  CacheService.getScriptCache().remove('sheets_ready'); // force re-check after setup
  _sheetsReady = false;
  ensureSheets_();

  // ── Tab colors ──────────────────────────────────────────────────────────
  var tabColors = {
    'Assets':            '#1B3A5C',
    'Liabilities':       '#7B2D2D',
    'Entities':          '#4A235A',
    'FX Rates':          '#145A32',
    'History':           '#5D6D7E',
    'Snapshots':         '#424949',
    'Asset Details':     '#1F618D',
    'Liability Details': '#922B21',
    'Org Chart':         '#6C3483',
    'Balances':          '#1A7341',
    'NW Snapshots':      '#2C3E50',
    'Net Worth History': '#1A5276'
  };
  Object.keys(tabColors).forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) s.setTabColor(tabColors[name]);
  });

  // ── Assets sheet ────────────────────────────────────────────────────────
  // Columns: ID(1) Name(2) Category(3) Entity(4) Currency(5) Local Value(6)
  //   USD Rate(7) USD Value(8) My Share %(9) My Share USD(10) Date Added(11)
  //   Last Updated(12) Notes(13) Plaid Account ID(14) Address(15) Cost Basis(16) Details(17)
  var assets = ss.getSheetByName('Assets');
  if (assets) {
    [[1,30],[2,220],[3,170],[4,160],[5,70],[6,120],[7,80],[8,120],
     [9,90],[10,130],[11,100],[12,110],[13,200],[15,180],[16,120]]
      .forEach(function(w){ assets.setColumnWidth(w[0], w[1]); });
    assets.hideColumns(14);  // Plaid Account ID — internal
    assets.hideColumns(17);  // Details JSON blob — internal

    var N = 2000;
    // Category dropdown
    assets.getRange(2, 3, N).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(CATEGORIES, true)
        .setAllowInvalid(false).setHelpText('Select a category').build());
    // Currency dropdown
    assets.getRange(2, 5, N).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(CURRENCIES, true)
        .setAllowInvalid(false).build());
    // Number formats
    assets.getRange(2,  6, N).setNumberFormat('#,##0.00');    // Local Value
    assets.getRange(2,  7, N).setNumberFormat('0.000000');    // USD Rate
    assets.getRange(2,  8, N).setNumberFormat('$#,##0.00');   // USD Value
    assets.getRange(2,  9, N).setNumberFormat('0.00');        // My Share %
    assets.getRange(2, 10, N).setNumberFormat('$#,##0.00');   // My Share USD
    assets.getRange(2, 16, N).setNumberFormat('$#,##0.00');   // Cost Basis
    assets.setFrozenRows(1);
    protectHeader_(assets);
  }

  // ── Liabilities sheet ───────────────────────────────────────────────────
  // Columns: ID(1) Name(2) Type(3) Currency(4) Amount(5) USD Value(6)
  //   Date Added(7) Last Updated(8) Notes(9) Location(10) Details(11)
  var liabs = ss.getSheetByName('Liabilities');
  if (liabs) {
    [[1,30],[2,220],[3,160],[4,70],[5,120],[6,120],[7,100],[8,110],[9,220],[10,160]]
      .forEach(function(w){ liabs.setColumnWidth(w[0], w[1]); });
    liabs.hideColumns(11); // Details JSON
    var liabTypes = ['Mortgage','Auto Loan','Personal Loan','Credit Card',
                     'Line of Credit','Business Loan','Student Loan','Other'];
    liabs.getRange(2, 3, 1000).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(liabTypes, true)
        .setAllowInvalid(true).build());
    liabs.getRange(2, 4, 1000).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(CURRENCIES, true)
        .setAllowInvalid(false).build());
    liabs.getRange(2, 5, 1000).setNumberFormat('#,##0.00');
    liabs.getRange(2, 6, 1000).setNumberFormat('$#,##0.00');
    liabs.setFrozenRows(1);
    protectHeader_(liabs);
  }

  // ── Entities sheet ──────────────────────────────────────────────────────
  var ents = ss.getSheetByName('Entities');
  if (ents) {
    [[1,200],[2,140],[3,140],[4,90],[5,260]]
      .forEach(function(w){ ents.setColumnWidth(w[0], w[1]); });
    ents.getRange(2, 4, 500).setNumberFormat('0.00');
    ents.setFrozenRows(1);
    protectHeader_(ents);
  }

  // ── FX Rates sheet ──────────────────────────────────────────────────────
  var fx = ss.getSheetByName('FX Rates');
  if (fx) {
    fx.setColumnWidth(1, 100); fx.setColumnWidth(2, 130); fx.setColumnWidth(3, 160);
    fx.getRange(2, 2, 100).setNumberFormat('0.000000');
    fx.setFrozenRows(1);
    protectHeader_(fx);
  }

  // ── History sheet ───────────────────────────────────────────────────────
  var hist = ss.getSheetByName('History');
  if (hist) {
    [[1,110],[2,220],[3,120],[4,120],[5,120],[6,80],[7,220]]
      .forEach(function(w){ hist.setColumnWidth(w[0], w[1]); });
    hist.getRange(2, 3, 5000, 3).setNumberFormat('$#,##0.00');
    hist.setFrozenRows(1);
    protectHeader_(hist);
  }

  // ── Snapshots sheet ─────────────────────────────────────────────────────
  var snap = ss.getSheetByName('Snapshots');
  if (snap) {
    snap.getRange(2, 6, 10000).setNumberFormat('$#,##0.00'); // My Share USD
    snap.setFrozenRows(1);
    protectHeader_(snap);
  }

  // ── Asset Details & Liability Details ────────────────────────────────────
  ['Asset Details', 'Liability Details'].forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) { s.setFrozenRows(1); protectHeader_(s); }
  });

  SpreadsheetApp.flush();
  ss.toast('Database structure configured! Tab colors, validation, formatting, and header protection applied.', 'Setup Complete', 8);
  return { success: true };
}

// ── Schema Reset ──────────────────────────────────────────────────────────────

function resetSchema() {
  var ss        = getSpreadsheet_();
  var ui        = SpreadsheetApp.getUi();
  var detSheet  = ss.getSheetByName('Asset Details');
  if (!detSheet) { ui.alert('Asset Details sheet not found. Run Setup Database Structure first.'); return; }

  var existingHeaders = detSheet.getLastColumn() > 0
    ? detSheet.getRange(1, 1, 1, detSheet.getLastColumn()).getValues()[0]
    : [];
  var dataRows = detSheet.getLastRow() - 1;
  var newHeaders = COL.ASSET_DETAILS;

  if (dataRows > 0) {
    // ── Has existing data: preserve all columns, mark removed ones [OLD] ──
    var response = ui.alert(
      'Asset Details has ' + dataRows + ' data row(s).',
      'Existing columns that are no longer in the schema will be renamed [OLD]. New columns will be added to the right. No data will be deleted. Continue?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;

    // Mark columns no longer in schema as [OLD]
    existingHeaders.forEach(function(h, i) {
      if (h && newHeaders.indexOf(h) === -1 && String(h).indexOf('[OLD]') === -1) {
        detSheet.getRange(1, i + 1).setValue('[OLD] ' + h);
      }
    });
    // Add missing new columns to the right
    var updatedHeaders = detSheet.getRange(1, 1, 1, detSheet.getLastColumn()).getValues()[0];
    newHeaders.forEach(function(h) {
      if (updatedHeaders.indexOf(h) === -1) {
        var nextCol = detSheet.getLastColumn() + 1;
        detSheet.getRange(1, nextCol).setValue(h)
          .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
      }
    });
    ss.toast('Schema updated. Old columns marked [OLD], new columns added. No data was deleted.', 'Schema Updated', 8);

  } else {
    // ── No data: safe to completely rebuild headers ─────────────────────
    var response2 = ui.alert(
      'Rebuild Asset Details schema?',
      'The sheet has no data. This will clear and replace all headers with the current schema.',
      ui.ButtonSet.YES_NO
    );
    if (response2 !== ui.Button.YES) return;

    detSheet.clearContents();
    detSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders])
      .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
    detSheet.setFrozenRows(1);

    // Style: group columns by category with slightly different header shades
    var groupColors = {
      'Asset ID': '#0d2137', 'Asset Name': '#0d2137', 'Status': '#0d2137',
      'Updated By': '#0d2137', 'Description': '#0d2137', 'Drive Folder': '#0d2137',
      'Project Leader': '#0d2137',
      'Occupancy': '#1B3A5C', 'Location': '#1B3A5C', 'Type': '#1B3A5C', 'Sqft': '#1B3A5C',
      'Purchase Price': '#1B3A5C', 'Purchase Date': '#1B3A5C', 'Closing Costs': '#1B3A5C',
      'Permits': '#1B3A5C', 'Revenue': '#1B3A5C', 'OpEx': '#1B3A5C',
      'Property Tax': '#1B3A5C', 'Insurance': '#1B3A5C', 'HOA': '#1B3A5C',
      'Maintenance': '#1B3A5C', 'Utilities': '#1B3A5C', 'Loan Info': '#1B3A5C',
      'Financial Notes': '#1B3A5C',
      'Contact 1 Type': '#2E4057', 'Contact 1 Name': '#2E4057',
      'Contact 2 Type': '#2E4057', 'Contact 2 Name': '#2E4057',
      'Contact 3 Type': '#2E4057', 'Contact 3 Name': '#2E4057',
      'Contact 4 Type': '#2E4057', 'Contact 4 Name': '#2E4057',
      'Borrower Name': '#5B1A1A', 'Borrower Contact': '#5B1A1A',
      'Original Amount': '#5B1A1A', 'Outstanding Balance': '#5B1A1A',
      'Interest Rate': '#5B1A1A', 'Loan Status': '#5B1A1A', 'Loan Date': '#5B1A1A',
      'Due Date': '#5B1A1A', 'Loan Terms': '#5B1A1A', 'Payment Schedule': '#5B1A1A',
      'Received To Date': '#5B1A1A', 'Collateral': '#5B1A1A', 'Drive Link': '#5B1A1A',
      'Attorney': '#5B1A1A', 'Loan Notes': '#5B1A1A',
      'Cash Bank': '#145A32', 'Cash Account Type': '#145A32',
      'Cash Account Number': '#145A32', 'Cash Interest Rate': '#145A32',
      'Custodian / Manager': '#1F618D', 'Equity Account Number': '#1F618D',
      'Shares / Units': '#1F618D', 'Avg Cost Per Share': '#1F618D', 'Equity Notes': '#1F618D',
      'PE Manager': '#4A235A', 'PE Tax Treatment': '#4A235A',
      'PE Year Invested': '#4A235A', 'PE Target Exit Year': '#4A235A',
      'PE Year Sold': '#4A235A', 'PE Year Written Off': '#4A235A',
      'PE Initial Investment': '#4A235A', 'PE Ownership %': '#4A235A',
      'PE Maturity Date': '#4A235A', 'PE Return Rate': '#4A235A',
      'PE Capital Calls': '#4A235A', 'PE Distributions': '#4A235A', 'PE Notes': '#4A235A'
    };
    newHeaders.forEach(function(h, i) {
      var col = groupColors[h] || '#0d2137';
      detSheet.getRange(1, i + 1).setBackground(col);
    });

    protectHeader_(detSheet);
    ss.toast('Asset Details schema rebuilt clean.', 'Done', 5);
  }
}

function protectHeader_(sheet) {
  // Remove any existing header protections first
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) {
    try {
      var r = p.getRange();
      if (r.getRow() === 1 && r.getNumRows() === 1) p.remove();
    } catch(e) {}
  });
  var protection = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).protect();
  protection.setDescription('Header row — managed by system');
  protection.setWarningOnly(true); // warns before editing but doesn't block
}
