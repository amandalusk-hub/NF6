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
  ENTITIES:    ['Name','Type','Jurisdiction','Ownership %','Notes'],
  FX:          ['Currency','Rate to USD','Last Fetched'],
  HISTORY:     ['Date','Asset Name','Old Value USD','New Value USD','Delta USD','Currency','Notes'],
  SNAPSHOTS:   ['Date','Month Key','Asset Name','Category','Currency','My Share USD'],
  ASSET_DETAILS:     ['Asset ID','Asset Name','Updated By','Project Leader','Occupancy','Description','Location','Type','Sqft','Drive Folder','Purchase Price','Purchase Date','Closing Costs','Permits','Revenue','OpEx','Property Tax','Insurance','HOA','Maintenance','Utilities','Loan Info','Financial Notes','Contact 1 Type','Contact 1 Name','Contact 2 Type','Contact 2 Name','Contact 3 Type','Contact 3 Name','Contact 4 Type','Contact 4 Name','Borrower Name','Borrower Contact','Original Amount','Outstanding Balance','Interest Rate','Loan Status','Loan Date','Due Date','Loan Terms','Payment Schedule','Received To Date','Collateral','Drive Link','Attorney','Loan Notes'],
  LIABILITY_DETAILS: ['Liability ID','Liability Name','Bank / Lender','Account Number','Interest Rate','Loan Type','Original Amount','Current Balance','Start Date','Maturity Date','Loan Term','Months Remaining','Monthly Payment','Principal','Interest Payment','Escrow','Property Tax','Insurance','HOA','Loan Officer','Attorney / Title','Insurance Agent','Other Contacts','Notes']
};

// Maps JS field names ↔ Asset Details sheet column names
var ASSET_DET_MAP = [
  ['updatedBy','Updated By'],['projectLeader','Project Leader'],['occupancy','Occupancy'],
  ['description','Description'],['location','Location'],['type','Type'],['sqft','Sqft'],
  ['folder','Drive Folder'],['purchasePrice','Purchase Price'],['purchaseDate','Purchase Date'],
  ['closingCosts','Closing Costs'],['permits','Permits'],['revenue','Revenue'],['opex','OpEx'],
  ['taxes','Property Tax'],['insurance','Insurance'],['hoa','HOA'],['maintenance','Maintenance'],
  ['utilities','Utilities'],['loan','Loan Info'],['finNotes','Financial Notes'],
  ['borrower','Borrower Name'],['borrowerContact','Borrower Contact'],
  ['loanOriginal','Original Amount'],['loanBalance','Outstanding Balance'],
  ['loanRate','Interest Rate'],['loanStatus','Loan Status'],['loanDate','Loan Date'],
  ['loanDue','Due Date'],['loanTerms','Loan Terms'],['loanPayment','Payment Schedule'],
  ['loanReceived','Received To Date'],['loanCollateral','Collateral'],
  ['loanDrive','Drive Link'],['loanAttorney','Attorney'],['loanNotes','Loan Notes']
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
  var ss = getSpreadsheet_();
  Object.keys(COL).forEach(function(key) {
    var name    = sheetName_(key);
    var headers = COL[key];
    var sheet   = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 220);
    } else {
      // Add any columns that exist in COL but are missing from the sheet
      var lastCol = sheet.getLastColumn();
      var existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
      var missing = headers.filter(function(h) { return existing.indexOf(h) === -1; });
      if (missing.length) {
        var startCol = lastCol + 1;
        sheet.getRange(1, startCol, 1, missing.length).setValues([missing])
          .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
      }
    }
  });
  _sheetsReady = true;
}

function sheetName_(key) {
  return { ASSETS: 'Assets', LIABILITIES: 'Liabilities', ENTITIES: 'Entities', FX: 'FX Rates', HISTORY: 'History', SNAPSHOTS: 'Snapshots', ASSET_DETAILS: 'Asset Details', LIABILITY_DETAILS: 'Liability Details' }[key];
}

function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

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

  // Auto-assign IDs to any asset rows that were manually entered without one
  var assetSheet = getSheet_('ASSETS');
  var assetRows  = assetSheet.getDataRange().getValues();
  for (var i = 1; i < assetRows.length; i++) {
    if (!assetRows[i][0]) {
      assetSheet.getRange(i + 1, 1).setValue(Utilities.getUuid());
    }
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

  // Build Asset Details map (id → det object with JS field names)
  var assetDetMap = {};
  var adSheet = getSheet_('ASSET_DETAILS');
  var adData  = adSheet.getDataRange().getValues();
  if (adData.length > 1) {
    var adHeaders = adData[0];
    for (var ai = 1; ai < adData.length; ai++) {
      var adRow = adData[ai];
      var detId = String(adRow[0]);
      if (!detId) continue;
      var colObj = {};
      adHeaders.forEach(function(h, j) { colObj[h] = adRow[j]; });
      var det = {};
      ASSET_DET_MAP.forEach(function(m) { det[m[0]] = colObj[m[1]] || ''; });
      det.contacts = [];
      for (var ci = 1; ci <= 4; ci++) {
        var cType = colObj['Contact ' + ci + ' Type'] || '';
        var cName = colObj['Contact ' + ci + ' Name'] || '';
        if (cType || cName) det.contacts.push({ type: cType, name: cName });
      }
      assetDetMap[detId] = det;
    }
  }

  // Build Liability Details map
  var liabDetMap = {};
  var ldSheet = getSheet_('LIABILITY_DETAILS');
  var ldData  = ldSheet.getDataRange().getValues();
  if (ldData.length > 1) {
    var ldHeaders = ldData[0];
    for (var li = 1; li < ldData.length; li++) {
      var ldRow = ldData[li];
      var ldId  = String(ldRow[0]);
      if (!ldId) continue;
      var ldColObj = {};
      ldHeaders.forEach(function(h, j) { ldColObj[h] = ldRow[j]; });
      var ldet = {};
      LIAB_DET_MAP.forEach(function(m) { ldet[m[0]] = ldColObj[m[1]] || ''; });
      liabDetMap[ldId] = ldet;
    }
  }

  // Merge det into each asset and liability
  var assets = clean(sheetToObjects_('ASSETS')).map(function(a) {
    a.det = assetDetMap[String(a.ID)] || null;
    return a;
  });
  var liabilities = clean(sheetToObjects_('LIABILITIES')).map(function(l) {
    l.det = liabDetMap[String(l.ID)] || null;
    return l;
  });

  var assetHeaderRow = assetSheet.getRange(1, 1, 1, Math.max(assetSheet.getLastColumn(), 1)).getValues()[0];
  return {
    assets:      assets,
    liabilities: liabilities,
    entities:    clean(sheetToObjects_('ENTITIES')),
    fxRates:     clean(sheetToObjects_('FX')),
    history:     clean(sheetToObjects_('HISTORY')),
    snapshots:   getSnapshotTrend(),
    categories:  CATEGORIES,
    currencies:  CURRENCIES,
    _debug: {
      assetHeaders:  assetHeaderRow,
      assetRowCount: Math.max(assetSheet.getLastRow() - 1, 0)
    }
  };
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
    updates.forEach(function(u) { sheet.getRange(i + 1, u[0]).setValue(u[1]); });

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
  var detailsCol   = assetsData[0].indexOf('Details') + 1;
  var lastUpdCol   = assetsData[0].indexOf('Last Updated') + 1;
  var assetName    = '';
  for (var i = 1; i < assetsData.length; i++) {
    if (String(assetsData[i][0]) === String(id)) {
      assetName = String(assetsData[i][1] || '');
      if (detailsCol > 0) assetsSheet.getRange(i + 1, detailsCol).setValue(detailsJson || '');
      if (lastUpdCol > 0) assetsSheet.getRange(i + 1, lastUpdCol).setValue(new Date());
      break;
    }
  }

  // 2. Upsert row in Asset Details sheet (flat columns, spreadsheet-editable)
  var detSheet   = getSheet_('ASSET_DETAILS');
  var detData    = detSheet.getDataRange().getValues();
  var detHeaders = detData[0];
  var contacts   = det.contacts || [];

  // Build a column-name → value lookup
  var colLookup = { 'Asset ID': id, 'Asset Name': assetName };
  ASSET_DET_MAP.forEach(function(m) { colLookup[m[1]] = det[m[0]] || ''; });
  for (var ci = 1; ci <= 4; ci++) {
    var c = contacts[ci - 1] || {};
    colLookup['Contact ' + ci + ' Type'] = c.type || '';
    colLookup['Contact ' + ci + ' Name'] = c.name || '';
  }
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

  return { success: true };
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
  getSheet_('ENTITIES').appendRow([
    data.name || '', data.type || '', data.jurisdiction || '',
    data.ownershipPct !== undefined ? Number(data.ownershipPct) : 100,
    data.notes || ''
  ]);
  return { success: true };
}

function updateEntity(data) {
  var sheet = getSheet_('ENTITIES');
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] !== data.originalName) continue;
    sheet.getRange(i + 1, 1).setValue(data.name         || rows[i][0]);
    sheet.getRange(i + 1, 2).setValue(data.type         || rows[i][1]);
    sheet.getRange(i + 1, 3).setValue(data.jurisdiction || rows[i][2]);
    sheet.getRange(i + 1, 4).setValue(data.ownershipPct !== undefined ? Number(data.ownershipPct) : rows[i][3]);
    sheet.getRange(i + 1, 5).setValue(data.notes        !== undefined ? data.notes : rows[i][4]);
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
  var sheet = getSheet_('LIABILITIES');
  var rows  = sheet.getDataRange().getValues();
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
    sheet.getRange(i + 1, 10).setValue(data.location !== undefined ? data.location : (rows[i][9]  || ''));
    sheet.getRange(i + 1, 11).setValue(data.details  !== undefined ? data.details  : (rows[i][10] || ''));
    return { success: true };
  }
  return { success: false, error: 'Liability not found' };
}

function saveLiabilityDetails(id, detailsJson) {
  var det = {};
  try { det = JSON.parse(detailsJson || '{}'); } catch(e) {}

  // 1. Write JSON blob to LIABILITIES.Details column
  var liabSheet  = getSheet_('LIABILITIES');
  var liabData   = liabSheet.getDataRange().getValues();
  var detailsCol = liabData[0].indexOf('Details') + 1;
  var lastUpdCol = liabData[0].indexOf('Last Updated') + 1;
  var liabName   = '';
  for (var i = 1; i < liabData.length; i++) {
    if (String(liabData[i][0]) === String(id)) {
      liabName = String(liabData[i][1] || '');
      if (detailsCol > 0) liabSheet.getRange(i + 1, detailsCol).setValue(detailsJson || '');
      if (lastUpdCol > 0) liabSheet.getRange(i + 1, lastUpdCol).setValue(new Date());
      break;
    }
  }

  // 2. Upsert row in Liability Details sheet (flat columns)
  var detSheet   = getSheet_('LIABILITY_DETAILS');
  var detData    = detSheet.getDataRange().getValues();
  var detHeaders = detData[0];
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

  return { success: true };
}

function deleteLiability(id) {
  var sheet = getSheet_('LIABILITIES');
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, error: 'Not found' };
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
  snapSheet.getRange(snapSheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  return { success: true, alreadyDone: false, count: rows.length, monthKey: monthKey };
}

function getSnapshotTrend() {
  var data = sheetToObjects_('SNAPSHOTS');
  var byMonth = {};
  data.forEach(function(row) {
    var mk = row['Month Key'] || '';
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
  var snapRows = sheetToObjects_('SNAPSHOTS');
  var monthSet  = {};
  var assetData = {}; // name -> { category, monthKey -> value }

  snapRows.forEach(function(row) {
    var mk   = row['Month Key'] || '';
    var name = row['Asset Name'] || '';
    if (!mk || !name) return;
    monthSet[mk] = true;
    if (!assetData[name]) assetData[name] = { category: row['Category'] || 'Other' };
    assetData[name][mk] = Number(row['My Share USD']) || 0;
  });

  var months = Object.keys(monthSet).sort().slice(-12);
  if (!months.length) return { months: [], assetTotals: [], liabTotals: [], netWorthTotals: [], assets: [] };

  var assetNames = Object.keys(assetData);
  var assets = assetNames.map(function(name) {
    return {
      name:     name,
      category: assetData[name].category,
      values:   months.map(function(m) { return assetData[name][m] || 0; })
    };
  });
  assets.sort(function(a, b) {
    if (a.category !== b.category) return a.category.localeCompare(b.category);
    return (b.values[b.values.length - 1] || 0) - (a.values[a.values.length - 1] || 0);
  });

  var assetTotals = months.map(function(m) {
    return assetNames.reduce(function(s, n) { return s + (assetData[n][m] || 0); }, 0);
  });

  // Use current liabilities total for all months (history accumulates over time)
  var liabs     = sheetToObjects_('LIABILITIES');
  var liabTotal = liabs.reduce(function(s, l) { return s + (Number(l['USD Value']) || 0); }, 0);
  var liabTotals      = months.map(function() { return liabTotal; });
  var netWorthTotals  = assetTotals.map(function(a, i) { return a - liabTotals[i]; });

  return { months: months, assetTotals: assetTotals, liabTotals: liabTotals, netWorthTotals: netWorthTotals, assets: assets };
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
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function syncPlaidAccounts() {
  var cfg    = getPlaidConfig_();
  var p      = PropertiesService.getScriptProperties();
  var tokens = JSON.parse(p.getProperty('PLAID_TOKENS') || '[]');
  if (!tokens.length) return { success: false, error: 'No Plaid accounts connected. Use Connect Bank first.' };

  var synced = 0;
  tokens.forEach(function(token) {
    try {
      var resp = UrlFetchApp.fetch(getPlaidBaseUrl_(cfg.env) + '/accounts/balance/get', {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({ client_id: cfg.clientId, secret: cfg.secret, access_token: token }),
        muteHttpExceptions: true
      });
      var data = JSON.parse(resp.getContentText());
      if (!data.accounts) return;

      data.accounts.forEach(function(acct) {
        var balance  = (acct.balances.current != null ? acct.balances.current : acct.balances.available) || 0;
        var acctName = (acct.name || 'Account') + ' ···' + (acct.mask || '');
        var acctId   = acct.account_id;
        var sheet    = getSheet_('ASSETS');
        var rows     = sheet.getDataRange().getValues();
        var found    = false;

        var matchRow = -1;
        // First pass: exact match by Plaid Account ID or name+category
        for (var i = 1; i < rows.length; i++) {
          if (rows[i][13] === acctId || (rows[i][1] === acctName && rows[i][2] === 'Cash')) {
            matchRow = i;
            break;
          }
        }
        // Second pass: claim any unlinked Cash asset (no Plaid ID set)
        if (matchRow === -1) {
          for (var i = 1; i < rows.length; i++) {
            if (rows[i][2] === 'Cash' && !rows[i][13]) {
              matchRow = i;
              break;
            }
          }
        }

        if (matchRow !== -1) {
          var oldUsd = Number(rows[matchRow][7]) || 0;
          sheet.getRange(matchRow + 1, 2).setValue(acctName);   // update Name from Plaid
          sheet.getRange(matchRow + 1, 6).setValue(balance);
          sheet.getRange(matchRow + 1, 7).setValue(1);
          sheet.getRange(matchRow + 1, 8).setValue(balance);
          sheet.getRange(matchRow + 1, 10).setValue(balance);
          sheet.getRange(matchRow + 1, 12).setValue(new Date());
          sheet.getRange(matchRow + 1, 14).setValue(acctId);
          if (Math.abs(balance - oldUsd) > 0.01) logHistory_(acctName, oldUsd, balance, 'USD', 'Plaid sync');
          found = true;
        }

        if (!found) {
          addAsset({ name: acctName, category: 'Cash', currency: 'USD', localValue: balance, mySharePct: 100, notes: 'Plaid: ' + acctId });
          var newRows = sheet.getDataRange().getValues();
          sheet.getRange(newRows.length, 14).setValue(acctId);
        }
        synced++;
      });
    } catch(e) {
      console.error('Plaid sync error:', e);
    }
  });

  return { success: true, synced: synced };
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

function handlePlaidSuccess(publicToken) {
  try {
    var exchResult = exchangePlaidToken(publicToken);
    if (!exchResult.success) return { success: false, message: exchResult.error };
    var syncResult = syncPlaidAccounts();
    return { success: true, message: 'Bank connected! ' + (syncResult.synced || 0) + ' account(s) synced to Assets tab.' };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
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
  if (new Date().getDate() === 1) takeMonthlySnapshot();
}
