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
  'Real Estate', 'Private Equity', 'Public Equity',
  'Cash', 'Crypto', 'Auto', 'Art/Jewelry',
  'VIP Medical', 'Insurance', 'Other'
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
  ASSETS:      ['ID','Name','Category','Entity','Currency','Local Value','USD Rate','USD Value','My Share %','My Share USD','Date Added','Last Updated','Notes','Plaid Account ID'],
  LIABILITIES: ['ID','Name','Type','Currency','Amount','USD Value','Date Added','Last Updated','Notes'],
  ENTITIES:    ['Name','Type','Jurisdiction','Ownership %','Notes'],
  FX:          ['Currency','Rate to USD','Last Fetched'],
  HISTORY:     ['Date','Asset Name','Old Value USD','New Value USD','Delta USD','Currency','Notes']
};

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
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Family Office — Wealth Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Sheet Bootstrapping ───────────────────────────────────────────────────────

function ensureSheets_() {
  var ss = getSpreadsheet_();
  Object.keys(COL).forEach(function(key) {
    var name  = sheetName_(key);
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      var headers = COL[key];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#0d2137')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 220);
    }
  });
}

function sheetName_(key) {
  return { ASSETS: 'Assets', LIABILITIES: 'Liabilities', ENTITIES: 'Entities', FX: 'FX Rates', HISTORY: 'History' }[key];
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

  function clean(arr) {
    return arr.map(function(obj) {
      var out = {};
      Object.keys(obj).forEach(function(k) {
        out[k] = obj[k] instanceof Date ? obj[k].toISOString() : obj[k];
      });
      return out;
    });
  }

  return {
    assets:      clean(sheetToObjects_('ASSETS')),
    liabilities: clean(sheetToObjects_('LIABILITIES')),
    entities:    clean(sheetToObjects_('ENTITIES')),
    fxRates:     clean(sheetToObjects_('FX')),
    history:     clean(sheetToObjects_('HISTORY')),
    categories:  CATEGORIES,
    currencies:  CURRENCIES
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

  sheet.appendRow([
    id, data.name || '', data.category || '', data.entity || '',
    data.currency || 'USD', localVal, fxRate, usdVal,
    sharePct, shareUsd, now, now, data.notes || '', ''
  ]);
  return { success: true, id: id };
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
      [13, data.notes !== undefined ? data.notes : rows[i][12]]
    ];
    updates.forEach(function(u) { sheet.getRange(i + 1, u[0]).setValue(u[1]); });

    if (Math.abs(usdVal - oldUsd) > 0.01) {
      logHistory_(data.name || rows[i][1], oldUsd, usdVal, currency, data.notes || 'Manual update');
    }
    return { success: true };
  }
  return { success: false, error: 'Asset not found' };
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
  sheet.appendRow([id, data.name || '', data.type || '', data.currency || 'USD', amount, amount * fxRate, now, now, data.notes || '']);
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
    return { success: true };
  }
  return { success: false, error: 'Liability not found' };
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
    var notes    = String(rows[i][12] || '').trim();

    if (category !== 'Real Estate' || currency !== 'USD') continue;

    var addrMatch = notes.match(/address:\s*([^\|]+?)(?:\s*\||$)/i);
    if (!addrMatch) continue;
    var address = addrMatch[1].trim();

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

    var rangeNote = 'Rentcast ' + formatDate_(new Date()) +
                    ': $' + formatNumber_(result.lowValue) + '–$' + formatNumber_(result.highValue);
    var newNotes = notes.replace(/Rentcast [^\|]*/g, rangeNote);
    if (newNotes === notes) newNotes = notes ? notes + ' | ' + rangeNote : rangeNote;
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
  return { success: true, updated: updated };
}

/**
 * Prompt for a US address and show the Rentcast estimate interactively.
 */
function lookupSingleProperty() {
  var ui     = SpreadsheetApp.getUi();
  var apiKey = PropertiesService.getScriptProperties().getProperty('RENTCAST_API_KEY');
  if (!apiKey) {
    ui.alert('RENTCAST_API_KEY not set.\n\nGo to: Extensions > Apps Script > Project Settings > Script Properties\nAdd: RENTCAST_API_KEY = your key from rentcast.io');
    return;
  }
  var resp = ui.prompt('Property Lookup', 'Enter full US address (e.g. 123 Main St, Houston, TX 77001):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var address = resp.getResponseText().trim();
  if (!address) return;
  var result = getRentcastEstimate_(address, apiKey);
  if (result.success) {
    ui.alert('Rentcast Estimate',
      address + '\n\n' +
      'Value:  $' + formatNumber_(result.value) + '\n' +
      'Range:  $' + formatNumber_(result.lowValue) + ' – $' + formatNumber_(result.highValue),
      ui.ButtonSet.OK);
  } else {
    ui.alert('Could not get estimate: ' + result.error);
  }
}

function getRentcastEstimate_(address, apiKey) {
  try {
    var url  = 'https://api.rentcast.io/v1/avm/value?address=' + encodeURIComponent(address);
    var resp = UrlFetchApp.fetch(url, {
      method: 'GET', headers: { 'X-Api-Key': apiKey }, muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code === 404) return { success: false, error: 'Address not found' };
    if (code === 429) return { success: false, error: 'Rate limit (50/month on free tier)' };
    if (code !== 200) return { success: false, error: 'HTTP ' + code };
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
        products:      ['transactions', 'accounts']
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

        for (var i = 1; i < rows.length; i++) {
          if (rows[i][13] === acctId || (rows[i][1] === acctName && rows[i][2] === 'Cash')) {
            var oldUsd = Number(rows[i][7]) || 0;
            sheet.getRange(i + 1, 6).setValue(balance);
            sheet.getRange(i + 1, 7).setValue(1);
            sheet.getRange(i + 1, 8).setValue(balance);
            sheet.getRange(i + 1, 10).setValue(balance);
            sheet.getRange(i + 1, 12).setValue(new Date());
            sheet.getRange(i + 1, 14).setValue(acctId);
            if (Math.abs(balance - oldUsd) > 0.01) logHistory_(acctName, oldUsd, balance, 'USD', 'Plaid sync');
            found = true;
            break;
          }
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
}
