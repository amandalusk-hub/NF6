/**
 * Code.gs — Family Office Wealth Tracker
 * Google Apps Script Web App Backend
 *
 * SCRIPT PROPERTIES TO SET:
 *   SPREADSHEET_ID   — ID of the backing Google Sheet
 *   RENTCAST_API_KEY — from rentcast.io (free tier: 50 req/month)
 *   PLAID_CLIENT_ID  — 6997417fe8a45f001e390093
 *   PLAID_SECRET     — 08b872e451d778ce028d2b5952693b
 *   PLAID_ENV        — sandbox | production
 */

// ── Constants ────────────────────────────────────────────────────────────────

var CATEGORIES = [
  'Real Estate', 'Private Equity', 'Public Equity',
  'Cash', 'Crypto', 'Auto', 'Art/Jewelry',
  'VIP Medical', 'Insurance', 'Other'
];

var CURRENCIES = ['USD','EUR','GBP','COP','BRL','MXN','CAD','JPY','CHF','AUD','DOP'];

var COL = {
  ASSETS:   ['ID','Name','Category','Entity','Currency','Local Value','USD Rate','USD Value','My Share %','My Share USD','Date Added','Last Updated','Notes','Plaid Account ID'],
  ENTITIES: ['Name','Type','Jurisdiction','Ownership %','Notes'],
  FX:       ['Currency','Rate to USD','Last Fetched'],
  HISTORY:  ['Date','Asset Name','Old Value USD','New Value USD','Delta USD','Currency','Notes']
};

// ── Menu ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tracker')
    .addItem('Check Setup', 'deploymentReadinessCheck')
    .addSeparator()
    .addItem('Refresh FX Rates', 'fetchExchangeRates')
    .addItem('Refresh US Property Values', 'refreshPropertyValues')
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

  // Helper to record a check result
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
    check('Sheets (Assets / Entities / FX Rates / History)', true, 'all present');
  } catch(e) {
    check('Sheets', false, e.message);
  }

  // 3. FX API (free, no key required)
  try {
    var fxResp = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/USD', { muteHttpExceptions: true });
    check('FX API (open.er-api.com)', fxResp.getResponseCode() === 200, 'HTTP ' + fxResp.getResponseCode());
  } catch(e) {
    check('FX API (open.er-api.com)', false, e.message);
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

  // 6. Plaid API connectivity (only if credentials are present)
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
  var triggers = ScriptApp.getProjectTriggers();
  var hasTrigger = triggers.some(function(t) { return t.getHandlerFunction() === 'dailySync_'; });
  check('Daily sync trigger', hasTrigger, hasTrigger ? 'installed' : 'not installed — run "Install Daily Trigger"');

  // 8. Web app deployment (informational only — can't verify programmatically)
  checks.push('ℹ️  Web app: deploy via Deploy > New deployment if using the web UI');

  var summary = (ok ? '✅ All required checks passed.' : '⚠️  Some checks failed — see details below.') +
    '\n\n' + checks.join('\n');

  SpreadsheetApp.getUi().alert('Deployment Readiness Check', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  return { ok: ok, checks: checks };
}

// ── Entry Point ──────────────────────────────────────────────────────────────

function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Family Office — Wealth Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Sheet Bootstrapping ──────────────────────────────────────────────────────

function ensureSheets_() {
  var ss = getSpreadsheet_();
  Object.keys(COL).forEach(function(key) {
    var name = sheetName_(key);
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
  return { ASSETS: 'Assets', ENTITIES: 'Entities', FX: 'FX Rates', HISTORY: 'History' }[key];
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

// ── Main Data Fetch ──────────────────────────────────────────────────────────

function getFullData() {
  ensureSheets_();
  var assets   = sheetToObjects_('ASSETS');
  var entities = sheetToObjects_('ENTITIES');
  var fx       = sheetToObjects_('FX');
  var history  = sheetToObjects_('HISTORY');

  // Serialize dates
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
    assets:     clean(assets),
    entities:   clean(entities),
    fxRates:    clean(fx),
    history:    clean(history),
    categories: CATEGORIES,
    currencies: CURRENCIES
  };
}

// ── FX Rates ─────────────────────────────────────────────────────────────────

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

    // Overwrite FX sheet
    var last = sheet.getLastRow();
    if (last > 1) sheet.getRange(2, 1, last - 1, 3).clearContent();
    if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);

    // Cache
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
      var obj = JSON.parse(cached);
      var ageHrs = (Date.now() - new Date(obj.fetched).getTime()) / 3600000;
      if (ageHrs < 4 && obj.rates[currency]) return obj.rates[currency];
    }
  } catch(e) {}
  var result = fetchExchangeRates();
  return (result.success && result.rates[currency]) ? result.rates[currency] : 1;
}

// ── Assets CRUD ──────────────────────────────────────────────────────────────

function addAsset(data) {
  var sheet      = getSheet_('ASSETS');
  var id         = Utilities.getUuid();
  var now        = new Date();
  var fxRate     = getFxRate_(data.currency || 'USD');
  var localVal   = Number(data.localValue) || 0;
  var usdVal     = localVal * fxRate;
  var sharePct   = data.mySharePct !== undefined ? Number(data.mySharePct) : 100;
  var shareUsd   = usdVal * sharePct / 100;

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

// ── Property Valuation (Rentcast) ─────────────────────────────────────────────
// Zillow deprecated their public API; Rentcast provides free property AVM.
// Free tier: 50 req/month. Get key at rentcast.io
// Set RENTCAST_API_KEY in Script Properties.

function getPropertyValue(address) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('RENTCAST_API_KEY');
  if (!apiKey) return { success: false, error: 'Set RENTCAST_API_KEY in Script Properties' };

  try {
    var url = 'https://api.rentcast.io/v1/avm/value?address=' + encodeURIComponent(address);
    var resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'X-Api-Key': apiKey },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() === 404) return { success: false, error: 'Address not found' };
    if (resp.getResponseCode() !== 200) return { success: false, error: 'API error ' + resp.getResponseCode() };

    var data = JSON.parse(resp.getContentText());
    return {
      success: true,
      value:      data.price      || data.value || null,
      lowValue:   data.priceLow   || null,
      highValue:  data.priceHigh  || null,
      address:    data.formattedAddress || address
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function refreshPropertyValues() {
  var sheet = getSheet_('ASSETS');
  var rows  = sheet.getDataRange().getValues();
  var updated = 0;

  for (var i = 1; i < rows.length; i++) {
    var category = rows[i][2];
    var notes    = (rows[i][12] || '').toString();
    var address  = '';

    // Look for US address in notes field (format: "address: 123 Main St, City, ST")
    var match = notes.match(/address:\s*(.+?)(?:\||$)/i);
    if (match) address = match[1].trim();

    if (category !== 'Real Estate' || !address) continue;
    if (!/\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY|DC)\b/i.test(address)) continue;

    var result = getPropertyValue(address);
    if (!result.success || !result.value) continue;

    var oldUsd = Number(rows[i][7]) || 0;
    var newUsd = result.value;
    sheet.getRange(i + 1, 6).setValue(newUsd);  // Local Value (USD for US properties)
    sheet.getRange(i + 1, 7).setValue(1);
    sheet.getRange(i + 1, 8).setValue(newUsd);
    var sharePct = Number(rows[i][8]) || 100;
    sheet.getRange(i + 1, 10).setValue(newUsd * sharePct / 100);
    sheet.getRange(i + 1, 12).setValue(new Date());

    if (Math.abs(newUsd - oldUsd) > 0.01) {
      logHistory_(rows[i][1], oldUsd, newUsd, 'USD', 'Auto-updated via Rentcast');
    }
    updated++;
    Utilities.sleep(500); // Rate limit
  }

  return { success: true, updated: updated };
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
    clientId: p.getProperty('PLAID_CLIENT_ID') || '6997417fe8a45f001e390093',
    secret:   p.getProperty('PLAID_SECRET')    || '08b872e451d778ce028d2b5952693b',
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
        client_id:    cfg.clientId,
        secret:       cfg.secret,
        client_name:  'MNW Family Office',
        country_codes: ['US'],
        language:     'en',
        user:         { client_user_id: 'mnw-family-office' },
        products:     ['transactions', 'accounts']
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
        var balance  = acct.balances.current || 0;
        var acctName = (acct.name || 'Account') + ' ···' + (acct.mask || '');
        var acctId   = acct.account_id;

        // Find existing asset by Plaid Account ID or name
        var sheet = getSheet_('ASSETS');
        var rows  = sheet.getDataRange().getValues();
        var found = false;

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
            found = true; break;
          }
        }

        if (!found) {
          addAsset({ name: acctName, category: 'Cash', currency: 'USD', localValue: balance, mySharePct: 100, notes: 'Plaid: ' + acctId });
          // Also store the Plaid account ID
          var newRows = sheet.getDataRange().getValues();
          var lastRow = newRows.length;
          sheet.getRange(lastRow, 14).setValue(acctId);
        }
        synced++;
      });
    } catch(e) { console.error('Plaid sync error:', e); }
  });

  return { success: true, synced: synced };
}

// ── Install Triggers ──────────────────────────────────────────────────────────

function installTriggers() {
  // Remove old
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (['dailySync_', 'fetchExchangeRates'].indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Daily at 7 AM: refresh FX + property values + Plaid
  ScriptApp.newTrigger('dailySync_').timeBased().everyDays(1).atHour(7).create();
  return { success: true };
}

function dailySync_() {
  fetchExchangeRates();
  syncPlaidAccounts();
  refreshPropertyValues();
}
