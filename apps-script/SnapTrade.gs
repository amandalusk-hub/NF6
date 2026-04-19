/**
 * SnapTrade API integration for NF6 Family Office Wealth Tracker
 *
 * Script Properties needed:
 *   SNAPTRADE_CLIENT_ID
 *   SNAPTRADE_CONSUMER_KEY
 *   SNAPTRADE_USER_SECRET
 */

const SNAPTRADE_BASE_URL = 'https://api.snaptrade.com';
const SNAPTRADE_USER_ID = 'mnfamilyoffice';

// ============================================================
// CREDENTIALS
// ============================================================

function getSnapTradeConfig_() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('SNAPTRADE_CLIENT_ID');
  const consumerKey = props.getProperty('SNAPTRADE_CONSUMER_KEY');
  if (!clientId || !consumerKey) {
    throw new Error('Missing SNAPTRADE_CLIENT_ID or SNAPTRADE_CONSUMER_KEY.');
  }
  return { clientId: clientId, consumerKey: consumerKey };
}

function getSnapTradeUserSecret_() {
  const userSecret = PropertiesService.getScriptProperties()
    .getProperty('SNAPTRADE_USER_SECRET');
  if (!userSecret) {
    throw new Error('Missing SNAPTRADE_USER_SECRET. Run registerSnapTradeUser() first.');
  }
  return userSecret;
}

// ============================================================
// JSON: sort object keys alphabetically, compact serialization
// ============================================================

function stableStringify_(obj) {
  if (obj === null || obj === undefined) return 'null';
  if (typeof obj !== 'object') return JSON.stringify(obj);
  if (Array.isArray(obj)) {
    return '[' + obj.map(stableStringify_).join(',') + ']';
  }
  const keys = Object.keys(obj).sort();
  const parts = keys.map(function(k) {
    return JSON.stringify(k) + ':' + stableStringify_(obj[k]);
  });
  return '{' + parts.join(',') + '}';
}

// ============================================================
// REQUEST SIGNING
// ============================================================

function snapTradeSign_(consumerKey, fullPath, queryString, bodyObj) {
  // Build signing object. Keys in sigObject top-level are sorted (content, path, query).
  // "content" is the body with its keys sorted, or null for GET.
  const sigObject = {
    content: bodyObj === undefined ? null : bodyObj,
    path: fullPath,
    query: queryString
  };

  // Serialize with alphabetically sorted keys throughout
  const sigContent = stableStringify_(sigObject);

  // KEY INSIGHT: Per SnapTrade's Node SDK, the consumer key is URI-encoded
  // before being used as the HMAC key.
  const hmacKey = encodeURI(consumerKey);

  const sigBytes = Utilities.computeHmacSha256Signature(sigContent, hmacKey);
  return Utilities.base64Encode(sigBytes);
}

function snapTradeRequest_(method, endpointPath, extraQuery, bodyObj) {
  const config = getSnapTradeConfig_();
  const timestamp = Math.floor(Date.now() / 1000).toString();
  const fullPath = '/api/v1' + endpointPath;

  // Canonical query string: keys alphabetically sorted
  const q = Object.assign({}, extraQuery || {}, {
    clientId: config.clientId,
    timestamp: timestamp
  });
  const queryString = Object.keys(q).sort().map(function(k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(q[k]);
  }).join('&');

  const bodyForSigning = (bodyObj === undefined ? null : bodyObj);
  const signature = snapTradeSign_(config.consumerKey, fullPath, queryString, bodyForSigning);

  const url = SNAPTRADE_BASE_URL + fullPath + '?' + queryString;

  // CRITICAL: The body sent over HTTP must be byte-identical to the body
  // that was included in the signing payload. Since we used stableStringify_
  // above (keys sorted), we use it here too.
  const payload = bodyObj ? stableStringify_(bodyObj) : null;

  Logger.log('URL: ' + url);
  Logger.log('Signed body: ' + stableStringify_(bodyForSigning));
  Logger.log('Sent payload: ' + payload);

  const options = {
    method: method,
    contentType: 'application/json',
    headers: { 'Signature': signature },
    muteHttpExceptions: true
  };
  if (payload) options.payload = payload;

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('SnapTrade API error ' + code + ' on ' + method + ' ' + endpointPath + ': ' + text);
  }
  return text ? JSON.parse(text) : {};
}

// ============================================================
// DIAGNOSTICS
// ============================================================

function testSnapTradeSignature() {
  try {
    const response = snapTradeRequest_(
      'POST',
      '/snapTrade/mockSignature',
      {},
      { userId: 'test@example.com', userSecret: 'TEST_SECRET' }
    );
    Logger.log('Signature test PASSED');
    Logger.log(JSON.stringify(response));
  } catch (e) {
    Logger.log('Signature test FAILED: ' + e.message);
  }
}

function diagnoseSnapTradeConfig() {
  const props = PropertiesService.getScriptProperties().getProperties();
  Logger.log('All SNAPTRADE_* properties found:');
  Object.keys(props).forEach(function(key) {
    if (key.indexOf('SNAPTRADE') === 0) {
      const val = props[key];
      const masked = val.length > 12
        ? val.substring(0, 6) + '...' + val.substring(val.length - 4) + ' (len=' + val.length + ')'
        : '(short, len=' + val.length + ')';
      Logger.log('  ' + key + ' = ' + masked);
    }
  });
  Logger.log('SNAPTRADE_USER_ID constant = ' + SNAPTRADE_USER_ID);
}

// ============================================================
// USER REGISTRATION (ONE-TIME)
// ============================================================

function registerSnapTradeUser() {
  const response = snapTradeRequest_(
    'POST',
    '/snapTrade/registerUser',
    {},
    { userId: SNAPTRADE_USER_ID }
  );
  Logger.log('=== SNAPTRADE USER REGISTRATION ===');
  Logger.log('userId: ' + response.userId);
  Logger.log('userSecret: ' + response.userSecret);
  Logger.log('>>> COPY the userSecret and save to Script Properties as SNAPTRADE_USER_SECRET');
  return response;
}

// ============================================================
// BROKERAGE CONNECTION
// ============================================================

function generateConnectionPortalUrl(broker) {
  const userSecret = getSnapTradeUserSecret_();

  // userId and userSecret go in query params, NOT body
  const extraQuery = {
    userId: SNAPTRADE_USER_ID,
    userSecret: userSecret
  };
  // Only broker (optional) goes in body
  const body = broker ? { broker: broker } : null;

  const response = snapTradeRequest_(
    'POST',
    '/snapTrade/login',
    extraQuery,
    body
  );

  Logger.log('=== CONNECTION PORTAL URL ===');
  Logger.log('Broker: ' + (broker || '(user picks)'));
  Logger.log('URL: ' + response.redirectURI);
  Logger.log('>>> Open within 5 minutes');

  return response;
}

function connectSchwab() {
  return generateConnectionPortalUrl('SCHWAB');
}

function connectAnyBroker() {
  return generateConnectionPortalUrl(null);
}
/**
 * Deletes the current SnapTrade user (removes all connections).
 * Safe to run when no brokerage accounts are connected yet.
 */
function deleteSnapTradeUser() {
  const userSecret = getSnapTradeUserSecret_();
  const response = snapTradeRequest_(
    'DELETE',
    '/snapTrade/deleteUser',
    { userId: SNAPTRADE_USER_ID, userSecret: userSecret },
    null
  );
  Logger.log('Delete response: ' + JSON.stringify(response));
  return response;
}

/**
 * Re-registers the user AND automatically saves the new userSecret
 * to Script Properties. This is the safe way to register.
 */
function registerAndSaveSnapTradeUser() {
  const response = snapTradeRequest_(
    'POST',
    '/snapTrade/registerUser',
    {},
    { userId: SNAPTRADE_USER_ID }
  );

  // Save the new secret immediately
  PropertiesService.getScriptProperties()
    .setProperty('SNAPTRADE_USER_SECRET', response.userSecret);

  Logger.log('=== USER REGISTERED AND SECRET SAVED ===');
  Logger.log('userId: ' + response.userId);
  Logger.log('userSecret (first 8 chars): ' + response.userSecret.substring(0, 8) + '...');
  Logger.log('Secret automatically saved to Script Properties.');

  return response;
}
/**
 * Test: call /snapTrade/login with userId/userSecret in query params (like delete),
 * not in the body.
 */
function connectSchwabQueryParams() {
  const userSecret = getSnapTradeUserSecret_();
  const response = snapTradeRequest_(
    'POST',
    '/snapTrade/login',
    {
      userId: SNAPTRADE_USER_ID,
      userSecret: userSecret
    },
    { broker: 'SCHWAB' }  // only broker in body
  );

  Logger.log('=== LOGIN SUCCESS ===');
  Logger.log('URL: ' + response.redirectURI);
  return response;
}
// ============================================================
// LIST ACCOUNTS & PULL DATA
// ============================================================

/**
 * Lists all brokerage accounts connected to the user.
 * Returns array of account objects with id, name, number, institution_name, etc.
 */
function listSnapTradeAccounts() {
  const userSecret = getSnapTradeUserSecret_();

  const response = snapTradeRequest_(
    'GET',
    '/accounts',
    {
      userId: SNAPTRADE_USER_ID,
      userSecret: userSecret
    },
    null
  );

  Logger.log('=== CONNECTED ACCOUNTS ===');
  Logger.log('Total: ' + response.length);
  Logger.log('');
  response.forEach(function(acct, i) {
    Logger.log('[' + i + '] ' + acct.institution_name + ' — ' + acct.name);
    Logger.log('    id: ' + acct.id);
    Logger.log('    number: ' + acct.number);
    Logger.log('    balance: $' + (acct.balance && acct.balance.total ? acct.balance.total.amount : 'n/a'));
    Logger.log('    currency: ' + (acct.balance && acct.balance.total ? acct.balance.total.currency : 'n/a'));
    Logger.log('    status: ' + acct.meta.status);
    Logger.log('');
  });

  return response;
}

/**
 * Gets full holdings (positions, balances, cash) for ALL connected accounts.
 * Useful for a single snapshot of everything.
 */
function getAllHoldings() {
  const userSecret = getSnapTradeUserSecret_();

  const response = snapTradeRequest_(
    'GET',
    '/holdings',
    {
      userId: SNAPTRADE_USER_ID,
      userSecret: userSecret
    },
    null
  );

  Logger.log('=== ALL HOLDINGS ===');
  Logger.log('Accounts returned: ' + response.length);
  Logger.log('');

  response.forEach(function(acct) {
    const info = acct.account || {};
    const balances = acct.balances || [];
    const positions = acct.positions || [];

    Logger.log('Account: ' + info.name + ' (' + info.number + ')');

    balances.forEach(function(bal) {
      Logger.log('  Cash: $' + bal.cash + ' ' + (bal.currency ? bal.currency.code : ''));
    });

    Logger.log('  Positions: ' + positions.length);
    positions.forEach(function(pos) {
      const sym = pos.symbol && pos.symbol.symbol ? pos.symbol.symbol.symbol : '?';
      const units = pos.units || 0;
      const price = pos.price || 0;
      const value = units * price;
      Logger.log('    ' + sym + ': ' + units + ' @ $' + price + ' = $' + value.toFixed(2));
    });
    Logger.log('');
  });

  return response;
}

// ============================================================
// ADDITIONAL BROKER CONNECTIONS
// ============================================================

function connectFidelity() {
  return generateConnectionPortalUrl('FIDELITY');
}

// ============================================================
// SPREADSHEET MENU DIALOG HELPERS
// ============================================================

// Opens a modal dialog with the SnapTrade connection URL so the user can
// click through to link their brokerage account from the spreadsheet menu.
function openSnapTradePortalDialog_(brokerCode, brokerLabel) {
  try {
    var result = generateConnectionPortalUrl(brokerCode);
    var url    = result && result.redirectURI;
    if (!url) throw new Error('No redirectURI returned from SnapTrade.');
    var html = HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:20px">' +
      '<p>Click below to connect your <b>' + brokerLabel + '</b> account.</p>' +
      '<p style="color:#888;font-size:12px">The link expires in 5 minutes. After connecting, ' +
      'close this dialog and run <b>Sync SnapTrade (Schwab + Fidelity)</b>.</p>' +
      '<div style="text-align:center;margin:24px 0">' +
      '<a href="' + url + '" target="_blank" ' +
      'style="background:#1a73e8;color:#fff;padding:12px 28px;border-radius:4px;' +
      'text-decoration:none;font-size:14px">Open Connection Portal</a>' +
      '</div></div>'
    ).setWidth(420).setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(html, 'Connect ' + brokerLabel);
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error generating connection URL: ' + e.message);
  }
}

function connectSchwabDialog() {
  openSnapTradePortalDialog_('SCHWAB', 'Charles Schwab');
}

function connectFidelityDialog() {
  openSnapTradePortalDialog_('FIDELITY', 'Fidelity');
}
