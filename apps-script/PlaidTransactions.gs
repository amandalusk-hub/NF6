/**
 * PlaidTransactions.gs — Master transaction feed from every Plaid-connected
 * account.
 *
 * Why this exists separately from TLMND_TRANSACTIONS:
 *   TLMND_TRANSACTIONS is a FILTERED view scoped to the 3 accounts Amanda
 *   configured for the TLMND cash-flow module (TLMND ···2001, NF USA CA
 *   ···2086, Fidelity ···6454). It carries TLMND-specific user-editable
 *   columns (Category, Recurring, Entity Tag, Notes) + rules-engine state.
 *
 *   But every OTHER Plaid connection Amanda has — Chase family accounts
 *   (Tiger Capital, Family Holdings, Joint Mgmt, Rev Trust, Mike Personal),
 *   Oriental Bank (Dorado), ATH, etc. — never flows into TLMND_TRANSACTIONS
 *   because they're not in that filter. Yet the Money Movement auto-check-
 *   off feature and the Properties / Dorado report both need to match
 *   wires / expenses against the real transactions on these accounts.
 *
 *   This sheet captures the raw transaction feed from EVERY Plaid connection,
 *   no filter. No user-editable columns, no TLMND business logic. Pure data.
 *
 * Readers (Money Movement, Properties) scan this sheet by account-name
 * substring + amount + date, same shape as the TLMND reader.
 */

var PLAID_TX_HEADERS = [
  'Transaction ID',   // A — primary key: 'plaid:' + pending_id or transaction_id
  'Date',             // B — YYYY-MM-DD
  'Account',          // C — "{institution} - {account name} ···{mask}"
  'Account ID',       // D — raw plaid account_id
  'Name',             // E — raw transaction name
  'Merchant',         // F — cleaned merchant name (when Plaid supplies one)
  'Amount USD',       // G — signed: + = money in, − = money out
                      //      (flipped from Plaid's convention where positive = debit)
  'Pending',          // H — Yes / blank
  'Plaid Category',   // I — comma-joined Plaid category breadcrumbs
  'Institution',      // J — e.g. "Chase", "Oriental Bank"
  'Last Synced',      // K — timestamp of last upsert
  'Raw JSON'          // L — full Plaid payload, for debugging / future fields
];


// ── Sheet lifecycle ────────────────────────────────────────────────────────

function _ensurePlaidTxSheet_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PLAID_TRANSACTIONS');
  if (!sheet) {
    sheet = ss.insertSheet('PLAID_TRANSACTIONS');
    sheet.getRange(1, 1, 1, PLAID_TX_HEADERS.length)
      .setValues([PLAID_TX_HEADERS])
      .setFontWeight('bold').setBackground('#14263d').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 220);
    return sheet;
  }
  // Auto-add any missing columns on schema drift.
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var missing = PLAID_TX_HEADERS.filter(function(h) { return existing.indexOf(h) < 0; });
  if (missing.length) {
    var startCol = existing.length + 1;
    sheet.getRange(1, startCol, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold').setBackground('#14263d').setFontColor('#ffffff');
  }
  return sheet;
}


// ── Main sync ──────────────────────────────────────────────────────────────
// Pulls transactions from EVERY Plaid access_token stored in PLAID_TOKENS,
// across EVERY account each token has, and upserts into PLAID_TRANSACTIONS.
// Idempotent — uses pending_transaction_id when available so pending → posted
// updates in place.

function syncAllPlaidTransactions(opts) {
  _requireEditor_();
  var cfg = getPlaidConfig_();
  if (!cfg.clientId || !cfg.secret) throw new Error('Plaid credentials not set. Set PLAID_CLIENT_ID and PLAID_SECRET in script properties.');
  var props  = PropertiesService.getScriptProperties();
  var tokens = JSON.parse(props.getProperty('PLAID_TOKENS') || '[]');
  var instMap = JSON.parse(props.getProperty('PLAID_INSTITUTIONS') || '{}');
  if (!tokens.length) throw new Error('No Plaid connections. Use Tracker → Connect Bank Account first.');

  var lookback = (opts && Number(opts.monthsBack)) || 12;   // default: 12 months
  var end   = new Date();
  var start = new Date(end.getFullYear(), end.getMonth() - lookback, 1);
  function fmt(d) { return d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2) + '-' + ('0'+d.getDate()).slice(-2); }

  var results = {
    success: true, tokens: tokens.length, accountsFound: 0, txnsFetched: 0,
    newRows: 0, updatedRows: 0, errors: [], perToken: []
  };

  var allRecords = [];

  tokens.forEach(function(token) {
    var tokenInfo = { tokenEnd: token.slice(-4), institution: instMap[token] || '', accounts: 0, txns: 0, error: null };

    // 1. /accounts/get — resolve account_id → human label for this token.
    var acctLabels = {};
    try {
      var accResp = UrlFetchApp.fetch(getPlaidBaseUrl_(cfg.env) + '/accounts/get', {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ client_id: cfg.clientId, secret: cfg.secret, access_token: token }),
        muteHttpExceptions: true
      });
      var accData = JSON.parse(accResp.getContentText());
      if (accData.error_code) {
        tokenInfo.error = accData.error_code + ' — ' + (accData.error_message || '');
        results.errors.push('Plaid /accounts/get (' + token.slice(-4) + '): ' + tokenInfo.error);
        results.perToken.push(tokenInfo);
        return;
      }
      (accData.accounts || []).forEach(function(a) {
        var label = (tokenInfo.institution ? tokenInfo.institution + ' - ' : '') +
                    (a.name || 'Account') + ' ···' + (a.mask || '');
        acctLabels[a.account_id] = label;
        tokenInfo.accounts++;
      });
      results.accountsFound += tokenInfo.accounts;
    } catch(e) {
      tokenInfo.error = e.message;
      results.errors.push('Plaid /accounts/get (' + token.slice(-4) + '): ' + e.message);
      results.perToken.push(tokenInfo);
      return;
    }

    // 2. /transactions/get — paginated pull for this token across all accounts.
    var offset = 0, pageSize = 500, pages = 0;
    while (pages < 40) {
      pages++;
      try {
        var txResp = UrlFetchApp.fetch(getPlaidBaseUrl_(cfg.env) + '/transactions/get', {
          method: 'POST', contentType: 'application/json',
          payload: JSON.stringify({
            client_id: cfg.clientId, secret: cfg.secret, access_token: token,
            start_date: fmt(start), end_date: fmt(end),
            options: { count: pageSize, offset: offset }
          }),
          muteHttpExceptions: true
        });
        var txData = JSON.parse(txResp.getContentText());
        if (txData.error_code) {
          tokenInfo.error = txData.error_code + ' — ' + (txData.error_message || '');
          results.errors.push('Plaid /transactions/get (' + token.slice(-4) + '): ' + tokenInfo.error);
          break;
        }
        var batch = txData.transactions || [];
        batch.forEach(function(tx) {
          var stableId = tx.pending_transaction_id || tx.transaction_id;
          allRecords.push({
            key:        'plaid:' + stableId,
            date:       tx.date,
            account:    acctLabels[tx.account_id] || tx.account_id,
            accountId:  tx.account_id,
            name:       tx.name || '',
            merchant:   tx.merchant_name || '',
            amountUsd:  -Number(tx.amount || 0),   // flip Plaid's convention
            pending:    tx.pending ? 'Yes' : '',
            plaidCat:   (tx.category || []).join(' › '),
            institution:tokenInfo.institution,
            raw:        tx
          });
          tokenInfo.txns++;
        });
        if (batch.length === 0 || (offset + batch.length) >= (txData.total_transactions || 0)) break;
        offset += pageSize;
      } catch(e) {
        tokenInfo.error = e.message;
        results.errors.push('Plaid /transactions/get (' + token.slice(-4) + '): ' + e.message);
        break;
      }
    }
    results.txnsFetched += tokenInfo.txns;
    results.perToken.push(tokenInfo);
  });

  // 3. Upsert into PLAID_TRANSACTIONS.
  try {
    var upsertRes = _upsertPlaidTxRows_(allRecords);
    results.newRows = upsertRes.newRows;
    results.updatedRows = upsertRes.updatedRows;
  } catch(e) {
    results.errors.push('Sheet upsert failed: ' + e.message);
    results.success = false;
  }

  return results;
}


// Upsert transactions into PLAID_TRANSACTIONS. Keyed on Transaction ID (col A).
// Rows with an existing key get updated in place; new rows get appended.
function _upsertPlaidTxRows_(records) {
  var sheet = _ensurePlaidTxSheet_();
  var lastRow = sheet.getLastRow();
  var hdr = sheet.getRange(1, 1, 1, PLAID_TX_HEADERS.length).getValues()[0];
  var iKey   = hdr.indexOf('Transaction ID');
  var iSynced = hdr.indexOf('Last Synced');

  // Build existing-key → row-number map (1-indexed sheet row).
  var existingMap = {};
  if (lastRow >= 2) {
    var keyCol = sheet.getRange(2, iKey + 1, lastRow - 1, 1).getValues();
    for (var r = 0; r < keyCol.length; r++) {
      var k = String(keyCol[r][0] || '');
      if (k) existingMap[k] = r + 2;
    }
  }

  var now = new Date();
  var newRows = [];
  var updates = [];
  records.forEach(function(rec) {
    var row = PLAID_TX_HEADERS.map(function(h) {
      switch (h) {
        case 'Transaction ID': return rec.key;
        case 'Date':           return rec.date;
        case 'Account':        return rec.account;
        case 'Account ID':     return rec.accountId;
        case 'Name':           return rec.name;
        case 'Merchant':       return rec.merchant;
        case 'Amount USD':     return rec.amountUsd;
        case 'Pending':        return rec.pending;
        case 'Plaid Category': return rec.plaidCat;
        case 'Institution':    return rec.institution;
        case 'Last Synced':    return now;
        case 'Raw JSON':       try { return JSON.stringify(rec.raw); } catch(e) { return ''; }
        default:               return '';
      }
    });
    if (existingMap[rec.key]) {
      updates.push({ rowNum: existingMap[rec.key], row: row });
    } else {
      newRows.push(row);
    }
  });

  updates.forEach(function(u) {
    sheet.getRange(u.rowNum, 1, 1, u.row.length).setValues([u.row]);
  });
  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, PLAID_TX_HEADERS[0].length ? PLAID_TX_HEADERS.length : newRows[0].length).setValues(newRows);
  }
  return { newRows: newRows.length, updatedRows: updates.length };
}


// Menu-callable wrapper that shows an alert at the end with counts.
function syncAllPlaidTransactionsMenu() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Sync ALL Plaid Transactions',
    'How many months of history to pull? (default: 12)\n\nThis writes every transaction from every Plaid-connected account into the PLAID_TRANSACTIONS sheet. Safe to run — existing rows are updated in place, not duplicated.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var months = Number(String(r.getResponseText() || '').trim()) || 12;
  var res = syncAllPlaidTransactions({ monthsBack: months });

  var lines = [
    'Sync complete.',
    '',
    'Tokens: ' + res.tokens,
    'Accounts: ' + res.accountsFound,
    'Transactions fetched: ' + res.txnsFetched,
    'New rows: ' + res.newRows,
    'Updated rows: ' + res.updatedRows,
    ''
  ];
  if (res.perToken && res.perToken.length) {
    lines.push('PER CONNECTION');
    res.perToken.forEach(function(t) {
      lines.push('  · ' + (t.institution || '(no name)') + ' [' + t.tokenEnd + ']: ' +
                 t.accounts + ' account(s), ' + t.txns + ' txn(s)' +
                 (t.error ? ' ⚠ ' + t.error : ''));
    });
    lines.push('');
  }
  if (res.errors && res.errors.length) {
    lines.push('ERRORS');
    res.errors.slice(0, 10).forEach(function(e) { lines.push('  ⚠ ' + e); });
  }
  ui.alert('PLAID_TRANSACTIONS Sync', lines.join('\n'), ui.ButtonSet.OK);
}


// Debug: list the distinct accounts that appear in PLAID_TRANSACTIONS so
// Amanda can confirm all expected accounts are flowing through.
function debugListPlaidTransactionAccounts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PLAID_TRANSACTIONS');
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('PLAID_TRANSACTIONS is empty. Run "Sync ALL Plaid Transactions" first.');
    return;
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var iAcct = headers.indexOf('Account');
  var iDate = headers.indexOf('Date');
  if (iAcct < 0) return;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  var acctMap = {};
  rows.forEach(function(r) {
    var a = String(r[iAcct] || '').trim();
    if (!a) return;
    if (!acctMap[a]) acctMap[a] = { count: 0, latest: null };
    acctMap[a].count++;
    var d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]);
    if (!isNaN(d.getTime()) && (!acctMap[a].latest || d > acctMap[a].latest)) acctMap[a].latest = d;
  });
  var lines = ['PLAID_TRANSACTIONS ACCOUNTS (' + Object.keys(acctMap).length + ')', ''];
  Object.keys(acctMap).sort().forEach(function(a) {
    var m = acctMap[a];
    lines.push('  • "' + a + '"  ·  ' + m.count + ' txn(s)  ·  latest: ' +
      (m.latest ? Utilities.formatDate(m.latest, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '—'));
  });
  SpreadsheetApp.getUi().alert('PLAID_TRANSACTIONS Accounts', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
