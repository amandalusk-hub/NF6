// ============================================================================
// TLMND CASH FLOW — Phase 1: raw transaction pull
//
// Pulls transactions from a fixed set of accounts (TLMND checking, NF
// California, Fidelity brokerage) into a TLMND_TRANSACTIONS sheet on demand
// and via a daily scheduled trigger. Phase 2 layers a category-rules engine
// over this raw data; Phase 3 renders the dashboard tab.
//
// Data sources:
//   - Plaid /transactions/get for TLMND checking + NF California
//   - SnapTrade /activities for Fidelity brokerage
//
// Idempotency: transactions are upserted by (source, transaction_id). If a
// row already exists for a given transaction_id, its user-editable columns
// (Category, Recurring, Entity Tag, Notes) are PRESERVED and the raw
// columns are refreshed with the latest data from the source.
// ============================================================================

// ── CONFIG ──────────────────────────────────────────────────────────────────
// Account IDs to track. Stored as a script property so we can update without
// a code change. Populated by setTLMNDConfig().
function getTLMNDConfig_() {
  var raw = PropertiesService.getScriptProperties().getProperty('TLMND_CONFIG');
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

function setTLMNDConfig(cfg) {
  PropertiesService.getScriptProperties().setProperty('TLMND_CONFIG', JSON.stringify(cfg));
  return { success: true };
}

// One-time initialization with the account IDs the user provided.
// Idempotent — safe to re-run; it fully replaces the stored config.
function initTLMNDConfigDefaults() {
  var cfg = {
    plaidAccounts: [
      {
        accountId:  'ZYkPYZr7P3CKpn5YoQr6HgnVKVbXdVsVEpBmA',
        label:      'TLMND ···2001',
        role:       'primary'    // this is TLMND's own account — money in/out counts directly
      },
      {
        accountId:  '5PvRPZg8Res95MzpZdKAUa87Z7k3e7s6R5Pvk',
        label:      'NF California',
        role:       'passthrough' // money hits here first, then journals to TLMND
      }
    ],
    snapTradeAccounts: [
      {
        accountId:  '74a46c33-1356-485d-baf9-85d4af3062fd',
        label:      'Fidelity ···6454',
        role:       'primary'
      }
    ],
    lookbackMonths: 3
  };
  setTLMNDConfig(cfg);
  SpreadsheetApp.getUi().alert('TLMND config initialized.\n\n' +
    'Plaid accounts: ' + cfg.plaidAccounts.length + '\n' +
    'SnapTrade accounts: ' + cfg.snapTradeAccounts.length + '\n' +
    'Lookback: ' + cfg.lookbackMonths + ' months');
}

// ── SHEET SCHEMA ────────────────────────────────────────────────────────────
// Column order matters — this is what the sync writes/reads.
var TLMND_TX_HEADERS = [
  'Transaction ID',    // A — primary key: <source>:<id>
  'Date',              // B — YYYY-MM-DD
  'Source',            // C — Plaid | SnapTrade
  'Account',           // D — human label from config
  'Account ID',        // E — raw account_id
  'Name',              // F — raw name from source
  'Merchant',          // G — cleaned merchant name (Plaid only)
  'Amount USD',        // H — signed: + = money in, − = money out
  'Type',              // I — In | Out | Transfer
  'Plaid Category',    // J — auto-category from Plaid (blank for SnapTrade)
  'Category',          // K — user-editable (preserved on re-sync)
  'Recurring',         // L — user-editable (Yes / blank)
  'Entity Tag',        // M — user-editable (TLM / NF / NF6 / Dr M / …)
  'Notes',             // N — user-editable
  'Pending',           // O — Yes / blank
  'Last Synced',       // P — timestamp of last upsert
  'Raw JSON'           // Q — original API payload for reference
];
var TLMND_TX_USER_COLS = ['Category','Recurring','Entity Tag','Notes'];

function _tlmndGetOrCreateTxSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TLMND_TRANSACTIONS');
  if (!sheet) {
    sheet = ss.insertSheet('TLMND_TRANSACTIONS');
    sheet.getRange(1, 1, 1, TLMND_TX_HEADERS.length).setValues([TLMND_TX_HEADERS]).setFontWeight('bold').setBackground('#f8f9fa');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 220);   // Transaction ID
    sheet.setColumnWidth(6, 300);   // Name
    sheet.setColumnWidth(7, 200);   // Merchant
    sheet.setColumnWidth(17, 60);   // Raw JSON — narrow so it doesn't hog space
    sheet.hideColumns(5);           // Account ID — internal
    sheet.hideColumns(17);          // Raw JSON — reference only
  } else {
    // Ensure headers are current (schema migration safety).
    var current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    var needsUpdate = false;
    for (var i = 0; i < TLMND_TX_HEADERS.length; i++) {
      if (current[i] !== TLMND_TX_HEADERS[i]) { needsUpdate = true; break; }
    }
    if (needsUpdate) sheet.getRange(1, 1, 1, TLMND_TX_HEADERS.length).setValues([TLMND_TX_HEADERS]);
  }
  return sheet;
}

// ── MAIN SYNC ───────────────────────────────────────────────────────────────
// Pulls from every configured account, upserts into TLMND_TRANSACTIONS.
// Preserves user-editable columns on re-sync.
function syncTLMNDCashFlow() {
  var cfg = getTLMNDConfig_();
  if (!cfg) return { success: false, error: 'TLMND config not initialized. Run initTLMNDConfigDefaults first.' };

  var lookback = cfg.lookbackMonths || 3;
  var end   = new Date();
  var start = new Date(end.getFullYear(), end.getMonth() - lookback, 1);

  var results = { success: true, plaidCount: 0, snapTradeCount: 0, upserts: 0, newRows: 0, errors: [] };

  // Collect all raw transactions from all sources into one array of upsert-ready records.
  var records = [];

  // Plaid side — batch pull per token, filter by account_id after fetch.
  try {
    var plaidRecords = _tlmndFetchPlaidRecords(cfg.plaidAccounts, start, end, results);
    records = records.concat(plaidRecords);
    results.plaidCount = plaidRecords.length;
  } catch(e) {
    results.errors.push('Plaid fetch failed: ' + e.message);
  }

  // SnapTrade side — one /activities call per account.
  try {
    var stRecords = _tlmndFetchSnapTradeRecords(cfg.snapTradeAccounts, start, end, results);
    records = records.concat(stRecords);
    results.snapTradeCount = stRecords.length;
  } catch(e) {
    results.errors.push('SnapTrade fetch failed: ' + e.message);
  }

  // Upsert into sheet.
  try {
    var upsertResult = _tlmndUpsertRecords(records);
    results.upserts = upsertResult.upserts;
    results.newRows = upsertResult.newRows;
  } catch(e) {
    results.errors.push('Sheet upsert failed: ' + e.message);
    results.success = false;
  }

  results.errorCount = results.errors.length;
  return results;
}

// Fetch transactions from Plaid for the configured accounts. We iterate all
// PLAID_TOKENS and keep only transactions whose account_id matches one in
// our config — a given account lives under one specific token, and we don't
// need to know which up front.
function _tlmndFetchPlaidRecords(plaidAccts, start, end, results) {
  var cfg = getPlaidConfig_();
  if (!cfg.clientId || !cfg.secret) throw new Error('Plaid credentials not set.');
  var props  = PropertiesService.getScriptProperties();
  var tokens = JSON.parse(props.getProperty('PLAID_TOKENS') || '[]');
  if (!tokens.length) throw new Error('No Plaid tokens.');

  function fmt(d) { return d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2) + '-' + ('0'+d.getDate()).slice(-2); }

  // Build lookup: account_id → { label, role }
  var wanted = {};
  plaidAccts.forEach(function(a) { wanted[a.accountId] = a; });

  var records = [];

  tokens.forEach(function(token) {
    // Skip early if none of the wanted accounts are in this Item — do a cheap
    // /accounts/get first to filter out irrelevant tokens.
    var hasWanted = false;
    try {
      var accResp = UrlFetchApp.fetch(getPlaidBaseUrl_(cfg.env) + '/accounts/get', {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ client_id: cfg.clientId, secret: cfg.secret, access_token: token }),
        muteHttpExceptions: true
      });
      var accData = JSON.parse(accResp.getContentText());
      if (accData.error_code) { return; }
      (accData.accounts || []).forEach(function(a) { if (wanted[a.account_id]) hasWanted = true; });
    } catch(e) { return; }
    if (!hasWanted) return;

    // Paginated /transactions/get for this token.
    var offset = 0, pageSize = 500, tries = 0;
    while (tries < 40) {
      tries++;
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
          results.errors.push('Plaid /transactions/get (' + token.slice(-4) + '): ' + txData.error_code + ' — ' + (txData.error_message || ''));
          return;
        }
        var batch = txData.transactions || [];
        batch.forEach(function(tx) {
          var w = wanted[tx.account_id];
          if (!w) return;
          records.push({
            key:        'plaid:' + tx.transaction_id,
            date:       tx.date,
            source:     'Plaid',
            account:    w.label,
            accountId:  tx.account_id,
            name:       tx.name || '',
            merchant:   tx.merchant_name || '',
            amountUsd:  -Number(tx.amount || 0),  // flip Plaid's convention: + = in
            type:       _tlmndInferType(tx),
            plaidCat:   (tx.category || []).join(' › '),
            pending:    tx.pending ? 'Yes' : '',
            raw:        tx
          });
        });
        if (batch.length === 0 || (offset + batch.length) >= (txData.total_transactions || 0)) break;
        offset += pageSize;
      } catch(e) {
        results.errors.push('Plaid tx (' + token.slice(-4) + '): ' + e.message);
        return;
      }
    }
  });

  return records;
}

function _tlmndInferType(tx) {
  var amt  = Number(tx.amount || 0);   // Plaid convention: + = out
  var code = (tx.transaction_code || '').toLowerCase();
  if (code === 'transfer' || code === 'wire') return 'Transfer';
  return amt >= 0 ? 'Out' : 'In';
}

// Fetch activities from SnapTrade for each configured account.
function _tlmndFetchSnapTradeRecords(snapAccts, start, end, results) {
  if (!snapAccts || !snapAccts.length) return [];
  function fmt(d) { return d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2) + '-' + ('0'+d.getDate()).slice(-2); }

  var userSecret;
  try { userSecret = getSnapTradeUserSecret_(); }
  catch(e) { throw new Error('SnapTrade not configured: ' + e.message); }

  var records = [];

  snapAccts.forEach(function(a) {
    try {
      var resp = snapTradeRequest_('GET', '/accounts/' + a.accountId + '/activities', {
        userId:     SNAPTRADE_USER_ID,
        userSecret: userSecret,
        startDate:  fmt(start),
        endDate:    fmt(end)
      }, null);
      // SnapTrade returns an array of activity objects.
      (resp || []).forEach(function(act) {
        var amtRaw   = act.amount != null ? Number(act.amount) : (act.price != null ? Number(act.price) * Number(act.units || 0) : 0);
        var currency = (act.currency && act.currency.code) || 'USD';
        // Only USD for now; skip anything else (family office is USD-book)
        if (currency !== 'USD') return;
        var actType   = String(act.type || '').toUpperCase();
        var dateStr   = String(act.trade_date || act.settlement_date || act.date || '').substring(0, 10);
        if (!dateStr) return;

        // Sign: for SnapTrade, deposits and dividends are IN, withdrawals and fees are OUT.
        // "amount" is typically already signed for CASH activities but positive for BUY/SELL.
        // Normalize: BUY = out (money out), SELL = in (money in), otherwise keep sign.
        var signedAmt = amtRaw;
        if (actType === 'BUY')            signedAmt = -Math.abs(amtRaw);
        else if (actType === 'SELL')      signedAmt =  Math.abs(amtRaw);
        else if (actType === 'DEPOSIT' || actType === 'DIVIDEND' || actType === 'INTEREST' || actType === 'TAX_REFUND')
                                          signedAmt =  Math.abs(amtRaw);
        else if (actType === 'WITHDRAWAL' || actType === 'FEE' || actType === 'TAX')
                                          signedAmt = -Math.abs(amtRaw);

        var typeLabel = signedAmt >= 0 ? 'In' : 'Out';
        if (actType === 'TRANSFER')       typeLabel = 'Transfer';

        var name = actType +
          (act.symbol && act.symbol.symbol ? ' ' + act.symbol.symbol : '') +
          (act.description ? ' — ' + act.description : '');

        records.push({
          key:        'st:' + (act.id || (a.accountId + '-' + dateStr + '-' + signedAmt)),
          date:       dateStr,
          source:     'SnapTrade',
          account:    a.label,
          accountId:  a.accountId,
          name:       name.trim(),
          merchant:   '',
          amountUsd:  signedAmt,
          type:       typeLabel,
          plaidCat:   '',
          pending:    '',
          raw:        act
        });
      });
    } catch(e) {
      results.errors.push('SnapTrade ' + a.label + ': ' + e.message);
    }
  });

  return records;
}

// Upsert into TLMND_TRANSACTIONS. Preserves user-editable columns on update;
// rewrites raw columns from the latest source data. Idempotent by
// Transaction ID (column A).
function _tlmndUpsertRecords(records) {
  var sheet   = _tlmndGetOrCreateTxSheet();
  var lastRow = sheet.getLastRow();
  var existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, TLMND_TX_HEADERS.length).getValues() : [];

  // Build map: key → { rowIdx (1-based), row (values) }
  var byKey = {};
  for (var i = 0; i < existing.length; i++) {
    var key = String(existing[i][0] || '');
    if (key) byKey[key] = { rowIdx: i + 2, row: existing[i] };
  }

  var now = new Date();
  var newRows = [];
  var updates = [];   // { rowIdx, values }
  var upsertCount = 0;

  records.forEach(function(r) {
    var newVals = [
      r.key,                              // A Transaction ID
      r.date,                             // B Date
      r.source,                           // C Source
      r.account,                          // D Account
      r.accountId,                        // E Account ID
      r.name,                             // F Name
      r.merchant,                         // G Merchant
      r.amountUsd,                        // H Amount USD
      r.type,                             // I Type
      r.plaidCat,                         // J Plaid Category
      '',                                 // K Category (preserved on update)
      '',                                 // L Recurring (preserved)
      '',                                 // M Entity Tag (preserved)
      '',                                 // N Notes (preserved)
      r.pending,                          // O Pending
      now,                                // P Last Synced
      JSON.stringify(r.raw || {})         // Q Raw JSON
    ];

    if (byKey[r.key]) {
      // UPDATE — preserve user-editable columns from existing row.
      var oldRow = byKey[r.key].row;
      newVals[10] = oldRow[10];  // Category
      newVals[11] = oldRow[11];  // Recurring
      newVals[12] = oldRow[12];  // Entity Tag
      newVals[13] = oldRow[13];  // Notes
      updates.push({ rowIdx: byKey[r.key].rowIdx, values: newVals });
    } else {
      // INSERT
      newRows.push(newVals);
    }
    upsertCount++;
  });

  // Apply updates one row at a time (typically few per sync).
  updates.forEach(function(u) {
    sheet.getRange(u.rowIdx, 1, 1, TLMND_TX_HEADERS.length).setValues([u.values]);
  });

  // Append new rows in a single batch.
  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, TLMND_TX_HEADERS.length).setValues(newRows);
  }

  return { upserts: upsertCount, newRows: newRows.length };
}

// ── MENU WRAPPERS + TRIGGER ─────────────────────────────────────────────────
function syncTLMNDCashFlowMenu() {
  var ui = SpreadsheetApp.getUi();
  var r  = syncTLMNDCashFlow();
  var msg = 'Plaid transactions:     ' + (r.plaidCount || 0) +
            '\nSnapTrade activities: ' + (r.snapTradeCount || 0) +
            '\nUpserts:                     ' + (r.upserts || 0) +
            '\nNew rows:                  ' + (r.newRows || 0) +
            '\nErrors:                        ' + (r.errorCount || 0);
  if (r.errors && r.errors.length) msg += '\n\nErrors:\n  • ' + r.errors.slice(0, 8).join('\n  • ');
  ui.alert(r.success ? 'TLMND Cash Flow Synced' : 'TLMND Cash Flow: Errors', msg, ui.ButtonSet.OK);
}

// Daily scheduled trigger — runs at ~4:30 AM in the script's timezone.
// Runs quietly (no UI); errors surface in the executions log.
function _tlmndDailySync() { syncTLMNDCashFlow(); }

function installTLMNDCashFlowTrigger() {
  var ui = SpreadsheetApp.getUi();
  var existing = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === '_tlmndDailySync'; });
  existing.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('_tlmndDailySync').timeBased().atHour(4).nearMinute(30).everyDays(1).create();
  ui.alert('Installed daily TLMND cash flow sync at 4:30 AM.\n\n' +
           'Replaced ' + existing.length + ' prior trigger(s) if any.');
}
