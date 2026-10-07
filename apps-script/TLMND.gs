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

function setTLMNDConfig(cfg) { _requireEditor_();
  PropertiesService.getScriptProperties().setProperty('TLMND_CONFIG', JSON.stringify(cfg));
  return { success: true };
}

// One-time initialization with the account IDs the user provided.
// Idempotent — safe to re-run; it fully replaces the stored config.
function initTLMNDConfigDefaults() { _requireEditor_();
  var cfg = {
    plaidAccounts: [
      {
        accountId:  'q7XnKYAAxKs1eXmRZZrAupZNM9REbEC1MbAzb',
        label:      'TLMND ···2001',
        role:       'primary'    // this is TLMND's own account — money in/out counts directly
      },
      {
        accountId:  'Bv9mLEzzVLhwxM5v66XRSX1P3Qr989HvXAOz1',
        label:      'NF USA CA ···2086',
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
    lookbackMonths: 6
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
function syncTLMNDCashFlow(opts) { _requireEditor_();
  var cfg = getTLMNDConfig_();
  if (!cfg) return { success: false, error: 'TLMND config not initialized. Run initTLMNDConfigDefaults first.' };

  // opts.monthsBack lets a caller override the config for a one-time deep
  // pull (backfilling old months that weren't captured by earlier syncs).
  // Defaults to the config's rolling lookback (3 months).
  var lookback = (opts && Number(opts.monthsBack)) || cfg.lookbackMonths || 3;
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

  // Auto-categorize via rules engine (if any rules exist). Non-fatal on failure.
  try {
    var ruleRes = applyTLMNDRules();
    if (ruleRes.success) {
      results.categorized   = ruleRes.categorized;
      results.excluded      = ruleRes.excluded;
      results.uncategorized = ruleRes.uncategorized;
    }
  } catch(e) {
    results.errors.push('Rules apply failed: ' + e.message);
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
  var found   = {};   // wanted accountId → token slice(-4) that has it
  var scanned = 0;

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
      if (accData.error_code) {
        results.errors.push('Plaid /accounts/get (' + token.slice(-4) + '): ' + accData.error_code + ' — ' + (accData.error_message || ''));
        return;
      }
      (accData.accounts || []).forEach(function(a) {
        scanned++;
        if (wanted[a.account_id]) { hasWanted = true; found[a.account_id] = token.slice(-4); }
      });
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
          // Key by pending_transaction_id when Plaid provides one — that's
          // the stable identifier across the pending → posted lifecycle.
          // When a pending tx becomes posted, Plaid issues a NEW
          // transaction_id but the pending_transaction_id stays the same,
          // so this key lets the upsert replace the pending row rather
          // than spawn a duplicate posted row.
          var stableId = tx.pending_transaction_id || tx.transaction_id;
          records.push({
            key:        'plaid:' + stableId,
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

  // Surface which configured accountIds we could NOT find in any token.
  // This is the most common "0 transactions returned" cause — the IDs
  // provided don't match any account under the current PLAID_TOKENS set
  // (possibly because Chase re-issued them during a re-link, or the ID
  // came from an Item that was later removed).
  var missing = plaidAccts.filter(function(a) { return !found[a.accountId]; });
  if (missing.length) {
    missing.forEach(function(m) {
      results.errors.push('Configured Plaid account NOT FOUND across ' + tokens.length +
        ' token(s) / ' + scanned + ' account(s): ' + m.label + ' (' + m.accountId.substring(0, 8) + '···' + m.accountId.slice(-6) + ')');
    });
  }

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
      // SnapTrade transactions/activities endpoint is /activities (top-level),
      // with `accounts` as a query param (comma-separated IDs), NOT
      // /accounts/{id}/activities. Response is either an array directly or an
      // object with .data / .transactions / .results wrapping the array —
      // handle all three shapes defensively.
      var resp = snapTradeRequest_('GET', '/activities', {
        userId:     SNAPTRADE_USER_ID,
        userSecret: userSecret,
        accounts:   a.accountId,
        startDate:  fmt(start),
        endDate:    fmt(end)
      }, null);
      var arr = Array.isArray(resp) ? resp
              : (resp && Array.isArray(resp.data))         ? resp.data
              : (resp && Array.isArray(resp.transactions)) ? resp.transactions
              : (resp && Array.isArray(resp.results))      ? resp.results
              : [];
      arr.forEach(function(act) {
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
// One-time deep backfill — pulls the last 24 months of Plaid transactions
// into TLMND_TRANSACTIONS. Use this to recover months that weren't captured
// by earlier rolling 3-month syncs (e.g. Solaris's January 2026 payment
// that came in before any sync ran).
function syncTLMNDCashFlowDeepMenu() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Deep TLMND Sync',
    'Pull how many months of Plaid history? (default: 24)\n\nThis backfills TLMND_TRANSACTIONS so older payments (like Solaris\'s January 2026 payment) can be matched. It\'s safe to run — existing rows are upserted, not duplicated.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var months = Number(resp.getResponseText()) || 24;
  if (months < 1 || months > 24) { ui.alert('Enter 1–24 months.'); return; }
  var r = syncTLMNDCashFlow({ monthsBack: months });
  var msg = 'Deep sync (' + months + ' months back):\n\n' +
            'Plaid transactions:       ' + (r.plaidCount || 0) +
            '\nSnapTrade activities:  ' + (r.snapTradeCount || 0) +
            '\nUpserts:                        ' + (r.upserts || 0) +
            '\nNew rows:                     ' + (r.newRows || 0);
  ui.alert(r.success ? 'Deep Sync Complete' : 'Deep Sync Errors', msg, ui.ButtonSet.OK);
}

function syncTLMNDCashFlowMenu() {
  var ui = SpreadsheetApp.getUi();
  var r  = syncTLMNDCashFlow();
  var msg = 'Plaid transactions:       ' + (r.plaidCount || 0) +
            '\nSnapTrade activities:  ' + (r.snapTradeCount || 0) +
            '\nUpserts:                        ' + (r.upserts || 0) +
            '\nNew rows:                     ' + (r.newRows || 0) +
            '\nCategorized:                 ' + (r.categorized || 0) +
            '\n  of which [EXCLUDED]: ' + (r.excluded || 0) +
            '\nStill uncategorized:      ' + (r.uncategorized || 0) +
            '\nErrors:                           ' + (r.errorCount || 0);
  if (r.errors && r.errors.length) msg += '\n\nErrors:\n  • ' + r.errors.slice(0, 8).join('\n  • ');
  ui.alert(r.success ? 'TLMND Cash Flow Synced' : 'TLMND Cash Flow: Errors', msg, ui.ButtonSet.OK);
}

// Daily scheduled trigger — runs at ~4:30 AM in the script's timezone.
// Runs quietly (no UI); errors surface in the executions log.
function _tlmndDailySync() { syncTLMNDCashFlow(); }

// Diagnostic — dumps every Plaid account_id across every current token +
// which SnapTrade account_ids are visible, and highlights whether the ones
// in TLMND_CONFIG were actually found. Use when Sync TLMND Cash Flow
// returns 0 transactions or "account NOT FOUND" errors.
function diagnoseTLMNDConfig() {
  var ui  = SpreadsheetApp.getUi();
  var cfg = getTLMNDConfig_();
  if (!cfg) { ui.alert('TLMND_CONFIG not set. Run Initialize first.'); return; }

  var out = [];
  out.push('TLMND CONFIG DIAGNOSTIC');
  out.push('Generated: ' + new Date().toISOString());
  out.push('');
  out.push('── CONFIGURED ACCOUNTS ──');
  cfg.plaidAccounts.forEach(function(a) {
    out.push('  Plaid    · ' + a.label + '  (role=' + a.role + ')');
    out.push('            id: ' + a.accountId);
  });
  cfg.snapTradeAccounts.forEach(function(a) {
    out.push('  SnapTrade · ' + a.label + '  (role=' + a.role + ')');
    out.push('            id: ' + a.accountId);
  });
  out.push('');
  out.push('── AVAILABLE PLAID ACCOUNTS (across all tokens) ──');

  var pcfg = getPlaidConfig_();
  var tokens = JSON.parse(PropertiesService.getScriptProperties().getProperty('PLAID_TOKENS') || '[]');
  var instMap = JSON.parse(PropertiesService.getScriptProperties().getProperty('PLAID_INSTITUTIONS') || '{}');
  var found = {};

  tokens.forEach(function(token, i) {
    var label = instMap[token] || '(unnamed)';
    out.push('');
    out.push('[' + (i+1) + '] ' + label + '   token ···' + token.slice(-4));
    try {
      var resp = UrlFetchApp.fetch(getPlaidBaseUrl_(pcfg.env) + '/accounts/get', {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ client_id: pcfg.clientId, secret: pcfg.secret, access_token: token }),
        muteHttpExceptions: true
      });
      var data = JSON.parse(resp.getContentText());
      if (data.error_code) {
        out.push('    ERROR: ' + data.error_code + ' — ' + (data.error_message || ''));
        return;
      }
      (data.accounts || []).forEach(function(a) {
        var wanted = cfg.plaidAccounts.filter(function(w) { return w.accountId === a.account_id; });
        var marker = wanted.length ? '  ✅ MATCHES ' + wanted[0].label : '';
        out.push('    · ' + (a.name || 'Account') + '  ···' + (a.mask || '????') + marker);
        out.push('      id: ' + a.account_id);
        if (wanted.length) found[a.account_id] = true;
      });
    } catch(e) {
      out.push('    fetch exception: ' + e.message);
    }
  });

  out.push('');
  out.push('── MATCH SUMMARY ──');
  cfg.plaidAccounts.forEach(function(a) {
    out.push('  ' + (found[a.accountId] ? '✅' : '❌') + ' ' + a.label + ' (' + a.accountId.substring(0, 8) + '···' + a.accountId.slice(-6) + ')');
  });

  var reportText = out.join('\n');
  var stamp = new Date().toISOString().substring(0, 19).replace(/[:T]/g, '-');
  var file  = DriveApp.createFile('TLMND Config Diagnostic ' + stamp + '.txt', reportText, MimeType.PLAIN_TEXT);

  var shown = reportText;
  if (shown.length > 4200) shown = shown.substring(0, 4200) + '\n\n… (truncated — full report in Drive)';
  ui.alert('TLMND Config Diagnostic', shown + '\n\nFull report:\n' + file.getUrl(), ui.ButtonSet.OK);
}

// ============================================================================
// PHASE 2 — CATEGORY RULES ENGINE
//
// The TLMND_CATEGORY_RULES sheet is the source of truth for turning raw
// transaction names into your line-item categories (Solaris-Fl Loan Repay,
// Ellison Medical, Payroll, etc). Applied automatically after every sync,
// and on demand via "Apply TLMND Rules" menu.
//
// Schema (columns A–L):
//   Priority       — lower runs first; first match wins per transaction
//   Match Field    — Name | Merchant | Account (what to test against)
//   Match Type     — contains | starts_with | regex | equals (case-insensitive)
//   Pattern        — the pattern to match
//   Amount Min     — optional; skip rule if amount < this (blank = no min)
//   Amount Max     — optional; skip rule if amount > this
//   Category       — line item name (e.g. "Solaris-Fl Holding LLC Loan Repayment Income")
//   Recurring      — Yes | No | blank
//   Entity Tag     — TLM | NF | NF6 | Dr M | TLN | TLW | blank (attribution)
//   Exclude        — Yes | blank; if Yes, the transaction is excluded from
//                    cash-flow totals (used for internal transfers and
//                    Fidelity SPAXX cash-mgmt noise that would double-count)
//   Enabled        — Yes | blank; blank disables the rule (leave in place
//                    but don't apply)
//   Notes          — free text for you
//
// Manual overrides: if you set the Category column on a transaction manually
// (a value that doesn't match any rule), it's preserved on re-sync. Only
// rows whose Category is either blank OR was set by a rule (matches some
// rule's Category) get re-categorized.
// ============================================================================

var TLMND_RULES_HEADERS = [
  'Priority',       // A
  'Match Field',    // B
  'Match Type',     // C
  'Pattern',        // D
  'Amount Min',     // E
  'Amount Max',     // F
  'Category',       // G
  'Recurring',      // H
  'Entity Tag',     // I
  'Exclude',        // J
  'Enabled',        // K
  'Notes',          // L
  'Source Filter',  // M — Plaid | SnapTrade | blank (any). Restricts the
                    //     rule to only match rows from a specific source.
  'Expected Monthly Override'  // N — signed number (positive = income,
                    //     negative = expense). When set, this value
                    //     overrides the auto-calculated forecast
                    //     average for the rule's Category in the
                    //     "Next 30 Days Expected" / "On-Track"
                    //     calculations. Useful when historical data
                    //     doesn't reflect the true monthly rate
                    //     (e.g. MacDonald paid a $8,333 lump but the
                    //     real monthly is $2,083).
];

function _tlmndGetOrCreateRulesSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TLMND_CATEGORY_RULES');
  if (!sheet) {
    sheet = ss.insertSheet('TLMND_CATEGORY_RULES');
    sheet.getRange(1, 1, 1, TLMND_RULES_HEADERS.length).setValues([TLMND_RULES_HEADERS])
      .setFontWeight('bold').setBackground('#f8f9fa');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(4, 300);   // Pattern
    sheet.setColumnWidth(7, 280);   // Category
    sheet.setColumnWidth(12, 200);  // Notes
  }
  return sheet;
}

// Starter rule set matched to the patterns actually observed in the user's
// TLMND_TRANSACTIONS data. Idempotent: replaces the sheet's rules entirely
// each time it's run — so you can re-seed after schema changes without
// creating duplicates. Any custom rules you added by hand will be lost — add
// them again after re-seeding.
function seedTLMNDRules() { _requireEditor_();
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Seed TLMND Rules',
    'This will REPLACE all rules in TLMND_CATEGORY_RULES with a starter set matched to your ' +
    'spreadsheet\'s line items. Any custom rules you added will be lost — you can re-add them after.\n\nContinue?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var sheet = _tlmndGetOrCreateRulesSheet();
  var last  = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, TLMND_RULES_HEADERS.length).clearContent();

  // Rule schema: [Priority, Match Field, Match Type, Pattern, Amt Min, Amt Max, Category, Recurring, Entity Tag, Exclude, Enabled, Notes]
  var rules = [
    // ── MONEY IN — recurring ─────────────────────────────────────────────
    [10, 'Name', 'contains', 'SOLARIS-FL HOLDI',                  '', '', 'Solaris-Fl Holding LLC Loan Repayment Income', 'Yes', 'TLMND',      '', 'Yes', 'Monthly ~$16,656'],
    // ELLISON MEDICAL: raw deposit lands on NF USA CA ···2086 but is
    // economically TLMND income (per user direction — the money is
    // earmarked for TLMND). Count the deposit itself as recurring
    // Money In; the internal transfer from NF USA CA to TLMND is
    // excluded below to prevent double-counting.
    [10, 'Name', 'contains', 'ELLISON MEDICAL',                   '', '', 'Ellison Medical - Customer (Carroll Canyon)',   'Yes', 'TLMND',      '', 'Yes', 'Deposit lands on NF USA CA ···2086, counted as TLMND income'],
    [10, 'Name', 'contains', 'BOOK TRANSFER CREDIT B/O: WASICA',  '', '', 'Wasica Holdings (Book Credit)',                 'Yes', 'TLMND',      '', 'Yes', 'Recurring inbound'],
    [10, 'Name', 'contains', 'CHERRY VALLEY',                     '', '', 'MacDonald Loan Repayment',                     'Yes', 'TLMND',      '', 'Yes', 'Recurring — from Cherry Valley Construction', '', 2083.33],
    [10, 'Name', 'contains', 'SA NJ REALTY',                      '', '', 'ASC Rental Income - TLMND Share (SA NJ Realty)', 'Yes', 'TLMND',   '', 'Yes', 'Recurring — Mike\'s real estate rent (comes in every so often)'],
    // WASKAR TEJEDA payments (CHIPS credits, wires) — categorize with Wasica
    // Holdings per user. Priority 15 keeps Penn Mutual Life Insurance
    // (priority 10) winning first for its specific pattern even though
    // that transaction also references Waskar in its wire memo.
    [15, 'Name', 'contains', 'WASKAR',                             '', '', 'Wasica Holdings (Book Credit)',                  'Yes', 'TLMND',   '', 'Yes', 'Waskar Tejeda transfers routed to Wasica Holdings category'],

    // ── MONEY OUT — recurring ────────────────────────────────────────────
    [10, 'Name', 'contains', 'UNITED HEALTHCAR',                  '', '', 'United Healthcare Insurance',                   'Yes', 'TLMND',        '', 'Yes', 'Monthly ~$9,764'],
    [10, 'Name', 'contains', 'THE GUARDIAN',                      '', '', 'The Guardian Insurance',                        'Yes', 'TLMND',        '', 'Yes', 'Monthly ~$625'],
    [10, 'Name', 'contains', 'DIVVY',                             '', '', 'Divvy Bill (Grand Total)',                      'Yes', 'TLMND',        '', 'Yes', 'Broad match — catches EWALLET Divvy ACH, DIVVY PEACH LLC wire payments, etc.'],
    [10, 'Name', 'contains', 'BSCAccountingLLC',                  '', '', 'BSC Accounting LLC (Accounting Fees)',          'Yes', 'TLMND',        '', 'Yes', 'Monthly -$3,500'],
    [10, 'Name', 'contains', 'PENN MUTUAL LIFE INS',              '', '', 'Life Insurance (Waskar Tejeda / Penn Mutual)',  'No',  'TLMND',      '', 'Yes', 'One-time yearly payment — not monthly recurring'],
    [10, 'Name', 'contains', 'To ManuEstrada',                    '', '', 'Manuela Estrada - Legal Fees',                  'Yes', 'TLMND',        '', 'Yes', ''],
    [10, 'Name', 'contains', 'To LynnNguyen',                     '', '', 'Lynn Repayment',                                'Yes', 'TLMND',        '', 'Yes', ''],
    // NOTE: We do NOT categorize the "MANUELA VALLEJO" international wire
    // here — that was a one-off business expense (Vietnam criminal record
    // certificate fees) that happened to reference her name, not her
    // recurring consulting pay. Her recurring $1,155/mo comes through Gusto
    // (see priority-15 rules below). International wires like that Vietnam
    // one land uncategorized so you can classify them per instance.

    // ── MONEY OUT — non-recurring / one-offs ─────────────────────────────
    // NF Europe wires around $1,600 are Paris Thacko apartment maintenance
    // (monthly recurring). Higher-dollar or oddly-sized wires to NF Europe
    // are inter-entity transfers proper. Split by amount range with the
    // Paris rule at higher priority (15) so it wins the match first.
    [15, 'Name', 'contains', 'NF EUROPE HOLDINGS',                -1750, -1500, 'Paris Thacko Apt Maintenance',                 'Yes', 'TLMND',        '', 'Yes', 'Wire ~$1,600/mo to NF Europe Holdings'],
    [20, 'Name', 'contains', 'NF EUROPE HOLDINGS',                '', '', 'NF Europe',                                     'No',  'NF',         '', 'Yes', 'Non-Paris wires to NF Europe'],
    [20, 'Name', 'contains', 'NF MDECO SAS',                      '', '', 'NF MDE CO',                                     'No',  'NF',         '', 'Yes', 'Via BTG Pactual'],
    [20, 'Name', 'contains', 'ROETZEL AND ANDRESS',               '', '', 'Legal Fees - Roetzel and Andress',              'No',  'TLMND',        '', 'Yes', ''],
    [20, 'Name', 'contains', 'THE HOUSE PROJECT FOUNDATION',      '', '', 'Charitable Donation - The House Project',       'No',  'TLMND',        '', 'Yes', 'Donation coordinated by Manuela E; not a payment to her'],

    // ── FIDELITY (SnapTrade) — real cash flow ────────────────────────────
    // Consultant-specific GUSTO splits FIRST (priority 15 — before the
    // generic GUSTO catch-alls at 30). Amount ranges are how we distinguish
    // who was paid. Contractors typically come through as GUSTO CND
    // (Contractor Non-Deposit) or GUSTO ICD (Contractor Deposit), not
    // GUSTO NET — matching on plain "GUSTO" catches whichever it is.
    //   Manuela Vallejo: exactly $1,155
    //   Mint Lusk:       exactly $2,000
    //   Amanda Lusk:     variable $2,001-$10,000 (per user, "usually 2-5k,
    //                    had an 8k payment once")
    // Anything else via GUSTO NET is W-2 team payroll (~$38k and ~$13k
    // per pay run, aggregating to ~$48k/mo).
    [15, 'Name', 'contains', 'GUSTO',                             -1160, -1150,  'Consulting - Manuela Vallejo (Gusto)',        'Yes', 'TLMND',      '', 'Yes', 'Fixed $1,155 via Gusto CND/ICD/NET'],
    [15, 'Name', 'contains', 'GUSTO',                             -2025, -1975,  'Mint Lusk (Consulting)',                       'Yes', 'TLMND',      '', 'Yes', 'Fixed $2,000 via Gusto CND/ICD/NET'],
    [15, 'Name', 'contains', 'GUSTO',                             -10000, -2001, 'Consulting - Amanda Lusk',                     'Yes', 'TLMND',      '', 'Yes', 'Range $2,001-$10,000 via Gusto CND/ICD/NET'],
    // Team payroll (W-2 employees) — everything else in the GUSTO NET
    // bucket. Usually two withdrawals per pay run (~$38k and ~$13k).
    [30, 'Name', 'contains', 'GUSTO NET',                         '', '', 'Payroll (Team W-2)',                            'Yes', 'TLMND',        '', 'Yes', 'Rest of team payroll, ~$48k/mo'],
    [30, 'Name', 'contains', 'GUSTO TAX',                         '', '', 'Payroll (Employer Taxes)',                      'Yes', 'TLMND',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'GUSTO ICD',                         '', '', 'Payroll (Contractor Deposits)',                 'Yes', 'TLMND',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'GUSTO FEE',                         '', '', 'Payroll (Gusto Fees)',                          'Yes', 'TLMND',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'GUSTO CND',                         '', '', 'Payroll (Contractor Non-Deposit)',              'Yes', 'TLMND',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'NEXT INSUR',                        '', '', 'Business Insurance (Next Insurance)',           'Yes', 'TLMND',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'DIVIDEND SPAXX',                    '', '', 'Fidelity Money Market Interest',                'Yes', 'TLMND',        '', 'Yes', ''],

    // ── EXCLUDE — internal cash mgmt / would double-count ────────────────
    // Fidelity SPAXX buy/sell/reinvest — internal cash sweep, not real flow.
    [40, 'Name', 'contains', 'BUY SPAXX',                         '', '', '(Fidelity SPAXX cash mgmt)',                     '',   '',           'Yes', 'Yes', 'Excluded — internal'],
    [40, 'Name', 'contains', 'SELL SPAXX',                        '', '', '(Fidelity SPAXX cash mgmt)',                     '',   '',           'Yes', 'Yes', 'Excluded — internal'],
    [40, 'Name', 'contains', 'REI SPAXX',                         '', '', '(Fidelity SPAXX reinvestment)',                  '',   '',           'Yes', 'Yes', 'Excluded — internal'],
    // TLMND ↔ Fidelity internal transfers — both directions excluded.
    [40, 'Name', 'contains', 'FidelityTLM',                       '', '', 'TLMND ···2001 → Fidelity ···6454',                    '',   '',           'Yes', 'Yes', 'Excluded — paired w/ Fidelity CONTRIBUTION; loose match catches "FidelityTLM" and "FidelityTLMND"'],
    [40, 'Name', 'contains', 'CONTRIBUTION — DIRECT DEPOSIT TLMND','', '', 'TLMND ···2001 → Fidelity ···6454',                    '',   '',           'Yes', 'Yes', 'Excluded — internal'],
    // Fidelity clearing broker (NFS = National Financial Services) returning
    // money to TLMND — internal move, exclude to avoid inflating income.
    [40, 'Name', 'contains', 'NATIONAL FINANCIAL SERVICES',       '', '', 'Fidelity ···6454 → TLMND ···2001',                '',   '',           'Yes', 'Yes', 'Excluded — Fidelity NFS book credit, internal move back to TLMND (Plaid side of pair)'],
    // Fidelity side of the same outbound wire back to TLMND checking —
    // pairs with the NFS book credit above (Plaid sees the credit, we
    // exclude that; SnapTrade sees the withdrawal, we exclude that too).
    [40, 'Name', 'contains', 'WITHDRAWAL — WIRE TRANSFER',        '', '', 'Fidelity ···6454 → TLMND ···2001',                '',   '',           'Yes', 'Yes', 'Excluded — Fidelity wire out to TLMND Chase (SnapTrade side of pair, funds payroll)'],

    // ── Bank noise ───────────────────────────────────────────────────────
    [50, 'Name', 'contains', 'SERVICE CHARGES FOR THE MONTH',     '', '', 'Bank Fees',                                     'Yes', 'TLMND',        '', 'Yes', ''],
    [50, 'Name', 'contains', 'ACCOUNT ANALYSIS SETTLEMENT',       '', '', 'Bank Fees',                                     'Yes', 'TLMND',        '', 'Yes', ''],

    // ── FIDELITY (SnapTrade) — CATCH-ALLS for non-SPAXX, non-Gusto activity
    // These match at priority 200 so specific rules above always win. They
    // group unfamiliar Fidelity activity into sensible buckets rather than
    // leaving them uncategorized. Names come from _tlmndFetchSnapTradeRecords
    // as "<ACTTYPE> [<symbol>] [— description]" so starts_with is reliable.
    [200, 'Name', 'starts_with', 'BUY ',                          '', '', '(Fidelity Investment Purchase)',           '',    '',      'Yes', 'Yes', 'Excluded — non-SPAXX buy (position change, not cash flow)'],
    [200, 'Name', 'starts_with', 'SELL ',                         '', '', '(Fidelity Investment Sale)',               '',    '',      'Yes', 'Yes', 'Excluded — non-SPAXX sell (position change, not cash flow)'],
    [200, 'Name', 'starts_with', 'DIVIDEND',                      '', '', 'Fidelity Investment Income (Dividends)',   'Yes', 'TLMND', '', 'Yes', 'Non-SPAXX dividends'],
    [200, 'Name', 'starts_with', 'INTEREST',                      '', '', 'Fidelity Investment Income (Interest)',    'Yes', 'TLMND', '', 'Yes', ''],
    [200, 'Name', 'starts_with', 'WITHDRAWAL',                    '', '', 'Fidelity Other Withdrawal',                'No',  'TLMND', '', 'Yes', 'Withdrawals not matched by Gusto/Next Insur'],
    [200, 'Name', 'starts_with', 'DEPOSIT',                       '', '', 'Fidelity Other Deposit',                   'No',  'TLMND', '', 'Yes', 'Deposits not matched by Contribution rule'],
    [200, 'Name', 'starts_with', 'TRANSFER',                      '', '', '(Fidelity Internal Transfer)',             '',    '',      'Yes', 'Yes', 'Excluded — internal Fidelity move'],
    [200, 'Name', 'starts_with', 'FEE',                           '', '', 'Fidelity Account Fees',                    'No',  'TLMND', '', 'Yes', ''],
    [200, 'Name', 'starts_with', 'TAX',                           '', '', 'Fidelity Tax Withholding',                 'No',  'TLMND', '', 'Yes', ''],
    [200, 'Name', 'starts_with', 'CONTRIBUTION',                  '', '', '(Fidelity Contribution — Other)',          '',    '',      'Yes', 'Yes', 'Excluded — non-TLMND contribution'],
    [200, 'Name', 'starts_with', 'REI ',                          '', '', '(Fidelity Reinvestment)',                  '',    '',      'Yes', 'Yes', 'Excluded — reinvested dividends'],

    // ── PLAID catch-alls — anything not matched by a specific rule above
    // gets a generic bucket so it lands in Non-Recurring but stays visible.
    // These match at priority 200 so the specific vendor/counterparty rules
    // (Solaris, Ellison, UHC, Divvy, etc.) always win.
    [200, 'Name', 'contains',    'DOMESTIC WIRE TRANSFER',        '', '', 'Other Wire Transfer (Domestic)',           'No',  'TLMND', '', 'Yes', 'Catch-all — add a specific rule if recurring'],
    [200, 'Name', 'contains',    'INTERNATIONAL WIRE',            '', '', 'Other Wire Transfer (International)',      'No',  'TLMND', '', 'Yes', 'Catch-all — add a specific rule if recurring'],
    [200, 'Name', 'contains',    'BOOK TRANSFER',                 '', '', 'Other Book Transfer',                      'No',  'TLMND', '', 'Yes', 'Catch-all'],
    [200, 'Name', 'contains',    'Online ACH Payment',            '', '', 'Other ACH Payment',                        'No',  'TLMND', '', 'Yes', 'Catch-all'],
    [200, 'Name', 'contains',    'ORIG CO NAME:',                 '', '', 'Other ACH (Deposit or Debit)',             'No',  'TLMND', '', 'Yes', 'Catch-all — inbound or outbound ACH not matched'],
    [200, 'Name', 'contains',    'REMOTE ONLINE DEPOSIT',         '', '', 'Remote Check Deposit',                     'No',  'TLMND', '', 'Yes', 'Deposited check via app/scanner'],
    [200, 'Name', 'contains',    'DEPOSIT ID NUMBER',             '', '', 'Branch Deposit',                           'No',  'TLMND', '', 'Yes', 'Cash/check deposit at branch'],
    [200, 'Name', 'starts_with', 'CHECK #',                       '', '', 'Check Payment',                            'No',  'TLMND', '', 'Yes', 'Hand-written check debited from the account'],

    // ── Inter-account journal transfers ──────────────────────────────────
    // NF USA CA ↔ TLMND internal journals: paired with the Ellison deposit
    // we already count as income. Excluding both sides prevents triple-count
    // (deposit + inbound + outbound = 3× the actual income).
    [90, 'Name', 'contains', 'Online Transfer from CHK ...2086',  '', '', 'NF USA CA ···2086 → TLMND ···2001',              '',   '',           'Yes', 'Yes', 'Excluded — TLMND-side inbound (paired with Ellison deposit already counted)'],
    [90, 'Name', 'contains', 'Online Transfer to CHK ...2001',    '', '', 'NF USA CA ···2086 → TLMND ···2001',              '',   '',           'Yes', 'Yes', 'Excluded — NF USA CA-side outbound (mirror of above)'],
    // The reverse direction (TLMND → NF USA CA) — added for completeness if
    // she moves money back that way. Same pair concept, other direction.
    [90, 'Name', 'contains', 'Online Transfer to CHK ...2086',    '', '', 'TLMND ···2001 → NF USA CA ···2086',              '',   '',           'Yes', 'Yes', 'Excluded — TLMND-side outbound to NF USA CA'],
    [90, 'Name', 'contains', 'Online Transfer from CHK ...2001',  '', '', 'TLMND ···2001 → NF USA CA ···2086',              '',   '',           'Yes', 'Yes', 'Excluded — NF USA CA-side inbound (mirror of above)'],
    // Blue Panda Family ···8686 → TLMND: real inter-entity funding, COUNT it.
    // Match both by account-mask (internal Chase transfer) AND by name
    // substring in case Blue Panda money arrives via a different mechanism
    // (wire, ACH) with a different name format.
    [15, 'Name', 'contains', 'BLUE PANDA',                        '', '', 'Blue Panda Family',               'No',  'TLMND',      '', 'Yes', 'Any Blue Panda inbound — catches wires/ACH by name'],
    [90, 'Name', 'contains', 'Online Transfer from CHK ...8686',  '', '', 'Blue Panda Family',               'No',  'TLMND',      '', 'Yes', 'Blue Panda Family ···8686 → TLMND (inbound Chase internal transfer)'],
    [90, 'Name', 'contains', 'Online Transfer to CHK ...8686',    '', '', 'Blue Panda Family',               'No',  'TLMND',      '', 'Yes', 'TLMND → Blue Panda Family ···8686 (outbound Chase internal transfer)'],
    // TLMND ↔ NF USA TX: real inter-entity movement, COUNT it. Match by
    // name substring first (catches wires/ACH) then by the internal
    // Chase transfer format as a fallback.
    [15, 'Name', 'contains', 'NF USA TX',                         '', '', 'NF Texas',                    'No',  'TLMND',      '', 'Yes', 'Any NF USA TX movement — catches wires/ACH by name'],
    [90, 'Name', 'contains', 'Online Transfer to CHK ...5155',    '', '', 'NF Texas',                    'No',  'TLMND',      '', 'Yes', 'TLMND ↔ NF USA TX ···5155 (Chase internal transfer)'],
    // NF6 Tiger Capital ···5319 — Chase internal transfer, real inter-entity flow.
    [90, 'Name', 'contains', 'Online Transfer from CHK ...5319',  '', '', 'NF6 Tiger Capital',           'No',  'TLMND',      '', 'Yes', 'NF6 Tiger Capital ···5319 → TLMND (inbound)'],
    [90, 'Name', 'contains', 'Online Transfer to CHK ...5319',    '', '', 'NF6 Tiger Capital',           'No',  'TLMND',      '', 'Yes', 'TLMND → NF6 Tiger Capital ···5319 (outbound)']
  ];

  // Pad every rule to the full header width so setValues stays rectangular.
  rules = rules.map(function(r) {
    return r.length < TLMND_RULES_HEADERS.length
      ? r.concat(new Array(TLMND_RULES_HEADERS.length - r.length).fill(''))
      : r;
  });

  // Post-process: mark catch-all rules (priority 200) with the correct
  // source filter so a SnapTrade pattern like starts_with 'DEPOSIT' can't
  // accidentally match a Plaid transaction that starts with the same word.
  var snapTradePatterns = ['BUY ','SELL ','DIVIDEND','INTEREST','WITHDRAWAL','DEPOSIT','TRANSFER','FEE','TAX','CONTRIBUTION','REI '];
  var plaidPatterns     = ['DOMESTIC WIRE TRANSFER','INTERNATIONAL WIRE','BOOK TRANSFER','Online ACH Payment','ORIG CO NAME:','REMOTE ONLINE DEPOSIT','DEPOSIT ID NUMBER','CHECK #'];
  rules.forEach(function(r) {
    if (r[0] !== 200) return;
    if (snapTradePatterns.indexOf(String(r[3])) >= 0) r[12] = 'SnapTrade';
    else if (plaidPatterns.indexOf(String(r[3])) >= 0) r[12] = 'Plaid';
  });

  sheet.getRange(2, 1, rules.length, TLMND_RULES_HEADERS.length).setValues(rules);
  ui.alert('Seeded ' + rules.length + ' rules into TLMND_CATEGORY_RULES.\n\n' +
           'Run Tracker → Apply TLMND Rules to categorize existing transactions, ' +
           'or run a full Sync — rules are applied automatically after every sync.');
}

function _tlmndLoadRules() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TLMND_CATEGORY_RULES');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, TLMND_RULES_HEADERS.length).getValues();
  var rules = [];
  vals.forEach(function(r) {
    if (String(r[10]).toLowerCase() !== 'yes') return;   // Enabled column
    if (!r[3]) return;                                    // no Pattern → skip
    rules.push({
      priority:        Number(r[0]) || 999,
      field:           String(r[1] || 'Name'),
      matchType:       String(r[2] || 'contains').toLowerCase(),
      pattern:         String(r[3]),
      amtMin:          r[4] === '' || r[4] == null ? null : Number(r[4]),
      amtMax:          r[5] === '' || r[5] == null ? null : Number(r[5]),
      category:        String(r[6] || ''),
      recurring:       String(r[7] || ''),
      entityTag:       String(r[8] || ''),
      exclude:         String(r[9]).toLowerCase() === 'yes',
      sourceFilter:    String(r[12] || '').trim(),
      expectedMonthly: r[13] === '' || r[13] == null ? null : Number(r[13])
    });
  });
  rules.sort(function(a, b) { return a.priority - b.priority; });
  return rules;
}

function _tlmndMatchOne(fieldValue, matchType, pattern) {
  var haystack = String(fieldValue || '').toLowerCase();
  var needle   = String(pattern || '').toLowerCase();
  switch (matchType) {
    case 'equals':      return haystack === needle;
    case 'starts_with': return haystack.indexOf(needle) === 0;
    case 'regex':
      try { return new RegExp(pattern, 'i').test(fieldValue); } catch(e) { return false; }
    case 'contains':
    default:            return haystack.indexOf(needle) >= 0;
  }
}

// Apply rules to TLMND_TRANSACTIONS. Preserves manually-set Category values
// (any Category that doesn't match some rule's Category is treated as manual).
function applyTLMNDRules() { _requireEditor_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TLMND_TRANSACTIONS');
  if (!sheet || sheet.getLastRow() < 2) return { success: false, error: 'No transactions to categorize.' };

  var rules = _tlmndLoadRules();
  if (!rules.length) return { success: false, error: 'No enabled rules in TLMND_CATEGORY_RULES. Run "Seed TLMND Rules" first.' };

  // Header index map for future-proofing.
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  function ci(name) { return headers.indexOf(name); }
  var idxName   = ci('Name');
  var idxMerch  = ci('Merchant');
  var idxAccount= ci('Account');
  var idxSource = ci('Source');
  var idxAmount = ci('Amount USD');
  var idxCat    = ci('Category');
  var idxRec    = ci('Recurring');
  var idxEnt    = ci('Entity Tag');
  var idxNotes  = ci('Notes');
  if (idxCat < 0) return { success: false, error: 'Category column missing.' };

  // Build min-priority-per-category lookup. Any category name whose
  // best-priority rule is >= CATCH_ALL_THRESHOLD is treated as a
  // catch-all bucket (e.g. "Other ACH Payment" from priority 200) and
  // eligible for re-evaluation on re-apply — so a newly-added
  // more-specific rule (e.g. priority 40 "FidelityTLM" →
  // "TLMND ···2001 → Fidelity ···6454") can win over the stale bucket.
  var CATCH_ALL_THRESHOLD = 100;
  var ruleCategories = {};              // category → true (any rule uses it)
  var catMinPriority = {};              // category → lowest priority observed
  rules.forEach(function(r) {
    if (!r.category) return;
    ruleCategories[r.category] = true;
    if (catMinPriority[r.category] == null || r.priority < catMinPriority[r.category]) {
      catMinPriority[r.category] = r.priority;
    }
  });
  function _isCatchAllCategory(cat) {
    var p = catMinPriority[cat];
    return p != null && p >= CATCH_ALL_THRESHOLD;
  }

  var last = sheet.getLastRow();
  var range = sheet.getRange(2, 1, last - 1, headers.length);
  var vals  = range.getValues();

  var categorized = 0, skippedManual = 0, uncategorized = 0, excluded = 0;

  vals.forEach(function(row, i) {
    var existing = String(row[idxCat] || '').trim();
    // If Category is already set AND it's a real manual entry (either
    // matches a low-priority rule, or matches no rule at all), preserve
    // it. Rule engine only runs on blank cells + catch-all bucketed rows.
    // This preserves manually-tagged rows (e.g. "MacDonald Loan Repayment"
    // on the $8,333 branch deposit) while still letting new more-specific
    // rules override stale catch-all assignments like "Other ACH Payment".
    if (existing && !_isCatchAllCategory(existing)) {
      for (var mk = 0; mk < rules.length; mk++) {
        if (rules[mk].category === existing) {
          if (rules[mk].recurring) row[idxRec] = rules[mk].recurring;
          if (rules[mk].entityTag) row[idxEnt] = rules[mk].entityTag;
          var mn = String(row[idxNotes] || '').replace(/^\[EXCLUDED\]\s*/, '');
          if (rules[mk].exclude) { mn = '[EXCLUDED] ' + mn; excluded++; }
          row[idxNotes] = mn.trim();
          break;
        }
      }
      skippedManual++;
      return;
    }

    var matched = null;
    for (var k = 0; k < rules.length; k++) {
      var r = rules[k];
      // Source filter — restrict a rule to Plaid or SnapTrade only.
      // Prevents e.g. the Fidelity SnapTrade "starts_with DEPOSIT" rule
      // from wrongly matching a Plaid Chase branch deposit named
      // "DEPOSIT ID NUMBER 553083".
      if (r.sourceFilter && r.sourceFilter !== String(row[idxSource] || '')) continue;
      var fieldVal = r.field === 'Merchant' ? row[idxMerch]
                   : r.field === 'Account'  ? row[idxAccount]
                   :                          row[idxName];
      if (!_tlmndMatchOne(fieldVal, r.matchType, r.pattern)) continue;
      var amt = Number(row[idxAmount] || 0);
      if (r.amtMin !== null && amt < r.amtMin) continue;
      if (r.amtMax !== null && amt > r.amtMax) continue;
      matched = r; break;
    }

    if (matched) {
      row[idxCat] = matched.category;
      if (matched.recurring) row[idxRec] = matched.recurring;
      if (matched.entityTag) row[idxEnt] = matched.entityTag;
      // Always strip any prior [EXCLUDED] marker before re-applying so
      // rows flip cleanly when a rule's exclude=Yes becomes exclude=blank.
      var n = String(row[idxNotes] || '').replace(/^\[EXCLUDED\]\s*/, '');
      if (matched.exclude) {
        n = '[EXCLUDED] ' + n;
        excluded++;
      }
      row[idxNotes] = n.trim();
      categorized++;
    } else {
      uncategorized++;
    }
  });

  range.setValues(vals);
  return {
    success: true,
    total: vals.length,
    categorized: categorized,
    excluded: excluded,
    skippedManual: skippedManual,
    uncategorized: uncategorized
  };
}

function applyTLMNDRulesMenu() {
  var ui = SpreadsheetApp.getUi();
  var r  = applyTLMNDRules();
  if (!r.success) { ui.alert('Failed', r.error, ui.ButtonSet.OK); return; }
  ui.alert('TLMND Rules Applied',
    'Total rows:                  ' + r.total +
    '\nCategorized (this run): ' + r.categorized +
    '\n  of which [EXCLUDED]: ' + r.excluded +
    '\nSkipped (manual override): ' + r.skippedManual +
    '\nStill uncategorized:      ' + r.uncategorized +
    (r.uncategorized > 0 ? '\n\nFor the uncategorized rows, either add a new rule in TLMND_CATEGORY_RULES ' +
     'or set the Category column manually — manual values are preserved on re-sync.' : ''),
    ui.ButtonSet.OK);
}

// ============================================================================
// PHASE 3 — DASHBOARD DATA ENDPOINT
//
// Called from the web app's TLMND Cash Flow tab. Reads TLMND_TRANSACTIONS,
// applies filters, and returns everything the dashboard needs to render in
// one round-trip:
//   - kpis: current-month In/Out/NetRecurring/NetAll + T3M/T6M/T12M averages
//   - monthlySeries: array of {ym, in, out, netRec, netAll} for the chart
//   - categoryMatrix: array of {category, entityTag, recurring, months:{ym→amt}, total}
//   - transactions: filtered raw rows for drill-down
// ============================================================================
function getTLMNDCashFlowData(opts) {
  opts = opts || {};
  var months        = Number(opts.months) || 6;             // window to render
  var entityFilter  = opts.entityTag || '';                 // '' = all
  // NOTE: excluded rows (NF USA CA deposits, Fidelity SPAXX cash mgmt,
  // internal transfer mirrors) are ALWAYS returned. Frontend routes them
  // into a dedicated "Reference — Related Account Activity" section that
  // does not roll into the main totals. This makes the passthrough
  // account visible without inflating TLMND's cash-flow numbers.

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TLMND_TRANSACTIONS');
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, error: 'TLMND_TRANSACTIONS sheet is empty. Run Sync TLMND Cash Flow first.' };
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  function ci(name) { return headers.indexOf(name); }
  var iDate    = ci('Date'),   iAccount = ci('Account'), iName = ci('Name');
  var iAmount  = ci('Amount USD'), iCat  = ci('Category'), iRec = ci('Recurring');
  var iEnt     = ci('Entity Tag'), iNotes = ci('Notes');

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  // Build the month window: last N months including current.
  var now = new Date();
  var monthKeys = [];
  for (var m = months - 1; m >= 0; m--) {
    var d = new Date(now.getFullYear(), now.getMonth() - m, 1);
    monthKeys.push(d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2));
  }
  var earliestYm = monthKeys[0];

  // Filter + normalize rows.
  var filtered = [];
  rows.forEach(function(r) {
    if (!r[iDate]) return;
    var d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]);
    if (isNaN(d.getTime())) return;
    var ym = d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2);
    if (ym < earliestYm) return;
    var notes = String(r[iNotes] || '');
    var isExcluded = notes.indexOf('[EXCLUDED]') >= 0;
    if (entityFilter && String(r[iEnt] || '') !== entityFilter) return;
    filtered.push({
      ym:        ym,
      date:      Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      account:   String(r[iAccount] || ''),
      name:      String(r[iName] || ''),
      amount:    Number(r[iAmount] || 0),
      category:  String(r[iCat] || '(Uncategorized)') || '(Uncategorized)',
      recurring: String(r[iRec] || '') === 'Yes',
      entityTag: String(r[iEnt] || ''),
      excluded:  isExcluded
    });
  });

  // Aggregate per (category, ym). Excluded transactions (Fidelity SPAXX
  // cash mgmt, TLMND↔Fidelity funding pairs, NF USA CA↔TLMND internal
  // journals) are filtered OUT entirely — the top-of-tab totals only
  // reflect actual TLMND income and expense flows, not internal mirrors
  // that would double-count.
  var catMap = {};
  var series = {};
  monthKeys.forEach(function(ym) { series[ym] = { ym: ym, in: 0, out: 0, netRec: 0, netAll: 0 }; });

  filtered.forEach(function(t) {
    if (t.excluded) return;
    if (!catMap[t.category]) {
      catMap[t.category] = { category: t.category, recurring: t.recurring, entityTag: t.entityTag, months: {}, total: 0 };
    }
    catMap[t.category].months[t.ym] = (catMap[t.category].months[t.ym] || 0) + t.amount;
    catMap[t.category].total += t.amount;
    if (t.recurring) catMap[t.category].recurring = true;

    var s = series[t.ym];
    if (!s) return;
    if (t.amount >= 0) s.in += t.amount; else s.out += t.amount;
    s.netAll += t.amount;
    if (t.recurring) s.netRec += t.amount;
  });

  var categoryMatrix = Object.keys(catMap).map(function(k) { return catMap[k]; })
    .sort(function(a, b) {
      if (a.recurring !== b.recurring) return a.recurring ? -1 : 1;
      return Math.abs(b.total) - Math.abs(a.total);
    });

  var monthlySeries = monthKeys.map(function(ym) { return series[ym]; });

  // KPIs: current month + trailing averages.
  var curYm = monthKeys[monthKeys.length - 1];
  var curr  = series[curYm] || { in: 0, out: 0, netRec: 0, netAll: 0 };
  function avg(nBack) {
    var total = 0, count = 0;
    for (var i = 0; i < nBack && i < monthlySeries.length; i++) {
      var s = monthlySeries[monthlySeries.length - 1 - i];
      total += s.netAll; count++;
    }
    return count ? total / count : 0;
  }
  var kpis = {
    ym:          curYm,
    monthIn:     curr.in,
    monthOut:    curr.out,
    monthNetRec: curr.netRec,
    monthNetAll: curr.netAll,
    avgT3M:      avg(3),
    avgT6M:      avg(6),
    avgT12M:     avg(12)
  };

  // Distinct entity tags for the filter dropdown.
  var entTagSet = {};
  rows.forEach(function(r) { var e = String(r[iEnt] || ''); if (e) entTagSet[e] = true; });

  // Category-level Expected Monthly overrides from the rules sheet — sent
  // to the client so the forecast + on-track panel can pin specific
  // categories to a known monthly rate (e.g. MacDonald = $2,083/mo)
  // instead of relying on the historical average.
  var overrides = {};
  try {
    var _r = _tlmndLoadRules();
    _r.forEach(function(rule) {
      if (rule.category && rule.expectedMonthly != null && !isNaN(rule.expectedMonthly)) {
        overrides[rule.category] = rule.expectedMonthly;
      }
    });
  } catch(e) {}

  return {
    success: true,
    kpis: kpis,
    monthlySeries: monthlySeries,
    categoryMatrix: categoryMatrix,
    transactions: filtered,
    monthKeys: monthKeys,
    entityTags: Object.keys(entTagSet).sort(),
    expectedOverrides: overrides,
    generatedAt: new Date().toISOString()
  };
}

// Client-callable sync trigger for the dashboard's Refresh button.
function refreshTLMNDCashFlow() {
  return syncTLMNDCashFlow();
}

// Read current balances for TLMND-tracked accounts from the Assets sheet.
// The Assets sheet is refreshed by syncPlaidAccounts / syncSnapTradeAccounts
// (both scheduled). Returns one entry per configured Plaid + SnapTrade
// account with its label, USD balance, and the freshest Last Updated
// timestamp across them all so the dashboard can show "as of ...".
function getTLMNDAccountBalances() {
  var cfg = getTLMNDConfig_();
  if (!cfg) return { success: false, error: 'TLMND config not initialized.' };

  var sheet = getSheet_('ASSETS');
  var rows  = sheet.getDataRange().getValues();
  var hdr   = rows[0];
  function ci(name) { return hdr.indexOf(name); }
  var iPlaid = ci('Plaid Account ID');
  var iSnap  = ci('SnapTrade ID');
  var iName  = ci('Name');
  var iUsd   = ci('USD Value');
  var iLU    = ci('Last Updated');

  var balances = [];
  var latestUpdate = null;

  function pushFromRow(i, label, source, role) {
    var val = Number(rows[i][iUsd]) || 0;
    var lu  = rows[i][iLU];
    balances.push({ label: label, value: val, source: source, role: role || 'primary' });
    if (lu && (!latestUpdate || new Date(lu) > new Date(latestUpdate))) latestUpdate = lu;
  }

  // Belt-and-suspenders: hardcode known passthrough accountIds so even if
  // the stored config predates the role field, NF USA CA is never surfaced
  // as a reserve balance. Its income already flows into TLMND's recurring
  // income baseline via categorization rules.
  var HARDCODED_PASSTHROUGH = { 'Bv9mLEzzVLhwxM5v66XRSX1P3Qr989HvXAOz1': true };

  (cfg.plaidAccounts || []).forEach(function(a) {
    var isPassthrough = a.role === 'passthrough' || HARDCODED_PASSTHROUGH[a.accountId];
    if (isPassthrough) return;  // don't return passthrough accounts at all
    for (var i = 1; i < rows.length; i++) {
      if (iPlaid >= 0 && String(rows[i][iPlaid]) === a.accountId) {
        pushFromRow(i, a.label, 'Plaid', 'primary');
        return;
      }
    }
    balances.push({ label: a.label, value: 0, source: 'Plaid', role: 'primary', notFound: true });
  });

  (cfg.snapTradeAccounts || []).forEach(function(a) {
    for (var i = 1; i < rows.length; i++) {
      if (iSnap >= 0 && String(rows[i][iSnap]) === a.accountId) {
        pushFromRow(i, a.label, 'SnapTrade', a.role);
        return;
      }
    }
    balances.push({ label: a.label, value: 0, source: 'SnapTrade', role: a.role || 'primary', notFound: true });
  });

  return {
    success: true,
    balances: balances,
    latestUpdate: latestUpdate ? new Date(latestUpdate).toISOString() : null
  };
}

// Nuclear option — deletes every data row in TLMND_TRANSACTIONS then
// runs a fresh sync. Useful after changing the sync keying (e.g. the
// pending_transaction_id fix) so old duplicates get flushed. Loses any
// manual Category/Notes edits since rules are re-applied after resync.
function clearAndResyncTLMND() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Clear & Resync TLMND',
    'This will DELETE every row in TLMND_TRANSACTIONS, then pull a fresh copy from Plaid/SnapTrade. ' +
    'Any manual Category / Notes edits will be lost (rules re-apply automatically after the pull).\n\n' +
    'Use this to flush pending/posted duplicates from earlier syncs. Continue?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TLMND_TRANSACTIONS');
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }

  var r = syncTLMNDCashFlow();
  var msg = 'Plaid transactions:       ' + (r.plaidCount || 0) +
            '\nSnapTrade activities:  ' + (r.snapTradeCount || 0) +
            '\nNew rows:                     ' + (r.newRows || 0) +
            '\nCategorized:                 ' + (r.categorized || 0) +
            '\n  of which [EXCLUDED]: ' + (r.excluded || 0) +
            '\nStill uncategorized:      ' + (r.uncategorized || 0) +
            '\nErrors:                           ' + (r.errorCount || 0);
  if (r.errors && r.errors.length) msg += '\n\nErrors:\n  • ' + r.errors.slice(0, 5).join('\n  • ');
  ui.alert('Clear & Resync complete', msg, ui.ButtonSet.OK);
}

// ============================================================================
// WEEKLY PDF REPORT — sent every Monday to the same distribution as the
// net worth weekly email (script property WEEKLY_PDF_RECIPIENT).
//
// Content is a static 2-page snapshot of the TLMND dashboard's key numbers
// (Funding Planner is intentionally excluded per user). Layout:
//   Page 1: hero net cash flow, top movers, comparison table, burn/forecast
//   Page 2: full category matrix (recurring / non-recurring / inter-entity)
// ============================================================================

// Server-side month label helper for PDF rendering.
function _tlmndMonthLabelSvr_(ym) {
  var names = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  if (!ym) return '';
  var p = String(ym).split('-');
  return (names[Number(p[1]) - 1] || '?') + ' ' + p[0];
}
function _tlmndFmtSvr_(v) {
  if (Math.abs(v || 0) < 0.5) return '$0';
  return (v < 0 ? '−$' : '+$') + Math.abs(Math.round(v)).toLocaleString();
}
function _tlmndFmtPos_(v) { return '$' + Math.abs(Math.round(v || 0)).toLocaleString(); }

// Server-side weekly aggregation helpers — mirrors the frontend ones so
// the PDF and dashboard produce identical numbers.
function _tlmndWeekStartMonday_(dateStr) {
  var d = new Date(dateStr + 'T12:00:00');
  var day = d.getDay();
  var diff = (day === 0) ? -6 : 1 - day;
  var m = new Date(d);
  m.setDate(d.getDate() + diff);
  return m.getFullYear() + '-' + ('0'+(m.getMonth()+1)).slice(-2) + '-' + ('0'+m.getDate()).slice(-2);
}
function _tlmndWeekLabel_(mondayStr) {
  var s = new Date(mondayStr + 'T12:00:00');
  var e = new Date(s); e.setDate(s.getDate() + 6);
  var mos = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  if (s.getMonth() === e.getMonth()) return mos[s.getMonth()] + ' ' + s.getDate() + '-' + e.getDate();
  return mos[s.getMonth()] + ' ' + s.getDate() + ' - ' + mos[e.getMonth()] + ' ' + e.getDate();
}
function _tlmndBuildWeekly_(txns) {
  var by = {};
  (txns || []).forEach(function(t) {
    if (t.excluded) return;
    var wk = _tlmndWeekStartMonday_(t.date);
    if (!by[wk]) by[wk] = { weekStart: wk, in: 0, out: 0, netAll: 0 };
    if (t.amount >= 0) by[wk].in += t.amount; else by[wk].out += t.amount;
    by[wk].netAll += t.amount;
  });
  return Object.keys(by).sort().map(function(k) { return by[k]; });
}
function _tlmndLastCompleteWeek_(weeks) {
  if (!weeks || !weeks.length) return null;
  var todayMon = _tlmndWeekStartMonday_((new Date()).toISOString().substring(0, 10));
  for (var i = weeks.length - 1; i >= 0; i--) {
    if (weeks[i].weekStart < todayMon) return weeks[i];
  }
  return weeks[0];
}
function _tlmndMonthTD_(txns, ym, dayCap) {
  var out = { in: 0, out: 0, netAll: 0 };
  (txns || []).forEach(function(t) {
    if (t.excluded) return;
    if (!t.date || t.date.substring(0, 7) !== ym) return;
    var day = Number(t.date.substring(8, 10));
    if (dayCap != null && day > dayCap) return;
    if (t.amount >= 0) out.in += t.amount; else out.out += t.amount;
    out.netAll += t.amount;
  });
  return out;
}
// Rolling N weeks ending at endIdx — mirror of the frontend helper.
function _tlmndRollingWeeks_(weeks, endIdx, n) {
  var out = { in: 0, out: 0, netAll: 0, startWeek: null, endWeek: null, weeksCounted: 0 };
  if (!weeks || !weeks.length || endIdx < 0) return out;
  var startIdx = Math.max(0, endIdx - n + 1);
  for (var i = startIdx; i <= endIdx; i++) {
    var w = weeks[i];
    out.in += w.in; out.out += w.out; out.netAll += w.netAll;
    out.weeksCounted++;
  }
  out.startWeek = weeks[startIdx].weekStart;
  out.endWeek   = weeks[endIdx].weekStart;
  return out;
}
// Rolling 30-day sum helpers, mirror of frontend.
function _tlmndSumDateRange_(txns, startISO, endISO) {
  var out = { in: 0, out: 0, netAll: 0, count: 0 };
  (txns || []).forEach(function(t) {
    if (t.excluded || !t.date) return;
    if (t.date < startISO || t.date > endISO) return;
    if (t.amount >= 0) out.in += t.amount; else out.out += t.amount;
    out.netAll += t.amount;
    out.count++;
  });
  return out;
}
function _tlmndDateOffset_(baseISO, offsetDays) {
  var d = new Date(baseISO + 'T12:00:00');
  d.setDate(d.getDate() + offsetDays);
  return d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2) + '-' + ('0'+d.getDate()).slice(-2);
}
function _tlmndDateShort_(iso) {
  if (!iso) return '';
  var d = new Date(iso + 'T12:00:00');
  var mos = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return mos[d.getMonth()] + ' ' + d.getDate();
}
// Next 30 days expected — recurring cats. Priority is an Expected Monthly
// Override from the rules sheet (per-category). If none set, uses the
// average of the last 3 completed months (divisor is always the full
// window so lump payments don't inflate the monthly rate).
function _tlmndForecast_(mat, months, overrides) {
  overrides = overrides || {};
  var now = new Date();
  var curYm = now.getFullYear() + '-' + ('0'+(now.getMonth()+1)).slice(-2);
  var recentYms = months.filter(function(m) { return m.ym !== curYm; })
    .slice(-3).map(function(m) { return m.ym; });
  var windowSize = recentYms.length || 1;
  function isIE(c) { return /blue panda|nf europe|nf mde co|nf texas|nf6 tiger|inter-entity/i.test(c); }
  var inItems = [], outItems = [];
  (mat || []).forEach(function(c) {
    if (!c.recurring) return;
    if (isIE(c.category)) return;
    var expected;
    if (overrides[c.category] != null) {
      expected = Number(overrides[c.category]);
    } else {
      var total = 0;
      recentYms.forEach(function(ym) { total += (c.months[ym] || 0); });
      expected = total / windowSize;
    }
    if (Math.abs(expected) < 1) return;
    if (expected >= 0) inItems.push({ name: c.category, val: expected });
    else outItems.push({ name: c.category, val: expected });
  });
  inItems.sort(function(a, b) { return b.val - a.val; });
  outItems.sort(function(a, b) { return a.val - b.val; });
  var totalIn = inItems.reduce(function(s, x) { return s + x.val; }, 0);
  var totalOut = outItems.reduce(function(s, x) { return s + x.val; }, 0);
  return { inItems: inItems, outItems: outItems, totalIn: totalIn, totalOut: totalOut, net: totalIn + totalOut };
}

function _tlmndEsc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function(c) {
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];
  });
}

function _tlmndBuildWeeklyPdfHtml_() {
  var d = getTLMNDCashFlowData({ months: 6 });
  if (!d || !d.success) return null;

  var kpis    = d.kpis;
  var months  = d.monthKeys || [];
  var mat     = d.categoryMatrix || [];
  var series  = d.monthlySeries || [];

  // Reference month for hero / top movers / comparison table = the LAST
  // COMPLETED month (not the in-progress current month). Otherwise a mid-
  // month partial like Sep 1-17 gets compared to full-month averages and
  // looks either great or terrible for the wrong reason.
  var _srvNow = new Date();
  var _srvCurYm = _srvNow.getFullYear() + '-' + ('0'+(_srvNow.getMonth()+1)).slice(-2);
  var refMonth = null;
  for (var _si = series.length - 1; _si >= 0; _si--) {
    if (series[_si].ym !== _srvCurYm) { refMonth = series[_si]; break; }
  }
  if (!refMonth) refMonth = series[series.length - 1] || { ym: _srvCurYm, in: 0, out: 0, netAll: 0 };
  var mtdEntry = (series.length && series[series.length - 1].ym === _srvCurYm && series[series.length - 1] !== refMonth) ? series[series.length - 1] : null;

  var monthLabel = _tlmndMonthLabelSvr_(refMonth.ym) + (mtdEntry ? ' (last complete)' : '');
  var reportDate = Utilities.formatDate(new Date(), 'America/New_York', 'MMMM d, yyyy');

  // Current balances for primary accounts (TLMND checking + Fidelity).
  // Passthrough accounts (NF USA CA) are already filtered out by
  // getTLMNDAccountBalances via the hardcoded passthrough list.
  var balancesRes = null;
  try { balancesRes = getTLMNDAccountBalances(); } catch(e) { balancesRes = null; }
  var balList = (balancesRes && balancesRes.success && balancesRes.balances) ? balancesRes.balances : [];
  var balTotal = 0;
  balList.forEach(function(b) { balTotal += Number(b.value) || 0; });
  var balAsOf = balancesRes && balancesRes.latestUpdate
    ? Utilities.formatDate(new Date(balancesRes.latestUpdate), 'America/New_York', 'MMM d, yyyy \'at\' h:mm a')
    : 'not yet synced';

  // === ROLLING MONTHLY HERO ===
  // Reference = LAST 30 DAYS (rolling), vs prior 30 days.
  var txns = d.transactions || [];
  var todayISO = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');
  var yesterdayISO = _tlmndDateOffset_(todayISO, -1);
  var start30 = _tlmndDateOffset_(todayISO, -30);
  var startPrev30 = _tlmndDateOffset_(todayISO, -60);
  var endPrev30   = _tlmndDateOffset_(todayISO, -31);
  var actCur   = _tlmndSumDateRange_(txns, start30, yesterdayISO);
  var actPrior = _tlmndSumDateRange_(txns, startPrev30, endPrev30);
  var actDelta = actCur.netAll - actPrior.netAll;
  var deltaTxt;
  if (Math.abs(actDelta) < 500) deltaTxt = 'about the same as the prior 30 days';
  else if (actDelta > 0) deltaTxt = _tlmndFmtPos_(actDelta) + ' better than prior 30 days';
  else                   deltaTxt = _tlmndFmtPos_(actDelta) + ' worse than prior 30 days';

  // === Coming Up forecast (recurring, next 30 days) ===
  var fc = _tlmndForecast_(mat, series, d.expectedOverrides || {});

  monthLabel = 'Last 30 Days · ' + _tlmndDateShort_(start30) + ' – ' + _tlmndDateShort_(yesterdayISO);
  var cur = actCur;

  // === Top movers — LAST 30 DAYS actual, grouped by CATEGORY ===
  var actualBy = {};
  txns.forEach(function(t) {
    if (t.excluded || !t.date) return;
    if (t.date < start30 || t.date > yesterdayISO) return;
    var nm = (t.category && t.category !== '(Uncategorized)') ? t.category : (t.name || '(unlabeled)');
    if (!actualBy[nm]) actualBy[nm] = { name: nm, val: 0, count: 0 };
    actualBy[nm].val += t.amount;
    actualBy[nm].count++;
  });
  var inItems = [], outItems = [];
  Object.keys(actualBy).forEach(function(k) {
    var e = actualBy[k];
    if (e.val > 0) inItems.push(e);
    else if (e.val < 0) outItems.push(e);
  });
  inItems.sort(function(a, b) { return b.val - a.val; });
  outItems.sort(function(a, b) { return a.val - b.val; });
  var totalIn  = inItems.reduce(function(s, x) { return s + x.val; }, 0);
  var totalOut = outItems.reduce(function(s, x) { return s + x.val; }, 0);

  // === Comparison table ===
  // Averages use completed months only.
  var refIdx = -1;
  for (var _pi = 0; _pi < series.length; _pi++) { if (series[_pi].ym === refMonth.ym) { refIdx = _pi; break; } }
  var seriesForAvg = series.slice(0, refIdx + 1);
  function avgField(field, n) {
    var arr = seriesForAvg.slice(-n);
    if (!arr.length) return 0;
    return arr.reduce(function(s, m) { return s + (m[field] || 0); }, 0) / arr.length;
  }
  var compareRows = [
    { label: 'Money In',  last30: actCur.in,     prior30: actPrior.in,     expected: fc.totalIn,  t3: avgField('in',3),     t6: avgField('in',6)  },
    { label: 'Money Out', last30: actCur.out,    prior30: actPrior.out,    expected: fc.totalOut, t3: avgField('out',3),    t6: avgField('out',6) },
    { label: 'Net',       last30: actCur.netAll, prior30: actPrior.netAll, expected: fc.net,      t3: avgField('netAll',3), t6: avgField('netAll',6), isNet: true }
  ];

  // Burn & Forecast (recurring only, exclude current partial month).
  var now = new Date();
  var currentYm = now.getFullYear() + '-' + ('0'+(now.getMonth()+1)).slice(-2);
  var completeMonths = series.filter(function(m) { return m.ym !== currentYm; });
  function isIE(c) { return /blue panda|nf europe|nf mde co|nf texas|nf6 tiger|inter-entity/i.test(c); }
  var opIn = {}, opOut = {};
  completeMonths.forEach(function(m) { opIn[m.ym] = 0; opOut[m.ym] = 0; });
  mat.forEach(function(c) {
    if (isIE(c.category)) return;
    if (!c.recurring) return;
    Object.keys(c.months).forEach(function(ym) {
      if (opIn[ym] === undefined) return;
      var v = c.months[ym] || 0;
      if (v >= 0) opIn[ym]  += v; else opOut[ym] += v;
    });
  });
  function avgN(dict, n) {
    var vals = completeMonths.slice(-n).map(function(m) { return dict[m.ym] || 0; });
    if (!vals.length) return 0;
    return vals.reduce(function(a, b) { return a + b; }, 0) / vals.length;
  }
  var avgRecIn  = avgN(opIn, 6);
  var avgRecOut = avgN(opOut, 6);
  var netRec    = avgRecIn + avgRecOut;
  var monthlyNeed = -netRec;

  // Category matrix — group by direction like the dashboard, with subgroups.
  function _grpOf(cat) {
    var c = String(cat || '').toLowerCase();
    if (/consulting|payroll|lusk|estrada|nguyen|vallejo|mint|lynn|bsc|accounting/.test(c)) return 'People & Payroll';
    if (/insurance|guardian|healthcar|life ins|next insur/.test(c))                       return 'Insurance';
    if (/divvy|bank fees|paris|apt maintenance|charity|donation/.test(c))                 return 'Operations & Facilities';
    if (/legal|law|roetzel/.test(c))                                                      return 'Legal';
    if (/ellison|solaris|wasica|book credit|macdonald|cherry valley|sa nj|realty/.test(c))return 'Customers & Loan Repayments';
    if (/fidelity|money market|dividend|interest/.test(c))                                return 'Fidelity Investments';
    if (/inter-entity|blue panda|nf europe|nf mde co|nf texas|nf6 tiger/.test(c))          return 'Inter-Entity';
    if (/wire|ach payment|book transfer|deposit|other/.test(c))                           return 'Uncategorized / Catch-all';
    return 'Other';
  }
  var recIn = [], recOut = [], nonIn = [], nonOut = [], ieIn = [], ieOut = [];
  mat.forEach(function(c) {
    var tIn = 0, tOut = 0;
    months.forEach(function(ym) {
      var v = c.months[ym] || 0;
      if (v >= 0) tIn += v; else tOut += v;
    });
    var dir = Math.abs(tIn) > Math.abs(tOut) ? 'in' : 'out';
    var isIEcat = /blue panda|nf europe|nf mde co|nf texas|nf6 tiger|inter-entity/i.test(c.category);
    if (isIEcat) (dir === 'in' ? ieIn : ieOut).push(c);
    else if (c.recurring) (dir === 'in' ? recIn : recOut).push(c);
    else (dir === 'in' ? nonIn : nonOut).push(c);
  });

  function fmtCell(v) {
    if (!v) return '<td>&middot;</td>';
    var cls = v < 0 ? 'neg' : 'pos';
    var sign = v < 0 ? '−$' : '$';
    return '<td class="' + cls + '">' + sign + Math.abs(Math.round(v)).toLocaleString() + '</td>';
  }
  function catRowHtml(c) {
    var tds = months.map(function(ym) { return fmtCell(c.months[ym] || 0); }).join('');
    return '<tr><td class="catname">' + _tlmndEsc_(c.category) + '</td>' + tds +
      '<td><strong>' + fmtCell(c.total).replace('<td class="', '').replace('">','">').replace('</td>','') + '</strong></td></tr>';
  }
  function catRowSimple(c) {
    var tds = months.map(function(ym) { return fmtCell(c.months[ym] || 0); }).join('');
    var totalCls = c.total < 0 ? 'neg' : 'pos';
    var totalSign = c.total < 0 ? '−$' : '$';
    return '<tr><td class="catname">' + _tlmndEsc_(c.category) + '</td>' + tds +
      '<td class="' + totalCls + '"><strong>' + totalSign + Math.abs(Math.round(c.total)).toLocaleString() + '</strong></td></tr>';
  }
  function totalRowHtml(label, arr, cls) {
    var perMonth = months.map(function(ym) {
      return arr.reduce(function(s, c) { return s + (c.months[ym] || 0); }, 0);
    });
    var grand = perMonth.reduce(function(s, v) { return s + v; }, 0);
    var tds = perMonth.map(function(v) { return fmtCell(v); }).join('');
    var gCls = grand < 0 ? 'neg' : 'pos';
    var gSign = grand < 0 ? '−$' : '$';
    return '<tr class="' + (cls || 'total') + '"><td class="catname">' + label + '</td>' + tds +
      '<td class="' + gCls + '"><strong>' + gSign + Math.abs(Math.round(grand)).toLocaleString() + '</strong></td></tr>';
  }
  function sectionRowHtml(label) {
    var span = months.length + 2;
    return '<tr class="section"><td colspan="' + span + '">' + label + '</td></tr>';
  }
  function groupHdrHtml(label) {
    var span = months.length + 2;
    return '<tr class="grp"><td colspan="' + span + '">' + label + '</td></tr>';
  }
  function renderSection(sectionLbl, items) {
    var out = sectionRowHtml(sectionLbl);
    if (!items.length) return out;
    var byGroup = {};
    items.forEach(function(c) { var g = _grpOf(c.category); (byGroup[g] = byGroup[g] || []).push(c); });
    var order = ['People & Payroll','Insurance','Operations & Facilities','Legal','Customers & Loan Repayments','Fidelity Investments','Inter-Entity','Uncategorized / Catch-all','Other'];
    order.forEach(function(g) {
      if (!byGroup[g] || !byGroup[g].length) return;
      byGroup[g].sort(function(a, b) { return Math.abs(b.total) - Math.abs(a.total); });
      out += groupHdrHtml(g);
      byGroup[g].forEach(function(c) { out += catRowSimple(c); });
    });
    return out;
  }

  // Header row for the matrix.
  var monthShort = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var matThead = '<thead><tr><th class="catname">Category</th>' +
    months.map(function(ym) {
      var p = ym.split('-'); return '<th>' + monthShort[Number(p[1])-1] + ' ' + p[0].slice(2) + '</th>';
    }).join('') + '<th>Total</th></tr></thead>';

  var matBody = '<tbody>' +
    renderSection('Recurring &mdash; Money In', recIn) +
    totalRowHtml('Total Recurring In', recIn) +
    renderSection('Recurring &mdash; Money Out', recOut) +
    totalRowHtml('Total Recurring Out', recOut) +
    totalRowHtml('Net Recurring', recIn.concat(recOut), 'net') +
    renderSection('Non-Recurring &mdash; Money In', nonIn) +
    totalRowHtml('Total Non-Recurring In', nonIn) +
    renderSection('Non-Recurring &mdash; Money Out', nonOut) +
    totalRowHtml('Total Non-Recurring Out', nonOut) +
    totalRowHtml('Net Non-Recurring', nonIn.concat(nonOut), 'net') +
    renderSection('Inter-Entity Transfers &mdash; In', ieIn) +
    totalRowHtml('Total Inter-Entity In', ieIn) +
    renderSection('Inter-Entity Transfers &mdash; Out', ieOut) +
    totalRowHtml('Total Inter-Entity Out', ieOut) +
    totalRowHtml('Net Inter-Entity', ieIn.concat(ieOut), 'net') +
    totalRowHtml('NET ALL', mat, 'grand') +
    '</tbody>';

  // Build the top movers HTML (Page 1) — individual transactions this
  // week with their transaction date prefix so Mike can see when it hit.
  function topMoverList(arr, max) {
    if (!arr.length) return '<div class="empty">No activity in the last 30 days</div>';
    var shown = arr.slice(0, max);
    return '<table class="movers">' + shown.map(function(x) {
      var cls = x.val < 0 ? 'neg' : 'pos';
      return '<tr><td>' + _tlmndEsc_(x.name || '') + '</td><td class="' + cls + '">' +
        (x.val < 0 ? '−$' : '+$') + Math.abs(Math.round(x.val)).toLocaleString() + '</td></tr>';
    }).join('') + '</table>';
  }

  // Build comparison table (Page 1).
  var cmpBody = compareRows.map(function(r) {
    var cls = r.isNet ? ' class="net"' : '';
    return '<tr' + cls + '><td class="lbl">' + r.label + '</td>' +
      '<td>' + _tlmndFmtSvr_(r.last30)   + '</td>' +
      '<td>' + _tlmndFmtSvr_(r.prior30)  + '</td>' +
      '<td>' + _tlmndFmtSvr_(r.expected) + '</td>' +
      '<td>' + _tlmndFmtSvr_(r.t3)       + '</td>' +
      '<td>' + _tlmndFmtSvr_(r.t6)       + '</td></tr>';
  }).join('');

  var heroCls = cur.netAll >= 0 ? 'pos' : 'neg';
  var heroVal = (cur.netAll < 0 ? '−$' : '+$') + Math.abs(Math.round(cur.netAll)).toLocaleString();

  return '' +
    '<!DOCTYPE html><html><head><meta charset="utf-8"><title>TLMND Cash Flow</title>' +
    '<style>' +
      '@page { size: letter; margin: 0.4in 0.4in 0.5in 0.4in; }' +
      'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;color:#0d2137;margin:0;padding:0;font-size:11px;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +
      '.report-hdr{border-bottom:3px solid #0d2137;padding-bottom:8px;margin-bottom:14px;display:flex;justify-content:space-between;align-items:baseline}' +
      '.report-hdr h1{font-size:16px;margin:0;color:#0d2137;letter-spacing:.3px}' +
      '.report-hdr .date{font-size:11px;color:#5f6368}' +
      '.balances{background:#eef4fa;border:1px solid #d0dae5;border-radius:6px;padding:10px 14px;margin-bottom:12px}' +
      '.balances .lbl{font-size:9px;color:#5f6368;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}' +
      '.balances .asof{float:right;font-size:9px;color:#5f6368;font-style:italic;text-transform:none;letter-spacing:0}' +
      '.balances table{width:100%;border-collapse:collapse;font-size:11px}' +
      '.balances td{padding:3px 0;font-variant-numeric:tabular-nums}' +
      '.balances td.acct{color:#3c4858}' +
      '.balances td.amt{text-align:right;font-weight:700;color:#0d2137;width:120px}' +
      '.balances tr.total td{border-top:1px solid #b8c7d4;padding-top:6px;margin-top:4px;font-weight:800}' +
      '.balances tr.total td.acct{color:#0d2137}' +
      '.hero{background:#0d2137;color:#fff;padding:16px 20px;border-radius:6px;margin-bottom:12px}' +
      '.hero .lbl{font-size:10px;text-transform:uppercase;letter-spacing:.6px;opacity:.75;margin-bottom:4px}' +
      '.hero .val{font-size:32px;font-weight:800;line-height:1;letter-spacing:-.5px}' +
      '.hero .val.pos{color:#7fdba0}' +
      '.hero .val.neg{color:#ff9d9d}' +
      '.hero .sub{margin-top:8px;font-size:11px;opacity:.9}' +
      '.two{display:table;width:100%;border-spacing:8px 0;margin-bottom:12px}' +
      '.two .col{display:table-cell;width:50%;background:#f8f9fa;border-radius:6px;padding:10px 14px;vertical-align:top;border-left:3px solid #ddd}' +
      '.two .col.in{border-left-color:#137333}' +
      '.two .col.out{border-left-color:#a50e0e}' +
      '.two h3{font-size:10px;color:#666;text-transform:uppercase;letter-spacing:.5px;margin:0 0 4px 0;font-weight:600}' +
      '.two .subtotal{font-size:16px;font-weight:700;margin-bottom:6px}' +
      '.two .subtotal.pos{color:#137333}' +
      '.two .subtotal.neg{color:#a50e0e}' +
      '.movers{width:100%;font-size:10px;border-collapse:collapse}' +
      '.movers td{padding:3px 0;border-bottom:1px solid #eee}' +
      '.movers td:last-child{text-align:right;font-variant-numeric:tabular-nums;font-weight:600;white-space:nowrap}' +
      '.pos{color:#137333}.neg{color:#a50e0e}' +
      '.compare,.forecast{background:#fff;border:1px solid #e0e5eb;border-radius:6px;padding:10px 14px;margin-bottom:12px}' +
      '.compare h3,.forecast h3{font-size:10px;color:#666;text-transform:uppercase;letter-spacing:.5px;margin:0 0 8px 0;font-weight:600}' +
      '.compare table{width:100%;border-collapse:collapse;font-size:11px}' +
      '.compare th{background:#f8f9fa;padding:5px 8px;text-align:right;color:#5f6368;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:.4px;white-space:nowrap;border-bottom:1px solid #dadce0}' +
      '.compare th:first-child,.compare td:first-child{text-align:left}' +
      '.compare td{padding:5px 8px;text-align:right;font-variant-numeric:tabular-nums;border-bottom:1px solid #f4f5f7}' +
      '.compare tr.net td{font-weight:700;border-top:2px solid #dadce0;background:#f5f7fa}' +
      '.forecast .tiles{display:table;width:100%;border-spacing:6px 0;margin-bottom:6px}' +
      '.forecast .tile{display:table-cell;width:33%;background:#f8f9fa;border-radius:4px;padding:8px 12px;vertical-align:top}' +
      '.forecast .tile .l{font-size:9px;color:#666;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px}' +
      '.forecast .tile .v{font-size:15px;font-weight:700;font-variant-numeric:tabular-nums}' +
      '.forecast .call{background:#eef4fa;border-radius:4px;padding:10px 14px}' +
      '.forecast .call .l{font-size:9px;color:#5f6368;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px}' +
      '.forecast .call .v{font-size:22px;font-weight:800;color:#0d2137;font-variant-numeric:tabular-nums}' +
      '.forecast .call .n{font-size:10px;color:#5f6368;margin-top:3px}' +
      'table.matrix{width:100%;border-collapse:collapse;font-size:9.5px;margin-top:6px}' +
      '.matrix-page{page-break-before:always}' +
      '.matrix-page .h2title{margin-top:0;padding-top:0}' +
      'table.matrix th{background:#f8f9fa;padding:5px 6px;text-align:right;color:#5f6368;font-weight:600;text-transform:uppercase;font-size:8px;letter-spacing:.4px;border-bottom:2px solid #dadce0;white-space:nowrap}' +
      'table.matrix th.catname,table.matrix td.catname{text-align:left}' +
      'table.matrix td{padding:4px 6px;text-align:right;font-variant-numeric:tabular-nums;border-bottom:1px solid #f4f5f7;white-space:nowrap}' +
      'table.matrix tr.section td{background:#0d2137;color:#fff;font-weight:700;font-size:9px;text-transform:uppercase;letter-spacing:.5px;padding:6px 8px;text-align:left}' +
      'table.matrix tr.grp td{background:#dfe4ea;color:#2c3e50;font-weight:700;text-transform:uppercase;font-size:8px;letter-spacing:.5px;padding:4px 8px 4px 20px;text-align:left}' +
      'table.matrix tr.total td{background:#dae5ee;font-weight:700;color:#0d2137;border-top:1px solid #b8c7d4}' +
      'table.matrix tr.net td{background:#c8d8e5;font-weight:800;font-size:10.5px;color:#0d2137;border-top:2px solid #0d2137;border-bottom:2px solid #0d2137;padding:6px 8px}' +
      'table.matrix tr.grand td{background:#0d2137;color:#fff;font-weight:800;font-size:11px;border-top:3px double #0d2137;padding:8px}' +
      '.h2title{font-size:12px;font-weight:700;color:#0d2137;margin:0 0 4px 0;padding-top:4px}' +
      '.footer{margin-top:10px;font-size:8px;color:#9aa0a6;font-style:italic}' +
    '</style></head><body>' +
    // ── PAGE 1 ────────────────────────────────────────────────
    '<div class="report-hdr">' +
      '<h1>TLMND Cash Flow &mdash; ' + monthLabel + '</h1>' +
      '<div class="date">Report generated ' + _tlmndEsc_(reportDate) + '</div>' +
    '</div>' +
    // Current Balances strip
    (balList.length
      ? '<div class="balances">' +
          '<div class="lbl">Current Balances <span class="asof">as of ' + _tlmndEsc_(balAsOf) + '</span></div>' +
          '<table>' +
            balList.map(function(b) {
              var v = Number(b.value) || 0;
              return '<tr><td class="acct">' + _tlmndEsc_(b.label) + '</td>' +
                '<td class="amt">$' + Math.round(v).toLocaleString() + '</td></tr>';
            }).join('') +
            '<tr class="total"><td class="acct">Total on hand</td>' +
              '<td class="amt">$' + Math.round(balTotal).toLocaleString() + '</td></tr>' +
          '</table>' +
        '</div>'
      : '') +
    // Hero (Last 30 Days actual) + Forecast (Next 30 Days expected)
    '<table style="width:100%;border-spacing:8px 0;margin-bottom:12px"><tr>' +
      '<td style="width:60%;vertical-align:top;padding:0">' +
        '<div class="hero" style="margin-bottom:0">' +
          '<div class="lbl">Net Cash Flow &middot; ' + _tlmndEsc_(monthLabel) + '</div>' +
          '<div class="val ' + heroCls + '">' + heroVal + '</div>' +
          '<div class="sub">' + _tlmndEsc_(deltaTxt) + '</div>' +
        '</div>' +
      '</td>' +
      '<td style="width:40%;vertical-align:top;padding:0">' +
        '<div style="background:#f8f9fa;border:1px solid #e0e5eb;border-radius:6px;padding:14px 16px;height:100%;box-sizing:border-box">' +
          '<div style="font-size:10px;color:#666;text-transform:uppercase;letter-spacing:.5px;font-weight:600;margin-bottom:4px">Next 30 Days &middot; Expected (recurring)</div>' +
          '<div style="font-size:22px;font-weight:800;font-variant-numeric:tabular-nums;line-height:1;margin-bottom:10px;color:' + (fc.net >= 0 ? '#137333' : '#a50e0e') + '">' + _tlmndFmtSvr_(fc.net) + '</div>' +
          '<div style="display:flex;justify-content:space-between;font-size:10.5px;color:#3c4858;padding:3px 0"><span>Expected in</span><span style="font-weight:700;color:#137333">+$' + Math.round(fc.totalIn).toLocaleString() + '</span></div>' +
          '<div style="display:flex;justify-content:space-between;font-size:10.5px;color:#3c4858;padding:3px 0"><span>Expected out</span><span style="font-weight:700;color:#a50e0e">-$' + Math.abs(Math.round(fc.totalOut)).toLocaleString() + '</span></div>' +
        '</div>' +
      '</td>' +
    '</tr></table>' +
    // Top movers (two cards side by side)
    '<div class="two">' +
      '<div class="col in">' +
        '<h3>Last 30 Days &middot; Money In (Top Sources)</h3>' +
        '<div class="subtotal pos">+$' + Math.round(totalIn).toLocaleString() + '</div>' +
        topMoverList(inItems, 6) +
      '</div>' +
      '<div class="col out">' +
        '<h3>Last 30 Days &middot; Top Expenses</h3>' +
        '<div class="subtotal neg">-$' + Math.abs(Math.round(totalOut)).toLocaleString() + '</div>' +
        topMoverList(outItems, 6) +
      '</div>' +
    '</div>' +
    // On-Track Analysis (Expected vs Actual, last 30 days recurring)
    (function() {
      function isIE(c) { return /blue panda|nf europe|nf mde co|nf texas|nf6 tiger|inter-entity/i.test(c || ''); }
      var actRecIn = 0, actRecOut = 0, actNonRecIn = 0, actNonRecOut = 0;
      txns.forEach(function(t) {
        if (t.excluded || !t.date) return;
        if (t.date < start30 || t.date > yesterdayISO) return;
        if (isIE(t.category)) return;
        if (t.recurring) {
          if (t.amount >= 0) actRecIn += t.amount; else actRecOut += t.amount;
        } else {
          if (t.amount >= 0) actNonRecIn += t.amount; else actNonRecOut += t.amount;
        }
      });
      var expNet = fc.net, actNet = actRecIn + actRecOut;
      var netVar = actNet - expNet;
      var varIn  = actRecIn  - fc.totalIn;
      var varOut = actRecOut - fc.totalOut;
      var status, statusBg, statusFg;
      if (Math.abs(netVar) < 3000)      { status = 'ON TRACK';       statusBg = '#e6f4ea'; statusFg = '#137333'; }
      else if (netVar > 0)              { status = 'AHEAD OF PLAN';  statusBg = '#e6f4ea'; statusFg = '#137333'; }
      else                              { status = 'BEHIND PLAN';    statusBg = '#fce8e6'; statusFg = '#a50e0e'; }
      function sVar(v, higherIsBetter) {
        var sign = v >= 0 ? '+' : '-';
        var absV = Math.abs(Math.round(v)).toLocaleString();
        var good = higherIsBetter ? v >= 0 : v >= 0;
        var color = good ? '#137333' : '#a50e0e';
        return '<span style="color:' + color + ';font-weight:700">' + sign + '$' + absV + '</span>';
      }
      return '<div style="background:#fff;border:1px solid #e0e5eb;border-radius:6px;padding:12px 16px;margin-bottom:12px;border-left:4px solid #4fc3f7">' +
        '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">' +
          '<div style="font-size:10px;color:#0d2137;text-transform:uppercase;letter-spacing:.5px;font-weight:700">Are We On Track? · Last 30 Days vs Expected Recurring</div>' +
          '<div style="padding:3px 10px;border-radius:20px;font-size:9px;font-weight:800;letter-spacing:.4px;background:' + statusBg + ';color:' + statusFg + '">' + status + '</div>' +
        '</div>' +
        '<table style="width:100%;border-collapse:collapse;font-size:10.5px">' +
          '<thead><tr>' +
            '<th style="text-align:left;padding:5px 8px;color:#5f6368;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:.4px;border-bottom:1px solid #dadce0"></th>' +
            '<th style="text-align:right;padding:5px 8px;color:#5f6368;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:.4px;border-bottom:1px solid #dadce0">Expected</th>' +
            '<th style="text-align:right;padding:5px 8px;color:#5f6368;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:.4px;border-bottom:1px solid #dadce0">Actual</th>' +
            '<th style="text-align:right;padding:5px 8px;color:#5f6368;font-weight:600;text-transform:uppercase;font-size:9px;letter-spacing:.4px;border-bottom:1px solid #dadce0">Variance</th>' +
          '</tr></thead><tbody>' +
            '<tr><td style="padding:6px 8px">Recurring Money In</td>' +
              '<td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums;color:#137333">+$' + Math.round(fc.totalIn).toLocaleString() + '</td>' +
              '<td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums;color:#137333">+$' + Math.round(actRecIn).toLocaleString() + '</td>' +
              '<td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums">' + sVar(varIn, true) + '</td></tr>' +
            '<tr><td style="padding:6px 8px">Recurring Money Out</td>' +
              '<td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums;color:#a50e0e">-$' + Math.abs(Math.round(fc.totalOut)).toLocaleString() + '</td>' +
              '<td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums;color:#a50e0e">-$' + Math.abs(Math.round(actRecOut)).toLocaleString() + '</td>' +
              '<td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums">' + sVar(varOut, false) + '</td></tr>' +
            '<tr style="border-top:2px solid #dadce0;background:#f5f7fa;font-weight:800"><td style="padding:6px 8px">Net Recurring</td>' +
              '<td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums">' + _tlmndFmtSvr_(expNet) + '</td>' +
              '<td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums">' + _tlmndFmtSvr_(actNet) + '</td>' +
              '<td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums">' + sVar(netVar, true) + '</td></tr>' +
          '</tbody></table>' +
          (Math.abs(actNonRecIn) + Math.abs(actNonRecOut) > 0
            ? '<div style="margin-top:8px;padding:8px 12px;background:#fafbfc;border-radius:4px;font-size:10px;color:#3c4858"><strong>Plus unplanned (non-recurring):</strong> Money in +$' + Math.round(actNonRecIn).toLocaleString() + ' &middot; Money out -$' + Math.abs(Math.round(actNonRecOut)).toLocaleString() + '</div>'
            : '') +
      '</div>';
    })() +
    // Comparison table — Rolling 30-day + Expected + T3/T6
    '<div class="compare">' +
      '<h3>Rolling comparison &middot; ' + _tlmndDateShort_(start30) + ' – ' + _tlmndDateShort_(yesterdayISO) + '</h3>' +
      '<table><thead><tr><th></th>' +
        '<th>Last 30 Days</th><th>Prior 30 Days</th>' +
        '<th>Next 30 (Exp.)</th>' +
        '<th>3-Mo Avg</th><th>6-Mo Avg</th>' +
      '</tr></thead>' +
      '<tbody>' + cmpBody + '</tbody></table>' +
    '</div>' +
    // Burn & Forecast
    '<div class="forecast">' +
      '<h3>Cash Needs &amp; Forecast (recurring, trailing 6 completed months)</h3>' +
      '<div class="tiles">' +
        '<div class="tile"><div class="l">Avg Recurring In</div><div class="v pos">+' + _tlmndFmtPos_(avgRecIn) + '</div></div>' +
        '<div class="tile"><div class="l">Avg Recurring Out</div><div class="v neg">-' + _tlmndFmtPos_(Math.abs(avgRecOut)) + '</div></div>' +
        '<div class="tile"><div class="l">Net Monthly Recurring</div><div class="v ' + (netRec >= 0 ? 'pos' : 'neg') + '">' + _tlmndFmtSvr_(netRec) + '</div></div>' +
      '</div>' +
      '<div class="call">' +
        (monthlyNeed > 0
          ? '<div class="l">Monthly transfer needed into TLMND</div><div class="v">' + _tlmndFmtPos_(monthlyNeed) + '</div>' +
            '<div class="n">Recurring money out exceeds money in by this much on average. Move in ~' + _tlmndFmtPos_(monthlyNeed) + '/mo from Blue Panda or other entities to keep TLMND self-funding.</div>'
          : '<div class="l">Monthly recurring surplus</div><div class="v" style="color:#137333">+' + _tlmndFmtPos_(-monthlyNeed) + '</div>' +
            '<div class="n">TLMND is self-funding on the recurring baseline.</div>') +
      '</div>' +
    '</div>' +
    // ── PAGE 2 ────────────────────────────────────────────────
    // Wrap title + matrix in one .matrix-page block so page-break stays
    // ABOVE the title (they render on the same new page together).
    '<div class="matrix-page">' +
      '<h2 class="h2title">Full Category Breakdown &mdash; Last 6 Months</h2>' +
      '<table class="matrix">' + matThead + matBody + '</table>' +
      '<div class="footer">Data pulled from Plaid + SnapTrade &middot; excludes internal-account transfers to avoid double-counting.</div>' +
    '</div>' +
    '</body></html>';
}

// Called by the Monday trigger. Reuses WEEKLY_PDF_RECIPIENT script property
// so it goes to the same distribution as the net-worth weekly email.
function weeklyTLMNDCashFlowEmail() {
  var recipient = PropertiesService.getScriptProperties().getProperty('WEEKLY_PDF_RECIPIENT');
  if (!recipient) {
    Logger.log('weeklyTLMNDCashFlowEmail: WEEKLY_PDF_RECIPIENT not set — skipping');
    return { success: false, error: 'No recipient configured. Set via Tracker > Set Weekly PDF Email Recipient.' };
  }
  return _tlmndSendPdfEmail_(recipient, 'Weekly TLMND Cash Flow');
}

function _tlmndSendPdfEmail_(recipient, subjectPrefix) {
  // Ensure a fresh sync so the PDF reflects the latest transactions.
  try { syncTLMNDCashFlow(); } catch(e) { Logger.log('TLMND PDF: pre-sync failed: ' + e.message); }

  var html = _tlmndBuildWeeklyPdfHtml_();
  if (!html) return { success: false, error: 'Failed to generate report HTML — TLMND config may not be set.' };

  var dateStr = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');
  var pdfBlob = Utilities.newBlob(html, 'text/html', 'TLMND_Cash_Flow_' + dateStr + '.html')
    .getAs('application/pdf').setName('TLMND_Cash_Flow_' + dateStr + '.pdf');

  var dateLabel = Utilities.formatDate(new Date(), 'America/New_York', 'MMMM d, yyyy');
  MailApp.sendEmail({
    to:          recipient,
    subject:     subjectPrefix + ' — ' + dateLabel,
    body:        'Your ' + subjectPrefix.toLowerCase() + ' PDF is attached.\n\n' +
                 'Includes:\n' +
                 '  • Net Cash Flow for the current month + comparison to averages\n' +
                 '  • Top sources of money in and top expenses this month\n' +
                 '  • Full category breakdown across the last 6 months (recurring, non-recurring, inter-entity)\n' +
                 '  • Cash needs forecast based on trailing 6 completed months\n\n' +
                 'The live dashboard is at your Family Office Tracker web app.',
    name:        'TLMND Cash Flow',
    attachments: [pdfBlob]
  });
  Logger.log('TLMND weekly PDF sent to ' + recipient);
  return { success: true };
}

// Menu-callable test send — goes to the current spreadsheet owner so
// you can preview the PDF before enabling the trigger.
function sendTLMNDWeeklyPdfTest() {
  var ui = SpreadsheetApp.getUi();
  var me = Session.getActiveUser().getEmail();
  if (!me) { ui.alert('Could not determine your email address.'); return; }
  var r = _tlmndSendPdfEmail_(me, 'TEST — TLMND Cash Flow');
  ui.alert(r.success ? 'Test PDF sent to ' + me : 'Send failed: ' + (r.error || 'unknown'), '', ui.ButtonSet.OK);
}

// Install/replace the Monday 8am trigger for the weekly TLMND PDF.
function installTLMNDWeeklyPdfTrigger() {
  var ui = SpreadsheetApp.getUi();
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'weeklyTLMNDCashFlowEmail') {
      ScriptApp.deleteTrigger(t); removed++;
    }
  });
  ScriptApp.newTrigger('weeklyTLMNDCashFlowEmail')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  ui.alert('Installed weekly TLMND PDF trigger.\n\nFires every Monday at 8 AM ET.\nReplaced ' + removed + ' prior trigger(s).\n\nEmail goes to whichever address is configured under Set Weekly PDF Email Recipient (same as the net-worth weekly email).');
}

function installTLMNDCashFlowTrigger() {
  var ui = SpreadsheetApp.getUi();
  var existing = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === '_tlmndDailySync'; });
  existing.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('_tlmndDailySync').timeBased().atHour(4).nearMinute(30).everyDays(1).create();
  ui.alert('Installed daily TLMND cash flow sync at 4:30 AM.\n\n' +
           'Replaced ' + existing.length + ' prior trigger(s) if any.');
}
