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
  'Priority',     // A
  'Match Field',  // B
  'Match Type',   // C
  'Pattern',      // D
  'Amount Min',   // E
  'Amount Max',   // F
  'Category',     // G
  'Recurring',    // H
  'Entity Tag',   // I
  'Exclude',      // J
  'Enabled',      // K
  'Notes'         // L
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
function seedTLMNDRules() {
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
    [10, 'Name', 'contains', 'SOLARIS-FL HOLDI',                  '', '', 'Solaris-Fl Holding LLC Loan Repayment Income', 'Yes', 'TLM',        '', 'Yes', 'Monthly ~$16,656'],
    [10, 'Name', 'contains', 'ELLISON MEDICAL',                   '', '', 'Ellison Medical - Customer (Carroll Canyon)',   'Yes', 'NF USA CA',  '', 'Yes', 'Lands on NF USA CA ···2086'],
    [10, 'Name', 'contains', 'BOOK TRANSFER CREDIT B/O: WASICA',  '', '', 'Wasica Holdings (Book Credit)',                 'No',  'TLM',        '', 'Yes', ''],

    // ── MONEY OUT — recurring ────────────────────────────────────────────
    [10, 'Name', 'contains', 'UNITED HEALTHCAR',                  '', '', 'United Healthcare Insurance',                   'Yes', 'TLM',        '', 'Yes', 'Monthly ~$9,764'],
    [10, 'Name', 'contains', 'THE GUARDIAN',                      '', '', 'The Guardian Insurance',                        'Yes', 'TLM',        '', 'Yes', 'Monthly ~$625'],
    [10, 'Name', 'contains', 'EWALLET - Divvy',                   '', '', 'Divvy Bill (Grand Total)',                      'Yes', 'TLM',        '', 'Yes', ''],
    [10, 'Name', 'contains', 'BSCAccountingLLC',                  '', '', 'BSC Accounting LLC (Accounting Fees)',          'Yes', 'TLM',        '', 'Yes', 'Monthly -$3,500'],
    [10, 'Name', 'contains', 'PENN MUTUAL LIFE INS',              '', '', 'Life Insurance (Waskar Tejeda / Penn Mutual)',  'Yes', 'TLM',        '', 'Yes', ''],
    [10, 'Name', 'contains', 'To ManuEstrada',                    '', '', 'Manuela Estrada - Legal Fees',                  'Yes', 'TLM',        '', 'Yes', ''],
    [10, 'Name', 'contains', 'To LynnNguyen',                     '', '', 'Lynn Repayment',                                'Yes', 'TLM',        '', 'Yes', ''],
    [10, 'Name', 'contains', 'MANUELA VALLEJO',                   '', '', 'Consulting - Manuela Vallejo',                  'Yes', 'TLM',        '', 'Yes', 'International wire, Vietnam'],

    // ── MONEY OUT — non-recurring / one-offs ─────────────────────────────
    [20, 'Name', 'contains', 'NF EUROPE HOLDINGS',                '', '', 'NF Europe Holdings (Inter-Entity Transfer)',    'No',  'NF',         '', 'Yes', ''],
    [20, 'Name', 'contains', 'NF MDECO SAS',                      '', '', 'NF Medellin (Inter-Entity Transfer)',           'No',  'NF',         '', 'Yes', 'Via BTG Pactual'],
    [20, 'Name', 'contains', 'ROETZEL AND ANDRESS',               '', '', 'Legal Fees - Roetzel and Andress',              'No',  'TLM',        '', 'Yes', ''],
    [20, 'Name', 'contains', 'THE HOUSE PROJECT FOUNDATION',      '', '', 'Consulting - Manuela Estrada (House Project)',  'No',  'TLM',        '', 'Yes', ''],

    // ── FIDELITY (SnapTrade) — real cash flow ────────────────────────────
    [30, 'Name', 'contains', 'GUSTO NET',                         '', '', 'Payroll (Net Wages)',                           'Yes', 'TLM',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'GUSTO TAX',                         '', '', 'Payroll (Employer Taxes)',                      'Yes', 'TLM',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'GUSTO ICD',                         '', '', 'Payroll (Contractor Deposits)',                 'Yes', 'TLM',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'GUSTO FEE',                         '', '', 'Payroll (Gusto Fees)',                          'Yes', 'TLM',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'GUSTO CND',                         '', '', 'Payroll (Contractor Non-Deposit)',              'Yes', 'TLM',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'NEXT INSUR',                        '', '', 'Business Insurance (Next Insurance)',           'Yes', 'TLM',        '', 'Yes', ''],
    [30, 'Name', 'contains', 'DIVIDEND SPAXX',                    '', '', 'Fidelity Money Market Interest',                'Yes', 'TLM',        '', 'Yes', ''],

    // ── EXCLUDE — internal cash mgmt / would double-count ────────────────
    // Fidelity SPAXX buy/sell/reinvest — internal cash sweep, not real flow.
    [40, 'Name', 'contains', 'BUY SPAXX',                         '', '', '(Fidelity SPAXX cash mgmt)',                     '',   '',           'Yes', 'Yes', 'Excluded — internal'],
    [40, 'Name', 'contains', 'SELL SPAXX',                        '', '', '(Fidelity SPAXX cash mgmt)',                     '',   '',           'Yes', 'Yes', 'Excluded — internal'],
    [40, 'Name', 'contains', 'REI SPAXX',                         '', '', '(Fidelity SPAXX reinvestment)',                  '',   '',           'Yes', 'Yes', 'Excluded — internal'],
    // TLMND → Fidelity transfer, both sides.
    [40, 'Name', 'contains', 'To FidelityTLMND',                  '', '', '(Transfer TLMND → Fidelity)',                    '',   '',           'Yes', 'Yes', 'Excluded — paired w/ Fidelity CONTRIBUTION'],
    [40, 'Name', 'contains', 'CONTRIBUTION — DIRECT DEPOSIT TLMND','', '', '(Transfer TLMND → Fidelity)',                    '',   '',           'Yes', 'Yes', 'Excluded — internal'],

    // ── Bank noise ───────────────────────────────────────────────────────
    [50, 'Name', 'contains', 'SERVICE CHARGES FOR THE MONTH',     '', '', 'Bank Fees',                                     'Yes', 'TLM',        '', 'Yes', ''],
    [50, 'Name', 'contains', 'ACCOUNT ANALYSIS SETTLEMENT',       '', '', 'Bank Fees',                                     'Yes', 'TLM',        '', 'Yes', ''],

    // ── Inter-account journal transfers ──────────────────────────────────
    // NF USA CA ↔ TLMND: these are Ellison Medical proceeds being moved from
    // where they land (NF USA CA ···2086) to where they belong (TLMND ···2001).
    // We already count the ELLISON MEDICAL deposit as income, so exclude the
    // internal journal to avoid double-counting.
    [90, 'Name', 'contains', 'Online Transfer from CHK ...2086',  '', '', '(Journal from NF USA CA → TLMND)',              '',   '',           'Yes', 'Yes', 'Excluded — paired with Ellison Medical'],
    [90, 'Name', 'contains', 'Online Transfer to CHK ...2001',    '', '', '(Journal from NF USA CA → TLMND)',              '',   '',           'Yes', 'Yes', 'Excluded — paired with Ellison Medical'],
    // Blue Panda Family ···8686 → TLMND: real inter-entity funding, COUNT it.
    [90, 'Name', 'contains', 'Online Transfer from CHK ...8686',  '', '', 'Transfer from Blue Panda Family',               'No',  'TLM',        '', 'Yes', 'Blue Panda Family ···8686 → TLMND funding'],
    // TLMND → NF USA TX ···5155: real inter-entity outflow, COUNT it.
    [90, 'Name', 'contains', 'Online Transfer to CHK ...5155',    '', '', 'Transfer to NF USA TX',                         'No',  'TLM',        '', 'Yes', 'TLMND → NF USA TX ···5155']
  ];

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
      priority:  Number(r[0]) || 999,
      field:     String(r[1] || 'Name'),
      matchType: String(r[2] || 'contains').toLowerCase(),
      pattern:   String(r[3]),
      amtMin:    r[4] === '' || r[4] == null ? null : Number(r[4]),
      amtMax:    r[5] === '' || r[5] == null ? null : Number(r[5]),
      category:  String(r[6] || ''),
      recurring: String(r[7] || ''),
      entityTag: String(r[8] || ''),
      exclude:   String(r[9]).toLowerCase() === 'yes'
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
function applyTLMNDRules() {
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
  var idxAmount = ci('Amount USD');
  var idxCat    = ci('Category');
  var idxRec    = ci('Recurring');
  var idxEnt    = ci('Entity Tag');
  var idxNotes  = ci('Notes');
  if (idxCat < 0) return { success: false, error: 'Category column missing.' };

  // Set of rule categories so we can detect manual overrides.
  var ruleCategories = {};
  rules.forEach(function(r) { if (r.category) ruleCategories[r.category] = true; });

  var last = sheet.getLastRow();
  var range = sheet.getRange(2, 1, last - 1, headers.length);
  var vals  = range.getValues();

  var categorized = 0, skippedManual = 0, uncategorized = 0, excluded = 0;

  vals.forEach(function(row, i) {
    var existing = String(row[idxCat] || '').trim();
    // Manual override: category set to a value no rule uses → preserve.
    if (existing && !ruleCategories[existing]) { skippedManual++; return; }

    var matched = null;
    for (var k = 0; k < rules.length; k++) {
      var r = rules[k];
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
      if (matched.exclude) {
        // Prepend [EXCLUDED] marker in Notes so users see it in the sheet at a glance.
        var n = String(row[idxNotes] || '');
        if (n.indexOf('[EXCLUDED]') < 0) row[idxNotes] = ('[EXCLUDED] ' + n).trim();
        excluded++;
      }
      categorized++;
    } else if (!existing) {
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

function installTLMNDCashFlowTrigger() {
  var ui = SpreadsheetApp.getUi();
  var existing = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === '_tlmndDailySync'; });
  existing.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('_tlmndDailySync').timeBased().atHour(4).nearMinute(30).everyDays(1).create();
  ui.alert('Installed daily TLMND cash flow sync at 4:30 AM.\n\n' +
           'Replaced ' + existing.length + ' prior trigger(s) if any.');
}
