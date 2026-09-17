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
  'Source Filter'   // M — Plaid | SnapTrade | blank (any). Restricts the
                    //     rule to only match rows from a specific source.
                    //     Fidelity catch-alls (BUY/SELL/DEPOSIT/etc.) use
                    //     'SnapTrade' so they don't accidentally match
                    //     Plaid rows whose Name starts with the same word
                    //     (e.g. "DEPOSIT ID NUMBER 553083" on Chase).
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
    [10, 'Name', 'contains', 'SOLARIS-FL HOLDI',                  '', '', 'Solaris-Fl Holding LLC Loan Repayment Income', 'Yes', 'TLMND',      '', 'Yes', 'Monthly ~$16,656'],
    // ELLISON MEDICAL: raw deposit lands on NF USA CA ···2086 but is
    // economically TLMND income (per user direction — the money is
    // earmarked for TLMND). Count the deposit itself as recurring
    // Money In; the internal transfer from NF USA CA to TLMND is
    // excluded below to prevent double-counting.
    [10, 'Name', 'contains', 'ELLISON MEDICAL',                   '', '', 'Ellison Medical - Customer (Carroll Canyon)',   'Yes', 'TLMND',      '', 'Yes', 'Deposit lands on NF USA CA ···2086, counted as TLMND income'],
    [10, 'Name', 'contains', 'BOOK TRANSFER CREDIT B/O: WASICA',  '', '', 'Wasica Holdings (Book Credit)',                 'Yes', 'TLMND',      '', 'Yes', 'Recurring inbound'],
    [10, 'Name', 'contains', 'CHERRY VALLEY',                     '', '', 'MacDonald Loan Repayment',                     'Yes', 'TLMND',      '', 'Yes', 'Recurring — from Cherry Valley Construction'],
    [10, 'Name', 'contains', 'SA NJ REALTY',                      '', '', 'ASC Rental Income - TLMND Share (SA NJ Realty)', 'Yes', 'TLMND',   '', 'Yes', 'Recurring — Mike\'s real estate rent (comes in every so often)'],

    // ── MONEY OUT — recurring ────────────────────────────────────────────
    [10, 'Name', 'contains', 'UNITED HEALTHCAR',                  '', '', 'United Healthcare Insurance',                   'Yes', 'TLMND',        '', 'Yes', 'Monthly ~$9,764'],
    [10, 'Name', 'contains', 'THE GUARDIAN',                      '', '', 'The Guardian Insurance',                        'Yes', 'TLMND',        '', 'Yes', 'Monthly ~$625'],
    [10, 'Name', 'contains', 'EWALLET - Divvy',                   '', '', 'Divvy Bill (Grand Total)',                      'Yes', 'TLMND',        '', 'Yes', ''],
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
    [20, 'Name', 'contains', 'NF EUROPE HOLDINGS',                '', '', 'NF Europe Holdings (Inter-Entity Transfer)',    'No',  'NF',         '', 'Yes', 'Non-Paris wires'],
    [20, 'Name', 'contains', 'NF MDECO SAS',                      '', '', 'NF Medellin (Inter-Entity Transfer)',           'No',  'NF',         '', 'Yes', 'Via BTG Pactual'],
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
    [40, 'Name', 'contains', 'FidelityTLM',                       '', '', '(Transfer TLMND → Fidelity)',                    '',   '',           'Yes', 'Yes', 'Excluded — paired w/ Fidelity CONTRIBUTION; loose match catches "FidelityTLM" and "FidelityTLMND"'],
    [40, 'Name', 'contains', 'CONTRIBUTION — DIRECT DEPOSIT TLMND','', '', '(Transfer TLMND → Fidelity)',                    '',   '',           'Yes', 'Yes', 'Excluded — internal'],
    // Fidelity clearing broker (NFS = National Financial Services) returning
    // money to TLMND — internal move, exclude to avoid inflating income.
    [40, 'Name', 'contains', 'NATIONAL FINANCIAL SERVICES',       '', '', '(Transfer Fidelity → TLMND)',                    '',   '',           'Yes', 'Yes', 'Excluded — Fidelity NFS book credit, internal move back to TLMND'],

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

    // ── Inter-account journal transfers ──────────────────────────────────
    // NF USA CA ↔ TLMND internal journals: paired with the Ellison deposit
    // we already count as income. Excluding both sides prevents triple-count
    // (deposit + inbound + outbound = 3× the actual income).
    [90, 'Name', 'contains', 'Online Transfer from CHK ...2086',  '', '', '(Journal from NF USA CA → TLMND)',              '',   '',           'Yes', 'Yes', 'Excluded — paired with Ellison deposit'],
    [90, 'Name', 'contains', 'Online Transfer to CHK ...2001',    '', '', '(Journal from NF USA CA → TLMND)',              '',   '',           'Yes', 'Yes', 'Excluded — mirror of above'],
    // Blue Panda Family ···8686 → TLMND: real inter-entity funding, COUNT it.
    // Match both by account-mask (internal Chase transfer) AND by name
    // substring in case Blue Panda money arrives via a different mechanism
    // (wire, ACH) with a different name format.
    [15, 'Name', 'contains', 'BLUE PANDA',                        '', '', 'Transfer from Blue Panda Family',               'No',  'TLMND',      '', 'Yes', 'Any Blue Panda inbound — catches wires/ACH by name'],
    [90, 'Name', 'contains', 'Online Transfer from CHK ...8686',  '', '', 'Transfer from Blue Panda Family',               'No',  'TLMND',      '', 'Yes', 'Blue Panda Family ···8686 → TLMND funding (Chase internal transfer)'],
    // TLMND ↔ NF USA TX: real inter-entity movement, COUNT it. Match by
    // name substring first (catches wires/ACH) then by the internal
    // Chase transfer format as a fallback.
    [15, 'Name', 'contains', 'NF USA TX',                         '', '', 'Transfer to/from NF USA TX',                    'No',  'TLMND',      '', 'Yes', 'Any NF USA TX movement — catches wires/ACH by name'],
    [90, 'Name', 'contains', 'Online Transfer to CHK ...5155',    '', '', 'Transfer to/from NF USA TX',                    'No',  'TLMND',      '', 'Yes', 'TLMND ↔ NF USA TX ···5155 (Chase internal transfer)']
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
  var plaidPatterns     = ['DOMESTIC WIRE TRANSFER','INTERNATIONAL WIRE','BOOK TRANSFER','Online ACH Payment','ORIG CO NAME:','REMOTE ONLINE DEPOSIT','DEPOSIT ID NUMBER'];
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
      priority:     Number(r[0]) || 999,
      field:        String(r[1] || 'Name'),
      matchType:    String(r[2] || 'contains').toLowerCase(),
      pattern:      String(r[3]),
      amtMin:       r[4] === '' || r[4] == null ? null : Number(r[4]),
      amtMax:       r[5] === '' || r[5] == null ? null : Number(r[5]),
      category:     String(r[6] || ''),
      recurring:    String(r[7] || ''),
      entityTag:    String(r[8] || ''),
      exclude:      String(r[9]).toLowerCase() === 'yes',
      sourceFilter: String(r[12] || '').trim()   // '' = any; 'Plaid' | 'SnapTrade'
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
  var idxSource = ci('Source');
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
      // Always strip any prior [EXCLUDED] marker before re-applying — this
      // is what keeps rows from being permanently stuck as excluded when a
      // rule is later flipped from exclude=Yes to exclude=blank (Ellison
      // Medical is the canonical case). Then re-add the marker only if
      // the CURRENT matched rule is still excluded.
      var n = String(row[idxNotes] || '').replace(/^\[EXCLUDED\]\s*/, '');
      if (matched.exclude) {
        n = '[EXCLUDED] ' + n;
        excluded++;
      }
      row[idxNotes] = n.trim();
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

  return {
    success: true,
    kpis: kpis,
    monthlySeries: monthlySeries,
    categoryMatrix: categoryMatrix,
    transactions: filtered,
    monthKeys: monthKeys,
    entityTags: Object.keys(entTagSet).sort(),
    generatedAt: new Date().toISOString()
  };
}

// Client-callable sync trigger for the dashboard's Refresh button.
function refreshTLMNDCashFlow() {
  return syncTLMNDCashFlow();
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

function installTLMNDCashFlowTrigger() {
  var ui = SpreadsheetApp.getUi();
  var existing = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === '_tlmndDailySync'; });
  existing.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('_tlmndDailySync').timeBased().atHour(4).nearMinute(30).everyDays(1).create();
  ui.alert('Installed daily TLMND cash flow sync at 4:30 AM.\n\n' +
           'Replaced ' + existing.length + ' prior trigger(s) if any.');
}
