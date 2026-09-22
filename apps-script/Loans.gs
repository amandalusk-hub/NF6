/**
 * Loans.gs — Loan registry + amortization + Plaid auto-matching.
 *
 * How it works:
 *  1. LOANS sheet holds one row per loan with terms + a "Plaid Match Pattern"
 *     text that identifies incoming payments from that borrower in the
 *     TLMND_TRANSACTIONS feed (e.g. "SOLARIS-FL HOLDI").
 *  2. For any active loan, _generateAmortizationSchedule_ produces the full
 *     payment schedule from the terms (N monthly rows: date, expected amount,
 *     interest portion, principal portion, remaining balance).
 *  3. _matchLoanPayments_ scans TLMND_TRANSACTIONS for the loan's Plaid
 *     pattern and returns actual received payments.
 *  4. getLoansStatus stitches schedule + payments together for the dashboard
 *     Loans tab: expected vs actual, on-track status, remaining balance.
 *
 * Adding a new loan is a Loans-tab "+ Add Loan" click — no code change.
 */

var LOANS_HEADERS = [
  'ID',
  'Name',
  'Entity',
  'Original Principal',
  'Accrued Interest',
  'Effective Principal',
  'Annual Rate',
  'Term (Months)',
  'First Payment Date',
  'Monthly Payment',
  'Payment Type',           // 'Beginning of Period' | 'End of Period'
  'Plaid Match Pattern',
  'Status',                 // Active | Paid Off | Delinquent | Deferred
  'Notes',
  'Linked Asset ID',        // Optional: the Assets row this loan tracks. When set,
                            // the asset's My Share USD auto-updates on load to the
                            // loan's current outstanding balance (× ownership %).
  'Date Added',
  'Last Updated'
];

function ensureLoansSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('LOANS');
  if (!sheet) {
    sheet = ss.insertSheet('LOANS');
    sheet.getRange(1, 1, 1, LOANS_HEADERS.length)
      .setValues([LOANS_HEADERS])
      .setFontWeight('bold')
      .setBackground('#14263d')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, LOANS_HEADERS.length);
    return sheet;
  }
  // Sheet exists — check the header row and add any missing columns (like
  // "Linked Asset ID" added after Amanda already seeded Solaris). Data in
  // existing rows stays intact; new columns just show up as blank cells.
  var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var missing = LOANS_HEADERS.filter(function(h){ return existing.indexOf(h) < 0; });
  if (missing.length) {
    var startCol = existing.length + 1;
    sheet.getRange(1, startCol, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold').setBackground('#14263d').setFontColor('#ffffff');
  }
  return sheet;
}

// Menu-callable one-time seed for Solaris — uses the exact terms Amanda
// pulled off her existing amortization schedule.
function initSolarisLoan() {
  ensureLoansSheet_();
  var loans = getLoans();
  if (loans.some(function(l){ return String(l.Name||'').toLowerCase().indexOf('solaris') >= 0; })) {
    SpreadsheetApp.getUi().alert('Solaris loan already exists in the LOANS sheet — no change made.');
    return;
  }
  addLoan({
    name: 'Solaris-FL Holding LLC Loan',
    entity: 'TLMND',
    originalPrincipal: 1170000,
    accruedInterest: 87901.85,
    effectivePrincipal: 1257901.85,
    annualRate: 10.25,
    termMonths: 120,
    firstPaymentDate: '2026-01-01',
    monthlyPayment: 16655.63,
    paymentType: 'Beginning of Period',
    plaidPattern: 'SOLARIS-FL HOLDI',
    status: 'Active',
    notes: 'Original $1,170,000 disbursed 04/17/25-04/22/25. ' +
           '258 days accrued interest at 10.25% = $87,901.85. ' +
           'Loan amount w/ accrued interest: $1,257,901.85. ' +
           'Est. total interest over life: $740,773.51.'
  });
  SpreadsheetApp.getUi().alert('Solaris loan added. Open the Loans tab on the dashboard.');
}

function getLoans() {
  var sheet = ensureLoansSheet_();
  if (sheet.getLastRow() < 2) return [];
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, LOANS_HEADERS.length).getValues();
  return vals.map(function(r) {
    var obj = {};
    LOANS_HEADERS.forEach(function(h, i){
      var v = r[i];
      // Normalize dates to ISO strings for the frontend.
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      obj[h] = v;
    });
    return obj;
  }).filter(function(o){ return o.ID; });
}

function addLoan(data) {
  _requireEditor_();
  _logAudit_('addLoan', 'loan', '', data && data.name, 'Added loan: ' + (data && data.name || ''));
  var sheet = ensureLoansSheet_();
  var now = new Date();
  var id = 'l_' + Utilities.getUuid().substring(0, 8);
  var row = LOANS_HEADERS.map(function(h){
    switch(h) {
      case 'ID': return id;
      case 'Name': return data.name || '';
      case 'Entity': return data.entity || '';
      case 'Original Principal': return Number(data.originalPrincipal) || 0;
      case 'Accrued Interest': return Number(data.accruedInterest) || 0;
      case 'Effective Principal': return Number(data.effectivePrincipal) || Number(data.originalPrincipal) || 0;
      case 'Annual Rate': return Number(data.annualRate) || 0;
      case 'Term (Months)': return Number(data.termMonths) || 0;
      case 'First Payment Date': return data.firstPaymentDate || '';
      case 'Monthly Payment': return Number(data.monthlyPayment) || 0;
      case 'Payment Type': return data.paymentType || 'End of Period';
      case 'Plaid Match Pattern': return data.plaidPattern || '';
      case 'Status': return data.status || 'Active';
      case 'Notes': return data.notes || '';
      case 'Linked Asset ID': return data.linkedAssetId || '';
      case 'Date Added': return now;
      case 'Last Updated': return now;
      default: return '';
    }
  });
  sheet.appendRow(row);
  return { success: true, id: id };
}

function updateLoan(id, data) {
  _requireEditor_();
  _logAudit_('updateLoan', 'loan', id, data && data.name, 'Updated loan: ' + (data && data.name || ''));
  var sheet = ensureLoansSheet_();
  if (sheet.getLastRow() < 2) return { success: false, error: 'No loans found.' };
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, LOANS_HEADERS.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) {
      var row = vals[i].slice();
      LOANS_HEADERS.forEach(function(h, idx) {
        switch(h) {
          case 'Name': if (data.name !== undefined) row[idx] = data.name; break;
          case 'Entity': if (data.entity !== undefined) row[idx] = data.entity; break;
          case 'Original Principal': if (data.originalPrincipal !== undefined) row[idx] = Number(data.originalPrincipal); break;
          case 'Accrued Interest': if (data.accruedInterest !== undefined) row[idx] = Number(data.accruedInterest); break;
          case 'Effective Principal': if (data.effectivePrincipal !== undefined) row[idx] = Number(data.effectivePrincipal); break;
          case 'Annual Rate': if (data.annualRate !== undefined) row[idx] = Number(data.annualRate); break;
          case 'Term (Months)': if (data.termMonths !== undefined) row[idx] = Number(data.termMonths); break;
          case 'First Payment Date': if (data.firstPaymentDate !== undefined) row[idx] = data.firstPaymentDate; break;
          case 'Monthly Payment': if (data.monthlyPayment !== undefined) row[idx] = Number(data.monthlyPayment); break;
          case 'Payment Type': if (data.paymentType !== undefined) row[idx] = data.paymentType; break;
          case 'Plaid Match Pattern': if (data.plaidPattern !== undefined) row[idx] = data.plaidPattern; break;
          case 'Status': if (data.status !== undefined) row[idx] = data.status; break;
          case 'Notes': if (data.notes !== undefined) row[idx] = data.notes; break;
          case 'Linked Asset ID': if (data.linkedAssetId !== undefined) row[idx] = data.linkedAssetId; break;
          case 'Last Updated': row[idx] = new Date(); break;
        }
      });
      sheet.getRange(i + 2, 1, 1, LOANS_HEADERS.length).setValues([row]);
      return { success: true };
    }
  }
  return { success: false, error: 'Loan not found: ' + id };
}

function deleteLoan(id) {
  _requireEditor_();
  var sheet = ensureLoansSheet_();
  if (sheet.getLastRow() < 2) return { success: false, error: 'No loans found.' };
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, LOANS_HEADERS.length).getValues();
  var nameIdx = LOANS_HEADERS.indexOf('Name');
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) {
      var name = nameIdx >= 0 ? vals[i][nameIdx] : '';
      _logAudit_('deleteLoan', 'loan', id, name, 'Deleted loan: ' + name);
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  return { success: false, error: 'Loan not found: ' + id };
}

// Amortization schedule generator. Matches the "Beginning of Period" convention
// Amanda's Solaris schedule uses: payment 1 has $0 interest (payment applied
// to principal before interest accrues), then subsequent payments compute
// interest on the pre-payment balance.
function _generateAmortizationSchedule_(loan) {
  var principal = Number(loan['Effective Principal']) || Number(loan['Original Principal']) || 0;
  var annualRate = Number(loan['Annual Rate']) || 0;
  var termMonths = Number(loan['Term (Months)']) || 0;
  var monthlyPayment = Number(loan['Monthly Payment']) || 0;
  var paymentType = String(loan['Payment Type'] || 'End of Period');
  var firstRaw = loan['First Payment Date'];
  var firstPayment = firstRaw instanceof Date ? firstRaw : new Date(firstRaw);

  if (!principal || !termMonths || !monthlyPayment || isNaN(firstPayment.getTime())) return [];

  // Normalize firstPayment to the intended calendar date. "2026-01-01" parsed
  // by new Date() lands on UTC midnight, which in ET reads as 12/31/2025 —
  // shifting all schedule rows a day earlier. Rebuild from the ISO parts so
  // the schedule dates line up with what Amanda entered.
  var firstYmd = Utilities.formatDate(firstPayment, 'UTC', 'yyyy-MM-dd').split('-');
  var firstY = Number(firstYmd[0]), firstM = Number(firstYmd[1]) - 1, firstD = Number(firstYmd[2]);

  var monthlyRate = (annualRate / 100) / 12;
  var balance = principal;
  var schedule = [];
  var isBOP = paymentType.toLowerCase().indexOf('beginning') >= 0;

  for (var n = 1; n <= termMonths; n++) {
    // Safe month-add: clamp the day to the target month's last day so a
    // loan starting on the 31st doesn't roll over ("Feb 31" → "Mar 3").
    var targetY = firstY, targetM = firstM + n - 1;
    while (targetM > 11) { targetY++; targetM -= 12; }
    var daysInTarget = new Date(targetY, targetM + 1, 0).getDate();
    var targetD = Math.min(firstD, daysInTarget);
    var dueDate = new Date(targetY, targetM, targetD);
    var interest, principalPaid, thisPayment;

    if (isBOP && n === 1) {
      // Beginning-of-period, first payment: no interest yet.
      interest = 0;
      thisPayment = Math.min(monthlyPayment, balance);
      principalPaid = thisPayment;
    } else {
      interest = balance * monthlyRate;
      // Final payment: adjust so balance zeroes out exactly.
      var normalPrincipal = monthlyPayment - interest;
      if (normalPrincipal >= balance) {
        principalPaid = balance;
        thisPayment = balance + interest;
      } else {
        principalPaid = normalPrincipal;
        thisPayment = monthlyPayment;
      }
    }

    balance = Math.max(0, balance - principalPaid);
    schedule.push({
      n: n,
      dueDate: Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      payment: Math.round(thisPayment * 100) / 100,
      interest: Math.round(interest * 100) / 100,
      principal: Math.round(principalPaid * 100) / 100,
      balanceAfter: Math.round(balance * 100) / 100
    });
    if (balance <= 0.01) break;
  }
  return schedule;
}

// Scan TLMND_TRANSACTIONS for received payments matching this loan's pattern.
// Returns oldest-first list.
function _matchLoanPayments_(loan) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TLMND_TRANSACTIONS');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var iDate = headers.indexOf('Date');
  var iAccount = headers.indexOf('Account');
  var iName = headers.indexOf('Name');
  var iAmount = headers.indexOf('Amount USD');
  if (iDate < 0 || iAmount < 0 || iName < 0) return [];

  var pattern = String(loan['Plaid Match Pattern'] || '').toLowerCase().trim();
  if (!pattern) return [];

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  var payments = [];
  rows.forEach(function(r) {
    var name = String(r[iName] || '').toLowerCase();
    if (name.indexOf(pattern) < 0) return;
    var d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]);
    if (isNaN(d.getTime())) return;
    var amount = Number(r[iAmount] || 0);
    // Loan repayment income should be positive (money in). Skip outflows.
    if (amount <= 0) return;
    payments.push({
      date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      amount: amount,
      account: String(r[iAccount] || ''),
      name: String(r[iName] || '')
    });
  });
  payments.sort(function(a, b) { return a.date.localeCompare(b.date); });
  return payments;
}

// Web-callable — full status for the dashboard Loans tab. Returns an array
// per active loan with:
//   { loan, schedule (rows w/ received flag + actual date+amount), summary }
function getLoansStatus() {
  var loans = getLoans().filter(function(l){ return String(l.Status||'').toLowerCase() !== 'deleted'; });
  return loans.map(function(loan) {
    var schedule = _generateAmortizationSchedule_(loan);
    var payments = _matchLoanPayments_(loan);

    // Greedy match: payment i → schedule row i. Simple and matches how
    // amortization is meant to work (one payment per period). If payments
    // arrive out of order or partial, we can refine later.
    // Attaches the full Plaid transaction (name, account) on matched rows
    // so the frontend drilldown can show the underlying transaction.
    var scheduleWithStatus = schedule.map(function(row, i) {
      var p = payments[i] || null;
      return {
        n: row.n,
        dueDate: row.dueDate,
        payment: row.payment,
        interest: row.interest,
        principal: row.principal,
        balanceAfter: row.balanceAfter,
        received: !!p,
        actualDate: p ? p.date : null,
        actualAmount: p ? p.amount : null,
        variance: p ? (p.amount - row.payment) : 0,
        txn: p ? { date: p.date, amount: p.amount, account: p.account, name: p.name } : null
      };
    });

    var received = scheduleWithStatus.filter(function(r){ return r.received; });
    var totalReceived = received.reduce(function(s, r){ return s + (r.actualAmount||0); }, 0);
    var totalScheduled = schedule.reduce(function(s, r){ return s + r.payment; }, 0);

    // Current outstanding balance based on payments received so far.
    var currentBalance;
    if (received.length === 0) {
      currentBalance = Number(loan['Effective Principal']) || 0;
    } else if (received.length <= schedule.length) {
      currentBalance = scheduleWithStatus[received.length - 1].balanceAfter;
    } else {
      currentBalance = 0;
    }

    // Next expected payment (first row that isn't yet received).
    var nextDue = scheduleWithStatus.find(function(r){ return !r.received; }) || null;

    // If this loan is linked to an asset, push the current outstanding
    // balance up to that asset row on the Assets sheet so the dashboard
    // stays consistent. This is a light writeback — only fires when the
    // asset's stored value actually differs from what the loan schedule
    // computes, so most Loans-tab loads are read-only.
    var linkedAssetId = String(loan['Linked Asset ID'] || '').trim();
    var linkSync = null;
    if (linkedAssetId) {
      try { linkSync = _syncLinkedAssetBalance_(linkedAssetId, currentBalance, loan['Name']); }
      catch(e) { Logger.log('_syncLinkedAssetBalance_ failed for ' + linkedAssetId + ': ' + e.message); }
    }

    return {
      loan: loan,
      schedule: scheduleWithStatus,
      summary: {
        paymentsReceived:  received.length,
        paymentsScheduled: schedule.length,
        totalReceived:     totalReceived,
        totalScheduled:    totalScheduled,
        currentBalance:    currentBalance,
        nextDue:           nextDue,
        pctPaid:           schedule.length ? (received.length / schedule.length) : 0,
        linkedAssetSync:   linkSync
      }
    };
  });
}

// Update the linked asset's stored balance to match the loan's current
// outstanding. Uses the asset's My Share % to compute the owner's share.
// Returns { updated: bool, oldValue, newValue } for the UI to display.
function _syncLinkedAssetBalance_(assetId, loanBalance, loanName) {
  var sheet = getSheet_('ASSETS');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var iId    = headers.indexOf('ID');
  var iLocal = headers.indexOf('Local Value');
  var iUsd   = headers.indexOf('USD Value');
  var iShare = headers.indexOf('My Share %');
  var iMine  = headers.indexOf('My Share USD');
  var iUpd   = headers.indexOf('Last Updated');
  if (iId < 0 || iMine < 0) return null;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iId]) !== String(assetId)) continue;
    var sharePct = iShare >= 0 ? (Number(data[r][iShare]) || 100) : 100;
    var newLocal = loanBalance;
    var newMine  = loanBalance * sharePct / 100;
    var oldMine  = iMine >= 0 ? (Number(data[r][iMine]) || 0) : 0;
    // Only write if the value materially differs (>$1 to avoid rounding chatter).
    if (Math.abs(newMine - oldMine) < 1) return { updated: false, oldValue: oldMine, newValue: newMine };
    if (iLocal >= 0) sheet.getRange(r + 1, iLocal + 1).setValue(newLocal);
    if (iUsd   >= 0) sheet.getRange(r + 1, iUsd + 1).setValue(newLocal);
    if (iMine  >= 0) sheet.getRange(r + 1, iMine + 1).setValue(newMine);
    if (iUpd   >= 0) sheet.getRange(r + 1, iUpd + 1).setValue(new Date());
    _logAudit_('syncLoanBalance', 'asset', assetId, loanName, 'Synced from loan: ' + newMine.toFixed(2));
    return { updated: true, oldValue: oldMine, newValue: newMine };
  }
  return null;
}

// Diagnostic — shows Amanda what the Plaid matcher is actually seeing for a
// given loan. Menu-callable. Answers: is the January payment missing because
// (a) the sheet doesn't have January data at all, or (b) the January
// transaction has a different name than the loan's Plaid Match Pattern?
function debugLoanMatches() {
  var loans = getLoans();
  if (!loans.length) { SpreadsheetApp.getUi().alert('No loans registered yet.'); return; }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TLMND_TRANSACTIONS');
  if (!sheet || sheet.getLastRow() < 2) { SpreadsheetApp.getUi().alert('TLMND_TRANSACTIONS is empty.'); return; }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var iDate = headers.indexOf('Date');
  var iName = headers.indexOf('Name');
  var iAmount = headers.indexOf('Amount USD');
  var iAccount = headers.indexOf('Account');
  var iNotes = headers.indexOf('Notes');
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  // Overall sheet stats.
  var minDate = null, maxDate = null;
  rows.forEach(function(r){
    var d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]);
    if (isNaN(d.getTime())) return;
    if (!minDate || d < minDate) minDate = d;
    if (!maxDate || d > maxDate) maxDate = d;
  });

  var report = 'TLMND_TRANSACTIONS: ' + rows.length + ' rows\n' +
               'Date range: ' + (minDate ? Utilities.formatDate(minDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '?') +
               ' → ' + (maxDate ? Utilities.formatDate(maxDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '?') + '\n\n';

  loans.forEach(function(loan) {
    var pattern = String(loan['Plaid Match Pattern'] || '').toLowerCase().trim();
    var name = String(loan.Name || '');
    report += '── ' + name + ' ──\n';
    report += 'Loan pattern: "' + loan['Plaid Match Pattern'] + '"\n';
    if (!pattern) { report += '  (no pattern set — cannot match)\n\n'; return; }

    // (1) Broader search: try the first significant word of the pattern
    //     (e.g. "solaris" from "SOLARIS-FL HOLDI") so we catch alternate
    //     spellings like "SOLARIS WIRE" or "SOLARIS PAYMENT".
    var firstWord = (pattern.match(/[a-z]{4,}/) || [''])[0];
    var broadMatches = [], exactMatches = [], excludedMatches = [], negativeMatches = [];
    rows.forEach(function(r) {
      var n = String(r[iName] || '').toLowerCase();
      var amt = Number(r[iAmount] || 0);
      var notes = String(r[iNotes] || '');
      var isExc = notes.indexOf('[EXCLUDED]') >= 0;
      var d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]);
      var ds = isNaN(d.getTime()) ? '?' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var rowSum = ds + '  $' + amt.toFixed(2) + '  ' + String(r[iName] || '').substring(0, 80) +
                   (isExc ? ' [EXCLUDED]' : '');
      if (firstWord && n.indexOf(firstWord) >= 0) {
        broadMatches.push(rowSum);
        if (n.indexOf(pattern) >= 0) {
          if (isExc) excludedMatches.push(rowSum);
          else if (amt <= 0) negativeMatches.push(rowSum);
          else exactMatches.push(rowSum);
        }
      }
    });
    report += 'Broad match ("' + firstWord + '" anywhere): ' + broadMatches.length + ' rows\n';
    report += 'Exact pattern match, matched: ' + exactMatches.length + ' rows\n';
    if (excludedMatches.length) report += 'Exact match but EXCLUDED: ' + excludedMatches.length + ' rows\n';
    if (negativeMatches.length) report += 'Exact match but amount <= 0 (outflow): ' + negativeMatches.length + ' rows\n';
    report += '\nAll broad matches (widest possible):\n';
    broadMatches.forEach(function(m){ report += '  ' + m + '\n'; });
    report += '\n';
  });

  // Log AND alert. Long output → open in the sheet's execution log.
  Logger.log(report);
  var ui = SpreadsheetApp.getUi();
  // Alert dialog has a size limit; truncate if huge.
  var alertText = report.length > 6000 ? report.substring(0, 6000) + '\n\n… (truncated — full report in Executions log)' : report;
  ui.alert('Loan Matcher Diagnostic', alertText, ui.ButtonSet.OK);
}

// Web-callable — for the "Linked Asset" dropdown in the Loans modal.
// Returns Loans Receivable / Promissory Notes / (all if requested) assets.
function getLoanableAssets() {
  var sheet = getSheet_('ASSETS');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var iId    = headers.indexOf('ID');
  var iName  = headers.indexOf('Name');
  var iCat   = headers.indexOf('Category');
  var iEnt   = headers.indexOf('Entity');
  var iShare = headers.indexOf('My Share %');
  var iMine  = headers.indexOf('My Share USD');
  var iArch  = headers.indexOf('Archived');
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var cat = String(data[r][iCat] || '');
    // Only offer Loans Receivable / Promissory Notes / Private Equity as
    // link targets — those are the categories where a loan schedule makes sense.
    if (!/loans receivable|promissory notes|private equity/i.test(cat)) continue;
    var isArch = iArch >= 0 ? String(data[r][iArch] || '').toLowerCase() === 'yes' || data[r][iArch] === true : false;
    if (isArch) continue;
    out.push({
      id: String(data[r][iId] || ''),
      name: String(data[r][iName] || ''),
      category: cat,
      entity: String(data[r][iEnt] || ''),
      sharePct: iShare >= 0 ? (Number(data[r][iShare]) || 100) : 100,
      currentValue: iMine >= 0 ? (Number(data[r][iMine]) || 0) : 0
    });
  }
  out.sort(function(a, b){ return a.name.localeCompare(b.name); });
  return out;
}
