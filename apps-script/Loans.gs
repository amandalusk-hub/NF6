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

// Menu-callable one-time seed for the Waskar loan (Wasica Holdings, LLC).
// Terms from Amanda's amortization calculator: $793,026.46 principal, 7%,
// 15 years, $7,127.95/mo. Plaid Match Pattern uses pipe-delimited "WASICA
// HOLDINGS|WASKAR" so both book transfers ("BOOK TRANSFER CREDIT B/O: WASICA
// HOLDINGS, LLC…") and wire descriptions ("WASKAR TEJEDA…") are caught.
//
// Also seeds ONE big manual catch-up entry for 04/18/2026 ($220,966.34)
// representing the back-payments covered by that lump wire.
//
// First payment date is a PLACEHOLDER (2023-07-01) — Amanda should edit it
// to the actual start date from the loan docs (Loans tab → Edit → change
// First Payment Date → Save). The schedule will regenerate.
function seedWaskarLoan() {
  _requireEditor_();
  var ui = SpreadsheetApp.getUi();
  var loans = getLoans();
  var existing = loans.filter(function(l){
    var n = String(l.Name||'').toLowerCase();
    return n.indexOf('waskar') >= 0 || n.indexOf('wasica') >= 0;
  })[0];
  if (existing) {
    ui.alert('A Waskar / Wasica loan already exists in the LOANS sheet — no change made.');
    return;
  }
  var resp = ui.alert(
    'Seed Waskar Loan?',
    'Add the Waskar / Wasica Holdings loan with these terms:\n\n' +
    '  Lender:           TLMND LLC\n' +
    '  Borrower:         Waskar Tejeda / Wasica Holdings, LLC\n' +
    '  Agreement date:   June 28, 2023\n' +
    '  Principal:        $793,026.46\n' +
    '  Rate:             7% annual\n' +
    '  Term:             180 months (15 years)\n' +
    '  Monthly payment:  $7,127.95\n' +
    '  First payment:    2023-07-01\n' +
    '  Match pattern:    WASICA HOLDINGS|WASKAR\n\n' +
    'Also adds ONE manual catch-up entry for 04/18/2026:\n' +
    '  Amount:     $220,966.34\n' +
    '  Principal:  $84,746.09\n' +
    '  Interest:   $136,220.25\n' +
    '  Represents back-payments covered by that lump wire.\n\n' +
    'Plaid should then match Feb/Mar/Jul/Aug 2026 monthly wires automatically.\n\n' +
    'Proceed?',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  var res = addLoan({
    name: 'Waskar Loan (Wasica Holdings, LLC)',
    entity: 'TLMND',
    originalPrincipal: 793026.46,
    accruedInterest: 0,
    effectivePrincipal: 793026.46,
    annualRate: 7,
    termMonths: 180,
    firstPaymentDate: '2023-07-01',
    monthlyPayment: 7127.95,
    paymentType: 'End of Period',
    plaidPattern: 'WASICA HOLDINGS|WASKAR',
    status: 'Active',
    notes: 'Loan Agreement dated June 28, 2023. Lender TLMND LLC, Borrower Waskar Tejeda ' +
           '(Wasica Holdings, LLC). Terms from amortization calculator: $793,026.46 @ 7% × ' +
           '15y, $7,127.95/mo. Plaid pattern catches both "BOOK TRANSFER CREDIT B/O: WASICA" ' +
           'and "WASKAR" wires. First payment July 1, 2023.'
  });
  if (!res.success) { ui.alert('Failed to add loan: ' + (res.error || 'unknown')); return; }
  // Add the catch-up manual entry WITH the real principal/interest split so
  // the current-balance calc correctly drops by $84,746 (not just the one
  // schedule row's ~$2,500 expected split).
  addManualLoanPayment(res.id, '2026-04-18', 220966.34,
    'Catch-up wire covering back-payments (Jul 2023 – Apr 2026). Per Amanda\'s Payment Log: ' +
    '$84,746.09 principal + $136,220.25 interest.',
    84746.09,   // principal
    136220.25); // interest
  ui.alert(
    'Waskar loan added.\n\n' +
    'Next steps:\n' +
    '1. Loans tab → Refresh — the schedule shows all 180 months. Feb/Mar/Jul/Aug 2026 rows should auto-match from Plaid.\n' +
    '2. The 04/18/2026 catch-up entry lands on the April 2026 row (blue "✓ Manual" badge).\n' +
    '3. Edit the loan → set Linked Asset if you want the asset balance to auto-sync.'
  );
}

// Menu-callable backfill for Solaris pre-Plaid payments (Jan / Feb / Mar
// 2026). Uses Jan 7 (from Amanda's bank statement) + estimated Feb/Mar
// dates matching the observed Plaid pattern (payments arrive ~2–7 days
// after the 1st). Amanda can edit dates later from the Loans tab by
// removing + re-adding a payment if she has the exact bank statement date.
function seedSolarisMissingPayments() {
  _requireEditor_();
  var ui = SpreadsheetApp.getUi();
  var loans = getLoans();
  var solaris = loans.filter(function(l){ return String(l.Name||'').toLowerCase().indexOf('solaris') >= 0; })[0];
  if (!solaris) { ui.alert('No Solaris loan found. Run "Loans → Init Solaris Loan" first.'); return; }
  var resp = ui.alert(
    'Backfill Solaris payments?',
    'This will add MANUAL payment entries for Jan / Feb / Mar 2026 (payments Plaid doesn\'t have).\n\n' +
    'Amount: $16,655.63 each\n' +
    'Dates: 1/7/2026 (from bank statement), 2/6/2026 (estimated), 3/5/2026 (estimated)\n\n' +
    'You can adjust dates later by removing + re-adding on the Loans tab.\n\nProceed?',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  var seeds = [
    { date: '2026-01-07', amount: 16655.63, notes: 'Backfilled: Jan 2026 payment (pre-Plaid, per bank statement)' },
    { date: '2026-02-06', amount: 16655.63, notes: 'Backfilled: Feb 2026 payment (pre-Plaid, estimated date)' },
    { date: '2026-03-05', amount: 16655.63, notes: 'Backfilled: Mar 2026 payment (pre-Plaid, estimated date)' }
  ];
  var added = 0;
  seeds.forEach(function(s) {
    addManualLoanPayment(solaris.ID, s.date, s.amount, s.notes);
    added++;
  });
  ui.alert('Seeded ' + added + ' backfill payments. Refresh the Loans tab to see them.');
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

// Scan TLMND_TRANSACTIONS for received payments matching this loan's pattern,
// plus any manually-entered payments (from LOAN_MANUAL_PAYMENTS sheet).
// Manual entries are for payments that Plaid doesn't have — e.g. loan
// payments received before the Plaid connection was set up, or into an
// account not synced by Plaid. Merged list is sorted oldest-first.
function _matchLoanPayments_(loan) {
  var payments = [];

  // (1) Plaid-matched payments from TLMND_TRANSACTIONS.
  // The Plaid Match Pattern supports multiple patterns separated by |
  // (e.g. "WASICA HOLDINGS|WASKAR") so loans that arrive under different
  // bank descriptions can all be caught. Match on ANY pattern (OR).
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TLMND_TRANSACTIONS');
  var rawPattern = String(loan['Plaid Match Pattern'] || '').trim();
  var patterns = rawPattern
    ? rawPattern.split('|').map(function(p){ return p.toLowerCase().trim(); }).filter(function(p){ return p.length > 0; })
    : [];
  if (sheet && sheet.getLastRow() >= 2 && patterns.length) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var iDate = headers.indexOf('Date');
    var iAccount = headers.indexOf('Account');
    var iName = headers.indexOf('Name');
    var iAmount = headers.indexOf('Amount USD');
    if (iDate >= 0 && iAmount >= 0 && iName >= 0) {
      var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
      rows.forEach(function(r) {
        var name = String(r[iName] || '').toLowerCase();
        // ANY of the patterns matches → match.
        var hit = false;
        for (var pi = 0; pi < patterns.length; pi++) {
          if (name.indexOf(patterns[pi]) >= 0) { hit = true; break; }
        }
        if (!hit) return;
        var d = r[iDate] instanceof Date ? r[iDate] : new Date(r[iDate]);
        if (isNaN(d.getTime())) return;
        var amount = Number(r[iAmount] || 0);
        if (amount <= 0) return;   // outflow — skip
        payments.push({
          date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          amount: amount,
          account: String(r[iAccount] || ''),
          name: String(r[iName] || ''),
          source: 'plaid'
        });
      });
    }
  }

  // (2) Manual entries for this loan.
  var manualSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOAN_MANUAL_PAYMENTS');
  if (manualSheet && manualSheet.getLastRow() >= 2) {
    var mh = manualSheet.getRange(1, 1, 1, manualSheet.getLastColumn()).getValues()[0];
    var mLoan = mh.indexOf('Loan ID');
    var mDate = mh.indexOf('Date');
    var mAmount = mh.indexOf('Amount');
    var mPrincipal = mh.indexOf('Principal');
    var mInterest = mh.indexOf('Interest');
    var mNotes = mh.indexOf('Notes');
    var mRows = manualSheet.getRange(2, 1, manualSheet.getLastRow() - 1, mh.length).getValues();
    var loanId = String(loan.ID || '');
    mRows.forEach(function(r) {
      if (String(r[mLoan] || '') !== loanId) return;
      var d = r[mDate] instanceof Date ? r[mDate] : new Date(r[mDate]);
      if (isNaN(d.getTime())) return;
      var amount = Number(r[mAmount] || 0);
      if (amount <= 0) return;
      var manualPrincipal = mPrincipal >= 0 && r[mPrincipal] !== '' ? Number(r[mPrincipal]) : null;
      var manualInterest = mInterest >= 0 && r[mInterest] !== '' ? Number(r[mInterest]) : null;
      payments.push({
        date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        amount: amount,
        account: '(manual entry)',
        name: mNotes >= 0 ? (String(r[mNotes] || '') || 'Manual payment entry') : 'Manual payment entry',
        source: 'manual',
        manualPrincipal: manualPrincipal,   // null = use schedule row's expected split
        manualInterest:  manualInterest
      });
    });
  }

  payments.sort(function(a, b) { return a.date.localeCompare(b.date); });
  return payments;
}

// ─── Manual Payments ──────────────────────────────────────────────────────
// For payments Plaid doesn't have (predates the connection, wrong account,
// etc.). Stored in LOAN_MANUAL_PAYMENTS. UI: click a Pending/Overdue row
// in the Loans tab → "Mark as received" → creates a row here.
var LOAN_MANUAL_HEADERS = [
  'ID', 'Loan ID', 'Date', 'Amount', 'Principal', 'Interest', 'Notes', 'Entered By', 'Entered At'
];

function ensureManualPaymentsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('LOAN_MANUAL_PAYMENTS');
  if (!sheet) {
    sheet = ss.insertSheet('LOAN_MANUAL_PAYMENTS');
    sheet.getRange(1, 1, 1, LOAN_MANUAL_HEADERS.length)
      .setValues([LOAN_MANUAL_HEADERS])
      .setFontWeight('bold').setBackground('#14263d').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Auto-add any headers we've since introduced (Principal, Interest were
  // added when Waskar's lump-sum catch-up needed a real principal split).
  var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var missing = LOAN_MANUAL_HEADERS.filter(function(h){ return existing.indexOf(h) < 0; });
  if (missing.length) {
    var startCol = existing.length + 1;
    sheet.getRange(1, startCol, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold').setBackground('#14263d').setFontColor('#ffffff');
  }
  return sheet;
}

// principal + interest are OPTIONAL. When provided (e.g. for a lump-sum
// catch-up that covers many periods at once), they override the schedule
// row's expected split when computing current balance. When omitted, the
// balance formula falls back to the schedule row's expected principal.
function addManualLoanPayment(loanId, dateStr, amount, notes, principal, interest) {
  _requireEditor_();
  if (!loanId || !dateStr || !amount) return { success: false, error: 'Loan ID, date, and amount required.' };
  var sheet = ensureManualPaymentsSheet_();
  var id = 'mp_' + Utilities.getUuid().substring(0, 8);
  var pVal = principal != null && principal !== '' ? Number(principal) : '';
  var iVal = interest  != null && interest  !== '' ? Number(interest)  : '';
  // Row must be built in LOAN_MANUAL_HEADERS order to survive schema drift.
  var row = LOAN_MANUAL_HEADERS.map(function(h){
    switch(h) {
      case 'ID':         return id;
      case 'Loan ID':    return loanId;
      case 'Date':       return dateStr;
      case 'Amount':     return Number(amount) || 0;
      case 'Principal':  return pVal;
      case 'Interest':   return iVal;
      case 'Notes':      return notes || '';
      case 'Entered By': return _currentUserEmail_() || '(unknown)';
      case 'Entered At': return new Date();
      default:           return '';
    }
  });
  sheet.appendRow(row);
  var splitNote = (pVal !== '' && iVal !== '') ? ' (P $' + pVal + ' / I $' + iVal + ')' : '';
  _logAudit_('addManualPayment', 'loan', loanId, '', 'Manual payment: ' + dateStr + ' $' + amount + splitNote + (notes ? ' — ' + notes : ''));
  return { success: true, id: id };
}

function deleteManualLoanPayment(id) {
  _requireEditor_();
  var sheet = ensureManualPaymentsSheet_();
  if (sheet.getLastRow() < 2) return { success: false, error: 'No manual payments.' };
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) {
      _logAudit_('deleteManualPayment', 'loan', vals[i][1], '', 'Removed manual payment: ' + id);
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  return { success: false, error: 'Manual payment not found: ' + id };
}

// Frontend-friendly: remove a manual payment by (loanId, date, amount) so the
// client doesn't need to know the mp_xxx id. Used from the drilldown's
// "Remove Manual Entry" button.
function _removeManualPaymentByLoanAndDate(loanId, dateStr, amount) {
  _requireEditor_();
  var sheet = ensureManualPaymentsSheet_();
  if (sheet.getLastRow() < 2) return { success: false, error: 'No manual payments.' };
  var mh = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var iId = mh.indexOf('ID');
  var iLoan = mh.indexOf('Loan ID');
  var iDate = mh.indexOf('Date');
  var iAmount = mh.indexOf('Amount');
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, mh.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][iLoan]) !== String(loanId)) continue;
    var rowDate = vals[i][iDate] instanceof Date
      ? Utilities.formatDate(vals[i][iDate], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(vals[i][iDate] || '');
    if (rowDate !== String(dateStr || '')) continue;
    if (Math.abs(Number(vals[i][iAmount]) - Number(amount)) > 0.01) continue;
    _logAudit_('deleteManualPayment', 'loan', loanId, '', 'Removed manual payment: ' + dateStr + ' $' + amount);
    sheet.deleteRow(i + 2);
    return { success: true };
  }
  return { success: false, error: 'Manual payment not found for that date/amount.' };
}

// Web-callable — full status for the dashboard Loans tab. Returns an array
// per active loan with:
//   { loan, schedule (rows w/ received flag + actual date+amount), summary }
function getLoansStatus() {
  var loans = getLoans().filter(function(l){ return String(l.Status||'').toLowerCase() !== 'deleted'; });
  return loans.map(function(loan) {
    var schedule = _generateAmortizationSchedule_(loan);
    var payments = _matchLoanPayments_(loan);

    // Month-based matching: each payment goes to the schedule row whose due
    // date is in the same calendar month. Handles late payments, missing
    // months (Amanda's Jan/Feb/Mar Solaris rows stay Pending until she
    // manually enters them), and multiple payments in one month (extras
    // spill to the next open month). Prior naive-greedy-by-index was
    // matching Apr 3 → Jan 1 which was very wrong.
    //
    // Build a month → row index map.
    var monthToRow = {};
    schedule.forEach(function(r, i) {
      var ym = r.dueDate.substring(0, 7);   // YYYY-MM
      if (monthToRow[ym] == null) monthToRow[ym] = i;
    });
    // Match each payment to a schedule row. Preference: same month; if the
    // month is already taken, spill to the next open row after it.
    var matchIdxByPayment = new Array(payments.length).fill(-1);
    var rowTaken = new Array(schedule.length).fill(false);
    payments.forEach(function(p, pi) {
      var ym = p.date.substring(0, 7);
      var start = monthToRow[ym];
      if (start == null) {
        // No schedule row for this month (payment before the first due, or
        // after the last). Fall back to closest un-taken row overall.
        var bestI = -1, bestDelta = Infinity;
        for (var i = 0; i < schedule.length; i++) {
          if (rowTaken[i]) continue;
          var delta = Math.abs(new Date(schedule[i].dueDate + 'T00:00:00').getTime() -
                               new Date(p.date + 'T00:00:00').getTime());
          if (delta < bestDelta) { bestDelta = delta; bestI = i; }
        }
        if (bestI >= 0) { matchIdxByPayment[pi] = bestI; rowTaken[bestI] = true; }
        return;
      }
      // Try same-month row first; if taken, walk forward.
      for (var j = start; j < schedule.length; j++) {
        if (!rowTaken[j]) { matchIdxByPayment[pi] = j; rowTaken[j] = true; return; }
      }
    });
    // Invert: rowIdx → paymentIdx.
    var paymentByRow = {};
    matchIdxByPayment.forEach(function(rowI, payI) {
      if (rowI >= 0) paymentByRow[rowI] = payments[payI];
    });

    var scheduleWithStatus = schedule.map(function(row, i) {
      var p = paymentByRow[i] || null;
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
        source: p ? (p.source || 'plaid') : null,
        // Effective principal reduction for THIS row. Manual entries can
        // override the schedule's expected split (needed for lump-sum
        // catch-ups covering many months of back-payments).
        principalPaid: p
          ? (p.manualPrincipal != null ? p.manualPrincipal : (row.principal || 0))
          : 0,
        txn: p ? { date: p.date, amount: p.amount, account: p.account, name: p.name, source: p.source || 'plaid', manualPrincipal: p.manualPrincipal, manualInterest: p.manualInterest } : null
      };
    });

    var received = scheduleWithStatus.filter(function(r){ return r.received; });
    var totalReceived = received.reduce(function(s, r){ return s + (r.actualAmount||0); }, 0);
    var totalScheduled = schedule.reduce(function(s, r){ return s + r.payment; }, 0);

    // Current outstanding balance = effective principal minus principal
    // actually paid down by received rows. Works whether the received rows
    // are contiguous or scattered (e.g. Apr–Sep matched, Jan–Mar still
    // pending), because we're summing what actually happened rather than
    // trusting the schedule's precomputed balanceAfter (which assumes all
    // prior rows were paid on time).
    // Uses each row's principalPaid (manual override if provided, else the
    // schedule's expected split) so a $220,966 catch-up wire correctly
    // credits its $84,746 principal portion instead of just $2,500.
    var principalPaid = received.reduce(function(s, r){ return s + (r.principalPaid || 0); }, 0);
    var effective = Number(loan['Effective Principal']) || 0;
    var currentBalance = Math.max(0, effective - principalPaid);

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

// Web-callable — returns the loan (if any) linked to a given asset ID.
// Used by the asset detail modal to show a banner directing users to the
// Loans tab instead of maintaining a separate manual Payment Log.
function getLoanLinkedToAsset(assetId) {
  if (!assetId) return null;
  var loans = getLoans();
  var linked = loans.filter(function(l){ return String(l['Linked Asset ID']||'') === String(assetId); })[0];
  if (!linked) return null;
  return {
    loanId: linked.ID,
    loanName: linked.Name,
    plaidPattern: linked['Plaid Match Pattern'],
    monthlyPayment: linked['Monthly Payment']
  };
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
