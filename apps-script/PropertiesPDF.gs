/**
 * PropertiesPDF.gs — monthly Dorado property report PDF + email.
 *
 * Mirrors Amanda's manual "Dorado Summary" Excel report that she sends Mike
 * monthly. Pulls from getPropertyMonthlyReport (already live for the web
 * tab), builds HTML in the same modern TLMND/NW aesthetic, converts to PDF
 * via newBlob(..).getAs('application/pdf'), and emails on the 20th of each
 * month covering the PRIOR month (that's when Amanda sends the real one).
 *
 * Entry points:
 *   generateDoradoMonthlyPdf(propertyId, year, month) → { blob, filename, html }
 *   sendDoradoMonthlyPdfTest()                        → menu-callable test send
 *   previewDoradoPdfHtml()                            → menu-callable preview (opens dialog)
 *   monthlyDoradoPdfEmail()                           → trigger handler
 *   installDoradoMonthlyPdfTrigger()                  → menu-callable one-time install
 *   setDoradoPdfRecipient()                           → menu-callable: who gets the email
 */

var _DORADO_PDF_RECIPIENT_KEY = 'DORADO_PDF_RECIPIENT';

// Build the HTML string for a given property + year + month. Called by both
// the PDF generator and the preview-in-browser helper.
function _buildDoradoPdfHtml_(propertyId, year, month) {
  var r = getPropertyMonthlyReport(propertyId, year, month);
  if (!r || r.error) return null;

  var prop   = r.property || {};
  var tiles  = r.tiles || {};
  var op     = r.operational || {};
  var rev    = r.revenue || {};
  var reportDate = Utilities.formatDate(new Date(), 'America/New_York', 'MMMM d, yyyy');

  // Section builders: use the same aggregation the web UI uses so numbers
  // tie perfectly between the two surfaces.
  var directItems = [];
  if (r.directCosts && r.directCosts.almaFee) {
    directItems.push({ label: 'Alma Fees (20%)', amount: r.directCosts.almaFee });
  }
  if (r.directCosts && r.directCosts.subcategories) {
    Object.keys(r.directCosts.subcategories).forEach(function(sub) {
      directItems.push({ label: sub, amount: r.directCosts.subcategories[sub].total });
    });
  }
  var opexItems = [];
  Object.keys(r.operatingExpenses || {}).forEach(function(sub) {
    opexItems.push({ label: sub, amount: r.operatingExpenses[sub].total });
  });
  var debtItems = [];
  Object.keys(r.debtService || {}).forEach(function(sub) {
    debtItems.push({ label: sub, amount: r.debtService[sub].total });
  });

  function fmt(n) {
    var num = Number(n) || 0;
    var neg = num < 0;
    var abs = Math.abs(num);
    var s = '$' + abs.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    return neg ? '(' + s + ')' : s;
  }
  function fmtInt(n) { return Number(n || 0).toLocaleString('en-US'); }
  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  // Row renderer for a titled section (Operating Revenue, Direct Costs, etc.)
  function sectionRows(items) {
    return items.map(function(item, i) {
      var isLast = i === items.length - 1;
      return '<tr' + (isLast ? ' class="last"' : '') + '>' +
        '<td class="lbl">' + esc(item.label) + '</td>' +
        '<td class="amt">' + fmt(item.amount) + '</td>' +
      '</tr>';
    }).join('');
  }

  function sectionHdr(label) {
    return '<tr class="sec"><td colspan="2">' + esc(label) + '</td></tr>';
  }

  // Four header tiles.
  var tileCards = [
    { label: 'Rents Collected',        value: tiles.rentsCollected,        cls: 'pos' },
    { label: 'Total Operating Exp',    value: tiles.totalOperatingExpense, cls: 'neg' },
    { label: 'Net Operating Income',   value: tiles.netOperatingIncome,    cls: (tiles.netOperatingIncome || 0) >= 0 ? 'pos' : 'neg' },
    { label: 'Debt Service',           value: tiles.debtService,           cls: 'neg' }
  ].map(function(t) {
    return '<div class="tile">' +
      '<div class="l">' + esc(t.label) + '</div>' +
      '<div class="v ' + t.cls + '">' + fmt(t.value || 0) + '</div>' +
    '</div>';
  }).join('');

  // Operational metric cells.
  var metricCells = [
    { label: 'Occupancy %',        value: (op.occupancyPct || 0) + '%' },
    { label: 'Reservations',       value: fmtInt(op.reservations) },
    { label: 'Nights Booked',      value: fmtInt(op.nightsBooked) },
    { label: 'Average Daily Rate', value: fmt(op.averageDailyRate || 0) }
  ].map(function(m) {
    return '<div class="metric">' +
      '<div class="l">' + esc(m.label) + '</div>' +
      '<div class="v">' + esc(m.value) + '</div>' +
    '</div>';
  }).join('');

  // Grand totals block (bottom summary that matches Amanda's Excel).
  var directTotal = (r.directCosts && r.directCosts.almaFee ? r.directCosts.almaFee : 0);
  if (r.directCosts && r.directCosts.subcategories) {
    Object.keys(r.directCosts.subcategories).forEach(function(s) {
      directTotal += r.directCosts.subcategories[s].total;
    });
  }
  var opexTotal = 0;
  Object.keys(r.operatingExpenses || {}).forEach(function(s) {
    opexTotal += r.operatingExpenses[s].total;
  });

  var totalRows = [
    { label: 'Rents Collected',      amount: tiles.rentsCollected       },
    { label: 'Direct Costs',         amount: directTotal                },
    { label: 'Operating Expenses',   amount: opexTotal                  },
    { label: 'Net Operating Income', amount: tiles.netOperatingIncome   },
    { label: 'Debt Service',         amount: tiles.debtService          }
  ].map(function(row) {
    return '<tr><td>' + esc(row.label) + '</td><td class="amt">' + fmt(row.amount || 0) + '</td></tr>';
  }).join('');
  var netCashFlowCls = (tiles.netCashFlow || 0) >= 0 ? 'pos' : 'neg';

  // Reservations list — only counting ones (Guest / Alma / Direct / unknown).
  var reservations = (r.reservations || []).filter(function(e) {
    return e.source === 'guest' || e.source === 'alma' || e.source === 'direct' || e.source === 'unknown';
  });
  var reservationsHtml = '';
  if (reservations.length) {
    reservationsHtml = '<div class="section">' +
      '<div class="section-hdr">Reservations This Month</div>' +
      '<table class="resv">' +
        '<thead><tr><th>Source</th><th>Guest</th><th>Check-in</th><th>Check-out</th><th>Nights</th></tr></thead>' +
        '<tbody>' + reservations.map(function(e) {
          return '<tr>' +
            '<td class="src">' + esc(e.source || 'unknown') + '</td>' +
            '<td>' + esc(e.title || '') + '</td>' +
            '<td>' + esc(e.startIso || '') + '</td>' +
            '<td>' + esc(e.endIso || '') + '</td>' +
            '<td class="n">' + fmtInt(e.nights) + '</td>' +
          '</tr>';
        }).join('') + '</tbody>' +
      '</table>' +
    '</div>';
  }

  return '' +
    '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + esc(prop['Name'] || 'Property') + ' — ' + esc(r.monthLabel) + '</title>' +
    '<style>' +
      '@page { size: letter; margin: 0.5in 0.5in 0.6in 0.5in; }' +
      'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;color:#0d2137;margin:0;padding:0;font-size:11px;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +
      '.report-hdr{border-bottom:3px solid #0d2137;padding-bottom:10px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:baseline}' +
      '.report-hdr h1{font-size:20px;margin:0;color:#0d2137;letter-spacing:.3px;font-weight:800}' +
      '.report-hdr .sub{font-size:13px;color:#5f6368;font-weight:600;margin-top:2px}' +
      '.report-hdr .date{font-size:11px;color:#5f6368}' +
      '.tiles{display:table;width:100%;border-spacing:8px 0;margin-bottom:10px;table-layout:fixed}' +
      '.tile{display:table-cell;background:#fff;border:1px solid #d0dae5;border-radius:6px;padding:12px 14px;text-align:center;vertical-align:middle}' +
      '.tile .l{font-size:9px;color:#5f6368;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px;font-weight:600}' +
      '.tile .v{font-size:17px;font-weight:800;font-variant-numeric:tabular-nums;line-height:1.1}' +
      '.hero{background:linear-gradient(135deg,#0d2137,#1a3a5c);color:#fff;padding:16px 22px;border-radius:6px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center}' +
      '.hero .l{font-size:11px;text-transform:uppercase;letter-spacing:1px;font-weight:700;opacity:.85}' +
      '.hero .v{font-size:28px;font-weight:800;font-variant-numeric:tabular-nums;letter-spacing:-.3px}' +
      '.hero .v.pos{color:#4ade80}' +
      '.hero .v.neg{color:#f87171}' +
      '.metrics-hdr{background:#0d2137;color:#fff;padding:7px 14px;font-size:10px;font-weight:700;letter-spacing:.6px;text-transform:uppercase}' +
      '.metrics{display:table;width:100%;border-spacing:1px 0;background:#e8ecf1;margin-bottom:16px;table-layout:fixed}' +
      '.metric{display:table-cell;background:#fff;padding:10px 12px;text-align:center}' +
      '.metric .l{font-size:9px;color:#5f6368;letter-spacing:.5px;text-transform:uppercase;margin-bottom:4px;font-weight:600}' +
      '.metric .v{font-size:16px;font-weight:700;color:#0d2137;font-variant-numeric:tabular-nums}' +
      '.section{margin-top:14px}' +
      '.section-hdr{background:#0d2137;color:#fff;padding:7px 14px;font-size:10px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;border-radius:4px 4px 0 0}' +
      '.section table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #d0dae5;border-top:none;border-radius:0 0 4px 4px}' +
      '.section td{padding:7px 14px;font-size:12px;border-bottom:1px solid #f0f3f7;font-variant-numeric:tabular-nums}' +
      '.section td.lbl{color:#3c4858}' +
      '.section td.amt{text-align:right;font-weight:600;color:#0d2137;width:140px}' +
      '.section tr.last td{border-bottom:none}' +
      '.totals{background:#0d2137;color:#fff;border-radius:6px;padding:12px 16px;margin-top:14px}' +
      '.totals table{width:100%;border-collapse:collapse}' +
      '.totals td{padding:4px 0;font-size:12px;color:#fff;opacity:.9;font-variant-numeric:tabular-nums}' +
      '.totals td.amt{text-align:right;font-weight:600;opacity:1}' +
      '.totals .divider td{border-top:1px solid rgba(255,255,255,.25);padding-top:8px;margin-top:4px;font-weight:700;opacity:1;font-size:13px}' +
      '.totals .ncf td{font-size:15px;font-weight:800;opacity:1;padding-top:10px}' +
      '.totals .ncf td.amt.pos{color:#4ade80}' +
      '.totals .ncf td.amt.neg{color:#f87171}' +
      '.resv{width:100%;border-collapse:collapse;background:#fff;border:1px solid #d0dae5;border-top:none;border-radius:0 0 4px 4px}' +
      '.resv th{background:#f0f3f7;padding:6px 10px;text-align:left;color:#5f6368;font-size:9px;font-weight:600;text-transform:uppercase;letter-spacing:.4px;border-bottom:1px solid #d0dae5}' +
      '.resv th:last-child,.resv td:last-child{text-align:right}' +
      '.resv td{padding:6px 10px;font-size:11px;border-bottom:1px solid #f4f5f7;color:#3c4858}' +
      '.resv td.src{text-transform:uppercase;font-size:9px;font-weight:700;color:#1a73e8;letter-spacing:.3px}' +
      '.resv td.n{font-variant-numeric:tabular-nums;font-weight:600;color:#0d2137}' +
      '.footer{margin-top:18px;padding-top:10px;border-top:1px solid #d0dae5;font-size:9px;color:#8a97a7;text-align:center}' +
      '.pos{color:#1e8e3e}' +
      '.neg{color:#c5221f}' +
    '</style></head><body>' +

    '<div class="report-hdr">' +
      '<div>' +
        '<h1>' + esc(prop['Name'] || 'Property') + '</h1>' +
        '<div class="sub">Monthly Summary · ' + esc(r.monthLabel) + '</div>' +
      '</div>' +
      '<div class="date">Generated ' + esc(reportDate) + '</div>' +
    '</div>' +

    '<div class="tiles">' + tileCards + '</div>' +

    '<div class="hero">' +
      '<div class="l">Net Cash Flow</div>' +
      '<div class="v ' + netCashFlowCls + '">' + fmt(tiles.netCashFlow || 0) + '</div>' +
    '</div>' +

    '<div class="metrics-hdr">Operational Metrics</div>' +
    '<div class="metrics">' + metricCells + '</div>' +

    '<div class="section">' +
      '<div class="section-hdr">Operating Revenue</div>' +
      '<table>' +
        '<tr><td class="lbl">Total Rental Amount</td><td class="amt">' + fmt(rev.grossRentalAmount || 0) + '</td></tr>' +
        '<tr class="last"><td class="lbl">Net Room Revenue</td><td class="amt">' + fmt(rev.netRoomRevenue || 0) + '</td></tr>' +
      '</table>' +
    '</div>' +

    (directItems.length ? (
      '<div class="section">' +
        '<div class="section-hdr">Direct Costs</div>' +
        '<table>' + sectionRows(directItems) + '</table>' +
      '</div>'
    ) : '') +

    (opexItems.length ? (
      '<div class="section">' +
        '<div class="section-hdr">Operating Expenses</div>' +
        '<table>' + sectionRows(opexItems) + '</table>' +
      '</div>'
    ) : '') +

    (debtItems.length ? (
      '<div class="section">' +
        '<div class="section-hdr">Debt Service</div>' +
        '<table>' + sectionRows(debtItems) + '</table>' +
      '</div>'
    ) : '') +

    '<div class="totals">' +
      '<table>' +
        totalRows +
        '<tr class="ncf"><td>Net Cash Flow</td><td class="amt ' + netCashFlowCls + '">' + fmt(tiles.netCashFlow || 0) + '</td></tr>' +
      '</table>' +
    '</div>' +

    reservationsHtml +

    '<div class="footer">Family Office Wealth Tracker · ' + esc(prop['Name'] || '') + ' · ' + esc(r.monthLabel) + '</div>' +

    '</body></html>';
}

// Generate the PDF blob. Returns { blob, filename, html } or null if the
// report couldn't be built.
function generateDoradoMonthlyPdf(propertyId, year, month) {
  var html = _buildDoradoPdfHtml_(propertyId, year, month);
  if (!html) return null;
  var r = getPropertyMonthlyReport(propertyId, year, month);
  var prop = r.property || {};
  var propSlug = String(prop['Name'] || 'property').replace(/[^\w-]+/g, '_').substring(0, 40);
  var monthSlug = r.monthKey;   // 'YYYY-MM'
  var filename = propSlug + '_' + monthSlug + '.pdf';
  var blob = Utilities.newBlob(html, 'text/html', filename.replace(/\.pdf$/, '.html'))
    .getAs('application/pdf')
    .setName(filename);
  return { blob: blob, filename: filename, html: html, report: r };
}

// Convenience: find the Dorado property by name (case-insensitive match on
// "dorado"), compute last-completed month, and build the PDF for it.
function _doradoPdfForLastMonth_() {
  var now = new Date();
  var y = now.getFullYear();
  var m = now.getMonth();        // 0-indexed; prior month = this value as 1-indexed
  if (m === 0) { y--; m = 12; }
  var dorado = getProperties().find(function(p) { return /dorado/i.test(String(p['Name'] || '')); });
  if (!dorado) return null;
  return generateDoradoMonthlyPdf(dorado['ID'], y, m);
}

// Web-callable: send the PDF for a SPECIFIC month to either Mike (testMode=
// false → the configured recipient) or just the current user (testMode=true).
// This is what the "📧 Send Monthly PDF to Mike" / "Test to me" buttons on
// the Property detail call.
function sendDoradoPdfForMonth(propertyId, year, month, testMode) {
  _requireEditor_();
  var recipient;
  if (testMode) {
    recipient = Session.getActiveUser().getEmail() || _currentUserEmail_();
    if (!recipient) return { success: false, error: 'Could not determine your email.' };
  } else {
    recipient = PropertiesService.getScriptProperties().getProperty(_DORADO_PDF_RECIPIENT_KEY);
    if (!recipient) return { success: false, error: 'No recipient set. Open Sheets → Tracker → Properties → Set Dorado PDF Recipient.' };
  }
  var pdf = generateDoradoMonthlyPdf(propertyId, year, month);
  if (!pdf) return { success: false, error: 'Could not build PDF for that property/month.' };
  var monthLabel = pdf.report && pdf.report.monthLabel ? pdf.report.monthLabel : '';
  var monthShort = monthLabel.split(' ')[0];
  var subject = (testMode ? 'TEST — ' : '') + 'Dorado Monthly Report — ' + monthLabel;
  var htmlBody = _buildDoradoEmailHtml_(pdf.report, monthLabel, monthShort);
  var plainBody = _buildDoradoEmailPlain_(pdf.report, monthLabel);
  MailApp.sendEmail({
    to:          recipient,
    subject:     subject,
    body:        plainBody,
    htmlBody:    htmlBody,
    name:        'Dorado Monthly Report',
    attachments: [pdf.blob]
  });
  return { success: true, recipient: recipient, monthLabel: monthLabel };
}

// Send the Dorado PDF to a recipient. Called by the test menu + the cron.
// Body is HTML — matches Amanda's existing hand-formatted email to Mike:
//   greeting · optional reservation note · "Dorado Beach · <Month> Property Report"
//   navy NCF banner with gold amount · 4 operational cards · Financial Summary
//   table · italic footer with reservation detail.
function _sendDoradoMonthlyPdf_(recipient, subjectPrefix) {
  _requireEditor_();
  var r = _doradoPdfForLastMonth_();
  if (!r) return { success: false, error: 'No Dorado property found, or report build failed.' };
  var monthLabel = r.report && r.report.monthLabel ? r.report.monthLabel : '';
  var monthShort = monthLabel.split(' ')[0];   // "July 2026" → "July"
  var subject = (subjectPrefix || 'Dorado Monthly Report') + ' — ' + monthLabel;

  var htmlBody = _buildDoradoEmailHtml_(r.report, monthLabel, monthShort);
  var plainBody = _buildDoradoEmailPlain_(r.report, monthLabel);

  MailApp.sendEmail({
    to:          recipient,
    subject:     subject,
    body:        plainBody,   // fallback for text-only clients
    htmlBody:    htmlBody,
    name:        'Dorado Monthly Report',
    attachments: [r.blob]
  });
  Logger.log('Dorado monthly PDF sent to ' + recipient + ' for ' + monthLabel);
  return { success: true, monthLabel: monthLabel };
}

// HTML email body — matches Amanda's hand-formatted template exactly.
function _buildDoradoEmailHtml_(report, monthLabel, monthShort) {
  var tiles = report.tiles || {};
  var op    = report.operational || {};
  var rev   = report.revenue || {};
  var ncf   = Number(tiles.netCashFlow) || 0;
  var rents = Number(tiles.rentsCollected) || 0;
  var totalOpex = Number(tiles.totalOperatingExpense) || 0;
  var noi   = Number(tiles.netOperatingIncome) || 0;
  var debt  = Number(tiles.debtService) || 0;

  // Days in the reported month (for Occupancy sub-label "X of Y nights").
  var daysInMonth = (function() {
    var y = Number(report.year), m = Number(report.month);
    if (!y || !m) return 0;
    return new Date(y, m, 0).getDate();
  })();
  var nightsBooked = Number(op.nightsBooked) || 0;
  var occupancyPct = Number(op.occupancyPct) || 0;
  // RevPAR = (gross rental revenue) / (days in month). Standard hotel metric.
  var revPar = daysInMonth > 0 ? (rents / daysInMonth) : 0;
  var adr = Number(op.averageDailyRate) || 0;

  // Reservation summary — paid vs non-paid stays. The report.reservations
  // array has every event with its (possibly overridden) source. Only
  // guest/alma/direct count as paid; owner/family/friend are non-paid stays.
  var allEvents = report.reservations || [];
  var paidEvents = allEvents.filter(function(e) {
    return e.source === 'guest' || e.source === 'alma' || e.source === 'direct' || e.source === 'unknown';
  });
  var nonPaidEvents = allEvents.filter(function(e) {
    return e.source === 'owner' || e.source === 'family' || e.source === 'friend';
  });

  // "Note: there were no reservations in July" style line at the top.
  var topNote = '';
  if (!paidEvents.length) {
    topNote = '<p style="margin:0 0 14px 0;font-style:italic;color:#5f6368;font-size:14px">' +
      'Note: there were no paid reservations in ' + _escH_(monthShort) +
    '</p>';
  }

  // Footer italic note — "No paid reservations for July. Brandon & Manuela
  // stayed at the property during the month."
  var footerNote = '';
  function _namesFromEvents(events) {
    var raw = events.map(function(e) {
      // Strip the property prefix and dashes. "Dorado - Brandon" → "Brandon"
      var t = String(e.title || '').replace(/^[^-]+-\s*/, '').trim();
      return t || '(unnamed stay)';
    });
    return _dedupeStayNames_(raw);
  }
  if (!paidEvents.length && nonPaidEvents.length) {
    var names = _namesFromEvents(nonPaidEvents);
    footerNote = '<p style="margin:16px 0 0 0;font-style:italic;color:#5f6368;font-size:13px">' +
      '*No paid reservations for ' + _escH_(monthShort) + '. ' +
      _escH_(_joinWithAnd_(names)) + ' stayed at the property during the month.' +
    '</p>';
  } else if (!paidEvents.length) {
    footerNote = '<p style="margin:16px 0 0 0;font-style:italic;color:#5f6368;font-size:13px">' +
      '*No paid reservations for ' + _escH_(monthShort) + '.' +
    '</p>';
  } else if (nonPaidEvents.length) {
    var names2 = _namesFromEvents(nonPaidEvents);
    footerNote = '<p style="margin:16px 0 0 0;font-style:italic;color:#5f6368;font-size:13px">' +
      '*' + _escH_(_joinWithAnd_(names2)) + ' also stayed at the property during the month (not counted toward occupancy).' +
    '</p>';
  }

  // Format helpers
  function fmt(n) {
    var num = Number(n) || 0, neg = num < 0, abs = Math.abs(num);
    var s = '$' + abs.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    return neg ? '$(' + abs.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + ')' : s;
  }
  function colorFor(n) { return (Number(n) || 0) < 0 ? '#c5221f' : '#1e8e3e'; }

  // 4 metric cards: ADR · Occupancy · RevPAR · Bookings
  var metricCards =
    _emailMetricCard_('ADR', adr > 0 ? fmt(adr) : 'N/A', null) +
    _emailMetricCard_('OCCUPANCY', occupancyPct + '%', nightsBooked + ' of ' + daysInMonth + ' nights') +
    _emailMetricCard_('REVPAR', fmt(revPar), null) +
    _emailMetricCard_('BOOKINGS', (paidEvents.length ? paidEvents.length + ' res / ' : '') + nightsBooked + ' nights', null);

  var navy = '#1a3a5c';
  var gold = '#c8a456';

  // Table rows — Rents / Opex / NOI / Debt / Net Cash Flow (last highlighted)
  function row(label, amount, options) {
    options = options || {};
    var isNegLabel = (Number(amount) || 0) < 0;
    var isNcf = options.isNcf;
    var bg = options.zebra ? '#f6f8fb' : '#ffffff';
    if (isNcf) bg = navy;
    var labelColor = isNcf ? '#ffffff' : '#0d2137';
    var amountColor = isNcf
      ? ((Number(amount) || 0) < 0 ? '#f87171' : '#4ade80')
      : (isNegLabel ? '#c5221f' : '#1e8e3e');
    return '<tr>' +
      '<td style="padding:12px 16px;font-size:14px;font-weight:' + (isNcf ? '700' : '500') + ';color:' + labelColor + ';background:' + bg + ';border-top:1px solid #e8ecf1">' + _escH_(label) + '</td>' +
      '<td style="padding:12px 16px;font-size:14px;font-weight:' + (isNcf ? '800' : '600') + ';color:' + amountColor + ';background:' + bg + ';text-align:right;border-top:1px solid #e8ecf1;font-variant-numeric:tabular-nums">' + fmt(amount) + '</td>' +
    '</tr>';
  }

  return '' +
    '<div style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;max-width:620px;margin:0 auto;padding:20px;color:#0d2137;line-height:1.5;font-size:14px">' +

      '<p style="margin:0 0 14px 0;font-size:14px;color:#0d2137">Hello Dr. Mike,</p>' +
      '<p style="margin:0 0 6px 0;font-size:14px;color:#0d2137">Attached is the ' + _escH_(monthShort) + ' Dorado Report for your review.</p>' +
      topNote +
      '<p style="margin:12px 0;font-size:14px;color:#0d2137">Here\'s a quick overview:</p>' +

      // Dorado Beach · July 2026 Property Report
      '<div style="margin:18px 0 12px 0;padding-bottom:8px;border-bottom:2px solid #0d2137">' +
        '<h2 style="margin:0;font-size:22px;color:#0d2137;font-weight:800">Dorado Beach <span style="font-weight:500;color:#5f6368">· ' + _escH_(monthLabel) + ' Property Report</span></h2>' +
      '</div>' +

      '<p style="margin:0 0 14px 0;font-size:14px;color:#3c4858">Please find below your ' + _escH_(monthLabel) + ' property performance summary.</p>' +

      // NET CASH FLOW banner
      '<div style="background:' + navy + ';border-radius:6px;padding:18px 24px;margin:0 0 24px 0;display:flex;justify-content:space-between;align-items:center;overflow:hidden">' +
        '<table style="width:100%;border-collapse:collapse"><tr>' +
          '<td style="font-size:12px;font-weight:700;color:' + gold + ';letter-spacing:1px;text-transform:uppercase">NET CASH FLOW — ' + _escH_(monthLabel.toUpperCase()) + '</td>' +
          '<td style="font-size:26px;font-weight:800;color:' + gold + ';text-align:right;font-variant-numeric:tabular-nums">' + fmt(ncf) + '</td>' +
        '</tr></table>' +
      '</div>' +

      // Operational Performance
      '<h3 style="margin:16px 0 10px 0;font-size:16px;color:#0d2137;font-weight:700">Operational Performance</h3>' +
      '<table style="width:100%;border-collapse:separate;border-spacing:6px 0;margin-bottom:18px"><tr>' +
        metricCards +
      '</tr></table>' +

      // Financial Summary
      '<h3 style="margin:16px 0 10px 0;font-size:16px;color:#0d2137;font-weight:700">Financial Summary</h3>' +
      '<table style="width:100%;border-collapse:collapse;border:1px solid #d0dae5;border-radius:4px;overflow:hidden">' +
        '<thead><tr>' +
          '<th style="padding:10px 16px;background:' + navy + ';color:#ffffff;font-size:12px;text-align:left;font-weight:700;letter-spacing:.3px">Item</th>' +
          '<th style="padding:10px 16px;background:' + navy + ';color:#ffffff;font-size:12px;text-align:right;font-weight:700;letter-spacing:.3px">Amount</th>' +
        '</tr></thead>' +
        '<tbody>' +
          row('Rents Collected',            rents) +
          row('Total Operating Expenses',   totalOpex,   { zebra: true }) +
          row('Net Operating Income',       noi) +
          row('Debt Service / Mortgage',    debt,        { zebra: true }) +
          row('Net Cash Flow',              ncf,         { isNcf: true }) +
        '</tbody>' +
      '</table>' +

      footerNote +

    '</div>';
}

function _emailMetricCard_(label, value, subLabel) {
  return '<td style="width:25%;padding:12px 10px;background:#ffffff;border:1px solid #d0dae5;border-radius:6px;text-align:center;vertical-align:top">' +
    '<div style="font-size:10px;color:#5f6368;letter-spacing:.5px;text-transform:uppercase;font-weight:600;margin-bottom:6px">' + _escH_(label) + '</div>' +
    '<div style="font-size:18px;font-weight:700;color:#0d2137;line-height:1.2">' + _escH_(value) + '</div>' +
    (subLabel ? '<div style="font-size:10px;color:#5f6368;margin-top:4px">' + _escH_(subLabel) + '</div>' : '') +
  '</td>';
}

function _buildDoradoEmailPlain_(report, monthLabel) {
  var tiles = report.tiles || {};
  function fmt(n) {
    var num = Number(n) || 0;
    return (num < 0 ? '-' : '') + '$' + Math.abs(num).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
  }
  return 'Hello Dr. Mike,\n\n' +
         'Attached is the ' + monthLabel + ' Dorado Report.\n\n' +
         'Net Cash Flow:            ' + fmt(tiles.netCashFlow || 0) + '\n' +
         'Rents Collected:          ' + fmt(tiles.rentsCollected || 0) + '\n' +
         'Total Operating Expenses: ' + fmt(tiles.totalOperatingExpense || 0) + '\n' +
         'Net Operating Income:     ' + fmt(tiles.netOperatingIncome || 0) + '\n' +
         'Debt Service / Mortgage:  ' + fmt(tiles.debtService || 0) + '\n\n' +
         'Full P&L, operational metrics, and reservations are in the attached PDF.';
}

function _escH_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function _joinWithAnd_(arr) {
  if (!arr || !arr.length) return '';
  if (arr.length === 1) return arr[0];
  if (arr.length === 2) return arr[0] + ' & ' + arr[1];
  return arr.slice(0, -1).join(', ') + ', & ' + arr[arr.length - 1];
}

// Collapse owner variants (Mike / Dr. Mike / Dr Michael Nguyen) to one
// canonical "Dr. Mike", then dedupe. Preserves friends with different last
// names (e.g. "Mike Macdonald" stays separate from "Dr. Mike").
var _OWNER_ALIAS_SET = {
  'mike': 1, 'dr mike': 1, 'michael': 1, 'michael nguyen': 1,
  'dr michael': 1, 'dr michael nguyen': 1,
  'nguyen': 1, 'dr nguyen': 1
};
function _dedupeStayNames_(names) {
  var seen = {};
  var out = [];
  (names || []).forEach(function(raw) {
    if (!raw) return;
    // Normalize: lowercase, strip punctuation, collapse spaces
    var low = String(raw).toLowerCase().replace(/[.,'"]/g, '').replace(/\s+/g, ' ').trim();
    var display = _OWNER_ALIAS_SET[low] ? 'Dr. Mike' : raw;
    var key = display.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    out.push(display);
  });
  return out;
}

// Menu-callable: send a TEST PDF to the current user's email.
function sendDoradoMonthlyPdfTest() {
  var ui = SpreadsheetApp.getUi();
  var me = Session.getActiveUser().getEmail();
  if (!me) { ui.alert('Could not determine your email address.'); return; }
  var r = _sendDoradoMonthlyPdf_(me, 'TEST — Dorado Monthly Report');
  if (r.success) ui.alert('Test PDF sent to ' + me + '\n\nCovers: ' + r.monthLabel);
  else ui.alert('Send failed: ' + (r.error || 'unknown'));
}

// Menu-callable: show the PDF HTML in a preview dialog so you can eyeball
// styling before sending. Doesn't produce a PDF, just the source HTML in a
// scrollable modal.
function previewDoradoPdfHtml() {
  var ui = SpreadsheetApp.getUi();
  var r = _doradoPdfForLastMonth_();
  if (!r) { ui.alert('No Dorado property found, or report build failed.'); return; }
  // Wrap in an iframe-free container with inline styles so Apps Script's
  // modal renders it directly.
  var html = HtmlService.createHtmlOutput(r.html)
    .setWidth(900)
    .setHeight(700)
    .setTitle('Dorado Monthly PDF Preview — ' + (r.report && r.report.monthLabel || ''));
  SpreadsheetApp.getUi().showModalDialog(html, 'Dorado Monthly PDF Preview — ' + (r.report && r.report.monthLabel || ''));
}

// Trigger handler — fires on the 20th of each month. Pulls the recipient
// from script properties (set via setDoradoPdfRecipient).
function monthlyDoradoPdfEmail() {
  var recipient = PropertiesService.getScriptProperties().getProperty(_DORADO_PDF_RECIPIENT_KEY);
  if (!recipient) {
    Logger.log('monthlyDoradoPdfEmail: no recipient set — skipping. Set via Tracker → Set Dorado PDF Recipient.');
    return;
  }
  _sendDoradoMonthlyPdf_(recipient, 'Dorado Monthly Report');
}

// Menu-callable: install the monthly trigger. Fires on the 20th at 8 AM ET.
function installDoradoMonthlyPdfTrigger() {
  var ui = SpreadsheetApp.getUi();
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'monthlyDoradoPdfEmail') {
      ScriptApp.deleteTrigger(t); removed++;
    }
  });
  // Monthly triggers fire on a specific day-of-month at a specific hour.
  // Fires on the 20th at 8 AM in the script's timezone.
  ScriptApp.newTrigger('monthlyDoradoPdfEmail')
    .timeBased().onMonthDay(20).atHour(8).create();
  var recipient = PropertiesService.getScriptProperties().getProperty(_DORADO_PDF_RECIPIENT_KEY) || '(not set)';
  ui.alert('Installed Dorado monthly PDF trigger.\n\n' +
           'Fires on the 20th of each month at 8 AM.\n' +
           'Replaced ' + removed + ' prior trigger(s).\n\n' +
           'Current recipient: ' + recipient + '\n' +
           '(Set / change via Tracker → Set Dorado PDF Recipient.)');
}

// Web-callable: read the current recipient(s) for the Properties UI input.
function getDoradoPdfRecipient() {
  return PropertiesService.getScriptProperties().getProperty(_DORADO_PDF_RECIPIENT_KEY) || '';
}

// Web-callable (from the Properties tab input): set the recipient(s).
// Accepts comma-separated multiple addresses. Validates each; rejects the
// whole save if any are malformed.
function setDoradoPdfRecipientFromUi(csv) {
  _requireEditor_();
  var val = String(csv == null ? '' : csv).trim();
  if (!val) {
    PropertiesService.getScriptProperties().deleteProperty(_DORADO_PDF_RECIPIENT_KEY);
    return { success: true, cleared: true };
  }
  var parts = val.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  var bad = parts.filter(function(e) { return !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); });
  if (bad.length) throw new Error('Invalid email(s): ' + bad.join(', '));
  var normalized = parts.join(', ');
  PropertiesService.getScriptProperties().setProperty(_DORADO_PDF_RECIPIENT_KEY, normalized);
  return { success: true, recipient: normalized };
}

// Menu-callable: set the email recipient for the monthly send.
function setDoradoPdfRecipient() {
  var ui = SpreadsheetApp.getUi();
  var cur = PropertiesService.getScriptProperties().getProperty(_DORADO_PDF_RECIPIENT_KEY) || '';
  var r = ui.prompt('Dorado Monthly PDF Recipient',
    'Email to send the Dorado monthly PDF to (goes out on the 20th of each month).\n\nCurrent: ' + (cur || '(not set)') + '\n\nEnter a new email, or leave blank + OK to clear:',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var email = String(r.getResponseText() || '').trim();
  if (!email) {
    PropertiesService.getScriptProperties().deleteProperty(_DORADO_PDF_RECIPIENT_KEY);
    ui.alert('Recipient cleared. Monthly email will skip until a new one is set.');
    return;
  }
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
    ui.alert('That doesn\'t look like a valid email address. Nothing saved.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty(_DORADO_PDF_RECIPIENT_KEY, email);
  ui.alert('Dorado PDF recipient set to:\n' + email);
}
