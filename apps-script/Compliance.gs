/**
 * Compliance.gs — Compliance Calendar: entries + reminders.
 *
 * Mirrors Amanda's existing "Compliance Calendar" Excel but gives it a web
 * UI, auto-computed urgency, and the two reminder cadences she asked for:
 *
 *   1. Daily (7 AM ET): for every item due within [30, 14, 7, 3, 1] days or
 *      overdue, group by Person Responsible, send ONE email per person
 *      listing their items. Amanda + Brandon CC'd on every send.
 *   2. Weekly digest (Mondays, 7:30 AM): single email to Amanda + Brandon
 *      summarizing everything due this week + anything overdue.
 *
 * Sheets:
 *   COMPLIANCE_ITEMS   — one row per obligation
 *   COMPLIANCE_PEOPLE  — name → email lookup (inline so each item can carry
 *                        a compound "Amanda/Brandon" person string and the
 *                        reminder resolves emails from this table)
 *
 * All data stays in the active spreadsheet (same sheet as the rest of the app).
 */

var COMPLIANCE_ITEMS_HEADERS = [
  'ID',
  'Status',                   // 'Not Started' | 'In Progress' | 'Up to Date' | 'Done' | 'Cancelled'
  'Item Type',
  'Category',
  'Location',
  'Entity',
  'Obligation',               // what to do
  'Person Responsible',       // 'Amanda/Brandon' etc. — parsed on reminder
  'Frequency',                // 'Yearly'|'bi-yearly'|'Quarterly'|'Monthly'|'One-time'|'Other'
  'Last Done',                // date — when it was last completed
  'Next Due Date',            // date — hard deadline
  'Description',              // SOP / how to do it
  'Link to Docs',
  'Cost',
  'Payment Method',
  'Point of Contact',
  'Website',
  'Account Number',
  'In Keeper',                // Yes/No/N/A
  'On FO Calendar',           // Yes/No
  'Calendar Event Link',
  'Notes',
  'Created At',
  'Last Updated'
];

var COMPLIANCE_PEOPLE_HEADERS = [
  'Name',                     // 'Amanda' (singular — compound names split on save)
  'Email',
  'Active'
];

// Script property keys for the CC list on reminder emails.
var _COMPLIANCE_CC_KEY = 'COMPLIANCE_CC_EMAILS';


// ── Sheet lifecycle ───────────────────────────────────────────────────────

function ensureComplianceSheets_() {
  _ensureComplianceSheet_('COMPLIANCE_ITEMS', COMPLIANCE_ITEMS_HEADERS);
  _ensureComplianceSheet_('COMPLIANCE_PEOPLE', COMPLIANCE_PEOPLE_HEADERS);
}

function _ensureComplianceSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  // Case/whitespace fallback — if getSheetByName returns null but a sheet
  // with the same case-insensitive trimmed name exists, use that one. This
  // avoids the "A sheet with the name 'X' already exists" error when the
  // sheet exists under a slightly different form.
  if (!sheet) {
    var allSheets = ss.getSheets();
    var lcTrimmed = String(name).toLowerCase().trim();
    for (var i = 0; i < allSheets.length; i++) {
      if (String(allSheets[i].getName()).toLowerCase().trim() === lcTrimmed) {
        sheet = allSheets[i];
        break;
      }
    }
  }
  if (!sheet) {
    try {
      sheet = ss.insertSheet(name);
    } catch (e) {
      // Race / weird-naming fallback: assume it was just created and look it up again
      sheet = ss.getSheetByName(name);
      if (!sheet) throw e;
    }
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold').setBackground('#14263d').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
    return sheet;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var missing = headers.filter(function(h) { return existing.indexOf(h) < 0; });
  if (missing.length) {
    var start = existing.length + 1;
    sheet.getRange(1, start, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold').setBackground('#14263d').setFontColor('#ffffff');
  }
  return sheet;
}

function _getComplianceRows_(sheetName, headers) {
  ensureComplianceSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), headers.length)).getValues();
  var hdr = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h) {
      var idx = hdr.indexOf(h);
      obj[h] = idx >= 0 ? row[idx] : '';
    });
    return obj;
  }).filter(function(o) {
    // Items keyed by ID; people keyed by Name
    return o.ID || o.Name;
  });
}

function _writeComplianceRow_(sheetName, headers, obj) {
  ensureComplianceSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var lastCol = Math.max(sheet.getLastColumn(), headers.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = hdr.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.appendRow(row);
}


// ── Status + urgency computation ──────────────────────────────────────────

// Returns a urgency bucket + numeric days-away for an item. Pure function —
// safe to call anywhere, no sheet reads.
function _complianceUrgency_(item, nowIso) {
  var today = nowIso || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var status = String(item.Status || '').toLowerCase();
  if (status === 'cancelled' || status === 'done') {
    return { bucket: status === 'done' ? 'done' : 'cancelled', daysAway: null, isActive: false };
  }
  var dueRaw = item['Next Due Date'];
  if (!dueRaw) return { bucket: 'no-date', daysAway: null, isActive: true };
  // Normalize date to yyyy-MM-dd for compare
  var dueIso;
  if (dueRaw instanceof Date) {
    dueIso = Utilities.formatDate(dueRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } else {
    // Allow strings like 'N/A' — treat as no date
    var s = String(dueRaw).trim();
    if (!/^\d{4}-\d{2}-\d{2}/.test(s)) return { bucket: 'no-date', daysAway: null, isActive: true };
    dueIso = s.substring(0, 10);
  }
  // Simple day diff (date math via ms epoch for safety)
  var dueMs = new Date(dueIso + 'T00:00:00Z').getTime();
  var todMs = new Date(today + 'T00:00:00Z').getTime();
  var daysAway = Math.round((dueMs - todMs) / 86400000);
  var bucket;
  if (daysAway < 0) bucket = 'overdue';
  else if (daysAway <= 7) bucket = 'critical';   // red
  else if (daysAway <= 30) bucket = 'due-soon';   // orange
  else bucket = 'ok';                              // green
  return { bucket: bucket, daysAway: daysAway, isActive: true, dueIso: dueIso };
}


// ── Public API ────────────────────────────────────────────────────────────

// Returns every item + its computed urgency. Web-callable.
function getComplianceItems() {
  var items = _getComplianceRows_('COMPLIANCE_ITEMS', COMPLIANCE_ITEMS_HEADERS);
  var out = items.map(function(it) {
    var u = _complianceUrgency_(it);
    // Normalize dates to ISO strings for client
    ['Last Done', 'Next Due Date', 'Created At', 'Last Updated'].forEach(function(k) {
      if (it[k] instanceof Date) {
        it[k] = Utilities.formatDate(it[k], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
    });
    it._urgency = u;
    return it;
  });
  return JSON.parse(JSON.stringify({ items: out, generatedAt: new Date().toISOString() }));
}

function getCompliancePeople() {
  var ppl = _getComplianceRows_('COMPLIANCE_PEOPLE', COMPLIANCE_PEOPLE_HEADERS);
  // Sort by name
  ppl.sort(function(a, b) { return String(a.Name).localeCompare(String(b.Name)); });
  return JSON.parse(JSON.stringify(ppl));
}

function upsertCompliancePerson(name, email, active) {
  _requireEditor_();
  ensureComplianceSheets_();
  name  = String(name  || '').trim();
  email = String(email || '').trim();
  if (!name) throw new Error('Name required');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COMPLIANCE_PEOPLE');
  var lastCol = Math.max(sheet.getLastColumn(), COMPLIANCE_PEOPLE_HEADERS.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var iName = hdr.indexOf('Name');
  var iEmail = hdr.indexOf('Email');
  var iActive = hdr.indexOf('Active');
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    for (var r = 0; r < data.length; r++) {
      if (String(data[r][iName]).toLowerCase() === name.toLowerCase()) {
        sheet.getRange(r + 2, iEmail + 1).setValue(email);
        if (iActive >= 0) sheet.getRange(r + 2, iActive + 1).setValue(active === false ? 'No' : 'Yes');
        return { success: true, updated: true };
      }
    }
  }
  _writeComplianceRow_('COMPLIANCE_PEOPLE', COMPLIANCE_PEOPLE_HEADERS, {
    'Name': name, 'Email': email, 'Active': active === false ? 'No' : 'Yes'
  });
  return { success: true, updated: false };
}

// Create or update a compliance item. If data.ID is set, updates in place;
// else creates a new row with a generated ID.
function upsertComplianceItem(data) {
  _requireEditor_();
  ensureComplianceSheets_();
  var now = new Date();
  data = data || {};
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COMPLIANCE_ITEMS');
  var lastCol = Math.max(sheet.getLastColumn(), COMPLIANCE_ITEMS_HEADERS.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var iId = hdr.indexOf('ID');

  // Build the write row (headers → values from data)
  function val(key) {
    if (data[key] === undefined) return '';
    // Dates come in as 'yyyy-MM-dd' strings from the client; coerce
    if ((key === 'Last Done' || key === 'Next Due Date') && typeof data[key] === 'string') {
      var s = data[key].trim();
      if (!s || /^n\/a$/i.test(s)) return '';
      // Store as Date so Google Sheets handles sorting + the daily cron
      // can read it back cleanly. Parse as UTC-noon to avoid tz slip.
      var d = new Date(s + 'T12:00:00Z');
      return isNaN(d.getTime()) ? s : d;
    }
    return data[key];
  }

  if (data.ID) {
    // Update existing by ID
    if (sheet.getLastRow() < 2) throw new Error('No items yet');
    var block = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol);
    var vals  = block.getValues();
    for (var r = 0; r < vals.length; r++) {
      if (String(vals[r][iId]) === String(data.ID)) {
        hdr.forEach(function(h, i) {
          if (!h) return;
          if (h === 'ID') return;    // never overwrite ID
          if (h === 'Created At') return;   // keep original
          if (h === 'Last Updated') { vals[r][i] = now; return; }
          if (data.hasOwnProperty(h)) vals[r][i] = val(h);
        });
        block.setValues(vals);
        return { success: true, id: data.ID, updated: true };
      }
    }
    throw new Error('Item not found: ' + data.ID);
  }

  // Create new
  var id = 'ci-' + _complianceNextId_();
  var row = {};
  hdr.forEach(function(h) {
    if (!h) return;
    if (h === 'ID') row[h] = id;
    else if (h === 'Created At') row[h] = now;
    else if (h === 'Last Updated') row[h] = now;
    else if (data.hasOwnProperty(h)) row[h] = val(h);
    else row[h] = '';
  });
  _writeComplianceRow_('COMPLIANCE_ITEMS', COMPLIANCE_ITEMS_HEADERS, row);
  return { success: true, id: id, updated: false };
}

function _complianceNextId_() {
  var items = _getComplianceRows_('COMPLIANCE_ITEMS', COMPLIANCE_ITEMS_HEADERS);
  var max = 0;
  items.forEach(function(i) {
    var m = String(i.ID || '').match(/^ci-(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  var n = String(max + 1);
  while (n.length < 5) n = '0' + n;
  return n;
}

// Mark an item done TODAY. Advances Next Due Date by its Frequency if
// recurring, else marks Status=Done.
function markComplianceItemDone(id) {
  _requireEditor_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COMPLIANCE_ITEMS');
  if (!sheet || sheet.getLastRow() < 2) throw new Error('No items yet');
  var lastCol = Math.max(sheet.getLastColumn(), COMPLIANCE_ITEMS_HEADERS.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var iId = hdr.indexOf('ID');
  var iStatus = hdr.indexOf('Status');
  var iLastDone = hdr.indexOf('Last Done');
  var iNextDue = hdr.indexOf('Next Due Date');
  var iFreq = hdr.indexOf('Frequency');
  var iUpd = hdr.indexOf('Last Updated');
  var block = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol);
  var vals = block.getValues();
  for (var r = 0; r < vals.length; r++) {
    if (String(vals[r][iId]) !== String(id)) continue;
    var today = new Date();
    vals[r][iLastDone] = today;
    var freq = String(vals[r][iFreq] || '').toLowerCase();
    var next = _complianceNextDue_(today, freq);
    if (next) {
      vals[r][iNextDue] = next;
      vals[r][iStatus] = 'Up to Date';
    } else {
      vals[r][iStatus] = 'Done';
    }
    if (iUpd >= 0) vals[r][iUpd] = today;
    block.setValues(vals);
    return { success: true, nextDue: next ? Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM-dd') : null };
  }
  throw new Error('Item not found: ' + id);
}

// Given a last-done date and a frequency string, return the next due Date.
// Returns null for one-time items that shouldn't recur.
function _complianceNextDue_(lastDone, freqLower) {
  var d = new Date(lastDone.getTime());
  switch (freqLower) {
    case 'monthly':   d.setMonth(d.getMonth() + 1); return d;
    case 'quarterly': d.setMonth(d.getMonth() + 3); return d;
    case 'bi-yearly':
    case 'biyearly':
    case 'semi-annual':
    case 'semiannual': d.setMonth(d.getMonth() + 6); return d;
    case 'yearly':
    case 'annual':    d.setFullYear(d.getFullYear() + 1); return d;
    case 'other':
    case 'one-time':
    case 'onetime':
    case '':
      return null;
    default:
      // Unknown — bump by a year as a sensible default
      d.setFullYear(d.getFullYear() + 1);
      return d;
  }
}

function deleteComplianceItem(id) {
  _requireEditor_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COMPLIANCE_ITEMS');
  if (!sheet || sheet.getLastRow() < 2) return { success: false };
  var lastCol = Math.max(sheet.getLastColumn(), COMPLIANCE_ITEMS_HEADERS.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var iId = hdr.indexOf('ID');
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  for (var r = data.length - 1; r >= 0; r--) {
    if (String(data[r][iId]) === String(id)) {
      sheet.deleteRow(r + 2);
      return { success: true };
    }
  }
  return { success: false };
}


// ── Reminder windows + email delivery ─────────────────────────────────────
// Default policy: send a reminder when days-away is EXACTLY in the set
// {30, 14, 7, 3, 1} OR when the item is overdue (any day >= 1 day late).
// This keeps the owner inbox-manageable while surfacing items a few times
// before they slip.

var _COMPLIANCE_REMINDER_WINDOWS = [30, 14, 7, 3, 1];

function _complianceShouldRemindToday_(daysAway) {
  if (daysAway == null) return false;
  if (daysAway < 0) return true;    // overdue: nag daily
  return _COMPLIANCE_REMINDER_WINDOWS.indexOf(daysAway) >= 0;
}

// Parse a "Amanda/Brandon" Person Responsible string into [{name, email}].
// Resolves each name via the COMPLIANCE_PEOPLE sheet. Entries without a
// known email are skipped (logged) so the cron doesn't crash on a typo.
function _complianceResolvePeople_(responsibleStr, peopleMap) {
  var raw = String(responsibleStr || '').trim();
  if (!raw) return [];
  var parts = raw.split(/[\/,&]+/).map(function(s) { return s.trim(); }).filter(Boolean);
  var out = [];
  parts.forEach(function(name) {
    var key = name.toLowerCase();
    var email = peopleMap[key];
    if (email) out.push({ name: name, email: email });
    else Logger.log('Compliance: no email on file for "' + name + '"');
  });
  return out;
}

// Daily cron handler. Fires 7 AM ET weekdays.
function dailyComplianceReminders() {
  ensureComplianceSheets_();
  var items = _getComplianceRows_('COMPLIANCE_ITEMS', COMPLIANCE_ITEMS_HEADERS);
  var people = _getComplianceRows_('COMPLIANCE_PEOPLE', COMPLIANCE_PEOPLE_HEADERS);
  var peopleMap = {};
  people.forEach(function(p) {
    var active = String(p.Active || 'Yes').toLowerCase();
    if (active === 'no' || active === 'false') return;
    if (p.Email) peopleMap[String(p.Name).toLowerCase()] = p.Email;
  });

  // Group items by resolved email
  var byEmail = {};   // email → { name, items: [] }
  items.forEach(function(it) {
    var u = _complianceUrgency_(it);
    if (!u.isActive) return;
    if (!_complianceShouldRemindToday_(u.daysAway)) return;
    var people = _complianceResolvePeople_(it['Person Responsible'], peopleMap);
    people.forEach(function(p) {
      var key = p.email.toLowerCase();
      byEmail[key] = byEmail[key] || { name: p.name, email: p.email, items: [] };
      byEmail[key].items.push(Object.assign({}, it, { _urgency: u }));
    });
  });

  var ccList = _complianceCcList_();
  var sent = 0;
  Object.keys(byEmail).forEach(function(k) {
    var pkg = byEmail[k];
    var html = _buildComplianceReminderHtml_(pkg.name, pkg.items);
    MailApp.sendEmail({
      to:       pkg.email,
      cc:       ccList,
      subject:  'Compliance Reminders — ' + pkg.items.length + ' item' + (pkg.items.length === 1 ? '' : 's'),
      htmlBody: html,
      body:     _buildComplianceReminderPlain_(pkg.name, pkg.items),
      name:     'Family Office Compliance'
    });
    sent++;
    Logger.log('Compliance reminder sent to ' + pkg.email + ' (' + pkg.items.length + ' items)');
  });
  return { sent: sent, buckets: Object.keys(byEmail).length };
}

// Weekly digest to Amanda + Brandon every Monday 7:30 AM ET.
function weeklyComplianceDigest() {
  ensureComplianceSheets_();
  var items = _getComplianceRows_('COMPLIANCE_ITEMS', COMPLIANCE_ITEMS_HEADERS);
  var dueThisWeek = [];
  var overdue = [];
  items.forEach(function(it) {
    var u = _complianceUrgency_(it);
    if (!u.isActive) return;
    if (u.bucket === 'overdue') overdue.push(Object.assign({}, it, { _urgency: u }));
    else if (u.daysAway != null && u.daysAway <= 7) dueThisWeek.push(Object.assign({}, it, { _urgency: u }));
  });
  var ccList = _complianceCcList_();
  if (!ccList) {
    Logger.log('Weekly digest: no CC addresses configured — skipping.');
    return { sent: 0 };
  }
  var html = _buildComplianceDigestHtml_(dueThisWeek, overdue);
  MailApp.sendEmail({
    to:       ccList,
    subject:  'Compliance Digest — ' + overdue.length + ' overdue / ' + dueThisWeek.length + ' due this week',
    htmlBody: html,
    body:     _buildComplianceDigestPlain_(dueThisWeek, overdue),
    name:     'Family Office Compliance'
  });
  return { sent: 1, overdue: overdue.length, dueThisWeek: dueThisWeek.length };
}

function _complianceCcList_() {
  return PropertiesService.getScriptProperties().getProperty(_COMPLIANCE_CC_KEY) || '';
}

function setComplianceCcList(csv) {
  _requireEditor_();
  var val = String(csv || '').trim();
  if (!val) { PropertiesService.getScriptProperties().deleteProperty(_COMPLIANCE_CC_KEY); return { success: true, cleared: true }; }
  var parts = val.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  var bad = parts.filter(function(e) { return !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); });
  if (bad.length) throw new Error('Invalid email(s): ' + bad.join(', '));
  var normalized = parts.join(', ');
  PropertiesService.getScriptProperties().setProperty(_COMPLIANCE_CC_KEY, normalized);
  return { success: true, cc: normalized };
}

function getComplianceCcList() { return _complianceCcList_(); }


// ── Email HTML builders ───────────────────────────────────────────────────

function _complianceEscH_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function _complianceUrgencyBadge_(bucket, daysAway) {
  var colorBg, colorFg, label;
  if (bucket === 'overdue') { colorBg = '#c5221f'; colorFg = '#fff'; label = Math.abs(daysAway) + 'd OVERDUE'; }
  else if (daysAway === 0)   { colorBg = '#c5221f'; colorFg = '#fff'; label = 'DUE TODAY'; }
  else if (daysAway === 1)   { colorBg = '#d93025'; colorFg = '#fff'; label = 'DUE TOMORROW'; }
  else if (bucket === 'critical') { colorBg = '#ea8600'; colorFg = '#fff'; label = 'DUE IN ' + daysAway + 'd'; }
  else if (bucket === 'due-soon') { colorBg = '#f9ab00'; colorFg = '#0d2137'; label = 'DUE IN ' + daysAway + 'd'; }
  else                        { colorBg = '#1e8e3e'; colorFg = '#fff'; label = daysAway + 'd AWAY'; }
  return '<span style="display:inline-block;background:' + colorBg + ';color:' + colorFg + ';padding:3px 8px;border-radius:3px;font-size:10px;font-weight:700;letter-spacing:.3px">' + _complianceEscH_(label) + '</span>';
}

function _complianceGroupByUrgency_(items) {
  var groups = { overdue: [], critical: [], 'due-soon': [], ok: [] };
  items.forEach(function(i) { (groups[i._urgency.bucket] || groups.ok).push(i); });
  // Sort each group by daysAway ascending (closest first; overdue most-overdue first)
  function sortBy(arr, invert) {
    arr.sort(function(a, b) {
      var ad = a._urgency.daysAway == null ? 999999 : a._urgency.daysAway;
      var bd = b._urgency.daysAway == null ? 999999 : b._urgency.daysAway;
      return invert ? (ad - bd) : (ad - bd);
    });
  }
  sortBy(groups.overdue, false);
  sortBy(groups.critical);
  sortBy(groups['due-soon']);
  sortBy(groups.ok);
  return groups;
}

function _buildComplianceReminderHtml_(personName, items) {
  var groups = _complianceGroupByUrgency_(items);
  function section(title, color, arr) {
    if (!arr.length) return '';
    return '<div style="margin:14px 0 6px 0;background:' + color + ';color:#fff;padding:7px 12px;border-radius:4px 4px 0 0;font-size:11px;font-weight:700;letter-spacing:.5px;text-transform:uppercase">' + title + ' (' + arr.length + ')</div>' +
      '<table style="width:100%;border-collapse:collapse;border:1px solid #d0dae5;border-top:none;border-radius:0 0 4px 4px;background:#fff">' +
        '<thead><tr>' +
          '<th style="padding:7px 12px;text-align:left;font-size:10px;color:#5f6368;background:#f0f3f7;border-bottom:1px solid #d0dae5;letter-spacing:.3px;text-transform:uppercase">When</th>' +
          '<th style="padding:7px 12px;text-align:left;font-size:10px;color:#5f6368;background:#f0f3f7;border-bottom:1px solid #d0dae5;letter-spacing:.3px;text-transform:uppercase">Obligation</th>' +
          '<th style="padding:7px 12px;text-align:left;font-size:10px;color:#5f6368;background:#f0f3f7;border-bottom:1px solid #d0dae5;letter-spacing:.3px;text-transform:uppercase">Entity</th>' +
          '<th style="padding:7px 12px;text-align:left;font-size:10px;color:#5f6368;background:#f0f3f7;border-bottom:1px solid #d0dae5;letter-spacing:.3px;text-transform:uppercase">Due</th>' +
        '</tr></thead>' +
        '<tbody>' + arr.map(function(it) {
          var dueIso = it._urgency.dueIso || '';
          var link = String(it['Link to Docs'] || '').trim();
          var obligation = _complianceEscH_(it.Obligation || '(no obligation set)');
          if (link) obligation = '<a href="' + _complianceEscH_(link) + '" style="color:#1a73e8;text-decoration:none">' + obligation + '</a>';
          return '<tr>' +
            '<td style="padding:9px 12px;font-size:11px;border-top:1px solid #f0f3f7;vertical-align:top;white-space:nowrap">' + _complianceUrgencyBadge_(it._urgency.bucket, it._urgency.daysAway) + '</td>' +
            '<td style="padding:9px 12px;font-size:12px;border-top:1px solid #f0f3f7;color:#0d2137;line-height:1.4">' + obligation +
              (it.Notes ? '<div style="color:#5f6368;font-size:11px;margin-top:3px">' + _complianceEscH_(String(it.Notes).substring(0, 160)) + '</div>' : '') +
            '</td>' +
            '<td style="padding:9px 12px;font-size:11px;border-top:1px solid #f0f3f7;color:#5f6368;vertical-align:top;white-space:nowrap">' + _complianceEscH_(it.Entity || '') + '</td>' +
            '<td style="padding:9px 12px;font-size:11px;border-top:1px solid #f0f3f7;color:#0d2137;vertical-align:top;white-space:nowrap;font-variant-numeric:tabular-nums">' + _complianceEscH_(dueIso) + '</td>' +
          '</tr>';
        }).join('') + '</tbody>' +
      '</table>';
  }
  return '<div style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;max-width:720px;margin:0 auto;padding:20px;color:#0d2137;font-size:14px">' +
    '<p style="margin:0 0 10px 0">Hi ' + _complianceEscH_(personName) + ',</p>' +
    '<p style="margin:0 0 16px 0">You have <strong>' + items.length + '</strong> compliance item' + (items.length === 1 ? '' : 's') + ' needing attention. Tap any obligation to open its docs.</p>' +
    section('Overdue',     '#c5221f', groups.overdue) +
    section('Due in ≤7 days',  '#ea8600', groups.critical) +
    section('Due in 8–30 days','#f9ab00', groups['due-soon']) +
    '<p style="margin:18px 0 0 0;font-size:11px;color:#8a97a7;font-style:italic">Full details + edit in the Compliance Calendar tab on the Family Office Tracker. CC\'d: Amanda + Brandon.</p>' +
  '</div>';
}

function _buildComplianceReminderPlain_(name, items) {
  return 'Hi ' + name + ',\n\n' +
    'You have ' + items.length + ' compliance item(s) needing attention:\n\n' +
    items.map(function(it) {
      var u = it._urgency;
      var label = u.bucket === 'overdue' ? Math.abs(u.daysAway) + 'd OVERDUE'
                : u.daysAway === 0 ? 'DUE TODAY'
                : u.daysAway === 1 ? 'DUE TOMORROW'
                : 'Due in ' + u.daysAway + 'd';
      return '  [' + label + '] ' + (it.Obligation || '(no obligation)') + ' — ' + (it.Entity || '') + ' — ' + (u.dueIso || '');
    }).join('\n') + '\n\nFull details in the Family Office Tracker Compliance Calendar.';
}

function _buildComplianceDigestHtml_(dueThisWeek, overdue) {
  var all = overdue.concat(dueThisWeek);
  var groups = _complianceGroupByUrgency_(all);
  function section(title, color, arr) {
    if (!arr.length) return '';
    return '<div style="margin:14px 0 6px 0;background:' + color + ';color:#fff;padding:7px 12px;border-radius:4px 4px 0 0;font-size:11px;font-weight:700;letter-spacing:.5px;text-transform:uppercase">' + title + ' (' + arr.length + ')</div>' +
      '<table style="width:100%;border-collapse:collapse;border:1px solid #d0dae5;border-top:none;border-radius:0 0 4px 4px;background:#fff">' +
        '<thead><tr>' +
          '<th style="padding:7px 12px;text-align:left;font-size:10px;color:#5f6368;background:#f0f3f7;border-bottom:1px solid #d0dae5;letter-spacing:.3px;text-transform:uppercase">When</th>' +
          '<th style="padding:7px 12px;text-align:left;font-size:10px;color:#5f6368;background:#f0f3f7;border-bottom:1px solid #d0dae5;letter-spacing:.3px;text-transform:uppercase">Who</th>' +
          '<th style="padding:7px 12px;text-align:left;font-size:10px;color:#5f6368;background:#f0f3f7;border-bottom:1px solid #d0dae5;letter-spacing:.3px;text-transform:uppercase">Obligation</th>' +
          '<th style="padding:7px 12px;text-align:left;font-size:10px;color:#5f6368;background:#f0f3f7;border-bottom:1px solid #d0dae5;letter-spacing:.3px;text-transform:uppercase">Entity</th>' +
          '<th style="padding:7px 12px;text-align:left;font-size:10px;color:#5f6368;background:#f0f3f7;border-bottom:1px solid #d0dae5;letter-spacing:.3px;text-transform:uppercase">Due</th>' +
        '</tr></thead>' +
        '<tbody>' + arr.map(function(it) {
          var dueIso = it._urgency.dueIso || '';
          var link = String(it['Link to Docs'] || '').trim();
          var obligation = _complianceEscH_(it.Obligation || '(no obligation set)');
          if (link) obligation = '<a href="' + _complianceEscH_(link) + '" style="color:#1a73e8;text-decoration:none">' + obligation + '</a>';
          return '<tr>' +
            '<td style="padding:9px 12px;font-size:11px;border-top:1px solid #f0f3f7;vertical-align:top;white-space:nowrap">' + _complianceUrgencyBadge_(it._urgency.bucket, it._urgency.daysAway) + '</td>' +
            '<td style="padding:9px 12px;font-size:11px;border-top:1px solid #f0f3f7;color:#5f6368;vertical-align:top">' + _complianceEscH_(it['Person Responsible'] || '') + '</td>' +
            '<td style="padding:9px 12px;font-size:12px;border-top:1px solid #f0f3f7;color:#0d2137;line-height:1.4">' + obligation + '</td>' +
            '<td style="padding:9px 12px;font-size:11px;border-top:1px solid #f0f3f7;color:#5f6368;vertical-align:top;white-space:nowrap">' + _complianceEscH_(it.Entity || '') + '</td>' +
            '<td style="padding:9px 12px;font-size:11px;border-top:1px solid #f0f3f7;color:#0d2137;vertical-align:top;white-space:nowrap;font-variant-numeric:tabular-nums">' + _complianceEscH_(dueIso) + '</td>' +
          '</tr>';
        }).join('') + '</tbody>' +
      '</table>';
  }
  return '<div style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;max-width:860px;margin:0 auto;padding:20px;color:#0d2137;font-size:14px">' +
    '<h2 style="margin:0 0 14px 0;font-size:18px;color:#0d2137">Compliance — Weekly Digest</h2>' +
    '<p style="margin:0 0 14px 0;color:#5f6368">' + overdue.length + ' overdue · ' + dueThisWeek.length + ' due in the next 7 days</p>' +
    section('Overdue', '#c5221f', groups.overdue) +
    section('Due in ≤7 days', '#ea8600', groups.critical) +
    (groups['due-soon'].length ? section('Due in 8–30 days', '#f9ab00', groups['due-soon']) : '') +
    '<p style="margin:18px 0 0 0;font-size:11px;color:#8a97a7;font-style:italic">Weekly Monday digest. Full list + edit in the Compliance Calendar tab.</p>' +
  '</div>';
}

function _buildComplianceDigestPlain_(dueThisWeek, overdue) {
  var lines = ['Compliance Weekly Digest', '', overdue.length + ' overdue, ' + dueThisWeek.length + ' due this week', ''];
  function fmt(it) {
    var u = it._urgency;
    var label = u.bucket === 'overdue' ? Math.abs(u.daysAway) + 'd OVERDUE'
              : 'Due in ' + u.daysAway + 'd';
    return '  [' + label + '] ' + (it.Obligation || '(no obligation)') + ' — ' + (it['Person Responsible'] || '') + ' — ' + (it.Entity || '') + ' — ' + (u.dueIso || '');
  }
  if (overdue.length) { lines.push('OVERDUE:'); overdue.forEach(function(i) { lines.push(fmt(i)); }); lines.push(''); }
  if (dueThisWeek.length) { lines.push('DUE THIS WEEK:'); dueThisWeek.forEach(function(i) { lines.push(fmt(i)); }); }
  return lines.join('\n');
}


// ── Trigger install ───────────────────────────────────────────────────────

function installComplianceReminderTriggers() {
  var ui = SpreadsheetApp.getUi();
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'dailyComplianceReminders' || fn === 'weeklyComplianceDigest') {
      ScriptApp.deleteTrigger(t); removed++;
    }
  });
  // Daily at 7 AM (every day — reminder windows filter which items qualify)
  ScriptApp.newTrigger('dailyComplianceReminders')
    .timeBased().everyDays(1).atHour(7).create();
  // Weekly Monday at 7:30 AM
  ScriptApp.newTrigger('weeklyComplianceDigest')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).nearMinute(30).create();
  var cc = _complianceCcList_() || '(not set)';
  ui.alert('Compliance triggers installed.\n\n' +
           '  • Daily reminders fire at 7 AM — one email per person with items due in 30/14/7/3/1 days or overdue.\n' +
           '  • Weekly digest fires Monday 7:30 AM to the CC list.\n\n' +
           'CC on all reminders: ' + cc + '\n' +
           '(Change via Tracker → Compliance → Set CC List)\n\n' +
           'Replaced ' + removed + ' prior trigger(s).');
}


// ── Menu-callable helpers ─────────────────────────────────────────────────

function setComplianceCcListFromMenu() {
  var ui = SpreadsheetApp.getUi();
  var cur = _complianceCcList_() || '(none)';
  var r = ui.prompt('Compliance CC Emails',
    'Comma-separated emails that get CC\'d on every reminder.\nCurrent: ' + cur + '\n\n(Usually Amanda + Brandon.)',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    var result = setComplianceCcList(r.getResponseText());
    ui.alert(result.cleared ? 'CC list cleared.' : 'CC list saved: ' + result.cc);
  } catch (e) {
    ui.alert('Save failed: ' + e.message);
  }
}

function sendComplianceRemindersTest() {
  var ui = SpreadsheetApp.getUi();
  var r = dailyComplianceReminders();
  ui.alert('Daily compliance reminder dry-run sent ' + r.sent + ' email(s) across ' + r.buckets + ' recipient(s).');
}

function sendComplianceDigestTest() {
  var ui = SpreadsheetApp.getUi();
  var r = weeklyComplianceDigest();
  ui.alert('Weekly digest sent ' + r.sent + ' email(s). Overdue: ' + (r.overdue||0) + ' · Due this week: ' + (r.dueThisWeek||0));
}
var COMPLIANCE_SEED = [
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Property Taxes",
    "Location": "United States",
    "Entity": "NF Texas",
    "Obligation": "Pay property taxes by 1/31 each year\nHarris County",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "2026-01-15",
    "Next Due Date": "2027-01-31",
    "Description": "Texas taxes SOP",
    "Link to Docs": "Property taxes",
    "Cost": "21562.0",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "https://myharriscountytax.com/",
    "Account Number": "069-124-002-0008",
    "In Keeper": "N/A",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Property Taxes",
    "Location": "United States",
    "Entity": "NF Texas",
    "Obligation": "Pay property taxes by 1/31 each year\nSpring Branch ISD",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "2026-01-15",
    "Next Due Date": "2027-01-31",
    "Description": "Texas taxes SOP",
    "Link to Docs": "Property taxes",
    "Cost": "42895.0",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "https://sbisd.propertytaxpayments.net/search",
    "Account Number": "691240020008.0",
    "In Keeper": "N/A",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Property Taxes",
    "Location": "United States",
    "Entity": "NF Texas",
    "Obligation": "Pay the proerty insurance",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "2025-09-01",
    "Next Due Date": "",
    "Description": "Coordinate in the account for 709 Kuhlman chat",
    "Link to Docs": "",
    "Cost": "28386.0",
    "Payment Method": "NF Texas account",
    "Point of Contact": "It is with Liberty Mutual",
    "Website": "Kuhlman Construction loan requires FULL payment of property insurance . Please pay the property insurance in full.\n\nhttps://www.libertymutual.com/     \nLogin: mnguy003@gmail.com    \npass: Kuhlman709?",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Donation",
    "Category": "Act 22 requirement",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Make a yearly donation in PR - For 2027 - \n\nFor 2027 - There’s this charity party every year in Dorado",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "2026-01-08",
    "Next Due Date": "2027-12-31",
    "Description": "Donation SOP",
    "Link to Docs": "PR annual donation",
    "Cost": "5000.0",
    "Payment Method": "",
    "Point of Contact": "Marta Rivera Rondón\nmarta@mareducados.org\n(787) 372-8808",
    "Website": "N/A",
    "Account Number": "N/A",
    "In Keeper": "N/A",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "529 Gift",
    "Category": "529 Gift",
    "Location": "United States",
    "Entity": "MN Personal",
    "Obligation": "Make a yearly contribution to Dr. Mikes nieces and nephews",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "2025-12-29",
    "Next Due Date": "2026-12-31",
    "Description": "529 SOP",
    "Link to Docs": "Ugift - 529",
    "Cost": "114000",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "https://www.ugift529.com/gifttpl/gifter/authentication/viewCollectUsername.cs",
    "Account Number": "N/A",
    "In Keeper": "Yes",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Car",
    "Category": "Tesla",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Tesla Car Insurance",
    "Person Responsible": "Tim",
    "Frequency": "Yearly",
    "Last Done": "2025-10-06",
    "Next Due Date": "2026-10-06",
    "Description": "Tesla - Maphre Certificate of Insurance - Expired 10-06-2026.pdf",
    "Link to Docs": "Tesla model Y 2026",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Car",
    "Category": "Tesla",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Tesla Registration",
    "Person Responsible": "Tim",
    "Frequency": "Yearly",
    "Last Done": "2026-08-31",
    "Next Due Date": "2027-09-30",
    "Description": "",
    "Link to Docs": "Tesla model Y 2026",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Car",
    "Category": "Tahoe",
    "Location": "Colombia",
    "Entity": "NF MDE CO",
    "Obligation": "Tahoe insurance SURA",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2025-12-23",
    "Next Due Date": "2026-12-28",
    "Description": "",
    "Link to Docs": "",
    "Cost": "3842.0",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Car",
    "Category": "Tahoe",
    "Location": "Colombia",
    "Entity": "NF MDE CO",
    "Obligation": "Tahoe vehicle tax - Gobernación de Antioquia",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2026-01-15",
    "Next Due Date": "2027-01-28",
    "Description": "Pay through Gobernación de Ant. website - Paid yearly within the Q1",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "https://www.vehiculosantioquia.com.co/impuestosweb/#/home-public",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Car",
    "Category": "Tahoe",
    "Location": "Colombia",
    "Entity": "NF MDE CO",
    "Obligation": "Tahoe - SOAT renewal",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2025-12-25",
    "Next Due Date": "2026-12-25",
    "Description": "Buy a new SOAT through Rappi or another website that sells that insurance. SOAT is a mandatory insurance that every car in Colombia must have - It's valid for a year",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Insurance",
    "Category": "Dorado Insurance",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Renew Insurance every year and make sure the hazard policy meets Oriental's requirements and send to them",
    "Person Responsible": "Tim/Amanda",
    "Frequency": "Yearly",
    "Last Done": "2026-01-23",
    "Next Due Date": "2027-05-14",
    "Description": "2026-05-14 – Michael Nguyen – MAPFRE PRAICO Homeowners Insurance Policy – Expires May 2027.pdf",
    "Link to Docs": "Personal",
    "Cost": "2814.0",
    "Payment Method": "",
    "Point of Contact": "Frances Padilla\nfpadilla@donatoinsurance.com\n(787) 781-8000 Ext. 427",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Insurance",
    "Category": "Dorado Insurance",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Obtain the Master Policy from Dorado\nDue to Oriental 60 days before it expires",
    "Person Responsible": "Tim",
    "Frequency": "Yearly",
    "Last Done": "2025-11-14",
    "Next Due Date": "2026-11-14",
    "Description": "Master Policy - 11-14-2025 - 11-14-2026.pdf",
    "Link to Docs": "Master policy",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "Carlos - cobros@martinalgroup.com",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Insurance",
    "Category": "Dorado Insurance",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Commercial insurance",
    "Person Responsible": "Tim",
    "Frequency": "Yearly",
    "Last Done": "2026-02-17",
    "Next Due Date": "2027-02-16",
    "Description": "02-16-2026 - Commercial Certificate Liability Insurance - Dorado PH.pdf",
    "Link to Docs": "Commercial",
    "Cost": "1650.0",
    "Payment Method": "",
    "Point of Contact": "Rafael Zequeria\n787-727-5035\nrezequeira@gmail.com",
    "Website": "",
    "Account Number": "Policy number: CG000920492",
    "In Keeper": "",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Insurance",
    "Category": "Dorado Insurance",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Golf cart insurance",
    "Person Responsible": "Tim",
    "Frequency": "Yearly",
    "Last Done": "2025-10-06",
    "Next Due Date": "2026-10-06",
    "Description": "2025 - 10 - Michael D Nguyen – MAPFRE Personal Package Insurance Policy – Tesla Model Y and Golf Carts.pdf",
    "Link to Docs": "Golf Cart",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Condado Lagoon Insurance",
    "Location": "Puerto Rico",
    "Entity": "NF PR SJ",
    "Obligation": "Personal/Commercial insurance",
    "Person Responsible": "Amanda",
    "Frequency": "Yearly",
    "Last Done": "2026-06-08",
    "Next Due Date": "2027-06-08",
    "Description": "",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "CRIM Property Taxes",
    "Location": "Puerto Rico",
    "Entity": "NF PR SJ",
    "Obligation": "Pay the bi-yearly CRIM taxes for Condado Lagoon Villas",
    "Person Responsible": "Amanda",
    "Frequency": "bi-yearly",
    "Last Done": "2026-06-19",
    "Next Due Date": "2027-04-01",
    "Description": "Invoice comes around of June & then Due by October",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Deeded parking spot",
    "Location": "Puerto Rico",
    "Entity": "NF PR SJ",
    "Obligation": "Pay the fees yearly",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "2026-06-05",
    "Next Due Date": "2027-06-05",
    "Description": "",
    "Link to Docs": "",
    "Cost": "420.0",
    "Payment Method": "Oriental Debit Card",
    "Point of Contact": "yaritza@lasbrisasproperty.com\nmargret@lasbrisasproperty.com",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Leased parking spot",
    "Location": "Puerto Rico",
    "Entity": "NF PR SJ",
    "Obligation": "Pay the fees yearly",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "2026-06-05",
    "Next Due Date": "2027-06-05",
    "Description": "",
    "Link to Docs": "",
    "Cost": "1740.0",
    "Payment Method": "Oriental Debit Card",
    "Point of Contact": "yaritza@lasbrisasproperty.com\nmargret@lasbrisasproperty.com",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "CRIM Property Taxes",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Pay the bi-yearly CRIM Dorado Taxes",
    "Person Responsible": "Tim/Amanda",
    "Frequency": "bi-yearly",
    "Last Done": "2026-06-19",
    "Next Due Date": "2027-04-01",
    "Description": "Invoice comes around of June & then Due by October",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Maintenance",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "AC maintenance quarterly",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Quarterly",
    "Last Done": "2026-04-17",
    "Next Due Date": "2026-10-17",
    "Description": "AC Maintenance Schedule",
    "Link to Docs": "AC",
    "Cost": "295.0",
    "Payment Method": "",
    "Point of Contact": "Gabriel Echevarria\n787-458-3334",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Maintenance",
    "Location": "France",
    "Entity": "MN Personal",
    "Obligation": "AC mainternance bi-yearly",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "bi-yearly",
    "Last Done": "2026-05-15",
    "Next Due Date": "2026-10-15",
    "Description": "AC Maintenance Schedule",
    "Link to Docs": "AC",
    "Cost": "938.2",
    "Payment Method": "",
    "Point of Contact": "Christophe\n+33621472993",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Maintenance",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Exterminator monthly service",
    "Person Responsible": "Brandon",
    "Frequency": "Monthly",
    "Last Done": "2026-07-12",
    "Next Due Date": "2026-08-12",
    "Description": "",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Maintenance",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Golf cart maintenance at Dorado",
    "Person Responsible": "Tim",
    "Frequency": "Other",
    "Last Done": "2026-03-19",
    "Next Due Date": "2026-09-30",
    "Description": "",
    "Link to Docs": "",
    "Cost": "700.0",
    "Payment Method": "ATH Movil",
    "Point of Contact": "Jayson Diaz\n787-904-8455\njjgolfcartcarepr@gmail.com",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": "Usually comes every 3 months"
  },
  {
    "Status": "Up to Date",
    "Item Type": "MN Personal",
    "Category": "Passport",
    "Location": "United States",
    "Entity": "MN Personal",
    "Obligation": "Renew passport",
    "Person Responsible": "Mike/Amanda/Brandon",
    "Frequency": "Other",
    "Last Done": "",
    "Next Due Date": "2034-03-25",
    "Description": "",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "MN Personal",
    "Category": "Medical License",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Renew Dr. Mikes medical license",
    "Person Responsible": "Mike",
    "Frequency": "Other",
    "Last Done": "2024-08-04",
    "Next Due Date": "2027-08-04",
    "Description": "",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "MN Personal",
    "Category": "Puerto Rico Drivers license",
    "Location": "Puerto Rico",
    "Entity": "MN Personal",
    "Obligation": "Renew Dr. Mikes PR Drivers License",
    "Person Responsible": "Mike/Amanda/Brandon",
    "Frequency": "Other",
    "Last Done": "2025-06-20",
    "Next Due Date": "2033-08-04",
    "Description": "",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Business expense",
    "Category": "Keeper security",
    "Location": "United States",
    "Entity": "TLMND",
    "Obligation": "Renew Keeper security every 3 years",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Other",
    "Last Done": "2025-06-24",
    "Next Due Date": "2028-06-25",
    "Description": "",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Entity",
    "Category": "Maibox",
    "Location": "France",
    "Entity": "Paris Thacko",
    "Obligation": "Mailbox subscription - Paris\nThe address is 253 rue Saint Honoré 75001 Paris France. It's registered as the SCI address",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2025-07-16",
    "Next Due Date": "2026-07-01",
    "Description": "Pay through Banco Santander portal via transfer",
    "Link to Docs": "X",
    "Cost": "1261.0",
    "Payment Method": "Transfer",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "House Insurance",
    "Location": "France",
    "Entity": "Paris Thacko",
    "Obligation": "House Insurance of 5 Quai Montebello AXXA",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2025-12-05",
    "Next Due Date": "2026-12-01",
    "Description": "Pay through Banco Santander portal via transfer",
    "Link to Docs": "X",
    "Cost": "528.0",
    "Payment Method": "Transfer",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Property Taxes",
    "Location": "France",
    "Entity": "Paris Thacko",
    "Obligation": "Tax Fonciere - RE Tax",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2025-10-27",
    "Next Due Date": "2026-10-01",
    "Description": "Automatic debit set up",
    "Link to Docs": "X",
    "Cost": "2428.0",
    "Payment Method": "Automatic debit",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Property Taxes",
    "Location": "France",
    "Entity": "Paris Thacko",
    "Obligation": "Vacant Housing Tax - To be paid if the property wasn't rented out during 90 consecutive days during the previous year",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2026-12-30",
    "Next Due Date": "2026-12-30",
    "Description": "Automatic debit set up",
    "Link to Docs": "X",
    "Cost": "5777.0",
    "Payment Method": "Automatic debit",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Entity",
    "Category": "Entity Taxes",
    "Location": "France",
    "Entity": "Paris Thacko",
    "Obligation": "Paris Thacko corporate tax. Amount depends on the company income",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2025-10-01",
    "Next Due Date": "2026-10-01",
    "Description": "Automatic debit set up",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "HOA",
    "Location": "France",
    "Entity": "Paris Thacko",
    "Obligation": "5 Quai Montebello HOA - Paid to the Syndic quarterly",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Quarterly",
    "Last Done": "2025-12-28",
    "Next Due Date": "2026-04-01",
    "Description": "Pay through Banco Santander portal via transfer",
    "Link to Docs": "X",
    "Cost": "1296.0",
    "Payment Method": "Transfer",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Entity",
    "Category": "Accountant fees",
    "Location": "France",
    "Entity": "Paris Thacko",
    "Obligation": "Fees paid to France accountants Groupe Prieur. Twice a year (In December and March)",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "bi-yearly",
    "Last Done": "2025-12-09",
    "Next Due Date": "2026-03-10",
    "Description": "Pay through Banco Santander portal via transfer",
    "Link to Docs": "X",
    "Cost": "1860.0",
    "Payment Method": "Transfer",
    "Point of Contact": "TROUVRAIS@groupe-prieur.fr",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Entity",
    "Category": "Tax returns filing",
    "Location": "France",
    "Entity": "Paris Thacko",
    "Obligation": "Tax returns filing of the previous accounting year - W/ accountants Groupe Prieur",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2025-04-01",
    "Next Due Date": "2026-04-01",
    "Description": "",
    "Link to Docs": "X",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "TROUVRAIS@groupe-prieur.fr",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "MN Personal",
    "Category": "Subscription",
    "Location": "United States",
    "Entity": "MN Personal",
    "Obligation": "United club membership",
    "Person Responsible": "Mike/Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "2025-09-30",
    "Next Due Date": "2026-09-30",
    "Description": "",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "TLMND business card 2001",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": "From Dr. Mike: 9/11/25: Fyi this is on auto renew now\n\nPls add this to our calendar, for some reason I can’t activate the auto-renew, maybe when its expiring and we buy the next one we can activate it on the TLMND business card"
  },
  {
    "Status": "Up to Date",
    "Item Type": "MN Personal",
    "Category": "Subscription",
    "Location": "United States",
    "Entity": "MN Personal",
    "Obligation": "AARP Membership renewal (every 5 years)",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Other",
    "Last Done": "2026-02-03",
    "Next Due Date": "2031-02-03",
    "Description": "Go to website and click renew membership:\nhttps://secure.aarp.org/",
    "Link to Docs": "AARP Membership Renewal.png",
    "Cost": "79.0",
    "Payment Method": "Chase credit card (4554)",
    "Point of Contact": "",
    "Website": "https://secure.aarp.org/",
    "Account Number": "",
    "In Keeper": "Yes",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "MN Personal",
    "Category": "Health insurance",
    "Location": "United States",
    "Entity": "MN Personal",
    "Obligation": "Renew Mike's Health insurance with GeoBlue",
    "Person Responsible": "Tim/Amanda",
    "Frequency": "Yearly",
    "Last Done": "2026-09-08",
    "Next Due Date": "2027-10-14",
    "Description": "",
    "Link to Docs": "GeoBlue insurance expires 10/14/26.jpg\n2026 - 10 - Michael Nguyen – Blue Cross Blue Shield Global Solutions – PPO Insurance ID Card.pdf",
    "Cost": "$762/month",
    "Payment Method": "Oriental debit 8813",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Not Started",
    "Item Type": "Property",
    "Category": "Real Estate Tax",
    "Location": "Colombia",
    "Entity": "NF MDE CO",
    "Obligation": "Quebrada Arriba RE Tax (Predial)",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2025-02-10",
    "Next Due Date": "2027-02-28",
    "Description": "Pay RE tax through Alcaldía de Med. portal. This tax can be either paid quarterly or yearly and can be paid anytime throughout the year",
    "Link to Docs": "",
    "Cost": "Variable",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "https://www.medellin.gov.co",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Real Estate Tax",
    "Location": "Colombia",
    "Entity": "NF MDE CO",
    "Obligation": "Cinturón Verde RE Tax (Predial)",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2026-02-10",
    "Next Due Date": "2027-02-10",
    "Description": "Pay RE tax through Alcaldía de Med. portal. This tax can be either paid quarterly or yearly and can be paid anytime throughout the year",
    "Link to Docs": "",
    "Cost": "Variable",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "https://www.medellin.gov.co",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Real Estate Tax",
    "Location": "Colombia",
    "Entity": "NF MDE CO",
    "Obligation": "Provincia Lot RE Tax (Predial)",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2026-02-10",
    "Next Due Date": "",
    "Description": "Pay RE tax through Alcaldía de Med. portal. This tax can be either paid quarterly or yearly and can be paid anytime throughout the year",
    "Link to Docs": "",
    "Cost": "Variable",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "https://www.medellin.gov.co",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Insurance",
    "Category": "Life insurance",
    "Location": "United States",
    "Entity": "MN Personal",
    "Obligation": "Renew Mikes life insurance every year through lincoln financial",
    "Person Responsible": "Tim",
    "Frequency": "Yearly",
    "Last Done": "2026-01-07",
    "Next Due Date": "2027-01-07",
    "Description": "",
    "Link to Docs": "",
    "Cost": "1609.58",
    "Payment Method": "Switch to Oriental personal",
    "Point of Contact": "",
    "Website": "https://www.lincolnfinancial.com/secure/login",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": "On calendar for 1/5"
  },
  {
    "Status": "Up to Date",
    "Item Type": "Entity",
    "Category": "Legal Requirement",
    "Location": "Colombia",
    "Entity": "NF MDE CO",
    "Obligation": "Renovación Matrícula Mercantil",
    "Person Responsible": "Manuela Vallejo",
    "Frequency": "Yearly",
    "Last Done": "2026-03-31",
    "Next Due Date": "2027-03-31",
    "Description": "Renew the company's \"Matrícula Mercantil\" annually. Due date: Until March 31st. This is a legal requirement for every company dully established in Colombia",
    "Link to Docs": "",
    "Cost": "$12.500.000 Cop",
    "Payment Method": "Bank transfer",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "Flood Insurance",
    "Location": "United States",
    "Entity": "NF Texas",
    "Obligation": "Pay the flood insurance premium by May 30 every year",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "2026-04-25",
    "Next Due Date": "2027-05-30",
    "Description": "Go to https://www.libertymutual.com/ & pay the premium. Confirm in group chat no changes and good to pay",
    "Link to Docs": "https://drive.google.com/drive/folders/1mHilGdrKWVSydsZydxc8ahzamDt-JIz7",
    "Cost": "1200.0",
    "Payment Method": "TLMND Credit Card",
    "Point of Contact": "",
    "Website": "https://www.libertymutual.com/",
    "Account Number": "",
    "In Keeper": "Yes",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Trust",
    "Category": "Bryn Mawr Trust Company",
    "Location": "United States",
    "Entity": "NF6 Family 2026 US Trust",
    "Obligation": "Pay the yearly fee to Bryn Mawr Trust Company",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "2026-06-15",
    "Next Due Date": "2027-06-15",
    "Description": "Email received from them to make the yearly payment & we send a wire",
    "Link to Docs": "",
    "Cost": "5000.0",
    "Payment Method": "MN Personal 1319",
    "Point of Contact": "AColeman2@bmt.com\nAnna Coleman\nO: 302-246-1082",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "In Progress",
    "Item Type": "MN Personal",
    "Category": "ABA quarterly questions",
    "Location": "United States",
    "Entity": "MN Personal",
    "Obligation": "Complete questions every quarter",
    "Person Responsible": "Peter",
    "Frequency": "Quarterly",
    "Last Done": "2026-07-01",
    "Next Due Date": "2027-01-01",
    "Description": "Every quarter, Peter to complete questions",
    "Link to Docs": "https://docs.google.com/document/d/14UmMuIsri5TwXfAb3UFqeYqOe58i03-gVBltiPA-BFA/edit?usp=drivesdk",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "Yes",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "In Progress",
    "Item Type": "MN Personal",
    "Category": "ABVLM quarterly questions",
    "Location": "United States",
    "Entity": "MN Personal",
    "Obligation": "Complete questions every quarter",
    "Person Responsible": "Peter",
    "Frequency": "Quarterly",
    "Last Done": "2026-10-03",
    "Next Due Date": "2026-10-01",
    "Description": "Every quarter, Peter to complete questions",
    "Link to Docs": "URL: https://abvlm.starttest.com \nUsername: dr.michael@vipmedicalgroup.com",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  },
  {
    "Status": "Up to Date",
    "Item Type": "Property",
    "Category": "April 15 annual fee reminder — $150 due to PR DOS each year​​​​​​​​​​​​​​​​",
    "Location": "Puerto Rico",
    "Entity": "NF PR SJ",
    "Obligation": "April 15 annual fee reminder — $150 due to PR DOS each year​​​​​​​​​​​​​​​​",
    "Person Responsible": "Amanda/Brandon",
    "Frequency": "Yearly",
    "Last Done": "",
    "Next Due Date": "2027-04-01",
    "Description": "April 15 annual fee reminder — $150 due to PR DOS each year​​​​​​​​​​​​​​​​",
    "Link to Docs": "",
    "Cost": "",
    "Payment Method": "",
    "Point of Contact": "",
    "Website": "",
    "Account Number": "",
    "In Keeper": "",
    "On FO Calendar": "No",
    "Calendar Event Link": "",
    "Notes": ""
  }
];

// Menu-callable: one-shot seed from Amanda's existing Compliance Calendar
// spreadsheet. Idempotent — skips if any items already exist. Also seeds
// COMPLIANCE_PEOPLE with the unique person names found (email left blank
// for Amanda to fill in).
function seedComplianceCalendar() {
  _requireEditor_();
  ensureComplianceSheets_();
  var ui = SpreadsheetApp.getUi();
  var existing = _getComplianceRows_('COMPLIANCE_ITEMS', COMPLIANCE_ITEMS_HEADERS);
  if (existing.length) {
    ui.alert('COMPLIANCE_ITEMS already has ' + existing.length + ' rows — seed skipped. Delete rows manually first if you want to re-seed.');
    return;
  }
  var now = new Date();
  var count = 0;
  COMPLIANCE_SEED.forEach(function(it, i) {
    var num = String(i + 1); while (num.length < 5) num = '0' + num;
    var id = 'ci-' + num;
    // Normalize date fields to Date objects for sheet storage
    ['Last Done', 'Next Due Date'].forEach(function(k) {
      if (typeof it[k] === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(it[k])) {
        it[k] = new Date(it[k] + 'T12:00:00Z');
      }
    });
    it['ID'] = id;
    it['Created At'] = now;
    it['Last Updated'] = now;
    _writeComplianceRow_('COMPLIANCE_ITEMS', COMPLIANCE_ITEMS_HEADERS, it);
    count++;
  });
  // Seed COMPLIANCE_PEOPLE from unique parsed names
  var names = {};
  COMPLIANCE_SEED.forEach(function(it) {
    String(it['Person Responsible'] || '').split(/[\/,&]+/).forEach(function(n) {
      var k = n.trim();
      if (k) names[k] = true;
    });
  });
  var seededPeople = 0;
  Object.keys(names).sort().forEach(function(name) {
    try { upsertCompliancePerson(name, '', true); seededPeople++; } catch(e) {}
  });
  ui.alert('Compliance Calendar seeded.\n\n' +
           '  • ' + count + ' items added to COMPLIANCE_ITEMS\n' +
           '  • ' + seededPeople + ' unique names added to COMPLIANCE_PEOPLE (fill in emails from the web UI)\n\n' +
           'Next: open the web app → Compliance tab → fill in email addresses for each person.');
}
