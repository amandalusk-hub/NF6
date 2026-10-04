/**
 * MoneyMovement.gs — Multi-hop wire tracker.
 *
 * Amanda's family runs complex flow-of-funds structures — money that needs
 * to arrive at Blue Panda FLP, NF PR SJ, or Tiger Capital rarely comes from
 * ONE account. It comes from multiple people (Mike, Nancy, siblings) and
 * flows through several intermediate entities (personal accounts → trusts →
 * FLPs / LLCs → destination). Ownership percentages at each parent entity
 * determine how a total distribution splits across contributors, and each
 * contributor's dollars still have to hop through their chain to land.
 *
 * Today she does this in Excel: rows are movement events, columns are the
 * individual wires, checkboxes track completion. Manual math for the amounts.
 *
 * This module models that as data:
 *
 *   MOVEMENT_TEMPLATES  — one row per flow definition (e.g. "Blue Panda FLP
 *                         Contribution", "NF6 Tiger Capital Distribution",
 *                         "NF PR SJ Wire"). Fixed structure that can be reused
 *                         every time money flows through this path.
 *
 *   MOVEMENT_HOPS       — the individual wires that make up a template. Each
 *                         hop has: an order, a from/to account, a responsible
 *                         person, and a "% of total" — when Amanda enters a
 *                         movement total of $X, each hop's dollar amount is
 *                         computed as X × its stored percentage.
 *
 *   MOVEMENTS           — one row per instance the template is executed:
 *                         template ID + total amount + date + status.
 *
 *   MOVEMENT_WIRES      — one row per wire in a specific movement. Links back
 *                         to MOVEMENTS + MOVEMENT_HOPS, holds the actual
 *                         amount, status (Pending / Sent / Confirmed), sent
 *                         date, and (later) the auto-matched Plaid txn ID.
 *
 * This file (commit 1): schemas + seed all 4 known templates + basic web-
 * callable CRUD. UI, Plaid auto-match, and email alerts are follow-on commits.
 */

// ── Schemas ────────────────────────────────────────────────────────────────

var MOVEMENT_TEMPLATES_HEADERS = [
  'ID',
  'Name',                     // "Blue Panda FLP Contribution", "NF6 Tiger Capital Distribution (Down)"
  'Destination',              // Where the money lands (or comes from, for Down direction).
  'Direction',                // 'up' = contributions INTO destination.
                              // 'down' = distributions OUT of destination.
  'Description',
  'Active',                   // Yes / No
  'Date Added',
  'Last Updated'
];

var MOVEMENT_HOPS_HEADERS = [
  'ID',
  'Template ID',
  'Order',                    // Display + execution order within the template.
                              // Multiple hops can share an order if they're
                              // independent parallel wires (different chains).
  'Chain',                    // Logical grouping — hops in the same chain
                              // belong to one contributor's path. Used to
                              // render the checklist grouped by contributor.
                              // e.g. "Mike (via MN Trust)", "Mike (via BPMGMT)",
                              // "Nancy (via BPMGMT)".
  'From Account',             // The account/entity sending money in this hop.
  'To Account',               // The account/entity receiving.
  'Responsible',              // Who kicks off this wire — Amanda, Ben, Gagan,
                              // Tim, etc. Powers the "your wires" view.
  'Amount % of Total',        // e.g. 99.00 means this hop moves 99% of the
                              // movement's total amount. Sum of the LAST hop
                              // in each chain should = 100% (money in) or
                              // = 100% (money out), matching physical reality.
  'Notes'
];

var MOVEMENTS_HEADERS = [
  'ID',
  'Template ID',
  'Total Amount',             // The USD total the movement moves.
  'Date Needed',              // When it needs to be done by.
  'Status',                   // 'Planning' | 'In Progress' | 'Complete' | 'Cancelled'
  'Notes',
  'Created By',
  'Created At',
  'Last Updated'
];

var MOVEMENT_WIRES_HEADERS = [
  'ID',
  'Movement ID',
  'Hop ID',                   // FK back to MOVEMENT_HOPS for from/to/chain/responsible
  'Amount',                   // Snapshot at movement creation time — hop's
                              // "Amount % of Total" × movement total. Frozen
                              // even if the template's percentage is later
                              // edited, so history is preserved.
  'Status',                   // 'Pending' | 'Sent' | 'Confirmed' | 'Skipped'
  'Sent Date',
  'Plaid Txn ID',             // If auto-matched to a real transaction (later
                              // commit — Plaid matching).
  'Notes',
  'Marked Done By',
  'Marked Done At'
];


// ── Sheet lifecycle ────────────────────────────────────────────────────────

function ensureMoneyMovementSheets_() {
  _ensureMMSheet_('MOVEMENT_TEMPLATES', MOVEMENT_TEMPLATES_HEADERS);
  _ensureMMSheet_('MOVEMENT_HOPS',      MOVEMENT_HOPS_HEADERS);
  _ensureMMSheet_('MOVEMENTS',          MOVEMENTS_HEADERS);
  _ensureMMSheet_('MOVEMENT_WIRES',     MOVEMENT_WIRES_HEADERS);
}

function _ensureMMSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
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
    var startCol = existing.length + 1;
    sheet.getRange(1, startCol, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold').setBackground('#14263d').setFontColor('#ffffff');
  }
  return sheet;
}


// ── Header-name read/write helpers ─────────────────────────────────────────

function _getMMRows_(sheetName, headers) {
  ensureMoneyMovementSheets_();
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
  }).filter(function(o) { return o.ID; });
}

function _writeMMRow_(sheetName, headers, obj) {
  ensureMoneyMovementSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var lastCol = Math.max(sheet.getLastColumn(), headers.length);
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row = hdr.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.appendRow(row);
}

function _updateMMRow_(sheetName, headers, id, patch) {
  ensureMoneyMovementSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet.getLastRow() < 2) return false;
  var lastCol = Math.max(sheet.getLastColumn(), headers.length);
  var data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  var hdr = data[0];
  var iId = hdr.indexOf('ID');
  if (iId < 0) return false;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iId]) !== String(id)) continue;
    Object.keys(patch).forEach(function(k) {
      var idx = hdr.indexOf(k);
      if (idx >= 0) sheet.getRange(r + 1, idx + 1).setValue(patch[k]);
    });
    return true;
  }
  return false;
}


// ── Web-callable CRUD ──────────────────────────────────────────────────────

// Returns every template (for the "New Movement" picker). Only rows whose
// Active column is EXPLICITLY 'No' (or false/0) are excluded — an empty or
// missing Active value counts as active, so a hand-added template row that
// forgot to fill it still shows up.
function getMovementTemplates() {
  var rows = _getMMRows_('MOVEMENT_TEMPLATES', MOVEMENT_TEMPLATES_HEADERS)
    .filter(function(t) {
      var raw = t['Active'];
      if (raw === false || raw === 0) return false;
      var a = String(raw == null ? '' : raw).toLowerCase().trim();
      return a !== 'no' && a !== 'false' && a !== '0';
    });
  // JSON round-trip so Date columns + any nested Google Sheets values
  // serialize cleanly across google.script.run (same trick we needed on
  // getLoansStatus before).
  return JSON.parse(JSON.stringify(rows));
}

// Returns the hops for a given template, sorted by (Order, Chain, then insertion).
function getMovementHops(templateId) {
  return _getMMRows_('MOVEMENT_HOPS', MOVEMENT_HOPS_HEADERS)
    .filter(function(h) { return String(h['Template ID']) === String(templateId); })
    .sort(function(a, b) {
      var oa = Number(a['Order']) || 0, ob = Number(b['Order']) || 0;
      if (oa !== ob) return oa - ob;
      return String(a['Chain']).localeCompare(String(b['Chain']));
    });
}

// Full template detail — template + hops together (for the UI).
function getMovementTemplateDetail(templateId) {
  var templates = _getMMRows_('MOVEMENT_TEMPLATES', MOVEMENT_TEMPLATES_HEADERS);
  var template = templates.find(function(t) { return String(t['ID']) === String(templateId); });
  if (!template) return { error: 'Template not found: ' + templateId };
  return { template: template, hops: getMovementHops(templateId) };
}

// Create a new movement from a template. Snapshots each hop's amount at
// creation time (hopAmount = totalAmount × hop's % of total) so that
// subsequent edits to the template's percentages never rewrite historical
// wire amounts.
function createMovement(templateId, totalAmount, dateNeeded, notes) {
  _requireEditor_();
  var detail = getMovementTemplateDetail(templateId);
  if (detail.error) throw new Error(detail.error);
  var total = Number(totalAmount) || 0;
  if (total <= 0) throw new Error('Total amount must be > 0.');

  var moveId = 'move_' + Utilities.getUuid().substring(0, 8);
  var now = new Date();
  var user = _currentUserEmail_();

  _writeMMRow_('MOVEMENTS', MOVEMENTS_HEADERS, {
    'ID':           moveId,
    'Template ID':  templateId,
    'Total Amount': total,
    'Date Needed':  dateNeeded ? new Date(dateNeeded) : '',
    'Status':       'Planning',
    'Notes':        String(notes || ''),
    'Created By':   user,
    'Created At':   now,
    'Last Updated': now
  });

  var wiresCreated = 0;
  detail.hops.forEach(function(hop) {
    var pct = Number(hop['Amount % of Total']) || 0;
    var amt = Math.round((total * pct / 100) * 100) / 100;   // 2dp
    _writeMMRow_('MOVEMENT_WIRES', MOVEMENT_WIRES_HEADERS, {
      'ID':          'wire_' + Utilities.getUuid().substring(0, 8),
      'Movement ID': moveId,
      'Hop ID':      hop['ID'],
      'Amount':      amt,
      'Status':      'Pending',
      'Sent Date':   '',
      'Plaid Txn ID':'',
      'Notes':       '',
      'Marked Done By': '',
      'Marked Done At': ''
    });
    wiresCreated++;
  });

  _logAudit_('createMovement', 'movement', moveId, detail.template['Name'],
             'Created ' + detail.template['Name'] + ' movement: $' + total.toFixed(2) + ' (' + wiresCreated + ' wires)');
  return { success: true, movementId: moveId, wiresCreated: wiresCreated };
}

// Delete a movement + all its wires. Physical delete — the row goes away.
// The audit log preserves the fact the delete happened. Used by the trash
// button on the movement card / detail modal.
function deleteMovement(movementId) {
  _requireEditor_();
  ensureMoneyMovementSheets_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Look up the movement name for the audit log before we delete the row.
  var move = _getMMRows_('MOVEMENTS', MOVEMENTS_HEADERS)
    .find(function(m) { return String(m['ID']) === String(movementId); });
  var label = move ? (move['Notes'] || move['ID']) : movementId;

  // Delete every wire row belonging to this movement.
  var wSheet = ss.getSheetByName('MOVEMENT_WIRES');
  if (wSheet && wSheet.getLastRow() >= 2) {
    var lastCol = Math.max(wSheet.getLastColumn(), MOVEMENT_WIRES_HEADERS.length);
    var data = wSheet.getRange(1, 1, wSheet.getLastRow(), lastCol).getValues();
    var hdr = data[0];
    var iMove = hdr.indexOf('Movement ID');
    // Iterate bottom-up so row numbers stay valid after deletes.
    for (var r = data.length - 1; r >= 1; r--) {
      if (String(data[r][iMove]) === String(movementId)) {
        wSheet.deleteRow(r + 1);
      }
    }
  }

  // Delete the movement row itself.
  var mSheet = ss.getSheetByName('MOVEMENTS');
  if (mSheet && mSheet.getLastRow() >= 2) {
    var mLastCol = Math.max(mSheet.getLastColumn(), MOVEMENTS_HEADERS.length);
    var mData = mSheet.getRange(1, 1, mSheet.getLastRow(), mLastCol).getValues();
    var mHdr = mData[0];
    var mIid = mHdr.indexOf('ID');
    for (var mr = mData.length - 1; mr >= 1; mr--) {
      if (String(mData[mr][mIid]) === String(movementId)) {
        mSheet.deleteRow(mr + 1);
      }
    }
  }

  _logAudit_('deleteMovement', 'movement', movementId, label, 'Deleted movement');
  return { success: true };
}

// Update a movement's status directly ('Planning' | 'In Progress' |
// 'Complete' | 'Cancelled'). Used by the checklist page's "Mark Complete"
// and "Reopen" buttons.
function setMovementStatus(movementId, newStatus) {
  _requireEditor_();
  var ok = _updateMMRow_('MOVEMENTS', MOVEMENTS_HEADERS, movementId, {
    'Status': String(newStatus || 'Planning'),
    'Last Updated': new Date()
  });
  _logAudit_('setMovementStatus', 'movement', movementId, newStatus,
             'Status → ' + newStatus);
  return { success: ok };
}

// Update a single wire's status. Used by the checklist checkbox: Pending →
// Sent → Confirmed (or Skipped if a wire doesn't apply for a particular
// movement).
function updateMovementWire(wireId, patch) {
  _requireEditor_();
  var writePatch = {};
  if (patch.status !== undefined)   writePatch['Status']      = String(patch.status);
  if (patch.sentDate !== undefined) writePatch['Sent Date']   = patch.sentDate ? new Date(patch.sentDate) : '';
  if (patch.plaidId !== undefined)  writePatch['Plaid Txn ID']= String(patch.plaidId);
  if (patch.notes !== undefined)    writePatch['Notes']       = String(patch.notes);
  if (patch.amount !== undefined)   writePatch['Amount']      = Number(patch.amount);
  if (patch.status === 'Confirmed' || patch.status === 'Sent') {
    writePatch['Marked Done By'] = _currentUserEmail_();
    writePatch['Marked Done At'] = new Date();
  }
  var ok = _updateMMRow_('MOVEMENT_WIRES', MOVEMENT_WIRES_HEADERS, wireId, writePatch);
  // After any wire-status change, re-evaluate the parent movement's status
  // so Planning → In Progress → Complete transitions happen automatically.
  var statusChange = null;
  if (ok && patch.status !== undefined) {
    try {
      var wireRow = _getMMRows_('MOVEMENT_WIRES', MOVEMENT_WIRES_HEADERS)
        .find(function(w) { return String(w['ID']) === String(wireId); });
      if (wireRow) statusChange = _recomputeMovementStatus_(wireRow['Movement ID']);
    } catch (e) { Logger.log('updateMovementWire: status recompute failed: ' + e.message); }
  }
  return { success: ok, statusChange: statusChange };
}

// Load one movement + all its wires (joined with hop details) — the shape the
// checklist UI renders from.
function getMovementDetail(movementId) {
  var moves = _getMMRows_('MOVEMENTS', MOVEMENTS_HEADERS);
  var move = moves.find(function(m) { return String(m['ID']) === String(movementId); });
  if (!move) return { error: 'Movement not found: ' + movementId };
  var template = _getMMRows_('MOVEMENT_TEMPLATES', MOVEMENT_TEMPLATES_HEADERS)
    .find(function(t) { return String(t['ID']) === String(move['Template ID']); });
  var hops = getMovementHops(move['Template ID']);
  var hopById = {};
  hops.forEach(function(h) { hopById[h['ID']] = h; });
  var wires = _getMMRows_('MOVEMENT_WIRES', MOVEMENT_WIRES_HEADERS)
    .filter(function(w) { return String(w['Movement ID']) === String(movementId); })
    .map(function(w) {
      var hop = hopById[w['Hop ID']] || {};
      return {
        wireId:      w['ID'],
        status:      w['Status'],
        amount:      Number(w['Amount']) || 0,
        sentDate:    w['Sent Date'],
        plaidId:     w['Plaid Txn ID'],
        notes:       w['Notes'],
        markedBy:    w['Marked Done By'],
        markedAt:    w['Marked Done At'],
        hopId:       w['Hop ID'],
        order:       Number(hop['Order']) || 0,
        chain:       hop['Chain'] || '',
        fromAccount: hop['From Account'] || '',
        toAccount:   hop['To Account'] || '',
        responsible: hop['Responsible'] || '',
        pctOfTotal:  Number(hop['Amount % of Total']) || 0
      };
    })
    .sort(function(a, b) {
      if (a.order !== b.order) return a.order - b.order;
      return a.chain.localeCompare(b.chain);
    });
  var done = wires.filter(function(w) { return w.status === 'Confirmed' || w.status === 'Sent'; }).length;
  return JSON.parse(JSON.stringify({
    movement: move,
    template: template,
    wires:    wires,
    progress: { done: done, total: wires.length, pct: wires.length ? Math.round(done / wires.length * 100) : 0 }
  }));
}

// List all movements (for the tab's landing view). Newest first.
function getMovements() {
  var rows = _getMMRows_('MOVEMENTS', MOVEMENTS_HEADERS)
    .sort(function(a, b) {
      var da = a['Created At'] instanceof Date ? a['Created At'].getTime() : 0;
      var db = b['Created At'] instanceof Date ? b['Created At'].getTime() : 0;
      return db - da;
    });
  return JSON.parse(JSON.stringify(rows));
}


// ── Seed: known templates ──────────────────────────────────────────────────
// Menu-callable one-shot: creates the 4 templates from Amanda's Excel
// screenshots if they aren't already there. Idempotent — matches by name.
// Percentages are the ones I extracted from her screenshots and confirmed:
//
//   NF6 Tiger Capital LLC (5319) is 99% owned by Mike (via 2019 MN Family
//   Rev Trust → NF6 Family Holdings) and 1% by Michelle/Nancy/David combined
//   (via NF6 Joint Mgmt LLC → NF6 Family Holdings, split 1/3 each = 0.333%
//   per sibling).
//
//   NF PR SJ LLC (5297) has the same ultimate ownership (99% Mike, 0.333% per
//   sibling) through the same NF6 Family Holdings (7932) parent.
//
//   Blue Panda FLP (8686) is 99% owned by MN Trust Irrevocable (100% Mike)
//   and 1% by BPMGMT (5150). BPMGMT is 51% Mike (via 2019 MN Family Rev
//   Trust) and 49% Nancy.
//
// The templates are:
//   1. Blue Panda FLP Contribution (up)
//   2. NF PR SJ Wire (up)
//   3. NF6 Tiger Capital Distribution — DOWN direction
//   4. NF6 Tiger Capital Contribution — UP direction

function seedMovementTemplates() {
  _requireEditor_();
  ensureMoneyMovementSheets_();
  var existing = _getMMRows_('MOVEMENT_TEMPLATES', MOVEMENT_TEMPLATES_HEADERS);
  var haveByName = {};
  existing.forEach(function(t) { haveByName[String(t['Name']).toLowerCase()] = t['ID']; });

  var results = [];

  // 1. Blue Panda FLP - Down (money coming DOWN into Blue Panda from owners)
  results.push(_seedTemplate_({
    name: 'Blue Panda FLP - Down',
    destination: 'Blue Panda FLP (8686)',
    direction: 'down',
    description: 'Money flowing IN to Blue Panda FLP. 99% via MN Trust Irrevocable (100% Mike). 1% via BPMGMT (5150), which is 51% Mike / 49% Nancy.',
    hops: [
      // Blue Panda scaling model: Amanda's input = MN Trust wire amount
      // (= Mike's main contribution). Her Excel enters this number in
      // column E and derives everything else. So the MN Trust chain = 100%
      // of input, BPMGMT consolidated = 1.0101% (= 1/99 of input), with
      // sub-splits 51% Mike / 49% Nancy inside BPMGMT.
      //
      // Example at user input $1,000,000:
      //   Mike → MN Trust = $1,000,000 (100%)
      //   MN Trust → Blue Panda = $1,000,000 (100%)
      //   Mike → Rev Trust = $5,151.52 (0.515152%)
      //   Rev Trust → BPMGMT = $5,151.52 (0.515152%)
      //   Nancy → BPMGMT = $4,949.49 (0.494949%)
      //   BPMGMT → Blue Panda = $10,101.01 (1.010101%)
      //   → Destination amount at Blue Panda = $1,010,101.01

      // Chain A: Mike via MN Trust (100% of input)
      { order: 10, chain:'Mike (via MN Trust)',  from:'Michael Nguyen Personal (1319)',   to:'MN Trust Irrevocable',                  responsible:'Amanda', pct: 100.00, notes:'' },
      { order: 20, chain:'Mike (via MN Trust)',  from:'MN Trust Irrevocable',             to:'Blue Panda FLP (8686)',                 responsible:'Amanda', pct: 100.00, notes:'' },
      // Chain B: Mike via BPMGMT (= Excel H = I * 0.51 = E * 1.0101/100 * 0.51 = E * 0.515151%)
      { order: 30, chain:'Mike (via BPMGMT)',    from:'Michael Nguyen Personal (1319)',   to:'2019 MN Family Revocable Trust (3333)', responsible:'Amanda', pct:  0.515151, notes:'51% of the 1.0101% BPMGMT slice.' },
      { order: 40, chain:'Mike (via BPMGMT)',    from:'2019 MN Family Revocable Trust (3333)', to:'BPMGMT (5150)',                    responsible:'Amanda', pct:  0.515151, notes:'' },
      // Chain C: Nancy via BPMGMT (= Excel G = I * 0.49 = E * 1.0101/100 * 0.49 = E * 0.494949%)
      { order: 30, chain:'Nancy (via BPMGMT)',   from:'Nancy Nguyen Personal',            to:'BPMGMT (5150)',                         responsible:'Ben',    pct:  0.494949, notes:'49% of the 1.0101% BPMGMT slice.' },
      // BPMGMT consolidation (= Excel I = E * 1.0101/100 = E * 1.0101%)
      { order: 50, chain:'BPMGMT consolidation', from:'BPMGMT (5150)',                    to:'Blue Panda FLP (8686)',                 responsible:'Amanda', pct:  1.0101,   notes:'After Mike + Nancy contributions merge at BPMGMT. Formula: E * 1.0101%.' }
    ]
  }, haveByName));

  // 2. NF PR SJ - Down (money coming DOWN into NF PR SJ from owners)
  results.push(_seedTemplate_({
    name: 'NF PR SJ - Down',
    destination: 'NF PR SJ LLC (5297)',
    direction: 'down',
    description: 'Money flowing IN to NF PR SJ LLC. 99% via Mike (through 2019 MN Family Rev Trust → NF6 Family Holdings). 1% via siblings (each 1/3 through Personal → NF6 Joint Mgmt → NF6 Family Holdings).',
    hops: [
      // Mike chain (99%)
      { order: 10, chain:'Mike (via Rev Trust)', from:'Michael Nguyen Personal (1319)',     to:'2019 Family Revocable Trust (3333)', responsible:'Amanda', pct: 99.00, notes:'' },
      { order: 20, chain:'Mike (via Rev Trust)', from:'2019 MN Family Revocable Trust',     to:'NF6 Family Holdings (7932)',           responsible:'Amanda', pct: 99.00, notes:'' },
      // Siblings (1/3 of 1% each = 0.333333% each). Use more precision so
      // $100,000 × 0.333333% = $333.33 rather than $333.00 with a truncated
      // 0.333%. Third sibling absorbs the rounding to preserve the sum.
      // Sibling personal-account wires are Ben's responsibility.
      { order: 30, chain:'David (via Joint Mgmt)',    from:'David Personal',    to:'NF6 Joint Mgmt LLC (8972)', responsible:'Ben', pct: 0.333333, notes:'' },
      { order: 30, chain:'Nancy (via Joint Mgmt)',    from:'Nancy Personal',    to:'NF6 Joint Mgmt LLC (8972)', responsible:'Ben', pct: 0.333333, notes:'' },
      { order: 30, chain:'Michelle (via Joint Mgmt)', from:'Michelle Personal', to:'NF6 Joint Mgmt LLC (8972)', responsible:'Ben', pct: 0.333334, notes:'' },
      // Merger + final leg
      { order: 40, chain:'Joint Mgmt consolidation', from:'NF6 Joint Mgmt LLC (8972)', to:'NF6 Family Holdings (7932)', responsible:'Amanda', pct: 1.00, notes:'After 3 sibling contributions merge.' },
      { order: 50, chain:'Final leg',                from:'NF6 Family Holdings (7932)', to:'NF PR SJ LLC (5297)',        responsible:'Amanda', pct:100.00, notes:'Full amount to destination.' }
    ]
  }, haveByName));

  // 3. NF6 Tiger Capital — DOWN direction (distribute FROM Tiger Capital TO
  //    individuals). Same ownership as NF PR SJ but reversed hops.
  results.push(_seedTemplate_({
    name: 'NF6 Tiger Capital - Up',
    destination: 'NF6 Tiger Capital LLC (5319)',
    direction: 'up',
    description: 'Money flowing UP out of Tiger Capital to owners: 99% to Mike + 0.333% each to Michelle/Nancy/David. Passes through Family Holdings → (Rev Trust for Mike | Joint Mgmt for siblings) → personal accounts.',
    hops: [
      // Full amount out of Tiger to Family Holdings
      { order: 10, chain:'Tiger → Family Holdings', from:'NF6 Tiger Capital LLC (5319)', to:'NF6 Family Holdings (7932)',            responsible:'Amanda', pct:100.00, notes:'' },
      // Family Holdings splits: 99% to Rev Trust, 1% to Joint Mgmt
      { order: 20, chain:'Mike branch (99%)',       from:'NF6 Family Holdings (7932)',   to:'2019 MN Family Revocable Trust (3333)', responsible:'Amanda', pct: 99.00, notes:'' },
      { order: 20, chain:'Siblings branch (1%)',    from:'NF6 Family Holdings (7932)',   to:'NF6 Joint Mgmt LLC (8972)',             responsible:'Amanda', pct:  1.00, notes:'' },
      // Rev Trust → Mike
      { order: 30, chain:'Mike final leg',          from:'2019 MN Family Revocable Trust (3333)', to:'Michael Nguyen Personal (1319)', responsible:'Amanda', pct: 99.00, notes:'' },
      // Joint Mgmt → each sibling (1/3 of 1%, using 0.333333% precision so
      // $100k × 0.333333% = $333.33 rather than $333.00). Third sibling
      // absorbs the rounding to preserve the sum. These are money-OUT from
      // the family LLC to sibling personal accounts — Amanda dispatches from
      // Joint Mgmt, so Amanda (not Ben) is responsible on the down direction.
      { order: 30, chain:'Michelle final leg',      from:'NF6 Joint Mgmt LLC (8972)',    to:'Michelle Personal',               responsible:'Amanda',  pct: 0.333333, notes:'' },
      { order: 30, chain:'Nancy final leg',         from:'NF6 Joint Mgmt LLC (8972)',    to:'Nancy Personal',                  responsible:'Amanda',  pct: 0.333333, notes:'' },
      { order: 30, chain:'David final leg',         from:'NF6 Joint Mgmt LLC (8972)',    to:'David Personal',                  responsible:'Amanda',  pct: 0.333334, notes:'' }
    ]
  }, haveByName));

  // 5. Blue Panda FLP — DOWN direction (distributions OUT of Blue Panda).
  //    Inverse of #1. Amanda uses this more often than the UP version.
  results.push(_seedTemplate_({
    name: 'Blue Panda FLP - Up',
    destination: 'Blue Panda FLP (8686)',
    direction: 'up',
    description: 'Money flowing UP out of Blue Panda FLP to owners. 99% flows back through MN Trust Irrevocable to Mike. 1% flows through BPMGMT (51% Mike via 2019 Rev Trust, 49% Nancy).',
    hops: [
      // Same scaling as Blue Panda - Down (input = MN Trust wire). Produces
      // the same per-row amounts as Amanda's Excel, just with outbound hops.
      { order: 10, chain:'MN Trust branch',        from:'Blue Panda FLP (8686)', to:'MN Trust Irrevocable',                     responsible:'Amanda', pct: 100.00,   notes:'' },
      { order: 10, chain:'BPMGMT branch',          from:'Blue Panda FLP (8686)', to:'BPMGMT (5150)',                            responsible:'Amanda', pct:  1.0101,  notes:'E * 1.0101%' },
      { order: 20, chain:'MN Trust → Mike',        from:'MN Trust Irrevocable',   to:'Michael Nguyen Personal (1319)',           responsible:'Amanda', pct: 100.00,   notes:'' },
      { order: 20, chain:'BPMGMT → Mike branch',   from:'BPMGMT (5150)',          to:'2019 MN Family Revocable Trust (3333)',    responsible:'Amanda', pct:  0.515151, notes:'51% of the 1.0101% BPMGMT slice.' },
      { order: 20, chain:'BPMGMT → Nancy',         from:'BPMGMT (5150)',          to:'Nancy Nguyen Personal',                    responsible:'Amanda', pct:  0.494949, notes:'49% of the 1.0101% BPMGMT slice.' },
      { order: 30, chain:'Rev Trust → Mike',       from:'2019 MN Family Revocable Trust (3333)', to:'Michael Nguyen Personal (1319)', responsible:'Amanda', pct:  0.515151, notes:'' }
    ]
  }, haveByName));

  // 6. NF PR SJ — DOWN direction (distributions OUT of NF PR SJ).
  //    Inverse of #2. Amanda uses this more often than the UP version.
  results.push(_seedTemplate_({
    name: 'NF PR SJ - Up',
    destination: 'NF PR SJ LLC (5297)',
    direction: 'up',
    description: 'Money flowing UP out of NF PR SJ LLC to owners. 99% goes back to Mike (via NF6 Family Holdings → 2019 MN Family Rev Trust → Michael Personal). 1% split 1/3 each to Michelle/Nancy/David via NF6 Joint Mgmt.',
    hops: [
      // Full amount out of NF PR SJ to Family Holdings
      { order: 10, chain:'PR SJ → Family Holdings', from:'NF PR SJ LLC (5297)',     to:'NF6 Family Holdings (7932)',              responsible:'Amanda', pct:100.00, notes:'' },
      // Family Holdings splits: 99% to Rev Trust (Mike), 1% to Joint Mgmt (siblings)
      { order: 20, chain:'Mike branch (99%)',       from:'NF6 Family Holdings (7932)', to:'2019 MN Family Revocable Trust (3333)', responsible:'Amanda', pct: 99.00, notes:'' },
      { order: 20, chain:'Siblings branch (1%)',    from:'NF6 Family Holdings (7932)', to:'NF6 Joint Mgmt LLC (8972)',              responsible:'Amanda', pct:  1.00, notes:'' },
      // Rev Trust → Mike
      { order: 30, chain:'Mike final leg',          from:'2019 MN Family Revocable Trust (3333)', to:'Michael Nguyen Personal (1319)', responsible:'Amanda', pct: 99.00, notes:'' },
      // Joint Mgmt → each sibling (0.333333% each, precise to hit $333.33 on $100k)
      { order: 30, chain:'Michelle final leg',      from:'NF6 Joint Mgmt LLC (8972)', to:'Michelle Personal',                     responsible:'Amanda', pct:  0.333333, notes:'' },
      { order: 30, chain:'Nancy final leg',         from:'NF6 Joint Mgmt LLC (8972)', to:'Nancy Personal',                        responsible:'Amanda', pct:  0.333333, notes:'' },
      { order: 30, chain:'David final leg',         from:'NF6 Joint Mgmt LLC (8972)', to:'David Personal',                        responsible:'Amanda', pct:  0.333334, notes:'' }
    ]
  }, haveByName));

  // 4. NF6 Tiger Capital — UP direction (contributions INTO Tiger Capital).
  //    Same ownership, hops reversed from #3.
  results.push(_seedTemplate_({
    name: 'NF6 Tiger Capital - Down',
    destination: 'NF6 Tiger Capital LLC (5319)',
    direction: 'down',
    description: 'Money coming DOWN into Tiger Capital from owners. Mike wires 99%, each sibling wires 0.333%. Each contribution flows through Family Holdings into Tiger.',
    hops: [
      // Mike chain (99%)
      { order: 10, chain:'Mike (via Rev Trust)',    from:'Michael Nguyen Personal (1319)',       to:'2019 MN Family Revocable Trust (3333)', responsible:'Amanda', pct: 99.00, notes:'' },
      { order: 20, chain:'Mike (via Rev Trust)',    from:'2019 MN Family Revocable Trust (3333)', to:'NF6 Family Holdings (7932)',            responsible:'Amanda', pct: 99.00, notes:'' },
      // Sibling chains (0.333333% each — precise 1/3 of 1% so amounts round
      // to $333.33 rather than $333.00 on a $100k movement). Third sibling
      // absorbs the rounding. Sibling personal-account wires are Ben's.
      { order: 10, chain:'Michelle (via Joint Mgmt)', from:'Michelle Personal', to:'NF6 Joint Mgmt LLC (8972)', responsible:'Ben', pct: 0.333333, notes:'' },
      { order: 10, chain:'Nancy (via Joint Mgmt)',    from:'Nancy Personal',    to:'NF6 Joint Mgmt LLC (8972)', responsible:'Ben', pct: 0.333333, notes:'' },
      { order: 10, chain:'David (via Joint Mgmt)',    from:'David Personal',    to:'NF6 Joint Mgmt LLC (8972)', responsible:'Ben', pct: 0.333334, notes:'' },
      // Joint Mgmt consolidation → Family Holdings (1% total)
      { order: 20, chain:'Joint Mgmt consolidation', from:'NF6 Joint Mgmt LLC (8972)', to:'NF6 Family Holdings (7932)', responsible:'Amanda', pct: 1.00, notes:'After 3 sibling contributions merge.' },
      // Family Holdings → Tiger Capital (100% total)
      { order: 30, chain:'Final leg',                from:'NF6 Family Holdings (7932)', to:'NF6 Tiger Capital LLC (5319)', responsible:'Amanda', pct:100.00, notes:'Full amount to Tiger.' }
    ]
  }, haveByName));

  return { success: true, templates: results };
}

function _seedTemplate_(spec, haveByName) {
  var key = spec.name.toLowerCase();
  var templateId;
  if (haveByName[key]) {
    templateId = haveByName[key];
    // Refresh description/direction on existing template but leave hops alone
    // (Amanda may have hand-edited amounts / added a hop we don't know about).
    _updateMMRow_('MOVEMENT_TEMPLATES', MOVEMENT_TEMPLATES_HEADERS, templateId, {
      'Destination': spec.destination,
      'Direction':   spec.direction,
      'Description': spec.description,
      'Last Updated':new Date()
    });
    return { name: spec.name, id: templateId, action: 'updated (hops untouched)' };
  }
  templateId = 'tpl_' + Utilities.getUuid().substring(0, 8);
  _writeMMRow_('MOVEMENT_TEMPLATES', MOVEMENT_TEMPLATES_HEADERS, {
    'ID':          templateId,
    'Name':        spec.name,
    'Destination': spec.destination,
    'Direction':   spec.direction,
    'Description': spec.description,
    'Active':      'Yes',
    'Date Added':  new Date(),
    'Last Updated':new Date()
  });
  spec.hops.forEach(function(h) {
    _writeMMRow_('MOVEMENT_HOPS', MOVEMENT_HOPS_HEADERS, {
      'ID':                'hop_' + Utilities.getUuid().substring(0, 8),
      'Template ID':       templateId,
      'Order':             h.order,
      'Chain':             h.chain,
      'From Account':      h.from,
      'To Account':        h.to,
      'Responsible':       h.responsible,
      'Amount % of Total': h.pct,
      'Notes':             h.notes || ''
    });
  });
  _logAudit_('seedMovementTemplate', 'template', templateId, spec.name,
             'Seeded ' + spec.name + ' (' + spec.hops.length + ' hops)');
  return { name: spec.name, id: templateId, action: 'created', hops: spec.hops.length };
}


// ── Menu diagnostics ───────────────────────────────────────────────────────

// ── Wire message formatter ─────────────────────────────────────────────────
// Generates the numbered wire list Amanda pastes into her outgoing message.
// Just the numbered lines — she writes her own intro text ("the money is in"
// or similar) around it. Format:
//
//   1. NF6 Tiger Capital LLC (5319) to NF6 Family Holdings (7932): $100,000.00
//   2. NF6 Family Holdings (7932) to NF6 Joint Mgmt LLC (8972): $1,000.00
//   3. NF6 Joint Mgmt LLC (8972) to Michelle Personal: $333.33
//   ...
//
// The wire ordering matters — a hop shouldn't appear before all of its
// inputs. We use Kahn's algorithm (topological sort) with a tiebreak that
// puts smaller-percentage branches first, which produces the "small stuff
// first, then the big consolidation" flow Amanda used in her example.

function _wireExecutionOrder_(hops) {
  var byId = {};
  var deps = {};          // hopId → { producerId: true }
  var reverseDeps = {};   // hopId → { consumerId: true }
  var initialOrder = {};

  hops.forEach(function(h, i) {
    var id = String(h['ID'] || h.hopId);
    byId[id] = h;
    initialOrder[id] = i;
    deps[id] = {};
    reverseDeps[id] = {};
  });

  // Hop C depends on hop P if P.To == C.From (money must land at C.From
  // before C can wire it onward).
  var allFroms = {}, allTos = {};
  hops.forEach(function(consumer) {
    var cid = String(consumer['ID'] || consumer.hopId);
    var cFrom = String(consumer['From Account'] || consumer.fromAccount || '');
    allFroms[cFrom] = true;
    hops.forEach(function(producer) {
      var pid = String(producer['ID'] || producer.hopId);
      if (pid === cid) return;
      var pTo = String(producer['To Account'] || producer.toAccount || '');
      allTos[pTo] = true;
      if (pTo && pTo === cFrom) {
        deps[cid][pid] = true;
        reverseDeps[pid][cid] = true;
      }
    });
  });

  // Auto-detect direction from the graph shape:
  //   Fan-in  (many contributors → one destination) = "down" per Amanda's
  //           convention (money coming DOWN into the entity). Show LARGE
  //           chains first so Mike's 99% reads before the sibling 0.333%.
  //   Fan-out (one source → many recipients) = "up" (money going UP to
  //           owners). Show SMALL branches first so siblings read before
  //           the final Mike branch (matches Amanda's Tiger Down example).
  var roots  = Object.keys(allFroms).filter(function(a) { return !allTos[a]; });
  var leaves = Object.keys(allTos).filter(function(a) { return !allFroms[a]; });
  var isFanIn = roots.length > leaves.length;
  var sortSign = isFanIn ? -1 : 1;   // -1 = DESC, 1 = ASC

  var pending = Object.keys(byId);
  var result = [];
  while (pending.length) {
    var ready = pending.filter(function(id) { return Object.keys(deps[id]).length === 0; });
    if (!ready.length) {
      pending.sort(function(a, b) { return initialOrder[a] - initialOrder[b]; });
      pending.forEach(function(id) { result.push(byId[id]); });
      break;
    }
    ready.sort(function(a, b) {
      var pa = Number(byId[a]['Amount % of Total'] || byId[a].pctOfTotal) || 0;
      var pb = Number(byId[b]['Amount % of Total'] || byId[b].pctOfTotal) || 0;
      if (pa !== pb) return sortSign * (pa - pb);
      return initialOrder[a] - initialOrder[b];
    });
    var next = ready[0];
    result.push(byId[next]);
    pending.splice(pending.indexOf(next), 1);
    Object.keys(reverseDeps[next]).forEach(function(childId) {
      delete deps[childId][next];
    });
  }
  return result;
}

function _fmtUsdAmount_(n) {
  var neg = n < 0;
  var abs = Math.abs(n);
  return (neg ? '-' : '') + '$' + abs.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// Whole-dollar formatter (nearest integer, no decimals). Used in the
// movement-message output per Amanda's preference for clean round numbers.
function _fmtUsdWhole_(n) {
  var rounded = Math.round(Number(n) || 0);
  var neg = rounded < 0;
  var abs = Math.abs(rounded);
  return (neg ? '-' : '') + '$' + String(abs).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// Web-callable — generate the "here's the wire order" message for a template
// + total amount, WITHOUT needing an existing Movement record. Used by the
// New Movement wizard for a live preview. Same output the UI's "Copy Message"
// button will produce once a movement is saved.
function formatMovementPreview(templateId, totalAmount) {
  var detail = getMovementTemplateDetail(templateId);
  if (detail.error) return { error: detail.error };
  var total = Number(totalAmount) || 0;
  if (total <= 0) return { error: 'Total amount must be > 0.' };
  var ordered = _wireExecutionOrder_(detail.hops);
  var lines = [];
  ordered.forEach(function(h, i) {
    var pct = Number(h['Amount % of Total']) || 0;
    var amt = Math.round(total * pct / 100);   // whole dollars
    lines.push((i + 1) + '. ' + h['From Account'] + ' to ' + h['To Account'] + ': ' + _fmtUsdWhole_(amt));
  });
  // Footer: total landing at (or leaving from) the destination entity.
  var dest = String(detail.template && detail.template['Destination'] || '').trim();
  var dir  = String(detail.template && detail.template['Direction']   || '').toLowerCase();
  if (dest) {
    var destTotal = 0;
    ordered.forEach(function(h) {
      var pct = Number(h['Amount % of Total']) || 0;
      var amt = Math.round(total * pct / 100);
      if (dir === 'down' && String(h['To Account']||'') === dest) destTotal += amt;
      else if (dir === 'up' && String(h['From Account']||'') === dest) destTotal += amt;
    });
    if (destTotal) {
      var label = dir === 'up' ? 'Total leaving ' + dest : 'Total landing at ' + dest;
      lines.push('');
      lines.push(label + ': ' + _fmtUsdWhole_(destTotal));
    }
  }
  return {
    template: detail.template,
    totalAmount: total,
    wireCount: ordered.length,
    message: lines.join('\n')
  };
}

// Same idea but for an existing movement — uses the actual snapshotted wire
// amounts (which may differ from a fresh template calc if the template's
// percentages have been edited since the movement was created).
function formatMovementMessage(movementId) {
  var d = getMovementDetail(movementId);
  if (d.error) return { error: d.error };
  // getMovementDetail's wires already have hop fields (fromAccount, toAccount,
  // pctOfTotal, order). Feed those into the topo sort.
  var normalized = d.wires.map(function(w) {
    return {
      ID: w.wireId, hopId: w.hopId,
      'From Account': w.fromAccount, fromAccount: w.fromAccount,
      'To Account':   w.toAccount,   toAccount:   w.toAccount,
      'Amount % of Total': w.pctOfTotal, pctOfTotal: w.pctOfTotal,
      _amount: w.amount
    };
  });
  var ordered = _wireExecutionOrder_(normalized);
  var lines = [];
  ordered.forEach(function(w, i) {
    lines.push((i + 1) + '. ' + w['From Account'] + ' to ' + w['To Account'] + ': ' + _fmtUsdWhole_(w._amount));
  });
  var dest = String(d.template && d.template['Destination'] || '').trim();
  var dir  = String(d.template && d.template['Direction']   || '').toLowerCase();
  if (dest) {
    var destTotal = 0;
    ordered.forEach(function(w) {
      var amt = Math.round(Number(w._amount) || 0);
      if (dir === 'down' && String(w['To Account']||'') === dest) destTotal += amt;
      else if (dir === 'up' && String(w['From Account']||'') === dest) destTotal += amt;
    });
    if (destTotal) {
      var label = dir === 'up' ? 'Total leaving ' + dest : 'Total landing at ' + dest;
      lines.push('');
      lines.push(label + ': ' + _fmtUsdWhole_(destTotal));
    }
  }
  return {
    movement:  d.movement,
    template:  d.template,
    wireCount: ordered.length,
    message:   lines.join('\n')
  };
}

// Menu-callable — shows a numbered list of every template, prompts for the
// number, then prompts for a total amount, then pops the generated message.
// The tab UI (commit 2) will replace this with a proper dropdown; this is
// the placeholder for testing until then.
function menuPreviewMovementMessage() {
  var ui = SpreadsheetApp.getUi();
  var templates = getMovementTemplates();
  if (!templates.length) {
    ui.alert('No templates yet. Run "Money Movement → Seed Templates" first.');
    return;
  }
  var list = templates.map(function(t, i) {
    return '  ' + (i + 1) + '. ' + t['Name'] + '  [' + (t['Direction'] || '?') + ']';
  }).join('\n');
  var r1 = ui.prompt('Preview Movement Message',
    'Pick a template — enter a number 1-' + templates.length + ' (or type a name substring):\n\n' + list,
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var raw = String(r1.getResponseText() || '').trim();
  if (!raw) return;
  var match;
  var asNum = Number(raw);
  if (!isNaN(asNum) && asNum >= 1 && asNum <= templates.length) {
    match = templates[asNum - 1];
  } else {
    var q = raw.toLowerCase();
    match = templates.find(function(t) { return String(t['Name']).toLowerCase().indexOf(q) >= 0; });
  }
  if (!match) {
    ui.alert('No template found for "' + raw + '".\n\nAvailable:\n' + list);
    return;
  }
  var r2 = ui.prompt('Preview Movement Message',
    'Selected: ' + match['Name'] + '\n\nEnter the total amount to move (e.g. 100000 or $100,000):',
    ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var amt = Number(String(r2.getResponseText() || '').replace(/[$,\s]/g, '')) || 0;
  if (amt <= 0) { ui.alert('Amount must be > 0.'); return; }
  var preview = formatMovementPreview(match['ID'], amt);
  if (preview.error) { ui.alert(preview.error); return; }
  ui.alert(match['Name'] + ' — ' + _fmtUsdAmount_(amt), preview.message, ui.ButtonSet.OK);
}


// Menu-callable — rescale the Blue Panda templates' percentages so the user's
// input amount represents the MN Trust wire ($1,000,000) rather than the
// destination amount ($1,010,101). Matches Amanda's Excel which also takes
// the MN Trust wire as the input number and derives everything else.
// Applies to BOTH "Blue Panda FLP - Down" and "Blue Panda FLP - Up".
// Idempotent.
function fixBluePandaInputScale() {
  _requireEditor_();
  ensureMoneyMovementSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MOVEMENT_HOPS');
  if (!sheet || sheet.getLastRow() < 2) return;
  var templates = _getMMRows_('MOVEMENT_TEMPLATES', MOVEMENT_TEMPLATES_HEADERS);
  var bluePandaIds = templates
    .filter(function(t) { return /blue panda/i.test(String(t['Name']||'')); })
    .map(function(t) { return String(t['ID']); });
  if (!bluePandaIds.length) {
    try { SpreadsheetApp.getUi().alert('No Blue Panda templates found. Run Seed Templates first.'); } catch(e) {}
    return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), MOVEMENT_HOPS_HEADERS.length);
  var data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  var hdr = data[0];
  var iTpl   = hdr.indexOf('Template ID');
  var iFrom  = hdr.indexOf('From Account');
  var iTo    = hdr.indexOf('To Account');
  var iPct   = hdr.indexOf('Amount % of Total');
  if (iTpl < 0 || iFrom < 0 || iTo < 0 || iPct < 0) return;

  // Rule: given a hop's from/to accounts, return the correct pct for the
  // new "MN Trust = 100% of input" scaling. Returns null if no change.
  // Target pcts match Amanda's Excel formulas exactly:
  //   I = E * 1.0101/100     (BPMGMT consolidation = 1.0101% of input)
  //   H = I * 0.51            (Mike's BPMGMT leg = 0.515151% of input)
  //   G = I * 0.49            (Nancy's BPMGMT leg = 0.494949% of input)
  //   F = H                   (Mike → Rev Trust passes through same amount)
  //   D = E                   (Mike → MN Trust passes through same amount)
  function targetPct(from, to) {
    // Mike ↔ MN Trust / MN Trust ↔ Blue Panda — Mike's main chain
    if (/mn trust irrevocable/i.test(from + ' ' + to) &&
        (/michael.*personal/i.test(from + ' ' + to) || /blue panda/i.test(from + ' ' + to))) return 100.00;
    // BPMGMT consolidation (either direction)
    if (/bpmgmt/i.test(from) && /blue panda/i.test(to)) return 1.0101;
    if (/blue panda/i.test(from) && /bpmgmt/i.test(to)) return 1.0101;
    // Mike's BPMGMT branch: Mike ↔ Rev Trust ↔ BPMGMT (0.515151% of input)
    if ((/michael.*personal/i.test(from) && /revocable trust/i.test(to)) ||
        (/revocable trust/i.test(from) && /michael.*personal/i.test(to))) return 0.515151;
    if ((/revocable trust/i.test(from) && /bpmgmt/i.test(to)) ||
        (/bpmgmt/i.test(from) && /revocable trust/i.test(to))) return 0.515151;
    // Nancy's BPMGMT branch (0.494949% of input)
    if ((/nancy.*personal/i.test(from) && /bpmgmt/i.test(to)) ||
        (/bpmgmt/i.test(from) && /nancy.*personal/i.test(to))) return 0.494949;
    return null;
  }

  var changed = 0;
  var summary = [];
  for (var r = 1; r < data.length; r++) {
    if (bluePandaIds.indexOf(String(data[r][iTpl])) < 0) continue;
    var from = String(data[r][iFrom] || '');
    var to   = String(data[r][iTo] || '');
    var want = targetPct(from, to);
    if (want === null) continue;
    var cur = Number(data[r][iPct]) || 0;
    if (Math.abs(cur - want) > 0.0000001) {
      sheet.getRange(r + 1, iPct + 1).setValue(want);
      changed++;
      summary.push('  ' + from + ' → ' + to + ': ' + cur + ' → ' + want);
    }
  }
  var msg = 'Updated ' + changed + ' Blue Panda hop(s). Input now represents the MN Trust wire amount.\n\n' + (summary.length ? summary.join('\n') : '(already canonical)');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Blue Panda Rescaled', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}


// Known account-label corrections. Add a new entry any time Amanda confirms
// an account number from Plaid. Existing rows are updated in-place by the
// resync menu function; the seed's own labels stay current for fresh seeds.
// Keys are the OLD (or short/no-ID) label, values are the CORRECT label.
var _MOVEMENT_ACCOUNT_RELABELS = {
  // 2019 MN Family Revocable Trust — correct ID is 3333 (was misseeded 7013)
  '2019 MN Family Revocable Trust (7013)':  '2019 MN Family Revocable Trust (3333)',
  '2019 Family Revocable Trust (7013)':     '2019 Family Revocable Trust (3333)',
  '2019 MN Family Revocable Trust':         '2019 MN Family Revocable Trust (3333)',
  // Mike's personal account — correct is 1319 (was misseeded 6916)
  'Michael Nguyen Personal (6916)':         'Michael Nguyen Personal (1319)',
  'Michael Nguyen Personal':                'Michael Nguyen Personal (1319)',
  // NF6 Family Holdings — Amanda uses the Chase account (7932); the Charles
  // Schwab one (316) is retired. Normalize all variants to 7932.
  'NF6 Family Holdings (316)':              'NF6 Family Holdings (7932)',
  'NF6 Family Holding (316)':               'NF6 Family Holdings (7932)',
  'NF6 Family Holdings':                    'NF6 Family Holdings (7932)',
  'NF6 Family Holding':                     'NF6 Family Holdings (7932)',
  // NF PR SJ LLC — Amanda corrected the account number in her sheet from
  // my seed's (808) to the actual (5297). Add both directions of the
  // normalization so any leftover (808) rows get pulled up.
  'NF PR SJ LLC (808)':                     'NF PR SJ LLC (5297)',
  'NF PR SJ LLC':                           'NF PR SJ LLC (5297)'
};

// ── Auto-checkoff (Plaid → wire matching) ──────────────────────────────────
// For every movement whose status is Planning or In Progress, scan every
// wire whose status is Pending. Try to find a matching real transaction in
// PLAID_TRANSACTIONS. If exactly one candidate matches on amount + account
// + date window, auto-mark the wire Sent and stamp the matched Plaid txn ID.
//
// Matching rules:
//   amount:  real Plaid txn amount within $1 of the wire's expected amount.
//            Compared using absolute values since Plaid flips signs for
//            inflow/outflow but our wires store signed amounts anyway (the
//            amount itself is always positive on the wire row).
//   account: real txn's Account field contains EITHER the last 4 digits of
//            the wire's From Account OR the To Account. This way wires
//            between two of Amanda's own accounts match from either side.
//   date:    real txn's Date is between (movement.Created At - 2 days) and
//            today + 1 day. A 2-day lookback handles the case where she
//            created the movement after the wire already went through.
//   dedup:   a Plaid txn ID that's already linked to any wire is skipped
//            for all other wires — no double-counting one deposit as two
//            wires.
//
// Idempotent + safe to re-run. Only touches wires still in Pending status.
// Already-matched wires (Sent / Confirmed) are left alone.

function autoCheckoffMovementWires() {
  _requireEditor_();
  ensureMoneyMovementSheets_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Load every pending wire across every active movement, with the
  //    accounts + amount info from their hop joined in.
  var hopsRows = _getMMRows_('MOVEMENT_HOPS', MOVEMENT_HOPS_HEADERS);
  var hopById = {};
  hopsRows.forEach(function(h) { hopById[String(h['ID'])] = h; });
  var movesRows = _getMMRows_('MOVEMENTS', MOVEMENTS_HEADERS)
    .filter(function(m) {
      var s = String(m['Status'] || '').toLowerCase();
      return s === 'planning' || s === 'in progress';
    });
  var moveById = {};
  movesRows.forEach(function(m) { moveById[String(m['ID'])] = m; });
  var activeMoveIds = Object.keys(moveById);
  if (!activeMoveIds.length) return { success: true, scanned: 0, matched: 0, note: 'No active movements.' };

  var wiresSheet = ss.getSheetByName('MOVEMENT_WIRES');
  if (!wiresSheet || wiresSheet.getLastRow() < 2) {
    return { success: true, scanned: 0, matched: 0, note: 'MOVEMENT_WIRES is empty.' };
  }
  var wLastCol = Math.max(wiresSheet.getLastColumn(), MOVEMENT_WIRES_HEADERS.length);
  var wData = wiresSheet.getRange(1, 1, wiresSheet.getLastRow(), wLastCol).getValues();
  var wHdr = wData[0];
  var iwId     = wHdr.indexOf('ID');
  var iwMove   = wHdr.indexOf('Movement ID');
  var iwHop    = wHdr.indexOf('Hop ID');
  var iwAmt    = wHdr.indexOf('Amount');
  var iwStatus = wHdr.indexOf('Status');
  var iwPlaid  = wHdr.indexOf('Plaid Txn ID');
  var iwSent   = wHdr.indexOf('Sent Date');
  var iwMarkedBy = wHdr.indexOf('Marked Done By');
  var iwMarkedAt = wHdr.indexOf('Marked Done At');

  // Track all Plaid txn IDs already linked to any wire — don't rematch them.
  var alreadyLinked = {};
  for (var r = 1; r < wData.length; r++) {
    var pid = String(wData[r][iwPlaid] || '').trim();
    if (pid) alreadyLinked[pid] = true;
  }

  var candidateWires = [];
  for (var r2 = 1; r2 < wData.length; r2++) {
    var mid = String(wData[r2][iwMove] || '');
    if (!moveById[mid]) continue;
    var st = String(wData[r2][iwStatus] || '').toLowerCase();
    if (st !== 'pending') continue;
    var hop = hopById[String(wData[r2][iwHop] || '')];
    if (!hop) continue;
    candidateWires.push({
      rowNum:    r2 + 1,
      wireId:    String(wData[r2][iwId]),
      move:      moveById[mid],
      hop:       hop,
      amount:    Math.abs(Number(wData[r2][iwAmt]) || 0)
    });
  }
  if (!candidateWires.length) return { success: true, scanned: 0, matched: 0, note: 'No Pending wires on active movements.' };

  // 2. Load PLAID_TRANSACTIONS once.
  var pSheet = ss.getSheetByName('PLAID_TRANSACTIONS');
  if (!pSheet || pSheet.getLastRow() < 2) {
    return { success: true, scanned: candidateWires.length, matched: 0, note: 'PLAID_TRANSACTIONS is empty — run Sync ALL Plaid Transactions first.' };
  }
  var pHdr = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0];
  var ipId    = pHdr.indexOf('Transaction ID');
  var ipDate  = pHdr.indexOf('Date');
  var ipAcct  = pHdr.indexOf('Account');
  var ipAmt   = pHdr.indexOf('Amount USD');
  var pData = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, pHdr.length).getValues();

  // 3. Try to match each pending wire.
  var matched = 0;
  var summary = [];
  var now = new Date();
  var me  = _currentUserEmail_();
  var ipName = pHdr.indexOf('Name');

  candidateWires.forEach(function(cw) {
    var fromAcct = String(cw.hop['From Account'] || '');
    var toAcct   = String(cw.hop['To Account']   || '');
    var fromKey  = _lastFourDigits_(fromAcct);
    var toKey    = _lastFourDigits_(toAcct);
    if (!fromKey && !toKey) return;   // no account signature to match

    // Date window: created - 2 days through today + 1 day.
    var createdAt = cw.move['Created At'] instanceof Date ? cw.move['Created At'] : new Date(cw.move['Created At']);
    var dateMin = new Date(createdAt.getTime() - 2 * 86400000);
    var dateMax = new Date(now.getTime() + 86400000);

    var hits = [];
    for (var pr = 0; pr < pData.length; pr++) {
      var pTxnId = String(pData[pr][ipId] || '');
      if (!pTxnId || alreadyLinked[pTxnId]) continue;
      var acct = String(pData[pr][ipAcct] || '');
      var acctLc = acct.toLowerCase();
      var matchesFrom = fromKey && acctLc.indexOf(fromKey) >= 0;
      var matchesTo   = toKey   && acctLc.indexOf(toKey)   >= 0;
      if (!matchesFrom && !matchesTo) continue;
      var rawAmt = Number(pData[pr][ipAmt]) || 0;
      var pamt = Math.abs(rawAmt);
      if (Math.abs(pamt - cw.amount) > 1.0) continue;   // $1 tolerance
      var d = pData[pr][ipDate] instanceof Date ? pData[pr][ipDate] : new Date(pData[pr][ipDate]);
      if (isNaN(d.getTime())) continue;
      if (d < dateMin || d > dateMax) continue;
      hits.push({
        txnId: pTxnId, date: d, account: acct, amount: pamt,
        signedAmount: rawAmt,
        name: ipName >= 0 ? String(pData[pr][ipName] || '') : '',
        matchesFrom: matchesFrom,
        matchesTo:   matchesTo
      });
    }

    // Disambiguation step 1: internal transfer dedup. If exactly 2 hits
    // on the same date with matching absolute amount but opposite signs
    // and one matches From account / other matches To account, they're
    // the same wire seen from both ends. Pick the OUTGOING (negative)
    // side as the canonical confirmation.
    if (hits.length === 2) {
      var h0 = hits[0], h1 = hits[1];
      var sameDate = h0.date.getTime() === h1.date.getTime();
      var sameAmt  = Math.abs(h0.amount - h1.amount) < 0.5;
      var oppositeSigns = (h0.signedAmount < 0) !== (h1.signedAmount < 0);
      var twoEnds = (h0.matchesFrom && h1.matchesTo) || (h0.matchesTo && h1.matchesFrom);
      if (sameDate && sameAmt && oppositeSigns && twoEnds) {
        hits = [h0.signedAmount < 0 ? h0 : h1];
      }
    }

    // Disambiguation step 2: sibling name hint. For sibling wires that
    // are all identical amount + account (e.g. Joint Mgmt → three siblings
    // all at $333.33 on same day), narrow by matching the sibling's first
    // name against the real Plaid transaction's recipient name.
    if (hits.length > 1) {
      var nameHint = _siblingNameHint_(toAcct) || _siblingNameHint_(fromAcct);
      if (nameHint) {
        var narrowed = hits.filter(function(h) {
          return h.name.toLowerCase().indexOf(nameHint) >= 0;
        });
        if (narrowed.length === 1) hits = narrowed;
      }
    }

    if (hits.length === 1) {
      var hit = hits[0];
      wiresSheet.getRange(cw.rowNum, iwStatus + 1).setValue('Sent');
      wiresSheet.getRange(cw.rowNum, iwPlaid + 1).setValue(hit.txnId);
      wiresSheet.getRange(cw.rowNum, iwSent + 1).setValue(hit.date);
      if (iwMarkedBy >= 0) wiresSheet.getRange(cw.rowNum, iwMarkedBy + 1).setValue(me + ' (auto)');
      if (iwMarkedAt >= 0) wiresSheet.getRange(cw.rowNum, iwMarkedAt + 1).setValue(now);
      alreadyLinked[hit.txnId] = true;
      matched++;
      summary.push('  ✓ ' + fromAcct + ' → ' + toAcct + '  $' + cw.amount.toFixed(2) +
                   '  ← ' + Utilities.formatDate(hit.date, Session.getScriptTimeZone(), 'yyyy-MM-dd') + '  ' + hit.account);
    } else if (hits.length > 1) {
      summary.push('  ? ' + fromAcct + ' → ' + toAcct + '  $' + cw.amount.toFixed(2) +
                   '  (' + hits.length + ' candidates — left Pending)');
    }
  });

  // 4. Auto-advance movement status based on how many wires are now done.
  //    Planning  → In Progress    when at least one wire is Sent/Confirmed
  //    In Progress → Complete     when every wire on the movement is Sent/Confirmed
  //    Doesn't downgrade — a movement that got to Complete stays Complete even
  //    if someone later toggles a wire back to Pending (that re-triggers a
  //    reopen on the next status re-eval).
  var affectedMoveIds = {};
  candidateWires.forEach(function(cw) { affectedMoveIds[String(cw.move['ID'])] = true; });
  var statusChanges = [];
  Object.keys(affectedMoveIds).forEach(function(mid) {
    var change = _recomputeMovementStatus_(mid);
    if (change) statusChanges.push(change);
  });

  return {
    success: true,
    scanned: candidateWires.length,
    matched: matched,
    summary: summary,
    statusChanges: statusChanges
  };
}

// Recompute a movement's status from its wires' current statuses. Returns
// { movementId, oldStatus, newStatus } if the status changed, else null.
// Safe to call anytime (after auto-match, after a manual wire toggle, from
// a menu button).
function _recomputeMovementStatus_(movementId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var moves = _getMMRows_('MOVEMENTS', MOVEMENTS_HEADERS);
  var move = moves.find(function(m) { return String(m['ID']) === String(movementId); });
  if (!move) return null;
  var currentStatus = String(move['Status'] || 'Planning');
  // Don't auto-touch a Cancelled movement.
  if (currentStatus === 'Cancelled') return null;

  var wires = _getMMRows_('MOVEMENT_WIRES', MOVEMENT_WIRES_HEADERS)
    .filter(function(w) { return String(w['Movement ID']) === String(movementId); });
  if (!wires.length) return null;
  var total = wires.length;
  var done  = wires.filter(function(w) {
    var s = String(w['Status']||'').toLowerCase();
    return s === 'sent' || s === 'confirmed';
  }).length;
  var skipped = wires.filter(function(w) {
    return String(w['Status']||'').toLowerCase() === 'skipped';
  }).length;

  var newStatus;
  if (done + skipped === total)         newStatus = 'Complete';
  else if (done === 0)                   newStatus = 'Planning';
  else                                   newStatus = 'In Progress';

  if (newStatus === currentStatus) return null;
  _updateMMRow_('MOVEMENTS', MOVEMENTS_HEADERS, movementId, {
    'Status': newStatus,
    'Last Updated': new Date()
  });
  _logAudit_('autoStatusChange', 'movement', movementId, move['Notes'] || '',
             'Status ' + currentStatus + ' → ' + newStatus + ' (' + done + '/' + total + ' done)');
  return { movementId: movementId, oldStatus: currentStatus, newStatus: newStatus };
}

// Web-callable: look up a Plaid transaction by its Transaction ID so the
// UI can show "here's the real transaction that matched this wire" when
// Amanda taps a ✓P indicator. Returns the row as a plain object, or null.
function getPlaidTxnDetails(txnId) {
  if (!txnId) return null;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PLAID_TRANSACTIONS');
  if (!sheet || sheet.getLastRow() < 2) return null;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var iId = headers.indexOf('Transaction ID');
  if (iId < 0) return null;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][iId]) !== String(txnId)) continue;
    var obj = {};
    headers.forEach(function(h, i) {
      var v = data[r][i];
      if (v instanceof Date) v = v.toISOString();
      // Omit the Raw JSON column — it's huge and the UI doesn't need it.
      if (h === 'Raw JSON') return;
      obj[h] = v;
    });
    return obj;
  }
  return null;
}


// Pull the last 4 digits off an account label like "NF6 Tiger Capital LLC
// (5319)" or "Michael Nguyen Personal (1319)". Returns '' if the label has
// no 4-digit tail. Used as the substring key for matching against the
// Plaid account label (which always ends in "···<mask>").
function _lastFourDigits_(label) {
  var m = String(label || '').match(/(\d{4})\D*$/);
  return m ? m[1] : '';
}

// Extract a sibling first-name token from a wire account label if present.
// Used to disambiguate between identical sibling wires on the same day
// (e.g. three $333.33 outflows from Joint Mgmt → each sibling). The Plaid
// transaction description for the real wire will contain the recipient's
// first name, so narrowing on this field flips an ambiguous "3 candidates"
// case into "1 match".
function _siblingNameHint_(label) {
  var m = String(label || '').toLowerCase().match(/\b(michelle|nancy|david|michael|mike)\b/);
  return m ? m[1] : '';
}

// Menu-callable wrapper with an alert summary.
function autoCheckoffMovementWiresMenu() {
  var res = autoCheckoffMovementWires();
  var lines = ['Auto-checkoff complete.', '',
               'Pending wires scanned: ' + res.scanned,
               'Wires auto-matched:    ' + res.matched];
  if (res.note) { lines.push(''); lines.push(res.note); }
  if (res.summary && res.summary.length) {
    lines.push('');
    lines.push('DETAIL');
    res.summary.forEach(function(s) { lines.push(s); });
  }
  try { SpreadsheetApp.getUi().alert('Auto-checkoff', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); }
  catch(e) { Logger.log(lines.join('\n')); }
}


// Menu-callable — walk every hop row and replace any From/To Account whose
// exact string matches a known-wrong label with the correct one. Idempotent:
// rows already carrying the correct label are untouched. Uses exact-string
// match (not substring) so a cell like "Michael Nguyen Personal (1319)"
// won't accidentally re-match the "Michael Nguyen Personal" no-ID entry.
function resyncMovementAccountLabels() {
  _requireEditor_();
  ensureMoneyMovementSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MOVEMENT_HOPS');
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('MOVEMENT_HOPS empty.');
    return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), MOVEMENT_HOPS_HEADERS.length);
  var data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  var hdr = data[0];
  var iFrom  = hdr.indexOf('From Account');
  var iTo    = hdr.indexOf('To Account');
  var iChain = hdr.indexOf('Chain');
  if (iFrom < 0 || iTo < 0) return;

  var changed = 0;
  var summary = [];
  for (var r = 1; r < data.length; r++) {
    var chain = String(data[r][iChain] || '');
    ['From', 'To'].forEach(function(side) {
      var col = side === 'From' ? iFrom : iTo;
      var old = String(data[r][col] || '');
      var _new = _MOVEMENT_ACCOUNT_RELABELS[old];
      if (_new && _new !== old) {
        sheet.getRange(r + 1, col + 1).setValue(_new);
        changed++;
        summary.push('  [' + chain + '] ' + side + ': "' + old + '" → "' + _new + '"');
      }
    });
  }
  var msg = 'Updated ' + changed + ' account label(s).\n\n' + (summary.length ? summary.join('\n') : '(all labels already correct)');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Account Labels Resynced', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}


// Menu-callable — rename + swap direction on existing templates so they
// match Amanda's mental model:
//   Up   = money going UP out of the entity to owners  (was called "down")
//   Down = money coming DOWN into the entity from owners (was called "up")
// Handles every prior naming scheme in one shot: the original
// Contribution/Distribution names AND the intermediate " - Up / - Down"
// names I had backwards. Idempotent.
function renameMovementTemplatesToDirection() {
  _requireEditor_();
  ensureMoneyMovementSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MOVEMENT_TEMPLATES');
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastCol = Math.max(sheet.getLastColumn(), MOVEMENT_TEMPLATES_HEADERS.length);
  var data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  var hdr = data[0];
  var iName = hdr.indexOf('Name');
  var iDir  = hdr.indexOf('Direction');
  if (iName < 0 || iDir < 0) return;

  // Map to CANONICAL name + direction (Amanda's convention).
  // Every prior variant should resolve to one of these targets.
  var CANONICAL = {
    // Blue Panda FLP
    'Blue Panda FLP Contribution':                { name: 'Blue Panda FLP - Down', dir: 'down' },
    'Blue Panda FLP Distribution (Down)':         { name: 'Blue Panda FLP - Up',   dir: 'up'   },
    'Blue Panda FLP - Up':                        { name: 'Blue Panda FLP - Down', dir: 'down' },
    'Blue Panda FLP - Down':                      { name: 'Blue Panda FLP - Up',   dir: 'up'   },
    // NF PR SJ
    'NF PR SJ Wire':                              { name: 'NF PR SJ - Down',       dir: 'down' },
    'NF PR SJ Distribution (Down)':               { name: 'NF PR SJ - Up',         dir: 'up'   },
    'NF PR SJ - Up':                              { name: 'NF PR SJ - Down',       dir: 'down' },
    'NF PR SJ - Down':                            { name: 'NF PR SJ - Up',         dir: 'up'   },
    // Tiger Capital
    'NF6 Tiger Capital Distribution (Down)':      { name: 'NF6 Tiger Capital - Up',   dir: 'up'   },
    'NF6 Tiger Capital Contribution (Up)':        { name: 'NF6 Tiger Capital - Down', dir: 'down' },
    'NF6 Tiger Capital - Down':                   { name: 'NF6 Tiger Capital - Up',   dir: 'up'   },
    'NF6 Tiger Capital - Up':                     { name: 'NF6 Tiger Capital - Down', dir: 'down' }
  };
  var changed = 0;
  var summary = [];
  for (var r = 1; r < data.length; r++) {
    var oldName = String(data[r][iName] || '');
    var oldDir  = String(data[r][iDir] || '');
    var target = CANONICAL[oldName];
    if (!target) continue;
    var didChange = false;
    if (target.name !== oldName) { sheet.getRange(r + 1, iName + 1).setValue(target.name); didChange = true; }
    if (target.dir  !== oldDir)  { sheet.getRange(r + 1, iDir + 1).setValue(target.dir);   didChange = true; }
    if (didChange) {
      changed++;
      summary.push('  "' + oldName + '" [' + oldDir + '] → "' + target.name + '" [' + target.dir + ']');
    }
  }
  var msg = 'Updated ' + changed + ' template(s).\n\n' + (summary.length ? summary.join('\n') : '(all already canonical)');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Templates Renamed', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}


// Menu-callable — canonicalize known percentages on every existing hop. Now
// scoped correctly so it doesn't accidentally overwrite non-1/3 sibling
// wires (Nancy's Blue Panda 0.49% got clobbered by an earlier broad version
// of this function — see the Nancy Blue Panda restore below).
//
// Rules applied:
//   • Sibling → NF6 Joint Mgmt LLC (or reverse): 0.333333% for
//     Michelle/Nancy, 0.333334% for David (third sibling absorbs the
//     rounding so $100k → $333.33 not $333.00).
//   • Nancy Nguyen Personal → BPMGMT (5150): 0.49% (Blue Panda BPMGMT
//     split with Mike getting 51% / Nancy 49% of the 1% slice).
//
// Rows that don't match either rule are left alone. Idempotent.
function fixSiblingPercentages() {
  _requireEditor_();
  ensureMoneyMovementSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MOVEMENT_HOPS');
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('MOVEMENT_HOPS empty.');
    return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), MOVEMENT_HOPS_HEADERS.length);
  var data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  var hdr = data[0];
  var iFrom  = hdr.indexOf('From Account');
  var iTo    = hdr.indexOf('To Account');
  var iPct   = hdr.indexOf('Amount % of Total');
  var iChain = hdr.indexOf('Chain');
  if (iFrom < 0 || iTo < 0 || iPct < 0) return;

  var SIBLING_RE   = /\b(david|michelle|nancy)\b[\s\w()]*personal/i;
  var JOINT_MGMT_RE = /nf6 joint mgmt/i;
  var changed = 0;
  var summary = [];
  for (var r = 1; r < data.length; r++) {
    var from = String(data[r][iFrom] || '');
    var to   = String(data[r][iTo] || '');
    var chain = String(data[r][iChain] || '');
    var oldPct = Number(data[r][iPct]) || 0;

    var targetPct = null;
    // Rule 1: Nancy Nguyen Personal → BPMGMT (Blue Panda 49% slice)
    if (/nancy nguyen personal/i.test(from) && /bpmgmt/i.test(to)) {
      targetPct = 0.49;
    }
    // Rule 2: sibling ↔ Joint Mgmt (the actual 1/3 splits)
    else if ((SIBLING_RE.test(from) || SIBLING_RE.test(to)) &&
             (JOINT_MGMT_RE.test(from) || JOINT_MGMT_RE.test(to))) {
      var isDavid = /\bdavid\b/i.test(from) || /\bdavid\b/i.test(to);
      targetPct = isDavid ? 0.333334 : 0.333333;
    }
    // Other rows (Nancy → Joint Mgmt is caught by Rule 2, Nancy → BPMGMT
    // by Rule 1, etc.): don't touch.
    if (targetPct === null) continue;

    if (Math.abs(oldPct - targetPct) > 0.0000001) {
      sheet.getRange(r + 1, iPct + 1).setValue(targetPct);
      changed++;
      summary.push('  ' + chain + ': ' + oldPct + ' → ' + targetPct);
    }
  }
  var msg = 'Updated ' + changed + ' wire percentage(s).\n\n' + (summary.length ? summary.join('\n') : '(already correct)');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Wire Percentages Canonicalized', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}


// Menu-callable — apply Amanda's responsible-person rule to every existing
// hop in MOVEMENT_HOPS. Rule based on FROM Account (not Chain):
//
//   • If From Account looks like a sibling personal account (contains
//     "David", "Michelle", or "Nancy" AND contains "Personal") → Ben.
//     These are money-OUT wires from a sibling's own account — Ben handles
//     those with the siblings directly.
//   • Every other wire → Amanda. This includes wires that ARRIVE at a
//     sibling's personal account (e.g. Joint Mgmt → David Personal on a
//     Tiger DOWN distribution) — Amanda dispatches those from the family
//     LLC, so they're Amanda's responsibility even though a sibling is on
//     the receiving end.
//
// Idempotent + safe to re-run any time preferences change.
function updateMovementResponsibles() {
  _requireEditor_();
  ensureMoneyMovementSheets_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MOVEMENT_HOPS');
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('MOVEMENT_HOPS is empty — nothing to update.');
    try { SpreadsheetApp.getUi().alert('MOVEMENT_HOPS is empty. Run Seed Templates first.'); } catch(e) {}
    return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), MOVEMENT_HOPS_HEADERS.length);
  var data    = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  var hdr     = data[0];
  var iFrom   = hdr.indexOf('From Account');
  var iResp   = hdr.indexOf('Responsible');
  if (iFrom < 0 || iResp < 0) {
    Logger.log('MOVEMENT_HOPS missing From Account or Responsible column.');
    return;
  }
  // Regex: sibling FIRST name + word char + Personal (matches "David Personal",
  // "Nancy Nguyen Personal", "Michelle Personal", etc.). Excludes cases where
  // the sibling's name appears elsewhere in the account label but "Personal"
  // is not in the string, to be safe.
  var SIBLING_PERSONAL_RE = /\b(david|michelle|nancy)\b[\s\w()]*personal/i;
  var changed = 0;
  var summary = [];
  for (var r = 1; r < data.length; r++) {
    var from    = String(data[r][iFrom] || '');
    var oldResp = String(data[r][iResp] || '');
    var newResp = SIBLING_PERSONAL_RE.test(from) ? 'Ben' : 'Amanda';
    if (oldResp !== newResp) {
      sheet.getRange(r + 1, iResp + 1).setValue(newResp);
      changed++;
      summary.push('  from "' + from + '": ' + (oldResp || '(blank)') + ' → ' + newResp);
    }
  }
  var msg = 'Updated ' + changed + ' hop(s).\n\n' + (summary.length ? summary.join('\n') : '(everything was already correct)');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Responsibles Updated', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
}


function debugMovementTemplates() {
  var templates = _getMMRows_('MOVEMENT_TEMPLATES', MOVEMENT_TEMPLATES_HEADERS);
  var lines = ['MOVEMENT TEMPLATES (' + templates.length + ')', ''];
  templates.forEach(function(t) {
    var hops = getMovementHops(t['ID']);
    lines.push('  • ' + t['Name'] + '  [' + t['Direction'] + ']');
    lines.push('      Destination: ' + t['Destination']);
    lines.push('      ID: ' + t['ID'] + '  ·  ' + hops.length + ' hop(s)');
    var byChain = {};
    hops.forEach(function(h) {
      var c = h['Chain'] || '(no chain)';
      byChain[c] = byChain[c] || [];
      byChain[c].push(h);
    });
    Object.keys(byChain).forEach(function(chain) {
      lines.push('      ─── ' + chain + ' ───');
      byChain[chain].forEach(function(h) {
        lines.push('        [' + h['Order'] + '] ' + h['From Account'] + ' → ' + h['To Account'] +
                   '  ·  ' + h['Amount % of Total'] + '%  ·  ' + (h['Responsible']||'?'));
      });
    });
    lines.push('');
  });
  Logger.log(lines.join('\n'));
  try { SpreadsheetApp.getUi().alert('Movement Templates', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); }
  catch(e) { Logger.log('(No UI — output above is in execution log.)'); }
}
