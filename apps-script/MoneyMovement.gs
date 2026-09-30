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

// Returns every active template (for the "New Movement" picker).
function getMovementTemplates() {
  return _getMMRows_('MOVEMENT_TEMPLATES', MOVEMENT_TEMPLATES_HEADERS)
    .filter(function(t) {
      var a = String(t['Active'] || '').toLowerCase();
      return a !== 'no' && a !== 'false' && a !== '0';
    });
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
  return { success: ok };
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
  return {
    movement: move,
    template: template,
    wires:    wires,
    progress: { done: done, total: wires.length, pct: wires.length ? Math.round(done / wires.length * 100) : 0 }
  };
}

// List all movements (for the tab's landing view). Newest first.
function getMovements() {
  return _getMMRows_('MOVEMENTS', MOVEMENTS_HEADERS)
    .sort(function(a, b) {
      var da = a['Created At'] instanceof Date ? a['Created At'].getTime() : 0;
      var db = b['Created At'] instanceof Date ? b['Created At'].getTime() : 0;
      return db - da;
    });
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
//   NF PR SJ LLC (808) has the same ultimate ownership (99% Mike, 0.333% per
//   sibling) through the same NF6 Family Holdings (316) parent.
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

  // 1. Blue Panda FLP Contribution
  results.push(_seedTemplate_({
    name: 'Blue Panda FLP Contribution',
    destination: 'Blue Panda FLP (8686)',
    direction: 'up',
    description: 'Money flowing IN to Blue Panda FLP. 99% via MN Trust Irrevocable (100% Mike). 1% via BPMGMT (5150), which is 51% Mike / 49% Nancy.',
    hops: [
      // Chain A: Mike via MN Trust (99% of total)
      { order: 10, chain:'Mike (via MN Trust)',  from:'Michael Nguyen Personal (1319)',   to:'MN Trust Irrevocable',                  responsible:'Amanda', pct: 99.00, notes:'' },
      { order: 20, chain:'Mike (via MN Trust)',  from:'MN Trust Irrevocable',             to:'Blue Panda FLP (8686)',                 responsible:'Amanda', pct: 99.00, notes:'' },
      // Chain B: Mike via BPMGMT (0.51% of total = 51% of the 1% BPMGMT slice)
      { order: 30, chain:'Mike (via BPMGMT)',    from:'Michael Nguyen Personal (1319)',   to:'2019 MN Family Revocable Trust (3333)', responsible:'Amanda', pct:  0.51, notes:'51% of the 1% BPMGMT slice.' },
      { order: 40, chain:'Mike (via BPMGMT)',    from:'2019 MN Family Revocable Trust (3333)', to:'BPMGMT (5150)',                    responsible:'Amanda', pct:  0.51, notes:'' },
      // Chain C: Nancy via BPMGMT (0.49% of total = 49% of the 1% BPMGMT slice)
      // Amanda's rule: sibling personal-account wires are Ben's responsibility.
      { order: 30, chain:'Nancy (via BPMGMT)',   from:'Nancy Nguyen Personal',            to:'BPMGMT (5150)',                         responsible:'Ben',    pct:  0.49, notes:'49% of the 1% BPMGMT slice.' },
      // Merger + final leg: BPMGMT → Blue Panda FLP (1% total)
      { order: 50, chain:'BPMGMT consolidation', from:'BPMGMT (5150)',                    to:'Blue Panda FLP (8686)',                 responsible:'Amanda', pct:  1.00, notes:'After Mike + Nancy contributions merge at BPMGMT.' }
    ]
  }, haveByName));

  // 2. NF PR SJ Wire (up)
  results.push(_seedTemplate_({
    name: 'NF PR SJ Wire',
    destination: 'NF PR SJ LLC (808)',
    direction: 'up',
    description: 'Money flowing IN to NF PR SJ LLC. 99% via Mike (through 2019 MN Family Rev Trust → NF6 Family Holdings). 1% via siblings (each 1/3 through Personal → NF6 Joint Mgmt → NF6 Family Holdings).',
    hops: [
      // Mike chain (99%)
      { order: 10, chain:'Mike (via Rev Trust)', from:'Michael Nguyen Personal (1319)',     to:'2019 Family Revocable Trust (3333)', responsible:'Amanda', pct: 99.00, notes:'' },
      { order: 20, chain:'Mike (via Rev Trust)', from:'2019 MN Family Revocable Trust',     to:'NF6 Family Holdings (316)',           responsible:'Amanda', pct: 99.00, notes:'' },
      // Siblings (1/3 of 1% each = 0.333333% each). Use more precision so
      // $100,000 × 0.333333% = $333.33 rather than $333.00 with a truncated
      // 0.333%. Third sibling absorbs the rounding to preserve the sum.
      // Sibling personal-account wires are Ben's responsibility.
      { order: 30, chain:'David (via Joint Mgmt)',    from:'David Personal',    to:'NF6 Joint Mgmt LLC (8972)', responsible:'Ben', pct: 0.333333, notes:'' },
      { order: 30, chain:'Nancy (via Joint Mgmt)',    from:'Nancy Personal',    to:'NF6 Joint Mgmt LLC (8972)', responsible:'Ben', pct: 0.333333, notes:'' },
      { order: 30, chain:'Michelle (via Joint Mgmt)', from:'Michelle Personal', to:'NF6 Joint Mgmt LLC (8972)', responsible:'Ben', pct: 0.333334, notes:'' },
      // Merger + final leg
      { order: 40, chain:'Joint Mgmt consolidation', from:'NF6 Joint Mgmt LLC (8972)', to:'NF6 Family Holdings (316)', responsible:'Amanda', pct: 1.00, notes:'After 3 sibling contributions merge.' },
      { order: 50, chain:'Final leg',                from:'NF6 Family Holdings (316)', to:'NF PR SJ LLC (808)',        responsible:'Amanda', pct:100.00, notes:'Full amount to destination.' }
    ]
  }, haveByName));

  // 3. NF6 Tiger Capital — DOWN direction (distribute FROM Tiger Capital TO
  //    individuals). Same ownership as NF PR SJ but reversed hops.
  results.push(_seedTemplate_({
    name: 'NF6 Tiger Capital Distribution (Down)',
    destination: 'NF6 Tiger Capital LLC (5319)',
    direction: 'down',
    description: 'Distribution OUT of Tiger Capital, 99% to Mike + 0.333% each to Michelle/Nancy/David. Money flows down through Family Holdings → (Rev Trust for Mike | Joint Mgmt for siblings) → personal accounts.',
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

  // 4. NF6 Tiger Capital — UP direction (contributions INTO Tiger Capital).
  //    Same ownership, hops reversed from #3.
  results.push(_seedTemplate_({
    name: 'NF6 Tiger Capital Contribution (Up)',
    destination: 'NF6 Tiger Capital LLC (5319)',
    direction: 'up',
    description: 'Contribution INTO Tiger Capital. Mike wires 99%, each sibling wires 0.333%. Each contribution flows up through Family Holdings.',
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
// Amanda writes a message like this every time she does a movement, so
// everyone involved knows what's happening in what order:
//
//   The money is in. I'll move it as follows
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
  hops.forEach(function(consumer) {
    var cid = String(consumer['ID'] || consumer.hopId);
    var cFrom = String(consumer['From Account'] || consumer.fromAccount || '');
    hops.forEach(function(producer) {
      var pid = String(producer['ID'] || producer.hopId);
      if (pid === cid) return;
      var pTo = String(producer['To Account'] || producer.toAccount || '');
      if (pTo && pTo === cFrom) {
        deps[cid][pid] = true;
        reverseDeps[pid][cid] = true;
      }
    });
  });

  var pending = Object.keys(byId);
  var result = [];
  while (pending.length) {
    var ready = pending.filter(function(id) { return Object.keys(deps[id]).length === 0; });
    if (!ready.length) {
      // Cycle or unresolvable dep — emit remaining in original order.
      pending.sort(function(a, b) { return initialOrder[a] - initialOrder[b]; });
      pending.forEach(function(id) { result.push(byId[id]); });
      break;
    }
    // Tiebreak: percentage ascending (small branches first) then insertion order.
    ready.sort(function(a, b) {
      var pa = Number(byId[a]['Amount % of Total'] || byId[a].pctOfTotal) || 0;
      var pb = Number(byId[b]['Amount % of Total'] || byId[b].pctOfTotal) || 0;
      if (pa !== pb) return pa - pb;
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
  var lines = ['The money is in. I' + String.fromCharCode(8217) + 'll move it as follows', ''];
  ordered.forEach(function(h, i) {
    var pct = Number(h['Amount % of Total']) || 0;
    var amt = Math.round(total * pct / 100 * 100) / 100;
    lines.push((i + 1) + '. ' + h['From Account'] + ' to ' + h['To Account'] + ': ' + _fmtUsdAmount_(amt));
  });
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
  var lines = ['The money is in. I' + String.fromCharCode(8217) + 'll move it as follows', ''];
  ordered.forEach(function(w, i) {
    lines.push((i + 1) + '. ' + w['From Account'] + ' to ' + w['To Account'] + ': ' + _fmtUsdAmount_(w._amount));
  });
  return {
    movement:  d.movement,
    template:  d.template,
    wireCount: ordered.length,
    message:   lines.join('\n')
  };
}

// Menu-callable — prompts for a template name substring + a total amount,
// then shows the generated message so Amanda can eyeball the format before
// the tab UI is built. Useful for testing the topo-sort logic on real data.
function menuPreviewMovementMessage() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('Preview Movement Message', 'Template name (substring, e.g. "NF PR SJ" or "Tiger Down"):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var query = String(r1.getResponseText() || '').trim().toLowerCase();
  if (!query) return;
  var templates = getMovementTemplates();
  var match = templates.find(function(t) { return String(t['Name']).toLowerCase().indexOf(query) >= 0; });
  if (!match) {
    ui.alert('No template found matching "' + query + '".\n\nAvailable:\n' + templates.map(function(t){return '  • ' + t['Name'];}).join('\n'));
    return;
  }
  var r2 = ui.prompt('Preview Movement Message', 'Total amount for "' + match['Name'] + '":', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var amt = Number(String(r2.getResponseText() || '').replace(/[$,\s]/g, '')) || 0;
  if (amt <= 0) { ui.alert('Amount must be > 0.'); return; }
  var preview = formatMovementPreview(match['ID'], amt);
  if (preview.error) { ui.alert(preview.error); return; }
  ui.alert(match['Name'] + ' — ' + _fmtUsdAmount_(amt), preview.message, ui.ButtonSet.OK);
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
  'Michael Nguyen Personal':                'Michael Nguyen Personal (1319)'
};

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


// Menu-callable — refresh sibling wire percentages on every existing hop
// where a David / Michelle / Nancy personal account appears as From or To.
// Sets Michelle + Nancy to 0.333333% and David to 0.333334% so the sum still
// equals 1% but $100k movements produce $333.33 wires (not $333.00 with a
// truncated 0.333%). Idempotent; runs once after Amanda seeds and any time
// pct precision needs correction later.
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

  var SIBLING_RE = /\b(david|michelle|nancy)\b[\s\w()]*personal/i;
  var changed = 0;
  var summary = [];
  for (var r = 1; r < data.length; r++) {
    var from = String(data[r][iFrom] || '');
    var to   = String(data[r][iTo] || '');
    var chain = String(data[r][iChain] || '');
    // Only touch rows involving a sibling personal account.
    if (!SIBLING_RE.test(from) && !SIBLING_RE.test(to)) continue;
    var isDavid = /\bdavid\b/i.test(from) || /\bdavid\b/i.test(to);
    var newPct = isDavid ? 0.333334 : 0.333333;
    var oldPct = Number(data[r][iPct]) || 0;
    if (Math.abs(oldPct - newPct) > 0.0000001) {
      sheet.getRange(r + 1, iPct + 1).setValue(newPct);
      changed++;
      summary.push('  ' + chain + ' (' + (isDavid ? 'David' : 'Michelle/Nancy') + '): ' + oldPct + ' → ' + newPct);
    }
  }
  var msg = 'Updated ' + changed + ' sibling wire percentage(s).\n\n' + (summary.length ? summary.join('\n') : '(already precise)');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Sibling Percentages Fixed', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
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
