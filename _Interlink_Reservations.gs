/**
 * _Interlink_Reservations.gs  (v0.4 Phase 4/6, IMS side)
 * -----------------------------------------------------------------------------
 * ONE sheet for all reservations: 'Reservation Register'. PM appends a row with
 * Status "Requested" (no Reservation No yet); a 1-minute time trigger turns it into a
 * normal Active reservation in IMS's own context (counter, document lock, snapshot
 * refresh all behave as usual). An unusable request becomes Status "Error" with the
 * reason in Notes.
 *
 * Register columns used here (A:N are the original ones):
 *   A Sr | B Reservation No | C Date | D Code | E Description | F Qty Reserved | G UOM |
 *   H For (Reference) | I Status | J Notes | K Created By | L Created At |
 *   M Released/Consumed At | N Released By | O Consumed Qty | P Request ID (PM's PMR-xxxx)
 *
 * Status: Requested → Active → Partially Consumed → Consumed, or Released / Error.
 * Active and Partially Consumed HOLD stock: held = Qty Reserved − Consumed Qty.
 *
 * One-time: run IM_installReservationRegister() from the editor (columns, dropdown, colours,
 * FG Master formula, trigger; moves anything left in the old 'PM Reservation Requests' sheet).
 */

var IM_REG = 'Reservation Register';
var IM_REG_WIDTH = 16;
var IM_OLD_QUEUE = 'PM Reservation Requests';

/** Active and Partially Consumed reservations hold stock. */
function IM_isOpenRsv_(status) { return status === 'Active' || status === 'Partially Consumed'; }

function IM_round3_(n) { return Math.round(n * 1000) / 1000; }

function IM_installReservationRegister() {
  var ss = SpreadsheetApp.getActive();
  var reg = ss.getSheetByName(IM_REG);
  if (!reg) { IM_sayRsv_('Reservation Register sheet not found.'); return; }

  // 1. columns O, P + Status dropdown + colours
  if (reg.getMaxColumns() < IM_REG_WIDTH + 2) reg.insertColumnsAfter(reg.getMaxColumns(), IM_REG_WIDTH + 2 - reg.getMaxColumns());
  reg.getRange(1, 15, 1, 2).setValues([['Consumed Qty', 'Request ID']]);
  reg.getRange(1, 14).copyTo(reg.getRange(1, 15, 1, 2), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  reg.setColumnWidth(15, 100); reg.setColumnWidth(16, 110);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Requested', 'Active', 'Partially Consumed', 'Consumed', 'Released', 'Error'], true).setAllowInvalid(true).build();
  reg.getRange(2, 9, Math.max(reg.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  if (typeof _cfReservations === 'function') _cfReservations(ss);

  // 2. FG Master Available must subtract only what is still held
  var fg = ss.getSheetByName('FG Master');
  if (fg && ON_HAND_FORMULAS['FG Master']) {
    var spec = ON_HAND_FORMULAS['FG Master'];
    var cell = fg.getRange(2, spec.col);
    if (cell.getFormula().indexOf('Reservation Register') >= 0) cell.setFormula(spec.formula);
  }

  // 3. fold the old queue sheet into the Register, then retire it
  var moved = IM_migrateOldQueue_(ss, reg);

  // 4. processing trigger
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'IM_processReservationRequests') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('IM_processReservationRequests').timeBased().everyMinutes(1).create();
  var res = IM_processReservationRequests();
  IM_sayRsv_('Reservation Register ready (columns O:P, status list, colours, trigger every minute).\n' + moved +
             '\nProcessed now: ' + res.done + ' reserved, ' + res.failed + ' failed.');
}

/**
 * Old 'PM Reservation Requests' → Register. Pending requests become Requested rows; a Done request
 * hands its Request ID to the reservation it created (so PM rows still pointing at PMR-xxxx resolve).
 * The old sheet is then renamed and hidden, not deleted: delete it yourself once PM looks right.
 */
function IM_migrateOldQueue_(ss, reg) {
  var old = ss.getSheetByName(IM_OLD_QUEUE);
  if (!old) return 'No old queue sheet to move.';
  var rows = old.getLastRow() >= 2 ? old.getRange(2, 1, old.getLastRow() - 1, 11).getValues() : [];
  var rv = reg.getLastRow() >= 2 ? reg.getRange(2, 1, reg.getLastRow() - 1, IM_REG_WIDTH).getValues() : [];
  var byRes = {}; rv.forEach(function (r, i) { if (r[1]) byRes[r[1]] = i + 2; });
  var haveReq = {}; rv.forEach(function (r) { if (r[15]) haveReq[r[15]] = true; });
  var tagged = 0, requeued = 0;
  rows.forEach(function (q) {  // A Req ID, B At, C By, D Project, E Item, F Qty, G Source Ref, H Status, I Res No
    if (!q[0] || haveReq[q[0]]) return;
    if (q[7] === 'Done' && byRes[q[8]]) { reg.getRange(byRes[q[8]], 16).setValue(q[0]); tagged++; }
    else if (q[7] === 'Pending') {
      IM_appendRequestRows_(reg, [{ reqId: q[0], by: q[2], code: q[4], qty: q[5], forRef: String(q[3]) + (q[6] ? ' / ' + q[6] : ''), at: q[1] }]);
      requeued++;
    }
  });
  old.setName('_old PM Reservation Requests'); old.hideSheet();
  return 'Old queue: ' + tagged + ' finished request(s) linked, ' + requeued + ' pending moved. Sheet renamed "_old PM Reservation Requests" and hidden (delete it when you are happy).';
}

/** Append "Requested" rows to the Register. Used by the migration; PM does the same from its side. */
function IM_appendRequestRows_(reg, reqs) {
  var start = nextRowByCol(reg, 1);
  var out = reqs.map(function (r, i) {
    var row = new Array(IM_REG_WIDTH); for (var k = 0; k < IM_REG_WIDTH; k++) row[k] = '';
    row[0] = start + i - 1; row[2] = r.at || new Date(); row[3] = r.code; row[5] = r.qty; row[7] = r.forRef;
    row[8] = 'Requested'; row[10] = r.by; row[11] = r.at || new Date(); row[15] = r.reqId;
    return row;
  });
  if (out.length) reg.getRange(start, 1, out.length, IM_REG_WIDTH).setValues(out);
}

function IM_sayRsv_(msg) {
  try { SpreadsheetApp.getUi().alert('Reservations', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }  // run from editor: no UI context
}

/** Recreate the 1-minute trigger if it has gone missing. Called when IMS opens. */
function IM_ensureQueueTrigger_() {
  var has = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'IM_processReservationRequests';
  });
  var reg = SpreadsheetApp.getActive().getSheetByName(IM_REG);
  if (!has && reg && reg.getRange(1, 16).getValue() === 'Request ID') {
    ScriptApp.newTrigger('IM_processReservationRequests').timeBased().everyMinutes(1).create();
    Logger.log('queue trigger was missing; recreated');
  }
}

/**
 * Consume `q` of one reservation row (vals = its A:O values). Consumed Qty += q; the status becomes
 * Consumed when nothing is left, otherwise Partially Consumed. Returns the new status.
 */
function IM_consumeRsvRow_(reg, rowNum, vals, q, note) {
  var total = Number(vals[5]) || 0, done = IM_round3_((Number(vals[14]) || 0) + q);
  var full = done >= total - 0.0005;
  if (full) done = total;
  var st = full ? 'Consumed' : 'Partially Consumed';
  reg.getRange(rowNum, 9).setValue(st);                        // I Status
  reg.getRange(rowNum, 13).setValue(new Date());               // M Released/Consumed At
  reg.getRange(rowNum, 14).setValue(currentUserEmail());       // N Released By
  reg.getRange(rowNum, 15).setValue(done);                     // O Consumed Qty
  if (note) reg.getRange(rowNum, 10).setValue((vals[9] ? vals[9] + ' | ' : '') + note);  // J Notes
  return st;
}

/**
 * Called by postRequisitionRows (20_Registers.gs) right after the requisition rows are written.
 * A Material Requisition is the outward event, so it consumes the matching open PM reservations
 * instead of someone marking them Consumed by hand. A reservation matches when:
 *   - it was made from PM: For (Reference) = "<project> / HW-nnn",
 *   - its project matches the requisition's by PRJ-### id (or by exact name when there is no id),
 *   - the item code matches.
 * Oldest first, each takes min(still held, qty issued and not yet used): issuing less than the
 * whole reservation makes it "Partially Consumed" and the rest stays reserved.
 * Runs before the snapshot refresh in postRequisitionRows, so one refresh covers both.
 * Never throws: a requisition must always post. Returns how many reservations it touched.
 */
function IM_consumePmReservations_(project, lines, reqNo) {
  try {
    var reg = SpreadsheetApp.getActive().getSheetByName(IM_REG);
    if (!reg || reg.getLastRow() < 2 || !project) return 0;
    var key = IM_projKey_(project);
    var issued = {};  // item code → qty issued in this requisition, not yet matched
    lines.forEach(function (l) { issued[l.itemCode] = (issued[l.itemCode] || 0) + Number(l.qty); });

    var data = reg.getRange(2, 1, reg.getLastRow() - 1, 15).getValues();
    var touched = 0;
    var PM_REF = / \/ HW-\d+$/;
    data.forEach(function (r, i) {  // Register rows are in reservation-number order, so oldest first
      var code = r[3], ref = String(r[7]);
      if (!IM_isOpenRsv_(r[8]) || !PM_REF.test(ref) || !(issued[code] > 0)) return;
      if (IM_projKey_(ref.replace(PM_REF, '')) !== key) return;
      var left = IM_round3_((Number(r[5]) || 0) - (Number(r[14]) || 0));
      var take = IM_round3_(Math.min(left, issued[code]));
      if (!(take > 0)) return;
      issued[code] = IM_round3_(issued[code] - take);
      IM_consumeRsvRow_(reg, i + 2, r, take, 'Consumed ' + take + ' by ' + reqNo);
      touched++;
    });
    return touched;
  } catch (e) {
    Logger.log('IM_consumePmReservations_: ' + e.message);
    return 0;
  }
}

/** Project comparison key: the PRJ-### id when there is one, else the trimmed lower-case name. */
function IM_projKey_(s) {
  var m = /PRJ-\d+/i.exec(String(s));
  return m ? m[0].toUpperCase() : String(s).trim().toLowerCase();
}

/**
 * A project delete (Masters.executeProjectDelete, started from PM) removes requisition / reservation rows
 * from here and lists their item codes in the hidden sheet _PD_Refresh (column A). Refresh those items'
 * stock snapshots in IMS's own context, then clear the list.
 */
function IM_refreshPendingSnapshots_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('_PD_Refresh');
  if (!sh || sh.getLastRow() < 1) return 0;
  var codes = {};
  sh.getRange(1, 1, sh.getLastRow(), 1).getValues().forEach(function (r) { if (r[0]) codes[String(r[0])] = true; });
  var list = Object.keys(codes), rm = [], sfg = [];
  var items = ss.getSheetByName('Item Master'), sfgs = ss.getSheetByName('SFG Master');
  list.forEach(function (c) {
    if (items && getRowByCode(items, c)) rm.push(c); else if (sfgs && getRowByCode(sfgs, c)) sfg.push(c);
  });
  if (rm.length) updateItemSnapshotsBatch(rm);
  if (sfg.length) updateSfgSnapshotsBatch(sfg);
  sh.clearContents();
  return list.length;
}

/** Menu action: process waiting PM requests right now. */
function IM_processReservationRequestsNow() {
  var r = IM_processReservationRequests();
  SpreadsheetApp.getActive().toast(r.done + ' reserved, ' + r.failed + ' failed', 'PM reservation requests', 6);
}

/**
 * Turn every "Requested" Register row into an Active reservation. Returns { done, failed }.
 * No duplicates: if an open reservation already exists for the same item + reference, the request
 * row is dropped and its Request ID is handed to the existing reservation (PM then finds it).
 */
function IM_processReservationRequests() {
  var out = { done: 0, failed: 0 };
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) { Logger.log('queue: busy, next run will do it'); return out; }
  try {
    var ss = SpreadsheetApp.getActive();
    var reg = ss.getSheetByName(IM_REG);
    if (!reg) return out;
    // Heartbeat (Q1): if this timestamp is stale, the time trigger is not running.
    var n = ScriptApp.getProjectTriggers().filter(function (t) {
      return t.getHandlerFunction() === 'IM_processReservationRequests';
    }).length;
    reg.getRange('Q1').setValue('Heartbeat: last run ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm:ss') + ' · triggers: ' + n);
    try { IM_refreshPendingSnapshots_(); } catch (e) { Logger.log('pending snapshots: ' + e.message); }
    if (reg.getLastRow() < 2) return out;
    var rows = reg.getRange(2, 1, reg.getLastRow() - 1, IM_REG_WIDTH).getValues();

    var open = {};  // "code|forRef" → row number of an open reservation
    rows.forEach(function (x, i) { if (IM_isOpenRsv_(x[8])) open[x[3] + '|' + x[7]] = i + 2; });

    var drop = [], rmCodes = [], sfgCodes = [];
    rows.forEach(function (r, i) {
      if (r[8] !== 'Requested') return;
      var row = i + 2, code = extractCodePrefix(r[3]), forRef = String(r[7]), qty = Number(r[5]);
      var fail = function (msg) { reg.getRange(row, 9).setValue('Error'); reg.getRange(row, 10).setValue(msg); out.failed++; };
      try {
        if (open[code + '|' + forRef]) {  // duplicate request
          var ex = open[code + '|' + forRef];
          if (r[15] && !reg.getRange(ex, 16).getValue()) reg.getRange(ex, 16).setValue(r[15]);
          drop.push(row); out.done++; return;
        }
        if (!(qty > 0)) return fail('Qty must be > 0');
        var master = ss.getSheetByName('Item Master'), mrow = getRowByCode(master, code), kind = 'RM';
        if (!mrow) { master = ss.getSheetByName('SFG Master'); mrow = getRowByCode(master, code); if (mrow) kind = 'SFG'; }
        if (!mrow) { master = ss.getSheetByName('FG Master'); mrow = getRowByCode(master, code); if (mrow) kind = 'FG'; }
        if (!mrow) return fail('Code not found: ' + code);
        var resNo = getNextCounter('Next Reservation Counter', 'RSV', 4);
        reg.getRange(row, 2).setValue(resNo);                                  // B Reservation No
        reg.getRange(row, 3).setValue(new Date());                             // C Date
        reg.getRange(row, 4).setValue(code);                                   // D Code (label → code)
        reg.getRange(row, 5).setValue(mrow[2] || '');                          // E Description
        reg.getRange(row, 7).setValue(kind === 'FG' ? (mrow[7] || '') : (mrow[6] || ''));  // G UOM
        reg.getRange(row, 9).setValue('Active');                               // I Status
        reg.getRange(row, 10).setValue('Requested from PM by ' + r[10]);       // J Notes
        open[code + '|' + forRef] = row;
        if (kind === 'RM') rmCodes.push(code); else if (kind === 'SFG') sfgCodes.push(code);
        out.done++;
      } catch (e) { fail(e.message); }
    });
    drop.sort(function (a, b) { return b - a; }).forEach(function (row) { reg.deleteRow(row); });
    if (rmCodes.length)  updateItemSnapshotsBatch(rmCodes);
    if (sfgCodes.length) updateSfgSnapshotsBatch(sfgCodes);
    if (out.done || out.failed) Logger.log('queue: ' + out.done + ' reserved, ' + out.failed + ' failed');
    return out;
  } finally { lock.releaseLock(); }
}
