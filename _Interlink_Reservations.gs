/**
 * _Interlink_Reservations.gs  (v0.4 Phase 4, IMS side)
 * -----------------------------------------------------------------------------
 * PM → IMS reservation queue. PM appends rows to 'PM Reservation Requests'; a
 * 1-minute time trigger runs the real addReservation() here, in IMS's own
 * context (counter, document lock, snapshot refresh all behave as usual), and
 * writes the RSV number / error back onto the request row. No existing IMS
 * code is modified.
 *
 * Queue columns:
 *   A Request ID | B Requested At | C Requested By | D Project Label |
 *   E Item (code or "CODE — desc" label) | F Qty | G Source Ref (e.g. HW-001) |
 *   H Status (Pending / Done / Error) | I Reservation No | J Processed At | K Message
 *
 * One-time: run IM_installReservationQueue() from the editor.
 */

var IM_RSV_QUEUE = 'PM Reservation Requests';
var IM_RSV_HEADERS = ['Request ID', 'Requested At', 'Requested By', 'Project Label', 'Item',
                      'Qty', 'Source Ref', 'Status', 'Reservation No', 'Processed At', 'Message'];

function IM_installReservationQueue() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(IM_RSV_QUEUE) || ss.insertSheet(IM_RSV_QUEUE);
  sh.getRange(1, 1, 1, IM_RSV_HEADERS.length).setValues([IM_RSV_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange('J:J').setNumberFormat('dd/mm/yyyy hh:mm');
  sh.setColumnWidth(5, 260); sh.setColumnWidth(11, 300);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'IM_processReservationRequests') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('IM_processReservationRequests').timeBased().everyMinutes(1).create();
  sh.getRange('M1').setValue('Heartbeat: waiting for first run');
  sh.setColumnWidth(13, 320);

  // Process anything already waiting, so this one run also clears a stuck request.
  var res = IM_processReservationRequests();
  var msg = '"' + IM_RSV_QUEUE + '" ready; trigger recreated (every minute). ' +
            'Processed now: ' + res.done + ' reserved, ' + res.failed + ' failed.' +
            (res.failed ? ' See the Message column on the queue sheet.' : '');
  try { SpreadsheetApp.getUi().alert('Reservation queue', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }  // run from editor: no UI context
}

/** Recreate the 1-minute trigger if it has gone missing. Called when IMS opens. */
function IM_ensureQueueTrigger_() {
  var has = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'IM_processReservationRequests';
  });
  if (!has && SpreadsheetApp.getActive().getSheetByName(IM_RSV_QUEUE)) {
    ScriptApp.newTrigger('IM_processReservationRequests').timeBased().everyMinutes(1).create();
    Logger.log('queue trigger was missing; recreated');
  }
}

/**
 * Called by postRequisitionRows (20_Registers.gs) right after the requisition rows are written.
 * A Material Requisition is the outward event, so it consumes the matching Active PM reservations
 * instead of someone marking them Consumed by hand. A reservation matches when:
 *   - it was made from PM: For (Reference) = "<project> / HW-nnn",
 *   - its project matches the requisition's by PRJ-### id (or by exact name when there is no id),
 *   - the item code matches, and
 *   - the issued qty of that item in this requisition covers the WHOLE reservation (oldest first).
 * A smaller issue leaves the reservation Active (it would otherwise free stock that is still
 * reserved). Runs before the snapshot refresh in postRequisitionRows, so one refresh covers both.
 * Never throws: a requisition must always post. Returns how many reservations it consumed.
 */
function IM_consumePmReservations_(project, lines, reqNo) {
  try {
    var reg = SpreadsheetApp.getActive().getSheetByName('Reservation Register');
    if (!reg || reg.getLastRow() < 2 || !project) return 0;
    var key = IM_projKey_(project);
    var issued = {};  // item code → qty issued in this requisition
    lines.forEach(function (l) { issued[l.itemCode] = (issued[l.itemCode] || 0) + Number(l.qty); });

    var data = reg.getRange(2, 1, reg.getLastRow() - 1, 14).getValues();
    var user = currentUserEmail(), now = new Date(), done = 0;
    var PM_REF = / \/ HW-\d+$/;
    data.forEach(function (r, i) {  // Register rows are in reservation-number order, so oldest first
      var code = r[3], ref = String(r[7]), qty = Number(r[5]);
      if (r[8] !== 'Active' || !PM_REF.test(ref) || !(qty > 0) || !(issued[code] >= qty)) return;
      if (IM_projKey_(ref.replace(PM_REF, '')) !== key) return;
      issued[code] -= qty;
      var row = i + 2;
      reg.getRange(row, 9).setValue('Consumed');   // I Status
      reg.getRange(row, 13).setValue(now);         // M Released/Consumed At
      reg.getRange(row, 14).setValue(user);        // N Released By
      reg.getRange(row, 10).setValue((r[9] ? r[9] + ' | ' : '') + 'Auto-consumed by ' + reqNo);  // J Notes
      done++;
    });
    return done;
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

/** Menu action: process waiting PM requests right now. */
function IM_processReservationRequestsNow() {
  var r = IM_processReservationRequests();
  SpreadsheetApp.getActive().toast(r.done + ' reserved, ' + r.failed + ' failed', 'PM reservation requests', 6);
}

/** Process every Pending request. Returns { done, failed } counts. */
function IM_processReservationRequests() {
  var out = { done: 0, failed: 0 };
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log('queue: previous run still going'); return out; }
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(IM_RSV_QUEUE);
    if (!sh) return out;
    // Heartbeat (M1): if this timestamp is stale, the time trigger is not running.
    var n = ScriptApp.getProjectTriggers().filter(function (t) {
      return t.getHandlerFunction() === 'IM_processReservationRequests';
    }).length;
    sh.getRange('M1').setValue('Heartbeat: last run ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm:ss') + ' · triggers: ' + n);
    if (sh.getLastRow() < 2) return out;
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, IM_RSV_HEADERS.length).getValues();

    // No duplicates: an Active reservation for the same item + reference already counts as
    // done. Keyed "code|forRef" → Reservation No; updated as this run adds more.
    var active = {};
    var reg = SpreadsheetApp.getActive().getSheetByName('Reservation Register');
    if (reg && reg.getLastRow() >= 2) {
      reg.getRange(2, 1, reg.getLastRow() - 1, 9).getValues().forEach(function (x) {
        if (x[8] === 'Active') active[x[3] + '|' + x[7]] = x[1];  // D code, H For (Reference), I Status, B Res No
      });
    }

    rows.forEach(function (r, i) {
      if (r[7] !== 'Pending') return;
      var code = extractCodePrefix(r[4]);
      var forRef = String(r[3]) + (r[6] ? ' / ' + r[6] : '');
      var res;
      if (active[code + '|' + forRef]) {
        res = { success: true, code: active[code + '|' + forRef], message: 'Already reserved — no duplicate created' };
      } else {
        try {
          res = addReservation({
            date:   new Date(),
            code:   code,
            qty:    r[5],
            forRef: forRef,
            notes:  'Requested from PM by ' + r[2]
          });
          if (res.success) active[code + '|' + forRef] = res.code;
        } catch (e) { res = { success: false, message: e.message }; }
      }

      // H Status, I Reservation No, J Processed At, K Message
      sh.getRange(i + 2, 8, 1, 4).setValues([[
        res.success ? 'Done' : 'Error', res.success ? res.code : '', new Date(), res.message || ''
      ]]);
      if (res.success) out.done++; else out.failed++;
    });
    if (out.done || out.failed) Logger.log('queue: ' + out.done + ' reserved, ' + out.failed + ' failed');
    return out;
  } finally { lock.releaseLock(); }
}
