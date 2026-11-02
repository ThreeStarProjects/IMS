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

  var msg = '"' + IM_RSV_QUEUE + '" sheet ready; requests are processed every minute.';
  try { SpreadsheetApp.getUi().alert('Reservation queue', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }  // run from editor: no UI context
}

function IM_processReservationRequests() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;  // a previous run is still going
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(IM_RSV_QUEUE);
    if (!sh || sh.getLastRow() < 2) return;
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, IM_RSV_HEADERS.length).getValues();

    rows.forEach(function (r, i) {
      if (r[7] !== 'Pending') return;
      var res;
      try {
        res = addReservation({
          date:   new Date(),
          code:   extractCodePrefix(r[4]),
          qty:    r[5],
          forRef: String(r[3]) + (r[6] ? ' / ' + r[6] : ''),
          notes:  'Requested from PM by ' + r[2]
        });
      } catch (e) { res = { success: false, message: e.message }; }

      // H Status, I Reservation No, J Processed At, K Message
      sh.getRange(i + 2, 8, 1, 4).setValues([[
        res.success ? 'Done' : 'Error', res.success ? res.code : '', new Date(), res.message || ''
      ]]);
    });
  } finally { lock.releaseLock(); }
}
