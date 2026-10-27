/**
 * 3 Star Projects IMS v3 — 60_Utils_Audit_Triggers.gs
 * Utility helpers, audit + error logging, counter management, trigger stubs.
 * Used by every other .gs file.
 *
 * Public helpers:
 *   getNextCounter(counterKey, prefix, pad)     → 'GRN-0001'
 *   getRowByCode(sheet, code)                   → row array (0-indexed) or null
 *   getRowsByValue(sheet, colNum, value)        → array of 1-indexed row numbers
 *   toast(msg, title, seconds)
 *   dateString(date)                            → 'YYYY-MM-DD'
 */

// ---------------------------------------------------------------------------
// Counters (Config-backed) — reservation is atomic per call via LockService
// ---------------------------------------------------------------------------

function getNextCounter(counterKey, prefix, pad) {
  // Callers (post* in 20_) already hold LockService.getDocumentLock().
  // Nested lock release would end outer lock early — no inner lock here.
  const current = Number(getConfig(counterKey)) || 0;
  const next = current + 1;
  setConfig(counterKey, next);
  if (!prefix) return next;
  return prefix + '-' + String(next).padStart(pad || 4, '0');
}

// ---------------------------------------------------------------------------
// Sheet lookup helpers — used by every form handler
// ---------------------------------------------------------------------------

function getRowByCode(sheet, code) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) if (data[i][0] === code) return data[i];
  return null;
}

function getRowsByValue(sheet, colNum, value) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const col = sheet.getRange(2, colNum, sheet.getLastRow() - 1, 1).getValues();
  const rows = [];
  for (let i = 0; i < col.length; i++) if (col[i][0] === value) rows.push(i + 2);
  return rows;
}

// Return next writable row based on the CODE column (col A by default).
// Bypasses sh.getLastRow() which is inflated by ARRAYFORMULA spills that
// output '' — causing appendRow / getLastRow+1 to land far below real data.
function nextRowByCol(sh, col) {
  col = col || 1;
  if (!sh) return 2;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 2;
  const vals = sh.getRange(1, col, lastRow, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '' && vals[i][0] != null) return i + 2;
  }
  return 2;
}


// ---------------------------------------------------------------------------
// User-facing helpers
// ---------------------------------------------------------------------------

function toast(msg, title, seconds) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, title || '3 Star IMS', seconds || 5);
}

function dateString(d) {
  const dt = (d instanceof Date) ? d : new Date(d);
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function currentUserEmail() {
  return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'unknown';
}

// Extract "CODE" from a composite "CODE — Desc | Attr | Attr" dropdown value.
// Idempotent: no separator → returns input trimmed. Used by every server
// handler that receives a dropdown pick from a Search Label source.
function extractCodePrefix(v) {
  if (!v) return '';
  const s = String(v).trim();
  const idx = s.indexOf(' — ');
  return idx > 0 ? s.substring(0, idx).trim() : s;
}

// Read col B (Photo) CellImage from a master sheet, batched.
// Returns { code: photoUrl } for all rows. Uses CellImage.getUrl() (original
// source URL); falls back to getContentUrl() (Google-hosted, may need auth).
function getMasterPhotoUrlMap(sheetName) {
  const out = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return out;
  const range = sh.getRange(2, 1, sh.getLastRow() - 1, 2);
  const vals = range.getValues();
  for (let i = 0; i < vals.length; i++) {
    const code = vals[i][0];
    const photo = vals[i][1];
    if (!code || !photo || typeof photo !== 'object') continue;
    try {
      if (typeof photo.getUrl === 'function') {
        const u = photo.getUrl();
        if (u) { out[code] = u; continue; }
      }
      if (typeof photo.getContentUrl === 'function') {
        const cu = photo.getContentUrl();
        if (cu) out[code] = cu;
      }
    } catch (e) { /* skip */ }
  }
  return out;
}

// Parse a date input (Date object or 'YYYY-MM-DD' string). Throws on invalid.
function parseDate(d) {
  if (d instanceof Date) return d;
  if (typeof d === 'string' && d) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    const parsed = new Date(d);
    if (!isNaN(parsed)) return parsed;
  }
  throw new Error('Invalid date: ' + d);
}





// ---------------------------------------------------------------------------
// Triggers — install / uninstall the nightly job.
// onEdit is a simple trigger (auto-registered by function name), no install
// needed. Only the time-driven nightlyRebuild needs setupTriggers().
// ---------------------------------------------------------------------------

function setupTriggers() {
  // Remove any existing time-driven triggers we own, then reinstall.
  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(function(t) {
    if (t.getHandlerFunction() === 'nightlyRebuild') ScriptApp.deleteTrigger(t);
  });
  if (getConfig('Nightly Rebuild Enabled') === false) {
    toast('Nightly rebuild disabled in Config. No trigger installed.');
    return;
  }
  const hour = Number(getConfig('Nightly Rebuild Hour (IST)'));
  const safeHour = (hour >= 0 && hour < 24) ? hour : 2;
  ScriptApp.newTrigger('nightlyRebuild').timeBased().atHour(safeHour).everyDays(1).create();
  toast('Nightly rebuild installed for ' + safeHour + ':00 (script timezone).');
}

function nightlyRebuild() {
  // Everything wrapped so a viewer failure doesn't kill the whole run.
  const results = { fullRebuild: false, viewers: 0, lowStock: 0, emailed: false,
                    errors: [], startTime: new Date(), endTime: null };
  function step(name, fn) {
    try { fn(); results[name] = true; }
    catch (e) {
      results.errors.push({ step: name, error: e.message });
      Logger.log('NIGHTLY REBUILD FAILED AT ' + name + ': ' + e.message);
    }
  }
  step('fullRebuild', function() { fullRebuild(); });
  step('dashboard', function() { rebuildDashboard(); results.viewers++; });
  step('explorer', function() {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Explorer');
    if (sh) {
      const view = sh.getRange('B1').getValue();
      const arg  = sh.getRange('D1').getValue();
      if (view) { rebuildExplorer(view, arg); results.viewers++; }
    }
  });
  // Low-stock email (best-effort)
  step('emailLowStock', function() { results.emailed = emailLowStockAlert(); });

  results.endTime = new Date();

  // Alert admin on ANY error so silent failures surface
  if (results.errors.length > 0) {
    _sendRebuildFailureEmail(results);
  }

  return results;
}

function _sendRebuildFailureEmail(results) {
  const admin = String(getConfig('Admin Email') || '').trim();
  if (!admin) return;  // No admin configured — nothing to send
  const company = getConfig('Company Name') || '3 Star Projects';
  const errorList = results.errors.map(function(e) {
    return '  • ' + e.step + ': ' + e.error;
  }).join('\n');
  const duration = Math.round((results.endTime - results.startTime) / 1000);
  const body =
    'The IMS nightly rebuild completed with errors.\n\n' +
    'Started:  ' + results.startTime + '\n' +
    'Finished: ' + results.endTime + '\n' +
    'Duration: ' + duration + ' seconds\n\n' +
    'Errors (' + results.errors.length + '):\n' + errorList + '\n\n' +
    'Steps that succeeded:\n' +
    '  fullRebuild: ' + (results.fullRebuild ? 'OK' : 'FAILED') + '\n' +
    '  dashboard:   ' + (results.dashboard   ? 'OK' : 'FAILED') + '\n' +
    '  explorer:    ' + (results.explorer    ? 'OK' : 'FAILED') + '\n' +
    '  emailLowStock: ' + (results.emailLowStock ? 'OK' : 'FAILED') + '\n\n' +
    'Fix: open the sheet and run "Full Stock Rebuild" from the 3SP Setup menu.\n';
  try {
    MailApp.sendEmail({
      to: admin,
      subject: '[' + company + '] Nightly Rebuild FAILED — ' + results.errors.length + ' error(s)',
      body: body,
    });
  } catch (e) {
    Logger.log('Failed to send rebuild failure email: ' + e.message);
  }
}

// ---------------------------------------------------------------------------
// Low-stock email — sends summary of RM + SFG below MIN.
// Returns true if an email was sent, false if no recipients / no low items.
// ---------------------------------------------------------------------------

function emailLowStockAlert() {
  const csv = String(getConfig('Low Stock Alert Recipients') || '').trim();
  if (!csv) return false;
  const recipients = csv.split(',').map(function(s) { return s.trim(); }).filter(String);
  if (!recipients.length) return false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snap = ss.getSheetByName('Stock Snapshot');
  if (!snap || snap.getLastRow() < 2) return false;
  // Vertical layout: RM section A:P, SFG section A:I (below SFG header)
  const rmData  = _readRmSection(snap, 17);
  const sfgData = _readSfgSection(snap, 9);
  const rows = [];
  // RM (v3.6): A=code, B=desc, I=Available(8), K=MIN(10), Q=BelowMIN(16)
  rmData.forEach(function(r) {
    if (r[0] && r[16] === true) rows.push([r[0], 'RM', r[1], r[8], r[10], (r[10] - r[8]), 'Reorder']);
  });
  // SFG: A=code, B=desc, F=Available (idx 5), G=MIN (idx 6), I=Below MIN (idx 8)
  sfgData.forEach(function(r) {
    if (r[0] && r[8] === true) rows.push([r[0], 'SFG', r[1], r[5], r[6], (r[6] - r[5]), 'Produce / return from JW']);
  });
  if (!rows.length) return false;

  const company = getConfig('Company Name') || '3 Star Projects';
  const html = '<h2 style="color:#1F4E79;">Low-Stock Alert — ' + _htmlEscape(company) + '</h2>' +
    '<p>' + rows.length + ' item(s) below MIN LVL as of ' + dateString(new Date()) + ':</p>' +
    '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Arial;font-size:11pt;">' +
    '<tr style="background:#1F4E79;color:#fff;"><th>Code</th><th>Type</th><th>Description</th><th>On Hand</th><th>MIN</th><th>Shortfall</th><th>Action</th></tr>' +
    rows.map(function(r) {
      return '<tr>' + r.map(function(c) { return '<td>' + _htmlEscape(c) + '</td>'; }).join('') + '</tr>';
    }).join('') + '</table>' +
    '<p style="font-size:9pt;color:#666;">Sent automatically by 3 Star IMS. See the Low Stock Alerts tab for the live list.</p>';

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: '[' + company + '] Low-Stock Alert — ' + rows.length + ' item(s) below MIN',
    htmlBody: html,
  });
  return true;
}


// ---------------------------------------------------------------------------
// Regenerate missing PDFs — polish. Scans each register for rows whose
// PDF Link cell is empty and re-runs the appropriate generator. Handy when
// PDF gen was originally skipped (e.g. bad folder ID at the time).
// ---------------------------------------------------------------------------

function regenerateMissingPdfs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jobs = [
    { sheet: 'Inward Register',              docCol: 2, linkCol: 18, docType: 'GRN',           fn: _regenGrn },
    { sheet: 'Material Requisition Register', docCol: 2, linkCol: 12, docType: 'Requisition',   fn: _regenReq },
    { sheet: 'Job Work Outward Register',     docCol: 2, linkCol: 13, docType: 'JobWorkChallan', fn: _regenJwc },
    { sheet: 'Job Work Inward Register',      docCol: 2, linkCol: 13, docType: 'JobWorkGRN',    fn: _regenJwgrn },
    { sheet: 'Purchase Order Register',       docCol: 2, linkCol:  9, docType: 'PO',            fn: _regenPo },
  ];
  let total = 0, ok = 0;
  jobs.forEach(function(j) {
    const sh = ss.getSheetByName(j.sheet);
    if (!sh || sh.getLastRow() < 2) return;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    // Group rows by doc No
    const groups = {};
    data.forEach(function(r, i) {
      const docNo = r[j.docCol - 1];
      const link = r[j.linkCol - 1];
      if (!docNo || link) return;
      (groups[docNo] = groups[docNo] || []).push(r);
    });
    Object.keys(groups).forEach(function(docNo) {
      total++;
      try {
        const url = j.fn(docNo, groups[docNo]);
        writeBackPdfLink(sh, j.docCol, docNo, j.linkCol, url);
        ok++;
      } catch (e) {
        Logger.log('PDF regen failed for ' + docNo + ': ' + e.message);
      }
    });
  });
  const msg = 'Regenerated ' + ok + ' / ' + total + ' missing PDFs. Check View → Executions for any failure details.';
  toast(msg);
  return msg;
}

// Each _regen* takes a docNo + its already-committed register rows,
// rebuilds a synthetic formData and lineRows, and calls the generator.
function _regenGrn(grnNo, rows) {
  const first = rows[0];
  const formData = { date: first[2], supplierCode: first[5], supplierName: first[6], invoiceNo: first[15] };
  const lineRows = rows.map(function(r) {
    return { itemCode: r[3], itemDesc: r[4], qty: r[7], uom: r[8], rate: r[9], discPct: r[10], gstPct: r[12], total: r[14] };
  });
  return generateGrnPdf(grnNo, formData, lineRows).url;
}
function _regenReq(reqNo, rows) {
  const first = rows[0];
  const formData = { date: first[2], project: first[3], takenBy: first[10] };  // col K = Taken By
  const lineRows = rows.map(function(r) {
    return { itemCode: r[4], itemDesc: r[5], qty: r[6], uom: r[7], purpose: r[8] };
  });
  return generateReqSlipPdf(reqNo, formData, lineRows).url;
}
function _regenJwc(jwcNo, rows) {
  const first = rows[0];
  const formData = { date: first[2], jobWorkerCode: first[3], jobWorkerName: first[4], processType: first[5] };
  const lineRows = rows.map(function(r) {
    return { sfgCode: r[6], desc: r[7], qty: r[8], uom: r[9], expReturn: r[10], remarks: r[11],
             weight: r[16] || 0 };  // col Q = Weight (KG)
  });
  return generateJobWorkChallanPdf(jwcNo, formData, lineRows).url;
}
function _regenPo(poNo, rows) {
  // rows here = PO Register header rows (should be one per PO No).
  // We need to also read PO Lines for this PO.
  const first = rows[0];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const linesSh = ss.getSheetByName('PO Lines');
  const lineData = linesSh && linesSh.getLastRow() >= 2
    ? linesSh.getRange(2, 1, linesSh.getLastRow() - 1, 12).getValues().filter(function(r) { return r[0] === poNo; })
    : [];
  const formData = {
    date: first[2], supplierCode: first[3], supplierName: first[4],
    deliveryDate: first[5], remarks: first[9],
    supplierGstin: '', supplierAddress: '', // not stored on header; regenerated PDF won't have GSTIN
  };
  const lineRows = lineData.map(function(r) {
    return { itemCode: r[2], itemDesc: r[3], qty: r[4], uom: r[5], rate: r[6],
             discPct: r[7], gstPct: r[9], total: r[11] };
  });
  return generatePoPdf(poNo, formData, lineRows).url;
}

function _regenJwgrn(jwgrnNo, rows) {
  const first = rows[0];
  const formData = { date: first[2], referenceChallan: first[3], jobWorkerName: first[4] };
  const lineRows = rows.map(function(r) {
    return { sfgCode: r[5], desc: r[6], qtyRecv: r[7], uom: r[8], yieldLoss: r[9], rejection: r[10], remarks: r[11] };
  });
  return generateJobWorkGrnPdf(jwgrnNo, formData, lineRows).url;
}


// ---------------------------------------------------------------------------
// UNDO LAST TRANSACTION
// Deletes all rows of the most recent doc# in the chosen register + refreshes
// affected item/SFG snapshots. Multi-line docs (a GRN with 10 items shares one
// GRN No) get fully removed. PDF files in Drive are left orphaned — clean up
// manually if needed.
// ---------------------------------------------------------------------------

const UNDO_TARGETS = {
  'Inward Register':                { docCol: 2, itemCol: 4, kind: 'RM'  },
  'Material Requisition Register':  { docCol: 2, itemCol: 5, kind: 'RM'  },
  'Job Work Outward Register':      { docCol: 2, itemCol: 7, kind: 'SFG' },
  'Job Work Inward Register':       { docCol: 2, itemCol: 6, kind: 'SFG' },
  'Adjustments Register':           { docCol: 0, itemCol: 4, kind: 'MIX' },
  'Purchase Order Register':        { docCol: 2, itemCol: 0, kind: 'PO'  },
};

function undoLastTransaction() {
  const ui = SpreadsheetApp.getUi();
  const names = Object.keys(UNDO_TARGETS);
  const list = names.map(function(n, i) { return (i + 1) + '. ' + n; }).join('\n');
  const pick = ui.prompt('Undo Last Transaction',
    'Which register?\n\n' + list + '\n\nEnter number 1..' + names.length + ':',
    ui.ButtonSet.OK_CANCEL);
  if (pick.getSelectedButton() !== ui.Button.OK) return;
  const idx = parseInt(pick.getResponseText().trim(), 10) - 1;
  if (isNaN(idx) || idx < 0 || idx >= names.length) { ui.alert('Invalid selection.'); return; }
  const name = names[idx];
  const spec = UNDO_TARGETS[name];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) { ui.alert(name + ' not found.'); return; }
  const lastRow = nextRowByCol(sh, 1) - 1;
  if (lastRow < 2) { ui.alert('Register empty — nothing to undo.'); return; }

  // Doc# lives at spec.docCol (0 = use Sr No = row index). Read the last row's doc.
  let lastDoc, lastDocRows = [];
  if (spec.docCol > 0) {
    lastDoc = sh.getRange(lastRow, spec.docCol).getValue();
    // Find all rows sharing this doc# (may not be contiguous but usually are)
    const docs = sh.getRange(2, spec.docCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < docs.length; i++) {
      if (docs[i][0] === lastDoc) lastDocRows.push(i + 2);
    }
  } else {
    // No doc# — treat last single row as the unit
    lastDoc = '(last row)';
    lastDocRows = [lastRow];
  }

  // Read item codes for snapshot refresh before deletion
  const itemCodes = [];
  if (spec.itemCol > 0) {
    lastDocRows.forEach(function(r) {
      const c = sh.getRange(r, spec.itemCol).getValue();
      if (c) itemCodes.push(c);
    });
  }

  const confirm = ui.alert('Confirm Undo',
    'Delete ' + lastDocRows.length + ' row(s) with doc# "' + lastDoc + '" from ' + name + '?\n\n' +
    'Affected item codes: ' + (itemCodes.length ? Array.from(new Set(itemCodes)).join(', ') : '(none)') + '\n\n' +
    'This cannot be undone. PDF file (if any) will be orphaned in Drive.',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  // Delete rows bottom-up so indices stay valid
  lastDocRows.sort(function(a, b) { return b - a; }).forEach(function(r) { sh.deleteRow(r); });

  // Refresh snapshots for touched codes
  const unique = Array.from(new Set(itemCodes));
  unique.forEach(function(c) {
    if (spec.kind === 'SFG') { updateSfgSnapshot(c); }
    else if (spec.kind === 'RM') { updateItemSnapshot(c); }
    else if (spec.kind === 'MIX') {
      // Adjustments touches both — try both, ignore misses.
      try { updateItemSnapshot(c); } catch (e) {}
      try { updateSfgSnapshot(c);  } catch (e) {}
    }
  });

  ui.alert('Undone: ' + lastDocRows.length + ' row(s) removed from ' + name + '.\n\n' +
    'Refreshed snapshots for ' + unique.length + ' code(s).' +
    (spec.kind === 'PO' ? '\n\nNote: PO undo also leaves PO Lines orphaned — clean manually if needed.' : ''));
}


// ---------------------------------------------------------------------------
// DELETE A SINGLE REGISTER ROW
// Fixes "1 wrong line in a 10-line GRN" without undoing the whole transaction.
// Reads the affected item/SFG code before deletion, refreshes its snapshot after.
// ---------------------------------------------------------------------------

function deleteRegisterRow() {
  const ui = SpreadsheetApp.getUi();
  const names = Object.keys(UNDO_TARGETS);
  const list = names.map(function(n, i) { return (i + 1) + '. ' + n; }).join('\n');
  const pick = ui.prompt('Delete Register Row',
    'Which register?\n\n' + list + '\n\nEnter number 1..' + names.length + ':',
    ui.ButtonSet.OK_CANCEL);
  if (pick.getSelectedButton() !== ui.Button.OK) return;
  const idx = parseInt(pick.getResponseText().trim(), 10) - 1;
  if (isNaN(idx) || idx < 0 || idx >= names.length) { ui.alert('Invalid selection.'); return; }
  const name = names[idx];
  const spec = UNDO_TARGETS[name];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) { ui.alert(name + ' not found.'); return; }
  const lastRow = nextRowByCol(sh, 1) - 1;
  if (lastRow < 2) { ui.alert('Register empty — nothing to delete.'); return; }

  const rowPrompt = ui.prompt('Delete Register Row',
    name + ' — enter the SHEET ROW NUMBER to delete (row 1 is header, so data starts at row 2). Last data row is ' + lastRow + '.',
    ui.ButtonSet.OK_CANCEL);
  if (rowPrompt.getSelectedButton() !== ui.Button.OK) return;
  const rowNum = parseInt(rowPrompt.getResponseText().trim(), 10);
  if (isNaN(rowNum) || rowNum < 2 || rowNum > lastRow) { ui.alert('Row must be between 2 and ' + lastRow + '.'); return; }

  const rowVals = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];
  const docNo = spec.docCol > 0 ? rowVals[spec.docCol - 1] : '(no doc#)';
  const affectedCode = spec.itemCol > 0 ? rowVals[spec.itemCol - 1] : '';

  const confirm = ui.alert('Confirm Delete',
    'Delete row ' + rowNum + ' from ' + name + '?\n\n' +
    'Doc#: ' + docNo + '\n' +
    'Affected code: ' + (affectedCode || '(none)') + '\n\n' +
    'This cannot be undone. Snapshot will refresh automatically.',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  sh.deleteRow(rowNum);

  // Refresh snapshot for the affected code
  if (affectedCode) {
    if (spec.kind === 'SFG') updateSfgSnapshotsBatch([affectedCode]);
    else if (spec.kind === 'RM') updateItemSnapshotsBatch([affectedCode]);
    else if (spec.kind === 'MIX') {
      // Try both — one will be a no-op
      updateItemSnapshotsBatch([affectedCode]);
      updateSfgSnapshotsBatch([affectedCode]);
    }
  }

  ui.alert('Row ' + rowNum + ' deleted from ' + name + '.\n\n' +
    (affectedCode ? 'Snapshot refreshed for ' + affectedCode + '.' : '') +
    (spec.kind === 'PO' ? '\n\nNote: If this was a PO Register row, also delete matching PO Lines rows manually.' : ''));
}

// ---------------------------------------------------------------------------
// REGENERATE PDF FOR A SPECIFIC DOC
// Handy after Drive deletion or template changes. Overwrites any existing link.
// ---------------------------------------------------------------------------

const PDF_REGEN_TARGETS = {
  'Inward Register':                 { docCol: 2, linkCol: 18, fn: _regenGrn,  label: 'GRN' },
  'Material Requisition Register':   { docCol: 2, linkCol: 12, fn: _regenReq,  label: 'Requisition' },
  'Job Work Outward Register':       { docCol: 2, linkCol: 13, fn: _regenJwc,  label: 'Job Work Challan' },
  'Job Work Inward Register':        { docCol: 2, linkCol: 13, fn: _regenJwgrn,label: 'Job Work GRN' },
  'Purchase Order Register':         { docCol: 2, linkCol:  9, fn: _regenPo,   label: 'PO' },
};

function regeneratePdfForDoc() {
  const ui = SpreadsheetApp.getUi();
  const names = Object.keys(PDF_REGEN_TARGETS);
  const list = names.map(function(n, i) { return (i + 1) + '. ' + n + ' (' + PDF_REGEN_TARGETS[n].label + ')'; }).join('\n');
  const pick = ui.prompt('Regenerate PDF',
    'Which register?\n\n' + list + '\n\nEnter number 1..' + names.length + ':',
    ui.ButtonSet.OK_CANCEL);
  if (pick.getSelectedButton() !== ui.Button.OK) return;
  const idx = parseInt(pick.getResponseText().trim(), 10) - 1;
  if (isNaN(idx) || idx < 0 || idx >= names.length) { ui.alert('Invalid selection.'); return; }
  const name = names[idx];
  const spec = PDF_REGEN_TARGETS[name];

  const docPrompt = ui.prompt('Regenerate PDF',
    name + ' — enter the Doc No (e.g., GRN-0042, PO-0007):',
    ui.ButtonSet.OK_CANCEL);
  if (docPrompt.getSelectedButton() !== ui.Button.OK) return;
  const docNo = docPrompt.getResponseText().trim();
  if (!docNo) { ui.alert('Doc No required.'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) { ui.alert(name + ' is empty.'); return; }
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const rows = data.filter(function(r) { return r[spec.docCol - 1] === docNo; });
  if (!rows.length) { ui.alert('Doc "' + docNo + '" not found in ' + name + '.'); return; }

  try {
    const url = spec.fn(docNo, rows);
    writeBackPdfLink(sh, spec.docCol, docNo, spec.linkCol, url);
    ui.alert('PDF regenerated for ' + docNo + ':\n\n' + url);
  } catch (e) {
    ui.alert('PDF regeneration failed:\n\n' + e.message);
  }
}

// ---------------------------------------------------------------------------
// EXPORT EXPLORER AS CSV
// Reads the Explorer sheet's data area (from row 2, its used range) → writes
// CSV to the PDF Docs Folder in Drive → shows link.
// ---------------------------------------------------------------------------

function exportExplorerAsCsv() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Explorer');
  if (!sh) { ui.alert('Explorer sheet missing.'); return; }
  const view = sh.getRange('B1').getValue() || 'view';
  const arg  = sh.getRange('D1').getValue();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 3 || lastCol < 1) { ui.alert('Explorer is empty — pick a view first.'); return; }

  // CSV escape: wrap in quotes if contains comma/quote/newline; double internal quotes.
  function esc(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const s = String(v);
    return /[,"\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const csv = data.map(function(row) { return row.map(esc).join(','); }).join('\r\n');

  const folderId = getConfig('PDF Docs Folder ID');
  const parent = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const label = String(view).replace(/[^A-Za-z0-9]+/g, '_') + (arg ? '_' + String(arg).replace(/[^A-Za-z0-9]+/g, '_') : '');
  const filename = 'Explorer_' + label + '_' + stamp + '.csv';
  const file = parent.createFile(filename, csv, MimeType.CSV);
  ui.alert('CSV exported:\n\n' + file.getUrl());
}