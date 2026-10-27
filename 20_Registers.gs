/**
 * 3 Star Projects IMS v3 — 20_Registers.gs
 * Phase 2: Inward, Requisition, Adjustments — validated posting.
 * Phase 3 will add Job Work Outward/Inward posting to this same file.
 *
 * Contract: each post* fn takes an already-validated formData shape from
 * 40_Forms.gs, writes rows atomically (LockService), calls the snapshot
 * updater for every affected item/SFG, and returns a result summary that
 * the form handler wraps as { success, message, code }.
 *
 * Deps: 00_Config_Menu_Setup.gs (getConfig/setConfig), 30_Snapshots.gs
 *       (updateItemSnapshot), 60_Utils_Audit_Triggers.gs.
 */

// ============================================================================
// SHARED HELPERS
// ============================================================================

// Read a master sheet once and return { code: row } — replaces N per-line
// getRowByCode calls with a single sheet read. Cuts most-forms save time by
// 200-400ms per line for large masters.
function _buildCodeMap(sh) {
  if (!sh || sh.getLastRow() < 2) return {};
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const out = {};
  for (let i = 0; i < data.length; i++) {
    if (data[i][0]) out[data[i][0]] = data[i];
  }
  return out;
}

// Board-type conversion. GRN/PO forms capture pcs + sqm/unit + per-sqm rate
// (bill style). Stock is tracked in sheets, so convert before storing:
//   qty_stored = pcs,  rate_stored = rate × sqm/unit
// Line total is algebraically unchanged. Non-board rows pass through.
// Shared by postInwardRows, postPoRows, amendPoRows, and the moving-avg loop.
function _boardConvert(line) {
  const sqmU = Number(line.sqmPerUnit) || 0;
  const pcs  = Number(line.pcs) || 0;
  if (sqmU > 0 && pcs > 0) return { qty: pcs, rate: (Number(line.rate) || 0) * sqmU };
  return { qty: Number(line.qty), rate: Number(line.rate) };
}

// Find the 1-indexed column number of 'Avg Rate (Rs.)' in Item Master by
// header lookup. Returns 0 if missing. Resilient to column position changes
// (migration, user edits, extra columns). Index into row arrays: col - 1.
function _avgRateCol(itemsSh) {
  if (!itemsSh) return 0;
  const lastCol = itemsSh.getLastColumn();
  if (lastCol < 1) return 0;
  const headers = itemsSh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf('Avg Rate (Rs.)');
  return idx >= 0 ? (idx + 1) : 0;
}

// ============================================================================
// INWARD — GRN posting
// ============================================================================

function postInwardRows(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inward = ss.getSheetByName('Inward Register');
  const items = ss.getSheetByName('Item Master');
  const suppliers = ss.getSheetByName('Supplier Master');
  if (!inward || !items || !suppliers) throw new Error('Missing sheet. Run Initialize Workbook.');

  const supplierRow = getRowByCode(suppliers, formData.supplierCode);
  if (!supplierRow) throw new Error('Supplier not found: ' + formData.supplierCode);
  const supplierName = supplierRow[1];
  const dateObj = parseDate(formData.date);
  const user = currentUserEmail();
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const grnNo = getNextCounter('Next GRN Counter', 'GRN', 4);
    const startRow = nextRowByCol(inward, 1);
    const rows = [];
    const touchedItems = [];

    // PERF: read Item Master ONCE and build a code-indexed map.
    // Was calling getRowByCode per line (each call scans all 391 rows) — reduces
    // 2N master reads to 1 for N-line GRNs.
    const itemsMap = _buildCodeMap(items);

    formData.lineItems.forEach(function(line, idx) {
      const itemRow = itemsMap[line.itemCode];
      if (!itemRow) throw new Error('Item not found: ' + line.itemCode);
      const desc = itemRow[2] || '';
      const uom = itemRow[6] || ''; // v3.5: Desc col C(2), UOM col G(6)
      const { qty, rate } = _boardConvert(line);
      const gstPct = Number(line.gst) || 0;
      const gross = qty * rate;
      // ponytail: discountMode 'amt' means user entered a flat Rs figure instead
      // of a percent. Both cols (K=%, L=Rs) still get populated — % is derived
      // for display when amt is given, so GRN PDF + reports stay consistent.
      let discPct, discAmt;
      if (line.discountMode === 'amt') {
        discAmt = Number(line.discount) || 0;
        discPct = gross > 0 ? (discAmt / gross) * 100 : 0;
      } else {
        discPct = Number(line.discount) || 0;
        discAmt = gross * discPct / 100;
      }
      const gstAmt = (gross - discAmt) * gstPct / 100;
      const total = gross - discAmt + gstAmt;

      rows.push([
        startRow + idx - 1,             // Sr No = row index (1-based excl header)
        grnNo,
        dateObj,
        line.itemCode, desc,
        formData.supplierCode, supplierName,
        qty, uom, rate,
        discPct, discAmt,
        gstPct, gstAmt,
        total,
        formData.invoiceNo || '',
        line.remarks || '',
        '',                             // GRN PDF Link (Phase 4 fills)
        user, now,
      ]);
      touchedItems.push(line.itemCode);
    });

    // Capture pre-GRN Item Master Avg Rate + prior physical qty (from Snapshot)
    // for the moving-weighted-avg update. MUST happen BEFORE snapshot rebuild.
    const avgRateCol = _avgRateCol(items);
    const avgRateIdx = avgRateCol > 0 ? avgRateCol - 1 : -1;
    const priorAvg = {};
    const priorPhysical = {};
    formData.lineItems.forEach(function(l) {
      const im = itemsMap[l.itemCode];
      if (im && avgRateIdx >= 0) priorAvg[l.itemCode] = Number(im[avgRateIdx]) || 0;
    });
    const snapSh = ss.getSheetByName('Stock Snapshot');
    if (snapSh) {
      try {
        _readRmSection(snapSh, 10).forEach(function(r) {
          if (r[0]) priorPhysical[r[0]] = (Number(r[8]) || 0) + (Number(r[9]) || 0);  // Available + Reserved
        });
      } catch (e) { Logger.log('Prior physical read failed: ' + e.message); }
    }

    inward.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    // ponytail: force "%" display format on the two percent cols at write time —
    // belt+suspenders in case the legacy '0.00%' format (which multiplies ×100)
    // is still stuck on the sheet from before the Formatting fix was applied.
    inward.getRange(startRow, 11, rows.length, 1).setNumberFormat('0.00"%"');  // K Discount %
    inward.getRange(startRow, 13, rows.length, 1).setNumberFormat('0.00"%"');  // M GST %

    // Taxable Amount col V — pre-tax line value (Qty × Rate − Discount). Matches
    // what Indian supplier bills show in their "Amount" column so GRN reconciles
    // 1:1 with the invoice. Written separately (not in rows[]) so PO Ref at col U
    // doesn't get shifted.
    const taxCol = _colByHeader(inward, 'Taxable Amount');
    if (taxCol > 0) {
      const taxVals = rows.map(function(r) {
        return [(Number(r[7]) || 0) * (Number(r[9]) || 0) - (Number(r[11]) || 0)];
      });
      inward.getRange(startRow, taxCol, taxVals.length, 1).setValues(taxVals)
            .setNumberFormat('₹#,##0.00;[Red]-₹#,##0.00;""');
    }

    // Moving weighted average: update Item Master col U per code. Board-mode
    // rows are converted to per-sheet via _boardConvert so moving-avg stays
    // in sheet terms and never mixes sqm into a per-sheet rate.
    //   new_avg = (old_qty × old_avg + Σ(grn_qty × effective_rate)) / (old_qty + Σ grn_qty)
    //   effective_rate = rate × (1 − discount%). GST excluded (recoverable).
    const grnAgg = {};
    formData.lineItems.forEach(function(l) {
      const { qty: q, rate: r } = _boardConvert(l);
      const eff = r * (1 - (Number(l.discount) || 0) / 100);
      if (q <= 0 || eff <= 0) return;
      if (!grnAgg[l.itemCode]) grnAgg[l.itemCode] = { qty: 0, sum: 0 };
      grnAgg[l.itemCode].qty += q;
      grnAgg[l.itemCode].sum += q * eff;
    });
    Object.keys(grnAgg).forEach(function(code) {
      const g = grnAgg[code];
      const oldQty = priorPhysical[code] || 0;
      const oldAvg = priorAvg[code] || 0;
      const totalQty = oldQty + g.qty;
      const newAvg = totalQty > 0
        ? (oldQty * oldAvg + g.sum) / totalQty
        : (g.sum / g.qty);
      const row = _findRowByCode(items, code);
      if (row > 0 && avgRateCol > 0) items.getRange(row, avgRateCol).setValue(Math.round(newAvg * 100) / 100);
    });

    // Recompute snapshot for every item touched (dedupe) — uses the fresh Avg Rate above
    updateItemSnapshotsBatch(touchedItems);

    // PDF (only if user didn't opt to skip) — skip the expensive photo read otherwise.
    let pdf = null;
    if (!formData.skipPdf) {
      const photoMap = getMasterPhotoUrlMap('Item Master');
      const pdfLines = formData.lineItems.map(function(line, idx) {
        const r = rows[idx];
        const im = itemsMap[r[3]];   // reuse cached map — was per-line getRowByCode
        return { itemCode: r[3], itemDesc: r[4], qty: r[7], uom: r[8], rate: r[9],
                 discPct: r[10], gstPct: r[12], total: r[14],
                 hsn: im ? (im[11] || '') : '',
                 photoUrl: photoMap[r[3]] || '' };
      });
      // Supplier extras for GST + address block: cols D=Contact(3), E=Mobile(4), F=Email(5),
      // G=GST(6), H=Address(7). Company state comes from Config.
      pdf = _safePdf(function() {
        return generateGrnPdf(grnNo, {
          date: dateObj, supplierCode: formData.supplierCode, supplierName: supplierName,
          supplierGstin: supplierRow[6] || '', supplierAddress: supplierRow[7] || '',
          supplierState: '', // spec has no state col; derived from GSTIN prefix by _isIntraState
          supplierContact: supplierRow[3] || '', supplierMobile: supplierRow[4] || '', supplierEmail: supplierRow[5] || '',
          invoiceNo: formData.invoiceNo,
        }, pdfLines);
      }, 'generateGrnPdf', grnNo);
      if (pdf) writeBackPdfLink(inward, 2, grnNo, 18, pdf.url); // col B = GRN No, col R = PDF Link
    }

    // PO → GRN linkage: stamp PO Ref (col U/21) for every inserted row + recompute PO status.
    if (formData.poRef) {
      const poRefCol = _colByHeader(inward, 'PO Ref');
      if (poRefCol > 0) {
        inward.getRange(startRow, poRefCol, rows.length, 1).setValues(
          rows.map(function() { return [formData.poRef]; })
        );
      }
      try { updatePoStatus(formData.poRef); } catch (e) { Logger.log('updatePoStatus failed: ' + e.message); }
    }

    return { grnNo: grnNo, rows: rows.length, pdfUrl: pdf ? pdf.url : null, pdfDownloadUrl: pdf ? pdf.downloadUrl : null };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// MATERIAL REQUISITION — internal outward
// ============================================================================

function postRequisitionRows(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = ss.getSheetByName('Material Requisition Register');
  const items = ss.getSheetByName('Item Master');
  if (!reg || !items) throw new Error('Missing sheet. Run Initialize Workbook.');

  const dateObj = parseDate(formData.date);
  const user = currentUserEmail();
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const reqNo = getNextCounter('Next Requisition Counter', 'REQ', 4);
    const startRow = nextRowByCol(reg, 1);
    const rows = [];
    const touched = [];

    // PERF: read Item Master once
    const itemsMap = _buildCodeMap(items);

    formData.lineItems.forEach(function(line, idx) {
      const itemRow = itemsMap[line.itemCode];
      if (!itemRow) throw new Error('Item not found: ' + line.itemCode);
      rows.push([
        startRow + idx - 1,             // Sr No
        reqNo,
        dateObj,
        formData.project || '',
        line.itemCode, itemRow[2] || '',
        Number(line.qty), itemRow[6] || '', // v3.5: Desc C(2), UOM G(6)
        line.purpose || '',
        user,                            // Requested By
        formData.takenBy || formData.approvedBy || '',  // Taken By (accepts legacy 'approvedBy')
        '',                              // Slip PDF Link
        line.remarks || '',
        user, now,
      ]);
      touched.push(line.itemCode);
    });

    reg.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    updateItemSnapshotsBatch(touched);

    const pdfLines = rows.map(function(r) {
      return { itemCode: r[4], itemDesc: r[5], qty: r[6], uom: r[7], purpose: r[8] };
    });
    const pdf = formData.skipPdf ? null : _safePdf(function() {
      return generateReqSlipPdf(reqNo, { date: dateObj, project: formData.project, takenBy: formData.takenBy || formData.approvedBy }, pdfLines);
    }, 'generateReqSlipPdf', reqNo);
    if (pdf) writeBackPdfLink(reg, 2, reqNo, 12, pdf.url); // col B = REQ No, col L = Slip PDF Link

    return { reqNo: reqNo, rows: rows.length, pdfUrl: pdf ? pdf.url : null, pdfDownloadUrl: pdf ? pdf.downloadUrl : null };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// ADJUSTMENTS — opening balances, damage, corrections
// ============================================================================

function postAdjustment(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = ss.getSheetByName('Adjustments Register');
  if (!reg) throw new Error('Missing sheet. Run Initialize Workbook.');

  const dateObj = parseDate(formData.date);
  const user = currentUserEmail();
  const now = new Date();

  // Lookup description + UOM — Item Master first, then SFG
  let master = ss.getSheetByName('Item Master');
  let row = getRowByCode(master, formData.code);
  let isSfg = false;
  if (!row) {
    master = ss.getSheetByName('SFG Master');
    row = getRowByCode(master, formData.code);
    isSfg = !!row;
  }
  if (!row) throw new Error('Code not found in Item or SFG Master: ' + formData.code);

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const startRow = nextRowByCol(reg, 1);
    const newRow = [
      startRow - 1,                      // Sr No
      dateObj,
      formData.type,
      formData.code,
      row[2] || '',                      // v3.5: Description col C (formula)
      Number(formData.qty),              // signed
      row[6] || '',                      // v3.5: UOM col G
      formData.reason || '',
      formData.remarks || '',
      user, now,
    ];
    reg.getRange(startRow, 1, 1, newRow.length).setValues([newRow]);

    if (isSfg) updateSfgSnapshot(formData.code);
    else       updateItemSnapshot(formData.code);
    return { code: formData.code, sr: startRow - 1 };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// JOB WORK OUTWARD — SFG to processor (powder coat, etc.)
// ============================================================================

function postJobWorkOutwardRows(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = ss.getSheetByName('Job Work Outward Register');
  const jws = ss.getSheetByName('Job Worker Master');
  const sfgs = ss.getSheetByName('SFG Master');
  const items = ss.getSheetByName('Item Master');  // RM may also be sent for JW
  if (!reg || !jws || !sfgs) throw new Error('Missing sheet. Run Initialize Workbook.');

  const jwRow = getRowByCode(jws, formData.jobWorkerCode);
  if (!jwRow) throw new Error('Job Worker not found: ' + formData.jobWorkerCode);
  const jwName = jwRow[1];
  const processType = jwRow[2];
  const dateObj = parseDate(formData.date);
  const user = currentUserEmail();
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const jwcNo = getNextCounter('Next Job Work Challan Counter', 'JWC', 4);
    const startRow = nextRowByCol(reg, 1);
    const rows = [];
    const touchedSfg = [], touchedRm = [];

    // PERF: read SFG Master + Item Master once each
    const sfgMap = _buildCodeMap(sfgs);
    const itemsMap = items ? _buildCodeMap(items) : {};

    formData.lineItems.forEach(function(line, idx) {
      // ponytail: field is still called sfgCode for backward compat with existing
      // callers; it may hold either an SFG or an RM (Item Master) code. SFG prefix
      // doesn't collide with any Item Master prefix, so look up in SFG first then
      // fall back to Item Master.
      const code = line.sfgCode;
      let masterRow = sfgMap[code];
      let kind = 'SFG';
      if (!masterRow && itemsMap[code]) { masterRow = itemsMap[code]; kind = 'RM'; }
      if (!masterRow) throw new Error('Code not found in SFG Master or Item Master: ' + code);
      const expReturn = line.expectedReturn ? parseDate(line.expectedReturn) : '';

      rows.push([
        startRow + idx - 1,                       // Sr No           A
        jwcNo,                                    //                 B
        dateObj,                                  //                 C
        formData.jobWorkerCode, jwName, processType,  //             D E F
        code, masterRow[2] || '',                 // Code, Desc      G H
        Number(line.qty), masterRow[6] || '',     // Qty, UOM        I J
        expReturn,                                // Expected Return K
        line.remarks || '',                       // Remarks         L
        '',                                        // Challan PDF     M
        'Open',                                    // Status          N
        user, now,                                // Created By/At   O P
        Number(line.weight) || 0,                 // Weight (KG)     Q
      ]);
      if (kind === 'SFG') touchedSfg.push(code);
      else                touchedRm.push(code);
    });

    reg.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    if (touchedSfg.length) updateSfgSnapshotsBatch(touchedSfg);
    if (touchedRm.length)  updateItemSnapshotsBatch(touchedRm);

    // Only load heavy photo map when actually generating a PDF
    let pdf = null;
    if (!formData.skipPdf) {
      // Pull photos from both masters — JW line may be an RM code.
      const sfgPhotos = getMasterPhotoUrlMap('SFG Master');
      const itemPhotos = getMasterPhotoUrlMap('Item Master');
      const pdfLines = rows.map(function(r) {
        const photo = sfgPhotos[r[6]] || itemPhotos[r[6]] || '';
        return { sfgCode: r[6], desc: r[7], qty: r[8], uom: r[9], expReturn: r[10], remarks: r[11],
                 weight: r[16] || 0, photoUrl: photo };
      });
      pdf = _safePdf(function() {
        return generateJobWorkChallanPdf(jwcNo, {
          date: dateObj, jobWorkerCode: formData.jobWorkerCode, jobWorkerName: jwName,
          processType: processType, jobWorkerGstin: jwRow[6] || '', jobWorkerAddress: jwRow[7] || '',
        }, pdfLines);
      }, 'generateJobWorkChallanPdf', jwcNo);
      if (pdf) writeBackPdfLink(reg, 2, jwcNo, 13, pdf.url); // col B = JWC No, col M = Challan PDF Link
    }

    return { jwcNo: jwcNo, rows: rows.length, pdfUrl: pdf ? pdf.url : null, pdfDownloadUrl: pdf ? pdf.downloadUrl : null };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// JOB WORK INWARD — SFG returned from processor. Closes matching outward.
// ============================================================================

function postJobWorkInwardRows(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = ss.getSheetByName('Job Work Inward Register');
  const outReg = ss.getSheetByName('Job Work Outward Register');
  if (!reg || !outReg) throw new Error('Missing sheet. Run Initialize Workbook.');

  // Look up job worker name once from any outward row with this JWC No
  const outRows = getRowsByValue(outReg, 2, formData.referenceChallan);
  if (!outRows.length) throw new Error('Reference challan not found: ' + formData.referenceChallan);
  const firstOut = outReg.getRange(outRows[0], 1, 1, outReg.getLastColumn()).getValues()[0];
  const jwName = firstOut[4];

  const dateObj = parseDate(formData.date);
  const user = currentUserEmail();
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const jwgrnNo = getNextCounter('Next Job Work GRN Counter', 'JWGRN', 4);
    const startRow = nextRowByCol(reg, 1);
    const rows = [];
    const touched = [];

    // Build a map of outward rows keyed by SFG for qty-sent lookup
    const outBySfg = {};
    outRows.forEach(function(r) {
      const rowVals = outReg.getRange(r, 1, 1, outReg.getLastColumn()).getValues()[0];
      outBySfg[rowVals[6]] = { row: r, qtySent: rowVals[8], uom: rowVals[9], desc: rowVals[7] };
    });

    formData.lineItems.forEach(function(line, idx) {
      const outInfo = outBySfg[line.sfgCode];
      if (!outInfo) throw new Error('SFG ' + line.sfgCode + ' not on challan ' + formData.referenceChallan);
      const qtyRecv = Number(line.qtyReceived);
      const yieldLoss = Number(line.yieldLoss) || 0;
      const rejection = Number(line.rejection) || 0;

      rows.push([
        startRow + idx - 1,                       // Sr No
        jwgrnNo,
        dateObj,
        formData.referenceChallan,
        jwName,
        line.sfgCode, outInfo.desc,
        qtyRecv, outInfo.uom,
        yieldLoss, rejection,
        line.remarks || '',
        '',                                        // GRN PDF Link
        user, now,
      ]);
      touched.push(line.sfgCode);
    });

    reg.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    // Flip every matching outward row to Closed (col N = 14)
    outRows.forEach(function(r) { outReg.getRange(r, 14).setValue('Closed'); });

    // ponytail: a JW challan may now contain RM codes too (Change 2: items→JW).
    // Split touched codes by which master owns them so the right snapshot runs.
    const sfgSh = ss.getSheetByName('SFG Master');
    const itemsSh = ss.getSheetByName('Item Master');
    const sfgMap = sfgSh ? _buildCodeMap(sfgSh) : {};
    const itemsMap = itemsSh ? _buildCodeMap(itemsSh) : {};
    const sfgTouched = [], rmTouched = [];
    touched.forEach(function(c) {
      if (sfgMap[c])        sfgTouched.push(c);
      else if (itemsMap[c]) rmTouched.push(c);
    });
    if (sfgTouched.length) updateSfgSnapshotsBatch(sfgTouched);
    if (rmTouched.length)  updateItemSnapshotsBatch(rmTouched);

    const pdfLines = rows.map(function(r) {
      return { sfgCode: r[5], desc: r[6], qtyRecv: r[7], uom: r[8],
               yieldLoss: r[9], rejection: r[10], remarks: r[11] };
    });
    const pdf = formData.skipPdf ? null : _safePdf(function() {
      return generateJobWorkGrnPdf(jwgrnNo, { date: dateObj, referenceChallan: formData.referenceChallan,
        jobWorkerName: jwName }, pdfLines);
    }, 'generateJobWorkGrnPdf', jwgrnNo);
    if (pdf) writeBackPdfLink(reg, 2, jwgrnNo, 13, pdf.url); // col B = JWGRN No, col M = GRN PDF Link

    return { jwgrnNo: jwgrnNo, rows: rows.length, closedChallan: formData.referenceChallan, pdfUrl: pdf ? pdf.url : null, pdfDownloadUrl: pdf ? pdf.downloadUrl : null };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// PURCHASE ORDER — writes PO Register header + PO Lines children atomically.
// No stock update (POs are intent; stock moves on GRN).
// Status defaults to 'Issued' since generating the PDF = sending to supplier.
// ============================================================================

function postPoRows(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const poReg = ss.getSheetByName('Purchase Order Register');
  const poLines = ss.getSheetByName('PO Lines');
  const items = ss.getSheetByName('Item Master');
  const suppliers = ss.getSheetByName('Supplier Master');
  if (!poReg || !poLines || !items || !suppliers) throw new Error('Missing sheet. Run Initialize Workbook.');

  const supplierRow = getRowByCode(suppliers, formData.supplierCode);
  if (!supplierRow) throw new Error('Supplier not found: ' + formData.supplierCode);
  const supplierName = supplierRow[1];
  const dateObj = parseDate(formData.date);
  const deliveryDate = formData.deliveryDate ? parseDate(formData.deliveryDate) : '';
  const user = currentUserEmail();
  const now = new Date();

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const poNo = getNextCounter('Next PO Counter', 'PO', 4);

    // Build line rows first so we know the grand total
    const lineRows = [];
    let grandTotal = 0;
    formData.lineItems.forEach(function(line, idx) {
      const itemRow = getRowByCode(items, line.itemCode);
      if (!itemRow) throw new Error('Item not found: ' + line.itemCode);
      const desc = itemRow[2] || '';                              // v3.5: col C (formula)
      const uom  = itemRow[6] || '';                              // v3.5: col G
      const { qty, rate } = _boardConvert(line);
      const gstPct  = Number(line.gst) || 0;
      const gross   = qty * rate;
      // ponytail: see postInwardRows — same discountMode branch for PO lines.
      let discPct, discAmt;
      if (line.discountMode === 'amt') {
        discAmt = Number(line.discount) || 0;
        discPct = gross > 0 ? (discAmt / gross) * 100 : 0;
      } else {
        discPct = Number(line.discount) || 0;
        discAmt = gross * discPct / 100;
      }
      const gstAmt  = (gross - discAmt) * gstPct / 100;
      const total   = gross - discAmt + gstAmt;
      grandTotal   += total;
      lineRows.push([
        poNo, idx + 1,
        line.itemCode, desc,
        qty, uom, rate,
        discPct, discAmt, gstPct, gstAmt, total,
      ]);
    });

    // PO Register header row (cols A-L):
    //   Sr, PO No, Date, Supplier Code, Supplier Name, Delivery Date, Status,
    //   Total Value, PO PDF Link, Remarks, Created By, Created At
    const headerRow = nextRowByCol(poReg, 1);
    const poHeaderVals = [
      headerRow - 1,             // Sr No
      poNo,
      dateObj,
      formData.supplierCode,
      supplierName,
      deliveryDate,
      'Issued',                  // Status
      grandTotal,                // Total Value
      '',                        // PO PDF Link (filled below)
      formData.remarks || '',
      user, now,
    ];
    poReg.getRange(headerRow, 1, 1, poHeaderVals.length).setValues([poHeaderVals]);

    // PO Lines child rows
    const linesStartRow = nextRowByCol(poLines, 1);
    poLines.getRange(linesStartRow, 1, lineRows.length, lineRows[0].length).setValues(lineRows);

    // Best-effort PDF
    const poPhotoMap = getMasterPhotoUrlMap('Item Master');
    const pdfLines = formData.lineItems.map(function(line, idx) {
      const r = lineRows[idx];
      const im = getRowByCode(items, r[2]);
      return {
        itemCode: r[2], itemDesc: r[3], qty: r[4], uom: r[5], rate: r[6],
        discPct: r[7], gstPct: r[9], total: r[11],
        hsn: im ? (im[11] || '') : '',
        photoUrl: poPhotoMap[r[2]] || '',
      };
    });
    const pdf = formData.skipPdf ? null : _safePdf(function() {
      return generatePoPdf(poNo, {
        date: dateObj, supplierCode: formData.supplierCode,
        // Supplier — verify-block overrides win over Master lookup when provided
        supplierName:    formData.supplierNameOverride    || supplierName,
        supplierGstin:   formData.supplierGstinOverride   || (supplierRow[6] || ''),
        supplierAddress: formData.supplierAddressOverride || (supplierRow[7] || ''),
        supplierContact: formData.supplierContactOverride ||
          [supplierRow[3], supplierRow[4], supplierRow[5]].filter(String).join(' | '),
        // Contact is a single composed string when overridden; also keep raw fields
        supplierMobile: supplierRow[4] || '', supplierEmail: supplierRow[5] || '',
        deliveryDate: deliveryDate, remarks: formData.remarks,
        // Bill To overrides (falls back to Config in the PDF layer)
        billToName:    formData.billToName    || '',
        billToAddress: formData.billToAddress || '',
        billToGstin:   formData.billToGstin   || '',
        billToContact: formData.billToContact || '',
        // Ship To overrides (falls back to Bill To / Config)
        shipToName:    formData.shipToName    || '',
        shipToAddress: formData.shipToAddress || '',
        shipToGstin:   formData.shipToGstin   || '',
        shipToContact: formData.shipToContact || '',
      }, pdfLines);
    }, 'generatePoPdf', poNo);
    if (pdf) writeBackPdfLink(poReg, 2, poNo, 9, pdf.url); // col B = PO No, col I = PDF Link

    return { poNo: poNo, rows: lineRows.length, total: grandTotal, pdfUrl: pdf ? pdf.url : null, pdfDownloadUrl: pdf ? pdf.downloadUrl : null };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// Internal helpers
// ============================================================================

// PDF gen is best-effort. Register writes are already committed by this point;
// a Drive permission / template error must not surface as a data-loss error to
// the user. Logs to Errors sheet, returns null on failure.
function _safePdf(fn, fnName, docNo) {
  try { return fn(); }
  catch (e) {  return null; }
}


// ============================================================================
// FG DELIVERY (Delivery Challan) — v3.5
// Delivers FGs to a customer + optionally sends hardware for installation.
// - Main-section lines with valid FG code → FG On Hand deducts via formula
// - Hardware-section lines with valid Item Code → RM stock deducts via
//   updateItemSnapshotsBatch (recomputes via _computeRmRowFromMaps which
//   includes FG Delivery Lines hardware section as an outward)
// - DC No = <Prefix>/<FY>/<seq>, e.g., TSP/2026-27/031. Counter per FY,
//   auto-created in Config as "DC Counter <FY>".
// ============================================================================

function _fiscalYear(d) {
  // Indian FY: Apr 1 → Mar 31. Returns "2026-27" for a date in Apr 2026..Mar 2027.
  const dt = d instanceof Date ? d : new Date();
  const m = dt.getMonth();  // 0-based
  const y = dt.getFullYear();
  const startY = (m >= 3) ? y : y - 1;
  const endY = String(startY + 1).slice(-2).padStart(2, '0');
  return startY + '-' + endY;
}

function _nextDcNo(dateObj) {
  // Caller (postFgDeliveryRows) already holds LockService.getDocumentLock().
  const fy = _fiscalYear(dateObj);
  const key = 'DC Counter ' + fy;
  let cur = Number(getConfig(key)) || 0;
  cur += 1;
  setConfig(key, cur);
  const prefix = getConfig('DC No Prefix') || 'TSP';
  return prefix + '/' + fy + '/' + String(cur).padStart(3, '0');
}

function postFgDeliveryRows(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = ss.getSheetByName('FG Delivery Register');
  const lines = ss.getSheetByName('FG Delivery Lines');
  if (!reg || !lines) throw new Error('FG Delivery sheets missing. Run Initialize Workbook.');
  if (!formData.mainItems || !formData.mainItems.length) throw new Error('At least one main item required.');

  const dateObj = parseDate(formData.date);
  const user = currentUserEmail();
  const now = new Date();
  const items = ss.getSheetByName('Item Master');
  const fgs   = ss.getSheetByName('FG Master');

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    // Allow manual DC No override from form. Reject blank/duplicate.
    let dcNo;
    if (formData.manualDcNo && String(formData.manualDcNo).trim()) {
      dcNo = String(formData.manualDcNo).trim();
      const dupReg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FG Delivery Register');
      if (dupReg && dupReg.getLastRow() >= 2) {
        const dcVals = dupReg.getRange(2, 2, dupReg.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < dcVals.length; i++) {
          if (dcVals[i][0] === dcNo) throw new Error('DC No "' + dcNo + '" already exists — pick a different one.');
        }
      }
    } else {
      dcNo = _nextDcNo(dateObj);
    }

    // Combine + normalize lines. Section = 'Main' (FGs) or 'Hardware' (Items).
    const allLines = [];
    let totalQty = 0, totalWt = 0;
    (formData.mainItems || []).forEach(function(l, i) {
      const qty = Number(l.qty) || 0;
      const wt  = Number(l.weight) || 0;
      totalQty += qty; totalWt += wt;
      allLines.push({
        section: 'Main', code: l.itemCode || '', desc: l.description || '',
        qty: qty, uom: l.uom || 'Nos', weight: wt, remark: l.remark || '',
      });
    });
    (formData.hardwareItems || []).forEach(function(l, i) {
      const qty = Number(l.qty) || 0;
      if (qty <= 0) return;
      totalQty += qty;
      allLines.push({
        section: 'Hardware', code: l.itemCode || '', desc: l.description || '',
        qty: qty, uom: l.uom || 'Nos', weight: 0, remark: l.remark || '',
      });
    });

    // Register row — 21 cols with Type at end. Type defaults to 'Customer' for
    // backward compat with older callers that don't send it.
    const deliveryType = (formData.deliveryType === 'Internal') ? 'Internal' : 'Customer';
    const regStart = nextRowByCol(reg, 1);
    reg.getRange(regStart, 1, 1, 21).setValues([[
      regStart - 1, dcNo, dateObj,
      formData.customerName || '', formData.project || '',
      formData.paymentTerms || '', formData.otherRef || '',
      formData.billTo || '', formData.billToGstin || '',
      formData.shipTo || '', formData.shipToGstin || '',
      formData.vehicleNo || '', formData.driver || '',
      formData.lrNo || '', formData.eWayBillNo || '',
      totalQty, totalWt, '', user, now, deliveryType,
    ]]);

    // Lines rows
    const lineStart = nextRowByCol(lines, 1);
    const lineRows = allLines.map(function(l, i) {
      return [dcNo, i + 1, l.section, l.code, l.desc, l.qty, l.uom, l.weight, l.remark, 0];
    });
    if (lineRows.length) lines.getRange(lineStart, 1, lineRows.length, 10).setValues(lineRows);

    // Stock refresh:
    //   FG main lines → FG Master On Hand formula auto-updates
    //   Main lines that are actually SFG codes → refresh SFG snapshot (fgDelMain deducts)
    //   Hardware lines with valid Item Code → refresh RM snapshot
    const sfgMaster = ss.getSheetByName('SFG Master');
    const mainSfgCodes = Array.from(new Set(allLines.filter(function(l) {
      return l.section === 'Main' && l.code && sfgMaster && getRowByCode(sfgMaster, l.code);
    }).map(function(l) { return l.code; })));
    updateSfgSnapshotsBatch(mainSfgCodes);

    const hardwareCodes = Array.from(new Set(allLines.filter(function(l) {
      return l.section === 'Hardware' && l.code;
    }).map(function(l) { return l.code; })));
    const validHwCodes = hardwareCodes.filter(function(c) { return items && getRowByCode(items, c); });
    updateItemSnapshotsBatch(validHwCodes);

    // PDF
    const pdf = formData.skipPdf ? null : _safePdf(function() {
      return generateFgDeliveryPdf(dcNo, {
        date: dateObj,
        customerName: formData.customerName, project: formData.project,
        paymentTerms: formData.paymentTerms, otherRef: formData.otherRef,
        billTo: formData.billTo, billToGstin: formData.billToGstin,
        shipTo: formData.shipTo, shipToGstin: formData.shipToGstin,
        vehicleNo: formData.vehicleNo, driver: formData.driver,
        lrNo: formData.lrNo, eWayBillNo: formData.eWayBillNo,
        totalQty: totalQty, totalWt: totalWt,
      }, allLines);
    }, 'generateFgDeliveryPdf', dcNo);
    if (pdf) writeBackPdfLink(reg, 2, dcNo, 18, pdf.url);  // col B = DC No, col R = PDF Link

    return { dcNo: dcNo, rows: lineRows.length, pdfUrl: pdf ? pdf.url : null, pdfDownloadUrl: pdf ? pdf.downloadUrl : null };
  } finally {
    lock.releaseLock();
  }
}