/**
 * 3 Star Projects IMS v3 — 40_Forms.gs
 * Form submit handlers called from 05_Forms.html via google.script.run.
 *
 * Phase 1 (this file): masters — addItem, addSfg, addSupplier, addJobWorker.
 * Phase 2/3 will add:  addInwardMultiLine, addRequisitionMultiLine,
 *                      addJobWorkOutMultiLine,
 *                      addJobWorkInMultiLine.
 *
 * All handlers return { success, message, code? } shape. Errors log to Errors
 * sheet and rethrow — 05_Forms.html shows a toast either way.
 *
 * Deps: 00_Config_Menu_Setup.gs, 10_Masters.gs, 60_Utils_Audit_Triggers.gs.
 */

// ============================================================================
// Master forms
// ============================================================================


// ---------------------------------------------------------------------------
// Duplicate detection — matches on normalized (Item Name + Size + Grade/Color).
// Returns { code, desc } of the existing row if found, else null.
// Skips inactive rows. Case-insensitive, whitespace-tolerant.
// ---------------------------------------------------------------------------
function _findDuplicateMaster(sheetName, itemName, size, color) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return null;
  const lastRow = nextRowByCol(sh, 1) - 1;
  if (lastRow < 2) return null;
  // _colByHeader returns -1 on miss; coerce to 0 so `||` fallbacks work
  // (in JS, -1 is truthy — without this, SFG/FG never fall through to their
  // color-column header and duplicate detection silently ignores color).
  const orZero = function(v) { return v > 0 ? v : 0; };
  const nameCol = orZero(_colByHeader(sh, 'Item Name'));
  const sizeCol = orZero(_colByHeader(sh, 'Size'));
  // Grade col varies: Item = Grade/Color, SFG = Color/Finish, FG = Colour
  const colorCol = orZero(_colByHeader(sh, 'Grade/Color'))
                || orZero(_colByHeader(sh, 'Color/Finish'))
                || orZero(_colByHeader(sh, 'Colour'));
  const descCol = orZero(_colByHeader(sh, 'Description'));
  if (nameCol < 1) return null;
  const width = Math.max(nameCol, sizeCol, colorCol, descCol);
  const data = sh.getRange(2, 1, lastRow - 1, width).getValues();
  const norm = function(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); };
  const key = norm(itemName) + '|' + norm(size) + '|' + norm(color);
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    const rowKey = norm(r[nameCol - 1]) + '|' + norm(sizeCol > 0 ? r[sizeCol - 1] : '') + '|' + norm(colorCol > 0 ? r[colorCol - 1] : '');
    if (rowKey === key) return { code: r[0], desc: descCol > 0 ? r[descCol - 1] : (r[nameCol - 1] || '') };
  }
  return null;
}

function addItem(d) {
  try {
    _require(d, ['category', 'itemName', 'uom', 'materialType']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Item Master');
    if (!sh) throw new Error('Item Master sheet missing — run Initialize Workbook.');
    _assertValueInDropdown(ss, 'A', d.category, 'Category');
    _assertValueInDropdown(ss, 'E', d.uom, 'UOM');
    _assertValueInDropdown(ss, 'C', d.materialType, 'Material Type');
    // Supplier field accepts comma-separated codes (e.g., 'SUPP-0001, SUPP-0002').
    // Validate each one exists.
    if (d.supplier) {
      String(d.supplier).split(',').map(function(c) { return c.trim(); }).filter(Boolean)
        .forEach(function(code) {
          _assertMasterCodeExists(ss, 'Supplier Master', code, 'Supplier');
        });
    }

    // Duplicate detection — return early with existing match unless forced.
    if (!d.forceCreate) {
      const dup = _findDuplicateMaster('Item Master', d.itemName, d.size, d.grade);
      if (dup) return { success: false, duplicate: true, existingCode: dup.code, existingDesc: dup.desc,
        message: 'Possible duplicate of ' + dup.code + ' (' + dup.desc + ')' };
    }
    const code = assignItemCode(d.category);
    // v3.5 Item Master schema:
    //   A=Code, B=Photo (manual CellImage — skip), C=Description (FORMULA — skip),
    //   D=Item Name, E=Size, F=Grade, G=UOM, H=Category, I=MatType, J=Storage,
    //   K=Rack, L=HSN, M=PrefSupplier, N=AvgDaily, O=LeadTime, P=Safety,
    //   Q=MIN, R=MAX, S=Active, T=Search Label (FORMULA — skip).
    const rowNum = nextRowByCol(sh, 1);
    sh.getRange(rowNum, 1).setValue(code);
    const dToS = [
      String(d.itemName).trim(),  // D: Item Name
      d.size || '',                // E
      d.grade || '',               // F
      d.uom,                       // G
      d.category,                  // H
      d.materialType,              // I
      d.storage || '',             // J
      d.rack || '',                // K
      d.hsn || '',                 // L
      d.supplier || '',            // M
      _num(d.avgDaily, 0),         // N
      _num(d.leadTime, 0),         // O
      _num(d.safety, 0),           // P
      _num(d.minLvl, 0),           // Q
      _num(d.maxLvl, 0),           // R
      true,                        // S Active
    ];
    sh.getRange(rowNum, 4, 1, dToS.length).setValues([dToS]);
    // Avg Rate — per-unit cost (per-sheet for board items). Dynamic column
    // lookup so this still works if the Avg Rate column ever shifts position.
    const rate = _num(d.avgRate, 0);
    const _arCol = (typeof _avgRateCol === 'function') ? _avgRateCol(sh) : 21;
    if (rate > 0 && _arCol > 0) sh.getRange(rowNum, _arCol).setValue(rate);
    // Opening Qty (optional) → logged as Adjustments Register row (type=Opening).
    // postAdjustment also refreshes the snapshot, so the item shows correct stock immediately.
    const openingQty = _num(d.openingQty, 0);
    if (openingQty > 0) {
      postAdjustment({
        date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        type: 'Opening', code: code, qty: openingQty,
        reason: 'Initial opening balance seeded on item creation', remarks: ''
      });
    }
    return { success: true, message: 'Item ' + code + ' added: ' + d.itemName +
      (openingQty > 0 ? ' (opening: ' + openingQty + ' ' + d.uom + ')' : ''), code: code };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addSfg(d) {
  try {
    _require(d, ['itemName', 'uom']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('SFG Master');
    if (!sh) throw new Error('SFG Master sheet missing — run Initialize Workbook.');
    _assertValueInDropdown(ss, 'E', d.uom, 'UOM');

    if (!d.forceCreate) {
      const dup = _findDuplicateMaster('SFG Master', d.itemName, d.size, d.color);
      if (dup) return { success: false, duplicate: true, existingCode: dup.code, existingDesc: dup.desc,
        message: 'Possible duplicate of ' + dup.code + ' (' + dup.desc + ')' };
    }
    const code = assignSfgCode();
    // v3.5 SFG Master schema:
    //   A=Code, B=Photo (manual), C=Desc(F), D=Item Name, E=Size, F=Color, G=UOM,
    //   H=Storage, I=MIN, J=MAX, K=Active, L=Opening Qty, M=On Hand(F), N=Search Label(F).
    const rowNum = nextRowByCol(sh, 1);
    sh.getRange(rowNum, 1).setValue(code);
    const dToK = [
      String(d.itemName).trim(),   // D
      d.size || '',                 // E
      d.color || '',                // F
      d.uom,                        // G
      d.storage || '',              // H
      _num(d.minLvl, 0),            // I
      _num(d.maxLvl, 0),            // J
      true,                         // K Active
    ];
    sh.getRange(rowNum, 4, 1, dToK.length).setValues([dToK]);
    // Opening Qty — write to SFG Master col L for reference AND post an "Opening"
    // adjustment so Snapshot's SFG section picks it up (Snapshot reads adjustments,
    // not the Master's Opening col).
    const openingQty = _num(d.openingQty, 0);
    sh.getRange(rowNum, _colByHeader(sh, 'Opening Qty')).setValue(openingQty);
    if (openingQty > 0) {
      postAdjustment({
        date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        type: 'Opening', code: code, qty: openingQty,
        reason: 'Initial opening balance seeded on SFG creation', remarks: ''
      });
    }
    return { success: true, message: 'SFG ' + code + ' added: ' + d.itemName +
      (openingQty > 0 ? ' (opening: ' + openingQty + ' ' + d.uom + ')' : ''), code: code };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addSupplier(d) {
  try {
    _require(d, ['firmName']);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Supplier Master');
    if (!sh) throw new Error('Supplier Master sheet missing — run Initialize Workbook.');
    if (d.email && !_looksLikeEmail(d.email)) throw new Error('Email format invalid');
    if (d.gst) _assertGstFormat(d.gst);

    const code = assignSupplierCode();
    const row = [
      code, String(d.firmName).trim(), d.category || '', d.contact || '',
      d.mobile || '', d.email || '', d.gst || '', d.address || '', true];
    const rowNum = nextRowByCol(sh, 1);
    sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Supplier ' + code + ' added: ' + d.firmName, code: code };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addFg(d) {
  try {
    _require(d, ['itemName', 'uom']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('FG Master');
    if (!sh) throw new Error('FG Master sheet missing — run Initialize Workbook.');
    _assertValueInDropdown(ss, 'E', d.uom, 'UOM');

    if (!d.forceCreate) {
      const dup = _findDuplicateMaster('FG Master', d.itemName, d.size, d.colour);
      if (dup) return { success: false, duplicate: true, existingCode: dup.code, existingDesc: dup.desc,
        message: 'Possible duplicate of ' + dup.code + ' (' + dup.desc + ')' };
    }
    const code = assignFgCode();
    // v3.5 FG Master schema:
    //   A=Code, B=Photo (manual), C=Desc(F), D=Item Name, E=Size, F=Colour,
    //   G=Finish, H=UOM, I=Active, J=Opening Qty, K=On Hand(F), L=Search Label(F).
    const rowNum = nextRowByCol(sh, 1);
    sh.getRange(rowNum, 1).setValue(code);
    const dToI = [
      String(d.itemName).trim(),   // D
      d.size || '',                 // E
      d.colour || '',               // F
      d.finish || '',               // G
      d.uom,                        // H
      true,                         // I Active
    ];
    sh.getRange(rowNum, 4, 1, dToI.length).setValues([dToI]);
    // Opening Qty — for FG, the Master's Opening Qty column IS the source of
    // truth (formula reads J directly). No adjustment record needed.
    const openingQty = _num(d.openingQty, 0);
    sh.getRange(rowNum, _colByHeader(sh, 'Opening Qty')).setValue(openingQty);
    return { success: true, message: 'FG ' + code + ' added: ' + d.itemName +
      (openingQty > 0 ? ' (opening: ' + openingQty + ' ' + d.uom + ')' : ''), code: code };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addTool(d) {
  try {
    _require(d, ['itemName', 'type']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Tool Master');
    if (!sh) throw new Error('Tool Master sheet missing — run Initialize Workbook.');
    if (['Power Tool','Hand Tool','Machine','Consumable','Other'].indexOf(d.type) < 0) {
      throw new Error('Type must be Power Tool / Hand Tool / Machine / Consumable / Other');
    }
    const code = assignToolCode();
    // v3.5 Tool Master: A=Code, B=Photo (manual), C=Desc(F), D=Item Name,
    //   E=Type, F=Serial, G=PurchaseDate, H=PurchaseValue, I=Vendor,
    //   J=Status, K=CurrentlyWith, L=Opening Qty, M=On Hand(F), N=Active,
    //   O=Search Label(F).
    const rowNum = nextRowByCol(sh, 1);
    sh.getRange(rowNum, 1).setValue(code);
    const dToK = [
      String(d.itemName).trim(),        // D
      d.type,                            // E
      d.serial || '',                    // F
      d.purchaseDate ? new Date(d.purchaseDate) : '',  // G
      _num(d.purchaseValue, 0),          // H
      d.vendor || '',                    // I
      'In Store',                        // J default status
      '',                                // K Currently With — blank until issued
    ];
    sh.getRange(rowNum, 4, 1, dToK.length).setValues([dToK]);
    sh.getRange(rowNum, _colByHeader(sh, 'Opening Qty')).setValue(1);
    sh.getRange(rowNum, _colByHeader(sh, 'Active?')).setValue(true);
    return { success: true, message: 'Tool ' + code + ' added: ' + d.itemName, code: code };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function issueTool(d) {
  try {
    _require(d, ['date', 'toolCode', 'person']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const toolSh = ss.getSheetByName('Tool Master');
    const movSh  = ss.getSheetByName('Tool Movement Register');
    if (!toolSh || !movSh) throw new Error('Tool sheets missing — run Initialize Workbook.');
    const toolRow = getRowByCode(toolSh, d.toolCode);
    if (!toolRow) throw new Error('Tool not found: ' + d.toolCode);
    if (toolRow[9] === 'Issued' || toolRow[9] === 'At Site') {   // v3.5: Status col J(9)
      throw new Error('Tool ' + d.toolCode + ' is already out with ' + (toolRow[10] || 'someone'));  // CurrWith K(10)
    }

    const user = currentUserEmail();
    const now = new Date();
    const movNo = getNextCounter('Next Movement Counter', 'TM', 4);
    const dateObj = parseDate(d.date);
    const expReturn = d.expectedReturn ? parseDate(d.expectedReturn) : '';

    // Update Tool Master status + currently-with (cols I=9, J=10)
    const toolRowNum = _findRowByCode(toolSh, d.toolCode);
    toolSh.getRange(toolRowNum, 10).setValue(d.atSite ? 'At Site' : 'Issued');  // v3.5: Status J(10)
    toolSh.getRange(toolRowNum, 11).setValue(d.person);  // CurrentlyWith K(11)

    // Append movement row (using col A/B key to survive ARRAYFORMULA spills)
    const startRow = nextRowByCol(movSh, 1);
    const movRowVals = [
      startRow - 1, movNo, dateObj, d.toolCode, toolRow[2] || '',  // v3.5: Desc C(2)
      'Issue', d.person, d.purpose || '', expReturn, '',
      '', d.remarks || '', '',
      user, now];
    movSh.getRange(startRow, 1, 1, movRowVals.length).setValues([movRowVals]);
    return { success: true, message: 'Tool ' + d.toolCode + ' issued to ' + d.person + ' (' + movNo + ')', code: movNo };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function returnTool(d) {
  try {
    _require(d, ['date', 'refMovementNo']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const toolSh = ss.getSheetByName('Tool Master');
    const movSh  = ss.getSheetByName('Tool Movement Register');
    if (!toolSh || !movSh) throw new Error('Tool sheets missing.');
    // Find the original Issue movement row
    const movRow = _findRowByCode(movSh, d.refMovementNo, 2); // col B = Movement No
    if (!movRow) throw new Error('Movement not found: ' + d.refMovementNo);
    const origRow = movSh.getRange(movRow, 1, 1, movSh.getLastColumn()).getValues()[0];
    if (origRow[9]) throw new Error('Movement ' + d.refMovementNo + ' already marked returned on ' + dateString(origRow[9]));

    const toolCode = origRow[3];
    const toolDesc = origRow[4];
    const user = currentUserEmail();
    const now = new Date();
    const newMovNo = getNextCounter('Next Movement Counter', 'TM', 4);
    const dateObj = parseDate(d.date);

    // Update Tool Master back to In Store, clear "currently with"
    const toolRowNum = _findRowByCode(toolSh, toolCode);
    if (toolRowNum) {
      toolSh.getRange(toolRowNum, 10).setValue(d.condition === 'Needs Repair' ? 'Under Repair' : (d.condition === 'Lost' ? 'Lost' : 'In Store'));  // v3.5: Status J(10)
      toolSh.getRange(toolRowNum, 11).setValue('');  // CurrentlyWith K(11)
    }
    // Mark original Issue with Actual Return Date (col J = 10)
    movSh.getRange(movRow, 10).setValue(dateObj);
    // Append Return movement row
    const startRow = nextRowByCol(movSh, 1);
    const retRowVals = [
      startRow - 1, newMovNo, dateObj, toolCode, toolDesc,
      'Return', origRow[6] || '', origRow[7] || '', '', dateObj,
      d.condition || 'Good', d.remarks || '', d.refMovementNo,
      user, now];
    movSh.getRange(startRow, 1, 1, retRowVals.length).setValues([retRowVals]);
    return { success: true, message: 'Tool ' + toolCode + ' returned (' + newMovNo + ' ← ' + d.refMovementNo + ')', code: newMovNo };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- BOM CREATE ---
function addBom(d) {
  try {
    _require(d, ['parentType', 'parentCode', 'lineItems']);
    if (['FG', 'SFG'].indexOf(d.parentType) < 0) throw new Error('Makes Type must be FG or SFG');
    if (!Array.isArray(d.lineItems) || !d.lineItems.length) throw new Error('At least one Needs line required');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bom = ss.getSheetByName('BOM Master');
    if (!bom) throw new Error('BOM Master sheet missing.');
    // Verify makes-product exists
    const parentSh = ss.getSheetByName(d.parentType === 'FG' ? 'FG Master' : 'SFG Master');
    if (!getRowByCode(parentSh, d.parentCode)) throw new Error(d.parentCode + ' not found in ' + parentSh.getName());

    // Find next available BOM ID (max existing + 1) so we don't collide
    let nextBom = 1;
    if (bom.getLastRow() >= 2) {
      const existing = bom.getRange(2, 1, bom.getLastRow() - 1, 1).getValues();
      existing.forEach(function(r) {
        const m = /^BOM-(\d+)$/.exec(String(r[0] || ''));
        if (m) { const n = parseInt(m[1], 10); if (n >= nextBom) nextBom = n + 1; }
      });
    }

    const rows = [];
    d.lineItems.forEach(function(line, i) {
      if (['RM', 'SFG'].indexOf(line.componentType) < 0) throw new Error('Line ' + (i+1) + ': Needs Type must be RM or SFG');
      _require(line, ['componentType', 'componentCode', 'qty']);
      if (Number(line.qty) <= 0) throw new Error('Line ' + (i+1) + ': Qty must be > 0');
      const compSh = ss.getSheetByName(line.componentType === 'RM' ? 'Item Master' : 'SFG Master');
      const compRow = getRowByCode(compSh, line.componentCode);
      if (!compRow) throw new Error('Line ' + (i+1) + ': Component ' + line.componentCode + ' not found');
      // UOM: use whatever the component sheet has at its UOM col (lookup by header — resilient to shifts)
      const uomCol = _colByHeader(compSh, 'UOM');
      const uom = (uomCol > 0) ? (compRow[uomCol - 1] || '') : '';
      const bomId = 'BOM-' + String(nextBom + rows.length).padStart(4, '0');
      rows.push([
        bomId, d.parentType, d.parentCode, line.componentType, line.componentCode,
        Number(line.qty), uom, line.notes || '']);
    });
    const writeStart = nextRowByCol(bom, 1);
    bom.getRange(writeStart, 1, rows.length, rows[0].length).setValues(rows);
    return { success: true, message: 'BOM created: ' + rows.length + ' component(s) for ' + d.parentCode, code: d.parentCode };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- LIST all BOM parents (for Edit BOM picker) ---
function listBomParents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bom = ss.getSheetByName('BOM Master');
  if (!bom || bom.getLastRow() < 2) return [];
  const data = bom.getRange(2, 1, bom.getLastRow() - 1, 8).getValues();
  const fgs = ss.getSheetByName('FG Master');
  const sfgs = ss.getSheetByName('SFG Master');
  const groups = {};
  data.forEach(function(r) {
    if (!r[1] || !r[2]) return;
    const key = r[1] + '|' + r[2];
    if (!groups[key]) groups[key] = { parentType: r[1], parentCode: r[2], lineCount: 0 };
    groups[key].lineCount++;
  });
  return Object.keys(groups).map(function(k) {
    const g = groups[k];
    const master = g.parentType === 'FG' ? fgs : sfgs;
    const mrow = getRowByCode(master, g.parentCode);
    g.description = mrow ? (mrow[2] || '') : '';
    g.label = g.parentType + ' ' + g.parentCode + ' — ' + g.description + '  (' + g.lineCount + ' needs)';
    return g;
  }).sort(function(a, b) { return a.label < b.label ? -1 : 1; });
}

// --- GET all lines of one BOM (for Edit BOM form load) ---
function getBomForEdit(parentType, parentCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bom = ss.getSheetByName('BOM Master');
  if (!bom || bom.getLastRow() < 2) return [];
  const data = bom.getRange(2, 1, bom.getLastRow() - 1, 8).getValues();
  const out = [];
  data.forEach(function(r) {
    if (r[1] !== parentType || r[2] !== parentCode) return;
    const compSheet = r[3] === 'RM' ? 'Item Master' : 'SFG Master';
    const compRow = getRowByCode(ss.getSheetByName(compSheet), r[4]);
    out.push({
      componentType: r[3], componentCode: r[4],
      componentDesc: compRow ? (compRow[2] || '') : '',
      qty: r[5], uom: r[6], notes: r[7] || '',
    });
  });
  return out;
}

// --- EDIT BOM: replace all rows for the parent with new lines (atomic) ---
function editBom(d) {
  try {
    _require(d, ['parentType', 'parentCode', 'lineItems']);
    if (['FG', 'SFG'].indexOf(d.parentType) < 0) throw new Error('Parent Type must be FG or SFG');
    if (!Array.isArray(d.lineItems) || !d.lineItems.length) throw new Error('At least one component required');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bom = ss.getSheetByName('BOM Master');
    if (!bom) throw new Error('BOM Master sheet missing.');

    // Validate all lines first (fail before any mutation)
    d.lineItems.forEach(function(line, i) {
      if (['RM','SFG'].indexOf(line.componentType) < 0) throw new Error('Line ' + (i+1) + ': Component Type must be RM or SFG');
      if (!line.componentCode) throw new Error('Line ' + (i+1) + ': Component Code required');
      if (!(Number(line.qty) > 0)) throw new Error('Line ' + (i+1) + ': Qty must be > 0');
      const compSh = ss.getSheetByName(line.componentType === 'RM' ? 'Item Master' : 'SFG Master');
      if (!getRowByCode(compSh, line.componentCode)) throw new Error('Line ' + (i+1) + ': Component ' + line.componentCode + ' not found');
    });

    const lock = LockService.getDocumentLock();
    lock.waitLock(15000);
    try {
      // Delete existing rows for this parent (bottom-up to keep indices valid)
      const lastRow = bom.getLastRow();
      if (lastRow >= 2) {
        const scan = bom.getRange(2, 1, lastRow - 1, 3).getValues();
        for (let i = scan.length - 1; i >= 0; i--) {
          if (scan[i][1] === d.parentType && scan[i][2] === d.parentCode) {
            bom.deleteRow(i + 2);
          }
        }
      }
      // Find next BOM ID
      let nextBom = 1;
      if (bom.getLastRow() >= 2) {
        const existing = bom.getRange(2, 1, bom.getLastRow() - 1, 1).getValues();
        existing.forEach(function(r) {
          const m = /^BOM-(\d+)$/.exec(String(r[0] || ''));
          if (m) { const n = parseInt(m[1], 10); if (n >= nextBom) nextBom = n + 1; }
        });
      }
      const rows = d.lineItems.map(function(line, i) {
        const compSh = ss.getSheetByName(line.componentType === 'RM' ? 'Item Master' : 'SFG Master');
        const compRow = getRowByCode(compSh, line.componentCode);
        const uomCol = _colByHeader(compSh, 'UOM');
        const uom = (uomCol > 0 && compRow) ? (compRow[uomCol - 1] || '') : '';
        return ['BOM-' + String(nextBom + i).padStart(4, '0'),
                d.parentType, d.parentCode,
                line.componentType, line.componentCode,
                Number(line.qty), uom, line.notes || ''];
      });
      const writeStart = nextRowByCol(bom, 1);
      bom.getRange(writeStart, 1, rows.length, rows[0].length).setValues(rows);
      return { success: true, message: 'BOM for ' + d.parentCode + ' updated: ' + rows.length + ' component(s).', code: d.parentCode };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- DELETE BOM: remove all rows for a parent ---
function deleteBom(d) {
  try {
    _require(d, ['parentType', 'parentCode']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bom = ss.getSheetByName('BOM Master');
    if (!bom || bom.getLastRow() < 2) throw new Error('BOM Master empty.');
    const lock = LockService.getDocumentLock();
    lock.waitLock(15000);
    try {
      const data = bom.getRange(2, 1, bom.getLastRow() - 1, 3).getValues();
      let deleted = 0;
      for (let i = data.length - 1; i >= 0; i--) {
        if (data[i][1] === d.parentType && data[i][2] === d.parentCode) {
          bom.deleteRow(i + 2);
          deleted++;
        }
      }
      if (!deleted) throw new Error('No BOM found for ' + d.parentType + ' ' + d.parentCode);
      return { success: true, message: 'BOM for ' + d.parentCode + ' deleted (' + deleted + ' rows).' };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- KIT BUILDER: pick FG + qty → auto-create Requisition covering all components ---
// Validates a BOM before Kit Builder uses it. Throws on missing components,
// zero qtys, or invalid types. Also checks SFG child BOMs referenced by parent.
function _validateBomComplete(bomData, fgCode) {
  const parentLines = bomData.filter(function(r) { return r[1] === 'FG' && r[2] === fgCode; });
  if (!parentLines.length) throw new Error('No BOM found for FG ' + fgCode);
  parentLines.forEach(function(r, i) {
    if (!r[3])                 throw new Error('BOM ' + fgCode + ' line ' + (i+1) + ': Component Type is empty');
    if (!r[4])                 throw new Error('BOM ' + fgCode + ' line ' + (i+1) + ': Component Code is empty');
    if (['RM','SFG'].indexOf(r[3]) < 0) throw new Error('BOM ' + fgCode + ' line ' + (i+1) + ': Component Type must be RM or SFG (got "' + r[3] + '")');
    if (!(Number(r[5]) > 0))   throw new Error('BOM ' + fgCode + ' line ' + (i+1) + ' (' + r[4] + '): Qty per Unit must be > 0');
    // If component is SFG, its own BOM must exist AND be complete
    if (r[3] === 'SFG') {
      const childLines = bomData.filter(function(x) { return x[1] === 'SFG' && x[2] === r[4]; });
      if (!childLines.length) throw new Error('SFG ' + r[4] + ' has no BOM — required by ' + fgCode);
      childLines.forEach(function(cr, ci) {
        if (!cr[3] || !cr[4] || !(Number(cr[5]) > 0)) {
          throw new Error('SFG ' + r[4] + ' BOM line ' + (ci+1) + ': incomplete (type/code/qty missing)');
        }
      });
    }
  });
}

function buildKit(d) {
  try {
    _require(d, ['date', 'fgCode', 'qty']);
    const orderQty = Number(d.qty);
    if (orderQty <= 0) throw new Error('Order Qty must be > 0');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bom = ss.getSheetByName('BOM Master');
    if (!bom || bom.getLastRow() < 2) throw new Error('BOM Master empty — add a BOM first.');

    // Two-level explosion: for each FG BOM line, if it's SFG, recurse into its BOM.
    const bomData = bom.getRange(2, 1, bom.getLastRow() - 1, 8).getValues();
    _validateBomComplete(bomData, d.fgCode);  // fail fast on incomplete BOM
    const consumption = {}; // itemCode → qty
    function addComponent(code, qty) {
      consumption[code] = (consumption[code] || 0) + qty;
    }
    const l0 = bomData.filter(function(r) { return r[1] === 'FG' && r[2] === d.fgCode; });
    if (!l0.length) throw new Error('No BOM found for FG ' + d.fgCode);
    l0.forEach(function(r) {
      const type = r[3], compCode = r[4], qtyPer = Number(r[5]) || 0;
      const totalQty = qtyPer * orderQty;
      if (type === 'RM') {
        addComponent(compCode, totalQty);
      } else if (type === 'SFG') {
        // Recurse: SFG's own BOM (all RM)
        const l1 = bomData.filter(function(x) { return x[1] === 'SFG' && x[2] === compCode; });
        l1.forEach(function(x) {
          const childCode = x[4], childQtyPer = Number(x[5]) || 0;
          addComponent(childCode, childQtyPer * totalQty);
        });
      }
    });

    // Build requisition lines from consumption map
    const lineItems = Object.keys(consumption).map(function(code) {
      return { itemCode: code, qty: consumption[code], purpose: 'Kit for ' + d.fgCode + ' × ' + orderQty };
    });
    if (!lineItems.length) throw new Error('BOM explosion yielded zero components. Check BOM.');
    const reqRes = postRequisitionRows({
      date: d.date, project: 'Kit: ' + d.fgCode + ' × ' + orderQty, takenBy: '',
      lineItems: lineItems,
    });
    return { success: true, message: 'Requisition ' + reqRes.reqNo + ' created for ' + lineItems.length + ' components (kit of ' + d.fgCode + ' × ' + orderQty + ').', code: reqRes.reqNo, pdfDownloadUrl: reqRes.pdfDownloadUrl };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- EDIT MASTER handlers — pick record by code, update named fields ---
function editItem(d) {
  try {
    _require(d, ['code']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Item Master');
    const row = _findRowByCode(sh, d.code);
    if (!row) throw new Error('Item ' + d.code + ' not found');
    // v3.5 cols: D=Item Name, E=Size, F=Grade, G=UOM, H=Cat, I=MatType, J=Storage,
    //             K=Rack, L=HSN, M=PrefSupplier, Q=MIN, R=MAX, S=Active
    if (d.itemName    !== undefined) sh.getRange(row,  4).setValue(String(d.itemName).trim());
    if (d.size        !== undefined) sh.getRange(row,  5).setValue(d.size || '');
    if (d.grade       !== undefined) sh.getRange(row,  6).setValue(d.grade || '');
    if (d.uom         !== undefined) sh.getRange(row,  7).setValue(d.uom);
    if (d.category    !== undefined) sh.getRange(row,  8).setValue(d.category);
    if (d.materialType!== undefined) sh.getRange(row,  9).setValue(d.materialType);
    if (d.storage     !== undefined) sh.getRange(row, 10).setValue(d.storage || '');
    if (d.rack        !== undefined) sh.getRange(row, 11).setValue(d.rack || '');
    if (d.hsn         !== undefined) sh.getRange(row, 12).setValue(d.hsn || '');
    if (d.supplier    !== undefined) sh.getRange(row, 13).setValue(d.supplier || '');
    if (d.minLvl      !== undefined) sh.getRange(row, 17).setValue(_num(d.minLvl, 0));
    if (d.maxLvl      !== undefined) sh.getRange(row, 18).setValue(_num(d.maxLvl, 0));
    if (d.active      !== undefined) sh.getRange(row, 19).setValue(!!d.active);
    // Avg Rate — dynamic column lookup (survives migration / column shuffles).
    if (d.avgRate !== undefined) {
      const _arCol = (typeof _avgRateCol === 'function') ? _avgRateCol(sh) : 21;
      if (_arCol > 0) sh.getRange(row, _arCol).setValue(_num(d.avgRate, 0) || '');
    }
    return { success: true, message: 'Item ' + d.code + ' updated', code: d.code };
  } catch (e) {  return { success: false, message: e.message }; }
}

function editSfg(d) {
  try {
    _require(d, ['code']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('SFG Master');
    const row = _findRowByCode(sh, d.code);
    if (!row) throw new Error('SFG ' + d.code + ' not found');
    // v3.5 cols: D=Name, E=Size, F=Color, G=UOM, H=Storage, I=MIN, J=MAX, K=Active
    if (d.itemName !== undefined) sh.getRange(row,  4).setValue(String(d.itemName).trim());
    if (d.size     !== undefined) sh.getRange(row,  5).setValue(d.size || '');
    if (d.color    !== undefined) sh.getRange(row,  6).setValue(d.color || '');
    if (d.uom      !== undefined) sh.getRange(row,  7).setValue(d.uom);
    if (d.storage  !== undefined) sh.getRange(row,  8).setValue(d.storage || '');
    if (d.minLvl   !== undefined) sh.getRange(row,  9).setValue(_num(d.minLvl, 0));
    if (d.maxLvl   !== undefined) sh.getRange(row, 10).setValue(_num(d.maxLvl, 0));
    if (d.active   !== undefined) sh.getRange(row, 11).setValue(!!d.active);
    return { success: true, message: 'SFG ' + d.code + ' updated', code: d.code };
  } catch (e) {  return { success: false, message: e.message }; }
}

function editFg(d) {
  try {
    _require(d, ['code']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('FG Master');
    const row = _findRowByCode(sh, d.code);
    if (!row) throw new Error('FG ' + d.code + ' not found');
    // v3.5 cols: D=Name, E=Size, F=Colour, G=Finish, H=UOM, I=Active
    if (d.itemName !== undefined) sh.getRange(row, 4).setValue(String(d.itemName).trim());
    if (d.size     !== undefined) sh.getRange(row, 5).setValue(d.size || '');
    if (d.colour   !== undefined) sh.getRange(row, 6).setValue(d.colour || '');
    if (d.finish   !== undefined) sh.getRange(row, 7).setValue(d.finish || '');
    if (d.uom      !== undefined) sh.getRange(row, 8).setValue(d.uom);
    if (d.active   !== undefined) sh.getRange(row, 9).setValue(!!d.active);
    return { success: true, message: 'FG ' + d.code + ' updated', code: d.code };
  } catch (e) {  return { success: false, message: e.message }; }
}

function editSupplier(d) {
  try {
    _require(d, ['code']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Supplier Master');
    const row = _findRowByCode(sh, d.code);
    if (!row) throw new Error('Supplier ' + d.code + ' not found');
    if (d.firmName !== undefined) sh.getRange(row, 2).setValue(String(d.firmName).trim());
    if (d.category !== undefined) sh.getRange(row, 3).setValue(d.category || '');
    if (d.contact  !== undefined) sh.getRange(row, 4).setValue(d.contact || '');
    if (d.mobile   !== undefined) sh.getRange(row, 5).setValue(d.mobile || '');
    if (d.email    !== undefined) sh.getRange(row, 6).setValue(d.email || '');
    if (d.gst      !== undefined) sh.getRange(row, 7).setValue(d.gst || '');
    if (d.address  !== undefined) sh.getRange(row, 8).setValue(d.address || '');
    if (d.active   !== undefined) sh.getRange(row, 9).setValue(!!d.active);
    return { success: true, message: 'Supplier ' + d.code + ' updated', code: d.code };
  } catch (e) {  return { success: false, message: e.message }; }
}

function editJobWorker(d) {
  try {
    _require(d, ['code']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Job Worker Master');
    const row = _findRowByCode(sh, d.code);
    if (!row) throw new Error('Job Worker ' + d.code + ' not found');
    if (d.firmName    !== undefined) sh.getRange(row, 2).setValue(String(d.firmName).trim());
    if (d.processType !== undefined) sh.getRange(row, 3).setValue(d.processType);
    if (d.contact     !== undefined) sh.getRange(row, 4).setValue(d.contact || '');
    if (d.mobile      !== undefined) sh.getRange(row, 5).setValue(d.mobile || '');
    if (d.email       !== undefined) sh.getRange(row, 6).setValue(d.email || '');
    if (d.gst         !== undefined) sh.getRange(row, 7).setValue(d.gst || '');
    if (d.address     !== undefined) sh.getRange(row, 8).setValue(d.address || '');
    if (d.active      !== undefined) sh.getRange(row, 9).setValue(!!d.active);
    return { success: true, message: 'Job Worker ' + d.code + ' updated', code: d.code };
  } catch (e) {  return { success: false, message: e.message }; }
}

// --- Get row-by-code for the Edit form to pre-populate fields ---
function getMasterRowForEdit(sheetName, code) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return null;
  const row = getRowByCode(sh, code);
  return row;
}

// Helpers
function _findRowByCode(sh, code, colNum) {
  if (!sh || sh.getLastRow() < 2) return null;
  const col = colNum || 1;
  const vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === code) return i + 2;
  return null;
}

function addJobWorker(d) {
  try {
    _require(d, ['firmName', 'processType']);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Job Worker Master');
    if (!sh) throw new Error('Job Worker Master sheet missing — run Initialize Workbook.');
    _assertValueInDropdown(ss, 'G', d.processType, 'Process Type');
    if (d.email && !_looksLikeEmail(d.email)) throw new Error('Email format invalid');
    if (d.gst) _assertGstFormat(d.gst);

    const code = assignJobWorkerCode();
    const row = [
      code, String(d.firmName).trim(), d.processType, d.contact || '',
      d.mobile || '', d.email || '', d.gst || '', d.address || '', true];
    const rowNum = nextRowByCol(sh, 1);
    sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Job Worker ' + code + ' added: ' + d.firmName, code: code };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================================
// Register forms (Phase 2) — thin validation + delegation to 20_Registers.gs
// ============================================================================

function addInwardMultiLine(d) {
  try {
    _require(d, ['date', 'supplierCode', 'lineItems']);
    if (!Array.isArray(d.lineItems) || !d.lineItems.length) throw new Error('At least one line item required');
    d.lineItems.forEach(function(line, i) {
      _require(line, ['itemCode', 'qty', 'rate']);
      if (Number(line.qty) <= 0)  throw new Error('Line ' + (i+1) + ': Qty must be > 0');
      if (Number(line.rate) < 0)  throw new Error('Line ' + (i+1) + ': Rate must be >= 0');
    });
    const res = postInwardRows(d);
    return { success: true, message: "GRN " + res.grnNo + " created (" + res.rows + " lines). Stock updated.", code: res.grnNo, pdfDownloadUrl: res.pdfDownloadUrl };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addRequisitionMultiLine(d) {
  try {
    _require(d, ['date', 'lineItems']);
    if (!Array.isArray(d.lineItems) || !d.lineItems.length) throw new Error('At least one line item required');
    d.lineItems.forEach(function(line, i) {
      _require(line, ['itemCode', 'qty']);
      if (Number(line.qty) <= 0) throw new Error('Line ' + (i+1) + ': Qty must be > 0');
    });
    // Overselling preflight — skip if client already confirmed
    if (!d.bypassStockCheck) {
      const check = _checkStockAvailability(d.lineItems.map(function(l) { return { code: l.itemCode, qty: l.qty }; }));
      if (!check.ok) return { success: false, needsConfirm: true, warnings: check.warnings, message: 'Some lines exceed Available stock. Confirm to proceed anyway.' };
    }
    const res = postRequisitionRows(d);
    return { success: true, message: 'Requisition ' + res.reqNo + ' created (' + res.rows + ' lines).', code: res.reqNo, pdfDownloadUrl: res.pdfDownloadUrl };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addJobWorkOutMultiLine(d) {
  try {
    _require(d, ['date', 'jobWorkerCode', 'lineItems']);
    if (!Array.isArray(d.lineItems) || !d.lineItems.length) throw new Error('At least one line item required');
    d.lineItems.forEach(function(line, i) {
      _require(line, ['sfgCode', 'qty']);
      if (Number(line.qty) <= 0) throw new Error('Line ' + (i+1) + ': Qty must be > 0');
    });
    if (!d.bypassStockCheck) {
      // ponytail: _checkStockAvailability already handles RM + SFG + FG, so this
      // works unchanged now that line.sfgCode may hold an RM code.
      const check = _checkStockAvailability(d.lineItems.map(function(l) { return { code: l.sfgCode, qty: l.qty }; }));
      if (!check.ok) return { success: false, needsConfirm: true, warnings: check.warnings, message: 'Some lines exceed Available stock. Confirm to proceed anyway.' };
    }
    const res = postJobWorkOutwardRows(d);
    return { success: true, message: 'Job Work Challan ' + res.jwcNo + ' created (' + res.rows + ' line(s)). Stock updated.', code: res.jwcNo, pdfDownloadUrl: res.pdfDownloadUrl };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addJobWorkInMultiLine(d) {
  try {
    _require(d, ['date', 'referenceChallan', 'lineItems']);
    if (!Array.isArray(d.lineItems) || !d.lineItems.length) throw new Error('Enter Qty Received on at least one line');
    d.lineItems.forEach(function(line, i) {
      _require(line, ['sfgCode', 'qtyReceived']);
      if (Number(line.qtyReceived) <= 0) throw new Error('Line ' + (i+1) + ': Qty Received must be > 0');
    });
    const res = postJobWorkInwardRows(d);
    return { success: true, message: 'JWGRN ' + res.jwgrnNo + ' created; challan ' + res.closedChallan + ' closed.', code: res.jwgrnNo, pdfDownloadUrl: res.pdfDownloadUrl };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================================
// RFQ — Request for Quotation. No vendor field (broadcast to multiple suppliers).
// Client validates presence of at least one line with itemCode + qty > 0.
// ============================================================================
function addRfq(d) {
  try {
    _require(d, ['date', 'lineItems']);
    if (!Array.isArray(d.lineItems) || !d.lineItems.length) throw new Error('At least one line item required');
    d.lineItems.forEach(function(line, i) {
      _require(line, ['itemCode', 'qty']);
      if (Number(line.qty) <= 0) throw new Error('Line ' + (i + 1) + ': Qty must be > 0');
    });
    const res = postRfqRows(d);
    return {
      success: true,
      message: 'RFQ ' + res.rfqNo + ' created (' + res.rows + ' items). Share the PDF with your vendors.',
      code: res.rfqNo,
      pdfDownloadUrl: res.pdfDownloadUrl,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// RFQ is PDF-only — no register writes. Counter still used for a unique RFQ No.
// Line items can be RM (Item Master), SFG or FG — dropdown on the form combines all three.
function postRfqRows(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dateObj = parseDate(formData.date);
  const rfqNo = getNextCounter('Next RFQ Counter', 'RFQ', 4);

  // Build one lookup that covers all three master types. Photos + descriptions
  // all live in col B / C respectively on each master, so a unified per-code lookup works.
  // Per master, extract: desc (C=2), size, colour, uom — column indexes vary by master.
  // Item Master: desc=C(2), size=E(4), grade=F(5), uom=G(6)
  // SFG Master:  desc=C(2), size=E(4), color=F(5), uom=G(6)
  // FG Master:   desc=C(2), size=E(4), colour=F(5), finish=G(6), uom=H(7)
  const lookup = {};
  function addFromMaster(sh, uomIdx) {
    if (!sh || sh.getLastRow() < 2) return;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(uomIdx + 1, 8)).getValues();
    data.forEach(function(r) {
      if (r[0] && !lookup[r[0]]) lookup[r[0]] = { desc: r[2] || '', size: r[4] || '', colour: r[5] || '', uom: r[uomIdx] || '' };
    });
  }
  addFromMaster(ss.getSheetByName('Item Master'), 6);
  addFromMaster(ss.getSheetByName('SFG Master'),  6);
  addFromMaster(ss.getSheetByName('FG Master'),   7);

  // Photos: pull from whichever master owns the code (RM/SFG/FG).
  const photoMap = getMasterPhotoUrlMap('Item Master');
  const sfgPhotos = getMasterPhotoUrlMap('SFG Master');
  const fgPhotos  = getMasterPhotoUrlMap('FG Master');
  Object.keys(sfgPhotos).forEach(function(k) { if (!photoMap[k]) photoMap[k] = sfgPhotos[k]; });
  Object.keys(fgPhotos).forEach(function(k)  { if (!photoMap[k]) photoMap[k] = fgPhotos[k]; });

  const pdfLines = formData.lineItems.map(function(line) {
    const info = lookup[line.itemCode];
    if (!info) throw new Error('Code not found in any master: ' + line.itemCode);
    return {
      itemCode: line.itemCode,
      desc:    info.desc,
      size:    line.size   || info.size,
      colour:  line.colour || info.colour,
      uom:     info.uom,
      qty:     Number(line.qty) || 0,
      remarks: line.remarks || '',
      photoUrl: photoMap[line.itemCode] || '',
    };
  });

  const pdf = formData.skipPdf ? null : _safePdf(function() {
    return generateRfqPdf(rfqNo, {
      date: dateObj, project: formData.project, purpose: formData.purpose,
      validUntil: formData.validUntil, remarks: formData.remarks,
    }, pdfLines);
  }, 'generateRfqPdf', rfqNo);

  return {
    rfqNo: rfqNo, rows: pdfLines.length,
    pdfUrl: pdf ? pdf.url : null,
    pdfDownloadUrl: pdf ? pdf.downloadUrl : null,
  };
}

function addPoMultiLine(d) {
  try {
    _require(d, ['date', 'supplierCode', 'lineItems']);
    if (!Array.isArray(d.lineItems) || !d.lineItems.length) throw new Error('At least one line item required');
    d.lineItems.forEach(function(line, i) {
      _require(line, ['itemCode', 'qty', 'rate']);
      if (Number(line.qty) <= 0)  throw new Error('Line ' + (i+1) + ': Qty must be > 0');
      if (Number(line.rate) < 0)  throw new Error('Line ' + (i+1) + ': Rate must be >= 0');
    });
    const res = postPoRows(d);
    return {
      success: true,
      message: 'PO ' + res.poNo + ' created (' + res.rows + ' items, Rs. ' + res.total.toFixed(2) + '). No stock change until GRN.',
      code: res.poNo,
      pdfDownloadUrl: res.pdfDownloadUrl,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Return current PHYSICAL stock for a code (RM/SFG/FG). Used by Add Adjustment
// UI to show "what's on the shelf right now". Physical = Available + Reserved
// (RM/FG) OR On Hand (SFG — physical at warehouse, excludes JW-out).
function getPhysicalStock(code) {
  try {
    if (!code) throw new Error('Code required');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const snap = ss.getSheetByName('Stock Snapshot');
    if (!snap) throw new Error('Stock Snapshot missing');
    // RM: Available (idx 8) + Reserved (idx 9)
    const rm = _readRmSection(snap, 10);
    for (let i = 0; i < rm.length; i++) {
      if (rm[i][0] === code) return { success: true, code: code, physical: (Number(rm[i][8]) || 0) + (Number(rm[i][9]) || 0), uom: '' };
    }
    // SFG: On Hand (idx 3) is physical at warehouse
    const sfg = _readSfgSection(snap, 4);
    for (let j = 0; j < sfg.length; j++) {
      if (sfg[j][0] === code) return { success: true, code: code, physical: Number(sfg[j][3]) || 0, uom: '' };
    }
    // FG: Available (K) + active reservations
    const fgSh = ss.getSheetByName('FG Master');
    if (fgSh && fgSh.getLastRow() >= 2) {
      const fd = fgSh.getRange(2, 1, fgSh.getLastRow() - 1, 11).getValues();
      const resMap = _aggregateReservations(ss);
      for (let k = 0; k < fd.length; k++) {
        if (fd[k][0] === code) return { success: true, code: code, physical: (Number(fd[k][10]) || 0) + (Number(resMap[code]) || 0), uom: '' };
      }
    }
    return { success: false, message: 'Not found in Snapshot', physical: 0, uom: '' };
  } catch (e) {
    return { success: false, message: e.message, physical: 0, uom: '' };
  }
}

function addAdjustment(d) {
  try {
    _require(d, ['date', 'type', 'code', 'qty']);
    const validTypes = ['Opening', 'Damage', 'Correction', 'Return', 'Other'];
    if (validTypes.indexOf(d.type) < 0) throw new Error('Type must be one of: ' + validTypes.join(', '));
    if (Number(d.qty) === 0) throw new Error('Qty Adjustment cannot be zero (use +/-)');
    const res = postAdjustment(d);
    return { success: true, message: 'Adjustment posted for ' + res.code + ' (' + d.type + ' ' + d.qty + ').', code: res.code, pdfDownloadUrl: res.pdfDownloadUrl };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================================
// BULK ADJUSTMENTS — post multiple adjustments in one submission
// ============================================================================

// Returns current on-hand for a code from Stock Snapshot.
// Vertical layout: RM at rows 2..(sfgHdr-2), SFG at (sfgHdr+1)..end.
// Both sections have code in col A; RM On Hand at col I (idx 8), SFG On Hand at col D (idx 3).
function getOnHandForCode(code) {
  if (!code) return null;
  const snap = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock Snapshot');
  if (!snap || snap.getLastRow() < 2) return 0;
  const rm = _readRmSection(snap, 9);
  for (let i = 0; i < rm.length; i++) if (rm[i][0] === code) return Number(rm[i][8]) || 0;
  const sfg = _readSfgSection(snap, 4);
  for (let i = 0; i < sfg.length; i++) if (sfg[i][0] === code) return Number(sfg[i][3]) || 0;
  return 0;
}

// Physical count reconciliation. Each line provides `targetQty` (what the
// physical count shows). Server looks up current on-hand from Stock Snapshot,
// computes delta = target - current, and posts as a Correction adjustment.
// Lines where delta = 0 are skipped (no change needed).
function addBulkAdjustment(d) {
  try {
    _require(d, ['date', 'lineItems']);
    if (!Array.isArray(d.lineItems) || !d.lineItems.length) throw new Error('At least one adjustment line required');
    // ponytail: unified Adjustment form (bulk + single merged) sends d.type.
    // Legacy bulk callers that didn't send it default to 'Correction', the only
    // type the old bulk form ever wrote.
    const adjType = d.type || 'Correction';
    const validTypes = ['Opening', 'Damage', 'Correction', 'Return', 'Other'];
    if (validTypes.indexOf(adjType) < 0) throw new Error('Type must be one of: ' + validTypes.join(', '));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reg = ss.getSheetByName('Adjustments Register');
    const snap = ss.getSheetByName('Stock Snapshot');
    if (!reg) throw new Error('Adjustments Register missing.');
    if (!snap) throw new Error('Stock Snapshot missing — run Full Stock Rebuild first.');

    // Build PHYSICAL on-hand lookup maps for delta computation.
    // "Set to X" means physical count on shelf (includes reserved units — they're still present).
    // RM: Available (I) + Reserved (J) = Physical
    // SFG: Snapshot col D = physical at warehouse (already excludes JW-out and includes reserved)
    // FG: Master col K (Available) + Active reservations for that FG code = Physical
    const rmOnHand = {}, sfgOnHand = {}, fgOnHand = {}, toolOnHand = {};
    _readRmSection(snap, 10).forEach(function(r) {
      if (r[0]) rmOnHand[r[0]] = (Number(r[8]) || 0) + (Number(r[9]) || 0);
    });
    _readSfgSection(snap, 4).forEach(function(r) {
      if (r[0]) sfgOnHand[r[0]] = Number(r[3]) || 0;
    });
    const reservationMap = _aggregateReservations(ss);
    const fgSh = ss.getSheetByName('FG Master');
    if (fgSh && fgSh.getLastRow() >= 2) {
      const fd = fgSh.getRange(2, 1, fgSh.getLastRow() - 1, 11).getValues();
      fd.forEach(function(r) {
        if (r[0]) fgOnHand[r[0]] = (Number(r[10]) || 0) + (Number(reservationMap[r[0]]) || 0);
      });
    }
    const toolSh = ss.getSheetByName('Tool Master');
    if (toolSh && toolSh.getLastRow() >= 2) {
      // Tool Master: A=Code(0), M=OnHand(12)
      const td = toolSh.getRange(2, 1, toolSh.getLastRow() - 1, 13).getValues();
      td.forEach(function(r) { if (r[0]) toolOnHand[r[0]] = Number(r[12]) || 0; });
    }

    // Validate + compute deltas
    const dateObj = parseDate(d.date);
    const user = currentUserEmail();
    const now = new Date();
    const lock = LockService.getDocumentLock();
    lock.waitLock(15000);
    try {
      const startRow = nextRowByCol(reg, 1);
      const rows = [];
      const touchedCodes = [];
      const skipped = [];

      d.lineItems.forEach(function(line, idx) {
        _require(line, ['code']);
        // Two modes:
        //   deltaMode=true  → line.qty is a signed delta (+/-), posted as-is
        //   deltaMode=false → line.targetQty is the final on-hand, delta computed
        const deltaMode = !!line.deltaMode;
        const rawQty = Number(deltaMode ? line.qty : line.targetQty);
        if (isNaN(rawQty)) throw new Error('Line ' + (idx+1) + ': Qty must be a number');
        if (!deltaMode && rawQty < 0) throw new Error('Line ' + (idx+1) + ': Set-to qty must be >= 0');
        if (deltaMode && rawQty === 0) {
          skipped.push(line.code + ' (delta = 0)');
          return;
        }

        // Lookup description + UOM across all masters: Item → SFG → FG → Tool
        let master = ss.getSheetByName('Item Master');
        let masterRow = getRowByCode(master, line.code);
        let kind = 'RM';
        if (!masterRow) { master = ss.getSheetByName('SFG Master');  masterRow = getRowByCode(master, line.code); if (masterRow) kind = 'SFG'; }
        if (!masterRow) { master = ss.getSheetByName('FG Master');   masterRow = getRowByCode(master, line.code); if (masterRow) kind = 'FG'; }
        if (!masterRow) { master = ss.getSheetByName('Tool Master'); masterRow = getRowByCode(master, line.code); if (masterRow) kind = 'Tool'; }
        if (!masterRow) throw new Error('Line ' + (idx+1) + ': Code not found in Item / SFG / FG / Tool Master: ' + line.code);

        // UOM column varies: Item/SFG col G (6), FG col H (7), Tool has none — use blank
        const uom = kind === 'FG' ? (masterRow[7] || '') : (kind === 'Tool' ? '' : (masterRow[6] || ''));

        let delta, reason;
        if (deltaMode) {
          delta = rawQty;
          reason = line.reason || ('Bulk adjustment: ' + (rawQty > 0 ? '+' : '') + rawQty);
        } else {
          let currentOnHand;
          if (kind === 'RM')       currentOnHand = rmOnHand[line.code] || 0;
          else if (kind === 'SFG') currentOnHand = sfgOnHand[line.code] || 0;
          else if (kind === 'FG')  currentOnHand = fgOnHand[line.code] || 0;
          else                     currentOnHand = toolOnHand[line.code] || 0;
          delta = rawQty - currentOnHand;
          if (delta === 0) {
            skipped.push(line.code + ' (already at ' + rawQty + ')');
            return;
          }
          reason = line.reason || ('Bulk adjustment: set to ' + rawQty);
        }

        rows.push([
          startRow + rows.length - 1,
          dateObj,
          adjType,
          line.code,
          masterRow[2] || '',
          delta,
          uom,
          reason,
          line.remarks || '',
          user, now,
        ]);
        touchedCodes.push({ code: line.code, kind: kind });
      });

      if (!rows.length) {
        return { success: true, message: 'No changes needed. All ' + d.lineItems.length + ' items already at target qty.' };
      }

      reg.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

      // Refresh snapshots — RM/SFG go via batch; FG/Tool are formula-driven
      // (their master's On Hand col auto-recomputes from Adjustments Register).
      const rmCodes = [], sfgCodes = [];
      touchedCodes.forEach(function(item) {
        if (item.kind === 'RM')  rmCodes.push(item.code);
        if (item.kind === 'SFG') sfgCodes.push(item.code);
      });
      updateItemSnapshotsBatch(rmCodes);
      updateSfgSnapshotsBatch(sfgCodes);

      const msg = rows.length + ' item(s) reconciled.' +
        (skipped.length ? ' Skipped ' + skipped.length + ' (no change): ' + skipped.join(', ') : '');
      return { success: true, message: msg };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================================
// OVERSELLING VALIDATION — preflight check for outward transactions.
// Returns { ok: true } if all lines fit within Available; else { ok: false, warnings: [...] }.
// Called from post* handlers before writing rows.
// ============================================================================

// lines: array of { code, qty } — checks each against Available on-hand.
// masters: 'rm' (Item Master) or 'sfg' (SFG Master) — determines which snapshot section to read.
function _checkStockAvailability(lines) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snap = ss.getSheetByName('Stock Snapshot');
  if (!snap) return { ok: true, warnings: [] };  // Fresh install, no snapshot yet

  // Build maps: RM+SFG from Stock Snapshot, FG from FG Master's On Hand col.
  const rmAvail = {}, sfgAvail = {}, fgAvail = {};
  _readRmSection(snap, 9).forEach(function(r) {
    if (r[0]) rmAvail[r[0]] = Number(r[8]) || 0;   // v3.6 RM: idx 8 = Available
  });
  _readSfgSection(snap, 6).forEach(function(r) {
    if (r[0]) sfgAvail[r[0]] = Number(r[5]) || 0;  // SFG: idx 5 = Available
  });
  const fgSh = ss.getSheetByName('FG Master');
  if (fgSh && fgSh.getLastRow() >= 2) {
    // FG Master: A=Code(0), K=On Hand(10) — formula-driven, subtracts deliveries + reservations
    const fd = fgSh.getRange(2, 1, fgSh.getLastRow() - 1, 11).getValues();
    fd.forEach(function(r) { if (r[0]) fgAvail[r[0]] = Number(r[10]) || 0; });
  }

  // Pull descriptions once for helpful warning text
  const descMap = _buildDescMap(ss);

  const warnings = [];
  const totals = {};
  lines.forEach(function(l) {
    if (!l.code || !(Number(l.qty) > 0)) return;
    totals[l.code] = (totals[l.code] || 0) + Number(l.qty);
  });
  Object.keys(totals).forEach(function(code) {
    const requested = totals[code];
    let available, section;
    if (rmAvail[code]  != null) { available = rmAvail[code];  section = 'RM';  }
    else if (sfgAvail[code] != null) { available = sfgAvail[code]; section = 'SFG'; }
    else if (fgAvail[code]  != null) { available = fgAvail[code];  section = 'FG';  }
    else return;  // Code not in any master — new/manual item, skip check
    // Round to 4 dp — avoids false positives on floating-point noise (10.0000001 vs 10).
    if (Math.round(requested * 10000) > Math.round(available * 10000)) {
      warnings.push({
        code: code, desc: descMap[code] || '', section: section,
        requested: requested, available: available, shortfall: requested - available,
        systemNote: available < 0
          ? 'System shows NEGATIVE. Likely no Opening Qty was ever seeded — add an Opening Adjustment for the real physical stock.'
          : '',
      });
    }
  });
  return { ok: warnings.length === 0, warnings: warnings };
}

// Code → description map across all masters (used in warnings).
function _buildDescMap(ss) {
  const out = {};
  ['Item Master', 'SFG Master', 'FG Master'].forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    data.forEach(function(r) { if (r[0]) out[r[0]] = r[2] || ''; });  // col C = Desc (formula)
  });
  return out;
}

// ============================================================================
// RESERVATION — soft-lock stock for a specific order/customer/project.
// Snapshot's Available = physical on-hand − sum(Active reservations).
// Release is manual only — user picks a reservation and marks it Released or Consumed.
// ============================================================================

function addReservation(d) {
  try {
    _require(d, ['date', 'code', 'qty', 'forRef']);
    const qty = Number(d.qty);
    if (!(qty > 0)) throw new Error('Qty must be > 0');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reg = ss.getSheetByName('Reservation Register');
    if (!reg) throw new Error('Reservation Register missing — run Initialize Workbook.');

    // Lookup description + UOM across Item → SFG → FG masters
    let master = ss.getSheetByName('Item Master');
    let masterRow = getRowByCode(master, d.code);
    let kind = 'RM';
    if (!masterRow) { master = ss.getSheetByName('SFG Master'); masterRow = getRowByCode(master, d.code); if (masterRow) kind = 'SFG'; }
    if (!masterRow) { master = ss.getSheetByName('FG Master');  masterRow = getRowByCode(master, d.code); if (masterRow) kind = 'FG'; }
    if (!masterRow) throw new Error('Code not found: ' + d.code);
    const uom = kind === 'FG' ? (masterRow[7] || '') : (masterRow[6] || '');

    const lock = LockService.getDocumentLock();
    lock.waitLock(15000);
    try {
      const resNo = getNextCounter('Next Reservation Counter', 'RSV', 4);
      const startRow = nextRowByCol(reg, 1);
      const user = currentUserEmail();
      const now = new Date();
      reg.getRange(startRow, 1, 1, 14).setValues([[
        startRow - 1,          // A Sr
        resNo,                 // B Reservation No
        parseDate(d.date),     // C Date
        d.code,                // D Code
        masterRow[2] || '',    // E Description
        qty,                   // F Qty Reserved
        uom,                   // G UOM
        d.forRef,              // H For (Reference)
        'Active',              // I Status
        d.notes || '',         // J Notes
        user, now,             // K, L
        '', '',                // M Released/Consumed At, N Released By
      ]]);
      // Refresh snapshot so Available reflects the new reservation
      if (kind === 'RM')  updateItemSnapshotsBatch([d.code]);
      if (kind === 'SFG') updateSfgSnapshotsBatch([d.code]);
      return { success: true, message: 'Reserved ' + qty + ' of ' + d.code + ' as ' + resNo, code: resNo };
    } finally { lock.releaseLock(); }
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Return list of active reservations for picker
function getActiveReservations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = ss.getSheetByName('Reservation Register');
  if (!reg || reg.getLastRow() < 2) return [];
  const data = reg.getRange(2, 1, reg.getLastRow() - 1, 14).getValues();
  const out = [];
  data.forEach(function(r) {
    if (r[8] !== 'Active') return;
    out.push({
      resNo: r[1], code: r[3], desc: r[4], qty: r[5], uom: r[6], forRef: r[7],
      label: r[1] + '  |  ' + r[3] + ' × ' + r[5] + ' ' + (r[6] || '') + '  |  ' + r[7],
    });
  });
  return out.sort(function(a, b) { return a.resNo < b.resNo ? 1 : -1; });
}

// Mark reservation as Released or Consumed. Refreshes snapshot for the code.
function releaseReservation(d) {
  try {
    _require(d, ['resNo', 'status']);
    if (['Released','Consumed'].indexOf(d.status) < 0) throw new Error('Status must be Released or Consumed');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reg = ss.getSheetByName('Reservation Register');
    if (!reg) throw new Error('Reservation Register missing.');
    const rows = getRowsByValue(reg, 2, d.resNo);
    if (!rows.length) throw new Error('Reservation not found: ' + d.resNo);
    const rowNum = rows[0];
    const currentStatus = reg.getRange(rowNum, 9).getValue();
    if (currentStatus !== 'Active') throw new Error('Reservation ' + d.resNo + ' is already ' + currentStatus);
    const code = reg.getRange(rowNum, 4).getValue();

    reg.getRange(rowNum, 9).setValue(d.status);       // I Status
    reg.getRange(rowNum, 13).setValue(new Date());    // M Released/Consumed At
    reg.getRange(rowNum, 14).setValue(currentUserEmail());  // N Released By
    if (d.notes) {
      const existing = reg.getRange(rowNum, 10).getValue() || '';
      reg.getRange(rowNum, 10).setValue((existing ? existing + ' | ' : '') + d.notes);
    }

    // Refresh snapshot — determine kind via master lookup
    const items = ss.getSheetByName('Item Master');
    const sfgs  = ss.getSheetByName('SFG Master');
    if (items && getRowByCode(items, code)) updateItemSnapshotsBatch([code]);
    if (sfgs && getRowByCode(sfgs, code))   updateSfgSnapshotsBatch([code]);

    return { success: true, message: 'Reservation ' + d.resNo + ' marked ' + d.status };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================================
// QC RECORD — inspection log for GRN / JWGRN / Pre-Dispatch DC / Manual.
// No auto-actions on Reject/Rework — pure logging. User posts adjustments
// manually if stock changes are needed.
// ============================================================================

// Return recent doc numbers for the picker (last 90 days) sorted newest first.
function getQcRefNoList(sourceType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cutoff = new Date(Date.now() - 90 * 86400000);
  const specs = {
    'GRN':               { sheet: 'Inward Register',              docCol: 2, dateCol: 3 },
    'JWGRN':             { sheet: 'Job Work Inward Register',     docCol: 2, dateCol: 3 },
    'Pre-Dispatch DC':   { sheet: 'FG Delivery Register',         docCol: 2, dateCol: 3 },
  };
  const spec = specs[sourceType];
  if (!spec) return [];
  const sh = ss.getSheetByName(spec.sheet);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const seen = {}, out = [];
  data.forEach(function(r) {
    const docNo = r[spec.docCol - 1];
    const date  = r[spec.dateCol - 1];
    if (!docNo || seen[docNo]) return;
    if (date instanceof Date && date < cutoff) return;
    seen[docNo] = true;
    out.push({ docNo: docNo, date: date instanceof Date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MMM-yy') : '' });
  });
  return out.sort(function(a, b) { return a.docNo < b.docNo ? 1 : -1; });
}

// Return lines (materials) inside a picked doc so user can QC one specific item.
// Populates auto-fill fields on the QC form.
function getQcSourceLines(sourceType, refNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!refNo) return [];
  if (sourceType === 'GRN') {
    const sh = ss.getSheetByName('Inward Register');
    if (!sh || sh.getLastRow() < 2) return [];
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    // Inward cols: A=Sr,B=GRN,C=Date,D=ItemCode,E=Desc,F=SupCode,G=SupName,H=Qty,I=UOM,J=Rate,U=PORef
    const out = [];
    const poLines = ss.getSheetByName('PO Lines');
    let poData = null;
    data.forEach(function(r) {
      if (r[1] !== refNo) return;
      // Try to pull ordered qty from PO Lines using PO Ref (col U, idx 20)
      const poRef = r[20] || '';
      let orderedQty = '';
      if (poRef && poLines) {
        if (!poData) poData = poLines.getRange(2, 1, Math.max(1, poLines.getLastRow() - 1), 13).getValues();
        for (let i = 0; i < poData.length; i++) {
          if (poData[i][0] === poRef && poData[i][2] === r[3]) { orderedQty = poData[i][4]; break; }
        }
      }
      out.push({
        line: r[3],  // Item Code as line ID
        vendor: r[6], material: r[4], po: poRef, deliveryDate: r[2],
        orderedQty: orderedQty, receivedQty: r[7],
      });
    });
    return out;
  }
  if (sourceType === 'JWGRN') {
    const sh = ss.getSheetByName('Job Work Inward Register');
    if (!sh || sh.getLastRow() < 2) return [];
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    // JW In cols: A=Sr,B=JWGRN,C=Date,D=RefChallan,E=JWName,F=SFGCode,G=Desc,H=QtyRecv,I=UOM
    const out = [];
    const jwOut = ss.getSheetByName('Job Work Outward Register');
    let jwoData = null;
    data.forEach(function(r) {
      if (r[1] !== refNo) return;
      // Try to pull qty sent (as "ordered") from matching JW Outward row
      let orderedQty = '';
      if (r[3] && jwOut) {
        if (!jwoData) jwoData = jwOut.getRange(2, 1, Math.max(1, jwOut.getLastRow() - 1), jwOut.getLastColumn()).getValues();
        for (let i = 0; i < jwoData.length; i++) {
          if (jwoData[i][1] === r[3] && jwoData[i][6] === r[5]) { orderedQty = jwoData[i][8]; break; }
        }
      }
      out.push({
        line: r[5],  // SFG Code as line ID
        vendor: r[4], material: r[6], po: r[3], deliveryDate: r[2],
        orderedQty: orderedQty, receivedQty: r[7],
      });
    });
    return out;
  }
  if (sourceType === 'Pre-Dispatch DC') {
    const reg = ss.getSheetByName('FG Delivery Register');
    const lines = ss.getSheetByName('FG Delivery Lines');
    if (!reg || !lines || lines.getLastRow() < 2) return [];
    // Header row for context
    const regRow = getRowByCode(reg, refNo);
    const customer = regRow ? (regRow[3] || '') : '';
    const project  = regRow ? (regRow[4] || '') : '';
    const date     = regRow ? (regRow[2] || '') : '';
    const lineData = lines.getRange(2, 1, lines.getLastRow() - 1, 10).getValues();
    const out = [];
    lineData.forEach(function(r) {
      if (r[0] !== refNo) return;
      // DC Lines: A=DCNo, B=LineNo, C=Section, D=Code, E=Desc, F=Qty, G=UOM
      out.push({
        line: (r[3] || r[4] || 'Line ' + r[1]),  // code or desc as ID
        vendor: customer, material: r[4], po: project, deliveryDate: date,
        orderedQty: '', receivedQty: r[5],
      });
    });
    return out;
  }
  return [];
}

// Save a QC row. All 35 cols in schema order.
function addQcRecord(d) {
  try {
    _require(d, ['sourceType', 'overallOutcome']);
    const validOutcomes = ['Accept', 'Reject', 'Rework'];
    if (validOutcomes.indexOf(d.overallOutcome) < 0) throw new Error('Overall Outcome must be Accept / Reject / Rework');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const qc = ss.getSheetByName('QC Register');
    if (!qc) throw new Error('QC Register missing — run Initialize Workbook.');
    const user = currentUserEmail();
    const now = new Date();

    const row = [
      d.sourceType, d.refNo || '', d.refLine || '', d.overallOutcome,
      d.vendorName || '', d.materialName || '', d.poNo || '',
      d.deliveryDate ? parseDate(d.deliveryDate) : '', d.checkedByBasic || user,
      d.orderedQty || '', d.receivedQty || '', d.shortExcess || '',
      d.statusQty || '', d.checkedByQty || user,
      d.anyDamage || '', d.packagingCondition || '', d.leakageBreakage || '',
      d.properLabeling || '', d.statusVisual || '', d.checkedByVisual || user,
      d.sizePlanned || '', d.sizeActual || '', d.weightPlanned || '', d.weightActual || '',
      d.colourPlanned || '', d.colourActual || '', d.materialGradePlanned || '', d.materialGradeActual || '',
      d.statusSpec || '', d.checkedBySpec || user,
      d.worksAsExpected || '', d.trialDone || '', d.outputAcceptable || '', d.documentAttached || '',
      now,
    ];
    const startRow = nextRowByCol(qc, 1);
    qc.getRange(startRow, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'QC record saved (' + d.overallOutcome + ') for ' + (d.materialName || d.refLine || d.refNo) };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function _require(obj, keys) {
  const missing = keys.filter(function(k) { return !obj || obj[k] === '' || obj[k] === null || obj[k] === undefined; });
  if (missing.length) throw new Error('Missing required field(s): ' + missing.join(', '));
}

function _num(v, fallback) {
  if (v === '' || v === null || v === undefined) return fallback;
  const n = Number(v);
  if (isNaN(n)) throw new Error('Invalid number: ' + v);
  return n;
}

function _looksLikeEmail(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s));
}

function _assertGstFormat(gst) {
  // India GST: 15 chars, digits + capital letters. Loose check to avoid
  // blocking data entry when a supplier is exempt or unregistered.
  const g = String(gst).replace(/\s+/g, '');
  if (g.length !== 15) throw new Error('GST No must be 15 characters (got ' + g.length + ')');
  if (!/^[0-9A-Z]{15}$/.test(g)) throw new Error('GST No must be letters+digits, uppercase');
}

function _assertValueInDropdown(ss, colLetter, value, label) {
  const sh = ss.getSheetByName('Dropdowns');
  if (!sh) return; // if Dropdowns doesn't exist, skip validation
  const col = colLetter.charCodeAt(0) - 64;
  const vals = sh.getRange(2, col, Math.max(1, sh.getLastRow() - 1), 1).getValues().map(function(r) { return r[0]; });
  if (vals.indexOf(value) < 0) throw new Error(label + ' "' + value + '" not in allowed list');
}

function _assertMasterCodeExists(ss, sheetName, code, label) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh || !getRowByCode(sh, code)) throw new Error(label + ' code "' + code + '" not found in ' + sheetName);
}


// ============================================================================
// BULK EDIT — powers form-bulk-edit modal.
// bulkEditGetRecords(sheetName) → { headers, rows, editableCols, codeCol }
// bulkEditApply({sheetName, codes, field, value}) → { updated, skipped }
// ============================================================================

// Which cols are safely editable per master. Excludes formula cols (Description,
// Search Label, On Hand) and identity cols (Code col A). Also excludes CellImage
// Photo col (write Photo URL instead — CellImage rebuilds via bulk photo tool).
const BULK_EDITABLE = {
  'Item Master':       ['Item Name','Size','Grade/Color','UOM','Category','Material Type','Storage No.','Rack No.','HSN/SAC Code','Preferred Supplier Code','Avg Daily Consumption','Lead Time (Days)','Safety Factor','MIN LVL','MAX LVL','Active?','Avg Rate (Rs.)'],
  'SFG Master':        ['Item Name','Size','Color/Finish','UOM','Storage No.','MIN LVL','MAX LVL','Active?','Opening Qty'],
  'FG Master':         ['Item Name','Size','Colour','Finish','UOM','Active?','Opening Qty'],
  'Tool Master':       ['Item Name','Type','Serial No','Purchase Date','Purchase Value','Vendor','Current Status','Currently With','Opening Qty','Active?'],
  'Supplier Master':   ['Firm Name','Category','Contact Person','Mobile No','Email','GST No','Address','Active?'],
  'Job Worker Master': ['Firm Name','Process Type','Contact Person','Mobile No','Email','GST No','Address','Active?'],
};

function bulkEditGetRecords(sheetName) {
  if (!BULK_EDITABLE[sheetName]) throw new Error('Sheet not supported for bulk edit: ' + sheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet missing: ' + sheetName);
  const lastRow = nextRowByCol(sh, 1) - 1;
  if (lastRow < 2) return { headers: ['Code','Search Label'], rows: [], editableFields: BULK_EDITABLE[sheetName] };
  const labelCol = SEARCH_LABEL_COL[sheetName] || 2;
  const codes = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const labels = sh.getRange(2, labelCol, lastRow - 1, 1).getValues();
  const rows = [];
  for (let i = 0; i < codes.length; i++) {
    const c = codes[i][0];
    if (!c) continue;
    rows.push([c, labels[i][0] || c]);
  }
  return { headers: ['Code','Search Label'], rows: rows, editableFields: BULK_EDITABLE[sheetName] };
}

function bulkEditApply(d) {
  try {
    _require(d, ['sheetName', 'codes', 'field']);
    if (!Array.isArray(d.codes) || !d.codes.length) throw new Error('No records selected');
    if (!BULK_EDITABLE[d.sheetName] || BULK_EDITABLE[d.sheetName].indexOf(d.field) < 0) {
      throw new Error('Field not editable in bulk: ' + d.field);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(d.sheetName);
    if (!sh) throw new Error('Sheet missing: ' + d.sheetName);
    const col = _colByHeader(sh, d.field);
    if (col < 1) throw new Error('Column not found: ' + d.field);
    const lastRow = nextRowByCol(sh, 1) - 1;
    if (lastRow < 2) return { success: true, updated: 0, skipped: d.codes.length };
    const allCodes = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    const codeSet = {};
    d.codes.forEach(function(c) { codeSet[extractCodePrefix(c) || c] = true; });
    // Coerce value type: booleans + numbers pass through JSON; date strings parseDate
    let value = d.value;
    if (d.field.toLowerCase().indexOf('date') >= 0 && typeof value === 'string' && value) {
      try { value = parseDate(value); } catch (e) { /* leave as string */ }
    }
    let updated = 0;
    for (let i = 0; i < allCodes.length; i++) {
      const c = allCodes[i][0];
      if (c && codeSet[c]) {
        sh.getRange(i + 2, col).setValue(value);
        updated++;
      }
    }
    return { success: true, updated: updated, skipped: d.codes.length - updated,
             message: 'Updated ' + updated + ' record(s). ' + (d.codes.length - updated) + ' code(s) not found.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================================
// PO → GRN linkage helpers (v3.4).
// getPoDraftForInward(poNo) — reads PO Register + PO Lines, returns a draft
//   payload the Inward form pre-fills. Line rate/discount/GST carry over so
//   the user only has to confirm what was actually delivered.
// updatePoStatus(poNo) — recomputes 'Open' | 'Partial' | 'Received' by
//   comparing ordered qty vs delivered qty (sum across every Inward row that
//   references this PO). Also writes 'Received Qty' col in PO Lines.
// ============================================================================

// ============================================================================
// PO AMENDMENT — pick existing PO, edit header + lines, overwrite in place,
// bump Rev counter, regenerate PDF with -R{n} suffix. Original PDF stays in
// Drive as audit trail (orphaned link). Blocks if PO has any Inward (GRN) rows
// referencing it unless `d.forceAmend` is true (client prompts the user).
// ============================================================================

// Return full PO payload (header + lines) for pre-filling the Amend form.
function getPoForAmend(poNo) {
  try {
    if (!poNo) throw new Error('PO No required');
    const code = extractCodePrefix(poNo);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const po = ss.getSheetByName('Purchase Order Register');
    const linesSh = ss.getSheetByName('PO Lines');
    const inward = ss.getSheetByName('Inward Register');
    if (!po || !linesSh) throw new Error('PO sheets missing.');
    // PO Register cols: A=Sr(0), B=PO No(1), C=Date(2), D=Sup Code(3), E=Sup Name(4),
    //   F=Delivery Date(5), G=Status(6), H=Total(7), I=PDF(8), J=Remarks(9), M=Rev(12)
    let poRow = null, poRowIdx = -1;
    if (po.getLastRow() >= 2) {
      const pd = po.getRange(2, 1, po.getLastRow() - 1, Math.max(13, po.getLastColumn())).getValues();
      for (let i = 0; i < pd.length; i++) {
        if (pd[i][1] === code) { poRow = pd[i]; poRowIdx = i + 2; break; }
      }
    }
    if (!poRow) throw new Error('PO ' + code + ' not found');

    // Count GRN rows against this PO so the client can show a warning.
    let grnCount = 0;
    if (inward && inward.getLastRow() >= 2) {
      const poRefCol = _colByHeader(inward, 'PO Ref');
      if (poRefCol > 0) {
        const inw = inward.getRange(2, poRefCol, inward.getLastRow() - 1, 1).getValues();
        inw.forEach(function(r) { if (r[0] === code) grnCount++; });
      }
    }

    // PO Lines cols: A=PO No(0), B=Line(1), C=Code(2), D=Desc(3), E=Qty(4),
    //   F=UOM(5), G=Rate(6), H=Disc%(7), I=Disc Amt(8), J=GST%(9), K=GST Amt(10),
    //   L=Line Total(11), M=Received Qty(12)
    const lastLine = linesSh.getLastRow();
    const lineOut = [];
    if (lastLine >= 2) {
      const ld = linesSh.getRange(2, 1, lastLine - 1, 13).getValues();
      ld.forEach(function(r) {
        if (r[0] !== code) return;
        lineOut.push({
          itemCode: r[2], itemDesc: r[3], qty: Number(r[4]) || 0, uom: r[5] || '',
          rate: Number(r[6]) || 0, discount: Number(r[7]) || 0, gst: Number(r[9]) || 0,
          receivedQty: Number(r[12]) || 0,
        });
      });
    }
    return {
      success: true,
      poNo: code,
      supplierCode: poRow[3] || '',
      supplierName: poRow[4] || '',
      date: poRow[2] instanceof Date ? Utilities.formatDate(poRow[2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      deliveryDate: poRow[5] instanceof Date ? Utilities.formatDate(poRow[5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      status: poRow[6] || '',
      remarks: poRow[9] || '',
      currentRev: Number(poRow[12]) || 0,
      grnCount: grnCount,
      lines: lineOut,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Overwrite the PO header + lines with the amended data, bump Rev, regen PDF.
function amendPoRows(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const poReg = ss.getSheetByName('Purchase Order Register');
  const poLines = ss.getSheetByName('PO Lines');
  const items = ss.getSheetByName('Item Master');
  const suppliers = ss.getSheetByName('Supplier Master');
  const inward = ss.getSheetByName('Inward Register');
  if (!poReg || !poLines || !items || !suppliers) throw new Error('Missing sheet. Run Initialize Workbook.');
  const poNo = extractCodePrefix(formData.poNo);
  if (!poNo) throw new Error('PO No required');

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    // Find PO row
    const lastReg = poReg.getLastRow();
    if (lastReg < 2) throw new Error('PO Register empty');
    const regData = poReg.getRange(2, 1, lastReg - 1, Math.max(15, poReg.getLastColumn())).getValues();
    let regRowIdx = -1, origRow = null;
    for (let i = 0; i < regData.length; i++) {
      if (regData[i][1] === poNo) { regRowIdx = i + 2; origRow = regData[i]; break; }
    }
    if (regRowIdx < 0) throw new Error('PO ' + poNo + ' not found');
    const status = origRow[6] || '';
    if (status === 'Cancelled') throw new Error('PO ' + poNo + ' is Cancelled — cannot amend.');

    // Block if GRN rows exist (unless user forced)
    if (!formData.forceAmend && inward && inward.getLastRow() >= 2) {
      const poRefCol = _colByHeader(inward, 'PO Ref');
      if (poRefCol > 0) {
        const inw = inward.getRange(2, poRefCol, inward.getLastRow() - 1, 1).getValues();
        const grnCount = inw.filter(function(r) { return r[0] === poNo; }).length;
        if (grnCount > 0) {
          return { success: false, needsConfirm: true, grnCount: grnCount,
            message: 'PO ' + poNo + ' has ' + grnCount + ' GRN row(s) against it. Amending changes the ordered qty/rate — Received Qty tracking may become inconsistent. Confirm to proceed.' };
        }
      }
    }

    const supplierRow = getRowByCode(suppliers, origRow[3]);
    const supplierName = supplierRow ? supplierRow[1] : origRow[4];
    const dateObj = formData.date ? parseDate(formData.date) : origRow[2];
    const deliveryDate = formData.deliveryDate ? parseDate(formData.deliveryDate) : (origRow[5] || '');

    // Build new line rows
    const lineRows = [];
    let grandTotal = 0;
    const itemsMap = _buildCodeMap(items);
    formData.lineItems.forEach(function(line, idx) {
      const itemRow = itemsMap[line.itemCode];
      if (!itemRow) throw new Error('Item not found: ' + line.itemCode);
      const desc = itemRow[2] || '';
      const uom  = itemRow[6] || '';
      const { qty, rate } = _boardConvert(line);
      const gstPct  = Number(line.gst) || 0;
      const gross   = qty * rate;
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
      grandTotal += total;
      lineRows.push([poNo, idx + 1, line.itemCode, desc, qty, uom, rate, discPct, discAmt, gstPct, gstAmt, total, Number(line.receivedQty) || 0]);
    });

    // Delete old lines for this PO (bottom-up) — simpler + safer than overwrite
    const lastLines = poLines.getLastRow();
    if (lastLines >= 2) {
      const ld = poLines.getRange(2, 1, lastLines - 1, 1).getValues();
      for (let i = ld.length - 1; i >= 0; i--) {
        if (ld[i][0] === poNo) poLines.deleteRow(i + 2);
      }
    }
    // Append new lines
    if (lineRows.length) {
      const start = nextRowByCol(poLines, 1);
      poLines.getRange(start, 1, lineRows.length, lineRows[0].length).setValues(lineRows);
    }

    // Overwrite PO header row (keep Sr + PO No + Status + created-by/at, bump Rev)
    const newRev = (Number(origRow[12]) || 0) + 1;
    const user = currentUserEmail();
    const now = new Date();
    poReg.getRange(regRowIdx, 3).setValue(dateObj);                    // C
    poReg.getRange(regRowIdx, 4).setValue(origRow[3]);                 // D supplier code (unchanged)
    poReg.getRange(regRowIdx, 5).setValue(supplierName);               // E supplier name
    poReg.getRange(regRowIdx, 6).setValue(deliveryDate);               // F
    poReg.getRange(regRowIdx, 8).setValue(grandTotal);                 // H
    poReg.getRange(regRowIdx, 10).setValue(formData.remarks || '');    // J
    poReg.getRange(regRowIdx, 13).setValue(newRev);                    // M Rev No
    poReg.getRange(regRowIdx, 14).setValue(now);                       // N Last Amended At
    poReg.getRange(regRowIdx, 15).setValue(user);                      // O Last Amended By

    // Regenerate PDF — filename includes -R{rev} suffix via docNo trick.
    // We pass the real poNo for register lookup but the PDF file carries R-tag.
    let pdf = null;
    if (!formData.skipPdf) {
      const poPhotoMap = getMasterPhotoUrlMap('Item Master');
      const pdfLines = lineRows.map(function(r) {
        const im = itemsMap[r[2]];
        return { itemCode: r[2], itemDesc: r[3], qty: r[4], uom: r[5], rate: r[6],
                 discPct: r[7], gstPct: r[9], total: r[11],
                 hsn: im ? (im[11] || '') : '',
                 photoUrl: poPhotoMap[r[2]] || '' };
      });
      pdf = _safePdf(function() {
        return generatePoPdf(poNo + '-R' + newRev, {
          date: dateObj, supplierCode: origRow[3],
          supplierName: supplierName,
          supplierGstin: supplierRow ? (supplierRow[6] || '') : '',
          supplierAddress: supplierRow ? (supplierRow[7] || '') : '',
          supplierContact: supplierRow ? [supplierRow[3], supplierRow[4], supplierRow[5]].filter(String).join(' | ') : '',
          supplierMobile: supplierRow ? (supplierRow[4] || '') : '',
          supplierEmail:  supplierRow ? (supplierRow[5] || '') : '',
          deliveryDate: deliveryDate, remarks: formData.remarks,
        }, pdfLines);
      }, 'generatePoPdf', poNo + '-R' + newRev);
      if (pdf) poReg.getRange(regRowIdx, 9).setFormula('=HYPERLINK("' + pdf.url + '","View PDF (R' + newRev + ')")');
    }

    return { success: true, poNo: poNo, rev: newRev, rows: lineRows.length, total: grandTotal,
             pdfUrl: pdf ? pdf.url : null, pdfDownloadUrl: pdf ? pdf.downloadUrl : null };
  } finally {
    lock.releaseLock();
  }
}

function amendPo(d) {
  try {
    _require(d, ['poNo', 'lineItems']);
    if (!Array.isArray(d.lineItems) || !d.lineItems.length) throw new Error('At least one line item required');
    d.lineItems.forEach(function(line, i) {
      _require(line, ['itemCode', 'qty', 'rate']);
      if (Number(line.qty) <= 0)  throw new Error('Line ' + (i+1) + ': Qty must be > 0');
      if (Number(line.rate) < 0)  throw new Error('Line ' + (i+1) + ': Rate must be >= 0');
    });
    const res = amendPoRows(d);
    if (res && res.needsConfirm) return res;  // bubble up the GRN-exists confirm
    return {
      success: true,
      message: 'PO ' + res.poNo + ' amended (R' + res.rev + ', ' + res.rows + ' line(s), Rs. ' + res.total.toFixed(2) + ').',
      code: res.poNo + '-R' + res.rev,
      pdfDownloadUrl: res.pdfDownloadUrl,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getPoDraftForInward(poNo) {
  try {
    if (!poNo) throw new Error('PO No required');
    const code = extractCodePrefix(poNo);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const po = ss.getSheetByName('Purchase Order Register');
    const lines = ss.getSheetByName('PO Lines');
    if (!po || !lines) throw new Error('PO sheets missing.');
    // PO Register cols: A=Sr No, B=PO No, C=Date, D=Supplier Code, E=Supplier Name.
    // Codes live in col B, not col A — scan col B directly.
    let poRow = null;
    if (po.getLastRow() >= 2) {
      const pd = po.getRange(2, 1, po.getLastRow() - 1, po.getLastColumn()).getValues();
      for (let i = 0; i < pd.length; i++) {
        if (pd[i][1] === code) { poRow = pd[i]; break; }
      }
    }
    if (!poRow) throw new Error('PO ' + code + ' not found');
    const lastLine = nextRowByCol(lines, 1) - 1;
    if (lastLine < 2) return { success: true, supplierCode: poRow[3] || '', lines: [] };
    const data = lines.getRange(2, 1, lastLine - 1, 13).getValues();
    const out = [];
    data.forEach(function(r) {
      if (r[0] !== code) return;
      out.push({
        itemCode: r[2], itemDesc: r[3], qty: Number(r[4]) || 0, uom: r[5] || '',
        rate: Number(r[6]) || 0, discount: Number(r[7]) || 0, gst: Number(r[9]) || 0,
        receivedQty: Number(r[12]) || 0,
      });
    });
    return { success: true, poNo: code, supplierCode: poRow[3] || '', lines: out };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updatePoStatus(poNo) {
  const code = extractCodePrefix(poNo);
  if (!code) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const po = ss.getSheetByName('Purchase Order Register');
  const lines = ss.getSheetByName('PO Lines');
  const inward = ss.getSheetByName('Inward Register');
  if (!po || !lines || !inward) return;
  // Ordered qty by item, from PO Lines
  const lastLine = nextRowByCol(lines, 1) - 1;
  if (lastLine < 2) return;
  const lineData = lines.getRange(2, 1, lastLine - 1, 13).getValues();
  const ordered = {}; const rowByItem = {};
  lineData.forEach(function(r, i) {
    if (r[0] !== code) return;
    ordered[r[2]] = (ordered[r[2]] || 0) + (Number(r[4]) || 0);
    rowByItem[r[2]] = i + 2;  // sheet row
  });
  if (!Object.keys(ordered).length) return;
  // Delivered qty by item, from Inward Register where PO Ref = code
  const poRefCol = _colByHeader(inward, 'PO Ref');
  const lastInward = nextRowByCol(inward, 1) - 1;
  const delivered = {};
  if (poRefCol > 0 && lastInward >= 2) {
    const inw = inward.getRange(2, 1, lastInward - 1, poRefCol).getValues();
    inw.forEach(function(r) {
      if (r[poRefCol - 1] !== code) return;
      const item = r[3], qty = Number(r[7]) || 0;
      delivered[item] = (delivered[item] || 0) + qty;
    });
  }
  // Write Received Qty back to PO Lines + compute status
  let allComplete = true, anyDelivered = false;
  Object.keys(ordered).forEach(function(item) {
    const recv = delivered[item] || 0;
    lines.getRange(rowByItem[item], 13).setValue(recv);
    if (recv > 0) anyDelivered = true;
    if (recv < ordered[item]) allComplete = false;
  });
  const status = allComplete ? 'Received' : (anyDelivered ? 'Partial' : 'Open');
  // Locate PO row + write status (col G = 7). PO codes live in col B (2), not col A.
  const poRow = _findRowByCode(po, code, 2);
  if (poRow) po.getRange(poRow, 7).setValue(status);
}


// ============================================================================
// SITE RETURN — bring back surplus material from an earlier FG Delivery.
// Updates FG Delivery Lines col J (Returned Qty) in place, then refreshes the
// snapshot for the affected codes. The register / snapshot engines already
// treat `net = delivered − returned`, so no new register writes are needed.
//
// Multiple returns against the same DC line are additive (new return added to
// existing Returned Qty). Partial returns supported. Validation prevents
// returning more than Still Out.
// ============================================================================

// Called from form on DC pick — returns one row per DC line with current state.
function getDcLinesForReturn(dcNo) {
  try {
    if (!dcNo) throw new Error('DC No required');
    const code = extractCodePrefix(dcNo);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lines = ss.getSheetByName('FG Delivery Lines');
    if (!lines) throw new Error('FG Delivery Lines sheet missing');
    const lastData = nextRowByCol(lines, 1) - 1;
    if (lastData < 2) return { success: true, dcNo: code, lines: [] };
    const data = lines.getRange(2, 1, lastData - 1, 10).getValues();
    const out = [];
    data.forEach(function(r) {
      if (r[0] !== code) return;
      const qty = Number(r[5]) || 0;
      const returned = Number(r[9]) || 0;
      out.push({
        lineNo: r[1], section: r[2] || '', code: r[3] || '', desc: r[4] || '',
        qty: qty, uom: r[6] || '', returned: returned,
        stillOut: Math.max(0, qty - returned),
      });
    });
    return { success: true, dcNo: code, lines: out };
  } catch (e) {
    return { success: false, message: e.message, lines: [] };
  }
}

// Thin wrapper called from the form submit — validates then delegates.
function addSiteReturn(d) {
  try {
    _require(d, ['dcNo', 'date']);
    if (!Array.isArray(d.lineReturns) || !d.lineReturns.length) {
      throw new Error('No line returns provided');
    }
    const haveQty = d.lineReturns.some(function(l) { return Number(l.returnedQty) > 0; });
    if (!haveQty) throw new Error('Enter Returned Qty > 0 on at least one line');
    const res = postSiteReturn(d);
    return {
      success: true,
      message: 'Site return posted: ' + res.linesUpdated + ' line(s) restored on DC ' + res.dcNo + '. Stock updated.',
      code: res.dcNo,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Real worker — locks, validates against current Still Out, writes col J, refreshes snapshot.
function postSiteReturn(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lines = ss.getSheetByName('FG Delivery Lines');
  if (!lines) throw new Error('FG Delivery Lines sheet missing');
  const dcNo = extractCodePrefix(formData.dcNo);

  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const lastData = nextRowByCol(lines, 1) - 1;
    if (lastData < 2) throw new Error('DC not found: ' + dcNo);
    const data = lines.getRange(2, 1, lastData - 1, 10).getValues();

    // Index DC's lines by lineNo for lookup
    const byLineNo = {};
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === dcNo) byLineNo[String(data[i][1])] = { rowNum: i + 2, row: data[i] };
    }
    if (!Object.keys(byLineNo).length) throw new Error('DC not found: ' + dcNo);

    // Validate + plan updates. Fail BEFORE any write.
    const sfgSh = ss.getSheetByName('SFG Master');
    const rmCodes = [], sfgCodes = [];
    const updates = [];  // { rowNum, newReturned }
    formData.lineReturns.forEach(function(l) {
      const retQty = Number(l.returnedQty) || 0;
      if (retQty <= 0) return;
      const info = byLineNo[String(l.lineNo)];
      if (!info) throw new Error('Line #' + l.lineNo + ' not found on DC ' + dcNo);
      const qty = Number(info.row[5]) || 0;
      const existing = Number(info.row[9]) || 0;
      const stillOut = qty - existing;
      // Allow a 0.0001 slack for float noise on sqm-based returns
      if (retQty > stillOut + 0.0001) {
        throw new Error('Line #' + l.lineNo + ' (' + (info.row[3] || info.row[4]) + '): return qty ' +
                        retQty + ' exceeds still-out qty ' + stillOut);
      }
      updates.push({ rowNum: info.rowNum, newReturned: existing + retQty });
      const code = info.row[3];
      const section = info.row[2];
      if (!code) return;
      if (section === 'Hardware') {
        rmCodes.push(code);
      } else if (sfgSh && getRowByCode(sfgSh, code)) {
        sfgCodes.push(code);
      }
      // FG Main lines: FG Master's Available formula auto-recomputes from lines.
    });

    if (!updates.length) throw new Error('Nothing to update');

    // Apply updates one by one — each is a single-cell write.
    updates.forEach(function(u) { lines.getRange(u.rowNum, 10).setValue(u.newReturned); });

    // Refresh snapshots for touched RM / SFG codes.
    if (rmCodes.length)  updateItemSnapshotsBatch(Array.from(new Set(rmCodes)));
    if (sfgCodes.length) updateSfgSnapshotsBatch(Array.from(new Set(sfgCodes)));

    return { dcNo: dcNo, linesUpdated: updates.length };
  } finally {
    lock.releaseLock();
  }
}

function addFgDelivery(d) {
  try {
    _require(d, ['date']);
    if (!d.mainItems || !d.mainItems.length) throw new Error('At least one main item required');
    // Validate mainItems: must have description + qty > 0
    d.mainItems.forEach(function(l, i) {
      if (!String(l.description || '').trim()) throw new Error('Main item ' + (i + 1) + ': description required');
      if (!(Number(l.qty) > 0)) throw new Error('Main item ' + (i + 1) + ': Qty must be > 0');
    });
    (d.hardwareItems || []).forEach(function(l, i) {
      if (l.qty && Number(l.qty) > 0 && !String(l.description || '').trim())
        throw new Error('Hardware item ' + (i + 1) + ': description required');
    });
    // Overselling preflight for both Main + Hardware sections (only lines with a code)
    if (!d.bypassStockCheck) {
      const allLines = (d.mainItems || []).concat(d.hardwareItems || [])
        .filter(function(l) { return l.itemCode && Number(l.qty) > 0; })
        .map(function(l) { return { code: l.itemCode, qty: l.qty }; });
      if (allLines.length) {
        const check = _checkStockAvailability(allLines);
        if (!check.ok) return { success: false, needsConfirm: true, warnings: check.warnings, message: 'Some lines exceed Available stock. Confirm to proceed anyway.' };
      }
    }
    const res = postFgDeliveryRows(d);
    return { success: true, message: 'DC ' + res.dcNo + ' created (' + res.rows + ' lines).',
             code: res.dcNo, pdfDownloadUrl: res.pdfDownloadUrl, pdfUrl: res.pdfUrl };
  } catch (e) {
    return { success: false, message: e.message };
  }
}