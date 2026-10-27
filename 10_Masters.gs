/**
 * 3 Star Projects IMS v3 — 10_Masters.gs
 * Phase 1: code assignment, BOM helpers, onEdit auto-population for masters.
 *
 * Deps: 00_Config_Menu_Setup.gs (getConfig), 60_Utils_Audit_Triggers.gs (helpers).
 */

// ============================================================================
// CATEGORY → CODE PREFIX map
// Users can override any of these in the Dropdowns sheet by editing category
// names — code prefix falls back to first-3-uppercase-letters if a new
// category isn't in this map.
// ============================================================================

const CATEGORY_PREFIX = {
  'Fasteners & Connectors':    'FST',
  'Furniture Hardware':        'HRD',
  'Edge Banding':              'EGB',
  'Board Materials':           'BRD',
  'Surface Materials':         'SUR',
  'Fabric & Upholstery':       'FAB',
  'Foam & Cushioning':         'FOM',
  'Adhesives & Chemicals':     'ADH',
  'Installation Materials':    'INS',
  'Glass & Mirror':            'GLS',
  'Drawer Systems':            'DRW',
  'Wall & Ceiling Materials':  'WCM',
  'Wardrobe Accessories':      'WRD',
  'Sliding Systems':           'SLD',
  'Locks & Security Hardware': 'LCK',
  'Interior Decor Materials':  'DEC',
  'PVC & Plastic Materials':   'PVC',
  'Steel & Metal Sections':    'STL',
  'Machinery & Equipment':     'MCH',
  'Welding Materials':         'WLD',
  'Cutting Tools & Abrasives': 'CUT',
  'Power Tools':               'PWR',
  'Hand Tools':                'HND',
  'Tools & Equipment':         'TLS',
  'IT & Automation Materials': 'ITA',
  'Office Supplies':           'OFF',
  'Safety Materials':          'SFT',
  'Machine Consumables':       'MCN',
  'Electrical Materials':      'ELC',
  'Packaging Materials':       'PKG',
  'Repair & Service Items':    'RSI',
  'Other':                     'OTH',
};

function prefixForCategory(category) {
  if (!category) return 'OTH';
  if (CATEGORY_PREFIX[category]) return CATEGORY_PREFIX[category];
  // Fallback: first 3 alphanumeric chars uppercase
  const cleaned = String(category).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  return (cleaned.substr(0, 3) || 'OTH');
}

// ============================================================================
// CODE ASSIGNMENT — scans existing master for max counter per prefix, +1.
// Sheet is source of truth. Config counters (Phase 0) are used for GRN/REQ etc,
// NOT for master codes since master codes are per-prefix and scanning is safer.
// LockService prevents two simultaneous adds from colliding.
// ============================================================================

function assignItemCode(category) {
  const prefix = prefixForCategory(category);
  return _nextCodeForPrefix('Item Master', 1, prefix, 4);
}

function assignSfgCode()       { return _nextCodeForPrefix('SFG Master', 1, 'SFG', 4); }
function assignSupplierCode()  { return _nextCodeForPrefix('Supplier Master', 1, 'SUPP', 4); }
function assignJobWorkerCode() { return _nextCodeForPrefix('Job Worker Master', 1, 'JW', 4); }
function assignToolCode()      { return _nextCodeForPrefix('Tool Master', 1, 'TOOL', 4); }
function assignFgCode()        { return _nextCodeForPrefix('FG Master', 1, 'FG', 4); }

function _nextCodeForPrefix(sheetName, colNum, prefix, pad) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sh) throw new Error('Sheet not found: ' + sheetName);
    let max = 0;
    if (sh.getLastRow() >= 2) {
      const vals = sh.getRange(2, colNum, sh.getLastRow() - 1, 1).getValues();
      const re = new RegExp('^' + prefix + '-(\\d+)$');
      for (let i = 0; i < vals.length; i++) {
        const m = re.exec(String(vals[i][0] || ''));
        if (m) { const n = parseInt(m[1], 10); if (n > max) max = n; }
      }
    }
    return prefix + '-' + String(max + 1).padStart(pad, '0');
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// BOM helpers — used by Phase 5 Kitting viewer, but exposed here so BOM
// Master onEdit can auto-fill Component UOM.
// ============================================================================

function lookupComponentUom(componentType, componentCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = (componentType === 'RM') ? 'Item Master' : (componentType === 'SFG' ? 'SFG Master' : null);
  if (!sheetName) return '';
  const sh = ss.getSheetByName(sheetName);
  const row = getRowByCode(sh, componentCode);
  if (!row) return '';
  // v3.5: UOM col G (idx 6) in Item/SFG Master.
  return row[6] || '';
}

// ============================================================================
// onEdit ROUTER — auto-populate codes + UOMs when users edit master sheets.
// Registered by Google as a simple trigger (no explicit install needed).
// ============================================================================

function onEdit(e) {
  if (!e || !e.range) return;
  try {
    const sh = e.range.getSheet();
    const name = sh.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    // ---- Row 1 pickers (Explorer, Kitting) — must run BEFORE the row<2 skip ----
    if (row === 1) {
      if (name === 'Explorer') {
        if (col === 2) {
          // View changed — clear old D1 value + validation so the new view's
          // dropdown appears fresh, then rebuild with empty arg.
          const d1 = sh.getRange('D1');
          d1.clearDataValidations();
          d1.clearContent();
          const view = sh.getRange('B1').getValue();
          if (view) return rebuildExplorer(view, '');
        }
        if (col === 4) {
          const view = sh.getRange('B1').getValue();
          const arg  = sh.getRange('D1').getValue();
          if (view) return rebuildExplorer(view, arg);
        }
      }
      if (name === 'Kitting Sheet' && (col === 3 || col === 5)) {
        const fg = sh.getRange('C1').getValue();
        const q  = Number(sh.getRange('E1').getValue()) || 1;
        if (fg) return rebuildKittingSheet(fg, q);
      }
      return; // No other row-1 handlers
    }

    // ---- Row 2+ handlers (masters, registers with in-sheet edits) ----
    if (name === 'Item Master')          return _onEditItemMaster(sh, row, col, e);
    if (name === 'SFG Master')           return _onEditSfgMaster(sh, row, col, e);
    if (name === 'Supplier Master')      return _onEditSupplierMaster(sh, row, col, e);
    if (name === 'Job Worker Master')    return _onEditJobWorkerMaster(sh, row, col, e);
    if (name === 'Tool Master')          return _onEditToolMaster(sh, row, col, e);
    if (name === 'BOM Master')           return _onEditBomMaster(sh, row, col, e);
    if (name === 'Adjustments Register') return _onEditAdjustments(sh, row, col, e);
    if (name === 'Stock Snapshot')       return _onEditStockSnapshot(sh, row, col, e);
    if (name === 'FG Delivery Lines')    return _onEditFgDeliveryLines(sh, row, col, e);
  } catch (err) {
    Logger.log('onEdit error: ' + err.message + ' | ' + (err.stack || ''));
  }
}

// Item Master: if Category (col G, was F) filled and Item Code (col A) blank → assign.
function _onEditItemMaster(sh, row, col, e) {
  if (col !== 8) return; // Category col H (v3.5: Photo B pushed everything +1)
  const codeCell = sh.getRange(row, 1);
  if (codeCell.getValue()) return;
  const category = e.value;
  if (!category) return;
  codeCell.setValue(assignItemCode(category));
  // Default Active? col S = 19
  const activeCell = sh.getRange(row, 19);
  if (activeCell.getValue() === '') activeCell.setValue(true);
}

// SFG Master: trigger on Item Name (col C = 3) — Description (col B) is now formula-driven.
function _onEditSfgMaster(sh, row, col, e) {
  if (col !== 4) return; // Item Name col D (v3.5)
  const codeCell = sh.getRange(row, 1);
  if (codeCell.getValue()) return;
  if (!e.value) return;
  codeCell.setValue(assignSfgCode());
  const activeCell = sh.getRange(row, 11); // Active col K
  if (activeCell.getValue() === '') activeCell.setValue(true);
}

// Supplier Master: if Firm Name (col B) filled and Supplier Code (col A) blank → assign.
function _onEditSupplierMaster(sh, row, col, e) {
  if (col !== 2) return;
  const codeCell = sh.getRange(row, 1);
  if (codeCell.getValue()) return;
  if (!e.value) return;
  codeCell.setValue(assignSupplierCode());
  const activeCell = sh.getRange(row, 9);
  if (activeCell.getValue() === '') activeCell.setValue(true);
}

// Tool Master: trigger on Item Name (col C) — auto-assign TOOL-####.
// Default status = 'In Store', Active = TRUE.
function _onEditToolMaster(sh, row, col, e) {
  if (col !== 4) return; // Item Name col D (v3.5)
  const codeCell = sh.getRange(row, 1);
  if (codeCell.getValue()) return;
  if (!e.value) return;
  codeCell.setValue(assignToolCode());
  const statusCell = sh.getRange(row, 10); // Status col J
  if (!statusCell.getValue()) statusCell.setValue('In Store');
  const activeCell = sh.getRange(row, 14); // Active col N
  if (activeCell.getValue() === '') activeCell.setValue(true);
}

// Job Worker Master: if Firm Name (col B) filled and Code (col A) blank → assign.
function _onEditJobWorkerMaster(sh, row, col, e) {
  if (col !== 2) return;
  const codeCell = sh.getRange(row, 1);
  if (codeCell.getValue()) return;
  if (!e.value) return;
  codeCell.setValue(assignJobWorkerCode());
  const activeCell = sh.getRange(row, 9);
  if (activeCell.getValue() === '') activeCell.setValue(true);
}

// BOM Master: when Component Code (col E) selected, auto-fill UOM (col G).
function _onEditBomMaster(sh, row, col, e) {
  if (col !== 5) return; // only on Component Code
  const compType = sh.getRange(row, 4).getValue(); // col D
  const compCode = e.value;
  if (!compType || !compCode) return;
  const uom = lookupComponentUom(compType, compCode);
  if (uom) sh.getRange(row, 7).setValue(uom);
}

// Adjustments Register: when Item/SFG Code (col D) selected, auto-fill
// Description (col E) and UOM (col G). Handles composite Search Label values
// like "FST-0001 — PHILIP HEAD | 13 X 6 | WHITE" by extracting the code prefix.
function _onEditAdjustments(sh, row, col, e) {
  if (col !== 4) return;
  const raw = e.value;
  if (!raw) return;
  const code = extractCodePrefix(raw);
  // Rewrite the cell to just the code (so downstream lookups stay clean)
  if (code !== raw) sh.getRange(row, 4).setValue(code);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let master = ss.getSheetByName('Item Master');
  let masterRow = getRowByCode(master, code);
  if (!masterRow) {
    master = ss.getSheetByName('SFG Master');
    masterRow = getRowByCode(master, code);
  }
  if (!masterRow) return;
  sh.getRange(row, 5).setValue(masterRow[2] || ''); // Description (col C — formula, v3.5)
  sh.getRange(row, 7).setValue(masterRow[6] || ''); // UOM (col G in Item/SFG Master, v3.5)
}

// Stock Snapshot manual edit → auto-post Correction adjustment for the delta.
// Vertical layout: RM section at rows 2..(sfgHdr-2), SFG section at (sfgHdr+1)..end.
// Watches col I (RM On Hand) in RM section, col D (SFG On Hand) in SFG section.
// Other cols/rows are formulas/derived/headers — silently ignored.
function _onEditStockSnapshot(sh, row, col, e) {
  if (row < 2) return;
  const sfgHeaderRow = _findSfgSectionStart(sh);
  const inSfgSection = sfgHeaderRow > 0 && row > sfgHeaderRow;
  const inSpacerOrHeader = sfgHeaderRow > 0 && (row === sfgHeaderRow - 1 || row === sfgHeaderRow);
  if (inSpacerOrHeader) return;

  let kind;
  if (!inSfgSection && col === 9)      kind = 'RM';   // RM: col I = On Hand
  else if (inSfgSection && col === 4)  kind = 'SFG';  // SFG: col D = On Hand
  else return;

  const code = sh.getRange(row, 1).getValue();  // col A holds code in both sections
  if (!code) return;

  const newVal = Number(e.value);
  const oldVal = Number(e.oldValue);
  if (isNaN(newVal)) return;  // Invalid input (letters, blank clear, etc.) — skip
  const delta = newVal - (isNaN(oldVal) ? 0 : oldVal);
  if (delta === 0) return;    // No net change

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reg = ss.getSheetByName('Adjustments Register');
  const master = ss.getSheetByName(kind === 'RM' ? 'Item Master' : 'SFG Master');
  if (!reg || !master) return;
  const masterRow = getRowByCode(master, code);
  if (!masterRow) return;

  // Post Correction adjustment
  const user = currentUserEmail();
  const now = new Date();
  const startRow = nextRowByCol(reg, 1);
  reg.getRange(startRow, 1, 1, 11).setValues([[
    startRow - 1,
    now,
    'Correction',
    code,
    masterRow[2] || '',
    delta,
    masterRow[6] || '',
    'Manual edit in Stock Snapshot: ' + (isNaN(oldVal) ? 0 : oldVal) + ' → ' + newVal,
    '',
    user, now,
  ]]);

  // Refresh snapshot row (also updates SFG Master / FG Master On Hand formulas
  // indirectly by re-triggering the sheet's recalculation).
  if (kind === 'RM') updateItemSnapshotsBatch([code]);
  else               updateSfgSnapshotsBatch([code]);

  toast('Adjustment posted: ' + code + ' ' + (delta > 0 ? '+' : '') + delta + ' (was ' + (isNaN(oldVal) ? 0 : oldVal) + ', now ' + newVal + ')', '3 Star IMS', 4);
}

// FG Delivery Lines: a manual edit to col J (Returned Qty) is a site return
// shortcut — refresh the snapshot for that line's code so stock reflects it
// without needing Full Rebuild. Only col J is watched; other edits are ignored.
function _onEditFgDeliveryLines(sh, row, col, e) {
  if (col !== 10) return;  // Returned Qty col J only
  if (row < 2) return;
  const code = sh.getRange(row, 4).getValue();      // D = Item Code
  const section = sh.getRange(row, 3).getValue();   // C = Section
  if (!code) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (section === 'Hardware') {
    updateItemSnapshotsBatch([code]);
  } else {
    const sfgs = ss.getSheetByName('SFG Master');
    if (sfgs && getRowByCode(sfgs, code)) updateSfgSnapshotsBatch([code]);
    // FG Main lines auto-recompute via FG Master's Available formula.
  }
  toast('Snapshot refreshed for ' + code + ' (DC return edit)', '3 Star IMS', 3);
}