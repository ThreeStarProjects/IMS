/**
 * 80_Formatting.gs — one-shot visual polish for the whole workbook.
 *
 * Everything is idempotent: safe to run any number of times. If a sheet is
 * missing (older/newer workbook), that block is silently skipped.
 *
 * Menu: 3 Star IMS → Setup → Apply Visual Formatting
 */

// ============================================================================
// PALETTE — sheet tab groups
// ============================================================================
const TAB_COLORS = {
  master:   '#1F4E79',  // deep blue
  register: '#1E8449',  // green
  report:   '#F47C2C',  // orange
  system:   '#7F8C8D',  // grey
};

const SHEETS_BY_GROUP = {
  master: ['Item Master', 'SFG Master', 'FG Master', 'Supplier Master', 'Job Worker Master', 'BOM Master', 'Tool Master'],
  register: [
    'Inward Register', 'Material Requisition Register', 'Job Work Outward Register',
    'Job Work Inward Register', 'FG Delivery Register', 'FG Delivery Lines',
    'Adjustments Register', 'QC Register', 'Reservation Register',
    'Purchase Order Register', 'PO Lines', 'Tool Movement Register',
  ],
  report: ['Dashboard', 'Stock Snapshot', 'Explorer', 'Kitting Sheet'],
  system: ['Dropdowns', 'Config'],
};

// ============================================================================
// MASTER ENTRY POINT
// ============================================================================
function applyVisualFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _colorTabs(ss);
  _reorderTabs(ss);
  _applyGlobalFreezePanes(ss);
  _applyNumberFormats(ss);
  _applyConditionalFormatting(ss);
  _applyBanding(ss);
  _applyColumnWidths(ss);
  toast('Visual formatting applied to all sheets.');
}

// ============================================================================
// TAB COLORS + ORDER
// ============================================================================
function _colorTabs(ss) {
  Object.keys(SHEETS_BY_GROUP).forEach(function(group) {
    SHEETS_BY_GROUP[group].forEach(function(name) {
      const sh = ss.getSheetByName(name);
      if (sh) sh.setTabColor(TAB_COLORS[group]);
    });
  });
}

// Left-to-right order: Dashboard, Snapshot, Explorer, Masters..., Registers..., System
function _reorderTabs(ss) {
  const order = []
    .concat(SHEETS_BY_GROUP.report)     // Dashboard first
    .concat(SHEETS_BY_GROUP.master)
    .concat(SHEETS_BY_GROUP.register)
    .concat(SHEETS_BY_GROUP.system);
  let pos = 1;
  order.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try { ss.setActiveSheet(sh); ss.moveActiveSheet(pos++); } catch (e) {}
  });
}

// ============================================================================
// FREEZE PANES
// ============================================================================
function _applyGlobalFreezePanes(ss) {
  // Freeze header row 1 on every sheet
  ss.getSheets().forEach(function(sh) {
    try { sh.setFrozenRows(1); } catch (e) {}
  });
  // Masters: also freeze col A (code) + col B (photo)
  SHEETS_BY_GROUP.master.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (sh) try { sh.setFrozenColumns(2); } catch (e) {}
  });
  // Registers: freeze col A (Sr No)
  SHEETS_BY_GROUP.register.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (sh) try { sh.setFrozenColumns(1); } catch (e) {}
  });
}

// ============================================================================
// NUMBER FORMATS — currency, quantity, dates
// ============================================================================
const FMT_QTY   = '#,##0.00;[Red]-#,##0.00;""';
const FMT_MONEY = '₹#,##0.00;[Red]-₹#,##0.00;""';
const FMT_INT   = '#,##0;[Red]-#,##0;""';
const FMT_DATE  = 'dd-mmm-yyyy';
const FMT_TS    = 'dd-mmm-yyyy hh:mm';

// (sheetName, cols[], format) triples — one per register/master
const NUMBER_FORMAT_RULES = [
  ['Inward Register',            ['H','J','L','N','O','P','T'], null],  // Handled below
  ['Material Requisition Register', ['G','O'], null],
  ['Job Work Outward Register',  ['I','K','P','Q'], null],
  ['Job Work Inward Register',   ['H','J','K','O'], null],
  ['FG Delivery Register',       ['P','Q','T'], null],
  ['FG Delivery Lines',          ['F','H','J'], null],
  ['Adjustments Register',       ['B','F','K'], null],
  ['Reservation Register',       ['C','F','L','M'], null],
  ['Purchase Order Register',    ['C','F','H','L'], null],
  ['PO Lines',                   ['E','G','I','K','L','M'], null],
  ['Stock Snapshot',             ['E','F','G','H','I','J','K','L','M','N','O','P'], null],
  ['Tool Movement Register',     ['C','I','J','O'], null],
  ['QC Register',                ['H','J','K','L','AI'], null],
];

function _applyNumberFormats(ss) {
  // Inward Register: qty(H)=qty, rate(J)=money, disc%(K)=pct, disc(L)=money, gst%(M)=pct, gst(N)=money, total(O)=money, date(C)=date
  // ponytail: '0.00%' multiplies the stored value by 100 (30 → "3000%");
  // we store raw percent numbers, so use a literal "%" suffix instead.
  _fmtCols(ss, 'Inward Register', {
    'C': FMT_DATE, 'H': FMT_QTY, 'J': FMT_MONEY, 'K': '0.00"%"', 'L': FMT_MONEY,
    'M': '0.00"%"', 'N': FMT_MONEY, 'O': FMT_MONEY, 'T': FMT_TS
  });
  _fmtCols(ss, 'Material Requisition Register', { 'C': FMT_DATE, 'G': FMT_QTY, 'O': FMT_TS });
  _fmtCols(ss, 'Job Work Outward Register', { 'C': FMT_DATE, 'I': FMT_QTY, 'K': FMT_DATE, 'P': FMT_TS, 'Q': FMT_QTY });
  _fmtCols(ss, 'Job Work Inward Register', { 'C': FMT_DATE, 'H': FMT_QTY, 'J': FMT_QTY, 'K': FMT_QTY, 'O': FMT_TS });
  _fmtCols(ss, 'FG Delivery Register', { 'C': FMT_DATE, 'P': FMT_QTY, 'Q': FMT_QTY, 'T': FMT_TS });
  _fmtCols(ss, 'FG Delivery Lines', { 'F': FMT_QTY, 'H': FMT_QTY, 'J': FMT_QTY });
  _fmtCols(ss, 'Adjustments Register', { 'B': FMT_DATE, 'F': FMT_QTY, 'K': FMT_TS });
  _fmtCols(ss, 'Reservation Register', { 'C': FMT_DATE, 'F': FMT_QTY, 'L': FMT_TS, 'M': FMT_TS });
  _fmtCols(ss, 'Purchase Order Register', { 'C': FMT_DATE, 'F': FMT_DATE, 'H': FMT_MONEY, 'L': FMT_TS });
  _fmtCols(ss, 'PO Lines', { 'E': FMT_QTY, 'G': FMT_MONEY, 'I': FMT_MONEY, 'K': FMT_MONEY, 'L': FMT_MONEY, 'M': FMT_QTY });
  // Stock Snapshot has two sections stacked vertically (RM cols A:Q, SFG cols A:I with
  // different semantics). Apply RM formats to the RM section only; then apply SFG formats
  // to the SFG section. Prevents SFG col I ("Below MIN?" bool) rendering as 0.00/1.00.
  _fmtStockSnapshot(ss);
  _fmtCols(ss, 'Tool Movement Register', { 'C': FMT_DATE, 'I': FMT_DATE, 'J': FMT_DATE, 'O': FMT_TS });
  _fmtCols(ss, 'QC Register', { 'H': FMT_DATE, 'J': FMT_QTY, 'K': FMT_QTY, 'L': FMT_QTY, 'AI': FMT_TS });
  // Item/SFG/FG Master — MIN/MAX/Opening as qty
  _fmtCols(ss, 'Item Master',  { 'N': FMT_QTY, 'O': FMT_INT, 'P': FMT_QTY, 'Q': FMT_QTY, 'R': FMT_QTY });
  _fmtCols(ss, 'SFG Master',   { 'I': FMT_QTY, 'J': FMT_QTY, 'L': FMT_QTY, 'M': FMT_QTY });
  _fmtCols(ss, 'FG Master',    { 'J': FMT_QTY, 'K': FMT_QTY });
  _fmtCols(ss, 'Tool Master',  { 'G': FMT_DATE, 'H': FMT_MONEY, 'L': FMT_QTY, 'M': FMT_QTY });
}

function _fmtStockSnapshot(ss) {
  const sh = ss.getSheetByName('Stock Snapshot');
  if (!sh) return;
  const maxRow = Math.max(sh.getMaxRows(), 2);
  let sfgHeaderRow = 0;
  try { sfgHeaderRow = _findSfgSectionStart(sh); } catch (e) {}
  const rmEnd = sfgHeaderRow > 0 ? sfgHeaderRow - 2 : maxRow;
  // ── RM section (rows 2..rmEnd) ──
  const rmFmts = {
    'E': FMT_QTY, 'F': FMT_QTY, 'G': FMT_QTY, 'H': FMT_QTY, 'I': FMT_QTY, 'J': FMT_QTY,
    'K': FMT_QTY, 'L': FMT_QTY, 'M': FMT_MONEY, 'N': FMT_DATE, 'O': FMT_DATE, 'P': FMT_INT
  };
  if (rmEnd >= 2) {
    // Nuke stuck legacy per-cell formats first (MAX LVL stuck as currency,
    // Stock Value stuck as Date, etc.). Then apply the correct format per col.
    try { sh.getRange('A2:Q' + rmEnd).setNumberFormat('0.############'); } catch (e) {}
    Object.keys(rmFmts).forEach(function(col) {
      try { sh.getRange(col + '2:' + col + rmEnd).setNumberFormat(rmFmts[col]); } catch (e) {}
    });
    // ponytail: Storage/Rack (col D) must stay TEXT. If a user types "20" or
    // "46132" as a storage/rack id, Sheets otherwise coerces it to a number and
    // the numeric format from the sweep above makes it render as "20.00" or
    // "46,132.00". '@' format forces text so it displays verbatim.
    try { sh.getRange('D2:D' + rmEnd).setNumberFormat('@'); } catch (e) {}
    // Force col M to a fresh numeric value — if any cell is stuck as a Date-typed
    // value (getValues returns Date obj), read + rewrite as pure number so the
    // dashboard's Number() coercion works. Idempotent for already-numeric cells.
    try {
      const rng = sh.getRange('M2:M' + rmEnd);
      const cur = rng.getValues();
      const fixed = cur.map(function(r) {
        const v = r[0];
        if (v instanceof Date) return [0];          // date object → 0 (real value re-written by next rebuild)
        return [(typeof v === 'number' && isFinite(v)) ? v : (v === '' || v == null ? '' : Number(v) || 0)];
      });
      rng.setValues(fixed);
    } catch (e) {}
  }
  // ── SFG section (rows sfgHeader+1..maxRow) ──
  // SFG cols: A=Code, B=Desc, C=Size/Color, D=OnHand, E=InProcess, F=Available,
  //           G=MIN, H=MAX, I=Below MIN? (bool — leave as default/text)
  if (sfgHeaderRow > 0 && maxRow > sfgHeaderRow) {
    const sfgStart = sfgHeaderRow + 1;
    ['D','E','F','G','H'].forEach(function(col) {
      try { sh.getRange(col + sfgStart + ':' + col + maxRow).setNumberFormat(FMT_QTY); } catch (e) {}
    });
    // Clear any format leakage in col I of SFG section back to auto
    try { sh.getRange('I' + sfgStart + ':I' + maxRow).setNumberFormat('@'); } catch (e) {}
  }
}

function _fmtCols(ss, sheetName, colFmts) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const lastRow = Math.max(sh.getMaxRows(), 2);
  Object.keys(colFmts).forEach(function(col) {
    try {
      sh.getRange(col + '2:' + col + lastRow).setNumberFormat(colFmts[col]);
    } catch (e) {}
  });
}

// ============================================================================
// CONDITIONAL FORMATTING — the "wow" factor
// ============================================================================
function _applyConditionalFormatting(ss) {
  _cfStockSnapshot(ss);
  _cfAdjustments(ss);
  _cfReservations(ss);
  _cfJwOutward(ss);
  _cfFgDelivery(ss);
  _cfQcRegister(ss);
  _cfMasters(ss);
}

function _cfStockSnapshot(ss) {
  const sh = ss.getSheetByName('Stock Snapshot');
  if (!sh) return;
  sh.clearConditionalFormatRules();
  const maxRow = Math.max(sh.getMaxRows(), 100);
  const rules = [];
  // Below MIN (col Q) TRUE → whole row red tint
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$Q2=TRUE')
    .setBackground('#FADBD8')
    .setRanges([sh.getRange('A2:Q' + maxRow)]).build());
  // Available (I) <= MIN × 1.5 (amber warn zone)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($I2<>"",$K2<>"",$I2<=$K2*1.5,$Q2<>TRUE)')
    .setBackground('#FDEBD0')
    .setRanges([sh.getRange('A2:Q' + maxRow)]).build());
  // Days since movement > 180 → red text; > 90 → orange text
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(180)
    .setFontColor('#C0392B').setBold(true)
    .setRanges([sh.getRange('P2:P' + maxRow)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(91, 180)
    .setFontColor('#F47C2C')
    .setRanges([sh.getRange('P2:P' + maxRow)]).build());
  // Negative Available → red bold
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#C0392B').setBold(true)
    .setRanges([sh.getRange('I2:I' + maxRow)]).build());
  // Reservations > 0 → light blue tint
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#D6EAF8')
    .setRanges([sh.getRange('J2:J' + maxRow)]).build());
  sh.setConditionalFormatRules(rules);
}

function _cfAdjustments(ss) {
  const sh = ss.getSheetByName('Adjustments Register');
  if (!sh) return;
  sh.clearConditionalFormatRules();
  const maxRow = Math.max(sh.getMaxRows(), 100);
  const rules = [
    // Type=Damage → light red row
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="Damage"')
      .setBackground('#FADBD8')
      .setRanges([sh.getRange('A2:K' + maxRow)]).build(),
    // Type=Opening → light green (informational, one-time)
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="Opening"')
      .setBackground('#D5F5E3')
      .setRanges([sh.getRange('A2:K' + maxRow)]).build(),
    // Negative qty adjustments → red text on qty col F
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor('#C0392B').setBold(true)
      .setRanges([sh.getRange('F2:F' + maxRow)]).build(),
  ];
  sh.setConditionalFormatRules(rules);
}

function _cfReservations(ss) {
  const sh = ss.getSheetByName('Reservation Register');
  if (!sh) return;
  sh.clearConditionalFormatRules();
  const maxRow = Math.max(sh.getMaxRows(), 100);
  const rules = [
    // Active → light blue
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I2="Active"')
      .setBackground('#D6EAF8')
      .setRanges([sh.getRange('A2:N' + maxRow)]).build(),
    // Released → grey
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I2="Released"')
      .setBackground('#EAEDED').setFontColor('#7F8C8D')
      .setRanges([sh.getRange('A2:N' + maxRow)]).build(),
    // Consumed → light green
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I2="Consumed"')
      .setBackground('#D5F5E3')
      .setRanges([sh.getRange('A2:N' + maxRow)]).build(),
  ];
  sh.setConditionalFormatRules(rules);
}

function _cfJwOutward(ss) {
  const sh = ss.getSheetByName('Job Work Outward Register');
  if (!sh) return;
  sh.clearConditionalFormatRules();
  const maxRow = Math.max(sh.getMaxRows(), 100);
  const rules = [
    // Status=Open + date > 21 days → red (overdue)
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($N2="Open",TODAY()-$C2>21)')
      .setBackground('#FADBD8')
      .setRanges([sh.getRange('A2:Q' + maxRow)]).build(),
    // Status=Open (not overdue) → soft amber
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$N2="Open"')
      .setBackground('#FDEBD0')
      .setRanges([sh.getRange('A2:Q' + maxRow)]).build(),
    // Status=Closed → green tint
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$N2="Closed"')
      .setBackground('#D5F5E3')
      .setRanges([sh.getRange('A2:Q' + maxRow)]).build(),
  ];
  sh.setConditionalFormatRules(rules);
}

function _cfFgDelivery(ss) {
  const sh = ss.getSheetByName('FG Delivery Register');
  if (!sh) return;
  sh.clearConditionalFormatRules();
  const maxRow = Math.max(sh.getMaxRows(), 100);
  const rules = [
    // Type=Return → light red
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$U2="Return"')
      .setBackground('#FADBD8')
      .setRanges([sh.getRange('A2:U' + maxRow)]).build(),
  ];
  sh.setConditionalFormatRules(rules);
}

function _cfQcRegister(ss) {
  const sh = ss.getSheetByName('QC Register');
  if (!sh) return;
  sh.clearConditionalFormatRules();
  const maxRow = Math.max(sh.getMaxRows(), 100);
  const rules = [
    // Overall Outcome = Pass (col D) → green
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="Pass"')
      .setBackground('#D5F5E3')
      .setRanges([sh.getRange('A2:AI' + maxRow)]).build(),
    // Fail → red
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="Fail"')
      .setBackground('#FADBD8')
      .setRanges([sh.getRange('A2:AI' + maxRow)]).build(),
    // Rework / Conditional → amber
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($D2="Rework",$D2="Conditional")')
      .setBackground('#FDEBD0')
      .setRanges([sh.getRange('A2:AI' + maxRow)]).build(),
  ];
  sh.setConditionalFormatRules(rules);
}

function _cfMasters(ss) {
  // Inactive rows on all masters → grey out
  const masterActiveCols = {
    'Item Master': 'S', 'SFG Master': 'K', 'FG Master': 'I',
    'Supplier Master': 'I', 'Job Worker Master': 'I', 'Tool Master': 'N',
  };
  Object.keys(masterActiveCols).forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const col = masterActiveCols[name];
    const maxRow = Math.max(sh.getMaxRows(), 100);
    const lastCol = String.fromCharCode(64 + sh.getLastColumn());  // rough letter
    sh.clearConditionalFormatRules();
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$' + col + '2=FALSE')
        .setBackground('#EAEDED').setFontColor('#7F8C8D').setItalic(true)
        .setRanges([sh.getRange('A2:' + lastCol + maxRow)]).build(),
    ]);
  });
}

// ============================================================================
// ROW BANDING — subtle alternating row colors (readability)
// ============================================================================
function _applyBanding(ss) {
  const bandTargets = []
    .concat(SHEETS_BY_GROUP.master)
    .concat(SHEETS_BY_GROUP.register);
  bandTargets.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    // Remove existing bandings first (idempotency)
    sh.getBandings().forEach(function(b) { try { b.remove(); } catch(e){} });
    try {
      const range = sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), sh.getLastColumn());
      range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    } catch (e) {}
  });
}

// ============================================================================
// COLUMN WIDTHS — standardized so nothing looks cramped
// ============================================================================
function _applyColumnWidths(ss) {
  // Masters: A=code(110), B=photo(80), C=desc wide(320)
  SHEETS_BY_GROUP.master.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try {
      sh.setColumnWidth(1, 110);
      if (name !== 'Supplier Master' && name !== 'Job Worker Master' && name !== 'BOM Master') {
        sh.setColumnWidth(2, 80);   // Photo column
        sh.setColumnWidth(3, 320);  // Description
      }
    } catch (e) {}
  });
  // Registers: A=Sr(50), B=DocNo(120), C=Date(100)
  SHEETS_BY_GROUP.register.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try {
      sh.setColumnWidth(1, 50);
      if (name === 'FG Delivery Lines' || name === 'PO Lines') return;  // no doc-no col
      sh.setColumnWidth(2, 120);
      sh.setColumnWidth(3, 100);
    } catch (e) {}
  });
  // Stock Snapshot: item code, desc wide, then narrow qty cols
  const snap = ss.getSheetByName('Stock Snapshot');
  if (snap) {
    try {
      snap.setColumnWidth(1, 110);
      snap.setColumnWidth(2, 280);
      snap.setColumnWidth(3, 130);
      snap.setColumnWidth(4, 100);
      for (let c = 5; c <= 12; c++) snap.setColumnWidth(c, 90);
      snap.setColumnWidth(13, 130);
      snap.setColumnWidth(14, 105);
      snap.setColumnWidth(15, 105);
      snap.setColumnWidth(16, 80);
      snap.setColumnWidth(17, 90);
    } catch (e) {}
  }
}