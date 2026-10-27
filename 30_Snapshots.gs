/**
 * 3 Star Projects IMS v3 — 30_Snapshots.gs
 * Stock Snapshot compute + refresh. Explorer/Kitting rebuilds also here.
 *
 * Layout — vertical stack (single tab):
 *   Row 1              : RM header (cols A:P from SHEET_SPECS)
 *   Rows 2..N          : RM data
 *   Row N+1            : blank spacer
 *   Row N+2            : SFG header (SFG_SNAPSHOT_HEADERS, cols A:I)
 *   Rows N+3..M        : SFG data
 *
 * All SFG readers locate the section via _findSfgSectionStart() which scans
 * col A for the header text 'SFG Code'.
 *
 * Deps: 00_Config_Menu_Setup.gs, 60_Utils_Audit_Triggers.gs.
 */

// SFG section headers — written dynamically at (rmRows + 3) by fullRebuild.
const SFG_SNAPSHOT_HEADERS = ['SFG Code','SFG Description','Size/Color','On Hand (SFG)','In Process (SFG)','Available (SFG)','MIN LVL','MAX LVL','Below MIN?'];

// Returns the sheet row of the SFG header ('SFG Code' in col A), or -1 if
// not yet written. SFG data starts at (returnedRow + 1).
function _findSfgSectionStart(snap) {
  if (!snap || snap.getLastRow() < 2) return -1;
  const col = snap.getRange(1, 1, snap.getLastRow(), 1).getValues();
  for (let i = 0; i < col.length; i++) {
    if (col[i][0] === 'SFG Code') return i + 1;  // 1-based row
  }
  return -1;
}

// Read just the RM section rows (2..sfgHeader-2) with N cols. Returns [] if
// no RM data.
function _readRmSection(snap, numCols) {
  if (!snap || snap.getLastRow() < 2) return [];
  const sfgHeaderRow = _findSfgSectionStart(snap);
  const rmEnd = sfgHeaderRow > 0 ? sfgHeaderRow - 2 : snap.getLastRow();
  if (rmEnd < 2) return [];
  return snap.getRange(2, 1, rmEnd - 1, numCols || 17).getValues();
}

// Read just the SFG section rows ((sfgHeader+1)..lastRow) with N cols. Returns [] if
// no SFG data.
function _readSfgSection(snap, numCols) {
  if (!snap) return [];
  const sfgHeaderRow = _findSfgSectionStart(snap);
  if (sfgHeaderRow < 0) return [];
  const lastRow = snap.getLastRow();
  if (lastRow <= sfgHeaderRow) return [];
  return snap.getRange(sfgHeaderRow + 1, 1, lastRow - sfgHeaderRow, numCols || 9).getValues();
}

// ============================================================================
// FULL REBUILD — from all registers, called nightly + from menu.
// Reads every register once, aggregates into maps keyed by code, then
// writes both sections in one setValues() per section (fast).
// ============================================================================

function fullRebuild() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snap = ss.getSheetByName('Stock Snapshot');
  if (!snap) throw new Error('Stock Snapshot sheet missing — run Initialize Workbook.');

  const items = _readMaster(ss, 'Item Master');
  const sfgs  = _readMaster(ss, 'SFG Master');

  // Aggregate registers ONCE. Share adjustments + reservations across RM/SFG
  // paths — previously both branches read those sheets separately.
  const sharedAdj = _aggregateAdjustments(ss);
  const sharedRes = _aggregateReservations(ss);
  const rmMaps  = _readAllRmAggregates(ss, { adjAll: sharedAdj, reserved: sharedRes });
  const sfgMaps = _readAllSfgAggregates(ss, { adjAll: sharedAdj, reserved: sharedRes });

  // Clear ALL existing data + formatting (keep row 1 RM header intact).
  const lastRow = snap.getLastRow();
  const lastCol = Math.max(snap.getLastColumn(), 16);
  if (lastRow > 1) snap.getRange(2, 1, lastRow - 1, lastCol).clear();

  const rmRows  = items.map(function(it) { return _computeRmRowFromMaps(it, rmMaps); });
  const sfgRows = sfgs.map(function(sf) { return _computeSfgRowFromMaps(sf, sfgMaps); });

  // RM section at rows 2..N (cols A:P)
  if (rmRows.length) snap.getRange(2, 1, rmRows.length, 17).setValues(rmRows);

  // SFG section: 1 spacer row after RM, then header, then data (cols A:I)
  const sfgHeaderRow = 2 + rmRows.length + 1;  // +1 spacer
  snap.getRange(sfgHeaderRow, 1, 1, SFG_SNAPSHOT_HEADERS.length)
      .setValues([SFG_SNAPSHOT_HEADERS])
      .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  if (sfgRows.length) snap.getRange(sfgHeaderRow + 1, 1, sfgRows.length, 9).setValues(sfgRows);

  // Force-reapply column formats so MIN/MAX/Stock Value never drift wrong.
  try { _fmtStockSnapshot(ss); } catch (e) {}
  // Dashboard reads from Snapshot — refresh it too so numbers stay in sync.
  try { rebuildDashboard(); } catch (e) {}
  toast('Stock rebuilt: ' + rmRows.length + ' items, ' + sfgRows.length + ' SFGs.', '3 Star IMS', 4);
  return { items: rmRows.length, sfgs: sfgRows.length };
}

// ============================================================================
// BATCH SNAPSHOT UPDATE — reads every register ONCE and updates N codes.
// Called from post* handlers instead of looping updateItemSnapshot per code.
// For 100 unique codes: 5 sheet reads instead of 500. 20-40× faster on multi-
// line transactions.
// ============================================================================

// Read all registers once, return aggregation maps for RM computations.
// Optional `shared` param — passes pre-computed adjAll + reserved maps to
// avoid double-reading those sheets when both RM and SFG aggregates are needed
// in the same call (fullRebuild). Standalone callers pass nothing → same behavior.
function _readAllRmAggregates(ss, shared) {
  shared = shared || {};
  return {
    inward:   _aggregate(ss, 'Inward Register',                { codeCol: 4, dateCol: 3, qtyCol: 8, rateCol: 10 }),
    req:      _aggregate(ss, 'Material Requisition Register',  { codeCol: 5, dateCol: 3, qtyCol: 7 }),
    adjAll:   shared.adjAll   || _aggregateAdjustments(ss),
    fgDelHw:  _aggregateFgDelHardware(ss),
    reserved: shared.reserved || _aggregateReservations(ss),
    // ponytail: RM may be sent to a Job Worker (Change 2: items→JW). The JW
    // Outward / Inward registers contain mixed SFG + RM codes; the compute
    // function filters by it.code so SFG entries silently miss on RM lookups.
    // Net (jwOut − jwIn) is "currently out at JW" — deducted from physical.
    jwOut:    shared.jwOut || _aggregate(ss, 'Job Work Outward Register', { codeCol: 7, dateCol: 3, qtyCol: 9 }),
    jwIn:     shared.jwIn  || _aggregate(ss, 'Job Work Inward Register',  { codeCol: 6, dateCol: 3, qtyCol: 8 }),
  };
}

// Read all registers once, return aggregation maps for SFG computations.
// fgDelMain accounts for SFG codes shipped via FG Delivery (main section) —
// these are SFG outward and must reduce SFG on-hand.
function _readAllSfgAggregates(ss, shared) {
  shared = shared || {};
  return {
    jwOut:  _aggregate(ss, 'Job Work Outward Register', { codeCol: 7, dateCol: 3, qtyCol: 9, statusCol: 14 }),
    jwIn:   _aggregate(ss, 'Job Work Inward Register',  { codeCol: 6, dateCol: 3, qtyCol: 8 }),
    adjAll: shared.adjAll     || _aggregateAdjustments(ss),
    fgDelMain: _aggregateFgDelMain(ss),
    reserved: shared.reserved || _aggregateReservations(ss),
  };
}

// Sum ACTIVE reservations by code. Returns { code: totalReservedQty }.
// Only Status='Active' counts — Released and Consumed are ignored.
function _aggregateReservations(ss) {
  const sh = ss.getSheetByName('Reservation Register');
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  // Cols: A=Sr, B=ResNo, C=Date, D=Code, E=Desc, F=Qty, G=UOM, H=For, I=Status
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  data.forEach(function(r) {
    if (!r[3]) return;
    if (r[8] !== 'Active') return;
    const qty = Number(r[5]) || 0;
    if (qty <= 0) return;
    out[r[3]] = (out[r[3]] || 0) + qty;
  });
  return out;
}

// Aggregate FG Delivery Lines Main section (any code — matched later by code).
// Returns { code: { qty, lastDate } } where net = delivered - returned.
function _aggregateFgDelMain(ss, onlyCode) {
  const sh = ss.getSheetByName('FG Delivery Lines');
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  const reg = ss.getSheetByName('FG Delivery Register');
  const dateByDc = {};
  if (reg && reg.getLastRow() >= 2) {
    const rd = reg.getRange(2, 2, reg.getLastRow() - 1, 2).getValues();  // B=DC No, C=Date
    rd.forEach(function(r) { if (r[0]) dateByDc[r[0]] = r[1]; });
  }
  data.forEach(function(r) {
    if (r[2] !== 'Main') return;
    const code = r[3]; if (!code) return;
    if (onlyCode && code !== onlyCode) return;
    const net = (Number(r[5]) || 0) - (Number(r[9]) || 0);
    if (net <= 0) return;
    if (!out[code]) out[code] = { qty: 0, lastDate: null };
    out[code].qty += net;
    const d = dateByDc[r[0]];
    if (d && (!out[code].lastDate || d > out[code].lastDate)) out[code].lastDate = d;
  });
  return out;
}

// Compute an RM snapshot row from pre-aggregated maps (no register reads).
// 17 cols: physical on-hand becomes Available (physical - reserved). Reserved
// shown separately. Below MIN check uses Available (safer — accounts for locked stock).
function _computeRmRowFromMaps(it, maps) {
  const today = new Date();
  const openingQty = (maps.adjAll.opening[it.code] || 0);
  const inQty  = (maps.inward[it.code] && maps.inward[it.code].qty) || 0;
  const outQty = ((maps.req[it.code] && maps.req[it.code].qty) || 0)
               + ((maps.fgDelHw[it.code] && maps.fgDelHw[it.code].qty) || 0);
  const adjQty = (maps.adjAll.nonOpening[it.code] || 0);
  // RM-at-JW net: everything sent minus everything returned. Positive = still at JW.
  const jwOutNet = (((maps.jwOut && maps.jwOut[it.code] && maps.jwOut[it.code].qty) || 0)
                 - ((maps.jwIn  && maps.jwIn[it.code]  && maps.jwIn[it.code].qty)  || 0));
  const physical = openingQty + inQty - outQty + adjQty - jwOutNet;
  const reserved = (maps.reserved && maps.reserved[it.code]) || 0;
  const available = physical - reserved;
  // Avg Rate priority: Item Master col U (user-set) → inward-derived → 0.
  // Clamp physical at 0 for value calc — negative stock shouldn't produce
  // negative "value tied up" numbers that inflate the dashboard total.
  const masterRate = Number(it.avgRate) || 0;
  const derivedRate = (maps.inward[it.code] && maps.inward[it.code].avgQty)
    ? (maps.inward[it.code].avgSum / maps.inward[it.code].avgQty) : 0;
  const avgRate = masterRate > 0 ? masterRate : derivedRate;
  const stockValue = Math.max(0, physical) * avgRate;
  const lastIn  = (maps.inward[it.code] && maps.inward[it.code].lastDate) || '';
  const lastOut = _maxDate(
    (maps.req[it.code] && maps.req[it.code].lastDate),
    (maps.fgDelHw[it.code] && maps.fgDelHw[it.code].lastDate));
  const lastAny = _maxDate(lastIn, lastOut);
  const daysSince = lastAny ? Math.floor((today - lastAny) / 86400000) : '';
  const belowMin = (available < (it.minLvl || 0));
  return [
    it.code, it.desc, it.category, it.storage,
    openingQty, inQty, outQty, adjQty, available, reserved,
    it.minLvl, it.maxLvl, stockValue,
    lastIn || '', lastOut || '', daysSince, belowMin,
  ];
}

// Compute an SFG snapshot row from pre-aggregated maps.
// SFG on-hand deductions: JW outward (in-process) + FG Delivery Main (SFG shipped) + Reserved (soft-lock).
function _computeSfgRowFromMaps(sf, maps) {
  const opening = maps.adjAll.opening[sf.code] || 0;
  const adj     = maps.adjAll.nonOpening[sf.code] || 0;
  const backIn  = (maps.jwIn[sf.code]  && maps.jwIn[sf.code].qty)  || 0;
  const sentOut = (maps.jwOut[sf.code] && maps.jwOut[sf.code].qty) || 0;
  const delOut  = (maps.fgDelMain && maps.fgDelMain[sf.code] && maps.fgDelMain[sf.code].qty) || 0;
  const reserved = (maps.reserved && maps.reserved[sf.code]) || 0;
  const onHand = opening + adj + backIn - sentOut - delOut;
  const inProcess = (maps.jwOut[sf.code] && maps.jwOut[sf.code].openQty) || 0;
  const available = onHand - inProcess - reserved;
  const belowMin = (available < (sf.minLvl || 0));
  return [
    sf.code, sf.desc, [sf.size, sf.color].filter(String).join(' '),
    onHand, inProcess, available,
    sf.minLvl, sf.maxLvl, belowMin,
  ];
}

// Batch update N RM snapshot rows in one pass. Silent no-op for missing codes.
// RM section is bounded above by SFG header row (if present) so RM appends
// don't overwrite the spacer or SFG data.
function updateItemSnapshotsBatch(itemCodes) {
  if (!itemCodes || !itemCodes.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snap = ss.getSheetByName('Stock Snapshot');
  const items = ss.getSheetByName('Item Master');
  if (!snap || !items) return;
  const maps = _readAllRmAggregates(ss);
  const unique = Array.from(new Set(itemCodes));
  const toStr = function(v) {
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'd/M');
    return String(v == null ? '' : v).trim();
  };
  // Bound RM section to avoid clobbering spacer / SFG header when appending
  const sfgHeaderRow = _findSfgSectionStart(snap);
  const rmStopBefore = sfgHeaderRow > 0 ? sfgHeaderRow - 1 : 0;  // spacer row
  // PERF: read Item Master ONCE. Previously called getRowByCode per code
  // (each call re-scans the whole master — 391 rows × 22 cols for N codes).
  const itemsMap = {};
  if (items.getLastRow() >= 2) {
    const idata = items.getRange(2, 1, items.getLastRow() - 1, items.getLastColumn()).getValues();
    idata.forEach(function(r) { if (r[0]) itemsMap[r[0]] = r; });
  }
  // Dynamic Avg Rate column lookup — survives migration / column shuffles.
  const _arIdx = (typeof _avgRateCol === 'function') ? _avgRateCol(items) - 1 : 20;
  unique.forEach(function(code) {
    const itemRow = itemsMap[code];
    if (!itemRow) return;
    // ponytail: apostrophe-prefix storage so Sheets doesn't auto-parse "4/20" as
    // a date on write (serial 46132). Mirror of the _readMaster fix.
    const sParts = [toStr(itemRow[9]), toStr(itemRow[10])].filter(Boolean);
    const row = _computeRmRowFromMaps({
      code: code,
      desc: itemRow[2] || '',
      category: itemRow[7] || '',
      storage: sParts.length ? "'" + sParts.join('/') : '',
      minLvl: Number(itemRow[16]) || 0,
      maxLvl: Number(itemRow[17]) || 0,
      avgRate: (_arIdx >= 0) ? (Number(itemRow[_arIdx]) || 0) : 0,
    }, maps);
    const targetRow = _findOrAppendRowBounded(snap, 1, code, 2, rmStopBefore);
    if (targetRow > 0) snap.getRange(targetRow, 1, 1, 17).setValues([row]);
  });
}

// Batch update N SFG snapshot rows in one pass.
// SFG rows live below the SFG header (found dynamically via _findSfgSectionStart).
function updateSfgSnapshotsBatch(sfgCodes) {
  if (!sfgCodes || !sfgCodes.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snap = ss.getSheetByName('Stock Snapshot');
  const sfgs = ss.getSheetByName('SFG Master');
  if (!snap || !sfgs) return;
  const sfgHeaderRow = _findSfgSectionStart(snap);
  if (sfgHeaderRow < 0) return;  // SFG section not yet written — run fullRebuild
  const dataStart = sfgHeaderRow + 1;
  const maps = _readAllSfgAggregates(ss);
  const unique = Array.from(new Set(sfgCodes));
  // PERF: read SFG Master ONCE (was: per-code full-scan)
  const sfgsMap = {};
  if (sfgs.getLastRow() >= 2) {
    const sdata = sfgs.getRange(2, 1, sfgs.getLastRow() - 1, sfgs.getLastColumn()).getValues();
    sdata.forEach(function(r) { if (r[0]) sfgsMap[r[0]] = r; });
  }
  unique.forEach(function(code) {
    const sfgRow = sfgsMap[code];
    if (!sfgRow) return;
    const row = _computeSfgRowFromMaps({
      code: code,
      desc: sfgRow[2] || '',
      size: sfgRow[4] || '',
      color: sfgRow[5] || '',
      minLvl: Number(sfgRow[8]) || 0,
      maxLvl: Number(sfgRow[9]) || 0,
    }, maps);
    const targetRow = _findOrAppendRow(snap, 1, code, dataStart);
    snap.getRange(targetRow, 1, 1, 9).setValues([row]);
  });
}

// Bounded find-or-append: like _findOrAppendRow but never returns a row >=
// stopBefore. Returns -1 if no room (RM section full). Used for RM updates
// to protect the SFG section from being overwritten.
function _findOrAppendRowBounded(sh, codeCol, code, startRow, stopBefore) {
  if (!stopBefore) return _findOrAppendRow(sh, codeCol, code, startRow);
  const scanEnd = Math.min(sh.getLastRow(), stopBefore - 1);
  if (scanEnd >= startRow) {
    const vals = sh.getRange(startRow, codeCol, scanEnd - startRow + 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) if (vals[i][0] === code) return startRow + i;
    for (let i = 0; i < vals.length; i++) if (!vals[i][0]) return startRow + i;
  }
  // No blank + not found. Insert a new row before the spacer to make room.
  sh.insertRowsBefore(stopBefore, 1);
  return stopBefore;
}

// Single-code convenience wrappers around the batch versions.
// Kept for callers (postAdjustment, tool issue/return) that only touch one code.
function updateItemSnapshot(itemCode) { updateItemSnapshotsBatch([itemCode]); }
function updateSfgSnapshot(sfgCode)   { updateSfgSnapshotsBatch([sfgCode]); }

// ============================================================================
// Master + register readers
// ============================================================================

function _readMaster(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  if (name === 'Item Master') {
    // v3.5 idx: Desc [2], Item Name [3], Category [7], Storage [9], Rack [10], MIN [16], MAX [17], Active [18]
    // Avg Rate: dynamic header lookup (idx = col - 1; -1 if missing).
    const _arIdx = (typeof _avgRateCol === 'function') ? _avgRateCol(sh) - 1 : 20;
    const toStr = function(v) {
      if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'd/M');
      return String(v == null ? '' : v).trim();
    };
    return rows.filter(function(r) { return r[0] && r[18] !== false; }).map(function(r) {
      // ponytail: prefix storage with ' so Sheets treats the joined string as
      // literal text on write. Without this, "4/20" gets auto-parsed as a date
      // (April 20) and stored as serial 46132 — see EGB-0019 case.
      const sParts = [toStr(r[9]), toStr(r[10])].filter(Boolean);
      const storage = sParts.length ? "'" + sParts.join('/') : '';
      return { code: r[0], desc: r[2], category: r[7], storage: storage,
               minLvl: Number(r[16]) || 0, maxLvl: Number(r[17]) || 0,
               avgRate: (_arIdx >= 0) ? (Number(r[_arIdx]) || 0) : 0 };
    });
  }
  if (name === 'SFG Master') {
    // v3.5 idx: Desc [2], Item Name [3], Size [4], Color [5], MIN [8], MAX [9], Active [10]
    return rows.filter(function(r) { return r[0] && r[10] !== false; }).map(function(r) {
      return { code: r[0], desc: r[2], size: r[4], color: r[5],
               minLvl: Number(r[8]) || 0, maxLvl: Number(r[9]) || 0 };
    });
  }
  return rows;
}

// Aggregate a register into { code: { qty, avgSum, avgQty, lastDate, openQty } }
// opts: codeCol/dateCol/qtyCol/rateCol/statusCol (1-based); openOnly=true adds openQty;
//       only=<code> restricts scan to a single code (surgical update path).
function _aggregate(ss, sheetName, opts) {
  const sh = ss.getSheetByName(sheetName);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const code = row[opts.codeCol - 1];
    if (!code) continue;
    if (opts.only && code !== opts.only) continue;
    const qty = Number(row[opts.qtyCol - 1]) || 0;
    const date = row[opts.dateCol - 1];
    const rec = out[code] = out[code] || { qty: 0, avgSum: 0, avgQty: 0, lastDate: null, openQty: 0 };
    rec.qty += qty;
    if (opts.rateCol) {
      const rate = Number(row[opts.rateCol - 1]) || 0;
      rec.avgSum += qty * rate;
      rec.avgQty += qty;
    }
    if (date instanceof Date && (!rec.lastDate || date > rec.lastDate)) rec.lastDate = date;
    if (opts.statusCol) {
      const status = row[opts.statusCol - 1];
      if (status === 'Open') rec.openQty += qty;
    }
  }
  return out;
}

// Adjustments split by Type: "Opening" separately from everything else.
function _aggregateAdjustments(ss, onlyCode) {
  const sh = ss.getSheetByName('Adjustments Register');
  const out = { opening: {}, nonOpening: {} };
  if (!sh || sh.getLastRow() < 2) return out;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    // Cols: A=Sr, B=Date, C=Type, D=Code, E=Desc, F=Qty (signed)
    const code = data[i][3];
    if (!code) continue;
    if (onlyCode && code !== onlyCode) continue;
    const type = data[i][2];
    const qty = Number(data[i][5]) || 0;
    if (type === 'Opening') out.opening[code] = (out.opening[code] || 0) + qty;
    else                    out.nonOpening[code] = (out.nonOpening[code] || 0) + qty;
  }
  return out;
}

function _maxDate(a, b) {
  if (!a) return b || null;
  if (!b) return a;
  return a > b ? a : b;
}

function _findOrAppendRow(sh, codeCol, code, startRow) {
  const lastRow = sh.getLastRow();
  if (lastRow >= startRow) {
    const vals = sh.getRange(startRow, codeCol, lastRow - startRow + 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) if (vals[i][0] === code) return startRow + i;
    // Not found — first blank in that column
    for (let i = 0; i < vals.length; i++) if (!vals[i][0]) return startRow + i;
  }
  return Math.max(lastRow + 1, startRow);
}

// ============================================================================
// VIEWERS — rebuilt on picker change (via onEdit) or menu button.
// Each fn clears everything except header row 1 + picker rows, then rewrites.
// ============================================================================

// ---------- DASHBOARD ----------
// Layout: title band (row 1), refresh timestamp (row 2), then sections.
// All sections use grid A:H (8 cols). KPI cards span 2 cols each, 4 per row.
function rebuildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Dashboard');
  if (!sh) return;

  // Wipe old data + formatting (safe re-run)
  const lastRow = Math.max(sh.getLastRow(), 60);
  sh.getRange(1, 1, lastRow, 8).clear();
  sh.getBandings().forEach(function(b) { try { b.remove(); } catch(e){} });
  [180, 120, 180, 120, 180, 120, 180, 120].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  // ---- Title band ----
  sh.getRange('A1:H1').merge().setValue('3 STAR PROJECTS — IMS DASHBOARD')
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);
  sh.getRange('A2:H2').merge()
    .setValue('Last refreshed: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') +
              '  |  Menu → 3 Star IMS → Refresh → Dashboard to update')
    .setBackground('#F47C2C').setFontColor('#FFFFFF').setFontSize(10).setHorizontalAlignment('center');
  sh.setRowHeight(2, 22);

  // ---- Gather data ----
  const items = _readMaster(ss, 'Item Master');
  const sfgs = _readMaster(ss, 'SFG Master');
  const fgSh = ss.getSheetByName('FG Master');
  const fgs = fgSh && fgSh.getLastRow() > 1
    ? fgSh.getRange(2, 1, fgSh.getLastRow() - 1, 1).getValues().filter(function(r) { return r[0]; }).length : 0;
  const supSh = ss.getSheetByName('Supplier Master');
  const suppliers = supSh && supSh.getLastRow() > 1
    ? supSh.getRange(2, 1, supSh.getLastRow() - 1, 1).getValues().filter(function(r) { return r[0]; }).length : 0;
  const jwSh = ss.getSheetByName('Job Worker Master');
  const jobWorkers = jwSh && jwSh.getLastRow() > 1
    ? jwSh.getRange(2, 1, jwSh.getLastRow() - 1, 1).getValues().filter(function(r) { return r[0]; }).length : 0;

  const snap = ss.getSheetByName('Stock Snapshot');
  let totalValue = 0, itemsBelowMin = 0, nonMoving = 0, sfgInProcessUnits = 0, sfgInProcessJobs = 0;
  const nonMovingThreshold = Number(getConfig('Non-Moving Days Threshold')) || 115;
  // PERF: read Snapshot RM section ONCE — every downstream aggregator receives
  // this via param instead of re-reading. Cuts ~5× full-sheet getValues to 1.
  let rmDataShared = [];
  if (snap && snap.getLastRow() >= 2) {
    const sfgHeaderRow = _findSfgSectionStart(snap);
    const rmEnd = sfgHeaderRow > 0 ? sfgHeaderRow - 2 : snap.getLastRow();
    if (rmEnd >= 2) {
      rmDataShared = snap.getRange(2, 1, rmEnd - 1, 17).getValues();
      rmDataShared.forEach(function(r) {
        if (!r[0]) return;
        totalValue += _safeStockValue(r[12]);   // Stock Value (col M, idx 12)
        if (r[16] === true) itemsBelowMin++;    // Below MIN (col Q, idx 16)
        if (r[15] !== '' && Number(r[15]) >= nonMovingThreshold) nonMoving++;  // Days Since (col P, idx 15)
      });
    }
    // SFG section: rows (sfgHeader+1)..lastRow, cols A:I
    if (sfgHeaderRow > 0 && snap.getLastRow() > sfgHeaderRow) {
      const sfgData = snap.getRange(sfgHeaderRow + 1, 1, snap.getLastRow() - sfgHeaderRow, 9).getValues();
      sfgData.forEach(function(r) {
        if (!r[0]) return;
        sfgInProcessUnits += Number(r[4]) || 0;  // In Process (col E, idx 4)
      });
    }
  }
  const jwOutSh = ss.getSheetByName('Job Work Outward Register');
  let openJwc = 0, overdueJwc = 0, rmInProcessUnits = 0;
  const today = new Date();
  if (jwOutSh && jwOutSh.getLastRow() >= 2) {
    const jwData = jwOutSh.getRange(2, 1, jwOutSh.getLastRow() - 1, 16).getValues();
    const seenJwc = {};
    jwData.forEach(function(r) {
      if (r[13] === 'Open' && !seenJwc[r[1]]) {
        seenJwc[r[1]] = true;
        openJwc++;
        if (r[2] instanceof Date) {
          const daysOut = Math.floor((today - r[2]) / 86400000);
          if (daysOut > 21) overdueJwc++;
        }
      }
    });
    sfgInProcessJobs = openJwc;
    // ponytail: RM-at-JW units = net (sent − received) per RM code, summed only
    // where positive. Share the jwData already in memory for the SFG counts above.
    const itemsSh = ss.getSheetByName('Item Master');
    if (itemsSh && itemsSh.getLastRow() >= 2) {
      const rmCodes = {};
      itemsSh.getRange(2, 1, itemsSh.getLastRow() - 1, 1).getValues()
        .forEach(function(r) { if (r[0]) rmCodes[r[0]] = true; });
      const netByCode = {};
      jwData.forEach(function(r) {
        const code = r[6];  // col G = Code
        if (!code || !rmCodes[code]) return;
        netByCode[code] = (netByCode[code] || 0) + (Number(r[8]) || 0);  // col I = Qty
      });
      const jwInSh = ss.getSheetByName('Job Work Inward Register');
      if (jwInSh && jwInSh.getLastRow() >= 2) {
        jwInSh.getRange(2, 1, jwInSh.getLastRow() - 1, 10).getValues().forEach(function(r) {
          const code = r[5];  // col F = Code
          if (!code || !rmCodes[code]) return;
          netByCode[code] = (netByCode[code] || 0) - (Number(r[7]) || 0);  // col H = Qty Received
        });
      }
      Object.keys(netByCode).forEach(function(c) {
        if (netByCode[c] > 0) rmInProcessUnits += netByCode[c];
      });
    }
  }
  const categoriesUsed = Array.from(new Set(items.map(function(i) { return i.category; }).filter(String))).length;

  // ---- Section: INVENTORY SNAPSHOT ----
  _dashSectionHeader(sh, 4, 'INVENTORY SNAPSHOT', '#1F4E79');
  _dashKpiCard(sh, 5, 1, 'ITEMS (RM)',    items.length,          '#EBF3FB');
  _dashKpiCard(sh, 5, 3, 'SFGs',           sfgs.length,           '#EBF3FB');
  _dashKpiCard(sh, 5, 5, 'FGs',            fgs,                   '#EBF3FB');
  _dashKpiCard(sh, 5, 7, 'STOCK VALUE',    '₹ ' + _fmtIndian(totalValue), '#EBF3FB', 14);

  // ---- Section: ATTENTION NEEDED ----
  _dashSectionHeader(sh, 9, 'ATTENTION NEEDED', '#C0392B');
  _dashKpiCard(sh, 10, 1, 'BELOW MIN',     itemsBelowMin + ' items',
    itemsBelowMin > 5 ? '#FADBD8' : (itemsBelowMin ? '#FDEBD0' : '#D5F5E3'));
  _dashKpiCard(sh, 10, 3, 'AT JOB WORK',
    sfgInProcessUnits + ' SFG + ' + rmInProcessUnits + ' RM\n@ ' + openJwc + ' job(s)',
    (openJwc || rmInProcessUnits) ? '#FDEBD0' : '#D5F5E3', 11);
  _dashKpiCard(sh, 10, 5, 'OVERDUE JW',    overdueJwc + ' job(s)',
    overdueJwc ? '#FADBD8' : '#D5F5E3');
  _dashKpiCard(sh, 10, 7, 'NON-MOVING',    nonMoving + ' items\n> ' + nonMovingThreshold + ' days',
    nonMoving > 10 ? '#FADBD8' : (nonMoving ? '#FDEBD0' : '#D5F5E3'), 12);

  // ---- Section: MASTERS ----
  _dashSectionHeader(sh, 14, 'MASTERS', '#1F4E79');
  sh.getRange('A15:H15').merge()
    .setValue('Suppliers: ' + suppliers + '   |   Job Workers: ' + jobWorkers +
              '   |   Categories in use: ' + categoriesUsed + '   |   Config: ' +
              (getConfig('Company GSTIN') ? 'GSTIN set' : 'GSTIN missing'))
    .setBackground('#F8F8F8').setFontSize(11).setHorizontalAlignment('center');
  sh.setRowHeight(15, 28);

  // ---- Section: RECENT ACTIVITY ----
  _dashSectionHeader(sh, 17, 'RECENT ACTIVITY (last 10)', '#1F4E79');
  // Recent GRNs (left, A:D)
  sh.getRange('A18:D18').setValues([['GRN No', 'Date', 'Supplier', 'Value']])
    .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  const recentGrns = _recentGrns(ss, 10);
  if (recentGrns.length) {
    sh.getRange(19, 1, recentGrns.length, 4).setValues(recentGrns);
    try { sh.getRange(19, 1, recentGrns.length, 4).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); } catch(e){}
  }
  // Recent Requisitions (right, E:H)
  sh.getRange('E18:H18').setValues([['Req No', 'Date', 'Project', 'Items']])
    .setFontWeight('bold').setBackground('#F47C2C').setFontColor('#FFFFFF');
  const recentReqs = _recentReqs(ss, 10);
  if (recentReqs.length) {
    sh.getRange(19, 5, recentReqs.length, 4).setValues(recentReqs);
    try { sh.getRange(19, 5, recentReqs.length, 4).applyRowBanding(SpreadsheetApp.BandingTheme.ORANGE, false, false); } catch(e){}
  }

  // ---- Section: STOCK VALUE BY CATEGORY ----
  const catStartRow = 32;
  _dashSectionHeader(sh, catStartRow, 'STOCK VALUE BY CATEGORY', '#1F4E79');
  sh.getRange(catStartRow + 1, 1, 1, 4).setValues([['Category', 'Items', 'Value (Rs.)', '% Share ▓']])
    .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  const byCategory = _stockByCategory(ss, rmDataShared);
  let catEndRow = catStartRow + 1;
  if (byCategory.length) {
    const rows = byCategory.map(function(c) {
      const pct = totalValue ? (c.value / totalValue * 100) : 0;
      return [c.category, c.count, '₹ ' + _fmtIndian(c.value), pct.toFixed(1) + '%  ' + '█'.repeat(Math.round(pct / 5))];
    });
    sh.getRange(catStartRow + 2, 1, rows.length, 4).setValues(rows);
    try { sh.getRange(catStartRow + 2, 1, rows.length, 4).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); } catch(e){}
    for (let i = 0; i < rows.length; i++) {
      sh.getRange(catStartRow + 2 + i, 5, 1, 4).setBackground('#FFFFFF');
    }
    catEndRow = catStartRow + 1 + rows.length;
  }

  // ---- Section: OPERATIONS (reservations + this month's flow) ----
  const opsRow = catEndRow + 2;
  const opsData = _dashOperationsData(ss);
  _dashSectionHeader(sh, opsRow, 'OPERATIONS (last 30 days)', '#1F4E79');
  _dashKpiCard(sh, opsRow + 1, 1, 'ACTIVE RESERVATIONS', opsData.activeReservations + ' item(s)',
    opsData.activeReservations ? '#D6EAF8' : '#EBF3FB', 14);
  _dashKpiCard(sh, opsRow + 1, 3, 'RESERVED VALUE', '₹ ' + _fmtIndian(opsData.reservedValue), '#D6EAF8', 14);
  _dashKpiCard(sh, opsRow + 1, 5, 'GRNs (30d)', opsData.grnsThisMonth + ' | ₹ ' + _fmtIndian(opsData.grnValueThisMonth),
    '#EBF3FB', 12);
  _dashKpiCard(sh, opsRow + 1, 7, 'DELIVERIES (30d)', opsData.deliveriesThisMonth + ' | ₹ ' + _fmtIndian(opsData.deliveryValueThisMonth),
    '#EBF3FB', 12);

  // ---- Section: STOCK AGING (value tied up by days-since-movement) ----
  const ageRow = opsRow + 5;
  const buckets = _dashAgingBuckets(ss, rmDataShared);
  _dashSectionHeader(sh, ageRow, 'STOCK AGING (RM value by days since last movement)', '#F47C2C');
  sh.getRange(ageRow + 1, 1, 1, 4).setValues([['Age Bucket', 'Items', 'Value (Rs.)', 'Share ▓']])
    .setFontWeight('bold').setBackground('#F47C2C').setFontColor('#FFFFFF');
  const ageRows = buckets.map(function(b) {
    const pct = totalValue ? (b.value / totalValue * 100) : 0;
    return [b.label, b.count, '₹ ' + _fmtIndian(b.value), pct.toFixed(1) + '%  ' + '█'.repeat(Math.round(pct / 5))];
  });
  sh.getRange(ageRow + 2, 1, ageRows.length, 4).setValues(ageRows);
  try { sh.getRange(ageRow + 2, 1, ageRows.length, 4).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); } catch(e){}
  // color the "> 180 days" row red
  const oldRow = ageRow + 2 + ageRows.length - 1;
  sh.getRange(oldRow, 1, 1, 4).setBackground('#FADBD8').setFontColor('#C0392B');

  // ---- Section: TOP 10 NON-MOVING RM ITEMS ----
  const topRow = ageRow + 2 + ageRows.length + 2;
  const topNonMoving = _dashTopNonMoving(ss, 10, rmDataShared);
  _dashSectionHeader(sh, topRow, 'TOP 10 NON-MOVING RM ITEMS (by value tied up)', '#C0392B');
  sh.getRange(topRow + 1, 1, 1, 6).setValues([['Item Code', 'Description', 'Category', 'Available', 'Days Since Movement', 'Value (Rs.)']])
    .setFontWeight('bold').setBackground('#C0392B').setFontColor('#FFFFFF');
  if (topNonMoving.length) {
    sh.getRange(topRow + 2, 1, topNonMoving.length, 6).setValues(topNonMoving);
    try { sh.getRange(topRow + 2, 1, topNonMoving.length, 6).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false); } catch(e){}
    // Right-align qty/days/value cols
    sh.getRange(topRow + 2, 4, topNonMoving.length, 3).setHorizontalAlignment('right');
  } else {
    sh.getRange(topRow + 2, 1, 1, 6).merge().setValue('No non-moving items — everything is moving!')
      .setBackground('#D5F5E3').setFontColor('#1E8449').setHorizontalAlignment('center');
  }

  sh.setFrozenRows(2);
}

// Aggregate operational metrics for Dashboard OPERATIONS section.
function _dashOperationsData(ss) {
  const out = { activeReservations: 0, reservedValue: 0,
                grnsThisMonth: 0, grnValueThisMonth: 0,
                deliveriesThisMonth: 0, deliveryValueThisMonth: 0 };
  const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 30);
  // Reservations
  const resSh = ss.getSheetByName('Reservation Register');
  if (resSh && resSh.getLastRow() >= 2) {
    const rd = resSh.getRange(2, 1, resSh.getLastRow() - 1, 9).getValues();
    const rmCosts = _itemCostMap(ss, rmDataShared);
    const seenRes = {};
    rd.forEach(function(r) {
      if (r[8] !== 'Active') return;
      if (!r[3]) return;
      seenRes[r[1]] = true;  // count distinct reservation numbers
      out.reservedValue += (Number(r[5]) || 0) * (rmCosts[r[3]] || 0);
    });
    out.activeReservations = Object.keys(seenRes).length;
  }
  // GRNs last 30 days
  const inSh = ss.getSheetByName('Inward Register');
  if (inSh && inSh.getLastRow() >= 2) {
    const gd = inSh.getRange(2, 1, inSh.getLastRow() - 1, 15).getValues();
    const seenGrn = {};
    gd.forEach(function(r) {
      if (!(r[2] instanceof Date) || r[2] < cutoff) return;
      if (r[1] && !seenGrn[r[1]]) { seenGrn[r[1]] = true; out.grnsThisMonth++; }
      out.grnValueThisMonth += Number(r[14]) || 0;
    });
  }
  // Deliveries last 30 days — from FG Delivery Register (one row per DC)
  const dcSh = ss.getSheetByName('FG Delivery Register');
  if (dcSh && dcSh.getLastRow() >= 2) {
    const dd = dcSh.getRange(2, 1, dcSh.getLastRow() - 1, 17).getValues();
    dd.forEach(function(r) {
      if (!(r[2] instanceof Date) || r[2] < cutoff) return;
      out.deliveriesThisMonth++;
      out.deliveryValueThisMonth += 0;  // No line-total on DC register; skip value or compute from lines
    });
  }
  return out;
}

// Cost per RM code from Stock Snapshot (Value / Available) — used for reserved-value calc.
function _itemCostMap(ss, rmData) {
  const out = {};
  const snap = ss.getSheetByName('Stock Snapshot');
  if (!snap) return out;
  const data = rmData || _readRmSection(snap, 13);
  data.forEach(function(r) {
    if (!r[0]) return;
    const avail = Number(r[8]) || 0;
    const val = Number(r[12]) || 0;
    if (avail > 0) out[r[0]] = val / avail;
  });
  return out;
}

// Group RM items into age buckets by "Days Since Movement" (Snapshot col P).
// Coerce a Snapshot col-M value to a safe positive number.
// Handles: numbers, Date objects (from leftover Date-formatting), '₹800.00'
// text strings, and blanks. Anything unparseable or absurd returns 0.
function _safeStockValue(x) {
  if (x == null || x === '') return 0;
  if (x instanceof Date) return 0;   // date-typed cell — not a real amount
  if (typeof x === 'number') return (isFinite(x) && x > 0 && x < 1e12) ? x : 0;
  if (typeof x === 'string') {
    const n = Number(String(x).replace(/[₹Rs.,\s]/g, ''));
    return (isFinite(n) && n > 0 && n < 1e12) ? n : 0;
  }
  return 0;
}

function _dashAgingBuckets(ss, rmData) {
  const buckets = [
    { label: '0-30 days',    max: 30,  count: 0, value: 0 },
    { label: '31-60 days',   max: 60,  count: 0, value: 0 },
    { label: '61-90 days',   max: 90,  count: 0, value: 0 },
    { label: '91-180 days',  max: 180, count: 0, value: 0 },
    { label: '> 180 days',   max: Infinity, count: 0, value: 0 },
  ];
  const snap = ss.getSheetByName('Stock Snapshot');
  if (!snap) return buckets;
  const data = rmData || _readRmSection(snap, 17);
  data.forEach(function(r) {
    if (!r[0]) return;
    const days = Number(r[15]) || 0;
    const value = _safeStockValue(r[12]);
    for (let i = 0; i < buckets.length; i++) {
      if (days <= buckets[i].max) {
        buckets[i].count++;
        buckets[i].value += value;
        break;
      }
    }
  });
  return buckets;
}

// Top-N RM items with highest stock value AND days since movement > threshold.
function _dashTopNonMoving(ss, limit, rmData) {
  const threshold = Number(getConfig('Non-Moving Days Threshold')) || 115;
  const snap = ss.getSheetByName('Stock Snapshot');
  if (!snap) return [];
  const data = rmData || _readRmSection(snap, 17);
  const rows = data
    .filter(function(r) { return r[0] && (Number(r[15]) || 0) >= threshold && (Number(r[12]) || 0) > 0; })
    .map(function(r) {
      return [r[0], r[1] || '', r[2] || '', Number(r[8]) || 0, Number(r[15]) || 0, '₹ ' + _fmtIndian(r[12])];
    })
    .sort(function(a, b) {
      // Sort by numeric value descending — need to parse back from the ₹-formatted string
      const av = parseFloat(String(a[5]).replace(/[^\d.-]/g, '')) || 0;
      const bv = parseFloat(String(b[5]).replace(/[^\d.-]/g, '')) || 0;
      return bv - av;
    });
  return rows.slice(0, limit);
}

// Section header row — merged, colored underline
function _dashSectionHeader(sh, row, text, color) {
  sh.getRange(row, 1, 1, 8).merge()
    .setValue(text)
    .setFontSize(12).setFontWeight('bold').setFontColor(color)
    .setBorder(false, false, true, false, false, false, color, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.setRowHeight(row, 26);
}

// KPI card = 2 cols wide, 3 rows tall (label row + big number row + spacer)
function _dashKpiCard(sh, row, col, label, value, bg, valueFontSize) {
  sh.getRange(row, col, 1, 2).merge().setValue(label)
    .setFontSize(9).setFontWeight('bold').setFontColor('#555555')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(bg);
  sh.getRange(row + 1, col, 2, 2).merge().setValue(value)
    .setFontSize(valueFontSize || 22).setFontWeight('bold').setFontColor('#1F4E79')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(bg)
    .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID)
    .setWrap(true);
  sh.setRowHeight(row, 20);
  sh.setRowHeight(row + 1, 34);
  sh.setRowHeight(row + 2, 20);
}

// Indian numbering format: 12,34,567.89
function _fmtIndian(n) {
  const num = Number(n) || 0;
  const parts = num.toFixed(2).split('.');
  const x = parts[0]; const last3 = x.slice(-3); const rest = x.slice(0, -3);
  const grouped = rest ? rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + last3 : last3;
  return grouped + '.' + parts[1];
}

function _recentGrns(ss, limit) {
  const sh = ss.getSheetByName('Inward Register');
  if (!sh || sh.getLastRow() < 2) return [];
  // One GRN may span multiple rows — dedupe by GRN No, sum values
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 20).getValues();
  const map = {};
  data.forEach(function(r) {
    const grn = r[1]; if (!grn) return;
    if (!map[grn]) map[grn] = { grn: grn, date: r[2], supplier: r[6] || r[5] || '', value: 0 };
    map[grn].value += Number(r[14]) || 0;
  });
  const list = Object.keys(map).map(function(k) {
    const m = map[k];
    return [m.grn, m.date, m.supplier, '₹ ' + _fmtIndian(m.value)];
  });
  list.sort(function(a, b) { return (a[1] < b[1]) ? 1 : -1; });
  return list.slice(0, limit);
}

function _recentReqs(ss, limit) {
  const sh = ss.getSheetByName('Material Requisition Register');
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 15).getValues();
  const map = {};
  data.forEach(function(r) {
    const req = r[1]; if (!req) return;
    if (!map[req]) map[req] = { req: req, date: r[2], project: r[3] || '', items: 0 };
    map[req].items++;
  });
  const list = Object.keys(map).map(function(k) {
    const m = map[k];
    return [m.req, m.date, m.project, m.items + ' item(s)'];
  });
  list.sort(function(a, b) { return (a[1] < b[1]) ? 1 : -1; });
  return list.slice(0, limit);
}

function _stockByCategory(ss, rmData) {
  const snap = ss.getSheetByName('Stock Snapshot');
  const data = rmData || _readRmSection(snap, 17);
  const bucket = {};
  data.forEach(function(r) {
    if (!r[0]) return;
    const cat = r[2] || 'Uncategorized';
    if (!bucket[cat]) bucket[cat] = { category: cat, count: 0, value: 0 };
    bucket[cat].count++;
    bucket[cat].value += _safeStockValue(r[12]);
  });
  const out = Object.keys(bucket).map(function(k) { return bucket[k]; });
  out.sort(function(a, b) { return b.value - a.value; });
  return out;
}

// ---------- EXPLORER ----------
// One sheet, 6 view modes. B1 = view picker, D1 = view-dependent input,
// C1 = auto-label for D1. Data starts row 3.
function rebuildExplorer(view, arg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Explorer');
  if (!sh) return;

  // Clear rows 2+, reset formatting, kill bandings
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, Math.max(sh.getLastColumn(), 12)).clear();
  sh.getBandings().forEach(function(b) { try { b.remove(); } catch(e){} });

  // Row 1 toolbar — subtle band under the view picker + input.
  // A1 = title chip, B1 = view picker (kept — user selects here), C1/D1 set by _setExplorerInput.
  sh.getRange('A1:I1').setBackground('#F7F9FB').setBorder(false, false, true, false, false, false, '#D6DBE3', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('A1').setValue('🔍  EXPLORER').setFontWeight('bold').setFontColor('#1F4E79')
    .setFontSize(11).setBackground('#EBF3FB').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 32);

  // Reset C1 + D1 to a clean slate BEFORE setting new label/validation.
  // Prevents old validation from clashing with new value when view swaps.
  sh.getRange('C1').clearContent().clearFormat();
  sh.getRange('D1').clearDataValidations();
  _setExplorerInput(sh, view, ss);
  // Row 2 = header area; data starts row 3
  const dispatch = {
    'Item Ledger':       _renderItemLedger,
    'Supplier Ledger':   _renderSupplierLedger,
    'Stock by Category': _renderCategoryStock,
    'Non-Moving Stock':  _renderNonMoving,
    'Low Stock Alerts':  _renderLowStock,
    'Pending Job Work':  _renderPendingJobWork,
    'Reorder Suggestions': _renderReorderSuggestions,
    'Site Consumption':    _renderSiteConsumption,
    'Project Requisitions': _renderProjectReqs,
    'QC Log':              _renderQcLog,
  };
  const fn = dispatch[view];
  if (fn) {
    // Views that need input from D1 — show a hint until user picks
    const needsInput = ['Item Ledger', 'Supplier Ledger', 'Stock by Category', 'Site Consumption', 'Project Requisitions'].indexOf(view) >= 0;
    if (needsInput && !arg) {
      sh.getRange('A3:H3').merge()
        .setValue('← Pick a value in cell D1 to load ' + view + '.')
        .setFontStyle('italic').setFontColor('#888888').setHorizontalAlignment('center')
        .setBackground('#FFFBE6');
      sh.setRowHeight(3, 32);
    } else {
      fn(sh, ss, arg);
    }
  }
  sh.setFrozenRows(2);
  // Auto-fit columns so long values (item names, descriptions, dates) stop
  // getting clipped. Covers whatever the current view wrote, up to col M.
  try { sh.autoResizeColumns(1, Math.min(13, Math.max(1, sh.getLastColumn()))); } catch (e) {}
}

// Update C1 label and D1 validation based on the current view
function _setExplorerInput(sh, view, ss) {
  const im  = ss.getSheetByName('Item Master');
  const sup = ss.getSheetByName('Supplier Master');
  const dd  = ss.getSheetByName('Dropdowns');
  const cLabel = sh.getRange('C1');
  const dInput = sh.getRange('D1');
  cLabel.setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF').setHorizontalAlignment('right');
  dInput.setBackground('#FFF2CC').setFontWeight('bold');
  dInput.clearDataValidations();

  // Use SEARCH_LABEL_COL from 00_Config_Menu_Setup.gs — dropdown source
  // becomes the composite Search Label, so users can type "white" to filter.
  const imLabelCol  = (typeof SEARCH_LABEL_COL !== 'undefined' && SEARCH_LABEL_COL['Item Master'])     || 1;
  const supLabelCol = (typeof SEARCH_LABEL_COL !== 'undefined' && SEARCH_LABEL_COL['Supplier Master']) || 1;
  switch (view) {
    case 'Item Ledger':
      cLabel.setValue('Item (search):');
      if (im && im.getLastRow() > 1) {
        dInput.setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInRange(im.getRange(2, imLabelCol, im.getLastRow() - 1, 1), true).setAllowInvalid(false).build());
      }
      break;
    case 'Supplier Ledger':
      cLabel.setValue('Supplier (search):');
      if (sup && sup.getLastRow() > 1) {
        dInput.setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInRange(sup.getRange(2, supLabelCol, sup.getLastRow() - 1, 1), true).setAllowInvalid(false).build());
      }
      break;
    case 'Stock by Category':
      cLabel.setValue('Category:');
      if (dd) {
        dInput.setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInRange(dd.getRange(2, 1, Math.max(1, dd.getLastRow() - 1), 1), true).setAllowInvalid(false).build());
      }
      break;
    case 'Non-Moving Stock':
      cLabel.setValue('Days threshold:');
      if (!dInput.getValue()) dInput.setValue(Number(getConfig('Non-Moving Days Threshold')) || 115);
      break;
    case 'Low Stock Alerts':
    case 'Pending Job Work':
    case 'Reorder Suggestions':
    case 'QC Log':
      cLabel.setValue('(no input)');
      dInput.setValue('').setBackground('#F0F0F0');
      break;
    case 'Site Consumption':
    case 'Project Requisitions':
      cLabel.setValue('Project:');
      if (dd) {
        // Projects live at Dropdowns col K (index 11)
        const projLastRow = Math.max(1, dd.getLastRow() - 1);
        dInput.setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInRange(dd.getRange(2, 11, projLastRow, 1), true).setAllowInvalid(true).build());
      }
      break;
  }
}

// ----- 6 render fns for Explorer -----

function _renderItemLedger(sh, ss, itemCode) {
  itemCode = extractCodePrefix(itemCode);
  if (!itemCode) return;
  const itemRow = getRowByCode(ss.getSheetByName('Item Master'), itemCode);
  if (!itemRow) { sh.getRange('A3').setValue('Item not found.').setFontStyle('italic'); return; }

  // Gather all movements:
  //   Inward (In), Requisition (Out), Adjustment (signed In/Out),
  //   FG Delivery hardware (Out), Reservations (soft lock — doesn't move stock).
  const events = [];
  _fetchRegisterRowsCustom(ss, 'Inward Register', 4, itemCode, function(r) {
    events.push({ date: r[2], type: 'Inward (GRN)', ref: r[1] || '', party: r[6] || '', qtyIn: Number(r[7]) || 0, qtyOut: 0, reserved: 0, remarks: r[16] || '' });
  });
  _fetchRegisterRowsCustom(ss, 'Material Requisition Register', 5, itemCode, function(r) {
    events.push({ date: r[2], type: 'Requisition', ref: r[1] || '', party: r[3] || '', qtyIn: 0, qtyOut: Number(r[6]) || 0, reserved: 0, remarks: r[8] || '' });
  });
  _fetchRegisterRowsCustom(ss, 'Adjustments Register', 4, itemCode, function(r) {
    const q = Number(r[5]) || 0;
    events.push({ date: r[1], type: 'Adjustment', ref: r[2] || '', party: r[7] || '', qtyIn: q > 0 ? q : 0, qtyOut: q < 0 ? -q : 0, reserved: 0, remarks: r[8] || '' });
  });
  // FG Delivery hardware section — RM outward via a DC. One row per hardware line
  // in FG Delivery Lines where Section='Hardware'. Date + customer come from the
  // parent FG Delivery Register row.
  const dcLines = ss.getSheetByName('FG Delivery Lines');
  const dcReg   = ss.getSheetByName('FG Delivery Register');
  if (dcLines && dcLines.getLastRow() >= 2) {
    const dateByDc = {}, custByDc = {};
    if (dcReg && dcReg.getLastRow() >= 2) {
      dcReg.getRange(2, 1, dcReg.getLastRow() - 1, 5).getValues().forEach(function(r) {
        if (r[1]) { dateByDc[r[1]] = r[2]; custByDc[r[1]] = r[3] || ''; }
      });
    }
    dcLines.getRange(2, 1, dcLines.getLastRow() - 1, 10).getValues().forEach(function(r) {
      if (r[2] !== 'Hardware' || r[3] !== itemCode) return;
      const netQty = (Number(r[5]) || 0) - (Number(r[9]) || 0);
      if (netQty <= 0) return;
      events.push({ date: dateByDc[r[0]] || '', type: 'FG Delivery (HW)', ref: r[0] || '', party: custByDc[r[0]] || '', qtyIn: 0, qtyOut: netQty, reserved: 0, remarks: r[8] || '' });
    });
  }
  // Reservations — Active rows add to the soft lock; Released/Consumed rows are
  // informational (the lock is already gone). Status appears in the Type column.
  _fetchRegisterRowsCustom(ss, 'Reservation Register', 4, itemCode, function(r) {
    const qty = Number(r[5]) || 0;
    const status = r[8] || 'Active';
    events.push({
      date: r[2],
      type: 'Reservation (' + status + ')',
      ref: r[1] || '',
      party: r[7] || '',
      qtyIn: 0, qtyOut: 0,
      reserved: status === 'Active' ? qty : 0,
      remarks: r[9] || '',
    });
  });
  // Job Work events — RM may be sent to a Job Worker (Change 2: items→JW).
  // JW Out cols: A=Sr,B=JWC,C=Date,D=JWCode,E=JWName,F=Process,G=Code,H=Desc,I=Qty
  _fetchRegisterRowsCustom(ss, 'Job Work Outward Register', 7, itemCode, function(r) {
    events.push({ date: r[2], type: 'JW Outward', ref: r[1] || '', party: r[4] || '', qtyIn: 0, qtyOut: Number(r[8]) || 0, reserved: 0, remarks: r[11] || '' });
  });
  // JW In cols: A=Sr,B=JWGRN,C=Date,D=RefChallan,E=JWName,F=Code,G=Desc,H=QtyRecv
  _fetchRegisterRowsCustom(ss, 'Job Work Inward Register', 6, itemCode, function(r) {
    events.push({ date: r[2], type: 'JW Inward', ref: r[1] || '', party: r[4] || '', qtyIn: Number(r[7]) || 0, qtyOut: 0, reserved: 0, remarks: r[11] || '' });
  });

  events.sort(function(a, b) {
    const ad = (a.date instanceof Date) ? a.date.getTime() : 0;
    const bd = (b.date instanceof Date) ? b.date.getTime() : 0;
    return ad - bd;
  });

  let totalIn = 0, totalOut = 0, netReserved = 0;
  events.forEach(function(e) { totalIn += e.qtyIn; totalOut += e.qtyOut; netReserved += e.reserved; });
  const balance = totalIn - totalOut;              // physical on-hand
  const available = balance - netReserved;         // what you can actually commit
  // Avg Rate — Item Master Avg Rate (Rs.) col (user-set) wins; else derive from Inward history.
  const _arIdx = (typeof _avgRateCol === 'function') ? _avgRateCol(ss.getSheetByName('Item Master')) - 1 : 20;
  const avgRateFromMaster = (_arIdx >= 0) ? (Number(itemRow[_arIdx]) || 0) : 0;
  const avgRate = avgRateFromMaster > 0 ? avgRateFromMaster : Number((_avgRateMap(ss) || {})[itemCode] || 0);

  // Summary strip (row 2) — 6 KPI chips: Item | IN | OUT | Bal | Rsvd | Avg Rate | Available
  sh.getRange('A2:J2').clearContent().clearFormat();
  sh.getRange('A2:B2').merge().setValue('  ' + itemCode + '  ›  ' + (itemRow[2] || ''))
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.getRange('C2').setValue('IN: ' + _fmtIndian(totalIn))
    .setBackground('#D5F5E3').setFontColor('#1E8449').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('D2').setValue('OUT: ' + _fmtIndian(totalOut))
    .setBackground('#FADBD8').setFontColor('#C0392B').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('E2').setValue('BAL: ' + _fmtIndian(balance))
    .setBackground(balance < 0 ? '#C0392B' : '#F47C2C').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.getRange('F2').setValue('RSVD: ' + _fmtIndian(netReserved))
    .setBackground('#D6EAF8').setFontColor('#1F4E79').setFontWeight('bold').setHorizontalAlignment('center');
  // Unit label at G1 (above the ₹/u chip) so you can tell meter vs no.s vs kg.
  const uomLbl = (itemRow[6] || '').toString().trim();
  sh.getRange('G1').setValue(uomLbl ? ('Unit: ' + uomLbl) : '')
    .setBackground('#FDEBD0').setFontColor('#8E6600').setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center');
  sh.getRange('G2').setValue('₹/u: ' + (avgRate > 0 ? _fmtIndian(avgRate) : '—'))
    .setBackground('#FDEBD0').setFontColor('#8E6600').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('H2:J2').merge().setValue('AVAILABLE: ' + _fmtIndian(available))
    .setBackground(available < 0 ? '#C0392B' : '#1E8449').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(11);
  sh.setRowHeight(2, 30);

  // Unified chronological ledger — now 10 cols with Reserved + Balance side by side.
  const headers = ['Sr','Date','Type','Ref No','Party / Purpose','Qty IN','Qty OUT','Reserved','Balance','Remarks'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);

  let runningBal = 0;
  const rows = events.map(function(e, i) {
    runningBal += e.qtyIn - e.qtyOut;  // Balance tracks physical only
    return [i + 1, e.date, e.type, e.ref, e.party, e.qtyIn || '', e.qtyOut || '', e.reserved || '', runningBal, e.remarks];
  });
  if (rows.length) sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows.length),
    widths: [40, 95, 150, 110, 220, 80, 80, 85, 95, 220], freezeRows: 3,
    dateCols: [2], centerCols: [1, 3], qtyCols: [6, 7, 8, 9], boldCols: [9] });

  if (rows.length) {
    sh.getRange(4, 9, rows.length, 1).setBackground('#EBF3FB');  // Balance col tint
    sh.getRange(4, 8, rows.length, 1).setBackground('#F4F8FB');  // Reserved col tint
    // Highlight reservation rows with a soft blue banding override
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][2]).indexOf('Reservation') === 0) {
        sh.getRange(4 + i, 1, 1, headers.length).setBackground('#F4F8FB').setFontStyle('italic');
      }
    }
  }
}

function _renderSupplierLedger(sh, ss, supCode) {
  supCode = extractCodePrefix(supCode);
  if (!supCode) return;
  const sup = getRowByCode(ss.getSheetByName('Supplier Master'), supCode);
  if (!sup) { sh.getRange('A3').setValue('Supplier not found.').setFontStyle('italic'); return; }
  _explorerBanner(sh, 'Supplier', supCode + ' — ' + sup[1] + (sup[6] ? '  |  GST: ' + sup[6] : ''), '#1F4E79');

  const headers = ['Sr','Date','GRN No','Item Code','Description','On Hand','Qty','UOM','Rate','Total','Invoice'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  const onHandMap = _onHandMap(ss);
  const rows = _fetchRegisterRowsCustom(ss, 'Inward Register', 6, supCode, function(r) {
    const code = r[3];
    return [r[0], r[2], r[1], code, r[4], (onHandMap[code] != null ? onHandMap[code] : ''), r[7], r[8], r[9], r[14], r[15]];
  }).sort(function(a, b) { return (a[1] < b[1]) ? 1 : -1; });
  if (rows.length) sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows.length),
    widths: [50, 95, 105, 105, 240, 90, 70, 60, 90, 100, 130], freezeRows: 3,
    dateCols: [2], centerCols: [1, 8], qtyCols: [6, 7], moneyCols: [9, 10] });
}

function _renderCategoryStock(sh, ss, category) {
  if (!category) return;
  const snap = ss.getSheetByName('Stock Snapshot');
  const data = _readRmSection(snap, 17);
  const rate = _canonicalAvgRateMap(ss);  // master col U wins, else inward-derived
  // v3.6 RM idx: 0=Code, 1=Desc, 2=Cat, 3=Storage, 8=Available, 9=Reserved,
  //              10=MIN, 11=MAX, 12=Value, 16=BelowMIN
  const rows = data.filter(function(r) { return r[0] && r[2] === category; })
    .map(function(r) {
      const avail = Number(r[8]) || 0;
      const unit = Number(rate[r[0]] || 0);
      return [r[0], r[1], r[2], r[3], avail, r[9], unit, r[10], r[11], r[12], r[16] ? 'YES' : ''];
    });
  _explorerBanner(sh, 'Category', category + '  |  ' + rows.length + ' item(s)', '#1F4E79');

  const headers = ['Item Code','Description','Category','Storage','Available','Reserved','Avg Rate','MIN','MAX','Stock Value (Rs.)','Below MIN?'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    // ponytail: set Storage col to text format BEFORE setValues. Order matters —
    // if Sheets receives "4/20" into an unformatted cell it auto-parses as a
    // date (serial 46132). Formatting after the write only cosmetically changes
    // the already-corrupted number.
    sh.getRange(4, 4, rows.length, 1).setNumberFormat('@');
    sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  }
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows.length),
    widths: [110, 220, 160, 95, 85, 85, 95, 65, 65, 125, 85], freezeRows: 3,
    qtyCols: [5, 6, 8, 9], moneyCols: [7, 10], centerCols: [11] });
  if (rows.length) {
    // Re-apply text format on Storage after beautify in case _beautify touched col D.
    sh.getRange(4, 4, rows.length, 1).setNumberFormat('@');
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][10] === 'YES') sh.getRange(4 + i, 1, 1, headers.length).setBackground('#FADBD8');
    }
  }
}

function _renderNonMoving(sh, ss, days) {
  days = Number(days) || Number(getConfig('Non-Moving Days Threshold')) || 115;
  const snap = ss.getSheetByName('Stock Snapshot');
  const data = _readRmSection(snap, 17);
  const rate = _canonicalAvgRateMap(ss);  // master col U wins, else inward-derived
  // v3.6 idx: 8=Available, 12=Value, 13=LastIn, 14=LastOut, 15=Days
  const rows = data
    .filter(function(r) { return r[0] && r[15] !== '' && r[15] >= days; })
    .sort(function(a, b) { return b[15] - a[15]; })
    .map(function(r) {
      const unit = Number(rate[r[0]] || 0);
      return [r[0], r[1], r[8], unit, r[14] || r[13] || '', r[15], r[12]];
    });
  _explorerBanner(sh, 'Non-Moving Threshold', '≥ ' + days + ' days idle  |  ' + rows.length + ' item(s)', '#F47C2C');

  const headers = ['Item Code','Description','Available','Avg Rate','Last Movement','Days Idle','Stock Value (Rs.)'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows.length),
    widths: [110, 240, 90, 100, 115, 90, 130], freezeRows: 3, headerColor: '#F47C2C',
    qtyCols: [3], moneyCols: [4, 7], dateCols: [5], intCols: [6] });
  if (rows.length) {
    for (let i = 0; i < rows.length; i++) {
      if (Number(rows[i][5]) >= days * 2) sh.getRange(4 + i, 1, 1, headers.length).setBackground('#FDEBD0');
    }
  }
}

function _renderLowStock(sh, ss) {
  const snap = ss.getSheetByName('Stock Snapshot');
  const rmData  = _readRmSection(snap, 17);
  const sfgData = _readSfgSection(snap, 9);
  const rows = [];
  // RM (v3.6): A=code, B=desc, I=Available(8), K=MIN(10), Q=BelowMIN(16)
  rmData.forEach(function(r) {
    if (r[0] && r[16] === true) rows.push([r[0], 'RM', r[1], r[8], r[10], (r[10] - r[8]), 'Reorder']);
  });
  // SFG: A=code, B=desc, F=Available (idx 5), G=MIN (idx 6), I=Below MIN (idx 8)
  sfgData.forEach(function(r) {
    if (r[0] && r[8] === true) rows.push([r[0], 'SFG', r[1], r[5], r[6], (r[6] - r[5]), 'Produce/Return from JW']);
  });
  _explorerBanner(sh, 'Low Stock', rows.length + ' item(s) below MIN LVL — needs attention', '#C0392B');

  const headers = ['Item/SFG Code','Type','Description','Current','MIN LVL','Shortfall','Action'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  }
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows.length),
    widths: [110, 70, 240, 90, 90, 100, 180], freezeRows: 3, headerColor: '#C0392B',
    centerCols: [2], qtyCols: [4, 5, 6], boldCols: [6] });
  if (rows.length) sh.getRange(4, 1, rows.length, headers.length).setBackground('#FADBD8');
}

function _renderPendingJobWork(sh, ss) {
  const reg = ss.getSheetByName('Job Work Outward Register');
  if (!reg || reg.getLastRow() < 2) return;
  const data = reg.getRange(2, 1, reg.getLastRow() - 1, reg.getLastColumn()).getValues();
  const today = new Date();
  const rows = data
    .filter(function(r) { return r[13] === 'Open'; })
    .map(function(r) {
      const daysOut = (r[2] instanceof Date) ? Math.floor((today - r[2]) / 86400000) : '';
      return [r[1], r[2], r[4], r[6], r[7], r[8], r[10] || '', daysOut, r[13]];
    })
    .sort(function(a, b) { return (b[7] || 0) - (a[7] || 0); });
  _explorerBanner(sh, 'Open Job Work', rows.length + ' challan(s) — sorted by days-out (overdue first)', '#F47C2C');

  const headers = ['Challan No','Date','Job Worker','Code','Description','On Hand','Qty','Expected Return','Days Out','Status'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  // ponytail: a challan line may be RM or SFG now (Change 2). Try SFG map first,
  // then fall back to the RM on-hand map so the column stays accurate for both.
  const sfgOnHand = _sfgOnHandMap(ss);
  const rmOnHand  = _onHandMap(ss);
  const rows2 = rows.map(function(r) {
    const code = r[3];
    const oh = (sfgOnHand[code] != null) ? sfgOnHand[code]
             : (rmOnHand[code]  != null ? rmOnHand[code]  : '');
    return [r[0], r[1], r[2], code, r[4], oh, r[5], r[6], r[7], r[8]];
  });
  if (rows2.length) sh.getRange(4, 1, rows2.length, headers.length).setValues(rows2);
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows2.length),
    widths: [110, 95, 160, 100, 200, 90, 70, 115, 85, 85], freezeRows: 3, headerColor: '#F47C2C',
    dateCols: [2, 8], qtyCols: [6, 7], intCols: [9], centerCols: [10] });
  if (rows2.length) {
    for (let i = 0; i < rows2.length; i++) {
      if ((Number(rows2[i][8]) || 0) > 21) sh.getRange(4 + i, 1, 1, headers.length).setBackground('#FADBD8');
    }
  }
}


// ---------- Reorder Suggestions view ----------
// Reads Stock Snapshot + Item Master. Lists items with On Hand < MIN LVL,
// suggests reorder qty = MAX(MIN_LVL*2, MAX_LVL) - On Hand (rounded up to Avg
// Daily × Lead Time × 1.2 if that's higher — captures replenish cycle).
function _renderReorderSuggestions(sh, ss) {
  const snap = ss.getSheetByName('Stock Snapshot');
  const items = ss.getSheetByName('Item Master');
  const snapData = _readRmSection(snap, 17);
  if (!snapData.length) { sh.getRange('A3').setValue('Stock Snapshot empty — run Full Rebuild.').setFontStyle('italic'); return; }
  // Item Master fields for extras
  const iLastRow = nextRowByCol(items, 1) - 1;
  const itemExtras = {};
  if (iLastRow >= 2) {
    const supCol = _colByHeader(items, 'Preferred Supplier Code');
    const avgCol = _colByHeader(items, 'Avg Daily Consumption');
    const ltCol  = _colByHeader(items, 'Lead Time (Days)');
    const width = Math.max(1, supCol || 0, avgCol || 0, ltCol || 0);
    const idata = items.getRange(2, 1, iLastRow - 1, width).getValues();
    idata.forEach(function(r) {
      if (!r[0]) return;
      itemExtras[r[0]] = {
        supplier: supCol > 0 ? (r[supCol - 1] || '') : '',
        avgDaily: avgCol > 0 ? (Number(r[avgCol - 1]) || 0) : 0,
        leadTime: ltCol > 0 ? (Number(r[ltCol - 1]) || 0) : 0,
      };
    });
  }

  const rate = _canonicalAvgRateMap(ss);
  const rows = [];
  // v3.6 idx: 8=Available, 9=Reserved, 10=MIN, 11=MAX, 16=BelowMIN
  snapData.forEach(function(r) {
    if (!r[0] || r[16] !== true) return;
    const available = Number(r[8]) || 0;
    const minLvl    = Number(r[10]) || 0;
    const maxLvl    = Number(r[11]) || 0;
    const ex = itemExtras[r[0]] || {};
    const cycleQty = ex.avgDaily * ex.leadTime * 1.2;
    const targetQty = Math.max(minLvl * 2, maxLvl, cycleQty);
    const suggested = Math.max(0, Math.ceil(targetQty - available));
    const unitRate = Number(rate[r[0]] || 0);
    const estCost = suggested * unitRate;
    rows.push([r[0], r[1], ex.supplier || '(none)', available, minLvl, maxLvl, suggested, unitRate, estCost]);
  });
  rows.sort(function(a, b) { return b[8] - a[8]; });

  const grandCost = rows.reduce(function(s, r) { return s + (r[8] || 0); }, 0);
  _explorerBanner(sh, 'Reorder', rows.length + ' item(s)  |  Est. cost: ₹ ' + _fmtIndian(grandCost) + '  |  Setup → Create Draft PO', '#1E8449');

  const headers = ['Item Code','Description','Preferred Supplier','Available','MIN','MAX','Suggested Qty','Unit Rate','Est. Cost'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows.length),
    widths: [110, 260, 140, 80, 70, 70, 110, 100, 120], freezeRows: 3, headerColor: '#1E8449',
    qtyCols: [4, 5, 6, 7], moneyCols: [8, 9], boldCols: [7, 9] });
  // Highlight rows with no preferred supplier — action needed
  if (rows.length) {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][2] === '(none)') sh.getRange(4 + i, 3).setBackground('#FADBD8').setFontColor('#900');
    }
  }
}

// Server-facing: groups Reorder rows by preferred supplier + creates one Draft
// PO per supplier. Called from the Setup menu.
function createDraftPoFromReorder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const snap = ss.getSheetByName('Stock Snapshot');
  const items = ss.getSheetByName('Item Master');
  if (!snap || snap.getLastRow() < 2) { ui.alert('Stock Snapshot empty. Run Full Stock Rebuild first.'); return; }
  const iLastRow = nextRowByCol(items, 1) - 1;
  const supCol = _colByHeader(items, 'Preferred Supplier Code');
  const avgCol = _colByHeader(items, 'Avg Daily Consumption');
  const ltCol  = _colByHeader(items, 'Lead Time (Days)');
  const width = Math.max(1, supCol || 0, avgCol || 0, ltCol || 0);
  const extras = {};
  if (iLastRow >= 2) {
    items.getRange(2, 1, iLastRow - 1, width).getValues().forEach(function(r) {
      if (!r[0]) return;
      extras[r[0]] = { supplier: supCol > 0 ? (r[supCol - 1] || '') : '',
                       avgDaily: avgCol > 0 ? (Number(r[avgCol - 1]) || 0) : 0,
                       leadTime: ltCol > 0 ? (Number(r[ltCol - 1]) || 0) : 0 };
    });
  }
  const rate = _canonicalAvgRateMap(ss);
  const snapData = _readRmSection(snap, 17);
  const bySupplier = {};
  snapData.forEach(function(r) {
    if (!r[0] || r[16] !== true) return;  // v3.6: BelowMIN at idx 16
    const ex = extras[r[0]] || {};
    if (!ex.supplier) return;  // skip items with no preferred supplier
    // Multi-supplier support: field may be 'SUPP-0001, SUPP-0002'. Auto-PO
    // routes to the FIRST supplier; user can manually route to others.
    const primarySupplier = String(ex.supplier).split(',')[0].trim();
    if (!primarySupplier) return;
    const available = Number(r[8]) || 0;
    const minLvl    = Number(r[10]) || 0;
    const maxLvl    = Number(r[11]) || 0;
    const suggested = Math.max(0, Math.ceil(Math.max(minLvl * 2, maxLvl, ex.avgDaily * ex.leadTime * 1.2) - available));
    if (suggested <= 0) return;
    (bySupplier[primarySupplier] = bySupplier[primarySupplier] || []).push({
      itemCode: r[0], qty: suggested, rate: Number(rate[r[0]] || 0), gst: 18, discount: 0, remarks: 'Auto-reorder',
    });
  });
  const suppliers = Object.keys(bySupplier);
  if (!suppliers.length) { ui.alert('No items below MIN with a Preferred Supplier set.'); return; }
  const created = [];
  const today = new Date();
  suppliers.forEach(function(supCode) {
    try {
      const res = postPoRows({
        date: today, supplierCode: supCode, deliveryDate: '',
        remarks: 'Auto-generated draft from Reorder Suggestions',
        lineItems: bySupplier[supCode],
        skipPdf: true,
      });
      created.push(supCode + ' → ' + res.poNo + ' (' + bySupplier[supCode].length + ' items)');
    } catch (e) {
      created.push(supCode + ' → FAILED: ' + e.message);
    }
  });
  ui.alert('Draft POs created:\n\n' + created.join('\n') + '\n\nOpen Purchase Order Register to review, then generate PDFs individually.');
}



// ---------- Site Consumption view ----------
// D1 = project name. Aggregates Requisition rows for that project,
// pulls avg rate from Inward Register for cost. Shows item-level totals + a
// grand total at the bottom.
function _renderSiteConsumption(sh, ss, project) {
  if (!project) return;
  const proj = String(project).trim().toLowerCase();
  const req  = ss.getSheetByName('Material Requisition Register');
  const rate = _canonicalAvgRateMap(ss);
  const agg = {}; // itemCode → { desc, qty, uom, reqs }

  function add(code, desc, qty, uom) {
    if (!code || !(qty > 0)) return;
    const key = code;
    if (!agg[key]) agg[key] = { code: code, desc: desc || '', qty: 0, uom: uom || '', reqs: 0 };
    agg[key].qty += Number(qty) || 0;
    agg[key].reqs++;
  }
  if (req && req.getLastRow() >= 2) {
    const d = req.getRange(2, 1, req.getLastRow() - 1, 15).getValues();
    d.forEach(function(r) {
      if (String(r[3] || '').trim().toLowerCase() !== proj) return;  // col D = Project
      add(r[4], r[5], r[6], r[7]);   // Item Code, Desc, Qty, UOM
    });
  }

  const rows = Object.keys(agg).map(function(k) {
    const r = agg[k];
    const unit = Number(rate[r.code] || 0);
    const value = unit * r.qty;
    return [r.code, r.desc, r.qty, r.uom, unit, value, r.reqs];
  }).sort(function(a, b) { return b[5] - a[5]; });  // biggest value first

  const totalQty = rows.reduce(function(s, r) { return s + r[2]; }, 0);
  const totalVal = rows.reduce(function(s, r) { return s + r[5]; }, 0);

  _explorerBanner(sh, 'Project', project + '  |  ' + rows.length + ' item(s)  |  Total: ₹ ' + _fmtIndian(totalVal), '#1F4E79');

  const headers = ['Item Code','Description','Qty','UOM','Avg Rate','Value (Rs.)','Requisitions'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows.length),
    widths: [110, 260, 80, 60, 95, 130, 100], freezeRows: 3,
    qtyCols: [3], moneyCols: [5, 6], intCols: [7], centerCols: [4], boldCols: [6] });
  // Totals row — placed BELOW the beautified data so it doesn't get overwritten
  if (rows.length) {
    const totalRow = 4 + rows.length;
    sh.getRange(totalRow, 1, 1, 7).setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.getRange(totalRow, 1).setValue('TOTAL');
    sh.getRange(totalRow, 3).setValue(totalQty).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
    sh.getRange(totalRow, 6).setValue(totalVal).setNumberFormat('₹#,##0.00').setHorizontalAlignment('right');
    sh.setRowHeight(totalRow, 28);
  }
}

// ---------- Project Requisitions view ----------
// Unlike Site Consumption (which aggregates per item), this lists EVERY
// requisition line for the picked project, chronologically, with Avg Rate and
// Line Value per row. Useful when you need to tally issuance against site visits.
function _renderProjectReqs(sh, ss, project) {
  if (!project) return;
  const proj = String(project).trim().toLowerCase();
  const req = ss.getSheetByName('Material Requisition Register');
  if (!req || req.getLastRow() < 2) {
    sh.getRange('A3').setValue('Requisition Register empty.').setFontStyle('italic');
    return;
  }
  const rate = _canonicalAvgRateMap(ss);
  // Req cols: A=Sr, B=ReqNo, C=Date, D=Project, E=ItemCode, F=ItemDesc,
  //           G=Qty, H=UOM, I=Purpose, J=RequestedBy, K=TakenBy, L=PDF, M=Remarks
  const data = req.getRange(2, 1, req.getLastRow() - 1, 15).getValues();
  const matched = data.filter(function(r) {
    return String(r[3] || '').trim().toLowerCase() === proj;
  });
  // Sort newest first so recent requisitions surface
  matched.sort(function(a, b) {
    const ad = (a[2] instanceof Date) ? a[2].getTime() : 0;
    const bd = (b[2] instanceof Date) ? b[2].getTime() : 0;
    return bd - ad;
  });

  const rows = matched.map(function(r, i) {
    const code = r[4] || '';
    const qty  = Number(r[6]) || 0;
    const unit = Number(rate[code] || 0);
    const value = qty * unit;
    return [i + 1, r[2], r[1], code, r[5] || '', qty, r[7] || '', unit, value, r[10] || r[9] || '', r[8] || ''];
  });

  const totalQty = rows.reduce(function(s, r) { return s + r[5]; }, 0);
  const totalVal = rows.reduce(function(s, r) { return s + r[8]; }, 0);

  _explorerBanner(sh, 'Project', project + '  |  ' + rows.length + ' line(s)  |  Total: ₹ ' + _fmtIndian(totalVal), '#1F4E79');

  const headers = ['Sr','Date','Req No','Item Code','Description','Qty','UOM','Avg Rate','Value (Rs.)','Taken By','Purpose'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows.length),
    widths: [40, 95, 105, 105, 240, 80, 60, 95, 110, 130, 220], freezeRows: 3,
    dateCols: [2], centerCols: [1, 7], qtyCols: [6], moneyCols: [8, 9], boldCols: [9] });
  // Totals row
  if (rows.length) {
    const totalRow = 4 + rows.length;
    sh.getRange(totalRow, 1, 1, headers.length).setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.getRange(totalRow, 1).setValue('TOTAL');
    sh.getRange(totalRow, 6).setValue(totalQty).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
    sh.getRange(totalRow, 9).setValue(totalVal).setNumberFormat('₹#,##0.00').setHorizontalAlignment('right');
    sh.setRowHeight(totalRow, 28);
  }
}

// ---------- QC Log view ----------
// Shows recent QC records with source/ref/verdict summary + counts by outcome.
function _renderQcLog(sh, ss) {
  const qc = ss.getSheetByName('QC Register');
  if (!qc || qc.getLastRow() < 2) {
    sh.getRange('A3').setValue('QC Register empty — no records yet.').setFontStyle('italic');
    return;
  }
  // QC cols: A=Source, B=RefNo, C=RefLine, D=Outcome, E=Vendor, F=Material,
  //   G=PO, H=Date, I=CheckedBy, ... AI=Created At
  const data = qc.getRange(2, 1, qc.getLastRow() - 1, 35).getValues();
  // Sort newest first (col AI = Created At)
  const sorted = data.filter(function(r) { return r[0]; })
    .sort(function(a, b) {
      const at = a[34] instanceof Date ? a[34].getTime() : 0;
      const bt = b[34] instanceof Date ? b[34].getTime() : 0;
      return bt - at;
    });

  // Counts by outcome for summary
  let acc = 0, rej = 0, rew = 0;
  sorted.forEach(function(r) {
    if (r[3] === 'Accept') acc++;
    else if (r[3] === 'Reject') rej++;
    else if (r[3] === 'Rework') rew++;
  });

  _explorerBanner(sh, 'QC Log', sorted.length + ' record(s)  |  ✓ Accept: ' + acc + '   ✗ Reject: ' + rej + '   ↻ Rework: ' + rew, '#1F4E79');

  const headers = ['Source','Ref No','Ref Line','Outcome','Vendor / Party','Material','PO','Date','Checked By'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  const rows = sorted.slice(0, 500).map(function(r) {
    return [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8]];
  });
  if (rows.length) sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows.length),
    widths: [110, 105, 120, 100, 180, 220, 105, 100, 140], freezeRows: 3,
    dateCols: [8], centerCols: [1, 3, 4] });
  // Colour-code Outcome column
  if (rows.length) {
    for (let i = 0; i < rows.length; i++) {
      const outcome = rows[i][3];
      const cell = sh.getRange(4 + i, 4);
      if (outcome === 'Accept') cell.setBackground('#D5F5E3').setFontColor('#1E8449').setFontWeight('bold');
      else if (outcome === 'Reject') cell.setBackground('#FADBD8').setFontColor('#C0392B').setFontWeight('bold');
      else if (outcome === 'Rework') cell.setBackground('#FDEBD0').setFontColor('#F47C2C').setFontWeight('bold');
    }
  }
}


// Aggregate FG Delivery Lines HARDWARE section: netQty = delivered - returned.
// Returns { itemCode: { qty, lastDate } } for RM outward tally.
function _aggregateFgDelHardware(ss, onlyCode) {
  const sh = ss.getSheetByName('FG Delivery Lines');
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  // Cols: A=DC No, B=Line No, C=Section, D=Item Code, E=Desc, F=Qty, G=UOM,
  //       H=Weight, I=Remark, J=Returned Qty
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  // Correlate DC No → Date via FG Delivery Register (for lastDate)
  const reg = ss.getSheetByName('FG Delivery Register');
  const dateByDc = {};
  if (reg && reg.getLastRow() >= 2) {
    const rd = reg.getRange(2, 2, reg.getLastRow() - 1, 2).getValues();  // B=DC No, C=Date
    rd.forEach(function(r) { if (r[0]) dateByDc[r[0]] = r[1]; });
  }
  data.forEach(function(r) {
    if (r[2] !== 'Hardware') return;
    const code = r[3]; if (!code) return;
    if (onlyCode && code !== onlyCode) return;
    const net = (Number(r[5]) || 0) - (Number(r[9]) || 0);
    if (net <= 0) return;
    if (!out[code]) out[code] = { qty: 0, lastDate: null };
    out[code].qty += net;
    const d = dateByDc[r[0]];
    if (d && (!out[code].lastDate || d > out[code].lastDate)) out[code].lastDate = d;
  });
  return out;
}

// ---------- On Hand lookup helpers — read Stock Snapshot once per call ----------
// Vertical layout: RM at rows 2..(sfgHeader-2), SFG data at (sfgHeader+1)..end.
function _onHandMap(ss) {
  const snap = ss.getSheetByName('Stock Snapshot');
  const out = {};
  if (!snap || snap.getLastRow() < 2) return out;
  const sfgHeaderRow = _findSfgSectionStart(snap);
  const rmEnd = sfgHeaderRow > 0 ? sfgHeaderRow - 2 : snap.getLastRow();
  if (rmEnd < 2) return out;
  const data = snap.getRange(2, 1, rmEnd - 1, 9).getValues();  // A..I → RM code + On Hand
  for (let i = 0; i < data.length; i++) if (data[i][0]) out[data[i][0]] = data[i][8];
  return out;
}
function _sfgOnHandMap(ss) {
  const snap = ss.getSheetByName('Stock Snapshot');
  const out = {};
  if (!snap) return out;
  const sfgHeaderRow = _findSfgSectionStart(snap);
  if (sfgHeaderRow < 0) return out;
  const dataStart = sfgHeaderRow + 1;
  const lastRow = snap.getLastRow();
  if (lastRow < dataStart) return out;
  const data = snap.getRange(dataStart, 1, lastRow - dataStart + 1, 4).getValues();  // A..D → SFG code + On Hand
  for (let i = 0; i < data.length; i++) if (data[i][0]) out[data[i][0]] = data[i][3];
  return out;
}

// ---------- Kitting Sheet (2-level BOM explosion) ----------
// C1 = FG picker, C2 = order qty, E2 = total kit cost (computed)
function rebuildKittingSheet(fgCode, orderQty) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Kitting Sheet');
  const snap = ss.getSheetByName('Stock Snapshot');
  if (!sh) return;
  _clearBelow(sh, 3);
  orderQty = Number(orderQty) || 1;
  const fg = getRowByCode(ss.getSheetByName('FG Master'), fgCode);
  sh.getRange('D2').setValue(fg ? (fg[2] || '') : 'FG not found');
  if (!fg) return;

  const headers = ['Level','Component','Type','Description','Qty/Kit','Qty for Order','UOM','Available','Short?','Unit Cost','Line Cost'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFF');

  const stockByCode = _stockAvailableMap(snap);
  const avgRateByCode = _canonicalAvgRateMap(ss);
  const rows = [];
  const l0 = _bomEntriesFor('FG', fgCode);
  l0.forEach(function(e) {
    const line = _bomRow(0, e, orderQty, stockByCode, avgRateByCode);
    rows.push(line);
    if (e.componentType === 'SFG') {
      // Recurse into SFG's own BOM as level 1
      const l1 = _bomEntriesFor('SFG', e.componentCode);
      // effective qty per kit at level 1 = child.qtyPer * parent.qtyPer * orderQty
      l1.forEach(function(c) {
        const scaled = { componentType: c.componentType, componentCode: c.componentCode, desc: c.desc, qtyPer: c.qtyPer * e.qtyPer, uom: c.uom };
        rows.push(_bomRow(1, scaled, orderQty, stockByCode, avgRateByCode));
      });
    }
  });

  const totalCost = rows.reduce(function(s, r) { return s + (Number(r[10]) || 0); }, 0);
  sh.getRange('F1').setValue('Total Kit Cost:').setFontWeight('bold');
  sh.getRange('G1').setValue(totalCost).setFontWeight('bold').setBackground('#F47C2C').setFontColor('#FFFFFF').setHorizontalAlignment('right');
  if (rows.length) sh.getRange(4, 1, rows.length, headers.length).setValues(rows);
  _beautify(sh, { headerRow: 3, dataStartRow: 4, dataEndRow: Math.max(3, 3 + rows.length),
    widths: [50, 110, 60, 240, 80, 90, 60, 80, 60, 90, 100], freezeRows: 3 });
  // Highlight Short=YES rows in red
  if (rows.length) {
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][8] === 'YES') sh.getRange(4 + i, 1, 1, headers.length).setBackground('#FADBD8');
    }
    // Indent level-1 (child) rows visually
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === 1) sh.getRange(4 + i, 2).setFontStyle('italic');
    }
  }
}

function _bomRow(level, e, orderQty, stockMap, rateMap) {
  const qtyOrder = e.qtyPer * orderQty;
  const avail = stockMap[e.componentCode] || 0;
  const rate = rateMap[e.componentCode] || 0;
  return [
    level, e.componentCode, e.componentType, e.desc || '',
    e.qtyPer, qtyOrder, e.uom || '',
    avail, (avail < qtyOrder ? 'YES' : ''),
    rate, rate * qtyOrder,
  ];
}

function _bomEntriesFor(parentType, parentCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bom = ss.getSheetByName('BOM Master');
  if (!bom || bom.getLastRow() < 2) return [];
  const data = bom.getRange(2, 1, bom.getLastRow() - 1, bom.getLastColumn()).getValues();
  const out = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (r[1] !== parentType || r[2] !== parentCode) continue;
    // cols: A BOM ID, B Parent Type, C Parent Code, D Component Type, E Component Code, F Qty, G UOM, H Notes
    const componentSheet = (r[3] === 'RM') ? 'Item Master' : 'SFG Master';
    const compRow = getRowByCode(ss.getSheetByName(componentSheet), r[4]);
    out.push({
      componentType: r[3], componentCode: r[4],
      // v3.5: BOM row cols A..H (unchanged); Item/SFG Master UOM at col G(6).
      qtyPer: Number(r[5]) || 0, uom: r[6] || (compRow ? compRow[6] : ''),
      desc: compRow ? (compRow[2] || '') : '',
    });
  }
  return out;
}

function _stockAvailableMap(snap) {
  const out = {};
  if (!snap || snap.getLastRow() < 2) return out;
  const sfgHeaderRow = _findSfgSectionStart(snap);
  const rmEnd = sfgHeaderRow > 0 ? sfgHeaderRow - 2 : snap.getLastRow();
  // RM section: A=Code, I=On Hand (idx 8)
  if (rmEnd >= 2) {
    const rmData = snap.getRange(2, 1, rmEnd - 1, 9).getValues();
    rmData.forEach(function(r) { if (r[0]) out[r[0]] = r[8]; });
  }
  // SFG section: A=Code, F=Available (idx 5)
  if (sfgHeaderRow > 0) {
    const lastRow = snap.getLastRow();
    if (lastRow > sfgHeaderRow) {
      const sfgData = snap.getRange(sfgHeaderRow + 1, 1, lastRow - sfgHeaderRow, 6).getValues();
      sfgData.forEach(function(r) { if (r[0]) out[r[0]] = r[5]; });
    }
  }
  return out;
}

function _avgRateMap(ss) {
  const inward = ss.getSheetByName('Inward Register');
  const out = {};
  if (!inward || inward.getLastRow() < 2) return out;
  const data = inward.getRange(2, 1, inward.getLastRow() - 1, inward.getLastColumn()).getValues();
  const sums = {};
  data.forEach(function(r) {
    const code = r[3]; if (!code) return;
    const qty = Number(r[7]) || 0, rate = Number(r[9]) || 0;
    if (!sums[code]) sums[code] = { qtyRate: 0, qty: 0 };
    sums[code].qtyRate += qty * rate;
    sums[code].qty += qty;
  });
  Object.keys(sums).forEach(function(k) { out[k] = sums[k].qty ? sums[k].qtyRate / sums[k].qty : 0; });
  return out;
}

// Canonical Avg Rate per Item Master code — matches _computeRmRowFromMaps:
// user-set Item Master col U wins; else fall back to inward-derived rate.
// Explorer views must use THIS, not _avgRateMap alone, otherwise displayed Avg
// Rate × Available won't equal the Snapshot's Stock Value column when a master
// rate has been set.
function _canonicalAvgRateMap(ss) {
  const derived = _avgRateMap(ss);
  const items = ss.getSheetByName('Item Master');
  const out = {};
  if (items && items.getLastRow() >= 2) {
    const _arIdx = (typeof _avgRateCol === 'function') ? _avgRateCol(items) - 1 : 20;
    const data = items.getRange(2, 1, items.getLastRow() - 1, items.getLastColumn()).getValues();
    data.forEach(function(r) {
      if (!r[0]) return;
      const master = (_arIdx >= 0) ? (Number(r[_arIdx]) || 0) : 0;
      out[r[0]] = master > 0 ? master : (Number(derived[r[0]]) || 0);
    });
  }
  // Codes present in inward but not in master still get their derived rate.
  Object.keys(derived).forEach(function(k) { if (out[k] == null) out[k] = Number(derived[k]) || 0; });
  return out;
}

// ---- Viewer helpers ----
function _clearBelow(sh, startRow) {
  if (sh.getLastRow() >= startRow) {
    const range = sh.getRange(startRow, 1, sh.getLastRow() - startRow + 1, sh.getLastColumn());
    range.clearContent().clearFormat();
  }
  // Remove existing bandings so re-runs don't stack them
  sh.getBandings().forEach(function(b) { try { b.remove(); } catch(e){} });
}

// Beautify a data range: apply column widths, header style, borders, banding, freeze.
// opts: { headerRow, dataStartRow, dataEndRow, widths:[..], headerColor?, freezeRows? }
function _beautify(sh, opts) {
  const numCols = opts.widths.length;
  opts.widths.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  // Header row styling — uppercase, letter-spacing via font weight, colored bar
  const headerColor = opts.headerColor || '#1F4E79';
  const headerRange = sh.getRange(opts.headerRow, 1, 1, numCols);
  headerRange.setFontWeight('bold')
    .setBackground(headerColor)
    .setFontColor('#FFFFFF')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, false, false, headerColor, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.setRowHeight(opts.headerRow, 30);

  // Data range styling
  if (opts.dataEndRow >= opts.dataStartRow) {
    const rowCount = opts.dataEndRow - opts.dataStartRow + 1;
    const dataRange = sh.getRange(opts.dataStartRow, 1, rowCount, numCols);
    dataRange.setBorder(true, true, true, true, true, true, '#E5E8EC', SpreadsheetApp.BorderStyle.SOLID)
             .setVerticalAlignment('middle')
             .setFontFamily('-apple-system,Segoe UI,Roboto,Arial').setFontSize(10);
    // Row banding (idempotent — drop existing first to avoid stacking)
    try {
      dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    } catch (e) {}
    // Column number formats
    (opts.moneyCols || []).forEach(function(c) {
      sh.getRange(opts.dataStartRow, c, rowCount, 1)
        .setNumberFormat('₹#,##0.00;[Red]-₹#,##0.00;""')
        .setHorizontalAlignment('right');
    });
    (opts.qtyCols || []).forEach(function(c) {
      sh.getRange(opts.dataStartRow, c, rowCount, 1)
        .setNumberFormat('#,##0.00;[Red]-#,##0.00;""')
        .setHorizontalAlignment('right');
    });
    (opts.intCols || []).forEach(function(c) {
      sh.getRange(opts.dataStartRow, c, rowCount, 1)
        .setNumberFormat('#,##0;[Red]-#,##0;""')
        .setHorizontalAlignment('right');
    });
    (opts.dateCols || []).forEach(function(c) {
      sh.getRange(opts.dataStartRow, c, rowCount, 1)
        .setNumberFormat('dd-mmm-yyyy')
        .setHorizontalAlignment('center');
    });
    (opts.centerCols || []).forEach(function(c) {
      sh.getRange(opts.dataStartRow, c, rowCount, 1).setHorizontalAlignment('center');
    });
    (opts.boldCols || []).forEach(function(c) {
      sh.getRange(opts.dataStartRow, c, rowCount, 1).setFontWeight('bold');
    });
    // Header alignment mirrors data alignment for numeric/center cols
    (opts.moneyCols || []).concat(opts.qtyCols || []).concat(opts.intCols || []).forEach(function(c) {
      sh.getRange(opts.headerRow, c).setHorizontalAlignment('right');
    });
    (opts.centerCols || []).concat(opts.dateCols || []).forEach(function(c) {
      sh.getRange(opts.headerRow, c).setHorizontalAlignment('center');
    });
  }
  if (opts.freezeRows) sh.setFrozenRows(opts.freezeRows);
}

// Consistent banner strip for Explorer row 2 — label chip + description chip.
// Called by every view. Replaces the ad-hoc plain-text banner.
function _explorerBanner(sh, label, value, accentColor) {
  const bg = accentColor || '#1F4E79';
  const range = sh.getRange('A2:I2').clearContent().clearFormat();
  range.merge();
  sh.getRange('A2').setValue('  ' + label + '   ›   ' + value)
    .setBackground(bg).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(2, 30);
}

// Fetch register rows where col=codeCol matches value, apply custom row-transform.
function _fetchRegisterRowsCustom(ss, sheetName, codeCol, value, transform) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const out = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][codeCol - 1] !== value) continue;
    out.push(transform(data[i]));
  }
  return out;
}

// Convenience: project 1-indexed columns from matching rows (thin wrapper).
function _fetchRegisterRows(ss, sheetName, codeCol, value, projectCols) {
  return _fetchRegisterRowsCustom(ss, sheetName, codeCol, value, function(r) {
    return projectCols.map(function(c) { return r[c - 1]; });
  });
}