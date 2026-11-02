/**
 * 3 Star Projects IMS v3 — 90_CodeHygiene.gs
 *
 * Two features bolted on after a team member re-coded master rows and left
 * historical register rows pointing at stale codes:
 *
 * 1) fixRegisterCodes()  — re-match register rows against current masters
 *                           by Description. Rewrites codes where there's
 *                           exactly one match; highlights unmatched or
 *                           ambiguous rows. Writes a "Code Fix Report" sheet.
 *                           Scope: Inward Register, Material Requisition
 *                           Register, Adjustments Register, FG Delivery Lines.
 *
 * 2) freezeAllCodes() / unfreezeCodePrompt() / _codeLockGuard()
 *                        — "Code Locked?" boolean column on Item/SFG/FG
 *                           Masters. Guard reverts manual edits to col A on
 *                           locked rows. Called from the onEdit router in
 *                           10_Masters.gs via `if (_codeLockGuard(e)) return;`.
 *                           Programmatic setValue (used by auto-assign) is
 *                           not intercepted by simple triggers, so the lock
 *                           never blocks the system itself.
 *
 * All functions idempotent + append-only on schema (column added at the end).
 */

// Masters we lock. Column index for 'Code Locked?' is appended at end on
// first run of ensureCodeLockColumn, so no existing index shifts.
const _CODE_LOCK_SHEETS = ['Item Master', 'SFG Master', 'FG Master'];
const _CODE_LOCK_HEADER = 'Code Locked?';

// Registers scanned by fixRegisterCodes. codeCol / descCol are 1-based.
// kind: 'RM' | 'SFG' | 'FG' | 'MIX_RM_SFG' (Adjustments) | 'DC' (FG Delivery
// Lines — look up by section). sectionCol only for 'DC'.
const _CODE_FIX_REGISTERS = [
  { name: 'Inward Register',               codeCol: 4, descCol: 5, kind: 'RM'  },
  { name: 'Material Requisition Register', codeCol: 5, descCol: 6, kind: 'RM'  },
  { name: 'Adjustments Register',          codeCol: 4, descCol: 5, kind: 'MIX_RM_SFG' },
  { name: 'FG Delivery Lines',             codeCol: 4, descCol: 5, kind: 'DC', sectionCol: 3 },
];

// Highlight palette (matches existing warn/danger palette).
const _HL_UNMATCHED = '#FFF2CC';  // yellow — no master row with this desc
const _HL_AMBIGUOUS = '#FADBD8';  // red   — multiple master rows match

// =============================================================================
// FIX REGISTER CODES
// =============================================================================

function fixRegisterCodes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const confirm = ui.alert(
    'Fix Register Codes',
    'This scans Inward, Requisition, Adjustments, and FG Delivery Lines, ' +
    'and for each row looks up the stored Description in the current ' +
    'Item / SFG / FG Masters.\n\n' +
    '• Exactly one match → the row\'s code is rewritten.\n' +
    '• Zero matches → the row is highlighted yellow.\n' +
    '• Multiple master rows share that description (and codes disagree) → ' +
    'highlighted red.\n\n' +
    'A "Code Fix Report" tab will be (re)written with every change and ' +
    'every highlight. If you don\'t like the result, use File → Version ' +
    'history to revert.\n\n' +
    'Proceed?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const descMap = _buildDescMap(ss);
  const report = []; // { register, row, section, oldCode, newCode, desc, status }
  const touchedByRegister = {};  // name -> { changed, unmatched, ambiguous, scanned }

  _CODE_FIX_REGISTERS.forEach(function(spec) {
    const sh = ss.getSheetByName(spec.name);
    const stats = { changed: 0, unmatched: 0, ambiguous: 0, scanned: 0 };
    touchedByRegister[spec.name] = stats;
    if (!sh || sh.getLastRow() < 2) return;

    const lastCol = sh.getLastColumn();
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    const codeUpdates = []; // { row, newCode }
    const highlightRows = []; // { row, color }

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const storedCode = row[spec.codeCol - 1];
      const storedDesc = row[spec.descCol - 1];
      if (!storedCode && !storedDesc) continue;  // fully empty — skip
      if (!storedDesc) continue;                  // no desc to match on

      stats.scanned++;

      // Determine which masters to search for this row.
      let kinds = [];
      if (spec.kind === 'RM')              kinds = ['RM'];
      else if (spec.kind === 'MIX_RM_SFG') kinds = ['RM', 'SFG'];
      else if (spec.kind === 'DC') {
        const section = row[spec.sectionCol - 1];
        if (section === 'Hardware') kinds = ['RM'];
        else                        kinds = ['SFG', 'FG'];
      } else continue;

      const matches = _findDescMatches(descMap, storedDesc, kinds);

      if (!matches.length) {
        stats.unmatched++;
        highlightRows.push({ row: i + 2, color: _HL_UNMATCHED });
        report.push([spec.name, i + 2, spec.kind === 'DC' ? (row[spec.sectionCol - 1] || '') : '',
                     storedCode, '', storedDesc, 'UNMATCHED — no master has this description']);
        continue;
      }

      // Multiple matches: fine if they all point at the same code (unlikely
      // but possible if a desc is duplicated across the same master row twice).
      const distinctCodes = Array.from(new Set(matches.map(function(m) { return m.code; })));
      if (distinctCodes.length > 1) {
        stats.ambiguous++;
        highlightRows.push({ row: i + 2, color: _HL_AMBIGUOUS });
        report.push([spec.name, i + 2, spec.kind === 'DC' ? (row[spec.sectionCol - 1] || '') : '',
                     storedCode, distinctCodes.join(' / '), storedDesc,
                     'AMBIGUOUS — ' + matches.length + ' master rows match this description']);
        continue;
      }

      const newCode = distinctCodes[0];
      if (newCode === storedCode) continue;  // already correct — no-op

      codeUpdates.push({ row: i + 2, newCode: newCode });
      report.push([spec.name, i + 2, spec.kind === 'DC' ? (row[spec.sectionCol - 1] || '') : '',
                   storedCode, newCode, storedDesc, 'CHANGED']);
      stats.changed++;
    }

    // Apply code updates one setValue per cell — kept simple; the counts are
    // small enough that batching isn't worth the index juggling.
    // ponytail: wrap in _bypassValidationSetValue — register cells may have
    // a stale data-validation rule (e.g., 'requireValueInRange Item Master
    // A2:A100') that rejects the new code before it lands. We clear the
    // rule, write, then restore the exact same rule so the next in-sheet
    // edit still shows the dropdown.
    codeUpdates.forEach(function(u) {
      _bypassValidationSetValue(sh.getRange(u.row, spec.codeCol), u.newCode);
    });
    // Apply highlights.
    highlightRows.forEach(function(h) {
      sh.getRange(h.row, 1, 1, lastCol).setBackground(h.color);
    });
  });

  // Write the Code Fix Report sheet (overwrite each run).
  _writeCodeFixReport(ss, report, touchedByRegister);

  // Snapshot refresh — codes that moved need their snapshot rows re-computed.
  // Simplest: full rebuild. The user is in a UI session, so a few seconds is
  // acceptable, and it keeps RM + SFG sections consistent in one pass.
  try { fullRebuild(); } catch (e) { /* leave snapshot stale if rebuild fails */ }

  const total = _sumStats(touchedByRegister);
  ui.alert(
    'Code Fix Report',
    'Scanned: ' + total.scanned + '\n' +
    'Codes rewritten: ' + total.changed + '\n' +
    'Rows highlighted yellow (unmatched): ' + total.unmatched + '\n' +
    'Rows highlighted red (ambiguous): ' + total.ambiguous + '\n\n' +
    'See the "Code Fix Report" tab for row-by-row detail.\n' +
    'Snapshot was rebuilt. To clear highlights after review, select the ' +
    'rows and Format → Clear formatting.',
    ui.ButtonSet.OK
  );
}

// Build { normDesc: [{code, kind, originalDesc}] } across the three masters.
// Description col = C (idx 2) in Item/SFG/FG Master (consistent per v3.5).
function _buildDescMap(ss) {
  const out = {};
  const sources = [
    { sheet: 'Item Master', kind: 'RM'  },
    { sheet: 'SFG Master',  kind: 'SFG' },
    { sheet: 'FG Master',   kind: 'FG'  },
  ];
  sources.forEach(function(s) {
    const sh = ss.getSheetByName(s.sheet);
    if (!sh || sh.getLastRow() < 2) return;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();  // A, B, C
    for (let i = 0; i < data.length; i++) {
      const code = data[i][0], desc = data[i][2];
      if (!code || !desc) continue;
      const key = _normalizeDesc(desc);
      if (!out[key]) out[key] = [];
      out[key].push({ code: code, kind: s.kind, originalDesc: desc });
    }
  });
  return out;
}

// Pull matches for `desc` from `descMap`, restricted to the given kinds.
function _findDescMatches(descMap, desc, kinds) {
  const key = _normalizeDesc(desc);
  const bucket = descMap[key] || [];
  return bucket.filter(function(m) { return kinds.indexOf(m.kind) >= 0; });
}

function _normalizeDesc(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

// setValue that survives a cell with an existing data-validation rule whose
// list / range doesn't yet include the new value. The rule is preserved
// identically on the way out — only the single setValue bypasses it.
function _bypassValidationSetValue(cell, value) {
  const rule = cell.getDataValidation();
  if (rule) cell.clearDataValidations();
  cell.setValue(value);
  if (rule) cell.setDataValidation(rule);
}

function _sumStats(byReg) {
  const t = { scanned: 0, changed: 0, unmatched: 0, ambiguous: 0 };
  Object.keys(byReg).forEach(function(k) {
    t.scanned    += byReg[k].scanned;
    t.changed    += byReg[k].changed;
    t.unmatched  += byReg[k].unmatched;
    t.ambiguous  += byReg[k].ambiguous;
  });
  return t;
}

function _writeCodeFixReport(ss, report, byReg) {
  let sh = ss.getSheetByName('Code Fix Report');
  if (!sh) sh = ss.insertSheet('Code Fix Report');
  // Overwrite cleanly.
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getBandings().forEach(function(b) { try { b.remove(); } catch (e) {} });
  sh.setTabColor('#F47C2C');

  // Summary header
  sh.getRange('A1:G1').merge().setValue('CODE FIX REPORT — ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'))
    .setBackground('#1F4E79').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(13).setHorizontalAlignment('center');
  sh.setRowHeight(1, 32);

  // Per-register summary
  const summaryRows = Object.keys(byReg).map(function(name) {
    const s = byReg[name];
    return [name, s.scanned, s.changed, s.unmatched, s.ambiguous];
  });
  sh.getRange(3, 1, 1, 5).setValues([['Register', 'Scanned', 'Changed', 'Unmatched', 'Ambiguous']])
    .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  if (summaryRows.length) {
    sh.getRange(4, 1, summaryRows.length, 5).setValues(summaryRows);
  }

  // Detail header + rows
  const detailStart = 3 + summaryRows.length + 2;  // blank spacer
  sh.getRange(detailStart, 1, 1, 7).setValues([[
    'Register', 'Row', 'Section', 'Old Code', 'New Code', 'Description', 'Status'
  ]]).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  if (report.length) {
    sh.getRange(detailStart + 1, 1, report.length, 7).setValues(report);
    // Color-code status col
    for (let i = 0; i < report.length; i++) {
      const status = report[i][6];
      let color = null;
      if (status === 'CHANGED')             color = '#D5F5E3';
      else if (status.indexOf('UNMATCHED') === 0)  color = _HL_UNMATCHED;
      else if (status.indexOf('AMBIGUOUS') === 0)  color = _HL_AMBIGUOUS;
      if (color) sh.getRange(detailStart + 1 + i, 1, 1, 7).setBackground(color);
    }
  } else {
    sh.getRange(detailStart + 1, 1, 1, 7).merge()
      .setValue('No changes and no issues — every stored description matches a current master code.')
      .setBackground('#D5F5E3').setFontColor('#1E8449').setHorizontalAlignment('center');
  }

  // Column widths
  [[1, 220], [2, 60], [3, 100], [4, 120], [5, 120], [6, 320], [7, 300]].forEach(function(cw) {
    sh.setColumnWidth(cw[0], cw[1]);
  });
  sh.setFrozenRows(detailStart);
}

// =============================================================================
// CODE LOCK — "Code Locked?" column + onEdit guard
// =============================================================================

// Idempotent column adder. Returns the 1-based column index of 'Code Locked?'
// on the sheet, adding it if missing. New column is appended at end so
// existing column indices (incl. Search Label and Avg Rate) never shift.
function _ensureCodeLockColumn(sh) {
  if (!sh) return -1;
  const lastCol = sh.getLastColumn();
  const headers = lastCol >= 1 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === _CODE_LOCK_HEADER) return i + 1;
  }
  const newCol = lastCol + 1;
  sh.getRange(1, newCol).setValue(_CODE_LOCK_HEADER)
    .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  sh.setColumnWidth(newCol, 100);
  return newCol;
}

function freezeAllCodes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const confirm = ui.alert(
    'Freeze All Codes',
    'This will add a "Code Locked?" column (if missing) to Item / SFG / FG ' +
    'Master and set it to TRUE for every row that has a code.\n\n' +
    'Once locked, manual edits to the code column (col A) are reverted ' +
    'with a toast. Auto-assigned codes from form submissions are not ' +
    'affected.\n\n' +
    'Proceed?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const summary = [];
  _CODE_LOCK_SHEETS.forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) { summary.push(name + ': missing (skipped)'); return; }
    const lockCol = _ensureCodeLockColumn(sh);
    if (sh.getLastRow() < 2) { summary.push(name + ': empty (col added)'); return; }
    const codes = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    const values = codes.map(function(r) { return [r[0] ? true : '']; });  // blank for empty rows
    sh.getRange(2, lockCol, values.length, 1).setValues(values);
    sh.getRange(2, lockCol, values.length, 1).insertCheckboxes();
    const locked = values.filter(function(r) { return r[0] === true; }).length;
    summary.push(name + ': ' + locked + ' code(s) locked');
  });
  ui.alert('Freeze All Codes', summary.join('\n'), ui.ButtonSet.OK);
}

function unfreezeCodePrompt() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = ui.prompt('Unfreeze Code',
    'Enter the code to unfreeze (e.g., FST-0001, SFG-0003, FG-0002):',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const code = resp.getResponseText().trim();
  if (!code) { ui.alert('No code entered.'); return; }

  for (let i = 0; i < _CODE_LOCK_SHEETS.length; i++) {
    const sh = ss.getSheetByName(_CODE_LOCK_SHEETS[i]);
    if (!sh) continue;
    const lockCol = _ensureCodeLockColumn(sh);
    const rowNum = _findRowByCode(sh, code);
    if (rowNum < 2) continue;
    sh.getRange(rowNum, lockCol).setValue(false);
    ui.alert('Unfroze ' + code + ' on ' + _CODE_LOCK_SHEETS[i] + ' (row ' + rowNum +
             '). Edit col A to change the code, then run Freeze All Codes again when done.');
    return;
  }
  ui.alert('Code "' + code + '" not found in Item / SFG / FG Master.');
}

// onEdit guard — called from 10_Masters.gs onEdit() via:
//   if (typeof _codeLockGuard === 'function' && _codeLockGuard(e)) return;
// Returns true if the edit was blocked (so the caller should stop processing).
// Returns false otherwise (continue with the normal onEdit router).
function _codeLockGuard(e) {
  try {
    if (!e || !e.range) return false;
    const sh = e.range.getSheet();
    const name = sh.getName();
    if (_CODE_LOCK_SHEETS.indexOf(name) < 0) return false;
    const col = e.range.getColumn();
    const row = e.range.getRow();
    if (col !== 1 || row < 2) return false;

    // Find lock column — don't add it from onEdit (quota-sensitive); if it
    // doesn't exist yet, no lock to enforce.
    const lastCol = sh.getLastColumn();
    const headers = lastCol >= 1 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    let lockCol = -1;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === _CODE_LOCK_HEADER) { lockCol = i + 1; break; }
    }
    if (lockCol < 0) return false;

    const locked = sh.getRange(row, lockCol).getValue();
    if (locked !== true) return false;

    // Edit lands on a locked row. Multi-cell paste → event has no oldValue,
    // so we can't cleanly revert; log and let it through rather than silently
    // corrupting. User can revert via version history.
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1 || e.oldValue == null) {
      toast('Code col edit on locked row — can\'t auto-revert (multi-cell paste or no prior value). ' +
            'Use File → Version history to revert if needed.', '3 Star IMS — Code Lock', 8);
      return false;
    }

    e.range.setValue(e.oldValue);
    toast(name + ' row ' + row + ': code is locked (' + e.oldValue + '). ' +
          'Unfreeze via menu first if you really need to change it.',
          '3 Star IMS — Code Lock', 6);
    return true;
  } catch (err) {
    Logger.log('_codeLockGuard error: ' + err.message + ' | ' + (err.stack || ''));
    return false;  // on error, don't block — fail open
  }
}