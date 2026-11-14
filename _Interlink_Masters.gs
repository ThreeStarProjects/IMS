/**
 * _Interlink_Masters.gs
 * -----------------------------------------------------------------------------
 * Thin wrapper around the Masters library (added as `Masters` dependency in
 * this script's Libraries). Caches reads in CacheService for 5 min so repeated
 * dropdown reads don't thrash the library.
 *
 * Paste this file verbatim into Finance / IMS / PM. It's self-contained —
 * no touches to any existing production code.
 *
 * One-time setup per system (do once after pasting this file):
 *   1. Script editor → Libraries → Add library
 *      Script ID: 1zWL2xvoZIZ1HAk-RrngKIeA4Ps4fdByYW9_hsMEm8lMPA2P8rOgTzSxC
 *      Version:   1   (or latest)
 *      Identifier: Masters
 *   2. Run IM_installInterlinkMenu() once from the editor — adds the
 *      🔗 Interlink menu to the sheet (via installable onOpen trigger, so
 *      the system's existing onOpen stays untouched).
 */

var IM_VERSION    = '0.4.0';
var IM_CACHE_TTL  = 300;   // 5 min
var IM_CACHE_KEY  = 'IM_v' + IM_VERSION + '_';

// -----------------------------------------------------------------------------
// Reads (cached)
// -----------------------------------------------------------------------------
function IM_getProjects()  { return IM_cache_('projects',  function () { return Masters.getProjects(); }); }
function IM_getCustomers() { return IM_cache_('customers', function () { return Masters.getCustomers(); }); }
function IM_getEmployees() { return IM_cache_('employees', function () { return Masters.getEmployees(); }); }
function IM_getSites()     { return IM_cache_('sites',     function () { return Masters.getSites(); }); }

/**
 * Parties with role flags. Call with a flag object to narrow:
 *   IM_getParties({ isSupplier:  true })  → suppliers only
 *   IM_getParties({ isJobWorker: true })  → job workers only
 *   IM_getParties({})                     → all parties
 */
function IM_getParties(filter) {
  var key = 'parties_' + JSON.stringify(filter || {});
  return IM_cache_(key, function () { return Masters.getParties(filter || {}); });
}

/** Freeform name → Masters record. For migration-time resolution. */
function IM_resolveName(entity, text) {
  return Masters.resolveFreeformName(entity, text);
}

// -----------------------------------------------------------------------------
// Dropdown-friendly label lists — "{ID} — {Name}"
// -----------------------------------------------------------------------------
function IM_projectLabels()  { return IM_asLabels_(IM_getProjects(),  'Project_ID',  'Name'); }
function IM_customerLabels() { return IM_asLabels_(IM_getCustomers(), 'Customer_ID', 'Name'); }
function IM_partyLabels(f)   { return IM_asLabels_(IM_getParties(f),  'Party_ID',    'Name'); }
function IM_employeeLabels() { return IM_asLabels_(IM_getEmployees(), 'Emp_ID',      'Name'); }
function IM_siteLabels()     { return IM_asLabels_(IM_getSites(),     'Site_ID',     'Name'); }

function IM_asLabels_(rows, idCol, nameCol) {
  return (rows || [])
    .filter(function (r) { return String(r['Active?'] || 'Yes').toLowerCase() !== 'no'; })
    .map(function (r) { return r[idCol] + ' — ' + r[nameCol]; });
}

/**
 * Customers for the delivery-challan picker: active ones, A to Z, with the details that fill Bill To / Ship To.
 * Finance is the master for customers (Masters holds a copy), so this only lists; typing a new name is still allowed.
 */
function IM_customerPicks_() {
  var t = function (v) { return String(v == null ? '' : v).trim(); };
  return (IM_getCustomers() || [])
    .filter(function (r) { return t(r.Name) && String(r['Active?'] || 'Yes').toLowerCase() !== 'no'; })
    .map(function (r) { return { name: t(r.Name), legal: t(r.Legal_Name), gstin: t(r.GSTIN), bill: t(r.Billing_Address), ship: t(r.Shipping_Address), terms: t(r.Payment_Terms_Days) }; })
    .sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : a.name.toLowerCase() > b.name.toLowerCase() ? 1 : 0; });
}

// Just the IDs (for lookups where you don't want the label)
function IM_projectIds()  { return IM_getProjects().map(function (r)  { return r.Project_ID;  }); }
function IM_customerIds() { return IM_getCustomers().map(function (r) { return r.Customer_ID; }); }

// -----------------------------------------------------------------------------
// Cache plumbing
// -----------------------------------------------------------------------------
function IM_cache_(key, factory) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(IM_CACHE_KEY + key);
  if (hit) {
    try { return JSON.parse(hit); } catch (_) {}
  }
  var fresh = factory();
  try {
    var str = JSON.stringify(fresh);
    if (str.length < 95000) cache.put(IM_CACHE_KEY + key, str, IM_CACHE_TTL);
  } catch (_) {}
  return fresh;
}

function IM_clearCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll([
    IM_CACHE_KEY + 'projects',
    IM_CACHE_KEY + 'customers',
    IM_CACHE_KEY + 'employees',
    IM_CACHE_KEY + 'sites',
    IM_CACHE_KEY + 'parties_{}',
    IM_CACHE_KEY + 'parties_{"isSupplier":true}',
    IM_CACHE_KEY + 'parties_{"isJobWorker":true}',
    IM_CACHE_KEY + 'parties_{"isTransporter":true}',
    IM_CACHE_KEY + 'parties_{"isServiceProv":true}'
  ]);
  SpreadsheetApp.getActive().toast('Masters cache cleared', '🔗 Interlink', 3);
}

// -----------------------------------------------------------------------------
// Menu — attached via installable onOpen trigger so the system's existing
// simple onOpen stays untouched. Run IM_installInterlinkMenu() ONCE.
// -----------------------------------------------------------------------------
function IM_installInterlinkMenu() {
  var ss = SpreadsheetApp.getActive();
  // De-dup: remove any existing IM_onOpen_ trigger first
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'IM_onOpen_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('IM_onOpen_').forSpreadsheet(ss).onOpen().create();
  // Also run once now so the menu appears without reopening
  IM_onOpen_();
  SpreadsheetApp.getUi().alert(
    'Interlink menu installed',
    'A 🔗 Interlink menu will appear every time this sheet is opened. ' +
    'The system\'s existing menu is unchanged.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function IM_onOpen_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔗 Interlink')
    .addItem('Refresh dropdowns from Masters', 'IM_refreshDropdowns')
    .addItem('➕ Add employee',                'IM_addEmployeeDialog')
    .addItem('Send suppliers and job workers to Masters', 'IM_syncPartiesNow')
    .addSeparator()
    .addItem('Process PM reservation requests now', 'IM_processReservationRequestsNow')
    .addSeparator()
    .addItem('❓ Start here',                  'IM_showStartHere')
    .addItem('About',                          'IM_about_')
    .addToUi();
  // Self-heal the PM reservation queue: recreate its trigger if missing, and drain anything waiting.
  try { IM_ensureQueueTrigger_(); IM_processReservationRequests(); } catch (e) {}
  IM_refreshDropdowns(true);  // keep PRJ-### project suggestions current (never throws when quiet)
  try {                       // Supplier Master / Job Worker Master are a copy of Masters' Parties (an editor also sends rows Masters lacks)
    var pr = IM_syncPartiesToIms_();
    if (pr.error || (pr.supplier && pr.supplier.onlyHere.length && pr.linkBlocked) || (pr.jobworker && pr.jobworker.onlyHere.length && pr.linkBlocked)) SpreadsheetApp.getActive().toast(IM_partySummary_(pr), '🔗 Interlink', 12);
  } catch (e) {}
  try { Masters.registerSystemSheet('ims', SpreadsheetApp.getActive().getId()); } catch (e) {}  // lets PM's project delete find this sheet
}

function IM_about_() {
  var pv = '?';
  try { pv = Masters.getVersion(); } catch (e) {}
  SpreadsheetApp.getUi().alert(
    '🔗 Interlink v' + IM_VERSION,
    'This system reads master data (Projects, Customers, Parties, Employees, ' +
    'Sites) from the Masters library.\n\n' +
    'Masters version: ' + pv + '\n' +
    'Cache TTL: ' + IM_CACHE_TTL + ' sec',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// -----------------------------------------------------------------------------
// Dropdown refresher (v0.4 Phase 1b, IMS). Keeps the active Masters project labels
// ("PRJ-001 — Test") in the Projects list on the Dropdowns sheet (col K), which feeds the
// project suggestions on the requisition and reservation forms. Without it nobody picks a
// PRJ-### when issuing material, and the PM reservation / Finance cost links cannot match.
// The list becomes: the plain names typed by hand (kept, in their order) + one label per active
// Masters project. So new projects appear, a renamed project gets its new label, and the labels
// of deleted / finished projects disappear. If Masters returns no active project at all, nothing
// is removed. Runs when IMS opens (quiet) and from 🔗 Interlink → Refresh dropdowns.
// -----------------------------------------------------------------------------
function IM_refreshDropdowns(quiet) {
  try {
    var ss = SpreadsheetApp.getActive();
    var dd = ss.getSheetByName('Dropdowns');
    if (!dd) return;
    if (!quiet) CacheService.getScriptCache().remove(IM_CACHE_KEY + 'projects');  // manual refresh: skip the 5-min cache
    CacheService.getScriptCache().remove(IM_CACHE_KEY + 'employees');              // people are always read fresh: someone added in Masters a minute ago must show up now

    var col = dd.getLastRow() >= 2 ? dd.getRange(2, 11, dd.getLastRow() - 1, 1).getValues() : [];
    var LABEL = /^PRJ-\d+ — /;
    var labels = IM_projectLabels(), cur = [], plain = [];
    col.forEach(function (r) {
      var v = String(r[0]);
      if (v === '') return;
      cur.push(v);
      if (!LABEL.test(v) || !labels.length) plain.push(v);   // typed-by-hand names stay; labels are rebuilt from Masters
    });
    var next = plain.concat(labels.filter(function (l) { return plain.indexOf(l) < 0; }));
    var changed = next.join('\u0001') !== cur.join('\u0001');
    if (changed) {
      var n = Math.max(next.length, col.length), out = [];
      for (var i = 0; i < n; i++) out.push([i < next.length ? next[i] : '']);
      dd.getRange(2, 11, n, 1).setValues(out);
    }
    var addedEmp = IM_refreshEmployeeDropdown_(dd);
    var partyMsg = quiet ? '' : IM_partySummary_(IM_syncPartiesToIms_(true));   // manual refresh: suppliers and job workers too
    if (!quiet) ss.toast((partyMsg ? partyMsg + ' ' : '') +(changed ? 'Project list updated from Masters (' + labels.length + ' project(s)).' : 'Projects already up to date.') +
      (addedEmp ? ' Employee list refreshed from Masters (' + addedEmp + ' people' + (IM_EMP_NEW_.length ? '; new: ' + IM_EMP_NEW_.slice(0, 5).join(', ') + (IM_EMP_NEW_.length > 5 ? ', …' : '') : '') + ').' : ' Employee list not changed (Masters returned no employees).'), '🔗 Interlink', 8);
  } catch (e) {
    if (!quiet) SpreadsheetApp.getUi().alert('Could not refresh projects from Masters', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/** Menu: ❓ Start here. Creates the tab the first time, then just opens it. */
function IM_showStartHere() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('Start here');
  if (!sh) {
    sh = ss.insertSheet('Start here', 0);
    var rows = [
      ['Start here: IMS (Inventory)', 'title'],
      ['The normal flow, step by step', 'h'],
      ['1.  Goods arrive: Add Inward (GRN). Stock goes up and the item shows the new quantity.', ''],
      ['2.  Reservations normally come from PM: when someone ticks Reserve? on the PM Hardware List, a row appears in the Reservation Register as Requested and becomes Active within about a minute. You do not have to do anything. You can also reserve by hand with the Reservation form.', ''],
      ["3.  Issuing material to a project: Add Material Requisition. Pick the project from the dropdown (it starts with PRJ-…). Then the project's reservation reduces by itself, and Finance shows the cost within about 10 minutes. The form remembers the last project you used.", ''],
      ['4.  If a requisition is more than the free stock, the form warns you and asks you to confirm before saving.', ''],
      ['5.  Reservation form, Release or Consume: Released frees the stock again. Consumed records the issue by posting a requisition for you (a partial quantity is allowed; the rest stays reserved).', ''],
      ['6.  A new item needed by a project can be added from PM (🔗 Interlink → ➕ Add item to IMS). Enter its Avg Rate in Item Master, so Finance can cost it when it is issued.', ''],
      ['What the words mean', 'h'],
      ['Free stock  what is in stock minus what projects have reserved.', ''],
      ['Requisition  issuing material from the store to a project. Reservation  holding stock for a project before it is issued.', ''],
      ['Reservation status  Requested (PM asked, waiting for IMS), Active (held), Partially Consumed (some issued, the rest still held), Consumed (all issued), Released (given back), Error (see the Notes column for why).', ''],
      ['If something looks wrong', 'h'],
      ['•  🔗 Interlink → Process PM reservation requests now: answers PM requests at once if they are waiting.', ''],
      ['•  🔗 Interlink → Refresh dropdowns from Masters: new PM projects appear in the project dropdown.', ''],
      ['•  In PM, 🔗 Interlink → 📋 Needs attention and 🩺 Health check show what is stuck, in plain words.', '']
    ];
    sh.getRange(1, 1, rows.length, 1).setValues(rows.map(function (r) { return [r[0]]; })).setWrap(true).setVerticalAlignment('top');
    sh.setColumnWidth(1, 900);
    rows.forEach(function (r, i) {
      if (r[1] === 'title') sh.getRange(i + 1, 1).setFontSize(16).setFontWeight('bold');
      if (r[1] === 'h') sh.getRange(i + 1, 1).setFontWeight('bold').setBackground('#DDEBF7');
    });
  }
  ss.setActiveSheet(sh);
}

/**
 * Employees (Dropdowns col I, used for Taken By / Requested By): GENERATED from Masters' Employees. The column is rewritten
 * completely (active people, A to Z); it exists only because the form dropdowns read this sheet, and it is not a list to edit.
 * Names not active in Masters disappear from it (requisitions already saved keep their text). If Masters cannot be read, or returns
 * nobody, the column is left exactly as it is. Returns how many names are in the list now; 0 when nothing was done.
 */
var IM_EMP_NEW_ = [];   // names that were not in the column before the last refresh (for the toast)
function IM_refreshEmployeeDropdown_(dd) {
  IM_EMP_NEW_ = [];
  try {
    var norm = function (s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); };
    var names = [], seen = {};
    (IM_getEmployees() || []).forEach(function (r) {
      var n = String(r.Name || '').replace(/\s+/g, ' ').trim(), k = norm(n);
      if (!k || String(r.Status) === 'Left' || seen[k]) return;
      seen[k] = true; names.push(n);
    });
    if (!names.length) return 0;
    names.sort(function (a, b) { return a.toLowerCase() < b.toLowerCase() ? -1 : a.toLowerCase() > b.toLowerCase() ? 1 : 0; });
    var last = dd.getLastRow();
    var cur = last >= 2 ? dd.getRange(2, 9, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
    var had = {}; cur.forEach(function (v) { if (v) had[norm(v)] = true; });
    IM_EMP_NEW_ = names.filter(function (n) { return !had[norm(n)]; });
    var same = cur.length >= names.length && names.every(function (n, i) { return cur[i] === n; }) && cur.slice(names.length).every(function (v) { return v === ''; });
    if (!same) {
      if (last >= 2) dd.getRange(2, 9, last - 1, 1).clearContent();
      dd.getRange(2, 9, names.length, 1).setValues(names.map(function (n) { return [n]; }));
    }
    return names.length;
  } catch (e) { return 0; }
}

// -----------------------------------------------------------------------------
// ➕ Add employee. The person is created in Masters (the one official list) and this system's own list is updated at
// once. Masters editors only (every Masters write is). Masters refuses an exact duplicate and warns on a similar name.
// -----------------------------------------------------------------------------
function IM_addEmployeeDialog() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(IM_addEmployeeHtml_()).setWidth(430).setHeight(500), '➕ Add employee');
}
function IM_addEmployeeHtml_() { return "<div style=\"font:14px Arial,sans-serif;padding:14px;color:#222\">\n<p style=\"margin:0 0 12px;color:#555\">The person is saved in <b>Masters</b>, the official employee list. The other systems pick the name up from there.</p>\n<label>Full name *</label><input id=\"n\" maxlength=\"80\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Role (optional)</label><input id=\"r\" maxlength=\"60\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Department (optional)</label><input id=\"d\" maxlength=\"60\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Phone (optional)</label><input id=\"p\" maxlength=\"30\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<div id=\"m\" style=\"min-height:42px;margin:6px 0;font-size:13px\"></div>\n<button id=\"go\" onclick=\"send(false)\" style=\"padding:8px 16px;background:#1F4E79;color:#fff;border:0;border-radius:4px;cursor:pointer\">Add employee</button>\n<button id=\"any\" onclick=\"send(true)\" style=\"display:none;padding:8px 16px;background:#b45309;color:#fff;border:0;border-radius:4px;cursor:pointer\">Add anyway (different person)</button>\n<button onclick=\"google.script.host.close()\" style=\"padding:8px 16px;margin-left:6px;border:1px solid #bbb;background:#fff;border-radius:4px;cursor:pointer\">Close</button>\n</div>\n<script>\nfunction $(i){return document.getElementById(i);}\nfunction say(t,c){var m=$('m');m.textContent=t;m.style.color=c||'#222';}\nfunction busy(b){$('go').disabled=b;$('any').disabled=b;}\nfunction send(confirm){\n  var name=$('n').value.replace(/\\s+/g,' ').trim();\n  if(!name){say('Please type the name.','#b91c1c');return;}\n  busy(true);say('Saving\u2026');\n  google.script.run.withSuccessHandler(done).withFailureHandler(function(e){busy(false);say('Could not save: '+(e&&e.message||e),'#b91c1c');})\n    .IM_addEmployee({name:name,role:$('r').value,department:$('d').value,phone:$('p').value,confirmSimilar:confirm===true});\n}\nfunction done(r){\n  busy(false);$('any').style.display='none';\n  if(r&&r.ok){say('Added '+r.name+' ('+r.id+'). '+(r.note||''),'#166534');$('go').style.display='none';$('n').disabled=$('r').disabled=$('d').disabled=$('p').disabled=true;}\n  else if(r&&r.duplicate){say(r.duplicate.name+' is already in Masters ('+r.duplicate.id+', '+r.duplicate.status+'). Nothing was added.','#b45309');}\n  else if(r&&r.similar){say('Similar name(s) already in Masters: '+r.similar.map(function(s){return s.name+' ('+s.id+', '+s.status+')';}).join('; ')+'. If this is the same person, close this window. If it is a different person, click Add anyway.','#b45309');$('any').style.display='inline-block';}\n  else{say((r&&r.error)||'Something went wrong.','#b91c1c');}\n}\n</script>"; }
/** Called by the dialog. Returns { ok, id, name, note } | { ok:false, duplicate | similar | error }. Never throws. */
function IM_addEmployee(d) {
  try {
    var r = Masters.addEmployeeChecked({ Name: d.name, Role: d.role, Department: d.department, Phone: d.phone, confirmSimilar: d.confirmSimilar === true });
    if (r.ok) { try { IM_clearCache(); } catch (e) {} r.note = IM_addEmployeePost_(r, d); }
    return r;
  } catch (e) {
    var m = String((e && e.message) || e);
    return { ok: false, error: /Editors only|Admin only|permission|access/i.test(m)
      ? 'Only Masters editors can add employees. Ask the administrator to add you as an editor, or ask them to add this person.' : m };
  }
}
/** IMS: the name goes into the employee list (Taken By / Requested By) straight away. */
function IM_addEmployeePost_(r, d) {
  try {
    var dd = SpreadsheetApp.getActive().getSheetByName('Dropdowns');
    if (dd) IM_refreshEmployeeDropdown_(dd);
    return r.name + ' now appears in the employee list (Taken By / Requested By). Payroll: the accountant adds the salary in Finance.';
  } catch (e) { return 'Saved in Masters. The employee list updates the next time IMS is opened.'; }
}

// -----------------------------------------------------------------------------
// SUPPLIERS and JOB WORKERS: Masters (Parties) is the master. The IMS sheets 'Supplier Master' and 'Job Worker Master' are kept as a
// GENERATED COPY (columns A to I), so every form, PO, GRN and job work challan that reads them keeps working unchanged.
//   * A party's IMS code (SUPP-xxxx / JW-xxxx) is stored in Masters, so old documents stay valid. A party that has no IMS code yet
//     (added in Masters, or copied from Finance) shows its Masters code (VEND-xxxx) here.
//   * The copy never deletes a row and never blanks a cell: a row that Masters does not know yet (typed here by hand, or not linked yet) is
//     left alone, and a Masters party whose NAME equals such a row is not added a second time (it waits to be linked).
//   * Linking: IMS rows Masters does not know yet are sent to Masters (needs a Masters editor; runs by itself when an editor opens IMS).
// -----------------------------------------------------------------------------
var IM_PARTY_KINDS_ = [
  { kind: 'supplier',  sheet: 'Supplier Master',   flag: { isSupplier: true },  code: 'IMS_Supplier_Code', label: 'supplier' },
  { kind: 'jobworker', sheet: 'Job Worker Master', flag: { isJobWorker: true }, code: 'IMS_JW_Code',       label: 'job worker' }
];
function IM_partyCacheKeys_() { return ['parties_{}', 'parties_{"isSupplier":true}', 'parties_{"isJobWorker":true}'].map(function (k) { return IM_CACHE_KEY + k; }); }
function IM_partyClearCache_() { IM_partyCacheKeys_().forEach(function (key) { try { CacheService.getScriptCache().remove(key); } catch (e) {} }); }
function IM_partyTxt_(v) { return String(v == null ? '' : v).trim(); }
function IM_partyActive_(v) { return v === true || String(v).toLowerCase() === 'true' || String(v).toLowerCase() === 'yes'; }
function IM_partyNorm_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }

/** Rewrites the generated copy of one sheet from Masters. Returns { rows, added, changed, waiting:[names], onlyHere:[names] } or null (sheet missing / Masters returned nobody). */
function IM_mirrorParties_(k) {
  var t = IM_partyTxt_, sh = SpreadsheetApp.getActive().getSheetByName(k.sheet);
  if (!sh) return null;
  var parties = (IM_getParties(k.flag) || []).filter(function (p) { return t(p.Name); });
  if (!parties.length) return null;                      // Masters returned nothing: leave the sheet exactly as it is
  var next = nextRowByCol(sh, 1), cur = next > 2 ? sh.getRange(2, 1, next - 2, 9).getValues() : [];
  var rowOf = {}; cur.forEach(function (r, i) { if (t(r[0])) rowOf[t(r[0])] = i; });
  var inMasters = {}, want = [];
  parties.forEach(function (p) {
    var code = t(p[k.code]) || t(p.Party_ID);
    inMasters[code] = true;
    want.push([code, t(p.Name).replace(/\s+/g, ' '), t(p.Category), t(p.Contact_Person), t(p.Phone), t(p.Email), t(p.GSTIN), t(p.Address), String(p['Active?'] || 'Yes').toLowerCase() !== 'no']);
  });
  var hereNames = {}, onlyHere = [];
  cur.forEach(function (r) { if (t(r[0]) && !inMasters[t(r[0])]) { hereNames[IM_partyNorm_(r[1])] = true; onlyHere.push(t(r[1])); } });
  var res = { rows: parties.length, added: 0, changed: 0, waiting: [], onlyHere: onlyHere }, add = [];
  want.forEach(function (w) {
    var i = rowOf[w[0]];
    if (i === undefined) {
      if (hereNames[IM_partyNorm_(w[1])]) { res.waiting.push(w[1]); return; }   // same firm as a row here that is not linked yet: do not add it twice
      add.push(w); return;
    }
    var r = cur[i], diff = false;
    for (var j = 1; j < 8; j++) { if (w[j] === '') continue; if (t(r[j]) !== w[j]) { r[j] = w[j]; diff = true; } }   // a blank in Masters never blanks a cell here
    if (IM_partyActive_(r[8]) !== w[8]) { r[8] = w[8]; diff = true; }
    if (diff) res.changed++;
  });
  if (res.changed) sh.getRange(2, 1, cur.length, 9).setValues(cur);
  if (add.length) { sh.getRange(next, 1, add.length, 9).setValues(add); res.added = add.length; }
  try { if (!sh.getRange('A1').getNote()) sh.getRange('A1').setNote('Kept in step with Masters (Parties) automatically. Add or change ' + k.label + 's with the IMS forms or in Masters: a change typed straight into this sheet is overwritten the next time Masters is read.'); } catch (e) {}
  return res;
}

/** Sends the rows of one sheet that Masters does not know yet (by IMS code or Masters code) to Masters. Returns { sent, blocked, error }. */
function IM_linkPartyRows_(k) {
  var t = IM_partyTxt_, sh = SpreadsheetApp.getActive().getSheetByName(k.sheet), out = { sent: 0 };
  if (!sh) return out;
  var known = {}; (IM_getParties(k.flag) || []).forEach(function (p) { if (t(p[k.code])) known[t(p[k.code])] = true; known[t(p.Party_ID)] = true; });
  var next = nextRowByCol(sh, 1); if (next <= 2) return out;
  var list = sh.getRange(2, 1, next - 2, 9).getValues().filter(function (r) { return t(r[0]) && t(r[1]) && !known[t(r[0])]; })
    .map(function (r) { return { code: t(r[0]), name: t(r[1]), category: t(r[2]), contact: t(r[3]), phone: t(r[4]), email: t(r[5]), gstin: t(r[6]), address: t(r[7]), active: IM_partyActive_(r[8]) }; });
  if (!list.length) return out;
  try {
    var r2, n = 0;
    do { r2 = Masters.syncPartiesFromIms(k.kind, list, 20); n++; } while (r2.more && n < 6);
    out.sent = list.length; out.more = r2.more;
    IM_partyClearCache_();
  } catch (e) { out.blocked = true; out.error = /Editors only|Admin only|permission|access/i.test(String(e && e.message)) ? 'only Masters editors can send suppliers and job workers' : String(e && e.message); }
  return out;
}

/**
 * Link (editors) then copy, for both sheets, under the document lock (a second person opening IMS at the same time just skips it).
 * Returns { supplier, jobworker, sent, linkBlocked, error }, each of supplier / jobworker being the copy result; never throws.
 */
function IM_syncPartiesToIms_(fresh) {
  var out = { linkBlocked: '', sent: 0 }, lock = LockService.getDocumentLock();
  try { if (!lock.tryLock(8000)) return out; } catch (e) { return out; }
  try {
    if (fresh) IM_partyClearCache_();
    IM_PARTY_KINDS_.forEach(function (k) {
      try {
        var l = IM_linkPartyRows_(k);
        if (l.blocked) out.linkBlocked = l.error; else out.sent += l.sent;
        out[k.kind] = IM_mirrorParties_(k);
      } catch (e) { out[k.kind] = null; out.error = String((e && e.message) || e); }
    });
  } finally { try { lock.releaseLock(); } catch (e) {} }
  return out;
}

/** Plain-words summary of IM_syncPartiesToIms_ ('' when there is nothing to say). */
function IM_partySummary_(r) {
  var one = function (x, label) { return x ? x.rows + ' ' + label + (x.added || x.changed ? ' (' + (x.added ? x.added + ' added' : '') + (x.added && x.changed ? ', ' : '') + (x.changed ? x.changed + ' updated' : '') + ')' : '') : ''; };
  var parts = [one(r.supplier, 'suppliers'), one(r.jobworker, 'job workers')].filter(String);
  var s = r.supplier, j = r.jobworker;
  var only = (s ? s.onlyHere.length : 0) + (j ? j.onlyHere.length : 0), wait = (s ? s.waiting.length : 0) + (j ? j.waiting.length : 0);
  var msg = parts.length ? 'Suppliers and job workers, from Masters: ' + parts.join(', ') + '.' : '';
  if (only) msg += ' ' + only + ' row(s) exist only in IMS' + (r.linkBlocked ? ': a Masters editor must open IMS (or use 🔗 Interlink → Send suppliers and job workers to Masters) to add them to Masters.' : ' (sent to Masters just now).');
  else if (wait) msg += ' ' + wait + ' Masters party(ies) wait to be linked to an IMS row of the same name.';
  if (r.error) msg += ' Problem: ' + r.error;
  return msg.trim();
}

/** Menu: 🔗 Send suppliers and job workers to Masters (and copy Masters back). */
function IM_syncPartiesNow() {
  var r = IM_syncPartiesToIms_(true);
  SpreadsheetApp.getActive().toast(IM_partySummary_(r) || (r.linkBlocked ? 'Not sent: ' + r.linkBlocked : 'Suppliers and job workers are already in step with Masters.'), '🔗 Interlink', 12);
}

/**
 * Called by the Add / Edit supplier and job worker forms. Writes to Masters (the master) and refreshes the sheet copy.
 * kind: 'supplier' | 'jobworker'. Returns { ok, code, message } or { ok:false, notInMasters:true } (edit of a row Masters does not know yet). Never throws.
 */
function IM_partyFromForm_(kind, code, d, isEdit) {
  var k = IM_PARTY_KINDS_.filter(function (x) { return x.kind === kind; })[0], t = IM_partyTxt_;
  try {
    var cat = kind === 'jobworker' ? d.processType : d.category;
    var res;
    if (isEdit) {
      IM_partyClearCache_();
      var p = (IM_getParties(k.flag) || []).filter(function (x) { return t(x[k.code]) === code || t(x.Party_ID) === code; })[0];
      if (!p) return { ok: false, notInMasters: true };
      var upd = {};
      if (d.firmName !== undefined) upd.Name = t(d.firmName);
      if (cat !== undefined) upd.Category = t(cat);
      if (d.contact !== undefined) upd.Contact_Person = t(d.contact);
      if (d.mobile !== undefined) upd.Phone = t(d.mobile);
      if (d.email !== undefined) upd.Email = t(d.email);
      if (d.gst !== undefined) upd.GSTIN = t(d.gst);
      if (d.address !== undefined) upd.Address = t(d.address);
      if (d.active !== undefined) upd['Active?'] = d.active ? 'Yes' : 'No';
      Masters.updateParty(p.Party_ID, upd);
      res = { ok: true, id: p.Party_ID };
    } else {
      res = Masters.addPartyChecked({ Name: t(d.firmName), role: kind, code: code, Category: t(cat), Contact_Person: t(d.contact), Phone: t(d.mobile), Email: t(d.email), GSTIN: t(d.gst), Address: t(d.address) });
      if (!res.ok) return { ok: false, message: res.duplicate ? res.duplicate.name + ' is already in Masters as a ' + k.label + ' (' + res.duplicate.id + '). Nothing was added; it is in the list already.' : 'Could not add.' };
    }
    var s = IM_syncPartiesToIms_(true);
    var now = (IM_getParties(k.flag) || []).filter(function (x) { return x.Party_ID === res.id; })[0];
    return { ok: true, code: now ? (t(now[k.code]) || t(now.Party_ID)) : (code || res.id), message: s.error || '' };
  } catch (e) {
    var m = String((e && e.message) || e);
    return { ok: false, message: /Editors only|Admin only|permission|access/i.test(m) ? 'Only Masters editors can add or change suppliers and job workers. Ask the administrator to add you as an editor, or ask them to do it.' : m };
  }
}
