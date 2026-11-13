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
    .addSeparator()
    .addItem('Process PM reservation requests now', 'IM_processReservationRequestsNow')
    .addSeparator()
    .addItem('❓ Start here',                  'IM_showStartHere')
    .addItem('About',                          'IM_about_')
    .addToUi();
  // Self-heal the PM reservation queue: recreate its trigger if missing, and drain anything waiting.
  try { IM_ensureQueueTrigger_(); IM_processReservationRequests(); } catch (e) {}
  IM_refreshDropdowns(true);  // keep PRJ-### project suggestions current (never throws when quiet)
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
    if (!quiet) ss.toast(changed ? 'Project list updated from Masters (' + labels.length + ' project(s)).' : 'Projects already up to date.', '🔗 Interlink', 5);
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
