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
    .addItem('Clear Masters cache',            'IM_clearCache')
    .addSeparator()
    .addItem('About',                          'IM_about_')
    .addToUi();
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
// Dropdown refresher — stub for v0.4 Phase 1. Each system will fill in its
// own IM_refreshDropdowns() in a sibling _Interlink_Dropdowns.gs file that
// knows which sheets+columns to re-validate.
// -----------------------------------------------------------------------------
function IM_refreshDropdowns() {
  SpreadsheetApp.getUi().alert(
    'Dropdowns not wired yet',
    'This system doesn\'t have per-sheet dropdown rules configured yet. ' +
    'That lives in _Interlink_Dropdowns.gs (next step).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
