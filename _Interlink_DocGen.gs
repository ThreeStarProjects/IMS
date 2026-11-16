/**
 * _Interlink_DocGen.gs
 * -----------------------------------------------------------------------------
 * The doorway between IMS and the DocGen app (3SP_DocGen on Vercel). Only for Purchase Orders and Delivery Challans.
 *
 *   feeds     DocGen asks for the lists it offers in its PO / Delivery Challan forms: suppliers, items, customers, projects.
 *   recordPo  a PO saved in DocGen is posted here with the REAL IMS function (addPoMultiLine): IMS numbers it, writes the
 *             Purchase Order Register + PO Lines, and no stock changes until a GRN, exactly like a PO made in IMS.
 *   recordDc  a Delivery Challan saved in DocGen is posted with the REAL addFgDelivery: IMS numbers it, writes the FG Delivery
 *             Register + Lines and reduces stock, exactly like a challan made in IMS. If lines exceed available stock IMS answers
 *             needsConfirm (with the warnings) and nothing is written until DocGen sends it again with bypassStockCheck.
 *
 * IMS makes NO PDF for these (DocGen draws its own), and the document number DocGen shows is the IMS number.
 * Safety: every request carries a secret key (script property DOCGEN_SECRET); only the three actions above exist; only whitelisted
 * fields are passed on; each DocGen document has an id (docRef) and a repeated request returns the number already issued
 * (hidden sheet _DocGen_Log), so a retry can never post twice.
 * Setup (once): 🔗 Interlink → DocGen link: show key. Then deploy IMS as a Web App (Execute as: Me, Access: Anyone).
 */

var IM_DG_MAX_LINES = 300;

function doGet() { return IM_dgJson_({ ok: true, service: 'IMS DocGen link' }); }

function doPost(e) {
  var out;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = IM_dgHandle_(body);
  } catch (err) { out = { ok: false, error: String((err && err.message) || err) }; }
  return IM_dgJson_(out);
}

function IM_dgJson_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Compares two strings without stopping at the first difference. */
function IM_dgSame_(a, b) {
  a = String(a || ''); b = String(b || '');
  var d = a.length ^ b.length;
  for (var i = 0; i < Math.max(a.length, b.length); i++) d |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return d === 0;
}

function IM_dgHandle_(body) {
  var secret = PropertiesService.getScriptProperties().getProperty('DOCGEN_SECRET');
  if (!secret) return { ok: false, error: 'The DocGen link is not set up in IMS yet (🔗 Interlink → DocGen link).' };
  if (!IM_dgSame_(body && body.secret, secret)) return { ok: false, error: 'Wrong key.' };
  var d = (body && body.data) || {};
  switch (body.action) {
    case 'feeds':    return { ok: true, feeds: IM_dgFeeds_() };
    case 'recordPo': return IM_dgRecord_('PO', d, IM_dgPo_);
    case 'recordDc': return IM_dgRecord_('DC', d, IM_dgDc_);
    default:         return { ok: false, error: 'Unknown action.' };
  }
}

// ---------------------------------------------------------------- feeds
function IM_dgCol_(hdr, names) {
  for (var i = 0; i < names.length; i++) { var k = hdr.indexOf(names[i]); if (k >= 0) return k; }
  return -1;
}
function IM_dgRows_(sheetName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return { hdr: [], rows: [] };
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  return { hdr: v[0].map(function (h) { return String(h).trim(); }), rows: v.slice(1) };
}
function IM_dgFeeds_() {
  var t = function (v) { return String(v == null ? '' : v).trim(); };
  var sup = IM_dgRows_('Supplier Master'), s = function (names) { return IM_dgCol_(sup.hdr, names); };
  var cC = s(['Supplier Code']), cN = s(['Firm Name']), cCat = s(['Category']), cCp = s(['Contact Person']), cM = s(['Mobile No']), cE = s(['Email']), cG = s(['GST No']), cA = s(['Address']), cAct = s(['Active?']);
  var suppliers = sup.rows.filter(function (r) { return t(r[cC]) && t(r[cN]) && r[cAct] !== false && String(r[cAct]).toLowerCase() !== 'false'; }).map(function (r) {
    var g = t(r[cG]).toUpperCase();
    return { code: t(r[cC]), name: t(r[cN]), category: t(r[cCat]), contact: t(r[cCp]), mobile: t(r[cM]), email: t(r[cE]), gst: g, stateCode: /^\d{2}/.test(g) ? g.slice(0, 2) : '', address: t(r[cA]) };
  });
  var it = IM_dgRows_('Item Master'), i = function (names) { return IM_dgCol_(it.hdr, names); };
  var kC = i(['Item Code']), kD = i(['Description']), kN = i(['Item Name']), kS = i(['Size']), kG = i(['Grade/Color']), kU = i(['UOM']), kH = i(['HSN/SAC Code']), kA = i(['Active?']), kR = i(['Avg Rate (Rs.)']);
  var items = it.rows.filter(function (r) { return t(r[kC]) && r[kA] !== false && String(r[kA]).toLowerCase() !== 'false'; }).map(function (r) {
    return { code: t(r[kC]), description: t(r[kD]) || t(r[kN]), name: t(r[kN]), size: t(r[kS]), grade: t(r[kG]), uom: t(r[kU]), hsn: t(r[kH]), rate: Number(r[kR]) || 0 };
  });
  // Finished goods + semi-finished goods: the codes a Delivery Challan's MAIN lines use (stock goes down for these)
  var goods = [['FG Master', 'FG Code', 'FG'], ['SFG Master', 'SFG Code', 'SFG']].reduce(function (acc, m) {
    var g = IM_dgRows_(m[0]), q = function (names) { return IM_dgCol_(g.hdr, names); }, gc = q([m[1]]), gd = q(['Description']), gn = q(['Item Name']), gu = q(['UOM']), ga = q(['Active?']);
    g.rows.forEach(function (r) {
      if (t(r[gc]) && r[ga] !== false && String(r[ga]).toLowerCase() !== 'false') acc.push({ code: t(r[gc]), description: t(r[gd]) || t(r[gn]), uom: t(r[gu]), kind: m[2] });
    });
    return acc;
  }, []);
  var customers = IM_customerPicks_().map(function (c) { return { name: c.name, legalName: c.legal, gst: c.gstin, billAddress: c.bill, shipAddress: c.ship, paymentTerms: c.terms }; });
  var projects = (IM_getProjects() || []).filter(function (p) { return t(p.Name) && String(p['Active?'] || 'Yes').toLowerCase() !== 'no'; })
    .map(function (p) { return { code: t(p.Project_ID), name: t(p.Name), label: t(p.Project_ID) + ' — ' + t(p.Name) }; });
  return { suppliers: suppliers, items: items, goods: goods, customers: customers, projects: projects, at: new Date().toISOString() };
}

// ---------------------------------------------------------------- recording (idempotent)
function IM_dgLog_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('_DocGen_Log');
  if (!sh) { sh = ss.insertSheet('_DocGen_Log'); sh.appendRow(['DocGen id', 'Type', 'IMS number', 'At']); sh.hideSheet(); }
  return sh;
}

/** kind 'PO' | 'DC'; build(d) -> the object for the real IMS function's result: { success, code, message, needsConfirm?, warnings? } */
function IM_dgRecord_(kind, d, build) {
  var ref = String((d && d.docRef) || '').trim();
  if (!ref || ref.length > 80) return { ok: false, error: 'docRef (the DocGen document id) is required.' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return { ok: false, error: 'IMS is busy, try again in a minute.' }; }
  try {
    var log = IM_dgLog_(), last = log.getLastRow();
    if (last >= 2) {
      var hit = log.getRange(2, 1, last - 1, 3).getValues().filter(function (r) { return r[0] === ref && r[1] === kind; })[0];
      if (hit) return { ok: true, number: hit[2], already: true };
    }
    var r = build(d);
    if (r && r.needsConfirm) return { ok: false, needsConfirm: true, warnings: r.warnings || [], error: r.message };
    if (!r || !r.success || !r.code) return { ok: false, error: (r && r.message) || 'IMS did not accept the document.' };
    log.appendRow([ref, kind, r.code, new Date()]);
    return { ok: true, number: r.code, message: r.message };
  } catch (e) {
    return { ok: false, error: String((e && e.message) || e) };
  } finally { lock.releaseLock(); }
}

function IM_dgText_(v) { return String(v == null ? '' : v).trim(); }
function IM_dgTag_(d) { return d.by ? '[DocGen: ' + IM_dgText_(d.by) + '] ' : '[DocGen] '; }

function IM_dgPo_(d) {
  var lines = d.lineItems;
  if (!Array.isArray(lines) || !lines.length) throw new Error('At least one line is required.');
  if (lines.length > IM_DG_MAX_LINES) throw new Error('Too many lines.');
  var t = IM_dgText_;
  return addPoMultiLine({
    date: d.date, supplierCode: t(d.supplierCode), deliveryDate: d.deliveryDate || '',
    remarks: IM_dgTag_(d) + t(d.remarks),
    lineItems: lines.map(function (l) { return { itemCode: t(l.itemCode), qty: Number(l.qty), rate: Number(l.rate), gst: Number(l.gst) || 0, discount: Number(l.discount) || 0, discountMode: l.discountMode === 'amt' ? 'amt' : 'pct' }; }),
    billToName: t(d.billToName), billToAddress: t(d.billToAddress), billToGstin: t(d.billToGstin), billToContact: t(d.billToContact),
    shipToName: t(d.shipToName), shipToAddress: t(d.shipToAddress), shipToGstin: t(d.shipToGstin), shipToContact: t(d.shipToContact),
    skipPdf: true
  });
}

function IM_dgDc_(d) {
  var t = IM_dgText_, main = d.mainItems, hw = d.hardwareItems || [];
  if (!Array.isArray(main) || !main.length) throw new Error('At least one main item is required.');
  if (main.length + hw.length > IM_DG_MAX_LINES) throw new Error('Too many lines.');
  var line = function (l) { return { itemCode: t(l.itemCode), description: t(l.description), qty: Number(l.qty), uom: t(l.uom) || 'Nos', weight: Number(l.weight) || 0, remark: t(l.remark) }; };
  return addFgDelivery({
    date: d.date, deliveryType: d.deliveryType === 'Internal' ? 'Internal' : 'Customer',
    customerName: t(d.customerName), project: t(d.project), paymentTerms: t(d.paymentTerms), otherRef: IM_dgTag_(d) + t(d.otherRef),
    billTo: t(d.billTo), billToGstin: t(d.billToGstin), shipTo: t(d.shipTo), shipToGstin: t(d.shipToGstin),
    vehicleNo: t(d.vehicleNo), driver: t(d.driver), lrNo: t(d.lrNo), eWayBillNo: t(d.eWayBillNo),
    mainItems: main.map(line), hardwareItems: hw.map(line),
    bypassStockCheck: d.bypassStockCheck === true, skipPdf: true
  });
}

// ---------------------------------------------------------------- setup (menu)
/** Menu: 🔗 Interlink → DocGen link: show key. Creates the key the first time; shows it again later. */
function IM_docgenSetup() {
  var ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties(), key = props.getProperty('DOCGEN_SECRET'), made = false;
  if (!key) { key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); props.setProperty('DOCGEN_SECRET', key); made = true; }
  ui.alert('DocGen link' + (made ? ' — key created' : ''),
    'DocGen needs two settings (Vercel → Project → Settings → Environment Variables):\n\n' +
    'IMS_API_SECRET\n' + key + '\n\n' +
    'IMS_API_URL\nthe Web App URL of IMS (Extensions ▸ Apps Script ▸ Deploy ▸ New deployment ▸ Web app ▸ Execute as: Me ▸ Who has access: Anyone). ' +
    'Anyone can reach the address, but nothing works without the key above.\n\n' +
    'Keep the key private. If it ever leaks, delete the script property DOCGEN_SECRET and open this menu again for a new one.', ui.ButtonSet.OK);
}
