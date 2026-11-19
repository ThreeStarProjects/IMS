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
    case 'feeds':    if (d && d.refresh === true) { try { IM_clearCache(); } catch (e) {} }   // 'refresh': forget IMS's 5-minute copy of the Masters lists first
                     return { ok: true, feeds: IM_dgFeeds_() };
    case 'peek':     return IM_dgPeek_(d);
    case 'recordPo': return IM_dgRecord_('PO', d, IM_dgPo_, IM_dgRevisePo_);
    case 'recordDc': return IM_dgRecord_('DC', d, IM_dgDc_, IM_dgReviseDc_);
    case 'recordRfq': return IM_dgRecord_('RFQ', d, IM_dgRfq_, IM_dgReviseRfq_);
    case 'recordJwc': return IM_dgRecord_('JWC', d, IM_dgJwc_, IM_dgReviseJwc_);
    case 'photos':   return IM_dgPhotos_(d);
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
  var jwm = IM_dgRows_('Job Worker Master'), w = function (names) { return IM_dgCol_(jwm.hdr, names); };
  var wC = w(['Job Worker Code']), wN = w(['Firm Name']), wP = w(['Process Type']), wCp = w(['Contact Person']), wM = w(['Mobile No']), wE = w(['Email']), wG = w(['GST No']), wA = w(['Address']), wAct = w(['Active?']);
  var jobWorkers = jwm.rows.filter(function (r) { return t(r[wC]) && t(r[wN]) && r[wAct] !== false && String(r[wAct]).toLowerCase() !== 'false'; }).map(function (r) {
    return { code: t(r[wC]), name: t(r[wN]), processType: t(r[wP]), contact: t(r[wCp]), mobile: t(r[wM]), email: t(r[wE]), gst: t(r[wG]).toUpperCase(), address: t(r[wA]) };
  });
  return { suppliers: suppliers, items: items, goods: goods, customers: customers, projects: projects, jobWorkers: jobWorkers, at: new Date().toISOString() };
}

// ---------------------------------------------------------------- recording (idempotent)
function IM_dgLog_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('_DocGen_Log');
  if (!sh) { sh = ss.insertSheet('_DocGen_Log'); sh.appendRow(['DocGen id', 'Type', 'IMS number', 'At', 'Content']); sh.hideSheet(); }
  if (sh.getMaxColumns() < 5) { sh.insertColumnsAfter(sh.getMaxColumns(), 5 - sh.getMaxColumns()); sh.getRange(1, 5).setValue('Content'); }
  return sh;
}

/** A fingerprint of what the document says (not who saved it or when), so saving the same content again is recognised and does nothing. */
function IM_dgDigest_(d) {
  var c = {}; Object.keys(d).sort().forEach(function (k) { if (['docRef', 'by', 'imsNumber', 'bypassStockCheck'].indexOf(k) < 0) c[k] = d[k]; });
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(c)));
}

/** TSP/PO/26-27/043, 043A, 043B all belong to one document: the "family" is the number without its version letters. */
function IM_dgFamily_(no) { var m = /^(.*\d)[A-Z]*$/.exec(String(no || '')); return m ? m[1] : String(no || ''); }
/** The next version of a number: 043 -> 043A -> 043B ... 043Z -> 043AA. */
function IM_dgNextVersion_(no) {
  var m = /^(.*\d)([A-Z]*)$/.exec(String(no || '')), base = m ? m[1] : String(no), s = m ? m[2] : '';
  if (!s) return base + 'A';
  var out = '', carry = true;
  for (var i = s.length - 1; i >= 0; i--) { var c = s.charCodeAt(i) - 65; if (carry) { c++; carry = c > 25; if (carry) c = 0; } out = String.fromCharCode(65 + c) + out; }
  return base + (carry ? 'A' : '') + out;
}

/** The number a document is stored under in column `col` of `sh`: exactly `no`, or (DocGen holding an older version) its current version. */
function IM_dgResolve_(sh, col, no) {
  if (!sh || sh.getLastRow() < 2) return no;
  var v = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues(), fam = IM_dgFamily_(no), alt = '';
  for (var i = 0; i < v.length; i++) { if (v[i][0] === no) return no; if (!alt && IM_dgFamily_(v[i][0]) === fam) alt = v[i][0]; }
  return alt || no;
}
/** Rename every row of `sh` whose column `col` is `from` (a new version letter). Refuses to take a number that exists. */
function IM_dgRename_(sh, col, from, to) {
  if (!sh || sh.getLastRow() < 2) return;
  var v = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  if (v.some(function (r) { return r[0] === to; })) throw new Error(to + ' already exists in IMS.');
  v.forEach(function (r, i) { if (r[0] === from) sh.getRange(i + 2, col).setValue(to); });
}

/**
 * kind 'PO' | 'DC'. create(d) -> { success, code, message, needsConfirm?, warnings? } is a NEW document (idempotent by docRef).
 * When d.imsNumber is given it is a NEW VERSION of a document DocGen already recorded: revise(d, number) replaces that document's lines
 * in place (never adds a second copy, never deducts twice). The same content saved again is a no-op.
 */
function IM_dgRecord_(kind, d, create, revise) {
  var ref = String((d && d.docRef) || '').trim();
  if (!ref || ref.length > 80) return { ok: false, error: 'docRef (the DocGen document id) is required.' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return { ok: false, error: 'IMS is busy, try again in a minute.' }; }
  try {
    var log = IM_dgLog_(), last = log.getLastRow(), rows = last >= 2 ? log.getRange(2, 1, last - 1, 5).getValues() : [];
    var digest = IM_dgDigest_(d), number = extractCodePrefix(IM_dgText_(d.imsNumber)), r;
    if (!number) {
      // The same DocGen id again (a retry after a later step failed). Same content = nothing to do. Changed content (the user fixed the
      // form and saved again) = a new version of that document, never a silent "already done" that would leave IMS with the old lines.
      var hit = rows.filter(function (x) { return x[0] === ref && x[1] === kind; }).pop();   // the latest one (a PO may have become 043A, 043B ...)
      if (hit) { if (hit[4] === digest) return { ok: true, number: hit[2], already: true }; number = hit[2]; }
    }
    if (number) {
      var fam = IM_dgFamily_(number);
      var mine = rows.filter(function (x) { return x[1] === kind && IM_dgFamily_(x[2]) === fam; });   // every version of this document, oldest first
      if (!mine.length) return { ok: false, error: number + ' was not recorded through DocGen, so DocGen cannot change it. Amend it in IMS.' };
      if (mine[mine.length - 1][4] === digest) return { ok: true, number: mine[mine.length - 1][2], already: true };   // same content as the latest version: nothing to do (and DocGen learns the current number)
      r = revise(d, mine[mine.length - 1][2]);   // always the latest version's number, even if DocGen still holds an older one
    } else {
      r = create(d);
    }
    if (r && r.needsConfirm) return { ok: false, needsConfirm: true, warnings: r.warnings || [], error: r.message };
    if (!r || !r.success || !r.code) return { ok: false, error: (r && r.message) || 'IMS did not accept the document.' };
    log.appendRow([ref, kind, r.code, new Date(), digest]);
    return { ok: true, number: r.code, message: r.message, revised: !!number };
  } catch (e) {
    return { ok: false, error: String((e && e.message) || e) };
  } finally { lock.releaseLock(); }
}

/**
 * A number DocGen asks IMS to use, if any. Normally there is none: IMS issues the standard number itself (TSP/PO/26-27/043) when it creates the
 * document and returns it, so IMS-made and DocGen-made documents share ONE series and can never collide.
 */
function IM_dgNumber_(d) {
  var n = IM_dgText_(d.number);
  if (!n) return '';
  if (n.length > 40 || n.indexOf(' \u2014 ') >= 0) throw new Error('The document number can be at most 40 characters.');
  return n;
}

/** The number the next PO / Challan / RFQ / Job Work Challan will get, for the form to show before saving. It is only really taken when the document is saved. */
function IM_dgPeek_(d) {
  var kind = IM_dgText_(d.kind);
  if (!IM_DOC_TYPES_[kind]) return { ok: false, error: 'Unknown kind.' };
  var dt; try { dt = d.date ? parseDate(d.date) : new Date(); } catch (e) { dt = new Date(); }
  return { ok: true, number: IM_docNumber_(kind, dt, false) };
}

function IM_dgText_(v) { return String(v == null ? '' : v).trim(); }
function IM_dgTag_(d) { return d.by ? '[DocGen: ' + IM_dgText_(d.by) + '] ' : '[DocGen] '; }

function IM_dgPo_(d) {
  var lines = d.lineItems;
  if (!Array.isArray(lines) || !lines.length) throw new Error('At least one line is required.');
  if (lines.length > IM_DG_MAX_LINES) throw new Error('Too many lines.');
  var t = IM_dgText_;
  return addPoMultiLine({
    manualPoNo: IM_dgNumber_(d),
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
    manualDcNo: IM_dgNumber_(d),
    date: d.date, deliveryType: d.deliveryType === 'Internal' ? 'Internal' : 'Customer',
    customerName: t(d.customerName), project: t(d.project), paymentTerms: t(d.paymentTerms), otherRef: IM_dgTag_(d) + t(d.otherRef),
    billTo: t(d.billTo), billToGstin: t(d.billToGstin), shipTo: t(d.shipTo), shipToGstin: t(d.shipToGstin),
    vehicleNo: t(d.vehicleNo), driver: t(d.driver), lrNo: t(d.lrNo), eWayBillNo: t(d.eWayBillNo),
    mainItems: main.map(line), hardwareItems: hw.map(line),
    bypassStockCheck: d.bypassStockCheck === true, skipPdf: true
  });
}

// ---------------------------------------------------------------- new versions (edit in DocGen)
/** A new version of a PO: IMS's own "Amend PO" (lines replaced in place, revision counter up). Refused once goods have been received. */
function IM_dgRevisePo_(d, poNo) {
  var lines = d.lineItems, ss = SpreadsheetApp.getActive(), t = IM_dgText_;
  if (!Array.isArray(lines) || !lines.length) throw new Error('At least one line is required.');
  var reg = ss.getSheetByName('Purchase Order Register'), pl = ss.getSheetByName('PO Lines');
  var regRows = reg && reg.getLastRow() >= 2 ? reg.getRange(2, 1, reg.getLastRow() - 1, 7).getValues() : [];
  var head = regRows.filter(function (r) { return r[1] === poNo; })[0] || regRows.filter(function (r) { return IM_dgFamily_(r[1]) === IM_dgFamily_(poNo); })[0];   // DocGen may still hold an older version's number
  if (!head) throw new Error('PO ' + poNo + ' not found in IMS.');
  poNo = head[1];
  if (head[3] !== t(d.supplierCode)) throw new Error('The supplier of ' + poNo + ' cannot be changed. Make a new PO instead.');
  if (pl && pl.getLastRow() >= 2 && pl.getRange(2, 1, pl.getLastRow() - 1, 13).getValues().some(function (r) { return r[0] === poNo && Number(r[12]) > 0; }))
    throw new Error(poNo + ' already has goods received against it, so it cannot be changed from DocGen. Amend it in IMS.');
  var res = amendPo({
    poNo: poNo, date: d.date, deliveryDate: d.deliveryDate || '', remarks: IM_dgTag_(d) + t(d.remarks), skipPdf: true,
    lineItems: lines.map(function (l) { return { itemCode: t(l.itemCode), qty: Number(l.qty), rate: Number(l.rate), gst: Number(l.gst) || 0, discount: Number(l.discount) || 0, discountMode: l.discountMode === 'amt' ? 'amt' : 'pct' }; })
  });
  if (res.needsConfirm) throw new Error(res.message + ' Amend it in IMS.');
  if (!res.success) throw new Error(res.message);
  // The edited PO becomes the next version: TSP/PO/26-27/043 -> 043A -> 043B. It is the SAME PO renamed (no second copy), so the GRN picker
  // lists only the latest version, with its latest lines. (No goods can have been received, that was refused above.)
  var next = IM_dgNextVersion_(poNo), lock = LockService.getDocumentLock(); lock.waitLock(15000);
  try {
    if (regRows.some(function (r) { return r[1] === next; })) throw new Error(next + ' already exists in IMS.');
    var hr = reg.getRange(2, 2, reg.getLastRow() - 1, 1).getValues(), hit = -1;
    hr.forEach(function (r, i) { if (r[0] === poNo) hit = i + 2; });
    if (hit < 0) throw new Error('PO ' + poNo + ' not found in IMS.');
    reg.getRange(hit, 2).setValue(next);
    if (pl && pl.getLastRow() >= 2) {
      var pr = pl.getRange(2, 1, pl.getLastRow() - 1, 1).getValues();
      pr.forEach(function (r, i) { if (r[0] === poNo) pl.getRange(i + 2, 1).setValue(next); });
    }
  } finally { lock.releaseLock(); }
  return { success: true, code: next, message: poNo + ' updated: now version ' + next + '.' };
}

/**
 * A new version of a Delivery Challan: its lines are REPLACED (stock is worked out from the lines, so the old quantities stop counting and
 * the new ones start; the difference is all that changes). Only the extra quantity is checked against stock. Refused once anything was returned.
 */
function IM_dgReviseDc_(d, dcNo) {
  var ss = SpreadsheetApp.getActive(), reg = ss.getSheetByName('FG Delivery Register'), ls = ss.getSheetByName('FG Delivery Lines'), t = IM_dgText_;
  var main = d.mainItems, hw = d.hardwareItems || [];
  if (!Array.isArray(main) || !main.length) throw new Error('At least one main item is required.');
  if (main.length + hw.length > IM_DG_MAX_LINES) throw new Error('Too many lines.');
  dcNo = IM_dgResolve_(reg, 2, dcNo);
  var lock = LockService.getDocumentLock(); lock.waitLock(15000);
  try {
    var regRow = -1;
    if (reg && reg.getLastRow() >= 2) reg.getRange(2, 2, reg.getLastRow() - 1, 1).getValues().some(function (r, i) { if (r[0] === dcNo) { regRow = i + 2; return true; } return false; });
    if (regRow < 0) throw new Error('Challan ' + dcNo + ' not found in IMS.');
    var oldRows = [], oldQty = {}, codes = {};
    if (ls.getLastRow() >= 2) ls.getRange(2, 1, ls.getLastRow() - 1, 10).getValues().forEach(function (r, i) {
      if (r[0] !== dcNo) return;
      oldRows.push(i + 2);
      if (Number(r[9]) > 0) throw new Error(dcNo + ' has goods returned against it, so it cannot be changed from DocGen. Fix it in IMS.');
      if (r[3]) { oldQty[r[3]] = (oldQty[r[3]] || 0) + (Number(r[5]) || 0); codes[r[3]] = 1; }
    });
    var all = [], totQty = 0, totWt = 0;
    main.forEach(function (l) { var q = Number(l.qty) || 0, w = Number(l.weight) || 0; totQty += q; totWt += w; all.push({ section: 'Main', code: t(l.itemCode), desc: t(l.description), qty: q, uom: t(l.uom) || 'Nos', weight: w, remark: t(l.remark) }); });
    hw.forEach(function (l) { var q = Number(l.qty) || 0; if (q <= 0) return; totQty += q; all.push({ section: 'Hardware', code: t(l.itemCode), desc: t(l.description), qty: q, uom: t(l.uom) || 'Nos', weight: 0, remark: t(l.remark) }); });
    all.forEach(function (l) { if (!(l.qty > 0)) throw new Error('Every line needs a quantity above zero.'); if (!l.desc) throw new Error('Every line needs a description.'); });

    // only the EXTRA quantity (new minus what this challan already took) can run the stock short
    if (d.bypassStockCheck !== true) {
      var newQty = {}; all.forEach(function (l) { if (l.code) newQty[l.code] = (newQty[l.code] || 0) + l.qty; });
      var extra = Object.keys(newQty).map(function (c) { return { code: c, qty: newQty[c] - (oldQty[c] || 0) }; }).filter(function (x) { return x.qty > 0; });
      if (extra.length) { var chk = _checkStockAvailability(extra, t(d.project)); if (!chk.ok) return { success: false, needsConfirm: true, warnings: chk.warnings, message: 'Some lines exceed Available stock. Confirm to proceed anyway.' }; }
    }

    oldRows.reverse().forEach(function (r) { ls.deleteRow(r); });
    var start = nextRowByCol(ls, 1);
    ls.getRange(start, 1, all.length, 10).setValues(all.map(function (l, i) { return [dcNo, i + 1, l.section, l.code, l.desc, l.qty, l.uom, l.weight, l.remark, 0]; }));
    reg.getRange(regRow, 3, 1, 15).setValues([[
      parseDate(d.date), t(d.customerName), t(d.project), t(d.paymentTerms), IM_dgTag_(d) + t(d.otherRef),
      t(d.billTo), t(d.billToGstin), t(d.shipTo), t(d.shipToGstin), t(d.vehicleNo), t(d.driver), t(d.lrNo), t(d.eWayBillNo), totQty, totWt
    ]]);

    all.forEach(function (l) { if (l.code) codes[l.code] = 1; });
    // only the EXTRA quantity is a new outward: it consumes the project's open reservations (the rest was handled when the challan was first made)
    var addQ = {}; all.forEach(function (l) { if (l.code) addQ[l.code] = (addQ[l.code] || 0) + l.qty; });
    var addLines = Object.keys(addQ).map(function (c) { return { itemCode: c, qty: addQ[c] - (oldQty[c] || 0) }; }).filter(function (x) { return x.qty > 0; });
    if (addLines.length) IM_consumePmReservations_(t(d.project), addLines, dcNo, null);
    var sfg = ss.getSheetByName('SFG Master'), items = ss.getSheetByName('Item Master');
    var list = Object.keys(codes);
    updateSfgSnapshotsBatch(list.filter(function (c) { return sfg && getRowByCode(sfg, c); }));
    updateItemSnapshotsBatch(list.filter(function (c) { return items && getRowByCode(items, c); }));
    // the edited challan becomes the next version (043 -> 043A -> 043B): the SAME challan renamed (register + lines), so nothing is counted twice
    var nextDc = IM_dgNextVersion_(dcNo);
    IM_dgRename_(reg, 2, dcNo, nextDc); IM_dgRename_(ls, 1, dcNo, nextDc);
    return { success: true, code: nextDc, message: 'Challan ' + dcNo + ' updated: now version ' + nextDc + ' (' + all.length + ' lines).' };
  } finally { lock.releaseLock(); }
}

// ---------------------------------------------------------------- Request for Quotation (record only: IMS keeps no RFQ register)
/** Every code a RFQ line may use: Item Master, SFG Master and FG Master. */
function IM_dgCodes_() {
  var all = {};
  ['Item Master', 'SFG Master', 'FG Master'].forEach(function (n) { var r = IM_dgRows_(n); r.rows.forEach(function (x) { if (x[0]) all[IM_dgText_(x[0])] = true; }); });
  return all;
}
/**
 * An RFQ has no stock effect and IMS has no RFQ register, so "recording" it means: check every line against the masters (the same codes
 * the IMS RFQ form offers) and keep the number in the DocGen log. The number is DocGen's. A new version just re-checks.
 */
function IM_dgRfqCheck_(d) {
  var lines = d.lineItems;
  if (!Array.isArray(lines) || !lines.length) throw new Error('At least one line is required.');
  if (lines.length > IM_DG_MAX_LINES) throw new Error('Too many lines.');
  var codes = IM_dgCodes_();
  lines.forEach(function (l, i) {
    if (!codes[IM_dgText_(l.itemCode)]) throw new Error('Line ' + (i + 1) + ': ' + IM_dgText_(l.itemCode) + ' is not an IMS item, semi-finished or finished goods code.');
    if (!(Number(l.qty) > 0)) throw new Error('Line ' + (i + 1) + ': quantity must be above zero.');
  });
  return lines.length;
}
function IM_dgRfq_(d) {
  var count = IM_dgRfqCheck_(d), n = IM_dgNumber_(d);
  if (!n) {   // IMS issues it: TSP/RFQ/26-27/015
    var lock = LockService.getDocumentLock(); lock.waitLock(15000);
    try { n = IM_docNumber_('RFQ', parseDate(d.date), true); } finally { lock.releaseLock(); }
  }
  return { success: true, code: n, message: 'RFQ ' + n + ' recorded (' + count + ' lines).' };
}
function IM_dgReviseRfq_(d, number) {
  var count = IM_dgRfqCheck_(d), next = IM_dgNextVersion_(number);   // IMS keeps no RFQ rows: the new version is the number in the DocGen log
  return { success: true, code: next, message: 'RFQ ' + number + ' updated: now version ' + next + ' (' + count + ' lines).' };
}

// ---------------------------------------------------------------- Job Work Challan (material sent to a job worker: stock goes down)
function IM_dgJwcLines_(d) {
  var lines = d.lineItems, t = IM_dgText_;
  if (!Array.isArray(lines) || !lines.length) throw new Error('At least one line is required.');
  if (lines.length > IM_DG_MAX_LINES) throw new Error('Too many lines.');
  return lines.map(function (l) { return { sfgCode: t(l.sfgCode), qty: Number(l.qty), expectedReturn: l.expectedReturn || '', remarks: t(l.remarks), weight: Number(l.weight) || 0 }; });
}
function IM_dgJwc_(d) {
  return addJobWorkOutMultiLine({
    manualJwcNo: IM_dgNumber_(d), date: d.date, jobWorkerCode: IM_dgText_(d.jobWorkerCode), lineItems: IM_dgJwcLines_(d),
    bypassStockCheck: d.bypassStockCheck === true, skipPdf: true
  });
}
/**
 * A new version of a Job Work Challan: the challan's lines are REPLACED (stock is worked out from the register rows, so only the
 * difference moves). Only the extra quantity is checked against stock. Refused once anything was received back (status no longer Open)
 * and when the job worker would change. If writing the new lines fails, the old ones are put back.
 * ponytail: not under IMS's document lock while rows are swapped (the write helper takes it); the window is a fraction of a second.
 */
function IM_dgReviseJwc_(d, jwcNo) {
  var ss = SpreadsheetApp.getActive(), reg = ss.getSheetByName('Job Work Outward Register'), lines = IM_dgJwcLines_(d);
  jwcNo = IM_dgResolve_(reg, 2, jwcNo);
  var all = reg && reg.getLastRow() >= 2 ? reg.getRange(2, 1, reg.getLastRow() - 1, 17).getValues() : [];
  var rowNos = [], old = [];
  all.forEach(function (r, i) { if (r[1] === jwcNo) { rowNos.push(i + 2); old.push(r); } });
  if (!old.length) throw new Error('Job Work Challan ' + jwcNo + ' not found in IMS.');
  if (old[0][3] !== IM_dgText_(d.jobWorkerCode)) throw new Error('The job worker of ' + jwcNo + ' cannot be changed. Make a new challan instead.');
  if (old.some(function (r) { return r[13] !== 'Open'; })) throw new Error(jwcNo + ' already has goods received back, so it cannot be changed from DocGen. Fix it in IMS.');
  if (d.bypassStockCheck !== true) {
    var oldQty = {}, newQty = {};
    old.forEach(function (r) { oldQty[r[6]] = (oldQty[r[6]] || 0) + (Number(r[8]) || 0); });
    lines.forEach(function (l) { newQty[l.sfgCode] = (newQty[l.sfgCode] || 0) + l.qty; });
    var extra = Object.keys(newQty).map(function (c) { return { code: c, qty: newQty[c] - (oldQty[c] || 0) }; }).filter(function (x) { return x.qty > 0; });
    if (extra.length) { var chk = _checkStockAvailability(extra); if (!chk.ok) return { success: false, needsConfirm: true, warnings: chk.warnings, message: 'Some lines exceed Available stock. Confirm to proceed anyway.' }; }
  }
  var refresh = function (codes) {
    var sfg = ss.getSheetByName('SFG Master'), items = ss.getSheetByName('Item Master');
    updateSfgSnapshotsBatch(codes.filter(function (c) { return sfg && getRowByCode(sfg, c); }));
    updateItemSnapshotsBatch(codes.filter(function (c) { return !(sfg && getRowByCode(sfg, c)) && items && getRowByCode(items, c); }));
  };
  rowNos.slice().reverse().forEach(function (r) { reg.deleteRow(r); });
  var res;
  try {
    res = postJobWorkOutwardRows({ manualJwcNo: jwcNo, date: d.date, jobWorkerCode: IM_dgText_(d.jobWorkerCode), lineItems: lines, skipPdf: true });
  } catch (e) {
    reg.getRange(reg.getLastRow() + 1, 1, old.length, old[0].length).setValues(old);   // put the old lines back
    refresh(old.map(function (r) { return r[6]; }));
    throw e;
  }
  var gone = old.map(function (r) { return r[6]; }).filter(function (c) { return !lines.some(function (l) { return l.sfgCode === c; }); });
  if (gone.length) refresh(gone);
  var nextJ = IM_dgNextVersion_(jwcNo); IM_dgRename_(reg, 2, jwcNo, nextJ);   // the SAME challan renamed to its next version
  return { success: true, code: nextJ, message: 'Job Work Challan ' + jwcNo + ' updated: now version ' + nextJ + ' (' + res.rows + ' lines).' };
}

// ---------------------------------------------------------------- photos (pictures inside sheet cells, shrunk for DocGen)
/**
 * code -> photo address, for ONLY the codes asked for. Reading every picture cell of the three masters (what getMasterPhotoUrlMap does for
 * the PDFs) takes about 50 seconds, which is the whole time DocGen is willing to wait. Here only the matching rows are read, and the answer
 * is remembered: found 6 hours, "no photo" 10 minutes (so a photo added in IMS shows up soon).
 */
function IM_dgPhotoUrls_(codes) {
  var cache = CacheService.getScriptCache(), urls = {}, todo = [];
  var hit = cache.getAll(codes.map(function (c) { return 'dgu_' + c; }));
  codes.forEach(function (c) { var v = hit['dgu_' + c]; if (v === undefined || v === null) todo.push(c); else if (v !== '-') urls[c] = v; });
  if (!todo.length) return urls;
  var ss = SpreadsheetApp.getActive(), left = todo.slice();
  ['Item Master', 'SFG Master', 'FG Master'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!left.length || !sh || sh.getLastRow() < 2) return;
    var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    left = left.filter(function (c) {
      for (var i = 0; i < ids.length; i++) {
        if (ids[i][0] !== c) continue;
        try {
          var p = sh.getRange(i + 2, 2).getValue();
          if (p && typeof p === 'object') { var u = (typeof p.getUrl === 'function' && p.getUrl()) || (typeof p.getContentUrl === 'function' && p.getContentUrl()); if (u) urls[c] = u; }
        } catch (e) { /* no usable picture in this cell */ }
        return false;   // found in this master: done with this code
      }
      return true;
    });
  });
  var found = {}, none = {};
  todo.forEach(function (c) { if (urls[c]) found['dgu_' + c] = urls[c]; else none['dgu_' + c] = '-'; });
  try {
    if (Object.keys(found).length) cache.putAll(found, 21600);
    if (Object.keys(none).length) cache.putAll(none, 600);
  } catch (e) { /* could not cache: just slower next time */ }
  return urls;
}

/** codes -> { code: 'data:image/...;base64,...' }, about 600px wide. A code without a photo is simply left out. */
function IM_dgPhotos_(d) {
  var codes = (Array.isArray(d.codes) ? d.codes : []).map(IM_dgText_).filter(function (c) { return /^[A-Za-z0-9_-]{1,30}$/.test(c); });
  if (!codes.length) return { ok: false, error: 'codes is required.' };
  if (codes.length > 40) return { ok: false, error: 'At most 40 codes per call.' };
  var urls = IM_dgPhotoUrls_(codes);
  _prefetchImages_(Object.keys(urls).map(function (c) { return urls[c]; }));
  var photos = {};
  Object.keys(urls).forEach(function (c) { var u = _driveImageDataUrl(urls[c]); if (u && u.length < 150000) photos[c] = u; });
  return { ok: true, photos: photos };
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
