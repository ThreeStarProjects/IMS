/**
 * 3 Star Projects IMS v3 — 50_Docs.gs
 * Tax-invoice-format PDF generation. Brand + logo + GST + bank + T&C from Config.
 *
 * Layout (all 5 transaction docs):
 *   [LOGO]                             [COMPANY NAME]
 *   Address / Email / Phone            GSTIN | PAN
 *   ─────────────────────────────────────────────
 *                    DOC TITLE
 *                    DOC-0001                            Date: dd-mm-yyyy
 *   [Party block: Supplier / Job Worker / Site — GSTIN, State, Address]
 *   [Line items table — HSN, qty, rate, disc, CGST/SGST or IGST, total]
 *   [Grand Total row]
 *   Amount in Words: Rupees ...
 *   ─────────────────────────────────────────────
 *   [Bank Details block]
 *   [Terms & Conditions block]
 *   For 3 Star Projects           For [Vendor]
 *   (Authorised Signatory)        (Received By)
 *   Generated timestamp — footer
 *
 * Deps: 00_Config_Menu_Setup.gs, 30_Snapshots.gs, 60_Utils_Audit_Triggers.gs.
 */

// ============================================================================
// Branding + Config accessors
// ============================================================================

function _brand() {
  return {
    company:   getConfig('Company Name')      || '3 Star Projects',
    email:     getConfig('Company Email')     || '',
    address:   getConfig('Company Address')   || '',
    phone:     getConfig('Company Phone')     || '',
    logoUrl:   getConfig('Company Logo URL')  || '',
    primary:   getConfig('Brand Primary Color')   || '#1F4E79',
    secondary: getConfig('Brand Secondary Color') || '#F47C2C',
    state:     getConfig('Company State')     || '',
    gstin:     getConfig('Company GSTIN')     || '',
    pan:       getConfig('Company PAN')       || '',
    bankName:  getConfig('Bank Name')         || '',
    bankAcct:  getConfig('Bank Account No')   || '',
    bankIfsc:  getConfig('Bank IFSC')         || '',
    bankBranch:getConfig('Bank Branch')       || '',
    terms:     getConfig('Terms and Conditions') || '',
    poTerms:   getConfig('PO Terms') || _DEFAULT_PO_TERMS,
  };
}

// Default T&Cs printed on every PO PDF. User can override by adding a "PO Terms"
// row to Config (multi-line, one clause per line, blank line = new paragraph).
const _DEFAULT_PO_TERMS = [
  '1. Material shall strictly conform to the agreed specifications, drawings, samples and quality standards.',
  '2. Three Star Projects Pvt. Ltd. reserves the right to inspect the material before dispatch and/or at the site; rejected material shall be returned at supplier\'s cost.',
  '3. Material shall be properly packed to avoid damage during transportation and handling.',
  '4. Supplier shall provide the proper GST Tax Invoice / Delivery Challan along with the material.',
  '5. Any damaged, defective or incorrect material received at site shall be replaced by the supplier at no additional cost.',
  '6. Delivery shall be made as per the agreed project schedule and site requirement.',
  '7. Payment Terms: [30/45/60] days from receipt of material and valid GST invoice, as mutually agreed.',
  '8. Any change in material specification, design/code or quantity shall require prior written approval from Three Star Projects Pvt. Ltd.',
  '9. The supplier shall mention our PO Number on the invoice and delivery documents.',
  '10. This Purchase Order is subject to the jurisdiction of courts at Thane / Mumbai.',
].join('\n');

function _htmlEscape(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// Format qty — drop trailing zeros so "10.00" shows as "10", "4.4652" stays as-is.
function _fmtQty(n) {
  const num = Number(n) || 0;
  if (Number.isInteger(num)) return String(num);
  return String(Math.round(num * 10000) / 10000);
}

// Indian numbering: 12,34,567.89
function _money(n) {
  const num = Number(n) || 0;
  const parts = num.toFixed(2).split('.');
  const x = parts[0]; const last3 = x.slice(-3); const rest = x.slice(0, -3);
  const grouped = rest ? rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + last3 : last3;
  return grouped + '.' + parts[1];
}

// ============================================================================
// Logo — fetch once, cache base64 in ScriptProperties keyed by file ID.
// This is why the logo actually renders: HtmlService PDF doesn't fetch
// remote URLs reliably, but inlined base64 does.
// ============================================================================

function _logoDataUrl() { return _driveImageDataUrl(getConfig('Company Logo URL')); }

// Fetch any image URL, return base64 data URL. Handles two sources:
//   1. Drive share links  → DriveApp (works for private files)
//   2. Any other https URL → UrlFetchApp (works for Cloudinary, S3, CDN, etc.)
// Cached in CacheService (100KB per value, 6hr TTL) so multi-line PDFs
// don't re-fetch the same image N times.
function _driveImageDataUrl(url) {
  if (!url) return '';
  const trimmed = String(url).trim();
  const cache = CacheService.getScriptCache();
  // Cache key must be a proper hash of the FULL URL. Previously used
  // base64(url).substr(0,60) — that's just truncated encoding, not a hash. Every
  // photo URL from Google Sheets shares a long common prefix (lh3.googleusercontent.com
  // or docs.google.com/spreadsheets/...), so different URLs collided on the same
  // key and served the wrong image. MD5 fixes that — same input → same key,
  // different input → different key, always.
  // Version-salt the cache — bumping _photoCacheEpoch in ScriptProperties invalidates
  // every stored image entry immediately (used by clearPhotoCache below).
  const epoch = PropertiesService.getScriptProperties().getProperty('_photoCacheEpoch') || '0';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, epoch + '|' + trimmed);
  const cacheKey = 'img_' + Utilities.base64EncodeWebSafe(digest);
  const cached = cache.get(cacheKey);
  if (cached) return cached;
  try {
    let blob;
    const driveId = _extractDriveId(trimmed);
    if (driveId) {
      // Drive file — auth'd fetch, works for private/shared files
      blob = DriveApp.getFileById(driveId).getBlob();
    } else if (/^https?:\/\//i.test(trimmed)) {
      // Public URL — direct HTTPS fetch (Cloudinary, S3, etc.)
      const response = UrlFetchApp.fetch(trimmed, { muteHttpExceptions: true, followRedirects: true });
      if (response.getResponseCode() !== 200) throw new Error('HTTP ' + response.getResponseCode() + ' for ' + trimmed);
      blob = response.getBlob();
    } else {
      return '';
    }
    const mime = blob.getContentType() || 'image/png';
    const b64 = Utilities.base64Encode(blob.getBytes());
    const dataUrl = 'data:' + mime + ';base64,' + b64;
    if (dataUrl.length < 100000) cache.put(cacheKey, dataUrl, 21600);
    return dataUrl;
  } catch (e) {
    return '';
  }
}

function _extractDriveId(url) {
  // Handles /d/FILE_ID/view, ?id=FILE_ID, /file/d/FILE_ID/
  const m1 = /\/d\/([a-zA-Z0-9_-]+)/.exec(url);
  if (m1) return m1[1];
  const m2 = /[?&]id=([a-zA-Z0-9_-]+)/.exec(url);
  if (m2) return m2[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(url.trim())) return url.trim();
  return '';
}

// ============================================================================
// GST helpers — determine intra-state (CGST+SGST) vs inter-state (IGST).
// State code = first 2 chars of GSTIN. Company state name → code via table.
// ============================================================================

const GST_STATE_CODES = {
  'JAMMU AND KASHMIR': '01', 'HIMACHAL PRADESH': '02', 'PUNJAB': '03', 'CHANDIGARH': '04',
  'UTTARAKHAND': '05', 'HARYANA': '06', 'DELHI': '07', 'RAJASTHAN': '08',
  'UTTAR PRADESH': '09', 'BIHAR': '10', 'SIKKIM': '11', 'ARUNACHAL PRADESH': '12',
  'NAGALAND': '13', 'MANIPUR': '14', 'MIZORAM': '15', 'TRIPURA': '16',
  'MEGHALAYA': '17', 'ASSAM': '18', 'WEST BENGAL': '19', 'JHARKHAND': '20',
  'ODISHA': '21', 'CHHATTISGARH': '22', 'MADHYA PRADESH': '23', 'GUJARAT': '24',
  'DADRA AND NAGAR HAVELI AND DAMAN AND DIU': '26', 'MAHARASHTRA': '27',
  'KARNATAKA': '29', 'GOA': '30', 'LAKSHADWEEP': '31', 'KERALA': '32',
  'TAMIL NADU': '33', 'PUDUCHERRY': '34', 'ANDAMAN AND NICOBAR ISLANDS': '35',
  'TELANGANA': '36', 'ANDHRA PRADESH': '37', 'LADAKH': '38',
};

function _companyStateCode(brand) {
  if (!brand.state) return '';
  return GST_STATE_CODES[brand.state.toUpperCase()] || '';
}

function _supplierStateCode(supplierGstin) {
  if (!supplierGstin) return '';
  return String(supplierGstin).substr(0, 2);
}

// intraState=true → split as CGST/SGST at half rate each. false → single IGST at full rate.
function _isIntraState(brand, partyGstin) {
  const cc = _companyStateCode(brand);
  const pc = _supplierStateCode(partyGstin);
  if (!cc || !pc) return true; // default: assume intra-state if we can't tell
  return cc === pc;
}

// ============================================================================
// Amount in words — Indian numbering (crores/lakhs/thousands).
// ============================================================================

function _amountInWords(n) {
  const num = Math.round(Number(n) * 100) / 100;
  const rupees = Math.floor(num);
  const paise = Math.round((num - rupees) * 100);
  let out = 'Rupees ' + _intToWords(rupees);
  if (paise) out += ' and ' + _intToWords(paise) + ' Paise';
  return out + ' Only';
}

function _intToWords(n) {
  if (n === 0) return 'Zero';
  const ones = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine','Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen'];
  const tens = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
  function twoDigit(x) {
    if (x < 20) return ones[x];
    return tens[Math.floor(x / 10)] + (x % 10 ? ' ' + ones[x % 10] : '');
  }
  function threeDigit(x) {
    if (x === 0) return '';
    if (x < 100) return twoDigit(x);
    return ones[Math.floor(x / 100)] + ' Hundred' + (x % 100 ? ' ' + twoDigit(x % 100) : '');
  }
  const parts = [];
  const crore  = Math.floor(n / 10000000); n = n % 10000000;
  const lakh   = Math.floor(n / 100000);   n = n % 100000;
  const thou   = Math.floor(n / 1000);     n = n % 1000;
  const rest   = n;
  if (crore) parts.push(twoDigit(crore) + ' Crore');
  if (lakh)  parts.push(twoDigit(lakh) + ' Lakh');
  if (thou)  parts.push(twoDigit(thou) + ' Thousand');
  if (rest)  parts.push(threeDigit(rest));
  return parts.join(' ').trim();
}

// ============================================================================
// Shared HTML shell — variables are optional; empty ones auto-skip their block.
// docSpec: {
//   title, docNo, dateStr,
//   partyBlock: { label, name, code, gstin, state, address, contact },
//   cols, rows, totalsRow?, amountInWords?, taxSummary?,
//   showBankDetails, showTerms,
//   signatures: [['For 3 Star Projects','(Authorised Signatory)'], ['For Vendor','(Received By)']],
//   footerNote?
// }
// ============================================================================

function _buildDocHtml(docSpec) {
  const b = _brand();
  const logo = _logoDataUrl();
  const primary = b.primary || '#1F4E79';

  const logoImg = logo
    ? '<img src="' + logo + '" class="logo">'
    : '<div class="logo-fallback">' + _htmlEscape(b.company.substring(0, 3).toUpperCase()) + '</div>';

  // Header — logo left, company info right (Ramdev-style)
  const headerHtml =
    '<div class="hdr">' +
      '<div class="hdr-logo">' + logoImg + '</div>' +
      '<div class="hdr-info">' +
        '<div class="hdr-name">' + _htmlEscape(b.company) + '</div>' +
        (b.address ? '<div class="hdr-line">' + _htmlEscape(b.address).replace(/\n/g, ' ') + '</div>' : '') +
        '<div class="hdr-ids">' +
          (b.phone ? '<span><b>Mobile:</b> ' + _htmlEscape(b.phone) + '</span>' : '') +
          (b.gstin ? '<span><b>GSTIN:</b> ' + _htmlEscape(b.gstin) + '</span>' : '') +
          (b.pan ? '<span><b>PAN Number:</b> ' + _htmlEscape(b.pan) + '</span>' : '') +
        '</div>' +
        (b.email ? '<div class="hdr-ids"><span><b>Email:</b> ' + _htmlEscape(b.email) + '</span></div>' : '') +
      '</div>' +
    '</div>' +
    '<div class="hdr-divider"></div>';

  // Doc metadata band — grey background with No / Date / (optional Expiry)
  const metaFields = [
    { label: (docSpec.docNoLabel || 'Document No.'), value: docSpec.docNo || '' },
    { label: (docSpec.dateLabel  || 'Date'),         value: docSpec.dateStr || '' },
  ];
  if (docSpec.extraDateLabel && docSpec.extraDateStr) {
    metaFields.push({ label: docSpec.extraDateLabel, value: docSpec.extraDateStr });
  }
  const metaHtml = '<div class="meta-band">' +
    metaFields.map(function(f) {
      return '<div class="meta-cell"><span class="meta-lbl">' + _htmlEscape(f.label) + ':</span> <b>' + _htmlEscape(f.value) + '</b></div>';
    }).join('') +
    '</div>';

  // Party blocks — supports single {label,name,...} or [Bill To, Ship To]
  const parties = Array.isArray(docSpec.partyBlock) ? docSpec.partyBlock
                : (docSpec.partyBlock ? [docSpec.partyBlock] : []);
  const partyHtml = parties.length ? '<div class="parties">' +
    parties.map(function(p) {
      return '<div class="party">' +
        '<div class="party-lbl">' + _htmlEscape(p.label) + '</div>' +
        '<div class="party-name">' + _htmlEscape(p.name) + '</div>' +
        (p.address ? '<div class="party-line">' + _htmlEscape(p.address).replace(/\n/g, '<br>') + '</div>' : '') +
        (p.gstin   ? '<div class="party-line">GSTIN: <b>' + _htmlEscape(p.gstin) + '</b></div>' : '') +
        (p.pan     ? '<div class="party-line">PAN Number: <b>' + _htmlEscape(p.pan) + '</b></div>' : '') +
        (p.state   ? '<div class="party-line">Place of Supply: <b>' + _htmlEscape(p.state) + '</b></div>' : '') +
        (p.contact ? '<div class="party-line">' + _htmlEscape(p.contact) + '</div>' : '') +
      '</div>';
    }).join('') + '</div>' : '';

  // Sections — supports multi-table docs (e.g., FG DC has Main + Hardware).
  // Legacy single-table callers still work via the fallback below.
  const sections = docSpec.sections || [{
    title: '', cols: docSpec.cols, rows: docSpec.rows,
    subtotalRow: docSpec.subtotalRow, numericColIdxs: docSpec.numericColIdxs
  }];
  function _renderSection(sec) {
    const numIdxs = sec.numericColIdxs || [];
    const thead = '<tr>' + (sec.cols || []).map(function(c, i) {
      const align = numIdxs.indexOf(i) >= 0 ? ' class="r"' : '';
      return '<th' + align + '>' + _htmlEscape(c) + '</th>';
    }).join('') + '</tr>';
    const tbody = (sec.rows || []).map(function(r) {
      return '<tr>' + r.map(function(cell, i) {
        const align = numIdxs.indexOf(i) >= 0 ? ' class="r"' : '';
        if (typeof cell === 'string' && cell.indexOf('data:') === 0) {
          return '<td class="thumb"><img src="' + cell + '"></td>';
        }
        if (typeof cell === 'string' && cell.indexOf('<') === 0) {
          return '<td' + align + '>' + cell + '</td>';
        }
        return '<td' + align + '>' + _htmlEscape(cell) + '</td>';
      }).join('') + '</tr>';
    }).join('');
    const subtotal = sec.subtotalRow
      ? '<tr class="subtotal">' + sec.subtotalRow.map(function(c, i) {
          const align = numIdxs.indexOf(i) >= 0 ? ' class="r"' : '';
          return '<td' + align + '>' + (typeof c === 'string' && c.indexOf('<') === 0 ? c : _htmlEscape(c)) + '</td>';
        }).join('') + '</tr>'
      : '';
    return (sec.title ? '<div class="section-title">' + _htmlEscape(sec.title) + '</div>' : '') +
           '<table class="items"><thead>' + thead + '</thead>' +
           '<tbody>' + tbody + subtotal + '</tbody></table>';
  }
  const tablesHtml = sections.map(_renderSection).join('');
  const preTableHtml = docSpec.preTableBlock || '';
  const postTableHtml = docSpec.postTableBlock || '';

  // Totals block (right side)
  const totalsHtml = docSpec.taxSummary ? (function() {
    const t = docSpec.taxSummary;
    const rows = [];
    rows.push({ lbl: 'Taxable Amount', val: 'Rs. ' + _money(t.taxable) });
    if (t.intra) {
      const cgstPct = t.taxable > 0 ? Math.round(t.cgst / t.taxable * 100 * 100) / 100 : 9;
      const sgstPct = t.taxable > 0 ? Math.round(t.sgst / t.taxable * 100 * 100) / 100 : 9;
      rows.push({ lbl: 'CGST @' + cgstPct + '%', val: 'Rs. ' + _money(t.cgst) });
      rows.push({ lbl: 'SGST @' + sgstPct + '%', val: 'Rs. ' + _money(t.sgst) });
    } else {
      const igstPct = t.taxable > 0 ? Math.round(t.igst / t.taxable * 100 * 100) / 100 : 18;
      rows.push({ lbl: 'IGST @' + igstPct + '%', val: 'Rs. ' + _money(t.igst) });
    }
    // Round Off — difference between rounded grand total and exact sum
    const exact = t.total;
    const rounded = Math.round(exact);
    const roundOff = rounded - exact;
    if (Math.abs(roundOff) > 0.001) {
      rows.push({ lbl: 'Round Off', val: (roundOff < 0 ? '- ' : '') + 'Rs. ' + _money(Math.abs(roundOff)) });
    }
    rows.push({ lbl: '<b>Total Amount</b>', val: '<b>Rs. ' + _money(rounded) + '</b>', highlight: true });
    return '<table class="totals-tbl"><tbody>' +
      rows.map(function(r) {
        return '<tr' + (r.highlight ? ' class="grand"' : '') + '><td>' + r.lbl + '</td><td class="r">' + r.val + '</td></tr>';
      }).join('') +
      '</tbody></table>';
  })() : '';

  // Bank details block (left)
  const bankHtml = (docSpec.showBankDetails !== false && b.bankName) ? (
    '<div class="bank-block">' +
      '<div class="fb-title">Bank Details</div>' +
      '<table class="bank-tbl">' +
        '<tr><td>Name:</td><td><b>' + _htmlEscape(b.company) + '</b></td></tr>' +
        (b.bankIfsc ? '<tr><td>IFSC Code:</td><td>' + _htmlEscape(b.bankIfsc) + '</td></tr>' : '') +
        (b.bankAcct ? '<tr><td>Account No:</td><td>' + _htmlEscape(b.bankAcct) + '</td></tr>' : '') +
        '<tr><td>Bank:</td><td>' + _htmlEscape(b.bankName) + (b.bankBranch ? ', ' + _htmlEscape(b.bankBranch) : '') + '</td></tr>' +
      '</table>' +
    '</div>'
  ) : '';

  const wordsHtml = docSpec.amountInWords ?
    '<div class="words-block"><div class="words-lbl">Total Amount (in words)</div>' +
    '<div class="words-val">' + _htmlEscape(docSpec.amountInWords) + '</div></div>' : '';

  const termsText = docSpec.customTerms || b.terms || '';
  const termsHtml = (docSpec.showTerms !== false && termsText) ? (
    '<div class="terms-block">' +
      '<div class="fb-title">Terms and Conditions</div>' +
      '<div class="terms-body">' + _htmlEscape(termsText).replace(/\\n/g, '<br>').replace(/\n/g, '<br>') + '</div>' +
    '</div>'
  ) : '';

  // Signature block (bottom right)
  const sigHtml = docSpec.showSignature !== false ? (
    '<div class="sig-block">' +
      (logo ? '<img src="' + logo + '" class="sig-logo">' : '') +
      '<div class="sig-title">AUTHORISED SIGNATORY FOR</div>' +
      '<div class="sig-company">' + _htmlEscape(b.company) + '</div>' +
    '</div>'
  ) : '';

  const bottomRow = '<div class="footers">' +
    '<div class="footers-l">' + bankHtml + '</div>' +
    '<div class="footers-r">' + totalsHtml + '</div>' +
  '</div>';

  const veryBottom = '<div class="footers">' +
    '<div class="footers-l">' + termsHtml + '</div>' +
    '<div class="footers-r">' + wordsHtml + sigHtml + '</div>' +
  '</div>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:10pt;color:#222;margin:0.25in 0.3in;line-height:1.4;}' +
    // Header
    '.hdr{display:table;width:100%;margin-bottom:2px;}' +
    '.hdr-logo{display:table-cell;width:110px;vertical-align:top;padding-right:14px;}' +
    '.hdr-info{display:table-cell;vertical-align:top;}' +
    '.logo{max-height:90px;max-width:100px;}' +
    '.logo-fallback{font-size:32pt;font-weight:bold;color:' + primary + ';border:2px solid ' + primary + ';padding:8px 14px;display:inline-block;}' +
    '.hdr-name{font-size:22pt;font-weight:bold;color:' + primary + ';line-height:1.1;margin-bottom:4px;}' +
    '.hdr-line{font-size:9.5pt;color:#555;line-height:1.35;}' +
    '.hdr-ids{font-size:9.5pt;color:#222;margin-top:3px;}' +
    '.hdr-ids span{margin-right:20px;}' +
    '.hdr-divider{border-top:3px solid ' + primary + ';margin:8px 0 10px;}' +
    // Meta band
    '.meta-band{background:#EFEFEF;padding:8px 12px;margin-bottom:14px;display:table;width:100%;box-sizing:border-box;}' +
    '.meta-cell{display:table-cell;padding-right:20px;font-size:10pt;}' +
    '.meta-lbl{color:#333;font-weight:bold;}' +
    // Party
    '.parties{display:table;width:100%;margin-bottom:14px;table-layout:fixed;}' +
    '.party{display:table-cell;vertical-align:top;padding-right:16px;}' +
    '.party-lbl{font-size:9pt;color:#666;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;}' +
    '.party-name{font-size:13pt;font-weight:bold;color:#222;margin-bottom:4px;}' +
    '.party-line{font-size:9.5pt;color:#333;line-height:1.35;}' +
    // Items table
    'table.items{width:100%;border-collapse:collapse;margin:0 0 6px;font-size:9.5pt;}' +
    'table.items th{color:' + primary + ';padding:8px 6px;text-align:left;font-size:9pt;text-transform:uppercase;letter-spacing:0.5px;font-weight:bold;border-bottom:2px solid ' + primary + ';}' +
    'table.items th.r{text-align:right;}' +
    'table.items td{padding:8px 6px;border-bottom:1px solid #E5E5E5;vertical-align:top;}' +
    'table.items td.r{text-align:right;}' +
    'table.items td.thumb{padding:2px;text-align:center;width:44px;}' +
    'table.items td.thumb img{max-height:38px;max-width:40px;object-fit:contain;}' +
    'table.items tr.subtotal td{border-top:1px solid #999;border-bottom:none;padding:8px 6px;font-weight:bold;color:#222;font-size:10pt;background:#FAFAFA;}' +
    '.section-title{background:' + primary + ';color:#fff;padding:6px 10px;margin:12px 0 0;font-weight:bold;font-size:10pt;letter-spacing:0.5px;text-transform:uppercase;}' +
    '.section-title + table.items th{border-top:none;}' +
    '.transport-block{border:1px solid #DDD;padding:8px 12px;margin:8px 0;background:#F9F9F9;font-size:9.5pt;line-height:1.5;}' +
    '.transport-block b{color:' + primary + ';}' +
    '.tax-cell{line-height:1.15;}' +
    '.tax-cell .pct{color:#888;font-size:8.5pt;}' +
    // Footers
    '.footers{display:table;width:100%;margin-top:14px;}' +
    '.footers-l{display:table-cell;width:55%;vertical-align:top;padding-right:20px;}' +
    '.footers-r{display:table-cell;width:45%;vertical-align:top;}' +
    // Bank
    '.bank-block .fb-title{font-weight:bold;font-size:10pt;color:#222;margin-bottom:6px;}' +
    'table.bank-tbl{border-collapse:collapse;font-size:9.5pt;}' +
    'table.bank-tbl td{padding:2px 12px 2px 0;color:#333;vertical-align:top;}' +
    // Totals block
    'table.totals-tbl{width:100%;border-collapse:collapse;font-size:10pt;}' +
    'table.totals-tbl td{padding:5px 8px;color:#333;}' +
    'table.totals-tbl td.r{text-align:right;font-variant-numeric:tabular-nums;}' +
    'table.totals-tbl tr.grand td{border-top:1px solid #333;font-size:11pt;padding-top:8px;}' +
    // Terms
    '.terms-block .fb-title{font-weight:bold;font-size:10pt;color:#222;margin-bottom:4px;}' +
    '.terms-body{font-size:9.5pt;color:#333;line-height:1.5;}' +
    // Words + Signature
    '.words-block{margin-bottom:20px;text-align:right;}' +
    '.words-lbl{font-size:9.5pt;color:#333;}' +
    '.words-val{font-size:10pt;color:#111;margin-top:2px;}' +
    '.sig-block{text-align:right;margin-top:20px;}' +
    '.sig-logo{max-height:60px;max-width:80px;margin-bottom:4px;}' +
    '.sig-title{font-weight:bold;font-size:9pt;color:#222;letter-spacing:0.5px;}' +
    '.sig-company{font-weight:bold;font-size:10pt;color:#222;}' +
    '</style></head><body>' +
    headerHtml +
    metaHtml +
    partyHtml +
    preTableHtml +
    tablesHtml +
    postTableHtml +
    bottomRow +
    veryBottom +
    (docSpec.footerNote ? '<div style="margin-top:12px;font-size:9pt;color:#666;font-style:italic;">' + _htmlEscape(docSpec.footerNote) + '</div>' : '') +
    '</body></html>';
}

// ============================================================================
// Drive save
// ============================================================================

function savePdfToDrive(html, docNo, docType) {
  const rootId = getConfig('PDF Docs Folder ID');
  const root = (rootId && rootId.trim())
    ? DriveApp.getFolderById(rootId.trim())
    : DriveApp.getRootFolder();
  const now = new Date();
  const yearMonth = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
  const typeFolder = _getOrCreateFolder(root, docType);
  const monthFolder = _getOrCreateFolder(typeFolder, yearMonth);
  const filename = docType + '_' + docNo + '_' + dateString(now) + '.pdf';
  const pdfBlob = Utilities.newBlob(html, 'text/html', filename)
    .getAs('application/pdf').setName(filename);
  const file = monthFolder.createFile(pdfBlob);
  const id = file.getId();
  return {
    url: file.getUrl(),
    fileId: id,
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + id,
  };
}

function _getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// Update the PDF link column of every row matching docNo. Used after PDF gen.
function writeBackPdfLink(sheet, docNoCol, docNo, linkCol, url) {
  const rows = getRowsByValue(sheet, docNoCol, docNo);
  rows.forEach(function(r) {
    sheet.getRange(r, linkCol).setFormula('=HYPERLINK("' + url + '","View PDF")');
  });
}

// ============================================================================
// 5 transaction PDFs
// ============================================================================

function generateGrnPdf(grnNo, formData, lineRows) {
  const b = _brand();
  const intra = _isIntraState(b, formData.supplierGstin);

  const cols = ['ITEMS','HSN','QTY','RATE','DISC %','TAX','AMOUNT'];

  let sumTaxable = 0, sumCgst = 0, sumSgst = 0, sumIgst = 0, sumQty = 0, sumTax = 0, sumAmount = 0;
  const rows = lineRows.map(function(l) {
    const qty = Number(l.qty) || 0;
    const rate = Number(l.rate) || 0;
    const discPct = Number(l.discPct) || 0;
    const gstPct  = Number(l.gstPct)  || 0;
    const gross = qty * rate;
    const disc  = gross * discPct / 100;
    const taxable = gross - disc;
    const gstAmt = taxable * gstPct / 100;
    const amount = taxable + gstAmt;  // amount with tax
    sumTaxable += taxable; sumQty += qty; sumTax += gstAmt; sumAmount += amount;
    if (intra) { const half = gstAmt / 2; sumCgst += half; sumSgst += half; }
    else       { sumIgst += gstAmt; }
    return [
      _htmlEscape(l.itemDesc || l.itemCode),
      l.hsn || '—',
      _fmtQty(qty) + ' ' + (l.uom || ''),
      _money(rate),
      discPct + ' %',
      '<div class="tax-cell">' + _money(gstAmt) + '<br><span class="pct">(' + gstPct + '%)</span></div>',
      _money(amount),
    ];
  });

  const sumTotal = sumTaxable + sumCgst + sumSgst + sumIgst;
  const taxSummary = { intra: intra, taxable: sumTaxable, cgst: sumCgst, sgst: sumSgst, igst: sumIgst, total: sumTotal };

  return _renderAndSave({
    title: 'Goods Receipt Note (GRN)',
    docNo: grnNo, docNoLabel: 'GRN No.',
    dateStr: dateString(formData.date), dateLabel: 'GRN Date',
    partyBlock: {
      label: 'Supplier',
      name: formData.supplierName || formData.supplierCode,
      gstin: formData.supplierGstin || '',
      state: formData.supplierState || '',
      address: formData.supplierAddress || '',
      contact: [formData.supplierContact, formData.supplierMobile, formData.supplierEmail].filter(String).join(' | '),
    },
    cols: cols,
    numericColIdxs: [2, 3, 4, 5, 6],
    rows: rows,
    subtotalRow: ['SUBTOTAL', '', _fmtQty(sumQty), '', '', _money(sumTax), _money(sumAmount)],
    taxSummary: taxSummary,
    amountInWords: _amountInWords(Math.round(sumTotal)) + ' Only',
    footerNote: formData.invoiceNo ? 'Supplier Invoice No: ' + formData.invoiceNo : '',
  }, grnNo, 'GRN');
}

function generatePoPdf(poNo, formData, lineRows) {
  const b = _brand();
  const intra = _isIntraState(b, formData.supplierGstin);

  const cols = ['ITEMS','HSN','QTY','RATE','DISC %','TAX','AMOUNT'];

  let sumTaxable = 0, sumCgst = 0, sumSgst = 0, sumIgst = 0, sumQty = 0, sumTax = 0, sumAmount = 0;
  const rows = lineRows.map(function(l) {
    const qty = Number(l.qty) || 0;
    const rate = Number(l.rate) || 0;
    const discPct = Number(l.discPct) || 0;
    const gstPct  = Number(l.gstPct)  || 0;
    const gross = qty * rate;
    const disc  = gross * discPct / 100;
    const taxable = gross - disc;
    const gstAmt = taxable * gstPct / 100;
    const amount = taxable + gstAmt;
    sumTaxable += taxable; sumQty += qty; sumTax += gstAmt; sumAmount += amount;
    if (intra) { const half = gstAmt / 2; sumCgst += half; sumSgst += half; }
    else       { sumIgst += gstAmt; }
    return [
      _htmlEscape(l.itemDesc || l.itemCode),
      l.hsn || '—',
      _fmtQty(qty) + ' ' + (l.uom || ''),
      _money(rate),
      discPct + ' %',
      '<div class="tax-cell">' + _money(gstAmt) + '<br><span class="pct">(' + gstPct + '%)</span></div>',
      _money(amount),
    ];
  });

  const sumTotal = sumTaxable + sumCgst + sumSgst + sumIgst;
  const taxSummary = { intra: intra, taxable: sumTaxable, cgst: sumCgst, sgst: sumSgst, igst: sumIgst, total: sumTotal };

  // Three-block party layout: Supplier (vendor) + Bill To (us) + Ship To (us or site).
  // Bill To defaults to company HQ from Config. Ship To defaults to same; if a
  // future PO form adds a "Ship To" field, formData.shipToAddress overrides.
  const supplierContactLine = [formData.supplierContact, formData.supplierMobile, formData.supplierEmail].filter(String).join(' | ');
  const parties = [
    { label: 'SUPPLIER',
      name: formData.supplierName || formData.supplierCode,
      address: formData.supplierAddress || '',
      gstin: formData.supplierGstin || '',
      contact: supplierContactLine },
    { label: 'BILL TO',
      name:    formData.billToName    || b.company,
      address: formData.billToAddress || b.address || '',
      gstin:   formData.billToGstin   || b.gstin || '',
      pan:     b.pan || '',
      contact: formData.billToContact || [b.phone, b.email].filter(String).join(' | ') },
    { label: 'SHIP TO',
      name:    formData.shipToName    || formData.billToName    || b.company,
      address: formData.shipToAddress || formData.billToAddress || b.address || '',
      gstin:   formData.shipToGstin   || formData.billToGstin   || b.gstin || '',
      contact: formData.shipToContact || '' },
  ];

  return _renderAndSave({
    title: 'Purchase Order',
    docNo: poNo, docNoLabel: 'PO No.',
    dateStr: dateString(formData.date), dateLabel: 'PO Date',
    extraDateLabel: formData.deliveryDate ? 'Expected Delivery' : '',
    extraDateStr: formData.deliveryDate ? dateString(formData.deliveryDate) : '',
    partyBlock: parties,
    cols: cols,
    numericColIdxs: [2, 3, 4, 5, 6],
    rows: rows,
    subtotalRow: ['SUBTOTAL', '', _fmtQty(sumQty), '', '', _money(sumTax), _money(sumAmount)],
    taxSummary: taxSummary,
    amountInWords: _amountInWords(Math.round(sumTotal)) + ' Only',
    customTerms: b.poTerms,   // 10-point PO T&Cs (from Config or hardcoded default)
    footerNote: formData.remarks ? 'Remarks: ' + formData.remarks : '',
  }, poNo, 'PO');
}

function generateReqSlipPdf(reqNo, formData, lineRows) {
  const cols = ['ITEMS','ITEM CODE','QTY','PURPOSE'];
  let sumQty = 0;
  const rows = lineRows.map(function(l) {
    sumQty += Number(l.qty) || 0;
    return [
      _htmlEscape(l.itemDesc || l.itemCode),
      l.itemCode,
      _fmtQty(l.qty) + ' ' + (l.uom || ''),
      l.purpose || '—',
    ];
  });
  return _renderAndSave({
    title: 'Material Requisition Slip',
    docNo: reqNo, docNoLabel: 'Req No.',
    dateStr: dateString(formData.date), dateLabel: 'Date',
    partyBlock: {
      label: 'Project',
      name: formData.project || '—',
      contact: 'Requested By: ' + currentUserEmail() + (formData.takenBy ? '  |  Taken By: ' + formData.takenBy : ''),
    },
    cols: cols,
    numericColIdxs: [2],
    rows: rows,
    subtotalRow: ['TOTAL', '', _fmtQty(sumQty), ''],
    showBankDetails: false,
  }, reqNo, 'Requisition');
}

function generateJobWorkChallanPdf(jwcNo, formData, lineRows) {
  const hasWeight = lineRows.some(function(l) { return Number(l.weight) > 0; });
  const cols = hasWeight
    ? ['ITEMS','CODE','QTY','WEIGHT (KG)','EXPECTED RETURN','REMARKS']
    : ['ITEMS','CODE','QTY','EXPECTED RETURN','REMARKS'];
  let sumQty = 0, sumWeight = 0;
  const rows = lineRows.map(function(l) {
    sumQty += Number(l.qty) || 0;
    sumWeight += Number(l.weight) || 0;
    const r = [
      _htmlEscape(l.desc || l.sfgCode),
      l.sfgCode,
      _fmtQty(l.qty) + ' ' + (l.uom || ''),
    ];
    if (hasWeight) r.push(Number(l.weight) ? Number(l.weight).toFixed(2) : '');
    r.push(l.expReturn ? dateString(l.expReturn) : '—', l.remarks || '');
    return r;
  });
  const subtotal = hasWeight
    ? ['TOTAL', '', _fmtQty(sumQty), _fmtQty(sumWeight), '', '']
    : ['TOTAL', '', _fmtQty(sumQty), '', ''];
  return _renderAndSave({
    title: 'Job Work Challan',
    docNo: jwcNo, docNoLabel: 'Challan No.',
    dateStr: dateString(formData.date), dateLabel: 'Date',
    partyBlock: {
      label: 'Job Worker',
      name: formData.jobWorkerName || formData.jobWorkerCode,
      gstin: formData.jobWorkerGstin || '',
      address: formData.jobWorkerAddress || '',
      contact: (formData.processType ? 'Process: ' + formData.processType : ''),
    },
    cols: cols,
    numericColIdxs: hasWeight ? [2, 3] : [2],
    rows: rows,
    subtotalRow: subtotal,
    showBankDetails: false,
    footerNote: 'Issued under Rule 45 of CGST Rules. Goods to be returned within 1 year of dispatch.',
  }, jwcNo, 'JobWorkChallan');
}

function generateJobWorkGrnPdf(jwgrnNo, formData, lineRows) {
  const b = _brand();
  const cols = ['#','Code','Description','Qty Recd','UOM','Yield Loss','Rejection','Remarks'];
  const rows = lineRows.map(function(l, i) {
    return [ i+1, l.sfgCode, l.desc, l.qtyRecv, l.uom,
             l.yieldLoss || 0, l.rejection || 0, l.remarks || '' ];
  });
  return _renderAndSave({
    title: 'Job Work GRN',
    docNo: jwgrnNo,
    dateStr: dateString(formData.date),
    partyBlock: {
      label: 'Received From (Job Worker)',
      name: formData.jobWorkerName || '—',
      code: '',
      contact: 'Reference Challan: ' + formData.referenceChallan,
    },
    cols: cols, rows: rows,
    signatures: [['Received By ' + b.company, currentUserEmail()], ['Job Worker (Sent)', '(Signature)']],
  }, jwgrnNo, 'JobWorkGRN');
}

// ============================================================================
// 2 report PDFs
// ============================================================================

function generateStockReportPdf(opts) {
  opts = opts || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snap = ss.getSheetByName('Stock Snapshot');
  if (!snap) throw new Error('Stock Snapshot missing. Run Initialize Workbook.');
  fullRebuild();

  const data = _readRmSection(snap, 17);
  if (!data.length) throw new Error('No stock data.');

  // v3.6 idx: 3=Storage/Rack, 8=Available, 9=Reserved, 12=Value, 15=Days, 16=BelowMIN
  // ponytail: swapped MIN/MAX for Storage/Rack — users consult the report at the
  // bin, not at a reorder review. Belo-MIN flag stays as the "< MIN" col.
  const cols = ['Item Code','Description','Category','Storage/Rack','Available','Reserved','Value (Rs.)','Days Idle','< MIN'];
  let rows = data.filter(function(r) { return r[0]; }).map(function(r) {
    return [ r[0], r[1], r[2], r[3] || '', r[8], r[9], _money(r[12]), r[15], (r[16] ? 'YES' : '') ];
  });
  if (opts.lowStockOnly)  rows = rows.filter(function(r) { return r[8] === 'YES'; });
  if (opts.category)      rows = rows.filter(function(r) { return r[2] === opts.category; });

  const grandValue = rows.reduce(function(s, r) { return s + Number(String(r[6]).replace(/,/g, '')); }, 0);  // col 6 = Value now
  const docNo = 'STK-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');

  return _renderAndSave({
    title: 'Stock Report' + (opts.lowStockOnly ? ' (Low Stock Only)' : '') + (opts.category ? ' — ' + opts.category : ''),
    docNo: docNo,
    dateStr: dateString(new Date()),
    partyBlock: {
      label: 'Report Details',
      name: (opts.lowStockOnly ? 'Below MIN LVL' : (opts.category || 'All Items')),
      contact: rows.length + ' items | Generated by ' + currentUserEmail(),
    },
    cols: cols, rows: rows,
    totalsRow: ['', '', '', '', '', '', 'Rs. ' + _money(grandValue), '', ''],
    signatures: [['Prepared By', currentUserEmail()], ['Approved By', '(Signature)']],
  }, docNo, 'StockReport');
}

function generateNonMovingReportPdf(days) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snap = ss.getSheetByName('Stock Snapshot');
  if (!snap) throw new Error('Stock Snapshot missing.');
  fullRebuild();
  days = Number(days) || Number(getConfig('Non-Moving Days Threshold')) || 115;

  const data = _readRmSection(snap, 17);
  if (!data.length) throw new Error('No stock data.');

  // v3.6 idx: 8=Available, 12=Value, 13=LastIn, 14=LastOut, 15=Days
  const cols = ['Item Code','Description','Available','Last Movement','Days Idle','Stock Value (Rs.)'];
  const rows = data
    .filter(function(r) { return r[0] && r[15] !== '' && r[15] >= days; })
    .sort(function(a, b) { return b[15] - a[15]; })
    .map(function(r) {
      const lastAny = r[14] || r[13] || '';
      return [ r[0], r[1], r[8], lastAny ? dateString(lastAny) : 'Never', r[15], _money(r[12]) ];
    });

  const totalIdleValue = rows.reduce(function(s, r) { return s + Number(String(r[5]).replace(/,/g, '')); }, 0);
  const docNo = 'NMS-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');

  return _renderAndSave({
    title: 'Non-Moving Stock Report',
    docNo: docNo,
    dateStr: dateString(new Date()),
    partyBlock: {
      label: 'Report Details',
      name: 'Threshold: ' + days + ' days',
      contact: rows.length + ' items flagged | Generated by ' + currentUserEmail(),
    },
    cols: cols, rows: rows,
    totalsRow: ['', '', '', '', 'Total Idle Value', 'Rs. ' + _money(totalIdleValue)],
    signatures: [['Prepared By', currentUserEmail()], ['Approved By', '(Signature)']],
    footerNote: 'Consider clearance / disposal / reallocation for items idle beyond threshold.',
  }, docNo, 'NonMoving');
}

// ============================================================================
// Render + save wrapper
// ============================================================================

function _renderAndSave(docSpec, docNo, docType) {
  const html = _buildDocHtml(docSpec);
  return savePdfToDrive(html, docNo, docType);
}

// Clear cached logo (call from editor if you swap the Company Logo URL and want
// the PDF to pick up the new image immediately).


// ============================================================================
// FG DELIVERY PDF — matches TSP DC template layout.
// Company block (left) + Doc info (right) → Bill To | Ship To → Items table
// → Hardware table (optional) → Transport block (optional) → Declaration +
// Signature footer. "ORIGINAL FOR CONSIGNEE" tag top right.
// ============================================================================

// RFQ — no vendor block, photo + spec table, quote-request framing.
function generateRfqPdf(rfqNo, formData, lineRows) {
  const cols = ['#','PHOTO','ITEM CODE','DESCRIPTION','SIZE','COLOUR','UOM','QTY REQUIRED','REMARKS'];
  let sumQty = 0;
  const rows = lineRows.map(function(l, i) {
    sumQty += Number(l.qty) || 0;
    return [
      i + 1,
      l.photoUrl ? _driveImageDataUrl(l.photoUrl) : '',
      l.itemCode,
      _htmlEscape(l.desc || ''),
      l.size || '—',
      l.colour || '—',
      l.uom || '',
      _fmtQty(l.qty),
      l.remarks || '',
    ];
  });

  const preTable = '<div style="padding:10px 14px;background:#FFF8ED;border-left:3px solid #F47C2C;margin-bottom:10px;font-size:10pt;line-height:1.5;">' +
    '<b>To all interested vendors —</b> Please quote your best price and delivery timeline for the items below. ' +
    'Include: unit rate, GST %, discount if any, freight terms, and lead time. ' +
    'Send your quote to the contact details in the header.' +
    (formData.validUntil ? '<br><b>Quote validity required until:</b> ' + _htmlEscape(dateString(formData.validUntil)) : '') +
    '</div>';

  return _renderAndSave({
    title: 'Request for Quotation',
    docNo: rfqNo, docNoLabel: 'RFQ No.',
    dateStr: dateString(formData.date), dateLabel: 'RFQ Date',
    // No partyBlock — this is a broadcast doc. Company info in the header is enough.
    preTableBlock: preTable,
    cols: cols,
    numericColIdxs: [0, 7],
    rows: rows,
    subtotalRow: ['', '', '', 'TOTAL', '', '', '', _fmtQty(sumQty), ''],
    showBankDetails: false,
    footerNote: [formData.purpose && 'Purpose: ' + formData.purpose,
                 formData.remarks && 'Notes: ' + formData.remarks].filter(Boolean).join('  |  '),
  }, rfqNo, 'RFQ');
}

function generateFgDeliveryPdf(dcNo, formData, allLines) {
  const mainLines = (allLines || []).filter(function(l) { return l.section === 'Main'; });
  const hwLines   = (allLines || []).filter(function(l) { return l.section === 'Hardware'; });

  const totalMainQty = mainLines.reduce(function(s, l) { return s + (Number(l.qty) || 0); }, 0);
  const totalMainWt  = mainLines.reduce(function(s, l) { return s + (Number(l.weight) || 0); }, 0);
  const totalHwQty   = hwLines.reduce(function(s, l) { return s + (Number(l.qty) || 0); }, 0);

  const mainRows = mainLines.map(function(l) {
    return [
      _htmlEscape(l.desc || ''),
      _fmtQty(l.qty),
      Number(l.weight) ? Number(l.weight).toFixed(2) : '',
      l.remark || '',
    ];
  });
  const sections = [{
    title: 'Items to be Delivered',
    cols: ['ITEMS','QTY','WEIGHT (KG)','REMARK'],
    numericColIdxs: [1, 2],
    rows: mainRows,
    subtotalRow: ['TOTAL', _fmtQty(totalMainQty), (totalMainWt ? totalMainWt.toFixed(2) : ''), ''],
  }];
  if (hwLines.length) {
    const hwRows = hwLines.map(function(l) {
      return [ _htmlEscape(l.desc || ''), _fmtQty(l.qty) ];
    });
    sections.push({
      title: 'Hardware List',
      cols: ['ITEMS','QTY'],
      numericColIdxs: [1],
      rows: hwRows,
      subtotalRow: ['TOTAL', _fmtQty(totalHwQty)],
    });
  }

  const transportBlock = (formData.vehicleNo || formData.driver || formData.lrNo || formData.eWayBillNo) ? (
    '<div class="transport-block"><b>Transport Details:</b>' +
    (formData.vehicleNo ? ' &nbsp; Vehicle: <b>' + _htmlEscape(formData.vehicleNo) + '</b>' : '') +
    (formData.driver    ? ' &nbsp; Driver: ' + _htmlEscape(formData.driver) : '') +
    (formData.lrNo      ? ' &nbsp; LR No: ' + _htmlEscape(formData.lrNo) : '') +
    (formData.eWayBillNo? ' &nbsp; E-Way Bill No: <b>' + _htmlEscape(formData.eWayBillNo) + '</b>' : '') +
    '</div>'
  ) : '';

  const parties = [
    { label: 'BILL TO', name: formData.customerName || '',
      address: formData.billTo || '',
      gstin: formData.billToGstin || '' },
    { label: 'SHIP TO', name: formData.customerName || '',
      address: formData.shipTo || formData.billTo || '',
      gstin: formData.shipToGstin || formData.billToGstin || '' },
  ];

  return _renderAndSave({
    title: 'Delivery Challan',
    docNo: dcNo, docNoLabel: 'Delivery Challan No.',
    dateStr: dateString(formData.date), dateLabel: 'Dated',
    extraDateLabel: formData.paymentTerms ? 'Payment Terms' : '',
    extraDateStr: formData.paymentTerms || '',
    partyBlock: parties,
    sections: sections,
    postTableBlock: transportBlock,
    showBankDetails: false,
    footerNote: formData.otherRef ? 'Other Reference: ' + formData.otherRef : '',
  }, dcNo, 'FGDelivery');
}

function clearLogoCache() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  toast('Logo cache cleared. Next PDF will re-fetch.');
}

// Bump the cache epoch — every existing image cache key becomes unreachable so
// the next PDF re-fetches all photos fresh. Use after fixing a wrong photo in
// Item Master or when a PDF shows a stale/incorrect image.
function clearPhotoCache() {
  PropertiesService.getScriptProperties().setProperty('_photoCacheEpoch', String(Date.now()));
  toast('Photo cache cleared. Next PDF will re-fetch all images.');
}