/**
 * 3 Star Projects IMS v3 — 00_Config_Menu_Setup.gs
 * Phase 0: Config sheet, custom menu, one-time workbook initialization.
 *
 * Deps: 60_Utils_Audit_Triggers.gs (toast)
 */

// ============================================================================
// SHEET SPECIFICATIONS — single source of truth for initializeWorkbook()
// ============================================================================
// Each entry: name, headers (row 1), optional freezeRows, optional widths.
// Adding a sheet? Add it here. initializeWorkbook() creates missing ones on
// re-run without disturbing existing data.

const SHEET_SPECS = [
  // -- Masters --
  // ponytail: 'Sqm per Unit' col removed Oct-2026 — per-row sqm value entered
  // on GRN/PO form instead; backend converts user's per-sqm rate to per-sheet.
  { name: 'Item Master', headers: ['Item Code','Photo','Description','Item Name','Size','Grade/Color','UOM','Category','Material Type','Storage No.','Rack No.','HSN/SAC Code','Preferred Supplier Code','Avg Daily Consumption','Lead Time (Days)','Safety Factor','MIN LVL','MAX LVL','Active?','Search Label','Avg Rate (Rs.)'] },
  { name: 'SFG Master', headers: ['SFG Code','Photo','Description','Item Name','Size','Color/Finish','UOM','Storage No.','MIN LVL','MAX LVL','Active?','Opening Qty','Available','Search Label'] },
  { name: 'FG Master', headers: ['FG Code','Photo','Description','Item Name','Size','Colour','Finish','UOM','Active?','Opening Qty','Available','Search Label'] },
  { name: 'Supplier Master', headers: ['Supplier Code','Firm Name','Category','Contact Person','Mobile No','Email','GST No','Address','Active?','Search Label'] },
  { name: 'Job Worker Master', headers: ['Job Worker Code','Firm Name','Process Type','Contact Person','Mobile No','Email','GST No','Address','Active?','Search Label'] },
  { name: 'Tool Master', headers: ['Tool Code','Photo','Description','Item Name','Type','Serial No','Purchase Date','Purchase Value','Vendor','Current Status','Currently With','Opening Qty','On Hand','Active?','Search Label'] },
  { name: 'Tool Movement Register', headers: ['Sr No','Movement No','Date','Tool Code','Tool Description','Action','Person / Site','Purpose','Expected Return','Actual Return','Condition','Remarks','Ref Movement No','Created By','Created At'] },
  { name: 'BOM Master', headers: ['BOM ID','Makes Type','Makes','Needs Type','Needs','Qty per Unit','UOM','Notes'] },
  { name: 'Dropdowns', headers: ['Categories','','Material Types','','UOM','','Process Types','','Employees','','Projects'] },

  // -- Registers --
  { name: 'Inward Register', headers: ['Sr No','GRN No','Date','Item Code','Item Description','Supplier Code','Supplier Name','Qty','UOM','Rate/Unit','Discount %','Discount Amount','GST %','GST Amount','Total Amount','Invoice No','Remarks','GRN PDF Link','Created By','Created At','PO Ref','Taxable Amount'] },
  { name: 'Material Requisition Register', headers: ['Sr No','Requisition No','Date','Project','Item Code','Item Description','Qty','UOM','Purpose / Work Description','Requested By','Taken By','Requisition Slip PDF Link','Remarks','Created By','Created At'] },
  { name: 'Job Work Outward Register', headers: ['Sr No','Job Work Challan No','Date','Job Worker Code','Job Worker Name','Process Type','SFG Code','SFG Description','Qty Sent','UOM','Expected Return Date','Remarks','Job Work Challan PDF Link','Status','Created By','Created At','Weight (KG)'] },
  { name: 'Job Work Inward Register', headers: ['Sr No','Job Work GRN No','Date','Reference Challan No','Job Worker Name','SFG Code','SFG Description','Qty Received','UOM','Yield Loss Qty','Rejection Qty','Remarks','Job Work GRN PDF Link','Created By','Created At'] },
  { name: 'FG Delivery Register', headers: ['Sr No','DC No','Date','Customer Name','Project','Payment Terms','Other Ref','Bill To','Bill To GSTIN','Ship To','Ship To GSTIN','Vehicle No','Driver','LR No','E-Way Bill No','Total Qty','Total Weight (KG)','DC PDF Link','Created By','Created At','Type'] },
  { name: 'FG Delivery Lines', headers: ['DC No','Line No','Section','Item Code','Description','Qty','UOM','Weight (KG)','Remark','Returned Qty'] },
  { name: 'Adjustments Register', headers: ['Sr No','Date','Type','Item/SFG Code','Description','Qty Adjustment','UOM','Reason / Details','Remarks','Created By','Created At'] },
  // Reservation Register — soft-locks stock for a specific customer/project/order.
  // Status: Active | Released (freed back) | Consumed (fulfilled by outward).
  // Manual release only — user updates Status via Release Reservation form.
  { name: 'Reservation Register', headers: ['Sr No','Reservation No','Date','Item/SFG Code','Description','Qty Reserved','UOM','For (Reference)','Status','Notes','Created By','Created At','Released/Consumed At','Released By'] },
  // QC Register — 35 cols. Source/Ref/Line/Outcome at front for easy scanning;
  // then Basic Details, Quantity Check, Visual Inspection, Specification Check,
  // Functional/Performance, Document, Created At.
  { name: 'QC Register', headers: [
    'Source Type','Reference No','Reference Line','Overall Outcome',
    'Vendor Name','Material Name','PO No','Delivery Date','Checked By (Basic)',
    'Ordered Qty','Received Qty','Short/Excess','Status (Qty)','Checked By (Qty)',
    'Any Damage','Packaging Condition','Leakage/Breakage','Proper Labeling','Status (Visual)','Checked By (Visual)',
    'Size/Dim Planned','Size/Dim Actual','Weight Planned','Weight Actual',
    'Colour/Finish Planned','Colour/Finish Actual','Material Grade Planned','Material Grade Actual',
    'Status (Spec)','Checked By (Spec)',
    'Does It Work as Expected','Trial/Sample Test Done','Output Quality Acceptable','Document Attached',
    'Created At',
  ]},
  { name: 'Purchase Order Register', headers: ['Sr No','PO No','Date','Supplier Code','Supplier Name','Delivery Date','Status','Total Value','PO PDF Link','Remarks','Created By','Created At','Rev No','Last Amended At','Last Amended By'] },
  { name: 'PO Lines', headers: ['PO No','Line No','Item Code','Item Description','Qty','UOM','Rate/Unit','Discount %','Discount Amount','GST %','GST Amount','Line Total','Received Qty'] },
  // RFQ (Request for Quotation) — PDF only, no register. See generateRfqPdf.

  // -- Snapshots / Viewers (consolidated) --
  //   Dashboard, Explorer (6 views), Stock Snapshot, Kitting Sheet.
  //   Deprecated separate viewers (Item Ledger, Supplier Ledger, Make & Grade
  //   Wise Stock, Non-Moving Stock, Low Stock Alerts, Pending Job Work) are
  //   no longer created. If you have them from a previous version, delete
  //   the tabs manually — nothing writes to them anymore.
  { name: 'Dashboard', headers: [] },       // Layout built by rebuildDashboard()
  { name: 'Explorer',  headers: [] },       // Layout built by rebuildExplorer(view, arg)
  // Stock Snapshot: RM section (cols A:Q, 17 cols) fills rows 2..N. SFG section is
  // written BELOW the RM section by fullRebuild (spacer row, then SFG header,
  // then SFG data). SFG headers are the SFG_SNAPSHOT_HEADERS const below —
  // NOT stored here in SHEET_SPECS since they're dynamic-row.
  //
  // Col I: Available (= physical on-hand − Reserved). Col J: Reserved (sum of
  // active reservations). Below MIN is now at col Q (was P).
  { name: 'Stock Snapshot', headers: ['Item Code','Description','Category','Storage/Rack','Opening Bal','Inward Qty','Outward Qty','Adjustment Qty','Available','Reserved','MIN LVL','MAX LVL','Stock Value (Rs.)','Last Inward Date','Last Outward Date','Days Since Movement','Below MIN?'] },
  { name: 'Kitting Sheet', headers: [] },

  // -- System --
  { name: 'Config', headers: ['Setting','Value','Notes'] },
];

// Default Config key/value/notes rows. Written only when Config is fresh.
const CONFIG_DEFAULTS = [
  ['Company Name', '3 Star Projects', ''],
  ['Company Email', 'info@3starprojects.com', 'For doc signature blocks'],
  ['Company Address', '', 'Fill this in — appears on all PDFs'],
  ['Company Phone', '', 'Optional, for doc footer'],
  ['Company Logo URL', '', 'Drive image link for PDF headers'],
  ['Brand Primary Color', '#1F4E79', 'Blue'],
  ['Brand Secondary Color', '#F47C2C', 'Orange'],
  ['PDF Docs Folder ID', '', 'Root Drive folder where PDFs are saved'],
  ['Next GRN Counter', 0, ''],
  ['Next Requisition Counter', 0, ''],
  ['Next Delivery Challan Counter', 0, ''],
  ['Next Job Work Challan Counter', 0, ''],
  ['Next Job Work GRN Counter', 0, ''],
  ['Next PO Counter', 0, ''],
  ['Next RFQ Counter', 0, 'For RFQ-#### (Request for Quotation)'],
  ['Next Movement Counter', 0, 'For TM-#### (tool movements)'],
  ['Next Reservation Counter', 0, 'For RSV-#### (stock reservations)'],
  ['Project List', '', 'Comma-separated: "Silvassa Site, Mumbai Office, ..."'],
  ['Low Stock Alert Recipients', '', 'Comma-separated emails'],
  ['Admin Email', 'info@3starprojects.com', 'Gets alerts when nightly rebuild fails'],
  ['High Value Threshold', 100000, 'GRN/PO total above this triggers confirmation dialog (Rs.)'],
  ['Nightly Rebuild Enabled', true, ''],
  ['Nightly Rebuild Hour (IST)', 2, '0–23'],
  ['Non-Moving Days Threshold', 115, ''],
  // -- GST + banking (added for tax-invoice-format PDFs) --
  ['Company State', 'Maharashtra', 'State name matching GSTIN prefix. Used to decide CGST+SGST (intra) vs IGST (inter).'],
  ['Company GSTIN', '', '15-char GSTIN — shows on every PDF'],
  ['Company PAN', '', 'Optional, shows below GSTIN'],
  ['Bank Name', '', 'Shows in Bank Details footer'],
  ['Bank Account No', '', ''],
  ['Bank IFSC', '', ''],
  ['Bank Branch', '', ''],
  ['Terms and Conditions', 'Goods once accepted will not be taken back. Interest @18% p.a. on delayed payments. Subject to Thane jurisdiction.', 'Shows in footer of GRN and Delivery Challan. Use \\n for line breaks.'],
  ['DC No Prefix', 'TSP', 'Prefix for FG Delivery Challan numbers (e.g., TSP/2026-27/031)'],
  ['DC Declaration Text', 'We declare that this delivery challan shows the actual price of the goods described and that all particulars are true and correct.', 'Bottom-left declaration on every FG Delivery PDF'],
];

// Dropdowns sheet contents. Cols B, D, F left blank as visual separators.
// Dropdowns sheet is CODE-MANAGED — _seedDropdowns overwrites cols A/C/E/G
// on every Initialize Workbook. User-edits to those cols will be wiped.
const DROPDOWN_LISTS = {
  A: [
    'Board Materials','Surface Materials','Edge Banding','Fabric & Upholstery','Foam & Cushioning',
    'Glass & Mirror','Aluminium Profiles','Steel & Metal Sections','Stainless Steel Materials',
    'PVC & Plastic Materials','Acrylic Materials','Decorative Panels','Wall & Ceiling Materials',
    'Stone & Solid Surfaces','Furniture Hardware','Fasteners & Connectors','Sliding Systems',
    'Drawer Systems','Locks & Security Hardware','Wardrobe Accessories','Kitchen Accessories',
    'Glass Hardware','Adhesives & Chemicals','Sealants & Silicone','Paints & Coatings',
    'Finishing Materials','Polishing Materials','Electrical Materials','Lighting Materials',
    'Wiring & Switchgear','Sensors & Automation','Packaging Materials','Protective Packaging',
    'Storage & Handling Materials','Fabrication Consumables','Welding Materials','Machine Consumables',
    'Cutting Tools & Abrasives','Safety Materials','Housekeeping Materials','Maintenance Materials',
    'Spare Parts','Tools & Equipment','Power Tools','Hand Tools','Machinery & Equipment',
    'Printing & Branding Materials','Office Supplies','IT & Automation Materials','Plumbing Materials',
    'HVAC Materials','Acoustic Materials','Interior Decor Materials','Modular Furniture Components',
    'Office Chair Components','School Furniture Components','Institutional Furniture Components',
    'Project & Site Materials','Installation Materials','Construction Consumables',
    'Imported Specialty Materials','Eco-Friendly Materials','Scrap & Recyclable Materials',
    'Finished Furniture','Semi Finished Components','Work In Progress Materials','Trading Materials',
    'Other',
  ],
  C: [
    'Raw Material','Consumables','Semi Finished Goods','Finished Goods','Trading Goods',
    'Packing Material','Work In Progress','Assets','Tools & Equipment','Maintenance Items',
    'Safety Items','Electrical Consumables','Fabrication Consumables','Office Consumables',
    'Housekeeping Consumables','Spare Parts','Returnable Items','Scrap Material',
    'Installation Material','Project Material','Imported Material','Custom Fabricated Items',
    'OEM Components','Sample Materials','Repair & Service Items','Rental Items',
    'Reusable Materials','Non-Stock Items','Capital Goods','Production Supplies',
  ],
  E: [
    'No.s','Pieces','Units','Sets','Pairs','Boxes','Packets','Bundles','Rolls','Sheets',
    'Boards','Panels','Plates','Bars','Rods','Pipes','Tubes','Coils','Reels','Meters',
    'Running Feet','Square Feet','Square Meters','Cubic Feet','Cubic Meters','Kilograms',
    'Grams','Metric Tons','Liters','Milliliters','Bags','Cartons','Drums','Buckets','Cans',
    'Bottles','Pouches','Trays','Pallets','Containers','Loads','Lots','Hours','Days',
    'Services','Jobs','Projects','Kits','Assemblies','Length','Weight','Volume','Area',
  ],
  G: ['Powder Coating','Lamination','Polishing','Upholstery','Plating','Engraving','Other'],
};

// ============================================================================
// MENU — built every open
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // ── 1) Add & Edit Data ──────────────────────────────────────────────
  // ponytail: Masters is now ONE submenu — each entry opens a unified form
  // with a New/Edit mode toggle (Tool stays single — no edit form).
  ui.createMenu('📝 3SP Data')
    .addSubMenu(ui.createMenu('Masters')
      .addItem('Item',                   'openMasterItemForm')
      .addItem('SFG',                    'openMasterSfgForm')
      .addItem('FG',                     'openMasterFgForm')
      .addItem('Supplier',               'openMasterSupplierForm')
      .addItem('Job Worker',             'openMasterJobWorkerForm')
      .addItem('Tool',                   'openAddToolForm'))
    .addSubMenu(ui.createMenu('Add Transaction')
      .addItem('Inward (GRN)',           'openAddInwardForm')
      .addItem('Material Requisition',   'openAddRequisitionForm')
      .addItem('Delivery (Internal / Customer)', 'openAddFgDeliveryForm')
      .addItem('Return from Site',       'openSiteReturnForm')
      .addSeparator()
      .addItem('Job Work Outward',       'openAddJobWorkOutForm')
      .addItem('Job Work Inward',        'openAddJobWorkInForm')
      .addSeparator()
      .addItem('Purchase Order',         'openAddPoForm')
      .addItem('Amend Purchase Order',   'openAmendPoForm')
      .addItem('Request for Quotation',  'openAddRfqForm')
      .addItem('Adjustment',             'openAdjustmentForm')
      .addSeparator()
      .addItem('QC Record',              'openAddQcRecordForm')
      .addSeparator()
      .addItem('Reservation',            'openReservationForm')
      .addSeparator()
      .addItem('Tool Movement',          'openToolMovementForm'))
    .addSubMenu(ui.createMenu('Recipes (BOM)')
      .addItem('Recipe (New or Edit)',   'openRecipeForm')
      .addItem('Kit Builder',            'openKitBuilderForm'))
    // (Edit Master submenu removed — merged into Masters above)
    .addSubMenu(ui.createMenu('Bulk')
      .addItem('Bulk Edit',              'openBulkEditForm'))
    .addToUi();

  // ── 2) Setup ─────────────────────────────────────────────────────────
  ui.createMenu('⚙️ 3SP Setup')
    .addItem('Initialize Workbook',       'initializeWorkbook')
    .addItem('Install Triggers',          'setupTriggers')
    .addSeparator()
    .addItem('Rebuild Formulas',          'rebuildAllFormulas')
    .addItem('Apply Visual Formatting',   'applyVisualFormatting')
    .addItem('Diagnose Dropdowns',        'diagnoseDropdowns')
    .addSeparator()
    .addItem('Full Stock Rebuild',        'fullRebuild')
    .addSeparator()
    .addItem('Regenerate Missing PDFs',   'regenerateMissingPdfs')
    .addItem('Clear Photo Cache',         'clearPhotoCache')
    .addItem('Regenerate PDF (by Doc No)','regeneratePdfForDoc')
    .addSeparator()
    .addItem('Create Draft PO from Reorder', 'createDraftPoFromReorder')
    .addSeparator()
    .addItem('Undo Last Transaction',     'undoLastTransaction')
    .addItem('Delete Single Register Row','deleteRegisterRow')
    .addSeparator()

    .addSeparator()    .addToUi();

  // ── 3) Report Generation ────────────────────────────────────────────
  ui.createMenu('📊 3SP Reports')
    .addItem('Stock Report (PDF)',        'menuStockReport')
    .addItem('Non-Moving Report (PDF)',   'menuNonMovingReport')
    .addSeparator()
    .addItem('Export Explorer as CSV',    'exportExplorerAsCsv')
    .addToUi();

  // ── 4) Dashboard direct refresh (single-click) ──────────────────────
  ui.createMenu('🔄 3SP Dashboard')
    .addItem('Refresh Dashboard now',     'menuRefreshDashboard')
    .addItem('Refresh Explorer (current view)', 'menuRefreshExplorer')
    .addToUi();
}

// ============================================================================
// FORM OPENERS — every menu item that opens 05_Forms.html
// ============================================================================

function _openForm(formId, title, width, height) {
  const html = HtmlService.createTemplateFromFile('05_Forms');
  html.formId = formId;
  const out = html.evaluate().setWidth(width || 900).setHeight(height || 650).setTitle(title);
  SpreadsheetApp.getUi().showModalDialog(out, title);
}

function openMasterItemForm()       { _openForm('form-master-item',       'Item Master (New or Edit)',       820, 720); }
function openMasterSfgForm()        { _openForm('form-master-sfg',        'SFG Master (New or Edit)',        820, 640); }
function openMasterSupplierForm()   { _openForm('form-master-supplier',   'Supplier Master (New or Edit)',   820, 680); }
function openMasterJobWorkerForm()  { _openForm('form-master-jobworker',  'Job Worker Master (New or Edit)', 820, 680); }
function openAddInwardForm()      { _openForm('form-add-inward',     'Add Inward (GRN)',        1250, 720); }
function openAddRequisitionForm() { _openForm('form-add-requisition','Add Material Requisition',1000, 680); }
function openAddJobWorkOutForm()  { _openForm('form-add-jw-out',     'Add Job Work Outward',    1000, 680); }
function openAddJobWorkInForm()   { _openForm('form-add-jw-in',      'Add Job Work Inward',     1000, 680); }
function openAdjustmentForm()        { _openForm('form-adjustment', 'Adjustment (single or bulk)', 1000, 680); }
function openAddQcRecordForm()       { _openForm('form-add-qc-record',   'QC Record',           1050, 720); }
function openReservationForm()       { _openForm('form-reservation', 'Reservation (New or Release)', 900, 580); }
function openAddPoForm()          { _openForm('form-add-po',         'Add Purchase Order (PO)', 1250, 720); }
function openAmendPoForm()        { _openForm('form-amend-po',       'Amend Purchase Order',    1250, 720); }
function openAddRfqForm()         { _openForm('form-add-rfq',        'Request for Quotation',   1000, 680); }
function openMasterFgForm()         { _openForm('form-master-fg',         'FG Master (New or Edit)',         820, 640); }
function openAddToolForm()          { _openForm('form-add-tool',          'Add Tool',                        720, 620); }
function openToolMovementForm()   { _openForm('form-tool-movement', 'Tool Movement (Issue or Return)', 760, 560); }
function openRecipeForm()         { _openForm('form-recipe',         'Recipe (BOM — New or Edit)', 1000, 680); }
function openKitBuilderForm()     { _openForm('form-kit-builder',    'Kit Builder',             900, 620); }
// Edit-only master openers removed — folded into openMaster*Form above (New/Edit toggle)
function openBulkEditForm()       { _openForm('form-bulk-edit',     'Bulk Edit',               900, 640); }
function openAddFgDeliveryForm() { _openForm('form-add-fg-delivery', 'Add FG Delivery (DC)',  1050, 700); }
function openSiteReturnForm()    { _openForm('form-site-return',     'Return from Site',       1050, 680); }

// Menu handlers — phased implementation.
function menuRefreshDashboard()      { fullRebuild(); rebuildDashboard(); toast('Dashboard refreshed.'); }
function menuRefreshExplorer() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Explorer');
  if (!sh) { toast('Explorer sheet missing — run Initialize Workbook.'); return; }
  const view = sh.getRange('B1').getValue();
  const arg  = sh.getRange('D1').getValue();
  fullRebuild(); rebuildExplorer(view, arg);
  toast('Explorer refreshed (' + view + ').');
}
function menuStockReport() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Stock Report',
    'Filter: enter category name for that category only, "low" for below-MIN only, or leave blank for all items.',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const filter = resp.getResponseText().trim();
  const opts = {};
  if (filter.toLowerCase() === 'low') opts.lowStockOnly = true;
  else if (filter) opts.category = filter;
  const pdf = generateStockReportPdf(opts);
  ui.alert('Stock Report ready:\n\n' + pdf.url);
}

function menuNonMovingReport() {
  const ui = SpreadsheetApp.getUi();
  const def = String(getConfig('Non-Moving Days Threshold') || 115);
  const resp = ui.prompt('Non-Moving Report', 'Days threshold (default ' + def + '):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const days = Number(resp.getResponseText().trim()) || Number(def);
  const pdf = generateNonMovingReportPdf(days);
  ui.alert('Non-Moving Report ready (>= ' + days + ' days):\n\n' + pdf.url);
}
// setupTriggers() lives in 60_Utils_Audit_Triggers.gs (real installer)


// ============================================================================
// ONE-TIME SETUP — run once from menu after pasting code
// ============================================================================

function initializeWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let created = 0, skipped = 0;

  let extended = 0;
  SHEET_SPECS.forEach(function(spec) {
    let sh = ss.getSheetByName(spec.name);
    if (sh) {
      // Sheet exists — append any missing header columns (safe for existing data)
      if (spec.headers && spec.headers.length) {
        const currentHeaders = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
        const missing = spec.headers.slice(currentHeaders.length).filter(function(h) { return h; });
        if (missing.length) {
          sh.getRange(1, currentHeaders.length + 1, 1, missing.length).setValues([missing])
            .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
          extended++;
        }
      }
      skipped++;
      return;
    }
    sh = ss.insertSheet(spec.name);
    if (spec.headers && spec.headers.length) {
      sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers])
        .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
    }
    created++;
  });

  _seedConfig(ss);
  _seedDropdowns(ss);
  _migrateBomHeaders(ss);  // Rename Parent/Component → Makes/Needs (idempotent)
  _migrateRequisitionHeaders(ss);  // Rename Approved By → Taken By (idempotent)
  _migrateOnHandToAvailable(ss);   // Rename SFG/FG Master 'On Hand' → 'Available' (idempotent)
  // ponytail: order matters. Add Avg Rate first (so it exists at col V on old
  // sheets), then delete the Sqm per Unit col (shifting Avg Rate V→U). On a
  // fresh sheet both are no-ops. On already-migrated sheets the delete is a
  // one-shot irreversible drop of that column.
  _migrateItemMasterSqmToAvgRate(ss);  // Avg Rate added, Sqm per Unit removed
  _forceRewriteStockSnapshotHeaders(ss);  // Ensure RM header reflects Available/Reserved (v3.6)
  _applyMasterValidations(ss);
  _seedDescriptionFormulas(ss);
  _seedSearchLabels(ss);
  _seedOnHandFormulas(ss);
  _seedViewerPickers(ss);

  // Nuke the default Sheet1 if it's still empty and lingering
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);

  // Remove deprecated sheets (Audit Log + Errors + Site Material Send Register,
  // the last consolidated into FG Delivery Register with Type='Internal').
  // Only deletes if the sheet has 0 data rows (data-safety). If data exists,
  // user must manually back up + delete the tab.
  let removed = 0;
  ['Audit Log', 'Errors', 'Site Material Send Register'].forEach(function(name) {
    const s = ss.getSheetByName(name);
    if (!s) return;
    if (s.getLastRow() > 1) {
      SpreadsheetApp.getUi().alert('Skipped deleting "' + name + '" — it still has data rows. Back it up and delete the tab manually.');
      return;
    }
    ss.deleteSheet(s); removed++;
  });

  SpreadsheetApp.getUi().alert(
    'Workbook initialized.\n\nSheets created: ' + created +
    '\nSheets already present (skipped): ' + skipped +
    '\nExisting sheets with new columns added: ' + extended +
    (removed ? '\nDeprecated sheets removed: ' + removed + ' (Audit Log/Errors)' : '') +
    '\n\nNext: fill in Config values, add photos to Item/SFG/FG Master rows if desired.'
  );
}

function _seedConfig(ss) {
  const sh = ss.getSheetByName('Config');
  if (!sh) return;
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 380);

  if (sh.getLastRow() < 2) {
    // Fresh — seed everything
    sh.getRange(2, 1, CONFIG_DEFAULTS.length, 3).setValues(CONFIG_DEFAULTS);
    return;
  }
  // Already populated — only APPEND keys that don't exist yet. Preserves user values.
  const existing = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function(r) { return r[0]; });
  const missing = CONFIG_DEFAULTS.filter(function(row) { return existing.indexOf(row[0]) < 0; });
  if (missing.length) {
    sh.getRange(sh.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
  }
}

// One-shot rename of legacy BOM Master headers (Parent/Component → Makes/Needs).
// Called from initializeWorkbook. Safe to re-run — only writes if old name found.
function _migrateBomHeaders(ss) {
  const bom = ss.getSheetByName('BOM Master');
  if (!bom || bom.getLastColumn() === 0) return;
  const renames = {
    'Parent Type':    'Makes Type',
    'Parent Code':    'Makes',
    'Component Type': 'Needs Type',
    'Component Code': 'Needs',
  };
  const headers = bom.getRange(1, 1, 1, bom.getLastColumn()).getValues()[0];
  let changed = false;
  for (let i = 0; i < headers.length; i++) {
    if (renames[headers[i]]) {
      bom.getRange(1, i + 1).setValue(renames[headers[i]]);
      changed = true;
    }
  }
  if (changed) toast('BOM Master headers migrated (Parent/Component → Makes/Needs).');
}

// Force-rewrite Stock Snapshot header row 1 to match current SHEET_SPECS.
// Data rows (2+) get overwritten by fullRebuild anyway. Headers are the only
// user-visible artifact of schema changes — this ensures upgrades pick up new
// col names like Available/Reserved without user having to run Full Rebuild first.
// Also clears any stale header cells past the current schema width.
function _forceRewriteStockSnapshotHeaders(ss) {
  const sh = ss.getSheetByName('Stock Snapshot');
  if (!sh) return;
  const spec = SHEET_SPECS.filter(function(s) { return s.name === 'Stock Snapshot'; })[0];
  if (!spec || !spec.headers) return;
  const currentLastCol = Math.max(sh.getLastColumn(), spec.headers.length);
  // Write new headers
  sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers])
    .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  // Clear any stale header cells beyond new schema width (e.g., old side-by-side cols)
  if (currentLastCol > spec.headers.length) {
    sh.getRange(1, spec.headers.length + 1, 1, currentLastCol - spec.headers.length).clearContent().clearFormat();
  }
}

// Rename Material Requisition Register 'Approved By' → 'Taken By'. Idempotent.
function _migrateRequisitionHeaders(ss) {
  const sh = ss.getSheetByName('Material Requisition Register');
  if (!sh || sh.getLastColumn() === 0) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === 'Approved By') {
      sh.getRange(1, i + 1).setValue('Taken By');
      toast('Requisition Register: renamed "Approved By" → "Taken By".');
      return;
    }
  }
}

// Rename SFG Master col M and FG Master col K 'On Hand' → 'Available' (v3.7).
// Idempotent — only acts on headers still named 'On Hand'.
function _migrateOnHandToAvailable(ss) {
  [['SFG Master', 13], ['FG Master', 11]].forEach(function(pair) {
    const sh = ss.getSheetByName(pair[0]);
    if (!sh || sh.getLastColumn() < pair[1]) return;
    const cell = sh.getRange(1, pair[1]);
    if (cell.getValue() === 'On Hand') {
      cell.setValue('Available');
      toast(pair[0] + ': renamed "On Hand" → "Available".');
    }
  });
}

// Oct-2026: delete 'Sqm per Unit' column from Item Master. The sqm/unit value
// is now entered per-row on GRN/PO forms; backend converts user's per-sqm rate
// to per-sheet storage. Idempotent — looks up the column by header, deletes it
// only if found. Sheets auto-shifts subsequent columns left, so Avg Rate moves
// from col V → col U after this runs.
// Idempotent Item Master migration: ensure Avg Rate (Rs.) exists, then drop
// Sqm per Unit if still present. Order matters — add first, delete after, so
// Avg Rate ends up at col U (21) regardless of starting state.
function _migrateItemMasterSqmToAvgRate(ss) {
  const sh = ss.getSheetByName('Item Master');
  if (!sh) return;
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('Avg Rate (Rs.)') < 0) {
    const newCol = sh.getLastColumn() + 1;
    sh.getRange(1, newCol).setValue('Avg Rate (Rs.)')
      .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
    toast('Item Master: added "Avg Rate (Rs.)" column.');
    headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  }
  const sqmIdx = headers.indexOf('Sqm per Unit');
  if (sqmIdx >= 0) {
    sh.deleteColumn(sqmIdx + 1);
    toast('Item Master: removed "Sqm per Unit" — Avg Rate now at col U.');
  }
}

function _seedDropdowns(ss) {
  const sh = ss.getSheetByName('Dropdowns');
  if (!sh) return;
  // Row 1 headers — authoritative. Overwrites any duplicates the sheet may
  // have acquired from column inserts. Matches SHEET_SPECS['Dropdowns'].
  const DD_HEADERS = ['Categories','','Material Types','','UOM','','Process Types','','Employees','','Projects'];
  sh.getRange(1, 1, 1, DD_HEADERS.length).setValues([DD_HEADERS])
    .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  // Column widths for readability
  [[1,180],[2,20],[3,180],[4,20],[5,140],[6,20],[7,180],[8,20],[9,180],[10,20],[11,180]]
    .forEach(function(cw){ sh.setColumnWidth(cw[0], cw[1]); });

  Object.keys(DROPDOWN_LISTS).forEach(function(colLetter) {
    const values = DROPDOWN_LISTS[colLetter];
    const col = colLetter.charCodeAt(0) - 64; // A=1
    // Wipe entire col beyond header, then write fresh (authoritative — user
    // edits to these cols are overwritten on every Init).
    const lastRow = Math.max(sh.getLastRow(), values.length + 1);
    if (lastRow >= 2) sh.getRange(2, col, lastRow - 1, 1).clearContent();
    sh.getRange(2, col, values.length, 1).setValues(values.map(function(v) { return [v]; }));
  });
}

// Locate a column by its header text — future-proof against column inserts.
// Returns 1-based col index, or -1 if not found.
function _colByHeader(sh, headerText) {
  if (!sh || sh.getLastColumn() === 0) return -1;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  for (let i = 0; i < headers.length; i++) if (headers[i] === headerText) return i + 1;
  return -1;
}

// Wipe existing validations from data-entry area before re-applying — prevents
// stale validations from earlier column layouts landing on the wrong cell.
function _wipeMasterValidations(sh, cols) {
  if (!sh) return;
  const c = cols || Math.max(20, sh.getLastColumn());
  sh.getRange(2, 1, 500, c).clearDataValidations();
}

function _applyMasterValidations(ss) {
  const dropdowns = ss.getSheetByName('Dropdowns');
  if (!dropdowns) return;
  const bigRange = 500;

  // ---- Item Master: UOM, Category, Material Type ----
  const im = ss.getSheetByName('Item Master');
  if (im) {
    _wipeMasterValidations(im);
    const uomCol = _colByHeader(im, 'UOM');
    const catCol = _colByHeader(im, 'Category');
    const matCol = _colByHeader(im, 'Material Type');
    if (uomCol > 0) _dv(im, 2, uomCol, bigRange, 1, dropdowns.getRange('E2:E' + (1 + DROPDOWN_LISTS.E.length)));
    if (catCol > 0) _dv(im, 2, catCol, bigRange, 1, dropdowns.getRange('A2:A' + (1 + DROPDOWN_LISTS.A.length)));
    if (matCol > 0) _dv(im, 2, matCol, bigRange, 1, dropdowns.getRange('C2:C' + (1 + DROPDOWN_LISTS.C.length)));
  }
  // ---- SFG Master: UOM ----
  const sfg = ss.getSheetByName('SFG Master');
  if (sfg) {
    _wipeMasterValidations(sfg);
    const uomCol = _colByHeader(sfg, 'UOM');
    if (uomCol > 0) _dv(sfg, 2, uomCol, bigRange, 1, dropdowns.getRange('E2:E' + (1 + DROPDOWN_LISTS.E.length)));
  }
  // ---- FG Master: UOM ----
  const fg = ss.getSheetByName('FG Master');
  if (fg) {
    _wipeMasterValidations(fg);
    const uomCol = _colByHeader(fg, 'UOM');
    if (uomCol > 0) _dv(fg, 2, uomCol, bigRange, 1, dropdowns.getRange('E2:E' + (1 + DROPDOWN_LISTS.E.length)));
  }
  // ---- Job Worker Master: Process Type ----
  const jw = ss.getSheetByName('Job Worker Master');
  if (jw) {
    _wipeMasterValidations(jw);
    const ptCol = _colByHeader(jw, 'Process Type');
    if (ptCol > 0) _dv(jw, 2, ptCol, bigRange, 1, dropdowns.getRange('G2:G' + (1 + DROPDOWN_LISTS.G.length)));
  }
  // ---- BOM Master: Makes Type + Needs Type ----
  const bom = ss.getSheetByName('BOM Master');
  if (bom) {
    _wipeMasterValidations(bom);
    const ptCol = _colByHeader(bom, 'Makes Type');
    const ctCol = _colByHeader(bom, 'Needs Type');
    if (ptCol > 0) _dvList(bom, 2, ptCol, bigRange, 1, ['FG','SFG']);
    if (ctCol > 0) _dvList(bom, 2, ctCol, bigRange, 1, ['RM','SFG']);
  }
  // ---- Job Work Outward Register: Status ----
  const jwo = ss.getSheetByName('Job Work Outward Register');
  if (jwo) {
    _wipeMasterValidations(jwo);
    const statusCol = _colByHeader(jwo, 'Status');
    if (statusCol > 0) _dvList(jwo, 2, statusCol, bigRange, 1, ['Open','Closed','Cancelled']);
  }
  // ---- Adjustments Register: Type ----
  const adj = ss.getSheetByName('Adjustments Register');
  if (adj) {
    _wipeMasterValidations(adj);
    const typeCol = _colByHeader(adj, 'Type');
    if (typeCol > 0) _dvList(adj, 2, typeCol, bigRange, 1, ['Opening','Damage','Correction','Return','Other']);
  }
  // ---- Purchase Order Register: Status ----
  const po = ss.getSheetByName('Purchase Order Register');
  if (po) {
    _wipeMasterValidations(po);
    const statusCol = _colByHeader(po, 'Status');
    if (statusCol > 0) _dvList(po, 2, statusCol, bigRange, 1, ['Draft','Issued','Received','Cancelled']);
  }
  // ---- Tool Master: Type + Current Status ----
  const tm = ss.getSheetByName('Tool Master');
  if (tm) {
    _wipeMasterValidations(tm);
    const typeCol   = _colByHeader(tm, 'Type');
    const statusCol = _colByHeader(tm, 'Current Status');
    if (typeCol > 0)   _dvList(tm, 2, typeCol,   bigRange, 1, ['Power Tool','Hand Tool','Machine','Consumable','Other']);
    if (statusCol > 0) _dvList(tm, 2, statusCol, bigRange, 1, ['In Store','Issued','At Site','Under Repair','Lost','Retired']);
  }
  // ---- Tool Movement Register: Action + Condition ----
  const tmr = ss.getSheetByName('Tool Movement Register');
  if (tmr) {
    _wipeMasterValidations(tmr);
    const actionCol = _colByHeader(tmr, 'Action');
    const condCol   = _colByHeader(tmr, 'Condition');
    if (actionCol > 0) _dvList(tmr, 2, actionCol, bigRange, 1, ['Issue','Return','Transfer','Repair Send','Repair Return','Retire']);
    if (condCol > 0)   _dvList(tmr, 2, condCol,   bigRange, 1, ['Good','Damaged','Needs Repair','Lost']);
  }
  // ---- Reservation Register: Status dropdown ----
  const res = ss.getSheetByName('Reservation Register');
  if (res) {
    _wipeMasterValidations(res);
    const stCol = _colByHeader(res, 'Status');
    if (stCol > 0) _dvList(res, 2, stCol, bigRange, 1, ['Active','Released','Consumed']);
  }
  // ---- QC Register: dropdowns for Source Type, Outcome, and all Status/inspection cols ----
  const qc = ss.getSheetByName('QC Register');
  if (qc) {
    _wipeMasterValidations(qc);
    const srcCol = _colByHeader(qc, 'Source Type');
    if (srcCol > 0) _dvList(qc, 2, srcCol, bigRange, 1, ['GRN','JWGRN','Pre-Dispatch DC','Manual']);
    const outCol = _colByHeader(qc, 'Overall Outcome');
    if (outCol > 0) _dvList(qc, 2, outCol, bigRange, 1, ['Accept','Reject','Rework']);
    // Status cols (3 sections) — Pass/Fail/Hold
    ['Status (Qty)','Status (Visual)','Status (Spec)'].forEach(function(h) {
      const c = _colByHeader(qc, h);
      if (c > 0) _dvList(qc, 2, c, bigRange, 1, ['Pass','Fail','Hold']);
    });
    // Damage cols — No Damage / Minor Damage / Major Damage
    ['Any Damage','Packaging Condition'].forEach(function(h) {
      const c = _colByHeader(qc, h);
      if (c > 0) _dvList(qc, 2, c, bigRange, 1, ['No Damage','Minor Damage','Major Damage']);
    });
    // Yes/No/Not Applicable cols
    ['Leakage/Breakage','Proper Labeling','Does It Work as Expected','Trial/Sample Test Done','Output Quality Acceptable','Document Attached'].forEach(function(h) {
      const c = _colByHeader(qc, h);
      if (c > 0) _dvList(qc, 2, c, bigRange, 1, ['Yes','No','Not Applicable']);
    });
    // Colour/Finish Actual
    const cfCol = _colByHeader(qc, 'Colour/Finish Actual');
    if (cfCol > 0) _dvList(qc, 2, cfCol, bigRange, 1, ['Match Approved Sample','Slight Variation','Wrong Shade','Not Applicable']);
    // Material Grade Actual
    const mgCol = _colByHeader(qc, 'Material Grade Actual');
    if (mgCol > 0) _dvList(qc, 2, mgCol, bigRange, 1, ['Correct','Incorrect']);
  }
}

// Seed the picker labels + dropdowns on viewer sheets. Idempotent: only
// writes labels if the cell is empty, so re-running Initialize doesn't
// wipe a value the user selected.
function _seedViewerPickers(ss) {
  const fg  = ss.getSheetByName('FG Master');

  // Explorer: A1 label, B1 view picker, C1 dynamic label, D1 dynamic input
  const exp = ss.getSheetByName('Explorer');
  if (exp) {
    if (!exp.getRange('A1').getValue()) exp.getRange('A1').setValue('View:').setFontWeight('bold').setBackground('#1F4E79').setFontColor('#FFFFFF').setHorizontalAlignment('right');
    if (!exp.getRange('B1').getValue()) exp.getRange('B1').setValue('Item Ledger');
    const viewRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(EXPLORER_VIEWS, true).setAllowInvalid(false).build();
    exp.getRange('B1').setDataValidation(viewRule);
    exp.setColumnWidth(1, 90);
    exp.setColumnWidth(2, 200);
    exp.setColumnWidth(3, 130);
    exp.setColumnWidth(4, 260);
    exp.setRowHeight(1, 32);
    // Programmatic setValue does NOT trigger onEdit — do the initial rebuild here
    // so C1 label + D1 dropdown appear immediately.
    if (typeof rebuildExplorer === 'function') {
      rebuildExplorer(exp.getRange('B1').getValue(), exp.getRange('D1').getValue());
    }
  }

  // Kitting: B1 label, C1 FG picker, D1 qty label, E1 qty, F1/G1 total
  const kt = ss.getSheetByName('Kitting Sheet');
  if (kt) {
    if (!kt.getRange('B1').getValue()) kt.getRange('B1').setValue('Select FG:').setFontWeight('bold');
    if (!kt.getRange('D1').getValue()) kt.getRange('D1').setValue('Order Qty:').setFontWeight('bold');
    if (!kt.getRange('E1').getValue()) kt.getRange('E1').setValue(1);
    if (!kt.getRange('C2').getValue()) kt.getRange('C2').setValue('Description:').setFontWeight('bold');
    if (fg) {
      // Use Search Label column (col F=6) so typing filters by substring
      const col = SEARCH_LABEL_COL['FG Master'] || 1;
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(fg.getRange(2, col, Math.max(1, fg.getLastRow() - 1), 1), true)
        .setAllowInvalid(false).build();
      kt.getRange('C1').setDataValidation(rule);
    }
  }
}

// Search Label spec per master: {col, formula}. Col is 1-based; formula is
// written to row 2 of that col (ARRAYFORMULA spills down). SEARCH_LABEL_COL
// and SEARCH_LABEL_FORMULAS below are derived aliases — cross-file readers
// (30_, 40_) use them unchanged.
const SEARCH_LABEL_SPECS = {
  'Item Master':       { col: 20, formula: '=ARRAYFORMULA(IF(A2:A="","", A2:A & " — " & D2:D & IF(E2:E<>""," | "&E2:E,"") & IF(F2:F<>""," | "&F2:F,"")))' },
  'SFG Master':        { col: 14, formula: '=ARRAYFORMULA(IF(A2:A="","", A2:A & " — " & D2:D & IF(E2:E<>""," | "&E2:E,"") & IF(F2:F<>""," | "&F2:F,"")))' },
  'FG Master':         { col: 12, formula: '=ARRAYFORMULA(IF(A2:A="","", A2:A & " — " & D2:D & IF(E2:E<>""," | "&E2:E,"") & IF(F2:F<>""," | "&F2:F,"") & IF(G2:G<>""," | "&G2:G,"")))' },
  'Supplier Master':   { col: 10, formula: '=ARRAYFORMULA(IF(A2:A="","", A2:A & " — " & B2:B & IF(C2:C<>""," | "&C2:C,"")))' },
  'Job Worker Master': { col: 10, formula: '=ARRAYFORMULA(IF(A2:A="","", A2:A & " — " & B2:B & IF(C2:C<>""," | "&C2:C,"")))' },
  'Tool Master':       { col: 15, formula: '=ARRAYFORMULA(IF(A2:A="","", A2:A & " — " & D2:D & IF(E2:E<>""," | "&E2:E,"") & IF(F2:F<>""," | "&F2:F,"")))' },
};
const SEARCH_LABEL_COL = Object.fromEntries(Object.entries(SEARCH_LABEL_SPECS).map(function(e){ return [e[0], e[1].col]; }));
const SEARCH_LABEL_FORMULAS = Object.fromEntries(Object.entries(SEARCH_LABEL_SPECS).map(function(e){ return [e[0], e[1].formula]; }));

// Description ARRAYFORMULA per master — writes to B2 of each master. Combines
// Item Name (col C) + Size + Grade/Color (or just Item Name for FG).
const DESCRIPTION_FORMULAS = {
  'Item Master': '=ARRAYFORMULA(IF(D2:D="","", D2:D & IF(E2:E<>""," - "&E2:E,"") & IF(F2:F<>""," - "&F2:F,"")))',
  'SFG Master':  '=ARRAYFORMULA(IF(D2:D="","", D2:D & IF(E2:E<>""," - "&E2:E,"") & IF(F2:F<>""," - "&F2:F,"")))',
  'FG Master':   '=ARRAYFORMULA(IF(D2:D="","", D2:D & IF(E2:E<>""," - "&E2:E,"") & IF(F2:F<>""," - "&F2:F,"") & IF(G2:G<>""," - "&G2:G,"")))',
  'Tool Master': '=ARRAYFORMULA(IF(D2:D="","", D2:D & IF(E2:E<>""," ("&E2:E&")","") & IF(F2:F<>""," SN:"&F2:F,"")))',
};

// Write the Description ARRAYFORMULA to row 2 col C of each master.
// v3.5: col B = Photo (manual CellImage), col C = Description (this formula),
// col D = Item Name. Idempotent — skips if formula already present.
function _seedDescriptionFormulas(ss) {
  Object.keys(DESCRIPTION_FORMULAS).forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const cell = sh.getRange(2, 3); // col C
    if (!cell.getFormula()) cell.setFormula(DESCRIPTION_FORMULAS[name]);
    sh.setColumnWidth(2, 100);    // Photo — square-ish
    if (sh.getLastRow() > 1) sh.setRowHeights(2, sh.getLastRow() - 1, 80);
    sh.setColumnWidth(3, 260);    // Description
    sh.setColumnWidth(4, 180);    // Item Name
  });
}

// Write the Search Label ARRAYFORMULA to row 2 of the Search Label col.
// Idempotent: only writes if the formula cell is empty.
function _seedSearchLabels(ss) {
  Object.keys(SEARCH_LABEL_FORMULAS).forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const col = SEARCH_LABEL_COL[name];
    const cell = sh.getRange(2, col);
    if (!cell.getFormula()) cell.setFormula(SEARCH_LABEL_FORMULAS[name]);
    sh.setColumnWidth(col, 320);
  });
}


// ============================================================================
// ON HAND FORMULAS — SFG pulls from Stock Snapshot; FG = Opening minus Site
// Sends; Tool = 1 when Currently With is blank, else 0. Idempotent; only
// writes if the cell is empty.
// ============================================================================

const ON_HAND_FORMULAS = {
  'SFG Master': {
    col: 13,  // M (Available) — after Photo B, Desc C, Name D, Size E, Color F, UOM G, Storage H, MIN I, MAX J, Active K, Opening L
    // v3.6 SFG Master col M = Available (physical − in-process − reserved).
    // VLOOKUP col 6 = F in SFG snapshot section (Available). RM/SFG codes don't
    // collide (different prefixes) so global A:F lookup finds SFG rows.
    formula: "=ARRAYFORMULA(IF(A2:A=\"\",\"\",IFERROR(VLOOKUP(A2:A,'Stock Snapshot'!A:F,6,FALSE),0)))"
  },
  'FG Master': {
    col: 11,  // K (Available) — subtracts active reservations
    // Available = Opening + Adjustments − FG Delivery Main + Returned − Active Reservations
    // Reservation Register: D=Code, F=Qty, I=Status (Active only)
    formula: "=ARRAYFORMULA(IF(A2:A=\"\",\"\",IF(J2:J=\"\",0,J2:J)+IFERROR(SUMIF('Adjustments Register'!D:D,A2:A,'Adjustments Register'!F:F),0)-IFERROR(SUMIFS('FG Delivery Lines'!F:F,'FG Delivery Lines'!D:D,A2:A,'FG Delivery Lines'!C:C,\"Main\"),0)+IFERROR(SUMIFS('FG Delivery Lines'!J:J,'FG Delivery Lines'!D:D,A2:A,'FG Delivery Lines'!C:C,\"Main\"),0)-IFERROR(SUMIFS('Reservation Register'!F:F,'Reservation Register'!D:D,A2:A,'Reservation Register'!I:I,\"Active\"),0)))"
  },
  'Tool Master': {
    col: 13,  // M (On Hand) — Currently With at K, Opening at L. In-store on-hand = Opening + Adjustments; issued = 0.
    formula: "=ARRAYFORMULA(IF(A2:A=\"\",\"\",IF(K2:K=\"\",IF(L2:L=\"\",1,L2:L)+IFERROR(SUMIF('Adjustments Register'!D:D,A2:A,'Adjustments Register'!F:F),0),0)))"
  },
};

function _seedOnHandFormulas(ss) {
  Object.keys(ON_HAND_FORMULAS).forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const spec = ON_HAND_FORMULAS[name];
    const cell = sh.getRange(2, spec.col);
    if (!cell.getFormula()) cell.setFormula(spec.formula);
    sh.setColumnWidth(spec.col, 90);
    sh.setColumnWidth(spec.col - 1, 90); // Opening Qty width too
  });
}

// ============================================================================
// REBUILD ALL FORMULAS — one-click after CSV paste. Clears the data-row range
// of the Description col (C) + Search Label col + On Hand col per master,
// then re-writes the ARRAYFORMULA at row 2 so it spills cleanly.
// Empty-string cells from a CSV paste block the spill; this scrubs them.
// Safe to re-run.
// ============================================================================

function rebuildAllFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = [];

  // Description formulas — col C for Item/SFG/FG/Tool (v3.5: col B is Photo)
  Object.keys(DESCRIPTION_FORMULAS).forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) { report.push(name + ': empty, skipped'); return; }
    sh.getRange(2, 3, sh.getLastRow() - 1, 1).clearContent();
    sh.getRange(2, 3).setFormula(DESCRIPTION_FORMULAS[name]);
    report.push(name + ' Description: refilled');
  });

  // Search Label formulas per SEARCH_LABEL_SPECS
  Object.keys(SEARCH_LABEL_SPECS).forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const spec = SEARCH_LABEL_SPECS[name];
    sh.getRange(2, spec.col, sh.getLastRow() - 1, 1).clearContent();
    sh.getRange(2, spec.col).setFormula(spec.formula);
    report.push(name + ' Search Label (col ' + spec.col + '): refilled');
  });

  // On Hand formulas per ON_HAND_FORMULAS (SFG/FG/Tool)
  Object.keys(ON_HAND_FORMULAS).forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const spec = ON_HAND_FORMULAS[name];
    sh.getRange(2, spec.col, sh.getLastRow() - 1, 1).clearContent();
    sh.getRange(2, spec.col).setFormula(spec.formula);
    report.push(name + ' On Hand (col ' + spec.col + '): refilled');
  });

  SpreadsheetApp.getUi().alert('Formulas rebuilt:\n\n' + report.join('\n'));
}

// The six views inside the Explorer sheet. Order matters — first is default.
const EXPLORER_VIEWS = [
  'Item Ledger',
  'Supplier Ledger',
  'Stock by Category',
  'Non-Moving Stock',
  'Low Stock Alerts',
  'Pending Job Work',
  'Reorder Suggestions',
  'Site Consumption',
  'Project Requisitions',
  'QC Log',
];

function _dv(sheet, row, col, numRows, numCols, sourceRange) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange, true).setAllowInvalid(false).build();
  sheet.getRange(row, col, numRows, numCols).setDataValidation(rule);
}

function _dvList(sheet, row, col, numRows, numCols, list) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  sheet.getRange(row, col, numRows, numCols).setDataValidation(rule);
}

// ============================================================================
// CONFIG READ/WRITE — used everywhere
// ============================================================================

function getConfig(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sh) return null;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) if (data[i][0] === key) return data[i][1];
  return null;
}

function setConfig(key, value) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sh) throw new Error('Config sheet missing');
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i + 2, 2).setValue(value); return; }
  }
  sh.appendRow([key, value, '']);
}

// ============================================================================
// FORM BOOTSTRAP — populates dropdowns in 05_Forms.html on load.
// Real form data (items, sfgs, suppliers, job workers) fills in as Phase 1
// creates the masters. This function survives across phases.
// ============================================================================

// Run from menu — shows the exact size of every dropdown list + any bootstrap error.
// Use when a form's dropdowns are empty and you don't know why.
function diagnoseDropdowns() {
  const b = getFormBootstrap('diagnose');
  const lines = [];
  lines.push('Bootstrap counts:');
  Object.keys(b).forEach(function(k) {
    const v = b[k];
    if (Array.isArray(v)) lines.push('  ' + k + ': ' + v.length + ' item(s)');
    else lines.push('  ' + k + ': ' + JSON.stringify(v));
  });
  if (b._bootstrapError) lines.push('\n⚠ OUTER FAILURE: ' + b._bootstrapError);
  lines.push('\nCheck View → Executions in Apps Script for individual helper failures (logged with Logger.log).');
  SpreadsheetApp.getUi().alert('Dropdown Diagnostic', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function getFormBootstrap(formId) {
  // Outermost try/catch — GUARANTEES the client always receives a valid object,
  // even if the whole workbook is misconfigured. Empty arrays render as
  // empty dropdowns rather than causing the form to fail silently.
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    function _safe(fn, fallback) {
      try { return fn(); }
      catch (e) { Logger.log('getFormBootstrap helper failed [' + fn.name + ']: ' + e.message + '\n' + e.stack); return fallback; }
    }
    const out = {
      categories:    _safe(function() { return _colValues(ss, 'Dropdowns', 1); }, []),
      materialTypes: _safe(function() { return _colValues(ss, 'Dropdowns', 3); }, []),
      uoms:          _safe(function() { return _colValues(ss, 'Dropdowns', 5); }, []),
      processTypes:  _safe(function() { return _colValues(ss, 'Dropdowns', 7); }, []),
      employees:     _safe(function() { return _colValues(ss, 'Dropdowns', 9); }, []),
      projects:      _safe(function() { return _colValues(ss, 'Dropdowns', 11).concat(_splitCsv(getConfig('Project List'))); }, []),
      items:         _safe(function() { return _codeDescUom(ss, 'Item Master', 1, 3, 7, 19); }, []),
      sfgs:          _safe(function() { return _codeDescUom(ss, 'SFG Master',  1, 3, 7, 11); }, []),
      suppliers:     _safe(function() { return _codeDescActive(ss, 'Supplier Master', 1, 2, 9); }, []),
      jobWorkers:    _safe(function() { return _codeDescActive(ss, 'Job Worker Master', 1, 2, 9); }, []),
      openJwc:       _safe(function() { return _openJwcList(ss); }, []),
      activePos:     _safe(function() { return _activePosList(ss); }, []),
      amendablePos:  _safe(function() { return _amendablePosList(ss); }, []),
      recentDcs:     _safe(function() { return _recentDcsList(ss); }, []),
      toolsInStore:  _safe(function() { return _toolsList(ss, 'InStore'); }, []),
      toolsOutstanding: _safe(function() { return _openMovementsList(ss); }, []),
      tools:         _safe(function() { return _toolsList(ss, null); }, []),
      highValueThreshold: _safe(function() { return Number(getConfig('High Value Threshold')) || 100000; }, 100000),
      company: _safe(function() { return {
        name:    getConfig('Company Name')    || '',
        address: getConfig('Company Address') || '',
        gstin:   getConfig('Company GSTIN')   || '',
        pan:     getConfig('Company PAN')     || '',
        phone:   getConfig('Company Phone')   || '',
        email:   getConfig('Company Email')   || '',
      }; }, {}),
    };
    out.fgs = _safe(function() { return _codeDescUom(ss, 'FG Master', 1, 3, 8, 9); }, []);
    out.itemsAndFgs = (out.items || []).concat(out.fgs || []);
    return out;
  } catch (e) {
    Logger.log('getFormBootstrap outer failure: ' + e.message + '\n' + e.stack);
    // Return a fully-formed empty shape so client dropdowns don't crash
    return {
      categories: [], materialTypes: [], uoms: [], processTypes: [], employees: [], projects: [],
      items: [], sfgs: [], fgs: [], itemsAndFgs: [],
      suppliers: [], jobWorkers: [], openJwc: [], activePos: [], amendablePos: [], recentDcs: [],
      toolsInStore: [], toolsOutstanding: [], tools: [],
      highValueThreshold: 100000,
      _bootstrapError: e.message,
    };
  }
}

function _colValues(ss, sheetName, colNum) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const lastData = nextRowByCol(sh, colNum) - 1;
  if (lastData < 2) return [];
  return sh.getRange(2, colNum, lastData - 1, 1).getValues()
    .map(function(r) { return r[0]; }).filter(String);
}

// Returns [{code, desc, uom, label}] where label is the composite from Search
// Label col (falls back to code + desc if Search Label is blank).
// Uses actual last data row (via nextRowByCol) to avoid reading phantom rows
// from ARRAYFORMULA spill — was making getFormBootstrap slow.
function _codeDescUom(ss, name, codeCol, descCol, uomCol, activeCol) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const lastData = nextRowByCol(sh, codeCol) - 1;
  if (lastData < 2) return [];
  const data = sh.getRange(2, 1, lastData - 1, sh.getLastColumn()).getValues();
  const labelCol = SEARCH_LABEL_COL[name];
  // ponytail: v.sqm bootstrap payload dropped Oct-2026 — Sqm per Unit column
  // removed from Item Master. GRN/PO forms accept per-row sqm entry instead.
  const out = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[codeCol - 1]) continue;
    if (activeCol && row[activeCol - 1] === false) continue;
    const composite = (labelCol && row[labelCol - 1])
      ? row[labelCol - 1]
      : row[codeCol - 1] + ' — ' + (row[descCol - 1] || '');
    out.push({
      code: row[codeCol - 1],
      desc: row[descCol - 1] || '',
      uom:  row[uomCol - 1]  || '',
      label: composite,
    });
  }
  return out;
}

function _codeDescActive(ss, name, codeCol, descCol, activeCol) {
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const lastData = nextRowByCol(sh, codeCol) - 1;
  if (lastData < 2) return [];
  const data = sh.getRange(2, 1, lastData - 1, sh.getLastColumn()).getValues();
  const labelCol = SEARCH_LABEL_COL[name];
  const out = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[codeCol - 1]) continue;
    if (activeCol && row[activeCol - 1] === false) continue;
    const composite = (labelCol && row[labelCol - 1])
      ? row[labelCol - 1]
      : row[codeCol - 1] + ' — ' + (row[descCol - 1] || '');
    out.push({ code: row[codeCol - 1], label: composite });
  }
  return out;
}

// Full supplier details for the PO form's "Verify Party" block.
// Supplier Master cols: A=Code, B=Firm Name, C=Category, D=Contact, E=Mobile,
//   F=Email, G=GST No, H=Address, I=Active.
function getSupplierDetails(code) {
  try {
    if (!code) return { success: false };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Supplier Master');
    if (!sh || sh.getLastRow() < 2) return { success: false };
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === code) {
        return { success: true,
          name: data[i][1] || '', category: data[i][2] || '',
          contact: data[i][3] || '', mobile: data[i][4] || '',
          email: data[i][5] || '', gstin: data[i][6] || '',
          address: data[i][7] || '' };
      }
    }
    return { success: false };
  } catch (e) { return { success: false, message: e.message }; }
}

// Active POs for GRN form dropdown — anything not fully closed. Newest first.
// PO Register cols: A=Sr, B=PO No, C=Date, D=Sup Code, E=Sup Name, F=Delivery Date,
//   G=Status, H=Total Value.
// Status vocabulary mixes 'Draft'/'Issued' (initial) with 'Open'/'Partial'/'Received'
// (after updatePoStatus runs). Anything except 'Received' or 'Cancelled' is active.
function _activePosList(ss) {
  const sh = ss.getSheetByName('Purchase Order Register');
  if (!sh) return [];
  const lastData = nextRowByCol(sh, 1) - 1;
  if (lastData < 2) return [];
  // Never read beyond the sheet's actual column count (old workbooks may have fewer cols)
  const width = Math.min(8, sh.getLastColumn() || 8);
  if (width < 7) return [];  // need at least through Status col G
  const data = sh.getRange(2, 1, lastData - 1, width).getValues();
  const CLOSED = { 'Received': true, 'Cancelled': true, 'Closed': true };
  const out = [];
  data.forEach(function(r) {
    const poNo = r[1], status = r[6] || 'Draft';
    if (!poNo) return;
    if (CLOSED[status]) return;
    const ts = (r[2] instanceof Date) ? r[2].getTime() : 0;
    const dateStr = (r[2] instanceof Date)
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'dd-MMM-yy') : '';
    // NOTE: send only primitives to the client — Date objects can break
    // google.script.run's serialization on some workbooks.
    out.push({
      code: poNo,
      ts: ts,
      label: poNo + '  —  ' + (r[4] || r[3] || '') + '  (' + status + (dateStr ? ', ' + dateStr : '') + ')',
    });
  });
  // Newest first (sort by numeric timestamp)
  out.sort(function(a, b) { return (b.ts || 0) - (a.ts || 0); });
  return out;
}

// POs amendable through the Amend form. Includes everything NOT Cancelled —
// 'Received' POs CAN be amended but the server shows a confirm (goods already
// came in). The caller (amendPoRows) decides whether to allow it.
function _amendablePosList(ss) {
  const sh = ss.getSheetByName('Purchase Order Register');
  if (!sh) return [];
  const lastData = nextRowByCol(sh, 1) - 1;
  if (lastData < 2) return [];
  const width = Math.min(13, sh.getLastColumn() || 12);
  if (width < 7) return [];
  const data = sh.getRange(2, 1, lastData - 1, width).getValues();
  const out = [];
  data.forEach(function(r) {
    const poNo = r[1], status = r[6] || 'Draft';
    if (!poNo) return;
    if (status === 'Cancelled') return;
    const rev = (width >= 13) ? (Number(r[12]) || 0) : 0;
    const ts = (r[2] instanceof Date) ? r[2].getTime() : 0;
    const dateStr = (r[2] instanceof Date)
      ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'dd-MMM-yy') : '';
    const revTag = rev > 0 ? ' R' + rev : '';
    out.push({
      code: poNo, ts: ts,
      label: poNo + revTag + '  —  ' + (r[4] || r[3] || '') + '  (' + status + (dateStr ? ', ' + dateStr : '') + ')',
    });
  });
  out.sort(function(a, b) { return (b.ts || 0) - (a.ts || 0); });
  return out;
}

// Recent FG Delivery Challans for the Site Return picker. Newest first; capped
// at 100 so the dropdown stays usable. Includes customer and date in label.
function _recentDcsList(ss) {
  const sh = ss.getSheetByName('FG Delivery Register');
  if (!sh) return [];
  const lastData = nextRowByCol(sh, 1) - 1;
  if (lastData < 2) return [];
  const width = Math.min(5, sh.getLastColumn() || 5);  // A..E = Sr, DC, Date, Cust, Project
  if (width < 4) return [];
  const data = sh.getRange(2, 1, lastData - 1, width).getValues();
  const out = [];
  data.forEach(function(r) {
    const dcNo = r[1], date = r[2], cust = r[3] || '', project = (width >= 5 ? r[4] : '') || '';
    if (!dcNo) return;
    const ts = (date instanceof Date) ? date.getTime() : 0;
    const dateStr = (date instanceof Date)
      ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MMM-yy') : '';
    out.push({
      code: dcNo, ts: ts,
      label: dcNo + '  —  ' + (cust || project || '(no customer)') + (dateStr ? '  (' + dateStr + ')' : ''),
    });
  });
  out.sort(function(a, b) { return (b.ts || 0) - (a.ts || 0); });
  return out.slice(0, 100);
}

function _openJwcList(ss) {
  const sh = ss.getSheetByName('Job Work Outward Register');
  if (!sh) return [];
  const lastData = nextRowByCol(sh, 1) - 1;
  if (lastData < 2) return [];
  const data = sh.getRange(2, 1, lastData - 1, sh.getLastColumn()).getValues();
  const seen = {}, out = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const jwc = row[1], status = row[13], jwName = row[4]; // cols B, N, E
    if (status !== 'Open' || !jwc || seen[jwc]) continue;
    seen[jwc] = true;
    out.push({ code: jwc, label: jwc + ' — ' + jwName });
  }
  return out;
}

// Tool Master → composite list, filter by status.
// v3.5 idx: [2]=Desc, [9]=Status(J), [13]=Active(N), [14]=Search Label(O).
function _toolsList(ss, statusFilter) {
  const sh = ss.getSheetByName('Tool Master');
  if (!sh) return [];
  const lastData = nextRowByCol(sh, 1) - 1;
  if (lastData < 2) return [];
  const data = sh.getRange(2, 1, lastData - 1, 15).getValues();
  const out = [];
  data.forEach(function(r) {
    if (!r[0] || r[13] === false) return;                                   // Active col N (idx 13)
    if (statusFilter === 'InStore' && r[9] !== 'In Store') return;          // Status col J (idx 9)
    out.push({ code: r[0], label: r[14] || (r[0] + ' — ' + (r[2] || '')) }); // Search Label col O (idx 14)
  });
  return out;
}

// Tool Movement Register — outstanding Issue movements (Actual Return blank)
function _openMovementsList(ss) {
  const sh = ss.getSheetByName('Tool Movement Register');
  if (!sh) return [];
  const lastData = nextRowByCol(sh, 1) - 1;
  if (lastData < 2) return [];
  const data = sh.getRange(2, 1, lastData - 1, 15).getValues();
  const out = [];
  data.forEach(function(r) {
    if (r[5] !== 'Issue') return; // only Issue rows
    if (r[9]) return; // already returned
    const label = r[1] + ' — ' + r[3] + ' with ' + r[6] +
      (r[8] ? ' (due ' + dateString(r[8]) + ')' : '');
    out.push({ code: r[1], label: label });
  });
  return out;
}

function _splitCsv(s) {
  if (!s) return [];
  return String(s).split(',').map(function(x) { return x.trim(); }).filter(String);
}

// Called by JW-Inward form when a reference challan is picked.
function getOpenJwcLines(jwcNo) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Job Work Outward Register');
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const out = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row[1] === jwcNo && row[13] === 'Open') {
      out.push({ sfgCode: row[6], desc: row[7], qtySent: row[8] });
    }
  }
  return out;
}

// ============================================================================
// FORM HANDLER STUBS — Phase 1+ replaces these with real logic.
// Kept here so Phase 0 forms don't throw "function not found" on submit.
// ============================================================================

// All form handlers live in 40_Forms.gs