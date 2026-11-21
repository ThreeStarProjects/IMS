# IMS - Inventory Management System

Google Apps Script bound to the "3SP IMS" spreadsheet. Stock, suppliers, purchasing, job work, deliveries, reservations and the PDFs that go with them. Runs from the sheet's menus and forms (it is not a library; cross-system calls reach it through its own queues and a web-app doorway).

> Part of the 3SP family: Masters-System, Project-Management, Finance-System, Ticketing, 3SP_DocGen. See "How the systems fit together" at the bottom.

## What it holds

- **Masters:** Item Master (raw material / hardware), SFG Master (semi-finished goods), FG Master (finished goods), Supplier Master, Job Worker Master, Dropdowns. Suppliers and job workers are mastered here (SUPP- / JW- codes) and copied to the Masters system; employees and projects come from Masters.
- **Registers:** Inward (GRN), Purchase Orders + lines, Material Requisition, Adjustments (opening stock and corrections), Job Work Outward / Inward, FG Delivery (Delivery Challans) + lines with returns from site, Reservation Register, QC Register.
- **Stock Snapshot** (raw material and SFG sections), dashboards and **Explorer views** (Item Ledger, Supplier Ledger, Stock by Category, Site Consumption, Project Requisitions, QC Log).
- **Documents** (PDF to Drive): PO (and amendments), GRN, requisition slip, Delivery Challan, RFQ, Job Work Challan.

## The stock maths (the rules that matter)

- Raw material **physical** = opening + GRNs - requisitions - hardware sent on Delivery Challans (net of returns) +/- adjustments - material currently at a job worker.
- **Available = physical - open reservations** (a reservation holds `Qty - Consumed Qty`; statuses Requested, Active, Partially Consumed, Consumed, Released, Error).
- SFG: on hand = opening + adjustments + job-work returns - sent to job work - shipped on challans; Available = on hand - reserved.
- A **Material Requisition** or a **Delivery Challan** for a project **consumes** that project's open reservations of the same item (oldest first, partial allowed, `PRJ-###` id or name match). The oversell warning adds back what the issuing project itself reserved.
- Project cost views (Site Consumption, Project Requisitions) include hardware sent on Customer Delivery Challans, net of returns. Finance imports the same lines.
- Returns from site restore stock but do not recreate the reservation.

## Document numbers

Standard format `TSP/<TYPE>/<FY>/<NNN>` (for example `TSP/PO/26-27/043`). IMS issues **PO, DC, RFQ and JWC** numbers (counters kept in Config as `<KIND> Counter <FY>`, FY April-March); DocGen issues the rest. When a document made in DocGen is edited, the same document is renamed to the next version letter (`043A`, `043B`, ...), so a GRN picker only ever lists the latest version and nothing is counted twice.

## Connections to the other systems

- **Masters** (library, pinned version in `appsscript.json`): projects, employees, customers, suppliers copy. Dropdowns refresh on open and from the Interlink menu.
- **Project-Management:** PM leaves reservation requests in the Reservation Register and commands in a hidden sheet; a one-minute trigger turns them into real reservations / items / suppliers (`_Interlink_Reservations.gs`). Opening IMS recreates the trigger if it went missing.
- **Finance:** pulls requisition and delivery-hardware costs for `PRJ-###` projects every 10 minutes.
- **3SP_DocGen (web app doorway, `_Interlink_DocGen.gs`):** `POST {secret, action, data}` with a key kept in the script property `DOCGEN_SECRET` (menu Interlink > *DocGen link: show key*). Actions: `feeds`, `projects`, `peek`, `recordPo`, `recordDc`, `recordRfq`, `recordJwc`, `photos`. Everything is idempotent by the DocGen document id (log sheet `_DocGen_Log`); the same content saved twice does nothing, a changed version replaces the lines in place. A new web-app version must be deployed after code changes to this file.

## Files

`00_Config_Menu_Setup.gs` config, menus, form launchers, workbook setup · `05_Forms.html` all forms · `10_Masters.gs` master data and lookups · `20_Registers.gs` the posting functions (GRN, PO, requisition, delivery, job work, ...) · `30_Snapshots.gs` stock snapshots, dashboards, explorer views · `40_Forms.gs` form handlers, reservations, returns, kit builder · `50_Docs.gs` PDFs · `60_Utils_Audit_Triggers.gs` helpers, document numbering, audit, triggers · `80_Formatting.gs` · `90_CodeHygiene.gs` · `_Interlink_Masters.gs`, `_Interlink_Reservations.gs`, `_Interlink_DocGen.gs` the connections.

## Deploy / development

- `clasp push --force`, commit, push to `main`. The GitHub Action (`.github/workflows/deploy.yml`) repeats the clasp push and needs the repo secret `CLASPRC_JSON` (if it fails with `invalid_rapt`, run `clasp login` and refresh the secret).
- Code that runs inside IMS uses the active spreadsheet, so other systems cannot call it as a library; keep cross-system work in the `_Interlink_*` queues or the doorway.
- Edit existing `.gs` files; add a new file only when the code would slow the current system down.
- Offline tests (Node `vm`) are kept in `_tests` next to the repos, not in git.

## How the systems fit together

```
              Masters-System  (library: projects, customers, parties, employees, sites)
             /     |      |      \
     Finance   IMS   Project-Mgmt  Ticketing         3SP_DocGen (Next.js) --web app--> IMS
        ^        ^        |                              (documents, numbers, PDFs)
        |        |--reservations / requests-----|
        |--- 10-minute import of project material costs from IMS
```
