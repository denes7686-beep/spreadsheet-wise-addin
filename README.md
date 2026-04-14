# Spreadsheet Wise Office Add-in MVP

This folder contains a first-pass Office Add-in scaffold for migrating the current Google Apps Script workflow to Excel Online.

## Scope of this MVP

Implemented in the stable base version:

- `Create New Invoice`
- `Submit Invoice`
- `Add To Transactions`
- `Update Formats`
- workbook validation for required sheets

Deferred to later phases:

- PDF export workflow
- email workflow
- saving PDF to OneDrive or SharePoint
- storing public PDF links

Those features are better handled through Microsoft Graph, Power Automate, or a small backend service.

## Workbook assumptions

The add-in is built against the workbook structure found in:

- [Spreadsheet Wise.xlsx](C:/Users/March/Documents/Invoice%20-%20Excell/Spreadsheet%20Wise.xlsx)

Required sheets:

- `Transactions`
- `Invoice Template`
- `Configuration`
- `Currencies`

Expected key cells:

- `Configuration!C4` payment terms
- `Configuration!C6` last invoice number
- `Configuration!C10` date format label
- `Configuration!C12` currency symbol
- `Invoice Template!C5:C7` status and dates
- invoice client block in column `D`
- `Transactions` table starting on row `14`

## Files

- [manifest.xml](C:/Users/March/Documents/Invoice%20-%20Excell/office-addin/manifest.xml)
- [src/taskpane.html](C:/Users/March/Documents/Invoice%20-%20Excell/office-addin/src/taskpane.html)
- [src/taskpane.css](C:/Users/March/Documents/Invoice%20-%20Excell/office-addin/src/taskpane.css)
- [src/taskpane.js](C:/Users/March/Documents/Invoice%20-%20Excell/office-addin/src/taskpane.js)

## Setup path

1. Host the contents of `src/` and `assets/` on HTTPS.
2. Update the URLs inside `manifest.xml` to match that host.
3. Upload or sideload the manifest into Excel on the web.
4. Open the workbook and launch the add-in from the ribbon.

## Recommended hosting options

- GitHub Pages for static task pane files
- Azure Static Web Apps
- Any HTTPS web host

## Phase 2 options

For `Send Email + PDF + Storage`, use one of these:

- Office Add-in + Power Automate
- Office Add-in + Microsoft Graph
- Office Add-in + lightweight backend API

For the fastest MVP, I recommend `Office Add-in + Power Automate`.

## Current stable flow

1. Create and finalize the invoice in Excel.
2. Click `Submit Invoice`.
3. Review the saved row in `Transactions`.

## Saving a stable version

Recommended:

1. Keep this exact `office-addin` folder in GitHub.
2. Create a separate backup copy named something like `office-addin-stable-v1`.
3. When you start future changes, work in the main folder and leave the stable copy untouched.

Best long-term option:

- create a GitHub release or a separate `stable-v1` branch
- keep a zip backup of the stable folder on your computer
