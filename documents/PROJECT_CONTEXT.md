# Excel Merger Project Context

## Purpose
This app is a client-side Excel operations portal for customs workflows. It runs as a static web app served by Express, with all XLSX processing in the browser via SheetJS.

## Runtime and stack
- Runtime: Node.js + Express static server
- Frontend: plain HTML/CSS/JS
- Excel processing: `xlsx` (SheetJS) in browser
- File download: `FileSaver`
- Entrypoint server: `server.js`
- Main UI/logic: `public/index.html`, `public/styles.css`, `public/script.js`, `public/dt-header-workflow.js`

## Scope (current)
The project currently has 3 user-facing tools:

1. Excel Merger
- Accepts multiple `.xlsx` files.
- Row slicing controls:
  - First file rows to remove: default `4`
  - Other files rows to remove: default `5`
- Optional "Add filename as first column".
- Merges rows into a single sheet and exports `merged.xlsx`.
- Generates a summary report on the right panel:
  - Total rows
  - Total duty
  - Total GST
  - Brokerage value counts (`0.0175`, `0.085`, `0.71`, `0.28`)
- Number formatting normalization:
  - DutiesHeader: converts `Value for Duty -> Exchange Rate` range (or fallback columns J..Q) to numeric General format.
  - DutiesItem: converts `Duty` and `Gov. Sales Tax` to numeric General format.

2. D/T Header File Modifier
- Inputs:
  - `CLIENT`
  - `RPT NAME`
  - `RPT DATE` (kept as text)
  - Source (SFTP file)
  - Target (`_DutiesHeader` file)
  - Optional DutiesItem (`_DutiesItem` file)
- Rewrites rows 1..3 in both header/item outputs:
  - `CLIENT:`
  - `RPT NAME:`
  - `RPT DATE :`
- Detects whether DutiesHeader and DutiesItem headers are on Excel row 4 or 5.
- If headers are found on row 4, inserts one blank row so output headers land on row 5.
- Analyze step remains first:
  - compares DutiesHeader `8308...` entries in H/J against SFTP AC/AS
  - blocks user review behind `Analyze 8308 Values` before modify/download
- Reads target CCNs from column H starting below the detected header row, with `8308` prefix cleanup on target side.
- Reads source values:
  - AC (CCN candidate)
  - AS (value for column J)
- Deduplicates source AC values before insert.
- Inserts only new CCNs (exact match check against cleaned target set).
- New row mapping:
  - A = `CLVS`
  - B and H = source AC
  - J = source AS
  - C..F copied from last existing non-empty target data row
  - K..Q = `0`
  - R = `DDP`
- Inserts rows right after last non-empty target row (avoids large blank-gap append issue).
- Converts `Value for Duty -> Exchange Rate` to numeric General format in output.
- Optional DutiesItem processing:
  - ensures a `CCN` column exists
  - builds a `Transaction Number -> CCN` lookup from the modified DutiesHeader output
  - writes matching CCNs directly into DutiesItem
  - leaves unmatched transaction numbers blank
- Auto-downloads output(s) with refreshed 12-digit timestamp before `_DutiesHeader` / `_DutiesItem` when applicable.

3. Header/Item Analyzer
- Inputs:
  - DutiesHeader required
  - DutiesItem optional
- Header report sections:
  - Total CCNs
  - Total CLVS, LVS, PGA by brokerage fee buckets
  - Empty Brokerage Fee CCNs
  - Empty Value for Duty CCNs
  - GST=0 with threshold CCNs
    - Threshold `>20.1` normally
    - Threshold `>40.1` when brokerage fee is `2.25`
  - Value for Duty `<20` with Duty/GST `>0` CCNs
- Item section display is currently removed from report body.
- If item file is provided, bottom "Totals Match" compares Header vs Item totals for Duty and GST.

## UI structure
- Left sidebar navigation switches sections:
  - `showDisplay('merger'|'modify'|'analyzer')`
- Main sections:
  - `#excel-merger`
  - `#modifyTool`
  - `#analyzerTool`
- Scroll-to-top floating button appears after vertical scroll.

## File map
- `server.js`: static hosting only (no backend Excel processing).
- `public/index.html`: all tool layouts and controls.
- `public/styles.css`: layout/responsive styling, tool sections, report formatting.
- `public/script.js`: all app behavior and Excel logic.
- `public/dt-header-workflow.js`: shared D/T header/item normalization and CCN propagation helpers.
- `views/index.ejs`: legacy minimal file, not primary UI.

## Key implementation assumptions
- Merger assumes first file contributes the merged header row.
- Merger currently skips first post-slice row only for file index 0 (first file), keeps row 0 for subsequent files.
- SFTP AC/AS start row for modifier is fixed at row 3 (index 2).
- D/T workflow header detection scans the first 10 rows and expects:
  - Header mode: `Transaction Number` and `CCN`
  - Item mode: `Transaction Number` and `Goods Description`
- Analyzer header detection scans first 15 rows and chooses best match row by keyword scoring.

## Known risks for future work
- In merger analysis, `analyzeData` expects `mergedData[0]`; if all inputs trim to empty, this can fail.
- Slicing inputs are parsed with `parseInt`; invalid/blank values should be guarded if stricter UX is required.
- Column matching relies on text labels; variant header names may require matcher expansion.

## Local run
- Install dependencies: `npm install`
- Dev: `npm run dev`
- Prod-like: `npm start`
- Default server port: `8080` (or `process.env.PORT`)

## Guidance for agents adding features
- Preserve current client-side processing model unless explicitly migrating to backend.
- Keep feature modules isolated in `public/script.js` (existing pattern: IIFE per tool).
- Avoid changing column index constants without documenting affected workflow rules.
- When changing row slicing/header rules, test with:
  - single file merge
  - multi-file merge with/without leading blank rows
  - add-filename on/off
