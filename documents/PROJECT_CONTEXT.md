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
  - DutiesItem: converts recognized numeric fields, including `Duty`, `GST` / `Gov. Sales Tax`, `HST`, `PST`, `SIMA`, and `Surtax`, to numeric format.

## Current progress (2026-09-25)
- The D/T Header + Item workflow supports the expanded tax schema found in the September 2026 reports:
  - Header accepts `GST` as an alias for `Gov. Sales Tax` and preserves `HST`, `PST`, `SIMA`, and `Surtax` between GST and Brokerage Total.
  - Item accepts the same GST alias and preserves `HST`, `PST`, `SIMA`, and `Surtax` between GST and Inco Terms.
  - Newly generated Header rows are populated by header label rather than fixed A:R positions. The added tax fields are initialized to numeric zero and `DDP` is written to the actual Inco Terms column.
  - The completion report compares Header and Item totals for each added tax field when that field exists in both reports.
  - Added tax fields participate in blank/zero validation and appear in the Header and Item field-count breakdowns.
- D/T Header and Item post-processing validation is implemented and hardened.
- Header validation reports blank values for all monitored fields, but zero `Duty` and `Gov. Sales Tax` findings only for LVS/PGA rows. CLVS zero Duty/GST values are expected and excluded; Header `Value for Duty` remains checked for every row.
- The final D/T report UI is now split into `Overview`, `Header checks`, and `Item checks` tabs. Each validation tab shows its warning count, field totals, provenance/expected-zero notes, and a scrollable affected-record list. Tabs support keyboard navigation and responsive layouts.
- Verification against `Test files/112-05240631`:
  - Final Header findings: 36 (2 `Value for Duty` zeros and 34 `Duty` zeros).
  - Expected CLVS zero values excluded: 282 (141 Duty + 141 GST).
  - Final Item findings: 40 (38 Duty zeros and 2 Duty blanks); no Quantity, Value for Duty, Value for Tax, or GST findings in that fixture.
- Verification completed locally with `node --check` for the changed JavaScript files, `git diff --check`, and `npm.cmd run test:regression` (all smoke sections passed).
- Rollout boundary: these checks are local fixture/CLI evidence only. No packaged desktop rebuild, deployment, or authenticated production/runtime verification has been performed. Before release, restart/rebuild the app from the current working tree and manually exercise Analyze 8308 -> Proceed with Modify, including Header-only and Header+Item runs.

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
  - compares DutiesHeader `8308...` CCNs and column J values against SFTP AC/AS
  - blocks user review behind `Analyze 8308 Values` before modify/download
- Resolves target CCNs per row using `CCN`, then `Cargo Control Number`, then `Order Number` (first nonblank value), with `8308` prefix cleanup on the target side. The same precedence applies to the 8308 review, insertion deduplication, classification, validation identifiers, and Item mapping.
- Reads source values:
  - AC (CCN candidate)
  - AS (value for column J)
- Deduplicates source AC values before insert.
- Inserts only new CCNs (exact match check against cleaned target set).
- New row mapping:
  - The D/T Header template retains its existing fixed column layout; alias support accepts alternate identifier labels/blank CCN cells and does not imply arbitrary column reordering or deletion.
  - A = `CLVS`
  - B and H = source AC
  - J = source AS
  - C..F copied from last existing non-empty target data row
  - K..Q = `0`
  - R = `DDP`
- Inserts rows right after last non-empty target row (avoids large blank-gap append issue).
- Applies brokerage automation from `public/brokerage-rates.json` using `CLIENT`:
  - `CCN` starting with `8308` => `PGA`
  - `Transaction Number` starting with `LV` => `LVS`
  - `Transaction Number === CLVS` => `CLVS`
  - if client is not found in JSON, brokerage fee is left blank for classified rows
- Overwrites `Shipment Date`, `Arrival Date`, and `Release Date` with `RPT DATE`.
- Sets `Exchange Rate` to `0`.
- Sorts header data rows by brokerage fee descending.
- Converts `Value for Duty -> Exchange Rate` to numeric accounting format in Header output, including `HST`, `PST`, `SIMA`, and `Surtax` when present.
- Converts the recognized numeric Item fields to numbers, including `GST` / `Gov. Sales Tax`, `HST`, `PST`, `SIMA`, and `Surtax` when present.
- Optional DutiesItem processing:
  - ensures a `CCN` column exists
  - builds a `Transaction Number -> CCN` lookup from the modified DutiesHeader output
  - writes matching CCNs directly into DutiesItem
  - leaves unmatched transaction numbers blank
- Auto-downloads output(s) using the SFTP filename base with refreshed 12-digit timestamp and `_DutiesHeader` / `_DutiesItem`.
- Final modify report includes:
  - inserted header row count
  - PGA / LVS / CLVS counts
  - header Duty and GST totals
  - blank brokerage fee row count
  - whether the client matched the brokerage JSON
  - header vs item Duty/GST total match when DutiesItem is provided
  - header vs item HST/PST/SIMA/Surtax total matches when those columns exist in both reports
- Post-processing validation runs on the final transformed rows before workbook download:
  - DutiesHeader required fields: `Value for Duty`, `Duty`, and `GST` / `Gov. Sales Tax`; optional `HST`, `PST`, `SIMA`, and `Surtax` fields are checked for blanks and zeros when present.
  - DutiesItem required fields: `Quantity`, `Value for Duty`, `Duty`, `Value for Tax`, and `GST` / `Gov. Sales Tax`; optional `HST`, `PST`, `SIMA`, and `Surtax` fields are checked for blanks and zeros when present.
  - Blank/null cells and exact numeric zero values (including formatted/accounting zero) are warnings. Tiny nonzero values are not zero. This is a blank/zero check, not a general numeric-format validator.
  - Header `Duty` and `Gov. Sales Tax` zero warnings apply to LVS/PGA rows; CLVS rows are expected to have zero amounts in those fields and are excluded. Header `Value for Duty` remains checked for every row.
  - Findings show the output Excel row and resolved CCN; Item findings also show `Transaction Number` and `Line #`.
  - All findings remain visible. Counts and record labels distinguish uploaded report rows from generated Header rows. Expected CLVS zero exclusions are reported separately. Provenance is tracked separately through sorting and does not add workbook columns.
  - Missing validation columns are blocking errors before either download; findings themselves do not block downloads.

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
- Standalone Header/Item Analyzer item section display is currently removed from its report body.
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
- `public/dt-header-workflow.js`: shared D/T header/item normalization, brokerage automation, sorting, and summary helpers.
- `public/brokerage-rates.json`: client-side brokerage rate lookup for D/T Header automation.
- `views/index.ejs`: legacy minimal file, not primary UI.

## Key implementation assumptions
- Merger assumes first file contributes the merged header row.
- Merger currently skips first post-slice row only for file index 0 (first file), keeps row 0 for subsequent files.
- SFTP AC/AS start row for modifier is fixed at row 3 (index 2).
- D/T workflow header detection scans the first 10 rows and expects:
  - Header mode: `Transaction Number` and at least one of `CCN`, `Cargo Control Number`, or `Order Number`
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
