# D/T Header and DutiesItem CCN Automation Design

## Summary

Extend the existing `D/T Header File` tool so one workflow can:

- accept `SFTP`, `DutiesHeader`, and optional `DutiesItem` workbooks
- collect `CLIENT`, `RPT NAME`, and `RPT DATE` from UI inputs
- overwrite rows `1-3` in generated outputs with those user-provided values
- normalize `DutiesHeader` and `DutiesItem` layouts when their column headers are on row `4`
- keep the current `8308` analysis gate for the header flow
- generate the updated `DutiesHeader` workbook
- optionally generate a paired `DutiesItem` workbook with direct `CCN` values written from the header workbook

The goal is to eliminate the manual pre-processing steps that users currently do in Excel before running the tool. The user should be able to drop the files into the app, let the app do the cleanup and carryover work, then review the downloaded outputs.

## Goals

- Preserve the current fully client-side XLSX processing model.
- Keep the `D/T Header File` page as the single entry point for this workflow.
- Support both row-`4` and row-`5` header layouts for `DutiesHeader` and `DutiesItem`.
- Let the user provide `CLIENT`, `RPT NAME`, and `RPT DATE` once and apply them to both outputs.
- Write `RPT DATE` as text so Excel does not auto-convert the stored value unexpectedly.
- Continue using the existing `8308` analysis before allowing header modification.
- When a `DutiesItem` workbook is supplied, write final `CCN` values directly instead of formulas.

## Non-Goals

- No backend processing.
- No new standalone page or tool for item processing.
- No fallback matching from `Order Number`, SFTP tracking, or other columns for item `CCN` fill.
- No fuzzy matching of `Transaction Number`.
- No change to SFTP column positions in this feature.

## User Workflow

The `D/T Header File` section becomes a three-file workflow plus metadata inputs:

1. User enters:
   - `CLIENT`
   - `RPT NAME`
   - `RPT DATE`
2. User drops:
   - one SFTP workbook
   - one `DutiesHeader` workbook
   - optionally one `DutiesItem` workbook
3. User clicks `Analyze 8308 Values`.
4. The app runs the existing `8308` value-for-duty comparison and shows the report.
5. User clicks `Proceed with Modify`.
6. The app:
   - rewrites rows `1-3` in the generated header output
   - normalizes header row placement when needed after the metadata rewrite
   - performs the existing header insert/update logic
   - if `DutiesItem` was provided, rewrites rows `1-3`, normalizes row placement when needed after the metadata rewrite, adds or reuses `CCN`, fills `CCN` values directly from the header workbook, and downloads the item output
7. User reviews the downloaded file or files after generation.

Compatibility rules:

- `SFTP` and `DutiesHeader` remain required.
- `DutiesItem` is optional so the current header-only workflow still works.
- If only header is provided, only the updated header workbook is downloaded.
- If header and item are provided, both outputs are downloaded in one run.

## Input Metadata Rules

The UI will collect three user inputs for the output metadata:

- `CLIENT`
- `RPT NAME`
- `RPT DATE`

These inputs apply to both generated outputs.

### Row 1 to 3 write rules

For both `DutiesHeader` and `DutiesItem` outputs:

- row `1`, column `A` = `CLIENT:`
- row `1`, column `B` = user-entered `CLIENT`
- row `2`, column `A` = `RPT NAME:`
- row `2`, column `B` = user-entered `RPT NAME`
- row `3`, column `A` = `RPT DATE :`
- row `3`, column `B` = user-entered `RPT DATE`

The app should explicitly write these values as text cells.

### Date handling

`RPT DATE` must be stored as text, not as an Excel date serial.

Rules:

- preserve exactly what the user typed after trimming outer whitespace
- do not parse or reformat it into an Excel date value
- do not auto-convert `05/27/2026` into a date cell

This avoids the current manual step where users pre-fill the workbook and rely on Excel formatting behavior before running the tool.

## Header Row Detection and Normalization

`DutiesHeader` and `DutiesItem` currently appear in two supported layouts:

- headers on row `5`
- headers on row `4`

The tool must detect the actual header row instead of assuming row `5`.

### Detection rule

Scan the first `6` rows and choose the row that best matches the expected labels.

For `DutiesHeader`, anchor on:

- `Transaction Number`
- `CCN`

For `DutiesItem`, anchor on:

- `Transaction Number`
- `Goods Description`

The detected header row determines the data start row for processing.

### Normalization rule

If the detected header row is row `4`:

- insert one blank row before the header row in the output workbook
- resulting output header row becomes row `5`
- resulting data start row becomes row `6`

If the detected header row is already row `5`:

- do not insert any row

Normalization order:

1. load workbook rows
2. detect whether header row is `4` or `5`
3. write rows `1-3` metadata values
4. if row `4`, insert one blank row before the detected header row
5. then continue with the rest of processing

This mirrors the current manual cleanup step while keeping the output consistent.

## D/T Header Processing Rules

The existing `D/T Header File` logic remains the baseline behavior.

### Existing rules to preserve

- use SFTP `AC` as the source CCN candidate
- use SFTP `AS` as the source value for column `J`
- deduplicate source `AC` values before insert
- build the target reference set from header `CCN`
- remove leading `8308` only when it appears at the start of the existing target `CCN`
- append only new CCNs
- create synthesized rows with:
  - `A = CLVS`
  - `B = source AC`
  - `H = source AC`
  - `J = source AS`
  - `C..F` copied from the last non-empty existing target row
  - `K..Q = 0`
  - `R = DDP`
- force `Value for Duty` through `Exchange Rate` cells to numeric `General` format when parseable
- keep the `8308` analysis report and explicit proceed step

### Required change

The header module must stop assuming:

- header row index `0`
- target data starts at row index `5`

Instead it must:

- detect the actual header row
- normalize row `4` to row `5` when needed
- derive the target data start row from the normalized header row

After normalization, the downstream logic can continue using the expected row-`5` structure.

## DutiesItem CCN Fill Rules

If a `DutiesItem` workbook is provided, the app must generate an updated item workbook as part of the same run.

### Source of truth

The item `CCN` values come from the header workbook, not from SFTP.

Lookup map:

- key = `Transaction Number` from `DutiesHeader`
- value = `CCN` from `DutiesHeader`

Matching rule:

- trim whitespace on both sides before comparing
- exact match only

### Item processing steps

1. Load the item workbook.
2. Detect whether its headers are on row `4` or row `5`.
3. If on row `4`, insert one blank row so the output header is row `5`.
4. Write rows `1-3` from the user inputs as text.
5. Resolve item column indexes from the normalized header row.
6. Build the header lookup map from the normalized header workbook data.
7. Fill the item `CCN` values directly.

### Column rules

Required item column:

- `Transaction Number`

Optional existing item column:

- `CCN`

Behavior:

- if `CCN` already exists in the item header row, reuse it
- if `CCN` does not exist, append a new `CCN` column at the end of the used header columns
- write the `CCN` header label as text

### Row-level fill rule

For each item data row:

- read `Transaction Number`
- if blank, write blank `CCN`
- if the transaction number exists in the header lookup map, write the mapped `CCN`
- if not found, leave `CCN` blank

No formula should be written to the workbook.

The manual Excel step this replaces is conceptually:

`VLOOKUP(item Transaction Number, header Transaction Number:CCN, 2, 0)`

but the app must resolve it directly and write final values.

## Unmatched Item Rows

If some item rows do not find a matching header `Transaction Number`:

- leave `CCN` blank
- still generate and download the item workbook
- report the unmatched count to the user in the completion message or result area

This is not a fatal error for the first version.

## Output File Naming

### DutiesHeader output

Keep the current naming behavior:

- if the filename already contains a 12-digit timestamp immediately before `_DutiesHeader`, replace that timestamp
- otherwise insert a generated 12-digit timestamp before `_DutiesHeader`
- if `_DutiesHeader` is not present, append the timestamp before the extension

### DutiesItem output

Apply the same pattern for item files:

- if the filename already contains a 12-digit timestamp immediately before `_DutiesItem`, replace that timestamp
- otherwise insert a generated 12-digit timestamp before `_DutiesItem`
- if `_DutiesItem` is not present, append the timestamp before the extension

If the input item filename is in a CLVS report style instead of `_DutiesItem`, preserve the original basename and append a timestamp before `.xlsx`.

## UI and UX Changes

The existing `D/T Header File` section remains the entry point.

Add:

- a third drop zone for `DutiesItem`
- three text inputs for:
  - `CLIENT`
  - `RPT NAME`
  - `RPT DATE`

Behavior:

- `Analyze 8308 Values` remains available when `SFTP` and `DutiesHeader` are selected and metadata inputs are present
- `Proceed with Modify` continues to appear only after analysis runs
- `DutiesItem` is optional
- the reset action clears:
  - the three metadata inputs
  - all selected files
  - the analysis report
  - the proceed button state

User experience goal:

- no need to open files beforehand
- no need to manually add a blank row
- no need to manually prefill rows `1-3`
- no need to add a formula column in Excel

## Error Handling

Alert and abort the run when:

- SFTP is missing
- `DutiesHeader` is missing
- any metadata input is blank
- the header workbook cannot be read
- the SFTP workbook cannot be read
- the detected `DutiesHeader` layout does not expose required columns
- `8308` analysis fails

If `DutiesItem` is provided, alert and abort the item-generation part when:

- the item workbook cannot be read
- the item header row cannot be detected
- the normalized item header row does not expose `Transaction Number`
- the normalized header workbook does not expose `Transaction Number` and `CCN` needed for the lookup map

For the combined run:

- header generation remains the primary workflow
- if an item file is supplied and item processing fails, the run should stop and report the item-processing error clearly before download completion is claimed

## Implementation Shape

Implementation stays in the existing D/T Header module:

- [public/index.html](C:/Users/Benjamin/Desktop/WORK/IT/excel-merger/public/index.html:108)
- [public/script.js](C:/Users/Benjamin/Desktop/WORK/IT/excel-merger/public/script.js:535)
- [public/styles.css](C:/Users/Benjamin/Desktop/WORK/IT/excel-merger/public/styles.css:613)

Recommended helper additions inside the current module:

- header-row detector for header and item workbooks
- row-`4` to row-`5` normalizer
- metadata writer for rows `1-3`
- column resolver by label
- header lookup-map builder
- duties-item workbook generator
- timestamped output-name generator reusable for both header and item

Keep the current module as one user-facing workflow, but isolate new logic into small helpers so the existing header behavior stays readable.

## Testing

Manual and regression coverage should include:

1. Existing row-`5` header-only workflow still works.
2. `8308` analysis still runs before modify.
3. Header output rows `1-3` are overwritten from user input.
4. Item output rows `1-3` are overwritten from user input.
5. `RPT DATE` is written as text, not an Excel date value.
6. Row-`4` CLVS samples from `Test files\\112-05240631` normalize to header row `5`.
7. Row-`5` APC-style samples remain unchanged structurally.
8. Item output appends `CCN` when missing.
9. Item output reuses `CCN` when already present.
10. Known `Transaction Number` values map to the expected header `CCN`.
11. Unmatched item rows keep blank `CCN`.
12. Two outputs are downloaded when all three files are provided.
13. One output is downloaded when only SFTP and header are provided.

Regression data to use:

- row-`4` CLVS samples in `Test files\\112-05240631`
- existing row-`5` header and item samples in `Test files\\83082142460`

## Acceptance Criteria

- A user can supply metadata and files without opening Excel first.
- The app rewrites rows `1-3` for both outputs from UI input.
- The app supports both row-`4` and row-`5` header layouts.
- Row-`4` header layouts are normalized to row `5` in the output.
- Existing `8308` header analysis still gates the modify action.
- Header processing still appends only new CCNs from SFTP.
- When an item file is supplied, the app writes direct `CCN` values from header `Transaction Number -> CCN` mapping.
- The user only needs to review the downloaded outputs after generation.
