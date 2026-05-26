# 112-05240631 Final Output Observations

## Scope
This note records observed workbook behavior from:

- Source header: `Test files/112-05240631/CLVS_Report_Header_10133_017927245_260526112458466.xlsx`
- Source item: `Test files/112-05240631/CLVS_Report_Detail_10133_017927245_260526112458608.xlsx`
- Final header: `Test files/112-05240631/final output/RLBE_50_11205240631_PVG_YYZ_260427045723_DutiesHeader.xlsx`
- Final item: `Test files/112-05240631/final output/RLBE_50_11205240631_PVG_YYZ_260427045723_DutiesItem.xlsx`

The goal here is observation only, not a design decision.

## Input vs Final Layout

### Source files
- Both source workbooks place their category header row on Excel row `4`.
- Source header has `18` active columns and ends at `Inco Terms`.
- Source item has `16` active columns and ends at `Inco Terms`.
- Source item has no `CCN` column.

### Final output files
- Both final workbooks place their category header row on Excel row `5`.
- Rows `1` to `3` are metadata rows in both outputs.
- Final header has `25` columns in the sheet, although the meaningful business headers still stop at `Inco Terms`.
- Final item has `17` active columns and ends with a new `CCN` column.

## Metadata Rewrite

### Source metadata
- Header source:
  - `CLIENT:` = `RELIABLE LOGISTICS`
  - `RPT NAME:` = `WEEKLY HEADER REPORT`
  - `RPT DATE :` = `05/26/2026`
- Item source:
  - `CLIENT:` = `RELIABLE LOGISTICS`
  - `RPT NAME:` = `WEEKLY DETAIL REPORT`
  - `RPT DATE :` = `05/26/2026`

### Final metadata
- Header final:
  - `CLIENT:` = `SF EXPRESS`
  - `RPT NAME:` = `AWB# 112-05240631`
  - `RPT DATE :` = `4/28/2026`
- Item final:
  - `CLIENT:` = `SF EXPRESS`
  - `RPT NAME:` = `AWB# 112-05240631`
  - `RPT DATE :` = `4/28/2026`

Observed behavior:
- The final process rewrites metadata consistently across both header and item outputs.
- The final process also normalizes row-4 source layouts into row-5 output layouts.

## Header Output Behavior

This section was rechecked against the corrected workbook saved at:

- `Test files/112-05240631/final output/RLBE_50_11205240631_PVG_YYZ_260427045723_DutiesHeader.xlsx`

### Overall shape
- Source header starts with `74` source transaction-number rows.
- Final header contains `215` populated rows under the header.
- Final header contains `75` distinct transaction numbers.
- Final header contains `215` distinct CCN values.

### Pattern in generated rows
- Many appended rows use `CLVS` in column `A` under the `Transaction Number` header.
- In those appended rows:
  - Column `B` holds the CCN-like value.
  - Port/date fields are populated and repeated in the same pattern.
  - Column `J` contains the value-for-duty amount.
  - `Inco Terms` is `DDP`.
- Sample trailing rows look like:
  - `A = CLVS`
  - `B = SF5199209249395`
  - `J = 1.61`
  - `R/Inco Terms = DDP`

Observed implication:
- The final header output is not a simple one-row-per-transaction copy of the original header file.
- It behaves like a mixed workbook:
  - original-style rows remain present
  - many generated CLVS rows are appended afterward

## Corrected Business Rules Observed In Current Header File

### Exchange rate
- No data row now contains an exchange-rate value of `1`.
- The exchange-rate cells are currently displayed in accounting-style zero form as `$-`.

### Brokerage grouping by row type
- The corrected workbook currently contains exactly three brokerage-total groups, in this order:
  - `$0.71`
  - `$0.28`
  - `$0.02`
- The rows are already grouped from largest brokerage total to smallest.

### PGA rows
- There are `2` rows whose `CCN` starts with `8308`.
- All `8308...` rows currently have `Brokerage Total = $0.71`.

### LVS rows
- There are `72` rows whose `Transaction Number` starts with `LV`.
- All `LV...` rows currently have `Brokerage Total = $0.28`.

### CLVS rows
- There are `141` rows whose `Transaction Number` is `CLVS`.
- In the current corrected file, all `CLVS` rows have `Brokerage Total = $0.02`.
- User-stated process note:
  - CLVS brokerage should ultimately be derived from the client/value basis.
  - The exact basis for choosing the CLVS fee will be provided later.
- Important distinction:
  - the current workbook shows a single CLVS fee bucket of `$0.02`
  - the future rule basis is not yet fully specified in the repo documentation

### Date normalization
- `RPT DATE :` in metadata is `4/28/2026`.
- After normalizing formatting differences such as `4/28/2026` vs `04/28/2026`, all populated data rows use that same date for:
  - `Shipment Date`
  - `Arrival Date`
  - `Release Date`

### Numeric presentation
- After row generation and business-rule adjustments, the numeric range from `Value for Duty` through `Exchange Rate` is formatted as accounting.
- This is a presentation step applied after the data values have already been finalized.

## Item Output Behavior

### Overall shape
- Source item has `116` populated transaction-number rows.
- Final item also has `116` populated transaction-number rows.
- Final item adds one new column:
  - `CCN` as the last active column

### CCN propagation
- Final item `CCN` column is at Excel column `Q` (`17th` column).
- `116 / 116` item rows have `CCN` populated.
- Comparing final item rows against the final header `Transaction Number -> CCN` relationship shows:
  - `0` missing transaction numbers
  - `0` CCN mismatches

Observed behavior:
- The item workbook row count is preserved.
- The main structural change on the item side is the addition and fill of the `CCN` column.

## Concrete Transformation Summary
- Source header row `4` became final header row `5`.
- Source item row `4` became final item row `5`.
- Source metadata was overwritten in both files.
- Final item gained a `CCN` column and that column is fully filled.
- Final header expanded substantially and includes many appended `CLVS` rows.
- In the corrected header file:
  - `8308...` rows are grouped first at `$0.71`
  - `LV...` rows follow at `$0.28`
  - `CLVS` rows follow at `$0.02`
  - all three date columns align to `RPT DATE`
  - exchange-rate `1` values are no longer present
  - the numeric band from `Value for Duty` to `Exchange Rate` is presented in accounting format

## Output Naming Behavior
- Final export naming is based on the SFTP filename, not the original header/item filename.
- Header output naming behavior:
  - take the SFTP-based name
  - refresh the timestamp segment
  - end the file name with `_DutiesHeader`
- Item output naming behavior:
  - use the same SFTP-based naming pattern
  - end the file name with `_DutiesItem`
- Observed intent:
  - header and item outputs should remain a paired set derived from the same shipment/SFTP filename context

## Implementation-Relevant Notes
- Row-4 and row-5 header handling is required for this workflow.
- Metadata rewrite is part of the observed final-output workflow, not a separate manual prep step.
- Item `CCN` fill should be treated as a direct output transformation, not a formula left for the user.
- The final header output pattern includes generated rows whose first visible business value is `CLVS` in column `A`, so any future automation that reads the final header must account for that shape explicitly.
- Brokerage assignment is not only a formatting concern; it is part of the business classification of row groups.
- The current verified CLVS output is a single `$0.02` bucket, but the intended CLVS rule basis is still incomplete and must be finalized before implementing that part of automation.
- Final number formatting and final filename generation both happen at the end of the workflow and should be treated as required output steps, not optional polish.
