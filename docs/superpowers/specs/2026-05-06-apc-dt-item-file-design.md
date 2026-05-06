# APC D/T Item File Design

## Summary

Add a new client-side conversion tool, `APC D/T Item File`, that takes:

- one SFTP workbook
- one `_DutiesItem` workbook

The tool appends new `CLVS` rows into the item workbook by comparing existing item-file `CCN` values against SFTP `Reliable_tracking` values.

This feature follows the same high-level principle as the existing `D/T Header File` tool:

1. derive a reference set from the target workbook
2. remove matching rows from the SFTP source
3. append only the remaining rows

The item-file version differs from the header-file version in three important ways:

- it uses item `CCN` instead of header `Order Number`
- it normalizes existing item `CCN` values by stripping leading `8308` only when present
- it appends row-level item records, not one synthesized row per unique key

## Goals

- Preserve the current fully client-side XLSX processing model.
- Reuse the existing module pattern in `public/script.js`.
- Append only new `CLVS` rows into the `_DutiesItem` workbook.
- Add six new output columns after `CCN`, populated only for newly appended rows.
- Keep existing item rows intact except for widened sheet structure when new columns are added.

## Non-Goals

- No backend processing.
- No attempt to reconcile or update existing item rows.
- No aggregation or deduplication of unmatched SFTP rows beyond the compare rule.
- No change to existing `D/T Header File` behavior.

## Inputs

### Target workbook

- `_DutiesItem` workbook
- first sheet only
- row 5 (index `4`) is the header row
- data begins on row 6 (index `5`)

Expected existing headers in row 5:

- `Transaction Number`
- `Goods Description`
- `Line #`
- `Country of Origin`
- `Tariff Treatment`
- `Part Number`
- `Quantity`
- `Port #`
- `Vendor Name`
- `Value for Duty`
- `HS #`
- `Duty Rate`
- `Duty`
- `Value for Tax`
- `Gov. Sales Tax`
- `Inco Terms`
- `CCN`

### Source workbook

- SFTP workbook
- first sheet only
- header row is row 1 (index `0`)
- data begins on row 3 (index `2`)

Required SFTP columns:

- `Reliable_tracking`
- `Goods_Description`
- `Package_no`
- `Country_of_origin`
- `Product_part`
- `Quantity`
- `CBSA_Port_of_Release`
- `Seller_name`
- `Total_value_of_parcel`
- `HS_code`
- `Inco_term`
- `Buyer_name`
- `Buyer_address`
- `Buyer_city`
- `Buyer_postal_code`
- `Buyer_province`
- `Order_number`

## Key Compare Rule

Build a reference set from existing item-file `CCN` values:

- if `CCN` starts with `8308`, remove that prefix
- otherwise keep the `CCN` unchanged

Examples:

- `8308APCP0001284765` -> `APCP0001284765`
- `APC1709600226264` -> `APC1709600226264`

For each SFTP data row:

- read `Reliable_tracking`
- trim whitespace
- if `Reliable_tracking` is blank, skip the row
- if the value exists in the normalized item `CCN` reference set, skip the row
- otherwise keep the row for append

This means:

- rows matching existing `PGA` keys are removed because the target-side `8308` prefix is stripped
- rows matching existing `LVS` keys are removed because non-`8308` values are kept as-is
- only new `CLVS` rows remain for append

## Output Shape

### Existing columns

New appended rows populate the item-file columns as follows:

- `Transaction Number` = literal `CLVS`
- `Goods Description` <- SFTP `Goods_Description`
- `Line #` <- SFTP `Package_no`
- `Country of Origin` <- SFTP `Country_of_origin`
- `Tariff Treatment` = blank
- `Part Number` <- SFTP `Product_part`
- `Quantity` <- SFTP `Quantity`
- `Port #` <- SFTP `CBSA_Port_of_Release`
- `Vendor Name` <- SFTP `Seller_name`
- `Value for Duty` <- SFTP `Total_value_of_parcel`
- `HS #` <- SFTP `HS_code`
- `Duty Rate` = `0`
- `Duty` = `0`
- `Value for Tax` <- SFTP `Total_value_of_parcel`
- `Gov. Sales Tax` = `0`
- `Inco Terms` <- SFTP `Inco_term`
- `CCN` <- raw SFTP `Reliable_tracking`

### Added columns

Append six new columns immediately after `CCN` in this exact order:

1. `Buyer Name`
2. `Buyer Address`
3. `Buyer City`
4. `Buyer Postal Code`
5. `Buyer Province`
6. `Order Number`

Population rules:

- existing item rows remain blank for all six new columns
- newly appended `CLVS` rows populate them from SFTP:
  - `Buyer Name` <- `Buyer_name`
  - `Buyer Address` <- `Buyer_address`
  - `Buyer City` <- `Buyer_city`
  - `Buyer Postal Code` <- `Buyer_postal_code`
  - `Buyer Province` <- `Buyer_province`
  - `Order Number` <- `Order_number`

If the target workbook already contains these six headers immediately after `CCN` in the same order, reuse them and do not append duplicate columns.

## Append Behavior

- Preserve rows 1 to 5 exactly.
- Find the last non-empty row in the target workbook.
- Append new `CLVS` rows after the last non-empty existing row, before trailing blank rows.
- Do not modify existing item rows other than widening row length so the new columns exist in the sheet.
- Do not deduplicate the remaining SFTP rows. If multiple unmatched SFTP rows share the same `Reliable_tracking`, append each remaining row.

## Formatting Rules

For newly appended rows, write numeric cells as numeric worksheet values with `General` format when parseable:

- `Quantity`
- `Value for Duty`
- `Duty Rate`
- `Duty`
- `Value for Tax`
- `Gov. Sales Tax`

Text cells should be written as text values. Text sourced from SFTP should be sanitized consistently with existing workbook-writing protections used elsewhere in the app to avoid Excel formula injection when needed.

## Filename Rule

Generate output from the target filename using the same timestamp-refresh pattern as the existing header tool, but for `_DutiesItem`:

- if the filename already contains a 12-digit timestamp immediately before `_DutiesItem`, replace that timestamp
- otherwise insert the generated 12-digit timestamp before `_DutiesItem`
- if `_DutiesItem` is not present, append the timestamp before the extension

## UI / UX Behavior

The existing UI shell for `APC D/T Item File` remains the entry point.

Behavior to add:

- enable the primary button only when both files are selected
- support click-to-select and drag/drop on both zones
- display selected filenames under each zone
- clicking run performs conversion and downloads the new workbook
- after successful generation, show a reset action similar to the existing tools

Validation failures should use the app's existing alert-based pattern.

## Error Handling

Alert and abort when:

- either input file is missing
- a required SFTP column is missing
- the item file does not expose the expected `CCN` column in row 5
- no sheet data can be read
- workbook parsing fails

If there are zero unmatched SFTP rows:

- still generate an output workbook with widened columns if needed
- do not append any new data rows
- download proceeds normally

## Implementation Plan Shape

Implementation should stay isolated to the new tool:

- `public/index.html`
  - extend the new tool section with a reset button container if needed
- `public/script.js`
  - add a dedicated IIFE module for `APC D/T Item File`
  - reuse small helpers where practical without changing existing tool behavior
- `public/styles.css`
  - extend existing button/reset styling only as needed

Do not refactor the existing header converter as part of this change.

## Testing

Manual verification should cover:

1. Both files selected, run succeeds, file downloads.
2. Existing `8308...` item `CCN` values correctly match SFTP `Reliable_tracking` after prefix strip.
3. Existing non-`8308` item `CCN` values match SFTP `Reliable_tracking` without modification.
4. Remaining unmatched SFTP rows are appended.
5. New columns are appended after `CCN` in the specified order.
6. Existing rows remain blank in the new columns.
7. New rows use `CLVS` in `Transaction Number`.
8. New rows use raw `Reliable_tracking` in `CCN`.
9. Numeric cells in new rows are exported as numeric values.
10. Output filename follows the `_DutiesItem` timestamp rule.
