# D/T Header Brokerage Automation Design

## Summary

Extend the current `D/T Header File` workflow so the app not only rewrites metadata, normalizes row-4 layouts, updates header rows from SFTP, and optionally fills item `CCN`, but also applies the observed business rules that users currently finish manually in Excel:

- assign `Brokerage Total` from a client-rate JSON reference
- set date columns from `RPT DATE`
- zero out `Exchange Rate`
- sort header rows by brokerage bucket descending
- format the numeric band from `Value for Duty` through `Exchange Rate` as accounting
- generate a final report showing brokerage bucket counts plus header/item duty and GST totals

The workflow remains fully client-side and uses `public/brokerage-rates.json` as the editable reference source for brokerage mapping.

## Goals

- Preserve the current browser-only XLSX model.
- Keep the existing D/T page and `8308` analysis gate.
- Apply header business rules automatically after row generation.
- Use `CLIENT` in `A2` / metadata as the basis for client brokerage lookup.
- Support the current row-4 and row-5 header normalization behavior.
- Reuse the existing analyzer-style reporting concepts for the final D/T completion report.

## Non-Goals

- No backend or database.
- No admin UI for editing rates in this phase.
- No live rate download.
- No fuzzy client lookup beyond explicit normalization.
- No change to item row count or item business classification logic beyond `CCN` carryover and total comparison.

## Inputs

The workflow continues to accept:

- required `SFTP`
- required `DutiesHeader`
- optional `DutiesItem`
- required metadata:
  - `CLIENT`
  - `RPT NAME`
  - `RPT DATE`

It now also depends on:

- `public/brokerage-rates.json`

## Brokerage Rate Reference

The app uses [brokerage-rates.json](C:/Users/Benjamin/Desktop/WORK/IT/excel-merger/public/brokerage-rates.json:1) as the source of truth for brokerage mapping.

### JSON structure assumptions

The file exposes:

- `defaults`
  - baseline `pga`
  - baseline `lvs`
- `groups`
  - grouped client sets where only `clvs` differs
- `overrides`
  - explicit per-client `pga`, `lvs`, and `clvs`
- `clientRateLookup`
  - flattened direct lookup by client name

### Client matching rules

Client matching must:

- use the user-entered `CLIENT` value
- trim leading and trailing whitespace
- collapse repeated internal whitespace to a single space
- compare case-insensitively

The app should normalize the user value and the JSON keys the same way before lookup.

### Lookup precedence

1. Match against `clientRateLookup`
2. If no match, treat the client as unknown

No fallback group inference beyond the JSON contents.

## Unknown Client Rule

If the `CLIENT` value is not found in the JSON:

- do not stop the workflow
- do not synthesize fallback brokerage values into header rows
- leave brokerage blank for rows that depend on the client lookup
- surface the missing client / blank brokerage condition in the final report

Rows whose brokerage is determined directly by non-client rules may still be populated if the design later requires it, but in this phase the safest rule is:

- if the row needs the configured client rates, and the client is unknown, leave `Brokerage Total` blank

## Header Business Rules

These rules apply to the normalized and generated header output after metadata rewrite and after the existing SFTP-to-header row insertion logic.

### Classification rules

For each populated header data row:

1. If `CCN` starts with `8308`
   - classify row as `PGA`
   - set `Brokerage Total` from client `pga`
2. Else if `Transaction Number` starts with `LV`
   - classify row as `LVS`
   - set `Brokerage Total` from client `lvs`
3. Else if `Transaction Number` is exactly `CLVS`
   - classify row as `CLVS`
   - set `Brokerage Total` from client `clvs`
4. Else
   - leave `Brokerage Total` unchanged unless a future rule is defined

If the client is unknown:

- rows that would otherwise receive `pga`, `lvs`, or `clvs` from the JSON should keep `Brokerage Total` blank

### Date overwrite rule

For every populated header data row:

- `Shipment Date = RPT DATE`
- `Arrival Date = RPT DATE`
- `Release Date = RPT DATE`

The written display can preserve the current workbook-style formatting, but the logical source of truth is the user-entered `RPT DATE`.

### Exchange rate rule

For every populated header data row:

- set `Exchange Rate` to zero

Presentation may render as accounting zero such as `$-`.

### Sorting rule

After brokerage assignment:

- sort header data rows from largest brokerage total to smallest

Observed desired order with the current rate set:

1. `PGA`
2. `LVS`
3. `CLVS`
4. blank / unmatched brokerage rows last

Within a brokerage bucket, preserve current relative order where practical instead of introducing a second arbitrary sort key.

### Numeric formatting rule

After values are final and after sorting:

- apply accounting formatting from `Value for Duty` through `Exchange Rate`

This is the final presentation step for the header workbook.

## Item Rules

The item workflow remains the same base behavior:

- rewrite metadata
- normalize row-4 to row-5 when needed
- ensure `CCN` column exists
- fill item `CCN` from header `Transaction Number -> CCN`

No brokerage assignment is written into the item workbook in this phase.

## Output Naming

Final naming remains SFTP-based for both outputs.

### Header output

- derive name from the SFTP filename
- refresh or insert the 12-digit timestamp
- end with `_DutiesHeader`

### Item output

- derive name from the same SFTP filename pattern
- refresh or insert the 12-digit timestamp
- end with `_DutiesItem`

## Final Report

After the modify run completes, the D/T tool should show a final report in the same area currently used for completion status.

### Header section

Show:

- total `PGA` rows
- total `LVS` rows
- total `CLVS` rows
- total `Duty`
- total `GST`
- count of rows with blank `Brokerage Total`
- whether the `CLIENT` lookup matched the JSON or not

### Header vs item comparison

If an item file was provided, compare:

- header `Duty` total vs item `Duty` total
- header `GST` total vs item `GST` total

Show each comparison as:

- `Match`
- `Mismatch`

The report should reuse the analyzer’s numeric parsing logic so totals are consistent with existing app behavior.

### Report behavior without item file

If no item file is provided:

- still show the header summary
- omit the comparison section or mark it as not available

## Implementation Shape

Use option 1:

- put the new business logic into `public/dt-header-workflow.js`
- keep `public/script.js` focused on:
  - reading files
  - calling shared helpers
  - downloading workbooks
  - rendering the report

### Shared workflow additions

Recommended additions in `public/dt-header-workflow.js`:

- client-name normalization helper
- brokerage-rate lookup helper
- header column resolver
- header row classifier
- brokerage assignment pass
- date overwrite pass
- exchange-rate zeroing pass
- stable brokerage sort helper
- final D/T summary builder

### Script-side responsibilities

In `public/script.js`:

- load brokerage JSON once when the D/T module initializes or lazily before modify
- pass rate config plus metadata into the shared workflow
- keep download generation and final report rendering in the module

## Error Handling

Hard-stop errors:

- brokerage JSON cannot be loaded or parsed
- required header columns are missing from the normalized workbook
- required metadata values are missing
- existing `8308` analysis fails

Soft warnings:

- client not found in JSON
- resulting blank brokerage rows
- item totals mismatch against header totals

Unknown client is not a hard-stop condition.

## Testing

Use `Test files/112-05240631` as the main regression fixture for this phase.

Add regression coverage for:

1. Known client lookup from JSON:
   - `SF EXPRESS` resolves to `pga=0.71`, `lvs=0.28`, `clvs=0.0175`
2. Row-4 normalization still lands on row 5
3. `8308...` rows receive `PGA`
4. `LV...` rows receive `LVS`
5. `CLVS` rows receive client `CLVS`
6. All three header date columns are rewritten from `RPT DATE`
7. Exchange rate is zeroed
8. Brokerage buckets are sorted descending
9. Accounting formatting config still targets `Value for Duty -> Exchange Rate`
10. Item `CCN` fill still succeeds
11. Final report totals can compare header/item duty and GST
12. Unknown client leaves brokerage blank and is reflected in summary output
13. Output naming remains SFTP-based with `_DutiesHeader` / `_DutiesItem`

## Acceptance Criteria

- The app automatically applies observed header business rules after modify.
- Brokerage totals are assigned from `brokerage-rates.json` using `CLIENT`.
- Unknown clients do not crash the run and leave brokerage blank.
- Header rows are sorted by brokerage total descending.
- Header numeric fields from `Value for Duty` through `Exchange Rate` are presented as accounting.
- Final completion view shows PGA/LVS/CLVS counts and header/item duty-GST comparison.
- Existing row-4 normalization, metadata rewrite, 8308 analysis gate, and item `CCN` fill all still work.
