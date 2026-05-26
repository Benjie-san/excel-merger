# D/T Header Brokerage Automation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Automate D/T header brokerage assignment, date normalization, exchange-rate zeroing, brokerage sorting, accounting formatting, and final summary reporting using `public/brokerage-rates.json`, while preserving the existing 8308 analysis gate and item `CCN` carryover workflow.

**Architecture:** Keep all D/T business transformations in `public/dt-header-workflow.js` and leave `public/script.js` as orchestration plus download/report rendering. Extend `scripts/regression-smoke.js` first so the new brokerage behavior is driven by failing regression checks against `Test files/112-05240631`, then implement the shared helpers, then wire the UI and reporting.

**Tech Stack:** Plain browser JavaScript, SheetJS `xlsx`, static `public/*.json` assets, Node-based regression harness in `scripts/regression-smoke.js`.

---

## File Structure

- Modify: `public/dt-header-workflow.js`
  - Shared D/T business rules, client normalization, JSON rate lookup, header transforms, sort helpers, summary generation.
- Modify: `public/script.js`
  - D/T module orchestration only: load rate JSON, call shared workflow, download workbooks, render final report.
- Modify: `scripts/regression-smoke.js`
  - New red/green coverage for client brokerage mapping, unknown-client blanks, header transforms, and final summary expectations.
- Reference only: `public/brokerage-rates.json`
  - Existing static brokerage source of truth; do not reshape unless implementation reveals a concrete bug.
- Reference only: `Test files/112-05240631/*`
  - Primary regression fixtures for row-4 workflow and corrected final-output observations.
- Optional doc touch if implementation changes assumptions: `documents/PROJECT_CONTEXT.md`

### Task 1: Extend Regression Coverage First

**Files:**
- Modify: `scripts/regression-smoke.js`
- Reference: `public/dt-header-workflow.js`
- Reference: `public/brokerage-rates.json`

- [ ] **Step 1: Add fixture loading for brokerage JSON and current 112-05240631 expectations**

Add near the D/T regression helpers:

```js
function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function normalizeClientName(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
}
```

Add fixture references inside `runDtHeaderWorkflowRegression`:

```js
const brokerageRatesPath = path.join(process.cwd(), "public", "brokerage-rates.json");
const brokerageRates = readJson(brokerageRatesPath);
const finalHeaderObservedPath = path.join(rootDir, "112-05240631", "final output", "RLBE_50_11205240631_PVG_YYZ_260427045723_DutiesHeader.xlsx");
```

- [ ] **Step 2: Write failing regression assertions for the new shared exports**

Add export checks before execution:

```js
const {
  detectHeaderRowIndex,
  prepareHeaderRowsForModify,
  prepareItemRowsWithCcn,
  applyBrokerageAutomation,
  summarizeDtOutputs
} = workflow;

if (typeof applyBrokerageAutomation !== "function") {
  issues.push("Missing export: applyBrokerageAutomation");
}
if (typeof summarizeDtOutputs !== "function") {
  issues.push("Missing export: summarizeDtOutputs");
}
```

- [ ] **Step 3: Add a failing brokerage automation test using `SF EXPRESS`**

Append a red test block after existing metadata/item-CCN checks:

```js
let transformed;
try {
  transformed = applyBrokerageAutomation({
    sourceRows: readFirstSheetRows(path.join(rootDir, "112-05240631", "RLBE_50_11205240631_PVG_YYZ_260427045723.xlsx")).rows,
    preparedHeader: actualHeaderResult,
    metadata: {
      client: "SF EXPRESS",
      reportName: "AWB# 112-05240631",
      reportDate: "4/28/2026"
    },
    brokerageRates
  });
} catch (err) {
  issues.push(`applyBrokerageAutomation threw: ${err.message}`);
}
```

Add assertions:

```js
if (!transformed || !Array.isArray(transformed.rows)) {
  issues.push("applyBrokerageAutomation should return an object with rows.");
}
if (!transformed || !transformed.summary) {
  issues.push("applyBrokerageAutomation should return a summary object.");
}
```

- [ ] **Step 4: Assert the transformed header matches key observed business rules**

Add focused checks:

```js
if (transformed.summary.clientMatched !== true) {
  issues.push(`Expected known client match for SF EXPRESS, got ${JSON.stringify(transformed.summary.clientMatched)}`);
}
if (transformed.summary.counts.pga !== 2) {
  issues.push(`Expected PGA count 2, got ${JSON.stringify(transformed.summary.counts.pga)}`);
}
if (transformed.summary.counts.lvs !== 72) {
  issues.push(`Expected LVS count 72, got ${JSON.stringify(transformed.summary.counts.lvs)}`);
}
if (transformed.summary.counts.clvs <= 0) {
  issues.push(`Expected positive CLVS count, got ${JSON.stringify(transformed.summary.counts.clvs)}`);
}
if (transformed.summary.blankBrokerageCount !== 0) {
  issues.push(`Expected no blank brokerage rows for SF EXPRESS, got ${JSON.stringify(transformed.summary.blankBrokerageCount)}`);
}
```

Add row-level samples:

```js
const sampleRows = transformed.rows.slice(transformed.headerRowIndex + 1, transformed.headerRowIndex + 6);
if (!String((sampleRows[0] || [])[1] || "").trim().startsWith("8308")) {
  issues.push("Expected first transformed data row to be a PGA 8308 row after sort.");
}
```

- [ ] **Step 5: Add failing checks for date overwrite, exchange-rate zeroing, and unknown-client blanks**

Use resolved column indexes:

```js
const transformedHeader = transformed.rows[transformed.headerRowIndex] || [];
const shipmentIdx = findHeaderColumnIndex(transformedHeader, "Shipment Date");
const arrivalIdx = findHeaderColumnIndex(transformedHeader, "Arrival Date");
const releaseIdx = findHeaderColumnIndex(transformedHeader, "Release Date");
const exchangeIdx = findHeaderColumnIndex(transformedHeader, "Exchange Rate");
const brokerageIdx = findHeaderColumnIndex(transformedHeader, "Brokerage Total");
```

Add checks:

```js
for (let r = transformed.headerRowIndex + 1; r < Math.min(transformed.rows.length, transformed.headerRowIndex + 15); r++) {
  const row = transformed.rows[r] || [];
  if (isEmptyRow(row)) continue;
  if (String(row[shipmentIdx] || "").trim() !== "4/28/2026") issues.push(`Shipment date mismatch at row ${r + 1}`);
  if (String(row[arrivalIdx] || "").trim() !== "4/28/2026") issues.push(`Arrival date mismatch at row ${r + 1}`);
  if (String(row[releaseIdx] || "").trim() !== "4/28/2026") issues.push(`Release date mismatch at row ${r + 1}`);
  if (parseNumberZero(row[exchangeIdx]) !== 0) issues.push(`Exchange rate should be zero at row ${r + 1}`);
}
```

Add unknown-client test:

```js
const unknownClientResult = applyBrokerageAutomation({
  sourceRows: readFirstSheetRows(path.join(rootDir, "112-05240631", "RLBE_50_11205240631_PVG_YYZ_260427045723.xlsx")).rows,
  preparedHeader: actualHeaderResult,
  metadata: { client: "UNKNOWN CLIENT", reportName: "AWB# 112-05240631", reportDate: "4/28/2026" },
  brokerageRates
});
if (unknownClientResult.summary.clientMatched !== false) {
  issues.push("Unknown client should report clientMatched=false.");
}
if (unknownClientResult.summary.blankBrokerageCount <= 0) {
  issues.push("Unknown client should leave brokerage blank for at least one row.");
}
```

- [ ] **Step 6: Add failing summary comparison checks for header vs item totals**

After preparing item rows:

```js
const summary = summarizeDtOutputs({
  headerRows: transformed.rows,
  itemRows: actualItemRows
});

if (!summary || !summary.header || !summary.compare) {
  issues.push("summarizeDtOutputs should return header/item/compare sections.");
}
if (typeof summary.compare.dutyMatch !== "boolean") {
  issues.push(`Expected boolean dutyMatch, got ${typeof summary.compare.dutyMatch}`);
}
if (typeof summary.compare.gstMatch !== "boolean") {
  issues.push(`Expected boolean gstMatch, got ${typeof summary.compare.gstMatch}`);
}
```

- [ ] **Step 7: Run regression to verify the new assertions fail**

Run: `node scripts/regression-smoke.js`

Expected:
- `D/T Header Workflow: FAIL`
- missing-export and/or behavior failures referencing `applyBrokerageAutomation`, `summarizeDtOutputs`, brokerage counts, or unknown-client blanks

- [ ] **Step 8: Commit the failing regression harness**

```bash
git add scripts/regression-smoke.js
git commit -m "test(regression): cover dt brokerage automation rules"
```

### Task 2: Implement Shared Brokerage Automation In `dt-header-workflow.js`

**Files:**
- Modify: `public/dt-header-workflow.js`
- Reference: `public/brokerage-rates.json`
- Test: `scripts/regression-smoke.js`

- [ ] **Step 1: Add client normalization and rate lookup helpers**

Add near existing normalization helpers:

```js
function normalizeClientName(value) {
  return normalizeCell(value).replace(/\s+/g, " ").toLowerCase();
}

function lookupClientRates(brokerageRates, clientName) {
  var normalizedClient = normalizeClientName(clientName);
  var lookup = brokerageRates && brokerageRates.clientRateLookup ? brokerageRates.clientRateLookup : {};
  var keys = Object.keys(lookup);
  for (var i = 0; i < keys.length; i++) {
    if (normalizeClientName(keys[i]) === normalizedClient) {
      return {
        matched: true,
        clientKey: keys[i],
        rates: lookup[keys[i]]
      };
    }
  }
  return {
    matched: false,
    clientKey: "",
    rates: null
  };
}
```

- [ ] **Step 2: Add header-column helpers and money/date utilities**

Add focused helpers:

```js
function parseNumber(value) {
  var s = normalizeCell(value).replace(/[$,]/g, "");
  if (!s || s === "-" || s === "--") return null;
  var num = parseFloat(s);
  return isNaN(num) ? null : num;
}

function formatMoney(value) {
  if (value === null || value === undefined || isNaN(value)) return "";
  return value;
}

function normalizeDateText(value) {
  return normalizeCell(value);
}
```

Add a header resolver:

```js
function resolveHeaderColumns(headerRow) {
  return {
    transactionNumber: findColumnIndex(headerRow, "Transaction Number"),
    ccn: findColumnIndex(headerRow, "CCN"),
    shipmentDate: findColumnIndex(headerRow, "Shipment Date"),
    arrivalDate: findColumnIndex(headerRow, "Arrival Date"),
    releaseDate: findColumnIndex(headerRow, "Release Date"),
    valueForDuty: findColumnIndex(headerRow, "Value for Duty"),
    duty: findColumnIndex(headerRow, "Duty"),
    gst: findColumnIndex(headerRow, "Gov. Sales Tax"),
    brokerageTotal: findColumnIndex(headerRow, "Brokerage Total"),
    exchangeRate: findColumnIndex(headerRow, "Exchange Rate")
  };
}
```

- [ ] **Step 3: Extract the existing header-row insertion logic from `script.js` into the shared workflow**

Add a shared function that receives normalized header rows and source rows:

```js
function insertMissingHeaderRows(options) {
  // move the current buildModifiedHeaderRows_MOD row-generation rules here
}
```

It should return:

```js
{
  rows: finalRows,
  headerRowIndex: headerRowIndex,
  insertedCount: insertedRows.length
}
```

- [ ] **Step 4: Implement the brokerage automation pass**

Add:

```js
function applyBrokerageAutomation(options) {
  var preparedHeader = normalizePreparedHeaderInput(options && options.preparedHeader);
  var metadata = options && options.metadata ? options.metadata : {};
  var brokerageRates = options && options.brokerageRates ? options.brokerageRates : null;
  var rateLookup = lookupClientRates(brokerageRates, metadata.client);
  var rows = cloneRows(preparedHeader.rows);
  var headerRowIndex = preparedHeader.headerRowIndex;
  var columns = resolveHeaderColumns(rows[headerRowIndex] || []);
  // validate required columns
  // assign PGA/LVS/CLVS brokerage
  // overwrite date columns from metadata.reportDate
  // zero exchange rate
  // stable sort data rows by brokerage desc, blanks last
  // return summary metadata
}
```

Classification rule implementation:

```js
if (ccn.indexOf("8308") === 0) {
  classification = "PGA";
  brokerageValue = rateLookup.matched ? rateLookup.rates.pga : null;
} else if (transaction.indexOf("LV") === 0) {
  classification = "LVS";
  brokerageValue = rateLookup.matched ? rateLookup.rates.lvs : null;
} else if (transaction === "CLVS") {
  classification = "CLVS";
  brokerageValue = rateLookup.matched ? rateLookup.rates.clvs : null;
}
```

- [ ] **Step 5: Implement the final summary builder**

Add:

```js
function summarizeDtOutputs(options) {
  var headerSummary = summarizeHeaderRows(options && options.headerRows);
  var itemSummary = options && options.itemRows ? summarizeItemRows(options.itemRows) : null;
  return {
    header: headerSummary,
    item: itemSummary,
    compare: {
      dutyMatch: itemSummary ? Math.abs(headerSummary.totalDutyValue - itemSummary.totalDutyValue) <= 0.0001 : null,
      gstMatch: itemSummary ? Math.abs(headerSummary.totalGstValue - itemSummary.totalGstValue) <= 0.0001 : null
    }
  };
}
```

Use analyzer-style bucket counting with the exact brokerage values coming from classified rows rather than free-form heuristics.

- [ ] **Step 6: Export the new shared functions**

Append to the returned API:

```js
return {
  detectHeaderRowIndex: detectHeaderRowIndex,
  prepareHeaderRowsForModify: prepareHeaderRowsForModify,
  prepareItemRowsWithCcn: prepareItemRowsWithCcn,
  insertMissingHeaderRows: insertMissingHeaderRows,
  applyBrokerageAutomation: applyBrokerageAutomation,
  summarizeDtOutputs: summarizeDtOutputs
};
```

- [ ] **Step 7: Run regression to verify shared logic now passes the new D/T checks**

Run: `node scripts/regression-smoke.js`

Expected:
- `D/T Header Workflow: PASS`
- other existing sections remain `PASS`

- [ ] **Step 8: Commit the shared workflow implementation**

```bash
git add public/dt-header-workflow.js
git commit -m "feat(workflow): add dt brokerage automation helpers"
```

### Task 3: Wire The D/T Module To The Shared Workflow And JSON

**Files:**
- Modify: `public/script.js`
- Reference: `public/brokerage-rates.json`
- Test: `scripts/regression-smoke.js`

- [ ] **Step 1: Add a lazy JSON loader inside the D/T module**

In the `ExcelModifyModule` IIFE, add:

```js
let brokerageRatesPromise = null;

async function loadBrokerageRates_MOD() {
  if (!brokerageRatesPromise) {
    brokerageRatesPromise = fetch("brokerage-rates.json").then(async (response) => {
      if (!response.ok) {
        throw new Error(`Failed to load brokerage rates: ${response.status}`);
      }
      return response.json();
    });
  }
  return brokerageRatesPromise;
}
```

- [ ] **Step 2: Replace local header-row build logic with the shared insertion + automation pipeline**

Inside `runModifyWorkflow_MOD`, replace the old path with:

```js
const brokerageRates = await loadBrokerageRates_MOD();
const sourceRows = await readExcelFile_MOD(sourceFile);
const targetRows = await readExcelFile_MOD(targetFile);
const preparedHeader = workflow.prepareHeaderRowsForModify({ targetRows, metadata });
const insertedHeader = workflow.insertMissingHeaderRows({
  sourceRows,
  preparedHeader
});
const automatedHeader = workflow.applyBrokerageAutomation({
  preparedHeader: insertedHeader,
  metadata,
  brokerageRates
});
```

- [ ] **Step 3: Keep item generation but ensure it uses the automated header rows as its source of truth**

Pass:

```js
const preparedItem = workflow.prepareItemRowsWithCcn({
  itemRows,
  preparedHeader: {
    rows: automatedHeader.rows,
    headerRowIndex: automatedHeader.headerRowIndex
  },
  metadata
});
```

- [ ] **Step 4: Render the new final report instead of the minimal completion message**

Replace `renderModifyCompletion` with a summary-aware version:

```js
function renderModifyCompletion(summary) {
  const compareHtml = summary.item
    ? `<p><strong>Duty Match:</strong> ${summary.compare.dutyMatch ? "Match" : "Mismatch"}</p>
       <p><strong>GST Match:</strong> ${summary.compare.gstMatch ? "Match" : "Mismatch"}</p>`
    : `<p><strong>Header/Item Compare:</strong> not available</p>`;

  const clientHtml = summary.header.clientMatched
    ? `<p><strong>Client Rate Source:</strong> ${summary.header.clientKey}</p>`
    : `<p><strong>Client Rate Source:</strong> not found, brokerage left blank where applicable</p>`;

  reportEl.innerHTML = `...`;
}
```

The rendered header section must include:

- PGA count
- LVS count
- CLVS count
- Duty total
- GST total
- blank brokerage count

- [ ] **Step 5: Keep existing output naming behavior based on the SFTP filename**

Do not change the naming helper behavior beyond ensuring both outputs call:

```js
buildOutputName_MOD(targetFile.name || "updated_target.xlsx", "_DutiesHeader")
buildOutputName_MOD(itemFile.name || "updated_item.xlsx", "_DutiesItem")
```

If the implementation reveals the current helper is still target-file based instead of SFTP-based, patch the caller to pass `sourceFile.name` as the naming source instead.

- [ ] **Step 6: Run regression again after the UI orchestration changes**

Run: `node scripts/regression-smoke.js`

Expected:
- all smoke sections `PASS`

- [ ] **Step 7: Commit the D/T module wiring**

```bash
git add public/script.js
git commit -m "feat(ui): wire dt brokerage automation and summary report"
```

### Task 4: Update Regression Expectations And Project Context

**Files:**
- Modify: `scripts/regression-smoke.js`
- Modify: `documents/PROJECT_CONTEXT.md`
- Reference: `documents/2026-05-27-112-05240631-final-output-observations.md`

- [ ] **Step 1: Remove or update any stale regression assumptions that conflict with the new brokerage behavior**

Search for stale expectations such as blank client/header brokerage placeholders and replace them with D/T-specific assertions only where still valid.

Use:

```js
rg -n "blank|brokerage|client should be blank" scripts/regression-smoke.js
```

- [ ] **Step 2: Update project context for the new D/T workflow**

Revise the D/T section in `documents/PROJECT_CONTEXT.md` to include:

```md
- loads `public/brokerage-rates.json` client-side
- assigns `Brokerage Total` by `CLIENT` + row classification (`8308 => PGA`, `LV => LVS`, `CLVS => CLVS`)
- blanks brokerage when client is unknown
- rewrites header date columns from `RPT DATE`
- zeroes exchange rate
- sorts header rows by brokerage descending
- shows a final summary report with brokerage counts and header/item duty-GST comparison
```

- [ ] **Step 3: Run the full regression one final time**

Run: `node scripts/regression-smoke.js`

Expected:
- `Candata Header: PASS`
- `Candata Item: PASS`
- `GETS Header Carryover: PASS`
- `D/T Header Workflow: PASS`
- `Merge Module: PASS`

- [ ] **Step 4: Commit the docs and final regression alignment**

```bash
git add scripts/regression-smoke.js documents/PROJECT_CONTEXT.md
git commit -m "docs: document dt brokerage automation workflow"
```

## Spec Coverage Self-Check

- Client-rate JSON lookup: Task 2, Task 3
- Unknown client blank brokerage: Task 1, Task 2, Task 3
- PGA/LVS/CLVS assignment: Task 1, Task 2
- Date overwrite: Task 1, Task 2
- Exchange-rate zeroing: Task 1, Task 2
- Brokerage-desc sort: Task 1, Task 2
- Accounting formatting band: Task 2, Task 3
- Item `CCN` carryover preserved: Task 1, Task 3
- Final summary report: Task 1, Task 2, Task 3
- Output naming behavior: Task 3
- Context docs: Task 4

