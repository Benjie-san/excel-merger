# D/T Header and DutiesItem CCN Automation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extend the existing D/T Header workflow so users can enter report metadata, process row-4 or row-5 header layouts, update the header workbook, and optionally generate a paired DutiesItem workbook with direct `Transaction Number -> CCN` carryover.

**Architecture:** Extract the workbook transformation logic into a small shared plain-JS module that can run in both the browser and Node regression smoke tests. Keep `public/script.js` as the DOM/controller layer for the D/T Header page, while the shared module owns metadata row rewriting, row-4 normalization, header modification, item CCN fill, and filename generation.

**Tech Stack:** Express static app, plain HTML/CSS/JS, SheetJS `xlsx`, Node regression smoke script

---

## File Structure

- Create: `public/dt-header-workflow.js`
  - Shared pure workbook helpers for metadata rewrite, row detection, row insertion, header modify, item `CCN` fill, and filename generation.
- Modify: `public/index.html`
  - Add metadata inputs and optional DutiesItem drop zone.
  - Load `dt-header-workflow.js` before `script.js`.
- Modify: `public/styles.css`
  - Add layout and spacing for the metadata inputs and third drop zone.
- Modify: `public/script.js`
  - Replace hardcoded D/T Header assumptions with calls into `window.DtHeaderWorkflow`.
  - Validate metadata inputs and optional DutiesItem file.
  - Keep the current `8308` analysis gate and dual-download flow.
- Modify: `scripts/regression-smoke.js`
  - Import the shared workflow module.
  - Add regressions for row-4 normalization, metadata rewrite, text date handling, and item `CCN` carryover.
- Modify: `documents/PROJECT_CONTEXT.md`
  - Update the current D/T Header workflow description to match the new behavior.

## Task 1: Add Regression Coverage For The Shared Workflow

**Files:**
- Modify: `scripts/regression-smoke.js`
- Test: `scripts/regression-smoke.js`

- [ ] **Step 1: Add a failing import and new regression entry point**

```js
const dtHeaderWorkflow = require("../public/dt-header-workflow.js");

function runDtHeaderWorkflowRegression(rootDir) {
  const issues = [];
  const clvsDir = path.join(rootDir, "112-05240631");
  const apcDir = path.join(rootDir, "83082142460");

  if (!dtHeaderWorkflow || typeof dtHeaderWorkflow.detectHeaderRowIndex !== "function") {
    issues.push("Shared D/T workflow module is missing detectHeaderRowIndex export.");
    return { issues };
  }

  return { issues };
}
```

- [ ] **Step 2: Add failing assertions for row-4 and row-5 header detection**

```js
const clvsHeader = readFirstSheetRows(path.join(clvsDir, "CLVS_Report_Header_10133_017927245_260526112458466.xlsx")).rows;
const clvsItem = readFirstSheetRows(path.join(clvsDir, "CLVS_Report_Detail_10133_017927245_260526112458608.xlsx")).rows;
const apcHeader = readFirstSheetRows(path.join(apcDir, "RLBE_161_8308214246_EWR_DutiesHeader.xlsx")).rows;
const apcItem = readFirstSheetRows(path.join(apcDir, "RLBE_161_8308214246_EWR_DutiesItem.xlsx")).rows;

const clvsHeaderRow = dtHeaderWorkflow.detectHeaderRowIndex(clvsHeader, "header");
const clvsItemRow = dtHeaderWorkflow.detectHeaderRowIndex(clvsItem, "item");
const apcHeaderRow = dtHeaderWorkflow.detectHeaderRowIndex(apcHeader, "header");
const apcItemRow = dtHeaderWorkflow.detectHeaderRowIndex(apcItem, "item");

if (clvsHeaderRow !== 3) issues.push(`Expected CLVS header row index 3, got ${clvsHeaderRow}`);
if (clvsItemRow !== 3) issues.push(`Expected CLVS item row index 3, got ${clvsItemRow}`);
if (apcHeaderRow !== 4) issues.push(`Expected APC header row index 4, got ${apcHeaderRow}`);
if (apcItemRow !== 4) issues.push(`Expected APC item row index 4, got ${apcItemRow}`);
```

- [ ] **Step 3: Add failing assertions for metadata rewrite, text date, row-4 normalization, and item CCN fill**

```js
const metadata = {
  client: "RELIABLE LOGISTICS",
  reportName: "WEEKLY DETAIL REPORT",
  reportDate: "05/27/2026"
};

const normalizedHeader = dtHeaderWorkflow.prepareHeaderRowsForModify({
  targetRows: clvsHeader,
  metadata
});
const normalizedItem = dtHeaderWorkflow.prepareItemRowsWithCcn({
  itemRows: clvsItem,
  headerRows: normalizedHeader.rows,
  metadata
});

if ((normalizedHeader.rows[0] || [])[0] !== "CLIENT:") issues.push("Header row 1 label was not rewritten.");
if ((normalizedHeader.rows[0] || [])[1] !== "RELIABLE LOGISTICS") issues.push("Header CLIENT value mismatch.");
if ((normalizedHeader.rows[2] || [])[1] !== "05/27/2026") issues.push("Header RPT DATE text mismatch.");
if ((normalizedHeader.rows[4] || [])[0] !== "Transaction Number") issues.push("Row-4 header was not normalized to row 5.");
if ((normalizedItem.rows[4] || [])[0] !== "Transaction Number") issues.push("Row-4 item header was not normalized to row 5.");
if ((normalizedItem.rows[4] || []).indexOf("CCN") === -1) issues.push("Item output did not include CCN column.");

const itemCcnIndex = (normalizedItem.rows[4] || []).indexOf("CCN");
if ((normalizedItem.rows[5] || [])[itemCcnIndex] !== "SF5199310030033") {
  issues.push("First CLVS item row did not receive the expected CCN.");
}
```

- [ ] **Step 4: Wire the new regression block into `main()` and verify the test fails**

```js
const dtHeader = runDtHeaderWorkflowRegression(rootDir);

if (dtHeader.issues.length === 0) {
  console.log("D/T Header Workflow: PASS");
} else {
  failed = true;
  console.log("D/T Header Workflow: FAIL");
  dtHeader.issues.forEach((issue) => console.log(`  - ${issue}`));
}
```

Run: `node scripts/regression-smoke.js`  
Expected: `FAIL` with at least one of:
- `Cannot find module '../public/dt-header-workflow.js'`
- `Shared D/T workflow module is missing detectHeaderRowIndex export.`

- [ ] **Step 5: Commit the failing-test harness**

```bash
git add scripts/regression-smoke.js
git commit -m "test: add dt header workflow regression harness"
```

## Task 2: Implement Shared Workbook Helpers In A Reusable Module

**Files:**
- Create: `public/dt-header-workflow.js`
- Test: `scripts/regression-smoke.js`

- [ ] **Step 1: Create the module shell with browser and Node exports**

```js
(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  if (root) {
    root.DtHeaderWorkflow = api;
  }
})(typeof window !== "undefined" ? window : globalThis, function () {
  function cloneRows(rows) {
    return (rows || []).map((row) => Array.isArray(row) ? row.slice() : []);
  }

  return {
    cloneRows
  };
});
```

- [ ] **Step 2: Implement metadata rewrite, header detection, and row normalization helpers**

```js
function normalizeText(value) {
  return String(value == null ? "" : value).trim();
}

function setTextCell(row, index, value) {
  while (row.length <= index) row.push("");
  row[index] = normalizeText(value);
}

function applyMetadataRows(rows, metadata) {
  while (rows.length < 3) rows.push([]);
  setTextCell(rows[0], 0, "CLIENT:");
  setTextCell(rows[0], 1, metadata.client);
  setTextCell(rows[1], 0, "RPT NAME:");
  setTextCell(rows[1], 1, metadata.reportName);
  setTextCell(rows[2], 0, "RPT DATE :");
  setTextCell(rows[2], 1, metadata.reportDate);
  return rows;
}

function detectHeaderRowIndex(rows, mode) {
  const expected = mode === "header"
    ? ["transaction number", "ccn"]
    : ["transaction number", "goods description"];

  for (let r = 0; r < Math.min(rows.length, 6); r++) {
    const normalized = (rows[r] || []).map((cell) => normalizeText(cell).toLowerCase());
    if (expected.every((label) => normalized.includes(label))) return r;
  }
  return -1;
}

function normalizeHeaderRowToFive(rows, headerRowIndex) {
  if (headerRowIndex === 3) {
    rows.splice(3, 0, []);
    return { rows, headerRowIndex: 4 };
  }
  return { rows, headerRowIndex };
}
```

- [ ] **Step 3: Implement shared header and item transformation helpers**

```js
function findColumnIndex(headers, label) {
  return (headers || []).findIndex((cell) => normalizeText(cell).toLowerCase() === label.toLowerCase());
}

function buildTransactionToCcnMap(rows, headerRowIndex) {
  const headers = rows[headerRowIndex] || [];
  const txIndex = findColumnIndex(headers, "Transaction Number");
  const ccnIndex = findColumnIndex(headers, "CCN");
  if (txIndex === -1 || ccnIndex === -1) {
    throw new Error("Header workbook is missing Transaction Number or CCN.");
  }

  const map = new Map();
  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const tx = normalizeText(row[txIndex]);
    const ccn = normalizeText(row[ccnIndex]);
    if (!tx) continue;
    map.set(tx, ccn);
  }
  return map;
}

function ensureItemCcnColumn(headers) {
  const existing = findColumnIndex(headers, "CCN");
  if (existing !== -1) return existing;
  headers.push("CCN");
  return headers.length - 1;
}

function prepareItemRowsWithCcn({ itemRows, headerRows, metadata }) {
  let rows = cloneRows(itemRows);
  applyMetadataRows(rows, metadata);
  const itemHeaderRowIndex = detectHeaderRowIndex(rows, "item");
  if (itemHeaderRowIndex === -1) throw new Error("Could not locate DutiesItem header row.");
  const normalizedItem = normalizeHeaderRowToFive(rows, itemHeaderRowIndex);

  const headerHeaderRowIndex = detectHeaderRowIndex(headerRows, "header");
  if (headerHeaderRowIndex === -1) throw new Error("Could not locate DutiesHeader header row.");
  const lookup = buildTransactionToCcnMap(headerRows, headerHeaderRowIndex);

  const headers = normalizedItem.rows[normalizedItem.headerRowIndex] || [];
  const txIndex = findColumnIndex(headers, "Transaction Number");
  const ccnIndex = ensureItemCcnColumn(headers);
  if (txIndex === -1) throw new Error("DutiesItem workbook is missing Transaction Number.");

  let unmatchedCount = 0;
  for (let r = normalizedItem.headerRowIndex + 1; r < normalizedItem.rows.length; r++) {
    const row = normalizedItem.rows[r] || [];
    const tx = normalizeText(row[txIndex]);
    while (row.length <= ccnIndex) row.push("");
    row[ccnIndex] = tx ? (lookup.get(tx) || "") : "";
    if (tx && !lookup.has(tx)) unmatchedCount++;
  }

  return {
    rows: normalizedItem.rows,
    headerRowIndex: normalizedItem.headerRowIndex,
    unmatchedCount
  };
}
```

- [ ] **Step 4: Implement header modification and filename helpers, then rerun regression smoke**

```js
function prepareHeaderRowsForModify({ targetRows, metadata }) {
  let rows = cloneRows(targetRows);
  applyMetadataRows(rows, metadata);
  const headerRowIndex = detectHeaderRowIndex(rows, "header");
  if (headerRowIndex === -1) throw new Error("Could not locate DutiesHeader header row.");
  return normalizeHeaderRowToFive(rows, headerRowIndex);
}

function modifyHeaderRowsFromSftp({ sourceRows, targetRows, headerRowIndex }) {
  const rows = cloneRows(targetRows);
  const headers = rows[headerRowIndex] || [];
  const ccnIndex = findColumnIndex(headers, "CCN");
  const valueForDutyIndex = findColumnIndex(headers, "Value for Duty");
  const exchangeRateIndex = findColumnIndex(headers, "Exchange Rate");
  if (ccnIndex === -1 || valueForDutyIndex === -1) {
    throw new Error("DutiesHeader workbook is missing required columns.");
  }

  const refSet = new Set();
  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const existing = normalizeText((rows[r] || [])[ccnIndex]);
    if (!existing) continue;
    refSet.add(existing.startsWith("8308") ? existing.slice(4) : existing);
  }

  const insertedRows = [];
  const lastExistingRow = rows[rows.length - 1] || [];
  for (let r = 2; r < sourceRows.length; r++) {
    const sourceRow = sourceRows[r] || [];
    const ac = normalizeText(sourceRow[28]);
    const as = normalizeText(sourceRow[44]);
    if (!ac || refSet.has(ac)) continue;

    const newRow = new Array(Math.max(headers.length, 18)).fill("");
    newRow[0] = "CLVS";
    newRow[1] = ac;
    newRow[2] = lastExistingRow[2] || "";
    newRow[3] = lastExistingRow[3] || "";
    newRow[4] = lastExistingRow[4] || "";
    newRow[5] = lastExistingRow[5] || "";
    newRow[7] = ac;
    newRow[9] = as;
    for (let c = 10; c <= 16; c++) newRow[c] = 0;
    newRow[17] = "DDP";
    insertedRows.push(newRow);
    refSet.add(ac);
  }

  return {
    rows: rows.concat(insertedRows),
    insertedCount: insertedRows.length,
    headerRowIndex,
    numericRange: {
      startRowIndex: headerRowIndex + 1,
      startColIndex: valueForDutyIndex,
      endColIndex: exchangeRateIndex === -1 ? 16 : exchangeRateIndex
    }
  };
}

function buildTimestampedOutputName(fileName, marker, timestamp) {
  const stampRegex = new RegExp(`(\\d{12})(?=${marker})`, "i");
  if (new RegExp(marker, "i").test(fileName)) {
    return stampRegex.test(fileName)
      ? fileName.replace(stampRegex, timestamp)
      : fileName.replace(new RegExp(marker, "i"), `${timestamp}${marker}`);
  }
  const dot = fileName.lastIndexOf(".");
  const base = dot === -1 ? fileName : fileName.slice(0, dot);
  const ext = dot === -1 ? ".xlsx" : fileName.slice(dot);
  return `${base}_${timestamp}${ext}`;
}
```

Run: `node scripts/regression-smoke.js`  
Expected: `D/T Header Workflow: PASS`

- [ ] **Step 5: Commit the shared module implementation**

```bash
git add public/dt-header-workflow.js scripts/regression-smoke.js
git commit -m "feat: add shared dt header workflow helpers"
```

## Task 3: Update The D/T Header UI For Metadata And Optional DutiesItem Input

**Files:**
- Modify: `public/index.html`
- Modify: `public/styles.css`
- Test: `public/index.html`

- [ ] **Step 1: Add metadata inputs and the optional DutiesItem drop zone to the D/T Header form**

```html
<div class="modify-metadata-grid">
  <label class="modify-field">
    <span>CLIENT</span>
    <input type="text" id="modifyClientInput" autocomplete="off" />
  </label>
  <label class="modify-field">
    <span>RPT NAME</span>
    <input type="text" id="modifyReportNameInput" autocomplete="off" />
  </label>
  <label class="modify-field">
    <span>RPT DATE</span>
    <input type="text" id="modifyReportDateInput" autocomplete="off" />
  </label>
</div>

<div id="modifyItemDropZone" class="drop-area">
  <i id="upload-icon" class="bi bi-file-earmark-arrow-up-fill"></i>
  <p>Drag & Drop Duties Item Here</p>
  <p>or click to select</p>
  <input type="file" id="modifyItemFileInput" accept=".xlsx" hidden />
</div>
<p id="modifyItemFileName" class="filename"></p>
```

- [ ] **Step 2: Load the shared module before the main browser script**

```html
<script src="dt-header-workflow.js"></script>
<script src="script.js"></script>
```

- [ ] **Step 3: Add styles for the metadata grid and inputs**

```css
.modify-metadata-grid {
  width: 100%;
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 12px;
  margin: 0 0 18px;
}

.modify-field {
  display: flex;
  flex-direction: column;
  gap: 6px;
  color: #e5e7eb;
  font-size: 13px;
}

.modify-field input[type="text"] {
  width: 100%;
  padding: 10px 12px;
  border: 1px solid #4b5563;
  border-radius: 6px;
  background: #111827;
  color: #fff;
}

.modify-field input[type="text"]:focus {
  outline: 2px solid #ff581f;
  border-color: #ff581f;
}
```

- [ ] **Step 4: Open the page and verify the new controls render without breaking the layout**

Run: `npm run dev`  
Then verify in browser:
- three metadata inputs appear above the drop zones
- SFTP and DutiesHeader zones still render
- new DutiesItem zone appears below DutiesHeader
- existing buttons remain visible

Expected: no console parse errors and no missing-element errors in the D/T Header section

- [ ] **Step 5: Commit the UI scaffold**

```bash
git add public/index.html public/styles.css
git commit -m "feat: add dt header metadata and item inputs"
```

## Task 4: Integrate The Browser Workflow With The Shared Module

**Files:**
- Modify: `public/script.js`
- Test: `public/script.js`

- [ ] **Step 1: Replace hardcoded D/T Header DOM bindings with the expanded set**

```js
const itemDrop = document.getElementById("modifyItemDropZone");
const itemInput = document.getElementById("modifyItemFileInput");
const itemFileName = document.getElementById("modifyItemFileName");

const clientInput = document.getElementById("modifyClientInput");
const reportNameInput = document.getElementById("modifyReportNameInput");
const reportDateInput = document.getElementById("modifyReportDateInput");

if (!modifyForm || !sourceDrop || !targetDrop || !runModifyBtn || !window.DtHeaderWorkflow) return;
```

- [ ] **Step 2: Require metadata inputs during the analysis gate and keep the existing 8308 flow**

```js
function readModifyMetadata() {
  return {
    client: String(clientInput.value || "").trim(),
    reportName: String(reportNameInput.value || "").trim(),
    reportDate: String(reportDateInput.value || "").trim()
  };
}

function validateModifyInputs({ requireItem = false } = {}) {
  const metadata = readModifyMetadata();
  if (!metadata.client || !metadata.reportName || !metadata.reportDate) {
    throw new Error("Please provide CLIENT, RPT NAME, and RPT DATE.");
  }
  if (!sourceInput.files[0] || !targetInput.files[0]) {
    throw new Error("Please provide both SFTP and DutiesHeader files.");
  }
  if (requireItem && !itemInput.files[0]) {
    throw new Error("Please provide a DutiesItem file.");
  }
  return metadata;
}
```

- [ ] **Step 3: Refactor modify execution so header normalization, metadata rewrite, and optional item output all use the shared module**

```js
async function buildWorkbookRows(file) {
  return await readExcelFile_MOD(file);
}

function downloadWorkbookFromRows({ rows, fileName, numericRange }) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(rows);

  if (numericRange) {
    for (let r = numericRange.startRowIndex; r < rows.length; r++) {
      for (let c = numericRange.startColIndex; c <= numericRange.endColIndex; c++) {
        const parsed = parseNumberFromCell_MOD((rows[r] || [])[c]);
        if (parsed === null) continue;
        setWorksheetNumber(worksheet, r, c, parsed);
      }
    }
  }

  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "binary", compression: true, bookSST: false });
  downloadBlobFile(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), fileName);
}

function renderModifyCompletion({ unmatchedItemCount }) {
  const reportEl = document.getElementById("modifyReport");
  if (!reportEl) return;
  reportEl.innerHTML = unmatchedItemCount > 0
    ? `<div class="analyze-report-container"><div class="analyze-8308-report"><h4>Modify Complete</h4><p>Header output downloaded.</p><p>Item output downloaded with ${unmatchedItemCount} unmatched transaction number(s) left blank.</p></div></div>`
    : `<div class="analyze-report-container"><div class="analyze-8308-report"><h4>Modify Complete</h4><p>All requested outputs were generated successfully.</p></div></div>`;
  reportEl.style.display = "block";
}

async function runModifyWorkflow() {
  const metadata = validateModifyInputs();
  const sourceRows = await buildWorkbookRows(sourceInput.files[0]);
  const targetRows = await buildWorkbookRows(targetInput.files[0]);

  const preparedHeader = window.DtHeaderWorkflow.prepareHeaderRowsForModify({
    targetRows,
    metadata
  });

  const modifiedHeader = window.DtHeaderWorkflow.modifyHeaderRowsFromSftp({
    sourceRows,
    targetRows: preparedHeader.rows,
    headerRowIndex: preparedHeader.headerRowIndex
  });

  downloadWorkbookFromRows({
    rows: modifiedHeader.rows,
    fileName: window.DtHeaderWorkflow.buildTimestampedOutputName(targetInput.files[0].name, "_DutiesHeader", generateTimestamp12()),
    numericRange: modifiedHeader.numericRange
  });

  if (itemInput.files && itemInput.files[0]) {
    const itemRows = await buildWorkbookRows(itemInput.files[0]);
    const preparedItem = window.DtHeaderWorkflow.prepareItemRowsWithCcn({
      itemRows,
      headerRows: modifiedHeader.rows,
      metadata
    });
    downloadWorkbookFromRows({
      rows: preparedItem.rows,
      fileName: window.DtHeaderWorkflow.buildTimestampedOutputName(itemInput.files[0].name, "_DutiesItem", generateTimestamp12())
    });
    renderModifyCompletion({ unmatchedItemCount: preparedItem.unmatchedCount });
  } else {
    renderModifyCompletion({ unmatchedItemCount: 0 });
  }
}
```

- [ ] **Step 4: Extend reset behavior and verify the workflow manually**

```js
if (itemFileName) itemFileName.textContent = "";
if (clientInput) clientInput.value = "";
if (reportNameInput) reportNameInput.value = "";
if (reportDateInput) reportDateInput.value = "";
```

Run: `npm run dev`  
Manual verification:
- use `Test files\112-05240631\RLBE_50_11205240631_PVG_YYZ_260427045723.xlsx`
- use `Test files\112-05240631\CLVS_Report_Header_10133_017927245_260526112458466.xlsx`
- use `Test files\112-05240631\CLVS_Report_Detail_10133_017927245_260526112458608.xlsx`
- enter metadata values
- run analysis, then proceed

Expected:
- two downloads
- rows `1-3` rewritten in both files
- headers appear on row `5`
- item file contains `CCN` values resolved from header `Transaction Number`

- [ ] **Step 5: Commit the browser integration**

```bash
git add public/script.js
git commit -m "feat: automate dt header and item workbook outputs"
```

## Task 5: Final Regression, Docs Update, And Release Notes

**Files:**
- Modify: `scripts/regression-smoke.js`
- Modify: `documents/PROJECT_CONTEXT.md`
- Test: `scripts/regression-smoke.js`

- [ ] **Step 1: Add one row-5 regression for existing item CCN reuse**

```js
const existingCcnItem = dtHeaderWorkflow.prepareItemRowsWithCcn({
  itemRows: apcItem,
  headerRows: apcHeader,
  metadata: {
    client: "APC",
    reportName: "AWB # 83082142460",
    reportDate: "03/22/2026"
  }
});

const apcCcnIndex = (existingCcnItem.rows[4] || []).indexOf("CCN");
if ((existingCcnItem.rows[5] || [])[apcCcnIndex] !== "8308APCP0001284765") {
  issues.push("Existing APC item CCN column was not reused correctly.");
}
```

- [ ] **Step 2: Update `documents/PROJECT_CONTEXT.md` with the new D/T Header behavior**

```md
2. D/T Header File Modifier
- Inputs:
  - Source (SFTP file)
  - Target (`_DutiesHeader` file)
  - Optional `DutiesItem` file
  - User-entered `CLIENT`, `RPT NAME`, and `RPT DATE`
- Rewrites rows 1 to 3 in generated outputs from user input.
- Supports row-4 and row-5 header layouts for header and item workbooks.
- If headers are on row 4, inserts one blank row so outputs normalize to row 5.
- If item is provided, writes `CCN` directly from Header `Transaction Number -> CCN`.
```

- [ ] **Step 3: Run the full regression suite and confirm all sections pass**

Run: `node scripts/regression-smoke.js`  
Expected:
- `Candata Header: PASS`
- `Candata Item: PASS`
- `GETS Header Carryover: PASS`
- `Merge Module: PASS` or `PASS (smoke-only...)`
- `D/T Header Workflow: PASS`

- [ ] **Step 4: Sanity-check the working tree and summarize the user-facing change**

Run: `git status --short`

Expected:
- only the intended D/T workflow and docs files are modified

Release summary to capture in commit body or PR notes:
- D/T Header now accepts metadata inputs and optional DutiesItem
- row-4 reports are normalized automatically
- item `CCN` carryover is written directly without formulas

- [ ] **Step 5: Commit the docs and regression completion**

```bash
git add scripts/regression-smoke.js documents/PROJECT_CONTEXT.md
git commit -m "docs: update dt header workflow context"
```

## Self-Review

### Spec coverage

- Metadata inputs for `CLIENT`, `RPT NAME`, and text `RPT DATE`: covered in Task 3 and Task 4.
- Row-4 and row-5 support for header and item workbooks: covered in Task 1 and Task 2.
- Rewrite rows `1-3` before downstream processing: covered in Task 2 and Task 4.
- Existing `8308` analysis gate remains: covered in Task 4.
- Optional DutiesItem generation with direct `Transaction Number -> CCN` fill: covered in Task 2 and Task 4.
- Documentation update: covered in Task 5.

### Placeholder scan

- No `TODO`, `TBD`, or “handle later” placeholders remain.
- Each code step includes actual file paths, function names, code, and commands.

### Type consistency

- Shared module API names are consistent across tasks:
  - `detectHeaderRowIndex`
  - `prepareHeaderRowsForModify`
  - `prepareItemRowsWithCcn`
  - `buildTransactionToCcnMap`
  - `buildTimestampedOutputName`
