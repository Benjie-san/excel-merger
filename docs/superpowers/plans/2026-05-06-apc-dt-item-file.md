# APC D/T Item File Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the `APC D/T Item File` converter UI and client-side XLSX transformation that filters existing item CCNs against SFTP `Reliable_tracking` values and appends only new CLVS rows.

**Architecture:** Extend the existing static HTML tool section with reset controls, then add a dedicated IIFE in `public/script.js` that handles file selection, workbook parsing, item-CCN normalization, SFTP filtering, row mapping, workbook output, and filename generation. Keep all logic isolated from the current D/T Header and Candata modules.

**Tech Stack:** Plain HTML/CSS/JavaScript, SheetJS `xlsx`, browser `FileReader`, existing download helper.

---

### Task 1: Finish APC Item UI Shell

**Files:**
- Modify: `public/index.html`
- Modify: `public/styles.css`

- [ ] **Step 1: Add result/reset container below APC action button**

```html
<div id="apc-item-result-btn">
  <button type="button" id="resetApcItemBtn" style="display:none;">
    <span class="material-symbols-outlined">
    refresh
    </span>
    <p>Start Over</p>
  </button>
</div>
```

- [ ] **Step 2: Extend shared button container selectors for APC item controls**

```css
#result-btn, #modify-result-btn, #analyze-result-btn, #candata-result-btn, #apc-item-result-btn {
  display: flex;
  justify-content: center;
  align-items: center;
  flex-direction: column;
  border-radius: 10px;
  gap: 10px;
  width: 100%;
}

#resetBtn, #resetModifyBtn, #resetAnalyzeBtn, #resetCandataBtn, #resetApcItemBtn {
  background-color: #4b5563;
  display: flex;
  align-items: center;
  justify-content: center;
  flex-direction: start;
  padding: 0px 10px;
}
```

- [ ] **Step 3: Keep APC primary button disabled by default in markup**

```html
<button type="button" id="runApcItem" class="btn-primary" disabled>Generate APC D/T Item File</button>
```

### Task 2: Implement APC Item Conversion Module

**Files:**
- Modify: `public/script.js`

- [ ] **Step 1: Add DOM wiring for APC item form and reset button**

```js
const apcItemForm = document.getElementById("apcItemForm");
const apcSourceDrop = document.getElementById("apcSourceDropZone");
const apcSourceInput = document.getElementById("apcSourceFileInput");
const apcSourceFileName = document.getElementById("apcSourceFileName");
const apcItemDrop = document.getElementById("apcItemDropZone");
const apcItemInput = document.getElementById("apcItemFileInput");
const apcItemFileName = document.getElementById("apcItemFileName");
const runApcItemBtn = document.getElementById("runApcItem");
const resetApcItemBtn = document.getElementById("resetApcItemBtn");
```

- [ ] **Step 2: Add file/drop handlers that update filenames and button enabled state**

```js
function syncApcItemButtonState() {
  const ready = !!(apcSourceInput.files && apcSourceInput.files[0] && apcItemInput.files && apcItemInput.files[0]);
  if (runApcItemBtn) runApcItemBtn.disabled = !ready;
}
```

- [ ] **Step 3: Add workbook helpers for AoA reading, numeric parsing, row emptiness, timestamp generation, and worksheet numeric writes**

```js
async function readExcelFile_APC(file) { /* FileReader -> AoA */ }
function parseNumberFromCell_APC(value) { /* same numeric handling pattern as modify module */ }
function findLastNonEmptyRow_APC(rows) { /* scan from bottom */ }
function setWorksheetNumber_APC(ws, r, c, value) { /* General numeric cell */ }
function generateTimestamp12_APC() { /* yymmddHHmmss */ }
```

- [ ] **Step 4: Add header lookup helpers for row-5 item headers and row-1 SFTP headers**

```js
function normalizeHeaderCell_APC(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function findRequiredIndex_APC(headers, labels, contextLabel) {
  const normalized = headers.map(normalizeHeaderCell_APC);
  for (const label of labels) {
    const idx = normalized.indexOf(normalizeHeaderCell_APC(label));
    if (idx !== -1) return idx;
  }
  throw new Error(`${contextLabel} column not found: ${labels[0]}`);
}
```

- [ ] **Step 5: Implement CCN normalization and appended-row mapping**

```js
function normalizeExistingItemCCN_APC(raw) {
  const value = String(raw || "").trim();
  if (!value) return "";
  return value.startsWith("8308") ? value.slice(4) : value;
}

function buildApcItemRow_APC(sourceRow, indexMap, outputIndexes) {
  const row = new Array(outputIndexes.totalColumns).fill("");
  row[outputIndexes.transactionNumber] = "CLVS";
  row[outputIndexes.goodsDescription] = sanitizeExcelText(sourceRow[indexMap.goodsDescription]);
  row[outputIndexes.lineNumber] = sourceRow[indexMap.packageNo];
  row[outputIndexes.countryOfOrigin] = sanitizeExcelText(sourceRow[indexMap.countryOfOrigin]);
  row[outputIndexes.tariffTreatment] = "";
  row[outputIndexes.partNumber] = sanitizeExcelText(sourceRow[indexMap.productPart]);
  row[outputIndexes.quantity] = sourceRow[indexMap.quantity];
  row[outputIndexes.port] = sanitizeExcelText(sourceRow[indexMap.cbsaPort]);
  row[outputIndexes.vendorName] = sanitizeExcelText(sourceRow[indexMap.sellerName]);
  row[outputIndexes.valueForDuty] = sourceRow[indexMap.totalValueOfParcel];
  row[outputIndexes.hs] = sanitizeExcelText(sourceRow[indexMap.hsCode]);
  row[outputIndexes.dutyRate] = 0;
  row[outputIndexes.duty] = 0;
  row[outputIndexes.valueForTax] = sourceRow[indexMap.totalValueOfParcel];
  row[outputIndexes.gst] = 0;
  row[outputIndexes.incoTerms] = sanitizeExcelText(sourceRow[indexMap.incoTerm]);
  row[outputIndexes.ccn] = sanitizeExcelText(sourceRow[indexMap.reliableTracking]);
  row[outputIndexes.buyerName] = sanitizeExcelText(sourceRow[indexMap.buyerName]);
  row[outputIndexes.buyerAddress] = sanitizeExcelText(sourceRow[indexMap.buyerAddress]);
  row[outputIndexes.buyerCity] = sanitizeExcelText(sourceRow[indexMap.buyerCity]);
  row[outputIndexes.buyerPostalCode] = sanitizeExcelText(sourceRow[indexMap.buyerPostalCode]);
  row[outputIndexes.buyerProvince] = sanitizeExcelText(sourceRow[indexMap.buyerProvince]);
  row[outputIndexes.orderNumber] = sanitizeExcelText(sourceRow[indexMap.orderNumber]);
  return row;
}
```

- [ ] **Step 6: Implement main conversion flow and filename output rule**

```js
async function runApcItemConversion() {
  // validate both files
  // read item + SFTP workbooks
  // resolve item header indexes from row 5
  // resolve SFTP indexes from row 1
  // build normalized existing CCN reference set from item rows 6+
  // keep only SFTP rows 3+ where Reliable_tracking is non-blank and not in ref set
  // ensure Buyer headers exist after CCN or append them
  // append mapped CLVS rows after last non-empty item row
  // write workbook, convert numeric columns on appended rows, download file
  // reveal reset button
}
```

- [ ] **Step 7: Wire run/reset actions**

```js
runApcItemBtn.addEventListener("click", runApcItemConversion);
resetApcItemBtn.addEventListener("click", () => {
  apcItemForm.reset();
  apcSourceFileName.textContent = "";
  apcItemFileName.textContent = "";
  runApcItemBtn.disabled = true;
  resetApcItemBtn.style.display = "none";
});
```

### Task 3: Verify Against Sample Workbooks

**Files:**
- Verify: `Test files/83082142460/RLBE_161_8308214246_EWR_DutiesItem.xlsx`
- Verify: `Test files/83082142460/RLBE_161_8308214246_EWR_FIRST_SECOND_FILE.xlsx`
- Verify: `public/script.js`

- [ ] **Step 1: Run syntax verification**

Run: `node --check public\script.js`
Expected: exit code `0`

- [ ] **Step 2: Run a local Node smoke script against the sample files using the same normalization and filtering rules**

```js
const existing = normalizeExistingItemCCN_APC("8308APCP0001284765");
if (existing !== "APCP0001284765") throw new Error("Normalization failed");
```

Expected:
- existing-item normalized keys include both stripped `8308...` and raw LVS values
- unmatched SFTP row count is non-zero for the sample data

- [ ] **Step 3: Manually verify output expectations in browser or equivalent script inspection**

Check:
- appended rows use `CLVS` in `Transaction Number`
- appended rows use raw `Reliable_tracking` in `CCN`
- added headers appear after `CCN`
- old rows stay blank in new columns

