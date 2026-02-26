//Toggle to show the section
const excelMerger = document.getElementById("excel-merger");
const DTHeaderFile = document.getElementById("modifyTool");
const analyzerTool = document.getElementById("analyzerTool");

function showDisplay(view){
  let mode = view;
  if (view === true) mode = "merger";
  if (view === false) mode = "modify";

  excelMerger.style.display = mode === "merger" ? "flex" : "none";
  DTHeaderFile.style.display = mode === "modify" ? "flex" : "none";
  if (analyzerTool) {
    analyzerTool.style.display = mode === "analyzer" ? "flex" : "none";
  }
}

/*******************************************************************************
 *  CLIENT-SIDE EXCEL MERGER + ANALYZER
 *  - Processes XLSX files entirely in the browser
 *  - Removes top rows
 *  - Adds filename column (optional)
 *  - Removes empty rows
 *  - Merges all files
 *  - Generates Detailed Report (Duty, GST, custom target values)
 *  - Creates downloadable merged.xlsx
 ******************************************************************************/

// ========== DOM ELEMENTS ==========>
const progressContainer = document.getElementById("progressContainer");
const progressBar = document.getElementById("progressBar");
const uploadForm = document.getElementById("uploadForm");
const filesInput = document.getElementById("files");
const downloadLink = document.getElementById("downloadLink");
const resetBtn = document.getElementById("resetBtn");
const dropZone = document.getElementById("dropZone");
const fileList = document.getElementById("fileList");
const fileCount = document.getElementById("fileCount");
const reportDiv = document.getElementById("report");
const billingSlicingRadio = document.getElementById("billingSlicing");
const customSlicingRadio = document.getElementById("customSlicing");
const customSlicingOptions = document.getElementById("customSlicingOptions");
const firstFileRowsInput = document.getElementById("firstFileRows");
const restFileRowsInput = document.getElementById("restFileRows");
const scrollTopBtn = document.getElementById("scrollTopBtn");

// ========== EVENT LISTENERS ==========

billingSlicingRadio.addEventListener("change", () => {
  if (billingSlicingRadio.checked) {
    customSlicingOptions.style.display = "none";
  }
});

customSlicingRadio.addEventListener("change", () => {
  if (customSlicingRadio.checked) {
    customSlicingOptions.style.display = "block";
  }
});

// Scroll-to-top button
if (scrollTopBtn) {
  window.addEventListener("scroll", () => {
    if (window.scrollY > 300) {
      scrollTopBtn.classList.add("show");
    } else {
      scrollTopBtn.classList.remove("show");
    }
  });

  scrollTopBtn.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
}


// =============================================================================
//  Helper: Fake progress effect
// =============================================================================
function simulateProgress() {
  let progress = 0;
  progressBar.style.width = "0%";
  progressBar.textContent = "0%";

  return new Promise((resolve) => {
    const interval = setInterval(() => {
      progress += Math.floor(Math.random() * 10) + 5;
      if (progress >= 95) {
        progress = 95;
        clearInterval(interval);
        resolve();
      }

      progressBar.style.width = progress + "%";
      progressBar.textContent = progress + "%";
    }, 200);
  });
}


// =============================================================================
//  Show file names in list
// =============================================================================
function updateFileList(files) {
  fileList.innerHTML = "";
  Array.from(files).forEach((file) => {
    const li = document.createElement("li");
    li.textContent = file.name;
    fileList.appendChild(li);
  });
}


// =============================================================================
//  Handle manual file selection
// =============================================================================
filesInput.addEventListener("change", () => {
  updateFileList(filesInput.files);
  fileCount.innerHTML = `${filesInput.files.length} files selected`;
});


// =============================================================================
//  Drag and Drop Support
// =============================================================================
dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");

  const files = e.dataTransfer.files;
  if (files.length > 0) {
    filesInput.files = files;
    updateFileList(files);
    fileCount.innerHTML = `${files.length} files selected`;
  }
});


// =============================================================================
//  EXCEL READER (browser-based)
// =============================================================================
async function readExcelFile(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      const data = e.target.result;
      const workbook = XLSX.read(data, { type: "binary" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      resolve(rows);
    };

    reader.readAsBinaryString(file);
  });
}


// =============================================================================
//  MAIN MERGE + CLEANER FUNCTION
// =============================================================================
async function processExcelFiles(files, addFilename) {
  let mergedData = [];
  let headersAdded = false;

  const slicingMode = document.querySelector('input[name="slicing"]:checked').value;
  const firstFileRows = parseInt(firstFileRowsInput.value, 10);
  const restFileRows = parseInt(restFileRowsInput.value, 10);

  for (let i = 0; i < files.length; i++) {
    const file = files[i];

    let sheetData = await readExcelFile(file);
    if (sheetData.length === 0) continue;

    let rowsToRemove = 0;
    if (slicingMode === 'billing') {
        rowsToRemove = i === 0 ? 4 : 5;
    } else {
        rowsToRemove = i === 0 ? firstFileRows : restFileRows;
    }

    let trimmed = sheetData.slice(rowsToRemove);

    // Remove empty rows
    trimmed = trimmed.filter((row) =>
      row.some((c) => c !== null && c !== undefined && c !== "")
    );

    if (trimmed.length === 0) continue;

    // Add header once
    if (!headersAdded) {
      const header = trimmed[0];
      mergedData.push(addFilename ? ["Source File", ...header] : header);
      headersAdded = true;
    }

    // Data rows
    for (let r = 1; r < trimmed.length; r++) {
      mergedData.push(
        addFilename ? [file.name, ...trimmed[r]] : trimmed[r]
      );
    }
  }

  return mergedData;
}


// =============================================================================
//  ANALYSIS: Count target values (only in "Brokerage Total" column)
// =============================================================================
function analyzeData(mergedData) {
  const headers = mergedData[0];
  const rows = mergedData.slice(1);

  const report = {
    totalRows: rows.length,
    columnSummary: {},
    targetValueCounts: {}
  };

  // Target values to check
  const targetValues = [0.0175, 0.085, 0.71, 0.28];
  const tolerance = 1e-3;
  targetValues.forEach(v => report.targetValueCounts[v] = 0);

  // Column selection
  const colIndex = headers.indexOf("Brokerage Total");

  if (colIndex === -1) {
    console.warn("Column 'Brokerage Total' not found");
    return report;
  }

  // Count matching values
  rows.forEach(row => {
    const cell = row[colIndex];
    if (cell == null) return;

    const num = parseFloat(String(cell).trim());
    if (isNaN(num)) return;

    const rounded = parseFloat(num.toFixed(3));
    targetValues.forEach(v => {
      if (Math.abs(rounded - v) < tolerance) {
        report.targetValueCounts[v]++;
      }
    });
  });

  // SUMMARIES (Duty, GST)
  headers.forEach((colName, idx) => {
    const colValues = rows
      .map(row => row[idx])
      .filter(v => v !== "" && v !== undefined && v !== null);

    const numeric = colValues
      .map(v => parseFloat(v))
      .filter(n => !isNaN(n));

    const sum = numeric.reduce((a, b) => a + b, 0);

    report.columnSummary[colName] = { sum };
  });

  return report;
}

// ===============================
// FORCE GENERAL NUMBER FORMAT
// ===============================
function forceGeneralNumber(ws, r, c, value) {
  const ref = XLSX.utils.encode_cell({ r, c });
  ws[ref] = {
    t: "n",
    v: value,
    z: "General" // General numeric (no currency)
  };
}

function parseNumberFromCell(value) {
  if (value === undefined || value === null) return null;
  let s = String(value).trim();
  if (s === "") return null;
  if (/^\$?\s*-\s*\$?$/.test(s)) return 0;

  let neg = false;
  if (s.startsWith("(") && s.endsWith(")")) {
    neg = true;
    s = s.slice(1, -1);
  }
  if (s.endsWith("-")) {
    neg = true;
    s = s.slice(0, -1);
  }

  s = s.replace(/[$,]/g, "").replace(/\s+/g, "");
  if (s === "") return null;
  if (s === "-") return 0;

  const num = parseFloat(s);
  if (isNaN(num)) return null;
  return neg ? -num : num;
}

function convertColumnRangeToNumbers(aoa, startRow, colStart, colEnd, ws) {
  for (let r = startRow; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row) continue;

    for (let c = colStart; c <= colEnd; c++) {
      const v = row[c];
      const num = parseNumberFromCell(v);

      if (num !== null) {
        aoa[r][c] = num;          // update AoA
        forceGeneralNumber(ws, r, c, num); // override Excel formatting
      }
    }
  }
}

function exportMergedExcel(mergedData) {
  if (!mergedData.length) return;

  const aoa = mergedData.map(r => [...r]);

  // detect file type using filenames
  const headerMode  = Array.from(filesInput.files).some(f => f.name.includes("_DutiesHeader"));
  const itemMode    = Array.from(filesInput.files).some(f => f.name.includes("_DutiesItem"));

  // workbook and sheet
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // row 1 downward (because you removed first 5 rows earlier)
  const startRow = 1;

  // DutiesHeader -> convert Value for Duty -> Exchange Rate when available
  if (headerMode) {
    const headers = aoa[0] || [];
    const startIdx = headers.indexOf("Value for Duty");
    const endIdx = headers.indexOf("Exchange Rate");
    if (startIdx !== -1 && endIdx !== -1 && startIdx <= endIdx) {
      convertColumnRangeToNumbers(aoa, startRow, startIdx, endIdx, ws);
    } else {
      // fallback to J (9) -> Q (16)
      convertColumnRangeToNumbers(aoa, startRow, 9, 16, ws);
    }
  }

  // DutiesItem → convert only Duty + GST
  if (itemMode) {
    const headers = aoa[0];
    const dutyIdx = headers.indexOf("Duty");
    const gstIdx  = headers.indexOf("Gov. Sales Tax");

    if (dutyIdx !== -1)
      convertColumnRangeToNumbers(aoa, startRow, dutyIdx, dutyIdx, ws);

    if (gstIdx !== -1)
      convertColumnRangeToNumbers(aoa, startRow, gstIdx, gstIdx, ws);
  }

  XLSX.utils.book_append_sheet(wb, ws, "Merged");

  const wbout = XLSX.write(wb, {
    bookType: "xlsx",
    type: "binary",
    compression: true,
    bookSST: false
  });

  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }

  saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), "merged.xlsx");
}


// =============================================================================
//  MAIN SUBMIT HANDLER (NO BACKEND)
// =============================================================================
uploadForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  reportDiv.innerHTML = "";
  downloadLink.style.display = "none";

  const files = filesInput.files;
  if (!files.length) {
    alert("Please select files");
    return;
  }

  const addFilename = document.getElementById("addFilename").checked;

  progressContainer.style.display = "block";
  await simulateProgress();

  // 1. Merge files
  const mergedData = await processExcelFiles(files, addFilename);

  // Complete progress
  progressBar.style.width = "100%";
  progressBar.textContent = "100%";

  // 2. Analyze data
  const report = analyzeData(mergedData);

  // 3. Show report
  reportDiv.innerHTML = `
      <h3>📊 Report</h3>
      <p><b>Total Rows:</b> ${report.totalRows}</p>
      <p><b>Total Duty:</b> ${Math.abs(report.columnSummary.Duty?.sum || 0).toFixed(2)}</p>
      <p><b>Total GST:</b> ${Math.abs(report.columnSummary["Gov. Sales Tax"]?.sum || 0).toFixed(2)}</p>

      <h4>Value Counts (Brokerage Total)</h4>
      <table border="1" cellpadding="5">
        <tr><th>Value</th><th>Count</th></tr>
        ${Object.entries(report.targetValueCounts)
          .map(([v, c]) => `<tr><td>${v}</td><td>${c}</td></tr>`)
          .join("")}
      </table>
  `;

  // 4. Export merged.xlsx
  exportMergedExcel(mergedData);

  // Show reset
  resetBtn.style.display = "flex";
});


// =============================================================================
//  RESET BUTTON
// =============================================================================
resetBtn.addEventListener("click", () => {
  uploadForm.reset();
  downloadLink.style.display = "none";
  resetBtn.style.display = "none";

  progressBar.style.width = "0%";
  progressBar.textContent = "0%";
  progressContainer.style.display = "none";

  fileList.innerHTML = "";
  fileCount.innerHTML = "No files selected";
  reportDiv.innerHTML = "";
  billingSlicingRadio.checked = true;
  customSlicingRadio.checked = false;
  customSlicingOptions.style.display = "none";
  firstFileRowsInput.value = "0";
  restFileRowsInput.value = "0";

});


function generateTimestamp12() {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(-2);
  const MM = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const HH = String(now.getHours()).padStart(2, "0");
  const mm = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");
  console.log(yy + MM + dd + HH + mm + ss)
  return yy + MM + dd + HH + mm + ss;
}


/**************************************************************
 * Excel Modifier Module (Full) — Exact-match (Option A)
 * - Isolated (IIFE) to avoid colliding with merger code
 * - Exact CCN matching: remove "8308" only if at start of target H
 * - Insert rows from source (AC -> B/H, AS -> J), set A="CLVS"
 * - Copy C-F from last non-empty existing target row
 * - K..Q = 0, R = "DDP"
 * - Convert J->Q (row 6 onward) to true numbers and force General format
 * - Auto-download with updated timestamp in filename (12-digit YYMMDDHHmmSS)
 **************************************************************/
(function ExcelModifyModule() {
  // DOM bindings (must exist in your HTML)
  const modifyForm = document.getElementById("modifyForm");
  const sourceDrop = document.getElementById("sourceDropZone");
  const sourceInput = document.getElementById("sourceFileInput");
  const sourceFileName = document.getElementById("sourceFileName");

  const targetDrop = document.getElementById("targetDropZone");
  const targetInput = document.getElementById("targetFileInput");
  const targetFileName = document.getElementById("targetFileName");

  const runModifyBtn = document.getElementById("runModify");
  const resetModifyBtn = document.getElementById("resetModifyBtn");

  /* -------------------------
     Dropzone setup (isolated)
     ------------------------- */
  function setupDropZone_MOD(dropArea, fileInput, fileNameDisplay) {
    dropArea.addEventListener("dragover", (e) => { e.preventDefault(); dropArea.classList.add("dragover"); });
    dropArea.addEventListener("dragleave", () => dropArea.classList.remove("dragover"));
    dropArea.addEventListener("drop", (e) => {
      e.preventDefault(); dropArea.classList.remove("dragover");
      const f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (!f) return;
      if (!f.name.toLowerCase().endsWith(".xlsx")) { alert("Please drop a .xlsx file"); return; }
      try { const dt = new DataTransfer(); dt.items.add(f); fileInput.files = dt.files; } catch (err) { /*ignore*/ }
      fileNameDisplay.textContent = f.name;
      console.log("Drop set:", f.name);
    });
    dropArea.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", (e) => {
      const f = e.target.files && e.target.files[0]; if (!f) return;
      if (!f.name.toLowerCase().endsWith(".xlsx")) { alert("Please select a .xlsx file"); e.target.value = ""; return; }
      fileNameDisplay.textContent = f.name;
      console.log("Input set:", f.name);
    });
  }

  setupDropZone_MOD(sourceDrop, sourceInput, sourceFileName);
  setupDropZone_MOD(targetDrop, targetInput, targetFileName);

  /* -------------------------
     Helpers
     ------------------------- */
  function generateTimestamp12() {
    const now = new Date();
    const yy = String(now.getFullYear()).slice(-2);
    const MM = String(now.getMonth() + 1).padStart(2, "0");
    const dd = String(now.getDate()).padStart(2, "0");
    const HH = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    return yy + MM + dd + HH + mm + ss;
  }

  // read Excel file -> AoA (unique to module)
  async function readExcelFile_MOD(file) {
    return new Promise((resolve, reject) => {
      if (!file) return resolve([]);
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target.result;
          const wb = XLSX.read(data, { type: "binary" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
          resolve(rows);
        } catch (err) { reject(err); }
      };
      reader.onerror = (err) => reject(err);
      reader.readAsBinaryString(file);
    });
  }

  // force a worksheet cell to numeric general
  function setWorksheetNumber(ws, r, c, value) {
    const ref = XLSX.utils.encode_cell({ r, c });
    ws[ref] = { t: "n", v: value, z: "General" };
  }

  function parseNumberFromCell_MOD(value) {
    if (value === undefined || value === null) return null;
    let s = String(value).trim();
    if (s === "") return null;
    if (/^\$?\s*-\s*\$?$/.test(s)) return 0;

    let neg = false;
    if (s.startsWith("(") && s.endsWith(")")) {
      neg = true;
      s = s.slice(1, -1);
    }
    if (s.endsWith("-")) {
      neg = true;
      s = s.slice(0, -1);
    }

    s = s.replace(/[$,]/g, "").replace(/\s+/g, "");
    if (s === "") return null;
    if (s === "-") return 0;

    const num = parseFloat(s);
    if (isNaN(num)) return null;
    return neg ? -num : num;
  }

  // find last non-empty row in AoA (search from bottom)
  function findLastNonEmptyRow(rows) {
    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i];
      if (!row) continue;
      if (row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== "")) return i;
    }
    return -1;
  }

  /* -------------------------
     Clean target CCN per Option A
     - remove "8308" only if at start, then keep EXACT rest
     ------------------------- */
  function cleanTargetCCN(raw) {
    if (raw === undefined || raw === null) return "";
    let s = String(raw).trim();
    if (s.startsWith("8308")) s = s.substring(4); // remove ONLY prefix
    return s;
  }

  /* -------------------------
     Core modify function
     ------------------------- */
  async function modifyAndDownloadExactMatch_MOD({
    sourceFileObj,
    targetFileObj,
    ccnColumnIndex = 7,        // H
    ccnStartRowIndex = 5,      // H6 -> index 5
    sourceACStartIndex = 2,    // AC3 -> index 2
    sourceASStartIndex = 2     // AS3 -> index 2
  } = {}) {
    try {
      if (!sourceFileObj || !targetFileObj) { alert("Please provide Source and Target files."); return; }

      // column constants
      const COL_AC = 28, COL_AS = 44;
      const COL_A = 0, COL_B = 1, COL_C = 2, COL_D = 3, COL_E = 4, COL_F = 5, COL_H = 7;
      const COL_J = 9, COL_K = 10, COL_Q = 16, COL_R = 17;

      // read files
      console.log("Reading target...");
      const tgtRows = await readExcelFile_MOD(targetFileObj);
      console.log("Reading source...");
      const srcRows = await readExcelFile_MOD(sourceFileObj);

      const targetRows = Array.isArray(tgtRows) ? tgtRows : [];
      const sourceRows = Array.isArray(srcRows) ? srcRows : [];

      // Find last data row; we'll insert new rows *before* trailing blanks
      const lastNonEmptyIndex = findLastNonEmptyRow(targetRows);
      const dataTargetRows = lastNonEmptyIndex >= 0
        ? targetRows.slice(0, lastNonEmptyIndex + 1)
        : targetRows;
      console.log("Target rows (raw):", targetRows.length, "data:", dataTargetRows.length);

      // build refSet from target H (exact cleaned strings)
      const refSet = new Set();
      for (let r = ccnStartRowIndex; r < dataTargetRows.length; r++) {
        const row = dataTargetRows[r] || [];
        const raw = row[ccnColumnIndex];
        if (raw === undefined || raw === null) continue;
        const cleaned = cleanTargetCCN(raw);
        if (cleaned !== "") refSet.add(cleaned);
      }
      console.log("refSet size:", refSet.size);

      // build source items (keep AC as EXACT trimmed string)
      const sourceItems = [];
      const sourceSeen = new Set();
      for (let r = sourceACStartIndex; r < sourceRows.length; r++) {
        const row = sourceRows[r] || [];
        const acRaw = (row[COL_AC] === undefined || row[COL_AC] === null) ? "" : String(row[COL_AC]).trim();
        const asRaw = (row[COL_AS] === undefined || row[COL_AS] === null) ? "" : String(row[COL_AS]).trim();
        if (acRaw === "" && asRaw === "") continue;
        if (acRaw !== "") {
          if (sourceSeen.has(acRaw)) continue;
          sourceSeen.add(acRaw);
        }
        sourceItems.push({ rowIndex: r, acRaw, asRaw });
      }
      console.log("sourceItems:", sourceItems.length);

      // determine copy-from row for C-F: last non-empty row in trimmed target
      const lastExistingRow = lastNonEmptyIndex >= 0 ? dataTargetRows[lastNonEmptyIndex] : [];

      // prepare inserted rows array
      // targetRowLen: base width derived from header or safe default
      const headerIndex = 0;
      const headerRow = dataTargetRows[headerIndex] || [];
      const targetRowLen = Math.max(headerRow.length, COL_R + 1, COL_Q + 1, COL_J + 1, 25);

      const insertedRows = [];
      let skippedExact = 0;

      for (const item of sourceItems) {
        const acTrim = item.acRaw; // exact trimmed string

        // skip if exact match exists in refSet
        if (acTrim !== "" && refSet.has(acTrim)) {
          skippedExact++;
          continue;
        }

        // build new row
        const newRow = new Array(targetRowLen).fill("");

        newRow[COL_A] = "CLVS";            // Column A
        newRow[COL_B] = item.acRaw;        // Column B
        newRow[COL_H] = item.acRaw;        // Column H (CCN)
        newRow[COL_J] = item.asRaw;        // Column J

        // copy C-F from lastExistingRow if present
        newRow[COL_C] = lastExistingRow[COL_C] ?? "";
        newRow[COL_D] = lastExistingRow[COL_D] ?? "";
        newRow[COL_E] = lastExistingRow[COL_E] ?? "";
        newRow[COL_F] = lastExistingRow[COL_F] ?? "";

        // K..Q -> 0
        for (let c = COL_K; c <= COL_Q; c++) newRow[c] = 0;

        // R -> "DDP"
        newRow[COL_R] = "DDP";

        insertedRows.push(newRow);
      }

      console.log("Inserted:", insertedRows.length, "Skipped:", skippedExact);

      // Build final AoA (preserve original target rows, then appended inserted rows)
      const insertAt = lastNonEmptyIndex >= 0 ? lastNonEmptyIndex + 1 : 0;
      const finalAoA = targetRows
        .slice(0, insertAt)
        .concat(insertedRows, targetRows.slice(insertAt));

      // Create workbook and worksheet from finalAoA
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(finalAoA);

      // Convert Value for Duty -> Exchange Rate to numbers starting at row index 5 (Excel row 6)
      const startIndex = 5;
      const headerRowFinal = finalAoA[0] || [];
      let colStart = headerRowFinal.indexOf("Value for Duty");
      let colEnd = headerRowFinal.indexOf("Exchange Rate");
      if (colStart === -1 || colEnd === -1 || colStart > colEnd) {
        colStart = 9;
        colEnd = 16;
      }
      for (let r = startIndex; r < finalAoA.length; r++) {
        const row = finalAoA[r] || [];
        for (let c = colStart; c <= colEnd; c++) {
          const rawVal = row[c];
          const num = parseNumberFromCell_MOD(rawVal);
          if (num !== null) {
            // update worksheet cell as number and force General numeric format
            setWorksheetNumber(ws, r, c, num);
            // also update AoA to keep consistent (optional)
            finalAoA[r][c] = num;
          }
        }
      }

      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      // Determine output filename from targetFileObj and replace/insert 12-digit timestamp before _DutiesHeader
      let outName = targetFileObj.name || "updated_target.xlsx";
      const stampRegex = /(\d{12})(?=_DutiesHeader)/;
      const newStamp = generateTimestamp12();
      if (/_DutiesHeader/i.test(outName)) {
        if (stampRegex.test(outName)) {
          outName = outName.replace(stampRegex, newStamp);
        } else {
          outName = outName.replace(/_DutiesHeader/i, `${newStamp}_DutiesHeader`);
        }
      } else {
        // fallback: append timestamp
        const dotIdx = outName.lastIndexOf(".");
        const base = dotIdx === -1 ? outName : outName.slice(0, dotIdx);
        const ext = dotIdx === -1 ? "" : outName.slice(dotIdx);
        outName = `${base}_${newStamp}${ext || ".xlsx"}`;
      }

      // Write and download
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary", compression: true, bookSST: false });
      function s2ab(s) { const buf = new ArrayBuffer(s.length); const view = new Uint8Array(buf); for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff; return buf; }
      saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), outName);

      // UI report
      if (resetModifyBtn) {
        resetModifyBtn.style.display = 'flex';
      }

      console.log("Modify complete. Downloaded:", outName);
    } catch (err) {
      console.error("modify error:", err);
      alert("Modify error: " + (err && err.message ? err.message : err));
    }
  }

  /* wire run button */
  runModifyBtn.addEventListener("click", async () => {
    const src = sourceInput.files && sourceInput.files[0] ? sourceInput.files[0] : null;
    const tgt = targetInput.files && targetInput.files[0] ? targetInput.files[0] : null;
    console.log("Run modify: src=", src && src.name, "tgt=", tgt && tgt.name);
    await modifyAndDownloadExactMatch_MOD({ sourceFileObj: src, targetFileObj: tgt });
  });

  /* wire reset button */
  resetModifyBtn.addEventListener("click", () => {
    // Reset the form, which clears the file inputs
    if(modifyForm) modifyForm.reset();

    // Clear file name displays
    if(sourceFileName) sourceFileName.textContent = "";
    if(targetFileName) targetFileName.textContent = "";

    // Hide the reset button
    resetModifyBtn.style.display = "none";

    console.log("Modify Tool has been reset.");
  });

})(); // end IIFE

/**************************************************************
 * Header/Item Analyzer Module (UI Only)
 * - Accepts DutiesHeader + DutiesItem files
 * - Shows placeholder report until analysis rules are provided
 **************************************************************/
(function HeaderItemAnalyzerModule() {
  const analyzerForm = document.getElementById("analyzerForm");
  const headerDrop = document.getElementById("headerDropZone");
  const headerInput = document.getElementById("headerFileInput");
  const headerFileName = document.getElementById("headerFileName");

  const itemDrop = document.getElementById("itemDropZone");
  const itemInput = document.getElementById("itemFileInput");
  const itemFileName = document.getElementById("itemFileName");

  const runAnalyzeBtn = document.getElementById("runAnalyze");
  const analyzeReportEl = document.getElementById("analyzeReport");
  const resetAnalyzeBtn = document.getElementById("resetAnalyzeBtn");

  if (!analyzerForm || !headerDrop || !headerInput || !itemDrop || !itemInput) return;

  function setupDropZone_AN(dropArea, fileInput, fileNameDisplay) {
    dropArea.addEventListener("dragover", (e) => { e.preventDefault(); dropArea.classList.add("dragover"); });
    dropArea.addEventListener("dragleave", () => dropArea.classList.remove("dragover"));
    dropArea.addEventListener("drop", (e) => {
      e.preventDefault(); dropArea.classList.remove("dragover");
      const f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (!f) return;
      if (!f.name.toLowerCase().endsWith(".xlsx")) { alert("Please drop a .xlsx file"); return; }
      try { const dt = new DataTransfer(); dt.items.add(f); fileInput.files = dt.files; } catch (err) { /*ignore*/ }
      if (fileNameDisplay) fileNameDisplay.textContent = f.name;
    });
    dropArea.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", (e) => {
      const f = e.target.files && e.target.files[0]; if (!f) return;
      if (!f.name.toLowerCase().endsWith(".xlsx")) { alert("Please select a .xlsx file"); e.target.value = ""; return; }
      if (fileNameDisplay) fileNameDisplay.textContent = f.name;
    });
  }

  setupDropZone_AN(headerDrop, headerInput, headerFileName);
  setupDropZone_AN(itemDrop, itemInput, itemFileName);

  function renderAnalyzerReport(data) {
    if (!analyzeReportEl) return;
    const safe = (v) => (v === undefined || v === null ? "-" : v);
    const header = data.header || {};
    const item = data.item || {};
    const compare = data.compare || {};
    const itemAvailable = !!data.itemAvailable;
    analyzeReportEl.innerHTML = `
      <div class="analyze-report-grid header-only">
        <div class="analyze-report-col">
          <h4>Header File</h4>
          <p><b>Total CCNs:</b> ${safe(header.totalCCNs)}</p>
          <p><b>Total CLVS:</b> ${safe(header.totalCLVS)}</p>
          <p><b>Total LVS:</b> ${safe(header.totalLVS)}</p>
          <p><b>Total PGA:</b> ${safe(header.totalPGA)}</p>
          <p><b>Empty Brokerage Fee (CCNs):</b></p>
          <p>${safe(header.emptyBrokerageCCNs)}</p>
        </div>
        <div class="analyze-report-col">
          <h4>Exceptions</h4>
          <p><b>Empty Value for Duty (CCNs):</b></p>
          <p>${safe(header.emptyValueForDutyCCNs)}</p>
          <p><b>GST = 0 with Value for Duty threshold (CCNs):</b></p>
          <p>${safe(header.gstZeroCCNs)}</p>
          <p><b>Value for Duty &lt; 20 with Duty/GST &gt; 0 (CCNs):</b></p>
          <p>${safe(header.lowValueDutyCCNs)}</p>
        </div>
      </div>
      ${
        itemAvailable
          ? `
      <hr>
      <div class="analyze-report-compare">
        <h4>Totals Match</h4>
        <p><b>Duty Totals:</b> ${safe(compare.duty)}</p>
        <p><b>GST Totals:</b> ${safe(compare.gst)}</p>
      </div>
      `
          : ""
      }
    `;
  }

  function readExcelFile_AN(file) {
    return new Promise((resolve, reject) => {
      if (!file) return resolve([]);
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target.result;
          const wb = XLSX.read(data, { type: "binary" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
          resolve(rows);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = (err) => reject(err);
      reader.readAsBinaryString(file);
    });
  }

  function findHeaderRow(rows) {
    const maxScan = Math.min(rows.length, 15);
    let bestIndex = 0;
    let bestScore = -1;
    for (let r = 0; r < maxScan; r++) {
      const row = rows[r] || [];
      const textRow = row.map(v => String(v || "").trim().toLowerCase());
      const score =
        (textRow.some(v => v.includes("ccn")) ? 1 : 0) +
        (textRow.some(v => v.includes("brokerage")) ? 1 : 0) +
        (textRow.some(v => v.includes("value for duty")) ? 1 : 0) +
        (textRow.some(v => v.includes("gov. sales")) || textRow.some(v => v.includes("gst")) ? 1 : 0);
      if (score > bestScore) {
        bestScore = score;
        bestIndex = r;
      }
    }
    return bestScore > 0 ? bestIndex : 0;
  }

  function findColumnIndex(headers, matchers) {
    for (let i = 0; i < headers.length; i++) {
      const cell = String(headers[i] || "").trim().toLowerCase();
      if (!cell) continue;
      if (matchers.some(m => m.test(cell))) return i;
    }
    return -1;
  }

  function findDutyColumnIndex(headers) {
    let fallback = -1;
    for (let i = 0; i < headers.length; i++) {
      const cell = String(headers[i] || "").trim().toLowerCase();
      if (!cell) continue;
      if (cell.includes("value for duty")) continue;
      if (cell === "duty") return i;
      if (cell.includes("duty rate")) continue;
      if (cell.includes("duty") && fallback === -1) fallback = i;
    }
    return fallback;
  }

  function isEmptyCell(v) {
    return v === undefined || v === null || String(v).trim() === "";
  }

  function approxEqual(a, b, eps = 1e-4) {
    return Math.abs(a - b) <= eps;
  }

  function formatNumber(n) {
    if (n === undefined || n === null || isNaN(n)) return "-";
    return Number(n).toFixed(2);
  }

  function analyzeDutiesRows(rows, { hasBrokerage }) {
    const headerRowIndex = findHeaderRow(rows);
    const headers = rows[headerRowIndex] || [];
    const dataRows = rows.slice(headerRowIndex + 1);

    const ccnIdx = findColumnIndex(headers, [/ccn/i]);
    const brokerageIdx = hasBrokerage ? findColumnIndex(headers, [/brokerage/i]) : -1;
    const valueForDutyIdx = findColumnIndex(headers, [/value for duty/i]);
    const dutyIdx = findDutyColumnIndex(headers);
    const gstIdx = findColumnIndex(headers, [/gov\.?\s*sales/i, /\bgst\b/i]);

    const missing = [];
    if (ccnIdx === -1) missing.push("CCN");
    if (valueForDutyIdx === -1) missing.push("Value for Duty");
    if (dutyIdx === -1) missing.push("Duty");
    if (gstIdx === -1) missing.push("GST");
    if (hasBrokerage && brokerageIdx === -1) missing.push("Brokerage");
    if (missing.length) {
      return { error: `Missing columns: ${missing.join(", ")}` };
    }

    const clvsValues = [0.0175, 0.35, 0.085];
    const lvsValues = [0.28, 0.35];
    const pgaValues = [0.71, 0.5, 0.6, 2.25];

    const uniqueCCNs = new Set();
    let totalCLVS = 0;
    let totalLVS = 0;
    let totalPGA = 0;

    let totalDuty = 0;
    let totalGST = 0;

    const emptyBrokerageSet = new Set();
    const emptyValueForDutySet = new Set();
    const gstZeroSet = new Set();
    const lowValueDutySet = new Set();

    for (const row of dataRows) {
      if (!row) continue;
      const ccnRaw = row[ccnIdx];
      const ccn = ccnRaw === undefined || ccnRaw === null ? "" : String(ccnRaw).trim();
      if (ccn !== "") uniqueCCNs.add(ccn);

      const brokerageRaw = hasBrokerage ? row[brokerageIdx] : null;
      const valueForDutyRaw = row[valueForDutyIdx];
      const dutyRaw = row[dutyIdx];
      const gstRaw = row[gstIdx];

      if (hasBrokerage && isEmptyCell(brokerageRaw) && ccn) {
        emptyBrokerageSet.add(ccn);
      }

      if (isEmptyCell(valueForDutyRaw) && ccn) {
        emptyValueForDutySet.add(ccn);
      }

      const brokerageVal = hasBrokerage ? parseNumberFromCell(brokerageRaw) : null;
      const valueForDutyVal = parseNumberFromCell(valueForDutyRaw);
      const dutyVal = parseNumberFromCell(dutyRaw);
      const gstVal = parseNumberFromCell(gstRaw);

      if (hasBrokerage && brokerageVal !== null) {
        if (clvsValues.some(v => approxEqual(brokerageVal, v))) totalCLVS++;
        if (lvsValues.some(v => approxEqual(brokerageVal, v))) totalLVS++;
        if (pgaValues.some(v => approxEqual(brokerageVal, v))) totalPGA++;
      }

      if (dutyVal !== null) totalDuty += dutyVal;
      if (gstVal !== null) totalGST += gstVal;

      if (gstVal !== null && approxEqual(gstVal, 0) && valueForDutyVal !== null) {
        const isPga225 = hasBrokerage && brokerageVal !== null && approxEqual(brokerageVal, 2.25);
        const threshold = isPga225 ? 40.1 : 20.1;
        if (valueForDutyVal > threshold && ccn) {
          gstZeroSet.add(ccn);
        }
      }

      if (valueForDutyVal !== null && valueForDutyVal < 20) {
        const hasDutyOrGst = (dutyVal !== null && Math.abs(dutyVal) > 0) || (gstVal !== null && Math.abs(gstVal) > 0);
        if (hasDutyOrGst && ccn) {
          lowValueDutySet.add(ccn);
        }
      }
    }

    return {
      totalCCNs: uniqueCCNs.size,
      totalCLVS: hasBrokerage ? totalCLVS : "-",
      totalLVS: hasBrokerage ? totalLVS : "-",
      totalPGA: hasBrokerage ? totalPGA : "-",
      totalDuty: formatNumber(totalDuty),
      totalGST: formatNumber(totalGST),
      totalDutyValue: totalDuty,
      totalGSTValue: totalGST,
      emptyBrokerageCCNs: hasBrokerage ? (emptyBrokerageSet.size ? Array.from(emptyBrokerageSet).join(", ") : "-") : "-",
      emptyValueForDutyCCNs: emptyValueForDutySet.size ? Array.from(emptyValueForDutySet).join(", ") : "-",
      gstZeroCCNs: gstZeroSet.size ? Array.from(gstZeroSet).join(", ") : "-",
      lowValueDutyCCNs: lowValueDutySet.size ? Array.from(lowValueDutySet).join(", ") : "-"
    };
  }

  function emptyItemStats() {
    return {
      totalCCNs: "-",
      totalDuty: "-",
      totalGST: "-",
      totalDutyValue: null,
      totalGSTValue: null,
      emptyValueForDutyCCNs: "-",
      gstZeroCCNs: "-",
      lowValueDutyCCNs: "-"
    };
  }

  // Initial empty report
  renderAnalyzerReport({
    header: {
      totalCCNs: "",
      totalCLVS: "",
      totalLVS: "",
      totalPGA: "",
      totalDuty: "",
      totalGST: "",
      totalDutyValue: null,
      totalGSTValue: null,
      emptyBrokerageCCNs: "",
      emptyValueForDutyCCNs: "",
      gstZeroCCNs: "",
      lowValueDutyCCNs: ""
    },
    item: {
      totalCCNs: "",
      totalDuty: "",
      totalGST: "",
      totalDutyValue: null,
      totalGSTValue: null,
      emptyValueForDutyCCNs: "",
      gstZeroCCNs: "",
      lowValueDutyCCNs: ""
    },
    compare: {
      duty: "",
      gst: ""
    },
    itemAvailable: false
  });

  if (runAnalyzeBtn) {
    runAnalyzeBtn.addEventListener("click", async () => {
      const headerFile = headerInput.files && headerInput.files[0] ? headerInput.files[0] : null;
      const itemFile = itemInput.files && itemInput.files[0] ? itemInput.files[0] : null;
      if (!headerFile) {
        alert("Please provide a DutiesHeader file.");
        return;
      }

      try {
        const headerRows = await readExcelFile_AN(headerFile);
        if (!headerRows.length) {
          alert("Header file appears to be empty.");
          return;
        }

        const headerStats = analyzeDutiesRows(headerRows, { hasBrokerage: true });
        if (headerStats.error) {
          alert(`Header file issue: ${headerStats.error}`);
          return;
        }

        let itemStats = emptyItemStats();
        if (itemFile) {
          const itemRows = await readExcelFile_AN(itemFile);
          if (!itemRows.length) {
            alert("Item file appears to be empty.");
            return;
          }
          itemStats = analyzeDutiesRows(itemRows, { hasBrokerage: false });
          if (itemStats.error) {
            alert(`Item file issue: ${itemStats.error}`);
            return;
          }
        }

        const dutyHeader = headerStats.totalDutyValue;
        const dutyItem = itemStats.totalDutyValue;
        const gstHeader = headerStats.totalGSTValue;
        const gstItem = itemStats.totalGSTValue;
        const tol = 0.01;

        const dutyMatch = dutyHeader !== null && dutyItem !== null && Math.abs(dutyHeader - dutyItem) <= tol;
        const gstMatch = gstHeader !== null && gstItem !== null && Math.abs(gstHeader - gstItem) <= tol;

        renderAnalyzerReport({
          header: headerStats,
          item: itemStats,
          compare: {
            duty: itemFile
              ? `Header ${headerStats.totalDuty} vs Item ${itemStats.totalDuty} (${dutyMatch ? "Matched" : "Not Matched"})`
              : "Item file not provided",
            gst: itemFile
              ? `Header ${headerStats.totalGST} vs Item ${itemStats.totalGST} (${gstMatch ? "Matched" : "Not Matched"})`
              : "Item file not provided"
          },
          itemAvailable: !!itemFile
        });

        if (resetAnalyzeBtn) {
          resetAnalyzeBtn.style.display = "flex";
        }
      } catch (err) {
        console.error("Analyzer error:", err);
        alert("Analyzer error: " + (err && err.message ? err.message : err));
      }
    });
  }

  if (resetAnalyzeBtn) {
    resetAnalyzeBtn.addEventListener("click", () => {
      analyzerForm.reset();
      if (headerFileName) headerFileName.textContent = "";
      if (itemFileName) itemFileName.textContent = "";
      renderAnalyzerReport({
        header: {
          totalCCNs: "",
          totalCLVS: "",
          totalLVS: "",
          totalPGA: "",
          totalDuty: "",
          totalGST: "",
          totalDutyValue: null,
          totalGSTValue: null,
          emptyBrokerageCCNs: "",
          emptyValueForDutyCCNs: "",
          gstZeroCCNs: "",
          lowValueDutyCCNs: ""
        },
        item: {
          totalCCNs: "",
          totalDuty: "",
          totalGST: "",
          totalDutyValue: null,
          totalGSTValue: null,
          emptyValueForDutyCCNs: "",
          gstZeroCCNs: "",
          lowValueDutyCCNs: ""
        },
        compare: {
          duty: "",
          gst: ""
        },
        itemAvailable: false
      });
      resetAnalyzeBtn.style.display = "none";
    });
  }
})();



