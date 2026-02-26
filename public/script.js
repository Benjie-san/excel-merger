//Toggle to show the section
const excelMerger = document.getElementById("excel-merger");
const DTHeaderFile = document.getElementById("modifyTool");

function showDisplay(item){
  if(item){
    excelMerger.style.display = "flex";
    DTHeaderFile.style.display = "none";
  }else{
    excelMerger.style.display = "none";
    DTHeaderFile.style.display = "flex";
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
    z: "0"       // TRUE GENERAL numeric (no currency)
  };
}

function convertColumnRangeToNumbers(aoa, startRow, colStart, colEnd, ws) {
  for (let r = startRow; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row) continue;

    for (let c = colStart; c <= colEnd; c++) {
      let v = row[c];
      if (v === undefined || v === null || v === "") continue;

      let s = String(v).trim().replace(/,/g, "");
      const num = parseFloat(s);

      if (!isNaN(num)) {
        aoa[r][c] = num;          // update AoA
        forceGeneralNumber(ws, r, c, num); // override Excel’s formatting
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

  // DutiesHeader → convert J (9) to Q (16)
  if (headerMode) {
    convertColumnRangeToNumbers(aoa, startRow, 9, 16, ws);
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
  const modifyReportEl = document.getElementById("modifyReport");
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
    ws[ref] = { t: "n", v: value, z: "0" };
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
    sourceACStartIndex = 3,    // AC3 -> index 3
    sourceASStartIndex = 3     // AS3 -> index 3
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

      // build refSet from target H (exact cleaned strings)
      const refSet = new Set();
      for (let r = ccnStartRowIndex; r < targetRows.length; r++) {
        const row = targetRows[r] || [];
        const raw = row[ccnColumnIndex];
        if (raw === undefined || raw === null) continue;
        const cleaned = cleanTargetCCN(raw);
        if (cleaned !== "") refSet.add(cleaned);
      }
      console.log("refSet size:", refSet.size);

      // build source items (keep AC as EXACT trimmed string)
      const sourceItems = [];
      for (let r = sourceACStartIndex; r < sourceRows.length; r++) {
        const row = sourceRows[r] || [];
        const acRaw = (row[COL_AC] === undefined || row[COL_AC] === null) ? "" : String(row[COL_AC]).trim();
        const asRaw = (row[COL_AS] === undefined || row[COL_AS] === null) ? "" : String(row[COL_AS]).trim();
        if (acRaw === "" && asRaw === "") continue;
        sourceItems.push({ rowIndex: r, acRaw, asRaw });
      }
      console.log("sourceItems:", sourceItems.length);

      // determine copy-from row for C-F: last non-empty row in targetRows
      const lastNonEmptyIndex = findLastNonEmptyRow(targetRows);
      const lastExistingRow = lastNonEmptyIndex >= 0 ? targetRows[lastNonEmptyIndex] : [];

      // prepare inserted rows array
      // targetRowLen: base width derived from header or safe default
      const headerIndex = 0;
      const headerRow = targetRows[headerIndex] || [];
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
      const finalAoA = targetRows.concat(insertedRows);

      // Create workbook and worksheet from finalAoA
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(finalAoA);

      // Convert J(9) -> Q(16) to numbers starting at row index 5 (Excel row 6)
      const startIndex = 5;
      for (let r = startIndex; r < finalAoA.length; r++) {
        const row = finalAoA[r] || [];
        for (let c = 9; c <= 16; c++) {
          const rawVal = row[c];
          if (rawVal === undefined || rawVal === null || rawVal === "") continue;
          const s = String(rawVal).trim().replace(/,/g, "");
          const num = parseFloat(s);
          if (!isNaN(num)) {
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
      if (modifyReportEl) {
        modifyReportEl.innerHTML = `
          <strong>Done</strong><br>
          Source candidates: ${sourceItems.length}<br>
          Inserted rows appended: ${insertedRows.length}<br>
          Skipped (exact): ${skippedExact}<br>
          Final rows (approx): ${finalAoA.length}
        `;
      }
      
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

    // Clear the report
    if(modifyReportEl) modifyReportEl.innerHTML = "";

    // Hide the reset button
    resetModifyBtn.style.display = "none";

    console.log("Modify Tool has been reset.");
  });

})(); // end IIFE
