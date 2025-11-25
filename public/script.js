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

// ========== DOM ELEMENTS FOR EXCEL MERGER ==========
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

  for (let i = 0; i < files.length; i++) {
    const file = files[i];

    let sheetData = await readExcelFile(file);
    if (sheetData.length === 0) continue;

    // Remove header rows (3 for first file, 4 for rest)
    const rowsToRemove = i === 0 ? 3 : 4;
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


// =============================================================================
//  EXPORT MERGED EXCEL (browser download)
// =============================================================================
function exportMergedExcel(mergedData) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(mergedData);
  XLSX.utils.book_append_sheet(wb, ws, "Merged");

  const wbout = XLSX.write(wb, {
    bookType: "xlsx",
    type: "binary",
    compression: true,  // zip compression enabled
    WTF: false,
    cellStyles: false,
    cellNF: false,
    cellDates: false,
    bookSST: false,     // *** turn off shared strings ***
  });

  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++)
      view[i] = s.charCodeAt(i) & 0xff;
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
});

/*****************************************************
 * FIXED: Excel Modify Tool - Client Side
 * - Fixes dropzone -> file flow
 * - No variable shadowing
 * - Top-level helper functions
 * - Single Run button handler
 *****************************************************/

/* ---------------- DOM (modify UI) ---------------- */
const sourceDrop = document.getElementById("sourceDropZone");
const sourceInput = document.getElementById("sourceFileInput");
const sourceFileName = document.getElementById("sourceFileName");

const targetDrop = document.getElementById("targetDropZone");
const targetInput = document.getElementById("targetFileInput");
const targetFileName = document.getElementById("targetFileName");

const runModifyBtn = document.getElementById("runModify");
const modifyReport = document.getElementById("modifyReport");

// Keep global references in sync (set by dropzone)
let sourceFile = null;
let targetFile = null;

/* ---------------- Dropzone Setup ----------------
   Ensures both the hidden file input and the global
   variable are set when user drops or selects file.
*/
function setupDropZone(dropArea, fileInput, fileNameDisplay, storageVarSetter) {
  // dragover
  dropArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropArea.classList.add("dragover");
  });

  // dragleave
  dropArea.addEventListener("dragleave", () => {
    dropArea.classList.remove("dragover");
  });

  // drop
  dropArea.addEventListener("drop", (e) => {
    e.preventDefault();
    dropArea.classList.remove("dragover");

    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!file) return;

    // accept .xlsx only
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      alert("Please drop a .xlsx file");
      return;
    }

    // set both the file input and the global var
    try {
      // set file input files using DataTransfer so .files is populated
      const dt = new DataTransfer();
      dt.items.add(file);
      fileInput.files = dt.files;
    } catch (err) {
      // Some browsers don't allow constructing DataTransfer — still set global var
      console.warn("Could not set fileInput.files programmatically", err);
    }

    storageVarSetter(file);
    fileNameDisplay.textContent = file.name;
    console.log("Dropzone set file:", file.name);
  });

  // click to open file chooser
  dropArea.addEventListener("click", () => fileInput.click());

  // when user selects via file picker
  fileInput.addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      alert("Please select a .xlsx file");
      e.target.value = ""; // reset
      return;
    }
    storageVarSetter(file);
    fileNameDisplay.textContent = file.name;
    console.log("File input set file:", file.name);
  });
}

// initialize dropzones
setupDropZone(sourceDrop, sourceInput, sourceFileName, (f) => sourceFile = f);
setupDropZone(targetDrop, targetInput, targetFileName, (f) => targetFile = f);

/***********************************************************************
 * Final Modify Tool (2-file flow) — Exact-match dedupe (normalized digits)
 * CCN source: Column H, start at row index 6 (H6 onward)
 * Source AC: column AC (index 28) starting at row index 3
 * Source AS: column AS (index 44) starting at row index 3
 *
 * Drop-in: paste into script.js after your dropzone + DOM setup.
 ***********************************************************************/

/* DOM elements (must exist) */
const sourceInputEl = document.getElementById("sourceFileInput");
const targetInputEl = document.getElementById("targetFileInput");
const runModifyButton = document.getElementById("runModify");
const modifyReportEl = document.getElementById("modifyReport");

/* helper: normalize to digits-only and strip leading zeros */
function normalizeDigits(val) {
  if (val === undefined || val === null) return "";
  let s = String(val).trim();
  s = s.replace(/[^\d]/g, ""); // digits only
  s = s.replace(/^0+/, "");    // strip leading zeros
  return s;
}

/* read file to AoA */
async function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    if (!file) return reject(new Error("No file provided"));
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

/* safe download AoA */
function downloadAoA(rows, filename = "updated_target.xlsx", targetRowsLength = 0) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);

  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary", compression: true, bookSST: false });

  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  }
  saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), filename);
}

/* find header row (optional; we won't rely on it for CCN extraction) */
function findHeaderRowIndex(rows, fallback = 0) {
  const keywords = ["Value for Duty", "Duty", "Brokerage Total", "Gov. Sales Tax", "Value"];
  const limit = Math.min(rows.length, 12);
  for (let i = 0; i < limit; i++) {
    const row = rows[i] || [];
    for (const cell of row) {
      if (!cell) continue;
      const s = String(cell);
      for (const kw of keywords) if (s.includes(kw)) return i;
    }
  }
  return fallback;
}

/* Core function: exact-match dedupe using Column H (index 7) starting at row index 6 */
async function modifyAndDownloadExactMatch({
  sourceFileObj,
  targetFileObj,
  // fixed indices per your spec:
  ccnColumnIndex = 7,         // Column H
  ccnStartRowIndex = 5,       // start scanning H6 onward (zero-based index 6)
  sourceACStartIndex = 3,     // AC3 onward
  sourceASStartIndex = 3,     // AS3 onward
  suffixMatchLength = 0       // 0 = disabled (exact-match only)
} = {}) {
  if (!sourceFileObj || !targetFileObj) {
    alert("Please provide Source and Target files.");
    return;
  }

  // Column positions (zero-based)
   // column indices
   const COL_AC = 28;
   const COL_AS = 44;
   const COL_A  = 0;
   const COL_B  = 1;
   const COL_C  = 2;
   const COL_D  = 3;
   const COL_E  = 4;
   const COL_F  = 5;
   const COL_H  = 7;
   const COL_J  = 9;
   const COL_K  = 10;
   const COL_Q  = 16;
   const COL_R  = 17;

  try {
    // 1) read AoA for both files
    console.log("Reading target (DutiesHeader)...");
    const targetRows = await readExcelFile(targetFileObj); // AoA
    console.log("Reading source...");
    const sourceRows = await readExcelFile(sourceFileObj); // AoA

        /**********************************************************************
     * Convert EXISTING target rows J6 → Q(end) to true numbers
     * (Excel Text-to-Columns behavior)
     **********************************************************************/
    for (let r = 5; r < targetRows.length; r++) {   // J6 = row index 5
      const row = targetRows[r];
      if (!row) continue;

      for (let c = 9; c <= 16; c++) {              // J (9) → Q (16)
        let v = row[c];
        if (v === undefined || v === null) continue;

        let s = String(v).trim();
        s = s.replace(/,/g, "");   // remove thousand separators

        const num = parseFloat(s);
        if (!isNaN(num)) {
          row[c] = num;            // store as NUMBER
        }
      }
    }



    // 2) extract CCNs from Column H starting at row index 6
    const refSet = new Set();
    for (let r = ccnStartRowIndex; r < targetRows.length; r++) {
      const row = targetRows[r] || [];
      const cell = row[ccnColumnIndex];
      if (cell === undefined || cell === null) continue;
      const raw = String(cell).trim();
      if (raw === "") continue;

      // If starts with 8308 => remove only that prefix
      let cleaned = raw;
      if (raw.startsWith("8308")) cleaned = raw.replace(/^8308/, "");

      const norm = normalizeDigits(cleaned);
      if (norm) refSet.add(norm);
    }
    console.log("Extracted normalized CCNs (exact-match set size):", refSet.size);

    // 3) build source items list (from AC3 & AS3 onward)
    const sourceItems = [];
    for (let r = sourceACStartIndex; r < sourceRows.length; r++) {
      const row = sourceRows[r] || [];
      const acRaw = (row[COL_AC] === undefined || row[COL_AC] === null) ? "" : String(row[COL_AC]).trim();
      const asRaw = (row[COL_AS] === undefined || row[COL_AS] === null) ? "" : String(row[COL_AS]).trim();
      if (acRaw === "" && asRaw === "") continue;
      const acNorm = normalizeDigits(acRaw);
      sourceItems.push({ rowIndex: r, acRaw, asRaw, acNorm });
    }
    console.log("Source candidate rows:", sourceItems.length);

    // 4) Decide insert vs skip (exact-match only)
    const headerIndex = findHeaderRowIndex ? findHeaderRowIndex(targetRows, 0) : 0;
    const headerRow = targetRows[headerIndex] || [];
    const targetRowLen = Math.max(headerRow.length, COL_R + 1, COL_Q + 1, COL_J + 1);

    const insertedRows = [];
    let skippedExact = 0;

    const lastExistingRow = targetRows.length ? targetRows[targetRows.length - 1] : [];

    for (const item of sourceItems) {
      // if normalized AC empty -> current behaviour = insert. (Change to skip if desired)
      

      // exact normalized match -> skip
      if (item.acNorm && refSet.has(item.acNorm)) {
        skippedExact++;
        continue;
      }

      // not matched -> insert
      const newRow = new Array(targetRowLen).fill("");
      newRow[0] = "CLVS";
      newRow[COL_B] = item.acRaw;
      newRow[COL_H] = item.acRaw;
    
      // Copy C-F from lastExistingRow
      newRow[COL_C] = lastExistingRow[COL_C] ?? "";
      newRow[COL_D] = lastExistingRow[COL_D] ?? "";
      newRow[COL_E] = lastExistingRow[COL_E] ?? "";
      newRow[COL_F] = lastExistingRow[COL_F] ?? "";
      newRow[COL_J] = item.asRaw;
      for (let c = COL_K; c <= COL_Q; c++) newRow[c] = 0;
      newRow[COL_R] = "DDP";
    
      insertedRows.push(newRow);
    }
    console.log("Inserted rows:", insertedRows.length, "Skipped exact matches:", skippedExact);

    // 5) Build final AoA by PRESERVING original target rows exactly, then appending inserted rows
    const finalAoA = [].concat(targetRows, insertedRows);

    // // 6) (Optional) trim trailing empty cells of inserted rows only to reduce file size
    // function trimTrailingInsertedOnly(aoa) {
    //   const trimmed = [];
    //   for (let i = 0; i < aoa.length; i++) {
    //     const row = aoa[i];
    //     // If row index is in original target (i < targetRows.length) -> keep as-is
    //     if (i < targetRows.length) {
    //       trimmed.push(row);
    //       continue;
    //     }
    //     // inserted rows: trim trailing empties
    //     if (!Array.isArray(row)) {
    //       trimmed.push(row);
    //       continue;
    //     }
    //     let last = row.length - 1;
    //     while (last >= 0 && (row[last] === "" || row[last] === undefined || row[last] === null)) last--;
    //     trimmed.push(row.slice(0, last + 1));
    //   }
    //   return trimmed;
    // }
    // const compactAoA = trimTrailingInsertedOnly(finalAoA);

        /**********************************************************************
     * Convert ALL rows J6 → Q(end) to true numbers (existing + inserted)
     * (Excel Text-to-Columns behavior for entire sheet)
     **********************************************************************/
    for (let r = 5; r < finalAoA.length; r++) {  // row index 5 = Excel row 6
      const row = finalAoA[r];
      if (!row) continue;

      for (let c = 9; c <= 16; c++) {  // J(9) → Q(16)
        let v = row[c];
        if (v === undefined || v === null) continue;

        let s = String(v).trim().replace(/,/g, "");

        const num = parseFloat(s);
        if (!isNaN(num)) {
          row[c] = num;    // FINAL: numeric value
        }
      }
    }


    // 7) download
    downloadAoA(finalAoA, "updated_target.xlsx", targetRows.length);

    // 8) UI report
    if (modifyReportEl) {
      modifyReportEl.innerHTML = `
        <strong>Done</strong><br>
        Source candidates: ${sourceItems.length}<br>
        Inserted rows appended: ${insertedRows.length}<br>
        Skipped (exact): ${skippedExact}<br>
        Final rows (approx): ${finalAoA.length}
      `;
    }

    console.log("modifyAndDownloadExactMatch complete.");
  } catch (err) {
    console.error("Error in modifyAndDownloadExactMatch:", err);
    alert("Error: " + (err && err.message ? err.message : err));
  }
}

/* Wire Run button */
runModifyButton.addEventListener("click", async () => {
  const srcFile = sourceInputEl.files && sourceInputEl.files[0] ? sourceInputEl.files[0] : null;
  const tgtFile = targetInputEl.files && targetInputEl.files[0] ? targetInputEl.files[0] : null;
  console.log("Run: src =", srcFile && srcFile.name, "tgt =", tgtFile && tgtFile.name);
  await modifyAndDownloadExactMatch({ sourceFileObj: srcFile, targetFileObj: tgtFile });
});
