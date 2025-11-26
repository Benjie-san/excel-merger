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

// ========== DOM ELEMENTS ==========
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

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });

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

/* ===========================================================================
   EXCEL MODIFY TOOL (ISOLATED) — safe, no global collisions
   - Wraps all code in an IIFE to prevent overriding merger helpers
   - Uses unique function names (suffix _MOD)
   - No accidental global writes
   ========================================================================== */
   (function ExcelModifyModule() {
    // DOM (modify UI)
    const sourceDrop_MOD = document.getElementById("sourceDropZone");
    const sourceInput_MOD = document.getElementById("sourceFileInput");
    const sourceFileName_MOD = document.getElementById("sourceFileName");
  
    const targetDrop_MOD = document.getElementById("targetDropZone");
    const targetInput_MOD = document.getElementById("targetFileInput");
    const targetFileName_MOD = document.getElementById("targetFileName");
  
    const runModifyBtn_MOD = document.getElementById("runModify");
    const modifyReport_MOD = document.getElementById("modifyReport");
  
    // local references (no globals)
    let sourceFile_MOD = null;
    let targetFile_MOD = null;
  
    // Setup dropzone (isolated)
    function setupDropZone_MOD(dropArea, fileInput, fileNameDisplay, storageSetter) {
      dropArea.addEventListener("dragover", (e) => { e.preventDefault(); dropArea.classList.add("dragover"); });
      dropArea.addEventListener("dragleave", () => dropArea.classList.remove("dragover"));
      dropArea.addEventListener("drop", (e) => {
        e.preventDefault(); dropArea.classList.remove("dragover");
        const file = e.dataTransfer.files && e.dataTransfer.files[0];
        if (!file) return;
        if (!file.name.toLowerCase().endsWith(".xlsx")) { alert("Please drop a .xlsx file"); return; }
        try { const dt = new DataTransfer(); dt.items.add(file); fileInput.files = dt.files; } catch (err) { console.warn("DT not allowed", err); }
        storageSetter(file);
        fileNameDisplay.textContent = file.name;
      });
      dropArea.addEventListener("click", () => fileInput.click());
      fileInput.addEventListener("change", (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        if (!file.name.toLowerCase().endsWith(".xlsx")) { alert("Please select a .xlsx file"); e.target.value = ""; return; }
        storageSetter(file);
        fileNameDisplay.textContent = file.name;
      });
    }
  
    setupDropZone_MOD(sourceDrop_MOD, sourceInput_MOD, sourceFileName_MOD, (f) => { sourceFile_MOD = f; });
    setupDropZone_MOD(targetDrop_MOD, targetInput_MOD, targetFileName_MOD, (f) => { targetFile_MOD = f; });
  
    // helper: normalize digits
    function normalizeDigits_MOD(val) {
      if (val === undefined || val === null) return "";
      let s = String(val).trim();
      s = s.replace(/[^\d]/g, "");
      s = s.replace(/^0+/, "");
      return s;
    }
  
    // read file -> AoA (unique name)
    async function readExcelFile_MOD(file) {
      return new Promise((resolve, reject) => {
        if (!file) return resolve([]); // return empty rather than reject (safer)
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
  
    // download AoA (unique name)
    function downloadAoA_MOD(rows, filename = "updated_target.xlsx") {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary", compression: true, bookSST: false });
      function s2ab(s) { const buf = new ArrayBuffer(s.length); const view = new Uint8Array(buf); for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff; return buf; }
      saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), filename);
    }
  
    // Core modify function (unique name)
    async function modifyAndDownloadExactMatch_MOD({
      sourceFileObj,
      targetFileObj,
      ccnColumnIndex = 7,
      ccnStartRowIndex = 5,
      sourceACStartIndex = 3,
      sourceASStartIndex = 3
    } = {}) {
      try {
        if (!sourceFileObj || !targetFileObj) { alert("Please provide Source and Target files."); return; }
  
        const COL_AC = 28, COL_AS = 44;
        const COL_A = 0, COL_B = 1, COL_C = 2, COL_D = 3, COL_E = 4, COL_F = 5, COL_H = 7;
        const COL_J = 9, COL_K = 10, COL_Q = 16, COL_R = 17;
  
        // read files (use local read)
        const tgtRows = await readExcelFile_MOD(targetFileObj);
        const srcRows = await readExcelFile_MOD(sourceFileObj);
  
        const targetRows = Array.isArray(tgtRows) ? tgtRows : [];
        const sourceRows = Array.isArray(srcRows) ? srcRows : [];
  
        // robust column width
        let maxCols = 0;
        for (const r of targetRows) if (Array.isArray(r)) maxCols = Math.max(maxCols, r.length);
        const targetRowLen = Math.max(maxCols, COL_R + 1, COL_Q + 1, COL_J + 1, 30);
  
        // convert existing target J6->Q to numbers (in-place)
        for (let r = 5; r < targetRows.length; r++) {
          const row = targetRows[r] || [];
          for (let c = 9; c <= 16; c++) {
            let v = row[c];
            if (v === undefined || v === null || v === "") continue;
            let s = String(v).trim().replace(/,/g, "");
            const num = parseFloat(s);
            if (!isNaN(num)) row[c] = num;
          }
        }
  
        // extract CCNs
        const refSet = new Set();
        for (let r = ccnStartRowIndex; r < targetRows.length; r++) {
          const row = targetRows[r] || [];
          const cell = row[ccnColumnIndex];
          if (cell === undefined || cell === null) continue;
          const raw = String(cell).trim();
          if (!raw) continue;
          let cleaned = raw;
          if (raw.startsWith("8308")) cleaned = raw.replace(/^8308/, "");
          const norm = normalizeDigits_MOD(cleaned);
          if (norm) refSet.add(norm);
        }
  
        // source items
        const sourceItems = [];
        for (let r = sourceACStartIndex; r < sourceRows.length; r++) {
          const row = sourceRows[r] || [];
          const acRaw = (row[COL_AC] === undefined || row[COL_AC] === null) ? "" : String(row[COL_AC]).trim();
          const asRaw = (row[COL_AS] === undefined || row[COL_AS] === null) ? "" : String(row[COL_AS]).trim();
          if (acRaw === "" && asRaw === "") continue;
          const acNorm = normalizeDigits_MOD(acRaw);
          sourceItems.push({ rowIndex: r, acRaw, asRaw, acNorm });
        }
  
        // last existing row for copying C-F
        const lastExistingRow = targetRows.length ? targetRows[targetRows.length - 1] : [];
  
        // build inserted rows
        const insertedRows = [];
        let skippedExact = 0;
        for (const item of sourceItems) {
          if (item.acNorm && refSet.has(item.acNorm)) { skippedExact++; continue; }
  
          const newRow = new Array(targetRowLen).fill("");
          newRow[COL_A] = "CLVS";
          newRow[COL_B] = item.acRaw;
          newRow[COL_H] = item.acRaw;
          newRow[COL_C] = lastExistingRow[COL_C] ?? "";
          newRow[COL_D] = lastExistingRow[COL_D] ?? "";
          newRow[COL_E] = lastExistingRow[COL_E] ?? "";
          newRow[COL_F] = lastExistingRow[COL_F] ?? "";
          newRow[COL_J] = item.asRaw;
          for (let c = COL_K; c <= COL_Q; c++) newRow[c] = 0;
          newRow[COL_R] = "DDP";
          insertedRows.push(newRow);
        }
  
        // final AoA and convert J6->Q both existing+inserted
        const finalAoA = targetRows.concat(insertedRows);
        for (let r = 5; r < finalAoA.length; r++) {
          const row = finalAoA[r] || [];
          for (let c = 9; c <= 16; c++) {
            let v = row[c];
            if (v === undefined || v === null || v === "") continue;
            let s = String(v).trim().replace(/,/g, "");
            const num = parseFloat(s);
            if (!isNaN(num)) row[c] = num;
          }
        }
  
        // download
        downloadAoA_MOD(finalAoA, "updated_target.xlsx");
  
        if (modifyReport_MOD) {
          modifyReport_MOD.innerHTML = `
            <strong>Done</strong><br>
            Source candidates: ${sourceItems.length}<br>
            Inserted rows appended: ${insertedRows.length}<br>
            Skipped (exact): ${skippedExact}<br>
            Final rows (approx): ${finalAoA.length}
          `;
        }
        console.log("modify complete (isolated).");
      } catch (err) {
        console.error("modify error (isolated):", err);
        alert("Error: " + (err && err.message ? err.message : err));
      }
    }
  
    // wire button (use the isolated files)
    runModifyBtn_MOD.addEventListener("click", async () => {
      const src = sourceFile_MOD || (sourceInput_MOD.files && sourceInput_MOD.files[0]) || null;
      const tgt = targetFile_MOD || (targetInput_MOD.files && targetInput_MOD.files[0]) || null;
      console.log("Run modify (isolated), src:", src && src.name, "tgt:", tgt && tgt.name);
      await modifyAndDownloadExactMatch_MOD({ sourceFileObj: src, targetFileObj: tgt });
    });
  
  })(); // end IIFE
  