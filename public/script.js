const progressContainer = document.getElementById("progressContainer");
const progressBar = document.getElementById("progressBar");
const uploadForm = document.getElementById("uploadForm");
const filesInput = document.getElementById("files");
const downloadLink = document.getElementById("downloadLink");
const resetBtn = document.getElementById("resetBtn");
const dropZone = document.getElementById("dropZone");
const fileList = document.getElementById("fileList");
const fileCount = document.getElementById("fileCount");
const reportDiv = document.getElementById("report"); // <-- 👈 add this to your HTML too

//fake progress 0% - 100%
function simulateProgress() {
  let progress = 0;
  progressBar.style.width = "0%";
  progressBar.textContent = "0%";

  return new Promise((resolve) => {
    const interval = setInterval(() => {
      progress += Math.floor(Math.random() * 10) + 5; // add 5–15% each step
      if (progress >= 95) {
        progress = 95; // stop at 95%, wait for server
        clearInterval(interval);
        resolve();
      }
      progressBar.style.width = progress + "%";
      progressBar.textContent = progress + "%";
    }, 300);
  });
}

function updateFileList(files) {
  fileList.innerHTML = "";
  Array.from(files).forEach((file) => {
    const li = document.createElement("li");
    li.textContent = file.name;
    fileList.appendChild(li);
  });
}

// Handle manual file selection
filesInput.addEventListener("change", () => {
  updateFileList(filesInput.files);
  fileCount.innerHTML = `${filesInput.files.length} files selected`
  console.log(filesInput.files.length)
});

// Handle drag events
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
    // Trick: update the <input type="file"> programmatically
    filesInput.files = files;
    updateFileList(files);
  }
});

// ---- MAIN MERGE + REPORT ----
uploadForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    downloadLink.style.display = "none";
    reportDiv.innerHTML = "";

    const files = document.getElementById("files").files;
    if (!files.length) {
      alert("Please select at least one file");
      return;
    }

    const formData = new FormData();
    for (let file of files) {
      formData.append("files", file);
    }

    // send checkbox state
    const addFilename = document.getElementById("addFilename").checked;
    formData.append("addFilename", addFilename);
    
    progressContainer.style.display = "block";
    await simulateProgress();
  
    const response = await fetch("/merge", {
      method: "POST",
      body: formData,
    });
  
    // if (response.ok) {
    //   const blob = await response.blob();
    //   const url = window.URL.createObjectURL(blob);
    //   //const link = document.getElementById("downloadLink");

    //   // Complete progress
    //   progressBar.style.width = "100%";
    //   progressBar.textContent = "100%";
  

    //   setTimeout(() => {
    //     console.log("rrun???")
    //     downloadLink.href = url;
    //     downloadLink.download = "merged.xlsx";
    //     downloadLink.style.display = "block";
    //     downloadLink.textContent = "Download Merged File";

    //     resetBtn.style.display = "flex"; // show reset button

    //   }, 300); // half a second delay looks smooth

    // } else {
    //   alert("Error merging files");
    //   progressContainer.style.display = "none";
    // }
    try {
      // --- Updated: expect JSON with report and base64 file ---
      const response = await fetch("/merge", {
        method: "POST",
        body: formData,
      });
  
      const result = await response.json();
      console.log("Response:", result.report);
      // Complete progress
      progressBar.style.width = "100%";
      progressBar.textContent = "100%";
      clearInterval;
  
      if (result.success) {
        // --- ✅ Display Download Link ---
        const blob = new Blob(
          [Uint8Array.from(atob(result.mergedFile), (c) => c.charCodeAt(0))],
          { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
        );
        const url = window.URL.createObjectURL(blob);
        downloadLink.href = url;
        downloadLink.download = "merged.xlsx";
        downloadLink.style.display = "block";
        downloadLink.textContent = "Download Merged File";
        resetBtn.style.display = "flex";
  
        // --- ✅ Show report summary ---
        const report = result.report;
        reportDiv.innerHTML = `
          <h3>📊 Report</h3>
          <p><b>Total Rows:</b> ${report.totalRows}</p>
          <p><b>Total Columns:</b> ${report.totalColumns}</p>
          <h4>Specific Value Counts:</h4>
          <table border="1" cellspacing="0" cellpadding="5">
            <tr><th>Value</th><th>Count</th></tr>
      
          </table>
        `;
      } else {
        alert("Error merging files");
      }
    } catch (error) {
      alert("Error: " + error.message);
      console.error("Error:", error);
    }
  });

//Start over button
resetBtn.addEventListener("click", () => {
  uploadForm.reset();                 // clears the file input
  downloadLink.style.display = "none";
  resetBtn.style.display = "none";

  // reset progress bar
  progressBar.style.width = "0%";
  progressBar.textContent = "0%";
  progressContainer.style.display = "none";
  fileList.innerHTML = ""; 
  fileCount.innerHTML = "No files selected";
  reportDiv.innerHTML = ""; 
});