const express = require('express');
const path = require('path')
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const app = express();

const port = parseInt(process.env.PORT) || process.argv[3] || 8080;

app.use(express.static(path.join(__dirname, 'public')))
  .set('views', path.join(__dirname, 'views'))
  .set('view engine', 'ejs');

app.get('/', (req, res) => {
  res.render('index');
});

app.get('/api', (req, res) => {
  res.json({"msg": "Hello world"});
});

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.post("/merge", upload.array("files"), (req, res) => {
  try {
    const mergedData = [];
    let headersAdded = false; //bool for adding filename checkbox
    const addFilename = req.body.addFilename === "true"; // checkbox state

    req.files.forEach((file, index) => {
      // Read workbook directly from memory buffer
      const workbook = xlsx.read(file.buffer, { type: "buffer" });
      const sheetName = workbook.SheetNames[0];
      const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

      if (sheetData.length === 0) return; // skip empty sheets

       // 👇 Remove top rows (4 for first file, 5 for the rest)
       let rowsToRemove = index === 0 ? 3 : 4; //changed from first and fifth row due to missing data
       let trimmedData = sheetData.slice(rowsToRemove);
      
       // 👇 Remove empty rows (rows that are entirely blank)
       trimmedData = trimmedData.filter(
         (row) => row.some((cell) => cell !== null && cell !== undefined && cell !== "")
       );

      if (trimmedData.length === 0) return;

      // Add headers only once
      if (!headersAdded) {
        if (addFilename) {
          mergedData.push(["Source File", ...trimmedData[0]]);
        } else {
          mergedData.push(trimmedData[0]);
        }
        headersAdded = true;
      }

      // Add rows
      for (let r = 1; r < trimmedData.length; r++) {
        if (addFilename) {
          mergedData.push([file.originalname, ...trimmedData[r]]);
        } else {
          mergedData.push(trimmedData[r]);
        }
      }

    });

    // Create new workbook
    const newWB = xlsx.utils.book_new();
    const newWS = xlsx.utils.aoa_to_sheet(mergedData);
    xlsx.utils.book_append_sheet(newWB, newWS, "Merged");
    const buffer = xlsx.write(newWB, { type: "buffer", bookType: "xlsx" });

    // --- Generate Report ---
    const headers = mergedData[0];
    const dataRows = mergedData.slice(1);
   
   // Basic report
    const report = {
      totalRows: dataRows.length,
      totalColumns: headers.length,
      columns: headers,
      sample: dataRows.slice(0, 5),
      columnSummary: {},
      targetValueCounts: {}, // 👈 new section
    };

    // ✅ Define target values to count
    const targetValues = [0.0175, 0.085, 0.71, 0.28];
    const tolerance = 1e-3;
    // Initialize counters
    targetValues.forEach((val) => {
      report.targetValueCounts[val] = 0;
    });
    let targetColumnIndex = 0;
    targetColumnIndex = report.columns.indexOf("Brokerage Total");
        // Count occurrences across all cells
    dataRows.forEach((row) => {
      const cell = row[targetColumnIndex];
      if (cell === undefined || cell === null) return;
    
      const cellValue = String(cell).trim();
      const num = parseFloat(cellValue);
    
      if (!isNaN(num)) {
        const rounded = parseFloat(num.toFixed(3));
        targetValues.forEach((val) => {
          if (Math.abs(rounded - val) < tolerance) {
            report.targetValueCounts[val]++;
          }
        });
      }
    });


    // Analyze each column
    headers.forEach((colName, colIndex) => {
      const values = dataRows.map((row) => row[colIndex]).filter((v) => v !== undefined && v !== null && v !== "");

      const uniqueCount = new Set(values).size;
      const numericValues = values
        .map((v) => parseFloat(v))
        .filter((n) => !isNaN(n));

      const sum = numericValues.reduce((a, b) => a + b, 0);
      const avg = numericValues.length ? sum / numericValues.length : null;

      report.columnSummary[colName] = {
        uniqueValues: uniqueCount,
        totalEntries: values.length,
        numericCount: numericValues.length,
        sum: numericValues.length ? sum : null,
        average: numericValues.length ? avg : null,
      };
    });

    // // Example: count unique values in column 2 (index 1)
    // const uniqueValues = new Set(dataRows.map((r) => r[1]));
    // report.uniqueInColumn2 = uniqueValues.size;

    // Send file to client
 
    // res.setHeader("Content-Disposition", "attachment; filename=merged.xlsx");
    // res.setHeader(
    //   "Content-Type",
    //   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    // );
    // // res.send(buffer);
    // res.json({
    //   success: true,
    //   report,
    //   mergedFile: buffer.toString("base64"), // send file as base64
    // });
     // Send both the file and report
     res.json({
      success: true,
      report,
      mergedFile: buffer.toString("base64"),
    });
  } catch (err) {
    console.error("Merge error:", err);
    res.status(500).send("Error merging files");
  }
});


app.listen(port, () => console.log(`Server running on http://localhost:${port}`));