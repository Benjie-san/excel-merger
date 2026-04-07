const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const XLSX = require("xlsx");

function readFirstSheetRows(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  return { rows: XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" }), sheetName };
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
  if (Number.isNaN(num)) return null;
  return neg ? -num : num;
}

function parseNumberZero(value) {
  const parsed = parseNumberFromCell(value);
  return parsed === null ? 0 : parsed;
}

function parseMaybeNumber(value) {
  if (value === undefined || value === null || String(value).trim() === "") return "";
  const parsed = parseNumberFromCell(value);
  return parsed === null ? String(value).trim() : parsed;
}

function isEmptyCell(v) {
  return v === undefined || v === null || String(v).trim() === "";
}

function isEmptyRow(row) {
  if (!Array.isArray(row)) return true;
  return row.every((cell) => isEmptyCell(cell));
}

function normalizeHeaderCell(v) {
  return String(v || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function findHeaderRowAndColumns(rows, specs, maxScan = 25) {
  const scanLimit = Math.min(rows.length, maxScan);
  let best = { rowIndex: -1, score: -1, indexMap: {} };

  for (let r = 0; r < scanLimit; r++) {
    const row = rows[r] || [];
    const normalized = row.map((cell) => normalizeHeaderCell(cell));
    const indexMap = {};
    let score = 0;

    specs.forEach((spec) => {
      let found = -1;
      for (let c = 0; c < normalized.length; c++) {
        const cell = normalized[c];
        if (!cell) continue;
        if (spec.matchers.some((rx) => rx.test(cell))) {
          found = c;
          break;
        }
      }
      indexMap[spec.key] = found;
      if (found !== -1) score++;
    });

    if (score > best.score) {
      best = { rowIndex: r, score, indexMap };
    }
  }

  return best;
}

function validateRequiredColumns(rows, specs, fileLabel) {
  const found = findHeaderRowAndColumns(rows, specs);
  if (found.rowIndex === -1 || found.score <= 0) {
    throw new Error(`${fileLabel}: Could not locate the header row.`);
  }

  const missing = specs
    .filter((spec) => found.indexMap[spec.key] === -1)
    .map((spec) => spec.label);

  if (missing.length) {
    throw new Error(`${fileLabel}: Missing columns: ${missing.join(", ")}`);
  }

  return found;
}

function sanitizeExcelText(value) {
  const s = String(value || "").trim();
  if (!s) return "";
  return /^[=+\-@\t\r]/.test(s) ? `'${s}` : s;
}

function trimTransaction(value) {
  const s = String(value || "").trim();
  if (s.length <= 5) return sanitizeExcelText(s);
  return sanitizeExcelText(s.slice(5));
}

function stableSortByGstDesc(records, gstKeyName) {
  return records
    .map((rec, idx) => ({ rec, idx }))
    .sort((a, b) => {
      const gstDiff = (b.rec[gstKeyName] || 0) - (a.rec[gstKeyName] || 0);
      if (Math.abs(gstDiff) > 1e-9) return gstDiff;
      return a.idx - b.idx;
    })
    .map((x) => x.rec);
}

function formatDateMMDDYYYY(value) {
  if (value === undefined || value === null) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const mm = String(value.getMonth() + 1).padStart(2, "0");
    const dd = String(value.getDate()).padStart(2, "0");
    const yyyy = String(value.getFullYear());
    return `${mm}/${dd}/${yyyy}`;
  }

  const s = String(value).trim();
  if (!s) return "";

  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return `${isoMatch[2]}/${isoMatch[3]}/${isoMatch[1]}`;
  }

  const usMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (usMatch) {
    const mm = String(usMatch[1]).padStart(2, "0");
    const dd = String(usMatch[2]).padStart(2, "0");
    return `${mm}/${dd}/${usMatch[3]}`;
  }

  const dt = new Date(s);
  if (!Number.isNaN(dt.getTime())) {
    const mm = String(dt.getMonth() + 1).padStart(2, "0");
    const dd = String(dt.getDate()).padStart(2, "0");
    const yyyy = String(dt.getFullYear());
    return `${mm}/${dd}/${yyyy}`;
  }

  return s;
}

function buildItemRecords(itemRows, itemIndexMap, dataStart) {
  const records = [];
  for (let r = dataStart; r < itemRows.length; r++) {
    const row = itemRows[r] || [];
    if (isEmptyRow(row)) continue;

    const ccn = String(row[itemIndexMap.ccn] || "").trim();
    if (!ccn) continue;

    const gst = parseNumberZero(row[itemIndexMap.gst]);
    const pst = parseNumberZero(row[itemIndexMap.pst]);
    const govSalesTax = gst + pst;

    records.push({
      transactionNumber: trimTransaction(row[itemIndexMap.transaction]),
      goodsDescription: sanitizeExcelText(row[itemIndexMap.productDescription]),
      lineNumber: sanitizeExcelText(row[itemIndexMap.cciLine]),
      countryOfOrigin: sanitizeExcelText(row[itemIndexMap.countryOfOrigin]),
      tariffTreatment: sanitizeExcelText(row[itemIndexMap.tariffTreatment]),
      partNumber: "",
      quantity: parseMaybeNumber(row[itemIndexMap.quantity]),
      port: sanitizeExcelText(row[itemIndexMap.port]),
      vendorName: sanitizeExcelText(row[itemIndexMap.vendorName]),
      valueForDuty: parseNumberZero(row[itemIndexMap.valueForDuty]),
      hs: sanitizeExcelText(row[itemIndexMap.classification]),
      dutyRate: parseMaybeNumber(row[itemIndexMap.dutyRate]),
      duty: parseNumberZero(row[itemIndexMap.customsDuty]),
      valueForTax: parseNumberZero(row[itemIndexMap.valueForTax]),
      govSalesTax,
      incoTerms: sanitizeExcelText(row[itemIndexMap.paymentTerms]),
      ccn,
      safeCcn: sanitizeExcelText(ccn)
    });
  }
  return stableSortByGstDesc(records, "govSalesTax");
}

function buildItemAggregatesByCcn(itemRecords) {
  const map = new Map();
  for (const rec of itemRecords) {
    if (!map.has(rec.ccn)) {
      map.set(rec.ccn, { valueForDuty: 0, duty: 0, govSalesTax: 0 });
    }
    const agg = map.get(rec.ccn);
    agg.valueForDuty += rec.valueForDuty;
    agg.duty += rec.duty;
    agg.govSalesTax += rec.govSalesTax;
  }
  return map;
}

function buildHeaderRecords(headerRows, headerIndexMap, dataStart, itemAggByCcn) {
  const records = [];
  for (let r = dataStart; r < headerRows.length; r++) {
    const row = headerRows[r] || [];
    if (isEmptyRow(row)) continue;

    const ccn = String(row[headerIndexMap.ccn] || "").trim();
    if (!ccn) continue;

    const agg = itemAggByCcn.get(ccn) || { valueForDuty: 0, duty: 0, govSalesTax: 0 };
    const releaseFormatted = formatDateMMDDYYYY(row[headerIndexMap.releaseDate]);
    const headerGst = parseNumberZero(row[headerIndexMap.totalGst]);
    const headerPst = parseNumberZero(row[headerIndexMap.totalProvincialSalesTax]);
    const recomputedHeaderGst = headerGst + headerPst;

    records.push({
      transactionNumber: trimTransaction(row[headerIndexMap.transaction]),
      ccn: sanitizeExcelText(ccn),
      port: sanitizeExcelText(row[headerIndexMap.portNumber]),
      shipmentDate: releaseFormatted,
      arrivalDate: releaseFormatted,
      releaseDate: releaseFormatted,
      cartons: "",
      orderNumber: sanitizeExcelText(ccn),
      otherReference: "",
      valueForDuty: agg.valueForDuty,
      duty: agg.duty,
      govSalesTax: recomputedHeaderGst,
      brokerageTotal: "",
      addlChargesTotal: 0,
      assessmentTotal: 0,
      exciseTaxTotal: 0,
      exchangeRate: 0,
      incoTerms: sanitizeExcelText(row[headerIndexMap.paymentTerms])
    });
  }
  return stableSortByGstDesc(records, "govSalesTax");
}

function buildItemOutputAoA(itemRecords, reportName, reportDate, clientName) {
  const aoa = [];
  aoa.push(["CLIENT:", clientName]);
  aoa.push(["RPT NAME:", reportName || "AWB #"]);
  aoa.push(["RPT DATE :", reportDate || ""]);
  aoa.push([]);
  aoa.push([
    "Transaction Number", "Goods Description", "Line #", "Country of Origin", "Tariff Treatment",
    "Part Number", "Quantity", "Port #", "Vendor Name", "Value for Duty", "HS #", "Duty Rate",
    "Duty", "Value for Tax", "Gov. Sales Tax", "Inco Terms", "CCN"
  ]);

  itemRecords.forEach((rec) => {
    aoa.push([
      rec.transactionNumber, rec.goodsDescription, rec.lineNumber, rec.countryOfOrigin,
      rec.tariffTreatment, rec.partNumber, rec.quantity, rec.port, rec.vendorName,
      rec.valueForDuty, rec.hs, rec.dutyRate, rec.duty, rec.valueForTax, rec.govSalesTax,
      rec.incoTerms, rec.safeCcn
    ]);
  });
  return aoa;
}

function buildHeaderOutputAoA(headerRecords, reportName, reportDate, clientName) {
  const aoa = [];
  aoa.push(["CLIENT:", clientName]);
  aoa.push(["RPT NAME:", reportName]);
  aoa.push(["RPT DATE :", reportDate]);
  aoa.push([]);
  aoa.push([
    "Transaction Number", "CCN", "Port #", "Shipment Date", "Arrival Date", "Release Date",
    "No. of Cartons", "Order Number", "Other Reference", "Value for Duty", "Duty", "Gov. Sales Tax",
    "Brokerage Total", "Addl. Charges Total", "Assessment Total", "Excise Tax Total", "Exchange Rate", "Inco Terms"
  ]);
  headerRecords.forEach((rec) => {
    aoa.push([
      rec.transactionNumber, rec.ccn, rec.port, rec.shipmentDate, rec.arrivalDate, rec.releaseDate,
      rec.cartons, rec.orderNumber, rec.otherReference, rec.valueForDuty, rec.duty, rec.govSalesTax,
      rec.brokerageTotal, rec.addlChargesTotal, rec.assessmentTotal, rec.exciseTaxTotal, rec.exchangeRate, rec.incoTerms
    ]);
  });
  return aoa;
}

function summarizeCandataAoA(aoa, type) {
  const headerRow = aoa[4] || [];
  const dataRows = aoa.slice(5);
  const indices = type === "item"
    ? { ccn: 16, vfd: 9, duty: 12, gst: 14, brokerage: null }
    : { ccn: 1, vfd: 9, duty: 10, gst: 11, brokerage: 12 };

  const summary = {
    client: String((aoa[0] || [])[1] || "").trim(),
    reportName: String((aoa[1] || [])[1] || "").trim(),
    reportDate: String((aoa[2] || [])[1] || "").trim(),
    rowCount: dataRows.length,
    headerHash: hashRows([headerRow]),
    ccnFirst5: [],
    sumVfd: 0,
    sumDuty: 0,
    sumGst: 0,
    sumBrokerage: 0
  };

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i] || [];
    if (summary.ccnFirst5.length < 5) summary.ccnFirst5.push(String(row[indices.ccn] || "").trim());
    summary.sumVfd += parseNumberZero(row[indices.vfd]);
    summary.sumDuty += parseNumberZero(row[indices.duty]);
    summary.sumGst += parseNumberZero(row[indices.gst]);
    if (indices.brokerage !== null) {
      summary.sumBrokerage += parseNumberZero(row[indices.brokerage]);
    }
  }

  summary.sumVfd = round2(summary.sumVfd);
  summary.sumDuty = round2(summary.sumDuty);
  summary.sumGst = round2(summary.sumGst);
  summary.sumBrokerage = round2(summary.sumBrokerage);
  return summary;
}

function round2(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function hashRows(rows) {
  const canonical = rows.map((row) => (row || []).map((cell) => {
    const n = parseNumberFromCell(cell);
    if (n !== null) return `#${round2(n)}`;
    return String(cell === undefined || cell === null ? "" : cell).trim();
  }));
  return crypto.createHash("sha256").update(JSON.stringify(canonical)).digest("hex");
}

function processMergeFiles(files, firstFileRows, restFileRows) {
  let mergedData = [];
  let headersAdded = false;

  files.forEach((filePath, i) => {
    let { rows } = readFirstSheetRows(filePath);
    if (!rows.length) return;

    const rowsToRemove = i === 0 ? firstFileRows : restFileRows;
    let trimmed = rows.slice(rowsToRemove);

    let firstNonEmpty = 0;
    while (firstNonEmpty < trimmed.length) {
      const row = trimmed[firstNonEmpty];
      const isEmpty = !row || row.every((c) => c === null || c === undefined || c === "");
      if (!isEmpty) break;
      firstNonEmpty++;
    }
    if (firstNonEmpty > 0) trimmed = trimmed.slice(firstNonEmpty);

    let lastNonEmpty = trimmed.length - 1;
    while (lastNonEmpty >= 0) {
      const row = trimmed[lastNonEmpty];
      const isEmpty = !row || row.every((c) => c === null || c === undefined || c === "");
      if (!isEmpty) break;
      lastNonEmpty--;
    }
    if (lastNonEmpty < trimmed.length - 1) trimmed = trimmed.slice(0, lastNonEmpty + 1);

    if (!trimmed.length) return;

    if (!headersAdded) {
      mergedData.push(trimmed[0]);
      headersAdded = true;
    }

    const startRow = i === 0 ? 1 : 0;
    for (let r = startRow; r < trimmed.length; r++) {
      mergedData.push(trimmed[r]);
    }
  });

  return mergedData;
}

function compareObjects(actual, expected) {
  const issues = [];
  const keys = Object.keys(expected);
  keys.forEach((key) => {
    const av = actual[key];
    const ev = expected[key];
    const same = JSON.stringify(av) === JSON.stringify(ev);
    if (!same) {
      issues.push(`${key}: actual=${JSON.stringify(av)} expected=${JSON.stringify(ev)}`);
    }
  });
  return issues;
}

function runCandataRegression(rootDir) {
  const dir = path.join(rootDir, "CandataToGetsFormat");
  const headerIn = path.join(dir, "17WS114V6GF79_DutiesHeader_INPUT.xlsx");
  const itemIn = path.join(dir, "17WS114V6GF79_DutiesItem_INPUT.xlsx");
  const headerOutExpected = path.join(dir, "17WS114V6GF79_DutiesHeader_OUTPUT.xlsx");
  const itemOutExpected = path.join(dir, "17WS114V6GF79_DutiesItem_OUTPUT.xlsx");

  const headerSpecs = [
    { key: "transaction", label: "Transaction Number", matchers: [/^transaction number$/i] },
    { key: "ccn", label: "Cargo Control Number", matchers: [/cargo control number/i, /^ccn$/i] },
    { key: "portNumber", label: "Port Number", matchers: [/^port number$/i, /^port #$/i] },
    { key: "directShipDate", label: "Direct Ship Date", matchers: [/direct ship date/i] },
    { key: "etaDate", label: "ETA Date", matchers: [/^eta date$/i, /\barrival date\b/i] },
    { key: "releaseDate", label: "Release Date", matchers: [/^release date$/i] },
    { key: "orderNumber", label: "Order Number", matchers: [/^order number$/i] },
    { key: "totalValueForDuty", label: "Total Value For Duty (CAD)", matchers: [/total value for duty/i] },
    { key: "totalCustomsDuties", label: "Total Customs Duties (CAD)", matchers: [/total customs duties/i] },
    { key: "totalGst", label: "Total GST (CAD)", matchers: [/^total gst/i] },
    { key: "totalProvincialSalesTax", label: "Total Provincial Sales Tax (CAD)", matchers: [/total provincial sales tax/i] },
    { key: "paymentTerms", label: "Payment Terms", matchers: [/^payment terms$/i, /^inco terms$/i] },
    { key: "billOfLading", label: "Bill of Lading", matchers: [/bill of lading/i, /\bawb\b/i] }
  ];
  const itemSpecs = [
    { key: "transaction", label: "Transaction Number", matchers: [/^transaction number$/i] },
    { key: "productDescription", label: "Product Description", matchers: [/product description/i, /goods description/i] },
    { key: "cciLine", label: "CCI Line#", matchers: [/cci line#?/i, /\bline #\b/i] },
    { key: "countryOfOrigin", label: "Country of Origin", matchers: [/country of origin/i] },
    { key: "tariffTreatment", label: "Tariff Treatment", matchers: [/tariff treatment/i] },
    { key: "quantity", label: "Quantity", matchers: [/^quantity$/i] },
    { key: "port", label: "Port Number", matchers: [/^port number$/i, /^port #$/i] },
    { key: "vendorName", label: "Vendor Name", matchers: [/vendor name/i] },
    { key: "valueForDuty", label: "Value For Duty (CAD)", matchers: [/value for duty/i] },
    { key: "classification", label: "Classification", matchers: [/^classification$/i, /^hs #$/i] },
    { key: "dutyRate", label: "Duty Rate", matchers: [/^duty rate$/i] },
    { key: "customsDuty", label: "Customs Duty (CAD)", matchers: [/customs duty/i, /^duty$/i] },
    { key: "valueForTax", label: "Value for Tax (CAD)", matchers: [/value for tax/i] },
    { key: "gst", label: "GST (CAD)", matchers: [/^gst/i, /gov\.?\s*sales/i] },
    { key: "pst", label: "Provincial Sales Tax (CAD)", matchers: [/provincial sales tax/i, /\bpst\b/i] },
    { key: "paymentTerms", label: "Payment Terms", matchers: [/^payment terms$/i, /^inco terms$/i] },
    { key: "ccn", label: "Cargo Control Number", matchers: [/cargo control number/i, /^ccn$/i] },
    { key: "billOfLading", label: "Bill of Lading", matchers: [/bill of lading/i, /\bawb\b/i] }
  ];

  const { rows: headerRows } = readFirstSheetRows(headerIn);
  const { rows: itemRows } = readFirstSheetRows(itemIn);
  const headerFound = validateRequiredColumns(headerRows, headerSpecs, "DutiesHeader");
  const itemFound = validateRequiredColumns(itemRows, itemSpecs, "Candata Duties Item");
  const itemRecords = buildItemRecords(itemRows, itemFound.indexMap, itemFound.rowIndex + 1);
  const itemAggByCcn = buildItemAggregatesByCcn(itemRecords);
  const headerRecords = buildHeaderRecords(headerRows, headerFound.indexMap, headerFound.rowIndex + 1, itemAggByCcn);

  let firstBOL = "";
  let firstReleaseDate = "";
  for (let r = headerFound.rowIndex + 1; r < headerRows.length; r++) {
    const row = headerRows[r] || [];
    if (!firstBOL) {
      const bol = sanitizeExcelText(row[headerFound.indexMap.billOfLading]);
      if (bol) firstBOL = bol;
    }
    if (!firstReleaseDate) {
      const releaseDate = row[headerFound.indexMap.releaseDate];
      const releaseFormatted = formatDateMMDDYYYY(releaseDate);
      if (releaseFormatted) firstReleaseDate = releaseFormatted;
    }
    if (firstBOL && firstReleaseDate) break;
  }

  const reportName = `AWB # ${firstBOL || ""}`.trim();
  const actualHeaderAoA = buildHeaderOutputAoA(headerRecords, reportName, firstReleaseDate || "", "");
  const actualItemAoA = buildItemOutputAoA(itemRecords, reportName, firstReleaseDate || "", "");
  const expectedHeaderAoA = readFirstSheetRows(headerOutExpected).rows;
  const expectedItemAoA = readFirstSheetRows(itemOutExpected).rows;

  const actualHeaderSummary = summarizeCandataAoA(actualHeaderAoA, "header");
  const expectedHeaderSummary = summarizeCandataAoA(expectedHeaderAoA, "header");
  const actualItemSummary = summarizeCandataAoA(actualItemAoA, "item");
  const expectedItemSummary = summarizeCandataAoA(expectedItemAoA, "item");

  // Known intended rule change in app: report date is now based on Release Date.
  // Legacy sample outputs may still reflect older date source.
  delete actualHeaderSummary.reportDate;
  delete expectedHeaderSummary.reportDate;
  delete actualItemSummary.reportDate;
  delete expectedItemSummary.reportDate;
  delete actualHeaderSummary.client;
  delete expectedHeaderSummary.client;
  delete actualItemSummary.client;
  delete expectedItemSummary.client;
  delete actualHeaderSummary.sumBrokerage;
  delete expectedHeaderSummary.sumBrokerage;

  const explicitIssues = [];
  const headerClientCell = String((actualHeaderAoA[0] || [])[1] || "").trim();
  const itemClientCell = String((actualItemAoA[0] || [])[1] || "").trim();
  if (headerClientCell !== "") {
    explicitIssues.push(`Header A2 client should be blank, got ${JSON.stringify(headerClientCell)}`);
  }
  if (itemClientCell !== "") {
    explicitIssues.push(`Item A2 client should be blank, got ${JSON.stringify(itemClientCell)}`);
  }

  const headerDataRows = actualHeaderAoA.slice(5);
  for (let i = 0; i < headerDataRows.length; i++) {
    const brokerageCell = String((headerDataRows[i] || [])[12] || "").trim();
    if (brokerageCell !== "") {
      explicitIssues.push(`Header brokerage cell should be blank at data row ${i + 6}, got ${JSON.stringify(brokerageCell)}`);
      break;
    }
  }

  return {
    headerIssues: compareObjects(actualHeaderSummary, expectedHeaderSummary).concat(explicitIssues),
    itemIssues: compareObjects(actualItemSummary, expectedItemSummary)
  };
}

function runMergeRegression(rootDir) {
  const dir = path.join(rootDir, "MERGE");
  const expectedPath = path.join(dir, "merged.xlsx");
  const mergeInputFiles = fs.readdirSync(dir)
    .filter((name) => name.toLowerCase().endsWith(".xlsx"))
    .filter((name) => name.toLowerCase() !== "merged.xlsx")
    .map((name) => path.join(dir, name))
    .sort((a, b) => a.localeCompare(b));

  const actual = processMergeFiles(mergeInputFiles, 4, 5);
  if (!fs.existsSync(expectedPath)) {
    return {
      mode: "smoke-only",
      rowCount: actual.length,
      columnCount: actual[0] ? actual[0].length : 0,
      hash: hashRows(actual)
    };
  }

  const expected = readFirstSheetRows(expectedPath).rows;

  const actualHash = hashRows(actual);
  const expectedHash = hashRows(expected);
  return {
    mode: "expected-compare",
    rowCountMatch: actual.length === expected.length,
    hashMatch: actualHash === expectedHash,
    actualRows: actual.length,
    expectedRows: expected.length,
    actualHash,
    expectedHash
  };
}

function main() {
  const rootDir = path.join(process.cwd(), "Test files");
  const candata = runCandataRegression(rootDir);
  const merge = runMergeRegression(rootDir);

  let failed = false;
  console.log("=== Regression Smoke Results ===");

  if (candata.headerIssues.length === 0) {
    console.log("Candata Header: PASS");
  } else {
    failed = true;
    console.log("Candata Header: FAIL");
    candata.headerIssues.forEach((issue) => console.log(`  - ${issue}`));
  }

  if (candata.itemIssues.length === 0) {
    console.log("Candata Item: PASS");
  } else {
    failed = true;
    console.log("Candata Item: FAIL");
    candata.itemIssues.forEach((issue) => console.log(`  - ${issue}`));
  }

  if (merge.mode === "smoke-only") {
    if (merge.rowCount > 1 && merge.columnCount > 0) {
      console.log("Merge Module: PASS (smoke-only, no expected merged.xlsx found)");
      console.log(`  - Rows=${merge.rowCount}, Cols=${merge.columnCount}`);
    } else {
      failed = true;
      console.log("Merge Module: FAIL (smoke-only)");
      console.log(`  - Rows=${merge.rowCount}, Cols=${merge.columnCount}`);
    }
  } else if (merge.rowCountMatch && merge.hashMatch) {
    console.log("Merge Module: PASS");
  } else {
    failed = true;
    console.log("Merge Module: FAIL");
    console.log(`  - Rows actual=${merge.actualRows} expected=${merge.expectedRows}`);
    console.log(`  - Hash actual=${merge.actualHash}`);
    console.log(`  - Hash expected=${merge.expectedHash}`);
  }

  if (failed) {
    process.exitCode = 1;
  }
}

main();
