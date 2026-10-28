const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const XLSX = require("xlsx");
const CandataWorkflow = require("../public/candata-workflow.js");

function readFirstSheetRows(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  return { rows: XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" }), sheetName };
}

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function cloneRows(rows) {
  return (rows || []).map((row) => (Array.isArray(row) ? row.slice() : row));
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

function normalizeClientName(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
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
    .filter((spec) => !spec.optional && found.indexMap[spec.key] === -1)
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

function routeProvinceTax(provinceValue, provincialSalesTaxValue, hstValue) {
  return CandataWorkflow.routeProvinceTax(provinceValue, parseNumberZero(provincialSalesTaxValue), parseNumberZero(hstValue));
}

function transactionLookupKeys(ccnValue, transactionValue) {
  const ccn = String(ccnValue || "").trim().toLowerCase();
  const transaction = trimTransaction(transactionValue).toLowerCase();
  const keys = [];
  if (ccn && transaction) keys.push(`pair:${ccn}|${transaction}`);
  if (ccn) keys.push(`ccn:${ccn}`);
  if (transaction) keys.push(`transaction:${transaction}`);
  return keys;
}

function buildPaymentTermsLookup(headerRows, headerIndexMap, dataStart) {
  const lookup = new Map();
  for (let r = dataStart; r < headerRows.length; r++) {
    const row = headerRows[r] || [];
    if (isEmptyRow(row)) continue;
    const terms = sanitizeExcelText(row[headerIndexMap.paymentTerms]);
    if (!terms) continue;
    transactionLookupKeys(row[headerIndexMap.ccn], row[headerIndexMap.transaction]).forEach((key) => {
      if (!lookup.has(key)) lookup.set(key, terms);
    });
  }
  return lookup;
}

function findPaymentTerms(lookup, ccn, transaction) {
  for (const key of transactionLookupKeys(ccn, transaction)) {
    if (lookup.has(key)) return lookup.get(key);
  }
  return "";
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

function buildItemRecords(itemRows, itemIndexMap, dataStart, paymentTermsLookup) {
  const records = [];
  for (let r = dataStart; r < itemRows.length; r++) {
    const row = itemRows[r] || [];
    if (isEmptyRow(row)) continue;

    const ccn = String(row[itemIndexMap.ccn] || "").trim();

    const gst = parseNumberZero(row[itemIndexMap.gst]);
    const provinceTax = routeProvinceTax(row[itemIndexMap.province], row[itemIndexMap.pst], row[itemIndexMap.hst]);

    records.push({
      transactionNumber: trimTransaction(row[itemIndexMap.transaction]),
      goodsDescription: sanitizeExcelText(row[itemIndexMap.productDescription]),
      lineNumber: sanitizeExcelText(row[itemIndexMap.cciLine]),
      countryOfOrigin: sanitizeExcelText(row[itemIndexMap.countryOfOrigin]),
      tariffTreatment: sanitizeExcelText(row[itemIndexMap.tariffTreatment]),
      partNumber: sanitizeExcelText(row[itemIndexMap.productNumber]),
      quantity: parseMaybeNumber(row[itemIndexMap.quantity]),
      port: sanitizeExcelText(row[itemIndexMap.port]),
      vendorName: sanitizeExcelText(row[itemIndexMap.vendorName]),
      valueForDuty: parseNumberZero(row[itemIndexMap.valueForDuty]),
      hs: sanitizeExcelText(row[itemIndexMap.classification]),
      dutyRate: parseMaybeNumber(row[itemIndexMap.dutyRate]),
      duty: parseNumberZero(row[itemIndexMap.customsDuty]),
      valueForTax: parseNumberZero(row[itemIndexMap.valueForTax]),
      gst,
      hst: provinceTax.hst,
      pst: provinceTax.pst,
      sima: parseNumberZero(row[itemIndexMap.sima]),
      surtax: parseNumberZero(row[itemIndexMap.surtax]),
      incoTerms: findPaymentTerms(paymentTermsLookup, ccn, row[itemIndexMap.transaction]),
      ccn,
      safeCcn: sanitizeExcelText(ccn)
    });
  }
  return records;
}

function buildHeaderRecords(headerRows, headerIndexMap, dataStart) {
  const entries = [];
  for (let r = dataStart; r < headerRows.length; r++) {
    const row = headerRows[r] || [];
    if (isEmptyRow(row)) continue;

    const ccn = String(row[headerIndexMap.ccn] || "").trim();
    if (!ccn) continue;

    const transaction = trimTransaction(row[headerIndexMap.transaction]);
    const provinceTax = routeProvinceTax(row[headerIndexMap.province], row[headerIndexMap.totalProvincialSalesTax], row[headerIndexMap.hst]);
    const headerDuty = parseNumberZero(row[headerIndexMap.totalCustomsDuties]);
    const headerGst = parseNumberZero(row[headerIndexMap.totalGst]);
    const releaseFormatted = formatDateMMDDYYYY(row[headerIndexMap.releaseDate]);
    const headerValueForDuty = parseNumberZero(row[headerIndexMap.totalValueForDuty]);

    const record = {
      transactionNumber: transaction,
      ccn: sanitizeExcelText(ccn),
      port: sanitizeExcelText(row[headerIndexMap.portNumber]),
      shipmentDate: formatDateMMDDYYYY(row[headerIndexMap.directShipDate]),
      arrivalDate: formatDateMMDDYYYY(row[headerIndexMap.etaDate]),
      releaseDate: releaseFormatted,
      cartons: "",
      orderNumber: sanitizeExcelText(row[headerIndexMap.orderNumber]),
      otherReference: "",
      valueForDuty: headerValueForDuty,
      duty: headerDuty,
      gst: headerGst,
      hst: provinceTax.hst,
      pst: provinceTax.pst,
      sima: parseNumberZero(row[headerIndexMap.sima]),
      surtax: parseNumberZero(row[headerIndexMap.surtax]),
      brokerageTotal: 0,
      addlChargesTotal: 0,
      exciseTaxTotal: 0,
      exchangeRate: 1,
      incoTerms: sanitizeExcelText(row[headerIndexMap.paymentTerms])
    };
    entries.push({
      row: r + 1, ccn, transaction: String(row[headerIndexMap.transaction] || "").trim(),
      province: String(row[headerIndexMap.province] || "").trim(),
      duty: headerDuty, gst: headerGst,
      pst: parseNumberZero(row[headerIndexMap.totalProvincialSalesTax]),
      hst: parseNumberZero(row[headerIndexMap.hst]), record
    });
  }
  return CandataWorkflow.buildHeaderGroups(entries).records;
}

function buildItemOutputAoA(itemRecords, reportName, reportDate, clientName) {
  const aoa = [];
  aoa.push(["CLIENT:", clientName]);
  aoa.push(["RPT NAME:", reportName || "AWB #"]);
  aoa.push(["RPT DATE :", reportDate || ""]);
  aoa.push([
    "Transaction Number", "Goods Description", "Line #", "Country of Origin", "Tariff Treatment",
    "Part Number", "Quantity", "Port #", "Vendor Name", "Value for Duty", "HS #", "Duty Rate",
    "Duty", "Value for Tax", "GST", "HST", "PST", "SIMA", "Surtax", "Inco Terms"
  ]);

  itemRecords.forEach((rec) => {
    aoa.push([
      rec.transactionNumber, rec.goodsDescription, rec.lineNumber, rec.countryOfOrigin,
      rec.tariffTreatment, rec.partNumber, rec.quantity, rec.port, rec.vendorName,
      rec.valueForDuty, rec.hs, rec.dutyRate, rec.duty, rec.valueForTax, rec.gst, rec.hst,
      rec.pst, rec.sima, rec.surtax, rec.incoTerms
    ]);
  });
  return aoa;
}

function buildHeaderOutputAoA(headerRecords, reportName, reportDate, clientName) {
  const aoa = [];
  aoa.push(["CLIENT:", clientName]);
  aoa.push(["RPT NAME:", reportName]);
  aoa.push(["RPT DATE :", reportDate]);
  aoa.push([
    "Transaction Number", "CCN", "Port #", "Shipment Date", "Arrival Date", "Release Date",
    "No. of Cartons", "Order Number", "Other Reference", "Value for Duty", "Duty", "GST", "HST", "PST",
    "SIMA", "Surtax", "Brokerage Total", "Addl. Charges Total", "Excise Tax Total", "Exchange Rate", "Inco Terms"
  ]);
  headerRecords.forEach((rec) => {
    aoa.push([
      rec.transactionNumber, rec.ccn, rec.port, rec.shipmentDate, rec.arrivalDate, rec.releaseDate,
      rec.cartons, rec.orderNumber, rec.otherReference, rec.valueForDuty, rec.duty, rec.gst, rec.hst,
      rec.pst, rec.sima, rec.surtax, rec.brokerageTotal, rec.addlChargesTotal, rec.exciseTaxTotal,
      rec.exchangeRate, rec.incoTerms
    ]);
  });
  return aoa;
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

function findHeaderColumnIndex(row, expectedLabel) {
  return (row || []).findIndex((cell) => normalizeHeaderCell(cell) === normalizeHeaderCell(expectedLabel));
}

function runExpandedDtTaxSchemaRegression(workflowModule, brokerageRates) {
  const issues = [];
  const headerLabels = [
    "Transaction Number", "CCN", "Port #", "Shipment Date", "Arrival Date", "Release Date",
    "No. of Cartons", "Order Number", "Other Reference", "Value for Duty", "Duty", "GST",
    "HST", "PST", "SIMA", "Surtax", "Brokerage Total", "Addl. Charges Total",
    "Excise Tax Total", "Exchange Rate", "Inco Terms"
  ];
  const itemLabels = [
    "Transaction Number", "Goods Description", "Line #", "Country of Origin", "Tariff Treatment",
    "Part Number", "Quantity", "Port #", "Vendor Name", "Value for Duty", "HS #", "Duty Rate",
    "Duty", "Value for Tax", "GST", "HST", "PST", "SIMA", "Surtax", "Inco Terms"
  ];
  const headerRows = [
    ["CLIENT:", "SF EXPRESS"], ["RPT NAME:", "TEST"], ["RPT DATE :", "09/24/2026"], headerLabels,
    ["LV-1", "CCN-1", "0497", "09/24/2026", "09/24/2026", "09/24/2026", "", "CCN-1", "", 10, 1, 2, 3, 0, 0, 0, 0.28, 0, 0, 1, "DDP"]
  ];
  const itemRows = [
    ["CLIENT:", "SF EXPRESS"], ["RPT NAME:", "TEST"], ["RPT DATE :", "09/24/2026"], itemLabels,
    ["LV-1", "ITEM", 1, "CN", "", "", 1, "0497", "VENDOR", 10, "", 0, 1, 13, 2, 3, 0, 0, 0, "DDP"]
  ];
  const sourceRows = [[], [], new Array(45).fill("")];
  sourceRows[2][28] = "NEW-CCN";
  sourceRows[2][44] = "12.34";
  const metadata = { client: "SF EXPRESS", reportName: "TEST", reportDate: "09/24/2026" };

  try {
    const prepared = workflowModule.prepareHeaderRowsForModify({ targetRows: headerRows, metadata });
    const modified = workflowModule.applyBrokerageAutomation({ sourceRows, preparedHeader: prepared, metadata, brokerageRates });
    const item = workflowModule.prepareItemRowsWithCcn({ itemRows, preparedHeader: modified, metadata });
    const summary = workflowModule.summarizeDtOutputs({ headerRows: modified.rows, itemRows: item.rows });
    const header = modified.rows[modified.headerRowIndex];
    const generated = modified.rows.find((row, index) => index > modified.headerRowIndex && row[header.indexOf("CCN")] === "NEW-CCN");

    if (!generated) issues.push("Expanded tax schema did not generate the new Header row.");
    if (generated) {
      for (const label of ["Duty", "GST", "HST", "PST", "SIMA", "Surtax", "Addl. Charges Total", "Excise Tax Total", "Exchange Rate"]) {
        if (generated[header.indexOf(label)] !== 0) issues.push(`Generated ${label} should be numeric zero.`);
      }
      if (generated[header.indexOf("Inco Terms")] !== "DDP") issues.push("Generated Inco Terms should be DDP in the expanded Header schema.");
    }
    if (summary.header.totalHstValue !== 3 || summary.item.totalHstValue !== 3 || summary.compare.hstMatch !== true) {
      issues.push(`Expanded HST totals did not map between Header and Item: ${JSON.stringify(summary.compare)}`);
    }
    for (const key of ["pstMatch", "simaMatch", "surtaxMatch"]) {
      if (summary.compare[key] !== true) issues.push(`Expanded tax comparison ${key} should match.`);
    }
    if (summary.validation.header.error || summary.validation.item.error) {
      issues.push(`Expanded GST alias validation failed: ${summary.validation.header.error || summary.validation.item.error}`);
    }
    const headerFieldCounts = summary.validation.header.fieldCounts;
    const itemFieldCounts = summary.validation.item.fieldCounts;
    if (headerFieldCounts.HST?.zero !== 1 || headerFieldCounts.PST?.zero !== 2 || headerFieldCounts.SIMA?.zero !== 2 || headerFieldCounts.Surtax?.zero !== 2) {
      issues.push(`Expanded Header zero breakdown is incorrect: ${JSON.stringify(headerFieldCounts)}`);
    }
    if (itemFieldCounts.HST?.zero !== 0 || itemFieldCounts.PST?.zero !== 1 || itemFieldCounts.SIMA?.zero !== 1 || itemFieldCounts.Surtax?.zero !== 1) {
      issues.push(`Expanded Item zero breakdown is incorrect: ${JSON.stringify(itemFieldCounts)}`);
    }
    for (const label of ["HST", "PST", "SIMA", "Surtax"]) {
      if (headerFieldCounts[label]?.blank !== 0 || itemFieldCounts[label]?.blank !== 0) {
        issues.push(`${label} should have no blank findings in the expanded-schema fixture.`);
      }
    }
  } catch (err) {
    issues.push(`Expanded D/T tax schema threw: ${err.message}`);
  }

  return issues;
}

function runDtHeaderWorkflowRegression(rootDir) {
  const issues = [];
  let workflowModule;

  try {
    workflowModule = require("../public/dt-header-workflow.js");
  } catch (err) {
    return {
      issues: [`Unable to import ../public/dt-header-workflow.js: ${err.message}`]
    };
  }

  const {
    detectHeaderRowIndex,
    prepareHeaderRowsForModify,
    prepareItemRowsWithCcn,
    applyBrokerageAutomation,
    summarizeDtOutputs
  } = workflowModule || {};

  if (typeof detectHeaderRowIndex !== "function") {
    issues.push("Missing export: detectHeaderRowIndex");
  }
  if (typeof prepareHeaderRowsForModify !== "function") {
    issues.push("Missing export: prepareHeaderRowsForModify");
  }
  if (typeof prepareItemRowsWithCcn !== "function") {
    issues.push("Missing export: prepareItemRowsWithCcn");
  }
  if (typeof applyBrokerageAutomation !== "function") {
    issues.push("Missing export: applyBrokerageAutomation");
  }
  if (typeof summarizeDtOutputs !== "function") {
    issues.push("Missing export: summarizeDtOutputs");
  }
  if (issues.length) return { issues };

  const row4HeaderPath = path.join(rootDir, "112-05240631", "CLVS_Report_Header_10133_017927245_260526112458466.xlsx");
  const row4ItemPath = path.join(rootDir, "112-05240631", "CLVS_Report_Detail_10133_017927245_260526112458608.xlsx");
  const row4SourcePath = path.join(rootDir, "112-05240631", "RLBE_50_11205240631_PVG_YYZ_260427045723.xlsx");
  const row5HeaderPath = path.join(rootDir, "83082142460", "RLBE_161_8308214246_EWR_DutiesHeader.xlsx");
  const row5ItemPath = path.join(rootDir, "83082142460", "RLBE_161_8308214246_EWR_DutiesItem.xlsx");
  const brokerageRatesPath = path.join(process.cwd(), "public", "brokerage-rates.json");
  const brokerageRates = readJson(brokerageRatesPath);
  issues.push(...runExpandedDtTaxSchemaRegression(workflowModule, brokerageRates));

  const fixturePaths = [row4HeaderPath, row4ItemPath, row4SourcePath, row5HeaderPath, row5ItemPath];
  const missingFixturePaths = fixturePaths.filter((filePath) => !fs.existsSync(filePath));
  if (missingFixturePaths.length) {
    return {
      issues,
      skippedFixtureReason: `${missingFixturePaths.length} legacy D/T fixture(s) are unavailable; expanded-schema synthetic coverage still ran.`
    };
  }

  const { rows: row4HeaderRows } = readFirstSheetRows(row4HeaderPath);
  const { rows: row4ItemRows } = readFirstSheetRows(row4ItemPath);
  const { rows: row4SourceRows } = readFirstSheetRows(row4SourcePath);
  const { rows: row5HeaderRows } = readFirstSheetRows(row5HeaderPath);
  const { rows: row5ItemRows } = readFirstSheetRows(row5ItemPath);

  const row4HeaderIndex = detectHeaderRowIndex(row4HeaderRows, "header");
  const row4ItemIndex = detectHeaderRowIndex(row4ItemRows, "item");
  const row5HeaderIndex = detectHeaderRowIndex(row5HeaderRows, "header");
  const row5ItemIndex = detectHeaderRowIndex(row5ItemRows, "item");

  if (row4HeaderIndex !== 3) issues.push(`Row-4 header detection failed for CLVS header: got ${row4HeaderIndex}, expected 3`);
  if (row4ItemIndex !== 3) issues.push(`Row-4 header detection failed for CLVS item: got ${row4ItemIndex}, expected 3`);
  if (row5HeaderIndex !== 4) issues.push(`Row-5 header detection failed for RLBE header: got ${row5HeaderIndex}, expected 4`);
  if (row5ItemIndex !== 4) issues.push(`Row-5 header detection failed for RLBE item: got ${row5ItemIndex}, expected 4`);
  if (issues.length) return { issues };

  const metadata = {
    client: "RELIABLE LOGISTICS",
    reportName: "WEEKLY DETAIL REPORT",
    reportDate: "05/27/2026"
  };

  let actualHeaderResult;
  try {
    actualHeaderResult = prepareHeaderRowsForModify({
      targetRows: cloneRows(row4HeaderRows),
      metadata
    });
  } catch (err) {
    return {
      issues: [`prepareHeaderRowsForModify threw: ${err.message}`]
    };
  }

  if (!actualHeaderResult || typeof actualHeaderResult !== "object" || !Array.isArray(actualHeaderResult.rows)) {
    issues.push("prepareHeaderRowsForModify should return an object with rows.");
  }
  if (issues.length) return { issues };

  if (actualHeaderResult.headerRowIndex !== 4) {
    issues.push(`prepareHeaderRowsForModify should expose normalized headerRowIndex=4, got ${JSON.stringify(actualHeaderResult.headerRowIndex)}`);
  }

  const actualHeaderRows = actualHeaderResult.rows;

  if (String((actualHeaderRows[0] || [])[1] || "").trim() !== metadata.client) {
    issues.push(`Header metadata client mismatch: got ${JSON.stringify((actualHeaderRows[0] || [])[1] || "")}`);
  }
  if (String((actualHeaderRows[1] || [])[1] || "").trim() !== metadata.reportName) {
    issues.push(`Header metadata report name mismatch: got ${JSON.stringify((actualHeaderRows[1] || [])[1] || "")}`);
  }
  if ((actualHeaderRows[2] || [])[1] !== metadata.reportDate) {
    issues.push(`Header metadata report date mismatch: got ${JSON.stringify((actualHeaderRows[2] || [])[1] || "")}`);
  }
  if (typeof ((actualHeaderRows[2] || [])[1]) !== "string") {
    issues.push(`Header report date should remain text, got type ${typeof ((actualHeaderRows[2] || [])[1])}`);
  }

  let actualItemResult;
  try {
    actualItemResult = prepareItemRowsWithCcn({
      itemRows: cloneRows(row4ItemRows),
      preparedHeader: actualHeaderResult,
      metadata
    });
  } catch (err) {
    return {
      issues: [`prepareItemRowsWithCcn threw: ${err.message}`]
    };
  }

  if (!actualItemResult || typeof actualItemResult !== "object" || !Array.isArray(actualItemResult.rows)) {
    issues.push("prepareItemRowsWithCcn should return an object with rows.");
  }
  if (issues.length) return { issues };

  if (actualItemResult.headerRowIndex !== 4) {
    issues.push(`prepareItemRowsWithCcn should expose normalized headerRowIndex=4, got ${JSON.stringify(actualItemResult.headerRowIndex)}`);
  }
  if (typeof actualItemResult.unmatchedCount !== "number") {
    issues.push(`prepareItemRowsWithCcn should expose numeric unmatchedCount, got ${typeof actualItemResult.unmatchedCount}`);
  }

  const actualItemRows = actualItemResult.rows;

  if (String((actualItemRows[0] || [])[1] || "").trim() !== metadata.client) {
    issues.push(`Item metadata client mismatch: got ${JSON.stringify((actualItemRows[0] || [])[1] || "")}`);
  }
  if (String((actualItemRows[1] || [])[1] || "").trim() !== metadata.reportName) {
    issues.push(`Item metadata report name mismatch: got ${JSON.stringify((actualItemRows[1] || [])[1] || "")}`);
  }
  if ((actualItemRows[2] || [])[1] !== metadata.reportDate) {
    issues.push(`Item metadata report date mismatch: got ${JSON.stringify((actualItemRows[2] || [])[1] || "")}`);
  }
  if (typeof ((actualItemRows[2] || [])[1]) !== "string") {
    issues.push(`Item report date should remain text, got type ${typeof ((actualItemRows[2] || [])[1])}`);
  }

  const normalizedHeaderRow = actualHeaderRows[4] || [];
  const normalizedItemRow = actualItemRows[4] || [];
  if (normalizeHeaderCell(normalizedHeaderRow[0]) !== "transaction number") {
    issues.push(`Header row normalization failed: row 5 starts with ${JSON.stringify(normalizedHeaderRow[0] || "")}`);
  }
  if (!isEmptyRow(actualHeaderRows[3] || [])) {
    issues.push("Header row normalization failed: row 4 should be blank after normalization.");
  }
  if (normalizeHeaderCell(normalizedItemRow[0]) !== "transaction number") {
    issues.push(`Item row normalization failed: row 5 starts with ${JSON.stringify(normalizedItemRow[0] || "")}`);
  }
  if (!isEmptyRow(actualItemRows[3] || [])) {
    issues.push("Item row normalization failed: row 4 should be blank after normalization.");
  }

  const itemHeaderRow = actualItemRows[4] || [];
  const transactionIdx = itemHeaderRow.findIndex((cell) => normalizeHeaderCell(cell) === "transaction number");
  const ccnIdx = itemHeaderRow.findIndex((cell) => normalizeHeaderCell(cell) === "ccn");
  if (transactionIdx === -1) issues.push("Item workflow output is missing Transaction Number.");
  if (ccnIdx === -1) issues.push("Item workflow output is missing CCN.");

  if (transactionIdx !== -1 && ccnIdx !== -1) {
    const sourceHeaderRow = row4HeaderRows[row4HeaderIndex] || [];
    const sourceTransactionIdx = findHeaderColumnIndex(sourceHeaderRow, "Transaction Number");
    const sourceCcnIdx = findHeaderColumnIndex(sourceHeaderRow, "CCN");
    if (sourceTransactionIdx === -1) issues.push("Source header workbook is missing Transaction Number.");
    if (sourceCcnIdx === -1) issues.push("Source header workbook is missing CCN.");

    const expectedCcnByTransaction = new Map();
    if (sourceTransactionIdx !== -1 && sourceCcnIdx !== -1) {
      for (let r = row4HeaderIndex + 1; r < row4HeaderRows.length; r++) {
        const row = row4HeaderRows[r] || [];
        const transaction = String(row[sourceTransactionIdx] || "").trim();
        const ccn = String(row[sourceCcnIdx] || "").trim();
        if (transaction && ccn) expectedCcnByTransaction.set(transaction, ccn);
      }

      for (let r = 5; r < actualItemRows.length; r++) {
        const row = actualItemRows[r] || [];
        if (isEmptyRow(row)) continue;
        const transaction = String(row[transactionIdx] || "").trim();
        if (!transaction) continue;
        const expectedCcn = expectedCcnByTransaction.get(transaction);
        if (!expectedCcn) continue;

        const actualCcn = String(row[ccnIdx] || "").trim();
        if (actualCcn !== expectedCcn) {
          issues.push(`Item CCN fill mismatch for transaction ${transaction}: got ${JSON.stringify(actualCcn)} expected ${JSON.stringify(expectedCcn)}`);
          break;
        }
      }
    }
  }

  let automatedHeaderResult;
  try {
    automatedHeaderResult = applyBrokerageAutomation({
      sourceRows: cloneRows(row4SourceRows),
      preparedHeader: actualHeaderResult,
      metadata: {
        client: "SF EXPRESS",
        reportName: "AWB# 112-05240631",
        reportDate: "4/28/2026"
      },
      brokerageRates
    });
  } catch (err) {
    return {
      issues: [`applyBrokerageAutomation threw: ${err.message}`]
    };
  }

  if (!automatedHeaderResult || typeof automatedHeaderResult !== "object" || !Array.isArray(automatedHeaderResult.rows)) {
    issues.push("applyBrokerageAutomation should return an object with rows.");
  }
  if (!automatedHeaderResult || !automatedHeaderResult.summary || typeof automatedHeaderResult.summary !== "object") {
    issues.push("applyBrokerageAutomation should return a summary object.");
  }
  if (issues.length) return { issues };

  const automatedHeaderRows = automatedHeaderResult.rows;
  const automatedHeaderRow = automatedHeaderRows[automatedHeaderResult.headerRowIndex] || [];
  const automatedTxnIdx = findHeaderColumnIndex(automatedHeaderRow, "Transaction Number");
  const automatedCcnIdx = findHeaderColumnIndex(automatedHeaderRow, "CCN");
  const automatedShipmentIdx = findHeaderColumnIndex(automatedHeaderRow, "Shipment Date");
  const automatedArrivalIdx = findHeaderColumnIndex(automatedHeaderRow, "Arrival Date");
  const automatedReleaseIdx = findHeaderColumnIndex(automatedHeaderRow, "Release Date");
  const automatedBrokerageIdx = findHeaderColumnIndex(automatedHeaderRow, "Brokerage Total");
  const automatedExchangeIdx = findHeaderColumnIndex(automatedHeaderRow, "Exchange Rate");

  if (automatedHeaderResult.summary.clientMatched !== true) {
    issues.push(`Expected known client match for SF EXPRESS, got ${JSON.stringify(automatedHeaderResult.summary.clientMatched)}`);
  }
  if (normalizeClientName(automatedHeaderResult.summary.clientKey) !== normalizeClientName("SF EXPRESS")) {
    issues.push(`Expected matched client key SF EXPRESS, got ${JSON.stringify(automatedHeaderResult.summary.clientKey)}`);
  }
  if (!automatedHeaderResult.summary.counts || automatedHeaderResult.summary.counts.pga !== 2) {
    issues.push(`Expected PGA count 2, got ${JSON.stringify(automatedHeaderResult.summary && automatedHeaderResult.summary.counts && automatedHeaderResult.summary.counts.pga)}`);
  }
  if (!automatedHeaderResult.summary.counts || automatedHeaderResult.summary.counts.lvs !== 72) {
    issues.push(`Expected LVS count 72, got ${JSON.stringify(automatedHeaderResult.summary && automatedHeaderResult.summary.counts && automatedHeaderResult.summary.counts.lvs)}`);
  }
  if (!automatedHeaderResult.summary.counts || automatedHeaderResult.summary.counts.clvs <= 0) {
    issues.push(`Expected positive CLVS count, got ${JSON.stringify(automatedHeaderResult.summary && automatedHeaderResult.summary.counts && automatedHeaderResult.summary.counts.clvs)}`);
  }
  if (automatedHeaderResult.summary.blankBrokerageCount !== 0) {
    issues.push(`Expected no blank brokerage rows for SF EXPRESS, got ${JSON.stringify(automatedHeaderResult.summary.blankBrokerageCount)}`);
  }

  const firstAutomatedDataRows = automatedHeaderRows.slice(automatedHeaderResult.headerRowIndex + 1, automatedHeaderResult.headerRowIndex + 6);
  if (!String((firstAutomatedDataRows[0] || [])[automatedCcnIdx] || "").trim().startsWith("8308")) {
    issues.push("Expected first transformed data row to be a PGA 8308 row after brokerage sort.");
  }

  for (let r = automatedHeaderResult.headerRowIndex + 1; r < Math.min(automatedHeaderRows.length, automatedHeaderResult.headerRowIndex + 15); r++) {
    const row = automatedHeaderRows[r] || [];
    if (isEmptyRow(row)) continue;
    if (String(row[automatedShipmentIdx] || "").trim() !== "4/28/2026") {
      issues.push(`Shipment date mismatch at row ${r + 1}: got ${JSON.stringify(row[automatedShipmentIdx] || "")}`);
      break;
    }
    if (String(row[automatedArrivalIdx] || "").trim() !== "4/28/2026") {
      issues.push(`Arrival date mismatch at row ${r + 1}: got ${JSON.stringify(row[automatedArrivalIdx] || "")}`);
      break;
    }
    if (String(row[automatedReleaseIdx] || "").trim() !== "4/28/2026") {
      issues.push(`Release date mismatch at row ${r + 1}: got ${JSON.stringify(row[automatedReleaseIdx] || "")}`);
      break;
    }
    if (parseNumberZero(row[automatedExchangeIdx]) !== 0) {
      issues.push(`Exchange rate should be zero at row ${r + 1}, got ${JSON.stringify(row[automatedExchangeIdx] || "")}`);
      break;
    }
  }

  const actualDistinctBrokerage = [];
  for (const row of automatedHeaderRows.slice(automatedHeaderResult.headerRowIndex + 1)) {
    if (isEmptyRow(row)) continue;
    const brokerageValue = parseNumberFromCell(row[automatedBrokerageIdx]);
    if (brokerageValue !== null && !actualDistinctBrokerage.some((value) => Math.abs(value - brokerageValue) <= 0.0000001)) {
      actualDistinctBrokerage.push(brokerageValue);
    }
  }
  const expectedDistinctBrokerage = [0.71, 0.28, 0.0175];
  if (JSON.stringify(actualDistinctBrokerage.slice(0, 3)) !== JSON.stringify(expectedDistinctBrokerage)) {
    issues.push(`Brokerage sort/order mismatch: actual=${JSON.stringify(actualDistinctBrokerage.slice(0, 3))} expected=${JSON.stringify(expectedDistinctBrokerage)}`);
  }

  let automatedItemResult;
  try {
    automatedItemResult = prepareItemRowsWithCcn({
      itemRows: cloneRows(row4ItemRows),
      preparedHeader: {
        rows: automatedHeaderRows,
        headerRowIndex: automatedHeaderResult.headerRowIndex
      },
      metadata: {
        client: "SF EXPRESS",
        reportName: "AWB# 112-05240631",
        reportDate: "4/28/2026"
      }
    });
  } catch (err) {
    return {
      issues: [`prepareItemRowsWithCcn with automated header threw: ${err.message}`]
    };
  }

  let combinedSummary;
  try {
    combinedSummary = summarizeDtOutputs({
      headerRows: automatedHeaderRows,
      itemRows: automatedItemResult.rows
    });
  } catch (err) {
    return {
      issues: [`summarizeDtOutputs threw: ${err.message}`]
    };
  }

  if (!combinedSummary || !combinedSummary.header || !combinedSummary.compare) {
    issues.push("summarizeDtOutputs should return header/item/compare sections.");
  } else {
    if (typeof combinedSummary.compare.dutyMatch !== "boolean") {
      issues.push(`Expected boolean dutyMatch, got ${typeof combinedSummary.compare.dutyMatch}`);
    }
    if (typeof combinedSummary.compare.gstMatch !== "boolean") {
      issues.push(`Expected boolean gstMatch, got ${typeof combinedSummary.compare.gstMatch}`);
    }
    if (!combinedSummary.validation || !combinedSummary.validation.header || !combinedSummary.validation.item) {
      issues.push("summarizeDtOutputs should return Header and Item validation results.");
    }
  }

  const validationHeaderRows = [
    ["CLIENT:", "TEST"],
    ["RPT NAME:", "TEST REPORT"],
    ["RPT DATE :", "05/27/2026"],
    ["Transaction Number", "CCN", "Order Number", "Value for Duty", "Duty", "Gov. Sales Tax"],
    ["TX-1", "CCN-1", "ORD-1", "0.00", "", "$0.00"],
    ["TX-2", "", "ORD-2", "0.00", "1.00", "2.00"],
    ["", "", "", "", "", ""]
  ];
  const headerValidation = workflowModule.validateReportRows(validationHeaderRows, "header");
  if (headerValidation.error) {
    issues.push(`Header validation fixture unexpectedly failed: ${headerValidation.error}`);
  } else {
    if (headerValidation.issueCount !== 4 || headerValidation.blankCount !== 1 || headerValidation.zeroCount !== 3) {
      issues.push(`Header validation counts mismatch: ${JSON.stringify({ issueCount: headerValidation.issueCount, blankCount: headerValidation.blankCount, zeroCount: headerValidation.zeroCount })}`);
    }
    const orderFallbackIssue = headerValidation.issues.find((issue) => issue.transactionNumber === "TX-2");
    if (!orderFallbackIssue || orderFallbackIssue.ccn !== "ORD-2") {
      issues.push(`Header validation should use Order Number as the CCN fallback: ${JSON.stringify(orderFallbackIssue)}`);
    }
  }

  const clvsExpectedZeroRows = [
    ["CLIENT:", "TEST"],
    ["RPT NAME:", "TEST REPORT"],
    ["RPT DATE :", "05/27/2026"],
    ["Transaction Number", "CCN", "Value for Duty", "Duty", "Gov. Sales Tax"],
    ["CLVS", "CLVS-1", "0.00", "0.00", "0.00"],
    ["LV-1", "LV-CCN-1", "0.00", "0.00", "0.00"]
  ];
  const clvsExpectedZeroValidation = workflowModule.validateReportRows(clvsExpectedZeroRows, "header");
  if (
    clvsExpectedZeroValidation.error ||
    clvsExpectedZeroValidation.issueCount !== 4 ||
    clvsExpectedZeroValidation.ignoredExpectedZeroCount !== 2 ||
    clvsExpectedZeroValidation.fieldCounts["Value for Duty"]?.zero !== 2 ||
    clvsExpectedZeroValidation.fieldCounts.Duty?.zero !== 1 ||
    clvsExpectedZeroValidation.fieldCounts["Gov. Sales Tax"]?.zero !== 1
  ) {
    issues.push(`CLVS Duty/GST zero exemptions should preserve Value for Duty checks: ${JSON.stringify(clvsExpectedZeroValidation)}`);
  }

  const validationItemRows = [
    ["CLIENT:", "TEST"],
    ["RPT NAME:", "TEST REPORT"],
    ["RPT DATE :", "05/27/2026"],
    ["Transaction Number", "Goods Description", "Line #", "Quantity", "Value for Duty", "Duty", "Value for Tax", "Gov. Sales Tax", "CCN"],
    ["TX-1", "ITEM A", "1", "0", "", "$0.00", "12.00", "1.00", "CCN-1"],
    ["TX-2", "ITEM B", "2", "", "5.00", "2.00", "0.00", "0.00", "CCN-2"],
    ["", "", "", "", "", "", "", "", ""]
  ];
  const itemValidation = workflowModule.validateReportRows(validationItemRows, "item");
  if (itemValidation.error) {
    issues.push(`Item validation fixture unexpectedly failed: ${itemValidation.error}`);
  } else {
    if (itemValidation.issueCount !== 6 || itemValidation.blankCount !== 2 || itemValidation.zeroCount !== 4) {
      issues.push(`Item validation counts mismatch: ${JSON.stringify({ issueCount: itemValidation.issueCount, blankCount: itemValidation.blankCount, zeroCount: itemValidation.zeroCount })}`);
    }
    const itemIssue = itemValidation.issues.find((issue) => issue.transactionNumber === "TX-1" && issue.field === "Quantity");
    if (!itemIssue || itemIssue.ccn !== "CCN-1" || itemIssue.lineNumber !== "1") {
      issues.push(`Item validation should identify CCN, Transaction Number, and Line Number: ${JSON.stringify(itemIssue)}`);
    }
  }

  const missingValueForTax = validationItemRows.map((row) => row.slice(0, 6).concat(row.slice(7)));
  const missingColumnValidation = workflowModule.validateReportRows(missingValueForTax, "item");
  if (!missingColumnValidation.error || !/Value for Tax/i.test(missingColumnValidation.error)) {
    issues.push(`Missing Item validation columns should be reported: ${JSON.stringify(missingColumnValidation)}`);
  }

  let headerOnlySummary;
  try {
    headerOnlySummary = summarizeDtOutputs({ headerRows: automatedHeaderRows });
  } catch (err) {
    issues.push(`Header-only summarizeDtOutputs threw: ${err.message}`);
  }
  if (headerOnlySummary && headerOnlySummary.validation && headerOnlySummary.validation.item !== null) {
    issues.push("Header-only validation should not create an Item result.");
  }

  // Exercise every monitored field, including fractional quantities below the old tolerance.
  for (const [mode, base, fieldLabels] of [
    ["header", validationHeaderRows, ["Value for Duty", "Duty", "Gov. Sales Tax"]],
    ["item", validationItemRows, ["Quantity", "Value for Duty", "Duty", "Value for Tax", "Gov. Sales Tax"]]
  ]) {
    const headerIndex = detectHeaderRowIndex(base, mode);
    const header = base[headerIndex];
    const fieldIndexes = fieldLabels.map((label) => header.indexOf(label));
    for (const value of [0, -0, "0", "$0.00", "(0.00)", "0.00-", "$ -", "-", "0e-8", 1e-8, -1e-8, "0.00000001", "-0.00000001", "$0.00000001", "0oops"]) {
      const rows = cloneRows(base.slice(0, headerIndex + 1));
      const row = new Array(header.length).fill("IDENTIFIER");
      fieldIndexes.forEach((index) => { row[index] = value; });
      rows.push(row);
      const result = workflowModule.validateReportRows(rows, mode);
      const expected = [1e-8, -1e-8, "0.00000001", "-0.00000001", "$0.00000001", "0oops"].includes(value) ? 0 : fieldLabels.length;
      if (result.issueCount !== expected) issues.push(`${mode} exact-zero case ${JSON.stringify(value)}: ${result.issueCount} warnings, expected ${expected}`);
    }
  }

  // Real supplied uploads and their final transformed outputs retain every warning.
  const rawHeaderValidation = workflowModule.validateReportRows(row4HeaderRows, "header");
  const rawItemValidation = workflowModule.validateReportRows(row4ItemRows, "item");
  if (
    rawHeaderValidation.zeroCount !== 36 ||
    rawHeaderValidation.blankCount !== 0 ||
    rawHeaderValidation.rowsChecked !== 74 ||
    rawHeaderValidation.fieldCounts["Duty"]?.zero !== 34 ||
    rawHeaderValidation.fieldCounts["Value for Duty"]?.zero !== 2
  ) {
    issues.push("Supplied raw Header should contain 74 rows with 34 Duty zeros and 2 Value for Duty zeros.");
  }
  if (rawItemValidation.zeroCount !== 38 || rawItemValidation.blankCount !== 2 || rawItemValidation.rowsChecked !== 116) {
    issues.push("Supplied raw Item should contain 116 rows with 38 Duty zeros and 2 Duty blanks.");
  }
  const finalValidation = summarizeDtOutputs({
    headerRows: automatedHeaderRows,
    itemRows: actualItemResult.rows,
    generatedHeaderRowNumbers: automatedHeaderResult.generatedRowNumbers
  }).validation;
  if (
    finalValidation.header.issueCount !== 36 ||
    finalValidation.header.uploadedIssueCount !== 36 ||
    finalValidation.header.generatedIssueCount !== 0 ||
    finalValidation.header.ignoredExpectedZeroCount !== 282
  ) {
    issues.push(`Processed Header warning provenance mismatch: ${JSON.stringify({ total: finalValidation.header.issueCount, uploaded: finalValidation.header.uploadedIssueCount, generated: finalValidation.header.generatedIssueCount, expectedExcluded: finalValidation.header.ignoredExpectedZeroCount })}`);
  }
  if (finalValidation.item.issueCount !== 40 || finalValidation.item.generatedIssueCount !== 0) issues.push("Processed Item should retain its 40 uploaded-row warnings.");
  for (const finding of finalValidation.header.issues) {
    const row = automatedHeaderRows[finding.rowNumber - 1];
    if (finding.ccn !== workflowModule.getRecordCcn(row, automatedHeaderRows[automatedHeaderResult.headerRowIndex])) issues.push("Sorted finding CCN does not match its output Excel row.");
    if ((finding.origin === "generated") !== automatedHeaderResult.generatedRowNumbers.includes(finding.rowNumber)) issues.push("Generated provenance did not survive sorting.");
  }

  // Alias precedence is shared by Header classification, deduplication, and Item propagation.
  for (const alias of ["Cargo Control Number", ""]) {
    const aliasRows = cloneRows(actualHeaderResult.rows);
    const headerIndex = actualHeaderResult.headerRowIndex;
    const explicitCcnIndex = aliasRows[headerIndex].indexOf("CCN");
    aliasRows[headerIndex][explicitCcnIndex] = alias; // Empty label exercises the existing Order Number column.
    const aliasOutput = applyBrokerageAutomation({ sourceRows: row4SourceRows, preparedHeader: { rows: aliasRows }, metadata: { client: "SF EXPRESS", reportDate: "4/28/2026" }, brokerageRates });
    if (JSON.stringify(aliasOutput.summary) !== JSON.stringify(automatedHeaderResult.summary) || aliasOutput.insertedCount !== automatedHeaderResult.insertedCount) issues.push(`${alias || "Order Number"} alias changed classification, totals, or inserted count.`);
    const aliasItem = workflowModule.prepareItemRowsWithCcn({ itemRows: row4ItemRows, preparedHeader: aliasOutput });
    const aliasItemHeader = aliasItem.rows[aliasItem.headerRowIndex];
    const itemCcnIndex = aliasItemHeader.indexOf("CCN");
    const aliasValidation = workflowModule.validateReportRows(aliasOutput.rows, "header");
    if (aliasValidation.issueCount !== 36 || aliasItem.unmatchedCount !== actualItemResult.unmatchedCount) issues.push(`${alias || "Order Number"} alias changed validation or Item mapping.`);
    for (let index = aliasItem.headerRowIndex + 1; index < aliasItem.rows.length; index++) {
      if (aliasItem.rows[index][itemCcnIndex] !== actualItemResult.rows[index][itemCcnIndex]) { issues.push(`${alias || "Order Number"} alias changed an Item CCN.`); break; }
    }
  }
  const precedenceHeader = ["Transaction Number", "CCN", "Cargo Control Number", "Order Number"];
  for (const [row, expected] of [[["TX", "CCN-A", "CARGO-B", "ORDER-C"], "CCN-A"], [["TX", " ", "CARGO-B", "ORDER-C"], "CARGO-B"], [["TX", "", "", "ORDER-C"], "ORDER-C"]]) {
    if (workflowModule.getRecordCcn(row, precedenceHeader) !== expected) issues.push("CCN alias precedence or blank-cell fallback failed.");
    const mapped = workflowModule.prepareItemRowsWithCcn({ headerRows: [["CLIENT:"], ["RPT NAME:"], ["RPT DATE :"], [], precedenceHeader, row], itemRows: [["CLIENT:"], ["RPT NAME:"], ["RPT DATE :"], [], ["Transaction Number", "Goods Description"], ["TX", "Item"]] });
    if (mapped.rows[5][2] !== expected) issues.push("Item mapping did not use per-row CCN fallback.");
  }

  for (const field of ["Quantity", "Value for Duty", "Value for Tax", "Duty", "Gov. Sales Tax"]) {
    const invalidItem = cloneRows(validationItemRows);
    const header = invalidItem[detectHeaderRowIndex(invalidItem, "item")];
    header[header.indexOf(field)] = "Missing field";
    try {
      summarizeDtOutputs({ headerRows: automatedHeaderRows, itemRows: invalidItem });
      issues.push(`Missing Item ${field} must throw before downloads.`);
    } catch (err) {
      if (!String(err.message).includes(field)) issues.push(`Missing ${field} error did not identify the field.`);
    }
  }

  let unknownClientResult;
  try {
    unknownClientResult = applyBrokerageAutomation({
      sourceRows: cloneRows(row4SourceRows),
      preparedHeader: actualHeaderResult,
      metadata: {
        client: "UNKNOWN CLIENT",
        reportName: "AWB# 112-05240631",
        reportDate: "4/28/2026"
      },
      brokerageRates
    });
  } catch (err) {
    return {
      issues: [`applyBrokerageAutomation unknown-client run threw: ${err.message}`]
    };
  }

  if (unknownClientResult.summary.clientMatched !== false) {
    issues.push("Unknown client should report clientMatched=false.");
  }
  if (unknownClientResult.summary.blankBrokerageCount <= 0) {
    issues.push("Unknown client should leave brokerage blank for at least one row.");
  }

  try {
    detectHeaderRowIndex(row4HeaderRows, "invalid");
    issues.push("detectHeaderRowIndex should throw on invalid mode.");
  } catch (err) {
    if (!/mode/i.test(String(err && err.message))) {
      issues.push(`detectHeaderRowIndex invalid-mode error should mention mode, got ${JSON.stringify(err.message)}`);
    }
  }

  return { issues };
}

function runCandataRegression(rootDir) {
  const headerIn = path.join(rootDir, "Header CANDATA.xlsx");
  const itemIn = path.join(rootDir, "Item CANDATA.xlsx");
  const headerSample = path.join(rootDir, "SAMPLE  NEW FORMAT HEADER GETS.xlsx");
  const itemSample = path.join(rootDir, "SAMPLE  NEW FORMAT ITEM GETS.xlsx");
  const headerIssues = [];
  const itemIssues = [];
  const addIssue = (issues, condition, message) => {
    if (!condition) issues.push(message);
  };

  for (const file of [headerIn, itemIn, headerSample, itemSample]) {
    if (!fs.existsSync(file)) {
      const issues = file === itemIn || file === itemSample ? itemIssues : headerIssues;
      issues.push(`Missing test file: ${file}`);
    }
  }
  if (headerIssues.length || itemIssues.length) return { headerIssues, itemIssues };

  try {
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
      { key: "totalProvincialSalesTax", label: "Provincial Sales Tax (CAD)", matchers: [/^(?:total )?provincial sales tax/i] },
      CandataWorkflow.hstColumnSpec(),
      { key: "sima", label: "SIMA (CAD)", matchers: [/^sima(?: \(cad\))?$/i] },
      { key: "surtax", label: "Surtax (CAD)", matchers: [/^surtax(?: \(cad\))?$/i] },
      { key: "province", label: "Province (CAD)", matchers: [/^province(?: \(cad\))?$/i] },
      { key: "paymentTerms", label: "Payment Terms", matchers: [/^payment terms$/i, /^inco terms$/i] },
      { key: "billOfLading", label: "Bill of Lading", matchers: [/bill of lading/i, /\bawb\b/i] }
    ];
    const itemSpecs = [
      { key: "transaction", label: "Transaction Number", matchers: [/^transaction number$/i] },
      { key: "productDescription", label: "Product Description", matchers: [/product description/i, /goods description/i] },
      { key: "cciLine", label: "CCI Line#", matchers: [/cci line#?/i, /\bline #\b/i] },
      { key: "productNumber", label: "Product Number", matchers: [/^product number$/i, /^part number$/i] },
      { key: "countryOfOrigin", label: "Country of Origin", matchers: [/country of origin/i] },
      { key: "tariffTreatment", label: "Tariff Treatment", matchers: [/tariff treatment/i] },
      { key: "quantity", label: "Quantity", matchers: [/^quantity$/i] },
      { key: "port", label: "Port Number", matchers: [/^port number$/i, /^port #$/i] },
      { key: "vendorName", label: "Vendor Name", matchers: [/^vendor(?: name)?$/i] },
      { key: "valueForDuty", label: "Value For Duty (CAD)", matchers: [/value for duty/i] },
      { key: "classification", label: "Classification", matchers: [/^classification$/i, /^hs #$/i] },
      { key: "dutyRate", label: "Duty Rate", matchers: [/^duty rate$/i] },
      { key: "customsDuty", label: "Customs Duty (CAD)", matchers: [/customs duty/i, /^duty$/i] },
      { key: "valueForTax", label: "Value for Tax (CAD)", matchers: [/value for tax/i] },
      { key: "gst", label: "GST (CAD)", matchers: [/^gst/i, /gov\.?\s*sales/i] },
      { key: "pst", label: "Provincial Sales Tax (CAD)", matchers: [/provincial sales tax/i, /\bpst\b/i] },
      CandataWorkflow.hstColumnSpec(),
      { key: "sima", label: "SIMA (CAD)", matchers: [/^sima(?: \(cad\))?$/i] },
      { key: "surtax", label: "Surtax (CAD)", matchers: [/^surtax(?: \(cad\))?$/i] },
      { key: "province", label: "Province (CAD)", matchers: [/^province(?: \(cad\))?$/i] },
      { key: "ccn", label: "Cargo Control Number", matchers: [/cargo control number/i, /^ccn$/i] },
      { key: "billOfLading", label: "Bill of Lading", matchers: [/bill of lading/i, /\bawb\b/i] }
    ];

    const { rows: rawHeaderRows } = readFirstSheetRows(headerIn);
    const { rows: rawItemRows } = readFirstSheetRows(itemIn);
    const { rows: headerSampleRows } = readFirstSheetRows(headerSample);
    const { rows: itemSampleRows } = readFirstSheetRows(itemSample);
    const headerDetected = validateRequiredColumns(rawHeaderRows, headerSpecs, "DutiesHeader");
    const itemDetected = validateRequiredColumns(rawItemRows, itemSpecs, "Candata Duties Item");
    const headerFound = CandataWorkflow.normalizeTaxColumns(rawHeaderRows, headerDetected, "totalGst", "totalProvincialSalesTax");
    const itemFound = CandataWorkflow.normalizeTaxColumns(rawItemRows, itemDetected, "gst", "pst");
    const headerRows = headerFound.rows;
    const itemRows = itemFound.rows;
    const headerDataStart = headerFound.rowIndex + 1;
    const itemDataStart = itemFound.rowIndex + 1;
    const termsLookup = buildPaymentTermsLookup(headerRows, headerFound.indexMap, headerDataStart);
    const headerRecords = buildHeaderRecords(headerRows, headerFound.indexMap, headerDataStart);
    const itemRecords = buildItemRecords(itemRows, itemFound.indexMap, itemDataStart, termsLookup);

    let firstBOL = "";
    let firstReleaseDate = "";
    for (let r = headerDataStart; r < headerRows.length; r++) {
      const row = headerRows[r] || [];
      if (!firstBOL) firstBOL = sanitizeExcelText(row[headerFound.indexMap.billOfLading]);
      if (!firstReleaseDate) firstReleaseDate = formatDateMMDDYYYY(row[headerFound.indexMap.releaseDate]);
      if (firstBOL && firstReleaseDate) break;
    }
    const reportName = `AWB # ${firstBOL || ""}`.trim();
    const actualHeaderAoA = buildHeaderOutputAoA(headerRecords, reportName, firstReleaseDate, "");
    const actualItemAoA = buildItemOutputAoA(itemRecords, reportName, firstReleaseDate, "");

    const expectedHeaderHeaders = [
      "Transaction Number", "CCN", "Port #", "Shipment Date", "Arrival Date", "Release Date",
      "No. of Cartons", "Order Number", "Other Reference", "Value for Duty", "Duty", "GST", "HST", "PST",
      "SIMA", "Surtax", "Brokerage Total", "Addl. Charges Total", "Excise Tax Total", "Exchange Rate", "Inco Terms"
    ];
    const expectedItemHeaders = [
      "Transaction Number", "Goods Description", "Line #", "Country of Origin", "Tariff Treatment",
      "Part Number", "Quantity", "Port #", "Vendor Name", "Value for Duty", "HS #", "Duty Rate", "Duty",
      "Value for Tax", "GST", "HST", "PST", "SIMA", "Surtax", "Inco Terms"
    ];
    const sampleHeaderHeaders = headerSampleRows[3] || [];
    const sampleItemHeaders = itemSampleRows[3] || [];
    addIssue(headerIssues, JSON.stringify(sampleHeaderHeaders) === JSON.stringify(expectedHeaderHeaders), "Header sample does not match the requested fixed schema at row 4.");
    addIssue(itemIssues, JSON.stringify(sampleItemHeaders) === JSON.stringify(expectedItemHeaders), "Item sample does not match the requested fixed schema at row 4.");
    addIssue(headerIssues, JSON.stringify(actualHeaderAoA[3]) === JSON.stringify(expectedHeaderHeaders), "Header output columns or row-4 placement differ from the requested schema.");
    addIssue(itemIssues, JSON.stringify(actualItemAoA[3]) === JSON.stringify(expectedItemHeaders), "Item output columns or row-4 placement differ from the requested schema.");
    addIssue(headerIssues, actualHeaderAoA.length === 4 + headerRecords.length, "Header output contains an unexpected blank row or row offset.");
    addIssue(itemIssues, actualItemAoA.length === 4 + itemRecords.length, "Item output contains an unexpected blank row or row offset.");
    addIssue(headerIssues, actualHeaderAoA[3]?.includes("Assessment Total") === false, "Header output still contains Assessment Total.");
    addIssue(itemIssues, actualItemAoA[3]?.includes("CCN") === false, "Item output must not contain CCN.");

    const firstHeaderSourceRow = headerRows.slice(headerDataStart).find((row) => !isEmptyRow(row));
    const firstItemSourceRow = itemRows.slice(itemDataStart).find((row) => !isEmptyRow(row));
    const firstHeaderOutput = actualHeaderAoA[4] || [];
    const firstItemOutput = actualItemAoA[4] || [];
    addIssue(headerIssues, headerRecords.length === 12, `Expected one Header row for each of 12 shipments; got ${headerRecords.length}.`);
    const expectedHeaderOrder = [];
    const seenHeaderOrder = new Set();
    for (let r = headerDataStart; r < headerRows.length; r++) {
      const row = headerRows[r] || [];
      const ccn = String(row[headerFound.indexMap.ccn] || "").trim();
      if (!ccn) continue;
      const key = ccn.toLowerCase();
      if (!seenHeaderOrder.has(key)) {
        seenHeaderOrder.add(key);
        expectedHeaderOrder.push(`${trimTransaction(row[headerFound.indexMap.transaction])}|${ccn}`);
      }
    }
    addIssue(headerIssues, JSON.stringify(headerRecords.map((record) => `${record.transactionNumber}|${record.ccn}`)) === JSON.stringify(expectedHeaderOrder), "Header shipment order changed from first-seen source order.");
    addIssue(headerIssues, firstHeaderOutput[0] === trimTransaction(firstHeaderSourceRow[headerFound.indexMap.transaction]), "Header source order or transaction mapping changed.");
    addIssue(headerIssues, firstHeaderOutput[1] === firstHeaderSourceRow[headerFound.indexMap.ccn], "Header CCN mapping changed.");
    addIssue(headerIssues, firstHeaderOutput[3] === formatDateMMDDYYYY(firstHeaderSourceRow[headerFound.indexMap.directShipDate]), "Header Direct Ship Date mapping changed.");
    addIssue(headerIssues, firstHeaderOutput[4] === formatDateMMDDYYYY(firstHeaderSourceRow[headerFound.indexMap.etaDate]), "Header ETA Date mapping changed.");
    addIssue(headerIssues, firstHeaderOutput[5] === formatDateMMDDYYYY(firstHeaderSourceRow[headerFound.indexMap.releaseDate]), "Header Release Date mapping changed.");
    addIssue(headerIssues, firstHeaderOutput[7] === firstHeaderSourceRow[headerFound.indexMap.orderNumber], "Header Order Number mapping changed.");
    addIssue(headerIssues, firstHeaderOutput[16] === 0 && firstHeaderOutput[17] === 0 && firstHeaderOutput[18] === 0 && firstHeaderOutput[19] === 1, "Header non-source totals or Exchange Rate defaults changed.");

    const sourceItemRows = itemRows.slice(itemDataStart).filter((row) => !isEmptyRow(row));
    addIssue(itemIssues, itemRecords.length === sourceItemRows.length, "Item rows were dropped or duplicated.");
    addIssue(itemIssues, firstItemOutput[0] === trimTransaction(firstItemSourceRow[itemFound.indexMap.transaction]), "Item source order or transaction mapping changed.");
    addIssue(itemIssues, firstItemOutput[1] === firstItemSourceRow[itemFound.indexMap.productDescription], "Item Product Description mapping changed.");
    addIssue(itemIssues, firstItemOutput[2] === firstItemSourceRow[itemFound.indexMap.cciLine], "Item CCI Line mapping changed.");
    addIssue(itemIssues, firstItemOutput[5] === firstItemSourceRow[itemFound.indexMap.productNumber], "Item Product Number to Part Number mapping changed.");
    addIssue(itemIssues, firstItemOutput[8] === firstItemSourceRow[itemFound.indexMap.vendorName], "Item Vendor to Vendor Name mapping changed.");
    addIssue(itemIssues, firstItemOutput[19] === findPaymentTerms(termsLookup, firstItemSourceRow[itemFound.indexMap.ccn], firstItemSourceRow[itemFound.indexMap.transaction]), "Item Inco Terms did not join from Header Payment Terms.");
    addIssue(itemIssues, actualItemAoA.slice(4).every((row, i) => row[0] === trimTransaction(sourceItemRows[i][itemFound.indexMap.transaction]) && row[2] === sourceItemRows[i][itemFound.indexMap.cciLine]), "Item source grouping or line order changed.");
    addIssue(itemIssues, itemRecords.every((record, i) => record.incoTerms === findPaymentTerms(termsLookup, sourceItemRows[i][itemFound.indexMap.ccn], sourceItemRows[i][itemFound.indexMap.transaction])), "One or more Item rows did not inherit matching Header Payment Terms.");

    const sumFor = (records, key) => round2(records.reduce((total, record) => total + parseNumberZero(record[key]), 0));
    const groupedDutyAndGstMatch = headerRecords.every((headerRecord) => {
      const shipmentItems = itemRecords.filter((itemRecord) =>
        itemRecord.ccn.toLowerCase() === headerRecord.ccn.toLowerCase() &&
        itemRecord.transactionNumber === headerRecord.transactionNumber
      );
      return shipmentItems.length > 0 && ["valueForDuty", "duty", "gst"]
        .every((key) => sumFor(shipmentItems, key) === round2(parseNumberZero(headerRecord[key])));
    });
    addIssue(headerIssues, groupedDutyAndGstMatch, "Header Value for Duty, Duty, or GST does not reconcile to grouped Item rows for every shipment.");

    const expectedHeaderTaxes = new Map();
    for (let r = headerDataStart; r < headerRows.length; r++) {
      const row = headerRows[r] || [];
      const ccn = String(row[headerFound.indexMap.ccn] || "").trim();
      if (!ccn) continue;
      const key = ccn.toLowerCase();
      const routed = routeProvinceTax(row[headerFound.indexMap.province], row[headerFound.indexMap.totalProvincialSalesTax], row[headerFound.indexMap.hst]);
      const expected = expectedHeaderTaxes.get(key) || { hst: routed.hst, pst: routed.pst, sima: 0, surtax: 0 };
      expected.sima = round2(expected.sima + parseNumberZero(row[headerFound.indexMap.sima]));
      expected.surtax = round2(expected.surtax + parseNumberZero(row[headerFound.indexMap.surtax]));
      expectedHeaderTaxes.set(key, expected);
    }
    const expectedTaxesInFirstSeenOrder = Array.from(expectedHeaderTaxes.values());
    addIssue(headerIssues, headerRecords.every((record, index) => {
      const expected = expectedTaxesInFirstSeenOrder[index];
      return expected && ["hst", "pst", "sima", "surtax"].every((field) => record[field] === expected[field]);
    }), "Header did not retain first PST/HST and sum only SIMA/Surtax by CCN.");

    const headerAliasRow = headerRows[headerFound.rowIndex].map((cell) => String(cell).replace(/^Provincial Sales Tax \(CAD\)$/i, "Total Provincial Sales Tax (CAD)"));
    const headerAliasFound = validateRequiredColumns([headerAliasRow], headerSpecs, "Header aliases");
    addIssue(headerIssues, headerAliasFound.indexMap.totalProvincialSalesTax >= 0, "Header Provincial Sales Tax alias validation failed.");
    const itemAliasRow = itemRows[itemFound.rowIndex].map((cell) => {
      if (/^Vendor$/i.test(String(cell))) return "Vendor Name";
      if (/^Product Number$/i.test(String(cell))) return "Part Number";
      return cell;
    });
    const itemAliasFound = validateRequiredColumns([itemAliasRow], itemSpecs, "Item aliases");
    addIssue(itemIssues, itemAliasFound.indexMap.vendorName >= 0 && itemAliasFound.indexMap.productNumber >= 0, "Item Vendor or Part Number alias validation failed.");

    for (const province of ["NB", "NL", "NS", "ON", "PE"]) {
      const routed = routeProvinceTax(` ${province.toLowerCase()} `, "$7.89");
      addIssue(itemIssues, routed.moved && routed.hst === 7.89 && routed.pst === 0, `Listed province ${province} did not route PST to HST.`);
    }
    const controlTax = routeProvinceTax(" qc ", "$7.89");
    addIssue(itemIssues, !controlTax.moved && controlTax.hst === 0 && controlTax.pst === 7.89, "Non-listed QC control did not preserve source PST.");

    const syntheticItem = firstItemSourceRow.slice();
    syntheticItem[itemFound.indexMap.gst] = "$1.11";
    syntheticItem[itemFound.indexMap.pst] = "$2.22";
    syntheticItem[itemFound.indexMap.sima] = "$3.33";
    syntheticItem[itemFound.indexMap.surtax] = "$4.44";
    syntheticItem[itemFound.indexMap.province] = " ON ";
    const syntheticItemRecord = buildItemRecords([itemRows[itemFound.rowIndex], syntheticItem], itemFound.indexMap, 1, termsLookup)[0];
    addIssue(itemIssues, syntheticItemRecord.gst === 1.11 && syntheticItemRecord.hst === 2.22 && syntheticItemRecord.pst === 0 && syntheticItemRecord.sima === 3.33 && syntheticItemRecord.surtax === 4.44, "Item GST, HST, PST, SIMA, or Surtax values were combined or mis-mapped.");
    const nextSyntheticItem = syntheticItem.slice();
    nextSyntheticItem[itemFound.indexMap.cciLine] = "LINE-SECOND";
    nextSyntheticItem[itemFound.indexMap.gst] = "$5.55";
    nextSyntheticItem[itemFound.indexMap.pst] = "$6.66";
    nextSyntheticItem[itemFound.indexMap.sima] = "$7.77";
    nextSyntheticItem[itemFound.indexMap.surtax] = "$8.88";
    const syntheticItemRows = buildItemRecords([itemRows[itemFound.rowIndex], syntheticItem, nextSyntheticItem], itemFound.indexMap, 1, termsLookup);
    const syntheticItemOutput = buildItemOutputAoA(syntheticItemRows, "", "", "").slice(4);
    addIssue(itemIssues, syntheticItemRows.length === 2 && syntheticItemOutput.length === 2 &&
      syntheticItemOutput[0][2] === syntheticItem[itemFound.indexMap.cciLine] && syntheticItemOutput[1][2] === "LINE-SECOND" &&
      syntheticItemOutput[0][14] === 1.11 && syntheticItemOutput[0][15] === 2.22 && syntheticItemOutput[0][16] === 0 && syntheticItemOutput[0][17] === 3.33 && syntheticItemOutput[0][18] === 4.44 &&
      syntheticItemOutput[1][14] === 5.55 && syntheticItemOutput[1][15] === 6.66 && syntheticItemOutput[1][16] === 0 && syntheticItemOutput[1][17] === 7.77 && syntheticItemOutput[1][18] === 8.88,
    "Duplicate-shipment Item rows lost their source order or individual amounts.");

    const syntheticHeader = firstHeaderSourceRow.slice();
    syntheticHeader[headerFound.indexMap.directShipDate] = "2026-10-01";
    syntheticHeader[headerFound.indexMap.etaDate] = "2026-10-02";
    syntheticHeader[headerFound.indexMap.releaseDate] = "2026-10-03";
    syntheticHeader[headerFound.indexMap.orderNumber] = "ORDER-SYNTH";
    syntheticHeader[headerFound.indexMap.totalValueForDuty] = "$11.11";
    syntheticHeader[headerFound.indexMap.totalCustomsDuties] = "$2.22";
    syntheticHeader[headerFound.indexMap.totalGst] = "$3.33";
    syntheticHeader[headerFound.indexMap.totalProvincialSalesTax] = "$4.44";
    syntheticHeader[headerFound.indexMap.sima] = "$5.55";
    syntheticHeader[headerFound.indexMap.surtax] = "$6.66";
    syntheticHeader[headerFound.indexMap.province] = " ON ";
    const syntheticHeaderRecord = buildHeaderRecords([headerRows[headerFound.rowIndex], syntheticHeader], headerFound.indexMap, 1)[0];
    addIssue(headerIssues, syntheticHeaderRecord.shipmentDate === "10/01/2026" && syntheticHeaderRecord.arrivalDate === "10/02/2026" && syntheticHeaderRecord.releaseDate === "10/03/2026", "Header shipment, arrival, and release date mappings changed.");
    addIssue(headerIssues, syntheticHeaderRecord.orderNumber === "ORDER-SYNTH" && syntheticHeaderRecord.valueForDuty === 11.11 && syntheticHeaderRecord.duty === 2.22 && syntheticHeaderRecord.gst === 3.33 && syntheticHeaderRecord.hst === 4.44 && syntheticHeaderRecord.pst === 0 && syntheticHeaderRecord.sima === 5.55 && syntheticHeaderRecord.surtax === 6.66, "Header source financial fields were combined or mis-mapped.");
    const otherShipment = syntheticHeader.slice();
    otherShipment[headerFound.indexMap.transaction] = "SYNTH-TXN-2";
    otherShipment[headerFound.indexMap.ccn] = "SYNTH-CCN-2";
    const duplicateHeader = syntheticHeader.slice();
    duplicateHeader[headerFound.indexMap.orderNumber] = "ORDER-LATER";
    duplicateHeader[headerFound.indexMap.totalValueForDuty] = "$99.99";
    duplicateHeader[headerFound.indexMap.totalProvincialSalesTax] = "$5.56";
    duplicateHeader[headerFound.indexMap.sima] = "$4.45";
    duplicateHeader[headerFound.indexMap.surtax] = "$3.34";
    duplicateHeader[headerFound.indexMap.province] = "QC";
    const syntheticHeaderRecords = buildHeaderRecords([headerRows[headerFound.rowIndex], syntheticHeader, otherShipment, duplicateHeader], headerFound.indexMap, 1);
    const syntheticHeaderOutput = buildHeaderOutputAoA(syntheticHeaderRecords, "", "", "").slice(4);
    addIssue(headerIssues, syntheticHeaderRecords.length === 2 && syntheticHeaderRecords[0].orderNumber === "ORDER-SYNTH" && syntheticHeaderRecords[1].orderNumber === "ORDER-SYNTH" && syntheticHeaderRecords[1].transactionNumber === trimTransaction(otherShipment[headerFound.indexMap.transaction]), "Duplicate Header grouping changed first-seen shipment order or first-row fields.");
    addIssue(headerIssues, syntheticHeaderOutput[0][9] === 11.11 && syntheticHeaderOutput[0][10] === 2.22 && syntheticHeaderOutput[0][11] === 3.33 &&
      syntheticHeaderOutput[0][12] === 4.44 && syntheticHeaderOutput[0][13] === 0 && syntheticHeaderOutput[0][14] === 10 && syntheticHeaderOutput[0][15] === 10,
    "Duplicate Header tax fields changed or SIMA/Surtax were not summed.");
    const differingHeader = duplicateHeader.slice();
    differingHeader[headerFound.indexMap.totalCustomsDuties] = "$999.99";
    differingHeader[headerFound.indexMap.totalGst] = "$888.88";
    const retained = buildHeaderRecords([headerRows[headerFound.rowIndex], syntheticHeader, differingHeader], headerFound.indexMap, 1);
    addIssue(headerIssues, retained.length === 1 && retained[0].duty === 2.22 && retained[0].gst === 3.33,
      "Conflicting duplicate Header Duty/GST did not retain the first row.");
  } catch (error) {
    headerIssues.push(`Candata Header regression error: ${error.message}`);
    itemIssues.push(`Candata Item regression error: ${error.message}`);
  }

  return { headerIssues, itemIssues };
}

function runGetsHeaderCarryoverRegression(rootDir) {
  const rawHeaderPath = path.join(rootDir, "raw duties header.xlsx");
  if (!fs.existsSync(rawHeaderPath)) {
    return { issues: [`Missing test file: ${rawHeaderPath}`] };
  }

  const headerSpecs = [
    { key: "transaction", label: "Transaction Number", matchers: [/^transaction number$/i] },
    { key: "ccn", label: "Cargo Control Number", matchers: [/cargo control number/i, /^ccn$/i] },
    { key: "portNumber", label: "Port Number", matchers: [/^port number$/i, /^port #$/i] },
    { key: "releaseDate", label: "Release Date", matchers: [/^release date$/i] },
    { key: "totalValueForDuty", label: "Total Value For Duty (CAD)", matchers: [/total value for duty/i] },
    { key: "totalCustomsDuties", label: "Total Customs Duties (CAD)", matchers: [/total customs duties/i] },
    { key: "totalGst", label: "Total GST (CAD)", matchers: [/^total gst/i] },
    { key: "totalProvincialSalesTax", label: "Total Provincial Sales Tax (CAD)", matchers: [/total provincial sales tax/i] },
    { key: "paymentTerms", label: "Payment Terms", matchers: [/^payment terms$/i, /^inco terms$/i] }
  ];

  const { rows: headerRows } = readFirstSheetRows(rawHeaderPath);
  const headerFound = validateRequiredColumns(headerRows, headerSpecs, "raw duties header.xlsx");
  const headerRecords = buildHeaderRecords(headerRows, headerFound.indexMap, headerFound.rowIndex + 1);

  let sourceVfd = 0;
  let sourceDuty = 0;
  let sourceRowCount = 0;
  for (let r = headerFound.rowIndex + 1; r < headerRows.length; r++) {
    const row = headerRows[r] || [];
    const ccn = String(row[headerFound.indexMap.ccn] || "").trim();
    if (!ccn) continue;
    sourceVfd += parseNumberZero(row[headerFound.indexMap.totalValueForDuty]);
    sourceDuty += parseNumberZero(row[headerFound.indexMap.totalCustomsDuties]);
    sourceRowCount++;
  }

  let transformedVfd = 0;
  let transformedDuty = 0;
  for (const rec of headerRecords) {
    transformedVfd += parseNumberZero(rec.valueForDuty);
    transformedDuty += parseNumberZero(rec.duty);
  }

  sourceVfd = round2(sourceVfd);
  sourceDuty = round2(sourceDuty);
  transformedVfd = round2(transformedVfd);
  transformedDuty = round2(transformedDuty);

  const issues = [];
  if (sourceRowCount === 0) {
    issues.push("raw duties header.xlsx produced zero data rows with CCN.");
  }
  if (sourceVfd <= 0) {
    issues.push(`Expected positive source Value for Duty sum, got ${sourceVfd}`);
  }
  if (Math.abs(transformedVfd - sourceVfd) > 0.01) {
    issues.push(`Value for Duty carryover mismatch: transformed=${transformedVfd} source=${sourceVfd}`);
  }
  if (Math.abs(transformedDuty - sourceDuty) > 0.01) {
    issues.push(`Duty carryover mismatch: transformed=${transformedDuty} source=${sourceDuty}`);
  }

  return {
    issues,
    summary: {
      rows: sourceRowCount,
      sourceVfd,
      sourceDuty,
      transformedVfd,
      transformedDuty
    }
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
  const getsHeaderCarryover = runGetsHeaderCarryoverRegression(rootDir);
  const dtHeaderWorkflow = runDtHeaderWorkflowRegression(rootDir);
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

  if (getsHeaderCarryover.issues.length === 0) {
    console.log("GETS Header Carryover: PASS");
    console.log(
      `  - Rows=${getsHeaderCarryover.summary.rows}, VFD=${getsHeaderCarryover.summary.transformedVfd}, Duty=${getsHeaderCarryover.summary.transformedDuty}`
    );
  } else {
    failed = true;
    console.log("GETS Header Carryover: FAIL");
    getsHeaderCarryover.issues.forEach((issue) => console.log(`  - ${issue}`));
  }

  if (dtHeaderWorkflow.issues.length === 0) {
    console.log("D/T Header Workflow: PASS");
    if (dtHeaderWorkflow.skippedFixtureReason) {
      console.log(`  - ${dtHeaderWorkflow.skippedFixtureReason}`);
    }
  } else {
    failed = true;
    console.log("D/T Header Workflow: FAIL");
    dtHeaderWorkflow.issues.forEach((issue) => console.log(`  - ${issue}`));
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
