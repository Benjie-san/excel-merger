(function (root, factory) {
  var api = factory();

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }

  if (root) {
    root.DtHeaderWorkflow = api;
  }
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  function cloneRows(rows) {
    return (rows || []).map(function (row) {
      return Array.isArray(row) ? row.slice() : [];
    });
  }

  function normalizeCell(value) {
    return String(value == null ? "" : value).trim();
  }

  function normalizeHeaderCell(value) {
    return normalizeCell(value).toLowerCase().replace(/\s+/g, " ");
  }

  function normalizeClientName(value) {
    return normalizeCell(value).replace(/\s+/g, " ").toLowerCase();
  }

  function ensureRow(rows, index) {
    while (rows.length <= index) {
      rows.push([]);
    }
    if (!Array.isArray(rows[index])) {
      rows[index] = [];
    }
    return rows[index];
  }

  function ensureCell(row, index) {
    while (row.length <= index) {
      row.push("");
    }
    return row;
  }

  function setCellText(row, index, value) {
    ensureCell(row, index);
    row[index] = normalizeCell(value);
  }

  function rewriteMetadataRows(rows, metadata) {
    var safeMetadata = metadata || {};
    var row0 = ensureRow(rows, 0);
    var row1 = ensureRow(rows, 1);
    var row2 = ensureRow(rows, 2);

    setCellText(row0, 0, "CLIENT:");
    setCellText(row0, 1, safeMetadata.client);
    setCellText(row1, 0, "RPT NAME:");
    setCellText(row1, 1, safeMetadata.reportName);
    setCellText(row2, 0, "RPT DATE :");
    setCellText(row2, 1, safeMetadata.reportDate);

    return rows;
  }

  function rowIncludesAll(row, expectedLabels) {
    var normalized = (row || []).map(normalizeHeaderCell);
    return expectedLabels.every(function (label) {
      return normalized.indexOf(label) !== -1;
    });
  }

  function getExpectedLabelsForMode(mode) {
    if (mode === "header") {
      return ["transaction number"];
    }
    if (mode === "item") {
      return ["transaction number", "goods description"];
    }
    throw new Error('Invalid mode "' + mode + '". Expected "header" or "item".');
  }

  function detectHeaderRowIndex(rows, mode) {
    var expectedLabels = getExpectedLabelsForMode(mode);
    var scanLimit = Math.min((rows || []).length, 10);

    for (var i = 0; i < scanLimit; i++) {
      if (rowIncludesAll(rows[i], expectedLabels) && (mode !== "header" || resolveCcnColumnIndexes(rows[i]).length > 0)) {
        return i;
      }
    }

    return -1;
  }

  function normalizeHeaderRows(rows, mode) {
    var headerRowIndex = detectHeaderRowIndex(rows, mode);
    if (headerRowIndex === -1) {
      throw new Error("Could not locate the " + (mode === "header" ? "DutiesHeader" : "DutiesItem") + " header row.");
    }

    if (headerRowIndex === 3) {
      rows.splice(3, 0, []);
      headerRowIndex = 4;
    }

    return {
      rows: rows,
      headerRowIndex: headerRowIndex
    };
  }

  function findColumnIndex(headerRow, expectedLabel) {
    var expected = normalizeHeaderCell(expectedLabel);
    for (var i = 0; i < (headerRow || []).length; i++) {
      if (normalizeHeaderCell(headerRow[i]) === expected) {
        return i;
      }
    }
    return -1;
  }

  function findColumnIndexAny(headerRow, expectedLabels) {
    for (var i = 0; i < expectedLabels.length; i++) {
      var index = findColumnIndex(headerRow, expectedLabels[i]);
      if (index !== -1) return index;
    }
    return -1;
  }

  function normalizePreparedHeaderInput(headerInput) {
    if (headerInput && Array.isArray(headerInput.rows)) {
      return {
        rows: cloneRows(headerInput.rows),
        headerRowIndex: typeof headerInput.headerRowIndex === "number" ? headerInput.headerRowIndex : detectHeaderRowIndex(headerInput.rows, "header")
      };
    }

    var clonedRows = cloneRows(headerInput);
    return normalizeHeaderRows(clonedRows, "header");
  }

  // Prefer the explicit CCN, then its long label, then Order Number, per row.
  function resolveCcnColumnIndexes(headerRow) {
    return ["CCN", "Cargo Control Number", "Order Number"]
      .map(function (label) { return findColumnIndex(headerRow, label); })
      .filter(function (index) { return index !== -1; });
  }

  function getRecordCcn(row, headerRow) {
    var indexes = resolveCcnColumnIndexes(headerRow);
    for (var i = 0; i < indexes.length; i++) {
      var value = normalizeCell(row[indexes[i]]);
      if (value) return value;
    }
    return "";
  }

  function buildTransactionToCcnMap(headerInput) {
    var normalizedHeader = normalizePreparedHeaderInput(headerInput);
    var headerRowIndex = normalizedHeader.headerRowIndex;
    var headerRow = normalizedHeader.rows[headerRowIndex] || [];
    var transactionIndex = findColumnIndex(headerRow, "Transaction Number");
    var ccnIndex = resolveCcnColumnIndexes(headerRow)[0];

    if (transactionIndex === -1 || ccnIndex === undefined) {
      throw new Error("Header workbook is missing Transaction Number or CCN / Cargo Control Number / Order Number.");
    }

    var lookup = new Map();
    for (var r = headerRowIndex + 1; r < normalizedHeader.rows.length; r++) {
      var row = normalizedHeader.rows[r] || [];
      var transaction = normalizeCell(row[transactionIndex]);
      var ccn = getRecordCcn(row, headerRow);
      if (!transaction || !ccn) {
        continue;
      }
      lookup.set(transaction, ccn);
    }

    return lookup;
  }

  function ensureItemCcnColumn(rows, headerRowIndex) {
    var headerRow = ensureRow(rows, headerRowIndex);
    var ccnIndex = findColumnIndex(headerRow, "CCN");
    if (ccnIndex !== -1) {
      return ccnIndex;
    }

    headerRow.push("CCN");
    ccnIndex = headerRow.length - 1;

    for (var r = 0; r < rows.length; r++) {
      var row = ensureRow(rows, r);
      while (row.length <= ccnIndex) {
        row.push("");
      }
    }

    return ccnIndex;
  }

  function isEmptyRow(row) {
    if (!Array.isArray(row)) return true;
    for (var i = 0; i < row.length; i++) {
      if (normalizeCell(row[i]) !== "") {
        return false;
      }
    }
    return true;
  }

  function parseNumber(value) {
    if (value === undefined || value === null) return null;
    var s = String(value).trim();
    if (s === "") return null;
    if (/^\$?\s*-\s*\$?$/.test(s)) return 0;

    var neg = false;
    if (s.charAt(0) === "(" && s.charAt(s.length - 1) === ")") {
      neg = true;
      s = s.slice(1, -1);
    }
    if (s.charAt(s.length - 1) === "-") {
      neg = true;
      s = s.slice(0, -1);
    }

    s = s.replace(/[$,]/g, "").replace(/\s+/g, "");
    if (s === "") return null;
    if (s === "-") return 0;

    var num = parseFloat(s);
    if (isNaN(num)) return null;
    return neg ? -num : num;
  }

  function roundToDisplay(value) {
    if (value === null || value === undefined || isNaN(value)) return value;
    return Math.round((value + Number.EPSILON) * 100) / 100;
  }

  function findLastNonEmptyRow(rows) {
    for (var i = rows.length - 1; i >= 0; i--) {
      if (!isEmptyRow(rows[i])) {
        return i;
      }
    }
    return -1;
  }

  function cleanTargetCCN(raw) {
    var s = normalizeCell(raw);
    if (s.indexOf("8308") === 0) {
      return s.slice(4);
    }
    return s;
  }

  function lookupClientRates(brokerageRates, clientName) {
    var normalizedClient = normalizeClientName(clientName);
    var lookup = brokerageRates && brokerageRates.clientRateLookup ? brokerageRates.clientRateLookup : {};
    var keys = Object.keys(lookup);

    for (var i = 0; i < keys.length; i++) {
      if (normalizeClientName(keys[i]) === normalizedClient) {
        return {
          matched: true,
          clientKey: keys[i],
          rates: lookup[keys[i]]
        };
      }
    }

    return {
      matched: false,
      clientKey: "",
      rates: null
    };
  }

  function resolveHeaderColumns(headerRow) {
    return {
      transactionNumber: findColumnIndex(headerRow, "Transaction Number"),
      ccn: resolveCcnColumnIndexes(headerRow).length ? resolveCcnColumnIndexes(headerRow)[0] : -1,
      port: findColumnIndexAny(headerRow, ["Port #", "Port Number"]),
      shipmentDate: findColumnIndex(headerRow, "Shipment Date"),
      arrivalDate: findColumnIndex(headerRow, "Arrival Date"),
      releaseDate: findColumnIndex(headerRow, "Release Date"),
      orderNumber: findColumnIndex(headerRow, "Order Number"),
      valueForDuty: findColumnIndex(headerRow, "Value for Duty"),
      duty: findColumnIndex(headerRow, "Duty"),
      gst: findColumnIndexAny(headerRow, ["Gov. Sales Tax", "GST"]),
      hst: findColumnIndex(headerRow, "HST"),
      pst: findColumnIndex(headerRow, "PST"),
      sima: findColumnIndex(headerRow, "SIMA"),
      surtax: findColumnIndex(headerRow, "Surtax"),
      brokerageTotal: findColumnIndex(headerRow, "Brokerage Total"),
      additionalChargesTotal: findColumnIndex(headerRow, "Addl. Charges Total"),
      assessmentTotal: findColumnIndex(headerRow, "Assessment Total"),
      exciseTaxTotal: findColumnIndex(headerRow, "Excise Tax Total"),
      exchangeRate: findColumnIndex(headerRow, "Exchange Rate"),
      incoTerms: findColumnIndex(headerRow, "Inco Terms")
    };
  }

  function assertRequiredHeaderColumns(columns) {
    var missing = [];
    if (columns.transactionNumber === -1) missing.push("Transaction Number");
    if (columns.ccn === -1) missing.push("CCN / Cargo Control Number / Order Number");
    if (columns.shipmentDate === -1) missing.push("Shipment Date");
    if (columns.arrivalDate === -1) missing.push("Arrival Date");
    if (columns.releaseDate === -1) missing.push("Release Date");
    if (columns.valueForDuty === -1) missing.push("Value for Duty");
    if (columns.duty === -1) missing.push("Duty");
    if (columns.gst === -1) missing.push("GST / Gov. Sales Tax");
    if (columns.brokerageTotal === -1) missing.push("Brokerage Total");
    if (columns.exchangeRate === -1) missing.push("Exchange Rate");
    if (missing.length) {
      throw new Error("Header workbook is missing required columns: " + missing.join(", ") + ".");
    }
  }

  function prepareHeaderRowsForModify(options) {
    var rows = cloneRows(options && options.targetRows);
    rewriteMetadataRows(rows, options && options.metadata);
    return normalizeHeaderRows(rows, "header");
  }

  function prepareItemRowsWithCcn(options) {
    var rows = cloneRows(options && options.itemRows);
    var headerInput = options && (options.preparedHeader || options.headerInput || options.headerRows);
    var normalizedHeader = normalizePreparedHeaderInput(headerInput);
    var lookup = buildTransactionToCcnMap(normalizedHeader);

    rewriteMetadataRows(rows, options && options.metadata);
    var normalizedRows = normalizeHeaderRows(rows, "item");
    var headerRowIndex = normalizedRows.headerRowIndex;
    var headerRow = normalizedRows.rows[headerRowIndex] || [];
    var transactionIndex = findColumnIndex(headerRow, "Transaction Number");
    if (transactionIndex === -1) {
      throw new Error("Item workbook is missing Transaction Number.");
    }

    var ccnIndex = ensureItemCcnColumn(normalizedRows.rows, headerRowIndex);
    var unmatchedCount = 0;

    for (var r = headerRowIndex + 1; r < normalizedRows.rows.length; r++) {
      var row = ensureRow(normalizedRows.rows, r);
      var transaction = normalizeCell(row[transactionIndex]);
      var ccn = transaction ? (lookup.get(transaction) || "") : "";
      if (transaction && !ccn) {
        unmatchedCount++;
      }
      row[ccnIndex] = ccn;
    }

    return {
      rows: normalizedRows.rows,
      headerRowIndex: headerRowIndex,
      unmatchedCount: unmatchedCount
    };
  }

  function insertMissingHeaderRows(options) {
    var preparedHeader = normalizePreparedHeaderInput(options && options.preparedHeader);
    var targetRows = cloneRows(preparedHeader.rows);
    var headerRowIndex = preparedHeader.headerRowIndex;
    var sourceRows = cloneRows(options && options.sourceRows);
    var headerRow = targetRows[headerRowIndex] || [];
    var columns = resolveHeaderColumns(headerRow);
    assertRequiredHeaderColumns(columns);

    var ccnStartRowIndex = headerRowIndex + 1;
    var sourceACStartIndex = 2;
    var sourceASStartIndex = 2;

    var COL_AC = 28;
    var COL_AS = 44;

    var lastNonEmptyIndex = findLastNonEmptyRow(targetRows);
    var dataTargetRows = lastNonEmptyIndex >= 0 ? targetRows.slice(0, lastNonEmptyIndex + 1) : targetRows.slice();
    var refSet = new Set();

    for (var r = ccnStartRowIndex; r < dataTargetRows.length; r++) {
      var row = dataTargetRows[r] || [];
      var cleaned = cleanTargetCCN(getRecordCcn(row, headerRow));
      if (cleaned !== "") {
        refSet.add(cleaned);
      }
    }

    var sourceItems = [];
    var sourceSeen = new Set();
    for (r = sourceACStartIndex; r < sourceRows.length; r++) {
      row = sourceRows[r] || [];
      var acRaw = normalizeCell(row[COL_AC]);
      var asRaw = normalizeCell(row[COL_AS]);
      if (acRaw === "" && asRaw === "") continue;
      if (acRaw !== "") {
        if (sourceSeen.has(acRaw)) continue;
        sourceSeen.add(acRaw);
      }
      sourceItems.push({ acRaw: acRaw, asRaw: asRaw });
    }

    var lastExistingRow = lastNonEmptyIndex >= 0 ? (dataTargetRows[lastNonEmptyIndex] || []) : [];
    var targetRowLen = headerRow.length;
    var insertedRows = [];

    for (var i = 0; i < sourceItems.length; i++) {
      var item = sourceItems[i];
      if (item.acRaw !== "" && refSet.has(item.acRaw)) {
        continue;
      }

      var newRow = new Array(targetRowLen).fill("");
      newRow[columns.transactionNumber] = "CLVS";
      newRow[columns.ccn] = item.acRaw;
      if (columns.port !== -1) newRow[columns.port] = lastExistingRow[columns.port] || "";
      newRow[columns.shipmentDate] = lastExistingRow[columns.shipmentDate] || "";
      newRow[columns.arrivalDate] = lastExistingRow[columns.arrivalDate] || "";
      newRow[columns.releaseDate] = lastExistingRow[columns.releaseDate] || "";
      if (columns.orderNumber !== -1) newRow[columns.orderNumber] = item.acRaw;
      newRow[columns.valueForDuty] = item.asRaw;
      [
        columns.duty,
        columns.gst,
        columns.hst,
        columns.pst,
        columns.sima,
        columns.surtax,
        columns.additionalChargesTotal,
        columns.assessmentTotal,
        columns.exciseTaxTotal,
        columns.exchangeRate
      ].forEach(function (index) {
        if (index !== -1) newRow[index] = 0;
      });
      if (columns.incoTerms !== -1) newRow[columns.incoTerms] = "DDP";
      insertedRows.push(newRow);
      if (item.acRaw !== "") {
        refSet.add(item.acRaw);
      }
    }

    var insertAt = lastNonEmptyIndex >= 0 ? lastNonEmptyIndex + 1 : 0;
    var finalRows = targetRows.slice(0, insertAt).concat(insertedRows, targetRows.slice(insertAt));
    return {
      rows: finalRows,
      headerRowIndex: headerRowIndex,
      insertedCount: insertedRows.length,
      generatedRowIndexes: insertedRows.map(function (_, index) { return insertAt + index; })
    };
  }

  function classifyHeaderRow(transaction, ccn) {
    if (ccn.indexOf("8308") === 0) {
      return "PGA";
    }
    if (transaction.indexOf("LV") === 0) {
      return "LVS";
    }
    if (transaction === "CLVS") {
      return "CLVS";
    }
    return "";
  }

  function stableSortRowsByBrokerage(dataRows, brokerageIndex) {
    return dataRows
      .map(function (row, idx) {
        var brokerageNumber = parseNumber(row[brokerageIndex]);
        var sortValue = brokerageNumber === null ? Number.NEGATIVE_INFINITY : brokerageNumber;
        return {
          row: row,
          idx: idx,
          sortValue: sortValue,
          empty: normalizeCell(row[brokerageIndex]) === ""
        };
      })
      .sort(function (a, b) {
        if (a.empty !== b.empty) {
          return a.empty ? 1 : -1;
        }
        if (Math.abs(b.sortValue - a.sortValue) > 0.0000001) {
          return b.sortValue - a.sortValue;
        }
        return a.idx - b.idx;
      })
      .map(function (entry) {
        return entry.row;
      });
  }

  function buildHeaderSummary(rows, headerRowIndex, clientLookup) {
    var headerRow = rows[headerRowIndex] || [];
    var columns = resolveHeaderColumns(headerRow);
    assertRequiredHeaderColumns(columns);

    var counts = {
      pga: 0,
      lvs: 0,
      clvs: 0
    };
    var blankBrokerageCount = 0;
    var totalDutyValue = 0;
    var totalGstValue = 0;
    var additionalTaxTotals = { hst: 0, pst: 0, sima: 0, surtax: 0 };
    var additionalTaxColumns = {
      hst: columns.hst !== -1,
      pst: columns.pst !== -1,
      sima: columns.sima !== -1,
      surtax: columns.surtax !== -1
    };

    for (var r = headerRowIndex + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      if (isEmptyRow(row)) continue;

      var transaction = normalizeCell(row[columns.transactionNumber]);
      var ccn = getRecordCcn(row, headerRow);
      var classification = classifyHeaderRow(transaction, ccn);
      if (classification === "PGA") counts.pga++;
      if (classification === "LVS") counts.lvs++;
      if (classification === "CLVS") counts.clvs++;

      if (normalizeCell(row[columns.brokerageTotal]) === "") {
        blankBrokerageCount++;
      }

      var dutyValue = parseNumber(row[columns.duty]);
      var gstValue = parseNumber(row[columns.gst]);
      if (dutyValue !== null) totalDutyValue += dutyValue;
      if (gstValue !== null) totalGstValue += gstValue;
      Object.keys(additionalTaxTotals).forEach(function (key) {
        if (!additionalTaxColumns[key]) return;
        var value = parseNumber(row[columns[key]]);
        if (value !== null) additionalTaxTotals[key] += value;
      });
    }

    return {
      clientMatched: !!(clientLookup && clientLookup.matched),
      clientKey: clientLookup && clientLookup.clientKey ? clientLookup.clientKey : "",
      counts: counts,
      blankBrokerageCount: blankBrokerageCount,
      totalDutyValue: roundToDisplay(totalDutyValue),
      totalGstValue: roundToDisplay(totalGstValue),
      additionalTaxColumns: additionalTaxColumns,
      totalHstValue: roundToDisplay(additionalTaxTotals.hst),
      totalPstValue: roundToDisplay(additionalTaxTotals.pst),
      totalSimaValue: roundToDisplay(additionalTaxTotals.sima),
      totalSurtaxValue: roundToDisplay(additionalTaxTotals.surtax)
    };
  }

  function buildItemSummary(rows) {
    var normalizedRows = normalizeHeaderRows(cloneRows(rows), "item");
    var headerRowIndex = normalizedRows.headerRowIndex;
    var headerRow = normalizedRows.rows[headerRowIndex] || [];
    var dutyIndex = findColumnIndex(headerRow, "Duty");
    var gstIndex = findColumnIndexAny(headerRow, ["Gov. Sales Tax", "GST"]);
    var additionalTaxIndexes = {
      hst: findColumnIndex(headerRow, "HST"),
      pst: findColumnIndex(headerRow, "PST"),
      sima: findColumnIndex(headerRow, "SIMA"),
      surtax: findColumnIndex(headerRow, "Surtax")
    };

    if (dutyIndex === -1 || gstIndex === -1) {
      throw new Error("Item workbook is missing Duty or GST / Gov. Sales Tax.");
    }

    var totalDutyValue = 0;
    var totalGstValue = 0;
    var additionalTaxTotals = { hst: 0, pst: 0, sima: 0, surtax: 0 };

    for (var r = headerRowIndex + 1; r < normalizedRows.rows.length; r++) {
      var row = normalizedRows.rows[r] || [];
      if (isEmptyRow(row)) continue;

      var dutyValue = parseNumber(row[dutyIndex]);
      var gstValue = parseNumber(row[gstIndex]);
      if (dutyValue !== null) totalDutyValue += dutyValue;
      if (gstValue !== null) totalGstValue += gstValue;
      Object.keys(additionalTaxTotals).forEach(function (key) {
        if (additionalTaxIndexes[key] === -1) return;
        var value = parseNumber(row[additionalTaxIndexes[key]]);
        if (value !== null) additionalTaxTotals[key] += value;
      });
    }

    return {
      totalDutyValue: roundToDisplay(totalDutyValue),
      totalGstValue: roundToDisplay(totalGstValue),
      additionalTaxColumns: {
        hst: additionalTaxIndexes.hst !== -1,
        pst: additionalTaxIndexes.pst !== -1,
        sima: additionalTaxIndexes.sima !== -1,
        surtax: additionalTaxIndexes.surtax !== -1
      },
      totalHstValue: roundToDisplay(additionalTaxTotals.hst),
      totalPstValue: roundToDisplay(additionalTaxTotals.pst),
      totalSimaValue: roundToDisplay(additionalTaxTotals.sima),
      totalSurtaxValue: roundToDisplay(additionalTaxTotals.surtax)
    };
  }

  var validationFields = {
    header: [
      { key: "valueForDuty", label: "Value for Duty" },
      { key: "duty", label: "Duty" },
      { key: "gst", label: "Gov. Sales Tax", aliases: ["Gov. Sales Tax", "GST"] },
      { key: "hst", label: "HST", optional: true },
      { key: "pst", label: "PST", optional: true },
      { key: "sima", label: "SIMA", optional: true },
      { key: "surtax", label: "Surtax", optional: true }
    ],
    item: [
      { key: "quantity", label: "Quantity" },
      { key: "valueForDuty", label: "Value for Duty" },
      { key: "duty", label: "Duty" },
      { key: "valueForTax", label: "Value for Tax" },
      { key: "gst", label: "Gov. Sales Tax", aliases: ["Gov. Sales Tax", "GST"] },
      { key: "hst", label: "HST", optional: true },
      { key: "pst", label: "PST", optional: true },
      { key: "sima", label: "SIMA", optional: true },
      { key: "surtax", label: "Surtax", optional: true }
    ]
  };

  function resolveValidationColumns(headerRow, mode) {
    var fields = validationFields[mode];
    if (!fields) {
      throw new Error('Invalid validation mode "' + mode + '". Expected "header" or "item".');
    }

    var columns = {};
    fields.forEach(function (field) {
      columns[field.key] = findColumnIndexAny(headerRow, field.aliases || [field.label]);
    });
    columns.transactionNumber = findColumnIndex(headerRow, "Transaction Number");
    columns.ccn = findColumnIndex(headerRow, "CCN");
    columns.orderNumber = findColumnIndex(headerRow, "Order Number");
    columns.lineNumber = findColumnIndex(headerRow, "Line #");
    if (columns.lineNumber === -1) {
      columns.lineNumber = findColumnIndex(headerRow, "Line Number");
    }
    return columns;
  }

  function classifyValidationValue(value) {
    if (normalizeCell(value) === "") {
      return "blank";
    }
    // Do not let parseFloat's numeric-prefix parsing turn invalid text into zero.
    var token = normalizeCell(value).replace(/[$,\s]/g, "");
    if (token.charAt(0) === "(" && token.charAt(token.length - 1) === ")") token = token.slice(1, -1);
    if (token === "-") return "zero";
    if (token.charAt(token.length - 1) === "-") token = token.slice(0, -1);
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(token)) return null;
    var numeric = parseNumber(value);
    return numeric === 0 ? "zero" : null;
  }

  function validationRecord(row, rowNumber, columns, mode) {
    var transaction = columns.transactionNumber === -1
      ? ""
      : normalizeCell(row[columns.transactionNumber]);
    var ccn = columns.ccn === -1 ? "" : normalizeCell(row[columns.ccn]);
    var orderNumber = columns.orderNumber === -1 ? "" : normalizeCell(row[columns.orderNumber]);
    var lineNumber = columns.lineNumber === -1 ? "" : normalizeCell(row[columns.lineNumber]);

    return {
      mode: mode,
      rowNumber: rowNumber,
      transactionNumber: transaction,
      ccn: ccn || orderNumber,
      orderNumber: orderNumber,
      lineNumber: lineNumber
    };
  }

  function validateReportRows(rows, mode, options) {
    var normalizedRows = normalizeHeaderRows(cloneRows(rows), mode);
    var headerRowIndex = normalizedRows.headerRowIndex;
    var headerRow = normalizedRows.rows[headerRowIndex] || [];
    var columns = resolveValidationColumns(headerRow, mode);
    var fields = validationFields[mode];
    var missing = fields
      .filter(function (field) { return !field.optional && columns[field.key] === -1; })
      .map(function (field) { return field.label; });

    if (missing.length) {
      return {
        mode: mode,
        headerRowIndex: headerRowIndex,
        error: "Missing required validation columns: " + missing.join(", ") + ".",
        missingColumns: missing,
        rowsChecked: 0,
        issueCount: 0,
        rowsWithIssues: 0,
        blankCount: 0,
        zeroCount: 0,
        ignoredExpectedZeroCount: 0,
        fieldCounts: {},
        issues: []
      };
    }

    var fieldCounts = {};
    var activeFields = fields.filter(function (field) { return columns[field.key] !== -1; });
    activeFields.forEach(function (field) {
      fieldCounts[field.label] = { blank: 0, zero: 0, total: 0 };
    });

    var issues = [];
    var rowsWithIssues = new Set();
    var rowsChecked = 0;
    var blankCount = 0;
    var zeroCount = 0;
    var generatedIssueCount = 0;
    var ignoredExpectedZeroCount = 0;
    var generatedRowNumbers = new Set(options && options.generatedRowNumbers || []);

    for (var r = headerRowIndex + 1; r < normalizedRows.rows.length; r++) {
      var row = normalizedRows.rows[r] || [];
      if (isEmptyRow(row)) continue;
      rowsChecked++;

      var record = validationRecord(row, r + 1, columns, mode);
      record.ccn = getRecordCcn(row, headerRow);
      var origin = generatedRowNumbers.has(r + 1) ? "generated" : "uploaded";
      var classification = mode === "header"
        ? classifyHeaderRow(record.transactionNumber, record.ccn)
        : "";
      activeFields.forEach(function (field) {
        var status = classifyValidationValue(row[columns[field.key]]);
        if (!status) return;

        // CLVS records legitimately carry zero Duty and GST values. Keep the
        // row in the output, but do not report those expected amounts as
        // validation findings. Other Header fields, including Value for Duty,
        // remain subject to the normal blank/zero checks.
        if (
          mode === "header" &&
          status === "zero" &&
          classification === "CLVS" &&
          (field.key === "duty" || field.key === "gst")
        ) {
          ignoredExpectedZeroCount++;
          return;
        }

        fieldCounts[field.label][status]++;
        fieldCounts[field.label].total++;
        if (status === "blank") blankCount++;
        if (status === "zero") zeroCount++;
        if (origin === "generated") generatedIssueCount++;
        rowsWithIssues.add(r);
        issues.push({
          mode: record.mode,
          origin: origin,
          rowNumber: record.rowNumber,
          transactionNumber: record.transactionNumber,
          ccn: record.ccn,
          orderNumber: record.orderNumber,
          lineNumber: record.lineNumber,
          field: field.label,
          type: status,
          value: row[columns[field.key]]
        });
      });
    }

    return {
      mode: mode,
      headerRowIndex: headerRowIndex,
      error: null,
      missingColumns: [],
      rowsChecked: rowsChecked,
      issueCount: issues.length,
      rowsWithIssues: rowsWithIssues.size,
      blankCount: blankCount,
      zeroCount: zeroCount,
      generatedIssueCount: generatedIssueCount,
      ignoredExpectedZeroCount: ignoredExpectedZeroCount,
      uploadedIssueCount: issues.length - generatedIssueCount,
      fieldCounts: fieldCounts,
      issues: issues
    };
  }

  function applyBrokerageAutomation(options) {
    var metadata = options && options.metadata ? options.metadata : {};
    var brokerageRates = options && options.brokerageRates ? options.brokerageRates : null;
    var preparedHeader = normalizePreparedHeaderInput(options && options.preparedHeader);
    var insertedHeader = options && options.sourceRows
      ? insertMissingHeaderRows({
          sourceRows: options.sourceRows,
          preparedHeader: preparedHeader
        })
      : normalizePreparedHeaderInput(preparedHeader);

    var rows = cloneRows(insertedHeader.rows);
    var headerRowIndex = insertedHeader.headerRowIndex;
    rewriteMetadataRows(rows, metadata);

    var headerRow = rows[headerRowIndex] || [];
    var columns = resolveHeaderColumns(headerRow);
    assertRequiredHeaderColumns(columns);

    var clientLookup = lookupClientRates(brokerageRates, metadata.client);
    var reportDate = normalizeCell(metadata.reportDate);
    var dataRows = [];
    var generatedIndexes = new Set(insertedHeader.generatedRowIndexes || []);
    var generatedRows = new Set();

    for (var r = headerRowIndex + 1; r < rows.length; r++) {
      var row = ensureRow(rows, r).slice();
      if (isEmptyRow(row)) continue;

      ensureCell(row, columns.exchangeRate);
      ensureCell(row, columns.brokerageTotal);
      ensureCell(row, columns.releaseDate);

      var transaction = normalizeCell(row[columns.transactionNumber]);
      var ccn = getRecordCcn(row, headerRow);
      var classification = classifyHeaderRow(transaction, ccn);

      row[columns.shipmentDate] = reportDate;
      row[columns.arrivalDate] = reportDate;
      row[columns.releaseDate] = reportDate;
      row[columns.exchangeRate] = 0;

      if (classification) {
        if (clientLookup.matched) {
          var brokerageValue = null;
          if (classification === "PGA") brokerageValue = clientLookup.rates.pga;
          if (classification === "LVS") brokerageValue = clientLookup.rates.lvs;
          if (classification === "CLVS") brokerageValue = clientLookup.rates.clvs;
          row[columns.brokerageTotal] = brokerageValue === null || brokerageValue === undefined ? "" : brokerageValue;
        } else {
          row[columns.brokerageTotal] = "";
        }
      }

      dataRows.push(row);
      if (generatedIndexes.has(r)) generatedRows.add(row);
    }

    dataRows = stableSortRowsByBrokerage(dataRows, columns.brokerageTotal);
    var finalRows = rows.slice(0, headerRowIndex + 1).concat(dataRows);
    var summary = buildHeaderSummary(finalRows, headerRowIndex, clientLookup);

    return {
      rows: finalRows,
      headerRowIndex: headerRowIndex,
      insertedCount: insertedHeader.insertedCount || 0,
      generatedRowNumbers: dataRows.reduce(function (numbers, row, index) {
        if (generatedRows.has(row)) numbers.push(headerRowIndex + index + 2);
        return numbers;
      }, []),
      summary: summary
    };
  }

  function summarizeDtOutputs(options) {
    var headerRows = options && options.headerRows ? options.headerRows : [];
    var itemRows = options && options.itemRows ? options.itemRows : null;
    var headerInput = normalizePreparedHeaderInput({ rows: headerRows });
    var headerSummary = buildHeaderSummary(headerInput.rows, headerInput.headerRowIndex, {
      matched: null,
      clientKey: ""
    });
    var itemSummary = itemRows ? buildItemSummary(itemRows) : null;
    var headerValidation = validateReportRows(headerRows, "header", {
      generatedRowNumbers: options && options.generatedHeaderRowNumbers
    });
    var itemValidation = itemRows ? validateReportRows(itemRows, "item") : null;

    if (headerValidation.error) {
      throw new Error("Header workbook validation failed: " + headerValidation.error);
    }
    if (itemValidation && itemValidation.error) {
      throw new Error("Item workbook validation failed: " + itemValidation.error);
    }

    return {
      header: headerSummary,
      item: itemSummary,
      validation: {
        header: headerValidation,
        item: itemValidation,
        totalIssues: headerValidation.issueCount + (itemValidation ? itemValidation.issueCount : 0),
        hasIssues: headerValidation.issueCount > 0 || !!(itemValidation && itemValidation.issueCount > 0)
      },
      compare: {
        dutyMatch: itemSummary ? Math.abs(headerSummary.totalDutyValue - itemSummary.totalDutyValue) <= 0.0001 : false,
        gstMatch: itemSummary ? Math.abs(headerSummary.totalGstValue - itemSummary.totalGstValue) <= 0.0001 : false,
        hstMatch: itemSummary && headerSummary.additionalTaxColumns.hst && itemSummary.additionalTaxColumns.hst
          ? Math.abs(headerSummary.totalHstValue - itemSummary.totalHstValue) <= 0.0001 : null,
        pstMatch: itemSummary && headerSummary.additionalTaxColumns.pst && itemSummary.additionalTaxColumns.pst
          ? Math.abs(headerSummary.totalPstValue - itemSummary.totalPstValue) <= 0.0001 : null,
        simaMatch: itemSummary && headerSummary.additionalTaxColumns.sima && itemSummary.additionalTaxColumns.sima
          ? Math.abs(headerSummary.totalSimaValue - itemSummary.totalSimaValue) <= 0.0001 : null,
        surtaxMatch: itemSummary && headerSummary.additionalTaxColumns.surtax && itemSummary.additionalTaxColumns.surtax
          ? Math.abs(headerSummary.totalSurtaxValue - itemSummary.totalSurtaxValue) <= 0.0001 : null
      }
    };
  }

  return {
    detectHeaderRowIndex: detectHeaderRowIndex,
    getRecordCcn: getRecordCcn,
    prepareHeaderRowsForModify: prepareHeaderRowsForModify,
    prepareItemRowsWithCcn: prepareItemRowsWithCcn,
    insertMissingHeaderRows: insertMissingHeaderRows,
    applyBrokerageAutomation: applyBrokerageAutomation,
    summarizeDtOutputs: summarizeDtOutputs,
    validateReportRows: validateReportRows
  };
});
