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
      return ["transaction number", "ccn"];
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
      if (rowIncludesAll(rows[i], expectedLabels)) {
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

  function buildTransactionToCcnMap(headerInput) {
    var normalizedHeader = normalizePreparedHeaderInput(headerInput);
    var headerRowIndex = normalizedHeader.headerRowIndex;
    var headerRow = normalizedHeader.rows[headerRowIndex] || [];
    var transactionIndex = findColumnIndex(headerRow, "Transaction Number");
    var ccnIndex = findColumnIndex(headerRow, "CCN");

    if (transactionIndex === -1 || ccnIndex === -1) {
      throw new Error("Header workbook is missing Transaction Number or CCN.");
    }

    var lookup = new Map();
    for (var r = headerRowIndex + 1; r < normalizedHeader.rows.length; r++) {
      var row = normalizedHeader.rows[r] || [];
      var transaction = normalizeCell(row[transactionIndex]);
      var ccn = normalizeCell(row[ccnIndex]);
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
      ccn: findColumnIndex(headerRow, "CCN"),
      shipmentDate: findColumnIndex(headerRow, "Shipment Date"),
      arrivalDate: findColumnIndex(headerRow, "Arrival Date"),
      releaseDate: findColumnIndex(headerRow, "Release Date"),
      valueForDuty: findColumnIndex(headerRow, "Value for Duty"),
      duty: findColumnIndex(headerRow, "Duty"),
      gst: findColumnIndex(headerRow, "Gov. Sales Tax"),
      brokerageTotal: findColumnIndex(headerRow, "Brokerage Total"),
      exchangeRate: findColumnIndex(headerRow, "Exchange Rate")
    };
  }

  function assertRequiredHeaderColumns(columns) {
    var missing = [];
    if (columns.transactionNumber === -1) missing.push("Transaction Number");
    if (columns.ccn === -1) missing.push("CCN");
    if (columns.shipmentDate === -1) missing.push("Shipment Date");
    if (columns.arrivalDate === -1) missing.push("Arrival Date");
    if (columns.releaseDate === -1) missing.push("Release Date");
    if (columns.valueForDuty === -1) missing.push("Value for Duty");
    if (columns.duty === -1) missing.push("Duty");
    if (columns.gst === -1) missing.push("Gov. Sales Tax");
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

    var ccnStartRowIndex = headerRowIndex + 1;
    var sourceACStartIndex = 2;
    var sourceASStartIndex = 2;

    var COL_AC = 28;
    var COL_AS = 44;
    var COL_A = 0;
    var COL_B = 1;
    var COL_C = 2;
    var COL_D = 3;
    var COL_E = 4;
    var COL_F = 5;
    var COL_H = 7;
    var COL_J = 9;
    var COL_K = 10;
    var COL_Q = 16;
    var COL_R = 17;

    var lastNonEmptyIndex = findLastNonEmptyRow(targetRows);
    var dataTargetRows = lastNonEmptyIndex >= 0 ? targetRows.slice(0, lastNonEmptyIndex + 1) : targetRows.slice();
    var refSet = new Set();

    for (var r = ccnStartRowIndex; r < dataTargetRows.length; r++) {
      var row = dataTargetRows[r] || [];
      var cleaned = cleanTargetCCN(row[COL_B]);
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
    var headerRow = dataTargetRows[headerRowIndex] || [];
    var targetRowLen = Math.max(headerRow.length, COL_R + 1, COL_Q + 1, COL_J + 1, 18);
    var insertedRows = [];

    for (var i = 0; i < sourceItems.length; i++) {
      var item = sourceItems[i];
      if (item.acRaw !== "" && refSet.has(item.acRaw)) {
        continue;
      }

      var newRow = new Array(targetRowLen).fill("");
      newRow[COL_A] = "CLVS";
      newRow[COL_B] = item.acRaw;
      newRow[COL_C] = lastExistingRow[COL_C] || "";
      newRow[COL_D] = lastExistingRow[COL_D] || "";
      newRow[COL_E] = lastExistingRow[COL_E] || "";
      newRow[COL_F] = lastExistingRow[COL_F] || "";
      newRow[COL_H] = item.acRaw;
      newRow[COL_J] = item.asRaw;
      for (var c = COL_K; c <= COL_Q; c++) {
        newRow[c] = 0;
      }
      newRow[COL_R] = "DDP";
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
      insertedCount: insertedRows.length
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

    for (var r = headerRowIndex + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      if (isEmptyRow(row)) continue;

      var transaction = normalizeCell(row[columns.transactionNumber]);
      var ccn = normalizeCell(row[columns.ccn]);
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
    }

    return {
      clientMatched: !!(clientLookup && clientLookup.matched),
      clientKey: clientLookup && clientLookup.clientKey ? clientLookup.clientKey : "",
      counts: counts,
      blankBrokerageCount: blankBrokerageCount,
      totalDutyValue: roundToDisplay(totalDutyValue),
      totalGstValue: roundToDisplay(totalGstValue)
    };
  }

  function buildItemSummary(rows) {
    var normalizedRows = normalizeHeaderRows(cloneRows(rows), "item");
    var headerRowIndex = normalizedRows.headerRowIndex;
    var headerRow = normalizedRows.rows[headerRowIndex] || [];
    var dutyIndex = findColumnIndex(headerRow, "Duty");
    var gstIndex = findColumnIndex(headerRow, "Gov. Sales Tax");

    if (dutyIndex === -1 || gstIndex === -1) {
      throw new Error("Item workbook is missing Duty or Gov. Sales Tax.");
    }

    var totalDutyValue = 0;
    var totalGstValue = 0;

    for (var r = headerRowIndex + 1; r < normalizedRows.rows.length; r++) {
      var row = normalizedRows.rows[r] || [];
      if (isEmptyRow(row)) continue;

      var dutyValue = parseNumber(row[dutyIndex]);
      var gstValue = parseNumber(row[gstIndex]);
      if (dutyValue !== null) totalDutyValue += dutyValue;
      if (gstValue !== null) totalGstValue += gstValue;
    }

    return {
      totalDutyValue: roundToDisplay(totalDutyValue),
      totalGstValue: roundToDisplay(totalGstValue)
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

    for (var r = headerRowIndex + 1; r < rows.length; r++) {
      var row = ensureRow(rows, r).slice();
      if (isEmptyRow(row)) continue;

      ensureCell(row, columns.exchangeRate);
      ensureCell(row, columns.brokerageTotal);
      ensureCell(row, columns.releaseDate);

      var transaction = normalizeCell(row[columns.transactionNumber]);
      var ccn = normalizeCell(row[columns.ccn]);
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
    }

    dataRows = stableSortRowsByBrokerage(dataRows, columns.brokerageTotal);
    var finalRows = rows.slice(0, headerRowIndex + 1).concat(dataRows);
    var summary = buildHeaderSummary(finalRows, headerRowIndex, clientLookup);

    return {
      rows: finalRows,
      headerRowIndex: headerRowIndex,
      insertedCount: insertedHeader.insertedCount || 0,
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

    return {
      header: headerSummary,
      item: itemSummary,
      compare: {
        dutyMatch: itemSummary ? Math.abs(headerSummary.totalDutyValue - itemSummary.totalDutyValue) <= 0.0001 : false,
        gstMatch: itemSummary ? Math.abs(headerSummary.totalGstValue - itemSummary.totalGstValue) <= 0.0001 : false
      }
    };
  }

  return {
    detectHeaderRowIndex: detectHeaderRowIndex,
    prepareHeaderRowsForModify: prepareHeaderRowsForModify,
    prepareItemRowsWithCcn: prepareItemRowsWithCcn,
    insertMissingHeaderRows: insertMissingHeaderRows,
    applyBrokerageAutomation: applyBrokerageAutomation,
    summarizeDtOutputs: summarizeDtOutputs
  };
});
