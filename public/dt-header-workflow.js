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

  function ensureRow(rows, index) {
    while (rows.length <= index) {
      rows.push([]);
    }
    if (!Array.isArray(rows[index])) {
      rows[index] = [];
    }
    return rows[index];
  }

  function setCellText(row, index, value) {
    while (row.length <= index) {
      row.push("");
    }
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

  return {
    detectHeaderRowIndex: detectHeaderRowIndex,
    prepareHeaderRowsForModify: prepareHeaderRowsForModify,
    prepareItemRowsWithCcn: prepareItemRowsWithCcn
  };
});
