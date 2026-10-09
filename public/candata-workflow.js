(function (root, factory) {
  const workflow = factory();
  if (typeof module === "object" && module.exports) module.exports = workflow;
  if (root) root.CandataWorkflow = workflow;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  const hstProvinces = new Set(["NB", "NL", "NS", "ON", "PE"]);
  const taxFields = ["duty", "gst", "pst", "hst"];

  function sumMoney(a, b) {
    return Math.round((a + b) * 100) / 100;
  }

  function routeProvinceTax(provinceValue, pst, hst = 0) {
    const province = String(provinceValue || "").trim().toUpperCase();
    if (!hstProvinces.has(province) || pst === 0) {
      return { pst, hst, moved: false, collision: false };
    }
    return { pst: 0, hst: sumMoney(hst, pst), moved: true, collision: hst !== 0 };
  }

  function hstColumnSpec() {
    return {
      key: "hst",
      label: "HST (CAD)",
      optional: true,
      matchers: [/^(?:total )?(?:hst|harmonized sales tax)(?: \(cad\))?$/i]
    };
  }

  // Reorder only the processing copy. Every detected column, including fields after
  // the tax block, must follow its source value to the new position.
  function normalizeTaxColumns(rows, found, gstKey, pstKey) {
    const indexMap = found.indexMap;
    const taxKeys = [gstKey, "hst", pstKey, "surtax", "sima"];
    for (const key of [gstKey, pstKey, "surtax", "sima"]) {
      if (!Number.isInteger(indexMap[key]) || indexMap[key] < 0) {
        throw new Error(`Cannot normalize Candata taxes: ${key} column is missing.`);
      }
    }
    const sourceHadHst = found.sourceHadHst ?? indexMap.hst >= 0;
    const taxIndices = new Set(taxKeys.map((key) => indexMap[key]).filter((index) => index >= 0));
    const width = rows.reduce((max, row) => Math.max(max, row?.length || 0),
      Math.max(...Object.values(indexMap).filter((index) => index >= 0)) + 1);
    const order = [];
    let insertAt = 0;
    for (let index = 0; index < width; index++) {
      if (index < indexMap[gstKey] && !taxIndices.has(index)) insertAt++;
      if (!taxIndices.has(index)) order.push(index);
    }
    order.splice(insertAt, 0, ...taxKeys.map((key) => indexMap[key] >= 0 ? indexMap[key] : null));
    const newIndexByOld = new Map(order.map((oldIndex, newIndex) => [oldIndex, newIndex]));
    const newIndexMap = Object.fromEntries(Object.entries(indexMap).map(([key, oldIndex]) => [
      key, oldIndex >= 0 ? newIndexByOld.get(oldIndex) : -1
    ]));
    newIndexMap.hst = insertAt + 1;

    const normalizedRows = rows.map((row, rowNumber) => {
      if (!Array.isArray(row)) return row;
      if (rowNumber < found.rowIndex) return row.slice();
      const normalized = [];
      normalized.length = order.reduce((length, oldIndex, newIndex) =>
        oldIndex !== null && oldIndex < row.length ? newIndex + 1 : length, 0);
      for (let oldIndex = 0; oldIndex < row.length; oldIndex++) {
        if (oldIndex in row) normalized[newIndexByOld.get(oldIndex)] = row[oldIndex];
      }
      if (indexMap.hst < 0) {
        if (rowNumber === found.rowIndex) normalized[newIndexMap.hst] = hstColumnSpec().label;
        else if (row.some((value) => value !== undefined && value !== null && String(value).trim() !== "")) {
          normalized[newIndexMap.hst] = 0;
        }
      }
      return normalized;
    });
    return { ...found, rows: normalizedRows, indexMap: newIndexMap, sourceHadHst };
  }

  function buildHeaderGroups(entries) {
    const byCcn = new Map();
    const records = [];
    for (const entry of entries) {
      const key = String(entry.ccn).trim().toLowerCase();
      let group = byCcn.get(key);
      const source = {
        row: entry.row,
        transaction: entry.transaction,
        province: entry.province,
        duty: entry.duty,
        gst: entry.gst,
        pst: entry.pst,
        hst: entry.hst,
        raw: entry.raw
      };
      if (!group) {
        group = { ccn: entry.ccn, rows: [], record: entry.record };
        byCcn.set(key, group);
        records.push(entry.record);
      } else {
        group.record.sima = sumMoney(group.record.sima, entry.record.sima);
        group.record.surtax = sumMoney(group.record.surtax, entry.record.surtax);
      }
      group.rows.push(source);
    }
    const duplicateGroups = Array.from(byCcn.values())
      .filter((group) => group.rows.length > 1)
      .map((group) => ({
        ccn: group.ccn,
        rows: group.rows,
        equality: Object.fromEntries(taxFields.map((field) => [
          field, group.rows.every((row) => row[field] === group.rows[0][field])
        ]))
      }));
    return { records, duplicateGroups };
  }

  return { routeProvinceTax, hstColumnSpec, normalizeTaxColumns, buildHeaderGroups };
});
