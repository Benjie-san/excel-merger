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

  return { routeProvinceTax, hstColumnSpec, buildHeaderGroups };
});
