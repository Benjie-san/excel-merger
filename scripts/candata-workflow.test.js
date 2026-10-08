const assert = require("node:assert/strict");
const { routeProvinceTax, hstColumnSpec, buildHeaderGroups } = require("../public/candata-workflow.js");

const hstSpec = hstColumnSpec();
assert.equal(hstSpec.optional, true);
for (const alias of ["HST", "HST (CAD)", "Total HST (CAD)", "Harmonized Sales Tax", "Total Harmonized Sales Tax (CAD)"]) {
  assert.ok(hstSpec.matchers.some((matcher) => matcher.test(alias.toLowerCase())), alias);
}

for (const province of ["NB", "NL", "NS", "ON", "PE"]) {
  assert.deepEqual(routeProvinceTax(` ${province.toLowerCase()} `, 7.89),
    { pst: 0, hst: 7.89, moved: true, collision: false });
}
assert.deepEqual(routeProvinceTax(" QC ", 7.89, 2),
  { pst: 7.89, hst: 2, moved: false, collision: false });
assert.deepEqual(routeProvinceTax("ON", 7.89, 2.11),
  { pst: 0, hst: 10, moved: true, collision: true });
assert.deepEqual(routeProvinceTax("ON", -3.25, 1),
  { pst: 0, hst: -2.25, moved: true, collision: true });
assert.deepEqual(routeProvinceTax("ON", 0, 4),
  { pst: 0, hst: 4, moved: false, collision: false });

const firstTax = routeProvinceTax("ON", 5, 0);
const secondTax = routeProvinceTax("ON", 0, 5);
const makeRecord = (ccn, duty, gst, tax, sima, surtax) => ({
  ccn, duty, gst, pst: tax.pst, hst: tax.hst, sima, surtax,
  orderNumber: `ORDER-${ccn}`
});
const entries = [
  { row: 5, ccn: " CCN-1 ", transaction: "TX-1", province: "ON", duty: 2, gst: 3,
    pst: 5, hst: 0, record: makeRecord("CCN-1", 2, 3, firstTax, 1.25, 2.25) },
  { row: 6, ccn: "CCN-2", transaction: "TX-2", province: "QC", duty: 7, gst: 8,
    pst: 1, hst: 0, record: makeRecord("CCN-2", 7, 8, routeProvinceTax("QC", 1), 0, 0) },
  { row: 7, ccn: "ccn-1", transaction: "TX-OTHER", province: "ON", duty: 99, gst: 88,
    pst: 0, hst: 5, record: makeRecord("CCN-1", 99, 88, secondTax, 3.5, 4.5) }
];
const result = buildHeaderGroups(entries);
assert.equal(result.records.length, 2);
assert.deepEqual(result.records.map((row) => row.orderNumber), ["ORDER-CCN-1", "ORDER-CCN-2"]);
assert.deepEqual(
  [result.records[0].duty, result.records[0].gst, result.records[0].pst, result.records[0].hst,
    result.records[0].sima, result.records[0].surtax],
  [2, 3, 0, 5, 4.75, 6.75]
);
assert.equal(result.duplicateGroups.length, 1);
assert.deepEqual(result.duplicateGroups[0].equality,
  { duty: false, gst: false, pst: false, hst: false });
assert.deepEqual(result.duplicateGroups[0].rows.map((row) => [row.row, row.transaction, row.pst, row.hst]),
  [[5, "TX-1", 5, 0], [7, "TX-OTHER", 0, 5]]);
assert.equal(firstTax.hst, secondTax.hst, "Routed equality must not hide original differences.");

console.log("Candata production workflow: PASS");
