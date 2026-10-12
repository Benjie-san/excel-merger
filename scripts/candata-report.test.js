const assert = require("node:assert/strict");
const report = require("../public/candata-report.js");

const same = { ccn: "CCN-100", rows: [{ transaction: "Txn-A" }],
  equality: { duty: true, gst: true, pst: true, hst: true } };
const differing = { ccn: "CCN-200", rows: [{ transaction: "Txn-B" }],
  equality: { duty: false, gst: true, pst: true, hst: true } };
const groups = [same, differing];

assert.equal(report.PAGE_SIZE, 25);
assert.equal(report.isDiffering(same), false);
assert.equal(report.isDiffering(differing), true);
assert.deepEqual(report.matchingGroups(groups, "ccn-200", "all"), [differing]);
assert.deepEqual(report.matchingGroups(groups, "txn-a", "all"), [same]);
assert.deepEqual(report.matchingGroups(groups, "", "differing"), [differing]);
assert.deepEqual(report.matchingGroups(groups, "txn-a", "differing"), []);

const findings = Array.from({ length: 53 }, (_, index) => index + 1);
assert.deepEqual(report.pageOf(findings, 1).items, findings.slice(0, 25));
assert.deepEqual(report.pageOf(findings, 2).items, findings.slice(25, 50));
assert.deepEqual(report.pageOf(findings, 99).items, findings.slice(50));
assert.equal(report.pageOf(findings, 99).page, 3);
assert.equal(report.pageOf([1], 3).page, 1);
assert.equal(report.pageOf([], 5).page, 1);

console.log("Candata report filtering and pagination passed.");
