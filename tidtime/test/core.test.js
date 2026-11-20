const { test } = require("node:test");
const assert = require("node:assert/strict");
const { matchRows, compareNumbers, observationDate } = require("../core");

test("stock counts in January and June use only SKU, observation day and quantity", () => {
  const january = [{ sku: "DEMO", observed: "2026-01-10", quantity: 12 }];
  const june = [{ sku: "DEMO", observed: "2026-06-10", quantity: 9 }];
  const day = (row) => row.observed;
  assert.ok(observationDate(january, day) < observationDate(june, day));
  const [group] = matchRows(january, june, (row) => row.sku);
  assert.deepEqual(compareNumbers(group, ["quantity"]), [
    { field: "quantity", before: 12, after: 9, delta: "-3" },
  ]);
});

test("delta keeps decimal places beyond Number and default decimal precision", () => {
  const [group] = matchRows(
    [{ id: 1, x: "0.0000000000000000000001" }],
    [{ id: 1, x: "1" }],
    (row) => row.id,
  );
  assert.equal(
    compareNumbers(group, ["x"])[0].delta,
    "0.9999999999999999999999",
  );
});
