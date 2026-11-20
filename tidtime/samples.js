const { matchRows, compareNumbers } = require("./core");

const january = [{ sku: "PAPER-A4", boxes: 12 }];
const june = [{ sku: "PAPER-A4", boxes: 9 }];
for (const group of matchRows(january, june, (item) => item.sku)) {
  console.log(group.key, compareNumbers(group, ["boxes"]));
}
