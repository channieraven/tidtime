// ข้อมูลชั้นวางอุปกรณ์ ใช้ชื่อคอลัมน์ตามสัญญาของหน้าจอนำเข้าเพื่อทดสอบแอป
const ExcelJS = require("exceljs");
const fs = require("node:fs");
const path = require("node:path");
const { columns } = require("../excel");

// ตำแหน่งบนชั้น, SKU, ประเภทกล่อง, ความสูง, ความกว้าง
const previousStock = [
  [1, "BOX-RED", "กล่องจดหมาย", 8, 30],
  [2, "BOX-BLUE", "กล่องเอกสาร", 12, 40],
  [3, "BOX-GREEN", "กล่องเครื่องเขียน", 15, 24],
  [4, "BOX-WHITE", "กล่องอะไหล่", 19, 32],
  [5, "BIN-XL", "ลังกระดาษ", 26, 60],
];
const current = [
  [3, "BOX-GREEN", "กล่องเครื่องเขียน", 16, 22.75],
  [1, "BOX-RED", "กล่องจดหมาย", 7.5, 30],
  [4, "BOX-WHITE", "กล่องอะไหล่", null, 29],
  [2, "TRAY-GREY", "ถาดเอกสาร", 4, 18],
  [6, "BIN-S", "ลังพลาสติก", 13, 25],
];
async function sample(directory, year, rows) {
  const wb = new ExcelJS.Workbook();
  function sheet(name, records) {
    const s = wb.addWorksheet(name);
    s.addRow(Object.values(columns));
    records.forEach(([no, id, category, height, width]) =>
      s.addRow([
        "RACK-A",
        no,
        id,
        category,
        year,
        year === 2568
          ? null
          : no === 2
            ? new Date("2026-03-10T00:00:00Z")
            : "10-03-2569",
        height,
        width,
        no === 2 ? "นับซ้ำที่ชั้นวาง" : "",
      ]),
    );
    s.columns.forEach((c) => {
      c.width = 24;
    });
    s.getRow(1).font = { bold: true };
    s.views = [{ state: "frozen", ySplit: 1 }];
  }
  sheet("ชั้นวาง", rows);
  if (year === 2569)
    sheet("ทดสอบคีย์ซ้ำ", [...rows, rows[1], [null, "BIN-M", "ถาดเทา", 1, 2]]);
  await wb.xlsx.writeFile(path.join(directory, "RACK-A-" + year + ".xlsx"));
}
async function writeSamples(directory) {
  fs.mkdirSync(directory, { recursive: true });
  await Promise.all([
    sample(directory, 2568, previousStock),
    sample(directory, 2569, current),
  ]);
}
module.exports = { writeSamples };
