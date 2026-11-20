const ExcelJS = require("exceljs");
const fs = require("node:fs");
const path = require("node:path");
const { columns } = require("./excel");
const directory = path.join(__dirname, "examples");
const specimens = [
  [1, "T001", "มะม่วง", 1.5, 3],
  [2, "T002", "มะขามป้อม", 1.8, 4],
  [3, "T003", "สัก", 2, 5],
  [4, "T004", "ยางนา", 2, 4],
  [5, "T005", "แดง", 1, 3],
];
const current = [
  [3, "T003", "สัก", 2.4, 4.75],
  [1, "T001", "มะม่วง", 1, 3],
  [4, "T004", "ยางนา", null, 3.6],
  [2, "T019", "หว้า", 0.7, 1.4],
  [6, "T006", "ประดู่", 1, 2],
];
async function sample(year, rows) {
  const wb = new ExcelJS.Workbook();
  function sheet(name, records) {
    const s = wb.addWorksheet(name);
    s.addRow(Object.values(columns));
    records.forEach(([no, id, species, h, d]) =>
      s.addRow([
        "P01",
        no,
        id,
        species,
        year,
        year === 2568
          ? null
          : no === 2
            ? new Date("2026-07-25T00:00:00Z")
            : "25-07-2569",
        h,
        d,
        no === 2 ? "สำรวจซ้ำจุดเดิม" : "",
      ]),
    );
    s.columns.forEach((c) => {
      c.width = 24;
    });
    s.getRow(1).font = { bold: true };
    s.views = [{ state: "frozen", ySplit: 1 }];
  }
  sheet("สำรวจ", rows);
  if (year === 2569)
    sheet("ทดสอบคีย์ซ้ำ", [...rows, rows[1], [null, "T007", "หว้า", 1, 2]]);
  await wb.xlsx.writeFile(path.join(directory, "P01-" + year + ".xlsx"));
}
fs.mkdirSync(directory, { recursive: true });
Promise.all([sample(2568, specimens), sample(2569, current)]).catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
