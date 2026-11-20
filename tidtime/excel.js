const ExcelJS = require("exceljs");
const path = require("node:path");

const columns = {
  plot: "รหัสแปลง",
  no: "ลำดับต้น",
  id: "รหัสต้นไม้",
  species: "ชนิดพันธุ์",
  year: "ปีรอบสำรวจ (พ.ศ.)",
  date: "วันที่สำรวจ",
  height: "ความสูง (ม.)",
  rcd: "ความโตโคนต้น RCD (ซม.)",
  note: "หมายเหตุ",
};

function valueOf(cell) {
  let value = cell.value;
  if (
    value &&
    typeof value === "object" &&
    ("formula" in value || "sharedFormula" in value)
  )
    value = value.result;
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (value?.richText) return value.richText.map((part) => part.text).join("");
  return value?.error ?? value?.text ?? value ?? "";
}
function displayDate(raw) {
  if (!raw) return "ไม่ระบุวันที่";
  const iso = /^(\d{4})-(\d\d)-(\d\d)$/.exec(String(raw));
  return iso ? iso[3] + "-" + iso[2] + "-" + (+iso[1] + 543) : String(raw);
}
async function openExcel(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  return { file: path.basename(file), workbook };
}
function readSheet(source, name) {
  const sheet = source.workbook.getWorksheet(name);
  if (!sheet) throw Error("ไม่พบชีตที่เลือก");
  const indices = {};
  sheet.getRow(1).eachCell((cell, index) => {
    const title = String(valueOf(cell)).trim();
    if (!Object.values(columns).includes(title)) return;
    if (indices[title]) throw Error("หัวตารางซ้ำ " + title);
    indices[title] = index;
  });
  const missing = Object.entries(columns).filter(
    ([f, title]) => f !== "note" && !indices[title],
  );
  if (missing.length)
    throw Error("ขาดคอลัมน์ " + missing.map((pair) => pair[1]).join(", "));
  const rows = [];
  sheet.eachRow((line, number) => {
    if (number === 1) return;
    const row = { line: number };
    for (const [field, title] of Object.entries(columns)) {
      row[field] = indices[title] ? valueOf(line.getCell(indices[title])) : "";
    }
    if (Object.keys(columns).some((f) => String(row[f]).trim())) rows.push(row);
  });
  if (!rows.length) throw Error("ชีตไม่มีข้อมูล");
  const plots = new Set(rows.map((r) => String(r.plot).trim()).filter(Boolean));
  const years = new Set(rows.map((r) => String(r.year).trim()));
  for (const r of rows)
    if (!/^\d{4}$/.test(String(r.year).trim()))
      throw Error("ปีรอบสำรวจไม่ครบหรือไม่ถูกต้องที่แถว " + r.line);
  if (plots.size !== 1)
    throw Error(
      plots.size
        ? "พบหลายแปลงในชีต"
        : "ขาดรหัสแปลงที่แถว " + rows.map((r) => r.line).join(", "),
    );
  if (years.size !== 1)
    throw Error("พบหลายรอบสำรวจในชีต " + [...years].join(", "));
  return {
    file: source.file,
    sheet: name,
    plot: [...plots][0],
    year: [...years][0],
    rows,
  };
}
module.exports = {
  openExcel,
  readSheet,
  columns,
  displayDate,
};
