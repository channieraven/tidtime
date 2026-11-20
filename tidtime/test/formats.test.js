const { test } = require("node:test");
const assert = require("node:assert/strict");
const ExcelJS = require("exceljs");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const {
  openExcel,
  readSheet,
  setSurveyYear,
  sheetOptions,
} = require("../excel");
const { pairRounds, compareTree } = require("../compare");

test("จับคู่ข้าม Snapshot และแบบฟอร์มเมื่อจุดเดิมย้ายหมวด โดยคงชีตและแถวทั้งคู่", async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "tidtime-formats-"));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const db = new ExcelJS.Workbook(),
    field = new ExcelJS.Workbook();
  function table(name, header, rows) {
    const s = db.addWorksheet(name);
    s.getRow(3).values = header;
    rows.forEach((r) => s.addRow(r));
  }
  table(
    "trees",
    ["tree_id", "plot_id", "tree_code", "tree_subtype_id"],
    [["uuid", "P99", "P99A01987", "sp"]],
  );
  table("tree_subtypes", ["tree_subtype_id", "name_th"], [["sp", "สัก"]]);
  table(
    "growth_forest",
    ["tree_id", "survey_date", "notes", "height_m", "rcd_cm"],
    [["uuid", new Date("2025-01-10"), "เดิม", 2, 4]],
  );
  function form(fruit) {
    const s = field.addWorksheet(
      fruit ? "ผลผลิตของผลไม้+กล้วย" : "ไม้ป่า+ยางพารา+ไผ่",
    );
    const cells = {
      A3: "วันที่",
      B3: "10/01/2569",
      A4: "ชื่อแปลง",
      B4: "P99",
      A7: "ต้นที่",
      B7: "เลขแท็กต้นไม้",
      C7: "เลขรหัส\nต้นไม้",
      ...(fruit
        ? {
            D7: "ชนิดพันธุ์ไม้หรือ\nชนิดกล้วย",
            F7: "ไม้ผล",
            F8: "ความโต",
            G8: "ความสูง",
            I7: "กล้วย",
            Q7: "หมายเหตุ",
            A10: 987,
            C10: "P99B01987",
            D10: "หว้า",
            G10: 0.7,
            F10: 1,
          }
        : {
            D7: "ชนิดพันธุ์",
            F7: "ไม้ป่า/ยางพารา",
            F8: "ความโตที่ระดับคอราก",
            K7: "ความสูง\n(ม.)",
            M7: "หมายเหตุ",
            A10: 1,
            C10: "P99A01001",
            D10: "สัก",
            B12: "ไม้ป่า 1",
          }),
    };
    for (const [addr, value] of Object.entries(cells))
      s.getCell(addr).value = value;
  }
  form(false);
  form(true);
  async function source(wb, file) {
    const p = path.join(dir, file);
    await wb.xlsx.writeFile(p);
    return openExcel(p);
  }
  const [a, b] = await Promise.all([
    source(db, "snapshot.xlsx"),
    source(field, "form.xlsx"),
  ]);
  const prior = readSheet(a, sheetOptions(a)[0]),
    next = readSheet(b, sheetOptions(b)[0]);
  assert.equal(prior.year, "");
  assert.throws(() => setSurveyYear(next, ""), /ปีรอบ/);
  const points = pairRounds(
    setSurveyYear(prior, "2568"),
    setSurveyYear(next, "2569"),
  );
  assert.equal(points.length, 2);
  const match = points.find((p) => p.no === "987");
  assert.equal(match.before[0].sourceSheet, "growth_forest");
  assert.equal(match.before[0].line, 4);
  assert.equal(match.after[0].sourceSheet, "ผลผลิตของผลไม้+กล้วย");
  assert.equal(match.after[0].line, 10);
  assert.equal(match.after[0].date, "10/01/2569");
  assert.match(compareTree(match).status, /รอตรวจ/);
  assert.equal(compareTree(match).lines.length, 2);
});
