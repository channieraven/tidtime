const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { openExcel, readSheet, displayDate } = require("../excel");
const { pairRounds, compareTree } = require("../compare");
const { Store } = require("../store");
const { writeSamples } = require("../samples");
const { before, after } = require("node:test");
let directory;
before(async () => {
  directory = fs.mkdtempSync(path.join(os.tmpdir(), "tidtime-samples-"));
  await writeSamples(directory);
});
after(() => fs.rmSync(directory, { recursive: true, force: true }));
const file = (y) => path.join(directory, "P01-" + y + ".xlsx");
async function rounds(sheet = "สำรวจ") {
  const [a, b] = await Promise.all([
    openExcel(file(2568)),
    openExcel(file(2569)),
  ]);
  return [readSheet(a, "สำรวจ"), readSheet(b, sheet)];
}
test("อ่าน Excel จริง จับคู่ตามจุดและตรวจแต่ละค่าวัด", async () => {
  const [a, b] = await rounds(),
    points = pairRounds(a, b);
  assert.deepEqual(
    points,
    pairRounds(
      { ...a, rows: [...a.rows].reverse() },
      { ...b, rows: [...b.rows].reverse() },
    ),
  );
  const results = points.map((p) => compareTree(p));
  assert.match(results[0].lines[0], /ความสูงลดลง 0.5 เมตร/);
  assert.equal(results[1].lines.length, 2);
  assert.match(results[1].status, /รอตรวจสอบ/);
  assert.deepEqual(results[2].lines, [
    "RCDลดลง 0.25 เซนติเมตร (5 → 4.75 เซนติเมตร)",
  ]);
  assert.match(results[3].lines[0], /ไม่พอ/);
  assert.match(results[3].lines[1], /RCDลดลง/);
  assert.equal(results[4].status, "พบเฉพาะรอบก่อน");
  assert.equal(results[5].status, "พบเฉพาะรอบปัจจุบัน");
  assert.equal(displayDate(points[1].after[0].date), "25-07-2569");
  assert.equal(displayDate(points[0].after[0].date), "25-07-2569");
  assert.equal(displayDate(a.rows[0].date), "ไม่ระบุวันที่");
  assert.equal(points[1].after[0].line, 5);
  assert.equal(
    compareTree(points[1], { decision: "same" }).lines.filter((l) =>
      l.includes("ลดลง"),
    ).length,
    2,
  );
  points[3].after[0].height = "วัดไม่ได้";
  assert.equal(compareTree(points[3]).lines.length, 2);
});
test("พักจุดซ้ำพร้อมทุกแถวและพักแถวคีย์ว่าง", async () => {
  const points = pairRounds(...(await rounds("ทดสอบคีย์ซ้ำ")));
  const duplicate = points.find((p) => p.no === "1");
  assert.deepEqual(
    duplicate.after.map((r) => r.line),
    [3, 7],
  );
  assert.match(compareTree(duplicate).status, /ซ้ำ/);
  assert.deepEqual(compareTree(duplicate, { decision: "same" }).lines, []);
  const empty = points.find((p) => !p.no);
  assert.equal(empty.after[0].line, 8);
  assert.match(compareTree(empty).status, /ว่าง/);
});
test("ปฏิเสธแปลงต่างกัน ผิดลำดับปี หลายรอบ และหัวตารางขาด", async () => {
  const [a, b] = await rounds();
  assert.throws(() => pairRounds(a, { ...b, plot: "P02" }), /คนละแปลง/);
  assert.throws(() => pairRounds(b, a), /ปี/);
  const source = await openExcel(file(2569)),
    s = source.workbook.getWorksheet("สำรวจ");
  s.getCell("E3").value = 2568;
  assert.throws(() => readSheet(source, s.name), /หลายรอบ/);
  s.getCell("A1").value = "";
  assert.throws(() => readSheet(source, s.name), /ขาดคอลัมน์/);
});
test("เปิดคลังใหม่หลังบันทึก อ่านหลักฐานและประวัติการเปลี่ยนผลตรวจ", async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "tidtime-"));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const original = fs.readFileSync(file(2569));
  const store = new Store(dir),
    [a, b] = await rounds(),
    job = store.create(a, b);
  const p = pairRounds(a, b)[1],
    note = "เจ้าหน้าที่ประจำแปลงยืนยันทางโทรศัพท์";
  store.review(job.id, p.key, {
    decision: "same",
    checked: "2026-07-27",
    note: "",
  });
  store.review(job.id, p.key, {
    decision: "replacement",
    checked: "2026-07-28",
    planted: "2026-06-01",
    note,
  });
  const reopened = new Store(dir);
  assert.equal(reopened.list()[0].id, job.id);
  const restored = reopened.load(job.id),
    [first, last] = restored.history;
  assert.equal(first.decision, "same");
  assert.equal(last.key, p.key);
  assert.equal(last.note, note);
  assert.equal(last.checked, "2026-07-28");
  assert.ok(Number.isFinite(Date.parse(last.saved)));
  assert.equal(last.planted, "2026-06-01");
  const baseline = restored.after.rows.find(
    (r) => r.line === last.baselineLine,
  );
  assert.equal(baseline.height, 0.7);
  assert.equal(
    restored.before.rows.find((r) => r.line === last.previousLine).id,
    "T002",
  );
  const result = compareTree(p, last);
  assert.match(result.lines.join(" "), /ค่าแรกของต้นใหม่.*0.7/);
  assert.doesNotMatch(result.lines.join(" "), /ลดลง/);
  assert.deepEqual(fs.readFileSync(file(2569)), original);
});

test("เรียงตามวันสำรวจจริง และพักเมื่อวันในปีเดียวกันกำกวม", async () => {
  const [a, b] = await rounds();
  const dated = (round, date) => ({
    ...round,
    year: "2569",
    rows: round.rows.map((row) => ({ ...row, date })),
  });
  const january = dated(a, "15/01/2569"),
    june = dated(b, "2026-06-15");
  assert.match(compareTree(pairRounds(january, june)[0]).lines[0], /0.5 เมตร/);
  assert.throws(() => pairRounds(june, january), /วันที่สำรวจ/);
  for (const bad of ["", "31-02-2569", "2026-01-15"]) {
    assert.throws(() => pairRounds(january, dated(b, bad)), /ปีเดียวกัน/);
  }
  june.rows[1].date = "2026-06-16";
  assert.throws(() => pairRounds(january, june), /ปีเดียวกัน/);
  june.rows[1].date = "";
  assert.throws(() => pairRounds(january, june), /ปีเดียวกัน/);
});

test("เปิด JSON โครงสร้าง 0.1.1 และเชื่อมเหตุการณ์ทดแทนด้วยคีย์เดิม", async (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "tidtime-legacy-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const [before, after] = await rounds();
  const id = "10000000-0000-4000-8000-000000000001";
  const event = {
    key: '["P01","2"]',
    decision: "replacement",
    note: "ทดสอบ",
    checked: "2026-08-01",
    saved: "2026-08-02T02:00:00Z",
    planted: null,
    previousSheet: "สำรวจ",
    previousLine: 3,
    baselineSheet: "สำรวจ",
    baselineLine: 5,
  };
  fs.writeFileSync(
    path.join(directory, id + ".json"),
    JSON.stringify({
      id,
      created: event.saved,
      before,
      after,
      history: [event],
    }),
  );
  const store = new Store(directory);
  const restored = store.load(id);
  const point = require("../compare")
    .matchPoints(restored.before, restored.after)
    .find((p) => p.key === restored.history[0].key);
  assert.equal(point.before[0].line, event.previousLine);
  assert.equal(point.after[0].line, event.baselineLine);
  assert.equal(point.before[0].id, "T002");
  assert.match(
    compareTree(point, restored.history[0]).lines.join(" "),
    /ค่าแรกของต้นใหม่.*0.7/,
  );
  const revised = store.review(id, event.key, {
    decision: "same",
    note: "ทบทวน",
    checked: "2026-08-03",
  });
  assert.deepEqual(revised.history[0], event);
});
