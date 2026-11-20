const Decimal = require("decimal.js");

const measurements = [
  ["height", "ความสูง", "เมตร"],
  ["rcd", "RCD", "เซนติเมตร"],
];
const clean = (value) => String(value ?? "").trim();

function pairRounds(before, after) {
  if (before.form && after.form && before.form !== after.form)
    throw Error("กรุณาเลือกชีตแบบเดียวกันทั้งสองรอบ");
  if (before.plot !== after.plot) throw Error("สองชีตมาจากคนละแปลง กรุณาเลือกใหม่");
  if (+before.year >= +after.year)
    throw Error("ปีรอบปัจจุบันต้องอยู่หลังปีรอบก่อน กรุณาตรวจไฟล์และชีต");
  const points = new Map();
  for (const [side, round] of [
    ["before", before],
    ["after", after],
  ]) {
    for (const row of round.rows) {
      const plot = clean(row.plot),
        no = clean(row.no);
      const key =
        plot && no
          ? JSON.stringify([plot, no])
          : JSON.stringify([side, row.sourceSheet || "", row.line, null]);
      if (!points.has(key))
        points.set(key, { key, plot, no, before: [], after: [] });
      points.get(key)[side].push(row);
    }
  }
  return [...points.values()].sort(
    (a, b) =>
      a.no.localeCompare(b.no, "th", { numeric: true }) ||
      a.key.localeCompare(b.key),
  );
}

function decimal(raw) {
  const value = clean(raw);
  return /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(value)
    ? new Decimal(value)
    : null;
}

function compareTree(point, review) {
  const lines = [];
  function result(status, needsReview = true) {
    return { status, needsReview, lines };
  }
  if (!point.plot || !point.no) return result("คีย์จุดปลูกว่าง · พักการเปรียบเทียบ");
  if (point.before.length > 1 || point.after.length > 1)
    return result("คีย์จุดปลูกซ้ำ · พักการเปรียบเทียบ");
  const [a] = point.before,
    [b] = point.after;
  const metrics =
    b?.kind === "bamboo"
      ? [
          measurements[0],
          ...[1, 2, 3].map((i) => ["culm" + i, "RCD ไผ่ ลำ " + i, "เซนติเมตร"]),
        ]
      : b && "fruitGirth" in b
        ? [
            measurements[0],
            [
              "fruitGirth",
              b.kind === "banana" ? "ความโตกล้วย" : "ความโตไม้ผล",
              "เซนติเมตร",
            ],
          ]
        : measurements;
  if (!a || !b) return result(a ? "พบเฉพาะรอบก่อน" : "พบเฉพาะรอบปัจจุบัน", !review);
  let incomplete = false;
  for (const [field, label] of [
    ["id", "รหัสต้นไม้"],
    ["species", "ชนิดพันธุ์"],
  ]) {
    const old = clean(a[field]),
      next = clean(b[field]);
    if (!old || !next) incomplete = true;
    if (old !== next)
      lines.push(
        label + " จาก " + (old || "ไม่ระบุ") + " เป็น " + (next || "ไม่ระบุ"),
      );
  }
  if (review?.decision === "replacement") {
    lines.push(
      "ค่าแรกของต้นใหม่ · " +
        metrics
          .map(([f, l, u]) => l + " " + (clean(b[f]) || "ไม่ระบุ") + " " + u)
          .join(" · "),
    );
    return result("ยืนยันว่าปลูกทดแทน", false);
  }
  if ((incomplete || lines.length) && review?.decision !== "same") {
    if (incomplete) lines.push("ข้อมูลรหัสหรือชนิดพันธุ์ไม่ครบ");
    return result("รอตรวจสอบว่าเป็นต้นเดิมหรือปลูกทดแทน");
  }
  let warnings = 0;
  for (const [field, label, unit] of metrics) {
    const old = decimal(a[field]),
      next = decimal(b[field]);
    if (!old || !next) {
      lines.push(label + " มีข้อมูลไม่พอ");
    } else if (next.lt(old)) {
      lines.push(
        label +
          "ลดลง " +
          old.minus(next).toFixed() +
          " " +
          unit +
          " (" +
          a[field] +
          " → " +
          b[field] +
          " " +
          unit +
          ")",
      );
    } else continue;
    warnings++;
  }
  return result(
    warnings ? "พบข้อสังเกตจากการวัด" : "ไม่พบค่าวัดลดลง",
    warnings > 0 && !review,
  );
}

module.exports = { pairRounds, compareTree };
