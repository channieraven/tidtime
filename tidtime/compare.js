const { matchRows, compareNumbers, observationDate } = require("./core");

const measurements = [
  ["height", "ความสูง", "เมตร"],
  ["rcd", "RCD", "เซนติเมตร"],
];
const clean = (value) => String(value ?? "").trim();

function pairRounds(before, after) {
  if (before.form && after.form && before.form !== after.form)
    throw Error("กรุณาเลือกชีตแบบเดียวกันทั้งสองรอบ");
  if (before.plot !== after.plot) throw Error("สองชีตมาจากคนละแปลง กรุณาเลือกใหม่");
  const earlier = observationDate(before.rows, surveyDay);
  const later = observationDate(after.rows, surveyDay);
  if (earlier && later && earlier !== later) {
    if (earlier > later) throw Error("วันที่สำรวจรอบปัจจุบันต้องอยู่หลังรอบก่อน");
  } else if (+before.year === +after.year) {
    throw Error(
      "รอบปีเดียวกันต้องมีวันที่สำรวจครบทุกแถว เป็นวันเดียวภายในรอบ และเป็นคนละวันระหว่างรอบ",
    );
  } else if (+before.year > +after.year) {
    throw Error("ปีรอบปัจจุบันต้องอยู่หลังปีรอบก่อน");
  }
  return matchPoints(before, after);
}

function surveyDay(row) {
  const text = clean(row.date);
  const thai = /^(\d{1,2})([-/])(\d{1,2})\2(\d{4})$/.exec(text);
  if (!thai) return text;
  return (
    String(Number(thai[4]) - 543) +
    "-" +
    thai[3].padStart(2, "0") +
    "-" +
    thai[1].padStart(2, "0")
  );
}

// Saved reviews reference the 0.1.1 key, including sheet and line for blank keys.
function matchPoints(before, after) {
  const groups = matchRows(before.rows, after.rows, (row) => {
    const plot = clean(row.plot),
      no = clean(row.no);
    return plot && no ? JSON.stringify([plot, no]) : null;
  });
  return groups
    .map((group) => {
      const side = group.before.length ? "before" : "after";
      const row = group[side][0];
      const key =
        group.key ??
        JSON.stringify([side, row.sourceSheet || "", row.line, null]);
      return { ...group, key, plot: clean(row.plot), no: clean(row.no) };
    })
    .sort(
      (a, b) =>
        a.no.localeCompare(b.no, "th", { numeric: true }) ||
        a.key.localeCompare(b.key),
    );
}

function compareTree(point, review) {
  const lines = [];
  function result(status, needsReview = true) {
    return { status, needsReview, lines };
  }
  if (point.status === "missing-key")
    return result("คีย์จุดปลูกว่าง · พักการเปรียบเทียบ");
  if (point.status === "duplicate") return result("คีย์จุดปลูกซ้ำ · พักการเปรียบเทียบ");
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
  if (point.status !== "paired")
    return result(
      point.status === "before-only" ? "พบเฉพาะรอบก่อน" : "พบเฉพาะรอบปัจจุบัน",
      !review,
    );
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
  const changes = compareNumbers(
    point,
    metrics.map(([field]) => field),
  );
  for (const [index, change] of changes.entries()) {
    const [, label, unit] = metrics[index];
    if (change.delta === null) {
      lines.push(label + " มีข้อมูลไม่พอ");
    } else if (change.delta.startsWith("-")) {
      lines.push(
        label +
          "ลดลง " +
          change.delta.slice(1) +
          " " +
          unit +
          " (" +
          change.before +
          " → " +
          change.after +
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

module.exports = { pairRounds, matchPoints, compareTree };
