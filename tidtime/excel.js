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
  const filename = path.basename(file);
  const rounds = require("./mris").snapshotRounds(workbook, filename, valueOf);
  const source = { file: filename, workbook, rounds };
  const names = ["ไม้ป่า+ยางพารา+ไผ่", "ผลผลิตของผลไม้+กล้วย"];
  if (!rounds && names.every((n) => workbook.getWorksheet(n)))
    source.fieldSheets = names;
  return source;
}
const combinedName = "ทั้งแปลง · รวมชีตไม้ป่าและไม้ผล";
function sheetOptions(source) {
  if (source.rounds)
    return [...source.rounds.keys()].sort(
      (a, b) =>
        a.split(" · ")[0].localeCompare(b.split(" · ")[0]) ||
        Number(source.rounds.get(b).form === "combined") -
          Number(source.rounds.get(a).form === "combined") ||
        a.localeCompare(b),
    );
  return [
    ...(source.fieldSheets ? [combinedName] : []),
    ...source.workbook.worksheets.map((s) => s.name),
  ];
}
// แบบฟอร์มภาคสนามใช้ข้อมูลประจำแปลงร่วมกันทั้งชีต และมีหัวตารางสามชั้น
function readFieldForm(source, sheet) {
  const title = (address) =>
    String(valueOf(sheet.getCell(address))).replace(/\s/g, "");
  const fruit = title("F7") === "ไม้ผล";
  const required = {
    A3: "วันที่",
    A4: "ชื่อแปลง",
    A7: "ต้นที่",
    B7: "เลขแท็กต้นไม้",
    C7: "เลขรหัสต้นไม้",
    ...(fruit
      ? {
          D7: "ชนิดพันธุ์ไม้หรือชนิดกล้วย",
          F8: "ความโต",
          G8: "ความสูง",
          Q7: "หมายเหตุ",
          I7: "กล้วย",
        }
      : {
          D7: "ชนิดพันธุ์",
          F7: "ไม้ป่า/ยางพารา",
          F8: "ความโตที่ระดับคอราก",
          K7: "ความสูง(ม.)",
          M7: "หมายเหตุ",
        }),
  };
  for (const [addr, expected] of Object.entries(required)) {
    if (title(addr) !== expected)
      throw Error(
        "รูปแบบหัวตารางต่างจากแบบฟอร์มที่รองรับที่ " +
          addr +
          " (คาดว่าเป็น " +
          expected +
          ")",
      );
  }
  const plot = String(valueOf(sheet.getCell("B4"))).trim();
  if (!plot) throw Error("ขาดชื่อแปลงที่ B4");
  const date = valueOf(sheet.getCell("B3"));
  const fields = fruit
    ? {
        F: "ความโตไม้ผล (ซม.)",
        H: "การติดดอกออกผล",
        I: "กล้วยทั้งหมด (ต้น)",
        J: "หน่อ",
        K: "ต้น 1 ปี",
        L: "ความโตกล้วย (ซม.)",
        M: "ความสูงต้นที่สูงที่สุดในกอ (ม.)",
        N: "ผลผลิต (เครือ)",
        O: "ผลผลิต (หวี)",
        P: "ราคา/หวี",
        R: "รุ่นปลูก",
      }
    : {
        G: "จำนวนไผ่ (ลำ)",
        H: "ความโตไผ่ ลำ 1 (ซม.)",
        I: "ความโตไผ่ ลำ 2 (ซม.)",
        J: "ความโตไผ่ ลำ 3 (ซม.)",
        L: "การติดดอกออกผล",
        N: "รุ่นปลูก",
      };
  const rows = [];
  sheet.eachRow((line, number) => {
    if (number < 10) return;
    const read = (col) => valueOf(line.getCell(col));
    const extras = {};
    for (const [col, label] of Object.entries({
      B: "เลขแท็กต้นไม้",
      E: "ระยะปลูก",
      ...fields,
    })) {
      if (read(col) !== "") extras[label] = read(col);
    }
    const row = {
      line: number,
      sourceSheet: sheet.name,
      plot,
      year: "",
      date,
      no: read("A"),
      id: read("C"),
      species: read("D"),
      height: read(fruit ? "G" : "K"),
      rcd: fruit ? "" : read("F"),
      note: read(fruit ? "Q" : "M"),
      extras,
    };
    if (
      [
        row.no,
        row.id,
        row.species,
        row.height,
        row.rcd,
        row.note,
        ...Object.values(extras),
      ].every((v) => v === "")
    )
      return;
    if (
      !row.no &&
      !row.id &&
      !row.species &&
      Object.keys(extras).length === 1 &&
      /^(ไม้ป่า|ไม้ผล)\s+\d+$/.test(String(read("B")).trim())
    )
      return;
    if (fruit) row.fruitGirth = read("F");
    if (
      fruit &&
      (String(row.species).includes("กล้วย") ||
        ["I", "J", "K", "L", "M"].some((col) => read(col) !== ""))
    ) {
      row.kind = "banana";
      row.height = read("M");
      row.fruitGirth = read("L");
      if (read("G") !== "") extras["ความสูงในช่องไม้ผล (ม.)"] = read("G");
    }
    if (
      !fruit &&
      (String(row.species).includes("ไผ่") ||
        ["G", "H", "I", "J"].some((col) => read(col) !== ""))
    ) {
      row.kind = "bamboo";
      for (const [i, col] of ["H", "I", "J"].entries())
        row["culm" + (i + 1)] = read(col);
    }
    rows.push(row);
  });
  if (!rows.length) throw Error("ชีตไม่มีข้อมูลต้นไม้");
  return {
    file: source.file,
    sheet: sheet.name,
    plot,
    year: "",
    form: fruit ? "fruit" : "forest",
    rows,
  };
}
function setSurveyYear(round, year) {
  if (!round.form) return round;
  if (!/^\d{4}$/.test(year || ""))
    throw Error("กรุณาระบุปีรอบสำรวจ พ.ศ. 4 หลักของแต่ละไฟล์");
  return { ...round, year, rows: round.rows.map((row) => ({ ...row, year })) };
}
function readSheet(source, name) {
  if (source.fieldSheets && name === combinedName) {
    const [a, b] = source.fieldSheets.map((n) => readSheet(source, n));
    if (a.plot !== b.plot || a.rows[0].date !== b.rows[0].date)
      throw Error("แปลงหรือวันที่ของสองชีตต่างกัน กรุณาเลือกทีละชีต");
    return {
      ...a,
      sheet: name,
      form: "combined",
      rows: [...a.rows, ...b.rows],
    };
  }
  if (source.rounds) {
    const round = source.rounds.get(name);
    if (!round) throw Error("กรุณาเลือกชุดข้อมูลแปลงจาก Snapshot");
    return round;
  }
  const sheet = source.workbook.getWorksheet(name);
  if (!sheet) throw Error("ไม่พบชีตที่เลือก");
  if (String(valueOf(sheet.getCell("A7"))).trim() === "ต้นที่")
    return readFieldForm(source, sheet);
  if (sheet.name.trim() === "สรุป" && valueOf(sheet.getCell("A2")) === "แปลง")
    throw Error("ชีตสรุปไม่มีข้อมูลรายต้น กรุณาเลือกชีตสำรวจ");
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
  sheetOptions,
  readSheet,
  setSurveyYear,
  columns,
  displayDate,
};
