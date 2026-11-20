// Snapshot เป็นหลายตารางในสมุดงานเดียว เชื่อมด้วย UUID ก่อนแบ่งชุดสำหรับเปรียบเทียบ
const groups = {
  growth_forest: "forest",
  growth_rubber: "forest",
  growth_bamboo: "forest",
  growth_fruit: "fruit",
  growth_banana: "fruit",
};

function table(sheet, required, read) {
  if (!sheet) throw Error("Snapshot ขาดชีตที่จำเป็น");
  const indices = new Map();
  sheet
    .getRow(3)
    .eachCell((cell, index) => indices.set(cell.text.trim(), index));
  for (const name of required)
    if (!indices.has(name)) throw Error(sheet.name + " ขาดคอลัมน์ " + name);
  const rows = [];
  sheet.eachRow((row, line) => {
    if (line <= 3) return;
    const record = { line };
    for (const [key, index] of indices) record[key] = read(row.getCell(index));
    rows.push(record);
  });
  return rows;
}
function indexBy(rows, key, sheet) {
  const index = new Map();
  for (const row of rows) {
    if (!row[key] || index.has(row[key]))
      throw Error(sheet + " มีคีย์ว่างหรือซ้ำ แถว " + row.line);
    index.set(row[key], row);
  }
  return index;
}
function snapshotRounds(workbook, file, read) {
  if (
    !workbook.getWorksheet("trees") ||
    !Object.keys(groups).some((s) => workbook.getWorksheet(s))
  )
    return null;
  const trees = indexBy(
    table(
      workbook.getWorksheet("trees"),
      ["tree_id", "plot_id", "tree_code", "tree_subtype_id"],
      read,
    ),
    "tree_id",
    "trees",
  );
  const species = indexBy(
    table(
      workbook.getWorksheet("tree_subtypes"),
      ["tree_subtype_id", "name_th"],
      read,
    ),
    "tree_subtype_id",
    "tree_subtypes",
  );
  const rounds = new Map();
  for (const [sheetName, form] of Object.entries(groups)) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) continue;
    const bamboo = sheetName === "growth_bamboo",
      banana = sheetName === "growth_banana";
    const height = bamboo
      ? "max_culm_height_m"
      : banana
        ? "max_pseudostem_height_m"
        : "height_m";
    const keys = bamboo ? ["rcd_cm_1", "rcd_cm_2", "rcd_cm_3"] : ["rcd_cm"];
    for (const record of table(
      sheet,
      ["tree_id", "survey_date", "notes", height, ...keys],
      read,
    )) {
      const tree = trees.get(record.tree_id);
      if (!tree || !tree.plot_id)
        throw Error(
          sheetName + " แถว " + record.line + " เชื่อมแปลงใน trees ไม่ได้",
        );
      const subtype = species.get(tree.tree_subtype_id);
      const plot = String(tree.plot_id).trim(),
        date = record.survey_date;
      const name = [
        plot,
        form === "forest" ? "ไม้ป่า+ยางพารา+ไผ่" : "ผลไม้+กล้วย",
        date || "ไม่ระบุวันที่",
      ].join(" · ");
      if (!rounds.has(name))
        rounds.set(name, { file, sheet: name, plot, year: "", form, rows: [] });
      // ตรวจรูปแบบเต็มและแปลงก่อนใช้เลขท้าย สามหลักนี้ตรงกับ "ต้นที่" ในแบบฟอร์ม P01
      const code = String(tree.tree_code).trim();
      const match = /^(P\d+)[A-Z]\d{2}(\d{3})$/.exec(code);
      const no = match && match[1] === plot ? Number(match[2]) : "";
      const extras = {
        ทะเบียนต้นไม้: "trees แถว " + tree.line,
        ชนิดพันธุ์ต้นทาง: subtype
          ? "tree_subtypes แถว " + subtype.line
          : "ไม่พบในทะเบียน",
        ที่มาลำดับต้น: "เลขท้าย 3 หลักของ tree_code",
        สถานะรอดจากต้นทาง: record.survival_status,
      };
      const row = {
        plot,
        no,
        id: code,
        species: subtype?.name_th || "",
        date,
        year: "",
        line: record.line,
        sourceSheet: sheetName,
        height: record[height],
        rcd: record.rcd_cm ?? "",
        note: record.notes,
        extras,
      };
      if (form === "fruit") row.fruitGirth = row.rcd;
      if (banana) row.kind = "banana";
      if (bamboo) {
        row.kind = "bamboo";
        for (let i = 1; i <= 3; i++) row["culm" + i] = record["rcd_cm_" + i];
      }
      rounds.get(name).rows.push(row);
    }
  }
  if (!rounds.size) throw Error("Snapshot ไม่มีบันทึกการวัดรายต้น");
  const combined = new Map();
  for (const round of rounds.values()) {
    const name = [
      round.plot,
      "รวมกลุ่มไม้ป่าและไม้ผล",
      round.rows[0].date || "ไม่ระบุวันที่",
    ].join(" · ");
    if (!combined.has(name))
      combined.set(name, { ...round, form: "combined", sheet: name, rows: [] });
    combined.get(name).rows.push(...round.rows);
  }
  return new Map([...combined, ...rounds]);
}
module.exports = { snapshotRounds };
