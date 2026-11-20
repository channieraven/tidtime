const api = window.tidtime;
const $ = (id) => document.getElementById(id);
const titles = { before: "รอบก่อน", after: "รอบปัจจุบัน" };
const decisions = {
  pending: "รอตรวจเพิ่มเติม",
  checked: "ตรวจรับทราบ",
  same: "ยืนยันว่าเป็นต้นเดิม",
  replacement: "ยืนยันว่าปลูกทดแทน",
};
const previews = {};
const years = {};
function updateCreate() {
  $("create").disabled = !Object.keys(titles).every(
    (side) =>
      previews[side] && (!previews[side].form || /^\d{4}$/.test(years[side])),
  );
}
let job, selected;
function el(tag, text, parent) {
  const node = document.createElement(tag);
  if (text != null) node.textContent = text;
  if (parent) parent.append(node);
  return node;
}
function option(select, value, text) {
  const node = el("option", text, select);
  node.value = value;
}
function panel(id) {
  for (const p of ["importer", "library", "work"]) $(p).hidden = p !== id;
}
function message(text) {
  $("message").textContent = text;
}
let busy = false;
async function run(action) {
  if (busy) return;
  busy = true;
  document.querySelector("main").inert = true;
  message("");
  try {
    await action();
  } catch (error) {
    message(error.message.replace(/^.*?Error: /, ""));
  } finally {
    busy = false;
    document.querySelector("main").inert = false;
  }
}
function previewText(p) {
  return (
    p.file +
    " · ชีต " +
    p.sheet +
    "\nแปลง " +
    p.plot +
    " · ปี " +
    (p.year || "กรุณาระบุ") +
    " · " +
    p.count +
    " แถว\nวันที่สำรวจ " +
    p.dates
  );
}
for (const side of Object.keys(titles)) {
  const card = el("div", null, $("slots"));
  card.className = "card";
  el("h2", titles[side], card);
  const choose = el("button", "เลือกไฟล์ .xlsx", card);
  const filename = el("p", "ยังไม่ได้เลือกไฟล์", card);
  const sheets = el("select", null, card);
  sheets.hidden = true;
  sheets.ariaLabel = "ชีต" + titles[side];
  const meta = el("p", "", card);
  const yearLabel = el("label", "ปีรอบสำรวจ (พ.ศ.) ตามทะเบียนรอบ ", card);
  const year = el("input", null, yearLabel);
  year.placeholder = "พ.ศ. 4 หลัก";
  year.maxLength = 4;
  year.inputMode = "numeric";
  yearLabel.hidden = true;
  year.oninput = () => {
    years[side] = year.value.trim();
    updateCreate();
  };
  function reset() {
    yearLabel.hidden = true;
    year.value = "";
    delete years[side];
    delete previews[side];
    $("create").disabled = true;
    meta.textContent = "";
  }
  async function read() {
    reset();
    try {
      previews[side] = await api.sheet(side, sheets.value);
      meta.textContent = previewText(previews[side]);
      yearLabel.hidden = !previews[side].form;
      updateCreate();
    } catch (e) {
      meta.textContent = e.message.replace(/^.*?Error: /, "");
    }
  }
  choose.onclick = () =>
    run(async () => {
      let file;
      try {
        file = await api.choose(side);
      } catch (e) {
        reset();
        filename.textContent = "เปิดไฟล์ไม่สำเร็จ";
        sheets.hidden = true;
        throw e;
      }
      if (!file) return;
      filename.textContent = file.file;
      sheets.replaceChildren();
      sheets.hidden = false;
      for (const s of file.sheets) option(sheets, s, s);
      await read();
    });
  sheets.onchange = () => run(read);
}
function openJob(next) {
  job = next;
  selected = null;
  $("filter").checked = false;
  panel("work");
  $("heading").textContent =
    "แปลง " +
    job.before.plot +
    " · " +
    job.before.year +
    " → " +
    job.after.year;
  $("rounds").textContent = Object.entries(titles)
    .map(([s, t]) => t + " " + previewText(job[s]).replaceAll("\n", " · "))
    .join("  |  ");
  points();
}
$("newJob").onclick = () => run(() => panel("importer"));
$("archive").onclick = () =>
  run(async () => {
    const jobs = await api.list();
    panel("library");
    $("jobs").replaceChildren();
    if (!jobs.length) el("p", "ยังไม่มีงานที่บันทึก", $("jobs"));
    for (const j of jobs)
      el(
        "button",
        "แปลง " +
          j.plot +
          " · " +
          j.from +
          " → " +
          j.to +
          " · " +
          localTime(j.created),
        $("jobs"),
      ).onclick = () => run(async () => openJob(await api.load(j.id)));
  });
$("create").onclick = () => run(async () => openJob(await api.create(years)));
$("filter").onchange = points;
function points() {
  const list = job.points.filter((p) => !$("filter").checked || p.needsReview);
  $("count").textContent =
    list.length + " จาก " + job.points.length + " รายการ";
  $("points").replaceChildren();
  if (!list.some((p) => p.key === selected)) selected = list[0]?.key;
  for (const p of list) {
    const btn = el(
      "button",
      (p.plot || "ขาดแปลง") + " / " + (p.no || "ขาดลำดับ") + " · " + p.status,
      $("points"),
    );
    btn.classList.toggle("active", p.key === selected);
    btn.onclick = () => {
      selected = p.key;
      points();
    };
  }
  detail(list.find((p) => p.key === selected));
}
function localTime(iso) {
  return new Date(iso).toLocaleString("th-TH");
}
function evidence(parent, p, side) {
  const block = el("div", null, parent);
  block.className = "evidence";
  const r = job[side];
  el("h3", titles[side] + " · " + r.year, block);
  el("p", r.file + " · ชีต " + r.sheet, block);
  if (!p[side].length) el("p", "ไม่มีแถวจากรอบนี้", block);
  for (const [index, row] of p[side].entries()) {
    el("strong", "แถว Excel " + row.line, block);
    if (row.sourceSheet) el("p", "ชีตต้นทาง " + row.sourceSheet, block);
    const dl = el("dl", null, block);
    for (const [key, label] of Object.entries(job.columns)) {
      if (key === "rcd" && ("fruitGirth" in row || row.kind === "bamboo"))
        continue;
      el(
        "dt",
        key === "note"
          ? "หมายเหตุใน Excel"
          : key === "year" && r.form
            ? "ปีรอบที่ระบุเมื่อนำเข้า"
            : label,
        dl,
      );
      el(
        "dd",
        key === "date"
          ? p.displayDates[side][index]
          : String(row[key] ?? "") || "—",
        dl,
      );
    }
    for (const [label, value] of Object.entries(row.extras || {})) {
      el("dt", label, dl);
      el("dd", String(value), dl);
    }
    if ("fruitGirth" in row) {
      el(
        "dt",
        row.kind === "banana" ? "ความโตกล้วย (ซม.)" : "ความโตไม้ผลที่ใช้เทียบ (ซม.)",
        dl,
      );
      el("dd", String(row.fruitGirth) || "—", dl);
    }
    if (row.kind === "bamboo")
      for (let i = 1; i <= 3; i++) {
        el("dt", "RCD ไผ่ ลำ " + i + " (ซม.)", dl);
        el("dd", String(row["culm" + i]) || "—", dl);
      }
  }
}
function detail(p) {
  const box = $("detail");
  box.replaceChildren();
  if (!p) {
    el("p", "ไม่มีรายการที่ต้องตรวจ", box);
    return;
  }
  el("h2", "จุดปลูก " + (p.plot || "—") + " / " + (p.no || "—"), box);
  el("p", p.status, box).className = "status";
  for (const l of p.lines) el("p", l, box);
  const pair = el("div", null, box);
  pair.className = "pair";
  for (const side of Object.keys(titles)) evidence(pair, p, side);
  el("h3", "ผลตรวจของเจ้าหน้าที่", box);
  const form = el("form", null, box);
  function field(label, tag, type) {
    const wrap = el("label", label, form);
    const input = el(tag, null, wrap);
    if (type) input.type = type;
    return input;
  }
  const choice = field("ผลสรุป", "select");
  const paired =
    p.plot && p.no && p.before.length === 1 && p.after.length === 1;
  const latest = p.history.at(-1);
  for (const [v, label] of Object.entries(decisions)) {
    if (!paired && ["same", "replacement"].includes(v)) continue;
    if (v === "checked" && p.status === "รอตรวจสอบว่าเป็นต้นเดิมหรือปลูกทดแทน")
      continue;
    option(choice, v, label);
  }
  choice.value =
    latest?.decision || (p.status.startsWith("รอตรวจ") ? "pending" : "checked");
  const checked = field("วันที่ตรวจสอบ ", "input", "date");
  checked.required = true;
  const today = new Date();
  today.setMinutes(today.getMinutes() - today.getTimezoneOffset());
  checked.value = latest?.checked || today.toISOString().slice(0, 10);
  const planted = field("วันที่ปลูกจริง (ถ้าทราบ) ", "input", "date");
  planted.value = latest?.planted || "";
  choice.onchange = () => {
    planted.parentElement.hidden = choice.value !== "replacement";
  };
  choice.onchange();
  const note = field("หมายเหตุเจ้าหน้าที่", "textarea");
  note.value = latest?.note || "";
  el("button", "บันทึกผลตรวจ", form).className = "primary";
  form.onsubmit = (event) => {
    event.preventDefault();
    run(async () => {
      job = await api.review(job.id, p.key, {
        decision: choice.value,
        note: note.value,
        checked: checked.value,
        planted: choice.value === "replacement" ? planted.value : "",
      });
      points();
      message("บันทึกผลตรวจแล้ว");
    });
  };
  el("h3", "ประวัติการตรวจ (" + p.history.length + ")", box);
  for (const h of [...p.history].reverse()) {
    el(
      "div",
      decisions[h.decision] +
        " · วันที่ตรวจ " +
        h.checked +
        "\nระบบบันทึก " +
        localTime(h.saved) +
        (h.decision === "replacement"
          ? "\nวันที่ปลูกจริง " +
            (h.planted || "ไม่ระบุ") +
            " · ฐานต้นใหม่จากแถว " +
            h.baselineLine
          : "") +
        "\n" +
        h.note,
      box,
    ).className = "event";
  }
}
