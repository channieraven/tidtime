const previous = { plotId: "P01", treeNo: 2, treeId: "A02", species: "มะขามป้อม", heightM: 5.0, rcdCm: 4.5, surveyYear: 2568, surveyDate: null };
const current = { plotId: "P01", treeNo: 2, treeId: "A02", species: "มะขามป้อม", heightM: 3.2, rcdCm: 4.0, surveyYear: 2569, surveyDate: "25-07-2569" };

function reportDecrease(label, before, after, unit) {
  if (!Number.isFinite(before) || !Number.isFinite(after)) {
    console.log(`เปรียบเทียบ${label}ไม่ได้ เพราะข้อมูลการวัดไม่ครบหรือไม่ใช่ตัวเลข`);
    return;
  }
  if (after < before) {
    console.log(`${label}ลดลง ${(before - after).toFixed(1)} ${unit} (จาก ${before} ${unit} เป็น ${after} ${unit})`);
  }
}

function compareTree(previous, current) {
  if (!previous?.plotId || !current?.plotId || previous.treeNo == null || current.treeNo == null) {
    console.log("เปรียบเทียบไม่ได้ เพราะข้อมูลจุดปลูกไม่ครบ");
    return;
  }
  if (previous.plotId !== current.plotId || previous.treeNo !== current.treeNo) {
    console.log("ไม่สามารถเปรียบเทียบต้นไม้ได้ เนื่องจากข้อมูลมาจากคนละจุดปลูก");
    return;
  }
  console.log(`เปรียบเทียบรอบปี ${previous.surveyYear ?? "ไม่ระบุปี"} (${previous.surveyDate ?? "ไม่ระบุวันที่"}) กับรอบปี ${current.surveyYear ?? "ไม่ระบุปี"} (${current.surveyDate ?? "ไม่ระบุวันที่"})`);
  if (!previous.treeId || !current.treeId || !previous.species || !current.species) {
    console.log("เปรียบเทียบไม่ได้ เพราะข้อมูลรหัสหรือชนิดต้นไม้ไม่ครบ");
    return;
  }
  if (previous.treeId !== current.treeId || previous.species !== current.species) {
    if (previous.treeId !== current.treeId) {
      console.log(`รหัสต้นไม้เปลี่ยนจาก ${previous.treeId} เป็น ${current.treeId}`);
    }
    if (previous.species !== current.species) {
      console.log(`ชนิดต้นไม้เปลี่ยนจาก ${previous.species} เป็น ${current.species}`);
    }
    console.log("โปรดตรวจสอบว่าเป็นต้นเดิมหรือปลูกทดแทน");
    return;
  }
  reportDecrease("ความสูงต้นไม้", previous.heightM, current.heightM, "เมตร");
  reportDecrease("ความโตที่ระดับคอรากต้นไม้", previous.rcdCm, current.rcdCm, "เซนติเมตร");
}

compareTree(previous, current);

const changedTree = { plotId: "P01", treeNo: 3, treeId: "A03", species: "หว้า", heightM: 1, rcdCm: 1 };
compareTree(previous, changedTree);

const previousRows = [
  { plotId: "P01", treeNo: 1, treeId: "A01", species: "มะม่วง", heightM: 4.5, rcdCm: 3.0 },
  previous,
];
const target = { plotId: "P01", treeNo: 1 };
const match = previousRows.find(row => row.plotId === target.plotId && row.treeNo === target.treeNo);
if (match) {
  console.log(match);
} else {
  console.log("ไม่พบข้อมูลต้นไม้ดังกล่าว");
}


