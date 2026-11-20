// รันทดสอบ Electron โดยแทนเฉพาะกล่องเลือกไฟล์ของ Windows
const { app, dialog } = require("electron");
const path = require("node:path");
const assert = require("node:assert/strict");
const [phase, directory] = process.argv.slice(-2);
app.setPath("userData", directory);
const fieldTest = Boolean(
  process.env.TIDTIME_TEST_SNAPSHOT && process.env.TIDTIME_TEST_FIELD,
);
const samples = fieldTest
  ? [process.env.TIDTIME_TEST_SNAPSHOT, process.env.TIDTIME_TEST_FIELD]
  : [2568, 2569].map((year) =>
      path.join(directory, "P01-" + year + ".xlsx"),
    );
dialog.showOpenDialog = async () => ({
  canceled: false,
  filePaths: [samples.shift()],
});
app.on("browser-window-created", (_, win) => {
  win.hide();
  win.webContents.once("did-finish-load", async () => {
    try {
      const result = await win.webContents.executeJavaScript(
        "(" +
          exercise.toString() +
          ")(" +
          JSON.stringify(phase) +
          "," +
          fieldTest +
          ")",
      );
      assert.match(result, /เจ้าหน้าที่ประจำแปลงยืนยันทางโทรศัพท์/);
      assert.match(result, fieldTest ? /ค่าแรกของต้นใหม่/ : /ค่าแรกของต้นใหม่.*0.7/);
      if (fieldTest) {
        assert.match(result, /growth_forest/);
        assert.match(result, /ชีตต้นทาง/);
      }
      assert.doesNotMatch(result, /ลดลง 1.1/);
      console.log("Electron UI passed");
      app.exit(0);
    } catch (error) {
      console.error(error);
      app.exit(1);
    }
  });
});
async function exercise(phase, fieldTest) {
  const $ = (selector) => document.querySelector(selector);
  async function wait(check) {
    const end = Date.now() + 15000;
    while (!check()) {
      if (Date.now() > end)
        throw Error("หมดเวลารอหน้าจอ");
      await new Promise((resolve) => setTimeout(resolve, 40));
    }
  }
  const wanted = () => fieldTest ? localStorage.getItem("testPoint") : "/ 2 ·";
  if (phase === "write") {
    const buttons = document.querySelectorAll("#slots button");
    buttons[0].click();
    await wait(() => !$("main").inert);
    buttons[1].click();
    if (fieldTest) {
      await wait(() => !$("main").inert);
      const inputs = document.querySelectorAll("#slots input");
      for (const [i, year] of ["2568", "2569"].entries()) {
        inputs[i].value = year;
        inputs[i].dispatchEvent(new Event("input"));
      }
    }
    await wait(() => !$("#create").disabled && !$("main").inert);
    $("#create").click();
    await wait(() => !$("#work").hidden);
    [...document.querySelectorAll("#points button")]
      .find((n) => n.textContent.includes(fieldTest ? "รอตรวจสอบว่าเป็นต้นเดิม" : "/ 2 ·"))
      .click();
    if (fieldTest) localStorage.setItem("testPoint", $("#points .active").textContent.split(" · ")[0] + " ·");
    const choice = $("form select");
    choice.value = "replacement";
    choice.dispatchEvent(new Event("change"));
    $("form input[type=date]").value = "2026-07-27";
    $("form textarea").value = "เจ้าหน้าที่ประจำแปลงยืนยันทางโทรศัพท์";
    $("form button").click();
    await wait(() => $("#message").textContent.includes("บันทึกผลตรวจแล้ว"));
  }
  $("#archive").click();
  await wait(() => $("#jobs button"));
  $("#jobs button").click();
  await wait(() => !$("#work").hidden);
  [...document.querySelectorAll("#points button")]
    .find((n) => n.textContent.includes(wanted()))
    .click();
  return $("#detail").innerText;
}
require("../main");
