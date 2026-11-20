// รันทดสอบ Electron โดยแทนเฉพาะกล่องเลือกไฟล์ของ Windows
const { app, dialog } = require("electron");
const path = require("node:path");
const assert = require("node:assert/strict");
const [phase, directory] = process.argv.slice(-2);
app.setPath("userData", directory);
const samples = [2568, 2569].map((year) =>
  path.join(directory, "RACK-A-" + year + ".xlsx"),
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
        "(" + exercise.toString() + ")(" + JSON.stringify(phase) + ")",
      );
      assert.match(result, /ตรวจรายการกับคลังอุปกรณ์/);
      assert.match(result, /ค่าแรกของต้นใหม่.*4/);
      assert.doesNotMatch(result, /ลดลง/);
      console.log("Electron UI passed");
      app.exit(0);
    } catch (error) {
      console.error(error);
      app.exit(1);
    }
  });
});
async function exercise(phase) {
  const $ = (selector) => document.querySelector(selector);
  async function wait(check) {
    const end = Date.now() + 15000;
    while (!check()) {
      if (Date.now() > end) throw Error("หมดเวลารอหน้าจอ");
      await new Promise((resolve) => setTimeout(resolve, 40));
    }
  }
  if (phase === "write") {
    const buttons = document.querySelectorAll("#slots button");
    buttons[0].click();
    await wait(() => !$("main").inert);
    buttons[1].click();
    await wait(() => !$("#create").disabled && !$("main").inert);
    $("#create").click();
    await wait(() => !$("#work").hidden);
    [...document.querySelectorAll("#points button")]
      .find((n) => n.textContent.includes("/ 2 ·"))
      .click();
    const choice = $("form select");
    choice.value = "replacement";
    choice.dispatchEvent(new Event("change"));
    $("form input[type=date]").value = "2026-07-27";
    $("form textarea").value = "ตรวจรายการกับคลังอุปกรณ์";
    $("form button").click();
    await wait(() => $("#message").textContent.includes("บันทึกผลตรวจแล้ว"));
  }
  $("#archive").click();
  await wait(() => $("#jobs button"));
  $("#jobs button").click();
  await wait(() => !$("#work").hidden);
  [...document.querySelectorAll("#points button")]
    .find((n) => n.textContent.includes("/ 2 ·"))
    .click();
  return $("#detail").innerText;
}
require("../main");
