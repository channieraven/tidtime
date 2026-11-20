const { spawnSync } = require("node:child_process");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const directory = fs.mkdtempSync(path.join(os.tmpdir(), "tidtime-ui-"));
async function run() {
  await require("../samples").writeSamples(directory);
  // Exercise the import screen with January and June observations in one year.
  const ExcelJS = require("exceljs");
  for (const [year, date] of [
    [2568, "15/01/2569"],
    [2569, new Date("2026-06-15")],
  ]) {
    const file = path.join(directory, "P01-" + year + ".xlsx");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(file);
    workbook.getWorksheet("สำรวจ").eachRow((row, line) => {
      if (line === 1) return;
      row.getCell(5).value = 2569;
      row.getCell(6).value = date;
    });
    await workbook.xlsx.writeFile(file);
  }

  for (const phase of ["write", "reopen"]) {
    const child = spawnSync(
      require("electron"),
      ["--headless", path.join(__dirname, "ui.cjs"), phase, directory],
      { stdio: "inherit", timeout: 60000, windowsHide: true },
    );
    if (child.error) throw child.error;
    if (child.status !== 0) throw Error("Electron test failed: " + phase);
  }
}
run()
  .finally(() => {
    fs.rmSync(directory, {
      recursive: true,
      force: true,
      maxRetries: 5,
      retryDelay: 100,
    });
  })
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
