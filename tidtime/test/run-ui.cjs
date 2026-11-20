const { spawnSync } = require("node:child_process");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const directory = fs.mkdtempSync(path.join(os.tmpdir(), "tidtime-ui-"));
try {
  for (const phase of ["write", "reopen"]) {
    const child = spawnSync(
      require("electron"),
      ["--headless", path.join(__dirname, "ui.cjs"), phase, directory],
      { stdio: "inherit", timeout: 60000, windowsHide: true },
    );
    if (child.error) throw child.error;
    if (child.status !== 0) throw Error("Electron test failed: " + phase);
  }
} finally {
  fs.rmSync(directory, {
    recursive: true,
    force: true,
    maxRetries: 5,
    retryDelay: 100,
  });
}
