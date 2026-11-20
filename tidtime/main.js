const { app, BrowserWindow, dialog, ipcMain, session } = require("electron");
const path = require("node:path");
const {
  openExcel,
  sheetOptions,
  readSheet,
  setSurveyYear,
  columns,
  displayDate,
} = require("./excel");
const { pairRounds, compareTree } = require("./compare");
const { Store } = require("./store");
const slots = new Map();
let store;
function slot(side) {
  if (!["before", "after"].includes(side)) throw Error("รอบไม่ถูกต้อง");
  return slots.get(side);
}
function view(job) {
  const histories = new Map();
  for (const e of job.history) {
    if (!histories.has(e.key)) histories.set(e.key, []);
    histories.get(e.key).push(e);
  }
  const points = pairRounds(job.before, job.after).map((p) => {
    const history = histories.get(p.key) || [];
    const latest = history.at(-1);
    const result = compareTree(
      p,
      latest?.decision === "pending" ? undefined : latest,
    );
    if (latest?.decision === "pending") result.needsReview = true;
    return {
      ...p,
      ...result,
      history,
      displayDates: {
        before: p.before.map((row) => displayDate(row.date)),
        after: p.after.map((row) => displayDate(row.date)),
      },
    };
  });
  return {
    id: job.id,
    before: summary(job.before),
    after: summary(job.after),
    points,
    columns,
  };
}
function summary(round) {
  const { file, sheet, plot, year, rows } = round;
  return {
    file,
    sheet,
    plot,
    year,
    form: round.form,
    count: rows.length,
    dates: [...new Set(rows.map((r) => displayDate(r.date)))].join(", "),
  };
}
function register() {
  ipcMain.handle("choose", async (_, side) => {
    slot(side);
    const pick = await dialog.showOpenDialog({
      properties: ["openFile"],
      filters: [{ name: "Excel", extensions: ["xlsx"] }],
    });
    if (pick.canceled) return null;
    slots.delete(side);
    const source = await openExcel(pick.filePaths[0]);
    slots.set(side, { source });
    return {
      file: source.file,
      sheets: sheetOptions(source),
    };
  });
  ipcMain.handle("sheet", (_, side, name) => {
    const item = slot(side);
    item.round = null;
    item.round = readSheet(item.source, name);
    return summary(item.round);
  });
  ipcMain.handle("create", (_, years = {}) => {
    const a = slot("before")?.round,
      b = slot("after")?.round;
    if (!a || !b) throw Error("กรุณาเลือกทั้งสองชีต");
    return view(
      store.create(
        setSurveyYear(a, years.before),
        setSurveyYear(b, years.after),
      ),
    );
  });
  ipcMain.handle("list", () => store.list());
  ipcMain.handle("load", (_, id) => view(store.load(id)));
  ipcMain.handle("review", (_, id, key, data) =>
    view(store.review(id, key, data)),
  );
}
if (!app.requestSingleInstanceLock()) app.quit();
else
  app.whenReady().then(() => {
    store = new Store(path.join(app.getPath("userData"), "surveys"));
    register();
    session.defaultSession.webRequest.onBeforeRequest(
      { urls: ["http://*/*", "https://*/*", "ws://*/*", "wss://*/*"] },
      (_, cb) => cb({ cancel: true }),
    );
    const win = new BrowserWindow({
      width: 1250,
      height: 850,
      minWidth: 840,
      minHeight: 620,
      autoHideMenuBar: true,
      webPreferences: {
        preload: path.join(__dirname, "preload.js"),
        sandbox: true,
        contextIsolation: true,
        nodeIntegration: false,
      },
    });
    win.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
    win.webContents.on("will-navigate", (event) => event.preventDefault());
    win.loadFile(path.join(__dirname, "index.html"));
  });
app.on("window-all-closed", () => app.quit());
