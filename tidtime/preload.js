const { contextBridge, ipcRenderer } = require("electron");
const api = {};
for (const name of ["choose", "sheet", "create", "list", "load", "review"]) {
  api[name] = (...args) => ipcRenderer.invoke(name, ...args);
}
contextBridge.exposeInMainWorld("tidtime", api);
