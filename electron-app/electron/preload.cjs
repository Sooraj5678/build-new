const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  isElectron: true,
  loadSettings: () => ipcRenderer.invoke('load-settings'),
  saveSettings: (settings) => ipcRenderer.invoke('save-settings', settings),
  printReport: () => ipcRenderer.invoke('print-report'),
  savePDF: (htmlContent) => ipcRenderer.invoke('save-pdf', htmlContent),
});
