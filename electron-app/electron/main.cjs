const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');

const isDev = !app.isPackaged;
const SETTINGS_FILE = path.join(app.getPath('userData'), 'vyapar-settings.json');

const DEFAULT_SETTINGS = {
  companyName: '',
  addressLine: '',
  mobile1: '',
  mobile2: '',
  reportTitle: 'City Wise Outstanding Receivable',
  fromDate: '',
  toDate: '',
};

function readSettings() {
  try {
    if (fs.existsSync(SETTINGS_FILE)) {
      return { ...DEFAULT_SETTINGS, ...JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf-8')) };
    }
  } catch (_e) {}
  return { ...DEFAULT_SETTINGS };
}

function writeSettings(settings) {
  try {
    fs.writeFileSync(SETTINGS_FILE, JSON.stringify(settings, null, 2), 'utf-8');
    return true;
  } catch (_e) {
    return false;
  }
}

let mainWindow = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 860,
    minWidth: 800,
    minHeight: 600,
    title: 'Vyapar Outstanding Formatter',
    icon: path.join(__dirname, '..', 'build', 'icon.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
    },
    backgroundColor: '#f0ede8',
    show: false,
  });

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.setMenuBarVisibility(false);

  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ── IPC Handlers ────────────────────────────────────────────

ipcMain.handle('load-settings', () => readSettings());

ipcMain.handle('save-settings', (_e, settings) => writeSettings(settings));

ipcMain.handle('print-report', () => {
  if (!mainWindow) return false;
  mainWindow.webContents.print(
    { silent: false, printBackground: true, pageSize: 'A4' },
    (success, reason) => { if (!success) console.error('Print failed:', reason); }
  );
  return true;
});

ipcMain.handle('save-pdf', async (_e, htmlContent) => {
  if (!mainWindow) return { success: false, error: 'No window' };

  const { filePath, canceled } = await dialog.showSaveDialog(mainWindow, {
    title: 'Save Outstanding Report as PDF',
    defaultPath: path.join(app.getPath('documents'), 'outstanding-report.pdf'),
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
  });

  if (canceled || !filePath) return { success: false, error: 'Cancelled' };

  try {
    const pdfWin = new BrowserWindow({
      show: false,
      width: 900,
      height: 1200,
      webPreferences: { nodeIntegration: false, contextIsolation: true },
    });

    await pdfWin.loadURL(
      'data:text/html;charset=utf-8,' + encodeURIComponent(htmlContent)
    );

    const pdfData = await pdfWin.webContents.printToPDF({
      pageSize: 'A4',
      landscape: false,
      printBackground: true,
      margins: {
        marginType: 'custom',
        top: 15,
        bottom: 15,
        left: 15,
        right: 15,
      },
    });

    pdfWin.close();
    fs.writeFileSync(filePath, pdfData);
    shell.showItemInFolder(filePath);
    return { success: true, filePath };
  } catch (e) {
    return { success: false, error: String(e) };
  }
});
