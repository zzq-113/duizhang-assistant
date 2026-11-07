'use strict';

const { app, BrowserWindow, dialog, ipcMain, session } = require('electron');
const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const APP_NAME = '清核对账助手';
const MAX_REPORT_BYTES = 200 * 1024 * 1024;
const indexPath = path.join(__dirname, 'index.html');
const trustedPage = pathToFileURL(indexPath).href;

app.enableSandbox();
app.setName(APP_NAME);

function isTrustedSender(event) {
  const url = event.senderFrame?.url || '';
  return url === trustedPage || url.startsWith(`${trustedPage}#`);
}

function safeReportName(value) {
  const name = path.basename(String(value || '订单收款对账报告.xlsx'))
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_');
  return name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`;
}

ipcMain.handle('save-workbook', async (event, payload) => {
  if (!isTrustedSender(event)) throw new Error('不受信任的保存请求');
  const bytes = payload?.bytes;
  if (!bytes || typeof bytes.byteLength !== 'number' || bytes.byteLength > MAX_REPORT_BYTES) {
    throw new Error('报告数据无效或超过 200 MB');
  }
  const fileName = safeReportName(payload.fileName);
  const choice = await dialog.showSaveDialog({
    title: '保存 Excel 对账报告',
    defaultPath: path.join(app.getPath('downloads'), fileName),
    filters: [{ name: 'Excel 工作簿', extensions: ['xlsx'] }],
    properties: ['createDirectory', 'showOverwriteConfirmation']
  });
  if (choice.canceled || !choice.filePath) return { saved: false };
  await fs.writeFile(choice.filePath, Buffer.from(bytes));
  return { saved: true };
});

function createWindow() {
  const window = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 900,
    minHeight: 650,
    title: APP_NAME,
    backgroundColor: '#f7f9fa',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      webSecurity: true
    }
  });

  window.loadFile(indexPath);
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.webContents.on('will-navigate', (event, url) => {
    if (!(url === trustedPage || url.startsWith(`${trustedPage}#`))) event.preventDefault();
  });
}

app.whenReady().then(() => {
  session.defaultSession.setPermissionRequestHandler((_webContents, _permission, callback) => callback(false));
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
