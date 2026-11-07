'use strict';

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('desktopBridge', Object.freeze({
  saveWorkbook(fileName, bytes) {
    return ipcRenderer.invoke('save-workbook', { fileName, bytes });
  }
}));
