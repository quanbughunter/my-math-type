// Cầu nối an toàn giữa giao diện và clipboard của Windows.
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('mymathNative', {
  platform: process.platform,
  copyText: (text) => ipcRenderer.invoke('clip:text', text),
  copyImage: (dataUrl) => ipcRenderer.invoke('clip:image', dataUrl),
});
