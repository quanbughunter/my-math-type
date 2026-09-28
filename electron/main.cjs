// Ứng dụng Windows (Electron): mở giao diện MyMath đã build trong thư mục dist/.
const { app, BrowserWindow, ipcMain, clipboard, nativeImage, shell, Menu } = require('electron');
const path = require('node:path');

if (!app.requestSingleInstanceLock()) app.quit();

let win;
function createWindow() {
  win = new BrowserWindow({
    width: 1120,
    height: 800,
    minWidth: 380,
    minHeight: 520,
    title: 'MyMath',
    backgroundColor: '#f3f4f8',
    icon: path.join(__dirname, '..', 'build', 'icon.png'),
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      sandbox: true,
      spellcheck: false,
    },
  });
  Menu.setApplicationMenu(null);
  win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  // liên kết ngoài mở bằng trình duyệt
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (!url.startsWith('file:')) { e.preventDefault(); shell.openExternal(url); }
  });
  // F12 mở công cụ gỡ lỗi, Ctrl +/- phóng to
  win.webContents.on('before-input-event', (e, input) => {
    if (input.type === 'keyDown' && input.key === 'F12') win.webContents.toggleDevTools();
  });
}

ipcMain.handle('clip:text', (_e, text) => { clipboard.writeText(String(text)); return true; });
ipcMain.handle('clip:image', (_e, dataUrl) => {
  clipboard.writeImage(nativeImage.createFromDataURL(String(dataUrl)));
  return true;
});

app.on('second-instance', () => { if (win) { if (win.isMinimized()) win.restore(); win.focus(); } });
app.whenReady().then(createWindow);
app.on('window-all-closed', () => app.quit());
