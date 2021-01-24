const {app, BrowserWindow, Menu, protocol, ipcMain, MenuItem, shell, dialog, webContents} = require('electron');
const log = require('electron-log');
const Store = require('electron-store');
const {autoUpdater} = require("electron-updater");

//-------------------------------------------------------------------
// Logging
//-------------------------------------------------------------------
autoUpdater.logger = log;
autoUpdater.logger.transports.file.level = 'info';
log.info('App starting...');
const store = new Store({
  name: 'storage',
  clearInvalidConfig: false
});

//-------------------------------------------------------------------
// Define the menu
//-------------------------------------------------------------------
let template = []
if (process.platform === 'darwin') {
  // OS X
  const name = app.getName();
  template.unshift({
    label: 'Menu',
    submenu: [
      {
        label: 'About ' + name,
        click() { sendStatusToWindow('You clicked about :)')}
      },
      {
        label: 'Quit',
        accelerator: 'Command+Q',
        click() { app.quit(); }
      },
    ]
  })
}
else if (process.platform == 'win32') {
  // windows
  const name = app.getName();
  template.unshift({
    label: 'Menu',
    submenu: [
      {
        label: 'About ' + name,
        role: 'about'
      },
      {
        label: 'Quit',
        accelerator: "Command+Q",
        click() { app.quit(); }
      },
    ]
  })
}


//-------------------------------------------------------------------
// Open a window that displays the version
//
// THIS SECTION IS NOT REQUIRED
//
// This isn't required for auto-updates to work, but it's easier
// for the app to show a window than to have to click "About" to see
// that updates are working.
//-------------------------------------------------------------------
let win;

function sendStatusToWindow(text) {
  log.info(text);
  win.webContents.send('message', text);
}
function createButtonAtWindow(text) {
  win.webContents.send('add-update-button', text);
}
function sendUpdateStatus(text) {
  win.webContents.send('update-status', text);
}
function startInstaller(fileName) {
  var installerProcess = require('child_process').execFile;
  var executablePath = app.getAppPath();
  var resPos = executablePath.search("resources");
  executablePath = executablePath.slice(0, resPos);
  executablePath = executablePath + "files\\" + fileName;
  installerProcess(executablePath, function(err, data) {
    if(err) {
      sendStatusToWindow("Error occured: " + err);
      return;
    }
  });
}

function createDefaultWindow() {
  win = new BrowserWindow({ width: 400, height: 320});
  //win = new BrowserWindow({ width: 1200, height: 1000});
  //win.webContents.openDevTools();
  win.on('closed', () => {
    win = null;
  });
  win.loadURL(`file://${__dirname}/index.html#v${app.getVersion()}`);
  return win;
}
autoUpdater.on('checking-for-update', () => {
  sendUpdateStatus('Checking for update...');
})
autoUpdater.on('update-available', (info) => {
  //sendStatusToWindow('Update available: ' + info.version);
})
autoUpdater.on('update-not-available', (info) => {
  //sendStatusToWindow('Update not available.');
})
autoUpdater.on('error', (err) => {
  sendUpdateStatus('Error in auto-updater. ' + err);
})
autoUpdater.on('download-progress', (progressObj) => {
  let log_message = "Download speed: " + Math.round(progressObj.bytesPerSecond / (1024 * 2) * 1000) / 1000 + ' KB/s';
  log_message = log_message + ' - Downloaded ' + Math.round(progressObj.percent * 10) / 10 + '%';
  log_message = log_message + ' (' + Math.round(progressObj.transferred / (1024 * 2) * 1000) / 1000 + ' KB/' + Math.round(progressObj.total / (1024 * 2) * 1000) / 1000 + ' KB)';
  sendUpdateStatus(log_message);
})
autoUpdater.on('update-downloaded', (info) => {
  sendUpdateStatus('Update downloaded. Application will be relaunched, soon!');
  app.relaunch();
  app.exit();
});

//-------------------------------------------------------------------
// Windows creation
//-------------------------------------------------------------------
app.on('ready', function() {
  // Create the Menu
  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);

  createDefaultWindow();
  if (store.get('download-done') == 1) {
    startInstaller(`CalcToolInstaller_V${app.getVersion()}.EXE`);
    store.set('download-done', 0);
  }
});

app.on('window-all-closed', () => {
  app.quit();
});

//-------------------------------------------------------------------
// ipc events
//-------------------------------------------------------------------
ipcMain.on('installButton', function(event, arg) {
  startInstaller(`CalcToolInstaller_V${app.getVersion()}.EXE`);
})
ipcMain.on('debugbtnclick', function(event, arg) {
  sendStatusToWindow('download-done = ' + store.get('download-done'));
  store.set('download-done', 1);
  sendStatusToWindow('download-done = ' + store.get('download-done'));
})
ipcMain.on('downloadUpdate', function(event, arg) {
  autoUpdater.downloadUpdate();
})

//-------------------------------------------------------------------
// Auto updates 
//-------------------------------------------------------------------
app.on('ready', function()  {
  autoUpdater.autoDownload = false;
  autoUpdater.allowPrerelease = true;
  autoUpdater.checkForUpdates();
});
autoUpdater.on('checking-for-update', () => {
})
autoUpdater.on('update-available', (info) => {
  createButtonAtWindow('v' + info.version + ' herunterladen');
})
autoUpdater.on('update-not-available', (info) => {
})
autoUpdater.on('error', (err) => {
})
autoUpdater.on('download-progress', (progressObj) => {
})
autoUpdater.on('update-downloaded', (info) => {
  store.set('download-done', true);
})
