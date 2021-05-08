const { info } = require('console');
const {app, BrowserWindow, Menu, protocol, ipcMain, MenuItem, shell, dialog, webContents, ipcRenderer} = require('electron');
const log = require('electron-log');
const Store = require('electron-store');
const {autoUpdater} = require("electron-updater");
const regedit = require('regedit');
const fs = require('fs');
const convert = require('xml-js');
const {_Version, installCalcTool} = require('./modules/installer');

const debug = true;


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
// Files-Path
//-------------------------------------------------------------------

var g_sResFolder = app.getAppPath();
{
  var nResourcePos = g_sResFolder.search("resources");
  if (nResourcePos == -1) { // debug only
    g_sResFolder = g_sResFolder + "\\res\\";
  }
  else {
    g_sResFolder = g_sResFolder.slice(0,nResourcePos) + "res\\";
  }  
}

//-------------------------------------------------------------------
// Registry Keys
//-------------------------------------------------------------------

var g_registryItems = Array(0);
regedit.setExternalVBSLocation(g_sResFolder + "winreg-vbs\\");
regedit.list(['HKCU\\SOFTWARE\\CalculationTool\\Language', 
              'HKCU\\SOFTWARE\\CalculationTool\\Version',
              'HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\Excel'])
.on('data', function(entry) {
  g_registryItems.push(entry);
})
.on('finish', function() {
  var bExcelFound = false;
  var AddInVersion = new _Version("");
  var UpdaterVersion = new _Version(app.getVersion());

  AddInVersion._getVersionFromRegistry(g_registryItems);
  if (AddInVersion.raw != "") {
    setTimeout(sendMessage, 1000, 'excel-addin-version' , "v" + AddInVersion.raw);
  }

  if (AddInVersion.isOlderThan(UpdaterVersion)) {
    setTimeout(sendMessage, 1000, 'excel-addin-version-old', true);
  }

  g_registryItems.forEach(element => {
    // get Excel version
    if (element.key.includes('Microsoft\\Office\\16.0\\Excel')) {
      try {
        if (element.data.values['ExcelName']) {
          bExcelFound = true;
        }
      } catch (error) {
        bExcelFound = false
      }
    }
  });

  setTimeout(sendMessage, 1000, 'excel-installation-status', bExcelFound);
})
.on('error', function(err) {
  ;
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
        click() { dialog.showMessageBox(win, { message: 'You clicked about :)' }) }
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
// Window content
//-------------------------------------------------------------------
let win;
var winWidth = 500, winHeight = 340;

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
function sendMessage(event, arg) {
  log.info("ipcRenderer event: " + event + " - " + toString(arg));
  win.webContents.send(event, arg);
}

function createDefaultWindow() {
  if (debug) {
    winWidth = 1200;
    winHeight = 1000;
  }
  win = new BrowserWindow({ width: winWidth, height: winHeight});
  if (debug) {
    win.webContents.openDevTools();
  }
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
  sendStatusToWindow('Error in auto-updater. ' + err);
})
autoUpdater.on('download-progress', (progressObj) => {
  let log_message = "Download speed: " + Math.round(progressObj.bytesPerSecond / (1024 * 2) * 1000) / 1000 + ' KB/s';
  log_message = log_message + ' - Downloaded ' + Math.round(progressObj.percent * 10) / 10 + '%';
  log_message = log_message + ' (' + Math.round(progressObj.transferred / (1024 * 2) * 1000) / 1000 + ' KB/' + Math.round(progressObj.total / (1024 * 2) * 1000) / 1000 + ' KB)';
  sendUpdateStatus(log_message);
  win.setSize(winWidth, winHeight + 30);
})
autoUpdater.on('update-downloaded', (info) => {
  sendUpdateStatus("Update heruntergeladen! Um das AddIn zu updaten,\nstarten sie den CalcToolClient neu!");
  win.setSize(winWidth, winHeight + 60);
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
    setTimeout(installCalcTool, 2000, win, g_sResFolder, g_registryItems);
    store.set('download-done', 0);
  }
});

app.on('window-all-closed', () => {
  app.quit();
});

//-------------------------------------------------------------------
// ipc events
//-------------------------------------------------------------------
ipcMain.on('install-button', function(event, arg) {
  if (arg = 'click') {
    installCalcTool(win, g_sResFolder, g_registryItems);
    win.setProgressBar(0.1);
  }
})
ipcMain.on('debugbtnclick', function(event, arg) {
  autoUpdater.checkForUpdates();
  //importSettings(g_sResFolder, app.getPath('appData') + "\\Microsoft\\AddIns\\CalculationTool\\", win);
})
ipcMain.on('download-update-req', function(event, arg) {
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
