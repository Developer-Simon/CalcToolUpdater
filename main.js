const { info } = require('console');
const {app, BrowserWindow, Menu, protocol, ipcMain, MenuItem, shell, dialog, webContents, ipcRenderer} = require('electron');
const log = require('electron-log');
const Store = require('electron-store');
const {autoUpdater} = require("electron-updater");
const regedit = require('regedit');
const fs = require('fs');
const convert = require('xml-js');
const { installCalcTool } = require('./modules/installer');

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
              'HKCU\\SOFTWARE\\CalculationTool\\Version'])
.on('data', function(entry) {
  g_registryItems.push(entry);
})
.on('finish', function() {
  //dialog.showMessageBox(win, {message: "All Registry Keys loaded!"});
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
// Window content
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

function createDefaultWindow() {
  //win = new BrowserWindow({ width: 400, height: 320});
  win = new BrowserWindow({ width: 1200, height: 1000});
  win.webContents.openDevTools();
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
  sendUpdateStatus('Update heruntergeladen! Starten sie den CalcToolClient neu, um das CalculationTool Update zu beginnen.');
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
ipcMain.on('installButton', function(event, arg) {
  installCalcTool(win, g_sResFolder, g_registryItems);
  win.setProgressBar(0.1);
})
ipcMain.on('debugbtnclick', function(event, arg) {
  var sAddinPath = app.getPath('appData') + "\\Microsoft\\AddIns\\CalculationTool\\";
  var sNewConfig = fs.readFileSync(sAddinPath + "CalculationTool.xml", { flag: 'r'});
  var sOldConfig = fs.readFileSync(sAddinPath + "CalculationTool_old.xml", { flag: 'r'})
  var NewConfig = convert.xml2js(sNewConfig, {compact: false});
  var OldConfig = convert.xml2js(sOldConfig, {compact: false});
  copySettings(NewConfig, OldConfig);
  sNewConfig = convert.js2xml(NewConfig, {compact: false})
  fs.writeFileSync(sAddinPath + "CalculationTool.xml", sNewConfig);
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
