const {app, BrowserWindow, Menu, protocol, ipcMain, MenuItem, shell, dialog, webContents} = require('electron');
const log = require('electron-log');
const {autoUpdater} = require("electron-updater");

//-------------------------------------------------------------------
// Logging
//-------------------------------------------------------------------
autoUpdater.logger = log;
autoUpdater.logger.transports.file.level = 'info';
log.info('App starting...');

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
  sendStatusToWindow('Checking for update...');
})
autoUpdater.on('update-available', (info) => {
  sendStatusToWindow('Update available.');
})
autoUpdater.on('update-not-available', (info) => {
  sendStatusToWindow('Update not available.');
})
autoUpdater.on('error', (err) => {
  sendStatusToWindow('Error in auto-updater. ' + err);
})
autoUpdater.on('download-progress', (progressObj) => {
  let log_message = "Download speed: " + progressObj.bytesPerSecond;
  log_message = log_message + ' - Downloaded ' + progressObj.percent + '%';
  log_message = log_message + ' (' + progressObj.transferred + "/" + progressObj.total + ')';
  sendStatusToWindow(log_message);
})
autoUpdater.on('update-downloaded', (info) => {
  sendStatusToWindow('Update downloaded');
});

//-------------------------------------------------------------------
// Program execution (main)
//-------------------------------------------------------------------
app.on('ready', function() {
  // Create the Menu
  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);

  createDefaultWindow();

  autoUpdater.autoDownload = false;
  autoUpdater.allowPrerelease = true;
  autoUpdater.checkForUpdates();
});

app.on('window-all-closed', () => {
  app.quit();
});

//-------------------------------------------------------------------
// ipc events
//-------------------------------------------------------------------
ipcMain.on('btnclick', function(event, arg) {
  var installerProcess = require('child_process').execFile;
  var executablePath = app.getAppPath();
  var resPos = executablePath.search("resources");
  executablePath = executablePath.slice(0, resPos);
  executablePath = executablePath + "files\\CalcToolInstaller_V1.1.0.1.EXE"
  sendStatusToWindow(executablePath);
  installerProcess(executablePath, function(err, data) {
    if(err) {
      sendStatusToWindow("Error occured: " + err);
      return;
    }
 
    sendStatusToWindow(data.toString());
  });
})

//-------------------------------------------------------------------
// Auto updates 
//-------------------------------------------------------------------
/*
app.on('ready', function()  {
  autoUpdater.autoDownload = false;
  autoUpdater.allowPrerelease = true;
  autoUpdater.checkForUpdates();
});
*/
autoUpdater.on('checking-for-update', () => {
})
autoUpdater.on('update-available', (info) => {
})
autoUpdater.on('update-not-available', (info) => {
})
autoUpdater.on('error', (err) => {
})
autoUpdater.on('download-progress', (progressObj) => {
})
autoUpdater.on('update-downloaded', (info) => {
})
