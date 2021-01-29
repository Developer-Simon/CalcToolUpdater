const { info } = require('console');
const {app, BrowserWindow, Menu, protocol, ipcMain, MenuItem, shell, dialog, webContents, ipcRenderer} = require('electron');
const log = require('electron-log');
const Store = require('electron-store');
const {autoUpdater} = require("electron-updater");
const regedit = require('regedit');
const path = require('path');
const fs = require('fs');

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

// ------------------------------------------------------------------
// Installation process
// ------------------------------------------------------------------

function formatVersion(versionObj) {
  var versionSub = new String(versionObj.raw);
  versionObj.major = Number(versionSub.substring(1, versionSub.search('.')));
  versionSub = versionSub.substring(versionSub.search('.') + 2, versionSub.length);
  versionObj.minor = Number(versionSub.substring(1, versionSub.search('.')));
  versionSub = versionSub.substring(versionSub.search('.') + 2, versionSub.length);
  versionObj.patch = Number(versionSub.substring(1, versionSub.search('.')));
  return versionObj;
}

function installCalcTool()
{
  // needed vars
  var sAddinPath = app.getPath('appData') + "\\Microsoft\\AddIns\\CalculationTool\\";
  var OldVersion = {raw: "", major: -1, minor: -1, patch: -1};
  var NewVersion = {raw: app.getVersion(), major: -1, minor: -1, patch: -1};
  var sLanguage = "DE";

  // for later use
  if (g_sResFolder == "") {
    dialog.showMessageBox(win, {
      type: "error",
      title: "Dateien nicht gefunden!",
      message: "Das Zielverzeichnis konnte nicht gefunden werden!"
    });
    return;
  }

  // get version out of registry
  g_registryItems.forEach(element => {
    if (element.key.includes('CalculationTool\\Version')) {
      OldVersion.raw = element.data.values[''].value;
    }
  });
  // check installation
  if (OldVersion.raw == "") { 
    var msgResult = dialog.showMessageBox(win, { // for version V1.0.4.0 or older
      type: "warning",
      buttons: ["Abbrechen","Fortfahren"],
      defaultId: 1,
      title: "Update für ältere Versionen akutell nicht verfügbar!",
      message: "Bei einer Neuinstallation drücken Sie einfach \"Fortfahren\".\nFalls Sie bereits eine ältere Version des CalculationTool (V1.0.4.1 oder älter) installiert haben, drücken Sie \"Abbrechen\"!"
    });
    if (msgResult != 1) {
      dialog.showMessageBox(win, {
        type: "info",
        title: "Weiteres Vorgehen",
        message: "Zum Updaten von älteren Installation haben Sie zwei Optionen:\n- Deinstallieren Sie die alte Version des CalculationTool (Achtung: Alle Einstellungen und Datenbanken werden gelöscht!)\n- Updaten sie ihre bestehende Version auf die Version V1.1.0.0 oder höher"
      })
      return;
    }
  } else { 
    // format versions
    OldVersion = formatVersion(OldVersion);
    NewVersion = formatVersion(NewVersion);

    // check version update intensity
    if (OldVersion.major < NewVersion.major) {
      var msgResult = dialog.showMessageBox(win, {
        type: "warning",
        buttons: ["Abbrechen","Fortfahren"],
        defaultId: 1,
        title: "Updates auf neue Haupt-Version",
        message: "Achtung, bei Updates auf eine neue Haupt-Version kann es zu Inkompatibilitäten zu alten Versionen kommen. Dabei könnten sowohl Daten als auch Einstellungen verloren gehen. Es wird empfohlen diese im Vorfeld manuell zu sichern!"
      });
      if (msgResult != 1) return;
    } else if (OldVersion.minor < NewVersion.minor ) {
      var msgResult = dialog.showMessageBox(win, {
        type: "info",
        buttons: ["Abbrechen", "Ja"],
        defaultId: 1,
        title: "Neue Version installieren",
        message: "Update von Version " + OldVersion.raw + " auf Version " + NewVersion.raw + " jetzt durchführen?"
      });
      if (msgResult != 1) return;
    } else if (OldVersion.patch < NewVersion.patch) {
      var msgResult = dialog.showMessageBox(win, {
        type: "info",
        buttons: ["Abbrechen", "Ja"],
        defaultId: 1,
        title: "Neue Version installieren",
        message: "Update von Version " + OldVersion.raw + " auf Version " + NewVersion.raw + " jetzt durchführen?"
      });
      if (msgResult != 1) return;
    } else if (OldVersion.raw == NewVersion.raw) {
      var msgResult = dialog.showMessageBox(win, {
        type: "warning",
        buttons: ["Abbrechen", "Fortfahren"],
        defaultId: 0,
        title: "Version bereits installiert",
        message: "Sie haben die aktuelle Version bereits installiert! Wollen Sie trotzdem die Installation fortführen?"
      });
      if (msgResult != 1) return;
    } else if (OldVersion.major == NewVersion.major) {
      var msgResult = dialog.showMessageBox(win, {
        type: "warning",
        buttons: ["Abbrechen", "Trotzdem durchführen"],
        defaultId: 0,
        title: "Alte Version installieren",
        message: "Sie haben bereits die Version " + OldVersion.raw + " installiert. Sind Sie sicher, dass sie auf eine alte Version (V" + NewVersion.raw + ") ein Downgrade durchführen wollen?"
      });
      if (msgResult != 1) return;
    } else {
      dialog.showMessageBox(win, {
        type: "error",
        title: "Updates auf alte Haupt-Version",
        message: "Ein Downgrade auf eine alte Version ist mit diesem Updater nicht möglich!"
      })
      return;
    }
  }

  // install Add-In, .dbasync file-information and -icon in registry
  // TODO: check admin permissions
  var aCreateKeys = 
  ["HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Add-in Manager\\",
   "HKCU\\Software\\CalculationTool\\Version\\",
   "HKCU\\Software\\CalculationTool\\Language\\",
   "HKCR\\.dbasync\\DefaultIcon\\"];
  aCreateKeys.forEach(key => {
    regedit.createKey(key, function(err) {
      if (err !== undefined) {
        dialog.showMessageBox(win, {message: "Fehler beim Registrieren: " + err.toString()});
      }
    });
  });
  var aSetValues = {
    'HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Add-in Manager\\': {
      [sAddinPath + "CalculationTool.xlam"]: {
        value: '',
        type: 'REG_SZ'
      }
    },
    'HKCU\\Software\\CalculationTool\\Version\\': {
      'default': {
        value: NewVersion.raw,
        type: 'REG_DEFAULT'
      }
    },
    'HKCU\\Software\\CalculationTool\\Language\\': {
      'default': {
        value: sLanguage,
        type: 'REG_DEFAULT'
      }
    },
    'HKCR\\.dbasync\\': {
      'default': {
        value: 'CalculationTool Synchronisationsdatei',
        type: 'REG_DEFAULT'
      },
      'EditFlags': {
        value: 0x00120401,
        type: 'REG_DWORD'
      }
    },
    'HKCR\\.dbasync\\DefaultIcon\\': {
      'default': {
        value: 'C:\\Windows\\System32\\imageres.dll,215',
        type: 'REG_DEFAULT'
      }
    }
  }
  regedit.putValue(aSetValues, function(err) {
    if (err !== undefined) {
      dialog.showMessageBox(win, {message: err.toString()});
    }
  });
  
  // Create Folders
  if (!fs.existsSync(sAddinPath)) fs.mkdirSync(sAddinPath);
  if (!fs.existsSync(sAddinPath + "DataBases\\")) fs.mkdirSync(sAddinPath + "DataBases\\");
  if (!fs.existsSync(sAddinPath + "Request\\")) fs.mkdirSync(sAddinPath + "Request\\");
  
  // Delete V1.0.1 or older
  if (fs.existsSync(path.join(sAddinPath, "\\..\\CalculationTool.xlam"))) fs.unlinkSync(path.join(sAddinPath, "\\..\\CalculationTool.xlam"));
  // Delete CalculationTool.inf of V1.0.4.1
  if (fs.existsSync(path.join(sAddinPath, "\\CalculationTool.inf"))) fs.unlinkSync(path.join(sAddinPath, "\\CalculationTool.inf"));

  // Rename old calcTool for saving Settings
  if (fs.existsSync(sAddinPath + "CalculationTool_old.xlam")) fs.unlinkSync(sAddinPath + "CalculationTool_old.xlam");
  if (fs.existsSync(sAddinPath + "CalculationTool.xlam")) {
    fs.rename(sAddinPath + "CalculationTool.xlam", sAddinPath + "CalculationTool_old.xlam", function(err) {
      if (err != null && err.code == 'EBUSY') {
        dialog.showMessageBox(win, {
          type: "warning",
          title: "Datei wird von einer anderen Anwendung verwendet",
          message: "Während Excel ausgeführt wird, kann das CalculationTool nicht aktualisiert werden. Beenden Sie Excel und starten die Installation erneut!"});
      } else if (err != null) {
        dialog.showMessageBox(win, {message: "Unerwarter Fehler: " + err.message});
      }
    });
  } 

  // Copy files
  fs.readdirSync(g_sResFolder + "calctool-files\\").forEach(file => {
    fs.copyFile(g_sResFolder + "calctool-files\\" + file, sAddinPath + file, function(err) {
      if (err != null) {
        dialog.showMessageBox(win, {message: "Unerwarter Fehler: " + err.message});
      }
    });
  });

  // Move DataBases
  if (OldVersion.raw != "") {
    fs.rename(sAddinPath + "ProductDataBase.xlsx", sAddinPath + "DataBases\\ProductDataBase.xlsx", function(err) {
      if (err != null) {
        dialog.showMessageBox(win, {message: "Unerwarter Fehler: " + err.message});
      }
    })
    fs.rename(sAddinPath + "ProjectDataBase.xlsx", sAddinPath + "DataBases\\ProjectDataBase.xlsx", function(err) {
      if (err != null) {
        dialog.showMessageBox(win, {message: "Unerwarter Fehler: " + err.message});
      }
    })
  }

  // TODO - Settings in xml

  var sResult = dialog.showMessageBox(win, {
    type: "info",
    buttons: ["Nein", "Ja"],
    defaultId: 1,
    title: "Installation erfolgreich!",
    message: "Das Calculation-Tool wurde erfolgreich installiert. Wollen Sie jetzt die Änderungen einsehen?"});
  if (sResult == 1) shell.openItem(sAddinPath + "VersionLog_DE.txt");

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
  installCalcTool();
  //startInstaller(`CalcToolInstaller_V${app.getVersion()}.EXE`);
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
