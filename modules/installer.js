const fs = require('fs');
const regedit = require('regedit');
const path = require('path');
const convert = require('xml-js');
const { app, dialog, shell, BrowserWindow } = require('electron');
const child_process = require('child_process');

exports.importSettings = function(sResFolder, sAddinPath, win) {
  dialog.showMessageBox(win, {message: sResFolder + "calctool-settings-import\\CreateXMLFromOldSettings.vbs"});
  var vbs = child_process.spawnSync('cscript.exe', [sResFolder + "calctool-settings-import\\CreateXMLFromOldSettings.vbs", sAddinPath])
  if (vbs.status.valueOf() != 0) {
    dialog.showMessageBox(win, {
      type: "warning",
      title: "Fehler beim Importieren der Einstellungen",
      message: "Beim Kopieren der Einstellungen ist ein Fehler aufgetreten. Die Einstellungen konnten nicht übernommen werden und es wurden die Standardeinstellungen zurückgesetzt.\nPrüfen Sie die Einstellungen auf Änderungen, insbesondere die Funktion DBAutoSync, welche standardmäßig deaktiviert ist.\nResultierender Fehlercode: " + vbs.status.toString()
    });
  }
}

/**
 * Install the Excel-AddIn
 * @param {BrowserWindow} win - The current window for dialogObject and progressBar
 * @param {string} sResFolder - Resource-folder, containing the files to install.
 * @param {Array} registryItems - Array of specific Registry-Items.
 */
exports.installCalcTool = function(win, sResFolder, registryItems) {
  // required vars
  var sAddinPath = app.getPath('appData') + "\\Microsoft\\AddIns\\CalculationTool\\";
  var OldVersion = {raw: "", major: -1, minor: -1, patch: -1};
  var NewVersion = {raw: app.getVersion(), major: -1, minor: -1, patch: -1};
  var sLanguage = "DE";

  // for later use
  if (sResFolder == "") {
    dialog.showMessageBox(win, {
      type: "error",
      title: "Dateien nicht gefunden!",
      message: "Das Zielverzeichnis konnte nicht gefunden werden!"
    });
    return;
  }

  // get version out of registry
  win.setProgressBar(0.1);
  registryItems.forEach(element => {
    if (element.key.includes('CalculationTool\\Version')) {
      OldVersion.raw = element.data.values[''].value;
    }
  });

  // get version out of CalcTool
  if (fs.existsSync(sAddinPath + "CalculationTool.inf")) {
    var fileContent = fs.readFileSync(sAddinPath + "CalculationTool.inf", { flag: 'r'});
    var versionPos = fileContent.toString().search("Version=");
    if (versionPos > 1) {
      versionPos = versionPos + "Version=".length;
      var sVersion = fileContent.toString().substr(versionPos, fileContent.toString().indexOf(".", versionPos + 5) - versionPos);
      OldVersion.raw = sVersion;
    }
  }

  // format versions
  OldVersion = formatVersion(OldVersion);
  NewVersion = formatVersion(NewVersion);

  // disable for old Versions
  if (OldVersion.raw != "" && ((OldVersion.major == 1 && OldVersion.minor == 0) || OldVersion.major == 0)) { 
    var msgResult = dialog.showMessageBox(win, { // for version V1.0.4.0 or older
      type: "warning",
      buttons: ["Abbrechen"],
      title: "Update für ältere Versionen akutell nicht verfügbar!",
      message: "Sie versuchen gerade von V" + OldVersion.raw + ".X auf V" + NewVersion.raw + " zu updaten. Leider ist für ältere Versionen des CalculationTool (V1.0.4.1 oder älter) aktuell noch keine Update-Funktion verfügbar!\nZum Updaten von älteren Installationen haben Sie zwei Optionen:\n- Deinstallieren Sie die alte Version des CalculationTool (Achtung: Alle Einstellungen und Datenbanken werden gelöscht!)\n- Updaten sie ihre bestehende Version auf die Version V1.1.0.0 oder höher"
    });
    return;
  } else { 
    // check version update intensity
    if (OldVersion.raw == "") {
      var msgResult = dialog.showMessageBox(win, {
        type: "info",
        buttons: ["Abbrechen", "Ja"],
        defaultId: 1,
        title: "Installation",
        message: "Wollen Sie das Excel-CalcTool jetzt installieren?"
      });
      if (msgResult != 1) return;
    }
    else if (OldVersion.major < NewVersion.major) {
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
  win.setProgressBar(0.2);
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
  win.setProgressBar(0.3);
  if (!fs.existsSync(sAddinPath)) fs.mkdirSync(sAddinPath);
  if (!fs.existsSync(sAddinPath + "DataBases\\")) fs.mkdirSync(sAddinPath + "DataBases\\");
  if (!fs.existsSync(sAddinPath + "Request\\")) fs.mkdirSync(sAddinPath + "Request\\");
  
  // Delete V1.0.1 or older
  if (fs.existsSync(path.join(sAddinPath, "\\..\\CalculationTool.xlam"))) fs.unlinkSync(path.join(sAddinPath, "\\..\\CalculationTool.xlam"));
  // Delete CalculationTool.inf of V1.0.4.1
  if (fs.existsSync(path.join(sAddinPath, "\\CalculationTool.inf"))) fs.unlinkSync(path.join(sAddinPath, "\\CalculationTool.inf"));

  // Get CalculationTool.xml content for transfering old settings
  var sOldConfig, OldConfig
  if (fs.existsSync(sAddinPath + "CalculationTool.xml")) {
    sOldConfig = fs.readFileSync(sAddinPath + "CalculationTool.xml", { flag: 'r'})
    OldConfig = convert.xml2js(sOldConfig, {compact: false});
    fs.unlinkSync(sAddinPath + "CalculationTool.xml");
  }

  // Copy files
  setTimeout(copyFiles, 1000, sResFolder, sAddinPath, win);

  // Move DataBases
  setTimeout(moveDataBases, 2000, OldVersion, sAddinPath, win);

  // TODO - import from old included settings (V1.0.4.1 or older)
  setTimeout(importSettings, 3000, OldVersion, sResFolder, sAddinPath, win);

  // insert old .xml settings into new settings file
  setTimeout(insertSettings, 4000, sAddinPath, OldConfig, win);

  // final result
  setTimeout(finalResult, 5000, sAddinPath, win)

}

// Further Process functions
function copyFiles(sResFolder, sAddinPath, win) {
  win.setProgressBar(0.4);
  fs.readdirSync(sResFolder + "calctool-files\\").forEach(file => {
    fs.copyFile(sResFolder + "calctool-files\\" + file, sAddinPath + file, function(err) {
      if (err != null && err.code == 'EBUSY') {
        dialog.showMessageBox(win, {
          type: "warning",
          title: "Datei wird von einer anderen Anwendung verwendet",
          message: "Während Excel ausgeführt wird, kann das CalculationTool nicht aktualisiert werden. Beenden Sie Excel und starten die Installation erneut!"});
      } else if (err != null) {
        dialog.showMessageBox(win, {message: "Unerwarter Fehler: " + err.message});
      }
    });
  });
}

function moveDataBases(OldVersion, sAddinPath, win) {
  win.setProgressBar(0.6);
  if (OldVersion.raw == "") {
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
  } else {
    fs.unlinkSync(sAddinPath + "ProductDataBase.xlsx");
    fs.unlinkSync(sAddinPath + "ProjectDataBase.xlsx");
  }
}

function importSettings(OldVersion, sResFolder, sAddinPath, win) {
  win.setProgressBar(0.65);
  if (OldVersion.raw != '') {
    if (OldVersion.major == 1 && (OldVersion.minor == 0 && OldVersion.patch >= 3 || OldVersion.minor == 1 && OldVersion.patch == 0)) {
      var vbs = child_process.spawnSync('cscript.exe', [sResFolder + "calctool-settings-import\\CreateXMLFromOldSettings.vbs", sAddinPath])
      if (vbs.status.valueOf() != 0) {
        dialog.showMessageBox(win, {
          type: "warning",
          title: "Fehler beim Importieren der Einstellungen",
          message: "Beim Kopieren der Einstellungen ist ein Fehler aufgetreten. Die Einstellungen konnten nicht übernommen werden und es wurden die Standardeinstellungen zurückgesetzt.\nPrüfen Sie die Einstellungen auf Änderungen, insbesondere die Funktion DBAutoSync, welche standardmäßig deaktiviert ist.\nResultierender Fehlercode: " + vbs.status.toString()
        });
      }
    }
  }
}

function insertSettings(sAddinPath, OldConfig, win) {
  win.setProgressBar(0.8);
  var sNewConfig = fs.readFileSync(sAddinPath + "CalculationTool.xml", { flag: 'r'});
  var NewConfig = convert.xml2js(sNewConfig, {compact: false});
  copySettings(NewConfig, OldConfig);
  sNewConfig = convert.js2xml(NewConfig, {compact: false})
  fs.writeFileSync(sAddinPath + "CalculationTool.xml", sNewConfig);
}

function finalResult(sAddinPath, win) {
  win.setProgressBar(1);
  var sResult = dialog.showMessageBox(win, {
    type: "info",
    buttons: ["Nein", "Ja"],
    defaultId: 1,
    title: "Installation erfolgreich!",
    message: "Das Calculation-Tool wurde erfolgreich installiert. Wollen Sie jetzt die Änderungen einsehen?"});
  if (sResult == 1) shell.openItem(sAddinPath + "VersionLog_DE.txt");
  win.setProgressBar(0);
}

// other functions
function formatVersion(versionObj) {
    var versionSub = new String(versionObj.raw);
    versionObj.major = Number(versionSub.substring(1, versionSub.search('.')));
    versionSub = versionSub.substring(versionSub.search('.') + 2, versionSub.length);
    versionObj.minor = Number(versionSub.substring(1, versionSub.search('.')));
    versionSub = versionSub.substring(versionSub.search('.') + 2, versionSub.length);
    versionObj.patch = Number(versionSub.substring(1, versionSub.search('.')));
    return versionObj;
}

function copySettings(NewSettings, OldSettings) {
    for (x in NewSettings.elements) {
        for (k in OldSettings.elements) { 
            // value
            if (NewSettings.elements[x].name == "value" && OldSettings.elements[k].name == "value") {
                if (NewSettings.elements[x].attributes.index == OldSettings.elements[k].attributes.index) {
                    if (NewSettings.elements[x].hasOwnProperty("elements") && OldSettings.elements[k].hasOwnProperty("elements")) {
                        NewSettings.elements[x].elements[0].text = OldSettings.elements[k].elements[0].text;
                        break;
                    }
                }
            }
            // type with id
            else if (NewSettings.elements[x].hasOwnProperty("attributes") && OldSettings.elements[k].hasOwnProperty("attributes")) {
                if (NewSettings.elements[x].attributes.hasOwnProperty("id") && OldSettings.elements[k].attributes.hasOwnProperty("id")) {
                    if (NewSettings.elements[x].attributes.id == OldSettings.elements[k].attributes.id) {
                        copySettings(NewSettings.elements[x], OldSettings.elements[k]);
                        break;
                    }
                }
            }
            else if (NewSettings.elements[x].name == OldSettings.elements[k].name) {
                if (NewSettings.elements[x].hasOwnProperty("elements") && NewSettings.elements[k].hasOwnProperty("elements")) {
                    copySettings(NewSettings.elements[x], OldSettings.elements[k]);
                    break;
                }
            }
        }
    }
}