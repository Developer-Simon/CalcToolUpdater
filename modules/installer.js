const fs = require('fs');
const regedit = require('regedit');
const path = require('path');
const convert = require('xml-js');
const { app, dialog } = require('electron');

/**
 * Install the Excel-AddIn
 * @param {Window} win - The current window for dialogObject
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
  registryItems.forEach(element => {
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
  
  // file actions (sync)
  // Create Folders
  if (!fs.existsSync(sAddinPath)) fs.mkdirSync(sAddinPath);
  if (!fs.existsSync(sAddinPath + "DataBases\\")) fs.mkdirSync(sAddinPath + "DataBases\\");
  if (!fs.existsSync(sAddinPath + "Request\\")) fs.mkdirSync(sAddinPath + "Request\\");
  
  // Delete V1.0.1 or older
  if (fs.existsSync(path.join(sAddinPath, "\\..\\CalculationTool.xlam"))) fs.unlinkSync(path.join(sAddinPath, "\\..\\CalculationTool.xlam"));
  // Delete CalculationTool.inf of V1.0.4.1
  if (fs.existsSync(path.join(sAddinPath, "\\CalculationTool.inf"))) fs.unlinkSync(path.join(sAddinPath, "\\CalculationTool.inf"));

  // file actions (async)

  // Get CalculationTool.xml content for transfering old settings
  var sOldConfig = fs.readFileSync(sAddinPath + "CalculationTool.xml", { flag: 'r'})
  var OldConfig = convert.xml2js(sOldConfig, {compact: false});
  if (fs.existsSync(sAddinPath + "CalculationTool.xml")) fs.unlinkSync(sAddinPath + "CalculationTool.xml");

  // Copy files
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

  // Move DataBases
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

  // TODO - import from old included settings (V1.0.4.1 or older)

  // insert old .xml settings into new settings file
  var sNewConfig = fs.readFileSync(sAddinPath + "CalculationTool.xml", { flag: 'r'});
  var NewConfig = convert.xml2js(sNewConfig, {compact: false});
  copySettings(NewConfig, OldConfig);
  sNewConfig = convert.js2xml(NewConfig, {compact: false})
  fs.writeFileSync(sAddinPath + "CalculationTool.xml", sNewConfig);

  var sResult = dialog.showMessageBox(win, {
    type: "info",
    buttons: ["Nein", "Ja"],
    defaultId: 1,
    title: "Installation erfolgreich!",
    message: "Das Calculation-Tool wurde erfolgreich installiert. Wollen Sie jetzt die Änderungen einsehen?"});
  if (sResult == 1) shell.openItem(sAddinPath + "VersionLog_DE.txt");

}

function GetOldSettings(sAddinPath) {

}

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