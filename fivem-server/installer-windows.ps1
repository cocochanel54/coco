# Installe (ou met à jour) FXServer et les ressources de base de FiveM, pour Windows.
# Le plus simple : double-cliquer sur installer-windows.bat
param(
    [ValidateSet('recommended', 'latest')]
    [string] $Canal = 'recommended'
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Dossier = $PSScriptRoot
$Tmp = Join-Path $Dossier '.installation'
$Artifacts = Join-Path $Dossier 'artifacts'
# Les crochets de [cfx] gênent PowerShell : on passe par .NET pour les dossiers
$Cfx = Join-Path $Dossier 'resources\[cfx]'

function Supprimer([string] $chemin) {
    if ([IO.Directory]::Exists($chemin)) { [IO.Directory]::Delete($chemin, $true) }
}

try {
    Supprimer $Tmp
    [void][IO.Directory]::CreateDirectory($Tmp)

    Write-Host "==> Recherche de la version « $Canal » de FXServer..."
    $infos = Invoke-RestMethod -UseBasicParsing -Uri 'https://changelogs-live.fivem.net/api/changelog/versions/win32/server'
    $url = $infos."$($Canal)_download"
    $version = $infos.$Canal
    if (-not $url) { throw 'Lien de téléchargement de FXServer introuvable.' }

    Write-Host "==> Téléchargement de FXServer $version (patiente un peu)..."
    $archive = Join-Path $Tmp 'server.7z'
    Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $archive

    Write-Host '==> Décompression...'
    Supprimer $Artifacts
    [void][IO.Directory]::CreateDirectory($Artifacts)
    $septZip = @("$env:ProgramFiles\7-Zip\7z.exe", "${env:ProgramFiles(x86)}\7-Zip\7z.exe") |
        Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if ($septZip) {
        & $septZip x $archive "-o$Artifacts" -y | Out-Null
    } else {
        # Windows 11 sait ouvrir les .7z avec tar ; sur Windows 10 il faut 7-Zip
        & tar.exe -xf $archive -C $Artifacts
    }
    if (-not (Test-Path -LiteralPath (Join-Path $Artifacts 'FXServer.exe'))) {
        throw "Impossible de décompresser server.7z. Installe 7-Zip (https://www.7-zip.org) puis relance installer-windows.bat."
    }

    Write-Host '==> Téléchargement des ressources de base (chat, spawnmanager...)...'
    $zip = Join-Path $Tmp 'cfx.zip'
    Invoke-WebRequest -UseBasicParsing -Uri 'https://github.com/citizenfx/cfx-server-data/archive/refs/heads/master.zip' -OutFile $zip
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::ExtractToDirectory($zip, $Tmp)
    Supprimer $Cfx
    [IO.Directory]::Move((Join-Path $Tmp 'cfx-server-data-master\resources'), $Cfx)

    $secrets = Join-Path $Dossier 'secrets.cfg'
    if (-not [IO.File]::Exists($secrets)) {
        [IO.File]::Copy((Join-Path $Dossier 'secrets.cfg.example'), $secrets)
    }

    Write-Host ''
    Write-Host 'Installation terminée !' -ForegroundColor Green
    if ((Get-Content -LiteralPath $secrets -Raw) -match 'colle_ta_cle_ici') {
        Write-Host 'Il reste à mettre ta clé de licence dans secrets.cfg (voir README.md), puis à lancer demarrer-windows.bat.'
    } else {
        Write-Host 'Lance le serveur avec demarrer-windows.bat.'
    }
}
catch {
    Write-Host ''
    Write-Host "Erreur : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Tu peux aussi installer à la main : voir README.md, partie « Installation manuelle ».'
    exit 1
}
finally {
    Supprimer $Tmp
}
