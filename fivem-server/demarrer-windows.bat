@echo off
rem Demarre le serveur FiveM (double-clic)
cd /d "%~dp0"

if not exist "artifacts\FXServer.exe" (
  echo FXServer n'est pas installe : lance d'abord installer-windows.bat
  pause
  exit /b 1
)

if not exist "secrets.cfg" copy "secrets.cfg.example" "secrets.cfg" >nul

findstr /c:"colle_ta_cle_ici" "secrets.cfg" >nul
if not errorlevel 1 (
  echo Mets d'abord ta cle de licence dans secrets.cfg ^(voir README.md^).
  start "" notepad "secrets.cfg"
  pause
  exit /b 1
)

"artifacts\FXServer.exe" +set onesync on +exec server.cfg
pause
