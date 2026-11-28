@echo off
rem Demarre le panneau web txAdmin (il s'ouvre dans le navigateur)
cd /d "%~dp0"

if not exist "artifacts\FXServer.exe" (
  echo FXServer n'est pas installe : lance d'abord installer-windows.bat
  pause
  exit /b 1
)

echo Dans txAdmin, choisis "Existing Server Data" et indique ce dossier :
echo %~dp0
"artifacts\FXServer.exe"
pause
