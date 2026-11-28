@echo off
rem Installe FXServer et les ressources de base (double-clic)
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer-windows.ps1" %*
pause
