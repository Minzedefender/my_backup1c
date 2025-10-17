@echo off
chcp 65001 >NUL
cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "setup\ConfigEditor.GUI.ps1"

if %ERRORLEVEL% NEQ 0 (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "setup\ConfigEditor.ps1"
)

pause