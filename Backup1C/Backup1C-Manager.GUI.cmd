@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo [INFO] Запуск менеджера системы резервного копирования 1С...

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Backup1C-Manager.GUI.ps1"

if %ERRORLEVEL% neq 0 (
    echo [ERROR] Произошла ошибка при выполнении менеджера
    pause
) else (
    echo [INFO] Менеджер завершен успешно
    timeout /t 2 > nul
)