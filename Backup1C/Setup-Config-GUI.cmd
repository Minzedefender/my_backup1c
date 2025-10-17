@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo [INFO] Запуск графического мастера настройки системы резервного копирования 1С...

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "setup\SetupWizard.GUI.ps1"

if %ERRORLEVEL% neq 0 (
    echo [ERROR] Произошла ошибка при выполнении мастера настройки
    pause
) else (
    echo [INFO] Мастер настройки завершен успешно
    timeout /t 2 > nul
)