@echo off
chcp 65001 >nul
setlocal
set SCRIPT_DIR=%~dp0
set PS_SCRIPT=%SCRIPT_DIR%Run-Backup.ps1
set TEMP_LOG=%TEMP%\backup_result_%RANDOM%.tmp

echo.
echo [INFO] Запуск процесса резервного копирования...
echo.

:: Запускаем PowerShell и сохраняем код возврата
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ^
  "& '%PS_SCRIPT%'; exit $LASTEXITCODE" 2>&1

:: Сохраняем код возврата
set EXIT_CODE=%ERRORLEVEL%

:: Проверяем результат
if %EXIT_CODE% EQU 0 (
    echo.
    echo [SUCCESS] Резервное копирование завершено успешно.
    echo Окно закроется через 5 секунд...
    timeout /t 5 >nul
    exit /b 0
) else (
    echo.
    echo [ERROR] Резервное копирование завершено с ошибками ^(код: %EXIT_CODE%^)
    echo.
    echo Нажмите любую клавишу для закрытия окна...
    pause >nul
    exit /b %EXIT_CODE%
)