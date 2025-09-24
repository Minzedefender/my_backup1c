@echo off
chcp 65001 >NUL
setlocal
pushd "%~dp0"

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SCRIPT_GUI=setup\ConfigEditor.GUI.ps1"
set "SCRIPT_CLI=setup\ConfigEditor.ps1"

:: Проверяем наличие GUI версии
if exist "%SCRIPT_GUI%" (
    echo [INFO] Запуск графического редактора конфигураций...
    "%PS%" -NoLogo -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT_GUI%"
) else if exist "%SCRIPT_CLI%" (
    echo [INFO] GUI версия не найдена, запуск консольного редактора...
    "%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_CLI%"
) else (
    echo [ERROR] Редактор конфигураций не найден
    echo.
    pause
    exit /b 1
)

popd
endlocal