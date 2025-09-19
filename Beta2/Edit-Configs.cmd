@echo off
chcp 65001 >NUL
setlocal
pushd "%~dp0"

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SCRIPT=setup\ConfigEditor.ps1"

if not exist "%SCRIPT%" (
  echo [ERROR] %SCRIPT% not found
  echo.
  pause
  exit /b 1
)

echo [INFO] Launching config editor...
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "EC=%ERRORLEVEL%"

echo.
if not "%EC%"=="0" (
  echo [ERROR] Editor exited with code %EC%.
) else (
  echo [INFO] Done.
)

echo.
pause
popd
endlocal
