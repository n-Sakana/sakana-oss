@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Install-Addin.ps1" -Uninstall
set "result=%ERRORLEVEL%"
if not "%result%"=="0" (
  echo.
  echo xltoolrack uninstall failed.
  pause
)
exit /b %result%
