@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Install-Addin.ps1"
set "result=%ERRORLEVEL%"
if not "%result%"=="0" (
  echo.
  echo xltoolrack setup failed.
  pause
)
exit /b %result%
