@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0test\Run-All.ps1"
exit /b %ERRORLEVEL%
