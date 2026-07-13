@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Build-Addin.ps1" -OutputFormat all
exit /b %ERRORLEVEL%

