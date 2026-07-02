@echo off
setlocal EnableExtensions

set "SCRIPT=%~dp0Install-PromptPackContextMenu.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: Install-PromptPackContextMenu.ps1 not found.
    echo Path: %SCRIPT%
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
pause
exit /b %EXIT_CODE%
