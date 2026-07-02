@echo off
setlocal EnableExtensions

chcp 65001 >nul

set "SCRIPT=%~dp0PromptPack.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PromptPack.ps1 not found.
    echo Path: %SCRIPT%
    echo.
    pause
    exit /b 1
)

if "%~1"=="" (
    echo PromptPack
    echo ----------
    echo Drag and drop files or folders onto this BAT file.
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo PromptPack completed successfully.
) else (
    echo PromptPack completed with errors.
)

pause
exit /b %EXIT_CODE%
