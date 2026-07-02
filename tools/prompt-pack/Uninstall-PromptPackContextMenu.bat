@echo off
setlocal EnableExtensions

if /I not "%PROMPTPACK_NO_RELAUNCH%"=="1" (
    if /I not "%~1"=="--inner" (
        start "PromptPack Uninstaller" "%ComSpec%" /d /k call "%~f0" --inner
        exit /b 0
    )
)

cd /d "%~dp0" || (
    echo ERROR: Failed to enter the PromptPack directory.
    echo Path: %~dp0
    echo.
    pause
    exit /b 1
)

set "SCRIPT=%~dp0Uninstall-PromptPackContextMenu.ps1"

echo PromptPack Context Menu Uninstaller
echo -----------------------------------
echo Tool directory:
echo %~dp0
echo.

if not exist "%SCRIPT%" (
    echo ERROR: Uninstall-PromptPackContextMenu.ps1 was not found.
    echo Path: %SCRIPT%
    echo.
    echo This uninstaller removes fixed HKCU registry keys.
    echo If this file is missing, use the manual reg delete commands in README.md.
    echo.
    pause
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo SUCCESS: PromptPack context menu removed.
) else (
    echo ERROR: Uninstaller failed with exit code %EXIT_CODE%.
)
echo.
echo This window is kept open so the result is visible.
echo You can close it after reading the message above.
echo.
pause
exit /b %EXIT_CODE%
