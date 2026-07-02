@echo off
setlocal EnableExtensions

if /I not "%PROMPTPACK_NO_RELAUNCH%"=="1" (
    if /I not "%~1"=="--inner" (
        start "PromptPack Installer" "%ComSpec%" /d /k call "%~f0" --inner
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

set "SCRIPT=%~dp0Install-PromptPackContextMenu.ps1"

echo PromptPack Context Menu Installer
echo ---------------------------------
echo Tool directory:
echo %~dp0
echo.

if not exist "%SCRIPT%" (
    echo ERROR: Install-PromptPackContextMenu.ps1 was not found.
    echo Path: %SCRIPT%
    echo.
    echo Make sure all PromptPack files are in the same folder.
    echo.
    pause
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo SUCCESS: PromptPack context menu installed.
) else (
    echo ERROR: Installer failed with exit code %EXIT_CODE%.
)
echo.
echo This window is kept open so the result is visible.
echo You can close it after reading the message above.
echo.
pause
exit /b %EXIT_CODE%
