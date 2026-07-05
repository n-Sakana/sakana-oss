@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul

cd /d "%~dp0"
if errorlevel 1 goto PromptPackInstallCdFailed

set "SCRIPT=%~dp0Install-PromptPackContextMenu.ps1"

echo PromptPack Context Menu Installer
echo ---------------------------------
echo Tool directory:
echo "%~dp0"
echo.

if not exist "%SCRIPT%" goto PromptPackInstallMissingScript

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"
goto PromptPackInstallDone

:PromptPackInstallCdFailed
echo ERROR: Failed to enter the PromptPack directory.
echo Path: "%~dp0"
echo.
set "EXIT_CODE=1"
goto PromptPackInstallDonePause

:PromptPackInstallMissingScript
echo ERROR: Install-PromptPackContextMenu.ps1 was not found.
echo Path: "%SCRIPT%"
echo.
echo Make sure all PromptPack files are in the same folder.
echo.
set "EXIT_CODE=1"
goto PromptPackInstallDonePause

:PromptPackInstallDone
echo.
if "%EXIT_CODE%"=="0" goto PromptPackInstallSuccess
goto PromptPackInstallFailure

:PromptPackInstallSuccess
echo SUCCESS: PromptPack context menu installed.
goto PromptPackInstallDonePause

:PromptPackInstallFailure
echo ERROR: Installer failed with exit code %EXIT_CODE%.
goto PromptPackInstallDonePause

:PromptPackInstallDonePause
if /I "%PROMPTPACK_NO_PAUSE%"=="1" exit /b %EXIT_CODE%

echo.
echo This window is kept open so the result is visible.
echo You can close it after reading the message above.
echo.
pause
exit /b %EXIT_CODE%
