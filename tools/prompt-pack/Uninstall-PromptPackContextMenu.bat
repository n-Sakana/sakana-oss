@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul

cd /d "%~dp0"
if errorlevel 1 goto PromptPackUninstallCdFailed

set "SCRIPT=%~dp0Uninstall-PromptPackContextMenu.ps1"

echo PromptPack Context Menu Uninstaller
echo -----------------------------------
echo Tool directory:
echo "%~dp0"
echo.

if not exist "%SCRIPT%" goto PromptPackUninstallMissingScript

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"
goto PromptPackUninstallDone

:PromptPackUninstallCdFailed
echo ERROR: Failed to enter the PromptPack directory.
echo Path: "%~dp0"
echo.
set "EXIT_CODE=1"
goto PromptPackUninstallDonePause

:PromptPackUninstallMissingScript
echo ERROR: Uninstall-PromptPackContextMenu.ps1 was not found.
echo Path: "%SCRIPT%"
echo.
echo This uninstaller removes fixed HKCU registry keys.
echo If this file is missing, use the manual reg delete commands in README.md.
echo.
set "EXIT_CODE=1"
goto PromptPackUninstallDonePause

:PromptPackUninstallDone
echo.
if "%EXIT_CODE%"=="0" goto PromptPackUninstallSuccess
goto PromptPackUninstallFailure

:PromptPackUninstallSuccess
echo SUCCESS: PromptPack context menu removed.
goto PromptPackUninstallDonePause

:PromptPackUninstallFailure
echo ERROR: Uninstaller failed with exit code %EXIT_CODE%.
goto PromptPackUninstallDonePause

:PromptPackUninstallDonePause
if /I "%PROMPTPACK_NO_PAUSE%"=="1" exit /b %EXIT_CODE%

echo.
echo This window is kept open so the result is visible.
echo You can close it after reading the message above.
echo.
pause
exit /b %EXIT_CODE%
