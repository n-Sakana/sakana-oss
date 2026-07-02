@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul

cd /d "%~dp0"
if errorlevel 1 goto PromptPackCdFailed

set "SCRIPT=%~dp0PromptPack.ps1"
if not exist "%SCRIPT%" goto PromptPackMissingScript

if "%~1"=="" goto PromptPackNoInput

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "EXIT_CODE=%ERRORLEVEL%"
goto PromptPackDone

:PromptPackCdFailed
echo ERROR: Failed to enter the PromptPack directory.
echo Path: "%~dp0"
echo.
set "EXIT_CODE=1"
goto PromptPackDonePause

:PromptPackMissingScript
echo ERROR: PromptPack.ps1 was not found.
echo Path: "%SCRIPT%"
echo.
set "EXIT_CODE=1"
goto PromptPackDonePause

:PromptPackNoInput
echo PromptPack
echo ----------
echo Drag and drop files or folders onto this BAT file.
echo You can also run it from the installed context menu.
echo.
set "EXIT_CODE=1"
goto PromptPackDonePause

:PromptPackDone
echo.
if "%EXIT_CODE%"=="0" goto PromptPackSuccess
goto PromptPackFailure

:PromptPackSuccess
echo PromptPack completed successfully.
goto PromptPackDonePause

:PromptPackFailure
echo PromptPack completed with errors.
goto PromptPackDonePause

:PromptPackDonePause
if /I "%PROMPTPACK_NO_PAUSE%"=="1" exit /b %EXIT_CODE%

echo.
echo This window is kept open so the result is visible.
echo You can close it after reading the message above.
echo.
pause
exit /b %EXIT_CODE%
