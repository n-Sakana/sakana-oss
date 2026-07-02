@echo off
setlocal EnableExtensions

if /I "%~1"=="--inner" (
    shift /1
) else (
    if /I not "%PROMPTPACK_NO_RELAUNCH%"=="1" (
        start "PromptPack" "%ComSpec%" /d /k call "%~f0" --inner %*
        exit /b 0
    )
)

chcp 65001 >nul
cd /d "%~dp0" || (
    echo ERROR: Failed to enter the PromptPack directory.
    echo Path: %~dp0
    echo.
    pause
    exit /b 1
)

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
    echo You can also run it from the installed context menu.
    echo.
    pause
    exit /b 1
)

set ARGS=
:PromptPackBuildArgs
if "%~1"=="" goto PromptPackArgsDone
set ARGS=%ARGS% "%~1"
shift /1
goto PromptPackBuildArgs
:PromptPackArgsDone

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %ARGS%
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo PromptPack completed successfully.
) else (
    echo PromptPack completed with errors.
)
echo.
echo This window is kept open so the result is visible.
echo You can close it after reading the message above.
echo.
pause
exit /b %EXIT_CODE%
