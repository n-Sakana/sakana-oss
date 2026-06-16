@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%Export-RepoText.ps1"
set "FAILED=0"

if not exist "%PS1%" (
    echo ERROR: Export-RepoText.ps1 not found:
    echo %PS1%
    pause
    exit /b 1
)

if "%~1"=="" (
    echo Drag and drop a repository folder onto this BAT.
    echo.
    pause
    exit /b 1
)

:LOOP
if "%~1"=="" goto DONE

call :EXPORT_ONE "%~1"
shift
goto LOOP

:EXPORT_ONE
set "ROOT=%~1"

if not exist "%ROOT%" (
    echo ERROR: Path not found: %ROOT%
    set "FAILED=1"
    exit /b 0
)

rem If a file is dropped, export its parent folder.
if not exist "%ROOT%\*" (
    for %%I in ("%ROOT%") do set "ROOT=%%~dpI"
)

for %%I in ("%ROOT%") do set "ROOT=%%~fI"

echo.
echo ========================================
echo Exporting:
echo %ROOT%
echo ========================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Root "%ROOT%"

if errorlevel 1 (
    echo FAILED: %ROOT%
    set "FAILED=1"
) else (
    echo DONE: %ROOT%
)

exit /b 0

:DONE
echo.
if "%FAILED%"=="0" (
    echo All exports completed.
) else (
    echo Completed with errors.
)

pause
exit /b %FAILED%
