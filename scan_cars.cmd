@echo off
rem Run scan_cars.py from any working directory.
rem Usage: scan_cars.cmd "C:\path\to\photos" --dry-run
rem Or drag a folder/file onto this file.
rem Keep this file ASCII-only: cmd.exe misreads UTF-8 batch files.
chcp 65001 >nul
setlocal
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"

set "BUNDLED=%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
if exist "%BUNDLED%" (
    set "PY_CMD="%BUNDLED%""
    goto run
)

where py >nul 2>nul
if not errorlevel 1 (
    set "PY_CMD=py -3"
    goto run
)

set "PY_CMD=python"

:run
%PY_CMD% "%~dp0scan_cars.py" %*
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" echo [error] exit code %RC%
echo Log file: %~dp0logs\scan_cars.log
if defined SCAN_CARS_NO_PAUSE goto end
echo.
echo Press any key to close this window...
pause >nul
:end
endlocal & exit /b %RC%
