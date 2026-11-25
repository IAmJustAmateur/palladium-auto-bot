@echo off
rem Launch the Telegram bot. Restarts it after a crash; close the window to stop.
rem Usage: run_bot.cmd [hidden]
rem   hidden - no pause at the end (used by the autostart launcher).
rem Exit code 1 means a permanent problem (no token, no key, already running) - do not loop.
rem ASCII only: cmd.exe misreads UTF-8 batch files.
chcp 65001 >nul
setlocal
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
if /i "%~1"=="hidden" set "BOT_NO_PAUSE=1"

set "BUNDLED=%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
if exist "%BUNDLED%" (
    set "PY_CMD="%BUNDLED%""
    goto loop
)

where py >nul 2>nul
if not errorlevel 1 (
    set "PY_CMD=py -3"
    goto loop
)

set "PY_CMD=python"

:loop
%PY_CMD% "%~dp0tg_bot.py"
set "RC=%ERRORLEVEL%"
if "%RC%"=="1" goto end
echo.
echo Bot stopped (code %RC%). Restart in 10 seconds. Press Ctrl+C twice to abort.
timeout /t 10 /nobreak >nul
goto loop

:end
echo.
echo Bot finished with code %RC%. See logs\tg_bot.log for details.
if defined BOT_NO_PAUSE goto quit
pause >nul
:quit
endlocal & exit /b %RC%
