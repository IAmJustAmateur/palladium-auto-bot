@echo off
rem Launch the Telegram bot. Restarts it after a crash; close the window to stop.
rem ASCII only: cmd.exe misreads UTF-8 batch files.
chcp 65001 >nul
setlocal
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"

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
%PY_CMD% "%~dp0tg_bot.py" %*
echo.
echo Bot stopped, exit code %ERRORLEVEL%. Restart in 10 seconds (Ctrl+C to abort).
timeout /t 10 /nobreak >nul
goto loop
