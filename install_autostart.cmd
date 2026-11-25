@echo off
rem Creates the "Palladium Auto Bot" scheduled task: start bot at user logon.
rem ASCII only: cmd.exe misreads UTF-8 batch files.
chcp 65001 >nul
setlocal

set "TARGET=%~dp0run_bot.cmd"
if not exist "%TARGET%" (
    echo ERROR: run_bot.cmd not found next to this script.
    exit /b 1
)

echo Target: %TARGET%
schtasks /Create /TN "Palladium Auto Bot" /TR "\"%TARGET%\"" /SC ONLOGON /F
if errorlevel 1 (
    echo.
    echo Failed. Try running this file as Administrator.
    exit /b 1
)

echo.
echo Task created. It starts the bot at next logon.
echo Check:  schtasks /Query /TN "Palladium Auto Bot"
echo Remove: schtasks /Delete /TN "Palladium Auto Bot" /F
echo.
echo To start it right now (only if the bot is NOT already running):
echo         schtasks /Run /TN "Palladium Auto Bot"
if defined BOT_NO_PAUSE goto end
pause >nul
:end
endlocal
