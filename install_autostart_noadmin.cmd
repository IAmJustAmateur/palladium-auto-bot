@echo off
rem Autostart WITHOUT administrator rights: puts a hidden launcher into the
rem user's Startup folder. Remove it to disable autostart.
rem ASCII only: cmd.exe misreads UTF-8 batch files.
chcp 65001 >nul
setlocal

set "TARGET=%~dp0run_bot.cmd"
if not exist "%TARGET%" (
    echo ERROR: run_bot.cmd not found next to this script.
    exit /b 1
)

set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS=%STARTUP%\PalladiumAutoBot.vbs"

> "%VBS%" echo CreateObject("WScript.Shell").Run """%TARGET%"" hidden", 0, False

if not exist "%VBS%" (
    echo ERROR: could not write %VBS%
    exit /b 1
)

echo Autostart installed.
echo   Launcher: %VBS%
echo   Target:   %TARGET%
echo.
echo The bot will start hidden at next logon.
echo To stop autostart: delete the file above.
if defined BOT_NO_PAUSE goto end
pause >nul
:end
endlocal
