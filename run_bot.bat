@echo off
cd /d "%~dp0"
title Receipt Tax Bot

:loop
python bot.py

rem Exit codes that mean "do not restart":
rem   0 = you stopped it (Ctrl+C),  3 = already running,  4 = no token set
if "%errorlevel%"=="0" goto end
if "%errorlevel%"=="3" goto end
if "%errorlevel%"=="4" goto end

echo.
echo The bot stopped unexpectedly. Restarting in 10 seconds...
echo (Close this window to stop it for good.)
timeout /t 10 /nobreak >nul
goto loop

:end
echo.
pause
