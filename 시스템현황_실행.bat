@echo off
title SystemStatus Updater
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-SystemStatus.ps1"
echo.
pause
