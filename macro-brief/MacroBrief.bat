@echo off
REM Lance le briefing macro du jour (double-clic)
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0MacroBrief.ps1" %*
echo.
pause
