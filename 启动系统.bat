@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ============================
echo   AI排单系统启动中...
echo ============================
echo.
powershell -ExecutionPolicy Bypass -File "watchdog.ps1"
pause
