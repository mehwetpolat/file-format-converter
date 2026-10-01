@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
if errorlevel 1 (
    echo Kaldirma tamamlanamadi.
    pause
    exit /b 1
)
pause
