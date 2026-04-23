@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File ".\NewAssetTool.Wpf.ps1"
echo.
echo (Done) Press any key to close...
pause >nul
endlocal
