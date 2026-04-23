@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0NewAssetTool.Wpf.matching.v14.ps1" -XamlPath "%~dp0NewAssetTool.matching.v14.xaml"
