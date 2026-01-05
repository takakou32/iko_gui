@echo off
chcp 932 > nul
powershell.exe -ExecutionPolicy Bypass -File "%~dp0Main.ps1"
pause
