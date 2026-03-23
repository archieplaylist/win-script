@echo off
:: This tiny script forces PowerShell to bypass the strict execution policy 
:: and run your app smoothly without showing a black console window.
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0WingetUI.ps1"