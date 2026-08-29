@echo off
title Window Sweeper - Remove from Start Menu
rem Removes the Start Menu shortcut created by "Add to Start Menu.bat".
rem This folder and WindowSweeper.ps1 itself are left alone - only the
rem shortcut goes away.
cd /d "%~dp0"
powershell -NoProfile -Command "$lnkPath = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Window Sweeper.lnk'; if (Test-Path $lnkPath) { Remove-Item -Path $lnkPath -Force; Write-Host 'Shortcut removed.' } else { Write-Host 'No Start Menu shortcut was found - nothing to remove.' }"
echo.
echo Done - Window Sweeper is no longer in the Start Menu.
echo.
pause
