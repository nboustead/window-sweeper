@echo off
title Window Sweeper - Remove from Start Menu
rem Removes the Start Menu shortcut created by "Add to Start Menu.bat".
rem This folder and WindowSweeper.ps1 itself are left alone - only the
rem shortcut goes away.
cd /d "%~dp0"
powershell -NoProfile -Command "$ErrorActionPreference = 'Stop'; try { $programsDir = [Environment]::GetFolderPath('Programs'); $lnkPath = Join-Path $programsDir 'Window Sweeper.lnk'; if (Test-Path -LiteralPath $lnkPath) { Remove-Item -LiteralPath $lnkPath -Force; Write-Host 'Shortcut removed.' } else { Write-Host 'No Start Menu shortcut was found - nothing to remove.' }; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
    echo.
    echo Something went wrong - the shortcut was not removed.
    echo.
    pause
    exit /b 1
)
echo.
echo Done - Window Sweeper is no longer in the Start Menu.
echo.
pause
