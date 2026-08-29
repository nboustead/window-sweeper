@echo off
title Window Sweeper - Add to Start Menu
rem Creates a Start Menu shortcut for Window Sweeper, so you can find it
rem by pressing the Windows key and typing "Window Sweeper" - just like
rem any other installed program, instead of digging through this folder.
rem Safe to run more than once; it just replaces the shortcut each time.
cd /d "%~dp0"
set "WS_DIR=%~dp0"
powershell -NoProfile -Command "$ErrorActionPreference = 'Stop'; try { $folder = $env:WS_DIR; $psPath = Join-Path $folder 'WindowSweeper.ps1'; $icoPath = Join-Path $folder 'sweeper.ico'; if (-not (Test-Path -LiteralPath $psPath)) { throw ('WindowSweeper.ps1 not found in ' + $folder) }; if (-not (Test-Path -LiteralPath $icoPath)) { throw ('sweeper.ico not found in ' + $folder) }; $programsDir = [Environment]::GetFolderPath('Programs'); $lnkPath = Join-Path $programsDir 'Window Sweeper.lnk'; $ws = New-Object -ComObject WScript.Shell; $lnk = $ws.CreateShortcut($lnkPath); $lnk.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'; $q = [char]34; $lnk.Arguments = '-STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ' + $q + $psPath + $q; $lnk.IconLocation = $icoPath + ',0'; $lnk.WorkingDirectory = $folder; $lnk.WindowStyle = 7; $lnk.Description = 'Sweep open Excel, Word, Outlook and Chrome windows'; $lnk.Save(); exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
    echo.
    echo Something went wrong - the shortcut was not created.
    echo.
    pause
    exit /b 1
)
echo.
echo Done - press the Windows key and type "Window Sweeper"
echo.
pause
