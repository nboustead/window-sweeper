@echo off
title Window Sweeper - Add to Start Menu
rem Creates a Start Menu shortcut for Window Sweeper, so you can find it
rem by pressing the Windows key and typing "Window Sweeper" - just like
rem any other installed program, instead of digging through this folder.
rem Safe to run more than once; it just replaces the shortcut each time.
cd /d "%~dp0"
powershell -NoProfile -Command "$folder = '%~dp0'; $ws = New-Object -ComObject WScript.Shell; $lnk = $ws.CreateShortcut((Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Window Sweeper.lnk')); $lnk.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'; $q = [char]34; $lnk.Arguments = '-STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ' + $q + $folder + 'WindowSweeper.ps1' + $q; $lnk.IconLocation = $folder + 'sweeper.ico,0'; $lnk.WorkingDirectory = $folder; $lnk.WindowStyle = 7; $lnk.Description = 'Sweep open Excel, Word, Outlook and Chrome windows'; $lnk.Save()"
echo.
echo Done - press the Windows key and type "Window Sweeper"
echo.
pause
