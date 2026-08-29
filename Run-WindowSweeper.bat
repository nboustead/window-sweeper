@echo off
title Window Sweeper
rem Window Sweeper - shows every open Excel workbook and Outlook message
rem window in one list, so you can save-and-close or just close each one
rem without answering a "Save changes?" prompt by hand every time.
rem Double-click this file (or a Desktop shortcut to it) to open the tool.
rem See README.txt in this folder for details.
cd /d "%~dp0"
start "" /min powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0WindowSweeper.ps1"
exit
