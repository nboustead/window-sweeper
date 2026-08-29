# v0.4.1 Start Menu helper scripts - post-fix test results

All tests executed 2026-08-29 against the exact shipped artifacts (commit 39fe7a3,
zip SHA256 3B938C616923FE358E831BA6060EA63281B6739D6E32897FF946081B51A9868A),
covering the hardened logic in `Add to Start Menu.bat` (path-as-data via WS_DIR
env var, ErrorActionPreference Stop + try/catch + errorlevel branch, -LiteralPath,
GetFolderPath('Programs')) and `Remove from Start Menu.bat` (same error handling,
-LiteralPath removal).

## A. Success cycle (repo folder)

1. `Add to Start Menu.bat` run from the repo working copy: exit 0, success line
   printed. Created `.lnk` inspected: TargetPath = powershell.exe,
   Arguments -File path, IconLocation, and WorkingDirectory all resolve to the
   repo folder; WindowStyle 7; description set.
2. Launched via the `.lnk`: window "Window Sweeper v0.4.1" appeared; the live
   process `Win32_Process.CommandLine` verified the `-File` argument pointed at
   the repo copy (checked twice, before and at close). Closed via
   `CloseMainWindow()` - clean exit in 2.4s.
3. `Remove from Start Menu.bat`: exit 0, `.lnk` confirmed gone via
   `Test-Path -LiteralPath`.

## B. Apostrophe-containing install path (the round-1 Major)

1. `WindowSweeper-v0.4.1.zip` extracted into a folder literally named
   `O'Brien Test`.
2. `Add to Start Menu.bat` run from that folder: exit 0. The created `.lnk`
   Arguments embed the apostrophe path verbatim as data:
   `-File "...\O'Brien Test\WindowSweeper.ps1"` - no parse error, no
   command alteration.
3. Launched via that `.lnk`: the live process CommandLine confirmed execution
   from the apostrophe path; clean close.
4. `Remove from Start Menu.bat` from the same folder: exit 0, `.lnk` gone.

## C. Forced failure path (the round-1 Major on false success)

A byte-identical scratch copy of `Add to Start Menu.bat` with `WS_DIR` pointed
at a nonexistent folder: the PowerShell block threw
("WindowSweeper.ps1 not found in ..."), exited 1; the batch printed
"Something went wrong - the shortcut was not created." and did NOT print the
success line; overall exit code 1. The Start Menu was untouched (the throw
fires before any shortcut-creation code).

## Isolation note

The tester's own pre-existing personal `Window Sweeper.lnk` (pointing at a
different install location) was backed up before each cycle and restored
byte-identically afterward (SHA256 C9DC8D7F... verified after every cycle).
No non-test windows or processes were touched.
