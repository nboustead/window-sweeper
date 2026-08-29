# Window Sweeper - Architecture

Condensed reference for anyone extending the tool. The source of truth is
always `WindowSweeper.ps1` itself - this document explains the shape of it,
not every line.

## Worker runspace + watchdog

All COM automation (Excel, Word, Outlook) runs inside **one persistent STA
runspace**, created lazily and reused across calls. Every touch goes through
`Invoke-ComTask`, which:

- Refuses to start a second call while one is already in flight
  (`$script:Busy` reentrancy guard) - this is what keeps a Timer tick or a
  grid event from racing a call that is still pumping.
- Pumps `Application.DoEvents()` every 15ms while it waits, so the UI stays
  responsive during a multi-second COM call.
- On timeout: tries `BeginStop()`, gives it a short grace period, and if the
  pipeline is still blocked inside an unmanaged COM call, **abandons the
  runspace** rather than disposing it (`Dispose()`/`EndInvoke()` would hang
  the UI too). A fresh runspace is created lazily on the next call; the
  abandoned one's thread rots harmlessly until Office itself unblocks it.

The worker script is self-contained source text (`$script:WorkerScript`),
shipped to the runspace via `AddScript().ToString()`, not a closure - every
helper function and COM object lives entirely inside that scriptblock's own
scope. **COM objects never cross the runspace boundary**; every op returns a
plain `PSCustomObject`/hashtable.

Application-level COM handles (the Excel/Word/Outlook `Application` object
itself) are cached in-runspace and revalidated with one cheap property read
before reuse. Documents/workbooks/items are never cached - they are always
re-resolved by key on every call, and that re-resolve doubles as the
liveness check.

## Action routing table

`$script:ActionRouting`, keyed `"App|RowType"`, is the single place that
maps a row to its ops:

```
@{ SaveCloseOp; CloseOp; SaveAsOp; ActivateStyle }
```

`Invoke-RowSaveClose` / `Invoke-RowClose` / `Invoke-RowActivate` /
`Invoke-BatchOp` all resolve through `Get-RoutingEntry` instead of
hardcoding per-app branches. An `(App, RowType)` with no entry (or a `$null`
op for the action being attempted) no-ops with a logged warning rather than
guessing. `ActivateStyle` is `'Win32Front'` (Hwnd-only fronting, no COM - the
row-click contract) or `'CdpActivate'` (Chrome CDP tabs, which have no HWND
of their own and activate over HTTP through the worker instead).

## Per-adapter notes

**Excel** - `Attach-ExcelInstances` walks visible `XLMAIN` windows,
resolving each to its `Application` object via
`XLMAIN -> XLDESK -> EXCEL7 -> AccessibleObjectFromWindow(OBJID_NATIVEOM)`.
This is what makes multi-instance Excel work (one Application COM pointer
per process id). Row-to-window matching handles the fact that a saved
file's XLMAIN title never carries its extension.

**Word** - only the ROT-registered instance is reachable
(`GetActiveObject('Word.Application')`, no launch fallback, no
multi-instance trick the way Excel has one). Hwnd comes straight from the
object model (`doc.Windows.Item(1).Hwnd`), not title matching - Word titles
include the file extension, unlike Excel, so there is no shared
title-matching helper. Protected View documents enumerate separately
(`$word.ProtectedViewWindows`) as their own row type - read-only, no save
state, closed via `pvw.Close()`.

**Outlook** - `Attach-Outlook` is `GetActiveObject`-only, same discipline as
Word (never launches Outlook as a side effect of a refresh). Inspector
resolution prefers `EntryID`; a not-yet-saved item falls back to a
`Subject + MessageClass + Index` composite, and if more than one live
inspector matches that composite, the resolve refuses rather than guessing
which one the stale key meant. Explorer (folder) windows are a separate row
type with their own liveness rule: the worker refuses to close the last
remaining Explorer, since that would exit Outlook's UI.

**Chrome** - two independent paths. The **CDP path** is config-gated,
never automatic (`chromeDebugPorts` defaults empty): a configured port is
queried over HTTP (`/json/list`), filtered strictly to `type == 'page'` so
extension service workers/background pages never appear as fake tabs, and
tab close/activate go through `/json/close/<id>` / `/json/activate/<id>`.
The **window fallback** (default when no port is configured) enumerates
visible `Chrome_WidgetWin_1` windows - filtered to windows whose owning
process is actually `chrome.exe`, since that window class is not exclusive
to Chrome (Electron apps share it) - and closes with a plain `WM_CLOSE`
(Chrome's own "close all tabs?" prompt, if any, is not suppressed).

## Safety contracts

- **Fail-closed discard gate.** Every Close op re-reads live save-state
  right before discarding, not the row's cached status from the last
  refresh. If that live read itself fails (property inaccessible), the
  default is **unsaved**, never saved - an unreadable state always forces
  the confirm gate rather than silently letting an unconfirmed close
  through. This applies uniformly to Excel, Word, and Outlook, and to the
  enum-time status label as well as the close-time re-check.
- **Batch scoping.** *Close All Saved* only ever touches rows already
  showing `Saved` (Chrome rows never qualify - they have no save state).
  *Save & Close All* excludes never-saved (pathless) rows, tallying them as
  skipped rather than guessing at a filename. *Close All and Discard* is
  the one batch that reaches every closable row regardless of status - the
  confirm dialog splits the list into "unsaved changes" vs. a separate
  "browser tabs (nothing to save)" section for any Chrome rows in the
  batch.
- **Test fences.** Five environment variables gate test behavior, never
  read from settings, and are documented here so a fence is never
  mistaken for a bug:
  - `WINDOWSWEEPER_TEST_PID` - scopes Excel enumeration to one instance.
  - `WINDOWSWEEPER_TEST_FILTER` - a name regex applied to Excel/Word/
    Outlook/Chrome rows alike, for fixture-scoped GUI testing.
  - `WINDOWSWEEPER_CHROME_DEBUG_PORT` - when set, is the *only* CDP port
    ever probed; settings-configured `chromeDebugPorts` are ignored
    entirely for the whole call.
  - `WINDOWSWEEPER_CHROME_TEST_PID` - fences the Chrome window-fallback
    path to one PID's windows, and both `ChromeWindowClose` and the
    Win32Front activate path refuse any hwnd whose live PID (or, for
    activate, whose owning process image) does not match, re-checked at
    action time.
  - `WINDOWSWEEPER_DEBUG=1` - turns on the trace log at
    `%USERPROFILE%\.window-sweeper\logs\debug.log`; zero cost when unset.

## Known limitations (v0.4.0)

- Word: only the single ROT-registered instance is reachable; no
  multi-instance support.
- Chrome CDP tab fronting is best-effort only (no HWND from CDP) and backs
  off silently when more than one visible window shares a tab's title.
- The Chrome window fallback closes an entire browser window, not
  individual tabs - Chrome's own confirmation prompt, if it appears, is
  not suppressed or answered for you.
- No PowerPoint support, no EXE packaging (still a `.ps1` + launcher
  `.bat`), no auto-update, no code signing.
- Windows PowerShell 5.1 only (not PowerShell 7 / pwsh).
