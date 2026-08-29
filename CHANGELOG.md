# Changelog

## 0.4.1 - 2026-08-29

Small follow-up release - no code changes to `WindowSweeper.ps1` beyond the
version string.

- Added `Add to Start Menu.bat` and `Remove from Start Menu.bat`, so the
  tool can show up in Windows Start search and be pinned to the taskbar
  like an installed program, instead of only being reachable by
  double-clicking a file in the extracted folder. `README.txt`/`README.md`
  updated accordingly (Desktop shortcut instructions kept as an
  alternative).

## 0.4.0 - 2026-08-29

Publish release. Four phases on top of the 0.3.1 baseline:

- **Phase 0 - generic action routing.** Replaced the two-app (Excel vs.
  everything-else) hardcoding with a routing table keyed `(App, RowType)`,
  so every save/close/activate op resolves generically instead of guessing.
  Filter UI moved from Both/Excel/Outlook radios to per-app checkboxes
  (`selectedApps` in settings, migrated automatically from the old
  `defaultFilter`).
- **Phase 1 - UI glow-up.** Per-app icon column with a 4px accent stripe,
  status rendered as rounded pills instead of plain colored text, app-
  grouped row ordering, Segoe UI typography, flat-styled buttons, a
  generated window icon.
- **Phase 2 - Word adapter.** Documents and Protected View windows, Hwnd
  resolved from the object model (not title matching), a guard against
  Word's Save() blocking mid-op on a pathless document.
- **Phase 3 - Chrome tabs adapter.** Config-gated CDP tab listing
  (`chromeDebugPorts`, empty by default) plus a window-level fallback for
  the common case of no debug port configured. Both close/activate paths
  verify the target window is chrome.exe-owned before acting, since
  Electron apps share Chrome's window class.
- **Cross-app hardening.** The unsaved-state re-check that gates a discard
  now fails *closed* everywhere (Excel, Word, Outlook) - an unreadable save
  state is always treated as unsaved, never silently treated as safe to
  discard.
- **Publish.** MIT license, public README, this changelog, an architecture
  doc, and a reproducible release-zip build script.

## 0.3.x - responsiveness + reliability rework

- Replaced COM-driven row-click activation with Win32-only window fronting
  (`ShowWindow` + `SetForegroundWindow` on a captured Hwnd) - row clicks no
  longer freeze the UI waiting on a cross-apartment COM call.
- Scoped, per-app refresh (`Do-ScopedRefresh`) replaced always-refreshing
  both apps on every action, so a busy Outlook no longer taxes every single
  Excel action with a 10-20s wait.
- Outlook busy-backoff: once found busy, skip touching it again for a
  cooldown window instead of re-triggering the same block on every refresh.
- Deferred-refresh timer: grid rebuilds now happen a tick after the
  triggering click, outside the DataGridView's own event dispatch, fixing
  a reentrancy crash that could leave the grid showing fewer rows than the
  counts label reported.
- Persistent Application-handle caching in the worker runspace, cutting
  the per-action COM attach cost.
- Debug trace log (`WINDOWSWEEPER_DEBUG=1`) for diagnosing UI-thread
  timing issues that leave no other trace.

## 0.2.0 - batch actions

- Save & Close All / Close All Saved / Close All and Discard, with a
  discard-confirmation dialog that lists every affected file/email by name
  before anything closes.
- Per-action log (`logs/actions-YYYY-MM-DD.csv`).

## 0.1.0 - initial internal version

- Single persistent STA worker runspace for all COM automation, with a
  timeout/watchdog around every call.
- Excel workbook and Outlook inspector/folder enumeration, one row per
  open item; single-row save-and-close or close.
