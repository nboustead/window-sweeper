WINDOW SWEEPER
==============

What this does:
You end up with a lot of Excel workbooks and Outlook email windows open
at once. Window Sweeper lists every one of them in a single window so
you can save-and-close or just close each one, without clicking through
them one at a time and answering "Do you want to save changes?" for each.

How to use it:
1. Double-click Run-WindowSweeper.bat. Actions are quick, usually 1-3
   seconds each.
2. Pick Both, Excel, or Outlook at the top; click Refresh any time.
3. Click "Save & Close" or "X" on a row to act on just that window, or
   use Save & Close All / Close All Saved / Close All and Discard at
   the bottom for many at once. Close All and Discard always lists what
   it is about to throw away first.
4. Keep on top (checked by default) keeps the sweeper above other
   windows, so clicking a row's name brings it up underneath instead of
   hiding it.
5. See an orange "Outlook busy" banner? Click "Show the dialog", handle
   whatever Outlook is asking, then click Refresh.

Rows in red have unsaved changes; closing one with X always confirms
first, listing exactly what will be discarded.

Having trouble and want Claude to look into it? Close the tool, set
WINDOWSWEEPER_DEBUG=1, and reopen it - a detailed log gets written to
C:\Users\taver\.window-sweeper\logs\debug.log.

MAKE A DESKTOP SHORTCUT
========================
1. Right-click Run-WindowSweeper.bat in this folder.
2. Choose "Show more options", then "Send to", then
   "Desktop (create shortcut)". (On some versions of Windows, "Send to"
   is already on the first right-click menu.)
3. Rename the new Desktop icon to "Window Sweeper" if you like.

MANUAL VERIFICATION CHECKLIST (run after any change to the tool)
==================================================================
- Both Excel instances are listed, each workbook shown once, with
  [#1] / [#2] tags; PERSONAL.XLSB never appears.
- Clicking X on an UNSAVED workbook closes it with no Excel save prompt.
- A never-saved Book1: clicking Save & Close opens Excel's Save As
  dialog (in front of the sweeper, even with Keep on top checked);
  clicking Cancel in that dialog leaves the workbook open.
- Saving and closing an open Outlook draft lands it in the Drafts folder.
- Closing the last Outlook folder window is refused with a message
  instead of exiting Outlook's UI.
- Hang drill: open Outlook's Import/Export wizard, then click Refresh.
  The Window Sweeper window stays usable, an orange busy banner appears
  within a few seconds, "Show the dialog" brings the wizard to the
  front, and Excel rows are unaffected. Close the wizard, click Refresh
  again, and the tool recovers cleanly.
- "Close All Saved" only touches rows already marked Saved.
- "Close All and Discard" lists every affected file/email name before
  asking you to confirm.
- Every action (save, close, discard) adds one row to
  C:\Users\taver\.window-sweeper\logs\actions-YYYY-MM-DD.csv
- Closing excel.exe from outside the tool (Task Manager), then clicking
  Refresh, updates the list cleanly with no error.
