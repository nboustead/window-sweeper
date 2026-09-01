<#
Window Sweeper
Triage and force-close open Excel workbooks and Outlook message/folder windows
without answering a "Save changes?" prompt for each one by hand.

Usage:
  Run-WindowSweeper.bat              normal use (hidden console, WinForms GUI)
  WindowSweeper.ps1 -Diag            console-only enumeration, no GUI, exit 0
  WINDOWSWEEPER_NOGUI=1 + dot-source everything defined, script returns, no GUI/-Diag

Windows PowerShell 5.1 only. See the design plan for the full architecture.
#>

param(
    [switch]$Diag
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:StateDir     = Join-Path $env:USERPROFILE ".window-sweeper"
$script:LogDir        = Join-Path $script:StateDir "logs"
$script:SettingsPath  = Join-Path $script:StateDir "settings.json"

# v0.4 Phase 0: the fixed set of apps the tool knows about and the display/
# enumeration order used throughout (filter checkboxes, row grouping, counts
# label). Word and Chrome are seeded here ahead of their Phase 2/3 adapters -
# they appear as filter checkboxes and grouping slots that simply produce no
# rows until EnumWord/EnumChrome exist.
$script:KnownApps = @('Excel', 'Word', 'Outlook', 'Chrome')
$script:AppOrder = @{}
for ($i = 0; $i -lt $script:KnownApps.Count; $i++) { $script:AppOrder[$script:KnownApps[$i]] = $i }

# RowTypes that represent an actual open document/item (as opposed to a
# Busy/Info/Explorer placeholder) - what Get-CountsSummary tallies per app.
# Word adds Document/ProtectedView in Phase 2; Chrome adds Tab/BrowserWindow
# in Phase 3.
$script:CountableRowTypes = @('Workbook', 'Inspector', 'Document', 'ProtectedView', 'Tab', 'BrowserWindow')

# v0.4 Phase 0: generic action routing, keyed "App|RowType" -> the op names
# and activation style for that combination. Invoke-RowSaveClose/
# Invoke-RowClose/Invoke-RowActivate/Invoke-BatchOp all resolve through this
# instead of hardcoding "Excel vs everything else" - the two-app assumption
# Sol flagged as a P1 blocker for adding Word/Chrome. An (App, RowType) with
# no entry here (or a $null op for the action being attempted) no-ops with a
# logged warning rather than guessing at an op name.
# ActivateStyle 'Win32Front' = the existing Hwnd-fronting-only click behavior
# (all COM apps, and Chrome windows once Phase 3 lands). 'CdpActivate' is
# reserved for Phase 3 Chrome CDP tabs, which have no HWND of their own.
$script:ActionRouting = @{
    'Excel|Workbook'      = @{ SaveCloseOp = 'ExcelSaveClose'; CloseOp = 'ExcelClose'; SaveAsOp = 'ExcelSaveAs'; ActivateStyle = 'Win32Front' }
    'Word|Document'       = @{ SaveCloseOp = 'WordSaveClose'; CloseOp = 'WordClose'; SaveAsOp = 'WordSaveAs'; ActivateStyle = 'Win32Front' }
    'Word|ProtectedView'  = @{ SaveCloseOp = $null; CloseOp = 'WordProtectedViewClose'; SaveAsOp = $null; ActivateStyle = 'Win32Front' }
    'Outlook|Inspector'   = @{ SaveCloseOp = 'OutlookSaveClose'; CloseOp = 'OutlookClose'; SaveAsOp = $null; ActivateStyle = 'Win32Front' }
    'Outlook|Explorer'    = @{ SaveCloseOp = $null; CloseOp = 'ExplorerClose'; SaveAsOp = $null; ActivateStyle = 'Win32Front' }
    'Chrome|Tab'          = @{ SaveCloseOp = $null; CloseOp = 'ChromeTabClose'; SaveAsOp = $null; ActivateStyle = 'CdpActivate' }
    'Chrome|BrowserWindow'= @{ SaveCloseOp = $null; CloseOp = 'ChromeWindowClose'; SaveAsOp = $null; ActivateStyle = 'Win32Front' }
}

function Get-RoutingEntry {
    param([string]$App, [string]$RowType)
    $key = "$App|$RowType"
    if ($script:ActionRouting.ContainsKey($key)) { return $script:ActionRouting[$key] }
    return $null
}

# ===== Region 1: Bootstrap + Win32 =====

$win32Src = @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class WinInfo
{
    public IntPtr Hwnd;
    public uint Pid;
    public string ClassName;
    public string Title;
    public bool Visible;
}

public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

// v0.4.2: multi-monitor fronting - land the target window on a different
// monitor than the sweeper itself when more than one exists.
public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

public struct MONITORINFO {
    public uint cbSize;
    public RECT rcMonitor;
    public RECT rcWork;
    public uint dwFlags;
}

public delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);

public class SweeperWin32
{
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lp);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder buf, int max);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder buf, int max);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string win);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);

    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] ref object ppv);

    // v0.4 Phase 3: Chrome window-fallback close (WM_CLOSE, fire-and-forget
    // - PostMessage never blocks on the target's message pump the way
    // SendMessage would) and liveness check. SetLastError=true (Sol MAJOR)
    // so a $false return can be paired with Marshal.GetLastWin32Error() on
    // the PowerShell side instead of silently reporting success.
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);

    // v0.4.2: row-click fronting maximize fix - IsIconic/IsZoomed let the
    // caller tell a minimized window from a merely-normal one before
    // deciding whether to force-restore or force-maximize it.
    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsZoomed(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

    [DllImport("user32.dll")]
    public static extern int GetSystemMetrics(int nIndex);

    [DllImport("user32.dll")]
    public static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    public const int SW_RESTORE = 9;
    public const int SW_SHOWMAXIMIZED = 3;
    public const uint WM_NULL = 0x0000;
    public const uint WM_CLOSE = 0x0010;
    public const uint SMTO_ABORTIFHUNG = 0x0002;
    public const uint MONITOR_DEFAULTTONEAREST = 2;
    public const int SM_CMONITORS = 80;
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOZORDER = 0x0004;
    private const uint OBJID_NATIVEOM = 0xFFFFFFF0;

    // Same "enumerate in C#, not a PS scriptblock delegate" reasoning as
    // GetTopLevelWindows - the callback must not be GC'd mid-enumeration.
    public static List<IntPtr> GetAllMonitors()
    {
        List<IntPtr> result = new List<IntPtr>();
        MonitorEnumProc cb = delegate(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData)
        {
            result.Add(hMonitor);
            return true;
        };
        EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, cb, IntPtr.Zero);
        return result;
    }

    // Enumeration lives in C#, not a PowerShell scriptblock delegate, so the
    // callback cannot be garbage-collected mid-enumeration.
    public static List<WinInfo> GetTopLevelWindows()
    {
        List<WinInfo> result = new List<WinInfo>();
        EnumWindowsProc cb = delegate(IntPtr hWnd, IntPtr lp)
        {
            StringBuilder cls = new StringBuilder(256);
            GetClassName(hWnd, cls, cls.Capacity);
            StringBuilder title = new StringBuilder(512);
            GetWindowText(hWnd, title, title.Capacity);
            uint pid;
            GetWindowThreadProcessId(hWnd, out pid);
            WinInfo info = new WinInfo();
            info.Hwnd = hWnd;
            info.Pid = pid;
            info.ClassName = cls.ToString();
            info.Title = title.ToString();
            info.Visible = IsWindowVisible(hWnd);
            result.Add(info);
            return true;
        };
        EnumWindows(cb, IntPtr.Zero);
        return result;
    }

    // XLMAIN -> XLDESK -> EXCEL7 -> AccessibleObjectFromWindow(native OM) -> Window;
    // caller reads .Application off the returned Window object.
    public static object GetExcelDispatchFromXlMain(IntPtr xlmain)
    {
        IntPtr xldesk = FindWindowEx(xlmain, IntPtr.Zero, "XLDESK", null);
        if (xldesk == IntPtr.Zero) return null;
        IntPtr excel7 = FindWindowEx(xldesk, IntPtr.Zero, "EXCEL7", null);
        if (excel7 == IntPtr.Zero) return null;
        Guid iidDispatch = new Guid("00020400-0000-0000-C000-000000000046");
        object ppv = null;
        int hr = AccessibleObjectFromWindow(excel7, OBJID_NATIVEOM, ref iidDispatch, ref ppv);
        if (hr != 0) return null;
        return ppv;
    }
}
"@
# Guarded: a second dot-source in the same session (the documented
# WINDOWSWEEPER_NOGUI test-harness path) would otherwise hit a terminating
# "type already exists" error before ever reaching the NOGUI early return.
if (-not ('SweeperWin32' -as [type])) {
    Add-Type -TypeDefinition $win32Src -ReferencedAssemblies System.Windows.Forms, System.Drawing
}

function Initialize-StateDir {
    try {
        if (-not (Test-Path $script:StateDir)) { New-Item -ItemType Directory -Path $script:StateDir -Force | Out-Null }
        if (-not (Test-Path $script:LogDir)) { New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null }
    } catch {
        # settings/logging just will not persist this run - never fatal
    }
}

function Get-DefaultSettings {
    # chromeDebugPorts default EMPTY (Sol P1) - the CDP path is config-gated,
    # never automatic; Nick's :9223 DealerCONNECT Chrome is a live
    # automation surface for ~9 other tools and must never be probed
    # uninvited. No UI control in v0.4 - README documents the manual edit.
    return [PSCustomObject]@{ selectedApps = @($script:KnownApps); confirmDiscard = $true; keepOnTop = $true; chromeDebugPorts = @() }
}

# v0.4 Phase 0: selectedApps (string array) replaces the old defaultFilter
# ('Both'/'Excel'/'Outlook') string. A settings.json still carrying the
# legacy key (no selectedApps yet) is migrated in place - 'Both' -> every
# known app, 'Excel'/'Outlook' -> just that one - and rewritten immediately
# so the on-disk file reflects the v0.4 shape from the very next load, not
# just whenever the window next closes.
function Get-Settings {
    if (-not (Test-Path $script:SettingsPath)) { return Get-DefaultSettings }
    try {
        $obj = Get-Content -Path $script:SettingsPath -Raw | ConvertFrom-Json
        $result = Get-DefaultSettings
        $needsRewrite = $false
        $hasSelectedApps = $obj.PSObject.Properties.Match('selectedApps').Count -gt 0
        $hasDefaultFilter = $obj.PSObject.Properties.Match('defaultFilter').Count -gt 0
        if ($hasSelectedApps -and ($obj.selectedApps -is [System.Array])) {
            # An explicit selectedApps ARRAY is authoritative - including an
            # explicit [] (Nick unchecked every app). Only a real JSON array
            # counts: null / a bare string / an object are malformed and fall
            # through to the legacy-filter or default paths instead of
            # silently becoming "no apps". A non-empty array whose entries
            # are ALL unrecognized is likewise treated as malformed.
            $apps = @($obj.selectedApps | Where-Object { $script:KnownApps -contains $_ } | Select-Object -Unique)
            if ($apps.Count -gt 0 -or @($obj.selectedApps).Count -eq 0) {
                $result.selectedApps = $apps
            } elseif ($hasDefaultFilter) {
                switch ($obj.defaultFilter) {
                    'Excel'   { $result.selectedApps = @('Excel') }
                    'Outlook' { $result.selectedApps = @('Outlook') }
                    default   { $result.selectedApps = @($script:KnownApps) }
                }
            }
        } elseif ($hasDefaultFilter) {
            switch ($obj.defaultFilter) {
                'Excel'   { $result.selectedApps = @('Excel') }
                'Outlook' { $result.selectedApps = @('Outlook') }
                default   { $result.selectedApps = @($script:KnownApps) }
            }
        }
        # Legacy key removal is unconditional on defaultFilter being present
        # at all - even a file that already has a (now-authoritative)
        # selectedApps can still be carrying a stale defaultFilter alongside
        # it, and that must not survive to the next load either.
        if ($hasDefaultFilter) { $needsRewrite = $true }
        # chromeDebugPorts: array-only, ints only (1-65535), same discipline
        # as selectedApps - anything else (missing, not an array, non-numeric
        # entries) leaves the Get-DefaultSettings empty-array default in
        # place rather than guessing at a port to probe.
        if ($obj.PSObject.Properties.Match('chromeDebugPorts').Count -gt 0 -and ($obj.chromeDebugPorts -is [System.Array])) {
            $ports = @()
            foreach ($p in $obj.chromeDebugPorts) {
                $parsedPort = 0
                if ([int]::TryParse("$p", [ref]$parsedPort) -and $parsedPort -gt 0 -and $parsedPort -le 65535) { $ports += $parsedPort }
            }
            $result.chromeDebugPorts = @($ports | Select-Object -Unique)
        }
        if ($obj.PSObject.Properties.Match('confirmDiscard').Count -gt 0) {
            $result.confirmDiscard = [bool]$obj.confirmDiscard
        }
        if ($obj.PSObject.Properties.Match('keepOnTop').Count -gt 0) {
            $result.keepOnTop = [bool]$obj.keepOnTop
        }
        if ($needsRewrite) {
            # Migration write: drops the legacy defaultFilter key for good
            # (the migrated $result object never had one) while preserving
            # whatever selectedApps was already authoritative above.
            Save-Settings -Settings $result
        }
        return $result
    } catch {
        return Get-DefaultSettings
    }
}

function Save-Settings {
    param($Settings)
    try {
        Initialize-StateDir
        $Settings | ConvertTo-Json -Depth 5 | Set-Content -Path $script:SettingsPath -Encoding UTF8
    } catch {
        # non-fatal
    }
}

# v0.3.1: env-gated trace log for diagnosing UI-thread reentrancy bugs (grid
# corruption, close-during-busy hangs) that leave no trace in the action log
# and whose exceptions get swallowed into the runspace's own $Error when
# raised from inside a WinForms event handler. Zero effect/zero cost when
# WINDOWSWEEPER_DEBUG is unset - callers should still avoid calling this
# inside tight loops (e.g. the DoEvents wait loop) even though the env check
# is cheap, since Add-Content per iteration would itself distort timing.
function Write-DebugLog {
    param([string]$Message)
    if ($env:WINDOWSWEEPER_DEBUG -ne '1') { return }
    try {
        Initialize-StateDir
        $logFile = Join-Path $script:LogDir 'debug.log'
        $line = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + ' [' + [System.Threading.Thread]::CurrentThread.ManagedThreadId + '] ' + $Message
        Add-Content -Path $logFile -Value $line -Encoding UTF8
    } catch {
        # debug logging must never be fatal
    }
}

function Write-ActionLog {
    param([string]$App, [string]$Name, [string]$Action, $Result, [string]$Path = '')
    try {
        Initialize-StateDir
        $logFile = Join-Path $script:LogDir ("actions-{0}.csv" -f (Get-Date -Format 'yyyy-MM-dd'))
        $entry = [PSCustomObject]@{
            Timestamp = Get-Date -Format 's'
            App       = $App
            Name      = $Name
            FullPath  = $Path
            Action    = $Action
            Ok        = $Result.Ok
            Message   = $Result.Message
        }
        $entry | Export-Csv -Path $logFile -Append -NoTypeInformation -Encoding UTF8
    } catch {
        # logging must never break the app
    }
}

# ===== Region 2: Watchdog =====

$script:ComRunspace = $null
$script:ZombieCount = 0
$script:Busy = $false                       # true while a pipeline is actually in flight
$script:ClosingRequested = $false           # set by FormClosing so an in-flight wait bails out
$script:OutlookBackoffUntil = [datetime]::MinValue  # v0.3: skip Outlook enum until this time passes

# v0.3 progress feedback: disables the three content panels (not the Form
# itself, so it can still be moved/closed) and shows a marquee + status text
# while a worker call is in flight. Every $script:XXX reference here is set
# later by New-MainForm; the try/catch plus per-control existence checks make
# this a harmless no-op in -Diag/NOGUI mode, where no Form ever exists.
function Set-BusyUi {
    param([bool]$Busy, [string]$StatusText = '')
    try {
        if ($script:TopPanel) { $script:TopPanel.Enabled = -not $Busy }
        if ($script:Grid) { $script:Grid.Enabled = -not $Busy }
        if ($script:BottomPanel) { $script:BottomPanel.Enabled = -not $Busy }
        if ($script:ProgressBar) { $script:ProgressBar.Visible = $Busy }
        if ($script:LblStatus) {
            $text = ''
            if ($Busy) { $text = $StatusText }
            $script:LblStatus.Text = $text
        }
        if ($script:MainForm) {
            $cur = [System.Windows.Forms.Cursors]::Default
            if ($Busy) { $cur = [System.Windows.Forms.Cursors]::WaitCursor }
            $script:MainForm.Cursor = $cur
        }
    } catch {}
}

function New-ComWorker {
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()
    return $rs
}

function Get-ComRunspace {
    if (-not $script:ComRunspace -or $script:ComRunspace.RunspaceStateInfo.State -ne 'Opened') {
        $script:ComRunspace = New-ComWorker
    }
    return $script:ComRunspace
}

# An open STA runspace pumps a background message loop that keeps the host
# process alive even after the console/GUI is done - a plain "exit" prints
# fine but the process never actually terminates. Every exit path must call
# this (or, if the runspace was already abandoned as a zombie, the process
# is force-terminated separately - see Region 6).
# Guarded on $script:Busy: if a pipeline is actually in flight, Runspace.Close()
# can block on the exact same "stuck inside an unmanaged COM call" condition
# Invoke-ComTask's own BeginStop+grace+abandon exists to survive - and unlike
# Invoke-ComTask, FormClosing has no grace period of its own and can run
# reentrantly (from FormClosing firing inside Invoke-ComTask's own DoEvents
# pump) on the UI thread. When busy, skip cleanup entirely - same abandon
# policy as the watchdog - and let Region 6's post-Application.Run finally +
# [Environment]::Exit(0) force-terminate the process instead.
function Close-ComRunspace {
    if ($script:Busy) { return }
    try {
        if ($script:ComRunspace -and $script:ComRunspace.RunspaceStateInfo.State -eq 'Opened') {
            $script:ComRunspace.Close()
            $script:ComRunspace.Dispose()
        }
    } catch {}
    $script:ComRunspace = $null
}

# The choke point every COM touch runs through. Runs the worker scriptblock on
# a persistent STA runspace, pumping DoEvents on the caller so the UI stays
# alive while it waits. On timeout it tries BeginStop, then a 1s grace period,
# then abandons the runspace outright rather than disposing it: the pipeline
# thread is blocked inside an unmanaged COM call, so Dispose()/EndInvoke()
# would just hang the UI too. The blocked thread rots harmlessly until Office
# unblocks it; a fresh runspace is created lazily on the next call.
function Invoke-ComTask {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Command,
        [int]$TimeoutMs = 5000,
        $OwnerForm = $null,
        [string]$StatusText = 'Working...'
    )

    $opName = "$($Command.Op)"
    Write-DebugLog "Invoke-ComTask ENTER Op=$opName Busy=$script:Busy"
    # Reentrancy guard: a Timer tick, Activated, or (pre-fix) grid event could
    # fire while a previous call is still pumping DoEvents in its own wait
    # loop below - this is exactly the mechanism suspected of corrupting the
    # grid (two overlapping refreshes racing on $script:Grid). Refuse rather
    # than stack a second worker call on top of the first.
    if ($script:Busy) {
        Write-DebugLog "Invoke-ComTask REENTRANT-BLOCKED Op=$opName"
        return @{ Ok = $false; Reentered = $true; Message = 'busy' }
    }

    $swTotal = [System.Diagnostics.Stopwatch]::StartNew()
    $rs = Get-ComRunspace
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($script:WorkerScript.ToString())
    [void]$ps.AddArgument($Command)

    if ($OwnerForm) { Set-BusyUi -Busy $true -StatusText $StatusText }

    # $script:Busy brackets the whole in-flight window so Close-ComRunspace
    # (which can run reentrantly, from FormClosing firing inside the DoEvents
    # pump below) knows never to touch the runspace while this is running.
    $script:Busy = $true
    try {
        $async = $ps.BeginInvoke()
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while (-not $async.IsCompleted -and $sw.ElapsedMilliseconds -lt $TimeoutMs -and -not $script:ClosingRequested) {
            [System.Windows.Forms.Application]::DoEvents()
            # v0.3.1: 15ms, not 50ms - a single CloseMainWindow() call (Task
            # Manager, the coordinator's own test harness) reads
            # Process.Responding at the instant it is called, and a 50ms
            # sleep gap between DoEvents pumps was wide enough that it could
            # land mid-sleep and see "not responding," returning false
            # without ever delivering WM_CLOSE at all (confirmed via debug
            # log - FormClosing never fired). Shrinking the gap makes the
            # window responsive far more continuously; the CPU cost is
            # trivial and bounded to however long a COM wait actually runs.
            Start-Sleep -Milliseconds 15
        }

        # A close request (Task Manager End Task, taskbar "Close window",
        # logoff) bails out of the wait immediately rather than riding out the
        # full 15-20s timeout - it is treated exactly like a timeout below.
        $timedOut = -not $async.IsCompleted
        $closing = $script:ClosingRequested
        Write-DebugLog "Invoke-ComTask WAIT-DONE Op=$opName IsCompleted=$($async.IsCompleted) ClosingRequested=$closing ElapsedMs=$($swTotal.ElapsedMilliseconds)"
        if ($timedOut) {
            try {
                $stopResult = $ps.BeginStop($null, $null)
                # If we are closing, skip the bounded DoEvents-pumped grace
                # wait entirely - every extra DoEvents call here is another
                # chance to re-enter WinForms machinery mid-shutdown, and
                # there is no UI left to keep responsive for. Go straight to
                # abandon so Region 6's finally + [Environment]::Exit(0) runs
                # as fast as possible.
                if (-not $closing) {
                    $graceSw = [System.Diagnostics.Stopwatch]::StartNew()
                    while (-not $stopResult.IsCompleted -and $graceSw.ElapsedMilliseconds -lt 1000) {
                        [System.Windows.Forms.Application]::DoEvents()
                        Start-Sleep -Milliseconds 50
                    }
                }
            } catch {}
        }

        if ($OwnerForm -and -not $closing) {
            # The form may already be disposed if this is unwinding because
            # FormClosing fired reentrantly while this call was waiting.
            try { Set-BusyUi -Busy $false } catch {}
        }

        # Once timed out, never call EndInvoke: whether the grace-period stop
        # landed or not, the result (if any) is unreliable, and EndInvoke on a
        # stopped pipeline just throws "the pipeline has been stopped". Dispose
        # is safe if the stop actually completed; if it did not, the pipeline is
        # still blocked inside an unmanaged COM call, so the runspace is
        # abandoned rather than disposed (Dispose/EndInvoke would hang the UI too)
        # - its thread rots harmlessly until Office unblocks it, and a fresh
        # runspace is created lazily on the next call.
        if ($timedOut) {
            if ($async.IsCompleted) {
                try { $ps.Dispose() } catch {}
            } else {
                $script:ZombieCount++
                $script:ComRunspace = $null
            }
            Write-DebugLog "Invoke-ComTask EXIT Op=$opName Ok=false TimedOut=true Closing=$closing ElapsedMs=$($swTotal.ElapsedMilliseconds) Zombies=$script:ZombieCount"
            return @{ Ok = $false; TimedOut = $true; Message = 'Timed out'; Zombies = $script:ZombieCount }
        }

        try {
            $output = $ps.EndInvoke($async)
            if ($ps.HadErrors) {
                $errText = ($ps.Streams.Error | ForEach-Object { $_.ToString() }) -join '; '
                Write-DebugLog "Invoke-ComTask EXIT Op=$opName Ok=false HadErrors ElapsedMs=$($swTotal.ElapsedMilliseconds) Message=$errText"
                return @{ Ok = $false; Message = $errText }
            }
            if ($output -and $output.Count -gt 0) {
                $last = $output[$output.Count - 1]
                Write-DebugLog "Invoke-ComTask EXIT Op=$opName Ok=$($last.Ok) ElapsedMs=$($swTotal.ElapsedMilliseconds)"
                return $last
            }
            Write-DebugLog "Invoke-ComTask EXIT Op=$opName Ok=false NoResult ElapsedMs=$($swTotal.ElapsedMilliseconds)"
            return @{ Ok = $false; Message = 'Worker returned no result' }
        } catch {
            Write-DebugLog "Invoke-ComTask EXIT Op=$opName EXCEPTION ElapsedMs=$($swTotal.ElapsedMilliseconds) Message=$($_.Exception.Message) Stack=$($_.ScriptStackTrace)"
            return @{ Ok = $false; Message = $_.Exception.Message }
        } finally {
            $ps.Dispose()
        }
    } finally {
        $script:Busy = $false
    }
}

# Cheap pre-flight gate: WM_NULL with SMTO_ABORTIFHUNG on one visible window
# of the process. Catches a dead message pump fast, before paying for the
# full Invoke-ComTask timeout. Not sufficient alone - a process with a modal
# dialog still answers WM_NULL but still parks COM calls - so Invoke-ComTask
# is the real safety net.
# Callers must check Get-Process themselves before calling this - a process
# that is not running is not "responsive" (returning true here previously let
# Get-Snapshot fall through to EnumOutlook when Outlook was not running).
function Test-AppResponsive {
    param([string]$ProcessName, [int]$TimeoutMs = 1000)
    try {
        $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if (-not $procs) { return $false }
        $pids = @($procs | ForEach-Object { $_.Id })
        $win = [SweeperWin32]::GetTopLevelWindows() | Where-Object { $_.Visible -and ($pids -contains $_.Pid) } | Select-Object -First 1
        if (-not $win) { return $true }
        $resultPtr = [UIntPtr]::Zero
        $r = [SweeperWin32]::SendMessageTimeout($win.Hwnd, [SweeperWin32]::WM_NULL, [IntPtr]::Zero, [IntPtr]::Zero, [SweeperWin32]::SMTO_ABORTIFHUNG, [uint32]$TimeoutMs, [ref]$resultPtr)
        return ($r -ne [IntPtr]::Zero)
    } catch {
        return $true
    }
}

function Find-BlockingDialog {
    param([string]$ProcessName)
    $found = @()
    try {
        $pids = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        if ($pids.Count -eq 0) { return $found }
        foreach ($w in [SweeperWin32]::GetTopLevelWindows()) {
            if ($w.ClassName -eq '#32770' -and $w.Visible -and ($pids -contains $w.Pid)) { $found += $w }
        }
    } catch {}
    return $found
}

function Show-BlockingDialog {
    param([string]$ProcessName)
    $dialogs = Find-BlockingDialog -ProcessName $ProcessName
    foreach ($d in $dialogs) {
        try {
            [SweeperWin32]::ShowWindow($d.Hwnd, [SweeperWin32]::SW_RESTORE) | Out-Null
            [SweeperWin32]::SetForegroundWindow($d.Hwnd) | Out-Null
        } catch {}
    }
    return ($dialogs.Count -gt 0)
}

# ===== Region 3: Worker scriptblock =====
# Runs inside the persistent STA runspace via Invoke-ComTask (or directly, on
# a console STA thread, when WINDOWSWEEPER_NOGUI=1). Self-contained: every
# helper it needs is defined inside the scriptblock body itself, since it is
# shipped to the runspace as source text (.ToString()), not a closure.
# COM objects never leave this scriptblock - every returned row/result is a
# plain PSCustomObject/hashtable. Every action re-resolves its target by a
# stable key on every call; that re-resolve doubles as the liveness check.

$script:WorkerScript = {
    param($cmd)

    # v0.3: persistent attachment cache. $script: here resolves to the
    # runspace's own top-level scope, which DOES survive across separate
    # AddScript().Invoke() calls on the same reused runspace (confirmed by
    # direct test) - so this cache is warm for every op after the first,
    # cutting out the AccessibleObjectFromWindow + .Application marshaling
    # cost per call. Application objects only, ever - never workbooks/items,
    # which are always re-resolved fresh by key (that re-resolve is the
    # liveness/staleness check the rest of the file depends on). If the
    # runspace is ever abandoned as a zombie, the next one just starts with
    # no $script:CachedExcelApps at all - an empty cache, not a safety change.
    if (-not $script:CachedExcelApps) { $script:CachedExcelApps = @{} }

    function Attach-ExcelInstances {
        # One Application COM pointer per unique process id. Excel SDI means
        # one XLMAIN per workbook, so a process can own several XLMAIN
        # windows; only one is needed to reach that process's Application.
        $liveXlmainHwnd = @{}
        foreach ($w in [SweeperWin32]::GetTopLevelWindows()) {
            if ($w.ClassName -ne 'XLMAIN' -or -not $w.Visible) { continue }
            if ($liveXlmainHwnd.ContainsKey($w.Pid)) { continue }
            $liveXlmainHwnd[$w.Pid] = $w.Hwnd
        }

        $result = @{}
        # Validate cache hits with one cheap property read each; drop
        # anything whose process no longer has a live XLMAIN, or whose
        # cached Application object no longer answers.
        foreach ($cachedPid in @($script:CachedExcelApps.Keys)) {
            if (-not $liveXlmainHwnd.ContainsKey($cachedPid)) {
                $script:CachedExcelApps.Remove($cachedPid)
                continue
            }
            $app = $script:CachedExcelApps[$cachedPid]
            $stillGood = $false
            try { $null = $app.Visible; $stillGood = $true } catch {}
            if ($stillGood) {
                $result[$cachedPid] = [PSCustomObject]@{ ExcelPid = $cachedPid; App = $app }
            } else {
                $script:CachedExcelApps.Remove($cachedPid)
            }
        }

        # Attach fresh only for a live pid the cache did not already cover.
        foreach ($livePid in $liveXlmainHwnd.Keys) {
            if ($result.ContainsKey($livePid)) { continue }
            try {
                $winObj = [SweeperWin32]::GetExcelDispatchFromXlMain($liveXlmainHwnd[$livePid])
                if (-not $winObj) { continue }
                $app = $winObj.Application
                $script:CachedExcelApps[$livePid] = $app
                $result[$livePid] = [PSCustomObject]@{ ExcelPid = $livePid; App = $app }
            } catch { }
        }

        if ($result.Count -eq 0) {
            try {
                $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
                $hwnd = [IntPtr]$app.Hwnd
                [uint32]$procId = 0
                [void][SweeperWin32]::GetWindowThreadProcessId($hwnd, [ref]$procId)
                if ($procId -ne 0) {
                    $script:CachedExcelApps[$procId] = $app
                    $result[$procId] = [PSCustomObject]@{ ExcelPid = $procId; App = $app }
                }
            } catch { }
        }
        return @($result.Values)
    }

    function Resolve-Workbook {
        param($key)
        if ($key.Kind -ne 'PidName') { return $null }
        $inst = Attach-ExcelInstances | Where-Object { $_.ExcelPid -eq $key.Pid } | Select-Object -First 1
        if (-not $inst) { return $null }
        try {
            foreach ($wb in @($inst.App.Workbooks)) {
                try { if ($wb.Name -eq $key.Name) { return $wb } } catch {}
            }
        } catch {}
        return $null
    }

    # The Outlook COM attach that once hung indefinitely - it only ever
    # executes here, inside the watchdog-guarded worker. GetActiveObject only:
    # this must never start a new Outlook process as a side effect of what
    # looks like a read-only enumeration/refresh. v0.3: cached like Excel -
    # validated with one cheap property read before reuse.
    function Attach-Outlook {
        if ($script:CachedOutlook) {
            try { $null = $script:CachedOutlook.Version; return $script:CachedOutlook }
            catch { $script:CachedOutlook = $null }
        }
        try {
            $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
            $script:CachedOutlook = $ol
            return $ol
        } catch { return $null }
    }

    # v0.4 Phase 2: Word COM attach - GetActiveObject ONLY, same discipline as
    # Attach-Outlook, and for the same reason: this must never launch a new
    # Word process as a side effect of what looks like a read-only
    # enumeration/refresh. Cached like Outlook, validated with one cheap
    # property read before reuse. Documented limitation (Sol): only the
    # ROT-registered instance is ever reachable - multi-instance Word is rare
    # and there is no OBJID_NATIVEOM trick for it the way Attach-ExcelInstances
    # has for Excel.
    function Attach-Word {
        if ($script:CachedWord) {
            try { $null = $script:CachedWord.Version; return $script:CachedWord }
            catch { $script:CachedWord = $null }
        }
        try {
            $wd = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
            $script:CachedWord = $wd
            return $wd
        } catch { return $null }
    }

    # Key = @{ Kind='WordName'; Name; FullName }. A key carrying a non-empty
    # FullName (Sol P1 - stale-key wrong-target) MUST match FullName exactly
    # - Name alone is never sufficient once a path exists, because the
    # original document could have closed and a DIFFERENT file that happens
    # to share the same filename (in another folder) could now be open,
    # silently becoming the target of the action. Name-alone matching is
    # only acceptable for a pathless key (a never-saved document at
    # snapshot time, where FullName was empty), and even then more than one
    # live document sharing that placeholder name is ambiguous - refuse
    # rather than guess, same contract as Resolve-Inspector's Ambiguous flag.
    function Resolve-WordDocument {
        param($word, $key)
        if ($key.Kind -ne 'WordName') { return @{ Doc = $null; Ambiguous = $false } }
        $nameMatches = @()
        foreach ($doc in @($word.Documents)) {
            try { if ($doc.Name -eq $key.Name) { $nameMatches += $doc } } catch {}
        }
        if ($key.FullName) {
            $pathMatches = @($nameMatches | Where-Object {
                $fn = $null
                try { $fn = $_.FullName } catch {}
                $fn -eq $key.FullName
            })
            if ($pathMatches.Count -eq 1) { return @{ Doc = $pathMatches[0]; Ambiguous = $false } }
            if ($pathMatches.Count -gt 1) { return @{ Doc = $null; Ambiguous = $true } }
            return @{ Doc = $null; Ambiguous = $false }
        }
        if ($nameMatches.Count -eq 1) { return @{ Doc = $nameMatches[0]; Ambiguous = $false } }
        if ($nameMatches.Count -gt 1) { return @{ Doc = $null; Ambiguous = $true } }
        return @{ Doc = $null; Ambiguous = $false }
    }

    # Same Index+VerifyCaption liveness contract as Resolve-Explorer.
    function Resolve-ProtectedView {
        param($word, $key)
        if ($key.Kind -ne 'Index') { return $null }
        $pvws = @($word.ProtectedViewWindows)
        if ($key.Index -ge $pvws.Count) { return $null }
        $pvw = $pvws[$key.Index]
        if ($key.VerifyCaption) {
            $cap = $null
            try { $cap = $pvw.Caption } catch {}
            if ($cap -ne $key.VerifyCaption) { return $null }
        }
        return $pvw
    }

    # EntryID when the item has one (anything already saved to a store);
    # otherwise a composite Subject+MessageClass+Index - Resolve-Inspector
    # uses the Index only to break a tie when Subject+MessageClass alone
    # matches more than one live inspector (e.g. two blank compose windows).
    function Get-ItemKey {
        param($item, $idx)
        $entryId = $null
        try { $entryId = $item.EntryID } catch {}
        if ($entryId) { return @{ Kind = 'EntryID'; Value = $entryId } }
        $subj = $null; $mc = $null
        try { $subj = $item.Subject } catch {}
        try { $mc = $item.MessageClass } catch {}
        return @{ Kind = 'SubjectIndex'; Subject = $subj; MessageClass = $mc; Index = $idx }
    }

    # Nick hit Outlook rows that never front on click: Inspector.Caption and
    # the live rctrl_renwnd32 window title often carry trailing/doubled
    # spaces that don't line up character-for-character even though they
    # are "the same" window. Trim + collapse whitespace runs before any
    # comparison, on both sides.
    function Get-NormalizedTitle {
        param([string]$Text)
        if (-not $Text) { return '' }
        return ($Text.Trim() -replace '\s+', ' ')
    }

    # Returns @{ Item = <resolved item or $null>; Ambiguous = <bool> } rather
    # than a bare item, so callers can tell "not found" (report "no longer
    # open") apart from "more than one live item looks the same and the
    # snapshot Index no longer identifies which one" (report a different,
    # more honest message rather than silently acting on the wrong window).
    function Resolve-Inspector {
        param($ol, $key)
        $inspectors = @($ol.Inspectors)
        if ($key.Kind -eq 'EntryID') {
            for ($i = 0; $i -lt $inspectors.Count; $i++) {
                $item = $null
                try { $item = $inspectors[$i].CurrentItem } catch { continue }
                if (-not $item) { continue }
                $eid = $null
                try { $eid = $item.EntryID } catch {}
                if ($eid -and $eid -eq $key.Value) { return @{ Item = $item; Ambiguous = $false } }
            }
            return @{ Item = $null; Ambiguous = $false }
        }
        # SubjectIndex: collect every live inspector whose Subject+MessageClass
        # still match. Exactly one -> unambiguous. More than one -> only trust
        # the snapshot Index if the item currently sitting there is itself one
        # of the matches; otherwise refuse rather than guess.
        $matches = @()
        for ($i = 0; $i -lt $inspectors.Count; $i++) {
            $item = $null
            try { $item = $inspectors[$i].CurrentItem } catch { continue }
            if (-not $item) { continue }
            $s = $null; $mcv = $null
            try { $s = $item.Subject } catch {}
            try { $mcv = $item.MessageClass } catch {}
            if ($s -eq $key.Subject -and $mcv -eq $key.MessageClass) {
                $matches += [PSCustomObject]@{ Index = $i; Item = $item }
            }
        }
        if ($matches.Count -eq 1) { return @{ Item = $matches[0].Item; Ambiguous = $false } }
        if ($matches.Count -gt 1) {
            $atIndex = $matches | Where-Object { $_.Index -eq $key.Index } | Select-Object -First 1
            if ($atIndex) { return @{ Item = $atIndex.Item; Ambiguous = $false } }
            return @{ Item = $null; Ambiguous = $true }
        }
        return @{ Item = $null; Ambiguous = $false }
    }

    function Resolve-Explorer {
        param($ol, $key)
        $explorers = @($ol.Explorers)
        if ($key.Kind -ne 'Index' -or $key.Index -ge $explorers.Count) { return $null }
        $exp = $explorers[$key.Index]
        if ($key.VerifyCaption) {
            $cap = $null
            try { $cap = $exp.Caption } catch {}
            if ($cap -ne $key.VerifyCaption) { return $null }
        }
        return $exp
    }

    # v0.4 Phase 3: minimal CDP HTTP client. HttpWebRequest (not
    # Invoke-WebRequest) so Timeout/Proxy are controlled precisely - a call
    # that must stay fast and silent has no use for a progress bar or the
    # cmdlet's extra parsing overhead. 2s timeout and no proxy per plan.
    # These are HTTP-only, no COM - never launches or touches anything
    # unless the caller already resolved a real port/id.
    function Invoke-ChromeCdpGet {
        param([string]$Url, [int]$TimeoutMs = 2000)
        $resp = $null
        try {
            $req = [System.Net.HttpWebRequest]::Create($Url)
            $req.Method = 'GET'
            $req.Timeout = $TimeoutMs
            $req.Proxy = $null
            $resp = $req.GetResponse()
            $stream = $resp.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $body = $reader.ReadToEnd()
            $reader.Dispose()
            return @{ Ok = $true; StatusCode = [int]$resp.StatusCode; Body = $body }
        } catch [System.Net.WebException] {
            $statusCode = 0
            $body = ''
            try {
                if ($_.Exception.Response) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                    $errStream = $_.Exception.Response.GetResponseStream()
                    $errReader = New-Object System.IO.StreamReader($errStream)
                    $body = $errReader.ReadToEnd()
                    $errReader.Dispose()
                }
            } catch {}
            return @{ Ok = $false; StatusCode = $statusCode; Body = $body; Message = $_.Exception.Message }
        } catch {
            return @{ Ok = $false; StatusCode = 0; Body = ''; Message = $_.Exception.Message }
        } finally {
            if ($resp) { try { $resp.Close() } catch {} }
        }
    }

    # Filters strictly to type='page' (panel P2 blocker) - extension
    # service workers/background pages otherwise appear as fake tabs.
    function Get-ChromeCdpPages {
        param([int]$Port)
        $r = Invoke-ChromeCdpGet -Url "http://127.0.0.1:$Port/json/list"
        if (-not $r.Ok) { return @() }
        $parsed = $null
        try { $parsed = $r.Body | ConvertFrom-Json } catch { return @() }
        if (-not $parsed) { return @() }
        $items = @($parsed)
        # Defensive consistency check: PowerShell 5.1's ConvertFrom-Json was
        # observed, rarely and non-deterministically during testing, to
        # collapse a genuine N-element JSON array into a single merged
        # object whose properties become space-joined concatenations of
        # every element's value (root cause not pinned down - reproduced
        # with byte-identical input succeeding in one process and failing
        # in another). PowerShell's `-eq` array-filter semantics would then
        # let that single garbled object slip through a `.type -eq 'page'`
        # filter. A raw substring count of `"id":` occurrences is a cheap,
        # independent cross-check; a mismatch means the parse is not
        # trustworthy - fail closed to zero rows (a merged/garbled id would
        # only ever 404 gracefully on close/activate anyway, since Chrome
        # tabs carry no data-loss risk, but a wrong on-screen count/title is
        # still worth never showing).
        $rawIdCount = ([regex]::Matches($r.Body, '"id"\s*:')).Count
        if ($items.Count -ne $rawIdCount) { return @() }
        return @($items | Where-Object { $_.type -eq 'page' })
    }

    # v0.4 Phase 3 hardening (Sol): Chrome_WidgetWin_1 is NOT exclusive to
    # Chrome - Electron apps (Claude, Cursor, VS Code, Slack, etc.) share
    # the same window class. Every window the fallback path touches, at
    # BOTH enum time and action time, must be confirmed to actually belong
    # to a chrome.exe process before it is ever shown as a row or closed -
    # otherwise an X on what looks like a "Chrome" row could WM_CLOSE an
    # unrelated application. Drop-on-failure: a window whose PID cannot be
    # resolved to a live process is never treated as Chrome.
    function Test-ChromeOwnedWindow {
        param([IntPtr]$Hwnd)
        try {
            [uint32]$winPid = 0
            [void][SweeperWin32]::GetWindowThreadProcessId($Hwnd, [ref]$winPid)
            if ($winPid -eq 0) { return $false }
            $proc = Get-Process -Id $winPid -ErrorAction Stop
            return ($proc.ProcessName -eq 'chrome')
        } catch {
            return $false
        }
    }

    if (-not $cmd -or -not $cmd.Op) { return @{ Ok = $false; Message = 'Invalid command' } }

    try {
        switch ($cmd.Op) {

            'EnumExcel' {
                $rows = @()
                $instances = @(Attach-ExcelInstances)
                $multi = $instances.Count -gt 1
                $idx = 0
                foreach ($inst in $instances) {
                    $idx++
                    $label = 'Excel'
                    if ($multi) { $label = "Excel [#$idx]" }
                    $wbList = $null
                    try { $wbList = @($inst.App.Workbooks) } catch {
                        $msg = 'Busy - could not read workbook list'
                        if ($_.Exception.Message -match '800AC472' -or $_.Exception.Message -match 'RPC_E_CALL_REJECTED') {
                            $msg = 'Press Enter or Esc in Excel first (it is in cell-edit mode)'
                        }
                        $rows += [PSCustomObject]@{ App = 'Excel'; RowType = 'Busy'; Name = $label; Detail = $msg; Status = 'Busy'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }
                        continue
                    }
                    $wins = @([SweeperWin32]::GetTopLevelWindows() | Where-Object { $_.ClassName -eq 'XLMAIN' -and $_.Pid -eq $inst.ExcelPid -and $_.Visible })
                    foreach ($wb in $wbList) {
                        try {
                            $name = $wb.Name
                            # The XLMAIN title bar drops the file extension (Book1.xlsx shows as
                            # just "Book1.xlsx - Excel" is wrong too - it is really "Book1 - Excel";
                            # a saved file shows as "Report - Excel", never "Report.xlsx - Excel"),
                            # but wb.Name always includes it - match on the name without extension.
                            $nameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($name)
                            # Prefer an exact "Name - Excel" / "Name - ..." match first - a bare
                            # StartsWith($nameNoExt) alone would let "Report" bind to "Report (2)"'s
                            # window. Only fall back to the loose prefix match for suffix variants
                            # that don't fit the "Name - " shape (e.g. "Name  [Compatibility Mode] - Excel").
                            $matchWin = $wins | Where-Object { $_.Title -eq "$nameNoExt - Excel" -or $_.Title.StartsWith("$nameNoExt - ") } | Select-Object -First 1
                            if (-not $matchWin) {
                                $matchWin = $wins | Where-Object { $_.Title.StartsWith($nameNoExt) } | Select-Object -First 1
                            }
                            if (-not $matchWin) { continue }  # hidden window (PERSONAL.XLSB, add-ins) - never listed
                            $hasPath = $false
                            try { $hasPath = [bool]$wb.Path } catch {}
                            # Fail CLOSED (Sol cross-app hardening): default
                            # to unsaved, not saved, so an unreadable .Saved
                            # never mislabels a row 'Saved' when the true
                            # state is unknown.
                            $saved = $false
                            try { $saved = [bool]$wb.Saved } catch {}
                            $status = 'Saved'
                            if (-not $hasPath) { $status = 'Never saved' }
                            elseif (-not $saved) { $status = 'UNSAVED' }
                            $display = $name
                            if ($multi) { $display = "[#$idx] $name" }
                            $detail = '(never saved)'
                            if ($hasPath) { try { $detail = $wb.Path } catch {} }
                            $rows += [PSCustomObject]@{
                                App = 'Excel'; RowType = 'Workbook'; Name = $display; Detail = $detail; Status = $status
                                Key = @{ Kind = 'PidName'; Pid = $inst.ExcelPid; Name = $name }
                                Hwnd = $matchWin.Hwnd.ToInt64(); CanSave = $true; CanClose = $true
                            }
                        } catch { continue }
                    }
                }
                # v0.3 test hooks - zero effect unless the env var is set.
                # WINDOWSWEEPER_TEST_PID scopes to one Excel instance (airtight
                # even for a never-saved fixture like "Book2" that a name regex
                # cannot tell apart from a real never-saved Book1); the name
                # filter, if also set, narrows further on top of that.
                $testPidRaw = $env:WINDOWSWEEPER_TEST_PID
                if ($testPidRaw) {
                    $parsedPid = 0
                    if ([int]::TryParse($testPidRaw, [ref]$parsedPid)) {
                        $rows = @($rows | Where-Object { $_.RowType -ne 'Workbook' -or $_.Key.Pid -eq $parsedPid })
                    }
                }
                $testFilter = $env:WINDOWSWEEPER_TEST_FILTER
                if ($testFilter) {
                    try { $rows = @($rows | Where-Object { $_.RowType -ne 'Workbook' -or $_.Name -match $testFilter }) } catch {}
                }
                return @{ Ok = $true; Message = ''; Rows = $rows }
            }

            # v0.4 Phase 2: Documents rows + ProtectedViewWindows rows.
            # Get-Snapshot pre-gates on WINWORD actually running before ever
            # calling this, so a $null Attach-Word here means either a rare
            # attach race or a non-ROT instance - report it rather than
            # crash, but never try to launch one.
            'EnumWord' {
                $word = Attach-Word
                if (-not $word) {
                    # Sol minor: surface an Info row (same contract as
                    # Outlook's "not running" row) rather than silently
                    # returning zero rows - a running-but-unreachable Word
                    # must never just vanish from the grid.
                    $msg = 'Word is running but not reachable - try again'
                    return @{ Ok = $true; Message = $msg; Rows = @([PSCustomObject]@{ App = 'Word'; RowType = 'Info'; Name = 'Word'; Detail = $msg; Status = 'Info'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }) }
                }
                $rows = @()
                $docs = @()
                try { $docs = @($word.Documents) } catch {
                    $rows += [PSCustomObject]@{ App = 'Word'; RowType = 'Busy'; Name = 'Word'; Detail = 'Busy - could not read document list'; Status = 'Busy'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }
                    return @{ Ok = $true; Message = ''; Rows = $rows }
                }
                foreach ($doc in $docs) {
                    try {
                        # Nick hit blank-Name rows in the wild: doc.Name can
                        # come back empty (not throw) for some live document
                        # states. Never emit a nameless row - skip it with a
                        # debug breadcrumb either way (throw or empty).
                        $name = $null
                        try { $name = $doc.Name } catch {}
                        if (-not $name) {
                            Write-DebugLog "EnumWord SKIP-NAMELESS-DOC (doc.Name threw or returned empty)"
                            continue
                        }
                        $fullName = $null
                        try { $fullName = $doc.FullName } catch {}
                        $hasPath = $false
                        try { $hasPath = [bool]$doc.Path } catch {}
                        # Fail CLOSED (Sol P1/cross-app hardening): default
                        # to unsaved so an unreadable .Saved labels the row
                        # 'UNSAVED', never 'Saved'.
                        $saved = $false
                        try { $saved = [bool]$doc.Saved } catch {}
                        $status = 'Saved'
                        if (-not $hasPath) { $status = 'Never saved' }
                        elseif (-not $saved) { $status = 'UNSAVED' }
                        $detail = '(never saved)'
                        if ($hasPath) { try { $detail = $doc.Path } catch {} }
                        # Hwnd from the OBJECT MODEL, not title matching (Sol
                        # P2): Word titles include the extension, unlike
                        # Excel, so no shared title-matching helper is needed
                        # or wanted here - Window.Hwnd kills the problem
                        # entirely.
                        $hwndVal = 0
                        try { $hwndVal = [int64]$doc.Windows.Item(1).Hwnd } catch {}
                        $rows += [PSCustomObject]@{
                            App = 'Word'; RowType = 'Document'; Name = $name; Detail = $detail; Status = $status
                            Key = @{ Kind = 'WordName'; Name = $name; FullName = $fullName }
                            Hwnd = $hwndVal; CanSave = $true; CanClose = $true
                        }
                    } catch { continue }
                }
                $pvws = @()
                try { $pvws = @($word.ProtectedViewWindows) } catch {}
                for ($i = 0; $i -lt $pvws.Count; $i++) {
                    try {
                        $pvw = $pvws[$i]
                        $caption = ''
                        try { $caption = $pvw.Caption } catch {}
                        # Nick hit a generic 'Protected View document'
                        # placeholder in the wild: .SourceName is not a real
                        # ProtectedViewWindow property (always threw,
                        # silently swallowed). .Document.Name (the limited
                        # Document proxy PVW exposes) is the real filename -
                        # prefer it; fall back to Caption, then the generic
                        # placeholder only if both are empty.
                        $docName = ''
                        try { $docName = $pvw.Document.Name } catch {}
                        $srcName = $docName
                        if (-not $srcName) { $srcName = $caption }
                        if (-not $srcName) { $srcName = 'Protected View document' }
                        $rows += [PSCustomObject]@{
                            App = 'Word'; RowType = 'ProtectedView'; Name = $srcName; Detail = 'read-only preview'; Status = 'Protected View'
                            Key = @{ Kind = 'Index'; Index = $i; VerifyCaption = $caption }; Hwnd = 0; CanSave = $false; CanClose = $true
                        }
                    } catch { continue }
                }
                # v0.4 test hook - same name-regex contract as Outlook's
                # Inspector filtering, so fixture-fenced GUI tests can scope
                # to only their own Word documents.
                $testFilter = $env:WINDOWSWEEPER_TEST_FILTER
                if ($testFilter) {
                    try { $rows = @($rows | Where-Object { $_.Name -match $testFilter }) } catch {}
                }
                return @{ Ok = $true; Message = ''; Rows = $rows }
            }

            'EnumOutlook' {
                $ol = Attach-Outlook
                if (-not $ol) {
                    $msg = 'Outlook not running'
                    try {
                        if (Get-Process -Name 'olk' -ErrorAction SilentlyContinue) {
                            $msg = 'New Outlook is running - it does not support automation. Open classic Outlook instead.'
                        }
                    } catch {}
                    return @{ Ok = $true; Message = $msg; Rows = @([PSCustomObject]@{ App = 'Outlook'; RowType = 'Info'; Name = 'Outlook'; Detail = $msg; Status = 'Info'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }) }
                }
                $rows = @()
                $inspectors = @()
                try { $inspectors = @($ol.Inspectors) } catch {
                    return @{ Ok = $true; Message = 'Outlook busy'; Rows = @([PSCustomObject]@{ App = 'Outlook'; RowType = 'Busy'; Name = 'Outlook'; Detail = 'Busy - a dialog may be open'; Status = 'Busy'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }) }
                }
                # Row-click fronting is Win32-only (no COM) - capture each
                # inspector's own top-level window once here, at enum time,
                # by matching its Caption against visible rctrl_renwnd32
                # windows. Explorer (folder) windows share this same class,
                # but their captions look nothing like an inspector's, so a
                # unique match naturally lands on the right window; ambiguous
                # or no match -> 0, which Invoke-RowActivate treats as a
                # silent no-op rather than guessing.
                $inspWins = @([SweeperWin32]::GetTopLevelWindows() | Where-Object { $_.ClassName -eq 'rctrl_renwnd32' -and $_.Visible })
                for ($i = 0; $i -lt $inspectors.Count; $i++) {
                    try {
                        $item = $inspectors[$i].CurrentItem
                        if (-not $item) { continue }
                        $rawSubj = ''
                        try { $rawSubj = $item.Subject } catch {}
                        $subj = $rawSubj
                        if (-not $subj) { $subj = '(no subject)' }
                        $mc = ''
                        try { $mc = $item.MessageClass } catch {}
                        $isReport = $mc -like 'REPORT.*'
                        # Fail CLOSED (Sol cross-app hardening): default to
                        # unsaved so an unreadable/missing .Saved never
                        # mislabels a row 'Saved' when the true state is
                        # unknown.
                        $saved = $false
                        try { if ($item.PSObject.Properties.Match('Saved').Count -gt 0) { $saved = [bool]$item.Saved } } catch {}
                        $status = 'Saved'
                        if (-not $saved) { $status = 'UNSAVED' }
                        $hwndVal = 0
                        $caption = ''
                        try { $caption = $inspectors[$i].Caption } catch {}
                        if ($caption) {
                            $normCaption = Get-NormalizedTitle $caption
                            $capMatches = @($inspWins | Where-Object { (Get-NormalizedTitle $_.Title) -eq $normCaption })
                            if ($capMatches.Count -eq 1) { $hwndVal = $capMatches[0].Hwnd.ToInt64() }
                        }
                        # Fallback: a unique window title starting with the
                        # item's own (normalized) Subject - only when the
                        # caption match above was not unique and there is a
                        # real subject to key on. Still ambiguous -> Hwnd=0,
                        # same as today; never guess.
                        if ($hwndVal -eq 0 -and $rawSubj) {
                            $normSubj = Get-NormalizedTitle $rawSubj
                            if ($normSubj) {
                                $subjMatches = @($inspWins | Where-Object { (Get-NormalizedTitle $_.Title).StartsWith($normSubj) })
                                if ($subjMatches.Count -eq 1) { $hwndVal = $subjMatches[0].Hwnd.ToInt64() }
                            }
                        }
                        $rows += [PSCustomObject]@{
                            App = 'Outlook'; RowType = 'Inspector'; Name = $subj; Detail = $mc; Status = $status
                            Key = (Get-ItemKey $item $i); Hwnd = $hwndVal; CanSave = (-not $isReport); CanClose = $true
                        }
                    } catch { continue }
                }
                $explorers = @()
                try { $explorers = @($ol.Explorers) } catch {}
                for ($i = 0; $i -lt $explorers.Count; $i++) {
                    try {
                        $exp = $explorers[$i]
                        $caption = ''
                        try { $caption = $exp.Caption } catch {}
                        if (-not $caption) { $caption = 'Outlook folder window' }
                        $folderName = ''
                        try { $folderName = $exp.CurrentFolder.Name } catch {}
                        $detail = ''
                        if ($folderName) { $detail = "Folder: $folderName" }
                        # Same normalized-caption match as Inspectors -
                        # Explorer (folder) windows share the rctrl_renwnd32
                        # class. No Subject-style fallback (Explorer has no
                        # equivalent); ambiguous or no match -> Hwnd=0.
                        $expHwndVal = 0
                        $normExpCaption = Get-NormalizedTitle $caption
                        if ($normExpCaption) {
                            $expMatches = @($inspWins | Where-Object { (Get-NormalizedTitle $_.Title) -eq $normExpCaption })
                            if ($expMatches.Count -eq 1) { $expHwndVal = $expMatches[0].Hwnd.ToInt64() }
                        }
                        $rows += [PSCustomObject]@{
                            App = 'Outlook'; RowType = 'Explorer'; Name = $caption; Detail = $detail; Status = 'Window'
                            Key = @{ Kind = 'Index'; Index = $i; VerifyCaption = $caption }; Hwnd = $expHwndVal; CanSave = $false; CanClose = $true
                        }
                    } catch { continue }
                }
                # v0.3 test hook - zero effect unless the env var is set.
                # Explorer rows have no meaningful Subject to match, so all of
                # them drop when a filter is active (keeps automated GUI
                # testing from ever touching a real folder window).
                $testFilter = $env:WINDOWSWEEPER_TEST_FILTER
                if ($testFilter) {
                    try {
                        $rows = @($rows | Where-Object {
                            if ($_.RowType -eq 'Explorer') { return $false }
                            if ($_.RowType -eq 'Inspector') { return ($_.Name -match $testFilter) }
                            return $true
                        })
                    } catch {}
                }
                return @{ Ok = $true; Message = ''; Rows = $rows }
            }

            # v0.4 Phase 3: two independent paths. CDP path is config-gated,
            # never automatic - see the port-selection logic below for the
            # exact test-isolation contract (Sol P1 + panel P2). Window
            # fallback is the default when no CDP port is in play at all.
            'EnumChrome' {
                $rows = @()
                # Sol MAJOR fix: EITHER Chrome test env var being active puts
                # us into full exclusive test mode - settings-configured
                # chromeDebugPorts are ignored ENTIRELY the moment either var
                # is set, not just when TEST_PORT is. The original bug: a
                # TEST_PID-only run (no TEST_PORT) fell into the "normal"
                # else-branch below and still probed real configured ports,
                # so a machine with :9223 configured could expose/close real
                # automation tabs during what was meant to be a PID-fenced
                # fallback-only test. CDP probing now happens ONLY when
                # WINDOWSWEEPER_CHROME_DEBUG_PORT itself is set (and only
                # that one port); TEST_PID-only mode probes no CDP port at
                # all (window-fallback only, fenced to that PID - see below).
                $testPortRaw = $env:WINDOWSWEEPER_CHROME_DEBUG_PORT
                $testPidRaw = $env:WINDOWSWEEPER_CHROME_TEST_PID
                $chromeTestModeActive = (-not [string]::IsNullOrEmpty($testPortRaw)) -or (-not [string]::IsNullOrEmpty($testPidRaw))

                $ports = @()
                if ($chromeTestModeActive) {
                    if (-not [string]::IsNullOrEmpty($testPortRaw)) {
                        $parsedPort = 0
                        if ([int]::TryParse($testPortRaw, [ref]$parsedPort)) { $ports = @($parsedPort) }
                    }
                } else {
                    $ports = @($cmd.ChromeDebugPorts | Where-Object { $_ })
                }

                foreach ($port in $ports) {
                    $pages = Get-ChromeCdpPages -Port $port
                    foreach ($tab in $pages) {
                        try {
                            $title = "$($tab.title)".Trim()
                            if (-not $title) { $title = '(untitled tab)' }
                            $url = "$($tab.url)".Trim()
                            $detailUrl = $url
                            if ($detailUrl.Length -gt 80) { $detailUrl = $detailUrl.Substring(0, 77) + '...' }
                            $tabId = "$($tab.id)"
                            if (-not $tabId) { continue }
                            $rows += [PSCustomObject]@{
                                App = 'Chrome'; RowType = 'Tab'; Name = $title; Detail = $detailUrl; Status = 'Tab'
                                Key = @{ Kind = 'CdpTab'; Port = $port; Id = $tabId }
                                Hwnd = 0; CanSave = $false; CanClose = $true
                            }
                        } catch { continue }
                    }
                }

                # Window fallback: the default when no CDP ports are
                # configured at all (real usage) - once any port is in play,
                # the fallback is superseded for this call. In Chrome test
                # mode, the fallback is disabled UNLESS TEST_PID is also set
                # (Sol P1) - otherwise a CDP-only (TEST_PORT-only) test run
                # could also enumerate Nick's real Chrome windows via the
                # fallback path. TEST_PID-only mode (no TEST_PORT) lands here
                # too: $ports is empty (no CDP probing at all per the fix
                # above), so the fallback is what carries that test entirely.
                $fallbackAllowed = $false
                if (-not $chromeTestModeActive -and $ports.Count -eq 0) { $fallbackAllowed = $true }
                if ($chromeTestModeActive -and -not [string]::IsNullOrEmpty($testPidRaw)) { $fallbackAllowed = $true }

                if ($fallbackAllowed) {
                    $winList = @([SweeperWin32]::GetTopLevelWindows() | Where-Object { $_.ClassName -eq 'Chrome_WidgetWin_1' -and $_.Visible -and $_.Title })
                    # Sol hardening: keep a candidate only if its owning
                    # process image is actually chrome.exe - Electron apps
                    # (Claude, Cursor, etc.) share the same window class and
                    # would otherwise appear as fake "Chrome" rows.
                    $winList = @($winList | Where-Object { Test-ChromeOwnedWindow -Hwnd $_.Hwnd })
                    if (-not [string]::IsNullOrEmpty($testPidRaw)) {
                        # TEST_PID fences the fallback to only the fixture's
                        # own windows (mirrors the Excel WINDOWSWEEPER_TEST_PID
                        # pattern) - an unparsable value fences to nothing
                        # rather than falling through to "everything."
                        $parsedTestPid = 0
                        if ([int]::TryParse($testPidRaw, [ref]$parsedTestPid)) {
                            $winList = @($winList | Where-Object { $_.Pid -eq $parsedTestPid })
                        } else {
                            $winList = @()
                        }
                    }
                    foreach ($w in $winList) {
                        $rows += [PSCustomObject]@{
                            App = 'Chrome'; RowType = 'BrowserWindow'; Name = $w.Title; Detail = 'browser window - tabs unknown'; Status = 'Window'
                            Key = @{ Kind = 'ChromeWindow'; Hwnd = $w.Hwnd.ToInt64() }
                            Hwnd = $w.Hwnd.ToInt64(); CanSave = $false; CanClose = $true
                        }
                    }
                }

                # v0.4 test hook - same name-regex contract as Word/Outlook,
                # so fixture-fenced GUI tests can scope to only fixture rows.
                $testFilter = $env:WINDOWSWEEPER_TEST_FILTER
                if ($testFilter) {
                    try { $rows = @($rows | Where-Object { $_.Name -match $testFilter }) } catch {}
                }
                return @{ Ok = $true; Message = ''; Rows = $rows }
            }

            'ExcelSaveClose' {
                $wb = Resolve-Workbook $cmd.Key
                if (-not $wb) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                $hasPath = $false
                try { $hasPath = [bool]$wb.Path } catch {}
                if (-not $hasPath) { return @{ Ok = $false; NeedsSaveAs = $true; Message = 'Never saved - use Save As' } }
                try {
                    $wb.Save()
                    $wb.Close($false)
                    return @{ Ok = $true; Message = 'Saved and closed' }
                } catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            # Never-saved-workbook flow only: blocks on purpose for up to 180s
            # (per Invoke-ComTask's TimeoutMs for this op) while Nick uses the
            # native Save As dialog.
            'ExcelSaveAs' {
                $wb = Resolve-Workbook $cmd.Key
                if (-not $wb) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                try {
                    $app = $wb.Application
                    $wb.Activate()
                    $shown = $app.Dialogs(5).Show()
                    if (-not $shown) { return @{ Ok = $true; Cancelled = $true; Message = 'Save As cancelled' } }
                    $wb.Close($false)
                    return @{ Ok = $true; Message = 'Saved and closed' }
                } catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            'ExcelClose' {
                $wb = Resolve-Workbook $cmd.Key
                if (-not $wb) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                # Re-read live save-state right before discarding - the row's
                # cached Status can be stale if Nick edited the workbook after
                # the last Refresh without clicking Refresh again.
                $hasPath = $false
                try { $hasPath = [bool]$wb.Path } catch {}
                # Fail CLOSED (Sol cross-app hardening): an unreadable .Saved
                # must force the discard-confirm gate, not skip it - the old
                # default-true let an unconfirmed Close($false) through on
                # read failure.
                $saved = $false
                try { $saved = [bool]$wb.Saved } catch {}
                $liveUnsaved = (-not $hasPath) -or (-not $saved)
                if ($liveUnsaved -and -not $cmd.ConfirmedDiscard) {
                    return @{ Ok = $false; NeedsConfirm = $true; Message = 'Has unsaved changes' }
                }
                try { $wb.Close($false); return @{ Ok = $true; Message = 'Closed without saving' } }
                catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            'ExcelActivate' {
                $wb = Resolve-Workbook $cmd.Key
                if (-not $wb) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                try { $wb.Activate(); $wb.Application.Visible = $true } catch {}
                return @{ Ok = $true; Message = '' }
            }

            'WordSaveClose' {
                $word = Attach-Word
                if (-not $word) { return @{ Ok = $false; Message = 'Word not reachable' } }
                $resolved = Resolve-WordDocument $word $cmd.Key
                if (-not $resolved.Doc) {
                    if ($resolved.Ambiguous) { return @{ Ok = $false; Message = 'Multiple documents look alike - refresh and try again' } }
                    return @{ Ok = $false; Message = 'No longer open - refresh and try again' }
                }
                $doc = $resolved.Doc
                $hasPath = $false
                try { $hasPath = [bool]$doc.Path } catch {}
                # Guard FIRST (Sol P1): doc.Save() on a pathless document
                # BLOCKS inside this 10s op showing Word's own Save As dialog
                # - exactly the ExcelSaveClose pathless guard, ported.
                if (-not $hasPath) { return @{ Ok = $false; NeedsSaveAs = $true; Message = 'Never saved - use Save As' } }
                try {
                    $doc.Save()
                    # Sol P1 (close-after-failed-save): a DocumentBeforeSave
                    # add-in handler can CANCEL Save() without throwing -
                    # re-verify the save actually landed before closing, or
                    # Close(0) would silently discard the edits.
                    $postPath = $false
                    try { $postPath = [bool]$doc.Path } catch {}
                    $postSaved = $false
                    try { $postSaved = [bool]$doc.Saved } catch {}
                    if (-not $postPath -or -not $postSaved) {
                        return @{ Ok = $false; Message = 'Save did not complete - document kept open' }
                    }
                    $doc.Close([ref]0)
                    return @{ Ok = $true; Message = 'Saved and closed' }
                } catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            # Never-saved-document flow only: blocks on purpose for up to
            # 180s (per Invoke-ComTask's TimeoutMs for this op) while Nick
            # uses the native Save As dialog.
            'WordSaveAs' {
                $word = Attach-Word
                if (-not $word) { return @{ Ok = $false; Message = 'Word not reachable' } }
                $resolved = Resolve-WordDocument $word $cmd.Key
                if (-not $resolved.Doc) {
                    if ($resolved.Ambiguous) { return @{ Ok = $false; Message = 'Multiple documents look alike - refresh and try again' } }
                    return @{ Ok = $false; Message = 'No longer open - refresh and try again' }
                }
                $doc = $resolved.Doc
                try {
                    $doc.Activate()
                    $shown = $word.Dialogs.Item(84).Show()  # wdDialogFileSaveAs
                    if (-not $shown) { return @{ Ok = $true; Cancelled = $true; Message = 'Save As cancelled' } }
                    # Sol P1: only -1 (OK/committed) means the dialog
                    # actually drove a save. Show() can return other nonzero
                    # values too (e.g. -2 when the user chose Close from the
                    # dialog) - those are NOT a committed save and must not
                    # fall through to Close(0), which would discard the edits.
                    if ($shown -ne -1) {
                        return @{ Ok = $false; Message = 'Save did not complete - document kept open' }
                    }
                    # Re-verify the save actually landed before closing - a
                    # DocumentBeforeSave add-in handler can cancel the
                    # dialog's internal save without throwing.
                    $postPath = $false
                    try { $postPath = [bool]$doc.Path } catch {}
                    $postSaved = $false
                    try { $postSaved = [bool]$doc.Saved } catch {}
                    if (-not $postPath -or -not $postSaved) {
                        return @{ Ok = $false; Message = 'Save did not complete - document kept open' }
                    }
                    $doc.Close([ref]0)
                    return @{ Ok = $true; Message = 'Saved and closed' }
                } catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            'WordClose' {
                $word = Attach-Word
                if (-not $word) { return @{ Ok = $false; Message = 'Word not reachable' } }
                $resolved = Resolve-WordDocument $word $cmd.Key
                if (-not $resolved.Doc) {
                    if ($resolved.Ambiguous) { return @{ Ok = $false; Message = 'Multiple documents look alike - refresh and try again' } }
                    return @{ Ok = $false; Message = 'No longer open - refresh and try again' }
                }
                $doc = $resolved.Doc
                # Re-read live save-state right before discarding - same
                # staleness protection as ExcelClose.
                $hasPath = $false
                try { $hasPath = [bool]$doc.Path } catch {}
                # Fail CLOSED (Sol cross-app hardening) - see ExcelClose.
                $saved = $false
                try { $saved = [bool]$doc.Saved } catch {}
                $liveUnsaved = (-not $hasPath) -or (-not $saved)
                if ($liveUnsaved -and -not $cmd.ConfirmedDiscard) {
                    return @{ Ok = $false; NeedsConfirm = $true; Message = 'Has unsaved changes' }
                }
                try { $doc.Close([ref]0); return @{ Ok = $true; Message = 'Closed without saving' } }  # wdDoNotSaveChanges, promptless
                catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            'WordProtectedViewClose' {
                $word = Attach-Word
                if (-not $word) { return @{ Ok = $false; Message = 'Word not reachable' } }
                $pvw = Resolve-ProtectedView $word $cmd.Key
                if (-not $pvw) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                try { $pvw.Close(); return @{ Ok = $true; Message = 'Closed' } }
                catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            'OutlookSaveClose' {
                $ol = Attach-Outlook
                if (-not $ol) { return @{ Ok = $false; Message = 'Outlook not reachable' } }
                $resolved = Resolve-Inspector $ol $cmd.Key
                if (-not $resolved.Item) {
                    if ($resolved.Ambiguous) { return @{ Ok = $false; Message = 'Multiple windows look alike - refresh and try again' } }
                    return @{ Ok = $false; Message = 'No longer open - refresh and try again' }
                }
                $item = $resolved.Item
                try {
                    $mc = ''
                    try { $mc = $item.MessageClass } catch {}
                    if ($mc -like 'REPORT.*') { return @{ Ok = $false; Message = 'Report items cannot be saved' } }
                    $item.Save()
                    $note = ''
                    try { if ($item.Sent -eq $false) { $note = ' (saved to Drafts)' } } catch {}
                    $item.Close(1)
                    return @{ Ok = $true; Message = "Saved and closed$note" }
                } catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            'OutlookClose' {
                $ol = Attach-Outlook
                if (-not $ol) { return @{ Ok = $false; Message = 'Outlook not reachable' } }
                $resolved = Resolve-Inspector $ol $cmd.Key
                if (-not $resolved.Item) {
                    if ($resolved.Ambiguous) { return @{ Ok = $false; Message = 'Multiple windows look alike - refresh and try again' } }
                    return @{ Ok = $false; Message = 'No longer open - refresh and try again' }
                }
                $item = $resolved.Item
                # Re-read live save-state right before discarding - same
                # staleness protection as ExcelClose.
                # Fail CLOSED (Sol cross-app hardening) - see ExcelClose.
                $saved = $false
                try { if ($item.PSObject.Properties.Match('Saved').Count -gt 0) { $saved = [bool]$item.Saved } } catch {}
                if (-not $saved -and -not $cmd.ConfirmedDiscard) {
                    return @{ Ok = $false; NeedsConfirm = $true; Message = 'Has unsaved changes' }
                }
                try { $item.Close(1); return @{ Ok = $true; Message = 'Closed without saving' } }
                catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            'OutlookActivate' {
                $ol = Attach-Outlook
                if (-not $ol) { return @{ Ok = $false; Message = 'Outlook not reachable' } }
                $resolved = Resolve-Inspector $ol $cmd.Key
                if (-not $resolved.Item) {
                    if ($resolved.Ambiguous) { return @{ Ok = $false; Message = 'Multiple windows look alike - refresh and try again' } }
                    return @{ Ok = $false; Message = 'No longer open - refresh and try again' }
                }
                try { $resolved.Item.GetInspector.Activate() } catch {}
                return @{ Ok = $true; Message = '' }
            }

            'ExplorerClose' {
                $ol = Attach-Outlook
                if (-not $ol) { return @{ Ok = $false; Message = 'Outlook not reachable' } }
                $explorers = @()
                try { $explorers = @($ol.Explorers) } catch {}
                # Live re-check: closing the last folder window would exit Outlook's UI.
                if ($explorers.Count -le 1) { return @{ Ok = $false; Message = 'Cannot close the last Outlook folder window' } }
                $exp = Resolve-Explorer $ol $cmd.Key
                if (-not $exp) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                try { $exp.Close(); return @{ Ok = $true; Message = 'Closed' } }
                catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            # v0.4 Phase 3: no confirm/live-recheck gate here the way
            # Excel/Word/Outlook Close ops have one - Chrome has no save
            # prompts at all, so there is no "changed since refresh" state
            # to catch. Liveness is just "did the close succeed" (a 404 =
            # already gone, reported gracefully, not an error).
            'ChromeTabClose' {
                $port = $cmd.Key.Port
                $id = $cmd.Key.Id
                if (-not $port -or -not $id) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                $r = Invoke-ChromeCdpGet -Url "http://127.0.0.1:$port/json/close/$id"
                if ($r.Ok) { return @{ Ok = $true; Message = 'Closed' } }
                if ($r.StatusCode -eq 404) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                return @{ Ok = $false; Message = 'Could not close tab' }
            }

            'ChromeTabActivate' {
                $port = $cmd.Key.Port
                $id = $cmd.Key.Id
                if (-not $port -or -not $id) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                $r = Invoke-ChromeCdpGet -Url "http://127.0.0.1:$port/json/activate/$id"
                if (-not $r.Ok) {
                    if ($r.StatusCode -eq 404) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                    return @{ Ok = $false; Message = 'Could not activate tab' }
                }
                # Best-effort window fronting only (Sol P2): CDP supplies no
                # HWND, so front ONLY when exactly one visible
                # Chrome_WidgetWin_1 title matches this tab's title - same
                # "Name - <app>" shape EnumExcel already matches on.
                # Ambiguous or no match -> activate without fronting, never
                # guess which window to bring forward. Sol MINOR: the
                # candidate filter must also confirm chrome.exe ownership -
                # class+title alone could let a uniquely-matching Electron
                # window (same class, coincidentally identical title) get
                # foregrounded instead.
                try {
                    $tabTitle = "$($cmd.TabTitle)"
                    if ($tabTitle) {
                        $titleMatches = @([SweeperWin32]::GetTopLevelWindows() | Where-Object {
                            $_.ClassName -eq 'Chrome_WidgetWin_1' -and $_.Visible -and
                            ($_.Title -eq $tabTitle -or $_.Title.StartsWith("$tabTitle - ")) -and
                            (Test-ChromeOwnedWindow -Hwnd $_.Hwnd)
                        })
                        if ($titleMatches.Count -eq 1) {
                            [SweeperWin32]::ShowWindow($titleMatches[0].Hwnd, [SweeperWin32]::SW_RESTORE) | Out-Null
                            [SweeperWin32]::SetForegroundWindow($titleMatches[0].Hwnd) | Out-Null
                        }
                    }
                } catch {}
                return @{ Ok = $true; Message = 'Activated' }
            }

            # Pure Win32 - kept as a worker op only for routing uniformity
            # with every other Close op (it cannot hang, so there is no
            # watchdog benefit, but Invoke-BatchOp/Invoke-RowClose need a
            # single consistent call shape).
            'ChromeWindowClose' {
                $hwndRaw = $cmd.Key.Hwnd
                if (-not $hwndRaw) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                $h = [IntPtr]$hwndRaw
                if (-not [SweeperWin32]::IsWindow($h)) { return @{ Ok = $false; Message = 'No longer open - refresh and try again' } }
                # Sol hardening: re-confirm the owning process is actually
                # chrome.exe at ACTION time, not just at enum time - a stale
                # or reused hwnd (the window closed and the OS handed that
                # numeric value to an unrelated Electron app) must never
                # reach WM_CLOSE just because it once belonged to Chrome.
                if (-not (Test-ChromeOwnedWindow -Hwnd $h)) {
                    return @{ Ok = $false; Message = 'No longer open - refresh and try again' }
                }
                # v0.4 Phase 3 test isolation: when a fixture PID is pinned,
                # refuse any hwnd whose LIVE pid differs - re-checked here at
                # action time (not just at enum time), so a stale key can
                # never reach a real Chrome window.
                $testPidRaw = $env:WINDOWSWEEPER_CHROME_TEST_PID
                if (-not [string]::IsNullOrEmpty($testPidRaw)) {
                    $parsedTestPid = 0
                    if ([int]::TryParse($testPidRaw, [ref]$parsedTestPid)) {
                        [uint32]$actualPid = 0
                        [void][SweeperWin32]::GetWindowThreadProcessId($h, [ref]$actualPid)
                        if ($actualPid -ne $parsedTestPid) {
                            return @{ Ok = $false; Message = 'Refused - window PID does not match the test fixture' }
                        }
                    }
                }
                try {
                    # Chrome may show its own "close all tabs?" prompt - not
                    # suppressed, per plan.
                    # Sol MAJOR: PostMessage's bool return was previously
                    # discarded and Ok=$true reported unconditionally - an
                    # access-denied or close-race would report success and
                    # inflate a batch "Closed" tally. Check the return value
                    # and surface the real Win32 error on failure.
                    $posted = [SweeperWin32]::PostMessage($h, [SweeperWin32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
                    if (-not $posted) {
                        $errCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
                        return @{ Ok = $false; Message = "Could not close window (Win32 error $errCode)" }
                    }
                    return @{ Ok = $true; Message = 'Close requested' }
                } catch { return @{ Ok = $false; Message = $_.Exception.Message } }
            }

            default { return @{ Ok = $false; Message = "Unknown op: $($cmd.Op)" } }
        }
    } catch {
        return @{ Ok = $false; Message = "Worker error: $($_.Exception.Message)" }
    }
}

# ===== Region 4: Snapshot orchestration =====

# v0.4 Phase 0/1: app-agnostic - loops $script:KnownApps in display order so
# the label reads "12 Excel | 3 Word | 6 Outlook - 4 unsaved" with an app
# dropped entirely once it has zero rows, instead of a fixed two-app
# "N workbooks, N emails" sentence.
function Get-CountsSummary {
    param($rows)
    $countable = @($rows | Where-Object { $script:CountableRowTypes -contains $_.RowType })
    $parts = @()
    foreach ($app in $script:KnownApps) {
        $appRows = @($countable | Where-Object { $_.App -eq $app })
        if ($appRows.Count -eq 0) { continue }
        $parts += "$($appRows.Count) $app"
    }
    $text = $parts -join ' | '
    if (-not $text) { $text = 'Nothing open' }
    $unsaved = @($countable | Where-Object { $_.Status -eq 'UNSAVED' -or $_.Status -eq 'Never saved' })
    if ($unsaved.Count -gt 0) { $text += " - $($unsaved.Count) unsaved" }
    return $text
}

# Each app is enumerated through its own Invoke-ComTask call, so a busy
# Outlook never blocks Excel results (or vice versa). Apps is a string array
# (v0.4 Phase 0 - replaces the old 'Both'/'Excel'/'Outlook' -Filter string so
# Get-Snapshot generalizes to N apps); Word/Chrome are valid entries with no
# case below yet - they simply contribute no rows until their Phase 2/3
# adapters land.
function Get-Snapshot {
    param([string[]]$Apps, $OwnerForm, [bool]$ForceOutlookRetry = $false)

    $rows = @()
    $busyApps = @()
    $messages = @()

    if ($Apps -contains 'Excel') {
        if (-not (Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue)) {
            # Not running - nothing to enumerate, nothing to report as busy.
        } elseif (-not (Test-AppResponsive -ProcessName 'EXCEL' -TimeoutMs 1000)) {
            $busyApps += 'Excel'
            $rows += [PSCustomObject]@{ App = 'Excel'; RowType = 'Busy'; Name = 'Excel'; Detail = 'Not responding - a dialog may be open'; Status = 'Busy'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }
        } else {
            # v0.3: 15s, down from 20s now the worker caches its Application
            # attachment (see Attach-ExcelInstances) - still generous because
            # cross-apartment marshaling into the background runspace is
            # measured 10-40x slower per call than the same call from the
            # main thread, and a from-scratch enumeration touches every
            # open workbook.
            $r = Invoke-ComTask -Command @{ Op = 'EnumExcel' } -TimeoutMs 15000 -OwnerForm $OwnerForm -StatusText 'Scanning Excel...'
            if ($r.TimedOut) {
                $busyApps += 'Excel'
                $rows += [PSCustomObject]@{ App = 'Excel'; RowType = 'Busy'; Name = 'Excel'; Detail = 'Not responding - a dialog may be open'; Status = 'Busy'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }
            } elseif (-not $r.Ok) {
                $messages += "Excel: $($r.Message)"
            } else {
                $rows += $r.Rows
                if (@($r.Rows | Where-Object { $_.RowType -eq 'Busy' }).Count -gt 0) { $busyApps += 'Excel' }
            }
        }
    }

    # A close request that arrived during Excel enumeration should not go on
    # to start a fresh Word wait too - bail before ever calling
    # Invoke-ComTask again.
    if ($script:ClosingRequested) {
        Write-DebugLog "Get-Snapshot ABORT-before-word (closing)"
        return @{ Rows = $rows; BusyApps = @($busyApps | Select-Object -Unique); Messages = $messages }
    }

    if ($Apps -contains 'Word') {
        # v0.4 Phase 2: same pre-flight + watchdog shape as Excel - no
        # Outlook-style backoff timer, since Word has no equivalent history
        # of a blocking-dialog tax on every subsequent action.
        if (-not (Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue)) {
            # Not running - nothing to enumerate, nothing to report as busy.
        } elseif (-not (Test-AppResponsive -ProcessName 'WINWORD' -TimeoutMs 1000)) {
            $busyApps += 'Word'
            $rows += [PSCustomObject]@{ App = 'Word'; RowType = 'Busy'; Name = 'Word'; Detail = 'Not responding - a dialog may be open'; Status = 'Busy'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }
        } else {
            $r = Invoke-ComTask -Command @{ Op = 'EnumWord' } -TimeoutMs 15000 -OwnerForm $OwnerForm -StatusText 'Scanning Word...'
            if ($r.TimedOut) {
                $busyApps += 'Word'
                $rows += [PSCustomObject]@{ App = 'Word'; RowType = 'Busy'; Name = 'Word'; Detail = 'Not responding - a dialog may be open'; Status = 'Busy'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }
            } elseif (-not $r.Ok) {
                $messages += "Word: $($r.Message)"
            } else {
                $rows += $r.Rows
                if (@($r.Rows | Where-Object { $_.RowType -eq 'Busy' }).Count -gt 0) { $busyApps += 'Word' }
            }
        }
    }

    # Same bail-out shape again, for the Word -> Outlook boundary.
    if ($script:ClosingRequested) {
        Write-DebugLog "Get-Snapshot ABORT-before-outlook (closing)"
        return @{ Rows = $rows; BusyApps = @($busyApps | Select-Object -Unique); Messages = $messages }
    }

    if ($Apps -contains 'Outlook') {
        # v0.3 backoff: once Outlook is found busy/timed-out, skip actually
        # touching it for 60s - just keep showing a busy placeholder. This is
        # the fix for the +20s tax that was landing on every single action
        # (auto-refresh re-running EnumOutlook right after, against an
        # Outlook that is still sitting on the same blocking dialog). Only an
        # explicit Refresh-button click (ForceOutlookRetry) bypasses it.
        if (-not $ForceOutlookRetry -and (Get-Date) -lt $script:OutlookBackoffUntil) {
            $busyApps += 'Outlook'
            $rows += [PSCustomObject]@{ App = 'Outlook'; RowType = 'Busy'; Name = 'Outlook'; Detail = 'Outlook busy - will retry on Refresh'; Status = 'Busy'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }
        }
        # Explicit gate, checked before ever touching Outlook: Refresh must
        # never be what starts a new Outlook process. Attach-Outlook itself
        # is GetActiveObject-only (no launch fallback), but this also saves
        # the cost of a full COM round-trip whenever Outlook simply is not up.
        elseif (-not (Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)) {
            $script:OutlookBackoffUntil = [datetime]::MinValue
            $msg = 'Outlook is not running'
            try {
                if (Get-Process -Name 'olk' -ErrorAction SilentlyContinue) {
                    $msg = 'New Outlook is running - it does not support automation. Open classic Outlook instead.'
                }
            } catch {}
            $rows += [PSCustomObject]@{ App = 'Outlook'; RowType = 'Info'; Name = 'Outlook'; Detail = $msg; Status = 'Info'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }
        } elseif (-not (Test-AppResponsive -ProcessName 'OUTLOOK' -TimeoutMs 1000)) {
            $busyApps += 'Outlook'
            $script:OutlookBackoffUntil = (Get-Date).AddSeconds(60)
            $rows += [PSCustomObject]@{ App = 'Outlook'; RowType = 'Busy'; Name = 'Outlook'; Detail = 'Not responding - a dialog may be open'; Status = 'Busy'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }
        } else {
            $r = Invoke-ComTask -Command @{ Op = 'EnumOutlook' } -TimeoutMs 10000 -OwnerForm $OwnerForm -StatusText 'Scanning Outlook...'
            if ($r.TimedOut) {
                $busyApps += 'Outlook'
                $script:OutlookBackoffUntil = (Get-Date).AddSeconds(60)
                $rows += [PSCustomObject]@{ App = 'Outlook'; RowType = 'Busy'; Name = 'Outlook'; Detail = 'Not responding - a dialog may be open'; Status = 'Busy'; Key = $null; Hwnd = 0; CanSave = $false; CanClose = $false }
            } elseif (-not $r.Ok) {
                $script:OutlookBackoffUntil = [datetime]::MinValue
                $messages += "Outlook: $($r.Message)"
            } else {
                $rows += $r.Rows
                $stillBusy = @($r.Rows | Where-Object { $_.RowType -eq 'Busy' }).Count -gt 0
                if ($stillBusy) {
                    $busyApps += 'Outlook'
                    $script:OutlookBackoffUntil = (Get-Date).AddSeconds(60)
                } else {
                    $script:OutlookBackoffUntil = [datetime]::MinValue
                }
            }
        }
    }

    # Same bail-out shape again, for the Outlook -> Chrome boundary.
    if ($script:ClosingRequested) {
        Write-DebugLog "Get-Snapshot ABORT-before-chrome (closing)"
        return @{ Rows = $rows; BusyApps = @($busyApps | Select-Object -Unique); Messages = $messages }
    }

    if ($Apps -contains 'Chrome') {
        # v0.4 Phase 3: no Test-AppResponsive/busy-row gating here - the CDP
        # path is a stateless HTTP call (its own 2s-per-port timeout is the
        # ceiling) and the window fallback is pure Win32 enumeration;
        # neither has an equivalent to Excel/Outlook's "COM call parked
        # behind a modal dialog" busy state. A dead/unconfigured port just
        # yields zero CDP rows - the window fallback (if not excluded by
        # test isolation) still shows. chrome.exe not running -> nothing to
        # enumerate, no Info row needed (Chrome, unlike Outlook, has no
        # "wrong app is running" ambiguity to explain).
        if (Get-Process -Name 'chrome' -ErrorAction SilentlyContinue) {
            $chromePorts = @()
            try { $chromePorts = @((Get-Settings).chromeDebugPorts) } catch {}
            $r = Invoke-ComTask -Command @{ Op = 'EnumChrome'; ChromeDebugPorts = $chromePorts } -TimeoutMs 10000 -OwnerForm $OwnerForm -StatusText 'Scanning Chrome...'
            if ($r.TimedOut) {
                $messages += 'Chrome: timed out'
            } elseif (-not $r.Ok) {
                $messages += "Chrome: $($r.Message)"
            } else {
                $rows += $r.Rows
            }
        }
    }

    return @{ Rows = $rows; BusyApps = @($busyApps | Select-Object -Unique); Messages = $messages }
}

function Invoke-DiagMode {
    Initialize-StateDir
    $snap = Get-Snapshot -Apps $script:KnownApps -OwnerForm $null
    Write-Host ''
    Write-Host 'Window Sweeper v0.4.1 - Diagnostic Enumeration' -ForegroundColor Cyan
    Write-Host ('-' * 74)
    $fmt = '{0,-8} {1,-30} {2,-24} {3,-10}'
    Write-Host ($fmt -f 'App', 'Name', 'Detail', 'Status')
    Write-Host ('-' * 74)
    foreach ($row in $snap.Rows) {
        $name = $row.Name
        if ($name.Length -gt 30) { $name = $name.Substring(0, 27) + '...' }
        $detail = "$($row.Detail)"
        if ($detail.Length -gt 24) { $detail = $detail.Substring(0, 21) + '...' }
        Write-Host ($fmt -f $row.App, $name, $detail, $row.Status)
    }
    Write-Host ('-' * 74)
    Write-Host (Get-CountsSummary -rows $snap.Rows)
    if ($snap.BusyApps.Count -gt 0) { Write-Host ('Busy: ' + ($snap.BusyApps -join ', ')) -ForegroundColor Yellow }
    foreach ($m in $snap.Messages) { Write-Host $m -ForegroundColor Yellow }
}

if ($env:WINDOWSWEEPER_NOGUI -eq '1') {
    return
}

if ($Diag) {
    Invoke-DiagMode
    Close-ComRunspace
    [Environment]::Exit(0)
}

# ===== Region 4c: Visual assets (v0.4 Phase 1) =====
# GUI-only helpers - app icons, the generated window icon, and the shared
# rounded-rect path used by both the status pills and the .ico glyph. Never
# referenced from -Diag/NOGUI, so no cost when the GUI never shows.

# App Paths is normally not WOW64-redirected for these installers, but a
# WOW6432Node fallback costs nothing and covers the odd machine where it is.
function Get-AppIconExePath {
    param([string]$ExeName)
    $candidates = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$ExeName",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\$ExeName"
    )
    foreach ($regPath in $candidates) {
        try {
            if (-not (Test-Path $regPath)) { continue }
            $item = Get-Item -Path $regPath -ErrorAction Stop
            $exePath = $item.GetValue('')
            if ($exePath -and (Test-Path $exePath)) { return $exePath }
        } catch {}
    }
    return $null
}

# Builds the App=>20x20 bitmap cache once at startup. Any failure for a given
# app (no App Paths entry, ExtractAssociatedIcon throws, etc.) leaves that
# app's slot $null - Update-Grid and the CellPainting icon column both treat
# a $null image as "draw nothing," never an error.
function Initialize-AppIcons {
    $script:AppIcons = @{}
    $exeMap = @{ 'Excel' = 'EXCEL.EXE'; 'Outlook' = 'OUTLOOK.EXE'; 'Word' = 'WINWORD.EXE'; 'Chrome' = 'chrome.exe' }
    foreach ($appName in $script:KnownApps) {
        $img = $null
        try {
            $exeName = $exeMap[$appName]
            if ($exeName) {
                $exePath = Get-AppIconExePath -ExeName $exeName
                if ($exePath) {
                    $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
                    if ($icon) {
                        $bmp = New-Object System.Drawing.Bitmap 20, 20
                        $g = [System.Drawing.Graphics]::FromImage($bmp)
                        try {
                            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                            $g.DrawIcon($icon, (New-Object System.Drawing.Rectangle 0, 0, 20, 20))
                        } finally { $g.Dispose() }
                        $img = $bmp
                        $icon.Dispose()
                    }
                }
            }
        } catch {
            Write-DebugLog "Initialize-AppIcons EXCEPTION App=$appName Message=$($_.Exception.Message)"
            $img = $null
        }
        $script:AppIcons[$appName] = $img
    }
}

# Shared by the Status-pill CellPainting and the generated window icon -
# GDI+ has no native rounded-rect fill, so both build one from four corner
# arcs. The path's geometry always depends on the caller's rectangle, so
# (unlike the brushes/pens/fonts) this one is necessarily built per call.
function New-RoundedRectPath {
    param([System.Drawing.Rectangle]$Rect, [int]$Radius)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $Radius * 2
    if ($d -gt $Rect.Width) { $d = $Rect.Width }
    if ($d -gt $Rect.Height) { $d = $Rect.Height }
    if ($d -lt 1) {
        $path.AddRectangle($Rect)
        return $path
    }
    $path.AddArc($Rect.X, $Rect.Y, $d, $d, 180, 90)
    $path.AddArc($Rect.Right - $d, $Rect.Y, $d, $d, 270, 90)
    $path.AddArc($Rect.Right - $d, $Rect.Bottom - $d, $d, $d, 0, 90)
    $path.AddArc($Rect.X, $Rect.Bottom - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

# Raw BMP-format (DIB) ICO writer - not the newer PNG-compressed ICO entry
# format, which System.Drawing.Icon has historically been unreliable loading
# back for anything other than a 256px entry. A 32bpp DIB with alpha, bottom
# up, followed by an all-zero (fully opaque) AND mask is the traditional
# format every Windows/.NET Icon reader has always supported.
function Get-DibBytesFromBitmap {
    param([System.Drawing.Bitmap]$Bmp)
    $w = $Bmp.Width; $h = $Bmp.Height
    $rect = New-Object System.Drawing.Rectangle 0, 0, $w, $h
    $bmpData = $Bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    try {
        $stride = [Math]::Abs($bmpData.Stride)
        $totalBytes = $stride * $h
        $buf = New-Object byte[] $totalBytes
        [System.Runtime.InteropServices.Marshal]::Copy($bmpData.Scan0, $buf, 0, $totalBytes)
    } finally {
        $Bmp.UnlockBits($bmpData)
    }
    # .NET hands back rows top-down; a DIB wants them bottom-up - reverse row order.
    $out = New-Object byte[] $totalBytes
    for ($r = 0; $r -lt $h; $r++) {
        $srcOffset = $r * $stride
        $dstOffset = ($h - 1 - $r) * $stride
        [Array]::Copy($buf, $srcOffset, $out, $dstOffset, $stride)
    }
    return $out
}

function New-IcoImageData {
    param([System.Drawing.Bitmap]$Bmp)
    $w = $Bmp.Width; $h = $Bmp.Height
    # [byte[]] cast is load-bearing: a bare `return $out` from
    # Get-DibBytesFromBitmap crosses a PowerShell function-return boundary,
    # which unrolls the array and hands the caller a generic System.Object[]
    # instead of a true System.Byte[] - BinaryWriter.Write() below then binds
    # to the wrong overload and silently writes far fewer bytes than expected
    # (confirmed by direct repro), producing a truncated, unloadable .ico.
    $xorBytes = [byte[]](Get-DibBytesFromBitmap -Bmp $Bmp)
    $maskRowBytes = [int]([Math]::Ceiling($w / 32.0) * 4)
    $maskBytes = New-Object byte[] ($maskRowBytes * $h)  # all zero => fully opaque via the alpha channel
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter $ms
    $bw.Write([int32]40)                                            # biSize (BITMAPINFOHEADER)
    $bw.Write([int32]$w)                                            # biWidth
    $bw.Write([int32]($h * 2))                                      # biHeight (XOR + AND)
    $bw.Write([uint16]1)                                            # biPlanes
    $bw.Write([uint16]32)                                           # biBitCount
    $bw.Write([int32]0)                                             # biCompression (BI_RGB)
    $bw.Write([int32]($xorBytes.Length + $maskBytes.Length))        # biSizeImage
    $bw.Write([int32]0); $bw.Write([int32]0)                        # biXPelsPerMeter/biYPelsPerMeter
    $bw.Write([int32]0); $bw.Write([int32]0)                        # biClrUsed/biClrImportant
    $bw.Write($xorBytes)
    $bw.Write($maskBytes)
    $bw.Flush()
    $result = $ms.ToArray()
    $bw.Dispose(); $ms.Dispose()
    return $result
}

# The 32x32/16x16 glyph itself: a #217346-to-#0F6CBD split rounded square
# with a white 'W', deterministic (no randomness) so re-running this never
# produces a visibly different icon.
function New-SweeperIconBitmap {
    param([int]$Size)
    $bmp = New-Object System.Drawing.Bitmap $Size, $Size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.Clear([System.Drawing.Color]::Transparent)
        $rect = New-Object System.Drawing.Rectangle 0, 0, $Size, $Size
        $radius = [Math]::Max(2, [int]($Size * 0.18))
        $path = New-RoundedRectPath -Rect $rect -Radius $radius
        $g.SetClip($path)
        $b1 = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml('#217346'))
        $b2 = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml('#0F6CBD'))
        $half = [int]($Size / 2)
        $g.FillRectangle($b1, 0, 0, $Size, $half)
        $g.FillRectangle($b2, 0, $half, $Size, $Size - $half)
        $g.ResetClip()
        $fontSize = [Math]::Max(6, [Math]::Round($Size * 0.5, 1))
        $font = New-Object System.Drawing.Font('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold)
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = [System.Drawing.StringAlignment]::Center
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $g.DrawString('W', $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF 0, 0, $Size, $Size), $sf)
        $b1.Dispose(); $b2.Dispose(); $font.Dispose(); $path.Dispose()
    } finally {
        $g.Dispose()
    }
    return $bmp
}

# Generated once and shipped beside the ps1 - a no-op if the file already
# exists (repo copy, or a previous run already made one).
function New-SweeperIcoFile {
    param([string]$Path)
    if (Test-Path $Path) { return }
    try {
        $sizes = @(32, 16)
        $entries = @()
        foreach ($sz in $sizes) {
            $bmp = New-SweeperIconBitmap -Size $sz
            # [byte[]] cast - same function-return unrolling gotcha as above.
            $data = [byte[]](New-IcoImageData -Bmp $bmp)
            $entries += [PSCustomObject]@{ Size = $sz; Data = $data }
            $bmp.Dispose()
        }
        $ms = New-Object System.IO.MemoryStream
        $bw = New-Object System.IO.BinaryWriter $ms
        $bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$entries.Count)  # ICONDIR
        $offset = 6 + (16 * $entries.Count)
        foreach ($e in $entries) {
            $wByte = $e.Size; if ($wByte -ge 256) { $wByte = 0 }
            $bw.Write([byte]$wByte); $bw.Write([byte]$wByte)   # width/height (0 = 256)
            $bw.Write([byte]0); $bw.Write([byte]0)             # color count, reserved
            $bw.Write([uint16]1); $bw.Write([uint16]32)        # planes, bit count
            $bw.Write([uint32]$e.Data.Length)
            $bw.Write([uint32]$offset)
            $offset += $e.Data.Length
        }
        foreach ($e in $entries) { $bw.Write($e.Data) }
        $bw.Flush()
        [System.IO.File]::WriteAllBytes($Path, $ms.ToArray())
        $bw.Dispose(); $ms.Dispose()
    } catch {
        Write-DebugLog "New-SweeperIcoFile EXCEPTION $($_.Exception.Message)"
    }
}

# Top/bottom-bar buttons only - grid BUTTON columns (SaveClose/Close) stay
# system-drawn per plan, so this is never called on those.
function Set-FlatButtonStyle {
    param($Button)
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(232, 240, 254)
}

function New-HtmlBrush {
    param([string]$Hex)
    return New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml($Hex))
}

# ===== Region 5: UI =====

$script:SelectedApps = @($script:KnownApps)
$script:SkipDiscardConfirmThisSession = $false
$script:MainForm = $null
$script:Grid = $null
$script:Banner = $null
$script:LblBanner = $null
$script:BtnShowDialog = $null
$script:LblCounts = $null
$script:LblTimestamp = $null
$script:TopPanel = $null
$script:BottomPanel = $null
$script:ProgressBar = $null
$script:LblStatus = $null

# v0.3 scoped-refresh state: the last full merged snapshot (so a scoped
# refresh can replace just one app's rows and leave the other app's rows as
# they were), plus staleness/throttle timestamps for Form.Add_Activated.
$script:LastSnapshotRows = @()
$script:LastBusyApps = @()
$script:LastSnapshotTime = [datetime]::MinValue
$script:LastActivatedRefreshAt = [datetime]::MinValue

function Update-BusyBanner {
    param([string[]]$BusyApps)
    if (-not $BusyApps -or $BusyApps.Count -eq 0) {
        $script:Banner.Visible = $false
        return
    }
    $names = $BusyApps -join ' and '
    $text = "$names busy - a dialog may be open"
    if ($script:ZombieCount -ge 3) { $text += ' (multiple stuck operations - consider restarting Window Sweeper)' }
    $script:LblBanner.Text = $text
    $script:BtnShowDialog.Tag = $BusyApps[0]
    $script:Banner.Visible = $true
}

function Update-Grid {
    param($rows)
    Write-DebugLog "Update-Grid ENTER rows=$($rows.Count)"
    # Rows.Clear() + re-Add resets the viewport to row 0 - restore the
    # scroll position afterward so a click on a row below the fold doesn't
    # snap Nick back to the top of an 18+ row grid on every single action.
    $scrollRow = 0
    try { $scrollRow = $script:Grid.FirstDisplayedScrollingRowIndex } catch {}
    if ($scrollRow -lt 0) { $scrollRow = 0 }

    # A DataGridView mutated reentrantly (e.g. from inside its own click
    # event) can throw partway through Clear()/Add() - "reentrant call to
    # SetCurrentCellAddressCore". The structural fix is deferring every
    # grid-mutating refresh outside grid-event context (see
    # Request-DeferredRefresh); this try/catch is the second line of
    # defense so a failure here is loud (logged, real row count visible in
    # the log) instead of silently leaving the grid showing fewer rows than
    # the counts label reports.
    try { $script:Grid.Rows.Clear() } catch { Write-DebugLog "Update-Grid Clear EXCEPTION $($_.Exception.Message)" }
    $script:Grid.Tag = $rows
    $addedCount = 0
    foreach ($row in $rows) {
        try {
            $i = $script:Grid.Rows.Add()
            $gridRow = $script:Grid.Rows[$i]
            $icon = $null
            if ($script:AppIcons -and $script:AppIcons.ContainsKey($row.App)) { $icon = $script:AppIcons[$row.App] }
            $gridRow.Cells['Icon'].Value = $icon
            $gridRow.Cells['App'].Value = $row.App
            $gridRow.Cells['Name'].Value = $row.Name
            $gridRow.Cells['Detail'].Value = "$($row.Detail)"
            $gridRow.Cells['Status'].Value = $row.Status
            $saveText = '-'
            if ($row.CanSave) { $saveText = 'Save & Close' }
            $closeText = '-'
            if ($row.CanClose) { $closeText = 'X' }
            $gridRow.Cells['SaveClose'].Value = $saveText
            $gridRow.Cells['Close'].Value = $closeText
            $gridRow.Tag = $row
            # v0.4 Phase 1: per-row Status.Style coloring is superseded by the
            # CellPainting rounded-pill rendering (accent stripe + badge) -
            # no per-row Font/Brush allocation here anymore.
            $addedCount++
        } catch {
            Write-DebugLog "Update-Grid ROW-EXCEPTION name=$($row.Name) app=$($row.App) Message=$($_.Exception.Message) Stack=$($_.ScriptStackTrace)"
            continue
        }
    }

    if ($script:Grid.Rows.Count -gt 0) {
        if ($scrollRow -ge $script:Grid.Rows.Count) { $scrollRow = $script:Grid.Rows.Count - 1 }
        try { $script:Grid.FirstDisplayedScrollingRowIndex = $scrollRow } catch {}
    }
    Write-DebugLog "Update-Grid EXIT requested=$($rows.Count) added=$addedCount gridRows=$($script:Grid.Rows.Count)"
}

# Applies the three post-snapshot UI updates independently (each its own
# try/catch) so one failing step (e.g. a grid reentrancy exception) cannot
# silently prevent the others from running - a partial failure ends up
# logged and visible (counts/timestamp/banner still reflect reality) instead
# of leaving the grid and the rest of the UI disagreeing with each other.
# v0.4 Phase 1: rows are sorted app-grouped (fixed order Excel, Word,
# Outlook, Chrome; then by Name) right before they hit the grid - the single
# convergence point both Do-FullRefresh and Do-ScopedRefresh funnel through,
# so scoped (per-app) refreshes stay correctly grouped too.
function Sort-RowsForDisplay {
    param($Rows)
    return @($Rows | Sort-Object -Property @{ Expression = {
        if ($script:AppOrder.ContainsKey($_.App)) { $script:AppOrder[$_.App] } else { 99 }
    } }, @{ Expression = 'Name' })
}

function Update-RefreshedUi {
    param($Rows, $BusyApps)
    $sorted = $Rows
    try { $sorted = Sort-RowsForDisplay -Rows $Rows } catch { Write-DebugLog "Update-RefreshedUi Sort EXCEPTION $($_.Exception.Message)"; $sorted = $Rows }
    try { Update-Grid -rows $sorted } catch { Write-DebugLog "Update-RefreshedUi Update-Grid EXCEPTION $($_.Exception.Message) Stack=$($_.ScriptStackTrace)" }
    try { $script:LblCounts.Text = Get-CountsSummary -rows $Rows } catch { Write-DebugLog "Update-RefreshedUi LblCounts EXCEPTION $($_.Exception.Message)" }
    try { $script:LblTimestamp.Text = 'Last refreshed: ' + (Get-Date -Format 'h:mm:ss tt') } catch {}
    try { Update-BusyBanner -BusyApps $BusyApps } catch { Write-DebugLog "Update-RefreshedUi Update-BusyBanner EXCEPTION $($_.Exception.Message)" }
}

function Do-FullRefresh {
    param([bool]$ForceOutlookRetry = $false)
    Write-DebugLog "Do-FullRefresh ENTER ForceOutlookRetry=$ForceOutlookRetry Apps=$($script:SelectedApps -join ',') Closing=$script:ClosingRequested"
    if ($script:ClosingRequested) { Write-DebugLog "Do-FullRefresh ABORT (closing)"; return }
    $snap = Get-Snapshot -Apps $script:SelectedApps -OwnerForm $script:MainForm -ForceOutlookRetry $ForceOutlookRetry
    if ($script:ClosingRequested) { Write-DebugLog "Do-FullRefresh ABORT-post-snapshot (closing)"; return }
    $script:LastSnapshotRows = $snap.Rows
    $script:LastBusyApps = $snap.BusyApps
    $script:LastSnapshotTime = Get-Date
    Update-RefreshedUi -Rows $snap.Rows -BusyApps $snap.BusyApps
    Write-DebugLog "Do-FullRefresh EXIT rows=$($snap.Rows.Count)"
}

# v0.3: refreshes rows for ONE app only, merging them into whatever rows are
# already shown for the other app. This is the fix for the responsiveness
# complaint - every single row/batch action used to trigger a full both-apps
# refresh, which on a machine where Outlook is dialog-blocked meant paying an
# extra ~10-20s frozen wait after every click for an app that action never
# touched. Full refresh is reserved for startup, the Refresh button, and a
# filter change.
function Do-ScopedRefresh {
    param([string]$App, [bool]$ForceOutlookRetry = $false)
    Write-DebugLog "Do-ScopedRefresh ENTER App=$App ForceOutlookRetry=$ForceOutlookRetry Closing=$script:ClosingRequested"
    if ($script:ClosingRequested) { Write-DebugLog "Do-ScopedRefresh ABORT (closing)"; return }
    if ($script:SelectedApps -notcontains $App) { return }
    $snap = Get-Snapshot -Apps @($App) -OwnerForm $script:MainForm -ForceOutlookRetry $ForceOutlookRetry
    if ($script:ClosingRequested) { Write-DebugLog "Do-ScopedRefresh ABORT-post-snapshot (closing)"; return }
    $otherRows = @($script:LastSnapshotRows | Where-Object { $_.App -ne $App })
    $mergedRows = @($otherRows) + @($snap.Rows)
    $otherBusy = @($script:LastBusyApps | Where-Object { $_ -ne $App })
    $mergedBusy = @($otherBusy + $snap.BusyApps | Select-Object -Unique)

    $script:LastSnapshotRows = $mergedRows
    $script:LastBusyApps = $mergedBusy
    $script:LastSnapshotTime = Get-Date

    Update-RefreshedUi -Rows $mergedRows -BusyApps $mergedBusy
    Write-DebugLog "Do-ScopedRefresh EXIT App=$App otherRows=$($otherRows.Count) newRows=$($snap.Rows.Count) merged=$($mergedRows.Count)"
}

# Batch ops can touch rows from both apps (or just one) - scope-refresh only
# the app(s) actually present among the rows that were acted on.
function Do-RefreshForApps {
    param([string[]]$Apps)
    foreach ($a in @($Apps | Select-Object -Unique)) { Do-ScopedRefresh -App $a }
}

function Show-DiscardConfirm {
    param($Rows)
    if ($script:SkipDiscardConfirmThisSession) { return $true }
    $settings = Get-Settings
    if (-not $settings.confirmDiscard) { return $true }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Confirm Discard'
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MinimizeBox = $false
    $dlg.MaximizeBox = $false
    $dlg.Width = 440
    $dlg.Height = 260

    # v0.4 Phase 3: Chrome tabs have no save state, so lumping them under
    # "This will discard unsaved changes in:" would be misleading -
    # Close All and Discard can sweep a mix of genuinely-unsaved documents,
    # already-Saved rows, and Chrome tabs. Split into two labeled sections;
    # either can be empty (a single non-Chrome row still reads naturally).
    $chromeRows = @($Rows | Where-Object { $_.App -eq 'Chrome' })
    $otherRows = @($Rows | Where-Object { $_.App -ne 'Chrome' })
    $lines = @()
    if ($otherRows.Count -gt 0) {
        $lines += 'This will discard unsaved changes in:'
        $lines += ''
        $lines += @($otherRows | ForEach-Object { " - $($_.Name)" })
    }
    if ($chromeRows.Count -gt 0) {
        if ($lines.Count -gt 0) { $lines += '' }
        $lines += 'Browser tabs (nothing to save):'
        $lines += @($chromeRows | ForEach-Object { " - $($_.Name)" })
    }
    if ($lines.Count -eq 0) { $lines = @('This will close the selected item(s).') }

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = ($lines -join "`r`n")
    $lbl.Location = New-Object System.Drawing.Point(15, 15)
    $lbl.Size = New-Object System.Drawing.Size(400, 140)

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = "Don't ask again this session"
    $chk.Location = New-Object System.Drawing.Point(15, 165)
    $chk.AutoSize = $true

    $btnYes = New-Object System.Windows.Forms.Button
    $btnYes.Text = 'Discard'
    $btnYes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $btnYes.Location = New-Object System.Drawing.Point(225, 195)
    $btnYes.Width = 90

    $btnNo = New-Object System.Windows.Forms.Button
    $btnNo.Text = 'Cancel'
    $btnNo.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnNo.Location = New-Object System.Drawing.Point(325, 195)
    $btnNo.Width = 90

    $dlg.Controls.AddRange([System.Windows.Forms.Control[]]@($lbl, $chk, $btnYes, $btnNo))
    $dlg.AcceptButton = $btnNo
    $dlg.CancelButton = $btnNo

    $result = $dlg.ShowDialog($script:MainForm)
    if ($chk.Checked) { $script:SkipDiscardConfirmThisSession = $true }
    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

# v0.3.1: a grid rebuild (Update-Grid's Rows.Clear()/Add()) must never run
# synchronously from inside a DataGridView event that the grid itself is
# still dispatching (CellContentClick) - DataGridView is not reentrant for
# that, and it can throw "reentrant call to SetCurrentCellAddressCore"
# partway through the rebuild, which is the leading suspect for the grid
# ending up showing fewer rows than the counts label reports. A row-action
# handler now only arms this and returns; a one-shot Timer fires the actual
# scoped refresh a tick later, once the grid's own event dispatch has fully
# unwound. Two requests before the timer fires coalesce to a Both refresh
# rather than silently dropping one.
$script:PendingRefreshApp = $null
$script:DeferredRefreshTimer = $null

function Request-DeferredRefresh {
    param([string]$App)
    if ($script:PendingRefreshApp -and $script:PendingRefreshApp -ne $App) {
        $script:PendingRefreshApp = 'Both'
    } else {
        $script:PendingRefreshApp = $App
    }
    Write-DebugLog "Request-DeferredRefresh App=$App Pending=$script:PendingRefreshApp"
    if ($script:DeferredRefreshTimer) {
        $script:DeferredRefreshTimer.Stop()
        $script:DeferredRefreshTimer.Start()
    }
}

# v0.3: Win32-only, no COM, ever. Row-click used to fire a COM
# ExcelActivate/OutlookActivate through the slow cross-apartment worker with
# the form disabled - a multi-second freeze on every single click, on the
# single most frequent interaction in the tool (peeking at a row before
# deciding what to do with it). The Hwnd was already captured at enum time
# for exactly this; a click now just fronts it directly. ExcelActivate/
# OutlookActivate worker ops still exist but nothing calls them from here.
# v0.4 Phase 3 hardening (Sol): top-level twin of the worker's own
# Test-ChromeOwnedWindow (that one is scoped inside $script:WorkerScript's
# closure and unreachable from here) - Invoke-RowActivate's Win32Front path
# runs entirely on the UI thread with no worker round-trip, so it needs its
# own copy of the same owning-process check. Chrome_WidgetWin_1 is shared
# by Electron apps (Claude, Cursor, etc.); drop-on-failure.
function Test-ChromeOwnedWindow {
    param([IntPtr]$Hwnd)
    try {
        [uint32]$winPid = 0
        [void][SweeperWin32]::GetWindowThreadProcessId($Hwnd, [ref]$winPid)
        if ($winPid -eq 0) { return $false }
        $proc = Get-Process -Id $winPid -ErrorAction Stop
        return ($proc.ProcessName -eq 'chrome')
    } catch {
        return $false
    }
}

# v0.4 Phase 0: which (App, RowType) combos are even clickable, and how, now
# comes from the routing table rather than a hardcoded RowType whitelist -
# an unrecognized combo (or a future CdpActivate style not implemented yet)
# no-ops with a logged warning instead of guessing. Rows with no Hwnd
# captured at enum time (Busy/Info/Explorer) still silently no-op exactly as
# before - the routing lookup just adds a debug breadcrumb on top.
function Invoke-RowActivate {
    param($Row)
    $route = Get-RoutingEntry -App $Row.App -RowType $Row.RowType
    if (-not $route) {
        Write-DebugLog "Invoke-RowActivate NO-ROUTE App=$($Row.App) RowType=$($Row.RowType)"
        return
    }
    if ($route.ActivateStyle -eq 'Win32Front') {
        if (-not $Row.Hwnd -or $Row.Hwnd -eq 0) { return }  # no unique window captured - silent no-op
        $h = [IntPtr]$Row.Hwnd
        # Scoped strictly to Chrome BrowserWindow rows so none of this ever
        # touches Excel/Word/Outlook fronting.
        if ($Row.App -eq 'Chrome' -and $Row.RowType -eq 'BrowserWindow') {
            # Sol hardening: re-confirm the owning process is actually
            # chrome.exe at ACTION time, unconditionally (not just under
            # test isolation) - a stale/reused hwnd (the window closed and
            # the OS handed that numeric value to an unrelated Electron
            # app) must never be fronted just because it once belonged to
            # Chrome.
            if (-not (Test-ChromeOwnedWindow -Hwnd $h)) {
                Write-DebugLog "Invoke-RowActivate REFUSED Chrome BrowserWindow hwnd is not chrome.exe-owned"
                return
            }
            # v0.4 Phase 3 test isolation: while WINDOWSWEEPER_CHROME_TEST_PID
            # is pinned, ALSO refuse any hwnd whose live PID does not match
            # the fixture - re-checked here, not just at enum time,
            # mirroring ChromeWindowClose's own re-check.
            $testPidRaw = $env:WINDOWSWEEPER_CHROME_TEST_PID
            if (-not [string]::IsNullOrEmpty($testPidRaw)) {
                $parsedTestPid = 0
                if ([int]::TryParse($testPidRaw, [ref]$parsedTestPid)) {
                    [uint32]$actualPid = 0
                    [void][SweeperWin32]::GetWindowThreadProcessId($h, [ref]$actualPid)
                    if ($actualPid -ne $parsedTestPid) {
                        Write-DebugLog "Invoke-RowActivate REFUSED Chrome BrowserWindow hwnd PID mismatch"
                        return
                    }
                }
            }
        }
        # v0.4.2: land the fronted window on a DIFFERENT monitor than the
        # sweeper itself when more than one exists - a fronted window
        # landing on the same screen as (behind/under) the sweeper was
        # Nick's complaint. Single monitor -> this block is inert.
        try {
            if ([SweeperWin32]::GetSystemMetrics([SweeperWin32]::SM_CMONITORS) -gt 1 -and $script:MainForm) {
                $sweeperMonitor = [SweeperWin32]::MonitorFromWindow($script:MainForm.Handle, [SweeperWin32]::MONITOR_DEFAULTTONEAREST)
                $targetMonitor = [SweeperWin32]::MonitorFromWindow($h, [SweeperWin32]::MONITOR_DEFAULTTONEAREST)
                if ($targetMonitor -eq $sweeperMonitor) {
                    $otherMonitor = [SweeperWin32]::GetAllMonitors() | Where-Object { $_ -ne $sweeperMonitor } | Select-Object -First 1
                    if ($otherMonitor) {
                        $mi = New-Object MONITORINFO
                        # Marshal.SizeOf([Type]) resolves ambiguously in PS
                        # 5.1 (tries to marshal the RuntimeType itself and
                        # throws) - pass the instance, not the type.
                        $mi.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($mi)
                        if ([SweeperWin32]::GetMonitorInfo($otherMonitor, [ref]$mi)) {
                            # Un-maximize first - SetWindowPos on a maximized
                            # window does not usefully relocate it.
                            if ([SweeperWin32]::IsZoomed($h)) { [SweeperWin32]::ShowWindow($h, [SweeperWin32]::SW_RESTORE) | Out-Null }
                            [SweeperWin32]::SetWindowPos($h, [IntPtr]::Zero, $mi.rcWork.Left, $mi.rcWork.Top, 0, 0, ([SweeperWin32]::SWP_NOSIZE -bor [SweeperWin32]::SWP_NOZORDER)) | Out-Null
                        }
                    }
                }
                # Already on a different monitor - leave position alone,
                # only the show-behavior below applies.
            }
        } catch { Write-DebugLog "Invoke-RowActivate MONITOR-MOVE EXCEPTION $($_.Exception.Message)" }

        # Nick hit a degenerate tiny-bar restore rect on Excel/Word (SDI
        # windows sometimes carry a junk restore rectangle) - unconditional
        # SW_RESTORE landed there. Excel/Word AND Outlook Inspector rows
        # (Nick wants Outlook to behave like Excel) now always come up
        # maximized; Chrome and Outlook Explorer (folder) rows are left at
        # whatever size they already are and only un-minimized if actually
        # minimized (never force-resized).
        $shouldMaximize = ($Row.App -eq 'Excel') -or ($Row.App -eq 'Word') -or ($Row.App -eq 'Outlook' -and $Row.RowType -eq 'Inspector')
        if ($shouldMaximize) {
            if ([SweeperWin32]::IsIconic($h)) { [SweeperWin32]::ShowWindow($h, [SweeperWin32]::SW_RESTORE) | Out-Null }
            [SweeperWin32]::ShowWindow($h, [SweeperWin32]::SW_SHOWMAXIMIZED) | Out-Null
        } elseif ([SweeperWin32]::IsIconic($h)) {
            [SweeperWin32]::ShowWindow($h, [SweeperWin32]::SW_RESTORE) | Out-Null
        }
        [SweeperWin32]::SetForegroundWindow($h) | Out-Null
        return
    }
    if ($route.ActivateStyle -eq 'CdpActivate') {
        # v0.4 Phase 3: CDP tabs have no HWND of their own - activation is
        # an HTTP call (/json/activate/<id>), routed through Invoke-ComTask
        # for the same uniformity as every other action even though it is
        # not COM and cannot hang the UI thread the way a COM call can (the
        # no-COM-on-click guarantee is preserved in spirit, not in letter).
        Invoke-ComTask -Command @{ Op = 'ChromeTabActivate'; Key = $Row.Key; TabTitle = $Row.Name } -TimeoutMs 5000 -OwnerForm $script:MainForm -StatusText "Activating $($Row.Name)..." | Out-Null
        return
    }
    Write-DebugLog "Invoke-RowActivate UNHANDLED-STYLE App=$($Row.App) RowType=$($Row.RowType) Style=$($route.ActivateStyle)"
}

# v0.4 Phase 0: op names now come from the routing table instead of an
# Excel-vs-everything-else hardcode. A row with no SaveCloseOp routed for its
# (App, RowType) no-ops with a logged warning rather than guessing - it
# should never actually happen, since CellContentClick only calls this for
# rows with CanSave=true, which only ever comes from RowTypes that have a
# SaveCloseOp seeded.
function Invoke-RowSaveClose {
    param($Row)
    $route = Get-RoutingEntry -App $Row.App -RowType $Row.RowType
    if (-not $route -or -not $route.SaveCloseOp) {
        Write-DebugLog "Invoke-RowSaveClose NO-ROUTE App=$($Row.App) RowType=$($Row.RowType)"
        return
    }
    $result = Invoke-ComTask -Command @{ Op = $route.SaveCloseOp; Key = $Row.Key } -TimeoutMs 10000 -OwnerForm $script:MainForm -StatusText "Saving $($Row.Name)..."
    if ($result.TimedOut) { Request-DeferredRefresh -App $Row.App; return }
    if ($result.NeedsSaveAs) {
        if (-not $route.SaveAsOp) {
            Write-DebugLog "Invoke-RowSaveClose NEEDS-SAVEAS-NO-ROUTE App=$($Row.App) RowType=$($Row.RowType)"
        } else {
            # Excel's native Save As dialog opens behind a TopMost sweeper -
            # observed live (the dialog existed, just invisible, foreground
            # stuck on the sweeper). Drop TopMost only for this one blocking
            # call so the dialog can actually come to the front, then restore
            # whatever the user had it set to.
            $wasTopMost = $false
            try { $wasTopMost = $script:MainForm.TopMost } catch {}
            try {
                if ($wasTopMost) { try { $script:MainForm.TopMost = $false } catch {} }
                $result = Invoke-ComTask -Command @{ Op = $route.SaveAsOp; Key = $Row.Key } -TimeoutMs 180000 -OwnerForm $script:MainForm -StatusText 'Waiting for you to finish Save As...'
            } finally {
                if ($wasTopMost) { try { $script:MainForm.TopMost = $true } catch {} }
            }
        }
    }
    Write-ActionLog -App $Row.App -Name $Row.Name -Action 'SaveClose' -Result $result -Path $Row.Detail
    Request-DeferredRefresh -App $Row.App
}

# Discard-confirm gate re-checks LIVE save-state at the moment of the click,
# not the row's cached Status from the last Refresh (which can go stale if
# Nick edits the workbook/email after refreshing without clicking Refresh
# again). A Saved-looking row is sent with ConfirmedDiscard=$false; only if
# the worker's live re-check (ExcelClose/OutlookClose) finds it is actually
# unsaved does it come back with NeedsConfirm, at which point the confirm
# dialog is shown and the op is re-issued with ConfirmedDiscard=$true. Rows
# already showing UNSAVED/Never saved confirm up front as before, then send
# ConfirmedDiscard=$true directly (no need for the round trip).
# v0.4 Phase 0: the RowType 'Explorer' special case folds into the generic
# routing lookup below (Outlook|Explorer -> CloseOp='ExplorerClose') - no
# functional change, since Explorer rows' Status is always 'Window' (never
# UNSAVED/Never saved) so the discard-confirm gate below was always a no-op
# for them anyway, and the worker's ExplorerClose case never reads
# cmd.ConfirmedDiscard.
function Invoke-RowClose {
    param($Row)
    $route = Get-RoutingEntry -App $Row.App -RowType $Row.RowType
    if (-not $route -or -not $route.CloseOp) {
        Write-DebugLog "Invoke-RowClose NO-ROUTE App=$($Row.App) RowType=$($Row.RowType)"
        return
    }
    $op = $route.CloseOp

    $confirmedDiscard = $false
    if ($Row.Status -eq 'UNSAVED' -or $Row.Status -eq 'Never saved') {
        if (-not (Show-DiscardConfirm -Rows @($Row))) { return }
        $confirmedDiscard = $true
    }

    $result = Invoke-ComTask -Command @{ Op = $op; Key = $Row.Key; ConfirmedDiscard = $confirmedDiscard } -TimeoutMs 10000 -OwnerForm $script:MainForm -StatusText "Closing $($Row.Name)..."
    if ($result.NeedsConfirm) {
        if (-not (Show-DiscardConfirm -Rows @($Row))) { return }
        $result = Invoke-ComTask -Command @{ Op = $op; Key = $Row.Key; ConfirmedDiscard = $true } -TimeoutMs 10000 -OwnerForm $script:MainForm -StatusText "Closing $($Row.Name)..."
    }
    Write-ActionLog -App $Row.App -Name $Row.Name -Action 'Close' -Result $result -Path $Row.Detail
    Request-DeferredRefresh -App $Row.App
}

# Shared driver for the three batch buttons. Tracks a timed-out app and skips
# its remaining rows for the rest of THIS batch, while the other app's rows
# keep going. ConfirmedDiscard is forwarded to Close ops unchanged for every
# row in the batch (Save ops ignore it); a row whose live state turns out to
# have changed since the last Refresh comes back as NeedsConfirm and is
# tallied as Skipped rather than silently forced through.
# v0.4 Phase 0: OpKind ('SaveCloseOp' or 'CloseOp') selects which field of
# the row's routing-table entry to use, replacing the old App-only
# OpNameForApp scriptblock - resolution is now per (App, RowType), so a
# batch spanning multiple RowTypes (e.g. Outlook Inspectors + Explorers)
# routes each row correctly. A row with no route (should not happen given
# the CanSave/CanClose gating upstream, but never guessed at) is tallied
# Failed with a logged warning.
function Invoke-BatchOp {
    param([array]$Rows, [string]$ActionLabel, [string]$OpKind, [bool]$ConfirmedDiscard = $false)
    $timedOutApps = @{}
    $counts = @{ Ok = 0; Failed = 0; Skipped = 0 }
    foreach ($row in $Rows) {
        if ($timedOutApps.ContainsKey($row.App)) { continue }
        $route = Get-RoutingEntry -App $row.App -RowType $row.RowType
        $op = $null
        if ($route) { $op = $route[$OpKind] }
        if (-not $op) {
            Write-DebugLog "Invoke-BatchOp NO-ROUTE App=$($row.App) RowType=$($row.RowType) OpKind=$OpKind"
            $counts.Failed++
            Write-ActionLog -App $row.App -Name $row.Name -Action $ActionLabel -Result @{ Ok = $false; Message = 'No routing for this row type' } -Path $row.Detail
            continue
        }
        $verb = 'Processing'
        if ($ActionLabel -eq 'SaveClose') { $verb = 'Saving' } elseif ($ActionLabel -eq 'Close') { $verb = 'Closing' }
        $result = Invoke-ComTask -Command @{ Op = $op; Key = $row.Key; ConfirmedDiscard = $ConfirmedDiscard } -TimeoutMs 10000 -OwnerForm $script:MainForm -StatusText "$verb $($row.Name)..."
        if ($result.TimedOut) {
            $timedOutApps[$row.App] = $true
            $counts.Failed++
            Write-ActionLog -App $row.App -Name $row.Name -Action $ActionLabel -Result $result -Path $row.Detail
            continue
        }
        if ($result.NeedsConfirm) {
            $counts.Skipped++
            Write-ActionLog -App $row.App -Name $row.Name -Action $ActionLabel -Result @{ Ok = $false; Message = 'changed since refresh - skipped' } -Path $row.Detail
            continue
        }
        Write-ActionLog -App $row.App -Name $row.Name -Action $ActionLabel -Result $result -Path $row.Detail
        if ($result.Ok) { $counts.Ok++ } else { $counts.Failed++ }
    }
    return $counts
}

# v0.4 Phase 0: Save & Close All scoping is unchanged in effect - CanSave
# rows, excluding pathless never-saved ones, which are tallied separately
# rather than silently attempted (a never-saved row's own Save&Close is what
# drives the native Save As dialog; the batch path has nowhere to show it).
# The RowType -eq 'Workbook'/'Inspector' restriction is dropped since it was
# always redundant with CanSave (only those two RowTypes ever set it true)
# and this generalizes cleanly to Word's Document rows in Phase 2.
function Invoke-SaveAll {
    $all = @($script:Grid.Tag)
    $rows = @($all | Where-Object { $_.CanSave -and $_.Status -ne 'Never saved' })
    $skipped = @($all | Where-Object { $_.CanSave -and $_.Status -eq 'Never saved' }).Count
    $counts = Invoke-BatchOp -Rows $rows -ActionLabel 'SaveClose' -OpKind 'SaveCloseOp'
    Do-RefreshForApps -Apps @($rows | ForEach-Object { $_.App })
    [System.Windows.Forms.MessageBox]::Show($script:MainForm, "Saved and closed: $($counts.Ok)`r`nNever-saved - skipped (use the row's own Save & Close): $skipped`r`nFailed: $($counts.Failed)", 'Save & Close All', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

# "Zero-risk, no confirm" per design - but that premise only holds if the
# cached Status is still accurate at click time, which nothing guarantees
# (Nick could have edited a row after the last Refresh). Every row is sent
# with ConfirmedDiscard=$false; the worker's live re-check catches any row
# that has actually gone unsaved since, and it is skipped rather than closed.
# Per plan: Status 'Saved' rows only - Chrome tabs (Phase 3) will never carry
# that status, so they never appear in this batch.
function Invoke-CloseAllSaved {
    $rows = @($script:Grid.Tag | Where-Object { $_.CanClose -and $_.Status -eq 'Saved' })
    $counts = Invoke-BatchOp -Rows $rows -ActionLabel 'Close' -OpKind 'CloseOp' -ConfirmedDiscard $false
    Do-RefreshForApps -Apps @($rows | ForEach-Object { $_.App })
    [System.Windows.Forms.MessageBox]::Show($script:MainForm, "Closed: $($counts.Ok)`r`nChanged since refresh - skipped: $($counts.Skipped)`r`nFailed: $($counts.Failed)", 'Close All Saved', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

# v0.4 Phase 0: broadened from "unsaved/never-saved rows only" to ALL
# closable rows, per plan - Chrome tabs (Phase 3) have no save state at all,
# so this batch is the only way they ever get swept in bulk. For today's
# Excel/Outlook-only world this also means an already-Saved row (or an
# Outlook folder window - Explorer rows are CanClose too) can be included in
# the batch alongside truly-unsaved ones; the discard confirm still lists
# every row by name before anything closes.
function Invoke-CloseAllDiscard {
    $rows = @($script:Grid.Tag | Where-Object { $_.CanClose })
    if ($rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($script:MainForm, 'Nothing to close.', 'Close All and Discard', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    if (-not (Show-DiscardConfirm -Rows $rows)) { return }
    # Names were already listed and confirmed above, so every row in this
    # batch is sent pre-confirmed.
    $counts = Invoke-BatchOp -Rows $rows -ActionLabel 'Close' -OpKind 'CloseOp' -ConfirmedDiscard $true
    Do-RefreshForApps -Apps @($rows | ForEach-Object { $_.App })
    [System.Windows.Forms.MessageBox]::Show($script:MainForm, "Closed: $($counts.Ok)`r`nFailed: $($counts.Failed)", 'Close All and Discard', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

function New-MainForm {
    $settings = Get-Settings
    $script:SelectedApps = @($settings.selectedApps)

    # v0.3 test hooks - zero effect unless one of the env vars is set. Read
    # once here for the banner/title; the actual row-dropping happens inside
    # the worker's EnumExcel/EnumOutlook (same env vars, inherited by the
    # runspace since environment variables are process-wide).
    $testFilterPattern = $env:WINDOWSWEEPER_TEST_FILTER
    $testPidEnv = $env:WINDOWSWEEPER_TEST_PID
    # v0.4 Phase 3: the Chrome fixture env vars are their own test-isolation
    # signal (see EnumChrome) - either one active also lights up the
    # existing TEST banner, same as the Excel/Outlook test vars.
    $testChromePortEnv = $env:WINDOWSWEEPER_CHROME_DEBUG_PORT
    $testChromePidEnv = $env:WINDOWSWEEPER_CHROME_TEST_PID
    $testActive = (-not [string]::IsNullOrEmpty($testFilterPattern)) -or (-not [string]::IsNullOrEmpty($testPidEnv)) -or (-not [string]::IsNullOrEmpty($testChromePortEnv)) -or (-not [string]::IsNullOrEmpty($testChromePidEnv))

    $form = New-Object System.Windows.Forms.Form
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9.75)
    $form.Text = 'Window Sweeper v0.4.1'
    if ($testActive) { $form.Text = $form.Text + ' [TEST]' }
    $form.Width = 920
    $form.Height = 600
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(720, 420)
    $form.TopMost = $settings.keepOnTop

    # v0.4 Phase 1: generated once beside the ps1 if missing; try/catch means
    # a failure anywhere in generation/load just leaves the form without an
    # icon rather than breaking startup.
    try {
        $icoPath = Join-Path $PSScriptRoot 'sweeper.ico'
        New-SweeperIcoFile -Path $icoPath
        if (Test-Path $icoPath) { $form.Icon = New-Object System.Drawing.Icon $icoPath }
    } catch { Write-DebugLog "Window icon EXCEPTION $($_.Exception.Message)" }

    $script:TopPanel = New-Object System.Windows.Forms.Panel
    $script:TopPanel.Dock = 'Top'
    $script:TopPanel.Height = 40

    # v0.4 Phase 0: per-app checkboxes replace the old Both/Excel/Outlook
    # radios - Word and Chrome appear now (all four checked by default) even
    # though their adapters don't land until Phase 2/3; toggling them today
    # just yields zero rows for that app. Laid out left-to-right using each
    # checkbox's own preferred width so "Outlook"/"Chrome" never overlap the
    # next control, rather than guessing fixed pixel gaps.
    $script:ChkApps = @{}
    $xPos = 10
    foreach ($appName in $script:KnownApps) {
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text = $appName
        $chk.AutoSize = $true
        $chk.Checked = ($script:SelectedApps -contains $appName)
        $chk.Location = New-Object System.Drawing.Point($xPos, 11)
        $script:ChkApps[$appName] = $chk
        $script:TopPanel.Controls.Add($chk)
        $pref = $chk.GetPreferredSize([System.Drawing.Size]::Empty)
        $xPos += $pref.Width + 14
    }

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = 'Refresh'; $btnRefresh.Location = New-Object System.Drawing.Point($xPos, 7); $btnRefresh.Width = 80
    $xPos += $btnRefresh.Width + 20

    # Default checked - this is what makes row-click peeking actually usable:
    # the target workbook/email fronts UNDER the sweeper instead of hiding it.
    $chkKeepOnTop = New-Object System.Windows.Forms.CheckBox
    $chkKeepOnTop.Text = 'Keep on top'; $chkKeepOnTop.Location = New-Object System.Drawing.Point($xPos, 11); $chkKeepOnTop.AutoSize = $true
    $chkKeepOnTop.Checked = $settings.keepOnTop
    $prefKot = $chkKeepOnTop.GetPreferredSize([System.Drawing.Size]::Empty)
    $xPos += $prefKot.Width + 20

    $script:LblTimestamp = New-Object System.Windows.Forms.Label
    $script:LblTimestamp.Location = New-Object System.Drawing.Point($xPos, 12); $script:LblTimestamp.AutoSize = $true
    $script:LblTimestamp.ForeColor = [System.Drawing.Color]::Gray

    $script:TopPanel.Controls.AddRange([System.Windows.Forms.Control[]]@($btnRefresh, $chkKeepOnTop, $script:LblTimestamp))
    Set-FlatButtonStyle -Button $btnRefresh

    $script:Banner = New-Object System.Windows.Forms.Panel
    $script:Banner.Dock = 'Top'; $script:Banner.Height = 32
    $script:Banner.BackColor = [System.Drawing.Color]::FromArgb(255, 229, 180)
    $script:Banner.Visible = $false

    $script:LblBanner = New-Object System.Windows.Forms.Label
    $script:LblBanner.AutoSize = $true; $script:LblBanner.Location = New-Object System.Drawing.Point(10, 8)
    $script:LblBanner.ForeColor = [System.Drawing.Color]::FromArgb(153, 76, 0)

    $script:BtnShowDialog = New-Object System.Windows.Forms.Button
    $script:BtnShowDialog.Text = 'Show the dialog'
    $script:BtnShowDialog.Location = New-Object System.Drawing.Point(700, 3); $script:BtnShowDialog.Width = 150

    $script:Banner.Controls.AddRange([System.Windows.Forms.Control[]]@($script:LblBanner, $script:BtnShowDialog))
    Set-FlatButtonStyle -Button $script:BtnShowDialog

    # v0.3 test filter banner - only built when a test env var is set.
    $testBanner = $null
    if ($testActive) {
        $testParts = @()
        if (-not [string]::IsNullOrEmpty($testFilterPattern)) { $testParts += "pattern='$testFilterPattern'" }
        if (-not [string]::IsNullOrEmpty($testPidEnv)) { $testParts += "pid=$testPidEnv" }
        if (-not [string]::IsNullOrEmpty($testChromePortEnv)) { $testParts += "chrome port=$testChromePortEnv" }
        if (-not [string]::IsNullOrEmpty($testChromePidEnv)) { $testParts += "chrome pid=$testChromePidEnv" }
        $testBanner = New-Object System.Windows.Forms.Panel
        $testBanner.Dock = 'Top'; $testBanner.Height = 24
        $testBanner.BackColor = [System.Drawing.Color]::FromArgb(255, 210, 210)
        $lblTest = New-Object System.Windows.Forms.Label
        $lblTest.Text = 'TEST FILTER ACTIVE: ' + ($testParts -join ', ')
        $lblTest.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
        $lblTest.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
        $lblTest.AutoSize = $true
        $lblTest.Location = New-Object System.Drawing.Point(10, 3)
        $testBanner.Controls.Add($lblTest)
    }

    $script:BottomPanel = New-Object System.Windows.Forms.Panel
    $script:BottomPanel.Dock = 'Bottom'; $script:BottomPanel.Height = 64

    $script:LblCounts = New-Object System.Windows.Forms.Label
    $script:LblCounts.Location = New-Object System.Drawing.Point(10, 12); $script:LblCounts.AutoSize = $true
    # Sol P2 (colored text, per plan): a single neutral dark accent rather
    # than per-app color spans - a plain Label can't paint multi-color runs
    # without becoming an owner-drawn control, and the per-app breakdown is
    # already conveyed in the text itself ("12 Excel | 3 Word | ...").
    $script:LblCounts.ForeColor = [System.Drawing.Color]::FromArgb(51, 51, 51)

    # "&&" so the literal ampersand displays (a plain Button, unlike a grid
    # cell, treats a single "&" as a mnemonic-underline prefix and hides it).
    $btnSaveAll = New-Object System.Windows.Forms.Button
    $btnSaveAll.Text = 'Save && Close All'; $btnSaveAll.Location = New-Object System.Drawing.Point(420, 8); $btnSaveAll.Width = 120

    $btnCloseSaved = New-Object System.Windows.Forms.Button
    $btnCloseSaved.Text = 'Close All Saved'; $btnCloseSaved.Location = New-Object System.Drawing.Point(550, 8); $btnCloseSaved.Width = 120

    $btnCloseDiscard = New-Object System.Windows.Forms.Button
    $btnCloseDiscard.Text = 'Close All and Discard...'; $btnCloseDiscard.Location = New-Object System.Drawing.Point(680, 8); $btnCloseDiscard.Width = 170

    # v0.3 progress feedback - marquee + status text, visible only while a
    # worker call is in flight (toggled by Set-BusyUi inside Invoke-ComTask).
    $script:LblStatus = New-Object System.Windows.Forms.Label
    $script:LblStatus.Location = New-Object System.Drawing.Point(10, 42); $script:LblStatus.AutoSize = $true
    $script:LblStatus.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 160)

    $script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $script:ProgressBar.Style = 'Marquee'
    $script:ProgressBar.MarqueeAnimationSpeed = 30
    $script:ProgressBar.Location = New-Object System.Drawing.Point(200, 40)
    $script:ProgressBar.Size = New-Object System.Drawing.Size(160, 16)
    $script:ProgressBar.Visible = $false

    $script:BottomPanel.Controls.AddRange([System.Windows.Forms.Control[]]@($script:LblCounts, $btnSaveAll, $btnCloseSaved, $btnCloseDiscard, $script:LblStatus, $script:ProgressBar))
    Set-FlatButtonStyle -Button $btnSaveAll
    Set-FlatButtonStyle -Button $btnCloseSaved
    Set-FlatButtonStyle -Button $btnCloseDiscard

    # v0.4 Phase 1: precreated GDI+ resources for the grid's CellPainting
    # handler - built once here, per form lifetime, never per-paint. Shared
    # by every row/cell repaint rather than allocated inside the handler.
    $script:AccentBrushes = @{
        'Excel'   = New-HtmlBrush '#217346'
        'Word'    = New-HtmlBrush '#2B579A'
        'Outlook' = New-HtmlBrush '#0F6CBD'
        'Chrome'  = New-HtmlBrush '#F4B400'
        'Busy'    = New-HtmlBrush '#9E9E9E'
        'Default' = New-HtmlBrush '#9E9E9E'
    }
    $pillFontRegular = New-Object System.Drawing.Font('Segoe UI', 8.25, [System.Drawing.FontStyle]::Regular)
    $pillFontBold = New-Object System.Drawing.Font('Segoe UI', 8.25, [System.Drawing.FontStyle]::Bold)
    $script:StatusStyles = @{
        'UNSAVED'     = @{ Fill = (New-HtmlBrush '#C62828'); Text = [System.Drawing.Brushes]::White; Font = $pillFontBold }
        'Never saved' = @{ Fill = (New-HtmlBrush '#EF6C00'); Text = [System.Drawing.Brushes]::White; Font = $pillFontRegular }
        'Saved'       = @{ Fill = (New-HtmlBrush '#E8F5E9'); Text = (New-HtmlBrush '#1B5E20'); Font = $pillFontRegular }
    }
    $script:StatusDefaultStyle = @{ Fill = (New-HtmlBrush '#EEEEEE'); Text = (New-HtmlBrush '#555555'); Font = $pillFontRegular }
    $script:CenterFormat = New-Object System.Drawing.StringFormat
    $script:CenterFormat.Alignment = [System.Drawing.StringAlignment]::Center
    $script:CenterFormat.LineAlignment = [System.Drawing.StringAlignment]::Center

    Initialize-AppIcons

    $script:Grid = New-Object System.Windows.Forms.DataGridView
    $script:Grid.Dock = 'Fill'
    $script:Grid.Font = New-Object System.Drawing.Font('Segoe UI', 9.75)
    $script:Grid.AllowUserToAddRows = $false
    $script:Grid.AllowUserToDeleteRows = $false
    $script:Grid.ReadOnly = $true
    $script:Grid.RowHeadersVisible = $false
    $script:Grid.MultiSelect = $false
    $script:Grid.SelectionMode = 'FullRowSelect'
    $script:Grid.AutoSizeColumnsMode = 'Fill'
    $script:Grid.RowTemplate.Height = 26
    $script:Grid.ColumnHeadersHeight = 30
    $script:Grid.EnableHeadersVisualStyles = $false
    $script:Grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9.75, [System.Drawing.FontStyle]::Bold)
    $script:Grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $script:Grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $script:Grid.ColumnHeadersBorderStyle = 'Single'
    $script:Grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $script:Grid.CellBorderStyle = 'SingleHorizontal'
    $script:Grid.AllowUserToResizeRows = $false

    # v0.4 Phase 1: fixed-width icon column, excluded from Fill so it never
    # competes with the text columns for space. NullValue=$null keeps a row
    # with no resolved icon (Initialize-AppIcons failure) genuinely blank
    # instead of WinForms' default broken-image glyph.
    $colIcon = New-Object System.Windows.Forms.DataGridViewImageColumn
    $colIcon.Name = 'Icon'; $colIcon.HeaderText = ''
    $colIcon.Width = 34
    $colIcon.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    $colIcon.Resizable = [System.Windows.Forms.DataGridViewTriState]::False
    $colIcon.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    $colIcon.ImageLayout = [System.Windows.Forms.DataGridViewImageCellLayout]::Zoom
    $colIcon.DefaultCellStyle.NullValue = $null
    $colIcon.ReadOnly = $true

    $colApp = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colApp.Name = 'App'; $colApp.HeaderText = 'App'; $colApp.FillWeight = 55; $colApp.ReadOnly = $true

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.Name = 'Name'; $colName.HeaderText = 'Name'; $colName.FillWeight = 150; $colName.ReadOnly = $true

    $colDetail = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDetail.Name = 'Detail'; $colDetail.HeaderText = 'Detail'; $colDetail.FillWeight = 230; $colDetail.ReadOnly = $true

    $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStatus.Name = 'Status'; $colStatus.HeaderText = 'Status'; $colStatus.FillWeight = 90; $colStatus.ReadOnly = $true

    $colSave = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $colSave.Name = 'SaveClose'; $colSave.HeaderText = 'Save'; $colSave.FillWeight = 95
    $colSave.UseColumnTextForButtonValue = $false

    $colClose = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $colClose.Name = 'Close'; $colClose.HeaderText = 'Close'; $colClose.FillWeight = 45
    $colClose.UseColumnTextForButtonValue = $false

    $script:Grid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colIcon, $colApp, $colName, $colDetail, $colStatus, $colSave, $colClose))

    # v0.4 Phase 1: (a) the 4px per-app accent stripe, painted into the icon
    # column alongside the icon itself; (b) the Status column rendered as a
    # rounded pill instead of plain text. Every branch is wrapped in the
    # single outer try/catch - on any exception e.Handled is never set, so
    # DataGridView falls back to its own default cell painting instead of
    # leaving a half-drawn cell or crashing the UI thread.
    # $script:-scoped, not function-local: a plain scriptblock handed to
    # Add_CellPainting does not reliably close over New-MainForm's own local
    # variables (confirmed - they read back $null once the event actually
    # fires, well after New-MainForm has returned), only over $script:/
    # $global: scope. Every other per-form CellPainting resource above
    # (AccentBrushes, StatusStyles, CenterFormat, AppIcons) was already
    # $script:-scoped for the same reason; these two just close the gap.
    $script:IconColIndex = $script:Grid.Columns['Icon'].Index
    $script:StatusColIndex = $script:Grid.Columns['Status'].Index
    $script:Grid.Add_CellPainting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
            if ($e.ColumnIndex -eq $script:IconColIndex) {
                $gridRow = $script:Grid.Rows[$e.RowIndex]
                $row = $gridRow.Tag
                $appName = $null; $rowType = $null
                if ($row) { $appName = $row.App; $rowType = $row.RowType }
                $accent = $script:AccentBrushes['Default']
                if ($rowType -eq 'Busy' -or $rowType -eq 'Info') {
                    $accent = $script:AccentBrushes['Busy']
                } elseif ($appName -and $script:AccentBrushes.ContainsKey($appName)) {
                    $accent = $script:AccentBrushes[$appName]
                }
                $e.PaintBackground($e.CellBounds, $true)
                $stripeRect = New-Object System.Drawing.Rectangle($e.CellBounds.X, $e.CellBounds.Y, 4, $e.CellBounds.Height)
                $e.Graphics.FillRectangle($accent, $stripeRect)
                $img = $e.Value
                if ($img) {
                    $imgSize = 20
                    $availX = $e.CellBounds.X + 4
                    $availW = $e.CellBounds.Width - 4
                    $drawX = $availX + [int](($availW - $imgSize) / 2)
                    $drawY = $e.CellBounds.Y + [int](($e.CellBounds.Height - $imgSize) / 2)
                    $e.Graphics.DrawImage($img, $drawX, $drawY, $imgSize, $imgSize)
                }
                $e.Handled = $true
                return
            }
            if ($e.ColumnIndex -eq $script:StatusColIndex) {
                $statusText = ''
                if ($e.Value) { $statusText = "$($e.Value)" }
                $style = $script:StatusDefaultStyle
                if ($script:StatusStyles.ContainsKey($statusText)) { $style = $script:StatusStyles[$statusText] }
                $e.PaintBackground($e.CellBounds, $true)
                $textSize = $e.Graphics.MeasureString($statusText, $style.Font)
                $padX = 10; $padY = 3
                $pillW = [Math]::Min($e.CellBounds.Width - 6, [int]$textSize.Width + $padX * 2)
                $pillH = [Math]::Min($e.CellBounds.Height - 6, [int]$textSize.Height + $padY * 2)
                if ($pillW -lt 1) { $pillW = 1 }
                if ($pillH -lt 1) { $pillH = 1 }
                $pillX = $e.CellBounds.X + [int](($e.CellBounds.Width - $pillW) / 2)
                $pillY = $e.CellBounds.Y + [int](($e.CellBounds.Height - $pillH) / 2)
                $pillRect = New-Object System.Drawing.Rectangle($pillX, $pillY, $pillW, $pillH)
                $radius = [Math]::Max(2, [Math]::Min(9, [int]($pillH / 2)))
                $path = New-RoundedRectPath -Rect $pillRect -Radius $radius
                $oldMode = $e.Graphics.SmoothingMode
                # try/finally so a throw from FillPath/DrawString (caught by
                # the outer catch, which falls back to default painting)
                # still disposes $path and restores SmoothingMode instead of
                # leaking the GDI+ path object and leaving the Graphics in
                # AntiAlias mode for whatever paints next.
                try {
                    $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                    $e.Graphics.FillPath($style.Fill, $path)
                    # PowerShell's .NET overload binding does not implicitly
                    # widen a Rectangle to RectangleF the way C# does -
                    # without this it silently resolves to the
                    # (string,Font,Brush,PointF,StringFormat) overload
                    # instead and throws on the argument type mismatch.
                    $pillRectF = New-Object System.Drawing.RectangleF($pillX, $pillY, $pillW, $pillH)
                    $e.Graphics.DrawString($statusText, $style.Font, $style.Text, $pillRectF, $script:CenterFormat)
                } finally {
                    $path.Dispose()
                    $e.Graphics.SmoothingMode = $oldMode
                }
                $e.Handled = $true
                return
            }
        } catch {
            Write-DebugLog "CellPainting EXCEPTION col=$($e.ColumnIndex) row=$($e.RowIndex) Message=$($_.Exception.Message) Stack=$($_.ScriptStackTrace)"
            # e.Handled stays false - DataGridView falls back to default painting.
        }
    })

    $script:Grid.Add_CellContentClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        # Panels being disabled while busy should already block mouse input,
        # but a keyboard Enter/Space on a focused button cell can still raise
        # this event - refuse explicitly rather than stacking a second
        # in-flight action (and a second, competing grid rebuild) on top of
        # whatever Invoke-ComTask call is already running.
        if ($script:Busy) { Write-DebugLog "CellContentClick SKIP (busy) row=$($e.RowIndex) col=$($e.ColumnIndex)"; return }
        try {
            $gridRow = $script:Grid.Rows[$e.RowIndex]
            $row = $gridRow.Tag
            $colName2 = $script:Grid.Columns[$e.ColumnIndex].Name
            Write-DebugLog "CellContentClick ENTER row=$($e.RowIndex) col=$colName2 name=$($row.Name)"
            if ($colName2 -eq 'SaveClose') {
                if (-not $row.CanSave) { return }
                Invoke-RowSaveClose -Row $row
            } elseif ($colName2 -eq 'Close') {
                if (-not $row.CanClose) { return }
                Invoke-RowClose -Row $row
            }
        } catch {
            Write-DebugLog "CellContentClick EXCEPTION $($_.Exception.Message) Stack=$($_.ScriptStackTrace)"
        }
    })

    $script:Grid.Add_CellClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $colName2 = $script:Grid.Columns[$e.ColumnIndex].Name
        if ($colName2 -ne 'Name' -and $colName2 -ne 'Detail') { return }
        try { Invoke-RowActivate -Row $script:Grid.Rows[$e.RowIndex].Tag }
        catch { Write-DebugLog "CellClick EXCEPTION $($_.Exception.Message) Stack=$($_.ScriptStackTrace)" }
    })

    if ($testBanner) { $form.Controls.Add($testBanner) }
    $form.Controls.Add($script:TopPanel)
    $form.Controls.Add($script:Banner)
    $form.Controls.Add($script:BottomPanel)
    $form.Controls.Add($script:Grid)
    # Dock layout processes controls back-to-front, so the Fill-docked grid must
    # be FRONTMOST or it lays out first and fills the whole client area with the
    # top-docked banners painted over its header and first rows (the "grid shows
    # only the last row" bug - rows 0..n-1 were hidden UNDER the panels).
    $script:Grid.BringToFront()

    # v0.4 Phase 0: identical handler on all four checkboxes - it just
    # recomputes $script:SelectedApps from whichever boxes are checked right
    # now, so it does not matter which single checkbox actually fired it.
    foreach ($appName in $script:KnownApps) {
        $script:ChkApps[$appName].Add_CheckedChanged({
            $script:SelectedApps = @($script:KnownApps | Where-Object { $script:ChkApps[$_].Checked })
            Do-FullRefresh
        })
    }
    # Refresh button is the one and only click that always retries Outlook,
    # bypassing an active backoff - every other refresh path respects it.
    $btnRefresh.Add_Click({ Do-FullRefresh -ForceOutlookRetry $true })
    $btnSaveAll.Add_Click({ Invoke-SaveAll })
    $btnCloseSaved.Add_Click({ Invoke-CloseAllSaved })
    $btnCloseDiscard.Add_Click({ Invoke-CloseAllDiscard })
    # $this = the sender checkbox; a function-local $chkKeepOnTop is NOT visible
    # inside the handler scriptblock (PS closure gap), so the old form no-op'd.
    $chkKeepOnTop.Add_CheckedChanged({ $script:MainForm.TopMost = $this.Checked })
    $script:BtnShowDialog.Add_Click({
        $app = $script:BtnShowDialog.Tag
        $procName = 'OUTLOOK'
        if ($app -eq 'Excel') { $procName = 'EXCEL' }
        elseif ($app -eq 'Word') { $procName = 'WINWORD' }
        Show-BlockingDialog -ProcessName $procName | Out-Null
    })

    # v0.3 staleness auto-refresh: coming back to the window (alt-tab, click
    # from another app) silently rescans Excel if the last snapshot is more
    # than 30s old, so a grid left open for a while does not show stale "no
    # longer open" errors on the first click. Deliberately Excel-only - never
    # Outlook, which would reintroduce the exact freeze this rework exists to
    # remove - and throttled so re-activating repeatedly cannot refire it more
    # than once per 10s.
    $form.Add_Activated({
        Write-DebugLog "Activated FIRE Busy=$script:Busy"
        if ($script:Busy) { return }
        $now = Get-Date
        if (($now - $script:LastActivatedRefreshAt).TotalSeconds -lt 10) { return }
        if (($now - $script:LastSnapshotTime).TotalSeconds -lt 30) { return }
        $script:LastActivatedRefreshAt = $now
        try { Do-ScopedRefresh -App 'Excel' }
        catch { Write-DebugLog "Activated EXCEPTION $($_.Exception.Message) Stack=$($_.ScriptStackTrace)" }
    })

    # v0.3.1: fires the actual grid-mutating refresh a tick after a grid-row
    # action requests it (Request-DeferredRefresh), so Update-Grid's
    # Rows.Clear()/Add() never runs while the DataGridView is still
    # dispatching the very click that triggered it.
    $script:DeferredRefreshTimer = New-Object System.Windows.Forms.Timer
    $script:DeferredRefreshTimer.Interval = 30
    $script:DeferredRefreshTimer.Add_Tick({
        $script:DeferredRefreshTimer.Stop()
        $app = $script:PendingRefreshApp
        $script:PendingRefreshApp = $null
        Write-DebugLog "DeferredRefreshTimer TICK app=$app Busy=$script:Busy"
        if (-not $app) { return }
        if ($script:Busy) { Request-DeferredRefresh -App $app; return }
        try {
            if ($app -eq 'Both') { Do-FullRefresh } else { Do-ScopedRefresh -App $app }
        } catch {
            Write-DebugLog "DeferredRefreshTimer EXCEPTION $($_.Exception.Message) Stack=$($_.ScriptStackTrace)"
        }
    })

    return $form
}

# ===== Region 6: Entry =====

try {
    Initialize-StateDir
    $script:MainForm = New-MainForm

    $startTimer = New-Object System.Windows.Forms.Timer
    $startTimer.Interval = 200
    $startTimer.Add_Tick({
        $startTimer.Stop()
        Write-DebugLog "StartTimer TICK Busy=$script:Busy"
        if ($script:Busy) { return }
        try { Do-FullRefresh } catch { Write-DebugLog "StartTimer EXCEPTION $($_.Exception.Message) Stack=$($_.ScriptStackTrace)" }
    })
    $script:MainForm.Add_Shown({ $startTimer.Start() })

    $script:MainForm.Add_FormClosing({
        $swClose = [System.Diagnostics.Stopwatch]::StartNew()
        Write-DebugLog "FormClosing ENTER Busy=$script:Busy"
        $script:ClosingRequested = $true
        try {
            $s = Get-Settings
            $s.selectedApps = $script:SelectedApps
            try { $s.keepOnTop = $script:MainForm.TopMost } catch {}
            Save-Settings -Settings $s
        } catch {
            Write-DebugLog "FormClosing SETTINGS-EXCEPTION $($_.Exception.Message) Stack=$($_.ScriptStackTrace)"
        }
        Write-DebugLog "FormClosing PRE-CloseComRunspace ElapsedMs=$($swClose.ElapsedMilliseconds) Busy=$script:Busy"
        Close-ComRunspace
        Write-DebugLog "FormClosing EXIT ElapsedMs=$($swClose.ElapsedMilliseconds)"
    })

    [void][System.Windows.Forms.Application]::Run($script:MainForm)
    Write-DebugLog "Application.Run RETURNED"
} catch {
    Write-DebugLog "Region6 CATCH $($_.Exception.Message) Stack=$($_.ScriptStackTrace)"
    [System.Windows.Forms.MessageBox]::Show("Window Sweeper hit an error and needs to close:`r`n`r`n$($_.Exception.Message)", 'Window Sweeper', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
} finally {
    # Belt and suspenders: Close-ComRunspace above handles the normal case,
    # but if a runspace was ever abandoned as a zombie (Invoke-ComTask), its
    # background STA/COM thread would otherwise keep this process alive
    # forever after the window closes. Force-terminate rather than leave a
    # zombie powershell.exe behind.
    Write-DebugLog "Region6 FINALLY Busy=$script:Busy ClosingRequested=$script:ClosingRequested - calling Close-ComRunspace + Exit(0)"
    Close-ComRunspace
    [Environment]::Exit(0)
}
