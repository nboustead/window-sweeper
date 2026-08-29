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

    public const int SW_RESTORE = 9;
    public const uint WM_NULL = 0x0000;
    public const uint SMTO_ABORTIFHUNG = 0x0002;
    private const uint OBJID_NATIVEOM = 0xFFFFFFF0;

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
    return [PSCustomObject]@{ defaultFilter = 'Both'; confirmDiscard = $true; keepOnTop = $true }
}

function Get-Settings {
    if (-not (Test-Path $script:SettingsPath)) { return Get-DefaultSettings }
    try {
        $obj = Get-Content -Path $script:SettingsPath -Raw | ConvertFrom-Json
        $result = Get-DefaultSettings
        if ($obj.PSObject.Properties.Match('defaultFilter').Count -gt 0 -and @('Both', 'Excel', 'Outlook') -contains $obj.defaultFilter) {
            $result.defaultFilter = $obj.defaultFilter
        }
        if ($obj.PSObject.Properties.Match('confirmDiscard').Count -gt 0) {
            $result.confirmDiscard = [bool]$obj.confirmDiscard
        }
        if ($obj.PSObject.Properties.Match('keepOnTop').Count -gt 0) {
            $result.keepOnTop = [bool]$obj.keepOnTop
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
        $Settings | ConvertTo-Json | Set-Content -Path $script:SettingsPath -Encoding UTF8
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
                            $saved = $true
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
                        $subj = ''
                        try { $subj = $item.Subject } catch {}
                        if (-not $subj) { $subj = '(no subject)' }
                        $mc = ''
                        try { $mc = $item.MessageClass } catch {}
                        $isReport = $mc -like 'REPORT.*'
                        $saved = $true
                        try { if ($item.PSObject.Properties.Match('Saved').Count -gt 0) { $saved = [bool]$item.Saved } } catch {}
                        $status = 'Saved'
                        if (-not $saved) { $status = 'UNSAVED' }
                        $hwndVal = 0
                        $caption = ''
                        try { $caption = $inspectors[$i].Caption } catch {}
                        if ($caption) {
                            $capMatches = @($inspWins | Where-Object { $_.Title -eq $caption -or $_.Title.StartsWith($caption) })
                            if ($capMatches.Count -eq 1) { $hwndVal = $capMatches[0].Hwnd.ToInt64() }
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
                        $rows += [PSCustomObject]@{
                            App = 'Outlook'; RowType = 'Explorer'; Name = $caption; Detail = $detail; Status = 'Window'
                            Key = @{ Kind = 'Index'; Index = $i; VerifyCaption = $caption }; Hwnd = 0; CanSave = $false; CanClose = $true
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
                $saved = $true
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
                $saved = $true
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

            default { return @{ Ok = $false; Message = "Unknown op: $($cmd.Op)" } }
        }
    } catch {
        return @{ Ok = $false; Message = "Worker error: $($_.Exception.Message)" }
    }
}

# ===== Region 4: Snapshot orchestration =====

function Get-CountsSummary {
    param($rows)
    $wbRows = @($rows | Where-Object { $_.RowType -eq 'Workbook' })
    $mailRows = @($rows | Where-Object { $_.RowType -eq 'Inspector' })
    $unsaved = @($rows | Where-Object { ($_.RowType -eq 'Workbook' -or $_.RowType -eq 'Inspector') -and ($_.Status -eq 'UNSAVED' -or $_.Status -eq 'Never saved') })
    $wbWord = 'workbooks'
    if ($wbRows.Count -eq 1) { $wbWord = 'workbook' }
    $mailWord = 'emails'
    if ($mailRows.Count -eq 1) { $mailWord = 'email' }
    $text = "$($wbRows.Count) $wbWord, $($mailRows.Count) $mailWord"
    if ($unsaved.Count -gt 0) { $text += " - $($unsaved.Count) unsaved" }
    return $text
}

# Each app is enumerated through its own Invoke-ComTask call, so a busy
# Outlook never blocks Excel results (or vice versa).
function Get-Snapshot {
    param([string]$Filter, $OwnerForm, [bool]$ForceOutlookRetry = $false)

    $rows = @()
    $busyApps = @()
    $messages = @()

    if ($Filter -eq 'Both' -or $Filter -eq 'Excel') {
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
    # to start a fresh Outlook wait too - bail before ever calling
    # Invoke-ComTask again.
    if ($script:ClosingRequested) {
        Write-DebugLog "Get-Snapshot ABORT-before-outlook (closing)"
        return @{ Rows = $rows; BusyApps = @($busyApps | Select-Object -Unique); Messages = $messages }
    }

    if ($Filter -eq 'Both' -or $Filter -eq 'Outlook') {
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

    return @{ Rows = $rows; BusyApps = @($busyApps | Select-Object -Unique); Messages = $messages }
}

function Invoke-DiagMode {
    Initialize-StateDir
    $snap = Get-Snapshot -Filter 'Both' -OwnerForm $null
    Write-Host ''
    Write-Host 'Window Sweeper v0.3 - Diagnostic Enumeration' -ForegroundColor Cyan
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

# ===== Region 5: UI =====

$script:CurrentFilter = 'Both'
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
            if ($row.Status -eq 'UNSAVED') {
                $gridRow.Cells['Status'].Style.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
                $gridRow.Cells['Status'].Style.Font = New-Object System.Drawing.Font($script:Grid.Font, [System.Drawing.FontStyle]::Bold)
            } elseif ($row.Status -eq 'Never saved') {
                $gridRow.Cells['Status'].Style.ForeColor = [System.Drawing.Color]::FromArgb(180, 95, 0)
            } elseif ($row.Status -eq 'Saved') {
                $gridRow.Cells['Status'].Style.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 0)
            }
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
function Update-RefreshedUi {
    param($Rows, $BusyApps)
    try { Update-Grid -rows $Rows } catch { Write-DebugLog "Update-RefreshedUi Update-Grid EXCEPTION $($_.Exception.Message) Stack=$($_.ScriptStackTrace)" }
    try { $script:LblCounts.Text = Get-CountsSummary -rows $Rows } catch { Write-DebugLog "Update-RefreshedUi LblCounts EXCEPTION $($_.Exception.Message)" }
    try { $script:LblTimestamp.Text = 'Last refreshed: ' + (Get-Date -Format 'h:mm:ss tt') } catch {}
    try { Update-BusyBanner -BusyApps $BusyApps } catch { Write-DebugLog "Update-RefreshedUi Update-BusyBanner EXCEPTION $($_.Exception.Message)" }
}

function Do-FullRefresh {
    param([bool]$ForceOutlookRetry = $false)
    Write-DebugLog "Do-FullRefresh ENTER ForceOutlookRetry=$ForceOutlookRetry Filter=$script:CurrentFilter Closing=$script:ClosingRequested"
    if ($script:ClosingRequested) { Write-DebugLog "Do-FullRefresh ABORT (closing)"; return }
    $snap = Get-Snapshot -Filter $script:CurrentFilter -OwnerForm $script:MainForm -ForceOutlookRetry $ForceOutlookRetry
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
    if ($script:CurrentFilter -ne 'Both' -and $script:CurrentFilter -ne $App) { return }
    $snap = Get-Snapshot -Filter $App -OwnerForm $script:MainForm -ForceOutlookRetry $ForceOutlookRetry
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
    param([string[]]$Names)
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

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "This will discard unsaved changes in:`r`n`r`n" + (($Names | ForEach-Object { " - $_" }) -join "`r`n")
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
function Invoke-RowActivate {
    param($Row)
    if ($Row.RowType -ne 'Workbook' -and $Row.RowType -ne 'Inspector') { return }
    if (-not $Row.Hwnd -or $Row.Hwnd -eq 0) { return }  # no unique window captured - silent no-op
    $h = [IntPtr]$Row.Hwnd
    [SweeperWin32]::ShowWindow($h, [SweeperWin32]::SW_RESTORE) | Out-Null
    [SweeperWin32]::SetForegroundWindow($h) | Out-Null
}

function Invoke-RowSaveClose {
    param($Row)
    $op = 'OutlookSaveClose'
    if ($Row.App -eq 'Excel') { $op = 'ExcelSaveClose' }
    $result = Invoke-ComTask -Command @{ Op = $op; Key = $Row.Key } -TimeoutMs 10000 -OwnerForm $script:MainForm -StatusText "Saving $($Row.Name)..."
    if ($result.TimedOut) { Request-DeferredRefresh -App $Row.App; return }
    if ($result.NeedsSaveAs) {
        # Excel's native Save As dialog opens behind a TopMost sweeper -
        # observed live (the dialog existed, just invisible, foreground
        # stuck on the sweeper). Drop TopMost only for this one blocking
        # call so the dialog can actually come to the front, then restore
        # whatever the user had it set to.
        $wasTopMost = $false
        try { $wasTopMost = $script:MainForm.TopMost } catch {}
        try {
            if ($wasTopMost) { try { $script:MainForm.TopMost = $false } catch {} }
            $result = Invoke-ComTask -Command @{ Op = 'ExcelSaveAs'; Key = $Row.Key } -TimeoutMs 180000 -OwnerForm $script:MainForm -StatusText 'Waiting for you to finish Save As in Excel...'
        } finally {
            if ($wasTopMost) { try { $script:MainForm.TopMost = $true } catch {} }
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
function Invoke-RowClose {
    param($Row)
    if ($Row.RowType -eq 'Explorer') {
        $result = Invoke-ComTask -Command @{ Op = 'ExplorerClose'; Key = $Row.Key } -TimeoutMs 10000 -OwnerForm $script:MainForm -StatusText "Closing $($Row.Name)..."
        Write-ActionLog -App $Row.App -Name $Row.Name -Action 'Close' -Result $result -Path $Row.Detail
        Request-DeferredRefresh -App $Row.App
        return
    }
    $op = 'OutlookClose'
    if ($Row.App -eq 'Excel') { $op = 'ExcelClose' }

    $confirmedDiscard = $false
    if ($Row.Status -eq 'UNSAVED' -or $Row.Status -eq 'Never saved') {
        if (-not (Show-DiscardConfirm -Names @($Row.Name))) { return }
        $confirmedDiscard = $true
    }

    $result = Invoke-ComTask -Command @{ Op = $op; Key = $Row.Key; ConfirmedDiscard = $confirmedDiscard } -TimeoutMs 10000 -OwnerForm $script:MainForm -StatusText "Closing $($Row.Name)..."
    if ($result.NeedsConfirm) {
        if (-not (Show-DiscardConfirm -Names @($Row.Name))) { return }
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
function Invoke-BatchOp {
    param([array]$Rows, [string]$ActionLabel, [scriptblock]$OpNameForApp, [bool]$ConfirmedDiscard = $false)
    $timedOutApps = @{}
    $counts = @{ Ok = 0; Failed = 0; Skipped = 0 }
    foreach ($row in $Rows) {
        if ($timedOutApps.ContainsKey($row.App)) { continue }
        $op = & $OpNameForApp $row.App
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

function Invoke-SaveAll {
    $all = @($script:Grid.Tag)
    $rows = @($all | Where-Object { $_.CanSave -and ($_.RowType -eq 'Workbook' -or $_.RowType -eq 'Inspector') -and $_.Status -ne 'Never saved' })
    $skipped = @($all | Where-Object { $_.RowType -eq 'Workbook' -and $_.Status -eq 'Never saved' }).Count
    $counts = Invoke-BatchOp -Rows $rows -ActionLabel 'SaveClose' -OpNameForApp { param($app) if ($app -eq 'Excel') { 'ExcelSaveClose' } else { 'OutlookSaveClose' } }
    Do-RefreshForApps -Apps @($rows | ForEach-Object { $_.App })
    [System.Windows.Forms.MessageBox]::Show($script:MainForm, "Saved and closed: $($counts.Ok)`r`nSkipped (never saved): $skipped`r`nFailed: $($counts.Failed)", 'Save & Close All', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

# "Zero-risk, no confirm" per design - but that premise only holds if the
# cached Status is still accurate at click time, which nothing guarantees
# (Nick could have edited a row after the last Refresh). Every row is sent
# with ConfirmedDiscard=$false; the worker's live re-check catches any row
# that has actually gone unsaved since, and it is skipped rather than closed.
function Invoke-CloseAllSaved {
    $rows = @($script:Grid.Tag | Where-Object { $_.CanClose -and $_.Status -eq 'Saved' -and ($_.RowType -eq 'Workbook' -or $_.RowType -eq 'Inspector') })
    $counts = Invoke-BatchOp -Rows $rows -ActionLabel 'Close' -OpNameForApp { param($app) if ($app -eq 'Excel') { 'ExcelClose' } else { 'OutlookClose' } } -ConfirmedDiscard $false
    Do-RefreshForApps -Apps @($rows | ForEach-Object { $_.App })
    [System.Windows.Forms.MessageBox]::Show($script:MainForm, "Closed: $($counts.Ok)`r`nChanged since refresh - skipped: $($counts.Skipped)`r`nFailed: $($counts.Failed)", 'Close All Saved', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

function Invoke-CloseAllDiscard {
    $rows = @($script:Grid.Tag | Where-Object { $_.CanClose -and ($_.Status -eq 'UNSAVED' -or $_.Status -eq 'Never saved') -and ($_.RowType -eq 'Workbook' -or $_.RowType -eq 'Inspector') })
    if ($rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($script:MainForm, 'Nothing unsaved to discard.', 'Close All and Discard', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    if (-not (Show-DiscardConfirm -Names @($rows | ForEach-Object { $_.Name }))) { return }
    # Names were already listed and confirmed above, so every row in this
    # batch is sent pre-confirmed.
    $counts = Invoke-BatchOp -Rows $rows -ActionLabel 'Close' -OpNameForApp { param($app) if ($app -eq 'Excel') { 'ExcelClose' } else { 'OutlookClose' } } -ConfirmedDiscard $true
    Do-RefreshForApps -Apps @($rows | ForEach-Object { $_.App })
    [System.Windows.Forms.MessageBox]::Show($script:MainForm, "Closed (discarded): $($counts.Ok)`r`nFailed: $($counts.Failed)", 'Close All and Discard', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

function New-MainForm {
    $settings = Get-Settings
    $script:CurrentFilter = $settings.defaultFilter

    # v0.3 test hooks - zero effect unless one of the env vars is set. Read
    # once here for the banner/title; the actual row-dropping happens inside
    # the worker's EnumExcel/EnumOutlook (same env vars, inherited by the
    # runspace since environment variables are process-wide).
    $testFilterPattern = $env:WINDOWSWEEPER_TEST_FILTER
    $testPidEnv = $env:WINDOWSWEEPER_TEST_PID
    $testActive = (-not [string]::IsNullOrEmpty($testFilterPattern)) -or (-not [string]::IsNullOrEmpty($testPidEnv))

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Window Sweeper v0.3'
    if ($testActive) { $form.Text = $form.Text + ' [TEST]' }
    $form.Width = 920
    $form.Height = 600
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(720, 420)
    $form.TopMost = $settings.keepOnTop

    $script:TopPanel = New-Object System.Windows.Forms.Panel
    $script:TopPanel.Dock = 'Top'
    $script:TopPanel.Height = 40

    $rbBoth = New-Object System.Windows.Forms.RadioButton
    $rbBoth.Text = 'Both'; $rbBoth.Location = New-Object System.Drawing.Point(10, 11); $rbBoth.AutoSize = $true
    $rbExcel = New-Object System.Windows.Forms.RadioButton
    $rbExcel.Text = 'Excel'; $rbExcel.Location = New-Object System.Drawing.Point(75, 11); $rbExcel.AutoSize = $true
    $rbOutlook = New-Object System.Windows.Forms.RadioButton
    $rbOutlook.Text = 'Outlook'; $rbOutlook.Location = New-Object System.Drawing.Point(150, 11); $rbOutlook.AutoSize = $true

    if ($script:CurrentFilter -eq 'Excel') { $rbExcel.Checked = $true }
    elseif ($script:CurrentFilter -eq 'Outlook') { $rbOutlook.Checked = $true }
    else { $rbBoth.Checked = $true }

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = 'Refresh'; $btnRefresh.Location = New-Object System.Drawing.Point(240, 7); $btnRefresh.Width = 80

    # Default checked - this is what makes row-click peeking actually usable:
    # the target workbook/email fronts UNDER the sweeper instead of hiding it.
    $chkKeepOnTop = New-Object System.Windows.Forms.CheckBox
    $chkKeepOnTop.Text = 'Keep on top'; $chkKeepOnTop.Location = New-Object System.Drawing.Point(330, 11); $chkKeepOnTop.AutoSize = $true
    $chkKeepOnTop.Checked = $settings.keepOnTop

    $script:LblTimestamp = New-Object System.Windows.Forms.Label
    $script:LblTimestamp.Location = New-Object System.Drawing.Point(460, 12); $script:LblTimestamp.AutoSize = $true
    $script:LblTimestamp.ForeColor = [System.Drawing.Color]::Gray

    $script:TopPanel.Controls.AddRange([System.Windows.Forms.Control[]]@($rbBoth, $rbExcel, $rbOutlook, $btnRefresh, $chkKeepOnTop, $script:LblTimestamp))

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

    # v0.3 test filter banner - only built when a test env var is set.
    $testBanner = $null
    if ($testActive) {
        $testParts = @()
        if (-not [string]::IsNullOrEmpty($testFilterPattern)) { $testParts += "pattern='$testFilterPattern'" }
        if (-not [string]::IsNullOrEmpty($testPidEnv)) { $testParts += "pid=$testPidEnv" }
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

    $script:Grid = New-Object System.Windows.Forms.DataGridView
    $script:Grid.Dock = 'Fill'
    $script:Grid.AllowUserToAddRows = $false
    $script:Grid.AllowUserToDeleteRows = $false
    $script:Grid.ReadOnly = $true
    $script:Grid.RowHeadersVisible = $false
    $script:Grid.MultiSelect = $false
    $script:Grid.SelectionMode = 'FullRowSelect'
    $script:Grid.AutoSizeColumnsMode = 'Fill'
    $script:Grid.AllowUserToResizeRows = $false

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

    $script:Grid.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colApp, $colName, $colDetail, $colStatus, $colSave, $colClose))

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

    $rbBoth.Add_CheckedChanged({ if ($rbBoth.Checked) { $script:CurrentFilter = 'Both'; Do-FullRefresh } })
    $rbExcel.Add_CheckedChanged({ if ($rbExcel.Checked) { $script:CurrentFilter = 'Excel'; Do-FullRefresh } })
    $rbOutlook.Add_CheckedChanged({ if ($rbOutlook.Checked) { $script:CurrentFilter = 'Outlook'; Do-FullRefresh } })
    # Refresh button is the one and only click that always retries Outlook,
    # bypassing an active backoff - every other refresh path respects it.
    $btnRefresh.Add_Click({ Do-FullRefresh -ForceOutlookRetry $true })
    $btnSaveAll.Add_Click({ Invoke-SaveAll })
    $btnCloseSaved.Add_Click({ Invoke-CloseAllSaved })
    $btnCloseDiscard.Add_Click({ Invoke-CloseAllDiscard })
    $chkKeepOnTop.Add_CheckedChanged({ $script:MainForm.TopMost = $chkKeepOnTop.Checked })
    $script:BtnShowDialog.Add_Click({
        $app = $script:BtnShowDialog.Tag
        $procName = 'OUTLOOK'
        if ($app -eq 'Excel') { $procName = 'EXCEL' }
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
            $s.defaultFilter = $script:CurrentFilter
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
