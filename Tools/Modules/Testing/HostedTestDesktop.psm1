Set-StrictMode -Version Latest

function Initialize-RSHostedDesktopNative {
    if ('RedSalamander.Testing.HostedDesktopNative' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
namespace RedSalamander.Testing {
    // DEVMODEW is a fixed 220-byte public Win32 structure. Preserve the complete
    // enumerated mode, including fields the test fixture does not change.
    [StructLayout(LayoutKind.Explicit, Size = 220)]
    public struct DisplayMode {
        [FieldOffset(68)] public ushort Size;
        [FieldOffset(72)] public uint Fields;
        [FieldOffset(168)] public uint BitsPerPixel;
        [FieldOffset(172)] public uint Width;
        [FieldOffset(176)] public uint Height;
        [FieldOffset(184)] public uint Frequency;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct Rectangle { public int Left, Top, Right, Bottom; }
    public sealed class SafeDesktopHandle : Microsoft.Win32.SafeHandles.SafeHandleZeroOrMinusOneIsInvalid {
        public SafeDesktopHandle() : base(true) {}
        public int ReleaseError;
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CloseDesktop(IntPtr desktop);
        protected override bool ReleaseHandle() {
            bool closed = CloseDesktop(handle);
            if (!closed) ReleaseError = Marshal.GetLastWin32Error();
            return closed;
        }
    }
    // Launcher-only foreground boundary: no application executable dependency and
    // no input or window on the user's original desktop. The window and its queue
    // belong to one dedicated thread for their entire lifetime.
    internal sealed class OwnedDesktopAnchor : IDisposable {
        private sealed class SafeWindowHandle : Microsoft.Win32.SafeHandles.SafeHandleZeroOrMinusOneIsInvalid {
            public SafeWindowHandle() : base(true) {}
            internal int ReleaseError { get; private set; }
            [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)] private static extern bool DestroyWindow(IntPtr window);
            protected override bool ReleaseHandle() {
                bool destroyed = DestroyWindow(handle);
                if (!destroyed) ReleaseError = Marshal.GetLastWin32Error();
                return destroyed;
            }
        }
        [StructLayout(LayoutKind.Sequential)] private struct Point { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] private struct Message {
            public IntPtr Window; public uint Id; public UIntPtr WParam; public IntPtr LParam;
            public uint Time; public Point Position; public uint Private;
        }
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool SetThreadDesktop(SafeDesktopHandle desktop);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        private static extern SafeWindowHandle CreateWindowExW(uint extended, string className, string name, uint style,
            int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr parameter);
        [DllImport("user32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll", ExactSpelling = true)] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", ExactSpelling = true)] private static extern IntPtr SetFocus(IntPtr window);
        [DllImport("user32.dll", ExactSpelling = true)] private static extern IntPtr GetFocus();
        [DllImport("user32.dll", ExactSpelling = true)] private static extern int GetSystemMetrics(int index);
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        private static extern uint MsgWaitForMultipleObjectsEx(uint count, IntPtr[] handles, uint milliseconds, uint mask, uint flags);
        [DllImport("user32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool PeekMessageW(out Message message, IntPtr window, uint first, uint last, uint remove);
        [DllImport("user32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool TranslateMessage(ref Message message);
        [DllImport("user32.dll", ExactSpelling = true)] private static extern IntPtr DispatchMessageW(ref Message message);
        private readonly SafeDesktopHandle desktop;
        private readonly AutoResetEvent stop = new AutoResetEvent(false);
        private readonly ManualResetEventSlim ready = new ManualResetEventSlim(false);
        private readonly Thread thread;
        private volatile Exception failure;
        private bool disposed;
        private OwnedDesktopAnchor(SafeDesktopHandle desktop) {
            this.desktop = desktop;
            thread = new Thread(Run) { IsBackground = true, Name = "RedSalamander hosted foreground boundary" };
        }
        internal static OwnedDesktopAnchor Start(SafeDesktopHandle desktop) {
            var anchor = new OwnedDesktopAnchor(desktop);
            anchor.thread.Start();
            if (!anchor.ready.Wait(TimeSpan.FromSeconds(10))) {
                anchor.Dispose();
                throw new TimeoutException("The owned hosted foreground boundary did not initialize.");
            }
            if (anchor.failure != null) {
                Exception error = anchor.failure;
                anchor.Dispose();
                throw new InvalidOperationException("Cannot establish the owned hosted foreground boundary.", error);
            }
            return anchor;
        }
        private void Run() {
            try {
                if (!SetThreadDesktop(desktop))
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot attach the owned foreground thread to its desktop.");
                var window = CreateWindowExW(0x80, "STATIC", "RedSalamander hosted test boundary", 0x90000000,
                    Math.Max(0, GetSystemMetrics(0) - 8), Math.Max(0, GetSystemMetrics(1) - 8), 8, 8,
                    IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                using (window) {
                    if (window == null || window.IsInvalid)
                        throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot create the owned hosted foreground boundary.");
                    IntPtr handle = window.DangerousGetHandle();
                    if (!SetForegroundWindow(handle))
                        throw new InvalidOperationException("The owned hosted boundary did not acquire real foreground ownership.");
                    SetFocus(handle);
                    if (GetForegroundWindow() != handle || GetFocus() != handle)
                        throw new InvalidOperationException("The owned hosted boundary did not acquire real native focus.");
                    ready.Set();
                    var waits = new[] { stop.SafeWaitHandle.DangerousGetHandle() };
                    for (;;) {
                        uint result = MsgWaitForMultipleObjectsEx(1, waits, 0xffffffff, 0x04ff, 0x0004);
                        if (result == 0) break;
                        if (result != 1) throw new Win32Exception(Marshal.GetLastWin32Error(), "The owned foreground queue wait failed.");
                        Message message;
                        while (PeekMessageW(out message, IntPtr.Zero, 0, 0, 1)) {
                            TranslateMessage(ref message);
                            DispatchMessageW(ref message);
                        }
                    }
                }
                if (window.ReleaseError != 0)
                    throw new Win32Exception(window.ReleaseError, "Cannot retire the owned hosted foreground window.");
            } catch (Win32Exception error) { failure = error; }
              catch (InvalidOperationException error) { failure = error; }
            finally { ready.Set(); }
        }
        public void Dispose() {
            if (disposed) return;
            stop.Set();
            if (!thread.Join(TimeSpan.FromSeconds(10)))
                throw new TimeoutException("The owned hosted foreground thread did not retire.");
            disposed = true;
            ready.Dispose(); stop.Dispose();
            if (failure != null) throw new InvalidOperationException("The owned hosted foreground boundary failed.", failure);
        }
    }
    public sealed class HostedOwnedDesktopLease : IDisposable {
        private readonly SafeDesktopHandle original;
        private readonly SafeDesktopHandle owned;
        private readonly string originalName;
        private OwnedDesktopAnchor anchor;
        private bool disposed;
        public readonly string Name;
        public bool Restored { get; private set; }
        public bool ForegroundBoundaryReady { get; private set; }
        internal HostedOwnedDesktopLease(SafeDesktopHandle original, string originalName, SafeDesktopHandle owned, string name) {
            this.original = original; this.originalName = originalName; this.owned = owned; Name = name;
        }
        internal void StartForegroundBoundary() {
            anchor = OwnedDesktopAnchor.Start(owned);
            ForegroundBoundaryReady = true;
        }
        public void Dispose() {
            if (disposed) return;
            disposed = true;
            try {
                string current = HostedDesktopNative.InputDesktopName();
                if (current == Name) {
                    if (!HostedDesktopNative.SwitchDesktop(original))
                        throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot restore the original hosted input desktop.");
                } else if (current != originalName) {
                    throw new InvalidOperationException("Another input desktop became active; it must not be replaced by the test fixture.");
                }
                Restored = HostedDesktopNative.InputDesktopName() == originalName;
                if (!Restored) throw new InvalidOperationException("Hosted input-desktop restoration did not match the original.");
            } finally {
                try { if (anchor != null) anchor.Dispose(); }
                finally { owned.Dispose(); original.Dispose(); }
            }
            if (owned.ReleaseError != 0 || original.ReleaseError != 0)
                throw new Win32Exception(owned.ReleaseError != 0 ? owned.ReleaseError : original.ReleaseError, "Cannot retire hosted desktop handles.");
        }
    }
    public sealed class DesktopContext {
        public string WindowStation = "";
        public int WindowStationError;
        public string ThreadDesktop = "";
        public int ThreadDesktopError;
        public string InputDesktop = "";
        public int InputDesktopError;
        public string ForegroundDesktop = "";
        public int ForegroundDesktopError;
    }
    public static class HostedDesktopNative {
        [DllImport("kernel32.dll", ExactSpelling = true)] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)] private static extern IntPtr GetProcessWindowStation();
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)] private static extern IntPtr GetThreadDesktop(uint thread);
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)] private static extern SafeDesktopHandle OpenInputDesktop(uint flags, [MarshalAs(UnmanagedType.Bool)] bool inherit, uint access);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        private static extern SafeDesktopHandle CreateDesktopW(string name, IntPtr device, IntPtr mode, uint flags, uint access, IntPtr security);
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)] public static extern bool SwitchDesktop(SafeDesktopHandle desktop);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool GetUserObjectInformationW(IntPtr handle, int index, StringBuilder text, uint bytes, out uint needed);
        private static string ObjectName(IntPtr handle, out int error) {
            error = 0;
            if (handle == IntPtr.Zero) { error = Marshal.GetLastWin32Error(); return ""; }
            var name = new StringBuilder(256);
            uint needed;
            if (!GetUserObjectInformationW(handle, 2, name, (uint)(name.Capacity * 2), out needed)) {
                error = Marshal.GetLastWin32Error(); return "";
            }
            return name.ToString();
        }
        public static string InputDesktopName() {
            using (var desktop = OpenInputDesktop(0, false, 1)) {
                if (desktop == null || desktop.IsInvalid)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot identify the current input desktop.");
                int error;
                string name = ObjectName(desktop.DangerousGetHandle(), out error);
                if (error != 0 || String.IsNullOrEmpty(name)) throw new Win32Exception(error, "Cannot read the current input desktop name.");
                return name;
            }
        }
        public static HostedOwnedDesktopLease EnterOwnedDesktop() {
            int error;
            string station = ObjectName(GetProcessWindowStation(), out error);
            if (error != 0 || station != "WinSta0") throw new InvalidOperationException("Hosted test desktops require WinSta0.");
            SafeDesktopHandle original = OpenInputDesktop(0, false, 0x101); // READOBJECTS | SWITCHDESKTOP
            if (original == null || original.IsInvalid) {
                int openError = Marshal.GetLastWin32Error();
                if (original != null) original.Dispose();
                throw new Win32Exception(openError, "Cannot retain the original hosted input desktop.");
            }
            SafeDesktopHandle owned = null;
            HostedOwnedDesktopLease lease = null;
            bool transferred = false;
            try {
                string originalName = ObjectName(original.DangerousGetHandle(), out error);
                if (error != 0 || originalName != "Default") throw new InvalidOperationException("The hosted input desktop must initially be Default.");
                string name = "RedSalamander.Test." + Guid.NewGuid().ToString("N");
                owned = CreateDesktopW(name, IntPtr.Zero, IntPtr.Zero, 0, 0x1ff, IntPtr.Zero);
                if (owned == null || owned.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot create the owned hosted test desktop.");
                lease = new HostedOwnedDesktopLease(original, originalName, owned, name);
                if (!SwitchDesktop(owned)) throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot activate the owned hosted test desktop.");
                lease.StartForegroundBoundary();
                transferred = true;
                return lease;
            } finally {
                if (!transferred) {
                    if (lease != null) lease.Dispose();
                    else { if (owned != null) owned.Dispose(); original.Dispose(); }
                }
            }
        }
        public static DesktopContext ReadDesktopContext(uint foregroundThread) {
            var result = new DesktopContext();
            result.WindowStation = ObjectName(GetProcessWindowStation(), out result.WindowStationError);
            result.ThreadDesktop = ObjectName(GetThreadDesktop(GetCurrentThreadId()), out result.ThreadDesktopError);
            // READOBJECTS only: neither desktop switching nor foreground permission is requested.
            using (var input = OpenInputDesktop(0, false, 1)) {
                if (input == null || input.IsInvalid) result.InputDesktopError = Marshal.GetLastWin32Error();
                else result.InputDesktop = ObjectName(input.DangerousGetHandle(), out result.InputDesktopError);
            }
            if (foregroundThread != 0)
                result.ForegroundDesktop = ObjectName(GetThreadDesktop(foregroundThread), out result.ForegroundDesktopError);
            return result;
        }
        [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumDisplaySettingsW(string device, int index, ref DisplayMode mode);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern int ChangeDisplaySettingsExW(string device, ref DisplayMode mode, IntPtr window, uint flags, IntPtr reserved);
        [DllImport("user32.dll", EntryPoint = "SystemParametersInfoW", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWorkArea(uint action, uint parameter, ref Rectangle area, uint flags);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern int GetSystemMetrics(int index);
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern int GetClassNameW(IntPtr window, StringBuilder name, int count);
    }
}
'@
}

function Get-RSHostedDesktopState {
    Initialize-RSHostedDesktopNative
    $mode = [RedSalamander.Testing.DisplayMode]::new()
    $mode.Size = 220
    if (-not [RedSalamander.Testing.HostedDesktopNative]::EnumDisplaySettingsW([NullString]::Value, -1, [ref]$mode)) {
        throw 'Cannot read the current hosted display mode.'
    }
    $work = [RedSalamander.Testing.Rectangle]::new()
    if (-not [RedSalamander.Testing.HostedDesktopNative]::GetWorkArea(0x30, 0, [ref]$work, 0)) {
        throw 'Cannot read the hosted desktop work area.'
    }
    $foreground = [RedSalamander.Testing.HostedDesktopNative]::GetForegroundWindow()
    $foregroundProcess = [uint32]0
    $foregroundThread = [RedSalamander.Testing.HostedDesktopNative]::GetWindowThreadProcessId($foreground, [ref]$foregroundProcess)
    $className = [Text.StringBuilder]::new(256)
    $null = [RedSalamander.Testing.HostedDesktopNative]::GetClassNameW($foreground, $className, $className.Capacity)
    $foregroundOwner = if ($foregroundProcess -eq 0) { $null } else { Get-Process -Id $foregroundProcess -ErrorAction SilentlyContinue }
    $processName = if ($null -eq $foregroundOwner) { '' } else { $foregroundOwner.ProcessName }
    $screenWidth = [RedSalamander.Testing.HostedDesktopNative]::GetSystemMetrics(0)
    if ($screenWidth -le 0) { throw 'Cannot determine the hosted display coordinate scale.' }
    [pscustomobject]@{
        Mode = $mode
        Width = [int]$mode.Width
        Height = [int]$mode.Height
        CoordinateScale = [double]$mode.Width / $screenWidth
        WorkWidth = $work.Right - $work.Left
        WorkHeight = $work.Bottom - $work.Top
        Session = [Diagnostics.Process]::GetCurrentProcess().SessionId
        ForegroundThread = $foregroundThread
        ForegroundProcess = $foregroundProcess
        ForegroundName = $processName
        ForegroundClass = $className.ToString()
        DesktopContext = [RedSalamander.Testing.HostedDesktopNative]::ReadDesktopContext($foregroundThread)
    }
}

function Get-RSHostedDesktopDiagnostics {
    <#
    .SYNOPSIS
        Reads hosted desktop identities and individual access errors.
    .DESCRIPTION
        Refuses outside GitHub Actions. Reads the calling window station and desktop,
        input desktop and foreground thread desktop, alongside display/session facts.
        Unavailable names retain their Windows error; they do not establish a mismatch.
        Does not send input, activate windows or change display/session settings.
    .OUTPUTS
        A diagnostic object, excluding the native display-mode restoration structure.
    .NOTES
        Read-only Windows GitHub Actions diagnostics; creates no qualification receipt.
    .EXAMPLE
        Get-RSHostedDesktopDiagnostics | ConvertTo-Json -Depth 5
    #>
    [CmdletBinding()]
    param()
    if ($env:GITHUB_ACTIONS -cne 'true') { throw 'Hosted desktop diagnostics are restricted to GitHub Actions.' }
    Get-RSHostedDesktopState | Select-Object * -ExcludeProperty Mode
}

function Get-RSHostedDisplayModes {
    Initialize-RSHostedDesktopNative
    for ($index = 0; $index -lt 4096; $index++) {
        $mode = [RedSalamander.Testing.DisplayMode]::new()
        $mode.Size = 220
        if (-not [RedSalamander.Testing.HostedDesktopNative]::EnumDisplaySettingsW([NullString]::Value, $index, [ref]$mode)) { return }
        $mode
    }
    throw 'Hosted display mode enumeration exceeded its safety bound.'
}

function Get-RSHostedInputDesktopRequest {
    <#
    .SYNOPSIS
        Selects the owned input-desktop fixture for this project's hosted native entries.
    .DESCRIPTION
        Returns owned only on this project's GitHub-hosted runner. Paired comparisons
        retain the original desktop so a historical baseline and current candidate use
        the same fixture. This policy performs no native calls or environment mutation.
    .PARAMETER PairedComparison
        Preserves the original input desktop for both sides of a matched study.
    .OUTPUTS
        The string owned, or null to retain the original desktop.
    .EXAMPLE
        Get-RSHostedInputDesktopRequest -PairedComparison:$false
    #>
    [CmdletBinding()]
    param([switch]$PairedComparison)
    if ($PairedComparison -or $env:GITHUB_ACTIONS -cne 'true' -or
        $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander') {
        return $null
    }
    'owned'
}

function New-RSHostedOwnedDesktop {
    Initialize-RSHostedDesktopNative
    [RedSalamander.Testing.HostedDesktopNative]::EnterOwnedDesktop()
}

function New-RSHostedDesktopProcessScope {
    param([Parameter(Mandatory)][string]$Name)
    Import-Module (Join-Path $PSScriptRoot '../Build/SanitizedEnvironment.psm1')
    [RedSalamander.Tooling.ContainedProcess]::UseOwnedDesktop(('WinSta0\' + $Name))
}

function Invoke-RSWithHostedOwnedDesktop {
    <#
    .SYNOPSIS
        Runs one leased interactive entry on a temporary owned hosted input desktop.
    .DESCRIPTION
        Hosted-only and called inside the runner's interactive mutex. Creates one uniquely
        named desktop, directs contained children there and uses the existing native input
        warning. After the synchronous action drains its contained children, restores the
        original input desktop and retires both handles. A thread-owned neutral window must
        acquire real foreground and focus before the action starts; its message queue stays
        live until restoration and is joined before its desktop handle closes.
        Refuses secure or unknown desktops.
        Native focus/input assertions are unchanged. A foreign desktop switch is preserved
        and reported as a fixture failure. Local use is refused before native calls.
    .PARAMETER Action
        Synchronous entry operation that closes its containment job before returning.
    .OUTPUTS
        Action output and desktop creation/restoration diagnostics.
    .NOTES
        Changes only the ephemeral hosted runner's input desktop during the leased entry.
        No foreign windows, processes, hooks or foreground-lock settings are changed.
    .EXAMPLE
        Invoke-RSWithHostedOwnedDesktop -Action { Invoke-OwnedNativeTestEntry }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Action)
    if ($env:GITHUB_ACTIONS -cne 'true' -or $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or
        $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander') {
        throw 'Owned input desktops are restricted to the project GitHub-hosted runner.'
    }
    $desktop = $null; $processScope = $null; $actionError = $null; $restoreError = $null
    try {
        $desktop = New-RSHostedOwnedDesktop
        if ($desktop.ForegroundBoundaryReady -cne $true) { throw 'The hosted owned input desktop has no verified foreground boundary.' }
        Write-Host ('Hosted owned input desktop entered: ' + $desktop.Name + '; foregroundBoundaryReady=' + $desktop.ForegroundBoundaryReady)
        $processScope = New-RSHostedDesktopProcessScope -Name $desktop.Name
        $result = & $Action
    } catch { $actionError = $_ } finally {
        try {
            try { if ($null -ne $processScope) { $processScope.Dispose() } }
            finally { if ($null -ne $desktop) { $desktop.Dispose() } }
            if ($null -ne $desktop) {
                Write-Host ('Hosted owned input desktop retired: ' + $desktop.Name + '; restored=' + $desktop.Restored)
                if ($desktop.Restored -ne $true) { throw 'The hosted owned input desktop did not restore its original input desktop.' }
            }
        } catch { $restoreError = $_ }
    }
    if ($null -ne $actionError -and $null -ne $restoreError) {
        throw [AggregateException]::new('Native entry and hosted desktop restoration both failed.', [Exception[]]@($actionError.Exception, $restoreError.Exception))
    }
    if ($null -ne $restoreError) { throw $restoreError }
    if ($null -ne $actionError) { throw $actionError }
    $result
}

function Set-RSHostedDisplayMode {
    param([Parameter(Mandatory)][object]$Mode)
    Initialize-RSHostedDesktopNative
    $nativeMode = $Mode
    $probe = [RedSalamander.Testing.HostedDesktopNative]::ChangeDisplaySettingsExW([NullString]::Value, [ref]$nativeMode, [IntPtr]::Zero, 2, [IntPtr]::Zero)
    if ($probe -ne 0) { throw "Hosted display mode validation failed: $probe." }
    $result = [RedSalamander.Testing.HostedDesktopNative]::ChangeDisplaySettingsExW([NullString]::Value, [ref]$nativeMode, [IntPtr]::Zero, 0, [IntPtr]::Zero)
    if ($result -ne 0) { throw "Hosted display mode change failed: $result." }
}

function Invoke-RSWithHostedTestDesktop {
    <#
    .SYNOPSIS
        Runs hosted native UI qualification with a sufficient display and restores it.
    .DESCRIPTION
        Refuses outside GitHub Actions. Reads the current display/foreground state, uses
        only a driver-enumerated 32-bit mode of at least 1920x1080 scaled for the display
        coordinates when required, verifies at least 1800x1000 of work area, runs the action and restores the original
        mode in finally. Does not change animation preferences, take screenshots, activate
        a window, install tools or reduce native test requirements. Unsupported setup and
        failed restoration are errors. Native harness warnings and desktop leases still apply.
    .PARAMETER Action
        Synchronous qualification action. Exceptions propagate after restoration.
    .OUTPUTS
        The action's output, with desktop diagnostics on the host stream.
    .NOTES
        Windows GitHub-hosted runner fixture only. No persistent display/registry changes.
        Callers: build-reusable.yml and nightly-flake.yml. Contract: Testing_SelfTests.md.
    .EXAMPLE
        Invoke-RSWithHostedTestDesktop -Action { ./Tools/Run-AllTests.ps1 -Suite CI -SkipBuild }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Action)
    if ($env:GITHUB_ACTIONS -cne 'true') { throw 'Hosted desktop setup is restricted to GitHub Actions.' }
    $before = Get-RSHostedDesktopState
    Write-Host ('Hosted desktop before: ' + ($before | Select-Object * -ExcludeProperty Mode | ConvertTo-Json -Compress))
    $changed = $false
    try {
        $minimumWidth = [Math]::Ceiling(1920 * $before.CoordinateScale)
        $minimumHeight = [Math]::Ceiling(1080 * $before.CoordinateScale)
        if ($before.Width -lt $minimumWidth -or $before.Height -lt $minimumHeight) {
            $mode = Get-RSHostedDisplayModes |
                Where-Object { $_.Width -ge $minimumWidth -and $_.Height -ge $minimumHeight -and $_.BitsPerPixel -eq 32 } |
                Sort-Object @{ Expression = { [int64]$_.Width * $_.Height } }, Frequency |
                Select-Object -First 1
            if ($null -eq $mode) { throw 'No supported hosted display mode can contain the required native capture viewports.' }
            $changed = $true
            Set-RSHostedDisplayMode -Mode $mode
        }
        $timer = [Diagnostics.Stopwatch]::StartNew()
        do {
            $ready = Get-RSHostedDesktopState
            if ($ready.WorkWidth -ge 1800 -and $ready.WorkHeight -ge 1000) { break }
            Start-Sleep -Milliseconds 100
        } while ($timer.ElapsedMilliseconds -lt 3000)
        Write-Host ('Hosted desktop ready: ' + ($ready | Select-Object * -ExcludeProperty Mode | ConvertTo-Json -Compress))
        if ($ready.WorkWidth -lt 1800 -or $ready.WorkHeight -lt 1000) { throw 'The hosted desktop work area cannot contain the required native UI fixtures.' }
        & $Action
    } finally {
        if ($changed) {
            Set-RSHostedDisplayMode -Mode $before.Mode
            $restored = Get-RSHostedDesktopState
            if ($restored.Width -ne $before.Width -or $restored.Height -ne $before.Height) { throw 'Hosted desktop display restoration did not converge.' }
            Write-Host ('Hosted desktop restored: ' + ($restored | Select-Object * -ExcludeProperty Mode | ConvertTo-Json -Compress))
        }
    }
}

Export-ModuleMember -Function Invoke-RSWithHostedTestDesktop, Get-RSHostedDesktopDiagnostics, Invoke-RSWithHostedOwnedDesktop, Get-RSHostedInputDesktopRequest
