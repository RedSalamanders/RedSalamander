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
    // A launcher-owned foreground window on the rendered Default desktop. WGC
    // on the hosted images cannot capture alternate Win32 desktops. All window
    // operations and its live message queue remain on one dedicated thread.
    internal sealed class HostedForegroundBoundary : IDisposable {
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
        [StructLayout(LayoutKind.Sequential)] private struct GuiThreadInfo {
            public uint Size, Flags;
            public IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
            public Rectangle CaretBounds;
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
        [DllImport("user32.dll", ExactSpelling = true)] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool GetGUIThreadInfo(uint thread, ref GuiThreadInfo info);
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
        private readonly TimeSpan joinThreshold;
        private volatile bool canceled;
        private volatile Exception failure;
        private bool threadStarted;
        private bool disposed;
        internal bool Restored { get; private set; }
        internal HostedForegroundBoundary(SafeDesktopHandle desktop) {
            this.desktop = desktop;
            joinThreshold = TimeSpan.FromSeconds(10);
            thread = new Thread(Run) { IsBackground = true, Name = "RedSalamander hosted foreground boundary" };
        }
        // Managed-only control: an already-started synthetic thread and short threshold;
        // it never invokes Run or opens a desktop/window. Production always uses ten seconds.
        internal HostedForegroundBoundary(Thread managedWorker, TimeSpan managedThreshold) {
            if (managedWorker == null || !managedWorker.IsAlive || managedThreshold <= TimeSpan.Zero)
                throw new ArgumentException("The managed drain control requires a live worker and positive threshold.");
            thread = managedWorker;
            threadStarted = true;
            joinThreshold = managedThreshold;
        }
        internal void Start() {
            thread.Start();
            threadStarted = true;
            if (!ready.Wait(TimeSpan.FromSeconds(10))) {
                Cancel();
                throw new TimeoutException("The hosted foreground boundary did not initialize.");
            }
            if (failure != null) {
                throw new InvalidOperationException("Cannot establish the hosted foreground boundary.", failure);
            }
        }
        private void Cancel() {
            // Do not wait for a native window call: a stuck call must not hold up
            // cancellation or the bounded join. An already executing call cannot
            // be revoked; the worker checks this sticky state before proceeding.
            canceled = true;
            stop.Set();
        }
        private static void VerifyOriginalForeground(IntPtr original, uint originalThread, uint originalProcess, IntPtr originalFocus) {
            if (HostedDesktopNative.InputDesktopName() != "Default" || GetForegroundWindow() != original)
                throw new InvalidOperationException("The hosted desktop or foreground changed before boundary acquisition.");
            if (original != IntPtr.Zero) {
                uint process;
                if (!IsWindow(original) || GetWindowThreadProcessId(original, out process) != originalThread || process != originalProcess ||
                    ReadFocus(originalThread).Focus != originalFocus || GetForegroundWindow() != original)
                    throw new InvalidOperationException("The hosted foreground identity or focus changed before boundary acquisition.");
            }
        }
        private static GuiThreadInfo ReadFocus(uint thread) {
            var info = new GuiThreadInfo { Size = (uint)Marshal.SizeOf<GuiThreadInfo>() };
            if (thread == 0) throw new InvalidOperationException("The original hosted foreground thread is unavailable.");
            if (!GetGUIThreadInfo(thread, ref info))
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot verify the original hosted foreground focus.");
            return info;
        }
        private void RestoreForeground(IntPtr handle, IntPtr original, uint originalThread, uint originalProcess, IntPtr originalFocus) {
            if (HostedDesktopNative.InputDesktopName() != "Default")
                throw new InvalidOperationException("A foreign input-desktop change must not be replaced by the hosted fixture.");
            IntPtr current = GetForegroundWindow();
            if (original != IntPtr.Zero) {
                uint ownerProcess;
                if (!IsWindow(original) || GetWindowThreadProcessId(original, out ownerProcess) != originalThread ||
                    ownerProcess != originalProcess)
                    throw new InvalidOperationException("The original hosted foreground window changed identity.");
            }
            if (current == handle && original != IntPtr.Zero) {
                if (!SetForegroundWindow(original))
                    throw new InvalidOperationException("Cannot restore the original hosted foreground window.");
                var timer = System.Diagnostics.Stopwatch.StartNew();
                do {
                    current = GetForegroundWindow();
                    if (current == original) break;
                    if (current != handle || timer.ElapsedMilliseconds >= 1500)
                        throw new InvalidOperationException("The original hosted foreground window was not restored.");
                    Thread.Sleep(10);
                } while (true);
            }
            // A previously empty foreground is verified after retiring our HWND.
            if (original == IntPtr.Zero && current == handle) return;
            if (current != original)
                throw new InvalidOperationException("A foreign foreground change must not be replaced by the hosted fixture.");
            if (original != IntPtr.Zero) {
                var timer = System.Diagnostics.Stopwatch.StartNew();
                while (ReadFocus(originalThread).Focus != originalFocus) {
                    if (GetForegroundWindow() != original || timer.ElapsedMilliseconds >= 1500)
                        throw new InvalidOperationException("The original hosted native focus was not restored.");
                    Thread.Sleep(10);
                }
            }
        }
        private void Run() {
            try {
                if (!SetThreadDesktop(desktop))
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot attach the hosted foreground thread to Default.");
                IntPtr original = GetForegroundWindow();
                uint originalProcess = 0;
                uint originalThread = original == IntPtr.Zero ? 0 : GetWindowThreadProcessId(original, out originalProcess);
                IntPtr originalFocus = original == IntPtr.Zero ? IntPtr.Zero : ReadFocus(originalThread).Focus;
                if (canceled) throw new InvalidOperationException("Hosted boundary startup was canceled before window creation.");
                var window = CreateWindowExW(0x08040080, "STATIC", "RedSalamander hosted test boundary", 0x90000000,
                    Math.Max(0, GetSystemMetrics(0) - 8), Math.Max(0, GetSystemMetrics(1) - 8), 8, 8,
                    IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                Exception actionFailure = null;
                Exception restoreFailure = null;
                using (window) {
                    if (window == null || window.IsInvalid)
                        throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot create the hosted foreground boundary.");
                    IntPtr handle = window.DangerousGetHandle();
                    try {
                        if (canceled) throw new InvalidOperationException("Hosted boundary startup was canceled before foreground acquisition.");
                        VerifyOriginalForeground(original, originalThread, originalProcess, originalFocus);
                        if (canceled) throw new InvalidOperationException("Hosted boundary startup was canceled before foreground acquisition.");
                        if (!SetForegroundWindow(handle))
                            throw new InvalidOperationException("The hosted boundary did not acquire real foreground ownership.");
                        if (canceled) throw new InvalidOperationException("Hosted boundary startup was canceled during foreground acquisition.");
                        SetFocus(handle);
                        if (canceled) throw new InvalidOperationException("Hosted boundary startup was canceled during focus acquisition.");
                        if (GetForegroundWindow() != handle || GetFocus() != handle)
                            throw new InvalidOperationException("The hosted boundary did not acquire real native focus.");
                        if (canceled) throw new InvalidOperationException("Hosted boundary startup was canceled before readiness.");
                        ready.Set();
                        var waits = new[] { stop.SafeWaitHandle.DangerousGetHandle() };
                        for (;;) {
                            uint result = MsgWaitForMultipleObjectsEx(1, waits, 0xffffffff, 0x04ff, 0x0004);
                            if (result == 0) break;
                            if (result != 1) throw new Win32Exception(Marshal.GetLastWin32Error(), "The hosted foreground queue wait failed.");
                            Message message;
                            while (PeekMessageW(out message, IntPtr.Zero, 0, 0, 1)) {
                                TranslateMessage(ref message);
                                DispatchMessageW(ref message);
                            }
                        }
                    } catch (Win32Exception error) { actionFailure = error; }
                      catch (InvalidOperationException error) { actionFailure = error; }
                    finally {
                        try { RestoreForeground(handle, original, originalThread, originalProcess, originalFocus); }
                        catch (Win32Exception error) { restoreFailure = error; }
                        catch (InvalidOperationException error) { restoreFailure = error; }
                    }
                }
                var failures = new System.Collections.Generic.List<Exception>();
                if (actionFailure != null) failures.Add(actionFailure);
                if (restoreFailure != null) failures.Add(restoreFailure);
                if (window.ReleaseError != 0) failures.Add(new Win32Exception(window.ReleaseError, "Cannot retire the hosted foreground window."));
                if (failures.Count != 0)
                    throw new AggregateException("Hosted boundary operation or retirement failed.", failures);
                Restored = GetForegroundWindow() == original;
                if (!Restored) throw new InvalidOperationException("The hosted foreground restoration did not match its original window.");
            } catch (Win32Exception error) { failure = error; }
              catch (InvalidOperationException error) { failure = error; }
              catch (AggregateException error) { failure = error; }
            finally { ready.Set(); }
        }
        public void Dispose() {
            if (disposed) return;
            Cancel();
            bool timeout = threadStarted && !thread.Join(joinThreshold);
            // Preserve the caller-owned interactive mutex while a boundary can still
            // touch foreground. Hosted supervision is the fail-stop if it never exits.
            if (timeout) {
                try { Console.Error.WriteLine("HostedDefaultDesktop.BoundaryJoinThresholdExceeded=1; retaining interactive lease until worker exit."); }
                catch (System.IO.IOException) { /* Preserve the timeout and continue safe drain if diagnostics fail. */ }
                catch (ObjectDisposedException) { /* Preserve the timeout and continue safe drain if diagnostics are closed. */ }
                thread.Join();
            }
            disposed = true;
            ready.Dispose(); stop.Dispose();
            if (timeout) throw new TimeoutException("The hosted foreground thread exceeded its retirement threshold and later drained.", failure);
            if (failure != null) throw new InvalidOperationException("The hosted foreground boundary failed.", failure);
        }
    }
    public sealed class HostedDefaultDesktopLease : IDisposable {
        private readonly SafeDesktopHandle desktop;
        private HostedForegroundBoundary boundary;
        private readonly object disposalLock = new object();
        private bool disposed;
        public readonly string Name = "Default";
        public bool Restored { get; private set; }
        public bool ForegroundBoundaryReady { get; private set; }
        internal HostedDefaultDesktopLease(SafeDesktopHandle desktop) { this.desktop = desktop; }
        internal void StartForegroundBoundary() {
            boundary = new HostedForegroundBoundary(desktop);
            boundary.Start();
            ForegroundBoundaryReady = true;
        }
        public void Dispose() {
            lock (disposalLock) {
                if (disposed) return;
                Exception desktopFailure = null;
                Exception boundaryFailure = null;
                try {
                    if (HostedDesktopNative.InputDesktopName() != Name)
                        throw new InvalidOperationException("A foreign input-desktop change must not be replaced by the hosted fixture.");
                } catch (Win32Exception error) { desktopFailure = error; }
                  catch (InvalidOperationException error) { desktopFailure = error; }
                try { if (boundary != null) boundary.Dispose(); }
                catch (TimeoutException timeout) { boundaryFailure = timeout; }
                catch (InvalidOperationException error) { boundaryFailure = error; }
                // The join completed before closing the handle; CloseDesktop rejects handles
                // still in use by a thread, including the boundary's own thread.
                desktop.Dispose();
                disposed = true;
                if (desktop.ReleaseError != 0) {
                    var closeFailure = new Win32Exception(desktop.ReleaseError, "Cannot retire the hosted input desktop handle.");
                    if (desktopFailure != null) desktopFailure = new AggregateException(desktopFailure, closeFailure);
                    else desktopFailure = closeFailure;
                }
                if (desktopFailure != null && boundaryFailure != null)
                    throw new AggregateException("Hosted desktop and boundary retirement both failed.", desktopFailure, boundaryFailure);
                if (desktopFailure != null) throw new InvalidOperationException("Hosted desktop retirement failed.", desktopFailure);
                if (boundaryFailure != null) throw new InvalidOperationException("Hosted boundary retirement failed.", boundaryFailure);
                Restored = boundary != null && boundary.Restored && HostedDesktopNative.InputDesktopName() == Name;
                if (!Restored) throw new InvalidOperationException("The hosted desktop and foreground restoration were not verified.");
            }
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
        public static HostedDefaultDesktopLease EnterDefaultDesktop() {
            int error;
            string station = ObjectName(GetProcessWindowStation(), out error);
            if (error != 0 || station != "WinSta0") throw new InvalidOperationException("Hosted UI tests require WinSta0.");
            string threadDesktop = ObjectName(GetThreadDesktop(GetCurrentThreadId()), out error);
            if (error != 0 || threadDesktop != "Default") throw new InvalidOperationException("The hosted caller must belong to Default.");
            SafeDesktopHandle desktop = OpenInputDesktop(0, false, 0x83); // READOBJECTS | CREATEWINDOW | WRITEOBJECTS
            if (desktop == null || desktop.IsInvalid) {
                int openError = Marshal.GetLastWin32Error();
                if (desktop != null) desktop.Dispose();
                throw new Win32Exception(openError, "Cannot retain the rendered hosted input desktop.");
            }
            HostedDefaultDesktopLease lease = null;
            Exception setupFailure = null;
            try {
                string name = ObjectName(desktop.DangerousGetHandle(), out error);
                if (error != 0 || name != "Default") throw new InvalidOperationException("The hosted input desktop must be Default.");
                lease = new HostedDefaultDesktopLease(desktop);
                lease.StartForegroundBoundary();
                return lease;
            }
            catch (Win32Exception failure) { setupFailure = failure; }
            catch (InvalidOperationException failure) { setupFailure = failure; }
            catch (TimeoutException failure) { setupFailure = failure; }
            catch (AggregateException failure) { setupFailure = failure; }
            Exception cleanupFailure = null;
            try { if (lease != null) lease.Dispose(); else desktop.Dispose(); }
            catch (Win32Exception failure) { cleanupFailure = failure; }
            catch (InvalidOperationException failure) { cleanupFailure = failure; }
            catch (TimeoutException failure) { cleanupFailure = failure; }
            catch (AggregateException failure) { cleanupFailure = failure; }
            if (cleanupFailure != null)
                throw new AggregateException("Hosted desktop setup and cleanup both failed.", setupFailure, cleanupFailure);
            throw new InvalidOperationException("Hosted desktop setup failed.", setupFailure);
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
        Selects the rendered Default-desktop fixture for this project's hosted native entries.
    .DESCRIPTION
        Returns default only on this project's GitHub-hosted runner. Paired comparisons
        retain the original desktop so a historical baseline and current candidate use
        the same fixture. This policy performs no native calls or environment mutation.
    .PARAMETER PairedComparison
        Preserves the original input desktop for both sides of a matched study.
    .OUTPUTS
        The string default, or null to retain the original desktop without a boundary.
    .EXAMPLE
        Get-RSHostedInputDesktopRequest -PairedComparison:$false
    #>
    [CmdletBinding()]
    param([switch]$PairedComparison)
    if ($PairedComparison -or $env:GITHUB_ACTIONS -cne 'true' -or
        $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander') {
        return $null
    }
    'default'
}

function New-RSHostedDefaultDesktop {
    Initialize-RSHostedDesktopNative
    [RedSalamander.Testing.HostedDesktopNative]::EnterDefaultDesktop()
}

function New-RSHostedDesktopProcessScope {
    param([Parameter(Mandatory)][string]$Name)
    Import-Module (Join-Path $PSScriptRoot '../Build/SanitizedEnvironment.psm1')
    if ($Name -cne 'Default') { throw 'Only the rendered Default desktop may be selected.' }
    [RedSalamander.Tooling.ContainedProcess]::UseHostedDefaultDesktop()
}

function Invoke-RSWithHostedDefaultDesktop {
    <#
    .SYNOPSIS
        Runs one leased hosted interactive entry on the rendered Default desktop.
    .DESCRIPTION
        Hosted-only and called inside the runner's interactive mutex. Requires WinSta0 and
        Default for both the caller and input desktop, explicitly directs contained children
        there and uses the existing native input warning. A thread-owned neutral window must
        acquire real foreground and focus before the action starts. After contained child drain,
        restores the original foreground and native focus, joins the live window thread and
        closes its retained desktop handle. No input desktop is created or switched.
        A ten-second retirement threshold records a failure; the caller retains its interactive
        lease until the worker exits. Hosted workflow supervision handles a worker that never drains.
        Refuses secure or unknown desktops.
        Native focus/input assertions are unchanged. A foreign desktop switch is preserved
        and reported as a fixture failure. Local use is refused before native calls.
    .PARAMETER Action
        Synchronous entry operation that closes its containment job before returning.
    .OUTPUTS
        Action output and boundary/readiness/restoration diagnostics.
    .NOTES
        Owns only a neutral foreground window on the ephemeral hosted runner's Default desktop.
        No foreign windows, processes, hooks or foreground-lock settings are changed.
    .EXAMPLE
        Invoke-RSWithHostedDefaultDesktop -Action { Invoke-OwnedNativeTestEntry }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Action)
    if ($env:GITHUB_ACTIONS -cne 'true' -or $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or
        $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander') {
        throw 'Hosted foreground boundaries are restricted to the project GitHub-hosted runner.'
    }
    $desktop = $null; $processScope = $null; $actionError = $null; $restoreError = $null
    try {
        $desktop = New-RSHostedDefaultDesktop
        if ($desktop.ForegroundBoundaryReady -cne $true) { throw 'The hosted Default desktop has no verified foreground boundary.' }
        Write-Host ('Hosted Default desktop entered: ' + $desktop.Name + '; foregroundBoundaryReady=' + $desktop.ForegroundBoundaryReady)
        $processScope = New-RSHostedDesktopProcessScope -Name $desktop.Name
        $result = & $Action
    } catch { $actionError = $_ } finally {
        try {
            try { if ($null -ne $processScope) { $processScope.Dispose() } }
            finally { if ($null -ne $desktop) { $desktop.Dispose() } }
            if ($null -ne $desktop) {
                Write-Host ('Hosted Default desktop retired: ' + $desktop.Name + '; restored=' + $desktop.Restored)
                if ($desktop.Restored -ne $true) { throw 'The hosted Default desktop did not restore its original foreground and focus.' }
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
        Direct callers include build-reusable.yml, nightly-flake.yml and
        ghostty-runtime-product-qualification.yml. Contract: Testing_SelfTests.md.
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

Export-ModuleMember -Function Invoke-RSWithHostedTestDesktop, Get-RSHostedDesktopDiagnostics, Invoke-RSWithHostedDefaultDesktop, Get-RSHostedInputDesktopRequest
