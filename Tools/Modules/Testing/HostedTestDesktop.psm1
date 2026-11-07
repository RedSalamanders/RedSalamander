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
        [StructLayout(LayoutKind.Sequential)] private struct KeyboardInput {
            public ushort VirtualKey, ScanCode;
            public uint Flags, Time;
            public UIntPtr ExtraInfo;
        }
        [StructLayout(LayoutKind.Explicit, Size = 40)] private struct NativeInput {
            [FieldOffset(0)] public uint Type;
            [FieldOffset(8)] public KeyboardInput Keyboard;
        }
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr BoundarySubclassProcedure(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam, UIntPtr subclass, UIntPtr data);
        [DllImport("comctl32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool SetWindowSubclass(IntPtr window, BoundarySubclassProcedure procedure, UIntPtr subclass, UIntPtr data);
        [DllImport("comctl32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool RemoveWindowSubclass(IntPtr window, BoundarySubclassProcedure procedure, UIntPtr subclass);
        [DllImport("comctl32.dll", ExactSpelling = true)] private static extern IntPtr DefSubclassProc(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool UnregisterHotKey(IntPtr window, int id);
        [DllImport("user32.dll", ExactSpelling = true)] private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)] private static extern uint SendInput(uint count, NativeInput[] input, int size);
        [DllImport("kernel32.dll", ExactSpelling = true)] private static extern uint GetCurrentProcessId();
        [DllImport("kernel32.dll", ExactSpelling = true)] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        private static extern SafeWindowHandle CreateWindowExW(uint extended, string className, string name, uint style,
            int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr parameter);
        [DllImport("user32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll", ExactSpelling = true)] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", ExactSpelling = true)] private static extern IntPtr SetFocus(IntPtr window);
        [DllImport("user32.dll", ExactSpelling = true)] private static extern IntPtr GetFocus();
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)] private static extern IntPtr GetWindowLongPtrW(IntPtr window, int index);
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)] private static extern IntPtr SetWindowLongPtrW(IntPtr window, int index, IntPtr value);
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);
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
        private readonly BoundarySubclassProcedure subclassProcedure;
        private static readonly UIntPtr BoundarySubclassId = new UIntPtr(0x52534244u);
        private const int BoundaryHotkeyId = 0x5253;
        private IntPtr activationWindow, activationOriginal, activationOriginalFocus;
        private uint activationOriginalThread, activationOriginalProcess;
        private bool hotkeyArmed, hotkeyHandled, hotkeyAcquired;
        private Exception hotkeyCallbackFailure;
        internal bool Restored { get; private set; }
        internal HostedForegroundBoundary(SafeDesktopHandle desktop) {
            this.desktop = desktop;
            subclassProcedure = BoundarySubclassCallback;
            joinThreshold = TimeSpan.FromSeconds(10);
            thread = new Thread(Run) { IsBackground = true, Name = "RedSalamander hosted foreground boundary" };
        }
        // Managed-only control: an already-started synthetic thread and short threshold;
        // it never invokes Run or opens a desktop/window. Production always uses ten seconds.
        internal HostedForegroundBoundary(Thread managedWorker, TimeSpan managedThreshold) {
            if (managedWorker == null || !managedWorker.IsAlive || managedThreshold <= TimeSpan.Zero)
                throw new ArgumentException("The managed drain control requires a live worker and positive threshold.");
            thread = managedWorker;
            subclassProcedure = BoundarySubclassCallback;
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
        internal static ushort[] RemainingOwnedChordKeyUps(uint downs, uint ups) {
            if (downs > 2) throw new ArgumentOutOfRangeException("downs");
            if (ups > downs) throw new ArgumentOutOfRangeException("ups");
            if (downs == ups) return Array.Empty<ushort>();
            if (downs == 1 || ups == 1) return new ushort[] { 0x12 };
            return new ushort[] { 0x87, 0x12 };
        }
        internal static bool AllowsOwnedBoundaryChord(string request, string actions, string runner, string repository, Architecture os, Architecture process) {
            return request == "1" && actions == "true" && runner == "github-hosted" && repository == "RedSalamanders/RedSalamander" &&
                os == Architecture.Arm64 && process == Architecture.Arm64;
        }
        private IntPtr BoundarySubclassCallback(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam, UIntPtr subclass, UIntPtr data) {
            if (message == 0x0312 && subclass == BoundarySubclassId && wParam.ToUInt64() == BoundaryHotkeyId && hotkeyArmed) {
                hotkeyHandled = true;
                // Capture named failures for the worker; never throw across WndProc.
                try {
                    uint process;
                    if (!canceled && window == activationWindow && IsWindow(window) && IsWindowVisible(window) &&
                        GetWindowThreadProcessId(window, out process) == GetCurrentThreadId() && process == GetCurrentProcessId()) {
                        if (GetForegroundWindow() == window && GetFocus() == window) {
                            if (HostedDesktopNative.InputDesktopName() != "Default")
                                throw new InvalidOperationException("Input desktop changed during owned hotkey admission.");
                            hotkeyAcquired = true;
                        } else {
                            VerifyOriginalForeground(activationOriginal, activationOriginalThread, activationOriginalProcess, activationOriginalFocus);
                            if (!canceled) {
                                bool accepted = SetForegroundWindow(window);
                                if (!canceled) {
                                    SetFocus(window);
                                    hotkeyAcquired = accepted && GetForegroundWindow() == window && GetFocus() == window;
                                }
                            }
                        }
                    }
                } catch (Win32Exception error) { hotkeyCallbackFailure = error; }
                  catch (InvalidOperationException error) { hotkeyCallbackFailure = error; }
                return IntPtr.Zero;
            }
            return DefSubclassProc(window, message, wParam, lParam);
        }
        private void AcquireWithOwnedHotkey(IntPtr handle, IntPtr original, uint originalThread, uint originalProcess, IntPtr originalFocus) {
            if (!AllowsOwnedBoundaryChord(Environment.GetEnvironmentVariable("RS_HOSTED_ARM64_STOCK_PREPARATION"),
                Environment.GetEnvironmentVariable("GITHUB_ACTIONS"), Environment.GetEnvironmentVariable("RUNNER_ENVIRONMENT"),
                Environment.GetEnvironmentVariable("GITHUB_REPOSITORY"), RuntimeInformation.OSArchitecture, RuntimeInformation.ProcessArchitecture))
                throw new InvalidOperationException("The hosted boundary did not acquire real foreground ownership; keyboard admission is limited to the selected native ARM64 role.");
            // ABI/layer exception: the compiler-free inline .NET host cannot include
            // TestSupport.Tests.DirectedSelfTestInputWarning.h. Mirror its owned
            // ALT+F24 acknowledgement/release protocol for this existing neutral
            // boundary only. Product suites keep their canonical native warning.
            VerifyOriginalForeground(original, originalThread, originalProcess, originalFocus);
            if (Marshal.SizeOf<NativeInput>() != 40 || Marshal.SizeOf<KeyboardInput>() != 24)
                throw new InvalidOperationException("Unexpected native keyboard input ABI.");
            activationWindow = handle; activationOriginal = original;
            activationOriginalThread = originalThread; activationOriginalProcess = originalProcess;
            activationOriginalFocus = originalFocus;
            hotkeyHandled = false; hotkeyAcquired = false; hotkeyCallbackFailure = null;
            bool subclassed = false, registered = false;
            uint downs = 0, ups = 0;
            Exception operationFailure = null;
            var cleanupFailures = new System.Collections.Generic.List<Exception>();
            try {
                if (canceled) throw new InvalidOperationException("Hosted keyboard admission was canceled before registration.");
                if (!SetWindowSubclass(handle, subclassProcedure, BoundarySubclassId, UIntPtr.Zero))
                    throw new InvalidOperationException("Cannot install the owned hosted hotkey callback.");
                subclassed = true;
                if (!RegisterHotKey(handle, BoundaryHotkeyId, 0x4001, 0x87))
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot register the owned hosted hotkey.");
                registered = true;
                VerifyOriginalForeground(original, originalThread, originalProcess, originalFocus);
                foreach (int key in new[] { 0x87, 0x10, 0x11, 0x12, 0x5b, 0x5c })
                    if ((GetAsyncKeyState(key) & 0x8000) != 0)
                        throw new InvalidOperationException("Held input prevents the owned hosted hotkey.");
                if (canceled) throw new InvalidOperationException("Hosted keyboard admission was canceled before input.");
                hotkeyArmed = true;
                var input = new[] {
                    new NativeInput { Type = 1, Keyboard = new KeyboardInput { VirtualKey = 0x12 } },
                    new NativeInput { Type = 1, Keyboard = new KeyboardInput { VirtualKey = 0x87 } }
                };
                downs = SendInput(2, input, Marshal.SizeOf<NativeInput>());
                if (downs != 2) throw new Win32Exception(Marshal.GetLastWin32Error(), "The owned hosted chord was only partially inserted.");
                var timer = System.Diagnostics.Stopwatch.StartNew();
                while (!canceled && !hotkeyHandled && timer.ElapsedMilliseconds < 500) {
                    Message message;
                    while (PeekMessageW(out message, handle, 0x0312, 0x0312, 1))
                        DispatchMessageW(ref message);
                    if (!hotkeyHandled && !canceled) {
                        uint wait = MsgWaitForMultipleObjectsEx(0, null, 10, 0x0080, 0x0004);
                        if (wait == 0xffffffff) throw new Win32Exception(Marshal.GetLastWin32Error(), "The owned hosted hotkey wait failed.");
                    }
                }
                if (hotkeyCallbackFailure != null)
                    throw new InvalidOperationException("The owned hosted hotkey callback refused activation.", hotkeyCallbackFailure);
                if (canceled || !hotkeyHandled || !hotkeyAcquired)
                    throw new InvalidOperationException("The real owned hosted hotkey did not acquire foreground and focus.");
            } catch (Win32Exception error) { operationFailure = error; }
              catch (InvalidOperationException error) { operationFailure = error; }
            finally {
                hotkeyArmed = false;
                // SendInput acknowledges an inserted prefix. Release only our own
                // inserted keys, in reverse order, including partial key-up insertion.
                for (int attempt = 0; attempt != 2 && ups < downs; ++attempt) {
                    ushort[] keys = RemainingOwnedChordKeyUps(downs, ups);
                    var release = new NativeInput[keys.Length];
                    for (int i = 0; i < keys.Length; ++i)
                        release[i] = new NativeInput { Type = 1, Keyboard = new KeyboardInput { VirtualKey = keys[i], Flags = 2 } };
                    ups += SendInput((uint)release.Length, release, Marshal.SizeOf<NativeInput>());
                }
                if (ups != downs) cleanupFailures.Add(new InvalidOperationException("The owned hosted chord did not acknowledge every key release."));
                if (registered && !UnregisterHotKey(handle, BoundaryHotkeyId))
                    cleanupFailures.Add(new Win32Exception(Marshal.GetLastWin32Error(), "Cannot unregister the owned hosted hotkey."));
                if (subclassed && !RemoveWindowSubclass(handle, subclassProcedure, BoundarySubclassId))
                    cleanupFailures.Add(new InvalidOperationException("Cannot retire the owned hosted hotkey callback."));
                try { Console.Error.WriteLine("HostedDefaultDesktop.BoundaryHotkey handled=" + hotkeyHandled + "; acquired=" + hotkeyAcquired + "; downs=" + downs + "; ups=" + ups + "; foreground=" + GetForegroundWindow().ToInt64() + "; focus=" + GetFocus().ToInt64()); }
                catch (System.IO.IOException) { /* Preserve input/drain results if the diagnostic stream is closed. */ }
                catch (ObjectDisposedException) { /* Preserve input/drain results if diagnostics were disposed. */ }
            }
            if (operationFailure != null) cleanupFailures.Insert(0, operationFailure);
            if (cleanupFailures.Count != 0) throw new AggregateException("Owned hosted keyboard admission failed.", cleanupFailures);
            if (HostedDesktopNative.InputDesktopName() != "Default" || GetForegroundWindow() != handle || GetFocus() != handle)
                throw new InvalidOperationException("The owned hosted hotkey foreground or focus was lost after release.");
        }
        internal static long BoundaryReturnStyle(long style) { return style & ~0x08000000L; }
        private void EnableBoundaryReturn(IntPtr handle) {
            uint process;
            if (canceled || !IsVerifiedInheritedDesktop(HostedDesktopNative.ReadDesktopContext(0)) || !IsWindow(handle) ||
                GetWindowThreadProcessId(handle, out process) != GetCurrentThreadId() || process != GetCurrentProcessId() ||
                GetForegroundWindow() != handle || GetFocus() != handle)
                throw new InvalidOperationException("Hosted boundary ownership changed before enabling automatic return.");
            Marshal.SetLastPInvokeError(0);
            long before = GetWindowLongPtrW(handle, -20).ToInt64();
            int error = Marshal.GetLastWin32Error();
            if (before == 0 && error != 0) throw new Win32Exception(error, "Cannot read hosted boundary extended style.");
            if ((before & 0x08000000L) == 0) throw new InvalidOperationException("Hosted boundary was not created without automatic activation.");
            if (canceled) throw new InvalidOperationException("Hosted boundary startup was canceled before style change.");
            long expected = BoundaryReturnStyle(before);
            Marshal.SetLastPInvokeError(0);
            IntPtr previous = SetWindowLongPtrW(handle, -20, new IntPtr(expected));
            error = Marshal.GetLastWin32Error();
            if (previous == IntPtr.Zero && error != 0) throw new Win32Exception(error, "Cannot enable hosted boundary return.");
            if (previous.ToInt64() != before) throw new InvalidOperationException("Hosted boundary extended style changed unexpectedly.");
            if (canceled) throw new InvalidOperationException("Hosted boundary startup was canceled during style change.");
            if (!SetWindowPos(handle, IntPtr.Zero, 0, 0, 0, 0, 0x0037))
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Cannot apply hosted boundary extended style.");
            Marshal.SetLastPInvokeError(0);
            long after = GetWindowLongPtrW(handle, -20).ToInt64();
            error = Marshal.GetLastWin32Error();
            if (after == 0 && error != 0) throw new Win32Exception(error, "Cannot verify hosted boundary extended style.");
            if (canceled || after != expected || !IsVerifiedInheritedDesktop(HostedDesktopNative.ReadDesktopContext(0)) ||
                !IsWindow(handle) || GetWindowThreadProcessId(handle, out process) != GetCurrentThreadId() || process != GetCurrentProcessId() ||
                GetForegroundWindow() != handle || GetFocus() != handle)
                throw new InvalidOperationException("Hosted boundary ownership or style changed during return setup.");
            try { Console.Error.WriteLine("HostedDefaultDesktop.BoundaryStyle before=" + before + "; after=" + after); }
            catch (System.IO.IOException) { /* Preserve setup result if diagnostics are unavailable. */ }
            catch (ObjectDisposedException) { /* Preserve setup result if diagnostics were disposed. */ }
        }
        internal static uint RestoreWaitMilliseconds(long elapsed) {
            if (elapsed < 0) throw new ArgumentOutOfRangeException(nameof(elapsed));
            return elapsed >= 1500 ? 0u : (uint)Math.Min(10, 1500 - elapsed);
        }
        private static void WaitOneRestoreMessage(System.Diagnostics.Stopwatch timer) {
            uint interval = RestoreWaitMilliseconds(timer.ElapsedMilliseconds);
            if (interval == 0) return;
            uint wait = MsgWaitForMultipleObjectsEx(0, null, interval, 0x04ff, 0x0004);
            if (wait == 0xffffffff) throw new Win32Exception(Marshal.GetLastWin32Error(), "The hosted restoration queue wait failed.");
            if (wait == 0) {
                Message message;
                if (PeekMessageW(out message, IntPtr.Zero, 0, 0, 1)) {
                    if (message.Id == 0x0012) throw new InvalidOperationException("The hosted restoration queue was quit.");
                    TranslateMessage(ref message);
                    DispatchMessageW(ref message);
                }
            }
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
                    WaitOneRestoreMessage(timer);
                    uint checkedProcess;
                    if (!IsVerifiedInheritedDesktop(HostedDesktopNative.ReadDesktopContext(0)) || !IsWindow(original) ||
                        GetWindowThreadProcessId(original, out checkedProcess) != originalThread || checkedProcess != originalProcess)
                        throw new InvalidOperationException("The hosted desktop or original identity changed during restoration.");
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
                    WaitOneRestoreMessage(timer);
                    uint checkedProcess;
                    if (!IsVerifiedInheritedDesktop(HostedDesktopNative.ReadDesktopContext(0)) || !IsWindow(original) ||
                        GetWindowThreadProcessId(original, out checkedProcess) != originalThread || checkedProcess != originalProcess)
                        throw new InvalidOperationException("The hosted desktop or original identity changed during focus restoration.");
                }
            }
        }
        // Pure predicate shared by the actual worker and managed-only controls.
        internal static bool IsVerifiedInheritedDesktop(DesktopContext context) {
            return context != null && context.WindowStation == "WinSta0" && context.WindowStationError == 0 &&
                context.ThreadDesktop == "Default" && context.ThreadDesktopError == 0 &&
                context.InputDesktop == "Default" && context.InputDesktopError == 0;
        }
        private void Run() {
            try {
                // The fresh worker inherits its process desktop. Verify the actual
                // worker and input context without narrowing its inherited handle rights.
                if (desktop == null || desktop.IsInvalid || !IsVerifiedInheritedDesktop(HostedDesktopNative.ReadDesktopContext(0)))
                    throw new InvalidOperationException("The hosted foreground worker did not inherit WinSta0/Default and the Default input desktop.");
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
                            AcquireWithOwnedHotkey(handle, original, originalThread, originalProcess, originalFocus);
                        if (canceled) throw new InvalidOperationException("Hosted boundary startup was canceled during foreground acquisition.");
                        SetFocus(handle);
                        if (canceled) throw new InvalidOperationException("Hosted boundary startup was canceled during focus acquisition.");
                        if (GetForegroundWindow() != handle || GetFocus() != handle)
                            throw new InvalidOperationException("The hosted boundary did not acquire real native focus.");
                        EnableBoundaryReturn(handle);
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
            // A failed subclass retirement may leave the callback installed until
            // the owner window/thread dies. Keep the delegate rooted through join.
            GC.KeepAlive(subclassProcedure);
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
                // The worker has drained before the retained lease handle closes; no
                // foreground work may outlive the interactive mutex.
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

function Initialize-RSHostedStockSurfaceNative {
    if ('RedSalamander.Testing.HostedStockSurface' -as [type]) { return }
    Initialize-RSHostedDesktopNative
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32.SafeHandles;
namespace RedSalamander.Testing {
    // Only the three explicitly approved hosted stock windows may be closed.
    // A retained process handle and a fresh HWND/thread/session check precede WM_CLOSE.
    public sealed class HostedStockSurface : IDisposable {
        private sealed class DesktopHandle : SafeHandleZeroOrMinusOneIsInvalid {
            public DesktopHandle() : base(true) {}
            [DllImport("user32.dll", ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool CloseDesktop(IntPtr h);
            protected override bool ReleaseHandle() { return CloseDesktop(handle); }
        }
        [DllImport("user32.dll", ExactSpelling=true, SetLastError=true)] static extern IntPtr GetProcessWindowStation();
        [DllImport("user32.dll", ExactSpelling=true, SetLastError=true)] static extern IntPtr GetThreadDesktop(uint tid);
        [DllImport("user32.dll", ExactSpelling=true, SetLastError=true)] static extern DesktopHandle OpenInputDesktop(uint flags,bool inherit,uint access);
        [DllImport("user32.dll", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool GetUserObjectInformationW(IntPtr h,int index,StringBuilder name,uint bytes,out uint needed);
        [DllImport("user32.dll", ExactSpelling=true)] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", ExactSpelling=true)] static extern IntPtr GetShellWindow();
        [DllImport("user32.dll", ExactSpelling=true)] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
        [DllImport("user32.dll", CharSet=CharSet.Unicode, ExactSpelling=true)] static extern int GetWindowTextW(IntPtr h,StringBuilder s,int n);
        [DllImport("user32.dll", CharSet=CharSet.Unicode, ExactSpelling=true)] static extern int GetClassNameW(IntPtr h,StringBuilder s,int n);
        [DllImport("user32.dll", ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool IsWindow(IntPtr h);
        [DllImport("user32.dll", ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll", ExactSpelling=true, SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool PostMessageW(IntPtr h,uint m,UIntPtr w,IntPtr l);
        [DllImport("kernel32.dll", ExactSpelling=true, SetLastError=true)] static extern SafeProcessHandle OpenProcess(uint access,bool inherit,uint pid);
        [DllImport("kernel32.dll", ExactSpelling=true)] static extern uint GetCurrentProcessId();
        [DllImport("kernel32.dll", ExactSpelling=true)] static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll", ExactSpelling=true, SetLastError=true)] static extern uint GetProcessId(SafeProcessHandle process);
        [DllImport("kernel32.dll", ExactSpelling=true, SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool ProcessIdToSessionId(uint pid,out uint session);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool QueryFullProcessImageNameW(SafeProcessHandle p,uint flags,StringBuilder s,ref uint n);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, ExactSpelling=true)] static extern int GetApplicationUserModelId(SafeProcessHandle p,ref uint n,StringBuilder s);
        [DllImport("dwmapi.dll", ExactSpelling=true)] static extern int DwmGetWindowAttribute(IntPtr h,uint attr,out uint value,uint size);
        [DllImport("kernel32.dll", ExactSpelling=true, SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool GetProcessTimes(SafeProcessHandle p,out long created,out long exited,out long kernel,out long user);
        private sealed class ArgumentHandle : SafeHandleZeroOrMinusOneIsInvalid {
            public ArgumentHandle():base(true) {}
            [DllImport("kernel32.dll", ExactSpelling=true)] static extern IntPtr LocalFree(IntPtr h);
            protected override bool ReleaseHandle(){return LocalFree(handle)==IntPtr.Zero;}
        }
        [DllImport("shell32.dll", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)] static extern ArgumentHandle CommandLineToArgvW(string command,out int count);
        private readonly SafeProcessHandle process;
        private readonly IntPtr window;
        private bool closed, wslAuthorized;
        private long wslCreationFileTime;
        public bool CloseRequested { get { return closed; } }
        public int WslArgumentCount { get; private set; }
        public bool WslArgumentMatch { get; private set; }
        public long WslCreationUtcTicks { get; private set; }
        public bool WslCreationMatch { get; private set; }
        public string Kind { get; private set; }
        public long Window { get { return window.ToInt64(); } }
        public uint Thread { get; private set; }
        public uint Process { get; private set; }
        public uint Session { get; private set; }
        public string Title { get; private set; }
        public string Class { get; private set; }
        public string Image { get; private set; }
        public string ApplicationId { get; private set; }
        public int ApplicationIdStatus { get; private set; }
        public string ForegroundDesktop { get; private set; }
        public int ForegroundDesktopError { get; private set; }
        private HostedStockSurface(IntPtr h,uint tid,uint pid,uint session,SafeProcessHandle p,string title,string cls,string image,string app,string kind) {
            window=h;Thread=tid;Process=pid;Session=session;process=p;Title=title;Class=cls;Image=image;ApplicationId=app;Kind=kind;
        }
        private static void RequireHostedArm64() {
            if(Environment.GetEnvironmentVariable("RS_HOSTED_ARM64_STOCK_PREPARATION")!="1" ||
               Environment.GetEnvironmentVariable("GITHUB_ACTIONS")!="true" || Environment.GetEnvironmentVariable("RUNNER_ENVIRONMENT")!="github-hosted" ||
               Environment.GetEnvironmentVariable("GITHUB_REPOSITORY")!="RedSalamanders/RedSalamander" ||
               RuntimeInformation.OSArchitecture!=Architecture.Arm64 || RuntimeInformation.ProcessArchitecture!=Architecture.Arm64)
                throw new InvalidOperationException("Stock surface preparation is limited to the explicitly selected project-hosted native ARM64 role.");
        }
        private static string ObjectName(IntPtr h) {
            if(h==IntPtr.Zero)throw new InvalidOperationException("Desktop or window station is unavailable.");
            var name=new StringBuilder(256);uint needed;
            if(!GetUserObjectInformationW(h,2,name,(uint)(name.Capacity*2),out needed))throw new Win32Exception(Marshal.GetLastWin32Error(),"Cannot identify desktop or station.");
            return name.ToString();
        }
        private static void RequireDefault(uint thread) {
            if(ObjectName(GetProcessWindowStation())!="WinSta0" || ObjectName(GetThreadDesktop(GetCurrentThreadId()))!="Default")
                throw new InvalidOperationException("Stock preparation requires WinSta0 and the Default caller desktop.");
            using(var input=OpenInputDesktop(0,false,1)) {
                if(input==null || input.IsInvalid || ObjectName(input.DangerousGetHandle())!="Default")
                    throw new InvalidOperationException("Stock preparation requires the Default input desktop.");
            }
        }
        private static string ReadTitle(IntPtr h) {var s=new StringBuilder(1024);GetWindowTextW(h,s,s.Capacity);return s.ToString();}
        private static string ReadClass(IntPtr h) {var s=new StringBuilder(256);GetClassNameW(h,s,s.Capacity);return s.ToString();}
        private static string ReadImage(SafeProcessHandle p) {uint n=4096;var s=new StringBuilder((int)n);if(!QueryFullProcessImageNameW(p,0,s,ref n))throw new Win32Exception(Marshal.GetLastWin32Error(),"Cannot identify the hosted foreground image.");return s.ToString();}
        private static string ReadApp(SafeProcessHandle p,out int status) {uint n=2048;var s=new StringBuilder((int)n);status=GetApplicationUserModelId(p,ref n,s);return status==0?s.ToString():"";}
        private static string Classify(IntPtr h,string title,string cls,string image,string app,int appStatus) {
            string windows=Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if(h==GetShellWindow() && cls=="Progman" && String.Equals(image,System.IO.Path.Combine(windows,"explorer.exe"),StringComparison.OrdinalIgnoreCase))return "Shell";
            if(title=="Microsoft account" && cls=="Windows.UI.Core.CoreWindow" && app=="Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy!App" &&
               String.Equals(image,System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),"WWAHost.exe"),StringComparison.OrdinalIgnoreCase))return "Account";
            if(title=="Search" && cls=="Windows.UI.Core.CoreWindow" && app=="MicrosoftWindows.Client.CBS_cw5n1h2txyewy!CortanaUI" &&
               String.Equals(image,System.IO.Path.Combine(windows,"SystemApps","MicrosoftWindows.Client.CBS_cw5n1h2txyewy","SearchHost.exe"),StringComparison.OrdinalIgnoreCase))return "Search";
            if(WslIdentityMatches(title,cls,image,app,appStatus,System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),"wsl.exe")))return "WslStartup";
            return "Unknown";
        }
        internal static bool WslIdentityMatches(string title,string cls,string image,string app,int appStatus,string expected) {
            return cls=="ConsoleWindowClass" && String.IsNullOrEmpty(app) && appStatus==15703 &&
                String.Equals(title,expected,StringComparison.OrdinalIgnoreCase) && String.Equals(image,expected,StringComparison.OrdinalIgnoreCase);
        }
        internal static bool WslArgumentsMatch(string[] args,string image) {
            return args!=null && args.Length==4 &&
                (String.Equals(args[0],image,StringComparison.OrdinalIgnoreCase) || String.Equals(args[0],"wsl.exe",StringComparison.OrdinalIgnoreCase)) &&
                args[1]=="--update" && args[2]=="--confirm" && args[3]=="--prompt-before-exit";
        }
        internal static bool WslCreationMatches(long capturedTicks,long queriedTicks) {
            // CIM timestamps have microsecond precision; accept only its sub-microsecond rounding difference.
            return capturedTicks>0 && queriedTicks>0 && Math.Abs(capturedTicks-queriedTicks)<10;
        }
        public void AuthorizeWsl(string command,long queriedCreationTicks) {
            RequireHostedArm64(); RequireDefault(Thread);
            if(Kind!="WslStartup" || closed || process==null || process.IsInvalid)throw new InvalidOperationException("WSL authorization requires the captured approved console.");
            long created,exited,kernel,user;
            if(!GetProcessTimes(process,out created,out exited,out kernel,out user))throw new Win32Exception(Marshal.GetLastWin32Error(),"Cannot bind WSL process creation.");
            WslCreationUtcTicks=DateTime.FromFileTimeUtc(created).Ticks;
            WslCreationMatch=WslCreationMatches(WslCreationUtcTicks,queriedCreationTicks);
            if(!WslCreationMatch)throw new InvalidOperationException("WSL process creation identity does not match CIM.");
            if(String.IsNullOrWhiteSpace(command))throw new InvalidOperationException("WSL startup argument sequence is unavailable or mismatched.");
            int count=0; using(var argv=CommandLineToArgvW(command,out count)) {
                WslArgumentCount=count;
                if(argv==null || argv.IsInvalid || count!=4)throw new InvalidOperationException("WSL startup argument sequence is unavailable or mismatched.");
                var values=new string[count];for(int i=0;i<count;i++)values[i]=Marshal.PtrToStringUni(Marshal.ReadIntPtr(argv.DangerousGetHandle(),i*IntPtr.Size));
                WslArgumentMatch=WslArgumentsMatch(values,Image);
                if(!WslArgumentMatch)throw new InvalidOperationException("WSL startup argument sequence is unavailable or mismatched.");
            }
            wslCreationFileTime=created;wslAuthorized=true;
        }
        public static HostedStockSurface Capture() {
            RequireHostedArm64();IntPtr h=GetForegroundWindow();uint pid;uint tid=GetWindowThreadProcessId(h,out pid);RequireDefault(tid);
            if(h==IntPtr.Zero && pid==0 && tid==0 && GetForegroundWindow()==IntPtr.Zero)return new HostedStockSurface(h,0,0,0,null,"","","","","Empty");
            if(h==IntPtr.Zero || pid==0 || tid==0)throw new InvalidOperationException("Foreground identity is unavailable.");
            uint session,own;if(!ProcessIdToSessionId(pid,out session) || !ProcessIdToSessionId(GetCurrentProcessId(),out own) || session!=own)
                throw new InvalidOperationException("Foreground and runner session differ or are unavailable.");
            var p=OpenProcess(0x1000,false,pid);if(p==null || p.IsInvalid){int error=Marshal.GetLastWin32Error();if(p!=null)p.Dispose();throw new Win32Exception(error,"Cannot retain the foreground process.");}
            bool transferred=false;
            try {
                int appStatus;string title=ReadTitle(h),cls=ReadClass(h),image=ReadImage(p),app=ReadApp(p,out appStatus);
                uint currentPid;uint currentTid=GetWindowThreadProcessId(h,out currentPid);
                if(GetProcessId(p)!=pid || GetForegroundWindow()!=h || currentPid!=pid || currentTid!=tid)throw new InvalidOperationException("Foreground changed while captured.");
                string kind=Classify(h,title,cls,image,app,appStatus);
                if(kind=="WslStartup" && !IsWindowVisible(h))throw new InvalidOperationException("WSL startup console is not visible.");
                var result=new HostedStockSurface(h,tid,pid,session,p,title,cls,image,app,kind);result.ApplicationIdStatus=appStatus;
                // Foreign thread desktop access is often denied by hosted Windows. It is
                // diagnostic only; WinSta0, caller/input Default and session are required.
                IntPtr foregroundDesktop=GetThreadDesktop(tid);
                if(foregroundDesktop==IntPtr.Zero)result.ForegroundDesktopError=Marshal.GetLastWin32Error();
                else {var name=new StringBuilder(256);uint needed;if(GetUserObjectInformationW(foregroundDesktop,2,name,(uint)(name.Capacity*2),out needed))result.ForegroundDesktop=name.ToString();else result.ForegroundDesktopError=Marshal.GetLastWin32Error();}
                transferred=true;return result;
            } finally { if(!transferred)p.Dispose(); }
        }
        private bool Retired() {
            if(!IsWindow(window))return true;
            uint pid;uint tid=GetWindowThreadProcessId(window,out pid);
            if(pid!=Process || tid!=Thread || GetForegroundWindow()==window)return false;
            bool hidden=!IsWindowVisible(window);
            if(!hidden && Kind=="Search"){uint mask;int hr=DwmGetWindowAttribute(window,14,out mask,sizeof(uint));hidden=hr>=0 && mask==2u;}
            if(!hidden)return false;
            if(!IsWindow(window))return true;
            tid=GetWindowThreadProcessId(window,out pid);return pid==Process && tid==Thread && GetForegroundWindow()!=window;
        }
        public void RequestClose() {
            RequireHostedArm64();RequireDefault(Thread);
            if(closed || (Kind!="Account" && Kind!="Search" && Kind!="WslStartup") || process==null || process.IsInvalid || (Kind=="WslStartup" && !wslAuthorized))throw new InvalidOperationException("Only one close for a captured approved stock surface is allowed.");
            uint pid;uint tid=GetWindowThreadProcessId(window,out pid);uint session, ownSession;int appStatus;string app=ReadApp(process,out appStatus);
            if(GetForegroundWindow()!=window || pid!=Process || tid!=Thread || GetProcessId(process)!=Process ||
               !ProcessIdToSessionId(pid,out session) || !ProcessIdToSessionId(GetCurrentProcessId(),out ownSession) || session!=Session || session!=ownSession || ReadTitle(window)!=Title || ReadClass(window)!=Class ||
               app!=ApplicationId || appStatus!=ApplicationIdStatus || !String.Equals(ReadImage(process),Image,StringComparison.OrdinalIgnoreCase) ||
               Classify(window,Title,Class,Image,ApplicationId,appStatus)!=Kind)
                throw new InvalidOperationException("Approved stock surface changed before close.");
            if(Kind=="WslStartup") {
                if(!IsWindowVisible(window))throw new InvalidOperationException("WSL startup console is no longer visible.");
                long created,exited,kernel,user;
                if(!GetProcessTimes(process,out created,out exited,out kernel,out user) || created!=wslCreationFileTime)
                    throw new InvalidOperationException("WSL process creation identity changed before close.");
            }
            closed=true;
            if(!PostMessageW(window,0x0010,UIntPtr.Zero,IntPtr.Zero))throw new Win32Exception(Marshal.GetLastWin32Error(),"Approved stock surface refused WM_CLOSE.");
            var timer=System.Diagnostics.Stopwatch.StartNew();
            while(!Retired() && timer.ElapsedMilliseconds<5000)System.Threading.Thread.Sleep(25);
            if(!Retired())throw new InvalidOperationException("Approved stock surface did not retire within five seconds.");
        }
        public void Dispose() { if(process!=null)process.Dispose(); }
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

function Get-RSHostedStockPreparationRequest {
    <#
    .SYNOPSIS
        Selects the approved stock-surface preparation for an eligible hosted ARM64 native role.
    .DESCRIPTION
        Returns 1 only for this repository's native ARM64 hosted ReleaseTest or an interactive
        Qualification. Local, x64, self-hosted, other repositories and paired studies return null.
        This pure policy performs no desktop calls. The native guard independently verifies the host.
    .PARAMETER Phase
        Canonical project adapter phase.
    .PARAMETER Platform
        Canonical native target platform.
    .PARAMETER QualificationScope
        Canonical project adapter scope; paired studies and read-only probes are excluded.
    .PARAMETER OsArchitecture
        Policy-test override; the native action independently verifies the actual OS architecture.
    .PARAMETER ProcessArchitecture
        Policy-test override; the native action independently verifies the actual process architecture.
    .OUTPUTS
        String 1 for the approved role, otherwise null.
    .EXAMPLE
        Get-RSHostedStockPreparationRequest -Phase ReleaseTest -Platform ARM64 -QualificationScope Normal
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Phase,[Parameter(Mandatory)][string]$Platform,
          [Parameter(Mandatory)][string]$QualificationScope,
          [string]$OsArchitecture=[Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString(),
          [string]$ProcessArchitecture=[Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString())
    if($Platform -cne 'ARM64' -or $OsArchitecture -cne 'Arm64' -or $ProcessArchitecture -cne 'Arm64' -or
       $env:GITHUB_ACTIONS -cne 'true' -or $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or
       $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander' -or
       $QualificationScope -cin @('R0ePaired','DesktopProbe') -or
       ($Phase -cne 'ReleaseTest' -and ($Phase -cne 'Qualification' -or $QualificationScope -cnotin @('Normal','Full','CommandsHandoff','FileOpsVolumes','ViewerSpaceCapture')))) {return $null}
    '1'
}

# WSL alone needs query-only command-line and creation-time evidence; the held
# native process handle validates that CIM refers to the same process instance.
function Get-RSHostedWslProcess {
    param([Parameter(Mandatory)][uint32]$ProcessId)
    Get-CimInstance -ClassName Win32_Process -Filter ("ProcessId = {0}" -f $ProcessId) -ErrorAction Stop
}

function Get-RSHostedStockSurface {
    Initialize-RSHostedStockSurfaceNative
    [RedSalamander.Testing.HostedStockSurface]::Capture()
}

function Invoke-RSHostedArm64StockPreparation {
    <#
    .SYNOPSIS
        Prepare the hosted ARM64 Default desktop for an interactive native entry.
    .DESCRIPTION
        Under the runner's existing interactive mutex, inspects only the foreground and may
        request one ordinary close each for the explicitly approved stock Microsoft account,
        Search and exact WSL update startup windows. WSL additionally requires matching
        held-process/CIM creation identity and the exact parsed update arguments. Exact image, AUMID, class, title, session, desktop and captured
        HWND/thread/process identity are checked with retained process handles. Unknown surfaces
        fail closed. A bounded five-second observation after account close catches the delayed
        Search transition; no input, process termination, generic window dismissal or policy
        change is performed. Each close has a create-only intent and result under the marked root.
    .PARAMETER RepoRoot
        Project checkout for the current runner entry and sandbox ownership check.
    .OUTPUTS
        Hosted preparation audit directory path. Does not establish a test or qualification pass.
    .EXAMPLE
        Invoke-RSHostedArm64StockPreparation -RepoRoot $PWD
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RepoRoot)
    if($env:RS_HOSTED_ARM64_STOCK_PREPARATION -cne '1' -or
       $env:GITHUB_ACTIONS -cne 'true' -or $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or
       $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander' -or
       [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -cne 'Arm64' -or
       [Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString() -cne 'Arm64'){
        throw 'Stock-surface preparation is restricted to the selected project-hosted native ARM64 role.'
    }
    Import-Module (Join-Path $PSScriptRoot 'TestSandbox.psm1')
    $root=Assert-RSTestSandboxContainedPath -TestRoot $env:REDSALAMANDER_TEST_ROOT -Path $env:REDSALAMANDER_TEST_ROOT
    $directory=Assert-RSTestSandboxContainedPath -TestRoot $root -Path (Join-Path $root ('evidence/hosted-stock-preparation/'+[guid]::NewGuid().ToString('N'))) -RequireDescendant
    if(Test-Path -LiteralPath $directory){throw 'Stock preparation audit directory already exists.'}
    [void][IO.Directory]::CreateDirectory($directory)
    $sourceHead = (& git -C $RepoRoot rev-parse HEAD | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or $sourceHead -cnotmatch '^[0-9a-f]{40}$') { throw 'Cannot bind stock preparation to the exact checkout head.' }
    Invoke-RSHostedStockPreparationCore -Directory $directory -SourceHead $sourceHead
    Write-Host ("Hosted ARM64 stock preparation audit: $directory")
    return $directory
}

# Internal state machine is separated so fixtures can mock observations and closes;
# the guarded entry above always supplies the fixed five-second observation window.
function Invoke-RSHostedStockPreparationCore {
    param([Parameter(Mandatory)][string]$Directory,[Parameter(Mandatory)][string]$SourceHead,[int]$ObservationMilliseconds=5000)
    $seen=@{};$timer=[Diagnostics.Stopwatch]::StartNew();$preparationError=$null;$completed=$false;$lastObservation=$null;$failure=$null;$wslCloseRequested=$false
    $start=[ordered]@{schema='red-salamander.hosted-stock-preparation.v1';source_head=$SourceHead;started_utc=[DateTimeOffset]::UtcNow.ToString('o')}
    $startBytes=[Text.UTF8Encoding]::new($false).GetBytes(($start|ConvertTo-Json));$startStream=[IO.File]::Open((Join-Path $directory 'start.json'),[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
    try{$startStream.Write($startBytes,0,$startBytes.Length);$startStream.Flush($true)}finally{$startStream.Dispose()}
    try { while($true){
        $surface=Get-RSHostedStockSurface
        try {
            $lastObservation=[ordered]@{kind=$surface.Kind;window=$surface.Window;thread=$surface.Thread;process=$surface.Process;session=$surface.Session;title=$surface.Title;class=$surface.Class;image=$surface.Image;application_id=$surface.ApplicationId;foreground_desktop=$surface.ForegroundDesktop;foreground_desktop_error=$surface.ForegroundDesktopError}
            if($surface.Kind -cnotin @('Empty','Shell','Account','Search','WslStartup')){throw 'An unapproved foreground surface prevents hosted preparation.'}
            if($surface.Kind -cin @('Account','Search','WslStartup')){
                if($seen.ContainsKey($surface.Kind)){throw 'An approved stock surface reappeared after its one permitted close.'}
                $seen[$surface.Kind]=$true
                $intentPath=Join-Path $directory ($surface.Kind.ToLowerInvariant()+'-intent.json')
                $resultPath=Join-Path $directory ($surface.Kind.ToLowerInvariant()+'-result.json')
                $intent=[ordered]@{schema='red-salamander.hosted-stock-intent.v1';kind=$surface.Kind;window=$surface.Window;thread=$surface.Thread;process=$surface.Process;session=$surface.Session;
                    title=$surface.Title;class=$surface.Class;image=$surface.Image;application_id=$surface.ApplicationId;
                    foreground_desktop=$surface.ForegroundDesktop;foreground_desktop_error=$surface.ForegroundDesktopError;source_head=$SourceHead;requested_utc=[DateTimeOffset]::UtcNow.ToString('o')}
                $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($intent|ConvertTo-Json -Depth 4))
                $stream=[IO.File]::Open($intentPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
                try{$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)}finally{$stream.Dispose()}
                $errorText=$null;$success=$false;$closeFailure=$null;$resultFailure=$null;$wslCimCreationTicks=$null
                try{
                    if($surface.Kind -ceq 'WslStartup'){
                        $wsl=Get-RSHostedWslProcess -ProcessId $surface.Process
                        if($null -eq $wsl -or [uint32]$wsl.ProcessId -ne $surface.Process -or $null -eq $wsl.CreationDate -or
                           -not [string]::Equals([string]$wsl.ExecutablePath,$surface.Image,[StringComparison]::OrdinalIgnoreCase)){
                            throw 'WSL CIM process identity is unavailable or mismatched.'
                        }
                        $wslCimCreationTicks=$wsl.CreationDate.ToUniversalTime().Ticks
                        $surface.AuthorizeWsl([string]$wsl.CommandLine,$wslCimCreationTicks)
                    }
                    $surface.RequestClose();$success=$true
                }catch{$closeFailure=$_;$errorText=$_.Exception.Message}
                if($surface.Kind -ceq 'WslStartup'){$wslCloseRequested=$surface.CloseRequested}
                try {
                    $result=[ordered]@{schema='red-salamander.hosted-stock-result.v1';intent_sha256=(Get-FileHash -LiteralPath $intentPath -Algorithm SHA256).Hash;retired=$success;error=$errorText;observed_utc=[DateTimeOffset]::UtcNow.ToString('o')}
                    if($surface.Kind -ceq 'WslStartup'){$result['close_requested']=$wslCloseRequested;$result['creation_utc_ticks']=$surface.WslCreationUtcTicks;$result['cim_creation_utc_ticks']=$wslCimCreationTicks;$result['creation_match']=$surface.WslCreationMatch;$result['application_id_status']=$surface.ApplicationIdStatus;$result['argument_count']=$surface.WslArgumentCount;$result['arguments_match']=$surface.WslArgumentMatch}
                    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($result|ConvertTo-Json -Depth 4))
                    $stream=[IO.File]::Open($resultPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
                    try{$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)}finally{$stream.Dispose()}
                } catch { $resultFailure=$_ }
                if($closeFailure -and $resultFailure){throw [AggregateException]::new('Stock close and audit result both failed.', [Exception[]]@($closeFailure.Exception,$resultFailure.Exception))}
                if($closeFailure){throw $closeFailure}
                if($resultFailure){throw $resultFailure}
                if($surface.Kind -cin @('Account','WslStartup')){$timer.Restart()}
                continue
            }
        } finally {if($null -ne $surface){$surface.Dispose()}}
        if($timer.ElapsedMilliseconds -ge $ObservationMilliseconds){break}
        Start-Sleep -Milliseconds 100
    };$completed=$true } catch {$failure=$_;$preparationError=$_.Exception.Message}
    try {
        $completion=[ordered]@{schema='red-salamander.hosted-stock-preparation-result.v1';completed=$completed;error=$preparationError;
            account_close_requested=$seen.ContainsKey('Account');search_close_requested=$seen.ContainsKey('Search');wsl_startup_close_requested=$wslCloseRequested;last_observation=$lastObservation;finished_utc=[DateTimeOffset]::UtcNow.ToString('o')}
        $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($completion|ConvertTo-Json -Depth 5));$stream=[IO.File]::Open((Join-Path $directory 'completion.json'),[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
        try{$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)}finally{$stream.Dispose()}
    } catch {
        if($failure){throw [AggregateException]::new('Stock preparation and completion audit both failed.', [Exception[]]@($failure.Exception,$_.Exception))}
        throw
    }
    if($failure){throw $failure}
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
        After direct foreground refusal, the opted-in project-hosted native ARM64 role may
        use one guarded registered Alt+F24 handoff to the owned neutral window in 500 ms.
        This runner prelude precedes the child warning; product input retains the shared
        native warning. Held input, partial-release, focus and retirement failures reject it.
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
        With the project adapter's approved ARM64 opt-in, prepares only the three identified
        stock foreground surfaces before creating the boundary. Other windows, processes,
        hooks and foreground-lock settings are unchanged.
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
        if ($env:RS_HOSTED_ARM64_STOCK_PREPARATION -ceq '1') {
            [void](Invoke-RSHostedArm64StockPreparation -RepoRoot ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..'))))
        }
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

Export-ModuleMember -Function Invoke-RSWithHostedTestDesktop, Get-RSHostedDesktopDiagnostics, Invoke-RSWithHostedDefaultDesktop, Get-RSHostedInputDesktopRequest, Get-RSHostedStockPreparationRequest
