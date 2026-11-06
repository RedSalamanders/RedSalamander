Set-StrictMode -Version Latest

function Initialize-RSHostedDesktopNative {
    if ('RedSalamander.Testing.HostedDesktopNative' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
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
        [DllImport("user32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CloseDesktop(IntPtr desktop);
        protected override bool ReleaseHandle() { return CloseDesktop(handle); }
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

Export-ModuleMember -Function Invoke-RSWithHostedTestDesktop, Get-RSHostedDesktopDiagnostics
