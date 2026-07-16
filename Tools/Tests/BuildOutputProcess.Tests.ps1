Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$buildScript = Join-Path $repoRoot 'build.ps1'

Describe 'Build output process preflight' {
    BeforeAll {
        $tokens = $null
        $parseErrors = $null
        $buildAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $buildScript,
            [ref]$tokens,
            [ref]$parseErrors)
        if ($parseErrors.Count -ne 0) {
            throw "build.ps1 did not parse: $($parseErrors[0].Message)"
        }

        foreach ($functionName in @('Test-BuildOutputSelfTestCommandLine', 'Stop-BuildOutputProcess')) {
            $functionAst = $buildAst.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -eq $functionName
            }, $true)
            if ($null -eq $functionAst) {
                throw "build.ps1 function '$functionName' was not found."
            }
            Invoke-Expression $functionAst.Extent.Text
        }
    }

    BeforeEach {
        $script:buildOutputProcesses = @()
        Mock Get-CimInstance { return $script:buildOutputProcesses }
        Mock Stop-Process {}
    }

    It 'recognizes suite and option self-test command-line flags without matching ordinary arguments' {
        $exe = 'Z:\src\RedSalamander\.build\x64\Debug\RedSalamander.exe'
        foreach ($commandLine in @(
            "`"$exe`" --selftest",
            "`"$exe`" --commands-selftest --selftest-repeat=100",
            "`"$exe`" --selftest-case=cmd_pane_fileops_completed_group_and_navigation",
            "`"$exe`" `"--monitor-scrollback-selftest`""
        )) {
            (Test-BuildOutputSelfTestCommandLine -CommandLine $commandLine) | Should Be $true
        }

        (Test-BuildOutputSelfTestCommandLine -CommandLine "`"$exe`"") | Should Be $false
        (Test-BuildOutputSelfTestCommandLine -CommandLine "`"$exe`" --output=selftest") | Should Be $false
    }

    It 'aborts with diagnostics and kills nothing when any exact-path process is a self-test' {
        $expectedPath = Join-Path $TestDrive 'Debug\RedSalamander.exe'
        $script:buildOutputProcesses = @(
            [pscustomobject]@{
                ProcessId = 4100
                ExecutablePath = $expectedPath
                CommandLine = "`"$expectedPath`""
            },
            [pscustomobject]@{
                ProcessId = 4101
                ExecutablePath = $expectedPath
                CommandLine = "`"$expectedPath`" --commands-selftest --selftest-repeat=100"
            }
        )

        $errorMessage = $null
        try {
            Stop-BuildOutputProcess -ProcessName 'RedSalamander.exe' -ExpectedExePath $expectedPath
        }
        catch {
            $errorMessage = $_.Exception.Message
        }

        $errorMessage | Should Match 'Build canceled because an active self-test may be using a target output'
        $errorMessage | Should Match 'PID=4101'
        $errorMessage | Should Match ([regex]::Escape([System.IO.Path]::GetFullPath($expectedPath)))
        $errorMessage | Should Match ([regex]::Escape('--commands-selftest --selftest-repeat=100'))
        Assert-MockCalled Stop-Process -Times 0
    }

    It 'refuses to kill an exact-path process when its command line is unavailable' {
        $expectedPath = Join-Path $TestDrive 'Debug\RedSalamander.exe'
        $script:buildOutputProcesses = @(
            [pscustomobject]@{
                ProcessId = 4201
                ExecutablePath = $expectedPath
                CommandLine = $null
            }
        )

        $errorMessage = $null
        try {
            Stop-BuildOutputProcess -ProcessName 'RedSalamander.exe' -ExpectedExePath $expectedPath
        }
        catch {
            $errorMessage = $_.Exception.Message
        }

        $errorMessage | Should Match 'PID=4201'
        $errorMessage | Should Match "CommandLine='<unavailable>'"
        Assert-MockCalled Stop-Process -Times 0
    }

    It 'still force-closes an exact-path interactive instance and ignores other paths' {
        $expectedPath = Join-Path $TestDrive 'Debug\RedSalamander.exe'
        $otherPath = Join-Path $TestDrive 'Other\RedSalamander.exe'
        $script:buildOutputProcesses = @(
            [pscustomobject]@{
                ProcessId = 4301
                ExecutablePath = $otherPath
                CommandLine = "`"$otherPath`" --interactive"
            },
            [pscustomobject]@{
                ProcessId = 4302
                ExecutablePath = $expectedPath
                CommandLine = "`"$expectedPath`""
            }
        )

        Stop-BuildOutputProcess -ProcessName 'RedSalamander.exe' -ExpectedExePath $expectedPath

        Assert-MockCalled Stop-Process -Times 1 -ParameterFilter { $Id -eq 4302 -and $Force }
    }
}

Describe 'Artifact operation lock' {
    BeforeAll {
        $sanitizedEnvironmentScript = Join-Path $repoRoot 'Tools\SanitizedEnvironment.ps1'
        $artifactLockScript = Join-Path $repoRoot 'Tools\ArtifactOperationLock.ps1'
        . $sanitizedEnvironmentScript
        . $artifactLockScript

        if (-not ('RedSalamander.ToolingTests.InheritableEvent' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace RedSalamander.ToolingTests
{
    public sealed class InheritableEvent : IDisposable
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct SecurityAttributes
        {
            public int Length;
            public IntPtr SecurityDescriptor;
            [MarshalAs(UnmanagedType.Bool)]
            public bool InheritHandle;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateEventW(
            ref SecurityAttributes attributes,
            [MarshalAs(UnmanagedType.Bool)] bool manualReset,
            [MarshalAs(UnmanagedType.Bool)] bool initialState,
            string name);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);

        private IntPtr handle;

        public InheritableEvent()
        {
            SecurityAttributes attributes = new SecurityAttributes();
            attributes.Length = Marshal.SizeOf(typeof(SecurityAttributes));
            attributes.InheritHandle = true;
            handle = CreateEventW(ref attributes, true, false, null);
            if (handle == IntPtr.Zero || handle == new IntPtr(-1))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "CreateEventW failed.");
            }
        }

        public long HandleValue { get { return handle.ToInt64(); } }

        public bool IsSignaled
        {
            get
            {
                uint result = WaitForSingleObject(handle, 0);
                if (result == 0u) return true;
                if (result == 258u) return false;
                throw new Win32Exception(Marshal.GetLastWin32Error(), "WaitForSingleObject failed.");
            }
        }

        public void Dispose()
        {
            if (handle != IntPtr.Zero)
            {
                CloseHandle(handle);
                handle = IntPtr.Zero;
            }
            GC.SuppressFinalize(this);
        }
    }
}
'@
        }
    }

    It 'uses a canonical repository identity and supports balanced same-thread reentrancy' {
        $testRepo = Join-Path $TestDrive 'Repo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        $lockPath = Get-RSArtifactOperationLockPath -RepoRoot $testRepo
        $caseVariantLockPath = Get-RSArtifactOperationLockPath -RepoRoot $testRepo.ToUpperInvariant()
        $lockPath.ToUpperInvariant() | Should Be $caseVariantLockPath.ToUpperInvariant()

        $outer = $null
        $inner = $null
        $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $testRepo
        try {
            $outer = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'outer-test'
            (Test-Path -LiteralPath $metadataPath) | Should Be $true
            $metadata = Read-RSArtifactOperationOwnerMetadata -RepoRoot $testRepo
            $metadata.owner_pid | Should Be $PID
            $metadata.operation | Should Be 'outer-test'
            [string]::IsNullOrWhiteSpace([string]$metadata.started_utc) | Should Be $false
            (Get-Content -LiteralPath $metadataPath -Raw) |
                Should Match '"started_utc"\s*:\s*"[^"]+Z"'
            (Remove-RSArtifactOperationOwnerMetadata -RepoRoot $testRepo -Token 'not-the-owner') |
                Should Be $false
            (Test-Path -LiteralPath $metadataPath) | Should Be $true

            $inner = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'inner-test'
            $outer.WasAbandoned | Should Be $false
            $inner.WasAbandoned | Should Be $false
            $outer.IsRootOwner | Should Be $true
            $inner.IsRootOwner | Should Be $false

            Exit-RSArtifactOperationLock -Lock $inner
            $inner = $null
            (Test-Path -LiteralPath $metadataPath) | Should Be $true
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $inner
            Exit-RSArtifactOperationLock -Lock $outer
        }
        (Test-Path -LiteralPath $metadataPath) | Should Be $false

        $afterRelease = $null
        try {
            $afterRelease = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'after-release-test'
            $afterRelease.WasAbandoned | Should Be $false
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $afterRelease
        }
    }

    It 'creates the process inside the exact containment job and preserves Unicode arguments and redirected streams' {
        $childScript = Join-Path $TestDrive 'ContainedUnicodeChild.ps1'
        $workingDirectory = Join-Path $TestDrive 'work café 漢'
        New-Item -ItemType Directory -Path $workingDirectory -Force | Out-Null
        @'
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Values)
$utf8 = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = $utf8
[Console]::Error.WriteLine('stderr-café-漢')
[pscustomobject]@{
    values = @($Values)
    workingDirectory = (Get-Location).Path
} | ConvertTo-Json -Compress
exit 37
'@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $values = @('', 'two words', 'quote"value', 'trailing\', 'café', '漢')
        $pwsh = (Get-Process -Id $PID).Path
        $startInfo = New-RSProcessStartInfo `
            -FilePath $pwsh `
            -Arguments (@('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $childScript) + $values) `
            -WorkingDirectory $workingDirectory
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $startInfo.CreateNoWindow = $true
        $utf8 = [System.Text.UTF8Encoding]::new($false)
        $startInfo.StandardOutputEncoding = $utf8
        $startInfo.StandardErrorEncoding = $utf8

        $process = $null
        try {
            $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
            $process.IsInContainmentJob | Should Be $true
            $stdoutTask = $process.StandardOutput.ReadToEndAsync()
            $stderrTask = $process.StandardError.ReadToEndAsync()
            $process.WaitForExit(10000) | Should Be $true
            $process.ExitCode | Should Be 37
            $payload = $stdoutTask.GetAwaiter().GetResult().Trim() | ConvertFrom-Json
            $stderrTask.GetAwaiter().GetResult().Trim() | Should Be 'stderr-café-漢'
            (@($payload.values) | ConvertTo-Json -Compress) | Should Be ($values | ConvertTo-Json -Compress)
            [System.IO.Path]::GetFullPath([string]$payload.workingDirectory) | Should Be ([System.IO.Path]::GetFullPath($workingDirectory))
        }
        finally {
            Close-RSContainedProcess -Process $process
        }
    }

    It 'allows only the explicit standard handles to cross process creation' {
        $childScript = Join-Path $TestDrive 'ContainedHandleProbe.ps1'
        @'
param([long]$HandleValue)
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class ChildHandleProbe
{
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetEvent(IntPtr handle);
    public static void Probe(long value)
    {
        SetEvent(new IntPtr(value));
    }
}
"@
[ChildHandleProbe]::Probe($HandleValue)
'probe-complete'
'@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $sentinel = [RedSalamander.ToolingTests.InheritableEvent]::new()
        $process = $null
        try {
            $pwsh = (Get-Process -Id $PID).Path
            $startInfo = New-RSProcessStartInfo `
                -FilePath $pwsh `
                -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $childScript, '-HandleValue', [string]$sentinel.HandleValue) `
                -WorkingDirectory $TestDrive
            $startInfo.RedirectStandardOutput = $true
            $startInfo.RedirectStandardError = $true
            $startInfo.CreateNoWindow = $true
            $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
            $stdoutTask = $process.StandardOutput.ReadToEndAsync()
            $stderrTask = $process.StandardError.ReadToEndAsync()
            $process.WaitForExit(10000) | Should Be $true
            $process.ExitCode | Should Be 0
            $stderrTask.GetAwaiter().GetResult() | Should Be ''
            $stdoutTask.GetAwaiter().GetResult().Trim() | Should Be 'probe-complete'
            $sentinel.IsSignaled | Should Be $false
        }
        finally {
            Close-RSContainedProcess -Process $process
            $sentinel.Dispose()
        }
    }

    It 'kills an immediately spawned grandchild when the containment job closes' {
        $childScript = Join-Path $TestDrive 'ContainedTreeChild.ps1'
        $grandchildScript = Join-Path $TestDrive 'ContainedTreeGrandchild.ps1'
        $pidPath = Join-Path $TestDrive 'contained-grandchild.pid'
        'Start-Sleep -Seconds 60' | Set-Content -LiteralPath $grandchildScript -Encoding UTF8
        @'
param([string]$PowerShellPath, [string]$GrandchildScript, [string]$PidPath)
$grandchild = Start-Process -FilePath $PowerShellPath -ArgumentList @('-NoProfile', '-File', ('"' + $GrandchildScript + '"')) -PassThru -WindowStyle Hidden
[System.IO.File]::WriteAllText($PidPath, [string]$grandchild.Id)
while ($true) { Start-Sleep -Milliseconds 100 }
'@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $pwsh = (Get-Process -Id $PID).Path
        $startInfo = New-RSProcessStartInfo `
            -FilePath $pwsh `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $childScript, '-PowerShellPath', $pwsh, '-GrandchildScript', $grandchildScript, '-PidPath', $pidPath) `
            -WorkingDirectory $TestDrive
        $startInfo.CreateNoWindow = $true

        $process = $null
        $rootPid = 0
        $grandchildPid = 0
        try {
            $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
            $rootPid = $process.Id
            $process.IsInContainmentJob | Should Be $true
            for ($attempt = 0; $attempt -lt 200 -and -not (Test-Path -LiteralPath $pidPath); $attempt++) {
                Start-Sleep -Milliseconds 25
            }
            (Test-Path -LiteralPath $pidPath) | Should Be $true
            $grandchildPid = [int](Get-Content -LiteralPath $pidPath -Raw)
            $null -ne (Get-Process -Id $grandchildPid -ErrorAction SilentlyContinue) | Should Be $true
            Close-RSContainedProcess -Process $process
            $process = $null

            for ($attempt = 0; $attempt -lt 200; $attempt++) {
                if ($null -eq (Get-Process -Id $rootPid -ErrorAction SilentlyContinue) -and
                    $null -eq (Get-Process -Id $grandchildPid -ErrorAction SilentlyContinue)) {
                    break
                }
                Start-Sleep -Milliseconds 25
            }
            $null -eq (Get-Process -Id $rootPid -ErrorAction SilentlyContinue) | Should Be $true
            $null -eq (Get-Process -Id $grandchildPid -ErrorAction SilentlyContinue) | Should Be $true
        }
        finally {
            Close-RSContainedProcess -Process $process
            if ($grandchildPid -gt 0) {
                Stop-Process -Id $grandchildPid -Force -ErrorAction SilentlyContinue
            }
            if ($rootPid -gt 0) {
                Stop-Process -Id $rootPid -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'cleans native launch resources after repeated creation failures' {
        $beforeHandles = (Get-Process -Id $PID).HandleCount
        for ($attempt = 0; $attempt -lt 10; $attempt++) {
            $startInfo = New-RSProcessStartInfo `
                -FilePath (Join-Path $TestDrive 'missing-contained-process.exe') `
                -WorkingDirectory $TestDrive
            $launchMessage = $null
            try {
                $unexpected = Start-RSContainedProcess -ProcessStartInfo $startInfo
                Close-RSContainedProcess -Process $unexpected
            }
            catch {
                $launchMessage = $_.Exception.Message
            }
            $launchMessage | Should Match 'CreateProcessW failed'
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        $afterHandles = (Get-Process -Id $PID).HandleCount
        ($afterHandles - $beforeHandles) | Should BeLessThan 24
    }

    It 'rejects a second process while the repository artifact lock is owned' {
        $testRepo = Join-Path $TestDrive 'ContendedRepo'
        $readyPath = Join-Path $TestDrive 'owner.ready'
        $releasePath = Join-Path $TestDrive 'owner.release'
        $childScript = Join-Path $TestDrive 'HoldArtifactLock.ps1'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        $escapedHelper = $artifactLockScript.Replace("'", "''")
        $escapedRepo = $testRepo.Replace("'", "''")
        $escapedReady = $readyPath.Replace("'", "''")
        $escapedRelease = $releasePath.Replace("'", "''")
        @"
`$ErrorActionPreference = 'Stop'
. '$escapedHelper'
`$lock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'child-owner'
try {
    [System.IO.File]::WriteAllText('$escapedReady', 'ready')
    while (-not (Test-Path -LiteralPath '$escapedRelease')) {
        Start-Sleep -Milliseconds 25
    }
}
finally {
    Exit-RSArtifactOperationLock -Lock `$lock
}
"@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $pwsh = (Get-Process -Id $PID).Path
        $child = Start-Process -FilePath $pwsh `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$childScript`"") `
            -PassThru `
            -WindowStyle Hidden
        try {
            $ready = $false
            for ($attempt = 0; $attempt -lt 200; $attempt++) {
                if (Test-Path -LiteralPath $readyPath) {
                    $ready = $true
                    break
                }
                Start-Sleep -Milliseconds 25
            }
            $ready | Should Be $true

            $message = $null
            try {
                $unexpected = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'second-owner'
                Exit-RSArtifactOperationLock -Lock $unexpected
            }
            catch {
                $message = $_.Exception.Message
            }
            $message | Should Match 'another RedSalamander build or test operation owns'
            $message | Should Match "Owner PID=$($child.Id)"
            $message | Should Match "Operation='child-owner'"
            $message | Should Match "StartedUtc='[^']+'"
        }
        finally {
            New-Item -ItemType File -Path $releasePath -Force | Out-Null
            if (-not $child.WaitForExit(5000)) {
                Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
            }
            $child.Dispose()
        }
    }

    It 'rejects a parallel runspace in the owner process instead of treating it as local nesting' {
        $testRepo = Join-Path $TestDrive 'ParallelRunspaceRepo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null
        $escapedHelper = $artifactLockScript.Replace("'", "''")
        $escapedRepo = $testRepo.Replace("'", "''")

        $outer = $null
        $parallelPowerShell = $null
        try {
            $outer = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'owning-runspace'
            $parallelPowerShell = [PowerShell]::Create()
            [void]$parallelPowerShell.AddScript(@"
`$ErrorActionPreference = 'Stop'
. '$escapedHelper'
try {
    `$unexpected = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'parallel-runspace'
    Exit-RSArtifactOperationLock -Lock `$unexpected
    Write-Output 'ACQUIRED'
}
catch {
    Write-Output ('ERROR: ' + `$_.Exception.Message)
}
"@)

            $parallelResult = @($parallelPowerShell.Invoke()) -join "`n"
            $parallelResult | Should Match 'another RedSalamander build or test operation owns'
            $parallelResult | Should Not Match 'ACQUIRED'
        }
        finally {
            if ($null -ne $parallelPowerShell) {
                $parallelPowerShell.Dispose()
            }
            Exit-RSArtifactOperationLock -Lock $outer
        }
    }

    It 'delegates only to a direct child after kill-on-close containment is active' {
        $testRepo = Join-Path $TestDrive 'InheritedRepo'
        $childScript = Join-Path $TestDrive 'InheritArtifactLock.ps1'
        $resultPath = Join-Path $TestDrive 'inherited-result.txt'
        $releasePath = Join-Path $TestDrive 'inherited-release.txt'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        $escapedEnvironment = $sanitizedEnvironmentScript.Replace("'", "''")
        $escapedHelper = $artifactLockScript.Replace("'", "''")
        $escapedRepo = $testRepo.Replace("'", "''")
        $escapedResult = $resultPath.Replace("'", "''")
        $escapedRelease = $releasePath.Replace("'", "''")
        @"
`$ErrorActionPreference = 'Stop'
. '$escapedEnvironment'
. '$escapedHelper'
`$lock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'descendant-build'
try {
    [System.IO.File]::WriteAllText('$escapedResult', [string]`$lock.IsDelegated)
    while (-not (Test-Path -LiteralPath '$escapedRelease')) {
        Start-Sleep -Milliseconds 25
    }
}
finally {
    Exit-RSArtifactOperationLock -Lock `$lock
}
"@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $testRepo
        $rootLock = $null
        try {
            $rootLock = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'root-runner'
            $rootMetadata = Read-RSArtifactOperationOwnerMetadata -RepoRoot $testRepo

            $pwsh = (Get-Process -Id $PID).Path
            $startInfo = New-RSProcessStartInfo `
                -FilePath $pwsh `
                -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $childScript) `
                -WorkingDirectory $testRepo
            $startInfo.CreateNoWindow = $true
            $child = $null
            try {
                $child = Start-RSContainedProcess `
                    -ProcessStartInfo $startInfo `
                    -DelegateArtifactOperation

                $resultReady = $false
                for ($attempt = 0; $attempt -lt 200; $attempt++) {
                    if (Test-Path -LiteralPath $resultPath) {
                        $resultReady = $true
                        break
                    }
                    Start-Sleep -Milliseconds 25
                }
                $resultReady | Should Be $true
                $releaseMessage = $null
                try {
                    Exit-RSArtifactOperationLock -Lock $rootLock
                }
                catch {
                    $releaseMessage = $_.Exception.Message
                }
                $releaseMessage | Should Match 'contained child delegation.*remain active'
                (Test-Path -LiteralPath $metadataPath) | Should Be $true

                New-Item -ItemType File -Path $releasePath -Force | Out-Null
                $child.WaitForExit(5000) | Should Be $true
                $child.ExitCode | Should Be 0
            }
            finally {
                New-Item -ItemType File -Path $releasePath -Force | Out-Null
                Close-RSContainedProcess -Process $child
            }

            (Get-Content -LiteralPath $resultPath -Raw) | Should Be 'True'
            (Test-Path -LiteralPath $metadataPath) | Should Be $true
            $metadataAfterChild = Read-RSArtifactOperationOwnerMetadata -RepoRoot $testRepo
            $metadataAfterChild.token | Should Be $rootMetadata.token
            $metadataAfterChild.operation | Should Be 'root-runner'
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $rootLock
        }

        (Test-Path -LiteralPath $metadataPath) | Should Be $false
    }

    It 'makes a surviving stale descendant acquire abandoned ownership instead of bypassing it' {
        $testRepo = Join-Path $TestDrive 'StaleDescendantRepo'
        $ownerScript = Join-Path $TestDrive 'LaunchStaleDescendant.ps1'
        $survivorScript = Join-Path $TestDrive 'StaleDescendant.ps1'
        $survivorReadyPath = Join-Path $TestDrive 'stale-descendant.ready'
        $survivorResultPath = Join-Path $TestDrive 'stale-descendant.result'
        $survivorReleasePath = Join-Path $TestDrive 'stale-descendant.release'
        $survivorPidPath = Join-Path $TestDrive 'stale-descendant.pid'
        $ownerStdoutPath = Join-Path $TestDrive 'stale-owner.stdout.log'
        $ownerStderrPath = Join-Path $TestDrive 'stale-owner.stderr.log'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        $escapedHelper = $artifactLockScript.Replace("'", "''")
        $escapedRepo = $testRepo.Replace("'", "''")
        $escapedSurvivor = $survivorScript.Replace("'", "''")
        $escapedReady = $survivorReadyPath.Replace("'", "''")
        $escapedResult = $survivorResultPath.Replace("'", "''")
        $escapedRelease = $survivorReleasePath.Replace("'", "''")
        $escapedPid = $survivorPidPath.Replace("'", "''")
        @"
param([int]`$OwnerPid)
`$ErrorActionPreference = 'Stop'
. '$escapedHelper'
[System.IO.File]::WriteAllText('$escapedReady', 'ready')
while (`$null -ne (Get-Process -Id `$OwnerPid -ErrorAction SilentlyContinue)) {
    Start-Sleep -Milliseconds 25
}
`$lock = `$null
try {
    `$lock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'surviving-descendant' -TimeoutMilliseconds 5000
    [System.IO.File]::WriteAllText(
        '$escapedResult',
        "`$(`$lock.IsDelegated)|`$(`$lock.WasAbandoned)")
    while (-not (Test-Path -LiteralPath '$escapedRelease')) {
        Start-Sleep -Milliseconds 25
    }
}
catch {
    [System.IO.File]::WriteAllText('$escapedResult', 'ERROR: ' + `$_.Exception.Message)
}
finally {
    Exit-RSArtifactOperationLock -Lock `$lock
}
"@ | Set-Content -LiteralPath $survivorScript -Encoding UTF8

        @"
`$ErrorActionPreference = 'Stop'
. '$escapedHelper'
`$lock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'doomed-owner'
`$pwsh = (Get-Process -Id `$PID).Path
`$survivor = Start-Process -FilePath `$pwsh -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', '"$escapedSurvivor"', '-OwnerPid', `$PID) -PassThru -WindowStyle Hidden
[System.IO.File]::WriteAllText('$escapedPid', [string]`$survivor.Id)
while (-not (Test-Path -LiteralPath '$escapedReady')) {
    Start-Sleep -Milliseconds 25
}
[Environment]::Exit(0)
"@ | Set-Content -LiteralPath $ownerScript -Encoding UTF8

        $pwsh = (Get-Process -Id $PID).Path
        $owner = Start-Process -FilePath $pwsh `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$ownerScript`"") `
            -PassThru `
            -RedirectStandardOutput $ownerStdoutPath `
            -RedirectStandardError $ownerStderrPath `
            -WindowStyle Hidden
        $survivorProcess = $null
        try {
            $owner.WaitForExit(5000) | Should Be $true
            $owner.WaitForExit()
            $owner.Refresh()
            if ($null -ne $owner.ExitCode -and $owner.ExitCode -ne 0) {
                $ownerStdout = if (Test-Path -LiteralPath $ownerStdoutPath) {
                    Get-Content -LiteralPath $ownerStdoutPath -Raw
                } else { '' }
                $ownerStderr = if (Test-Path -LiteralPath $ownerStderrPath) {
                    Get-Content -LiteralPath $ownerStderrPath -Raw
                } else { '' }
                throw "Stale owner fixture exited $($owner.ExitCode). stdout='$ownerStdout' stderr='$ownerStderr'"
            }

            $resultReady = $false
            for ($attempt = 0; $attempt -lt 200; $attempt++) {
                if (Test-Path -LiteralPath $survivorResultPath) {
                    $resultReady = $true
                    break
                }
                Start-Sleep -Milliseconds 25
            }
            $resultReady | Should Be $true
            (Get-Content -LiteralPath $survivorResultPath -Raw) | Should Be 'False|True'

            $survivorPid = [int](Get-Content -LiteralPath $survivorPidPath -Raw)
            $survivorProcess = Get-Process -Id $survivorPid -ErrorAction Stop
            $message = $null
            try {
                $unexpected = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'observer'
                Exit-RSArtifactOperationLock -Lock $unexpected
            }
            catch {
                $message = $_.Exception.Message
            }
            $message | Should Match 'another RedSalamander build or test operation owns'
            $message | Should Match "Owner PID=$survivorPid"
        }
        finally {
            New-Item -ItemType File -Path $survivorReleasePath -Force | Out-Null
            if ($null -ne $survivorProcess) {
                if (-not $survivorProcess.WaitForExit(5000)) {
                    Stop-Process -Id $survivorProcess.Id -Force -ErrorAction SilentlyContinue
                }
                $survivorProcess.Dispose()
            }
            if (-not $owner.HasExited) {
                Stop-Process -Id $owner.Id -Force -ErrorAction SilentlyContinue
            }
            $owner.Dispose()
        }
    }

    It 'reports abandoned ownership and persists contamination until explicitly cleared' {
        $testRepo = Join-Path $TestDrive 'AbandonedRepo'
        $childScript = Join-Path $TestDrive 'AbandonArtifactLock.ps1'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null

        $escapedHelper = $artifactLockScript.Replace("'", "''")
        $escapedRepo = $testRepo.Replace("'", "''")
        @"
`$ErrorActionPreference = 'Stop'
. '$escapedHelper'
`$lock = Enter-RSArtifactOperationLock -RepoRoot '$escapedRepo' -Operation 'abandoned-owner'
[Environment]::Exit(0)
"@ | Set-Content -LiteralPath $childScript -Encoding UTF8

        $pwsh = (Get-Process -Id $PID).Path
        $child = Start-Process -FilePath $pwsh `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$childScript`"") `
            -PassThru `
            -WindowStyle Hidden
        $lock = $null
        try {
            $child.WaitForExit(5000) | Should Be $true
            $child.ExitCode | Should Be 0

            $lock = Enter-RSArtifactOperationLock -RepoRoot $testRepo -Operation 'abandonment-observer'
            $lock.WasAbandoned | Should Be $true
            $markerPath = Set-RSArtifactOperationContaminated -RepoRoot $testRepo -Reason 'test abandonment'
            (Test-Path -LiteralPath $markerPath) | Should Be $true
            (Test-RSArtifactOperationContaminated -RepoRoot $testRepo) | Should Be $true
            Clear-RSArtifactOperationContaminated -RepoRoot $testRepo
            (Test-RSArtifactOperationContaminated -RepoRoot $testRepo) | Should Be $false
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $lock
            if (-not $child.HasExited) {
                Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
            }
            $child.Dispose()
        }

        $legacyRepo = Join-Path $TestDrive 'LegacyOwnerRepo'
        $legacyMetadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $legacyRepo
        New-Item -ItemType Directory -Path (Split-Path -Parent $legacyMetadataPath) -Force | Out-Null
        [System.IO.File]::WriteAllText(
            $legacyMetadataPath,
            '{"schema":"red-salamander.artifact-operation-owner.v1","owner_pid":82648}',
            [System.Text.UTF8Encoding]::new($false))

        $legacyLock = $null
        try {
            $legacyLock = Enter-RSArtifactOperationLock -RepoRoot $legacyRepo -Operation 'legacy-sidecar-observer'
            $legacyLock.WasAbandoned | Should Be $true
            (Read-RSArtifactOperationOwnerMetadata -RepoRoot $legacyRepo).schema |
                Should Be 'red-salamander.artifact-operation-owner.v3'
            [void](Set-RSArtifactOperationContaminated -RepoRoot $legacyRepo -Reason 'stale v1 owner')
            (Test-RSArtifactOperationContaminated -RepoRoot $legacyRepo) | Should Be $true
            Clear-RSArtifactOperationContaminated -RepoRoot $legacyRepo
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $legacyLock
        }
    }

    It 'rejects residual compiler tools whose command line targets the repository' {
        $testRepo = Join-Path $TestDrive 'ResidualRepo'
        New-Item -ItemType Directory -Path $testRepo -Force | Out-Null
        Mock Get-CimInstance {
            return @(
                [pscustomobject]@{
                    ProcessId = 7711
                    ParentProcessId = 7700
                    Name = 'cl.exe'
                    CommandLine = "cl.exe /Fo`"$testRepo\.build\Intermediate\sample.obj`" sample.cpp"
                },
                [pscustomobject]@{
                    ProcessId = 8811
                    ParentProcessId = 8800
                    Name = 'MSBuild.exe'
                    CommandLine = 'MSBuild.exe C:\OtherRepo\Other.sln'
                }
            )
        }

        $message = $null
        try {
            Assert-RSNoResidualArtifactToolProcesses -RepoRoot $testRepo
        }
        catch {
            $message = $_.Exception.Message
        }

        $message | Should Match 'build tool is still touching this repository'
        $message | Should Match 'PID=7711'
        $message | Should Not Match 'PID=8811'
    }
}
