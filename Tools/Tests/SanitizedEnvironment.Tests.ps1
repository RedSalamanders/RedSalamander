Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperModule = Join-Path $repoRoot 'Tools\Modules\Build\SanitizedEnvironment.psm1'

Describe 'SanitizedEnvironment helper' {
    BeforeAll {
        Import-Module $helperModule -Force -ErrorAction Stop
    }

    It 'creates a child ProcessStartInfo without mutating the caller Path' {
        $originalPath = [System.Environment]::GetEnvironmentVariable('Path', 'Process')
        try {
            [System.Environment]::SetEnvironmentVariable('Path', 'C:\Temp\One;C:\Temp\Two', 'Process')

            $null = New-RSProcessStartInfo `
                -FilePath 'cmd.exe' `
                -Arguments @('/d', '/c', 'exit', '0')

            [System.Environment]::GetEnvironmentVariable('Path', 'Process') | Should Be 'C:\Temp\One;C:\Temp\Two'
        }
        finally {
            [System.Environment]::SetEnvironmentVariable('Path', $originalPath, 'Process')
        }
    }

    It 'creates a child environment with a canonical Path entry and extra variables' {
        $psi = New-RSProcessStartInfo `
            -FilePath 'cmd.exe' `
            -Arguments @('/d', '/c', 'exit', '0') `
            -AdditionalEnvironment @{ 'RS_TEST_FLAG' = '1' }

        $pathEntry = $psi.Environment.Keys | Where-Object { $_ -ceq 'Path' }
        $duplicatePathAlias = $psi.Environment.Keys | Where-Object { $_ -ceq 'PATH' }

        $pathEntry | Should Be 'Path'
        $duplicatePathAlias | Should BeNullOrEmpty
        $psi.Environment['Path'] | Should Not BeNullOrEmpty
        $psi.Environment['RS_TEST_FLAG'] | Should Be '1'
    }

    It 'preserves the effective process Path order while merging duplicate aliases' {
        $processEnvironment = [System.Collections.Specialized.OrderedDictionary]::new([System.StringComparer]::Ordinal)
        $processEnvironment.Add('PATH', 'C:\UpperOnly;C:\Shared')
        $processEnvironment.Add('Path', 'C:\LowerOnly;C:\Shared')
        $processEnvironment.Add('RS_TEST_FLAG', '1')

        $normalized = Get-RSNormalizedEnvironmentMap `
            -ProcessEnvironment $processEnvironment `
            -ProcessPathValue 'C:\EffectiveFirst;C:\Shared'

        $pathEntry = $normalized.Keys | Where-Object { $_ -ceq 'Path' }
        $duplicatePathAlias = $normalized.Keys | Where-Object { $_ -ceq 'PATH' }

        $pathEntry | Should Be 'Path'
        $duplicatePathAlias | Should BeNullOrEmpty
        $normalized['Path'] | Should Be 'C:\EffectiveFirst;C:\Shared;C:\UpperOnly;C:\LowerOnly'
        $normalized['RS_TEST_FLAG'] | Should Be '1'
    }

    It 'canonicalizes uppercase PATH overrides for child processes' {
        $psi = New-RSProcessStartInfo `
            -FilePath 'cmd.exe' `
            -Arguments @('/d', '/c', 'exit', '0') `
            -AdditionalEnvironment @{ 'PATH' = 'C:\Override' }

        $pathEntry = $psi.Environment.Keys | Where-Object { $_ -ceq 'Path' }
        $duplicatePathAlias = $psi.Environment.Keys | Where-Object { $_ -ceq 'PATH' }

        $pathEntry | Should Be 'Path'
        $duplicatePathAlias | Should BeNullOrEmpty
        $psi.Environment['Path'] | Should Be 'C:\Override'
    }

    It 'removes an absent null override without creating an empty child key' {
        $name = 'RS_NULL_OVERRIDE_' + [Guid]::NewGuid().ToString('N')
        $psi = New-RSProcessStartInfo -FilePath 'cmd.exe' -AdditionalEnvironment @{ $name = $null }
        $psi.Environment.ContainsKey($name) | Should Be $false
        [System.Environment]::GetEnvironmentVariables('Process').Contains($name) | Should Be $false
    }

    It 'removes an inherited child key while preserving its caller value' {
        $name = 'RS_INHERITED_OVERRIDE_' + [Guid]::NewGuid().ToString('N')
        try {
            [System.Environment]::SetEnvironmentVariable($name, 'caller-value', 'Process')
            $psi = New-RSProcessStartInfo -FilePath 'cmd.exe' -AdditionalEnvironment @{ $name = $null }
            $psi.Environment.ContainsKey($name) | Should Be $false
            [System.Environment]::GetEnvironmentVariable($name, 'Process') | Should Be 'caller-value'
        }
        finally {
            [System.Environment]::SetEnvironmentVariable($name, $null, 'Process')
        }
    }

    It 'keeps an explicit empty child value distinct from removing the key' {
        $name = 'RS_EMPTY_OVERRIDE_' + [Guid]::NewGuid().ToString('N')
        $psi = New-RSProcessStartInfo -FilePath 'cmd.exe' -AdditionalEnvironment @{ $name = '' }
        $psi.Environment.ContainsKey($name) | Should Be $true
        $psi.Environment[$name] | Should Be ''
        [System.Environment]::GetEnvironmentVariables('Process').Contains($name) | Should Be $false
    }

    It 'preserves absent empty and nonempty overrides in the actual contained child' {
        $suffix = [Guid]::NewGuid().ToString('N')
        $removedName = 'RS_REMOVED_' + $suffix
        $emptyName = 'RS_EMPTY_' + $suffix
        $presentName = 'RS_PRESENT_' + $suffix
        $scriptPath = Join-Path $TestDrive 'CaptureOverrideStates.ps1'
        $outputPath = Join-Path $TestDrive 'override-states.json'
        @'
param([string]$OutputPath, [string]$RemovedName, [string]$EmptyName, [string]$PresentName)
$environmentMap = [System.Environment]::GetEnvironmentVariables('Process')
[pscustomobject]@{
    removedPresent = $environmentMap.Contains($RemovedName)
    emptyPresent = $environmentMap.Contains($EmptyName)
    emptyValue = $environmentMap[$EmptyName]
    presentPresent = $environmentMap.Contains($PresentName)
    presentValue = $environmentMap[$PresentName]
} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
'@ | Set-Content -LiteralPath $scriptPath -Encoding ASCII
        try {
            [System.Environment]::SetEnvironmentVariable($removedName, 'caller-only', 'Process')
            $exitCode = Invoke-RSProcess `
                -FilePath 'powershell.exe' `
                -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, $outputPath, $removedName, $emptyName, $presentName) `
                -WorkingDirectory $TestDrive `
                -AdditionalEnvironment @{ $removedName = $null; $emptyName = ''; $presentName = 'child-only' }
            $payload = Get-Content -LiteralPath $outputPath -Raw | ConvertFrom-Json
            $exitCode | Should Be 0
            $payload.removedPresent | Should Be $false
            $payload.emptyPresent | Should Be $true
            $payload.emptyValue | Should Be ''
            $payload.presentPresent | Should Be $true
            $payload.presentValue | Should Be 'child-only'
            [System.Environment]::GetEnvironmentVariable($removedName, 'Process') | Should Be 'caller-only'
            [System.Environment]::GetEnvironmentVariables('Process').Contains($emptyName) | Should Be $false
            [System.Environment]::GetEnvironmentVariables('Process').Contains($presentName) | Should Be $false
        }
        finally {
            [System.Environment]::SetEnvironmentVariable($removedName, $null, 'Process')
        }
    }

    It 'launches direct child processes with a sanitized Path key' {
        $scriptPath = Join-Path $TestDrive 'CaptureEnvironment.ps1'
        $outputPath = Join-Path $TestDrive 'environment.json'
        @'
param(
    [string]$OutputPath
)

$keys = @([System.Environment]::GetEnvironmentVariables('Process').Keys | Where-Object { $_ -ieq 'Path' })
$payload = [pscustomobject]@{
    keys = $keys
    path = [System.Environment]::GetEnvironmentVariable('Path', 'Process')
}
$payload | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
'@ | Set-Content -Path $scriptPath -Encoding ASCII

        $exitCode = Invoke-RSProcess `
            -FilePath 'powershell.exe' `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, $outputPath) `
            -WorkingDirectory $TestDrive `
            -AdditionalEnvironment @{ 'PATH' = 'C:\DirectChild' }

        $payload = Get-Content -Path $outputPath -Raw | ConvertFrom-Json

        $exitCode | Should Be 0
        @($payload.keys) | Should Be @('Path')
        $payload.path | Should Be 'C:\DirectChild'
    }

    It 'contains each child in a job that kills on close and ends an unhandled exception without a Windows Error Reporting dialog' {
        # The child reads the limit flags of the job it runs in. DIE_ON_UNHANDLED_EXCEPTION applies SEM_NOGPFAULTERRORBOX to every
        # process of the job, so a crash in an unattended build or test run ends the process instead of waiting on a dialog.
        $scriptPath = Join-Path $TestDrive 'QueryJobLimits.ps1'
        $outputPath = Join-Path $TestDrive 'job-limits.json'
        @'
param(
    [string]$OutputPath
)

Add-Type -TypeDefinition @"
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

public static class RSContainmentProbe
{
    [StructLayout(LayoutKind.Sequential)]
    private struct BasicLimits
    {
        public long PerProcessUserTimeLimit;
        public long PerJobUserTimeLimit;
        public uint LimitFlags;
        public UIntPtr MinimumWorkingSetSize;
        public UIntPtr MaximumWorkingSetSize;
        public uint ActiveProcessLimit;
        public UIntPtr Affinity;
        public uint PriorityClass;
        public uint SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ExtendedLimits
    {
        public BasicLimits Basic;
        public ulong ReadOperationCount;
        public ulong WriteOperationCount;
        public ulong OtherOperationCount;
        public ulong ReadTransferCount;
        public ulong WriteTransferCount;
        public ulong OtherTransferCount;
        public UIntPtr ProcessMemoryLimit;
        public UIntPtr JobMemoryLimit;
        public UIntPtr PeakProcessMemoryUsed;
        public UIntPtr PeakJobMemoryUsed;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool QueryInformationJobObject(IntPtr job, int informationClass, out ExtendedLimits information, int length, IntPtr returned);

    public static uint LimitFlags()
    {
        ExtendedLimits limits;
        // A null job handle queries the job the calling process runs in.
        if (!QueryInformationJobObject(IntPtr.Zero, 9, out limits, Marshal.SizeOf(typeof(ExtendedLimits)), IntPtr.Zero))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error(), "QueryInformationJobObject failed.");
        }
        return limits.Basic.LimitFlags;
    }
}
"@

[pscustomobject]@{ limitFlags = [RSContainmentProbe]::LimitFlags() } | ConvertTo-Json | Set-Content -LiteralPath $OutputPath -Encoding UTF8
'@ | Set-Content -Path $scriptPath -Encoding ASCII

        $exitCode = Invoke-RSProcess `
            -FilePath 'powershell.exe' `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, $outputPath) `
            -WorkingDirectory $TestDrive

        $payload = Get-Content -Path $outputPath -Raw | ConvertFrom-Json
        $limitFlags = [uint32]$payload.limitFlags

        $exitCode | Should Be 0
        ($limitFlags -band 0x00002000) | Should Be 0x00002000 # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        ($limitFlags -band 0x00000400) | Should Be 0x00000400 # JOB_OBJECT_LIMIT_DIE_ON_UNHANDLED_EXCEPTION
    }
}
