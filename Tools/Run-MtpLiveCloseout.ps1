<#
.SYNOPSIS
    Runs and archives the opt-in MTP/PTP live-device closeout smoke.

.DESCRIPTION
    This helper wraps the gated `mtp_live_device_smoke` CompareDirectories case.
    Live mode requires both an approved device selector and an explicit scratch
    folder so the destructive write/read/overwrite/rename/copy/move/delete
    matrix never targets arbitrary user data. Probe mode performs the same
    production WPD enumeration path with an impossible requested device name and
    is useful for archiving "no device visible" evidence without writes.

.PARAMETER Device
    Selects the approved live MTP/PTP device name or pattern for the destructive smoke.

.PARAMETER Scratch
    Names the dedicated disposable device folder in which the live mutation matrix may run.

.PARAMETER ProbeNoDevice
    Runs the production enumeration path with an impossible selector and performs no device writes.

.PARAMETER SkipBuild
    Reuses an existing test-enabled build instead of building before the focused case.

.PARAMETER TimeoutMultiplier
    Scales the focused CompareDirectories case timeout for slow devices or hosts.

.PARAMETER Platform
    Selects the x64 or ARM64 build lane.

.PARAMETER Configuration
    Selects the test-enabled build configuration.

.PARAMETER ArchiveRoot
    Selects the evidence directory below the current X:\RedSalamander.Perf run.
    Relative values are resolved below that run. The default is
    runs\<runId>\artifacts\mtp-live-closeout.

.PARAMETER TestRoot
    Exact X:\RedSalamander.Perf root. The repository drive is used by default.

.PARAMETER AllowExternalTestRoot
    Allows first-time initialization of the exact selected X:\RedSalamander.Perf root.

.OUTPUTS
    No supported pipeline objects. Writes command, environment, device-probe, test logs, and copied self-test artifacts below ArchiveRoot and exits with the focused runner's exit code.

.NOTES
    Prerequisites: Windows WPD support, an approved live device and disposable scratch folder for Live mode, an initialized X:\RedSalamander.Perf root, and the normal test toolchain unless -SkipBuild is used. Side effects: Live mode creates, reads, overwrites, renames, copies, moves, and deletes only inside the supplied device scratch folder; both modes write local evidence only below the selected test root. Primary use is explicit operator closeout, never default CI.

.EXAMPLE
    .\Tools\Run-MtpLiveCloseout.ps1 -ProbeNoDevice -SkipBuild

.EXAMPLE
    .\Tools\Run-MtpLiveCloseout.ps1 -Device "*" -Scratch "RedSalamanderMtpSmoke" -SkipBuild
#>

[CmdletBinding(DefaultParameterSetName = 'Live')]
param(
    [Parameter(ParameterSetName = 'Live', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Device,

    [Parameter(ParameterSetName = 'Live', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Scratch,

    [Parameter(ParameterSetName = 'Probe', Mandatory = $true)]
    [switch]$ProbeNoDevice,

    [switch]$SkipBuild,

    [double]$TimeoutMultiplier = 3.0,

    [ValidateSet('x64', 'ARM64')]
    [string]$Platform = 'x64',

    [string]$Configuration = 'Debug',

    [string]$ArchiveRoot = '',

    [string]$TestRoot = $env:REDSALAMANDER_TEST_ROOT,

    [switch]$AllowExternalTestRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $repoRoot) {
    throw 'Unable to resolve repository root.'
}

$runAll = Join-Path $repoRoot 'Tools\Run-AllTests.ps1'
if (-not (Test-Path $runAll)) {
    throw "Run-AllTests.ps1 not found at $runAll"
}

$testRunPlanModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
Import-Module $testRunPlanModule -Force -ErrorAction Stop
$selectedTestRoot = Get-RSTestSandboxRoot `
    -RepoRoot $repoRoot `
    -TestRootOverride $TestRoot `
    -SelfTestRootOverride '' `
    -AllowExternalInitialization:$AllowExternalTestRoot
[void](Initialize-RSTestSandboxRoot `
        -TestRoot $selectedTestRoot `
        -RepoRoot $repoRoot `
        -AllowExternalInitialization:$AllowExternalTestRoot)
$context = New-RSTestRunContext `
    -RepoRoot $repoRoot `
    -RunId (New-RSTestRunId) `
    -TestRootOverride $selectedTestRoot `
    -SelfTestRootOverride ''

if ([string]::IsNullOrWhiteSpace($ArchiveRoot)) {
    $ArchiveRoot = Join-Path $context.RunRoot 'artifacts\mtp-live-closeout'
} elseif (-not [System.IO.Path]::IsPathRooted($ArchiveRoot)) {
    $ArchiveRoot = Join-Path $context.RunRoot $ArchiveRoot
}
$ArchiveRoot = Assert-RSTestSandboxContainedPath `
    -TestRoot $context.TestRoot `
    -Path $ArchiveRoot `
    -RequireDescendant
$selfTestArtifactRoot = [IO.Path]::GetFullPath($context.ArtifactRoot).TrimEnd('\')
if ($ArchiveRoot.StartsWith("$selfTestArtifactRoot\", [System.StringComparison]::OrdinalIgnoreCase) -or
    [string]::Equals($ArchiveRoot.TrimEnd('\'), $selfTestArtifactRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "ArchiveRoot must not be the self-test artifact source: $ArchiveRoot"
}

New-Item -ItemType Directory -Force -Path $ArchiveRoot | Out-Null

$deviceValue = if ($ProbeNoDevice) { '__redsal_mtp_probe_no_such_device__' } else { $Device }
$scratchValue = if ($ProbeNoDevice) { '' } else { $Scratch }

$previousEnv = @{
    REDSALAMANDER_TEST_ROOT           = [Environment]::GetEnvironmentVariable('REDSALAMANDER_TEST_ROOT', 'Process')
    REDSALAMANDER_TEST_RUN_ID         = [Environment]::GetEnvironmentVariable('REDSALAMANDER_TEST_RUN_ID', 'Process')
    REDSALAMANDER_SELFTEST_MTP_DEVICE  = [Environment]::GetEnvironmentVariable('REDSALAMANDER_SELFTEST_MTP_DEVICE', 'Process')
    REDSALAMANDER_SELFTEST_MTP_PROFILE = [Environment]::GetEnvironmentVariable('REDSALAMANDER_SELFTEST_MTP_PROFILE', 'Process')
    REDSALAMANDER_SELFTEST_MTP_ROOT    = [Environment]::GetEnvironmentVariable('REDSALAMANDER_SELFTEST_MTP_ROOT', 'Process')
    REDSALAMANDER_SELFTEST_MTP_SCRATCH = [Environment]::GetEnvironmentVariable('REDSALAMANDER_SELFTEST_MTP_SCRATCH', 'Process')
}

function Set-ProcessEnvironmentValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [AllowNull()]
        [string]$Value
    )

    [Environment]::SetEnvironmentVariable($Name, $Value, 'Process')
}

function Write-DeviceProbe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory
    )

    try {
        Get-PnpDevice -PresentOnly |
            Where-Object {
                $_.Class -in @('WPD', 'Image', 'Camera', 'PortableDevice') -or
                $_.FriendlyName -match 'MTP|PTP|Portable|Phone|Camera|Android|iPhone|Pixel|Samsung|Nikon|Canon' -or
                $_.InstanceId -match 'WPD|MTP|USBSTOR|VID_'
            } |
            Select-Object Status, Class, FriendlyName, InstanceId |
            Format-List |
            Out-File -FilePath (Join-Path $OutputDirectory 'pnp-device-probe.txt') -Encoding utf8
    } catch {
        "Get-PnpDevice probe failed: $($_.Exception.Message)" |
            Out-File -FilePath (Join-Path $OutputDirectory 'pnp-device-probe.txt') -Encoding utf8
    }

    try {
        Get-CimInstance Win32_PnPEntity |
            Where-Object {
                $_.PNPClass -eq 'WPD' -or
                $_.Name -match 'MTP|PTP|Portable Device|Phone|Android|iPhone|Pixel|Samsung|Nikon|Canon' -or
                $_.DeviceID -match 'WPD|MTP'
            } |
            Select-Object Name, PNPClass, Status, DeviceID |
            Format-List |
            Out-File -FilePath (Join-Path $OutputDirectory 'cim-wpd-probe.txt') -Encoding utf8
    } catch {
        "CIM WPD probe failed: $($_.Exception.Message)" |
            Out-File -FilePath (Join-Path $OutputDirectory 'cim-wpd-probe.txt') -Encoding utf8
    }
}

$exitCode = 1
try {
    Set-ProcessEnvironmentValue -Name 'REDSALAMANDER_TEST_ROOT' -Value $context.TestRoot
    Set-ProcessEnvironmentValue -Name 'REDSALAMANDER_TEST_RUN_ID' -Value $context.RunId
    Set-ProcessEnvironmentValue -Name 'REDSALAMANDER_SELFTEST_MTP_DEVICE' -Value $deviceValue
    Set-ProcessEnvironmentValue -Name 'REDSALAMANDER_SELFTEST_MTP_SCRATCH' -Value $scratchValue

    $envSnapshot = @(
        "REDSALAMANDER_SELFTEST_MTP_DEVICE=$env:REDSALAMANDER_SELFTEST_MTP_DEVICE",
        "REDSALAMANDER_SELFTEST_MTP_PROFILE=$env:REDSALAMANDER_SELFTEST_MTP_PROFILE",
        "REDSALAMANDER_SELFTEST_MTP_ROOT=$env:REDSALAMANDER_SELFTEST_MTP_ROOT",
        "REDSALAMANDER_SELFTEST_MTP_SCRATCH=$env:REDSALAMANDER_SELFTEST_MTP_SCRATCH"
    )
    $envSnapshot | Out-File -FilePath (Join-Path $ArchiveRoot 'mtp-env.txt') -Encoding utf8

    Write-DeviceProbe -OutputDirectory $ArchiveRoot

    $stdout = Join-Path $ArchiveRoot 'run-all-tests.stdout.log'
    $stderr = Join-Path $ArchiveRoot 'run-all-tests.stderr.log'
    $pwsh = (Get-Process -Id $PID).Path
    if ([string]::IsNullOrWhiteSpace($pwsh)) {
        throw 'Unable to locate current PowerShell executable.'
    }

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $runAll,
        '-Suite', 'Compare',
        '-ValidationMode', 'Fresh',
        '-CaseFilter', 'mtp_live_device_smoke',
        '-FailFast',
        '-TimeoutMultiplier', ([string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0}', $TimeoutMultiplier)),
        '-Platform', $Platform,
        '-Configuration', $Configuration,
        '-TestRoot', $context.TestRoot,
        '-RunId', $context.RunId
    )
    if ($AllowExternalTestRoot) {
        $arguments += '-AllowExternalTestRoot'
    }
    if ($SkipBuild) {
        $arguments += '-SkipBuild'
    }

    $commandLine = "$pwsh " + ($arguments -join ' ')
    $commandLine | Out-File -FilePath (Join-Path $ArchiveRoot 'command.txt') -Encoding utf8

    $process = Start-Process -FilePath $pwsh `
        -ArgumentList $arguments `
        -WorkingDirectory $repoRoot `
        -PassThru `
        -Wait `
        -NoNewWindow `
        -RedirectStandardOutput $stdout `
        -RedirectStandardError $stderr

    $exitCode = if ($null -ne $process.ExitCode) { [int]$process.ExitCode } else { 0 }

    $lastRun = $context.ArtifactRoot
    if (Test-Path $lastRun) {
        $lastRunArchive = Join-Path $ArchiveRoot 'last_run'
        New-Item -ItemType Directory -Force -Path $lastRunArchive | Out-Null
        Copy-Item -Path (Join-Path $lastRun '*') -Destination $lastRunArchive -Recurse -Force
    }

    $summaryLines = @(
        '# MTP Live Closeout Run',
        '',
        "- Date: $(Get-Date -Format o)",
        "- Mode: $(if ($ProbeNoDevice) { 'ProbeNoDevice' } else { 'Live' })",
        ('- Device selector: `{0}`' -f $deviceValue),
        ('- Scratch: `{0}`' -f $scratchValue),
        "- Exit code: $exitCode",
        ('- Archive: `{0}`' -f $ArchiveRoot),
        ''
    )

    $resultsJson = Join-Path $ArchiveRoot 'last_run\run-all-tests-results.json'
    if (Test-Path $resultsJson) {
        try {
            $summary = Get-Content $resultsJson | ConvertFrom-Json
            if ($summary.schema -eq 'red-salamander.run-all-tests.v2') {
                $summaryLines += "- Result: repository=$($summary.repository_verdict) change=$($summary.change_verdict) executed=$($summary.counts.executed) reused=$($summary.counts.reused) total=$($summary.case_counts.total) passed=$($summary.case_counts.passed) failed=$($summary.case_counts.failed) skipped=$($summary.case_counts.skipped)"
            } else {
                $summaryLines += "- Result: total=$($summary.total) passed=$($summary.passed) failed=$($summary.failed) skipped=$($summary.skipped)"
            }
        } catch {
            $summaryLines += "- Result: unable to parse run-all-tests-results.json ($($_.Exception.Message))"
        }
    } else {
        $summaryLines += '- Result: run-all-tests-results.json was not archived.'
    }

    $summaryLines | Out-File -FilePath (Join-Path $ArchiveRoot 'summary.md') -Encoding utf8

    Write-Host "MTP live closeout archive: $ArchiveRoot"
    Write-Host "Exit code: $exitCode"
} finally {
    foreach ($entry in $previousEnv.GetEnumerator()) {
        Set-ProcessEnvironmentValue -Name $entry.Key -Value $entry.Value
    }
}

exit $exitCode
