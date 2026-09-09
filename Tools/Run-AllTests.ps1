<#
.SYNOPSIS
    Builds and runs all RedSalamander self-test suites, then displays a comprehensive summary.

.DESCRIPTION
    This script:
      1. Optionally builds the Debug configuration of RedSalamander.
      2. Runs each self-test suite (Compare, Commands, FileOps) or a selected subset.
         Suite Full also runs standalone native tests, CppUnitTest DLLs, and
         PowerShell test scripts.
      3. Parses the results.json artifacts from each suite.
      4. Displays a color-coded summary with pass/fail/skip counts, timing, and failure details.
      5. Returns a typed exit code: 0 PASSED, 1 FAILED, 2 NOT_EVALUATED,
         3 invalid CLI/plan, 4 blocked attestation, or 5 corrupt evidence.

    Self-test artifacts are written under:
      REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\selftest\last_run\
    The runner uses exact X:\RedSalamander.Perf on the repository drive by default.
    X: may be replaced by any fixed local drive through -TestRoot or
    REDSALAMANDER_TEST_ROOT. No test-generated local data is written outside it.

.PARAMETER Suite
    Which suite(s) to run:
      All       - Run all three self-test suites (default)
      Compare   - Run only --compare-selftest
      Commands  - Run only --commands-selftest
      FileOps   - Run only --fileops-selftest
      CI        - Run the GitHub Actions PR gate through the unified runner
      Full      - Run all self-tests plus standalone/native and script tests

.PARAMETER SkipBuild
    When set, skips the build step only when a schema-valid full-solution receipt
    matches this checkout/profile and every attested output still matches its digest.

.PARAMETER ApplyLegacySandboxCleanup
    Retired compatibility switch. The runner rejects it because test tooling never
    cleans paths outside X:\RedSalamander.Perf.

.PARAMETER SkipLegacySandboxCleanup
    Deprecated no-op compatibility switch. Outside-root cleanup no longer exists.

.PARAMETER AllowExternalTestRoot
    Allows first-time initialization of exact X:\RedSalamander.Perf on a fixed local
    drive. Arbitrary absolute roots remain forbidden.

.PARAMETER TestRoot
    Exact X:\RedSalamander.Perf root. When omitted, REDSALAMANDER_TEST_ROOT is used;
    if that is also unset, X: is the repository drive.

.PARAMETER RunId
    Optional exact run identifier used by an orchestrating test helper. Normal callers
    should omit it so the runner creates a unique identifier.

.PARAMETER AllowAlternateVolumeTestRoot
    Allows File Operations selftests to create their marked
    X:\RedSalamander.Perf root on one alternate fixed local volume. An enabled
    path.local.alternate machine resource restricts selection to that exact root.

.PARAMETER FailFast
    When set, passes --selftest-fail-fast to abort after first case failure.

.PARAMETER TimeoutMultiplier
    Scales all test timeouts. Use >1.0 on slow CI machines. Default: 1.0.

.PARAMETER CaseFilter
    Run only the matching case name or prefix (passed as --selftest-case=...).

.PARAMETER CommandsFamily
    With Suite Commands, runs one governed Commands source family in its own process.
    Family evidence is focused iteration coverage and never completes a Full repository verdict.

.PARAMETER SelfTestRepeat
    Repeat each matched in-product self-test case N times in-process (passed as --selftest-repeat=N).

.PARAMETER SelfTestShuffleSeed
    Shuffle matched in-product self-test case/phase order with the supplied seed (passed as --selftest-shuffle=SEED).

.PARAMETER SelfTestFlakyProofCase
    Debug self-test classifier proof hook. The named case fails only in the suite context and passes isolated/shuffle retries.

.PARAMETER SelfTestOrderProofCase
    Debug self-test classifier proof hook. The named case fails in suite/shuffle context and passes isolated retries.

.PARAMETER PerfBudgetPath
    Optional JSON5 perf budget file passed to native in-product self-test suites.

.PARAMETER RequirePerfBudgets
    Fail native perf selftests when the budget file is missing, has no current-machine entry, or has no hard entry for the current build.

.PARAMETER ClassifyFailures
    Rerun failed entries/cases once to classify blocking failures as FLAKY, REGRESSION,
    or ISOLATION_SUSPECT. Suite CI enables this automatically.

.PARAMETER QuarantinePath
    Optional JSONL blocking repair ledger. When omitted, the runner reads
    Tools\test-quarantine.jsonl if that file exists; absence of the default file
    is the canonical no-quarantine state. Invalid or active entries in an
    explicit or default ledger keep the aggregate run red.

.PARAMETER Platform
    Target platform. Default: x64.

.PARAMETER Configuration
    Build configuration. Default: Debug (self-tests require Debug builds).

.PARAMETER ExePath
    Compatibility spelling for explicitly naming the selected profile's canonical
    .build\<Platform>\<Configuration>\RedSalamander.exe path. It cannot select a
    foreign repository, profile, or arbitrary executable.

.PARAMETER ExplainPlan
    Computes and publishes the deterministic affected-set shadow decision without
    building or launching any validation entry. It never reuses test evidence.

.PARAMETER ImpactBase
    Optional local commit-ish whose merge base with HEAD defines committed impact.
    The planner never fetches a missing revision.

.PARAMETER ExplainPlanPath
    Optional output path for -ExplainPlan. It must remain beneath
    X:\RedSalamander.Perf\evidence.

.PARAMETER BuildNumber
    Optional explicit build number forwarded to build.ps1. Use this for reproducible
    closeout or when local persisted version state is intentionally unavailable.

.PARAMETER ValidationMode
    Fresh executes every entry. Resume is available only for Suite Full with an explicit
    ResumeFrom origin and reuses only exact-artifact promoted evidence. Affected is an
    opt-in Suite Full iteration mode that executes the governed impacted set and reports
    repository status NOT_EVALUATED.

.PARAMETER ResumeFrom
    Run ID or run-directory path beneath X:\RedSalamander.Perf\evidence\runs used by Resume.

.OUTPUTS
    No supported pipeline objects. Writes a color summary and runner-owned JSON, JSONL,
    Markdown, trace, and self-test artifacts beneath REDSALAMANDER_TEST_ROOT. Exit codes:
    0 PASSED, 1 FAILED, 2 NOT_EVALUATED, 3 invalid CLI/plan,
    4 blocked/missing attestation, and 5 corrupt evidence.

.NOTES
    Prerequisites: PowerShell 7, the repository build toolchain unless -SkipBuild is used, and an interactive desktop for UI-driven suites. Side effects: acquires the selected platform/configuration artifact lock, holds the Windows-session mutex only around marked interactive entries, may build the solution, launches contained child test processes, and replaces runner-owned directories beneath the selected marked sandbox. Alternate-volume roots require an explicit switch. It does not terminate unrelated processes or block other profiles/worktrees. Primary consumers: local closeout, GitHub CI, and nightly test workflows.

.EXAMPLE
    .\Tools\Run-AllTests.ps1
    Build + run all suites with summary.

.EXAMPLE
    .\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild
    Run Compare suite only (skip build).

.EXAMPLE
    .\Tools\Run-AllTests.ps1 -FailFast -TimeoutMultiplier 2.0
    Run all suites, stop on first failure, with 2x timeouts.
#>

[CmdletBinding()]
param(
    [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'CI', 'Full')]
    [string]$Suite = 'All',

    [switch]$SkipBuild,

    [switch]$ApplyLegacySandboxCleanup,

    [switch]$SkipLegacySandboxCleanup,

    [switch]$AllowExternalTestRoot,

    [string]$TestRoot = $env:REDSALAMANDER_TEST_ROOT,

    [string]$RunId = '',

    [switch]$AllowAlternateVolumeTestRoot,

    [switch]$FailFast,

    [double]$TimeoutMultiplier = 1.0,

    [string]$CaseFilter = '',

    [string]$CommandsFamily = '',

    [uint32]$SelfTestRepeat = 1,

    [string]$SelfTestShuffleSeed = '',

    [string]$SelfTestFlakyProofCase = '',

    [string]$SelfTestOrderProofCase = '',

    [string]$PerfBudgetPath = '',


    [switch]$RequirePerfBudgets,

    [switch]$ClassifyFailures,

    [string]$QuarantinePath = '',

    [ValidateSet('x64', 'ARM64')]
    [string]$Platform = 'x64',

    [string]$Configuration = 'Debug',

    [string]$ExePath = '',

    [switch]$ExplainPlan,

    [string]$ImpactBase = '',

    [string]$ExplainPlanPath = '',

    [ValidateRange(0, [int]::MaxValue)]
    [int]$BuildNumber = 0,

    [ValidateSet('Fresh', 'Resume', 'Affected')]
    [string]$ValidationMode = 'Fresh',

    [string]$ResumeFrom = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Exit-RSInvalidRunnerInvocation {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "INVALID CLI/PLAN: $Message" -ForegroundColor Red
    exit 3
}

if ($ValidationMode -eq 'Resume') {
    if ($Suite -ne 'Full') { Exit-RSInvalidRunnerInvocation 'Resume is supported only for -Suite Full.' }
    if ([string]::IsNullOrWhiteSpace($ResumeFrom)) { Exit-RSInvalidRunnerInvocation 'Resume requires -ResumeFrom <run-id-or-directory>.' }
    if ($ExplainPlan) { Exit-RSInvalidRunnerInvocation 'Use -ExplainPlan without Resume; explanation never consumes result evidence.' }
} elseif (-not [string]::IsNullOrWhiteSpace($ResumeFrom)) {
    Exit-RSInvalidRunnerInvocation '-ResumeFrom is valid only with -ValidationMode Resume.'
}
if ($ValidationMode -eq 'Affected' -and $Suite -ne 'Full') {
    Exit-RSInvalidRunnerInvocation 'Affected is supported only for -Suite Full.'
}
if ($ApplyLegacySandboxCleanup) {
    Exit-RSInvalidRunnerInvocation '-ApplyLegacySandboxCleanup is retired: normal test execution never mutates paths outside X:\RedSalamander.Perf.'
}
if ($Configuration -notin @('Debug', 'Release', 'ASan Debug')) {
    Exit-RSInvalidRunnerInvocation "Unsupported configuration '$Configuration'."
}
if ($Platform -eq 'ARM64' -and $Configuration -eq 'ASan Debug') {
    Exit-RSInvalidRunnerInvocation 'ASan Debug is supported only for x64.'
}

# --- Helpers ---

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $repoRoot) { $repoRoot = $PSScriptRoot | Split-Path -Parent }

$testRunPlanModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
if (-not (Test-Path $testRunPlanModule)) {
    Write-Host "ERROR: test run plan helper not found at $testRunPlanModule" -ForegroundColor Red
    exit 1
}

Import-Module $testRunPlanModule -Force -ErrorAction Stop

if (-not [string]::IsNullOrWhiteSpace($CommandsFamily)) {
    $knownCommandsFamilies = @(Get-RSCommandsSelfTestFamilies | ForEach-Object { $_.name })
    if ($Suite -ne 'Commands' -or $CommandsFamily -notin $knownCommandsFamilies) {
        Exit-RSInvalidRunnerInvocation "-CommandsFamily requires -Suite Commands and one of: $($knownCommandsFamilies -join ', ')."
    }
}

$selectedTestRoot = Get-RSTestSandboxRoot `
    -RepoRoot $repoRoot `
    -TestRootOverride $TestRoot `
    -SelfTestRootOverride '' `
    -AllowExternalInitialization:$AllowExternalTestRoot
[void](Initialize-RSTestSandboxRoot `
        -TestRoot $selectedTestRoot `
        -RepoRoot $repoRoot `
        -AllowExternalInitialization:$AllowExternalTestRoot)
$validationEvidenceRoot = Join-Path $selectedTestRoot 'evidence'
$testMachineResources = Read-RSTestMachineResources `
    -TestRoot $selectedTestRoot `
    -SchemaPath (Join-Path $repoRoot 'Specs\Testing\TestMachineResources.schema.json')

$processStreamingModule = Join-Path $repoRoot 'Tools\Modules\Build\ProcessStreaming.psm1'
if (-not (Test-Path $processStreamingModule)) {
    Write-Host "ERROR: process streaming helper not found at $processStreamingModule" -ForegroundColor Red
    exit 1
}

Import-Module $processStreamingModule -Force -ErrorAction Stop

$sanitizedEnvironmentModule = Join-Path $repoRoot 'Tools\Modules\Build\SanitizedEnvironment.psm1'
if (-not (Test-Path $sanitizedEnvironmentModule)) {
    Write-Host "ERROR: contained process helper not found at $sanitizedEnvironmentModule" -ForegroundColor Red
    exit 1
}

Import-Module $sanitizedEnvironmentModule -Force -ErrorAction Stop

$artifactOperationLockModule = Join-Path $repoRoot 'Tools\Modules\Build\ArtifactOperationLock.psm1'
if (-not (Test-Path $artifactOperationLockModule)) {
    Write-Host "ERROR: artifact operation lock helper not found at $artifactOperationLockModule" -ForegroundColor Red
    exit 1
}

Import-Module $artifactOperationLockModule -Force -ErrorAction Stop

$buildEvidenceModule = Join-Path $repoRoot 'Tools\Modules\Build\BuildEvidence.psm1'
if (-not (Test-Path $buildEvidenceModule)) {
    Write-Host "ERROR: build evidence helper not found at $buildEvidenceModule" -ForegroundColor Red
    exit 1
}

Import-Module $buildEvidenceModule -Force -ErrorAction Stop

$validationFingerprintModule = Join-Path $repoRoot 'Tools\Modules\Testing\ValidationFingerprint.psm1'
Import-Module $validationFingerprintModule -Force -ErrorAction Stop

$validationImpactModule = Join-Path $repoRoot 'Tools\Modules\Testing\ValidationImpact.psm1'
Import-Module $validationImpactModule -Force -ErrorAction Stop

$validationEvidenceModule = Join-Path $repoRoot 'Tools\Modules\Testing\ValidationEvidence.psm1'
Import-Module $validationEvidenceModule -Force -ErrorAction Stop

function Import-RSRunnerValidationModules {
    # build.ps1 and in-process Pester entries import modules in nested script
    # scopes. Restore the runner-owned command surface after either boundary so
    # PowerShell scope teardown cannot remove Startrail commands mid-campaign.
    Import-Module $buildEvidenceModule -Force -ErrorAction Stop
    Import-Module $validationFingerprintModule -Force -ErrorAction Stop
    Import-Module $validationImpactModule -Force -ErrorAction Stop
    Import-Module $validationEvidenceModule -Force -ErrorAction Stop
}

function Write-Header([string]$text) {
    $line = '=' * 70
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  $text" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
}

function Write-SubHeader([string]$text) {
    Write-Host ""
    Write-Host "--- $text ---" -ForegroundColor Yellow
}

function Format-Duration([uint64]$ms) {
    if ($ms -lt 1000) { return "${ms}ms" }
    if ($ms -lt 60000) { return "{0:N1}s" -f ($ms / 1000.0) }
    $minutes = [math]::Floor($ms / 60000)
    $seconds = ($ms % 60000) / 1000.0
    return "{0}m {1:N1}s" -f $minutes, $seconds
}

function Get-StatusColor([string]$status) {
    switch ($status) {
        'passed'  { return 'Green' }
        'failed'  { return 'Red' }
        'crashed' { return 'Red' }
        'skipped' { return 'DarkYellow' }
        default   { return 'Gray' }
    }
}

function Get-JsonValue($object, [string[]]$names, $defaultValue = $null) {
    if ($null -eq $object) { return $defaultValue }
    if ($object -is [System.Collections.IDictionary]) {
        foreach ($name in $names) {
            if ($object.Contains($name)) {
                return $object[$name]
            }
        }
        return $defaultValue
    }
    foreach ($name in $names) {
        $property = $object.PSObject.Properties[$name]
        if ($null -ne $property) {
            return $property.Value
        }
    }
    return $defaultValue
}

function Test-RSTestEntryRequiresInteractiveDesktop {
    param(
        [AllowNull()]
        [object]$Entry
    )

    return [bool](Get-RSObjectValue `
        -Object $Entry `
        -Names @('RequiresInteractiveDesktop', 'requires_interactive_desktop') `
        -DefaultValue $false)
}

function Invoke-RSInteractiveDesktopOperation {
    param(
        [bool]$RequiresInteractiveDesktop,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Operation
    )

    if (-not $RequiresInteractiveDesktop -or
        [Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        return & $Operation
    }

    # Foreground, UI Automation, pointer, and clipboard cases share one interactive
    # Windows desktop across worktrees. Only the marked entry execution owns this mutex.
    $interactiveTestMutex = [System.Threading.Mutex]::new(
        $false,
        'Local\RedSalamander.InteractiveTestRun.v1')
    $interactiveTestMutexAcquired = $false
    $executionStateArmed = $false
    try {
        try {
            $interactiveTestMutexAcquired = $interactiveTestMutex.WaitOne([TimeSpan]::FromHours(3))
        }
        catch [System.Threading.AbandonedMutexException] {
            $interactiveTestMutexAcquired = $true
            Write-Warning 'Acquired an abandoned interactive-test mutex from a terminated entry.'
        }
        if (-not $interactiveTestMutexAcquired) {
            throw 'Timed out after three hours waiting for another RedSalamander interactive test entry in this Windows session.'
        }

        if (-not ('RedSalamander.TestExecutionState' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace RedSalamander
{
    public static class TestExecutionState
    {
        [DllImport("kernel32.dll")]
        public static extern uint SetThreadExecutionState(uint executionState);
    }
}
'@
        }

        $executionState = [Convert]::ToUInt32('80000003', 16) # ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        $executionStateArmed = [RedSalamander.TestExecutionState]::SetThreadExecutionState($executionState) -ne 0
        if (-not $executionStateArmed) {
            Write-Warning 'Unable to prevent system/display idle while the interactive test entry runs.'
        }

        return & $Operation
    }
    finally {
        if ($executionStateArmed) {
            $continuousExecutionState = [Convert]::ToUInt32('80000000', 16) # ES_CONTINUOUS
            [void][RedSalamander.TestExecutionState]::SetThreadExecutionState($continuousExecutionState)
        }
        if ($interactiveTestMutexAcquired) {
            $interactiveTestMutex.ReleaseMutex()
        }
        $interactiveTestMutex.Dispose()
    }
}

function Resolve-VSTestConsole {
    $command = Get-Command 'vstest.console.exe' -ErrorAction SilentlyContinue
    if ($command -and $command.Source) {
        return $command.Source
    }

    $vswhereCandidates = @(
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\Installer\vswhere.exe"
    )

    $vswhere = $vswhereCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $vswhere) {
        return $null
    }

    $installations = @()
    try {
        $json = & $vswhere -all -products '*' -prerelease -format json 2>$null
        if ($LASTEXITCODE -eq 0 -and $json) {
            $installations = @($json | ConvertFrom-Json)
        }
    } catch {
        return $null
    }

    foreach ($installation in @($installations | Sort-Object { try { [version]$_.installationVersion } catch { [version]'0.0' } } -Descending)) {
        if (-not $installation.installationPath) {
            continue
        }

        $candidate = Join-Path $installation.installationPath 'Common7\IDE\Extensions\TestPlatform\vstest.console.exe'
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Invoke-RSTestPlanEntry {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [string]$ArtifactRoot = ''
    )

    $start = Get-Date
    $exitCode = 0
    $failureReason = ''
    $parsedResult = $null
    $resultParsed = $false
    $coverageValid = $true
    $passedCount = [int64]0
    $failedCount = [int64]0
    $skippedCount = [int64]0
    $totalCount = [int64]0
    $pendingCount = [int64]0
    $inconclusiveCount = [int64]0
    $outputLogPath = ''
    if (-not [string]::IsNullOrWhiteSpace($ArtifactRoot) -and $Entry.Kind -in @('Executable', 'CppUnitTest')) {
        $safeName = ([string]$Entry.Name) -replace '[^A-Za-z0-9_.-]', '_'
        $outputLogPath = Join-Path $ArtifactRoot "$safeName.output.log"
    }

    try {
        switch ($Entry.Kind) {
            'Executable' {
                if (-not (Test-Path $Entry.Path)) {
                    throw "Executable not found: $($Entry.Path)"
                }
                $exitCode = Invoke-RSInteractiveDesktopOperation `
                    -RequiresInteractiveDesktop (Test-RSTestEntryRequiresInteractiveDesktop -Entry $Entry) `
                    -Operation {
                        Invoke-RSStreamingProcess `
                            -FilePath $Entry.Path `
                            -Arguments @($Entry.Arguments) `
                            -WorkingDirectory $Entry.WorkingDirectory `
                            -LogPath $outputLogPath
                    }
            }
            'CppUnitTest' {
                if (-not (Test-Path $Entry.Path)) {
                    throw "CppUnitTest DLL not found: $($Entry.Path)"
                }

                $vstest = Resolve-VSTestConsole
                if (-not $vstest) {
                    throw 'vstest.console.exe was not found. Install Visual Studio Test Platform or run from a Developer PowerShell.'
                }

                $arguments = @($Entry.Path)
                $exitCode = Invoke-RSStreamingProcess -FilePath $vstest -Arguments $arguments -WorkingDirectory $Entry.WorkingDirectory -LogPath $outputLogPath
            }
            'Pester' {
                if (-not (Test-Path $Entry.Path)) {
                    throw "Pester test path not found: $($Entry.Path)"
                }

                $primaryTestRoot = [string]$env:REDSALAMANDER_TEST_ROOT
                $runId = [string]$env:REDSALAMANDER_TEST_RUN_ID
                if ([string]::IsNullOrWhiteSpace($primaryTestRoot) -or
                    [string]::IsNullOrWhiteSpace($runId)) {
                    throw 'Pester execution requires the runner-owned TestSandbox root and run id.'
                }
                $pesterTestRoot = Get-RSToolingPesterSandboxRoot `
                    -RepoRoot $RepoRoot `
                    -PrimaryTestRoot $primaryTestRoot
                if (-not [string]::Equals(
                    $pesterTestRoot,
                    [IO.Path]::GetFullPath($primaryTestRoot).TrimEnd('\'),
                    [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw "Pester must use the selected X:\RedSalamander.Perf root without an auxiliary sandbox."
                }

                $pesterParameters = New-RSPesterInvokeParameters -Path $Entry.Path -Arguments @($Entry.Arguments)
                $pesterResult = Invoke-Pester @pesterParameters
                $parsedResult = $pesterResult
                $resultParsed = $null -ne $pesterResult
                $passedCount = [int64](Get-JsonValue $pesterResult @('PassedCount', 'Passed') 0)
                $failedCount = [int64](Get-JsonValue $pesterResult @('FailedCount', 'Failed') 0)
                $skippedCount = [int64](Get-JsonValue $pesterResult @('SkippedCount', 'Skipped') 0)
                $pendingCount = [int64](Get-JsonValue $pesterResult @('PendingCount', 'Pending') 0)
                $inconclusiveCount = [int64](Get-JsonValue $pesterResult @('InconclusiveCount', 'Inconclusive') 0)
                $totalCount = [int64](Get-JsonValue $pesterResult @('TotalCount', 'Total') `
                    ($passedCount + $failedCount + $skippedCount + $pendingCount + $inconclusiveCount))
                $exitCode = if ($failedCount -gt 0) { 1 } else { 0 }
            }
            'PowerShellScript' {
                if (-not (Test-Path $Entry.Path)) {
                    throw "PowerShell test script not found: $($Entry.Path)"
                }

                Push-Location $Entry.WorkingDirectory
                try {
                    $scriptArguments = @($Entry.Arguments)
                    & $Entry.Path @scriptArguments
                    $exitCode = if ($null -ne $LASTEXITCODE) { $LASTEXITCODE } else { 0 }
                } finally {
                    Pop-Location
                }
            }
            default {
                throw "Unsupported test plan entry kind: $($Entry.Kind)"
            }
        }
    } catch {
        $exitCode = 1
        $failureReason = $_.Exception.Message
    }

    $end = Get-Date
    $wallMs = [uint64](($end - $start).TotalMilliseconds)

    if ($Entry.Kind -ne 'Pester') {
        $passedCount = if ($exitCode -eq 0) { 1 } else { 0 }
        $failedCount = if ($exitCode -eq 0) { 0 } else { 1 }
        $totalCount = 1
    }

    $failures = @()
    if ($exitCode -ne 0) {
        if ([string]::IsNullOrWhiteSpace($failureReason)) {
            $failureReason = "Process exited with code $exitCode."
        }

        $failures += [pscustomobject]@{
            name = $Entry.Name
            status = 'failed'
            reason = $failureReason
            durationMs = $wallMs
        }
    }

    return @{
        Name       = $Entry.Name
        Kind       = $Entry.Kind
        ExitCode   = $exitCode
        WallMs     = $wallMs
        Parsed     = $parsedResult
        ResultParsed = $resultParsed
        CoverageValid = $coverageValid
        Cases      = @()
        Passed     = $passedCount
        Failed     = $failedCount
        Skipped    = $skippedCount
        Total      = $totalCount
        Pending    = $pendingCount
        Inconclusive = $inconclusiveCount
        DurationMs = $wallMs
        Failures   = $failures
        OutputLogPath = $outputLogPath
    }
}

function New-RSTestRetryPlanEntry {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [string[]]$Arguments = @()
    )

    [pscustomobject]@{
        Name = $Name
        Kind = [string](Get-RSObjectValue -Object $Entry -Names @('Kind', 'kind') -DefaultValue '')
        Path = [string](Get-RSObjectValue -Object $Entry -Names @('Path', 'path') -DefaultValue '')
        Arguments = @($Arguments)
        WorkingDirectory = [string](Get-RSObjectValue -Object $Entry -Names @('WorkingDirectory', 'working_directory') -DefaultValue '')
        JsonName = [string](Get-RSObjectValue -Object $Entry -Names @('JsonName', 'json_name') -DefaultValue '')
        RequiresInteractiveDesktop = [bool](Get-RSObjectValue -Object $Entry -Names @('RequiresInteractiveDesktop', 'requires_interactive_desktop') -DefaultValue $false)
    }
}

function New-RSTestRetryAttempt {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Mode,

        [Parameter(Mandatory = $true)]
        [int]$ExitCode,

        [uint64]$DurationMs = 0,

        [string]$OutputLogPath = '',

        [string]$Reason = '',

        [string]$ShuffleSeed = ''
    )

    [pscustomobject]@{
        name = $Name
        mode = $Mode
        exit_code = $ExitCode
        passed = ($ExitCode -eq 0)
        duration_ms = $DurationMs
        output_log_path = $OutputLogPath
        reason = $Reason
        shuffle_seed = $ShuffleSeed
    }
}

function Get-RSSelfTestShuffleTriageSeeds {
    param(
        [string[]]$Arguments = @(),

        [int]$Count = 3
    )

    $seeds = @()
    foreach ($argument in @($Arguments)) {
        if ($argument -match '^--selftest-shuffle=(.+)$') {
            $seed = $Matches[1]
            if (-not [string]::IsNullOrWhiteSpace($seed) -and $seed -notin $seeds) {
                $seeds += $seed
            }
        }
    }

    while (@($seeds).Count -lt $Count) {
        $seed = [string](Get-Random -Minimum 1 -Maximum ([int]::MaxValue))
        if ($seed -notin $seeds) {
            $seeds += $seed
        }
    }

    return $seeds
}

function Get-RSTestResultFailureReason {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Result
    )

    $failureReasons = @(Get-RSObjectValue -Object $Result -Names @('Failures', 'failures') -DefaultValue @() | ForEach-Object {
            [string](Get-RSObjectValue -Object $_ -Names @('reason', 'Reason') -DefaultValue '')
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if (@($failureReasons).Count -eq 0) {
        return ''
    }

    return ($failureReasons -join '; ')
}

function Invoke-RSTestEntryClassificationRetry {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [string]$ArtifactRoot = ''
    )

    $retryEntry = New-RSTestRetryPlanEntry `
        -Entry $Entry `
        -Name "$($Entry.Name).retry1" `
        -Arguments @($Entry.Arguments)
    $retryResult = Invoke-RSTestPlanEntry -Entry $retryEntry -RepoRoot $RepoRoot -ArtifactRoot $ArtifactRoot
    $retryPassed = Test-RSTestResultPassed -Result $retryResult
    $retryExitCode = if ($retryPassed) { 0 } else { [int](Get-RSObjectValue -Object $retryResult -Names @('ExitCode', 'exit_code') -DefaultValue 1) }

    return New-RSTestRetryAttempt `
        -Name $Entry.Name `
        -Mode 'entry' `
        -ExitCode $retryExitCode `
        -DurationMs ([uint64](Get-RSObjectValue -Object $retryResult -Names @('WallMs', 'DurationMs', 'duration_ms') -DefaultValue 0)) `
        -OutputLogPath ([string](Get-RSObjectValue -Object $retryResult -Names @('OutputLogPath', 'output_log_path') -DefaultValue '')) `
        -Reason (Get-RSTestResultFailureReason -Result $retryResult)
}

function Invoke-RSSelfTestClassificationRetry {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [object]$SuiteResult
    )

    $failureNames = @(Get-RSObjectValue -Object $SuiteResult -Names @('Failures', 'failures') -DefaultValue @() | ForEach-Object {
            [string](Get-RSObjectValue -Object $_ -Names @('name', 'Name') -DefaultValue '')
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    $nonCaseFailureNames = @($failureNames | Where-Object { $_ -in @('selftest_result_coverage', 'setup') })
    $failedCaseNames = @($failureNames | Where-Object { $_ -notin @('selftest_result_coverage', 'setup') })

    if (@($failedCaseNames).Count -eq 0 -or @($nonCaseFailureNames).Count -gt 0) {
        $started = Get-Date
        $exitCode = Invoke-RSSelfTestProcess -Entry $Entry -Arguments @($Entry.Arguments)
        $ended = Get-Date
        return @(
            New-RSTestRetryAttempt `
                -Name $Entry.Name `
                -Mode 'entry' `
                -ExitCode $exitCode `
                -DurationMs ([uint64](($ended - $started).TotalMilliseconds)) `
                -Reason $(if ($exitCode -ne 0) { "Process exited with code $exitCode." } else { '' })
        )
    }

    $retryAttempts = @()
    $retryPlan = @(Get-RSSelfTestClassificationRetryPlan `
            -Entry $Entry `
            -FailureNames $failedCaseNames)

    foreach ($plan in @($retryPlan | Where-Object { [string]$_.mode -eq 'failed-case' })) {
        $started = Get-Date
        $exitCode = Invoke-RSSelfTestProcess -Entry $Entry -Arguments @($plan.arguments)
        $ended = Get-Date
        $retryAttempts += New-RSTestRetryAttempt `
            -Name $plan.name `
            -Mode 'failed-case' `
            -ExitCode $exitCode `
            -DurationMs ([uint64](($ended - $started).TotalMilliseconds)) `
            -Reason $(if ($exitCode -ne 0) { "Process exited with code $exitCode." } else { '' })
    }

    $allFailedCaseRetriesPassed = $true
    foreach ($retryAttempt in @($retryAttempts)) {
        if (-not [bool](Get-RSObjectValue -Object $retryAttempt -Names @('passed', 'Passed') -DefaultValue $false)) {
            $allFailedCaseRetriesPassed = $false
            break
        }
    }

    if ($allFailedCaseRetriesPassed) {
        $shuffleRetryPlan = @(Get-RSSelfTestClassificationRetryPlan `
                -Entry $Entry `
                -FailureNames $failedCaseNames `
                -ShuffleTriageSeeds @(Get-RSSelfTestShuffleTriageSeeds -Arguments @($Entry.Arguments) -Count 3) | Where-Object { [string]$_.mode -eq 'shuffle-triage' })
        foreach ($plan in $shuffleRetryPlan) {
            $started = Get-Date
            $exitCode = Invoke-RSSelfTestProcess -Entry $Entry -Arguments @($plan.arguments)
            $ended = Get-Date
            $retryAttempts += New-RSTestRetryAttempt `
                -Name $Entry.Name `
                -Mode 'shuffle-triage' `
                -ExitCode $exitCode `
                -DurationMs ([uint64](($ended - $started).TotalMilliseconds)) `
                -Reason $(if ($exitCode -ne 0) { "Shuffle triage seed $($plan.shuffle_seed) exited with code $exitCode." } else { '' }) `
                -ShuffleSeed $plan.shuffle_seed
        }
    }

    return $retryAttempts
}

function Invoke-RSSelfTestCaseList {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry
    )

    $arguments = @(Get-RSSelfTestListArguments -SelfTestArguments @($Entry.Arguments))

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $Entry.Path
    $startInfo.WorkingDirectory = $Entry.WorkingDirectory
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true
    if ($startInfo.PSObject.Properties['ArgumentList'] -and $null -ne $startInfo.ArgumentList) {
        foreach ($argument in $arguments) {
            [void]$startInfo.ArgumentList.Add($argument)
        }
    } else {
        $startInfo.Arguments = (($arguments | ForEach-Object { ConvertTo-RSStreamingQuotedArgument $_ }) -join ' ')
    }

    $processResult = Invoke-RSInteractiveDesktopOperation `
        -RequiresInteractiveDesktop (Test-RSTestEntryRequiresInteractiveDesktop -Entry $Entry) `
        -Operation {
            $process = $null
            try {
                $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
                $stdoutTask = $process.StandardOutput.ReadToEndAsync()
                $stderrTask = $process.StandardError.ReadToEndAsync()
                $process.WaitForExit()
                return [pscustomobject]@{
                    Stdout = $stdoutTask.GetAwaiter().GetResult()
                    Stderr = $stderrTask.GetAwaiter().GetResult()
                    ExitCode = $process.ExitCode
                }
            }
            finally {
                Close-RSContainedProcess -Process $process
            }
        }

    $stdout = [string]$processResult.Stdout
    $stderr = [string]$processResult.Stderr
    $exitCode = [int]$processResult.ExitCode

    if ($exitCode -ne 0) {
        throw "Case listing failed with exit code ${exitCode}: $stderr"
    }

    if ([string]::IsNullOrWhiteSpace($stdout)) {
        throw 'Case listing produced no JSON output.'
    }

    return ($stdout | ConvertFrom-Json)
}

function Invoke-RSSelfTestProcess {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $Entry.Path
    $startInfo.WorkingDirectory = $Entry.WorkingDirectory
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    if ($startInfo.PSObject.Properties['ArgumentList'] -and $null -ne $startInfo.ArgumentList) {
        foreach ($argument in $Arguments) {
            [void]$startInfo.ArgumentList.Add($argument)
        }
    } else {
        $startInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-RSStreamingQuotedArgument $_ }) -join ' ')
    }

    return Invoke-RSInteractiveDesktopOperation `
        -RequiresInteractiveDesktop (Test-RSTestEntryRequiresInteractiveDesktop -Entry $Entry) `
        -Operation {
            $process = $null
            try {
                $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
                $process.WaitForExit()
                return $process.ExitCode
            }
            finally {
                Close-RSContainedProcess -Process $process
            }
        }
}

function Invoke-RSTestQuarantineRepairAttempt {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Attempt,

        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$ArtifactRoot
    )

    $harness = [string](Get-RSObjectValue -Object $Attempt -Names @('harness') -DefaultValue '')
    $name = [string](Get-RSObjectValue -Object $Attempt -Names @('name') -DefaultValue '')
    $mode = [string](Get-RSObjectValue -Object $Attempt -Names @('mode') -DefaultValue '')
    $safeName = ("$harness.$name" -replace '[^A-Za-z0-9_.-]', '_')
    $repairArtifactRoot = Join-Path $ArtifactRoot 'quarantine'
    New-Item -ItemType Directory -Path $repairArtifactRoot -Force | Out-Null

    $attemptEntry = [pscustomobject]@{
        Name = "$safeName.repair"
        Kind = [string](Get-RSObjectValue -Object $Attempt -Names @('kind') -DefaultValue '')
        Path = [string](Get-RSObjectValue -Object $Attempt -Names @('path') -DefaultValue '')
        Arguments = @(Get-RSObjectValue -Object $Attempt -Names @('arguments') -DefaultValue @())
        WorkingDirectory = [string](Get-RSObjectValue -Object $Attempt -Names @('working_directory') -DefaultValue '')
        JsonName = [string](Get-RSObjectValue -Object $Attempt -Names @('json_name') -DefaultValue '')
        RequiresInteractiveDesktop = [bool](Get-RSObjectValue -Object $Attempt -Names @('requires_interactive_desktop') -DefaultValue $false)
    }

    $started = Get-Date
    $exitCode = 1
    $outputLogPath = ''
    $reason = ''

    try {
        if ($mode -eq 'selftest-case') {
            $previousSelfTestRoot = [Environment]::GetEnvironmentVariable('REDSALAMANDER_SELFTEST_ROOT', 'Process')
            $repairSelfTestRoot = Join-Path $repairArtifactRoot $safeName
            try {
                [Environment]::SetEnvironmentVariable('REDSALAMANDER_SELFTEST_ROOT', $repairSelfTestRoot, 'Process')
                New-Item -ItemType Directory -Path $repairSelfTestRoot -Force | Out-Null
                $exitCode = Invoke-RSSelfTestProcess -Entry $attemptEntry -Arguments @($attemptEntry.Arguments)
                $outputLogPath = Join-Path $repairSelfTestRoot 'last_run'
            } finally {
                [Environment]::SetEnvironmentVariable('REDSALAMANDER_SELFTEST_ROOT', $previousSelfTestRoot, 'Process')
            }
        } else {
            $result = Invoke-RSTestPlanEntry -Entry $attemptEntry -RepoRoot $RepoRoot -ArtifactRoot $repairArtifactRoot
            $exitCode = [int](Get-RSObjectValue -Object $result -Names @('ExitCode', 'exit_code') -DefaultValue 1)
            $outputLogPath = [string](Get-RSObjectValue -Object $result -Names @('OutputLogPath', 'output_log_path') -DefaultValue '')
            $reason = Get-RSTestResultFailureReason -Result $result
        }
    } catch {
        $exitCode = 1
        $reason = $_.Exception.Message
    }

    $ended = Get-Date
    if ($exitCode -ne 0 -and [string]::IsNullOrWhiteSpace($reason)) {
        $reason = "Repair lane exited with code $exitCode."
    }

    [pscustomobject]@{
        harness = $harness
        name = $name
        mode = $mode
        owner = [string](Get-RSObjectValue -Object $Attempt -Names @('owner') -DefaultValue '')
        expires = [string](Get-RSObjectValue -Object $Attempt -Names @('expires') -DefaultValue '')
        issue = [string](Get-RSObjectValue -Object $Attempt -Names @('issue') -DefaultValue '')
        exit_code = $exitCode
        passed = ($exitCode -eq 0)
        duration_ms = [uint64](($ended - $started).TotalMilliseconds)
        output_log_path = $outputLogPath
        reason = $reason
    }
}

function Format-RSCoverageReason {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Coverage
    )

    $parts = @()
    if (($Coverage.PSObject.Properties['NoExpectedCases']) -and [bool]$Coverage.NoExpectedCases) {
        $parts += 'no expected cases matched the requested filter'
    }
    if (@($Coverage.DuplicateExpected).Count -gt 0) {
        $parts += "duplicate expected: $(@($Coverage.DuplicateExpected) -join ', ')"
    }
    if (@($Coverage.DuplicateActual).Count -gt 0) {
        $parts += "duplicate actual: $(@($Coverage.DuplicateActual) -join ', ')"
    }
    if (@($Coverage.Missing).Count -gt 0) {
        $parts += "missing: $(@($Coverage.Missing) -join ', ')"
    }
    if (@($Coverage.Extra).Count -gt 0) {
        $parts += "extra: $(@($Coverage.Extra) -join ', ')"
    }

    if (@($parts).Count -eq 0) {
        return 'result coverage mismatch'
    }

    return ($parts -join '; ')
}

$expectedExeFullPath = [IO.Path]::GetFullPath(
    (Join-Path $repoRoot ".build\$Platform\$Configuration\RedSalamander.exe"))
if ($ExplainPlan) {
    if (-not [string]::IsNullOrWhiteSpace($ExePath)) {
        Exit-RSInvalidRunnerInvocation '-ExplainPlan does not accept -ExePath because it does not launch an executable.'
    }
    $explainEntries = @(Get-RSTestRunPlan `
            -Suite $Suite -RepoRoot $repoRoot -Platform $Platform -Configuration $Configuration `
            -RedSalamanderExePath $expectedExeFullPath -TimeoutMultiplier $TimeoutMultiplier `
            -FailFast:$FailFast -CaseFilter $CaseFilter -CommandsFamily $CommandsFamily -RepeatCount $SelfTestRepeat `
            -ShuffleSeed $SelfTestShuffleSeed -PerfBudgetPath $PerfBudgetPath `
            -RequirePerfBudgets:$RequirePerfBudgets -SelfTestFlakyProofCase $SelfTestFlakyProofCase `
            -SelfTestOrderProofCase $SelfTestOrderProofCase)
    $coverageScope = if (-not [string]::IsNullOrWhiteSpace($CaseFilter)) {
        'filtered'
    } elseif (-not [string]::IsNullOrWhiteSpace($CommandsFamily) -or $Suite -in @('Compare', 'Commands', 'FileOps')) {
        'family'
    } else {
        'complete'
    }
    $explainValidationPlan = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite $Suite `
        -Entries $explainEntries -ValidationMode Affected -CoverageScope $coverageScope
    $explainSnapshot = Get-RSWorkspaceSnapshot -RepoRoot $repoRoot -ImpactBase $ImpactBase
    $impactManifestPath = Join-Path $repoRoot 'Tools\validation-impact.json'
    $impactSchemaPath = Join-Path $repoRoot 'Specs\Testing\ValidationImpact.schema.json'
    $impactManifest = Read-RSValidationImpactManifest -Path $impactManifestPath `
        -SchemaPath $impactSchemaPath -RepoRoot $repoRoot
    $trackedGraphInputs = @(& git -C $repoRoot ls-files '*.sln' '*.slnx' '*.vcxproj' '*.props' '*.targets' 'RuntimeDependencies.props')
    $governance = Test-RSValidationImpactGovernance -Manifest $impactManifest -TrackedPaths $trackedGraphInputs -RepoRoot $repoRoot
    if (-not $governance.IsValid) {
        Exit-RSInvalidRunnerInvocation "Validation impact governance is incomplete: $(@($governance.UnmappedBuildInputs) -join ', ')"
    }
    $explainDurationEstimates = Get-RSValidationDurationEstimates `
        -EvidenceRoot $validationEvidenceRoot `
        -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') `
        -EntryId @($explainValidationPlan.entries | ForEach-Object { [string]$_.id })
    $explainDecision = Get-RSValidationImpactDecision -Manifest $impactManifest `
        -WorkspaceSnapshot $explainSnapshot -ValidationPlan $explainValidationPlan `
        -EstimatedDurationMs $explainDurationEstimates -PlannerMode shadow
    $explainRoot = $validationEvidenceRoot
    [void](New-Item -ItemType Directory -Path $explainRoot -Force)
    $resolvedExplainPath = if ([string]::IsNullOrWhiteSpace($ExplainPlanPath)) {
        Join-Path $explainRoot 'explain-plan.json'
    } else {
        [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathFullyQualified($ExplainPlanPath)) {
                    $ExplainPlanPath
                } else {
                    Join-Path $explainRoot $ExplainPlanPath
                }))
    }
    Write-RSAtomicCanonicalJsonFile -LiteralPath $resolvedExplainPath -InputObject $explainDecision `
        -SchemaPath (Join-Path $repoRoot 'Specs\Testing\ValidationPlanDecision.schema.json') `
        -AllowedRoot $explainRoot
    Write-Host "Shadow validation plan: $($explainDecision.decision_digest)" -ForegroundColor Cyan
    Write-Host "Changed paths: $(@($explainDecision.changed_paths).Count); advisory build: $($explainDecision.advisory_build_action)" -ForegroundColor DarkGray
    foreach ($entry in $explainDecision.entries) {
        $scope = if ($entry.affected) { 'AFFECTED' } else { 'UNCHANGED' }
        $estimate = if ($entry.estimated_duration_ms -gt 0) { "$($entry.estimated_duration_ms) ms" } else { 'unknown' }
        Write-Host ("  {0,-10} {1,-42} {2,12}  {3}" -f $scope, $entry.id, $estimate, (@($entry.reason_codes) -join ','))
    }
    Write-Host "Estimated fresh duration: $($explainDecision.estimated_fresh_duration_ms) ms (advisory only)." -ForegroundColor DarkGray
    Write-Host "Decision: $resolvedExplainPath" -ForegroundColor DarkGray
    exit (Get-RSValidationExitCode -TerminalState NOT_EVALUATED)
}
if ($ExePath) {
    $candidateExePath = if ([IO.Path]::IsPathFullyQualified($ExePath)) {
        $ExePath
    }
    else {
        Join-Path $repoRoot $ExePath
    }
    $exeFullPath = [IO.Path]::GetFullPath($candidateExePath)
    if (-not [string]::Equals(
            $exeFullPath,
            $expectedExeFullPath,
            [StringComparison]::OrdinalIgnoreCase)) {
        Exit-RSInvalidRunnerInvocation (("The -ExePath compatibility parameter must name the selected repository profile's " +
                "canonical executable: $expectedExeFullPath. Refusing uncoordinated executable: $exeFullPath"))
    }
}
else {
    $exeFullPath = $expectedExeFullPath
}

$artifactOperationLock = $null
$testMachineResourceEnvironmentLease = $null
$runnerEnvironmentNames = @(
    'REDSALAMANDER_TEST_ROOT',
    'REDSALAMANDER_TEST_RUN_ID',
    'REDSALAMANDER_SELFTEST_ROOT',
    'REDSALAMANDER_TEST_ALLOW_ALTERNATE_VOLUME',
    'TEMP',
    'TMP'
)
$runnerEnvironmentPrevious = [ordered]@{}
foreach ($environmentName in $runnerEnvironmentNames) {
    $runnerEnvironmentPrevious[$environmentName] = [Environment]::GetEnvironmentVariable($environmentName, 'Process')
}
try {
    $operationMode = if ($SkipBuild) { 'existing artifacts' } else { 'build and test' }
    $artifactOperationScope = @{
        kind = 'test'
        target = $operationMode
        suite = $Suite
        configuration = $Configuration
        platform = $Platform
    }
    $artifactOperationLock = Enter-RSArtifactOperationLock `
        -RepoRoot $repoRoot `
        -Operation "Run-AllTests $Suite ($operationMode) $Configuration|$Platform" `
        -Scope $artifactOperationScope

    if ($artifactOperationLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $repoRoot `
                -Reason "The previous build/test owner exited without clearing the exclusive artifact-operation lock." `
                -AbandonedOwner $artifactOperationLock.AbandonedOwner `
                -Scope $artifactOperationScope)
    }

    if (Test-RSArtifactOperationContaminated -RepoRoot $repoRoot -Scope $artifactOperationScope) {
        $markerPath = Get-RSArtifactContaminationMarkerPath `
            -RepoRoot $repoRoot `
            -Scope $artifactOperationScope
        Write-Host "BLOCKED ATTESTATION: shared build artifacts may be mixed after an interrupted operation. Run build.ps1 -Rebuild before starting tests. Marker: $markerPath" -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
    }

    Assert-RSNoResidualArtifactToolProcesses `
        -RepoRoot $repoRoot `
        -Scope $artifactOperationScope

$testRunContext = New-RSTestRunContext `
    -RepoRoot $repoRoot `
    -RunId $RunId `
    -TestRootOverride $selectedTestRoot `
    -SelfTestRootOverride '' `
    -AllowExternalInitialization:$AllowExternalTestRoot
[void](Initialize-RSTestSandboxRoot `
        -TestRoot $testRunContext.TestRoot `
        -RepoRoot $repoRoot `
        -AllowExternalInitialization:$AllowExternalTestRoot)
$env:REDSALAMANDER_TEST_ROOT = $testRunContext.TestRoot
$env:REDSALAMANDER_TEST_RUN_ID = $testRunContext.RunId
[Environment]::SetEnvironmentVariable('REDSALAMANDER_SELFTEST_ROOT', $null, 'Process')
[Environment]::SetEnvironmentVariable(
    'REDSALAMANDER_TEST_ALLOW_ALTERNATE_VOLUME',
    $(if ($AllowAlternateVolumeTestRoot) { '1' } else { $null }),
    'Process')
$processTempRoot = Join-Path $testRunContext.RunRoot 'scratch\process-temp'
New-Item -ItemType Directory -Path $processTempRoot -Force | Out-Null
$env:TEMP = $processTempRoot
$env:TMP = $processTempRoot
$testMachineResourceEnvironmentLease = Set-RSTestMachineResourceEnvironment -Configuration $testMachineResources
$artifactRoot = $testRunContext.ArtifactRoot
New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null

$staleRunCleanupResults = @(Remove-RSTestSandboxStaleRunDirectories `
        -TestRoot $testRunContext.TestRoot `
        -RunId $testRunContext.RunId)
if (@($staleRunCleanupResults).Count -gt 0) {
    $removedStaleRuns = @($staleRunCleanupResults | Where-Object { $_.Status -eq 'Removed' }).Count
    $failedStaleRuns = @($staleRunCleanupResults | Where-Object { $_.Status -eq 'Failed' }).Count
    Write-Host "Stale TestSandbox run cleanup: $removedStaleRuns removed, $failedStaleRuns failed." -ForegroundColor DarkGray
}

if ([string]::IsNullOrWhiteSpace($QuarantinePath)) {
    $QuarantinePath = Join-Path $repoRoot 'Tools\test-quarantine.jsonl'
}
$quarantineStatus = Read-RSTestQuarantineFile -Path $QuarantinePath

# --- Build step ---

if (-not $SkipBuild) {
    Write-Header "Building $Configuration|$Platform"
    $buildScript = Join-Path $repoRoot 'build.ps1'
    if (-not (Test-Path $buildScript)) {
        Write-Host "ERROR: build.ps1 not found at $buildScript" -ForegroundColor Red
        exit 1
    }

    $buildArgs = Get-RSBuildScriptArguments -Suite $Suite -Configuration $Configuration -Platform $Platform
    if ($BuildNumber -gt 0) { $buildArgs.BuildNumber = $BuildNumber }
    $buildEnvironmentOverrides = Get-RSBuildEnvironmentOverrides -Suite $Suite
    $previousBuildEnvironment = @{}
    foreach ($key in @($buildEnvironmentOverrides.Keys)) {
        $previousBuildEnvironment[$key] = [Environment]::GetEnvironmentVariable($key, 'Process')
        [Environment]::SetEnvironmentVariable($key, [string]$buildEnvironmentOverrides[$key], 'Process')
    }

    $buildExitCode = 1
    try {
        & $buildScript @buildArgs
        $buildExitCode = $LASTEXITCODE
    } finally {
        foreach ($key in @($buildEnvironmentOverrides.Keys)) {
            [Environment]::SetEnvironmentVariable($key, $previousBuildEnvironment[$key], 'Process')
        }
    }

    if ($buildExitCode -ne 0) {
        Write-Host "BUILD FAILED (exit code $buildExitCode)" -ForegroundColor Red
        exit 1
    }
    Write-Host "Build succeeded." -ForegroundColor Green
} else {
    $buildOutputDir = Join-Path $repoRoot ".build\$Platform\$Configuration"
    $receiptSchema = Join-Path $repoRoot 'Specs\Build\BuildReceipt.schema.json'
    $receipt = Read-RSBuildReceipt -BuildOutputDir $buildOutputDir -SchemaPath $receiptSchema
    $sourceSnapshot = Get-RSBuildSourceSnapshot -RepoRoot $repoRoot
    $requiredExecutable = [IO.Path]::GetRelativePath($repoRoot, $exeFullPath).Replace('\', '/')
    if (-not (Test-RSBuildReceiptForArtifactUse -Receipt $receipt -SourceSnapshot $sourceSnapshot `
            -Platform $Platform -Configuration $Configuration -RepoRoot $repoRoot `
            -RequiredRelativePath @($requiredExecutable))) {
        Write-Host ("BLOCKED ATTESTATION: -SkipBuild requires a compatible full-solution build receipt and byte-identical artifacts. " +
                "Run build.ps1 for $Configuration|$Platform, or import a verified same-workflow portable attestation.") -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
    }
    Write-Host "Skipping build (verified receipt $($receipt.receipt_id))." -ForegroundColor DarkGray
}

$buildOutputDir = Join-Path $repoRoot ".build\$Platform\$Configuration"
$receiptSchema = Join-Path $repoRoot 'Specs\Build\BuildReceipt.schema.json'
if (-not $SkipBuild) {
    $receipt = Read-RSBuildReceipt -BuildOutputDir $buildOutputDir -SchemaPath $receiptSchema
}
if ($null -eq $receipt) {
    Write-Host "BLOCKED ATTESTATION: validation requires a schema-valid build receipt for $Configuration|$Platform." -ForegroundColor Red
    exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
}

Import-RSRunnerValidationModules

if (-not (Test-Path $exeFullPath)) {
    Write-Host "ERROR: Executable not found: $exeFullPath" -ForegroundColor Red
    Write-Host "  Build a Debug configuration first, or use -ExePath to specify the path." -ForegroundColor Yellow
    exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
}

# --- Determine test plan ---

$testPlan = @(Get-RSTestRunPlan `
        -Suite $Suite `
        -RepoRoot $repoRoot `
        -Platform $Platform `
        -Configuration $Configuration `
        -RedSalamanderExePath $exeFullPath `
        -TimeoutMultiplier $TimeoutMultiplier `
        -FailFast:$FailFast `
        -CaseFilter $CaseFilter `
        -CommandsFamily $CommandsFamily `
        -RepeatCount $SelfTestRepeat `
        -ShuffleSeed $SelfTestShuffleSeed `
        -PerfBudgetPath $PerfBudgetPath `
        -RequirePerfBudgets:$RequirePerfBudgets `
        -SelfTestFlakyProofCase $SelfTestFlakyProofCase `
        -SelfTestOrderProofCase $SelfTestOrderProofCase)
$requiredPlanArtifactPaths = @($testPlan | ForEach-Object {
        $candidate = [IO.Path]::GetFullPath([string]$_.Path)
        $outputRoot = [IO.Path]::GetFullPath($buildOutputDir).TrimEnd('\')
        if ($candidate.StartsWith($outputRoot + '\', [StringComparison]::OrdinalIgnoreCase) -and
            (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            [IO.Path]::GetRelativePath($repoRoot, $candidate).Replace('\', '/')
        }
    } | Sort-Object -Unique)
if ($Suite -in @('CI', 'Full') -and
    -not (Test-RSBuildReceiptForArtifactUse -Receipt $receipt -SourceSnapshot (Get-RSBuildSourceSnapshot -RepoRoot $repoRoot) `
        -Platform $Platform -Configuration $Configuration -RepoRoot $repoRoot `
        -RequiredRelativePath $requiredPlanArtifactPaths)) {
    Write-Host 'BLOCKED ATTESTATION: the verified solution receipt does not contain the complete selected executable/plugin test plan closure.' -ForegroundColor Red
    exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
}
$validationCoverageScope = if (-not [string]::IsNullOrWhiteSpace($CaseFilter)) {
    'filtered'
} elseif (-not [string]::IsNullOrWhiteSpace($CommandsFamily) -or $Suite -in @('Compare', 'Commands', 'FileOps')) {
    'family'
} else {
    'complete'
}
$validationPlan = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite $Suite `
    -Entries $testPlan -ValidationMode $ValidationMode -CoverageScope $validationCoverageScope
$validationPlanPath = Join-Path $artifactRoot 'validation-plan.json'
Write-RSAtomicCanonicalJsonFile -LiteralPath $validationPlanPath -InputObject $validationPlan `
    -SchemaPath (Join-Path $repoRoot 'Specs\Testing\ValidationPlan.schema.json') -AllowedRoot $artifactRoot
Write-Host "Validation plan: $($validationPlan.plan_digest)" -ForegroundColor DarkGray
$workspaceSnapshotSchema = Join-Path $repoRoot 'Specs\Testing\WorkspaceSnapshot.schema.json'
function Publish-RSRunnerWorkspaceSnapshot {
    param([Parameter(Mandatory = $true)][string]$Label)

    $snapshot = Get-RSWorkspaceSnapshot -RepoRoot $repoRoot -ImpactBase $ImpactBase
    $path = Join-Path $artifactRoot "workspace-$Label.json"
    Write-RSAtomicCanonicalJsonFile -LiteralPath $path -InputObject $snapshot `
        -SchemaPath $workspaceSnapshotSchema -AllowedRoot $artifactRoot
    return $snapshot
}
$initialWorkspaceSnapshot = Publish-RSRunnerWorkspaceSnapshot -Label 'initial'
$workspaceSnapshotRows = @([pscustomobject]@{
        entry_id = $null; phase = 'initial'; snapshot_id = $initialWorkspaceSnapshot.snapshot_id
        computation_ms = [int64]$initialWorkspaceSnapshot.computation_ms
    })
$sourceMutationDetected = $false

$suiteConfigs = @($testPlan | Where-Object { $_.Kind -eq 'SelfTest' })
$standaloneConfigs = @($testPlan | Where-Object { $_.Kind -ne 'SelfTest' })
$failureClassificationEnabled = ($ClassifyFailures -or $Suite -eq 'CI')
$runnerDependencyPaths = @(Get-RSValidationDependencyClosure -RepoRoot $repoRoot `
        -RootRelativePath 'Tools/Run-AllTests.ps1')
$runnerDependencyDigests = @{}
foreach ($relativePath in $runnerDependencyPaths) {
    $dependencyPath = Join-Path $repoRoot $relativePath
    if (-not (Test-Path -LiteralPath $dependencyPath -PathType Leaf)) {
        throw "Validation runner dependency is missing: $relativePath"
    }
    $runnerDependencyDigests[$relativePath.ToLowerInvariant()] = `
        (Get-FileHash -LiteralPath $dependencyPath -Algorithm SHA256).Hash.ToLowerInvariant()
}
$quarantineDigest = if (Test-Path -LiteralPath $QuarantinePath -PathType Leaf) {
    (Get-FileHash -LiteralPath $QuarantinePath -Algorithm SHA256).Hash.ToLowerInvariant()
} else {
    Get-RSCanonicalJsonDigest -InputObject ([ordered]@{ state = 'absent'; path = 'Tools/test-quarantine.jsonl' })
}
$artifactEpoch = 0L
$resumeOriginState = $null
$resumeOriginDirectory = $null
if ($ValidationMode -eq 'Resume') {
    # Fingerprints bind the artifact epoch. Resolve and validate the origin before
    # creating current fingerprints so an exact Resume can address promotions
    # published after the origin writer established its terminal epoch.
    $evidenceRoot = [IO.Path]::GetFullPath($validationEvidenceRoot).TrimEnd('\')
    try {
        $resolvedOrigin = Resolve-RSValidationRunDirectory -EvidenceRoot $evidenceRoot `
            -RunReference $ResumeFrom -SchemaRoot (Join-Path $repoRoot 'Specs\Testing')
        $resumeOriginDirectory = $resolvedOrigin.RunDirectory
        $resumeOriginState = $resolvedOrigin.State
    } catch {
        Write-Host "CORRUPT EVIDENCE: $($_.Exception.Message)" -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState CORRUPT_EVIDENCE)
    }
    $artifactEpoch = [int64]$resumeOriginState.artifact_epoch
}
$buildReceiptBindingDigest = Get-RSValidationBuildReceiptBindingDigest -BuildReceipt $receipt
function New-RSRunnerEntryFingerprintRecord {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$ExecutionEntries,
        [Parameter(Mandatory = $true)][object]$PortablePlan
    )

    $timer = [Diagnostics.Stopwatch]::StartNew()
    $rows = @()
    for ($entryOrdinal = 0; $entryOrdinal -lt $ExecutionEntries.Count; $entryOrdinal++) {
        $executionEntry = $ExecutionEntries[$entryOrdinal]
        $portableEntry = @($PortablePlan.entries | Where-Object id -eq $executionEntry.Id)
        if ($portableEntry.Count -ne 1) {
            throw "Validation plan identity mismatch for entry '$($executionEntry.Id)'."
        }
        $inputSetDigests = @{}
        foreach ($inputSet in @($portableEntry[0].input_sets)) {
            $inputSetDigests[[string]$inputSet] = $initialWorkspaceSnapshot.snapshot_id
        }
        $capabilityInputs = @{}
        foreach ($environmentInput in @($portableEntry[0].environment_inputs)) {
            switch ([string]$environmentInput) {
                'env.profile' { $capabilityInputs[$environmentInput] = "$Platform|$Configuration" }
                'env.windows-11' { $capabilityInputs[$environmentInput] = [Environment]::OSVersion.Version.ToString() }
                'env.interactive-desktop' { $capabilityInputs[$environmentInput] = [Environment]::UserInteractive }
                'env.network' { $capabilityInputs[$environmentInput] = [Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable() }
                'env.test-resources' { $capabilityInputs[$environmentInput] = $testMachineResources.CapabilityDigest }
                default { throw "Validation entry '$($executionEntry.Id)' declares an unknown environment input '$environmentInput'." }
            }
        }
        $fingerprint = Get-RSValidationEntryFingerprint -EntryContract $portableEntry[0] `
            -CoveragePlanDigest $PortablePlan.coverage_plan_digest -WorkspaceSnapshot $initialWorkspaceSnapshot `
            -BuildReceipt $receipt -BuildReceiptBindingDigest $buildReceiptBindingDigest `
            -InputSetDigests $inputSetDigests `
            -RunnerDependencyDigests $runnerDependencyDigests -CapabilityInputs $capabilityInputs `
            -QuarantineDigest $quarantineDigest `
            -ClassificationMode $(if ($failureClassificationEnabled) { 'enabled' } else { 'disabled' }) `
            -ArtifactEpoch $artifactEpoch
        $executionEntry | Add-Member -NotePropertyName EntryFingerprint -NotePropertyValue $fingerprint.Fingerprint -Force
        $executionEntry | Add-Member -NotePropertyName EntryFingerprintComponentVersion `
            -NotePropertyValue $fingerprint.ComponentDigestVersion -Force
        $executionEntry | Add-Member -NotePropertyName EntryFingerprintComponents `
            -NotePropertyValue $fingerprint.ComponentDigests -Force
        $executionEntry | Add-Member -NotePropertyName ArtifactEpoch -NotePropertyValue $artifactEpoch -Force
        $rows += [pscustomobject][ordered]@{
            id = $executionEntry.Id
            fingerprint = $fingerprint.Fingerprint
            component_digest_version = $fingerprint.ComponentDigestVersion
            component_digests = $fingerprint.ComponentDigests
            computation_ms = [int64]$fingerprint.ComputationMs
        }
    }
    $timer.Stop()
    $record = [ordered]@{
        schema = 'red-salamander.validation-entry-fingerprints.v1'
        canonicalization = 'rs-canonical-json-v1'
        fingerprints_digest = '0' * 64
        plan_digest = $PortablePlan.plan_digest
        coverage_plan_digest = $PortablePlan.coverage_plan_digest
        snapshot_id = $initialWorkspaceSnapshot.snapshot_id
        build_receipt_id = $receipt.receipt_id
        artifact_epoch = $artifactEpoch
        entries = $rows
        computation_ms = [int64][Math]::Max(0, [Math]::Round($timer.Elapsed.TotalMilliseconds))
    }
    $record.fingerprints_digest = Get-RSCanonicalJsonDigest -InputObject ([ordered]@{
            schema = $record.schema; canonicalization = $record.canonicalization
            plan_digest = $record.plan_digest; coverage_plan_digest = $record.coverage_plan_digest
            snapshot_id = $record.snapshot_id
            build_receipt_id = $record.build_receipt_id; artifact_epoch = $record.artifact_epoch
            entries = $record.entries
        })
    return [pscustomobject]$record
}

$entryFingerprintRecord = New-RSRunnerEntryFingerprintRecord -ExecutionEntries $testPlan -PortablePlan $validationPlan
$entryFingerprintRows = @($entryFingerprintRecord.entries)
$impactManifest = Read-RSValidationImpactManifest `
    -Path (Join-Path $repoRoot 'Tools\validation-impact.json') `
    -SchemaPath (Join-Path $repoRoot 'Specs\Testing\ValidationImpact.schema.json') `
    -RepoRoot $repoRoot
$trackedGraphInputs = @(& git -C $repoRoot ls-files '*.sln' '*.slnx' '*.vcxproj' '*.props' '*.targets' 'RuntimeDependencies.props')
$impactGovernance = Test-RSValidationImpactGovernance -Manifest $impactManifest -TrackedPaths $trackedGraphInputs -RepoRoot $repoRoot
if (-not $impactGovernance.IsValid) {
    Exit-RSInvalidRunnerInvocation "Validation impact governance is incomplete: $(@($impactGovernance.UnmappedBuildInputs) -join ', ')"
}
$durationEstimates = Get-RSValidationDurationEstimates `
    -EvidenceRoot $validationEvidenceRoot `
    -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') `
    -EntryId @($validationPlan.entries | ForEach-Object { [string]$_.id })
$fingerprintsById = @{}
foreach ($fingerprintRow in $entryFingerprintRows) { $fingerprintsById[$fingerprintRow.id] = $fingerprintRow.fingerprint }
$fingerprintComponentsById = @{}
foreach ($fingerprintRow in $entryFingerprintRows) {
    $fingerprintComponentsById[$fingerprintRow.id] = [pscustomobject]@{
        version = $fingerprintRow.component_digest_version
        digests = $fingerprintRow.component_digests
    }
}
$shadowDecision = Get-RSValidationImpactDecision -Manifest $impactManifest `
    -WorkspaceSnapshot $initialWorkspaceSnapshot -ValidationPlan $validationPlan `
    -EntryFingerprints $fingerprintsById -EstimatedDurationMs $durationEstimates -PlannerMode shadow
if ($ValidationMode -eq 'Affected') {
    $affectedEntryIds = @($shadowDecision.entries | Where-Object affected | ForEach-Object { [string]$_.id })
    $testPlan = @($testPlan | Where-Object { $_.Id -in $affectedEntryIds })
    $validationCoverageScope = 'filtered'
    $validationPlan = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite $Suite `
        -Entries $testPlan -ValidationMode Affected -CoverageScope $validationCoverageScope
    Write-RSAtomicCanonicalJsonFile -LiteralPath $validationPlanPath -InputObject $validationPlan `
        -SchemaPath (Join-Path $repoRoot 'Specs\Testing\ValidationPlan.schema.json') -AllowedRoot $artifactRoot
    $entryFingerprintRecord = New-RSRunnerEntryFingerprintRecord -ExecutionEntries $testPlan -PortablePlan $validationPlan
    $entryFingerprintRows = @($entryFingerprintRecord.entries)
    $fingerprintsById = @{}
    foreach ($fingerprintRow in $entryFingerprintRows) { $fingerprintsById[$fingerprintRow.id] = $fingerprintRow.fingerprint }
    $fingerprintComponentsById = @{}
    foreach ($fingerprintRow in $entryFingerprintRows) {
        $fingerprintComponentsById[$fingerprintRow.id] = [pscustomobject]@{
            version = $fingerprintRow.component_digest_version
            digests = $fingerprintRow.component_digests
        }
    }
    $shadowDecision = Get-RSValidationImpactDecision -Manifest $impactManifest `
        -WorkspaceSnapshot $initialWorkspaceSnapshot -ValidationPlan $validationPlan `
        -EntryFingerprints $fingerprintsById -EstimatedDurationMs $durationEstimates -PlannerMode enforced
}
Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $artifactRoot 'entry-fingerprints.json') `
    -InputObject $entryFingerprintRecord `
    -SchemaPath (Join-Path $repoRoot 'Specs\Testing\ValidationEntryFingerprints.schema.json') `
    -AllowedRoot $artifactRoot
Write-Host "Entry fingerprints: $($entryFingerprintRecord.fingerprints_digest) ($($entryFingerprintRecord.computation_ms) ms)" -ForegroundColor DarkGray
Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $artifactRoot 'validation-plan-decision.json') `
    -InputObject $shadowDecision `
    -SchemaPath (Join-Path $repoRoot 'Specs\Testing\ValidationPlanDecision.schema.json') `
    -AllowedRoot $artifactRoot
if ($ValidationMode -eq 'Affected') {
    Write-Host "Affected execution set: $(@($shadowDecision.entries).Count) entries; advisory-build=$($shadowDecision.advisory_build_action). Repository verdict remains NOT_EVALUATED." -ForegroundColor Cyan
} else {
    Write-Host "Shadow affected set: $(@($shadowDecision.entries | Where-Object affected).Count)/$(@($shadowDecision.entries).Count) entries; advisory-build=$($shadowDecision.advisory_build_action). Fresh still executes all." -ForegroundColor DarkGray
}
$suiteConfigs = @($testPlan | Where-Object { $_.Kind -eq 'SelfTest' })
$standaloneConfigs = @($testPlan | Where-Object { $_.Kind -ne 'SelfTest' })
if ($ValidationMode -eq 'Resume') {
    try {
        $resumeCandidates = @(Get-RSValidationResumeCandidates -OriginRunDirectory $resumeOriginDirectory `
                -CurrentPlan $validationPlan -CurrentDecision $shadowDecision `
                -CurrentSnapshot $initialWorkspaceSnapshot -BuildReceipt $receipt `
                -CurrentFingerprintComponentsById $fingerprintComponentsById `
                -SchemaRoot (Join-Path $repoRoot 'Specs\Testing'))
    } catch {
        Write-Host "CORRUPT EVIDENCE: $($_.Exception.Message)" -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState CORRUPT_EVIDENCE)
    }
    if (@($resumeCandidates | Where-Object reason_code -eq 'CORRUPT_EVIDENCE').Count -gt 0) {
        Write-Host 'CORRUPT EVIDENCE: Resume origin contains an invalid promotion or evidence record.' -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState CORRUPT_EVIDENCE)
    }
    $shadowDecision = New-RSValidationResumeDecision -FreshDecision $shadowDecision -Candidate $resumeCandidates
    Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $artifactRoot 'validation-plan-decision.json') `
        -InputObject $shadowDecision `
        -SchemaPath (Join-Path $repoRoot 'Specs\Testing\ValidationPlanDecision.schema.json') `
        -AllowedRoot $artifactRoot
    Write-Host "Exact resume: $(@($resumeCandidates | Where-Object reusable).Count)/$($resumeCandidates.Count) entries reusable from $($resumeOriginState.run_id)." -ForegroundColor Cyan
}
$validationEvidenceRun = $null
if ($Suite -in @('CI', 'Full')) {
    $validationEvidenceRun = New-RSValidationEvidenceRun `
        -EvidenceRoot $validationEvidenceRoot `
        -RunId $testRunContext.RunId -Plan $validationPlan -Decision $shadowDecision `
        -BuildReceipt $receipt -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') `
        -ArtifactEpoch $artifactEpoch -ValidationMode $ValidationMode `
        -CampaignId $(if ($null -eq $resumeOriginState) { $null } else { $resumeOriginState.campaign_id }) `
        -CampaignSequence $(if ($null -eq $resumeOriginState) { 1 } else { [int64]$resumeOriginState.campaign_sequence + 1L }) `
        -ParentRunId $(if ($null -eq $resumeOriginState) { $null } else { $resumeOriginState.run_id })
    Write-Host "Validation checkpoint: $($validationEvidenceRun.RunDirectory)" -ForegroundColor DarkGray
}
$reusedEntryIds = @($shadowDecision.entries | Where-Object decision -eq 'REUSE' | ForEach-Object { [string]$_.id })
if ($reusedEntryIds.Count -gt 0) {
    $suiteConfigs = @($suiteConfigs | Where-Object { $_.Id -notin $reusedEntryIds })
    $standaloneConfigs = @($standaloneConfigs | Where-Object { $_.Id -notin $reusedEntryIds })
}
function Publish-RSRunnerEntryEvidence {
    param(
        [Parameter(Mandatory = $true)][object]$Entry,
        [Parameter(Mandatory = $true)][object]$Result,
        [Parameter(Mandatory = $true)][datetime]$StartedUtc,
        [Parameter(Mandatory = $true)][object]$PostSnapshot,
        [string[]]$CandidateArtifact = @(),
        [switch]$RequireArtifact
    )

    $portableEntry = @($validationPlan.entries | Where-Object id -eq $Entry.Id)
    if ($portableEntry.Count -ne 1) { throw "Evidence entry contract is missing for '$($Entry.Id)'." }
    $candidateFiles = @($CandidateArtifact | Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_) -and (Test-Path -LiteralPath $_ -PathType Leaf)
        })
    $preserved = if ($null -eq $validationEvidenceRun) { @() } else {
        @(Publish-RSValidationResultArtifacts -Run $validationEvidenceRun `
                -EntryId $Entry.Id -CandidatePath $candidateFiles)
    }
    $requiredArtifactsPresent = -not $RequireArtifact -or $(if ($null -eq $validationEvidenceRun) {
            $candidateFiles.Count -gt 0
        } else {
            $preserved.Count -gt 0
        })
    $resultParsed = [bool](Get-RSObjectValue -Object $Result -Names @('ResultParsed', 'result_parsed') -DefaultValue $false)
    $coverageValid = [bool](Get-RSObjectValue -Object $Result -Names @('CoverageValid', 'coverage_valid') -DefaultValue $true)
    $terminalResult = New-RSValidationTerminalResult -EntryContract $portableEntry[0] -Result $Result `
        -ResultParsed $resultParsed -CoverageValid $coverageValid `
        -RequiredArtifactsPresent $requiredArtifactsPresent
    if ($terminalResult.status -ne 'PASSED') {
        if ($Result -is [Collections.IDictionary]) {
            $Result['ExitCode'] = 1
            $Result['Failed'] = [Math]::Max(1, [int](Get-RSObjectValue -Object $Result -Names @('Failed', 'failed') -DefaultValue 0))
        } else {
            $Result | Add-Member -NotePropertyName ExitCode -NotePropertyValue 1 -Force
            $Result | Add-Member -NotePropertyName Failed `
                -NotePropertyValue ([Math]::Max(1, [int](Get-RSObjectValue -Object $Result -Names @('Failed', 'failed') -DefaultValue 0))) -Force
        }
    }
    if ($null -eq $validationEvidenceRun) {
        return [pscustomobject]@{ TerminalResult = $terminalResult; PublishedEvidence = $null }
    }
    $required = if ($RequireArtifact -and $preserved.Count -eq 0) {
        @("entries/$($Entry.Id)/results/required-terminal-artifact.missing")
    } elseif ($RequireArtifact) { @($preserved.path) } else { @() }
    $published = Publish-RSValidationEntryEvidence -Run $validationEvidenceRun `
        -EntryContract $portableEntry[0] -EntryFingerprint $Entry.EntryFingerprint `
        -EntryFingerprintComponentVersion $Entry.EntryFingerprintComponentVersion `
        -EntryFingerprintComponents $Entry.EntryFingerprintComponents `
        -TerminalResult $terminalResult `
        -PreSnapshot $initialWorkspaceSnapshot -PostSnapshot $PostSnapshot `
        -BuildReceipt $receipt -StartedUtc $StartedUtc -EndedUtc (Get-Date) `
        -ResultArtifact $preserved -RequiredArtifact $required
    return [pscustomobject]@{ TerminalResult = $terminalResult; PublishedEvidence = $published }
}
$quarantineHarnessCaseMap = @{}
$activeQuarantineEntries = @(Get-RSObjectValue -Object $quarantineStatus -Names @('ActiveEntries', 'active_entries') -DefaultValue @())
$activeQuarantineHarnesses = @($activeQuarantineEntries | ForEach-Object {
        [string](Get-RSObjectValue -Object $_ -Names @('harness') -DefaultValue '')
    } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
foreach ($selfTestEntry in @($suiteConfigs | Where-Object { [string]$_.Name -in $activeQuarantineHarnesses })) {
    try {
        $caseList = Invoke-RSSelfTestCaseList -Entry $selfTestEntry
        $suiteCases = @($caseList.suites | Where-Object { $_.suite -eq $selfTestEntry.Name } | Select-Object -First 1)
        if (@($suiteCases).Count -eq 0) {
            $quarantineHarnessCaseMap[$selfTestEntry.Name] = @()
        } else {
            $quarantineHarnessCaseMap[$selfTestEntry.Name] = @($suiteCases[0].cases | ForEach-Object { [string]$_.name })
        }
    } catch {
        $quarantineHarnessCaseMap[$selfTestEntry.Name] = @()
    }
}
$quarantineRepairPlan = Get-RSTestQuarantineRepairPlan -QuarantineStatus $quarantineStatus -TestPlan $testPlan -HarnessCaseMap $quarantineHarnessCaseMap
if (-not $quarantineRepairPlan.IsValid) {
    $quarantineStatus = [pscustomobject]@{
        IsValid = $false
        HasBlockingEntries = $true
        TotalCount = [int](Get-RSObjectValue -Object $quarantineStatus -Names @('TotalCount', 'total_count') -DefaultValue 0)
        ActiveEntries = @(Get-RSObjectValue -Object $quarantineStatus -Names @('ActiveEntries', 'active_entries') -DefaultValue @())
        InvalidEntries = @($quarantineRepairPlan.InvalidEntries)
        Errors = @($quarantineRepairPlan.Errors)
    }
}

# --- Run suites ---

$allResults = @()
$quarantineRepairAttempts = @()
$overallStartTime = Get-Date
$overallExitCode = 0

function Assert-RSRunnerArtifactReadBarrier {
    param([Parameter(Mandatory = $true)][object]$Entry)
    if ($null -eq $validationEvidenceRun -or @($Entry.ArtifactReadIds).Count -eq 0 -or
        @($Entry.ArtifactWriteIds).Count -gt 0) {
        return
    }
    if (-not (Test-RSValidationArtifactReadBarrier -Run $validationEvidenceRun `
            -BuildReceipt $receipt -ExpectedEpoch $artifactEpoch)) {
        Write-Host "BLOCKED ATTESTATION: artifact reader '$($Entry.Id)' cannot enter an unstable or mismatched epoch." -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
    }
}

function Invoke-RSRunnerStandaloneValidationEntry {
    param([Parameter(Mandatory = $true)][object]$Entry, [switch]$DeferEvidence)

    Write-Header "Running: $($Entry.Name)"
    if ($null -ne $validationEvidenceRun) {
        Set-RSValidationEntryCheckpoint -Run $validationEvidenceRun -EntryId $Entry.Id -Status running
    }
    Write-Host "  Kind: $($Entry.Kind)" -ForegroundColor DarkGray
    Write-Host "  Path: $($Entry.Path)" -ForegroundColor DarkGray
    if (@($Entry.Arguments).Count -gt 0) { Write-Host "  Args: $(@($Entry.Arguments) -join ' ')" -ForegroundColor DarkGray }
    Write-Host ""

    $startedUtc = Get-Date
    Assert-RSRunnerArtifactReadBarrier -Entry $Entry
    $profileEnvironment = [ordered]@{}
    if (@($Entry.ArtifactWriteIds).Count -gt 0) {
        $profileEnvironment['RS_VALIDATION_PLATFORM'] = $Platform
        $profileEnvironment['RS_VALIDATION_CONFIGURATION'] = $Configuration
    }
    $previousProfileEnvironment = @{}
    try {
        foreach ($key in $profileEnvironment.Keys) {
            $previousProfileEnvironment[$key] = [Environment]::GetEnvironmentVariable($key, 'Process')
            [Environment]::SetEnvironmentVariable($key, [string]$profileEnvironment[$key], 'Process')
        }
        $result = Invoke-RSTestPlanEntry -Entry $Entry -RepoRoot $repoRoot -ArtifactRoot $artifactRoot
        if ($failureClassificationEnabled -and $result.ExitCode -ne 0) {
            Write-SubHeader "CLASSIFYING FAILURE: $($Entry.Name)"
            $retryAttempts = @(Invoke-RSTestEntryClassificationRetry -Entry $Entry -RepoRoot $repoRoot -ArtifactRoot $artifactRoot)
            $result = Add-RSTestResultClassification -Result $result -RetryResults $retryAttempts
            Write-Host "  Classification: $($result.Classification)" -ForegroundColor Yellow
            Write-Host "  Reason: $($result.ClassificationReason)" -ForegroundColor DarkGray
        } else {
            $result = Add-RSTestResultClassification -Result $result
        }
    } finally {
        foreach ($key in $profileEnvironment.Keys) {
            [Environment]::SetEnvironmentVariable($key, $previousProfileEnvironment[$key], 'Process')
        }
    }

    Import-RSRunnerValidationModules
    $postSnapshot = Publish-RSRunnerWorkspaceSnapshot -Label "after-$($Entry.Id)"
    $result | Add-Member -NotePropertyName WorkspaceSnapshotId -NotePropertyValue $postSnapshot.snapshot_id -Force
    $logPath = [string](Get-RSObjectValue -Object $result -Names @('OutputLogPath', 'output_log_path') -DefaultValue '')
    if (-not $DeferEvidence) {
        [void](Publish-RSRunnerEntryEvidence -Entry $Entry -Result $result -StartedUtc $startedUtc `
            -PostSnapshot $postSnapshot -CandidateArtifact @($logPath) `
            -RequireArtifact:(-not [string]::IsNullOrWhiteSpace($logPath)))
    }
    return [pscustomobject]@{
        Entry = $Entry; Result = $result; StartedUtc = $startedUtc; PostSnapshot = $postSnapshot
        LogPath = $logPath; SourceMutation = ($postSnapshot.snapshot_id -ne $initialWorkspaceSnapshot.snapshot_id)
    }
}

function Add-RSRunnerStandaloneExecutionResult {
    param([Parameter(Mandatory = $true)][object]$Execution)
    $script:workspaceSnapshotRows += [pscustomobject]@{
        entry_id = $Execution.Entry.Id; phase = 'post-entry'; snapshot_id = $Execution.PostSnapshot.snapshot_id
        computation_ms = [int64]$Execution.PostSnapshot.computation_ms
    }
    if ($Execution.Result.ExitCode -ne 0) { $script:overallExitCode = 1 }
    if ($Execution.SourceMutation) {
        $script:sourceMutationDetected = $true
        $script:overallExitCode = 1
        Write-Host "  SOURCE MUTATION: workspace changed during $($Execution.Entry.Id)." -ForegroundColor Red
    }
    $script:allResults += $Execution.Result
}

$artifactMutatorConfigs = @($standaloneConfigs | Where-Object { @($_.ArtifactWriteIds).Count -gt 0 })
$standaloneConfigs = @($standaloneConfigs | Where-Object { @($_.ArtifactWriteIds).Count -eq 0 })
foreach ($artifactMutator in $artifactMutatorConfigs) {
    $preMutationReceipt = $receipt
    if (-not (Test-RSValidationArtifactReadBarrier -Run $validationEvidenceRun `
            -BuildReceipt $preMutationReceipt -ExpectedEpoch $artifactEpoch)) {
        Write-Host "BLOCKED ATTESTATION: artifact writer '$($artifactMutator.Id)' cannot open from an unstable epoch." -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
    }
    # Fail closed before the writer can delete or replace any selected-profile byte.
    # A crash after this point leaves no usable receipt and a durable mutating epoch.
    Suspend-RSBuildReceipt -BuildOutputDir $buildOutputDir
    $receipt = $null
    $artifactEpoch = Start-RSValidationArtifactMutation -Run $validationEvidenceRun `
        -Plan $validationPlan -WrittenArtifactId @($artifactMutator.ArtifactWriteIds)
    $mutationExecution = Invoke-RSRunnerStandaloneValidationEntry -Entry $artifactMutator -DeferEvidence
    if ($mutationExecution.Result.ExitCode -ne 0) { $overallExitCode = 1 }

    Write-Header "Re-attesting profile after: $($artifactMutator.Name)"
    $buildNumberArgument = @($preMutationReceipt.build_arguments | Where-Object { $_ -match '^build-number=([0-9]+)$' } | Select-Object -First 1)
    if ($buildNumberArgument.Count -ne 1) {
        Write-Host 'BLOCKED ATTESTATION: the pre-mutation receipt does not bind one build number.' -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
    }
    $reattestBuildNumber = [int]($buildNumberArgument[0] -replace '^build-number=', '')
    $reattestEnvironmentOverrides = Get-RSBuildEnvironmentOverrides -Suite $Suite
    $previousReattestEnvironment = @{}
    foreach ($key in @($reattestEnvironmentOverrides.Keys)) {
        $previousReattestEnvironment[$key] = [Environment]::GetEnvironmentVariable($key, 'Process')
        [Environment]::SetEnvironmentVariable($key, [string]$reattestEnvironmentOverrides[$key], 'Process')
    }

    $reattestExitCode = 1
    try {
        & (Join-Path $repoRoot 'build.ps1') -Configuration $Configuration -Platform $Platform `
            -BuildNumber $reattestBuildNumber -MaxCpuCount 4
        $reattestExitCode = $LASTEXITCODE
    } finally {
        foreach ($key in @($reattestEnvironmentOverrides.Keys)) {
            [Environment]::SetEnvironmentVariable($key, $previousReattestEnvironment[$key], 'Process')
        }
    }
    if ($reattestExitCode -ne 0) {
        Write-Host "BLOCKED ATTESTATION: profile re-attestation failed after '$($artifactMutator.Id)'." -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
    }
    Import-RSRunnerValidationModules
    $restoredReceipt = Read-RSBuildReceipt -BuildOutputDir $buildOutputDir -SchemaPath $receiptSchema
    $restoredSnapshot = Get-RSBuildSourceSnapshot -RepoRoot $repoRoot
    if ($null -eq $restoredReceipt -or
        -not (Test-RSBuildReceiptForArtifactUse -Receipt $restoredReceipt -SourceSnapshot $restoredSnapshot `
            -Platform $Platform -Configuration $Configuration -RepoRoot $repoRoot `
            -RequiredRelativePath $requiredPlanArtifactPaths)) {
        Write-Host "BLOCKED ATTESTATION: the re-attested profile failed byte-for-byte receipt verification after '$($artifactMutator.Id)'." -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
    }
    $receipt = $restoredReceipt

    # Native PE/PDB identities can change across an equivalent rebuild. Bind all
    # evidence to the verified post-mutation receipt/epoch before publishing this
    # writer or launching any artifact reader. Any Resume reuse selected against
    # the previous receipt is conservatively converted back to execution.
    $entryFingerprintRecord = New-RSRunnerEntryFingerprintRecord -ExecutionEntries $testPlan -PortablePlan $validationPlan
    $entryFingerprintRows = @($entryFingerprintRecord.entries)
    $fingerprintsById = @{}
    foreach ($fingerprintRow in $entryFingerprintRows) { $fingerprintsById[$fingerprintRow.id] = $fingerprintRow.fingerprint }
    $shadowDecision = Get-RSValidationImpactDecision -Manifest $impactManifest `
        -WorkspaceSnapshot $initialWorkspaceSnapshot -ValidationPlan $validationPlan `
        -EntryFingerprints $fingerprintsById -EstimatedDurationMs $durationEstimates `
        -PlannerMode $(if ($ValidationMode -eq 'Affected') { 'enforced' } else { 'shadow' })
    Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $artifactRoot 'entry-fingerprints.json') `
        -InputObject $entryFingerprintRecord `
        -SchemaPath (Join-Path $repoRoot 'Specs\Testing\ValidationEntryFingerprints.schema.json') `
        -AllowedRoot $artifactRoot
    Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $artifactRoot 'validation-plan-decision.json') `
        -InputObject $shadowDecision `
        -SchemaPath (Join-Path $repoRoot 'Specs\Testing\ValidationPlanDecision.schema.json') `
        -AllowedRoot $artifactRoot
    $artifactEpoch = Complete-RSValidationArtifactMutation -Run $validationEvidenceRun `
        -Decision $shadowDecision -BuildReceipt $receipt -MutatorEntryId $artifactMutator.Id `
        -ExpectedEpoch $artifactEpoch
    $buildReceiptBindingDigest = Get-RSValidationBuildReceiptBindingDigest -BuildReceipt $receipt
    if (-not (Test-RSValidationArtifactReadBarrier -Run $validationEvidenceRun `
            -BuildReceipt $receipt -ExpectedEpoch $artifactEpoch)) {
        Write-Host "BLOCKED ATTESTATION: re-attested artifact epoch did not become readable." -ForegroundColor Red
        exit (Get-RSValidationExitCode -TerminalState BLOCKED_ATTESTATION)
    }
    $suiteConfigs = @($testPlan | Where-Object { $_.Kind -eq 'SelfTest' })
    $standaloneConfigs = @($testPlan | Where-Object { $_.Kind -ne 'SelfTest' -and @($_.ArtifactWriteIds).Count -eq 0 })
    [void](Publish-RSRunnerEntryEvidence -Entry $artifactMutator -Result $mutationExecution.Result `
        -StartedUtc $mutationExecution.StartedUtc -PostSnapshot $mutationExecution.PostSnapshot `
        -CandidateArtifact @($mutationExecution.LogPath) `
        -RequireArtifact:(-not [string]::IsNullOrWhiteSpace($mutationExecution.LogPath)))
    Add-RSRunnerStandaloneExecutionResult -Execution $mutationExecution
}

$lateResourceClasses = @('interactive-desktop', 'performance-exclusive', 'stateful', 'network')
$earlyStandaloneConfigs = @($standaloneConfigs | Where-Object {
        @($_.ResourceClasses | Where-Object { $_ -in $lateResourceClasses }).Count -eq 0
    } | Sort-Object `
        @{ Expression = { $estimate = [int64]$durationEstimates[$_.Id]; if ($estimate -gt 0) { $estimate } else { [int64]::MaxValue } } }, `
        @{ Expression = { $_.Id } })
$standaloneConfigs = @($standaloneConfigs | Where-Object { $_.Id -notin @($earlyStandaloneConfigs.Id) })
if ($earlyStandaloneConfigs.Count -gt 0) {
    Write-Host "Scheduler: running $($earlyStandaloneConfigs.Count) cheap noninteractive entries before broad interactive/stateful work." -ForegroundColor DarkGray
}
foreach ($entry in $earlyStandaloneConfigs) {
    $execution = Invoke-RSRunnerStandaloneValidationEntry -Entry $entry
    Add-RSRunnerStandaloneExecutionResult -Execution $execution
}

foreach ($sc in $suiteConfigs) {
    Write-Header "Running: $($sc.Name)"
    Assert-RSRunnerArtifactReadBarrier -Entry $sc

    if ($null -ne $validationEvidenceRun) {
        Set-RSValidationEntryCheckpoint -Run $validationEvidenceRun -EntryId $sc.Id -Status running
    }

    $args = @($sc.Arguments)

    $suiteStart = Get-Date
    Write-Host "  Exe:  $($sc.Path)" -ForegroundColor DarkGray
    Write-Host "  Args: $($args -join ' ')" -ForegroundColor DarkGray
    Write-Host ""

    $jsonCandidates = @(
        (Join-Path $artifactRoot "$($sc.JsonName)\results.json"),
        (Join-Path $artifactRoot "$($sc.JsonName)_results.json"),
        (Join-Path $artifactRoot 'results.json')
    )
    foreach ($jsonPath in $jsonCandidates) {
        if (Test-Path $jsonPath) {
            Remove-Item -LiteralPath $jsonPath -Force
        }
    }

    $suiteExitCode = Invoke-RSSelfTestProcess -Entry $sc -Arguments $args
    $suiteEnd = Get-Date
    $suiteWallMs = [uint64](($suiteEnd - $suiteStart).TotalMilliseconds)

    if ($suiteExitCode -ne 0) {
        $overallExitCode = 1
    }

    # --- Parse results.json ---
    $resultsParsed = $null
    $resultArtifactPath = $null

    foreach ($jsonPath in $jsonCandidates) {
        if (Test-Path $jsonPath) {
            if ($null -eq $resultArtifactPath) {
                $resultArtifactPath = $jsonPath
            }
            try {
                $resultsParsed = Get-Content $jsonPath -Raw | ConvertFrom-Json
                $resultArtifactPath = $jsonPath
                break
            } catch {
                Write-Host "  WARNING: Failed to parse $jsonPath : $_" -ForegroundColor Yellow
            }
        }
    }

    $suiteResult = @{
        Name       = $sc.Name
        Kind       = $sc.Kind
        ExitCode   = $suiteExitCode
        WallMs     = $suiteWallMs
        Parsed     = $resultsParsed
        ResultParsed = $false
        CoverageValid = $false
        Cases      = @()
        Passed     = 0
        Failed     = 0
        Skipped    = 0
        Total      = 0
        DurationMs = 0
        Failures   = @()
        ShuffleSeed = ''
        RepeatCount = 1
    }

    if ($resultsParsed) {
        $suiteResult.ShuffleSeed = Get-JsonValue $resultsParsed @('shuffle_seed') ''
        $suiteResult.RepeatCount = Get-JsonValue $resultsParsed @('repeat_count') 1

        # Handle both direct suite results and aggregated run results
        $casesArray = $null
        $hasCasesArray = $false
        $directCasesProperty = $resultsParsed.PSObject.Properties['cases']
        $directSuiteName = [string](Get-JsonValue $resultsParsed @('suite') '')
        if ($null -ne $directCasesProperty -and $directSuiteName -eq $sc.Name) {
            $casesArray = @($directCasesProperty.Value)
            $hasCasesArray = $true
            $suiteResult.Passed     = Get-JsonValue $resultsParsed @('passed') 0
            $suiteResult.Failed     = Get-JsonValue $resultsParsed @('failed') 0
            $suiteResult.Skipped    = Get-JsonValue $resultsParsed @('skipped') 0
            $suiteResult.DurationMs = Get-JsonValue $resultsParsed @('durationMs', 'duration_ms') 0
        } else {
            # Aggregated run result - find our suite
            $parsedSuites = Get-JsonValue $resultsParsed @('suites') $null
            foreach ($s in @($parsedSuites)) {
                $suiteName  = Get-JsonValue $s @('suite') ''
                $suiteCasesProperty = $s.PSObject.Properties['cases']
                if ($suiteName -eq $sc.Name -and $null -ne $suiteCasesProperty) {
                    $casesArray = if ($null -ne $suiteCasesProperty) { @($suiteCasesProperty.Value) } else { @() }
                    $hasCasesArray = $true
                    $suiteResult.Passed     = Get-JsonValue $s @('passed') 0
                    $suiteResult.Failed     = Get-JsonValue $s @('failed') 0
                    $suiteResult.Skipped    = Get-JsonValue $s @('skipped') 0
                    $suiteResult.DurationMs = Get-JsonValue $s @('durationMs', 'duration_ms') 0
                    $suiteResult.ShuffleSeed = Get-JsonValue $s @('shuffle_seed') $suiteResult.ShuffleSeed
                    $suiteResult.RepeatCount = Get-JsonValue $s @('repeat_count') $suiteResult.RepeatCount
                    break
                }
            }
        }

        if ($hasCasesArray) {
            $suiteResult.ResultParsed = $true
            $suiteResult.Cases = @($casesArray)
            $suiteResult.Total = [int64](Get-JsonValue $(if ($null -ne $directCasesProperty -and $directSuiteName -eq $sc.Name) { $resultsParsed } else { $s }) `
                @('total') ($suiteResult.Passed + $suiteResult.Failed + $suiteResult.Skipped))
            $suiteResult.Failures = @($suiteResult.Cases | Where-Object { (Get-JsonValue $_ @('status') '') -in @('failed', 'crashed') })

            try {
                $caseList = Invoke-RSSelfTestCaseList -Entry $sc
                $expectedSuite = @($caseList.suites | Where-Object { $_.suite -eq $sc.Name } | Select-Object -First 1)
                if (@($expectedSuite).Count -eq 0) {
                    throw "Case listing did not include suite '$($sc.Name)'."
                }

                $expectedNames = @($expectedSuite[0].cases | ForEach-Object { [string]$_.name })
                $allowedExtraNames = if ($sc.Name -eq 'CompareDirectories') { @('setup') } else { @() }
                $coverage = Test-RSSelfTestResultCoverage `
                    -ExpectedCaseNames $expectedNames `
                    -ActualCases $suiteResult.Cases `
                    -AllowedExtraCaseNames $allowedExtraNames `
                    -ExpectedRepeatCount $SelfTestRepeat `
                    -RequireExpectedCases:(-not [string]::IsNullOrWhiteSpace($CaseFilter))
                if (-not $coverage.IsValid) {
                    $suiteResult.Failures += [pscustomobject]@{
                        name = 'selftest_result_coverage'
                        status = 'failed'
                        reason = Format-RSCoverageReason -Coverage $coverage
                        durationMs = 0
                    }
                    ++$suiteResult.Failed
                    $overallExitCode = 1
                } else {
                    $suiteResult.CoverageValid = $true
                }
            } catch {
                $suiteResult.Failures += [pscustomobject]@{
                    name = 'selftest_result_coverage'
                    status = 'failed'
                    reason = $_.Exception.Message
                    durationMs = 0
                }
                ++$suiteResult.Failed
                $overallExitCode = 1
            }
        }
    }
    if ($suiteResult.Failed -gt 0 -or @($suiteResult.Failures).Count -gt 0) {
        $suiteResult.ExitCode = 1
        $overallExitCode = 1
    }

    if ($failureClassificationEnabled -and $suiteResult.ExitCode -ne 0) {
        Write-SubHeader "CLASSIFYING FAILURE: $($sc.Name)"
        $retryAttempts = @(Invoke-RSSelfTestClassificationRetry -Entry $sc -SuiteResult $suiteResult)
        $suiteResult = Add-RSTestResultClassification -Result $suiteResult -RetryResults $retryAttempts
        Write-Host "  Classification: $($suiteResult.Classification)" -ForegroundColor Yellow
        Write-Host "  Reason: $($suiteResult.ClassificationReason)" -ForegroundColor DarkGray
    } else {
        $suiteResult = Add-RSTestResultClassification -Result $suiteResult
    }

    $postWorkspaceSnapshot = Publish-RSRunnerWorkspaceSnapshot -Label "after-$($sc.Id)"
    $suiteResult.WorkspaceSnapshotId = $postWorkspaceSnapshot.snapshot_id
    $workspaceSnapshotRows += [pscustomobject]@{
        entry_id = $sc.Id; phase = 'post-entry'; snapshot_id = $postWorkspaceSnapshot.snapshot_id
        computation_ms = [int64]$postWorkspaceSnapshot.computation_ms
    }
    if ($postWorkspaceSnapshot.snapshot_id -ne $initialWorkspaceSnapshot.snapshot_id) {
        $sourceMutationDetected = $true
        $overallExitCode = 1
        Write-Host "  SOURCE MUTATION: workspace changed during $($sc.Id)." -ForegroundColor Red
    }

    $terminalPublication = Publish-RSRunnerEntryEvidence -Entry $sc -Result $suiteResult `
        -StartedUtc $suiteStart -PostSnapshot $postWorkspaceSnapshot `
        -CandidateArtifact @($resultArtifactPath) -RequireArtifact
    if ($terminalPublication.TerminalResult.status -ne 'PASSED') {
        $overallExitCode = 1
        Write-Host "  TERMINAL RESULT: FAILED ($(@($terminalPublication.TerminalResult.reason_codes) -join ', '))" -ForegroundColor Red
    }

    $allResults += $suiteResult
}

foreach ($entry in $standaloneConfigs) {
    $execution = Invoke-RSRunnerStandaloneValidationEntry -Entry $entry
    Add-RSRunnerStandaloneExecutionResult -Execution $execution
}

if (@($quarantineRepairPlan.Attempts).Count -gt 0) {
    Write-Header "Running quarantine repair lane"
    $repairIndex = 0
    foreach ($attempt in @($quarantineRepairPlan.Attempts)) {
        ++$repairIndex
        $harness = Get-RSObjectValue -Object $attempt -Names @('harness') -DefaultValue '(missing harness)'
        $name = Get-RSObjectValue -Object $attempt -Names @('name') -DefaultValue '(missing name)'
        $owner = Get-RSObjectValue -Object $attempt -Names @('owner') -DefaultValue '(missing owner)'
        $expires = Get-RSObjectValue -Object $attempt -Names @('expires') -DefaultValue '(missing expiry)'
        Write-Host "  Repair: [$harness] $name owner=$owner expires=$expires" -ForegroundColor Yellow

        $attemptResult = Invoke-RSTestQuarantineRepairAttempt -Attempt $attempt -RepoRoot $repoRoot -ArtifactRoot $artifactRoot
        $quarantineRepairAttempts += $attemptResult
        if (-not [bool]$attemptResult.passed) {
            $overallExitCode = 1
        }
        $repairId = [string](Get-RSObjectValue -Object $attempt -Names @('id') -DefaultValue 'repair')
        $postWorkspaceSnapshot = Publish-RSRunnerWorkspaceSnapshot -Label "after-repair-$repairIndex-$repairId"
        $workspaceSnapshotRows += [pscustomobject]@{
            entry_id = $repairId; phase = 'post-repair'; snapshot_id = $postWorkspaceSnapshot.snapshot_id
            computation_ms = [int64]$postWorkspaceSnapshot.computation_ms
        }
        if ($postWorkspaceSnapshot.snapshot_id -ne $initialWorkspaceSnapshot.snapshot_id) {
            $sourceMutationDetected = $true
            $overallExitCode = 1
            Write-Host "  SOURCE MUTATION: workspace changed during repair $repairId." -ForegroundColor Red
        }
    }
}

$finalWorkspaceSnapshot = Publish-RSRunnerWorkspaceSnapshot -Label 'final'
$workspaceSnapshotRows += [pscustomobject]@{
    entry_id = $null; phase = 'final'; snapshot_id = $finalWorkspaceSnapshot.snapshot_id
    computation_ms = [int64]$finalWorkspaceSnapshot.computation_ms
}
if ($finalWorkspaceSnapshot.snapshot_id -ne $initialWorkspaceSnapshot.snapshot_id) {
    $sourceMutationDetected = $true
    $overallExitCode = 1
}
$validationEvidenceStatus = $null
if ($null -ne $validationEvidenceRun) {
    $validationEvidenceStatus = Complete-RSValidationEvidenceRun -Run $validationEvidenceRun `
        -ForceFailed:($overallExitCode -ne 0 -or $sourceMutationDetected -or $quarantineStatus.HasBlockingEntries)
    Write-Host "Validation evidence state: $validationEvidenceStatus" -ForegroundColor DarkGray
    if ($validationEvidenceStatus -ne 'passed') { $overallExitCode = 1 }
}
$overallEndTime = Get-Date
$overallWallMs = [uint64](($overallEndTime - $overallStartTime).TotalMilliseconds)
if ($quarantineStatus.HasBlockingEntries) {
    $overallExitCode = 1
}

$testSandboxAudit = Get-RSTestSandboxDiskAudit `
    -TestRoot $testRunContext.TestRoot `
    -RunId $testRunContext.RunId

# Preserve a runner-owned aggregate artifact before any suite can overwrite the
# native root results.json on a multi-suite run.
$runSummaryPath = Join-Path $artifactRoot 'run-all-tests-results.json'
New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null
$runSummary = New-RSTestRunSummary `
    -Suite $Suite `
    -Platform $Platform `
    -Configuration $Configuration `
    -ExePath $exeFullPath `
    -ArtifactRoot $artifactRoot `
    -RunStartedUtc $($overallStartTime.ToUniversalTime()) `
    -RunEndedUtc $($overallEndTime.ToUniversalTime()) `
    -DurationMs $overallWallMs `
    -ExitCode $overallExitCode `
    -TimeoutMultiplier $TimeoutMultiplier `
    -FailFast:$FailFast `
    -CaseFilter $CaseFilter `
    -TestRoot $testRunContext.TestRoot `
    -RunId $testRunContext.RunId `
    -QuarantineStatus $quarantineStatus `
    -QuarantineRepairAttempts $quarantineRepairAttempts `
    -TestSandboxAudit $testSandboxAudit `
    -Results $allResults
$summaryForConsumers = $runSummary
if ($null -ne $validationEvidenceRun) {
    $legacySummaryPath = Join-Path $artifactRoot 'run-all-tests-results-v1.json'
    $runSummary | ConvertTo-Json -Depth 32 | Set-Content -LiteralPath $legacySummaryPath -Encoding UTF8
    $summaryForConsumers = New-RSValidationRunSummaryV2 -Run $validationEvidenceRun `
        -Plan $validationPlan -Decision $shadowDecision -LegacySummary $runSummary `
        -EvidenceRoot $validationEvidenceRoot `
        -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') -CoverageScope $validationCoverageScope
    Write-RSAtomicCanonicalJsonFile -LiteralPath $runSummaryPath -InputObject $summaryForConsumers `
        -SchemaPath (Join-Path $repoRoot 'Specs\Testing\RunAllTestsSummary.schema.json') `
        -AllowedRoot $artifactRoot
} else {
    $runSummary | ConvertTo-Json -Depth 32 | Set-Content -LiteralPath $runSummaryPath -Encoding UTF8
}
$terminalState = if ($summaryForConsumers.schema -eq 'red-salamander.run-all-tests.v2') {
    switch ([string]$summaryForConsumers.repository_verdict) {
        'PASSED' { 'PASSED' }
        'FAILED' { 'FAILED' }
        'NOT_EVALUATED' { 'NOT_EVALUATED' }
        'BLOCKED' { 'CORRUPT_EVIDENCE' }
        default { 'CORRUPT_EVIDENCE' }
    }
} elseif ($overallExitCode -eq 0) {
    'PASSED'
} else {
    'FAILED'
}
$overallExitCode = Get-RSValidationExitCode -TerminalState $terminalState
$caseHistoryRows = @(Convert-RSTestRunSummaryToCaseHistoryRows -Summary $runSummary)
$caseHistoryPath = Join-Path $artifactRoot 'run-all-tests-case-history.jsonl'
Convert-RSTestCaseHistoryRowsToJsonl -Rows $caseHistoryRows | Set-Content -LiteralPath $caseHistoryPath -Encoding UTF8
$caseDashboardPath = Join-Path $artifactRoot 'run-all-tests-dashboard.md'
Convert-RSTestCaseHistoryRowsToDashboardMarkdown -Rows $caseHistoryRows -Summary $runSummary | Set-Content -LiteralPath $caseDashboardPath -Encoding UTF8

$githubStepSummaryPath = [Environment]::GetEnvironmentVariable('GITHUB_STEP_SUMMARY', 'Process')
if (-not [string]::IsNullOrWhiteSpace($githubStepSummaryPath)) {
    try {
        $githubSummaryMarkdown = Convert-RSTestRunSummaryToGitHubStepSummary -Summary $summaryForConsumers
        Add-Content -LiteralPath $githubStepSummaryPath -Value $githubSummaryMarkdown -Encoding UTF8
    } catch {
        Write-Host "WARNING: Failed to write GitHub step summary: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# --- Display comprehensive summary ---

Write-Header "TEST RESULTS SUMMARY"

$totalPassed  = 0
$totalFailed  = 0
$totalSkipped = 0
$totalCases   = 0

foreach ($r in $allResults) {
    $totalPassed  += $r.Passed
    $totalFailed  += $r.Failed
    $totalSkipped += $r.Skipped
    $totalCases   += ($r.Passed + $r.Failed + $r.Skipped)

    $statusIcon  = if ($r.ExitCode -eq 0) { '[PASS]' } else { '[FAIL]' }
    $statusColor = if ($r.ExitCode -eq 0) { 'Green' } else { 'Red' }

    Write-Host ""
    Write-Host "  $statusIcon " -ForegroundColor $statusColor -NoNewline
    Write-Host "$($r.Name)" -ForegroundColor White -NoNewline
    Write-Host "  ($(Format-Duration $r.WallMs))" -ForegroundColor DarkGray

    if ($r.Parsed) {
        Write-Host "         Passed: $($r.Passed)" -ForegroundColor Green -NoNewline
        Write-Host "  Failed: $($r.Failed)" -ForegroundColor $(if ($r.Failed -gt 0) { 'Red' } else { 'Green' }) -NoNewline
        Write-Host "  Skipped: $($r.Skipped)" -ForegroundColor $(if ($r.Skipped -gt 0) { 'DarkYellow' } else { 'Green' })
    } else {
        Write-Host "         (results.json not found - exit code: $($r.ExitCode))" -ForegroundColor Yellow
    }
    $classification = Get-JsonValue $r @('Classification', 'classification') ''
    if (-not [string]::IsNullOrWhiteSpace($classification) -and $classification -ne 'PASSED') {
        $classificationReason = Get-JsonValue $r @('ClassificationReason', 'classification_reason') ''
        Write-Host "         Classification: $classification" -ForegroundColor Yellow
        if (-not [string]::IsNullOrWhiteSpace($classificationReason)) {
            Write-Host "         $classificationReason" -ForegroundColor DarkGray
        }
    }
    $retryAttempts = @(Get-JsonValue $r @('RetryAttempts', 'retry_attempts') @())
    if (@($retryAttempts).Count -gt 0) {
        Write-Host "         Retry attempts: $(@($retryAttempts).Count)" -ForegroundColor DarkGray
    }
    $outputLogPath = Get-JsonValue $r @('OutputLogPath', 'output_log_path') ''
    if (-not [string]::IsNullOrWhiteSpace($outputLogPath)) {
        Write-Host "         Output: $outputLogPath" -ForegroundColor DarkGray
    }
}

# --- Overall totals ---

Write-Host ""
$overallLine = '-' * 55
Write-Host "  $overallLine" -ForegroundColor DarkGray
$overallIcon = $terminalState.Replace('_', ' ')
$overallColor = switch ($terminalState) {
    'PASSED' { 'Green' }
    'NOT_EVALUATED' { 'Yellow' }
    default { 'Red' }
}

Write-Host ""
Write-Host "  Overall: " -NoNewline
Write-Host $overallIcon -ForegroundColor $overallColor -NoNewline
Write-Host "  Total: $totalCases" -ForegroundColor White -NoNewline
Write-Host "  Passed: $totalPassed" -ForegroundColor Green -NoNewline
Write-Host "  Failed: $totalFailed" -ForegroundColor $(if ($totalFailed -gt 0) { 'Red' } else { 'Green' }) -NoNewline
Write-Host "  Skipped: $totalSkipped" -ForegroundColor $(if ($totalSkipped -gt 0) { 'DarkYellow' } else { 'Green' })
Write-Host "  Wall time: $(Format-Duration $overallWallMs)" -ForegroundColor DarkGray

if ($runSummary.PSObject.Properties['classifications']) {
    Write-Host "  Classifications: " -ForegroundColor DarkGray -NoNewline
    Write-Host "flaky=$($runSummary.classifications.flaky) " -ForegroundColor Yellow -NoNewline
    Write-Host "regression=$($runSummary.classifications.regression) " -ForegroundColor Red -NoNewline
    Write-Host "isolation_suspect=$($runSummary.classifications.isolation_suspect) " -ForegroundColor Yellow -NoNewline
    Write-Host "unclassified_failure=$($runSummary.classifications.unclassified_failure)" -ForegroundColor Yellow
}
if ($runSummary.PSObject.Properties['quarantine']) {
    Write-Host "  Quarantine: " -ForegroundColor DarkGray -NoNewline
    Write-Host "active=$($runSummary.quarantine.active_count) " -ForegroundColor Yellow -NoNewline
    Write-Host "invalid=$($runSummary.quarantine.invalid_count) " -ForegroundColor Yellow -NoNewline
    Write-Host "repair_attempts=$($runSummary.quarantine.repair_attempt_count) " -ForegroundColor Yellow -NoNewline
    Write-Host "repair_reproduced=$($runSummary.quarantine.repair_reproduced_count)" -ForegroundColor Yellow
}

# --- Failure details ---

$allFailures = @()
foreach ($r in $allResults) {
    foreach ($f in $r.Failures) {
        $allFailures += @{
            Suite    = $r.Name
            CaseName = Get-JsonValue $f @('name') '(unnamed case)'
            Reason   = Get-JsonValue $f @('reason') '(no reason provided)'
            Duration = Get-JsonValue $f @('durationMs', 'duration_ms') 0
        }
    }
}

if ($allFailures.Count -gt 0) {
    Write-SubHeader "FAILURE DETAILS ($($allFailures.Count) failing case$(if ($allFailures.Count -ne 1) { 's' }))"

    foreach ($failure in $allFailures) {
        Write-Host ""
        Write-Host "  FAIL " -ForegroundColor Red -NoNewline
        Write-Host "[$($failure.Suite)] " -ForegroundColor Yellow -NoNewline
        Write-Host "$($failure.CaseName)" -ForegroundColor White
        Write-Host "       Reason: " -ForegroundColor DarkGray -NoNewline
        Write-Host "$($failure.Reason)" -ForegroundColor Red
        if ($failure.Duration -gt 0) {
            Write-Host "       Duration: $(Format-Duration $failure.Duration)" -ForegroundColor DarkGray
        }
    }
}

# --- Skipped case summary (if any) ---

$allSkipped = @()
foreach ($r in $allResults) {
    if ($r.Cases) {
        foreach ($c in $r.Cases) {
            if ((Get-JsonValue $c @('status') '') -eq 'skipped') {
                $allSkipped += @{
                    Suite  = $r.Name
                    Name   = Get-JsonValue $c @('name') '(unnamed case)'
                    Reason = Get-JsonValue $c @('reason') '(no reason)'
                }
            }
        }
    }
}

if ($allSkipped.Count -gt 0) {
    Write-SubHeader "SKIPPED CASES ($($allSkipped.Count) case$(if ($allSkipped.Count -ne 1) { 's' }))"

    foreach ($skip in $allSkipped) {
        Write-Host "  SKIP " -ForegroundColor DarkYellow -NoNewline
        Write-Host "[$($skip.Suite)] " -ForegroundColor Yellow -NoNewline
        Write-Host "$($skip.Name)" -ForegroundColor White -NoNewline
        Write-Host " - $($skip.Reason)" -ForegroundColor DarkGray
    }
}

if ($quarantineStatus.HasBlockingEntries) {
    Write-SubHeader "QUARANTINE BLOCKERS"
    foreach ($errorMessage in @($quarantineStatus.Errors)) {
        Write-Host "  INVALID " -ForegroundColor Red -NoNewline
        Write-Host $errorMessage -ForegroundColor Red
    }
    foreach ($entry in @($quarantineStatus.ActiveEntries)) {
        $harness = Get-RSObjectValue -Object $entry -Names @('harness') -DefaultValue '(missing harness)'
        $name = Get-RSObjectValue -Object $entry -Names @('name') -DefaultValue '(missing name)'
        $owner = Get-RSObjectValue -Object $entry -Names @('owner') -DefaultValue '(missing owner)'
        $expires = Get-RSObjectValue -Object $entry -Names @('expires') -DefaultValue '(missing expiry)'
        Write-Host "  ACTIVE " -ForegroundColor Yellow -NoNewline
        Write-Host "[$harness] $name owner=$owner expires=$expires" -ForegroundColor Yellow
    }
    foreach ($attempt in @($quarantineRepairAttempts)) {
        $statusText = if ([bool]$attempt.passed) { 'REPAIR PASS' } else { 'REPAIR FAIL' }
        $statusColor = if ([bool]$attempt.passed) { 'Green' } else { 'Red' }
        Write-Host "  $statusText " -ForegroundColor $statusColor -NoNewline
        Write-Host "[$($attempt.harness)] $($attempt.name)" -ForegroundColor $statusColor
        if (-not [string]::IsNullOrWhiteSpace($attempt.reason)) {
            Write-Host "         $($attempt.reason)" -ForegroundColor DarkGray
        }
    }
}

# --- Artifact locations ---

Write-SubHeader "ARTIFACTS"
Write-Host "  Test root: $($testRunContext.TestRoot)" -ForegroundColor DarkGray
Write-Host "  Run id:    $($testRunContext.RunId)" -ForegroundColor DarkGray
Write-Host "  Disk audit issues: $($testSandboxAudit.issue_count)" -ForegroundColor $(if ($testSandboxAudit.issue_count -eq 0) { 'DarkGray' } else { 'Yellow' })
Write-Host "  Last run: $artifactRoot" -ForegroundColor DarkGray
if (Test-Path $artifactRoot) {
    $jsonFiles = Get-ChildItem $artifactRoot -Filter '*.json' -Recurse -ErrorAction SilentlyContinue
    foreach ($jf in $jsonFiles) {
        $relPath = $jf.FullName.Substring($artifactRoot.Length + 1)
        Write-Host "    $relPath" -ForegroundColor DarkGray
    }
}
Write-Host ""

# --- Hints ---

if ($terminalState -in @('FAILED', 'CORRUPT_EVIDENCE')) {
    Write-SubHeader "NEXT STEPS"
    Write-Host "  1. Review failure reasons above" -ForegroundColor Yellow
    Write-Host "  2. Check trace logs: $artifactRoot\<suite>\trace.txt" -ForegroundColor Yellow
    Write-Host "  3. Re-run a single failing case:" -ForegroundColor Yellow
    Write-Host "     $exeFullPath --<suite>-selftest --selftest-case=<case_name>" -ForegroundColor DarkGray
    Write-Host "  4. Compare with previous run:" -ForegroundColor Yellow
    Write-Host "     .\Tools\CompareTestRuns.ps1 <old_run_path> <new_run_path>" -ForegroundColor DarkGray
    Write-Host ""
}

    exit $overallExitCode
}
finally {
    Restore-RSTestMachineResourceEnvironment -Lease $testMachineResourceEnvironmentLease
    foreach ($entry in $runnerEnvironmentPrevious.GetEnumerator()) {
        [Environment]::SetEnvironmentVariable([string]$entry.Key, $entry.Value, 'Process')
    }
    Exit-RSArtifactOperationLock -Lock $artifactOperationLock
}
