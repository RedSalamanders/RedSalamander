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
      5. Returns exit code 0 on full success, 1 on any failure.

    Self-test artifacts are written under:
      REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\selftest\last_run\
    The runner sets REDSALAMANDER_TEST_ROOT to .build\TestSandbox by default.

.PARAMETER Suite
    Which suite(s) to run:
      All       - Run all three self-test suites (default)
      Compare   - Run only --compare-selftest
      Commands  - Run only --commands-selftest
      FileOps   - Run only --fileops-selftest
      CI        - Run the GitHub Actions PR gate through the unified runner
      Full      - Run all self-tests plus standalone/native and script tests

.PARAMETER SkipBuild
    When set, skips the build step and uses the existing Debug binary.

.PARAMETER SkipLegacySandboxCleanup
    When set, skips the pre-run legacy test-sandbox cleanup reaper.

.PARAMETER FailFast
    When set, passes --selftest-fail-fast to abort after first case failure.

.PARAMETER TimeoutMultiplier
    Scales all test timeouts. Use >1.0 on slow CI machines. Default: 1.0.

.PARAMETER CaseFilter
    Run only the matching case name or prefix (passed as --selftest-case=...).

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
    Optional JSONL quarantine metadata file. Defaults to Tools\test-quarantine.jsonl.
    Invalid or active entries keep the aggregate run red.

.PARAMETER Platform
    Target platform. Default: x64.

.PARAMETER Configuration
    Build configuration. Default: Debug (self-tests require Debug builds).

.PARAMETER ExePath
    Override the executable path. By default, resolves from .build\<Platform>\<Configuration>\.

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

    [switch]$SkipLegacySandboxCleanup,

    [switch]$FailFast,

    [double]$TimeoutMultiplier = 1.0,

    [string]$CaseFilter = '',

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

    [string]$ExePath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Helpers ---

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $repoRoot) { $repoRoot = $PSScriptRoot | Split-Path -Parent }

$testRunPlanScript = Join-Path $repoRoot 'Tools\TestRunPlan.ps1'
if (-not (Test-Path $testRunPlanScript)) {
    Write-Host "ERROR: test run plan helper not found at $testRunPlanScript" -ForegroundColor Red
    exit 1
}

. $testRunPlanScript

$processStreamingScript = Join-Path $repoRoot 'Tools\ProcessStreaming.ps1'
if (-not (Test-Path $processStreamingScript)) {
    Write-Host "ERROR: process streaming helper not found at $processStreamingScript" -ForegroundColor Red
    exit 1
}

. $processStreamingScript

$artifactOperationLockScript = Join-Path $repoRoot 'Tools\ArtifactOperationLock.ps1'
if (-not (Test-Path $artifactOperationLockScript)) {
    Write-Host "ERROR: artifact operation lock helper not found at $artifactOperationLockScript" -ForegroundColor Red
    exit 1
}

. $artifactOperationLockScript

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
                $exitCode = Invoke-RSStreamingProcess -FilePath $Entry.Path -Arguments @($Entry.Arguments) -WorkingDirectory $Entry.WorkingDirectory -LogPath $outputLogPath
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

                $pesterParameters = New-RSPesterInvokeParameters -Path $Entry.Path -Arguments @($Entry.Arguments)
                $pesterResult = Invoke-Pester @pesterParameters
                $failedCount = Get-JsonValue $pesterResult @('FailedCount', 'Failed') 0
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
        Parsed     = $null
        Cases      = @()
        Passed     = if ($exitCode -eq 0) { 1 } else { 0 }
        Failed     = if ($exitCode -eq 0) { 0 } else { 1 }
        Skipped    = 0
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

    $process = $null
    $stdout = ''
    $stderr = ''
    $exitCode = 1
    try {
        $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $process.WaitForExit()
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        $exitCode = $process.ExitCode
    }
    finally {
        Close-RSContainedProcess -Process $process
    }

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

$artifactOperationLock = $null
try {
    $operationMode = if ($SkipBuild) { 'existing artifacts' } else { 'build and test' }
    $artifactOperationLock = Enter-RSArtifactOperationLock `
        -RepoRoot $repoRoot `
        -Operation "Run-AllTests $Suite ($operationMode) $Configuration|$Platform" `
        -Scope @{
            kind = 'test'
            target = $operationMode
            suite = $Suite
            configuration = $Configuration
            platform = $Platform
        }

    if ($artifactOperationLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $repoRoot `
                -Reason "The previous build/test owner exited without clearing the exclusive artifact-operation lock." `
                -AbandonedOwner $artifactOperationLock.AbandonedOwner)
    }

    if (Test-RSArtifactOperationContaminated -RepoRoot $repoRoot) {
        $markerPath = Get-RSArtifactContaminationMarkerPath -RepoRoot $repoRoot
        throw "Shared build artifacts may be mixed after an interrupted operation. Run build.ps1 -Rebuild before starting tests. Marker: $markerPath"
    }

    Assert-RSNoResidualArtifactToolProcesses -RepoRoot $repoRoot

# --- Resolve paths ---

if ($ExePath) {
    $exeFullPath = $ExePath
} else {
    $exeFullPath = Join-Path $repoRoot ".build\$Platform\$Configuration\RedSalamander.exe"
}

$testRunContext = New-RSTestRunContext `
    -RepoRoot $repoRoot `
    -SelfTestRootOverride ''
$env:REDSALAMANDER_TEST_ROOT = $testRunContext.TestRoot
$env:REDSALAMANDER_TEST_RUN_ID = $testRunContext.RunId
[Environment]::SetEnvironmentVariable('REDSALAMANDER_SELFTEST_ROOT', $null, 'Process')
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

if (-not $SkipLegacySandboxCleanup) {
    $cleanupScript = Join-Path $repoRoot 'Tools\Clean-TestSandbox.ps1'
    & $cleanupScript -Apply -Confirm:$false | Out-Null
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
    Write-Host "Skipping build (using existing binary)." -ForegroundColor DarkGray
}

if (-not (Test-Path $exeFullPath)) {
    Write-Host "ERROR: Executable not found: $exeFullPath" -ForegroundColor Red
    Write-Host "  Build a Debug configuration first, or use -ExePath to specify the path." -ForegroundColor Yellow
    exit 1
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
        -RepeatCount $SelfTestRepeat `
        -ShuffleSeed $SelfTestShuffleSeed `
        -PerfBudgetPath $PerfBudgetPath `
        -RequirePerfBudgets:$RequirePerfBudgets `
        -SelfTestFlakyProofCase $SelfTestFlakyProofCase `
        -SelfTestOrderProofCase $SelfTestOrderProofCase)

$suiteConfigs = @($testPlan | Where-Object { $_.Kind -eq 'SelfTest' })
$standaloneConfigs = @($testPlan | Where-Object { $_.Kind -ne 'SelfTest' })
$failureClassificationEnabled = ($ClassifyFailures -or $Suite -eq 'CI')
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

foreach ($sc in $suiteConfigs) {
    Write-Header "Running: $($sc.Name)"

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

    foreach ($jsonPath in $jsonCandidates) {
        if (Test-Path $jsonPath) {
            try {
                $resultsParsed = Get-Content $jsonPath -Raw | ConvertFrom-Json
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
        Cases      = @()
        Passed     = 0
        Failed     = 0
        Skipped    = 0
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
        if ($null -ne $directCasesProperty) {
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
                if ($suiteName -eq $sc.Name -or $null -ne $suiteCasesProperty) {
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
            $suiteResult.Cases = @($casesArray)
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

    $allResults += $suiteResult
}

foreach ($entry in $standaloneConfigs) {
    Write-Header "Running: $($entry.Name)"
    Write-Host "  Kind: $($entry.Kind)" -ForegroundColor DarkGray
    Write-Host "  Path: $($entry.Path)" -ForegroundColor DarkGray
    if (@($entry.Arguments).Count -gt 0) {
        Write-Host "  Args: $(@($entry.Arguments) -join ' ')" -ForegroundColor DarkGray
    }
    Write-Host ""

    $suiteResult = Invoke-RSTestPlanEntry -Entry $entry -RepoRoot $repoRoot -ArtifactRoot $artifactRoot
    if ($suiteResult.ExitCode -ne 0) {
        $overallExitCode = 1
    }

    if ($failureClassificationEnabled -and $suiteResult.ExitCode -ne 0) {
        Write-SubHeader "CLASSIFYING FAILURE: $($entry.Name)"
        $retryAttempts = @(Invoke-RSTestEntryClassificationRetry -Entry $entry -RepoRoot $repoRoot -ArtifactRoot $artifactRoot)
        $suiteResult = Add-RSTestResultClassification -Result $suiteResult -RetryResults $retryAttempts
        Write-Host "  Classification: $($suiteResult.Classification)" -ForegroundColor Yellow
        Write-Host "  Reason: $($suiteResult.ClassificationReason)" -ForegroundColor DarkGray
    } else {
        $suiteResult = Add-RSTestResultClassification -Result $suiteResult
    }

    $allResults += $suiteResult
}

if (@($quarantineRepairPlan.Attempts).Count -gt 0) {
    Write-Header "Running quarantine repair lane"
    foreach ($attempt in @($quarantineRepairPlan.Attempts)) {
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
    }
}

$overallEndTime = Get-Date
$overallWallMs = [uint64](($overallEndTime - $overallStartTime).TotalMilliseconds)
if ($quarantineStatus.HasBlockingEntries) {
    $overallExitCode = 1
}

$testSandboxAudit = Get-RSTestSandboxDiskAudit `
    -TestRoot $testRunContext.TestRoot `
    -RunId $testRunContext.RunId `
    -DriveRoots (Get-RSFixedDriveRoots)

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
$runSummary | ConvertTo-Json -Depth 32 | Set-Content -LiteralPath $runSummaryPath -Encoding UTF8
$caseHistoryRows = @(Convert-RSTestRunSummaryToCaseHistoryRows -Summary $runSummary)
$caseHistoryPath = Join-Path $artifactRoot 'run-all-tests-case-history.jsonl'
Convert-RSTestCaseHistoryRowsToJsonl -Rows $caseHistoryRows | Set-Content -LiteralPath $caseHistoryPath -Encoding UTF8
$caseDashboardPath = Join-Path $artifactRoot 'run-all-tests-dashboard.md'
Convert-RSTestCaseHistoryRowsToDashboardMarkdown -Rows $caseHistoryRows -Summary $runSummary | Set-Content -LiteralPath $caseDashboardPath -Encoding UTF8

$githubStepSummaryPath = [Environment]::GetEnvironmentVariable('GITHUB_STEP_SUMMARY', 'Process')
if (-not [string]::IsNullOrWhiteSpace($githubStepSummaryPath)) {
    try {
        $githubSummaryMarkdown = Convert-RSTestRunSummaryToGitHubStepSummary -Summary $runSummary
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
$overallIcon  = if ($overallExitCode -eq 0) { 'PASSED' } else { 'FAILED' }
$overallColor = if ($overallExitCode -eq 0) { 'Green' } else { 'Red' }

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

if ($overallExitCode -ne 0) {
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
    Exit-RSArtifactOperationLock -Lock $artifactOperationLock
}
