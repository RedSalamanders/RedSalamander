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

    Self-test artifacts are written to:
      %LOCALAPPDATA%\RedSalamander\SelfTest\last_run\

.PARAMETER Suite
    Which suite(s) to run:
      All       - Run all three self-test suites (default)
      Compare   - Run only --compare-selftest
      Commands  - Run only --commands-selftest
      FileOps   - Run only --fileops-selftest
      Full      - Run all self-tests plus standalone/native and script tests

.PARAMETER SkipBuild
    When set, skips the build step and uses the existing Debug binary.

.PARAMETER FailFast
    When set, passes --selftest-fail-fast to abort after first case failure.

.PARAMETER TimeoutMultiplier
    Scales all test timeouts. Use >1.0 on slow CI machines. Default: 1.0.

.PARAMETER CaseFilter
    Run only the matching case name or prefix (passed as --selftest-case=...).

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
    [ValidateSet('All', 'Compare', 'Commands', 'FileOps', 'Full')]
    [string]$Suite = 'All',

    [switch]$SkipBuild,

    [switch]$FailFast,

    [double]$TimeoutMultiplier = 1.0,

    [string]$CaseFilter = '',

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
        'skipped' { return 'DarkYellow' }
        default   { return 'Gray' }
    }
}

function Get-JsonValue($object, [string[]]$names, $defaultValue = $null) {
    if ($null -eq $object) { return $defaultValue }
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

                $pesterResult = Invoke-Pester -Path $Entry.Path -PassThru
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

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    [void]$process.Start()
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    if ($process.ExitCode -ne 0) {
        throw "Case listing failed with exit code $($process.ExitCode): $stderr"
    }

    if ([string]::IsNullOrWhiteSpace($stdout)) {
        throw 'Case listing produced no JSON output.'
    }

    return ($stdout | ConvertFrom-Json)
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

# --- Resolve paths ---

if ($ExePath) {
    $exeFullPath = $ExePath
} else {
    $exeFullPath = Join-Path $repoRoot ".build\$Platform\$Configuration\RedSalamander.exe"
}

$artifactRoot = Get-RSSelfTestArtifactRoot

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
        -CaseFilter $CaseFilter)

$suiteConfigs = @($testPlan | Where-Object { $_.Kind -eq 'SelfTest' })
$standaloneConfigs = @($testPlan | Where-Object { $_.Kind -ne 'SelfTest' })

# --- Run suites ---

$allResults = @()
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

    $process = Start-Process -FilePath $sc.Path -ArgumentList $args -WorkingDirectory $sc.WorkingDirectory -Wait -PassThru
    $suiteExitCode = if ($null -ne $process.ExitCode) { $process.ExitCode } else { 0 }
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
    }

    if ($resultsParsed) {
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
                    break
                }
            }
        }

        if ($hasCasesArray) {
            $suiteResult.Cases = @($casesArray)
            $suiteResult.Failures = @($suiteResult.Cases | Where-Object { (Get-JsonValue $_ @('status') '') -eq 'failed' })

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

    $allResults += $suiteResult
}

$overallEndTime = Get-Date
$overallWallMs = [uint64](($overallEndTime - $overallStartTime).TotalMilliseconds)

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
    -Results $allResults
$runSummary | ConvertTo-Json -Depth 32 | Set-Content -LiteralPath $runSummaryPath -Encoding UTF8

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

# --- Artifact locations ---

Write-SubHeader "ARTIFACTS"
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
