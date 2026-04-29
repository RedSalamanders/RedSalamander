<#
.SYNOPSIS
    Builds and runs all RedSalamander self-test suites, then displays a comprehensive summary.

.DESCRIPTION
    This script:
      1. Optionally builds the Debug configuration of RedSalamander.
      2. Runs each self-test suite (Compare, Commands, FileOps) or a selected subset.
      3. Parses the results.json artifacts from each suite.
      4. Displays a color-coded summary with pass/fail/skip counts, timing, and failure details.
      5. Returns exit code 0 on full success, 1 on any failure.

    Self-test artifacts are written to:
      %LOCALAPPDATA%\RedSalamander\SelfTest\last_run\

.PARAMETER Suite
    Which suite(s) to run:
      All       - Run all three suites (default)
      Compare   - Run only --compare-selftest
      Commands  - Run only --commands-selftest
      FileOps   - Run only --fileops-selftest

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
    [ValidateSet('All', 'Compare', 'Commands', 'FileOps')]
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

# --- Resolve paths ---

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $repoRoot) { $repoRoot = $PSScriptRoot | Split-Path -Parent }

if ($ExePath) {
    $exeFullPath = $ExePath
} else {
    $exeFullPath = Join-Path $repoRoot ".build\$Platform\$Configuration\RedSalamander.exe"
}

$artifactRoot = Join-Path $env:LOCALAPPDATA 'RedSalamander\SelfTest\last_run'

# --- Build step ---

if (-not $SkipBuild) {
    Write-Header "Building $Configuration|$Platform"
    $buildScript = Join-Path $repoRoot 'build.ps1'
    if (-not (Test-Path $buildScript)) {
        Write-Host "ERROR: build.ps1 not found at $buildScript" -ForegroundColor Red
        exit 1
    }

    & $buildScript -Configuration $Configuration -ProjectName RedSalamander
    if ($LASTEXITCODE -ne 0) {
        Write-Host "BUILD FAILED (exit code $LASTEXITCODE)" -ForegroundColor Red
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

# --- Determine suites to run ---

$suiteConfigs = @()

switch ($Suite) {
    'All' {
        $suiteConfigs += @{ Name = 'CompareDirectories'; Flag = '--compare-selftest';  JsonName = 'compare' }
        $suiteConfigs += @{ Name = 'Commands';           Flag = '--commands-selftest';  JsonName = 'commands' }
        $suiteConfigs += @{ Name = 'FileOperations';     Flag = '--fileops-selftest';   JsonName = 'fileops' }
    }
    'Compare' {
        $suiteConfigs += @{ Name = 'CompareDirectories'; Flag = '--compare-selftest'; JsonName = 'compare' }
    }
    'Commands' {
        $suiteConfigs += @{ Name = 'Commands';           Flag = '--commands-selftest'; JsonName = 'commands' }
    }
    'FileOps' {
        $suiteConfigs += @{ Name = 'FileOperations';     Flag = '--fileops-selftest'; JsonName = 'fileops' }
    }
}

# --- Run suites ---

$allResults = @()
$overallStartTime = Get-Date
$overallExitCode = 0

foreach ($sc in $suiteConfigs) {
    Write-Header "Running: $($sc.Name)"

    $args = @($sc.Flag)
    if ($FailFast) { $args += '--selftest-fail-fast' }
    if ($TimeoutMultiplier -ne 1.0) { $args += "--selftest-timeout-multiplier=$TimeoutMultiplier" }
    if ($CaseFilter) { $args += "--selftest-case=$CaseFilter" }

    $suiteStart = Get-Date
    Write-Host "  Exe:  $exeFullPath" -ForegroundColor DarkGray
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

    $process = Start-Process -FilePath $exeFullPath -ArgumentList $args -WorkingDirectory $repoRoot -Wait -PassThru
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
        $directCases = Get-JsonValue $resultsParsed @('cases') $null
        if ($null -ne $directCases) {
            $casesArray = $directCases
            $suiteResult.Passed     = Get-JsonValue $resultsParsed @('passed') 0
            $suiteResult.Failed     = Get-JsonValue $resultsParsed @('failed') 0
            $suiteResult.Skipped    = Get-JsonValue $resultsParsed @('skipped') 0
            $suiteResult.DurationMs = Get-JsonValue $resultsParsed @('durationMs', 'duration_ms') 0
        } else {
            # Aggregated run result - find our suite
            $parsedSuites = Get-JsonValue $resultsParsed @('suites') $null
            foreach ($s in @($parsedSuites)) {
                $suiteName  = Get-JsonValue $s @('suite') ''
                $suiteCases = Get-JsonValue $s @('cases') $null
                if ($suiteName -eq $sc.Name -or $null -ne $suiteCases) {
                    $casesArray = $suiteCases
                    $suiteResult.Passed     = Get-JsonValue $s @('passed') 0
                    $suiteResult.Failed     = Get-JsonValue $s @('failed') 0
                    $suiteResult.Skipped    = Get-JsonValue $s @('skipped') 0
                    $suiteResult.DurationMs = Get-JsonValue $s @('durationMs', 'duration_ms') 0
                    break
                }
            }
        }

        if ($null -ne $casesArray) {
            $suiteResult.Cases = @($casesArray)
            $suiteResult.Failures = @($suiteResult.Cases | Where-Object { (Get-JsonValue $_ @('status') '') -eq 'failed' })
        }
    }
    if ($suiteResult.Failed -gt 0 -or @($suiteResult.Failures).Count -gt 0) {
        $overallExitCode = 1
    }

    $allResults += $suiteResult
}

$overallEndTime = Get-Date
$overallWallMs = [uint64](($overallEndTime - $overallStartTime).TotalMilliseconds)

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
