<#
.SYNOPSIS
    Compares two archived selftest run folders and reports differences.

.DESCRIPTION
    Given two archived run folders (each containing a suite results JSON like
    fileops_results.json / compare_results.json / commands_results.json), this script:

    1. Prints a summary comparing passed/failed/skipped counts and duration.
    2. SHA-256 hashes every file in both folders and lists files that were
       Added, Removed, or Changed.
    3. Diffs individual test cases by name, reporting status changes, reason
       changes, and duration deltas.
    4. Optionally shows a line-level diff of the trace logs.

    Run folders live under Specs\TestRuns\ and are timestamped directories
    created by the self-test harness.

.PARAMETER OldRun
    Path to the baseline (older) run folder.

.PARAMETER NewRun
    Path to the candidate (newer) run folder.

.PARAMETER Suite
    Which suite artifacts to compare. In 'Auto' mode, picks the first matching results JSON
    in this order:
        fileops_results.json, compare_results.json, commands_results.json,
        selftest_run_results.json, results.json

.PARAMETER ShowTraceDiff
    When set, loads trace.txt (or fileops_trace.txt) from both folders and
    prints a line-level diff using Compare-Object.

.PARAMETER MaxTraceDiffLines
    Maximum number of trace-diff lines to display. Default: 120.

.OUTPUTS
    Human-readable summary, artifact, case, and optional trace differences. No supported pipeline objects.

.NOTES
    Prerequisites: two archived self-test run folders with compatible result JSON. Side effects: none; both run trees are read-only. Missing required results fail explicitly, while trace artifacts are required only with ShowTraceDiff. Primary use is manual before/after evidence review.

.EXAMPLE
    .\Tools\CompareTestRuns.ps1 `
        Specs\TestRuns\<ComputerHashName>\FileOps\2026-02-27_085402 `
        Specs\TestRuns\<ComputerHashName>\FileOps\2026-02-28_000800

    Compares two adjacent runs and shows summary, file, and case differences.

.EXAMPLE
    .\Tools\CompareTestRuns.ps1 `
        Specs\TestRuns\<ComputerHashName>\FileOps\2026-02-27_121904 `
        Specs\TestRuns\<ComputerHashName>\FileOps\2026-02-27_141202 `
        -ShowTraceDiff -MaxTraceDiffLines 50

    Same comparison, plus a truncated trace-log diff.
#>
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$OldRun,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$NewRun,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Auto', 'FileOps', 'CompareDirectories', 'Commands', 'SelfTest')]
    [string]$Suite = 'Auto',

    [switch]$ShowTraceDiff,

    [int]$MaxTraceDiffLines = 120
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Modules\Reporting\TestRunReporting.psm1') -Force

function Get-RunFileMap([string]$RunRoot) {
    $map = @{}
    $root = (Resolve-Path -LiteralPath $RunRoot).Path
    $prefix = if ($root.EndsWith([System.IO.Path]::DirectorySeparatorChar)) { $root } else { ($root + [System.IO.Path]::DirectorySeparatorChar) }

    Get-ChildItem -LiteralPath $root -File -Recurse | ForEach-Object {
        $rel = $_.FullName.Substring($prefix.Length)
        $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        $map[$rel] = [pscustomobject]@{
            Name = $rel
            Size = $_.Length
            Hash = $hash
        }
    }

    return $map
}

function Compare-RunFiles([hashtable]$OldFiles, [hashtable]$NewFiles) {
    $all = @($OldFiles.Keys + $NewFiles.Keys) | Sort-Object -Unique

    $changed = @()
    foreach ($name in $all) {
        $a = $OldFiles[$name]
        $b = $NewFiles[$name]

        if (-not $a) {
            $changed += [pscustomobject]@{ Kind = 'Added'; Name = $name; OldSize = $null; NewSize = $b.Size }
            continue
        }
        if (-not $b) {
            $changed += [pscustomobject]@{ Kind = 'Removed'; Name = $name; OldSize = $a.Size; NewSize = $null }
            continue
        }
        if ($a.Hash -ne $b.Hash) {
            $changed += [pscustomobject]@{ Kind = 'Changed'; Name = $name; OldSize = $a.Size; NewSize = $b.Size }
        }
    }

    return $changed
}

function Format-DurationDelta([Nullable[int]]$OldMs, [Nullable[int]]$NewMs) {
    if ($null -eq $OldMs -or $null -eq $NewMs) {
        return $null
    }

    $delta = $NewMs - $OldMs
    if ($delta -eq 0) {
        return '0'
    }
    if ($delta -gt 0) {
        return "+$delta"
    }
    return "$delta"
}

$oldRoot = Resolve-RSExistingPath -Path $OldRun
$newRoot = Resolve-RSExistingPath -Path $NewRun

Write-Host "Old: $oldRoot"
Write-Host "New: $newRoot"

$oldResultsPath = Find-RSTestRunResultsJson -RunRoot $oldRoot -Suite $Suite -MissingPolicy Throw
$newResultsPath = Find-RSTestRunResultsJson -RunRoot $newRoot -Suite $Suite -MissingPolicy Throw

$oldResults = Read-RSJsonFile -Path $oldResultsPath
$newResults = Read-RSJsonFile -Path $newResultsPath

Write-Host ""
Write-Host "Summary"
Write-Host "-------"

$oldPassed = if ($null -eq $oldResults.passed) { 0 } else { [int]$oldResults.passed }
$oldFailed = if ($null -eq $oldResults.failed) { 0 } else { [int]$oldResults.failed }
$oldSkipped = if ($null -eq $oldResults.skipped) { 0 } else { [int]$oldResults.skipped }
$oldMs = if ($null -eq $oldResults.duration_ms) { $null } else { [int]$oldResults.duration_ms }
$oldTotal = Get-RSTestRunTotalCases -Results $oldResults
$oldColor = Get-RSTestRunSummaryColor $oldFailed $oldSkipped

$newPassed = if ($null -eq $newResults.passed) { 0 } else { [int]$newResults.passed }
$newFailed = if ($null -eq $newResults.failed) { 0 } else { [int]$newResults.failed }
$newSkipped = if ($null -eq $newResults.skipped) { 0 } else { [int]$newResults.skipped }
$newMs = if ($null -eq $newResults.duration_ms) { $null } else { [int]$newResults.duration_ms }
$newTotal = Get-RSTestRunTotalCases -Results $newResults
$newColor = Get-RSTestRunSummaryColor $newFailed $newSkipped

$oldSuite = $oldResults.suite
$newSuite = $newResults.suite
if ([string]::IsNullOrWhiteSpace($oldSuite)) { $oldSuite = '<unknown>' }
if ([string]::IsNullOrWhiteSpace($newSuite)) { $newSuite = '<unknown>' }

Write-Host ("Old: suite={0} cases={1} passed={2} ({3}) failed={4} ({5}) skipped={6} ({7}) duration_ms={8}" -f `
        $oldSuite, $oldTotal, $oldPassed, (Format-RSPercent $oldPassed $oldTotal), $oldFailed, (Format-RSPercent $oldFailed $oldTotal), $oldSkipped, (Format-RSPercent $oldSkipped $oldTotal), $oldMs) `
    -ForegroundColor $oldColor

Write-Host ("New: suite={0} cases={1} passed={2} ({3}) failed={4} ({5}) skipped={6} ({7}) duration_ms={8}" -f `
        $newSuite, $newTotal, $newPassed, (Format-RSPercent $newPassed $newTotal), $newFailed, (Format-RSPercent $newFailed $newTotal), $newSkipped, (Format-RSPercent $newSkipped $newTotal), $newMs) `
    -ForegroundColor $newColor

$durDelta = Format-DurationDelta $oldMs $newMs
$durDeltaPct = Format-RSDeltaPercent $oldMs $newMs
$durColor = if ($null -eq $durDelta -or $durDelta -eq '0') { 'Gray' } elseif ($durDelta.StartsWith('+')) { 'Red' } else { 'Green' }
Write-Host ("delta_duration_ms={0} ({1})" -f $durDelta, $durDeltaPct) -ForegroundColor $durColor

$caseDelta = $newTotal - $oldTotal
$caseDeltaText = if ($caseDelta -eq 0) { '0' } elseif ($caseDelta -gt 0) { "+$caseDelta" } else { "$caseDelta" }
$caseDeltaPct = Format-RSDeltaPercent $oldTotal $newTotal
Write-Host ("delta_cases={0} ({1})" -f $caseDeltaText, $caseDeltaPct)

Write-Host ""
Write-Host "Files"
Write-Host "-----"

$oldFiles = Get-RunFileMap $oldRoot
$newFiles = Get-RunFileMap $newRoot
$fileChanges = Compare-RunFiles $oldFiles $newFiles

if ($fileChanges.Count -eq 0) {
    Write-Host "No file-level changes."
} else {
    $fileChanges | Sort-Object Kind, Name | Format-Table -AutoSize
}

Write-Host ""
Write-Host "Cases"
Write-Host "-----"

$oldCasesByName = @{}
foreach ($c in ($oldResults.cases | ForEach-Object { $_ })) {
    if ($c -and $c.name) {
        $oldCasesByName[$c.name] = $c
    }
}

$newCasesByName = @{}
foreach ($c in ($newResults.cases | ForEach-Object { $_ })) {
    if ($c -and $c.name) {
        $newCasesByName[$c.name] = $c
    }
}

$allCaseNames = @($oldCasesByName.Keys + $newCasesByName.Keys) | Sort-Object -Unique

$caseChanges = @()
foreach ($name in $allCaseNames) {
    $a = $oldCasesByName[$name]
    $b = $newCasesByName[$name]

    if (-not $a) {
        $caseChanges += [pscustomobject]@{
            Kind       = 'Added'
            Name       = $name
            OldStatus  = $null
            NewStatus  = $b.status
            OldMs      = $null
            NewMs      = $b.duration_ms
            DeltaMs    = $null
            DeltaPct   = $null
            NewReason  = $b.reason
        }
        continue
    }
    if (-not $b) {
        $caseChanges += [pscustomobject]@{
            Kind       = 'Removed'
            Name       = $name
            OldStatus  = $a.status
            NewStatus  = $null
            OldMs      = $a.duration_ms
            NewMs      = $null
            DeltaMs    = $null
            DeltaPct   = $null
            NewReason  = $null
        }
        continue
    }

    $statusChanged = ($a.status -ne $b.status)
    $reasonChanged = ($a.PSObject.Properties.Match('reason').Count -gt 0) -or ($b.PSObject.Properties.Match('reason').Count -gt 0)
    $reasonChanged = $reasonChanged -and (($a.reason -ne $b.reason))

    $oldCaseMs = if ($null -eq $a.duration_ms) { $null } else { [int]$a.duration_ms }
    $newCaseMs = if ($null -eq $b.duration_ms) { $null } else { [int]$b.duration_ms }

    $delta = Format-DurationDelta $oldCaseMs $newCaseMs
    $durationChanged = ($delta -ne '0')

    if ($statusChanged -or $reasonChanged -or $durationChanged) {
        $caseChanges += [pscustomobject]@{
            Kind       = 'Changed'
            Name       = $name
            OldStatus  = $a.status
            NewStatus  = $b.status
            OldMs      = $oldCaseMs
            NewMs      = $newCaseMs
            DeltaMs    = $delta
            DeltaPct   = Format-RSDeltaPercent $oldCaseMs $newCaseMs
            NewReason  = $b.reason
        }
    }
}

if ($caseChanges.Count -eq 0) {
    Write-Host "No case-level changes."
} else {
    $caseChanges | Sort-Object Kind, Name | Format-Table -AutoSize
}

if ($ShowTraceDiff) {
    $oldTracePath = Find-RSTestRunTraceFile -RunRoot $oldRoot -Suite $Suite -MissingPolicy ReturnNull
    $newTracePath = Find-RSTestRunTraceFile -RunRoot $newRoot -Suite $Suite -MissingPolicy ReturnNull

    if (-not $oldTracePath -or -not $newTracePath) {
        Write-Host ""
        Write-Host "Trace"
        Write-Host "-----"
        Write-Host "Trace diff requested but trace.txt was not found in one or both run folders."
        exit 0
    }

    $oldLines = Get-Content -LiteralPath $oldTracePath
    $newLines = Get-Content -LiteralPath $newTracePath

    Write-Host ""
    Write-Host "Trace"
    Write-Host "-----"
    Write-Host ("Old trace lines: {0}" -f $oldLines.Count)
    Write-Host ("New trace lines: {0}" -f $newLines.Count)

    $diff = Compare-Object -ReferenceObject $oldLines -DifferenceObject $newLines -IncludeEqual:$false
    if (-not $diff -or $diff.Count -eq 0) {
        Write-Host "No line-level differences."
        exit 0
    }

    Write-Host ("Diff lines: {0} (showing first {1})" -f $diff.Count, $MaxTraceDiffLines)
    $diff | Select-Object -First $MaxTraceDiffLines | Format-Table -AutoSize
}
