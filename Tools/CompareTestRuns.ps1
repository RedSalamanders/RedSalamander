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

function Resolve-ExistingPath([string]$PathText) {
    $resolved = Resolve-Path -LiteralPath $PathText -ErrorAction Stop
    return $resolved.Path
}

function Find-ResultsJson([string]$RunRoot) {
    $autoCandidates = @(
        (Join-Path $RunRoot 'fileops_results.json'),
        (Join-Path $RunRoot 'compare_results.json'),
        (Join-Path $RunRoot 'commands_results.json'),
        (Join-Path $RunRoot 'selftest_run_results.json'),
        (Join-Path $RunRoot 'results.json')
    )

    $candidates = switch ($Suite) {
        'FileOps' { @((Join-Path $RunRoot 'fileops_results.json')) }
        'CompareDirectories' { @((Join-Path $RunRoot 'compare_results.json')) }
        'Commands' { @((Join-Path $RunRoot 'commands_results.json')) }
        'SelfTest' { @((Join-Path $RunRoot 'selftest_run_results.json'), (Join-Path $RunRoot 'results.json')) }
        'Auto' { $autoCandidates }
        default { $autoCandidates }
    }

    foreach ($p in $candidates) {
        if (Test-Path -LiteralPath $p) {
            return $p
        }
    }

    throw "Could not find results JSON in run folder: $RunRoot"
}

function Find-TraceTxt([string]$RunRoot) {
    $autoCandidates = @(
        (Join-Path $RunRoot 'fileops_trace.txt'),
        (Join-Path $RunRoot 'compare_trace.txt'),
        (Join-Path $RunRoot 'commands_trace.txt'),
        (Join-Path $RunRoot 'selftest_run_trace.txt'),
        (Join-Path $RunRoot 'trace.txt')
    )

    $candidates = switch ($Suite) {
        'FileOps' { @((Join-Path $RunRoot 'fileops_trace.txt')) }
        'CompareDirectories' { @((Join-Path $RunRoot 'compare_trace.txt')) }
        'Commands' { @((Join-Path $RunRoot 'commands_trace.txt')) }
        'SelfTest' { @((Join-Path $RunRoot 'selftest_run_trace.txt'), (Join-Path $RunRoot 'trace.txt')) }
        'Auto' { $autoCandidates }
        default { $autoCandidates }
    }

    foreach ($p in $candidates) {
        if (Test-Path -LiteralPath $p) {
            return $p
        }
    }

    return $null
}

function Load-Json([string]$PathText) {
    $raw = Get-Content -LiteralPath $PathText -Raw
    return $raw | ConvertFrom-Json -Depth 64
}

function Get-RunFileMap([string]$RunRoot) {
    $map = @{}
    $root = (Resolve-Path -LiteralPath $RunRoot).Path

    Get-ChildItem -LiteralPath $root -File | ForEach-Object {
        $rel = $_.Name
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

function Get-TotalCases($Results) {
    $passed = if ($null -eq $Results.passed) { 0 } else { [int]$Results.passed }
    $failed = if ($null -eq $Results.failed) { 0 } else { [int]$Results.failed }
    $skipped = if ($null -eq $Results.skipped) { 0 } else { [int]$Results.skipped }

    if ($Results.PSObject.Properties.Match('cases').Count -gt 0 -and $Results.cases -and $Results.cases.Count -gt 0) {
        return [int]$Results.cases.Count
    }

    $total = $passed + $failed + $skipped
    return $total
}

function Format-Percent([int]$Numerator, [int]$Denominator, [int]$Decimals = 1) {
    if ($Denominator -le 0) {
        return ''
    }

    $pct = 100.0 * $Numerator / [double]$Denominator
    $fmt = '0.' + ('0' * $Decimals)
    return ($pct.ToString($fmt, [System.Globalization.CultureInfo]::InvariantCulture) + '%')
}

function Format-DeltaPercent([Nullable[int]]$OldValue, [Nullable[int]]$NewValue, [int]$Decimals = 1) {
    if ($null -eq $OldValue -or $null -eq $NewValue) {
        return $null
    }
    if ($OldValue -le 0) {
        return $null
    }

    $delta = $NewValue - $OldValue
    if ($delta -eq 0) {
        return '0%'
    }

    $pct = 100.0 * $delta / [double]$OldValue
    $fmt = '0.' + ('0' * $Decimals)
    $text = $pct.ToString($fmt, [System.Globalization.CultureInfo]::InvariantCulture) + '%'
    if ($pct -gt 0) {
        return "+$text"
    }
    return $text
}

function Get-RunSummaryColor([int]$Failed, [int]$Skipped) {
    if ($Failed -gt 0) {
        return 'Red'
    }
    if ($Skipped -gt 0) {
        return 'Yellow'
    }
    return 'Green'
}

$oldRoot = Resolve-ExistingPath $OldRun
$newRoot = Resolve-ExistingPath $NewRun

Write-Host "Old: $oldRoot"
Write-Host "New: $newRoot"

$oldResultsPath = Find-ResultsJson $oldRoot
$newResultsPath = Find-ResultsJson $newRoot

$oldResults = Load-Json $oldResultsPath
$newResults = Load-Json $newResultsPath

Write-Host ""
Write-Host "Summary"
Write-Host "-------"

$oldPassed = if ($null -eq $oldResults.passed) { 0 } else { [int]$oldResults.passed }
$oldFailed = if ($null -eq $oldResults.failed) { 0 } else { [int]$oldResults.failed }
$oldSkipped = if ($null -eq $oldResults.skipped) { 0 } else { [int]$oldResults.skipped }
$oldMs = if ($null -eq $oldResults.duration_ms) { $null } else { [int]$oldResults.duration_ms }
$oldTotal = Get-TotalCases $oldResults
$oldColor = Get-RunSummaryColor $oldFailed $oldSkipped

$newPassed = if ($null -eq $newResults.passed) { 0 } else { [int]$newResults.passed }
$newFailed = if ($null -eq $newResults.failed) { 0 } else { [int]$newResults.failed }
$newSkipped = if ($null -eq $newResults.skipped) { 0 } else { [int]$newResults.skipped }
$newMs = if ($null -eq $newResults.duration_ms) { $null } else { [int]$newResults.duration_ms }
$newTotal = Get-TotalCases $newResults
$newColor = Get-RunSummaryColor $newFailed $newSkipped

$oldSuite = $oldResults.suite
$newSuite = $newResults.suite
if ([string]::IsNullOrWhiteSpace($oldSuite)) { $oldSuite = '<unknown>' }
if ([string]::IsNullOrWhiteSpace($newSuite)) { $newSuite = '<unknown>' }

Write-Host ("Old: suite={0} cases={1} passed={2} ({3}) failed={4} ({5}) skipped={6} ({7}) duration_ms={8}" -f `
        $oldSuite, $oldTotal, $oldPassed, (Format-Percent $oldPassed $oldTotal), $oldFailed, (Format-Percent $oldFailed $oldTotal), $oldSkipped, (Format-Percent $oldSkipped $oldTotal), $oldMs) `
    -ForegroundColor $oldColor

Write-Host ("New: suite={0} cases={1} passed={2} ({3}) failed={4} ({5}) skipped={6} ({7}) duration_ms={8}" -f `
        $newSuite, $newTotal, $newPassed, (Format-Percent $newPassed $newTotal), $newFailed, (Format-Percent $newFailed $newTotal), $newSkipped, (Format-Percent $newSkipped $newTotal), $newMs) `
    -ForegroundColor $newColor

$durDelta = Format-DurationDelta $oldMs $newMs
$durDeltaPct = Format-DeltaPercent $oldMs $newMs
$durColor = if ($null -eq $durDelta -or $durDelta -eq '0') { 'Gray' } elseif ($durDelta.StartsWith('+')) { 'Red' } else { 'Green' }
Write-Host ("Δduration_ms={0} ({1})" -f $durDelta, $durDeltaPct) -ForegroundColor $durColor

$caseDelta = $newTotal - $oldTotal
$caseDeltaText = if ($caseDelta -eq 0) { '0' } elseif ($caseDelta -gt 0) { "+$caseDelta" } else { "$caseDelta" }
$caseDeltaPct = Format-DeltaPercent $oldTotal $newTotal
Write-Host ("Δcases={0} ({1})" -f $caseDeltaText, $caseDeltaPct)

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
            DeltaPct   = Format-DeltaPercent $oldCaseMs $newCaseMs
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
    $oldTracePath = Find-TraceTxt $oldRoot
    $newTracePath = Find-TraceTxt $newRoot

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
