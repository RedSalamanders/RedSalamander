<#
.SYNOPSIS
    Analyzes archived selftest run folders under Specs\TestRuns and reports timing trends.

.DESCRIPTION
    Given an "Area" folder (e.g. Specs\TestRuns\<ComputerHashName>\FileOps) containing timestamped run
    subfolders, this script:

      - Loads the suite results JSON for each run (e.g. fileops_results.json).
      - Prints a timeline (pass/fail + duration + case-count).
      - Prints a "last run" report: slowest cases, regressions vs previous run,
        and long-term case timing deltas (first -> last).

    Notes:
      - Archived run folders are created automatically by the selftest harness (Debug build from a repo checkout),
        or by manually copying %LOCALAPPDATA%\RedSalamander\SelfTest\last_run\ artifacts.
      - This script is intentionally "read-only" and does not modify any run folders.

.PARAMETER RunsRoot
    Path to an area folder containing run subfolders (timestamped directories).
    Example: Specs\TestRuns\<ComputerHashName>\FileOps

.PARAMETER Suite
    Which suite results artifact to analyze. In 'Auto' mode, the script picks the first
    matching results JSON in this order:
        fileops_results.json, compare_results.json, commands_results.json,
        selftest_run_results.json, results.json

.PARAMETER TopN
    Number of cases to show in the slowest/regression tables. Default: 10.

.PARAMETER OutMarkdown
    Optional path to write a markdown report (UTF-8). If omitted, prints to the console only.

.EXAMPLE
    .\Tools\AnalyzeTestRuns.ps1 Specs\TestRuns\<ComputerHashName>\FileOps

.EXAMPLE
    .\Tools\AnalyzeTestRuns.ps1 Specs\TestRuns\<ComputerHashName>\FileOps -TopN 15 -OutMarkdown .\analysis.md
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$RunsRoot,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Auto', 'FileOps', 'CompareDirectories', 'Commands', 'SelfTest')]
    [string]$Suite = 'Auto',

    [Parameter(Mandatory = $false)]
    [int]$TopN = 10,

    [Parameter(Mandatory = $false)]
    [string]$OutMarkdown
)

$ErrorActionPreference = 'Stop'

function Resolve-ExistingPath([string]$PathText) {
    $resolved = Resolve-Path -LiteralPath $PathText -ErrorAction Stop
    return $resolved.Path
}

function TryParse-TimestampFromFolderName([string]$FolderName) {
    if ([string]::IsNullOrWhiteSpace($FolderName) -or $FolderName.Length -lt 17) {
        return $null
    }

    $prefix = $FolderName.Substring(0, 17)
    try {
        return [datetime]::ParseExact($prefix, 'yyyy-MM-dd_HHmmss', [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        return $null
    }
}

function Load-Json([string]$PathText) {
    $raw = Get-Content -LiteralPath $PathText -Raw
    return $raw | ConvertFrom-Json -Depth 64
}

function Load-EnvMap([string]$RunRoot) {
    $envPath = Join-Path $RunRoot 'env.txt'
    if (-not (Test-Path -LiteralPath $envPath)) {
        return @{}
    }

    $map = @{}
    $lines = Get-Content -LiteralPath $envPath -ErrorAction Stop
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $idx = $line.IndexOf(':')
        if ($idx -lt 0) {
            continue
        }
        $key = $line.Substring(0, $idx).Trim()
        $value = $line.Substring($idx + 1).Trim()
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $map[$key] = $value
        }
    }

    return $map
}

function Find-ResultsJson([string]$RunRoot, [string]$SuiteName) {
    $autoCandidates = @(
        (Join-Path $RunRoot 'fileops_results.json'),
        (Join-Path $RunRoot 'compare_results.json'),
        (Join-Path $RunRoot 'commands_results.json'),
        (Join-Path $RunRoot 'selftest_run_results.json'),
        (Join-Path $RunRoot 'results.json')
    )

    $candidates = switch ($SuiteName) {
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

    return $null
}

function Get-ShortCommit([string]$Commit) {
    if ([string]::IsNullOrWhiteSpace($Commit)) {
        return ''
    }
    if ($Commit.Length -le 8) {
        return $Commit
    }
    return $Commit.Substring(0, 8)
}

function Get-CasesByName($Results) {
    $map = @{}
    foreach ($c in ($Results.cases | ForEach-Object { $_ })) {
        if ($c -and $c.name) {
            $map[$c.name] = $c
        }
    }
    return $map
}

function Sum-CaseDurationsMs([hashtable]$CasesByName, [string[]]$CaseNames) {
    $sum = 0
    foreach ($name in $CaseNames) {
        $c = $CasesByName[$name]
        if (-not $c) {
            continue
        }
        if ($null -eq $c.duration_ms) {
            continue
        }
        $sum += [int]$c.duration_ms
    }
    return $sum
}

function Get-TotalCases([int]$Passed, [int]$Failed, [int]$Skipped, [int]$FallbackCaseCount) {
    $total = $Passed + $Failed + $Skipped
    if ($total -gt 0) {
        return $total
    }
    return $FallbackCaseCount
}

function Format-Percent([int]$Numerator, [int]$Denominator, [int]$Decimals = 1) {
    if ($Denominator -le 0) {
        return ''
    }

    $pct = 100.0 * $Numerator / [double]$Denominator
    $fmt = '0.' + ('0' * $Decimals)
    return ($pct.ToString($fmt, [System.Globalization.CultureInfo]::InvariantCulture) + '%')
}

function Format-SignedInt([Nullable[int]]$Value) {
    if ($null -eq $Value) {
        return ''
    }
    if ($Value -eq 0) {
        return '0'
    }
    if ($Value -gt 0) {
        return "+$Value"
    }
    return "$Value"
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

function Get-DeltaColor([Nullable[int]]$DeltaMs) {
    if ($null -eq $DeltaMs -or $DeltaMs -eq 0) {
        return 'Gray'
    }
    if ($DeltaMs -gt 0) {
        return 'Red'
    }
    return 'Green'
}

function Write-SectionHeader([string]$Title) {
    Write-Host ""
    Write-Host $Title -ForegroundColor Yellow
    Write-Host ('-' * $Title.Length) -ForegroundColor Yellow
}

function Get-BlockChar([int]$Value, [int]$Min, [int]$Max) {
    $blocks = @('▁', '▂', '▃', '▄', '▅', '▆', '▇', '█')
    if ($Max -le $Min) {
        return $blocks[$blocks.Count - 1]
    }

    $scaled = ($Value - $Min) * 7.0 / ($Max - $Min)
    $idx = [int][Math]::Floor($scaled)
    if ($idx -lt 0) { $idx = 0 }
    if ($idx -gt 7) { $idx = 7 }
    return $blocks[$idx]
}

function Build-Sparkline([object[]]$Durations, [int]$Min, [int]$Max) {
    $s = ''
    foreach ($d in $Durations) {
        if ($null -eq $d) {
            $s += '.'
            continue
        }
        $s += (Get-BlockChar ([int]$d) $Min $Max)
    }
    return $s
}

function Write-SparklineColored([object[]]$Durations, [int]$Min, [int]$Max) {
    $prev = $null
    $isFirst = $true

    foreach ($d in $Durations) {
        if ($null -eq $d) {
            Write-Host '.' -NoNewline -ForegroundColor DarkGray
            $prev = $null
            $isFirst = $false
            continue
        }

        $ms = [int]$d
        $ch = Get-BlockChar $ms $Min $Max

        $color = $null
        if ($isFirst) {
            $color = 'Gray'
        } elseif ($null -eq $prev) {
            $color = 'Yellow'
        } else {
            $color = Get-DeltaColor ($ms - [int]$prev)
        }

        Write-Host $ch -NoNewline -ForegroundColor $color
        $prev = $ms
        $isFirst = $false
    }
}

function Write-Cell([string]$Text, [int]$Width, [ValidateSet('Left', 'Right')] [string]$Align, [string]$Color, [switch]$NoNewline) {
    $t = if ($null -eq $Text) { '' } else { [string]$Text }
    if ($t.Length -gt $Width) {
        $t = $t.Substring(0, $Width)
    }

    if ($Align -eq 'Right') {
        $t = $t.PadLeft($Width)
    } else {
        $t = $t.PadRight($Width)
    }

    if ([string]::IsNullOrWhiteSpace($Color)) {
        Write-Host $t -NoNewline:$NoNewline
        return
    }

    Write-Host $t -NoNewline:$NoNewline -ForegroundColor $Color
}

function Write-ColorTable($Rows, $Columns) {
    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Host "_No data._"
        return
    }

    # Compute column widths.
    foreach ($c in $Columns) {
        $max = ([string]$c.Header).Length
        foreach ($r in $Rows) {
            $v = & $c.GetText $r
            $s = if ($null -eq $v) { '' } else { [string]$v }
            if ($s.Length -gt $max) { $max = $s.Length }
        }
        if ($c.PSObject.Properties.Match('Width').Count -eq 0) {
            $c | Add-Member -NotePropertyName 'Width' -NotePropertyValue $max
        } else {
            $c.Width = $max
        }
    }

    # Header row
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        $c = $Columns[$i]
        $sep = if ($i -lt ($Columns.Count - 1)) { '  ' } else { '' }
        Write-Cell $c.Header $c.Width 'Left' 'Yellow' -NoNewline
        if ($sep) { Write-Host $sep -NoNewline }
    }
    Write-Host ""

    # Data rows
    foreach ($r in $Rows) {
        for ($i = 0; $i -lt $Columns.Count; $i++) {
            $c = $Columns[$i]
            $text = & $c.GetText $r
            $color = $null
            if ($c.PSObject.Properties.Match('GetColor').Count -gt 0 -and $c.GetColor) {
                $color = & $c.GetColor $r
            }

            $sep = if ($i -lt ($Columns.Count - 1)) { '  ' } else { '' }
            Write-Cell $text $c.Width $c.Align $color -NoNewline
            if ($sep) { Write-Host $sep -NoNewline }
        }
        Write-Host ""
    }
}

function Write-MarkdownTable([string]$Title, $Rows, [string[]]$Columns) {
    if (-not $Rows -or $Rows.Count -eq 0) {
        return @("## $Title", '', '_No data._', '')
    }

    $lines = @("## $Title", '')
    $lines += '| ' + ($Columns -join ' | ') + ' |'
    $lines += '| ' + (($Columns | ForEach-Object { '---' }) -join ' | ') + ' |'

    foreach ($r in $Rows) {
        $cells = @()
        foreach ($col in $Columns) {
            $v = $r.$col
            if ($null -eq $v) { $v = '' }
            $cells += ("$v" -replace '\|', '\\|')
        }
        $lines += '| ' + ($cells -join ' | ') + ' |'
    }

    $lines += ''
    return $lines
}

$runsRootResolved = Resolve-ExistingPath $RunsRoot

$runDirs = Get-ChildItem -LiteralPath $runsRootResolved -Directory | ForEach-Object { $_ }
if (-not $runDirs -or $runDirs.Count -eq 0) {
    throw "No run subfolders found under: $runsRootResolved"
}

$runs = @()
foreach ($dir in $runDirs) {
    $resultsPath = Find-ResultsJson $dir.FullName $Suite
    if (-not $resultsPath) {
        continue
    }

    $results = Load-Json $resultsPath
    $env = Load-EnvMap $dir.FullName

    $ts = TryParse-TimestampFromFolderName $dir.Name
    $runs += [pscustomobject]@{
        Folder   = $dir.Name
        Time     = $ts
        Commit   = ($env['git_commit'])
        Branch   = ($env['git_branch'])
        Machine  = ($env['machine'])
        Scenario = ($env['scenario'])
        Suite    = ($results.suite)
        Ms       = ([int]$results.duration_ms)
        Passed   = ([int]$results.passed)
        Failed   = ([int]$results.failed)
        Skipped  = ([int]$results.skipped)
        Cases    = ($results.cases.Count)
        Results  = $results
        RunRoot  = $dir.FullName
    }
}

if (-not $runs -or $runs.Count -eq 0) {
    throw "No runs found under $runsRootResolved for suite '$Suite' (missing results JSON)."
}

$runs = $runs | Sort-Object @{ Expression = { if ($null -eq $_.Time) { [datetime]::MinValue } else { $_.Time } } }, Folder

$base = $runs[0]
$baseCases = Get-CasesByName $base.Results
$baseCaseNames = @($baseCases.Keys | Sort-Object)

$timeline = @()
for ($i = 0; $i -lt $runs.Count; $i++) {
    $r = $runs[$i]
    $casesByName = Get-CasesByName $r.Results
    $baseMs = Sum-CaseDurationsMs $casesByName $baseCaseNames
    $totalCases = Get-TotalCases $r.Passed $r.Failed $r.Skipped $r.Cases

    $newCaseNames = @($casesByName.Keys | Where-Object { -not $baseCases.ContainsKey($_) } | Sort-Object)
    $newCasesMs = Sum-CaseDurationsMs $casesByName $newCaseNames
    $newCasesText = ''
    if ($newCaseNames.Count -gt 0) {
        $newCasesText = ("{0} ({1})" -f $newCasesMs, (Format-Percent $newCasesMs $r.Ms))
    } else {
        $newCasesText = ("0 ({0})" -f (Format-Percent 0 $r.Ms))
    }

    $prevSuiteDelta = $null
    $prevBaseDelta = $null
    $prevSuiteDeltaText = ''
    $prevBaseDeltaText = ''
    if ($i -gt 0) {
        $prev = $runs[$i - 1]
        $prevCasesByName = Get-CasesByName $prev.Results
        $prevBaseMs = Sum-CaseDurationsMs $prevCasesByName $baseCaseNames

        $prevSuiteDelta = $r.Ms - $prev.Ms
        $prevBaseDelta = $baseMs - $prevBaseMs

        $prevSuiteDeltaText = ("{0} ({1})" -f (Format-SignedInt $prevSuiteDelta), (Format-DeltaPercent $prev.Ms $r.Ms))
        $prevBaseDeltaText = ("{0} ({1})" -f (Format-SignedInt $prevBaseDelta), (Format-DeltaPercent $prevBaseMs $baseMs))
    }

    $timeline += [pscustomobject]@{
        Folder     = $r.Folder
        Commit     = (Get-ShortCommit $r.Commit)
        SuiteMs    = $r.Ms
        BaseMs     = $baseMs
        NewCases   = $newCasesText
        PassPct    = (Format-Percent $r.Passed $totalCases)
        DeltaSuiteMs = $prevSuiteDelta
        DeltaSuite = $prevSuiteDeltaText
        DeltaBaseMs  = $prevBaseDelta
        DeltaBase  = $prevBaseDeltaText
        Cases      = $r.Cases
        Passed     = $r.Passed
        Failed     = $r.Failed
        Skipped    = $r.Skipped
    }
}

$last = $runs[$runs.Count - 1]
$prev = $null
if ($runs.Count -ge 2) {
    $prev = $runs[$runs.Count - 2]
}

Write-Host "Runs root: $runsRootResolved"
Write-Host ("Suite selection: {0}" -f $Suite)
Write-Host ("Runs analyzed: {0}" -f $runs.Count)

Write-SectionHeader "Timeline"
Write-Host "SuiteMs = reported duration_ms; BaseMs = sum of first-run case durations in this run"

$timelineColumns = @(
    [pscustomobject]@{ Header = 'Folder'; Align = 'Left'; GetText = { param($r) $r.Folder } }
    [pscustomobject]@{ Header = 'Commit'; Align = 'Left'; GetText = { param($r) $r.Commit } }
    [pscustomobject]@{ Header = 'SuiteMs'; Align = 'Right'; GetText = { param($r) $r.SuiteMs } }
    [pscustomobject]@{ Header = 'BaseMs'; Align = 'Right'; GetText = { param($r) $r.BaseMs } }
    [pscustomobject]@{ Header = 'NewCases'; Align = 'Right'; GetText = { param($r) $r.NewCases } }
    [pscustomobject]@{ Header = 'PassPct'; Align = 'Right'; GetText = { param($r) $r.PassPct }; GetColor = { param($r) Get-RunSummaryColor ([int]$r.Failed) ([int]$r.Skipped) } }
    [pscustomobject]@{ Header = 'ΔSuite'; Align = 'Right'; GetText = { param($r) $r.DeltaSuite }; GetColor = { param($r) Get-DeltaColor $r.DeltaSuiteMs } }
    [pscustomobject]@{ Header = 'ΔBase'; Align = 'Right'; GetText = { param($r) $r.DeltaBase }; GetColor = { param($r) Get-DeltaColor $r.DeltaBaseMs } }
    [pscustomobject]@{ Header = 'Cases'; Align = 'Right'; GetText = { param($r) $r.Cases } }
    [pscustomobject]@{ Header = 'Pass'; Align = 'Right'; GetText = { param($r) $r.Passed } }
    [pscustomobject]@{ Header = 'Fail'; Align = 'Right'; GetText = { param($r) $r.Failed } }
    [pscustomobject]@{ Header = 'Skip'; Align = 'Right'; GetText = { param($r) $r.Skipped } }
)
Write-ColorTable $timeline $timelineColumns

Write-SectionHeader "Last run"
$lastTotalCases = Get-TotalCases $last.Passed $last.Failed $last.Skipped $last.Cases
$lastColor = Get-RunSummaryColor $last.Failed $last.Skipped
Write-Host ("Folder: {0}" -f $last.Folder)
Write-Host ("Commit: {0}" -f $last.Commit)
Write-Host ("Passed/Failed/Skipped: {0}/{1}/{2} (pass {3})" -f $last.Passed, $last.Failed, $last.Skipped, (Format-Percent $last.Passed $lastTotalCases)) -ForegroundColor $lastColor
Write-Host ("SuiteMs: {0}" -f $last.Ms)

$lastCases = $last.Results.cases | ForEach-Object { $_ }
Write-SectionHeader ("Slowest cases (Top {0})" -f $TopN)
$lastCases | Sort-Object duration_ms -Descending | Select-Object -First $TopN `
    name, status, duration_ms, @{ n = 'pct_of_suite'; e = { Format-Percent ([int]$_.duration_ms) $last.Ms } }, reason `
    | Format-Table -AutoSize

if ($prev) {
    $prevCasesByName = Get-CasesByName $prev.Results
    $lastCasesByName = Get-CasesByName $last.Results

    $commonNames = @($prevCasesByName.Keys | Where-Object { $lastCasesByName.ContainsKey($_) } | Sort-Object)
    $addedNames = @($lastCasesByName.Keys | Where-Object { -not $prevCasesByName.ContainsKey($_) } | Sort-Object)
    $removedNames = @($prevCasesByName.Keys | Where-Object { -not $lastCasesByName.ContainsKey($_) } | Sort-Object)

    $commonPrevMs = Sum-CaseDurationsMs $prevCasesByName $commonNames
    $commonLastMs = Sum-CaseDurationsMs $lastCasesByName $commonNames
    $addedLastMs = Sum-CaseDurationsMs $lastCasesByName $addedNames
    $removedPrevMs = Sum-CaseDurationsMs $prevCasesByName $removedNames

    $changes = @()
    foreach ($name in $commonNames) {
        $a = $prevCasesByName[$name]
        $b = $lastCasesByName[$name]
        if (-not $a -or -not $b) {
            continue
        }

        $oldMs = if ($null -eq $a.duration_ms) { $null } else { [int]$a.duration_ms }
        $newMs = if ($null -eq $b.duration_ms) { $null } else { [int]$b.duration_ms }

        $delta = $null
        if ($null -ne $oldMs -and $null -ne $newMs) {
            $delta = $newMs - $oldMs
        }

        $statusChanged = ($a.status -ne $b.status)
        $durationChanged = ($null -ne $delta -and $delta -ne 0)
        $reasonChanged = (($a.PSObject.Properties.Match('reason').Count -gt 0) -or ($b.PSObject.Properties.Match('reason').Count -gt 0)) -and ($a.reason -ne $b.reason)

        if ($statusChanged -or $durationChanged -or $reasonChanged) {
            $changes += [pscustomobject]@{
                Name      = $name
                OldStatus = $a.status
                NewStatus = $b.status
                OldMs     = $oldMs
                NewMs     = $newMs
                DeltaMs   = $delta
                DeltaPct  = Format-DeltaPercent $oldMs $newMs
            }
        }
    }

    Write-SectionHeader ("Δ vs previous run ({0} -> {1})" -f $prev.Folder, $last.Folder)
    $suiteDeltaMs = ($last.Ms - $prev.Ms)
    Write-Host ("SuiteMs: {0} -> {1} (Δ {2} / {3})" -f $prev.Ms, $last.Ms, (Format-SignedInt $suiteDeltaMs), (Format-DeltaPercent $prev.Ms $last.Ms)) -ForegroundColor (Get-DeltaColor $suiteDeltaMs)
    Write-Host ("Cases:   {0} -> {1} (Added {2}, Removed {3})" -f $prev.Cases, $last.Cases, $addedNames.Count, $removedNames.Count)
    $commonDeltaMs = ($commonLastMs - $commonPrevMs)
    Write-Host ("Common-case sum: {0} -> {1} (Δ {2} / {3})" -f $commonPrevMs, $commonLastMs, (Format-SignedInt $commonDeltaMs), (Format-DeltaPercent $commonPrevMs $commonLastMs)) -ForegroundColor (Get-DeltaColor $commonDeltaMs)
    if ($addedNames.Count -gt 0) {
        Write-Host ("Added-case sum (new only): {0} ({1} of SuiteMs)" -f $addedLastMs, (Format-Percent $addedLastMs $last.Ms)) -ForegroundColor 'Yellow'
    }
    if ($removedNames.Count -gt 0) {
        Write-Host ("Removed-case sum (prev only): {0}" -f $removedPrevMs)
    }

    $regressions = $changes | Where-Object { $_.DeltaMs -gt 0 } | Sort-Object DeltaMs -Descending | Select-Object -First $TopN
    $improvements = $changes | Where-Object { $_.DeltaMs -lt 0 } | Sort-Object DeltaMs | Select-Object -First $TopN

    Write-SectionHeader ("Top regressions (Δms > 0, Top {0})" -f $TopN)
    if (-not $regressions -or $regressions.Count -eq 0) {
        Write-Host "None."
    } else {
        foreach ($r in ($regressions | ForEach-Object { $_ })) {
            Write-Host ("{0} ({1})  {2}  ({3} -> {4})" -f (Format-SignedInt $r.DeltaMs), $r.DeltaPct, $r.Name, $r.OldMs, $r.NewMs) -ForegroundColor 'Red'
        }
    }

    Write-SectionHeader ("Top improvements (Δms < 0, Top {0})" -f $TopN)
    if (-not $improvements -or $improvements.Count -eq 0) {
        Write-Host "None."
    } else {
        foreach ($r in ($improvements | ForEach-Object { $_ })) {
            Write-Host ("{0} ({1})  {2}  ({3} -> {4})" -f (Format-SignedInt $r.DeltaMs), $r.DeltaPct, $r.Name, $r.OldMs, $r.NewMs) -ForegroundColor 'Green'
        }
    }

    if ($addedNames.Count -gt 0) {
        Write-SectionHeader "Added cases"
        $addedNames | ForEach-Object { Write-Host "- $_" }
    }

    if ($removedNames.Count -gt 0) {
        Write-SectionHeader "Removed cases"
        $removedNames | ForEach-Object { Write-Host "- $_" }
    }
}

Write-SectionHeader "Long-term case timing deltas (first -> last)"

$allCaseNames = @()
foreach ($r in $runs) {
    foreach ($c in ($r.Results.cases | ForEach-Object { $_ })) {
        if ($c -and $c.name) {
            $allCaseNames += $c.name
        }
    }
}
$allCaseNames = $allCaseNames | Sort-Object -Unique

$caseTrendRows = @()
foreach ($name in $allCaseNames) {
    $durations = @()
    foreach ($r in $runs) {
        $case = ($r.Results.cases | Where-Object { $_.name -eq $name } | Select-Object -First 1)
        if ($case -and $null -ne $case.duration_ms) {
            $durations += [int]$case.duration_ms
        } else {
            $durations += $null
        }
    }

    $firstIdx = ($durations | ForEach-Object -Begin { $i = 0 } -Process { $v = $_; $out = [pscustomobject]@{ i = $i; v = $v }; $i++; $out } | Where-Object { $null -ne $_.v } | Select-Object -First 1)
    $lastIdx = ($durations | ForEach-Object -Begin { $i = 0 } -Process { $v = $_; $out = [pscustomobject]@{ i = $i; v = $v }; $i++; $out } | Where-Object { $null -ne $_.v } | Select-Object -Last 1)

    if (-not $firstIdx -or -not $lastIdx) {
        continue
    }

    $firstMs = [int]$firstIdx.v
    $lastMs = [int]$lastIdx.v
    $deltaMs = $lastMs - $firstMs

    $seen = ($durations | Where-Object { $null -ne $_ }).Count
    $minMs = [int](($durations | Where-Object { $null -ne $_ } | Measure-Object -Minimum).Minimum)
    $maxMs = [int](($durations | Where-Object { $null -ne $_ } | Measure-Object -Maximum).Maximum)

    $caseTrendRows += [pscustomobject]@{
        Name    = $name
        Seen    = $seen
        FirstMs = $firstMs
        LastMs  = $lastMs
        DeltaMs = $deltaMs
        DeltaPct = Format-DeltaPercent $firstMs $lastMs
        MinMs   = $minMs
        MaxMs   = $maxMs
        Durations = $durations
        Sparkline = (Build-Sparkline $durations $minMs $maxMs)
    }
}

$caseTrendRows = $caseTrendRows | Sort-Object @{ Expression = { [Math]::Abs([int]$_.DeltaMs) }; Descending = $true }, Name
Write-Host "Graph: each column is a run (oldest -> newest). Height is relative per-case (min..max)."
Write-Host "Color: green=faster vs previous run, red=slower, yellow=new/missing-to-present, gray=same, darkgray=missing."
Write-Host ("Runs:  {0}" -f (@($runs | ForEach-Object { $_.Folder }) -join '  '))

$nameWidth = ($caseTrendRows | ForEach-Object { ([string]$_.Name).Length } | Measure-Object -Maximum).Maximum
if ($nameWidth -lt 24) { $nameWidth = 24 }
if ($nameWidth -gt 56) { $nameWidth = 56 }

Write-Host ""
Write-Cell 'Δms' 7 'Right' 'Yellow' -NoNewline
Write-Host " " -NoNewline
Write-Cell 'Δ%' 9 'Right' 'Yellow' -NoNewline
Write-Host " " -NoNewline
Write-Cell 'Name' $nameWidth 'Left' 'Yellow' -NoNewline
Write-Host " " -NoNewline
Write-Cell 'Graph' ($runs.Count + 2) 'Left' 'Yellow' -NoNewline
Write-Host " " -NoNewline
Write-Cell 'First->Last' 18 'Left' 'Yellow'

foreach ($r in ($caseTrendRows | ForEach-Object { $_ })) {
    $deltaColor = Get-DeltaColor ([int]$r.DeltaMs)
    Write-Cell (Format-SignedInt ([int]$r.DeltaMs)) 7 'Right' $deltaColor -NoNewline
    Write-Host " " -NoNewline
    Write-Cell $r.DeltaPct 9 'Right' $deltaColor -NoNewline
    Write-Host " " -NoNewline
    Write-Cell $r.Name $nameWidth 'Left' $null -NoNewline
    Write-Host " |" -NoNewline
    Write-SparklineColored $r.Durations $r.MinMs $r.MaxMs
    Write-Host "| " -NoNewline
    Write-Host ("{0}->{1} ms" -f $r.FirstMs, $r.LastMs)
}

if (-not [string]::IsNullOrWhiteSpace($OutMarkdown)) {
    $mdLines = @(
        ('# TestRun Analysis'),
        '',
        ('Runs root: `{0}`' -f $runsRootResolved),
        ('Suite: `{0}`' -f $Suite),
        ('Generated: {0}' -f (Get-Date -Format o)),
        ''
    )

    $mdLines += Write-MarkdownTable 'Timeline' $timeline @('Folder', 'Commit', 'SuiteMs', 'BaseMs', 'NewCases', 'PassPct', 'DeltaSuite', 'DeltaBase', 'Cases', 'Passed', 'Failed', 'Skipped')

    $slowRows = $lastCases | Sort-Object duration_ms -Descending | Select-Object -First $TopN | ForEach-Object {
        [pscustomobject]@{
            Name       = $_.name
            Status     = $_.status
            DurationMs = $_.duration_ms
            PctOfSuite = (Format-Percent ([int]$_.duration_ms) $last.Ms)
        }
    }
    $mdLines += Write-MarkdownTable ("Last run slowest cases ({0})" -f $last.Folder) $slowRows @('Name', 'Status', 'DurationMs', 'PctOfSuite')

    if ($prev) {
        $regRows = @()
        foreach ($r in ($regressions | ForEach-Object { $_ })) {
            $regRows += [pscustomobject]@{
                Name    = $r.Name
                OldMs   = $r.OldMs
                NewMs   = $r.NewMs
                DeltaMs = $r.DeltaMs
                DeltaPct = $r.DeltaPct
            }
        }
        $impRows = @()
        foreach ($r in ($improvements | ForEach-Object { $_ })) {
            $impRows += [pscustomobject]@{
                Name    = $r.Name
                OldMs   = $r.OldMs
                NewMs   = $r.NewMs
                DeltaMs = $r.DeltaMs
                DeltaPct = $r.DeltaPct
            }
        }

        $mdLines += Write-MarkdownTable ("Regressions vs previous ({0} -> {1})" -f $prev.Folder, $last.Folder) $regRows @('Name', 'OldMs', 'NewMs', 'DeltaMs', 'DeltaPct')
        $mdLines += Write-MarkdownTable ("Improvements vs previous ({0} -> {1})" -f $prev.Folder, $last.Folder) $impRows @('Name', 'OldMs', 'NewMs', 'DeltaMs', 'DeltaPct')
    }

    $mdLines += Write-MarkdownTable 'Long-term case deltas (first -> last)' $caseTrendRows @('Name', 'Seen', 'FirstMs', 'LastMs', 'DeltaMs', 'DeltaPct', 'MinMs', 'MaxMs', 'Sparkline')

    $outPathResolved = $OutMarkdown
    if (-not ([System.IO.Path]::IsPathRooted($outPathResolved))) {
        $outPathResolved = Join-Path (Get-Location).Path $OutMarkdown
    }

    $mdLines | Out-File -LiteralPath $outPathResolved -Encoding utf8
    Write-Host ""
    Write-Host ("Wrote: {0}" -f $outPathResolved)
}
