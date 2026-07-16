<#
.SYNOPSIS
    Lists, summarizes, compares, and trends archived perf runs.

.DESCRIPTION
    Reads perf JSONL artifacts archived under Specs\TestRuns\<ComputerHash>\<Area>\<Run>\perf\perf_metrics.jsonl.
    Supports:
      - listing discovered runs
      - summarizing one run
      - comparing two runs
      - showing metric evolution over time
#>

param(
    [Parameter(Position = 0)]
    [string]$RunsRoot = 'Specs\TestRuns',

    [string]$Scenario,
    [string]$Area,
    [string]$Metric,
    [switch]$FolderViewPreset,
    [string[]]$CompareRun,
    [string]$Run,
    [string]$BudgetPath,
    [int]$MinimumSamplesForP95 = 200,
    [int]$MinimumSamplesForP99 = 1000,
    [switch]$FailOnQuality,
    [switch]$ShowBuildFlavor,
    [switch]$Trend,
    [Alias('h', '?')]
    [switch]$Help
)

$ErrorActionPreference = 'Stop'
$script:HasPSStyle = $null -ne (Get-Variable -Name PSStyle -ErrorAction SilentlyContinue)
$script:QualityFailure = $false
$script:MinimumSamplesForP95Explicit = $PSBoundParameters.ContainsKey('MinimumSamplesForP95')
$script:FolderViewPresetMetrics = @(
    'folder.frame.total_us',
    'folder.frame.present_us',
    'folder.frame.input_to_paint_us',
    'folder.frame.dirty_rect_area_px',
    'folder.scroll_input_to_paint_us',
    'folder.scroll.product_paint_render_count',
    'folder.scroll.product_paint_full_client_count',
    'folder.scroll.product_paint_dirty_rect_area_px',
    'folder.relayout_to_paint_us',
    'folder.sort_toggle_us',
    'folder.filter.keystroke_to_paint_us',
    'folder.scale.enumeration_us',
    'folder.scale.first_visible_paint_us',
    'folder.scale.working_set_bytes',
    'folder.scale.private_bytes',
    'folder.scale.working_set_bytes_per_item',
    'folder.scale.private_bytes_per_item',
    'folder.cold_first_visit.enumeration_us',
    'folder.cold_first_visit.first_paint_us',
    'folder.cold_first_visit.icon_index_lookup_count',
    'folder.cold_first_visit.icon_bitmap_queue_count',
    'folder.cold_first_visit.icon_settle_us',
    'folder.cold_first_visit.thumbnail_fallback_us',
    'folder.slow_provider.enumeration_us',
    'folder.slow_provider.first_visible_paint_us',
    'FolderView.ExecuteEnumeration.IconIndex.QueryExtensions',
    'FolderView.ExecuteEnumeration.IconIndex.QueryPerFileIcons',
    'FolderView.ExecuteEnumeration.IconIndex.BuildPerFilePaths',
    'FolderView.IconLoading.ProcessQueue',
    'FolderView.IconLoading.BatchUpdate',
    'FolderView.IconLoading.BitmapConversion',
    'icons.queue_wait_to_dequeue_us',
    'icons.extract_us',
    'icons.batch_update_scan_us',
    'icons.failed_lookup_first_us',
    'icons.failed_lookup_repeat_us',
    'icons.repeated_failed_lookup_count',
    'icons.recall_avoided_count',
    'iconcache.shgetfileinfo_us',
    'iconcache.lock_wait_slow_us',
    'iconcache.lock_hold_slow_us',
    'thumbnails.extract_us',
    'thumbnails.close_to_idle_us',
    'thumbnails.provider_allowed_lookup_us',
    'thumbnails.shell_provider_allowed_count',
    'thumbnails.shell_provider_timeout_count',
    'folder.refresh.request_to_paint_us'
)
$script:MetricPresetSet = $null
if ($FolderViewPreset) {
    $script:MetricPresetSet = @{}
    foreach ($metricName in $script:FolderViewPresetMetrics) {
        $script:MetricPresetSet[$metricName] = $true
    }
}

function Get-Style([string]$Kind) {
    if (-not $script:HasPSStyle) {
        return ''
    }

    switch ($Kind) {
        'Header'      { return $PSStyle.Foreground.BrightWhite }
        'Info'        { return $PSStyle.Foreground.BrightCyan }
        'Improvement' { return $PSStyle.Foreground.BrightGreen }
        'Regression'  { return $PSStyle.Foreground.BrightRed }
        'Noise'       { return $PSStyle.Foreground.BrightYellow }
        'Muted'       { return $PSStyle.Foreground.BrightBlack }
        default       { return '' }
    }
}

function Write-StyledLine([string]$Text, [string]$Kind = 'Default') {
    $style = Get-Style $Kind
    if ($style) {
        Write-Host "$style$Text$($PSStyle.Reset)"
        return
    }

    $color = switch ($Kind) {
        'Header'      { 'White' }
        'Info'        { 'Cyan' }
        'Improvement' { 'Green' }
        'Regression'  { 'Red' }
        'Noise'       { 'Yellow' }
        'Muted'       { 'DarkGray' }
        default       { 'Gray' }
    }
    Write-Host $Text -ForegroundColor $color
}

function Get-StatusGlyph([string]$Kind) {
    switch ($Kind) {
        'Regression'  { return '▲' }
        'Improvement' { return '▼' }
        'Noise'       { return '•' }
        'Muted'       { return '○' }
        'Info'        { return '◆' }
        default       { return '·' }
    }
}

function Get-MetricFamily([string]$Metric) {
    if ([string]::IsNullOrWhiteSpace($Metric)) {
        return 'other'
    }

    if ($Metric.StartsWith('App.Startup.', [System.StringComparison]::OrdinalIgnoreCase)) {
        return 'App.Startup'
    }

    $firstDot = $Metric.IndexOf('.')
    if ($firstDot -lt 0) {
        return $Metric
    }

    $firstSegment = $Metric.Substring(0, $firstDot)
    if ($firstSegment -eq 'FileOps') {
        $secondDot = $Metric.IndexOf('.', $firstDot + 1)
        if ($secondDot -gt $firstDot) {
            return $Metric.Substring(0, $secondDot)
        }
    }

    return $firstSegment
}

function Get-FamilyRank([string]$Family) {
    switch ($Family) {
        'App.Startup'        { return 0 }
        'render'             { return 1 }
        'icons'              { return 2 }
        'iconcache'          { return 3 }
        'folder'             { return 4 }
        'FolderView'         { return 5 }
        'thumbnails'         { return 6 }
        'FileOps.Scheduler'  { return 7 }
        'FileOps.Queue'      { return 8 }
        'FileOps.Progress'   { return 9 }
        'FileOps.PreCalc'    { return 10 }
        'FileOps.ItemCompleted' { return 11 }
        'FileOps'            { return 12 }
        'search'             { return 13 }
        'find'               { return 14 }
        default              { return 99 }
    }
}

function Get-FamilyGroups([object[]]$Entries) {
    return @(
        $Entries |
            Group-Object { Get-MetricFamily $_.Metric } |
            Sort-Object @{Expression = { Get-FamilyRank $_.Name }}, Name
    )
}

function Write-FamilyTitle([string]$Family, [int]$Count) {
    Write-StyledLine ("{0} {1}  ({2} metrics)" -f (Get-StatusGlyph 'Info'), $Family, $Count) 'Header'
}

function Format-CompactValue($Value, [string]$Unit) {
    if ($null -eq $Value) {
        return 'n/a'
    }

    $numericValue = [double]$Value
    $magnitude = [Math]::Abs($numericValue)
    if ($magnitude -ge 1000000000.0) {
        $text = '{0:N2}G' -f ($numericValue / 1000000000.0)
    } elseif ($magnitude -ge 1000000.0) {
        $text = '{0:N2}M' -f ($numericValue / 1000000.0)
    } elseif ($magnitude -ge 1000.0) {
        $text = '{0:N2}K' -f ($numericValue / 1000.0)
    } else {
        $text = '{0:N2}' -f $numericValue
    }

    if ([string]::IsNullOrWhiteSpace($Unit)) {
        return $text
    }

    return "$text $Unit"
}

function Get-DeltaStatus($DeltaPct) {
    if ($null -eq $DeltaPct) {
        return 'Muted'
    }
    $numericDeltaPct = [double]$DeltaPct
    if ($numericDeltaPct -ge 10.0) {
        return 'Regression'
    }
    if ($numericDeltaPct -le -10.0) {
        return 'Improvement'
    }
    return 'Noise'
}

function Format-DeltaPct($DeltaPct) {
    if ($null -eq $DeltaPct) {
        return 'n/a'
    }
    return '{0:+0.0;-0.0;0.0}%' -f ([double]$DeltaPct)
}

function Format-DeltaValue($DeltaValue, [string]$Unit) {
    if ($null -eq $DeltaValue) {
        return 'n/a'
    }
    $numericDeltaValue = [double]$DeltaValue
    $sign = if ($numericDeltaValue -gt 0) { '+' } elseif ($numericDeltaValue -lt 0) { '' } else { '' }
    return "$sign$(Format-CompactValue $numericDeltaValue $Unit)"
}

function New-Bar([double]$Fraction, [int]$Width = 18) {
    $clamped = [Math]::Max(0.0, [Math]::Min(1.0, $Fraction))
    $filled = [Math]::Round($clamped * $Width)
    if ($filled -gt $Width) {
        $filled = $Width
    }
    return ('█' * $filled) + ('░' * ($Width - $filled))
}

function New-AsciiSparkline([double[]]$Values) {
    if (-not $Values -or $Values.Count -eq 0) {
        return ''
    }

    if ($Values.Count -eq 1) {
        return '█'
    }

    $levels = @('▁', '▂', '▃', '▄', '▅', '▆', '▇', '█')
    $min = ($Values | Measure-Object -Minimum).Minimum
    $max = ($Values | Measure-Object -Maximum).Maximum
    $span = [double]($max - $min)
    if ($span -le 0.0) {
        return ('▄' * $Values.Count)
    }

    $chars = foreach ($value in $Values) {
        $ratio = ($value - $min) / $span
        $idx = [Math]::Floor($ratio * ($levels.Count - 1))
        if ($idx -lt 0) { $idx = 0 }
        if ($idx -ge $levels.Count) { $idx = $levels.Count - 1 }
        $levels[$idx]
    }

    return -join $chars
}

function Write-Section([string]$Title) {
    Write-StyledLine ''
    $lineWidth = [Math]::Max(18, $Title.Length + 4)
    $rule = '─' * $lineWidth
    Write-StyledLine "┌$rule┐" 'Muted'
    Write-StyledLine "│ $Title │" 'Header'
    Write-StyledLine "└$rule┘" 'Muted'
}

function Show-Usage() {
    Write-StyledLine 'Show-PerfRuns.ps1' 'Header'
    Write-StyledLine 'List, summarize, compare, and trend archived performance runs.' 'Info'

    Write-Section 'Usage'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 [-RunsRoot <path>] [-Area <name>] [-Scenario <name>] [-Metric <name>] [-FolderViewPreset]'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 -Run <run-folder> [-Scenario <name>] [-Metric <name>] [-FolderViewPreset]'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 -CompareRun <baseline-run>,<candidate-run> [-Scenario <name>] [-Metric <name>] [-FolderViewPreset]'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 -Trend [-Area <name>] [-Scenario <name>] [-Metric <name>] [-FolderViewPreset]'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 -Help'

    Write-Section 'Modes'
    Write-Host '  default     Lists discovered perf runs.'
    Write-Host '  -Run        Shows one-run summary with detailed table then Unicode heat bars.'
    Write-Host '  -CompareRun Compares two runs with detailed table then colored delta chart.'
    Write-Host '  -Trend      Shows metric evolution over time with detailed table then Unicode trend graphs.'

    Write-Section 'Filters'
    Write-Host '  -RunsRoot   Root folder to scan for archived perf runs. Default: Specs\TestRuns'
    Write-Host '  -Area       Restrict discovery to one area, for example Commands, CompareDirectories, or FileOps.'
    Write-Host '  -Scenario   Filter metrics by scenario field from the JSONL records.'
    Write-Host '  -Metric     Filter to a single metric name.'
    Write-Host '  -FolderViewPreset Filter to the core FolderView frame, scale/cold/slow, scroll, icon, thumbnail, and refresh metrics.'
    Write-Host '  -BudgetPath Override the FolderView performance-budget file used for sample minimums.'
    Write-Host '  -MinimumSamplesForP95 Minimum count for p95 quality. Explicit values override budget minimums. Default: 200.'
    Write-Host '  -MinimumSamplesForP99 Minimum count for p99 quality. Default: 1000.'
    Write-Host '  -FailOnQuality Exit non-zero when a requested metric misses the p95 sample minimum.'
    Write-Host '  -ShowBuildFlavor Include build labels while listing discovered runs.'

    Write-Section 'Examples'
    Write-Host '  .\Tools\Show-PerfRuns.ps1'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 -Run D:\RedSalamander\Specs\TestRuns\7d3a1247382a\FileOps\2026-03-26_174352'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 -Run D:\RedSalamander\Specs\TestRuns\7d3a1247382a\FileOps\2026-03-26_174352 -Metric render.present_us'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 -Run D:\...\Commands\2026-06-28_160000 -FolderViewPreset'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 -CompareRun D:\...\160446,D:\...\174352'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 -Trend -Area FileOps -Metric FileOps.Operation'
    Write-Host '  .\Tools\Show-PerfRuns.ps1 -Trend -Area CompareDirectories'

    Write-Section 'Interpretation'
    Write-Host '  Regression  p95 increased by 10% or more.'
    Write-Host '  Improvement p95 decreased by 10% or more.'
    Write-Host '  Noise       change stayed within +/-10%.'
    Write-Host '  Muted       missing comparison side or mixed-unit trend; review manually.'
}

function Resolve-ExistingPath([string]$PathText) {
    return (Resolve-Path -LiteralPath $PathText -ErrorAction Stop).Path
}

function Try-ResolvePath([string]$PathText) {
    try {
        return (Resolve-Path -LiteralPath $PathText -ErrorAction Stop).Path
    } catch {
        return $null
    }
}

function Get-EnvMap([string]$RunRoot) {
    $envPath = Join-Path $RunRoot 'env.txt'
    $map = @{}
    if (-not (Test-Path -LiteralPath $envPath)) {
        return $map
    }

    foreach ($line in (Get-Content -LiteralPath $envPath)) {
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

function Get-RunFolders([string]$Root) {
    $resolvedRoot = Resolve-ExistingPath $Root
    $perfFiles = Get-ChildItem -LiteralPath $resolvedRoot -Recurse -File -Filter 'perf_metrics.jsonl'
    $runs = foreach ($file in $perfFiles) {
        $perfDir = Split-Path -Parent $file.FullName
        $runRoot = Split-Path -Parent $perfDir
        $areaName = Split-Path -Leaf (Split-Path -Parent $runRoot)
        $runName = Split-Path -Leaf $runRoot
        $env = Get-EnvMap $runRoot
        [pscustomobject]@{
            RunRoot = $runRoot
            RunName = $runName
            Area = $areaName
            PerfPath = $file.FullName
            Env = $env
        }
    }

    return $runs | Sort-Object Area, RunName
}

function Get-PerfRecords([string]$JsonlPath) {
    $records = New-Object System.Collections.Generic.List[object]
    foreach ($line in (Get-Content -LiteralPath $JsonlPath)) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $records.Add(($line | ConvertFrom-Json))
    }
    return $records.ToArray()
}

function Get-PercentileValue([double[]]$SortedValues, [double]$Percentile) {
    if ($SortedValues.Count -eq 0) {
        return $null
    }
    $index = [Math]::Ceiling(($Percentile / 100.0) * $SortedValues.Count) - 1
    if ($index -lt 0) { $index = 0 }
    if ($index -ge $SortedValues.Count) { $index = $SortedValues.Count - 1 }
    return $SortedValues[$index]
}

function Get-RecordBuildFlavor([object[]]$Records) {
    $builds = @(
        $Records | ForEach-Object {
            if ($null -ne $_ -and $_.PSObject.Properties.Name -contains 'build') {
                $build = [string]$_.build
                if (-not [string]::IsNullOrWhiteSpace($build)) {
                    $build
                }
            }
        } | Sort-Object -Unique
    )

    if ($builds.Count -eq 0) {
        return 'unknown'
    }
    if ($builds.Count -eq 1) {
        return $builds[0]
    }
    return 'mixed'
}

function Get-SampleQuality([int]$Count, [int]$MinimumSamples) {
    if ($Count -ge $MinimumSamples) {
        return 'pass'
    }
    return 'fail'
}

function Get-FolderViewBudgetMinimumSamples() {
    if ($null -ne $script:FolderViewBudgetMinimumSamples) {
        return $script:FolderViewBudgetMinimumSamples
    }

    $script:FolderViewBudgetMinimumSamples = @{}
    $resolvedBudgetPath = $BudgetPath
    if ([string]::IsNullOrWhiteSpace($resolvedBudgetPath)) {
        $repoRoot = Split-Path -Parent $PSScriptRoot
        $resolvedBudgetPath = Join-Path $repoRoot 'Specs\Testing\FolderViewPerfBudgets.json5'
    }
    if (-not (Test-Path -LiteralPath $resolvedBudgetPath)) {
        return $script:FolderViewBudgetMinimumSamples
    }

    try {
        $budgetDoc = Get-Content -LiteralPath $resolvedBudgetPath -Raw | ConvertFrom-Json
        $budgetEntries = New-Object System.Collections.Generic.List[object]
        $topLevelBudgets = $budgetDoc.PSObject.Properties['budgets']
        if ($null -ne $topLevelBudgets) {
            foreach ($entry in @($topLevelBudgets.Value)) {
                [void]$budgetEntries.Add($entry)
            }
        }

        $machines = $budgetDoc.PSObject.Properties['machines']
        if ($null -ne $machines) {
            foreach ($machine in @($machines.Value)) {
                if ($null -eq $machine) {
                    continue
                }

                $machineBudgets = $machine.PSObject.Properties['budgets']
                if ($null -eq $machineBudgets) {
                    continue
                }

                foreach ($entry in @($machineBudgets.Value)) {
                    [void]$budgetEntries.Add($entry)
                }
            }
        }

        foreach ($entry in $budgetEntries) {
            if ($null -eq $entry -or [string]::IsNullOrWhiteSpace([string]$entry.metric) -or $null -eq $entry.minimumSamples -or
                [string]$entry.stat -ne 'p95') {
                continue
            }

            $metricName = [string]$entry.metric
            $minimum = [int]$entry.minimumSamples
            if (-not $script:FolderViewBudgetMinimumSamples.ContainsKey($metricName) -or $minimum -gt $script:FolderViewBudgetMinimumSamples[$metricName]) {
                $script:FolderViewBudgetMinimumSamples[$metricName] = $minimum
            }
        }
    } catch {
        $script:FolderViewBudgetMinimumSamples = @{}
        $script:QualityFailure = $true
        Write-Warning "Unable to parse existing perf budget file '$resolvedBudgetPath': $($_.Exception.Message)"
    }

    return $script:FolderViewBudgetMinimumSamples
}

function Test-OneShotMetric([string]$MetricName) {
    if ([string]::IsNullOrWhiteSpace($MetricName)) {
        return $false
    }

    if ($MetricName.StartsWith('folder.cold_first_visit.', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    if ($MetricName.StartsWith('folder.slow_provider.', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    if ($MetricName.StartsWith('folder.scale.', [System.StringComparison]::OrdinalIgnoreCase) -and
        ($MetricName.EndsWith('_bytes', [System.StringComparison]::OrdinalIgnoreCase) -or
            $MetricName.EndsWith('_bytes_per_item', [System.StringComparison]::OrdinalIgnoreCase))) {
        return $true
    }
    if ($MetricName.EndsWith('_count', [System.StringComparison]::OrdinalIgnoreCase) -or
        $MetricName.EndsWith('_area_px', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    return $false
}

function Get-MetricQualityMinimum([string]$MetricName) {
    if ($script:MinimumSamplesForP95Explicit) {
        return $MinimumSamplesForP95
    }
    $budgetMinimums = Get-FolderViewBudgetMinimumSamples
    if ($budgetMinimums.ContainsKey($MetricName)) {
        return [int]$budgetMinimums[$MetricName]
    }
    if ($FolderViewPreset -and [string]::IsNullOrWhiteSpace($Metric)) {
        return 0
    }
    if (Test-OneShotMetric $MetricName) {
        return 0
    }
    return $MinimumSamplesForP95
}

function Get-MetricSampleQuality([int]$Count, [int]$MinimumSamples) {
    if ($MinimumSamples -le 0) {
        return 'n/a'
    }
    return Get-SampleQuality $Count $MinimumSamples
}

function Register-SampleQuality([int]$Count, [string]$MetricName, [bool]$Enabled = $true) {
    if (-not $Enabled) {
        return
    }

    $minimumSamples = Get-MetricQualityMinimum $MetricName
    if ($minimumSamples -gt 0 -and $Count -lt $minimumSamples) {
        $script:QualityFailure = $true
    }
}

function Complete-WithQualityStatus() {
    if ($FailOnQuality -and $script:QualityFailure) {
        exit 2
    }
    exit 0
}

function Measure-Metric([object[]]$Records, [string]$MetricFilter, [string]$ScenarioFilter, [bool]$RegisterQuality = $true) {
    $filtered = @($Records | ForEach-Object { $_ } | Where-Object {
        ([string]::IsNullOrWhiteSpace($MetricFilter) -or $_.metric -eq $MetricFilter) -and
        ($null -eq $script:MetricPresetSet -or $script:MetricPresetSet.ContainsKey([string]$_.metric)) -and
        ([string]::IsNullOrWhiteSpace($ScenarioFilter) -or $_.scenario -eq $ScenarioFilter)
    })

    $groups = $filtered | Group-Object -Property metric
    $summary = foreach ($group in $groups) {
        $values = @($group.Group | ForEach-Object { [double]$_.value } | Sort-Object)
        $count = $values.Count
        if ($count -eq 0) {
            continue
        }

        $qualityMinimum = Get-MetricQualityMinimum $group.Name
        Register-SampleQuality $count $group.Name $RegisterQuality

        [pscustomobject]@{
            Metric = $group.Name
            Count = $count
            Build = Get-RecordBuildFlavor @($group.Group)
            P95Quality = Get-MetricSampleQuality $count $qualityMinimum
            P99Quality = if ($qualityMinimum -le 0) { 'n/a' } else { Get-SampleQuality $count $MinimumSamplesForP99 }
            Unit = ($group.Group | Select-Object -First 1).unit
            P50 = Get-PercentileValue $values 50
            P95 = Get-PercentileValue $values 95
            P99 = Get-PercentileValue $values 99
            Max = $values[-1]
            Sum = ($values | Measure-Object -Sum).Sum
            Avg = ($values | Measure-Object -Average).Average
        }
    }

    return $summary | Sort-Object Metric
}

function Show-RunList([object[]]$Runs, [switch]$IncludeBuild) {
    $rows = foreach ($run in $Runs) {
        $row = [ordered]@{
            Area = $run.Area
            RunName = $run.RunName
            Branch = $run.Env.git_branch
            Commit = $run.Env.git_commit
        }

        if ($IncludeBuild) {
            $row['Build'] = Get-RecordBuildFlavor (Get-PerfRecords $run.PerfPath)
        }

        $row['RunRoot'] = $run.RunRoot
        [pscustomobject]$row
    }

    $rows | Format-Table -AutoSize
}

function Show-RunSummary([object]$RunInfo, [string]$MetricFilter, [string]$ScenarioFilter) {
    Write-StyledLine "Run: $($RunInfo.RunRoot)" 'Header'
    if ($RunInfo.Env.Count -gt 0) {
        Write-StyledLine "Branch: $($RunInfo.Env.git_branch)  Commit: $($RunInfo.Env.git_commit)  Area: $($RunInfo.Area)" 'Info'
    }

    $records = Get-PerfRecords $RunInfo.PerfPath
    $summary = Measure-Metric $records $MetricFilter $ScenarioFilter
    if (-not $summary) {
        Write-StyledLine 'No matching perf metrics found.' 'Noise'
        $script:QualityFailure = $true
        return
    }

    $ordered = @($summary | Sort-Object P95 -Descending)
    $maxP95 = ($ordered | Measure-Object -Property P95 -Maximum).Maximum

    Write-Section 'Detailed Table'
    foreach ($familyGroup in (Get-FamilyGroups $ordered)) {
        Write-FamilyTitle $familyGroup.Name $familyGroup.Count
        @($familyGroup.Group) | Select-Object Metric, Count, Build, P95Quality, P99Quality, Unit,
            @{N='P50';E={[math]::Round($_.P50, 2)}},
            @{N='P95';E={[math]::Round($_.P95, 2)}},
            @{N='P99';E={[math]::Round($_.P99, 2)}},
            @{N='Max';E={[math]::Round($_.Max, 2)}},
            @{N='Avg';E={[math]::Round($_.Avg, 2)}} | Format-Table -AutoSize
    }

    Write-Section 'Visual Heatmap'
    foreach ($familyGroup in (Get-FamilyGroups @($ordered | Select-Object -First 24))) {
        Write-FamilyTitle $familyGroup.Name $familyGroup.Count
        foreach ($entry in @($familyGroup.Group)) {
            $bar = New-Bar -Fraction ($(if ($maxP95 -gt 0) { [double]$entry.P95 / [double]$maxP95 } else { 0.0 })) -Width 22
            $line = '{0} {1} {2,-38} p95={3,-12} p50={4,-12} count={5,5} build={6,-8} p95q={7}' -f `
                (Get-StatusGlyph 'Info'),
                $bar,
                $entry.Metric,
                (Format-CompactValue $entry.P95 $entry.Unit),
                (Format-CompactValue $entry.P50 $entry.Unit),
                $entry.Count,
                $entry.Build,
                $entry.P95Quality
            Write-StyledLine $line 'Info'
        }
    }
}

function Compare-PerfRuns([object]$OldRun, [object]$NewRun, [string]$MetricFilter, [string]$ScenarioFilter) {
    $oldSummary = @(Measure-Metric (Get-PerfRecords $OldRun.PerfPath) $MetricFilter $ScenarioFilter $false)
    $newSummary = @(Measure-Metric (Get-PerfRecords $NewRun.PerfPath) $MetricFilter $ScenarioFilter $true)

    $oldMap = @{}
    foreach ($entry in $oldSummary) { $oldMap[$entry.Metric] = $entry }
    $newMap = @{}
    foreach ($entry in $newSummary) { $newMap[$entry.Metric] = $entry }

    $allMetrics = @($oldMap.Keys + $newMap.Keys) | Sort-Object -Unique
    $rows = foreach ($name in $allMetrics) {
        $old = $oldMap[$name]
        $new = $newMap[$name]
        $oldP95 = if ($old) { [double]$old.P95 } else { $null }
        $newP95 = if ($new) { [double]$new.P95 } else { $null }
        $deltaPct = $null
        if ($null -ne $oldP95 -and $oldP95 -ne 0 -and $null -ne $newP95) {
            $deltaPct = (($newP95 - $oldP95) / $oldP95) * 100.0
        }

        [pscustomobject]@{
            Metric = $name
            Unit = if ($new) { $new.Unit } elseif ($old) { $old.Unit } else { '' }
            OldP95 = $oldP95
            NewP95 = $newP95
            DeltaP95 = if ($null -ne $oldP95 -and $null -ne $newP95) { $newP95 - $oldP95 } else { $null }
            DeltaPct = $deltaPct
            OldCount = if ($old) { $old.Count } else { 0 }
            NewCount = if ($new) { $new.Count } else { 0 }
            OldBuild = if ($old) { $old.Build } else { '' }
            NewBuild = if ($new) { $new.Build } else { '' }
            OldP95Quality = if ($old) { $old.P95Quality } else { '' }
            NewP95Quality = if ($new) { $new.P95Quality } else { '' }
            Status = Get-DeltaStatus $deltaPct
        }
    }

    $ordered = @($rows | Sort-Object @{Expression = { [Math]::Abs($(if ($null -ne $_.DeltaPct) { [double]$_.DeltaPct } else { -1.0 })) }; Descending = $true }, Metric)
    $maxAbsDelta = ($ordered | Where-Object { $null -ne $_.DeltaPct } | ForEach-Object { [Math]::Abs([double]$_.DeltaPct) } | Measure-Object -Maximum).Maximum

    Write-StyledLine "Baseline:  $($OldRun.RunRoot)" 'Header'
    Write-StyledLine "Candidate: $($NewRun.RunRoot)" 'Header'
    Write-StyledLine 'Legend: regression >= +10%, improvement <= -10%, noise otherwise. Lower is treated as better.' 'Muted'

    $regressions = @($ordered | Where-Object { $_.Status -eq 'Regression' }).Count
    $improvements = @($ordered | Where-Object { $_.Status -eq 'Improvement' }).Count
    $noise = @($ordered | Where-Object { $_.Status -eq 'Noise' }).Count
    Write-StyledLine "Summary: regressions=$regressions  improvements=$improvements  noise=$noise" 'Info'

    Write-Section 'Detailed Table'
    foreach ($familyGroup in (Get-FamilyGroups $ordered)) {
        Write-FamilyTitle $familyGroup.Name $familyGroup.Count
        @($familyGroup.Group) | Select-Object Metric, Status, Unit,
            @{N='OldP95';E={if ($null -ne $_.OldP95) {[math]::Round($_.OldP95, 2)} else {$null}}},
            @{N='NewP95';E={if ($null -ne $_.NewP95) {[math]::Round($_.NewP95, 2)} else {$null}}},
            @{N='Delta';E={if ($null -ne $_.DeltaP95) {[math]::Round($_.DeltaP95, 2)} else {$null}}},
            @{N='DeltaPct';E={Format-DeltaPct $_.DeltaPct}},
            OldCount, NewCount, OldBuild, NewBuild, OldP95Quality, NewP95Quality | Format-Table -AutoSize
    }

    Write-Section 'Visual Delta'
    foreach ($familyGroup in (Get-FamilyGroups $ordered)) {
        Write-FamilyTitle $familyGroup.Name $familyGroup.Count
        foreach ($row in @($familyGroup.Group)) {
            $absDelta = if ($null -ne $row.DeltaPct) { [Math]::Abs([double]$row.DeltaPct) } else { 0.0 }
            $bar = New-Bar -Fraction ($(if ($maxAbsDelta -gt 0) { $absDelta / $maxAbsDelta } else { 0.0 })) -Width 22
            $line = '{0} {1} {2,-38} {3,11} → {4,-11} {5,8}  {6}' -f `
                (Get-StatusGlyph $row.Status),
                $bar,
                $row.Metric,
                (Format-CompactValue $row.OldP95 $row.Unit),
                (Format-CompactValue $row.NewP95 $row.Unit),
                (Format-DeltaPct $row.DeltaPct),
                (Format-DeltaValue $row.DeltaP95 $row.Unit)
            Write-StyledLine $line $row.Status
        }
    }
}

function Show-PerfTrend([object[]]$Runs, [string]$MetricFilter, [string]$ScenarioFilter) {
    $rows = foreach ($run in $Runs) {
        $summary = @(Measure-Metric (Get-PerfRecords $run.PerfPath) $MetricFilter $ScenarioFilter)
        foreach ($entry in $summary) {
            [pscustomobject]@{
                Area = $run.Area
                Run = $run.RunName
                Metric = $entry.Metric
                Unit = $entry.Unit
                Build = $entry.Build
                P95Quality = $entry.P95Quality
                P99Quality = $entry.P99Quality
                P50 = [math]::Round($entry.P50, 2)
                P95 = [math]::Round($entry.P95, 2)
                P99 = [math]::Round($entry.P99, 2)
                Max = [math]::Round($entry.Max, 2)
                Avg = [math]::Round($entry.Avg, 2)
                Count = $entry.Count
            }
        }
    }

    $ordered = @($rows | Sort-Object Metric, Run)
    if (-not $ordered) {
        Write-StyledLine 'No trend data found.' 'Noise'
        return
    }

    if ($MetricFilter) {
        $groups = @($ordered | Group-Object Metric)
        foreach ($group in $groups) {
            $items = @($group.Group | Sort-Object Run)
            $values = @($items | ForEach-Object { [double]$_.P95 })
            $spark = New-AsciiSparkline $values
            $first = [double]$items[0].P95
            $last = [double]$items[-1].P95
            $deltaPct = if ($first -ne 0.0) { (($last - $first) / $first) * 100.0 } else { $null }
            $units = @($items | Select-Object -ExpandProperty Unit -Unique)
            $status = if ($units.Count -gt 1) { 'Muted' } else { Get-DeltaStatus $deltaPct }
            $unitNote = if ($units.Count -gt 1) { ' [mixed units]' } else { '' }

            Write-Section "Detailed Table: $($group.Name)"
            $items | Format-Table Area, Run, Metric, Unit, Build, P95Quality, P99Quality, P50, P95, P99, Max, Avg, Count -AutoSize

            Write-Section "Visual Trend: $($group.Name)"
            $metricLine = ("{0} trend={1}  first={2}  last={3}  delta={4}" -f `
                    (Get-StatusGlyph $status),
                    $spark,
                    (Format-CompactValue $first $items[0].Unit),
                    (Format-CompactValue $last $items[0].Unit),
                    (Format-DeltaPct $deltaPct)) + $unitNote
            Write-StyledLine $metricLine $status

            $maxP95 = ($items | Measure-Object -Property P95 -Maximum).Maximum
            foreach ($item in $items) {
                $bar = New-Bar -Fraction ($(if ($maxP95 -gt 0) { [double]$item.P95 / [double]$maxP95 } else { 0.0 })) -Width 22
                $line = '  {0} {1} p95={2,-12} p50={3,-12} avg={4,-12} count={5,5}' -f `
                    $item.Run,
                    $bar,
                    (Format-CompactValue $item.P95 $item.Unit),
                    (Format-CompactValue $item.P50 $item.Unit),
                    (Format-CompactValue $item.Avg $item.Unit),
                    $item.Count
                Write-StyledLine $line 'Info'
            }
        }
        return
    }

    $overview = foreach ($group in ($ordered | Group-Object Metric)) {
        $items = @($group.Group | Sort-Object Run)
        $values = @($items | ForEach-Object { [double]$_.P95 })
        $first = [double]$items[0].P95
        $last = [double]$items[-1].P95
        $deltaPct = if ($first -ne 0.0) { (($last - $first) / $first) * 100.0 } else { $null }
        $units = @($items | Select-Object -ExpandProperty Unit -Unique)
        [pscustomobject]@{
            Metric = $group.Name
            Unit = if ($units.Count -gt 1) { 'mixed' } else { $items[0].Unit }
            Runs = $items.Count
            Spark = New-AsciiSparkline $values
            First = $first
            Last = $last
            DeltaPct = $deltaPct
            Status = if ($units.Count -gt 1) { 'Muted' } else { Get-DeltaStatus $deltaPct }
        }
    }

    Write-Section 'Detailed Table'
    foreach ($familyGroup in (Get-FamilyGroups $ordered)) {
        Write-FamilyTitle $familyGroup.Name $familyGroup.Count
        @($familyGroup.Group) | Format-Table Area, Run, Metric, Unit, Build, P95Quality, P99Quality, P50, P95, P99, Max, Avg, Count -AutoSize
    }

    Write-Section 'Trend Overview'
    Write-StyledLine 'Legend: regression >= +10%, improvement <= -10%, noise otherwise. Mixed-unit metrics are muted.' 'Muted'
    $sortedOverview = @($overview | Sort-Object @{Expression = { [Math]::Abs($(if ($null -ne $_.DeltaPct) { [double]$_.DeltaPct } else { -1.0 })) }; Descending = $true }, Metric)
    foreach ($familyGroup in (Get-FamilyGroups $sortedOverview)) {
        Write-FamilyTitle $familyGroup.Name $familyGroup.Count
        foreach ($entry in @($familyGroup.Group)) {
            $line = '{0} {1,-38} {2,-12} runs={3,2} first={4,-12} last={5,-12} delta={6,8}' -f `
                (Get-StatusGlyph $entry.Status),
                $entry.Metric,
                $entry.Spark,
                $entry.Runs,
                (Format-CompactValue $entry.First $entry.Unit),
                (Format-CompactValue $entry.Last $entry.Unit),
                (Format-DeltaPct $entry.DeltaPct)
            Write-StyledLine $line $entry.Status
        }
    }
}

if ($Help) {
    Show-Usage
    exit 0
}

$runs = @(Get-RunFolders $RunsRoot)
if (-not $runs) {
    throw "No perf runs found under $RunsRoot"
}

if ($Area) {
    $runs = @($runs | Where-Object { $_.Area -eq $Area })
}

if (-not $Run -and -not $CompareRun -and -not $Trend -and -not $Metric -and -not $Scenario -and -not $FolderViewPreset) {
    Show-RunList $runs -IncludeBuild:$ShowBuildFlavor
    exit 0
}

if ($CompareRun) {
    if ($CompareRun.Count -ne 2) {
        throw '-CompareRun requires exactly two run folder paths.'
    }
    $oldRunPath = Resolve-ExistingPath $CompareRun[0]
    $newRunPath = Resolve-ExistingPath $CompareRun[1]
    $oldRun = $runs | Where-Object { $_.RunRoot -eq $oldRunPath } | Select-Object -First 1
    $newRun = $runs | Where-Object { $_.RunRoot -eq $newRunPath } | Select-Object -First 1
    if (-not $oldRun -or -not $newRun) {
        throw 'One or both compare runs were not found in the discovered perf run set.'
    }
    Compare-PerfRuns $oldRun $newRun $Metric $Scenario
    Complete-WithQualityStatus
}

if ($Trend) {
    Show-PerfTrend $runs $Metric $Scenario
    Complete-WithQualityStatus
}

if (-not $Run) {
    $Run = $runs[-1].RunRoot
}

$resolvedRun = Resolve-ExistingPath $Run
$runInfo = $runs | Where-Object { $_.RunRoot -eq $resolvedRun } | Select-Object -First 1
if (-not $runInfo) {
    throw "Run not found in discovered perf runs: $Run"
}

Show-RunSummary $runInfo $Metric $Scenario
Complete-WithQualityStatus
