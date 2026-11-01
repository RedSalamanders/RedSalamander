param()
$ErrorActionPreference = 'Stop'
$taskRoot = Join-Path $PSScriptRoot 'paired-current-r21'
function Get-TaskMedian([double[]]$Values) {
    if (-not $Values.Count) { throw 'Empty measurement set.' }
    $taskSorted = @($Values | Sort-Object)
    $taskMiddle = [int][Math]::Floor($taskSorted.Count / 2)
    if ($taskSorted.Count % 2) { return $taskSorted[$taskMiddle] }
    return ($taskSorted[$taskMiddle - 1] + $taskSorted[$taskMiddle]) / 2.0
}
$taskCaptures = [Collections.Generic.List[object]]::new()
foreach ($taskVariant in @('baseline', 'candidate')) {
    foreach ($taskIndex in 1..5) {
        $taskPath = Join-Path $taskRoot "$taskVariant-$taskIndex/artifacts/selftest/last_run/perf/perf_metrics.jsonl"
        $taskRows = @([IO.File]::ReadAllLines($taskPath) | ForEach-Object {
            if ($_ -match '"metric"\s*:\s*"FileOps\.(IntegratedPresentation\.Cpu|Popup\.)') { $_ | ConvertFrom-Json }
        })
        $taskCpu = @($taskRows | Where-Object metric -eq 'FileOps.IntegratedPresentation.Cpu')
        if ($taskCpu.Count -ne 29) { throw 'CPU phase set differs from matched fixture.' }
        $taskIntervals = [Collections.Generic.List[object]]::new()
        for ($taskI = 1; $taskI -lt $taskCpu.Count; ++$taskI) {
            $taskIntervals.Add([pscustomobject]@{
                phase = $taskCpu[$taskI].detail
                ui_us = [int64]$taskCpu[$taskI].value0 - [int64]$taskCpu[$taskI - 1].value0
                process_us = [int64]$taskCpu[$taskI].value1 - [int64]$taskCpu[$taskI - 1].value1
            })
        }
        $taskCosts = @($taskRows | Where-Object metric -in @(
            'FileOps.Popup.HostedControls.SyncUs', 'FileOps.Popup.HostedPaintUs', 'FileOps.Popup.WmPaintUs',
            'FileOps.Popup.Render.BuildSnapshotUs', 'FileOps.Popup.Render.CardLayoutUs',
            'FileOps.Popup.Render.AutoResizeUs', 'FileOps.Popup.Render.ScrollLayoutUs',
            'FileOps.Popup.Render.DrawUs', 'FileOps.Popup.Render.TotalUs') | Group-Object metric, detail | ForEach-Object {
                [pscustomobject]@{
                    metric = $_.Group[0].metric; detail = $_.Group[0].detail; samples = $_.Count
                    median_us = Get-TaskMedian @($_.Group.durationUs | ForEach-Object { [double]$_ })
                    min_us = ($_.Group.durationUs | Measure-Object -Minimum).Minimum
                    max_us = ($_.Group.durationUs | Measure-Object -Maximum).Maximum
                }
            })
        $taskCaptures.Add([pscustomobject]@{
            variant = $taskVariant; capture = $taskIndex; intervals = $taskIntervals.ToArray(); popup_costs = $taskCosts
            raw_sha256 = (Get-FileHash -LiteralPath $taskPath -Algorithm SHA256).Hash.ToLowerInvariant()
        })
    }
}
$taskComparisons = [Collections.Generic.List[object]]::new()
foreach ($taskPhase in @($taskCaptures[0].intervals.phase)) {
    $taskBefore = @($taskCaptures | Where-Object variant -eq 'baseline' | ForEach-Object {
        ($_.intervals | Where-Object phase -eq $taskPhase).ui_us
    })
    $taskAfter = @($taskCaptures | Where-Object variant -eq 'candidate' | ForEach-Object {
        ($_.intervals | Where-Object phase -eq $taskPhase).ui_us
    })
    $taskB = Get-TaskMedian $taskBefore; $taskA = Get-TaskMedian $taskAfter
    $taskComparisons.Add([pscustomobject]@{
        phase = $taskPhase; baseline_ui_median_us = $taskB; candidate_ui_median_us = $taskA
        delta_percent = if ($taskB -gt 0) { 100.0 * ($taskA - $taskB) / $taskB } else { $null }
        baseline_samples = $taskBefore; candidate_samples = $taskAfter
    })
}
$taskResult = [ordered]@{
    schema = 'red-salamander.i27-interleaved-cost-diagnostic.v1'
    method = 'Five matched fresh processes per variant. CPU phase endpoint differences and per-capture popup component medians; no FPS or GPU inference.'
    limitations = @('CPU values have 15.625 ms sampling granularity.',
        'Popup component summaries span the complete fixture; they are not isolated graph-only costs.',
        'Elapsed paint durations include presentation synchronization and scheduler delay; this is diagnostic attribution, not a new acceptance baseline.')
    phase_comparisons = $taskComparisons.ToArray(); captures = $taskCaptures.ToArray()
}
[IO.File]::WriteAllText((Join-Path $taskRoot 'cost-diagnostic.json'), ($taskResult | ConvertTo-Json -Depth 10), [Text.UTF8Encoding]::new($false))
$taskComparisons | Where-Object phase -in @('popup-open-before', 'popup-open-after', 'popup-collapsed-before', 'popup-collapsed-after', 'popup-after-churn') | Format-Table -AutoSize
foreach ($taskVariant in @('baseline', 'candidate')) {
    Write-Output "Whole-fixture component medians across captures: $taskVariant"
    $taskCaptures | Where-Object variant -eq $taskVariant | ForEach-Object popup_costs | Group-Object metric, detail | ForEach-Object {
        [pscustomobject]@{ metric = $_.Group[0].metric; detail = $_.Group[0].detail; median_us = Get-TaskMedian @($_.Group.median_us) }
    } | Format-Table -AutoSize
}
