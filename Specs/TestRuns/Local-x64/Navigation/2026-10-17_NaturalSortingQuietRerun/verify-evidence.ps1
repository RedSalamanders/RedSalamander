$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$packet='Z:/src/RedSalamander/Specs/TestRuns/Local-x64/Navigation/2026-10-17_NaturalSortingQuietRerun'
$binding=Get-Content (Join-Path $packet 'artifact-bindings.json') -Raw|ConvertFrom-Json
$checked=0
foreach($copy in $binding.original_copies){
 if((Get-FileHash -LiteralPath (Join-Path $packet $copy.archive)).Hash-ne$copy.sha256-or(Get-FileHash -LiteralPath $copy.source).Hash-ne$copy.sha256){throw "Copy binding mismatch: $($copy.archive)"}
 $checked++
}
foreach($extract in @($binding.derived_extracts)){
 if((Get-FileHash -LiteralPath (Join-Path $packet $extract.archive)).Hash-ne$extract.archive_sha256-or(Get-FileHash -LiteralPath $extract.source).Hash-ne$extract.source_sha256){throw 'Derived binding mismatch'}
 $checked++
}
$summary=Get-Content (Join-Path $packet 'performance-comparison.json') -Raw|ConvertFrom-Json
$allowed=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach($name in @('folder.sort_toggle_us','FolderView.ApplyCurrentSort','FolderView.ExecuteEnumeration.BuildItems','FolderView.ExecuteEnumeration.SortMerge','FolderView.ExecuteEnumeration.NaturalSortRanks')){$null=$allowed.Add($name)}
$verdicts=[Collections.Generic.List[object]]::new()
foreach($phase in @('quiet-debug-02','quiet-debug-03','quiet-release-01','quiet-release-02')){
 $rawLines=[IO.File]::ReadAllLines((Join-Path $root "$phase/perf/perf_metrics.jsonl"))
 $selected=[Collections.Generic.List[string]]::new()
 $sorts=[Collections.Generic.List[object]]::new()
 $rows=0
 foreach($line in $rawLines){if([string]::IsNullOrWhiteSpace($line)){continue};$row=$line|ConvertFrom-Json;$rows++;if($allowed.Contains($row.metric)){$selected.Add($line)};if($row.metric-eq'folder.sort_toggle_us'){$sorts.Add($row)}}
 $archiveLines=[IO.File]::ReadAllLines((Join-Path $packet "$phase/perf/perf-selected.jsonl"))
 if($sorts.Count-ne50-or$archiveLines.Count-ne$selected.Count){throw 'Metric family count mismatch'}
 for($i=0;$i-lt$archiveLines.Count;$i++){if(![StringComparer]::Ordinal.Equals($archiveLines[$i],$selected[$i])){throw 'Metric row or order changed'}}
 $computed=$summary.phases.$phase
 for($i=0;$i-lt50;$i++){if($sorts[$i].durationUs-ne$computed.samples[$i].duration_us){throw 'Sample summary differs from raw'}}
 for($requestIndex=0;$requestIndex-lt10;$requestIndex++){
  $values=@($sorts[$requestIndex].durationUs,$sorts[$requestIndex+10].durationUs,$sorts[$requestIndex+20].durationUs,$sorts[$requestIndex+30].durationUs,$sorts[$requestIndex+40].durationUs)|Sort-Object
  $stats=$computed.requests[$requestIndex].stats
  if($stats.count-ne5-or$stats.min_us-ne$values[0]-or$stats.median_us-ne$values[2]-or$stats.max_us-ne$values[4]){throw 'Per-request summary mismatch'}
 }
 $rawMonitor=@(Get-Content (Join-Path $root "$phase/build-process-samples.jsonl")|ConvertFrom-Json)
 $rawBenchmark=@($rawMonitor|Where-Object stage -eq 'benchmark')
 $keptBenchmark=@(Get-Content (Join-Path $packet "$phase/build-process-qualifying-samples.jsonl")|ConvertFrom-Json|Where-Object stage -eq 'benchmark')
 $observation=Get-Content (Join-Path $packet "$phase/observation.json") -Raw|ConvertFrom-Json
 $contaminated=@($rawBenchmark|Where-Object {$_.tools.Count}).Count
 if($rawBenchmark.Count-ne$observation.measurement_sample_count-or$keptBenchmark.Count-ne$rawBenchmark.Count-or$contaminated-ne$observation.measurement_contaminated_polls){throw 'Monitor row count mismatch'}
 if(($rawBenchmark.gap_ms|Measure-Object -Maximum).Maximum-ne$observation.maximum_measurement_gap_ms){throw 'Maximum poll gap mismatch'}
 for($i=0;$i-lt$rawBenchmark.Count;$i++){if(($rawBenchmark[$i]|ConvertTo-Json -Depth 6 -Compress)-ne($keptBenchmark[$i]|ConvertTo-Json -Depth 6 -Compress)){throw 'Monitor measurement row changed'}}
 $cases=Get-Content (Join-Path $packet "$phase/commands/results.json") -Raw|ConvertFrom-Json
 if($cases.passed-ne5-or$cases.failed-or$cases.skipped-or$cases.cases.Count-ne5){throw 'Case verdicts differ'}
 $accepted=$phase-in@('quiet-debug-03','quiet-release-02')
 if($accepted-and($contaminated-or!$observation.build_free_eligible)){throw 'Accepted run is not build-free eligible'}
 $verdicts.Add([ordered]@{phase=$phase;accepted=$accepted;raw_metric_rows=$rows;selected_metric_rows=$selected.Count;all_sort_samples=$sorts.Count;benchmark_snapshots=$rawBenchmark.Count;build_tool_or_launcher_polls=$contaminated;cases_passed=$cases.passed;metric_row_order_preserved=$true;all_measurement_snapshots_preserved=$true;summary_verified_from_raw=$true})
}
[ordered]@{verified_utc=[DateTime]::UtcNow.ToString('o');binding_count=$checked;status='passed';phases=$verdicts}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $root 'independent-evidence-verification.json') -Encoding utf8
[pscustomobject]@{status='passed';verified_bindings=$checked;verified_sort_samples=200;accepted_sort_samples=100;verified_benchmark_attempts=4}
