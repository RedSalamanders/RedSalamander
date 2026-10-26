$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$packet='Z:/src/RedSalamander/Specs/TestRuns/Local-x64/Navigation/2026-10-20_DebugSortingRowMoves'
$binding=Get-Content (Join-Path $packet 'artifact-bindings.json') -Raw|ConvertFrom-Json
$checked=0
foreach($copy in $binding.original_copies){
 if((Get-FileHash -LiteralPath $copy.source).Hash-ne$copy.sha256-or(Get-FileHash -LiteralPath (Join-Path $packet $copy.archive)).Hash-ne$copy.sha256){throw "Original binding failed: $($copy.archive)"}
 $checked++
}
foreach($extract in $binding.derived_extracts){
 if((Get-FileHash -LiteralPath $extract.source).Hash-ne$extract.source_sha256-or(Get-FileHash -LiteralPath (Join-Path $packet $extract.archive)).Hash-ne$extract.archive_sha256){throw 'Extract binding failed'}
 if((Get-Item -LiteralPath $extract.source).Length-ne$extract.source_bytes){throw 'Raw size differs'}
 $checked++
}
$summary=Get-Content (Join-Path $packet 'performance-comparison.json') -Raw|ConvertFrom-Json
$allowed=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach($metric in @('folder.sort_toggle_us','FolderView.ApplyCurrentSort','FolderView.ApplyCurrentSort.OrderItems','FolderView.ApplyCurrentSort.RowMoves','FolderView.ApplyCurrentSort.ScratchBytes','FolderView.ExecuteEnumeration.BuildItems','FolderView.ExecuteEnumeration.SortMerge','FolderView.ExecuteEnumeration.NaturalSortRanks')){$null=$allowed.Add($metric)}
$verdicts=@()
foreach($property in $summary.phases.PSObject.Properties){
 $phase=$property.Name;$computed=$property.Value
 $raw=[IO.File]::ReadAllLines((Join-Path $root "$phase/perf/perf_metrics.jsonl"))
 $expected=[Collections.Generic.List[string]]::new();$sorts=[Collections.Generic.List[object]]::new()
 foreach($line in $raw){if([string]::IsNullOrWhiteSpace($line)){continue};$row=$line|ConvertFrom-Json;if($allowed.Contains($row.metric)){$expected.Add($line)};if($row.metric-eq'folder.sort_toggle_us'){$sorts.Add($row)}}
 $kept=[IO.File]::ReadAllLines((Join-Path $packet "$phase/perf/perf-selected.jsonl"))
 if($kept.Count-ne$expected.Count-or$sorts.Count-ne50){throw 'Relevant metric completeness failed'}
 for($i=0;$i-lt$kept.Count;$i++){if(![StringComparer]::Ordinal.Equals($kept[$i],$expected[$i])){throw 'Metric content/order differs'}}
 for($i=0;$i-lt50;$i++){if($sorts[$i].durationUs-ne$computed.samples[$i].duration_us){throw 'Raw sample differs from summary'}}
 for($i=0;$i-lt10;$i++){
  $values=@($sorts[$i].durationUs,$sorts[$i+10].durationUs,$sorts[$i+20].durationUs,$sorts[$i+30].durationUs,$sorts[$i+40].durationUs)|Sort-Object
  $stats=$computed.requests[$i].stats
  if($stats.count-ne5-or$stats.min_us-ne$values[0]-or$stats.median_us-ne$values[2]-or$stats.max_us-ne$values[4]){throw 'Request statistics differ'}
 }
 $metricRows=@($kept|ForEach-Object {$_|ConvertFrom-Json})
 foreach($stage in $computed.stages.PSObject.Properties){
  $metric=$stage.Name
  $values=@($metricRows|Where-Object {$_.metric-eq$metric-and$(if($metric-eq'FolderView.ExecuteEnumeration.SortMerge'){$_.value0-eq0-and$_.value1-eq5000}else{$_.value0-eq5000})}|ForEach-Object durationUs|Sort-Object)
  if(!$values.Count){if($null-ne$stage.Value){throw 'Empty stage has statistics'};continue}
  $middle=[int][Math]::Floor($values.Count/2)
  $median=if($values.Count%2){$values[$middle]}else{([decimal]$values[$middle-1]+[decimal]$values[$middle])/2}
  if($stage.Value.count-ne$values.Count-or$stage.Value.min_us-ne$values[0]-or$stage.Value.max_us-ne$values[-1]-or$stage.Value.median_us-ne$median){throw 'Core stage statistics differ from original rows'}
 }
 $observation=Get-Content (Join-Path $packet "$phase/observation.json") -Raw|ConvertFrom-Json
 $rawMonitor=@(Get-Content (Join-Path $root "$phase/build-process-samples.jsonl")|ConvertFrom-Json)
 $rawBenchmark=@($rawMonitor|Where-Object stage -eq 'benchmark')
 $keptBenchmark=@(Get-Content (Join-Path $packet "$phase/build-process-qualifying-samples.jsonl")|ConvertFrom-Json|Where-Object stage -eq 'benchmark')
 $contaminated=@($rawBenchmark|Where-Object {$_.tools.Count}).Count
 if($rawBenchmark.Count-ne$observation.measurement_sample_count-or$keptBenchmark.Count-ne$rawBenchmark.Count-or$contaminated-ne$observation.measurement_contaminated_polls){throw 'Observer completeness failed'}
 if(($rawBenchmark.gap_ms|Measure-Object -Maximum).Maximum-ne$observation.maximum_measurement_gap_ms){throw 'Observer maximum gap differs'}
 for($i=0;$i-lt$rawBenchmark.Count;$i++){if(($rawBenchmark[$i]|ConvertTo-Json -Depth 6 -Compress)-ne($keptBenchmark[$i]|ConvertTo-Json -Depth 6 -Compress)){throw 'Observer row differs'}}
 if($computed.accepted-and($contaminated-or!$observation.build_free_eligible-or$observation.maximum_measurement_gap_ms-gt1000)){throw 'Accepted attempt fails observation rules'}
 $cases=Get-Content (Join-Path $packet "$phase/commands/results.json") -Raw|ConvertFrom-Json
 if($cases.passed-ne5-or$cases.failed-or$cases.skipped){throw 'Native benchmark verdict differs'}
 $verdicts += [ordered]@{phase=$phase;accepted=$computed.accepted;selected_rows=$kept.Count;sort_samples=50;benchmark_snapshots=$rawBenchmark.Count;contaminated_snapshots=$contaminated;all_relevant_rows_preserved=$true;all_samples_and_request_statistics_verified=$true;all_measurement_snapshots_preserved=$true}
}
foreach($configuration in @('Debug','Release')){
 $pair=$summary.pairs.$configuration
 $b=Get-Content (Join-Path $packet "$($pair.baseline)/build-receipt.json") -Raw|ConvertFrom-Json
 $c=Get-Content (Join-Path $packet "$($pair.candidate)/build-receipt.json") -Raw|ConvertFrom-Json
 foreach($field in @('configuration','platform','repo_identity','toolchain_identity_digest','vcpkg_identity_digest','runtime_dependency_digest','project_graph_digest')){if($b.$field-ne$c.$field){throw "Pair identity differs: $configuration $field"}}
 if(($b.build_arguments|ConvertTo-Json -Compress)-ne($c.build_arguments|ConvertTo-Json -Compress)-or!$b.tests_enabled-or!$c.tests_enabled){throw 'Build flags differ'}
 $bo=Get-Content (Join-Path $packet "$($pair.baseline)/observation.json") -Raw|ConvertFrom-Json
 $co=Get-Content (Join-Path $packet "$($pair.candidate)/observation.json") -Raw|ConvertFrom-Json
 if($bo.observer_sha256-ne$co.observer_sha256-or$bo.scanner_sha256-ne$co.scanner_sha256){throw 'Observer code differs within pair'}
 foreach($request in $pair.requests){
  $change=[Math]::Round(100*($request.candidate_median_us-$request.baseline_median_us)/$request.baseline_median_us,3)
  if($change-ne$request.change_percent){throw 'Pair request percentage differs'}
 }
}
[ordered]@{verified_utc=[DateTime]::UtcNow.ToString('o');status='passed';verified_bindings=$checked;phases=$verdicts;pair_profile_checkout_toolchain_dependency_flags_and_observer_identities_match=$true;scope='Independent raw-row, copy, summary, observation and receipt comparison. Native functional verdicts and current executable hashes are separate evidence.'}|ConvertTo-Json -Depth 9|Set-Content (Join-Path $root 'independent-evidence-verification.json')
Copy-Item -LiteralPath (Join-Path $root 'independent-evidence-verification.json') -Destination (Join-Path $packet 'independent-evidence-verification.json')
[pscustomobject]@{status='passed';verified_bindings=$checked;verified_benchmark_attempts=$verdicts.Count}
