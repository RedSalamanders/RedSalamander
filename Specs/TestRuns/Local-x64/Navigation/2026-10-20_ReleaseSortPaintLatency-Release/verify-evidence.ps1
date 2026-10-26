$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$reports=@()
$allowed=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach($name in @('folder.sort_toggle_us','folder.sort_product_paint_us','folder.sort_product_warm_calls','folder.sort_legacy_warm_calls','FolderView.ApplyCurrentSort','FolderView.ApplyCurrentSort.OrderItems','FolderView.ApplyCurrentSort.RowMoves','FolderView.ApplyCurrentSort.ScratchBytes','render.present_us','render.begin_to_enddraw_us','render.layout_items_us','dwrite.text_layout.frame_create_us','dwrite.text_layout.frame_create_count')){$null=$allowed.Add($name)}
foreach($name in @('FolderView.ExecuteEnumeration.BuildItems','FolderView.ExecuteEnumeration.SortMerge','FolderView.ExecuteEnumeration.NaturalSortRanks')){$null=$allowed.Add($name)}
foreach($configuration in @('Debug','Release')){
 $packet="Z:/src/RedSalamander/Specs/TestRuns/Local-x64/Navigation/2026-10-20_ReleaseSortPaintLatency-$configuration"
 $binding=Get-Content (Join-Path $packet 'artifact-bindings.json') -Raw|ConvertFrom-Json
 $checked=0
 foreach($entry in $binding.original_copies){
  if((Get-FileHash -LiteralPath $entry.source).Hash-ne$entry.sha256-or(Get-FileHash -LiteralPath (Join-Path $packet $entry.archive)).Hash-ne$entry.sha256){throw 'Original copy binding differs'}
  $checked++
 }
 foreach($entry in $binding.derived_extracts){
  if((Get-FileHash -LiteralPath $entry.source).Hash-ne$entry.source_sha256-or(Get-FileHash -LiteralPath (Join-Path $packet $entry.archive)).Hash-ne$entry.archive_sha256-or(Get-Item -LiteralPath $entry.source).Length-ne$entry.source_bytes){throw 'Extract source/copy binding differs'}
  $checked++
 }
 $summary=Get-Content (Join-Path $packet 'performance-comparison.json') -Raw|ConvertFrom-Json
 $phaseChecks=@();$candidateReceipts=@()
 foreach($property in $summary.phases.PSObject.Properties){
  $phase=$property.Name;$s=$property.Value
  $rawPath=Join-Path $root "$phase/perf/perf_metrics.jsonl"
  $raw=[IO.File]::ReadAllLines($rawPath)
  $expected=[Collections.Generic.List[string]]::new();$rows=[Collections.Generic.List[object]]::new()
  foreach($line in $raw){if([string]::IsNullOrWhiteSpace($line)){continue};$row=$line|ConvertFrom-Json;if($allowed.Contains($row.metric)){$expected.Add($line);$rows.Add($row)}}
  $kept=[IO.File]::ReadAllLines((Join-Path $packet "$phase/perf/perf-selected.jsonl"))
  if($kept.Count-ne$expected.Count){throw 'Relevant metric row coverage differs'}
  for($i=0;$i-lt$kept.Count;$i++){if(![StringComparer]::Ordinal.Equals($kept[$i],$expected[$i])){throw 'Relevant metric content/order differs'}}
  $metric=if($s.protocol-eq'product_paint_v2'){'folder.sort_product_paint_us'}else{'folder.sort_toggle_us'}
  $sorts=@($rows|Where-Object metric -eq $metric)
  if($sorts.Count-ne50-or$s.samples.Count-ne50){throw 'Sort sample count differs'}
  for($i=0;$i-lt50;$i++){if($sorts[$i].durationUs-ne$s.samples[$i].duration_us-or$sorts[$i].value1-ne$s.samples[$i].render_calls-or$sorts[$i].value0-ne5000){throw 'Original request sample differs'}}
  $core=@($rows|Where-Object {$_.metric-eq'FolderView.ApplyCurrentSort.OrderItems'-and$_.value0-eq5000}|ForEach-Object durationUs|Sort-Object)
  if($core.Count){$mid=[int][Math]::Floor($core.Count/2);$median=if($core.Count%2){$core[$mid]}else{([decimal]$core[$mid-1]+[decimal]$core[$mid])/2};if($s.core_ordering.count-ne$core.Count-or$s.core_ordering.min_us-ne$core[0]-or$s.core_ordering.max_us-ne$core[-1]-or$s.core_ordering.median_us-ne$median){throw 'Core ordering statistics differ'}}
  $rawMonitor=@(Get-Content (Join-Path $root "$phase/build-process-samples.jsonl")|ConvertFrom-Json|Where-Object stage -eq 'benchmark')
  $keptMonitor=@(Get-Content (Join-Path $packet "$phase/build-process-qualifying-samples.jsonl")|ConvertFrom-Json|Where-Object stage -eq 'benchmark')
  $o=Get-Content (Join-Path $packet "$phase/observation.json") -Raw|ConvertFrom-Json
  $overlap=@($rawMonitor|Where-Object {$_.tools.Count}).Count
  if($rawMonitor.Count-ne$o.measurement_sample_count-or$rawMonitor.Count-ne$keptMonitor.Count-or$overlap-ne$o.measurement_contaminated_polls){throw 'Measurement observation completeness differs'}
  for($i=0;$i-lt$rawMonitor.Count;$i++){if(($rawMonitor[$i]|ConvertTo-Json -Depth 6 -Compress)-ne($keptMonitor[$i]|ConvertTo-Json -Depth 6 -Compress)){throw 'Observation content/order differs'}}
  if(($rawMonitor.gap_ms|Measure-Object -Maximum).Maximum-ne$o.maximum_measurement_gap_ms){throw 'Observation maximum gap differs'}
  if($s.selected-and(!$o.build_free_eligible-or$overlap-or$o.maximum_measurement_gap_ms-gt1000)){throw 'Selected phase fails observation rules'}
  $ref=Get-Content (Join-Path $packet "$phase-receipt-ref.json") -Raw|ConvertFrom-Json
  $receipt=Get-Content (Join-Path $packet $ref.archive) -Raw|ConvertFrom-Json
  if((Get-FileHash -LiteralPath (Join-Path $packet $ref.archive)).Hash-ne$ref.source_sha256-or$receipt.receipt_id-ne$s.receipt_id-or!$receipt.tests_enabled-or$receipt.exit_code-ne0){throw 'Receipt reference differs'}
  $cases=Get-Content (Join-Path $packet "$phase/commands/results.json") -Raw|ConvertFrom-Json
  if($cases.passed-ne5-or$cases.failed-or$cases.skipped){throw 'Native benchmark verdict differs'}
  if($phase-like'candidate-*'-and$s.selected){
   $candidateReceipts+= $receipt.receipt_id
   $warmName=if($s.protocol-eq'product_paint_v2'){'folder.sort_product_warm_calls'}else{'folder.sort_legacy_warm_calls'}
   $warm=@($rows|Where-Object metric -eq $warmName)
   if($warm.Count-ne50){throw 'Counter sample coverage differs'}
   for($i=0;$i-lt50;$i++){
    if($warm[$i].value0-ne$s.samples[$i].warm_calls-or$warm[$i].value1-ne$s.samples[$i].successful_presents-or$warm[$i].durationUs-ne0){throw 'Counter sample differs'}
    if($s.protocol-eq'product_paint_v2'){
     if($warm[$i].value0-ne0-or$(if($i%10-eq0){$sorts[$i].value1-ne0-or$warm[$i].value1-ne0}else{$warm[$i].value1-lt1})){throw 'Product paint proof fails'}
    }elseif($warm[$i].value0-ne1){throw 'Legacy warm control did not retain synthetic warm call'}
   }
  }
  $phaseChecks += [ordered]@{phase=$phase;sort_samples=50;all_relevant_rows=$kept.Count;measurement_snapshots=$rawMonitor.Count;overlap_polls=$overlap;selected=$s.selected}
 }
 if(@($candidateReceipts|Sort-Object -Unique).Count-ne1){throw 'Candidate blocks do not share one exact executable receipt per profile'}
 foreach($property in $summary.role_statistics.PSObject.Properties){
  $role=$property.Name;$group=@($summary.phases.PSObject.Properties.Value|Where-Object {$_.selected-and$_.role-eq$role})
  foreach($request in $property.Value){
   $values=@($group|ForEach-Object samples|Where-Object request_index -eq $request.request_index|ForEach-Object duration_us|Sort-Object)
   $mid=[int][Math]::Floor($values.Count/2);$median=if($values.Count%2){$values[$mid]}else{([decimal]$values[$mid-1]+[decimal]$values[$mid])/2}
   if($request.stats.count-ne$values.Count-or$request.stats.min_us-ne$values[0]-or$request.stats.max_us-ne$values[-1]-or$request.stats.median_us-ne$median){throw 'Role request statistics differ'}
  }
 }
 $native=Get-Content (Join-Path $packet "functional-$($configuration.ToLowerInvariant())-01/commands/results.json") -Raw|ConvertFrom-Json
 if($native.passed-ne9-or$native.failed-or$native.skipped){throw 'Functional verdict differs'}
 $reports += [ordered]@{configuration=$configuration;status='passed';verified_bindings=$checked;phases=$phaseChecks;exact_candidate_receipt_across_counterbalanced_blocks=$candidateReceipts[0]}
}
$verification=[ordered]@{verified_utc=[DateTime]::UtcNow.ToString('o');status='passed';profiles=$reports;scope='Independent original/extract digests, relevant-row coverage/order, every request and counter, grouped/core statistics, complete measurement snapshots, native verdicts and shared executable receipt within each counterbalanced profile. Initial v1 baselines retain their narrower observer scope.'}
$verification|ConvertTo-Json -Depth 10|Set-Content (Join-Path $root 'independent-verification.json')
foreach($configuration in @('Debug','Release')){Copy-Item -LiteralPath (Join-Path $root 'independent-verification.json') -Destination "Z:/src/RedSalamander/Specs/TestRuns/Local-x64/Navigation/2026-10-20_ReleaseSortPaintLatency-$configuration/independent-verification.json"}
$reports|ForEach-Object {[pscustomobject]@{configuration=$_.configuration;status=$_.status;verified_bindings=$_.verified_bindings;phases=$_.phases.Count}}
