$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$repo='Z:/src/RedSalamander'
$identities=Get-Content (Join-Path $root 'final-identities.json') -Raw|ConvertFrom-Json
$verified=@(foreach($profile in @('Debug','Release')){
 $packet=Join-Path $repo "Specs/TestRuns/Local-x64/Navigation/2026-10-20_DebugSortWriter-$profile"
 $binding=Get-Content (Join-Path $packet 'artifact-bindings.json') -Raw|ConvertFrom-Json
 $decodedObservers=@{}
 foreach($record in $binding.bindings){
  $archive=Join-Path $packet $record.archive
  if($record.kind-eq'complete-original'){
   if((Get-FileHash -LiteralPath $record.source).Hash-ne$record.sha256-or(Get-FileHash -LiteralPath $archive).Hash-ne$record.sha256){throw 'Original artifact hash differs'}
  }elseif($record.kind-eq'gzip-complete-original'){
   if((Get-FileHash -LiteralPath $record.source).Hash-ne$record.source_sha256-or(Get-FileHash -LiteralPath $archive).Hash-ne$record.compressed_sha256){throw 'Original/compressed observation hash differs'}
   $compressedStream=$null;$gzip=$null;$decoded=$null
   try{
    $compressedStream=[IO.File]::OpenRead($archive)
    $gzip=[IO.Compression.GZipStream]::new($compressedStream,[IO.Compression.CompressionMode]::Decompress)
    $decoded=[IO.MemoryStream]::new()
    $gzip.CopyTo($decoded)
    $bytes=$decoded.ToArray()
    if($bytes.Length-ne$record.source_bytes-or[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))-ne$record.source_sha256){throw 'Compressed observations lost or changed original bytes'}
    $decodedObservers[$record.archive]=[Text.Encoding]::UTF8.GetString($bytes)
   }finally{if($decoded){$decoded.Dispose()};if($gzip){$gzip.Dispose()};if($compressedStream){$compressedStream.Dispose()}}
  }else{
   if((Get-FileHash -LiteralPath $record.source).Hash-ne$record.source_sha256-or(Get-FileHash -LiteralPath $archive).Hash-ne$record.selected_sha256){throw 'Raw/extract metric hash differs'}
   $original=[IO.File]::ReadAllLines($record.source)
   $expected=@(foreach($line in $original){if(($line|ConvertFrom-Json).metric-match$record.metric_filter){$line}})
   $actual=[IO.File]::ReadAllLines($archive)
   if($expected.Count-ne$actual.Count-or$actual.Count-ne$record.selected_rows){throw 'Metric-family coverage differs'}
   for($index=0;$index-lt$actual.Count;$index++){if($actual[$index]-cne$expected[$index]){throw 'Metric content/order changed'}}
  }
 }
 $data=Get-Content (Join-Path $packet 'performance-comparison.json') -Raw|ConvertFrom-Json
 foreach($phase in $data.phases){
  $receipt=Get-Content (Join-Path $packet "receipts/$($phase.receipt_id).json") -Raw|ConvertFrom-Json
  $identity=@($identities.profiles|Where-Object {$_.configuration-eq$profile-and$_.role-eq$phase.role})
  if($identity.Count-ne1-or$receipt.receipt_id-ne$identity[0].receipt_id){throw 'A counterbalanced role used a different executable'}
  $snapshot=Get-Content (Join-Path $packet "$($phase.phase)/workspace-final.json") -Raw|ConvertFrom-Json
  if($snapshot.snapshot_id-ne$receipt.source_snapshot_id){throw 'Run source differs from compiled source'}
  foreach($compiledInput in $identity[0].compiled_inputs){
   $entry=@($snapshot.files|Where-Object path -eq $compiledInput.path)
   if($entry.Count-ne1-or$entry[0].worktree_sha256-ne$compiledInput.sha256){throw 'Compiled C++ identity differs'}
   if($phase.role-eq'candidate'-and(Get-FileHash -LiteralPath (Join-Path $repo $compiledInput.path)).Hash-ne$compiledInput.sha256){throw 'User-facing C++ differs from the qualified candidate'}
  }
  $raw=@(Get-Content (Join-Path $root "$($phase.phase)/perf/perf_metrics.jsonl")|ForEach-Object{$_|ConvertFrom-Json})
  $requests=@($raw|Where-Object metric -eq 'folder.sort_product_paint_us')
  $writers=@($raw|Where-Object metric -eq 'folder.sort_jsonl_write_us')
  $warms=@($raw|Where-Object metric -eq 'folder.sort_product_warm_calls')
  $samples=@($data.samples|Where-Object phase -eq $phase.phase)
  if($requests.Count-ne50-or$writers.Count-ne50-or$warms.Count-ne50-or$samples.Count-ne50){throw 'Incomplete request samples'}
  for($i=0;$i-lt50;$i++){
   $sample=$samples[$i]
   if($sample.duration_us-ne$requests[$i].durationUs-or$sample.writer_us-ne$writers[$i].durationUs-or$sample.remaining_us-ne([long]$requests[$i].durationUs-[long]$writers[$i].durationUs)-or$sample.writer_calls-ne$writers[$i].value0-or$sample.render_calls-ne$requests[$i].value1-or$sample.warm_calls-ne$warms[$i].value0-or$sample.successful_presents-ne$warms[$i].value1){throw 'A per-request value changed'}
  }
  $prepRaw=@($raw|Where-Object {($_.metric-eq'FolderView.ExecuteEnumeration.NaturalSortRanks'-and$_.value0-eq5000)-or($_.metric-eq'FolderView.ExecuteEnumeration.SortMerge'-and$_.value1-eq5000)})
  $prepSamples=@($data.preparation_samples|Where-Object phase -eq $phase.phase)
  if($prepRaw.Count-ne$prepSamples.Count){throw 'Preparation row coverage differs'}
  for($i=0;$i-lt$prepRaw.Count;$i++){
   if($prepRaw[$i].metric-ne$prepSamples[$i].metric-or$prepRaw[$i].durationUs-ne$prepSamples[$i].duration_us-or$prepRaw[$i].value0-ne$prepSamples[$i].value0-or$prepRaw[$i].value1-ne$prepSamples[$i].value1){throw 'Preparation row order/value differs'}
  }
  $o=Get-Content (Join-Path $packet "$($phase.phase)/observation.json") -Raw|ConvertFrom-Json
  if($o.observer_sha256-ne(Get-FileHash -LiteralPath (Join-Path $packet 'run-isolated-benchmark.ps1')).Hash-or$o.scanner_sha256-ne(Get-FileHash -LiteralPath (Join-Path $packet 'scan-processes.cs')).Hash){throw 'Observer/scanner version changed between roles'}
  $pollText=$decodedObservers["$($phase.phase)/build-process-samples.jsonl.gz"]
  if([string]::IsNullOrWhiteSpace($pollText)){throw 'Complete process observations are missing'}
  $polls=@($pollText-split'\r?\n'|Where-Object Length -gt 0|ForEach-Object{$_|ConvertFrom-Json}|Where-Object stage -eq 'benchmark')
  if(!$o.build_free_eligible-or$polls.Count-ne$o.measurement_sample_count-or@($polls|Where-Object {$_.tools.Count}).Count-or@($polls|Where-Object gap_ms -gt 1000).Count){throw 'Process observer coverage/eligibility differs'}
  $native=Get-Content (Join-Path $packet "$($phase.phase)/commands/results.json") -Raw|ConvertFrom-Json
  if($native.passed-ne5-or$native.failed-or$native.skipped){throw 'Benchmark native cases failed'}
 }
 foreach($group in $data.statistics){
  $samples=@($data.samples|Where-Object {$_.role-eq$group.role-and$_.request_index-eq$group.request_index})
  [long[]]$paint=@($samples.duration_us);[array]::Sort($paint)
  [long[]]$writer=@($samples.writer_us);[array]::Sort($writer)
  [long[]]$remaining=@($samples.remaining_us);[array]::Sort($remaining)
  if($paint.Length-ne10-or$group.paint.median_us-ne([double]$paint[4]+$paint[5])/2-or$group.writer.median_us-ne([double]$writer[4]+$writer[5])/2-or$group.remaining.median_us-ne([double]$remaining[4]+$remaining[5])/2-or$group.paint.min_us-ne$paint[0]-or$group.paint.max_us-ne$paint[9]){throw 'Grouped statistics differ from every raw request'}
 }
 foreach($group in $data.preparation_statistics){
  [long[]]$values=@($data.preparation_samples|Where-Object {$_.role-eq$group.role-and$_.metric-eq$group.metric}|ForEach-Object duration_us)
  [array]::Sort($values)
  if(!$values.Length){throw 'Preparation distribution is empty'}
  $middle=[int][Math]::Floor($values.Length/2)
  $median=if($values.Length%2){[double]$values[$middle]}else{([double]$values[$middle-1]+$values[$middle])/2}
  if($group.duration.count-ne$values.Length-or$group.duration.median_us-ne$median-or$group.duration.min_us-ne$values[0]-or$group.duration.max_us-ne$values[-1]){throw 'Preparation distribution differs from original rows'}
 }
 $native=Get-Content (Join-Path $packet "functional-final-$($profile.ToLowerInvariant())-01/commands/results.json") -Raw|ConvertFrom-Json
 if($native.passed-ne10-or$native.failed-or$native.skipped){throw 'Final native correctness failed'}
 [ordered]@{configuration=$profile;status='passed';bindings=$binding.bindings.Count;blocks=$data.phases.Count;request_samples=$data.samples.Count;scope='All original/extract bindings and metric-family row coverage/order, every request/counter, source/receipt identity, complete measurement snapshots, native verdicts and independently recomputed medians verified.'}
})
$result=[ordered]@{verified_utc=[DateTime]::UtcNow.ToString('o');status='passed';profiles=$verified}
$result|ConvertTo-Json -Depth 8|Set-Content (Join-Path $root 'independent-verification.json')
foreach($profile in @('Debug','Release')){Copy-Item -LiteralPath (Join-Path $root 'independent-verification.json') -Destination (Join-Path $repo "Specs/TestRuns/Local-x64/Navigation/2026-10-20_DebugSortWriter-$profile/independent-verification.json")}
$result|ConvertTo-Json -Depth 8
