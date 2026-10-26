$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$accepted=Get-Content (Join-Path $root 'accepted-blocks.json') -Raw|ConvertFrom-Json
function Get-Statistics($values){
 $ordered=@($values|Sort-Object)
 if(!$ordered.Count){return $null}
 $middle=[int][Math]::Floor($ordered.Count/2)
 $median=if($ordered.Count%2){$ordered[$middle]}else{([double]$ordered[$middle-1]+[double]$ordered[$middle])/2}
 [ordered]@{count=$ordered.Count;median_us=$median;min_us=$ordered[0];max_us=$ordered[-1]}
}
$samples=[Collections.Generic.List[object]]::new()
$phases=[Collections.Generic.List[object]]::new()
$preparation=[Collections.Generic.List[object]]::new()
$keys=@('name','extension','time','size','none','name','extension','time','size','none')
$directions=@('ascending','ascending','descending','descending','ascending','descending','descending','ascending','ascending','ascending')
foreach($block in $accepted.blocks){
 $phaseRoot=Join-Path $root $block.phase
 $o=Get-Content (Join-Path $phaseRoot 'observation.json') -Raw|ConvertFrom-Json
 $r=Get-Content (Join-Path $phaseRoot 'commands/results.json') -Raw|ConvertFrom-Json
 if(!$o.build_free_eligible-or$o.measurement_contaminated_polls-or$o.maximum_measurement_gap_ms-gt1000-or$r.passed-ne5-or$r.failed-or$r.skipped){throw "Ineligible block $($block.phase)"}
 $rawPath=Join-Path $phaseRoot 'perf/perf_metrics.jsonl'
 $rows=@(Get-Content -LiteralPath $rawPath|ForEach-Object{$_|ConvertFrom-Json})
 foreach($row in $rows){
  if(($row.metric-eq'FolderView.ExecuteEnumeration.NaturalSortRanks'-and$row.value0-eq5000)-or($row.metric-eq'FolderView.ExecuteEnumeration.SortMerge'-and$row.value1-eq5000)){
   $preparation.Add([pscustomobject]@{phase=$block.phase;configuration=$block.configuration;role=$block.role;metric=$row.metric;duration_us=$row.durationUs;value0=$row.value0;value1=$row.value1})
  }
 }
 $requestIndices=@(for($i=0;$i-lt$rows.Count;$i++){if($rows[$i].metric-eq'folder.sort_product_paint_us'){$i}})
 $writerIndices=@(for($i=0;$i-lt$rows.Count;$i++){if($rows[$i].metric-eq'folder.sort_jsonl_write_us'){$i}})
 $warms=@($rows|Where-Object metric -eq 'folder.sort_product_warm_calls')
 if($requestIndices.Count-ne50-or$writerIndices.Count-ne50-or$warms.Count-ne50){throw 'Incomplete request/observer metrics'}
 $previousWriter=-1
 for($i=0;$i-lt50;$i++){
  $request=$rows[$requestIndices[$i]]
  $writer=$rows[$writerIndices[$i]]
  $requestIndex=$i%10
  if($request.detail-ne$keys[$requestIndex]-or$writer.detail-ne$request.detail-or$request.value0-ne5000-or$writerIndices[$i]-le$requestIndices[$i]){throw 'Fixture/request order changed'}
  $frameContract=if($requestIndex-eq0){$request.value1-eq0-and$warms[$i].value1-eq0}else{$request.value1-ge1-and$warms[$i].value1-ge1-and$request.value1-ge$warms[$i].value1}
  if(!$frameContract-or$warms[$i].value0-ne0-or$writer.value1-ne$request.value1){throw 'Product paint/warm counters changed'}
  if($writer.durationUs-gt$request.durationUs){throw 'Writer attribution exceeds enclosing request'}
  $stages=@{}
  if($requestIndex-ne0){
   for($j=$previousWriter+1;$j-lt$requestIndices[$i];$j++){
    $metric=$rows[$j].metric
    if($metric -match '^(FolderView\.ApplyCurrentSort(?:\.|$)|render\.(layout_items_us|begin_to_enddraw_us|present_us|draw_item_us)$|folder\.layout\.)'){
     if(!$stages.ContainsKey($metric)){$stages[$metric]=0L}
     $stages[$metric]+=[long]$rows[$j].durationUs
    }
   }
  }
  $previousWriter=$writerIndices[$i]
  $samples.Add([pscustomobject]@{phase=$block.phase;configuration=$block.configuration;role=$block.role;block=$block.block;repeat=([int][Math]::Floor($i/10)+1);request_index=($requestIndex+1);key=$request.detail;direction=$directions[$requestIndex];duration_us=$request.durationUs;writer_us=$writer.durationUs;remaining_us=([long]$request.durationUs-[long]$writer.durationUs);writer_calls=$writer.value0;render_calls=$request.value1;warm_calls=$warms[$i].value0;successful_presents=$warms[$i].value1;stages=$stages})
 }
 $phases.Add([ordered]@{phase=$block.phase;role=$block.role;configuration=$block.configuration;run_id=$o.run_id;receipt_id=(Get-Content (Join-Path $phaseRoot 'build-receipt.json') -Raw|ConvertFrom-Json).receipt_id;raw_sha256=(Get-FileHash -LiteralPath $rawPath).Hash;raw_bytes=(Get-Item -LiteralPath $rawPath).Length;raw_rows=$rows.Count;sort_samples=50;overlap_polls=$o.measurement_contaminated_polls;maximum_gap_ms=$o.maximum_measurement_gap_ms;observer_cpu_ms=$o.measurement_observer_cpu_ms;metric_counts=@($rows|Group-Object metric|Sort-Object Name|Select-Object Name,Count)})
}
$statistics=@(foreach($profile in @('Debug','Release')){foreach($role in @('baseline','candidate')){foreach($requestIndex in 1..10){
 $group=@($samples|Where-Object {$_.configuration-eq$profile-and$_.role-eq$role-and$_.request_index-eq$requestIndex})
 if($group.Count-ne10){throw 'Counterbalanced sample group incomplete'}
 $stageNames=@($group|ForEach-Object {$_.stages.Keys}|Sort-Object -Unique)
 $stageStats=[ordered]@{}
 foreach($stage in $stageNames){$stageStats[$stage]=Get-Statistics @($group|ForEach-Object {$_.stages[$stage]})}
 [ordered]@{configuration=$profile;role=$role;request_index=$requestIndex;key=$group[0].key;direction=$group[0].direction;paint=Get-Statistics @($group.duration_us);writer=Get-Statistics @($group.writer_us);remaining=Get-Statistics @($group.remaining_us);writer_calls=@($group.writer_calls);stages=$stageStats}
}}})
$preparationStatistics=@(foreach($profile in @('Debug','Release')){foreach($role in @('baseline','candidate')){foreach($metric in @('FolderView.ExecuteEnumeration.NaturalSortRanks','FolderView.ExecuteEnumeration.SortMerge')){
 $group=@($preparation|Where-Object {$_.configuration-eq$profile-and$_.role-eq$role-and$_.metric-eq$metric})
 [ordered]@{configuration=$profile;role=$role;metric=$metric;duration=Get-Statistics @($group.duration_us)}
}}})
[ordered]@{protocol='Unchanged product_paint_v2; identical test-only writer observation';selection=$accepted.selection;sequence=$accepted.sequence;phases=$phases.ToArray();statistics=$statistics;samples=$samples.ToArray();preparation_samples=$preparation.ToArray();preparation_statistics=$preparationStatistics;scope='Writer time is this UI thread inclusive elapsed time, including contention; it is not CPU time or uncaptured latency. Nested stages are not additive. Preparation retains every 5000-item enumeration rank/sort-merge record separately from request latency; it is not end-to-end folder-open time. All original metric rows and rejected attempts remain retained.'}|ConvertTo-Json -Depth 12|Set-Content (Join-Path $root 'performance-comparison.json')
$statistics|Where-Object request_index -eq 2|ForEach-Object {[pscustomobject]@{profile=$_.configuration;role=$_.role;paint_ms=$_.paint.median_us/1000;writer_ms=$_.writer.median_us/1000;minimum_ms=$_.paint.min_us/1000;maximum_ms=$_.paint.max_us/1000}}|Format-Table|Out-String|Write-Output
