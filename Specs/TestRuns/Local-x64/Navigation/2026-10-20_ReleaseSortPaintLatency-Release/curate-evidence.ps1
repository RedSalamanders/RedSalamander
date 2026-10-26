$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$repo='Z:/src/RedSalamander'
$baseline=Get-Content (Join-Path $root 'accepted-baselines.json') -Raw|ConvertFrom-Json
$candidate=Get-Content (Join-Path $root 'accepted-candidates.json') -Raw|ConvertFrom-Json
$keys=@('name','extension','time','size','none','name','extension','time','size','none')
$directions=@('ascending','ascending','descending','descending','ascending','descending','descending','ascending','ascending','ascending')
$pattern='"metric"\s*:\s*"(?:folder\.sort(?:_toggle_us|_product_paint_us|_product_warm_calls|_legacy_warm_calls)|FolderView\.ApplyCurrentSort(?:\.(?:OrderItems|RowMoves|ScratchBytes))?|FolderView\.ExecuteEnumeration\.(?:BuildItems|SortMerge|NaturalSortRanks)|render\.(?:present_us|begin_to_enddraw_us|layout_items_us)|dwrite\.text_layout\.frame_create_(?:us|count))"'
function Stats([object[]]$values){
 $v=@($values|Sort-Object)
 if(!$v.Count){return $null}
 $m=[int][Math]::Floor($v.Count/2)
 $median=if($v.Count%2){$v[$m]}else{([decimal]$v[$m-1]+[decimal]$v[$m])/2}
 [ordered]@{count=$v.Count;min_us=$v[0];median_us=$median;max_us=$v[-1]}
}
foreach($configuration in @('Debug','Release')){
 $packet=Join-Path $repo "Specs/TestRuns/Local-x64/Navigation/2026-10-20_ReleaseSortPaintLatency-$configuration"
 New-Item -ItemType Directory -Path $packet -Force|Out-Null
 $copies=[Collections.Generic.List[object]]::new()
 $extracts=[Collections.Generic.List[object]]::new()
 function Copy-Bound([string]$source,[string]$relative){
  $target=Join-Path $packet $relative
  New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force|Out-Null
  Copy-Item -LiteralPath $source -Destination $target
  $hash=(Get-FileHash -LiteralPath $source).Hash
  if((Get-FileHash -LiteralPath $target).Hash-ne$hash){throw 'Copy digest differs'}
  $copies.Add([ordered]@{source=$source;archive=$relative;sha256=$hash;bytes=(Get-Item -LiteralPath $source).Length})
 }
 function Extract-Bound([string]$source,[string]$relative,[string[]]$lines,[int]$rawCount,[string]$policy){
  $target=Join-Path $packet $relative
  New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force|Out-Null
  [IO.File]::WriteAllLines($target,$lines,[Text.UTF8Encoding]::new($false))
  $extracts.Add([ordered]@{source=$source;source_sha256=(Get-FileHash -LiteralPath $source).Hash;source_bytes=(Get-Item -LiteralPath $source).Length;source_rows=$rawCount;archive=$relative;archive_sha256=(Get-FileHash -LiteralPath $target).Hash;archive_rows=$lines.Count;policy=$policy})
 }
 $accepted=@($baseline.$configuration)+@($candidate.blocks|Where-Object configuration -eq $configuration|ForEach-Object phase)
 $phases=[ordered]@{}
 foreach($dir in @(Get-ChildItem -LiteralPath $root -Directory|Where-Object {Test-Path -LiteralPath (Join-Path $_.FullName 'observation.json')}|Sort-Object Name)){
  $o=Get-Content (Join-Path $dir.FullName 'observation.json') -Raw|ConvertFrom-Json
  if($o.configuration-ne$configuration){continue}
  $phase=$dir.Name
  $selected=$phase-in$accepted
  if($selected-and!$o.build_free_eligible){throw 'Selected phase fails observer gate'}
  $binding=Get-Content (Join-Path $dir.FullName 'bindings.json') -Raw|ConvertFrom-Json
  $receipt=Get-Content (Join-Path $dir.FullName 'build-receipt.json') -Raw|ConvertFrom-Json
  $receiptPath="receipts/$($receipt.receipt_id).json"
  if(!(Test-Path -LiteralPath (Join-Path $packet $receiptPath))){Copy-Bound (Join-Path $dir.FullName 'build-receipt.json') $receiptPath}
  [ordered]@{receipt_id=$receipt.receipt_id;archive=$receiptPath;source_sha256=(Get-FileHash -LiteralPath (Join-Path $dir.FullName 'build-receipt.json')).Hash}|ConvertTo-Json|Set-Content (Join-Path $packet "$phase-receipt-ref.json")
  foreach($relative in @('observation.json','bindings.json','results.json','run-all-tests-results.json','workspace-final.json','trace.txt','commands/results.json','commands/trace.txt')){Copy-Bound (Join-Path $dir.FullName $relative) "$phase/$relative"}
  $control=$phase-match'-legacy-'
  $artifact=if($control){'folderView_perf_sort_toggle_legacy_warm_control_metrics.json'}else{'folderView_perf_sort_toggle_stress_metrics.json'}
  Copy-Bound (Join-Path $dir.FullName "perf/$artifact") "$phase/perf/$artifact"
  $rawPath=Join-Path $dir.FullName 'perf/perf_metrics.jsonl'
  $raw=[IO.File]::ReadAllLines($rawPath)
  $lines=[Collections.Generic.List[string]]::new();$rows=[Collections.Generic.List[object]]::new()
  foreach($line in $raw){if($line-match$pattern){$lines.Add($line);$rows.Add(($line|ConvertFrom-Json))}}
  Extract-Bound $rawPath "$phase/perf/perf-selected.jsonl" $lines.ToArray() $raw.Count 'Every row from the named sorting timer/counter, ApplyCurrentSort, enumeration/preparation, layout-total, drawing, presentation and frame text-creation families, with original line content/order. Detailed layout phases and duplicate whole-frame totals remain in the digest-bound external raw file.'
  $metric=if($phase-match'-product-'){'folder.sort_product_paint_us'}else{'folder.sort_toggle_us'}
  $sorts=@($rows|Where-Object metric -eq $metric)
  if($sorts.Count-ne50){throw 'Five repetitions require 50 sort samples'}
  for($i=0;$i-lt50;$i++){if($sorts[$i].detail-ne$keys[$i%10]-or$sorts[$i].value0-ne5000){throw 'Fixture request order/count differs'}}
  $warmMetric=if($metric-eq'folder.sort_product_paint_us'){'folder.sort_product_warm_calls'}else{'folder.sort_legacy_warm_calls'}
  $warm=@($rows|Where-Object metric -eq $warmMetric)
  if($phase-like'candidate-*'){
   if($warm.Count-ne50){throw 'Candidate timing counter coverage differs'}
   foreach($row in $warm){if($row.durationUs-ne0){throw 'Counter has a duration'}}
   if($metric-eq'folder.sort_product_paint_us'){
    for($i=0;$i-lt50;$i++){if($warm[$i].value0-ne0-or$(if($i%10-eq0){$sorts[$i].value1-ne0-or$warm[$i].value1-ne0}else{$warm[$i].value1-lt1})){throw 'Product timing manufactured warming/no-op frames or accepted no successful paint'}}
   }elseif(@($warm|Where-Object value0 -ne 1).Count){throw 'Legacy control did not retain one warm call per request'}
  }
  $monitorPath=Join-Path $dir.FullName 'build-process-samples.jsonl'
  $monitor=[IO.File]::ReadAllLines($monitorPath)
  $kept=[Collections.Generic.List[string]]::new();$observations=[Collections.Generic.List[object]]::new()
  $gateEnd=([DateTime]$o.gate_ended_utc).ToUniversalTime()
  foreach($line in $monitor){$entry=$line|ConvertFrom-Json;if($entry.stage-eq'benchmark'){$kept.Add($line);$observations.Add($entry)}elseif((([DateTime]$entry.utc).ToUniversalTime()-$gateEnd).TotalSeconds-ge-31){$kept.Add($line)}}
  Extract-Bound $monitorPath "$phase/build-process-qualifying-samples.jsonl" $kept.ToArray() $monitor.Count 'Every measurement snapshot and final 31 seconds of gate snapshots, preserving original line order/content. Earlier gate activity remains externally digest-bound.'
  $cases=Get-Content (Join-Path $dir.FullName 'commands/results.json') -Raw|ConvertFrom-Json
  if($cases.passed-ne5-or$cases.failed-or$cases.skipped){throw 'Benchmark native verdict failed'}
  $core=@($rows|Where-Object {$_.metric-eq'FolderView.ApplyCurrentSort.OrderItems'-and$_.value0-eq5000}|ForEach-Object durationUs)
  $phases[$phase]=[ordered]@{selected=$selected;role=$(if($phase-like'baseline-*'){'fresh legacy baseline'}elseif($control){'candidate legacy control'}else{'candidate product paint'});protocol=$(if($metric-eq'folder.sort_product_paint_us'){'product_paint_v2'}else{'legacy_warm_v1'});configuration=$configuration;run_id=$o.run_id;receipt_id=$receipt.receipt_id;all_sort_samples=50;core_ordering=(Stats $core);samples=@(for($i=0;$i-lt50;$i++){[ordered]@{repeat=[int][Math]::Floor($i/10)+1;request_index=$i%10+1;key=$keys[$i%10];direction=$directions[$i%10];duration_us=$sorts[$i].durationUs;render_calls=$sorts[$i].value1;warm_calls=$(if($warm.Count){$warm[$i].value0}else{$null});successful_presents=$(if($warm.Count){$warm[$i].value1}else{$null})}});measurement_samples=$observations.Count;build_overlap_polls=$o.measurement_contaminated_polls;maximum_poll_gap_ms=$o.maximum_measurement_gap_ms;observer_cpu_ms=$o.measurement_observer_cpu_ms}
 }
 foreach($entry in $phases.Keys){
  $phases[$entry]['observer_scope']=if($entry-like'baseline-*'){'v1 build-only observation; absence of concurrent native test runtimes not established'}else{'v2 build and known native-test-runtime observation'}
 }
 $roleStats=[ordered]@{}
 foreach($role in @('fresh legacy baseline','candidate legacy control','candidate product paint')){
  $selectedPhases=@($phases.Values|Where-Object {$_.selected-and$_.role-eq$role})
  $roleStats[$role]=@(for($i=0;$i-lt10;$i++){
   $samples=@($selectedPhases|ForEach-Object samples|Where-Object request_index -eq ($i+1))
   [ordered]@{request_index=$i+1;key=$keys[$i];direction=$directions[$i];stats=(Stats @($samples.duration_us));render_calls=@($samples.render_calls);warm_calls=@($samples.warm_calls);successful_presents=@($samples.successful_presents)}
  })
 }
 [ordered]@{configuration=$configuration;fixture='Unchanged 5000 files and ten requests';selection='First native/observer-eligible attempts, never latency selection';counterbalance=$candidate.sequence;role_statistics=$roleStats;phases=$phases;scope='The new product-paint protocol corrects timed synthetic work. Its difference from legacy warm timing is a measurement-protocol difference, not a production speedup. Core ordering, layout and presentation remain separate measurements.'}|ConvertTo-Json -Depth 16|Set-Content (Join-Path $packet 'performance-comparison.json')
 $functional="functional-$($configuration.ToLowerInvariant())-01"
 foreach($relative in @('bindings.json','results.json','run-all-tests-results.json','workspace-final.json','trace.txt','commands/results.json','commands/trace.txt')){Copy-Bound (Join-Path $root "$functional/$relative") "$functional/$relative"}
 foreach($name in @('run-isolated-benchmark.ps1','run-legacy-control-benchmark.ps1','run-candidates.ps1','scan-processes.cs','observer-control-equivalence.json','ordering-and-fixture-checks.json','candidate-source-copies.json','initial-baseline-identities.json','tooling-pester-isolated.json','tooling-pester-isolated.xml')){Copy-Bound (Join-Path $root $name) $name}
 foreach($name in @('v1-run-isolated-benchmark.ps1','v1-run-legacy-control-benchmark.ps1','v1-scan-processes.cs')){Copy-Bound (Join-Path $root $name) $name}
 foreach($name in @('source-and-format-proofs.json','source-and-format-proofs.ps1','final-identities.json','final-identities-check.ps1','sandbox-audit-comparison.json','tooling-pester-main.json','tooling-pester-main.xml','curate-evidence.ps1','verify-evidence.ps1','render-report.ps1')){Copy-Bound (Join-Path $root $name) $name}
 foreach($name in @('original-attribution-bindings.json','paired-baseline-release-01-stage-attribution.json','paired-candidate-release-01-stage-attribution.json')){Copy-Bound (Join-Path $root $name) $name}
 foreach($name in @('probe-result.json','probe-context.json','build-process-samples.jsonl','observed-build-processes.jsonl')){Copy-Bound (Join-Path $root "observer-runtime-positive-control/$name") "observer-runtime-positive-control/$name"}
 foreach($dir in @(Get-ChildItem -LiteralPath $root -Directory|Where-Object {Test-Path -LiteralPath (Join-Path $_.FullName 'setup-retirement.json')})){
  Copy-Bound (Join-Path $dir.FullName 'setup-retirement.json') "$($dir.Name)/setup-retirement.json"
  $path=Join-Path $dir.FullName 'build-process-samples.jsonl';$raw=[IO.File]::ReadAllLines($path)
  $tail=@($raw|Select-Object -Last 8|ForEach-Object {try{$null=$_|ConvertFrom-Json;$_}catch{}})
  if($tail.Count){Extract-Bound $path "$($dir.Name)/gate-tail.jsonl" $tail $raw.Count 'Last eight available gate lines retaining only complete JSON objects. An owned gate retired before any benchmark is diagnostic only and may end with a buffered partial row; full raw bytes remain externally digest-bound.'}
 }
 if($configuration-eq'Debug'){
  foreach($relative in @('bindings.json','results.json','run-all-tests-results.json','commands/results.json','commands/trace.txt','perf/folderView_perf_sort_toggle_stress_metrics.json')){Copy-Bound (Join-Path $root "smoke-debug-product-01/$relative") "smoke-debug-product-01/$relative"}
 }
 foreach($name in @('candidate-build-debug.txt','candidate-build-release.txt','spec-inventory-isolated.txt','tool-inventory-isolated.txt')){$path=Join-Path $root $name;$raw=[IO.File]::ReadAllLines($path);Extract-Bound $path "validation-tails/$name" @($raw|Select-Object -Last 24) $raw.Count 'Last 24 original validation lines; full output remains externally digest-bound.'}
 foreach($name in @('spec-inventory-main.txt','tool-inventory-main.txt')){$path=Join-Path $root $name;$raw=[IO.File]::ReadAllLines($path);Extract-Bound $path "validation-tails/$name" @($raw|Select-Object -Last 24) $raw.Count 'Last 24 original final main metadata validation lines; full output remains externally digest-bound.'}
 [ordered]@{original_copies=$copies;derived_extracts=$extracts}|ConvertTo-Json -Depth 10|Set-Content (Join-Path $packet 'artifact-bindings.json')
 $files=@(Get-ChildItem -LiteralPath $packet -File -Recurse)
 $bytes=($files|Measure-Object -Property Length -Sum).Sum
 if(@($files|Where-Object Length -gt 2MB).Count-or$bytes-gt5MB){throw "Governed archive size exceeded: $configuration $bytes"}
 [pscustomobject]@{configuration=$configuration;files=$files.Count;bytes=$bytes;selected_phases=$accepted.Count}
}
