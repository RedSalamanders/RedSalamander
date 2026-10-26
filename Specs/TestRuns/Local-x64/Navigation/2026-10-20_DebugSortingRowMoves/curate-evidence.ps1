$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$packet='Z:/src/RedSalamander/Specs/TestRuns/Local-x64/Navigation/2026-10-20_DebugSortingRowMoves'
New-Item -ItemType Directory -Path $packet -Force|Out-Null
$copies=[Collections.Generic.List[object]]::new()
$extracts=[Collections.Generic.List[object]]::new()
function Copy-Evidence([string]$source,[string]$relative){
 if(!(Test-Path -LiteralPath $source)){throw "Missing evidence: $source"}
 $target=Join-Path $packet $relative
 New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force|Out-Null
 Copy-Item -LiteralPath $source -Destination $target
 $hash=(Get-FileHash -LiteralPath $source).Hash
 if((Get-FileHash -LiteralPath $target).Hash-ne$hash){throw 'Copy hash mismatch'}
 $copies.Add([ordered]@{source=$source;archive=$relative;sha256=$hash;bytes=(Get-Item -LiteralPath $source).Length})
}
function Write-Extract([string]$source,[string]$relative,[string[]]$lines,[string]$policy,[int]$rawRows){
 $target=Join-Path $packet $relative
 New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force|Out-Null
 [IO.File]::WriteAllLines($target,$lines,[Text.UTF8Encoding]::new($false))
 $extracts.Add([ordered]@{source=$source;source_sha256=(Get-FileHash -LiteralPath $source).Hash;source_bytes=(Get-Item -LiteralPath $source).Length;source_rows=$rawRows;archive=$relative;archive_sha256=(Get-FileHash -LiteralPath $target).Hash;archive_rows=$lines.Count;policy=$policy})
}
function Get-Stats([object[]]$values){
 $v=@($values|Sort-Object)
 if(!$v.Count){return $null}
 $middle=[int][Math]::Floor($v.Count/2)
 $median=if($v.Count%2){$v[$middle]}else{([decimal]$v[$middle-1]+[decimal]$v[$middle])/2}
 return [ordered]@{count=$v.Count;min_us=$v[0];median_us=$median;max_us=$v[-1]}
}
$baseline=Get-Content (Join-Path $root 'accepted-baselines.json') -Raw|ConvertFrom-Json
$candidate=Get-Content (Join-Path $root 'accepted-candidates.json') -Raw|ConvertFrom-Json
$accepted=@($baseline.Debug,$baseline.Release,$candidate.Debug,$candidate.Release)
if(@($accepted|Where-Object {[string]::IsNullOrWhiteSpace($_)}).Count){throw 'Accepted pair missing'}
$keys=@('Name','Extension','Time','Size','None','Name','Extension','Time','Size','None')
$directions=@('Ascending','Ascending','Descending','Descending','Ascending','Descending','Descending','Ascending','Ascending','Ascending')
$phases=[ordered]@{}
$pattern='"metric"\s*:\s*"(?:folder\.sort_toggle_us|FolderView\.ApplyCurrentSort(?:\.(?:OrderItems|RowMoves|ScratchBytes))?|FolderView\.ExecuteEnumeration\.(?:BuildItems|SortMerge|NaturalSortRanks))"'
foreach($dir in @(Get-ChildItem -LiteralPath $root -Directory|Where-Object {Test-Path -LiteralPath (Join-Path $_.FullName 'observation.json')}|Sort-Object Name)){
 $phase=$dir.Name
 $observation=Get-Content (Join-Path $dir.FullName 'observation.json') -Raw|ConvertFrom-Json
 $isAccepted=$phase-in$accepted
 if($isAccepted-and!$observation.build_free_eligible){throw 'Accepted contaminated attempt'}
 foreach($relative in @('observation.json','bindings.json','build-receipt.json','results.json','run-all-tests-results.json','workspace-final.json','trace.txt','commands/results.json','commands/trace.txt','perf/folderView_perf_sort_toggle_stress_metrics.json')){
  Copy-Evidence (Join-Path $dir.FullName $relative) "$phase/$relative"
 }
 Copy-Evidence (Join-Path $dir.FullName 'runner.log') "$phase/runner.txt"
 $rawPath=Join-Path $dir.FullName 'perf/perf_metrics.jsonl'
 $raw=[IO.File]::ReadAllLines($rawPath)
 $selected=[Collections.Generic.List[string]]::new()
 $rows=[Collections.Generic.List[object]]::new()
 foreach($line in $raw){if($line-match$pattern){$selected.Add($line);$rows.Add(($line|ConvertFrom-Json))}}
 Write-Extract $rawPath "$phase/perf/perf-selected.jsonl" $selected.ToArray() 'Every row from all eight named sorting metric families, byte-identical line content in original order; other metric families remain in the digest-bound external raw file.' $raw.Count
 $sorts=@($rows|Where-Object metric -eq 'folder.sort_toggle_us')
 if($sorts.Count-ne50){throw 'Five complete repetitions require 50 observations'}
 $cases=Get-Content (Join-Path $dir.FullName 'commands/results.json') -Raw|ConvertFrom-Json
 if($cases.passed-ne5-or$cases.failed-or$cases.skipped){throw 'Benchmark cases did not all pass'}
 $requests=@(for($index=0;$index-lt10;$index++){
  $values=@(for($repeat=0;$repeat-lt5;$repeat++){$row=$sorts[$index+10*$repeat];if($row.detail-ne$keys[$index]-or$row.value0-ne5000){throw 'Fixture request/count differs'};$row.durationUs})
  [ordered]@{request_index=$index+1;key=$keys[$index];direction=$directions[$index];stats=(Get-Stats $values)}
 })
 $stages=[ordered]@{}
 foreach($metric in @('FolderView.ApplyCurrentSort','FolderView.ApplyCurrentSort.OrderItems','FolderView.ExecuteEnumeration.BuildItems','FolderView.ExecuteEnumeration.NaturalSortRanks','FolderView.ExecuteEnumeration.SortMerge')){
  $values=@($rows|Where-Object {$_.metric-eq$metric-and$(if($metric-eq'FolderView.ExecuteEnumeration.SortMerge'){$_.value0-eq0-and$_.value1-eq5000}else{$_.value0-eq5000})}|ForEach-Object durationUs)
  $stages[$metric]=Get-Stats $values
 }
 $moveRows=@($rows|Where-Object metric -eq 'FolderView.ApplyCurrentSort.RowMoves')
 $scratchRows=@($rows|Where-Object metric -eq 'FolderView.ApplyCurrentSort.ScratchBytes')
 foreach($row in $moveRows){if($row.durationUs-ne0-or$row.value1-gt$row.value0+[Math]::Floor($row.value0/2)){throw 'Payload movement bound failed'}}
 foreach($row in $scratchRows){if($row.durationUs-ne0-or$row.value1-gt16*$row.value0-or$row.value1-lt8*$row.value0){throw 'Index scratch bound failed'}}
 $monitorPath=Join-Path $dir.FullName 'build-process-samples.jsonl'
 $monitor=[IO.File]::ReadAllLines($monitorPath)
 $qualifying=[Collections.Generic.List[string]]::new()
 $benchmark=@()
 $gateEnd=([DateTime]$observation.gate_ended_utc).ToUniversalTime()
 foreach($line in $monitor){$r=$line|ConvertFrom-Json;if($r.stage-eq'benchmark'){$qualifying.Add($line);$benchmark+=$r}elseif((([DateTime]$r.utc).ToUniversalTime()-$gateEnd).TotalSeconds-ge-31){$qualifying.Add($line)}}
 Write-Extract $monitorPath "$phase/build-process-qualifying-samples.jsonl" $qualifying.ToArray() 'All benchmark snapshots plus the final 31 seconds of gate snapshots, in original order. Earlier gate activity remains in the digest-bound raw file.' $monitor.Count
 $observed=Join-Path $dir.FullName 'observed-build-processes.jsonl'
 if(Test-Path -LiteralPath $observed){Copy-Evidence $observed "$phase/observed-build-processes.jsonl"}
 $contaminated=@($benchmark|Where-Object {$_.tools.Count}).Count
 if($benchmark.Count-ne$observation.measurement_sample_count-or$contaminated-ne$observation.measurement_contaminated_polls){throw 'Observer counts differ from original rows'}
 $phases[$phase]=[ordered]@{accepted=$isAccepted;disposition=$(if($isAccepted){'selected stable-checkout pair'}elseif($observation.build_free_eligible){'eligible supplemental run superseded by source-isolation decision'}else{'rejected by observation predicates'});configuration=$observation.configuration;run_id=$observation.run_id;receipt_id=(Get-Content (Join-Path $dir.FullName 'bindings.json') -Raw|ConvertFrom-Json).receipt_id;cases_passed=$cases.passed;all_sort_samples=$sorts.Count;samples=@(for($i=0;$i-lt$sorts.Count;$i++){[ordered]@{repeat=[int][Math]::Floor($i/10)+1;request_index=$i%10+1;key=$keys[$i%10];direction=$directions[$i%10];duration_us=$sorts[$i].durationUs}});requests=$requests;stages=$stages;payload_move_records=$moveRows.Count;payload_moves_min=($moveRows.value1|Measure-Object -Minimum).Minimum;payload_moves_max=($moveRows.value1|Measure-Object -Maximum).Maximum;index_scratch_bytes_min=($scratchRows.value1|Measure-Object -Minimum).Minimum;index_scratch_bytes_max=($scratchRows.value1|Measure-Object -Maximum).Maximum;benchmark_snapshots=$benchmark.Count;contaminated_snapshots=$contaminated;maximum_poll_gap_ms=$observation.maximum_measurement_gap_ms}
}
$pairs=[ordered]@{}
foreach($configuration in @('Debug','Release')){
 $b=$phases[$baseline.$configuration];$c=$phases[$candidate.$configuration]
 if(!$b-or!$c){throw 'Pair data missing'}
 $pairs[$configuration]=[ordered]@{baseline=$baseline.$configuration;candidate=$candidate.$configuration;requests=@(for($i=0;$i-lt10;$i++){$old=$b.requests[$i].stats.median_us;$new=$c.requests[$i].stats.median_us;[ordered]@{key=$keys[$i];direction=$directions[$i];baseline_median_us=$old;candidate_median_us=$new;change_percent=[Math]::Round(100*($new-$old)/$old,3)}});ordering=[ordered]@{baseline=$b.stages['FolderView.ApplyCurrentSort.OrderItems'];candidate=$c.stages['FolderView.ApplyCurrentSort.OrderItems']};apply_current_sort=[ordered]@{baseline=$b.stages['FolderView.ApplyCurrentSort'];candidate=$c.stages['FolderView.ApplyCurrentSort']}}
}
[ordered]@{created_utc=[DateTime]::UtcNow.ToString('o');baseline_commit='d7dd952d629985b48d25b23b56d1726fbb2a1ec5';machine_hash='4cb089111a23';fixture='Unchanged folderView_perf_sort_toggle_stress, 5000 files, ten requests, five repetitions';phase_selection='First eligible attempt by native verdict and observer predicates, never selected by latency';pairs=$pairs;phases=$phases;stage_scope='Stage records include setup, accepted-listing application and fixture restoration. Enumeration records can include canceled work. Counts are not cold-open trial counts; separate stage medians are not additive.'}|ConvertTo-Json -Depth 15|Set-Content (Join-Path $packet 'performance-comparison.json')
foreach($dir in @(Get-ChildItem -LiteralPath $root -Directory|Where-Object Name -Like 'functional-*')){
 foreach($relative in @('bindings.json','build-receipt.json','results.json','run-all-tests-results.json','workspace-final.json','commands/results.json','commands/trace.txt')){Copy-Evidence (Join-Path $dir.FullName $relative) "$($dir.Name)/$relative"}
 $tracePath=Join-Path $dir.FullName 'trace.txt'
 $trace=[IO.File]::ReadAllLines($tracePath)
 $kept=@(for($index=0;$index-lt$trace.Count;$index++){if($index-lt20-or$index-ge$trace.Count-20-or$trace[$index]-match'(?i)selftest|assert|\b(?:error|warn(?:ing)?|fail(?:ed|ure)?)\b'){$trace[$index]}})
 Write-Extract $tracePath "$($dir.Name)/trace.txt" $kept 'First and last 20 application-trace lines plus every selftest/assert/error/warning/failure line, once in original order. Complete focused commands/trace.txt and every case verdict/reason remain unabridged; the verbose application trace remains externally digest-bound.' $trace.Count
}
$retiredGates=@()
foreach($dir in @(Get-ChildItem -LiteralPath $root -Directory|Where-Object {Test-Path -LiteralPath (Join-Path $_.FullName 'setup-retirement.json')})){
 Copy-Evidence (Join-Path $dir.FullName 'setup-retirement.json') "$($dir.Name)/setup-retirement.json"
 $retired=Get-Content (Join-Path $dir.FullName 'setup-retirement.json') -Raw|ConvertFrom-Json
 if(!$retired.benchmark_started-and(Test-Path (Join-Path $dir.FullName 'runner.log'))){throw 'Retirement incorrectly claims no benchmark launch'}
 $rawPath=Join-Path $dir.FullName 'build-process-samples.jsonl'
 if(Test-Path -LiteralPath $rawPath){
  $raw=[IO.File]::ReadAllLines($rawPath)
  $valid=@($raw|Select-Object -Last 8|ForEach-Object {try{$null=$_|ConvertFrom-Json;$_}catch{}})
  $retiredGates += [ordered]@{phase=$dir.Name;benchmark_started=$retired.benchmark_started;raw_source=$rawPath;raw_sha256=(Get-FileHash -LiteralPath $rawPath).Hash;raw_bytes=(Get-Item -LiteralPath $rawPath).Length;raw_lines=$raw.Count;retained_complete_tail_rows=$valid.Count}
  if($valid.Count){Write-Extract $rawPath "$($dir.Name)/gate-final-valid-samples.jsonl" $valid 'Last eight available lines, retaining only complete JSON objects. An owned gate stopped before measurement can end with a buffered partial row; these are diagnostic gate samples, not a quiet-window qualification. Full earlier gate activity remains externally digest-bound.' $raw.Count}
 }
}
ConvertTo-Json -InputObject @($retiredGates) -Depth 6|Set-Content (Join-Path $packet 'retired-gates.json')
foreach($name in @('run-benchmark.ps1','run-main-benchmark.ps1','run-isolated-benchmark.ps1','scan-processes.cs','wait-idle-nodes.ps1','run-functional.ps1','baseline-source-hashes.json','candidate-source-copies.json','candidate-patch-check.json','non-build-process-exclusion.json','idle-node-verification-debug-baseline.json','idle-node-verification-release-baseline.json','tooling-pester.json','tooling-pester.xml','ordering-preservation-source-checks.json','format-validation.json','staging-proxy-review.json','isolated-pairing-decision.json','tooling-pester-final.json','tooling-pester-final.xml')){
 Copy-Evidence (Join-Path $root $name) $name
}
foreach($file in @(Get-ChildItem -LiteralPath $root -File|Where-Object {$_.Name-match'^(paired-preparation|release-candidate-preparation).*\.json$'})){Copy-Evidence $file.FullName $file.Name}
foreach($name in @('final-identities.json','final-sandbox-audit-comparison.json','verify-final-identities.ps1')){Copy-Evidence (Join-Path $root $name) $name}
foreach($name in @('tooling-pester-closeout.json','tooling-pester-closeout.xml','tracked-archive-closeout.json','final-closeout.json','curate-evidence.ps1','verify-evidence.ps1','render-report.ps1')){Copy-Evidence (Join-Path $root $name) $name}
foreach($name in @('spec-inventory-closeout.txt','tool-inventory-closeout.txt','archive-inventory-closeout.txt')){
 $source=Join-Path $root $name;$raw=[IO.File]::ReadAllLines($source)
 Write-Extract $source "validation-tails/$name" @($raw|Select-Object -Last 28) 'Last 28 original closeout validation lines; full output retained externally with digest and byte/line counts.' $raw.Count
}
Copy-Evidence (Join-Path $root 'isolated-observer-control-01/probe-result.json') 'observer-control/probe-result.json'
foreach($name in @('spec-inventory.txt','tool-inventory.txt','archive-inventory.txt','baseline-build-Debug.log','baseline-build-Release.log','baseline-build-Release-retry.log','baseline-build-Release-isolated.log','candidate-build-Debug.log','candidate-build-Release-main.log','candidate-build-Release-worktree.log','candidate-build-Debug-worktree.log','baseline-build-Debug-worktree.log','candidate-v2-main-Debug-attestation-failed.log','spec-inventory-final.txt','tool-inventory-final.txt')){
 $source=Join-Path $root $name;$raw=[IO.File]::ReadAllLines($source)
 Write-Extract $source "validation-tails/$($name.Replace('.log','.txt'))" @($raw|Select-Object -Last 28) 'Last 28 original lines; full output retained externally with digest and byte/line counts.' $raw.Count
}
[ordered]@{original_copies=$copies;derived_extracts=$extracts}|ConvertTo-Json -Depth 10|Set-Content (Join-Path $packet 'artifact-bindings.json')
$files=@(Get-ChildItem -LiteralPath $packet -File -Recurse)
$archiveBytes=($files|Measure-Object -Property Length -Sum).Sum
if(@($files|Where-Object Length -gt 2MB).Count-or$archiveBytes-gt5MB){throw 'Curated archive exceeds governed limits'}
[pscustomobject]@{packet=$packet;files=$files.Count;bytes=$archiveBytes;benchmark_attempts=$phases.Count;accepted_attempts=$accepted.Count}
