param([string]$DebugPhase='quiet-debug-03',[string]$ReleasePhase='quiet-release-02')
$ErrorActionPreference='Stop'
$repo='Z:/src/RedSalamander'
$packet=Join-Path $repo 'Specs/TestRuns/Local-x64/Navigation/2026-10-17_NaturalSortingQuietRerun'
$prior=Get-Content 'Z:/RedSalamander.Perf/evidence/natural-sort-optimization-2026-10-17/performance-comparison.json' -Raw|ConvertFrom-Json
$requests=@('Name Ascending','Extension Ascending','Time Descending','Size Descending','None Ascending (first)','Name Descending','Extension Descending','Time Ascending','Size Ascending','None Ascending (last)')
$expectedDetails=@('name','extension','time','size','none','name','extension','time','size','none')
function Stats($values){$v=@($values|Sort-Object); if(!$v.Count){return $null}; [ordered]@{count=$v.Count;min_us=$v[0];median_us=($v[[int][Math]::Floor(($v.Count-1)/2)]+$v[[int][Math]::Ceiling(($v.Count-1)/2)])/2;max_us=$v[-1]}}
$copies=[Collections.Generic.List[object]]::new()
$extracts=[Collections.Generic.List[object]]::new()
function Copy-Bound([string]$source,[string]$relative){
 $target=Join-Path $packet $relative
 if((Test-Path -LiteralPath $target)-and$relative-notin@('performance-comparison.json','analyze-and-publish.ps1','build-validation.json')-and(Get-FileHash -LiteralPath $source).Hash-ne(Get-FileHash -LiteralPath $target).Hash){throw "Archive destination content differs: $relative"}
 New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force|Out-Null
 Copy-Item -LiteralPath $source -Destination $target
 $hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
 if((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash-ne$hash){throw 'Archive copy hash mismatch'}
 $copies.Add([pscustomobject]@{source=$source;archive=$relative.Replace('\','/');sha256=$hash})
}
$phases=[ordered]@{}
foreach($phase in @('quiet-debug-02',$DebugPhase,'quiet-release-01',$ReleasePhase)){
 $phaseRoot=Join-Path $PSScriptRoot $phase
 $observation=Get-Content (Join-Path $phaseRoot 'observation.json') -Raw|ConvertFrom-Json
 if($phase-in@($DebugPhase,$ReleasePhase)-and!$observation.build_free_eligible){throw "$phase is not eligible build-free evidence"}
 $caseResults=Get-Content (Join-Path $phaseRoot 'commands/results.json') -Raw|ConvertFrom-Json
 if($caseResults.passed-ne5-or$caseResults.failed-ne0-or$caseResults.skipped-ne0-or$caseResults.cases.Count-ne5-or@($caseResults.cases|Where-Object {$_.status-ne'passed'-or$_.name-ne'folderView_perf_sort_toggle_stress'}).Count){throw 'Unexpected case verdicts'}
 $rawPath=Join-Path $phaseRoot 'perf/perf_metrics.jsonl'
 $raw=[IO.File]::ReadAllText($rawPath)
 $all=@($raw -split '\r?\n'|Where-Object {$_}|ConvertFrom-Json)
 $sorts=@($all|Where-Object metric -eq 'folder.sort_toggle_us')
 if($sorts.Count-ne50){throw "Expected 50 sort events in $phase"}
 $samples=for($i=0;$i-lt50;$i++){
  if($sorts[$i].detail-ne$expectedDetails[$i%10]){throw 'Unexpected request sequence'}
  [ordered]@{repeat=1+[int][Math]::Floor($i/10);request=$requests[$i%10];duration_us=$sorts[$i].durationUs;render_calls=$sorts[$i].value1;timestamp=$sorts[$i].timestamp}
 }
 $requestStats=for($i=0;$i-lt10;$i++){$v=for($j=$i;$j-lt50;$j+=10){$sorts[$j].durationUs};[ordered]@{request=$requests[$i];stats=(Stats $v)}}
 $merge=@($all|Where-Object {$_.metric-eq'FolderView.ExecuteEnumeration.SortMerge'-and$_.value0-eq0-and$_.value1-eq5000})
 $ranks=@($all|Where-Object {$_.metric-eq'FolderView.ExecuteEnumeration.NaturalSortRanks'-and$_.value0-eq5000})
 $build=@($all|Where-Object {$_.metric-eq'FolderView.ExecuteEnumeration.BuildItems'-and$_.value0-eq5000-and$_.value1-eq5000})
 $apply=@($all|Where-Object {$_.metric-eq'FolderView.ApplyCurrentSort'-and$_.value0-eq5000})
 $lastCase=Get-Content (Join-Path $phaseRoot 'perf/folderView_perf_sort_toggle_stress_metrics.json') -Raw|ConvertFrom-Json
 if($lastCase.itemCount-ne5000-or$lastCase.enumerationCount-ne1-or$lastCase.sortOperations.Count-ne10-or@($lastCase.sortOperations|Where-Object {$_.itemCount-ne5000-or$_.incrementalSearchEffectUpdates-ne0}).Count){throw 'Stress invariants failed'}
 $phases[$phase]=[ordered]@{configuration=$observation.configuration;run_id=$observation.run_id;raw_sha256=(Get-FileHash $rawPath -Algorithm SHA256).Hash;machine_hashes=@($all.machineHash|Sort-Object -Unique);case_verdicts=[ordered]@{passed=$caseResults.passed;failed=$caseResults.failed;skipped=$caseResults.skipped};observation=$observation;samples=$samples;requests=$requestStats;enumeration_sort_merge=(Stats $merge.durationUs);enumeration_build_items=(Stats $build.durationUs);rank_preparation=(Stats $ranks.durationUs);retained_rank_bytes_per_listing=if($ranks.Count){$ranks[0].value1}else{0};temporary_rank_index_bytes=40000;apply_current_sort=(Stats $apply.durationUs)}
 foreach($file in Get-ChildItem $phaseRoot -File -Recurse){
  $relative=$file.FullName.Substring($phaseRoot.Length+1).Replace('\','/')
  if($relative-eq'perf/perf_metrics.jsonl'){
   $pattern='folder\.sort_toggle_us|FolderView\.ApplyCurrentSort|FolderView\.ExecuteEnumeration\.(BuildItems|SortMerge|NaturalSortRanks)'
   $rows=[regex]::Matches($raw,'(?m)^.*(?:'+$pattern+').*(?:\r?\n|$)')
   $archive=$phase+'/perf/perf-selected.jsonl'
   $target=Join-Path $packet $archive
   New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force|Out-Null
   [IO.File]::WriteAllText($target,(($rows|ForEach-Object {$_.Value})-join''),[Text.UTF8Encoding]::new($false))
   $extracts.Add([ordered]@{source=$file.FullName;source_sha256=(Get-FileHash $file.FullName -Algorithm SHA256).Hash;source_bytes=$file.Length;source_nonblank_rows=$all.Count;archive=$archive;archive_sha256=(Get-FileHash $target -Algorithm SHA256).Hash;selection=$pattern;selected_rows=$rows.Count;all_sort_samples=50})
  }elseif($relative-eq'build-process-samples.jsonl'){
   $lines=[IO.File]::ReadAllLines($file.FullName)
   $quietEnd=([DateTime]$observation.gate_ended_utc).ToUniversalTime()
   $quietStart=$quietEnd.AddSeconds(-30)
   $rows=@($lines|Where-Object { $sample=$_|ConvertFrom-Json; $sample.stage-eq'benchmark'-or(([DateTime]$sample.utc).ToUniversalTime()-ge$quietStart) })
   $archive=$phase+'/build-process-qualifying-samples.jsonl'
   $target=Join-Path $packet $archive
   [IO.File]::WriteAllLines($target,$rows,[Text.UTF8Encoding]::new($false))
   $extracts.Add([ordered]@{source=$file.FullName;source_sha256=(Get-FileHash $file.FullName -Algorithm SHA256).Hash;source_bytes=$file.Length;source_rows=$lines.Count;archive=$archive;archive_sha256=(Get-FileHash $target -Algorithm SHA256).Hash;selection='All benchmark samples plus final 30 seconds of gate by recorded UTC; no benchmark samples omitted';selected_rows=$rows.Count})
  }else{Copy-Bound $file.FullName ($phase+'/'+$relative)}
 }
}
$comparison=for($i=0;$i-lt10;$i++){
 $debug=$phases[$DebugPhase].requests[$i].stats.median_us
 $release=$phases[$ReleasePhase].requests[$i].stats.median_us
 $oldDebug=$prior.phases.'debug-qualified'.requests[$i].stats.median_us
 $oldRelease=$prior.phases.'candidate-release'.requests[$i].stats.median_us
 $baseline=$prior.phases.'baseline-debug'.requests[$i].stats.median_us
 [ordered]@{request=$requests[$i];quiet_debug_median_us=$debug;quiet_release_median_us=$release;historical_contended_debug_baseline_median_us=$baseline;historical_contended_debug_candidate_median_us=$oldDebug;historical_contended_release_candidate_median_us=$oldRelease;debug_change_from_historical_contended_baseline_percent=[Math]::Round(100*($debug-$baseline)/$baseline,2);debug_change_from_historical_contended_candidate_percent=[Math]::Round(100*($debug-$oldDebug)/$oldDebug,2);release_change_from_historical_contended_candidate_percent=[Math]::Round(100*($release-$oldRelease)/$oldRelease,2)}
}
[ordered]@{scope='Same-machine optimized candidate rerun: five repetitions and 50 sort requests per profile. Accepted phases are quiet-debug-03 and quiet-release-02; no compiler or identified build-launcher process observed during their runner windows. Earlier attempts are preserved separately and excluded from the accepted comparison. Historical baseline/candidate rows had concurrent builds, so comparisons are directional and not a new build-free baseline/candidate A/B.';prior_comparison_path='Z:/RedSalamander.Perf/evidence/natural-sort-optimization-2026-10-17/performance-comparison.json';prior_comparison_sha256=(Get-FileHash 'Z:/RedSalamander.Perf/evidence/natural-sort-optimization-2026-10-17/performance-comparison.json').Hash;phases=$phases;comparison=$comparison}|ConvertTo-Json -Depth 18|Set-Content (Join-Path $PSScriptRoot 'performance-comparison.json') -Encoding utf8
foreach($name in @('performance-comparison.json','tested-code.json','runtime-verification.json','preflight.json','non-build-process-exclusion.json','run-quiet-benchmark.ps1','archive-inventory-validation.log','build-Debug.log','build-Release.log','attempt-history.json','build-validation.json','analyze-and-publish.ps1')){Copy-Bound (Join-Path $PSScriptRoot $name) $name}
foreach($name in @('probe-result.json','build-process-samples.jsonl','observed-build-processes.jsonl')){Copy-Bound (Join-Path $PSScriptRoot ('observer-control-03/'+$name)) ('observer-control/'+$name)}
Copy-Bound 'Z:/RedSalamander.Perf/evidence/natural-sort-optimization-2026-10-17/performance-comparison.json' 'historical-contended-comparison.json'
[ordered]@{original_copies=$copies;derived_extracts=$extracts;complete_bulk_metrics_and_monitor_logs_preserved_at=$PSScriptRoot}|ConvertTo-Json -Depth 10|Set-Content (Join-Path $packet 'artifact-bindings.json') -Encoding utf8
$comparison|ForEach-Object {[pscustomobject]$_}|Format-Table request,quiet_debug_median_us,quiet_release_median_us,debug_change_from_historical_contended_baseline_percent -AutoSize
Get-ChildItem $packet -File -Recurse|Measure-Object Length -Sum|Select-Object Count,Sum




