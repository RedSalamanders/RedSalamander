$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$repo='Z:/src/RedSalamander'
$accepted=Get-Content (Join-Path $root 'accepted-blocks.json') -Raw|ConvertFrom-Json
$comparison=Get-Content (Join-Path $root 'performance-comparison.json') -Raw|ConvertFrom-Json
$filter='^(folder\.sort_|FolderView\.ApplyCurrentSort(?:\.|$)|FolderView\.ExecuteEnumeration\.(SortMerge|NaturalSortRanks)|render\.(layout_items_us|begin_to_enddraw_us|present_us)$|folder\.layout\.|dwrite\.text_layout\.frame_create)'
foreach($profile in @('Debug','Release')){
 $packet=Join-Path $repo "Specs/TestRuns/Local-x64/Navigation/2026-10-20_DebugSortWriter-$profile"
 if(Test-Path -LiteralPath $packet){
  $existing=@(Get-ChildItem -LiteralPath $packet -File -Recurse)
  if(@($existing|Where-Object Name -ne 'README.md').Count){throw 'Evidence packet already exists; do not overwrite retained data'}
 }else{New-Item -ItemType Directory -Path $packet|Out-Null}
 $bindings=[Collections.Generic.List[object]]::new()
 $receipts=[Collections.Generic.HashSet[string]]::new()
 $copy={param($source,$relative)
  if(!(Test-Path -LiteralPath $source)){throw "Missing source artifact: $source"}
  $destination=Join-Path $packet $relative
  New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force|Out-Null
  Copy-Item -LiteralPath $source -Destination $destination
  $hash=(Get-FileHash -LiteralPath $source).Hash
  if((Get-FileHash -LiteralPath $destination).Hash-ne$hash){throw 'Copied artifact differs'}
  $bindings.Add([ordered]@{source=$source;archive=$relative;sha256=$hash;bytes=(Get-Item -LiteralPath $source).Length;kind='complete-original'})
 }
 $extract={param($source,$relative)
  $lines=[IO.File]::ReadAllLines($source)
  $selected=@(foreach($line in $lines){$row=$line|ConvertFrom-Json;if($row.metric-match$filter){$line}})
  $destination=Join-Path $packet $relative
  New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force|Out-Null
  [IO.File]::WriteAllText($destination,($selected-join"`n")+"`n",[Text.UTF8Encoding]::new($false))
  $bindings.Add([ordered]@{source=$source;archive=$relative;source_sha256=(Get-FileHash -LiteralPath $source).Hash;source_bytes=(Get-Item -LiteralPath $source).Length;source_rows=$lines.Count;selected_sha256=(Get-FileHash -LiteralPath $destination).Hash;selected_rows=$selected.Count;metric_filter=$filter;kind='complete-metric-families-original-order';omitted='Full original includes per-item draw, all other frame and preparation families. It remains intact at source; analysis uses the complete original.'})
 }
 $compress={param($source,$relative)
  $destination=Join-Path $packet $relative
  New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force|Out-Null
  $inputStream=$null;$outputStream=$null;$gzip=$null
  try{
   $inputStream=[IO.File]::OpenRead($source)
   $outputStream=[IO.File]::Open($destination,[IO.FileMode]::CreateNew)
   $gzip=[IO.Compression.GZipStream]::new($outputStream,[IO.Compression.CompressionMode]::Compress,$true)
   $inputStream.CopyTo($gzip)
  }finally{if($gzip){$gzip.Dispose()};if($outputStream){$outputStream.Dispose()};if($inputStream){$inputStream.Dispose()}}
  $bindings.Add([ordered]@{source=$source;archive=$relative;source_sha256=(Get-FileHash -LiteralPath $source).Hash;source_bytes=(Get-Item -LiteralPath $source).Length;compressed_sha256=(Get-FileHash -LiteralPath $destination).Hash;compressed_bytes=(Get-Item -LiteralPath $destination).Length;kind='gzip-complete-original';retention='Every original byte, including all gate and measurement observations, is retained losslessly.'})
 }
 foreach($block in @($accepted.blocks|Where-Object configuration -eq $profile)){
  $phaseRoot=Join-Path $root $block.phase
  foreach($relative in @('observation.json','bindings.json','results.json','run-all-tests-results.json','workspace-final.json','commands/results.json','commands/trace.txt','build-process-samples.jsonl','observed-build-processes.jsonl','perf/folderView_perf_sort_toggle_stress_metrics.json')){
   $source=Join-Path $phaseRoot $relative
   if(Test-Path -LiteralPath $source){
    if($relative-eq'build-process-samples.jsonl'){&$compress $source "$($block.phase)/$relative.gz"}else{&$copy $source "$($block.phase)/$relative"}
   }
  }
  &$extract (Join-Path $phaseRoot 'perf/perf_metrics.jsonl') "$($block.phase)/perf/perf-selected.jsonl"
  $receiptPath=Join-Path $phaseRoot 'build-receipt.json'
  $receipt=Get-Content -LiteralPath $receiptPath -Raw|ConvertFrom-Json
  if($receipts.Add($receipt.receipt_id)){&$copy $receiptPath "receipts/$($receipt.receipt_id).json"}
  @{receipt_id=$receipt.receipt_id;archive="receipts/$($receipt.receipt_id).json"}|ConvertTo-Json|Set-Content (Join-Path $packet "$($block.phase)-receipt-ref.json")
 }
 $native="functional-final-$($profile.ToLowerInvariant())-01"
 foreach($relative in @('bindings.json','results.json','run-all-tests-results.json','workspace-final.json','commands/results.json','commands/trace.txt')){
  &$copy (Join-Path $root "$native/$relative") "$native/$relative"
 }
 foreach($name in @('source-preservation.json','source-preservation-pre-measurement.json','final-identities.json','format-verification.json','orchestration-parse.json','metadata-verification.json','tooling-main-closeout.json','tooling-main-closeout.xml','attempt-ledger.json','observer-native-positive-context.json','observer-positive-context.json','non-build-process-exclusion.json','analyze-results.ps1','verify-source.ps1','run-matrix.ps1','run-matrix-before-resume.ps1','AcceptedPrefix.psm1','AcceptedPrefix.Tests.ps1','resume-guard-tests.json','resume-guard-tests.xml','orchestrator-resumes.jsonl','run-isolated-benchmark.ps1','scan-processes.cs','run-functional.ps1','curate-evidence.ps1','verify-evidence.ps1','render-report.ps1','write-attempt-ledger.ps1')){
  &$copy (Join-Path $root $name) $name
 }
 foreach($relative in @('probe-result.json','build-process-samples.jsonl')){
  &$copy (Join-Path $root "observer-native-positive-01/$relative") "observer-native-positive-01/$relative"
 }
 [ordered]@{profile=$profile;protocol=$comparison.protocol;selection=$comparison.selection;sequence=$comparison.sequence;phases=@($comparison.phases|Where-Object configuration -eq $profile);statistics=@($comparison.statistics|Where-Object configuration -eq $profile);samples=@($comparison.samples|Where-Object configuration -eq $profile);preparation_samples=@($comparison.preparation_samples|Where-Object configuration -eq $profile);preparation_statistics=@($comparison.preparation_statistics|Where-Object configuration -eq $profile);scope=$comparison.scope}|ConvertTo-Json -Depth 12|Set-Content (Join-Path $packet 'performance-comparison.json')
 [ordered]@{profile=$profile;bindings=$bindings.ToArray();retention='All timing/request/counter samples are retained. Entire named metric families are selected without latency selection. Full originals and all failed/rejected attempts remain under the marked external evidence root.'}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $packet 'artifact-bindings.json')
 $files=@(Get-ChildItem -LiteralPath $packet -Recurse -File)
 [pscustomobject]@{profile=$profile;files=$files.Count;bytes=($files|Measure-Object Length -Sum).Sum}|ConvertTo-Json -Compress
}
