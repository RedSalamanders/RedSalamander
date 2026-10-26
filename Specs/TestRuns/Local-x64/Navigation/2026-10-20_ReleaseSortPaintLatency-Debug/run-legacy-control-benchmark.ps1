param([string]$Configuration,[string]$Phase,[switch]$ProbeOnly,[int]$ExpectedProcessId=0)
$ErrorActionPreference='Stop'
$repo='Z:/src/RedSalamander-worktrees/debug-sort-row-moves-validation'
$observerSourceHash=(Get-FileHash -LiteralPath $PSCommandPath).Hash
$scannerSourceHash=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'scan-processes.cs')).Hash
Add-Type -Path (Join-Path $PSScriptRoot 'scan-processes.cs')
$phaseRoot=Join-Path $PSScriptRoot $Phase
if(Test-Path $phaseRoot){throw 'Attempt destination already exists'}
New-Item -ItemType Directory -Path $phaseRoot|Out-Null
$toolNames=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach($name in @('MSBuild','cl','link','lib','rc','mt','clang','clang++','clang-cl','lld','lld-link','ninja','cmake','vcpkg','zig','gcc','g++','cc1','cc1plus','rustc','cargo','csc','vbc','dotnet','make','gmake','clang-scan-deps','DxUi.ControlTests','DxUi.EmbeddedTests','DxUi.FoundationTests','DxUi.PerformanceBench')){$null=$toolNames.Add($name)}
$observerProcess=[Diagnostics.Process]::GetCurrentProcess()
$observer=[IO.StreamWriter]::new((Join-Path $phaseRoot 'build-process-samples.jsonl'),$false,[Text.UTF8Encoding]::new($false))
$script:lastPoll=[Diagnostics.Stopwatch]::StartNew()
$script:sampleCount=0
$script:maxGap=0.0
$script:observerWorkMs=0.0
$script:observedPids=[Collections.Generic.HashSet[int]]::new()
$script:shellCache=@{}
$script:lastGap=0.0
$script:nonBuild=Get-Content (Join-Path $PSScriptRoot 'non-build-process-exclusion.json') -Raw|ConvertFrom-Json
$script:idleNodes=@(Get-Content (Join-Path $PSScriptRoot 'idle-node-exclusions.json') -Raw|ConvertFrom-Json)
$script:detailCache=@{}
$script:unknownShellFirstSeen=@{}
$script:lastUnknownShellCount=0
$script:measurementUnknownShellPolls=0
$script:idleCandidates=@{}
function Test-OriginalBuildParentExited([int]$parentId,[long]$childStartTicks){
 try{
  $parent=[Diagnostics.Process]::GetProcessById($parentId)
  try{return $parent.StartTime.ToUniversalTime().Ticks-gt$childStartTicks}finally{$parent.Dispose()}
 }catch [ArgumentException]{return $true}catch [InvalidOperationException]{return $true}catch [ComponentModel.Win32Exception]{return $false}
}
function Sample-BuildProcesses([string]$stage){
 $gap=$script:lastPoll.Elapsed.TotalMilliseconds
 $script:lastGap=$gap
 $script:lastPoll.Restart()
 $script:maxGap=[Math]::Max($script:maxGap,$gap)
 $work=[Diagnostics.Stopwatch]::StartNew()
 $interesting=[Collections.Generic.List[object]]::new()
 $unknown=[Collections.Generic.List[int]]::new()
 foreach($process in [SortProcessScanner]::Scan($PID)){
   $name=$process.Name
   $ticks=$process.StartTicks
   if($name-eq$script:nonBuild.name-and$process.Id-eq$script:nonBuild.pid-and$ticks-eq$script:nonBuild.start_ticks){continue}
   $idleNode=@($script:idleNodes|Where-Object {$_.pid-eq$process.Id-and$_.start_ticks-eq$ticks-and$_.cpu_ticks-eq$process.CpuTicks})
   if($idleNode.Count){continue}
   $key='{0}:{1}'-f$process.Id,$ticks
   $interesting.Add([pscustomobject]@{name=$name;pid=$process.Id;key=$key;is_tool=$toolNames.Contains($name);start_ticks=$ticks;cpu_ticks=$process.CpuTicks})
   if(!$script:detailCache.ContainsKey($key)-or[string]::IsNullOrEmpty($script:detailCache[$key].CommandLine)){$unknown.Add($process.Id)}
 }
 if($unknown.Count){
  $filter=(@($unknown|ForEach-Object {"ProcessId=$_"})-join' OR ')
  $details=@{}
  foreach($detail in @(Get-CimInstance Win32_Process -Filter $filter -ErrorAction Stop)){$details[[int]$detail.ProcessId]=$detail}
  foreach($entry in $interesting){
   if(!$script:detailCache.ContainsKey($entry.key)-or[string]::IsNullOrEmpty($script:detailCache[$entry.key].CommandLine)){
    $script:detailCache[$entry.key]=if($details.ContainsKey($entry.pid)){$details[$entry.pid]}else{$null}
   }
  }
 }
 $found=[Collections.Generic.List[object]]::new()
 $unknownShells=[Collections.Generic.List[object]]::new()
 foreach($entry in $interesting){
  $detail=$script:detailCache[$entry.key]
  $unresolved=($null-eq$detail)-or[string]::IsNullOrEmpty($detail.CommandLine)
  if($unresolved-and!$entry.is_tool){
   if(!$script:unknownShellFirstSeen.ContainsKey($entry.key)){$script:unknownShellFirstSeen[$entry.key]=[Diagnostics.Stopwatch]::StartNew()}
   $unknownAge=$script:unknownShellFirstSeen[$entry.key].Elapsed.TotalMilliseconds
   $unknownShells.Add([pscustomobject]@{name=$entry.name;pid=$entry.pid;unresolved_age_ms=[Math]::Round($unknownAge,3)})
  }else{$unknownAge=0.0}
  # Only the gate can admit a newly discovered idle reusable worker. During
  # measurement every new tool, and any resumed CPU on an admitted identity,
  # remains contamination. Parent PID reuse cannot impersonate the old parent.
  if($stage-eq'gate'-and$entry.name-eq'MSBuild'-and!$unresolved-and$detail.CommandLine-match'/nodemode:1 /nodeReuse:true'){
   if(Test-OriginalBuildParentExited $detail.ParentProcessId $entry.start_ticks){
    if(!$script:idleCandidates.ContainsKey($entry.key)-or$script:idleCandidates[$entry.key].cpu_ticks-ne$entry.cpu_ticks){
     $script:idleCandidates[$entry.key]=[pscustomobject]@{cpu_ticks=$entry.cpu_ticks;stable=[Diagnostics.Stopwatch]::StartNew()}
    }
    if($script:idleCandidates[$entry.key].stable.Elapsed.TotalSeconds-ge30){
     $script:idleNodes+= [pscustomobject]@{pid=$entry.pid;start_ticks=$entry.start_ticks;cpu_ticks=$entry.cpu_ticks;parent_pid=$detail.ParentProcessId;command_line=$detail.CommandLine;observed_utc=[DateTime]::UtcNow.ToString('o');stable_seconds=$script:idleCandidates[$entry.key].stable.Elapsed.TotalSeconds;origin='Original parent exited; unchanged CPU verified during this gate'}
     continue
    }
   }else{$script:idleCandidates.Remove($entry.key)}
  }
  $buildShell=($unresolved-and$unknownAge-ge1000)-or[bool]($detail.CommandLine-match'(?i)(?:build|build-dxui|restore-dxui|update-dxui)\.ps1\b|run-build\.py\b|cmake\s+--build|dotnet\s+(?:build|msbuild)|/t:(?:rebuild|build)|[\\/]test\.ps1\b')
  if($entry.is_tool-or$buildShell){
   $found.Add([pscustomobject]@{name=$entry.name;pid=$entry.pid;build_shell=$buildShell})
   if($script:observedPids.Add($entry.pid)-and$detail){
    $detail|Select-Object Name,ProcessId,ParentProcessId,CreationDate,ExecutablePath,CommandLine|ConvertTo-Json -Compress|Add-Content (Join-Path $phaseRoot 'observed-build-processes.jsonl') -Encoding utf8
   }
  }
 }
 $script:lastUnknownShellCount=$unknownShells.Count
 $work.Stop()
 $script:observerWorkMs+=$work.Elapsed.TotalMilliseconds
 $script:sampleCount++
 $observer.WriteLine(([ordered]@{utc=[DateTime]::UtcNow.ToString('o');stage=$stage;gap_ms=[Math]::Round($gap,3);work_ms=[Math]::Round($work.Elapsed.TotalMilliseconds,3);tools=@($found.ToArray());unresolved_shells=@($unknownShells.ToArray())}|ConvertTo-Json -Depth 4 -Compress))
 return @($found.ToArray())
}
try{
 if($ProbeOnly){
  $foundExpected=$false
  for($probeIndex=0;$probeIndex-lt5;$probeIndex++){
   $tools=@(Sample-BuildProcesses 'synthetic-observer-control')
   if(@($tools|Where-Object pid -eq $ExpectedProcessId).Count){$foundExpected=$true}
   Start-Sleep -Milliseconds 200
  }
  [ordered]@{scope='Synthetic observer positive control, not a performance benchmark';expected_pid=$ExpectedProcessId;detected_expected_pid=$foundExpected;sample_count=$script:sampleCount}|ConvertTo-Json|Set-Content (Join-Path $phaseRoot 'probe-result.json') -Encoding utf8
  if(!$foundExpected){throw 'Observer missed the synthetic build-script process'}
  return
 }
 $gateStarted=[DateTime]::UtcNow
 $quietSince=[Diagnostics.Stopwatch]::StartNew()
 while($true){
  $tools=@(Sample-BuildProcesses 'gate')
  if($tools.Count-or$script:lastGap-gt1000){$quietSince.Restart()}
  if($quietSince.Elapsed.TotalSeconds-ge30){break}
  Start-Sleep -Milliseconds 200
 }
 $gateEnded=[DateTime]::UtcNow
 $observer.Flush()
 $script:maxGap=0.0
 $script:lastPoll.Restart()
 $measurementStart=[DateTime]::UtcNow
 $observerCpuBefore=$observerProcess.TotalProcessorTime
 $stdout=Join-Path $phaseRoot 'runner.log'
 $stderr=Join-Path $phaseRoot 'runner-stderr.log'
 $taskPwsh=Join-Path $PSHOME 'pwsh.exe'
 $arguments=@('-NoLogo','-NoProfile','-File',(Join-Path $repo 'Tools/Run-AllTests.ps1'),'-Suite','Commands','-Configuration',$Configuration,'-SkipBuild','-CaseFilter','folderView_perf_sort_toggle_legacy_warm_control','-SelfTestRepeat','5')
 $child=Start-Process -FilePath $taskPwsh -ArgumentList $arguments -WorkingDirectory $repo -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr -PassThru
 $contaminatingPolls=0
 $measurementSamples=0
 try{
  do{
   $tools=@(Sample-BuildProcesses 'benchmark')
   $measurementSamples++
   if($script:lastUnknownShellCount){$script:measurementUnknownShellPolls++}
   if($tools.Count){$contaminatingPolls++}
   Start-Sleep -Milliseconds 200
   $child.Refresh()
  }while(!$child.HasExited)
  $child.WaitForExit()
  $exitCode=$child.ExitCode
 }finally{$child.Dispose()}
 $measurementEnd=[DateTime]::UtcNow
 $observer.Flush()
 $runMatch=[regex]::Matches([IO.File]::ReadAllText($stdout),'Run id:\s+([0-9]{8}T[0-9]{6}Z-[0-9]+-[a-f0-9]{32})')
 $runId=if($runMatch.Count){$runMatch[$runMatch.Count-1].Groups[1].Value}else{''}
 $summary=[ordered]@{phase=$Phase;configuration=$Configuration;observer_sha256=$observerSourceHash;scanner_sha256=$scannerSourceHash;idle_msbuild_nodes=$script:idleNodes;idle_node_policy='Only exact reused-node identities whose original parent exited and CPU stayed unchanged for 30 seconds are excluded. New identities can qualify only during the gate, followed by the separate 30-second quiet interval. During measurement any new tool or CPU change remains contamination. Parent PID reuse is checked by start time.';observer='Typed System.Diagnostics.Process.GetProcesses snapshots, nominal 200 ms, and batched cached command-line classification; same observer for baseline and candidate';event_subscription='Snapshot observer by design; no privileged event subscription requested in this follow-up';gate_started_utc=$gateStarted.ToString('o');gate_ended_utc=$gateEnded.ToString('o');required_build_free_gate_seconds=30;runner_started_utc=$measurementStart.ToString('o');runner_ended_utc=$measurementEnd.ToString('o');measurement_sample_count=$measurementSamples;measurement_contaminated_polls=$contaminatingPolls;measurement_unresolved_shell_polls=$script:measurementUnknownShellPolls;unresolved_shell_policy='Retry classification every poll; unclassified orchestration shell lasting at least 1000 ms is a blocker. Shorter launches are logged as unresolved, not asserted to be build processes. Compiler/tool name detection is unconditional.';maximum_measurement_gap_ms=[Math]::Round($script:maxGap,3);sampling_queries_work_ms=[Math]::Round($script:observerWorkMs,3);measurement_observer_cpu_ms=($observerProcess.TotalProcessorTime-$observerCpuBefore).TotalMilliseconds;all_sample_count=$script:sampleCount;tool_names=@($toolNames);exit_code=$exitCode;run_id=$runId;build_free_eligible=($contaminatingPolls-eq0-and$exitCode-eq0-and$runId-ne''-and$script:maxGap-le1000)}
 $summary|ConvertTo-Json -Depth 6|Set-Content (Join-Path $phaseRoot 'observation.json') -Encoding utf8
 if($exitCode-ne0){Get-Content $stdout -Tail 20;throw "Canonical benchmark exit $exitCode"}
 if(!$runId){throw 'Canonical run id missing'}
 $last=Join-Path "Z:/RedSalamander.Perf/runs/$runId" 'artifacts/selftest/last_run'
 $bindings=@()
 foreach($relative in @('results.json','run-all-tests-results.json','workspace-final.json','trace.txt','commands/results.json','commands/trace.txt','perf/perf_metrics.jsonl','perf/folderView_perf_sort_toggle_legacy_warm_control_metrics.json')){
  $source=Join-Path $last $relative
  if(Test-Path -LiteralPath $source){
   $target=Join-Path $phaseRoot $relative
   New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force|Out-Null
   Copy-Item -LiteralPath $source -Destination $target
   $hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
   if((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash-ne$hash){throw 'Artifact copy mismatch'}
   $bindings+=[pscustomobject]@{source=$source;archive=$relative;sha256=$hash}
  }
 }
 $pointer=Get-Content (Join-Path $repo ".build/x64/$Configuration/build-receipt.json") -Raw|ConvertFrom-Json
 Copy-Item -LiteralPath (Join-Path $repo ".build/x64/$Configuration/receipts/$($pointer.receipt_id).json") -Destination (Join-Path $phaseRoot 'build-receipt.json')
 [ordered]@{phase=$Phase;configuration=$Configuration;observer_sha256=$observerSourceHash;scanner_sha256=$scannerSourceHash;run_id=$runId;receipt_id=$pointer.receipt_id;artifacts=$bindings}|ConvertTo-Json -Depth 7|Set-Content (Join-Path $phaseRoot 'bindings.json') -Encoding utf8
 $summary|ConvertTo-Json -Depth 6
}finally{$observer.Dispose();$observerProcess.Dispose()}










