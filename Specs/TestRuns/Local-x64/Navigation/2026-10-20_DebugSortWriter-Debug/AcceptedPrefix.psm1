Set-StrictMode -Version Latest

function Get-SortBlockPlan {
 @(
  @{configuration='Debug';role='baseline';block=1},
  @{configuration='Debug';role='candidate';block=1},
  @{configuration='Debug';role='candidate';block=2},
  @{configuration='Debug';role='baseline';block=2},
  @{configuration='Release';role='candidate';block=1},
  @{configuration='Release';role='baseline';block=1},
  @{configuration='Release';role='baseline';block=2},
  @{configuration='Release';role='candidate';block=2}
 )
}

function Read-VerifiedSortPrefix {
 [CmdletBinding()]
 param([Parameter(Mandatory)][string]$Root,
       [Parameter(Mandatory)][object[]]$Blocks,
       [Parameter(Mandatory)][string]$ObserverSha256,
       [Parameter(Mandatory)][string]$ScannerSha256,
       [Parameter(Mandatory)][scriptblock]$VerifyCurrentProfile,
       [Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$MissingOriginals)
 $ErrorActionPreference='Stop'
 $checkpoint=Join-Path $Root 'accepted-blocks-progress.json'
 if(!(Test-Path -LiteralPath $checkpoint)){return}
 $prior=@(Get-Content -LiteralPath $checkpoint -Raw|ConvertFrom-Json)
 if($prior.Count-gt$Blocks.Count){throw 'Prefix exceeds the fixed block plan'}
 $identities=Get-Content -LiteralPath (Join-Path $Root 'final-identities.json') -Raw|ConvertFrom-Json
 $verified=[Collections.Generic.List[object]]::new()
 $verifiedProfiles=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
 $previousEnd=[DateTime]::MinValue
 $requiredArtifacts=@('results.json','run-all-tests-results.json','workspace-final.json','trace.txt','commands/results.json','commands/trace.txt','perf/perf_metrics.jsonl','perf/folderView_perf_sort_toggle_stress_metrics.json')
 for($index=0;$index-lt$prior.Count;$index++){
  $entry=$prior[$index]
  $expected=$Blocks[$index]
  if($entry.configuration-cne$expected.configuration-or$entry.role-cne$expected.role-or$entry.block-ne$expected.block){throw 'Prefix is not the exact fixed-order block prefix'}
  $phasePattern='^{0}-{1}-block{2}-\d+$'-f[regex]::Escape($expected.role),$expected.configuration.ToLowerInvariant(),$expected.block
  if($entry.phase-cnotmatch$phasePattern){throw 'Prefix phase path does not match its declared block'}
  $phaseRoot=Join-Path $Root $entry.phase
  $observation=Get-Content -LiteralPath (Join-Path $phaseRoot 'observation.json') -Raw|ConvertFrom-Json
  if($observation.phase-cne$entry.phase-or$observation.configuration-cne$entry.configuration-or$observation.build_free_eligible-isnot[bool]-or!$observation.build_free_eligible-or$observation.measurement_contaminated_polls-or$observation.maximum_measurement_gap_ms-gt1000-or$observation.exit_code-ne0-or[string]::IsNullOrWhiteSpace($observation.run_id)){throw 'Prefix observation is ineligible'}
  if($observation.observer_sha256-ne$ObserverSha256-or$observation.scanner_sha256-ne$ScannerSha256){throw 'Prefix observer or scanner changed'}
  $started=([DateTime]$observation.runner_started_utc).ToUniversalTime()
  $ended=([DateTime]$observation.runner_ended_utc).ToUniversalTime()
  if($started-lt$previousEnd-or$ended-le$started){throw 'Prefix measurement chronology changed'}
  $previousEnd=$ended
  $polls=@(Get-Content -LiteralPath (Join-Path $phaseRoot 'build-process-samples.jsonl')|ForEach-Object {$_|ConvertFrom-Json}|Where-Object stage -eq 'benchmark')
  if(!$polls.Count-or$polls.Count-ne$observation.measurement_sample_count-or@($polls|Where-Object {$_.tools.Count-or$_.gap_ms-gt1000}).Count){throw 'Prefix polling coverage or eligibility changed'}
  $native=Get-Content -LiteralPath (Join-Path $phaseRoot 'commands/results.json') -Raw|ConvertFrom-Json
  if($native.passed-ne5-or$native.failed-or$native.skipped-or$native.cases.Count-ne5){throw 'Prefix native verdict is not five passes'}
  $identity=@($identities.profiles|Where-Object {$_.configuration-ceq$entry.configuration-and$_.role-ceq$entry.role})
  if($identity.Count-ne1){throw 'Prefix has no unique expected profile identity'}
  $identity=$identity[0]
  $receipt=Get-Content -LiteralPath (Join-Path $phaseRoot 'build-receipt.json') -Raw|ConvertFrom-Json
  $snapshot=Get-Content -LiteralPath (Join-Path $phaseRoot 'workspace-final.json') -Raw|ConvertFrom-Json
  if($receipt.receipt_id-ne$identity.receipt_id-or$receipt.configuration-cne$entry.configuration-or$receipt.platform-cne'x64'-or!$receipt.tests_enabled-or$receipt.source_snapshot_id-ne$identity.source_snapshot_id-or$snapshot.snapshot_id-ne$receipt.source_snapshot_id-or$receipt.toolchain_identity_digest-ne$identity.toolchain_identity_digest-or$receipt.runtime_dependency_digest-ne$identity.runtime_dependency_digest){throw 'Prefix source, receipt, profile or dependency identity changed'}
  foreach($compiledInput in $identity.compiled_inputs){
   $source=@($snapshot.files|Where-Object path -eq $compiledInput.path)
   if($source.Count-ne1-or$source[0].worktree_sha256-ne$compiledInput.sha256){throw 'Prefix compiled source snapshot changed'}
  }
  $profileKey="$($entry.role):$($entry.configuration)"
  if($verifiedProfiles.Add($profileKey)){
   & $VerifyCurrentProfile $identity {
    param($profileIdentity)
    foreach($file in (@($profileIdentity.compiled_inputs)+@($profileIdentity.runtime_outputs))){
     if((Get-FileHash -LiteralPath (Join-Path $profileIdentity.repository $file.path)).Hash-ne$file.sha256){throw 'Prefix current compiled input or runtime artifact changed'}
    }
   }
  }
  $bindings=Get-Content -LiteralPath (Join-Path $phaseRoot 'bindings.json') -Raw|ConvertFrom-Json
  if($bindings.phase-cne$entry.phase-or$bindings.run_id-ne$observation.run_id-or$bindings.receipt_id-ne$receipt.receipt_id-or$bindings.observer_sha256-ne$ObserverSha256-or$bindings.scanner_sha256-ne$ScannerSha256){throw 'Prefix artifact bindings changed identity'}
  foreach($relative in $requiredArtifacts){
   $binding=@($bindings.artifacts|Where-Object archive -eq $relative)
   if($binding.Count-ne1){throw 'Prefix artifact family is missing or duplicated'}
   if((Get-FileHash -LiteralPath (Join-Path $phaseRoot $relative)).Hash-ne$binding[0].sha256){throw 'Prefix original or archived artifact hash changed'}
   if(Test-Path -LiteralPath $binding[0].source){
    if((Get-FileHash -LiteralPath $binding[0].source).Hash-ne$binding[0].sha256){throw 'Prefix original artifact hash changed'}
   }else{
    $MissingOriginals.Add([ordered]@{phase=$entry.phase;source=$binding[0].source;complete_copy=(Join-Path $phaseRoot $relative);sha256=$binding[0].sha256;scope='Ephemeral canonical source is absent. Required complete copied bytes match its originally recorded digest; no timing data is reconstructed.'})
   }
  }
  $verified.Add($entry)
 }
 $verified.ToArray()
}

Export-ModuleMember -Function Get-SortBlockPlan,Read-VerifiedSortPrefix
