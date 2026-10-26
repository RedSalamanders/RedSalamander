param([switch]$VerifyResumeOnly)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
Import-Module (Join-Path $root 'AcceptedPrefix.psm1')
Import-Module 'Z:/src/RedSalamander/Tools/Modules/Build/ArtifactOperationLock.psm1'
foreach($profile in @('debug','release')){
 $native=Get-Content (Join-Path $root "functional-final-$profile-01/commands/results.json") -Raw|ConvertFrom-Json
 if($native.passed-ne10-or$native.failed-or$native.skipped){throw 'Both native profiles must pass before benchmark launch'}
}
$blocks=@(Get-SortBlockPlan)
$accepted=[Collections.Generic.List[object]]::new()
$verifyCurrentProfile={
 param($identity,$check)
 $lease=Enter-RSArtifactOperationLock -RepoRoot $identity.repository -Operation 'sort-prefix-resume-readback' -Scope @{configuration=$identity.configuration;platform='x64'}
 try{
  $currentReceipt=Get-Content -LiteralPath (Join-Path $identity.repository ".build/x64/$($identity.configuration)/build-receipt.json") -Raw|ConvertFrom-Json
  if($currentReceipt.receipt_id-ne$identity.receipt_id){throw 'Prefix current receipt changed'}
  &$check $identity
 }finally{Exit-RSArtifactOperationLock -Lock $lease}
}
$missingOriginals=[Collections.Generic.List[object]]::new()
$prefix=@(Read-VerifiedSortPrefix -Root $root -Blocks $blocks -ObserverSha256 (Get-FileHash (Join-Path $root 'run-isolated-benchmark.ps1')).Hash -ScannerSha256 (Get-FileHash (Join-Path $root 'scan-processes.cs')).Hash -VerifyCurrentProfile $verifyCurrentProfile -MissingOriginals $missingOriginals)
foreach($entry in $prefix){$accepted.Add($entry)}
[ordered]@{utc=[DateTime]::UtcNow.ToString('o');verified_prefix_count=$accepted.Count;phases=@($prefix.phase);mode=if($VerifyResumeOnly){'verification only; no runner launched'}else{'resume fixed sequence'};module_sha256=(Get-FileHash (Join-Path $root 'AcceptedPrefix.psm1')).Hash;missing_ephemeral_sources=$missingOriginals.ToArray()}|ConvertTo-Json -Depth 6 -Compress|Add-Content (Join-Path $root 'orchestrator-resumes.jsonl')
if($VerifyResumeOnly){[pscustomobject]@{verified_prefix_count=$accepted.Count;runner_launched=$false}|ConvertTo-Json;return}
for($blockIndex=$accepted.Count;$blockIndex-lt$blocks.Count;$blockIndex++){
 $block=$blocks[$blockIndex]
 $eligible=$false
 for($attempt=1;$attempt-le8;$attempt++){
  $phase='{0}-{1}-block{2}-{3:d2}'-f$block.role,$block.configuration.ToLowerInvariant(),$block.block,$attempt
  if(Test-Path -LiteralPath (Join-Path $root $phase)){continue}
  & (Join-Path $root 'run-isolated-benchmark.ps1') -Role $block.role -Configuration $block.configuration -Phase $phase | Out-Host
  $o=Get-Content (Join-Path $root "$phase/observation.json") -Raw|ConvertFrom-Json
  $r=Get-Content (Join-Path $root "$phase/commands/results.json") -Raw|ConvertFrom-Json
  if($o.build_free_eligible-and$r.passed-eq5-and!$r.failed-and!$r.skipped-and$r.cases.Count-eq5){
   $accepted.Add([ordered]@{configuration=$block.configuration;role=$block.role;block=$block.block;phase=$phase})
   $accepted.ToArray()|ConvertTo-Json -Depth 5|Set-Content (Join-Path $root 'accepted-blocks-progress.json')
   $eligible=$true
   break
  }
 }
 if(!$eligible){throw 'No eligible block in eight attempts'}
}
[ordered]@{sequence='Debug ABBA; Release BAAB';repeats_per_block=5;selection='First native/observer-eligible attempt, never latency selection';blocks=$accepted.ToArray()}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $root 'accepted-blocks.json')
