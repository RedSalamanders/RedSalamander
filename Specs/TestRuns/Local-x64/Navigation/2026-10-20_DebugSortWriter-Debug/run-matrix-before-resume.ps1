$ErrorActionPreference='Stop'
$root=$PSScriptRoot
foreach($profile in @('debug','release')){
 $native=Get-Content (Join-Path $root "functional-final-$profile-01/commands/results.json") -Raw|ConvertFrom-Json
 if($native.passed-ne10-or$native.failed-or$native.skipped){throw 'Both native profiles must pass before benchmark launch'}
}
$blocks=@(
 @{configuration='Debug';role='baseline';block=1},
 @{configuration='Debug';role='candidate';block=1},
 @{configuration='Debug';role='candidate';block=2},
 @{configuration='Debug';role='baseline';block=2},
 @{configuration='Release';role='candidate';block=1},
 @{configuration='Release';role='baseline';block=1},
 @{configuration='Release';role='baseline';block=2},
 @{configuration='Release';role='candidate';block=2}
)
$accepted=[Collections.Generic.List[object]]::new()
foreach($block in $blocks){
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
