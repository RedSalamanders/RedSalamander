$ErrorActionPreference='Stop'
$root=$PSScriptRoot
foreach($configuration in @('debug','release')){
 $r=Get-Content (Join-Path $root "functional-$configuration-01/commands/results.json") -Raw|ConvertFrom-Json
 if($r.passed-ne9-or$r.failed-or$r.skipped){throw 'Both candidate native profiles must pass before timing'}
}
$blocks=@(
 @{configuration='Debug';role='legacy';block=1},
 @{configuration='Debug';role='product';block=1},
 @{configuration='Debug';role='product';block=2},
 @{configuration='Debug';role='legacy';block=2},
 @{configuration='Release';role='product';block=1},
 @{configuration='Release';role='legacy';block=1},
 @{configuration='Release';role='legacy';block=2},
 @{configuration='Release';role='product';block=2}
)
$accepted=[Collections.Generic.List[object]]::new()
foreach($block in $blocks){
 $eligible=$false
 for($attempt=1;$attempt-le8;$attempt++){
  $phase='candidate-{0}-{1}-block{2}-{3:d2}'-f$block.configuration.ToLowerInvariant(),$block.role,$block.block,$attempt
  if(Test-Path -LiteralPath (Join-Path $root $phase)){continue}
  $script=if($block.role-eq'legacy'){'run-legacy-control-benchmark.ps1'}else{'run-isolated-benchmark.ps1'}
  & (Join-Path $root $script) -Configuration $block.configuration -Phase $phase | Out-Host
  $o=Get-Content (Join-Path $root "$phase/observation.json") -Raw|ConvertFrom-Json
  $r=Get-Content (Join-Path $root "$phase/commands/results.json") -Raw|ConvertFrom-Json
  if($o.build_free_eligible-and$r.passed-eq5-and!$r.failed-and!$r.skipped-and$r.cases.Count-eq5){
   $accepted.Add([ordered]@{configuration=$block.configuration;role=$block.role;block=$block.block;phase=$phase})
   $eligible=$true
   break
  }
 }
 if(!$eligible){throw 'No eligible candidate block in eight attempts'}
}
[ordered]@{sequence='Debug legacy/product/product/legacy; Release product/legacy/legacy/product';repeats_per_block=5;selection='First native/observer-eligible attempt, never latency selection';blocks=$accepted.ToArray()}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $root 'accepted-candidates.json')
