param([string]$Configuration,[string]$Phase,[string]$Repo='Z:/src/RedSalamander')
$ErrorActionPreference='Stop'
$root=Join-Path $PSScriptRoot $Phase
if(Test-Path -LiteralPath $root){throw 'Attempt already exists'}
New-Item -ItemType Directory -Path $root|Out-Null
$filter='cmd_natural_name_sorting_contracts,folder_view_focus_memory_bounded_lru,folder_view_focus_memory_navigation_roundtrip,folder_view_selection_refresh_filter_contract,folder_view_selection_visibility_restore_contract,folderView_perf_focus_selection_state'
Push-Location $Repo
try{
 & (Join-Path $Repo 'Tools/Run-AllTests.ps1') -Suite Commands -Configuration $Configuration -Platform x64 -SkipBuild -CaseFilter $filter *> (Join-Path $root 'runner.log')
 $code=$LASTEXITCODE
}finally{Pop-Location}
$matches=[regex]::Matches([IO.File]::ReadAllText((Join-Path $root 'runner.log')),'Run id:\s+([0-9]{8}T[0-9]{6}Z-[0-9]+-[a-f0-9]{32})')
if(!$matches.Count){Get-Content (Join-Path $root 'runner.log') -Tail 20;throw 'Run id missing'}
$runId=$matches[$matches.Count-1].Groups[1].Value
$last="Z:/RedSalamander.Perf/runs/$runId/artifacts/selftest/last_run"
$bindings=@()
foreach($relative in @('results.json','run-all-tests-results.json','workspace-final.json','trace.txt','commands/results.json','commands/trace.txt','perf/perf_metrics.jsonl')){
 $source=Join-Path $last $relative
 if(Test-Path -LiteralPath $source){
  $target=Join-Path $root $relative
  New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force|Out-Null
  Copy-Item -LiteralPath $source -Destination $target
  $hash=(Get-FileHash -LiteralPath $source).Hash
  if((Get-FileHash -LiteralPath $target).Hash-ne$hash){throw 'Copy hash differs'}
  $bindings+=[ordered]@{source=$source;archive=$relative;sha256=$hash}
 }
}
$receipt=Get-Content (Join-Path $Repo ".build/x64/$Configuration/build-receipt.json") -Raw|ConvertFrom-Json
Copy-Item -LiteralPath (Join-Path $Repo ".build/x64/$Configuration/receipts/$($receipt.receipt_id).json") -Destination (Join-Path $root 'build-receipt.json')
$cases=Get-Content (Join-Path $root 'commands/results.json') -Raw|ConvertFrom-Json
[ordered]@{phase=$Phase;configuration=$Configuration;repository=$Repo;exit_code=$code;run_id=$runId;receipt_id=$receipt.receipt_id;case_filter=$filter;artifacts=$bindings;passed=$cases.passed;failed=$cases.failed;skipped=$cases.skipped}|ConvertTo-Json -Depth 7|Set-Content (Join-Path $root 'bindings.json')
Get-Content (Join-Path $root 'runner.log') -Tail 12
if($code-ne0-or$cases.passed-ne6-or$cases.failed-or$cases.skipped){throw 'Focused native validation did not pass all six cases'}
