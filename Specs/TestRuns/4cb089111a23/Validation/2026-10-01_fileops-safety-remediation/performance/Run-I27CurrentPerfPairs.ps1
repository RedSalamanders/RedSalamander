param()
$ErrorActionPreference='Stop'
$taskRoot=Join-Path $PSScriptRoot 'paired-current-r21'
if(Test-Path -LiteralPath $taskRoot){throw 'Current-condition pair directory already exists.'}
New-Item -ItemType Directory -Path $taskRoot | Out-Null
foreach($name in @('Run-I27Case.ps1','Copy-I27RunEvidence.ps1','baseline-build-receipt.json','candidate-build-receipt.json','baseline-source-binding.json','isolation-source-binding-final.json','candidate-perf-build-binding.json')){Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $taskRoot $name)}
$taskPwsh='C:/Users/eric/.cache/codex-runtimes/codex-primary-runtime/dependencies/native/powershell/pwsh.exe'
$taskBindings=[Collections.Generic.List[object]]::new()
$taskRuntime=@{}
foreach($variant in @('baseline','candidate')){
 $repo=if($variant -eq 'baseline'){'C:/RsI26Adopt'}else{'C:/RsI27Safety'}
 $build=Join-Path $repo '.build/x64/Release'
 $taskRuntime[$variant]=@(Get-ChildItem -LiteralPath $build -Recurse -File|Where-Object Extension -in @('.exe','.dll')|ForEach-Object {[pscustomobject]@{path=$_.FullName;sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()}})
}
foreach($index in 1..5){
 $order=if($index % 2){@('baseline','candidate')}else{@('candidate','baseline')}
 foreach($variant in $order){
  $repo=if($variant -eq 'baseline'){'C:/RsI26Adopt'}else{'C:/RsI27Safety'}
  foreach($module in $taskRuntime[$variant]){if((Get-FileHash -LiteralPath $module.path -Algorithm SHA256).Hash.ToLowerInvariant() -ne $module.sha256){throw 'Runtime changed before a paired capture.'}}
  $started=[DateTime]::UtcNow.ToString('o')
  & $taskPwsh -NoProfile -File (Join-Path $taskRoot 'Run-I27Case.ps1') -RepoRoot $repo -Configuration Release -Suite FileOps -CaseFilter FileOps_IntegratedPresentationPerf -EvidenceName "$variant-$index"
  if($LASTEXITCODE -ne 0){exit $LASTEXITCODE}
  foreach($module in $taskRuntime[$variant]){if((Get-FileHash -LiteralPath $module.path -Algorithm SHA256).Hash.ToLowerInvariant() -ne $module.sha256){throw 'Runtime changed after a paired capture.'}}
  $receipt=Get-Content -LiteralPath (Join-Path $repo '.build/x64/Release/build-receipt.json') -Raw|ConvertFrom-Json
  $run=Get-Content -LiteralPath (Join-Path $taskRoot "$variant-$index/artifacts/selftest/last_run/run-all-tests-results.json") -Raw|ConvertFrom-Json
  $taskBindings.Add([pscustomobject]@{variant=$variant;capture=$index;started_utc=$started;completed_utc=[DateTime]::UtcNow.ToString('o');run_id=$run.run_id;receipt_id=$receipt.receipt_id;source_snapshot_id=$receipt.source_snapshot_id;modules_checked_before_and_after=$taskRuntime[$variant].Count;runtime=$taskRuntime[$variant]})
  [IO.File]::WriteAllText((Join-Path $taskRoot 'interleaved-capture-bindings.json'),($taskBindings.ToArray()|ConvertTo-Json -Depth 7),[Text.UTF8Encoding]::new($false))
  Write-Output "Completed current-condition pair $index/$variant."
 }
}
& $taskPwsh -NoProfile -File (Join-Path $PSScriptRoot 'Compare-I27Performance.ps1') -EvidenceRoot $taskRoot
exit $LASTEXITCODE