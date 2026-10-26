$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$descriptions=[ordered]@{
 'initial-build-debug-unattested.txt'='Compilation completed, but source changed during build; no success receipt. Excluded from qualification.'
 'baseline-build-debug.txt'='Duplicate preserved first unattested Debug build log; not a second successful qualification.'
 'baseline-build-release.txt'='Refused while a same-repository shared-dependency operation was live; no timed result.'
 'baseline-instrumented-build-debug.txt'='New test array initialization produced C5246; repaired before final builds.'
 'baseline-instrumented-build-release.txt'='Disk-full PDB/compilation failure; no success receipt.'
 'candidate-build-debug.txt'='Disk-full dependency restore/clone failure; no success receipt.'
 'baseline-relocated-build-debug.txt'='Stale absolute generated DxUi props after owned worktree relocation; canonical restore regenerated the input.'
 'candidate-relocated-build-debug.txt'='Superseded candidate prototype build, zero warnings/errors; not the final nonblocking-observer binary.'
 'candidate-relocated-build-release.txt'='Superseded candidate prototype build, zero warnings/errors; not the final nonblocking-observer binary.'
 'baseline-final-build-debug.txt'='Final baseline Debug, zero warnings/errors; receipt-bound instrumentation only.'
 'baseline-final-build-release.txt'='Final baseline Release, zero warnings/errors; receipt-bound instrumentation only.'
 'candidate-final-build-debug.txt'='Final candidate Debug, zero warnings/errors; final lock-free test observation.'
 'candidate-final-build-release.txt'='Final candidate Release, zero warnings/errors; final lock-free test observation.'
}
function Bind-File([string]$path){
 [ordered]@{path=$path;sha256=(Get-FileHash -LiteralPath $path).Hash;bytes=(Get-Item -LiteralPath $path).Length}
}
$builds=@(foreach($name in $descriptions.Keys){[ordered]@{artifact=Bind-File (Join-Path $root $name);outcome=$descriptions[$name]}})
$blocks=@(Get-ChildItem -LiteralPath $root -Directory|Where-Object Name -match '^(baseline|candidate)-(debug|release)-block\d+-\d+$'|Sort-Object Name|ForEach-Object{
 $phase=$_.Name
 $files=@(Get-ChildItem -LiteralPath $_.FullName -Recurse -File|ForEach-Object {Bind-File $_.FullName})
 $observationPath=Join-Path $_.FullName 'observation.json'
 $interruptionPath=Join-Path $_.FullName 'interrupted-gate.json'
 $observation=if(Test-Path -LiteralPath $observationPath){Get-Content -LiteralPath $observationPath -Raw|ConvertFrom-Json}else{$null}
 [ordered]@{phase=$phase;build_free_eligible=($null-ne$observation-and$observation.build_free_eligible);status=if($observation){if($observation.build_free_eligible){'observer-eligible; selection recorded separately'}else{'rejected; native result cannot override observer failure'}}elseif(Test-Path -LiteralPath $interruptionPath){'pre-run gate deliberately interrupted; zero benchmark samples'}else{'incomplete; never eligible'};observation=$observation;files=$files}
})
$context=@(foreach($name in @('observer-positive-context.json','observer-native-positive-context.json','non-build-process-exclusion.json','relocated-stale-DxUi.resolved.x64.props','baseline-PerfJsonl.cpp','prototype-PerfJsonl.cpp','candidate-final-PerfJsonl.cpp','run-matrix-before-resume.ps1','resume-guard-pester-syntax-failure.json','resume-guard-pester-syntax-failure.xml','resume-ephemeral-source-refusal.json')){Bind-File (Join-Path $root $name)})
[ordered]@{
 status='All preparation/rejected attempts retained; no latency selection'
 root=$root
 baseline_commit='a6ced5f41ebbb8d08e8fc0f562866bf3b5e6c698'
 builds=$builds
 final_identities=Get-Content (Join-Path $root 'final-identities.json') -Raw|ConvertFrom-Json
 superseded_prototype_receipts=@('c0584346c5b35e7c2dcbc15b035c671d53552274088dc71a1c595e2ddd20446e','d7cd523c6f3d26ab43e7eade9017de71a6197c32fab5f19a76c9489d337783fd')
 prototype_native_recovery='First Debug prototype native result passed ten cases at S:; the wrapper initially assumed former Z: artifacts. Exact original run artifacts were recovered by run ID into functional-debug-01. Final functional-final-* runs supersede it and use the corrected explicit root.'
 prototype_artifacts=@(Get-ChildItem -LiteralPath (Join-Path $root 'functional-debug-01') -Recurse -File|ForEach-Object {Bind-File $_.FullName})
 observer_controls='The real compiler positive control predates the added RedPrism self-test classification; its exact observer version remains in its context. The real RedPrism self-test positive control uses the final scanner. Neither control launches compilation/native work and neither is a benchmark.'
 conservative_false_positive=[ordered]@{phase='baseline-release-block2-01';pid=9464;classification='Read-only tooling inspection referencing test.ps1 in its command text';disposition='Original flagged observation retained and attempt excluded under the fixed classifier; no eligibility override';evidence=(Join-Path $root 'baseline-release-block2-01/observed-build-processes.jsonl')}
 contexts=$context
 orchestration_tests=Get-Content (Join-Path $root 'resume-guard-tests.json') -Raw|ConvertFrom-Json
 orchestration_resumptions=@(Get-Content (Join-Path $root 'orchestrator-resumes.jsonl')|ForEach-Object {$_|ConvertFrom-Json})
 resumption_scope='An exact accepted prefix is preserved without latency reselection. Its source/receipt/profile/dependency identities, current compiled inputs/runtime closure under canonical profile locks, every complete copied artifact and observation/poll rows are verified. Absent ephemeral canonical source paths are explicitly recorded; the required complete captured copies must still match their originally recorded hashes. No native code, timer, corpus or observer is changed by resumption.'
 phases=$blocks
 selected_blocks=Get-Content (Join-Path $root 'accepted-blocks.json') -Raw|ConvertFrom-Json
 scope='An interrupted gate has incomplete polling-buffer flush and no timing authority. Successful native cases from a rejected measurement never establish build-free timing. All originals, logs, source prototypes and unaccepted observations remain external with content hashes.'
}|ConvertTo-Json -Depth 16|Set-Content -LiteralPath (Join-Path $root 'attempt-ledger.json')
