$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'AcceptedPrefix.psm1')
$script:prefixFixtureBase=Join-Path $PSScriptRoot ('resume-guard-fixtures/'+[guid]::NewGuid().ToString('N'))
if(!([IO.Path]::GetFullPath($script:prefixFixtureBase).StartsWith('Z:\RedSalamander.Perf\',[StringComparison]::OrdinalIgnoreCase))){throw 'Fixture must stay in the marked task sandbox'}
function Write-PrefixFixtureJson($path,$value){
 [IO.Directory]::CreateDirectory((Split-Path $path -Parent))|Out-Null
 [IO.File]::WriteAllText($path,($value|ConvertTo-Json -Depth 16),[Text.UTF8Encoding]::new($false))
}
function New-PrefixFixture {
 $fixture=Join-Path $script:prefixFixtureBase ([guid]::NewGuid().ToString('N'))
 $repo=Join-Path $fixture 'repository'
 [IO.Directory]::CreateDirectory($repo)|Out-Null
 foreach($name in @('source.cpp','runtime.bin')){[IO.File]::WriteAllText((Join-Path $repo $name),"owned synthetic fixture $name")}
 $sourceHash=(Get-FileHash (Join-Path $repo 'source.cpp')).Hash
 $runtimeHash=(Get-FileHash (Join-Path $repo 'runtime.bin')).Hash
 $entry=@{configuration='Debug';role='baseline';block=1;phase='baseline-debug-block1-01'}
 Write-PrefixFixtureJson (Join-Path $fixture 'accepted-blocks-progress.json') @($entry)
 $identity=@{role='baseline';configuration='Debug';platform='x64';repository=$repo;receipt_id='receipt';source_snapshot_id='snapshot';toolchain_identity_digest='toolchain';runtime_dependency_digest='runtime';tests_enabled=$true;compiled_inputs=@(@{path='source.cpp';sha256=$sourceHash});runtime_outputs=@(@{path='runtime.bin';sha256=$runtimeHash;bytes=10})}
 Write-PrefixFixtureJson (Join-Path $fixture 'final-identities.json') @{profiles=@($identity)}
 $phase=Join-Path $fixture $entry.phase
 $receipt=@{receipt_id='receipt';configuration='Debug';platform='x64';tests_enabled=$true;source_snapshot_id='snapshot';toolchain_identity_digest='toolchain';runtime_dependency_digest='runtime'}
 Write-PrefixFixtureJson (Join-Path $phase 'build-receipt.json') $receipt
 $snapshot=@{snapshot_id='snapshot';files=@(@{path='source.cpp';worktree_sha256=$sourceHash})}
 $observation=@{phase=$entry.phase;configuration='Debug';build_free_eligible=$true;measurement_contaminated_polls=0;maximum_measurement_gap_ms=20;measurement_sample_count=1;exit_code=0;run_id='run';observer_sha256='observer';scanner_sha256='scanner';runner_started_utc='2000-01-01T00:01:00Z';runner_ended_utc='2000-01-01T00:02:00Z'}
 Write-PrefixFixtureJson (Join-Path $phase 'observation.json') $observation
 $poll=@{stage='benchmark';tools=@();gap_ms=20}
 [IO.File]::WriteAllText((Join-Path $phase 'build-process-samples.jsonl'),($poll|ConvertTo-Json -Compress))
 $native=@{passed=5;failed=0;skipped=0;cases=@(1..5|ForEach-Object {@{name='folderView_perf_sort_toggle_stress';status='passed'}})}
 $values=@{'results.json'=@{status='passed'};'run-all-tests-results.json'=@{status='passed'};'workspace-final.json'=$snapshot;'trace.txt'='fixture';'commands/results.json'=$native;'commands/trace.txt'='fixture';'perf/perf_metrics.jsonl'=@{metric='fixture';durationUs=1};'perf/folderView_perf_sort_toggle_stress_metrics.json'=@{measurementProtocol='product_paint_v2'}}
 $artifacts=@(foreach($relative in $values.Keys){
  $original=Join-Path $fixture "originals/$relative"
  $copy=Join-Path $phase $relative
  Write-PrefixFixtureJson $original $values[$relative]
  [IO.Directory]::CreateDirectory((Split-Path $copy -Parent))|Out-Null
  Copy-Item -LiteralPath $original -Destination $copy
  @{source=$original;archive=$relative;sha256=(Get-FileHash $original).Hash}
 })
 Write-PrefixFixtureJson (Join-Path $phase 'bindings.json') @{phase=$entry.phase;run_id='run';receipt_id='receipt';observer_sha256='observer';scanner_sha256='scanner';artifacts=$artifacts}
 [pscustomobject]@{root=$fixture;phase=$phase;repo=$repo}
}
function Read-PrefixFixture($fixture){
 $script:missingFixtureOriginals=[Collections.Generic.List[object]]::new()
 @(Read-VerifiedSortPrefix -Root $fixture.root -Blocks @(Get-SortBlockPlan) -ObserverSha256 'observer' -ScannerSha256 'scanner' -VerifyCurrentProfile {param($identity,$check) &$check $identity} -MissingOriginals $script:missingFixtureOriginals)
}
function Assert-PrefixRejected($fixture,$message){
 $failure=$null
 try{Read-PrefixFixture $fixture|Out-Null}catch{$failure=$_.Exception.Message}
 $failure|Should Match $message
}
Describe 'Accepted sort prefix is exact and never selected by latency' {
 BeforeEach {$script:fixture=New-PrefixFixture}
 It 'accepts a complete unchanged prefix without launching a native process' {
  @(Read-PrefixFixture $script:fixture).Count|Should Be 1
 }
 It 'accepts an absent checkpoint as an empty prefix' {
  $empty=Join-Path $script:prefixFixtureBase ([guid]::NewGuid().ToString('N'))
  [IO.Directory]::CreateDirectory($empty)|Out-Null
  $missing=[Collections.Generic.List[object]]::new()
  @(Read-VerifiedSortPrefix -Root $empty -Blocks @(Get-SortBlockPlan) -ObserverSha256 'observer' -ScannerSha256 'scanner' -VerifyCurrentProfile {throw 'must not run'} -MissingOriginals $missing).Count|Should Be 0
 }
 It 'rejects a reordered or duplicate prefix instead of reselecting a block' {
  Write-PrefixFixtureJson (Join-Path $script:fixture.root 'accepted-blocks-progress.json') @(@{configuration='Debug';role='candidate';block=1;phase='candidate-debug-block1-01'})
  Assert-PrefixRejected $script:fixture 'exact fixed-order'
 }
 It 'rejects traversal or mismatched phase names' {
  Write-PrefixFixtureJson (Join-Path $script:fixture.root 'accepted-blocks-progress.json') @(@{configuration='Debug';role='baseline';block=1;phase='../baseline-debug-block1-01'})
  Assert-PrefixRejected $script:fixture 'phase path'
 }
 It 'rejects an ineligible completed observer result' {
  $p=Join-Path $script:fixture.phase 'observation.json'; $v=Get-Content $p -Raw|ConvertFrom-Json; $v.build_free_eligible=$false; Write-PrefixFixtureJson $p $v
  Assert-PrefixRejected $script:fixture 'observation is ineligible'
 }
 It 'rejects changed observer or scanner versions' {
  $p=Join-Path $script:fixture.phase 'observation.json'; $v=Get-Content $p -Raw|ConvertFrom-Json; $v.scanner_sha256='changed'; Write-PrefixFixtureJson $p $v
  Assert-PrefixRejected $script:fixture 'observer or scanner changed'
 }
 It 'checks real polling rows rather than trusting the eligibility boolean' {
  [IO.File]::WriteAllText((Join-Path $script:fixture.phase 'build-process-samples.jsonl'),(@{stage='benchmark';tools=@(@{pid=1});gap_ms=20}|ConvertTo-Json -Compress -Depth 5))
  Assert-PrefixRejected $script:fixture 'polling coverage'
 }
 It 'rejects excessive polling gaps' {
  [IO.File]::WriteAllText((Join-Path $script:fixture.phase 'build-process-samples.jsonl'),(@{stage='benchmark';tools=@();gap_ms=1001}|ConvertTo-Json -Compress -Depth 5))
  Assert-PrefixRejected $script:fixture 'polling coverage'
 }
 It 'rejects failed or skipped native cases' {
  $p=Join-Path $script:fixture.phase 'commands/results.json'; $v=Get-Content $p -Raw|ConvertFrom-Json; $v.skipped=1; Write-PrefixFixtureJson $p $v
  Assert-PrefixRejected $script:fixture 'native verdict'
 }
 It 'rejects a different source snapshot or receipt' {
  $p=Join-Path $script:fixture.phase 'build-receipt.json'; $v=Get-Content $p -Raw|ConvertFrom-Json; $v.source_snapshot_id='changed'; Write-PrefixFixtureJson $p $v
  Assert-PrefixRejected $script:fixture 'source, receipt'
 }
 It 'rejects changed current compiled inputs' {
  [IO.File]::AppendAllText((Join-Path $script:fixture.repo 'source.cpp'),'changed')
  Assert-PrefixRejected $script:fixture 'compiled input or runtime'
 }
 It 'rejects changed current executable closure bytes' {
  [IO.File]::AppendAllText((Join-Path $script:fixture.repo 'runtime.bin'),'changed')
  Assert-PrefixRejected $script:fixture 'compiled input or runtime'
 }
 It 'rejects altered raw data or archived artifacts' {
  [IO.File]::AppendAllText((Join-Path $script:fixture.phase 'perf/perf_metrics.jsonl'),'changed')
  Assert-PrefixRejected $script:fixture 'artifact hash changed'
 }
 It 'retains complete hash-verified bytes when an ephemeral canonical source is gone and records that absence' {
  Rename-Item -LiteralPath (Join-Path $script:fixture.root 'originals/perf/perf_metrics.jsonl') -NewName 'perf_metrics.retained.jsonl'
  @(Read-PrefixFixture $script:fixture).Count|Should Be 1
  $script:missingFixtureOriginals.Count|Should Be 1
 }
 It 'rejects changed ephemeral originals even when the complete copy is intact' {
  [IO.File]::AppendAllText((Join-Path $script:fixture.root 'originals/perf/perf_metrics.jsonl'),'changed')
  Assert-PrefixRejected $script:fixture 'original artifact hash changed'
 }
}
