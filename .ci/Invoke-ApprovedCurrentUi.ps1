<# PRIVATE MANUAL DIAGNOSTIC ONLY. Two fixed filtered cases; never Full or production qualification. #>
[CmdletBinding()]
param(
 [Parameter(Mandatory)][string]$CandidateRoot,
 [Parameter(Mandatory)][string]$BootstrapRoot,
 [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{40}$')][string]$ExpectedHead,
 [Parameter(Mandatory)][ValidateSet('PR200','PR201')][string]$Label,
 [switch]$ApproveAccountClose,
 [switch]$ApproveSearchClose
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$PSNativeCommandUseErrorActionPreference=$false
$expected=@{PR200='ab5f142c3ca54d4c2c559a2654f8a2e8df4229eb';PR201='efd597e40e26d17bfcd9dbb32fa8438ff757873b'}
if(-not $ApproveAccountClose -or -not $ApproveSearchClose -or $ExpectedHead -cne $expected[$Label] -or $env:GITHUB_ACTIONS -cne 'true' -or
 $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander' -or
 $env:GITHUB_EVENT_NAME -cne 'workflow_dispatch'){throw 'Explicit manual hosted candidate and account-close authorization required.'}
if([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -cne 'Arm64' -or
 [Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString() -cne 'Arm64'){throw 'Native ARM64 OS and PowerShell process are required.'}
$CandidateRoot=[IO.Path]::GetFullPath($CandidateRoot);$BootstrapRoot=[IO.Path]::GetFullPath($BootstrapRoot)
function Assert-CleanHead([string]$Root,[string]$Head){
 $actual=& git -C $Root rev-parse HEAD;if($LASTEXITCODE -ne 0 -or $actual -cne $Head){throw 'Checkout HEAD mismatch.'}
 $dirty=@(& git -C $Root status --porcelain);if($LASTEXITCODE -ne 0 -or $dirty.Count){throw 'Checkout is dirty or status failed.'}
}
Assert-CleanHead $CandidateRoot $ExpectedHead
Assert-CleanHead $BootstrapRoot $env:GITHUB_SHA
Import-Module (Join-Path $CandidateRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
$testRoot='C:\RedSalamander.Perf';$env:REDSALAMANDER_TEST_ROOT=$testRoot
[void](Initialize-RSTestSandboxRoot -TestRoot $testRoot -RepoRoot $CandidateRoot -AllowExternalInitialization)
$scratch=New-RSTestSandboxScratchDirectory -RepoRoot $CandidateRoot -Harness UnifiedCI -Case "approved-ui-$Label"
$env:TEMP=$scratch;$env:TMP=$scratch
$e=Join-Path $testRoot "evidence/approved-current-ui/$($env:GITHUB_RUN_ID)-$($env:GITHUB_RUN_ATTEMPT)-$Label"
if(Test-Path -LiteralPath $e){throw 'Diagnostic evidence already exists.'};[void][IO.Directory]::CreateDirectory($e)
$record=[ordered]@{schema='private.approved-current-ui.v1';label=$Label;head=$ExpectedHead;bootstrap_head=$env:GITHUB_SHA;
 status='INCOMPLETE';repository_coverage='NOT_EVALUATED';receipt=$null;source_snapshot=$null;cases=@();assertion_probe_exit=$null;
 foreground_snapshot=$null;foreground_snapshot_error=$null;error=$null}
function Child([string]$Repo,[string]$File,[string[]]$Arguments,[string]$Log){
 if(Test-Path -LiteralPath $Log){throw 'Child log already exists.'}
 $old=[IO.Directory]::GetCurrentDirectory();Push-Location -LiteralPath $Repo
 try{[IO.Directory]::SetCurrentDirectory($Repo);& (Join-Path $PSHOME 'pwsh.exe') -NoProfile -NonInteractive -File $File @Arguments *> $Log;return $LASTEXITCODE}
 finally{[IO.Directory]::SetCurrentDirectory($old);Pop-Location}
}
try {
 Import-Module (Join-Path $CandidateRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -Force
 $vcpkg=Initialize-RSVcpkgToolContext -RepoRoot $CandidateRoot
 $code=Child $CandidateRoot (Join-Path $CandidateRoot 'vcpkg-install.ps1') @('-Platform','ARM64','-VcpkgExe',$vcpkg.Executable,'-VcpkgRoot',$vcpkg.RegistryRoot) (Join-Path $e 'vcpkg.log')
 if($code -ne 0){throw "Dependency setup failed: $code"}
 $toolRoot=Resolve-RSVcpkgSafeChildPath -Root $CandidateRoot -Child '.build/vcpkg-tool'
 $environment=Get-RSVcpkgBuildEnvironment -RepoRoot $CandidateRoot -ToolRoot $toolRoot -Platform ARM64 -Configuration Release
 foreach($entry in $environment.GetEnumerator()){[Environment]::SetEnvironmentVariable($entry.Key,[string]$entry.Value,'Process')}
 $env:VCPKG_REGISTRY_ROOT=Resolve-RSVcpkgSafeChildPath -Root $CandidateRoot -Child '.build/vcpkg-registry'
 $env:RSBuildEnableTests='true';$env:RSBuildCompilerDebugInformation='Embedded'
 $buildLog=Join-Path $e 'build.log'
 $code=Child $CandidateRoot (Join-Path $CandidateRoot 'build.ps1') @('-Platform','ARM64','-Configuration','Release','-BuildNumber','1') $buildLog
 if($code -ne 0){throw "Full solution build failed: $code"}
 $buildText=[IO.File]::ReadAllText($buildLog)
 $diagnostics=[regex]::Matches($buildText,'(?m)^\s*Diagnostics:\s*(\d+) warning\(s\),\s*(\d+) error\(s\)\s*$')
 if($diagnostics.Count -lt 1 -or @($diagnostics|Where-Object {$_.Groups[1].Value -cne '0' -or $_.Groups[2].Value -cne '0'}).Count){throw 'Build diagnostic summary is missing or nonzero.'}
 Import-Module (Join-Path $CandidateRoot 'Tools/Modules/Build/BuildEvidence.psm1') -Force
 $output=Join-Path $CandidateRoot '.build/ARM64/Release';$schema=Join-Path $CandidateRoot 'Specs/Build/BuildReceipt.schema.json'
 $receipt=Read-RSBuildReceipt -BuildOutputDir $output -SchemaPath $schema;$snapshot=Get-RSBuildSourceSnapshot -RepoRoot $CandidateRoot
 $required=@('.build/ARM64/Release/RedSalamander.exe','.build/ARM64/Release/PluginContractTests.exe')
 if(-not $receipt -or $receipt.git_head -cne $ExpectedHead -or $receipt.target -cne 'solution' -or $receipt.platform -cne 'ARM64' -or
 $receipt.configuration -cne 'Release' -or -not $receipt.tests_enabled -or 'build-number=1' -cnotin $receipt.build_arguments -or
 'compiler-debug-information=embedded' -cnotin $receipt.build_arguments -or
 -not (Test-RSBuildReceiptForArtifactUse -Receipt $receipt -SourceSnapshot $snapshot -Platform ARM64 -Configuration Release -RepoRoot $CandidateRoot -RequiredRelativePath $required)){throw 'Candidate source/receipt/artifact identity mismatch.'}
 Assert-CleanHead $CandidateRoot $ExpectedHead
 $receiptPath=Join-Path $output 'build-receipt.json';$receiptHash=(Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash
 Copy-Item -LiteralPath $receiptPath -Destination (Join-Path $e 'build-receipt.json')
 $record.receipt=$receipt.receipt_id;$record.source_snapshot=$snapshot.SnapshotId
 # This child runs only the existing guarded close branch and returns before all unrelated API probes.
 $closeLog=Join-Path $e 'approved-account-close.log'
 $code=Child $CandidateRoot (Join-Path $BootstrapRoot '.ci/WgcAdmissionProbe.ps1') @('-TestRoot',$testRoot,'-RepoRoot',$CandidateRoot,'-CloseOnly','-DismissHostedAccountSurface') $closeLog
 $closeText=[IO.File]::ReadAllText($closeLog)
 $intents=[regex]::Matches($closeText,'(?m)^Close intent: (.+?)\s*$');$outcomes=[regex]::Matches($closeText,'(?m)^Close result: (.+?)\s*$')
 if($intents.Count -ne 1 -or $outcomes.Count -ne 1){throw "Approved close lacks one intent and result (exit $code)."}
 $intentPath=Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path $intents[0].Groups[1].Value -RequireDescendant
 $outcomePath=Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path $outcomes[0].Groups[1].Value -RequireDescendant
 Copy-Item -LiteralPath $intentPath -Destination (Join-Path $e 'account-dismissal-intent.json')
 Copy-Item -LiteralPath $outcomePath -Destination (Join-Path $e 'account-dismissal-result.json')
 if($code -ne 0){throw "Approved account-window check failed: $code"}
 $intent=Get-Content -LiteralPath $intentPath -Raw|ConvertFrom-Json -AsHashtable;$outcome=Get-Content -LiteralPath $outcomePath -Raw|ConvertFrom-Json -AsHashtable
 if($intent.Head -cne $ExpectedHead -or $intent.DesktopContext.WindowStation -cne 'WinSta0' -or $intent.DesktopContext.InputDesktop -cne 'Default' -or
    $intent.DesktopContext.ThreadDesktop -cne 'Default' -or $outcome.IntentSha256 -ine (Get-FileHash -LiteralPath $intentPath -Algorithm SHA256).Hash -or
    $outcome.Dismissed -isnot [bool] -or $outcome.NoTarget -isnot [bool] -or $outcome.Dismissed -eq $outcome.NoTarget -or
    $outcome.Error -or $outcome.SnapshotError -or $null -eq $outcome.PostCloseForeground){throw 'Approved close intent/result binding failed.'}
 $postAccount=$outcome.PostCloseForeground
 if($postAccount.Stable -isnot [bool] -or -not $postAccount.Stable){throw 'Post-account foreground is unstable; stop before UI.'}
 if(($postAccount.Window -ne 0 -or $postAccount.Process -ne 0) -and
    ($postAccount.VerifiedShellForeground -isnot [bool] -or -not $postAccount.VerifiedShellForeground)){
  $searchImage=[IO.Path]::Combine([Environment]::GetFolderPath([Environment+SpecialFolder]::Windows),'SystemApps','MicrosoftWindows.Client.CBS_cw5n1h2txyewy','SearchHost.exe')
  if($postAccount.Window -eq 0 -or $postAccount.Process -eq 0 -or $postAccount.Title -cne 'Search' -or
     $postAccount.Class -cne 'Windows.UI.Core.CoreWindow' -or $postAccount.ApplicationId -cne 'MicrosoftWindows.Client.CBS_cw5n1h2txyewy!CortanaUI' -or
     $postAccount.Image -ine $searchImage -or $null -eq $postAccount.Session -or $postAccount.QueryError){throw 'Unexpected post-account foreground; preserve metadata and stop before UI.'}
 }
 # A second separately authorized exact-identity guard owns any Search close.
 $searchLog=Join-Path $e 'approved-search-close.log'
 $code=Child $CandidateRoot (Join-Path $BootstrapRoot '.ci/WgcAdmissionProbe.ps1') @('-TestRoot',$testRoot,'-RepoRoot',$CandidateRoot,'-CloseSearchOnly','-DismissHostedSearchSurface') $searchLog
 $searchText=[IO.File]::ReadAllText($searchLog)
 $searchIntents=[regex]::Matches($searchText,'(?m)^Search intent: (.+?)\s*$');$searchOutcomes=[regex]::Matches($searchText,'(?m)^Search result: (.+?)\s*$')
 if($searchIntents.Count -ne 1 -or $searchOutcomes.Count -ne 1){throw "Approved Search close lacks one intent and result (exit $code)."}
 $searchIntentPath=Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path $searchIntents[0].Groups[1].Value -RequireDescendant
 $searchOutcomePath=Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path $searchOutcomes[0].Groups[1].Value -RequireDescendant
 Copy-Item -LiteralPath $searchIntentPath -Destination (Join-Path $e 'search-dismissal-intent.json')
 Copy-Item -LiteralPath $searchOutcomePath -Destination (Join-Path $e 'search-dismissal-result.json')
 if($code -ne 0){throw "Approved Search close failed: $code"}
 $searchIntent=Get-Content -LiteralPath $searchIntentPath -Raw|ConvertFrom-Json -AsHashtable;$searchOutcome=Get-Content -LiteralPath $searchOutcomePath -Raw|ConvertFrom-Json -AsHashtable
 if($searchIntent.Head -cne $ExpectedHead -or $searchIntent.DesktopContext.WindowStation -cne 'WinSta0' -or $searchIntent.DesktopContext.InputDesktop -cne 'Default' -or
    $searchIntent.DesktopContext.ThreadDesktop -cne 'Default' -or $searchOutcome.IntentSha256 -ine (Get-FileHash -LiteralPath $searchIntentPath -Algorithm SHA256).Hash -or
    $searchOutcome.Dismissed -isnot [bool] -or $searchOutcome.NoTarget -isnot [bool] -or $searchOutcome.Dismissed -eq $searchOutcome.NoTarget -or
    $searchOutcome.Error -or $searchOutcome.SnapshotError -or $null -eq $searchOutcome.PostCloseForeground){throw 'Approved Search intent/result binding failed.'}
 $postSearch=$searchOutcome.PostCloseForeground
 if($postSearch.Stable -isnot [bool] -or -not $postSearch.Stable -or
    (($postSearch.Window -ne 0 -or $postSearch.Process -ne 0) -and
     ($postSearch.VerifiedShellForeground -isnot [bool] -or -not $postSearch.VerifiedShellForeground))){throw 'Post-Search foreground is unstable or is not the verified shell; stop before UI.'}
 Import-Module (Join-Path $CandidateRoot 'Tools/Modules/Testing/HostedDiagnosticEvidence.psm1') -Force
 Import-Module (Join-Path $CandidateRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1') -Force
 $caseFailures=[Collections.Generic.List[string]]::new()
 $prior=[Environment]::GetEnvironmentVariable('REDSALAMANDER_HOSTED_TEST_DESKTOP','Process')
 try {
  [Environment]::SetEnvironmentVariable('REDSALAMANDER_HOSTED_TEST_DESKTOP',(Get-RSHostedInputDesktopRequest -PairedComparison:$false),'Process')
  Invoke-RSWithHostedTestDesktop -Action {
   foreach($scope in @('CommandsHandoff','ViewerSpaceCapture')){
    $observed=$null
    try {
    $selection=Get-RSHostedDiagnosticSelection -Scope $scope
    Assert-CleanHead $CandidateRoot $ExpectedHead
    $before=Get-RSBuildSourceSnapshot -RepoRoot $CandidateRoot
    $current=Read-RSBuildReceipt -BuildOutputDir $output -SchemaPath $schema
    if($before.SnapshotId -cne $snapshot.SnapshotId -or $current.receipt_id -cne $receipt.receipt_id -or
       (Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash -cne $receiptHash -or
       -not (Test-RSBuildReceiptForArtifactUse -Receipt $current -SourceSnapshot $before -Platform ARM64 -Configuration Release -RepoRoot $CandidateRoot -RequiredRelativePath $required)){throw 'Pre-case source or receipt changed.'}
    # Reserve the canonical identity in evidence before launch. New-RSTestRunContext
    # computes paths only; Run-AllTests remains the owner that creates the run root.
    $planned=New-RSTestRunContext -RepoRoot $CandidateRoot -TestRootOverride $testRoot -SelfTestRootOverride ''
    $expectedRunId=[string]$planned.RunId
    if($expectedRunId -cnotmatch '^[0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32}$' -or
       (Test-Path -LiteralPath $planned.RunRoot) -or (Test-Path -LiteralPath (Join-Path $testRoot "evidence/completed-runs/$expectedRunId")) -or
       (Test-Path -LiteralPath (Join-Path $testRoot "evidence/run-archive-records/$expectedRunId")) -or
       (Test-Path -LiteralPath (Join-Path $testRoot "evidence/runs/$expectedRunId"))){throw 'Planned run ID is invalid or already exists.'}
    $plannedPath=Join-Path $e "$scope-planned-run.json"
    $plannedBytes=[Text.UTF8Encoding]::new($false).GetBytes((@{scope=$scope;expected_run_id=$expectedRunId;run_root=$planned.RunRoot;head=$ExpectedHead;receipt=$receipt.receipt_id}|ConvertTo-Json))
    $plannedStream=[IO.File]::Open($plannedPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
    try{$plannedStream.Write($plannedBytes,0,$plannedBytes.Length);$plannedStream.Flush($true)}finally{$plannedStream.Dispose()}
    $log=Join-Path $e "$scope-runner.log";$started=[DateTimeOffset]::UtcNow;$launchError=$null;$code=-1
    try{$code=Child $CandidateRoot (Join-Path $CandidateRoot 'Tools/Run-AllTests.ps1') @('-Platform','ARM64','-Configuration','Release','-Suite',$selection.suite,'-CaseFilter',$selection.case,'-ValidationMode','Fresh','-SkipBuild','-BuildNumber','1','-TestRoot',$testRoot,'-RunId',$expectedRunId,'-TimeoutMultiplier','2.0','-AllowAlternateVolumeTestRoot') $log}
    catch{$launchError=$_.Exception.Message}
    $ended=[DateTimeOffset]::UtcNow;$runnerText=if(Test-Path -LiteralPath $log -PathType Leaf){[IO.File]::ReadAllText($log)}else{''}
    $found=[regex]::Matches($runnerText,'(?m)^\s*Run id:\s*([0-9]{8}T[0-9]{6}Z-[0-9]+-[0-9a-f]{32})\s*$')
    if($null -ne $launchError -or $found.Count -ne 1 -or $found[0].Groups[1].Value -cne $expectedRunId){
     $possible=@(@($expectedRunId)+@($found|ForEach-Object {$_.Groups[1].Value})|Sort-Object -Unique)
     $partial=Join-Path $e "$scope-partial";[void][IO.Directory]::CreateDirectory($partial)
     @{runner_exit=$code;launch_error=$launchError;expected_run_id=$expectedRunId;run_id_occurrences=$found.Count;run_ids=$possible;status='AMBIGUOUS_MISSING_OR_MISMATCHED_RUN_ID'} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $partial 'runner-state.json') -Encoding utf8
     foreach($candidateId in $possible){
      $partialCase=Join-Path $partial $candidateId;[void][IO.Directory]::CreateDirectory($partialCase)
      $partialRun=Join-Path $testRoot "evidence/completed-runs/$candidateId/artifacts/selftest/last_run"
      if(-not(Test-Path -LiteralPath $partialRun -PathType Container)){$partialRun=Join-Path $testRoot "runs/$candidateId/artifacts/selftest/last_run"}
      foreach($rel in @("$($selection.native)/results.json","$($selection.native)/trace.txt",'run-all-tests-results.json')){$source=Join-Path $partialRun $rel;if(Test-Path -LiteralPath $source -PathType Leaf){Copy-Item -LiteralPath $source -Destination (Join-Path $partialCase ([IO.Path]::GetFileName($rel)))}}
      foreach($name in @('intent.json','complete.json')){$source=Join-Path $testRoot "evidence/run-archive-records/$candidateId/$name";if(Test-Path -LiteralPath $source -PathType Leaf){Copy-Item -LiteralPath $source -Destination (Join-Path $partialCase $name)}}
      foreach($origin in @('runs','evidence/completed-runs')){
       $runBase=Join-Path $testRoot "$origin/$candidateId";$prefix=if($origin -ceq 'runs'){'live'}else{'archived'}
       $state=Join-Path $runBase 'run-state.json';if(Test-Path -LiteralPath $state -PathType Leaf){Copy-Item -LiteralPath $state -Destination (Join-Path $partialCase "$prefix-run-state.json")}
       $exits=Join-Path $runBase 'artifacts/native-process-exits'
       if(Test-Path -LiteralPath $exits -PathType Container){$exitDestination=Join-Path $partialCase "$prefix-native-process-exits";[void][IO.Directory]::CreateDirectory($exitDestination);Get-ChildItem -LiteralPath $exits -File -Filter '*.json' | ForEach-Object {Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $exitDestination $_.Name)}}
      }
      $validationState=Join-Path $testRoot "evidence/runs/$candidateId/run-state.json";if(Test-Path -LiteralPath $validationState -PathType Leaf){Copy-Item -LiteralPath $validationState -Destination (Join-Path $partialCase 'validation-run-state.json')}
      $partialCaptures=Join-Path $partialRun 'perf/viewer_space_captures';if($scope -ceq 'ViewerSpaceCapture' -and (Test-Path -LiteralPath $partialCaptures -PathType Container)){Copy-Item -LiteralPath $partialCaptures -Destination (Join-Path $partialCase 'viewer_space_captures') -Recurse}
     }
     throw 'Runner launch failed or did not report its exact planned run ID; available identified evidence retained.'
    }
    $id=Get-RSHostedDiagnosticRunId -Output $runnerText
    $observed=[ordered]@{scope=$scope;run_id=$id;runner_exit=$code;log=$log;status='INCOMPLETE';native_sha256=$null;summary_sha256=$null;intent_sha256=$null}
    $record.cases+=,$observed
    [IO.File]::WriteAllText((Join-Path $e "$scope-run-id.json"),($observed|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
    # Retain available raw evidence before validation, including incomplete/failed runs.
    $run=Join-Path $testRoot "evidence/completed-runs/$id/artifacts/selftest/last_run"
    if(-not (Test-Path -LiteralPath $run -PathType Container)){$run=Join-Path $testRoot "runs/$id/artifacts/selftest/last_run"}
    $caseDir=Join-Path $e $scope;[void][IO.Directory]::CreateDirectory($caseDir)
    foreach($rel in @("$($selection.native)/results.json","$($selection.native)/trace.txt",'run-all-tests-results.json')){
     $source=Join-Path $run $rel;if(Test-Path -LiteralPath $source -PathType Leaf){Copy-Item -LiteralPath $source -Destination (Join-Path $caseDir ([IO.Path]::GetFileName($rel)))}
    }
    foreach($name in @('intent.json','complete.json')){$source=Join-Path $testRoot "evidence/run-archive-records/$id/$name";if(Test-Path -LiteralPath $source -PathType Leaf){Copy-Item -LiteralPath $source -Destination (Join-Path $caseDir $name)}}
    foreach($origin in @('runs','evidence/completed-runs')){
     $runBase=Join-Path $testRoot "$origin/$id";$prefix=if($origin -ceq 'runs'){'live'}else{'archived'}
     $exits=Join-Path $runBase 'artifacts/native-process-exits'
     if(Test-Path -LiteralPath $exits -PathType Container){$exitDestination=Join-Path $caseDir "$prefix-native-process-exits";[void][IO.Directory]::CreateDirectory($exitDestination);Get-ChildItem -LiteralPath $exits -File -Filter '*.json' | ForEach-Object {Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $exitDestination $_.Name)}}
    }
    $validationState=Join-Path $testRoot "evidence/runs/$id/run-state.json";if(Test-Path -LiteralPath $validationState -PathType Leaf){Copy-Item -LiteralPath $validationState -Destination (Join-Path $caseDir 'validation-run-state.json')}
    $captures=Join-Path $run 'perf/viewer_space_captures';if($scope -ceq 'ViewerSpaceCapture' -and (Test-Path -LiteralPath $captures -PathType Container)){Copy-Item -LiteralPath $captures -Destination (Join-Path $caseDir 'viewer_space_captures') -Recurse}
    $evidence=Read-RSHostedDiagnosticEvidence -TestRoot $testRoot -RunId $id -Selection $selection
    $intentCopy=Join-Path $caseDir 'intent.json';$traceCopy=Join-Path $caseDir 'trace.txt'
    $nativeCopy=Join-Path $caseDir 'results.json';$summaryCopy=Join-Path $caseDir 'run-all-tests-results.json'
    if((Get-FileHash -LiteralPath $intentCopy -Algorithm SHA256).Hash -ine $evidence.archiveIntentSha256 -or
       (Get-FileHash -LiteralPath $nativeCopy -Algorithm SHA256).Hash -ine $evidence.nativeSha256 -or
       (Get-FileHash -LiteralPath $summaryCopy -Algorithm SHA256).Hash -ine $evidence.summarySha256){throw 'Retained native evidence differs from archive observations.'}
    $intent=Get-Content -LiteralPath $intentCopy -Raw | ConvertFrom-Json -AsHashtable
    $traceRelative="artifacts\selftest\last_run\$($selection.native)\trace.txt"
    $traceRows=@($intent.inventory|Where-Object {$_.path -ieq $traceRelative -and -not $_.directory})
    if($traceRows.Count -ne 1 -or [int64]$traceRows[0].length -ne (Get-Item -LiteralPath $traceCopy).Length -or
       $traceRows[0].sha256 -ine (Get-FileHash -LiteralPath $traceCopy -Algorithm SHA256).Hash){throw 'Retained native trace differs from closed archive inventory.'}
    Assert-CleanHead $CandidateRoot $ExpectedHead
    $after=Get-RSBuildSourceSnapshot -RepoRoot $CandidateRoot;$afterReceipt=Read-RSBuildReceipt -BuildOutputDir $output -SchemaPath $schema
    $verified=$after.SnapshotId -ceq $snapshot.SnapshotId -and $afterReceipt.receipt_id -ceq $receipt.receipt_id -and
     (Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash -ceq $receiptHash -and
     (Test-RSBuildReceiptForArtifactUse -Receipt $afterReceipt -SourceSnapshot $after -Platform ARM64 -Configuration Release -RepoRoot $CandidateRoot -RequiredRelativePath $required)
    [void](Test-RSHostedDiagnosticOutcome -ExitCode $code -RunId $id -Selection $selection -Summary $evidence.summary -Native $evidence.native -Platform ARM64 -Configuration Release -ReceiptVerified $verified -ExpectedExePath (Join-Path $output 'RedSalamander.exe') -InvocationStartedUtc $started -InvocationEndedUtc $ended)
    $observed.status='PASSED_FILTERED';$observed.native_sha256=$evidence.nativeSha256;$observed.summary_sha256=$evidence.summarySha256;$observed.intent_sha256=$evidence.archiveIntentSha256
    } catch {
     $caseFailures.Add("$scope`: $($_.Exception.Message)")
     if($null -ne $observed){$observed.status='FAILED_FILTERED'}else{$record.cases+=,[ordered]@{scope=$scope;run_id=$null;runner_exit=$null;log=(Join-Path $e "$scope-runner.log");status='FAILED_BEFORE_RUN_ID';native_sha256=$null;summary_sha256=$null;intent_sha256=$null}}
    }
   }
  }
 } finally {
  [Environment]::SetEnvironmentVariable('REDSALAMANDER_HOSTED_TEST_DESKTOP',$prior,'Process')
  try {
   $snapshotLog=Join-Path $e 'final-foreground-snapshot.log'
   $snapshotExit=Child $CandidateRoot (Join-Path $BootstrapRoot '.ci/WgcAdmissionProbe.ps1') @('-TestRoot',$testRoot,'-RepoRoot',$CandidateRoot,'-SnapshotOnly') $snapshotLog
   if($snapshotExit -ne 0){throw "Read-only foreground snapshot failed: $snapshotExit"}
   $matches=[regex]::Matches([IO.File]::ReadAllText($snapshotLog),'(?m)^Foreground snapshot: (.+?)\s*$')
   if($matches.Count -ne 1){throw 'Expected one read-only foreground snapshot.'}
   $snapshotPath=Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path $matches[0].Groups[1].Value -RequireDescendant
   $retained=Join-Path $e 'final-foreground-snapshot.json';Copy-Item -LiteralPath $snapshotPath -Destination $retained
   $record.foreground_snapshot=Get-Content -LiteralPath $retained -Raw | ConvertFrom-Json -AsHashtable
  } catch {$record.foreground_snapshot_error=$_.Exception.Message}
 }
 Assert-CleanHead $CandidateRoot $ExpectedHead
 $probeLog=Join-Path $e 'assertion-probe.log';& (Join-Path $output 'PluginContractTests.exe') --ci-assertion-probe *> $probeLog;$probeExit=$LASTEXITCODE
 $record.assertion_probe_exit=$probeExit
 if($probeExit -ne 21){throw "Assertion probe expected 21, got $probeExit"}
 $finalSnapshot=Get-RSBuildSourceSnapshot -RepoRoot $CandidateRoot;$finalReceipt=Read-RSBuildReceipt -BuildOutputDir $output -SchemaPath $schema
 if($finalSnapshot.SnapshotId -cne $snapshot.SnapshotId -or $finalReceipt.receipt_id -cne $receipt.receipt_id -or
    (Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash -cne $receiptHash -or
    -not (Test-RSBuildReceiptForArtifactUse -Receipt $finalReceipt -SourceSnapshot $finalSnapshot -Platform ARM64 -Configuration Release -RepoRoot $CandidateRoot -RequiredRelativePath $required)){throw 'Post-probe source or receipt changed.'}
 Assert-CleanHead $CandidateRoot $ExpectedHead
 if($record.foreground_snapshot_error -or $null -eq $record.foreground_snapshot){throw 'Final foreground snapshot is unavailable; see retained diagnostic.'}
 if($caseFailures.Count){throw ('Filtered diagnostics failed: '+($caseFailures -join '; '))}
 $record.status='PASSED_FILTERED_ONLY'
} catch {$record.error=$_.Exception.Message;throw}
finally {$record | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $e 'diagnostic.json') -Encoding utf8}
