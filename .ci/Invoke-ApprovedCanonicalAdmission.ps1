<# PRIVATE fixture admission diagnostic. This is not a product or native qualification. #>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$CandidateRoot,[Parameter(Mandatory)][string]$BootstrapRoot,
      [switch]$ApproveAccountClose,[switch]$ApproveSearchClose,
      [ValidateSet('Canonical','Explicit','Inherited')][string]$AttachmentVariant='Canonical')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop';$PSNativeCommandUseErrorActionPreference=$false
$expected='3871a05da6109106d26b53ede1937e5ef29ec341'
if(-not $ApproveAccountClose -or -not $ApproveSearchClose -or $env:GITHUB_ACTIONS -cne 'true' -or
   $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander' -or
   $env:GITHUB_EVENT_NAME -cne 'workflow_dispatch' -or
   [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -cne 'Arm64' -or
   [Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString() -cne 'Arm64'){
    throw 'Both approvals and the exact manual project-hosted native ARM64 role are required.'
}
$CandidateRoot=[IO.Path]::GetFullPath($CandidateRoot);$BootstrapRoot=[IO.Path]::GetFullPath($BootstrapRoot)
foreach($pair in @(@($CandidateRoot,$expected),@($BootstrapRoot,$env:GITHUB_SHA))){
    $head=(& git -C $pair[0] rev-parse HEAD | Out-String).Trim()
    if($LASTEXITCODE -ne 0 -or $head -cne $pair[1]){throw 'Exact checkout identity is unavailable or mismatched.'}
    $dirty=@(& git -C $pair[0] status --porcelain)
    if($LASTEXITCODE -ne 0 -or $dirty.Count){throw 'Checkout is dirty or status failed.'}
}
$modulePath=Join-Path $CandidateRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1'
$moduleHash=(Get-FileHash -LiteralPath $modulePath -Algorithm SHA256).Hash
$testRoot='C:\RedSalamander.Perf';$priorRoot=$env:REDSALAMANDER_TEST_ROOT;$priorTemp=$env:TEMP;$priorTmp=$env:TMP
$priorDesktop=$env:REDSALAMANDER_HOSTED_TEST_DESKTOP;$priorStock=$env:RS_HOSTED_ARM64_STOCK_PREPARATION
$env:REDSALAMANDER_TEST_ROOT=$testRoot
Import-Module (Join-Path $CandidateRoot 'Tools/Modules/Testing/TestSandbox.psm1')
[void](Initialize-RSTestSandboxRoot -TestRoot $testRoot -RepoRoot $CandidateRoot -AllowExternalInitialization)
$scratch=New-RSTestSandboxScratchDirectory -RepoRoot $CandidateRoot -Harness review-tools -Case approved-canonical-admission
$env:TEMP=$scratch;$env:TMP=$scratch
$e=Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path (Join-Path $testRoot "evidence/approved-canonical-admission/$($env:GITHUB_RUN_ID)-$($env:GITHUB_RUN_ATTEMPT)-$AttachmentVariant") -RequireDescendant
if(Test-Path -LiteralPath $e){throw 'Preserve prior admission evidence.'}
[void][IO.Directory]::CreateDirectory($e)
function Write-PrivateCreateOnlyJson([string]$Path,[object]$Value){
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($Value|ConvertTo-Json -Depth 8))
    $stream=[IO.File]::Open($Path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
    try{$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)}finally{$stream.Dispose()}
}
$variant=$null
if($AttachmentVariant -cne 'Canonical'){
 . (Join-Path $BootstrapRoot '.ci/Prepare-PrivateDesktopVariant.ps1')
 $variant=Get-PrivateBoundaryVariant -ModulePath $modulePath -Variant $AttachmentVariant
}
$intent=[ordered]@{schema='private.canonical-admission-intent.v1';candidate_head=$expected;candidate_module_sha256=$moduleHash;
    bootstrap_head=$env:GITHUB_SHA;run_id=$env:GITHUB_RUN_ID;run_attempt=$env:GITHUB_RUN_ATTEMPT;job=$env:GITHUB_JOB;
    account_approved=[bool]$ApproveAccountClose;search_approved=[bool]$ApproveSearchClose;
    os_architecture=[Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString();process_architecture=[Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString();
    fixture_scope=if($variant){'PRIVATE_TRANSFORMED_FIXTURE'}else{'CANONICAL_FIXTURE'};attachment_variant=$AttachmentVariant;
    original_csharp_sha256=if($variant){$variant.OriginalSha256}else{$null};compiled_csharp_sha256=if($variant){$variant.CompiledSha256}else{$null};
    exact_replacements=if($variant){$variant.Replacements}else{@()};
    coverage='NOT_EVALUATED';created_utc=[DateTimeOffset]::UtcNow.ToString('o')}
Write-PrivateCreateOnlyJson (Join-Path $e 'intent.json') $intent
$record=[ordered]@{schema='private.canonical-admission-result.v1';intent_sha256=(Get-FileHash -LiteralPath (Join-Path $e 'intent.json')).Hash;
    entered=$false;completed=$false;desktop_request=$null;stock_request=$null;error=$null;source_after_verified=$false;cleanup_errors=@();coverage='NOT_EVALUATED'}
$failure=$null;$transcriptStarted=$false;$cleanupFailures=[Collections.Generic.List[Exception]]::new()
try{
 Start-Transcript -LiteralPath (Join-Path $e 'fixture-transcript.txt') -NoClobber | Out-Null;$transcriptStarted=$true
 if($variant){Add-Type -TypeDefinition $variant.Compiled}
 Import-Module $modulePath
 . (Join-Path $BootstrapRoot '.ci/Read-HostedInputCapability.ps1')
 try { $beforeCapability=Get-PrivateHostedInputCapability } catch { $beforeCapability=@{schema='private.hosted-input-capability.v1';error=$_.Exception.ToString()} }
 Write-PrivateCreateOnlyJson (Join-Path $e 'input-capability-before.json') $beforeCapability
 $desktop=Get-RSHostedInputDesktopRequest -PairedComparison:$false
 $stock=Get-RSHostedStockPreparationRequest -Phase Qualification -Platform ARM64 -QualificationScope CommandsHandoff
 $record.desktop_request=$desktop;$record.stock_request=$stock
 if($desktop -cne 'default' -or $stock -cne '1'){throw 'Canonical fixture role did not select the expected hosted requests.'}
 $env:REDSALAMANDER_HOSTED_TEST_DESKTOP=$desktop;$env:RS_HOSTED_ARM64_STOCK_PREPARATION=$stock
 # Private caller only: the canonical runner lease is not exported. The same
 # session-local mutex remains on this thread until both canonical wrappers return.
 $mutex=[Threading.Mutex]::new($false,'Local\RedSalamander.InteractiveTestRun.v1');$held=$false
 try{
  try{$held=$mutex.WaitOne([TimeSpan]::FromMinutes(2))}catch [Threading.AbandonedMutexException]{$held=$true;Write-Warning 'Acquired abandoned interactive mutex.'}
  if(-not $held){throw 'Timed out waiting for the canonical interactive mutex.'}
  Invoke-RSWithHostedTestDesktop {
   Invoke-RSWithHostedDefaultDesktop { $record.entered=$true }
  }
  $record.completed=$true
 }finally{if($held){$mutex.ReleaseMutex()};$mutex.Dispose()}
}catch{$failure=$_;$record.error=$_.Exception.ToString()}
finally{
 try {
  if (Get-Command Get-PrivateHostedInputCapability -ErrorAction SilentlyContinue) {
   try { $afterCapability=Get-PrivateHostedInputCapability } catch { $afterCapability=@{schema='private.hosted-input-capability.v1';error=$_.Exception.ToString()} }
  } else { $afterCapability=@{schema='private.hosted-input-capability.v1';error='Capability helper unavailable before fixture operation.'} }
  Write-PrivateCreateOnlyJson (Join-Path $e 'input-capability-after.json') $afterCapability
 } catch { $cleanupFailures.Add($_.Exception) }
 try{
  $afterHead=(& git -C $CandidateRoot rev-parse HEAD | Out-String).Trim()
  if($LASTEXITCODE -ne 0 -or $afterHead -cne $expected){throw 'Candidate head changed during admission.'}
  $afterDirty=@(& git -C $CandidateRoot status --porcelain)
  if($LASTEXITCODE -ne 0 -or $afterDirty.Count -or (Get-FileHash -LiteralPath $modulePath -Algorithm SHA256).Hash -cne $moduleHash){throw 'Candidate source or fixture module changed during admission.'}
  $record.source_after_verified=$true
 }catch{$cleanupFailures.Add($_.Exception)}
 if($transcriptStarted){try{Stop-Transcript | Out-Null}catch{$cleanupFailures.Add($_.Exception)}}
 $env:REDSALAMANDER_HOSTED_TEST_DESKTOP=$priorDesktop;$env:RS_HOSTED_ARM64_STOCK_PREPARATION=$priorStock
 $env:REDSALAMANDER_TEST_ROOT=$priorRoot;$env:TEMP=$priorTemp;$env:TMP=$priorTmp
 $record.cleanup_errors=@($cleanupFailures | ForEach-Object {$_.ToString()})
 try{Write-PrivateCreateOnlyJson (Join-Path $e 'result.json') $record}catch{
  $allFailures=[Collections.Generic.List[Exception]]::new()
  if($failure){$allFailures.Add($failure.Exception)}
  foreach($cleanupFailure in $cleanupFailures){$allFailures.Add($cleanupFailure)}
  $allFailures.Add($_.Exception)
  throw [AggregateException]::new('Admission and/or evidence completion failed.', $allFailures.ToArray())
 }
}
if($failure -and $cleanupFailures.Count){throw [AggregateException]::new('Admission and cleanup failed.', [Exception[]](@($failure.Exception)+@($cleanupFailures.ToArray())))}
if($failure){throw $failure}
if($cleanupFailures.Count){throw [AggregateException]::new('Admission cleanup failed.', $cleanupFailures.ToArray())}
if(-not $record.entered -or -not $record.completed){throw 'Canonical admission did not complete.'}
