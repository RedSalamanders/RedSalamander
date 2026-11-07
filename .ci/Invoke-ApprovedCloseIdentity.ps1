<# PRIVATE metadata-only diagnostic: no test app, owned window, activation, input or capture. #>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$RepoRoot,[switch]$ApproveAccountClose,[switch]$ApproveSearchClose)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop';$PSNativeCommandUseErrorActionPreference=$false
if(-not $ApproveAccountClose -or $env:GITHUB_ACTIONS -cne 'true' -or $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or
   $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander' -or $env:GITHUB_EVENT_NAME -cne 'workflow_dispatch' -or
   [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -cne 'Arm64' -or
   [Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString() -cne 'Arm64') {throw 'Explicit manual hosted native ARM64 account-close authorization required.'}
$RepoRoot=[IO.Path]::GetFullPath($RepoRoot);$head=& git -C $RepoRoot rev-parse HEAD
if($LASTEXITCODE -ne 0 -or $head -cne $env:GITHUB_SHA){throw 'Bootstrap HEAD mismatch.'}
$dirty=@(& git -C $RepoRoot status --porcelain);if($LASTEXITCODE -ne 0 -or $dirty.Count){throw 'Bootstrap checkout dirty or status failed.'}
$testRoot='C:\RedSalamander.Perf';$env:REDSALAMANDER_TEST_ROOT=$testRoot
Import-Module (Join-Path $RepoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
[void](Initialize-RSTestSandboxRoot -TestRoot $testRoot -RepoRoot $RepoRoot -AllowExternalInitialization)
$scratch=New-RSTestSandboxScratchDirectory -RepoRoot $RepoRoot -Harness review-tools -Case approved-close-identity
$env:TEMP=$scratch;$env:TMP=$scratch
$e=Join-Path $testRoot "evidence/approved-close-identity/$($env:GITHUB_RUN_ID)-$($env:GITHUB_RUN_ATTEMPT)"
if(Test-Path -LiteralPath $e){throw 'Preserve previous metadata evidence.'};[void][IO.Directory]::CreateDirectory($e)
$record=[ordered]@{schema='private.approved-close-identity.v1';bootstrap_head=$head;status='INCOMPLETE';coverage='NOT_EVALUATED';close_exit=$null;snapshot_exit=$null;error=$null}
try {
 $probe=Join-Path $RepoRoot '.ci/WgcAdmissionProbe.ps1';$hostPath=Join-Path $PSHOME 'pwsh.exe'
 $closeLog=Join-Path $e 'close.log'
 & $hostPath -NoProfile -NonInteractive -File $probe -TestRoot $testRoot -RepoRoot $RepoRoot -CloseOnly -DismissHostedAccountSurface *> $closeLog
 $record.close_exit=$LASTEXITCODE
 $closeText=[IO.File]::ReadAllText($closeLog)
 foreach($item in @(@('Close intent','account-dismissal-intent.json'),@('Close result','account-dismissal-result.json'))){
  $matches=[regex]::Matches($closeText,'(?m)^'+[regex]::Escape($item[0])+': (.+?)\s*$')
  if($matches.Count -eq 1){$source=Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path $matches[0].Groups[1].Value -RequireDescendant;Copy-Item -LiteralPath $source -Destination (Join-Path $e $item[1])}
 }
 if($ApproveSearchClose){
  $record['search_requested']=$true;$record['search_exit']=$null;$record['search_dismissed']=$null;$record['search_no_target']=$null;$record['search_error']=$null
  try {
   if($record.close_exit -ne 0){throw 'Account close did not complete; Search close was not attempted.'}
   $accountIntentPath=Join-Path $e 'account-dismissal-intent.json';$accountResultPath=Join-Path $e 'account-dismissal-result.json'
   $accountIntent=Get-Content -LiteralPath $accountIntentPath -Raw|ConvertFrom-Json -AsHashtable
   $accountResult=Get-Content -LiteralPath $accountResultPath -Raw|ConvertFrom-Json -AsHashtable
   if($accountIntent.Head -cne $head -or $accountResult.IntentSha256 -ine (Get-FileHash -LiteralPath $accountIntentPath -Algorithm SHA256).Hash -or
      $accountResult.Dismissed -isnot [bool] -or $accountResult.NoTarget -isnot [bool] -or $accountResult.Dismissed -eq $accountResult.NoTarget -or $accountResult.Error){throw 'Account close proof invalid; Search close was not attempted.'}
   $searchLog=Join-Path $e 'search-close.log'
   & $hostPath -NoProfile -NonInteractive -File $probe -TestRoot $testRoot -RepoRoot $RepoRoot -CloseSearchOnly -DismissHostedSearchSurface *> $searchLog
   $record.search_exit=$LASTEXITCODE;$searchText=[IO.File]::ReadAllText($searchLog)
   foreach($item in @(@('Search intent','search-dismissal-intent.json'),@('Search result','search-dismissal-result.json'))){
    $searchMatches=[regex]::Matches($searchText,'(?m)^'+[regex]::Escape($item[0])+': (.+?)\s*$')
    if($searchMatches.Count -eq 1){$source=Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path $searchMatches[0].Groups[1].Value -RequireDescendant;Copy-Item -LiteralPath $source -Destination (Join-Path $e $item[1])}
   }
   $searchResultPath=Join-Path $e 'search-dismissal-result.json'
   if(Test-Path -LiteralPath $searchResultPath -PathType Leaf){
    $searchResult=Get-Content -LiteralPath $searchResultPath -Raw|ConvertFrom-Json -AsHashtable
    $record.search_dismissed=$searchResult.Dismissed;$record.search_no_target=$searchResult.NoTarget;$record.search_error=$searchResult.Error
   }
   $searchIntentPath=Join-Path $e 'search-dismissal-intent.json'
   if($record.search_exit -ne 0 -or -not(Test-Path -LiteralPath $searchIntentPath -PathType Leaf) -or -not(Test-Path -LiteralPath $searchResultPath -PathType Leaf)){throw 'Search close failed or did not retain its intent/result.'}
   $searchIntent=Get-Content -LiteralPath $searchIntentPath -Raw|ConvertFrom-Json -AsHashtable
   if($searchIntent.Head -cne $head -or $searchResult.IntentSha256 -ine (Get-FileHash -LiteralPath $searchIntentPath -Algorithm SHA256).Hash -or
      $searchResult.Dismissed -isnot [bool] -or $searchResult.NoTarget -isnot [bool] -or $searchResult.Dismissed -eq $searchResult.NoTarget -or $searchResult.Error){throw 'Search close intent/result did not match.'}
  } catch {if(-not $record.search_error){$record.search_error=$_.Exception.Message}}
 }
 # A fresh process records the revealed foreground even when the guarded close refused.
 $snapshotLog=Join-Path $e 'snapshot.log'
 & $hostPath -NoProfile -NonInteractive -File $probe -TestRoot $testRoot -RepoRoot $RepoRoot -SnapshotOnly *> $snapshotLog
 $record.snapshot_exit=$LASTEXITCODE
 $matches=[regex]::Matches([IO.File]::ReadAllText($snapshotLog),'(?m)^Foreground snapshot: (.+?)\s*$')
 if($matches.Count -eq 1){$source=Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path $matches[0].Groups[1].Value -RequireDescendant;Copy-Item -LiteralPath $source -Destination (Join-Path $e 'foreground-snapshot.json')}
 if($record.close_exit -ne 0 -or $record.snapshot_exit -ne 0 -or $matches.Count -ne 1 -or
    -not(Test-Path -LiteralPath (Join-Path $e 'account-dismissal-intent.json') -PathType Leaf) -or
    -not(Test-Path -LiteralPath (Join-Path $e 'account-dismissal-result.json') -PathType Leaf)){throw 'Approved close or metadata snapshot was incomplete; retained evidence is diagnostic only.'}
 $intentPath=Join-Path $e 'account-dismissal-intent.json';$resultPath=Join-Path $e 'account-dismissal-result.json'
 $intent=Get-Content -LiteralPath $intentPath -Raw|ConvertFrom-Json -AsHashtable;$result=Get-Content -LiteralPath $resultPath -Raw|ConvertFrom-Json -AsHashtable
 if($intent.Head -cne $head -or $result.IntentSha256 -ine (Get-FileHash -LiteralPath $intentPath -Algorithm SHA256).Hash -or
    $result.Dismissed -isnot [bool] -or $result.NoTarget -isnot [bool] -or $result.Dismissed -eq $result.NoTarget -or $result.Error){throw 'Close intent/result did not match.'}
 if($ApproveSearchClose -and $record.search_error){throw ('Search diagnostic failed: '+$record.search_error)}
 $record.status='OBSERVED_ONLY'
}catch{$record.error=$_.Exception.Message;throw}
finally{$record|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $e 'diagnostic.json') -Encoding utf8}
