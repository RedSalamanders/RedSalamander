#requires -Version 7.0
[CmdletBinding()]
param(
    [string]$RepoRoot='C:/RsI26Adopt',
    [string]$FrozenEvidenceRoot='C:/RedSalamander.Perf/evidence/i26-ui/interaction-v58-20261001',
    [string]$BuildReceiptPath='C:/RedSalamander.Perf/evidence/i26-ui/final-gates-drivers-20260930/debug-v58-build/build-receipt.json',
    [string]$EvidenceRoot='C:/RedSalamander.Perf/evidence/i26-ui/final-gates-drivers-20260930/debug-v58-source-guards',
    [string]$TestRoot='V:/RedSalamander.Perf',
    [string]$AlternateTestRoot='C:/RedSalamander.Perf',
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{64}$')][string]$ExpectedManifestSha256
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath($RepoRoot)
$frozen=[IO.Path]::GetFullPath($FrozenEvidenceRoot)
$out=[IO.Path]::GetFullPath($EvidenceRoot)
$test=[IO.Path]::GetFullPath($TestRoot)
$alternate=[IO.Path]::GetFullPath($AlternateTestRoot)
$receipt=[IO.Path]::GetFullPath($BuildReceiptPath)
$manifest=Join-Path $frozen 'candidate-inputs.json'
$inventory=Join-Path $frozen 'source-inventory.json'
$pinScopePath=Join-Path $frozen 'pin-scope.json'
$frozenSource=Join-Path $frozen 'frozen/source'
$guardPath=Join-Path $PSScriptRoot 'candidate-guards.ps1'
$runner=Join-Path $repo 'Tools/Run-AllTests.ps1'
$names=@('file_operations_phase0_typed_contract_source_guard','file_operations_preparing_lifecycle_source_guard')
$filter=[string]::Join(',',$names)
if(Test-Path -LiteralPath $out){throw "Refusing to overwrite evidence: $out"}
foreach($path in @($manifest,$inventory,$pinScopePath,$receipt,$guardPath,$runner)){
    if(-not(Test-Path -LiteralPath $path -PathType Leaf)){throw "Required input missing: $path"}
}
if(-not(Test-Path -LiteralPath $frozenSource -PathType Container)){throw 'Frozen source tree missing.'}
if((Get-FileHash -LiteralPath $manifest -Algorithm SHA256).Hash.ToLowerInvariant() -cne $ExpectedManifestSha256.ToLowerInvariant()){
    throw 'V58 frozen manifest does not match the root-reviewed identity.'
}
foreach($root in @($test,$alternate)){
    $mark=Join-Path $root '.red-salamander-test-root'
    if(-not(Test-Path -LiteralPath $root -PathType Container) -or
       -not(Test-Path -LiteralPath $mark -PathType Leaf) -or
       (Get-Content -LiteralPath $mark -Raw).Trim() -cne 'red-salamander.test-sandbox-root.v1'){
        throw "Unmarked test root: $root"
    }
    $volume=Get-Volume -DriveLetter ([IO.Path]::GetPathRoot($root)[0]) -ErrorAction Stop
    if($volume.SizeRemaining -lt 1GB){throw "Focused root has less than 1 GiB free: $root"}
}
. $guardPath
$candidateArgs=@{RepoRoot=$repo;Configuration='Debug';FrozenInputsPath=$manifest;
    FrozenSourceRoot=$frozenSource;FrozenInventoryPath=$inventory;BuildReceiptPath=$receipt}
$before=Assert-I26Candidate @candidateArgs
$rows=@(Get-Content -LiteralPath $manifest -Raw|ConvertFrom-Json)
$physical=@($rows|ForEach-Object{([string]$_.path).Replace('/','\')}|Sort-Object -Unique)
$scope=Get-Content -LiteralPath $inventory -Raw|ConvertFrom-Json
$pin=Get-Content -LiteralPath $pinScopePath -Raw|ConvertFrom-Json
$qualified='19ca44d95a0cd493fc7fb1a57f4d4ad227bdefc1'
if($rows.Count -ne 85 -or $physical.Count -ne 84 -or $scope.logicalInputCount -ne 85 -or
   $scope.physicalInputCount -ne 84 -or ([string]$scope.candidateDxUiSha).ToLowerInvariant() -cne $qualified -or
   ([string]$pin.candidatePin).ToLowerInvariant() -cne $qualified){throw 'V58 frozen scope/pin is not exact 85/84 pin19.'}
$exe=Join-Path $repo '.build/x64/Debug/RedSalamander.exe'
$exeSha=Get-I26Sha256 $exe
$source=Join-Path $repo 'RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp'
$sourceText=Get-Content -LiteralPath $source -Raw
foreach($name in $names){
    $pattern='SelfTest::RunCase\(\s*options\s*,\s*suite\s*,\s*L"'+[regex]::Escape($name)+'"'
    if([regex]::Matches($sourceText,$pattern).Count -ne 1){throw "Exact Commands registration missing/duplicate: $name"}
}
Import-Module (Join-Path $repo 'Tools/Modules/Testing/TestRunPlan.psm1') -Force -ErrorAction Stop
$plan=@(Get-RSTestRunPlan -Suite Commands -CaseFilter $filter -RepoRoot $repo -Platform x64 `
    -Configuration Debug -RedSalamanderExePath $exe -TimeoutMultiplier 3 -RepeatCount 1)
if($plan.Count -ne 1 -or -not[bool]$plan[0].RequiresInteractiveDesktop){throw 'Commands desktop lease classification drifted.'}
$arguments=@($plan[0].Arguments|ForEach-Object{[string]$_})
if($arguments -notcontains ('--selftest-case='+$filter) -or
   @($arguments|Where-Object{$_ -like '--selftest-shuffle=*'}).Count -ne 0){throw 'Focused Commands filter/shuffle plan drifted.'}
[void](New-Item -ItemType Directory -Path $out)
$sourceArchive=Join-Path $out 'frozen-source'
Copy-Item -LiteralPath $frozenSource -Destination $sourceArchive -Recurse
foreach($row in $rows){
    $archived=Join-Path $sourceArchive ([string]$row.path)
    if(-not(Test-Path -LiteralPath $archived -PathType Leaf) -or
       (Get-I26Sha256 $archived) -cne ([string]$row.sha256).ToLowerInvariant()){
        throw "Archived source fails manifest: $($row.path)"
    }
}
foreach($path in @($manifest,$inventory,$pinScopePath,$receipt,$guardPath,$PSCommandPath)){
    Copy-Item -LiteralPath $path -Destination (Join-Path $out (Split-Path -Leaf $path))
}
[ordered]@{schema='i26-v58-source-guard-preflight.v1';utc=[DateTimeOffset]::UtcNow.ToString('o');
    candidate=$before;manifestSha256=$ExpectedManifestSha256.ToLowerInvariant();receiptSha256=(Get-I26Sha256 $receipt);
    executableSha256=$exeSha;scriptSha256=(Get-I26Sha256 $PSCommandPath);caseFilter=$filter;
    expectedNames=$names;planArguments=$arguments;requiresInteractiveDesktop=$true;
    testRoot=$test;alternateTestRoot=$alternate}|
    ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $out 'preflight.json') -Encoding utf8
$oldLocation=Get-Location
$oldPrimary=[Environment]::GetEnvironmentVariable('REDSALAMANDER_TEST_ROOT','Process')
$oldAlternate=[Environment]::GetEnvironmentVariable('REDSALAMANDER_TEST_ALTERNATE_ROOT','Process')
try{
    $env:REDSALAMANDER_TEST_ROOT=$test
    $env:REDSALAMANDER_TEST_ALTERNATE_ROOT=$alternate
    Set-Location -LiteralPath $repo
    [void](Assert-I26Candidate @candidateArgs)
    if((Get-I26Sha256 $exe) -cne $exeSha){throw 'Debug executable changed before focused run.'}
    $log=Join-Path $out 'runner.log'
    & $runner -Suite Commands -CaseFilter $filter -Configuration Debug -Platform x64 `
        -SkipBuild -TestRoot $test -AllowAlternateVolumeTestRoot -TimeoutMultiplier 3 *> $log
    $exitCode=$LASTEXITCODE
    $archive=Join-Path $out 'runner-run'
    $runId=Save-I26RunnerRun -LogPath $log -TestRoot $test -Destination $archive
    [ordered]@{runId=$runId;exitCode=$exitCode;filter=$filter;archive=$archive}|
        ConvertTo-Json|Set-Content -LiteralPath (Join-Path $out 'invocation.json') -Encoding utf8
    $after=Assert-I26Candidate @candidateArgs
    if((Get-I26Sha256 $exe) -cne $exeSha){throw 'Debug executable changed during focused run.'}
    if($exitCode -ne 0){throw "Focused Commands exited $exitCode; original run retained."}
    $lastRun=Join-Path $archive 'artifacts/selftest/last_run'
    $v1=Get-Content -LiteralPath (Join-Path $lastRun 'run-all-tests-results-v1.json') -Raw|ConvertFrom-Json
    $native=Get-Content -LiteralPath (Join-Path $lastRun 'commands/results.json') -Raw|ConvertFrom-Json
    $cases=@($native.cases)
    $actual=@($cases|ForEach-Object{[string]$_.name}|Sort-Object -CaseSensitive)
    $expected=@($names|Sort-Object -CaseSensitive)
    if($v1.schema -cne 'red-salamander.run-all-tests.v1' -or $v1.run_id -cne $runId -or
       $v1.suite -cne 'Commands' -or $v1.platform -cne 'x64' -or $v1.configuration -cne 'Debug' -or
       $v1.case_filter -cne $filter -or $v1.total -ne 2 -or $v1.passed -ne 2 -or
       $v1.failed -ne 0 -or $v1.skipped -ne 0 -or $v1.exit_code -ne 0 -or
       -not[bool]$v1.test_sandbox_audit.is_clean -or $v1.test_sandbox_audit.issue_count -ne 0 -or
       $native.suite -cne 'Commands' -or $native.repeat_count -ne 1 -or $native.passed -ne 2 -or
       $native.failed -ne 0 -or $native.skipped -ne 0 -or $cases.Count -ne 2 -or
       [string]::Join([char]10,$actual) -cne [string]::Join([char]10,$expected) -or
       @($cases|Where-Object{$_.status -cne 'passed' -or $_.repeat_index -ne 1}).Count -ne 0){
        throw 'Two exact source-guard cases did not pass 2/0/0 in one clean Commands run.'
    }
    [ordered]@{status='passed';runId=$runId;caseFilter=$filter;cases=$actual;passed=2;failed=0;skipped=0;
        sandboxClean=$true;before=$before;after=$after;receiptSha256=(Get-I26Sha256 $receipt)}|
        ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $out 'qualification.json') -Encoding utf8
}catch{
    $_|Out-String|Set-Content -LiteralPath (Join-Path $out 'driver-error.txt') -Encoding utf8
    throw
}finally{
    Set-Location -LiteralPath $oldLocation
    [Environment]::SetEnvironmentVariable('REDSALAMANDER_TEST_ROOT',$oldPrimary,'Process')
    [Environment]::SetEnvironmentVariable('REDSALAMANDER_TEST_ALTERNATE_ROOT',$oldAlternate,'Process')
}
