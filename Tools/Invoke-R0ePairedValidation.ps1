<#
.SYNOPSIS
Runs a source-bound paired R0e diagnostic on one native hosted machine.
.DESCRIPTION
Builds an isolated baseline and measures five fresh filtered FileOps processes per side.
The existing workload, two-second ceiling, safety assertions and matched p95 bound are unchanged.
A timing-only RED baseline is retained and reports NOT_EVALUATED. No product, Full, or release
qualification receipt is produced. All generated test evidence stays in the marked sandbox.
.PARAMETER BaselineHead
Full lowercase reviewed baseline commit SHA available from origin.
.PARAMETER CandidateHead
Full lowercase current candidate SHA; must equal GITHUB_SHA and the clean checkout.
.PARAMETER Platform
Native x64 or ARM64 platform.
.PARAMETER BuildNumber
Frozen positive hosted batch build number used by both full solution builds.
.PARAMETER TestRoot
Existing marked X:\RedSalamander.Perf sandbox; this command never initializes it.
.PARAMETER EngineRoot
Existing checkout of the exact shared engine pin in .ci/project.json.
.PARAMETER WhatIf
Checks hosted identity and sandbox preconditions, then previews the build and test without creating evidence or changing the environment.
.OUTPUTS
A create-only cohort under TestRoot/evidence/r0e-paired containing status, logs,
source/receipt identities, archived-run references and a diagnostic assessment.
Exit 0 means green-pair diagnostic completed, never Full or production acceptance.
Exit 2 means a complete validated timing-only RED baseline, NOT_EVALUATED.
Exit 1 means incomplete, corrupt, failed, or otherwise invalid evidence.
.NOTES
Requires PowerShell 7.4, Git, the pinned native toolchain and an explicitly admitted hosted
native window. Creates and retains an isolated Git worktree and build outputs under .build;
may fetch the exact baseline object, install its pinned dependencies and build it. It never
cleans worktrees or kills unrelated processes. The existing test runner owns desktop leases.
.EXAMPLE
./Tools/Invoke-R0ePairedValidation.ps1 -BaselineHead 0123456789abcdef0123456789abcdef01234567 -CandidateHead 89abcdef0123456789abcdef0123456789abcdef -Platform x64 -BuildNumber 1 -TestRoot C:\RedSalamander.Perf -EngineRoot C:\src\engine
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact='Medium')]
param(
    [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{40}$')][string]$BaselineHead,
    [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{40}$')][string]$CandidateHead,
    [Parameter(Mandatory)][ValidateSet('x64','ARM64')][string]$Platform,
    [Parameter(Mandatory)][ValidateRange(1,65535)][int]$BuildNumber,
    [Parameter(Mandatory)][string]$TestRoot,
    [Parameter(Mandatory)][string]$EngineRoot
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
Import-Module (Join-Path $root 'Tools/Modules/Testing/R0ePairedValidation.psm1')
Import-Module (Join-Path $root 'Tools/Modules/Testing/TestSandbox.psm1')
Import-Module (Join-Path $root 'Tools/Modules/Build/BuildEvidence.psm1')
Import-Module (Join-Path $root 'Tools/Modules/Build/VcpkgInstallSafety.psm1')
$prior=@{}; $exitCode=1
function Set-ScopedEnvironment([string]$Name,[string]$Value) {
    if(-not $prior.ContainsKey($Name)){$prior[$Name]=[Environment]::GetEnvironmentVariable($Name,'Process')}
    [Environment]::SetEnvironmentVariable($Name,$Value,'Process')
}
function Invoke-R0eChild([string]$Repo,[string]$EntryPath,[string[]]$Arguments,[string]$LogPath) {
    $oldDirectory=[IO.Directory]::GetCurrentDirectory()
    Push-Location -LiteralPath $Repo
    try {
        [IO.Directory]::SetCurrentDirectory($Repo)
        & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $EntryPath @Arguments *> $LogPath
        return $LASTEXITCODE
    } finally {
        [IO.Directory]::SetCurrentDirectory($oldDirectory)
        Pop-Location
    }
}
function Require-CleanHead([string]$Repo,[string]$Head) {
    $actual=(& git -C $Repo rev-parse HEAD);if($LASTEXITCODE -ne 0 -or $actual -cne $Head){throw 'R0e checkout HEAD mismatch.'}
    $dirty=@(& git -C $Repo status --porcelain);if($LASTEXITCODE -ne 0 -or $dirty.Count -ne 0){throw 'R0e checkout is dirty or status failed.'}
}
function Read-VerifiedReceipt([string]$Repo,[string]$Head) {
    $out=Join-Path $Repo ".build/$Platform/Release"; $schema=Join-Path $Repo 'Specs/Build/BuildReceipt.schema.json'
    $receipt=Read-RSBuildReceipt -BuildOutputDir $out -SchemaPath $schema; $snapshot=Get-RSBuildSourceSnapshot -RepoRoot $Repo
    if(-not $receipt -or $receipt.git_head -cne $Head -or $receipt.target -cne 'solution' -or $receipt.platform -cne $Platform -or
       $receipt.configuration -cne 'Release' -or -not $receipt.tests_enabled -or "build-number=$BuildNumber" -cnotin $receipt.build_arguments -or
       'compiler-debug-information=embedded' -cnotin $receipt.build_arguments -or
       -not (Test-RSBuildReceiptForArtifactUse -Receipt $receipt -SourceSnapshot $snapshot -Platform $Platform -Configuration Release -RepoRoot $Repo -RequiredRelativePath ".build/$Platform/Release/RedSalamander.exe")) {throw 'R0e full Release receipt is unavailable or invalid.'}
    return $receipt
}
if($env:GITHUB_ACTIONS -cne 'true' -or $env:GITHUB_SHA -cne $CandidateHead -or
   $CandidateHead -ceq ('0' * 40) -or $BaselineHead -ceq ('0' * 40)){throw 'Hosted source identity is unavailable.'}
$arch=[System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
$processArch=[System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString()
if($arch -ine $Platform -or $processArch -ine $Platform){throw 'R0e requires a native matching process and OS architecture.'}
Require-CleanHead $root $CandidateHead
$engineProvider=$null;$engineDrive=$null
$EngineRoot=$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($EngineRoot,[ref]$engineProvider,[ref]$engineDrive)
if($engineProvider.Name -cne 'FileSystem'){throw 'EngineRoot must be a filesystem path.'}
$project=Get-Content -LiteralPath (Join-Path $root '.ci/project.json') -Raw|ConvertFrom-Json
Require-CleanHead $EngineRoot ([string]$project.engine.ref)
$TestRoot=Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path $TestRoot
if (-not $PSCmdlet.ShouldProcess("$Platform baseline $BaselineHead and candidate $CandidateHead", 'Create retained evidence and baseline worktree, build and run ten hosted diagnostics')) { return }
$cohortId='r0e-'+[Guid]::NewGuid().ToString('N')
$cohort=Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $TestRoot "evidence/r0e-paired/$cohortId") -RequireDescendant
if (Test-Path -LiteralPath $cohort) { throw 'R0e cohort already exists.' }
[void][IO.Directory]::CreateDirectory($cohort)
$running=Join-Path $cohort 'running.json'; $terminal=Join-Path $cohort 'terminal.json'
$runningBytes=[Text.UTF8Encoding]::new($false).GetBytes((@{schema='red-salamander.r0e-paired-running.v1';cohort=$cohortId;started_utc=[DateTime]::UtcNow.ToString('o');
    baseline=$BaselineHead;candidate=$CandidateHead;platform=$Platform;build_number=$BuildNumber}|ConvertTo-Json))
$runningStream=[IO.File]::Open($running,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
try { $runningStream.Write($runningBytes,0,$runningBytes.Length); $runningStream.Flush($true) } finally { $runningStream.Dispose() }
$state=[ordered]@{schema='red-salamander.r0e-paired-result.v1';cohort=$cohortId;outcome='INCOMPLETE';baseline_head=$BaselineHead;
    candidate_head=$CandidateHead;platform=$Platform;build_number=$BuildNumber;worktree=$null;candidate_receipt=$null;baseline_receipt=$null;
    baseline_case=$null;candidate_case=$null;current_run_id=$null;samples=@();assessment=$null;error=$null;ended_utc=$null}
try {
    Set-ScopedEnvironment 'RSBuildEnableTests' 'true';Set-ScopedEnvironment 'RSBuildCompilerDebugInformation' 'Embedded';Set-ScopedEnvironment 'REDSALAMANDER_TEST_ROOT' $TestRoot
    $candidateReceipt=Read-VerifiedReceipt $root $CandidateHead; $state.candidate_receipt=$candidateReceipt.receipt_id
    $candidateCase=Get-RSR0eCaseIdentity ([IO.File]::ReadAllText((Join-Path $root 'RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Phases05_06.cpp')))
    if($candidateCase.full -cne 'A3541BB88E72E6BC41EAD4C77D37F9136A39EF123C8AAB780ECEF63039F020CE'){throw 'Candidate R0e case is not the reviewed candidate.'}
    $state.candidate_case=$candidateCase
    $buildRoot=Join-Path $root '.build';$buildItem=Get-Item -LiteralPath $buildRoot -Force
    if(($buildItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){throw 'Build root is a reparse point.'}
    $baselineRoot=Join-Path $buildRoot $cohortId;$state.worktree=$baselineRoot
    if(Test-Path -LiteralPath $baselineRoot){throw 'Baseline checkout path already exists.'}
    & git -C $root cat-file -e "$BaselineHead^{commit}" 2>$null
    if($LASTEXITCODE -ne 0){
        & git -C $root fetch --no-tags --no-write-fetch-head origin $BaselineHead *> (Join-Path $cohort 'fetch.log')
        if($LASTEXITCODE -ne 0){throw 'Exact baseline commit fetch failed.'}
    }
    & git -C $root worktree add --detach $baselineRoot $BaselineHead *> (Join-Path $cohort 'worktree.log')
    if($LASTEXITCODE -ne 0){throw 'Baseline checkout failed; preserve the owned path.'}
    Require-CleanHead $baselineRoot $BaselineHead
    $baselineCase=Get-RSR0eCaseIdentity ([IO.File]::ReadAllText((Join-Path $baselineRoot 'RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Phases05_06.cpp')))
    $state.baseline_case=$baselineCase
    foreach($path in @('Dependencies/DxUi.lock.json','vcpkg.json','vcpkg-tool.json','vcpkg-configuration.json')){
        $candidateFile=Join-Path $root $path;$baselineFile=Join-Path $baselineRoot $path
        $candidateExists=Test-Path -LiteralPath $candidateFile -PathType Leaf;$baselineExists=Test-Path -LiteralPath $baselineFile -PathType Leaf
        if($candidateExists -ne $baselineExists -or (-not $candidateExists -and $path -cne 'vcpkg-configuration.json')){throw "Dependency input differs or is missing: $path"}
        if($candidateExists -and (Get-FileHash -LiteralPath $candidateFile -Algorithm SHA256).Hash -cne (Get-FileHash -LiteralPath $baselineFile -Algorithm SHA256).Hash){throw "Dependency input differs: $path"}
    }
    $tool=Initialize-RSVcpkgToolContext -RepoRoot $baselineRoot
    $installExit=Invoke-R0eChild $baselineRoot (Join-Path $baselineRoot 'vcpkg-install.ps1') @('-Platform',$Platform,'-VcpkgExe',$tool.Executable,'-VcpkgRoot',$tool.RegistryRoot) (Join-Path $cohort 'baseline-vcpkg.log')
    if($installExit -ne 0){throw 'Baseline dependency installation failed.'}
    $environment=Get-RSVcpkgBuildEnvironment -RepoRoot $baselineRoot -ToolRoot $tool.ToolRoot -Platform $Platform -Configuration Release
    foreach($entry in $environment.GetEnumerator()){Set-ScopedEnvironment $entry.Key ([string]$entry.Value)}
    Set-ScopedEnvironment 'VCPKG_REGISTRY_ROOT' $tool.RegistryRoot
    $buildExit=Invoke-R0eChild $baselineRoot (Join-Path $baselineRoot 'build.ps1') @('-Configuration','Release','-Platform',$Platform,'-BuildNumber',([string]$BuildNumber),'-MaxCpuCount','4') (Join-Path $cohort 'baseline-build.log')
    if($buildExit -ne 0){throw 'Baseline full-solution build failed.'}
    Require-CleanHead $baselineRoot $BaselineHead
    $baselineReceipt=Read-VerifiedReceipt $baselineRoot $BaselineHead;$state.baseline_receipt=$baselineReceipt.receipt_id
    foreach($property in @('toolchain_identity_digest','vcpkg_identity_digest','ghostty_runtime_identity','runtime_dependency_digest')){
        if($baselineReceipt.$property -cne $candidateReceipt.$property){throw "Baseline/candidate build identity differs: $property"}
    }
    # Restore the candidate environment before either side executes; build environment is not test data.
    foreach($name in @($prior.Keys)){[Environment]::SetEnvironmentVariable($name,$prior[$name],'Process')}
    Set-ScopedEnvironment 'RSBuildEnableTests' 'true';Set-ScopedEnvironment 'RSBuildCompilerDebugInformation' 'Embedded';Set-ScopedEnvironment 'REDSALAMANDER_TEST_ROOT' $TestRoot
    $all=@();$sideRows=@{baseline=@();candidate=@()}
    foreach($side in @('baseline','candidate')){
        $repo=if($side -ceq 'baseline'){$baselineRoot}else{$root};$head=if($side -ceq 'baseline'){$BaselineHead}else{$CandidateHead}
        $expected=if($side -ceq 'baseline'){$baselineReceipt}else{$candidateReceipt}
        for($index=1;$index -le 5;$index++){
            Require-CleanHead $repo $head
            $before=Read-VerifiedReceipt $repo $head;if($before.receipt_id -cne $expected.receipt_id){throw 'Pre-sample receipt changed.'}
            $log=Join-Path $cohort ("$side-{0:D2}.log" -f $index);$started=[DateTimeOffset]::UtcNow
            $sampleExit=Invoke-R0eChild $repo (Join-Path $repo 'Tools/Run-AllTests.ps1') @('-Suite','FileOps','-CaseFilter','FileOps_ProviderCapabilityMatrix','-Configuration','Release','-Platform',$Platform,'-BuildNumber',([string]$BuildNumber),'-TimeoutMultiplier','1','-SelfTestRepeat','1','-SkipBuild','-TestRoot',$TestRoot) $log
            $ended=[DateTimeOffset]::UtcNow
            $logText=[IO.File]::ReadAllText($log);$ids=[regex]::Matches($logText,'(?m)^\s*Run id:\s*(\d{8}T\d{6}Z-\d+-[0-9a-f]{32})\s*$')
            if($ids.Count -ne 1){throw 'R0e sample lacks a unique run ID; preserve log.'}
            $id=$ids[0].Groups[1].Value;$state.current_run_id=$id
            # Persist the runner identity before archive validation, including when closure is corrupt.
            $identified=[ordered]@{side=$side;sample=$index;head=$head;receipt=$expected.receipt_id;run_id=$id;runner_exit=$sampleExit;
                log=$log;log_sha256=(Get-FileHash -LiteralPath $log -Algorithm SHA256).Hash}
            $identifiedPath=Join-Path $cohort ("$side-{0:D2}-identified.json" -f $index)
            $identifiedBytes=[Text.UTF8Encoding]::new($false).GetBytes(($identified|ConvertTo-Json -Depth 10))
            $identifiedStream=[IO.File]::Open($identifiedPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
            try { $identifiedStream.Write($identifiedBytes,0,$identifiedBytes.Length); $identifiedStream.Flush($true) } finally { $identifiedStream.Dispose() }
            $archive=Read-RSR0eArchive -TestRoot $TestRoot -RunId $id
            $summaryCopy=Join-Path $cohort ("$side-{0:D2}-runner-summary.json" -f $index)
            $summaryStream=[IO.File]::Open($summaryCopy,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::Read)
            try { $summaryStream.Write($archive.summary_bytes,0,$archive.summary_bytes.Length); $summaryStream.Flush($true) } finally { $summaryStream.Dispose() }
            $archive.files['artifacts\selftest\last_run\run-all-tests-results.json'].retained_path=$summaryCopy
            $summary=$archive.summary_text | ConvertFrom-Json -AsHashtable
            if($summary.executable -ine (Join-Path $repo ".build/$Platform/Release/RedSalamander.exe") -or
                [DateTimeOffset]$summary.run_started_utc -lt $started -or [DateTimeOffset]$summary.run_ended_utc -gt $ended -or
                [DateTimeOffset]$summary.run_ended_utc -lt [DateTimeOffset]$summary.run_started_utc){throw 'R0e sample invocation identity mismatch.'}
            $partial=[ordered]@{side=$side;sample=$index;head=$head;receipt=$expected.receipt_id;run_id=$id;runner_exit=$sampleExit;
                archive=$archive.files;intent_sha256=$archive.intent_sha256;log=$log;log_sha256=(Get-FileHash -LiteralPath $log -Algorithm SHA256).Hash}
            $state.samples=@($all)+@([pscustomobject]$partial)
            [IO.File]::WriteAllText((Join-Path $cohort ("$side-{0:D2}-observed.json" -f $index)),($partial|ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
            $sample=Test-RSR0eSample -Summary $summary -Native $archive.native -Metrics $archive.metrics -Side $side -RunId $id -Platform $Platform -RunnerExit $sampleExit
            Require-CleanHead $repo $head;$after=Read-VerifiedReceipt $repo $head;if($after.receipt_id -cne $expected.receipt_id){throw 'Post-sample receipt changed.'}
            $record=[ordered]@{side=$side;sample=$index;head=$head;receipt=$expected.receipt_id;run_id=$id;runner_exit=$sampleExit;red=$sample.red;
                machine_hash=$sample.machine_hash;metric_run_id=$sample.metric_run_id;process=$sample.process;thread=$sample.thread;
                branch=$sample.branch;commit=$sample.commit;metrics=$sample.metrics;archive=$archive.files;intent_sha256=$archive.intent_sha256;log=$log;
                log_sha256=(Get-FileHash -LiteralPath $log -Algorithm SHA256).Hash}
            $all+=,[pscustomobject]$record;$state.samples=$all
            [IO.File]::WriteAllText((Join-Path $cohort ("$side-{0:D2}.json" -f $index)),($record|ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
            $sideRows[$side]+=,$sample
        }
    }
    $state.assessment=Get-RSR0eAssessment -Baseline $sideRows.baseline -Candidate $sideRows.candidate
    $state.outcome=$state.assessment.outcome
    if($state.assessment.baseline_red){$exitCode=2}else{$exitCode=0}
}catch{$state.error=$_.Exception.Message;Write-Error -ErrorAction Continue $state.error}
finally{
    foreach($name in @($prior.Keys)){[Environment]::SetEnvironmentVariable($name,$prior[$name],'Process')}
    $state.ended_utc=[DateTime]::UtcNow.ToString('o')
    $pendingTerminal=$terminal+'.pending'
    [IO.File]::WriteAllText($pendingTerminal,($state|ConvertTo-Json -Depth 15),[Text.UTF8Encoding]::new($false))
    [IO.File]::Move($pendingTerminal,$terminal)
}
Write-Host "R0e diagnostic $($state.outcome): $terminal"
exit $exitCode
