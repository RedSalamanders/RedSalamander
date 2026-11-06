<#
.SYNOPSIS
Adapts the project commands to the pinned organization CI engine.
.DESCRIPTION
Runs the reviewed compiler-free PR subset or the admitted native build/test/package
role. Native build roles restore pinned dependencies beneath .build and apply
process-local integration before invoking canonical project commands. Test scratch
and evidence use the governed sandbox. No product publication is performed here.
.PARAMETER Phase
Compiler-free PR checks, Release test build, regression execution, production ZIP, or explicit qualification.
.PARAMETER EngineRoot
Checkout of the exact shared engine commit in .ci/project.json. Relative paths resolve from the current PowerShell filesystem location.
.PARAMETER Platform
Matching native x64 or ARM64 hosted runner.
.PARAMETER Descriptor
Frozen main admission descriptor; required for native phases.
.PARAMETER Base
Immutable PR comparison base for changelog validation.
.PARAMETER Configuration
Debug, Release or ASan Debug for an explicit qualification only.
.PARAMETER QualificationScope
Normal regression, explicit Full qualification, a fixed filtered diagnostic, the manual R0e paired study, or read-only DesktopProbe; never advances automatic Release state.
.PARAMETER DevelopmentEngine
Allows a local engine under development for compiler-free checks only.
.OUTPUTS
Canonical test evidence, source-bound receipts and qualified ZIPs; GitHub job output paths.
.NOTES
Requires PowerShell 7.4 and Git. Native build phases bootstrap the pinned vcpkg
tool/baseline and install the selected platform dependencies before building.
Native work is hosted-only and uses the matching architecture.
Static never invokes compiler fixtures. No publication or source/pin updates.
Filtered diagnostics require an exact passing case and remain NOT_EVALUATED for repository coverage.
DesktopProbe records hosted desktop identities and access errors without building, input or display changes.
See Specs/Build/Build_UnifiedCI.md for roles, input rules, evidence and manual actions.
.EXAMPLE
./Tools/Invoke-ProjectCi.ps1 -Phase Static -EngineRoot ../RedSalamanders.CI -Base origin/main -DevelopmentEngine
#>
[CmdletBinding()]
param([Parameter(Mandatory)][ValidateSet('Static','ReleaseTestBuild','ReleaseTest','Production','Qualification')][string]$Phase,
    [Parameter(Mandatory)][string]$EngineRoot, [ValidateSet('x64','ARM64')][string]$Platform = 'x64',
    [string]$Descriptor = '', [string]$Base = 'origin/main',
    [ValidateSet('Debug','Release','ASan Debug')][string]$Configuration = 'Release', [ValidateSet('Normal','Full','CommandsHandoff','FileOpsVolumes','ViewerSpaceCapture','R0ePaired','DesktopProbe')][string]$QualificationScope = 'Normal', [switch]$DevelopmentEngine)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($QualificationScope -in @('CommandsHandoff','FileOpsVolumes','ViewerSpaceCapture','R0ePaired','DesktopProbe') -and $Phase -ne 'Qualification') {
    throw 'Filtered diagnostic scopes are only valid for explicit Qualification.'
}
if ($QualificationScope -ceq 'R0ePaired' -and $Configuration -cne 'Release') { throw 'R0e paired validation requires Release.' }
if ($QualificationScope -ceq 'DesktopProbe' -and $env:GITHUB_ACTIONS -cne 'true') { throw 'Hosted desktop diagnostics are restricted to GitHub Actions.' }
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$engineProvider = $null
$engineDrive = $null
$EngineRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($EngineRoot, [ref]$engineProvider, [ref]$engineDrive)
if ($engineProvider.Name -cne 'FileSystem') { throw 'EngineRoot requires the FileSystem provider.' }
Import-Module (Join-Path $EngineRoot 'Tools/Modules/Ci.Policy.psm1') -ErrorAction Stop
Import-Module (Join-Path $EngineRoot 'Tools/Modules/Ci.Project.psm1') -ErrorAction Stop
$context = New-CiProjectContext -Root $root -EngineRoot $EngineRoot -Phase $Phase -Platform $Platform -Descriptor $Descriptor -DevelopmentEngine:$DevelopmentEngine
$pairedDispatch = $null
if ($QualificationScope -ceq 'R0ePaired') {
    Import-Module (Join-Path $root 'Tools/Modules/Testing/HostedDiagnosticEvidence.psm1')
    if ([string]::IsNullOrWhiteSpace($env:GITHUB_EVENT_PATH) -or -not (Test-Path -LiteralPath $env:GITHUB_EVENT_PATH -PathType Leaf)) {
        throw 'R0e paired validation requires the manual dispatch event payload.'
    }
    $pairedDispatch = Get-RSHostedPairedDispatch -EventName $env:GITHUB_EVENT_NAME -EventJson ([IO.File]::ReadAllText($env:GITHUB_EVENT_PATH)) -CandidateHead $env:GITHUB_SHA
    $candidateHead = (Invoke-CiGit $root @('rev-parse','HEAD')).stdout.Trim()
    $candidateStatus = (Invoke-CiGit $root @('status','--porcelain')).stdout.Trim()
    $baselineCommit = (Invoke-CiGit $root @('rev-parse','--verify',($pairedDispatch.baseline + '^{commit}'))).stdout.Trim()
    if ($candidateHead -cne $pairedDispatch.candidate -or $context.batch.commit -cne $pairedDispatch.candidate -or $candidateStatus) {
        throw 'R0e candidate checkout must be clean and match the frozen descriptor and dispatch SHA.'
    }
    if ($baselineCommit -cne $pairedDispatch.baseline) { throw 'R0e baseline must name the exact available commit object.' }
}
function Run([string]$Script, [string[]]$CommandArguments = @()) {
    Invoke-CiProjectScript -Context $context -Script $Script -Arguments $CommandArguments
}
if ($QualificationScope -ceq 'DesktopProbe') {
    Import-Module (Join-Path $root 'Tools/Modules/Testing/HostedTestDesktop.psm1')
    Import-Module (Join-Path $root 'Tools/Modules/Testing/TestSandbox.psm1')
    $probeRoot = Join-Path ([IO.Path]::GetPathRoot($root)) 'RedSalamander.Perf'
    $null = Initialize-RSTestSandboxRoot -TestRoot $probeRoot -RepoRoot $root -AllowExternalInitialization
    $probeDirectory = Assert-RSTestSandboxContainedPath -TestRoot $probeRoot -Path (Join-Path $probeRoot "evidence/desktop-probe/$($env:GITHUB_RUN_ID)/$Platform") -RequireDescendant
    [void][IO.Directory]::CreateDirectory($probeDirectory)
    $probe = [ordered]@{ schema='red-salamander.hosted-desktop-probe.v1'; head=$context.batch.commit; platform=$Platform;
        observed_utc=[DateTime]::UtcNow.ToString('o'); repository_verdict='NOT_EVALUATED'; desktop=(Get-RSHostedDesktopDiagnostics) }
    $probe | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $probeDirectory 'desktop.json') -Encoding utf8
    if ($env:GITHUB_OUTPUT) { Add-Content -LiteralPath $env:GITHUB_OUTPUT -Value ('evidence_directory=' + (Join-Path $probeRoot 'evidence')) }
    Write-Host ('Read-only hosted desktop probe: ' + ($probe | ConvertTo-Json -Depth 8 -Compress))
    return
}
function Test-PinnedDxUiMatrix {
    Import-Module (Join-Path $root 'Tools/Modules/Build/DxUiDependency.psm1')
    $sourceIdentity=Get-RSDxUiSourceIdentity -RepoRoot $root
    $validator=Join-Path $sourceIdentity.Source 'validate-build-matrix.ps1'
    Run ([IO.Path]::GetRelativePath($root,$validator)) @('-Root',$root)
}
if ($Phase -eq 'Static') {
    $env:CI_NATIVE_BUILDS_DISABLED = 'true'
    Test-CiProjectConfiguration $context
    & (Join-Path $EngineRoot 'Tools/Test-Workflows.ps1') -Root $root
    $baseCommit = Resolve-CiCommit $root $Base
    Run 'Tools/Test-Changelog.ps1' @('-Base',$baseCommit)
    foreach ($script in @('Tools/Verify-NoSubclassManager.ps1','Tools/Audit-ComctlReportSurfaces.ps1','Tools/Audit-VisibleNativeSurfaces.ps1')) { Run $script }
    Run 'Tools/Get-ToolInventory.ps1' @('-FailOnFindings','-Format','Json')
    Run 'Tools/Get-SpecInventory.ps1' @('-FailOnFindings')
    return
}
if ($Phase -ne 'Qualification') { $Configuration = 'Release' }
Import-Module (Join-Path $root 'Tools/Modules/Build/VcpkgInstallSafety.psm1')
if ($Phase -in @('ReleaseTestBuild','Production','Qualification')) {
    $vcpkg = Initialize-RSVcpkgToolContext -RepoRoot $root
    Run 'vcpkg-install.ps1' @('-Platform',$Platform,'-VcpkgExe',$vcpkg.Executable,'-VcpkgRoot',$vcpkg.RegistryRoot)
}
# The test step must retain the build's environment identity without reinstalling.
$toolRoot = Resolve-RSVcpkgSafeChildPath -Root $root -Child '.build/vcpkg-tool'
$environment = Get-RSVcpkgBuildEnvironment -RepoRoot $root -ToolRoot $toolRoot -Platform $Platform -Configuration $Configuration
foreach ($entry in $environment.GetEnumerator()) { [Environment]::SetEnvironmentVariable($entry.Key,$entry.Value,'Process') }
$env:VCPKG_REGISTRY_ROOT = Resolve-RSVcpkgSafeChildPath -Root $root -Child '.build/vcpkg-registry'
$number = [string]$context.batch.buildNumber
$profile = @('-Platform',$Platform,'-Configuration',$Configuration)
$started = [DateTime]::UtcNow
Import-Module (Join-Path $root 'Tools/Modules/Testing/TestSandbox.psm1') -Force
$env:REDSALAMANDER_TEST_ROOT = Join-Path ([IO.Path]::GetPathRoot($root)) 'RedSalamander.Perf'
$null = Initialize-RSTestSandboxRoot -TestRoot $env:REDSALAMANDER_TEST_ROOT -RepoRoot $root -AllowExternalInitialization
$env:TEMP = New-RSTestSandboxScratchDirectory -RepoRoot $root -Harness UnifiedCI -Case "$Phase-$Platform"
$env:TMP = $env:TEMP
$receiptDirectory = Join-Path $env:REDSALAMANDER_TEST_ROOT "evidence/unified-ci/$($env:GITHUB_RUN_ID)/$Platform"
$evidenceDirectory = Join-Path $env:REDSALAMANDER_TEST_ROOT 'evidence'
if ($env:GITHUB_OUTPUT) { Add-Content -LiteralPath $env:GITHUB_OUTPUT -Value ('evidence_directory=' + $evidenceDirectory) }
if ($Phase -eq 'Production') {
    $env:RSBuildEnableTests = 'false'
    Run 'build.ps1' ($profile + @('-BuildNumber',$number,'-OfficialRelease','-ProductionOnly'))
    Test-PinnedDxUiMatrix
    # Reuse the test-stage smoke harness without recompiling it in production.
    $testReceipt = Get-Content (Join-Path $root ('.build/ci/test-tools/regression-' + $Platform + '.json')) -Raw | ConvertFrom-Json -AsHashtable
    $null = Test-CiRegressionReceipt $testReceipt $context.batch $Platform
    $harness = Join-Path $root '.build/ci/test-tools/PluginContractTests.exe'
    $expected = @($testReceipt.artifacts | Where-Object { $_.path -ceq ".build/$Platform/Release/PluginContractTests.exe" })
    if ($expected.Count -ne 1 -or (Get-FileHash -LiteralPath $harness).Hash.ToLowerInvariant() -cne $expected[0].sha256) { throw 'Smoke harness does not match the tested artifact.' }
    Copy-Item -LiteralPath $harness -Destination (Join-Path $root ".build/$Platform/Release/PluginContractTests.exe")
    Run 'Installer/zip/build-zip.ps1' ($profile + @('-BuildNumber',$number))
    $archives = @(Get-ChildItem (Join-Path $root '.build/AppPackages') -Filter "*-$Platform-Portable.zip" -File | ForEach-Object FullName)
    Write-CiProductionReceipt -Context $context -Archives $archives -TestModeDisabled $true -SmokePassed $true
    return
}
$env:RSBuildEnableTests = 'true'
if ($QualificationScope -ceq 'R0ePaired') { $env:RSBuildCompilerDebugInformation = 'Embedded' }
if ($Phase -in @('ReleaseTestBuild','Qualification')) {
    Run 'build.ps1' ($profile + @('-BuildNumber',$number))
    Test-PinnedDxUiMatrix
    if ($Phase -eq 'ReleaseTestBuild') { return }
}
Import-Module (Join-Path $root 'Tools/Modules/Testing/HostedTestDesktop.psm1') -Force
if (-not (Get-Module -ListAvailable Pester | Where-Object Version -EQ '3.4.0')) {
    Install-Module Pester -RequiredVersion 3.4.0 -Scope CurrentUser -Force -SkipPublisherCheck
}
if ($Configuration -eq 'ASan Debug') {
    $seed = Invoke-CiProcess (Join-Path $root ".build/$Platform/$Configuration/PluginContractTests.exe") @('--asan-seed-heap-overflow') $env:TEMP -TimeoutSeconds 60 -AllowFailure
    [void][IO.Directory]::CreateDirectory($receiptDirectory)
    [IO.File]::WriteAllText((Join-Path $receiptDirectory 'asan-detector-proof.log'),$seed.stdout+$seed.stderr)
    if ($seed.exitCode -eq 0 -or ($seed.stdout+$seed.stderr) -notmatch 'AddressSanitizer') {
        throw 'The seeded heap overflow was not rejected with an AddressSanitizer diagnostic.'
    }
}
Invoke-RSWithHostedTestDesktop -Action {
    if ($QualificationScope -ceq 'R0ePaired') {
        [void][IO.Directory]::CreateDirectory($receiptDirectory)
        $command = Assert-CiContainedPath $root (Join-Path $root 'Tools/Invoke-R0ePairedValidation.ps1')
        $hostPath = (Get-Process -Id $PID).Path
        $pairArgs = @('-NoProfile','-NonInteractive','-File',$command,'-BaselineHead',$pairedDispatch.baseline,
            '-CandidateHead',$pairedDispatch.candidate,'-Platform',$Platform,'-BuildNumber',$number,
            '-TestRoot',$env:REDSALAMANDER_TEST_ROOT,'-EngineRoot',$EngineRoot)
        $process = Invoke-CiProcess -FilePath $hostPath -Arguments $pairArgs -WorkingDirectory $root -TimeoutSeconds 18000 -Stream -AllowFailure
        [IO.File]::WriteAllText((Join-Path $receiptDirectory 'r0e-paired.stdout.log'), $process.stdout)
        [IO.File]::WriteAllText((Join-Path $receiptDirectory 'r0e-paired.stderr.log'), $process.stderr)
        if ($process.exitCode -eq 2) { throw 'R0e paired diagnostic retained a timing-only RED baseline; qualification remains NOT_EVALUATED.' }
        if ($process.exitCode -ne 0) { throw "R0e paired diagnostic failed with exit $($process.exitCode); preserve its evidence." }
        [void](Invoke-CiAssertionProbe -Context $context -Executable ".build/$Platform/$Configuration/PluginContractTests.exe" -ExpectedExitCode 21)
        Write-Host 'R0e paired diagnostic completed; repository coverage remains NOT_EVALUATED.'
        return
    }
    if ($QualificationScope -in @('CommandsHandoff','FileOpsVolumes','ViewerSpaceCapture')) {
        Import-Module (Join-Path $root 'Tools/Modules/Testing/HostedDiagnosticEvidence.psm1') -Force
        Import-Module (Join-Path $root 'Tools/Modules/Build/BuildEvidence.psm1')
        $selection = Get-RSHostedDiagnosticSelection -Scope $QualificationScope
        [void][IO.Directory]::CreateDirectory($receiptDirectory)
        $diagnosticMarker = [ordered]@{scope=$QualificationScope; suite=$selection.suite; case=$selection.case; run_id=$null;
            platform=$Platform; configuration=$Configuration; source_snapshot_id=$null; build_receipt_id=$null;
            native_result_sha256=$null; runner_summary_sha256=$null; archive_intent_sha256=$null;
            assertion_probe_passed=$false; status='incomplete'; repository_verdict='NOT_EVALUATED'; coverage='filtered'}
        try {
            $buildOutput = Join-Path $root ".build/$Platform/$Configuration"
            $receipt = Read-RSBuildReceipt -BuildOutputDir $buildOutput -SchemaPath (Join-Path $root 'Specs/Build/BuildReceipt.schema.json')
            $receiptHash = (Get-FileHash -LiteralPath (Join-Path $buildOutput 'build-receipt.json') -Algorithm SHA256).Hash
            $snapshot = Get-RSBuildSourceSnapshot -RepoRoot $root
            $requiredExe = ".build/$Platform/$Configuration/RedSalamander.exe"
            $receiptVerified = Test-RSBuildReceiptForArtifactUse -Receipt $receipt -SourceSnapshot $snapshot -Platform $Platform -Configuration $Configuration -RepoRoot $root -RequiredRelativePath @($requiredExe)
            if (-not $receiptVerified) { throw 'Diagnostic full-solution source/artifact receipt is invalid.' }
            $runner = Assert-CiContainedPath $root (Join-Path $root 'Tools/Run-AllTests.ps1')
            $hostPath = (Get-Process -Id $PID).Path
            $runnerArgs = @('-NoProfile','-NonInteractive','-File',$runner) + $profile + @('-Suite',$selection.suite,'-CaseFilter',$selection.case,
                '-ValidationMode','Fresh','-SkipBuild','-BuildNumber',$number,'-TestRoot',$env:REDSALAMANDER_TEST_ROOT,'-TimeoutMultiplier','2.0','-AllowAlternateVolumeTestRoot')
            $invocationStartedUtc = [DateTimeOffset]::UtcNow
            $process = Invoke-CiProcess -FilePath $hostPath -Arguments $runnerArgs -WorkingDirectory $root -TimeoutSeconds 18000 -Stream -AllowFailure
            $invocationEndedUtc = [DateTimeOffset]::UtcNow
            [IO.File]::WriteAllText((Join-Path $receiptDirectory 'diagnostic-runner.stdout.log'), $process.stdout)
            [IO.File]::WriteAllText((Join-Path $receiptDirectory 'diagnostic-runner.stderr.log'), $process.stderr)
            if ($process.exitCode -ne 2) { throw "Filtered diagnostic runner returned unexpected exit $($process.exitCode)." }
            $runId = Get-RSHostedDiagnosticRunId -Output $process.stdout
            $diagnosticMarker.run_id = $runId
            $evidence = Read-RSHostedDiagnosticEvidence -TestRoot $env:REDSALAMANDER_TEST_ROOT -RunId $runId -Selection $selection
            $after = Get-RSBuildSourceSnapshot -RepoRoot $root
            $afterReceipt = Read-RSBuildReceipt -BuildOutputDir $buildOutput -SchemaPath (Join-Path $root 'Specs/Build/BuildReceipt.schema.json')
            $stillVerified = $after.SnapshotId -ceq $snapshot.SnapshotId -and $null -ne $afterReceipt -and
                $receipt.receipt_id -ceq $afterReceipt.receipt_id -and
                $receiptHash -ceq (Get-FileHash -LiteralPath (Join-Path $buildOutput 'build-receipt.json') -Algorithm SHA256).Hash -and
                (Test-RSBuildReceiptForArtifactUse -Receipt $afterReceipt -SourceSnapshot $after -Platform $Platform -Configuration $Configuration -RepoRoot $root -RequiredRelativePath @($requiredExe))
            [void](Test-RSHostedDiagnosticOutcome -ExitCode $process.exitCode -RunId $runId -Selection $selection -Summary $evidence.summary -Native $evidence.native -Platform $Platform -Configuration $Configuration -ReceiptVerified $stillVerified -ExpectedExePath ([IO.Path]::GetFullPath((Join-Path $buildOutput 'RedSalamander.exe'))) -InvocationStartedUtc $invocationStartedUtc -InvocationEndedUtc $invocationEndedUtc)
            $diagnosticMarker.source_snapshot_id = $snapshot.SnapshotId
            $diagnosticMarker.build_receipt_id = $receipt.receipt_id
            $diagnosticMarker.native_result_sha256 = $evidence.nativeSha256
            $diagnosticMarker.runner_summary_sha256 = $evidence.summarySha256
            $diagnosticMarker.archive_intent_sha256 = $evidence.archiveIntentSha256
            $diagnosticMarker.assertion_probe_passed = Invoke-CiAssertionProbe -Context $context -Executable ".build/$Platform/$Configuration/PluginContractTests.exe" -ExpectedExitCode 21
            $diagnosticMarker.status = 'passed'
            Write-Host "Filtered diagnostic $QualificationScope passed its selected case; repository coverage remains NOT_EVALUATED."
        } catch {
            $diagnosticMarker.status = 'failed'
            throw
        } finally {
            $diagnosticMarker | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $receiptDirectory 'diagnostic-coverage.json') -Encoding utf8
        }
        return
    }
    $suite = if ($Phase -eq 'Qualification' -and $QualificationScope -eq 'Full') { 'Full' } else { 'CI' }
    Run 'Tools/Run-AllTests.ps1' ($profile + @('-Suite',$suite,'-ValidationMode','Fresh','-SkipBuild','-BuildNumber',$number,'-TestRoot',$env:REDSALAMANDER_TEST_ROOT,'-TimeoutMultiplier','2.0','-AllowAlternateVolumeTestRoot'))
}
if ($QualificationScope -in @('CommandsHandoff','FileOpsVolumes','ViewerSpaceCapture','R0ePaired')) { return }
$probe = Invoke-CiAssertionProbe -Context $context -Executable ".build/$Platform/$Configuration/PluginContractTests.exe" -ExpectedExitCode 21
if ($Phase -eq 'Qualification') { Write-Host 'Explicit native qualification finished; it does not authorize production or publication.'; return }
$summaries = @(Get-ChildItem (Join-Path $env:REDSALAMANDER_TEST_ROOT 'evidence/completed-runs') -Recurse -Filter 'run-all-tests-results.json' -File | Where-Object LastWriteTimeUtc -ge $started)
if ($summaries.Count -ne 1) { throw 'Expected exactly one fresh completed CI-suite summary.' }
$summary = Get-Content $summaries[0].FullName -Raw | ConvertFrom-Json -AsHashtable
if ($summary.schema -cne 'red-salamander.run-all-tests.v2' -or $summary.repository_verdict -cne 'PASSED' -or
    $summary.case_counts.total -le 0 -or $summary.case_counts.failed -ne 0 -or
    $summary.counts.incomplete -ne 0 -or $summary.counts.invalidated -ne 0) { throw 'CI regression case evidence is incomplete.' }
$evidence = @{ countUnit='required-case';registered=[int]$summary.case_counts.total;expectedSkipped=[int]$summary.case_counts.skipped;
    discovered = [int]$summary.case_counts.total-[int]$summary.case_counts.skipped; executed = [int]$summary.case_counts.passed; evidence = (Get-CiEvidenceReference @($summaries[0].FullName)) }
$binaries = Get-CiProfileArtifacts -Context $context
if ($env:GITHUB_OUTPUT) { Add-Content -LiteralPath $env:GITHUB_OUTPUT -Value ('smoke_harness=' + (Join-Path $root ".build/$Platform/Release/PluginContractTests.exe")) }


if ($Phase -eq 'Qualification') { Write-Host 'Explicit qualification finished; it does not authorize production or publication.'; return }
$countUnit=if($evidence.Contains('countUnit')){$evidence.countUnit}else{'case'}
$registered=if($evidence.Contains('registered')){[int]$evidence.registered}else{[int]$evidence.discovered}
$expectedSkipped=if($evidence.Contains('expectedSkipped')){[int]$evidence.expectedSkipped}else{0}
Write-CiRegressionReceipt -Context $context -Discovered $evidence.discovered -Executed $evidence.executed -AssertionProbePassed $probe -Artifacts $binaries -EvidenceDirectory $receiptDirectory -CanonicalEvidence $evidence.evidence -CountUnit $countUnit -Registered $registered -ExpectedSkipped $expectedSkipped
if ($env:GITHUB_OUTPUT) { Add-Content -LiteralPath $env:GITHUB_OUTPUT -Value ('evidence_directory=' + $evidenceDirectory) }
