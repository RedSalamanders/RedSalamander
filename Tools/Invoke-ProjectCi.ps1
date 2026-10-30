<#
.SYNOPSIS
Adapts the project commands to the pinned organization CI engine.
.PARAMETER Phase
Compiler-free PR checks, Release test build, regression execution, production ZIP, or explicit qualification.
.PARAMETER EngineRoot
Checkout of the exact shared engine commit in .ci/project.json.
.PARAMETER Platform
Matching native x64 or ARM64 hosted runner.
.PARAMETER Descriptor
Frozen main admission descriptor; required for native phases.
.PARAMETER Base
Immutable PR comparison base for changelog validation.
.PARAMETER Configuration
Debug, Release or ASan Debug for an explicit qualification only.
.PARAMETER QualificationScope
Normal regression or explicit Full qualification; never advances automatic Release state.
.PARAMETER DevelopmentEngine
Allows a local engine under development for compiler-free checks only.
.OUTPUTS
Canonical test evidence, source-bound receipts and qualified ZIPs; GitHub job output paths.
.NOTES
Requires PowerShell 7.4 and Git. Native work is hosted-only and uses the matching architecture.
Static never invokes compiler fixtures. No publication or source/pin updates.
See Specs/Build/Build_UnifiedCI.md for roles, input rules, evidence and manual actions.
.EXAMPLE
./Tools/Invoke-ProjectCi.ps1 -Phase Static -EngineRoot ../RedSalamanders.CI -Base origin/main -DevelopmentEngine
#>
[CmdletBinding()]
param([Parameter(Mandatory)][ValidateSet('Static','ReleaseTestBuild','ReleaseTest','Production','Qualification')][string]$Phase,
    [Parameter(Mandatory)][string]$EngineRoot, [ValidateSet('x64','ARM64')][string]$Platform = 'x64',
    [string]$Descriptor = '', [string]$Base = 'origin/main',
    [ValidateSet('Debug','Release','ASan Debug')][string]$Configuration = 'Release', [ValidateSet('Normal','Full')][string]$QualificationScope = 'Normal', [switch]$DevelopmentEngine)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$EngineRoot = [IO.Path]::GetFullPath($EngineRoot)
Import-Module (Join-Path $EngineRoot 'Tools/Modules/Ci.Policy.psm1') -ErrorAction Stop
Import-Module (Join-Path $EngineRoot 'Tools/Modules/Ci.Project.psm1') -ErrorAction Stop
$context = New-CiProjectContext -Root $root -EngineRoot $EngineRoot -Phase $Phase -Platform $Platform -Descriptor $Descriptor -DevelopmentEngine:$DevelopmentEngine
function Run([string]$Script, [string[]]$CommandArguments = @()) {
    Invoke-CiProjectScript -Context $context -Script $Script -Arguments $CommandArguments
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
    $suite = if ($Phase -eq 'Qualification' -and $QualificationScope -eq 'Full') { 'Full' } else { 'CI' }
    Run 'Tools/Run-AllTests.ps1' ($profile + @('-Suite',$suite,'-ValidationMode','Fresh','-SkipBuild','-BuildNumber',$number,'-TestRoot',$env:REDSALAMANDER_TEST_ROOT,'-TimeoutMultiplier','2.0','-AllowAlternateVolumeTestRoot'))
}
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
