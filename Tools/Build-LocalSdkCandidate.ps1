#Requires -Version 7.4
<#
.SYNOPSIS Builds the application and plugin closure against an explicit local SDK in isolated outputs.
.DESCRIPTION Uses the pinned shared broker's source, API, actual ABI, immutable reader, live output lease and linker-input checks. Resolves the native solution target through the canonical solution mapper. Does not run the application or tests.
.PARAMETER Descriptor Absolute producer local SDK snapshot descriptor.
.PARAMETER DescriptorSha256 Independently retained SHA-256 of the descriptor.
.PARAMETER MSBuildPath Explicit application-selected native Visual Studio MSBuild.exe; the broker verifies its actual toolchain and consumer ABI.
.PARAMETER Platform x64 or ARM64 Release target. Cross-compilation does not establish native execution.
.PARAMETER BuildNumber Explicit positive resource build number; zero would invoke the ordinary mutable version allocator.
.OUTPUTS Fresh .build/local-unpinned operation, build log, actual link proof and local-only build record.
.NOTES Requires local Windows, the exact clean pinned engine, matching API/ABI and existing ordinary platform dependencies, including prepared terminal inputs. Run ordinary build/dependency preparation first. Holds the existing platform dependency lease and refuses contamination; never repairs dependencies, changes managed pins or overwrites ordinary outputs. Every project checks its output lease before writes; actual DxUi consumers retain full ABI checks. Tests are compiled but not executed. Grants no repository, package, production or native-platform qualification.
.EXAMPLE ./Tools/Build-LocalSdkCandidate.ps1 -Descriptor Z:/src/DxUi/.build/local-sdk-capture/example/local-unpinned-snapshot.json -DescriptorSha256 aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa -MSBuildPath 'C:/BuildTools/MSBuild/Current/Bin/amd64/MSBuild.exe' -BuildNumber 123
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$Descriptor,
    [Parameter(Mandatory)][ValidatePattern('^[a-f0-9]{64}$')][string]$DescriptorSha256,
    [Parameter(Mandatory)][string]$MSBuildPath,
    [ValidateSet('x64','ARM64')][string]$Platform='x64',[Parameter(Mandatory)][ValidateRange(1,65535)][int]$BuildNumber)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(-not $IsWindows -or $env:GITHUB_ACTIONS -ceq 'true' -or $env:CI_NATIVE_BUILDS_DISABLED -ceq 'true'){throw 'Local SDK candidate builds require local Windows native execution.'}
if(-not [IO.Path]::IsPathFullyQualified($MSBuildPath) -or -not [IO.File]::Exists($MSBuildPath) -or [IO.Path]::GetFileName($MSBuildPath) -ine 'MSBuild.exe'){throw 'Select an existing absolute native Visual Studio MSBuild.exe.'}
$root=Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $root 'Tools/Modules/Build/ArtifactOperationLock.psm1')
Import-Module (Join-Path $root 'Tools/Modules/Build/BuildProjectSelection.psm1')
Import-Module (Join-Path $root '.ci-engine/Tools/Modules/Ci.LocalSdkBuild.psm1') -ErrorAction Stop
$selection=Get-RSBuildProjectUnion -SolutionPath (Join-Path $root 'RedSalamander.sln') -ProjectNames RedSalamander
$buildContract=[ordered]@{buildFile='RedSalamander.sln';metadataProject='RedSalamander/RedSalamander.vcxproj';target=$selection.MSBuildTarget;sdkProperty='RSDxUiSdkRoot';engineProperty='RSCiEngineRoot';outputProperty='RSOutputRoot';testsProperty='RSBuildEnableTests';versionProperty='RSVersionBuildNumber';executable='RedSalamander.exe'}
$scope=@{configuration='Release';platform=$Platform;coordination='shared-dependency'}
$lease=Enter-RSArtifactOperationLock -RepoRoot $root -Operation 'Local SDK dependency read' -Scope $scope
try{
    if($lease.WasAbandoned){
        [void](Set-RSArtifactOperationContaminated -RepoRoot $root -Reason 'The previous dependency owner exited before local SDK input admission.' -AbandonedOwner $lease.AbandonedOwner -Scope $scope)
    }
    if(Read-RSArtifactOperationContamination -RepoRoot $root -Scope $scope){throw 'Local SDK build refuses abandoned or contaminated dependency inputs; repair them through the ordinary build.'}
    Assert-RSNoResidualArtifactToolProcesses -RepoRoot $root -Scope $scope
    $result=Invoke-CiLocalSdkApplicationBuild -ProjectRoot $root -Descriptor $Descriptor -ExpectedDescriptorSha256 $DescriptorSha256 -Platform $Platform -MSBuildPath $MSBuildPath -BuildContract $buildContract -BuildNumber $BuildNumber
}finally{Exit-RSArtifactOperationLock -Lock $lease}
Write-Host 'LOCAL_UNPINNED_BUILD: tests, repository and production qualification remain NOT_EVALUATED.'
$result
