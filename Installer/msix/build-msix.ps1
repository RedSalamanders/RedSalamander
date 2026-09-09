<#
.SYNOPSIS
    Builds one unsigned RedSalamander MSIX package from a compiled Release profile.
.DESCRIPTION
    Regenerates the reviewed MSIX assets, stages a package manifest, and invokes
    the Windows Application Packaging Project while one artifact-profile lease and
    one repository-wide packaging lease cover the complete transaction.
.PARAMETER Version
    Three-part package version, for example 7.0.184.
.PARAMETER Configuration
    Build configuration. Only Release is supported.
.PARAMETER Platform
    Target platform, x64 or ARM64.
.PARAMETER MSBuildPath
    Optional exact MSBuild executable path. Defaults to MSBUILD_EXE_PATH or the
    first MSBuild.exe resolved from PATH.
.PARAMETER TargetPlatformVersion
    Optional installed four-part Windows SDK version. The newest installed UAP
    SDK is selected when omitted.
.EXAMPLE
    .\Installer\msix\build-msix.ps1 -Version 7.0.184 -Platform x64
.OUTPUTS
    None. MSBuild writes the unsigned package beneath .build\AppPackages.
.NOTES
    Mutates package outputs but never the tracked package manifest. Interrupted
    packaging may leave an unused unique staged manifest and fails closed for
    shared AppPackages outputs until the contamination marker is reviewed.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version,

    [ValidateSet('Release')]
    [string]$Configuration = 'Release',

    [Parameter(Mandatory = $true)]
    [ValidateSet('x64', 'ARM64')]
    [string]$Platform,

    [string]$MSBuildPath,

    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$TargetPlatformVersion
)

$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$artifactOperationLockModule = Join-Path $repoRoot 'Tools\Modules\Build\ArtifactOperationLock.psm1'
$sanitizedEnvironmentModule = Join-Path $repoRoot 'Tools\Modules\Build\SanitizedEnvironment.psm1'
$generateAssetsScript = Join-Path $PSScriptRoot 'GenerateAssets.ps1'
$updateManifestScript = Join-Path $PSScriptRoot 'UpdateManifestVersion.ps1'
$installerProject = Join-Path $PSScriptRoot 'RedSalamanderInstaller.wapproj'

foreach ($requiredPath in @(
        $artifactOperationLockModule,
        $sanitizedEnvironmentModule,
        $generateAssetsScript,
        $updateManifestScript,
        $installerProject)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Required MSIX packaging input not found: $requiredPath"
    }
}

Import-Module $artifactOperationLockModule -Force -ErrorAction Stop
Import-Module $sanitizedEnvironmentModule -Force -ErrorAction Stop

$resolvedMSBuildPath = if (-not [string]::IsNullOrWhiteSpace($MSBuildPath)) {
    [IO.Path]::GetFullPath($MSBuildPath)
}
elseif (-not [string]::IsNullOrWhiteSpace($env:MSBUILD_EXE_PATH)) {
    [IO.Path]::GetFullPath($env:MSBUILD_EXE_PATH)
}
else {
    $command = Get-Command 'MSBuild.exe' -ErrorAction Stop
    [IO.Path]::GetFullPath($command.Source)
}
if (-not (Test-Path -LiteralPath $resolvedMSBuildPath -PathType Leaf)) {
    throw "MSBuild executable not found: $resolvedMSBuildPath"
}

if ([string]::IsNullOrWhiteSpace($TargetPlatformVersion)) {
    $uapRoot = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\DesignTime\CommonConfiguration\Neutral\UAP'
    $installedSdk = Get-ChildItem -LiteralPath $uapRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
        Sort-Object { [version]$_.Name } -Descending |
        Select-Object -First 1
    if ($null -eq $installedSdk) {
        throw "Windows SDK with UAP props not found beneath: $uapRoot"
    }
    $TargetPlatformVersion = $installedSdk.Name
}

$artifactProfileScope = @{
    kind = 'packaging-input'
    target = 'msix'
    configuration = $Configuration
    platform = $Platform
}
$packagingScope = @{
    kind = 'packaging'
    target = 'msix'
    configuration = $Configuration
    platform = $Platform
    coordination = 'packaging'
}
$sharedDependencyScope = @{
    kind = 'packaging-dependency'
    target = 'msix'
    configuration = $Configuration
    platform = $Platform
    coordination = 'shared-dependency'
}
$artifactProfileLock = $null
$packagingLock = $null
$sharedDependencyLock = $null
$stagedManifestPath = $null
try {
    $artifactProfileLock = Enter-RSArtifactOperationLock `
        -RepoRoot $repoRoot `
        -Operation "standalone MSIX input $Configuration|$Platform" `
        -Scope $artifactProfileScope
    if ($artifactProfileLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $repoRoot `
                -Reason 'The previous build or packaging reader exited while using this compiled artifact profile.' `
                -AbandonedOwner $artifactProfileLock.AbandonedOwner `
                -Scope $artifactProfileScope)
    }
    if (Test-RSArtifactOperationContaminated -RepoRoot $repoRoot -Scope $artifactProfileScope) {
        $profileMarkerPath = Get-RSArtifactContaminationMarkerPath `
            -RepoRoot $repoRoot `
            -Scope $artifactProfileScope
        throw "Compiled packaging inputs may be incomplete after an interrupted operation. Run a full-solution build.ps1 -Rebuild for $Configuration|$Platform before packaging. Marker: $profileMarkerPath"
    }
    Assert-RSNoResidualArtifactToolProcesses `
        -RepoRoot $repoRoot `
        -Scope $artifactProfileScope

    $packagingLock = Enter-RSArtifactOperationLock `
        -RepoRoot $repoRoot `
        -Operation "standalone MSIX packaging $Configuration|$Platform" `
        -Scope $packagingScope
    if ($packagingLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $repoRoot `
                -Reason 'The previous packaging owner exited while mutating shared installer inputs or AppPackages outputs.' `
                -AbandonedOwner $packagingLock.AbandonedOwner `
                -Scope $packagingScope)
    }
    $packagingContamination = Read-RSArtifactOperationContamination `
        -RepoRoot $repoRoot `
        -Scope $packagingScope
    if ($null -ne $packagingContamination) {
        $markerPath = Get-RSArtifactContaminationMarkerPath `
            -RepoRoot $repoRoot `
            -Scope $packagingScope
        throw "Shared packaging state may be incomplete after an interrupted operation. Review and restore Installer inputs and .build\AppPackages, then remove the marker explicitly. Marker: $markerPath"
    }

    & $generateAssetsScript -Configuration $Configuration -Platform $Platform
    $stagedManifestPath = & $updateManifestScript `
        -Version $Version `
        -Platform $Platform `
        -Configuration $Configuration `
        -ManifestPath (Join-Path $PSScriptRoot 'Package.appxmanifest')

    $solutionDirWithSlash = $repoRoot.TrimEnd('\') + '\'
    $msixArguments = @(
        $installerProject,
        '/t:Build',
        "/p:Configuration=$Configuration",
        "/p:Platform=$Platform",
        "/p:TargetPlatformVersion=$TargetPlatformVersion",
        "/p:SolutionDir=$solutionDirWithSlash",
        "/p:RSAppxManifestPath=$stagedManifestPath",
        '/p:AppxPackageSigningEnabled=false',
        '/p:GenerateAppInstallerFile=false',
        '/p:AppxBundle=Never',
        '/v:minimal',
        '/nologo'
    )

    $sharedDependencyLock = Enter-RSArtifactOperationLock `
        -RepoRoot $repoRoot `
        -Operation "standalone MSIX shared dependency phase $Configuration|$Platform" `
        -Scope $sharedDependencyScope
    if ($sharedDependencyLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $repoRoot `
                -Reason 'The previous MSBuild owner exited while it could have been writing shared vcpkg dependencies.' `
                -AbandonedOwner $sharedDependencyLock.AbandonedOwner `
                -Scope $sharedDependencyScope)
    }
    $sharedDependencyContamination = Read-RSArtifactOperationContamination `
        -RepoRoot $repoRoot `
        -Scope $sharedDependencyScope
    if ($null -ne $sharedDependencyContamination) {
        $sharedDependencyMarkerPath = Get-RSArtifactContaminationMarkerPath `
            -RepoRoot $repoRoot `
            -Scope $sharedDependencyScope
        throw "Shared vcpkg dependencies may be incomplete after an interrupted MSBuild operation. Run a full-solution build.ps1 -Rebuild for platform $Platform before packaging. Marker: $sharedDependencyMarkerPath"
    }
    Assert-RSNoResidualArtifactToolProcesses `
        -RepoRoot $repoRoot `
        -Scope $sharedDependencyScope

    $exitCode = Invoke-RSProcess `
        -FilePath $resolvedMSBuildPath `
        -Arguments $msixArguments `
        -WorkingDirectory $repoRoot
    if ($exitCode -ne 0) {
        throw "MSIX build failed with exit code $exitCode"
    }
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($stagedManifestPath)) {
        $stagedManifestDirectory = Resolve-RSGuardedRepositoryPath `
            -RepoRoot $repoRoot -GuardedRelativeRoot '.build\AppPackages' `
            -Path (Split-Path $stagedManifestPath -Parent)
        if (Test-Path -LiteralPath $stagedManifestDirectory -PathType Container) {
            Remove-Item -LiteralPath $stagedManifestDirectory -Recurse -Force
        }
    }
    Exit-RSArtifactOperationLock -Lock $sharedDependencyLock
    Exit-RSArtifactOperationLock -Lock $packagingLock
    Exit-RSArtifactOperationLock -Lock $artifactProfileLock
}
