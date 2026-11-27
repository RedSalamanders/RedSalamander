<#
.SYNOPSIS
    Generates WinGet manifests with the shared publication renderer.
.DESCRIPTION
    Hashes both portable packages and generates the reviewed multi-file manifest
    from .ci/project.json using the checkout's pinned CI engine.
.PARAMETER Version
    Package version (e.g., 7.0.183). If not provided, derives it from the shared version helper.
.PARAMETER BuildNumber
    Positive compiled package build number used when Version is omitted. One of
    Version or BuildNumber is required.
.PARAMETER ZipPath
    Path to the x64 ZIP installer. Its leaf name must be exactly
    RedSalamander-<Version>-x64-Portable.zip.
.PARAMETER Arm64ZipPath
    Path to the ARM64 ZIP installer. Its leaf name must be exactly
    RedSalamander-<Version>-ARM64-Portable.zip.
.PARAMETER OutputDir
    Output directory within the selected package root. The compatibility caller
    defaults to .build\AppPackages\winget-manifest. The canonical caller uses a version subfolder.
.PARAMETER CanonicalPackageDirectory
    Use the canonical .build/packages root and discover either package layout when ZIP paths are omitted.
.PARAMETER ValidateManifest
    Run winget validate on the staged manifest before publication.
.PARAMETER ReleaseDate
    Manifest ReleaseDate as yyyy-MM-dd. The release workflow passes the GitHub
    release's published date so reruns generate identical manifests. Default: today.
.EXAMPLE
    .\Installer\winget\generate-manifest.ps1 -Version 7.0.183 -ZipPath .\.build\AppPackages\RedSalamander-7.0.183-x64-Portable.zip -Arm64ZipPath .\.build\AppPackages\RedSalamander-7.0.183-ARM64-Portable.zip
.OUTPUTS
    None. Writes the three Winget manifest files to OutputDir.
.NOTES
    Mutates shared packaging outputs while holding the repository-wide packaging
    coordination lock. Interrupted packaging fails closed until a reviewer restores
    Installer inputs and package roots and removes the reported marker.
#>
[CmdletBinding()]
param(
    [string]$Version,
    [int]$BuildNumber = 0,
    [string]$ZipPath,
    [string]$Arm64ZipPath,
    [string]$OutputDir = '',
    [string]$ReleaseDate = '',
    [switch]$CanonicalPackageDirectory,
    [switch]$ValidateManifest
)

$ErrorActionPreference = 'Stop'

function Resolve-WingetReleaseDate {
    param(
        [AllowEmptyString()]
        [string]$CandidateDate
    )

    if ([string]::IsNullOrWhiteSpace($CandidateDate)) {
        return (Get-Date -Format 'yyyy-MM-dd')
    }

    $parsed = [datetime]::MinValue
    if ($CandidateDate -notmatch '^\d{4}-\d{2}-\d{2}$' -or
        -not [datetime]::TryParseExact(
            $CandidateDate,
            'yyyy-MM-dd',
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None,
            [ref]$parsed)) {
        throw "Winget ReleaseDate must be a calendar date formatted yyyy-MM-dd: $CandidateDate"
    }

    return $CandidateDate
}

function Assert-WingetManifestVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CandidateVersion
    )

    $match = [regex]::Match($CandidateVersion, '^(?<major>\d+)\.(?<minor>\d+)\.(?<build>\d+)$')
    if (-not $match.Success) {
        throw "Winget Version must contain exactly three numeric components: $CandidateVersion"
    }

    $components = @{}
    foreach ($componentName in @('major', 'minor', 'build')) {
        $componentValue = 0
        if (-not [int]::TryParse($match.Groups[$componentName].Value, [ref]$componentValue) -or
            $componentValue -lt 0 -or
            $componentValue -gt 65535) {
            throw "Winget Version components must be in the range 0..65535: $CandidateVersion"
        }
        $components[$componentName] = $componentValue
    }

    if ($components.build -le 0) {
        throw "Winget Version build component must be positive: $CandidateVersion"
    }
}

function Resolve-WingetPortableZip {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CandidatePath,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedLeaf,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($CandidatePath)) {
        throw "$Label ZIP path is required. Expected leaf: $ExpectedLeaf"
    }

    $normalizedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CandidatePath)
    if (-not (Test-Path -LiteralPath $normalizedPath -PathType Leaf)) {
        throw "$Label ZIP must be an existing file: $normalizedPath"
    }

    $item = Get-Item -LiteralPath $normalizedPath -Force
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "$Label ZIP must not be a reparse point: $normalizedPath"
    }
    if (-not [string]::Equals(
            $item.Name,
            $ExpectedLeaf,
            [System.StringComparison]::Ordinal)) {
        throw "$Label ZIP leaf must be exactly '$ExpectedLeaf': $($item.Name)"
    }

    return $item.FullName
}

function Get-WingetPortableZipPath {
    param([string]$RepositoryRoot, [string]$LeafName)

    foreach ($directory in @('.build\packages', '.build\AppPackages')) {
        $candidate = Resolve-RSGuardedRepositoryPath -RepoRoot $RepositoryRoot -GuardedRelativeRoot $directory `
            -Path (Join-Path $directory $LeafName)
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    throw "No package named '$LeafName' exists in .build/packages or .build/AppPackages."
}

if ([string]::IsNullOrWhiteSpace($Version) -and ($BuildNumber -le 0 -or $BuildNumber -gt 65535)) {
    throw 'Winget manifest generation requires either -Version or -BuildNumber in the range 1..65535; it never allocates or guesses a package version.'
}

$RepoRoot = Split-Path -Parent $PSScriptRoot | Split-Path -Parent
$VersioningModule = Join-Path $RepoRoot "Tools\Modules\Build\Versioning.psm1"
$ArtifactOperationLockModule = Join-Path $RepoRoot "Tools\Modules\Build\ArtifactOperationLock.psm1"
$PackageRelativeRoot = if ($CanonicalPackageDirectory) { '.build\packages' } else { '.build\AppPackages' }
if (-not (Test-Path $ArtifactOperationLockModule)) {
    throw "Artifact-operation lock module not found: $ArtifactOperationLockModule"
}

Import-Module $ArtifactOperationLockModule -Force -ErrorAction Stop
$PackagingScope = @{
    kind = 'packaging'
    target = 'winget-manifest'
    configuration = 'Release'
    platform = 'x64'
    coordination = 'packaging'
}
$PackagingLock = $null
try {
    $PackagingLock = Enter-RSArtifactOperationLock `
        -RepoRoot $RepoRoot `
        -Operation 'standalone Winget manifest generation Release|multiarch' `
        -Scope $PackagingScope
    if ($PackagingLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $RepoRoot `
                -Reason 'The previous packaging owner exited while mutating shared installer inputs or package outputs.' `
                -AbandonedOwner $PackagingLock.AbandonedOwner `
                -Scope $PackagingScope)
    }
    $PackagingContamination = Read-RSArtifactOperationContamination `
        -RepoRoot $RepoRoot `
        -Scope $PackagingScope
    if ($null -ne $PackagingContamination) {
        $MarkerPath = Get-RSArtifactContaminationMarkerPath `
            -RepoRoot $RepoRoot `
            -Scope $PackagingScope
        throw "Shared packaging state may be incomplete after an interrupted operation. Review and restore Installer inputs and the package roots .build\packages and .build\AppPackages, then remove the marker explicitly. Marker: $MarkerPath"
    }

# Read version from the shared version helper if not provided
if (-not $Version) {
    if (-not (Test-Path $VersioningModule)) {
        throw "Version helper module not found: $VersioningModule"
    }

    Import-Module $VersioningModule -Force -ErrorAction Stop
    $VersionContext = New-RSVersionContext `
        -RepoRoot $RepoRoot `
        -Configuration Release `
        -Platform x64 `
        -BuildNumber $BuildNumber
    $Version = $VersionContext.PackagingVersion
}

Assert-WingetManifestVersion -CandidateVersion $Version
$ReleaseDate = Resolve-WingetReleaseDate -CandidateDate $ReleaseDate
if ($CanonicalPackageDirectory) {
    if (-not $ZipPath) {
        $ZipPath = Get-WingetPortableZipPath -RepositoryRoot $RepoRoot -LeafName "RedSalamander-$Version-x64-Portable.zip"
    }
    if (-not $Arm64ZipPath) {
        $Arm64ZipPath = Get-WingetPortableZipPath -RepositoryRoot $RepoRoot -LeafName "RedSalamander-$Version-ARM64-Portable.zip"
    }
}
if (-not $OutputDir) {
    $OutputDir = if ($CanonicalPackageDirectory) {
        Join-Path $PackageRelativeRoot "winget-manifest\$Version"
    } else {
        Join-Path $PackageRelativeRoot 'winget-manifest'
    }
}
$ZipPath = Resolve-WingetPortableZip `
    -CandidatePath $ZipPath `
    -ExpectedLeaf "RedSalamander-$Version-x64-Portable.zip" `
    -Label 'x64'
$Arm64ZipPath = Resolve-WingetPortableZip `
    -CandidatePath $Arm64ZipPath `
    -ExpectedLeaf "RedSalamander-$Version-ARM64-Portable.zip" `
    -Label 'ARM64'
$OutputDir = Resolve-RSGuardedRepositoryPath `
    -RepoRoot $RepoRoot `
    -GuardedRelativeRoot $PackageRelativeRoot `
    -Path $OutputDir

Write-Host "Generating winget manifest for version $Version..." -ForegroundColor Cyan

$binding = Join-Path $RepoRoot '.ci-engine/Tools/Modules/Ci.EngineBinding.psm1'
if (-not (Test-Path -LiteralPath $binding -PathType Leaf)) {
    throw 'WinGet manifest generation requires the .ci-engine checkout pinned by .ci/project.json.'
}
Import-Module $binding -ErrorAction Stop
$publication = Resolve-CiPinnedEngineModule -ProjectRoot $RepoRoot -Name 'Publication.psm1'
Import-Module (Join-Path (Split-Path -Parent $publication) 'Ci.Policy.psm1') -ErrorAction Stop
Import-Module $publication -ErrorAction Stop

$project = Read-CiProject $RepoRoot
if ($project.repository -cne 'RedSalamanders/RedSalamander') {
    throw 'WinGet manifest generation requires the RedSalamander project manifest.'
}

$plan = @{
    project = $project
    batch = @{ version = $Version }
    tag = "v$Version"
    assets = @(
        @{ platform = 'x64'; file = [IO.Path]::GetFileName($ZipPath); path = $ZipPath; sha256 = (Get-FileHash -LiteralPath $ZipPath -Algorithm SHA256).Hash }
        @{ platform = 'ARM64'; file = [IO.Path]::GetFileName($Arm64ZipPath); path = $Arm64ZipPath; sha256 = (Get-FileHash -LiteralPath $Arm64ZipPath -Algorithm SHA256).Hash }
    )
}
$StageDir = Resolve-RSGuardedRepositoryPath -RepoRoot $RepoRoot -GuardedRelativeRoot $PackageRelativeRoot `
    -Path (Join-Path $PackageRelativeRoot ('.winget-' + [Guid]::NewGuid().ToString('N'))) -CreateDirectory
try {
    $null = New-CiWingetManifest -Plan $plan -Directory $StageDir -ReleaseDate $ReleaseDate
    if ($ValidateManifest) {
        . (Join-Path $PSScriptRoot 'WingetValidation.ps1')
        Invoke-RSWingetManifestValidation -ManifestPath $StageDir
    }

    $OutputDir = Resolve-RSGuardedRepositoryPath -RepoRoot $RepoRoot -GuardedRelativeRoot $PackageRelativeRoot `
        -Path $OutputDir -CreateDirectory
    $destinations = @(foreach ($suffix in '.yaml', '.installer.yaml', '.locale.en-US.yaml') {
        $candidate = Resolve-RSGuardedRepositoryPath -RepoRoot $RepoRoot -GuardedRelativeRoot $PackageRelativeRoot `
            -Path (Join-Path $OutputDir "$($project.winget.identifier)$suffix")
        if (Test-Path -LiteralPath $candidate -PathType Container) {
            throw "A WinGet manifest destination is a directory: $candidate"
        }
        $candidate
    })
    foreach ($destination in $destinations) {
        $name = [IO.Path]::GetFileName($destination)
        $temporary = Resolve-RSGuardedRepositoryPath -RepoRoot $RepoRoot -GuardedRelativeRoot $PackageRelativeRoot `
            -Path (Join-Path $OutputDir ('.' + $name + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'))
        try {
            [IO.File]::Copy((Join-Path $StageDir $name), $temporary)
            $null = Publish-RSGuardedRepositoryFile -RepoRoot $RepoRoot -GuardedRelativeRoot $PackageRelativeRoot `
                -StagedPath $temporary -DestinationPath $destination
        } finally {
            if (Test-Path -LiteralPath $temporary -PathType Leaf) {
                Remove-Item -LiteralPath $temporary -Force
            }
        }
    }
} finally {
    $cleanup = Resolve-RSGuardedRepositoryPath -RepoRoot $RepoRoot -GuardedRelativeRoot $PackageRelativeRoot -Path $StageDir
    Remove-Item -LiteralPath $cleanup -Recurse -Force
}

Write-Host "WinGet manifests published to $OutputDir" -ForegroundColor Green
}
finally {
    Exit-RSArtifactOperationLock -Lock $PackagingLock
}
