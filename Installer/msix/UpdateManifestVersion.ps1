<#
.SYNOPSIS
    Updates the MSIX package identity version and architecture.
.DESCRIPTION
    Creates a uniquely staged MSIX manifest with the requested identity version
    and architecture. The tracked source manifest is read-only.
.PARAMETER Version
    Three-part package version, for example 7.0.184.
.PARAMETER Platform
    Target platform, x64 or ARM64.
.PARAMETER ManifestPath
    Source Package.appxmanifest. It must resolve beneath this repository's
    Installer\msix directory and is never modified.
.PARAMETER OutputManifestPath
    Optional unique output below .build\AppPackages. When omitted, the command
    creates a unique ManifestStaging directory.
.PARAMETER Configuration
    Build configuration associated with the packaging operation. Default: Release.
.EXAMPLE
    .\Installer\msix\UpdateManifestVersion.ps1 -Version 7.0.184 -Platform x64
.OUTPUTS
    The absolute path to the uniquely staged UTF-8 manifest.
.NOTES
    Writes only a unique ignored staging path while holding the repository-wide
    packaging coordination lock. An interruption may leave an unused staging
    directory, but cannot change the tracked manifest.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [ValidateSet("x64", "ARM64")]
    [string]$Platform,

    [string]$ManifestPath = (Join-Path $PSScriptRoot "Package.appxmanifest"),

    [string]$OutputManifestPath,

    [ValidateSet('Release')]
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\\..")).Path
$artifactOperationLockModule = Join-Path $repoRoot "Tools\Modules\Build\ArtifactOperationLock.psm1"
if (-not (Test-Path $artifactOperationLockModule)) {
    throw "Artifact-operation lock module not found: $artifactOperationLockModule"
}

Import-Module $artifactOperationLockModule -Force -ErrorAction Stop
$packagingScope = @{
    kind = 'packaging'
    target = 'msix-manifest'
    configuration = $Configuration
    platform = $Platform
    coordination = 'packaging'
}
$packagingLock = $null
try {
    $packagingLock = Enter-RSArtifactOperationLock `
        -RepoRoot $repoRoot `
        -Operation "standalone MSIX manifest stamping $Configuration|$Platform" `
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

$ManifestPath = Resolve-RSGuardedRepositoryPath `
    -RepoRoot $repoRoot `
    -GuardedRelativeRoot 'Installer\msix' `
    -Path $ManifestPath

if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    throw "MSIX version input must be a three-part version such as 7.0.184: $Version"
}

if (-not (Test-Path $ManifestPath)) {
    throw "MSIX manifest not found: $ManifestPath"
}

$identityVersion = "$Version.0"
$architecture = if ($Platform -eq "ARM64") { "arm64" } else { "x64" }

if ([string]::IsNullOrWhiteSpace($OutputManifestPath)) {
    $OutputManifestPath = Join-Path $repoRoot `
        ('.build\AppPackages\ManifestStaging\{0}\Package.appxmanifest' -f [Guid]::NewGuid().ToString('N'))
}
$outputParent = Resolve-RSGuardedRepositoryPath `
    -RepoRoot $repoRoot `
    -GuardedRelativeRoot '.build\AppPackages' `
    -Path (Split-Path $OutputManifestPath -Parent) `
    -CreateDirectory
$OutputManifestPath = Resolve-RSGuardedRepositoryPath `
    -RepoRoot $repoRoot `
    -GuardedRelativeRoot '.build\AppPackages' `
    -Path (Join-Path $outputParent (Split-Path $OutputManifestPath -Leaf))
if (Test-Path -LiteralPath $OutputManifestPath) {
    throw "MSIX staged manifest output already exists; a unique path is required: $OutputManifestPath"
}

$manifest = [xml](Get-Content -Path $ManifestPath -Raw)
$identity = $manifest.Package.Identity
if (-not $identity) {
    throw "MSIX manifest Identity element not found: $ManifestPath"
}

$identity.Version = $identityVersion
$identity.ProcessorArchitecture = $architecture

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$settings = [System.Xml.XmlWriterSettings]::new()
$settings.Encoding = $utf8NoBom
$settings.Indent = $true

$stream = $null
$writer = $null
$writeSucceeded = $false
try {
    $stream = [IO.File]::Open($OutputManifestPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
    $manifest.Save($writer)
    $writeSucceeded = $true
}
finally {
    if ($null -ne $writer) { $writer.Dispose() }
    if ($null -ne $stream) { $stream.Dispose() }
    if (-not $writeSucceeded -and (Test-Path -LiteralPath $OutputManifestPath -PathType Leaf)) {
        Remove-Item -LiteralPath $OutputManifestPath -Force
    }
}

Write-Host "Staged MSIX manifest: Version=$identityVersion ProcessorArchitecture=$architecture Path=$OutputManifestPath"
return $OutputManifestPath
}
finally {
    Exit-RSArtifactOperationLock -Lock $packagingLock
}
