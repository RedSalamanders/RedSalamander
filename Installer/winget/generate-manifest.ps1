<#
.SYNOPSIS
    Generates winget manifest files from templates.
.DESCRIPTION
    Hashes the x64 and ARM64 portable packages and writes the reviewed multi-file
    manifest from repository templates.
.PARAMETER Version
    Package version (e.g., 7.0.183). If not provided, derives it from the shared version helper.
.PARAMETER BuildNumber
    Positive compiled package build number used when Version is omitted. One of
    Version or BuildNumber is required.
.PARAMETER ZipPath
    Required path to the x64 ZIP installer. Its leaf name must be exactly
    RedSalamander-<Version>-x64-Portable.zip.
.PARAMETER Arm64ZipPath
    Required path to the ARM64 ZIP installer. Its leaf name must be exactly
    RedSalamander-<Version>-ARM64-Portable.zip.
.PARAMETER OutputDir
    Output directory for manifest files. It must resolve beneath the repository's
    .build\AppPackages root. Default: .build\AppPackages\winget-manifest.
.EXAMPLE
    .\Installer\winget\generate-manifest.ps1 -Version 7.0.183 -ZipPath .\.build\AppPackages\RedSalamander-7.0.183-x64-Portable.zip -Arm64ZipPath .\.build\AppPackages\RedSalamander-7.0.183-ARM64-Portable.zip
.OUTPUTS
    None. Writes the three Winget manifest files to OutputDir.
.NOTES
    Mutates shared packaging outputs while holding the repository-wide packaging
    coordination lock. Interrupted packaging fails closed until a reviewer restores
    Installer inputs and .build\AppPackages and removes the reported marker.
#>
[CmdletBinding()]
param(
    [string]$Version,
    [int]$BuildNumber = 0,
    [string]$ZipPath,
    [string]$Arm64ZipPath,
    [string]$OutputDir = ".build\AppPackages\winget-manifest"
)

$ErrorActionPreference = 'Stop'

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

if ([string]::IsNullOrWhiteSpace($Version) -and ($BuildNumber -le 0 -or $BuildNumber -gt 65535)) {
    throw 'Winget manifest generation requires either -Version or -BuildNumber in the range 1..65535; it never allocates or guesses a package version.'
}

$RepoRoot = Split-Path -Parent $PSScriptRoot | Split-Path -Parent
$VersioningModule = Join-Path $RepoRoot "Tools\Modules\Build\Versioning.psm1"
$ArtifactOperationLockModule = Join-Path $RepoRoot "Tools\Modules\Build\ArtifactOperationLock.psm1"
$TemplateDir = Join-Path $PSScriptRoot "templates"
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
                -Reason 'The previous packaging owner exited while mutating shared installer inputs or AppPackages outputs.' `
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
        throw "Shared packaging state may be incomplete after an interrupted operation. Review and restore Installer inputs and .build\AppPackages, then remove the marker explicitly. Marker: $MarkerPath"
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
    -GuardedRelativeRoot '.build\AppPackages' `
    -Path $OutputDir `
    -CreateDirectory

Write-Host "Generating winget manifest for version $Version..." -ForegroundColor Cyan

# Calculate SHA256 for the portable archive
Write-Host "  Calculating x64 ZIP SHA256..." -ForegroundColor Gray
$ZipSha256 = (Get-FileHash -LiteralPath $ZipPath -Algorithm SHA256).Hash
Write-Host "  Calculating ARM64 ZIP SHA256..." -ForegroundColor Gray
$Arm64ZipSha256 = (Get-FileHash -LiteralPath $Arm64ZipPath -Algorithm SHA256).Hash

# Release date (today)
$ReleaseDate = Get-Date -Format "yyyy-MM-dd"

# Template substitutions
$Replacements = @{
    '{VERSION}' = $Version
    '{ VERSION }' = $Version
    '{ZIP_SHA256}' = $ZipSha256
    '{ ZIP_SHA256 }' = $ZipSha256
    '{ARM64_ZIP_SHA256}' = $Arm64ZipSha256
    '{ ARM64_ZIP_SHA256 }' = $Arm64ZipSha256
    '{RELEASE_DATE}' = $ReleaseDate
    '{ RELEASE_DATE }' = $ReleaseDate
}

# Process each template
$Templates = @(
    "RedSalamanders.RedSalamander.installer.yaml",
    "RedSalamanders.RedSalamander.locale.en-US.yaml",
    "RedSalamanders.RedSalamander.yaml"
)

foreach ($Template in $Templates) {
    $TemplatePath = Join-Path $TemplateDir $Template
    $OutputPath = Join-Path $OutputDir $Template
    
    Write-Host "  Processing $Template..." -ForegroundColor Gray
    
    $Content = Get-Content $TemplatePath -Raw
    foreach ($Key in $Replacements.Keys) {
        $Content = $Content -replace [regex]::Escape($Key), $Replacements[$Key]
    }
    
    Set-Content -Path $OutputPath -Value $Content -Encoding UTF8 -NoNewline
}

Write-Host "✓ Winget manifest generated successfully!" -ForegroundColor Green
Write-Host "  Output: $OutputDir" -ForegroundColor Gray
Write-Host "`nNext steps:" -ForegroundColor Cyan
Write-Host "  1. Review the manifest files in $OutputDir" -ForegroundColor Gray
Write-Host "  2. Validate with: winget validate --manifest $OutputDir" -ForegroundColor Gray
Write-Host "  3. Test install with: winget install --manifest $OutputDir" -ForegroundColor Gray
Write-Host "  4. Submit to winget-pkgs repository (see docs/WingetIntegration.md)" -ForegroundColor Gray
}
finally {
    Exit-RSArtifactOperationLock -Lock $PackagingLock
}
