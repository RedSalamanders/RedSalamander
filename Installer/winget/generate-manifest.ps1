<#
.SYNOPSIS
    Generates winget manifest files from templates.
.PARAMETER Version
    Package version (e.g., 7.0.183). If not provided, derives it from the shared version helper.
.PARAMETER BuildNumber
    Build number override used when Version is omitted.
.PARAMETER ZipPath
    Path to the x64 ZIP installer (to calculate SHA256).
.PARAMETER Arm64ZipPath
    Path to the ARM64 ZIP installer (to calculate SHA256).
.PARAMETER OutputDir
    Output directory for manifest files. Default: .build\AppPackages\winget-manifest
#>
param(
    [string]$Version,
    [int]$BuildNumber = 0,
    [string]$ZipPath,
    [string]$Arm64ZipPath,
    [string]$OutputDir = ".build\AppPackages\winget-manifest"
)

$ErrorActionPreference = 'Stop'

$RepoRoot = Split-Path -Parent $PSScriptRoot | Split-Path -Parent
$VersioningScript = Join-Path $RepoRoot "Tools\Versioning.ps1"
$TemplateDir = Join-Path $PSScriptRoot "templates"

# Read version from the shared version helper if not provided
if (-not $Version) {
    if (-not (Test-Path $VersioningScript)) {
        throw "Version helper script not found: $VersioningScript"
    }

    . $VersioningScript
    $VersionContext = if ($BuildNumber -gt 0) {
        Get-RSVersionContext -RepoRoot $RepoRoot -Configuration Release -Platform x64 -BuildNumber $BuildNumber
    } else {
        $savedContext = Read-RSVersionContext -RepoRoot $RepoRoot
        if ($savedContext) { $savedContext } else { Get-RSVersionContext -RepoRoot $RepoRoot -Configuration Release -Platform x64 }
    }
    $Version = $VersionContext.PackagingVersion
}

Write-Host "Generating winget manifest for version $Version..." -ForegroundColor Cyan

# Calculate SHA256 for the portable archive
$ZipSha256 = ""
$Arm64ZipSha256 = "ARM64_ZIP_SHA256_PLACEHOLDER"

if ($ZipPath -and (Test-Path $ZipPath)) {
    Write-Host "  Calculating ZIP SHA256..." -ForegroundColor Gray
    $ZipSha256 = (Get-FileHash -Path $ZipPath -Algorithm SHA256).Hash
} else {
    Write-Warning "ZIP path not provided or not found. SHA256 will be placeholder."
    $ZipSha256 = "ZIP_SHA256_PLACEHOLDER"
}

if ($Arm64ZipPath -and (Test-Path $Arm64ZipPath)) {
    Write-Host "  Calculating ARM64 ZIP SHA256..." -ForegroundColor Gray
    $Arm64ZipSha256 = (Get-FileHash -Path $Arm64ZipPath -Algorithm SHA256).Hash
} else {
    throw "ARM64 ZIP path not provided or not found: $Arm64ZipPath"
}

# Release date (today)
$ReleaseDate = Get-Date -Format "yyyy-MM-dd"

# Create output directory
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

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
