<#
.SYNOPSIS
    Generates winget manifest files from templates.
.PARAMETER Version
    Package version (e.g., 7.0.183). If not provided, derives it from the shared version helper.
.PARAMETER BuildNumber
    Build number override used when Version is omitted.
.PARAMETER MsiPath
    Path to the MSI installer (to calculate SHA256).
.PARAMETER ZipPath
    Path to the ZIP installer (to calculate SHA256).
.PARAMETER OutputDir
    Output directory for manifest files. Default: .build\AppPackages\winget-manifest
.PARAMETER ProductCode
    MSI Product Code GUID. If not provided, extracts from MSI.
#>
param(
    [string]$Version,
    [int]$BuildNumber = 0,
    [string]$MsiPath,
    [string]$ZipPath,
    [string]$OutputDir = ".build\AppPackages\winget-manifest",
    [string]$ProductCode
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

# Calculate SHA256 for installers
$MsiSha256 = ""
$ZipSha256 = ""

if ($MsiPath -and (Test-Path $MsiPath)) {
    Write-Host "  Calculating MSI SHA256..." -ForegroundColor Gray
    $MsiSha256 = (Get-FileHash -Path $MsiPath -Algorithm SHA256).Hash
} else {
    Write-Warning "MSI path not provided or not found. SHA256 will be placeholder."
    $MsiSha256 = "MSI_SHA256_PLACEHOLDER"
}

if ($ZipPath -and (Test-Path $ZipPath)) {
    Write-Host "  Calculating ZIP SHA256..." -ForegroundColor Gray
    $ZipSha256 = (Get-FileHash -Path $ZipPath -Algorithm SHA256).Hash
} else {
    Write-Warning "ZIP path not provided or not found. SHA256 will be placeholder."
    $ZipSha256 = "ZIP_SHA256_PLACEHOLDER"
}

# Extract Product Code from MSI if not provided
if (-not $ProductCode -and $MsiPath -and (Test-Path $MsiPath)) {
    Write-Host "  Extracting Product Code from MSI..." -ForegroundColor Gray
    try {
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MsiPath, 0))
        $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, @("SELECT Value FROM Property WHERE Property='ProductCode'"))
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
        $ProductCode = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
    } catch {
        Write-Warning "Could not extract Product Code from MSI: $_"
        $ProductCode = "PRODUCT_CODE_PLACEHOLDER"
    }
} elseif (-not $ProductCode) {
    $ProductCode = "PRODUCT_CODE_PLACEHOLDER"
}

# Release date (today)
$ReleaseDate = Get-Date -Format "yyyy-MM-dd"

# Create output directory
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

# Template substitutions
$Replacements = @{
    '{VERSION}' = $Version
    '{MSI_SHA256}' = $MsiSha256
    '{ZIP_SHA256}' = $ZipSha256
    '{PRODUCT_CODE}' = $ProductCode
    '{RELEASE_DATE}' = $ReleaseDate
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
Write-Host "  4. Submit to winget-pkgs repository (see Docs/WingetIntegration.md)" -ForegroundColor Gray
