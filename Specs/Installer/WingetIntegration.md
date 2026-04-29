# Winget Integration Guide for RedSalamander

This document provides step-by-step instructions for integrating RedSalamander with Windows Package Manager (winget).

## Overview

Winget supports multiple installer types:
- **MSI** (recommended for traditional installs)
- **MSIX** (for Store-like installs)
- **Portable ZIP** (for users who prefer portable apps)

This guide covers all three approaches.

---

## Step 1: Create Portable ZIP Package

### 1.1 Create ZIP Build Script

Create `Installer\zip\build-zip.ps1`:

```powershell
<#
.SYNOPSIS
    Creates a portable ZIP package for RedSalamander.
.PARAMETER Configuration
    Build configuration (Debug or Release). Default: Release.
.PARAMETER Platform
    Target platform (x64 or ARM64). Default: x64.
#>
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    
    [ValidateSet('x64', 'ARM64')]
    [string]$Platform = 'x64'
)

$ErrorActionPreference = 'Stop'

# Paths
$RepoRoot = Split-Path -Parent $PSScriptRoot | Split-Path -Parent
$BuildOutputDir = Join-Path $RepoRoot ".build\$Platform\$Configuration"
$PackageOutputDir = Join-Path $RepoRoot ".build\AppPackages"
$TempDir = Join-Path $env:TEMP "RedSalamanderZip_$([Guid]::NewGuid())"

# Read version from Version.h
$VersionHeader = Join-Path $RepoRoot "Common\Version.h"
$VersionContent = Get-Content $VersionHeader -Raw
$VersionContent -match '#define VERSINFO_MAJOR\s+(\d+)' | Out-Null
$Major = $Matches[1]
$VersionContent -match '#define VERSINFO_MINORA\s+(\d+)' | Out-Null
$Minor = $Matches[1]
$VersionContent -match '#define VERSINFO_BUILDNUMBER\s+(\d+)' | Out-Null
$Build = $Matches[1]
$Version = "$Major.$Minor.$Build"

# Output ZIP path
$ZipFileName = "RedSalamander-$Version-$Platform-Portable.zip"
$ZipPath = Join-Path $PackageOutputDir $ZipFileName

Write-Host "Creating portable ZIP package..." -ForegroundColor Cyan
Write-Host "  Version: $Version" -ForegroundColor Gray
Write-Host "  Platform: $Platform" -ForegroundColor Gray
Write-Host "  Configuration: $Configuration" -ForegroundColor Gray
Write-Host "  Output: $ZipPath" -ForegroundColor Gray

# Verify build output exists
if (-not (Test-Path $BuildOutputDir)) {
    throw "Build output directory not found: $BuildOutputDir. Run build.ps1 first."
}

# Create temp directory
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

try {
    # Copy main executables
    Copy-Item (Join-Path $BuildOutputDir "RedSalamander.exe") $TempDir
    Copy-Item (Join-Path $BuildOutputDir "RedSalamanderMonitor.exe") $TempDir
    Copy-Item (Join-Path $BuildOutputDir "RedSalamanderSearchService.exe") $TempDir

    # Copy runtime DLLs (excluding system DLLs)
    Get-ChildItem (Join-Path $BuildOutputDir "*.dll") | Where-Object {
        $_.Name -notlike "vcruntime*.dll" -and
        $_.Name -notlike "msvcp*.dll" -and
        $_.Name -notlike "concrt*.dll"
    } | Copy-Item -Destination $TempDir

    # Copy Plugins folder
    $PluginsSource = Join-Path $BuildOutputDir "Plugins"
    if (Test-Path $PluginsSource) {
        $PluginsDest = Join-Path $TempDir "Plugins"
        Copy-Item $PluginsSource -Destination $PluginsDest -Recurse -Force
        
        # Remove build artifacts from Plugins
        Get-ChildItem $PluginsDest -Recurse -Include "*.pdb","*.lib","*.exp","*.ilk","*.iobj","*.ipdb" | Remove-Item -Force
    }

    # Copy Themes folder
    $ThemesSource = Join-Path $BuildOutputDir "Themes"
    if (Test-Path $ThemesSource) {
        Copy-Item $ThemesSource -Destination (Join-Path $TempDir "Themes") -Recurse -Force
    }

    # Copy SettingsStore.schema.json if present
    $SchemaSource = Join-Path $BuildOutputDir "SettingsStore.schema.json"
    if (Test-Path $SchemaSource) {
        Copy-Item $SchemaSource $TempDir
    }

    # Create README.txt for portable users
    $ReadmeContent = @"
RedSalamander $Version - Portable Edition
=============================================

This is a portable distribution of RedSalamander.

GETTING STARTED
---------------
1. Run RedSalamander.exe to launch the file manager
2. Run RedSalamanderMonitor.exe for debugging/monitoring

PORTABLE MODE
-------------
Settings are stored in:
  - AppData: %LOCALAPPDATA%\RedSalamander\Settings

To run in fully portable mode (settings alongside the executable),
create an empty file named 'portable.ini' in this directory.

PLUGINS
-------
All plugins are located in the Plugins\ subdirectory.

THEMES
------
Themes are located in the Themes\ subdirectory.
Place custom themes here and select them in Preferences.

For more information, visit:
https://github.com/RedSalamanders/RedSalamander

"@
    Set-Content -Path (Join-Path $TempDir "README.txt") -Value $ReadmeContent -Encoding UTF8

    # Create LICENSE.txt
    $LicenseSource = Join-Path $RepoRoot "LICENSE.txt"
    if (Test-Path $LicenseSource) {
        Copy-Item $LicenseSource (Join-Path $TempDir "LICENSE.txt")
    }

    # Ensure output directory exists
    New-Item -ItemType Directory -Path $PackageOutputDir -Force | Out-Null

    # Create ZIP archive
    Write-Host "Compressing files..." -ForegroundColor Cyan
    Compress-Archive -Path "$TempDir\*" -DestinationPath $ZipPath -Force

    Write-Host "✓ Portable ZIP created successfully!" -ForegroundColor Green
    Write-Host "  $ZipPath" -ForegroundColor Gray
    
    # Show size
    $ZipSize = (Get-Item $ZipPath).Length / 1MB
    Write-Host "  Size: $($ZipSize.ToString('F2')) MB" -ForegroundColor Gray

} finally {
    # Cleanup temp directory
    if (Test-Path $TempDir) {
        Remove-Item $TempDir -Recurse -Force
    }
}
```

### 1.2 Update build.ps1

Add a `-Zip` parameter to the main build script. Edit `build.ps1`:

```powershell
# Add to param block:
[switch]$Zip,

# Add after MSIX/MSI build logic:
if ($Zip -and $Configuration -eq 'Release') {
    Write-Host "`n=== Building ZIP Package ===" -ForegroundColor Cyan
    & "$PSScriptRoot\Installer\zip\build-zip.ps1" -Configuration $Configuration -Platform $Platform
    if ($LASTEXITCODE -ne 0) {
        throw "ZIP packaging failed"
    }
}
```

### 1.3 Test ZIP Package

```powershell
# Build and create ZIP
.\build.ps1 -Configuration Release -Zip

# Verify output
ls .build\AppPackages\*.zip
```

---

## Step 2: Create Winget Manifest Files

Winget manifests are YAML files that describe your package. You need three files:
1. **Installer manifest** (`.installer.yaml`) - installer details
2. **Locale manifest** (`.locale.en-US.yaml`) - localized metadata
3. **Version manifest** (`.yaml`) - version metadata

### 2.1 Create Manifest Template Directory

Create directory structure:

```
Installer/winget/
├── generate-manifest.ps1
└── templates/
    ├── RedSalamanders.RedSalamander.installer.yaml
    ├── RedSalamanders.RedSalamander.locale.en-US.yaml
    └── RedSalamanders.RedSalamander.yaml
```

### 2.2 Create Installer Manifest

Create `Installer\winget\templates\RedSalamanders.RedSalamander.installer.yaml`:

```yaml
# yaml-language-server: $schema=https://aka.ms/winget-manifest.installer.1.6.0.schema.json

PackageIdentifier: RedSalamanders.RedSalamander
PackageVersion: {VERSION}
Platform:
  - Windows.Desktop
MinimumOSVersion: 10.0.19041.0
InstallerType: wix
Scope: machine
InstallModes:
  - interactive
  - silent
  - silentWithProgress
UpgradeBehavior: install
Protocols:
  - red-salamander
FileExtensions:
  - red-theme
ReleaseDate: {RELEASE_DATE}
Installers:
  - Architecture: x64
    InstallerUrl: https://github.com/RedSalamanders/RedSalamander/releases/download/v{VERSION}/RedSalamander-{VERSION}-x64.msi
    InstallerSha256: {MSI_SHA256}
    ProductCode: '{PRODUCT_CODE}'
  - Architecture: x64
    InstallerType: portable
    InstallerUrl: https://github.com/RedSalamanders/RedSalamander/releases/download/v{VERSION}/RedSalamander-{VERSION}-x64-Portable.zip
    InstallerSha256: {ZIP_SHA256}
ManifestType: installer
ManifestVersion: 1.6.0
```

### 2.3 Create Locale Manifest

Create `Installer\winget\templates\RedSalamanders.RedSalamander.locale.en-US.yaml`:

```yaml
# yaml-language-server: $schema=https://aka.ms/winget-manifest.defaultLocale.1.6.0.schema.json

PackageIdentifier: RedSalamanders.RedSalamander
PackageVersion: {VERSION}
PackageLocale: en-US
Publisher: RedSalamanders
PublisherUrl: https://github.com/RedSalamanders
PublisherSupportUrl: https://github.com/RedSalamanders/RedSalamander/issues
# PrivacyUrl: 
Author: RedSalamanders
PackageName: RedSalamander
PackageUrl: https://github.com/RedSalamanders/RedSalamander
License: MIT
LicenseUrl: https://github.com/RedSalamanders/RedSalamander/blob/master/LICENSE.txt
# Copyright: 
# CopyrightUrl: 
ShortDescription: Windows file manager with advanced features and plugin support
Description: |-
  RedSalamander is a Windows-native dual-pane file manager with plugin-based virtual file systems,
  advanced text visualization, and real-time debugging capabilities. Features include:
  - Dual-pane file management with virtual file system support
  - Plugin architecture for custom file systems (7z, S3, cloud drives)
  - Advanced text viewer with Direct2D rendering
  - Comparison tools and search capabilities
  - Real-time monitoring with Event Tracing for Windows (ETW)
Moniker: red-salamander
Tags:
  - file-manager
  - dual-pane
  - file-explorer
  - file-viewer
  - direct2d
  - cpp23
  - windows
  - plugins
ReleaseNotes: See https://github.com/RedSalamanders/RedSalamander/releases/tag/v{VERSION}
ReleaseNotesUrl: https://github.com/RedSalamanders/RedSalamander/releases/tag/v{VERSION}
# PurchaseUrl: 
# InstallationNotes: 
Documentations:
  - DocumentLabel: Getting Started
    DocumentUrl: https://github.com/RedSalamanders/RedSalamander/blob/master/Docs/GettingStarted.md
ManifestType: defaultLocale
ManifestVersion: 1.6.0
```

### 2.4 Create Version Manifest

Create `Installer\winget\templates\RedSalamanders.RedSalamander.yaml`:

```yaml
# yaml-language-server: $schema=https://aka.ms/winget-manifest.version.1.6.0.schema.json

PackageIdentifier: RedSalamanders.RedSalamander
PackageVersion: {VERSION}
DefaultLocale: en-US
ManifestType: version
ManifestVersion: 1.6.0
```

---

## Step 3: Create Manifest Generation Script

Create `Installer\winget\generate-manifest.ps1`:

```powershell
<#
.SYNOPSIS
    Generates winget manifest files from templates.
.PARAMETER Version
    Package version (e.g., 7.0.183). If not provided, reads from Common/Version.h.
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
    [string]$MsiPath,
    [string]$ZipPath,
    [string]$OutputDir = ".build\AppPackages\winget-manifest",
    [string]$ProductCode
)

$ErrorActionPreference = 'Stop'

$RepoRoot = Split-Path -Parent $PSScriptRoot | Split-Path -Parent
$TemplateDir = Join-Path $PSScriptRoot "templates"

# Read version from Version.h if not provided
if (-not $Version) {
    $VersionHeader = Join-Path $RepoRoot "Common\Version.h"
    $VersionContent = Get-Content $VersionHeader -Raw
    $VersionContent -match '#define VERSINFO_MAJOR\s+(\d+)' | Out-Null
    $Major = $Matches[1]
    $VersionContent -match '#define VERSINFO_MINORA\s+(\d+)' | Out-Null
    $Minor = $Matches[1]
    $VersionContent -match '#define VERSINFO_BUILDNUMBER\s+(\d+)' | Out-Null
    $Build = $Matches[1]
    $Version = "$Major.$Minor.$Build"
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
```

---

## Step 4: Integrate with build.ps1

Update `build.ps1` to generate winget manifests after building installers:

```powershell
# Add after ZIP/MSI/MSIX build logic:
if ($GenerateWingetManifest -and $Configuration -eq 'Release') {
    Write-Host "`n=== Generating Winget Manifest ===" -ForegroundColor Cyan
    
    $MsiPath = Get-ChildItem ".build\AppPackages\RedSalamander-*.msi" | Select-Object -First 1 -ExpandProperty FullName
    $ZipPath = Get-ChildItem ".build\AppPackages\RedSalamander-*-Portable.zip" | Select-Object -First 1 -ExpandProperty FullName
    
    & "$PSScriptRoot\Installer\winget\generate-manifest.ps1" -MsiPath $MsiPath -ZipPath $ZipPath
}
```

Add parameter:
```powershell
[switch]$GenerateWingetManifest,
```

---

## Step 5: Test Locally

### 5.1 Build All Installers

```powershell
# Build Release with all installer types
.\build.ps1 -Configuration Release -Msi -Zip -GenerateWingetManifest
```

### 5.2 Validate Manifest

```powershell
# Install winget-create tool if not already installed
winget install --id Microsoft.WingetCreate

# Validate manifest
winget validate --manifest .build\AppPackages\winget-manifest
```

### 5.3 Test Local Install

```powershell
# Test install from local manifest
winget install --manifest .build\AppPackages\winget-manifest

# Test uninstall
winget uninstall RedSalamanders.RedSalamander
```

---

## Step 6: Submit to Winget Community Repository

### 6.1 Prerequisites

1. Fork the winget-pkgs repository:
   - Go to https://github.com/microsoft/winget-pkgs
   - Click "Fork"

2. Clone your fork:
   ```powershell
   git clone https://github.com/YOUR_USERNAME/winget-pkgs.git
   cd winget-pkgs
   ```

### 6.2 Add Your Package

```powershell
# Create directory structure
$PackageDir = "manifests\d\RedSalamanders\RedSalamander\7.0.183"
New-Item -ItemType Directory -Path $PackageDir -Force

# Copy manifest files
Copy-Item .build\AppPackages\winget-manifest\*.yaml $PackageDir\
```

### 6.3 Create Pull Request

```powershell
# Create branch
git checkout -b RedSalamanders.RedSalamander-7.0.183

# Add and commit
git add manifests/d/RedSalamanders/RedSalamander/7.0.183/
git commit -m "Add RedSalamanders.RedSalamander version 7.0.183"

# Push to your fork
git push origin RedSalamanders.RedSalamander-7.0.183
```

Then:
1. Go to https://github.com/YOUR_USERNAME/winget-pkgs
2. Click "Pull Request"
3. Create PR to microsoft/winget-pkgs:master
4. Wait for automated validation and maintainer review

### 6.4 Update Instructions

After your first package is accepted, future updates are easier:

```powershell
# Use winget-create to update
wingetcreate update RedSalamanders.RedSalamander `
  --version 7.0.184 `
  --urls "https://github.com/RedSalamanders/RedSalamander/releases/download/v7.0.184/RedSalamander-7.0.184-x64.msi|x64" `
         "https://github.com/RedSalamanders/RedSalamander/releases/download/v7.0.184/RedSalamander-7.0.184-x64-Portable.zip|x64|portable" `
  --submit --token YOUR_GITHUB_TOKEN
```

---

## Step 7: Automate with GitHub Actions

Create `.github/workflows/winget-release.yml`:

```yaml
name: Publish to Winget

on:
  release:
    types: [published]

jobs:
  publish-winget:
    runs-on: windows-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Get version from release
        id: version
        run: |
          $version = "${{ github.event.release.tag_name }}".TrimStart('v')
          echo "version=$version" >> $env:GITHUB_OUTPUT

      - name: Download release assets
        run: |
          $version = "${{ steps.version.outputs.version }}"
          $msiUrl = "https://github.com/${{ github.repository }}/releases/download/v$version/RedSalamander-$version-x64.msi"
          $zipUrl = "https://github.com/${{ github.repository }}/releases/download/v$version/RedSalamander-$version-x64-Portable.zip"
          
          Invoke-WebRequest -Uri $msiUrl -OutFile "RedSalamander.msi"
          Invoke-WebRequest -Uri $zipUrl -OutFile "RedSalamander.zip"

      - name: Submit to Winget
        run: |
          winget install --id Microsoft.WingetCreate
          
          $version = "${{ steps.version.outputs.version }}"
          wingetcreate update RedSalamanders.RedSalamander `
            --version $version `
            --urls "https://github.com/${{ github.repository }}/releases/download/v$version/RedSalamander-$version-x64.msi|x64" `
                   "https://github.com/${{ github.repository }}/releases/download/v$version/RedSalamander-$version-x64-Portable.zip|x64|portable" `
            --submit --token ${{ secrets.WINGET_TOKEN }}
```

Add `WINGET_TOKEN` secret to your repository:
1. Generate GitHub personal access token with `public_repo` scope
2. Add as secret in repository settings

---

## Summary

**Files to Create:**
1. `Installer\zip\build-zip.ps1` - ZIP package builder
2. `Installer\winget\templates\RedSalamanders.RedSalamander.installer.yaml` - Installer manifest template
3. `Installer\winget\templates\RedSalamanders.RedSalamander.locale.en-US.yaml` - Locale manifest template
4. `Installer\winget\templates\RedSalamanders.RedSalamander.yaml` - Version manifest template
5. `Installer\winget\generate-manifest.ps1` - Manifest generator
6. `.github\workflows\winget-release.yml` - Automated publishing workflow

**Build Commands:**
```powershell
# Build all installer types
.\build.ps1 -Configuration Release -Msi -Zip -GenerateWingetManifest

# Validate manifest
winget validate --manifest .build\AppPackages\winget-manifest

# Test install locally
winget install --manifest .build\AppPackages\winget-manifest
```

**After First PR:**
```powershell
# Automated updates via winget-create
wingetcreate update RedSalamanders.RedSalamander --version X.Y.Z --urls ... --submit
```

---

## References

- [Winget Manifest Documentation](https://learn.microsoft.com/en-us/windows/package-manager/package/manifest)
- [Winget Community Repository](https://github.com/microsoft/winget-pkgs)
- [WingetCreate Tool](https://github.com/microsoft/winget-create)
- [Contributing Guidelines](https://github.com/microsoft/winget-pkgs/blob/master/CONTRIBUTING.md)
