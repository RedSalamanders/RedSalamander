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
    [string]$Platform = 'x64',

    [int]$BuildNumber = 0
)

$ErrorActionPreference = 'Stop'

# Paths
$RepoRoot = Split-Path -Parent $PSScriptRoot | Split-Path -Parent
$VersioningScript = Join-Path $RepoRoot "Tools\Versioning.ps1"
$VcRuntimeScript = Join-Path $PSScriptRoot "VcRuntime.ps1"
$BuildOutputDir = Join-Path $RepoRoot ".build\$Platform\$Configuration"
$PackageOutputDir = Join-Path $RepoRoot ".build\AppPackages"
$TempDir = Join-Path $env:TEMP "RedSalamanderZip_$([Guid]::NewGuid())"

if (-not (Test-Path $VersioningScript)) {
    throw "Version helper script not found: $VersioningScript"
}

if (-not (Test-Path $VcRuntimeScript)) {
    throw "VC runtime helper script not found: $VcRuntimeScript"
}

. $VersioningScript
. $VcRuntimeScript
$VersionContext = if ($BuildNumber -gt 0) {
    Get-RSVersionContext -RepoRoot $RepoRoot -Configuration $Configuration -Platform $Platform -BuildNumber $BuildNumber
} else {
    $savedContext = Read-RSVersionContext -RepoRoot $RepoRoot
    if ($savedContext) { $savedContext } else { Get-RSVersionContext -RepoRoot $RepoRoot -Configuration $Configuration -Platform $Platform }
}
$Version = $VersionContext.PackagingVersion

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
    Copy-Item (Join-Path $BuildOutputDir "RedLauncher.exe") $TempDir
    Copy-Item (Join-Path $BuildOutputDir "RedLauncherConsole.exe") $TempDir
    Copy-Item (Join-Path $BuildOutputDir "RedSalamander.exe") $TempDir
    Copy-Item (Join-Path $BuildOutputDir "RedSalamanderMonitor.exe") $TempDir
    
    $SearchServiceExe = Join-Path $BuildOutputDir "RedSalamanderSearchService.exe"
    if (Test-Path $SearchServiceExe) {
        Copy-Item $SearchServiceExe $TempDir
    }

    # Copy runtime DLLs (excluding system DLLs)
    Get-ChildItem (Join-Path $BuildOutputDir "*.dll") | Where-Object {
        $_.Name -notlike "vcruntime*.dll" -and
        $_.Name -notlike "msvcp*.dll" -and
        $_.Name -notlike "concrt*.dll"
    } | Copy-Item -Destination $TempDir

    # Bundle the app-local MSVC CRT so Winget validation and direct ZIP installs
    # can launch RedSalamander on clean machines without preinstalled runtimes.
    $VcRuntimeDlls = Copy-RSVcRuntimeDependencies -Platform $Platform -DestinationDir $TempDir
    Write-Host "  Bundled MSVC runtime DLLs: $($VcRuntimeDlls -join ', ')" -ForegroundColor Gray

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
2. Winget's "RedSalamander" command alias runs the flash-free RedLauncher.exe,
   which starts RedSalamander.exe from this directory so app-local DLLs are found
3. Run RedLauncherConsole.exe when a console foreground wait and process exit
   code are required, for example self-test automation
4. Run RedSalamanderMonitor.exe for debugging/monitoring

RUNTIME REQUIREMENT
-------------------
This package includes the Microsoft Visual C++ runtime DLLs required for this
CPU architecture.

PORTABLE MODE
-------------
Settings are stored in:
  - AppData: %LOCALAPPDATA%\RedSalamander

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
