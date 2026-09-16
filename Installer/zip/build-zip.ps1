<#
.SYNOPSIS
    Creates a portable ZIP package for RedSalamander.
.DESCRIPTION
    Builds a versioned portable archive from one compiled output profile, then
    validates the archive from a clean extraction before accepting it.
.PARAMETER Configuration
    Build configuration (Debug or Release). Default: Release.
.PARAMETER Platform
    Target platform (x64 or ARM64). Default: x64.
.PARAMETER BuildNumber
    Required positive build number identifying the already compiled profile.
.EXAMPLE
    .\Installer\zip\build-zip.ps1 -Configuration Release -Platform x64 -BuildNumber 184
.OUTPUTS
    None. Writes the portable ZIP beneath .build\AppPackages.
.NOTES
    Mutates shared packaging outputs while holding the repository-wide packaging
    coordination lock. Interrupted packaging fails closed until a reviewer restores
    Installer inputs and .build\AppPackages and removes the reported marker.
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    
    [ValidateSet('x64', 'ARM64')]
    [string]$Platform = 'x64',

    [int]$BuildNumber = 0
)

$ErrorActionPreference = 'Stop'

if ($BuildNumber -le 0 -or $BuildNumber -gt 65535) {
    throw 'Standalone ZIP packaging requires -BuildNumber in the range 1..65535. Use build.ps1 -Zip to propagate the compiled profile version automatically.'
}

# Paths
$RepoRoot = Split-Path -Parent $PSScriptRoot | Split-Path -Parent
$VersioningModule = Join-Path $RepoRoot "Tools\Modules\Build\Versioning.psm1"
$ArtifactOperationLockModule = Join-Path $RepoRoot "Tools\Modules\Build\ArtifactOperationLock.psm1"
$VcRuntimeScript = Join-Path $PSScriptRoot "VcRuntime.ps1"
$RuntimeDependencyModule = Join-Path $RepoRoot "Tools\Modules\Packaging\RuntimeDependencies.psm1"
$PortablePackageSmokeModule = Join-Path $RepoRoot "Tools\Modules\Packaging\PortablePackageSmoke.psm1"
$BuildOutputDir = Join-Path $RepoRoot ".build\$Platform\$Configuration"
$PackageSmokeHarness = Join-Path $BuildOutputDir "PluginContractTests.exe"
$PackageOutputDir = Join-Path $RepoRoot ".build\AppPackages"
$TempDir = Join-Path $env:TEMP "RedSalamanderZip_$([Guid]::NewGuid())"

if (-not (Test-Path $VersioningModule)) {
    throw "Version helper module not found: $VersioningModule"
}

if (-not (Test-Path $ArtifactOperationLockModule)) {
    throw "Artifact-operation lock module not found: $ArtifactOperationLockModule"
}

if (-not (Test-Path $VcRuntimeScript)) {
    throw "VC runtime helper script not found: $VcRuntimeScript"
}

if (-not (Test-Path $RuntimeDependencyModule)) {
    throw "Runtime dependency helper not found: $RuntimeDependencyModule"
}

if (-not (Test-Path $PortablePackageSmokeModule)) {
    throw "Portable package smoke helper not found: $PortablePackageSmokeModule"
}

Import-Module $VersioningModule -Force -ErrorAction Stop
Import-Module $ArtifactOperationLockModule -Force -ErrorAction Stop
. $VcRuntimeScript
Import-Module $RuntimeDependencyModule -Force -ErrorAction Stop
Import-Module $PortablePackageSmokeModule -Force -ErrorAction Stop

# Reject a pre-existing .build or AppPackages reparse path before lock metadata
# or package output can be redirected outside the repository. Missing components
# are created only after the packaging lease is held.
$PackageOutputDir = Resolve-RSGuardedRepositoryPath `
    -RepoRoot $RepoRoot `
    -GuardedRelativeRoot '.build\AppPackages' `
    -Path $PackageOutputDir
$ArtifactProfileScope = @{
    kind = 'packaging-input'
    target = 'portable-zip'
    configuration = $Configuration
    platform = $Platform
}
$PackagingScope = @{
    kind = 'packaging'
    target = 'portable-zip'
    configuration = $Configuration
    platform = $Platform
    coordination = 'packaging'
}
$ArtifactProfileLock = $null
$PackagingLock = $null
try {
    $ArtifactProfileLock = Enter-RSArtifactOperationLock `
        -RepoRoot $RepoRoot `
        -Operation "standalone portable ZIP input $Configuration|$Platform" `
        -Scope $ArtifactProfileScope
    if ($ArtifactProfileLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $RepoRoot `
                -Reason 'The previous build or packaging reader exited while using this compiled artifact profile.' `
                -AbandonedOwner $ArtifactProfileLock.AbandonedOwner `
                -Scope $ArtifactProfileScope)
    }
    if (Test-RSArtifactOperationContaminated -RepoRoot $RepoRoot -Scope $ArtifactProfileScope) {
        $ProfileMarkerPath = Get-RSArtifactContaminationMarkerPath `
            -RepoRoot $RepoRoot `
            -Scope $ArtifactProfileScope
        throw "Compiled packaging inputs may be incomplete after an interrupted operation. Run a full-solution build.ps1 -Rebuild for $Configuration|$Platform before packaging. Marker: $ProfileMarkerPath"
    }
    Assert-RSNoResidualArtifactToolProcesses `
        -RepoRoot $RepoRoot `
        -Scope $ArtifactProfileScope

    $PackagingLock = Enter-RSArtifactOperationLock `
        -RepoRoot $RepoRoot `
        -Operation "standalone portable ZIP packaging $Configuration|$Platform" `
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

    $PackageOutputDir = Resolve-RSGuardedRepositoryPath `
        -RepoRoot $RepoRoot `
        -GuardedRelativeRoot '.build\AppPackages' `
        -Path $PackageOutputDir `
        -CreateDirectory

$VersionContext = New-RSVersionContext `
    -RepoRoot $RepoRoot `
    -Configuration $Configuration `
    -Platform $Platform `
    -BuildNumber $BuildNumber
$Version = $VersionContext.PackagingVersion

# Output ZIP path
$ZipFileName = "RedSalamander-$Version-$Platform-Portable.zip"
$ZipPath = Join-Path $PackageOutputDir $ZipFileName
$StagedZipPath = Join-Path $PackageOutputDir `
    ('.{0}.{1}.tmp.zip' -f $ZipFileName, [Guid]::NewGuid().ToString('N'))

Write-Host "Creating portable ZIP package..." -ForegroundColor Cyan
Write-Host "  Version: $Version" -ForegroundColor Gray
Write-Host "  Platform: $Platform" -ForegroundColor Gray
Write-Host "  Configuration: $Configuration" -ForegroundColor Gray
Write-Host "  Output: $ZipPath" -ForegroundColor Gray

# Verify build output exists
if (-not (Test-Path $BuildOutputDir)) {
    throw "Build output directory not found: $BuildOutputDir. Run build.ps1 first."
}

$null = Assert-RSRuntimeDependenciesInOutput -RepoRoot $RepoRoot -BuildOutputDir $BuildOutputDir -Configuration $Configuration -Platform $Platform
if (-not (Test-Path -LiteralPath $PackageSmokeHarness -PathType Leaf)) {
    throw "Receipt-bound package smoke harness not found: $PackageSmokeHarness. Rebuild the complete $Configuration|$Platform profile."
}

# Create temp directory
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

try {
    # Copy main executables
    Copy-Item (Join-Path $BuildOutputDir "RedLauncher.exe") $TempDir
    Copy-Item (Join-Path $BuildOutputDir "red.exe") $TempDir
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

    # Copy localization satellite DLLs that live beside the Plugins folder.
    $LangSource = Join-Path $BuildOutputDir "Lang"
    if (Test-Path $LangSource) {
        $LangDest = Join-Path $TempDir "Lang"
        Copy-Item $LangSource -Destination $LangDest -Recurse -Force

        # Remove build artifacts from Lang
        Get-ChildItem $LangDest -Recurse -Include "*.pdb","*.lib","*.exp","*.ilk","*.iobj","*.ipdb" | Remove-Item -Force
    }

    # Copy Themes folder
    Import-Module (Join-Path $RepoRoot 'Tools/Modules/Build/DxUiDependency.psm1')
    Copy-RSDxUiPackageProvenance -RepoRoot $RepoRoot -BuildOutputDir $BuildOutputDir -PackageRoot $TempDir

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
2. Winget's "RedSalamander" and "red" command aliases run detached-console
   launcher executables, which start RedSalamander.exe from this directory so
   app-local DLLs are found
3. Run RedLauncher.exe with self-test flags when a foreground wait and process
   exit code are required, for example self-test automation
4. Run RedSalamanderMonitor.exe for debugging/monitoring

RUNTIME REQUIREMENT
-------------------
RedSalamander requires Windows 11 build 22000.2600 or later.

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

    # Create and validate a unique same-directory archive before publishing it.
    # The prior complete ZIP remains untouched until the final atomic operation.
    Write-Host "Compressing files..." -ForegroundColor Cyan
    Compress-Archive -Path "$TempDir\*" -DestinationPath $StagedZipPath

    $HostCanExecutePackage = $Platform -eq 'x64' -or $env:PROCESSOR_ARCHITECTURE -eq 'ARM64'
    $SmokeResult = Test-RSPortablePackage `
        -RepoRoot $RepoRoot `
        -ZipPath $StagedZipPath `
        -BuildOutputDir $BuildOutputDir `
        -Configuration $Configuration `
        -Platform $Platform `
        -SkipExecution:(-not $HostCanExecutePackage)
    if ($SmokeResult.ExecutionSkipped) {
        Write-Host "  Portable package dependency closure validated; $Platform execution skipped on this host." -ForegroundColor Yellow
    } else {
        Write-Host "  Portable package app startup and all built-in plugin contracts passed." -ForegroundColor Gray
    }

    # Revalidate every component immediately before the atomic replacement.
    $PackageOutputDir = Resolve-RSGuardedRepositoryPath `
        -RepoRoot $RepoRoot `
        -GuardedRelativeRoot '.build\AppPackages' `
        -Path $PackageOutputDir `
        -CreateDirectory
    $ZipPath = Publish-RSGuardedRepositoryFile `
        -RepoRoot $RepoRoot `
        -GuardedRelativeRoot '.build\AppPackages' `
        -StagedPath $StagedZipPath `
        -DestinationPath $ZipPath

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
    if (-not [string]::IsNullOrEmpty($StagedZipPath) -and
        (Test-Path -LiteralPath $StagedZipPath -PathType Leaf)) {
        Remove-Item -LiteralPath $StagedZipPath -Force
    }
}
}
finally {
    Exit-RSArtifactOperationLock -Lock $PackagingLock
    Exit-RSArtifactOperationLock -Lock $ArtifactProfileLock
}
