<#
.SYNOPSIS
    Builds the RedSalamander symbols MSI with WiX Toolset v6 or later.
.DESCRIPTION
    Stages the reviewed Release PDB allowlist and creates a versioned symbols MSI.
.PARAMETER Configuration
    Build configuration. Only Release is supported.
.PARAMETER Platform
    Target platform. Only x64 is supported.
.PARAMETER OutputDirectory
    Destination directory. It must resolve to .build\AppPackages or one of its
    descendants; the default is the AppPackages root.
.PARAMETER BuildNumber
    Required positive build number identifying the already compiled Release profile.
.EXAMPLE
    .\Installer\msi\build-msi-symbols.ps1 -Configuration Release -Platform x64 -BuildNumber 184
.OUTPUTS
    None. Writes the versioned symbols MSI to OutputDirectory.
.NOTES
    Mutates shared packaging outputs while holding the repository-wide packaging
    coordination lock. Interrupted packaging fails closed until a reviewer restores
    Installer inputs and .build\AppPackages and removes the reported marker.
#>
[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Build configuration (Release only)")]
    [ValidateSet("Release")]
    [string]$Configuration = "Release",

    [Parameter(HelpMessage = "Target platform (x64 only)")]
    [ValidateSet("x64")]
    [string]$Platform = "x64",

    [Parameter(HelpMessage = "Output directory for the MSI")]
    [string]$OutputDirectory = $null,

    [Parameter(HelpMessage = "Build number to stamp into the MSI")]
    [int]$BuildNumber = 0
)

$ErrorActionPreference = "Stop"

if ($BuildNumber -le 0 -or $BuildNumber -gt 65535) {
    throw 'Standalone symbols MSI packaging requires -BuildNumber in the range 1..65535. Use build.ps1 -Msi to propagate the compiled profile version automatically.'
}

$solutionDir = (Resolve-Path (Join-Path $PSScriptRoot "..\\..")).Path
$solutionDirWithSlash = $solutionDir.TrimEnd('\') + '\'
$versioningModule = Join-Path $solutionDir "Tools\Modules\Build\Versioning.psm1"
$artifactOperationLockModule = Join-Path $solutionDir "Tools\Modules\Build\ArtifactOperationLock.psm1"
$sanitizedEnvironmentModule = Join-Path $solutionDir "Tools\Modules\Build\SanitizedEnvironment.psm1"
if (-not (Test-Path $versioningModule)) {
    throw "Version helper module not found: $versioningModule"
}
if (-not (Test-Path $artifactOperationLockModule)) {
    throw "Artifact-operation lock module not found: $artifactOperationLockModule"
}
if (-not (Test-Path $sanitizedEnvironmentModule)) {
    throw "Contained-process helper module not found: $sanitizedEnvironmentModule"
}

Import-Module $versioningModule -Force -ErrorAction Stop
Import-Module $artifactOperationLockModule -Force -ErrorAction Stop
Import-Module $sanitizedEnvironmentModule -Force -ErrorAction Stop
$artifactProfileScope = @{
    kind = 'packaging-input'
    target = 'msi-symbols'
    configuration = $Configuration
    platform = $Platform
}
$packagingScope = @{
    kind = 'packaging'
    target = 'msi-symbols'
    configuration = $Configuration
    platform = $Platform
    coordination = 'packaging'
}
$artifactProfileLock = $null
$packagingLock = $null
try {
    $artifactProfileLock = Enter-RSArtifactOperationLock `
        -RepoRoot $solutionDir `
        -Operation "standalone symbols MSI input $Configuration|$Platform" `
        -Scope $artifactProfileScope
    if ($artifactProfileLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $solutionDir `
                -Reason 'The previous build or packaging reader exited while using this compiled artifact profile.' `
                -AbandonedOwner $artifactProfileLock.AbandonedOwner `
                -Scope $artifactProfileScope)
    }
    if (Test-RSArtifactOperationContaminated -RepoRoot $solutionDir -Scope $artifactProfileScope) {
        $profileMarkerPath = Get-RSArtifactContaminationMarkerPath `
            -RepoRoot $solutionDir `
            -Scope $artifactProfileScope
        throw "Compiled packaging inputs may be incomplete after an interrupted operation. Run a full-solution build.ps1 -Rebuild for $Configuration|$Platform before packaging. Marker: $profileMarkerPath"
    }
    Assert-RSNoResidualArtifactToolProcesses `
        -RepoRoot $solutionDir `
        -Scope $artifactProfileScope

    $packagingLock = Enter-RSArtifactOperationLock `
        -RepoRoot $solutionDir `
        -Operation "standalone symbols MSI packaging $Configuration|$Platform" `
        -Scope $packagingScope
    if ($packagingLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $solutionDir `
                -Reason 'The previous packaging owner exited while mutating shared installer inputs or AppPackages outputs.' `
                -AbandonedOwner $packagingLock.AbandonedOwner `
                -Scope $packagingScope)
    }
    $packagingContamination = Read-RSArtifactOperationContamination `
        -RepoRoot $solutionDir `
        -Scope $packagingScope
    if ($null -ne $packagingContamination) {
        $markerPath = Get-RSArtifactContaminationMarkerPath `
            -RepoRoot $solutionDir `
            -Scope $packagingScope
        throw "Shared packaging state may be incomplete after an interrupted operation. Review and restore Installer inputs and .build\AppPackages, then remove the marker explicitly. Marker: $markerPath"
    }

$versionContext = New-RSVersionContext `
    -RepoRoot $solutionDir `
    -Configuration $Configuration `
    -Platform $Platform `
    -BuildNumber $BuildNumber

if ($versionContext.Major -gt 255 -or $versionContext.Minor -gt 255 -or $versionContext.BuildNumber -gt 65535) {
    throw "MSI version components out of range (major/minor must be <=255, build must be <=65535): $($versionContext.PackagingVersion)"
}

$productVersion = $versionContext.PackagingVersion

$releaseDir = Join-Path $solutionDir (".build\\{0}\\{1}" -f $Platform, $Configuration)
if (-not (Test-Path $releaseDir)) {
    throw "Build output not found: $releaseDir"
}

$requestedOutputDirectory = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    '.build\AppPackages'
} else { $OutputDirectory }
$outputDir = Resolve-RSGuardedRepositoryPath `
    -RepoRoot $solutionDir `
    -GuardedRelativeRoot '.build\AppPackages' `
    -Path $requestedOutputDirectory `
    -CreateDirectory

$msiPath = Join-Path $outputDir ("RedSalamanderSymbols-{0}-{1}.msi" -f $productVersion, $Platform)

$workRoot = Join-Path $env:TEMP ("RedSalamanderMsiSymbols_{0}_{1}_{2}" -f $productVersion, $Platform, ([guid]::NewGuid().ToString("N")))
$stageDir = Join-Path $workRoot "stage"
$objDir = Join-Path $workRoot "obj"
try {
    New-Item -ItemType Directory -Path $workRoot | Out-Null
    New-Item -ItemType Directory -Path $stageDir | Out-Null
    New-Item -ItemType Directory -Path $objDir | Out-Null

$shippingPdbs = @(
    "RedSalamander.pdb",
    "RedSalamanderMonitor.pdb",
    "RedSalamanderSearchService.pdb",
    "Common.pdb"
)

foreach ($pdb in $shippingPdbs) {
    $src = Join-Path $releaseDir $pdb
    if (-not (Test-Path $src)) {
        throw "Expected PDB not found: $src"
    }
    Copy-Item -Path $src -Destination (Join-Path $stageDir $pdb) -Force
}

$pluginsSrc = Join-Path $releaseDir "Plugins"
if (Test-Path $pluginsSrc) {
    $pluginsDst = Join-Path $stageDir "Plugins"
    New-Item -ItemType Directory -Path $pluginsDst -Force | Out-Null
    Copy-Item -Path (Join-Path $pluginsSrc "*.pdb") -Destination $pluginsDst -Force -ErrorAction SilentlyContinue
}

$productWxs = Join-Path $PSScriptRoot "ProductSymbols.wxs"
if (-not (Test-Path $productWxs)) {
    throw "MSI WiX source not found: $productWxs"
}

function Find-WixExe {
    $cmd = Get-Command "wix.exe" -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $roots = @()
    if ($env:ProgramFiles) {
        $roots += $env:ProgramFiles
    }
    if (${env:ProgramFiles(x86)}) {
        $roots += ${env:ProgramFiles(x86)}
    }

    foreach ($root in $roots) {
        $installDirs =
            Get-ChildItem -Path $root -Directory -Filter "WiX Toolset v*" -ErrorAction SilentlyContinue |
            Sort-Object -Property Name -Descending

        foreach ($dir in $installDirs) {
            $candidate = Join-Path $dir.FullName "bin\\wix.exe"
            if (Test-Path $candidate) {
                return $candidate
            }
        }
    }

    throw "WiX Toolset not found (wix.exe). Install WiX Toolset CLI v6+. Example: winget install --exact --id WiXToolset.WiXCLI"
}

$wixExe = Find-WixExe

Write-Host "Building Symbols MSI..." -ForegroundColor Yellow
$wixArgs = @(
    "build",
    $productWxs,
    "-arch", $Platform,
    "-ext", "WixToolset.UI.wixext",
    "-intermediatefolder", $objDir,
    "-d", "ProductVersion=$productVersion",
    "-d", "SolutionDir=$solutionDirWithSlash",
    "-d", "StageDir=$stageDir",
    "-out", $msiPath
)

$wixExitCode = Invoke-RSProcess `
    -FilePath $wixExe `
    -Arguments $wixArgs `
    -WorkingDirectory $solutionDir
if ($wixExitCode -ne 0) {
    throw "wix build failed with exit code $wixExitCode. Packaging does not mutate machine-global WiX extensions; install the matching WixToolset.UI.wixext version explicitly before retrying."
}

    Write-Host "MSI created: $msiPath" -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $workRoot) {
        Remove-Item -LiteralPath $workRoot -Recurse -Force
    }
}
}
finally {
    Exit-RSArtifactOperationLock -Lock $packagingLock
    Exit-RSArtifactOperationLock -Lock $artifactProfileLock
}
