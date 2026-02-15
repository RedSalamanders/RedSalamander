[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Build configuration (Release only)")]
    [ValidateSet("Release")]
    [string]$Configuration = "Release",

    [Parameter(HelpMessage = "Target platform (x64 only)")]
    [ValidateSet("x64")]
    [string]$Platform = "x64",

    [Parameter(HelpMessage = "Output directory for the MSI")]
    [string]$OutputDirectory = $null
)

$ErrorActionPreference = "Stop"

function Get-DefineInt {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HeaderPath,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $pattern = '^\s*#define\s+' + [Regex]::Escape($Name) + '\s+(\d+)\s*$'
    $line = Get-Content -Path $HeaderPath | Where-Object { $_ -match $pattern } | Select-Object -First 1
    if (-not $line) {
        throw "Failed to find $Name in $HeaderPath"
    }

    $match = [Regex]::Match($line, $pattern)
    return [int]$match.Groups[1].Value
}

$solutionDir = (Resolve-Path (Join-Path $PSScriptRoot "..\\..")).Path
$solutionDirWithSlash = $solutionDir.TrimEnd('\') + '\'

$versionHeader = Join-Path $solutionDir "Common\\Version.h"
$major = Get-DefineInt -HeaderPath $versionHeader -Name "VERSINFO_MAJOR"
$minor = Get-DefineInt -HeaderPath $versionHeader -Name "VERSINFO_MINORA"
$build = Get-DefineInt -HeaderPath $versionHeader -Name "VERSINFO_BUILDNUMBER"

if ($major -gt 255 -or $minor -gt 255 -or $build -gt 65535) {
    throw "MSI version components out of range (major/minor must be <=255, build must be <=65535): $major.$minor.$build"
}

$productVersion = "$major.$minor.$build"

$releaseDir = Join-Path $solutionDir (".build\\{0}\\{1}" -f $Platform, $Configuration)
if (-not (Test-Path $releaseDir)) {
    throw "Build output not found: $releaseDir"
}

$outputDir = if ($OutputDirectory) { $OutputDirectory } else { (Join-Path $solutionDir ".build\\AppPackages") }
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$msiPath = Join-Path $outputDir ("RedSalamanderSymbols-{0}-{1}.msi" -f $productVersion, $Platform)

$workRoot = Join-Path $env:TEMP ("RedSalamanderMsiSymbols_{0}_{1}_{2}" -f $productVersion, $Platform, ([guid]::NewGuid().ToString("N")))
$stageDir = Join-Path $workRoot "stage"
$objDir = Join-Path $workRoot "obj"
New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
New-Item -ItemType Directory -Path $objDir -Force | Out-Null

$shippingPdbs = @(
    "RedSalamander.pdb",
    "RedSalamanderMonitor.pdb",
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

function Get-WixSemVer {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WixExe
    )

    $versionText = & $WixExe --version
    if ($LASTEXITCODE -ne 0) {
        throw "wix.exe --version failed with exit code $LASTEXITCODE"
    }

    $semVer = (($versionText -split '\+')[0]).Trim()
    if (-not $semVer) {
        throw "Failed to parse wix.exe version from: $versionText"
    }

    return $semVer
}

function Ensure-WixUiExtension {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WixExe
    )

    $extensionsText = & $WixExe extension list -g 2>$null
    if ($LASTEXITCODE -ne 0) {
        $extensionsText = ""
    }

    if ($extensionsText -and ($extensionsText | Where-Object { $_ -like "WixToolset.UI.wixext *" })) {
        return
    }

    $semVer = Get-WixSemVer -WixExe $WixExe

    Write-Host "Installing WiX UI extension (WixToolset.UI.wixext/$semVer)..." -ForegroundColor Yellow
    & $WixExe extension add "WixToolset.UI.wixext/$semVer" -g | Out-Null

    if ($LASTEXITCODE -ne 0) {
        throw "Failed to install WiX UI extension (WixToolset.UI.wixext/$semVer)."
    }
}

$wixExe = Find-WixExe
Ensure-WixUiExtension -WixExe $wixExe

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

& $wixExe @wixArgs
if ($LASTEXITCODE -ne 0) {
    throw "wix build failed with exit code $LASTEXITCODE"
}

Write-Host "MSI created: $msiPath" -ForegroundColor Green
