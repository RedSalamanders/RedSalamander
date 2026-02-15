<#
.SYNOPSIS
    Build RedSalamander MSI installer (WiX Toolset v6+).
.DESCRIPTION
    Stages the Release output folder and builds an MSI from `Installer\msi\Product.wxs`.
    This is intended for simple distribution scenarios where MSIX is undesirable.
.PARAMETER Configuration
    Build configuration (Release only).
.PARAMETER Platform
    Target platform (x64 only).
.PARAMETER OutputDirectory
    Output directory for the MSI (default: .build\AppPackages\ under solution root).
.EXAMPLE
    .\Installer\msi\build-msi.ps1 -Configuration Release -Platform x64
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

$msiPath = Join-Path $outputDir ("RedSalamander-{0}-{1}.msi" -f $productVersion, $Platform)

$workRoot = Join-Path $env:TEMP ("RedSalamanderMsi_{0}_{1}_{2}" -f $productVersion, $Platform, ([guid]::NewGuid().ToString("N")))
$stageDir = Join-Path $workRoot "stage"
$objDir = Join-Path $workRoot "obj"
New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
New-Item -ItemType Directory -Path $objDir -Force | Out-Null

# Stage files into a clean directory so we can exclude build artifacts (PDB/etc) consistently.
# Also exclude non-shipping executables (PoC/test), ASAN artifacts, and unused third-party DLLs.
$exclude = @("*.pdb", "*.lib", "*.exp", "*.ilk", "*.iobj", "*.ipdb", "*.pch", "*.exe", "asan.supp", "aws-cpp-sdk-dynamodb.dll", "aws-cpp-sdk-kinesis.dll", "aws-cpp-sdk-s3.dll", "brotlienc.dll")
$robocopyArgs = @(
    $releaseDir,
    $stageDir,
    "/E",
    "/R:1",
    "/W:1",
    "/NFL",
    "/NDL",
    "/NJH",
    "/NJS",
    "/NP",
    "/XF"
) + $exclude

& robocopy.exe @robocopyArgs | Out-Null
if ($LASTEXITCODE -ge 8) {
    throw "robocopy failed with exit code $LASTEXITCODE"
}

# Copy shipping executables explicitly (we excluded *.exe above).
$shippingExes = @("RedSalamander.exe", "RedSalamanderMonitor.exe")
foreach ($exe in $shippingExes) {
    $src = Join-Path $releaseDir $exe
    if (-not (Test-Path $src)) {
        throw "Expected shipping executable not found: $src"
    }
}
& robocopy.exe $releaseDir $stageDir @shippingExes "/R:1" "/W:1" "/NFL" "/NDL" "/NJH" "/NJS" "/NP" | Out-Null
if ($LASTEXITCODE -ge 8) {
    throw "robocopy (shipping exes) failed with exit code $LASTEXITCODE"
}
foreach ($exe in $shippingExes) {
    $dst = Join-Path $stageDir $exe
    if (-not (Test-Path $dst)) {
        throw "Failed to stage shipping executable: $dst"
    }
}

$productWxs = Join-Path $PSScriptRoot "Product.wxs"
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

Write-Host "Building MSI..." -ForegroundColor Yellow
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
