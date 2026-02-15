<#
.SYNOPSIS
    Install vcpkg dependencies for RedSalamander without polluting the repo root.
.DESCRIPTION
    Runs `vcpkg install` in manifest mode, forcing the install root to:
      .build\vcpkg_installed\

    This matches the MSBuild/CI layout (see Directory.Build.props) so builds never create
    a top-level `vcpkg_installed\` folder.
.PARAMETER Platform
    Target platform used to pick the default vcpkg triplet (x64 -> x64-windows, ARM64 -> arm64-windows).
.PARAMETER Triplet
    Optional explicit vcpkg triplet (overrides -Platform mapping).
.PARAMETER Clean
    Deletes the install root (`.build\\vcpkg_installed\\`) before installing (all triplets).
.PARAMETER VcpkgExe
    Optional explicit path to vcpkg.exe (overrides discovery).
.EXAMPLE
    .\vcpkg-install.ps1
.EXAMPLE
    .\vcpkg-install.ps1 -Platform ARM64
.EXAMPLE
    .\vcpkg-install.ps1 -Triplet x64-windows -Clean
#>

[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Target platform (x64 or ARM64)")]
    [ValidateSet("x64", "ARM64")]
    [string]$Platform = "x64",

    [Parameter(HelpMessage = "Optional explicit vcpkg triplet (e.g. x64-windows, arm64-windows)")]
    [string]$Triplet = $null,

    [Parameter(HelpMessage = "Delete .build\\vcpkg_installed before installing")]
    [switch]$Clean,

    [Parameter(HelpMessage = "Optional explicit path to vcpkg.exe")]
    [string]$VcpkgExe = $null
)

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$manifestRoot = $repoRoot
$installRoot = Join-Path $repoRoot ".build\\vcpkg_installed"

if (-not (Test-Path (Join-Path $manifestRoot "vcpkg.json"))) {
    throw "vcpkg.json not found at: $manifestRoot"
}

if (-not $Triplet) {
    $Triplet = if ($Platform -eq "ARM64") { "arm64-windows" } else { "x64-windows" }
}

function Resolve-VcpkgExePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $false)]
        [string]$ExplicitVcpkgExe
    )

    if ($ExplicitVcpkgExe) {
        $candidate = $ExplicitVcpkgExe
        if (-not (Test-Path $candidate)) {
            throw "VcpkgExe not found: $candidate"
        }
        return (Resolve-Path $candidate).Path
    }

    $repoLocal = Join-Path $RepoRoot "vcpkg\\vcpkg.exe"
    if (Test-Path $repoLocal) {
        return (Resolve-Path $repoLocal).Path
    }

    if ($env:VCPKG_ROOT) {
        $root = $env:VCPKG_ROOT
        $fromRoot = Join-Path $root "vcpkg.exe"
        if (Test-Path $fromRoot) {
            return (Resolve-Path $fromRoot).Path
        }
    }

    $cmd = Get-Command "vcpkg.exe" -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) {
        return $cmd.Source
    }

    throw "vcpkg.exe not found. Install vcpkg and add it to PATH, or set VCPKG_ROOT, or pass -VcpkgExe."
}

$vcpkgExePath = Resolve-VcpkgExePath -RepoRoot $repoRoot -ExplicitVcpkgExe $VcpkgExe

Write-Host "vcpkg exe:     $vcpkgExePath" -ForegroundColor Cyan
Write-Host "manifest root: $manifestRoot" -ForegroundColor Cyan
Write-Host "install root:  $installRoot" -ForegroundColor Cyan
Write-Host "triplet:       $Triplet" -ForegroundColor Cyan

if ($Clean) {
    $repoFull = [System.IO.Path]::GetFullPath($repoRoot)
    $installFull = [System.IO.Path]::GetFullPath($installRoot)
    if (-not $installFull.StartsWith($repoFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to delete install root outside repo: $installFull"
    }

    if (Test-Path $installRoot) {
        Write-Host "Cleaning install root: $installRoot" -ForegroundColor Yellow
        Remove-Item -Path $installRoot -Recurse -Force
    }
}

New-Item -ItemType Directory -Path $installRoot -Force | Out-Null

& $vcpkgExePath install `
    --triplet $Triplet `
    --x-manifest-root $manifestRoot `
    --x-install-root $installRoot

if ($LASTEXITCODE -ne 0) {
    throw "vcpkg install failed with exit code $LASTEXITCODE"
}

$expectedHeader = Join-Path $installRoot "$Triplet\\include\\wil\\com.h"
if (Test-Path $expectedHeader) {
    Write-Host "OK: Found WIL headers at: $expectedHeader" -ForegroundColor Green
} else {
    Write-Host "Warning: Expected WIL header not found at: $expectedHeader" -ForegroundColor Yellow
    Write-Host "Install root contents:" -ForegroundColor Yellow
    Get-ChildItem -Path $installRoot -ErrorAction SilentlyContinue | Select-Object -First 20 Name
}
