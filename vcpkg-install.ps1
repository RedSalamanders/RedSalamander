<#
.SYNOPSIS
    Install vcpkg dependencies for RedSalamander without polluting the repo root.

.DESCRIPTION
    Runs `vcpkg install` in manifest mode using an isolated staging root per
    triplet, then merges the staged triplet into the canonical install root:
      .build\vcpkg_installed\
    Staging roots are created under:
      .build\vcpkg_install_staging\

    This matches the MSBuild/CI layout (see Directory.Build.props) so builds never create
    a top-level `vcpkg_installed\` folder.

    By default (no parameters), installs both supported triplets:
      - x64-windows
      - arm64-windows

    The ASan Debug build configuration links against these same triplets and uses
    _DISABLE_VECTOR_ANNOTATION / _DISABLE_STRING_ANNOTATION to avoid MSVC STL
    annotation mismatches, so no separate ASan-instrumented vcpkg triplet is needed.

    Automatically discovers vcpkg from: explicit -VcpkgExe parameter, repo-local vcpkg\,
    VCPKG_ROOT environment variable, PATH, Visual Studio bundled vcpkg, Chocolatey,
    Scoop, and common installation directories.

.PARAMETER Platform
    Target platform (x64, ARM64, or All). Default is "All" which installs both platforms.

.PARAMETER Triplet
    Optional explicit vcpkg triplet (overrides -Platform). Use this to install
    a specific triplet like "x64-windows-static" or any custom triplet.

.PARAMETER Clean
    Deletes the canonical install root (`.build\vcpkg_installed\`) and staging root
    (`.build\vcpkg_install_staging\`) before installing. Removes all previously installed
    triplets. Applies once before any installations begin.

.PARAMETER VcpkgExe
    Optional explicit path to vcpkg.exe (overrides automatic discovery). Use this if vcpkg
    is installed in a non-standard location or to use a specific vcpkg version.

.PARAMETER Help
    Show this help message and exit. Alias: -h

.EXAMPLE
    .\vcpkg-install.ps1
    Installs both supported triplets (default behavior):
    x64-windows, arm64-windows

.EXAMPLE
    .\vcpkg-install.ps1 -Platform x64
    Installs only x64-windows.

.EXAMPLE
    .\vcpkg-install.ps1 -Triplet x64-windows -Clean
    Cleans install root and installs only x64-windows (explicit triplet override).

.EXAMPLE
    .\vcpkg-install.ps1 -VcpkgExe C:\tools\vcpkg\vcpkg.exe
    Uses specific vcpkg.exe and installs the default triplets.

.EXAMPLE
    .\vcpkg-install.ps1 -Help
    Shows this help message.

.NOTES
    Help can be invoked using standard PowerShell conventions:
      .\vcpkg-install.ps1 -Help
      .\vcpkg-install.ps1 -h
      .\vcpkg-install.ps1 -?
      Get-Help .\vcpkg-install.ps1 -Full
      Get-Help .\vcpkg-install.ps1 -Examples
#>

[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Target platform (x64, ARM64, or All). Default is All.")]
    [string]$Platform = "All",

    [Parameter(HelpMessage = "Optional explicit vcpkg triplet (overrides -Platform)")]
    [string]$Triplet = $null,

    [Parameter(HelpMessage = "Delete .build\\vcpkg_installed and .build\\vcpkg_install_staging before installing")]
    [switch]$Clean,

    [Parameter(HelpMessage = "Optional explicit path to vcpkg.exe")]
    [string]$VcpkgExe = $null,

    [Parameter(HelpMessage = "Show help and exit")]
    [Alias("h")]
    [switch]$Help
)

$ErrorActionPreference = "Stop"

if ($Help) {
    Get-Help $PSCommandPath -Full
    return
}

. (Join-Path $PSScriptRoot "Tools\VcpkgInstallSafety.ps1")

$validPlatforms = @('x64', 'ARM64', 'All')
if ($Platform -notin $validPlatforms) {
    Write-Host "Error: Invalid -Platform '$Platform'. Must be one of: $($validPlatforms -join ', ')" -ForegroundColor Red
    Write-Host "Use -Help for usage information." -ForegroundColor Yellow
    exit 1
}

$repoRoot = $PSScriptRoot
$manifestRoot = $repoRoot
$installRoot = Resolve-RSVcpkgSafeChildPath -Root $repoRoot -Child ".build\vcpkg_installed" -Description "vcpkg install root"
$stagingRoot = Resolve-RSVcpkgSafeChildPath -Root $repoRoot -Child ".build\vcpkg_install_staging" -Description "vcpkg staging root"

if (-not (Test-Path (Join-Path $manifestRoot "vcpkg.json"))) {
    throw "vcpkg.json not found at: $manifestRoot"
}

$tripletsToInstall = @()

if ($Triplet) {
    $tripletsToInstall = @($Triplet)
} elseif ($Platform -eq "All") {
    $tripletsToInstall = @("x64-windows", "arm64-windows")
} else {
    $arch = if ($Platform -eq "ARM64") { "arm64" } else { "x64" }
    $tripletsToInstall = @("$arch-windows")
}

$tripletsToInstall = @($tripletsToInstall | ForEach-Object { Assert-RSVcpkgTripletLeafName -Triplet $_ })

function Resolve-VcpkgExePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $false)]
        [string]$ExplicitVcpkgExe
    )

    if ($ExplicitVcpkgExe) {
        $candidate = $ExplicitVcpkgExe
        if (-not (Test-Path $candidate -PathType Leaf)) {
            throw "VcpkgExe not found: $candidate"
        }
        Write-Verbose "vcpkg discovered at: $candidate (source: explicit -VcpkgExe parameter)"
        return (Resolve-Path $candidate).Path
    }

    $repoLocal = Join-Path $RepoRoot "vcpkg\vcpkg.exe"
    if (Test-Path $repoLocal -PathType Leaf) {
        Write-Verbose "vcpkg discovered at: $repoLocal (source: repo-local)"
        return (Resolve-Path $repoLocal).Path
    }

    if ($env:VCPKG_ROOT) {
        $root = $env:VCPKG_ROOT
        $fromRoot = Join-Path $root "vcpkg.exe"
        if (Test-Path $fromRoot -PathType Leaf) {
            Write-Verbose "vcpkg discovered at: $fromRoot (source: VCPKG_ROOT environment variable)"
            return (Resolve-Path $fromRoot).Path
        }
    }

    $cmd = Get-Command "vcpkg.exe" -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) {
        Write-Verbose "vcpkg discovered at: $($cmd.Source) (source: PATH)"
        return $cmd.Source
    }

    $vswherePath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswherePath -PathType Leaf) {
        try {
            $vsPath = & $vswherePath -latest -prerelease -products * -requires Microsoft.VisualStudio.Component.Vcpkg -property installationPath 2>$null
            if ($vsPath) {
                $vsVcpkg = Join-Path $vsPath "VC\vcpkg\vcpkg.exe"
                if (Test-Path $vsVcpkg -PathType Leaf) {
                    Write-Verbose "vcpkg discovered at: $vsVcpkg (source: Visual Studio bundled)"
                    return (Resolve-Path $vsVcpkg).Path
                }
            }
        } catch {
        }
    }

    $chocoPath1 = "C:\tools\vcpkg\vcpkg.exe"
    if (Test-Path $chocoPath1 -PathType Leaf) {
        Write-Verbose "vcpkg discovered at: $chocoPath1 (source: Chocolatey)"
        return (Resolve-Path $chocoPath1).Path
    }

    if ($env:ChocolateyInstall) {
        $chocoPath2 = Join-Path $env:ChocolateyInstall "lib\vcpkg\tools\vcpkg.exe"
        if (Test-Path $chocoPath2 -PathType Leaf) {
            Write-Verbose "vcpkg discovered at: $chocoPath2 (source: Chocolatey)"
            return (Resolve-Path $chocoPath2).Path
        }
    }

    $scoopPath1 = Join-Path $env:USERPROFILE "scoop\apps\vcpkg\current\vcpkg.exe"
    if (Test-Path $scoopPath1 -PathType Leaf) {
        Write-Verbose "vcpkg discovered at: $scoopPath1 (source: Scoop)"
        return (Resolve-Path $scoopPath1).Path
    }

    if ($env:SCOOP) {
        $scoopPath2 = Join-Path $env:SCOOP "apps\vcpkg\current\vcpkg.exe"
        if (Test-Path $scoopPath2 -PathType Leaf) {
            Write-Verbose "vcpkg discovered at: $scoopPath2 (source: Scoop)"
            return (Resolve-Path $scoopPath2).Path
        }
    }

    $commonRoots = @(
        "C:\vcpkg\vcpkg.exe",
        "C:\dev\vcpkg\vcpkg.exe",
        (Join-Path $env:USERPROFILE "vcpkg\vcpkg.exe"),
        (Join-Path $env:USERPROFILE "source\vcpkg\vcpkg.exe")
    )

    foreach ($commonPath in $commonRoots) {
        if (Test-Path $commonPath -PathType Leaf) {
            Write-Verbose "vcpkg discovered at: $commonPath (source: common installation directory)"
            return (Resolve-Path $commonPath).Path
        }
    }

    throw "vcpkg.exe not found. Install vcpkg and add it to PATH, or set VCPKG_ROOT, or pass -VcpkgExe. Auto-discovery checks: repo-local, VCPKG_ROOT, PATH, Visual Studio bundled, Chocolatey, Scoop, and common installation directories."
}

$vcpkgExePath = Resolve-VcpkgExePath -RepoRoot $repoRoot -ExplicitVcpkgExe $VcpkgExe

Write-Host "vcpkg exe:     $vcpkgExePath" -ForegroundColor Cyan
Write-Host "manifest root: $manifestRoot" -ForegroundColor Cyan
Write-Host "install root:  $installRoot" -ForegroundColor Cyan
Write-Host "staging root:  $stagingRoot" -ForegroundColor Cyan
Write-Host "triplets:      $($tripletsToInstall -join ', ')" -ForegroundColor Cyan

if ($Clean) {
    $repoFull = [System.IO.Path]::GetFullPath($repoRoot).TrimEnd([char]92)
    foreach ($pathToClean in @($installRoot, $stagingRoot)) {
        $cleanFull = [System.IO.Path]::GetFullPath($pathToClean).TrimEnd([char]92)
        if (($cleanFull -eq $repoFull) -or (-not $cleanFull.StartsWith("$repoFull\", [System.StringComparison]::OrdinalIgnoreCase))) {
            throw "Refusing to delete path outside repo: $cleanFull"
        }

        if (Test-Path -LiteralPath $pathToClean) {
            Write-Host "Cleaning: $pathToClean" -ForegroundColor Yellow
            Remove-Item -LiteralPath $pathToClean -Recurse -Force
        }
    }
}

New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null

foreach ($triplet in $tripletsToInstall) {
    Write-Host ""
    Write-Host "=== Installing $triplet ===" -ForegroundColor Cyan
    $tripletStagingRoot = Resolve-RSVcpkgSafeChildPath -Root $stagingRoot -Child $triplet -Description "triplet staging root"
    if (Test-Path -LiteralPath $tripletStagingRoot) {
        Remove-Item -LiteralPath $tripletStagingRoot -Recurse -Force
    }
    New-Item -ItemType Directory -Path $tripletStagingRoot -Force | Out-Null

    & $vcpkgExePath install `
        --triplet $triplet `
        --x-manifest-root $manifestRoot `
        --x-install-root $tripletStagingRoot

    if ($LASTEXITCODE -ne 0) {
        throw "vcpkg install failed for triplet $triplet with exit code $LASTEXITCODE"
    }

    $stagedTripletPath = Resolve-RSVcpkgSafeChildPath -Root $tripletStagingRoot -Child $triplet -Description "staged triplet output"
    if (-not (Test-Path -LiteralPath $stagedTripletPath -PathType Container)) {
        throw "vcpkg install for $triplet did not produce expected staged directory: $stagedTripletPath"
    }

    $destinationTripletPath = Resolve-RSVcpkgSafeChildPath -Root $installRoot -Child $triplet -Description "destination triplet path"

    if (Test-Path -LiteralPath $destinationTripletPath) {
        Write-Host "Merging $triplet into existing installation: $destinationTripletPath" -ForegroundColor Cyan
        Merge-RSVcpkgTripletSafe -SourcePath $stagedTripletPath -DestinationPath $destinationTripletPath -TripletName $triplet
    } else {
        Write-Host "Installing $triplet (new installation): $destinationTripletPath" -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $destinationTripletPath -Force | Out-Null
        Copy-Item -Path (Join-Path $stagedTripletPath '*') -Destination $destinationTripletPath -Recurse -Force
        Write-Host "Installed $triplet into: $destinationTripletPath" -ForegroundColor Green
    }
}

foreach ($triplet in $tripletsToInstall) {
    $expectedHeader = Join-Path $installRoot "$triplet\\include\\wil\\com.h"
    if (Test-Path $expectedHeader) {
        Write-Host "OK: Found WIL headers for $triplet at: $expectedHeader" -ForegroundColor Green
    } else {
        Write-Host "Warning: Expected WIL header for $triplet not found at: $expectedHeader" -ForegroundColor Yellow
        Write-Host "Install root contents:" -ForegroundColor Yellow
        Get-ChildItem -Path $installRoot -ErrorAction SilentlyContinue | Select-Object -First 20 Name
    }
}
