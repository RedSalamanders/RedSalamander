<#
.SYNOPSIS
    Install vcpkg dependencies for RedSalamander without polluting the repo root.

.DESCRIPTION
    Runs `vcpkg install` in manifest mode using an isolated staging root per
    triplet, then merges the staged triplet into its platform-private canonical
    install root:
      .build\vcpkg_installed\<platform>\<triplet>\
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

    Discovers vcpkg from the supported locations below, then requires the executable
    to belong to a Git checkout at the commit pinned by vcpkg-tool.json. This keeps
    local dependency installation on the same tool identity as CI. Before acquiring
    dependency locks or creating staging state, every manifest version floor and
    exact override is required to exist in the selected registry version database.

.PARAMETER Platform
    Target platform (x64, ARM64, or All). Default is "All" which installs both platforms.

.PARAMETER Triplet
    Optional explicit vcpkg triplet (overrides -Platform). Use this to install
    a specific x64 or ARM64 triplet such as "x64-windows-static". Other architecture
    prefixes are rejected because they have no reviewed shared-dependency lock scope.

.PARAMETER Clean
    Deletes only the selected triplet directories beneath their platform-scoped
    canonical install and staging roots before installing. With the default
    two-triplet selection, both supported triplets are cleaned. Unselected/custom
    triplets and the other platform's install metadata are preserved.

.PARAMETER VcpkgExe
    Optional explicit path to vcpkg.exe (overrides automatic discovery). Use this if vcpkg
    is installed in a non-standard location or to use a specific vcpkg version.

.PARAMETER VcpkgRoot
    Optional explicit vcpkg ports/version registry root. The directory must contain the
    ports and versions directories of a valid vcpkg checkout. This is passed through the
    command-line --vcpkg-root option so it overrides the executable checkout without
    weakening the independent vcpkg-tool.json executable identity check.

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
    Cleans and reinstalls only the selected x64-windows triplet.

.EXAMPLE
    .\vcpkg-install.ps1 -VcpkgExe C:\tools\vcpkg\vcpkg.exe
    Uses specific vcpkg.exe and installs the default triplets.

.EXAMPLE
    .\vcpkg-install.ps1 -Platform ARM64 -VcpkgExe C:\tools\vcpkg\vcpkg.exe -VcpkgRoot C:\ports\vcpkg
    Uses an independently pinned executable and ports/version registry checkout.

.EXAMPLE
    .\vcpkg-install.ps1 -Help
    Shows this help message.

.NOTES
    Prerequisites: the pinned vcpkg Git checkout and repository build helpers. Side
    effects: acquires shared-dependency locks in stable x64/ARM64 order, downloads or
    builds packages in staging, and merges only selected triplets into
    .build\vcpkg_installed\<platform>. It never races MSBuild for the same platform
    root and never shares mutable vcpkg metadata between x64 and ARM64.
    Primary consumers are local dependency setup and build reproducibility checks.

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

    [Parameter(HelpMessage = "Optional explicit x64 or ARM64 vcpkg triplet (overrides -Platform)")]
    [string]$Triplet = $null,

    [Parameter(HelpMessage = "Delete only selected triplet directories before installing")]
    [switch]$Clean,

    [Parameter(HelpMessage = "Optional explicit path to vcpkg.exe")]
    [string]$VcpkgExe = $null,

    [Parameter(HelpMessage = "Optional explicit vcpkg ports/version registry root")]
    [string]$VcpkgRoot = $null,

    [Parameter(HelpMessage = "Show help and exit")]
    [Alias("h")]
    [switch]$Help
)

$ErrorActionPreference = "Stop"

if ($Help) {
    Get-Help $PSCommandPath -Full
    return
}

Import-Module (Join-Path $PSScriptRoot "Tools\Modules\Build\VcpkgInstallSafety.psm1") -Force -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot "Tools\Modules\Build\VcpkgToolIdentity.psm1") -Force -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot "Tools\Modules\Build\ArtifactOperationLock.psm1") -Force -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot "Tools\Modules\Build\SanitizedEnvironment.psm1") -Force -ErrorAction Stop

$validPlatforms = @('x64', 'ARM64', 'All')
if ($Platform -notin $validPlatforms) {
    Write-Host "Error: Invalid -Platform '$Platform'. Must be one of: $($validPlatforms -join ', ')" -ForegroundColor Red
    Write-Host "Use -Help for usage information." -ForegroundColor Yellow
    exit 1
}

$repoRoot = $PSScriptRoot
$manifestRoot = $repoRoot
$installBaseRoot = Resolve-RSVcpkgSafeChildPath -Root $repoRoot -Child ".build\vcpkg_installed" -Description "vcpkg install base root"
$stagingRoot = Resolve-RSVcpkgSafeChildPath -Root $repoRoot -Child ".build\vcpkg_install_staging" -Description "vcpkg staging root"

if (-not (Test-Path (Join-Path $manifestRoot "vcpkg.json"))) {
    throw "vcpkg.json not found at: $manifestRoot"
}

$vcpkgRootPath = $null
if (-not [string]::IsNullOrWhiteSpace($VcpkgRoot)) {
    if (-not (Test-Path -LiteralPath $VcpkgRoot -PathType Container)) {
        throw "Explicit vcpkg root not found: $VcpkgRoot"
    }
    $vcpkgRootPath = (Resolve-Path -LiteralPath $VcpkgRoot).Path
    foreach ($requiredDirectory in @('ports', 'versions')) {
        $requiredPath = Join-Path $vcpkgRootPath $requiredDirectory
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Container)) {
            throw "Explicit vcpkg root is missing required directory '$requiredDirectory': $vcpkgRootPath"
        }
    }
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

function Get-RSVcpkgCoordinationPlatform {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TripletName
    )

    if ($TripletName -match '^(?i)arm64(?:-|$)') {
        return 'ARM64'
    }
    if ($TripletName -match '^(?i)x64(?:-|$)') {
        return 'x64'
    }

    throw "Triplet '$TripletName' has no reviewed RedSalamander platform lock mapping. Use an x64-* or arm64-* triplet."
}

$coordinationPlatforms = @($tripletsToInstall |
        ForEach-Object { Get-RSVcpkgCoordinationPlatform -TripletName $_ } |
        Sort-Object { if ($_ -eq 'x64') { 0 } else { 1 } } -Unique)

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
$vcpkgToolPin = Get-RSVcpkgToolPin -RepoRoot $repoRoot
$vcpkgExePath = Assert-RSVcpkgToolIdentity -VcpkgExe $vcpkgExePath -ExpectedCommit $vcpkgToolPin.Commit
$vcpkgRegistryRoot = if ($null -ne $vcpkgRootPath) { $vcpkgRootPath } else { Split-Path -Parent $vcpkgExePath }
Assert-RSVcpkgManifestVersionConstraintsAvailable `
    -ManifestPath (Join-Path $manifestRoot 'vcpkg.json') `
    -RegistryRoot $vcpkgRegistryRoot

Write-Host "vcpkg exe:     $vcpkgExePath" -ForegroundColor Cyan
Write-Host "vcpkg commit:  $($vcpkgToolPin.Commit)" -ForegroundColor Cyan
Write-Host "vcpkg root:    $vcpkgRegistryRoot" -ForegroundColor Cyan
Write-Host "manifest root: $manifestRoot" -ForegroundColor Cyan
Write-Host "install base:  $installBaseRoot" -ForegroundColor Cyan
Write-Host "staging root:  $stagingRoot" -ForegroundColor Cyan
Write-Host "triplets:      $($tripletsToInstall -join ', ')" -ForegroundColor Cyan

$sharedDependencyLocks = [System.Collections.Generic.List[object]]::new()
try {
    foreach ($coordinationPlatform in $coordinationPlatforms) {
        $sharedDependencyScope = @{
            kind = 'dependency-install'
            target = ($tripletsToInstall -join ',')
            configuration = 'dependency-install'
            platform = $coordinationPlatform
            coordination = 'shared-dependency'
        }
        $sharedDependencyLock = Enter-RSArtifactOperationLock `
            -RepoRoot $repoRoot `
            -Operation "vcpkg install shared dependency phase $coordinationPlatform" `
            -Scope $sharedDependencyScope
        [void]$sharedDependencyLocks.Add($sharedDependencyLock)
        if ($sharedDependencyLock.WasAbandoned) {
            [void](Set-RSArtifactOperationContaminated `
                    -RepoRoot $repoRoot `
                    -Reason "The previous dependency writer exited while mutating platform $coordinationPlatform vcpkg files." `
                    -AbandonedOwner $sharedDependencyLock.AbandonedOwner `
                    -Scope $sharedDependencyScope)
        }
        if (Test-RSArtifactOperationContaminated -RepoRoot $repoRoot -Scope $sharedDependencyScope) {
            $markerPath = Get-RSArtifactContaminationMarkerPath `
                -RepoRoot $repoRoot `
                -Scope $sharedDependencyScope
            throw "Shared vcpkg dependencies may be incomplete for platform $coordinationPlatform. Repair with a full-solution build.ps1 -Rebuild for that platform. Marker: $markerPath"
        }
        Assert-RSNoResidualArtifactToolProcesses `
            -RepoRoot $repoRoot `
            -Scope $sharedDependencyScope
    }

    New-Item -ItemType Directory -Path $installBaseRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null

foreach ($triplet in $tripletsToInstall) {
    Write-Host ""
    Write-Host "=== Installing $triplet ===" -ForegroundColor Cyan
    $platformScopeName = (Get-RSVcpkgCoordinationPlatform -TripletName $triplet).ToLowerInvariant()
    $platformInstallRoot = Resolve-RSVcpkgSafeChildPath -Root $installBaseRoot -Child $platformScopeName -Description "platform vcpkg install root"
    $tripletStagingRoot = Resolve-RSVcpkgSafeChildPath -Root $stagingRoot -Child $triplet -Description "triplet staging root"
    $destinationTripletPath = Resolve-RSVcpkgSafeChildPath -Root $platformInstallRoot -Child $triplet -Description "destination triplet path"
    New-Item -ItemType Directory -Path $platformInstallRoot -Force | Out-Null
    if ($Clean -and (Test-Path -LiteralPath $destinationTripletPath)) {
        Write-Host "Cleaning selected install: $destinationTripletPath" -ForegroundColor Yellow
        Remove-Item -LiteralPath $destinationTripletPath -Recurse -Force
    }
    if (Test-Path -LiteralPath $tripletStagingRoot) {
        Remove-Item -LiteralPath $tripletStagingRoot -Recurse -Force
    }
    New-Item -ItemType Directory -Path $tripletStagingRoot -Force | Out-Null

    $vcpkgArguments = [System.Collections.Generic.List[string]]::new()
    foreach ($argument in @(
            'install',
            '--triplet', $triplet,
            '--x-manifest-root', $manifestRoot,
            '--x-install-root', $tripletStagingRoot)) {
        [void]$vcpkgArguments.Add($argument)
    }
    if ($null -ne $vcpkgRootPath) {
        [void]$vcpkgArguments.Add("--vcpkg-root=$vcpkgRootPath")
    }

    $vcpkgExitCode = Invoke-RSProcess `
        -FilePath $vcpkgExePath `
        -Arguments $vcpkgArguments.ToArray() `
        -WorkingDirectory $repoRoot

    if ($vcpkgExitCode -ne 0) {
        throw "vcpkg install failed for triplet $triplet with exit code $vcpkgExitCode"
    }

    $stagedTripletPath = Resolve-RSVcpkgSafeChildPath -Root $tripletStagingRoot -Child $triplet -Description "staged triplet output"
    if (-not (Test-Path -LiteralPath $stagedTripletPath -PathType Container)) {
        throw "vcpkg install for $triplet did not produce expected staged directory: $stagedTripletPath"
    }

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
    $platformScopeName = (Get-RSVcpkgCoordinationPlatform -TripletName $triplet).ToLowerInvariant()
    $platformInstallRoot = Resolve-RSVcpkgSafeChildPath -Root $installBaseRoot -Child $platformScopeName -Description "platform vcpkg install root"
    $expectedHeader = Join-Path $platformInstallRoot "$triplet\\include\\wil\\com.h"
    if (Test-Path $expectedHeader) {
        Write-Host "OK: Found WIL headers for $triplet at: $expectedHeader" -ForegroundColor Green
    } else {
        Write-Host "Warning: Expected WIL header for $triplet not found at: $expectedHeader" -ForegroundColor Yellow
        Write-Host "Install root contents:" -ForegroundColor Yellow
        Get-ChildItem -Path $platformInstallRoot -ErrorAction SilentlyContinue | Select-Object -First 20 Name
    }
}
}
finally {
    for ($index = $sharedDependencyLocks.Count - 1; $index -ge 0; $index--) {
        Exit-RSArtifactOperationLock -Lock $sharedDependencyLocks[$index]
    }
}
