<#
.SYNOPSIS
    Build RedSalamander solution
.DESCRIPTION
    Builds the entire RedSalamander solution in Debug (default), Release, or ASan Debug configuration.
    Optionally builds one ProjectName or an explicit ProjectNames union.
    Captures native artifacts and provenance, optionally produces a portable ZIP,
    and verifies those bytes before publishing a successful build receipt.
.PARAMETER Configuration
    Build configuration: Debug, Release, or ASan Debug (default: Debug)
.PARAMETER Platform
    Target platform: x64 or ARM64 (default: x64)
.PARAMETER ProjectName
    Specific project to build. If not specified, builds the entire solution.
    Examples: "RedSalamander", "Common"
.PARAMETER ProjectNames
    Native solution project names to build in one evaluated solution operation.
    Includes their project-reference and runtime deployment closure. Publishes a
    selected-projects receipt, which cannot satisfy a full-solution test/package gate.
    Cannot be combined with ProjectName, ProductionOnly or release/package actions.
.PARAMETER ProductionOnly
    Build production projects without Tests or PoC (Release only); cannot be combined with Zip.
    Resolve every selected project through its full MSBuild solution-folder target name.
.PARAMETER Clean
    Perform a clean build
.PARAMETER Rebuild
    Rebuild all projects
.PARAMETER MaxCpuCount
    Optional maximum MSBuild worker count. Zero uses MSBuild's default.
.PARAMETER Zip
    Build a portable ZIP package after a successful build (Release only)
.PARAMETER GenerateWingetManifest
    Generate winget manifests from the x64 and ARM64 portable ZIP packages (Release only).
.PARAMETER BuildNumber
    Override the build number shared by all modules in the current build invocation.
    Required for local official release builds. In CI, GITHUB_RUN_NUMBER is used automatically.
.PARAMETER OfficialRelease
    Stamp binaries as an official release build (no local beta Debug/Release suffix in displayed versions).
    Official release mode requires -BuildNumber or a CI environment with GITHUB_RUN_NUMBER, and a
    test-disabled build: -Configuration Release without RSBuildEnableTests=true.
.PARAMETER MonitorDiagnostics
    Deprecated. Diagnostics are enabled at runtime with --etw and perf JSONL with --perf.
.PARAMETER DxUiSdkInputDescriptor
    Optional independently retained pinned qualified SDK input descriptor below repository .build.
    Release only; keeps actual compiler/ABI, source/API, linker and reader-lifetime checks.
.EXAMPLE
    .\build.ps1
    Builds entire solution in Debug configuration
.EXAMPLE
    .\build.ps1 -Configuration Release
    Builds entire solution in Release configuration
.EXAMPLE
    .\build.ps1 -ProjectName Common
    Builds only Common project
.EXAMPLE
    .\build.ps1 -Configuration Release -Clean
    Clean build of entire solution in Release configuration
.EXAMPLE
    .\build.ps1 -Zip
    Builds entire solution in Release and produces a portable ZIP package
.EXAMPLE
    .\build.ps1 -Configuration Release -BuildNumber 183 -OfficialRelease
    Builds an official release-stamped binary set using CI/release build number 183
#>

[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Build configuration (Debug or Release)")]
    [ValidateSet("Debug", "Release", "ASan Debug")]
    [string]$Configuration = "Debug",

    [Parameter(HelpMessage = "Target platform (x64 or ARM64)")]
    [ValidateSet("x64", "ARM64")]
    [string]$Platform = "x64",

    [Parameter(HelpMessage = "Specific project to build (builds entire solution if not specified)")]
    [string]$ProjectName = $null,

    [Parameter(HelpMessage = "Explicit native solution project union")]
    [ValidateNotNullOrEmpty()]
    [string[]]$ProjectNames = @(),

    [Parameter(HelpMessage = "Build the production solution projects, excluding Tests and PoC (Release only)")]
    [switch]$ProductionOnly,

    [Parameter(HelpMessage = "Perform a clean build")]
    [switch]$Clean,

    [Parameter(HelpMessage = "Rebuild all projects")]
    [switch]$Rebuild,

    [Parameter(HelpMessage = "Maximum MSBuild worker count (0 uses the default)")]
    [ValidateRange(0, 256)]
    [int]$MaxCpuCount = 0,

    [Parameter(HelpMessage = "Build ZIP package after build (Release only)")]
    [switch]$Zip,

    [Parameter(HelpMessage = "Generate winget manifest after build (Release only)")]
    [switch]$GenerateWingetManifest,

    [Parameter(HelpMessage = "Override the build number (required for local official release builds)")]
    [int]$BuildNumber = 0,

    [Parameter(HelpMessage = "Stamp binaries as an official release build (requires -BuildNumber or CI)")]
    [switch]$OfficialRelease,

    [Parameter(HelpMessage = "Deprecated; use the runtime --etw flag instead")]
    [switch]$MonitorDiagnostics,

    [string]$DxUiSdkInputDescriptor = ''
)

$ErrorActionPreference = "Stop"
if ($ProjectNames.Count -and ($ProjectName -or $ProductionOnly -or $Zip -or $GenerateWingetManifest -or $OfficialRelease)) {
    throw 'ProjectNames cannot be combined with ProjectName, ProductionOnly or release/package actions.'
}
if ($DxUiSdkInputDescriptor -and ($Configuration -cne 'Release' -or $Clean)) {throw 'Qualified SDK inputs require a Release build without Clean.'}
if ($DxUiSdkInputDescriptor -and ($Zip -or $GenerateWingetManifest)) {throw 'Build SDK-backed native outputs first; package them separately after their build receipt is sealed.'}
if ($env:CI_NATIVE_BUILDS_DISABLED -eq 'true') { throw 'Native builds are forbidden in compiler-free PR checks.' }
if ($ProductionOnly -and ($Configuration -ne 'Release' -or $ProjectName -or $Clean)) {
    throw 'ProductionOnly requires Release and cannot be combined with ProjectName or Clean.'
}
if ($ProductionOnly -and $Zip) {
    throw 'ProductionOnly cannot create a ZIP because it excludes the package smoke-test host. Use a full Release build with -Zip.'
}
$nativeEntryElapsed = [Diagnostics.Stopwatch]::StartNew()
Import-Module (Join-Path $PSScriptRoot 'Build/BuildConsole.psm1') -ErrorAction Stop
$BuildProjectSelectionModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\BuildProjectSelection.psm1"
$BuildEvidenceModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\BuildEvidence.psm1"
$MSBuildInvocationModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\MSBuildInvocation.psm1"
$ProcessStreamingModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\ProcessStreaming.psm1"
$SanitizedEnvironmentModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\SanitizedEnvironment.psm1"
$ArtifactOperationLockModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\ArtifactOperationLock.psm1"
$VcpkgInstallSafetyModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\VcpkgInstallSafety.psm1"
$ReleaseArtifactPolicyModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Packaging\ReleaseArtifactPolicy.psm1"
if (-not (Test-Path $BuildProjectSelectionModule)) {
    Write-Error "Build project selection helper not found: $BuildProjectSelectionModule"
    exit 1
}
if (-not (Test-Path $BuildEvidenceModule)) {
    Write-Error "Build evidence helper not found: $BuildEvidenceModule"
    exit 1
}
if (-not (Test-Path $MSBuildInvocationModule)) {
    Write-Error "MSBuild invocation helper not found: $MSBuildInvocationModule"
    exit 1
}
if (-not (Test-Path $ProcessStreamingModule)) {
    Write-Error "Process streaming helper not found: $ProcessStreamingModule"
    exit 1
}
if (-not (Test-Path $SanitizedEnvironmentModule)) {
    Write-Error "Sanitized environment helper not found: $SanitizedEnvironmentModule"
    exit 1
}
if (-not (Test-Path $ArtifactOperationLockModule)) {
    Write-Error "Artifact operation lock helper not found: $ArtifactOperationLockModule"
    exit 1
}
if (-not (Test-Path $VcpkgInstallSafetyModule)) {
    Write-Error "vcpkg install safety helper not found: $VcpkgInstallSafetyModule"
    exit 1
}
if (-not (Test-Path $ReleaseArtifactPolicyModule)) {
    Write-Error "Release artifact policy helper not found: $ReleaseArtifactPolicyModule"
    exit 1
}

Import-Module $BuildProjectSelectionModule -Force -ErrorAction Stop
Import-Module $BuildEvidenceModule -Force -ErrorAction Stop
Import-Module $MSBuildInvocationModule -Force -ErrorAction Stop
Import-Module $SanitizedEnvironmentModule -Force -ErrorAction Stop
Import-Module $ProcessStreamingModule -Force -ErrorAction Stop
Import-Module $ArtifactOperationLockModule -Force -ErrorAction Stop
Import-Module $VcpkgInstallSafetyModule -Force -ErrorAction Stop
Import-Module $ReleaseArtifactPolicyModule -Force -ErrorAction Stop

function Test-InteractiveTerminal {
    try {
        $canReadWindowTitle = $true
        if ($null -ne $Host -and $null -ne $Host.UI -and $null -ne $Host.UI.RawUI) {
            try {
                $null = $Host.UI.RawUI.WindowTitle
            }
            catch {
                $canReadWindowTitle = $false
            }
        }

        return Test-RSInteractiveTerminal `
            -IsOutputRedirected ([Console]::IsOutputRedirected) `
            -IsErrorRedirected ([Console]::IsErrorRedirected) `
            -HasRawUi ($null -ne $Host -and $null -ne $Host.UI -and $null -ne $Host.UI.RawUI) `
            -CanReadWindowTitle $canReadWindowTitle
    }
    catch {
        Write-Verbose "Interactive terminal detection failed: $($_.Exception.Message)"
        return $false
    }
}

function New-ProcessLogPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prefix
    )

    $safePrefix = if ([string]::IsNullOrWhiteSpace($Prefix)) {
        "process"
    }
    else {
        ($Prefix -replace '[^A-Za-z0-9._-]', '_')
    }

    $logDir = Join-Path $PSScriptRoot ".build\logs"
    [void](New-Item -ItemType Directory -Path $logDir -Force)

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
    $uniqueSuffix = [Guid]::NewGuid().ToString('N').Substring(0, 8)
    return Join-Path $logDir "$safePrefix-$timestamp-pid$PID-$uniqueSuffix.log"
}

$script:UseInteractiveTerminal = Test-InteractiveTerminal
$script:LastProcessLogPath = $null

function Write-BuildBanner {
    $art = @(
        '   ____          _ ____        _                                _           ',
        '  |  _ \ ___  __| / ___|  __ _| | __ _ _ __ ___   __ _ _ __   __| | ___ _ __ ',
        '  | |_) / _ \/ _` \___ \ / _` | |/ _` | ''_ ` _ \ / _` | ''_ \ / _` |/ _ \ ''__|',
        '  |  _ <  __/ (_| |___) | (_| | | (_| | | | | | | (_| | | | | (_| |  __/ |   ',
        '  |_| \_\___|\__,_|____/ \__,_|_|\__,_|_| |_| |_|\__,_|_| |_|\__,_|\___|_|   '
    )

    if ($script:UseInteractiveTerminal) {
        $colors = @("DarkCyan", "Cyan", "Green", "Yellow", "Magenta")
        for ($i = 0; $i -lt $art.Count; ++$i) {
            Write-Host $art[$i] -ForegroundColor $colors[$i]
        }
        Write-Host "                     Build, package, and validate with confidence" -ForegroundColor DarkGray
    }
    else {
        foreach ($line in $art) {
            Write-Host $line
        }
        Write-Host "                     Build, package, and validate with confidence"
    }

    Write-Host ""
}

# Validate packaging options early so we fail fast before attempting a build.


if ($Zip -and $ProjectName) {
    Write-Error "Packaging requires building the full solution. Remove -ProjectName."
    exit 1
}

$packageMode = if ($Zip) { "ZIP" } else { "None" }

if ($Zip) {
    if (-not $PSBoundParameters.ContainsKey("Configuration")) {
        $Configuration = "Release"
    } elseif ($Configuration -ne "Release") {
        Write-Error "Packaging requires -Configuration Release."
        exit 1
    }
}

if ($GenerateWingetManifest -and $Configuration -ne "Release") {
    Write-Error "Winget manifest generation requires -Configuration Release."
    exit 1
}

# A test-enabled (ENABLE_TESTS) build is never distributable, so nothing that packages or stamps
# an official release may build test-enabled: neither by override nor by the non-Release default
# that the effective tests setting below applies.
$requestedDistributableTests = [Environment]::GetEnvironmentVariable('RSBuildEnableTests', 'Process')
if (($Zip -or $GenerateWingetManifest -or $OfficialRelease) -and
    ($requestedDistributableTests -ieq 'true' -or
     ([string]::IsNullOrWhiteSpace($requestedDistributableTests) -and $Configuration -ne 'Release'))) {
    Write-Error "Packaging and -OfficialRelease require a test-disabled build (Release, RSBuildEnableTests not 'true'); a test-enabled build is never distributable."
    exit 1
}

# Script constants
$SolutionFile = Join-Path -Path $PSScriptRoot -ChildPath "RedSalamander.sln"
$VersioningModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\Versioning.psm1"

Write-BuildBanner

# Validate solution file exists
if (-not (Test-Path $SolutionFile)) {
    Write-Error "Solution file not found: $SolutionFile"
    exit 1
}

if (-not (Test-Path $VersioningModule)) {
    Write-Error "Version helper module not found: $VersioningModule"
    exit 1
}

Import-Module $VersioningModule -Force -ErrorAction Stop

$SolutionFullPath = (Resolve-Path $SolutionFile).Path
$SolutionDir = (Split-Path -Parent $SolutionFullPath)
$SolutionDirWithSlash = $SolutionDir.TrimEnd('\') + '\'
$unionSelection = if ($ProjectNames.Count) {
    Get-RSBuildProjectUnion -SolutionPath $SolutionFullPath -ProjectNames $ProjectNames -Rebuild:$Rebuild
} else { $null }
# Partial rebuilds must never acquire full-solution contamination repair rights.
$repairProjectName = if ($null -ne $unionSelection) { $unionSelection.ReceiptTarget } else { $ProjectName }
$artifactOperationLock = $null
$sdkInput = $null
$sdkModule = $null
$sdkBuildFailure = $null
$sdkProperties = @()
$stopwatch = [System.Diagnostics.Stopwatch]::new()
$contaminationRepairAuthorized = $false
$script:sharedDependencyRepairAuthorized = $false
try {
    $operationTarget = if ($null -ne $unionSelection) { $unionSelection.ReceiptTarget } elseif ($ProjectName) { $ProjectName } else { 'solution' }
    $artifactOperationScope = @{
        kind = 'build'
        target = $operationTarget
        configuration = $Configuration
        platform = $Platform
    }
    $artifactOperationLock = Enter-RSArtifactOperationLock `
        -RepoRoot $SolutionDir `
        -Operation "build $operationTarget $Configuration|$Platform" `
        -Scope $artifactOperationScope

    if ($artifactOperationLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $SolutionDir `
                -Reason "The previous build/test owner exited without clearing the exclusive artifact-operation lock." `
                -AbandonedOwner $artifactOperationLock.AbandonedOwner `
                -Scope $artifactOperationScope)
    }

    $contamination = Read-RSArtifactOperationContamination `
        -RepoRoot $SolutionDir `
        -Scope $artifactOperationScope
    if ($null -ne $contamination) {
        $contaminationRepairAuthorized = Test-RSArtifactOperationRepairAllowed `
            -Contamination $contamination `
            -Rebuild:$Rebuild `
            -ProjectName $repairProjectName `
            -Configuration $Configuration `
            -Platform $Platform
        if (-not $contaminationRepairAuthorized) {
            $markerPath = Get-RSArtifactContaminationMarkerPath `
                -RepoRoot $SolutionDir `
                -Scope $artifactOperationScope
            $scopeText = if ($null -ne $contamination.PSObject.Properties['abandoned_operation'] -and
                -not [string]::IsNullOrWhiteSpace([string]$contamination.abandoned_operation)) {
                " Previous operation: '$($contamination.abandoned_operation)'."
            } else { '' }
            throw ("Shared build artifacts may be mixed after an interrupted operation." +
                   "$scopeText Run a full-solution build.ps1 -Rebuild with the recorded configuration and platform; " +
                   "-Clean alone and targeted or wrong-scope rebuilds cannot clear this marker. Marker: $markerPath")
        }
    }

    Assert-RSNoResidualArtifactToolProcesses `
        -RepoRoot $SolutionDir `
        -Scope $artifactOperationScope

$versionContext = Resolve-RSVersionContext `
    -RepoRoot $SolutionDir `
    -Configuration $Configuration `
    -Platform $Platform `
    -BuildNumber $BuildNumber `
    -OfficialRelease:$OfficialRelease
$versionStatePath = Get-RSVersionStatePath -RepoRoot $SolutionDir

# Function to find MSBuild
function Get-RSHostOrderedMSBuildRelativePaths {
    # Bin\MSBuild.exe is the 32-bit x86 MSBuild. The VC toolset derives the compiler host from the
    # MSBuild process architecture and downgrades arm64/x64 hosts it cannot see, so every discovery
    # strategy prefers the native 64-bit MSBuild of this host, then every other 64-bit executable
    # the host can run (ARM64 Windows emulates x64; x64 Windows cannot run arm64\), and only then
    # an x86 executable.
    $isArm64Host = [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64'
    [string[]]$candidates = if ($isArm64Host) {
        @('MSBuild\Current\Bin\arm64\MSBuild.exe', 'MSBuild\Current\Bin\amd64\MSBuild.exe')
    } else {
        @('MSBuild\Current\Bin\amd64\MSBuild.exe')
    }
    return $candidates + [string[]]@(
        'MSBuild\15.0\Bin\amd64\MSBuild.exe',
        'MSBuild\Current\Bin\MSBuild.exe',
        'MSBuild\15.0\Bin\MSBuild.exe'
    )
}

function Resolve-RSHostOrderedMSBuildPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    # A PATH hit from a developer prompt is usually the x86 MSBuild\Current\Bin\MSBuild.exe; map
    # it to the best host-ordered executable of the same installation when one exists.
    $match = [regex]::Match($Path, '^(?<root>.+?)\\MSBuild\\(?:Current|15\.0)\\Bin\\(?:(?:amd64|arm64)\\)?MSBuild\.exe$', 'IgnoreCase')
    if (-not $match.Success) { return $Path }
    foreach ($relative in Get-RSHostOrderedMSBuildRelativePaths) {
        $candidate = Join-Path $match.Groups['root'].Value $relative
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $Path
}

function Find-MSBuild {
    Write-Host "Locating MSBuild..." -ForegroundColor Yellow

    # In GitHub Actions, consume the exact workflow-selected path first so the cache-key
    # identity probe and the build receipt cannot resolve different MSBuild executables.
    if ($env:GITHUB_ACTIONS -and ($env:GITHUB_ACTIONS -eq "true")) {
        if (-not [string]::IsNullOrWhiteSpace($env:MSBUILD_EXE_PATH)) {
            $selectedPath = [IO.Path]::GetFullPath($env:MSBUILD_EXE_PATH)
            if (-not (Test-Path -LiteralPath $selectedPath -PathType Leaf)) {
                throw "Workflow-selected MSBuild does not exist: $selectedPath"
            }
            $selectedVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($selectedPath)
            # A registered VS instance can live under RUNNER_TEMP or another custom
            # root. Validate the binary identity and MSBuild layout, not its parent name.
            if ($selectedVersion.FileMajorPart -lt 18 -or
                $selectedVersion.OriginalFilename -ne 'MSBuild.exe' -or
                $selectedPath -notmatch '\\MSBuild\\Current\\Bin\\(?:(?:amd64|arm64)\\)?MSBuild\.exe$') {
                throw "Workflow-selected MSBuild is not a Visual Studio 2026 executable: $selectedPath"
            }
            return @{
                Path = $selectedPath
                Version = "MSBuild $($selectedVersion.FileMajorPart) (workflow-selected)"
                Method = "MSBUILD_EXE_PATH"
            }
        }

        # Compatibility fallback for callers that only prepend the selected directory.
        $msbuildInPath = Get-Command msbuild.exe -ErrorAction SilentlyContinue
        if ($msbuildInPath -and $msbuildInPath.Source -and (Test-Path $msbuildInPath.Source)) {
            $candidatePath = Resolve-RSHostOrderedMSBuildPath -Path $msbuildInPath.Source
            $fileMajor = $null
            try {
                $ver = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($candidatePath)
                $fileMajor = $ver.FileMajorPart
            }
            catch {
                $fileMajor = $null
            }

            # VS 2026 MSBuild should report major version 18.
            if (($null -ne $fileMajor) -and ($fileMajor -ge 18) -and $ver.OriginalFilename -eq 'MSBuild.exe' -and
                $candidatePath -match '\\MSBuild\\Current\\Bin\\(?:(?:amd64|arm64)\\)?MSBuild\.exe$') {
                return @{
                    Path = $candidatePath
                    Version = "MSBuild $fileMajor (PATH)"
                    Method = "PATH"
                }
            }
        }
    }

    # Strategy 1: Try vswhere (preferred method for VS 2017+)
    $vswherePaths = @(
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\Installer\vswhere.exe"
    )

    $vswhere = $vswherePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($vswhere) {
        Write-Host "  Found vswhere: $vswhere" -ForegroundColor Gray

        # Prefer the newest Visual Studio instance that has MSBuild on disk.
        # Important for CI: the hosted runner may have VS 2022 preinstalled, but we may install newer Build Tools.
        $instances = @()
        try {
            $instancesJson = & $vswhere -all -products "*" -prerelease -format json 2>$null
            if ($LASTEXITCODE -eq 0 -and $instancesJson) {
                $instances = @($instancesJson | ConvertFrom-Json)
            }
        }
        catch {
            $instances = @()
        }

        $best = $null
        foreach ($instance in $instances) {
            $installPath = $instance.installationPath
            if (-not $installPath) {
                continue
            }

            $installVersion = [version]"0.0"
            try {
                if ($instance.installationVersion) {
                    $installVersion = [version]$instance.installationVersion
                }
            }
            catch {
                $installVersion = [version]"0.0"
            }

            $msbuildCandidates = @(Get-RSHostOrderedMSBuildRelativePaths | ForEach-Object { Join-Path $installPath $_ })

            $msbuildPath = $msbuildCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
            if (-not $msbuildPath) {
                continue
            }

            if (-not $best -or $installVersion -gt $best.InstallVersion) {
                $best = @{
                    InstallVersion = $installVersion
                    DisplayName = $instance.displayName
                    InstallationVersion = $instance.installationVersion
                    Path = $msbuildPath
                }
            }
        }

        if ($best) {
            $displayName = if ($best.DisplayName) { $best.DisplayName } else { "Visual Studio" }
            $versionText = if ($best.InstallationVersion) { "$displayName ($($best.InstallationVersion))" } else { $displayName }
            return @{
                Path = $best.Path
                Version = $versionText
                Method = "vswhere"
            }
        }
    }

    # Strategy 2: Search common Visual Studio installation paths
    Write-Host "  Searching Visual Studio installation paths..." -ForegroundColor Gray

    $vsYears = @("2026", "2022")
    $vsEditions = @("Enterprise", "Professional", "Community", "BuildTools")
    $basePaths = @(
        "${env:ProgramFiles}\Microsoft Visual Studio",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio"
    )

    foreach ($basePath in $basePaths) {
        foreach ($year in $vsYears) {
            foreach ($edition in $vsEditions) {
                $msbuildPaths = @(Get-RSHostOrderedMSBuildRelativePaths | ForEach-Object { "$basePath\$year\$edition\$_" })

                foreach ($msbuildPath in $msbuildPaths) {
                    if (Test-Path $msbuildPath) {
                        return @{
                            Path = $msbuildPath
                            Version = "Visual Studio $year $edition"
                            Method = "path search"
                        }
                    }
                }
            }
        }
    }

    # Strategy 3: Search PATH environment variable
    Write-Host "  Searching PATH environment variable..." -ForegroundColor Gray

    $msbuildInPath = Get-Command msbuild.exe -ErrorAction SilentlyContinue
    if ($msbuildInPath) {
        return @{
            Path = Resolve-RSHostOrderedMSBuildPath -Path $msbuildInPath.Source
            Version = "Found in PATH"
            Method = "PATH"
        }
    }

    # Strategy 4: Use Developer Command Prompt environment
    Write-Host "  Checking Developer Command Prompt environment..." -ForegroundColor Gray

    if ($env:VSINSTALLDIR) {
        $devMSBuildPaths = @(Get-RSHostOrderedMSBuildRelativePaths | ForEach-Object { "$env:VSINSTALLDIR\$_" })

        foreach ($devPath in $devMSBuildPaths) {
            if (Test-Path $devPath) {
                return @{
                    Path = $devPath
                    Version = "Developer Command Prompt"
                    Method = "VSINSTALLDIR"
                }
            }
        }
    }

    return $null
}


# Script-scope state is initialized before use: Run-AllTests and Test-Changes invoke this script under Set-StrictMode -Version Latest.
$script:vcpkgOverlayTripletRefreshed = $false

function Update-RSBuildVcpkgOverlayTriplet {
    # MSBuild's manifest install runs whenever vcpkg.json is newer than its stamp, and Directory.Build.props passes it the overlay
    # triplet vcpkg-install.ps1 writes. Refresh that file here the same way (same installation, its default MSVC toolset, same
    # path), so neither installer falls back to vcpkg's newest toolset. Once per build, inside the platform dependency lock.
    if ($script:vcpkgOverlayTripletRefreshed) {
        return
    }
    $script:vcpkgOverlayTripletRefreshed = $true
    if ($env:VcpkgManifestInstall -eq 'false') {
        # The caller (CI) installs with the governed installer first and disables MSBuild's install.
        return
    }

    $triplet = if ($Platform -eq 'ARM64') { 'arm64-windows' } else { 'x64-windows' }
    $visualStudioInstallation = Get-RSVcpkgVisualStudioInstallation
    # The vcpkg root MSBuild's integration installs with: an explicit VcpkgRoot, else the installation's bundled vcpkg. Their
    # x64-windows and arm64-windows triplets match the pinned tool's byte for byte, so both installers write the same file.
    $stockTriplet = @($env:VcpkgRoot, (Join-Path $visualStudioInstallation 'VC\vcpkg')) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { Join-Path $_ "triplets\$triplet.cmake" } |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
        Select-Object -First 1
    if (-not $stockTriplet) {
        # No vcpkg MSBuild integration to install with; an existing overlay from vcpkg-install.ps1 stays as it is.
        return
    }

    $msvcToolset = Get-RSVcpkgDefaultToolset -Installation $visualStudioInstallation
    $overlayDirectory = Resolve-RSVcpkgSafeChildPath -Root $SolutionDir -Child ".build\vcpkg-triplets\$($Platform.ToLowerInvariant())" -Description "vcpkg triplet overlay directory"
    $overlay = Update-RSVcpkgOverlayTriplet -StockTripletPath $stockTriplet -Toolset $msvcToolset -OutputDirectory $overlayDirectory
    if ($overlay.Changed) {
        Write-Host "vcpkg overlay triplet $($overlay.Path) written: MSVC $($msvcToolset.MajorMinor) of $visualStudioInstallation" -ForegroundColor Cyan
    }
}

function Invoke-MSBuild {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MSBuildPath,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [string]$WorkingDirectory = $PSScriptRoot
    )

    $sharedDependencyScope = $artifactOperationScope.Clone()
    $sharedDependencyScope['coordination'] = 'shared-dependency'
    $sharedDependencyLock = $null
    $effectiveArguments = @($Arguments)
    $hasNodeReuseSetting = $effectiveArguments | Where-Object {
        $_ -match '^(?i)/(nr|noder[e]?use):'
    }
    if (-not $hasNodeReuseSetting) {
        # Disable MSBuild node reuse for script-driven builds. Reused build nodes can keep
        # helper processes alive across invocations and make the outer script report a tool
        # failure even when the actual compile/link work succeeded.
        $effectiveArguments += "/nr:false"
    }

    try {
        $sharedDependencyLock = Enter-RSArtifactOperationLock `
            -RepoRoot $SolutionDir `
            -Operation "MSBuild shared dependency phase $Configuration|$Platform" `
            -Scope $sharedDependencyScope
        if ($sharedDependencyLock.WasAbandoned) {
            [void](Set-RSArtifactOperationContaminated `
                    -RepoRoot $SolutionDir `
                    -Reason 'The previous MSBuild owner exited while it could have been writing shared vcpkg dependencies.' `
                    -AbandonedOwner $sharedDependencyLock.AbandonedOwner `
                    -Scope $sharedDependencyScope)
        }

        $sharedDependencyContamination = Read-RSArtifactOperationContamination `
            -RepoRoot $SolutionDir `
            -Scope $sharedDependencyScope
        if ($null -ne $sharedDependencyContamination) {
            $script:sharedDependencyRepairAuthorized = Test-RSArtifactOperationRepairAllowed `
                -Contamination $sharedDependencyContamination `
                -Rebuild:$Rebuild `
                -ProjectName $repairProjectName `
                -Configuration $Configuration `
                -Platform $Platform
            if (-not $script:sharedDependencyRepairAuthorized) {
                $markerPath = Get-RSArtifactContaminationMarkerPath `
                    -RepoRoot $SolutionDir `
                    -Scope $sharedDependencyScope
                throw "Shared vcpkg dependencies may be incomplete after an interrupted MSBuild operation. Run a full-solution build.ps1 -Rebuild for platform $Platform. Marker: $markerPath"
            }
        }
        Assert-RSNoResidualArtifactToolProcesses `
            -RepoRoot $SolutionDir `
            -Scope $sharedDependencyScope
        Update-RSBuildVcpkgOverlayTriplet

        $logPath = New-ProcessLogPath -Prefix 'msbuild'
        $script:LastProcessLogPath = $logPath
        $invocationPlan = Get-RSMSBuildInvocationPlan -UseInteractiveTerminal $script:UseInteractiveTerminal -LogPath $logPath
        if ($invocationPlan.UseDirectConsole) {
            # Preserve MSBuild's native color and message ordering in interactive terminals while still writing a file log.
            $exitCode = Invoke-RSProcess `
                -FilePath $MSBuildPath `
                -Arguments @($effectiveArguments + @($invocationPlan.AdditionalArguments)) `
                -WorkingDirectory $WorkingDirectory
        }
        else {
            $exitCode = Invoke-RSStreamingProcess `
                -FilePath $MSBuildPath `
                -Arguments $effectiveArguments `
                -WorkingDirectory $WorkingDirectory `
                -LogPath $logPath `
                -OutputLineCallback {
                param(
                    [string]$Line,
                    [bool]$IsError
                )

                    Write-RSMSBuildStreamingLine -Line $Line -IsError $IsError
                }
        }
        Write-Host "Captured log: $logPath" -ForegroundColor DarkGray
        Write-RSMSBuildDiagnosticSummary -LogPath $logPath
        if($exitCode -eq 0){Assert-NativeBuildDiagnostics -LogPath $logPath}
        $global:LASTEXITCODE = $exitCode
        return $exitCode
    }
    finally {
        Exit-RSArtifactOperationLock -Lock $sharedDependencyLock
    }
}

# Find MSBuild
$msbuildInfo = Find-MSBuild

if (-not $msbuildInfo) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "MSBuild Not Found!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install one of the following:" -ForegroundColor Yellow
    Write-Host "  - Visual Studio 2022 (recommended)" -ForegroundColor Yellow
    Write-Host "  - Visual Studio 2019" -ForegroundColor Yellow
    Write-Host "  - Visual Studio Build Tools" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Download from: https://visualstudio.microsoft.com/downloads/" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Ensure 'Desktop development with C++' workload is installed." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

$msbuildPath = $msbuildInfo.Path
$vsVersion = $msbuildInfo.Version

Write-Host "Found: $vsVersion" -ForegroundColor Green
Write-Host "MSBuild: $msbuildPath" -ForegroundColor Green
Write-Host "Detection method: $($msbuildInfo.Method)" -ForegroundColor Gray
Write-Host ""

# Determine build target
$buildTarget = if ($Rebuild) {
    "Rebuild"
} elseif ($Clean) {
    "Clean;Build"
} else {
    "Build"
}

# Display build configuration
Write-Host "Build Configuration:" -ForegroundColor Cyan
Write-Host "  Solution:      $SolutionFile"
Write-Host "  Target:        $(if ($null -ne $unionSelection) { $unionSelection.ProjectNames -join ', ' } elseif ($ProjectName) { $ProjectName } else { 'All Projects' })"
Write-Host "  Configuration: $Configuration"
Write-Host "  Platform:      $Platform"
Write-Host "  Build workers: $(if ($MaxCpuCount -gt 0) { $MaxCpuCount } else { 'MSBuild default' })" -ForegroundColor Gray
Write-Host "  Version:       $($versionContext.DisplayVersion)" -ForegroundColor Gray
Write-Host "  Build number:  $($versionContext.BuildNumber)" -ForegroundColor Gray
Write-Host "  Version state: $versionStatePath" -ForegroundColor Gray
    Write-Host "  Action:        $buildTarget"
    Write-Host "  Package:       $packageMode"
    Write-Host "  Monitor diag:  Runtime flag (--etw)"
    Write-Host ""

$buildInput = $SolutionFile
$resolvedProjectPath = $null
$buildProjectDirectly = $false
$projectCleanTarget = 'Clean'
if ($ProjectName) {
    $buildSelection = Get-RSBuildSelection `
        -SolutionPath $SolutionFullPath `
        -SolutionDir $SolutionDir `
        -ProjectName $ProjectName `
        -Rebuild:$Rebuild

    $resolvedProjectPath = $buildSelection.ResolvedProjectPath
    $buildInput = $buildSelection.BuildInput
    $buildProjectDirectly = $buildSelection.BuildProjectDirectly
    $projectCleanTarget = $buildSelection.CleanTarget
} elseif ($null -ne $unionSelection) {
    $buildInput = $unionSelection.BuildInput
    $projectCleanTarget = $unionSelection.CleanTarget
}

$buildOutputDir = Join-Path -Path $SolutionDir -ChildPath (".build\{0}\{1}" -f $Platform, $Configuration)
$artifactManifestDir = Join-Path -Path $SolutionDir -ChildPath (".build\BuildArtifactManifests\{0}" -f ([Guid]::NewGuid().ToString('N')))
$buildReceiptSchema = Join-Path -Path $SolutionDir -ChildPath 'Specs\Build\BuildReceipt.schema.json'
$toolchainIdentitySchema = Join-Path -Path $SolutionDir -ChildPath 'Specs\Build\ToolchainIdentity.schema.json'
$receiptTarget = if ($null -ne $unionSelection) { $unionSelection.ReceiptTarget } elseif ($ProjectName) { $ProjectName.ToLowerInvariant() } else { 'solution' }
$receiptProjectSelection = if ($null -ne $unionSelection) {
    @($unionSelection.ProjectSelection)
} elseif ($resolvedProjectPath) {
    @([IO.Path]::GetRelativePath($SolutionDir, $resolvedProjectPath).Replace('\', '/'))
} else {
    @([IO.Path]::GetRelativePath($SolutionDir, $SolutionFullPath).Replace('\', '/'))
}
if ($ProductionOnly) {
    $productionProjects = @(Get-RSSolutionProjects -SolutionPath $SolutionFullPath | Where-Object {
        $_.RelativePath -match '\.vcxproj$' -and $_.RelativePath -notmatch '^(Tests|PoC)[\\/]'
    })
    if (-not $productionProjects.Count) { throw 'No production solution projects were selected.' }
    $receiptProjectSelection = @($productionProjects | ForEach-Object { $_.RelativePath.Replace('\','/') })
}
$receiptBuildArguments = @(
    "build-number=$($versionContext.BuildNumber)",
    "official-release=$([bool]$OfficialRelease)",
    "monitor-diagnostics=$([bool]$MonitorDiagnostics)"
)
$receiptBuildArguments += "production-only=$($ProductionOnly.IsPresent.ToString().ToLowerInvariant())"
$requestedCompilerDebugInformation = [Environment]::GetEnvironmentVariable('RSBuildCompilerDebugInformation', 'Process')
if ([string]::IsNullOrWhiteSpace($requestedCompilerDebugInformation)) {
    $effectiveCompilerDebugInformation = 'ProgramDatabase'
} elseif ($requestedCompilerDebugInformation -ieq 'Embedded') {
    $effectiveCompilerDebugInformation = 'Embedded'
} elseif ($requestedCompilerDebugInformation -ieq 'ProgramDatabase') {
    $effectiveCompilerDebugInformation = 'ProgramDatabase'
} else {
    throw "RSBuildCompilerDebugInformation must be 'Embedded' or 'ProgramDatabase' when specified."
}
$receiptBuildArguments += "compiler-debug-information=$($effectiveCompilerDebugInformation.ToLowerInvariant())"
$requestedTestsEnabled = [Environment]::GetEnvironmentVariable('RSBuildEnableTests', 'Process')
if ([string]::IsNullOrWhiteSpace($requestedTestsEnabled)) {
    $effectiveTestsEnabled = ($Configuration -ne 'Release')
} elseif ($requestedTestsEnabled -ieq 'true') {
    $effectiveTestsEnabled = $true
} elseif ($requestedTestsEnabled -ieq 'false') {
    $effectiveTestsEnabled = $false
} else {
    throw "RSBuildEnableTests must be 'true' or 'false' when specified."
}
$receiptTestsEnabled = $effectiveTestsEnabled
$receiptBuildArguments += "tests-enabled=$($effectiveTestsEnabled.ToString().ToLowerInvariant())"
Write-Host 'Computing build evidence identity...' -ForegroundColor Gray
Import-Module (Join-Path $SolutionDir 'Tools/Modules/Build/DxUiDependency.psm1') -Force -ErrorAction Stop
if ($DxUiSdkInputDescriptor) {
    $engine = Join-Path $SolutionDir '.ci-engine'
    $sdkModule = Import-Module (Join-Path $engine 'Tools/Modules/Ci.SdkConsumer.psm1') -PassThru
    $sdkInput = & $sdkModule {param($r,$d,$p,$c,$m) Open-CiApplicationSdkInput -ProjectRoot $r -Descriptor $d -Platform $p -Configuration $c -MSBuildPath $m} $SolutionDir $DxUiSdkInputDescriptor $Platform $Configuration $msbuildPath
    if ($null -ne $sdkInput -and $sdkInput.Evidence.recipe -cne 'PinnedQualifiedReleaseSdkV2') {
        & $sdkModule {param($r) Close-CiPinnedSdkInput $r} $sdkInput
        $sdkInput=$null
        Write-Host 'The optional SDK does not provide the qualified ABI contract; retaining pinned Source.'
    }
    if ($null -ne $sdkInput) {
        foreach ($path in @($sdkInput.SdkRoot,$sdkInput.LeasePath,$engine,$SolutionDir)) {if ($path -match '[%;"\r\n]') {throw 'SDK paths contain unsupported MSBuild delimiters.'}}
        $sdkProperties=@("/p:RSDxUiSdkRoot=$($sdkInput.SdkRoot)","/p:CiSdkRoot=$($sdkInput.SdkRoot)","/p:CiSdkProjectRoot=$SolutionDir","/p:CiSdkLeasePath=$($sdkInput.LeasePath)","/p:RSCiEngineRoot=$engine",'/p:CiSdkContract=QualifiedAbiV2')
        $receiptBuildArguments+=@('dxui-recipe=PinnedQualifiedReleaseSdkV2',"dxui-sdk-input-key=$($sdkInput.Evidence.inputKey)","dxui-sdk-manifest=$($sdkInput.Evidence.manifestSha256)","dxui-sdk-descriptor-sha256=$($sdkInput.Descriptor.Hash.ToLowerInvariant())")
    }
}
Restore-RSDxUiDependency -RepoRoot $SolutionDir -MSBuildPath $msbuildPath -Platform $Platform -Configuration $Configuration -CheckUpdates -Rebuild:$Rebuild -ProjectName $repairProjectName -SdkInput $sdkInput
if($null -eq $sdkInput){
    $oldSdkRecord=Resolve-RSGuardedRepositoryPath -RepoRoot $SolutionDir -GuardedRelativeRoot ".build/$Platform/$Configuration" -Path (Join-Path $buildOutputDir 'DxUi.sdk-input.json')
    if(Test-Path -LiteralPath $oldSdkRecord){Remove-Item -LiteralPath $oldSdkRecord}
}
$buildSourceSnapshotBefore = Get-RSBuildSourceSnapshot -RepoRoot $SolutionDir
$buildIdentity = Get-RSBuildIdentityBundle -RepoRoot $SolutionDir -MSBuildPath $msbuildPath `
    -Platform $Platform -Configuration $Configuration
$toolchainIdentityPath = Join-Path $buildOutputDir 'toolchain-identity.json'
$priorReceipt = $null
try {
    $priorReceipt = Read-RSBuildReceipt -BuildOutputDir $buildOutputDir -SchemaPath $buildReceiptSchema
} catch {
    Write-Warning "Existing build receipt is invalid and will not be trusted: $($_.Exception.Message)"
}
$receiptCompatible = Test-RSBuildReceiptCompatible `
    -Receipt $priorReceipt `
    -SourceSnapshot $buildSourceSnapshotBefore `
    -Identity $buildIdentity `
    -Platform $Platform `
    -Configuration $Configuration `
    -Target $receiptTarget `
    -RepoRoot $SolutionDir `
    -ProjectSelection $receiptProjectSelection `
    -BuildArguments $receiptBuildArguments `
    -TestsEnabled $receiptTestsEnabled `
    -OfficialRelease ([bool]$OfficialRelease)
if (-not $receiptCompatible -and -not $Rebuild) {
    Write-Host 'No compatible build receipt exists; first attestation requires Rebuild.' -ForegroundColor Yellow
    $Rebuild = $true
    if ($ProjectName) {
        $buildSelection = Get-RSBuildSelection `
            -SolutionPath $SolutionFullPath `
            -SolutionDir $SolutionDir `
            -ProjectName $ProjectName `
            -Rebuild
        $resolvedProjectPath = $buildSelection.ResolvedProjectPath
        $buildInput = $buildSelection.BuildInput
        $buildProjectDirectly = $buildSelection.BuildProjectDirectly
        $projectCleanTarget = $buildSelection.CleanTarget
    } elseif ($null -ne $unionSelection) {
        $unionSelection = Get-RSBuildProjectUnion -SolutionPath $SolutionFullPath -ProjectNames $ProjectNames -Rebuild
    }
}

$msbuildTarget = if ($null -ne $unionSelection) {
    $unionSelection.MSBuildTarget
} elseif ($ProjectName) {
    $buildSelection.MSBuildTarget
} elseif ($Rebuild) {
    "Rebuild"
} else {
    "Build"
}
if ($ProductionOnly) {
    $msbuildTarget = (@($productionProjects | ForEach-Object {
        if ($Rebuild) { $_.SolutionTarget + ':Rebuild' } else { $_.SolutionTarget }
    }) -join ';')
}

# Build parameters
$buildParams = @(
    $buildInput
    "/t:$msbuildTarget"
    "/p:Configuration=$Configuration"
    "/p:Platform=$Platform"
    $(if ($MaxCpuCount -gt 0) { "/m:$MaxCpuCount" } else { "/m" })
    "/v:minimal"      # Minimal verbosity
    "/nologo"         # Suppress MSBuild banner
)

$buildParams += "/p:SolutionDir=$SolutionDirWithSlash"
$buildParams += "/p:RSVersionBuildNumber=$($versionContext.BuildNumber)"
$buildParams += "/p:RSBuildEnableTests=$($effectiveTestsEnabled.ToString().ToLowerInvariant())"
$buildParams += "/p:RSBuildCompilerDebugInformation=$effectiveCompilerDebugInformation"
$buildParams += "/p:RSBuildArtifactManifestRoot=$($artifactManifestDir.TrimEnd('\'))\"
$buildParams += $sdkProperties
if ($null -ne $sdkInput) {
    & $msbuildPath (Join-Path $SolutionDir 'RedSalamander/RedSalamander.vcxproj') /nologo /v:minimal '/t:CiValidateSdkConsumer' "/p:Platform=$Platform" "/p:Configuration=$Configuration" "/p:SolutionDir=$SolutionDirWithSlash" @sdkProperties
    if ($LASTEXITCODE) {throw 'The actual application profile refused the qualified SDK before compilation.'}
}
if ($OfficialRelease) {
    $buildParams += "/p:RSVersionOfficialRelease=true"
}
if ($MonitorDiagnostics) {
    Write-Warning "-MonitorDiagnostics is deprecated; launch RedSalamander with --etw instead; RedMonitor is independently installed."
}

$cleanParams = $null
if ($Clean) {
    $cleanParams = @(
        $buildInput
        "/t:Clean"
        "/p:Configuration=$Configuration"
        "/p:Platform=$Platform"
        "/v:minimal"
        "/nologo"
    )
    if ($ProjectName -or $null -ne $unionSelection) {
        if ($buildProjectDirectly) {
            $cleanParams[0] = $buildInput
        } else {
            $cleanParams[1] = "/t:$projectCleanTarget"
        }
    }
    $cleanParams += "/p:SolutionDir=$SolutionDirWithSlash"
    $cleanParams += "/p:RSVersionBuildNumber=$($versionContext.BuildNumber)"
    $cleanParams += "/p:RSBuildEnableTests=$($effectiveTestsEnabled.ToString().ToLowerInvariant())"
    if ($OfficialRelease) {
        $cleanParams += "/p:RSVersionOfficialRelease=true"
    }
    if ($MonitorDiagnostics) {
        Write-Warning "-MonitorDiagnostics is deprecated; launch RedSalamander with --etw instead; RedMonitor is independently installed."
    }
}

# Start build
Write-Host "Starting build..." -ForegroundColor Yellow
$stopwatch.Restart()

function Assert-BuildOutputProcessNotRunning {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessName,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedExePath
    )

    $expectedFullPath = $null
    try {
        $expectedFullPath = [System.IO.Path]::GetFullPath($ExpectedExePath)
    }
    catch {
        throw "Build canceled because target-output process safety could not be verified: invalid expected executable path '$ExpectedExePath'. No process was terminated. $($_.Exception.Message)"
    }

    $escapedName = $ProcessName.Replace("'", "''")
    $processes = @()
    try {
        $processes = Get-CimInstance Win32_Process -Filter "Name='$escapedName'" -ErrorAction Stop
    }
    catch {
        throw ("Build canceled because target-output process safety could not be verified for '$expectedFullPath': " +
               "unable to enumerate '$ProcessName' process metadata. No process was terminated. $($_.Exception.Message)")
    }

    $matchingProcesses = @()
    foreach ($proc in $processes) {
        $exePath = $proc.ExecutablePath
        if (-not $exePath) {
            continue
        }

        $exeFullPath = $null
        try {
            $exeFullPath = [System.IO.Path]::GetFullPath($exePath)
        }
        catch {
            continue
        }

        if (-not [string]::Equals($exeFullPath, $expectedFullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $matchingProcesses += [pscustomobject]@{
            ProcessId = $proc.ProcessId
            ExecutablePath = $exeFullPath
            CommandLine = $proc.CommandLine
        }
    }

    if ($matchingProcesses.Count -gt 0) {
        $diagnostics = @($matchingProcesses | ForEach-Object {
            $commandLine = if ([string]::IsNullOrWhiteSpace($_.CommandLine)) { '<unavailable>' } else { $_.CommandLine }
            "  PID=$($_.ProcessId); Path='$($_.ExecutablePath)'; CommandLine='$commandLine'"
        })
        throw ("Build canceled because an independently launched process is using an exact target output. " +
               "It was not terminated because it is not proven to belong to this operation; close it and retry.`n" +
               ($diagnostics -join "`n"))
    }
}

    Assert-BuildOutputProcessNotRunning -ProcessName "RedSalamander.exe" -ExpectedExePath (Join-Path -Path $buildOutputDir -ChildPath "RedSalamander.exe")

    $buildStartedUtc = [datetime]::UtcNow
    Suspend-RSBuildReceipt -BuildOutputDir $buildOutputDir
    $toolchainIdentityPath = Publish-RSBuildToolchainIdentity `
        -BuildOutputDir $buildOutputDir `
        -IdentityRecord $buildIdentity.ToolchainIdentityRecord `
        -SchemaPath $toolchainIdentitySchema

    # Execute clean if requested
    if ($Clean) {
        Write-Host "Cleaning..." -ForegroundColor Yellow
        $null = Invoke-MSBuild -MSBuildPath $msbuildPath -Arguments $cleanParams -WorkingDirectory $SolutionDir

        if ($LASTEXITCODE -ne 0) {
            Write-Host "Clean failed, continuing with build..." -ForegroundColor Yellow
        }
        Write-Host "Building..." -ForegroundColor Yellow
    }

    $null = Invoke-MSBuild -MSBuildPath $msbuildPath -Arguments $buildParams -WorkingDirectory $SolutionDir

    if ($LASTEXITCODE -ne 0) {
        $stopwatch.Stop()
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "Build Failed!" -ForegroundColor Red
        Write-Host "Exit code: $LASTEXITCODE" -ForegroundColor Red
        Write-Host "Build time: $($stopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        exit $LASTEXITCODE
    }

    $stopwatch.Stop()
    $buildSourceSnapshotAfter = Get-RSBuildSourceSnapshot -RepoRoot $SolutionDir
    if ($buildSourceSnapshotAfter.SnapshotId -ne $buildSourceSnapshotBefore.SnapshotId) {
        throw 'Repository source changed while the build was running; no success receipt was published.'
    }
    $packageArtifacts = @()

    $packagingLock = $null
    $packageEvidenceBefore = @{}
    $appPackagesEvidenceRoot = Join-Path -Path $SolutionDir -ChildPath '.build\AppPackages'
    if ($Zip -or $GenerateWingetManifest) {
        $appPackagesEvidenceRoot = Resolve-RSGuardedRepositoryPath `
            -RepoRoot $SolutionDir -GuardedRelativeRoot '.build\AppPackages' `
            -Path $appPackagesEvidenceRoot
        if (Test-Path -LiteralPath $appPackagesEvidenceRoot -PathType Container) {
            foreach ($file in @(Get-ChildItem -LiteralPath $appPackagesEvidenceRoot -File -Recurse)) {
                $packageEvidenceBefore[$file.FullName] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            }
        }
    }
    if ($Zip -or $GenerateWingetManifest) {
        $packagingScope = $artifactOperationScope.Clone()
        $packagingScope['coordination'] = 'packaging'
        $packagingLock = Enter-RSArtifactOperationLock `
            -RepoRoot $SolutionDir `
            -Operation "packaging $packageMode winget=$GenerateWingetManifest" `
            -Scope $packagingScope
        if ($packagingLock.WasAbandoned) {
            [void](Set-RSArtifactOperationContaminated `
                    -RepoRoot $SolutionDir `
                    -Reason 'The previous packaging owner exited while mutating shared installer inputs or AppPackages outputs.' `
                    -AbandonedOwner $packagingLock.AbandonedOwner `
                    -Scope $packagingScope)
        }
        if (Test-RSArtifactOperationContaminated -RepoRoot $SolutionDir -Scope $packagingScope) {
            $markerPath = Get-RSArtifactContaminationMarkerPath `
                -RepoRoot $SolutionDir `
                -Scope $packagingScope
            throw "Shared packaging state may be incomplete after an interrupted operation. Review and remove the marker only after restoring Installer inputs and .build\AppPackages. Marker: $markerPath"
        }
    }

    try {
        # Capture native bytes and module provenance before staging the portable package.
        Write-RSDxUiModuleProvenance -RepoRoot $SolutionDir -ArtifactManifestDir $artifactManifestDir -Platform $Platform -Configuration $Configuration -SdkInput $sdkInput
        $sdkRecordPath=if($null -ne $sdkInput){Write-RSDxUiSdkBuildRecord -RepoRoot $SolutionDir -BuildOutputDir $buildOutputDir -SdkInput $sdkInput -MSBuildPath $msbuildPath -Platform $Platform}else{$null}
        $buildArtifacts = @(Get-RSBuildArtifactRecords `
                -RepoRoot $SolutionDir `
                -BuildOutputDir $buildOutputDir `
                -ArtifactManifestDir $artifactManifestDir `
                -Configuration $Configuration `
                -Platform $Platform `
                -ProjectName $(if ($ProjectName) { $ProjectName } else { '' }) `
                -ResolvedProjectPath $(if ($resolvedProjectPath) { $resolvedProjectPath } else { '' }) `
                -ProjectNames $(if ($null -ne $unionSelection) { @($unionSelection.ProjectNames) } else { @() }))
        $buildArtifacts += New-RSBuildArtifactRecord `
            -RepoRoot $SolutionDir `
            -LiteralPath $toolchainIdentityPath
        $buildArtifacts += @(if($sdkRecordPath){New-RSBuildArtifactRecord -RepoRoot $SolutionDir -LiteralPath $sdkRecordPath})
        $buildArtifacts = @($buildArtifacts | Sort-Object relative_path)
        $receiptTestsEnabled = Test-RSBuildArtifactManifestTestsEnabled -ArtifactManifestDir $artifactManifestDir -RequireUniform
        if (($Zip -or $GenerateWingetManifest -or $OfficialRelease) -and $receiptTestsEnabled) {
            throw 'A distributable build contains a test-enabled output; no package receipt will be published.'
        }
        $summaryArtifacts = if ($null -ne $unionSelection) {
            @($buildArtifacts | Where-Object role -eq 'launchable')
        } elseif ($ProjectName) {
            @($buildArtifacts | Where-Object {
                    [IO.Path]::GetFileNameWithoutExtension($_.relative_path) -eq $ProjectName
                })
        } else {
            @($buildArtifacts | Where-Object role -eq 'launchable')
        }

        if ($Zip) {
            $zipScript = Join-Path -Path $SolutionDir -ChildPath "Installer\\zip\\build-zip.ps1"
            if (-not (Test-Path $zipScript)) {
                Write-Error "ZIP build script not found: $zipScript"
                exit 1
            }

            Write-Host ""
            Write-Host "Building ZIP package..." -ForegroundColor Yellow
            $zipStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

            try {
                & $zipScript -Configuration $Configuration -Platform $Platform -BuildNumber $versionContext.BuildNumber
            }
            catch {
                $zipStopwatch.Stop()
                Write-Host ""
                Write-Host "========================================" -ForegroundColor Red
                Write-Host "ZIP Packaging Failed!" -ForegroundColor Red
                Write-Host "Error: $_" -ForegroundColor Red
                Write-Host "Packaging time: $($zipStopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Red
                Write-Host "========================================" -ForegroundColor Red
                exit 1
            }

            $zipStopwatch.Stop()
            Write-Host "ZIP packaging completed successfully! ($($zipStopwatch.Elapsed.ToString('mm\:ss')))" -ForegroundColor Green
        }

        if ($GenerateWingetManifest) {
            $wingetScript = Join-Path -Path $SolutionDir -ChildPath "Installer\\winget\\generate-manifest.ps1"
            if (-not (Test-Path $wingetScript)) {
                Write-Error "Winget manifest generation script not found: $wingetScript"
                exit 1
            }

            Write-Host ""
            Write-Host "Generating winget manifest..." -ForegroundColor Yellow

            $appPackagesDir = Join-Path -Path $SolutionDir -ChildPath ".build\\AppPackages"
            $packageVersion = $versionContext.PackagingVersion
            $ZipPath = Join-Path $appPackagesDir "RedSalamander-$packageVersion-x64-Portable.zip"
            $Arm64ZipPath = Join-Path $appPackagesDir "RedSalamander-$packageVersion-ARM64-Portable.zip"

            try {
                $wingetParams = @{
                    BuildNumber = $versionContext.BuildNumber
                    ZipPath = $ZipPath
                    Arm64ZipPath = $Arm64ZipPath
                }

                & $wingetScript @wingetParams
            }
            catch {
                Write-Host ""
                Write-Host "========================================" -ForegroundColor Red
                Write-Host "Winget Manifest Generation Failed!" -ForegroundColor Red
                Write-Host "Error: $_" -ForegroundColor Red
                Write-Host "========================================" -ForegroundColor Red
                exit 1
            }
        }
    }
    finally {
        Exit-RSArtifactOperationLock -Lock $packagingLock
    }

    if ($contaminationRepairAuthorized) {
        Clear-RSArtifactOperationContaminated `
            -RepoRoot $SolutionDir `
            -Scope $artifactOperationScope
    }
    if ($script:sharedDependencyRepairAuthorized) {
        $sharedDependencyScope = $artifactOperationScope.Clone()
        $sharedDependencyScope['coordination'] = 'shared-dependency'
        Clear-RSArtifactOperationContaminated `
            -RepoRoot $SolutionDir `
            -Scope $sharedDependencyScope
    }

    if ($Zip -or $GenerateWingetManifest) {
        if (-not (Test-Path -LiteralPath $appPackagesEvidenceRoot -PathType Container)) {
            throw "Requested packaging produced no AppPackages directory: $appPackagesEvidenceRoot"
        }
        # The packaging lease is released here: check only, never delete. The in-lease check
        # already checked the newly produced portable ZIP.
        foreach ($file in @(Get-ChildItem -LiteralPath $appPackagesEvidenceRoot -File -Recurse | Sort-Object FullName)) {
            $currentDigest = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            if (-not $packageEvidenceBefore.ContainsKey($file.FullName) -or $packageEvidenceBefore[$file.FullName] -ne $currentDigest) {
                if ($file.Extension -ieq '.zip') {
                    Assert-RSShippingPackageHasNoTestSurface -LiteralPath $file.FullName
                }
                $packageArtifact = New-RSBuildArtifactRecord -RepoRoot $SolutionDir -LiteralPath $file.FullName
                $packageArtifacts += $packageArtifact
                $summaryArtifacts += $packageArtifact
            }
        }
    }

    $operationDuration = [datetime]::UtcNow - $buildStartedUtc
    $buildSourceSnapshotFinal = Get-RSBuildSourceSnapshot -RepoRoot $SolutionDir
    if ($buildSourceSnapshotFinal.SnapshotId -ne $buildSourceSnapshotBefore.SnapshotId) {
        throw 'Repository source changed while the selected build/package operation was running; no success receipt was published.'
    }
    # Staging-only packagers must not change the native closure or its provenance after
    # capture. Reuse the canonical record reader and fail before publishing a stale receipt.
    foreach ($capturedArtifact in $buildArtifacts) {
        $currentArtifact = New-RSBuildArtifactRecord -RepoRoot $SolutionDir `
            -LiteralPath (Join-Path $SolutionDir $capturedArtifact.relative_path)
        if ($currentArtifact.sha256 -cne $capturedArtifact.sha256 -or
            $currentArtifact.size_bytes -ne $capturedArtifact.size_bytes) {
            throw "Build output changed after native artifact capture: $($capturedArtifact.relative_path)"
        }
    }
    if ((Test-RSBuildArtifactManifestTestsEnabled -ArtifactManifestDir $artifactManifestDir -RequireUniform) -ne $receiptTestsEnabled) {
        throw 'Build artifact test surface changed during packaging; no success receipt will be published.'
    }
    $receiptArtifacts = @(@($buildArtifacts) + @($packageArtifacts) | Sort-Object relative_path)
    if ($null -ne $sdkInput) {& $sdkModule {param($r) Close-CiPinnedSdkInput $r} $sdkInput}
    $buildReceipt = New-RSBuildReceipt `
        -SourceSnapshot $buildSourceSnapshotFinal `
        -Identity $buildIdentity `
        -Platform $Platform `
        -Configuration $Configuration `
        -Target $receiptTarget `
        -ProjectSelection $receiptProjectSelection `
        -BuildArguments $receiptBuildArguments `
        -TestsEnabled $receiptTestsEnabled `
        -OfficialRelease ([bool]$OfficialRelease) `
        -StartedUtc $buildStartedUtc `
        -EndedUtc ([datetime]::UtcNow) `
        -Duration $operationDuration `
        -Outputs $receiptArtifacts
    $buildReceipt = Publish-RSBuildReceipt `
        -BuildOutputDir $buildOutputDir `
        -Receipt $buildReceipt `
        -SchemaPath $buildReceiptSchema
    Write-Host "Build artifacts attested: $($buildReceipt.receipt_id)" -ForegroundColor Gray

    $visibleArtifacts=@($summaryArtifacts|Sort-Object relative_path|ForEach-Object{[string]$_.absolute_path_diagnostic})
    $app=@($summaryArtifacts|Where-Object{[IO.Path]::GetFileName([string]$_.absolute_path_diagnostic)-eq 'RedSalamander.exe'}|Select-Object -First 1)
    $launch=if($app.Count){[string]$app[0].absolute_path_diagnostic}else{''}
    $processLog=Get-Variable -Name LastProcessLogPath -Scope Script -ErrorAction SilentlyContinue
    if($processLog -and (Test-Path -LiteralPath $processLog.Value)){
        Write-NativeBuildResult -Project RedSalamander -Platform $Platform -Configuration $Configuration -LogPath $processLog.Value -Artifacts $visibleArtifacts -Executable $launch -Duration $nativeEntryElapsed.Elapsed
    }else{
        foreach($line in @(ConvertTo-RSBuildSummaryLines -Configuration $Configuration -Platform $Platform -Duration $operationDuration -Artifacts @($summaryArtifacts))){Write-Host $line -ForegroundColor Green}
    }
    exit 0
}
catch {
    $sdkBuildFailure=$_
    $stopwatch.Stop()
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Build Failed with Exception!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host "Build time: $($stopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Red
    exit 1
}
finally {
    try {
        if ($null -ne $sdkInput -and -not $sdkInput.Closed) {& $sdkModule {param($r) Close-CiPinnedSdkInput $r} $sdkInput}
    } catch {if ($null -eq $sdkBuildFailure) {throw} else {Write-Warning ('SDK close also failed: '+$_.Exception.Message)}}
    finally {Exit-RSArtifactOperationLock -Lock $artifactOperationLock}
}
