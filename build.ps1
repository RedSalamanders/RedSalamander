<#
.SYNOPSIS
    Build RedSalamander solution
.DESCRIPTION
    Builds the entire RedSalamander solution in Debug (default), Release, or ASan Debug configuration.
    Optionally builds a specific project if ProjectName is provided.
.PARAMETER Configuration
    Build configuration: Debug, Release, or ASan Debug (default: Debug)
.PARAMETER Platform
    Target platform: x64 or ARM64 (default: x64)
.PARAMETER ProjectName
    Specific project to build. If not specified, builds the entire solution.
    Examples: "RedSalamander", "RedSalamanderMonitor", "Common"
.PARAMETER Clean
    Perform a clean build
.PARAMETER Rebuild
    Rebuild all projects
.PARAMETER MaxCpuCount
    Optional maximum MSBuild worker count. Zero uses MSBuild's default.
.PARAMETER Msix
    Build MSIX package after a successful build (Release only)
.PARAMETER Msi
    Build MSI package after a successful build (Release only, requires WiX Toolset)
.PARAMETER BuildNumber
    Override the build number shared by all modules in the current build invocation.
    Required for local official release builds. In CI, GITHUB_RUN_NUMBER is used automatically.
.PARAMETER OfficialRelease
    Stamp binaries as an official release build (no local beta Debug/Release suffix in displayed versions).
    Official release mode requires -BuildNumber or a CI environment with GITHUB_RUN_NUMBER.
.PARAMETER MonitorDiagnostics
    Deprecated. Diagnostics are enabled at runtime with --etw and perf JSONL with --perf.
.EXAMPLE
    .\build.ps1
    Builds entire solution in Debug configuration
.EXAMPLE
    .\build.ps1 -Configuration Release
    Builds entire solution in Release configuration
.EXAMPLE
    .\build.ps1 -ProjectName RedSalamanderMonitor
    Builds only RedSalamanderMonitor project
.EXAMPLE
    .\build.ps1 -Configuration Release -Clean
    Clean build of entire solution in Release configuration
.EXAMPLE
    .\build.ps1 -Msix
    Builds entire solution in Release and produces an MSIX package
.EXAMPLE
    .\build.ps1 -Msi
    Builds entire solution in Release and produces an MSI installer
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
    
    [Parameter(HelpMessage = "Perform a clean build")]
    [switch]$Clean,
    
    [Parameter(HelpMessage = "Rebuild all projects")]
    [switch]$Rebuild,

    [Parameter(HelpMessage = "Maximum MSBuild worker count (0 uses the default)")]
    [ValidateRange(0, 256)]
    [int]$MaxCpuCount = 0,

    [Parameter(HelpMessage = "Build MSIX package after build (Release only)")]
    [switch]$Msix,

    [Parameter(HelpMessage = "Build MSI package after build (Release only, requires WiX Toolset)")]
    [switch]$Msi,

    [Parameter(HelpMessage = "Build ZIP package after build (Release only)")]
    [switch]$Zip,

    [Parameter(HelpMessage = "Generate winget manifest after build (Release only)")]
    [switch]$GenerateWingetManifest,

    [Parameter(HelpMessage = "Override the build number (required for local official release builds)")]
    [int]$BuildNumber = 0,

    [Parameter(HelpMessage = "Stamp binaries as an official release build (requires -BuildNumber or CI)")]
    [switch]$OfficialRelease,

    [Parameter(HelpMessage = "Deprecated; use the runtime --etw flag instead")]
    [switch]$MonitorDiagnostics
)

$ErrorActionPreference = "Stop"
$BuildProjectSelectionModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\BuildProjectSelection.psm1"
$BuildEvidenceModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\BuildEvidence.psm1"
$MSBuildInvocationModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\MSBuildInvocation.psm1"
$ProcessStreamingModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\ProcessStreaming.psm1"
$SanitizedEnvironmentModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\SanitizedEnvironment.psm1"
$ArtifactOperationLockModule = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Modules\Build\ArtifactOperationLock.psm1"
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

Import-Module $BuildProjectSelectionModule -Force -ErrorAction Stop
Import-Module $BuildEvidenceModule -Force -ErrorAction Stop
Import-Module $MSBuildInvocationModule -Force -ErrorAction Stop
Import-Module $SanitizedEnvironmentModule -Force -ErrorAction Stop
Import-Module $ProcessStreamingModule -Force -ErrorAction Stop
Import-Module $ArtifactOperationLockModule -Force -ErrorAction Stop

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
if ($Msix -and $Msi) {
    Write-Error "Specify only one of -Msix or -Msi."
    exit 1
}

if (($Msix -or $Msi -or $Zip) -and $ProjectName) {
    Write-Error "Packaging requires building the full solution. Remove -ProjectName."
    exit 1
}

$packageMode = if ($Msix) { "MSIX" } elseif ($Msi) { "MSI" } elseif ($Zip) { "ZIP" } else { "None" }

if ($Msix -or $Msi -or $Zip) {
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
$artifactOperationLock = $null
$stopwatch = [System.Diagnostics.Stopwatch]::new()
$contaminationRepairAuthorized = $false
$script:sharedDependencyRepairAuthorized = $false
try {
    $operationTarget = if ($ProjectName) { $ProjectName } else { 'solution' }
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
            -ProjectName $ProjectName `
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
            if ($selectedVersion.FileMajorPart -lt 18 -or
                $selectedPath -notmatch '\\Microsoft Visual Studio\\') {
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
            $candidatePath = $msbuildInPath.Source
            $fileMajor = $null
            try {
                $ver = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($candidatePath)
                $fileMajor = $ver.FileMajorPart
            }
            catch {
                $fileMajor = $null
            }

            # VS 2026 MSBuild should report major version 18.
            if (($null -ne $fileMajor) -and ($fileMajor -ge 18) -and ($candidatePath -match '\\Microsoft Visual Studio\\')) {
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

            $msbuildCandidates = @(
                (Join-Path $installPath "MSBuild\\Current\\Bin\\amd64\\MSBuild.exe"),
                (Join-Path $installPath "MSBuild\\Current\\Bin\\MSBuild.exe"),
                (Join-Path $installPath "MSBuild\\15.0\\Bin\\amd64\\MSBuild.exe"),
                (Join-Path $installPath "MSBuild\\15.0\\Bin\\MSBuild.exe")
            )

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
                $msbuildPaths = @(
                    "$basePath\$year\$edition\MSBuild\Current\Bin\MSBuild.exe",
                    "$basePath\$year\$edition\MSBuild\Current\Bin\amd64\MSBuild.exe",
                    "$basePath\$year\$edition\MSBuild\15.0\Bin\MSBuild.exe",
                    "$basePath\$year\$edition\MSBuild\15.0\Bin\amd64\MSBuild.exe"
                )
                
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
            Path = $msbuildInPath.Source
            Version = "Found in PATH"
            Method = "PATH"
        }
    }
    
    # Strategy 4: Use Developer Command Prompt environment
    Write-Host "  Checking Developer Command Prompt environment..." -ForegroundColor Gray
    
    if ($env:VSINSTALLDIR) {
        $devMSBuildPaths = @(
            "$env:VSINSTALLDIR\MSBuild\Current\Bin\MSBuild.exe",
            "$env:VSINSTALLDIR\MSBuild\15.0\Bin\MSBuild.exe"
        )
        
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

function Get-WappTargetPlatformVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectPath
    )

    $text = Get-Content -Path $ProjectPath -Raw
    $match = [Regex]::Match($text, '<TargetPlatformVersion>\s*([^<]+)\s*</TargetPlatformVersion>')
    if (-not $match.Success) {
        return $null
    }

    return $match.Groups[1].Value.Trim()
}

function Test-UapPropsAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPlatformVersion
    )

    $uapPropsPath = Join-Path ${env:ProgramFiles(x86)} ("Windows Kits\\10\\DesignTime\\CommonConfiguration\\Neutral\\UAP\\{0}\\UAP.props" -f $TargetPlatformVersion)
    return (Test-Path $uapPropsPath)
}

function Get-InstalledUapTargetPlatformVersions {
    $uapRoot = Join-Path ${env:ProgramFiles(x86)} "Windows Kits\\10\\DesignTime\\CommonConfiguration\\Neutral\\UAP"
    if (-not (Test-Path $uapRoot)) {
        return @()
    }

    return Get-ChildItem -Path $uapRoot -Directory -ErrorAction SilentlyContinue |
        ForEach-Object { $_.Name } |
        Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }
}

function Resolve-UapTargetPlatformVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WapProjPath
    )

    $requested = Get-WappTargetPlatformVersion -ProjectPath $WapProjPath
    if ($requested -and (Test-UapPropsAvailable -TargetPlatformVersion $requested)) {
        return $requested
    }

    $installed = Get-InstalledUapTargetPlatformVersions
    if (-not $installed -or $installed.Count -eq 0) {
        throw "Windows SDK folder containing 'UAP.props' was not found. Install the Windows 11 24H2 SDK (10.0.26100.0 UAP) or set TargetPlatformVersion to an installed version."
    }

    $selected = $installed | Sort-Object { [version]$_ } -Descending | Select-Object -First 1
    if ($requested) {
        Write-Host "Windows SDK for UAP $requested not found; using installed UAP $selected." -ForegroundColor Yellow
    }

    return $selected
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
                -ProjectName $ProjectName `
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
Write-Host "  Target:        $(if ($ProjectName) { $ProjectName } else { 'All Projects' })"
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
}

$buildOutputDir = Join-Path -Path $SolutionDir -ChildPath (".build\{0}\{1}" -f $Platform, $Configuration)
$artifactManifestDir = Join-Path -Path $SolutionDir -ChildPath (".build\BuildArtifactManifests\{0}" -f ([Guid]::NewGuid().ToString('N')))
$buildReceiptSchema = Join-Path -Path $SolutionDir -ChildPath 'Specs\Build\BuildReceipt.schema.json'
$toolchainIdentitySchema = Join-Path -Path $SolutionDir -ChildPath 'Specs\Build\ToolchainIdentity.schema.json'
$receiptTarget = if ($ProjectName) { $ProjectName.ToLowerInvariant() } else { 'solution' }
$receiptProjectSelection = if ($resolvedProjectPath) {
    @([IO.Path]::GetRelativePath($SolutionDir, $resolvedProjectPath).Replace('\', '/'))
} else {
    @([IO.Path]::GetRelativePath($SolutionDir, $SolutionFullPath).Replace('\', '/'))
}
$receiptBuildArguments = @(
    "build-number=$($versionContext.BuildNumber)",
    "official-release=$([bool]$OfficialRelease)",
    "monitor-diagnostics=$([bool]$MonitorDiagnostics)"
)
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
    }
}

$msbuildTarget = if ($ProjectName) {
    $buildSelection.MSBuildTarget
} elseif ($Rebuild) {
    "Rebuild"
} else {
    "Build"
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
if ($OfficialRelease) {
    $buildParams += "/p:RSVersionOfficialRelease=true"
}
if ($MonitorDiagnostics) {
    Write-Warning "-MonitorDiagnostics is deprecated; launch RedSalamander or RedSalamanderMonitor with --etw instead."
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
    if ($ProjectName) {
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
        Write-Warning "-MonitorDiagnostics is deprecated; launch RedSalamander or RedSalamanderMonitor with --etw instead."
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
    Assert-BuildOutputProcessNotRunning -ProcessName "RedSalamanderMonitor.exe" -ExpectedExePath (Join-Path -Path $buildOutputDir -ChildPath "RedSalamanderMonitor.exe")

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
    $buildArtifacts = @(Get-RSBuildArtifactRecords `
            -RepoRoot $SolutionDir `
            -BuildOutputDir $buildOutputDir `
            -ArtifactManifestDir $artifactManifestDir `
            -Configuration $Configuration `
            -Platform $Platform `
            -ProjectName $(if ($ProjectName) { $ProjectName } else { '' }) `
            -ResolvedProjectPath $(if ($resolvedProjectPath) { $resolvedProjectPath } else { '' }))
    $buildArtifacts += New-RSBuildArtifactRecord `
        -RepoRoot $SolutionDir `
        -LiteralPath $toolchainIdentityPath
    $buildArtifacts = @($buildArtifacts | Sort-Object relative_path)
    $receiptTestsEnabled = Test-RSBuildArtifactManifestTestsEnabled -ArtifactManifestDir $artifactManifestDir
    $summaryArtifacts = if ($ProjectName) {
        @($buildArtifacts | Where-Object {
                [IO.Path]::GetFileNameWithoutExtension($_.relative_path) -eq $ProjectName
            })
    } else {
        @($buildArtifacts | Where-Object role -eq 'launchable')
    }
    $packageArtifacts = @()

    $packagingLock = $null
    $stagedMsixManifest = $null
    $packageEvidenceBefore = @{}
    $appPackagesEvidenceRoot = Join-Path -Path $SolutionDir -ChildPath '.build\AppPackages'
    if ($Msix -or $Msi -or $Zip -or $GenerateWingetManifest) {
        if (Test-Path -LiteralPath $appPackagesEvidenceRoot -PathType Container) {
            foreach ($file in @(Get-ChildItem -LiteralPath $appPackagesEvidenceRoot -File -Recurse)) {
                $packageEvidenceBefore[$file.FullName] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            }
        }
    }
    if ($Msix -or $Msi -or $Zip -or $GenerateWingetManifest) {
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
    if ($Msix) {
        $msixAssetsScript = Join-Path -Path $SolutionDir -ChildPath "Installer\msix\GenerateAssets.ps1"
        if (-not (Test-Path $msixAssetsScript)) {
            Write-Error "MSIX assets generation script not found: $msixAssetsScript"
            exit 1
        }

        Write-Host ""
        Write-Host "Generating MSIX assets..." -ForegroundColor Yellow
        try {
            & $msixAssetsScript -Configuration $Configuration -Platform $Platform
        }
        catch {
            Write-Error "MSIX assets generation failed: $_"
            exit 1
        }

        $installerProject = Join-Path -Path $SolutionDir -ChildPath "Installer\msix\RedSalamanderInstaller.wapproj"
        if (-not (Test-Path $installerProject)) {
            Write-Error "MSIX packaging project not found: $installerProject"
            exit 1
        }

        $msixVersionScript = Join-Path -Path $SolutionDir -ChildPath "Installer\msix\UpdateManifestVersion.ps1"
        if (-not (Test-Path $msixVersionScript)) {
            Write-Error "MSIX version update script not found: $msixVersionScript"
            exit 1
        }

        try {
            $stagedMsixManifest = & $msixVersionScript `
                -Version $versionContext.PackagingVersion `
                -Platform $Platform `
                -Configuration $Configuration
        }
        catch {
            Write-Error "MSIX manifest version update failed: $_"
            exit 1
        }

        $targetPlatformVersion = $null
        try {
            $targetPlatformVersion = Resolve-UapTargetPlatformVersion -WapProjPath $installerProject
        }
        catch {
            Write-Error "Failed to resolve Windows SDK TargetPlatformVersion for MSIX packaging: $_"
            exit 1
        }

        Write-Host ""
        Write-Host "Building MSIX package..." -ForegroundColor Yellow
        $msixStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        $msixParams = @(
            $installerProject
            "/t:Build"
            "/p:Configuration=$Configuration"
            "/p:Platform=$Platform"
            "/p:TargetPlatformVersion=$targetPlatformVersion"
            "/p:SolutionDir=$SolutionDirWithSlash"
            "/p:RSAppxManifestPath=$stagedMsixManifest"
            "/p:AppxPackageSigningEnabled=false"
            "/p:GenerateAppInstallerFile=false"
            "/p:AppxBundle=Never"
            "/v:minimal"
            "/nologo"
        )

        $null = Invoke-MSBuild -MSBuildPath $msbuildPath -Arguments $msixParams -WorkingDirectory $SolutionDir

        if ($LASTEXITCODE -ne 0) {
            $msixStopwatch.Stop()
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Red
            Write-Host "MSIX Packaging Failed!" -ForegroundColor Red
            Write-Host "Exit code: $LASTEXITCODE" -ForegroundColor Red
            Write-Host "Packaging time: $($msixStopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Red
            Write-Host "========================================" -ForegroundColor Red
            exit $LASTEXITCODE
        }

        $msixStopwatch.Stop()
        Write-Host "MSIX packaging completed successfully! ($($msixStopwatch.Elapsed.ToString('mm\:ss')))" -ForegroundColor Green
    }

    if ($Msi) {
        $msiScript = Join-Path -Path $SolutionDir -ChildPath "Installer\msi\build-msi.ps1"
        if (-not (Test-Path $msiScript)) {
            Write-Error "MSI build script not found: $msiScript"
            exit 1
        }

        $msiSymbolsScript = Join-Path -Path $SolutionDir -ChildPath "Installer\msi\build-msi-symbols.ps1"
        if (-not (Test-Path $msiSymbolsScript)) {
            Write-Error "MSI symbols build script not found: $msiSymbolsScript"
            exit 1
        }

        Write-Host ""
        Write-Host "Building MSI installer..." -ForegroundColor Yellow
        $msiStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            $msiArgs = @(
                "-Configuration", $Configuration,
                "-Platform", $Platform,
                "-BuildNumber", "$($versionContext.BuildNumber)"
            )
            & $msiScript @msiArgs
        }
        catch {
            $msiStopwatch.Stop()
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Red
            Write-Host "MSI Packaging Failed!" -ForegroundColor Red
            Write-Host "Error: $_" -ForegroundColor Red
            Write-Host "Packaging time: $($msiStopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Red
            Write-Host "========================================" -ForegroundColor Red
            exit 1
        }

        $msiStopwatch.Stop()
        Write-Host "MSI packaging completed successfully! ($($msiStopwatch.Elapsed.ToString('mm\:ss')))" -ForegroundColor Green

        Write-Host ""
        Write-Host "Building MSI symbols package (PDB)..." -ForegroundColor Yellow
        $msiSymbolsStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            $msiSymbolsArgs = @(
                "-Configuration", $Configuration,
                "-Platform", $Platform,
                "-BuildNumber", "$($versionContext.BuildNumber)"
            )
            & $msiSymbolsScript @msiSymbolsArgs
        }
        catch {
            $msiSymbolsStopwatch.Stop()
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Red
            Write-Host "MSI Symbols Packaging Failed!" -ForegroundColor Red
            Write-Host "Error: $_" -ForegroundColor Red
            Write-Host "Packaging time: $($msiSymbolsStopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Red
            Write-Host "========================================" -ForegroundColor Red
            exit 1
        }

        $msiSymbolsStopwatch.Stop()
        Write-Host "MSI symbols packaging completed successfully! ($($msiSymbolsStopwatch.Elapsed.ToString('mm\:ss')))" -ForegroundColor Green
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
        if (-not [string]::IsNullOrWhiteSpace($stagedMsixManifest)) {
            $stagedMsixManifestDirectory = Resolve-RSGuardedRepositoryPath `
                -RepoRoot $SolutionDir -GuardedRelativeRoot '.build\AppPackages' `
                -Path (Split-Path $stagedMsixManifest -Parent)
            if (Test-Path -LiteralPath $stagedMsixManifestDirectory -PathType Container) {
                Remove-Item -LiteralPath $stagedMsixManifestDirectory -Recurse -Force
            }
        }
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

    if ($Msix -or $Msi -or $Zip -or $GenerateWingetManifest) {
        if (-not (Test-Path -LiteralPath $appPackagesEvidenceRoot -PathType Container)) {
            throw "Requested packaging produced no AppPackages directory: $appPackagesEvidenceRoot"
        }
        foreach ($file in @(Get-ChildItem -LiteralPath $appPackagesEvidenceRoot -File -Recurse | Sort-Object FullName)) {
            $currentDigest = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            if (-not $packageEvidenceBefore.ContainsKey($file.FullName) -or $packageEvidenceBefore[$file.FullName] -ne $currentDigest) {
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
    $receiptArtifacts = @(@($buildArtifacts) + @($packageArtifacts) | Sort-Object relative_path)
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

    Write-Host ''
    foreach ($line in @(ConvertTo-RSBuildSummaryLines `
            -Configuration $Configuration `
            -Platform $Platform `
            -Duration $operationDuration `
            -Artifacts @($summaryArtifacts))) {
        $color = if ($line -like '& *' -or $line -like 'Output: *') { 'Cyan' } else { 'Green' }
        Write-Host $line -ForegroundColor $color
    }

    exit 0
}
catch {
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
    Exit-RSArtifactOperationLock -Lock $artifactOperationLock
}
