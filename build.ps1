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
$BuildProjectSelectionScript = Join-Path -Path $PSScriptRoot -ChildPath "Tools\BuildProjectSelection.ps1"
$MSBuildInvocationScript = Join-Path -Path $PSScriptRoot -ChildPath "Tools\MSBuildInvocation.ps1"
$ProcessStreamingScript = Join-Path -Path $PSScriptRoot -ChildPath "Tools\ProcessStreaming.ps1"
$SanitizedEnvironmentScript = Join-Path -Path $PSScriptRoot -ChildPath "Tools\SanitizedEnvironment.ps1"
$ArtifactOperationLockScript = Join-Path -Path $PSScriptRoot -ChildPath "Tools\ArtifactOperationLock.ps1"
if (-not (Test-Path $BuildProjectSelectionScript)) {
    Write-Error "Build project selection helper not found: $BuildProjectSelectionScript"
    exit 1
}
if (-not (Test-Path $MSBuildInvocationScript)) {
    Write-Error "MSBuild invocation helper not found: $MSBuildInvocationScript"
    exit 1
}
if (-not (Test-Path $ProcessStreamingScript)) {
    Write-Error "Process streaming helper not found: $ProcessStreamingScript"
    exit 1
}
if (-not (Test-Path $SanitizedEnvironmentScript)) {
    Write-Error "Sanitized environment helper not found: $SanitizedEnvironmentScript"
    exit 1
}
if (-not (Test-Path $ArtifactOperationLockScript)) {
    Write-Error "Artifact operation lock helper not found: $ArtifactOperationLockScript"
    exit 1
}

. $BuildProjectSelectionScript
. $MSBuildInvocationScript
. $SanitizedEnvironmentScript
. $ProcessStreamingScript
. $ArtifactOperationLockScript

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
    return Join-Path $logDir "$safePrefix-$timestamp.log"
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
$VersioningScript = Join-Path -Path $PSScriptRoot -ChildPath "Tools\Versioning.ps1"

Write-BuildBanner

# Validate solution file exists
if (-not (Test-Path $SolutionFile)) {
    Write-Error "Solution file not found: $SolutionFile"
    exit 1
}

if (-not (Test-Path $VersioningScript)) {
    Write-Error "Version helper script not found: $VersioningScript"
    exit 1
}

. $VersioningScript

$SolutionFullPath = (Resolve-Path $SolutionFile).Path
$SolutionDir = (Split-Path -Parent $SolutionFullPath)
$SolutionDirWithSlash = $SolutionDir.TrimEnd('\') + '\'
$artifactOperationLock = $null
$stopwatch = [System.Diagnostics.Stopwatch]::new()
$contaminationRepairAuthorized = $false
try {
    $operationTarget = if ($ProjectName) { $ProjectName } else { 'solution' }
    $artifactOperationLock = Enter-RSArtifactOperationLock `
        -RepoRoot $SolutionDir `
        -Operation "build $operationTarget $Configuration|$Platform" `
        -Scope @{
            kind = 'build'
            target = $operationTarget
            configuration = $Configuration
            platform = $Platform
        }

    if ($artifactOperationLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $SolutionDir `
                -Reason "The previous build/test owner exited without clearing the exclusive artifact-operation lock." `
                -AbandonedOwner $artifactOperationLock.AbandonedOwner)
    }

    $contamination = Read-RSArtifactOperationContamination -RepoRoot $SolutionDir
    if ($null -ne $contamination) {
        $contaminationRepairAuthorized = Test-RSArtifactOperationRepairAllowed `
            -Contamination $contamination `
            -Rebuild:$Rebuild `
            -ProjectName $ProjectName `
            -Configuration $Configuration `
            -Platform $Platform
        if (-not $contaminationRepairAuthorized) {
            $markerPath = Get-RSArtifactContaminationMarkerPath -RepoRoot $SolutionDir
            $scopeText = if ($null -ne $contamination.PSObject.Properties['abandoned_operation'] -and
                -not [string]::IsNullOrWhiteSpace([string]$contamination.abandoned_operation)) {
                " Previous operation: '$($contamination.abandoned_operation)'."
            } else { '' }
            throw ("Shared build artifacts may be mixed after an interrupted operation." +
                   "$scopeText Run a full-solution build.ps1 -Rebuild with the recorded configuration and platform; " +
                   "-Clean alone and targeted or wrong-scope rebuilds cannot clear this marker. Marker: $markerPath")
        }
    }

    Assert-RSNoResidualArtifactToolProcesses -RepoRoot $SolutionDir

$versionState = Use-RSVersionStateLock -RepoRoot $SolutionDir -ScriptBlock {
    $context = Get-RSVersionContext -RepoRoot $SolutionDir -Configuration $Configuration -Platform $Platform -BuildNumber $BuildNumber -OfficialRelease:$OfficialRelease
    $statePath = Save-RSVersionContext -RepoRoot $SolutionDir -VersionContext $context

    [pscustomobject]@{
        Context = $context
        StatePath = $statePath
    }
}
$versionContext = $versionState.Context
$versionStatePath = $versionState.StatePath

# Function to find MSBuild
function Find-MSBuild {
    Write-Host "Locating MSBuild..." -ForegroundColor Yellow

    # In GitHub Actions, prefer MSBuild from PATH. The workflow can install a newer VS toolchain and
    # prepend its MSBuild directory to PATH. This avoids accidentally picking the preinstalled VS 2022 instance.
    if ($env:GITHUB_ACTIONS -and ($env:GITHUB_ACTIONS -eq "true")) {
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
    "/m"              # Multi-processor build
    "/v:minimal"      # Minimal verbosity
    "/nologo"         # Suppress MSBuild banner
)

$buildParams += "/p:SolutionDir=$SolutionDirWithSlash"
$buildParams += "/p:RSVersionBuildNumber=$($versionContext.BuildNumber)"
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

function Test-BuildOutputSelfTestCommandLine {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$CommandLine
    )

    if ([string]::IsNullOrWhiteSpace($CommandLine)) {
        return $false
    }

    return $CommandLine -match '(?i)(?:^|[\s"])--[a-z0-9-]*selftest[a-z0-9-]*(?:=[^\s"]*)?(?=$|[\s"])'
}

function Stop-BuildOutputProcess {
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
        return
    }

    $escapedName = $ProcessName.Replace("'", "''")
    $processes = @()
    try {
        $processes = Get-CimInstance Win32_Process -Filter "Name='$escapedName'" -ErrorAction SilentlyContinue
    }
    catch {
        return
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

    $protectedProcesses = @($matchingProcesses | Where-Object {
        [string]::IsNullOrWhiteSpace($_.CommandLine) -or (Test-BuildOutputSelfTestCommandLine -CommandLine $_.CommandLine)
    })
    if ($protectedProcesses.Count -gt 0) {
        $diagnostics = @($protectedProcesses | ForEach-Object {
            $commandLine = if ([string]::IsNullOrWhiteSpace($_.CommandLine)) { '<unavailable>' } else { $_.CommandLine }
            "  PID=$($_.ProcessId); Path='$($_.ExecutablePath)'; CommandLine='$commandLine'"
        })
        throw ("Build canceled because an active self-test may be using a target output. " +
               "Wait for the self-test to finish before rebuilding.`n" +
               ($diagnostics -join "`n"))
    }

    foreach ($proc in $matchingProcesses) {
        Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue
    }
}

    $buildOutputDir = Join-Path -Path $SolutionDir -ChildPath (".build\\{0}\\{1}" -f $Platform, $Configuration)
    Stop-BuildOutputProcess -ProcessName "RedSalamander.exe" -ExpectedExePath (Join-Path -Path $buildOutputDir -ChildPath "RedSalamander.exe")
    Stop-BuildOutputProcess -ProcessName "RedSalamanderMonitor.exe" -ExpectedExePath (Join-Path -Path $buildOutputDir -ChildPath "RedSalamanderMonitor.exe")

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
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Build completed successfully!" -ForegroundColor Green
    Write-Host "Configuration: $Configuration | Platform: $Platform" -ForegroundColor Green
    Write-Host "Build time: $($stopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    
    # Show output paths
    if ($ProjectName) {
        $isLanguageResourceProject = $false
        if ($resolvedProjectPath) {
            $normalizedProjectPath = $resolvedProjectPath.Replace('/', '\')
            $isLanguageResourceProject = $normalizedProjectPath -match '\\Lang\\[^\\]+\\[^\\]+\.vcxproj$'
        }

        if ($isLanguageResourceProject) {
            $languageOutput = ".build\\$Platform\\$Configuration\\Lang\\$ProjectName.dll"
            if (-not (Test-Path $languageOutput)) {
                Write-Host ""
                Write-Host "========================================" -ForegroundColor Red
                Write-Host "Language Resource Output Validation Failed!" -ForegroundColor Red
                Write-Host "Expected: $languageOutput" -ForegroundColor Red
                Write-Host "========================================" -ForegroundColor Red
                exit 1
            }

            $fileSize = (Get-Item $languageOutput).Length
            $fileSizeMB = [math]::Round($fileSize / 1MB, 2)
            Write-Host "Output: $languageOutput ($fileSizeMB MB)" -ForegroundColor Cyan
            Write-Host "Language resource output validated in Lang folder." -ForegroundColor Cyan
        }

        # Show specific project output
        $outputCandidates = if ($isLanguageResourceProject) {
            @()
        } elseif ($resolvedProjectPath -and $resolvedProjectPath -like '*\Plugins\*') {
            @(
                ".build\\$Platform\\$Configuration\\Plugins\\$ProjectName.exe",
                ".build\\$Platform\\$Configuration\\Plugins\\$ProjectName.dll",
                ".build\\$Platform\\$Configuration\\$ProjectName.exe",
                ".build\\$Platform\\$Configuration\\$ProjectName.dll"
            )
        } else {
            @(
                ".build\\$Platform\\$Configuration\\$ProjectName.exe",
                ".build\\$Platform\\$Configuration\\$ProjectName.dll",
                ".build\\$Platform\\$Configuration\\Plugins\\$ProjectName.exe",
                ".build\\$Platform\\$Configuration\\Plugins\\$ProjectName.dll"
            )
        }

        foreach ($candidate in $outputCandidates) {
            if (Test-Path $candidate) {
                $fileSize = (Get-Item $candidate).Length
                $fileSizeMB = [math]::Round($fileSize / 1MB, 2)
                Write-Host "Output: $candidate ($fileSizeMB MB)" -ForegroundColor Cyan
                break
            }
        }
    } else {
        # Show output paths for main executables
        $mainProjects = @("RedLauncher", "RedSalamander", "RedSalamanderMonitor")
        foreach ($project in $mainProjects) {
            $outputPath = ".build\\$Platform\\$Configuration\\$project.exe"
            if (Test-Path $outputPath) {
                $fileSize = (Get-Item $outputPath).Length
                $fileSizeMB = [math]::Round($fileSize / 1MB, 2)
                Write-Host "Output: $outputPath ($fileSizeMB MB)" -ForegroundColor Cyan
            }
        }
    }

    if ($Msix) {
        $msixAssetsScript = Join-Path -Path $SolutionDir -ChildPath "Installer\msix\GenerateAssets.ps1"
        if (-not (Test-Path $msixAssetsScript)) {
            Write-Error "MSIX assets generation script not found: $msixAssetsScript"
            exit 1
        }

        Write-Host ""
        Write-Host "Generating MSIX assets..." -ForegroundColor Yellow
        try {
            & $msixAssetsScript
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
            & $msixVersionScript -Version $versionContext.PackagingVersion -Platform $Platform
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

        $appPackagesDir = Join-Path -Path $SolutionDir -ChildPath ".build\\AppPackages"
        if (Test-Path $appPackagesDir) {
            $msixFiles = Get-ChildItem -Path $appPackagesDir -Filter *.msix -Recurse -ErrorAction SilentlyContinue |
                Sort-Object -Property LastWriteTime -Descending |
                Select-Object -First 5

            foreach ($msixFile in $msixFiles) {
                $relativePath = if ($msixFile.FullName.StartsWith($SolutionDirWithSlash, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $msixFile.FullName.Substring($SolutionDirWithSlash.Length)
                } else {
                    $msixFile.FullName
                }
                $fileSizeMB = [math]::Round($msixFile.Length / 1MB, 2)
                Write-Host "Output: $relativePath ($fileSizeMB MB)" -ForegroundColor Cyan
            }
        }
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

        $appPackagesDir = Join-Path -Path $SolutionDir -ChildPath ".build\\AppPackages"
        if (Test-Path $appPackagesDir) {
            $msiFiles = Get-ChildItem -Path $appPackagesDir -Filter *.msi -Recurse -ErrorAction SilentlyContinue |
                Sort-Object -Property LastWriteTime -Descending |
                Select-Object -First 5

            foreach ($msiFile in $msiFiles) {
                $relativePath = if ($msiFile.FullName.StartsWith($SolutionDirWithSlash, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $msiFile.FullName.Substring($SolutionDirWithSlash.Length)
                } else {
                    $msiFile.FullName
                }
                $fileSizeMB = [math]::Round($msiFile.Length / 1MB, 2)
                Write-Host "Output: $relativePath ($fileSizeMB MB)" -ForegroundColor Cyan
            }
        }
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

        $appPackagesDir = Join-Path -Path $SolutionDir -ChildPath ".build\\AppPackages"
        if (Test-Path $appPackagesDir) {
            $zipFiles = Get-ChildItem -Path $appPackagesDir -Filter *.zip -Recurse -ErrorAction SilentlyContinue |
                Sort-Object -Property LastWriteTime -Descending |
                Select-Object -First 5

            foreach ($zipFile in $zipFiles) {
                $relativePath = if ($zipFile.FullName.StartsWith($SolutionDirWithSlash, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $zipFile.FullName.Substring($SolutionDirWithSlash.Length)
                } else {
                    $zipFile.FullName
                }
                $fileSizeMB = [math]::Round($zipFile.Length / 1MB, 2)
                Write-Host "Output: $relativePath ($fileSizeMB MB)" -ForegroundColor Cyan
            }
        }
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
        $ZipPath = Get-ChildItem -Path $appPackagesDir -Filter "RedSalamander-*-x64-Portable.zip" -Recurse -ErrorAction SilentlyContinue |
            Sort-Object -Property LastWriteTime -Descending |
            Select-Object -First 1 -ExpandProperty FullName
        $Arm64ZipPath = Get-ChildItem -Path $appPackagesDir -Filter "RedSalamander-*-ARM64-Portable.zip" -Recurse -ErrorAction SilentlyContinue |
            Sort-Object -Property LastWriteTime -Descending |
            Select-Object -First 1 -ExpandProperty FullName

        try {
            $wingetParams = @{}
            $wingetParams['BuildNumber'] = $versionContext.BuildNumber
            if ($ZipPath) { $wingetParams['ZipPath'] = $ZipPath }
            if ($Arm64ZipPath) { $wingetParams['Arm64ZipPath'] = $Arm64ZipPath }
            
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
    
    if ($contaminationRepairAuthorized) {
        Clear-RSArtifactOperationContaminated -RepoRoot $SolutionDir
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
