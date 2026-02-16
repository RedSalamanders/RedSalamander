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
    [switch]$Msi
)

$ErrorActionPreference = "Stop"

# Validate packaging options early so we fail fast before attempting a build.
if ($Msix -and $Msi) {
    Write-Error "Specify only one of -Msix or -Msi."
    exit 1
}

if (($Msix -or $Msi) -and $ProjectName) {
    Write-Error "Packaging requires building the full solution. Remove -ProjectName."
    exit 1
}

$packageMode = if ($Msix) { "MSIX" } elseif ($Msi) { "MSI" } else { "None" }

if ($Msix -or $Msi) {
    if (-not $PSBoundParameters.ContainsKey("Configuration")) {
        $Configuration = "Release"
    } elseif ($Configuration -ne "Release") {
        Write-Error "Packaging requires -Configuration Release."
        exit 1
    }
}

# Script constants
$SolutionFile = Join-Path -Path $PSScriptRoot -ChildPath "RedSalamander.sln"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "RedSalamander Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Validate solution file exists
if (-not (Test-Path $SolutionFile)) {
    Write-Error "Solution file not found: $SolutionFile"
    exit 1
}

$SolutionFullPath = (Resolve-Path $SolutionFile).Path
$SolutionDir = (Split-Path -Parent $SolutionFullPath)
$SolutionDirWithSlash = $SolutionDir.TrimEnd('\') + '\'

function Resolve-ProjectFileFromSolution {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SolutionPath,

        [Parameter(Mandatory = $true)]
        [string]$ProjectName
    )

    $solutionText = Get-Content -Path $SolutionPath -Raw
    $pattern = 'Project\("\{[0-9A-Fa-f-]+\}"\)\s*=\s*"' + [Regex]::Escape($ProjectName) + '",\s*"([^"]+)"'
    $match = [Regex]::Match($solutionText, $pattern)
    if (-not $match.Success) {
        throw "Project '$ProjectName' not found in solution '$SolutionPath'."
    }

    $relativePath = $match.Groups[1].Value
    $projectPath = Join-Path -Path (Split-Path -Parent $SolutionPath) -ChildPath $relativePath
    if (-not (Test-Path $projectPath)) {
        throw "Project file not found: $projectPath"
    }

    return (Resolve-Path $projectPath).Path
}

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
            if (($fileMajor -ne $null) -and ($fileMajor -ge 18) -and ($candidatePath -match '\\Microsoft Visual Studio\\')) {
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
        throw "Windows SDK folder containing 'UAP.props' was not found. Install Windows 10/11 SDK (UAP) or set TargetPlatformVersion to an installed version."
    }

    $selected = $installed | Sort-Object { [version]$_ } -Descending | Select-Object -First 1
    if ($requested) {
        Write-Host "Windows SDK for UAP $requested not found; using installed UAP $selected." -ForegroundColor Yellow
    }

    return $selected
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
Write-Host "  Action:        $buildTarget"
Write-Host "  Package:       $packageMode"
Write-Host ""

# Resolve project build input (building a .vcxproj needs SolutionDir for include paths using $(SolutionDir)).
$isProjectBuild = $false
$buildInput = $SolutionFile
if ($ProjectName) {
    $buildInput = Resolve-ProjectFileFromSolution -SolutionPath $SolutionFullPath -ProjectName $ProjectName
    $isProjectBuild = $true
}

$msbuildTarget = if ($Rebuild) { "Rebuild" } else { "Build" }

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

if ($isProjectBuild) {
    $buildParams += "/p:SolutionDir=$SolutionDirWithSlash"
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
    if ($isProjectBuild) {
        $cleanParams += "/p:SolutionDir=$SolutionDirWithSlash"
    }
}

# Start build
Write-Host "Starting build..." -ForegroundColor Yellow
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

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

        Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue
    }
}

$buildOutputDir = Join-Path -Path $SolutionDir -ChildPath (".build\\{0}\\{1}" -f $Platform, $Configuration)
Stop-BuildOutputProcess -ProcessName "RedSalamander.exe" -ExpectedExePath (Join-Path -Path $buildOutputDir -ChildPath "RedSalamander.exe")
Stop-BuildOutputProcess -ProcessName "RedSalamanderMonitor.exe" -ExpectedExePath (Join-Path -Path $buildOutputDir -ChildPath "RedSalamanderMonitor.exe")

try {
    # Execute clean if requested
    if ($Clean) {
        Write-Host "Cleaning..." -ForegroundColor Yellow
        & $msbuildPath $cleanParams
        
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Clean failed, continuing with build..." -ForegroundColor Yellow
        }
        Write-Host "Building..." -ForegroundColor Yellow
    }
    
    & $msbuildPath $buildParams
    
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
        # Show specific project output
        $outputCandidates = @(
            ".build\\$Platform\\$Configuration\\$ProjectName.exe",
            ".build\\$Platform\\$Configuration\\$ProjectName.dll",
            ".build\\$Platform\\$Configuration\\Plugins\\$ProjectName.exe",
            ".build\\$Platform\\$Configuration\\Plugins\\$ProjectName.dll"
        )

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
        $mainProjects = @("RedSalamander", "RedSalamanderMonitor")
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

        & $msbuildPath $msixParams

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
            & $msiScript -Configuration $Configuration -Platform $Platform
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
            & $msiSymbolsScript -Configuration $Configuration -Platform $Platform
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
