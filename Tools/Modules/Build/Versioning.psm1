# Private build/release module; import through its reviewed callers.
Set-StrictMode -Version Latest

function Get-RSVersionLockPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Join-Path (Get-RSVersionStateDirectory -RepoRoot $RepoRoot) 'version-state.lock'
}

function Use-RSVersionStateLock {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $stateDirectory = Get-RSVersionStateDirectory -RepoRoot $RepoRoot
    [void](New-Item -ItemType Directory -Path $stateDirectory -Force)
    $lockPath = Get-RSVersionLockPath -RepoRoot $RepoRoot
    $lockStream = $null
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    try {
        while ($null -eq $lockStream -and $stopwatch.Elapsed -lt [TimeSpan]::FromSeconds(30)) {
            try {
                # The repository-local file lease works across Windows sessions and
                # through subst/junction aliases because every alias opens the same
                # physical file. A crashed owner releases the kernel file handle.
                $lockStream = [IO.File]::Open(
                    $lockPath,
                    [IO.FileMode]::OpenOrCreate,
                    [IO.FileAccess]::ReadWrite,
                    [IO.FileShare]::None)
            }
            catch [IO.IOException] {
                Start-Sleep -Milliseconds 50
            }
        }
        if ($null -eq $lockStream) {
            throw "Timed out waiting for repository version-state lease '$lockPath'."
        }

        return & $ScriptBlock
    }
    finally {
        $stopwatch.Stop()
        if ($null -ne $lockStream) {
            $lockStream.Dispose()
        }
    }
}

function Get-RSVersionHeaderPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Join-Path $RepoRoot "Common\Version.h"
}

function Get-RSVersionBaseInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $headerPath = Get-RSVersionHeaderPath -RepoRoot $RepoRoot
    $content = Get-Content -Path $headerPath -Raw

    $readDefine = {
        param([string]$name)

        $pattern = "(?m)^\s*#define\s+$([Regex]::Escape($name))\s+(\d+)\s*$"
        $match = [Regex]::Match($content, $pattern)
        if (-not $match.Success) {
            throw "Failed to find $name in $headerPath"
        }

        return [int]$match.Groups[1].Value
    }

    $major = & $readDefine "VERSINFO_MAJOR"
    $minor = & $readDefine "VERSINFO_MINOR"
    $displayBaseVersion = "$major.$minor"

    [pscustomobject]@{
        HeaderPath = $headerPath
        Major = $major
        Minor = $minor
        DisplayBaseVersion = $displayBaseVersion
        LastVersionOfSalamander = ($major * 100) + $minor
    }
}

function Get-RSVersionStateDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Join-Path $RepoRoot ".build\version"
}

function Get-RSVersionStatePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Join-Path (Get-RSVersionStateDirectory -RepoRoot $RepoRoot) "current-version.json"
}

function Get-RSLocalBuildCounterPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Join-Path (Get-RSVersionStateDirectory -RepoRoot $RepoRoot) "local-build-counter.txt"
}

function Write-RSAtomicVersionText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [System.Text.Encoding]$Encoding
    )

    $directory = Split-Path -Parent $Path
    [void](New-Item -ItemType Directory -Path $directory -Force)
    $temporaryPath = Join-Path $directory ('.{0}.{1}.{2}.tmp' -f
        ([IO.Path]::GetFileName($Path)), $PID, [Guid]::NewGuid().ToString('N'))
    $backupPath = "$temporaryPath.bak"
    try {
        [IO.File]::WriteAllText($temporaryPath, $Text, $Encoding)
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            # A concrete same-directory backup avoids PowerShell converting a
            # null backup argument to an invalid empty path on Windows PowerShell.
            Invoke-RSVersionFileReplace -SourcePath $temporaryPath -DestinationPath $Path -BackupPath $backupPath
        }
        else {
            [IO.File]::Move($temporaryPath, $Path)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
        if (Test-Path -LiteralPath $backupPath) {
            Remove-Item -LiteralPath $backupPath -Force
        }
    }
}

function Invoke-RSVersionFileReplace {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    [IO.File]::Replace($SourcePath, $DestinationPath, $BackupPath)
}

function Test-RSVersionJsonInteger {
    param(
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [long]$Minimum,

        [Parameter(Mandatory = $true)]
        [long]$Maximum
    )

    if ($null -eq $Value) { return $false }
    $integralTypes = @(
        [byte], [sbyte], [int16], [uint16], [int32], [uint32], [int64]
    )
    if ($integralTypes -notcontains $Value.GetType()) { return $false }
    $number = [long]$Value
    return $number -ge $Minimum -and $number -le $Maximum
}

function New-RSLocalBuildNumberUnlocked {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $stateDir = Get-RSVersionStateDirectory -RepoRoot $RepoRoot
    New-Item -ItemType Directory -Path $stateDir -Force | Out-Null

    $counterPath = Get-RSLocalBuildCounterPath -RepoRoot $RepoRoot
    $lastBuildNumber = 0
    if (Test-Path -LiteralPath $counterPath -PathType Leaf) {
        $rawText = Get-Content -LiteralPath $counterPath -Raw -ErrorAction Stop
        $raw = if ($null -eq $rawText) { '' } else { $rawText.Trim() }
        $parsed = 0
        if ([string]::IsNullOrWhiteSpace($raw) -or
            -not [int]::TryParse($raw, [ref]$parsed) -or
            $parsed -lt 1 -or
            $parsed -ge 65535) {
            throw "Invalid local build counter in ${counterPath}: '$raw'"
        }
        $lastBuildNumber = $parsed
    }

    $nextBuildNumber = $lastBuildNumber + 1
    Write-RSAtomicVersionText `
        -Path $counterPath `
        -Text $nextBuildNumber.ToString([Globalization.CultureInfo]::InvariantCulture) `
        -Encoding ([Text.Encoding]::ASCII)
    return $nextBuildNumber
}

function Read-RSVersionContextUnlocked {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $statePath = Get-RSVersionStatePath -RepoRoot $RepoRoot
    if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) {
        return $null
    }

    try {
        $context = Get-Content -LiteralPath $statePath -Raw -ErrorAction Stop |
            ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Invalid version state in '$statePath': $($_.Exception.Message)"
    }

    foreach ($requiredName in @('Major', 'Minor', 'BuildNumber', 'Configuration', 'Platform', 'OfficialRelease')) {
        if ($null -eq $context.PSObject.Properties[$requiredName]) {
            throw "Invalid version state in '$statePath': missing '$requiredName'."
        }
    }
    if (-not (Test-RSVersionJsonInteger -Value $context.Major -Minimum 0 -Maximum 65535) -or
        -not (Test-RSVersionJsonInteger -Value $context.Minor -Minimum 0 -Maximum 65535) -or
        -not (Test-RSVersionJsonInteger -Value $context.BuildNumber -Minimum 1 -Maximum 65535) -or
        @('Debug', 'ASan Debug', 'Release') -cnotcontains [string]$context.Configuration -or
        @('x64', 'ARM64') -cnotcontains [string]$context.Platform -or
        $context.OfficialRelease -isnot [bool]) {
        throw "Invalid version state in '$statePath': version components, profile metadata, or release kind has an invalid type or value."
    }
    return $context
}

function Resolve-RSBuildNumberUnlocked {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [int]$BuildNumber = 0,

        [string]$Configuration = "",

        [string]$Platform = "",

        [switch]$OfficialRelease
    )

    if ($BuildNumber -gt 0) {
        return $BuildNumber
    }

    if ($env:GITHUB_RUN_NUMBER) {
        $ciBuildNumber = 0
        if ([int]::TryParse($env:GITHUB_RUN_NUMBER, [ref]$ciBuildNumber) -and $ciBuildNumber -gt 0) {
            return $ciBuildNumber
        }
    }

    if ($OfficialRelease) {
        throw "Official release builds require an explicit build number or a CI environment with GITHUB_RUN_NUMBER. Local beta builds continue to use a local per-worktree counter."
    }

    $savedContext = Read-RSVersionContextUnlocked -RepoRoot $RepoRoot
    if ($savedContext) {
        $sameReleaseKind = ([bool]$savedContext.OfficialRelease) -eq ([bool]$OfficialRelease)
        if ($sameReleaseKind) {
            return [int]$savedContext.BuildNumber
        }
    }

    return New-RSLocalBuildNumberUnlocked -RepoRoot $RepoRoot
}

function Get-RSVersionContextUnlocked {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [int]$BuildNumber = 0,

        [switch]$OfficialRelease
    )

    $baseInfo = Get-RSVersionBaseInfo -RepoRoot $RepoRoot
    $resolvedBuildNumber = Resolve-RSBuildNumberUnlocked -RepoRoot $RepoRoot -BuildNumber $BuildNumber -Configuration $Configuration -Platform $Platform -OfficialRelease:$OfficialRelease
    if ($baseInfo.Major -gt 65535 -or $baseInfo.Minor -gt 65535 -or $resolvedBuildNumber -gt 65535) {
        throw "Version components must fit Windows VERSIONINFO 16-bit fields: $($baseInfo.Major).$($baseInfo.Minor).$resolvedBuildNumber"
    }

    $platformLabel = if ($Platform -eq "ARM64") { "ARM64" } else { "x64" }
    $configurationLabel = if ($Configuration -eq "ASan Debug") { "ASan Debug" } else { $Configuration }
    $displayVersion = if ($OfficialRelease) {
        "$($baseInfo.DisplayBaseVersion).$resolvedBuildNumber"
    } else {
        "$($baseInfo.DisplayBaseVersion).$resolvedBuildNumber beta $configurationLabel ($platformLabel)"
    }

    [pscustomobject]@{
        Major = $baseInfo.Major
        Minor = $baseInfo.Minor
        BuildNumber = $resolvedBuildNumber
        Configuration = $Configuration
        ConfigurationLabel = $configurationLabel
        Platform = $Platform
        PlatformLabel = $platformLabel
        OfficialRelease = [bool]$OfficialRelease
        DisplayBaseVersion = $baseInfo.DisplayBaseVersion
        DisplayVersion = $displayVersion
        PackagingVersion = "$($baseInfo.Major).$($baseInfo.Minor).$resolvedBuildNumber"
        LastVersionOfSalamander = $baseInfo.LastVersionOfSalamander
        HeaderPath = $baseInfo.HeaderPath
    }
}

function Save-RSVersionContextUnlocked {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [psobject]$VersionContext
    )

    $stateDir = Get-RSVersionStateDirectory -RepoRoot $RepoRoot
    New-Item -ItemType Directory -Path $stateDir -Force | Out-Null

    $statePath = Get-RSVersionStatePath -RepoRoot $RepoRoot
    $json = $VersionContext | ConvertTo-Json -Depth 4
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    Write-RSAtomicVersionText -Path $statePath -Text $json -Encoding $utf8NoBom
    return $statePath
}

function New-RSVersionContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Debug', 'ASan Debug', 'Release')]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 65535)]
        [int]$BuildNumber,

        [switch]$OfficialRelease
    )

    # An explicit build number makes this a pure operation: package callers can
    # construct metadata without reading, allocating, or persisting version state.
    return Get-RSVersionContextUnlocked `
        -RepoRoot $RepoRoot `
        -Configuration $Configuration `
        -Platform $Platform `
        -BuildNumber $BuildNumber `
        -OfficialRelease:$OfficialRelease
}

function Resolve-RSVersionContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Debug', 'ASan Debug', 'Release')]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [ValidateRange(0, 65535)]
        [int]$BuildNumber = 0,

        [switch]$OfficialRelease
    )

    return Use-RSVersionStateLock -RepoRoot $RepoRoot -ScriptBlock {
        $context = Get-RSVersionContextUnlocked `
            -RepoRoot $RepoRoot `
            -Configuration $Configuration `
            -Platform $Platform `
            -BuildNumber $BuildNumber `
            -OfficialRelease:$OfficialRelease
        [void](Save-RSVersionContextUnlocked -RepoRoot $RepoRoot -VersionContext $context)
        return $context
    }
}

function Read-RSVersionContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Use-RSVersionStateLock -RepoRoot $RepoRoot -ScriptBlock {
        Read-RSVersionContextUnlocked -RepoRoot $RepoRoot
    }
}

Export-ModuleMember -Function @(
    'Get-RSVersionStatePath',
    'New-RSVersionContext',
    'Resolve-RSVersionContext',
    'Read-RSVersionContext'
)
