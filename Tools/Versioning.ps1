Set-StrictMode -Version Latest

function Get-RSVersionLockName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $normalizedRoot = [System.IO.Path]::GetFullPath($RepoRoot).TrimEnd('\').ToUpperInvariant()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalizedRoot)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $sha256.ComputeHash($bytes)
    }
    finally {
        $sha256.Dispose()
    }
    $hash = ([System.BitConverter]::ToString($hashBytes)).Replace("-", "")
    return "Local\RedSalamander.Versioning.$hash"
}

function Use-RSVersionStateLock {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $mutexName = Get-RSVersionLockName -RepoRoot $RepoRoot
    $mutex = [System.Threading.Mutex]::new($false, $mutexName)
    $lockTaken = $false
    try {
        $lockTaken = $mutex.WaitOne([TimeSpan]::FromSeconds(30))
        if (-not $lockTaken) {
            throw "Timed out waiting for version state lock '$mutexName'"
        }

        return & $ScriptBlock
    }
    finally {
        if ($lockTaken) {
            $mutex.ReleaseMutex() | Out-Null
        }
        $mutex.Dispose()
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

function New-RSLocalBuildNumber {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $stateDir = Get-RSVersionStateDirectory -RepoRoot $RepoRoot
    New-Item -ItemType Directory -Path $stateDir -Force | Out-Null

    $counterPath = Get-RSLocalBuildCounterPath -RepoRoot $RepoRoot
    $lastBuildNumber = 0
    if (Test-Path $counterPath) {
        $raw = (Get-Content -Path $counterPath -Raw).Trim()
        if ($raw) {
            $parsed = 0
            if (-not [int]::TryParse($raw, [ref]$parsed)) {
                throw "Invalid local build counter in ${counterPath}: '$raw'"
            }
            $lastBuildNumber = $parsed
        }
    }

    $nextBuildNumber = $lastBuildNumber + 1
    Set-Content -Path $counterPath -Value $nextBuildNumber -Encoding ASCII -NoNewline
    return $nextBuildNumber
}

function Resolve-RSBuildNumber {
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

    $savedContext = Read-RSVersionContext -RepoRoot $RepoRoot
    if ($savedContext) {
        $sameReleaseKind = ([bool]$savedContext.OfficialRelease) -eq ([bool]$OfficialRelease)
        if ($sameReleaseKind) {
            return [int]$savedContext.BuildNumber
        }
    }

    return New-RSLocalBuildNumber -RepoRoot $RepoRoot
}

function Get-RSVersionContext {
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
    $resolvedBuildNumber = Resolve-RSBuildNumber -RepoRoot $RepoRoot -BuildNumber $BuildNumber -Configuration $Configuration -Platform $Platform -OfficialRelease:$OfficialRelease
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

function Save-RSVersionContext {
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
    [System.IO.File]::WriteAllText($statePath, $json, $utf8NoBom)
    return $statePath
}

function Read-RSVersionContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $statePath = Get-RSVersionStatePath -RepoRoot $RepoRoot
    if (-not (Test-Path $statePath)) {
        return $null
    }

    return Get-Content -Path $statePath -Raw | ConvertFrom-Json
}
