<#
.SYNOPSIS
    Invokes MSBuild after normalizing duplicate Path/PATH process environment entries.
.DESCRIPTION
    Some PowerShell-hosted sessions expose both Path and PATH in the process environment
    block. MSBuild later clones that environment when launching CL.exe, and the duplicate
    casing can trip ProcessStartInfo on Windows with:
      "Item has already been added. Key in dictionary: 'Path' Key being added: 'PATH'"

    This wrapper merges Path segments across Machine/User/Process, removes the duplicate
    PATH alias, resolves MSBuild via vswhere when not provided, and then forwards the
    remaining arguments to MSBuild unchanged.
.EXAMPLE
    .\Tools\Invoke-SanitizedMsbuild.ps1 Z:\src\RedSalamander\Common\DxUi\DxUi.vcxproj /t:Build /p:Configuration=Debug /p:Platform=x64
#>

[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true, Position = 0)]
    [string[]]$MSBuildArguments = @(),

    [string]$MSBuildPath = ""
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$sanitizedEnvironmentScript = Join-Path $PSScriptRoot 'SanitizedEnvironment.ps1'
$artifactOperationLockScript = Join-Path $PSScriptRoot 'ArtifactOperationLock.ps1'
if (-not (Test-Path $sanitizedEnvironmentScript)) {
    throw "Sanitized environment helper not found: $sanitizedEnvironmentScript"
}

. $sanitizedEnvironmentScript
if (-not (Test-Path $artifactOperationLockScript)) {
    throw "Artifact operation lock helper not found: $artifactOperationLockScript"
}
. $artifactOperationLockScript

function Resolve-MSBuildPath {
    if (-not [string]::IsNullOrWhiteSpace($MSBuildPath)) {
        return $MSBuildPath
    }

    $vswherePaths = @(
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\Installer\vswhere.exe"
    )

    $vswhere = $vswherePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($vswhere) {
        $instancesJson = & $vswhere -all -products "*" -prerelease -format json 2>$null
        if ($LASTEXITCODE -eq 0 -and $instancesJson) {
            $instances = @($instancesJson | ConvertFrom-Json) | Sort-Object {
                try {
                    [version]$_.installationVersion
                }
                catch {
                    [version]'0.0'
                }
            } -Descending

            foreach ($instance in $instances) {
                $installPath = $instance.installationPath
                if (-not $installPath) {
                    continue
                }

                $candidates = @(
                    (Join-Path $installPath "MSBuild\\Current\\Bin\\amd64\\MSBuild.exe"),
                    (Join-Path $installPath "MSBuild\\Current\\Bin\\MSBuild.exe")
                )

                $candidate = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
                if ($candidate) {
                    return $candidate
                }
            }
        }
    }

    $command = Get-Command msbuild.exe -ErrorAction SilentlyContinue
    if ($command -and $command.Source -and (Test-Path $command.Source)) {
        return $command.Source
    }

    throw "Unable to locate MSBuild.exe."
}

$configuration = ''
$platform = ''
$target = 'direct-msbuild'
foreach ($argument in $MSBuildArguments) {
    if ($argument -match '^(?i)/p:Configuration=(?<value>.+)$') {
        $configuration = $Matches.value.Trim('"')
    }
    elseif ($argument -match '^(?i)/p:Platform=(?<value>.+)$') {
        $platform = $Matches.value.Trim('"')
    }
    elseif ($target -eq 'direct-msbuild' -and $argument -notmatch '^[/-]') {
        $target = $argument
    }
}

$artifactOperationLock = $null
try {
    $artifactOperationLock = Enter-RSArtifactOperationLock `
        -RepoRoot $repoRoot `
        -Operation "direct MSBuild $target $configuration|$platform" `
        -Scope @{
            kind = 'build'
            target = $target
            configuration = $configuration
            platform = $platform
        }

    if ($artifactOperationLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $repoRoot `
                -Reason 'The previous build/test owner exited without clearing the exclusive artifact-operation lock.' `
                -AbandonedOwner $artifactOperationLock.AbandonedOwner)
    }
    if (Test-RSArtifactOperationContaminated -RepoRoot $repoRoot) {
        $markerPath = Get-RSArtifactContaminationMarkerPath -RepoRoot $repoRoot
        throw "Direct MSBuild cannot repair contaminated shared artifacts. Run a matching full-solution build.ps1 -Rebuild first. Marker: $markerPath"
    }

    Assert-RSNoResidualArtifactToolProcesses -RepoRoot $repoRoot
    $resolvedMsbuildPath = Resolve-MSBuildPath
    $exitCode = Invoke-RSProcess `
        -FilePath $resolvedMsbuildPath `
        -Arguments $MSBuildArguments `
        -WorkingDirectory $repoRoot
    exit $exitCode
}
finally {
    Exit-RSArtifactOperationLock -Lock $artifactOperationLock
}
