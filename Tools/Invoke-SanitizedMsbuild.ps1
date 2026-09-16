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
    remaining arguments to MSBuild unchanged. Artifact-mutating invocations must include
    explicit /p:Configuration and /p:Platform properties so coordination is scoped to
    the exact repository output profile.
.PARAMETER MSBuildArguments
    Supplies the project/solution, targets, properties, and other arguments forwarded
    unchanged to MSBuild. The list must include explicit Configuration and Platform
    MSBuild properties, either as separate switches or one semicolon/comma-delimited
    property list.
.PARAMETER MSBuildPath
    Overrides MSBuild discovery with an explicit executable path.
.OUTPUTS
    Streams MSBuild console output and exits with the child MSBuild exit code; no supported pipeline objects are returned.
.NOTES
    Prerequisites: MSBuild, explicit Configuration and Platform properties, the repository
    build helpers, and uncontaminated profile/shared-dependency artifact scopes. Side
    effects: normalizes the current process Path environment, acquires scoped coordination
    locks, launches contained MSBuild, and writes ordinary build outputs/logs. Primary use
    is direct build diagnosis; build.ps1 remains the canonical repository build entrypoint.
.EXAMPLE
    .\Tools\Invoke-SanitizedMsbuild.ps1 Z:\src\RedSalamander\Tests\ProductUiTests\ProductUiTests.vcxproj /t:Build /p:Configuration=Debug /p:Platform=x64
#>

[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true, Position = 0)]
    [string[]]$MSBuildArguments = @(),

    [string]$MSBuildPath = ""
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$sanitizedEnvironmentModule = Join-Path $PSScriptRoot 'Modules\Build\SanitizedEnvironment.psm1'
$artifactOperationLockModule = Join-Path $PSScriptRoot 'Modules\Build\ArtifactOperationLock.psm1'
if (-not (Test-Path $sanitizedEnvironmentModule)) {
    throw "Sanitized environment helper not found: $sanitizedEnvironmentModule"
}

Import-Module $sanitizedEnvironmentModule -Force -ErrorAction Stop
if (-not (Test-Path $artifactOperationLockModule)) {
    throw "Artifact operation lock helper not found: $artifactOperationLockModule"
}
Import-Module $artifactOperationLockModule -Force -ErrorAction Stop

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

$configuration = [string](Get-RSMSBuildPropertyValue `
        -Arguments $MSBuildArguments `
        -Name 'Configuration')
$platform = [string](Get-RSMSBuildPropertyValue `
        -Arguments $MSBuildArguments `
        -Name 'Platform')
$target = 'direct-msbuild'
foreach ($argument in $MSBuildArguments) {
    if ($target -eq 'direct-msbuild' -and $argument -notmatch '^[/-]') {
        $target = $argument
    }
}

$artifactOperationLock = $null
$sharedDependencyLock = $null
try {
    $artifactOperationScope = @{
        kind = 'build'
        target = $target
        configuration = $configuration
        platform = $platform
    }
    $artifactOperationLock = Enter-RSArtifactOperationLock `
        -RepoRoot $repoRoot `
        -Operation "direct MSBuild $target $configuration|$platform" `
        -Scope $artifactOperationScope

    if ($artifactOperationLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $repoRoot `
                -Reason 'The previous build/test owner exited without clearing the exclusive artifact-operation lock.' `
                -AbandonedOwner $artifactOperationLock.AbandonedOwner `
                -Scope $artifactOperationScope)
    }
    if (Test-RSArtifactOperationContaminated -RepoRoot $repoRoot -Scope $artifactOperationScope) {
        $markerPath = Get-RSArtifactContaminationMarkerPath `
            -RepoRoot $repoRoot `
            -Scope $artifactOperationScope
        throw "Direct MSBuild cannot repair contaminated shared artifacts. Run a matching full-solution build.ps1 -Rebuild first. Marker: $markerPath"
    }

    Assert-RSNoResidualArtifactToolProcesses `
        -RepoRoot $repoRoot `
        -Scope $artifactOperationScope
    $sharedDependencyScope = $artifactOperationScope.Clone()
    $sharedDependencyScope['coordination'] = 'shared-dependency'
    $sharedDependencyLock = Enter-RSArtifactOperationLock `
        -RepoRoot $repoRoot `
        -Operation "direct MSBuild shared dependency phase $configuration|$platform" `
        -Scope $sharedDependencyScope
    if ($sharedDependencyLock.WasAbandoned) {
        [void](Set-RSArtifactOperationContaminated `
                -RepoRoot $repoRoot `
                -Reason 'The previous MSBuild owner exited while it could have been writing shared vcpkg dependencies.' `
                -AbandonedOwner $sharedDependencyLock.AbandonedOwner `
                -Scope $sharedDependencyScope)
    }
    if (Test-RSArtifactOperationContaminated -RepoRoot $repoRoot -Scope $sharedDependencyScope) {
        $markerPath = Get-RSArtifactContaminationMarkerPath `
            -RepoRoot $repoRoot `
            -Scope $sharedDependencyScope
        throw "Direct MSBuild cannot repair interrupted shared vcpkg dependencies. Run a full-solution build.ps1 -Rebuild for platform $platform. Marker: $markerPath"
    }
    Assert-RSNoResidualArtifactToolProcesses `
        -RepoRoot $repoRoot `
        -Scope $sharedDependencyScope

    $resolvedMsbuildPath = Resolve-MSBuildPath
    $exitCode = Invoke-RSProcess `
        -FilePath $resolvedMsbuildPath `
        -Arguments $MSBuildArguments `
        -WorkingDirectory $repoRoot
    exit $exitCode
}
finally {
    Exit-RSArtifactOperationLock -Lock $sharedDependencyLock
    Exit-RSArtifactOperationLock -Lock $artifactOperationLock
}
