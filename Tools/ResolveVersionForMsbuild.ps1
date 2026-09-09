<#
.SYNOPSIS
Resolves and persists the build number used by MSBuild.

.DESCRIPTION
Loads the repository versioning implementation under its version-state lock, resolves the configuration/platform build context, saves the local version state, and writes only the numeric build number to standard output for MSBuild consumption.

.PARAMETER RepoRoot
Specifies the repository root containing the private Tools/Modules/Build/Versioning.psm1 module.

.PARAMETER Configuration
Specifies the MSBuild configuration whose version context is being resolved.

.PARAMETER Platform
Specifies the MSBuild platform whose version context is being resolved.

.PARAMETER OfficialRelease
Uses the official-release version policy instead of the ordinary local-build policy.

.OUTPUTS
Writes the numeric build number without an added newline to standard output.

.NOTES
Prerequisites: a valid repository checkout and writable repository version-state location. Side effects: updates the serialized local version context under the repository build state. Exit is nonzero on lock, configuration, state, or version-policy failure. Primary consumer: Directory.Build.targets/MSBuild; human invocation is diagnostic only.

.EXAMPLE
.\Tools\ResolveVersionForMsbuild.ps1 -RepoRoot . -Configuration Debug -Platform x64
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot,

    [Parameter(Mandatory = $true)]
    [string]$Configuration,

    [Parameter(Mandatory = $true)]
    [string]$Platform,

    [switch]$OfficialRelease
)

$ErrorActionPreference = "Stop"

$resolvedRepoRoot = (Resolve-Path $RepoRoot).Path
Import-Module (Join-Path $resolvedRepoRoot "Tools\Modules\Build\Versioning.psm1") -Force -ErrorAction Stop

$versionContext = Resolve-RSVersionContext `
    -RepoRoot $resolvedRepoRoot `
    -Configuration $Configuration `
    -Platform $Platform `
    -OfficialRelease:$OfficialRelease
[Console]::Out.Write($versionContext.BuildNumber)
