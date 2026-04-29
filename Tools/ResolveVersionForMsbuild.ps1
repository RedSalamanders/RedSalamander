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
. (Join-Path $resolvedRepoRoot "Tools\Versioning.ps1")

Use-RSVersionStateLock -RepoRoot $resolvedRepoRoot -ScriptBlock {
    $script:versionContext = Get-RSVersionContext -RepoRoot $resolvedRepoRoot -Configuration $Configuration -Platform $Platform -OfficialRelease:$OfficialRelease
    Save-RSVersionContext -RepoRoot $resolvedRepoRoot -VersionContext $script:versionContext | Out-Null
} | Out-Null
[Console]::Out.Write($versionContext.BuildNumber)
