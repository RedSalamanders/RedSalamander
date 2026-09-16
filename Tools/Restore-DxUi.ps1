<#
.SYNOPSIS
Restore RedSalamander's exact DxUi source and isolated build dependencies.
.DESCRIPTION
Clones only the locked commit into repository .build, restores WIL and writes platform-specific
MSBuild properties using the selected first-party compiler host, including explicit overrides.
Uses the shared-dependency lease and contained child processes. Does not edit
a sibling checkout or the lock. Unavailable pinned source or a dirty source fails closed; public HTTPS requires no secret.
.PARAMETER Platform
Target architecture, x64 or ARM64.
.PARAMETER Configuration
Debug, Release or ASan Debug; used for toolchain evaluation and operation coordination.
.PARAMETER CheckUpdates
Show one bounded advisory for a newer main revision without changing the pin. A validated candidate includes the
next product-regression command.
.OUTPUTS
Source, logs, resolved properties, identity and isolated dependency outputs under .build.
.NOTES
Requires Git, PowerShell 7 and Visual Studio C++ tools. Creates only repository-managed dependency
source, properties and outputs under .build; network access is required on a cold restore.
.EXAMPLE
.\Tools\Restore-DxUi.ps1 -Platform x64 -Configuration Debug -CheckUpdates
#>
[CmdletBinding()]
param(
    [ValidateSet('x64','ARM64')][string]$Platform='x64',
    [ValidateSet('Debug','Release','ASan Debug')][string]$Configuration='Debug',
    [switch]$CheckUpdates
)
$ErrorActionPreference='Stop'
$repo = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $PSScriptRoot 'Modules/Build/DxUiDependency.psm1') -Force
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$installation = & $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath
if (-not $installation) { throw 'Visual Studio C++ MSBuild is required.' }
$msbuildRelative = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64') {
    'MSBuild/Current/Bin/MSBuild.exe'
} else { 'MSBuild/Current/Bin/amd64/MSBuild.exe' }
Restore-RSDxUiDependency -RepoRoot $repo -MSBuildPath (Join-Path $installation $msbuildRelative) `
    -Platform $Platform -Configuration $Configuration -CheckUpdates:$CheckUpdates
$global:LASTEXITCODE=0
