<#
.SYNOPSIS
Restores the pinned RedSalamander, DxUi and RedMonitor client dependencies.
.PARAMETER Platform
Target x64 or ARM64 architecture. Defaults to x64.
.PARAMETER Configuration
Selects the DxUi profile and vcpkg triplet. Defaults to Debug.
.OUTPUTS
Dependency restore output from the existing product commands.
.NOTES
Requires PowerShell 7, Git and the supported Visual Studio toolchain. Managed
dependencies and logs go under .build. A cold restore needs network access.
.EXAMPLE
./restore.ps1 -Platform ARM64 -Configuration Release
#>
#Requires -Version 7.0
[CmdletBinding()]
param(
    [ValidateSet('x64', 'ARM64')]
    [string]$Platform = 'x64',

    [ValidateSet('Debug', 'Release', 'ASan Debug')]
    [string]$Configuration = 'Debug'
)

$ErrorActionPreference = 'Stop'
if ($env:CI_NATIVE_BUILDS_DISABLED -eq 'true') {
    throw 'Dependency restoration is forbidden in compiler-free PR checks.'
}

& (Join-Path $PSScriptRoot 'Verify-RedMonitorClient.ps1')
if (-not $?) { exit $LASTEXITCODE }

& (Join-Path $PSScriptRoot 'Tools/Restore-DxUi.ps1') -Platform $Platform -Configuration $Configuration
if (-not $?) { exit $LASTEXITCODE }

Import-Module (Join-Path $PSScriptRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -ErrorAction Stop
$tool = Initialize-RSVcpkgToolContext -RepoRoot $PSScriptRoot
& (Join-Path $PSScriptRoot 'vcpkg-install.ps1') -Platform $Platform -Asan:($Configuration -eq 'ASan Debug') -VcpkgExe $tool.Executable -VcpkgRoot $tool.RegistryRoot -ReuseCurrent
if (-not $?) { exit $LASTEXITCODE }
