<#
.SYNOPSIS
Explicitly updates the RedMonitor client header, notice, and immutable runtime pin.
.DESCRIPTION
Invokes the reviewed updater at its separate tooling pin. Only the two client
files are downloaded from the selected commit. The tooling pin is preserved.
.PARAMETER Commit
Full immutable RedSalamanders/RedMonitor commit reviewed for adoption.
.PARAMETER RepositoryRoot
Consumer repository to update.
.OUTPUTS
Console update result and updated client files and lock.
.NOTES
Requires PowerShell 7 and authenticated gh access. Review the resulting files and lock together.
.EXAMPLE
./Tools/Update-RedMonitorClient.ps1 -Commit 0000000000000000000000000000000000000000
#>
#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{40}$')][string]$Commit,
    [string]$RepositoryRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $RepositoryRoot) { $RepositoryRoot = Split-Path -Parent $PSScriptRoot }
$root = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($RepositoryRoot)
Import-Module (Join-Path $PSScriptRoot 'Modules/Build/RedMonitorClientTools.psm1') -Scope Local
$update = Get-RedMonitorClientCommand -Root $root -Name 'Update-RedMonitorClient.ps1'
$PSBoundParameters['RepositoryRoot'] = $root
& $update @PSBoundParameters
