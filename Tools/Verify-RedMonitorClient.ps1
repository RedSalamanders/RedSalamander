<#
.SYNOPSIS
Verifies the exact local RedMonitor client and tooling pins without network or build work.
.DESCRIPTION
Uses the reviewed verifier byte for byte after checking both tool hashes.
.PARAMETER RepositoryRoot
Consumer repository containing Dependencies/RedMonitor.lock.json and ThirdParty/RedMonitor.
.OUTPUTS
Console verification result.
.NOTES
Read-only. Never builds or launches the viewer.
.EXAMPLE
./Tools/Verify-RedMonitorClient.ps1
#>
[CmdletBinding()]
param([string]$RepositoryRoot)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $RepositoryRoot) { $RepositoryRoot = Split-Path -Parent $PSScriptRoot }
$root = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($RepositoryRoot)
Import-Module (Join-Path $PSScriptRoot 'Modules/Build/RedMonitorClientTools.psm1') -Scope Local
$verify = Get-RedMonitorClientCommand -Root $root -Name 'Verify-RedMonitorClient.ps1'
$PSBoundParameters['RepositoryRoot'] = $root
& $verify @PSBoundParameters
