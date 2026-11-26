<#
.SYNOPSIS
Verifies the pinned local RedMonitor client and tool bundle.
.PARAMETER RepositoryRoot
Consumer checkout. Defaults to this repository.
.OUTPUTS
Console verification result.
.NOTES
Read-only. The existing Tools entry point performs verification.
.EXAMPLE
./Verify-RedMonitorClient.ps1
#>
[CmdletBinding()]
param([string]$RepositoryRoot)

$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'Tools/Verify-RedMonitorClient.ps1') @PSBoundParameters
