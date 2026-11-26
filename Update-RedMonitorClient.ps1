<#
.SYNOPSIS
Updates the RedMonitor client header and notice at one reviewed commit.
.PARAMETER Commit
Full immutable RedMonitor commit reviewed for client adoption.
.PARAMETER RepositoryRoot
Consumer checkout. Defaults to this repository.
.OUTPUTS
Console update result, client header and notice, and the updated client lock.
.NOTES
Requires PowerShell 7 and authenticated gh access. The existing Tools entry point runs the update.
.EXAMPLE
./Update-RedMonitorClient.ps1 -Commit 0000000000000000000000000000000000000000
#>
#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{40}$')][string]$Commit,
    [string]$RepositoryRoot
)

$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'Tools/Update-RedMonitorClient.ps1') @PSBoundParameters
