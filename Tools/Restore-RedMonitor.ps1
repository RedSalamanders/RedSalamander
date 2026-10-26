<#
.SYNOPSIS
Compatibility entrypoint that verifies only the vendored RedMonitor client; no viewer restore/build occurs.
.PARAMETER RestoreOnly
Retained compatibility switch. Client verification is always read-only.
.DESCRIPTION
Checks or explicitly updates the minimal client dependency. Ordinary verification is local and read-only.
.OUTPUTS
Console result; the explicit updater also writes the two vendored inputs and immutable lock.
.NOTES
Never builds, stages or launches the independent RedMonitor helper. Does not terminate any process.
.EXAMPLE
./Tools/Restore-RedMonitor.ps1
#>
[CmdletBinding()]
param([switch]$RestoreOnly)
$ErrorActionPreference='Stop'
& (Join-Path $PSScriptRoot 'Verify-RedMonitorClient.ps1') -RepositoryRoot (Split-Path $PSScriptRoot)