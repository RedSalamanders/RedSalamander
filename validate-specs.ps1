<#
.SYNOPSIS
Validates RedSalamander's specification and plan inventories without writing reports.
.OUTPUTS
Inventory results. Validation failures are terminating errors.
.NOTES
Requires PowerShell 7, Git, and the product source and specification files.
.EXAMPLE
./validate-specs.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

& (Join-Path $PSScriptRoot 'Tools/Get-SpecInventory.ps1') -FailOnFindings -NoReports
