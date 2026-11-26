<#
.SYNOPSIS
Validates repository guidance and inventories.
.DESCRIPTION
Runs the guidance, tool inventory, and specification validators used by
compiler-free pull request CI.
.OUTPUTS
Guidance status and inventory results. Validation failures are terminating errors.
.NOTES
Requires PowerShell 7 and the exact clean .ci-engine checkout. No reports are written.
.EXAMPLE
./validate.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

& (Join-Path $PSScriptRoot 'validate-skills.ps1')
& (Join-Path $PSScriptRoot 'Tools/Get-ToolInventory.ps1') -FailOnFindings -Format Json
& (Join-Path $PSScriptRoot 'validate-specs.ps1')
