<#
.SYNOPSIS
    Emits the source-derived RedSalamander test inventory.

.DESCRIPTION
    Scans test source files for stable registration patterns and emits a JSON
    manifest. This is a lightweight bridge until each runner can emit its own
    authoritative case list.
#>

[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),

    [ValidateSet('Json', 'Text')]
    [string]$Format = 'Json'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$helper = Join-Path $PSScriptRoot 'TestInventory.ps1'
. $helper

$inventory = Get-RSTestInventory -RepoRoot $RepoRoot

if ($Format -eq 'Json') {
    ConvertTo-RSTestInventoryJson -Inventory $inventory
    exit 0
}

Write-Host "Commands RunCase registrations: $($inventory.SelfTests.Commands.RunCaseRegistrations)"
Write-Host "CompareDirectories RunCase registrations: $($inventory.SelfTests.CompareDirectories.RunCaseRegistrations)"
Write-Host "FileOperations active phases: $($inventory.SelfTests.FileOperations.ActivePhases)"
Write-Host "PerformanceTests2 TEST_METHOD count: $($inventory.Standalone.PerformanceTests2.TestMethods)"
Write-Host "Tools Pester cases: $($inventory.Scripts.ToolsPester.Cases)"
Write-Host "Tools Pester RequiresBuildToolchain cases: $($inventory.Scripts.ToolsPester.RequiresBuildToolchainCases)"
Write-Host "vcpkg synthetic merge cases: $($inventory.Scripts.VcpkgMergeSynthetic.Cases)"
Write-Host "vcpkg lock validation cases: $($inventory.Scripts.VcpkgMergeLockValidation.Cases)"
