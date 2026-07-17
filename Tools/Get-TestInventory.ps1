<#
.SYNOPSIS
    Emits the source-derived RedSalamander test inventory.

.DESCRIPTION
    Scans stable source registration patterns and reconciles native test projects
    with the canonical CI/Full run plan. The JSON manifest records test surfaces,
    execution kinds, and derived counts without making documentation own mutable
    totals. Runner-native case listing remains authoritative for in-product cases.
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
Write-Host "TestHarness source-contract classifications: $($inventory.Scripts.ToolsPester.SourceContracts.CategoryCounts | ConvertTo-Json -Compress)"
Write-Host "TestHarness behavioral replacement candidates: $(@($inventory.Scripts.ToolsPester.SourceContracts.ReplacementCandidates).Count)"
Write-Host "vcpkg synthetic merge cases: $($inventory.Scripts.VcpkgMergeSynthetic.Cases)"
Write-Host "vcpkg lock validation cases: $($inventory.Scripts.VcpkgMergeLockValidation.Cases)"
Write-Host "Project-backed test surfaces: $(@($inventory.RunPlan.ProjectBackedSurfaces).Count)"
foreach ($surface in $inventory.RunPlan.ProjectBackedSurfaces) {
    Write-Host "  $($surface.Name): $($surface.Kind)"
}
Write-Host "CI run-plan entries: $(@($inventory.RunPlan.CI).Count)"
Write-Host "Full run-plan entries: $(@($inventory.RunPlan.Full).Count)"
