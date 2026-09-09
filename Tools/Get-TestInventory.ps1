<#
.SYNOPSIS
    Emits the source-derived RedSalamander test inventory.

.DESCRIPTION
    Scans stable source registration patterns and reconciles native test projects
    with the canonical CI/Full run plan. The JSON manifest records test surfaces,
    execution kinds, and derived counts without making documentation own mutable
    totals. Runner-native case listing remains authoritative for in-product cases.

.PARAMETER RepoRoot
    Specifies the repository root whose source registrations and run plan are inventoried.

.PARAMETER Format
    Emits machine-readable JSON or a concise human-readable text summary.

.OUTPUTS
    A JSON string when Format is Json. Text mode writes a summary to the host and emits no supported pipeline objects.

.NOTES
    Prerequisites: a repository checkout with the canonical self-test registration and Tools Pester structure. Side effects: none; the command is read-only. Exit is nonzero on malformed registrations, missing run-plan surfaces, or inventory reconciliation failure. Primary consumers: Get-SpecInventory.ps1, test-coverage review, and inventory policy tests.

.EXAMPLE
    .\Tools\Get-TestInventory.ps1 -Format Json

.EXAMPLE
    .\Tools\Get-TestInventory.ps1 -Format Text
#>

[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),

    [ValidateSet('Json', 'Text')]
    [string]$Format = 'Json'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$helper = Join-Path $PSScriptRoot 'Modules\Testing\TestInventory.psm1'
Import-Module $helper -Force -ErrorAction Stop

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
