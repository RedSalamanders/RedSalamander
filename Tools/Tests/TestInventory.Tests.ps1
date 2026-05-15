Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Tools\TestInventory.ps1'

function Assert-RSEqual {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Actual,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected' but got '$Actual'."
    }
}

Describe 'Test inventory helper' {
    BeforeAll {
        . $helperScript
    }

    It 'counts the in-product self-test surfaces from source' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot

        Assert-RSEqual -Actual $inventory.SelfTests.Commands.RunCaseRegistrations -Expected 598 -Message 'Commands static RunCase count drifted.'
        Assert-RSEqual -Actual $inventory.SelfTests.CompareDirectories.RunCaseRegistrations -Expected 141 -Message 'CompareDirectories static RunCase count drifted.'
        Assert-RSEqual -Actual $inventory.SelfTests.FileOperations.ActivePhases -Expected 73 -Message 'FileOperations active phase count drifted.'
    }

    It 'counts standalone, performance, and script test surfaces from source' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot

        Assert-RSEqual -Actual $inventory.Standalone.PerformanceTests2.TestMethods -Expected 11 -Message 'PerformanceTests2 method count drifted.'
        Assert-RSEqual -Actual $inventory.Scripts.ToolsPester.Cases -Expected 74 -Message 'Tools Pester test count drifted.'
        Assert-RSEqual -Actual $inventory.Scripts.ToolsPester.RequiresBuildToolchainCases -Expected 1 -Message 'Build-toolchain Pester count drifted.'
        Assert-RSEqual -Actual $inventory.Scripts.VcpkgMergeSynthetic.Cases -Expected 5 -Message 'Synthetic vcpkg merge count drifted.'
        Assert-RSEqual -Actual $inventory.Scripts.VcpkgMergeLockValidation.Cases -Expected 3 -Message 'Lock-validation vcpkg count drifted.'
    }

    It 'exports JSON that can be used as a manifest artifact' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot
        $json = ConvertTo-RSTestInventoryJson -Inventory $inventory
        $roundTrip = $json | ConvertFrom-Json

        Assert-RSEqual -Actual $roundTrip.selfTests.commands.runCaseRegistrations -Expected 598 -Message 'JSON manifest should include Commands count.'
        Assert-RSEqual -Actual $roundTrip.scripts.toolsPester.cases -Expected 74 -Message 'JSON manifest should include Tools Pester count.'
    }

    It 'guards FileOperations Step enum values against phase-order drift' {
        $coordinator = Join-Path $repoRoot 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'
        $integrity = Get-RSFileOpsPhaseIntegrity -FilePath $coordinator

        Assert-RSEqual -Actual $integrity.ActiveEnumValues.Count -Expected 73 -Message 'FileOperations active Step count drifted.'
        Assert-RSEqual -Actual $integrity.MissingActivePhases.Count -Expected 0 -Message 'Every active Step enum value should appear in kFileOpsPhaseOrder.'
        Assert-RSEqual -Actual $integrity.DuplicateOrderedPhases.Count -Expected 0 -Message 'kFileOpsPhaseOrder should not list an active phase more than once.'
        Assert-RSEqual -Actual $integrity.ExtraOrderedPhases.Count -Expected 0 -Message 'kFileOpsPhaseOrder should not reference unknown Step values.'
    }

    It 'keeps documented inventory counts aligned with the source manifest' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot
        $coverageDoc = Get-RSTestInventoryDocSnapshot -Path (Join-Path $repoRoot 'Specs\Testing\Testing_TestCoverage.md')
        $readmeDoc = Get-RSTestInventoryDocSnapshot -Path (Join-Path $repoRoot 'Tests\README.md')

        Assert-RSEqual -Actual $coverageDoc.CommandsRunCases -Expected $inventory.SelfTests.Commands.RunCaseRegistrations -Message 'Coverage spec Commands count drifted from source.'
        Assert-RSEqual -Actual $coverageDoc.CompareRunCases -Expected $inventory.SelfTests.CompareDirectories.RunCaseRegistrations -Message 'Coverage spec CompareDirectories count drifted from source.'
        Assert-RSEqual -Actual $coverageDoc.FileOpsActivePhases -Expected $inventory.SelfTests.FileOperations.ActivePhases -Message 'Coverage spec FileOperations count drifted from source.'
        Assert-RSEqual -Actual $coverageDoc.PerformanceTestMethods -Expected $inventory.Standalone.PerformanceTests2.TestMethods -Message 'Coverage spec PerformanceTests2 count drifted from source.'
        Assert-RSEqual -Actual $coverageDoc.ToolsPesterCases -Expected $inventory.Scripts.ToolsPester.Cases -Message 'Coverage spec Tools Pester count drifted from source.'
        Assert-RSEqual -Actual $readmeDoc.ToolsPesterCases -Expected $inventory.Scripts.ToolsPester.Cases -Message 'Tests README Tools Pester count drifted from source.'
        Assert-RSEqual -Actual $readmeDoc.VcpkgSyntheticCases -Expected $inventory.Scripts.VcpkgMergeSynthetic.Cases -Message 'Tests README vcpkg synthetic count drifted from source.'
    }
}
