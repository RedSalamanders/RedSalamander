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

        Assert-RSEqual -Actual $inventory.SelfTests.Commands.RunCaseRegistrations -Expected 707 -Message 'Commands static RunCase count drifted.'
        Assert-RSEqual -Actual $inventory.SelfTests.CompareDirectories.RunCaseRegistrations -Expected 249 -Message 'CompareDirectories static RunCase count drifted.'
        Assert-RSEqual -Actual $inventory.SelfTests.FileOperations.ActivePhases -Expected 126 -Message 'FileOperations active phase count drifted.'

        $settingsHeader = Get-Content -LiteralPath (Join-Path $repoRoot 'Common\SettingsStore.h') -Raw
        $settingsCases = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp') -Raw
        $fileOperationsStruct = [Regex]::Match($settingsHeader, 'struct\s+FileOperationsSettings\s*\{(?<body>[\s\S]*?)\r?\n\};')
        Assert-RSEqual -Actual $fileOperationsStruct.Success -Expected $true -Message 'FileOperationsSettings definition was not found.'
        $fieldNames = @([Regex]::Matches(
                $fileOperationsStruct.Groups['body'].Value,
                '(?m)^\s*(?![#/])[^;\r\n()=]+?\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:=[^;\r\n]*)?;\s*$') |
            ForEach-Object { $_.Groups[1].Value } |
            Sort-Object -Unique)
        $fieldCaseBlock = [Regex]::Match($settingsCases, 'const\s+std::array\s+fieldCases\s*\{(?<body>[\s\S]*?)\r?\n\s*\};\s*\r?\n\s*\r?\n\s*for\s*\(const\s+FieldCase&')
        Assert-RSEqual -Actual $fieldCaseBlock.Success -Expected $true -Message 'FileOperationsSettings field-case table was not found.'
        $fieldCaseNames = @([Regex]::Matches($fieldCaseBlock.Groups['body'].Value, 'FieldCase\{L"([A-Za-z0-9_]+)"') |
            ForEach-Object { $_.Groups[1].Value } |
            Sort-Object -Unique)
        $fieldCaseDrift = @(Compare-Object -ReferenceObject $fieldNames -DifferenceObject $fieldCaseNames)
        Assert-RSEqual -Actual $fieldCaseDrift.Count -Expected 0 -Message 'Every FileOperationsSettings field must have exactly one isolated persistence case.'
    }

    It 'counts standalone, performance, and script test surfaces from source' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot

        Assert-RSEqual -Actual $inventory.Standalone.PerformanceTests2.TestMethods -Expected 14 -Message 'PerformanceTests2 method count drifted.'
        Assert-RSEqual -Actual $inventory.Standalone.DxUiTests.NativeTextInputCases -Expected 118 -Message 'NativeTextInput method count drifted.'
        Assert-RSEqual -Actual $inventory.Scripts.ToolsPester.Cases -Expected 294 -Message 'Tools Pester test count drifted.'
        Assert-RSEqual -Actual $inventory.Scripts.ToolsPester.RequiresBuildToolchainCases -Expected 1 -Message 'Build-toolchain Pester count drifted.'
        Assert-RSEqual -Actual $inventory.Scripts.VcpkgMergeSynthetic.Cases -Expected 5 -Message 'Synthetic vcpkg merge count drifted.'
        Assert-RSEqual -Actual $inventory.Scripts.VcpkgMergeLockValidation.Cases -Expected 3 -Message 'Lock-validation vcpkg count drifted.'
    }

    It 'exports JSON that can be used as a manifest artifact' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot
        $json = ConvertTo-RSTestInventoryJson -Inventory $inventory
        $roundTrip = $json | ConvertFrom-Json

        Assert-RSEqual -Actual $roundTrip.selfTests.commands.runCaseRegistrations -Expected 707 -Message 'JSON manifest should include Commands count.'
        Assert-RSEqual -Actual $roundTrip.standalone.dxUiTests.nativeTextInputCases -Expected 118 -Message 'JSON manifest should include NativeTextInput count.'
        Assert-RSEqual -Actual $roundTrip.scripts.toolsPester.cases -Expected 294 -Message 'JSON manifest should include Tools Pester count.'
    }

    It 'guards FileOperations Step enum values against phase-order drift' {
        $coordinator = Join-Path $repoRoot 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'
        $integrity = Get-RSFileOpsPhaseIntegrity -FilePath $coordinator

        Assert-RSEqual -Actual $integrity.ActiveEnumValues.Count -Expected 126 -Message 'FileOperations active Step count drifted.'
        Assert-RSEqual -Actual $integrity.MissingActivePhases.Count -Expected 0 -Message 'Every active Step enum value should appear in kFileOpsPhaseOrder.'
        Assert-RSEqual -Actual $integrity.DuplicateOrderedPhases.Count -Expected 0 -Message 'kFileOpsPhaseOrder should not list an active phase more than once.'
        Assert-RSEqual -Actual $integrity.ExtraOrderedPhases.Count -Expected 0 -Message 'kFileOpsPhaseOrder should not reference unknown Step values.'
    }

    It 'keeps documented inventory counts aligned with the source manifest' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot
        $coverageDoc = Get-RSTestInventoryDocSnapshot -Path (Join-Path $repoRoot 'Specs\Testing\Testing_TestCoverage.md')
        $readmeDoc = Get-RSTestInventoryDocSnapshot -Path (Join-Path $repoRoot 'Tests\README.md')
        $readmeInventory = Get-RSTestsReadmeInventorySnapshot -Path (Join-Path $repoRoot 'Tests\README.md')

        Assert-RSEqual -Actual $coverageDoc.CommandsRunCases -Expected $inventory.SelfTests.Commands.RunCaseRegistrations -Message 'Coverage spec Commands count drifted from source.'
        Assert-RSEqual -Actual $coverageDoc.CompareRunCases -Expected $inventory.SelfTests.CompareDirectories.RunCaseRegistrations -Message 'Coverage spec CompareDirectories count drifted from source.'
        Assert-RSEqual -Actual $coverageDoc.FileOpsActivePhases -Expected $inventory.SelfTests.FileOperations.ActivePhases -Message 'Coverage spec FileOperations count drifted from source.'
        Assert-RSEqual -Actual $coverageDoc.PerformanceTestMethods -Expected $inventory.Standalone.PerformanceTests2.TestMethods -Message 'Coverage spec PerformanceTests2 count drifted from source.'
        Assert-RSEqual -Actual $coverageDoc.NativeTextInputCases -Expected $inventory.Standalone.DxUiTests.NativeTextInputCases -Message 'Coverage spec NativeTextInput count drifted from source.'
        Assert-RSEqual -Actual $readmeDoc.NativeTextInputCases -Expected $inventory.Standalone.DxUiTests.NativeTextInputCases -Message 'Tests README NativeTextInput count drifted from source.'
        Assert-RSEqual -Actual $coverageDoc.ToolsPesterCases -Expected $inventory.Scripts.ToolsPester.Cases -Message 'Coverage spec Tools Pester count drifted from source.'
        Assert-RSEqual -Actual $readmeDoc.ToolsPesterCases -Expected $inventory.Scripts.ToolsPester.Cases -Message 'Tests README Tools Pester count drifted from source.'
        Assert-RSEqual -Actual $readmeDoc.VcpkgSyntheticCases -Expected $inventory.Scripts.VcpkgMergeSynthetic.Cases -Message 'Tests README vcpkg synthetic count drifted from source.'

        Assert-RSEqual -Actual $readmeInventory.CommandsListedCases.Count -Expected 3 -Message 'Tests README should expose Commands listed count in overview, heading, and narrative.'
        foreach ($count in $readmeInventory.CommandsListedCases) {
            Assert-RSEqual -Actual $count -Expected 809 -Message 'Tests README Commands listed counts disagree.'
        }
        Assert-RSEqual -Actual $readmeInventory.CommandsStaticRunCases.Count -Expected 1 -Message 'Tests README should expose one Commands static-registration count.'
        Assert-RSEqual -Actual $readmeInventory.CommandsStaticRunCases[0] -Expected $inventory.SelfTests.Commands.RunCaseRegistrations -Message 'Tests README Commands static count drifted from source.'

        Assert-RSEqual -Actual $readmeInventory.CompareListedCases.Count -Expected 3 -Message 'Tests README should expose Compare listed count in overview, heading, and narrative.'
        foreach ($count in $readmeInventory.CompareListedCases) {
            Assert-RSEqual -Actual $count -Expected 256 -Message 'Tests README Compare listed counts disagree.'
        }
        Assert-RSEqual -Actual $readmeInventory.CompareStaticRunCases.Count -Expected 1 -Message 'Tests README should expose one Compare static-registration count.'
        Assert-RSEqual -Actual $readmeInventory.CompareStaticRunCases[0] -Expected $inventory.SelfTests.CompareDirectories.RunCaseRegistrations -Message 'Tests README Compare static count drifted from source.'

        Assert-RSEqual -Actual $readmeInventory.FileOpsListedCases.Count -Expected 3 -Message 'Tests README should expose FileOps listed count in overview, heading, and narrative.'
        foreach ($count in $readmeInventory.FileOpsListedCases) {
            Assert-RSEqual -Actual $count -Expected ($inventory.SelfTests.FileOperations.ActivePhases + 2) -Message 'Tests README FileOps listed counts disagree.'
        }
        Assert-RSEqual -Actual $readmeInventory.FileOpsActivePhases.Count -Expected 1 -Message 'Tests README should expose one FileOps active-phase count.'
        Assert-RSEqual -Actual $readmeInventory.FileOpsActivePhases[0] -Expected $inventory.SelfTests.FileOperations.ActivePhases -Message 'Tests README FileOps active count drifted from source.'

        Assert-RSEqual -Actual $readmeInventory.ToolsPesterCases.Count -Expected 1 -Message 'Tests README should expose one Tools Pester aggregate count.'
        Assert-RSEqual -Actual $readmeInventory.ToolsPesterCases[0] -Expected $inventory.Scripts.ToolsPester.Cases -Message 'Tests README Tools Pester aggregate drifted from source.'

        foreach ($property in $inventory.SelfTests.Commands.FamilyRunCaseRegistrations.PSObject.Properties) {
            $documented = $readmeInventory.CommandsFamilyCases.PSObject.Properties[$property.Name]
            Assert-RSEqual -Actual ($null -ne $documented) -Expected $true -Message "Tests README is missing Commands family '$($property.Name)'."
            Assert-RSEqual -Actual $documented.Value -Expected $property.Value -Message "Tests README Commands family '$($property.Name)' drifted from source."
        }
        $commandsFamilyTotal = ($inventory.SelfTests.Commands.FamilyRunCaseRegistrations.PSObject.Properties.Value | Measure-Object -Sum).Sum
        Assert-RSEqual -Actual $inventory.SelfTests.Commands.CoordinatorRunCaseRegistrations -Expected 1 -Message 'Commands coordinator should retain one isolation-failure RunCase registration.'
        Assert-RSEqual -Actual ($commandsFamilyTotal + $inventory.SelfTests.Commands.CoordinatorRunCaseRegistrations) -Expected $inventory.SelfTests.Commands.RunCaseRegistrations -Message 'Commands family and coordinator registrations should account for the aggregate source count.'
        Assert-RSEqual -Actual @($readmeInventory.CommandsFamilyCases.PSObject.Properties).Count -Expected @($inventory.SelfTests.Commands.FamilyRunCaseRegistrations.PSObject.Properties).Count -Message 'Tests README Commands family table has extra or missing rows.'

        foreach ($property in $inventory.Scripts.ToolsPester.FileCases.PSObject.Properties) {
            $documented = $readmeInventory.ToolsPesterFileCases.PSObject.Properties[$property.Name]
            Assert-RSEqual -Actual ($null -ne $documented) -Expected $true -Message "Tests README is missing Tools Pester file '$($property.Name)'."
            Assert-RSEqual -Actual $documented.Value -Expected $property.Value -Message "Tests README Tools Pester file '$($property.Name)' drifted from source."
        }
        Assert-RSEqual -Actual @($readmeInventory.ToolsPesterFileCases.PSObject.Properties).Count -Expected @($inventory.Scripts.ToolsPester.FileCases.PSObject.Properties).Count -Message 'Tests README Tools Pester table has extra or missing rows.'
    }
}
