Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestInventory.psm1'
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force

Describe 'Test inventory helper' {
    BeforeAll {
        Import-Module $helperModule -Force -ErrorAction Stop
    }

    It 'derives in-product self-test surfaces without frozen registration totals' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot

        Assert-RSEqual -Actual ($inventory.SelfTests.Commands.RunCaseRegistrations -gt 0) -Expected $true -Message 'Commands source inventory should not be empty.'
        Assert-RSEqual -Actual ($inventory.SelfTests.CompareDirectories.RunCaseRegistrations -gt 0) -Expected $true -Message 'CompareDirectories source inventory should not be empty.'
        Assert-RSEqual -Actual ($inventory.SelfTests.FileOperations.ActivePhases -gt 0) -Expected $true -Message 'FileOperations source inventory should not be empty.'

        $commandsFamilyTotal = ($inventory.SelfTests.Commands.FamilyRunCaseRegistrations.PSObject.Properties.Value | Measure-Object -Sum).Sum
        Assert-RSEqual -Actual $inventory.SelfTests.Commands.CoordinatorRunCaseRegistrations -Expected 1 -Message 'Commands coordinator should retain one isolation-failure RunCase registration.'
        Assert-RSEqual -Actual ($commandsFamilyTotal + $inventory.SelfTests.Commands.CoordinatorRunCaseRegistrations) -Expected $inventory.SelfTests.Commands.RunCaseRegistrations -Message 'Commands family and coordinator registrations should account for the aggregate source inventory.'

        $settingsHeader = Get-Content -LiteralPath (Join-Path $repoRoot 'Common\SettingsStore.h') -Raw
        $settingsCases = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp') -Raw
        $fileOperationsStruct = [Regex]::Match($settingsHeader, 'struct\s+FileOperationsSettings\s*\{(?<body>[\s\S]*?)\r?\n\};')
        Assert-RSEqual -Actual $fileOperationsStruct.Success -Expected $true -Message 'FileOperationsSettings definition was not found.'
        $fieldNames = @([Regex]::Matches(
                $fileOperationsStruct.Groups['body'].Value,
                '(?m)^\s*(?![#/])[^;\r\n()=]+?\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:=[^;\r\n]*)?;\s*$') |
            ForEach-Object { $_.Groups[1].Value } |
            Sort-Object -Unique)
        $nonPersistedPhase0Names = @([Regex]::Matches(
                $fileOperationsStruct.Groups['body'].Value,
                '(?m)^\s*(?![#/])[^;\r\n()=]+?\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:=[^;\r\n]*)?;\s*//\s*NON_PERSISTED_PHASE0\s*$') |
            ForEach-Object { $_.Groups[1].Value } |
            Sort-Object -Unique)
        Assert-RSEqual -Actual $nonPersistedPhase0Names.Count -Expected 0 -Message 'No FileOperationsSettings field may bypass persistence coverage after the Phase 0 bridge retirement.'
        $fieldCaseBlock = [Regex]::Match($settingsCases, 'const\s+std::array\s+fieldCases\s*\{(?<body>[\s\S]*?)\r?\n\s*\};\s*\r?\n\s*\r?\n\s*for\s*\(const\s+FieldCase&')
        Assert-RSEqual -Actual $fieldCaseBlock.Success -Expected $true -Message 'FileOperationsSettings field-case table was not found.'
        $fieldCaseNames = @([Regex]::Matches($fieldCaseBlock.Groups['body'].Value, 'FieldCase\{L"([A-Za-z0-9_]+)"') |
            ForEach-Object { $_.Groups[1].Value } |
            Sort-Object -Unique)
        $fieldCaseDrift = @(Compare-Object -ReferenceObject $fieldNames -DifferenceObject $fieldCaseNames)
        Assert-RSEqual -Actual $fieldCaseDrift.Count -Expected 0 -Message 'Every FileOperationsSettings field must have exactly one isolated persistence case.'
    }

    It 'derives every native test project and its kind from the canonical run plan' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot
        $projectNames = @(Get-RSTestProjectNames -RepoRoot $repoRoot)
        $inventoryProjectNames = @($inventory.RunPlan.ProjectBackedSurfaces.Name)
        $projectDrift = @(Compare-Object -ReferenceObject $projectNames -DifferenceObject $inventoryProjectNames)
        Assert-RSEqual -Actual $projectDrift.Count -Expected 0 -Message 'Tests project files and run-plan-backed inventory surfaces should have set equality.'

        foreach ($surface in $inventory.RunPlan.ProjectBackedSurfaces) {
            $expectedKind = if ($surface.Name -eq 'PerformanceTests2') { 'CppUnitTest' } else { 'Executable' }
            Assert-RSEqual -Actual $surface.Kind -Expected $expectedKind -Message "Run-plan kind drifted for test project '$($surface.Name)'."
            Assert-RSEqual -Actual ($surface.PlanEntries.Count -gt 0) -Expected $true -Message "Test project '$($surface.Name)' should have at least one run-plan entry."
        }
    }

    It 'keeps critical CI and Full surfaces visible with their run-plan kinds' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot
        $allEntries = @($inventory.RunPlan.CI + $inventory.RunPlan.Full)
        $requiredKinds = [ordered]@{
            PluginContractTests = 'Executable'
            SettingsSchemaTests = 'Executable'
            CrashHandlingTests = 'Executable'
            RedSalamanderMonitorEtwLatency = 'Executable'
            PerformanceTests2 = 'CppUnitTest'
            ToolsPesterTests = 'Pester'
            ToolsPesterBuildToolchain = 'Pester'
            VcpkgMergeSynthetic = 'PowerShellScript'
        }

        foreach ($required in $requiredKinds.GetEnumerator()) {
            $matches = @($allEntries | Where-Object { $_.Name -eq $required.Key -and $_.Kind -eq $required.Value })
            Assert-RSEqual -Actual ($matches.Count -gt 0) -Expected $true -Message "Required test surface '$($required.Key)' with kind '$($required.Value)' is missing from CI/Full inventory."
        }
        foreach ($entry in $allEntries) {
            Assert-RSEqual -Actual $entry.MetadataComplete -Expected $true -Message "Run-plan metadata is incomplete for '$($entry.Name)'."
            Assert-RSEqual -Actual ([string]::IsNullOrWhiteSpace($entry.Id)) -Expected $false -Message "Run-plan ID is missing for '$($entry.Name)'."
        }
    }

    It 'exports JSON as a lossless manifest of source and run-plan inventory' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot
        $roundTrip = ConvertTo-RSTestInventoryJson -Inventory $inventory | ConvertFrom-Json

        Assert-RSEqual -Actual $roundTrip.selfTests.commands.runCaseRegistrations -Expected $inventory.SelfTests.Commands.RunCaseRegistrations -Message 'JSON manifest should preserve the derived Commands registration count.'
        Assert-RSEqual -Actual $roundTrip.standalone.productUiTests.cases -Expected $inventory.Standalone.ProductUiTests.Cases -Message 'JSON manifest should preserve the derived product UI count.'
        Assert-RSEqual -Actual $roundTrip.scripts.toolsPester.cases -Expected $inventory.Scripts.ToolsPester.Cases -Message 'JSON manifest should preserve the derived Tools Pester count.'
        Assert-RSEqual -Actual @($roundTrip.scripts.toolsPester.sourceContracts.cases).Count -Expected @($inventory.Scripts.ToolsPester.SourceContracts.Cases).Count -Message 'JSON manifest should preserve every classified source-contract case.'
        Assert-RSEqual -Actual @($roundTrip.scripts.toolsPester.sourceContracts.replacementCandidates).Count -Expected @($inventory.Scripts.ToolsPester.SourceContracts.ReplacementCandidates).Count -Message 'JSON manifest should preserve the behavioral replacement queue.'
        Assert-RSEqual -Actual @($roundTrip.runPlan.projectBackedSurfaces).Count -Expected @($inventory.RunPlan.ProjectBackedSurfaces).Count -Message 'JSON manifest should preserve every project-backed test surface.'
        Assert-RSEqual -Actual @($roundTrip.runPlan.ci).Count -Expected @($inventory.RunPlan.CI).Count -Message 'JSON manifest should preserve every CI run-plan entry.'
        Assert-RSEqual -Actual @($roundTrip.runPlan.full).Count -Expected @($inventory.RunPlan.Full).Count -Message 'JSON manifest should preserve every Full run-plan entry.'
        Assert-RSEqual -Actual @($roundTrip.runPlan.full | Where-Object { -not $_.metadataComplete }).Count -Expected 0 -Message 'JSON manifest should report complete metadata for every Full entry.'
    }

    It 'classifies every source-contract case for replacement decisions' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot
        $sourceContracts = $inventory.Scripts.ToolsPester.SourceContracts
        $sourceContractPath = Join-Path $repoRoot 'Tools\Tests\TestHarnessSourceContracts.Tests.ps1'
        $declaredCount = Get-RSSelectStringCount -Path @($sourceContractPath) -Pattern '^\s*It\s'

        Assert-RSEqual -Actual @($sourceContracts.Cases).Count -Expected $declaredCount -Message 'Every source-contract It block should have one live classification.'
        Assert-RSEqual -Actual @($sourceContracts.Cases.Name | Sort-Object -Unique).Count -Expected $declaredCount -Message 'Source-contract classification names should be unique.'
        Assert-RSEqual -Actual (($sourceContracts.CategoryCounts.PSObject.Properties.Value | Measure-Object -Sum).Sum) -Expected $declaredCount -Message 'Source-contract category totals should account for every case.'
        Assert-RSEqual -Actual (@($sourceContracts.ReplacementCandidates).Count -gt 0) -Expected $true -Message 'The inventory should expose behavioral source-shape checks that still need runtime replacement.'
    }

    It 'guards FileOperations Step enum values against phase-order drift' {
        $inventory = Get-RSTestInventory -RepoRoot $repoRoot
        $coordinator = Join-Path $repoRoot 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'
        $integrity = Get-RSFileOpsPhaseIntegrity -FilePath $coordinator

        Assert-RSEqual -Actual $integrity.ActiveEnumValues.Count -Expected $inventory.SelfTests.FileOperations.ActivePhases -Message 'FileOperations active Step inventory should be derived from the phase order.'
        Assert-RSEqual -Actual $integrity.MissingActivePhases.Count -Expected 0 -Message 'Every active Step enum value should appear in kFileOpsPhaseOrder.'
        Assert-RSEqual -Actual $integrity.DuplicateOrderedPhases.Count -Expected 0 -Message 'kFileOpsPhaseOrder should not list an active phase more than once.'
        Assert-RSEqual -Actual $integrity.ExtraOrderedPhases.Count -Expected 0 -Message 'kFileOpsPhaseOrder should not reference unknown Step values.'
        Assert-RSEqual -Actual ($integrity.OrderedPhases -contains 'FileOps_VisualGalleryInteraction') -Expected $false -Message 'Interactive documentation capture remains exact-case only.'
        Assert-RSEqual -Actual ($integrity.ActiveEnumValues -contains 'FileOps_VisualGalleryInteraction') -Expected $false -Message 'Interactive capture must not inflate broad behavioral coverage.'
        Assert-RSEqual -Actual ($integrity.EnumValues -contains 'FileOps_VisualGallery') -Expected $true -Message 'The explicit documentation gallery must remain a discoverable Step.'
        Assert-RSEqual -Actual ($integrity.OrderedPhases -contains 'FileOps_VisualGallery') -Expected $false -Message 'Documentation capture must not enter the broad native phase sequence.'
        Assert-RSEqual -Actual ($integrity.ActiveEnumValues -contains 'FileOps_VisualGallery') -Expected $false -Message 'Opt-in documentation capture must not inflate behavioral coverage inventory.'
    }

    It 'documents commands for live inventory instead of checked-in current totals' {
        $readme = Get-Content -LiteralPath (Join-Path $repoRoot 'Tests\README.md') -Raw
        $coverage = Get-Content -LiteralPath (Join-Path $repoRoot 'Specs\Testing\Testing_TestCoverage.md') -Raw

        foreach ($document in @($readme, $coverage)) {
            $document | Should Match 'Tools[\\/]Get-TestInventory\.ps1\s+-Format\s+Json'
            $document | Should Not Match 'Counts below are current as of'
            $document | Should Not Match 'Current runner-native inventory as of'
            $document | Should Not Match 'Current source-derived fallback counts'
        }

        $readme | Should Not Match 'runner-listed\s+(cases|phases)'
        $readme | Should Not Match 'source fallback scan reports\s+\d+'
    }
}
