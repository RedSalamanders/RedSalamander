Set-StrictMode -Version Latest

function Get-RSSelectStringCount {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    $matches = @(Select-String -Path $Path -Pattern $Pattern -ErrorAction Stop)
    return $matches.Count
}

function Get-RSFileOpsActivePhaseCount {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $phaseNames = Get-RSFileOpsPhaseOrderNames -FilePath $FilePath
    return @($phaseNames | Where-Object { $_ -ne 'Setup' -and $_ -ne 'Cleanup_RestorePluginConfig' }).Count
}

function Get-RSFileOpsStepEnumNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $lines = Get-Content -LiteralPath $FilePath
    $startMatch = $lines | Select-String -Pattern '^\s*enum class Step\b' | Select-Object -First 1
    if (-not $startMatch) {
        throw "SelfTestState::Step enum was not found in $FilePath."
    }

    $start = $startMatch.LineNumber
    $end = $start
    while ($end -le $lines.Count -and $lines[$end - 1] -notmatch '^\s*\};') {
        $end++
    }

    $stepNames = @()
    for ($lineIndex = $start; $lineIndex -le $end; $lineIndex++) {
        if ($lines[$lineIndex - 1] -match '^\s*([A-Za-z0-9_]+),') {
            $stepNames += $matches[1]
        }
    }

    return $stepNames
}

function Get-RSFileOpsPhaseOrderNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $lines = Get-Content -LiteralPath $FilePath
    $startMatch = $lines | Select-String -Pattern '^constexpr auto kFileOpsPhaseOrder' | Select-Object -First 1
    if (-not $startMatch) {
        throw "kFileOpsPhaseOrder was not found in $FilePath."
    }

    $start = $startMatch.LineNumber
    $end = $start
    while ($end -le $lines.Count -and $lines[$end - 1] -notmatch '^\s*\}\);') {
        $end++
    }

    $phaseNames = @()
    for ($lineIndex = $start; $lineIndex -le $end; $lineIndex++) {
        if ($lines[$lineIndex - 1] -match 'SelfTestState::Step::([A-Za-z0-9_]+)') {
            $phaseNames += $matches[1]
        }
    }

    return $phaseNames
}

function Get-RSFileOpsPhaseIntegrity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $excluded = @('Idle', 'Setup', 'Cleanup_RestorePluginConfig', 'Done', 'Failed')
    $enumValues = @(Get-RSFileOpsStepEnumNames -FilePath $FilePath)
    $orderedPhases = @(Get-RSFileOpsPhaseOrderNames -FilePath $FilePath)
    $activeEnumValues = @($enumValues | Where-Object { $_ -notin $excluded })
    $orderedActivePhases = @($orderedPhases | Where-Object { $_ -notin $excluded })

    return [pscustomobject]@{
        EnumValues = $enumValues
        OrderedPhases = $orderedPhases
        ActiveEnumValues = $activeEnumValues
        MissingActivePhases = @($activeEnumValues | Where-Object { $_ -notin $orderedActivePhases })
        DuplicateOrderedPhases = @($orderedActivePhases | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
        ExtraOrderedPhases = @($orderedActivePhases | Where-Object { $_ -notin $enumValues })
    }
}

function Get-RSPesterTagCaseCount {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$Files,

        [Parameter(Mandatory = $true)]
        [string]$Tag
    )

    $count = 0
    foreach ($file in $Files) {
        $hasTag = Select-String -Path $file.FullName -Pattern ("-Tag\s+{0}\b" -f [Regex]::Escape($Tag)) -Quiet
        if ($hasTag) {
            $count += Get-RSSelectStringCount -Path @($file.FullName) -Pattern '^\s*It\s'
        }
    }

    return $count
}

function Get-RSTestProjectNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $testsRoot = Join-Path $RepoRoot 'Tests'
    return @(Get-ChildItem -LiteralPath $testsRoot -Recurse -Filter '*.vcxproj' |
        Where-Object { $_.FullName -notmatch '[\\/]Lang[\\/]' } |
        ForEach-Object { $_.BaseName } |
        Sort-Object -Unique)
}

function Get-RSSourceContractInventory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($FilePath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -ne 0) {
        $messages = @($parseErrors | ForEach-Object { $_.Message }) -join '; '
        throw "Source-contract inventory could not parse '$FilePath': $messages"
    }

    $allowedCategories = @('LexicalSafetyBan', 'StructuralGraphInvariant', 'BehavioralShadow', 'MixedSourceShape')
    $caseCommands = @($ast.FindAll({
                param($node)
                return $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'It'
            }, $true))

    $cases = foreach ($command in $caseCommands) {
        $nameNode = @($command.CommandElements | Where-Object {
                $_ -is [System.Management.Automation.Language.StringConstantExpressionAst]
            } | Select-Object -Skip 1 -First 1)
        if ($nameNode.Count -ne 1) {
            throw "Source-contract inventory requires every It block in '$FilePath' to use a literal name."
        }

        $bodyNode = @($command.CommandElements | Where-Object {
                $_ -is [System.Management.Automation.Language.ScriptBlockExpressionAst]
            } | Select-Object -First 1)
        if ($bodyNode.Count -ne 1) {
            throw "Source-contract inventory could not find the body for '$($nameNode[0].Value)'."
        }

        $name = $nameNode[0].Value
        $body = $bodyNode[0].Extent.Text
        $hasPositiveRegex = $body -match 'Should\s+Match'
        $hasNegativeRegex = $body -match 'Should\s+Not\s+Match'
        $hasGraphEvidence = $name -match '(?i)build|project|run[- ]plan|inventory|translation unit|test artifact|package|dependency|runner|case listing|repeat|shuffle' -or
            $body -match '(?i)\.vcxproj|\.props|\.targets|Directory\.Build|Get-RSTestRunPlan|Get-RSTestInventory'

        $category = if ($hasGraphEvidence) {
            'StructuralGraphInvariant'
        }
        elseif ($hasNegativeRegex -and -not $hasPositiveRegex) {
            'LexicalSafetyBan'
        }
        elseif ($hasPositiveRegex -and $hasNegativeRegex) {
            'MixedSourceShape'
        }
        else {
            'BehavioralShadow'
        }

        if ($category -notin $allowedCategories) {
            throw "Source-contract case '$name' received unsupported category '$category'."
        }

        [pscustomobject]@{
            Name = $name
            Category = $category
            Line = $command.Extent.StartLineNumber
            PositiveRegexAssertions = ([regex]::Matches($body, 'Should\s+Match')).Count
            NegativeRegexAssertions = ([regex]::Matches($body, 'Should\s+Not\s+Match')).Count
        }
    }

    $duplicateNames = @($cases | Group-Object Name | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    if ($duplicateNames.Count -ne 0) {
        throw "Source-contract case names must be unique: $($duplicateNames -join ', ')."
    }

    $categoryCounts = [ordered]@{}
    foreach ($category in $allowedCategories) {
        $categoryCounts[$category] = @($cases | Where-Object { $_.Category -eq $category }).Count
    }

    return [pscustomobject]@{
        File = [System.IO.Path]::GetFileName($FilePath)
        Cases = @($cases)
        CategoryCounts = [pscustomobject]$categoryCounts
        ReplacementCandidates = @($cases | Where-Object { $_.Category -in @('BehavioralShadow', 'MixedSourceShape') } | ForEach-Object { $_.Name })
    }
}

function Get-RSTestRunPlanSurfaceInventory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    if (-not (Get-Command Get-RSTestRunPlan -CommandType Function -ErrorAction SilentlyContinue)) {
        . (Join-Path $RepoRoot 'Tools\TestRunPlan.ps1')
    }

    $redSalamanderExe = Join-Path $RepoRoot '.build\x64\Debug\RedSalamander.exe'
    $planArguments = @{
        RepoRoot = $RepoRoot
        Platform = 'x64'
        Configuration = 'Debug'
        RedSalamanderExePath = $redSalamanderExe
    }
    $ciPlan = @(Get-RSTestRunPlan -Suite CI @planArguments)
    $fullPlan = @(Get-RSTestRunPlan -Suite Full @planArguments)

    $projectNames = @(Get-RSTestProjectNames -RepoRoot $RepoRoot)
    $projectBackedSurfaces = foreach ($projectName in $projectNames) {
        $matchingEntries = @($ciPlan + $fullPlan | Where-Object {
                [System.IO.Path]::GetFileNameWithoutExtension([string]$_.Path) -eq $projectName
            })
        $kinds = @($matchingEntries.Kind | Sort-Object -Unique)
        if ($matchingEntries.Count -eq 0) {
            throw "Test project '$projectName' is not represented in the CI or Full run plan."
        }
        if ($kinds.Count -ne 1) {
            throw "Test project '$projectName' has inconsistent run-plan kinds: $($kinds -join ', ')."
        }

        [pscustomobject]@{
            Name = $projectName
            Kind = $kinds[0]
            PlanEntries = @($matchingEntries.Name | Sort-Object -Unique)
        }
    }

    $convertPlan = {
        param([object[]]$Entries)

        return @($Entries | ForEach-Object {
                [pscustomobject]@{
                    Name = $_.Name
                    Kind = $_.Kind
                    Artifact = [System.IO.Path]::GetFileName([string]$_.Path)
                }
            })
    }

    return [pscustomobject]@{
        CI = & $convertPlan $ciPlan
        Full = & $convertPlan $fullPlan
        ProjectBackedSurfaces = @($projectBackedSurfaces)
    }
}

function Get-RSTestInventory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $commandsFiles = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'RedSalamander\SelfTest\Commands') -Filter '*.cpp')
    $commandsCoordinator = Join-Path $RepoRoot 'RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp'
    $compareFiles = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'RedSalamander\SelfTest\CompareDirectories') -Filter '*.cpp')
    $fileOpsCoordinator = Join-Path $RepoRoot 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'
    $performanceFiles = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'Tests\PerformanceTests2') -Filter '*.cpp')
    $nativeTextInputTests = Join-Path $RepoRoot 'Tests\DxUiTests\DxUiTests.NativeTextInput.cpp'
    $toolTestFiles = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'Tools\Tests') -Filter '*.Tests.ps1')
    $sourceContractFile = Join-Path $RepoRoot 'Tools\Tests\TestHarnessSourceContracts.Tests.ps1'
    $syntheticScript = Join-Path $RepoRoot 'Tests\vcpkg-merge-synthetic-test.ps1'
    $lockValidationScript = Join-Path $RepoRoot 'Tests\vcpkg-merge-lock-validation.ps1'
    $runPlanInventory = Get-RSTestRunPlanSurfaceInventory -RepoRoot $RepoRoot

    $commandsFamilyFiles = [ordered]@{
        Settings = 'Commands.SelfTest.Settings.cpp'
        BatchRename = 'Commands.SelfTest.BatchRename.cpp'
        PluginConfig = 'Commands.SelfTest.PluginConfig.cpp'
        Connections = 'Commands.SelfTest.Connections.cpp'
        'Preferences Dispatch' = 'Commands.SelfTest.Preferences.Dispatch.cpp'
        CompareOptions = 'Commands.SelfTest.CompareOptions.cpp'
        Search = 'Commands.SelfTest.Search.cpp'
        Shortcuts = 'Commands.SelfTest.Shortcuts.cpp'
        ViewCommands = 'Commands.SelfTest.ViewCommands.cpp'
        FileOps = 'Commands.SelfTest.FileOps.cpp'
        Navigation = 'Commands.SelfTest.Navigation.cpp'
        Dialogs = 'Commands.SelfTest.Dialogs.cpp'
        ShellCommands = 'Commands.SelfTest.ShellCommands.cpp'
    }
    $commandsFamilyCases = [ordered]@{}
    foreach ($entry in $commandsFamilyFiles.GetEnumerator()) {
        $familyPath = Join-Path $RepoRoot (Join-Path 'RedSalamander\SelfTest\Commands' $entry.Value)
        $commandsFamilyCases[$entry.Key] = Get-RSSelectStringCount -Path @($familyPath) -Pattern 'SelfTest::RunCase\('
    }

    $toolsPesterFileCases = [ordered]@{}
    foreach ($file in ($toolTestFiles | Sort-Object Name)) {
        $toolsPesterFileCases[$file.Name] = Get-RSSelectStringCount -Path @($file.FullName) -Pattern '^\s*It\s'
    }

    return [pscustomobject]@{
        GeneratedAt = (Get-Date).ToString('o')
        SelfTests = [pscustomobject]@{
            Commands = [pscustomobject]@{
                RunCaseRegistrations = Get-RSSelectStringCount -Path @($commandsFiles.FullName) -Pattern 'SelfTest::RunCase\('
                CoordinatorRunCaseRegistrations = Get-RSSelectStringCount -Path @($commandsCoordinator) -Pattern 'SelfTest::RunCase\('
                FamilyRunCaseRegistrations = [pscustomobject]$commandsFamilyCases
            }
            CompareDirectories = [pscustomobject]@{
                RunCaseRegistrations = Get-RSSelectStringCount -Path @($compareFiles.FullName) -Pattern 'SelfTest::RunCase\('
            }
            FileOperations = [pscustomobject]@{
                ActivePhases = Get-RSFileOpsActivePhaseCount -FilePath $fileOpsCoordinator
            }
        }
        Standalone = [pscustomobject]@{
            NativeProjects = @($runPlanInventory.ProjectBackedSurfaces)
            NativeExecutables = @($runPlanInventory.ProjectBackedSurfaces | Where-Object { $_.Kind -eq 'Executable' } | ForEach-Object { $_.Name })
            PerformanceTests2 = [pscustomobject]@{
                TestMethods = Get-RSSelectStringCount -Path @($performanceFiles.FullName) -Pattern 'TEST_METHOD\('
            }
            DxUiTests = [pscustomobject]@{
                NativeTextInputCases = Get-RSSelectStringCount -Path @($nativeTextInputTests) -Pattern '^void\s+TestNativeTextInput'
            }
        }
        Scripts = [pscustomobject]@{
            ToolsPester = [pscustomobject]@{
                Cases = Get-RSSelectStringCount -Path @($toolTestFiles.FullName) -Pattern '^\s*It\s'
                RequiresBuildToolchainCases = Get-RSPesterTagCaseCount -Files $toolTestFiles -Tag 'RequiresBuildToolchain'
                FileCases = [pscustomobject]$toolsPesterFileCases
                SourceContracts = Get-RSSourceContractInventory -FilePath $sourceContractFile
            }
            VcpkgMergeSynthetic = [pscustomobject]@{
                Cases = Get-RSSelectStringCount -Path @($syntheticScript) -Pattern "Write-Host\s+'\[\d+\]"
            }
            VcpkgMergeLockValidation = [pscustomobject]@{
                Cases = Get-RSSelectStringCount -Path @($lockValidationScript) -Pattern 'Write-Host\s+"\[Test\s+\d+\]'
            }
        }
        RunPlan = $runPlanInventory
    }
}

function ConvertTo-RSTestInventoryJson {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Inventory
    )

    $jsonReady = [ordered]@{
        generatedAt = $Inventory.GeneratedAt
        selfTests = [ordered]@{
            commands = [ordered]@{
                runCaseRegistrations = $Inventory.SelfTests.Commands.RunCaseRegistrations
                coordinatorRunCaseRegistrations = $Inventory.SelfTests.Commands.CoordinatorRunCaseRegistrations
                familyRunCaseRegistrations = $Inventory.SelfTests.Commands.FamilyRunCaseRegistrations
            }
            compareDirectories = [ordered]@{
                runCaseRegistrations = $Inventory.SelfTests.CompareDirectories.RunCaseRegistrations
            }
            fileOperations = [ordered]@{
                activePhases = $Inventory.SelfTests.FileOperations.ActivePhases
            }
        }
        standalone = [ordered]@{
            nativeExecutables = @($Inventory.Standalone.NativeExecutables)
            nativeProjects = @($Inventory.Standalone.NativeProjects | ForEach-Object {
                    [ordered]@{
                        name = $_.Name
                        kind = $_.Kind
                        planEntries = @($_.PlanEntries)
                    }
                })
            performanceTests2 = [ordered]@{
                testMethods = $Inventory.Standalone.PerformanceTests2.TestMethods
            }
            dxUiTests = [ordered]@{
                nativeTextInputCases = $Inventory.Standalone.DxUiTests.NativeTextInputCases
            }
        }
        scripts = [ordered]@{
            toolsPester = [ordered]@{
                cases = $Inventory.Scripts.ToolsPester.Cases
                requiresBuildToolchainCases = $Inventory.Scripts.ToolsPester.RequiresBuildToolchainCases
                fileCases = $Inventory.Scripts.ToolsPester.FileCases
                sourceContracts = [ordered]@{
                    file = $Inventory.Scripts.ToolsPester.SourceContracts.File
                    categoryCounts = $Inventory.Scripts.ToolsPester.SourceContracts.CategoryCounts
                    replacementCandidates = @($Inventory.Scripts.ToolsPester.SourceContracts.ReplacementCandidates)
                    cases = @($Inventory.Scripts.ToolsPester.SourceContracts.Cases | ForEach-Object {
                            [ordered]@{
                                name = $_.Name
                                category = $_.Category
                                line = $_.Line
                                positiveRegexAssertions = $_.PositiveRegexAssertions
                                negativeRegexAssertions = $_.NegativeRegexAssertions
                            }
                        })
                }
            }
            vcpkgMergeSynthetic = [ordered]@{
                cases = $Inventory.Scripts.VcpkgMergeSynthetic.Cases
            }
            vcpkgMergeLockValidation = [ordered]@{
                cases = $Inventory.Scripts.VcpkgMergeLockValidation.Cases
            }
        }
        runPlan = [ordered]@{
            ci = @($Inventory.RunPlan.CI | ForEach-Object {
                    [ordered]@{ name = $_.Name; kind = $_.Kind; artifact = $_.Artifact }
                })
            full = @($Inventory.RunPlan.Full | ForEach-Object {
                    [ordered]@{ name = $_.Name; kind = $_.Kind; artifact = $_.Artifact }
                })
            projectBackedSurfaces = @($Inventory.RunPlan.ProjectBackedSurfaces | ForEach-Object {
                    [ordered]@{ name = $_.Name; kind = $_.Kind; planEntries = @($_.PlanEntries) }
                })
        }
    }

    return ($jsonReady | ConvertTo-Json -Depth 10)
}
