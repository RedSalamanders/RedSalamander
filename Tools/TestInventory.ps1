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

function Get-RSRegexInt {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    $match = [Regex]::Match($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $match.Success) {
        return $null
    }

    return [int]$match.Groups[1].Value
}

function Get-RSRegexInts {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    return @([Regex]::Matches($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline) |
        ForEach-Object { [int]$_.Groups[1].Value })
}

function Get-RSMarkdownTableCountMap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$StartHeading,

        [Parameter(Mandatory = $true)]
        [string]$EndHeading
    )

    $start = $Text.IndexOf($StartHeading, [StringComparison]::Ordinal)
    $end = $Text.IndexOf($EndHeading, [Math]::Max(0, $start + $StartHeading.Length), [StringComparison]::Ordinal)
    if ($start -lt 0 -or $end -le $start) {
        return [pscustomobject]@{}
    }

    $map = [ordered]@{}
    $section = $Text.Substring($start, $end - $start)
    foreach ($match in [Regex]::Matches($section, '(?m)^\|\s*([^|]+?)\s*\|\s*(?:`[^`]+`\s*\|\s*)?(\d+)\s*\|')) {
        $name = $match.Groups[1].Value.Trim().Trim('`')
        if ($name -ne 'Family' -and $name -ne 'File') {
            $map[$name] = [int]$match.Groups[2].Value
        }
    }

    return [pscustomobject]$map
}

function Get-RSTestInventoryDocSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $text = Get-Content -LiteralPath $Path -Raw

    return [pscustomobject]@{
        CommandsRunCases = Get-RSRegexInt -Text $text -Pattern 'Commands:\s+(\d+)\s+static'
        CompareRunCases = Get-RSRegexInt -Text $text -Pattern 'CompareDirectories:\s+(\d+)\s+static'
        FileOpsActivePhases = Get-RSRegexInt -Text $text -Pattern 'FileOperations:\s+(\d+)\s+active'
        PerformanceTestMethods = Get-RSRegexInt -Text $text -Pattern 'PerformanceTests2:\s+(\d+)\s+CppUnitTest'
        NativeTextInputCases = Get-RSRegexInt -Text $text -Pattern '\|\s+(?:NativeTextInput|\*\*DxUiTests / NativeTextInput\*\*)\s+\|(?:\s+`[^`]+`\s+\|)?\s+(\d+)\s+\|'
        ToolsPesterCases = Get-RSRegexInt -Text $text -Pattern '(\d+)\s+Pester-style'
        VcpkgSyntheticCases = Get-RSRegexInt -Text $text -Pattern '(\d+)\s+fast\s+synthetic'
    }
}

function Get-RSTestsReadmeInventorySnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $text = Get-Content -LiteralPath $Path -Raw

    return [pscustomobject]@{
        CommandsListedCases = @(
            Get-RSRegexInt -Text $text -Pattern '\|\s*Self-Tests \(in-process\)\s*\|\s*3\s*\|\s*(\d+)\s+Commands listed cases'
            Get-RSRegexInt -Text $text -Pattern '## 1\. Commands Self-Test Suite — (\d+) runner-listed cases'
            Get-RSRegexInt -Text $text -Pattern '--selftest-list-cases --commands-selftest`\s*\r?\nlists (\d+) cases'
        )
        CommandsStaticRunCases = @(Get-RSRegexInts -Text $text -Pattern '--selftest-list-cases --commands-selftest`\s*\r?\nlists\s+\d+\s+cases\.\s+The source fallback scan reports\s+(\d+)\s+static')
        CompareListedCases = @(
            Get-RSRegexInt -Text $text -Pattern '\|\s*Self-Tests \(in-process\)\s*\|\s*3\s*\|\s*\d+\s+Commands listed cases,\s*(\d+)\s+Compare listed cases'
            Get-RSRegexInt -Text $text -Pattern '## 2\. Compare Directories Self-Test Suite — (\d+) runner-listed cases'
            Get-RSRegexInt -Text $text -Pattern '--selftest-list-cases --compare-selftest`\s*\r?\nlists (\d+) cases'
        )
        CompareStaticRunCases = @(Get-RSRegexInts -Text $text -Pattern 'source fallback scan reports (\d+) static\s*\r?\n`SelfTest::RunCase` call sites plus explicit')
        FileOpsListedCases = @(
            Get-RSRegexInt -Text $text -Pattern '\|\s*Self-Tests \(in-process\)\s*\|\s*3\s*\|[^|]*,\s*(\d+)\s+FileOps listed phases'
            Get-RSRegexInt -Text $text -Pattern '## 3\. File Operations Self-Test Suite — (\d+) runner-listed phases'
            Get-RSRegexInt -Text $text -Pattern '--selftest-list-cases --fileops-selftest`\s*\r?\nlists (\d+) phases'
        )
        FileOpsActivePhases = @(Get-RSRegexInts -Text $text -Pattern 'phases: setup,\s*(\d+) active ordered phases')
        ToolsPesterCases = @(Get-RSRegexInts -Text $text -Pattern '(\d+) Pester-style/tool cases')
        CommandsFamilyCases = Get-RSMarkdownTableCountMap -Text $text -StartHeading '| Family | File | Tests | Coverage |' -EndHeading '## 2. Compare Directories'
        ToolsPesterFileCases = Get-RSMarkdownTableCountMap -Text $text -StartHeading '| File | Cases | Coverage |' -EndHeading 'Fast vcpkg merge coverage:'
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
    $syntheticScript = Join-Path $RepoRoot 'Tests\vcpkg-merge-synthetic-test.ps1'
    $lockValidationScript = Join-Path $RepoRoot 'Tests\vcpkg-merge-lock-validation.ps1'

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

    $standaloneNames = @(
        'DxUiTests',
        'FileSystemCurlTests',
        'ViewerPETests',
        'ViewerSqliteTests',
        'MonitorTest',
        'LocalizationTests',
        'RedConfigureTests'
    )

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
            NativeExecutables = $standaloneNames
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
            }
            VcpkgMergeSynthetic = [pscustomobject]@{
                Cases = Get-RSSelectStringCount -Path @($syntheticScript) -Pattern "Write-Host\s+'\[\d+\]"
            }
            VcpkgMergeLockValidation = [pscustomobject]@{
                Cases = Get-RSSelectStringCount -Path @($lockValidationScript) -Pattern 'Write-Host\s+"\[Test\s+\d+\]'
            }
        }
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
            }
            vcpkgMergeSynthetic = [ordered]@{
                cases = $Inventory.Scripts.VcpkgMergeSynthetic.Cases
            }
            vcpkgMergeLockValidation = [ordered]@{
                cases = $Inventory.Scripts.VcpkgMergeLockValidation.Cases
            }
        }
    }

    return ($jsonReady | ConvertTo-Json -Depth 8)
}
