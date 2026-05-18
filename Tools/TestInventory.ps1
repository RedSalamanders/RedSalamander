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

function Get-RSTestInventory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $commandsFiles = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'RedSalamander\SelfTest\Commands') -Filter '*.cpp')
    $compareFiles = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'RedSalamander\SelfTest\CompareDirectories') -Filter '*.cpp')
    $fileOpsCoordinator = Join-Path $RepoRoot 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'
    $performanceFiles = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'Tests\PerformanceTests2') -Filter '*.cpp')
    $nativeTextInputTests = Join-Path $RepoRoot 'Tests\DxUiTests\DxUiTests.NativeTextInput.cpp'
    $toolTestFiles = @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'Tools\Tests') -Filter '*.Tests.ps1')
    $syntheticScript = Join-Path $RepoRoot 'Tests\vcpkg-merge-synthetic-test.ps1'
    $lockValidationScript = Join-Path $RepoRoot 'Tests\vcpkg-merge-lock-validation.ps1'

    $standaloneNames = @(
        'DxUiTests',
        'ViewerPETests',
        'ViewerSqliteTests',
        'MonitorTest',
        'LocalizationTests'
    )

    return [pscustomobject]@{
        GeneratedAt = (Get-Date).ToString('o')
        SelfTests = [pscustomobject]@{
            Commands = [pscustomobject]@{
                RunCaseRegistrations = Get-RSSelectStringCount -Path @($commandsFiles.FullName) -Pattern 'SelfTest::RunCase\('
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
