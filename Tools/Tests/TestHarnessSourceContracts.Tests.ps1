Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Get-RSText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return Get-Content -LiteralPath (Join-Path $repoRoot $Path) -Raw
}

Describe 'Test harness source contracts' {
    It 'rejects unknown DxUiTests switches instead of silently running every suite' {
        $source = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.cpp'

        $source | Should Match 'Unknown argument'
        $source | Should Match 'arg\[0\]\s*==\s*L''-'''
    }

    It 'rejects empty DxUiTests option values with targeted diagnostics' {
        $source = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.cpp'

        $source | Should Match 'Missing suite name'
        $source | Should Match 'Missing perf JSONL path'
    }

    It 'documents and enforces bounded self-test timeout multipliers' {
        $source = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'

        $source | Should Match 'kSelfTestTimeoutMultiplierMin'
        $source | Should Match 'kSelfTestTimeoutMultiplierMin\s*=\s*0\.1'
        $source | Should Match 'kSelfTestTimeoutMultiplierMax'
        $source | Should Match 'kSelfTestTimeoutMultiplierMax\s*=\s*100\.0'
        $source | Should Match 'std::isfinite'
        $source | Should Match 'std::clamp\(parsed,\s*kSelfTestTimeoutMultiplierMin,\s*kSelfTestTimeoutMultiplierMax\)'
        $source | Should Match 'Clamped --selftest-timeout-multiplier'
        $source | Should Match 'Invalid --selftest-timeout-multiplier'
    }

    It 'keeps scaled self-test timeouts finite and nonzero for nonzero bases' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'

        $source | Should Match 'std::isfinite\(scaled\)'
        $source | Should Match 'baseMs\s*>\s*0u'
        $source | Should Match 'return 1u;'
    }

    It 'gives ViewerPETests nested shell churn its own timeout budget' {
        $source = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'

        $source | Should Match 'kViewerHarnessDefaultTimeout\s*=\s*120000ms'
        $source | Should Match 'kViewerShellComboLongRunTimeout\s*=\s*600000ms'
        $source | Should Match 'kViewerShellComboLongRunTimeout\s*>\s*kViewerHarnessDefaultTimeout'
        $source | Should Match 'TestViewerShellComboHostsLongRunOpenCloseStayStable",\s*kViewerShellComboLongRunTimeout'
    }

    It 'requires git identity before accepting self-test archive repo roots' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'

        $source | Should Match 'kRepoGitDirName'
        $source | Should Match 'IsRepoRootCandidate'
        $source | Should Match 'candidate\s*/\s*kRepoGitDirName'
    }

    It 'keeps self-test archive parent walking bounded by a named limit' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'

        $source | Should Match 'kRepoRootParentWalkLimit'
        $source | Should Match 'i\s*<\s*kRepoRootParentWalkLimit'
    }

    It 'keeps Commands Preferences chunks namespace-self-contained' {
        $coordinator = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.cpp'
        $coordinator.TrimEnd() | Should Not Match 'namespace\s*\{\s*$'

        $chunkNames = @(
            'Commands.SelfTest.Preferences.ChromeAndPlugins.cpp',
            'Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp',
            'Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp',
            'Commands.SelfTest.Preferences.PluginsThemesAdvanced.cpp',
            'Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp',
            'Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp',
            'Commands.SelfTest.Preferences.Dispatch.cpp'
        )

        foreach ($chunkName in $chunkNames) {
            $chunk = Get-RSText -Path "RedSalamander\SelfTest\Commands\$chunkName"
            $chunk | Should Match '^namespace\s*\{'
            $chunk.TrimEnd() | Should Match '\}\s*//\s*namespace'
        }
    }

    It 'keeps IconCache association LRU perf emission outside the association cache lock helper' {
        $source = Get-RSText -Path 'RedSalamander\IconCache.cpp'
        $bodyMatch = [regex]::Match($source, 'EvictAssociationQueryBatch\(\)[\s\S]*?\n\}')
        $bodyMatch.Success | Should Be $true
        $bodyMatch.Value | Should Not Match 'PerfEmit'
        $source | Should Match 'EmitAssociationLruEvictScanMetric'
    }

    It 'keeps FolderView thumbnail WIC factory ownership local to the processing thread' {
        $header = Get-RSText -Path 'RedSalamander\FolderView.h'
        $source = Get-RSText -Path 'RedSalamander\FolderView.Icons.cpp'

        $header | Should Not Match '_thumbnailWicFactory'
        $source | Should Match 'EnsureThumbnailWicFactory\(\s*wil::com_ptr<IWICImagingFactory>&'
        $source | Should Match 'thumbnailWicFactory'
    }

    It 'keeps IMAP summary repair budget exhaustion local to the current repair pass' {
        $source = Get-RSText -Path 'Plugins\FileSystemCurl\FileSystemCurl.Imap.cpp'
        $bodyMatch = [regex]::Match($source, 'auto\s+repairMissingSummaries\s*=\s*\[[\s\S]*?return\s+repairResult;\s*\n\s*\};')
        $bodyMatch.Success | Should Be $true

        $bodyMatch.Value | Should Not Match 'kRepairBudgetExceededHr'
        $bodyMatch.Value | Should Not Match 'tryConsumeRepairFetch[\s\S]*?noexcept'
        $bodyMatch.Value | Should Match 'break;'
    }

    It 'clears FolderView incremental-search layout effects over the full DWrite label text' {
        $source = Get-RSText -Path 'RedSalamander\FolderView.Interaction.cpp'
        $bodyMatch = [regex]::Match($source, 'void\s+FolderView::ClearIncrementalSearchLayoutEffects\(\)\s+noexcept[\s\S]*?\n\}')
        $bodyMatch.Success | Should Be $true
        $bodyMatch.Value | Should Match 'item\.displayName'
        $bodyMatch.Value | Should Not Match 'GetVisualDisplayName'
    }

    It 'matches FolderView incremental-search highlights against visual display names' {
        $source = Get-RSText -Path 'RedSalamander\FolderView.Interaction.cpp'
        $bodyMatch = [regex]::Match($source, 'void\s+FolderView::UpdateIncrementalSearchHighlightForFocusedItem\(\)[\s\S]*?\n\}')
        $bodyMatch.Success | Should Be $true

        $bodyMatch.Value | Should Match 'GetVisualDisplayName\(item\)'
        $bodyMatch.Value | Should Not Match 'FindIncrementalSearchMatchOffset\(item\.displayName\)'
    }

    It 'cleans stale FolderView incremental-search drawing effects when search is inactive' {
        $header = Get-RSText -Path 'RedSalamander\FolderView.h'
        $rendering = Get-RSText -Path 'RedSalamander\FolderView.Rendering.cpp'
        $interaction = Get-RSText -Path 'RedSalamander\FolderView.Interaction.cpp'

        $header | Should Match '_incrementalSearchLayoutEffectsDirty'
        $rendering | Should Match '_incrementalSearchLayoutEffectsDirty'
        $rendering | Should Match 'ClearIncrementalSearchLayoutEffects\(\)'
        $interaction | Should Match '_incrementalSearchLayoutEffectsDirty\s*=\s*true'
        $interaction | Should Match '_incrementalSearchLayoutEffectsDirty\s*=\s*false'
    }

    It 'exposes runner-native self-test case listing without executing case bodies' {
        $main = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'
        $common = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.h'
        $commandsHeader = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.h'
        $commandsSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp'
        $compareHeader = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.h'
        $compareSource = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'

        $main | Should Match '--selftest-list-cases'
        $main | Should Match 'BuildSelfTestCaseListJson'
        $main | Should Match 'FileOperationsSelfTest::BuildExpectedCaseNames'
        $common | Should Match 'listCasesOnly'
        $commandsHeader | Should Match 'ListCases'
        $commandsSource | Should Match 'listCasesOnly'
        $compareHeader | Should Match 'ListCases'
        $compareSource | Should Match 'kCompareCaseNames'
    }

    It 'uses shared self-test result emission for ad hoc and state-machine cases' {
        $commonHeader = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.h'
        $commonSource = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'
        $compareSource = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'
        $fileOpsSource = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'

        $commonHeader | Should Match 'AppendCaseResult'
        $commonSource | Should Match 'void\s+AppendCaseResult'
        $compareSource | Should Match 'SelfTest::AppendCaseResult'
        $fileOpsSource | Should Match 'SelfTest::AppendCaseResult'
    }

    It 'keeps CompareDirectories runner-listed case names unique' {
        $compareSource = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'
        $arrayMatch = [regex]::Match($compareSource, 'kCompareCaseNames[\s\S]*?=\s*\{(?<body>[\s\S]*?)\};')
        $arrayMatch.Success | Should Be $true

        $names = @([regex]::Matches($arrayMatch.Groups['body'].Value, 'L"(?<name>[^"]+)"') | ForEach-Object { $_.Groups['name'].Value })
        $duplicates = @($names | Group-Object | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name)

        @($duplicates).Count | Should Be 0
    }

    It 'keeps CompareDirectories literal RunCase names listed for runner coverage' {
        $compareSource = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'
        $arrayMatch = [regex]::Match($compareSource, 'kCompareCaseNames[\s\S]*?=\s*\{(?<body>[\s\S]*?)\};')
        $arrayMatch.Success | Should Be $true

        $listedNames = @([regex]::Matches($arrayMatch.Groups['body'].Value, 'L"(?<name>[^"]+)"') | ForEach-Object { $_.Groups['name'].Value })
        $listedSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($name in $listedNames) {
            [void]$listedSet.Add($name)
        }

        $compareDir = Join-Path $repoRoot 'RedSalamander\SelfTest\CompareDirectories'
        $literalRunCaseNames = @(
            Get-ChildItem -LiteralPath $compareDir -Filter '*.cpp' |
                ForEach-Object {
                    $source = Get-Content -LiteralPath $_.FullName -Raw
                    [regex]::Matches($source, 'SelfTest::RunCase\(\s*options,\s*suite,\s*L"(?<name>[^"]+)"') |
                        ForEach-Object { $_.Groups['name'].Value }
                } |
                Sort-Object -Unique
        )

        $missing = @($literalRunCaseNames | Where-Object { -not $listedSet.Contains($_) })
        @($missing).Count | Should Be 0
    }

    It 'opens the file-operations custom speed-limit prompt asynchronously for self-tests' {
        $popupSource = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.Popup.cpp'

        $popupSource | Should Match 'kFileOperationsPopupDeferredSpeedLimitPromptMessage'
        $popupSource | Should Match 'Let self-test callers advance their state before the modal prompt loop starts'
        $popupSource | Should Match 'PostMessageW\(hwnd,\s*kFileOperationsPopupDeferredSpeedLimitPromptMessage'
        $popupSource | Should Not Match 'payload->kind\s*==\s*PopupHitTest::Kind::TaskSpeedLimit\s*&&\s*payload->data\s*==\s*1u\)[\s\S]{0,160}ShowCustomSpeedLimitPromptForTask'
    }

    It 'keeps FileOperations self-test case prefixes runnable' {
        $fileOpsSource = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'

        $fileOpsSource | Should Match 'FindPhasesByPrefix'
        $fileOpsSource | Should Match 'StartsWithIgnoreCase\(StepToString\(step\),\s*prefix\)'
        $fileOpsSource | Should Match 'selection\.activePhases\s*=\s*prefixMatches'
        $fileOpsSource | Should Match 'ResolveRunSelection\(filter\)\.recognized'
    }

    It 'keeps native DxUi text hosts compatible with Win32 edit text and selection messages' {
        $windowHostSource = Get-RSText -Path 'Common\DxUi\DxUi.WindowHost.cpp'
        $nativeTextInputSource = Get-RSText -Path 'Common\DxUi\DxUi.NativeTextInput.cpp'

        $windowHostSource | Should Match 'case WM_GETTEXTLENGTH:'
        $windowHostSource | Should Match 'case WM_GETTEXT:'
        $windowHostSource | Should Match 'case WM_SETTEXT:'
        $windowHostSource | Should Match 'case EM_GETSEL:'
        $windowHostSource | Should Match 'case EM_SETSEL:'
        $windowHostSource | Should Match 'case EM_REPLACESEL:'

        $nativeTextInputSource | Should Match 'case WM_GETTEXTLENGTH:'
        $nativeTextInputSource | Should Match 'case WM_GETTEXT:'
        $nativeTextInputSource | Should Match 'case WM_SETTEXT:'
        $nativeTextInputSource | Should Match 'case EM_GETSEL:'
        $nativeTextInputSource | Should Match 'case EM_SETSEL:'
        $nativeTextInputSource | Should Match 'case EM_REPLACESEL:'
    }
}
