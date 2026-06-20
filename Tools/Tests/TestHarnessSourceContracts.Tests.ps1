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

    It 'bounds DirectoryInfoCache eviction when protected entries exceed the cache budget' {
        $source = Get-RSText -Path 'RedSalamander\DirectoryInfoCache.cpp'
        $bodyMatch = [regex]::Match($source, 'void\s+DirectoryInfoCache::MaybeEvictLocked\([\s\S]*?\n\}')
        $bodyMatch.Success | Should Be $true
        $bodyMatch.Value | Should Match 'scanLimit\s*=\s*_lru\.size\(\)'
        $bodyMatch.Value | Should Match 'scanned\s*<\s*scanLimit'
        $bodyMatch.Value | Should Match '\+\+scanned'
        $bodyMatch.Value | Should Match 'only pinned/borrowed/loading entries cannot spin'
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

    It 'clears FolderView incremental-search layout effects over the visual label text' {
        # Highlights are matched and applied against GetVisualDisplayName (see
        # the sibling contract below), so clearing must use the same text the
        # label layout and highlight ranges were built from.
        $source = Get-RSText -Path 'RedSalamander\FolderView.Interaction.cpp'
        $bodyMatch = [regex]::Match($source, 'void\s+FolderView::ClearIncrementalSearchLayoutEffects\(\)\s+noexcept[\s\S]*?\n\}')
        $bodyMatch.Success | Should Be $true
        $bodyMatch.Value | Should Match 'GetVisualDisplayName\(item\)'
        $bodyMatch.Value | Should Not Match 'item\.displayName'
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

    It 'keeps FileOperations pre-calc perf counters race-free under worker concurrency' {
        $header = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperationsInternal.h'
        $source = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'

        $header | Should Match 'std::atomic<uint64_t>\s+preCalcUs'
        $header | Should Match 'std::atomic<uint64_t>\s+preCalcCallbackCount'
        $header | Should Match 'std::atomic<uint64_t>\s+preCalcCallbackUs'
        $header | Should Match 'std::atomic<uint64_t>\s+preCalcLockWaitUs'

        $source | Should Match '_perf\.preCalcUs\.fetch_add'
        $source | Should Match '_perf\.preCalcCallbackCount\.fetch_add'
        $source | Should Match '_perf\.preCalcCallbackUs\.fetch_add'
        $source | Should Match '_perf\.preCalcLockWaitUs\.fetch_add'

        $source | Should Not Match '\+\+_perf\.preCalcCallbackCount'
        $source | Should Not Match '_perf\.preCalcCallbackCount\+\+'
        $source | Should Not Match '_perf\.preCalcUs\s*\+='
        $source | Should Not Match '_perf\.preCalcCallbackUs\s*\+='
        $source | Should Not Match '_perf\.preCalcLockWaitUs\s*\+='
    }

    It 'keeps FileSystem dynamic scheduler cleanup explicit and recursive-copy queueing unbounded' {
        $source = Get-RSText -Path 'Plugins\FileSystem\FileSystem.FileOps.cpp'

        $source | Should Not Match 'maxQueuedItems'
        $source | Should Not Match 'queue\.items\.size\(\)\s*<'
        $source | Should Not Match 'queue\.cv\.wait\(lock'

        $schedulerMatch = [regex]::Match($source, 'class\s+SharedFileOpsJobScheduler\s+final[\s\S]*?\n\};\s*\n\s*SharedFileOpsJobScheduler&')
        $schedulerMatch.Success | Should Be $true
        $scheduler = $schedulerMatch.Value

        $hasSchedulableMatch = [regex]::Match($scheduler, '\[\[nodiscard\]\]\s+bool\s+hasSchedulableWorkLocked\(\)\s+noexcept\s*\{(?<body>[\s\S]*?)\n\s*\}')
        $hasSchedulableMatch.Success | Should Be $true
        $hasSchedulableMatch.Groups['body'].Value | Should Not Match 'cleanupJobsLocked'

        $scheduler | Should Match 'while\s*\(!\s*job->done\.load[\s\S]*?cleanupJobsLocked\(\);[\s\S]*?_cv\.wait\(lock\);[\s\S]*?cleanupJobsLocked\(\);'
    }

    It 'keeps cross-filesystem MOVE source deletion behind one bridge-copy completion helper' {
        $source = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'

        $source | Should Match 'ShouldDeleteMoveSourceAfterBridgeCopy'
        $source | Should Match 'SUCCEEDED\(copyHr\)[\s\S]{0,360}skippedFileConflictCount\s*==\s*0'
        $source | Should Not Match 'moveCopyCompleted\s*=\s*bridgeSkippedDirectoryReparseCount\s*==\s*0\s*&&\s*!\s*bridgeRootDirectoryReparseSkipped'
        $source | Should Not Match 'bridgeLeftSourceAuthoritative\s*=\s*bridge\.skippedDirectoryReparseCount\s*>\s*0'
    }

    It 'keeps FileOperations conflict routing in one arbiter and one action-layout builder' {
        $header = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperationsInternal.h'
        $source = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'

        $header | Should Match 'struct\s+ConflictArbiter'
        $header | Should Match 'ConflictArbiter\s+_conflictArbiter'

        ([regex]::Matches($source, '\[\[nodiscard\]\]\s+ConflictActionLayout\s+BuildConflictActionLayout')).Count | Should Be 1
        $source | Should Not Match 'clearCachedDecision'
        $source | Should Not Match 'const\s+auto\s+setConflictPromptLocked'
        $source | Should Not Match 'const\s+auto\s+waitForConflictDecision'
        $source | Should Not Match 'const\s+auto\s+getCachedDecision'
        $source | Should Not Match 'const\s+auto\s+setCachedDecision'
    }

    It 'keeps Riptide low-risk hardening invariants wired' {
        $fileOps = Get-RSText -Path 'Plugins\FileSystem\FileSystem.FileOps.cpp'
        $runtime = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.Runtime.Part.cpp'
        $bridge = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'
        $storage = Get-RSText -Path 'Plugins\FileSystem\FileSystem.cpp'
        $popup = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.Popup.cpp'
        $microsoftDrive = Get-RSText -Path 'Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.cpp'
        $s3 = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.Directory.cpp'
        $findFiles = Get-RSText -Path 'RedSalamander\FindFilesWindow.cpp'
        $folderWindowHeader = Get-RSText -Path 'RedSalamander\FolderWindow.h'
        $folderWindow = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.cpp'

        $fileOps | Should Not Match 'IsDirectoryEmpty'
        $runtime | Should Match 'sourceCandidates'
        $runtime | Should Match 'destinationCandidates'
        $runtime | Should Match 'for\s*\(const std::wstring& sourceCandidate'
        $runtime | Should Match 'for\s*\(const std::wstring& destinationCandidate'
        $bridge | Should Match 'const\s+FileSystemFlags\s+tempFlags[\s\S]{0,260}~[\s\S]{0,160}FILESYSTEM_FLAG_ALLOW_OVERWRITE'
        $bridge | Should Not Match 'const\s+FileSystemFlags\s+tempFlags\s*=\s*static_cast<FileSystemFlags>\([\s\S]{0,220}\|\s*static_cast<uint32_t>\(FILESYSTEM_FLAG_ALLOW_OVERWRITE\)'
        $storage | Should Match 'else\s+if\s*\(\s*busTypeKnown\s*&&\s*busType\s*==\s*BusTypeNvme\s*\)'
        $popup | Should Match 'if\s*\(\s*weightCount\s*<\s*weights\.size\(\)\s*\)\s*\{[\s\S]{0,160}\+\+weightCount;'
        $popup | Should Not Match '\}\s*\+\+weightCount;\s*\}'

        $microsoftDrive | Should Match 'using\s+CancelProbe\s*=\s*std::function<HRESULT\(\)>'
        $microsoftDrive | Should Match 'MergeMoveFolderIntoExisting[\s\S]{0,800}const\s+CancelProbe&\s+checkCancel'
        $microsoftDrive | Should Match 'MoveOrRenameItem[\s\S]{0,800}const\s+CancelProbe&\s+checkCancel'
        $microsoftDrive | Should Match 'const\s+CancelProbe\s+checkCancel'
        $microsoftDrive | Should Match 'MoveOrRenameItem[\s\S]{0,520}checkCancel'

        $s3 | Should Match 'kMaxS3PerObjectConflictRetries'
        $s3 | Should Match 'unsigned\s+int\s+retryCount\s*=\s*0'
        $s3 | Should Match 'case\s+FileSystemIssueAction::Retry:[\s\S]{0,520}checkCancel[\s\S]{0,520}retryCount'

        $findFiles | Should Match 'uint64_t\s+taskId\s*=\s*0'
        $findFiles | Should Match 'uint64_t\s+createdTickMs\s*=\s*0'
        $findFiles | Should Match 'ReapExpiredPendingResultRemovals'
        $findFiles | Should Match 'PendingResultRemoval\{\.taskId\s*='
        $folderWindowHeader | Should Match 'std::weak_ptr<void>\s+lifetimeGuard'
        $folderWindowHeader | Should Match 'bool\s+hasLifetimeGuard\s*=\s*false'
        $folderWindow | Should Match 'hasLifetimeGuard\s*&&\s*subscription\.lifetimeGuard\.expired\(\)'
    }

    It 'keeps Floodgate data-safety invariants wired' {
        $bridge = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'
        $bridgeQueue = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.Queue.Part.cpp'
        $fileOpsSelfTest = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'
        $fairstreamSelfTest = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.Fairstream.cpp'
        $s3 = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.Directory.cpp'
        $curl = Get-RSText -Path 'Plugins\FileSystemCurl\FileSystemCurl.CopyMove.cpp'
        $microsoftDrive = Get-RSText -Path 'Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.cpp'
        $folderView = Get-RSText -Path 'RedSalamander\FolderView.cpp'
        $folderViewEnumeration = Get-RSText -Path 'RedSalamander\FolderView.Enumeration.cpp'
        $folderViewInteraction = Get-RSText -Path 'RedSalamander\FolderView.Interaction.cpp'
        $navigationCommands = Get-RSText -Path 'RedSalamander\FolderWindow.FileSystem.Navigation.Part.cpp'
        $mainWindow = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'

        $s3 | Should Match 'HasPlannedDestinationAncestorCollision'
        $s3 | Should Match 'EstimateTransferBytes[\s\S]{0,2200}HasPlannedDestinationAncestorCollision\(plan\)[\s\S]{0,160}ERROR_ALREADY_EXISTS'
        $s3 | Should Match 'ExecuteCopyOrMove[\s\S]{0,2600}HasPlannedDestinationAncestorCollision\(plan\)[\s\S]{0,160}ERROR_ALREADY_EXISTS'
        $s3 | Should Match 'RunDebugPlannedDestinationAncestorCollisionSelfTest'
        $s3 | Should Match 'src/a/b'

        $bridge | Should Match 'hrReaderSize\s*=\s*reader->GetSize\(&fileTotalBytes\)'
        $bridge | Should Match 'task\._operation\s*==\s*FILESYSTEM_MOVE[\s\S]{0,240}ERROR_PARTIAL_COPY[\s\S]{0,360}bridge\.integrity\.sourceSizeUnknown'
        $bridge | Should Match 'PromoteTempToFinalPath[\s\S]{0,2200}CreateFileReader\(destinationPath\.c_str\(\),\s*destinationReader\.addressof\(\)\)'
        $bridge | Should Match 'bridge\.integrity\.destinationSizeMismatch'
        $bridgeQueue | Should Match 'SetFileOpsBridgeFailNextSourceGetSizeForSelfTest'
        $fairstreamSelfTest | Should Match 'Floodgate_CrossFsMoveGetSizeFailurePreservesSource[\s\S]{0,1800}SetFileOpsBridgeFailNextSourceGetSizeForSelfTest\(1\)'
        $fairstreamSelfTest | Should Match 'Floodgate_CrossFsMoveGetSizeFailurePreservesSource'
        $fairstreamSelfTest | Should Match 'ERROR_PARTIAL_COPY'
        $fairstreamSelfTest | Should Match 'sourceText\s*!=\s*kSourceContents'
        $fileOpsSelfTest | Should Match 'Floodgate_CrossFsMoveGetSizeFailurePreservesSource'
        # The full per-family suite only runs cases that are MEMBERS of a kFileOpsFamily* array; a case present
        # only in the enum/name-map/kFileOpsPhaseOrder is reported but silently skipped. Assert the data-safety
        # case is actually inside the Fairstream family literal (between its '{{' and closing '}}').
        $fileOpsSelfTest | Should Match 'kFileOpsFamilyFairstream\{\{(?:(?!\}\})[\s\S])*?Floodgate_CrossFsMoveGetSizeFailurePreservesSource'

        $curl | Should Match 'CurlUploadFromFile\(destinationConn,\s*stagedRemotePath[\s\S]{0,500}GetEntryInfo\(destinationConn,\s*stagedRemotePath,\s*stagedInfo\)'
        $curl | Should Match 'stagedInfo\.sizeBytes\s*!=\s*fileSize[\s\S]{0,180}RemoteDeleteFile\(destinationConn,\s*stagedRemotePath\)[\s\S]{0,120}ERROR_PARTIAL_COPY'

        $microsoftDrive | Should Match 'incompleteDueToInvalidChildNameOut'
        $microsoftDrive | Should Match 'sourceEnumerationIncomplete[\s\S]{0,260}allChildrenMoved\s*=\s*!\s*sourceEnumerationIncomplete'
        $microsoftDrive | Should Match 'RunDebugEmptyNameChildBlocksSourceDeleteSelfTest'
        $microsoftDrive | Should Match 'RunDebugMoveRejectsInvalidOptionsSizeSelfTest'
        $microsoftDrive | Should Match 'options\s*!=\s*nullptr\s*&&\s*options->sizeBytes\s*!=\s*sizeof\(FileSystemOptions\)'

        $folderView | Should Match 'const\s+bool\s+folderChanged[\s\S]{0,420}if\s*\(\s*folderChanged\s*\)[\s\S]{0,120}ExitIncrementalSearch\(\)'
        $folderViewEnumeration | Should Match 'bool\s+isRefresh\s*=\s*false[\s\S]{0,420}if\s*\(\s*!\s*isRefresh\s*\)[\s\S]{0,120}ExitIncrementalSearch\(\)'
        $folderViewInteraction | Should Match 'focusStayedInView[\s\S]{0,220}newFocus\s*==\s*_hWnd\.get\(\)[\s\S]{0,160}IsChild\(_hWnd\.get\(\),\s*newFocus\)'
        $folderViewInteraction | Should Match 'newFocus\s*!=\s*nullptr\s*&&\s*!\s*focusStayedInView[\s\S]{0,120}ExitIncrementalSearch\(\)'
        $navigationCommands | Should Match 'void\s+FolderWindow::CommandQuickSearch[\s\S]*?PostMessageW\(_hWnd\.get\(\),\s*WndMsg::kPaneRestoreFolderFocus'
        $folderViewInteraction | Should Not Match 'QuickSearch OnKillFocus|QuickSearch Exit|DescribeQuickSearchFocusTarget'
        $navigationCommands | Should Not Match 'CommandQuickSearch:'
        $mainWindow | Should Not Match 'WM_COMMAND QuickSearch'
    }

    It 'keeps Riptide test-credibility gaps closed' {
        $fairstream = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.Fairstream.cpp'
        $popup = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.Popup.cpp'
        $popupHeader = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.Popup.h'
        $selfTest = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'

        $fairstream | Should Match 'shouldOverwriteConflict'
        $fairstream | Should Match 'Task::ConflictAction::Overwrite'
        $fairstream | Should Match 'Task::ConflictAction::Skip'
        $fairstream | Should Match 'kSourceBytes'
        $fairstream | Should Match 'kDestinationBytes'
        $fairstream | Should Not Match 'after four skips'

        $popupHeader | Should Match 'DebugBuildFileOperationsGraphFairnessHistorySnapshot'
        $popup | Should Match 'DebugBuildFileOperationsGraphFairnessHistorySnapshot'
        $popup | Should Match 'PopulateGraphHueDebugSummary'
        $fairstream | Should Match 'shape=4-equal-streams-deterministic-history'
        $fairstream | Should Not Match 'ShowWindow\(popup,\s*SW_SHOWNOACTIVATE\)'
        $fairstream | Should Not Match 'shape=4-equal-streams-live-history'

        $selfTest | Should Match 'TryReadPerfMetricSamples'
        $selfTest | Should Match 'fairstreamSaturationDispatchMetricBaseline'
        $selfTest | Should Match 'fairstreamSaturationMaxDistinctInFlightTrees'
        $fairstream | Should Match 'FileOps\.CopyRecursiveParallel\.WorkItemDispatches'
        $fairstream | Should Match 'fairstreamSaturationDispatchMetricBaseline'
        $fairstream | Should Match 'fairstreamSaturationMaxDistinctInFlightTrees'
        $fairstream | Should Match 'Fairstream_SaturationConcurrentCopiesMakeProgress[\s\S]*copyMoveMaxConcurrency":4'
        $fairstream | Should Match 'distinct source trees'
        $fairstream | Should Match 'robust dispatch'
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

    It 'keeps MainMenuBarHost HWND teardown inside the DxUi host lifetime' {
        $source = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'
        $classMatch = [regex]::Match($source, 'class\s+MainMenuBarHost\s+final[\s\S]*?\n\};')
        $classMatch.Success | Should Be $true
        $body = $classMatch.Value

        $body | Should Match '~MainMenuBarHost\(\)\s+noexcept[\s\S]{0,80}Destroy\(\);'
        $body | Should Match 'void\s+OnNcDestroy\(HWND\s+hwnd\)\s+noexcept[\s\S]*?_hwnd\.release\(\)'
        $body | Should Match 'SetWindowLongPtrW\(hwnd,\s*GWLP_USERDATA,\s*0\)'

        $hostIndex = $body.IndexOf('RedSalamander::DxUi::WindowHost _host;')
        $hwndIndex = $body.IndexOf('wil::unique_hwnd _hwnd;')
        ($hostIndex -ge 0) | Should Be $true
        ($hwndIndex -ge 0) | Should Be $true
        ($hostIndex -lt $hwndIndex) | Should Be $true
    }
}
