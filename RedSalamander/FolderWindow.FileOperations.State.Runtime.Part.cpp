bool IsAutoDismissableFileOperationCompletion(HRESULT resultHr, unsigned long warningCount, unsigned long errorCount) noexcept
{
    if (IsCancellationStatus(resultHr))
    {
        return true;
    }

    if (FAILED(resultHr))
    {
        return false;
    }

    return warningCount == 0 && errorCount == 0;
}

FolderWindow::FileOperationState::FileOperationState(FolderWindow& owner) : _owner(owner)
{
    _uiLifetime = std::make_shared<int>(0);
}

FolderWindow::FileOperationState::~FileOperationState()
{
    Shutdown();
}

HRESULT FolderWindow::FileOperationState::StartOperation(FileSystemOperation operation,
                                                         FolderWindow::Pane sourcePane,
                                                         std::optional<FolderWindow::Pane> destinationPane,
                                                         const wil::com_ptr<IFileSystem>& fileSystem,
                                                         std::vector<std::filesystem::path> sourcePaths,
                                                         std::filesystem::path destinationFolder,
                                                         FileSystemFlags flags,
                                                         bool waitForOthers,
                                                         uint64_t initialSpeedLimitBytesPerSecond,
                                                         ExecutionMode executionMode,
                                                         bool requireConfirmation,
                                                         wil::com_ptr<IFileSystem> destinationFileSystem)
{
    if (! fileSystem)
    {
        Debug::Error(L"FolderWindow StartOperation null filesystem");
        return E_POINTER;
    }

    if (sourcePaths.empty())
    {
        Debug::Error(L"FolderWindow StartOperation sourcePath empty");
        return S_FALSE;
    }

    if (! destinationFileSystem && ! CanSameFileSystemOperation(fileSystem, operation))
    {
        Debug::Error(L"FolderWindow StartOperation provider rejected same-filesystem operation op={}", static_cast<unsigned int>(operation));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const std::wstring& sourcePluginId      = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.pluginId : _owner._rightPane.pluginId;
    const std::wstring& sourcePluginShortId = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.pluginShortId : _owner._rightPane.pluginShortId;

    if (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE)
    {
        std::wstring_view destinationPluginId;
        if (destinationPane.has_value())
        {
            destinationPluginId = destinationPane.value() == FolderWindow::Pane::Left ? std::wstring_view(_owner._leftPane.pluginId)
                                                                                      : std::wstring_view(_owner._rightPane.pluginId);
        }

        SessionState::UpdateActiveFileSystemPluginIdsAndOperation({sourcePluginId, destinationPluginId}, SessionState::OperationKind::Copy);
    }

    const bool sourceIsLocalFilePlugin  = NavigationLocation::IsFilePluginShortId(sourcePluginShortId);
    const bool deleteBypassesRecycleBin = operation == FILESYSTEM_DELETE && ((flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) == 0 || ! sourceIsLocalFilePlugin);

    const bool allowPreCalcForOperation  = operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE ||
                                           // For Recycle Bin deletes, the shell can provide progress without blocking on a full recursive preflight scan.
                                           deleteBypassesRecycleBin;
    const unsigned int preCalcMaxWorkers = GetPreCalcMaxWorkersFromSettings(_owner._settings);
    const bool enablePreCalc             = allowPreCalcForOperation && GetPreCalcEnabledFromSettings(_owner._settings);
    const bool supportsBandwidthLimit    = operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE;
    const uint64_t taskDesiredSpeedLimit = (supportsBandwidthLimit && initialSpeedLimitBytesPerSecond == 0)
                                               ? GetDefaultBandwidthLimitBytesPerSecondFromSettings(_owner._settings)
                                               : initialSpeedLimitBytesPerSecond;

    std::vector<DWORD> sourcePathAttributesHint;

    if (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE)
    {
        unsigned long long fileCount    = 0;
        unsigned long long folderCount  = 0;
        unsigned long long unknownCount = 0;
        std::filesystem::path sampleFile;
        bool hasSampleFile = false;

        const FolderView& sourceFolderView = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.folderView : _owner._rightPane.folderView;
        const std::vector<FolderView::PathAttributes> selected = sourceFolderView.GetSelectedOrFocusedPathAttributes();
        bool selectionMatches                                  = ! selected.empty() && selected.size() == sourcePaths.size();
        if (selectionMatches)
        {
            for (size_t i = 0; i < selected.size(); ++i)
            {
                if (selected[i].path != sourcePaths[i])
                {
                    selectionMatches = false;
                    break;
                }
            }
        }

        if (selectionMatches)
        {
            for (const auto& item : selected)
            {
                const bool isDirectory = (item.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (isDirectory)
                {
                    ++folderCount;
                    continue;
                }

                ++fileCount;
                if (! hasSampleFile)
                {
                    sampleFile    = item.path;
                    hasSampleFile = true;
                }
            }

            sourcePathAttributesHint.reserve(selected.size());
            for (const auto& item : selected)
            {
                sourcePathAttributesHint.push_back(item.fileAttributes);
            }
        }
        else
        {
            unknownCount = static_cast<unsigned long long>(sourcePaths.size());
        }

        auto suffixFor = [](unsigned long long count) noexcept -> std::wstring_view
        { return count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s"); };

        const unsigned long long itemCount = static_cast<unsigned long long>(sourcePaths.size());
        std::wstring what;
        if (unknownCount > 0)
        {
            const std::wstring_view itemSuffix = suffixFor(itemCount);
            what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_ITEM, itemCount, itemSuffix);
        }
        else if (fileCount > 0 && folderCount > 0)
        {
            const std::wstring_view fileSuffix   = suffixFor(fileCount);
            const std::wstring_view folderSuffix = suffixFor(folderCount);
            what = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILES_FOLDERS, fileCount, fileSuffix, folderCount, folderSuffix);
        }
        else if (fileCount > 0)
        {
            const std::wstring_view fileSuffix = suffixFor(fileCount);
            what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILE, fileCount, fileSuffix);
        }
        else
        {
            const std::wstring_view folderSuffix = suffixFor(folderCount);
            what                                 = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FOLDER, folderCount, folderSuffix);
        }

        auto ensureTrailingSeparator = [](std::wstring text) noexcept -> std::wstring
        {
            if (text.empty())
            {
                return text;
            }

            const wchar_t last = text.back();
            if (last == L'\\' || last == L'/')
            {
                return text;
            }

            text.push_back(L'\\');
            return text;
        };

        auto normalizeSlashes = [](std::wstring& text) noexcept
        {
            for (auto& ch : text)
            {
                if (ch == L'/')
                {
                    ch = L'\\';
                }
            }
        };

        std::wstring fromText;
        if (sourcePaths.size() == 1u)
        {
            fromText = sourcePaths.front().wstring();
            if (unknownCount == 0 && folderCount == 1ull && fileCount == 0ull)
            {
                fromText = ensureTrailingSeparator(std::move(fromText));
            }
        }
        else
        {
            std::filesystem::path commonParent = sourcePaths.front().parent_path();
            bool multipleParents               = false;
            for (size_t index = 1; index < sourcePaths.size(); ++index)
            {
                const std::filesystem::path parent = sourcePaths[index].parent_path();
                if (CompareStringOrdinal(commonParent.c_str(), -1, parent.c_str(), -1, TRUE) != CSTR_EQUAL)
                {
                    multipleParents = true;
                    break;
                }
            }

            if (multipleParents)
            {
                fromText = LoadStringResource(nullptr, IDS_FILEOPS_LOCATION_MULTIPLE);
            }
            else if (unknownCount == 0 && fileCount > 0 && folderCount > 0 && hasSampleFile)
            {
                fromText = sampleFile.wstring();
            }
            else
            {
                fromText = ensureTrailingSeparator(commonParent.wstring());
            }
        }

        std::wstring toText;
        toText = ensureTrailingSeparator(destinationFolder.wstring());
        normalizeSlashes(fromText);
        normalizeSlashes(toText);

        const UINT messageId = operation == FILESYSTEM_COPY ? static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_COPY) : static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_MOVE);
        const std::wstring message = FormatStringResource(nullptr, messageId, what, fromText, toText);

        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_CONFIRM);

        HostPromptRequest prompt{};
        prompt.version       = 1;
        prompt.sizeBytes     = sizeof(prompt);
        prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
        prompt.severity      = HOST_ALERT_INFO;
        prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
        prompt.targetWindow  = _owner.GetHwnd();
        prompt.title         = caption.c_str();
        prompt.message       = message.c_str();
        prompt.defaultResult = HOST_PROMPT_RESULT_OK;

        HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
        const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
        if (FAILED(hrPrompt) || promptResult != HOST_PROMPT_RESULT_OK)
        {
            return S_FALSE;
        }

        const bool isRecursive = (flags & FILESYSTEM_FLAG_RECURSIVE) != 0;
        if (isRecursive && _owner._settings && GetReparsePointPolicyFromSettings(*_owner._settings, sourcePluginId) == ReparsePointPolicy::FollowTargets)
        {
            bool shouldPrompt = false;
            {
                std::scoped_lock lock(_followTargetsWarningMutex);
                if (_followTargetsWarningAccepted)
                {
                    shouldPrompt = false;
                }
                else if (_followTargetsWarningPromptActive)
                {
                    // Safety-first: if a warning prompt is already visible (possible re-entrancy), abort this start.
                    return S_FALSE;
                }
                else
                {
                    _followTargetsWarningPromptActive = true;
                    shouldPrompt                      = true;
                }
            }

            if (shouldPrompt)
            {
                const std::wstring warningCaption = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
                const std::wstring warningMessage = LoadStringResource(nullptr, IDS_MSG_FILEOPS_REPARSE_FOLLOW_WARNING);

                HostPromptRequest warningPrompt{};
                warningPrompt.version       = 1;
                warningPrompt.sizeBytes     = sizeof(warningPrompt);
                warningPrompt.scope         = HOST_ALERT_SCOPE_WINDOW;
                warningPrompt.severity      = HOST_ALERT_WARNING;
                warningPrompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
                warningPrompt.targetWindow  = _owner.GetHwnd();
                warningPrompt.title         = warningCaption.c_str();
                warningPrompt.message       = warningMessage.c_str();
                warningPrompt.defaultResult = HOST_PROMPT_RESULT_CANCEL;

                HostPromptResult warningResult = HOST_PROMPT_RESULT_NONE;
                const HRESULT hrWarning        = HostShowPrompt(warningPrompt, nullptr, &warningResult);
                if (FAILED(hrWarning) || warningResult != HOST_PROMPT_RESULT_OK)
                {
                    std::scoped_lock lock(_followTargetsWarningMutex);
                    _followTargetsWarningPromptActive = false;
                    return S_FALSE;
                }

                std::scoped_lock lock(_followTargetsWarningMutex);
                _followTargetsWarningPromptActive = false;
                _followTargetsWarningAccepted     = true;
            }
        }
    }
    else if (operation == FILESYSTEM_DELETE && (requireConfirmation || deleteBypassesRecycleBin))
    {
        unsigned long long fileCount    = 0;
        unsigned long long folderCount  = 0;
        unsigned long long unknownCount = 0;
        std::filesystem::path sampleFile;
        bool hasSampleFile = false;

        const FolderView& sourceFolderView = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.folderView : _owner._rightPane.folderView;
        const std::vector<FolderView::PathAttributes> selected = sourceFolderView.GetSelectedOrFocusedPathAttributes();
        bool selectionMatches                                  = ! selected.empty() && selected.size() == sourcePaths.size();
        if (selectionMatches)
        {
            for (size_t i = 0; i < selected.size(); ++i)
            {
                if (selected[i].path != sourcePaths[i])
                {
                    selectionMatches = false;
                    break;
                }
            }
        }

        if (selectionMatches)
        {
            for (const auto& item : selected)
            {
                const bool isDirectory = (item.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (isDirectory)
                {
                    ++folderCount;
                    continue;
                }

                ++fileCount;
                if (! hasSampleFile)
                {
                    sampleFile    = item.path;
                    hasSampleFile = true;
                }
            }
        }
        else
        {
            unknownCount = static_cast<unsigned long long>(sourcePaths.size());
        }

        auto suffixFor = [](unsigned long long count) noexcept -> std::wstring_view
        { return count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s"); };

        const unsigned long long itemCount = static_cast<unsigned long long>(sourcePaths.size());
        std::wstring what;
        if (unknownCount > 0)
        {
            const std::wstring_view itemSuffix = suffixFor(itemCount);
            what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_ITEM, itemCount, itemSuffix);
        }
        else if (fileCount > 0 && folderCount > 0)
        {
            const std::wstring_view fileSuffix   = suffixFor(fileCount);
            const std::wstring_view folderSuffix = suffixFor(folderCount);
            what = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILES_FOLDERS, fileCount, fileSuffix, folderCount, folderSuffix);
        }
        else if (fileCount > 0)
        {
            const std::wstring_view fileSuffix = suffixFor(fileCount);
            what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILE, fileCount, fileSuffix);
        }
        else
        {
            const std::wstring_view folderSuffix = suffixFor(folderCount);
            what                                 = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FOLDER, folderCount, folderSuffix);
        }

        auto ensureTrailingSeparator = [](std::wstring text) noexcept -> std::wstring
        {
            if (text.empty())
            {
                return text;
            }

            const wchar_t last = text.back();
            if (last == L'\\' || last == L'/')
            {
                return text;
            }

            text.push_back(L'\\');
            return text;
        };

        auto normalizeSlashes = [](std::wstring& text) noexcept
        {
            for (auto& ch : text)
            {
                if (ch == L'/')
                {
                    ch = L'\\';
                }
            }
        };

        std::wstring fromText;
        if (sourcePaths.size() == 1u)
        {
            fromText = sourcePaths.front().wstring();
            if (unknownCount == 0 && folderCount == 1ull && fileCount == 0ull)
            {
                fromText = ensureTrailingSeparator(std::move(fromText));
            }
        }
        else
        {
            std::filesystem::path commonParent = sourcePaths.front().parent_path();
            bool multipleParents               = false;
            for (size_t index = 1; index < sourcePaths.size(); ++index)
            {
                const std::filesystem::path parent = sourcePaths[index].parent_path();
                if (CompareStringOrdinal(commonParent.c_str(), -1, parent.c_str(), -1, TRUE) != CSTR_EQUAL)
                {
                    multipleParents = true;
                    break;
                }
            }

            if (multipleParents)
            {
                fromText = LoadStringResource(nullptr, IDS_FILEOPS_LOCATION_MULTIPLE);
            }
            else if (unknownCount == 0 && fileCount > 0 && folderCount > 0 && hasSampleFile)
            {
                fromText = sampleFile.wstring();
            }
            else
            {
                fromText = ensureTrailingSeparator(commonParent.wstring());
            }
        }

        normalizeSlashes(fromText);

        const std::wstring message = FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFIRM_PERMANENT_DELETE, what, fromText);
        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_CONFIRM);

        HostPromptRequest prompt{};
        prompt.version       = 1;
        prompt.sizeBytes     = sizeof(prompt);
        prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
        prompt.severity      = HOST_ALERT_WARNING;
        prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
        prompt.targetWindow  = _owner.GetHwnd();
        prompt.title         = caption.c_str();
        prompt.message       = message.c_str();
        prompt.defaultResult = HOST_PROMPT_RESULT_CANCEL;

        HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
        const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
        if (FAILED(hrPrompt) || promptResult != HOST_PROMPT_RESULT_OK)
        {
            return S_FALSE;
        }
    }

    if (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE)
    {
        auto normalizeSlashes = [](std::wstring& text) noexcept
        {
            for (auto& ch : text)
            {
                if (ch == L'/')
                {
                    ch = L'\\';
                }
            }
        };

        const std::wstring destinationFolderText = destinationFolder.wstring();

        const bool haveAttributesHint = sourcePathAttributesHint.size() == sourcePaths.size();

        std::wstring invalidSourceText;
        std::wstring invalidDestinationItemText;
        for (size_t index = 0; index < sourcePaths.size(); ++index)
        {
            const bool hintIsDirectory = haveAttributesHint && ((sourcePathAttributesHint[index] & FILE_ATTRIBUTE_DIRECTORY) != 0);
            // If we have hints, only directories can cause "copy into self/descendant" recursion.
            // If we don't have hints, be conservative and validate all sources.
            if (haveAttributesHint && ! hintIsDirectory)
            {
                continue;
            }

            const std::wstring sourceText = sourcePaths[index].wstring();
            const std::wstring_view leaf  = GetPathLeaf(sourceText);
            if (leaf.empty())
            {
                continue;
            }

            const std::wstring destinationItemText = JoinFolderAndLeaf(destinationFolderText, leaf);

            std::wstring sourceNormalized          = sourceText;
            std::wstring destinationItemNormalized = destinationItemText;
            normalizeSlashes(sourceNormalized);
            normalizeSlashes(destinationItemNormalized);

            if (! IsSameOrChildPath(sourceNormalized, destinationItemNormalized))
            {
                continue;
            }

            invalidSourceText          = sourceText;
            invalidDestinationItemText = destinationItemText;
            break;
        }

        if (! invalidSourceText.empty())
        {
            Debug::Error(L"FolderWindow StartOperation rejected overlapping destination op={} src:{} dstFolder:{} dstItem:{}",
                         OperationToString(operation),
                         invalidSourceText,
                         destinationFolder.native(),
                         invalidDestinationItemText);

            const std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
            const std::wstring message =
                FormatStringResource(nullptr, IDS_FMT_FILEOPS_INVALID_DESTINATION_OVERLAP, invalidSourceText, destinationFolder.native());
            FolderView& view = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.folderView : _owner._rightPane.folderView;
            view.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, title, message);
            return S_FALSE;
        }
    }

    auto task = std::make_unique<Task>(*this);
    {
        std::scoped_lock lock(_mutex);
        task->_taskId                   = _nextTaskId++;
        task->_operation                = operation;
        task->_executionMode            = executionMode;
        task->_sourcePane               = sourcePane;
        task->_destinationPane          = destinationPane;
        task->_fileSystem               = fileSystem;
        task->_destinationFileSystem    = std::move(destinationFileSystem);
        task->_sourcePaths              = std::move(sourcePaths);
        task->_sourcePathAttributesHint = std::move(sourcePathAttributesHint);
        task->_destinationFolder        = std::move(destinationFolder);
        task->_flags                    = flags;
        task->_enablePreCalc            = enablePreCalc;
        task->_preCalcMaxWorkers        = preCalcMaxWorkers;
        task->_crossFsBridgeBufferBytes = GetCrossFsBridgeBufferBytesFromSettings(_owner._settings);
        task->_waitForOthers.store(waitForOthers, std::memory_order_release);
        task->_desiredSpeedLimitBytesPerSecond.store(taskDesiredSpeedLimit, std::memory_order_release);
        // Mark as waiting in queue immediately if queuing, so UI shows "Waiting..." right away
        task->SetWaitingInQueue(waitForOthers);
    }

    {
        const size_t itemCount = task->_sourcePaths.size();
        {
            std::scoped_lock lock(task->_topLevelCompletionMutex);
            task->_topLevelItemKinds.assign(itemCount, Task::TopLevelItemKind::Unknown);
            task->_topLevelItemCompleted.assign(itemCount, 0);
            task->_plannedTopLevelFiles     = 0;
            task->_plannedTopLevelFolders   = 0;
            task->_completedTopLevelFiles   = 0;
            task->_completedTopLevelFolders = 0;
        }
        task->_publishedCompletedTopLevelFiles.store(0, std::memory_order_release);
        task->_publishedCompletedTopLevelFolders.store(0, std::memory_order_release);

        const bool haveHint = task->_sourcePathAttributesHint.size() == itemCount;
        if (haveHint)
        {
            std::scoped_lock lock(task->_topLevelCompletionMutex);
            for (size_t i = 0; i < itemCount; ++i)
            {
                const bool isDir = (task->_sourcePathAttributesHint[i] & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (isDir)
                {
                    task->_topLevelItemKinds[i] = Task::TopLevelItemKind::Folder;
                    if (task->_plannedTopLevelFolders < std::numeric_limits<unsigned long>::max())
                    {
                        ++task->_plannedTopLevelFolders;
                    }
                }
                else
                {
                    task->_topLevelItemKinds[i] = Task::TopLevelItemKind::File;
                    if (task->_plannedTopLevelFiles < std::numeric_limits<unsigned long>::max())
                    {
                        ++task->_plannedTopLevelFiles;
                    }
                }
            }
        }
    }

    {
        std::scoped_lock lock(task->_progressPathMutex);
        if (! task->_sourcePaths.empty())
        {
            task->_progressSourcePath = task->_sourcePaths.front().native();
        }

        if (! task->_destinationFolder.empty())
        {
            task->_progressDestinationPath = task->_destinationFolder.native();
        }

        PublishDiagnosticPathSnapshotLocked(*task);
    }

    Task* rawTask = task.get();

    {
        std::scoped_lock lock(_mutex);
        _tasks.emplace_back(std::move(task));
    }

    EnsurePopupVisible();

    rawTask->_thread = std::jthread([rawTask](std::stop_token stopToken) noexcept { rawTask->ThreadMain(stopToken); });
    return S_OK;
}

void FolderWindow::FileOperationState::ApplyTheme(const AppTheme& /*theme*/)
{
    HWND popup      = nullptr;
    HWND issuesPane = nullptr;
    {
        std::scoped_lock lock(_mutex);
        popup      = _popup.get();
        issuesPane = _issuesPane.get();
    }

    if (popup)
    {
        PostMessageW(popup, WM_THEMECHANGED, 0, 0);
    }

    if (issuesPane)
    {
        PostMessageW(issuesPane, WM_THEMECHANGED, 0, 0);
    }
}

void FolderWindow::FileOperationState::Shutdown() noexcept
{
    std::vector<std::unique_ptr<Task>> tasks;
    wil::unique_hwnd popupToClose;
    wil::unique_hwnd issuesPaneToClose;
    HWND popupHwnd      = nullptr;
    HWND issuesPaneHwnd = nullptr;
    {
        std::scoped_lock lock(_mutex);
        _uiLifetime.reset();
        tasks.swap(_tasks);
        popupHwnd         = _popup.get();
        issuesPaneHwnd    = _issuesPane.get();
        popupToClose      = std::move(_popup);
        issuesPaneToClose = std::move(_issuesPane);
    }

    if (_owner._settings)
    {
        if (popupHwnd)
        {
            SavePopupPlacement(popupHwnd);
        }

        if (issuesPaneHwnd)
        {
            SaveIssuesPanePlacement(issuesPaneHwnd);
        }

        if (popupHwnd || issuesPaneHwnd)
        {
            const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(kFileOpsAppId, *_owner._settings);
            if (FAILED(saveHr))
            {
                const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kFileOpsAppId);
                Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
            }
        }
    }

    for (auto& task : tasks)
    {
        if (task)
        {
            task->RequestCancel();
        }
    }

    // tasks destruct here; jthread joins automatically.
    FlushDiagnostics(true);
}

void FolderWindow::FileOperationState::NotifyQueueChanged()
{
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Queue.NotifyAllCount", L"", 0, 1u, 0u, S_OK);
    }
    _queueCv.notify_all();
}

bool FolderWindow::FileOperationState::HasActiveOperations() noexcept
{
    {
        std::scoped_lock lock(_mutex);
        if (! _tasks.empty())
        {
            return true;
        }
    }

    // Defensive fallback: active operations are expected to always have a task object.
    std::scoped_lock lock(_queueMutex);
    return _activeOperations > 0 || ! _queue.empty();
}

bool FolderWindow::FileOperationState::ShouldQueueNewTask() noexcept
{
    if (! _queueNewTasks.load(std::memory_order_acquire))
    {
        return false;
    }

    return HasActiveOperations();
}

void FolderWindow::FileOperationState::SetQueueNewTasks(bool queue) noexcept
{
    _queueNewTasks.store(queue, std::memory_order_release);
}

bool FolderWindow::FileOperationState::GetQueueNewTasks() const noexcept
{
    return _queueNewTasks.load(std::memory_order_acquire);
}

void FolderWindow::FileOperationState::ApplyQueueMode(bool queue) noexcept
{
    _queueNewTasks.store(queue, std::memory_order_release);

    std::vector<Task*> tasks;
    CollectTasks(tasks);

    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        if (! queue)
        {
            task->SetWaitForOthers(false);
            continue;
        }

        if (! task->HasStarted())
        {
            task->SetWaitForOthers(true);
            continue;
        }
    }

    UpdateQueuePausedTasks();
    NotifyQueueChanged();
}

void FolderWindow::FileOperationState::CancelAll() noexcept
{
    std::vector<Task*> tasks;
    {
        std::scoped_lock lock(_mutex);
        tasks.reserve(_tasks.size());
        for (const auto& task : _tasks)
        {
            if (task)
            {
                tasks.push_back(task.get());
            }
        }
    }

    for (Task* task : tasks)
    {
        if (task)
        {
            task->RequestCancel();
        }
    }
}

void FolderWindow::FileOperationState::CollectTasks(std::vector<Task*>& outTasks) noexcept
{
    std::scoped_lock lock(_mutex);
    outTasks.clear();
    outTasks.reserve(_tasks.size());
    for (const auto& task : _tasks)
    {
        if (task)
        {
            outTasks.push_back(task.get());
        }
    }
}

void FolderWindow::FileOperationState::CollectInformationalTasks(std::vector<FolderWindow::InformationalTaskUpdate>& outTasks) noexcept
{
    const uint64_t waitStartedUs = PerfNowUs();
    std::scoped_lock lock(_mutex);
    const uint64_t waitUs        = PerfElapsedUs(waitStartedUs);
    const uint64_t copyStartedUs = PerfNowUs();
    outTasks.clear();
    outTasks.reserve(_informationalTasks.size());
    for (const auto& task : _informationalTasks)
    {
        outTasks.push_back(task);
    }
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.InfoTask.Collect.LockWaitUs", L"", waitUs, static_cast<uint64_t>(_informationalTasks.size()), 0u, S_OK);
        Debug::Perf::Emit(L"FileOps.InfoTask.Collect.CopyUs", L"", PerfElapsedUs(copyStartedUs), static_cast<uint64_t>(_informationalTasks.size()), 0u, S_OK);
    }
}

void FolderWindow::FileOperationState::CollectCompletedTasks(std::vector<CompletedTaskSummary>& outTasks) noexcept
{
    std::scoped_lock lock(_mutex);
    outTasks.clear();
    outTasks.reserve(_completedTasks.size());
    for (const auto& summary : _completedTasks)
    {
        outTasks.push_back(summary);
    }
}

void FolderWindow::FileOperationState::CollectDiagnostics(std::vector<TaskDiagnosticEntry>& outEntries) noexcept
{
    std::scoped_lock lock(_diagnosticsMutex);
    outEntries.clear();
    outEntries.reserve(_diagnosticsInMemory.size());
    for (const auto& entry : _diagnosticsInMemory)
    {
        outEntries.push_back(entry);
    }
}

void FolderWindow::FileOperationState::DismissCompletedTask(uint64_t taskId) noexcept
{
    wil::unique_hwnd popupToClose;

    {
        std::scoped_lock lock(_mutex);
        _completedTasks.erase(std::remove_if(_completedTasks.begin(),
                                             _completedTasks.end(),
                                             [&](const CompletedTaskSummary& summary) noexcept { return summary.taskId == taskId; }),
                              _completedTasks.end());

        if (_tasks.empty() && _completedTasks.empty() && _informationalTasks.empty())
        {
            popupToClose = std::move(_popup);
        }
    }
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::DebugAppendCompletedTaskForSelfTest(CompletedTaskSummary summary) noexcept
{
    std::scoped_lock lock(_mutex);
    _completedTasks.push_back(std::move(summary));
}
#endif

uint64_t FolderWindow::FileOperationState::CreateOrUpdateInformationalTask(const FolderWindow::InformationalTaskUpdate& update) noexcept
{
    bool createdNew    = false;
    bool autoDismissed = false;
    bool needShowPopup = false;
    HWND popup         = nullptr;
    wil::unique_hwnd popupToClose;

    uint64_t taskId      = update.taskId;
    bool updatedExisting = false;

    const uint64_t lockWaitStartedUs = PerfNowUs();
    {
        std::scoped_lock lock(_mutex);
        const uint64_t lockWaitUs  = PerfElapsedUs(lockWaitStartedUs);
        const uint64_t lockHeldUs0 = PerfNowUs();
        if (taskId != 0)
        {
            for (auto& existing : _informationalTasks)
            {
                if (existing.taskId == taskId)
                {
                    existing        = update;
                    existing.taskId = taskId;
                    updatedExisting = true;
                    break;
                }
            }
        }

        if (taskId == 0 || ! updatedExisting)
        {
            taskId                                       = _nextTaskId++;
            FolderWindow::InformationalTaskUpdate stored = update;
            stored.taskId                                = taskId;
            _informationalTasks.push_back(std::move(stored));

            createdNew = true;
        }

        const bool autoDismissSuccess = _owner._settings ? GetAutoDismissSuccessFromSettings(*_owner._settings) : false;
        if (autoDismissSuccess && update.finished && IsAutoDismissableFileOperationCompletion(update.resultHr, 0, 0))
        {
            _informationalTasks.erase(std::remove_if(_informationalTasks.begin(),
                                                     _informationalTasks.end(),
                                                     [&](const FolderWindow::InformationalTaskUpdate& task) noexcept { return task.taskId == taskId; }),
                                      _informationalTasks.end());

            autoDismissed = true;
            if (_tasks.empty() && _completedTasks.empty() && _informationalTasks.empty())
            {
                popupToClose = std::move(_popup);
            }
        }

        popup = _popup.get();
        if (! autoDismissed && ! update.finished)
        {
            // Informational tasks should surface progress similarly to file operations: when the task starts, ensure the popup is shown.
            needShowPopup = createdNew || popup == nullptr || IsWindowVisible(popup) == FALSE || IsIconic(popup) != FALSE;
        }
        else
        {
            needShowPopup = createdNew && (popup == nullptr) && ! autoDismissed;
        }

        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.InfoTask.Update.LockWaitUs", L"", lockWaitUs, static_cast<uint64_t>(_informationalTasks.size()), 0u, S_OK);
            Debug::Perf::Emit(
                L"FileOps.InfoTask.Update.LockHoldUs", L"", PerfElapsedUs(lockHeldUs0), static_cast<uint64_t>(_informationalTasks.size()), 0u, S_OK);
        }
    }

    if (needShowPopup)
    {
        EnsurePopupVisible();
    }
    else if (popup)
    {
        InvalidateRect(popup, nullptr, FALSE);
    }

    return taskId;
}

void FolderWindow::FileOperationState::DismissInformationalTask(uint64_t taskId) noexcept
{
    if (taskId == 0)
    {
        return;
    }

    wil::unique_hwnd popupToClose;
    HWND popup = nullptr;
    {
        std::scoped_lock lock(_mutex);
        _informationalTasks.erase(std::remove_if(_informationalTasks.begin(),
                                                 _informationalTasks.end(),
                                                 [&](const FolderWindow::InformationalTaskUpdate& task) noexcept { return task.taskId == taskId; }),
                                  _informationalTasks.end());

        if (_tasks.empty() && _completedTasks.empty() && _informationalTasks.empty())
        {
            popupToClose = std::move(_popup);
        }
        else
        {
            popup = _popup.get();
        }
    }

    if (popup)
    {
        InvalidateRect(popup, nullptr, FALSE);
    }
}

bool FolderWindow::FileOperationState::GetAutoDismissSuccess() const noexcept
{
    if (! _owner._settings)
    {
        return false;
    }

    return GetAutoDismissSuccessFromSettings(*_owner._settings);
}

void FolderWindow::FileOperationState::SetAutoDismissSuccess(bool enabled) noexcept
{
    if (! _owner._settings)
    {
        return;
    }

    const bool previous = GetAutoDismissSuccessFromSettings(*_owner._settings);
    SetAutoDismissSuccessInSettings(*_owner._settings, enabled);

    if (enabled && ! previous)
    {
        wil::unique_hwnd popupToClose;
        HWND popup = nullptr;
        {
            std::scoped_lock lock(_mutex);
            _completedTasks.erase(std::remove_if(_completedTasks.begin(),
                                                 _completedTasks.end(),
                                                 [](const CompletedTaskSummary& summary) noexcept
            { return IsAutoDismissableFileOperationCompletion(summary.resultHr, summary.warningCount, summary.errorCount); }),
                                  _completedTasks.end());

            _informationalTasks.erase(std::remove_if(_informationalTasks.begin(),
                                                     _informationalTasks.end(),
                                                     [](const FolderWindow::InformationalTaskUpdate& task) noexcept
            { return task.finished && IsAutoDismissableFileOperationCompletion(task.resultHr, 0, 0); }),
                                      _informationalTasks.end());

            if (_tasks.empty() && _completedTasks.empty() && _informationalTasks.empty())
            {
                popupToClose = std::move(_popup);
            }
            else
            {
                popup = _popup.get();
            }
        }

        if (popup)
        {
            InvalidateRect(popup, nullptr, FALSE);
        }
    }
}
