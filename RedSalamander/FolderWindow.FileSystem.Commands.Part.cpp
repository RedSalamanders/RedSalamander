namespace
{
[[nodiscard]] std::wstring BuildCreateDirectorySuggestedName(std::wstring_view baseName, unsigned int suffix)
{
    if (suffix == 0u)
    {
        return std::wstring(baseName);
    }

    return std::format(L"{} ({})", baseName, suffix);
}

[[nodiscard]] bool DirectoryNamesMatch(std::wstring_view left, std::wstring_view right, bool ignoreCase) noexcept
{
    if (! ignoreCase)
    {
        return left == right;
    }

    return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool TryCollectExistingDirectoryNames(const wil::com_ptr<IFileSystem>& fileSystem,
                                                    const std::filesystem::path& folder,
                                                    std::vector<std::wstring>& outNames) noexcept
{
    outNames.clear();
    if (! fileSystem)
    {
        return false;
    }

    wil::com_ptr<IFilesInformation> filesInformation;
    const HRESULT hr = fileSystem->ReadDirectoryInfo(folder.c_str(), filesInformation.put());
    if (FAILED(hr) || ! filesInformation)
    {
        return false;
    }

    FileInfo* entry = nullptr;
    if (FAILED(filesInformation->GetBuffer(&entry)) || entry == nullptr)
    {
        return true;
    }

    while (entry != nullptr)
    {
        if ((entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
            const std::wstring_view name(entry->FileName, nameChars);
            if (name != L"." && name != L"..")
            {
                outNames.emplace_back(name);
            }
        }

        if (entry->NextEntryOffset == 0)
        {
            break;
        }

        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset);
    }

    return true;
}

[[nodiscard]] std::wstring ResolveInitialCreateDirectoryName(const wil::com_ptr<IFileSystem>& fileSystem,
                                                             const std::filesystem::path& folder,
                                                             std::wstring_view defaultName,
                                                             bool ignoreCase) noexcept
{
    std::vector<std::wstring> existingDirectoryNames;
    if (! TryCollectExistingDirectoryNames(fileSystem, folder, existingDirectoryNames))
    {
        return std::wstring(defaultName);
    }

    const auto isTaken = [&](std::wstring_view candidate) noexcept
    {
        for (const std::wstring& existingName : existingDirectoryNames)
        {
            if (DirectoryNamesMatch(existingName, candidate, ignoreCase))
            {
                return true;
            }
        }

        return false;
    };

    for (unsigned int suffix = 0u; suffix < 10000u; ++suffix)
    {
        std::wstring candidate = BuildCreateDirectorySuggestedName(defaultName, suffix);
        if (! isTaken(candidate))
        {
            return candidate;
        }
    }

    return std::format(L"{} ({})", defaultName, GetTickCount64());
}

[[nodiscard]] bool TryParseCreateDirectorySuggestedSuffix(std::wstring_view requestedName, std::wstring_view defaultName, unsigned int& outSuffix) noexcept
{
    if (requestedName == defaultName)
    {
        outSuffix = 0u;
        return true;
    }

    std::wstring prefix(defaultName);
    prefix.append(L" (");
    if (requestedName.size() <= prefix.size() || ! requestedName.starts_with(prefix) || requestedName.back() != L')')
    {
        return false;
    }

    const std::wstring_view digits = requestedName.substr(prefix.size(), requestedName.size() - prefix.size() - 1u);
    if (digits.empty())
    {
        return false;
    }

    uint64_t parsedValue = 0u;
    for (const wchar_t ch : digits)
    {
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }

        parsedValue = (parsedValue * 10u) + static_cast<uint64_t>(ch - L'0');
        if (parsedValue > static_cast<uint64_t>((std::numeric_limits<unsigned int>::max)()))
        {
            return false;
        }
    }

    if (parsedValue == 0u)
    {
        return false;
    }

    outSuffix = static_cast<unsigned int>(parsedValue);
    return true;
}
} // namespace

void FolderWindow::CommandCreateDirectory(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.fileSystem)
    {
        return;
    }

    HWND ownerWindow = GetOwnerWindowOrSelf(_hWnd.get());
    std::wstring pluginName;
    if (ownerWindow)
    {
        FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
        const auto& plugins                    = pluginManager.GetPlugins();
        pluginName                             = TryGetFileSystemPluginDisplayName(plugins, state.pluginId, state.pluginShortId);
    }

    const auto folder = state.folderView.GetFolderPath();
    if (! folder)
    {
        return;
    }

    const std::filesystem::path base = folder.value();

    wil::com_ptr<IFileSystemDirectoryOperations> dirOps;
    state.fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryOperations), dirOps.put_void());

    const bool canUseWin32 = IsFilePluginShortId(state.pluginShortId) && LooksLikeWindowsAbsolutePath(base.wstring());
    if (! dirOps && ! canUseWin32)
    {
        std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message;
        if (! pluginName.empty())
        {
            message = FormatStringResource(nullptr, IDS_FMT_PANE_CREATE_DIR_UNSUPPORTED_PLUGIN, pluginName);
        }
        if (message.empty())
        {
            message = LoadStringResource(nullptr, IDS_MSG_PANE_CREATE_DIR_UNSUPPORTED);
        }

        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
        return;
    }

    const std::wstring defaultNameBase = LoadStringResource(nullptr, IDS_NEW_FOLDER_DEFAULT_NAME);
    if (defaultNameBase.empty())
    {
        return;
    }

    const bool treatSuggestedNamesCaseInsensitive = IsFilePluginShortId(state.pluginShortId);
    const std::wstring initialName = ResolveInitialCreateDirectoryName(state.fileSystem, base, defaultNameBase, treatSuggestedNamesCaseInsensitive);

    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::filesystem::path displayPath = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, base);
    const auto folderName                   = PromptForCreateDirectoryName(ownerWindow, displayPath.wstring(), initialName, _theme);
    if (! folderName.has_value())
    {
        return;
    }

    const std::wstring requestedName = folderName.value();
    unsigned int suggestedSuffix     = 0u;
    const bool autoSuffix            = TryParseCreateDirectorySuggestedSuffix(requestedName, defaultNameBase, suggestedSuffix);

    const int maxAttempts = autoSuffix ? 1000 : 1;
    for (int attempt = 0; attempt < maxAttempts; ++attempt)
    {
        std::wstring candidateName = requestedName;
        if (autoSuffix)
        {
            candidateName = BuildCreateDirectorySuggestedName(defaultNameBase, suggestedSuffix + static_cast<unsigned int>(attempt));
        }

        const std::filesystem::path newFolderPath = base / std::filesystem::path(candidateName);
        if (newFolderPath.empty())
        {
            continue;
        }

        HRESULT hr = S_OK;
        if (dirOps)
        {
            hr = dirOps->CreateDirectory(newFolderPath.c_str());
        }
        else
        {
            if (::CreateDirectoryW(newFolderPath.c_str(), nullptr) == 0)
            {
                const DWORD error = GetLastError();
                hr                = HRESULT_FROM_WIN32(error);
            }
        }

        if (SUCCEEDED(hr))
        {
            const std::wstring focusName = newFolderPath.filename().wstring();
            if (! focusName.empty())
            {
                state.folderView.RememberFocusedItemForFolder(base, focusName);
            }

            DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
            cache.NotifyPathCreated(state.fileSystem.get(), newFolderPath);
            state.folderView.ForceRefresh();

            const Pane otherPane   = pane == Pane::Left ? Pane::Right : Pane::Left;
            PaneState& otherState  = otherPane == Pane::Left ? _leftPane : _rightPane;
            const auto otherFolder = otherState.folderView.GetFolderPath();
            if (otherState.fileSystem && otherFolder.has_value() && OrdinalString::EqualsNoCasePath(otherFolder.value(), base) &&
                EqualsNoCase(otherState.pluginId, state.pluginId) && EqualsNoCase(otherState.instanceContext, state.instanceContext) &&
                ! cache.IsFolderWatched(otherState.fileSystem.get(), base))
            {
                otherState.folderView.ForceRefresh();
            }
            return;
        }

        if (hr == E_NOTIMPL)
        {
            std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
            std::wstring message;
            if (! pluginName.empty())
            {
                message = FormatStringResource(nullptr, IDS_FMT_PANE_CREATE_DIR_UNSUPPORTED_PLUGIN, pluginName);
            }
            if (message.empty())
            {
                message = LoadStringResource(nullptr, IDS_MSG_PANE_CREATE_DIR_UNSUPPORTED);
            }

            state.folderView.ShowAlertOverlay(
                FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
            return;
        }

        constexpr HRESULT alreadyExistsHr = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        constexpr HRESULT fileExistsHr    = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        if (autoSuffix && (hr == alreadyExistsHr || hr == fileExistsHr))
        {
            continue;
        }

        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message = FormatStringResource(nullptr, IDS_FMT_PANE_CREATE_DIR_FAILED, newFolderPath.wstring(), static_cast<unsigned long>(hr));
        state.folderView.ShowAlertOverlay(
            FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message), hr);
        return;
    }
}

void FolderWindow::CommandRefresh(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.ForceRefresh();
}

#ifdef ENABLE_TESTS
uint64_t FolderWindow::DebugGetForceRefreshCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetForceRefreshCount();
}

std::wstring_view FolderWindow::DebugGetFocusedItemDisplayName(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetFocusedDisplayName();
}

bool FolderWindow::DebugHasItemDisplayName(Pane pane, std::wstring_view displayName) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugHasItemDisplayName(displayName);
}

size_t FolderWindow::DebugGetItemCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetItemCount();
}

size_t FolderWindow::DebugGetPaneBitmapIconCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetBitmapIconCount();
}

bool FolderWindow::DebugIsItemSelected(Pane pane, std::wstring_view displayName) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugIsItemSelectedByDisplayName(displayName);
}

size_t FolderWindow::DebugGetSelectedCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetSelectedItemCount();
}

uint64_t FolderWindow::DebugGetWarmPaneRenderingCallCount(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetWarmRenderingCallCount();
}

FolderView::DebugWarmPerfSnapshot FolderWindow::DebugGetWarmPanePerfSnapshot(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetWarmPerfSnapshot();
}

bool FolderWindow::DebugWarmPaneRendering(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugWarmRenderingForSelfTest();
}

bool FolderWindow::DebugIsEmptyFolderStateActive(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugIsEmptyFolderStateActive();
}

std::wstring_view FolderWindow::DebugGetEmptyFolderFunMessage(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetEmptyFolderFunMessage();
}

FolderView::DebugEmptyFolderItemMetrics FolderWindow::DebugGetEmptyFolderItemMetrics(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetEmptyFolderItemMetrics();
}

HWND FolderWindow::DebugGetNavigationViewHwnd(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.hNavigationView.get();
}

bool FolderWindow::DebugGetNavigationViewSnapshot(Pane pane, NavigationViewDebugSnapshot& out) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.navigationView.DebugGetSnapshot(out);
}

bool FolderWindow::DebugFocusNavigationViewRegion(Pane pane, NavigationView::FocusRegion region) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.navigationView.DebugFocusRegion(region);
}

bool FolderWindow::DebugFocusItemByDisplayName(Pane pane, std::wstring_view displayName) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.PrepareForExternalCommand(displayName);
}

FolderView::NameFilterState FolderWindow::DebugGetNameFilterState(Pane pane) const
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.GetNameFilterState();
}

bool FolderWindow::DebugIsNameFilterActive(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.IsNameFilterActive();
}

void FolderWindow::DebugResetPaneVisibilityState(Pane pane) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.ShowHiddenNames();
    state.folderView.SetNameFilterState(FolderView::NameFilterState{});
}

FolderView::FilterWatermarkVisualMode FolderWindow::DebugGetFilterWatermarkVisualMode(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.DebugGetFilterWatermarkVisualMode();
}
#endif

void FolderWindow::CommandCalculateDirectorySizes(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const std::optional<std::filesystem::path> folderPath = state.folderView.GetFolderPath();
    if (! folderPath.has_value() || folderPath.value().empty())
    {
        return;
    }

    if (! TryViewSpaceWithViewer(pane, folderPath.value()))
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_PANE_CALCULATE_DIRECTORY_SIZES_FAILED);

        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message));
    }
}

void FolderWindow::CommandSelectionSelectDialog(Pane pane)
{
    SetActivePane(pane);

    std::vector<std::wstring> history;
    if (_settings && _settings->selectionMasks.has_value())
    {
        history = _settings->selectionMasks->selectHistory;
    }
    MaskSyntax::NormalizeWildcardMaskHistory(history, MaskSyntax::kWildcardMaskHistoryMaxItems);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::optional<std::wstring> maskTextOpt =
        PromptForSelectionMask(ownerWindow, history, _theme, IDS_CAPTION_SELECTION_MASK_SELECT, IDS_LABEL_SELECTION_MASK_SELECT);
    if (! maskTextOpt.has_value())
    {
        return;
    }

    const std::wstring& maskText = maskTextOpt.value();

    if (_settings)
    {
        Common::Settings::SelectionMasksSettings& masks =
            _settings->selectionMasks.has_value() ? _settings->selectionMasks.value() : _settings->selectionMasks.emplace();

        MaskSyntax::AddToWildcardMaskHistory(masks.selectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems, maskText);
        MaskSyntax::NormalizeWildcardMaskHistory(masks.selectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems);
    }

    MaskSyntax::WildcardMask mask = MaskSyntax::ParseWildcardMask(maskText);

    SetPaneSelectionByDisplayNamePredicate(pane, [mask = std::move(mask)](std::wstring_view displayName) noexcept {
        return MaskSyntax::MatchesWildcardMask(displayName, mask);
    }, false /* clearExistingSelection */);
}

void FolderWindow::CommandSelectionUnselectDialog(Pane pane)
{
    SetActivePane(pane);

    std::vector<std::wstring> history;
    if (_settings && _settings->selectionMasks.has_value())
    {
        history = _settings->selectionMasks->unselectHistory;
    }
    MaskSyntax::NormalizeWildcardMaskHistory(history, MaskSyntax::kWildcardMaskHistoryMaxItems);

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::optional<std::wstring> maskTextOpt =
        PromptForSelectionMask(ownerWindow, history, _theme, IDS_CAPTION_SELECTION_MASK_UNSELECT, IDS_LABEL_SELECTION_MASK_UNSELECT);
    if (! maskTextOpt.has_value())
    {
        return;
    }

    const std::wstring& maskText = maskTextOpt.value();

    if (_settings)
    {
        Common::Settings::SelectionMasksSettings& masks =
            _settings->selectionMasks.has_value() ? _settings->selectionMasks.value() : _settings->selectionMasks.emplace();

        MaskSyntax::AddToWildcardMaskHistory(masks.unselectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems, maskText);
        MaskSyntax::NormalizeWildcardMaskHistory(masks.unselectHistory, MaskSyntax::kWildcardMaskHistoryMaxItems);
    }

    MaskSyntax::WildcardMask mask = MaskSyntax::ParseWildcardMask(maskText);

    ClearPaneSelectionByDisplayNamePredicate(
        pane, [mask = std::move(mask)](std::wstring_view displayName) noexcept { return MaskSyntax::MatchesWildcardMask(displayName, mask); });
}

void FolderWindow::CommandSelectionInvert(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.InvertSelection();
}

bool FolderWindow::HasSavedSelection() const noexcept
{
    return _savedSelection.has_value() && ! _savedSelection->displayNames.empty();
}

void FolderWindow::CommandSelectionSave(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const std::optional<std::filesystem::path> folderOpt = state.folderView.GetFolderPath();
    if (! folderOpt.has_value() || folderOpt.value().empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    std::vector<std::wstring> names = state.folderView.GetSelectedOrFocusedDisplayNames();
    if (names.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    SavedSelection saved{};
    saved.sourcePluginId        = std::wstring(state.folderView.GetFileSystemPluginId());
    saved.sourceInstanceContext = std::wstring(state.folderView.GetFileSystemInstanceContext());
    saved.sourceFolder          = folderOpt.value();
    saved.displayNames          = std::move(names);
    _savedSelection             = std::move(saved);

    std::wstring clipboardText;
    {
        const std::wstring folderText = folderOpt.value().native();
        size_t reserveChars           = folderText.size() + 2u;
        for (const auto& name : _savedSelection->displayNames)
        {
            reserveChars += name.size() + 2u;
        }

        clipboardText.reserve(reserveChars);
        clipboardText.append(folderText);
        for (const auto& name : _savedSelection->displayNames)
        {
            clipboardText.append(L"\r\n");
            clipboardText.append(name);
        }
        clipboardText.append(L"\r\n");
    }

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    if (! SetClipboardUnicodeText(ownerWindow, clipboardText))
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_SELECTION_SAVE_CLIPBOARD_FAILED);
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
    }
}

void FolderWindow::CopySelectionText(Pane pane, CopySelectionTextMode mode, UINT titleStringId)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    const std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    const bool preferUncPath       = mode == CopySelectionTextMode::UncPathAndName && IsFilePluginShortId(state.pluginShortId);
    const auto renderClipboardLine = [mode, preferUncPath](const std::filesystem::path& path) -> std::wstring
    {
        const std::wstring nativePath = path.native();
        switch (mode)
        {
            case CopySelectionTextMode::PathAndName: return nativePath;

            case CopySelectionTextMode::Name:
            {
                std::wstring fileName = path.filename().native();
                return fileName.empty() ? nativePath : fileName;
            }

            case CopySelectionTextMode::Path:
            {
                std::filesystem::path parentPath = path.parent_path();
                if (parentPath.empty() && path.has_root_path())
                {
                    parentPath = path.root_path();
                }

                std::wstring containingPath = parentPath.native();
                return containingPath.empty() ? nativePath : containingPath;
            }

            case CopySelectionTextMode::UncPathAndName: return preferUncPath ? GetUniversalPathOrOriginal(nativePath) : nativePath;
        }

        return nativePath;
    };

    std::vector<std::wstring> lines;
    lines.reserve(paths.size());

    size_t reserveChars = 2u;
    for (const auto& path : paths)
    {
        std::wstring line = renderClipboardLine(path);
        if (line.empty())
        {
            continue;
        }

        reserveChars += line.size() + 2u;
        lines.push_back(std::move(line));
    }

    if (lines.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    std::wstring clipboardText;
    clipboardText.reserve(reserveChars);
    for (size_t i = 0; i < lines.size(); ++i)
    {
        if (i != 0u)
        {
            clipboardText.append(L"\r\n");
        }
        clipboardText.append(lines[i]);
    }
    clipboardText.append(L"\r\n");

    const HWND ownerWindow = GetClipboardOwnerWindow(_hWnd.get());
    if (! SetClipboardUnicodeText(ownerWindow, clipboardText))
    {
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_SELECTION_SAVE_CLIPBOARD_FAILED);
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
        return;
    }

    std::wstring title = LoadStringResource(nullptr, titleStringId);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_INFORMATION);
    }

    const unsigned long long count = static_cast<unsigned long long>(lines.size());
    const std::wstring_view suffix = count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s");
    std::wstring message           = FormatStringResource(nullptr, IDS_FMT_COPY_PATH_AND_FILE_NAME_COPIED, count, suffix);
    if (message.empty())
    {
        message = std::format(L"Copied {} item{} to clipboard.", count, suffix);
    }

    ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Information, std::move(title), std::move(message), S_OK, false, false);
}

void FolderWindow::CommandCopyPathAndNameAsText(Pane pane)
{
    CopySelectionText(pane, CopySelectionTextMode::PathAndName, IDS_CMD_COPY_PATH_AND_NAME_AS_TEXT);
}

void FolderWindow::CommandCopyNameAsText(Pane pane)
{
    CopySelectionText(pane, CopySelectionTextMode::Name, IDS_CMD_COPY_NAME_AS_TEXT);
}

void FolderWindow::CommandCopyPathAsText(Pane pane)
{
    CopySelectionText(pane, CopySelectionTextMode::Path, IDS_CMD_COPY_PATH_AS_TEXT);
}

void FolderWindow::CommandCopyUncPathAndNameAsText(Pane pane)
{
    CopySelectionText(pane, CopySelectionTextMode::UncPathAndName, IDS_CMD_COPY_PATH_AND_FILE_NAME);
}

void FolderWindow::CommandSelectionRestore(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (! HasSavedSelection())
    {
        MessageBeep(MB_ICONWARNING);
        std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_WARNING);
        std::wstring message = LoadStringResource(nullptr, IDS_MSG_SELECTION_RESTORE_NO_SAVED);
        ShowPaneAlertOverlay(
            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Warning, std::move(title), std::move(message), S_OK, true, false);
        return;
    }

    const SavedSelection& saved = _savedSelection.value();

    std::unordered_set<std::wstring_view> remaining;
    remaining.reserve(saved.displayNames.size());
    for (const auto& name : saved.displayNames)
    {
        if (! name.empty())
        {
            remaining.emplace(name);
        }
    }

    SetPaneSelectionByDisplayNamePredicate(pane,
                                           [&](std::wstring_view displayName) noexcept
    {
        const auto it = remaining.find(displayName);
        if (it == remaining.end())
        {
            return false;
        }
        remaining.erase(it);
        return true;
    },
                                           true /* clearExistingSelection */);

    if (! remaining.empty())
    {
        struct WStringViewNoCaseLess final
        {
            bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
            {
                return OrdinalString::Compare(left, right, true) < 0;
            }
        };

        std::map<std::wstring_view, std::vector<std::wstring_view>, WStringViewNoCaseLess> remainingNoCase;
        for (const auto& name : remaining)
        {
            remainingNoCase[name].push_back(name);
        }

        SetPaneSelectionByDisplayNamePredicate(pane,
                                               [&](std::wstring_view displayName) noexcept
        {
            const auto it = remainingNoCase.find(displayName);
            if (it == remainingNoCase.end() || it->second.empty())
            {
                return false;
            }

            const std::wstring_view matched = it->second.back();
            it->second.pop_back();
            remaining.erase(matched);
            if (it->second.empty())
            {
                remainingNoCase.erase(it);
            }
            return true;
        },
                                               false /* clearExistingSelection */);
    }

    if (remaining.empty())
    {
        return;
    }

    std::wstring missingLines;
    for (const auto& name : saved.displayNames)
    {
        if (name.empty())
        {
            continue;
        }

        if (remaining.find(std::wstring_view(name)) == remaining.end())
        {
            continue;
        }

        missingLines.append(L"- ");
        missingLines.append(name);
        missingLines.append(L"\r\n");
    }

    std::wstring filterNote;
    if (state.folderView.IsNameFilterActive())
    {
        const std::wstring noteText = LoadStringResource(nullptr, IDS_MSG_SELECTION_RESTORE_FILTER_NOTE);
        if (! noteText.empty())
        {
            filterNote.reserve(4u + noteText.size());
            filterNote.append(L"\r\n\r\n");
            filterNote.append(noteText);
        }
    }

    std::wstring title = LoadStringResource(nullptr, IDS_CMD_SELECTION_RESTORE);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_OVERLAY_TITLE_INFORMATION);
    }

    std::wstring message = FormatStringResource(nullptr, IDS_FMT_SELECTION_RESTORE_INCOMPLETE, missingLines, filterNote);
    ShowPaneAlertOverlay(
        pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Information, std::move(title), std::move(message), S_OK, true, false);
}

void FolderWindow::CommandSelectionSelectSameExtension(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SelectSameExtension();
}

void FolderWindow::CommandSelectionSelectSameName(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SelectSameName();
}

void FolderWindow::CommandSelectionUnselectSameExtension(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.UnselectSameExtension();
}

void FolderWindow::CommandSelectionUnselectSameName(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.UnselectSameName();
}

void FolderWindow::CommandSelectionHideSelectedNames(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.HideSelectedNames();
}

void FolderWindow::CommandSelectionHideUnselectedNames(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.HideUnselectedNames();
}

void FolderWindow::CommandSelectionShowHiddenNames(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.ShowHiddenNames();
}

bool FolderWindow::CanShowHiddenNames(Pane pane) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    return state.folderView.HasHiddenNames();
}

void FolderWindow::CommandSelectionGoToPreviousSelectedName(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    static_cast<void>(state.folderView.GoToPreviousSelectedName());
}

void FolderWindow::CommandSelectionGoToNextSelectedName(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    static_cast<void>(state.folderView.GoToNextSelectedName());
}

void FolderWindow::CommandChangeCase(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    if (! state.fileSystem)
    {
        return;
    }

    if (state.changeCaseThread.joinable())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    HWND ownerWindow = _hWnd ? GetAncestor(_hWnd.get(), GA_ROOT) : nullptr;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::optional<ChangeCase::Options> dialogResult = PromptForChangeCase(ownerWindow, _theme, true);
    if (! dialogResult.has_value())
    {
        return;
    }

    const std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        MessageBeep(MB_ICONWARNING);
        return;
    }

    const HWND ownerHwnd                 = _hWnd.get();
    wil::com_ptr<IFileSystem> fileSystem = state.fileSystem;
    const ChangeCase::Options options    = dialogResult.value();
    const std::wstring title             = LoadStringResource(nullptr, IDS_CMD_CHANGE_CASE);

    state.changeCaseThread = std::jthread([ownerHwnd, pane, fileSystem, paths, options, title](std::stop_token stopToken) noexcept
    {
        const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        if (FAILED(coinitHr))
        {
            Debug::Error(L"ChangeCase task: CoInitializeEx(COINIT_MULTITHREADED) failed: 0x{:08X}", coinitHr);
            FAIL_FAST_IF_FAILED(coinitHr);
        }
        [[maybe_unused]] const wil::unique_couninitialize_call coUninit;

        struct ProgressState final
        {
            HWND hwnd               = nullptr;
            FolderWindow::Pane pane = FolderWindow::Pane::Left;
            std::wstring title;
            ULONGLONG startTick      = 0;
            ULONGLONG lastPostedTick = 0;
            uint64_t infoTaskId      = 0;
            ChangeCase::ProgressUpdate last{};

            void PostTaskUpdate(bool finished, HRESULT hr) noexcept
            {
                if (! hwnd || IsWindow(hwnd) == FALSE || infoTaskId == 0)
                {
                    return;
                }

                FolderWindow::InformationalTaskUpdate info{};
                info.kind                       = FolderWindow::InformationalTaskUpdate::Kind::ChangeCase;
                info.taskId                     = infoTaskId;
                info.title                      = title;
                info.changeCaseCurrentPath      = last.currentPath;
                info.changeCaseScannedFolders   = last.scannedFolders;
                info.changeCaseScannedEntries   = last.scannedEntries;
                info.changeCasePlannedRenames   = last.plannedRenames;
                info.changeCaseCompletedRenames = last.completedRenames;
                info.changeCaseEnumerating      = ! finished && last.phase == ChangeCase::ProgressUpdate::Phase::Enumerating;
                info.changeCaseRenaming         = ! finished && last.phase == ChangeCase::ProgressUpdate::Phase::Renaming;
                info.finished                   = finished;
                info.resultHr                   = hr;

                auto payload    = std::make_unique<ChangeCaseTaskPayload>();
                payload->update = std::move(info);
                static_cast<void>(PostMessagePayload(hwnd, WndMsg::kChangeCaseTaskUpdate, 0, std::move(payload)));
            }

            void EnsureTaskVisibleAfterThreshold() noexcept
            {
                if (! hwnd || IsWindow(hwnd) == FALSE || infoTaskId != 0)
                {
                    return;
                }

                const ULONGLONG nowTick = GetTickCount64();
                if (startTick == 0 || nowTick < startTick || (nowTick - startTick) < 700ull)
                {
                    return;
                }

                FolderWindow::InformationalTaskUpdate info{};
                info.kind                       = FolderWindow::InformationalTaskUpdate::Kind::ChangeCase;
                info.title                      = title;
                info.changeCaseCurrentPath      = last.currentPath;
                info.changeCaseScannedFolders   = last.scannedFolders;
                info.changeCaseScannedEntries   = last.scannedEntries;
                info.changeCasePlannedRenames   = last.plannedRenames;
                info.changeCaseCompletedRenames = last.completedRenames;
                info.changeCaseEnumerating      = last.phase == ChangeCase::ProgressUpdate::Phase::Enumerating;
                info.changeCaseRenaming         = last.phase == ChangeCase::ProgressUpdate::Phase::Renaming;

                auto payload    = std::make_unique<ChangeCaseTaskPayload>();
                payload->update = std::move(info);

                ChangeCaseTaskPayload* raw = payload.release();
                DWORD_PTR result           = 0;
                const LRESULT sendOk =
                    SendMessageTimeoutW(hwnd, WndMsg::kChangeCaseTaskUpdate, 0, reinterpret_cast<LPARAM>(raw), SMTO_ABORTIFHUNG | SMTO_BLOCK, 1000, &result);
                if (sendOk == 0)
                {
                    delete raw;
                    return;
                }

                infoTaskId     = static_cast<uint64_t>(result);
                lastPostedTick = nowTick;
            }
        };

        ProgressState progressState{};
        progressState.hwnd      = ownerHwnd;
        progressState.pane      = pane;
        progressState.title     = title;
        progressState.startTick = GetTickCount64();

        const auto onProgress = [](const ChangeCase::ProgressUpdate& update, void* cookie) noexcept
        {
            auto* state = static_cast<ProgressState*>(cookie);
            if (! state)
            {
                return;
            }

            state->last = update;
            state->EnsureTaskVisibleAfterThreshold();

            if (state->infoTaskId != 0)
            {
                const ULONGLONG nowTick = GetTickCount64();
                if (state->lastPostedTick != 0 && nowTick >= state->lastPostedTick && (nowTick - state->lastPostedTick) < 100ull)
                {
                    return;
                }

                state->lastPostedTick = nowTick;
                state->PostTaskUpdate(false, S_OK);
            }
        };

        const HRESULT operationHr = ChangeCase::ApplyToPaths(*fileSystem, paths, options, stopToken, onProgress, &progressState);

        progressState.EnsureTaskVisibleAfterThreshold();
        progressState.PostTaskUpdate(true, operationHr);

        if (ownerHwnd && IsWindow(ownerHwnd) != FALSE)
        {
            auto completed  = std::make_unique<ChangeCaseCompletedPayload>();
            completed->pane = pane;
            completed->hr   = operationHr;
            static_cast<void>(PostMessagePayload(ownerHwnd, WndMsg::kChangeCaseCompleted, 0, std::move(completed)));
        }
    });
}
