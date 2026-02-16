#include "FolderWindow.FileOperationsInternal.h"
#ifdef _DEBUG
#include "FolderWindow.FileOperations.SelfTest.h"
#endif
#include "HostServices.h"
#include "NavigationLocation.h"

#include <limits>
#include <yyjson.h>

namespace
{
struct FileSystemCapabilitiesV1
{
    bool read            = false;
    bool write           = false;
    bool deleteOperation = false;
    bool properties      = false;

    std::vector<std::wstring> exportCopy;
    std::vector<std::wstring> exportMove;
    std::vector<std::wstring> importCopy;
    std::vector<std::wstring> importMove;
};

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::vector<std::wstring> ParsePluginIdList(yyjson_val* value) noexcept
{
    std::vector<std::wstring> result;
    if (! value || ! yyjson_is_arr(value))
    {
        return result;
    }

    const size_t count = yyjson_arr_size(value);
    result.reserve(count);
    for (size_t i = 0; i < count; ++i)
    {
        yyjson_val* item = yyjson_arr_get(value, i);
        if (! item || ! yyjson_is_str(item))
        {
            continue;
        }

        const char* s = yyjson_get_str(item);
        if (! s || s[0] == '\0')
        {
            continue;
        }

        std::wstring wide = Utf16FromUtf8(s);
        if (! wide.empty())
        {
            result.emplace_back(std::move(wide));
        }
    }

    return result;
}

[[nodiscard]] std::optional<FileSystemCapabilitiesV1> TryParseCapabilitiesJson(std::string_view jsonUtf8) noexcept
{
    if (jsonUtf8.empty())
    {
        return std::nullopt;
    }

    // yyjson may modify the input buffer; it requires a mutable char*.
    std::string jsonCopy(jsonUtf8);
    std::unique_ptr<yyjson_doc, decltype(&yyjson_doc_free)> doc(
        yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, nullptr), &yyjson_doc_free);
    if (! doc)
    {
        return std::nullopt;
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    if (! root || ! yyjson_is_obj(root))
    {
        return std::nullopt;
    }

    yyjson_val* versionVal = yyjson_obj_get(root, "version");
    if (! versionVal || ! yyjson_is_int(versionVal) || yyjson_get_int(versionVal) != 1)
    {
        return std::nullopt;
    }

    FileSystemCapabilitiesV1 out{};

    if (yyjson_val* ops = yyjson_obj_get(root, "operations"); ops && yyjson_is_obj(ops))
    {
        if (yyjson_val* v = yyjson_obj_get(ops, "read"); v && yyjson_is_bool(v))
        {
            out.read = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(ops, "write"); v && yyjson_is_bool(v))
        {
            out.write = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(ops, "delete"); v && yyjson_is_bool(v))
        {
            out.deleteOperation = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(ops, "properties"); v && yyjson_is_bool(v))
        {
            out.properties = yyjson_get_bool(v);
        }
    }

    if (yyjson_val* cross = yyjson_obj_get(root, "crossFileSystem"); cross && yyjson_is_obj(cross))
    {
        if (yyjson_val* exp = yyjson_obj_get(cross, "export"); exp && yyjson_is_obj(exp))
        {
            out.exportCopy = ParsePluginIdList(yyjson_obj_get(exp, "copy"));
            out.exportMove = ParsePluginIdList(yyjson_obj_get(exp, "move"));
        }
        if (yyjson_val* imp = yyjson_obj_get(cross, "import"); imp && yyjson_is_obj(imp))
        {
            out.importCopy = ParsePluginIdList(yyjson_obj_get(imp, "copy"));
            out.importMove = ParsePluginIdList(yyjson_obj_get(imp, "move"));
        }
    }

    return out;
}

[[nodiscard]] std::optional<FileSystemCapabilitiesV1> TryGetCapabilities(const wil::com_ptr<IFileSystem>& fileSystem) noexcept
{
    if (! fileSystem)
    {
        return std::nullopt;
    }

    const char* jsonUtf8 = nullptr;
    const HRESULT hr     = fileSystem->GetCapabilities(&jsonUtf8);
    if (FAILED(hr) || ! jsonUtf8 || jsonUtf8[0] == '\0')
    {
        return std::nullopt;
    }

    const std::string_view jsonView(jsonUtf8);
    return TryParseCapabilitiesJson(jsonView);
}

[[nodiscard]] bool IdListAllows(const std::vector<std::wstring>& allowedIds, std::wstring_view otherPluginId) noexcept
{
    if (otherPluginId.empty())
    {
        return false;
    }

    for (const auto& id : allowedIds)
    {
        if (id == L"*")
        {
            return true;
        }

        const int result = CompareStringOrdinal(id.data(), static_cast<int>(id.size()), otherPluginId.data(), static_cast<int>(otherPluginId.size()), TRUE);
        if (result == CSTR_EQUAL)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool CanCrossFileSystemCopyMove(const wil::com_ptr<IFileSystem>& sourceFileSystem,
                                              std::wstring_view sourcePluginId,
                                              const wil::com_ptr<IFileSystem>& destinationFileSystem,
                                              std::wstring_view destinationPluginId,
                                              FileSystemOperation operation) noexcept
{
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return false;
    }

    const std::optional<FileSystemCapabilitiesV1> sourceCaps = TryGetCapabilities(sourceFileSystem);
    const std::optional<FileSystemCapabilitiesV1> destCaps   = TryGetCapabilities(destinationFileSystem);
    if (! sourceCaps.has_value() || ! destCaps.has_value())
    {
        return false;
    }

    if (! sourceCaps->read || ! destCaps->write)
    {
        return false;
    }

    if (operation == FILESYSTEM_MOVE && ! sourceCaps->deleteOperation)
    {
        return false;
    }

    const std::vector<std::wstring>& exportList = operation == FILESYSTEM_COPY ? sourceCaps->exportCopy : sourceCaps->exportMove;
    const std::vector<std::wstring>& importList = operation == FILESYSTEM_COPY ? destCaps->importCopy : destCaps->importMove;

    return IdListAllows(exportList, destinationPluginId) && IdListAllows(importList, sourcePluginId);
}
} // namespace

void FolderWindow::FileOperationStateDeleter::operator()(FileOperationState* state) const noexcept
{
    std::default_delete<FileOperationState>{}(state);
}

void FolderWindow::EnsureFileOperations()
{
    if (_fileOperations)
    {
        return;
    }

    auto state = std::make_unique<FileOperationState>(*this);
    _fileOperations.reset(state.release());
}

HRESULT FolderWindow::StartFileOperationFromFolderView(Pane pane, FolderView::FileOperationRequest request) noexcept
{
    PaneState& destinationState = pane == Pane::Left ? _leftPane : _rightPane;
    if (! destinationState.fileSystem)
    {
        return E_POINTER;
    }

    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_FAIL;
    }

    const bool isCopyMove = request.operation == FILESYSTEM_COPY || request.operation == FILESYSTEM_MOVE;

    Pane sourcePane                      = pane;
    std::optional<Pane> destinationPane  = std::nullopt;
    wil::com_ptr<IFileSystem> fileSystem = destinationState.fileSystem;
    wil::com_ptr<IFileSystem> destinationFileSystem;

    if (isCopyMove && request.sourceContextSpecified)
    {
        const auto contextMatches = [&](const PaneState& paneState) noexcept -> bool
        {
            return CompareStringOrdinal(paneState.pluginId.c_str(), -1, request.sourcePluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                   NavigationLocation::EqualsNoCase(paneState.instanceContext, request.sourceInstanceContext);
        };

        const bool leftMatches  = contextMatches(_leftPane);
        const bool rightMatches = contextMatches(_rightPane);

        if (leftMatches ^ rightMatches)
        {
            sourcePane = leftMatches ? Pane::Left : Pane::Right;
        }
        else if (leftMatches && rightMatches)
        {
            const auto isUnderFolder = [](std::wstring_view folder, std::wstring_view path) noexcept -> bool
            {
                while (! folder.empty() && (folder.back() == L'\\' || folder.back() == L'/'))
                {
                    folder.remove_suffix(1);
                }
                if (folder.empty() || path.size() <= folder.size())
                {
                    return false;
                }

                if (CompareStringOrdinal(path.data(), static_cast<int>(folder.size()), folder.data(), static_cast<int>(folder.size()), TRUE) != CSTR_EQUAL)
                {
                    return false;
                }

                const wchar_t next = path[folder.size()];
                return next == L'\\' || next == L'/';
            };

            bool inferredSourcePane = false;
            if (! request.sourcePaths.empty())
            {
                const std::wstring_view firstPath = request.sourcePaths.front().native();

                const auto leftFolder  = _leftPane.folderView.GetFolderPath();
                const auto rightFolder = _rightPane.folderView.GetFolderPath();

                const bool underLeft  = leftFolder.has_value() && isUnderFolder(leftFolder->native(), firstPath);
                const bool underRight = rightFolder.has_value() && isUnderFolder(rightFolder->native(), firstPath);

                if (underLeft ^ underRight)
                {
                    sourcePane         = underLeft ? Pane::Left : Pane::Right;
                    inferredSourcePane = true;
                }
            }

            if (! inferredSourcePane)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
        }
        else
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        if (sourcePane != pane)
        {
            PaneState& sourceState = sourcePane == Pane::Left ? _leftPane : _rightPane;
            if (! sourceState.fileSystem)
            {
                return E_POINTER;
            }

            if (! SanityCheckBothPanes(sourceState, destinationState, request.operation))
            {
                return E_FAIL;
            }

            fileSystem      = sourceState.fileSystem;
            destinationPane = pane;

            const bool contextSame = CompareStringOrdinal(sourceState.pluginId.c_str(), -1, destinationState.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                     NavigationLocation::EqualsNoCase(sourceState.instanceContext, destinationState.instanceContext);
            destinationFileSystem = contextSame ? nullptr : destinationState.fileSystem;
        }
    }

    const bool waitForOthers                = _fileOperations->ShouldQueueNewTask();
    std::filesystem::path destinationFolder = request.destinationFolder.value_or(std::filesystem::path{});
    return _fileOperations->StartOperation(request.operation,
                                           sourcePane,
                                           destinationPane,
                                           fileSystem,
                                           std::move(request.sourcePaths),
                                           std::move(destinationFolder),
                                           request.flags,
                                           waitForOthers,
                                           0,
                                           FileOperationState::ExecutionMode::PerItem,
                                           false,
                                           std::move(destinationFileSystem));
}

void FolderWindow::ShutdownFileOperations() noexcept
{
    _fileOperations.reset();
}

void FolderWindow::ApplyFileOperationsTheme() noexcept
{
    if (_fileOperations)
    {
        _fileOperations->ApplyTheme(_theme);
    }
}

void FolderWindow::CommandToggleFileOperationsIssuesPane()
{
    EnsureFileOperations();
    if (! _fileOperations)
    {
        return;
    }

    _fileOperations->ToggleIssuesPane();
}

bool FolderWindow::IsFileOperationsIssuesPaneVisible() noexcept
{
    if (! _fileOperations)
    {
        return false;
    }

    return _fileOperations->IsIssuesPaneVisible();
}

#ifdef _DEBUG
FolderWindow::FileOperationState* FolderWindow::DebugGetFileOperationState() noexcept
{
    EnsureFileOperations();
    return _fileOperations.get();
}
#endif

bool FolderWindow::ConfirmCancelAllFileOperations(HWND ownerWindow) noexcept
{
    if (! _fileOperations || ! _fileOperations->HasActiveOperations())
    {
        return true;
    }

    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_FILEOPS_EXIT);
    const std::wstring message = LoadStringResource(nullptr, IDS_MSG_FILEOPS_CANCEL_ALL_EXIT);

    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_INFO;
    prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = ownerWindow;
    prompt.title         = title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_CANCEL;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hrPrompt) || promptResult != HOST_PROMPT_RESULT_OK)
    {
        return false;
    }

    _fileOperations->CancelAll();
    return true;
}

void FolderWindow::CommandDelete(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! _fileOperations)
    {
        state.folderView.CommandDelete();
        return;
    }

    if (! state.fileSystem)
    {
        return;
    }

    std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN);

    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    static_cast<void>(_fileOperations->StartOperation(
        FILESYSTEM_DELETE, pane, std::nullopt, state.fileSystem, std::move(paths), {}, flags, waitForOthers, 0, FileOperationState::ExecutionMode::PerItem));
}

void FolderWindow::CommandPermanentDelete(Pane pane)
{
    SetActivePane(pane);
    EnsureFileOperations();

    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! _fileOperations || ! state.fileSystem)
    {
        return;
    }

    std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    static_cast<void>(_fileOperations->StartOperation(
        FILESYSTEM_DELETE, pane, std::nullopt, state.fileSystem, std::move(paths), {}, flags, waitForOthers, 0, FileOperationState::ExecutionMode::PerItem));
}

void FolderWindow::CommandPermanentDeleteWithValidation(Pane pane)
{
    SetActivePane(pane);
    EnsureFileOperations();

    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! _fileOperations || ! state.fileSystem)
    {
        return;
    }

    std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    static_cast<void>(_fileOperations->StartOperation(FILESYSTEM_DELETE,
                                                      pane,
                                                      std::nullopt,
                                                      state.fileSystem,
                                                      std::move(paths),
                                                      {},
                                                      flags,
                                                      waitForOthers,
                                                      0,
                                                      FileOperationState::ExecutionMode::PerItem,
                                                      true));
}

bool FolderWindow::SanityCheckBothPanes(FolderWindow::PaneState& src, FolderWindow::PaneState& dest, FileSystemOperation operation)
{
    bool ok             = true;
    bool sameFolder     = false;
    bool contextsDiffer = false;
    if (! _fileOperations)
    {
        Debug::Error(L"FolderWindow::SanityCheckBothPanes No active file operations.");
        ok = false;
    }

    if (ok && (! src.fileSystem || ! dest.fileSystem))
    {
        Debug::Error(L"FolderWindow::SanityCheckBothPanes Source or destination pane has no file system.");
        ok = false;
    }

    if (ok && (src.pluginId.empty() || dest.pluginId.empty()))
    {
        Debug::Error(L"FolderWindow::SanityCheckBothPanes Source or destination pane has no file system metadata.");
        ok = false;
    }

    if (ok && (! dest.folderView.GetFolderPath().has_value()))
    {
        Debug::Error(L"FolderWindow::SanityCheckBothPanes No destination path.");
        ok = false;
    }

    if (ok)
    {
        const bool contextSame = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                 NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
        contextsDiffer = ! contextSame;

        const auto srcFolder = src.folderView.GetFolderPath();
        const auto dstFolder = dest.folderView.GetFolderPath();
        if (contextSame && srcFolder.has_value() && dstFolder.has_value() &&
            NavigationLocation::EqualsNoCase(srcFolder.value().native(), dstFolder.value().native()))
        {
            Debug::Error(L"FolderWindow::SanityCheckBothPanes Source and destination folder are the same: {}.", srcFolder.value().native());
            sameFolder = true;
            ok         = false;
        }
    }

    if (ok && contextsDiffer && (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE))
    {
        if (! CanCrossFileSystemCopyMove(src.fileSystem, src.pluginId, dest.fileSystem, dest.pluginId, operation))
        {
            Debug::Error(L"FolderWindow::SanityCheckBothPanes Cross-filesystem operation not allowed src:{} dest:{} op:{}.",
                         src.pluginId,
                         dest.pluginId,
                         static_cast<unsigned int>(operation));
            ok = false;
        }
    }

    if (! ok && _hWnd)
    {
        const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        int messageId              = sameFolder       ? IDS_MSG_PANE_OP_REQUIRES_DIFFERENT_FOLDER
                                     : contextsDiffer ? IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS
                                                      : IDS_MSG_PANE_OP_REQUIRES_SAME_FS;
        const std::wstring message = LoadStringResource(nullptr, static_cast<UINT>(messageId));
        src.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, title, message);
        return false;
    }

    return ok;
}

void FolderWindow::CommandCopyToOtherPane(Pane sourcePane)
{
    SetActivePane(sourcePane);
    const Pane destPane = sourcePane == Pane::Left ? Pane::Right : Pane::Left;

    PaneState& src  = sourcePane == Pane::Left ? _leftPane : _rightPane;
    PaneState& dest = destPane == Pane::Left ? _leftPane : _rightPane;

    if (! SanityCheckBothPanes(src, dest, FILESYSTEM_COPY))
    {
        return;
    }

    std::vector<std::filesystem::path> paths = src.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        auto srcPath = src.currentPath.has_value() ? src.currentPath.value().c_str() : L"(unknown)";
        Debug::Error(L"FolderWindow::CommandCopyToOtherPane No selected paths: {}", srcPath);
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    const bool contextSame   = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                             NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
    wil::com_ptr<IFileSystem> destinationFileSystem = contextSame ? nullptr : dest.fileSystem;
    static_cast<void>(_fileOperations->StartOperation(FILESYSTEM_COPY,
                                                      sourcePane,
                                                      destPane,
                                                      src.fileSystem,
                                                      std::move(paths),
                                                      dest.folderView.GetFolderPath().value(),
                                                      flags,
                                                      waitForOthers,
                                                      0,
                                                      FileOperationState::ExecutionMode::PerItem,
                                                      false,
                                                      std::move(destinationFileSystem)));
}

void FolderWindow::CommandMoveToOtherPane(Pane sourcePane)
{
    SetActivePane(sourcePane);
    const Pane destPane = sourcePane == Pane::Left ? Pane::Right : Pane::Left;

    PaneState& src  = sourcePane == Pane::Left ? _leftPane : _rightPane;
    PaneState& dest = destPane == Pane::Left ? _leftPane : _rightPane;

    if (! SanityCheckBothPanes(src, dest, FILESYSTEM_MOVE))
    {
        return;
    }

    std::vector<std::filesystem::path> paths = src.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        auto srcPath = src.currentPath.has_value() ? src.currentPath.value().c_str() : L"(unknown)";
        Debug::Error(L"FolderWindow::CommandMoveToOtherPane No selected paths: {}", srcPath);
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    const bool contextSame   = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                             NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
    wil::com_ptr<IFileSystem> destinationFileSystem = contextSame ? nullptr : dest.fileSystem;
    static_cast<void>(_fileOperations->StartOperation(FILESYSTEM_MOVE,
                                                      sourcePane,
                                                      destPane,
                                                      src.fileSystem,
                                                      std::move(paths),
                                                      dest.folderView.GetFolderPath().value(),
                                                      flags,
                                                      waitForOthers,
                                                      0,
                                                      FileOperationState::ExecutionMode::PerItem,
                                                      false,
                                                      std::move(destinationFileSystem)));
}

LRESULT FolderWindow::OnFileOperationCompleted(LPARAM lp) noexcept
{
    auto payload = TakeMessagePayload<FileOperationState::TaskCompletedPayload>(lp);
    if (! payload)
    {
        return 0;
    }

    if (! _fileOperations)
    {
        return 0;
    }

#ifdef _DEBUG
    if (FileOperationsSelfTest::IsRunning())
    {
        FileOperationsSelfTest::NotifyTaskCompleted(payload->taskId, payload->hr);
    }
#endif

    FileOperationState::Task* task = _fileOperations->FindTask(payload->taskId);
    if (! task)
    {
        return 0;
    }

    const Pane sourcePane                     = task->GetSourcePane();
    const std::optional<Pane> destinationPane = task->GetDestinationPane();

    if (_fileOperationCompletedCallback)
    {
        FileOperationCompletedEvent e{};
        e.operation         = task->GetOperation();
        e.sourcePane        = sourcePane;
        e.destinationPane   = destinationPane;
        e.sourcePaths       = task->_sourcePaths;
        e.destinationFolder = task->GetDestinationFolder();
        e.hr                = payload->hr;
        _fileOperationCompletedCallback(e);
    }

    PaneState& src            = sourcePane == Pane::Left ? _leftPane : _rightPane;
    DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();

    const auto srcFolder = src.folderView.GetFolderPath();
    if (! src.fileSystem || ! srcFolder.has_value() || ! cache.IsFolderWatched(src.fileSystem.get(), srcFolder.value()))
    {
        src.folderView.ForceRefresh();
    }

    if (destinationPane.has_value())
    {
        PaneState& dst       = destinationPane.value() == Pane::Left ? _leftPane : _rightPane;
        const auto dstFolder = dst.folderView.GetFolderPath();
        if (! dst.fileSystem || ! dstFolder.has_value() || ! cache.IsFolderWatched(dst.fileSystem.get(), dstFolder.value()))
        {
            dst.folderView.ForceRefresh();
        }
    }

    const bool autoDismissSuccess = _fileOperations->GetAutoDismissSuccess();
    _fileOperations->RemoveTask(payload->taskId);
    constexpr HRESULT cancelledHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    if (autoDismissSuccess && (SUCCEEDED(payload->hr) || payload->hr == cancelledHr || payload->hr == E_ABORT))
    {
        _fileOperations->DismissCompletedTask(payload->taskId);
    }
    return 0;
}
