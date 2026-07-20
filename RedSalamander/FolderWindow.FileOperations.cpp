#include "FolderWindow.FileOperationsInternal.h"
#ifdef ENABLE_TESTS
#include "FolderWindow.FileOperations.SelfTest.h"
#endif
#include "FileSystemPathIdentity.h"
#include "HostServices.h"
#include "NavigationLocation.h"

#include <limits>
#include <unordered_set>

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

namespace
{
struct FileSystemCapabilitiesV1
{
    bool copyOperation   = false;
    bool moveOperation   = false;
    bool read            = false;
    bool write           = false;
    bool deleteOperation = false;
    bool properties      = false;

    std::vector<std::wstring> exportCopy;
    std::vector<std::wstring> exportMove;
    std::vector<std::wstring> importCopy;
    std::vector<std::wstring> importMove;
    std::optional<FileSystemPathIdentity> pathIdentity;
};

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
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

    // The mandatory shape requires all four sections (PlugInterfaces/FileSystem.h); a document
    // missing any of them is a provider contract violation and must fail closed rather than be
    // silently defaulted.
    yyjson_val* requiredOperations      = yyjson_obj_get(root, "operations");
    yyjson_val* requiredConcurrency     = yyjson_obj_get(root, "concurrency");
    yyjson_val* requiredCrossFileSystem = yyjson_obj_get(root, "crossFileSystem");
    yyjson_val* requiredPathIdentity    = yyjson_obj_get(root, "pathIdentity");
    if (! requiredOperations || ! yyjson_is_obj(requiredOperations) || ! requiredConcurrency || ! yyjson_is_obj(requiredConcurrency) ||
        ! requiredCrossFileSystem || ! yyjson_is_obj(requiredCrossFileSystem) || ! requiredPathIdentity || ! yyjson_is_obj(requiredPathIdentity))
    {
        return std::nullopt;
    }

    FileSystemCapabilitiesV1 out{};
    out.pathIdentity = TryParseFileSystemPathIdentityContractFromRoot(root, {});
    if (! out.pathIdentity.has_value())
    {
        return std::nullopt;
    }

    if (yyjson_val* ops = yyjson_obj_get(root, "operations"); ops && yyjson_is_obj(ops))
    {
        if (yyjson_val* v = yyjson_obj_get(ops, "copy"); v && yyjson_is_bool(v))
        {
            out.copyOperation = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(ops, "move"); v && yyjson_is_bool(v))
        {
            out.moveOperation = yyjson_get_bool(v);
        }
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

[[nodiscard]] bool HasStablePathIdentity(const FileSystemCapabilitiesV1& capabilities) noexcept
{
    return capabilities.pathIdentity.has_value() && capabilities.pathIdentity->pathTextStableIdentity;
}

// GetCapabilities is mandatory: every provider must return S_OK and a parseable version-1 document. Any other
// response is a provider contract violation; the host fails closed and surfaces the violation once per instance.
void ReportCapabilitiesContractViolationOnce(const IFileSystem* fileSystem, std::wstring_view pluginId, HRESULT hr) noexcept
{
    static std::mutex reportedMutex;
    static std::unordered_set<const IFileSystem*> reportedProviders;

    {
        // The only throwing operation here is the set insertion, and that only throws
        // std::bad_alloc; per the repo exception policy allocation failure terminates
        // (this function is noexcept), so no catch is needed.
        const std::lock_guard lock(reportedMutex);
        if (! reportedProviders.insert(fileSystem).second)
        {
            return;
        }
    }

    Debug::Error(L"FolderWindow filesystem provider '{}' violates the mandatory GetCapabilities contract (hr=0x{:08X}); "
                 L"capability-gated operations are disabled for this instance.",
                 pluginId.empty() ? std::wstring_view(L"<unknown>") : pluginId,
                 static_cast<unsigned long>(hr));
}

[[nodiscard]] std::optional<FileSystemCapabilitiesV1> TryGetCapabilities(const wil::com_ptr<IFileSystem>& fileSystem, std::wstring_view pluginId = {}) noexcept
{
    if (! fileSystem)
    {
        return std::nullopt;
    }

    const char* jsonUtf8 = nullptr;
    const HRESULT hr     = fileSystem->GetCapabilities(&jsonUtf8);
    if (FAILED(hr) || ! jsonUtf8 || jsonUtf8[0] == '\0')
    {
        ReportCapabilitiesContractViolationOnce(fileSystem.get(), pluginId, hr);
        return std::nullopt;
    }

    const std::string_view jsonView(jsonUtf8);
    std::optional<FileSystemCapabilitiesV1> capabilities = TryParseCapabilitiesJson(jsonView);
    if (! capabilities.has_value())
    {
        ReportCapabilitiesContractViolationOnce(fileSystem.get(), pluginId, hr);
    }

    return capabilities;
}

[[nodiscard]] bool CanSameFileSystemOperationFromCapabilities(const wil::com_ptr<IFileSystem>& fileSystem,
                                                              FileSystemOperation operation,
                                                              std::wstring_view pluginId = {}) noexcept
{
    const std::optional<FileSystemCapabilitiesV1> capabilities = TryGetCapabilities(fileSystem, pluginId);
    if (! capabilities.has_value())
    {
        return false;
    }

    switch (operation)
    {
        case FILESYSTEM_COPY: return capabilities->copyOperation && HasStablePathIdentity(capabilities.value());
        case FILESYSTEM_MOVE: return capabilities->moveOperation && HasStablePathIdentity(capabilities.value());
        case FILESYSTEM_DELETE: return capabilities->deleteOperation && HasStablePathIdentity(capabilities.value());
        case FILESYSTEM_RENAME: return true;
        default: return true;
    }
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

        if (OrdinalString::EqualsNoCase(id, otherPluginId))
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

    const std::optional<FileSystemCapabilitiesV1> sourceCaps = TryGetCapabilities(sourceFileSystem, sourcePluginId);
    const std::optional<FileSystemCapabilitiesV1> destCaps   = TryGetCapabilities(destinationFileSystem, destinationPluginId);
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

bool CanSameFileSystemOperation(const wil::com_ptr<IFileSystem>& fileSystem, FileSystemOperation operation, std::wstring_view pluginId) noexcept
{
    return CanSameFileSystemOperationFromCapabilities(fileSystem, operation, pluginId);
}

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

uint64_t FolderWindow::CreateOrUpdateInformationalTask(const InformationalTaskUpdate& update) noexcept
{
    EnsureFileOperations();
    if (! _fileOperations)
    {
        return 0;
    }

    return _fileOperations->CreateOrUpdateInformationalTask(update);
}

void FolderWindow::DismissInformationalTask(uint64_t taskId) noexcept
{
    if (taskId == 0 || ! _fileOperations)
    {
        return;
    }

    _fileOperations->DismissInformationalTask(taskId);
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
    std::wstring sourcePluginIdOverride;
    std::wstring sourcePluginShortIdOverride;

    if (isCopyMove && ! request.sourceContextSpecified &&
        CompareStringOrdinal(destinationState.pluginId.c_str(), -1, L"builtin/file-system", -1, TRUE) != CSTR_EQUAL)
    {
        const FileSystemPluginManager::PluginEntry* localEntry = FileSystemPluginManager::GetInstance().FindPluginById(L"builtin/file-system");
        if (! localEntry || ! localEntry->fileSystem || localEntry->disabled || ! localEntry->loadable || localEntry->shortId.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        if (! CanCrossFileSystemCopyMove(localEntry->fileSystem, localEntry->id, destinationState.fileSystem, destinationState.pluginId, request.operation))
        {
            destinationState.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                                         FolderView::OverlaySeverity::Error,
                                                         LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                                         LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        fileSystem                  = localEntry->fileSystem;
        destinationFileSystem       = destinationState.fileSystem;
        destinationPane             = pane;
        sourcePluginIdOverride      = localEntry->id;
        sourcePluginShortIdOverride = localEntry->shortId;
    }

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

                if (! OrdinalString::StartsWithNoCase(path, folder))
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
            destinationFileSystem  = contextSame ? nullptr : destinationState.fileSystem;
        }
    }

    const std::wstring& sourcePluginIdForGate =
        sourcePluginIdOverride.empty() ? (sourcePane == Pane::Left ? _leftPane.pluginId : _rightPane.pluginId) : sourcePluginIdOverride;
    if (! destinationFileSystem && ! CanSameFileSystemOperation(fileSystem, request.operation, sourcePluginIdForGate))
    {
        Debug::Error(L"FolderWindow::StartFileOperationFromFolderView provider rejected same-filesystem operation plugin:{} op:{}.",
                     sourcePluginIdForGate,
                     static_cast<unsigned int>(request.operation));
        destinationState.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                                     FolderView::OverlaySeverity::Error,
                                                     LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                                     LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const bool waitForOthers                = _fileOperations->ShouldQueueNewTask();
    std::filesystem::path destinationFolder = request.destinationFolder.value_or(std::filesystem::path{});
    uint64_t taskId                         = 0;
    HRESULT startHr                         = _fileOperations->StartOperation(request.operation,
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
                                                                              std::move(destinationFileSystem),
                                                                              &taskId,
                                                                              {},
                                                                              {},
                                                                              std::move(sourcePluginIdOverride),
                                                                              std::move(sourcePluginShortIdOverride));
    if (SUCCEEDED(startHr) && taskId != 0u && request.completionCallback)
    {
        _fileOperationRequestCompletionCallbacks.insert_or_assign(taskId, std::move(request.completionCallback));
    }
    return startHr;
}

std::optional<FolderWindow::Pane> FolderWindow::ResolveSourcePaneForResolvedPaths(std::wstring_view sourcePluginId,
                                                                                  std::wstring_view sourceInstanceContext,
                                                                                  const std::vector<std::filesystem::path>& sourcePaths) const noexcept
{
    if (sourcePluginId.empty() || sourcePaths.empty())
    {
        return std::nullopt;
    }

    const auto contextMatches = [&](const PaneState& state) noexcept -> bool
    {
        return CompareStringOrdinal(state.pluginId.c_str(), -1, sourcePluginId.data(), static_cast<int>(sourcePluginId.size()), TRUE) == CSTR_EQUAL &&
               NavigationLocation::EqualsNoCase(state.instanceContext, sourceInstanceContext);
    };

    const bool leftMatches  = contextMatches(_leftPane);
    const bool rightMatches = contextMatches(_rightPane);
    if (! leftMatches && ! rightMatches)
    {
        return std::nullopt;
    }

    const auto isSameOrUnderFolder = [](std::wstring_view folder, std::wstring_view path) noexcept -> bool
    {
        while (! folder.empty() && (folder.back() == L'\\' || folder.back() == L'/'))
        {
            folder.remove_suffix(1);
        }
        if (folder.empty() || path.size() < folder.size())
        {
            return false;
        }

        if (! OrdinalString::StartsWithNoCase(path, folder))
        {
            return false;
        }

        return path.size() == folder.size() || path[folder.size()] == L'\\' || path[folder.size()] == L'/';
    };

    const auto pathFitsPane = [&](Pane pane) noexcept -> bool
    {
        const PaneState& state                            = pane == Pane::Left ? _leftPane : _rightPane;
        const std::optional<std::filesystem::path> folder = state.folderView.GetFolderPath();
        return folder.has_value() && isSameOrUnderFolder(folder->native(), sourcePaths.front().native());
    };

    Pane sourcePane = Pane::Left;
    if (leftMatches && rightMatches)
    {
        const Pane focusedPane = GetFocusedPane();
        if (pathFitsPane(focusedPane))
        {
            sourcePane = focusedPane;
        }
        else if (pathFitsPane(_activePane))
        {
            sourcePane = _activePane;
        }
        else if (pathFitsPane(Pane::Left))
        {
            sourcePane = Pane::Left;
        }
        else if (pathFitsPane(Pane::Right))
        {
            sourcePane = Pane::Right;
        }
        else
        {
            sourcePane = focusedPane;
        }
    }
    else
    {
        sourcePane = leftMatches ? Pane::Left : Pane::Right;
    }

    return sourcePane;
}

std::optional<std::filesystem::path> FolderWindow::GetOtherPaneDestinationForResolvedPaths(std::wstring_view sourcePluginId,
                                                                                           std::wstring_view sourceInstanceContext,
                                                                                           const std::vector<std::filesystem::path>& sourcePaths) const noexcept
{
    const std::optional<Pane> sourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourcePaths);
    if (! sourcePane.has_value())
    {
        return std::nullopt;
    }

    const Pane destinationPane        = OppositePane(sourcePane.value());
    const PaneState& destinationState = destinationPane == Pane::Left ? _leftPane : _rightPane;
    return destinationState.folderView.GetFolderPath();
}

HRESULT FolderWindow::StartFileOperationForResolvedPaths(std::wstring_view sourcePluginId,
                                                         std::wstring_view sourceInstanceContext,
                                                         FileSystemOperation operation,
                                                         std::vector<std::filesystem::path> sourcePaths,
                                                         FileSystemFlags flags,
                                                         bool requireConfirmation,
                                                         uint64_t* taskIdOut) noexcept
{
    if (taskIdOut)
    {
        *taskIdOut = 0;
    }

    if (sourcePluginId.empty())
    {
        return E_INVALIDARG;
    }

    if (sourcePaths.empty())
    {
        return S_FALSE;
    }

    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_FAIL;
    }

    const std::optional<Pane> resolvedSourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourcePaths);
    if (! resolvedSourcePane.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const Pane sourcePane  = resolvedSourcePane.value();
    PaneState& sourceState = sourcePane == Pane::Left ? _leftPane : _rightPane;
    if (! sourceState.fileSystem)
    {
        return E_POINTER;
    }

    if (! CanSameFileSystemOperation(sourceState.fileSystem, operation, sourceState.pluginId))
    {
        Debug::Error(L"FolderWindow::StartFileOperationForResolvedPaths provider rejected operation plugin:{} op:{}.",
                     sourceState.pluginId,
                     static_cast<unsigned int>(operation));
        sourceState.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                                FolderView::OverlaySeverity::Error,
                                                LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                                LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    return _fileOperations->StartOperation(operation,
                                           sourcePane,
                                           std::nullopt,
                                           sourceState.fileSystem,
                                           std::move(sourcePaths),
                                           {},
                                           flags,
                                           waitForOthers,
                                           0,
                                           FileOperationState::ExecutionMode::PerItem,
                                           requireConfirmation,
                                           nullptr,
                                           taskIdOut);
}

HRESULT FolderWindow::StartFileOperationForResolvedItemsToOtherPane(std::wstring_view sourcePluginId,
                                                                    std::wstring_view sourceInstanceContext,
                                                                    FileSystemOperation operation,
                                                                    std::vector<ResolvedFileOperationItem> items,
                                                                    std::optional<std::filesystem::path>* outDestinationFolder,
                                                                    uint64_t* taskIdOut,
                                                                    std::wstring confirmationMessage) noexcept
{
    if (taskIdOut)
    {
        *taskIdOut = 0;
    }

    if (outDestinationFolder)
    {
        outDestinationFolder->reset();
    }

    if (sourcePluginId.empty() || (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE))
    {
        return E_INVALIDARG;
    }

    if (items.empty())
    {
        return S_FALSE;
    }

    std::vector<std::filesystem::path> sourcePaths;
    sourcePaths.reserve(items.size());
    for (const ResolvedFileOperationItem& item : items)
    {
        if (item.sourcePath.empty() || item.destinationPath.empty())
        {
            return E_INVALIDARG;
        }
        sourcePaths.push_back(item.sourcePath);
    }

    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_FAIL;
    }

    const std::optional<Pane> resolvedSourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourcePaths);
    if (! resolvedSourcePane.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const Pane sourcePane = resolvedSourcePane.value();
    const Pane destPane   = OppositePane(sourcePane);
    PaneState& src        = sourcePane == Pane::Left ? _leftPane : _rightPane;
    PaneState& dest       = destPane == Pane::Left ? _leftPane : _rightPane;
    if (! src.fileSystem || ! dest.fileSystem)
    {
        return E_POINTER;
    }

    if (! SanityCheckBothPanes(src, dest, operation))
    {
        return E_FAIL;
    }

    const std::optional<std::filesystem::path> destinationFolder = dest.folderView.GetFolderPath();
    if (! destinationFolder.has_value())
    {
        return E_FAIL;
    }
    if (outDestinationFolder)
    {
        *outDestinationFolder = destinationFolder.value();
    }

    const bool waitForOthers                        = _fileOperations->ShouldQueueNewTask();
    const bool contextSame                          = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                                      NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
    wil::com_ptr<IFileSystem> destinationFileSystem = contextSame ? nullptr : dest.fileSystem;
    if (! destinationFileSystem && ! CanSameFileSystemOperation(src.fileSystem, operation, src.pluginId))
    {
        Debug::Error(L"FolderWindow::StartFileOperationForResolvedItemsToOtherPane provider rejected same-filesystem operation plugin:{} op:{}.",
                     src.pluginId,
                     static_cast<unsigned int>(operation));
        src.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                        FolderView::OverlaySeverity::Error,
                                        LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                        LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const bool requireConfirmation = operation == FILESYSTEM_MOVE;
    return _fileOperations->StartOperation(operation,
                                           sourcePane,
                                           destPane,
                                           src.fileSystem,
                                           std::move(sourcePaths),
                                           destinationFolder.value(),
                                           FILESYSTEM_FLAG_NONE,
                                           waitForOthers,
                                           0,
                                           FileOperationState::ExecutionMode::PerItem,
                                           requireConfirmation,
                                           std::move(destinationFileSystem),
                                           taskIdOut,
                                           std::move(items),
                                           std::move(confirmationMessage));
}

HRESULT FolderWindow::StartFileOperationForResolvedPathsToOtherPane(std::wstring_view sourcePluginId,
                                                                    std::wstring_view sourceInstanceContext,
                                                                    FileSystemOperation operation,
                                                                    std::vector<std::filesystem::path> sourcePaths,
                                                                    FileSystemFlags flags,
                                                                    std::optional<std::filesystem::path>* outDestinationFolder,
                                                                    uint64_t* taskIdOut) noexcept
{
    if (taskIdOut)
    {
        *taskIdOut = 0;
    }

    if (outDestinationFolder)
    {
        outDestinationFolder->reset();
    }

    if (sourcePluginId.empty() || (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE))
    {
        return E_INVALIDARG;
    }

    if (sourcePaths.empty())
    {
        return S_FALSE;
    }

    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_FAIL;
    }

    const std::optional<Pane> resolvedSourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourcePaths);
    if (! resolvedSourcePane.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const Pane sourcePane = resolvedSourcePane.value();
    const Pane destPane   = OppositePane(sourcePane);
    PaneState& src        = sourcePane == Pane::Left ? _leftPane : _rightPane;
    PaneState& dest       = destPane == Pane::Left ? _leftPane : _rightPane;
    if (! src.fileSystem || ! dest.fileSystem)
    {
        return E_POINTER;
    }

    if (! SanityCheckBothPanes(src, dest, operation))
    {
        return E_FAIL;
    }

    const std::optional<std::filesystem::path> destinationFolder = dest.folderView.GetFolderPath();
    if (! destinationFolder.has_value())
    {
        return E_FAIL;
    }
    if (outDestinationFolder)
    {
        *outDestinationFolder = destinationFolder.value();
    }

    const bool waitForOthers                        = _fileOperations->ShouldQueueNewTask();
    const bool contextSame                          = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                                      NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
    wil::com_ptr<IFileSystem> destinationFileSystem = contextSame ? nullptr : dest.fileSystem;
    if (! destinationFileSystem && ! CanSameFileSystemOperation(src.fileSystem, operation, src.pluginId))
    {
        Debug::Error(L"FolderWindow::StartFileOperationForResolvedPathsToOtherPane provider rejected same-filesystem operation plugin:{} op:{}.",
                     src.pluginId,
                     static_cast<unsigned int>(operation));
        src.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                        FolderView::OverlaySeverity::Error,
                                        LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                        LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    return _fileOperations->StartOperation(operation,
                                           sourcePane,
                                           destPane,
                                           src.fileSystem,
                                           std::move(sourcePaths),
                                           destinationFolder.value(),
                                           flags,
                                           waitForOthers,
                                           0,
                                           FileOperationState::ExecutionMode::PerItem,
                                           false,
                                           std::move(destinationFileSystem),
                                           taskIdOut);
}

HRESULT FolderWindow::StartFileOperationForResolvedPathsToDestination(std::wstring_view sourcePluginId,
                                                                      std::wstring_view sourceInstanceContext,
                                                                      FileSystemOperation operation,
                                                                      std::vector<std::filesystem::path> sourcePaths,
                                                                      std::filesystem::path destinationFolder,
                                                                      FileSystemFlags flags,
                                                                      uint64_t* taskIdOut) noexcept
{
    if (taskIdOut)
    {
        *taskIdOut = 0;
    }

    if (sourcePluginId.empty() || (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE))
    {
        return E_INVALIDARG;
    }

    if (sourcePaths.empty())
    {
        return S_FALSE;
    }

    if (destinationFolder.empty())
    {
        return E_INVALIDARG;
    }

    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_FAIL;
    }

    const std::optional<Pane> resolvedSourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourcePaths);
    if (! resolvedSourcePane.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const Pane sourcePane = resolvedSourcePane.value();
    PaneState& src        = sourcePane == Pane::Left ? _leftPane : _rightPane;
    if (! src.fileSystem)
    {
        return E_POINTER;
    }

    if (! CanSameFileSystemOperation(src.fileSystem, operation, src.pluginId))
    {
        Debug::Error(L"FolderWindow::StartFileOperationForResolvedPathsToDestination provider rejected same-filesystem operation plugin:{} op:{}.",
                     src.pluginId,
                     static_cast<unsigned int>(operation));
        src.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                        FolderView::OverlaySeverity::Error,
                                        LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                        LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    return _fileOperations->StartOperation(operation,
                                           sourcePane,
                                           std::nullopt,
                                           src.fileSystem,
                                           std::move(sourcePaths),
                                           std::move(destinationFolder),
                                           flags,
                                           waitForOthers,
                                           0,
                                           FileOperationState::ExecutionMode::PerItem,
                                           false,
                                           nullptr,
                                           taskIdOut);
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

#ifdef ENABLE_TESTS
FolderWindow::FileOperationState* FolderWindow::DebugGetFileOperationState() noexcept
{
    EnsureFileOperations();
    return _fileOperations.get();
}

void FolderWindow::DebugSetFileOperationRequestCallbackEnabled(Pane pane, bool enabled) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! enabled)
    {
        state.folderView.SetFileOperationRequestCallback({});
        return;
    }

    state.folderView.SetFileOperationRequestCallback([this, pane](FolderView::FileOperationRequest request) noexcept -> HRESULT
    { return StartFileOperationFromFolderView(pane, std::move(request)); });
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

    if (! CanSameFileSystemOperation(state.fileSystem, FILESYSTEM_DELETE, state.pluginId))
    {
        Debug::Error(L"FolderWindow::CommandDelete provider rejected delete plugin:{}.", state.pluginId);
        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                          FolderView::OverlaySeverity::Error,
                                          LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                          LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN);

    // Recycle deletes go through ONE bulk DeleteItems call: the plugin batches same-parent
    // items into a single IFileOperation (up to 1000 per batch), which per-item routing would
    // degrade to one shell operation + one STA thread per item.
    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    static_cast<void>(_fileOperations->StartOperation(
        FILESYSTEM_DELETE, pane, std::nullopt, state.fileSystem, std::move(paths), {}, flags, waitForOthers, 0, FileOperationState::ExecutionMode::BulkItems));
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

    if (! CanSameFileSystemOperation(state.fileSystem, FILESYSTEM_DELETE, state.pluginId))
    {
        Debug::Error(L"FolderWindow::CommandPermanentDelete provider rejected delete plugin:{}.", state.pluginId);
        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                          FolderView::OverlaySeverity::Error,
                                          LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                          LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
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
    bool ok                   = true;
    bool sameFolder           = false;
    bool contextsDiffer       = false;
    bool destinationUnsettled = false;
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

    if (ok && ! dest.folderView.IsCurrentFolderEnumerated())
    {
        Debug::Warning(L"FolderWindow::SanityCheckBothPanes rejected an operation while the destination folder is still loading.");
        destinationUnsettled = true;
        ok                   = false;
    }

    if (ok)
    {
        const bool contextSame = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                 NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
        contextsDiffer         = ! contextSame;

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
    else if (ok && ! contextsDiffer && ! CanSameFileSystemOperation(src.fileSystem, operation, src.pluginId))
    {
        Debug::Error(L"FolderWindow::SanityCheckBothPanes provider rejected same-filesystem operation plugin:{} op:{}.",
                     src.pluginId,
                     static_cast<unsigned int>(operation));
        contextsDiffer = true;
        ok             = false;
    }

    if (! ok && _hWnd)
    {
        const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        int messageId              = destinationUnsettled ? IDS_MSG_PANE_OP_DESTINATION_LOADING
                                     : sameFolder         ? IDS_MSG_PANE_OP_REQUIRES_DIFFERENT_FOLDER
                                     : contextsDiffer     ? IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS
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

    const bool waitForOthers                        = _fileOperations->ShouldQueueNewTask();
    const bool contextSame                          = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
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

    const bool waitForOthers                        = _fileOperations->ShouldQueueNewTask();
    const bool contextSame                          = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
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

#ifdef ENABLE_TESTS
    if (FileOperationsSelfTest::IsRunning())
    {
        FileOperationsSelfTest::NotifyTaskCompleted(payload->taskId, payload->hr);
    }
#endif

    FileOperationState::Task* task = _fileOperations->FindTask(payload->taskId);
    if (! task)
    {
        _fileOperationRequestCompletionCallbacks.erase(payload->taskId);
        return 0;
    }

    if (auto completionIt = _fileOperationRequestCompletionCallbacks.find(payload->taskId); completionIt != _fileOperationRequestCompletionCallbacks.end())
    {
        std::function<void(HRESULT)> completion = std::move(completionIt->second);
        _fileOperationRequestCompletionCallbacks.erase(completionIt);
        if (completion)
        {
            completion(payload->hr);
        }
    }

    const Pane sourcePane                     = task->GetSourcePane();
    const std::optional<Pane> destinationPane = task->GetDestinationPane();
    const bool hasResolvedItems               = task->_resolvedItems.size() == task->_sourcePaths.size();

    const auto resolvedDestinationForIndex = [&](size_t index) noexcept -> std::optional<std::filesystem::path>
    {
        if (hasResolvedItems)
        {
            return task->_resolvedItems[index].destinationPath;
        }

        if (destinationPane.has_value() && index < task->_sourcePaths.size())
        {
            return task->GetDestinationFolder() / task->_sourcePaths[index].filename();
        }

        return std::nullopt;
    };

    const auto isResolvedDirectoryShell = [&](size_t index) noexcept -> bool
    { return hasResolvedItems && task->_resolvedItems[index].kind == ResolvedFileOperationItemKind::DirectoryShell; };

    if (! _fileOperationCompletedCallbacks.empty())
    {
        FileOperationCompletedEvent e{};
        e.taskId          = payload->taskId;
        e.operation       = task->GetOperation();
        e.sourcePane      = sourcePane;
        e.destinationPane = destinationPane;
        e.sourcePaths     = task->_sourcePaths;
        {
            std::scoped_lock lock(task->_sourceItemStatusMutex);
            e.itemOutcomes.reserve(task->_sourceItemStatuses.size());
            for (size_t index = 0; index < task->_sourceItemStatuses.size(); ++index)
            {
                if (task->_sourceItemStatuses[index].has_value())
                {
                    e.itemOutcomes.push_back(FileOperationItemOutcome{.sourceIndex = index, .status = task->_sourceItemStatuses[index].value()});
                }
            }
        }
        e.destinationFolder = task->GetDestinationFolder();
        if (destinationPane.has_value() && (task->GetOperation() == FILESYSTEM_COPY || task->GetOperation() == FILESYSTEM_MOVE))
        {
            e.destinationPaths.reserve(task->_sourcePaths.size());
            for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
            {
                if (const auto destination = resolvedDestinationForIndex(index); destination.has_value())
                {
                    e.destinationPaths.push_back(destination.value());
                }
            }
        }
        e.hr = payload->hr;

        // Iterate over a copy: a callback may unsubscribe (or subscribe) while handling the event.
        const auto subscriptions = _fileOperationCompletedCallbacks;
        for (const FileOperationCompletedSubscription& subscription : subscriptions)
        {
            if (subscription.hasLifetimeGuard && subscription.lifetimeGuard.expired())
            {
                continue;
            }
            if (subscription.callback)
            {
                subscription.callback(e);
            }
        }
    }

    PaneState& src            = sourcePane == Pane::Left ? _leftPane : _rightPane;
    DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();

    const auto trimTrailingSeparators = [](std::wstring_view path) noexcept -> std::wstring_view
    {
        while (path.size() > 1u && (path.back() == L'\\' || path.back() == L'/'))
        {
            path.remove_suffix(1);
        }
        return path;
    };

    const auto foldersEqual = [&](const std::filesystem::path& left, const std::filesystem::path& right) noexcept -> bool
    {
        const std::wstring_view leftText  = trimTrailingSeparators(left.native());
        const std::wstring_view rightText = trimTrailingSeparators(right.native());
        return NavigationLocation::EqualsNoCase(leftText, rightText);
    };

    const auto forceRefreshIfShowingFolder = [&](PaneState& paneState, const std::filesystem::path& folder) noexcept
    {
        if (folder.empty())
        {
            return;
        }

        const auto paneFolder = paneState.folderView.GetFolderPath();
        if (paneFolder.has_value() && foldersEqual(paneFolder.value(), folder))
        {
            paneState.folderView.ForceRefresh();
        }
    };

    const auto forceRefreshVisibleFolder = [&](const std::filesystem::path& folder) noexcept
    {
        forceRefreshIfShowingFolder(_leftPane, folder);
        forceRefreshIfShowingFolder(_rightPane, folder);
    };

    const auto forceRefreshVisibleSourceParents = [&]() noexcept
    {
        for (const auto& sourcePath : task->_sourcePaths)
        {
            const std::filesystem::path parent = sourcePath.parent_path();
            if (! parent.empty())
            {
                forceRefreshVisibleFolder(parent);
            }
        }
    };

    const auto notifyDestinationParentsChanged = [&](PaneState& dst) noexcept
    {
        if (! dst.fileSystem)
        {
            return;
        }

        for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
        {
            const auto destination = resolvedDestinationForIndex(index);
            if (! destination.has_value())
            {
                continue;
            }

            const std::filesystem::path parent = destination->parent_path();
            if (! parent.empty())
            {
                cache.NotifyFolderContentsChanged(dst.fileSystem.get(), parent);
                forceRefreshVisibleFolder(parent);
            }
        }
    };

    const auto forceRefreshPane = [&](PaneState& paneState)
    {
        const auto folder = paneState.folderView.GetFolderPath();
        if (! paneState.fileSystem || ! folder.has_value() || ! cache.IsFolderWatched(paneState.fileSystem.get(), folder.value()))
        {
            paneState.folderView.ForceRefresh();
        }
    };

    if (SUCCEEDED(payload->hr))
    {
        PaneState* dst         = destinationPane.has_value() ? &(destinationPane.value() == Pane::Left ? _leftPane : _rightPane) : nullptr;
        const bool sameContext = dst != nullptr && CompareStringOrdinal(task->_sourcePluginId.c_str(), -1, dst->pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                 NavigationLocation::EqualsNoCase(src.instanceContext, dst->instanceContext);

        switch (task->GetOperation())
        {
            case FILESYSTEM_COPY:
                if (dst != nullptr && dst->fileSystem)
                {
                    cache.NotifyFolderContentsChanged(dst->fileSystem.get(), task->GetDestinationFolder());
                    notifyDestinationParentsChanged(*dst);
                }
                forceRefreshVisibleFolder(task->GetDestinationFolder());
                break;
            case FILESYSTEM_MOVE:
                if (sameContext && src.fileSystem)
                {
                    for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
                    {
                        if (isResolvedDirectoryShell(index))
                        {
                            continue;
                        }
                        if (const auto destination = resolvedDestinationForIndex(index); destination.has_value())
                        {
                            cache.NotifyPathMoved(src.fileSystem.get(), task->_sourcePaths[index], destination.value());
                        }
                    }
                }
                else
                {
                    if (task->_fileSystem)
                    {
                        for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
                        {
                            if (! isResolvedDirectoryShell(index))
                            {
                                cache.NotifyPathDeleted(task->_fileSystem.get(), task->_sourcePaths[index]);
                            }
                        }
                    }
                    if (dst != nullptr && dst->fileSystem)
                    {
                        cache.NotifyFolderContentsChanged(dst->fileSystem.get(), task->GetDestinationFolder());
                        notifyDestinationParentsChanged(*dst);
                    }
                }
                forceRefreshVisibleSourceParents();
                forceRefreshVisibleFolder(task->GetDestinationFolder());
                break;
            case FILESYSTEM_DELETE:
                if (src.fileSystem)
                {
                    for (const auto& sourcePath : task->_sourcePaths)
                    {
                        cache.NotifyPathDeleted(src.fileSystem.get(), sourcePath);
                    }
                }
                forceRefreshVisibleSourceParents();
                break;
            case FILESYSTEM_RENAME:
                forceRefreshPane(src);
                if (dst != nullptr)
                {
                    forceRefreshPane(*dst);
                }
                break;
            default:
                forceRefreshPane(src);
                if (dst != nullptr)
                {
                    forceRefreshPane(*dst);
                }
                break;
        }
    }
    else
    {
        forceRefreshPane(src);
        if (destinationPane.has_value())
        {
            PaneState& dst = destinationPane.value() == Pane::Left ? _leftPane : _rightPane;
            forceRefreshPane(dst);
        }
    }

    const bool autoDismissSuccess = _fileOperations->GetAutoDismissSuccess();
    _fileOperations->RemoveTask(payload->taskId);
    if (autoDismissSuccess && IsAutoDismissableFileOperationCompletion(payload->hr, payload->warningCount, payload->errorCount))
    {
        _fileOperations->DismissCompletedTask(payload->taskId);
    }
    return 0;
}
