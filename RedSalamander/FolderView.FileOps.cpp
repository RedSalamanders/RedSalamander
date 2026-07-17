#include "FolderViewInternal.h"

#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"
#include "SelfTest/Common/SelfTestLatencyHooks.h"
#endif

namespace
{
#ifdef ENABLE_TESTS
std::mutex g_debugMoveSelectedItemsDestinationMutex;
std::optional<std::filesystem::path> g_debugMoveSelectedItemsDestination;
std::atomic_bool g_debugDirectFileOperationFallbackEnabled{false};

[[nodiscard]] std::optional<std::filesystem::path> TakeDebugMoveSelectedItemsDestinationForSelfTest() noexcept
{
    std::scoped_lock lock(g_debugMoveSelectedItemsDestinationMutex);
    std::optional<std::filesystem::path> destination = std::move(g_debugMoveSelectedItemsDestination);
    g_debugMoveSelectedItemsDestination.reset();
    return destination;
}
#endif

constexpr HRESULT kMissingFileOperationHostHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);

[[nodiscard]] bool OpenClipboardWithRetriesForFolderView(HWND ownerWindow) noexcept
{
    constexpr uint32_t kMaxAttempts = 6u;
    constexpr DWORD kRetryWaitMs    = 5u;
    for (uint32_t attempt = 0; attempt < kMaxAttempts; ++attempt)
    {
        if (OpenClipboard(ownerWindow) != 0)
        {
            return true;
        }

        if (attempt + 1u < kMaxAttempts)
        {
            Sleep(kRetryWaitMs);
        }
    }

    return false;
}

[[nodiscard]] bool IsDirectFileOperationFallbackEnabledForSelfTest() noexcept
{
#ifdef ENABLE_TESTS
    return g_debugDirectFileOperationFallbackEnabled.load(std::memory_order_acquire);
#else
    return false;
#endif
}

[[nodiscard]] bool IsBuiltinFileSystemPlugin(std::wstring_view pluginId) noexcept
{
    return CompareStringOrdinal(pluginId.data(), static_cast<int>(pluginId.size()), L"builtin/file-system", -1, TRUE) == CSTR_EQUAL;
}

void ShowClipboardOverlay(FolderView& view, UINT titleStringId, UINT messageStringId, FolderView::OverlaySeverity severity, HRESULT hr = S_OK) noexcept
{
    Debug::Perf::Scope perf(L"clipboard.feedback_us");
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, titleStringId);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, severity == FolderView::OverlaySeverity::Error ? IDS_CAPTION_ERROR : IDS_CAPTION_WARNING);
    }

    std::wstring message = LoadStringResource(nullptr, messageStringId);
    if (message.empty())
    {
        message = title;
    }

    view.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, severity, std::move(title), std::move(message), hr, true, false);
}

void ShowClipboardFormattedOverlay(FolderView& view, UINT titleStringId, UINT messageStringId, const std::filesystem::path& path, HRESULT hr) noexcept
{
    Debug::Perf::Scope perf(L"clipboard.feedback_us");
    perf.SetHr(hr);

    std::wstring title = LoadStringResource(nullptr, titleStringId);
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
    }

    std::wstring message = FormatStringResource(nullptr, messageStringId, path.filename().wstring(), static_cast<unsigned long>(static_cast<uint32_t>(hr)));
    if (message.empty())
    {
        message = title;
    }

    view.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message), hr, true, false);
}

[[nodiscard]] wil::unique_hglobal BuildFileDropHGlobal(const std::vector<std::filesystem::path>& paths) noexcept
{
    size_t totalChars = 1u;
    for (const auto& path : paths)
    {
        const size_t pathChars = path.native().size();
        if (pathChars == (std::numeric_limits<size_t>::max)() || totalChars > (std::numeric_limits<size_t>::max)() - pathChars - 1u)
        {
            return nullptr;
        }
        totalChars += pathChars + 1u;
    }

    if (totalChars > ((std::numeric_limits<size_t>::max)() - sizeof(DROPFILES)) / sizeof(wchar_t))
    {
        return nullptr;
    }
    const size_t bytes = sizeof(DROPFILES) + totalChars * sizeof(wchar_t);
    wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! memory)
    {
        return nullptr;
    }

    auto* drop = static_cast<DROPFILES*>(GlobalLock(memory.get()));
    if (! drop)
    {
        return nullptr;
    }

    drop->pFiles = sizeof(DROPFILES);
    drop->pt     = POINT{};
    drop->fNC    = FALSE;
    drop->fWide  = TRUE;

    auto* cursor = reinterpret_cast<wchar_t*>(reinterpret_cast<BYTE*>(drop) + drop->pFiles);
    for (const auto& path : paths)
    {
        const std::wstring text = path.native();
        std::copy(text.begin(), text.end(), cursor);
        cursor += text.size();
        *cursor++ = L'\0';
    }
    *cursor = L'\0';
    GlobalUnlock(memory.get());

    return memory;
}

[[nodiscard]] wil::unique_hglobal BuildPreferredDropEffectHGlobal(DWORD preferredEffect) noexcept
{
    wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, sizeof(DWORD)));
    if (! memory)
    {
        return nullptr;
    }

    auto* effect = static_cast<DWORD*>(GlobalLock(memory.get()));
    if (! effect)
    {
        return nullptr;
    }

    *effect = preferredEffect;
    GlobalUnlock(memory.get());
    return memory;
}

[[nodiscard]] bool SetFileDropClipboard(HWND ownerWindow, const std::vector<std::filesystem::path>& paths, DWORD preferredEffect) noexcept
{
    wil::unique_hglobal drop = BuildFileDropHGlobal(paths);
    if (! drop)
    {
        return false;
    }

    wil::unique_hglobal effect = BuildPreferredDropEffectHGlobal(preferredEffect);
    if (! effect)
    {
        return false;
    }

    if (! OpenClipboardWithRetriesForFolderView(ownerWindow))
    {
        Debug::Warning(L"FolderView::SetFileDropClipboard: OpenClipboard failed (error={}).", GetLastError());
        return false;
    }
    const auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });

    if (EmptyClipboard() == 0)
    {
        Debug::Warning(L"FolderView::SetFileDropClipboard: EmptyClipboard failed (error={}).", GetLastError());
        return false;
    }

    if (SetClipboardData(CF_HDROP, drop.get()) == nullptr)
    {
        Debug::Warning(L"FolderView::SetFileDropClipboard: SetClipboardData(CF_HDROP) failed (error={}).", GetLastError());
        return false;
    }
    drop.release();

    const UINT preferredDropEffectFormat = PreferredDropEffectFormat();
    if (preferredDropEffectFormat == 0u || SetClipboardData(preferredDropEffectFormat, effect.get()) == nullptr)
    {
        Debug::Warning(L"FolderView::SetFileDropClipboard: SetClipboardData(Preferred DropEffect) failed (format={}, error={}).",
                       preferredDropEffectFormat,
                       GetLastError());
        return false;
    }
    effect.release();

    return true;
}

[[nodiscard]] std::vector<std::filesystem::path> ReadFileDropClipboard(HWND ownerWindow) noexcept
{
    constexpr UINT kMaxClipboardDropPaths = 65'536u;
    constexpr UINT kMaxClipboardPathChars = 32'768u;
    constexpr size_t kMaxClipboardTotalChars = 16u * 1024u * 1024u;

    std::vector<std::filesystem::path> result;
    if (! OpenClipboardWithRetriesForFolderView(ownerWindow))
    {
        return result;
    }
    const auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });

    HANDLE handle = GetClipboardData(CF_HDROP);
    if (! handle)
    {
        return result;
    }

    const auto fileCount = DragQueryFileW(static_cast<HDROP>(handle), 0xFFFFFFFFu, nullptr, 0u);
    if (fileCount > kMaxClipboardDropPaths)
    {
        return result;
    }
    result.reserve(fileCount);
    size_t totalPathChars = 0u;
    for (UINT index = 0; index < fileCount; ++index)
    {
        const UINT length = DragQueryFileW(static_cast<HDROP>(handle), index, nullptr, 0u);
        if (length == 0u || length > kMaxClipboardPathChars || totalPathChars > kMaxClipboardTotalChars - static_cast<size_t>(length))
        {
            result.clear();
            return result;
        }

        std::wstring path(static_cast<size_t>(length) + 1u, L'\0');
        if (DragQueryFileW(static_cast<HDROP>(handle), index, path.data(), length + 1u) == length)
        {
            path.resize(length);
            result.emplace_back(path);
            totalPathChars += length;
        }
    }

    return result;
}

[[nodiscard]] std::optional<DWORD> ReadPreferredDropEffectClipboard(HWND ownerWindow) noexcept
{
    if (! OpenClipboardWithRetriesForFolderView(ownerWindow))
    {
        return std::nullopt;
    }
    const auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });

    const UINT format = PreferredDropEffectFormat();
    if (format == 0u)
    {
        return std::nullopt;
    }

    HANDLE handle = GetClipboardData(format);
    if (! handle)
    {
        return std::nullopt;
    }

    auto* effect = static_cast<DWORD*>(GlobalLock(handle));
    if (! effect)
    {
        return std::nullopt;
    }
    const auto unlock = wil::scope_exit([&] { GlobalUnlock(handle); });
    if (GlobalSize(handle) < sizeof(DWORD))
    {
        return std::nullopt;
    }

    const DWORD result = *effect;
    return result;
}

void InvalidateMoveClipboardAfterVerifiedCompletion(DWORD expectedSequenceNumber) noexcept
{
    if (expectedSequenceNumber == 0u || GetClipboardSequenceNumber() != expectedSequenceNumber || ! OpenClipboardWithRetriesForFolderView(nullptr))
    {
        return;
    }
    const auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });
    if (GetClipboardSequenceNumber() != expectedSequenceNumber)
    {
        return;
    }
    if (EmptyClipboard() == 0)
    {
        Debug::Warning(L"FolderView: failed to invalidate completed MOVE clipboard (error={}).", GetLastError());
    }
}

[[nodiscard]] bool IsShortcutSlotCollision(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) || hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
}

[[nodiscard]] HRESULT CreateShellShortcut(const std::filesystem::path& target,
                                          const std::filesystem::path& destinationFolder,
                                          std::filesystem::path& createdPath) noexcept
{
    createdPath.clear();

    wil::com_ptr<IShellLinkW> shellLink;
    HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(shellLink.addressof()));
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring targetPath = BuildShellPersistSavePath(target);
    hr                            = shellLink->SetPath(targetPath.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    const std::filesystem::path parent = target.parent_path();
    if (! parent.empty())
    {
        static_cast<void>(shellLink->SetWorkingDirectory(parent.c_str()));
    }

    const std::wstring description = target.filename().wstring();
    if (! description.empty())
    {
        static_cast<void>(shellLink->SetDescription(description.c_str()));
    }

    wil::com_ptr<IPersistFile> persist;
    hr = shellLink->QueryInterface(IID_PPV_ARGS(persist.addressof()));
    if (FAILED(hr))
    {
        return hr;
    }

    for (int attempt = 0; attempt < 256; ++attempt)
    {
        const std::filesystem::path candidate = GenerateShortcutPath(destinationFolder, target, attempt);
        bool exists                          = false;
        const HRESULT existsHr               = QueryShellShortcutExactPathExists(candidate, exists);
        if (FAILED(existsHr))
        {
            return existsHr;
        }

        if (! exists)
        {
            createdPath = candidate;
#ifdef ENABLE_TESTS
            SelfTestLatency::Consume(SelfTestLatency::Point::PasteShortcutAfterSlotProbe);
#endif
            hr = SaveShellShortcutExactPath(persist.get(), createdPath);
            if (SUCCEEDED(hr))
            {
                return S_OK;
            }

            if (IsShortcutSlotCollision(hr))
            {
                createdPath.clear();
                continue;
            }

            return hr;
        }
    }

    return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
}
} // namespace

#ifdef ENABLE_TESTS
void FolderView::DebugSetNextMoveSelectedItemsDestinationForSelfTest(std::optional<std::filesystem::path> destination) noexcept
{
    std::scoped_lock lock(g_debugMoveSelectedItemsDestinationMutex);
    g_debugMoveSelectedItemsDestination = std::move(destination);
}

void FolderView::DebugSetDirectFileOperationFallbackEnabledForSelfTest(bool enabled) noexcept
{
    g_debugDirectFileOperationFallbackEnabled.store(enabled, std::memory_order_release);
}

bool FolderView::DebugIsDirectFileOperationFallbackEnabledForSelfTest() noexcept
{
    return IsDirectFileOperationFallbackEnabledForSelfTest();
}
#endif

void FolderView::CommandRename()
{
    RenameFocusedItem();
}

void FolderView::CommandView()
{
    static_cast<void>(RequestViewFocusedItem(ViewFileRole::Primary, true));
}

bool FolderView::CommandAlternateView()
{
    return RequestViewFocusedItem(ViewFileRole::Alternate, false);
}

bool FolderView::CommandViewWith(std::wstring_view actionId)
{
    if (actionId.empty())
    {
        return false;
    }

    return RequestViewFocusedItem(ViewFileRole::Primary, false, actionId);
}

bool FolderView::RequestViewFocusedItem(ViewFileRole role, bool activateFallback, std::wstring_view actionId)
{
    if (_focusedIndex == static_cast<size_t>(-1) || _focusedIndex >= _items.size())
    {
        return false;
    }

    const auto& item = _items[_focusedIndex];
    if (item.isDirectory)
    {
        if (activateFallback)
        {
            SetFolderPath(GetItemFullPath(item));
            return true;
        }
        return false;
    }

    bool handled = false;
    if (_viewFileRequestCallback)
    {
        ViewFileRequest request;
        request.role        = role;
        request.actionId    = std::wstring(actionId);
        request.focusedPath = GetItemFullPath(item);

        for (const auto& candidate : _items)
        {
            if (candidate.isDirectory)
            {
                continue;
            }

            request.displayedFilePaths.push_back(GetItemFullPath(candidate));
            if (candidate.selected)
            {
                request.selectionPaths.push_back(GetItemFullPath(candidate));
            }
        }

        handled = _viewFileRequestCallback(request);
    }

    if (! handled && activateFallback)
    {
        ActivateFocusedItem();
        handled = true;
    }

    return handled;
}

void FolderView::CommandDelete()
{
    if (_hWnd)
    {
        SetFocus(_hWnd.get());
        const HWND root = GetAncestor(_hWnd.get(), GA_ROOT);
        if (root && PostMessageW(root, WM_COMMAND, MAKEWPARAM(IDM_PANE_DELETE, 0), 0) != 0)
        {
            return;
        }
    }

    DeleteSelectedItems();
}

HRESULT FolderView::CopySelectedItemsToFolder(const std::filesystem::path& destinationFolder)
{
    if (! _fileSystem)
    {
        return E_POINTER;
    }

    if (destinationFolder.empty())
    {
        return E_INVALIDARG;
    }

    std::vector<std::filesystem::path> paths;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            paths.push_back(GetItemFullPath(item));
        }
    }

    if (paths.empty() && _focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        paths.push_back(GetItemFullPath(_items[_focusedIndex]));
    }

    if (paths.empty())
    {
        return S_FALSE;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
    if (_fileOperationRequestCallback)
    {
        FileOperationRequest request{};
        request.operation         = FILESYSTEM_COPY;
        request.sourcePaths       = std::move(paths);
        request.destinationFolder = destinationFolder;
        request.flags             = flags;

        const HRESULT hrStart = _fileOperationRequestCallback(std::move(request));
        if (FAILED(hrStart))
        {
            ReportError(L"Copy", hrStart);
            return hrStart;
        }

        return hrStart;
    }

    if (! IsDirectFileOperationFallbackEnabledForSelfTest())
    {
        Debug::Error(L"FolderView::CopySelectedItemsToFolder missing FileOperationRequestCallback; refusing direct file-operation fallback");
        ReportError(L"Copy", kMissingFileOperationHostHr);
        return kMissingFileOperationHostHr;
    }

    if (! ConfirmNonRevertableFileOperation(_hWnd.get(), _fileSystem.get(), FILESYSTEM_COPY, paths, destinationFolder))
    {
        return S_FALSE;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** sourcePaths = nullptr;
    unsigned long count         = 0;
    HRESULT hr                  = BuildPathArrayArena(paths, arenaOwner, &sourcePaths, &count);
    if (FAILED(hr))
    {
        ReportError(L"Copy", hr);
        return hr;
    }

    hr = _fileSystem->CopyItems(sourcePaths, count, destinationFolder.c_str(), flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Copy", hr);
        return hr;
    }

    return S_OK;
}

HRESULT FolderView::MoveSelectedItemsToFolder(const std::filesystem::path& destinationFolder)
{
    if (! _fileSystem)
    {
        return E_POINTER;
    }

    if (destinationFolder.empty())
    {
        return E_INVALIDARG;
    }

    std::vector<std::filesystem::path> paths;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            paths.push_back(GetItemFullPath(item));
        }
    }

    if (paths.empty() && _focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        paths.push_back(GetItemFullPath(_items[_focusedIndex]));
    }

    if (paths.empty())
    {
        return S_FALSE;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
    if (_fileOperationRequestCallback)
    {
        FileOperationRequest request{};
        request.operation         = FILESYSTEM_MOVE;
        request.sourcePaths       = std::move(paths);
        request.destinationFolder = destinationFolder;
        request.flags             = flags;

        const HRESULT hrStart = _fileOperationRequestCallback(std::move(request));
        if (FAILED(hrStart))
        {
            ReportError(L"Move", hrStart);
            return hrStart;
        }

        return hrStart;
    }

    if (! IsDirectFileOperationFallbackEnabledForSelfTest())
    {
        Debug::Error(L"FolderView::MoveSelectedItemsToFolder missing FileOperationRequestCallback; refusing direct file-operation fallback");
        ReportError(L"Move", kMissingFileOperationHostHr);
        return kMissingFileOperationHostHr;
    }

    if (! ConfirmNonRevertableFileOperation(_hWnd.get(), _fileSystem.get(), FILESYSTEM_MOVE, paths, destinationFolder))
    {
        return S_FALSE;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** sourcePaths = nullptr;
    unsigned long count         = 0;
    HRESULT hr                  = BuildPathArrayArena(paths, arenaOwner, &sourcePaths, &count);
    if (FAILED(hr))
    {
        ReportError(L"Move", hr);
        return hr;
    }

    hr = _fileSystem->MoveItems(sourcePaths, count, destinationFolder.c_str(), flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Move", hr);
        return hr;
    }

    return S_OK;
}

void FolderView::DeleteSelectedItems()
{
    if (! _fileSystem)
    {
        return;
    }

    std::vector<std::filesystem::path> paths = GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        return;
    }

    if (_fileOperationRequestCallback)
    {
        FileOperationRequest request{};
        request.operation   = FILESYSTEM_DELETE;
        request.sourcePaths = std::move(paths);
        request.flags       = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN);

        const HRESULT hrStart = _fileOperationRequestCallback(std::move(request));
        if (FAILED(hrStart))
        {
            ReportError(L"Delete", hrStart);
        }
        return;
    }

    if (! IsDirectFileOperationFallbackEnabledForSelfTest())
    {
        Debug::Error(L"FolderView::DeleteSelectedItems missing FileOperationRequestCallback; refusing direct file-operation fallback");
        ReportError(L"Delete", kMissingFileOperationHostHr);
        return;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** pathArray = nullptr;
    unsigned long count       = 0;
    HRESULT hr                = BuildPathArrayArena(paths, arenaOwner, &pathArray, &count);
    if (FAILED(hr))
    {
        ReportError(L"Delete", hr);
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN);
    hr                          = _fileSystem->DeleteItems(pathArray, count, flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Delete", hr);
        return;
    }

    DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
    for (const auto& path : paths)
    {
        cache.NotifyPathDeleted(_fileSystem.get(), path);
    }
    if (! _currentFolder || ! cache.IsFolderWatched(_fileSystem.get(), _currentFolder.value()))
    {
        ForceRefresh();
    }
}

void FolderView::CopySelectionToClipboard()
{
    std::vector<std::filesystem::path> paths;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            paths.push_back(GetItemFullPath(item));
        }
    }

    if (paths.empty() && _focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        paths.push_back(GetItemFullPath(_items[_focusedIndex]));
    }

    if (paths.empty())
        return;

    static_cast<void>(SetFileDropClipboard(_hWnd.get(), paths, DROPEFFECT_COPY));
}

bool FolderView::CutSelectionToClipboard()
{
    Debug::Perf::Scope perf(L"clipboard.cut_us");

    if (! IsBuiltinFileSystemPlugin(_fileSystemPluginId))
    {
        ShowClipboardOverlay(*this, IDS_CMD_CLIPBOARD_CUT, IDS_MSG_CLIPBOARD_LOCAL_SELECTION_REQUIRED, OverlaySeverity::Warning);
        return false;
    }

    const std::vector<std::filesystem::path> paths = GetSelectedOrFocusedPaths();
    perf.SetValue0(static_cast<uint64_t>(paths.size()));
    if (paths.empty())
    {
        ShowClipboardOverlay(*this, IDS_CMD_CLIPBOARD_CUT, IDS_MSG_CLIPBOARD_SELECTION_REQUIRED, OverlaySeverity::Warning);
        return false;
    }

    if (! SetFileDropClipboard(_hWnd.get(), paths, DROPEFFECT_MOVE))
    {
        ShowClipboardOverlay(*this, IDS_CMD_CLIPBOARD_CUT, IDS_MSG_CLIPBOARD_WRITE_FAILED, OverlaySeverity::Warning);
        return false;
    }

    return true;
}

void FolderView::PasteItemsFromClipboard()
{
    std::vector<std::filesystem::path> sources = ReadFileDropClipboard(_hWnd.get());
    if (sources.empty())
    {
        return;
    }

    if (! _currentFolder || ! _fileSystem)
    {
        return;
    }
    if (! IsCurrentFolderEnumerated())
    {
        return;
    }

    const DWORD preferredDropEffect     = ReadPreferredDropEffectClipboard(_hWnd.get()).value_or(DROPEFFECT_COPY);
    const bool moveRequested            = preferredDropEffect == DROPEFFECT_MOVE;
    const FileSystemOperation operation = moveRequested ? FILESYSTEM_MOVE : FILESYSTEM_COPY;
    const DWORD clipboardSequenceNumber = moveRequested ? GetClipboardSequenceNumber() : 0u;

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"FolderView::PasteItemsFromClipboard sources={} preferredDropEffect=0x{:X} operation={} callback={}",
                                              sources.size(),
                                              static_cast<unsigned>(preferredDropEffect),
                                              operation == FILESYSTEM_MOVE ? L"move" : L"copy",
                                              _fileOperationRequestCallback ? L"yes" : L"no"));
#endif

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
    if (_fileOperationRequestCallback)
    {
        FileOperationRequest request{};
        request.operation         = operation;
        request.sourcePaths       = std::move(sources);
        request.destinationFolder = _currentFolder.value();
        request.flags             = flags;
        if (moveRequested)
        {
            request.completionCallback = [clipboardSequenceNumber](HRESULT completionHr) noexcept
            {
                if (SUCCEEDED(completionHr))
                {
                    InvalidateMoveClipboardAfterVerifiedCompletion(clipboardSequenceNumber);
                }
            };
        }

        const HRESULT hrStart = _fileOperationRequestCallback(std::move(request));
        if (FAILED(hrStart))
        {
            ReportError(moveRequested ? L"Move" : L"Copy", hrStart);
        }
        return;
    }

    if (! IsDirectFileOperationFallbackEnabledForSelfTest())
    {
        Debug::Error(L"FolderView::PasteItemsFromClipboard missing FileOperationRequestCallback; refusing direct file-operation fallback");
        ReportError(moveRequested ? L"Move" : L"Copy", kMissingFileOperationHostHr);
        return;
    }

    if (! ConfirmNonRevertableFileOperation(_hWnd.get(), _fileSystem.get(), operation, sources, _currentFolder.value()))
    {
        return;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** paths = nullptr;
    unsigned long count   = 0;
    HRESULT hr            = BuildPathArrayArena(sources, arenaOwner, &paths, &count);
    if (FAILED(hr))
    {
        ReportError(moveRequested ? L"Move" : L"Copy", hr);
        return;
    }

    hr = moveRequested ? _fileSystem->MoveItems(paths, count, _currentFolder->c_str(), flags, nullptr, nullptr, nullptr)
                       : _fileSystem->CopyItems(paths, count, _currentFolder->c_str(), flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(moveRequested ? L"Move" : L"Copy", hr);
        return;
    }

    DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
    if (moveRequested)
    {
        for (const auto& source : sources)
        {
            const std::filesystem::path parent = source.parent_path();
            if (! parent.empty())
            {
                cache.NotifyFolderContentsChanged(_fileSystem.get(), parent);
            }
        }
    }
    cache.NotifyFolderContentsChanged(_fileSystem.get(), _currentFolder.value());
    if (moveRequested)
    {
        InvalidateMoveClipboardAfterVerifiedCompletion(clipboardSequenceNumber);
    }
    if (! _currentFolder || ! cache.IsFolderWatched(_fileSystem.get(), _currentFolder.value()))
    {
        ForceRefresh();
    }
}

bool FolderView::PasteShortcutFromClipboard()
{
    Debug::Perf::Scope perf(L"clipboard.paste_shortcut_us");
    RecoverStalePasteShortcutWork();

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"PasteShortcutFromClipboard enter hwnd=0x{:X} focus=0x{:X} plugin='{}' currentFolder='{}'",
                                              reinterpret_cast<uintptr_t>(_hWnd.get()),
                                              reinterpret_cast<uintptr_t>(GetFocus()),
                                              _fileSystemPluginId,
                                              _currentFolder.has_value() ? _currentFolder->wstring() : std::wstring()));
#endif

    if (! _currentFolder.has_value() || _currentFolder.value().empty() || ! IsBuiltinFileSystemPlugin(_fileSystemPluginId))
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"PasteShortcutFromClipboard blocked: local built-in folder required");
#endif
        ShowClipboardOverlay(*this, IDS_CMD_CLIPBOARD_PASTE_SHORTCUT, IDS_MSG_CLIPBOARD_LOCAL_FOLDER_REQUIRED, OverlaySeverity::Warning);
        return false;
    }

    std::vector<std::filesystem::path> sources = ReadFileDropClipboard(_hWnd.get());
    const uint64_t sourceCount = static_cast<uint64_t>(sources.size());
    perf.SetValue0(sourceCount);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(
        std::format(L"PasteShortcutFromClipboard sources={} first='{}'", sources.size(), ! sources.empty() ? sources.front().wstring() : std::wstring()));
#endif
    if (sources.empty())
    {
        ShowClipboardOverlay(*this, IDS_CMD_CLIPBOARD_PASTE_SHORTCUT, IDS_MSG_CLIPBOARD_NO_SHORTCUT_SOURCE, OverlaySeverity::Warning);
        return false;
    }

    PasteShortcutRequest request{};
    request.fileSystem   = _fileSystem;
    request.requestId    = _nextPasteShortcutRequestId++;
    request.sources      = std::move(sources);
    request.targetFolder = _currentFolder.value();
    request.generation   = _folderPathGeneration;
#ifdef ENABLE_TESTS
    constexpr const wchar_t* kFailCompletionPostEnv = L"REDSALAMANDER_PASTE_SHORTCUT_FAIL_COMPLETION_POST";
    std::array<wchar_t, 8> failCompletionPostValue{};
    request.failFirstCompletionPostForSelfTest =
        ::GetEnvironmentVariableW(kFailCompletionPostEnv, failCompletionPostValue.data(), static_cast<DWORD>(failCompletionPostValue.size())) > 0u;
    if (request.failFirstCompletionPostForSelfTest)
    {
        static_cast<void>(::SetEnvironmentVariableW(kFailCompletionPostEnv, nullptr));
    }
#endif

    if (_pasteShortcutInFlight)
    {
        _pendingPasteShortcutRequests.push_back(std::move(request));
        perf.SetValue1(2u);
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(std::format(L"PasteShortcutFromClipboard pending sources={} generation={} queueDepth={}",
                                                  sourceCount,
                                                  _folderPathGeneration,
                                                  _pendingPasteShortcutRequests.size()));
#endif
        return true;
    }

    if (! StartPasteShortcutWork(std::move(request)))
    {
        return false;
    }

    perf.SetValue1(1u);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"PasteShortcutFromClipboard queued sources={} generation={}",
                                              sourceCount,
                                              _folderPathGeneration));
#endif
    return true;
}

bool FolderView::StartPasteShortcutWork(PasteShortcutRequest request)
{
    struct PasteShortcutWork
    {
        PasteShortcutWork() noexcept                              = default;
        PasteShortcutWork(const PasteShortcutWork&)               = delete;
        PasteShortcutWork& operator=(const PasteShortcutWork&)    = delete;
        PasteShortcutWork(PasteShortcutWork&&)                    = delete;
        PasteShortcutWork& operator=(PasteShortcutWork&&)         = delete;

        HWND hwnd = nullptr;
        PasteShortcutRequest request;
        PasteShortcutResult result;

        void Execute() noexcept
        {
            const auto startedAt = std::chrono::steady_clock::now();
            result.createdLinks.reserve(request.sources.size());

            const HRESULT coinitHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
            if (FAILED(coinitHr))
            {
                result.firstFailure = coinitHr;
                if (! request.sources.empty())
                {
                    result.failedSource = request.sources.front();
                }
            }
            else
            {
                [[maybe_unused]] const wil::unique_couninitialize_call coUninit;
                for (const auto& source : request.sources)
                {
#ifdef ENABLE_TESTS
                    SelfTestLatency::Consume(SelfTestLatency::Point::PasteShortcutSave);
#endif
                    std::filesystem::path createdPath;
#ifdef ENABLE_TESTS
                    HRESULT hr = SelfTestLatency::ConsumeFailure(SelfTestLatency::Point::PasteShortcutSave);
                    if (SUCCEEDED(hr))
                    {
                        hr = CreateShellShortcut(source, request.targetFolder, createdPath);
                    }
#else
                    const HRESULT hr = CreateShellShortcut(source, request.targetFolder, createdPath);
#endif
                    if (FAILED(hr))
                    {
                        result.firstFailure = hr;
                        result.failedSource = source;
                        break;
                    }

                    if (! createdPath.empty())
                    {
                        result.createdLinks.push_back(std::move(createdPath));
                    }
                }
            }

            result.elapsedUs = Debug::Perf::ElapsedUs(startedAt);
            auto recoveryPayload = std::make_unique<PasteShortcutResult>(result);
            auto payload         = std::make_unique<PasteShortcutResult>(std::move(result));
#ifdef ENABLE_TESTS
            const bool forceFirstPostFailure = request.failFirstCompletionPostForSelfTest;
#else
            constexpr bool forceFirstPostFailure = false;
#endif
            const bool posted = ! forceFirstPostFailure && PostMessagePayload(hwnd, WndMsg::kFolderViewPasteShortcutComplete, 0, std::move(payload));
            if (! posted)
            {
                const DWORD firstPostError = forceFirstPostFailure ? ERROR_NOT_ENOUGH_MEMORY : ::GetLastError();
                if (! PostMessagePayload(hwnd, WndMsg::kFolderViewPasteShortcutComplete, 0, std::move(recoveryPayload)))
                {
                    Debug::Error(L"Paste Shortcut completion delivery failed twice (firstError={}, retryError={}).",
                                 firstPostError,
                                 ::GetLastError());
                }
                else
                {
                    Debug::Warning(L"Paste Shortcut completion delivery recovered after PostMessage failure (error={}).", firstPostError);
                }
            }
        }
    };

    auto work                 = std::make_unique<PasteShortcutWork>();
    work->hwnd                = _hWnd.get();
    work->request             = std::move(request);
    work->result.fileSystem   = work->request.fileSystem;
    work->result.requestId    = work->request.requestId;
    work->result.generation   = work->request.generation;
    work->result.targetFolder = work->request.targetFolder;

    const std::filesystem::path firstSource = ! work->request.sources.empty() ? work->request.sources.front() : std::filesystem::path{};
    const uint64_t requestId = work->request.requestId;

    if (! SubmitOwnedThreadpoolCallback(work))
    {
        const HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
        ShowClipboardFormattedOverlay(*this, IDS_CMD_CLIPBOARD_PASTE_SHORTCUT, IDS_FMT_CLIPBOARD_PASTE_SHORTCUT_FAILED, firstSource, hr);
        return false;
    }

    _pasteShortcutInFlight = true;
    _activePasteShortcutRequestId = requestId;
    _pasteShortcutStartedAt = std::chrono::steady_clock::now();
    return true;
}

void FolderView::RecoverStalePasteShortcutWork()
{
    constexpr auto kStaleTimeout = std::chrono::seconds(30);
    if (! _pasteShortcutInFlight || _pasteShortcutStartedAt == std::chrono::steady_clock::time_point{} ||
        std::chrono::steady_clock::now() - _pasteShortcutStartedAt < kStaleTimeout)
    {
        return;
    }

    Debug::Warning(L"Paste Shortcut completion timed out; releasing stale request {} and continuing queued work.", _activePasteShortcutRequestId);
    _pasteShortcutInFlight = false;
    _activePasteShortcutRequestId = 0u;
    _pasteShortcutStartedAt = {};
    StartNextPasteShortcutRequest();
}

void FolderView::StartNextPasteShortcutRequest()
{
    if (_pasteShortcutInFlight)
    {
        return;
    }

    while (! _pendingPasteShortcutRequests.empty())
    {
        PasteShortcutRequest next = std::move(_pendingPasteShortcutRequests.front());
        _pendingPasteShortcutRequests.pop_front();
        if (StartPasteShortcutWork(std::move(next)))
        {
            return;
        }
    }
}

void FolderView::OnPasteShortcutComplete(PasteShortcutResult result)
{
    Debug::Perf::Emit(L"clipboard.paste_shortcut_worker_us",
                      L"",
                      result.elapsedUs,
                      static_cast<uint64_t>(result.createdLinks.size()),
                      0u,
                      result.firstFailure);

#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(std::format(L"PasteShortcutComplete created={} firstFailure=0x{:08X} target='{}'",
                                              result.createdLinks.size(),
                                              static_cast<unsigned long>(result.firstFailure),
                                              result.targetFolder.wstring()));
#endif

    const bool hasCurrentTarget = _currentFolder && OrdinalString::EqualsNoCasePath(result.targetFolder, _currentFolder.value());
    const bool generationMatches = result.generation == _folderPathGeneration;

    if (result.fileSystem && ! result.createdLinks.empty())
    {
        DirectoryInfoCache::GetInstance().NotifyFolderContentsChanged(result.fileSystem.get(), result.targetFolder);
    }

    if (_fileSystem && result.fileSystem.get() == _fileSystem.get() && ! result.createdLinks.empty() && hasCurrentTarget && generationMatches)
    {
        RememberFocusedItemForFolder(result.targetFolder, result.createdLinks.back().filename().wstring());
        EnumerateFolder();
    }

    if (FAILED(result.firstFailure))
    {
        ShowClipboardFormattedOverlay(*this, IDS_CMD_CLIPBOARD_PASTE_SHORTCUT, IDS_FMT_CLIPBOARD_PASTE_SHORTCUT_FAILED, result.failedSource, result.firstFailure);
    }

    if (result.requestId == _activePasteShortcutRequestId)
    {
        _pasteShortcutInFlight = false;
        _activePasteShortcutRequestId = 0u;
        _pasteShortcutStartedAt = {};
        StartNextPasteShortcutRequest();
    }
}

void FolderView::RenameFocusedItem()
{
    if (_focusedIndex == static_cast<size_t>(-1) || _focusedIndex >= _items.size())
        return;

    if (! _fileSystem)
        return;

    const auto& item                     = _items[_focusedIndex];
    const std::wstring originalName      = std::wstring(item.displayName);
    const bool originalIsDirectory       = item.isDirectory;
    const std::filesystem::path fullPath = GetItemFullPath(item);

    RenamePromptResult prompt = PromptForRename(_hWnd.get(), originalName, originalIsDirectory, _appTheme);
    if (prompt.action == RenamePromptAction::Cancel)
    {
        return;
    }

    if (prompt.action == RenamePromptAction::BatchRename)
    {
        if (_batchRenameRequestCallback)
        {
            // Folders root Batch Rename at the folder; files seed it with the prompted item.
            // The item path is captured before the modal prompt, so a folder refresh that
            // clears the live selection while the prompt is open cannot change the target.
            _batchRenameRequestCallback(fullPath, originalIsDirectory);
        }
        return;
    }

    if (prompt.text.empty())
    {
        return;
    }

    std::filesystem::path target = fullPath.parent_path() / prompt.text;
    const FileSystemFlags flags  = FILESYSTEM_FLAG_NONE;
    const HRESULT hr             = _fileSystem->RenameItem(fullPath.c_str(), target.c_str(), flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Rename", hr);
        return;
    }

    DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
    cache.NotifyPathMoved(_fileSystem.get(), fullPath, target);
    if (! _currentFolder || ! cache.IsFolderWatched(_fileSystem.get(), _currentFolder.value()))
    {
        ForceRefresh();
    }
}

void FolderView::ShowProperties()
{
    if (_focusedIndex == static_cast<size_t>(-1) || _focusedIndex >= _items.size())
        return;

    const auto& item                     = _items[_focusedIndex];
    const std::filesystem::path fullPath = GetItemFullPath(item);

    if (_propertiesRequestCallback)
    {
        const HRESULT hr = _propertiesRequestCallback(fullPath);
        if (FAILED(hr))
        {
            ReportError(L"Properties", hr);
        }
        return;
    }

    SHObjectProperties(_hWnd.get(), SHOP_FILEPATH, fullPath.c_str(), nullptr);
}

void FolderView::MoveSelectedItems()
{
    if (! _fileSystem)
    {
        return;
    }

    std::vector<std::filesystem::path> paths;
    for (const auto& item : _items)
    {
        if (item.selected)
        {
            paths.push_back(GetItemFullPath(item));
        }
    }

    if (paths.empty() && _focusedIndex != static_cast<size_t>(-1) && _focusedIndex < _items.size())
    {
        paths.push_back(GetItemFullPath(_items[_focusedIndex]));
    }

    if (paths.empty())
    {
        return;
    }

    std::filesystem::path destination;
#ifdef ENABLE_TESTS
    if (std::optional<std::filesystem::path> debugDestination = TakeDebugMoveSelectedItemsDestinationForSelfTest(); debugDestination.has_value())
    {
        destination = std::move(debugDestination.value());
        SelfTest::AppendSelfTestTrace(std::format(L"FolderView::MoveSelectedItems debug destination '{}'", destination.wstring()));
    }
    else
#endif
    {
        wil::com_ptr<IFileOpenDialog> dialog;
        if (FAILED(CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.addressof()))))
            return;

        DWORD options = 0;
        if (FAILED(dialog->GetOptions(&options)))
            return;
        dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
        if (FAILED(dialog->Show(_hWnd.get())))
            return;

        wil::com_ptr<IShellItem> result;
        if (FAILED(dialog->GetResult(result.addressof())))
            return;

        wil::unique_cotaskmem_string selectedPath;
        if (FAILED(result->GetDisplayName(SIGDN_FILESYSPATH, selectedPath.addressof())))
            return;

        destination = std::filesystem::path(selectedPath.get());
    }

    if (destination.empty())
    {
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);
    if (_fileOperationRequestCallback)
    {
        FileOperationRequest request{};
        request.operation         = FILESYSTEM_MOVE;
        request.sourcePaths       = std::move(paths);
        request.destinationFolder = destination;
        request.flags             = flags;

        const HRESULT hrStart = _fileOperationRequestCallback(std::move(request));
        if (FAILED(hrStart))
        {
            ReportError(L"Move", hrStart);
        }
        return;
    }

    if (! IsDirectFileOperationFallbackEnabledForSelfTest())
    {
        Debug::Error(L"FolderView::MoveSelectedItems missing FileOperationRequestCallback; refusing direct file-operation fallback");
        ReportError(L"Move", kMissingFileOperationHostHr);
        return;
    }

    if (! ConfirmNonRevertableFileOperation(_hWnd.get(), _fileSystem.get(), FILESYSTEM_MOVE, paths, destination))
    {
        return;
    }

    FileSystemArenaOwner arenaOwner;
    const wchar_t** sourcePaths = nullptr;
    unsigned long count         = 0;
    HRESULT hr                  = BuildPathArrayArena(paths, arenaOwner, &sourcePaths, &count);
    if (FAILED(hr))
    {
        ReportError(L"Move", hr);
        return;
    }

    hr = _fileSystem->MoveItems(sourcePaths, count, destination.c_str(), flags, nullptr, nullptr, nullptr);
    if (FAILED(hr))
    {
        ReportError(L"Move", hr);
        return;
    }

    DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
    for (const auto& path : paths)
    {
        cache.NotifyPathMoved(_fileSystem.get(), path, destination / path.filename());
    }
    if (! _currentFolder || ! cache.IsFolderWatched(_fileSystem.get(), _currentFolder.value()))
    {
        ForceRefresh();
    }
}
