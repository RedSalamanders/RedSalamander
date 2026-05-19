#include "FolderViewInternal.h"

#include "FluentIcons.h"
#include "NavigationLocation.h"
#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"
#endif

#ifdef ENABLE_TESTS
HWND GetFolderViewRenamePromptHandle() noexcept
{
    const HWND hwnd = FindWindowW(kFolderViewRenamePromptClassName, nullptr);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

bool DebugGetFolderViewRenamePromptSnapshot(FolderViewRenamePromptDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetFolderViewRenamePromptHandle();
    return hwnd && SendMessageW(hwnd,
                                WndMsg::kFolderViewRenamePromptDebug,
                                static_cast<WPARAM>(FolderViewRenamePromptDebugCommand::GetSnapshot),
                                reinterpret_cast<LPARAM>(&out)) != FALSE;
}

bool DebugSetFolderViewRenamePromptText(std::wstring_view text) noexcept
{
    const HWND hwnd = GetFolderViewRenamePromptHandle();
    const std::wstring payload(text);
    return hwnd && SendMessageW(hwnd,
                                WndMsg::kFolderViewRenamePromptDebug,
                                static_cast<WPARAM>(FolderViewRenamePromptDebugCommand::SetText),
                                reinterpret_cast<LPARAM>(&payload)) != FALSE;
}

bool DebugConfirmFolderViewRenamePrompt() noexcept
{
    const HWND hwnd = GetFolderViewRenamePromptHandle();
    return hwnd && SendMessageW(hwnd, WndMsg::kFolderViewRenamePromptDebug, static_cast<WPARAM>(FolderViewRenamePromptDebugCommand::Confirm), 0) != FALSE;
}

bool DebugCancelFolderViewRenamePrompt() noexcept
{
    const HWND hwnd = GetFolderViewRenamePromptHandle();
    return hwnd && SendMessageW(hwnd, WndMsg::kFolderViewRenamePromptDebug, static_cast<WPARAM>(FolderViewRenamePromptDebugCommand::Cancel), 0) != FALSE;
}
#endif

void FolderView::SetPaneFocused(bool focused) noexcept
{
    if (_paneFocused == focused)
    {
        return;
    }

    _paneFocused = focused;
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

FolderView::FolderView() : _appTheme(ResolveAppTheme(ThemeMode::System, L"")), _theme(_appTheme.folderView)
{
    _items.reserve(256);
    _alertOverlay = std::make_unique<RedSalamander::Ui::AlertOverlay>();
}

FolderView::~FolderView()
{
    Destroy();
}

void FolderView::RecordPendingInputToPaintStart(std::chrono::steady_clock::time_point inputStartedAt) noexcept
{
    _pendingInputToPaintStart = inputStartedAt;
}

void FolderView::RecordInputToPaintStartIfViewportOrFocusChanged(std::chrono::steady_clock::time_point inputStartedAt,
                                                                  float scrollBefore,
                                                                  float horizontalBefore,
                                                                  size_t focusedBefore) noexcept
{
    if (_scrollOffset != scrollBefore || _horizontalOffset != horizontalBefore || _focusedIndex != focusedBefore)
    {
        RecordPendingInputToPaintStart(inputStartedAt);
    }
}

void FolderView::EmitPendingInputToPaintMetricAfterPresent() noexcept
{
    if (! _pendingInputToPaintStart.has_value())
    {
        return;
    }

    const auto inputStartedAt = _pendingInputToPaintStart.value();
    _pendingInputToPaintStart.reset();
    Debug::Perf::Emit(L"folder.frame.input_to_paint_us", L"", Debug::Perf::ElapsedUs(inputStartedAt), 0u, 0u, S_OK);
}

std::filesystem::path FolderView::GetItemFullPath(const FolderItem& item) const
{
    if (_itemsFolder.empty())
    {
        return std::filesystem::path(item.displayName);
    }

    std::filesystem::path fullPath = _itemsFolder;
    fullPath /= item.displayName;
    return fullPath;
}

ATOM FolderView::RegisterWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = FolderView::WndProcThunk;
    wc.cbClsExtra    = 0;
    wc.cbWndExtra    = 0;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kFolderViewClassName;
    atom             = RegisterClassExW(&wc);
    return atom;
}

HWND FolderView::Create(HWND parent, int x, int y, int width, int height)
{
    if (_hWnd)
    {
        return _hWnd.get();
    }

    _hParent.reset(parent);
    auto hinst = reinterpret_cast<HINSTANCE>(GetWindowLongPtrW(parent, GWLP_HINSTANCE));
    RegisterWndClass(hinst);

    CreateWindowExW(0,
                    kFolderViewClassName,
                    L"",
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_HSCROLL,
                    x,
                    y,
                    width,
                    height,
                    parent,
                    nullptr,
                    hinst,
                    this);

    // _hwnd is set in WndProcThunk during WM_NCCREATE
    return _hWnd.get();
}

void FolderView::Destroy()
{
    CancelPendingEnumeration();
    StopEnumerationThread();

    _directoryCachePin = {};
    _items.clear();
    _itemsArenaBuffer.reset();
    _itemsFolder.clear();
    _currentFolder.reset();
    _displayedFolder.reset();
    _focusMemory.clear();
    _focusMemoryRootKey.clear();

    _fileSystem.reset(); // release before plugin DLL can unload

    DiscardDeviceResources();

    _hWnd.reset();

    if (_coInitialized)
    {
        CoUninitialize();
        _coInitialized = false;
    }
    if (_oleInitialized)
    {
        OleUninitialize();
        _oleInitialized = false;
    }
}

void FolderView::SetFolderPath(const std::optional<std::filesystem::path>& folderPath)
{
    ExitIncrementalSearch();

    if (! folderPath)
    {
        _hiddenNames.store(std::shared_ptr<const HiddenNamesFilter>{}, std::memory_order_release);
        _pendingExternalCommandAfterEnumeration.reset();
        ClearErrorOverlay(ErrorOverlayKind::Enumeration);
        _directoryCachePin = DirectoryInfoCache::Pin{};
        _currentFolder.reset();
        _displayedFolder.reset();
        _items.clear();
        _itemsArenaBuffer.reset();
        _itemsFolder.clear();
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
        return;
    }

    if (! _currentFolder.has_value() || ! OrdinalString::EqualsNoCasePath(_currentFolder.value(), folderPath.value()))
    {
        const auto hiddenNames = _hiddenNames.load(std::memory_order_acquire);
        if (hiddenNames && ! hiddenNames->names.empty())
        {
            _hiddenNames.store(std::shared_ptr<const HiddenNamesFilter>{}, std::memory_order_release);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), nullptr, FALSE);
            }
        }
    }

    _currentFolder = folderPath;
    if (_fileSystem && _hWnd)
    {
        _directoryCachePin =
            DirectoryInfoCache::GetInstance().PinFolder(_fileSystem.get(), _currentFolder.value(), _hWnd.get(), WndMsg::kFolderViewDirectoryImpact);
    }
    else
    {
        _directoryCachePin = DirectoryInfoCache::Pin{};
    }

    // Notify parent window of path change
    if (_pathChangedCallback)
    {
        _pathChangedCallback(_currentFolder);
    }

    EnumerateFolder();
}

void FolderView::ForceRefresh()
{
#ifdef ENABLE_TESTS
    ++_debugForceRefreshCount;
#endif

    if (_fileSystem && _currentFolder && _hWnd)
    {
        DirectoryInfoCache::GetInstance().InvalidateFolder(_fileSystem.get(), _currentFolder.value());
        _lastDirectoryCacheRefreshTick = GetTickCount64();
        RequestRefreshFromCache();
        return;
    }

    EnumerateFolder();
}

void FolderView::SetEmptyStateMessage(std::wstring message)
{
    if (message == _emptyStateMessage)
    {
        return;
    }

    _emptyStateMessage     = std::move(message);
    _emptyStateMessageKind = EmptyStateMessageKind::None;
    if (! _emptyStateMessage.empty())
    {
        const std::wstring compareNoDiff = LoadStringResource(nullptr, IDS_COMPARE_NO_DIFFERENCES);
        _emptyStateMessageKind           = (_emptyStateMessage == compareNoDiff) ? EmptyStateMessageKind::CompareNoDifferences : EmptyStateMessageKind::Generic;
    }

    _emptyMessageIconLayout.reset();
    _emptyMessageTitleLayout.reset();
    _emptyMessageFunLayout.reset();
    _emptyMessageLayoutClientSizePx = {};
    _emptyMessageLayoutDpi          = 0.0f;
    _emptyMessageLayoutMessageId    = 0;
    _emptyMessageIconFontSizeDip    = 0.0f;
    _emptyMessageIconMetrics        = {};
    _emptyMessageTitleMetrics       = {};
    _emptyMessageFunMetrics         = {};

    UpdateCompareNoDifferencesState();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::SetBackgroundWatermark(std::wstring message, bool animated)
{
    if (message == _backgroundWatermarkMessage && animated == _backgroundWatermarkAnimated)
    {
        return;
    }

    _backgroundWatermarkMessage  = std::move(message);
    _backgroundWatermarkAnimated = animated;

    _backgroundWatermarkLayout.reset();
    _backgroundWatermarkLayoutClientSizePx = {};
    _backgroundWatermarkLayoutDpi          = 0.0f;
    _backgroundWatermarkLayoutText.clear();
    _backgroundWatermarkLayoutFontSizeDip = 0.0f;

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool FolderView::CanShowEmptyFolderState() const noexcept
{
    if (! _emptyStateMessage.empty())
    {
        return false;
    }

    if (! _backgroundWatermarkMessage.empty())
    {
        return false;
    }

    if (IsNameFilterActive())
    {
        return false;
    }

    if (! _currentFolder || ! _displayedFolder)
    {
        return false;
    }

    if (! _items.empty())
    {
        return false;
    }

    if (_pendingBusyOverlay.has_value())
    {
        return false;
    }

    if (! OrdinalString::EqualsNoCasePath(_currentFolder.value(), _displayedFolder.value()))
    {
        return false;
    }

    bool hasOverlay = false;
    {
        std::lock_guard lock(_errorOverlayMutex);
        hasOverlay = _errorOverlay.has_value();
    }

    return ! hasOverlay;
}

void FolderView::RefreshDetailsText()
{
    const bool includeDetailsLine  = _displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed ||
                                    _displayMode == DisplayMode::Thumbnails;
    const bool includeMetadataLine = _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    if (! includeDetailsLine)
    {
        return;
    }

    if (! _detailsTextProvider && ! (includeMetadataLine && _metadataTextProvider))
    {
        return;
    }

    if (_items.empty())
    {
        return;
    }

    bool anyChanged = false;
    for (auto& item : _items)
    {
        if (item.displayName.empty())
        {
            continue;
        }

        if (_detailsTextProvider)
        {
            std::wstring details =
                _detailsTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
            if (details != item.detailsText)
            {
                anyChanged       = true;
                item.detailsText = std::move(details);
                item.detailsLayout.reset();
                item.detailsMetrics = {};
            }
        }

        if (includeMetadataLine && _metadataTextProvider)
        {
            std::wstring metadata =
                _metadataTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
            if (metadata != item.metadataText)
            {
                anyChanged        = true;
                item.metadataText = std::move(metadata);
                item.metadataLayout.reset();
                item.metadataMetrics = {};
            }
        }
    }

    if (! anyChanged)
    {
        return;
    }

    _itemMetricsCached = false;
    LayoutItems();
    UpdateScrollMetrics();
    ScheduleIdleLayoutCreation();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::OnDpiChanged(float newDpi)
{
    if (newDpi <= 0.0f)
        return;
    _dpi                   = newDpi;
    _itemMetricsCached     = false;
    _estimatedMetricsValid = false; // Recompute estimated metrics from font at new DPI
    if (_d2dContext)
    {
        _d2dContext->SetDpi(_dpi, _dpi);
    }
    // Update icon cache DPI (note: existing cached icons won't be updated)
    IconCache::GetInstance().SetDpi(_dpi);
    LayoutItems();
    UpdateScrollMetrics();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::SetFileSystem(const wil::com_ptr<IFileSystem>& fileSystem)
{
    _directoryCachePin = DirectoryInfoCache::Pin{};
    _fileSystem        = fileSystem;
    _displayedFolder.reset();
    _focusMemory.clear();
    _focusMemoryRootKey.clear();
    _fileSystemMetadata = nullptr;
    if (_fileSystem)
    {
        wil::com_ptr<IInformations> infos;
        auto hr = _fileSystem->QueryInterface(__uuidof(IInformations), infos.put_void());
        if (SUCCEEDED(hr) && infos)
        {
            _fileSystemMetadata  = nullptr;
            const HRESULT metaHr = infos->GetMetaData(&_fileSystemMetadata);
            if (FAILED(metaHr))
            {
                _fileSystemMetadata = nullptr;
                Debug::Error(L"FolderView::SetFileSystem: Failed to get file system metadata, hr=0x%08X", metaHr);
            }
        }
    }

    if (_currentFolder && _fileSystem && _hWnd)
    {
        const std::wstring currentFolderText = _currentFolder->wstring();
        const bool currentLooksWindows       = NavigationLocation::LooksLikeWindowsAbsolutePath(currentFolderText);
        const bool currentLooksPluginPath    = ! currentFolderText.empty() && (currentFolderText.front() == L'/' || currentFolderText.front() == L'\\');
        const std::wstring_view pluginShortId =
            (_fileSystemMetadata && _fileSystemMetadata->shortId) ? std::wstring_view(_fileSystemMetadata->shortId) : std::wstring_view{};
        const bool isFilePlugin = NavigationLocation::IsFilePluginShortId(pluginShortId);

        if ((isFilePlugin && currentLooksPluginPath && ! currentLooksWindows) || (! isFilePlugin && currentLooksWindows))
        {
            Debug::Info(L"FolderView::SetFileSystem: skipping repin for currentFolder='{}' pluginShortId='{}' isFilePlugin={}",
                        currentFolderText,
                        pluginShortId.empty() ? std::wstring_view(L"<unknown>") : pluginShortId,
                        isFilePlugin ? L"true" : L"false");
        }
        else
        {
            Debug::Info(L"FolderView::SetFileSystem: pinning currentFolder='{}' pluginShortId='{}' isFilePlugin={}",
                        currentFolderText,
                        pluginShortId.empty() ? std::wstring_view(L"<unknown>") : pluginShortId,
                        isFilePlugin ? L"true" : L"false");
            _directoryCachePin =
                DirectoryInfoCache::GetInstance().PinFolder(_fileSystem.get(), _currentFolder.value(), _hWnd.get(), WndMsg::kFolderViewDirectoryImpact);
        }
    }
    else
    {
        _directoryCachePin = DirectoryInfoCache::Pin{};
    }
}

const PluginMetaData* FolderView::GetFileSystemMetadata() const
{
    return _fileSystemMetadata;
}

LRESULT CALLBACK FolderView::WndProcThunk(HWND hWindow, UINT message, WPARAM wParam, LPARAM lParam)
{
    FolderView* self = nullptr;
    if (message == WM_NCCREATE)
    {
        auto create = reinterpret_cast<CREATESTRUCTW*>(lParam);
        self        = reinterpret_cast<FolderView*>(create->lpCreateParams);
        SetWindowLongPtrW(hWindow, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        self->_hWnd.reset(hWindow);
        InitPostedPayloadWindow(hWindow);
    }
    else
    {
        self = reinterpret_cast<FolderView*>(GetWindowLongPtrW(hWindow, GWLP_USERDATA));
    }

    if (! self)
    {
        return DefWindowProcW(hWindow, message, wParam, lParam);
    }

    return self->WndProc(hWindow, message, wParam, lParam);
}

LRESULT FolderView::WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
        case WM_CREATE: OnCreate(); return 0;
        case WndMsg::kFolderViewDeferredInit: OnDeferredInit(); return 0;
        case WndMsg::kFolderViewEnumerateComplete:
        {
            auto payload = TakeMessagePayload<EnumerationPayload>(lParam);
            ProcessEnumerationResult(std::move(payload));
            return 0;
        }
        case WndMsg::kFolderViewIconLoaded: OnIconLoaded(static_cast<size_t>(lParam)); return 0;
        case WndMsg::kFolderViewBatchIconUpdate: OnBatchIconUpdate(); return 0;
        case WndMsg::kFolderViewCreateIconBitmap:
        {
            auto request = TakeMessagePayload<IconBitmapRequest>(lParam);
            OnCreateIconBitmap(std::move(request));
            return 0;
        }
        case WndMsg::kFolderViewCreateThumbnailBitmap:
        {
            auto request = TakeMessagePayload<ThumbnailBitmapRequest>(lParam);
            OnCreateThumbnailBitmap(std::move(request));
            return 0;
        }
        case WndMsg::kFolderViewDirectoryImpact:
        {
            auto impact = TakeMessagePayload<DirectoryInfoCache::DirectoryImpact>(lParam);
            OnDirectoryImpact(std::move(impact));
            return 0;
        }
        case WndMsg::kFolderViewDirectoryCacheDirty: OnDirectoryCacheDirty(); return 0;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_NCDESTROY: static_cast<void>(DrainPostedPayloadsForWindow(hwnd)); break;
        case WM_DPICHANGED_AFTERPARENT: OnDpiChanged(static_cast<float>(GetDpiForWindow(hwnd))); return 0;
        case WM_SIZE: OnSize(LOWORD(lParam), HIWORD(lParam)); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: OnPaint(); return 0;
        case WM_MOUSEWHEEL: OnMeasuredMouseWheelMessage(LOWORD(wParam), GET_WHEEL_DELTA_WPARAM(wParam)); return 0;
        case WM_MOUSEHWHEEL: OnMeasuredMouseHWheelMessage(GET_WHEEL_DELTA_WPARAM(wParam)); return 0;
        case WM_LBUTTONDOWN: OnLButtonDown({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}, wParam); return 0;
        case WM_LBUTTONDBLCLK: OnLButtonDblClk({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}, wParam); return 0;
        case WM_LBUTTONUP: OnLButtonUp({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}); return 0;
        case WM_MOUSEMOVE: OnMouseMove({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}, wParam); return 0;
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_TIMER: OnTimerMessage(static_cast<UINT_PTR>(wParam)); return 0;
        case WM_KEYDOWN: OnMeasuredKeyDownMessage(wParam, lParam); return 0;
        case WM_CHAR: OnCharMessage(static_cast<wchar_t>(wParam)); return 0;
        case WM_SETFOCUS: return OnSetFocusMessage();
        case WM_KILLFOCUS: return OnKillFocusMessage();
        case WM_SYSKEYDOWN:
            if (OnMeasuredSysKeyDownMessage(wParam, lParam))
            {
                return 0;
            }
            break;
        case WM_SYSCHAR:
            if (wParam == 'D' || wParam == 'd')
            {
                return 0;
            }
            break;
        case WM_GETDLGCODE: return DLGC_WANTTAB | DLGC_WANTARROWS | DLGC_WANTCHARS;
        case WM_CONTEXTMENU: OnContextMenuMessage(hwnd, lParam); return 0;
        case WM_HSCROLL: OnMeasuredHScrollMessage(LOWORD(wParam)); return 0;
        case WM_COMMAND: OnCommandMessage(LOWORD(wParam)); return 0;
    }
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

void FolderView::OnCreate()
{
    const UINT windowDpi = GetDpiForWindow(_hWnd.get());
    if (windowDpi > 0)
    {
        _dpi = static_cast<float>(windowDpi);
    }
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (SUCCEEDED(hr))
    {
        _coInitialized = true;
    }
    else if (hr == RPC_E_CHANGED_MODE)
    {
        // Already initialized in different mode; continue without owning COM lifetime.
        _coInitialized = false;
    }
    else
    {
        ReportError(L"CoInitializeEx", hr);
    }

    HRESULT hrOle = OleInitialize(nullptr);
    if (SUCCEEDED(hrOle))
    {
        _oleInitialized = true;
    }
    else if (hrOle != RPC_E_CHANGED_MODE)
    {
        ReportError(L"OleInitialize", hrOle);
    }

    EnsureDropTarget();
}

void FolderView::OnDeferredInit()
{
#ifdef ENABLE_TESTS
    ++_debugDeferredInitCallCount;
#endif
    Debug::Perf::Scope perf(L"FolderView.DeferredInit");

    const auto computeMissingMask = [&]() noexcept -> uint32_t
    {
        // Bitmask for diagnosing why FolderView is still in fallback rendering.
        // 0x01: client size is zero
        // 0x02: missing D2D device context
        // 0x04: missing swap chain
        // 0x08: missing D2D target bitmap
        // 0x10: swap chain resize pending
        uint32_t mask = 0;
        if (_clientSize.cx <= 0 || _clientSize.cy <= 0)
        {
            mask |= 0x01u;
        }
        if (! _d2dContext)
        {
            mask |= 0x02u;
        }
        if (! _swapChain && ! _swapChainLegacy)
        {
            mask |= 0x04u;
        }
        if (! _d2dTarget)
        {
            mask |= 0x08u;
        }
        if (_swapChainResizePending)
        {
            mask |= 0x10u;
        }
        return mask;
    };

    const uint32_t missingBefore = computeMissingMask();
    perf.SetValue0(missingBefore);

    if (_currentFolder.has_value())
    {
        perf.SetDetail(_currentFolder->native());
    }
    else if (! _itemsFolder.empty())
    {
        perf.SetDetail(_itemsFolder.native());
    }

    EnsureDeviceIndependentResources();
    EnsureDeviceResources();
    EnsureSwapChain();

    const uint32_t missingAfter = computeMissingMask();
    perf.SetValue1(missingAfter);
    perf.SetHr(missingAfter == 0 ? S_OK : S_FALSE);

    // Mark message as consumed only after attempting initialization so we don't re-post while running.
    _deferredInitPosted = false;

    // Icon extraction/conversion only needs the D2D device context. Do not block it behind
    // swap-chain resize completion, or fast startup can leave items stuck on placeholders.
    if (_d2dContext)
    {
        IconCache::GetInstance().Initialize(_d2dContext.get(), _dpi);
        QueueIconLoading();
        if (_thumbnailsVisible)
        {
            QueueThumbnailLoading();
        }
    }

    if (missingAfter != 0)
    {
        // Still not ready (often due to 0x0 size or during active resize). Avoid invalidation loops.
        return;
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::OnDestroy()
{
    // Stop idle layout timer
    if (_idleLayoutTimer != 0 && _hWnd)
    {
        KillTimer(_hWnd.get(), kIdleLayoutTimerId);
        _idleLayoutTimer = 0;
    }
    if (_directoryCacheRefreshTimer != 0 && _hWnd)
    {
        KillTimer(_hWnd.get(), kDirectoryCacheRefreshTimerId);
        _directoryCacheRefreshTimer = 0;
    }

    StopOverlayAnimation();
    StopOverlayTimer();
    CancelPendingEnumeration();
    StopEnumerationThread();
    _directoryCachePin = DirectoryInfoCache::Pin{};
    if (_dropTargetRegistered && _hWnd)
    {
        RevokeDragDrop(_hWnd.get());
        _dropTargetRegistered = false;
    }
    _dropTarget.reset();
    ReleaseSwapChain();
    DiscardDeviceResources();

    if (_oleInitialized)
    {
        OleUninitialize();
        _oleInitialized = false;
    }
    if (_coInitialized)
    {
        CoUninitialize();
        _coInitialized = false;
    }
}

void FolderView::OnSize(UINT width, UINT height)
{
    _clientSize.cx = static_cast<int>(width);
    _clientSize.cy = static_cast<int>(height);

    // Debug::Info(L"FolderView::OnSize {}x{}", width, height);
    _swapChainResizePending = true;
    _pendingSwapChainWidth  = static_cast<UINT>(std::max(1L, _clientSize.cx));
    _pendingSwapChainHeight = static_cast<UINT>(std::max(1L, _clientSize.cy));

    LayoutItems();
    UpdateScrollMetrics();
    QueueMissingVisibleThumbnails();
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void FolderView::OnPaint()
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint paint_dc = wil::BeginPaint(_hWnd.get(), &ps);

    // Handle pending swap chain resize BEFORE rendering to ensure valid render target
    if (_swapChainResizePending && _clientSize.cx > 0 && _clientSize.cy > 0)
    {
        auto strInfo = std::format(L"{}x{}", _pendingSwapChainWidth, _pendingSwapChainHeight);
        TRACER_CTX(strInfo.c_str());
        Debug::Info(L"FolderView::OnPaint handling deferred swap-chain resize");

        if (_swapChain || _swapChainLegacy)
        {
            if (TryResizeSwapChain(_pendingSwapChainWidth, _pendingSwapChainHeight))
            {
                _swapChainResizePending = false;

                // Recreate the D2D target for the resized swap chain so we can render immediately
                // instead of falling back and posting another deferred init.
                EnsureSwapChain();
            }
        }
        else
        {
            _swapChainResizePending = false;
        }
    }

    RECT rcPaint = ps.rcPaint;

    if (! _d2dContext || (! _swapChain && ! _swapChainLegacy) || ! _d2dTarget)
    {
        HBRUSH fillBrush = _menuBackgroundBrush ? _menuBackgroundBrush.get() : nullptr;
        wil::unique_hbrush fallbackBrush;
        if (! fillBrush)
        {
            auto toByte = [](float value) -> BYTE
            {
                if (value <= 0.0f)
                {
                    return 0;
                }
                if (value >= 1.0f)
                {
                    return 255;
                }
                return static_cast<BYTE>(value * 255.0f + 0.5f);
            };

            const COLORREF rgb = RGB(toByte(_theme.backgroundColor.r), toByte(_theme.backgroundColor.g), toByte(_theme.backgroundColor.b));
            fallbackBrush.reset(CreateSolidBrush(rgb));
            fillBrush = fallbackBrush ? fallbackBrush.get() : static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH));
        }
        FillRect(paint_dc.get(), &rcPaint, fillBrush);
        if (! _deferredInitPosted && _hWnd && _clientSize.cx > 0 && _clientSize.cy > 0)
        {
            _deferredInitPosted = PostMessageW(_hWnd.get(), WndMsg::kFolderViewDeferredInit, 0, 0) != 0;
        }
        return;
    }

    Render(rcPaint);
}

void FolderView::SetAppTheme(const AppTheme& theme)
{
    const bool compactModeChanged = _appTheme.compactMode != theme.compactMode;
    _appTheme                     = theme;
    if (compactModeChanged)
    {
        LayoutItems();
        UpdateScrollMetrics();
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

#ifdef ENABLE_TESTS
bool FolderView::DebugWarmRenderingForSelfTest() noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    ++_debugWarmRenderingCallCount;

    if (! _hWnd || IsWindow(_hWnd.get()) == FALSE)
    {
        return false;
    }

    RECT rc{};
    GetClientRect(_hWnd.get(), &rc);
    _clientSize.cx = std::max<LONG>(0L, rc.right - rc.left);
    _clientSize.cy = std::max<LONG>(0L, rc.bottom - rc.top);
    if (_clientSize.cx <= 0 || _clientSize.cy <= 0)
    {
        return false;
    }

    _swapChainResizePending = true;
    _pendingSwapChainWidth  = static_cast<UINT>(std::max<LONG>(1L, _clientSize.cx));
    _pendingSwapChainHeight = static_cast<UINT>(std::max<LONG>(1L, _clientSize.cy));

    OnDeferredInit();
    const bool ready = _d2dContext && (_swapChain || _swapChainLegacy) && _d2dTarget;
    if (! ready)
    {
        Debug::Perf::Emit(L"folder.selftest.render_warmup_us", L"not-ready", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, S_FALSE);
        return false;
    }

    RECT fullRect{};
    fullRect.right  = _clientSize.cx;
    fullRect.bottom = _clientSize.cy;

    Render(fullRect);

    QueueIconLoading();
    if (_thumbnailsVisible)
    {
        QueueThumbnailLoading();
    }
    if (_iconLoadingActive.load(std::memory_order_acquire))
    {
        ProcessIconLoadQueue();
    }
    if (_thumbnailLoadingActive.load(std::memory_order_acquire))
    {
        ProcessThumbnailLoadQueue();
    }

    uint64_t drainedMessages = 0;
    MSG msg{};
    while (PeekMessageW(&msg, _hWnd.get(), 0, 0, PM_REMOVE) != FALSE)
    {
        ++drainedMessages;
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    OnBatchIconUpdate();
    Render(fullRect);

    Debug::Perf::Emit(L"folder.selftest.render_warmup_us", L"", Debug::Perf::ElapsedUs(startedAt), drainedMessages, _items.size(), S_OK);
    return true;
}
#endif

void FolderView::SetTheme(const FolderViewTheme& theme)
{
    _theme = theme;
    RecreateThemeBrushes();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::SetMenuTheme(const MenuTheme& menuTheme)
{
    _menuTheme = menuTheme;
    _menuBackgroundBrush.reset(CreateSolidBrush(_menuTheme.background));
}

void FolderView::SetShortcutManager(const ShortcutManager* shortcuts) noexcept
{
    _shortcutManager = shortcuts;
}

void FolderView::SetDisplayMode(DisplayMode mode)
{
    const bool thumbnailsVisible = mode == DisplayMode::Thumbnails;
    if (_displayMode == mode && _thumbnailsVisible == thumbnailsVisible)
    {
        return;
    }

    const bool thumbnailsChanged = _thumbnailsVisible != thumbnailsVisible;
    _displayMode            = mode;
    _thumbnailsVisible      = thumbnailsVisible;
    _itemMetricsCached      = false;
    _cachedMaxLabelWidth    = 0.0f;
    _cachedMaxLabelHeight   = 0.0f;
    _cachedMaxDetailsWidth  = 0.0f;
    _cachedMaxMetadataWidth = 0.0f;
    _detailsSizeSlotChars   = 0;
    _lastLayoutWidth        = 0.0f;

    if (thumbnailsChanged)
    {
        CancelThumbnailLoading();
        _iconSizeDip = thumbnailsVisible ? static_cast<float>(_thumbnailSizeDip) : kFolderViewListIconSizeDip;
        for (auto& item : _items)
        {
            item.icon.reset();
            item.thumbnail.reset();
            item.labelLayout.reset();
            item.labelMetrics = {};
            item.detailsLayout.reset();
            item.detailsMetrics = {};
            item.metadataLayout.reset();
            item.metadataMetrics = {};
        }
    }
    else if (_displayMode == DisplayMode::Brief)
    {
        for (auto& item : _items)
        {
            item.detailsLayout.reset();
            item.detailsMetrics = {};
            item.metadataLayout.reset();
            item.metadataMetrics = {};
        }
    }
    else if (_displayMode == DisplayMode::Detailed)
    {
        for (auto& item : _items)
        {
            item.metadataLayout.reset();
            item.metadataMetrics = {};
        }
    }

    LayoutItems();
    UpdateScrollMetrics();
    QueueIconLoading();
    if (_thumbnailsVisible)
    {
        QueueThumbnailLoading();
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::SetThumbnailSizeDip(uint32_t sizeDip)
{
    const uint32_t normalizedSizeDip = Common::Settings::Thumbnail::NormalizeSizeDip(sizeDip);
    if (_thumbnailSizeDip == normalizedSizeDip)
    {
        return;
    }

    Debug::Perf::Scope perf(L"thumbnails.size_change_us");
    perf.SetValue0(_thumbnailSizeDip);
    perf.SetValue1(normalizedSizeDip);

    _thumbnailSizeDip = normalizedSizeDip;
    if (! _thumbnailsVisible)
    {
        return;
    }

    CancelThumbnailLoading();
    _iconSizeDip = static_cast<float>(_thumbnailSizeDip);
    _itemMetricsCached      = false;
    _cachedMaxLabelWidth    = 0.0f;
    _cachedMaxLabelHeight   = 0.0f;
    _cachedMaxDetailsWidth  = 0.0f;
    _cachedMaxMetadataWidth = 0.0f;
    _lastLayoutWidth        = 0.0f;

    for (auto& item : _items)
    {
        item.thumbnail.reset();
        item.icon.reset();
        item.labelLayout.reset();
        item.labelMetrics = {};
        item.detailsLayout.reset();
        item.detailsMetrics = {};
        item.metadataLayout.reset();
        item.metadataMetrics = {};
    }

    LayoutItems();
    UpdateScrollMetrics();
    QueueIconLoading();
    QueueThumbnailLoading();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::SetFileExtensionsVisible(bool visible)
{
    if (_fileExtensionsVisible == visible)
    {
        return;
    }

    _fileExtensionsVisible = visible;
    _itemMetricsCached      = false;
    _cachedMaxLabelWidth    = 0.0f;
    _cachedMaxLabelHeight   = 0.0f;
    _lastLayoutWidth        = 0.0f;

    for (auto& item : _items)
    {
        item.labelLayout.reset();
        item.labelMetrics = {};
    }

    LayoutItems();
    UpdateScrollMetrics();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::SetThumbnailsVisible(bool visible)
{
    if (visible)
    {
        SetDisplayMode(DisplayMode::Thumbnails);
        return;
    }

    if (_displayMode == DisplayMode::Thumbnails || _thumbnailsVisible)
    {
        SetDisplayMode(DisplayMode::Brief);
    }
}

void FolderView::SetSort(SortBy sortBy, SortDirection direction)
{
    if (_sortBy == sortBy && _sortDirection == direction)
    {
        return;
    }

    _sortBy        = sortBy;
    _sortDirection = direction;
    ApplyCurrentSort();

    LayoutItems();
    UpdateScrollMetrics();
    QueueIconLoading();
    if (_thumbnailsVisible)
    {
        QueueThumbnailLoading();
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::SetShowHiddenFiles(bool show)
{
    if (_showHiddenFiles.load(std::memory_order_relaxed) == show)
    {
        return;
    }

    _showHiddenFiles.store(show, std::memory_order_release);
    RequestRefreshFromCache();
}

bool FolderView::GetShowHiddenFiles() const noexcept
{
    return _showHiddenFiles.load(std::memory_order_relaxed);
}

void FolderView::SetShowSystemFiles(bool show)
{
    if (_showSystemFiles.load(std::memory_order_relaxed) == show)
    {
        return;
    }

    _showSystemFiles.store(show, std::memory_order_release);
    RequestRefreshFromCache();
}

bool FolderView::GetShowSystemFiles() const noexcept
{
    return _showSystemFiles.load(std::memory_order_relaxed);
}

void FolderView::SetNameFilterState(const NameFilterState& state, bool refresh)
{
    const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(state.text);
#ifdef ENABLE_TESTS
    SelfTest::AppendSelfTestTrace(
        std::format(L"FolderView::SetNameFilterState: enabled={} trimmed='{}' refresh={}", state.enabled ? 1 : 0, trimmed, refresh ? 1 : 0));
#endif

    if (! state.enabled && trimmed.empty())
    {
        _nameFilter.store(std::shared_ptr<const CompiledNameFilter>{}, std::memory_order_release);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        if (refresh)
        {
#ifdef ENABLE_TESTS
            SelfTest::AppendSelfTestTrace(L"FolderView::SetNameFilterState: clearing filter and requesting refresh");
#endif
            RequestRefreshFromCache();
        }
        return;
    }

    auto compiled           = std::make_shared<CompiledNameFilter>();
    compiled->state.enabled = state.enabled;
    compiled->state.text    = trimmed;
    compiled->mask          = MaskSyntax::ParseWildcardMask(trimmed);
    compiled->hasMask       = ! compiled->mask.includePatterns.empty() || ! compiled->mask.excludePatterns.empty();

    std::shared_ptr<const CompiledNameFilter> compiledConst = std::move(compiled);
    _nameFilter.store(std::move(compiledConst), std::memory_order_release);

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    if (refresh)
    {
#ifdef ENABLE_TESTS
        SelfTest::AppendSelfTestTrace(L"FolderView::SetNameFilterState: stored compiled filter and requesting refresh");
#endif
        RequestRefreshFromCache();
    }
}

FolderView::NameFilterState FolderView::GetNameFilterState() const
{
    const auto filter = _nameFilter.load(std::memory_order_acquire);
    return filter ? filter->state : NameFilterState{};
}

bool FolderView::IsNameFilterActive() const noexcept
{
    const auto filter = _nameFilter.load(std::memory_order_acquire);
    if (filter && filter->state.enabled && filter->hasMask)
    {
        return true;
    }

    const auto hiddenNames = _hiddenNames.load(std::memory_order_acquire);
    return hiddenNames && ! hiddenNames->names.empty();
}

bool FolderView::HasHiddenNames() const noexcept
{
    const auto hiddenNames = _hiddenNames.load(std::memory_order_acquire);
    return hiddenNames && ! hiddenNames->names.empty();
}
