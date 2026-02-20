#include "FolderViewInternal.h"

#include "FluentIcons.h"

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

FolderView::FolderView() : _theme(ResolveAppTheme(ThemeMode::System, L"").folderView)
{
    _items.reserve(256);
    _alertOverlay = std::make_unique<RedSalamander::Ui::AlertOverlay>();
}

FolderView::~FolderView()
{
    Destroy();
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

    _currentFolder = folderPath;
    if (_fileSystem && _hWnd)
    {
        _directoryCachePin =
            DirectoryInfoCache::GetInstance().PinFolder(_fileSystem.get(), _currentFolder.value(), _hWnd.get(), WndMsg::kFolderViewDirectoryCacheDirty);
    }
    else
    {
        _directoryCachePin = DirectoryInfoCache::Pin{};
    }
    EnumerateFolder();

    // Notify parent window of path change
    if (_pathChangedCallback)
    {
        _pathChangedCallback(_currentFolder);
    }
}

void FolderView::ForceRefresh()
{
#ifdef _DEBUG
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

    _emptyStateMessage = std::move(message);
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void FolderView::RefreshDetailsText()
{
    if (_displayMode == DisplayMode::Brief)
    {
        return;
    }

    if (! _detailsTextProvider && ! (_displayMode == DisplayMode::ExtraDetailed && _metadataTextProvider))
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

        if (_displayMode == DisplayMode::ExtraDetailed && _metadataTextProvider)
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
    _dpi               = newDpi;
    _menuFont          = CreateMenuFontForDpi(static_cast<UINT>(_dpi));
    _menuIconFont      = FluentIcons::CreateFontForDpi(static_cast<UINT>(_dpi), FluentIcons::kDefaultSizeDip);
    _menuIconFontDpi   = static_cast<UINT>(_dpi);
    _menuIconFontValid = false;
    if (_menuIconFont && _hWnd)
    {
        auto hdc = wil::GetDC(_hWnd.get());
        if (hdc)
        {
            _menuIconFontValid = FluentIcons::FontHasGlyph(hdc.get(), _menuIconFont.get(), FluentIcons::kChevronRightSmall);
        }
    }
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
    _fileSystem = fileSystem;
    _displayedFolder.reset();
    _focusMemory.clear();
    _focusMemoryRootKey.clear();
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
        _directoryCachePin =
            DirectoryInfoCache::GetInstance().PinFolder(_fileSystem.get(), *_currentFolder, _hWnd.get(), WndMsg::kFolderViewDirectoryCacheDirty);
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
        case WndMsg::kFolderViewDirectoryCacheDirty: OnDirectoryCacheDirty(); return 0;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_NCDESTROY: static_cast<void>(DrainPostedPayloadsForWindow(hwnd)); break;
        case WM_SIZE: OnSize(LOWORD(lParam), HIWORD(lParam)); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: OnPaint(); return 0;
        case WM_MOUSEWHEEL: OnMouseWheelMessage(LOWORD(wParam), GET_WHEEL_DELTA_WPARAM(wParam)); return 0;
        case WM_MOUSEHWHEEL: OnMouseWheel(GET_WHEEL_DELTA_WPARAM(wParam), true); return 0;
        case WM_LBUTTONDOWN: OnLButtonDown({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}, wParam); return 0;
        case WM_LBUTTONDBLCLK: OnLButtonDblClk({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}, wParam); return 0;
        case WM_LBUTTONUP: OnLButtonUp({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}); return 0;
        case WM_MOUSEMOVE: OnMouseMove({GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)}, wParam); return 0;
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_TIMER: OnTimerMessage(static_cast<UINT_PTR>(wParam)); return 0;
        case WM_KEYDOWN: OnKeyDownMessage(wParam); return 0;
        case WM_CHAR: OnCharMessage(static_cast<wchar_t>(wParam)); return 0;
        case WM_SETFOCUS: return OnSetFocusMessage();
        case WM_KILLFOCUS: return OnKillFocusMessage();
        case WM_SYSKEYDOWN:
            if (OnSysKeyDownMessage(wParam))
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
        case WM_HSCROLL: OnHScrollMessage(LOWORD(wParam)); return 0;
        case WM_MEASUREITEM: OnMeasureItem(reinterpret_cast<MEASUREITEMSTRUCT*>(lParam)); return TRUE;
        case WM_DRAWITEM: OnDrawItem(reinterpret_cast<DRAWITEMSTRUCT*>(lParam)); return TRUE;
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
    _menuFont          = CreateMenuFontForDpi(static_cast<UINT>(_dpi));
    _menuIconFont      = FluentIcons::CreateFontForDpi(static_cast<UINT>(_dpi), FluentIcons::kDefaultSizeDip);
    _menuIconFontDpi   = static_cast<UINT>(_dpi);
    _menuIconFontValid = false;
    if (_menuIconFont && _hWnd)
    {
        auto hdc = wil::GetDC(_hWnd.get());
        if (hdc)
        {
            _menuIconFontValid = FluentIcons::FontHasGlyph(hdc.get(), _menuIconFont.get(), FluentIcons::kChevronRightSmall);
        }
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

    if (missingAfter != 0)
    {
        // Still not ready (often due to 0x0 size or during active resize). Avoid invalidation loops.
        return;
    }

    // Initialize application-wide icon cache
    if (_d2dContext)
    {
        IconCache::GetInstance().Initialize(_d2dContext.get(), _dpi);
    }

    // Icon loading can be queued before D2D resources exist (during early enumeration).
    // Re-queue now that we can actually convert icons to bitmaps.
    QueueIconLoading();

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
    if (_displayMode == mode)
    {
        return;
    }

    _displayMode            = mode;
    _itemMetricsCached      = false;
    _cachedMaxLabelWidth    = 0.0f;
    _cachedMaxLabelHeight   = 0.0f;
    _cachedMaxDetailsWidth  = 0.0f;
    _cachedMaxMetadataWidth = 0.0f;
    _detailsSizeSlotChars   = 0;
    _lastLayoutWidth        = 0.0f;

    if (_displayMode == DisplayMode::Brief)
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

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
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

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}
