#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cwctype>
#include <format>
#include <iterator>
#include <limits>
#include <new>
#include <string_view>

#include <windows.h>
#include <windowsx.h>

#include <commctrl.h>

#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "dwrite.lib")
#pragma comment(lib, "dxgi.lib")

#include "NavigationView.h"
#include "NavigationViewInternal.h"

#include "NavigationLocation.h"

#include "DirectoryInfoCache.h"
#include "Helpers.h"
#include "IconCache.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "ThemedControls.h"
#include "resource.h"

NavigationView::NavigationView() = default;

NavigationView::~NavigationView()
{
    Destroy();
}

HRESULT STDMETHODCALLTYPE NavigationView::NavigationMenuRequestNavigate(const wchar_t* path, void* cookie) noexcept
{
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (cookie != _fileSystemPlugin.get()) // check the sanity check from SetCallback
    {
        Debug::Error(L"NavigationView::RequestNavigate: Invalid cookie");
        return S_FALSE;
    }

    if (! _hWnd || ! IsWindow(_hWnd.get()))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }

    auto owned = std::make_unique<std::wstring>(path);

    if (! PostMessagePayload(_hWnd.get(), WndMsg::kNavigationMenuRequestPath, 0, std::move(owned)))
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    return S_OK;
}

void NavigationView::RequestPathChange(const std::filesystem::path& path)
{
    if (_pathChangedCallback)
    {
        _pathChangedCallback(path);
        return;
    }

    SetPath(path);
}

std::filesystem::path NavigationView::ToPluginPath(const std::filesystem::path& displayPath) const
{
    NavigationLocation::Location location;
    if (! NavigationLocation::TryParseLocation(displayPath.native(), location))
    {
        return displayPath;
    }

    if (_pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file"))
    {
        return location.pluginPath;
    }

    if (! location.pluginShortId.empty() && ! EqualsNoCase(location.pluginShortId, _pluginShortId))
    {
        return std::filesystem::path{};
    }

    return NavigationLocation::NormalizePluginPath(location.pluginPath.wstring());
}

ATOM NavigationView::RegisterWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
        return atom;

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr; // Custom painting
    wc.lpszClassName = kClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

HWND NavigationView::Create(HWND parent, int x, int y, int width, int height)
{
    Debug::Perf::Scope perf(L"NavigationView.Create");

    _hInstance = GetModuleHandle(nullptr);

    {
        Debug::Perf::Scope perfRegister(L"NavigationView.Create.RegisterWndClass");
        if (! RegisterWndClass(_hInstance))
        {
            return nullptr;
        }
    }

    {
        Debug::Perf::Scope perfCreateWindow(L"NavigationView.Create.CreateWindowExW");
        CreateWindowExW(0, kClassName, L"", WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, x, y, width, height, parent, nullptr, _hInstance, this);
    }

    // _hWnd is set in NCCREATE
    return _hWnd.get();
}

void NavigationView::Destroy()
{
    if (_editSuggestThread.joinable())
    {
        _editSuggestThread.request_stop();
        _editSuggestCv.notify_all();
        _editSuggestThread.join();
    }

    if (_siblingPrefetchThread.joinable())
    {
        _siblingPrefetchThread.request_stop();
        _siblingPrefetchCv.notify_all();
        _siblingPrefetchThread.join();
    }

    {
        std::lock_guard lock(_editSuggestMutex);
        _editSuggestPendingQuery.reset();
    }
    {
        std::lock_guard lock(_siblingPrefetchMutex);
        _siblingPrefetchPendingQuery.reset();
    }

    if (_navigationMenu)
    {
        _navigationMenu->SetCallback(nullptr, nullptr);
    }
    _navigationMenu.reset();
    _driveInfo.reset();
    _fileSystemIo.reset();
    _fileSystemPlugin.reset();

    _navigationMenuActions.clear();
    _driveMenuActions.clear();
    _menuBitmaps.clear();
    _menuIconBitmapD2D = nullptr;

    _hWnd.reset();
}

LRESULT CALLBACK NavigationView::WndProcThunk(HWND hWindow, UINT msg, WPARAM wp, LPARAM lp)
{
    NavigationView* self = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self    = reinterpret_cast<NavigationView*>(cs->lpCreateParams);
        SetWindowLongPtrW(hWindow, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        self->_hWnd.reset(hWindow);
        InitPostedPayloadWindow(hWindow);
    }
    else
    {
        self = reinterpret_cast<NavigationView*>(GetWindowLongPtrW(hWindow, GWLP_USERDATA));
    }

    if (self)
    {
        return self->WndProc(hWindow, msg, wp, lp);
    }

    return DefWindowProcW(hWindow, msg, wp, lp);
}

LRESULT NavigationView::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WndMsg::kNavigationViewDeferredInit: OnDeferredInit(); return 0;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_NCDESTROY: static_cast<void>(DrainPostedPayloadsForWindow(hwnd)); break;
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: OnPaint(); return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_COMMAND: OnCommand(LOWORD(wp), reinterpret_cast<HWND>(lp), HIWORD(wp)); return 0;
        case WM_MEASUREITEM: OnMeasureItem(reinterpret_cast<MEASUREITEMSTRUCT*>(lp)); return TRUE;
        case WM_DRAWITEM: OnDrawItem(reinterpret_cast<DRAWITEMSTRUCT*>(lp)); return TRUE;
        case WM_CTLCOLOREDIT: return OnCtlColorEdit(reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_LBUTTONDOWN: OnLButtonDown({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_LBUTTONDBLCLK: OnLButtonDblClk({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_MOUSEMOVE: OnMouseMove({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_SETCURSOR: OnSetCursor(reinterpret_cast<HWND>(wp), LOWORD(lp), HIWORD(lp)); return TRUE;
        case WM_TIMER: OnTimer(static_cast<UINT_PTR>(wp)); return 0;
        case WM_ENTERMENULOOP: OnEnterMenuLoop(static_cast<BOOL>(wp)); return 0;
        case WM_EXITMENULOOP: OnExitMenuLoop(static_cast<BOOL>(wp)); return 0;
        case WM_SETFOCUS: OnSetFocus(); return 0;
        case WM_KILLFOCUS: OnKillFocus(reinterpret_cast<HWND>(wp)); return 0;
        case WM_KEYDOWN:
            if (OnKeyDown(wp))
            {
                return 0;
            }
            break;
        case WM_SYSKEYDOWN:
            if (OnKeyDown(wp))
            {
                return 0;
            }
            break;
        case WM_SYSCHAR:
            if (wp == 'D' || wp == 'd')
            {
                return 0;
            }
            break;
        case WM_GETDLGCODE: return DLGC_WANTTAB | DLGC_WANTARROWS | DLGC_WANTCHARS;
        case WndMsg::kEditSuggestResults:
        {
            auto payload = TakeMessagePayload<EditSuggestResultsPayload>(lp);
            return OnEditSuggestResults(std::move(payload));
        }
        case WndMsg::kNavigationMenuRequestPath:
        {
            auto text = TakeMessagePayload<std::wstring>(lp);
            return OnNavigationMenuRequestPath(std::move(text));
        }
        case WndMsg::kNavigationMenuShowSiblingsDropdown:
            _pendingSeparatorMenuSwitchIndex = -1;
            ShowSiblingsDropdown(static_cast<size_t>(wp));
            return 0;                                                            // Deferred menu opening
        case WndMsg::kNavigationMenuShowFullPath: ShowFullPathPopup(); return 0; // Deferred full-path popup opening
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT NavigationView::OnNavigationMenuRequestPath(std::unique_ptr<std::wstring> text)
{
    if (text && ! text->empty())
    {
        RequestPathChange(std::filesystem::path(*text));
    }
    return 0;
}

void NavigationView::OnCreate(HWND hWindow)
{
    Debug::Perf::Scope perf(L"NavigationView.OnCreate");

    // _hWnd not yet initialize
    {
        Debug::Perf::Scope perfDpi(L"NavigationView.OnCreate.GetDpiForWindow");
        _dpi = GetDpiForWindow(hWindow);
    }

    // Create GDI resources
    {
        Debug::Perf::Scope perfFont(L"NavigationView.OnCreate.CreateFontW.PathFont");
        int pathFontHeight = -MulDiv(12, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI);
        _pathFont.reset(CreateFontW(pathFontHeight,
                                    0,
                                    0,
                                    0,
                                    FW_NORMAL,
                                    FALSE,
                                    FALSE,
                                    FALSE,
                                    DEFAULT_CHARSET,
                                    OUT_DEFAULT_PRECIS,
                                    CLIP_DEFAULT_PRECIS,
                                    CLEARTYPE_QUALITY,
                                    DEFAULT_PITCH,
                                    L"Segoe UI"));
    }

    {
        Debug::Perf::Scope perfTheme(L"NavigationView.OnCreate.SetTheme");
        const AppTheme resolvedTheme = ResolveAppTheme(ThemeMode::System, L"");
        SetTheme(resolvedTheme);
    }

    // GDI menus are NOT DPI-aware - they always expect physical pixels at 96 DPI
    // Do not scale menu icon size with DPI
    _menuIconSize = GetSystemMetrics(SM_CXSMICON);

    // Create tooltip window
    {
        Debug::Perf::Scope perfIcc(L"NavigationView.OnCreate.InitCommonControls");
        InitCommonControls();
    }

    if (! _navDropdownCombo)
    {
        Debug::Perf::Scope perfCombo(L"NavigationView.OnCreate.NavDropdownCombo.Create");
        _navDropdownCombo.reset(ThemedControls::CreateModernComboBox(hWindow, ID_NAV_DROPDOWN_COMBO, &_appTheme));
        if (_navDropdownCombo)
        {
            Debug::Perf::Scope perfComboInit(L"NavigationView.OnCreate.NavDropdownCombo.Initialize");
            HFONT fontToUse = _pathFont ? _pathFont.get() : (_menuFont ? _menuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
            SendMessageW(_navDropdownCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(fontToUse), FALSE);
            ThemedControls::SetModernComboCloseOnOutsideAccept(_navDropdownCombo.get(), false);
            ThemedControls::SetModernComboDropDownPreferBelow(_navDropdownCombo.get(), true);
            ThemedControls::SetModernComboCompactMode(_navDropdownCombo.get(), true);
            ThemedControls::SetModernComboUseMiddleEllipsis(_navDropdownCombo.get(), true);

            wil::unique_hrgn emptyRgn(CreateRectRgn(0, 0, 0, 0));
            if (emptyRgn)
            {
                SetWindowRgn(_navDropdownCombo.get(), emptyRgn.release(), TRUE);
            }

            SetWindowPos(_navDropdownCombo.get(), nullptr, -32000, -32000, 10, 10, SWP_NOZORDER | SWP_NOACTIVATE | SWP_HIDEWINDOW);
        }
    }
}

void NavigationView::OnDeferredInit()
{
    _deferredInitPosted = false;

    Debug::Perf::Scope perf(L"NavigationView.DeferredInit");
    perf.SetDetail(_hWnd ? L"Visible" : L"");
    perf.SetValue0(_hWnd ? static_cast<uint64_t>(GetWindowLongPtrW(_hWnd.get(), GWLP_ID)) : 0u);

    if (_swapChain && _d2dTarget)
    {
        return;
    }

    EnsureD2DResources();
    if (_d2dContext)
    {
        IconCache::GetInstance().Initialize(_d2dContext.get(), static_cast<float>(_dpi));
    }

    if (_currentPluginPath)
    {
        UpdateBreadcrumbLayout();
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::UpdateHoverTimerState() noexcept
{
    const bool shouldRun = _editMode || _inMenuLoop;
    if (! _hWnd)
    {
        _hoverTimer = 0;
        return;
    }

    if (shouldRun)
    {
        if (_hoverTimer == 0)
        {
            _hoverTimer = SetTimer(_hWnd.get(), HOVER_TIMER_ID, 1000 / HOVER_CHECK_FPS, nullptr);
        }
        return;
    }

    if (_hoverTimer != 0)
    {
        KillTimer(_hWnd.get(), HOVER_TIMER_ID);
        _hoverTimer = 0;
    }
}

void NavigationView::OnDestroy()
{
    // Kill timers
    if (_hoverTimer != 0)
    {
        KillTimer(_hWnd.get(), HOVER_TIMER_ID);
        _hoverTimer = 0;
    }

    StopSeparatorAnimation();

    CloseFullPathPopup();
    CloseEditSuggestPopup();

    if (_editSuggestThread.joinable())
    {
        _editSuggestThread.request_stop();
        _editSuggestCv.notify_all();
        _editSuggestThread.join();
    }

    // Clean up menu bitmaps (automatic with wil::unique_hbitmap)
    _menuBitmaps.clear();

    // Destroy child controls
    _pathEdit.reset();

    // Release Direct2D resources
    DiscardD2DResources();
}

void NavigationView::OnPaint()
{
    PAINTSTRUCT ps;
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd.get(), &ps);

    auto trace = std::format(
        L"[NavigationView] Paint rect: ({},{}) to ({},{}), editMode={}", ps.rcPaint.left, ps.rcPaint.top, ps.rcPaint.right, ps.rcPaint.bottom, _editMode);
    TRACER_CTX(trace.c_str());

    // Fill background
    FillRect(hdc.get(), &ps.rcPaint, _backgroundBrush.get());

    // Draw bottom border
    auto oldPen = wil::SelectObject(hdc.get(), _borderPen.get());
    MoveToEx(hdc.get(), 0, _clientSize.cy - 1, nullptr);
    LineTo(hdc.get(), _clientSize.cx, _clientSize.cy - 1);

    if (! _swapChain || ! _d2dTarget || ! _d2dContext)
    {
        if (! _deferredInitPosted && _hWnd)
        {
            _deferredInitPosted = PostMessageW(_hWnd.get(), WndMsg::kNavigationViewDeferredInit, 0, 0) != 0;
        }
        return;
    }

    // Render Section 1, 2, 3 & 4 with Direct2D
    _deferPresent      = true;
    _queuedPresentFull = false;
    _queuedPresentDirtyRect.reset();

    RenderDriveSection();
    RenderPathSection();
    RenderHistorySection();
    RenderDiskInfoSection();

    _deferPresent = false;

    if (_queuedPresentFull)
    {
        Present(std::nullopt);
    }
    else if (_queuedPresentDirtyRect.has_value())
    {
        RECT dirtyRect = _queuedPresentDirtyRect.value();
        Present(&dirtyRect);
    }

    _queuedPresentFull = false;
    _queuedPresentDirtyRect.reset();
}

void NavigationView::OnSize(UINT width, UINT height)
{
    if (width == 0 && height == 0)
        return; // init edge case

    _clientSize = {static_cast<LONG>(width), static_cast<LONG>(height)};
    // Calculate Layout
    // height is already DPI aware and adjust to the screen keep _clientSize.cy as is
    int scaledDriveSectionWidth    = MulDiv(kDriveSectionWidth, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI);
    int scaledDiskInfoSectionWidth = MulDiv(kDiskInfoSectionWidth, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI);
    int scaledHistoryWidth         = MulDiv(kHistoryButtonWidth, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI);

    if (! _showMenuSection)
    {
        scaledDriveSectionWidth = 0;
    }
    if (! _showDiskInfoSection)
    {
        scaledDiskInfoSectionWidth = 0;
    }

    // Section 1: Menu button (left)
    _sectionDriveRect = {0, 0, scaledDriveSectionWidth, _clientSize.cy};
    // Section 4: Disk info (right)
    _sectionDiskInfoRect = {_clientSize.cx - scaledDiskInfoSectionWidth, 0, _clientSize.cx, _clientSize.cy};
    // Section 2: Path display (middle)
    _sectionPathRect = {scaledDriveSectionWidth, 0, _clientSize.cx - scaledDiskInfoSectionWidth - scaledHistoryWidth, _clientSize.cy};
    // Section 3: History Button
    _sectionHistoryRect = {_sectionPathRect.right, 0, _sectionPathRect.right + scaledHistoryWidth, _clientSize.cy};

    bool hadSwapChain = static_cast<bool>(_swapChain);

    if (_d2dContext)
    {
        // Ensure resources so DirectWrite formats are ready before layout rebuild
        EnsureD2DResources();
    }

    // Recreate swap chain for new size (full window) if it already existed
    if (hadSwapChain && _swapChain)
    {
        _d2dContext->SetTarget(nullptr);
        _d2dTarget        = nullptr;
        auto bufferWidth  = static_cast<UINT>(_clientSize.cx);
        auto bufferHeight = static_cast<UINT>(_clientSize.cy);
        HRESULT hr        = _swapChain->ResizeBuffers(0, bufferWidth, bufferHeight, DXGI_FORMAT_UNKNOWN, 0);
        if (SUCCEEDED(hr))
        {
            _hasPresented = false; // Reset flag after ResizeBuffers
            wil::com_ptr<IDXGISurface> surface;
            hr = _swapChain->GetBuffer(0, IID_PPV_ARGS(&surface));
            if (SUCCEEDED(hr))
            {
                D2D1_BITMAP_PROPERTIES1 props = D2D1::BitmapProperties1(D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW,
                                                                        D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED),
                                                                        static_cast<float>(USER_DEFAULT_SCREEN_DPI),
                                                                        static_cast<float>(USER_DEFAULT_SCREEN_DPI));
                _d2dContext->CreateBitmapFromDxgiSurface(surface.get(), &props, &_d2dTarget);
            }
        }
    }

    if (_currentPluginPath && _d2dContext)
    {
        UpdateBreadcrumbLayout();
    }

    if (_pathEdit)
    {
        const auto chrome = ComputeEditChromeRects(_sectionPathRect, _dpi);
        LayoutSingleLineEditInRect(_pathEdit.get(), chrome.editRect);
    }

    if (_editSuggestPopup)
    {
        UpdateEditSuggestPopupWindow();
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::OnCommand(UINT id, [[maybe_unused]] HWND hwndCtl, UINT codeNotify)
{
    if (_editMode && id == ID_PATH_EDIT && codeNotify == EN_CHANGE && _pathEdit && hwndCtl == _pathEdit.get())
    {
        UpdateEditSuggest();
        return;
    }

    if (id == ID_NAV_DROPDOWN_COMBO && _navDropdownCombo && hwndCtl == _navDropdownCombo.get())
    {
        if (codeNotify == CBN_SELENDOK)
        {
            const int sel = static_cast<int>(SendMessageW(_navDropdownCombo.get(), CB_GETCURSEL, 0, 0));
            if (sel >= 0 && static_cast<size_t>(sel) < _navDropdownPaths.size())
            {
                const std::filesystem::path selectedPath = _navDropdownPaths[static_cast<size_t>(sel)];

                _navDropdownKind = ModernDropdownKind::None;
                _navDropdownPaths.clear();

                if (_menuOpenForSeparator != -1)
                {
                    _pendingSeparatorMenuSwitchIndex = -1;
                    StartSeparatorAnimation(static_cast<size_t>(_menuOpenForSeparator), 0.0f);
                    _menuOpenForSeparator = -1;
                    _activeSeparatorIndex = -1;
                    RenderPathSection();
                }

                SetWindowPos(_navDropdownCombo.get(), nullptr, -32000, -32000, 10, 10, SWP_NOZORDER | SWP_NOACTIVATE | SWP_HIDEWINDOW);
                RequestPathChange(selectedPath);
            }
            return;
        }

        if (codeNotify == CBN_SELENDCANCEL || codeNotify == CBN_CLOSEUP)
        {
            _navDropdownKind = ModernDropdownKind::None;
            _navDropdownPaths.clear();

            if (_menuOpenForSeparator != -1)
            {
                _pendingSeparatorMenuSwitchIndex = -1;
                StartSeparatorAnimation(static_cast<size_t>(_menuOpenForSeparator), 0.0f);
                _menuOpenForSeparator = -1;
                _activeSeparatorIndex = -1;
                RenderPathSection();
            }

            if (_requestFolderViewFocusCallback && _hWnd)
            {
                const HWND root = GetAncestor(_hWnd.get(), GA_ROOT);
                if (root && GetActiveWindow() == root)
                {
                    _requestFolderViewFocusCallback();
                }
            }

            if (_navDropdownCombo)
            {
                SetWindowPos(_navDropdownCombo.get(), nullptr, -32000, -32000, 10, 10, SWP_NOZORDER | SWP_NOACTIVATE | SWP_HIDEWINDOW);
            }
            return;
        }

        return;
    }

    if (ExecuteNavigationMenuAction(id))
    {
        _navigationMenuActions.clear();
        return;
    }

    if (ExecuteDriveMenuAction(id))
    {
        _driveMenuActions.clear();
        return;
    }

    // History button and disk static handlers removed - now handled in OnLButtonDown
    if (id >= ID_SIBLING_BASE)
    {
        // Handle sibling folder navigation - no limit on number of siblings
        // Actual navigation is handled in ShowSiblingsDropdown via TrackPopupMenu return value
    }
    else if (id == ID_PATH_EDIT && codeNotify == EN_KILLFOCUS)
    {
        ExitEditMode(false);
    }
}

void NavigationView::OnDpiChanged(float newDpi)
{
    _dpi = static_cast<UINT>(newDpi);
    IconCache::GetInstance().SetDpi(newDpi);

    InvalidateBreadcrumbLayoutCache();

    // Recreate fonts with new DPI
    int pathFontHeight = -MulDiv(12, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI);
    _pathFont.reset(CreateFontW(pathFontHeight,
                                0,
                                0,
                                0,
                                FW_NORMAL,
                                FALSE,
                                FALSE,
                                FALSE,
                                DEFAULT_CHARSET,
                                OUT_DEFAULT_PRECIS,
                                CLIP_DEFAULT_PRECIS,
                                CLEARTYPE_QUALITY,
                                DEFAULT_PITCH,
                                L"Segoe UI"));

    _menuFont    = CreateMenuFontForDpi(_dpi);
    _menuFontDpi = _dpi;

    // GDI menus are NOT DPI-aware - menu icon size does not change with DPI
    // It always stays at the system's base small icon size (96 DPI physical pixels)
    _menuIconSize = GetSystemMetrics(SM_CXSMICON);

    // Recreate DirectWrite resources
    _pathFormat      = nullptr;
    _separatorFormat = nullptr;
    EnsureD2DResources();

    // Regenerate menu icon bitmap at new DPI
    UpdateMenuIconBitmap();

    if (_navDropdownCombo)
    {
        HFONT fontToUse = _pathFont ? _pathFont.get() : (_menuFont ? _menuFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
        SendMessageW(_navDropdownCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(fontToUse), FALSE);
    }

    if (_pathEdit)
    {
        SendMessageW(_pathEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_pathFont.get()), TRUE);
        const auto chrome = ComputeEditChromeRects(_sectionPathRect, _dpi);
        LayoutSingleLineEditInRect(_pathEdit.get(), chrome.editRect);
    }

    if (_editSuggestPopup)
    {
        UpdateEditSuggestPopupWindow();
    }

    if (_fullPathPopupEdit)
    {
        SendMessageW(_fullPathPopupEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_pathFont.get()), TRUE);

        if (_fullPathPopup)
        {
            RECT rc{};
            GetClientRect(_fullPathPopup.get(), &rc);
            LayoutSingleLineEditInRect(_fullPathPopupEdit.get(), rc);
        }
    }

    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void NavigationView::SetPath(const std::optional<std::filesystem::path>& path)
{
    if (! path)
    {
        _currentPath       = std::nullopt;
        _currentPluginPath = std::nullopt;
        _currentEditPath   = std::nullopt;
        _currentInstanceContext.clear();
        _segments.clear();
        _separators.clear();
        _separatorRotationAngles.clear();
        _separatorTargetAngles.clear();
        InvalidateBreadcrumbLayoutCache();
        _hoveredSegmentIndex   = -1;
        _hoveredSeparatorIndex = -1;
        _activeSeparatorIndex  = -1;
        _menuOpenForSeparator  = -1;
        // Clear menu icon bitmap
        _menuIconBitmapD2D = nullptr;
        UpdateDiskInfo();
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        return;
    }

    NavigationLocation::Location location;
    const std::filesystem::path incomingPath = path.value();
    static_cast<void>(NavigationLocation::TryParseLocation(incomingPath.native(), location));

    const bool isFilePlugin = _pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file");

    if (isFilePlugin)
    {
        const std::filesystem::path normalizedPath = NormalizeDirectoryPath(incomingPath);
        _currentPath                               = normalizedPath;
        _currentPluginPath                         = normalizedPath;
        _currentEditPath                           = normalizedPath;
        _currentInstanceContext.clear();
    }
    else
    {
        std::wstring_view shortId = _pluginShortId;
        if (! location.pluginShortId.empty())
        {
            shortId = location.pluginShortId;
        }

        const std::filesystem::path pluginPath = location.pluginPath.empty() ? std::filesystem::path(L"/") : location.pluginPath;

        _currentInstanceContext = location.instanceContext;
        _currentPluginPath      = pluginPath;
        _currentEditPath        = NavigationLocation::FormatEditPath(shortId, pluginPath);
        _currentPath            = NavigationLocation::FormatHistoryPath(shortId, _currentInstanceContext, pluginPath);
    }

    if (! _dwriteFactory || ! _pathFormat || ! _separatorFormat)
    {
        EnsureD2DResources();
    }
    UpdateBreadcrumbLayout(); // Build layout when path changes
    if (_currentPluginPath.has_value())
    {
        QueueSiblingPrefetchForPath(_currentPluginPath.value());
    }

    UpdateMenuIconBitmap();
    UpdateDiskInfo();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::SetPathChangedCallback(PathChangedCallback callback)
{
    _pathChangedCallback = callback;
}

void NavigationView::SetRequestFolderViewFocusCallback(RequestFolderViewFocusCallback callback)
{
    _requestFolderViewFocusCallback = std::move(callback);
}

std::vector<std::filesystem::path> NavigationView::GetHistory() const
{
    return std::vector<std::filesystem::path>(_pathHistory.begin(), _pathHistory.end());
}

void NavigationView::SetHistory(const std::vector<std::filesystem::path>& history)
{
    _pathHistory.clear();

    for (const auto& entry : history)
    {
        if (entry.empty())
        {
            continue;
        }

        const bool exists = std::find(_pathHistory.begin(), _pathHistory.end(), entry) != _pathHistory.end();
        if (exists)
        {
            continue;
        }

        _pathHistory.push_back(entry);
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::SetFileSystem(const wil::com_ptr<IFileSystem>& fileSystem)
{
    if (_navigationMenu)
    {
        _navigationMenu->SetCallback(nullptr, nullptr);
    }

    _fileSystemPlugin = fileSystem;
    _fileSystemIo     = nullptr;
    _navigationMenu   = nullptr;
    _driveInfo        = nullptr;
    _pluginShortId.clear();

    if (_fileSystemPlugin)
    {
        wil::com_ptr<IInformations> informations;
        if (SUCCEEDED(_fileSystemPlugin->QueryInterface(__uuidof(IInformations), informations.put_void())) && informations)
        {
            const PluginMetaData* meta = nullptr;
            if (SUCCEEDED(informations->GetMetaData(&meta)) && meta && meta->shortId)
            {
                _pluginShortId = meta->shortId;
            }
        }

        _fileSystemPlugin->QueryInterface(__uuidof(INavigationMenu), _navigationMenu.put_void());
        _fileSystemPlugin->QueryInterface(__uuidof(IDriveInfo), _driveInfo.put_void());
        _fileSystemPlugin->QueryInterface(__uuidof(IFileSystemIO), _fileSystemIo.put_void());

        if (_navigationMenu)
        {
            _navigationMenu->SetCallback(this, _fileSystemPlugin.get()); // for sanity check
        }
    }

    _showMenuSection     = static_cast<bool>(_navigationMenu);
    _showDiskInfoSection = static_cast<bool>(_driveInfo);

    _menuButtonPressed = false;
    _menuButtonHovered = false;
    _diskInfoHovered   = false;
    _menuIconBitmapD2D = nullptr;
    _menuBitmaps.clear();
    _navigationMenuActions.clear();
    _driveMenuActions.clear();

    static_cast<void>(_editSuggestRequestId.fetch_add(1, std::memory_order_acq_rel));
    {
        std::lock_guard lock(_editSuggestMutex);
        _editSuggestPendingQuery.reset();
    }
    _editSuggestMountedInstance.reset();

    static_cast<void>(_siblingPrefetchRequestId.fetch_add(1, std::memory_order_acq_rel));
    {
        std::lock_guard lock(_siblingPrefetchMutex);
        _siblingPrefetchPendingQuery.reset();
    }
    _siblingPrefetchCv.notify_one();

    _editSuggestItems.clear();
    _editSuggestHighlightText.clear();
    CloseEditSuggestPopup();

    NormalizeFocusRegion();

    if (_clientSize.cx > 0 && _clientSize.cy > 0)
    {
        OnSize(static_cast<UINT>(_clientSize.cx), static_cast<UINT>(_clientSize.cy));
    }

    UpdateDiskInfo();
    if (_showMenuSection)
    {
        UpdateMenuIconBitmap();
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::SetSettings(Common::Settings::Settings* settings) noexcept
{
    _settings = settings;
}

void NavigationView::SetTheme(const AppTheme& theme)
{
    _appTheme  = theme;
    _baseTheme = _appTheme.navigationView;
    _menuTheme = _appTheme.menu;

    _menuBackgroundBrush.reset(CreateSolidBrush(_menuTheme.background));

    UpdateEffectiveTheme();
    InvalidateBreadcrumbLayoutCache();

    if (_d2dContext)
    {
        EnsureD2DResources();
    }

    if (_editSuggestPopup)
    {
        DiscardEditSuggestPopupD2DResources();
        InvalidateRect(_editSuggestPopup.get(), nullptr, TRUE);
    }

    if (_currentPluginPath.has_value() && _clientSize.cx > 0 && _clientSize.cy > 0)
    {
        UpdateBreadcrumbLayout();
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    if (_navDropdownCombo)
    {
        ThemedControls::ApplyThemeToComboBox(_navDropdownCombo.get(), _appTheme);
    }
}

void NavigationView::SetPaneFocused(bool focused) noexcept
{
    if (_paneFocused == focused)
    {
        return;
    }

    _paneFocused = focused;
    UpdateEffectiveTheme();

    if (_d2dContext)
    {
        EnsureD2DResources();
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::UpdateEffectiveTheme() noexcept
{
    _theme = _baseTheme;

    if (_paneFocused)
    {
        _theme.gdiBorderPen = ColorToCOLORREF(_theme.accent);
    }
    else
    {
        const float textBlend   = _theme.darkBase ? 0.45f : 0.35f;
        const float sepBlend    = _theme.darkBase ? 0.65f : 0.55f;
        const float accentBlend = _theme.darkBase ? 0.50f : 0.40f;

        if (! _theme.darkBase)
        {
            const float bgBlend         = 0.06f;
            const D2D1::ColorF baseText = _theme.text;
            _theme.background           = BlendColorF(_theme.background, baseText, bgBlend);
            _theme.backgroundHover      = BlendColorF(_theme.backgroundHover, baseText, bgBlend);
            _theme.backgroundPressed    = BlendColorF(_theme.backgroundPressed, baseText, bgBlend);
            _theme.hoverHighlight       = BlendColorF(_theme.hoverHighlight, baseText, bgBlend);
            _theme.pressedHighlight     = BlendColorF(_theme.pressedHighlight, baseText, bgBlend);
        }

        _theme.text               = BlendColorF(_theme.text, _theme.background, textBlend);
        _theme.separator          = BlendColorF(_theme.separator, _theme.background, sepBlend);
        _theme.accent             = BlendColorF(_theme.accent, _theme.background, accentBlend);
        _theme.progressOk         = BlendColorF(_theme.progressOk, _theme.background, accentBlend);
        _theme.progressWarn       = BlendColorF(_theme.progressWarn, _theme.background, accentBlend);
        _theme.progressBackground = BlendColorF(_theme.progressBackground, _theme.background, std::max(accentBlend, 0.65f));

        const float borderBlend = _theme.darkBase ? 0.70f : 0.82f;
        _theme.gdiBorderPen     = BlendColorRef(_theme.gdiBorderPen, ColorToCOLORREF(_theme.background), borderBlend);
    }

    _theme.gdiBackground = ColorToCOLORREF(_theme.background);
    _theme.gdiBorder     = _theme.gdiBackground;

    _backgroundBrush.reset(CreateSolidBrush(_theme.gdiBackground));
    _borderBrush.reset(CreateSolidBrush(_theme.gdiBorder));
    _borderPen.reset(CreatePen(PS_SOLID, 1, _theme.gdiBorderPen));
}

void NavigationView::SetFocusRegion(FocusRegion region)
{
    _focusedRegion = region;
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::FocusAddressBar()
{
    SetFocusRegion(FocusRegion::Path);
    if (_hWnd)
    {
        SetFocus(_hWnd.get());
    }
    EnterEditMode();
}

void NavigationView::OpenChangeDirectoryFromCommand()
{
    SetFocusRegion(FocusRegion::Path);
    if (_hWnd)
    {
        SetFocus(_hWnd.get());
    }

    const bool isFilePlugin = _pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file");
    if (! isFilePlugin && ! _currentInstanceContext.empty())
    {
        const std::filesystem::path pluginPath = _currentPluginPath.has_value() ? _currentPluginPath.value() : std::filesystem::path(L"/");
        const std::wstring pluginPathText      = NavigationLocation::NormalizePluginPathText(pluginPath.wstring());

        std::wstring editText;
        editText.reserve(_currentInstanceContext.size() + 1u + pluginPathText.size());
        editText.append(_currentInstanceContext);
        editText.push_back(L'|');
        editText.append(pluginPathText);
        _currentEditPath = std::filesystem::path(std::move(editText));
    }

    EnterEditMode();
}

void NavigationView::OpenHistoryDropdownFromKeyboard()
{
    SetFocusRegion(FocusRegion::History);
    if (_hWnd)
    {
        SetFocus(_hWnd.get());
    }
    ShowHistoryDropdown();
}

// Direct2D implementation continues in next part...
