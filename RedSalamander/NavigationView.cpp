#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cwctype>
#include <format>
#include <iterator>
#include <limits>
#include <new>
#include <optional>
#include <string_view>

#include <windows.h>
#include <windowsx.h>

#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "dwrite.lib")
#pragma comment(lib, "dxgi.lib")

#include "NavigationView.h"
#include "NavigationViewInternal.h"

#include "D2DHdcPaint.h"
#include "DxUi/DxUi.Typography.h"
#include "NavigationLocation.h"

#include "DirectoryInfoCache.h"
#include "Helpers.h"
#include "IconCache.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "UiMetrics.h"
#include "resource.h"

namespace
{
[[maybe_unused]] [[nodiscard]] HWND FindOwnedVisibleDxContextMenuWindow(HWND ownerHwnd) noexcept
{
    if (! ownerHwnd)
    {
        return nullptr;
    }

    const HWND rootOwner = GetAncestor(ownerHwnd, GA_ROOT);
    for (HWND popup = FindWindowW(L"DxUi_ContextMenu", nullptr); popup != nullptr; popup = FindWindowExW(nullptr, popup, L"DxUi_ContextMenu", nullptr))
    {
        if (IsWindowVisible(popup) == FALSE)
        {
            continue;
        }

        const HWND popupOwner = GetWindow(popup, GW_OWNER);
        if (popupOwner == ownerHwnd || (rootOwner && popupOwner == rootOwner))
        {
            return popup;
        }
    }

    return nullptr;
}

[[nodiscard]] const wchar_t* TraceNavigationWindowMessageName(UINT message) noexcept
{
    switch (message)
    {
        case WM_NCHITTEST: return L"WM_NCHITTEST";
        case WM_MOUSEACTIVATE: return L"WM_MOUSEACTIVATE";
        case WM_SETCURSOR: return L"WM_SETCURSOR";
        case WM_MOUSEMOVE: return L"WM_MOUSEMOVE";
        case WM_MOUSELEAVE: return L"WM_MOUSELEAVE";
        case WM_LBUTTONDOWN: return L"WM_LBUTTONDOWN";
        case WM_LBUTTONUP: return L"WM_LBUTTONUP";
        case WM_LBUTTONDBLCLK: return L"WM_LBUTTONDBLCLK";
        case WM_RBUTTONDOWN: return L"WM_RBUTTONDOWN";
        case WM_RBUTTONUP: return L"WM_RBUTTONUP";
        case WM_CAPTURECHANGED: return L"WM_CAPTURECHANGED";
        case WM_CANCELMODE: return L"WM_CANCELMODE";
        case WM_SETFOCUS: return L"WM_SETFOCUS";
        case WM_KILLFOCUS: return L"WM_KILLFOCUS";
        case WM_KEYDOWN: return L"WM_KEYDOWN";
        case WM_SYSKEYDOWN: return L"WM_SYSKEYDOWN";
        case WndMsg::kNavigationViewShowHistoryDropdown: return L"kNavigationViewShowHistoryDropdown";
        case WndMsg::kNavigationViewShowMenuDropdown: return L"kNavigationViewShowMenuDropdown";
        case WndMsg::kNavigationViewShowDiskInfoDropdown: return L"kNavigationViewShowDiskInfoDropdown";
        case WndMsg::kNavigationViewShowDriveMenuDropdown: return L"kNavigationViewShowDriveMenuDropdown";
        case WndMsg::kNavigationMenuShowSiblingsDropdown: return L"kNavigationMenuShowSiblingsDropdown";
        default: return L"message";
    }
}

[[nodiscard]] bool ShouldTraceNavigationWindowRaw(UINT message) noexcept
{
    switch (message)
    {
        case WM_MOUSEACTIVATE:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP:
        case WM_CAPTURECHANGED:
        case WM_CANCELMODE:
        case WM_SETFOCUS:
        case WM_KILLFOCUS:
        case WM_KEYDOWN:
        case WM_SYSKEYDOWN:
        case WndMsg::kNavigationViewShowHistoryDropdown:
        case WndMsg::kNavigationViewShowMenuDropdown:
        case WndMsg::kNavigationViewShowDiskInfoDropdown:
        case WndMsg::kNavigationViewShowDriveMenuDropdown:
        case WndMsg::kNavigationMenuShowSiblingsDropdown: return true;
        default: return false;
    }
}

[[nodiscard]] const wchar_t* TraceNavigationResultName(UINT message, LRESULT result) noexcept
{
    if (message == WM_MOUSEACTIVATE)
    {
        switch (result)
        {
            case MA_ACTIVATE: return L"MA_ACTIVATE";
            case MA_ACTIVATEANDEAT: return L"MA_ACTIVATEANDEAT";
            case MA_NOACTIVATE: return L"MA_NOACTIVATE";
            case MA_NOACTIVATEANDEAT: return L"MA_NOACTIVATEANDEAT";
            default: return L"result";
        }
    }

    if (message == WM_NCHITTEST)
    {
        switch (result)
        {
            case HTCLIENT: return L"HTCLIENT";
            case HTTRANSPARENT: return L"HTTRANSPARENT";
            case HTNOWHERE: return L"HTNOWHERE";
            default: return L"hit";
        }
    }

    return result != 0 ? L"nonzero" : L"zero";
}

void TraceNavigationWindowRaw(
    HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam, std::wstring_view phase, std::optional<LRESULT> result = std::nullopt) noexcept
{
    if (! RedSalamander::DxUi::IsContextMenuDiagnosticsEnabled())
    {
        return;
    }

    POINT cursorScreen{};
    const bool haveCursor       = GetCursorPos(&cursorScreen) != FALSE; // getcursorpos-allow: diagnostic-only
    POINT cursorClient          = cursorScreen;
    const bool haveCursorClient = hwnd && haveCursor && ScreenToClient(hwnd, &cursorClient) != FALSE;

    POINT messageScreen{};
    POINT messageClient{};
    bool haveMessageScreen = false;
    bool haveMessageClient = false;
    if (message == WM_NCHITTEST)
    {
        messageScreen     = POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        messageClient     = messageScreen;
        haveMessageScreen = true;
        haveMessageClient = hwnd && ScreenToClient(hwnd, &messageClient) != FALSE;
    }
    else if ((message >= WM_MOUSEFIRST && message <= WM_MOUSELAST) || message == WM_MOUSELEAVE)
    {
        messageClient     = POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
        messageScreen     = messageClient;
        haveMessageClient = true;
        haveMessageScreen = hwnd && ClientToScreen(hwnd, &messageScreen) != FALSE;
    }

    RECT clientRect{};
    const bool haveClientRect  = hwnd && GetClientRect(hwnd, &clientRect) != FALSE;
    const bool cursorInClient  = haveCursorClient && haveClientRect && PtInRect(&clientRect, cursorClient) != FALSE;
    const bool messageInClient = haveMessageClient && haveClientRect && PtInRect(&clientRect, messageClient) != FALSE;

    const HWND windowAtCursor = haveCursor ? WindowFromPoint(cursorScreen) : nullptr;
    HWND childAtCursor        = nullptr;
    if (hwnd && haveCursor)
    {
        const HWND root = GetAncestor(hwnd, GA_ROOT);
        if (root)
        {
            POINT rootClient = cursorScreen;
            if (ScreenToClient(root, &rootClient) != FALSE)
            {
                childAtCursor = ChildWindowFromPointEx(root, rootClient, CWP_SKIPINVISIBLE);
            }
        }
    }

    TraceNavigationViewMenuDiagnostics(L"navigation.wndproc.raw",
                                       L"phase={} hwnd={:#x} msg={} msgId={:#x} wParam={:#x} lParam={:#x} "
                                       L"cursorScreen=({}, {}) haveCursor={} cursorClient=({}, {}) cursorInClient={} "
                                       L"messageScreen=({}, {}) haveMessageScreen={} messageClient=({}, {}) messageInClient={} "
                                       L"windowAtCursor={:#x} childAtCursor={:#x} focus={:#x} active={:#x} foreground={:#x} capture={:#x} "
                                       L"result={} resultName={} ownedMenu={:#x}",
                                       phase,
                                       reinterpret_cast<uintptr_t>(hwnd),
                                       TraceNavigationWindowMessageName(message),
                                       static_cast<unsigned int>(message),
                                       static_cast<uintptr_t>(wParam),
                                       static_cast<uintptr_t>(lParam),
                                       cursorScreen.x,
                                       cursorScreen.y,
                                       haveCursor ? 1 : 0,
                                       cursorClient.x,
                                       cursorClient.y,
                                       cursorInClient ? 1 : 0,
                                       messageScreen.x,
                                       messageScreen.y,
                                       haveMessageScreen ? 1 : 0,
                                       messageClient.x,
                                       messageClient.y,
                                       messageInClient ? 1 : 0,
                                       reinterpret_cast<uintptr_t>(windowAtCursor),
                                       reinterpret_cast<uintptr_t>(childAtCursor),
                                       reinterpret_cast<uintptr_t>(GetFocus()),
                                       reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                       reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                       reinterpret_cast<uintptr_t>(GetCapture()),
                                       result.has_value() ? static_cast<int>(result.value()) : 0,
                                       result.has_value() ? TraceNavigationResultName(message, result.value()) : L"none",
                                       reinterpret_cast<uintptr_t>(FindOwnedVisibleDxContextMenuWindow(hwnd)));
}

[[nodiscard]] bool OptionalPathEquals(const std::optional<std::filesystem::path>& lhs, const std::optional<std::filesystem::path>& rhs) noexcept
{
    if (lhs.has_value() != rhs.has_value())
    {
        return false;
    }

    if (! lhs.has_value())
    {
        return true;
    }

    return OrdinalString::EqualsNoCasePath(lhs.value(), rhs.value());
}

[[nodiscard]] bool OptionalPathTextEquals(const std::optional<std::filesystem::path>& lhs, const std::optional<std::filesystem::path>& rhs) noexcept
{
    if (lhs.has_value() != rhs.has_value())
    {
        return false;
    }

    if (! lhs.has_value())
    {
        return true;
    }

    return lhs->native() == rhs->native();
}
} // namespace

NavigationView::NavigationView() = default;

NavigationView::~NavigationView()
{
    Destroy();
}

bool NavigationView::ShouldAcceptPointerEvent(const RedSalamander::DxUi::PointerInputEvent& event) const noexcept
{
    const HWND hwnd = _hWnd.get();
    if (! hwnd || ! event.targetHwnd || (event.targetHwnd != hwnd && IsChild(hwnd, event.targetHwnd) == FALSE))
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.pointer.reject",
                                           L"hwnd={:#x} reason=target target={:#x}",
                                           reinterpret_cast<uintptr_t>(hwnd),
                                           reinterpret_cast<uintptr_t>(event.targetHwnd));
        return false;
    }

    if (! event.hasClientPoint)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.pointer.reject",
                                           L"hwnd={:#x} reason=no-client-point target={:#x}",
                                           reinterpret_cast<uintptr_t>(hwnd),
                                           reinterpret_cast<uintptr_t>(event.targetHwnd));
        return false;
    }

    return true;
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

void NavigationView::RequestOwnerPaneFocus() const noexcept
{
    if (_requestFolderViewFocusCallback)
    {
        _requestFolderViewFocusCallback();
    }
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

ATOM NavigationView::RegisterDxHostWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = DxHostWndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_IBEAM);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kDxHostClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

LRESULT CALLBACK NavigationView::DxHostWndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* host = reinterpret_cast<RedSalamander::DxUi::WindowHost*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE)
    {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        host     = reinterpret_cast<RedSalamander::DxUi::WindowHost*>(cs->lpCreateParams);
        if (! host || ! host->Attach(hwnd))
        {
            return FALSE;
        }

        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(host));
    }

    if (! host)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    bool handled     = false;
    const LRESULT dx = host->HandleMessage(hwnd, msg, wp, lp, handled);
    if (msg == WM_NCDESTROY)
    {
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
    }

    return handled ? dx : DefWindowProcW(hwnd, msg, wp, lp);
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

    StopDriveInfoWorker();

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
    if (ShouldTraceNavigationWindowRaw(msg))
    {
        TraceNavigationWindowRaw(hwnd, msg, wp, lp, L"enter");
    }

    switch (msg)
    {
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WndMsg::kNavigationViewDeferredInit: OnDeferredInit(); return 0;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_NCDESTROY: static_cast<void>(DrainPostedPayloadsForWindow(hwnd)); break;
        case WM_ERASEBKGND: return 1;
        case WM_NCHITTEST: return HTCLIENT;
        case WM_MOUSEACTIVATE: TraceNavigationWindowRaw(hwnd, msg, wp, lp, L"return", MA_ACTIVATE); return MA_ACTIVATE;
        case WM_PAINT: OnPaint(); return 0;
        case WM_SETREDRAW:
        {
            const bool enableRedraw = wp != FALSE;
            if (! enableRedraw)
            {
                _redrawSuspended            = true;
                _redrawSuspendedUntilTickMs = GetTickCount64() + 1000u;
            }
            const LRESULT result = DefWindowProcW(hwnd, msg, wp, lp);
            if (enableRedraw)
            {
                if (! _embeddedDestinationMode && _editMode)
                {
                    _pathEditBlurSuppressActive      = true;
                    _pathEditBlurSuppressUntilTickMs = GetTickCount64() + 2000u;
                }
                _redrawSuspended            = false;
                _redrawSuspendedUntilTickMs = 0;
                RefreshActiveEditHostAfterParentPaint();
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            return result;
        }
        case WM_SHOWWINDOW:
            if (wp != FALSE)
            {
                RefreshActiveEditHostAfterParentPaint();
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            return 0;
        case WM_DISPLAYCHANGE:
        case WM_SETTINGCHANGE:
        case WM_THEMECHANGED:
            ApplyDxEditHostThemes();
            UpdatePathEditHostLayout();
            UpdateFullPathPopupEditHostLayout();
            RefreshActiveEditHostAfterParentPaint();
            InvalidateRect(hwnd, nullptr, FALSE);
            return 0;
        case WM_DPICHANGED_AFTERPARENT: OnDpiChanged(static_cast<float>(GetDpiForWindow(hwnd))); return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_COMMAND: OnCommand(LOWORD(wp), reinterpret_cast<HWND>(lp), HIWORD(wp)); return 0;
        case WM_CTLCOLOREDIT: return OnCtlColorEdit(reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_LBUTTONDOWN:
        {
            if (auto event = RedSalamander::DxUi::TryBuildPointerInputEvent(hwnd, msg, wp, lp); event.has_value())
            {
                OnLButtonDown(event.value());
            }
            return 0;
        }
        case WM_LBUTTONDBLCLK:
        {
            if (auto event = RedSalamander::DxUi::TryBuildPointerInputEvent(hwnd, msg, wp, lp); event.has_value())
            {
                OnLButtonDblClk(event.value());
            }
            return 0;
        }
        case WM_MOUSEMOVE:
        {
            if (auto event = RedSalamander::DxUi::TryBuildPointerInputEvent(hwnd, msg, wp, lp); event.has_value())
            {
                OnMouseMove(event.value());
            }
            return 0;
        }
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_SETCURSOR: OnSetCursor(reinterpret_cast<HWND>(wp), LOWORD(lp), HIWORD(lp)); return TRUE;
        case WM_TIMER: OnTimer(static_cast<UINT_PTR>(wp)); return 0;
        case WM_CANCELMODE: TraceNavigationInputState(L"cancel-mode"); break;
        case WM_CAPTURECHANGED:
            TraceNavigationViewMenuDiagnostics(L"navigation.capture-changed",
                                               L"hwnd={:#x} newCapture={:#x} focus={:#x} active={:#x} foreground={:#x}",
                                               reinterpret_cast<uintptr_t>(hwnd),
                                               reinterpret_cast<uintptr_t>(reinterpret_cast<HWND>(lp)),
                                               reinterpret_cast<uintptr_t>(GetFocus()),
                                               reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                               reinterpret_cast<uintptr_t>(GetForegroundWindow()));
            break;
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
            TraceNavigationViewMenuDiagnostics(L"navigation.siblings-dropdown.message",
                                               L"hwnd={:#x} index={} focus={:#x} active={:#x} foreground={:#x} capture={:#x}",
                                               reinterpret_cast<uintptr_t>(hwnd),
                                               static_cast<size_t>(wp),
                                               reinterpret_cast<uintptr_t>(GetFocus()),
                                               reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                               reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                               reinterpret_cast<uintptr_t>(GetCapture()));
            _pendingSeparatorMenuSwitchIndex = -1;
            ShowSiblingsDropdown(static_cast<size_t>(wp));
            return 0;                                                            // Deferred menu opening
        case WndMsg::kNavigationMenuShowFullPath: ShowFullPathPopup(); return 0; // Deferred full-path popup opening
        case WndMsg::kNavigationViewShowHistoryDropdown: ShowHistoryDropdown(false, wp != 0); return 0;
        case WndMsg::kNavigationViewShowMenuDropdown: ShowMenuDropdown(false, wp != 0); return 0;
        case WndMsg::kNavigationViewShowDiskInfoDropdown: ShowDiskInfoDropdown(false, wp != 0); return 0;
        case WndMsg::kNavigationViewShowDriveMenuDropdown: ShowFileSystemDriveMenuDropdown(); return 0;
        case WndMsg::kNavigationViewDriveInfoLoaded:
        {
            auto payload = TakeMessagePayload<DriveInfoResultPayload>(lp);
            return OnDriveInfoLoaded(std::move(payload));
        }
        case WndMsg::kNavigationViewRestoreFolderFocus:
            if (_editMode || _fullPathPopupEditMode)
            {
                return 0;
            }
            if (_requestFolderViewFocusCallback && _hWnd)
            {
                const HWND root = GetAncestor(_hWnd.get(), GA_ROOT);
                if (! root)
                {
                    return 0;
                }
                if (const HWND activeWindow = GetActiveWindow(); activeWindow && activeWindow != root)
                {
                    return 0;
                }
                if (const HWND foregroundWindow = GetForegroundWindow(); foregroundWindow && foregroundWindow != root)
                {
                    return 0;
                }

                _requestFolderViewFocusCallback();
            }
            return 0;
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

    {
        Debug::Perf::Scope perfTheme(L"NavigationView.OnCreate.SetTheme");
        const AppTheme resolvedTheme = ResolveAppTheme(ThemeMode::System, L"");
        SetTheme(resolvedTheme);
    }

    // GDI menus are NOT DPI-aware - they always expect physical pixels at 96 DPI
    // Do not scale menu icon size with DPI
    _menuIconSize = GetSystemMetrics(SM_CXSMICON);
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
        if (_showMenuSection && _currentPluginPath)
        {
            UpdateMenuIconBitmap();
        }
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
    const bool shouldRun = _inMenuLoop;
    if (! _hWnd)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-timer-state",
                                           L"hwnd=null action=no-window shouldRun={} editMode={} inMenuLoop={} oldTimer={}",
                                           shouldRun ? 1 : 0,
                                           _editMode ? 1 : 0,
                                           _inMenuLoop ? 1 : 0,
                                           static_cast<unsigned long long>(_hoverTimer));
        _hoverTimer = 0;
        return;
    }

    if (shouldRun)
    {
        if (_hoverTimer == 0)
        {
            const UINT_PTR oldTimer = _hoverTimer;
            _hoverTimer             = SetTimer(_hWnd.get(), HOVER_TIMER_ID, 1000 / HOVER_CHECK_FPS, nullptr);
            const DWORD error       = _hoverTimer == 0 ? GetLastError() : ERROR_SUCCESS;
            TraceNavigationViewMenuDiagnostics(L"navigation.hover-timer-state",
                                               L"hwnd={:#x} action=start shouldRun={} editMode={} inMenuLoop={} oldTimer={} newTimer={} lastError={}",
                                               reinterpret_cast<uintptr_t>(_hWnd.get()),
                                               shouldRun ? 1 : 0,
                                               _editMode ? 1 : 0,
                                               _inMenuLoop ? 1 : 0,
                                               static_cast<unsigned long long>(oldTimer),
                                               static_cast<unsigned long long>(_hoverTimer),
                                               error);
        }
        else
        {
            TraceNavigationViewMenuDiagnostics(L"navigation.hover-timer-state",
                                               L"hwnd={:#x} action=keep shouldRun={} editMode={} inMenuLoop={} timer={}",
                                               reinterpret_cast<uintptr_t>(_hWnd.get()),
                                               shouldRun ? 1 : 0,
                                               _editMode ? 1 : 0,
                                               _inMenuLoop ? 1 : 0,
                                               static_cast<unsigned long long>(_hoverTimer));
        }
        return;
    }

    if (_hoverTimer != 0)
    {
        const UINT_PTR oldTimer = _hoverTimer;
        KillTimer(_hWnd.get(), HOVER_TIMER_ID);
        _hoverTimer = 0;
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-timer-state",
                                           L"hwnd={:#x} action=stop shouldRun={} editMode={} inMenuLoop={} oldTimer={} newTimer=0",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           shouldRun ? 1 : 0,
                                           _editMode ? 1 : 0,
                                           _inMenuLoop ? 1 : 0,
                                           static_cast<unsigned long long>(oldTimer));
    }
    else
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.hover-timer-state",
                                           L"hwnd={:#x} action=idle shouldRun={} editMode={} inMenuLoop={} timer=0",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           shouldRun ? 1 : 0,
                                           _editMode ? 1 : 0,
                                           _inMenuLoop ? 1 : 0);
    }
}

void NavigationView::OnDestroy()
{
    _redrawSuspended            = false;
    _redrawSuspendedUntilTickMs = 0;
    _pathEditBlurSuppressActive = false;

    // Kill timers
    if (_hoverTimer != 0)
    {
        KillTimer(_hWnd.get(), HOVER_TIMER_ID);
        _hoverTimer = 0;
    }

    StopSeparatorAnimation();

    CloseFullPathPopup();
    CloseEditSuggestPopup();
    CloseEditValidationPopup();

    if (_editSuggestThread.joinable())
    {
        _editSuggestThread.request_stop();
        _editSuggestCv.notify_all();
        _editSuggestThread.join();
    }

    // Clean up menu bitmaps (automatic with wil::unique_hbitmap)
    _menuBitmaps.clear();

    if (_pathEdit && _pathEdit->field)
    {
        _pathEdit->field->SetOnTextChanged({});
        _pathEdit->field->SetOnBlur({});
    }
    _pathEdit.reset();
    _fullPathPopupEdit.reset();

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

    if (! _embeddedDestinationMode)
    {
        RECT client{};
        GetClientRect(_hWnd.get(), &client);
        D2DHdcPaint::Session borderPaint;
        if (borderPaint.Begin(hdc.get(), client))
        {
            const float y = static_cast<float>(std::max<LONG>(0, _clientSize.cy - 1));
            borderPaint.DrawLine(0.0f, y, static_cast<float>(std::max<LONG>(0, _clientSize.cx)), y, _theme.gdiBorderPen);
        }
    }

    if (! _swapChain || ! _d2dTarget || ! _d2dContext)
    {
        if (! _deferredInitPosted && _hWnd)
        {
            _deferredInitPosted = PostMessageW(_hWnd.get(), WndMsg::kNavigationViewDeferredInit, 0, 0) != 0;
        }
        RefreshActiveEditHostAfterParentPaint();
        return;
    }

    // Render Section 1, 2, 3 & 4 with Direct2D. Hover is owned by delivered pointer messages;
    // paint must not sample the global cursor and silently overwrite that state.
    _deferPresent      = true;
    _queuedPresentFull = false;
    _queuedPresentDirtyRect.reset();
    TraceNavigationViewMenuDiagnostics(L"navigation.paint",
                                       L"hwnd={:#x} action=defer-present-start paintRect=({}, {}, {}, {})",
                                       reinterpret_cast<uintptr_t>(_hWnd.get()),
                                       ps.rcPaint.left,
                                       ps.rcPaint.top,
                                       ps.rcPaint.right,
                                       ps.rcPaint.bottom);

    RenderDriveSection();
    RenderPathSection();
    RenderHistorySection();
    RenderDiskInfoSection();

    _deferPresent = false;

    if (_queuedPresentFull)
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.paint", L"hwnd={:#x} action=flush-full", reinterpret_cast<uintptr_t>(_hWnd.get()));
        Present(std::nullopt);
    }
    else if (_queuedPresentDirtyRect.has_value())
    {
        RECT dirtyRect = _queuedPresentDirtyRect.value();
        TraceNavigationViewMenuDiagnostics(L"navigation.paint",
                                           L"hwnd={:#x} action=flush-dirty rect=({}, {}, {}, {})",
                                           reinterpret_cast<uintptr_t>(_hWnd.get()),
                                           dirtyRect.left,
                                           dirtyRect.top,
                                           dirtyRect.right,
                                           dirtyRect.bottom);
        Present(&dirtyRect);
    }
    else
    {
        TraceNavigationViewMenuDiagnostics(L"navigation.paint", L"hwnd={:#x} action=flush-none", reinterpret_cast<uintptr_t>(_hWnd.get()));
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

    UpdatePathEditHostLayout();

    if (_editSuggestPopup)
    {
        UpdateEditSuggestPopupWindow();
    }

    UpdateFullPathPopupEditHostLayout();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::OnCommand(UINT id, [[maybe_unused]] HWND hwndCtl, [[maybe_unused]] UINT codeNotify)
{
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
        // Handle sibling folder navigation - selection is dispatched directly by the DxUi sibling dropdown.
    }
}

void NavigationView::OnDpiChanged(float newDpi)
{
    Debug::Perf::Scope perf(L"navigation.ui.dpi_change_us");
    _dpi = static_cast<UINT>(newDpi);
    IconCache::GetInstance().SetDpi(newDpi);

    InvalidateBreadcrumbLayoutCache();

    // GDI menus are NOT DPI-aware - menu icon size does not change with DPI
    // It always stays at the system's base small icon size (96 DPI physical pixels)
    _menuIconSize = GetSystemMetrics(SM_CXSMICON);

    // Recreate DirectWrite resources
    _pathFormat      = nullptr;
    _separatorFormat = nullptr;
    EnsureD2DResources();

    // Regenerate menu icon bitmap at new DPI
    UpdateMenuIconBitmap();

    ApplyDxEditHostThemes();

    // Re-run the full section/breadcrumb layout at the new DPI. The host window may have
    // already resized this child during its own WM_DPICHANGED handling (before our DPI state
    // was updated), which rebuilt the breadcrumb segments with stale-DPI fonts; OnSize
    // rebuilds the section rects, swap-chain target, breadcrumb segments, and the edit host
    // from the current client size, so the order of resize vs. DPI notification cannot leave
    // stale-scale content behind.
    if (_clientSize.cx > 0 && _clientSize.cy > 0)
    {
        OnSize(static_cast<UINT>(_clientSize.cx), static_cast<UINT>(_clientSize.cy));
    }
    else
    {
        UpdatePathEditHostLayout();
    }

    if (_editSuggestPopup)
    {
        DiscardEditSuggestPopupD2DResources();
        UpdateEditSuggestPopupWindow();
    }

    if (_fullPathPopup)
    {
        DiscardFullPathPopupD2DResources();
        UpdateFullPathPopupWindow();
    }

    UpdateFullPathPopupEditHostLayout();

    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void NavigationView::SetPath(const std::optional<std::filesystem::path>& path)
{
    if (! path)
    {
        if (_editMode)
        {
            ExitEditMode(false, L"path-clear");
        }
        CloseFullPathPopup();

        if (! _currentPath.has_value() && ! _currentPluginPath.has_value() && ! _currentEditPath.has_value() && _currentInstanceContext.empty() &&
            _segments.empty() && _separators.empty() && _hoveredSegmentIndex == -1 && _hoveredSeparatorIndex == -1)
        {
            return;
        }

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
    std::optional<std::filesystem::path> nextCurrentPath;
    std::optional<std::filesystem::path> nextPluginPath;
    std::optional<std::filesystem::path> nextEditPath;
    std::wstring nextInstanceContext;

    if (isFilePlugin)
    {
        const std::filesystem::path normalizedPath = NormalizeDirectoryPath(incomingPath);
        nextCurrentPath                            = normalizedPath;
        nextPluginPath                             = normalizedPath;
        nextEditPath                               = normalizedPath;
    }
    else
    {
        std::wstring_view shortId = _pluginShortId;
        if (! location.pluginShortId.empty())
        {
            shortId = location.pluginShortId;
        }

        const std::filesystem::path pluginPath = location.pluginPath.empty() ? std::filesystem::path(L"/") : location.pluginPath;

        nextInstanceContext = location.instanceContext;
        nextPluginPath      = pluginPath;
        nextEditPath        = NavigationLocation::FormatEditPath(shortId, pluginPath);
        nextCurrentPath     = NavigationLocation::FormatHistoryPath(shortId, nextInstanceContext, pluginPath);
    }

    const bool sameLocation = OptionalPathEquals(_currentPath, nextCurrentPath) && OptionalPathEquals(_currentPluginPath, nextPluginPath) &&
                              _currentInstanceContext == nextInstanceContext;
    const bool sameEditPath = OptionalPathEquals(_currentEditPath, nextEditPath);
    const bool samePath     = sameLocation && (_editMode || sameEditPath);
    // `samePath` is case-insensitive: it decides location identity (edit mode/popup handling below).
    // While path edit is active, `_currentEditPath` tracks user-typed text, not the pane's identity.
    // A same-location refresh must not retire edit mode just because the user has edited the buffer.
    // The short-circuit additionally requires byte-equal text so a case-only change (e.g. after a
    // case-only rename of the current location) still refreshes the displayed breadcrumb strings.
    const bool samePathText = samePath && OptionalPathTextEquals(_currentPath, nextCurrentPath) && OptionalPathTextEquals(_currentPluginPath, nextPluginPath) &&
                              (_editMode || OptionalPathTextEquals(_currentEditPath, nextEditPath));
    if (samePathText && _breadcrumbLayoutCacheValid)
    {
        return;
    }

    if (! samePath)
    {
        if (_editMode)
        {
            ExitEditMode(false, L"path-change");
        }
        CloseFullPathPopup();
    }

    _currentPath       = std::move(nextCurrentPath);
    _currentPluginPath = std::move(nextPluginPath);
    if (! _editMode || ! sameLocation)
    {
        _currentEditPath = std::move(nextEditPath);
    }
    _currentInstanceContext = std::move(nextInstanceContext);

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
    std::deque<std::filesystem::path> nextHistory;

    for (const auto& entry : history)
    {
        if (entry.empty())
        {
            continue;
        }

        const bool exists = std::find(nextHistory.begin(), nextHistory.end(), entry) != nextHistory.end();
        if (exists)
        {
            continue;
        }

        nextHistory.push_back(entry);
    }

    if (_pathHistory.size() == nextHistory.size() && std::equal(_pathHistory.begin(), _pathHistory.end(), nextHistory.begin()))
    {
        return;
    }

    _pathHistory = std::move(nextHistory);

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

    StopDriveInfoWorker();

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
    ApplyDxEditHostThemes();
}

void NavigationView::SetPaneFocused(bool focused) noexcept
{
    if (_paneFocused == focused)
    {
        return;
    }

    _paneFocused = focused;
    UpdateEffectiveTheme();
    ApplyDxEditHostThemes();

    if (_d2dContext)
    {
        EnsureD2DResources();
    }

    RefreshActiveEditHostAfterParentPaint();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::SetEmbeddedDestinationMode(bool embedded) noexcept
{
    if (_embeddedDestinationMode == embedded)
    {
        return;
    }

    _embeddedDestinationMode = embedded;
    UpdateEffectiveTheme();
    ApplyDxEditHostThemes();

    if (_d2dContext)
    {
        EnsureD2DResources();
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void NavigationView::UpdatePathEditHostLayout() noexcept
{
    if (! _pathEdit || ! _pathEdit->hwnd)
    {
        return;
    }

    const RECT editBounds     = GetPathEditBoundsRect(_sectionPathRect, _sectionHistoryRect);
    const auto chrome         = ComputeEditChromeRects(editBounds, _dpi);
    const int hostWidth       = static_cast<int>((std::max)(0L, chrome.editRect.right - chrome.editRect.left));
    const int hostHeight      = static_cast<int>((std::max)(0L, chrome.editRect.bottom - chrome.editRect.top));
    const UINT visibilityFlag = _editMode ? static_cast<UINT>(SWP_SHOWWINDOW) : static_cast<UINT>(SWP_HIDEWINDOW);
    if (_editMode && ! _embeddedDestinationMode)
    {
        _pathEditBlurSuppressActive      = true;
        _pathEditBlurSuppressUntilTickMs = GetTickCount64() + 2000u;
    }
    SetWindowPos(
        _pathEdit->hwnd.get(), nullptr, chrome.editRect.left, chrome.editRect.top, hostWidth, hostHeight, SWP_NOZORDER | SWP_NOACTIVATE | visibilityFlag);

    if (_editMode)
    {
        RefreshActiveEditHostAfterParentPaint();
    }

    if (_pathEdit->field && ! _pathEdit->field->GetAccessibleHelpText().empty())
    {
        UpdateEditValidationPopupWindow(*_pathEdit);
    }
}

void NavigationView::UpdateFullPathPopupEditHostLayout() noexcept
{
    if (! _fullPathPopupEdit || ! _fullPathPopupEdit->hwnd || ! _fullPathPopup)
    {
        return;
    }

    RECT rc{};
    GetClientRect(_fullPathPopup.get(), &rc);
    const int hostWidth       = static_cast<int>((std::max)(0L, rc.right - rc.left));
    const int hostHeight      = static_cast<int>((std::max)(0L, rc.bottom - rc.top));
    const UINT visibilityFlag = _fullPathPopupEditMode ? static_cast<UINT>(SWP_SHOWWINDOW) : static_cast<UINT>(SWP_HIDEWINDOW);
    SetWindowPos(_fullPathPopupEdit->hwnd.get(), nullptr, rc.left, rc.top, hostWidth, hostHeight, SWP_NOZORDER | SWP_NOACTIVATE | visibilityFlag);

    if (_fullPathPopupEdit->field && ! _fullPathPopupEdit->field->GetAccessibleHelpText().empty())
    {
        UpdateEditValidationPopupWindow(*_fullPathPopupEdit);
    }
}

void NavigationView::ApplyDxEditHostThemes() noexcept
{
    const auto palette = MakeNavigationDxEditPalette(_appTheme, _theme);
    if (_pathEdit)
    {
        _pathEdit->host.SetTheme(palette);
        if (_pathEdit->field)
        {
            _pathEdit->field->SetCaretColor(_theme.text);
        }
    }
    if (_fullPathPopupEdit)
    {
        _fullPathPopupEdit->host.SetTheme(palette);
        if (_fullPathPopupEdit->field)
        {
            _fullPathPopupEdit->field->SetCaretColor(_theme.text);
        }
    }
}

#ifdef ENABLE_TESTS
bool NavigationView::DebugGetSnapshot(NavigationViewDebugSnapshot& out) const noexcept
{
    out = {};
    if (! _hWnd || IsWindow(_hWnd.get()) == FALSE)
    {
        return false;
    }

    out.dpi                               = _dpi;
    out.editMode                          = _editMode;
    out.embeddedDestinationMode           = _embeddedDestinationMode;
    out.fullPathPopupVisible              = _fullPathPopup && IsWindowVisible(_fullPathPopup.get()) != FALSE;
    out.fullPathPopupEditMode             = _fullPathPopupEditMode;
    out.showMenuSection                   = _showMenuSection;
    out.showDiskInfoSection               = _showDiskInfoSection;
    out.menuIconBitmapLoaded              = _menuIconBitmapD2D != nullptr;
    out.historyCount                      = _pathHistory.size();
    out.historyDropdownOpenCount          = _debugHistoryDropdownOpenCount;
    out.menuRegionRect                    = _sectionDriveRect;
    out.pathRegionRect                    = _sectionPathRect;
    out.historyRegionRect                 = _sectionHistoryRect;
    out.diskInfoRegionRect                = _sectionDiskInfoRect;
    out.debugEnterEditAttemptCount        = _debugEnterEditAttemptCount;
    out.debugEnterEditSuccessCount        = _debugEnterEditSuccessCount;
    out.debugEnterEditAbortCount          = _debugEnterEditAbortCount;
    out.debugExitEditCount                = _debugExitEditCount;
    out.debugDoubleClickActivateCount     = _debugDoubleClickActivateCount;
    out.debugKeyboardActivateCount        = _debugKeyboardActivateCount;
    out.debugLastDoubleClickOnLastSegment = _debugLastDoubleClickOnLastSegment;
    out.debugLastDoubleClickInWhitespace  = _debugLastDoubleClickInWhitespace;
    out.debugLastDoubleClickPoint         = _debugLastDoubleClickPoint;
    out.debugLastDoubleClickLocalX        = _debugLastDoubleClickLocalX;
    out.debugLastDoubleClickLocalY        = _debugLastDoubleClickLocalY;
    out.debugLastExitEditAccepted         = _debugLastExitEditAccepted;
    out.debugLastEnterEditAbortReason     = _debugLastEnterEditAbortReason;
    out.debugLastExitEditReason           = _debugLastExitEditReason;
    if (! _segments.empty() && ! _segments.back().isEllipsis)
    {
        const auto& lastSegment        = _segments.back();
        out.pathLastSegmentVisible     = true;
        out.pathLastSegmentRect.left   = _sectionPathRect.left + static_cast<LONG>(std::floor(lastSegment.bounds.left));
        out.pathLastSegmentRect.top    = _sectionPathRect.top + static_cast<LONG>(std::floor(lastSegment.bounds.top));
        out.pathLastSegmentRect.right  = _sectionPathRect.left + static_cast<LONG>(std::ceil(lastSegment.bounds.right));
        out.pathLastSegmentRect.bottom = _sectionPathRect.top + static_cast<LONG>(std::ceil(lastSegment.bounds.bottom));
    }
    for (const auto& segment : _segments)
    {
        if (! segment.isEllipsis)
        {
            if (_currentPath.has_value() && segment.fullPath == _currentPath.value())
            {
                out.pathCurrentSegmentVisible     = true;
                out.pathCurrentSegmentRect.left   = _sectionPathRect.left + static_cast<LONG>(std::floor(segment.bounds.left));
                out.pathCurrentSegmentRect.top    = _sectionPathRect.top + static_cast<LONG>(std::floor(segment.bounds.top));
                out.pathCurrentSegmentRect.right  = _sectionPathRect.left + static_cast<LONG>(std::ceil(segment.bounds.right));
                out.pathCurrentSegmentRect.bottom = _sectionPathRect.top + static_cast<LONG>(std::ceil(segment.bounds.bottom));
            }
            else if (_currentPath.has_value())
            {
                out.pathAncestorSegmentVisible     = true;
                out.pathAncestorSegmentRect.left   = _sectionPathRect.left + static_cast<LONG>(std::floor(segment.bounds.left));
                out.pathAncestorSegmentRect.top    = _sectionPathRect.top + static_cast<LONG>(std::floor(segment.bounds.top));
                out.pathAncestorSegmentRect.right  = _sectionPathRect.left + static_cast<LONG>(std::ceil(segment.bounds.right));
                out.pathAncestorSegmentRect.bottom = _sectionPathRect.top + static_cast<LONG>(std::ceil(segment.bounds.bottom));
                out.pathAncestorTargetText         = segment.fullPath.wstring();
            }

            continue;
        }

        out.pathEllipsisVisible     = true;
        out.pathEllipsisRect.left   = _sectionPathRect.left + static_cast<LONG>(std::floor(segment.bounds.left));
        out.pathEllipsisRect.top    = _sectionPathRect.top + static_cast<LONG>(std::floor(segment.bounds.top));
        out.pathEllipsisRect.right  = _sectionPathRect.left + static_cast<LONG>(std::ceil(segment.bounds.right));
        out.pathEllipsisRect.bottom = _sectionPathRect.top + static_cast<LONG>(std::ceil(segment.bounds.bottom));
        break;
    }
    out.menuButtonHovered     = _menuButtonHovered;
    out.historyButtonHovered  = _historyButtonHovered;
    out.diskInfoHovered       = _diskInfoHovered;
    out.hoveredSegmentIndex   = _hoveredSegmentIndex;
    out.hoveredSeparatorIndex = _hoveredSeparatorIndex;
    HWND navDropdownPopup     = nullptr;
    if (_navDropdownKind != ModernDropdownKind::None)
    {
        const HWND popup = FindOwnedVisibleDxContextMenuWindow(_hWnd.get());
        if (popup && IsWindowVisible(popup) != FALSE)
        {
            navDropdownPopup           = popup;
            out.historyDropdownVisible = true;
            switch (_navDropdownKind)
            {
                case ModernDropdownKind::Menu: out.dropdownKind = NavigationViewDebugDropdownKind::Menu; break;
                case ModernDropdownKind::Drive: out.dropdownKind = NavigationViewDebugDropdownKind::Drive; break;
                case ModernDropdownKind::History: out.dropdownKind = NavigationViewDebugDropdownKind::History; break;
                case ModernDropdownKind::DiskInfo: out.dropdownKind = NavigationViewDebugDropdownKind::DiskInfo; break;
                case ModernDropdownKind::Siblings: out.dropdownKind = NavigationViewDebugDropdownKind::Siblings; break;
                case ModernDropdownKind::None: break;
            }
            out.historyDropdownItemCount     = _navDropdownPaths.size();
            out.historyDropdownSelectedIndex = _navDropdownSelectedIndex;
        }
    }
    if (_editSuggestPopup && IsWindowVisible(_editSuggestPopup.get()) != FALSE)
    {
        out.editSuggestPopupVisible    = true;
        out.editSuggestItemCount       = _editSuggestItems.size();
        out.editSuggestSelectedIndex   = _editSuggestSelectedIndex;
        out.editSuggestPopupClientSize = _editSuggestPopupClientSize;
    }
    if (_editValidationPopup && IsWindowVisible(_editValidationPopup.get()) != FALSE)
    {
        out.currentEditValidationPopupVisible        = true;
        out.currentEditValidationPopupHwnd           = _editValidationPopup.get();
        out.currentEditValidationPopupRoundedRegion  = _editValidationPopupRoundedRegion;
        out.currentEditValidationPopupUsesFluentIcon = _editValidationPopupIconUsesFluent;
        out.currentEditValidationPopupIconGlyph      = _editValidationPopupIconGlyph;
        if (GetWindowRect(_editValidationPopup.get(), &out.currentEditValidationPopupScreenRect) == FALSE)
        {
            out.currentEditValidationPopupScreenRect = _editValidationPopupScreenRect;
        }
    }
    if (_currentPath.has_value())
    {
        out.currentPathText = _currentPath->wstring();
    }

    const auto copyWindowText = [](const HWND hwnd, std::wstring& text) noexcept
    {
        text.clear();
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            return;
        }

        const int length = GetWindowTextLengthW(hwnd);
        if (length <= 0)
        {
            return;
        }

        text.resize(static_cast<size_t>(length));
        const int copied = GetWindowTextW(hwnd, text.data(), length + 1);
        if (copied <= 0)
        {
            text.clear();
            return;
        }
        text.resize(static_cast<size_t>(copied));
    };

    const auto captureCurrentEditDiagnostics = [&out](const NavigationDxTextHost& textHost) noexcept
    {
        out.currentEditHostHwnd            = textHost.hwnd.get();
        out.currentEditInputHwnd           = textHost.GetTextInputHwnd();
        out.currentEditUsesNativeTextInput = textHost.host.GetTextInputBackend() == RedSalamander::DxUi::TextInputBackend::Native;
        if (textHost.field)
        {
            out.currentEditHelpText.assign(textHost.field->GetAccessibleHelpText());
        }

        D2D1_RECT_F caretRectDip{};
        RECT caretScreenRectPx{};
        if (textHost.host.DebugGetNativeTextInputCaretRect(caretRectDip, caretScreenRectPx))
        {
            out.currentEditCaretScreenRectValid = true;
            out.currentEditCaretScreenRect      = caretScreenRectPx;
        }

        RedSalamander::DxUi::NativeTextInputState nativeState{};
        if (textHost.host.DebugGetNativeTextInputState(nativeState) && nativeState.compositionStartIndex.has_value() &&
            nativeState.compositionEndIndex.has_value())
        {
            out.currentEditHasActiveComposition = true;
            out.currentEditCompositionStart     = nativeState.compositionStartIndex.value();
            out.currentEditCompositionEnd       = nativeState.compositionEndIndex.value();
        }
    };

    if (_pathEdit && _pathEdit->field && _pathEdit->hwnd && IsWindowVisible(_pathEdit->hwnd.get()) != FALSE)
    {
        out.currentEditText.assign(_pathEdit->field->GetText());
        captureCurrentEditDiagnostics(*_pathEdit);
        const size_t caretIndex       = std::min(_pathEdit->field->GetCaretIndex(), out.currentEditText.size());
        out.currentEditSelectionStart = caretIndex;
        out.currentEditSelectionEnd   = caretIndex;
        if (const auto selection = _pathEdit->field->GetSelectionRange(); selection.has_value())
        {
            out.currentEditHasSelection   = true;
            out.currentEditSelectionStart = selection->first;
            out.currentEditSelectionEnd   = selection->second;
        }
    }

    if (_fullPathPopup && IsWindowVisible(_fullPathPopup.get()) != FALSE)
    {
        out.fullPathPopupClientSize = _fullPathPopupClientSize;
        for (const auto& segment : _fullPathPopupSegments)
        {
            if (_currentPath.has_value() && segment.fullPath != _currentPath.value())
            {
                out.fullPathPopupAncestorSegmentVisible     = true;
                out.fullPathPopupAncestorSegmentRect.left   = static_cast<LONG>(std::floor(segment.bounds.left));
                out.fullPathPopupAncestorSegmentRect.top    = static_cast<LONG>(std::floor(segment.bounds.top - _fullPathPopupScrollY));
                out.fullPathPopupAncestorSegmentRect.right  = static_cast<LONG>(std::ceil(segment.bounds.right));
                out.fullPathPopupAncestorSegmentRect.bottom = static_cast<LONG>(std::ceil(segment.bounds.bottom - _fullPathPopupScrollY));
                out.fullPathPopupAncestorTargetText         = segment.fullPath.wstring();
            }
        }
    }

    if (_fullPathPopupEdit && _fullPathPopupEdit->field && _fullPathPopupEdit->hwnd && IsWindowVisible(_fullPathPopupEdit->hwnd.get()) != FALSE)
    {
        out.currentEditText.assign(_fullPathPopupEdit->field->GetText());
        captureCurrentEditDiagnostics(*_fullPathPopupEdit);
        const size_t caretIndex       = std::min(_fullPathPopupEdit->field->GetCaretIndex(), out.currentEditText.size());
        out.currentEditSelectionStart = caretIndex;
        out.currentEditSelectionEnd   = caretIndex;
        if (const auto selection = _fullPathPopupEdit->field->GetSelectionRange(); selection.has_value())
        {
            out.currentEditHasSelection   = true;
            out.currentEditSelectionStart = selection->first;
            out.currentEditSelectionEnd   = selection->second;
        }
    }
    else if ((! _pathEdit || ! _pathEdit->field || ! _pathEdit->hwnd || IsWindowVisible(_pathEdit->hwnd.get()) == FALSE) && _currentEditPath.has_value())
    {
        out.currentEditText = _currentEditPath->wstring();
    }

    out.visibleChildWindowCount = 0u;
    if (_pathEdit && _pathEdit->hwnd && IsWindowVisible(_pathEdit->hwnd.get()) != FALSE)
    {
        ++out.visibleChildWindowCount;
    }
    if (out.historyDropdownVisible)
    {
        ++out.visibleChildWindowCount;
    }

    const HWND focused = GetFocus();
    if (_pathEdit && _pathEdit->hwnd && focused &&
        (focused == _pathEdit->hwnd.get() || focused == _pathEdit->GetTextInputHwnd() || IsChild(_pathEdit->hwnd.get(), focused) != FALSE))
    {
        out.focusTarget = NavigationViewDebugFocusTarget::PathEdit;
        return true;
    }

    if (navDropdownPopup && focused && (focused == navDropdownPopup || IsChild(navDropdownPopup, focused) != FALSE))
    {
        out.focusTarget = NavigationViewDebugFocusTarget::HistoryDropdown;
        return true;
    }

    if (_fullPathPopupEdit && _fullPathPopupEdit->hwnd && focused &&
        (focused == _fullPathPopupEdit->hwnd.get() || focused == _fullPathPopupEdit->GetTextInputHwnd() ||
         IsChild(_fullPathPopupEdit->hwnd.get(), focused) != FALSE))
    {
        out.focusTarget = NavigationViewDebugFocusTarget::FullPathPopupEdit;
        return true;
    }

    if (focused && (focused == _hWnd.get() || IsChild(_hWnd.get(), focused) != FALSE))
    {
        switch (_focusedRegion)
        {
            case FocusRegion::Menu: out.focusTarget = NavigationViewDebugFocusTarget::MenuRegion; break;
            case FocusRegion::Path: out.focusTarget = NavigationViewDebugFocusTarget::PathRegion; break;
            case FocusRegion::History: out.focusTarget = NavigationViewDebugFocusTarget::HistoryRegion; break;
            case FocusRegion::DiskInfo: out.focusTarget = NavigationViewDebugFocusTarget::DiskInfoRegion; break;
        }
    }

    return true;
}

bool NavigationView::DebugFocusRegion(FocusRegion region) noexcept
{
    if (! _hWnd || IsWindow(_hWnd.get()) == FALSE)
    {
        return false;
    }

    SetFocusRegion(region);
    SetFocus(_hWnd.get());
    return GetFocus() == _hWnd.get();
}

#endif

void NavigationView::UpdateEffectiveTheme() noexcept
{
    _theme = _baseTheme;

    if (_embeddedDestinationMode)
    {
        const D2D1::ColorF background = ColorFromCOLORREF(_appTheme.windowBackground);
        const float hoverBlend        = _theme.darkBase ? 0.10f : 0.06f;
        const float pressedBlend      = _theme.darkBase ? 0.16f : 0.10f;
        const float textBlend         = _theme.darkBase ? 0.45f : 0.35f;
        const float sepBlend          = _theme.darkBase ? 0.65f : 0.55f;
        const float accentBlend       = _theme.darkBase ? 0.50f : 0.40f;

        _theme.background         = background;
        _theme.backgroundHover    = BlendColorF(background, _theme.text, hoverBlend);
        _theme.backgroundPressed  = BlendColorF(background, _theme.text, pressedBlend);
        _theme.hoverHighlight     = _theme.backgroundHover;
        _theme.pressedHighlight   = _theme.backgroundPressed;
        _theme.text               = BlendColorF(_theme.text, _theme.background, textBlend);
        _theme.separator          = BlendColorF(_theme.separator, _theme.background, sepBlend);
        _theme.accent             = BlendColorF(_theme.accent, _theme.background, accentBlend);
        _theme.progressOk         = BlendColorF(_theme.progressOk, _theme.background, accentBlend);
        _theme.progressWarn       = BlendColorF(_theme.progressWarn, _theme.background, accentBlend);
        _theme.progressBackground = BlendColorF(_theme.progressBackground, _theme.background, std::max(accentBlend, 0.65f));
        _theme.gdiBorderPen       = ColorToCOLORREF(background);
    }
    else if (_paneFocused)
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
        PostMessageW(_hWnd.get(), WndMsg::kNavigationViewShowHistoryDropdown, 1, 0);
    }
}

// Direct2D implementation continues in next part...
