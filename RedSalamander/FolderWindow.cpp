#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <winsock2.h>

#include "FolderWindowInternal.h"
#include "HostServices.h"
#include "NavigationLocation.h"

#include <netioapi.h>

#pragma comment(lib, "Iphlpapi.lib")

namespace
{
LRESULT OnHostServicesMessage(UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    LRESULT result = 0;
    if (TryHandleHostServicesWindowMessage(msg, wp, lp, result))
    {
        return result;
    }
    return 0;
}

[[nodiscard]] COLORREF SplitterGripColor(const AppTheme& theme) noexcept
{
    if (theme.highContrast)
    {
        return theme.menu.text;
    }

    constexpr int kTowardTextWeight = 1;
    constexpr int kDenom            = 4;
    static_assert(kTowardTextWeight > 0 && kTowardTextWeight < kDenom);

    const int baseWeight           = kDenom - kTowardTextWeight;
    const COLORREF baseColor       = theme.menu.separator;
    const COLORREF towardTextColor = theme.menu.text;

    const int r = (static_cast<int>(GetRValue(baseColor)) * baseWeight + static_cast<int>(GetRValue(towardTextColor)) * kTowardTextWeight) / kDenom;
    const int g = (static_cast<int>(GetGValue(baseColor)) * baseWeight + static_cast<int>(GetGValue(towardTextColor)) * kTowardTextWeight) / kDenom;
    const int b = (static_cast<int>(GetBValue(baseColor)) * baseWeight + static_cast<int>(GetBValue(towardTextColor)) * kTowardTextWeight) / kDenom;

    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}
} // namespace

class FolderWindow::NetworkChangeSubscription final
{
public:
    explicit NetworkChangeSubscription(HWND hwnd) noexcept : _hwnd(hwnd)
    {
        if (! _hwnd)
        {
            return;
        }

        HANDLE handle                  = nullptr;
        const BOOL initialNotification = FALSE;

#pragma warning(push)
        // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
#pragma warning(disable : 5039)
        const NTSTATUS status = NotifyIpInterfaceChange(AF_UNSPEC, &NetworkChangeSubscription::OnIpInterfaceChanged, this, initialNotification, &handle);
#pragma warning(pop)

        if (status != NO_ERROR)
        {
            Debug::Warning(L"FolderWindow: NotifyIpInterfaceChange failed (status={})", status);
            return;
        }

        _handle.reset(handle);
    }

    NetworkChangeSubscription(const NetworkChangeSubscription&)            = delete;
    NetworkChangeSubscription& operator=(const NetworkChangeSubscription&) = delete;
    NetworkChangeSubscription(NetworkChangeSubscription&&)                 = delete;
    NetworkChangeSubscription& operator=(NetworkChangeSubscription&&)      = delete;

    ~NetworkChangeSubscription() = default;

private:
    static void CALLBACK OnIpInterfaceChanged(PVOID callerContext, PMIB_IPINTERFACE_ROW /*row*/, MIB_NOTIFICATION_TYPE notificationType) noexcept
    {
        if (notificationType == MibInitialNotification)
        {
            return;
        }

        auto* self = static_cast<NetworkChangeSubscription*>(callerContext);
        if (! self || ! self->_hwnd)
        {
            return;
        }

        PostMessageW(self->_hwnd, WndMsg::kNetworkConnectivityChanged, 0, 0);
    }

    struct NetworkHandleDeleter
    {
        void operator()(HANDLE handle) const noexcept
        {
            if (handle)
            {
                CancelMibChangeNotify2(handle);
            }
        }
    };

    HWND _hwnd = nullptr;
    std::unique_ptr<void, NetworkHandleDeleter> _handle;
};

FolderWindow::FolderWindow()
{
    _viewerCallback.owner = this;
}

FolderWindow::~FolderWindow()
{
    Destroy();
}

void FolderWindow::SetSettings(Common::Settings::Settings* settings) noexcept
{
    _settings = settings;
    _leftPane.navigationView.SetSettings(settings);
    _rightPane.navigationView.SetSettings(settings);
}

void FolderWindow::SetShortcutManager(const ShortcutManager* shortcuts) noexcept
{
    _shortcutManager = shortcuts;
    _functionBar.SetShortcutManager(shortcuts);
    _leftPane.folderView.SetShortcutManager(shortcuts);
    _rightPane.folderView.SetShortcutManager(shortcuts);
}

void FolderWindow::SetFunctionBarModifiers(uint32_t modifiers) noexcept
{
    _functionBar.SetModifiers(modifiers);
}

void FolderWindow::SetFunctionBarPressedKey(std::optional<uint32_t> vk) noexcept
{
    _functionBar.SetPressedFunctionKey(vk);
}

void FolderWindow::SetFunctionBarVisible(bool visible) noexcept
{
    if (_functionBarVisible == visible)
    {
        return;
    }

    _functionBarVisible = visible;

    if (_hWnd)
    {
        CalculateLayout();
        AdjustChildWindows();
    }

    if (const HWND bar = _functionBar.GetHwnd())
    {
        ShowWindow(bar, _functionBarVisible ? SW_SHOW : SW_HIDE);
    }
}

void FolderWindow::SetPanePathChangedCallback(PanePathChangedCallback callback)
{
    _panePathChangedCallback = std::move(callback);
}

void FolderWindow::SetPaneEnumerationCompletedCallback(Pane pane, FolderView::EnumerationCompletedCallback callback)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetEnumerationCompletedCallback(std::move(callback));
}

void FolderWindow::SetPaneDetailsTextProvider(Pane pane, FolderView::DetailsTextProvider provider)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetDetailsTextProvider(std::move(provider));
}

void FolderWindow::RefreshPaneDetailsText(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.RefreshDetailsText();
}

void FolderWindow::SetPaneSelectionByDisplayNamePredicate(Pane pane,
                                                          const std::function<bool(std::wstring_view)>& shouldSelect,
                                                          bool clearExistingSelection) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.SetSelectionByDisplayNamePredicate(shouldSelect, clearExistingSelection);
}

void FolderWindow::SetFileOperationCompletedCallback(FileOperationCompletedCallback callback)
{
    _fileOperationCompletedCallback = std::move(callback);
}

ATOM FolderWindow::RegisterWndClass(HINSTANCE instance)
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

HWND FolderWindow::Create(HWND parent, int x, int y, int width, int height)
{
    _hInstance = GetModuleHandle(nullptr);

    if (! RegisterWndClass(_hInstance))
    {
        return nullptr;
    }

    _clientSize = {static_cast<LONG>(width), static_cast<LONG>(height)};

    const HWND hwnd =
        CreateWindowExW(0, kClassName, L"", WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, x, y, width, height, parent, nullptr, _hInstance, this);
    if (! hwnd)
    {
        return nullptr;
    }

    _backgroundBrush.reset(CreateSolidBrush(_theme.windowBackground));
    _splitterBrush.reset(CreateSolidBrush(_theme.menu.separator));
    _splitterGripBrush.reset(CreateSolidBrush(SplitterGripColor(_theme)));

    return hwnd;
}

void FolderWindow::Destroy()
{
    ShutdownFileOperations();

    CancelSelectionSizeComputation(Pane::Left);
    CancelSelectionSizeComputation(Pane::Right);

    if (_leftPane.selectionSizeThread.joinable())
    {
        _leftPane.selectionSizeThread.request_stop();
        _leftPane.selectionSizeCv.notify_all();
        _leftPane.selectionSizeThread = std::jthread{};
    }

    if (_rightPane.selectionSizeThread.joinable())
    {
        _rightPane.selectionSizeThread.request_stop();
        _rightPane.selectionSizeCv.notify_all();
        _rightPane.selectionSizeThread = std::jthread{};
    }

    _backgroundBrush.reset();
    _splitterBrush.reset();
    _splitterGripBrush.reset();

    _functionBar.Destroy();

    if (_leftPane.hNavigationView)
    {
        _leftPane.navigationView.Destroy();
        _leftPane.hNavigationView.reset();
    }

    if (_leftPane.hFolderView)
    {
        _leftPane.folderView.Destroy();
        _leftPane.hFolderView.reset();
    }

    if (_leftPane.hStatusBar)
    {
        _leftPane.hStatusBar.reset();
    }

    if (_rightPane.hNavigationView)
    {
        _rightPane.navigationView.Destroy();
        _rightPane.hNavigationView.reset();
    }

    if (_rightPane.hFolderView)
    {
        _rightPane.folderView.Destroy();
        _rightPane.hFolderView.reset();
    }

    if (_rightPane.hStatusBar)
    {
        _rightPane.hStatusBar.reset();
    }

    if (_leftPane.fileSystem)
    {
        DirectoryInfoCache::GetInstance().ClearForFileSystem(_leftPane.fileSystem.get());
    }
    if (_rightPane.fileSystem)
    {
        DirectoryInfoCache::GetInstance().ClearForFileSystem(_rightPane.fileSystem.get());
    }

    _leftPane.fileSystem = nullptr;
    _leftPane.fileSystemModule.reset();
    _leftPane.pluginId.clear();
    _leftPane.currentPath.reset();
    _leftPane.updatingPath = false;

    _rightPane.fileSystem = nullptr;
    _rightPane.fileSystemModule.reset();
    _rightPane.pluginId.clear();
    _rightPane.currentPath.reset();
    _rightPane.updatingPath = false;

    _hWnd.reset();
}

LRESULT CALLBACK FolderWindow::WndProcThunk(HWND hWindow, UINT msg, WPARAM wp, LPARAM lp)
{
    FolderWindow* self = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self    = reinterpret_cast<FolderWindow*>(cs->lpCreateParams);
        SetWindowLongPtrW(hWindow, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        self->_hWnd.reset(hWindow);
        InitPostedPayloadWindow(hWindow);
    }
    else
    {
        self = reinterpret_cast<FolderWindow*>(GetWindowLongPtrW(hWindow, GWLP_USERDATA));
    }

    if (self)
    {
        return self->WndProc(hWindow, msg, wp, lp);
    }

    return DefWindowProcW(hWindow, msg, wp, lp);
}

LRESULT FolderWindow::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_CREATE: return OnCreate(hwnd) ? 0 : -1;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_NCDESTROY: static_cast<void>(DrainPostedPayloadsForWindow(hwnd)); break;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_SETFOCUS: OnSetFocus(); return 0;
        case WM_DEVICECHANGE: return OnDeviceChange(static_cast<UINT>(wp), lp);
        case WndMsg::kNetworkConnectivityChanged: OnNetworkConnectivityChanged(); return 0;
        case WM_ERASEBKGND: return 1; // no erase background
        case WM_PAINT: OnPaint(); return 0;
        case WM_DRAWITEM: return OnDrawItem(reinterpret_cast<DRAWITEMSTRUCT*>(lp));
        case WM_LBUTTONDOWN: OnLButtonDown({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_LBUTTONDBLCLK: OnLButtonDblClk({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_LBUTTONUP: OnLButtonUp(); return 0;
        case WM_MOUSEMOVE: OnMouseMove({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_CAPTURECHANGED: OnCaptureChanged(); return 0;
        case WM_PARENTNOTIFY: OnParentNotify(LOWORD(wp), HIWORD(wp)); return 0;
        case WM_NOTIFY: return OnNotify(reinterpret_cast<const NMHDR*>(lp));
        case WM_SETCURSOR: return OnSetCursor(reinterpret_cast<HWND>(wp), LOWORD(lp), HIWORD(lp));
        case WndMsg::kPaneFocusChanged: UpdatePaneFocusStates(); return 0;
        case WndMsg::kPaneSelectionSizeComputed: return OnPaneSelectionSizeComputed(lp);
        case WndMsg::kPaneSelectionSizeProgress: return OnPaneSelectionSizeProgress(lp);
        case WndMsg::kFileOperationCompleted: return OnFileOperationCompleted(lp);
        case WndMsg::kHostShowAlert: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostClearAlert: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostShowPrompt: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostShowConnectionManager: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostGetConnectionJsonUtf8: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostGetConnectionSecret: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostPromptConnectionSecret: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostClearCachedConnectionSecret: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostUpgradeFtpAnonymousToPassword: return OnHostServicesMessage(msg, wp, lp);
        case WndMsg::kHostExecuteInPane: return OnHostServicesMessage(msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT FolderWindow::OnDeviceChange(UINT event, LPARAM data) noexcept
{
    if (event != DBT_DEVICEARRIVAL && event != DBT_DEVICEREMOVECOMPLETE)
    {
        return static_cast<LRESULT>(TRUE);
    }

    const auto* hdr = reinterpret_cast<const DEV_BROADCAST_HDR*>(data);
    if (! hdr || hdr->dbch_devicetype != DBT_DEVTYP_VOLUME)
    {
        return static_cast<LRESULT>(TRUE);
    }

    const auto* volume   = reinterpret_cast<const DEV_BROADCAST_VOLUME*>(hdr);
    const DWORD unitmask = volume->dbcv_unitmask;
    if (unitmask == 0)
    {
        return static_cast<LRESULT>(TRUE);
    }

    auto refreshIfAffected = [&](PaneState& pane)
    {
        const std::optional<std::filesystem::path>& current = pane.currentPath;
        if (! current.has_value())
        {
            return;
        }

        const auto driveLetter = NavigationLocation::TryGetWindowsDriveLetter(current.value());
        if (! driveLetter.has_value() || ! NavigationLocation::DriveMaskContainsLetter(static_cast<uint32_t>(unitmask), driveLetter.value()))
        {
            return;
        }

        pane.folderView.ForceRefresh();
    };

    refreshIfAffected(_leftPane);
    refreshIfAffected(_rightPane);
    return static_cast<LRESULT>(TRUE);
}

void FolderWindow::OnNetworkConnectivityChanged() noexcept
{
    const uint64_t now             = GetTickCount64();
    constexpr uint64_t kDebounceMs = 500;
    if (_lastNetworkConnectivityRefreshTick != 0 && now - _lastNetworkConnectivityRefreshTick < kDebounceMs)
    {
        return;
    }
    _lastNetworkConnectivityRefreshTick = now;

    auto refreshIfNetworkPath = [&](PaneState& pane)
    {
        if (! pane.hFolderView)
        {
            return;
        }

        if (! NavigationLocation::IsFilePluginShortId(pane.pluginShortId))
        {
            return;
        }

        if (! pane.currentPath.has_value())
        {
            return;
        }

        const std::wstring_view pathText = pane.currentPath.value().native();
        if (NavigationLocation::LooksLikeUncPath(pathText))
        {
            pane.folderView.ForceRefresh();
            return;
        }

        const auto driveLetter = NavigationLocation::TryGetWindowsDriveLetter(pathText);
        if (! driveLetter.has_value())
        {
            return;
        }

        std::wstring driveRoot;
        driveRoot.push_back(driveLetter.value());
        driveRoot.append(L":\\");

        const UINT driveType = GetDriveTypeW(driveRoot.c_str());
        if (driveType == DRIVE_REMOTE)
        {
            pane.folderView.ForceRefresh();
        }
    };

    refreshIfNetworkPath(_leftPane);
    refreshIfNetworkPath(_rightPane);
}

LRESULT FolderWindow::OnDrawItem(DRAWITEMSTRUCT* dis)
{
    if (! _hWnd)
    {
        return 0;
    }

    const WPARAM controlId = dis ? static_cast<WPARAM>(dis->CtlID) : 0;
    return DefWindowProcW(_hWnd.get(), WM_DRAWITEM, controlId, reinterpret_cast<LPARAM>(dis));
}

LRESULT FolderWindow::OnNotify(const NMHDR* header)
{
    if (header && header->code == NM_CLICK && (header->idFrom == kLeftStatusBarId || header->idFrom == kRightStatusBarId))
    {
        const auto* mouse = reinterpret_cast<const NMMOUSE*>(header);
        const int part    = static_cast<int>(mouse->dwItemSpec);
        if (part == 1 && _showSortMenuCallback)
        {
            const Pane pane = header->idFrom == kLeftStatusBarId ? Pane::Left : Pane::Right;
            SetActivePane(pane);

            RECT partRect{};
            POINT screenPoint{};
            if (SendMessageW(header->hwndFrom, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&partRect)) != 0)
            {
                const int dpi      = static_cast<int>(GetDpiForWindow(header->hwndFrom));
                const int paddingX = MulDiv(kStatusBarSortPaddingXDip, dpi, USER_DEFAULT_SCREEN_DPI);
                screenPoint        = {std::max(partRect.left, partRect.right - paddingX), partRect.top};
            }
            else
            {
                screenPoint = mouse->pt;
            }
            ClientToScreen(header->hwndFrom, &screenPoint);
            _showSortMenuCallback(pane, screenPoint);
            return 0;
        }
    }

    if (! _hWnd)
    {
        return 0;
    }
    const WPARAM controlId = header ? static_cast<WPARAM>(header->idFrom) : 0;
    return DefWindowProcW(_hWnd.get(), WM_NOTIFY, controlId, reinterpret_cast<LPARAM>(header));
}

bool FolderWindow::OnCreate(HWND hwnd) noexcept
{
    // _hWnd not yet initialize
    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.GetDpiForWindow");
        _dpi = GetDpiForWindow(hwnd);
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.EnsureFileOperations");
        EnsureFileOperations();
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.InitCommonControlsEx");
        INITCOMMONCONTROLSEX icc{};
        icc.dwSize = sizeof(icc);
        icc.dwICC  = ICC_BAR_CLASSES;
        InitCommonControlsEx(&icc);
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CalculateLayout");
        CalculateLayout();
    }

    auto createPane = [&](Pane pane,
                          PaneState& state,
                          const RECT& navRect,
                          const RECT& folderRect,
                          const RECT& statusRect,
                          UINT_PTR navId,
                          UINT_PTR folderId,
                          UINT_PTR statusId) -> bool
    {
        const std::wstring_view paneName = pane == Pane::Left ? L"Left" : L"Right";

        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.NavigationView.Create");
            perf.SetDetail(paneName);
            state.hNavigationView.reset(
                state.navigationView.Create(hwnd, navRect.left, navRect.top, navRect.right - navRect.left, navRect.bottom - navRect.top));
        }
        if (! state.hNavigationView)
        {
            return false;
        }
        SetWindowLongPtrW(state.hNavigationView.get(), GWLP_ID, static_cast<LONG_PTR>(navId));

        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.FolderView.Create");
            perf.SetDetail(paneName);
            state.hFolderView.reset(
                state.folderView.Create(hwnd, folderRect.left, folderRect.top, folderRect.right - folderRect.left, folderRect.bottom - folderRect.top));
        }
        if (! state.hFolderView)
        {
            return false;
        }
        SetWindowLongPtrW(state.hFolderView.get(), GWLP_ID, static_cast<LONG_PTR>(folderId));

        const int statusWidth   = std::max(0L, statusRect.right - statusRect.left);
        const int statusHeight  = std::max(0L, statusRect.bottom - statusRect.top);
        const DWORD statusStyle = WS_CHILD | WS_VISIBLE | CCS_NOPARENTALIGN | CCS_NORESIZE | SBARS_TOOLTIPS;
        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.StatusBar.CreateWindowExW");
            perf.SetDetail(paneName);
            state.hStatusBar.reset(CreateWindowExW(0,
                                                   STATUSCLASSNAMEW,
                                                   nullptr,
                                                   statusStyle,
                                                   statusRect.left,
                                                   statusRect.top,
                                                   statusWidth,
                                                   statusHeight,
                                                   hwnd,
                                                   reinterpret_cast<HMENU>(statusId),
                                                   _hInstance,
                                                   nullptr));
        }
        if (! state.hStatusBar)
        {
            return false;
        }
        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.StatusBar.Initialize");
            perf.SetDetail(paneName);
            SetPropW(state.hStatusBar.get(), kStatusBarOwnerProp, reinterpret_cast<HANDLE>(this));
            SetPropW(state.hStatusBar.get(), kStatusBarSelectionTextProp, reinterpret_cast<HANDLE>(&state.statusSelectionText));
            SetPropW(state.hStatusBar.get(), kStatusBarSortTextProp, reinterpret_cast<HANDLE>(&state.statusSortText));
            SetPropW(state.hStatusBar.get(), kStatusBarFocusHueProp, reinterpret_cast<HANDLE>(&state.statusFocusHueDegrees));
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
            SetWindowSubclass(state.hStatusBar.get(), StatusBarSubclassProc, statusId, 0);
#pragma warning(pop)
        }

        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.SetFileSystem");
            perf.SetDetail(paneName);
            state.folderView.SetFileSystem(state.fileSystem);
            state.navigationView.SetFileSystem(state.fileSystem);
        }

        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.SetCallbacks");
            perf.SetDetail(paneName);
            state.navigationView.SetPathChangedCallback([this, pane](const std::optional<std::filesystem::path>& path)
                                                        { OnNavigationPathChanged(pane, path); });
            state.navigationView.SetRequestFolderViewFocusCallback(
                [this, pane]
                {
                    PaneState& s = pane == Pane::Left ? _leftPane : _rightPane;
                    if (s.hFolderView)
                    {
                        SetFocus(s.hFolderView.get());
                    }
                });

            state.folderView.SetPathChangedCallback([this, pane](const std::optional<std::filesystem::path>& path) { OnFolderViewPathChanged(pane, path); });
            state.folderView.SetNavigateUpFromRootRequestCallback([this, pane] { OnFolderViewNavigateUpFromRoot(pane); });
            state.folderView.SetOpenFileRequestCallback([this, pane](const std::filesystem::path& path) { return TryOpenFileAsVirtualFileSystem(pane, path); });
            state.folderView.SetViewFileRequestCallback([this, pane](const FolderView::ViewFileRequest& request)
                                                        { return TryViewFileWithViewer(pane, request); });
            state.folderView.SetFileOperationRequestCallback([this, pane](FolderView::FileOperationRequest request) noexcept -> HRESULT
                                                             { return StartFileOperationFromFolderView(pane, std::move(request)); });
            state.folderView.SetPropertiesRequestCallback([this, pane](std::filesystem::path path) noexcept -> HRESULT
                                                          { return ShowItemPropertiesFromFolderView(pane, std::move(path)); });
            state.folderView.SetNavigationRequestCallback(
                [this, pane](FolderView::NavigationRequest request)
                {
                    PaneState& s = pane == Pane::Left ? _leftPane : _rightPane;
                    switch (request)
                    {
                        case FolderView::NavigationRequest::FocusNavigationMenu:
                            s.navigationView.SetFocusRegion(NavigationView::FocusRegion::Menu);
                            if (s.hNavigationView)
                            {
                                SetFocus(s.hNavigationView.get());
                            }
                            break;
                        case FolderView::NavigationRequest::FocusNavigationDiskInfo:
                            s.navigationView.SetFocusRegion(NavigationView::FocusRegion::DiskInfo);
                            if (s.hNavigationView)
                            {
                                SetFocus(s.hNavigationView.get());
                            }
                            break;
                        case FolderView::NavigationRequest::FocusAddressBar: s.navigationView.FocusAddressBar(); break;
                        case FolderView::NavigationRequest::OpenHistoryDropdown: s.navigationView.OpenHistoryDropdownFromKeyboard(); break;
                        case FolderView::NavigationRequest::SwitchPane:
                        {
                            const Pane otherPane = pane == Pane::Left ? Pane::Right : Pane::Left;
                            PaneState& other     = otherPane == Pane::Left ? _leftPane : _rightPane;
                            if (other.hFolderView)
                            {
                                SetActivePane(otherPane);
                                SetFocus(other.hFolderView.get());
                            }
                            break;
                        }
                    }
                });

            state.folderView.SetSelectionChangedCallback(
                [this, pane](const FolderView::SelectionStats& stats)
                {
                    PaneState& s     = pane == Pane::Left ? _leftPane : _rightPane;
                    s.selectionStats = stats;
                    CancelSelectionSizeComputation(pane);
                    UpdatePaneStatusBar(pane);
                });

            state.folderView.SetIncrementalSearchChangedCallback([this, pane] { UpdatePaneStatusBar(pane); });
            state.folderView.SetSelectionSizeComputationRequestedCallback([this, pane] { RequestSelectionSizeComputation(pane); });
        }

        {
            Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane.StartSelectionSizeWorker");
            perf.SetDetail(paneName);
            StartSelectionSizeWorker(pane);
        }

        return true;
    };

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane");
        perf.SetDetail(L"Left");
        if (! createPane(
                Pane::Left, _leftPane, _leftNavigationRect, _leftFolderViewRect, _leftStatusBarRect, kLeftNavigationId, kLeftFolderViewId, kLeftStatusBarId))
        {
            Debug::Error(L"FolderWindow::OnCreate failed to create left pane.");
            return false;
        }
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.CreatePane");
        perf.SetDetail(L"Right");
        if (! createPane(Pane::Right,
                         _rightPane,
                         _rightNavigationRect,
                         _rightFolderViewRect,
                         _rightStatusBarRect,
                         kRightNavigationId,
                         kRightFolderViewId,
                         kRightStatusBarId))
        {
            Debug::Error(L"FolderWindow::OnCreate failed to create right pane.");
            return false;
        }
    }

    const int functionBarWidth  = std::max(0L, _functionBarRect.right - _functionBarRect.left);
    const int functionBarHeight = std::max(0L, _functionBarRect.bottom - _functionBarRect.top);
    HWND functionBarHwnd        = nullptr;
    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.FunctionBar.Create");
        functionBarHwnd = _functionBar.Create(hwnd, _functionBarRect.left, _functionBarRect.top, functionBarWidth, functionBarHeight);
    }
    if (functionBarHwnd)
    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.FunctionBar.Initialize");
        _functionBar.SetDpi(_dpi);
        _functionBar.SetShortcutManager(_shortcutManager);
        _functionBar.SetTheme(_theme);
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.UpdatePaneUI");
        UpdatePaneStatusBar(Pane::Left);
        UpdatePaneStatusBar(Pane::Right);
        UpdatePaneFocusStates();
    }

    const std::wstring_view defaultPluginId = FileSystemPluginManager::GetInstance().GetActivePluginId();
    if (! defaultPluginId.empty())
    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.EnsurePaneFileSystems");
        perf.SetDetail(defaultPluginId);
        static_cast<void>(EnsurePaneFileSystem(Pane::Left, defaultPluginId));
        static_cast<void>(EnsurePaneFileSystem(Pane::Right, defaultPluginId));
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.NetworkChangeSubscription");
        _networkChangeSubscription = std::make_unique<NetworkChangeSubscription>(hwnd);
    }

    {
        Debug::Perf::Scope perf(L"FolderWindow.OnCreate.ApplyTheme");
        ApplyTheme(ResolveAppTheme(ThemeMode::System, L"RedSalamander"));
    }
    return true;
}

void FolderWindow::DebugShowOverlaySample(Pane pane, FolderView::OverlaySeverity severity)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugShowOverlaySample(severity);
}

void FolderWindow::DebugShowOverlaySampleNonModal(Pane pane, FolderView::OverlaySeverity severity)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugShowOverlaySample(FolderView::ErrorOverlayKind::Operation, severity, false);
}

void FolderWindow::DebugShowOverlaySampleBusyWithCancel(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugShowOverlaySample(FolderView::ErrorOverlayKind::Enumeration, FolderView::OverlaySeverity::Busy, true);
}

void FolderWindow::DebugShowOverlaySampleCanceled(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugShowCanceledOverlaySample();
}

void FolderWindow::DebugHideOverlaySample(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.DebugHideOverlaySample();
}

void FolderWindow::ShowPaneAlertOverlay(Pane pane,
                                        FolderView::ErrorOverlayKind kind,
                                        FolderView::OverlaySeverity severity,
                                        std::wstring title,
                                        std::wstring message,
                                        HRESULT hr,
                                        bool closable,
                                        bool blocksInput)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.hFolderView)
    {
        return;
    }

    state.folderView.ShowAlertOverlay(kind, severity, std::move(title), std::move(message), hr, closable, blocksInput);
}

void FolderWindow::DismissPaneAlertOverlay(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.hFolderView)
    {
        return;
    }

    state.folderView.DismissAlertOverlay();
}

void FolderWindow::OnDestroy()
{
    _networkChangeSubscription.reset();

    ShutdownViewers();
    ShutdownFileOperations();

    CancelSelectionSizeComputation(Pane::Left);
    CancelSelectionSizeComputation(Pane::Right);

    if (_leftPane.selectionSizeThread.joinable())
    {
        _leftPane.selectionSizeThread.request_stop();
        _leftPane.selectionSizeCv.notify_all();
        _leftPane.selectionSizeThread = std::jthread{};
    }

    if (_rightPane.selectionSizeThread.joinable())
    {
        _rightPane.selectionSizeThread.request_stop();
        _rightPane.selectionSizeCv.notify_all();
        _rightPane.selectionSizeThread = std::jthread{};
    }

    if (_draggingSplitter)
    {
        ReleaseCapture();
        _draggingSplitter = false;
    }

    auto destroyPane = [](PaneState& state)
    {
        if (state.hNavigationView)
        {
            state.navigationView.Destroy();
            state.hNavigationView = nullptr;
        }

        if (state.hFolderView)
        {
            state.folderView.Destroy();
            state.hFolderView = nullptr;
        }

        if (state.hStatusBar)
        {
            state.hStatusBar = nullptr;
        }

        state.fileSystem = nullptr;
        state.fileSystemModule.reset();
        state.pluginId.clear();
        state.currentPath.reset();
        state.updatingPath = false;
    };

    destroyPane(_leftPane);
    destroyPane(_rightPane);
}

void FolderWindow::CommandRename(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.CommandRename();
}

void FolderWindow::CommandView(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    state.folderView.CommandView();
}

void FolderWindow::CommandViewSpace(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;

    std::filesystem::path targetPath;
    const std::vector<std::filesystem::path> selectedDirs = state.folderView.GetSelectedDirectoryPaths();
    if (selectedDirs.size() == 1)
    {
        targetPath = selectedDirs.front();
    }
    else
    {
        const std::optional<std::filesystem::path> currentFolder = state.folderView.GetFolderPath();
        if (currentFolder.has_value())
        {
            targetPath = currentFolder.value();
        }
    }

    if (targetPath.empty())
    {
        return;
    }

    static_cast<void>(TryViewSpaceWithViewer(pane, targetPath));
}

void FolderWindow::SetShowSortMenuCallback(ShowSortMenuCallback callback)
{
    _showSortMenuCallback = std::move(callback);
}
