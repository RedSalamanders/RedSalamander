#include "Framework.h"

#include "FloatingTerminalWindow.h"

#include "CommandRegistry.h"
#include "CommandRuntimeState.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "Helpers.h"
#include "Keyboard.h"
#include "TerminalHostSupport.h"
#include "ViewerPluginManager.h"
#include "Win32CallbackHelpers.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "resource.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cmath>
#include <cstring>
#include <memory>
#include <ranges>
#include <span>
#include <vector>

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TabControl;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kClassName[] = L"RedSalamander.FloatingTerminalWindow";
constexpr UINT_PTR kPersistTimerId = 1u;
constexpr UINT kPersistDelayMs = 250u;
constexpr float kTabStripHeightDip = 32.0f;

[[nodiscard]] int64_t CurrentSteadyTimestampNs() noexcept
{
    return std::chrono::duration_cast<std::chrono::nanoseconds>(
               std::chrono::steady_clock::now().time_since_epoch())
        .count();
}

struct FloatingTerminalEventPayload final
{
    std::wstring tabId;
    TerminalEvent event{};
    std::chrono::steady_clock::time_point publishedAt{};
};

class FloatingTerminalWindow final : public ITerminalEventCallback
{
public:
    FloatingTerminalWindow() = default;
    ~FloatingTerminalWindow() = default;
    FloatingTerminalWindow(const FloatingTerminalWindow&) = delete;
    FloatingTerminalWindow(FloatingTerminalWindow&&) = delete;
    FloatingTerminalWindow& operator=(const FloatingTerminalWindow&) = delete;
    FloatingTerminalWindow& operator=(FloatingTerminalWindow&&) = delete;

    [[nodiscard]] HWND Create(HWND activationSource, Common::Settings::Settings& settings, const AppTheme& theme) noexcept;
    [[nodiscard]] HRESULT RestoreRememberedTabs(const std::optional<FloatingTerminalOpenRequest>& invocation, bool addWhenMissing) noexcept;
    [[nodiscard]] HRESULT BeginRememberedTabsRestore() noexcept;
    [[nodiscard]] HRESULT AddTab(const FloatingTerminalOpenRequest& request, std::wstring_view requestedTabId = {}) noexcept;
    [[nodiscard]] bool IsInputTarget(HWND targetWindow) const noexcept;
    [[nodiscard]] HRESULT RouteShortcut(HWND targetWindow,
                                        std::wstring_view commandId,
                                        const MSG& message,
                                        uint32_t modifiers,
                                        TerminalShortcutRoute& route) noexcept;
    [[nodiscard]] bool QueryCommandState(std::wstring_view commandId, CommandRuntimeState& state) noexcept;
    [[nodiscard]] bool ExecuteCommand(std::wstring_view commandId) noexcept;
    [[nodiscard]] std::optional<FloatingTerminalOpenRequest> ActiveRequest() const noexcept;
    void UpdateTheme(const AppTheme& theme) noexcept;
    void PrepareForAppShutdown() noexcept;
    [[nodiscard]] HWND GetHwnd() const noexcept { return _window.get(); }
#if defined(ENABLE_TESTS)
    [[nodiscard]] bool DebugSnapshot(FloatingTerminalDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugCloseTab(size_t index) noexcept;
    [[nodiscard]] bool DebugReorderTab(size_t fromIndex, size_t toIndex) noexcept;
    [[nodiscard]] HRESULT DebugTerminateRootProcess(uint32_t exitCode) noexcept;
#endif

    void STDMETHODCALLTYPE OnTerminalEvent(const TerminalEvent* event, void* cookie) noexcept override;

private:
    struct Tab final
    {
        FloatingTerminalWindow* owner = nullptr;
        std::wstring tabId;
        std::wstring profileId;
        std::wstring providerId;
        std::wstring canonicalPath;
        wil::com_ptr<ITerminal> terminal;
        HWND child = nullptr;
    };

    static ATOM RegisterClass(HINSTANCE instance) noexcept;
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    LRESULT HandleMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    void BuildUi();
    void Layout() noexcept;
    void ApplyTheme() noexcept;
    void SelectTab(size_t index, bool focusChild) noexcept;
    [[nodiscard]] bool CloseTab(size_t index, bool removeRestoreRecord) noexcept;
    void CloseAllTabs() noexcept;
    void ReorderTab(size_t fromIndex, size_t toIndex) noexcept;
    void RestoreNextRememberedTab() noexcept;
    void SyncSettings(std::optional<bool> cleanShutdownOpen = std::nullopt) noexcept;
    void CapturePlacement() noexcept;
    [[nodiscard]] std::optional<size_t> FindTabByTarget(HWND targetWindow) const noexcept;
    [[nodiscard]] std::optional<size_t> FindTabById(std::wstring_view tabId) const noexcept;
    [[nodiscard]] std::wstring NewStableTabId() const noexcept;

    wil::unique_hwnd _window;
    HWND _activationSource = nullptr;
    Common::Settings::Settings* _settings = nullptr;
    AppTheme _theme{};
    WindowHost _host;
    Panel* _root = nullptr;
    TabControl* _tabsControl = nullptr;
    Button* _newTabButton = nullptr;
    std::vector<std::unique_ptr<Tab>> _tabs;
    std::vector<Common::Settings::FloatingTerminalTabSettings> _pendingRestoreTabs;
    std::wstring _pendingRestoreActiveTabId;
    size_t _pendingRestoreIndex = 0u;
    bool _restoreInProgress = false;
    size_t _selectedIndex = 0u;
    bool _appShutdown = false;
    bool _closing = false;
    uint64_t _coalescedPlacementMessageCount = 0u;
    uint64_t _placementWriteEpisodeCount = 0u;
    uint64_t _stateSnapshotCount = 0u;
    std::atomic_uint64_t _pendingExitPayloadCount{0u};
};

std::unique_ptr<FloatingTerminalWindow> g_floatingTerminal;
#if defined(ENABLE_TESTS)
std::atomic_uint64_t g_debugRootExitTimingGeneration{0u};
std::atomic_uint64_t g_debugRootExitToHostCloseUs{0u};
#endif

void DropDestroyedSingleton() noexcept
{
    if (g_floatingTerminal && g_floatingTerminal->GetHwnd() == nullptr)
    {
        g_floatingTerminal.reset();
    }
}

ATOM FloatingTerminalWindow::RegisterClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0u;
    if (atom != 0u)
    {
        return atom;
    }
    WNDCLASSEXW windowClass{};
    windowClass.cbSize = sizeof(windowClass);
    windowClass.lpfnWndProc = &FloatingTerminalWindow::WindowProc;
    windowClass.hInstance = instance;
    windowClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    windowClass.hIcon = LoadIconW(instance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
    windowClass.hIconSm = windowClass.hIcon;
    windowClass.lpszClassName = kClassName;
    atom = RegisterClassExW(&windowClass);
    return atom;
}

HWND FloatingTerminalWindow::Create(HWND activationSource, Common::Settings::Settings& settings, const AppTheme& theme) noexcept
{
    _activationSource = activationSource;
    _settings = &settings;
    _theme = theme;
    const HINSTANCE instance = GetModuleHandleW(nullptr);
    if (RegisterClass(instance) == 0u)
    {
        return nullptr;
    }

    const UINT dpi = activationSource != nullptr ? GetDpiForWindow(activationSource) : GetDpiForSystem();
    const int width = MulDiv(900, static_cast<int>(std::max(dpi, 96u)), USER_DEFAULT_SCREEN_DPI);
    const int height = MulDiv(620, static_cast<int>(std::max(dpi, 96u)), USER_DEFAULT_SCREEN_DPI);
    RECT sourceRect{80, 80, 80 + width, 80 + height};
    if (activationSource != nullptr)
    {
        GetWindowRect(activationSource, &sourceRect);
    }
    const int x = sourceRect.left + MulDiv(48, static_cast<int>(std::max(dpi, 96u)), USER_DEFAULT_SCREEN_DPI);
    const int y = sourceRect.top + MulDiv(48, static_cast<int>(std::max(dpi, 96u)), USER_DEFAULT_SCREEN_DPI);
    const HWND hwnd = CreateWindowExW(WS_EX_APPWINDOW,
                                      kClassName,
                                      LoadStringResource(nullptr, IDS_PREVIEW_TAB_TERMINAL).c_str(),
                                      WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
                                      x,
                                      y,
                                      width,
                                      height,
                                      nullptr,
                                      nullptr,
                                      instance,
                                      this);
    if (hwnd == nullptr)
    {
        return nullptr;
    }

    int showCommand = SW_SHOWNORMAL;
    if (settings.terminal.has_value() && settings.terminal->floatingWindow.has_value())
    {
        showCommand = WindowPlacementPersistence::Restore(settings.terminal->floatingWindow->placement, hwnd);
    }
    ShowWindow(hwnd, showCommand);
    UpdateWindow(hwnd);
    return hwnd;
}

LRESULT CALLBACK FloatingTerminalWindow::WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_NCCREATE)
    {
        const auto* create = reinterpret_cast<CREATESTRUCTW*>(lParam);
        auto* self = static_cast<FloatingTerminalWindow*>(create != nullptr ? create->lpCreateParams : nullptr);
        if (self == nullptr)
        {
            return FALSE;
        }
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        self->_window.reset(hwnd);
        InitPostedPayloadWindow(hwnd);
    }
    auto* self = reinterpret_cast<FloatingTerminalWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return self != nullptr ? self->HandleMessage(hwnd, message, wParam, lParam) : DefWindowProcW(hwnd, message, wParam, lParam);
}

LRESULT FloatingTerminalWindow::HandleMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    if (message == WM_NCDESTROY)
    {
        KillTimer(hwnd, kPersistTimerId);
        static_cast<void>(DrainPostedPayloadsForWindow(hwnd));
        _pendingExitPayloadCount.store(0u, std::memory_order_release);
        Debug::Perf::EmitValue(L"terminal.floating.placement_messages_coalesced", _coalescedPlacementMessageCount);
        Debug::Perf::EmitValue(L"terminal.floating.placement_write_episodes", _placementWriteEpisodeCount);
        Debug::Perf::EmitValue(L"terminal.floating.state_snapshots", _stateSnapshotCount);
        Debug::Perf::EmitValue(L"terminal.floating.retained_roots", 0u);
        Debug::Perf::EmitValue(L"terminal.floating.retained_tabs", _tabs.size());
        Debug::Perf::EmitValue(L"terminal.floating.retained_terminals", _tabs.size());
        Debug::Perf::EmitValue(L"terminal.floating.retained_callbacks", _tabs.size());
        Debug::Perf::EmitValue(L"terminal.floating.retained_payloads", 0u);
        _host.Detach();
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
        if (_window.get() == hwnd)
        {
            static_cast<void>(_window.release());
        }
        return DefWindowProcW(hwnd, message, wParam, lParam);
    }
    bool handled = false;
    const LRESULT hostResult = message == WM_CREATE ? 0 : _host.HandleMessage(hwnd, message, wParam, lParam, handled);
    if (handled)
    {
        if (message == WM_SIZE)
        {
            Layout();
            if (! _closing && wParam != SIZE_MINIMIZED)
            {
                ++_coalescedPlacementMessageCount;
                SetTimer(hwnd, kPersistTimerId, kPersistDelayMs, nullptr);
            }
        }
        else if (message == WM_DPICHANGED)
        {
            UpdateTheme(_theme);
            Layout();
        }
        else if (message == WM_ACTIVATE)
        {
            ApplyWindowChromeTheme(hwnd, _theme, WindowBackdropTarget::Primary, LOWORD(wParam) != WA_INACTIVE);
        }
        return hostResult;
    }
    switch (message)
    {
        case WM_CREATE:
            if (! _host.Attach(hwnd))
            {
                return -1;
            }
            BuildUi();
            ApplyTheme();
            Layout();
            return 0;
        case WM_GETMINMAXINFO:
        {
            auto* info = reinterpret_cast<MINMAXINFO*>(lParam);
            const UINT dpi = GetDpiForWindow(hwnd);
            info->ptMinTrackSize.x = MulDiv(480, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
            info->ptMinTrackSize.y = MulDiv(300, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
            return 0;
        }
        case WM_SIZE:
            Layout();
            if (! _closing && wParam != SIZE_MINIMIZED)
            {
                ++_coalescedPlacementMessageCount;
                SetTimer(hwnd, kPersistTimerId, kPersistDelayMs, nullptr);
            }
            return 0;
        case WM_MOVE:
            if (! _closing)
            {
                ++_coalescedPlacementMessageCount;
                SetTimer(hwnd, kPersistTimerId, kPersistDelayMs, nullptr);
            }
            return 0;
        case WM_TIMER:
            if (wParam == kPersistTimerId)
            {
                KillTimer(hwnd, kPersistTimerId);
                CapturePlacement();
                ++_placementWriteEpisodeCount;
                return 0;
            }
            break;
        case WM_DPICHANGED:
        {
            const auto* suggested = reinterpret_cast<const RECT*>(lParam);
            if (suggested != nullptr)
            {
                SetWindowPos(hwnd,
                             nullptr,
                             suggested->left,
                             suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            UpdateTheme(_theme);
            Layout();
            return 0;
        }
        case WM_ACTIVATE:
            ApplyWindowChromeTheme(hwnd, _theme, WindowBackdropTarget::Primary, LOWORD(wParam) != WA_INACTIVE);
            return 0;
        case WndMsg::kTerminalSessionExited:
        {
            std::unique_ptr<FloatingTerminalEventPayload> payload = TakeMessagePayload<FloatingTerminalEventPayload>(lParam);
            if (_pendingExitPayloadCount.load(std::memory_order_acquire) != 0u)
            {
                _pendingExitPayloadCount.fetch_sub(1u, std::memory_order_acq_rel);
            }
            if (! payload)
            {
                return 0;
            }
            const std::optional<size_t> index = FindTabById(payload->tabId);
            if (! index.has_value())
            {
                return 0;
            }
            Tab& tab = *_tabs[index.value()];
            TerminalViewState view{};
            view.sizeBytes = sizeof(view);
            const HRESULT stateHr = tab.terminal ? tab.terminal->GetViewState(&view) : E_HANDLE;
            const bool matches = SUCCEEDED(stateHr) &&
                memcmp(view.instanceId.bytes, payload->event.instanceId.bytes, sizeof(view.instanceId.bytes)) == 0 &&
                view.sessionGeneration == payload->event.sessionGeneration && view.activity.lifecycleState == TerminalLifecycleState::Exited &&
                view.finalSnapshotComplete != 0u && payload->event.finalSnapshotComplete != 0u;
            CoTaskMemFree(view.title.data);
            CoTaskMemFree(view.status.data);
            if (matches)
            {
                const size_t closingIndex = index.value();
                if (CloseTab(closingIndex, true) && _tabsControl != nullptr)
                {
                    const int64_t closedAtNs = CurrentSteadyTimestampNs();
                    if (payload->event.rootExitObservedTimestampNs > 0 &&
                        closedAtNs >= payload->event.rootExitObservedTimestampNs)
                    {
                        const uint64_t rootExitToHostCloseUs =
                            static_cast<uint64_t>((closedAtNs - payload->event.rootExitObservedTimestampNs) / 1'000);
                        Debug::Perf::EmitDurationUs(
                            L"terminal.session.root_exit_to_host_close_us",
                            rootExitToHostCloseUs,
                            1u,
                            _tabs.size());
#if defined(ENABLE_TESTS)
                        g_debugRootExitToHostCloseUs.store(rootExitToHostCloseUs, std::memory_order_release);
                        g_debugRootExitTimingGeneration.fetch_add(1u, std::memory_order_acq_rel);
#endif
                    }
                    Debug::Perf::EmitDurationUs(L"terminal.session.exit_to_host_close_us",
                                                Debug::Perf::ElapsedUs(payload->publishedAt),
                                                1u,
                                                _tabs.size());
                    _tabsControl->RemoveTab(closingIndex);
                    if (_tabs.empty())
                    {
                        static_cast<void>(PostMessageW(hwnd, WndMsg::kFloatingTerminalCloseEmpty, 0u, 0));
                    }
                    else
                    {
                        SelectTab(std::min(closingIndex, _tabs.size() - 1u), false);
                    }
                }
            }
            return 0;
        }
        case WndMsg::kFloatingTerminalCloseEmpty:
            if (_tabs.empty())
            {
                SyncSettings(false);
                _closing = true;
                _window.reset();
            }
            return 0;
        case WndMsg::kFloatingTerminalRestoreNextTab:
            RestoreNextRememberedTab();
            return 0;
        case WM_CLOSE:
            if (! _closing)
            {
                CapturePlacement();
                SyncSettings(_appShutdown && ! _tabs.empty());
                _closing = true;
                CloseAllTabs();
                _window.reset();
            }
            return 0;
    }
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

void FloatingTerminalWindow::BuildUi()
{
    auto root = std::make_unique<Panel>();
    _root = root.get();
    _tabsControl = root->AddChild<TabControl>();
    _tabsControl->SetTabReorderingEnabled(true);
    _tabsControl->SetOnSelectionChanged([this](size_t index) noexcept { SelectTab(index, true); });
    _tabsControl->SetOnTabCloseRequested([this](size_t index) noexcept { return CloseTab(index, true); });
    _tabsControl->SetOnTabClosed([this](size_t index) noexcept
    {
        if (_tabs.empty() && _window)
        {
            static_cast<void>(PostMessageW(_window.get(), WndMsg::kFloatingTerminalCloseEmpty, 0u, 0));
        }
        else if (! _tabs.empty())
        {
            SelectTab(std::min(index, _tabs.size() - 1u), true);
        }
    });
    _tabsControl->SetOnTabReordered([this](size_t fromIndex, size_t toIndex) noexcept { ReorderTab(fromIndex, toIndex); });
    _newTabButton = root->AddChild<Button>(L"+");
    _newTabButton->SetTooltipText(LoadStringResource(nullptr, IDS_CMD_TERMINAL_TAB_NEW));
    _newTabButton->SetOnClick([this]() noexcept
    {
        const std::optional<FloatingTerminalOpenRequest> active = ActiveRequest();
        if (active.has_value())
        {
            static_cast<void>(AddTab(active.value()));
        }
    });
    _host.SetRoot(std::move(root));
}

void FloatingTerminalWindow::Layout() noexcept
{
    if (! _window || _root == nullptr)
    {
        return;
    }
    RECT client{};
    GetClientRect(_window.get(), &client);
    const float dpiScale = static_cast<float>(std::max(GetDpiForWindow(_window.get()), 96u)) / static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    const float widthDip = static_cast<float>(std::max(0L, client.right - client.left)) / dpiScale;
    const float heightDip = static_cast<float>(std::max(0L, client.bottom - client.top)) / dpiScale;
    _root->SetBounds(D2D1::RectF(0.0f, 0.0f, widthDip, heightDip));
    if (_tabsControl != nullptr)
    {
        _tabsControl->SetBounds(D2D1::RectF(0.0f, 0.0f, widthDip, heightDip));
    }
    if (_newTabButton != nullptr)
    {
        _newTabButton->SetBounds(D2D1::RectF(std::max(0.0f, widthDip - 36.0f), 2.0f, std::max(32.0f, widthDip - 4.0f), 30.0f));
    }

    const int headerHeight = static_cast<int>(std::lround(_host.DipsToPixels(kTabStripHeightDip)));
    const int width = std::max(0L, client.right - client.left);
    const int height = std::max(0L, client.bottom - client.top - headerHeight);
    for (const std::unique_ptr<Tab>& tab : _tabs)
    {
        if (tab && tab->child != nullptr && IsWindow(tab->child) != FALSE)
        {
            SetWindowPos(tab->child, nullptr, 0, headerHeight, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }
    _host.Invalidate();
}

void FloatingTerminalWindow::ApplyTheme() noexcept
{
    if (! _window)
    {
        return;
    }
    ApplyWindowChromeTheme(_window.get(), _theme, WindowBackdropTarget::Primary, GetActiveWindow() == _window.get());
    _host.SetTheme(MakeAppThemeDxPalette(_theme, _theme.windowBackground));
    _host.Invalidate();
}

void FloatingTerminalWindow::UpdateTheme(const AppTheme& theme) noexcept
{
    _theme = theme;
    ApplyTheme();
    const TerminalTheme terminalTheme = TerminalHostSupport::BuildTerminalTheme(_theme, _window ? GetDpiForWindow(_window.get()) : 96u);
    for (const std::unique_ptr<Tab>& tab : _tabs)
    {
        if (tab && tab->terminal)
        {
            static_cast<void>(tab->terminal->SetTheme(&terminalTheme));
        }
    }
}

std::wstring FloatingTerminalWindow::NewStableTabId() const noexcept
{
    GUID id{};
    if (FAILED(CoCreateGuid(&id)))
    {
        return {};
    }
    std::array<wchar_t, 64u> text{};
    const int length = StringFromGUID2(id, text.data(), static_cast<int>(text.size()));
    return length > 1 ? std::wstring(text.data() + 1, static_cast<size_t>(length - 3)) : std::wstring{};
}

HRESULT FloatingTerminalWindow::AddTab(const FloatingTerminalOpenRequest& request, std::wstring_view requestedTabId) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    if (! _window || _settings == nullptr || request.profileId.empty() || request.providerId.empty() || request.canonicalPath.empty())
    {
        return E_INVALIDARG;
    }
    TerminalHostSupport::OwnedTerminalLocation location = TerminalHostSupport::MakeTerminalLocation(std::filesystem::path(request.canonicalPath));
    if (location.kind == TerminalLocationKind::Unsupported)
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
    }

    auto tab = std::make_unique<Tab>();
    tab->owner = this;
    tab->tabId = requestedTabId.empty() ? NewStableTabId() : std::wstring(requestedTabId);
    tab->profileId = request.profileId;
    tab->providerId = request.providerId;
    tab->canonicalPath = request.canonicalPath;
    if (tab->tabId.empty() || FindTabById(tab->tabId).has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_DUP_NAME);
    }

    HRESULT hr = ViewerPluginManager::GetInstance().CreateTerminalInstance(tab->profileId, *_settings, tab->terminal);
    if (FAILED(hr) || ! tab->terminal)
    {
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }
    GUID instanceGuid{};
    if (FAILED(hr = CoCreateGuid(&instanceGuid)))
    {
        return hr;
    }
    TerminalOpenContext context{};
    context.sizeBytes = sizeof(context);
    context.parentWindow = _window.get();
    memcpy(context.instanceId.bytes, &instanceGuid, sizeof(instanceGuid));
    context.originalSource.folderWindowInstanceId = reinterpret_cast<uint64_t>(this);
    context.originalSource.paneInstanceId = static_cast<uint64_t>(_tabs.size() + 1u);
    context.sourceGeneration = 1u;
    context.sourceLocation = location.View();
    context.launchLocation = location.View();
    const TerminalTheme terminalTheme = TerminalHostSupport::BuildTerminalTheme(_theme, GetDpiForWindow(_window.get()));
    static_cast<void>(tab->terminal->SetTheme(&terminalTheme));
    if (FAILED(hr = tab->terminal->Open(&context)))
    {
        static_cast<void>(tab->terminal->Close());
        return hr;
    }
    if (FAILED(hr = tab->terminal->GetChildWindow(&tab->child)) || tab->child == nullptr || IsWindow(tab->child) == FALSE ||
        GetParent(tab->child) != _window.get())
    {
        static_cast<void>(tab->terminal->Close());
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE);
    }
    if (FAILED(hr = tab->terminal->SetCallback(this, tab.get())))
    {
        static_cast<void>(tab->terminal->Close());
        return hr;
    }

    const std::wstring title = std::filesystem::path(tab->canonicalPath).filename().wstring();
    _tabs.push_back(std::move(tab));
    if (_tabsControl != nullptr)
    {
        _tabsControl->AddTab<Panel>(title.empty() ? LoadStringResource(nullptr, IDS_PREVIEW_TAB_TERMINAL) : title);
        _tabsControl->SetTabClosable(_tabs.size() - 1u, true);
        _tabsControl->SetTabTooltip(_tabs.size() - 1u, _tabs.back()->canonicalPath);
    }
    SelectTab(_tabs.size() - 1u, true);
    Layout();
    SyncSettings(false);
    Debug::Perf::EmitDurationUs(L"terminal.floating.tab_open_to_visible_us",
                                Debug::Perf::ElapsedUs(startedAt),
                                _tabs.size(),
                                1u);
    return S_OK;
}

HRESULT FloatingTerminalWindow::RestoreRememberedTabs(
    const std::optional<FloatingTerminalOpenRequest>& invocation, bool addWhenMissing) noexcept
{
    if (_settings == nullptr)
    {
        return E_HANDLE;
    }
    std::vector<Common::Settings::FloatingTerminalTabSettings> remembered;
    std::wstring activeTabId;
    if (_settings->terminal.has_value() && _settings->terminal->floatingWindow.has_value())
    {
        remembered = _settings->terminal->floatingWindow->tabs;
        activeTabId = _settings->terminal->floatingWindow->activeTabId;
    }
    HRESULT firstFailure = S_OK;
    for (const Common::Settings::FloatingTerminalTabSettings& saved : remembered)
    {
        const FloatingTerminalOpenRequest request{
            .profileId = saved.profileId, .providerId = saved.providerId, .canonicalPath = saved.canonicalPath};
        const HRESULT hr = AddTab(request, saved.tabId);
        if (FAILED(hr) && SUCCEEDED(firstFailure))
        {
            firstFailure = hr;
        }
    }

    std::optional<size_t> selection = FindTabById(activeTabId);
    if (invocation.has_value())
    {
        const auto match = std::ranges::find_if(_tabs, [&](const std::unique_ptr<Tab>& tab) noexcept
        {
            return tab && tab->profileId == invocation->profileId && tab->providerId == invocation->providerId &&
                CompareStringOrdinal(tab->canonicalPath.c_str(), -1, invocation->canonicalPath.c_str(), -1, TRUE) == CSTR_EQUAL;
        });
        if (match != _tabs.end())
        {
            selection = static_cast<size_t>(match - _tabs.begin());
        }
        else if (addWhenMissing)
        {
            const HRESULT addHr = AddTab(invocation.value());
            if (FAILED(addHr) && SUCCEEDED(firstFailure))
            {
                firstFailure = addHr;
            }
            if (SUCCEEDED(addHr))
            {
                selection = _tabs.size() - 1u;
            }
        }
    }
    if (_tabs.empty() && invocation.has_value())
    {
        const HRESULT addHr = AddTab(invocation.value());
        if (FAILED(addHr))
        {
            return addHr;
        }
        selection = 0u;
    }
    if (!_tabs.empty())
    {
        SelectTab(selection.value_or(0u), true);
        SyncSettings(false);
        return S_OK;
    }
    return FAILED(firstFailure) ? firstFailure : S_FALSE;
}

HRESULT FloatingTerminalWindow::BeginRememberedTabsRestore() noexcept
{
    if (_settings == nullptr || !_window)
    {
        return E_HANDLE;
    }
    if (!_settings->terminal.has_value() || !_settings->terminal->floatingWindow.has_value() ||
        _settings->terminal->floatingWindow->tabs.empty())
    {
        return S_FALSE;
    }
    _pendingRestoreTabs = _settings->terminal->floatingWindow->tabs;
    _pendingRestoreActiveTabId = _settings->terminal->floatingWindow->activeTabId;
    _pendingRestoreIndex = 0u;
    _restoreInProgress = true;
    if (PostMessageW(_window.get(), WndMsg::kFloatingTerminalRestoreNextTab, 0u, 0) == FALSE)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
        _restoreInProgress = false;
        _pendingRestoreTabs.clear();
        _pendingRestoreActiveTabId.clear();
        return hr;
    }
    return S_OK;
}

void FloatingTerminalWindow::RestoreNextRememberedTab() noexcept
{
    if (!_restoreInProgress || _closing || !_window)
    {
        return;
    }
    if (_pendingRestoreIndex < _pendingRestoreTabs.size())
    {
        const Common::Settings::FloatingTerminalTabSettings& saved = _pendingRestoreTabs[_pendingRestoreIndex++];
        const FloatingTerminalOpenRequest request{
            .profileId = saved.profileId, .providerId = saved.providerId, .canonicalPath = saved.canonicalPath};
        static_cast<void>(AddTab(request, saved.tabId));
    }
    if (_pendingRestoreIndex < _pendingRestoreTabs.size())
    {
        if (PostMessageW(_window.get(), WndMsg::kFloatingTerminalRestoreNextTab, 0u, 0) != FALSE)
        {
            return;
        }
    }

    _restoreInProgress = false;
    const std::optional<size_t> selection = FindTabById(_pendingRestoreActiveTabId);
    _pendingRestoreTabs.clear();
    _pendingRestoreActiveTabId.clear();
    if (!_tabs.empty())
    {
        SelectTab(selection.value_or(0u), true);
        SyncSettings(false);
    }
    else
    {
        static_cast<void>(PostMessageW(_window.get(), WndMsg::kFloatingTerminalCloseEmpty, 0u, 0));
    }
}

void FloatingTerminalWindow::SelectTab(size_t index, bool focusChild) noexcept
{
    if (index >= _tabs.size())
    {
        return;
    }
    _selectedIndex = index;
    for (size_t current = 0u; current < _tabs.size(); ++current)
    {
        const std::unique_ptr<Tab>& tab = _tabs[current];
        if (tab && tab->child != nullptr && IsWindow(tab->child) != FALSE)
        {
            ShowWindow(tab->child, current == index ? SW_SHOW : SW_HIDE);
        }
    }
    if (_tabsControl != nullptr && _tabsControl->GetSelectedIndex() != index)
    {
        _tabsControl->SetSelectedIndex(index);
        _host.Invalidate();
    }
    if (focusChild && _tabs[index]->child != nullptr)
    {
        SetFocus(_tabs[index]->child);
    }
    SyncSettings(false);
}

bool FloatingTerminalWindow::CloseTab(size_t index, bool removeRestoreRecord) noexcept
{
    if (index >= _tabs.size())
    {
        return false;
    }
    std::unique_ptr<Tab> tab = std::move(_tabs[index]);
    if (tab->terminal)
    {
        static_cast<void>(tab->terminal->SetCallback(nullptr, nullptr));
        static_cast<void>(tab->terminal->Close());
        tab->terminal.reset();
    }
    tab->child = nullptr;
    _tabs.erase(_tabs.begin() + static_cast<ptrdiff_t>(index));
    if (_tabs.empty())
    {
        _selectedIndex = 0u;
    }
    else if (_selectedIndex >= _tabs.size())
    {
        _selectedIndex = _tabs.size() - 1u;
    }
    if (removeRestoreRecord)
    {
        SyncSettings(false);
    }
    return true;
}

void FloatingTerminalWindow::CloseAllTabs() noexcept
{
    while (! _tabs.empty())
    {
        static_cast<void>(CloseTab(_tabs.size() - 1u, false));
    }
}

void FloatingTerminalWindow::ReorderTab(size_t fromIndex, size_t toIndex) noexcept
{
    if (fromIndex >= _tabs.size() || toIndex >= _tabs.size() || fromIndex == toIndex)
    {
        return;
    }
    std::unique_ptr<Tab> moved = std::move(_tabs[fromIndex]);
    _tabs.erase(_tabs.begin() + static_cast<ptrdiff_t>(fromIndex));
    _tabs.insert(_tabs.begin() + static_cast<ptrdiff_t>(toIndex), std::move(moved));
    if (_selectedIndex == fromIndex)
    {
        _selectedIndex = toIndex;
    }
    else if (fromIndex < _selectedIndex && toIndex >= _selectedIndex)
    {
        --_selectedIndex;
    }
    else if (fromIndex > _selectedIndex && toIndex <= _selectedIndex)
    {
        ++_selectedIndex;
    }
    SyncSettings(false);
}

void FloatingTerminalWindow::CapturePlacement() noexcept
{
    if (_settings == nullptr || !_window)
    {
        return;
    }
    if (!_settings->terminal.has_value())
    {
        _settings->terminal.emplace();
    }
    if (!_settings->terminal->floatingWindow.has_value())
    {
        _settings->terminal->floatingWindow.emplace();
    }
    const std::optional<Common::Settings::WindowPlacement> placement = WindowPlacementPersistence::Capture(_window.get());
    if (placement.has_value())
    {
        _settings->terminal->floatingWindow->placement = placement.value();
    }
}

void FloatingTerminalWindow::SyncSettings(std::optional<bool> cleanShutdownOpen) noexcept
{
    if (_settings == nullptr)
    {
        return;
    }
    if (!_settings->terminal.has_value())
    {
        _settings->terminal.emplace();
    }
    if (!_settings->terminal->floatingWindow.has_value())
    {
        _settings->terminal->floatingWindow.emplace();
    }
    ++_stateSnapshotCount;
    Common::Settings::FloatingTerminalWindowSettings& saved = _settings->terminal->floatingWindow.value();
    if (cleanShutdownOpen.has_value())
    {
        saved.wasOpenAtCleanShutdown = cleanShutdownOpen.value();
    }
    if (_restoreInProgress)
    {
        saved.activeTabId = _pendingRestoreActiveTabId;
        saved.tabs = _pendingRestoreTabs;
        return;
    }
    saved.activeTabId = _tabs.empty() ? std::wstring{} : _tabs[_selectedIndex]->tabId;
    saved.tabs.clear();
    saved.tabs.reserve(_tabs.size());
    for (const std::unique_ptr<Tab>& tab : _tabs)
    {
        if (tab)
        {
            saved.tabs.push_back(Common::Settings::FloatingTerminalTabSettings{
                .tabId = tab->tabId, .profileId = tab->profileId, .providerId = tab->providerId, .canonicalPath = tab->canonicalPath});
        }
    }
}

std::optional<size_t> FloatingTerminalWindow::FindTabByTarget(HWND targetWindow) const noexcept
{
    for (size_t index = 0u; index < _tabs.size(); ++index)
    {
        const HWND child = _tabs[index] ? _tabs[index]->child : nullptr;
        if (child != nullptr && IsWindow(child) != FALSE && (targetWindow == child || IsChild(child, targetWindow) != FALSE))
        {
            return index;
        }
    }
    return std::nullopt;
}

std::optional<size_t> FloatingTerminalWindow::FindTabById(std::wstring_view tabId) const noexcept
{
    const auto found = std::ranges::find_if(_tabs, [&](const std::unique_ptr<Tab>& tab) noexcept { return tab && tab->tabId == tabId; });
    return found != _tabs.end() ? std::optional<size_t>{static_cast<size_t>(found - _tabs.begin())} : std::nullopt;
}

bool FloatingTerminalWindow::IsInputTarget(HWND targetWindow) const noexcept
{
    return targetWindow != nullptr && FindTabByTarget(targetWindow).has_value();
}

HRESULT FloatingTerminalWindow::RouteShortcut(HWND targetWindow,
                                              std::wstring_view commandId,
                                              const MSG& message,
                                              uint32_t modifiers,
                                              TerminalShortcutRoute& route) noexcept
{
    route = TerminalShortcutRoute::PassThrough;
    const std::optional<size_t> index = FindTabByTarget(targetWindow);
    if (! index.has_value() || !_tabs[index.value()]->terminal)
    {
        return E_HANDLE;
    }
    wil::com_ptr<ITerminalActions> actions;
    HRESULT hr = _tabs[index.value()]->terminal->QueryInterface(__uuidof(ITerminalActions), actions.put_void());
    if (FAILED(hr) || !actions)
    {
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }
    TerminalViewState view{};
    view.sizeBytes = sizeof(view);
    hr = _tabs[index.value()]->terminal->GetViewState(&view);
    if (FAILED(hr))
    {
        return hr;
    }
    CoTaskMemFree(view.title.data);
    CoTaskMemFree(view.status.data);
    TerminalShortcutRequest request{};
    request.sizeBytes = sizeof(request);
    request.commandId = {commandId.data(), static_cast<uint32_t>(commandId.size())};
    request.instanceId = view.instanceId;
    request.sessionGeneration = view.sessionGeneration;
    request.message = message.message;
    request.virtualKey = static_cast<uint32_t>(message.wParam);
    request.scanCode = Common::Keyboard::ScanCodeFromKeyMessageLParam(message.lParam);
    request.extended = Common::Keyboard::IsExtendedKeyMessageLParam(message.lParam) ? 1u : 0u;
    request.systemKey = message.message == WM_SYSKEYDOWN ? 1u : 0u;
    request.repeatCount = static_cast<uint32_t>(message.lParam & 0xFFFFu);
    request.previousDown = (static_cast<ULONG_PTR>(message.lParam) & (1ull << 30u)) != 0u ? 1u : 0u;
    request.modifierFlags = modifiers & 0x7u;
    const auto addModifier = [&](int virtualKey, uint32_t flag) noexcept
    {
        if ((GetKeyState(virtualKey) & 0x8000) != 0)
        {
            request.modifierFlags |= flag;
        }
    };
    addModifier(VK_LCONTROL, TerminalShortcutModifierLeftCtrl);
    addModifier(VK_RCONTROL, TerminalShortcutModifierRightCtrl);
    addModifier(VK_LMENU, TerminalShortcutModifierLeftAlt);
    addModifier(VK_RMENU, TerminalShortcutModifierRightAlt);
    addModifier(VK_LSHIFT, TerminalShortcutModifierLeftShift);
    addModifier(VK_RSHIFT, TerminalShortcutModifierRightShift);
    return actions->RouteShortcut(&request, &route);
}

bool FloatingTerminalWindow::QueryCommandState(std::wstring_view commandId, CommandRuntimeState& state) noexcept
{
    state = {};
    state.enabled = false;
    if (_tabs.empty() || _selectedIndex >= _tabs.size() || ! _tabs[_selectedIndex])
    {
        return true;
    }

    const Tab& tab = *_tabs[_selectedIndex];
    TerminalViewState view{};
    view.sizeBytes = sizeof(view);
    if (tab.terminal && SUCCEEDED(tab.terminal->GetViewState(&view)))
    {
        wil::unique_cotaskmem_string title(view.title.data);
        wil::unique_cotaskmem_string status(view.status.data);
        state.terminalIdentityPresent = true;
        state.terminalInstanceId = view.instanceId;
        state.terminalSessionGeneration = view.sessionGeneration;

        const bool pluginAction = IsTerminalPluginActionId(commandId);
        if (pluginAction)
        {
            wil::com_ptr<ITerminalActions> actions;
            if (FAILED(tab.terminal->QueryInterface(__uuidof(ITerminalActions), actions.put_void())) || ! actions)
            {
                return true;
            }
            TerminalActionRequest request{};
            request.sizeBytes = sizeof(request);
            request.commandId = {commandId.data(), static_cast<uint32_t>(commandId.size())};
            request.instanceId = view.instanceId;
            request.sessionGeneration = view.sessionGeneration;
            TerminalActionState actionState{};
            actionState.sizeBytes = sizeof(actionState);
            if (SUCCEEDED(actions->GetActionState(&request, &actionState)))
            {
                state.enabled = actionState.enabled != 0u;
                state.checked = actionState.checked != 0u;
            }
            return true;
        }
    }

    if (commandId == L"cmd/terminal/close" || commandId == L"cmd/terminal/contextMenu" ||
        commandId == L"cmd/terminal/sessionMenu" || commandId == L"cmd/terminal/tab/next" ||
        commandId == L"cmd/terminal/tab/previous" || commandId == L"cmd/terminal/tab/last")
    {
        state.enabled = true;
        return true;
    }
    if (commandId == L"cmd/terminal/tab/new" || commandId == L"cmd/terminal/openFloatingWindow")
    {
        state.enabled = ActiveRequest().has_value();
        return true;
    }
    constexpr std::wstring_view selectPrefix = L"cmd/terminal/tab/select/";
    if (commandId.starts_with(selectPrefix) && commandId.size() == selectPrefix.size() + 1u)
    {
        const wchar_t digit = commandId.back();
        state.enabled = digit >= L'1' && digit <= L'8' && static_cast<size_t>(digit - L'1') < _tabs.size();
        return true;
    }
    return true;
}

bool FloatingTerminalWindow::ExecuteCommand(std::wstring_view commandId) noexcept
{
    if (_tabs.empty() || _selectedIndex >= _tabs.size())
    {
        return false;
    }
    const auto showCommandMenu = [&](std::span<const std::wstring_view> actionIds) noexcept
    {
        if (!_window)
        {
            return false;
        }
        wil::unique_hmenu menu(CreatePopupMenu());
        if (!menu)
        {
            return false;
        }
        constexpr UINT kFirstCommand = 0x7600u;
        for (size_t index = 0u; index < actionIds.size(); ++index)
        {
            const std::wstring_view actionId = actionIds[index];
            if (actionId.empty())
            {
                if (AppendMenuW(menu.get(), MF_SEPARATOR, 0u, nullptr) == FALSE) return false;
                continue;
            }
            const CommandInfo* info = FindCommandInfo(actionId);
            const std::wstring label = info != nullptr && info->displayNameStringId != 0u
                ? LoadStringResource(nullptr, info->displayNameStringId)
                : std::wstring(actionId);
            CommandRuntimeState actionState{};
            const bool enabled = QueryCommandState(actionId, actionState) && actionState.enabled;
            if (AppendMenuW(menu.get(), static_cast<UINT>(MF_STRING | (enabled ? MF_ENABLED : MF_GRAYED)),
                            kFirstCommand + static_cast<UINT>(index), label.c_str()) == FALSE)
            {
                return false;
            }
        }
        POINT anchor{};
        if (const HWND child = _tabs[_selectedIndex]->child; child != nullptr)
        {
            RECT bounds{};
            if (GetWindowRect(child, &bounds) != FALSE)
            {
                anchor = {bounds.left + MulDiv(12, static_cast<int>(GetDpiForWindow(_window.get())), USER_DEFAULT_SCREEN_DPI),
                          bounds.top + MulDiv(12, static_cast<int>(GetDpiForWindow(_window.get())), USER_DEFAULT_SCREEN_DPI)};
            }
        }
        if (anchor.x == 0 && anchor.y == 0)
        {
            RECT windowBounds{};
            if (GetWindowRect(_window.get(), &windowBounds) != FALSE)
            {
                anchor = {windowBounds.left + MulDiv(12, static_cast<int>(GetDpiForWindow(_window.get())), USER_DEFAULT_SCREEN_DPI),
                          windowBounds.top + MulDiv(12, static_cast<int>(GetDpiForWindow(_window.get())), USER_DEFAULT_SCREEN_DPI)};
            }
        }
        SetForegroundWindow(_window.get());
        const UINT selected = static_cast<UINT>(TrackPopupMenuEx(
            menu.get(), TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD,
            anchor.x, anchor.y, _window.get(), nullptr));
        if (selected < kFirstCommand || selected >= kFirstCommand + actionIds.size()) return true;
        const std::wstring_view selectedAction = actionIds[selected - kFirstCommand];
        return !selectedAction.empty() && ExecuteCommand(selectedAction);
    };
    if (commandId == L"cmd/terminal/contextMenu")
    {
        constexpr std::array<std::wstring_view, 7u> actions{{
            L"cmd/terminal/copy", L"cmd/terminal/paste", L"cmd/terminal/selectAll", std::wstring_view{},
            L"cmd/terminal/find", L"cmd/terminal/suggestions", L"cmd/terminal/close"}};
        return showCommandMenu(actions);
    }
    if (commandId == L"cmd/terminal/sessionMenu")
    {
        constexpr std::array<std::wstring_view, 3u> actions{{
            L"cmd/terminal/tab/new", std::wstring_view{}, L"cmd/terminal/close"}};
        return showCommandMenu(actions);
    }
    if (commandId == L"cmd/terminal/tab/new")
    {
        const std::optional<FloatingTerminalOpenRequest> request = ActiveRequest();
        return request.has_value() && SUCCEEDED(AddTab(request.value()));
    }
    if (commandId == L"cmd/terminal/close")
    {
        const size_t index = _selectedIndex;
        if (!CloseTab(index, true))
        {
            return false;
        }
        if (_tabsControl != nullptr)
        {
            _tabsControl->RemoveTab(index);
        }
        if (_tabs.empty() && _window)
        {
            static_cast<void>(PostMessageW(_window.get(), WndMsg::kFloatingTerminalCloseEmpty, 0u, 0));
        }
        else if (!_tabs.empty())
        {
            SelectTab(std::min(index, _tabs.size() - 1u), true);
        }
        return true;
    }
    if (commandId == L"cmd/terminal/tab/next" || commandId == L"cmd/terminal/tab/previous")
    {
        const bool forward = commandId == L"cmd/terminal/tab/next";
        const size_t target = forward ? (_selectedIndex + 1u) % _tabs.size() : (_selectedIndex + _tabs.size() - 1u) % _tabs.size();
        SelectTab(target, true);
        return true;
    }
    if (commandId == L"cmd/terminal/tab/last")
    {
        SelectTab(_tabs.size() - 1u, true);
        return true;
    }
    constexpr std::wstring_view selectPrefix = L"cmd/terminal/tab/select/";
    if (commandId.starts_with(selectPrefix) && commandId.size() == selectPrefix.size() + 1u)
    {
        const wchar_t digit = commandId.back();
        const size_t index = digit >= L'1' && digit <= L'8' ? static_cast<size_t>(digit - L'1') : _tabs.size();
        if (index >= _tabs.size())
        {
            return false;
        }
        SelectTab(index, true);
        return true;
    }

    const bool pluginAction = IsTerminalPluginActionId(commandId);
    if (!pluginAction || !_tabs[_selectedIndex]->terminal)
    {
        return false;
    }
    wil::com_ptr<ITerminalActions> actions;
    if (FAILED(_tabs[_selectedIndex]->terminal->QueryInterface(__uuidof(ITerminalActions), actions.put_void())) || !actions)
    {
        return false;
    }
    TerminalViewState view{};
    view.sizeBytes = sizeof(view);
    if (FAILED(_tabs[_selectedIndex]->terminal->GetViewState(&view)))
    {
        return false;
    }
    CoTaskMemFree(view.title.data);
    CoTaskMemFree(view.status.data);
    TerminalActionRequest request{};
    request.sizeBytes = sizeof(request);
    request.commandId = {commandId.data(), static_cast<uint32_t>(commandId.size())};
    request.instanceId = view.instanceId;
    request.sessionGeneration = view.sessionGeneration;
    return actions->ExecuteAction(&request) == S_OK;
}

std::optional<FloatingTerminalOpenRequest> FloatingTerminalWindow::ActiveRequest() const noexcept
{
    if (_tabs.empty() || _selectedIndex >= _tabs.size() || !_tabs[_selectedIndex])
    {
        return std::nullopt;
    }
    const Tab& tab = *_tabs[_selectedIndex];
    return FloatingTerminalOpenRequest{
        .profileId = tab.profileId, .providerId = tab.providerId, .canonicalPath = tab.canonicalPath};
}

void FloatingTerminalWindow::PrepareForAppShutdown() noexcept
{
    _appShutdown = true;
    CapturePlacement();
    SyncSettings(!_tabs.empty());
}

void FloatingTerminalWindow::OnTerminalEvent(const TerminalEvent* event, void* cookie) noexcept
{
    auto* tab = static_cast<Tab*>(cookie);
    if (event == nullptr || event->sizeBytes < sizeof(TerminalEvent) || event->kind != TerminalEventKind::RootSessionExited ||
        event->finalSnapshotComplete == 0u || tab == nullptr || tab->owner != this || !_window)
    {
        return;
    }
    auto payload = std::make_unique<FloatingTerminalEventPayload>();
    payload->tabId = tab->tabId;
    payload->event = *event;
    payload->publishedAt = std::chrono::steady_clock::now();
    if (PostMessagePayload(_window.get(), WndMsg::kTerminalSessionExited, 0u, std::move(payload)))
    {
        _pendingExitPayloadCount.fetch_add(1u, std::memory_order_acq_rel);
    }
}

#if defined(ENABLE_TESTS)
bool FloatingTerminalWindow::DebugSnapshot(FloatingTerminalDebugSnapshot& out) const noexcept
{
    out = {};
    out.root = _window.get();
    out.tabCount = _tabs.size();
    out.selectedIndex = _selectedIndex;
    out.pendingExitPayloadCount = _pendingExitPayloadCount.load(std::memory_order_acquire);
    out.selectedChild = _tabs.empty() || _selectedIndex >= _tabs.size() ? nullptr : _tabs[_selectedIndex]->child;
    if (!_tabs.empty() && _selectedIndex < _tabs.size() && _tabs[_selectedIndex] && _tabs[_selectedIndex]->terminal)
    {
        TerminalViewState view{};
        view.sizeBytes = sizeof(view);
        if (SUCCEEDED(_tabs[_selectedIndex]->terminal->GetViewState(&view)))
        {
            out.selectedLifecycle = view.activity.lifecycleState;
            out.selectedActivityTrust = view.activity.activityTrust;
            out.selectedSessionGeneration = view.sessionGeneration;
            out.selectedFinalSnapshotComplete = view.finalSnapshotComplete != 0u;
            out.selectedIdleAtPrimaryPrompt = view.activity.idleAtPrimaryPrompt != 0u;
        }
        CoTaskMemFree(view.title.data);
        CoTaskMemFree(view.status.data);
    }
    for (const std::unique_ptr<Tab>& tab : _tabs)
    {
        if (tab)
        {
            out.tabIds.push_back(tab->tabId);
            out.paths.push_back(tab->canonicalPath);
        }
    }
    return _window != nullptr;
}

bool FloatingTerminalWindow::DebugCloseTab(size_t index) noexcept
{
    if (!CloseTab(index, true))
    {
        return false;
    }
    if (_tabsControl != nullptr)
    {
        _tabsControl->RemoveTab(index);
    }
    if (_tabs.empty() && _window)
    {
        static_cast<void>(PostMessageW(_window.get(), WndMsg::kFloatingTerminalCloseEmpty, 0u, 0));
    }
    return true;
}

bool FloatingTerminalWindow::DebugReorderTab(size_t fromIndex, size_t toIndex) noexcept
{
    if (fromIndex >= _tabs.size() || toIndex >= _tabs.size())
    {
        return false;
    }
    ReorderTab(fromIndex, toIndex);
    return true;
}

HRESULT FloatingTerminalWindow::DebugTerminateRootProcess(uint32_t exitCode) noexcept
{
    if (_tabs.empty() || _selectedIndex >= _tabs.size() || !_tabs[_selectedIndex] || !_tabs[_selectedIndex]->terminal)
    {
        return E_HANDLE;
    }
    return TerminalHostSupport::DebugTerminateRootProcess(_tabs[_selectedIndex]->terminal.get(), exitCode);
}
#endif
} // namespace

HRESULT ShowFloatingTerminalWindow(HWND activationSource,
                                   Common::Settings::Settings& settings,
                                   const FloatingTerminalOpenRequest& request,
                                   const AppTheme& theme) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    DropDestroyedSingleton();
    if (g_floatingTerminal && g_floatingTerminal->GetHwnd() != nullptr)
    {
        const HRESULT addHr = g_floatingTerminal->AddTab(request);
        if (SUCCEEDED(addHr))
        {
            ShowWindow(g_floatingTerminal->GetHwnd(), SW_RESTORE);
            SetForegroundWindow(g_floatingTerminal->GetHwnd());
        }
        return addHr;
    }
    auto window = std::make_unique<FloatingTerminalWindow>();
    if (window->Create(activationSource, settings, theme) == nullptr)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    g_floatingTerminal = std::move(window);
    const HRESULT restoreHr = g_floatingTerminal->RestoreRememberedTabs(request, true);
    if (FAILED(restoreHr))
    {
        SendMessageW(g_floatingTerminal->GetHwnd(), WM_CLOSE, 0u, 0);
        DropDestroyedSingleton();
        return restoreHr;
    }
    SetForegroundWindow(g_floatingTerminal->GetHwnd());
    Debug::Perf::EmitDurationUs(L"terminal.floating.window_open_to_visible_us",
                                Debug::Perf::ElapsedUs(startedAt),
                                g_floatingTerminal ? 1u : 0u,
                                g_floatingTerminal ? 1u : 0u);
    return S_OK;
}

HRESULT RestoreFloatingTerminalWindowAfterStartup(
    HWND activationSource, Common::Settings::Settings& settings, const AppTheme& theme) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    DropDestroyedSingleton();
    if (g_floatingTerminal)
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }
    if (!settings.terminal.has_value() || !settings.terminal->floatingWindow.has_value() ||
        !settings.terminal->floatingWindow->wasOpenAtCleanShutdown || settings.terminal->floatingWindow->tabs.empty())
    {
        return S_FALSE;
    }
    auto window = std::make_unique<FloatingTerminalWindow>();
    if (window->Create(activationSource, settings, theme) == nullptr)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    g_floatingTerminal = std::move(window);
    const HRESULT restoreHr = g_floatingTerminal->BeginRememberedTabsRestore();
    if (FAILED(restoreHr) || restoreHr == S_FALSE)
    {
        SendMessageW(g_floatingTerminal->GetHwnd(), WM_CLOSE, 0u, 0);
        DropDestroyedSingleton();
        return FAILED(restoreHr) ? restoreHr : E_FAIL;
    }
    Debug::Perf::EmitDurationUs(L"terminal.floating.window_open_to_visible_us",
                                Debug::Perf::ElapsedUs(startedAt),
                                settings.terminal->floatingWindow->tabs.size(),
                                1u);
    return S_OK;
}

void PrepareFloatingTerminalWindowForAppShutdown() noexcept
{
    DropDestroyedSingleton();
    if (g_floatingTerminal)
    {
        g_floatingTerminal->PrepareForAppShutdown();
    }
}

void UpdateFloatingTerminalWindowTheme(const AppTheme& theme) noexcept
{
    DropDestroyedSingleton();
    if (g_floatingTerminal)
    {
        g_floatingTerminal->UpdateTheme(theme);
    }
}

HWND GetFloatingTerminalWindowHandle() noexcept
{
    DropDestroyedSingleton();
    return g_floatingTerminal ? g_floatingTerminal->GetHwnd() : nullptr;
}

bool IsFloatingTerminalInputTarget(HWND targetWindow) noexcept
{
    DropDestroyedSingleton();
    return g_floatingTerminal && g_floatingTerminal->IsInputTarget(targetWindow);
}

HRESULT RouteFloatingTerminalShortcut(HWND targetWindow,
                                      std::wstring_view commandId,
                                      const MSG& message,
                                      uint32_t modifiers,
                                      TerminalShortcutRoute& route) noexcept
{
    DropDestroyedSingleton();
    return g_floatingTerminal ? g_floatingTerminal->RouteShortcut(targetWindow, commandId, message, modifiers, route) : E_HANDLE;
}

bool ExecuteFloatingTerminalCommand(std::wstring_view commandId) noexcept
{
    DropDestroyedSingleton();
    return g_floatingTerminal && g_floatingTerminal->ExecuteCommand(commandId);
}

bool QueryFloatingTerminalCommandState(std::wstring_view commandId, CommandRuntimeState& state) noexcept
{
    DropDestroyedSingleton();
    if (! g_floatingTerminal)
    {
        state = {};
        state.enabled = false;
        return true;
    }
    return g_floatingTerminal->QueryCommandState(commandId, state);
}

std::optional<FloatingTerminalOpenRequest> GetActiveFloatingTerminalRequest() noexcept
{
    DropDestroyedSingleton();
    return g_floatingTerminal ? g_floatingTerminal->ActiveRequest() : std::nullopt;
}

#if defined(ENABLE_TESTS)
bool DebugGetFloatingTerminalSnapshot(FloatingTerminalDebugSnapshot& out) noexcept
{
    DropDestroyedSingleton();
    return g_floatingTerminal && g_floatingTerminal->DebugSnapshot(out);
}

bool DebugCloseFloatingTerminalTab(size_t index) noexcept
{
    DropDestroyedSingleton();
    return g_floatingTerminal && g_floatingTerminal->DebugCloseTab(index);
}

bool DebugReorderFloatingTerminalTab(size_t fromIndex, size_t toIndex) noexcept
{
    DropDestroyedSingleton();
    return g_floatingTerminal && g_floatingTerminal->DebugReorderTab(fromIndex, toIndex);
}

HRESULT DebugTerminateFloatingTerminalRootProcess(uint32_t exitCode) noexcept
{
    DropDestroyedSingleton();
    return g_floatingTerminal ? g_floatingTerminal->DebugTerminateRootProcess(exitCode) : E_HANDLE;
}

void DebugGetFloatingTerminalRootExitTiming(uint64_t& generation, uint64_t& durationUs) noexcept
{
    generation = g_debugRootExitTimingGeneration.load(std::memory_order_acquire);
    durationUs = g_debugRootExitToHostCloseUs.load(std::memory_order_acquire);
}
#endif
