#include <algorithm>
#include <array>
#include <atomic>
#include <cerrno>
#include <chrono>
#include <cmath>
#include <cstddef>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <span>
#include <string>
#include <system_error>
#include <unordered_map>
#include <vector>

#include "AppTheme.h"
#include "DxUi/DxUi.FocusRestore.h"
#include "FluentIcons.h"
#include "LocalizationManager.h"
#include "MaskSyntax.h"
#include "NavigationLocation.h"
#include "SettingsStore.h"
#include "resource.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <shlobj_core.h>
#include <shellapi.h>
#include <strsafe.h>
#include <winnetwk.h>

#pragma comment(lib, "Mpr.lib")

// Define the ETW provider for RedSalamander.exe
// Each executable must have its own provider instance with the same GUID
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"

#include "ExceptionHelpers.h"
#include "Version.h"

#include "CommandRegistry.h"
#include "CompareDirectoriesWindow.h"
#include "ConnectionCredentialPromptDialog.h"
#include "ConnectionManagerWindow.h"
#include "CrashHandler.h"
#include "CrashQuarantine.h"
#include "DirectoryInfoCache.h"
#include "DxUiThemePalette.h"
#include "FileActionLauncher.h"
#include "FileActionResolver.h"
#include "FileSystemPluginManager.h"
#include "FindFilesWindow.h"
#include "FolderWindow.h"
#include "Framework.h"
#include "HostServices.h"
#include "IconCache.h"
#include "ManagePluginsDialog.h"
#include "Preferences.h"
#include "RedSalamander.h"
#include "SessionState.h"
#include "SettingsHotReload.h"
#include "SettingsSave.h"
#include "SettingsSchemaExport.h"
#include "ShortcutDefaults.h"
#include "ShortcutManager.h"
#include "ShortcutsWindow.h"
#include "SplashScreen.h"
#include "StartupMetrics.h"
#include "Ui/AnimationDispatcher.h"
#include "ViewerPluginManager.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "WindowSizing.h"

#ifdef ENABLE_TESTS
#include "CommandDispatch.Debug.h"
#include "Commands.SelfTest.h"
#include "CompareDirectoriesEngine.SelfTest.h"
#include "FolderWindow.FileOperations.SelfTest.h"
#endif

PCWSTR REDSALAMANDER_TEXT_VERSION = L"RedSalamander " VERSINFO_VERSION;

constexpr int MAX_LOADSTRING = 100;

// Global Variables:
HINSTANCE g_hInstance = nullptr; // current instance
FolderWindow g_folderWindow;     // folder window (integrates NavigationView + FolderView)
std::atomic<HWND> g_hFolderWindow{nullptr};
ThemeMode g_themeMode = ThemeMode::System;
Common::Settings::Settings g_settings;

void ApplyThemeId(HWND hWnd, std::wstring_view themeId);

namespace
{
constexpr wchar_t kAppId[]                           = L"RedSalamander";
constexpr wchar_t kMainWindowId[]                    = L"MainWindow";
constexpr wchar_t kPreferencesWindowId[]             = L"PreferencesWindow";
constexpr wchar_t kConnectionManagerWindowId[]       = L"ConnectionManagerWindow";
constexpr wchar_t kShortcutsWindowId[]               = L"ShortcutsWindow";
constexpr wchar_t kFindFilesWindowId[]               = L"FindFilesWindow";
constexpr wchar_t kItemPropertiesWindowId[]          = L"ItemPropertiesWindow";
constexpr wchar_t kItemPropertiesWindowClassName[]   = L"RedSalamander.ItemPropertiesWindow";
constexpr wchar_t kAboutDialogWindowClassName[]      = L"RedSalamander.AboutWindow";
constexpr wchar_t kFatalErrorDialogWindowClassName[] = L"RedSalamander.FatalErrorWindow";
constexpr wchar_t kExternalHelpUrl[]                 = L"https://github.com/RedSalamanders/RedSalamander/tree/main/Docs#readme";
constexpr wchar_t kLeftPaneSlot[]                    = L"left";
constexpr wchar_t kRightPaneSlot[]                   = L"right";
#ifdef ENABLE_TESTS
constexpr UINT kFatalErrorDialogDebugGetSnapshotMessage   = WM_APP + 0x71;
constexpr UINT kFatalErrorDialogDebugScrollByWheelMessage = WM_APP + 0x72;
constexpr double kSelfTestTimeoutMultiplierDefault        = 1.0;
constexpr double kSelfTestTimeoutMultiplierMin            = 0.1;
constexpr double kSelfTestTimeoutMultiplierMax            = 100.0;
std::mutex g_debugConnectionManagerConnectMutex;
bool g_debugConnectionManagerConnectSeen    = false;
uint8_t g_debugConnectionManagerConnectPane = 0u;
std::wstring g_debugConnectionManagerConnectName;
std::mutex g_debugRereadAssociationsMutex;
const Common::Settings::Settings* g_debugRereadAssociationsSettingsOverride = nullptr;
RereadAssociationsDebugSnapshot g_debugRereadAssociationsSnapshot{};
bool g_debugRereadAssociationsSnapshotValid = false;
#endif

[[nodiscard]] HWND NormalizeOwnedWindow(HWND ownerWindow) noexcept
{
    return (ownerWindow && IsWindow(ownerWindow) != FALSE) ? ownerWindow : nullptr;
}
AppTheme ResolveConfiguredTheme() noexcept;

void CenterWindowOnOwner(HWND window, HWND owner) noexcept
{
    if (! window || IsWindow(window) == FALSE || ! owner || IsWindow(owner) == FALSE)
    {
        return;
    }

    RECT ownerRect{};
    RECT windowRect{};
    if (GetWindowRect(owner, &ownerRect) == FALSE || GetWindowRect(window, &windowRect) == FALSE)
    {
        return;
    }

    const int x = ownerRect.left + (((ownerRect.right - ownerRect.left) - (windowRect.right - windowRect.left)) / 2);
    const int y = ownerRect.top + (((ownerRect.bottom - ownerRect.top) - (windowRect.bottom - windowRect.top)) / 2);
    SetWindowPos(window, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

#ifdef ENABLE_TESTS
struct SelfTestTimeoutMultiplierParseResult final
{
    bool valid   = false;
    double value = kSelfTestTimeoutMultiplierDefault;
};

[[nodiscard]] SelfTestTimeoutMultiplierParseResult ParseSelfTestTimeoutMultiplier(std::wstring_view value) noexcept
{
    std::wstring valueCopy(value);
    wchar_t* end        = nullptr;
    errno               = 0;
    const double parsed = wcstod(valueCopy.c_str(), &end);
    if (valueCopy.empty() || end == valueCopy.c_str() || (end && *end != L'\0') || errno != 0 || ! std::isfinite(parsed))
    {
        Debug::Error(L"Invalid --selftest-timeout-multiplier value '{}'. Expected a finite number in [{}, {}].",
                     valueCopy,
                     kSelfTestTimeoutMultiplierMin,
                     kSelfTestTimeoutMultiplierMax);
        return {};
    }

    const double clamped = std::clamp(parsed, kSelfTestTimeoutMultiplierMin, kSelfTestTimeoutMultiplierMax);
    if (clamped != parsed)
    {
        Debug::Warning(L"Clamped --selftest-timeout-multiplier from {} to {}. Supported range is [{}, {}].",
                       parsed,
                       clamped,
                       kSelfTestTimeoutMultiplierMin,
                       kSelfTestTimeoutMultiplierMax);
    }

    return {.valid = true, .value = clamped};
}

[[nodiscard]] std::wstring EscapeJsonString(std::wstring_view value)
{
    std::wstring escaped;
    escaped.reserve(value.size() + 2u);
    for (const wchar_t ch : value)
    {
        switch (ch)
        {
            case L'"': escaped.append(L"\\\""); break;
            case L'\\': escaped.append(L"\\\\"); break;
            case L'\b': escaped.append(L"\\b"); break;
            case L'\f': escaped.append(L"\\f"); break;
            case L'\n': escaped.append(L"\\n"); break;
            case L'\r': escaped.append(L"\\r"); break;
            case L'\t': escaped.append(L"\\t"); break;
            default:
                if (ch < 0x20)
                {
                    escaped.append(std::format(L"\\u{:04X}", static_cast<unsigned int>(ch)));
                }
                else
                {
                    escaped.push_back(ch);
                }
                break;
        }
    }
    return escaped;
}

void AppendJsonStringProperty(std::wstring& json, std::wstring_view name, std::wstring_view value)
{
    json.append(L"\"");
    json.append(name);
    json.append(L"\":\"");
    json.append(EscapeJsonString(value));
    json.append(L"\"");
}

void AppendSelfTestCaseListSuiteJson(std::wstring& json, std::wstring_view suiteName, const std::vector<std::wstring>& names)
{
    json.append(L"{");
    AppendJsonStringProperty(json, L"suite", suiteName);
    json.append(std::format(L",\"count\":{},\"cases\":[", names.size()));

    bool first = true;
    for (const std::wstring& name : names)
    {
        if (! first)
        {
            json.append(L",");
        }
        first = false;

        json.append(L"{");
        AppendJsonStringProperty(json, L"name", name);
        json.append(L"}");
    }

    json.append(L"]}");
}

[[nodiscard]] std::wstring BuildSelfTestCaseListJson(const SelfTest::SelfTestOptions& options, bool includeCompare, bool includeCommands, bool includeFileOps)
{
    struct SuiteCases final
    {
        std::wstring_view name;
        std::vector<std::wstring> cases;
    };

    std::vector<SuiteCases> suites;
    suites.reserve(3u);
    if (includeCompare)
    {
        suites.push_back({L"CompareDirectories", CompareDirectoriesSelfTest::ListCases(options)});
    }
    if (includeCommands)
    {
        suites.push_back({L"Commands", CommandsSelfTest::ListCases(options)});
    }
    if (includeFileOps)
    {
        suites.push_back({L"FileOperations", FileOperationsSelfTest::BuildExpectedCaseNames(options)});
    }

    size_t total = 0u;
    for (const SuiteCases& suite : suites)
    {
        total += suite.cases.size();
    }

    std::wstring json;
    json.reserve(8192u + (total * 96u));
    json.append(L"{\"version\":1");
    if (! options.caseFilter.empty())
    {
        json.append(L",");
        AppendJsonStringProperty(json, L"case_filter", options.caseFilter);
    }
    json.append(std::format(L",\"total\":{},\"suites\":[", total));

    bool first = true;
    for (const SuiteCases& suite : suites)
    {
        if (! first)
        {
            json.append(L",");
        }
        first = false;
        AppendSelfTestCaseListSuiteJson(json, suite.name, suite.cases);
    }

    json.append(L"]}\r\n");
    return json;
}

BOOL CALLBACK CountVisibleOwnedDialogChildWindowsProc(HWND hwnd, LPARAM lParam) noexcept
{
    auto* count = reinterpret_cast<size_t*>(lParam);
    if (! count)
    {
        return FALSE;
    }

    if (IsWindowVisible(hwnd) != FALSE)
    {
        ++(*count);
    }

    return TRUE;
}

[[nodiscard]] size_t CountVisibleOwnedDialogChildWindows(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    size_t count = 0u;
    EnumChildWindows(hwnd, CountVisibleOwnedDialogChildWindowsProc, reinterpret_cast<LPARAM>(&count));
    return count;
}

[[nodiscard]] uint32_t PackArgb(const D2D1_COLOR_F& color) noexcept
{
    const auto toByte = [](float value) noexcept -> uint32_t { return static_cast<uint32_t>(std::clamp(std::lround(value * 255.0f), 0l, 255l)); };

    return (toByte(color.a) << 24u) | (toByte(color.r) << 16u) | (toByte(color.g) << 8u) | toByte(color.b);
}
#endif

class AboutDialogWindow final
{
public:
    AboutDialogWindow(const AboutDialogWindow&)            = delete;
    AboutDialogWindow& operator=(const AboutDialogWindow&) = delete;
    AboutDialogWindow(AboutDialogWindow&&)                 = delete;
    AboutDialogWindow& operator=(AboutDialogWindow&&)      = delete;

    AboutDialogWindow(HWND ownerWindow, const AppTheme& theme) noexcept : _ownerWindow(NormalizeOwnedWindow(ownerWindow)), _theme(theme)
    {
    }

    [[nodiscard]] HRESULT ShowModal() noexcept
    {
        if (const HWND existing = FindWindowW(kAboutDialogWindowClassName, nullptr); existing && IsWindow(existing) != FALSE)
        {
            ShowWindow(existing, IsIconic(existing) ? SW_RESTORE : SW_SHOW);
            SetForegroundWindow(existing);
            return S_FALSE;
        }

        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return classHr;
        }

        const DWORD style        = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle      = WS_EX_DLGMODALFRAME;
        const UINT dpi           = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();
        const int clientWidthPx  = MulDiv(420, static_cast<int>(dpi == 0u ? 96u : dpi), 96);
        const int clientHeightPx = MulDiv(170, static_cast<int>(dpi == 0u ? 96u : dpi), 96);

        RECT bounds{0, 0, clientWidthPx, clientHeightPx};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        const std::wstring caption = LoadStringResource(nullptr, IDS_ABOUT_WINDOW_CAPTION);
        const HWND hwnd            = CreateWindowExW(exStyle,
                                                     kAboutDialogWindowClassName,
                                                     caption.c_str(),
                                                     style,
                                                     CW_USEDEFAULT,
                                                     CW_USEDEFAULT,
                                                     bounds.right - bounds.left,
                                                     bounds.bottom - bounds.top,
                                                     nullptr,
                                                     nullptr,
                                                     g_hInstance,
                                                     this);
        if (! hwnd)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        CenterWindowOnOwner(_hWnd.get(), _ownerWindow);
        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            if (getMessageResult == 0)
            {
                _done   = true;
                _result = S_FALSE;
                PostQuitMessage(static_cast<int>(msg.wParam));
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<AboutDialogWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
            return TRUE;
        }

        auto* self = reinterpret_cast<AboutDialogWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_NCDESTROY)
            {
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                if (! self->_done)
                {
                    self->_done   = true;
                    self->_result = S_FALSE;
                }
            }
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, self->_theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE: self->Close(); return 0;
            case WM_NCDESTROY:
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                if (! self->_done)
                {
                    self->_done   = true;
                    self->_result = S_FALSE;
                }
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = AboutDialogWindow::WndProc;
        wc.hInstance     = g_hInstance;
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kAboutDialogWindowClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

        BuildUi();
        ApplyTheme();
        Layout();
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_okButton);
        _dxHost.SetFocusControl(_okButton);
        return true;
    }

    void BuildUi() noexcept
    {
        if (_root)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _titleLabel = _root->AddChild<Label>(L"RedSalamander");
        _titleLabel->SetFontRole(FontRole::Header);
        _titleLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _versionLabel = _root->AddChild<Label>(std::wstring(REDSALAMANDER_TEXT_VERSION));
        _versionLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _copyrightLabel = _root->AddChild<Label>(VERSINFO_COPYRIGHT);
        _copyrightLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetMnemonic(L'O');
        _okButton->SetOnClick([this] { Close(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _dxHost.SetTheme(MakeAppThemeDxPalette(_theme));
        if (_hWnd)
        {
            ApplyTitleBarTheme(_hWnd.get(), _theme, GetActiveWindow() == _hWnd.get());
        }
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const D2D1_RECT_F bounds = _dxHost.GetClientBoundsDip();
        _root->SetBounds(bounds);

        const float left         = 24.0f;
        const float right        = bounds.right - 24.0f;
        const float contentWidth = (std::max)(0.0f, right - left);

        _titleLabel->SetBounds(D2D1::RectF(left, 22.0f, left + contentWidth, 50.0f));
        _versionLabel->SetBounds(D2D1::RectF(left, 62.0f, left + contentWidth, 84.0f));
        _copyrightLabel->SetBounds(D2D1::RectF(left, 88.0f, left + contentWidth, 110.0f));

        const float buttonWidth  = 92.0f;
        const float buttonHeight = 32.0f;
        const float buttonRight  = bounds.right - 24.0f;
        const float buttonTop    = bounds.bottom - 24.0f - buttonHeight;
        _okButton->SetBounds(D2D1::RectF(buttonRight - buttonWidth, buttonTop, buttonRight, buttonTop + buttonHeight));
    }

    void Close() noexcept
    {
        if (_done)
        {
            return;
        }

        _done   = true;
        _result = S_OK;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            DestroyWindow(_hWnd.get());
        }
    }

private:
    HWND _ownerWindow = nullptr;
    AppTheme _theme{};
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root           = nullptr;
    RedSalamander::DxUi::Label* _titleLabel     = nullptr;
    RedSalamander::DxUi::Label* _versionLabel   = nullptr;
    RedSalamander::DxUi::Label* _copyrightLabel = nullptr;
    RedSalamander::DxUi::Button* _okButton      = nullptr;
    bool _done                                  = false;
    HRESULT _result                             = S_FALSE;
};

[[nodiscard]] HRESULT ShowAboutDialog(HWND ownerWindow, const AppTheme& theme) noexcept
{
    AboutDialogWindow dialog(ownerWindow, theme);
    return dialog.ShowModal();
}

class FatalErrorDialogWindow final
{
public:
    FatalErrorDialogWindow(HWND ownerWindow, const AppTheme& theme, std::wstring caption, std::wstring message) noexcept
        : _ownerWindow(NormalizeOwnedWindow(ownerWindow)),
          _theme(theme),
          _caption(std::move(caption)),
          _message(std::move(message))
    {
    }

    FatalErrorDialogWindow(const FatalErrorDialogWindow&)            = delete;
    FatalErrorDialogWindow& operator=(const FatalErrorDialogWindow&) = delete;
    FatalErrorDialogWindow(FatalErrorDialogWindow&&)                 = delete;
    FatalErrorDialogWindow& operator=(FatalErrorDialogWindow&&)      = delete;

    [[nodiscard]] HRESULT ShowModal() noexcept
    {
        const HRESULT classHr = EnsureWindowClass();
        if (FAILED(classHr))
        {
            return classHr;
        }

        const DWORD style        = WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
        const DWORD exStyle      = WS_EX_DLGMODALFRAME;
        const UINT dpi           = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : GetDpiForSystem();
        const int clientWidthPx  = MulDiv(480, static_cast<int>(dpi == 0u ? 96u : dpi), 96);
        const int clientHeightPx = MulDiv(220, static_cast<int>(dpi == 0u ? 96u : dpi), 96);

        RECT bounds{0, 0, clientWidthPx, clientHeightPx};
        if (AdjustWindowRectExForDpi(&bounds, style, FALSE, exStyle, dpi) == FALSE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        const bool restoreOwnerEnabled = _ownerWindow && IsWindow(_ownerWindow) != FALSE && IsWindowEnabled(_ownerWindow) != FALSE;
        if (restoreOwnerEnabled)
        {
            EnableWindow(_ownerWindow, FALSE);
        }
        const auto restoreOwner = wil::scope_exit([this, restoreOwnerEnabled]() noexcept
        {
            if (restoreOwnerEnabled && _ownerWindow && IsWindow(_ownerWindow) != FALSE)
            {
                EnableWindow(_ownerWindow, TRUE);
                SetActiveWindow(_ownerWindow);
            }
        });

        const std::wstring caption = _caption.empty() ? LoadStringResource(nullptr, IDS_APP_TITLE) : _caption;
        const HWND hwnd            = CreateWindowExW(exStyle,
                                                     kFatalErrorDialogWindowClassName,
                                                     caption.c_str(),
                                                     style,
                                                     CW_USEDEFAULT,
                                                     CW_USEDEFAULT,
                                                     bounds.right - bounds.left,
                                                     bounds.bottom - bounds.top,
                                                     _ownerWindow,
                                                     nullptr,
                                                     g_hInstance,
                                                     this);
        if (! hwnd)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (! _hWnd)
        {
            _hWnd.reset(hwnd);
        }

        CenterWindowOnOwner(_hWnd.get(), _ownerWindow);
        static_cast<void>(_dxHost.PrimeForShow());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        UpdateWindow(_hWnd.get());
        SetForegroundWindow(_hWnd.get());

        MSG msg{};
        while (! _done)
        {
            const BOOL getMessageResult = GetMessageW(&msg, nullptr, 0, 0);
            if (getMessageResult == -1)
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            if (getMessageResult == 0)
            {
                _done   = true;
                _result = S_FALSE;
                PostQuitMessage(static_cast<int>(msg.wParam));
                break;
            }

            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return _result;
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        if (message == WM_NCCREATE)
        {
            auto* cs   = reinterpret_cast<CREATESTRUCTW*>(lParam);
            auto* self = static_cast<FatalErrorDialogWindow*>(cs ? cs->lpCreateParams : nullptr);
            if (! self)
            {
                return FALSE;
            }

            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (! self->_hWnd)
            {
                self->_hWnd.reset(hwnd);
            }
        }

        auto* self = reinterpret_cast<FatalErrorDialogWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (! self)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        bool handled     = false;
        LRESULT dxResult = 0;
        if (message != WM_CREATE)
        {
            dxResult = self->_dxHost.HandleMessage(hwnd, message, wParam, lParam, handled);
        }
        if (handled)
        {
            if (message == WM_NCDESTROY)
            {
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                if (! self->_done)
                {
                    self->_done   = true;
                    self->_result = S_FALSE;
                }
            }
            if (message == WM_SIZE || message == WM_DPICHANGED)
            {
                self->Layout();
            }
            return dxResult;
        }

        switch (message)
        {
            case WM_CREATE: return self->OnCreate(hwnd) ? 0 : -1;
#ifdef ENABLE_TESTS
            case kFatalErrorDialogDebugGetSnapshotMessage:
            {
                auto* snapshot = reinterpret_cast<FatalErrorDialogDebugSnapshot*>(lParam);
                return (snapshot && self->DebugGetSnapshot(*snapshot)) ? TRUE : FALSE;
            }
            case kFatalErrorDialogDebugScrollByWheelMessage: return self->DebugScrollByWheelDetents(static_cast<int>(wParam)) ? TRUE : FALSE;
#endif
            case WM_SIZE: self->Layout(); return 0;
            case WM_DPICHANGED:
            {
                const auto* suggested = reinterpret_cast<const RECT*>(lParam);
                if (suggested)
                {
                    SetWindowPos(hwnd,
                                 nullptr,
                                 suggested->left,
                                 suggested->top,
                                 suggested->right - suggested->left,
                                 suggested->bottom - suggested->top,
                                 SWP_NOZORDER | SWP_NOACTIVATE);
                }
                self->Layout();
                return 0;
            }
            case WM_ERASEBKGND: return 1;
            case WM_NCACTIVATE: ApplyTitleBarTheme(hwnd, self->_theme, wParam != FALSE); return DefWindowProcW(hwnd, message, wParam, lParam);
            case WM_CLOSE: self->Close(); return 0;
            case WM_NCDESTROY:
                self->_dxHost.Detach();
                if (self->_hWnd.get() == hwnd)
                {
                    self->_hWnd.release();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                if (! self->_done)
                {
                    self->_done   = true;
                    self->_result = S_FALSE;
                }
                break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

private:
    [[nodiscard]] static HRESULT EnsureWindowClass() noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return S_OK;
        }

        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = FatalErrorDialogWindow::WndProc;
        wc.hInstance     = g_hInstance;
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kFatalErrorDialogWindowClassName;
        wc.style         = CS_DBLCLKS;

        atom = RegisterClassExW(&wc);
        return atom != 0 ? S_OK : HRESULT_FROM_WIN32(GetLastError());
    }

    [[nodiscard]] bool OnCreate(HWND hwnd) noexcept
    {
        if (! _dxHost.Attach(hwnd))
        {
            return false;
        }

        BuildUi();
        ApplyTheme();
        Layout();
        _dxHost.SetDefaultButton(_okButton);
        _dxHost.SetCancelButton(_okButton);
        _dxHost.SetFocusControl(_okButton);
        return true;
    }

    void BuildUi()
    {
        if (_root)
        {
            return;
        }

        using namespace RedSalamander::DxUi;

        _rootStorage = std::make_unique<Panel>();
        _root        = _rootStorage.get();

        _messageField = _root->AddChild<TextField>(_message);
        _messageField->SetMultiline(true);
        _messageField->SetReadOnly(true);

        _okButton = _root->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_OK));
        _okButton->SetPrimary(true);
        _okButton->SetMnemonic(L'O');
        _okButton->SetOnClick([this] { Close(); });

        _dxHost.SetRoot(std::move(_rootStorage));
    }

    void ApplyTheme() noexcept
    {
        _dxHost.SetTheme(MakeAppThemeDxPalette(_theme));
        if (_hWnd)
        {
            ApplyTitleBarTheme(_hWnd.get(), _theme, GetActiveWindow() == _hWnd.get());
        }
    }

    void Layout() noexcept
    {
        if (! _root)
        {
            return;
        }

        const D2D1_RECT_F bounds = _dxHost.GetClientBoundsDip();
        _root->SetBounds(bounds);

        const float outerPadding  = 16.0f;
        const float buttonWidth   = 92.0f;
        const float buttonHeight  = 32.0f;
        const float buttonBottom  = bounds.bottom - outerPadding;
        const float buttonTop     = buttonBottom - buttonHeight;
        const float buttonRight   = bounds.right - outerPadding;
        const float contentTop    = outerPadding;
        const float contentBottom = (std::max)(contentTop, buttonTop - 12.0f);

        _messageField->SetBounds(D2D1::RectF(outerPadding, contentTop, bounds.right - outerPadding, contentBottom));
        _okButton->SetBounds(D2D1::RectF(buttonRight - buttonWidth, buttonTop, buttonRight, buttonBottom));
    }

#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugGetSnapshot(FatalErrorDialogDebugSnapshot& out) const noexcept
    {
        const HWND hwnd             = _hWnd.get();
        out.usesDxUiHost            = hwnd && ::IsWindow(hwnd) != FALSE;
        out.visibleChildWindowCount = hwnd ? CountVisibleOwnedDialogChildWindows(hwnd) : 0u;
        out.themeDark               = _theme.dark;
        out.themeHighContrast       = _theme.highContrast;
        out.themeRainbow            = _theme.menu.rainbowMode;
        RedSalamander::DxUi::TextFieldDebugMultilineState multilineState{};
        if (_messageField && _messageField->DebugGetMultilineState(_dxHost, multilineState))
        {
            out.bodyFirstVisibleLine    = multilineState.firstVisibleLine;
            out.bodyVisibleLineCount    = multilineState.visibleLineCount;
            out.bodyTotalLineCount      = multilineState.totalLineCount;
            out.bodyCanScrollVertically = multilineState.canScrollVertically;
        }
        if (_messageField)
        {
            const auto bodyStyle = RedSalamander::DxUi::ResolveTextFieldVisualStyle(_dxHost.GetTheme(),
                                                                                    _messageField->IsEnabled(),
                                                                                    _messageField->IsHovered(),
                                                                                    _messageField->HasFocus(),
                                                                                    _messageField->HasFocus() && _dxHost.IsKeyboardFocusVisible());
            out.bodyFillArgb     = PackArgb(bodyStyle.fill);
            out.bodyTextArgb     = PackArgb(bodyStyle.text);
        }
        out.renderCount        = _dxHost.DebugGetRenderCount();
        out.resizeCount        = _dxHost.DebugGetResizeCount();
        out.resizeFailureCount = _dxHost.DebugGetResizeFailureCount();
        out.messageText        = _message;
        return out.usesDxUiHost;
    }

    [[nodiscard]] bool DebugScrollByWheelDetents(int detents) noexcept
    {
        const HWND hwnd = _hWnd.get();
        if (! hwnd || ::IsWindow(hwnd) == FALSE || ! _messageField)
        {
            return false;
        }

        if (::IsIconic(hwnd))
        {
            ::ShowWindow(hwnd, SW_RESTORE);
        }

        const float wheelDelta = detents > 0 ? static_cast<float>(WHEEL_DELTA) : -static_cast<float>(WHEEL_DELTA);
        const int stepCount    = detents > 0 ? detents : -detents;
        for (int remaining = stepCount; remaining > 0; --remaining)
        {
            if (! _messageField->OnMouseWheel(_dxHost, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u))
            {
                return false;
            }
        }
        return true;
    }
#endif

    void Close() noexcept
    {
        if (_done)
        {
            return;
        }

        _done   = true;
        _result = S_OK;
        if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
        {
            DestroyWindow(_hWnd.get());
        }
    }

private:
    HWND _ownerWindow = nullptr;
    AppTheme _theme{};
    std::wstring _caption;
    std::wstring _message;
    wil::unique_hwnd _hWnd;
    RedSalamander::DxUi::WindowHost _dxHost;
    std::unique_ptr<RedSalamander::DxUi::Panel> _rootStorage;
    RedSalamander::DxUi::Panel* _root             = nullptr;
    RedSalamander::DxUi::TextField* _messageField = nullptr;
    RedSalamander::DxUi::Button* _okButton        = nullptr;
    bool _done                                    = false;
    HRESULT _result                               = S_FALSE;
};

#ifdef ENABLE_TESTS
bool g_runFileOpsSelfTest            = false;
bool g_runCompareDirectoriesSelfTest = false;
bool g_runCommandsSelfTest           = false;
std::atomic<DWORD> g_selfTestMonitorProcessId{0};

struct SelfTestMonitorCloseContext
{
    DWORD processId    = 0;
    bool foundWindow   = false;
    bool closedWithMsg = false;
};

BOOL CALLBACK SelfTestMonitorWindowEnumProc(HWND hwnd, LPARAM lParam) noexcept
{
    auto* context = reinterpret_cast<SelfTestMonitorCloseContext*>(lParam);
    if (! context || hwnd == nullptr)
    {
        return TRUE;
    }

    DWORD windowPid = 0;
    GetWindowThreadProcessId(hwnd, &windowPid);
    if (windowPid != context->processId)
    {
        return TRUE;
    }

    context->foundWindow = true;
    if (PostMessageW(hwnd, WM_CLOSE, 0, 0))
    {
        context->closedWithMsg = true;
    }

    return TRUE;
}

void ShutdownSelfTestMonitor() noexcept
{
    if (! (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest))
    {
        return;
    }

    const DWORD monitorPid = g_selfTestMonitorProcessId.load(std::memory_order_acquire);
    if (monitorPid == 0)
    {
        return;
    }

    SelfTestMonitorCloseContext context{};
    context.processId = monitorPid;
    EnumWindows(SelfTestMonitorWindowEnumProc, reinterpret_cast<LPARAM>(&context));

    if (! context.closedWithMsg && ! context.foundWindow)
    {
        Debug::Info(L"SelfTest: monitor PID {} not found in window enumeration", monitorPid);
    }

    const wil::unique_handle monitorProcess{::OpenProcess(SYNCHRONIZE | PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION, FALSE, monitorPid)};
    if (! monitorProcess)
    {
        g_selfTestMonitorProcessId.store(0, std::memory_order_release);
        return;
    }

    if (WaitForSingleObject(monitorProcess.get(), 2000) == WAIT_TIMEOUT)
    {
        TerminateProcess(monitorProcess.get(), 0);
        WaitForSingleObject(monitorProcess.get(), 2000);
    }

    g_selfTestMonitorProcessId.store(0, std::memory_order_release);
}

[[nodiscard]] bool HasAnySelfTestArgInCommandLine() noexcept
{
    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
    if (! argv || argc <= 1)
    {
        return false;
    }

    constexpr std::wstring_view kSelfTestArgs[] = {
        L"--selftest",
        L"--compare-selftest",
        L"--commands-selftest",
        L"--fileops-selftest",
    };

    for (int i = 1; i < argc; ++i)
    {
        const wchar_t* arg = argv.get()[i];
        if (! arg || arg[0] == L'\0')
        {
            continue;
        }

        const std::wstring_view argView(arg);
        for (std::wstring_view needle : kSelfTestArgs)
        {
            if (OrdinalString::EqualsNoCase(argView, needle))
            {
                return true;
            }
        }
    }

    return false;
}

void QueueRedSalamanderMonitorLaunch() noexcept
{
    if (HasAnySelfTestArgInCommandLine())
    {
        return;
    }

    // Best-effort: launch the ETW viewer early in debug builds so startup ETW events are visible.
    // RedSalamanderMonitor has its own single-instance mutex, so extra launches will exit quickly.
    constexpr wchar_t kInstanceMutexName[] = L"Local\\RedSalamanderMonitor_Instance";
    wil::unique_handle existingInstance(::OpenMutexW(SYNCHRONIZE, FALSE, kInstanceMutexName));
    if (existingInstance)
    {
        return;
    }

    static_cast<void>(TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* /*context*/) noexcept
    {
        constexpr wchar_t kInstanceMutexName[] = L"Local\\RedSalamanderMonitor_Instance";
        wil::unique_handle existingInstance(::OpenMutexW(SYNCHRONIZE, FALSE, kInstanceMutexName));
        if (existingInstance)
        {
            return;
        }

        wchar_t exePath[MAX_PATH]{};
        const DWORD exeLen = ::GetModuleFileNameW(nullptr, exePath, static_cast<DWORD>(std::size(exePath)));
        if (exeLen == 0 || exeLen >= std::size(exePath))
        {
            return;
        }

        wchar_t* lastSlash = wcsrchr(exePath, L'\\');
        if (! lastSlash)
        {
            lastSlash = wcsrchr(exePath, L'/');
        }
        if (! lastSlash)
        {
            return;
        }
        *(lastSlash + 1) = L'\0';

        wchar_t monitorPath[MAX_PATH]{};
        if (wcscpy_s(monitorPath, std::size(monitorPath), exePath) != 0)
        {
            return;
        }
        if (wcscat_s(monitorPath, std::size(monitorPath), L"RedSalamanderMonitor.exe") != 0)
        {
            return;
        }

        if (::GetFileAttributesW(monitorPath) == INVALID_FILE_ATTRIBUTES)
        {
            return;
        }

        wchar_t cmdLine[(MAX_PATH * 2) + 4]{};
        if (wcscpy_s(cmdLine, std::size(cmdLine), L"\"") != 0)
        {
            return;
        }
        if (wcscat_s(cmdLine, std::size(cmdLine), monitorPath) != 0)
        {
            return;
        }
        if (wcscat_s(cmdLine, std::size(cmdLine), L"\"") != 0)
        {
            return;
        }

        STARTUPINFOW si{};
        si.cb = sizeof(si);

        PROCESS_INFORMATION pi{};
        if (! ::CreateProcessW(monitorPath, cmdLine, nullptr, nullptr, FALSE, 0, nullptr, nullptr, &si, &pi))
        {
            return;
        }

        wil::unique_handle process(pi.hProcess);
        wil::unique_handle thread(pi.hThread);
        g_selfTestMonitorProcessId.store(pi.dwProcessId, std::memory_order_release);
    },
        nullptr,
        nullptr));
}
#endif // ENABLE_TESTS

[[nodiscard]] bool IsRunningAnySelfTest() noexcept
{
#ifdef ENABLE_TESTS
    return g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest;
#else
    return false;
#endif
}

// Known folder GUID for OneDrive root (aka "SkyDrive").
constexpr GUID kKnownFolderIdOneDrive = {
    0xA52BBA46,
    0xE9E1,
    0x435F,
    {0xB3, 0xD9, 0x28, 0xDA, 0xA6, 0x48, 0xC0, 0xF6},
};

void ShowFatalErrorDialog(HWND owner, const wchar_t* caption, const wchar_t* message) noexcept
{
    const AppTheme theme = ResolveConfiguredTheme();
    FatalErrorDialogWindow dialog(owner, theme, caption ? std::wstring(caption) : std::wstring(), message ? std::wstring(message) : std::wstring());
    static_cast<void>(dialog.ShowModal());
}

HMENU g_mainMenuHandle       = nullptr;
HMENU g_filesMenu            = nullptr;
HMENU g_viewMenu             = nullptr;
HMENU g_editMenu             = nullptr;
HMENU g_editAdvancedMenu     = nullptr;
HMENU g_viewThemeMenu        = nullptr;
HMENU g_viewPluginsMenu      = nullptr;
HMENU g_viewWithMenu         = nullptr;
HMENU g_editWithMenu         = nullptr;
HMENU g_newTemplateMenu      = nullptr;
HMENU g_userMenu             = nullptr;
HMENU g_openFileExplorerMenu = nullptr;

std::unordered_map<UINT, std::wstring> g_viewWithMenuIdToActionId;
std::unordered_map<UINT, std::wstring> g_editWithMenuIdToActionId;
std::unordered_map<UINT, std::wstring> g_newTemplateMenuIdToTemplateId;
std::unordered_map<UINT, std::wstring> g_userMenuIdToActionId;

bool g_menuBarVisible          = true;
bool g_menuBarTemporarilyShown = false;
bool g_functionBarVisible      = true;

struct FullScreenState
{
    bool active        = false;
    DWORD savedStyle   = 0;
    DWORD savedExStyle = 0;
    WINDOWPLACEMENT savedPlacement{};
};
FullScreenState g_fullScreenState{};

void ToggleFullScreen(HWND hWnd) noexcept;

constexpr UINT_PTR kFunctionBarPressedKeyClearTimerId = 1001u;
constexpr UINT kFunctionBarPressedKeyClearDelayMs     = 200u;
std::optional<uint32_t> g_functionBarPressedKey;
std::optional<uint32_t> g_functionBarPressedKeyClearPending;

#ifdef ENABLE_TESTS
#include "SelfTestCommon.h"

constexpr UINT_PTR kFileOpsSelfTestTimerId     = 1002u;
constexpr UINT kFileOpsSelfTestTimerIntervalMs = 50u;
constexpr UINT kCommandsSelfTestStartMessage   = WM_APP + 0x6A;
constexpr wchar_t kSelfTestRunMutexName[]      = L"Local\\RedSalamander_SelfTestRun";
int g_selfTestExitCode                         = 0;
SelfTest::SelfTestOptions g_selfTestOptions{};
SelfTest::SelfTestRunResult g_selfTestRunResult{};
std::optional<std::chrono::time_point<std::chrono::steady_clock>> g_selfTestRunStart{};
bool g_selfTestRunFinalized = false;
wil::unique_handle g_selfTestRunMutex;
std::vector<std::wstring> g_fileOpsSelfTestRunFilters;
std::vector<std::wstring> g_fileOpsSelfTestExpectedCases;
size_t g_fileOpsSelfTestRunIndex = 0;
SelfTest::SelfTestSuiteResult g_fileOpsSelfTestAggregateResult{};
bool g_fileOpsSelfTestStartPending = false;

[[nodiscard]] bool TryAcquireSelfTestRunMutex() noexcept
{
    if (g_selfTestRunMutex)
    {
        return true;
    }

    SetLastError(ERROR_SUCCESS);
    wil::unique_handle mutex(::CreateMutexW(nullptr, FALSE, kSelfTestRunMutexName));
    if (! mutex)
    {
        Debug::ErrorWithLastError(L"CreateMutexW failed for the self-test run guard.");
        return false;
    }

    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        Debug::Error(L"Another RedSalamander self-test run is already active. Refusing to start a parallel self-test instance.");
        return false;
    }

    g_selfTestRunMutex = std::move(mutex);
    return true;
}

[[nodiscard]] std::wstring GetSelfTestUtcIso8601() noexcept
{
    const std::chrono::system_clock::time_point now = std::chrono::system_clock::now();
    const std::time_t nowUtc                        = std::chrono::system_clock::to_time_t(now);
    tm utc{};
    if (gmtime_s(&utc, &nowUtc) != 0)
    {
        return {};
    }

    const auto nowMs      = std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch());
    const auto millisPart = nowMs.count() % 1000;

    return std::format(
        L"{0:04}-{1:02}-{2:02}T{3:02}:{4:02}:{5:02}.{6:03}Z", utc.tm_year + 1900, utc.tm_mon + 1, utc.tm_mday, utc.tm_hour, utc.tm_min, utc.tm_sec, millisPart);
}

[[nodiscard]] std::wstring_view GetSelfTestArchiveArea() noexcept
{
    const int count = (g_runFileOpsSelfTest ? 1 : 0) + (g_runCompareDirectoriesSelfTest ? 1 : 0) + (g_runCommandsSelfTest ? 1 : 0);
    if (count != 1)
    {
        return L"SelfTest";
    }

    if (g_runFileOpsSelfTest)
    {
        return L"FileOps";
    }
    if (g_runCompareDirectoriesSelfTest)
    {
        return L"CompareDirectories";
    }
    if (g_runCommandsSelfTest)
    {
        return L"Commands";
    }

    return L"SelfTest";
}

void FinalizeSelfTestRun() noexcept
{
    if (g_selfTestRunFinalized || ! g_selfTestRunStart.has_value())
    {
        return;
    }

    ShutdownSelfTestMonitor();

    const auto now                   = std::chrono::steady_clock::now();
    g_selfTestRunResult.durationMs   = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(now - g_selfTestRunStart.value()).count());
    g_selfTestRunResult.failFast     = g_selfTestOptions.failFast;
    g_selfTestRunResult.timeoutScale = g_selfTestOptions.timeoutScale;
    g_selfTestRunResult.caseFilter   = g_selfTestOptions.caseFilter;

    const std::filesystem::path runJsonPath = SelfTest::SelfTestRoot() / L"last_run" / L"results.json";
    SelfTest::WriteRunJson(g_selfTestRunResult, runJsonPath);
    SelfTest::TryArchiveLastRunToRepo(GetSelfTestArchiveArea(), g_selfTestExitCode, g_selfTestRunResult.durationMs, &g_selfTestRunResult);
    g_selfTestRunMutex.reset();
    g_selfTestRunFinalized = true;
}

void TraceSelfTestExitCode(std::wstring_view source, int exitCode) noexcept
{
    SelfTest::AppendSelfTestTrace(std::format(L"{}: exit_code={}", source, exitCode));
}

void MergeFileOpsCase(SelfTest::SelfTestCaseResult& target, const SelfTest::SelfTestCaseResult& source) noexcept
{
    target.durationMs += source.durationMs;

    if (target.status == SelfTest::SelfTestCaseResult::Status::failed || source.status == SelfTest::SelfTestCaseResult::Status::failed)
    {
        target.status = SelfTest::SelfTestCaseResult::Status::failed;
        if (source.status == SelfTest::SelfTestCaseResult::Status::failed && ! source.reason.empty())
        {
            target.reason = source.reason;
        }
        else if (target.reason.empty())
        {
            target.reason = source.reason;
        }
        return;
    }

    if (target.status == SelfTest::SelfTestCaseResult::Status::passed || source.status == SelfTest::SelfTestCaseResult::Status::passed)
    {
        target.status = SelfTest::SelfTestCaseResult::Status::passed;
        target.reason.clear();
        return;
    }

    target.status = SelfTest::SelfTestCaseResult::Status::skipped;
    if (source.status == SelfTest::SelfTestCaseResult::Status::skipped && ! source.reason.empty())
    {
        target.reason = source.reason;
    }
}

void MergeFileOpsSuite(SelfTest::SelfTestSuiteResult& aggregate, const SelfTest::SelfTestSuiteResult& current) noexcept
{
    aggregate.suite = SelfTest::SelfTestSuite::FileOperations;
    aggregate.durationMs += current.durationMs;
    if (aggregate.failureMessage.empty() && ! current.failureMessage.empty())
    {
        aggregate.failureMessage = current.failureMessage;
    }

    for (const auto& item : current.cases)
    {
        const auto it = std::find_if(
            aggregate.cases.begin(), aggregate.cases.end(), [&](const SelfTest::SelfTestCaseResult& existing) noexcept { return existing.name == item.name; });
        if (it == aggregate.cases.end())
        {
            aggregate.cases.push_back(item);
            continue;
        }

        MergeFileOpsCase(*it, item);
    }
}

void FinalizeFileOpsAggregateResult(SelfTest::SelfTestSuiteResult& aggregate, const std::vector<std::wstring>& expectedCases, bool stoppedEarly) noexcept
{
    std::vector<SelfTest::SelfTestCaseResult> orderedCases;
    orderedCases.reserve(expectedCases.size());

    for (const std::wstring& expectedName : expectedCases)
    {
        const auto it = std::find_if(
            aggregate.cases.begin(), aggregate.cases.end(), [&](const SelfTest::SelfTestCaseResult& item) noexcept { return item.name == expectedName; });
        if (it != aggregate.cases.end())
        {
            orderedCases.push_back(*it);
            continue;
        }

        SelfTest::SelfTestCaseResult skipped{};
        skipped.name       = expectedName;
        skipped.status     = SelfTest::SelfTestCaseResult::Status::skipped;
        skipped.durationMs = 0;
        skipped.reason     = stoppedEarly ? L"not executed (suite stopped early)" : L"not executed";
        orderedCases.push_back(std::move(skipped));
    }

    aggregate.cases   = std::move(orderedCases);
    aggregate.passed  = 0;
    aggregate.failed  = 0;
    aggregate.skipped = 0;
    aggregate.failureMessage.clear();

    for (const auto& item : aggregate.cases)
    {
        switch (item.status)
        {
            case SelfTest::SelfTestCaseResult::Status::passed: ++aggregate.passed; break;
            case SelfTest::SelfTestCaseResult::Status::failed:
                ++aggregate.failed;
                if (aggregate.failureMessage.empty() && ! item.reason.empty())
                {
                    aggregate.failureMessage = item.reason;
                }
                break;
            case SelfTest::SelfTestCaseResult::Status::skipped: ++aggregate.skipped; break;
        }
    }
}

SelfTest::SelfTestOptions MakeFileOpsRunOptions(std::wstring_view runFilter)
{
    SelfTest::SelfTestOptions options = g_selfTestOptions;
    options.caseFilter                = std::wstring(runFilter);
    return options;
}

void StartNextFileOpsSelfTestRun(HWND hWnd) noexcept
{
    if (g_fileOpsSelfTestRunIndex >= g_fileOpsSelfTestRunFilters.size())
    {
        return;
    }

    const std::wstring& runFilter = g_fileOpsSelfTestRunFilters[g_fileOpsSelfTestRunIndex];
    SelfTest::AppendSuiteTrace(
        SelfTest::SelfTestSuite::FileOperations,
        std::format(L"FileOpsSelfTest: family {}/{} -> {}", g_fileOpsSelfTestRunIndex + 1, g_fileOpsSelfTestRunFilters.size(), runFilter));
    SelfTest::AppendSelfTestTrace(
        std::format(L"FileOpsSelfTest: family {}/{} -> {}", g_fileOpsSelfTestRunIndex + 1, g_fileOpsSelfTestRunFilters.size(), runFilter));
    FileOperationsSelfTest::Start(hWnd, MakeFileOpsRunOptions(runFilter));
    ++g_fileOpsSelfTestRunIndex;
}

void ResetSelfTestRunState() noexcept
{
    g_selfTestExitCode                = 0;
    g_selfTestRunFinalized            = false;
    g_selfTestRunStart                = std::chrono::steady_clock::now();
    g_selfTestRunResult               = {};
    g_selfTestRunResult.startedUtcIso = GetSelfTestUtcIso8601();
    SelfTest::SetRunStartedUtcIso(g_selfTestRunResult.startedUtcIso);
    g_selfTestRunResult.failFast     = g_selfTestOptions.failFast;
    g_selfTestRunResult.timeoutScale = g_selfTestOptions.timeoutScale;
    g_selfTestRunResult.caseFilter   = g_selfTestOptions.caseFilter;
    g_fileOpsSelfTestRunFilters.clear();
    g_fileOpsSelfTestExpectedCases.clear();
    g_fileOpsSelfTestRunIndex        = 0;
    g_fileOpsSelfTestAggregateResult = {};
    g_fileOpsSelfTestStartPending    = false;
}

void RecordSelfTestSuite(SelfTest::SelfTestSuiteResult result) noexcept
{
    g_selfTestRunResult.suites.push_back(std::move(result));
}

LRESULT RunCommandsSelfTestAndRequestShutdown(HWND hWnd) noexcept
{
    SplashScreen::IfExistSetText(L"Launching commands-selftest...");
    SelfTest::SelfTestSuiteResult commandsResult;
    Debug::Info(L"CommandsSelfTest: running");
    SelfTest::InitSelfTestRun(g_selfTestOptions);
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"CommandsSelfTest: begin");
    SelfTest::AppendSelfTestTrace(L"CommandsSelfTest: begin");
    SelfTest::AppendSelfTestTrace(L"CommandsSelfTest: calling CommandsSelfTest::Run");
    g_selfTestExitCode |= CommandsSelfTest::Run(hWnd, g_selfTestOptions, &commandsResult) ? 0 : 1;
    RecordSelfTestSuite(commandsResult);
    if (g_selfTestExitCode != 0)
    {
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"CommandsSelfTest: FAIL");
        SelfTest::AppendSelfTestTrace(L"CommandsSelfTest: FAIL");
        if (! commandsResult.failureMessage.empty())
        {
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, commandsResult.failureMessage);
            SelfTest::AppendSelfTestTrace(commandsResult.failureMessage);
        }
    }
    else
    {
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, L"CommandsSelfTest: PASS");
        SelfTest::AppendSelfTestTrace(L"CommandsSelfTest: PASS");
    }
    TraceSelfTestExitCode(L"CommandsSelfTest: end", g_selfTestExitCode);
    SplashScreen::CloseIfExist();
    if (PostMessageW(hWnd, WM_CLOSE, 0, 0) == 0)
    {
        FinalizeSelfTestRun();
        PostQuitMessage(g_selfTestExitCode);
    }
    return 0;
}
#endif

HMENU g_leftPaneMenu    = nullptr;
HMENU g_leftSortMenu    = nullptr;
HMENU g_leftDisplayMenu = nullptr;
HMENU g_leftGoToMenu    = nullptr;
HMENU g_leftShowMenu    = nullptr;

HMENU g_rightPaneMenu    = nullptr;
HMENU g_rightSortMenu    = nullptr;
HMENU g_rightDisplayMenu = nullptr;
HMENU g_rightGoToMenu    = nullptr;
HMENU g_rightShowMenu    = nullptr;

constexpr UINT kHistoryMenuMaxItems = 50u;

static void ResetMenuHandleCache() noexcept
{
    g_mainMenuHandle       = nullptr;
    g_filesMenu            = nullptr;
    g_viewMenu             = nullptr;
    g_editMenu             = nullptr;
    g_editAdvancedMenu     = nullptr;
    g_viewThemeMenu        = nullptr;
    g_viewPluginsMenu      = nullptr;
    g_viewWithMenu         = nullptr;
    g_editWithMenu         = nullptr;
    g_newTemplateMenu      = nullptr;
    g_userMenu             = nullptr;
    g_openFileExplorerMenu = nullptr;
    g_leftPaneMenu         = nullptr;
    g_leftSortMenu         = nullptr;
    g_leftDisplayMenu      = nullptr;
    g_leftGoToMenu         = nullptr;
    g_leftShowMenu         = nullptr;
    g_rightPaneMenu        = nullptr;
    g_rightSortMenu        = nullptr;
    g_rightDisplayMenu     = nullptr;
    g_rightGoToMenu        = nullptr;
    g_rightShowMenu        = nullptr;
    g_viewWithMenuIdToActionId.clear();
    g_editWithMenuIdToActionId.clear();
    g_newTemplateMenuIdToTemplateId.clear();
    g_userMenuIdToActionId.clear();
}

struct NavigatePathMenuTarget
{
    FolderWindow::Pane pane = FolderWindow::Pane::Left;
    std::filesystem::path path;
};

std::unordered_map<UINT, NavigatePathMenuTarget> g_navigatePathMenuTargets;

constexpr UINT kCustomThemeMenuIdFirst = 32800u;
constexpr UINT kCustomThemeMenuIdLast  = 32999u;

std::unordered_map<UINT, std::wstring> g_customThemeMenuIdToThemeId;
std::unordered_map<std::wstring, UINT> g_customThemeIdToMenuId;
std::vector<Common::Settings::ThemeDefinition> g_fileThemes;

constexpr UINT kPluginMenuIdFirst = 33500u;
constexpr UINT kPluginMenuIdLast  = 33699u;
std::unordered_map<UINT, std::wstring> g_pluginMenuIdToPluginId;
std::unordered_map<std::wstring, UINT> g_pluginIdToMenuId;

std::unordered_map<UINT, wil::unique_hbitmap> g_mainMenuIconBitmaps;

std::filesystem::path GetThemesDirectory()
{
    wil::unique_cotaskmem_string modulePath;
    const HRESULT hr = wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, modulePath);
    if (FAILED(hr) || ! modulePath)
    {
        return {};
    }
    return std::filesystem::path(modulePath.get()).parent_path() / L"Themes";
}

void CancelFunctionBarPressedKeyClearTimer(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    KillTimer(hwnd, kFunctionBarPressedKeyClearTimerId);
    g_functionBarPressedKeyClearPending = std::nullopt;
}

void SetFunctionBarPressedKeyState(std::optional<uint32_t> vk) noexcept
{
    g_functionBarPressedKey = vk;
    g_folderWindow.SetFunctionBarPressedKey(vk);
}

void ScheduleFunctionBarPressedKeyClear(HWND hwnd, uint32_t vk) noexcept
{
    if (! hwnd)
    {
        return;
    }

    g_functionBarPressedKeyClearPending = vk;
    SetTimer(hwnd, kFunctionBarPressedKeyClearTimerId, kFunctionBarPressedKeyClearDelayMs, nullptr);
}

LRESULT OnMainWindowTimer(HWND hWnd, UINT_PTR timerId) noexcept
{
#ifdef ENABLE_TESTS
    if (timerId == kFileOpsSelfTestTimerId)
    {
        static std::atomic<bool> tickInProgress{false};
        if (tickInProgress.exchange(true, std::memory_order_acq_rel))
        {
            return 0;
        }
        const auto clearTickInProgress = wil::scope_exit([&] { tickInProgress.store(false, std::memory_order_release); });

        if (g_fileOpsSelfTestStartPending && ! FileOperationsSelfTest::IsRunning())
        {
            g_fileOpsSelfTestStartPending = false;
            StartNextFileOpsSelfTestRun(hWnd);
            return 0;
        }

        const bool done = FileOperationsSelfTest::Tick(hWnd);
        if (done)
        {
            const bool currentRunFailed                          = FileOperationsSelfTest::DidFail();
            const SelfTest::SelfTestSuiteResult currentRunResult = FileOperationsSelfTest::GetSuiteResult();
            MergeFileOpsSuite(g_fileOpsSelfTestAggregateResult, currentRunResult);
            g_selfTestExitCode |= currentRunFailed ? 1 : 0;

            const bool hasMoreRuns = g_fileOpsSelfTestRunIndex < g_fileOpsSelfTestRunFilters.size();
            const bool stopEarly   = currentRunFailed && g_selfTestOptions.failFast;
            if (hasMoreRuns && ! stopEarly)
            {
                g_fileOpsSelfTestStartPending = true;
                return 0;
            }

            KillTimer(hWnd, kFileOpsSelfTestTimerId);
            FinalizeFileOpsAggregateResult(g_fileOpsSelfTestAggregateResult, g_fileOpsSelfTestExpectedCases, stopEarly);

            const bool fileOpsFailed = g_fileOpsSelfTestAggregateResult.failed != 0;
            g_selfTestExitCode |= fileOpsFailed ? 1 : 0;
            RecordSelfTestSuite(g_fileOpsSelfTestAggregateResult);
            if (SelfTest::GetSelfTestOptions().writeJsonSummary)
            {
                const std::filesystem::path jsonPath = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::FileOperations, L"results.json");
                SelfTest::WriteSuiteJson(g_fileOpsSelfTestAggregateResult, jsonPath);
            }

            if (fileOpsFailed)
            {
                SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, L"FileOpsSelfTest: FAIL");
                SelfTest::AppendSelfTestTrace(L"FileOpsSelfTest: FAIL");
                const std::wstring_view message = g_fileOpsSelfTestAggregateResult.failureMessage;
                if (! message.empty())
                {
                    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, message);
                    SelfTest::AppendSelfTestTrace(message);
                }
            }
            else
            {
                SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, L"FileOpsSelfTest: PASS");
                SelfTest::AppendSelfTestTrace(L"FileOpsSelfTest: PASS");
            }
            TraceSelfTestExitCode(L"FileOpsSelfTest: end", g_selfTestExitCode);
            // Route file-operations self-test shutdown through WM_CLOSE/WM_DESTROY as well so
            // plugin managers, hot-reload, and window-owned resources tear down before process exit.
            SplashScreen::CloseIfExist();
            if (PostMessageW(hWnd, WM_CLOSE, 0, 0) == 0)
            {
                FinalizeSelfTestRun();
                PostQuitMessage(g_selfTestExitCode);
            }
        }
        return 0;
    }
#endif

    if (timerId != kFunctionBarPressedKeyClearTimerId)
    {
        return DefWindowProcW(hWnd, WM_TIMER, timerId, 0);
    }

    KillTimer(hWnd, kFunctionBarPressedKeyClearTimerId);

    if (g_functionBarPressedKeyClearPending.has_value() && g_functionBarPressedKey.has_value() &&
        g_functionBarPressedKey.value() == g_functionBarPressedKeyClearPending.value())
    {
        SetFunctionBarPressedKeyState(std::nullopt);
    }

    g_functionBarPressedKeyClearPending = std::nullopt;
    return 0;
}

ThemeMode ThemeModeFromThemeId(std::wstring_view id) noexcept
{
    if (id == L"builtin/light")
    {
        return ThemeMode::Light;
    }
    if (id == L"builtin/dark")
    {
        return ThemeMode::Dark;
    }
    if (id == L"builtin/rainbow")
    {
        return ThemeMode::Rainbow;
    }
    if (id == L"builtin/highContrast")
    {
        return ThemeMode::HighContrast;
    }
    return ThemeMode::System;
}

FolderView::DisplayMode DisplayModeFromSettings(Common::Settings::FolderDisplayMode mode) noexcept
{
    switch (mode)
    {
        case Common::Settings::FolderDisplayMode::Brief: return FolderView::DisplayMode::Brief;
        case Common::Settings::FolderDisplayMode::Detailed: return FolderView::DisplayMode::Detailed;
        case Common::Settings::FolderDisplayMode::ExtraDetailed: return FolderView::DisplayMode::ExtraDetailed;
        case Common::Settings::FolderDisplayMode::Thumbnails: return FolderView::DisplayMode::Thumbnails;
    }
    return FolderView::DisplayMode::Brief;
}

FolderView::SortBy SortByFromSettings(Common::Settings::FolderSortBy sortBy) noexcept
{
    switch (sortBy)
    {
        case Common::Settings::FolderSortBy::Name: return FolderView::SortBy::Name;
        case Common::Settings::FolderSortBy::Extension: return FolderView::SortBy::Extension;
        case Common::Settings::FolderSortBy::Time: return FolderView::SortBy::Time;
        case Common::Settings::FolderSortBy::Size: return FolderView::SortBy::Size;
        case Common::Settings::FolderSortBy::Attributes: return FolderView::SortBy::Attributes;
        case Common::Settings::FolderSortBy::None: return FolderView::SortBy::None;
    }
    return FolderView::SortBy::Name;
}

FolderView::SortDirection SortDirectionFromSettings(Common::Settings::FolderSortDirection direction) noexcept
{
    switch (direction)
    {
        case Common::Settings::FolderSortDirection::Ascending: return FolderView::SortDirection::Ascending;
        case Common::Settings::FolderSortDirection::Descending: return FolderView::SortDirection::Descending;
    }
    return FolderView::SortDirection::Ascending;
}

Common::Settings::FolderDisplayMode DisplayModeToSettings(FolderView::DisplayMode mode) noexcept
{
    switch (mode)
    {
        case FolderView::DisplayMode::Brief: return Common::Settings::FolderDisplayMode::Brief;
        case FolderView::DisplayMode::Detailed: return Common::Settings::FolderDisplayMode::Detailed;
        case FolderView::DisplayMode::ExtraDetailed: return Common::Settings::FolderDisplayMode::ExtraDetailed;
        case FolderView::DisplayMode::Thumbnails: return Common::Settings::FolderDisplayMode::Thumbnails;
    }
    return Common::Settings::FolderDisplayMode::Brief;
}

Common::Settings::FolderSortBy SortByToSettings(FolderView::SortBy sortBy) noexcept
{
    switch (sortBy)
    {
        case FolderView::SortBy::Name: return Common::Settings::FolderSortBy::Name;
        case FolderView::SortBy::Extension: return Common::Settings::FolderSortBy::Extension;
        case FolderView::SortBy::Time: return Common::Settings::FolderSortBy::Time;
        case FolderView::SortBy::Size: return Common::Settings::FolderSortBy::Size;
        case FolderView::SortBy::Attributes: return Common::Settings::FolderSortBy::Attributes;
        case FolderView::SortBy::None: return Common::Settings::FolderSortBy::None;
    }
    return Common::Settings::FolderSortBy::Name;
}

Common::Settings::FolderSortDirection SortDirectionToSettings(FolderView::SortDirection direction) noexcept
{
    switch (direction)
    {
        case FolderView::SortDirection::Ascending: return Common::Settings::FolderSortDirection::Ascending;
        case FolderView::SortDirection::Descending: return Common::Settings::FolderSortDirection::Descending;
    }
    return Common::Settings::FolderSortDirection::Ascending;
}

std::wstring EscapeMenuLabel(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());

    for (wchar_t ch : text)
    {
        if (ch == L'\t')
        {
            result.push_back(L' ');
            continue;
        }

        result.push_back(ch);
        if (ch == L'&')
        {
            result.push_back(L'&');
        }
    }

    return result;
}

std::wstring ThemeIdFromThemeMode(ThemeMode mode)
{
    switch (mode)
    {
        case ThemeMode::Light: return L"builtin/light";
        case ThemeMode::Dark: return L"builtin/dark";
        case ThemeMode::Rainbow: return L"builtin/rainbow";
        case ThemeMode::HighContrast: return L"builtin/highContrast";
        case ThemeMode::System:
        default: return L"builtin/system";
    }
}

const Common::Settings::ThemeDefinition* FindThemeById(std::wstring_view id) noexcept;

void UpdateThemeModeFromCurrentSettings() noexcept
{
    std::wstring_view themeId = g_settings.theme.currentThemeId;
    if (themeId.rfind(L"user/", 0) == 0)
    {
        if (const auto* def = FindThemeById(themeId))
        {
            themeId = def->baseThemeId;
        }
    }

    g_themeMode = ThemeModeFromThemeId(themeId);
}

const Common::Settings::ThemeDefinition* FindThemeById(std::wstring_view id) noexcept
{
    for (const auto& def : g_settings.theme.themes)
    {
        if (def.id == id)
        {
            return &def;
        }
    }
    for (const auto& def : g_fileThemes)
    {
        if (def.id == id)
        {
            return &def;
        }
    }
    return nullptr;
}

struct CustomThemeGroups
{
    std::vector<const Common::Settings::ThemeDefinition*> fileThemes;
    std::vector<const Common::Settings::ThemeDefinition*> settingsThemes;
};

[[nodiscard]] CustomThemeGroups CollectCustomThemeGroups()
{
    CustomThemeGroups groups;

    std::unordered_map<std::wstring, const Common::Settings::ThemeDefinition*> settingsThemesById;
    settingsThemesById.reserve(g_settings.theme.themes.size());
    for (const auto& def : g_settings.theme.themes)
    {
        if (def.id.rfind(L"user/", 0) != 0)
        {
            continue;
        }
        settingsThemesById[def.id] = &def;
    }

    groups.settingsThemes.reserve(settingsThemesById.size());
    for (const auto& entry : settingsThemesById)
    {
        groups.settingsThemes.push_back(entry.second);
    }

    groups.fileThemes.reserve(g_fileThemes.size());
    for (const auto& def : g_fileThemes)
    {
        if (def.id.rfind(L"user/", 0) != 0)
        {
            continue;
        }
        if (settingsThemesById.contains(def.id))
        {
            continue;
        }
        groups.fileThemes.push_back(&def);
    }

    const auto byNameThenId = [](const Common::Settings::ThemeDefinition* a, const Common::Settings::ThemeDefinition* b)
    {
        if (a->name == b->name)
        {
            return a->id < b->id;
        }
        return a->name < b->name;
    };

    std::sort(groups.fileThemes.begin(), groups.fileThemes.end(), byNameThenId);
    std::sort(groups.settingsThemes.begin(), groups.settingsThemes.end(), byNameThenId);
    return groups;
}

COLORREF ColorRefFromArgb(uint32_t argb) noexcept
{
    const uint8_t r = static_cast<uint8_t>((argb >> 16) & 0xFFu);
    const uint8_t g = static_cast<uint8_t>((argb >> 8) & 0xFFu);
    const uint8_t b = static_cast<uint8_t>(argb & 0xFFu);
    return RGB(r, g, b);
}

float AlphaFromArgb(uint32_t argb) noexcept
{
    const uint8_t a = static_cast<uint8_t>((argb >> 24) & 0xFFu);
    return static_cast<float>(a) / 255.0f;
}

std::optional<uint32_t> FindColorOverride(const std::unordered_map<std::wstring, uint32_t>& colors, std::wstring_view key) noexcept
{
    const auto it = colors.find(std::wstring(key));
    if (it == colors.end())
    {
        return std::nullopt;
    }
    return it->second;
}

void ApplyThemeOverrides(AppTheme& theme, const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    const auto applyColorRef = [&](std::wstring_view key, COLORREF& target) noexcept
    {
        const auto argb = FindColorOverride(colors, key);
        if (! argb)
        {
            return;
        }
        target = ColorRefFromArgb(*argb);
    };

    const auto applyD2D = [&](std::wstring_view key, D2D1::ColorF& target) noexcept
    {
        const auto argb = FindColorOverride(colors, key);
        if (! argb)
        {
            return;
        }
        const COLORREF rgb = ColorRefFromArgb(*argb);
        target             = ColorFromCOLORREF(rgb, AlphaFromArgb(*argb));
    };

    applyD2D(L"app.accent", theme.accent);
    applyColorRef(L"window.background", theme.windowBackground);

    applyColorRef(L"menu.background", theme.menu.background);
    applyColorRef(L"menu.text", theme.menu.text);
    applyColorRef(L"menu.disabledText", theme.menu.disabledText);
    applyColorRef(L"menu.selectionBg", theme.menu.selectionBg);
    applyColorRef(L"menu.selectionText", theme.menu.selectionText);
    applyColorRef(L"menu.separator", theme.menu.separator);
    applyColorRef(L"menu.border", theme.menu.border);

    applyD2D(L"navigation.background", theme.navigationView.background);
    applyD2D(L"navigation.backgroundHover", theme.navigationView.backgroundHover);
    applyD2D(L"navigation.backgroundPressed", theme.navigationView.backgroundPressed);
    applyD2D(L"navigation.text", theme.navigationView.text);
    applyD2D(L"navigation.separator", theme.navigationView.separator);
    applyD2D(L"navigation.accent", theme.navigationView.accent);
    applyD2D(L"navigation.progressOk", theme.navigationView.progressOk);
    applyD2D(L"navigation.progressWarn", theme.navigationView.progressWarn);
    applyD2D(L"navigation.progressBackground", theme.navigationView.progressBackground);

    if (const auto argb = FindColorOverride(colors, L"navigation.background"))
    {
        const COLORREF rgb                 = ColorRefFromArgb(*argb);
        theme.navigationView.gdiBackground = rgb;
        theme.navigationView.gdiBorder     = rgb;
    }

    if (const auto argb = FindColorOverride(colors, L"navigation.separator"))
    {
        theme.navigationView.gdiBorderPen = ColorRefFromArgb(*argb);
    }

    applyD2D(L"folderView.background", theme.folderView.backgroundColor);
    applyD2D(L"folderView.itemBackgroundNormal", theme.folderView.itemBackgroundNormal);
    applyD2D(L"folderView.itemBackgroundHovered", theme.folderView.itemBackgroundHovered);
    applyD2D(L"folderView.itemBackgroundSelected", theme.folderView.itemBackgroundSelected);
    applyD2D(L"folderView.itemBackgroundSelectedInactive", theme.folderView.itemBackgroundSelectedInactive);
    applyD2D(L"folderView.itemBackgroundFocused", theme.folderView.itemBackgroundFocused);
    applyD2D(L"folderView.textNormal", theme.folderView.textNormal);
    applyD2D(L"folderView.textSelected", theme.folderView.textSelected);
    applyD2D(L"folderView.textSelectedInactive", theme.folderView.textSelectedInactive);
    applyD2D(L"folderView.textDisabled", theme.folderView.textDisabled);
    applyD2D(L"folderView.focusBorder", theme.folderView.focusBorder);
    applyD2D(L"folderView.gridLines", theme.folderView.gridLines);
    applyD2D(L"folderView.errorBackground", theme.folderView.errorBackground);
    applyD2D(L"folderView.errorText", theme.folderView.errorText);
    applyD2D(L"folderView.warningBackground", theme.folderView.warningBackground);
    applyD2D(L"folderView.warningText", theme.folderView.warningText);
    applyD2D(L"folderView.infoBackground", theme.folderView.infoBackground);
    applyD2D(L"folderView.infoText", theme.folderView.infoText);

    // Derive file operation colors from the effective theme (post-overrides).
    theme.fileOperations.progressBackground = theme.navigationView.progressBackground;
    theme.fileOperations.progressTotal      = theme.navigationView.progressOk;
    theme.fileOperations.progressItem       = theme.navigationView.accent;

    const D2D1::ColorF menuBorder   = ColorFromCOLORREF(theme.menu.border);
    const D2D1::ColorF menuDisabled = ColorFromCOLORREF(theme.menu.disabledText);

    theme.fileOperations.graphBackground =
        D2D1::ColorF(theme.fileOperations.progressBackground.r, theme.fileOperations.progressBackground.g, theme.fileOperations.progressBackground.b, 0.35f);
    theme.fileOperations.graphGrid      = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.35f);
    theme.fileOperations.graphLimit     = D2D1::ColorF(menuDisabled.r, menuDisabled.g, menuDisabled.b, 0.85f);
    theme.fileOperations.graphLine      = theme.fileOperations.progressItem;
    theme.fileOperations.scrollbarTrack = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.12f);
    theme.fileOperations.scrollbarThumb = D2D1::ColorF(menuBorder.r, menuBorder.g, menuBorder.b, 0.40f);

    applyD2D(L"fileOps.progressBackground", theme.fileOperations.progressBackground);
    applyD2D(L"fileOps.progressTotal", theme.fileOperations.progressTotal);
    applyD2D(L"fileOps.progressItem", theme.fileOperations.progressItem);
    applyD2D(L"fileOps.graphBackground", theme.fileOperations.graphBackground);
    applyD2D(L"fileOps.graphGrid", theme.fileOperations.graphGrid);
    applyD2D(L"fileOps.graphLimit", theme.fileOperations.graphLimit);
    applyD2D(L"fileOps.graphLine", theme.fileOperations.graphLine);
    applyD2D(L"fileOps.scrollbarTrack", theme.fileOperations.scrollbarTrack);
    applyD2D(L"fileOps.scrollbarThumb", theme.fileOperations.scrollbarThumb);

    applyD2D(L"viewer.diff.addedBackground", theme.viewerDiff.addedBackground);
    applyD2D(L"viewer.diff.removedBackground", theme.viewerDiff.removedBackground);
    applyD2D(L"viewer.diff.contextBackground", theme.viewerDiff.contextBackground);
    applyD2D(L"viewer.diff.headerBackground", theme.viewerDiff.headerBackground);
    applyD2D(L"viewer.diff.bannerBackground", theme.viewerDiff.bannerBackground);
    applyD2D(L"viewer.diff.placeholderBackground", theme.viewerDiff.placeholderBackground);
    applyD2D(L"viewer.diff.divider", theme.viewerDiff.divider);

    if (! FindColorOverride(colors, L"folderView.itemBackgroundSelectedInactive"))
    {
        if (const auto argb = FindColorOverride(colors, L"folderView.itemBackgroundSelected"))
        {
            const float inactiveSelectionAlphaScale = theme.highContrast ? 0.80f : 0.65f;
            const COLORREF rgb                      = ColorRefFromArgb(*argb);
            theme.folderView.itemBackgroundSelectedInactive =
                ColorFromCOLORREF(rgb, std::clamp(AlphaFromArgb(*argb) * inactiveSelectionAlphaScale, 0.0f, 1.0f));
        }
    }

    if (! FindColorOverride(colors, L"folderView.textSelectedInactive") && ! theme.highContrast)
    {
        const float alpha             = std::clamp(theme.folderView.itemBackgroundSelectedInactive.a, 0.0f, 1.0f);
        const D2D1::ColorF background = theme.folderView.backgroundColor;
        const D2D1::ColorF overlay    = theme.folderView.itemBackgroundSelectedInactive;

        const D2D1::ColorF composite = D2D1::ColorF(overlay.r * alpha + background.r * (1.0f - alpha),
                                                    overlay.g * alpha + background.g * (1.0f - alpha),
                                                    overlay.b * alpha + background.b * (1.0f - alpha),
                                                    1.0f);

        const COLORREF contrastText           = ChooseContrastingTextColor(ColorToCOLORREF(composite));
        theme.folderView.textSelectedInactive = ColorFromCOLORREF(contrastText);
    }
}

std::optional<D2D1::ColorF> FindAccentOverride(const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    const auto argb = FindColorOverride(colors, L"app.accent");
    if (! argb)
    {
        return std::nullopt;
    }
    const COLORREF rgb = ColorRefFromArgb(*argb);
    return ColorFromCOLORREF(rgb, AlphaFromArgb(*argb));
}

AppTheme ResolveConfiguredTheme() noexcept
{
    std::wstring_view themeId = g_settings.theme.currentThemeId;

    const Common::Settings::ThemeDefinition* custom = nullptr;
    if (themeId.rfind(L"user/", 0) == 0)
    {
        custom = FindThemeById(themeId);
    }

    ThemeMode baseMode = ThemeModeFromThemeId(themeId);
    std::optional<D2D1::ColorF> accentOverride;
    const std::unordered_map<std::wstring, uint32_t>* overrides = nullptr;

    if (custom)
    {
        baseMode       = ThemeModeFromThemeId(custom->baseThemeId);
        accentOverride = FindAccentOverride(custom->colors);
        overrides      = &custom->colors;
    }

    AppTheme theme = ResolveAppTheme(baseMode, L"RedSalamander", accentOverride);
    if (overrides)
    {
        ApplyThemeOverrides(theme, *overrides);
    }

    SettingsHotReload::ApplyUiPreferencesToTheme(g_settings, theme);

    return theme;
}

std::optional<std::filesystem::path> GetDefaultFolder() noexcept
{
    wil::unique_cotaskmem_string folderPath;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Documents, 0, nullptr, folderPath.put())) && folderPath)
    {
        return std::filesystem::path(folderPath.get());
    }
    return std::nullopt;
}

Localization::LanguagePreference GetLanguagePreferenceFromSettings(const Common::Settings::Settings& settings)
{
    Localization::LanguagePreference preference;
    if (! settings.ui || settings.ui->language.empty() || settings.ui->language == L"system")
    {
        preference.kind = Localization::LanguagePreferenceKind::System;
        return preference;
    }

    preference.kind    = Localization::LanguagePreferenceKind::Culture;
    preference.culture = settings.ui->language;
    return preference;
}

void CaptureRuntimeSettings(Common::Settings::Settings& settings, HWND hWnd) noexcept
{
    if (hWnd)
    {
        WINDOWPLACEMENT placement{};
        placement.length = sizeof(placement);
        if (GetWindowPlacement(hWnd, &placement))
        {
            Common::Settings::WindowPlacement wp;
            wp.state = placement.showCmd == SW_SHOWMAXIMIZED ? Common::Settings::WindowState::Maximized : Common::Settings::WindowState::Normal;

            const RECT rc         = placement.rcNormalPosition;
            wp.bounds.x           = rc.left;
            wp.bounds.y           = rc.top;
            const int savedWidth  = static_cast<int>(rc.right - rc.left);
            const int savedHeight = static_cast<int>(rc.bottom - rc.top);
            wp.bounds.width       = std::max(1, savedWidth);
            wp.bounds.height      = std::max(1, savedHeight);
            wp.dpi                = GetDpiForWindow(hWnd);

            settings.windows[kMainWindowId] = std::move(wp);
        }
    }

    if (const HWND prefs = GetPreferencesDialogHandle())
    {
        WindowPlacementPersistence::Save(settings, kPreferencesWindowId, prefs);
    }

    if (const HWND connections = GetConnectionManagerDialogHandle())
    {
        WindowPlacementPersistence::Save(settings, kConnectionManagerWindowId, connections);
    }

    if (const HWND shortcuts = GetShortcutsWindowHandle())
    {
        WindowPlacementPersistence::Save(settings, kShortcutsWindowId, shortcuts);
    }

    if (const HWND findFiles = GetFindFilesWindowHandle())
    {
        WindowPlacementPersistence::Save(settings, kFindFilesWindowId, findFiles);
    }

    const auto captureItemPropertiesIfMatch = [&](HWND hwnd) noexcept
    {
        if (! hwnd)
        {
            return false;
        }

        wchar_t className[128]{};
        const int len = GetClassNameW(hwnd, className, static_cast<int>(std::size(className)));
        if (len <= 0)
        {
            return false;
        }

        if (wcscmp(className, kItemPropertiesWindowClassName) != 0)
        {
            return false;
        }

        WindowPlacementPersistence::Save(settings, kItemPropertiesWindowId, hwnd);
        return true;
    };

    if (! captureItemPropertiesIfMatch(GetForegroundWindow()))
    {
        if (const HWND props = FindWindowW(kItemPropertiesWindowClassName, nullptr))
        {
            WindowPlacementPersistence::Save(settings, kItemPropertiesWindowId, props);
        }
    }

    if (g_hFolderWindow.load(std::memory_order_acquire))
    {
        std::vector<Common::Settings::FolderHistoryFilterState> historyFilters;
        if (settings.folders.has_value())
        {
            historyFilters = settings.folders->historyFilters;
        }

        Common::Settings::FoldersSettings folders;
        const FolderWindow::Pane activePane = g_folderWindow.GetFocusedPane();
        folders.active                      = activePane == FolderWindow::Pane::Right ? kRightPaneSlot : kLeftPaneSlot;
        folders.layout.splitRatio           = g_folderWindow.GetSplitRatio();
        folders.showHiddenFiles             = g_folderWindow.GetShowHiddenFiles();
        folders.showSystemFiles             = g_folderWindow.GetShowSystemFiles();
        if (const std::optional<FolderWindow::Pane> zoomedPane = g_folderWindow.GetZoomedPane())
        {
            folders.layout.zoomedPane            = zoomedPane.value() == FolderWindow::Pane::Left ? kLeftPaneSlot : kRightPaneSlot;
            folders.layout.zoomRestoreSplitRatio = g_folderWindow.GetZoomRestoreSplitRatio();
        }

        const std::filesystem::path safeDefault = GetDefaultFolder().value_or(std::filesystem::path(L"C:\\"));

        auto addPane = [&](FolderWindow::Pane paneId, std::wstring_view slot)
        {
            Common::Settings::FolderPane pane;
            pane.slot = std::wstring(slot);

            std::filesystem::path current                         = safeDefault;
            const std::optional<std::filesystem::path> currentOpt = g_folderWindow.GetCurrentPath(paneId);
            if (currentOpt.has_value() && ! currentOpt.value().empty())
            {
                current = currentOpt.value();
            }

            pane.current = current;

            pane.view.display          = DisplayModeToSettings(g_folderWindow.GetDisplayMode(paneId));
            pane.view.sortBy           = SortByToSettings(g_folderWindow.GetSortBy(paneId));
            pane.view.sortDirection    = SortDirectionToSettings(g_folderWindow.GetSortDirection(paneId));
            pane.view.fileExtensionsVisible = g_folderWindow.GetFileExtensionsVisible(paneId);
            pane.view.thumbnailsVisible     = false;
            pane.view.navigationBarVisible  = g_folderWindow.GetNavigationBarVisible(paneId);
            pane.view.filterBarVisible      = g_folderWindow.GetFilterBarVisible(paneId);
            pane.view.statusBarVisible = g_folderWindow.GetStatusBarVisible(paneId);

            folders.items.push_back(std::move(pane));
        };

        addPane(FolderWindow::Pane::Left, kLeftPaneSlot);
        addPane(FolderWindow::Pane::Right, kRightPaneSlot);

        folders.historyMax = g_folderWindow.GetFolderHistoryMax();
        folders.history    = g_folderWindow.GetFolderHistory();

        if (! historyFilters.empty() && ! folders.history.empty())
        {
            folders.historyFilters.reserve(folders.history.size());
            for (const auto& historyPath : folders.history)
            {
                if (historyPath.empty())
                {
                    continue;
                }

                const std::wstring_view historyText = historyPath.native();
                const auto it =
                    std::find_if(historyFilters.begin(), historyFilters.end(), [&](const Common::Settings::FolderHistoryFilterState& state) noexcept {
                    return ! state.path.empty() && OrdinalString::EqualsNoCasePath(state.path, historyText);
                });
                if (it == historyFilters.end())
                {
                    continue;
                }

                if (! it->enabled && it->text.empty())
                {
                    continue;
                }

                Common::Settings::FolderHistoryFilterState state = *it;
                state.path                                       = historyPath;
                folders.historyFilters.push_back(std::move(state));
            }
        }

        settings.folders = std::move(folders);
    }

    Common::Settings::MainMenuState menuState;
    menuState.menuBarVisible     = g_menuBarVisible;
    menuState.functionBarVisible = g_functionBarVisible;
    settings.mainMenu            = menuState;
}

void CaptureRuntimeSettings(HWND hWnd) noexcept
{
    CaptureRuntimeSettings(g_settings, hWnd);
}

void SaveAppSettings(HWND hWnd) noexcept
{
    CaptureRuntimeSettings(hWnd);

    g_folderWindow.CloseAllViewers();
    const auto pluginSchemas = CollectPluginConfigurationSchemas(g_settings);
    if (g_hFolderWindow.exchange(nullptr, std::memory_order_acq_rel))
    {
        g_folderWindow.Destroy();
    }
    FileSystemPluginManager::GetInstance().Shutdown(g_settings);
    ViewerPluginManager::GetInstance().Shutdown(g_settings);

    const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(kAppId, g_settings, pluginSchemas);
    if (SUCCEEDED(saveHr))
    {
        return;
    }

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kAppId);
    DBGOUT_ERROR(L"SaveSettings failed (hr=0x{:08X}) path={}\n", static_cast<unsigned long>(saveHr), settingsPath.wstring());
}

void UpdateThemeMenuChecks() noexcept
{
    if (! g_viewThemeMenu)
    {
        return;
    }

    const bool highContrastEnabled = IsHighContrastEnabled();

    EnableMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST, MF_BYCOMMAND | MF_GRAYED);
    CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST, static_cast<UINT>(MF_BYCOMMAND | (highContrastEnabled ? MF_CHECKED : MF_UNCHECKED)));

    for (const auto& [id, _] : g_customThemeMenuIdToThemeId)
    {
        CheckMenuItem(g_viewThemeMenu, id, MF_BYCOMMAND | MF_UNCHECKED);
    }

    const std::wstring currentThemeId = g_settings.theme.currentThemeId;
    const auto customIt               = g_customThemeIdToMenuId.find(currentThemeId);
    if (customIt != g_customThemeIdToMenuId.end())
    {
        CheckMenuItem(g_viewThemeMenu, customIt->second, MF_BYCOMMAND | MF_CHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_SYSTEM, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_LIGHT, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_DARK, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_RAINBOW, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP, MF_BYCOMMAND | MF_UNCHECKED);
        return;
    }

    const ThemeMode mode = ThemeModeFromThemeId(currentThemeId);
    if (mode == ThemeMode::HighContrast)
    {
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_SYSTEM, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_LIGHT, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_DARK, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_RAINBOW, MF_BYCOMMAND | MF_UNCHECKED);
        CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP, MF_BYCOMMAND | MF_CHECKED);
        return;
    }

    CheckMenuItem(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP, MF_BYCOMMAND | MF_UNCHECKED);
    UINT checkedId = IDM_VIEW_THEME_SYSTEM;
    switch (mode)
    {
        case ThemeMode::System: checkedId = IDM_VIEW_THEME_SYSTEM; break;
        case ThemeMode::Light: checkedId = IDM_VIEW_THEME_LIGHT; break;
        case ThemeMode::Dark: checkedId = IDM_VIEW_THEME_DARK; break;
        case ThemeMode::Rainbow: checkedId = IDM_VIEW_THEME_RAINBOW; break;
        case ThemeMode::HighContrast: break;
    }

    CheckMenuRadioItem(g_viewThemeMenu, IDM_VIEW_THEME_SYSTEM, IDM_VIEW_THEME_RAINBOW, checkedId, MF_BYCOMMAND);
}

void UpdatePluginsMenuChecks() noexcept
{
    if (! g_viewPluginsMenu)
    {
        return;
    }

    const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
    std::wstring activeId(g_folderWindow.GetFileSystemPluginId(pane));
    if (activeId.empty())
    {
        activeId = std::wstring(FileSystemPluginManager::GetInstance().GetActivePluginId());
    }

    for (const auto& [id, _] : g_pluginMenuIdToPluginId)
    {
        CheckMenuItem(g_viewPluginsMenu, id, MF_BYCOMMAND | MF_UNCHECKED);
    }

    const auto it = g_pluginIdToMenuId.find(activeId);
    if (it != g_pluginIdToMenuId.end())
    {
        CheckMenuItem(g_viewPluginsMenu, it->second, MF_BYCOMMAND | MF_CHECKED);
    }
}

static bool TryFindMenuPathToCommand(HMENU menu, UINT commandId, std::vector<HMENU>& path) noexcept
{
    if (! menu)
    {
        return false;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return false;
    }

    path.push_back(menu);

    for (int pos = 0; pos < count; ++pos)
    {
        const UINT id = GetMenuItemID(menu, pos);
        if (id == commandId)
        {
            return true;
        }

        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (TryFindMenuPathToCommand(subMenu, commandId, path))
        {
            return true;
        }
    }

    path.pop_back();
    return false;
}

static size_t CommonPrefixSize(const std::vector<HMENU>& a, const std::vector<HMENU>& b) noexcept
{
    size_t i = 0;
    while (i < a.size() && i < b.size() && a[i] == b[i])
    {
        ++i;
    }
    return i;
}

static int FindMenuItemPosById(HMENU menu, UINT id) noexcept
{
    if (! menu)
    {
        return -1;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return -1;
    }

    for (int pos = 0; pos < count; ++pos)
    {
        if (GetMenuItemID(menu, pos) == id)
        {
            return pos;
        }
    }

    return -1;
}

static HMENU FindSubMenuAfterCommand(HMENU menu, UINT commandId) noexcept
{
    if (! menu)
    {
        return nullptr;
    }

    const int commandPos = FindMenuItemPosById(menu, commandId);
    if (commandPos < 0)
    {
        return nullptr;
    }

    const int count = GetMenuItemCount(menu);
    if (count <= commandPos + 1)
    {
        return nullptr;
    }

    return GetSubMenu(menu, commandPos + 1);
}

static HMENU FindSubMenuAfterCommandRecursive(HMENU menu, UINT commandId) noexcept
{
    HMENU found = FindSubMenuAfterCommand(menu, commandId);
    if (found)
    {
        return found;
    }

    const int count = GetMenuItemCount(menu);
    if (count <= 0)
    {
        return nullptr;
    }

    for (int pos = 0; pos < count; ++pos)
    {
        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }

        found = FindSubMenuAfterCommandRecursive(subMenu, commandId);
        if (found)
        {
            return found;
        }
    }

    return nullptr;
}

static void DeleteMenuItemsFromPosition(HMENU menu, int startPos) noexcept
{
    if (! menu)
    {
        return;
    }

    const int count = GetMenuItemCount(menu);
    if (count < 0)
    {
        return;
    }

    for (int pos = count - 1; pos >= startPos; --pos)
    {
        DeleteMenu(menu, static_cast<UINT>(pos), MF_BYPOSITION);
    }
}

static bool IsOverlaySampleEnabled() noexcept
{
#if defined(_DEBUG) || defined(DEBUG)
    return true;
#else
    return false;
#endif
}

static bool IsMenuSeparatorAt(HMENU menu, int pos) noexcept
{
    if (! menu || pos < 0)
    {
        return false;
    }

    MENUITEMINFOW mii{};
    mii.cbSize = sizeof(mii);
    mii.fMask  = MIIM_FTYPE;
    if (! GetMenuItemInfoW(menu, static_cast<UINT>(pos), TRUE, &mii))
    {
        return false;
    }

    return (mii.fType & MFT_SEPARATOR) != 0;
}

static bool MenuContainsCommandIdRecursive(HMENU menu, UINT commandId) noexcept
{
    if (! menu)
    {
        return false;
    }

    if (FindMenuItemPosById(menu, commandId) >= 0)
    {
        return true;
    }

    const int count = GetMenuItemCount(menu);
    if (count <= 0)
    {
        return false;
    }

    for (int pos = 0; pos < count; ++pos)
    {
        HMENU subMenu = GetSubMenu(menu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (MenuContainsCommandIdRecursive(subMenu, commandId))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] static wchar_t GetMainMenuCommandIconGlyph(UINT menuCommandId) noexcept
{
    switch (menuCommandId)
    {
        case IDM_FILE_PREFERENCES: return FluentIcons::kSettings;
        case IDM_APP_COMPARE: return FluentIcons::kSyncFolder;
        case IDM_VIEW_PLUGINS_MANAGE: return FluentIcons::kPuzzle;
        case IDM_PANE_EXECUTE_OPEN: return FluentIcons::kOpenFile;
        case IDM_PANE_CONNECTION_MANAGER: return FluentIcons::kConnections;
        case IDM_PANE_CLIPBOARD_CUT: return FluentIcons::kCut;
        case IDM_PANE_CLIPBOARD_COPY: return FluentIcons::kCopy;
        case IDM_PANE_CLIPBOARD_PASTE: return FluentIcons::kPaste;
        case IDM_PANE_RENAME: return FluentIcons::kRename;
        case IDM_PANE_DELETE: return FluentIcons::kDelete;
        case IDM_PANE_OPEN_PROPERTIES: return FluentIcons::kInfo;
        case IDM_PANE_CONNECT: return FluentIcons::kMapDrive;
        case IDM_PANE_SHOW_FOLDERS_HISTORY: return FluentIcons::kHistory;
        case IDM_PANE_FIND: return FluentIcons::kFind;
        case IDM_PANE_OPEN_COMMAND_SHELL: return FluentIcons::kCommandPrompt;
    }

    return 0;
}

static void RemoveOverlaySampleSubmenu(HMENU paneMenu, UINT sampleErrorCommandId) noexcept
{
    if (! paneMenu)
    {
        return;
    }

    const int itemCount = GetMenuItemCount(paneMenu);
    if (itemCount <= 0)
    {
        return;
    }

    for (int pos = 0; pos < itemCount; ++pos)
    {
        HMENU subMenu = GetSubMenu(paneMenu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (! MenuContainsCommandIdRecursive(subMenu, sampleErrorCommandId))
        {
            continue;
        }

        DeleteMenu(paneMenu, static_cast<UINT>(pos), MF_BYPOSITION);
        if (pos > 0 && IsMenuSeparatorAt(paneMenu, pos - 1))
        {
            DeleteMenu(paneMenu, static_cast<UINT>(pos - 1), MF_BYPOSITION);
        }
        break;
    }
}

static void EnsurePaneMenuHandlesFor(
    HMENU mainMenu, UINT sortCommand, UINT displayCommand, HMENU& outPane, HMENU& outSort, HMENU& outDisplay, HMENU& outHistory) noexcept
{
    std::vector<HMENU> sortPath;
    std::vector<HMENU> displayPath;
    if (! TryFindMenuPathToCommand(mainMenu, sortCommand, sortPath) || ! TryFindMenuPathToCommand(mainMenu, displayCommand, displayPath))
    {
        return;
    }

    const size_t common = CommonPrefixSize(sortPath, displayPath);
    if (common < 2)
    {
        return;
    }

    HMENU paneMenu = sortPath[common - 1];
    outSort        = sortPath.back();
    outDisplay     = displayPath.back();
    outPane        = paneMenu;

    HMENU historyMenu   = nullptr;
    const int paneCount = GetMenuItemCount(paneMenu);
    for (int pos = 0; pos < paneCount; ++pos)
    {
        HMENU subMenu = GetSubMenu(paneMenu, pos);
        if (! subMenu)
        {
            continue;
        }

        if (subMenu == outSort || subMenu == outDisplay)
        {
            continue;
        }

        historyMenu = subMenu;
        break;
    }

    outHistory = historyMenu;
}

static void EnsureMenuHandles(HWND hWnd) noexcept
{
    HMENU mainMenu = GetMenu(hWnd);
    if (! mainMenu)
    {
        mainMenu = g_mainMenuHandle;
    }
    if (! mainMenu)
    {
        return;
    }

    if (! g_mainMenuHandle)
    {
        g_mainMenuHandle = mainMenu;
    }

    if (! g_filesMenu)
    {
        std::vector<HMENU> filesPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_PANE_VIEW, filesPath) && ! filesPath.empty())
        {
            g_filesMenu = filesPath.back();
        }
    }

    if (! g_viewMenu)
    {
        std::vector<HMENU> viewPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_VIEW_MENUBAR, viewPath) && ! viewPath.empty())
        {
            g_viewMenu = viewPath.back();
        }
    }

    if (! g_editMenu)
    {
        std::vector<HMENU> editPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_PANE_CLIPBOARD_COPY, editPath) && ! editPath.empty())
        {
            g_editMenu = editPath.back();
        }
    }

    if (! g_editAdvancedMenu)
    {
        std::vector<HMENU> advancedPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_PANE_SAVE_SELECTION, advancedPath) && ! advancedPath.empty())
        {
            g_editAdvancedMenu = advancedPath.back();
        }
    }

    if (! g_viewThemeMenu)
    {
        std::vector<HMENU> themePath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_VIEW_THEME_SYSTEM, themePath) && ! themePath.empty())
        {
            g_viewThemeMenu = themePath.back();
        }
    }

    if (! g_viewPluginsMenu)
    {
        std::vector<HMENU> pluginsPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_VIEW_PLUGINS_MANAGE, pluginsPath) && ! pluginsPath.empty())
        {
            g_viewPluginsMenu = pluginsPath.back();
        }
    }

    if (g_filesMenu)
    {
        if (! g_viewWithMenu)
        {
            g_viewWithMenu = FindSubMenuAfterCommand(g_filesMenu, IDM_PANE_ALTERNATE_VIEW);
        }

        if (! g_editWithMenu)
        {
            g_editWithMenu = FindSubMenuAfterCommand(g_filesMenu, IDM_PANE_ALTERNATE_EDIT);
        }

        if (! g_newTemplateMenu)
        {
            g_newTemplateMenu = FindSubMenuAfterCommand(g_filesMenu, IDM_PANE_UNPACK);
        }
    }

    if (! g_userMenu)
    {
        g_userMenu = FindSubMenuAfterCommandRecursive(mainMenu, IDM_APP_REREAD_ASSOCIATIONS);
    }

    if (! g_openFileExplorerMenu)
    {
        std::vector<HMENU> explorerPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_PANE_OPEN_CURRENT_FOLDER, explorerPath) && ! explorerPath.empty())
        {
            g_openFileExplorerMenu = explorerPath.back();
        }
    }

    if (! g_leftPaneMenu || ! g_leftSortMenu || ! g_leftDisplayMenu || ! g_leftGoToMenu)
    {
        EnsurePaneMenuHandlesFor(mainMenu, IDM_LEFT_SORT_NAME, IDM_LEFT_DISPLAY_BRIEF, g_leftPaneMenu, g_leftSortMenu, g_leftDisplayMenu, g_leftGoToMenu);
    }
    if (! g_leftShowMenu)
    {
        std::vector<HMENU> showPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_LEFT_SHOW_HIDDEN_FILES, showPath) && ! showPath.empty())
        {
            g_leftShowMenu = showPath.back();
        }
    }

    if (! g_rightPaneMenu || ! g_rightSortMenu || ! g_rightDisplayMenu || ! g_rightGoToMenu)
    {
        EnsurePaneMenuHandlesFor(mainMenu, IDM_RIGHT_SORT_NAME, IDM_RIGHT_DISPLAY_BRIEF, g_rightPaneMenu, g_rightSortMenu, g_rightDisplayMenu, g_rightGoToMenu);
    }
    if (! g_rightShowMenu)
    {
        std::vector<HMENU> showPath;
        if (TryFindMenuPathToCommand(mainMenu, IDM_RIGHT_SHOW_HIDDEN_FILES, showPath) && ! showPath.empty())
        {
            g_rightShowMenu = showPath.back();
        }
    }

    if (! IsOverlaySampleEnabled())
    {
        RemoveOverlaySampleSubmenu(g_leftPaneMenu, IDM_LEFT_OVERLAY_SAMPLE_ERROR);
        RemoveOverlaySampleSubmenu(g_rightPaneMenu, IDM_RIGHT_OVERLAY_SAMPLE_ERROR);
    }
}

static void RebuildPluginsMenuDynamicItems(HWND hWnd)
{
    if (! g_viewPluginsMenu)
    {
        return;
    }

    const int managePos = FindMenuItemPosById(g_viewPluginsMenu, IDM_VIEW_PLUGINS_MANAGE);
    if (managePos < 0)
    {
        return;
    }

    DeleteMenuItemsFromPosition(g_viewPluginsMenu, managePos + 1);

    g_pluginMenuIdToPluginId.clear();
    g_pluginIdToMenuId.clear();

    if (! AppendMenuW(g_viewPluginsMenu, MF_SEPARATOR, 0, nullptr))
    {
        return;
    }

    const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();

    std::vector<const FileSystemPluginManager::PluginEntry*> embedded;
    std::vector<const FileSystemPluginManager::PluginEntry*> optional;
    std::vector<const FileSystemPluginManager::PluginEntry*> custom;

    embedded.reserve(plugins.size());
    optional.reserve(plugins.size());
    custom.reserve(plugins.size());

    for (const auto& entry : plugins)
    {
        switch (entry.origin)
        {
            case FileSystemPluginManager::PluginOrigin::Embedded: embedded.push_back(&entry); break;
            case FileSystemPluginManager::PluginOrigin::Optional: optional.push_back(&entry); break;
            case FileSystemPluginManager::PluginOrigin::Custom: custom.push_back(&entry); break;
        }
    }

    const auto byNameThenId = [](const FileSystemPluginManager::PluginEntry* a, const FileSystemPluginManager::PluginEntry* b)
    {
        const std::wstring an = (a && ! a->name.empty()) ? a->name : (a ? a->path.filename().wstring() : std::wstring());
        const std::wstring bn = (b && ! b->name.empty()) ? b->name : (b ? b->path.filename().wstring() : std::wstring());

        const int cmp = _wcsicmp(an.c_str(), bn.c_str());
        if (cmp != 0)
        {
            return cmp < 0;
        }

        const std::wstring aid = a ? a->id : std::wstring();
        const std::wstring bid = b ? b->id : std::wstring();
        return aid < bid;
    };

    std::sort(embedded.begin(), embedded.end(), byNameThenId);
    std::sort(optional.begin(), optional.end(), byNameThenId);
    std::sort(custom.begin(), custom.end(), byNameThenId);

    const std::wstring disabledSuffix    = LoadStringResource(nullptr, IDS_PLUGIN_SUFFIX_DISABLED);
    const std::wstring unavailableSuffix = LoadStringResource(nullptr, IDS_PLUGIN_SUFFIX_UNAVAILABLE);

    UINT nextId   = kPluginMenuIdFirst;
    bool wroteAny = false;

    const auto appendSection = [&](const std::vector<const FileSystemPluginManager::PluginEntry*>& items) -> bool
    {
        for (const auto* entry : items)
        {
            if (! entry)
            {
                continue;
            }

            if (nextId > kPluginMenuIdLast)
            {
                break;
            }

            std::wstring label = entry->name.empty() ? entry->path.filename().wstring() : entry->name;
            label              = EscapeMenuLabel(label);

            if (! entry->loadable && ! unavailableSuffix.empty())
            {
                label.append(L" ");
                label.append(unavailableSuffix);
            }
            else if (entry->disabled && ! disabledSuffix.empty())
            {
                label.append(L" ");
                label.append(disabledSuffix);
            }

            UINT flags = MF_STRING;
            if (entry->disabled || ! entry->loadable || entry->id.empty())
            {
                flags |= MF_GRAYED;
            }

            if (! AppendMenuW(g_viewPluginsMenu, flags, nextId, label.c_str()))
            {
                return false;
            }

            if (! entry->id.empty())
            {
                g_pluginMenuIdToPluginId[nextId] = entry->id;
                g_pluginIdToMenuId[entry->id]    = nextId;
            }

            ++nextId;
            wroteAny = true;
        }

        return true;
    };

    const auto appendSeparatorIf =
        [&](const std::vector<const FileSystemPluginManager::PluginEntry*>& current, const std::vector<const FileSystemPluginManager::PluginEntry*>& next)
    {
        if (current.empty() || next.empty())
        {
            return;
        }

        AppendMenuW(g_viewPluginsMenu, MF_SEPARATOR, 0, nullptr);
    };

    if (! appendSection(embedded))
    {
        return;
    }
    appendSeparatorIf(embedded, optional);
    if (! appendSection(optional))
    {
        return;
    }
    appendSeparatorIf(optional, custom);
    if (! appendSection(custom))
    {
        return;
    }

    if (! wroteAny)
    {
        const std::wstring emptyLabel = LoadStringResource(nullptr, IDS_MENU_EMPTY);
        AppendMenuW(g_viewPluginsMenu, MF_STRING | MF_GRAYED, 0, emptyLabel.empty() ? L"" : emptyLabel.c_str());
    }

    UpdatePluginsMenuChecks();
    DrawMenuBar(hWnd);
}

static void RebuildThemeMenuDynamicItems(HWND hWnd)
{
    if (! g_viewThemeMenu)
    {
        return;
    }

    const int lastBuiltInPos = FindMenuItemPosById(g_viewThemeMenu, IDM_VIEW_THEME_HIGH_CONTRAST_APP);
    if (lastBuiltInPos < 0)
    {
        return;
    }

    DeleteMenuItemsFromPosition(g_viewThemeMenu, lastBuiltInPos + 1);

    g_customThemeMenuIdToThemeId.clear();
    g_customThemeIdToMenuId.clear();

    const auto appendThemeNavigation = []() -> bool
    {
        if (! AppendMenuW(g_viewThemeMenu, MF_SEPARATOR, 0, nullptr))
        {
            return false;
        }

        std::wstring previousTheme = LoadStringResource(nullptr, IDS_CMD_THEME_SELECT_PREV);
        if (previousTheme.empty())
        {
            previousTheme = L"Previous Theme";
        }
        previousTheme = EscapeMenuLabel(previousTheme);
        if (! AppendMenuW(g_viewThemeMenu, MF_STRING, IDM_VIEW_THEME_PREV, previousTheme.c_str()))
        {
            return false;
        }

        std::wstring nextTheme = LoadStringResource(nullptr, IDS_CMD_THEME_SELECT_NEXT);
        if (nextTheme.empty())
        {
            nextTheme = L"Next Theme";
        }
        nextTheme = EscapeMenuLabel(nextTheme);
        return AppendMenuW(g_viewThemeMenu, MF_STRING, IDM_VIEW_THEME_NEXT, nextTheme.c_str()) != FALSE;
    };

    const CustomThemeGroups customThemes = CollectCustomThemeGroups();

    if (! customThemes.fileThemes.empty() || ! customThemes.settingsThemes.empty())
    {
        if (! AppendMenuW(g_viewThemeMenu, MF_SEPARATOR, 0, nullptr))
        {
            return;
        }

        UINT nextId = kCustomThemeMenuIdFirst;

        const auto addThemes = [&](const std::vector<const Common::Settings::ThemeDefinition*>& themes)
        {
            for (const auto* def : themes)
            {
                if (nextId > kCustomThemeMenuIdLast)
                {
                    break;
                }

                std::wstring label = def->name.empty() ? def->id : def->name;
                label              = EscapeMenuLabel(label);

                if (! AppendMenuW(g_viewThemeMenu, MF_STRING, nextId, label.c_str()))
                {
                    return false;
                }

                g_customThemeMenuIdToThemeId[nextId] = def->id;
                g_customThemeIdToMenuId[def->id]     = nextId;
                ++nextId;
            }

            return true;
        };

        if (! addThemes(customThemes.fileThemes))
        {
            return;
        }

        if (! customThemes.fileThemes.empty() && ! customThemes.settingsThemes.empty() && nextId <= kCustomThemeMenuIdLast)
        {
            if (! AppendMenuW(g_viewThemeMenu, MF_SEPARATOR, 0, nullptr))
            {
                return;
            }
        }

        if (! addThemes(customThemes.settingsThemes))
        {
            return;
        }
    }

    if (! appendThemeNavigation())
    {
        return;
    }

    DrawMenuBar(hWnd);
}

void UpdatePaneMenuChecks() noexcept
{
    auto updateSortDisplay = [](FolderWindow::Pane pane, HMENU sortMenu, HMENU displayMenu, UINT sortBase, UINT displayBase)
    {
        const FolderView::SortBy sortBy = g_folderWindow.GetSortBy(pane);
        const UINT sortLast             = sortBase + static_cast<UINT>(FolderView::SortBy::None);
        const UINT sortChecked          = sortBase + static_cast<UINT>(sortBy);
        CheckMenuRadioItem(sortMenu, sortBase, sortLast, sortChecked, MF_BYCOMMAND);

        const FolderView::DisplayMode display = g_folderWindow.GetDisplayMode(pane);
        const UINT displayChecked             = displayBase + static_cast<UINT>(display);
        CheckMenuRadioItem(displayMenu, displayBase, displayBase + 3, displayChecked, MF_BYCOMMAND);
    };

    if (g_leftSortMenu && g_leftDisplayMenu)
    {
        updateSortDisplay(FolderWindow::Pane::Left, g_leftSortMenu, g_leftDisplayMenu, IDM_LEFT_SORT_NAME, IDM_LEFT_DISPLAY_BRIEF);
    }
    if (g_rightSortMenu && g_rightDisplayMenu)
    {
        updateSortDisplay(FolderWindow::Pane::Right, g_rightSortMenu, g_rightDisplayMenu, IDM_RIGHT_SORT_NAME, IDM_RIGHT_DISPLAY_BRIEF);
    }

    if (g_leftGoToMenu)
    {
        const UINT backFlags = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.CanHistoryBack(FolderWindow::Pane::Left) ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(g_leftGoToMenu, IDM_LEFT_GO_TO_BACK, backFlags);

        const UINT forwardFlags = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.CanHistoryForward(FolderWindow::Pane::Left) ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(g_leftGoToMenu, IDM_LEFT_GO_TO_FORWARD, forwardFlags);
    }

    if (g_rightGoToMenu)
    {
        const UINT backFlags = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.CanHistoryBack(FolderWindow::Pane::Right) ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(g_rightGoToMenu, IDM_RIGHT_GO_TO_BACK, backFlags);

        const UINT forwardFlags = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.CanHistoryForward(FolderWindow::Pane::Right) ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(g_rightGoToMenu, IDM_RIGHT_GO_TO_FORWARD, forwardFlags);
    }

    const UINT leftStatusCheck  = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Left) ? MF_CHECKED : MF_UNCHECKED));
    const UINT rightStatusCheck = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Right) ? MF_CHECKED : MF_UNCHECKED));
    const UINT leftFileExtensionsCheck =
        static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetFileExtensionsVisible(FolderWindow::Pane::Left) ? MF_CHECKED : MF_UNCHECKED));
    const UINT rightFileExtensionsCheck =
        static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetFileExtensionsVisible(FolderWindow::Pane::Right) ? MF_CHECKED : MF_UNCHECKED));
    const UINT leftPreviewPaneCheck =
        static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.IsPreviewPaneOpenForSource(FolderWindow::Pane::Left) ? MF_CHECKED : MF_UNCHECKED));
    const UINT rightPreviewPaneCheck =
        static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.IsPreviewPaneOpenForSource(FolderWindow::Pane::Right) ? MF_CHECKED : MF_UNCHECKED));
    const UINT leftFilterBarCheck =
        static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetFilterBarVisible(FolderWindow::Pane::Left) ? MF_CHECKED : MF_UNCHECKED));
    const UINT rightFilterBarCheck =
        static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetFilterBarVisible(FolderWindow::Pane::Right) ? MF_CHECKED : MF_UNCHECKED));
    const UINT leftNavigationCheck =
        static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Left) ? MF_CHECKED : MF_UNCHECKED));
    const UINT rightNavigationCheck =
        static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Right) ? MF_CHECKED : MF_UNCHECKED));
    const UINT hiddenFilesCheck = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetShowHiddenFiles() ? MF_CHECKED : MF_UNCHECKED));
    const UINT systemFilesCheck = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.GetShowSystemFiles() ? MF_CHECKED : MF_UNCHECKED));

    if (g_leftPaneMenu)
    {
        CheckMenuItem(g_leftPaneMenu, IDM_LEFT_PREVIEW_PANE, leftPreviewPaneCheck);
    }

    if (g_rightPaneMenu)
    {
        CheckMenuItem(g_rightPaneMenu, IDM_RIGHT_PREVIEW_PANE, rightPreviewPaneCheck);
    }

    if (g_leftShowMenu)
    {
        CheckMenuItem(g_leftShowMenu, IDM_LEFT_SHOW_HIDDEN_FILES, hiddenFilesCheck);
        CheckMenuItem(g_leftShowMenu, IDM_LEFT_SHOW_SYSTEM_FILES, systemFilesCheck);
        CheckMenuItem(g_leftShowMenu, IDM_LEFT_SHOW_FILE_EXTENSIONS, leftFileExtensionsCheck);
        CheckMenuItem(g_leftShowMenu, IDM_LEFT_FILTER_BAR, leftFilterBarCheck);
        CheckMenuItem(g_leftShowMenu, IDM_LEFT_NAVIGATION_BAR, leftNavigationCheck);
        CheckMenuItem(g_leftShowMenu, IDM_LEFT_STATUSBAR, leftStatusCheck);
    }

    if (g_rightShowMenu)
    {
        CheckMenuItem(g_rightShowMenu, IDM_RIGHT_SHOW_HIDDEN_FILES, hiddenFilesCheck);
        CheckMenuItem(g_rightShowMenu, IDM_RIGHT_SHOW_SYSTEM_FILES, systemFilesCheck);
        CheckMenuItem(g_rightShowMenu, IDM_RIGHT_SHOW_FILE_EXTENSIONS, rightFileExtensionsCheck);
        CheckMenuItem(g_rightShowMenu, IDM_RIGHT_FILTER_BAR, rightFilterBarCheck);
        CheckMenuItem(g_rightShowMenu, IDM_RIGHT_NAVIGATION_BAR, rightNavigationCheck);
        CheckMenuItem(g_rightShowMenu, IDM_RIGHT_STATUSBAR, rightStatusCheck);
    }

    if (g_viewMenu)
    {
        const UINT menuBarCheck = static_cast<UINT>(MF_BYCOMMAND | (g_menuBarVisible ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(g_viewMenu, IDM_VIEW_MENUBAR, menuBarCheck);

        const UINT functionBarCheck = static_cast<UINT>(MF_BYCOMMAND | (g_functionBarVisible ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(g_viewMenu, IDM_VIEW_FUNCTIONBAR, functionBarCheck);

        const UINT issuesPaneCheck = static_cast<UINT>(MF_BYCOMMAND | (g_folderWindow.IsFileOperationsIssuesPaneVisible() ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(g_viewMenu, IDM_VIEW_FILEOPS_FAILED_ITEMS, issuesPaneCheck);
    }

    if (g_editMenu)
    {
        const bool canRestore   = g_folderWindow.HasSavedSelection();
        const UINT restoreFlags = static_cast<UINT>(MF_BYCOMMAND | (canRestore ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(g_editMenu, IDM_PANE_SELECTION_RESTORE, restoreFlags);
    }

    if (g_editAdvancedMenu)
    {
        const bool canRestore   = g_folderWindow.HasSavedSelection();
        const UINT restoreFlags = static_cast<UINT>(MF_BYCOMMAND | (canRestore ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(g_editAdvancedMenu, IDM_PANE_SELECTION_RESTORE, restoreFlags);

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        const bool canShowHidden      = g_folderWindow.CanShowHiddenNames(pane);
        const UINT showHiddenFlags    = static_cast<UINT>(MF_BYCOMMAND | (canShowHidden ? MF_ENABLED : MF_GRAYED));
        EnableMenuItem(g_editAdvancedMenu, IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES, showHiddenFlags);
    }
}

void ShowSortMenuPopup(HWND hWnd, FolderWindow::Pane pane, POINT screenPoint) noexcept
{
    const auto loadLabel = [](UINT stringId, std::wstring_view fallback) noexcept -> std::wstring
    {
        std::wstring text = LoadStringResource(nullptr, stringId);
        if (text.empty())
        {
            text.assign(fallback);
        }
        return text;
    };

    EnsureMenuHandles(hWnd);
    UpdatePaneMenuChecks();

    const bool isLeft = pane == FolderWindow::Pane::Left;
    const UINT idName = isLeft ? IDM_LEFT_SORT_NAME : IDM_RIGHT_SORT_NAME;
    const UINT idExt  = isLeft ? IDM_LEFT_SORT_EXTENSION : IDM_RIGHT_SORT_EXTENSION;
    const UINT idTime = isLeft ? IDM_LEFT_SORT_TIME : IDM_RIGHT_SORT_TIME;
    const UINT idSize = isLeft ? IDM_LEFT_SORT_SIZE : IDM_RIGHT_SORT_SIZE;
    const UINT idAttr = isLeft ? IDM_LEFT_SORT_ATTRIBUTES : IDM_RIGHT_SORT_ATTRIBUTES;
    const UINT idNone = isLeft ? IDM_LEFT_SORT_NONE : IDM_RIGHT_SORT_NONE;

    UINT checkedId = idNone;
    switch (g_folderWindow.GetSortBy(pane))
    {
        case FolderView::SortBy::Name: checkedId = idName; break;
        case FolderView::SortBy::Extension: checkedId = idExt; break;
        case FolderView::SortBy::Time: checkedId = idTime; break;
        case FolderView::SortBy::Size: checkedId = idSize; break;
        case FolderView::SortBy::Attributes: checkedId = idAttr; break;
        case FolderView::SortBy::None: checkedId = idNone; break;
    }

    auto makeRadioItem = [&](UINT commandId, UINT stringId, std::wstring_view fallback) noexcept
    {
        RedSalamander::DxUi::MenuFlyoutItem item{};
        item.kind      = RedSalamander::DxUi::MenuItemKind::Radio;
        item.text      = loadLabel(stringId, fallback);
        item.commandId = static_cast<int>(commandId);
        item.checked   = commandId == checkedId;
        return item;
    };

    const std::array<RedSalamander::DxUi::MenuFlyoutItem, 6> items{
        makeRadioItem(idNone, IDS_PREFS_PANES_SORT_NONE, L"None"),
        makeRadioItem(idName, IDS_PREFS_PANES_SORT_NAME, L"Name"),
        makeRadioItem(idExt, IDS_PREFS_PANES_SORT_EXTENSION, L"Extension"),
        makeRadioItem(idTime, IDS_PREFS_PANES_SORT_TIME, L"Time"),
        makeRadioItem(idSize, IDS_PREFS_PANES_SORT_SIZE, L"Size"),
        makeRadioItem(idAttr, IDS_PREFS_PANES_SORT_ATTRIBUTES, L"Attributes"),
    };

    RedSalamander::DxUi::ContextMenuSessionCallbacks callbacks{};
    callbacks.rootHorizontalAlignment = RedSalamander::DxUi::ContextMenuRootHorizontalAlignment::End;
    callbacks.rootVerticalPlacement   = RedSalamander::DxUi::ContextMenuRootVerticalPlacement::Above;

    if (const auto result = RedSalamander::DxUi::ContextMenu::Show(hWnd, screenPoint, items, MakeAppThemeDxPalette(ResolveConfiguredTheme()), callbacks);
        result.has_value())
    {
        PostMessageW(hWnd, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(result.value()), 0), 0);
    }
}

void AppendEmptyMenuItem(HMENU menu) noexcept
{
    if (! menu)
    {
        return;
    }

    const std::wstring emptyLabel = LoadStringResource(nullptr, IDS_MENU_EMPTY);
    AppendMenuW(menu, MF_STRING | MF_GRAYED, 0, emptyLabel.empty() ? L"" : emptyLabel.c_str());
}

[[nodiscard]] std::wstring GetComputerNameText()
{
    wchar_t computerName[MAX_COMPUTERNAME_LENGTH + 1]{};
    DWORD computerNameLength = static_cast<DWORD>(std::size(computerName));
    if (GetComputerNameW(computerName, &computerNameLength) == FALSE || computerNameLength == 0u)
    {
        return {};
    }

    return std::wstring(computerName, computerNameLength);
}

[[nodiscard]] std::wstring GetFileActionMenuText(const Common::Settings::FileActionDefinition& action)
{
    if (! action.displayName.empty())
    {
        return action.displayName;
    }

    return action.id;
}

void RebuildFileActionMenuDynamicItems(HMENU menu, bool viewerActions)
{
    if (! menu)
    {
        return;
    }

    Debug::Perf::Scope perf(viewerActions ? L"fileaction.viewwith.menu_populate_us" : L"fileaction.editwith.menu_populate_us");

    DeleteMenuItemsFromPosition(menu, 0);
    auto& menuMap = viewerActions ? g_viewWithMenuIdToActionId : g_editWithMenuIdToActionId;
    menuMap.clear();

    const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
    const std::optional<std::filesystem::path> focusedPath = g_folderWindow.GetFocusedItemPath(pane);
    if (! focusedPath.has_value())
    {
        AppendEmptyMenuItem(menu);
        perf.SetValue0(0u);
        perf.SetValue1(0u);
        return;
    }

    const std::vector<Common::Settings::FileActionDefinition>& actionSettings =
        viewerActions ? g_settings.fileActions.viewers.actions : g_settings.fileActions.editors.actions;
    const std::wstring computerName = GetComputerNameText();
    const std::vector<const Common::Settings::FileActionDefinition*> actions =
        FileActionResolver::CollectApplicableActions(actionSettings, focusedPath.value(), computerName);

    const UINT firstId = viewerActions ? IDM_PANE_VIEW_WITH_BASE : IDM_PANE_EDIT_WITH_BASE;
    const UINT lastId  = viewerActions ? IDM_PANE_VIEW_WITH_LAST : IDM_PANE_EDIT_WITH_LAST;
    UINT nextId        = firstId;
    uint64_t written   = 0;

    for (const Common::Settings::FileActionDefinition* action : actions)
    {
        if (! action || action->id.empty() || nextId > lastId)
        {
            continue;
        }

        const std::wstring text = GetFileActionMenuText(*action);
        if (text.empty())
        {
            continue;
        }

        if (AppendMenuW(menu, MF_STRING, nextId, text.c_str()) != FALSE)
        {
            menuMap.emplace(nextId, action->id);
            ++nextId;
            ++written;
        }
    }

    if (written == 0)
    {
        AppendEmptyMenuItem(menu);
    }

    perf.SetValue0(static_cast<uint64_t>(actions.size()));
    perf.SetValue1(written);
}

void RebuildUserMenuDynamicItems(HMENU menu)
{
    if (! menu)
    {
        return;
    }

    Debug::Perf::Scope perf(L"usermenu.menu_populate_us");

    DeleteMenuItemsFromPosition(menu, 0);
    g_userMenuIdToActionId.clear();

    const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
    const std::vector<FolderWindow::UserMenuItem> items = g_folderWindow.CollectUserMenuItems(pane);

    UINT nextId      = IDM_PANE_USER_MENU_BASE;
    uint64_t written = 0;
    for (const FolderWindow::UserMenuItem& item : items)
    {
        if (item.id.empty() || item.displayName.empty() || nextId > IDM_PANE_USER_MENU_LAST)
        {
            continue;
        }

        const UINT flags = item.enabled ? MF_STRING : (MF_STRING | MF_GRAYED);
        if (AppendMenuW(menu, flags, nextId, item.displayName.c_str()) != FALSE)
        {
            g_userMenuIdToActionId.emplace(nextId, item.id);
            ++nextId;
            ++written;
        }
    }

    if (written == 0)
    {
        AppendEmptyMenuItem(menu);
    }

    perf.SetValue0(static_cast<uint64_t>(items.size()));
    perf.SetValue1(written);
}

[[nodiscard]] POINT GetUserMenuPopupAnchor(HWND ownerWindow) noexcept
{
    RECT rect{};
    if (const HWND focusedFolderView = g_folderWindow.GetFocusedFolderViewHwnd();
        focusedFolderView && GetWindowRect(focusedFolderView, &rect) != FALSE)
    {
        return POINT{rect.left + 8, rect.top + 8};
    }

    if (ownerWindow && GetWindowRect(ownerWindow, &rect) != FALSE)
    {
        return POINT{rect.left + 16, rect.top + GetSystemMetrics(SM_CYMENU) + 8};
    }

    POINT point{};
    static_cast<void>(GetCursorPos(&point));
    return point;
}

void ShowUserMenuPopup(HWND hWnd) noexcept
{
    EnsureMenuHandles(hWnd);
    if (! g_userMenu)
    {
        return;
    }

    RebuildUserMenuDynamicItems(g_userMenu);

    const POINT point = GetUserMenuPopupAnchor(hWnd);
    SetForegroundWindow(hWnd);
    const UINT commandId = static_cast<UINT>(
        TrackPopupMenuEx(g_userMenu, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, point.x, point.y, hWnd, nullptr));
    if (commandId != 0u)
    {
        SendMessageW(hWnd, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(commandId), 0), 0);
    }
}

void RebuildShellNewTemplateMenuDynamicItems(HMENU menu)
{
    if (! menu)
    {
        return;
    }

    Debug::Perf::Scope perf(L"shellnew.menu_populate_us");

    constexpr int kFirstDynamicPosition = 2;
    DeleteMenuItemsFromPosition(menu, kFirstDynamicPosition);
    g_newTemplateMenuIdToTemplateId.clear();

    const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
    const std::vector<FolderWindow::ShellNewTemplateMenuItem> items = g_folderWindow.CollectShellNewTemplateMenuItems(pane);

    UINT nextId      = IDM_PANE_NEW_TEMPLATE_BASE;
    uint64_t written = 0u;
    for (const FolderWindow::ShellNewTemplateMenuItem& item : items)
    {
        if (item.id.empty() || item.displayName.empty() || nextId > IDM_PANE_NEW_TEMPLATE_LAST)
        {
            continue;
        }

        if (AppendMenuW(menu, MF_STRING, nextId, item.displayName.c_str()) != FALSE)
        {
            g_newTemplateMenuIdToTemplateId.emplace(nextId, item.id);
            ++nextId;
            ++written;
        }
    }

    if (written == 0u)
    {
        AppendEmptyMenuItem(menu);
    }

    perf.SetValue0(static_cast<uint64_t>(items.size()));
    perf.SetValue1(written);
}

void RebuildGoToMenuDynamicItems(FolderWindow::Pane pane, HMENU goToMenu, UINT hotPathsCommandId, UINT hotPathBaseId, UINT historyBaseId)
{
    if (! goToMenu)
    {
        return;
    }
    const int hotPathsPos = FindMenuItemPosById(goToMenu, hotPathsCommandId);
    if (hotPathsPos < 0)
    {
        return;
    }

    // Reset everything after the fixed "Hot Paths..." entry (dynamic hot paths + separator + history).
    DeleteMenuItemsFromPosition(goToMenu, hotPathsPos + 1);

    // Append hot path items.
    UINT hotPathWritten = 0;
    if (g_settings.hotPaths.has_value())
    {
        const auto& slots = g_settings.hotPaths.value().slots;
        for (size_t i = 0; i < slots.size(); ++i)
        {
            if (! slots[i].has_value() || slots[i].value().path.empty())
            {
                continue;
            }

            const auto& slot = slots[i].value();
            const UINT id    = hotPathBaseId + static_cast<UINT>(i);

            g_navigatePathMenuTargets[id] = NavigatePathMenuTarget{pane, std::filesystem::path(slot.path)};

            const wchar_t digitChar = (i < 9) ? static_cast<wchar_t>(L'1' + i) : L'0';
            std::wstring label;
            if (! slot.label.empty())
            {
                label = std::format(L"&{}: {}", digitChar, EscapeMenuLabel(slot.label));
            }
            else
            {
                label = std::format(L"&{}: {}", digitChar, EscapeMenuLabel(slot.path));
            }

            AppendMenuW(goToMenu, MF_STRING, id, label.c_str());
            ++hotPathWritten;
        }
    }

    if (hotPathWritten == 0)
    {
        AppendEmptyMenuItem(goToMenu);
    }
    AppendMenuW(goToMenu, MF_SEPARATOR, 0, nullptr);

    const std::vector<std::filesystem::path> history   = g_folderWindow.GetFolderHistory(pane);
    const std::optional<std::filesystem::path> current = g_folderWindow.GetCurrentPath(pane);
    const std::wstring currentText                     = current.has_value() ? current.value().wstring() : std::wstring{};

    UINT written = 0;
    for (size_t i = 0; i < history.size() && written < kHistoryMenuMaxItems; ++i)
    {
        const std::filesystem::path& entry = history[i];
        if (entry.empty())
        {
            continue;
        }

        const UINT id                 = historyBaseId + written;
        g_navigatePathMenuTargets[id] = NavigatePathMenuTarget{pane, entry};

        UINT flags = MF_STRING;
        if (! currentText.empty() && _wcsicmp(entry.c_str(), currentText.c_str()) == 0)
        {
            flags |= MF_CHECKED;
        }

        std::wstring label = EscapeMenuLabel(entry.wstring());
        if (! AppendMenuW(goToMenu, flags, id, label.c_str()))
        {
            break;
        }

        ++written;
    }

    if (written == 0)
    {
        AppendEmptyMenuItem(goToMenu);
    }
}

[[nodiscard]] std::optional<std::wstring> TryGetKnownFolderDisplayName(REFKNOWNFOLDERID folderId) noexcept
{
    wil::com_ptr<IShellItem> folderItem;
    if (FAILED(SHGetKnownFolderItem(folderId, KF_FLAG_DEFAULT, nullptr, IID_PPV_ARGS(folderItem.put()))) || ! folderItem)
    {
        return std::nullopt;
    }

    wil::unique_cotaskmem_string rawName;
    if (FAILED(folderItem->GetDisplayName(SIGDN_NORMALDISPLAY, rawName.put())) || ! rawName)
    {
        return std::nullopt;
    }

    return EscapeMenuLabel(rawName.get());
}

[[nodiscard]] HBITMAP GetOrCreateMainMenuIconBitmap(UINT menuCommandId, const GUID& knownFolderId) noexcept
{
    const auto it = g_mainMenuIconBitmaps.find(menuCommandId);
    if (it != g_mainMenuIconBitmaps.end() && it->second)
    {
        return it->second.get();
    }

    int iconSize = GetSystemMetrics(SM_CXSMICON);
    if (iconSize <= 0)
    {
        iconSize = 16;
    }

    const auto iconIndex = IconCache::GetInstance().QuerySysIconIndexForKnownFolder(knownFolderId);
    if (! iconIndex.has_value())
    {
        return nullptr;
    }

    wil::unique_hbitmap bitmap = IconCache::GetInstance().CreateMenuBitmapFromIconIndex(iconIndex.value(), iconSize);
    if (! bitmap)
    {
        return nullptr;
    }

    auto existing = g_mainMenuIconBitmaps.find(menuCommandId);
    if (existing != g_mainMenuIconBitmaps.end())
    {
        existing->second = std::move(bitmap);
        return existing->second.get();
    }

    const HBITMAP rawBitmap = bitmap.get();
    g_mainMenuIconBitmaps.emplace(menuCommandId, std::move(bitmap));
    return rawBitmap;
}

void UpdateOpenFileExplorerMenuStockFolders(HMENU openExplorerMenu) noexcept
{
    if (! openExplorerMenu)
    {
        return;
    }

    struct FolderItem
    {
        UINT menuId;
        const GUID* knownFolderId;
    };

    static constexpr std::array<FolderItem, 7> kFolders = {
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_DESKTOP, &FOLDERID_Desktop},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_DOCUMENTS, &FOLDERID_Documents},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_DOWNLOADS, &FOLDERID_Downloads},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_PICTURES, &FOLDERID_Pictures},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_MUSIC, &FOLDERID_Music},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_VIDEOS, &FOLDERID_Videos},
        FolderItem{IDM_APP_OPEN_FILE_EXPLORER_ONEDRIVE, &kKnownFolderIdOneDrive},
    };

    for (const FolderItem& entry : kFolders)
    {
        const std::optional<std::wstring> nameOpt = TryGetKnownFolderDisplayName(*entry.knownFolderId);
        if (! nameOpt.has_value())
        {
            EnableMenuItem(openExplorerMenu, entry.menuId, MF_BYCOMMAND | MF_GRAYED);
            continue;
        }

        std::wstring name = nameOpt.value();

        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize     = sizeof(itemInfo);
        itemInfo.fMask      = MIIM_STRING;
        itemInfo.dwTypeData = name.data();
        itemInfo.cch        = static_cast<UINT>(name.size());
        SetMenuItemInfoW(openExplorerMenu, entry.menuId, FALSE, &itemInfo);

        static_cast<void>(GetOrCreateMainMenuIconBitmap(entry.menuId, *entry.knownFolderId));

        EnableMenuItem(openExplorerMenu, entry.menuId, MF_BYCOMMAND | MF_ENABLED);
    }
}

void OnInitMenuPopup(HWND hWnd, HMENU menu)
{
    if (! menu)
    {
        return;
    }

    EnsureMenuHandles(hWnd);

    if (menu == g_viewWithMenu)
    {
        RebuildFileActionMenuDynamicItems(menu, true);
        return;
    }

    if (menu == g_editWithMenu)
    {
        RebuildFileActionMenuDynamicItems(menu, false);
        return;
    }

    if (menu == g_userMenu)
    {
        RebuildUserMenuDynamicItems(menu);
        return;
    }

    if (menu == g_newTemplateMenu)
    {
        RebuildShellNewTemplateMenuDynamicItems(menu);
        return;
    }

    if (menu == g_viewPluginsMenu)
    {
        RebuildPluginsMenuDynamicItems(hWnd);
        return;
    }

    if (menu == g_openFileExplorerMenu)
    {
        UpdateOpenFileExplorerMenuStockFolders(menu);
        return;
    }

    if (menu != g_leftGoToMenu && menu != g_rightGoToMenu && menu != g_leftSortMenu && menu != g_rightSortMenu && menu != g_leftDisplayMenu &&
        menu != g_rightDisplayMenu && menu != g_leftShowMenu && menu != g_rightShowMenu && menu != g_viewMenu && menu != g_editMenu && menu != g_editAdvancedMenu &&
        menu != g_leftPaneMenu && menu != g_rightPaneMenu)
    {
        return;
    }

    g_navigatePathMenuTargets.clear();
    RebuildGoToMenuDynamicItems(FolderWindow::Pane::Left, g_leftGoToMenu, IDM_LEFT_HOT_PATHS, IDM_LEFT_HOT_PATH_BASE, IDM_LEFT_HISTORY_BASE);
    RebuildGoToMenuDynamicItems(FolderWindow::Pane::Right, g_rightGoToMenu, IDM_RIGHT_HOT_PATHS, IDM_RIGHT_HOT_PATH_BASE, IDM_RIGHT_HISTORY_BASE);
    UpdatePaneMenuChecks();
}

void SplitMenuText(std::wstring_view raw, std::wstring& text, std::wstring& shortcut)
{
    const size_t tabPos = raw.find(L'\t');
    if (tabPos == std::wstring_view::npos)
    {
        text.assign(raw);
        shortcut.clear();
        return;
    }

    text.assign(raw.substr(0, tabPos));
    shortcut.assign(raw.substr(tabPos + 1));
}

[[nodiscard]] std::wstring StripMenuMnemonicMarkers(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());

    for (size_t index = 0; index < text.size(); ++index)
    {
        if (text[index] != L'&')
        {
            result.push_back(text[index]);
            continue;
        }

        if ((index + 1u) < text.size() && text[index + 1u] == L'&')
        {
            result.push_back(L'&');
            ++index;
        }
    }

    return result;
}

[[nodiscard]] wchar_t FindMenuMnemonic(std::wstring_view text) noexcept
{
    for (size_t index = 0; index < text.size(); ++index)
    {
        if (text[index] != L'&')
        {
            continue;
        }

        if ((index + 1u) >= text.size())
        {
            break;
        }

        if (text[index + 1u] == L'&')
        {
            ++index;
            continue;
        }

        return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(text[index + 1u])));
    }

    return L'\0';
}

[[nodiscard]] std::wstring VkToMenuShortcutText(uint32_t vk) noexcept
{
    vk &= 0xFFu;

    if (vk >= VK_F1 && vk <= VK_F24)
    {
        return std::format(L"F{}", static_cast<unsigned>(vk - VK_F1 + 1));
    }

    if ((vk >= static_cast<uint32_t>('0') && vk <= static_cast<uint32_t>('9')) || (vk >= static_cast<uint32_t>('A') && vk <= static_cast<uint32_t>('Z')))
    {
        wchar_t buf[2]{};
        buf[0] = static_cast<wchar_t>(vk);
        buf[1] = L'\0';
        return buf;
    }

    UINT scanCode = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    if (scanCode == 0)
    {
        return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
    }

    bool extended = false;
    switch (vk)
    {
        case VK_LEFT:
        case VK_UP:
        case VK_RIGHT:
        case VK_DOWN:
        case VK_PRIOR:
        case VK_NEXT:
        case VK_END:
        case VK_HOME:
        case VK_INSERT:
        case VK_DELETE: extended = true; break;
    }

    LPARAM lParam = static_cast<LPARAM>(scanCode) << 16;
    if (extended)
    {
        lParam |= (1 << 24);
    }

    wchar_t keyName[64]{};
    const int length = GetKeyNameTextW(static_cast<LONG>(lParam), keyName, static_cast<int>(std::size(keyName)));
    if (length > 0)
    {
        return std::wstring(keyName, static_cast<size_t>(length));
    }

    return std::format(L"VK_{:02X}", static_cast<unsigned>(vk));
}

[[nodiscard]] std::wstring FormatMenuChordText(uint32_t vk, uint32_t modifiers) noexcept
{
    std::wstring result;
    auto appendPart = [&](std::wstring_view part)
    {
        if (part.empty())
        {
            return;
        }
        if (! result.empty())
        {
            result.append(L"+");
        }
        result.append(part);
    };

    const uint32_t maskedMods = modifiers & 0x7u;
    if ((maskedMods & ShortcutManager::kModCtrl) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_CTRL));
    }
    if ((maskedMods & ShortcutManager::kModAlt) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_ALT));
    }
    if ((maskedMods & ShortcutManager::kModShift) != 0)
    {
        appendPart(LoadStringResource(nullptr, IDS_MOD_SHIFT));
    }

    appendPart(VkToMenuShortcutText(vk));
    return result;
}

[[nodiscard]] std::optional<std::wstring_view> TryGetCommandIdForMenuShortcut(UINT menuCommandId) noexcept
{
    if (menuCommandId == 0)
    {
        return std::nullopt;
    }

    if (const CommandInfo* info = FindCommandInfoByWmCommandId(menuCommandId))
    {
        return info->id;
    }

    switch (menuCommandId)
    {
        case IDM_LEFT_CHANGE_DRIVE: return L"cmd/app/openLeftDriveMenu";
        case IDM_RIGHT_CHANGE_DRIVE: return L"cmd/app/openRightDriveMenu";

        case IDM_LEFT_GO_TO_BACK:
        case IDM_RIGHT_GO_TO_BACK: return L"cmd/pane/historyBack";
        case IDM_LEFT_GO_TO_FORWARD:
        case IDM_RIGHT_GO_TO_FORWARD: return L"cmd/pane/historyForward";
        case IDM_LEFT_GO_TO_PARENT_DIRECTORY:
        case IDM_RIGHT_GO_TO_PARENT_DIRECTORY: return L"cmd/pane/upOneDirectory";
        case IDM_LEFT_GO_TO_ROOT_DIRECTORY:
        case IDM_RIGHT_GO_TO_ROOT_DIRECTORY: return L"cmd/pane/goRootDirectory";
        case IDM_LEFT_GO_TO_PATH_FROM_OTHER_PANE:
        case IDM_RIGHT_GO_TO_PATH_FROM_OTHER_PANE: return L"cmd/pane/setPathFromOtherPane";
        case IDM_LEFT_HOT_PATHS:
        case IDM_RIGHT_HOT_PATHS: return L"cmd/pane/hotPaths";

        case IDM_LEFT_DISPLAY_BRIEF:
        case IDM_RIGHT_DISPLAY_BRIEF: return L"cmd/pane/display/brief";
        case IDM_LEFT_DISPLAY_DETAILED:
        case IDM_RIGHT_DISPLAY_DETAILED: return L"cmd/pane/display/detailed";
        case IDM_LEFT_DISPLAY_EXTRA_DETAILED:
        case IDM_RIGHT_DISPLAY_EXTRA_DETAILED: return L"cmd/pane/display/extraDetailed";
        case IDM_LEFT_DISPLAY_THUMBNAILS:
        case IDM_RIGHT_DISPLAY_THUMBNAILS: return L"cmd/pane/viewOptions/toggleThumbnails";
        case IDM_LEFT_PREVIEW_PANE:
        case IDM_RIGHT_PREVIEW_PANE: return L"cmd/pane/viewOptions/togglePreviewPane";

        case IDM_LEFT_SORT_NONE:
        case IDM_RIGHT_SORT_NONE: return L"cmd/pane/sort/none";
        case IDM_LEFT_SORT_NAME:
        case IDM_RIGHT_SORT_NAME: return L"cmd/pane/sort/name";
        case IDM_LEFT_SORT_EXTENSION:
        case IDM_RIGHT_SORT_EXTENSION: return L"cmd/pane/sort/extension";
        case IDM_LEFT_SORT_TIME:
        case IDM_RIGHT_SORT_TIME: return L"cmd/pane/sort/time";
        case IDM_LEFT_SORT_SIZE:
        case IDM_RIGHT_SORT_SIZE: return L"cmd/pane/sort/size";
        case IDM_LEFT_SORT_ATTRIBUTES:
        case IDM_RIGHT_SORT_ATTRIBUTES: return L"cmd/pane/sort/attributes";

        case IDM_LEFT_ZOOM_PANEL:
        case IDM_RIGHT_ZOOM_PANEL: return L"cmd/pane/zoomPanel";
        case IDM_LEFT_SHOW_HIDDEN_FILES:
        case IDM_RIGHT_SHOW_HIDDEN_FILES: return L"cmd/pane/viewOptions/toggleHiddenFiles";
        case IDM_LEFT_SHOW_SYSTEM_FILES:
        case IDM_RIGHT_SHOW_SYSTEM_FILES: return L"cmd/pane/viewOptions/toggleSystemFiles";
        case IDM_LEFT_SHOW_FILE_EXTENSIONS:
        case IDM_RIGHT_SHOW_FILE_EXTENSIONS: return L"cmd/pane/viewOptions/toggleFileExtensions";
        case IDM_LEFT_FILTER:
        case IDM_RIGHT_FILTER: return L"cmd/pane/filter";
        case IDM_LEFT_REFRESH:
        case IDM_RIGHT_REFRESH: return L"cmd/pane/refresh";
        case IDM_LEFT_FILTER_BAR:
        case IDM_RIGHT_FILTER_BAR: return L"cmd/pane/viewOptions/toggleFilterBar";
        case IDM_LEFT_NAVIGATION_BAR:
        case IDM_RIGHT_NAVIGATION_BAR: return L"cmd/pane/viewOptions/toggleNavigationBar";
        case IDM_LEFT_STATUSBAR:
        case IDM_RIGHT_STATUSBAR: return L"cmd/pane/viewOptions/toggleStatusBar";
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<std::wstring> TryGetShortcutTextForCommandId(std::wstring_view commandId) noexcept
{
    if (commandId.empty())
    {
        return std::nullopt;
    }

    if (g_settings.shortcuts.has_value())
    {
        const Common::Settings::ShortcutsSettings& shortcuts = g_settings.shortcuts.value();

        const auto findBinding = [&](const std::vector<Common::Settings::ShortcutBinding>& bindings) -> std::optional<std::wstring>
        {
            for (const auto& binding : bindings)
            {
                if (binding.commandId.empty())
                {
                    continue;
                }

                if (std::wstring_view(binding.commandId) != commandId)
                {
                    continue;
                }

                return FormatMenuChordText(binding.vk, binding.modifiers);
            }

            return std::nullopt;
        };

        if (std::optional<std::wstring> found = findBinding(shortcuts.functionBar))
        {
            return found;
        }

        if (std::optional<std::wstring> found = findBinding(shortcuts.folderView))
        {
            return found;
        }
    }

    return std::nullopt;
}

[[nodiscard]] bool TryGetMenuItemPresentationText(HMENU menu, UINT position, std::wstring& outText, std::wstring& outShortcut) noexcept
{
    outText.clear();
    outShortcut.clear();

    std::array<wchar_t, 512> buffer{};
    const int length = GetMenuStringW(menu, position, buffer.data(), static_cast<int>(buffer.size()), MF_BYPOSITION);
    if (length <= 0)
    {
        return false;
    }

    SplitMenuText(std::wstring_view(buffer.data(), static_cast<size_t>(length)), outText, outShortcut);
    return true;
}

[[nodiscard]] std::optional<std::wstring> TryGetDynamicMenuShortcutText(UINT menuCommandId) noexcept
{
    if (const std::optional<std::wstring_view> commandIdOpt = TryGetCommandIdForMenuShortcut(menuCommandId); commandIdOpt.has_value())
    {
        return TryGetShortcutTextForCommandId(commandIdOpt.value());
    }

    if ((menuCommandId >= IDM_LEFT_HOT_PATH_BASE && menuCommandId < (IDM_LEFT_HOT_PATH_BASE + 10u)) ||
        (menuCommandId >= IDM_RIGHT_HOT_PATH_BASE && menuCommandId < (IDM_RIGHT_HOT_PATH_BASE + 10u)))
    {
        const UINT base    = (menuCommandId >= IDM_RIGHT_HOT_PATH_BASE) ? IDM_RIGHT_HOT_PATH_BASE : IDM_LEFT_HOT_PATH_BASE;
        const UINT slotIdx = menuCommandId - base;

        const wchar_t digitChar = (slotIdx < 9u) ? static_cast<wchar_t>(L'1' + slotIdx) : L'0';
        std::wstring commandId  = L"cmd/pane/hotPath/";
        commandId.push_back(digitChar);
        return TryGetShortcutTextForCommandId(commandId);
    }

    return std::nullopt;
}

[[nodiscard]] std::vector<RedSalamander::DxUi::MenuFlyoutItem> ConvertHMenuToDxFlyoutItems(HMENU menu) noexcept
{
    using RedSalamander::DxUi::MenuFlyoutItem;
    using RedSalamander::DxUi::MenuItemKind;

    std::vector<MenuFlyoutItem> items;
    if (! menu)
    {
        return items;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return items;
    }

    items.reserve(static_cast<size_t>(itemCount));
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, position, TRUE, &itemInfo))
        {
            continue;
        }

        MenuFlyoutItem item{};
        if ((itemInfo.fType & MFT_SEPARATOR) != 0)
        {
            item.kind = MenuItemKind::Separator;
            items.push_back(std::move(item));
            continue;
        }

        std::wstring text;
        std::wstring shortcut;
        if (TryGetMenuItemPresentationText(menu, position, text, shortcut))
        {
            item.text            = StripMenuMnemonicMarkers(text);
            item.acceleratorText = shortcut;
        }

        item.commandId = static_cast<int>(itemInfo.wID);
        item.enabled   = (itemInfo.fState & MFS_GRAYED) == 0;
        item.checked   = (itemInfo.fState & MFS_CHECKED) != 0;

        if ((itemInfo.fType & MFT_RADIOCHECK) != 0)
        {
            item.kind = MenuItemKind::Radio;
        }
        else if (item.checked)
        {
            item.kind = MenuItemKind::Toggle;
        }

        if (const std::optional<std::wstring> shortcutText = TryGetDynamicMenuShortcutText(itemInfo.wID); shortcutText.has_value())
        {
            item.acceleratorText = shortcutText.value();
        }

        if (const wchar_t iconGlyph = GetMainMenuCommandIconGlyph(itemInfo.wID); iconGlyph != 0)
        {
            item.iconGlyph.assign(1u, iconGlyph);
        }

        if (itemInfo.hSubMenu)
        {
            item.children = ConvertHMenuToDxFlyoutItems(itemInfo.hSubMenu);
        }

        items.push_back(std::move(item));
    }

    return items;
}

[[nodiscard]] std::vector<RedSalamander::DxUi::MenuBarItem> BuildDxMenuBarItems(HMENU menu) noexcept
{
    std::vector<RedSalamander::DxUi::MenuBarItem> items;
    if (! menu)
    {
        return items;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return items;
    }

    items.reserve(static_cast<size_t>(itemCount));
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, position, TRUE, &itemInfo))
        {
            continue;
        }

        if ((itemInfo.fType & MFT_SEPARATOR) != 0 || itemInfo.hSubMenu == nullptr)
        {
            continue;
        }

        std::wstring text;
        std::wstring shortcut;
        if (! TryGetMenuItemPresentationText(menu, position, text, shortcut) || text.empty())
        {
            continue;
        }

        RedSalamander::DxUi::MenuBarItem item{};
        item.text           = StripMenuMnemonicMarkers(text);
        item.mnemonic       = FindMenuMnemonic(text);
        item.enabled        = (itemInfo.fState & MFS_GRAYED) == 0;
        item.rightJustified = (itemInfo.fType & MFT_RIGHTJUSTIFY) != 0;
        item.sourceIndex    = static_cast<size_t>(position);
        items.push_back(std::move(item));
    }

    return items;
}

class MainMenuBarHost final
{
public:
    MainMenuBarHost()                                  = default;
    MainMenuBarHost(const MainMenuBarHost&)            = delete;
    MainMenuBarHost& operator=(const MainMenuBarHost&) = delete;
    MainMenuBarHost(MainMenuBarHost&&)                 = delete;
    MainMenuBarHost& operator=(MainMenuBarHost&&)      = delete;

    [[nodiscard]] bool EnsureCreated(HWND ownerWindow) noexcept
    {
        if (! ownerWindow || IsWindow(ownerWindow) == FALSE)
        {
            return false;
        }

        _ownerWindow = ownerWindow;
        if (_hwnd && IsWindow(_hwnd.get()) != FALSE)
        {
            return true;
        }

        HINSTANCE instance = GetModuleHandleW(nullptr);
        if (! EnsureWindowClass(instance))
        {
            return false;
        }

        HWND hwnd = CreateWindowExW(0, kWindowClassName, L"", WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, 0, 0, 1, 1, ownerWindow, nullptr, instance, this);
        if (! hwnd)
        {
            return false;
        }

        _hwnd.reset(hwnd);
        if (! _host.Attach(hwnd))
        {
            _hwnd.reset();
            return false;
        }

        _host.SetOnEscape([this]() noexcept
        {
            if (g_menuBarTemporarilyShown && ! g_menuBarVisible)
            {
                DismissTemporaryBar();
            }
            static_cast<void>(g_folderWindow.TryRestoreActivePaneFolderViewFocus());
            return true;
        });

        auto menuBar = std::make_unique<RedSalamander::DxUi::MenuBar>();
        _menuBar     = menuBar.get();
        _host.SetRoot(std::move(menuBar));
        _host.SetTheme(MakeAppThemeDxPalette(_theme));
        SyncMenuModel();
        return true;
    }

    void Destroy() noexcept
    {
        _host.Detach();
        _menuBar = nullptr;
        _hwnd.reset();
        _ownerWindow      = nullptr;
        _focusRestoreHwnd = nullptr;
    }

    void SetTheme(const AppTheme& theme) noexcept
    {
        _theme = theme;
        if (_hwnd && IsWindow(_hwnd.get()) != FALSE)
        {
            _host.SetTheme(MakeAppThemeDxPalette(_theme));
            UpdateLayout();
        }
    }

    void SyncMenuModel() noexcept
    {
        if (! _menuBar)
        {
            return;
        }

        EnsureMenuHandles(_ownerWindow);
        _menuBar->SetItems(BuildDxMenuBarItems(g_mainMenuHandle));
        _menuBar->SetOnOpenItem([this](size_t index, POINT screenPoint, bool keyboardInvocation) { OpenPopup(index, screenPoint, keyboardInvocation); });
        UpdateSelectedIndexSnapshot();
    }

    void UpdateLayout() noexcept
    {
        if (! _ownerWindow || ! _hwnd || IsWindow(_hwnd.get()) == FALSE)
        {
            return;
        }

        RECT client{};
        if (! GetClientRect(_ownerWindow, &client))
        {
            return;
        }

        const bool visible = ShouldShow();
        const int width    = client.right - client.left;
        const int height   = visible ? GetVisibleHeightPx() : 0;
        MoveWindow(_hwnd.get(), client.left, client.top, width, height, TRUE);
        ShowWindow(_hwnd.get(), visible ? SW_SHOWNA : SW_HIDE);
        SyncMenuModel();
    }

    [[nodiscard]] int GetVisibleHeightPx() const noexcept
    {
        const UINT dpi      = _ownerWindow && IsWindow(_ownerWindow) != FALSE ? GetDpiForWindow(_ownerWindow) : USER_DEFAULT_SCREEN_DPI;
        const int heightDip = _theme.compactMode ? 26 : 32;
        return MulDiv(heightDip, static_cast<int>(dpi == 0u ? USER_DEFAULT_SCREEN_DPI : dpi), USER_DEFAULT_SCREEN_DPI);
    }

    [[nodiscard]] bool FocusFirstItem() noexcept
    {
        if (! _menuBar || ! _hwnd || IsWindow(_hwnd.get()) == FALSE)
        {
            return false;
        }

        const auto items = _menuBar->GetItems();
        if (items.empty())
        {
            return false;
        }

        CaptureFocusRestoreTarget();
        _menuBar->SetSelectedIndex(0u);
        UpdateSelectedIndexSnapshot();
        SetFocus(_hwnd.get());
        _host.SetFocusControl(_menuBar);
        return true;
    }

    [[nodiscard]] bool ActivateMnemonic(wchar_t mnemonic) noexcept
    {
        if (! _menuBar || ! _hwnd || IsWindow(_hwnd.get()) == FALSE)
        {
            return false;
        }

        CaptureFocusRestoreTarget();
        SetFocus(_hwnd.get());
        const bool activated = _menuBar->ActivateMnemonic(_host, mnemonic);
        UpdateSelectedIndexSnapshot();
        return activated;
    }

    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hwnd.get();
    }

    [[nodiscard]] std::optional<size_t> GetSelectedIndex() const noexcept
    {
        const int selectedIndex = _selectedIndexSnapshot.load(std::memory_order_acquire);
        return selectedIndex >= 0 ? std::optional<size_t>{static_cast<size_t>(selectedIndex)} : std::nullopt;
    }

    [[nodiscard]] size_t GetVisualHighlightCount() const noexcept
    {
        const int highlightCount = _visualHighlightCountSnapshot.load(std::memory_order_acquire);
        return highlightCount > 0 ? static_cast<size_t>(highlightCount) : 0u;
    }

    [[nodiscard]] uint64_t GetRenderCount() const noexcept
    {
#if defined(ENABLE_TESTS)
        return _host.DebugGetRenderCount();
#else
        return 0u;
#endif
    }

    [[nodiscard]] bool GetItemLabel(size_t index, std::wstring& outText) const noexcept
    {
        outText.clear();
        if (! _menuBar)
        {
            return false;
        }

        const std::span<const RedSalamander::DxUi::MenuBarItem> items = _menuBar->GetItems();
        if (index >= items.size())
        {
            return false;
        }

        outText.assign(items[index].text);
        return true;
    }

    [[nodiscard]] bool GetItemSourceIndex(size_t index, size_t& outSourceIndex) const noexcept
    {
        outSourceIndex = 0u;
        if (! _menuBar)
        {
            return false;
        }

        const std::span<const RedSalamander::DxUi::MenuBarItem> items = _menuBar->GetItems();
        if (index >= items.size())
        {
            return false;
        }

        outSourceIndex = items[index].sourceIndex;
        return true;
    }

    [[nodiscard]] bool GetItemScreenRect(size_t index, RECT& rectPx) const noexcept
    {
        return _menuBar && _hwnd && IsWindow(_hwnd.get()) != FALSE && _menuBar->TryGetItemScreenRect(_host, index, rectPx);
    }

    [[nodiscard]] bool HitTestItemScreenPoint(POINT screenPoint, size_t& outIndex) const noexcept
    {
        outIndex                             = 0u;
        const std::optional<size_t> hitIndex = HitTestScreenPoint(screenPoint);
        if (! hitIndex.has_value())
        {
            return false;
        }

        outIndex = hitIndex.value();
        return true;
    }

private:
    static constexpr wchar_t kWindowClassName[] = L"RedSalamander.DxMainMenuBar";

    [[nodiscard]] bool EnsureWindowClass(HINSTANCE instance) noexcept
    {
        static ATOM atom = 0;
        if (atom != 0)
        {
            return true;
        }

        WNDCLASSEXW windowClass{};
        windowClass.cbSize        = sizeof(windowClass);
        windowClass.style         = CS_HREDRAW | CS_VREDRAW;
        windowClass.lpfnWndProc   = &MainMenuBarHost::WndProc;
        windowClass.hInstance     = instance;
        windowClass.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        windowClass.hbrBackground = nullptr;
        windowClass.lpszClassName = kWindowClassName;
        atom                      = RegisterClassExW(&windowClass);
        return atom != 0;
    }

    [[nodiscard]] bool ShouldShow() const noexcept
    {
        return g_menuBarVisible || g_menuBarTemporarilyShown;
    }

    void CaptureFocusRestoreTarget() noexcept
    {
        RedSalamander::DxUi::CaptureFocusRestoreTarget(_ownerWindow, _hwnd.get(), _focusRestoreHwnd);
    }

    [[nodiscard]] bool RestoreCapturedFocus() noexcept
    {
        return RedSalamander::DxUi::RestoreCapturedFocus(_focusRestoreHwnd);
    }

    void ClearMenuSessionFocusState() noexcept
    {
        if (_menuBar)
        {
            _menuBar->SetSelectedIndex(std::nullopt);
        }
        _host.SetFocusControl(nullptr);
        UpdateSelectedIndexSnapshot();
    }

    [[nodiscard]] bool RestoreFocusAfterMenuSession() noexcept
    {
        ClearMenuSessionFocusState();
        if (RestoreCapturedFocus())
        {
            return true;
        }
        if (g_folderWindow.GetFocusedFolderViewHwnd() != nullptr)
        {
            return true;
        }
        if (g_folderWindow.TryRestoreActivePaneFolderViewFocus())
        {
            return true;
        }
        if (_ownerWindow && IsWindow(_ownerWindow) != FALSE)
        {
            SetFocus(_ownerWindow);
            return GetFocus() == _ownerWindow;
        }
        return false;
    }

    [[nodiscard]] std::optional<size_t> HitTestScreenPoint(POINT screenPoint) const noexcept
    {
        if (! _menuBar || ! _hwnd || IsWindow(_hwnd.get()) == FALSE)
        {
            return std::nullopt;
        }

        const std::optional<RedSalamander::DxUi::PointDip> pointDip = _host.ScreenPointToDipPoint(screenPoint);
        if (! pointDip.has_value())
        {
            return std::nullopt;
        }

        return _menuBar->HitTestPoint(_host, pointDip.value());
    }

    [[nodiscard]] std::optional<size_t> ResolveRawMenuIndex(size_t visualIndex) const noexcept
    {
        if (! _menuBar)
        {
            return std::nullopt;
        }

        const auto items = _menuBar->GetItems();
        if (visualIndex >= items.size())
        {
            return std::nullopt;
        }

        const size_t sourceIndex = items[visualIndex].sourceIndex;
        return sourceIndex == static_cast<size_t>(-1) ? std::optional<size_t>{visualIndex} : std::optional<size_t>{sourceIndex};
    }

    [[nodiscard]] std::optional<POINT> GetItemAnchorScreenPoint(size_t index) const noexcept
    {
        RECT itemRectPx{};
        if (! GetItemScreenRect(index, itemRectPx))
        {
            return std::nullopt;
        }

        return POINT{itemRectPx.left, itemRectPx.bottom};
    }

    [[nodiscard]] std::optional<size_t> FindNextEnabledItem(size_t currentIndex, bool forward) const noexcept
    {
        if (! _menuBar)
        {
            return std::nullopt;
        }

        const auto items = _menuBar->GetItems();
        if (items.empty() || currentIndex >= items.size())
        {
            return std::nullopt;
        }

        for (size_t step = 1u; step <= items.size(); ++step)
        {
            const size_t nextIndex = forward ? ((currentIndex + step) % items.size()) : ((currentIndex + items.size() - (step % items.size())) % items.size());
            if (items[nextIndex].enabled)
            {
                return nextIndex;
            }
        }

        return std::nullopt;
    }

    [[nodiscard]] std::optional<RedSalamander::DxUi::ContextMenuRootSwitchRequest> BuildRootSwitchRequest(size_t index) noexcept
    {
        if (! _ownerWindow || ! g_mainMenuHandle || ! _menuBar)
        {
            return std::nullopt;
        }

        const std::optional<size_t> rawIndex = ResolveRawMenuIndex(index);
        if (! rawIndex.has_value())
        {
            return std::nullopt;
        }

        const HMENU popupMenu = GetSubMenu(g_mainMenuHandle, static_cast<int>(rawIndex.value()));
        if (! popupMenu)
        {
            return std::nullopt;
        }

        OnInitMenuPopup(_ownerWindow, popupMenu);
        SyncMenuModel();
        _menuBar->SetSelectedIndex(index);
        UpdateSelectedIndexSnapshot();

        const std::optional<POINT> screenPoint = GetItemAnchorScreenPoint(index);
        if (! screenPoint.has_value())
        {
            return std::nullopt;
        }

        RedSalamander::DxUi::ContextMenuRootSwitchRequest request{};
        request.screenPoint = screenPoint.value();
        request.items       = ConvertHMenuToDxFlyoutItems(popupMenu);
        if (request.items.empty())
        {
            return std::nullopt;
        }
        return request;
    }

    void OpenPopup(size_t index, POINT screenPoint, bool keyboardInvocation) noexcept
    {
        if (! _ownerWindow || ! g_mainMenuHandle || ! _menuBar)
        {
            return;
        }

        CaptureFocusRestoreTarget();
        const std::optional<size_t> rawIndex = ResolveRawMenuIndex(index);
        if (! rawIndex.has_value())
        {
            return;
        }

        const HMENU popupMenu = GetSubMenu(g_mainMenuHandle, static_cast<int>(rawIndex.value()));
        if (! popupMenu)
        {
            return;
        }

        OnInitMenuPopup(_ownerWindow, popupMenu);
        SyncMenuModel();
        _menuBar->SetSelectedIndex(index);
        UpdateSelectedIndexSnapshot();

        const auto flyoutItems = ConvertHMenuToDxFlyoutItems(popupMenu);
        RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
        sessionCallbacks.focusFirstNavigableItem = keyboardInvocation;

        size_t activeIndex                     = index;
        sessionCallbacks.switchRootFromPointer = [this,
                                                  &activeIndex](POINT hoverScreenPoint) -> std::optional<RedSalamander::DxUi::ContextMenuRootSwitchRequest>
        {
            const std::optional<size_t> hitIndex = HitTestScreenPoint(hoverScreenPoint);
            if (! hitIndex.has_value() || hitIndex.value() == activeIndex || ! _menuBar)
            {
                return std::nullopt;
            }

            const auto items = _menuBar->GetItems();
            if (hitIndex.value() >= items.size() || ! items[hitIndex.value()].enabled)
            {
                return std::nullopt;
            }

            if (auto request = BuildRootSwitchRequest(hitIndex.value()); request.has_value())
            {
                activeIndex = hitIndex.value();
                return request;
            }

            return std::nullopt;
        };
        sessionCallbacks.switchRootFromDirection = [this, &activeIndex](bool forward) -> std::optional<RedSalamander::DxUi::ContextMenuRootSwitchRequest>
        {
            const std::optional<size_t> nextIndex = FindNextEnabledItem(activeIndex, forward);
            if (! nextIndex.has_value() || nextIndex.value() == activeIndex)
            {
                return std::nullopt;
            }

            if (auto request = BuildRootSwitchRequest(nextIndex.value()); request.has_value())
            {
                activeIndex = nextIndex.value();
                return request;
            }

            return std::nullopt;
        };

        const auto result = RedSalamander::DxUi::ContextMenu::Show(_ownerWindow, screenPoint, flyoutItems, MakeAppThemeDxPalette(_theme), sessionCallbacks);

        SyncMenuModel();
        if (g_menuBarTemporarilyShown && ! g_menuBarVisible)
        {
            DismissTemporaryBar();
        }
        else
        {
            static_cast<void>(RestoreFocusAfterMenuSession());
        }

        if (result.has_value())
        {
            PostMessageW(_ownerWindow, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(result.value()), 0), 0);
        }
    }

    void DismissTemporaryBar() noexcept
    {
        if (! g_menuBarTemporarilyShown || g_menuBarVisible || ! _ownerWindow)
        {
            return;
        }

        g_menuBarTemporarilyShown = false;
        const bool restored       = RestoreFocusAfterMenuSession();
        SendMessageW(_ownerWindow, WM_SIZE, 0, 0);
        if (! restored)
        {
            if (const HWND folderWindow = g_hFolderWindow.load(std::memory_order_acquire); folderWindow && IsWindow(folderWindow) != FALSE)
            {
                SetFocus(folderWindow);
            }
            else if (IsWindow(_ownerWindow) != FALSE)
            {
                SetFocus(_ownerWindow);
            }
        }
        UpdateSelectedIndexSnapshot();
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        auto* self = reinterpret_cast<MainMenuBarHost*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (message == WM_NCCREATE)
        {
            const auto* createStruct = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            self                     = createStruct ? static_cast<MainMenuBarHost*>(createStruct->lpCreateParams) : nullptr;
            if (self)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            }
        }

        if (self)
        {
            bool handled         = false;
            const LRESULT result = self->_host.HandleMessage(hwnd, message, wParam, lParam, handled);
            if (handled)
            {
                if (message == WM_NCDESTROY)
                {
                    self->_menuBar = nullptr;
                    self->_selectedIndexSnapshot.store(-1, std::memory_order_release);
                    self->_visualHighlightCountSnapshot.store(0, std::memory_order_release);
                    self->_focusRestoreHwnd = nullptr;
                }
                else
                {
                    self->UpdateSelectedIndexSnapshot();
                }
                if (message == WM_KILLFOCUS && g_menuBarTemporarilyShown && ! g_menuBarVisible)
                {
                    self->DismissTemporaryBar();
                    return 0;
                }
                return result;
            }

            if (message == WM_KILLFOCUS && g_menuBarTemporarilyShown && ! g_menuBarVisible)
            {
                self->DismissTemporaryBar();
                return 0;
            }
            if (message == WM_NCDESTROY)
            {
                self->_menuBar = nullptr;
                self->_selectedIndexSnapshot.store(-1, std::memory_order_release);
                self->_visualHighlightCountSnapshot.store(0, std::memory_order_release);
                self->_focusRestoreHwnd = nullptr;
            }
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }

    void UpdateSelectedIndexSnapshot() noexcept
    {
        if (! _menuBar)
        {
            _selectedIndexSnapshot.store(-1, std::memory_order_release);
            _visualHighlightCountSnapshot.store(0, std::memory_order_release);
            return;
        }

        const std::optional<size_t> selectedIndex = _menuBar->GetSelectedIndex();
        _selectedIndexSnapshot.store(selectedIndex.has_value() ? static_cast<int>(selectedIndex.value()) : -1, std::memory_order_release);
        _visualHighlightCountSnapshot.store(static_cast<int>(_menuBar->GetVisualHighlightCount()), std::memory_order_release);
    }

    HWND _ownerWindow = nullptr;
    wil::unique_hwnd _hwnd;
    AppTheme _theme{};
    RedSalamander::DxUi::WindowHost _host;
    RedSalamander::DxUi::MenuBar* _menuBar = nullptr;
    std::atomic<int> _selectedIndexSnapshot{-1};
    std::atomic<int> _visualHighlightCountSnapshot{0};
    HWND _focusRestoreHwnd = nullptr;
};

MainMenuBarHost g_mainMenuBarHost;

[[nodiscard]] bool ReloadLocalizedMainMenu(HWND hWnd, bool syncMenuBar) noexcept
{
    wil::unique_hmenu newMenu(Localization::LoadMenuResource(GetModuleHandleW(nullptr), IDC_REDSALAMANDER));
    if (! newMenu)
    {
        Debug::Warning(L"Failed to load localized main menu resource.");
        return false;
    }

    wil::unique_hmenu oldMenu;
    if (g_mainMenuHandle && g_mainMenuHandle != newMenu.get())
    {
        if (GetMenu(hWnd) == g_mainMenuHandle)
        {
            SetMenu(hWnd, nullptr);
            DrawMenuBar(hWnd);
        }
        oldMenu.reset(g_mainMenuHandle);
    }

    ResetMenuHandleCache();
    g_mainMenuHandle = newMenu.release();
    EnsureMenuHandles(hWnd);
    RebuildThemeMenuDynamicItems(hWnd);
    RebuildPluginsMenuDynamicItems(hWnd);

    if (syncMenuBar)
    {
        g_mainMenuBarHost.SyncMenuModel();
    }

    return true;
}

} // namespace

#ifdef ENABLE_TESTS
[[nodiscard]] wchar_t DebugGetMainMenuIconGlyph(UINT menuCommandId) noexcept
{
    return GetMainMenuCommandIconGlyph(menuCommandId);
}

[[nodiscard]] HMENU DebugGetMainMenuModelHandle() noexcept
{
    return g_mainMenuHandle;
}

[[nodiscard]] bool DebugIsMainMenuBarSurfaceVisible(HWND mainWindow) noexcept
{
    const HWND menuBar = FindWindowExW(mainWindow, nullptr, L"RedSalamander.DxMainMenuBar", nullptr);
    return menuBar && IsWindowVisible(menuBar) != FALSE;
}

[[nodiscard]] int DebugGetMainMenuBarSelectedIndex() noexcept
{
    const std::optional<size_t> selectedIndex = g_mainMenuBarHost.GetSelectedIndex();
    return selectedIndex.has_value() ? static_cast<int>(selectedIndex.value()) : -1;
}

[[nodiscard]] int DebugGetMainMenuBarVisualHighlightCount() noexcept
{
    return static_cast<int>(g_mainMenuBarHost.GetVisualHighlightCount());
}

[[nodiscard]] uint64_t DebugGetMainMenuBarRenderCount() noexcept
{
    return g_mainMenuBarHost.GetRenderCount();
}

[[nodiscard]] bool DebugGetMainMenuBarItemLabel(size_t index, std::wstring& outText) noexcept
{
    return g_mainMenuBarHost.GetItemLabel(index, outText);
}

[[nodiscard]] bool DebugGetMainMenuBarItemSourceIndex(size_t index, size_t& outSourceIndex) noexcept
{
    return g_mainMenuBarHost.GetItemSourceIndex(index, outSourceIndex);
}

[[nodiscard]] bool DebugGetMainMenuBarItemScreenRect(HWND mainWindow, size_t index, RECT& rectPx) noexcept
{
    const HWND menuBar = FindWindowExW(mainWindow, nullptr, L"RedSalamander.DxMainMenuBar", nullptr);
    if (! menuBar || IsWindowVisible(menuBar) == FALSE)
    {
        return false;
    }

    return g_mainMenuBarHost.GetItemScreenRect(index, rectPx);
}

[[nodiscard]] bool DebugHitTestMainMenuBarScreenPoint(HWND mainWindow, POINT screenPoint, size_t& outIndex) noexcept
{
    outIndex           = 0u;
    const HWND menuBar = FindWindowExW(mainWindow, nullptr, L"RedSalamander.DxMainMenuBar", nullptr);
    if (! menuBar || IsWindowVisible(menuBar) == FALSE)
    {
        return false;
    }

    return g_mainMenuBarHost.HitTestItemScreenPoint(screenPoint, outIndex);
}
#endif

// Forward declarations of functions included in this code module:
std::optional<HWND> InitInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
static void AdjustLayout(HWND hWnd);

namespace
{
ShortcutManager g_shortcutManager;

[[nodiscard]] std::filesystem::path GetDefaultFileSystemRoot() noexcept
{
    wchar_t buffer[MAX_PATH]{};
    const UINT bufferSize = static_cast<UINT>(std::size(buffer));
    const UINT length     = GetWindowsDirectoryW(buffer, bufferSize);
    if (length > 0 && length < bufferSize)
    {
        const std::filesystem::path root = std::filesystem::path(buffer).root_path();
        if (! root.empty())
        {
            return root;
        }
    }

    return std::filesystem::path(L"C:\\");
}

[[nodiscard]] bool LooksLikeUncPath(std::wstring_view text) noexcept
{
    return text.rfind(L"\\\\", 0) == 0 || text.rfind(L"//", 0) == 0;
}

[[nodiscard]] std::optional<std::wstring> TryGetUncShareRoot(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return std::nullopt;
    }

    std::wstring normalized(text);
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');

    if (normalized.rfind(L"\\\\", 0) != 0)
    {
        return std::nullopt;
    }

    const size_t serverStart = 2;
    const size_t serverEnd   = normalized.find(L'\\', serverStart);
    if (serverEnd == std::wstring::npos || serverEnd == serverStart)
    {
        return std::nullopt;
    }

    const size_t shareStart = serverEnd + 1;
    if (shareStart >= normalized.size())
    {
        return std::nullopt;
    }

    const size_t shareEnd = normalized.find(L'\\', shareStart);
    if (shareEnd == std::wstring::npos)
    {
        return normalized;
    }

    if (shareEnd <= shareStart)
    {
        return std::nullopt;
    }

    return normalized.substr(0, shareEnd);
}

[[nodiscard]] bool IsEditControl(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    wchar_t className[64]{};
    const int length = GetClassNameW(hwnd, className, static_cast<int>(std::size(className)));
    if (length <= 0)
    {
        return false;
    }

    const std::wstring_view name(className, static_cast<size_t>(length));
    if (name == L"Edit")
    {
        return true;
    }

    if (name.rfind(L"RichEdit", 0) == 0 || name.rfind(L"RICHEDIT", 0) == 0)
    {
        return true;
    }

    return false;
}

[[nodiscard]] HRESULT ShowConnectNetworkDriveDialog(HWND ownerWindow, const std::optional<std::wstring>& remoteNameOpt) noexcept
{
    DWORD error = ERROR_SUCCESS;
    if (remoteNameOpt.has_value() && ! remoteNameOpt.value().empty())
    {
        std::wstring remoteName;
        remoteName = remoteNameOpt.value();
        std::replace(remoteName.begin(), remoteName.end(), L'/', L'\\');

        NETRESOURCEW netResource{};
        netResource.dwType = RESOURCETYPE_DISK;

        CONNECTDLGSTRUCTW dialog{};
        dialog.cbStructure       = sizeof(dialog);
        dialog.hwndOwner         = ownerWindow;
        dialog.lpConnRes         = &netResource;
        dialog.dwFlags           = CONNDLG_RO_PATH;
        netResource.lpRemoteName = remoteName.data();
        dialog.dwFlags           = CONNDLG_USE_MRU;
        error                    = WNetConnectionDialog1W(&dialog);
    }
    else
    {
        error = WNetConnectionDialog(ownerWindow, RESOURCETYPE_DISK);
    }

    if (error == NO_ERROR)
    {
        return S_OK;
    }
    if (error == ERROR_CANCELLED)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    Debug::Error(L"ShowConnectNetworkDriveDialog failed (err=0x{:08X}).", static_cast<unsigned long>(error));
    return HRESULT_FROM_WIN32(error);
}

[[nodiscard]] HRESULT ShowDisconnectNetworkDriveDialog(HWND ownerWindow,
                                                       const std::optional<std::wstring>& localNameOpt,
                                                       const std::optional<std::wstring>& remoteNameOpt) noexcept
{
    DISCDLGSTRUCTW dialog{};
    dialog.cbStructure = sizeof(dialog);
    dialog.hwndOwner   = ownerWindow;

    std::wstring localName;
    if (localNameOpt.has_value() && ! localNameOpt.value().empty())
    {
        localName = localNameOpt.value();
        std::replace(localName.begin(), localName.end(), L'/', L'\\');
        dialog.lpLocalName = localName.data();
    }

    std::wstring remoteName;
    if (remoteNameOpt.has_value() && ! remoteNameOpt.value().empty())
    {
        remoteName = remoteNameOpt.value();
        std::replace(remoteName.begin(), remoteName.end(), L'/', L'\\');
        dialog.lpRemoteName = remoteName.data();
    }

    DWORD error = 0;
    if (dialog.lpLocalName || dialog.lpRemoteName)
    {
        error = WNetDisconnectDialog1W(&dialog);
    }
    else
    {
        error = WNetDisconnectDialog(ownerWindow, RESOURCETYPE_DISK);
    }

    if (error == NO_ERROR)
    {
        return S_OK;
    }
    if (error == ERROR_CANCELLED)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    Debug::Error(L"WNetDisconnectDialog1W failed (err=0x{:08X}).", static_cast<unsigned long>(error));
    return HRESULT_FROM_WIN32(error);
}

[[nodiscard]] bool IsEditControlFocused() noexcept
{
    return IsEditControl(GetFocus());
}

[[nodiscard]] uint32_t GetCurrentShortcutModifiers() noexcept
{
    uint32_t modifiers = 0;

    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModCtrl;
    }

    if ((GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModAlt;
    }

    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModShift;
    }

    return modifiers;
}

void ReloadShortcutsFromSettings()
{
    if (! g_settings.shortcuts.has_value())
    {
        g_shortcutManager.Clear();
        if (g_hFolderWindow.load(std::memory_order_acquire))
        {
            g_folderWindow.SetShortcutManager(nullptr);
        }
        return;
    }

    g_shortcutManager.Load(g_settings.shortcuts.value());
    if (g_hFolderWindow.load(std::memory_order_acquire))
    {
        g_folderWindow.SetShortcutManager(&g_shortcutManager);
    }
}

[[nodiscard]] bool SendKeyToFocusedFolderView(uint32_t vk) noexcept
{
    if (! g_hFolderWindow.load(std::memory_order_acquire))
    {
        return false;
    }

    const HWND folderView = g_folderWindow.GetFocusedFolderViewHwnd();
    if (! folderView)
    {
        return false;
    }

    SendMessageW(folderView, WM_KEYDOWN, static_cast<WPARAM>(vk), 0);
    return true;
}

[[nodiscard]] bool SendKeyToFolderView(FolderWindow::Pane pane, uint32_t vk) noexcept
{
    if (! g_hFolderWindow.load(std::memory_order_acquire))
    {
        return false;
    }

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(pane);
    if (! folderView)
    {
        return false;
    }

    SendMessageW(folderView, WM_KEYDOWN, static_cast<WPARAM>(vk), 0);
    return true;
}

[[nodiscard]] bool SendCommandToFolderView(FolderWindow::Pane pane, uint32_t commandId) noexcept
{
    if (! g_hFolderWindow.load(std::memory_order_acquire))
    {
        return false;
    }

    const HWND folderView = g_folderWindow.GetFolderViewHwnd(pane);
    if (! folderView)
    {
        return false;
    }

    SendMessageW(folderView, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(commandId), 0), 0);
    return true;
}

void ShowCommandNotImplementedMessage(HWND ownerWindow, std::wstring_view commandId) noexcept
{
    if (commandId.empty())
    {
        return;
    }

    commandId = CanonicalizeCommandId(commandId);

    std::wstring displayName;
    const std::optional<unsigned int> displayNameIdOpt = TryGetCommandDisplayNameStringId(commandId);
    if (displayNameIdOpt.has_value())
    {
        displayName = LoadStringResource(nullptr, displayNameIdOpt.value());
    }
    if (displayName.empty())
    {
        displayName = std::wstring(commandId);
    }

    const std::wstring text    = FormatStringResource(nullptr, IDS_FMT_CMD_NOT_IMPLEMENTED, displayName);
    const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_NOT_IMPLEMENTED);

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = (ownerWindow && IsWindow(ownerWindow)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = HOST_ALERT_INFO;
    request.targetWindow = (request.scope == HOST_ALERT_SCOPE_WINDOW) ? ownerWindow : nullptr;
    request.title        = caption.c_str();
    request.message      = text.c_str();
    request.closable     = TRUE;
    static_cast<void>(HostShowAlert(request));
}

[[nodiscard]] bool ExecuteCommandById(HWND ownerWindow, std::wstring_view commandId) noexcept
{
    if (commandId.empty())
    {
        return false;
    }

    const std::wstring_view originalCommandId = commandId;
    std::optional<std::wstring_view> themeId;
    {
        constexpr std::wstring_view kThemeSelectPrefix = L"cmd/app/theme/select/";
        if (originalCommandId.starts_with(kThemeSelectPrefix) && originalCommandId.size() > kThemeSelectPrefix.size())
        {
            themeId = originalCommandId.substr(kThemeSelectPrefix.size());
        }
    }

    std::optional<std::wstring_view> viewerActionId;
    {
        constexpr std::wstring_view kViewWithPrefix = L"cmd/pane/viewWith/";
        if (originalCommandId.starts_with(kViewWithPrefix) && originalCommandId.size() > kViewWithPrefix.size())
        {
            viewerActionId = originalCommandId.substr(kViewWithPrefix.size());
        }
    }

    std::optional<std::wstring_view> editorActionId;
    {
        constexpr std::wstring_view kEditWithPrefix = L"cmd/pane/editWith/";
        if (originalCommandId.starts_with(kEditWithPrefix) && originalCommandId.size() > kEditWithPrefix.size())
        {
            editorActionId = originalCommandId.substr(kEditWithPrefix.size());
        }
    }

    std::optional<std::wstring_view> userMenuActionId;
    {
        constexpr std::wstring_view kUserMenuPrefix = L"cmd/pane/userMenu/";
        if (originalCommandId.starts_with(kUserMenuPrefix) && originalCommandId.size() > kUserMenuPrefix.size())
        {
            userMenuActionId = originalCommandId.substr(kUserMenuPrefix.size());
        }
    }

    std::optional<std::wstring_view> shellNewTemplateId;
    {
        constexpr std::wstring_view kShellNewTemplatePrefix = L"cmd/pane/newFromShellTemplate/";
        if (originalCommandId.starts_with(kShellNewTemplatePrefix) && originalCommandId.size() > kShellNewTemplatePrefix.size())
        {
            shellNewTemplateId = originalCommandId.substr(kShellNewTemplatePrefix.size());
        }
    }

    std::optional<wchar_t> driveRootLetter;
    {
        constexpr std::wstring_view kGoDriveRootPrefix = L"cmd/pane/goDriveRoot/";
        if (originalCommandId.starts_with(kGoDriveRootPrefix) && originalCommandId.size() > kGoDriveRootPrefix.size())
        {
            const wchar_t rawLetter = originalCommandId[kGoDriveRootPrefix.size()];
            if (std::iswalpha(static_cast<wint_t>(rawLetter)) != 0)
            {
                const wchar_t upper = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(rawLetter)));
                if (upper >= L'A' && upper <= L'Z')
                {
                    driveRootLetter = upper;
                }
            }
        }
    }

    std::optional<int> hotPathSlotIndex;
    {
        constexpr std::wstring_view kHotPathPrefix    = L"cmd/pane/hotPath/";
        constexpr std::wstring_view kSetHotPathPrefix = L"cmd/pane/setHotPath/";
        std::wstring_view suffix;
        if (originalCommandId.starts_with(kHotPathPrefix) && originalCommandId.size() > kHotPathPrefix.size())
        {
            suffix = originalCommandId.substr(kHotPathPrefix.size());
        }
        else if (originalCommandId.starts_with(kSetHotPathPrefix) && originalCommandId.size() > kSetHotPathPrefix.size())
        {
            suffix = originalCommandId.substr(kSetHotPathPrefix.size());
        }

        if (! suffix.empty())
        {
            const wchar_t digit = suffix[0];
            if (digit >= L'1' && digit <= L'9')
            {
                hotPathSlotIndex = static_cast<int>(digit - L'1');
            }
            else if (digit == L'0')
            {
                hotPathSlotIndex = 9;
            }
        }
    }

    commandId = CanonicalizeCommandId(commandId);

    if (commandId == L"cmd/app/theme/select" && themeId.has_value())
    {
        ApplyThemeId(ownerWindow, themeId.value());
        return true;
    }

    if (commandId == L"cmd/pane/menu")
    {
        SendMessageW(ownerWindow, WM_SYSCOMMAND, SC_KEYMENU, 0);
        return true;
    }

    if (commandId == L"cmd/pane/focusAddressBar")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandFocusAddressBar(pane);
        return true;
    }

    if (commandId == L"cmd/pane/historyBack")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandHistoryBack(pane);
        return true;
    }

    if (commandId == L"cmd/pane/historyForward")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandHistoryForward(pane);
        return true;
    }

    if (commandId == L"cmd/pane/setPathFromOtherPane")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandSetPathFromOtherPane(pane);
        return true;
    }

    if (commandId == L"cmd/pane/goRootDirectory")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandGoRootDirectory(pane);
        return true;
    }

    if (commandId == L"cmd/pane/upOneDirectory")
    {
        return SendKeyToFocusedFolderView(VK_BACK);
    }
    if (commandId == L"cmd/pane/switchPaneFocus")
    {
        return SendKeyToFocusedFolderView(VK_TAB);
    }
    if (commandId == L"cmd/pane/selection/unselectAll")
    {
        return SendKeyToFocusedFolderView(VK_ESCAPE);
    }
    if (commandId == L"cmd/pane/zoomPanel")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.ToggleZoomPanel(pane);
        return true;
    }
    if (commandId == L"cmd/pane/viewOptions/toggleStatusBar")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
        const bool visible           = g_folderWindow.GetStatusBarVisible(pane);
        g_folderWindow.SetStatusBarVisible(pane, ! visible);
        UpdatePaneMenuChecks();
        return true;
    }
    if (commandId == L"cmd/pane/viewOptions/toggleFileExtensions")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
        const bool visible           = g_folderWindow.GetFileExtensionsVisible(pane);
        g_folderWindow.SetFileExtensionsVisible(pane, ! visible);
        UpdatePaneMenuChecks();
        return true;
    }
    if (commandId == L"cmd/pane/viewOptions/toggleNavigationBar")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
        const bool visible           = g_folderWindow.GetNavigationBarVisible(pane);
        g_folderWindow.SetNavigationBarVisible(pane, ! visible);
        UpdatePaneMenuChecks();
        return true;
    }
    if (commandId == L"cmd/pane/viewOptions/toggleThumbnails")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
        g_folderWindow.SetDisplayMode(pane, FolderView::DisplayMode::Thumbnails);
        UpdatePaneMenuChecks();
        return true;
    }
    if (commandId == L"cmd/pane/viewOptions/togglePreviewPane")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
        g_folderWindow.TogglePreviewPane(pane);
        UpdatePaneMenuChecks();
        return true;
    }
    if (commandId == L"cmd/pane/viewOptions/toggleFilterBar")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
        const bool visible           = g_folderWindow.GetFilterBarVisible(pane);
        g_folderWindow.SetFilterBarVisible(pane, ! visible);
        UpdatePaneMenuChecks();
        return true;
    }
    if (commandId == L"cmd/pane/viewWith" && viewerActionId.has_value())
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandViewWith(pane, viewerActionId.value());
        return true;
    }
    if (commandId == L"cmd/pane/editWith" && editorActionId.has_value())
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandEditWith(pane, editorActionId.value());
        return true;
    }
    if (commandId == L"cmd/pane/userMenu" && userMenuActionId.has_value())
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandUserMenu(pane, userMenuActionId.value());
        return true;
    }
    if (commandId == L"cmd/pane/newFromShellTemplate")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandNewFromShellTemplate(pane, shellNewTemplateId.value_or(std::wstring_view{}));
        return true;
    }
    if (commandId == L"cmd/pane/filter")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandFilter(pane);
        return true;
    }
    if (commandId == L"cmd/pane/refresh")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.CommandRefresh(pane);
        return true;
    }
    if (commandId == L"cmd/pane/executeOpen")
    {
        return SendKeyToFocusedFolderView(VK_RETURN);
    }
    if (commandId == L"cmd/pane/selectCalculateDirectorySizeNext")
    {
        return SendKeyToFocusedFolderView(VK_SPACE);
    }
    if (commandId == L"cmd/pane/selectNext")
    {
        return SendKeyToFocusedFolderView(VK_INSERT);
    }
    if (commandId == L"cmd/pane/moveToRecycleBin")
    {
        return SendKeyToFocusedFolderView(VK_DELETE);
    }

    if (commandId == L"cmd/pane/goDriveRoot")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire))
        {
            return false;
        }

        wchar_t driveLetter = 0;
        if (driveRootLetter.has_value())
        {
            driveLetter = driveRootLetter.value();
        }
        else
        {
            driveLetter = L'\0';
        }

        if (driveLetter == 0)
        {
            const std::filesystem::path defaultRoot = GetDefaultFileSystemRoot();
            g_folderWindow.SetFolderPath(g_folderWindow.GetFocusedPane(), defaultRoot);
            return true;
        }

        std::wstring driveRoot;
        driveRoot.push_back(driveLetter);
        driveRoot.append(L":\\");

        const UINT driveType = GetDriveTypeW(driveRoot.c_str());
        if (driveType == DRIVE_NO_ROOT_DIR)
        {
            return true;
        }

        const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
        g_folderWindow.SetActivePane(pane);
        g_folderWindow.SetFolderPath(pane, std::filesystem::path(driveRoot));
        return true;
    }

    if (commandId == L"cmd/pane/hotPath")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire) || ! hotPathSlotIndex.has_value())
        {
            return true;
        }

        const int slotIdx = hotPathSlotIndex.value();
        if (g_settings.hotPaths.has_value())
        {
            const auto& slots = g_settings.hotPaths.value().slots;
            if (slotIdx >= 0 && slotIdx < static_cast<int>(slots.size()) && slots[static_cast<size_t>(slotIdx)].has_value())
            {
                const auto& slot = slots[static_cast<size_t>(slotIdx)].value();
                if (! slot.path.empty())
                {
                    const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
                    g_folderWindow.SetActivePane(pane);
                    g_folderWindow.SetFolderPath(pane, std::filesystem::path(slot.path));
                }
            }
        }
        return true;
    }

    if (commandId == L"cmd/pane/setHotPath")
    {
        if (! g_hFolderWindow.load(std::memory_order_acquire) || ! hotPathSlotIndex.has_value())
        {
            return true;
        }

        const int slotIdx                                      = hotPathSlotIndex.value();
        const FolderWindow::Pane pane                          = g_folderWindow.GetFocusedPane();
        const std::optional<std::filesystem::path> currentPath = g_folderWindow.GetCurrentPath(pane);
        if (! currentPath.has_value() || currentPath.value().empty())
        {
            return true;
        }

        const std::wstring newPath = currentPath.value().wstring();

        // Check if slot is occupied and ask for confirmation.
        if (g_settings.hotPaths.has_value())
        {
            const auto& slots = g_settings.hotPaths.value().slots;
            if (slotIdx >= 0 && slotIdx < static_cast<int>(slots.size()) && slots[static_cast<size_t>(slotIdx)].has_value())
            {
                const auto& existingSlot = slots[static_cast<size_t>(slotIdx)].value();
                if (! existingSlot.path.empty())
                {
                    const wchar_t digitChar    = (slotIdx < 9) ? static_cast<wchar_t>(L'1' + slotIdx) : L'0';
                    const std::wstring title   = LoadStringResource(nullptr, IDS_HOT_PATH_CONFIRM_TITLE);
                    const std::wstring message = FormatStringResource(nullptr, IDS_HOT_PATH_CONFIRM_REPLACE, digitChar, existingSlot.path);

                    const int result = MessageBoxCenteredText(ownerWindow, message, title, MB_YESNO | MB_ICONQUESTION);
                    if (result != IDYES)
                    {
                        return true;
                    }
                }
            }
        }

        // Assign the slot.
        if (! g_settings.hotPaths.has_value())
        {
            g_settings.hotPaths = Common::Settings::HotPathsSettings{};
        }

        Common::Settings::HotPathSlot slot{};
        if (slotIdx >= 0 && slotIdx < static_cast<int>(g_settings.hotPaths.value().slots.size()) &&
            g_settings.hotPaths.value().slots[static_cast<size_t>(slotIdx)].has_value())
        {
            slot = g_settings.hotPaths.value().slots[static_cast<size_t>(slotIdx)].value();
        }
        slot.path                                                       = newPath;
        g_settings.hotPaths.value().slots[static_cast<size_t>(slotIdx)] = std::move(slot);

        static_cast<void>(SettingsHotReload::SaveSettingsAndSchema(kAppId, g_settings));

        // Optionally open the prefs page.
        if (g_settings.hotPaths.value().openPrefsOnAssign)
        {
            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowPreferencesDialogHotPaths(ownerWindow, kAppId, g_settings, theme));
            g_folderWindow.ResyncNavigationShellFromFolderView(g_folderWindow.GetActivePane());
        }
        return true;
    }

    if (commandId == L"cmd/pane/hotPaths")
    {
        const AppTheme theme = ResolveConfiguredTheme();
        static_cast<void>(ShowPreferencesDialogHotPaths(ownerWindow, kAppId, g_settings, theme));
        g_folderWindow.ResyncNavigationShellFromFolderView(g_folderWindow.GetActivePane());
        return true;
    }

    const std::optional<unsigned int> wmCommandOpt = TryGetWmCommandId(commandId);
    if (wmCommandOpt.has_value())
    {
        const WPARAM wp = MAKEWPARAM(static_cast<WORD>(wmCommandOpt.value()), 0);
        SendMessageW(ownerWindow, WM_COMMAND, wp, 0);
        return true;
    }

    ShowCommandNotImplementedMessage(ownerWindow, commandId);
    return true;
}

[[nodiscard]] bool DispatchShortcutCommand(HWND ownerWindow, std::wstring_view commandId) noexcept
{
    return ExecuteCommandById(ownerWindow, commandId);
}

LRESULT OnFunctionBarInvoke(HWND ownerWindow, WPARAM wParam, LPARAM lParam) noexcept
{
    const uint32_t vk        = static_cast<uint32_t>(wParam);
    const uint32_t modifiers = static_cast<uint32_t>(lParam) & 0x7u;

    CancelFunctionBarPressedKeyClearTimer(ownerWindow);
    SetFunctionBarPressedKeyState(vk);
    ScheduleFunctionBarPressedKeyClear(ownerWindow, vk);

    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFunctionBarCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        return 0;
    }

    static_cast<void>(DispatchShortcutCommand(ownerWindow, commandOpt.value()));
    return 0;
}

[[nodiscard]] bool TryHandleShortcutKeyDown(HWND ownerWindow, const MSG& msg) noexcept
{
    if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN)
    {
        return false;
    }

    const uint32_t vk = static_cast<uint32_t>(msg.wParam);
    if (vk < VK_F1 || vk > VK_F12)
    {
        return false;
    }

    CancelFunctionBarPressedKeyClearTimer(ownerWindow);
    SetFunctionBarPressedKeyState(vk);
    ScheduleFunctionBarPressedKeyClear(ownerWindow, vk);

    const uint32_t modifiers = GetCurrentShortcutModifiers();

    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFunctionBarCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        // Consume unbound function keys so the focused control and static accelerator table
        // do not apply a hard-coded behavior when shortcuts are reconfigured.
        return true;
    }

    return DispatchShortcutCommand(ownerWindow, commandOpt.value());
}

[[nodiscard]] bool TryHandleFolderViewShortcutKeyDown(HWND ownerWindow, const MSG& msg) noexcept
{
    if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN)
    {
        return false;
    }

    const HWND folderWindow = g_hFolderWindow.load(std::memory_order_acquire);
    if (! folderWindow)
    {
        return false;
    }

    const uint32_t vk        = static_cast<uint32_t>(msg.wParam);
    const uint32_t modifiers = GetCurrentShortcutModifiers();

    const bool folderViewFocused = (g_folderWindow.GetFocusedFolderViewHwnd() != nullptr);
    if (! folderViewFocused)
    {
        // Avoid stealing NavigationView keyboard traversal and text entry, but allow modified chords
        // (Ctrl/Alt/Shift) to execute settings-backed FolderWindow commands when focus is inside the FolderWindow.
        if (modifiers == 0u || vk == static_cast<uint32_t>(VK_TAB))
        {
            return false;
        }

        const HWND focus = GetFocus();
        if (! focus || (focus != folderWindow && ! IsChild(folderWindow, focus)))
        {
            return false;
        }
    }

    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFolderViewCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        return false;
    }

    return DispatchShortcutCommand(ownerWindow, commandOpt.value());
}

[[nodiscard]] bool TryReclaimMainFolderViewFocusOnEscape(HWND ownerWindow, const MSG& msg) noexcept
{
    if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN)
    {
        return false;
    }
    if (msg.wParam != static_cast<WPARAM>(VK_ESCAPE))
    {
        return false;
    }
    if (! ownerWindow || ! IsWindowEnabled(ownerWindow))
    {
        return false;
    }
    if (! g_hFolderWindow.load(std::memory_order_acquire) || g_folderWindow.GetFocusedFolderViewHwnd() != nullptr)
    {
        return false;
    }
    if (g_folderWindow.IsFocusInNavigationView())
    {
        return false;
    }

    return g_folderWindow.TryRestoreActivePaneFolderViewFocus();
}

[[nodiscard]] bool IsCompareDirectoriesWindowMessageRoot(HWND root) noexcept
{
    if (! root)
    {
        return false;
    }

    std::array<wchar_t, 96> className{};
    const int len = GetClassNameW(root, className.data(), static_cast<int>(className.size()));
    if (len <= 0)
    {
        return false;
    }

    const std::wstring_view classNameView(className.data(), static_cast<size_t>(len));
    return OrdinalString::EqualsNoCase(classNameView, L"RedSalamander.CompareDirectoriesWindow");
}

[[nodiscard]] HWND TryGetFocusedFolderViewInCompareWindow(HWND compareWindow) noexcept
{
    if (! compareWindow)
    {
        return nullptr;
    }

    const HWND focus = GetFocus();
    if (! focus || (focus != compareWindow && ! IsChild(compareWindow, focus)))
    {
        return nullptr;
    }

    std::array<wchar_t, 96> className{};
    HWND current = focus;
    while (current && current != compareWindow)
    {
        className.fill(L'\0');
        const int len = GetClassNameW(current, className.data(), static_cast<int>(className.size()));
        if (len > 0 && OrdinalString::EqualsNoCase(std::wstring_view(className.data(), static_cast<size_t>(len)), L"RedSalamanderFolderView"))
        {
            return current;
        }

        current = GetParent(current);
    }

    return nullptr;
}

[[nodiscard]] bool IsFolderViewFocusedInCompareWindow(HWND compareWindow) noexcept
{
    return TryGetFocusedFolderViewInCompareWindow(compareWindow) != nullptr;
}

[[nodiscard]] bool DispatchShortcutCommandToCompareWindow(HWND compareWindow, std::wstring_view commandId) noexcept
{
    if (! compareWindow || commandId.empty())
    {
        return false;
    }

    const std::wstring_view originalCommandId = commandId;
    commandId                                 = CanonicalizeCommandId(commandId);

    if (commandId == L"cmd/pane/selection/unselectAll")
    {
        const HWND folderView = TryGetFocusedFolderViewInCompareWindow(compareWindow);
        if (! folderView)
        {
            return false;
        }

        SendMessageW(folderView, WM_KEYDOWN, static_cast<WPARAM>(VK_ESCAPE), 0);
        return true;
    }

    const std::optional<unsigned int> wmCommandOpt = TryGetWmCommandId(commandId);
    if (! wmCommandOpt.has_value())
    {
        auto owned = std::make_unique<std::wstring>(originalCommandId);
        static_cast<void>(PostMessagePayload(compareWindow, WndMsg::kCompareDirectoriesExecuteCommand, 0, std::move(owned)));
        return true;
    }

    const WPARAM wp = MAKEWPARAM(static_cast<WORD>(wmCommandOpt.value()), 0);
    SendMessageW(compareWindow, WM_COMMAND, wp, 0);
    return true;
}

[[nodiscard]] bool TryHandleCompareWindowShortcutKeyDown(HWND mainWindow, HWND compareWindow, const MSG& msg) noexcept
{
    if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN)
    {
        return false;
    }

    const uint32_t vk = static_cast<uint32_t>(msg.wParam);
    if (vk < VK_F1 || vk > VK_F12)
    {
        return false;
    }

    const uint32_t modifiers = GetCurrentShortcutModifiers();

    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFunctionBarCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        // Consume unbound function keys so the focused control and static accelerator table
        // do not apply a hard-coded behavior when shortcuts are reconfigured.
        return true;
    }

    const std::wstring_view commandId = CanonicalizeCommandId(commandOpt.value());
    if (commandId.starts_with(L"cmd/app/"))
    {
        return DispatchShortcutCommand(mainWindow, commandId);
    }

    return DispatchShortcutCommandToCompareWindow(compareWindow, commandId);
}

[[nodiscard]] bool TryHandleCompareWindowFolderViewShortcutKeyDown(HWND mainWindow, HWND compareWindow, const MSG& msg) noexcept
{
    if (msg.message != WM_KEYDOWN && msg.message != WM_SYSKEYDOWN)
    {
        return false;
    }

    const uint32_t vk        = static_cast<uint32_t>(msg.wParam);
    const uint32_t modifiers = GetCurrentShortcutModifiers();

    const bool folderViewFocused = IsFolderViewFocusedInCompareWindow(compareWindow);
    if (! folderViewFocused)
    {
        // Avoid stealing keyboard traversal and text entry when focus isn't inside a FolderView,
        // but allow modified chords (Ctrl/Alt/Shift) to execute settings-backed commands.
        if (modifiers == 0u || vk == static_cast<uint32_t>(VK_TAB))
        {
            return false;
        }
    }

    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFolderViewCommand(vk, modifiers);
    if (! commandOpt.has_value())
    {
        return false;
    }

    const std::wstring_view commandId = CanonicalizeCommandId(commandOpt.value());
    if (commandId.starts_with(L"cmd/app/"))
    {
        return DispatchShortcutCommand(mainWindow, commandId);
    }

    return DispatchShortcutCommandToCompareWindow(compareWindow, commandId);
}

[[nodiscard]] bool TryDismissAlertOverlaysOnEscape(HWND root) noexcept
{
    if (! root || IsWindow(root) == FALSE)
    {
        return false;
    }

    struct EnumState
    {
        bool dismissed = false;
    };

    EnumState state{};
    EnumChildWindows(root,
                     [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto* state = reinterpret_cast<EnumState*>(lParam);
        if (! state)
        {
            return TRUE;
        }

        constexpr wchar_t kAlertOverlayWindowClassName[] = L"RedSalamander.AlertOverlayWindow";

        wchar_t className[64]{};
        const int classLen = GetClassNameW(child, className, static_cast<int>(_countof(className)));
        if (classLen <= 0 || wcscmp(className, kAlertOverlayWindowClassName) != 0)
        {
            return TRUE;
        }

        if (IsWindowVisible(child) == FALSE)
        {
            return TRUE;
        }

        SendMessageW(child, WM_KEYDOWN, VK_ESCAPE, 0);
        if (IsWindowVisible(child) == FALSE)
        {
            state->dismissed = true;
            return FALSE; // stop enumeration
        }

        return TRUE;
    },
                     reinterpret_cast<LPARAM>(&state));
    return state.dismissed;
}

#if defined(RS_ASAN_DEBUG_BUILD)
constexpr std::wstring_view kRedSalamanderBuildFlavor = L"ASan Debug";
#elif defined(_DEBUG)
constexpr std::wstring_view kRedSalamanderBuildFlavor = L"Debug";
#else
constexpr std::wstring_view kRedSalamanderBuildFlavor = L"Release";
#endif

[[nodiscard]] constexpr bool IsRedSalamanderDiagnosticsEnabledByDefault() noexcept
{
#if defined(_DEBUG) || defined(RS_ASAN_DEBUG_BUILD)
    return true;
#else
    return false;
#endif
}

constexpr wchar_t kRedSalamanderHelpText[] =
    L"RedSalamander\r\n"
    L"\r\n"
    L"Usage:\r\n"
    L"  RedSalamander.exe [options]\r\n"
    L"\r\n"
    L"Options:\r\n"
    L"  -h, --help, /?                 Show this help.\r\n"
    L"  --crash-test                    Trigger crash handler test.\r\n"
    L"  --etw                           Enable RedSalamander Info/Perf/Debug ETW diagnostics in Release; Debug and ASan Debug enable them by default.\r\n"
    L"  --perf                          Write RedSalamander perf metrics to the default JSONL path in Release; Debug and ASan Debug enable it by default.\r\n"
    L"  --perf=PATH                     Write RedSalamander perf metrics to a custom JSONL path.\r\n"
#ifdef ENABLE_TESTS
    L"  --selftest                      Run all debug self-test suites and exit.\r\n"
    L"  --compare-selftest              Run CompareDirectories self-test suite.\r\n"
    L"  --commands-selftest             Run Commands self-test suite.\r\n"
    L"  --fileops-selftest              Run FileOperations self-test suite.\r\n"
    L"  --selftest-fail-fast            Stop after first failing self-test case.\r\n"
    L"  --selftest-case=NAME            Run the exact matching self-test case, or a case-prefix family when NAME ends in '_'.\r\n"
    L"  --selftest-list-cases           Emit self-test case inventory JSON and exit; combine with suite flags or --selftest-case.\r\n"
    L"  --selftest-timeout-multiplier=N Multiply self-test timeouts by finite N, clamped to [0.1, 100.0] (default 1.0).\r\n"
#endif
    L"\r\n";

[[nodiscard]] constexpr std::wstring_view GetRedSalamanderHelpText() noexcept
{
    return kRedSalamanderHelpText;
}

[[nodiscard]] std::wstring GetLocalTimestampForFileName() noexcept
{
    SYSTEMTIME st{};
    ::GetLocalTime(&st);
    return std::format(L"{0:04}-{1:02}-{2:02}_{3:02}{4:02}{5:02}", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
}

[[nodiscard]] std::filesystem::path GetDefaultPerfJsonlPath(std::wstring_view appName) noexcept
{
    std::filesystem::path root;
    wil::unique_cotaskmem_string localAppData;
    if (SUCCEEDED(::SHGetKnownFolderPath(FOLDERID_LocalAppData, KF_FLAG_CREATE, nullptr, &localAppData)) && localAppData)
    {
        root = localAppData.get();
    }
    else
    {
        std::error_code ec;
        root = std::filesystem::temp_directory_path(ec);
        if (ec)
        {
            root = L".";
        }
    }

    std::wstring fileName(appName);
    fileName.push_back(L'_');
    fileName.append(GetLocalTimestampForFileName());
    fileName.append(L".jsonl");

    return root / L"RedSalamander" / L"Perf" / fileName;
}
} // namespace

[[nodiscard]] bool DispatchShortcutCommandFromWindow(HWND ownerWindow, std::wstring_view commandId) noexcept
{
    return DispatchShortcutCommand(ownerWindow, commandId);
}

#ifdef ENABLE_TESTS
bool DebugDispatchShortcutCommand(HWND ownerWindow, std::wstring_view commandId) noexcept
{
    return DispatchShortcutCommand(ownerWindow, commandId);
}

void DebugSetRereadAssociationsSettingsForTest(const Common::Settings::Settings* settings) noexcept
{
    const std::scoped_lock lock(g_debugRereadAssociationsMutex);
    g_debugRereadAssociationsSettingsOverride = settings;
}

void DebugResetRereadAssociationsSnapshot() noexcept
{
    const std::scoped_lock lock(g_debugRereadAssociationsMutex);
    g_debugRereadAssociationsSnapshot      = {};
    g_debugRereadAssociationsSnapshotValid = false;
}

bool DebugGetRereadAssociationsSnapshot(RereadAssociationsDebugSnapshot& out) noexcept
{
    const std::scoped_lock lock(g_debugRereadAssociationsMutex);
    out = g_debugRereadAssociationsSnapshot;
    return g_debugRereadAssociationsSnapshotValid;
}

std::wstring_view DebugGetRedSalamanderHelpText() noexcept
{
    return GetRedSalamanderHelpText();
}

bool DebugIsRedSalamanderDiagnosticsEnabledByDefault() noexcept
{
    return IsRedSalamanderDiagnosticsEnabledByDefault();
}
#endif // ENABLE_TESTS

// Separate function with C++ objects (cannot use __try/__except)
static int RunApplication(HINSTANCE hInstance, int nCmdShow)
{
    const auto hasArg = [](PCWSTR needle) noexcept -> bool
    {
        if (! needle || needle[0] == L'\0')
        {
            return false;
        }

        int argc = 0;
        wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
        if (! argv || argc <= 1)
        {
            return false;
        }

        for (int i = 1; i < argc; ++i)
        {
            const wchar_t* arg = argv.get()[i];
            if (! arg || arg[0] == L'\0')
            {
                continue;
            }

            if (CompareStringOrdinal(arg, -1, needle, -1, TRUE) == CSTR_EQUAL)
            {
                return true;
            }
        }

        return false;
    };

    const auto getArgValue = [](PCWSTR prefix, std::wstring& value) noexcept -> bool
    {
        if (! prefix || prefix[0] == L'\0')
        {
            return false;
        }

        int argc = 0;
        wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
        if (! argv || argc <= 1)
        {
            return false;
        }

        const size_t prefixLen = wcslen(prefix);
        if (prefixLen == 0)
        {
            return false;
        }

        for (int i = 1; i < argc; ++i)
        {
            const wchar_t* arg = argv.get()[i];
            if (! arg)
            {
                continue;
            }

            if (_wcsnicmp(arg, prefix, prefixLen) != 0)
            {
                continue;
            }

            value = arg + prefixLen;
            return true;
        }

        return false;
    };

    if (hasArg(L"--etw"))
    {
        Debug::detail::SetRuntimeMonitorDiagnosticsEnabled(true);
        static_cast<void>(::SetEnvironmentVariableW(L"REDSALAMANDER_DIAGNOSTICS_ETW", L"1"));
    }

    std::wstring perfJsonlPath;
    const bool customPerfPath = getArgValue(L"--perf=", perfJsonlPath);
    if (customPerfPath || hasArg(L"--perf") || IsRedSalamanderDiagnosticsEnabledByDefault())
    {
        const std::filesystem::path perfPath =
            (customPerfPath && ! perfJsonlPath.empty()) ? std::filesystem::path(perfJsonlPath) : GetDefaultPerfJsonlPath(L"RedSalamander");
        Debug::Perf::ConfigureJsonlOutput(perfPath, L"RedSalamander", kRedSalamanderBuildFlavor);
    }

    const auto writeHelpText = [](std::wstring_view text) -> void
    {
        auto tryWrite = [](HANDLE handle, std::wstring_view msg) -> bool
        {
            if (! handle || handle == INVALID_HANDLE_VALUE)
            {
                return false;
            }

            DWORD mode = 0;
            if (GetConsoleMode(handle, &mode) != FALSE)
            {
                DWORD written = 0;
                return WriteConsoleW(handle, msg.data(), static_cast<DWORD>(msg.size()), &written, nullptr) != FALSE;
            }

            const int bytesNeeded = WideCharToMultiByte(CP_UTF8, 0, msg.data(), static_cast<int>(msg.size()), nullptr, 0, nullptr, nullptr);
            if (bytesNeeded <= 0)
            {
                return false;
            }

            std::string utf8;
            utf8.resize(static_cast<size_t>(bytesNeeded));
            const int converted = WideCharToMultiByte(CP_UTF8, 0, msg.data(), static_cast<int>(msg.size()), utf8.data(), bytesNeeded, nullptr, nullptr);
            if (converted != bytesNeeded)
            {
                return false;
            }

            DWORD written = 0;
            return WriteFile(handle, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr) != FALSE;
        };

        if (tryWrite(GetStdHandle(STD_OUTPUT_HANDLE), text))
        {
            return;
        }

        const bool hadConsole = GetConsoleWindow() != nullptr;
        if (! hadConsole)
        {
            if (AttachConsole(ATTACH_PARENT_PROCESS) == FALSE)
            {
                const DWORD err = GetLastError();
                if (err != ERROR_ACCESS_DENIED)
                {
                    static_cast<void>(AllocConsole());
                }
            }
        }

        wil::unique_handle conout(CreateFileW(L"CONOUT$", GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr));
        if (conout && tryWrite(conout.get(), text))
        {
            return;
        }

        std::wstring boxed(text);
        ShowFatalErrorDialog(nullptr, L"RedSalamander Help", boxed.c_str());
    };

    if (hasArg(L"--help") || hasArg(L"-h") || hasArg(L"/?"))
    {
        writeHelpText(GetRedSalamanderHelpText());
        return 0;
    }

    std::optional<Debug::Perf::Scope> startupPerf;
    startupPerf.emplace(L"App.Startup.UntilMessageLoop");
    startupPerf->SetDetail(kAppId);

    StartupMetrics::Initialize();

#ifdef ENABLE_TESTS
    QueueRedSalamanderMonitorLaunch();
#endif

    if (hasArg(L"--crash-test"))
    {
        CrashHandler::TriggerCrashTest();
    }

#ifndef ENABLE_TESTS
    std::wstring unsupportedSelfTestArg;
    if (hasArg(L"--selftest") || hasArg(L"--compare-selftest") || hasArg(L"--commands-selftest") || hasArg(L"--fileops-selftest") ||
        hasArg(L"--selftest-fail-fast") || hasArg(L"--selftest-list-cases") || getArgValue(L"--selftest-case=", unsupportedSelfTestArg) ||
        getArgValue(L"--selftest-timeout-multiplier=", unsupportedSelfTestArg))
    {
        Debug::Error(L"Self-test command-line arguments require ENABLE_TESTS.");
        return 2;
    }
#endif

#ifdef ENABLE_TESTS
    g_selfTestOptions                  = SelfTest::GetSelfTestOptions();
    g_selfTestOptions.failFast         = hasArg(L"--selftest-fail-fast");
    g_selfTestOptions.timeoutScale     = kSelfTestTimeoutMultiplierDefault;
    g_selfTestOptions.writeJsonSummary = true;
    g_selfTestOptions.caseFilter.clear();

    std::wstring multiplierArg;
    if (getArgValue(L"--selftest-timeout-multiplier=", multiplierArg))
    {
        const SelfTestTimeoutMultiplierParseResult parsed = ParseSelfTestTimeoutMultiplier(multiplierArg);
        if (! parsed.valid)
        {
            return 2;
        }

        g_selfTestOptions.timeoutScale = parsed.value;
    }

    std::wstring caseFilterArg;
    if (getArgValue(L"--selftest-case=", caseFilterArg))
    {
        g_selfTestOptions.caseFilter = std::move(caseFilterArg);
    }

    if (hasArg(L"--selftest-list-cases"))
    {
        const bool selfTestArgSelected  = hasArg(L"--selftest");
        const bool compareArgSelected   = hasArg(L"--compare-selftest");
        const bool commandsArgSelected  = hasArg(L"--commands-selftest");
        const bool fileOpsArgSelected   = hasArg(L"--fileops-selftest");
        const bool explicitSuiteRequest = selfTestArgSelected || compareArgSelected || commandsArgSelected || fileOpsArgSelected;

        const bool includeCompare  = selfTestArgSelected || compareArgSelected || ! explicitSuiteRequest;
        const bool includeCommands = selfTestArgSelected || commandsArgSelected || ! explicitSuiteRequest;
        const bool includeFileOps  = selfTestArgSelected || fileOpsArgSelected || ! explicitSuiteRequest;

        g_selfTestOptions.writeJsonSummary = false;
        g_selfTestOptions.listCasesOnly    = true;
        writeHelpText(BuildSelfTestCaseListJson(g_selfTestOptions, includeCompare, includeCommands, includeFileOps));
        return 0;
    }

    if (hasArg(L"--selftest"))
    {
        g_runFileOpsSelfTest            = true;
        g_runCompareDirectoriesSelfTest = true;
        g_runCommandsSelfTest           = true;
        HostSetAutoAcceptPrompts(true);
    }

    if (hasArg(L"--fileops-selftest"))
    {
        g_runFileOpsSelfTest = true;
        HostSetAutoAcceptPrompts(true);
    }

    if (hasArg(L"--compare-selftest"))
    {
        g_runCompareDirectoriesSelfTest = true;
    }

    if (hasArg(L"--commands-selftest"))
    {
        g_runCommandsSelfTest = true;
    }

    if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
    {
        if (! TryAcquireSelfTestRunMutex())
        {
            return 3;
        }

        SelfTest::GetSelfTestOptions() = g_selfTestOptions;
        SelfTest::RotateSelfTestRuns();
        ResetSelfTestRunState();
        SelfTest::InitSelfTestRun(g_selfTestOptions);
    }
#endif

    HRESULT comHr       = S_OK;
    bool comInitialized = false;
    {
        SplashScreen::IfExistSetText(L"Initializing COM...");
        Debug::Perf::Scope perf(L"App.Startup.CoInitializeEx");
        comHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        perf.SetHr(comHr);

        if (SUCCEEDED(comHr))
        {
            comInitialized = true;
        }
        else if (comHr == RPC_E_CHANGED_MODE)
        {
            Debug::Warning(L"CoInitializeEx returned RPC_E_CHANGED_MODE; COM is already initialized with a different threading model.");
        }
        else
        {
            Debug::Warning(L"CoInitializeEx failed: 0x{:08X}", comHr);
        }
    }

    const auto comCleanup = wil::scope_exit([&]
    {
        if (comInitialized)
        {
            CoUninitialize();
        }
    });

    const auto shutdownProcessSingletons = wil::scope_exit([]
    {
        // The splash screen owns a WindowHost on a worker UI thread. Close and join it
        // before the global DxUi host sweep so teardown cannot race across threads.
        SplashScreen::CloseIfExist();
        RedSalamander::DxUi::ShutdownAllWindowHostsForProcessExit();

        if (IsRunningAnySelfTest())
        {
            return;
        }

        // Process-lifetime singletons are intentionally leaked to avoid static destruction order hazards.
        // Explicitly release their resources before COM/CRT teardown to keep shutdown smooth and memory-bounded.
        RedSalamander::Ui::AnimationDispatcher::GetInstance().Shutdown();
        DirectoryInfoCache::GetInstance().Shutdown();
        IconCache::GetInstance().Shutdown();
    });

    const ThemeMode envTheme = GetInitialThemeModeFromEnvironment();
    g_themeMode              = envTheme;

    const std::filesystem::path themesDir = GetThemesDirectory();
    if (! themesDir.empty())
    {
        SplashScreen::IfExistSetText(L"Loading theme definitions...");
        Debug::Perf::Scope perf(L"App.Startup.LoadThemeDefinitions");
        perf.SetDetail(themesDir.native());
        Common::Settings::LoadThemeDefinitionsFromDirectory(themesDir, g_fileThemes);
    }

    HRESULT settingsHr = S_OK;
    Common::Settings::SettingsLoadRecoveryInfo settingsRecovery{};
    {
        SplashScreen::IfExistSetText(L"Loading app settings...");
        Debug::Perf::Scope perf(L"App.Startup.LoadSettings");
        perf.SetDetail(kAppId);
        settingsHr = Common::Settings::LoadSettingsWithRecoveryInfo(kAppId, g_settings, &settingsRecovery);
        perf.SetHr(settingsHr);
    }
    if (SUCCEEDED(settingsHr))
    {
        Localization::RegisterResourceOwner(kAppId, GetModuleHandleW(nullptr));
        Localization::ApplyLanguagePreference(GetLanguagePreferenceFromSettings(g_settings));
    }
    if (settingsHr == S_OK)
    {
        std::wstring_view themeId = g_settings.theme.currentThemeId;
        if (themeId.rfind(L"user/", 0) == 0)
        {
            const auto* def = FindThemeById(themeId);
            if (def)
            {
                themeId = def->baseThemeId;
            }
        }
        g_themeMode = ThemeModeFromThemeId(themeId);
    }
    else
    {
        g_settings.theme.currentThemeId = ThemeIdFromThemeMode(envTheme);
    }

    CrashQuarantine::OfferPluginDisableIfPreviousCrashDetected(g_settings);

#ifdef ENABLE_TESTS
    const bool runHeadlessCompareSelfTest = g_runCompareDirectoriesSelfTest && ! g_runFileOpsSelfTest && ! g_runCommandsSelfTest;
    const bool anySelfTest                = g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest;
#else
    const bool runHeadlessCompareSelfTest = false;
    const bool anySelfTest                = false;
#endif

    if (! anySelfTest && settingsRecovery.backedUp && ! settingsRecovery.backupPath.empty())
    {
        const std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_SETTINGS_RESTORED_DEFAULTS);
        const std::wstring message =
            settingsRecovery.reason == Common::Settings::SettingsLoadRecoveryReason::UnsupportedSchemaVersion
                ? FormatStringResource(nullptr,
                                       IDS_FMT_SETTINGS_RESTORED_DEFAULTS_UNSUPPORTED_SCHEMA_BACKUP,
                                       settingsRecovery.unsupportedSchemaVersion,
                                       settingsRecovery.settingsPath.wstring(),
                                       settingsRecovery.backupPath.wstring())
                : FormatStringResource(nullptr,
                                       IDS_FMT_SETTINGS_RESTORED_DEFAULTS_BACKUP,
                                       settingsRecovery.settingsPath.wstring(),
                                       settingsRecovery.backupPath.wstring());
        MessageBoxCenteredText(nullptr, message, title, MB_OK | MB_ICONWARNING);
    }

    const bool showSplash      = ! anySelfTest && ! runHeadlessCompareSelfTest && (! g_settings.startup.has_value() || g_settings.startup->showSplash);
    const auto setSplashStatus = [&](std::wstring_view status) noexcept
    {
        if (showSplash)
        {
            SplashScreen::IfExistSetText(status);
        }
    };

    if (showSplash)
    {
        setSplashStatus(L"Starting RedSalamander...");
        SplashScreen::BeginDelayedOpen(std::chrono::milliseconds(300), hInstance);
    }

    {
        if (showSplash)
        {
            setSplashStatus(L"Warming visual resources...");
        }
        Debug::Perf::Scope perf(L"App.Startup.QueueNavigationViewWarmup");
        const BOOL queued = runHeadlessCompareSelfTest ? TRUE : TrySubmitThreadpoolCallback([](PTP_CALLBACK_INSTANCE /*instance*/, void* /*context*/) noexcept {
            NavigationView::WarmSharedDeviceResources();
        }, nullptr, nullptr);
        perf.SetHr(queued ? S_OK : E_FAIL);
    }

    {
        if (showSplash)
        {
            setSplashStatus(L"Preparing keyboard shortcuts...");
        }
        Debug::Perf::Scope perf(L"App.Startup.Shortcuts.Initialize");
        ShortcutDefaults::EnsureShortcutsInitialized(g_settings);
        ReloadShortcutsFromSettings();
    }

    HRESULT pluginHr = S_OK;
    {
        if (showSplash)
        {
            setSplashStatus(L"Initializing file-system plugins...");
        }
        Debug::Perf::Scope perf(L"App.Startup.Plugins.Initialize.FileSystems");
        perf.SetDetail(kAppId);
        pluginHr = FileSystemPluginManager::GetInstance().Initialize(g_settings);
        perf.SetHr(pluginHr);
    }
    if (FAILED(pluginHr))
    {
        Debug::Warning(L"FileSystemPluginManager::Initialize failed (hr=0x{:08X})", static_cast<unsigned long>(pluginHr));
    }

    HRESULT viewerHr = S_OK;
    {
        if (showSplash)
        {
            setSplashStatus(L"Initializing viewer plugins...");
        }
        Debug::Perf::Scope perf(L"App.Startup.Plugins.Initialize.Viewers");
        perf.SetDetail(kAppId);
        viewerHr = ViewerPluginManager::GetInstance().Initialize(g_settings);
        perf.SetHr(viewerHr);
    }
    if (FAILED(viewerHr))
    {
        Debug::Warning(L"ViewerPluginManager::Initialize failed (hr=0x{:08X})", static_cast<unsigned long>(viewerHr));
    }

    {
        if (showSplash)
        {
            setSplashStatus(L"Applying directory cache settings...");
        }
        Debug::Perf::Scope perf(L"App.Startup.DirectoryInfoCache.ApplySettings");
        DirectoryInfoCache::GetInstance().ApplySettings(g_settings);
    }

#ifdef ENABLE_TESTS
    if (runHeadlessCompareSelfTest)
    {
        SelfTest::SelfTestSuiteResult compareResult;
        Debug::Info(L"CompareSelfTest: running (headless)");
        SelfTest::InitSelfTestRun(g_selfTestOptions);
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, L"CompareSelfTest: begin");
        SelfTest::AppendSelfTestTrace(L"CompareSelfTest: begin");
        g_selfTestExitCode |= CompareDirectoriesSelfTest::Run(g_selfTestOptions, &compareResult) ? 0 : 1;
        RecordSelfTestSuite(compareResult);
        if (g_selfTestExitCode != 0)
        {
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, L"CompareSelfTest: FAIL");
            SelfTest::AppendSelfTestTrace(L"CompareSelfTest: FAIL");
            if (! compareResult.failureMessage.empty())
            {
                SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, compareResult.failureMessage);
                SelfTest::AppendSelfTestTrace(compareResult.failureMessage);
            }
        }
        else
        {
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, L"CompareSelfTest: PASS");
            SelfTest::AppendSelfTestTrace(L"CompareSelfTest: PASS");
        }
        TraceSelfTestExitCode(L"CompareSelfTest: end", g_selfTestExitCode);
        FinalizeSelfTestRun();

        FileSystemPluginManager::GetInstance().Shutdown(g_settings);
        ViewerPluginManager::GetInstance().Shutdown(g_settings);

        startupPerf->SetHr(S_OK);
        startupPerf.reset();
        return g_selfTestExitCode;
    }
#endif

    // Perform application initialization:
    std::optional<HWND> hWnd;
    {
        if (showSplash)
        {
            setSplashStatus(L"Creating main window...");
        }
#ifdef ENABLE_TESTS
        if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
        {
            SelfTest::AppendSelfTestTrace(L"RunApplication: InitInstance begin");
        }
#endif
        Debug::Perf::Scope perf(L"App.Startup.InitInstance");
        hWnd = InitInstance(hInstance, nCmdShow);
        perf.SetHr(hWnd.has_value() ? S_OK : E_FAIL);
    }
#ifdef ENABLE_TESTS
    if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
    {
        SelfTest::AppendSelfTestTrace(hWnd.has_value() ? L"RunApplication: InitInstance returned hwnd" : L"RunApplication: InitInstance returned null");
    }
#endif
    if (! hWnd)
    {
        startupPerf->SetHr(E_FAIL);
        return FALSE;
    }

    wil::unique_haccel hAccelTable;
    {
        if (showSplash)
        {
            setSplashStatus(L"Loading menu accelerators...");
        }
#ifdef ENABLE_TESTS
        if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
        {
            SelfTest::AppendSelfTestTrace(L"RunApplication: LoadAccelerators begin");
        }
#endif
        Debug::Perf::Scope perf(L"App.Startup.LoadAccelerators");
        hAccelTable.reset(Localization::LoadAcceleratorsResource(hInstance, MAKEINTRESOURCEW(IDC_REDSALAMANDER)));
        perf.SetHr(hAccelTable ? S_OK : E_FAIL);
    }
#ifdef ENABLE_TESTS
    if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
    {
        SelfTest::AppendSelfTestTrace(L"RunApplication: entering message loop");
    }
#endif

    startupPerf->SetHr(S_OK);
    startupPerf.reset();

    MSG msg;
    bool altDown                  = false;
    bool altUsed                  = false;
    uint32_t functionBarModifiers = 0;
    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        const HWND root                   = msg.hwnd ? GetAncestor(msg.hwnd, GA_ROOT) : nullptr;
        const bool isMainWindowMessage    = (root == *hWnd);
        const bool isCompareWindowMessage = IsCompareDirectoriesWindowMessageRoot(root);

        if ((msg.message == WM_SYSKEYDOWN || msg.message == WM_KEYDOWN) && msg.wParam == static_cast<WPARAM>(VK_ESCAPE))
        {
            // Host alerts are closable by design when request.closable==TRUE (see Common/PlugInterfaces/Host.h).
            // Modeless overlay windows don't take focus, so route Esc explicitly.
            if (TryDismissAlertOverlaysOnEscape(root))
            {
                continue;
            }
        }

        const HWND prefsDialog = GetPreferencesDialogHandle();
        if (prefsDialog && root == prefsDialog)
        {
            if (IsDialogMessageW(prefsDialog, &msg))
            {
                continue;
            }
        }
        const HWND connectionsDialog = GetConnectionManagerDialogHandle();
        if (connectionsDialog && root == connectionsDialog)
        {
            if (IsDialogMessageW(connectionsDialog, &msg))
            {
                continue;
            }
        }

        if (g_mainMenuHandle)
        {
            if (isMainWindowMessage)
            {
                if ((msg.message == WM_SYSKEYDOWN || msg.message == WM_KEYDOWN) && msg.wParam == VK_MENU)
                {
                    altDown = true;
                    altUsed = false;
                }
                else if (altDown)
                {
                    if (msg.message == WM_SYSCHAR && msg.wParam != VK_MENU)
                    {
                        altUsed = true;

                        EnsureMenuHandles(*hWnd);
                        UpdatePaneMenuChecks();
                        if (! g_menuBarVisible)
                        {
                            g_menuBarTemporarilyShown = true;
                            AdjustLayout(*hWnd);
                        }
                        if (g_mainMenuBarHost.ActivateMnemonic(static_cast<wchar_t>(msg.wParam)))
                        {
                            altDown = false;
                            altUsed = false;
                            continue;
                        }
                    }
                    else if ((msg.message == WM_SYSKEYDOWN || msg.message == WM_KEYDOWN || msg.message == WM_CHAR) && msg.wParam != VK_MENU)
                    {
                        altUsed = true;
                    }
                    else if ((msg.message == WM_SYSKEYUP || msg.message == WM_KEYUP) && msg.wParam == VK_MENU)
                    {
                        if (! altUsed)
                        {
                            EnsureMenuHandles(*hWnd);
                            UpdatePaneMenuChecks();
                            if (! g_menuBarVisible)
                            {
                                g_menuBarTemporarilyShown = true;
                            }
                            AdjustLayout(*hWnd);
                            static_cast<void>(g_mainMenuBarHost.FocusFirstItem());

                            altDown = false;
                            altUsed = false;
                            continue;
                        }

                        altDown = false;
                        altUsed = false;
                    }
                }
            }
            else if ((msg.message == WM_SYSKEYUP || msg.message == WM_KEYUP) && msg.wParam == VK_MENU)
            {
                altDown = false;
                altUsed = false;
            }
        }

        if (isMainWindowMessage && g_hFolderWindow.load(std::memory_order_acquire))
        {
            const uint32_t currentModifiers = GetCurrentShortcutModifiers();
            if (currentModifiers != functionBarModifiers)
            {
                functionBarModifiers = currentModifiers;
                g_folderWindow.SetFunctionBarModifiers(currentModifiers);
            }

            const uint32_t vk = static_cast<uint32_t>(msg.wParam);
            if ((msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN) && vk >= VK_F1 && vk <= VK_F12)
            {
                CancelFunctionBarPressedKeyClearTimer(*hWnd);
                SetFunctionBarPressedKeyState(vk);
            }
            else if ((msg.message == WM_KEYUP || msg.message == WM_SYSKEYUP) && vk >= VK_F1 && vk <= VK_F12)
            {
                if (g_functionBarPressedKey.has_value() && g_functionBarPressedKey.value() == vk)
                {
                    ScheduleFunctionBarPressedKeyClear(*hWnd, vk);
                }
            }
        }

        if (isMainWindowMessage && g_hFolderWindow.load(std::memory_order_acquire))
        {
            const uint32_t vk = static_cast<uint32_t>(msg.wParam);
            if (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN)
            {
                if (g_folderWindow.HandleViewWidthAdjustKey(vk))
                {
                    continue;
                }

                if (vk == static_cast<uint32_t>(VK_ESCAPE) && g_fullScreenState.active)
                {
                    ToggleFullScreen(*hWnd);
                    continue;
                }
            }
        }

        const bool editFocused = (isMainWindowMessage || isCompareWindowMessage) && IsEditControlFocused();
        if ((isMainWindowMessage || isCompareWindowMessage) && editFocused)
        {
            const uint32_t vk = static_cast<uint32_t>(msg.wParam);
            if ((msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN) && vk == static_cast<uint32_t>(VK_F1))
            {
                const uint32_t modifiers = GetCurrentShortcutModifiers();
                if (modifiers == 0u)
                {
                    const std::optional<std::wstring_view> commandOpt = g_shortcutManager.FindFunctionBarCommand(vk, modifiers);
                    if (commandOpt.has_value() && CanonicalizeCommandId(commandOpt.value()) == L"cmd/app/showShortcuts")
                    {
                        static_cast<void>(DispatchShortcutCommand(*hWnd, commandOpt.value()));
                        continue;
                    }
                }
            }
        }

        if (isMainWindowMessage && ! editFocused && TryReclaimMainFolderViewFocusOnEscape(*hWnd, msg))
        {
            continue;
        }

        if (isMainWindowMessage && ! editFocused)
        {
            if (TryHandleShortcutKeyDown(*hWnd, msg))
            {
                continue;
            }

            if (TryHandleFolderViewShortcutKeyDown(*hWnd, msg))
            {
                continue;
            }
        }
        else if (isCompareWindowMessage && ! editFocused)
        {
            if (TryHandleCompareWindowShortcutKeyDown(*hWnd, root, msg))
            {
                continue;
            }

            if (TryHandleCompareWindowFolderViewShortcutKeyDown(*hWnd, root, msg))
            {
                continue;
            }
        }

        if (! isMainWindowMessage || editFocused || ! TranslateAccelerator(*hWnd, hAccelTable.get(), &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    const int exitCode = static_cast<int>(msg.wParam);
#ifdef ENABLE_TESTS
    if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
    {
        SelfTest::AppendSelfTestTrace(std::format(L"RunApplication: message loop exit={}", exitCode));
        FinalizeSelfTestRun();
    }
#endif
    return exitCode;
}

namespace
{
void BuildFatalExceptionMessage(HINSTANCE hInstance, const wchar_t* exceptionName, DWORD exceptionCode, wchar_t* outMessage, size_t outMessageChars) noexcept
{
    if (! outMessage || outMessageChars == 0)
    {
        return;
    }
    outMessage[0] = L'\0';

    const std::wstring msg = FormatStringResource(hInstance, IDS_FATAL_EXCEPTION_FMT, exceptionName, static_cast<unsigned>(exceptionCode));
    if (! msg.empty())
    {
        static_cast<void>(wcsncpy_s(outMessage, outMessageChars, msg.c_str(), _TRUNCATE));
        return;
    }

    const std::ptrdiff_t maxChars = static_cast<std::ptrdiff_t>(outMessageChars - 1);
    const auto r = std::format_to_n(outMessage, maxChars, L"Fatal Exception ({0}, 0x{1:08X}).", exceptionName, static_cast<unsigned>(exceptionCode));
    const std::ptrdiff_t written             = (r.size < 0) ? 0 : ((r.size > maxChars) ? maxChars : r.size);
    outMessage[static_cast<size_t>(written)] = L'\0';
}
} // namespace

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, [[maybe_unused]] _In_opt_ HINSTANCE hPrevInstance, [[maybe_unused]] _In_ LPWSTR lpCmdLine, _In_ int nCmdShow)
{
    // Use SEH to catch all exceptions (no C++ objects in this scope)
    CrashHandler::Install();
    __try
    {
        return RunApplication(hInstance, nCmdShow);
    }
    __except (CrashHandler::WriteDumpForException(GetExceptionInformation()))
    {
        // Handle all exceptions including SEH exceptions
        const DWORD exceptionCode    = GetExceptionCode();
        const wchar_t* exceptionName = exception::GetExceptionName(exceptionCode);
#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            wchar_t exceptionTrace[160]{};
            if (SUCCEEDED(StringCchPrintfW(exceptionTrace,
                                           std::size(exceptionTrace),
                                           L"wWinMain: unhandled exception code=0x%08X name=%s",
                                           static_cast<unsigned>(exceptionCode),
                                           exceptionName ? exceptionName : L"unknown")))
            {
                SelfTest::AppendSelfTestTrace(exceptionTrace);
            }
            else
            {
                SelfTest::AppendSelfTestTrace(L"wWinMain: unhandled exception");
            }
        }
#endif
        wchar_t errorMsg[512]{};
        BuildFatalExceptionMessage(hInstance, exceptionName, exceptionCode, errorMsg, std::size(errorMsg));
        OutputDebugStringW(errorMsg);

        wchar_t caption[256]{};
        const int captionLength = LoadStringW(hInstance, IDS_FATAL_ERROR_CAPTION, caption, static_cast<int>(std::size(caption)));
        ShowFatalErrorDialog(nullptr, captionLength > 0 ? caption : L"", errorMsg);

        return -1;
    }
}

// Saves instance handle and creates main window
std::optional<HWND> InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    g_hInstance = hInstance; // Store instance handle in our global variable

#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"InitInstance: begin");
    }
#endif

    std::wstring szWindowClass(MAX_LOADSTRING, L'\0');
    LoadStringW(hInstance, IDC_REDSALAMANDER, szWindowClass.data(), MAX_LOADSTRING);

    WNDCLASSEXW wcex{};
    wcex.cbSize        = sizeof(WNDCLASSEX);
    wcex.style         = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc   = WndProc;
    wcex.cbClsExtra    = 0;
    wcex.cbWndExtra    = 0;
    wcex.hInstance     = hInstance;
    wcex.hIcon         = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_REDSALAMANDER));
    wcex.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wcex.lpszMenuName  = nullptr;
    wcex.lpszClassName = szWindowClass.c_str();
    wcex.hIconSm       = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    ATOM atom = 0;
    {
        Debug::Perf::Scope perf(L"App.Startup.InitInstance.RegisterClassExW");
        perf.SetDetail(szWindowClass);
        atom = RegisterClassExW(&wcex);
        perf.SetHr(atom ? S_OK : HRESULT_FROM_WIN32(GetLastError()));
    }
    if (! atom)
    {
        Debug::ErrorWithLastError(L"RegisterClassExW failed");
#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            SelfTest::AppendSelfTestTrace(std::format(L"InitInstance: RegisterClassExW failed gle={}", GetLastError()));
        }
#endif
        return std::nullopt;
    }
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"InitInstance: RegisterClassExW ok");
    }
#endif

    std::wstring szTitle(MAX_LOADSTRING, L'\0');
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle.data(), MAX_LOADSTRING);

    wil::unique_hwnd hWnd;
    {
        Debug::Perf::Scope perf(L"App.Startup.InitInstance.CreateWindowW");
        perf.SetDetail(szTitle);
        hWnd.reset(CreateWindowW(szWindowClass.c_str(),
                                 szTitle.c_str(),
                                 WS_OVERLAPPEDWINDOW,
                                 CW_USEDEFAULT,
                                 CW_USEDEFAULT,
                                 CW_USEDEFAULT,
                                 CW_USEDEFAULT,
                                 nullptr,
                                 nullptr,
                                 hInstance,
                                 nullptr));
        perf.SetHr(hWnd ? S_OK : HRESULT_FROM_WIN32(GetLastError()));
    }

    if (! hWnd)
    {
        Debug::ErrorWithLastError(L"CreateWindowW failed");
#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            SelfTest::AppendSelfTestTrace(std::format(L"InitInstance: CreateWindowW failed gle={}", GetLastError()));
        }
#endif
        return std::nullopt;
    }
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(std::format(L"InitInstance: CreateWindowW ok hwnd=0x{:X}", reinterpret_cast<uintptr_t>(hWnd.get())));
    }
#endif

    StartupMetrics::MarkFirstWindowCreated(kMainWindowId);

    int showCmd = nCmdShow;
    {
        Debug::Perf::Scope perf(L"App.Startup.InitInstance.RestoreWindowPlacement");
        perf.SetDetail(kMainWindowId);

        const auto it = g_settings.windows.find(kMainWindowId);
        if (it != g_settings.windows.end())
        {
            const UINT dpi                                     = GetDpiForWindow(hWnd.get());
            const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, dpi);

            SetWindowPos(hWnd.get(),
                         nullptr,
                         normalized.bounds.x,
                         normalized.bounds.y,
                         normalized.bounds.width,
                         normalized.bounds.height,
                         SWP_NOZORDER | SWP_NOACTIVATE);

            if (normalized.state == Common::Settings::WindowState::Maximized)
            {
                showCmd = SW_MAXIMIZE;
            }
            else
            {
                showCmd = SW_SHOWNORMAL;
            }
        }
    }
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"InitInstance: window placement restored");
    }
#endif

    if (SplashScreen::Exist())
    {
        SplashScreen::SetOwner(hWnd.get());
    }
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"InitInstance: splash owner set");
    }
#endif

    {
        Debug::Perf::Scope perf(L"App.Startup.InitInstance.ShowUpdateWindow");
        perf.SetValue0(static_cast<uint64_t>(showCmd));

#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            SelfTest::AppendSelfTestTrace(L"InitInstance: ShowWindow begin");
        }
#endif
        ShowWindow(hWnd.get(), showCmd);
        g_mainMenuBarHost.UpdateLayout();
        if (const HWND menuBarWindow = g_mainMenuBarHost.GetHwnd(); menuBarWindow && IsWindowVisible(menuBarWindow) != FALSE)
        {
            RedrawWindow(menuBarWindow, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE);
        }
#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            SelfTest::AppendSelfTestTrace(L"InitInstance: ShowWindow ok");
            SelfTest::AppendSelfTestTrace(L"InitInstance: UpdateWindow begin");
        }
#endif
#ifdef ENABLE_TESTS
        if (! IsRunningAnySelfTest())
#endif
        {
            UpdateWindow(hWnd.get());
        }
#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            SelfTest::AppendSelfTestTrace(L"InitInstance: UpdateWindow skipped for selftest");
        }
#endif
    }
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"InitInstance: ShowWindow/UpdateWindow ok");
    }
#endif
    static_cast<void>(PostMessageW(hWnd.get(), WndMsg::kAppStartupInputReady, 0, 0));

    {
        Debug::Perf::Scope perf(L"App.Startup.InitInstance.QueueIconCacheWarm");
        const BOOL queued = TrySubmitThreadpoolCallback(
            [](PTP_CALLBACK_INSTANCE /*instance*/, void* /*context*/) noexcept { IconCache::GetInstance().WarmCommonExtensions(); }, nullptr, nullptr);
        perf.SetHr(queued ? S_OK : E_FAIL);
    }

#ifdef ENABLE_TESTS
    DBGOUT_INFO(L"RedSalamander started, version {}\n", VERSINFO_VERSION);

    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
    for (int i = 0; argv && i < argc; ++i)
    {
        const wchar_t* arg = argv.get()[i];
        DBGOUT_ERROR(L"  argv[{}] = ({})\n", i, arg ? arg : L"(null)");
    }

    // for (int j = 0; j < 1000; j++)
    //{
    //     DBGOUT_WARNING(L"  len ({})\n", j);
    // }

#endif

    return hWnd.release();
}

static void AdjustLayout(HWND hWnd)
{
    const HWND folderWindow = g_hFolderWindow.load(std::memory_order_acquire);
    if (! hWnd)
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(hWnd, &client))
    {
        return;
    }

    const int width  = client.right - client.left;
    const int height = client.bottom - client.top;
    g_mainMenuBarHost.UpdateLayout();

    if (! folderWindow)
    {
        return;
    }

    const int menuBarHeight = (g_menuBarVisible || g_menuBarTemporarilyShown) ? g_mainMenuBarHost.GetVisibleHeightPx() : 0;
    MoveWindow(folderWindow, client.left, client.top + menuBarHeight, width, (std::max)(0, height - menuBarHeight), TRUE);
}

static void ApplyAppTheme(HWND hWnd)
{
    const AppTheme theme    = ResolveConfiguredTheme();
    const bool windowActive = GetActiveWindow() == hWnd;
    ApplyTitleBarTheme(hWnd, theme, windowActive);
    ApplyWindowBackdropTheme(hWnd, theme, WindowBackdropTarget::Primary);

    MessageBoxTheme messageBoxTheme{};
    messageBoxTheme.enabled      = true;
    messageBoxTheme.useDarkMode  = theme.dark;
    messageBoxTheme.highContrast = theme.highContrast;
    messageBoxTheme.background   = theme.windowBackground;
    messageBoxTheme.text         = theme.menu.text;
    SetDefaultMessageBoxTheme(messageBoxTheme);

    if (g_hFolderWindow.load(std::memory_order_acquire))
    {
        g_folderWindow.ApplyTheme(theme);
    }

    UpdateShortcutsWindowTheme(theme);
    UpdateConnectionManagerWindowsTheme(theme);
    UpdateConnectionCredentialPromptWindowsTheme(theme);
    UpdateCompareDirectoriesWindowsTheme(theme);
    UpdateFindFilesWindowsTheme(theme);
    UpdatePreferencesWindowsTheme(theme);
    UpdatePluginConfigurationWindowsTheme(theme);

    UpdateThemeMenuChecks();
    g_mainMenuBarHost.SetTheme(theme);
    g_mainMenuBarHost.SyncMenuModel();
    SendMessageW(hWnd, WM_SIZE, 0, 0);
    RedrawWindow(hWnd, nullptr, nullptr, RDW_INVALIDATE | RDW_FRAME | RDW_ERASE | RDW_ALLCHILDREN);

    if (const HWND prefs = GetPreferencesDialogHandle(); prefs && IsWindow(prefs))
    {
        PostMessageW(prefs, WM_THEMECHANGED, 0, 0);
    }
}

[[nodiscard]] std::vector<std::wstring_view> BuildThemeCycleIds()
{
    constexpr std::array<std::wstring_view, 5> kBuiltInThemeIds = {
        std::wstring_view{L"builtin/system"},
        std::wstring_view{L"builtin/light"},
        std::wstring_view{L"builtin/dark"},
        std::wstring_view{L"builtin/rainbow"},
        std::wstring_view{L"builtin/highContrast"},
    };

    const CustomThemeGroups customThemes = CollectCustomThemeGroups();

    std::vector<std::wstring_view> ids;
    ids.reserve(kBuiltInThemeIds.size() + customThemes.fileThemes.size() + customThemes.settingsThemes.size());
    ids.insert(ids.end(), kBuiltInThemeIds.begin(), kBuiltInThemeIds.end());

    for (const auto* def : customThemes.fileThemes)
    {
        ids.push_back(def->id);
    }
    for (const auto* def : customThemes.settingsThemes)
    {
        ids.push_back(def->id);
    }

    return ids;
}

void ApplyThemeId(HWND hWnd, std::wstring_view themeId)
{
    g_settings.theme.currentThemeId = std::wstring(themeId);
    if (const auto* def = FindThemeById(themeId))
    {
        g_themeMode = ThemeModeFromThemeId(def->baseThemeId);
    }
    else
    {
        g_themeMode = ThemeModeFromThemeId(themeId);
    }

    ApplyAppTheme(hWnd);
}

void SelectAdjacentTheme(HWND hWnd, int direction)
{
    const std::vector<std::wstring_view> ids = BuildThemeCycleIds();
    if (ids.empty())
    {
        return;
    }

    const auto findTheme = [&](std::wstring_view id)
    {
        return std::find(ids.begin(), ids.end(), id);
    };

    auto currentIt = findTheme(g_settings.theme.currentThemeId);
    if (currentIt == ids.end())
    {
        const std::wstring fallbackThemeId = ThemeIdFromThemeMode(g_themeMode);
        currentIt                          = findTheme(fallbackThemeId);
    }

    const size_t currentIndex = currentIt != ids.end() ? static_cast<size_t>(std::distance(ids.begin(), currentIt)) : 0u;
    const size_t count        = ids.size();
    const size_t nextIndex    = direction >= 0 ? (currentIndex + 1u) % count : (currentIndex + count - 1u) % count;
    ApplyThemeId(hWnd, ids[nextIndex]);
}

[[nodiscard]] bool HasOpenItemPropertiesWindow() noexcept
{
    const HWND props = FindWindowW(kItemPropertiesWindowClassName, nullptr);
    return props && IsWindow(props);
}

void ApplyCurrentSettingsToRunningApp(HWND hWnd) noexcept
{
    if (! hWnd)
    {
        return;
    }

    Localization::ApplyLanguagePreference(GetLanguagePreferenceFromSettings(g_settings));
    UpdateThemeModeFromCurrentSettings();
    if (! ReloadLocalizedMainMenu(hWnd, true))
    {
        EnsureMenuHandles(hWnd);
        RebuildThemeMenuDynamicItems(hWnd);
    }
    ApplyAppTheme(hWnd);

    const Common::Settings::MainMenuState menu = g_settings.mainMenu.value_or(Common::Settings::MainMenuState{});

    if (menu.menuBarVisible != g_menuBarVisible)
    {
        g_menuBarVisible          = menu.menuBarVisible;
        g_menuBarTemporarilyShown = false;
    }

    if (menu.functionBarVisible != g_functionBarVisible)
    {
        g_functionBarVisible = menu.functionBarVisible;
        g_folderWindow.SetFunctionBarVisible(g_functionBarVisible);
    }

    UpdatePaneMenuChecks();
    g_mainMenuBarHost.SyncMenuModel();
    AdjustLayout(hWnd);

    DirectoryInfoCache::GetInstance().ApplySettings(g_settings);

    if (g_hFolderWindow.load(std::memory_order_acquire))
    {
        const Common::Settings::FolderPane* leftSettings  = nullptr;
        const Common::Settings::FolderPane* rightSettings = nullptr;
        uint32_t folderHistoryMax                         = 20u;
        bool showHiddenFiles                              = true;
        bool showSystemFiles                              = true;

        if (g_settings.folders)
        {
            const auto& folders = *g_settings.folders;
            folderHistoryMax    = folders.historyMax;
            showHiddenFiles     = folders.showHiddenFiles;
            showSystemFiles     = folders.showSystemFiles;

            for (const auto& item : folders.items)
            {
                if (item.slot == kLeftPaneSlot)
                {
                    leftSettings = &item;
                }
                else if (item.slot == kRightPaneSlot)
                {
                    rightSettings = &item;
                }
            }
        }

        folderHistoryMax = std::clamp(folderHistoryMax, 1u, 50u);
        g_folderWindow.SetFolderHistoryMax(folderHistoryMax);
        g_folderWindow.SetShowHiddenFiles(showHiddenFiles);
        g_folderWindow.SetShowSystemFiles(showSystemFiles);

        auto applyPane = [&](FolderWindow::Pane pane, const Common::Settings::FolderPane* settingsPane)
        {
            FolderView::DisplayMode displayMode     = FolderView::DisplayMode::Brief;
            FolderView::SortBy sortBy               = FolderView::SortBy::Name;
            FolderView::SortDirection sortDirection = FolderView::SortDirection::Ascending;
            bool fileExtensionsVisible              = true;
            bool navigationBarVisible               = true;
            bool filterBarVisible                   = false;
            bool statusBarVisible                   = true;

            if (settingsPane)
            {
                displayMode           = DisplayModeFromSettings(settingsPane->view.display);
                sortBy                = SortByFromSettings(settingsPane->view.sortBy);
                sortDirection         = SortDirectionFromSettings(settingsPane->view.sortDirection);
                fileExtensionsVisible = settingsPane->view.fileExtensionsVisible;
                if (settingsPane->view.thumbnailsVisible)
                {
                    displayMode = FolderView::DisplayMode::Thumbnails;
                }
                navigationBarVisible  = settingsPane->view.navigationBarVisible;
                filterBarVisible      = settingsPane->view.filterBarVisible;
                statusBarVisible      = settingsPane->view.statusBarVisible;
            }

            g_folderWindow.SetFileExtensionsVisible(pane, fileExtensionsVisible);
            g_folderWindow.SetNavigationBarVisible(pane, navigationBarVisible);
            g_folderWindow.SetFilterBarVisible(pane, filterBarVisible);
            g_folderWindow.SetStatusBarVisible(pane, statusBarVisible);
            g_folderWindow.SetSort(pane, sortBy, sortDirection);
            g_folderWindow.SetDisplayMode(pane, displayMode);
        };

        applyPane(FolderWindow::Pane::Left, leftSettings);
        applyPane(FolderWindow::Pane::Right, rightSettings);

        UpdatePaneMenuChecks();
    }

    ReloadShortcutsFromSettings();
    if (g_settings.shortcuts.has_value())
    {
        UpdateShortcutsWindowData(g_settings.shortcuts.value(), g_shortcutManager);
    }
}

void RefreshRunningPluginsFromSettings(HWND hWnd) noexcept
{
    if (! hWnd)
    {
        return;
    }

    static_cast<void>(FileSystemPluginManager::GetInstance().Refresh(g_settings));
    static_cast<void>(ViewerPluginManager::GetInstance().Refresh(g_settings));
    static_cast<void>(g_folderWindow.ReloadFileSystemPlugins());
    RebuildPluginsMenuDynamicItems(hWnd);
}

[[nodiscard]] std::vector<std::wstring_view> CollectRuntimeSettingsWindowIds() noexcept
{
    std::vector<std::wstring_view> runtimeWindowIds;
    runtimeWindowIds.reserve(6);
    runtimeWindowIds.push_back(kMainWindowId);
    if (const HWND prefs = GetPreferencesDialogHandle(); prefs && IsWindow(prefs))
    {
        runtimeWindowIds.push_back(kPreferencesWindowId);
    }
    if (const HWND connections = GetConnectionManagerDialogHandle(); connections && IsWindow(connections))
    {
        runtimeWindowIds.push_back(kConnectionManagerWindowId);
    }
    if (const HWND shortcuts = GetShortcutsWindowHandle(); shortcuts && IsWindow(shortcuts))
    {
        runtimeWindowIds.push_back(kShortcutsWindowId);
    }
    if (const HWND findFiles = GetFindFilesWindowHandle(); findFiles && IsWindow(findFiles))
    {
        runtimeWindowIds.push_back(kFindFilesWindowId);
    }
    if (HasOpenItemPropertiesWindow())
    {
        runtimeWindowIds.push_back(kItemPropertiesWindowId);
    }
    return runtimeWindowIds;
}

[[nodiscard]] bool SettingsContainPanePath(
    const std::optional<Common::Settings::FoldersSettings>& folders, std::wstring_view slot, const std::optional<std::filesystem::path>& expected) noexcept
{
    if (! expected.has_value())
    {
        return true;
    }
    if (! folders.has_value())
    {
        return false;
    }

    for (const Common::Settings::FolderPane& pane : folders->items)
    {
        if (pane.slot == slot && pane.current == expected.value())
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] HRESULT LoadRereadAssociationsSettings(Common::Settings::Settings& settings) noexcept
{
#ifdef ENABLE_TESTS
    {
        const std::scoped_lock lock(g_debugRereadAssociationsMutex);
        if (g_debugRereadAssociationsSettingsOverride)
        {
            settings = *g_debugRereadAssociationsSettingsOverride;
            return S_OK;
        }
    }
#endif

    return Common::Settings::TryLoadSettingsNoRecovery(kAppId, settings);
}

#ifdef ENABLE_TESTS
void DebugPublishRereadAssociationsSnapshot(const RereadAssociationsDebugSnapshot& snapshot) noexcept
{
    const std::scoped_lock lock(g_debugRereadAssociationsMutex);
    g_debugRereadAssociationsSnapshot      = snapshot;
    g_debugRereadAssociationsSnapshotValid = true;
}
#endif

void RereadAssociations(HWND hWnd) noexcept
{
    Debug::Perf::Scope perf(L"rereadAssociations.total_us");

#ifdef ENABLE_TESTS
    RereadAssociationsDebugSnapshot snapshot{};
    snapshot.attempted                 = true;
    snapshot.leftRefreshCountBefore    = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    snapshot.rightRefreshCountBefore   = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Right);
    snapshot.associationIconCacheSizeBefore = DebugGetAssociationIconCacheSize();
#endif

    Common::Settings::Settings diskSettings;
    const HRESULT loadHr = LoadRereadAssociationsSettings(diskSettings);
#ifdef ENABLE_TESTS
    snapshot.hr     = loadHr;
    snapshot.loaded = (loadHr == S_OK);
#endif
    if (loadHr != S_OK)
    {
        perf.SetHr(loadHr);
        Debug::Warning(L"RereadAssociations: settings reload failed (hr=0x{:08X})", static_cast<unsigned long>(loadHr));
        SettingsHotReload::ShowInvalidReloadAlert(Common::Settings::GetSettingsPath(kAppId));
#ifdef ENABLE_TESTS
        DebugPublishRereadAssociationsSnapshot(snapshot);
#endif
        return;
    }

    const FolderWindow::Pane activePaneBefore = g_folderWindow.GetActivePane();
    const std::optional<std::filesystem::path> liveLeftBefore  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
    const std::optional<std::filesystem::path> liveRightBefore = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);

    Common::Settings::Settings runtimeSettings = g_settings;
    CaptureRuntimeSettings(runtimeSettings, hWnd);
    g_settings = SettingsHotReload::MergeDiskSettingsWithRuntimeSession(diskSettings, runtimeSettings, CollectRuntimeSettingsWindowIds());
    ShortcutDefaults::EnsureShortcutsInitialized(g_settings);

    SettingsHotReload::ClearInvalidReloadAlert();
    ApplyCurrentSettingsToRunningApp(hWnd);
    RefreshRunningPluginsFromSettings(hWnd);

    IconCache::GetInstance().Clear();
    IconCache::GetInstance().ClearAssociationCache();
#ifdef ENABLE_TESTS
    snapshot.associationIconCacheSizeAfterClear = DebugGetAssociationIconCacheSize();
#endif

    EnsureMenuHandles(hWnd);
    bool viewWithMenuRebuilt = false;
    bool editWithMenuRebuilt = false;
    bool userMenuRebuilt     = false;
    if (g_viewWithMenu)
    {
        RebuildFileActionMenuDynamicItems(g_viewWithMenu, true);
        viewWithMenuRebuilt = true;
    }
    if (g_editWithMenu)
    {
        RebuildFileActionMenuDynamicItems(g_editWithMenu, false);
        editWithMenuRebuilt = true;
    }
    if (g_userMenu)
    {
        RebuildUserMenuDynamicItems(g_userMenu);
        userMenuRebuilt = true;
    }
    if (g_newTemplateMenu)
    {
        RebuildShellNewTemplateMenuDynamicItems(g_newTemplateMenu);
    }
    UpdatePaneMenuChecks();
    g_mainMenuBarHost.SyncMenuModel();

    g_folderWindow.CommandRefresh(FolderWindow::Pane::Left);
    g_folderWindow.CommandRefresh(FolderWindow::Pane::Right);
    g_folderWindow.SetActivePane(activePaneBefore);

    Common::Settings::SettingsFileStamp stamp{};
    if (Common::Settings::TryGetSettingsFileStamp(kAppId, stamp) == S_OK)
    {
        SettingsHotReload::MarkAppliedStamp(stamp);
    }
    SettingsHotReload::NotifyParticipants();

#ifdef ENABLE_TESTS
    snapshot.viewerActionCount               = g_settings.fileActions.viewers.actions.size();
    snapshot.editorActionCount               = g_settings.fileActions.editors.actions.size();
    snapshot.userMenuActionCount             = g_settings.userMenu.actions.size();
    snapshot.viewerExtensionMappingCount     = g_settings.fileActions.viewers.associations.size();
    snapshot.fileSystemExtensionMappingCount = g_settings.extensions.openWithFileSystemByExtension.size();
    snapshot.leftRefreshCountAfter              = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Left);
    snapshot.rightRefreshCountAfter             = g_folderWindow.DebugGetForceRefreshCount(FolderWindow::Pane::Right);
    snapshot.dynamicFileActionMenusRebuilt      = viewWithMenuRebuilt && editWithMenuRebuilt;
    snapshot.userMenuRebuilt                    = userMenuRebuilt;
    snapshot.pluginsRefreshed                   = hWnd && IsWindow(hWnd) != FALSE;
    snapshot.runtimeFoldersPreserved =
        SettingsContainPanePath(g_settings.folders, kLeftPaneSlot, liveLeftBefore) && SettingsContainPanePath(g_settings.folders, kRightPaneSlot, liveRightBefore);
    DebugPublishRereadAssociationsSnapshot(snapshot);
#endif

    perf.SetValue0(static_cast<uint64_t>(g_settings.fileActions.viewers.associations.size()));
    perf.SetValue1(
        static_cast<uint64_t>(g_settings.fileActions.viewers.actions.size() + g_settings.fileActions.editors.actions.size() + g_settings.userMenu.actions.size()));
    perf.SetHr(S_OK);
}

LRESULT OnMainWindowSettingsFileChanged(HWND hWnd, LPARAM lParam) noexcept
{
    auto payload = TakeMessagePayload<SettingsHotReload::SettingsFileChangedPayload>(lParam);
    if (! payload || ! hWnd)
    {
        return 0;
    }

    const SettingsHotReload::ChangedSettingsLoadResult loadResult = SettingsHotReload::TryLoadChangedSettings();
    switch (loadResult.status)
    {
        case SettingsHotReload::ChangedSettingsStatus::NoChange:
        case SettingsHotReload::ChangedSettingsStatus::Missing: return 0;
        case SettingsHotReload::ChangedSettingsStatus::Invalid:
            if (loadResult.stamp.has_value())
            {
                SettingsHotReload::MarkRejectedStamp(loadResult.stamp.value());
            }
            SettingsHotReload::ShowInvalidReloadAlert(Common::Settings::GetSettingsPath(kAppId));
            return 0;
        case SettingsHotReload::ChangedSettingsStatus::Error:
            Debug::Warning(L"SettingsHotReload: TryLoadChangedSettings failed (hr=0x{:08X})", static_cast<unsigned long>(loadResult.hr));
            return 0;
        case SettingsHotReload::ChangedSettingsStatus::Loaded: break;
    }

    Common::Settings::Settings runtimeSettings = g_settings;
    CaptureRuntimeSettings(runtimeSettings, hWnd);

    g_settings = SettingsHotReload::MergeDiskSettingsWithRuntimeSession(loadResult.settings, runtimeSettings, CollectRuntimeSettingsWindowIds());
    ShortcutDefaults::EnsureShortcutsInitialized(g_settings);

    SettingsHotReload::ClearInvalidReloadAlert();
    ApplyCurrentSettingsToRunningApp(hWnd);
    RefreshRunningPluginsFromSettings(hWnd);

    if (loadResult.stamp.has_value())
    {
        SettingsHotReload::MarkAppliedStamp(loadResult.stamp.value());
    }

    SettingsHotReload::NotifyParticipants();
    return 0;
}

namespace
{
LRESULT OnMainWindowCreate(HWND hWnd, [[maybe_unused]] const CREATESTRUCTW* createStruct)
{
    Debug::Perf::Scope wmCreatePerf(L"App.Startup.MainWindow.WM_CREATE");
    wmCreatePerf.SetDetail(kMainWindowId);
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: begin");
    }
#endif

    g_menuBarVisible     = true;
    g_functionBarVisible = true;
    if (g_settings.mainMenu.has_value())
    {
        g_menuBarVisible     = g_settings.mainMenu.value().menuBarVisible;
        g_functionBarVisible = g_settings.mainMenu.value().functionBarVisible;
    }
    g_menuBarTemporarilyShown = false;

    {
        Debug::Perf::Scope perf(L"App.Startup.MainWindow.Menus");
        perf.SetDetail(kMainWindowId);
        static_cast<void>(ReloadLocalizedMainMenu(hWnd, false));
    }

    if (g_mainMenuHandle)
    {
        SetMenu(hWnd, nullptr);
        DrawMenuBar(hWnd);
    }

    const AppTheme initialTheme = ResolveConfiguredTheme();
    g_mainMenuBarHost.SetTheme(initialTheme);
    if (! g_mainMenuBarHost.EnsureCreated(hWnd))
    {
        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        const std::wstring message = L"Failed to create the DxUI main menu bar surface.";
        ShowFatalErrorDialog(hWnd, caption.c_str(), message.c_str());
        wmCreatePerf.SetHr(E_FAIL);
        return -1;
    }
    g_mainMenuBarHost.SyncMenuModel();

    {
        Debug::Perf::Scope perf(L"App.Startup.FolderWindow.Create");
        perf.SetDetail(kMainWindowId);
        const HWND folderWindow = g_folderWindow.Create(hWnd, 0, 0, 0, 0);
        g_hFolderWindow.store(folderWindow, std::memory_order_release);
        perf.SetHr(folderWindow ? S_OK : E_FAIL);
    }
    if (! g_hFolderWindow.load(std::memory_order_acquire))
    {
        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        const std::wstring message = LoadStringResource(nullptr, IDS_MSG_FAILED_CREATE_FOLDERWINDOW);
        ShowFatalErrorDialog(hWnd, caption.c_str(), message.c_str());
        wmCreatePerf.SetHr(E_FAIL);
#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: folder window create failed");
        }
#endif
        return -1;
    }
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: folder window created");
    }
#endif

    g_folderWindow.SetSettings(&g_settings);
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: folder settings pointer set");
    }
#endif
    g_folderWindow.SetShortcutManager(&g_shortcutManager);
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: shortcut manager set");
    }
#endif
    g_folderWindow.SetFunctionBarVisible(g_functionBarVisible);
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: function bar visibility set");
    }
#endif

    AdjustLayout(hWnd);
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: layout adjusted");
    }
#endif

    g_folderWindow.SetShowSortMenuCallback([hWnd](FolderWindow::Pane pane, POINT screenPoint) { ShowSortMenuPopup(hWnd, pane, screenPoint); });

    {
        Debug::Perf::Scope perf(L"App.Startup.ApplyAppTheme");
        perf.SetDetail(kMainWindowId);
        ApplyAppTheme(hWnd);
    }
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: theme applied");
    }
#endif

    const std::filesystem::path safeDefault = GetDefaultFolder().value_or(std::filesystem::path(L"C:\\"));

    const Common::Settings::FolderPane* leftSettings  = nullptr;
    const Common::Settings::FolderPane* rightSettings = nullptr;
    FolderWindow::Pane activePane                     = FolderWindow::Pane::Left;
    float splitRatio                                  = 0.5f;
    std::optional<FolderWindow::Pane> zoomedPane;
    std::optional<float> zoomRestoreSplitRatio;
    uint32_t folderHistoryMax = 20u;
    bool showHiddenFiles      = true;
    bool showSystemFiles      = true;
    std::vector<std::filesystem::path> folderHistory;

#ifdef ENABLE_TESTS
    const bool useDeterministicSelfTestPaneStartup = IsRunningAnySelfTest();
#else
    const bool useDeterministicSelfTestPaneStartup = false;
#endif

    if (! useDeterministicSelfTestPaneStartup && g_settings.folders)
    {
        const auto& folders = *g_settings.folders;
        splitRatio          = folders.layout.splitRatio;
        if (folders.active == kRightPaneSlot)
        {
            activePane = FolderWindow::Pane::Right;
        }

        if (folders.layout.zoomedPane.has_value())
        {
            if (folders.layout.zoomedPane.value() == kLeftPaneSlot)
            {
                zoomedPane = FolderWindow::Pane::Left;
            }
            else if (folders.layout.zoomedPane.value() == kRightPaneSlot)
            {
                zoomedPane = FolderWindow::Pane::Right;
            }
        }
        if (folders.layout.zoomRestoreSplitRatio.has_value())
        {
            zoomRestoreSplitRatio = folders.layout.zoomRestoreSplitRatio.value();
        }

        folderHistoryMax = folders.historyMax;
        showHiddenFiles  = folders.showHiddenFiles;
        showSystemFiles  = folders.showSystemFiles;
        folderHistory    = folders.history;

        for (const auto& item : folders.items)
        {
            if (item.slot == kLeftPaneSlot)
            {
                leftSettings = &item;
            }
            else if (item.slot == kRightPaneSlot)
            {
                rightSettings = &item;
            }
        }
    }

    g_folderWindow.SetSplitRatio(splitRatio);
    g_folderWindow.SetActivePane(activePane);
    g_folderWindow.SetZoomState(zoomedPane, zoomRestoreSplitRatio);
    g_folderWindow.SetFolderHistoryMax(folderHistoryMax);
    g_folderWindow.SetShowHiddenFiles(showHiddenFiles);
    g_folderWindow.SetShowSystemFiles(showSystemFiles);
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: base folder settings applied");
    }
#endif

    auto applyPane = [&](FolderWindow::Pane pane, const Common::Settings::FolderPane* settingsPane)
    {
        Debug::Perf::Scope perf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left" : L"App.Startup.ApplyPane.Right");

        std::filesystem::path current = safeDefault;

        if (! useDeterministicSelfTestPaneStartup && settingsPane && ! settingsPane->current.empty())
        {
            current = settingsPane->current;
        }

        FolderView::DisplayMode displayMode     = FolderView::DisplayMode::Brief;
        FolderView::SortBy sortBy               = FolderView::SortBy::Name;
        FolderView::SortDirection sortDirection = FolderView::SortDirection::Ascending;
        bool fileExtensionsVisible              = true;
        bool navigationBarVisible               = true;
        bool filterBarVisible                   = false;
        bool statusBarVisible                   = true;
        if (! useDeterministicSelfTestPaneStartup && settingsPane)
        {
            displayMode           = DisplayModeFromSettings(settingsPane->view.display);
            sortBy                = SortByFromSettings(settingsPane->view.sortBy);
            sortDirection         = SortDirectionFromSettings(settingsPane->view.sortDirection);
            fileExtensionsVisible = settingsPane->view.fileExtensionsVisible;
            if (settingsPane->view.thumbnailsVisible)
            {
                displayMode = FolderView::DisplayMode::Thumbnails;
            }
            navigationBarVisible  = settingsPane->view.navigationBarVisible;
            filterBarVisible      = settingsPane->view.filterBarVisible;
            statusBarVisible      = settingsPane->view.statusBarVisible;
        }

        perf.SetDetail(current.native());
        perf.SetValue0(static_cast<uint64_t>(displayMode));
        perf.SetValue1(static_cast<uint64_t>(sortBy));

        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetFileExtensionsVisible"
                                                                         : L"App.Startup.ApplyPane.Right.SetFileExtensionsVisible");
            callPerf.SetDetail(current.native());
            callPerf.SetValue0(fileExtensionsVisible ? 1u : 0u);
            g_folderWindow.SetFileExtensionsVisible(pane, fileExtensionsVisible);
        }
        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetNavigationBarVisible"
                                                                         : L"App.Startup.ApplyPane.Right.SetNavigationBarVisible");
            callPerf.SetDetail(current.native());
            callPerf.SetValue0(navigationBarVisible ? 1u : 0u);
            g_folderWindow.SetNavigationBarVisible(pane, navigationBarVisible);
        }
        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetFilterBarVisible"
                                                                         : L"App.Startup.ApplyPane.Right.SetFilterBarVisible");
            callPerf.SetDetail(current.native());
            callPerf.SetValue0(filterBarVisible ? 1u : 0u);
            g_folderWindow.SetFilterBarVisible(pane, filterBarVisible);
        }
        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetStatusBarVisible"
                                                                         : L"App.Startup.ApplyPane.Right.SetStatusBarVisible");
            callPerf.SetDetail(current.native());
            callPerf.SetValue0(statusBarVisible ? 1u : 0u);
            g_folderWindow.SetStatusBarVisible(pane, statusBarVisible);
        }
#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            SelfTest::AppendSelfTestTrace(pane == FolderWindow::Pane::Left ? L"OnMainWindowCreate: left status bar set"
                                                                           : L"OnMainWindowCreate: right status bar set");
        }
#endif
        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetSort" : L"App.Startup.ApplyPane.Right.SetSort");
            callPerf.SetDetail(current.native());
            callPerf.SetValue0(static_cast<uint64_t>(sortBy));
            callPerf.SetValue1(static_cast<uint64_t>(sortDirection));
            g_folderWindow.SetSort(pane, sortBy, sortDirection);
        }
#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            SelfTest::AppendSelfTestTrace(pane == FolderWindow::Pane::Left ? L"OnMainWindowCreate: left sort set" : L"OnMainWindowCreate: right sort set");
        }
#endif
        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetDisplayMode"
                                                                         : L"App.Startup.ApplyPane.Right.SetDisplayMode");
            callPerf.SetDetail(current.native());
            callPerf.SetValue0(static_cast<uint64_t>(displayMode));
            g_folderWindow.SetDisplayMode(pane, displayMode);
        }
#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            SelfTest::AppendSelfTestTrace(pane == FolderWindow::Pane::Left ? L"OnMainWindowCreate: left display mode set"
                                                                           : L"OnMainWindowCreate: right display mode set");
        }
#endif
        {
            Debug::Perf::Scope callPerf(pane == FolderWindow::Pane::Left ? L"App.Startup.ApplyPane.Left.SetFolderPath"
                                                                         : L"App.Startup.ApplyPane.Right.SetFolderPath");
            callPerf.SetDetail(current.native());
            if (! useDeterministicSelfTestPaneStartup)
            {
                g_folderWindow.SetFolderPath(pane, current);
            }
        }
#ifdef ENABLE_TESTS
        if (IsRunningAnySelfTest())
        {
            if (useDeterministicSelfTestPaneStartup)
            {
                SelfTest::AppendSelfTestTrace(pane == FolderWindow::Pane::Left ? L"OnMainWindowCreate: left folder path skipped for selftest"
                                                                               : L"OnMainWindowCreate: right folder path skipped for selftest");
            }
            else
            {
                SelfTest::AppendSelfTestTrace(pane == FolderWindow::Pane::Left ? L"OnMainWindowCreate: left folder path set"
                                                                               : L"OnMainWindowCreate: right folder path set");
            }
        }
#endif
    };

    applyPane(FolderWindow::Pane::Left, leftSettings);
    applyPane(FolderWindow::Pane::Right, rightSettings);
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: pane settings applied");
    }
#endif

    if (! useDeterministicSelfTestPaneStartup && ! folderHistory.empty())
    {
        Debug::Perf::Scope perf(L"App.Startup.FolderHistory.Set");
        perf.SetDetail(kMainWindowId);
        perf.SetValue0(folderHistory.size());
        g_folderWindow.SetFolderHistory(folderHistory);
    }
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: folder history processed");
    }
#endif

    const HRESULT hotReloadHr = useDeterministicSelfTestPaneStartup ? S_FALSE : SettingsHotReload::Start(hWnd, kAppId);
    if (FAILED(hotReloadHr))
    {
        Debug::Warning(L"SettingsHotReload::Start failed (hr=0x{:08X})", static_cast<unsigned long>(hotReloadHr));
    }
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(std::format(L"OnMainWindowCreate: hot reload start hr=0x{:08X}", static_cast<unsigned>(hotReloadHr)));
    }
#endif

#ifdef ENABLE_TESTS
    if (g_runCompareDirectoriesSelfTest)
    {
        SplashScreen::IfExistSetText(L"Launching compare-selftest...");
        SelfTest::SelfTestSuiteResult compareResult;
        Debug::Info(L"CompareSelfTest: running");
        SelfTest::InitSelfTestRun(g_selfTestOptions);
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, L"CompareSelfTest: begin");
        SelfTest::AppendSelfTestTrace(L"CompareSelfTest: begin");
        g_selfTestExitCode |= CompareDirectoriesSelfTest::Run(g_selfTestOptions, &compareResult) ? 0 : 1;
        RecordSelfTestSuite(compareResult);
        if (g_selfTestExitCode != 0)
        {
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, L"CompareSelfTest: FAIL");
            SelfTest::AppendSelfTestTrace(L"CompareSelfTest: FAIL");
            if (! compareResult.failureMessage.empty())
            {
                SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, compareResult.failureMessage);
                SelfTest::AppendSelfTestTrace(compareResult.failureMessage);
            }
        }
        else
        {
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, L"CompareSelfTest: PASS");
            SelfTest::AppendSelfTestTrace(L"CompareSelfTest: PASS");
        }
        TraceSelfTestExitCode(L"CompareSelfTest: end", g_selfTestExitCode);
    }

    const bool deferCommandsSelfTest = g_runCommandsSelfTest && ! g_runCompareDirectoriesSelfTest && ! g_runFileOpsSelfTest;

    if (g_runCommandsSelfTest && ! deferCommandsSelfTest)
    {
        static_cast<void>(RunCommandsSelfTestAndRequestShutdown(hWnd));
    }
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: command selftest dispatch processed");
    }
#endif

    if (g_runFileOpsSelfTest)
    {
        SplashScreen::IfExistSetText(L"Launching file-operations self-test...");
        Debug::Info(L"FileOpsSelfTest: scheduling");
        SelfTest::InitSelfTestRun(g_selfTestOptions);
        SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, L"FileOpsSelfTest: scheduling");
        g_fileOpsSelfTestRunFilters            = FileOperationsSelfTest::BuildRunFilters(g_selfTestOptions);
        g_fileOpsSelfTestExpectedCases         = FileOperationsSelfTest::BuildExpectedCaseNames(g_selfTestOptions);
        g_fileOpsSelfTestRunIndex              = 0;
        g_fileOpsSelfTestAggregateResult       = {};
        g_fileOpsSelfTestAggregateResult.suite = SelfTest::SelfTestSuite::FileOperations;

        if (g_fileOpsSelfTestRunFilters.empty() || g_fileOpsSelfTestExpectedCases.empty())
        {
            const std::wstring message = std::format(L"FileOpsSelfTest: unknown case/family filter '{}'.", g_selfTestOptions.caseFilter);
            Debug::Error(L"{}", message);
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, message);
            SelfTest::AppendSelfTestTrace(message);

            SelfTest::SelfTestSuiteResult failure{};
            failure.suite          = SelfTest::SelfTestSuite::FileOperations;
            failure.failed         = 1;
            failure.failureMessage = message;
            RecordSelfTestSuite(std::move(failure));

            g_selfTestExitCode |= 1;
            TraceSelfTestExitCode(L"FileOpsSelfTest: scheduling failed", g_selfTestExitCode);
            FinalizeSelfTestRun();
            PostQuitMessage(g_selfTestExitCode);
            return 0;
        }

        StartNextFileOpsSelfTestRun(hWnd);
        if (SetTimer(hWnd, kFileOpsSelfTestTimerId, kFileOpsSelfTestTimerIntervalMs, nullptr) == 0)
        {
            const HRESULT hr     = HRESULT_FROM_WIN32(GetLastError());
            std::wstring message = std::format(L"FileOpsSelfTest: SetTimer failed: 0x{:08X}", static_cast<unsigned>(hr));
            Debug::Error(L"{}", message);
            SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::FileOperations, message);
            SelfTest::AppendSelfTestTrace(message);

            SelfTest::SelfTestSuiteResult failure{};
            failure.suite          = SelfTest::SelfTestSuite::FileOperations;
            failure.failed         = 1;
            failure.failureMessage = std::move(message);
            RecordSelfTestSuite(std::move(failure));

            g_selfTestExitCode |= 1;
            TraceSelfTestExitCode(L"FileOpsSelfTest: scheduling failed", g_selfTestExitCode);
            FinalizeSelfTestRun();
            PostQuitMessage(g_selfTestExitCode);
        }
    }

    if (deferCommandsSelfTest)
    {
        SplashScreen::IfExistSetText(L"Scheduling commands-selftest...");
        if (PostMessageW(hWnd, kCommandsSelfTestStartMessage, 0, 0) == 0)
        {
            FinalizeSelfTestRun();
            PostQuitMessage(g_selfTestExitCode);
        }
    }
    else if ((g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest) && ! g_runFileOpsSelfTest)
    {
        // Selftests may exit before the app reaches its normal "input ready" phase.
        // Close the splash and route shutdown through WM_CLOSE/WM_DESTROY so graphics resources are torn down
        // before CRT/static destruction (avoids late D2D/D3D/icon-cache access during process shutdown).
        SplashScreen::CloseIfExist();
        if (PostMessageW(hWnd, WM_CLOSE, 0, 0) == 0)
        {
            PostQuitMessage(g_selfTestExitCode);
        }
    }
#endif

    wmCreatePerf.SetHr(S_OK);
#ifdef ENABLE_TESTS
    if (IsRunningAnySelfTest())
    {
        SelfTest::AppendSelfTestTrace(L"OnMainWindowCreate: success");
    }
#endif
    return 0;
}

void EnterFullScreen(HWND hWnd) noexcept
{
    if (! hWnd || IsWindow(hWnd) == FALSE || g_fullScreenState.active)
    {
        return;
    }

    g_fullScreenState.savedStyle   = static_cast<DWORD>(GetWindowLongPtrW(hWnd, GWL_STYLE));
    g_fullScreenState.savedExStyle = static_cast<DWORD>(GetWindowLongPtrW(hWnd, GWL_EXSTYLE));

    g_fullScreenState.savedPlacement        = {};
    g_fullScreenState.savedPlacement.length = sizeof(WINDOWPLACEMENT);
    if (GetWindowPlacement(hWnd, &g_fullScreenState.savedPlacement) == FALSE)
    {
        // Fallback: synthesize a placement from the current window rectangle.
        RECT windowRect{};
        if (GetWindowRect(hWnd, &windowRect) != FALSE)
        {
            g_fullScreenState.savedPlacement.length           = sizeof(WINDOWPLACEMENT);
            g_fullScreenState.savedPlacement.flags            = 0;
            g_fullScreenState.savedPlacement.showCmd          = SW_SHOWNORMAL;
            g_fullScreenState.savedPlacement.ptMinPosition.x  = windowRect.left;
            g_fullScreenState.savedPlacement.ptMinPosition.y  = windowRect.top;
            g_fullScreenState.savedPlacement.ptMaxPosition.x  = windowRect.left;
            g_fullScreenState.savedPlacement.ptMaxPosition.y  = windowRect.top;
            g_fullScreenState.savedPlacement.rcNormalPosition = windowRect;
        }
        else
        {
            // Indicate that there is no valid placement information.
            g_fullScreenState.savedPlacement.length = 0;
        }
    }

    MONITORINFO mi{};
    mi.cbSize              = sizeof(mi);
    const HMONITOR monitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST);
    if (! monitor || GetMonitorInfoW(monitor, &mi) == FALSE)
    {
        return;
    }

    DWORD newStyle = g_fullScreenState.savedStyle;
    newStyle &= ~WS_OVERLAPPEDWINDOW;
    newStyle |= WS_POPUP;

    DWORD newExStyle = g_fullScreenState.savedExStyle;
    newExStyle |= WS_EX_TOPMOST;

    SetWindowLongPtrW(hWnd, GWL_STYLE, static_cast<LONG_PTR>(newStyle));
    SetWindowLongPtrW(hWnd, GWL_EXSTYLE, static_cast<LONG_PTR>(newExStyle));

    const int x      = mi.rcMonitor.left;
    const int y      = mi.rcMonitor.top;
    const int width  = mi.rcMonitor.right - mi.rcMonitor.left;
    const int height = mi.rcMonitor.bottom - mi.rcMonitor.top;

    SetWindowPos(hWnd, HWND_TOPMOST, x, y, width, height, SWP_FRAMECHANGED | SWP_NOOWNERZORDER | SWP_NOACTIVATE);
    g_fullScreenState.active = true;
}

void ExitFullScreen(HWND hWnd) noexcept
{
    if (! hWnd || IsWindow(hWnd) == FALSE || ! g_fullScreenState.active)
    {
        return;
    }

    SetWindowLongPtrW(hWnd, GWL_STYLE, static_cast<LONG_PTR>(g_fullScreenState.savedStyle));
    SetWindowLongPtrW(hWnd, GWL_EXSTYLE, static_cast<LONG_PTR>(g_fullScreenState.savedExStyle));

    WINDOWPLACEMENT placement = g_fullScreenState.savedPlacement;
    if ((g_fullScreenState.savedStyle & WS_VISIBLE) == 0)
    {
        placement.showCmd = SW_HIDE;
    }

    if (placement.length == sizeof(WINDOWPLACEMENT))
    {
        static_cast<void>(SetWindowPlacement(hWnd, &placement));
    }
    else if ((g_fullScreenState.savedStyle & WS_VISIBLE) == 0)
    {
        ShowWindow(hWnd, SW_HIDE);
    }

    const HWND insertAfter = (g_fullScreenState.savedExStyle & WS_EX_TOPMOST) != 0 ? HWND_TOPMOST : HWND_NOTOPMOST;
    SetWindowPos(hWnd, insertAfter, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED | SWP_NOOWNERZORDER | SWP_NOACTIVATE);

    g_fullScreenState.active = false;
}

void ToggleFullScreen(HWND hWnd) noexcept
{
    if (g_fullScreenState.active)
    {
        ExitFullScreen(hWnd);
    }
    else
    {
        EnterFullScreen(hWnd);
    }
}

void OpenExternalHelp(HWND ownerWindow) noexcept
{
    const HINSTANCE result = ShellExecuteW(ownerWindow, L"open", kExternalHelpUrl, nullptr, nullptr, SW_SHOWNORMAL);
    const auto code        = reinterpret_cast<INT_PTR>(result);
    if (code <= 32)
    {
        Debug::Error(L"External Help: ShellExecuteW failed for '{}' (code={}).", kExternalHelpUrl, static_cast<long long>(code));
    }
}

LRESULT OnMainWindowCommand(HWND hWnd, UINT id, UINT codeNotify, HWND hwndCtl)
{
    const UINT wmId = id;
    switch (wmId)
    {
        case IDM_ABOUT:
        {
            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowAboutDialog(hWnd, theme));
            break;
        }
        case IDM_APP_EXTERNAL_HELP:
        {
            OpenExternalHelp(hWnd);
            break;
        }
        case IDM_FILE_PREFERENCES:
        {
            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowPreferencesDialog(hWnd, kAppId, g_settings, theme));
            break;
        }
        case IDM_APP_SHOW_SHORTCUTS:
        {
            const AppTheme theme = ResolveConfiguredTheme();
            if (g_settings.shortcuts.has_value())
            {
                ShowShortcutsWindow(hWnd, g_settings, g_settings.shortcuts.value(), g_shortcutManager, theme);
            }
            break;
        }
        case IDM_APP_FULL_SCREEN:
        {
            ToggleFullScreen(hWnd);
            break;
        }
        case IDM_APP_VIEW_WIDTH:
        {
            if (g_hFolderWindow.load(std::memory_order_acquire))
            {
                if (g_folderWindow.IsViewWidthAdjustActive())
                {
                    g_folderWindow.CommitViewWidthAdjust();
                }
                else
                {
                    g_folderWindow.BeginViewWidthAdjust();
                }
            }
            break;
        }
        case IDM_APP_REREAD_ASSOCIATIONS:
        {
            RereadAssociations(hWnd);
            break;
        }
        case IDM_EXIT: SendMessageW(hWnd, WM_CLOSE, 0, 0); break;
        case IDM_VIEW_MENUBAR:
        {
            EnsureMenuHandles(hWnd);

            g_menuBarVisible          = ! g_menuBarVisible;
            g_menuBarTemporarilyShown = false;

            UpdatePaneMenuChecks();
            g_mainMenuBarHost.SyncMenuModel();
            AdjustLayout(hWnd);
            break;
        }
        case IDM_VIEW_FUNCTIONBAR:
        {
            g_functionBarVisible = ! g_functionBarVisible;
            g_folderWindow.SetFunctionBarVisible(g_functionBarVisible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_VIEW_WINDOW_MENU: SendMessageW(hWnd, WM_SYSCOMMAND, SC_KEYMENU, static_cast<LPARAM>(' ')); break;
        case IDM_VIEW_SWITCH_PANE_FOCUS:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendKeyToFolderView(pane, VK_TAB));
            break;
        }
        case IDM_VIEW_FILEOPS_FAILED_ITEMS:
            g_folderWindow.CommandToggleFileOperationsIssuesPane();
            UpdatePaneMenuChecks();
            break;
        case IDM_APP_COMPARE:
        {
            const auto showErrorAlert = [&](unsigned int messageStringId) noexcept
            {
                const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
                const std::wstring message = LoadStringResource(nullptr, messageStringId);

                HostAlertRequest request{};
                request.version      = 1;
                request.sizeBytes    = sizeof(request);
                request.scope        = HOST_ALERT_SCOPE_WINDOW;
                request.modality     = HOST_ALERT_MODELESS;
                request.severity     = HOST_ALERT_ERROR;
                request.targetWindow = hWnd;
                request.title        = title.c_str();
                request.message      = message.c_str();
                request.closable     = TRUE;
                static_cast<void>(HostShowAlert(request));
            };
            const auto showInvalidPathAlert = [&](std::wstring_view pathText) noexcept
            {
                const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_INVALID_PATH);
                const std::wstring message = FormatStringResource(nullptr, IDS_FMT_INVALID_PATH, pathText);

                HostAlertRequest request{};
                request.version      = 1;
                request.sizeBytes    = sizeof(request);
                request.scope        = HOST_ALERT_SCOPE_WINDOW;
                request.modality     = HOST_ALERT_MODELESS;
                request.severity     = HOST_ALERT_ERROR;
                request.targetWindow = hWnd;
                request.title        = title.c_str();
                request.message      = message.c_str();
                request.closable     = TRUE;
                static_cast<void>(HostShowAlert(request));
            };

            const std::wstring_view leftPluginId  = g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Left);
            const std::wstring_view rightPluginId = g_folderWindow.GetFileSystemPluginId(FolderWindow::Pane::Right);
            if (leftPluginId.empty() || rightPluginId.empty())
            {
                showErrorAlert(IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS);
                break;
            }

            const std::optional<std::filesystem::path> leftRoot  = g_folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Left);
            const std::optional<std::filesystem::path> rightRoot = g_folderWindow.GetCurrentPluginPath(FolderWindow::Pane::Right);
            if (! leftRoot.has_value() || ! rightRoot.has_value() || leftRoot->empty() || rightRoot->empty())
            {
                std::wstring_view badPath;
                if (leftRoot.has_value() && ! leftRoot->empty())
                {
                    if (rightRoot.has_value())
                    {
                        badPath = rightRoot->native();
                    }
                }
                else if (leftRoot.has_value())
                {
                    badPath = leftRoot->native();
                }

                showInvalidPathAlert(badPath);
                break;
            }

            const std::optional<std::filesystem::path> leftCurrent  = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left);
            const std::optional<std::filesystem::path> rightCurrent = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Right);

            NavigationLocation::Location leftLocation{};
            NavigationLocation::Location rightLocation{};
            if (! leftCurrent.has_value() || ! rightCurrent.has_value() || ! NavigationLocation::TryParseLocation(leftCurrent->native(), leftLocation) ||
                ! NavigationLocation::TryParseLocation(rightCurrent->native(), rightLocation))
            {
                showInvalidPathAlert({});
                break;
            }

            const auto findPluginById = [](std::wstring_view id) noexcept -> const FileSystemPluginManager::PluginEntry*
            {
                for (const FileSystemPluginManager::PluginEntry& entry : FileSystemPluginManager::GetInstance().GetPlugins())
                {
                    if (CompareStringOrdinal(entry.id.c_str(), -1, id.data(), static_cast<int>(id.size()), TRUE) == CSTR_EQUAL)
                    {
                        return &entry;
                    }
                }

                return nullptr;
            };

            const FileSystemPluginManager::PluginEntry* leftPlugin  = findPluginById(leftPluginId);
            const FileSystemPluginManager::PluginEntry* rightPlugin = findPluginById(rightPluginId);
            if (! leftPlugin || ! rightPlugin || leftPlugin->disabled || rightPlugin->disabled || ! leftPlugin->loadable || ! rightPlugin->loadable)
            {
                showErrorAlert(IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS);
                break;
            }

            CompareDirectoriesPaneContext left{};
            left.pluginId        = std::wstring(leftPluginId);
            left.instanceContext = std::move(leftLocation.instanceContext);
            left.rootPluginPath  = leftRoot.value();

            CompareDirectoriesPaneContext right{};
            right.pluginId        = std::wstring(rightPluginId);
            right.instanceContext = std::move(rightLocation.instanceContext);
            right.rootPluginPath  = rightRoot.value();

            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowCompareDirectoriesWindow(hWnd, g_settings, theme, &g_shortcutManager, std::move(left), std::move(right)));
            break;
        }
        case IDM_PANE_FIND:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();

            FindFilesPaneContext context{};
            context.fileSystem      = g_folderWindow.GetFileSystem(pane);
            context.pluginId        = std::wstring(g_folderWindow.GetFileSystemPluginId(pane));
            context.pluginShortId   = std::wstring(g_folderWindow.GetFileSystemPluginShortId(pane));
            context.instanceContext = std::wstring(g_folderWindow.GetFileSystemInstanceContext(pane));
            context.rootPluginPath  = g_folderWindow.GetCurrentPluginPath(pane).value_or(std::filesystem::path{});

            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowFindFilesWindow(hWnd, g_settings, theme, std::move(context)));
            break;
        }
        case IDM_APP_SWAP_PANES: g_folderWindow.SwapPanes(); break;
        case IDM_VIEW_THEME_PREV:
            SelectAdjacentTheme(hWnd, -1);
            break;
        case IDM_VIEW_THEME_NEXT:
            SelectAdjacentTheme(hWnd, 1);
            break;
        case IDM_VIEW_THEME_SYSTEM:
            ApplyThemeId(hWnd, L"builtin/system");
            break;
        case IDM_VIEW_THEME_LIGHT:
            ApplyThemeId(hWnd, L"builtin/light");
            break;
        case IDM_VIEW_THEME_DARK:
            ApplyThemeId(hWnd, L"builtin/dark");
            break;
        case IDM_VIEW_THEME_RAINBOW:
            ApplyThemeId(hWnd, L"builtin/rainbow");
            break;
        case IDM_VIEW_THEME_HIGH_CONTRAST_APP:
            ApplyThemeId(hWnd, L"builtin/highContrast");
            break;
        case IDM_VIEW_PLUGINS_MANAGE:
        {
            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowPreferencesDialogPlugins(hWnd, kAppId, g_settings, theme));
            break;
        }
        case IDM_VIEW_PANE_STATUSBAR_LEFT:
        case IDM_LEFT_STATUSBAR:
        {
            const bool visible = g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Left);
            g_folderWindow.SetStatusBarVisible(FolderWindow::Pane::Left, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_VIEW_PANE_STATUSBAR_RIGHT:
        case IDM_RIGHT_STATUSBAR:
        {
            const bool visible = g_folderWindow.GetStatusBarVisible(FolderWindow::Pane::Right);
            g_folderWindow.SetStatusBarVisible(FolderWindow::Pane::Right, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_VIEW_PANE_HIDDEN_FILES:
        case IDM_LEFT_SHOW_HIDDEN_FILES:
        case IDM_RIGHT_SHOW_HIDDEN_FILES:
        {
            const bool visible = g_folderWindow.GetShowHiddenFiles();
            g_folderWindow.SetShowHiddenFiles(! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_VIEW_PANE_SYSTEM_FILES:
        case IDM_LEFT_SHOW_SYSTEM_FILES:
        case IDM_RIGHT_SHOW_SYSTEM_FILES:
        {
            const bool visible = g_folderWindow.GetShowSystemFiles();
            g_folderWindow.SetShowSystemFiles(! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_VIEW_PANE_FILE_EXTENSIONS:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
            const bool visible           = g_folderWindow.GetFileExtensionsVisible(pane);
            g_folderWindow.SetFileExtensionsVisible(pane, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_LEFT_SHOW_FILE_EXTENSIONS:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            const bool visible = g_folderWindow.GetFileExtensionsVisible(FolderWindow::Pane::Left);
            g_folderWindow.SetFileExtensionsVisible(FolderWindow::Pane::Left, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_RIGHT_SHOW_FILE_EXTENSIONS:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            const bool visible = g_folderWindow.GetFileExtensionsVisible(FolderWindow::Pane::Right);
            g_folderWindow.SetFileExtensionsVisible(FolderWindow::Pane::Right, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_VIEW_PANE_THUMBNAILS:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
            g_folderWindow.SetDisplayMode(pane, FolderView::DisplayMode::Thumbnails);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_LEFT_DISPLAY_THUMBNAILS:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Thumbnails);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_RIGHT_DISPLAY_THUMBNAILS:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Right, FolderView::DisplayMode::Thumbnails);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_VIEW_PANE_PREVIEW_PANE:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
            g_folderWindow.TogglePreviewPane(pane);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_LEFT_PREVIEW_PANE:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.TogglePreviewPane(FolderWindow::Pane::Left);
            UpdatePaneMenuChecks();
            break;
        case IDM_RIGHT_PREVIEW_PANE:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.TogglePreviewPane(FolderWindow::Pane::Right);
            UpdatePaneMenuChecks();
            break;
        case IDM_VIEW_PANE_FILTER_BAR:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
            const bool visible           = g_folderWindow.GetFilterBarVisible(pane);
            g_folderWindow.SetFilterBarVisible(pane, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_LEFT_FILTER_BAR:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            const bool visible = g_folderWindow.GetFilterBarVisible(FolderWindow::Pane::Left);
            g_folderWindow.SetFilterBarVisible(FolderWindow::Pane::Left, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_RIGHT_FILTER_BAR:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            const bool visible = g_folderWindow.GetFilterBarVisible(FolderWindow::Pane::Right);
            g_folderWindow.SetFilterBarVisible(FolderWindow::Pane::Right, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_LEFT_CHANGE_DRIVE: g_folderWindow.CommandOpenDriveMenu(FolderWindow::Pane::Left); break;
        case IDM_RIGHT_CHANGE_DRIVE: g_folderWindow.CommandOpenDriveMenu(FolderWindow::Pane::Right); break;
        case IDM_LEFT_GO_TO_BACK: g_folderWindow.CommandHistoryBack(FolderWindow::Pane::Left); break;
        case IDM_LEFT_GO_TO_FORWARD: g_folderWindow.CommandHistoryForward(FolderWindow::Pane::Left); break;
        case IDM_LEFT_GO_TO_PARENT_DIRECTORY:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            static_cast<void>(SendKeyToFolderView(FolderWindow::Pane::Left, VK_BACK));
            break;
        case IDM_LEFT_GO_TO_ROOT_DIRECTORY: g_folderWindow.CommandGoRootDirectory(FolderWindow::Pane::Left); break;
        case IDM_LEFT_GO_TO_PATH_FROM_OTHER_PANE: g_folderWindow.CommandSetPathFromOtherPane(FolderWindow::Pane::Left); break;
        case IDM_LEFT_HOT_PATHS:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            const AppTheme theme = ResolveConfiguredTheme();
#ifdef ENABLE_TESTS
            if (const auto beforePath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left); beforePath.has_value())
            {
                SelfTest::AppendSelfTestTrace(std::format(L"IDM_LEFT_HOT_PATHS before='{}'", beforePath->wstring()));
            }
#endif
            static_cast<void>(ShowPreferencesDialogHotPaths(hWnd, kAppId, g_settings, theme));
            g_folderWindow.ResyncNavigationShellFromFolderView(FolderWindow::Pane::Left);
#ifdef ENABLE_TESTS
            if (const auto afterPath = g_folderWindow.GetCurrentPath(FolderWindow::Pane::Left); afterPath.has_value())
            {
                SelfTest::AppendSelfTestTrace(std::format(L"IDM_LEFT_HOT_PATHS after='{}'", afterPath->wstring()));
            }
#endif
            break;
        }
        case IDM_LEFT_ZOOM_PANEL: g_folderWindow.ToggleZoomPanel(FolderWindow::Pane::Left); break;
        case IDM_LEFT_FILTER: g_folderWindow.CommandFilter(FolderWindow::Pane::Left); break;
        case IDM_LEFT_REFRESH: g_folderWindow.CommandRefresh(FolderWindow::Pane::Left); break;
        case IDM_RIGHT_GO_TO_BACK: g_folderWindow.CommandHistoryBack(FolderWindow::Pane::Right); break;
        case IDM_RIGHT_GO_TO_FORWARD: g_folderWindow.CommandHistoryForward(FolderWindow::Pane::Right); break;
        case IDM_RIGHT_GO_TO_PARENT_DIRECTORY:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            static_cast<void>(SendKeyToFolderView(FolderWindow::Pane::Right, VK_BACK));
            break;
        case IDM_RIGHT_GO_TO_ROOT_DIRECTORY: g_folderWindow.CommandGoRootDirectory(FolderWindow::Pane::Right); break;
        case IDM_RIGHT_GO_TO_PATH_FROM_OTHER_PANE: g_folderWindow.CommandSetPathFromOtherPane(FolderWindow::Pane::Right); break;
        case IDM_RIGHT_HOT_PATHS:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            const AppTheme theme = ResolveConfiguredTheme();
            static_cast<void>(ShowPreferencesDialogHotPaths(hWnd, kAppId, g_settings, theme));
            g_folderWindow.ResyncNavigationShellFromFolderView(FolderWindow::Pane::Right);
            break;
        }
        case IDM_RIGHT_ZOOM_PANEL: g_folderWindow.ToggleZoomPanel(FolderWindow::Pane::Right); break;
        case IDM_RIGHT_FILTER: g_folderWindow.CommandFilter(FolderWindow::Pane::Right); break;
        case IDM_RIGHT_REFRESH: g_folderWindow.CommandRefresh(FolderWindow::Pane::Right); break;
        case IDM_VIEW_PANE_NAVBAR_LEFT:
        case IDM_LEFT_NAVIGATION_BAR:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            const bool visible = g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Left);
            g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Left, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_VIEW_PANE_NAVBAR_RIGHT:
        case IDM_RIGHT_NAVIGATION_BAR:
        {
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            const bool visible = g_folderWindow.GetNavigationBarVisible(FolderWindow::Pane::Right);
            g_folderWindow.SetNavigationBarVisible(FolderWindow::Pane::Right, ! visible);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_PANE_MENU: SendMessageW(hWnd, WM_SYSCOMMAND, SC_KEYMENU, 0); break;
        case IDM_PANE_USER_MENU: ShowUserMenuPopup(hWnd); break;
        case IDM_PANE_EXECUTE_OPEN:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_OPEN));
            break;
        }
        case IDM_PANE_MOVE_TO_RECYCLE_BIN:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_DELETE));
            break;
        }
        case IDM_PANE_CLIPBOARD_CUT:
        {
            if (g_folderWindow.TryHandleNavigationEditClipboardCommand(id))
            {
                break;
            }
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            if (HWND folderView = g_folderWindow.GetFolderViewHwnd(pane))
            {
                auto* view = reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderView, GWLP_USERDATA));
                if (view)
                {
                    static_cast<void>(view->CutSelectionToClipboard());
                }
            }
            break;
        }
        case IDM_PANE_CLIPBOARD_COPY:
        {
            if (g_folderWindow.TryHandleNavigationEditClipboardCommand(id))
            {
                break;
            }
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_COPY));
            break;
        }
        case IDM_PANE_COPY_PATH_AND_FILE_NAME:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandCopyUncPathAndNameAsText(pane);
            break;
        }
        case IDM_PANE_COPY_PATH_AND_NAME_AS_TEXT:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandCopyPathAndNameAsText(pane);
            break;
        }
        case IDM_PANE_COPY_NAME_AS_TEXT:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandCopyNameAsText(pane);
            break;
        }
        case IDM_PANE_COPY_PATH_AS_TEXT:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandCopyPathAsText(pane);
            break;
        }
        case IDM_PANE_SELECTION_SELECT_ALL:
        {
            if (g_folderWindow.TryHandleNavigationEditClipboardCommand(id))
            {
                break;
            }
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_SELECT_ALL));
            break;
        }
        case IDM_PANE_SELECTION_SELECT_DIALOG:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionSelectDialog(pane);
            break;
        }
        case IDM_PANE_SELECTION_UNSELECT_ALL:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_UNSELECT_ALL));
            break;
        }
        case IDM_PANE_SELECTION_UNSELECT_DIALOG:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionUnselectDialog(pane);
            break;
        }
        case IDM_PANE_SAVE_SELECTION:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionSave(pane);
            break;
        }
        case IDM_PANE_SELECTION_RESTORE:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionRestore(pane);
            break;
        }
        case IDM_PANE_SELECTION_SELECT_SAME_NAME:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionSelectSameName(pane);
            break;
        }
        case IDM_PANE_SELECTION_UNSELECT_SAME_NAME:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionUnselectSameName(pane);
            break;
        }
        case IDM_PANE_SELECTION_INVERT:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionInvert(pane);
            break;
        }
        case IDM_PANE_SELECTION_SELECT_SAME_EXTENSION:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionSelectSameExtension(pane);
            break;
        }
        case IDM_PANE_SELECTION_UNSELECT_SAME_EXTENSION:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionUnselectSameExtension(pane);
            break;
        }
        case IDM_PANE_SELECTION_HIDE_SELECTED_NAMES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionHideSelectedNames(pane);
            break;
        }
        case IDM_PANE_SELECTION_HIDE_UNSELECTED_NAMES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionHideUnselectedNames(pane);
            break;
        }
        case IDM_PANE_SELECTION_SHOW_HIDDEN_NAMES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionShowHiddenNames(pane);
            break;
        }
        case IDM_PANE_SELECTION_GOTO_PREV_SELECTED_NAME:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionGoToPreviousSelectedName(pane);
            break;
        }
        case IDM_PANE_SELECTION_GOTO_NEXT_SELECTED_NAME:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSelectionGoToNextSelectedName(pane);
            break;
        }
        case IDM_PANE_CLIPBOARD_PASTE:
        {
            if (g_folderWindow.TryHandleNavigationEditClipboardCommand(id))
            {
                break;
            }
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_PASTE));
            break;
        }
        case IDM_PANE_CLIPBOARD_PASTE_SHORTCUT:
        {
#ifdef ENABLE_TESTS
            SelfTest::AppendSelfTestTrace(std::format(L"WM_COMMAND PasteShortcut enter focus=0x{:X} focusedPane={} focusedView=0x{:X}",
                                                      reinterpret_cast<uintptr_t>(GetFocus()),
                                                      static_cast<int>(g_folderWindow.GetFocusedPane()),
                                                      reinterpret_cast<uintptr_t>(g_folderWindow.GetFocusedFolderViewHwnd())));
#endif
            if (g_folderWindow.TryHandleNavigationEditClipboardCommand(id))
            {
#ifdef ENABLE_TESTS
                SelfTest::AppendSelfTestTrace(L"WM_COMMAND PasteShortcut handled by navigation edit");
#endif
                break;
            }
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            if (HWND folderView = g_folderWindow.GetFolderViewHwnd(pane))
            {
                auto* view = reinterpret_cast<FolderView*>(GetWindowLongPtrW(folderView, GWLP_USERDATA));
#ifdef ENABLE_TESTS
                SelfTest::AppendSelfTestTrace(std::format(L"WM_COMMAND PasteShortcut folderView=0x{:X} view={} pane={}",
                                                          reinterpret_cast<uintptr_t>(folderView),
                                                          view ? 1 : 0,
                                                          static_cast<int>(pane)));
#endif
                if (view)
                {
                    static_cast<void>(view->PasteShortcutFromClipboard());
                }
            }
            break;
        }
        case IDM_PANE_OPEN_PROPERTIES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendCommandToFolderView(pane, IDM_FOLDERVIEW_CONTEXT_PROPERTIES));
            break;
        }
        case IDM_PANE_OPEN_SECURITY:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandOpenSecurity(pane);
            break;
        }
        case IDM_PANE_GO_TO_SHORTCUT_OR_LINK_TARGET:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandGoToShortcutOrLinkTarget(pane);
            break;
        }
        case IDM_PANE_CHANGE_ATTRIBUTES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandChangeAttributes(pane);
            break;
        }
        case IDM_PANE_CONTEXT_MENU:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendKeyToFolderView(pane, VK_APPS));
            break;
        }
        case IDM_PANE_CONTEXT_MENU_CURRENT_DIRECTORY:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandContextMenuCurrentDirectory(pane);
            break;
        }
        case IDM_PANE_SELECT_NEXT:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendKeyToFolderView(pane, VK_INSERT));
            break;
        }
        case IDM_PANE_SELECT_CALC_DIR_SIZE_NEXT:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            static_cast<void>(SendKeyToFolderView(pane, VK_SPACE));
            break;
        }
        case IDM_PANE_CHANGE_DIRECTORY:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandChangeDirectory(pane);
            break;
        }
        case IDM_PANE_SHOW_FOLDERS_HISTORY:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandShowFolderHistory(pane);
            break;
        }
        case IDM_PANE_CONNECT:
        {
            std::optional<std::wstring> remoteName;

            const FolderWindow::Pane pane        = g_folderWindow.GetFocusedPane();
            const std::wstring_view fileSystemId = g_folderWindow.GetFileSystemPluginId(pane);
            if (fileSystemId == L"builtin/file-system")
            {
                const std::optional<std::filesystem::path> pathOpt = g_folderWindow.GetCurrentPluginPath(pane);
                if (pathOpt.has_value())
                {
                    const std::wstring pathText = pathOpt.value().wstring();
                    if (LooksLikeUncPath(pathText))
                    {
                        remoteName = pathText;
                    }
                }
            }

            const DWORD drivesBefore = GetLogicalDrives();
            const HRESULT hr         = ShowConnectNetworkDriveDialog(hWnd, remoteName);
            if (FAILED(hr) || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                break;
            }

            if (drivesBefore == 0)
            {
                break;
            }

            DWORD drivesAfter  = 0;
            DWORD newDriveMask = 0;
            for (int attempt = 0; attempt < 10 && newDriveMask == 0; ++attempt)
            {
                drivesAfter = GetLogicalDrives();
                if (drivesAfter != 0)
                {
                    newDriveMask = drivesAfter & ~drivesBefore;
                }

                if (newDriveMask == 0)
                {
                    Sleep(static_cast<DWORD>(50u));
                }
            }

            if (newDriveMask == 0)
            {
                break;
            }

            std::optional<std::filesystem::path> newDrivePath;
            std::optional<std::filesystem::path> firstNewDrivePath;

            for (unsigned int bitIndex = 0u; bitIndex < 26u; ++bitIndex)
            {
                const DWORD bit = static_cast<DWORD>(1u) << bitIndex;
                if ((newDriveMask & bit) == 0)
                {
                    continue;
                }

                const wchar_t driveLetter = static_cast<wchar_t>(L'A' + bitIndex);
                std::wstring driveRoot;
                driveRoot.push_back(driveLetter);
                driveRoot.append(L":\\");

                const UINT driveType = GetDriveTypeW(driveRoot.c_str());
                if (! firstNewDrivePath.has_value())
                {
                    firstNewDrivePath = std::filesystem::path(driveRoot);
                }
                if (driveType == DRIVE_REMOTE)
                {
                    newDrivePath = std::filesystem::path(driveRoot);
                    break;
                }
            }

            if (! newDrivePath.has_value() && firstNewDrivePath.has_value())
            {
                newDrivePath = firstNewDrivePath.value();
            }

            if (newDrivePath.has_value())
            {
                g_folderWindow.SetActivePane(pane);
                g_folderWindow.SetFolderPath(pane, newDrivePath.value());
            }
            break;
        }
        case IDM_PANE_DISCONNECT:
        {
            const DWORD drivesBefore = GetLogicalDrives();

            std::optional<unsigned int> focusedDriveBitIndex;
            std::optional<std::wstring> localName;
            std::optional<std::wstring> remoteName;

            const FolderWindow::Pane pane        = g_folderWindow.GetFocusedPane();
            const std::wstring_view fileSystemId = g_folderWindow.GetFileSystemPluginId(pane);
            if (fileSystemId == L"builtin/file-system")
            {
                const std::optional<std::filesystem::path> pathOpt = g_folderWindow.GetCurrentPluginPath(pane);
                if (pathOpt.has_value())
                {
                    std::wstring pathText = pathOpt.value().wstring();
                    std::replace(pathText.begin(), pathText.end(), L'/', L'\\');

                    if (pathText.size() >= 2 && std::iswalpha(static_cast<wint_t>(pathText[0])) != 0 && pathText[1] == L':')
                    {
                        const wchar_t driveLetter = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(pathText[0])));
                        if (driveLetter >= L'A' && driveLetter <= L'Z')
                        {
                            focusedDriveBitIndex = static_cast<unsigned int>(driveLetter - L'A');
                        }

                        std::wstring driveRoot;
                        driveRoot.push_back(driveLetter);
                        driveRoot.append(L":\\");

                        const UINT driveType = GetDriveTypeW(driveRoot.c_str());
                        if (driveType == DRIVE_REMOTE)
                        {
                            std::wstring local;
                            local.push_back(driveLetter);
                            local.push_back(L':');
                            localName = std::move(local);
                        }
                    }
                    else if (LooksLikeUncPath(pathText))
                    {
                        remoteName = TryGetUncShareRoot(pathText);
                    }
                }
            }

            g_folderWindow.PrepareForNetworkDriveDisconnect(pane);

            const HRESULT hr = ShowDisconnectNetworkDriveDialog(hWnd, localName, remoteName);
            if (FAILED(hr) || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                break;
            }

            if (drivesBefore == 0 || ! focusedDriveBitIndex.has_value())
            {
                break;
            }

            DWORD drivesAfter      = 0;
            DWORD removedDriveMask = 0;
            for (int attempt = 0; attempt < 10 && removedDriveMask == 0; ++attempt)
            {
                drivesAfter = GetLogicalDrives();
                if (drivesAfter != 0)
                {
                    removedDriveMask = drivesBefore & ~drivesAfter;
                }

                if (removedDriveMask == 0)
                {
                    Sleep(static_cast<DWORD>(50u));
                }
            }

            if (removedDriveMask == 0)
            {
                break;
            }

            const unsigned int bitIndex = focusedDriveBitIndex.value();
            if (bitIndex >= 26u)
            {
                break;
            }

            const DWORD focusedDriveBit = static_cast<DWORD>(1u) << bitIndex;
            if ((removedDriveMask & focusedDriveBit) == 0)
            {
                break;
            }

            g_folderWindow.SetActivePane(pane);
            g_folderWindow.SetFolderPath(pane, GetDefaultFileSystemRoot());
            break;
        }
        case IDM_PANE_CONNECTION_MANAGER:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetActivePane();
            const AppTheme theme          = ResolveConfiguredTheme();

            static_cast<void>(ShowConnectionManagerWindow(hWnd, kAppId, g_settings, theme, {}, static_cast<uint8_t>(pane)));
            break;
        }
        case IDM_PANE_MAKE_FILE_LIST:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandMakeFileList(pane);
            break;
        }
        case IDM_PANE_PACK:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandPack(pane);
            break;
        }
        case IDM_PANE_UNPACK:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandUnpack(pane);
            break;
        }
        case IDM_PANE_LIST_OPENED_FILES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandListOpenedFiles(pane);
            break;
        }
        case IDM_PANE_SHARES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandSharedDirectories(pane);
            break;
        }
        case IDM_PANE_CHANGE_CASE:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandChangeCase(pane);
            break;
        }
        case IDM_PANE_OPEN_COMMAND_SHELL:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandOpenCommandShell(pane);
            break;
        }
        case IDM_PANE_QUICK_SEARCH:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandQuickSearch(pane);
            break;
        }
        case IDM_PANE_BRING_CURRENT_DIR_TO_COMMAND_LINE:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandBringCurrentDirToCommandLine(pane);
            break;
        }
        case IDM_PANE_BRING_FILENAME_TO_COMMAND_LINE:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.CommandBringFilenameToCommandLine(pane);
            break;
        }
        case IDM_PANE_OPEN_CURRENT_FOLDER:
        {
            const FolderWindow::Pane pane                   = g_folderWindow.GetFocusedPane();
            const std::optional<std::filesystem::path> path = g_folderWindow.GetCurrentPath(pane);
            if (! path.has_value() || path.value().empty())
            {
                break;
            }

            const std::wstring pathText = path.value().wstring();
            ShellExecuteW(hWnd, L"open", pathText.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
            break;
        }
        case IDM_APP_OPEN_FILE_EXPLORER_DESKTOP:
        case IDM_APP_OPEN_FILE_EXPLORER_DOCUMENTS:
        case IDM_APP_OPEN_FILE_EXPLORER_DOWNLOADS:
        case IDM_APP_OPEN_FILE_EXPLORER_PICTURES:
        case IDM_APP_OPEN_FILE_EXPLORER_MUSIC:
        case IDM_APP_OPEN_FILE_EXPLORER_VIDEOS:
        case IDM_APP_OPEN_FILE_EXPLORER_ONEDRIVE:
        {
            const GUID* folderId = nullptr;
            switch (wmId)
            {
                case IDM_APP_OPEN_FILE_EXPLORER_DESKTOP: folderId = &FOLDERID_Desktop; break;
                case IDM_APP_OPEN_FILE_EXPLORER_DOCUMENTS: folderId = &FOLDERID_Documents; break;
                case IDM_APP_OPEN_FILE_EXPLORER_DOWNLOADS: folderId = &FOLDERID_Downloads; break;
                case IDM_APP_OPEN_FILE_EXPLORER_PICTURES: folderId = &FOLDERID_Pictures; break;
                case IDM_APP_OPEN_FILE_EXPLORER_MUSIC: folderId = &FOLDERID_Music; break;
                case IDM_APP_OPEN_FILE_EXPLORER_VIDEOS: folderId = &FOLDERID_Videos; break;
                case IDM_APP_OPEN_FILE_EXPLORER_ONEDRIVE: folderId = &kKnownFolderIdOneDrive; break;
            }

            if (! folderId)
            {
                break;
            }

            wil::unique_cotaskmem_string folderPath;
            if (SUCCEEDED(SHGetKnownFolderPath(*folderId, 0, nullptr, folderPath.put())) && folderPath)
            {
                ShellExecuteW(hWnd, L"open", folderPath.get(), nullptr, nullptr, SW_SHOWNORMAL);
            }
            break;
        }
        case IDM_PANE_RENAME:
        case IDM_PANE_VIEW:
        case IDM_PANE_ALTERNATE_VIEW:
        case IDM_PANE_EDIT:
        case IDM_PANE_ALTERNATE_EDIT:
        case IDM_PANE_EDIT_NEW:
        case IDM_PANE_VIEW_SPACE:
        case IDM_PANE_COPY_TO_OTHER:
        case IDM_PANE_MOVE_TO_OTHER:
        case IDM_PANE_CREATE_DIR:
        case IDM_PANE_DELETE:
        case IDM_PANE_PERMANENT_DELETE:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);

            switch (wmId)
            {
                case IDM_PANE_RENAME: g_folderWindow.CommandRename(pane); break;
                case IDM_PANE_VIEW: g_folderWindow.CommandView(pane); break;
                case IDM_PANE_ALTERNATE_VIEW: g_folderWindow.CommandAlternateView(pane); break;
                case IDM_PANE_EDIT: g_folderWindow.CommandEdit(pane); break;
                case IDM_PANE_ALTERNATE_EDIT: g_folderWindow.CommandAlternateEdit(pane); break;
                case IDM_PANE_EDIT_NEW: g_folderWindow.CommandEditNew(pane); break;
                case IDM_PANE_VIEW_SPACE: g_folderWindow.CommandViewSpace(pane); break;
                case IDM_PANE_COPY_TO_OTHER: g_folderWindow.CommandCopyToOtherPane(pane); break;
                case IDM_PANE_MOVE_TO_OTHER: g_folderWindow.CommandMoveToOtherPane(pane); break;
                case IDM_PANE_CREATE_DIR: g_folderWindow.CommandCreateDirectory(pane); break;
                case IDM_PANE_DELETE: g_folderWindow.CommandDelete(pane); break;
                case IDM_PANE_PERMANENT_DELETE: g_folderWindow.CommandPermanentDelete(pane); break;
            }
            break;
        }
        case IDM_PANE_SORT_NAME:
        case IDM_PANE_SORT_EXTENSION:
        case IDM_PANE_SORT_TIME:
        case IDM_PANE_SORT_SIZE:
        case IDM_PANE_SORT_ATTRIBUTES:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);

            FolderView::SortBy sortBy = FolderView::SortBy::Name;
            switch (wmId)
            {
                case IDM_PANE_SORT_NAME: sortBy = FolderView::SortBy::Name; break;
                case IDM_PANE_SORT_EXTENSION: sortBy = FolderView::SortBy::Extension; break;
                case IDM_PANE_SORT_TIME: sortBy = FolderView::SortBy::Time; break;
                case IDM_PANE_SORT_SIZE: sortBy = FolderView::SortBy::Size; break;
                case IDM_PANE_SORT_ATTRIBUTES: sortBy = FolderView::SortBy::Attributes; break;
            }

            g_folderWindow.CycleSortBy(pane, sortBy);
            break;
        }
        case IDM_PANE_SORT_NONE:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);
            g_folderWindow.SetSort(pane, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
            break;
        }
        case IDM_PANE_DISPLAY_BRIEF:
        case IDM_PANE_DISPLAY_DETAILED:
        case IDM_PANE_DISPLAY_EXTRA_DETAILED:
        {
            const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
            g_folderWindow.SetActivePane(pane);

            FolderView::DisplayMode mode = FolderView::DisplayMode::Brief;
            switch (wmId)
            {
                case IDM_PANE_DISPLAY_BRIEF: mode = FolderView::DisplayMode::Brief; break;
                case IDM_PANE_DISPLAY_DETAILED: mode = FolderView::DisplayMode::Detailed; break;
                case IDM_PANE_DISPLAY_EXTRA_DETAILED: mode = FolderView::DisplayMode::ExtraDetailed; break;
            }
            g_folderWindow.SetDisplayMode(pane, mode);
            UpdatePaneMenuChecks();
            break;
        }
        case IDM_LEFT_SORT_NAME:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Left, FolderView::SortBy::Name);
            break;
        case IDM_LEFT_SORT_EXTENSION:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Left, FolderView::SortBy::Extension);
            break;
        case IDM_LEFT_SORT_TIME:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Left, FolderView::SortBy::Time);
            break;
        case IDM_LEFT_SORT_SIZE:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Left, FolderView::SortBy::Size);
            break;
        case IDM_LEFT_SORT_ATTRIBUTES:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Left, FolderView::SortBy::Attributes);
            break;
        case IDM_LEFT_SORT_NONE:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.SetSort(FolderWindow::Pane::Left, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
            break;
        case IDM_LEFT_DISPLAY_BRIEF:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Brief);
            UpdatePaneMenuChecks();
            break;
        case IDM_LEFT_DISPLAY_DETAILED:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::Detailed);
            UpdatePaneMenuChecks();
            break;
        case IDM_LEFT_DISPLAY_EXTRA_DETAILED:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Left, FolderView::DisplayMode::ExtraDetailed);
            UpdatePaneMenuChecks();
            break;
        case IDM_RIGHT_SORT_NAME:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Right, FolderView::SortBy::Name);
            break;
        case IDM_RIGHT_SORT_EXTENSION:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Right, FolderView::SortBy::Extension);
            break;
        case IDM_RIGHT_SORT_TIME:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Right, FolderView::SortBy::Time);
            break;
        case IDM_RIGHT_SORT_SIZE:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Right, FolderView::SortBy::Size);
            break;
        case IDM_RIGHT_SORT_ATTRIBUTES:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.CycleSortBy(FolderWindow::Pane::Right, FolderView::SortBy::Attributes);
            break;
        case IDM_RIGHT_SORT_NONE:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.SetSort(FolderWindow::Pane::Right, FolderView::SortBy::None, FolderView::SortDirection::Ascending);
            break;
        case IDM_RIGHT_DISPLAY_BRIEF:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Right, FolderView::DisplayMode::Brief);
            UpdatePaneMenuChecks();
            break;
        case IDM_RIGHT_DISPLAY_DETAILED:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Right, FolderView::DisplayMode::Detailed);
            UpdatePaneMenuChecks();
            break;
        case IDM_RIGHT_DISPLAY_EXTRA_DETAILED:
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.SetDisplayMode(FolderWindow::Pane::Right, FolderView::DisplayMode::ExtraDetailed);
            UpdatePaneMenuChecks();
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_ERROR:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Error);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_WARNING:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Warning);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_INFORMATION:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Information);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_BUSY:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Busy);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_HIDE:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugHideOverlaySample(FolderWindow::Pane::Left);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_ERROR_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Error);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_WARNING_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Warning);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_INFORMATION_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Left, FolderView::OverlaySeverity::Information);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_CANCELED:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySampleCanceled(FolderWindow::Pane::Left);
            break;
        case IDM_LEFT_OVERLAY_SAMPLE_BUSY_WITH_CANCEL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Left);
            g_folderWindow.DebugShowOverlaySampleBusyWithCancel(FolderWindow::Pane::Left);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_ERROR:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Error);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_WARNING:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Warning);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_INFORMATION:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Information);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_BUSY:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySample(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Busy);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_HIDE:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugHideOverlaySample(FolderWindow::Pane::Right);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_ERROR_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Error);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_WARNING_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Warning);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_INFORMATION_NONMODAL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySampleNonModal(FolderWindow::Pane::Right, FolderView::OverlaySeverity::Information);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_CANCELED:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySampleCanceled(FolderWindow::Pane::Right);
            break;
        case IDM_RIGHT_OVERLAY_SAMPLE_BUSY_WITH_CANCEL:
            if (! IsOverlaySampleEnabled())
            {
                break;
            }
            g_folderWindow.SetActivePane(FolderWindow::Pane::Right);
            g_folderWindow.DebugShowOverlaySampleBusyWithCancel(FolderWindow::Pane::Right);
            break;
        default:
        {
            const UINT cmdId  = wmId;
            const auto pathIt = g_navigatePathMenuTargets.find(cmdId);
            if (pathIt != g_navigatePathMenuTargets.end())
            {
                g_folderWindow.SetActivePane(pathIt->second.pane);
                g_folderWindow.SetFolderPath(pathIt->second.pane, pathIt->second.path);
                break;
            }

            const auto it = g_customThemeMenuIdToThemeId.find(cmdId);
            if (it != g_customThemeMenuIdToThemeId.end())
            {
                ApplyThemeId(hWnd, it->second);
                break;
            }

            const auto pluginIt = g_pluginMenuIdToPluginId.find(cmdId);
            if (pluginIt != g_pluginMenuIdToPluginId.end())
            {
                FileSystemPluginManager& plugins = FileSystemPluginManager::GetInstance();
                const HRESULT hr                 = plugins.SetActivePlugin(pluginIt->second, g_settings);
                if (SUCCEEDED(hr))
                {
                    const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
                    g_folderWindow.SetActivePane(pane);
                    static_cast<void>(g_folderWindow.SetFileSystemPluginForPane(pane, pluginIt->second));
                    UpdatePluginsMenuChecks();

                    const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(kAppId, g_settings);
                    if (FAILED(saveHr))
                    {
                        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kAppId);
                        std::wstring title                       = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
                        std::wstring message =
                            FormatStringResource(nullptr, IDS_FMT_SETTINGS_SAVE_FAILED, settingsPath.wstring(), static_cast<unsigned long>(saveHr));
                        g_folderWindow.ShowPaneAlertOverlay(
                            pane, FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, std::move(title), std::move(message), saveHr);
                        DBGOUT_ERROR(L"SaveSettings failed (hr=0x{:08X}) path={}\n", static_cast<unsigned long>(saveHr), settingsPath.wstring());
                    }
                }
                break;
            }

            const auto viewWithIt = g_viewWithMenuIdToActionId.find(cmdId);
            if (viewWithIt != g_viewWithMenuIdToActionId.end())
            {
                const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
                g_folderWindow.CommandViewWith(pane, viewWithIt->second);
                break;
            }

            const auto editWithIt = g_editWithMenuIdToActionId.find(cmdId);
            if (editWithIt != g_editWithMenuIdToActionId.end())
            {
                const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
                g_folderWindow.CommandEditWith(pane, editWithIt->second);
                break;
            }

            const auto userMenuIt = g_userMenuIdToActionId.find(cmdId);
            if (userMenuIt != g_userMenuIdToActionId.end())
            {
                const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
                g_folderWindow.CommandUserMenu(pane, userMenuIt->second);
                break;
            }

            const auto newTemplateIt = g_newTemplateMenuIdToTemplateId.find(cmdId);
            if (newTemplateIt != g_newTemplateMenuIdToTemplateId.end())
            {
                const FolderWindow::Pane pane = g_folderWindow.GetFocusedPane();
                g_folderWindow.CommandNewFromShellTemplate(pane, newTemplateIt->second);
                break;
            }

            if (const CommandInfo* info = FindCommandInfoByWmCommandId(cmdId))
            {
                ShowCommandNotImplementedMessage(hWnd, info->id);
                break;
            }

            return DefWindowProcW(hWnd, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(id), static_cast<WORD>(codeNotify)), reinterpret_cast<LPARAM>(hwndCtl));
        }
    }

    return 0;
}

LRESULT OnMainWindowSysCommand(HWND hWnd, WPARAM wParam, LPARAM lParam)
{
    if ((wParam & 0xFFF0u) == SC_KEYMENU && g_mainMenuHandle)
    {
        EnsureMenuHandles(hWnd);
        UpdatePaneMenuChecks();
        if (! g_menuBarVisible)
        {
            g_menuBarTemporarilyShown = true;
        }
        AdjustLayout(hWnd);
        if (lParam != 0)
        {
            if (! g_mainMenuBarHost.ActivateMnemonic(static_cast<wchar_t>(lParam)))
            {
                static_cast<void>(g_mainMenuBarHost.FocusFirstItem());
            }
        }
        else
        {
            static_cast<void>(g_mainMenuBarHost.FocusFirstItem());
        }
        return 0;
    }
    return DefWindowProcW(hWnd, WM_SYSCOMMAND, wParam, lParam);
}

void RestoreMainWindowFolderFocus(HWND hWnd)
{
    if (! IsWindowEnabled(hWnd))
    {
        return;
    }

    const HWND activeWindow = GetActiveWindow();
    if (activeWindow && activeWindow != hWnd)
    {
        return;
    }

    const HWND folderWindow = g_hFolderWindow.load(std::memory_order_acquire);
    if (! folderWindow)
    {
        return;
    }

    const HWND focused = GetFocus();
    if (focused && (focused == folderWindow || IsChild(folderWindow, focused)))
    {
        return;
    }

    SetFocus(folderWindow);
}

LRESULT OnMainWindowExitMenuLoop(HWND hWnd, [[maybe_unused]] BOOL isTrackPopupMenu)
{
    if (g_menuBarTemporarilyShown && ! g_menuBarVisible)
    {
        g_menuBarTemporarilyShown = false;
        AdjustLayout(hWnd);
    }

    if (GetActiveWindow() == hWnd)
    {
        RestoreMainWindowFolderFocus(hWnd);
    }
    return 0;
}

LRESULT OnMainWindowActivate(HWND hWnd, WORD state)
{
    if (state != WA_INACTIVE)
    {
        RestoreMainWindowFolderFocus(hWnd);
    }
    return 0;
}

LRESULT OnMainWindowSetFocus(HWND hWnd)
{
    RestoreMainWindowFolderFocus(hWnd);
    return 0;
}

LRESULT OnMainWindowPaint(HWND hWnd)
{
    StartupMetrics::MarkFirstPaint(kMainWindowId);
    wil::unique_hdc_paint paint_dc = wil::BeginPaint(hWnd);
    return 0;
}

LRESULT OnMainWindowStartupInputReady([[maybe_unused]] HWND hWnd)
{
    StartupMetrics::MarkInputReady(kMainWindowId);
#ifdef ENABLE_TESTS
    if (! IsRunningAnySelfTest())
#endif
    {
        SplashScreen::RequestCloseIfExist();
    }
    CrashHandler::ShowPreviousCrashUiIfPresent(hWnd);
    return 0;
}

LRESULT OnMainWindowSize(HWND hWnd, [[maybe_unused]] UINT width, [[maybe_unused]] UINT height)
{
    AdjustLayout(hWnd);
    return 0;
}

LRESULT OnMainWindowDpiChanged(HWND hWnd, UINT newDpi, const RECT* prcNew)
{
    g_folderWindow.OnDpiChanged(static_cast<float>(newDpi));
    if (prcNew != nullptr)
    {
        SetWindowPos(hWnd, nullptr, prcNew->left, prcNew->top, prcNew->right - prcNew->left, prcNew->bottom - prcNew->top, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    AdjustLayout(hWnd);
    return 0;
}

LRESULT OnMainWindowThemeChanged(HWND hWnd)
{
    LocaleFormatting::InvalidateFormatLocaleCache();
    ApplyAppTheme(hWnd);
    return 0;
}

LRESULT OnMainWindowSettingsApplied(HWND hWnd)
{
    ApplyCurrentSettingsToRunningApp(hWnd);
    return 0;
}

LRESULT OnMainWindowPluginsChanged(HWND hWnd)
{
    RefreshRunningPluginsFromSettings(hWnd);
    return 0;
}

LRESULT OnMainWindowConnectionManagerConnect([[maybe_unused]] HWND hWnd, WPARAM wParam, LPARAM lParam)
{
    auto name = TakeMessagePayload<std::wstring>(lParam);
    if (! name || name->empty())
    {
        return 0;
    }

#ifdef ENABLE_TESTS
    {
        const std::scoped_lock lock(g_debugConnectionManagerConnectMutex);
        g_debugConnectionManagerConnectSeen = true;
        g_debugConnectionManagerConnectPane = static_cast<uint8_t>(wParam == 1u ? 1u : 0u);
        g_debugConnectionManagerConnectName = *name;
    }
#endif

    const FolderWindow::Pane pane = (wParam == 1u) ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
    g_folderWindow.SetActivePane(pane);

    std::wstring target = L"nav:";
    target.append(*name);
    g_folderWindow.SetFolderPath(pane, std::filesystem::path(std::move(target)));
    return 0;
}

struct ShutdownCloseWindowSnapshotContext final
{
    DWORD pid        = 0;
    HWND excludeHwnd = nullptr;
    std::vector<HWND> windows;
};

constexpr std::wstring_view kInternalWindowClassPrefix = L"RedSalamander.";
constexpr UINT kShutdownCloseTimeoutMs                 = 1000;
constexpr UINT kFinalizeMainWindowCloseMessage         = WM_APP + 0x47;

BOOL CALLBACK ShutdownCloseEnumWindowsProc(HWND hwnd, LPARAM lParam) noexcept
{
    auto* ctx = reinterpret_cast<ShutdownCloseWindowSnapshotContext*>(lParam);
    if (! ctx || ! hwnd || hwnd == ctx->excludeHwnd)
    {
        return TRUE;
    }

    DWORD windowPid = 0;
    static_cast<void>(GetWindowThreadProcessId(hwnd, &windowPid));
    if (windowPid != ctx->pid)
    {
        return TRUE;
    }

    // Only unowned top-level windows; owned windows will be torn down automatically with their owner.
    if (GetWindow(hwnd, GW_OWNER) != nullptr || GetParent(hwnd) != nullptr)
    {
        return TRUE;
    }

    wchar_t className[256]{};
    const int len = GetClassNameW(hwnd, className, static_cast<int>(std::size(className)));
    if (len <= 0)
    {
        return TRUE;
    }

    if (static_cast<size_t>(len) < kInternalWindowClassPrefix.size())
    {
        return TRUE;
    }

    const std::wstring_view classNameView(className, static_cast<size_t>(len));
    if (! OrdinalString::StartsWithNoCase(classNameView, kInternalWindowClassPrefix))
    {
        return TRUE;
    }

    ctx->windows.push_back(hwnd);
    return TRUE;
}

[[nodiscard]] bool CloseUnownedTopLevelRedSalamanderWindowsForShutdown(HWND excludeHwnd) noexcept
{
    // Our internal window classes follow the "RedSalamander.*" naming convention. Close any unowned top-level windows
    // in this process so they can tear down D2D/D3D resources before DLL/process shutdown.
    ShutdownCloseWindowSnapshotContext ctx{};
    ctx.pid         = GetCurrentProcessId();
    ctx.excludeHwnd = excludeHwnd;
    EnumWindows(&ShutdownCloseEnumWindowsProc, reinterpret_cast<LPARAM>(&ctx));

    bool allClosed = true;
    for (HWND hwnd : ctx.windows)
    {
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            continue;
        }

        DWORD_PTR ignoredResult = 0;
        const LRESULT sent      = SendMessageTimeoutW(hwnd, WM_CLOSE, 0, 0, SMTO_ABORTIFHUNG | SMTO_NORMAL, kShutdownCloseTimeoutMs, &ignoredResult);
        if (sent == 0)
        {
            if (IsWindow(hwnd) != FALSE)
            {
                allClosed = false;
            }
            continue;
        }

        if (IsWindow(hwnd) != FALSE)
        {
            allClosed = false;
        }
    }

    return allClosed;
}

LRESULT OnMainWindowClose(HWND hWnd)
{
#ifdef ENABLE_TESTS
    if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
    {
        if (! DestroyWindow(hWnd))
        {
            PostQuitMessage(g_selfTestExitCode);
        }
        return 0;
    }
#endif

    if (! g_folderWindow.ConfirmCancelAllFileOperations(hWnd))
    {
        return 0;
    }

    // Ensure non-owned top-level windows release graphics resources before we tear down the process.
    // The D2D debug layer breaks on shutdown if any D2D objects are still alive at DLL unload time.
    if (! CloseUnownedTopLevelRedSalamanderWindowsForShutdown(hWnd))
    {
        Debug::Warning(L"Shutdown: proceeding while one or more internal top-level windows are still closing.");
    }

    // Post the final destroy so any auxiliary windows that just received WM_CLOSE can process their own
    // posted teardown before the main window posts WM_QUIT.
    static_cast<void>(PostMessageW(hWnd, kFinalizeMainWindowCloseMessage, 0, 0));
    return 0;
}

LRESULT OnMainWindowDestroy(HWND hWnd)
{
    SettingsHotReload::Stop();
    g_mainMenuBarHost.Destroy();

#ifdef ENABLE_TESTS
    if (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest)
    {
        g_folderWindow.CloseAllViewers();
        g_folderWindow.SetSettings(nullptr);
        if (g_hFolderWindow.exchange(nullptr, std::memory_order_acq_rel))
        {
            g_folderWindow.Destroy();
        }
        FileSystemPluginManager::GetInstance().Shutdown(g_settings);
        ViewerPluginManager::GetInstance().Shutdown(g_settings);
        SessionState::Clear();
        ShutdownSelfTestMonitor();
        TraceSelfTestExitCode(L"OnMainWindowDestroy: PostQuitMessage",
                              (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest) ? g_selfTestExitCode : 0);
        PostQuitMessage((g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest) ? g_selfTestExitCode : 0);
        return 0;
    }
#endif

    SaveAppSettings(hWnd);
    SessionState::Clear();

#ifdef ENABLE_TESTS
    ShutdownSelfTestMonitor();
    TraceSelfTestExitCode(L"OnMainWindowDestroy: PostQuitMessage",
                          (g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest) ? g_selfTestExitCode : 0);
    PostQuitMessage((g_runFileOpsSelfTest || g_runCompareDirectoriesSelfTest || g_runCommandsSelfTest) ? g_selfTestExitCode : 0);
#else
    PostQuitMessage(0);
#endif

    return 0;
}

} // namespace

HWND GetAboutDialogHandle() noexcept
{
    const HWND hwnd = FindWindowW(kAboutDialogWindowClassName, nullptr);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

HWND GetFatalErrorDialogHandle() noexcept
{
    const HWND hwnd = FindWindowW(kFatalErrorDialogWindowClassName, nullptr);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

#ifdef ENABLE_TESTS
void DebugResetConnectionManagerConnectNavigation() noexcept
{
    const std::scoped_lock lock(g_debugConnectionManagerConnectMutex);
    g_debugConnectionManagerConnectSeen = false;
    g_debugConnectionManagerConnectPane = 0u;
    g_debugConnectionManagerConnectName.clear();
}

bool DebugGetConnectionManagerConnectNavigation(uint8_t& outPane, std::wstring& outName) noexcept
{
    const std::scoped_lock lock(g_debugConnectionManagerConnectMutex);
    if (! g_debugConnectionManagerConnectSeen)
    {
        return false;
    }

    outPane = g_debugConnectionManagerConnectPane;
    outName = g_debugConnectionManagerConnectName;
    return true;
}

void DebugShowFatalErrorDialog(HWND ownerWindow, const wchar_t* caption, const wchar_t* message) noexcept
{
    ShowFatalErrorDialog(ownerWindow, caption, message);
}

bool DebugGetFatalErrorDialogSnapshot(FatalErrorDialogDebugSnapshot& out) noexcept
{
    if (const HWND hwnd = GetFatalErrorDialogHandle(); hwnd && IsWindow(hwnd) != FALSE)
    {
        return SendMessageW(hwnd, kFatalErrorDialogDebugGetSnapshotMessage, 0, reinterpret_cast<LPARAM>(&out)) != FALSE;
    }

    out = {};
    return false;
}

bool DebugScrollFatalErrorDialogByWheelDetents(int detents) noexcept
{
    if (detents == 0)
    {
        return true;
    }

    if (const HWND hwnd = GetFatalErrorDialogHandle(); hwnd && IsWindow(hwnd) != FALSE)
    {
        return SendMessageW(hwnd, kFatalErrorDialogDebugScrollByWheelMessage, static_cast<WPARAM>(detents), 0) != FALSE;
    }

    return false;
}
#endif
// Processes messages for the main window.
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
        case WM_CREATE: InitPostedPayloadWindow(hWnd); return OnMainWindowCreate(hWnd, reinterpret_cast<const CREATESTRUCTW*>(lParam));
#ifdef ENABLE_TESTS
        case kCommandsSelfTestStartMessage: return RunCommandsSelfTestAndRequestShutdown(hWnd);
#endif
        case WM_COMMAND: return OnMainWindowCommand(hWnd, LOWORD(wParam), HIWORD(wParam), reinterpret_cast<HWND>(lParam));
        case WndMsg::kFunctionBarInvoke: return OnFunctionBarInvoke(hWnd, wParam, lParam);
        case WndMsg::kSettingsApplied: return OnMainWindowSettingsApplied(hWnd);
        case WndMsg::kPluginsChanged: return OnMainWindowPluginsChanged(hWnd);
        case WndMsg::kConnectionManagerConnect: return OnMainWindowConnectionManagerConnect(hWnd, wParam, lParam);
        case WndMsg::kSettingsFileChanged: return OnMainWindowSettingsFileChanged(hWnd, lParam);
        case WndMsg::kPreferencesRequestSettingsSnapshot: CaptureRuntimeSettings(hWnd); return 0;
        case WndMsg::kPaneRestoreFolderFocus:
        {
            if (IsIconic(hWnd) != FALSE)
            {
                ShowWindow(hWnd, SW_RESTORE);
            }
            static_cast<void>(SetActiveWindow(hWnd));
            static_cast<void>(SetForegroundWindow(hWnd));
            RestoreMainWindowFolderFocus(hWnd);
            if (g_folderWindow.GetFocusedFolderViewHwnd() == nullptr)
            {
                static_cast<void>(g_folderWindow.TryRestoreActivePaneFolderViewFocus());
            }
            return g_folderWindow.GetFocusedFolderViewHwnd() != nullptr ? 1 : 0;
        }
        case WM_TIMER: return OnMainWindowTimer(hWnd, static_cast<UINT_PTR>(wParam));
        case WM_NCDESTROY: static_cast<void>(DrainPostedPayloadsForWindow(hWnd)); return DefWindowProcW(hWnd, message, wParam, lParam);
        case WM_NCACTIVATE:
            if (g_hFolderWindow.load(std::memory_order_acquire))
            {
                ApplyTitleBarTheme(hWnd, g_folderWindow.GetTheme(), wParam != FALSE);
            }
            return DefWindowProcW(hWnd, message, wParam, lParam);
        case WM_SYSCOMMAND: return OnMainWindowSysCommand(hWnd, wParam, lParam);
        case WM_INITMENUPOPUP: OnInitMenuPopup(hWnd, reinterpret_cast<HMENU>(wParam)); return 0;
        case WM_ACTIVATE: return OnMainWindowActivate(hWnd, LOWORD(wParam));
        case WM_EXITMENULOOP: return OnMainWindowExitMenuLoop(hWnd, static_cast<BOOL>(wParam));
        case WM_SETFOCUS: return OnMainWindowSetFocus(hWnd);
        case WM_PAINT: return OnMainWindowPaint(hWnd);
        case WndMsg::kAppStartupInputReady: return OnMainWindowStartupInputReady(hWnd);
        case WM_SIZE: return OnMainWindowSize(hWnd, LOWORD(lParam), HIWORD(lParam));
        case WM_DPICHANGED: return OnMainWindowDpiChanged(hWnd, static_cast<UINT>(HIWORD(wParam)), reinterpret_cast<const RECT*>(lParam));
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lParam))
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hWnd, *info, 760, 480);
            }
            return 0;
        case WM_ERASEBKGND: return 1;
        case WM_DEVICECHANGE:
        {
            const HWND folderWindow = g_hFolderWindow.load(std::memory_order_acquire);
            if (folderWindow)
            {
                SendMessageW(folderWindow, WM_DEVICECHANGE, wParam, lParam);
            }
        }
            return DefWindowProcW(hWnd, message, wParam, lParam);
        case kFinalizeMainWindowCloseMessage:
            if (IsWindow(hWnd) != FALSE)
            {
                DestroyWindow(hWnd);
            }
            return 0;
        case WM_CLOSE: return OnMainWindowClose(hWnd);
        case WM_SETTINGCHANGE:
        case WM_THEMECHANGED:
        case WM_DWMCOLORIZATIONCOLORCHANGED:
        case WM_SYSCOLORCHANGE: return OnMainWindowThemeChanged(hWnd);
        case WM_DESTROY: return OnMainWindowDestroy(hWnd);
        default: return DefWindowProcW(hWnd, message, wParam, lParam);
    }
}
