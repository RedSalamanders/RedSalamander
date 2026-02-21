#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <filesystem>
#include <format>
#include <fstream>
#include <iterator>
#include <memory>
#include <optional>
#include <limits>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include <Windows.h>

#include <commctrl.h>
#include <commdlg.h>
#include <d2d1.h>
#include <d2d1helper.h>
#include <shellapi.h>
#include <shellscalingapi.h>
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shcore.lib")

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027) 
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <wincodec.h>
#pragma comment(lib, "Windowscodecs.lib")

// Define the ETW provider for RedSalamanderMonitor.exe
// Each executable must have its own provider instance with the same GUID
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h" // Must be included before EtwListener.h for InfoParam definition

#include "ColorTextView.h"
#include "Configuration.h"
#include "EtwListener.h"
#include "ExceptionHelpers.h" // Shared exception handling utilities
#include "RedSalamanderMonitor.h"
#include "SettingsStore.h"
#include "resource.h"

// Global Variables:
// All globals below are accessed exclusively from the UI thread (message loop).
// The only cross-thread interaction is EtwListener's worker thread calling
// g_colorView.QueueEtwEvent(), which is thread-safe via atomic HWND + critical section.
HINSTANCE g_hInstance = NULL;  // current instance
ColorTextView g_colorView;     // ColorTextView instance for the right panel
wil::unique_hwnd g_hColorView; // ColorTextView window handle
wil::unique_hwnd g_hToolbar;   // Toolbar window handle
wil::unique_hwnd g_hStatusBar; // Status bar window handle
using unique_himagelist = wil::unique_any<HIMAGELIST, decltype(&ImageList_Destroy), ImageList_Destroy>;
static unique_himagelist g_toolbarImageList; // Image list backing the toolbar
bool g_showIds            = true;            // Show Process/Thread IDs in output
bool g_alwaysOnTop        = false;           // Main window always-on-top flag
bool g_toolbarVisible     = true;            // Toolbar visibility (menu state)
bool g_lineNumbersVisible = true;            // Line numbers menu state
bool g_autoScrollEnabled  = true;            // Auto-scroll menu state
// Auto-scroll state is now managed by ColorTextView (_autoScrollEnabled member)
static std::unique_ptr<EtwListener> g_etwListener; // ETW real-time event listener
Common::Settings::Settings g_settings;

// Filter state: bitmask where bit N corresponds to InfoParam::Type value N (0x1F = all 5 types enabled)
static uint32_t g_filterMask  = Debug::InfoParam::Type::All; // All types enabled by default
static int g_lastFilterPreset = -1;                          // -1 = custom, 0 = Errors Only, 1 = Errors+Warnings, 2 = All, 3 = Errors+Debug

// Status bar update timer
static constexpr UINT_PTR kStatusBarTimerId      = 100;
static constexpr UINT kStatusBarUpdateIntervalMs = 500; // Update every 500ms
static uint64_t g_lastMessageCount               = 0;   // Track message rate for adaptive refresh

namespace
{
constexpr wchar_t kAppId[]    = L"RedSalamanderMonitor";
constexpr wchar_t kWindowId[] = L"MonitorWindow";

HMENU g_viewThemeMenu = nullptr;

constexpr UINT kCustomThemeMenuIdFirst = 32900u;
constexpr UINT kCustomThemeMenuIdLast  = 33099u;

std::unordered_map<UINT, std::wstring> g_customThemeMenuIdToThemeId;
std::unordered_map<std::wstring, UINT> g_customThemeIdToMenuId;
std::vector<Common::Settings::ThemeDefinition> g_fileThemes;

struct ModalMessageDialogState
{
    const wchar_t* caption = nullptr;
    const wchar_t* message = nullptr;
};

INT_PTR CALLBACK ModalMessageDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<ModalMessageDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
    switch (msg)
    {
        case WM_INITDIALOG:
        {
            state = reinterpret_cast<ModalMessageDialogState*>(lp);
            SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));

            if (state)
            {
                if (state->caption && state->caption[0] != L'\0')
                {
                    SetWindowTextW(dlg, state->caption);
                }

                if (state->message && state->message[0] != L'\0')
                {
                    SetDlgItemTextW(dlg, IDC_MODAL_MESSAGE_TEXT, state->message);
                }
            }

            wchar_t okText[64]{};
            const int okLength = LoadStringW(GetModuleHandleW(nullptr), IDS_BTN_OK, okText, static_cast<int>(sizeof(okText) / sizeof(okText[0])));
            if (okLength > 0)
            {
                SetDlgItemTextW(dlg, IDOK, okText);
            }

            return static_cast<INT_PTR>(TRUE);
        }
        case WM_COMMAND:
        {
            const WORD id = LOWORD(wp);
            if (id == IDOK || id == IDCANCEL)
            {
                EndDialog(dlg, id);
                return static_cast<INT_PTR>(TRUE);
            }
            break;
        }
    }
    return static_cast<INT_PTR>(FALSE);
}

void ShowModalMessageDialog(HINSTANCE instance, HWND owner, const wchar_t* caption, const wchar_t* message) noexcept
{
    ModalMessageDialogState state{};
    state.caption = caption;
    state.message = message;

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
    DialogBoxParamW(instance, MAKEINTRESOURCEW(IDD_MODAL_MESSAGE), owner, ModalMessageDialogProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)
}

std::filesystem::path GetThemesDirectory() noexcept
{
    wil::unique_cotaskmem_string modulePath;
    const HRESULT hr = wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, modulePath);
    if (FAILED(hr) || ! modulePath)
    {
        return {};
    }

    return std::filesystem::path(modulePath.get()).parent_path() / L"Themes";
}

std::wstring EscapeMenuLabel(std::wstring_view raw)
{
    std::wstring result;
    result.reserve(raw.size());

    for (const wchar_t ch : raw)
    {
        result.push_back(ch);
        if (ch == L'&')
        {
            result.push_back(L'&');
        }
    }

    return result;
}

// Convert a value expressed in DIPs (96-DPI) to physical pixels for a given window
inline int DipsToPx(HWND hwnd, int dip)
{
    UINT dpi = USER_DEFAULT_SCREEN_DPI;
    if (hwnd)
    {
        dpi = GetDpiForWindow(hwnd);
    }
    return MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
}

bool IsProcessElevated() noexcept
{
    wil::unique_handle token;
    if (! ::OpenProcessToken(::GetCurrentProcess(), TOKEN_QUERY, token.put()))
    {
        return false;
    }

    TOKEN_ELEVATION elevation{};
    DWORD size = 0;
    if (! ::GetTokenInformation(token.get(), TokenElevation, &elevation, sizeof(elevation), &size))
    {
        return false;
    }

    return elevation.TokenIsElevated != 0;
}

std::wstring QuoteCommandLineArg(std::wstring_view arg)
{
    if (arg.empty())
    {
        return L"\"\"";
    }

    const bool needsQuotes = arg.find_first_of(L" \t\n\v\"") != std::wstring_view::npos;
    if (! needsQuotes)
    {
        return std::wstring{arg};
    }

    std::wstring result;
    result.reserve(arg.size() + 2);
    result.push_back(L'"');

    size_t backslashCount = 0;
    for (const wchar_t ch : arg)
    {
        if (ch == L'\\')
        {
            ++backslashCount;
            continue;
        }

        if (ch == L'"')
        {
            result.append(backslashCount * 2 + 1, L'\\');
            result.push_back(L'"');
            backslashCount = 0;
            continue;
        }

        if (backslashCount != 0)
        {
            result.append(backslashCount, L'\\');
            backslashCount = 0;
        }

        result.push_back(ch);
    }

    if (backslashCount != 0)
    {
        result.append(backslashCount * 2, L'\\');
    }

    result.push_back(L'"');
    return result;
}

std::wstring BuildRelaunchParameters(std::wstring_view extraArg) noexcept
{
    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));

    std::wstring params;
    bool alreadyHasExtra = false;

    if (argv && argc > 1)
    {
        for (int i = 1; i < argc; ++i)
        {
            if (! extraArg.empty() && extraArg == argv.get()[i])
            {
                alreadyHasExtra = true;
            }

            if (! params.empty())
            {
                params.push_back(L' ');
            }
            params.append(QuoteCommandLineArg(argv.get()[i]));
        }
    }

    if (! extraArg.empty() && ! alreadyHasExtra)
    {
        if (! params.empty())
        {
            params.push_back(L' ');
        }
        params.append(QuoteCommandLineArg(extraArg));
    }

    return params;
}

bool HasCommandLineArg(std::wstring_view arg) noexcept
{
    if (arg.empty())
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
        if (arg == argv.get()[i])
        {
            return true;
        }
    }

    return false;
}

bool RelaunchSelfElevated(HWND owner, std::wstring_view extraArg) noexcept
{
    wil::unique_cotaskmem_string exePath;
    const HRESULT hr = wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, exePath);
    if (FAILED(hr) || ! exePath)
    {
        return false;
    }

    const std::wstring params = BuildRelaunchParameters(extraArg);

    SHELLEXECUTEINFOW execInfo{};
    execInfo.cbSize       = sizeof(execInfo);
    execInfo.fMask        = SEE_MASK_NOCLOSEPROCESS;
    execInfo.hwnd         = owner;
    execInfo.lpVerb       = L"runas";
    execInfo.lpFile       = exePath.get();
    execInfo.lpParameters = params.empty() ? nullptr : params.c_str();
    execInfo.nShow        = SW_SHOWNORMAL;

    if (! ::ShellExecuteExW(&execInfo))
    {
        return false;
    }

    wil::unique_handle launchedProcess(execInfo.hProcess);
    return true;
}

bool IsHighContrastEnabled() noexcept
{
    HIGHCONTRASTW hc{};
    hc.cbSize = sizeof(hc);
    if (! SystemParametersInfoW(SPI_GETHIGHCONTRAST, hc.cbSize, &hc, 0))
    {
        return false;
    }
    return (hc.dwFlags & HCF_HIGHCONTRASTON) != 0;
}

bool IsSystemDarkModeEnabled() noexcept
{
    DWORD appsUseLightTheme = 1;
    DWORD dataSize          = sizeof(appsUseLightTheme);
    const LSTATUS status    = RegGetValueW(HKEY_CURRENT_USER,
                                        L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                                        L"AppsUseLightTheme",
                                        RRF_RT_REG_DWORD,
                                        nullptr,
                                        &appsUseLightTheme,
                                        &dataSize);
    if (status != ERROR_SUCCESS)
    {
        return false;
    }
    return appsUseLightTheme == 0;
}

Common::Settings::MonitorFilterPreset PresetFromLegacy(int legacy) noexcept
{
    switch (legacy)
    {
        case 0: return Common::Settings::MonitorFilterPreset::ErrorsOnly;
        case 1: return Common::Settings::MonitorFilterPreset::ErrorsWarnings;
        case 2: return Common::Settings::MonitorFilterPreset::AllTypes;
        default: return Common::Settings::MonitorFilterPreset::Custom;
    }
}

int LegacyFromPreset(Common::Settings::MonitorFilterPreset preset) noexcept
{
    switch (preset)
    {
        case Common::Settings::MonitorFilterPreset::ErrorsOnly: return 0;
        case Common::Settings::MonitorFilterPreset::ErrorsWarnings: return 1;
        case Common::Settings::MonitorFilterPreset::AllTypes: return 2;
        case Common::Settings::MonitorFilterPreset::Custom:
        default: return -1;
    }
}

int InferLegacyPresetFromMask(uint32_t mask) noexcept
{
    // Mask bit mapping (Active Filter menu):
    // 0x01=Text, 0x02=Error, 0x04=Warning, 0x08=Info, 0x10=Debug
    mask &= Debug::InfoParam::Type::All;
    switch (mask)
    {
        case 0x02: return 0; // Errors only
        case 0x06: return 1; // Errors+Warnings
        case 0x1F: return 2; // All types
        case 0x12: return 3; // Errors+Debug
        default: return -1;  // Custom
    }
}

D2D1_COLOR_F D2DFromArgb(uint32_t argb) noexcept
{
    const float a = static_cast<float>((argb >> 24) & 0xFFu) / 255.0f;
    const float r = static_cast<float>((argb >> 16) & 0xFFu) / 255.0f;
    const float g = static_cast<float>((argb >> 8) & 0xFFu) / 255.0f;
    const float b = static_cast<float>(argb & 0xFFu) / 255.0f;
    return D2D1::ColorF(r, g, b, a);
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

void ApplyMonitorThemeOverrides(ColorTextView::Theme& theme, const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    const auto apply = [&](std::wstring_view key, D2D1_COLOR_F& target) noexcept
    {
        const auto it = colors.find(std::wstring(key));
        if (it == colors.end())
        {
            return;
        }
        target = D2DFromArgb(it->second);
    };

    apply(L"monitor.textView.bg", theme.bg);
    apply(L"monitor.textView.fg", theme.fg);
    apply(L"monitor.textView.caret", theme.caret);
    apply(L"monitor.textView.selection", theme.selection);
    apply(L"monitor.textView.searchHighlight", theme.searchHighlight);
    apply(L"monitor.textView.gutterBg", theme.gutterBg);
    apply(L"monitor.textView.gutterFg", theme.gutterFg);
    apply(L"monitor.textView.metaText", theme.metaText);
    apply(L"monitor.textView.metaError", theme.metaError);
    apply(L"monitor.textView.metaWarning", theme.metaWarning);
    apply(L"monitor.textView.metaInfo", theme.metaInfo);
    apply(L"monitor.textView.metaDebug", theme.metaDebug);
}

ColorTextView::Theme MakeMonitorThemeHighContrast() noexcept
{
    const auto sys = [](int idx, float alpha = 1.0f) noexcept
    {
        const COLORREF c = GetSysColor(idx);
        return D2D1::ColorF(
            static_cast<float>(GetRValue(c)) / 255.0f, static_cast<float>(GetGValue(c)) / 255.0f, static_cast<float>(GetBValue(c)) / 255.0f, alpha);
    };

    ColorTextView::Theme t;
    t.bg              = sys(COLOR_WINDOW);
    t.fg              = sys(COLOR_WINDOWTEXT);
    t.caret           = t.fg;
    t.selection       = sys(COLOR_HIGHLIGHT, 0.45f);
    t.searchHighlight = sys(COLOR_HIGHLIGHT, 0.35f);
    t.gutterBg        = sys(COLOR_BTNFACE);
    t.gutterFg        = sys(COLOR_WINDOWTEXT);
    t.metaText        = sys(COLOR_WINDOWTEXT, 0.85f);
    t.metaError       = sys(COLOR_HIGHLIGHTTEXT);
    t.metaWarning     = sys(COLOR_HIGHLIGHTTEXT);
    t.metaInfo        = sys(COLOR_HIGHLIGHTTEXT);
    t.metaDebug       = sys(COLOR_HIGHLIGHTTEXT);
    return t;
}

ColorTextView::Theme MakeMonitorThemeLight() noexcept
{
    return ColorTextView::Theme{};
}

ColorTextView::Theme MakeMonitorThemeDark() noexcept
{
    ColorTextView::Theme t;
    t.bg              = D2D1::ColorF(0.08f, 0.08f, 0.08f);
    t.fg              = D2D1::ColorF(0.92f, 0.92f, 0.92f);
    t.caret           = t.fg;
    t.selection       = D2D1::ColorF(0.20f, 0.55f, 0.95f, 0.35f);
    t.searchHighlight = D2D1::ColorF(1.00f, 0.85f, 0.05f, 0.35f);
    t.gutterBg        = D2D1::ColorF(0.12f, 0.12f, 0.12f);
    t.gutterFg        = D2D1::ColorF(0.65f, 0.65f, 0.65f);
    t.metaText        = D2D1::ColorF(0.65f, 0.65f, 0.65f);
    t.metaError       = D2D1::ColorF(D2D1::ColorF::Red);
    t.metaWarning     = D2D1::ColorF(D2D1::ColorF::Orange);
    t.metaInfo        = D2D1::ColorF(D2D1::ColorF::DodgerBlue);
    t.metaDebug       = D2D1::ColorF(D2D1::ColorF::MediumPurple);
    return t;
}

ColorTextView::Theme ResolveMonitorTheme() noexcept
{
    if (IsHighContrastEnabled())
    {
        return MakeMonitorThemeHighContrast();
    }

    std::wstring_view themeId                       = g_settings.theme.currentThemeId;
    const Common::Settings::ThemeDefinition* custom = nullptr;
    if (themeId.rfind(L"user/", 0) == 0)
    {
        custom = FindThemeById(themeId);
    }

    std::wstring_view baseThemeId                               = themeId;
    const std::unordered_map<std::wstring, uint32_t>* overrides = nullptr;
    if (custom)
    {
        baseThemeId = custom->baseThemeId;
        overrides   = &custom->colors;
    }

    const bool systemDark = IsSystemDarkModeEnabled();
    ColorTextView::Theme theme;

    if (baseThemeId == L"builtin/highContrast")
    {
        theme = MakeMonitorThemeHighContrast();
    }
    else if (baseThemeId == L"builtin/dark")
    {
        theme = MakeMonitorThemeDark();
    }
    else if (baseThemeId == L"builtin/light")
    {
        theme = MakeMonitorThemeLight();
    }
    else if (baseThemeId == L"builtin/rainbow")
    {
        theme = systemDark ? MakeMonitorThemeDark() : MakeMonitorThemeLight();
    }
    else
    {
        theme = systemDark ? MakeMonitorThemeDark() : MakeMonitorThemeLight();
    }

    if (overrides)
    {
        ApplyMonitorThemeOverrides(theme, *overrides);
    }

    return theme;
}

void ApplyMonitorTheme() noexcept
{
    const ColorTextView::Theme theme = ResolveMonitorTheme();
    g_colorView.SetTheme(theme);
}

bool TryFindMenuPathToCommand(HMENU menu, UINT commandId, std::vector<HMENU>& path) noexcept
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

int FindMenuItemPosById(HMENU menu, UINT id) noexcept
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

void DeleteMenuItemsFromPosition(HMENU menu, int startPos) noexcept
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

void EnsureThemeMenuHandle(HWND hWnd) noexcept
{
    if (g_viewThemeMenu)
    {
        return;
    }

    HMENU mainMenu = GetMenu(hWnd);
    if (! mainMenu)
    {
        return;
    }

    std::vector<HMENU> themePath;
    if (TryFindMenuPathToCommand(mainMenu, IDM_VIEW_THEME_SYSTEM, themePath) && ! themePath.empty())
    {
        g_viewThemeMenu = themePath.back();
    }
}

void RebuildThemeMenuDynamicItems(HWND hWnd)
{
    EnsureThemeMenuHandle(hWnd);
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

    std::vector<const Common::Settings::ThemeDefinition*> settingsThemes;
    settingsThemes.reserve(settingsThemesById.size());
    for (const auto& [id, def] : settingsThemesById)
    {
        settingsThemes.push_back(def);
    }

    std::vector<const Common::Settings::ThemeDefinition*> fileThemes;
    fileThemes.reserve(g_fileThemes.size());
    for (const auto& def : g_fileThemes)
    {
        if (def.id.rfind(L"user/", 0) != 0)
        {
            continue;
        }
        if (settingsThemesById.contains(def.id))
        {
            continue; // settings version wins
        }
        fileThemes.push_back(&def);
    }

    const auto byNameThenId = [](const Common::Settings::ThemeDefinition* a, const Common::Settings::ThemeDefinition* b)
    {
        if (a->name == b->name)
        {
            return a->id < b->id;
        }
        return a->name < b->name;
    };

    std::sort(fileThemes.begin(), fileThemes.end(), byNameThenId);
    std::sort(settingsThemes.begin(), settingsThemes.end(), byNameThenId);

    if (! fileThemes.empty() || ! settingsThemes.empty())
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

        if (! addThemes(fileThemes))
        {
            return;
        }

        if (! fileThemes.empty() && ! settingsThemes.empty() && nextId <= kCustomThemeMenuIdLast)
        {
            if (! AppendMenuW(g_viewThemeMenu, MF_SEPARATOR, 0, nullptr))
            {
                return;
            }
        }

        if (! addThemes(settingsThemes))
        {
            return;
        }
    }

    DrawMenuBar(hWnd);
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

    if (currentThemeId == L"builtin/highContrast")
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
    if (currentThemeId == L"builtin/light")
    {
        checkedId = IDM_VIEW_THEME_LIGHT;
    }
    else if (currentThemeId == L"builtin/dark")
    {
        checkedId = IDM_VIEW_THEME_DARK;
    }
    else if (currentThemeId == L"builtin/rainbow")
    {
        checkedId = IDM_VIEW_THEME_RAINBOW;
    }

    CheckMenuRadioItem(g_viewThemeMenu, IDM_VIEW_THEME_SYSTEM, IDM_VIEW_THEME_RAINBOW, checkedId, MF_BYCOMMAND);
}

void SaveMonitorSettings(HWND hWnd) noexcept
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

            g_settings.windows[kWindowId] = std::move(wp);
        }
    }

    if (! g_settings.monitor)
    {
        g_settings.monitor = Common::Settings::MonitorSettings{};
    }

    g_settings.monitor->menu.toolbarVisible     = g_toolbarVisible;
    g_settings.monitor->menu.lineNumbersVisible = g_colorView.IsLineNumbersEnabled();
    g_settings.monitor->menu.alwaysOnTop        = g_alwaysOnTop;
    g_settings.monitor->menu.showIds            = g_showIds;
    g_settings.monitor->menu.autoScroll         = g_colorView.GetAutoScroll();

    g_settings.monitor->filter.mask   = g_filterMask & 31u;
    g_settings.monitor->filter.preset = PresetFromLegacy(g_lastFilterPreset);

    const HRESULT saveHr = Common::Settings::SaveSettings(kAppId, g_settings);
    if (FAILED(saveHr))
    {
        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kAppId);
        DBGOUT_ERROR(L"SaveSettings failed (hr=0x{:08X}) path={}\n", static_cast<unsigned long>(saveHr), settingsPath.wstring());
    }
}

// (Re)create the toolbar with DPI-scaled bitmap and button sizes
void CreateOrRecreateToolbar(HWND hWnd)
{
    if (! hWnd)
        return;

    // If a toolbar already exists, destroy it before creating a new one
    if (g_hToolbar)
    {
        g_hToolbar.reset();
    }
    if (g_toolbarImageList)
    {
        g_toolbarImageList.reset();
    }

    // Base logical size for icons (in DIPs)
    constexpr int kIconDip = 24; // 24-DIP toolbar icons
    const int cxBitmap     = DipsToPx(hWnd, kIconDip);
    const int cyBitmap     = DipsToPx(hWnd, kIconDip);

    g_hToolbar.reset(CreateWindowExW(0,
                                     TOOLBARCLASSNAMEW,
                                     nullptr,
                                     WS_CHILD | WS_VISIBLE | TBSTYLE_TOOLTIPS | TBSTYLE_FLAT | CCS_TOP | CCS_NODIVIDER,
                                     0,
                                     0,
                                     0,
                                     0,
                                     hWnd,
                                     nullptr,
                                     g_hInstance,
                                     nullptr));

    if (! g_hToolbar)
        return;

    SendMessage(g_hToolbar.get(), TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);

    // Configure bitmap and button sizes based on DPI
    SendMessage(g_hToolbar.get(), TB_SETBITMAPSIZE, 0, MAKELPARAM(cxBitmap, cyBitmap));
    const int padding = DipsToPx(hWnd, 8);
    SendMessage(g_hToolbar.get(), TB_SETBUTTONSIZE, 0, MAKELPARAM(cxBitmap + padding, cyBitmap + padding));

    // Build a per-DPI image list from the PNG strip using WIC for high quality scaling
    // PNG strip layout is 6 icons, each 64x64 with 8 px spacing
    constexpr int kSrcIcon = 64;
    constexpr int kSrcGap  = 8;
    constexpr int kCount   = 6;

    // Ensure COM is initialized for WIC
    bool coinit   = SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED));
    auto coUninit = wil::scope_exit(
        [&]() noexcept
        {
            if (coinit)
                CoUninitialize();
        });

    wil::com_ptr<IWICImagingFactory> factory;
    if (SUCCEEDED(CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.addressof()))))
    {
        HRSRC hRes = FindResourceW(g_hInstance, MAKEINTRESOURCE(IDB_TOOLBAR_PNG), L"PNG");
        if (hRes)
        {
            HGLOBAL hData = LoadResource(g_hInstance, hRes);
            if (hData)
            {
                const void* pData  = LockResource(hData);
                const DWORD cbData = SizeofResource(g_hInstance, hRes);

                wil::com_ptr<IWICStream> stream;
                if (SUCCEEDED(factory->CreateStream(stream.addressof())) &&
                    SUCCEEDED(stream->InitializeFromMemory(reinterpret_cast<BYTE*>(const_cast<void*>(pData)), cbData)))
                {
                    wil::com_ptr<IWICBitmapDecoder> decoder;
                    if (SUCCEEDED(factory->CreateDecoderFromStream(stream.get(), nullptr, WICDecodeMetadataCacheOnLoad, decoder.addressof())))
                    {
                        wil::com_ptr<IWICBitmapFrameDecode> frame;
                        if (SUCCEEDED(decoder->GetFrame(0, frame.addressof())))
                        {
                            wil::com_ptr<IWICFormatConverter> converter;
                            if (SUCCEEDED(factory->CreateFormatConverter(converter.addressof())) &&
                                SUCCEEDED(converter->Initialize(
                                    frame.get(), GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, nullptr, 0.0f, WICBitmapPaletteTypeCustom)))
                            {
                                // Create the image list (32bpp with alpha)
                                g_toolbarImageList.reset(ImageList_Create(cxBitmap, cyBitmap, ILC_COLOR32, kCount, 1));
                                if (g_toolbarImageList)
                                {
                                    ImageList_SetBkColor(g_toolbarImageList.get(), CLR_NONE);

                                    for (int i = 0; i < kCount; ++i)
                                    {
                                        const int x = i * (kSrcIcon + kSrcGap);
                                        WICRect rc{x, 0, kSrcIcon, kSrcIcon};

                                        wil::com_ptr<IWICBitmapClipper> clip;
                                        wil::com_ptr<IWICBitmapScaler> scaler;

                                        if (SUCCEEDED(factory->CreateBitmapClipper(clip.addressof())) && SUCCEEDED(clip->Initialize(converter.get(), &rc)) &&
                                            SUCCEEDED(factory->CreateBitmapScaler(scaler.addressof())) &&
                                            SUCCEEDED(scaler->Initialize(
                                                clip.get(), static_cast<UINT>(cxBitmap), static_cast<UINT>(cyBitmap), WICBitmapInterpolationModeFant)))
                                        {
                                            const UINT stride  = static_cast<UINT>(cxBitmap) * 4u;
                                            const UINT bufSize = stride * static_cast<UINT>(cyBitmap);
                                            std::vector<BYTE> pixels(bufSize);
                                            if (SUCCEEDED(scaler->CopyPixels(nullptr, stride, bufSize, pixels.data())))
                                            {
                                                BITMAPINFO bmi{};
                                                bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
                                                bmi.bmiHeader.biWidth       = cxBitmap;
                                                bmi.bmiHeader.biHeight      = -cyBitmap; // top-down DIB
                                                bmi.bmiHeader.biPlanes      = 1;
                                                bmi.bmiHeader.biBitCount    = 32;
                                                bmi.bmiHeader.biCompression = BI_RGB;

                                                void* bits = nullptr;
                                                wil::unique_hbitmap hbm(CreateDIBSection(nullptr, &bmi, DIB_RGB_COLORS, &bits, nullptr, 0));
                                                if (hbm && bits)
                                                {
                                                    memcpy(bits, pixels.data(), pixels.size());
                                                    ImageList_Add(g_toolbarImageList.get(), hbm.get(), nullptr);
                                                }
                                            }
                                        }
                                    }

                                    // Attach the image list to the toolbar
                                    SendMessage(g_hToolbar.get(), TB_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(g_toolbarImageList.get()));
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Define buttons referencing image indices in the image list (0-based)
    TBBUTTON buttons[] = {
        {0, IDM_FILE_NEW, TBSTATE_ENABLED, TBSTYLE_BUTTON, {0}, 0, 0},
        {1, IDM_FILE_OPEN, TBSTATE_ENABLED, TBSTYLE_BUTTON, {0}, 0, 0},
        {2, IDM_FILE_SAVE_AS, TBSTATE_ENABLED, TBSTYLE_BUTTON, {0}, 0, 0},
        {0, 0, 0, TBSTYLE_SEP, {0}, 0, 0},
        {3, IDM_EDIT_COPY, TBSTATE_ENABLED, TBSTYLE_BUTTON, {0}, 0, 0},
        {0, 0, 0, TBSTYLE_SEP, {0}, 0, 0},
        {4, IDM_OPTION_ID, TBSTATE_ENABLED, TBSTYLE_BUTTON, {0}, 0, 0},
    };

    SendMessage(g_hToolbar.get(), TB_ADDBUTTONS, static_cast<WPARAM>(sizeof(buttons) / sizeof(buttons[0])), reinterpret_cast<LPARAM>(&buttons[0]));

    // Ensure the toolbar sizes itself properly
    SendMessage(g_hToolbar.get(), TB_AUTOSIZE, 0, 0);
}

std::wstring ReadFileAsTextUTF(const std::wstring& path)
{
    std::wstring result;
    std::ifstream f(path, std::ios::binary);
    if (! f)
        return result;
    std::vector<unsigned char> bytes((std::istreambuf_iterator<char>(f)), std::istreambuf_iterator<char>());
    if (bytes.empty())
        return result;

    // BOM detection: UTF-16LE, UTF-8
    if (bytes.size() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
    {
        // UTF-16 LE with BOM
        size_t wcharCount = (bytes.size() - 2) / 2;
        result.resize(wcharCount);
        memcpy(result.data(), bytes.data() + 2, wcharCount * sizeof(wchar_t));
        return result;
    }
    if (bytes.size() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
    {
        // UTF-8 BOM
        int wlen = MultiByteToWideChar(CP_UTF8, 0, reinterpret_cast<LPCCH>(bytes.data() + 3), static_cast<int>(bytes.size() - 3), nullptr, 0);
        if (wlen > 0)
        {
            result.resize(static_cast<size_t>(wlen));
            MultiByteToWideChar(CP_UTF8, 0, reinterpret_cast<LPCCH>(bytes.data() + 3), static_cast<int>(bytes.size() - 3), result.data(), wlen);
        }
        return result;
    }
    // Assume UTF-8 without BOM
    int wlen = MultiByteToWideChar(CP_UTF8, 0, reinterpret_cast<LPCCH>(bytes.data()), static_cast<int>(bytes.size()), nullptr, 0);
    if (wlen > 0)
    {
        result.resize(static_cast<size_t>(wlen));
        MultiByteToWideChar(CP_UTF8, 0, reinterpret_cast<LPCCH>(bytes.data()), static_cast<int>(bytes.size()), result.data(), wlen);
    }
    return result;
}

bool DoFileOpen(HWND owner)
{
    wchar_t file[MAX_PATH]        = L"";
    const std::wstring filter     = LoadStringResource(g_hInstance, IDS_FILE_FILTER_OPEN);
    const std::wstring defaultExt = LoadStringResource(g_hInstance, IDS_FILE_DEFAULT_EXT_TXT);

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = owner;
    ofn.lpstrFilter = filter.empty() ? nullptr : filter.c_str();
    ofn.lpstrFile   = file;
    ofn.nMaxFile    = static_cast<DWORD>(std::size(file));
    ofn.Flags       = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_EXPLORER;
    ofn.lpstrDefExt = defaultExt.empty() ? nullptr : defaultExt.c_str();
    if (! GetOpenFileNameW(&ofn))
        return false;

    auto text = ReadFileAsTextUTF(file);
    if (text.empty())
    {
        const std::wstring caption = LoadStringResource(g_hInstance, IDS_CAPTION_OPEN);
        const std::wstring message = LoadStringResource(g_hInstance, IDS_MSG_OPEN_FAILED_READ);
        ShowModalMessageDialog(g_hInstance, owner, caption.c_str(), message.c_str());
        return false;
    }
    g_colorView.ClearColoring();
    g_colorView.SetText(text);
    return true;
}

bool DoFileSaveAs(HWND owner)
{
    wchar_t file[MAX_PATH]        = L"";
    const std::wstring filter     = LoadStringResource(g_hInstance, IDS_FILE_FILTER_SAVE);
    const std::wstring defaultExt = LoadStringResource(g_hInstance, IDS_FILE_DEFAULT_EXT_TXT);

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = owner;
    ofn.lpstrFilter = filter.empty() ? nullptr : filter.c_str();
    ofn.lpstrFile   = file;
    ofn.nMaxFile    = static_cast<DWORD>(std::size(file));
    ofn.Flags       = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT | OFN_EXPLORER;
    ofn.lpstrDefExt = defaultExt.empty() ? nullptr : defaultExt.c_str();
    if (! GetSaveFileNameW(&ofn))
        return false;

    return g_colorView.SaveTextToFile(file);
}

void AdjustLayout(HWND hWnd)
{
    if (! hWnd)
        return;

    RECT clientRect{};
    GetClientRect(hWnd, &clientRect);

    int toolbarHeight         = 0;
    const bool toolbarVisible = g_hToolbar && IsWindowVisible(g_hToolbar.get());
    if (toolbarVisible)
    {
        // Ask toolbar to autosize and then query its ideal size for the current DPI
        SendMessage(g_hToolbar.get(), TB_AUTOSIZE, 0, 0);

        SIZE maxSize{};
        if (SendMessage(g_hToolbar.get(), TB_GETMAXSIZE, 0, reinterpret_cast<LPARAM>(&maxSize)))
        {
            toolbarHeight = static_cast<int>(maxSize.cy);
        }
        else
        {
            // Fallback: measure current window rect
            RECT rcTb{};
            if (GetWindowRect(g_hToolbar.get(), &rcTb))
            {
                ::MapWindowPoints(nullptr, hWnd, reinterpret_cast<LPPOINT>(&rcTb), 2);
                toolbarHeight = rcTb.bottom - rcTb.top;
            }
        }
    }

    int statusBarHeight = 0;
    if (g_hStatusBar)
    {
        // Status bar auto-sizes itself, just need to get its height
        SendMessage(g_hStatusBar.get(), WM_SIZE, 0, 0);
        RECT rcStatus{};
        if (GetWindowRect(g_hStatusBar.get(), &rcStatus))
        {
            statusBarHeight = rcStatus.bottom - rcStatus.top;
        }
    }

    const int width  = clientRect.right - clientRect.left;
    const int height = clientRect.bottom - clientRect.top;

    if (toolbarVisible)
    {
        MoveWindow(g_hToolbar.get(), 0, 0, width, toolbarHeight, TRUE);
    }

    const int contentHeight = std::max(0, height - toolbarHeight - statusBarHeight);

    if (g_hColorView)
    {
        MoveWindow(g_hColorView.get(), 0, toolbarHeight, width, contentHeight, TRUE);
    }

    if (g_hStatusBar)
    {
        MoveWindow(g_hStatusBar.get(), 0, toolbarHeight + contentHeight, width, statusBarHeight, TRUE);
    }
}

// Create status bar with 5 parts: auto-scroll, filter state, visible lines, total lines, ETW stats
void CreateStatusBar(HWND hWnd)
{
    g_hStatusBar.reset(CreateWindowExW(0, STATUSCLASSNAMEW, nullptr, WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0, hWnd, nullptr, g_hInstance, nullptr));
    if (g_hStatusBar)
    {
        // Define 5 parts: auto-scroll (90px), filter (220px), visible lines (150px), total lines (150px), ETW stats (remaining)
        constexpr int kAutoWidth    = 90;
        constexpr int kFilterWidth  = 220;
        constexpr int kVisibleWidth = 150;
        constexpr int kTotalWidth   = 150;

        RECT rcClient{};
        GetClientRect(hWnd, &rcClient);

        const int parts[] = {
            kAutoWidth,
            kAutoWidth + kFilterWidth,
            kAutoWidth + kFilterWidth + kVisibleWidth,
            kAutoWidth + kFilterWidth + kVisibleWidth + kTotalWidth,
            -1 // Remaining width
        };

        SendMessage(g_hStatusBar.get(), SB_SETPARTS, 5, reinterpret_cast<LPARAM>(parts));
    }
}

// Update status bar with current statistics and synchronize auto-scroll menu state
void UpdateStatusBar()
{
    if (! g_hStatusBar)
        return;

    // Synchronize auto-scroll menu checkmark with ColorTextView's auto-scroll state
    // ColorTextView manages its own auto-scroll state and we query it here
    bool isAutoScrollEnabled = g_colorView.GetAutoScroll();

    HWND hMainWnd = GetParent(g_hStatusBar.get());
    if (hMainWnd)
    {
        HMENU hMenu = GetMenu(hMainWnd);
        if (hMenu)
        {
            // Always sync menu to match actual auto-scroll state (no static variable needed)
            UINT checkState = isAutoScrollEnabled ? MF_CHECKED : MF_UNCHECKED;
            CheckMenuItem(hMenu, IDM_OPTION_AUTO_SCROLL, MF_BYCOMMAND | checkState);
        }
    }

    // Get line counts from ColorTextView
    const size_t visibleLines = g_colorView.GetVisibleLineCount();
    const size_t totalLines   = g_colorView.GetTotalLineCount();

    // Get ETW stats (listener-side, not writer-side)
    uint64_t etwReceived = 0;
    if (g_etwListener)
    {
        const auto s = g_etwListener->GetStatistics();
        etwReceived  = static_cast<uint64_t>(s.eventsProcessed);
    }

    // Format status bar text with specific filter names
    std::wstring filterText;
    const uint32_t filterMask = g_filterMask & Debug::InfoParam::Type::All;
    if (filterMask == Debug::InfoParam::Type::All)
    {
        filterText = LoadStringResource(g_hInstance, IDS_STATUS_FILTER_ALL);
    }
    else
    {
        // Build filter text showing enabled types
        std::vector<std::wstring> enabledTypeNames;
        if (filterMask & 0x01)
            enabledTypeNames.push_back(LoadStringResource(g_hInstance, IDS_FILTER_TYPE_TEXT));
        if (filterMask & 0x02)
            enabledTypeNames.push_back(LoadStringResource(g_hInstance, IDS_FILTER_TYPE_ERROR));
        if (filterMask & 0x04)
            enabledTypeNames.push_back(LoadStringResource(g_hInstance, IDS_FILTER_TYPE_WARNING));
        if (filterMask & 0x08)
            enabledTypeNames.push_back(LoadStringResource(g_hInstance, IDS_FILTER_TYPE_INFO));
        if (filterMask & 0x10)
            enabledTypeNames.push_back(LoadStringResource(g_hInstance, IDS_FILTER_TYPE_DEBUG));

        if (enabledTypeNames.empty())
        {
            filterText = LoadStringResource(g_hInstance, IDS_STATUS_FILTER_NONE);
        }
        else if (enabledTypeNames.size() == 1)
        {
            filterText = FormatStringResource(g_hInstance, IDS_STATUS_FILTER_ONE_FMT, enabledTypeNames[0]);
        }
        else
        {
            // Join with "+" separator for multiple types
            filterText                = LoadStringResource(g_hInstance, IDS_STATUS_FILTER_MULTI_PREFIX);
            const std::wstring joiner = LoadStringResource(g_hInstance, IDS_STATUS_FILTER_MULTI_JOINER);
            for (size_t i = 0; i < enabledTypeNames.size(); ++i)
            {
                if (i > 0)
                    filterText += joiner;
                filterText += enabledTypeNames[i];
            }
        }
    }

    const std::wstring autoText    = LoadStringResource(g_hInstance, isAutoScrollEnabled ? IDS_STATUS_AUTOSCROLL_ON : IDS_STATUS_AUTOSCROLL_OFF);
    const std::wstring visibleText = FormatStringResource(g_hInstance, IDS_STATUS_VISIBLE_FMT, visibleLines);
    const std::wstring totalText   = FormatStringResource(g_hInstance, IDS_STATUS_TOTAL_FMT, totalLines);
    const std::wstring etwText     = FormatStringResource(g_hInstance, IDS_STATUS_ETW_RECEIVED_FMT, etwReceived);

    SendMessageW(g_hStatusBar.get(), SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(autoText.c_str()));
    SendMessageW(g_hStatusBar.get(), SB_SETTEXTW, 1, reinterpret_cast<LPARAM>(filterText.c_str()));
    SendMessageW(g_hStatusBar.get(), SB_SETTEXTW, 2, reinterpret_cast<LPARAM>(visibleText.c_str()));
    SendMessageW(g_hStatusBar.get(), SB_SETTEXTW, 3, reinterpret_cast<LPARAM>(totalText.c_str()));
    SendMessageW(g_hStatusBar.get(), SB_SETTEXTW, 4, reinterpret_cast<LPARAM>(etwText.c_str()));

    // Track message count for adaptive refresh
    g_lastMessageCount = etwReceived;
}
} // namespace

// Forward declarations of functions included in this code module:
std::optional<HWND> InitInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK About(HWND, UINT, WPARAM, LPARAM);

// Ensure the process runs Per-Monitor (V2 if available) so GetDpiForWindow returns actual DPI
// this information is also in the manifest
// <dpiAware>true/PM</dpiAware>
// <dpiAwareness> PerMonitorV2</ dpiAwareness>
static void InitializeDpiAwareness()
{
#ifdef _DEBUG
    OutputDebugStringA("=== DPI Awareness Diagnostics ===\n");

    // Check what the thread DPI awareness context reports FIRST
    DPI_AWARENESS_CONTEXT currentContext = GetThreadDpiAwarenessContext();
    if (currentContext)
    {
        // Get the raw pointer value for comparison
        auto rawValue = reinterpret_cast<uintptr_t>(currentContext);
        auto msg      = std::format("Initial context raw value: 0x{:X}\n", rawValue);
        OutputDebugStringA(msg.c_str());

        OutputDebugStringA("Initial thread DPI awareness context: ");

        // Check against all known context values
        if (AreDpiAwarenessContextsEqual(currentContext, DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2))
            OutputDebugStringA("PER_MONITOR_AWARE_V2: Ok\n");
        else if (AreDpiAwarenessContextsEqual(currentContext, DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE))
            OutputDebugStringA("PER_MONITOR_AWARE: Ok\n");
        else if (AreDpiAwarenessContextsEqual(currentContext, DPI_AWARENESS_CONTEXT_SYSTEM_AWARE))
            OutputDebugStringA("SYSTEM_AWARE\n");
        else if (AreDpiAwarenessContextsEqual(currentContext, DPI_AWARENESS_CONTEXT_UNAWARE))
            OutputDebugStringA("UNAWARE\n");
        else
        {
            // Try to get DPI awareness from the context using the conversion function
            DPI_AWARENESS awareness  = GetAwarenessFromDpiAwarenessContext(currentContext);
            const char* awarenessStr = "Unknown";
            switch (awareness)
            {
                case DPI_AWARENESS_UNAWARE: awarenessStr = "UNAWARE"; break;
                case DPI_AWARENESS_SYSTEM_AWARE: awarenessStr = "SYSTEM_AWARE"; break;
                case DPI_AWARENESS_PER_MONITOR_AWARE: awarenessStr = "PER_MONITOR_AWARE (likely V2 from manifest)"; break;
                case DPI_AWARENESS_INVALID:
                default: awarenessStr = "Invalid"; break;
            }
            auto awarenessMsg = std::format("Decoded awareness: {}: Ok\n", awarenessStr);
            OutputDebugStringA(awarenessMsg.c_str());
        }

        // Also show what the constants are for comparison
        auto v2Value      = reinterpret_cast<intptr_t>(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        auto v1Value      = reinterpret_cast<intptr_t>(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE);
        auto systemValue  = reinterpret_cast<intptr_t>(DPI_AWARENESS_CONTEXT_SYSTEM_AWARE);
        auto unawareValue = reinterpret_cast<intptr_t>(DPI_AWARENESS_CONTEXT_UNAWARE);

        auto constantsMsg = std::format("Constants: V2={}, V1={}, SYS={}, UN={}\n", v2Value, v1Value, systemValue, unawareValue);
        OutputDebugStringA(constantsMsg.c_str());
    }

    OutputDebugStringA("Attempting to set DPI awareness programmatically\n");
#endif

    // Try the modern API first (Windows 10 1703+)
    if (SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2))
    {
#ifdef _DEBUG
        OutputDebugStringA("Successfully set DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2\n");
#endif
        return;
    }

    auto err = GetLastError();
    if (err == ERROR_ACCESS_DENIED)
    {
#ifdef _DEBUG
        OutputDebugStringA("DPI awareness already set (ACCESS_DENIED) - manifest is working!\n");

        // The legacy API may not report correctly after windows are created
        PROCESS_DPI_AWARENESS awareness = PROCESS_DPI_UNAWARE;
        auto hr                         = GetProcessDpiAwareness(GetCurrentProcess(), &awareness);
        if (SUCCEEDED(hr))
        {
            const char* awarenessStr = "Unknown";
            switch (awareness)
            {
                case PROCESS_DPI_UNAWARE: awarenessStr = "UNAWARE (legacy API - may be inaccurate)"; break;
                case PROCESS_SYSTEM_DPI_AWARE: awarenessStr = "SYSTEM_AWARE"; break;
                case PROCESS_PER_MONITOR_DPI_AWARE: awarenessStr = "PER_MONITOR_AWARE"; break;
            }
            auto msg = std::format("Legacy API reports: {}\n", awarenessStr);
            OutputDebugStringA(msg.c_str());
        }

        // Use the newer, more accurate API to decode the actual awareness level
        DPI_AWARENESS_CONTEXT context = GetThreadDpiAwarenessContext();
        if (context)
        {
            DPI_AWARENESS contextAwareness = GetAwarenessFromDpiAwarenessContext(context);
            OutputDebugStringA("Actual DPI awareness: ");
            switch (contextAwareness)
            {
                case DPI_AWARENESS_UNAWARE: OutputDebugStringA("UNAWARE (Something is wrong)\n"); break;
                case DPI_AWARENESS_SYSTEM_AWARE: OutputDebugStringA("SYSTEM_AWARE\n"); break;
                case DPI_AWARENESS_PER_MONITOR_AWARE: OutputDebugStringA("PER_MONITOR_AWARE (V1 or V2 - Manifest working!)\n"); break;
                case DPI_AWARENESS_INVALID: OutputDebugStringA("INVALID\n"); break;
                default: OutputDebugStringA("Invalid/Unknown\n"); break;
            }
        }

        OutputDebugStringA("=== End DPI Diagnostics ===\n");
#endif
        return; // Success - DPI awareness is already set
    }

    // ... rest of fallback code remains the same
#ifdef _DEBUG
    auto errorMsg = std::format("SetProcessDpiAwarenessContext V2 failed: {}\n", err);
    OutputDebugStringA(errorMsg.c_str());
#endif

    // Fallback to Per-Monitor V1 (Windows 8.1+)
    if (SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE))
    {
#ifdef _DEBUG
        OutputDebugStringA("Successfully set DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE\n");
#endif
        return;
    }

    err = GetLastError();
    if (err == ERROR_ACCESS_DENIED)
    {
#ifdef _DEBUG
        OutputDebugStringA("DPI awareness already set (V1 also ACCESS_DENIED)\n");
#endif
        return; // Success - DPI awareness is already set
    }

#ifdef _DEBUG
    auto errorMsg2 = std::format("SetProcessDpiAwarenessContext V1 failed: {}\n", err);
    OutputDebugStringA(errorMsg2.c_str());
#endif

    // Final fallback to legacy API (Windows Vista+)
    if (SetProcessDPIAware())
    {
#ifdef _DEBUG
        OutputDebugStringA("Successfully set legacy DPI awareness\n");
#endif
        return;
    }

#ifdef _DEBUG
    OutputDebugStringA("All DPI awareness methods failed - this is unexpected\n");
#endif
}

// Separate function with C++ objects (cannot use __try/__except)
static int RunApplication(HINSTANCE hInstance, int nCmdShow)
{
    constexpr wchar_t kInstanceMutexName[] = L"Local\\RedSalamanderMonitor_Instance";

    constexpr std::wstring_view kWaitInstanceArg = L"--wait-instance";
    constexpr ULONGLONG kWaitInstanceTimeoutMs   = 5000;
    constexpr DWORD kWaitInstancePollMs          = 50;

    wil::unique_handle instanceMutex;
    DWORD mutexCreationError = ERROR_SUCCESS;

    const auto tryCreateInstanceMutex = [&]() -> bool
    {
        instanceMutex.reset(::CreateMutexW(nullptr, FALSE, kInstanceMutexName));
        mutexCreationError = ::GetLastError();
        return static_cast<bool>(instanceMutex);
    };

    if (! tryCreateInstanceMutex())
    {
        const std::wstring caption = LoadStringResource(hInstance, IDS_APP_TITLE);
        const std::wstring message = LoadStringResource(hInstance, IDS_MSG_INSTANCE_GUARD_FAILED);
        ShowModalMessageDialog(hInstance, nullptr, caption.c_str(), message.c_str());
        return FALSE;
    }

    if (mutexCreationError == ERROR_ALREADY_EXISTS)
    {
        if (! HasCommandLineArg(kWaitInstanceArg))
        {
            OutputDebugStringW(L"Red Salamander Monitor is already running.");
            return FALSE;
        }

        instanceMutex.reset();

        const ULONGLONG startTick = ::GetTickCount64();
        while (true)
        {
            ::Sleep(kWaitInstancePollMs);

            if (! tryCreateInstanceMutex())
            {
                const std::wstring caption = LoadStringResource(hInstance, IDS_APP_TITLE);
                const std::wstring message = LoadStringResource(hInstance, IDS_MSG_INSTANCE_GUARD_FAILED);
                ShowModalMessageDialog(hInstance, nullptr, caption.c_str(), message.c_str());
                return FALSE;
            }

            if (mutexCreationError != ERROR_ALREADY_EXISTS)
            {
                break;
            }

            instanceMutex.reset();

            if ((::GetTickCount64() - startTick) > kWaitInstanceTimeoutMs)
            {
                OutputDebugStringW(L"Timed out waiting for previous instance to exit.");
                return FALSE;
            }
        }
    }

    // Set DPI awareness before creating any windows
    InitializeDpiAwareness();

    g_config.Load();

    const std::filesystem::path themesDir = GetThemesDirectory();
    if (! themesDir.empty())
    {
        Common::Settings::LoadThemeDefinitionsFromDirectory(themesDir, g_fileThemes);
    }

    Common::Settings::LoadSettings(kAppId, g_settings);

    if (g_settings.monitor)
    {
        g_toolbarVisible     = g_settings.monitor->menu.toolbarVisible;
        g_lineNumbersVisible = g_settings.monitor->menu.lineNumbersVisible;
        g_alwaysOnTop        = g_settings.monitor->menu.alwaysOnTop;
        g_showIds            = g_settings.monitor->menu.showIds;
        g_autoScrollEnabled  = g_settings.monitor->menu.autoScroll;

        g_filterMask       = g_settings.monitor->filter.mask & 31u;
        g_lastFilterPreset = LegacyFromPreset(g_settings.monitor->filter.preset);
        if (g_lastFilterPreset == -1)
        {
            const int inferred = InferLegacyPresetFromMask(g_filterMask);
            if (inferred != -1)
                g_lastFilterPreset = inferred;
        }
    }
    else
    {
        // Migration/defaults: prefer existing registry values for filter settings if present.
        g_filterMask       = g_config.filterMask;
        g_lastFilterPreset = g_config.lastFilterPreset;
        if (g_lastFilterPreset == -1)
        {
            const int inferred = InferLegacyPresetFromMask(g_filterMask);
            if (inferred != -1)
                g_lastFilterPreset = inferred;
        }

        Common::Settings::MonitorSettings monitorSettings;
        monitorSettings.menu.toolbarVisible     = g_toolbarVisible;
        monitorSettings.menu.lineNumbersVisible = g_lineNumbersVisible;
        monitorSettings.menu.alwaysOnTop        = g_alwaysOnTop;
        monitorSettings.menu.showIds            = g_showIds;
        monitorSettings.menu.autoScroll         = g_autoScrollEnabled;
        monitorSettings.filter.mask             = g_filterMask & 31u;
        monitorSettings.filter.preset           = PresetFromLegacy(g_lastFilterPreset);
        g_settings.monitor                      = std::move(monitorSettings);
    }

    // Perform application initialization:
    auto hWnd = InitInstance(hInstance, nCmdShow);
    if (! hWnd)
    {
        return FALSE;
    }

    wil::unique_haccel hAccelTable(LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_REDSALAMANDERMONITOR)));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (! TranslateAccelerator(hWnd.value(), hAccelTable.get(), &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return static_cast<int>(msg.wParam);
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

    const auto max = std::min<std::size_t>(outMessageChars - 1, static_cast<std::size_t>(std::numeric_limits<std::ptrdiff_t>::max()));
    const auto r   = std::format_to_n(outMessage, static_cast<std::ptrdiff_t>(max), L"Fatal Exception ({0}, 0x{1:08X}).", exceptionName, static_cast<unsigned>(exceptionCode));
    const std::ptrdiff_t cap     = static_cast<std::ptrdiff_t>(max);
    const std::ptrdiff_t written = (r.size < 0) ? 0 : ((r.size > cap) ? cap : r.size);
    outMessage[(written <= 0) ? 0u : static_cast<size_t>(written)] = L'\0';
}
} // namespace

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE /*hPrevInstance*/, _In_ LPWSTR /*lpCmdLine*/, _In_ int nCmdShow)
{
    // Use SEH to catch all exceptions (no C++ objects in this scope)
    __try
    {
        return RunApplication(hInstance, nCmdShow);
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        // Handle all exceptions including SEH exceptions
        const DWORD exceptionCode    = GetExceptionCode();
        const wchar_t* exceptionName = exception::GetExceptionName(exceptionCode);

        wchar_t errorMsg[512]{};
        BuildFatalExceptionMessage(hInstance, exceptionName, exceptionCode, errorMsg, std::size(errorMsg));
        OutputDebugStringW(errorMsg);

        wchar_t caption[256]{};
        const int captionLength = LoadStringW(hInstance, IDS_FATAL_ERROR_CAPTION, caption, static_cast<int>(sizeof(caption) / sizeof(caption[0])));
        ShowModalMessageDialog(hInstance, nullptr, captionLength > 0 ? caption : L"", errorMsg);

        return -1;
    }
}

// Saves instance handle and creates main window
std::optional<HWND> InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    g_hInstance = hInstance; // Store instance handle in our global variable

    // Initialize common controls (for toolbar)
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC  = ICC_BAR_CLASSES;
    InitCommonControlsEx(&icc);

    WNDCLASSEXW wcex{};
    wcex.cbSize        = sizeof(WNDCLASSEX);
    wcex.style         = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc   = WndProc;
    wcex.cbClsExtra    = 0;
    wcex.cbWndExtra    = 0;
    wcex.hInstance     = hInstance;
    wcex.hIcon         = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_REDSALAMANDERMONITOR));
    wcex.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wcex.lpszMenuName  = MAKEINTRESOURCEW(IDC_REDSALAMANDERMONITOR);
    wcex.lpszClassName = g_redSalamanderMonitorClassName;
    wcex.hIconSm       = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    auto atom = RegisterClassExW(&wcex);
    if (! atom)
    {
        // TODO: Handle error, e.g., log it or show a message box
        return std::nullopt;
    }

    wchar_t title[256]{};
    const int titleLength = LoadStringW(hInstance, IDS_APP_TITLE, title, static_cast<int>(sizeof(title) / sizeof(title[0])));
    PCWSTR windowTitle    = titleLength > 0 ? title : g_redSalamanderMonitor;

    wil::unique_hwnd hWnd(CreateWindowExW(WS_EX_NOACTIVATE,
                                          g_redSalamanderMonitorClassName,
                                          windowTitle,
                                          WS_OVERLAPPEDWINDOW,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          640,
                                          480,
                                          nullptr,
                                          nullptr,
                                          hInstance,
                                          nullptr));
    if (! hWnd)
    {
        // TODO: Handle error, e.g., log it or show a message box
        return std::nullopt;
    }

    // Remove WS_EX_NOACTIVATE after creation
    SetWindowLongPtr(hWnd.get(), GWL_EXSTYLE, GetWindowLongPtr(hWnd.get(), GWL_EXSTYLE) & ~WS_EX_NOACTIVATE);

    int showCmd   = nCmdShow;
    const auto it = g_settings.windows.find(kWindowId);
    if (it != g_settings.windows.end())
    {
        const UINT dpi                                     = GetDpiForWindow(hWnd.get());
        const Common::Settings::WindowPlacement normalized = Common::Settings::NormalizeWindowPlacement(it->second, dpi);

        SetWindowPos(
            hWnd.get(), nullptr, normalized.bounds.x, normalized.bounds.y, normalized.bounds.width, normalized.bounds.height, SWP_NOZORDER | SWP_NOACTIVATE);

        showCmd = normalized.state == Common::Settings::WindowState::Maximized ? SW_MAXIMIZE : SW_SHOWNORMAL;
    }

    ShowWindow(hWnd.get(), showCmd);
    UpdateWindow(hWnd.get());

    return hWnd.release();
}

static std::wstring NormalizeLineEndings(const std::wstring& input)
{
    if (input.empty())
    {
        return L"\n";
    }

    std::wstring result;
    result.reserve(input.length() + 1); // Reserve space to avoid reallocations

    for (size_t i = 0; i < input.length(); ++i)
    {
        if (input[i] == L'\r')
        {
            // Check if this is \r\n sequence
            if (i + 1 < input.length() && input[i + 1] == L'\n')
            {
                result += L'\n';
                ++i; // Skip the \n as we've handled the \r\n pair
            }
            else
            {
                result += L'\n'; // Convert standalone \r to \n
            }
        }
        else
        {
            result += input[i];
        }
    }

    // Ensure text ends with exactly one \n
    while (! result.empty() && result.back() == L'\n')
    {
        result.pop_back();
    }
    result += L'\n';

    return result;
}

void AddLine(PCWSTR line)
{
    // OutputDebugStringW(std::format(L"-- {}\n", line).c_str());

    if (! g_hColorView)
        return;

    std::wstring text = NormalizeLineEndings(line);
    g_colorView.AppendText(text); // ColorTextView handles scroll-to-bottom internally
}

namespace
{
LRESULT OnCreateMainWindow(HWND hWnd)
{
    g_hColorView.reset(g_colorView.Create(hWnd, 0, 0, 0, 0));
    if (! g_hColorView)
    {
        const std::wstring caption = LoadStringResource(g_hInstance, IDS_CAPTION_ERROR);
        const std::wstring message = LoadStringResource(g_hInstance, IDS_MSG_CREATE_COLORTEXTVIEW_FAILED);
        ShowModalMessageDialog(g_hInstance, hWnd, caption.c_str(), message.c_str());
        return -1;
    }

    CreateOrRecreateToolbar(hWnd);
    CreateStatusBar(hWnd);

    const std::wstring sampleText = LoadStringResource(g_hInstance, IDS_SAMPLE_TEXT);
    g_colorView.SetText(sampleText);
    g_colorView.ColorizeWord(L"ColorTextView", D2D1::ColorF(D2D1::ColorF::Blue));
    g_colorView.ColorizeWord(L"right", D2D1::ColorF(D2D1::ColorF::Green));

    g_colorView.EnableShowIds(g_showIds);
    g_colorView.EnableLineNumbers(g_lineNumbersVisible);
    g_colorView.SetAutoScroll(g_autoScrollEnabled);
    g_colorView.SetFilterMask(g_filterMask);
    ApplyMonitorTheme();

    if (g_hToolbar)
    {
        ShowWindow(g_hToolbar.get(), g_toolbarVisible ? SW_SHOW : SW_HIDE);
    }

    if (g_alwaysOnTop)
    {
        SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    }

    HMENU hMenu = GetMenu(hWnd);
    if (hMenu)
    {
        CheckMenuItem(hMenu, IDM_VIEW_TOOLBAR, static_cast<UINT>(MF_BYCOMMAND | (g_toolbarVisible ? MF_CHECKED : MF_UNCHECKED)));
        CheckMenuItem(hMenu, IDM_OPTION_TOP, static_cast<UINT>(MF_BYCOMMAND | (g_alwaysOnTop ? MF_CHECKED : MF_UNCHECKED)));
        CheckMenuItem(hMenu, IDM_OPTION_ID, static_cast<UINT>(MF_BYCOMMAND | (g_showIds ? MF_CHECKED : MF_UNCHECKED)));
        CheckMenuItem(hMenu, IDM_VIEW_LINE_NUMBERS, static_cast<UINT>(MF_BYCOMMAND | (g_colorView.IsLineNumbersEnabled() ? MF_CHECKED : MF_UNCHECKED)));

        CheckMenuItem(hMenu, IDM_FILTER_TEXT, static_cast<UINT>(MF_BYCOMMAND | ((g_filterMask & 0x01) ? MF_CHECKED : MF_UNCHECKED)));
        CheckMenuItem(hMenu, IDM_FILTER_ERROR, static_cast<UINT>(MF_BYCOMMAND | ((g_filterMask & 0x02) ? MF_CHECKED : MF_UNCHECKED)));
        CheckMenuItem(hMenu, IDM_FILTER_WARNING, static_cast<UINT>(MF_BYCOMMAND | ((g_filterMask & 0x04) ? MF_CHECKED : MF_UNCHECKED)));
        CheckMenuItem(hMenu, IDM_FILTER_INFO, static_cast<UINT>(MF_BYCOMMAND | ((g_filterMask & 0x08) ? MF_CHECKED : MF_UNCHECKED)));
        CheckMenuItem(hMenu, IDM_FILTER_DEBUG, static_cast<UINT>(MF_BYCOMMAND | ((g_filterMask & 0x10) ? MF_CHECKED : MF_UNCHECKED)));

        CheckMenuItem(hMenu, IDM_OPTION_AUTO_SCROLL, static_cast<UINT>(MF_BYCOMMAND | (g_colorView.GetAutoScroll() ? MF_CHECKED : MF_UNCHECKED)));

        RebuildThemeMenuDynamicItems(hWnd);
        UpdateThemeMenuChecks();
    }

    AdjustLayout(hWnd);
    UpdateStatusBar();

    g_etwListener         = std::make_unique<EtwListener>();
    const bool etwStarted = g_etwListener->Start(
        [](const Debug::InfoParam& info, const std::wstring& message)
        {
            std::wstring normalizedMsg = message;
            while (! normalizedMsg.empty() && (normalizedMsg.back() == L'\n' || normalizedMsg.back() == L'\r'))
            {
                normalizedMsg.pop_back();
            }

            g_colorView.QueueEtwEvent(info, std::move(normalizedMsg));
        });

    if (! etwStarted)
    {
        constexpr std::wstring_view kWaitInstanceArg = L"--wait-instance";

        const ULONG etwErrorCode = g_etwListener ? g_etwListener->GetLastErrorCode() : ERROR_SUCCESS;
        if (etwErrorCode == ERROR_ACCESS_DENIED && ! IsProcessElevated())
        {
            const std::wstring caption = LoadStringResource(g_hInstance, IDS_CAPTION_ETW_WARNING);
            const std::wstring message = LoadStringResource(g_hInstance, IDS_MSG_ETW_ELEVATE_PROMPT);

            const int choice = MessageBoxW(hWnd, message.c_str(), caption.empty() ? L"" : caption.c_str(), MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON1);

            if (choice == IDYES && RelaunchSelfElevated(hWnd, kWaitInstanceArg))
            {
                return -1;
            }
        }

        std::wstring errorMsg = FormatStringResource(g_hInstance, IDS_FMT_ETW_START_FAILED, g_etwListener->GetLastError());
        AddLine(errorMsg.c_str());

#ifdef _DEBUG
        OutputDebugString(errorMsg.c_str());
        OutputDebugStringA("\n");
#endif
    }
    else
    {
        const std::wstring startedText = LoadStringResource(g_hInstance, IDS_MSG_ETW_STARTED);
        AddLine(startedText.empty() ? L"" : startedText.c_str());
    }

    SetTimer(hWnd, kStatusBarTimerId, kStatusBarUpdateIntervalMs, nullptr);
    return 0;
}

LRESULT OnTimerMainWindow(HWND hWnd, UINT_PTR timerId)
{
    if (timerId != kStatusBarTimerId)
    {
        return DefWindowProcW(hWnd, WM_TIMER, static_cast<WPARAM>(timerId), 0);
    }

    uint64_t currentCount = 0;
    if (g_etwListener)
    {
        const auto s = g_etwListener->GetStatistics();
        currentCount = static_cast<uint64_t>(s.eventsProcessed);
    }

    if (currentCount != g_lastMessageCount || (currentCount % 10 == 0))
    {
        UpdateStatusBar();
    }

    return 0;
}

LRESULT OnSizeMainWindow(HWND hWnd, UINT /*width*/, UINT /*height*/)
{
    AdjustLayout(hWnd);
    return 0;
}

LRESULT OnDpiChangedMainWindow(HWND hWnd, [[maybe_unused]] UINT newDpi, const RECT* suggestedRect)
{
    CreateOrRecreateToolbar(hWnd);
    if (g_hToolbar)
    {
        ShowWindow(g_hToolbar.get(), g_toolbarVisible ? SW_SHOW : SW_HIDE);
    }

    if (suggestedRect != nullptr)
    {
        SetWindowPos(hWnd,
                     nullptr,
                     suggestedRect->left,
                     suggestedRect->top,
                     suggestedRect->right - suggestedRect->left,
                     suggestedRect->bottom - suggestedRect->top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

    AdjustLayout(hWnd);
    return 0;
}

LRESULT OnSystemThemeChangedMainWindow([[maybe_unused]] HWND hWnd)
{
    LocaleFormatting::InvalidateFormatLocaleCache();
    ApplyMonitorTheme();
    UpdateThemeMenuChecks();
    return 0;
}

LRESULT OnCommandMainWindow(HWND hWnd, UINT id, UINT codeNotify, HWND hwndCtl)
{
    switch (id)
    {
        case IDM_ABOUT:
#pragma warning(suppress : 5039)
            DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, reinterpret_cast<DLGPROC>(About));
            break;
        case IDM_EXIT: DestroyWindow(hWnd); break;
        case IDM_FILE_NEW:
            g_colorView.ClearColoring();
            g_colorView.ClearText();
            break;
        case IDM_FILE_OPEN: DoFileOpen(hWnd); break;
        case IDM_FILE_SAVE_AS: DoFileSaveAs(hWnd); break;
        case IDM_EDIT_COPY: g_colorView.CopySelection(); break;
        case IDM_EDIT_FIND: g_colorView.ShowFind(); break;
        case IDM_EDIT_FIND_NEXT: g_colorView.FindNext(false); break;
        case IDM_EDIT_FIND_PREV: g_colorView.FindNext(true); break;
        case IDM_FILE_PRINT:
        {
            const std::wstring caption = LoadStringResource(g_hInstance, IDS_CAPTION_PRINT);
            const std::wstring message = LoadStringResource(g_hInstance, IDS_MSG_PRINT_NOT_IMPLEMENTED);
            ShowModalMessageDialog(g_hInstance, hWnd, caption.c_str(), message.c_str());
            break;
        }
        case IDM_VIEW_LINE_NUMBERS:
            if (g_hColorView)
            {
                const bool currentState = g_colorView.IsLineNumbersEnabled();
                g_colorView.EnableLineNumbers(! currentState);
                g_lineNumbersVisible = ! currentState;

                HMENU hMenu = GetMenu(hWnd);
                CheckMenuItem(hMenu, IDM_VIEW_LINE_NUMBERS, static_cast<UINT>(MF_BYCOMMAND | (currentState ? MF_UNCHECKED : MF_CHECKED)));
            }
            break;
        case IDM_VIEW_THEME_SYSTEM:
            g_settings.theme.currentThemeId = L"builtin/system";
            ApplyMonitorTheme();
            UpdateThemeMenuChecks();
            break;
        case IDM_VIEW_THEME_LIGHT:
            g_settings.theme.currentThemeId = L"builtin/light";
            ApplyMonitorTheme();
            UpdateThemeMenuChecks();
            break;
        case IDM_VIEW_THEME_DARK:
            g_settings.theme.currentThemeId = L"builtin/dark";
            ApplyMonitorTheme();
            UpdateThemeMenuChecks();
            break;
        case IDM_VIEW_THEME_RAINBOW:
            g_settings.theme.currentThemeId = L"builtin/rainbow";
            ApplyMonitorTheme();
            UpdateThemeMenuChecks();
            break;
        case IDM_VIEW_THEME_HIGH_CONTRAST_APP:
            g_settings.theme.currentThemeId = L"builtin/highContrast";
            ApplyMonitorTheme();
            UpdateThemeMenuChecks();
            break;
        case IDM_OPTION_TOP:
        {
            g_alwaysOnTop = ! g_alwaysOnTop;
            SetWindowPos(hWnd, g_alwaysOnTop ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, IDM_OPTION_TOP, static_cast<UINT>(MF_BYCOMMAND | (g_alwaysOnTop ? MF_CHECKED : MF_UNCHECKED)));
            break;
        }
        case IDM_OPTION_ID:
        {
            g_showIds   = ! g_showIds;
            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, IDM_OPTION_ID, static_cast<UINT>(MF_BYCOMMAND | (g_showIds ? MF_CHECKED : MF_UNCHECKED)));
            g_colorView.EnableShowIds(g_showIds);
            break;
        }
        case IDM_OPTION_AUTO_SCROLL:
        {
            // From the End-key accelerator we treat this as "go to end and follow" (force enable),
            // not as a toggle.
            const bool invokedByAccelerator = (codeNotify == 1);
            const bool newState             = invokedByAccelerator ? true : ! g_colorView.GetAutoScroll();

            if (invokedByAccelerator)
            {
                g_colorView.GoToEnd(true);
            }
            else
            {
                g_colorView.SetAutoScroll(newState);
            }

            g_autoScrollEnabled = newState;

            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, IDM_OPTION_AUTO_SCROLL, static_cast<UINT>(MF_BYCOMMAND | (newState ? MF_CHECKED : MF_UNCHECKED)));
            break;
        }
        case IDM_VIEW_TOOLBAR:
        {
            HMENU hMenu                 = GetMenu(hWnd);
            const UINT state            = GetMenuState(hMenu, IDM_VIEW_TOOLBAR, MF_BYCOMMAND);
            const bool currentlyChecked = (state & MF_CHECKED) != 0;
            if (g_hToolbar)
            {
                ShowWindow(g_hToolbar.get(), currentlyChecked ? SW_HIDE : SW_SHOW);
                SendMessageW(g_hToolbar.get(), TB_AUTOSIZE, 0, 0);
            }
            g_toolbarVisible = ! currentlyChecked;
            CheckMenuItem(hMenu, IDM_VIEW_TOOLBAR, static_cast<UINT>(MF_BYCOMMAND | (currentlyChecked ? MF_UNCHECKED : MF_CHECKED)));
            AdjustLayout(hWnd);
            break;
        }
        case IDM_FILTER_TEXT:
        case IDM_FILTER_ERROR:
        case IDM_FILTER_WARNING:
        case IDM_FILTER_INFO:
        case IDM_FILTER_DEBUG:
        {
            const int typeIndex    = static_cast<int>(id) - IDM_FILTER_TEXT;
            const uint32_t bitMask = 1u << static_cast<uint32_t>(typeIndex);

            g_filterMask ^= bitMask;

            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, id, static_cast<UINT>(MF_BYCOMMAND | ((g_filterMask & bitMask) ? MF_CHECKED : MF_UNCHECKED)));

            g_lastFilterPreset = -1;

            g_colorView.SetFilterMask(g_filterMask);
            UpdateStatusBar();
            break;
        }
        case IDM_FILTER_PRESET_ERRORS_ONLY:
        {
            g_filterMask       = 0x02;
            g_lastFilterPreset = 0;

            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, IDM_FILTER_TEXT, MF_BYCOMMAND | MF_UNCHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_ERROR, MF_BYCOMMAND | MF_CHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_WARNING, MF_BYCOMMAND | MF_UNCHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_INFO, MF_BYCOMMAND | MF_UNCHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_DEBUG, MF_BYCOMMAND | MF_UNCHECKED);

            g_colorView.SetFilterMask(g_filterMask);
            UpdateStatusBar();
            break;
        }
        case IDM_FILTER_PRESET_ERRORS_WARNINGS:
        {
            g_filterMask       = 0x06;
            g_lastFilterPreset = 1;

            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, IDM_FILTER_TEXT, MF_BYCOMMAND | MF_UNCHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_ERROR, MF_BYCOMMAND | MF_CHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_WARNING, MF_BYCOMMAND | MF_CHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_INFO, MF_BYCOMMAND | MF_UNCHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_DEBUG, MF_BYCOMMAND | MF_UNCHECKED);

            g_colorView.SetFilterMask(g_filterMask);
            UpdateStatusBar();
            break;
        }
        case IDM_FILTER_PRESET_ERRORS_DEBUG:
        {
            g_filterMask       = 0x12;
            g_lastFilterPreset = 3;

            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, IDM_FILTER_TEXT, MF_BYCOMMAND | MF_UNCHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_ERROR, MF_BYCOMMAND | MF_CHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_WARNING, MF_BYCOMMAND | MF_UNCHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_INFO, MF_BYCOMMAND | MF_UNCHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_DEBUG, MF_BYCOMMAND | MF_CHECKED);

            g_colorView.SetFilterMask(g_filterMask);
            UpdateStatusBar();
            break;
        }
        case IDM_FILTER_PRESET_ALL:
        {
            g_filterMask       = Debug::InfoParam::Type::All;
            g_lastFilterPreset = 2;

            HMENU hMenu = GetMenu(hWnd);
            CheckMenuItem(hMenu, IDM_FILTER_TEXT, MF_BYCOMMAND | MF_CHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_ERROR, MF_BYCOMMAND | MF_CHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_WARNING, MF_BYCOMMAND | MF_CHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_INFO, MF_BYCOMMAND | MF_CHECKED);
            CheckMenuItem(hMenu, IDM_FILTER_DEBUG, MF_BYCOMMAND | MF_CHECKED);

            g_colorView.SetFilterMask(g_filterMask);
            UpdateStatusBar();
            break;
        }
        default:
        {
            const auto it = g_customThemeMenuIdToThemeId.find(id);
            if (it != g_customThemeMenuIdToThemeId.end())
            {
                g_settings.theme.currentThemeId = it->second;
                ApplyMonitorTheme();
                UpdateThemeMenuChecks();
                break;
            }

            return DefWindowProcW(hWnd, WM_COMMAND, MAKEWPARAM(static_cast<WORD>(id), static_cast<WORD>(codeNotify)), reinterpret_cast<LPARAM>(hwndCtl));
        }
    }

    return 0;
}

LRESULT OnPaintMainWindow(HWND hWnd)
{
    wil::unique_hdc_paint paint_dc = wil::BeginPaint(hWnd);
    return 0;
}

LRESULT OnDestroyMainWindow(HWND hWnd)
{
    SaveMonitorSettings(hWnd);
    KillTimer(hWnd, kStatusBarTimerId);

    // IMPORTANT: Shutdown order matters for thread safety.
    // 1. Stop ETW listener first (stops worker thread that calls QueueEtwEvent on g_colorView)
    // 2. Then destroy the color view (safe because no more cross-thread PostMessage calls)
    // Reversing this order risks use-after-free: worker thread could PostMessage to destroyed HWND.
    if (g_etwListener)
    {
        g_etwListener->Stop();
        g_etwListener.reset();
    }

    g_hColorView.reset();
    if (g_hToolbar)
    {
        g_hToolbar.reset();
    }
    if (g_toolbarImageList)
    {
        g_toolbarImageList.reset();
    }

    PostQuitMessage(0);
    return 0;
}
} // namespace

// Processes messages for the main window.
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
        case WM_CREATE: return OnCreateMainWindow(hWnd);
        case WM_TIMER: return OnTimerMainWindow(hWnd, static_cast<UINT_PTR>(wParam));
        case WM_SIZE: return OnSizeMainWindow(hWnd, LOWORD(lParam), HIWORD(lParam));
        case WM_DPICHANGED: return OnDpiChangedMainWindow(hWnd, static_cast<UINT>(HIWORD(wParam)), reinterpret_cast<const RECT*>(lParam));
        case WM_SETTINGCHANGE:
        case WM_THEMECHANGED:
        case WM_DWMCOLORIZATIONCOLORCHANGED:
        case WM_SYSCOLORCHANGE: return OnSystemThemeChangedMainWindow(hWnd);
        case WM_COMMAND: return OnCommandMainWindow(hWnd, LOWORD(wParam), HIWORD(wParam), reinterpret_cast<HWND>(lParam));
        case WM_PAINT: return OnPaintMainWindow(hWnd);
        case WM_DESTROY: return OnDestroyMainWindow(hWnd);
        default: return DefWindowProc(hWnd, message, wParam, lParam);
    }
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, [[maybe_unused]] LPARAM lParam)
{
    switch (message)
    {
        case WM_INITDIALOG: return static_cast<INT_PTR>(TRUE);

        case WM_COMMAND:
            if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
            {
                EndDialog(hDlg, LOWORD(wParam));
                return static_cast<INT_PTR>(TRUE);
            }
            break;
    }
    return static_cast<INT_PTR>(FALSE);
}
