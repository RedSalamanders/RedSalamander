#include "Framework.h"

#include "Preferences.Advanced.h"
#include "Preferences.Dialog.h"
#include "Preferences.Editors.h"
#include "Preferences.General.h"
#include "Preferences.Internal.h"
#include "Preferences.Keyboard.h"
#include "Preferences.Mouse.h"
#include "Preferences.Panes.h"
#include "Preferences.Plugins.h"
#include "Preferences.Themes.h"
#include "Preferences.Viewers.h"

#include <algorithm>
#include <array>
#include <charconv>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cwchar>
#include <cwctype>
#include <filesystem>
#include <format>
#include <functional>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <tuple>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include <commctrl.h>
#include <uxtheme.h>
#include <windowsx.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include "CommandRegistry.h"
#include "Helpers.h"
#include "HostServices.h"
#include "SettingsSave.h"
#include "SettingsSchemaExport.h"
#include "ShortcutDefaults.h"
#include "ShortcutManager.h"
#include "ShortcutText.h"
#include "ThemedControls.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "resource.h"

namespace
{
struct PreferencesDialogHost final : PreferencesDialogState
{
    PreferencesDialogHost()                                        = default;
    PreferencesDialogHost(const PreferencesDialogHost&)            = delete;
    PreferencesDialogHost& operator=(const PreferencesDialogHost&) = delete;

    GeneralPane _generalPane;
    PanesPane _panesPane;
    ViewersPane _viewersPane;
    EditorsPane _editorsPane;
    KeyboardPane _keyboardPane;
    MousePane _mousePane;
    ThemesPane _themesPane;
    PluginsPane _pluginsPane;
    AdvancedPane _advancedPane;
};

constexpr UINT_PTR kPrefsWheelRouteSubclassId   = 2u;
constexpr wchar_t kPrefsPageHostClassName[]     = L"RedSalamanderPrefsPageHost";
constexpr wchar_t kPreferencesWindowId[]        = L"PreferencesWindow";

[[nodiscard]] HWND GetActivePrefsPaneWindow(const PreferencesDialogState& state) noexcept
{
    const auto& hostState = static_cast<const PreferencesDialogHost&>(state);

    switch (state.currentCategory)
    {
        case PrefCategory::General: return hostState._generalPane.Hwnd();
        case PrefCategory::Panes: return hostState._panesPane.Hwnd();
        case PrefCategory::Viewers: return hostState._viewersPane.Hwnd();
        case PrefCategory::Editors: return hostState._editorsPane.Hwnd();
        case PrefCategory::Keyboard: return hostState._keyboardPane.Hwnd();
        case PrefCategory::Mouse: return hostState._mousePane.Hwnd();
        case PrefCategory::Themes: return hostState._themesPane.Hwnd();
        case PrefCategory::Plugins: return hostState._pluginsPane.Hwnd();
        case PrefCategory::Advanced: return hostState._advancedPane.Hwnd();
        default: return nullptr;
    }
}

[[nodiscard]] bool EnsurePrefsPageHostClassRegistered() noexcept
{
    const HINSTANCE instance = GetModuleHandleW(nullptr);
    WNDCLASSEXW existing{};
    existing.cbSize = sizeof(existing);
    if (GetClassInfoExW(instance, kPrefsPageHostClassName, &existing) != 0)
    {
        return true;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_DBLCLKS;
    wc.lpfnWndProc   = DefWindowProcW;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.lpszClassName = kPrefsPageHostClassName;

    const ATOM atom = RegisterClassExW(&wc);
    return atom != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

[[nodiscard]] HWND FindWheelTargetFromPoint(HWND root, POINT ptScreen) noexcept
{
    if (! root)
    {
        return nullptr;
    }

    HWND target = WindowFromPoint(ptScreen);
    if (! target || GetAncestor(target, GA_ROOT) != root)
    {
        return nullptr;
    }

    while (target && target != root)
    {
        const LONG_PTR style = GetWindowLongPtrW(target, GWL_STYLE);
        if ((style & WS_VSCROLL) != 0)
        {
            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_RANGE | SIF_PAGE;
            if (GetScrollInfo(target, SB_VERT, &si))
            {
                const int range = std::max(0, (si.nMax - si.nMin) + 1);
                if (range <= static_cast<int>(si.nPage))
                {
                    target = GetParent(target);
                    continue;
                }
            }

            std::array<wchar_t, 16> className{};
            const int len = GetClassNameW(target, className.data(), static_cast<int>(className.size()));
            if (len > 0 && _wcsicmp(className.data(), L"ComboBox") == 0)
            {
                if (SendMessageW(target, CB_GETDROPPEDSTATE, 0, 0) == 0)
                {
                    target = GetParent(target);
                    continue;
                }
            }
            return target;
        }
        target = GetParent(target);
    }

    return nullptr;
}

[[nodiscard]] bool HandlePageHostMouseWheel(HWND host, PreferencesDialogState& state, WPARAM wp) noexcept
{
    if (! host || state.pageScrollMaxY <= 0)
    {
        return false;
    }

    const int delta = GET_WHEEL_DELTA_WPARAM(wp);
    if (delta == 0)
    {
        return true;
    }

    state.pageWheelDeltaRemainder += delta;
    const int steps = state.pageWheelDeltaRemainder / WHEEL_DELTA;
    if (steps == 0)
    {
        return true;
    }
    state.pageWheelDeltaRemainder -= steps * WHEEL_DELTA;

    UINT linesPerNotch = 3;
    SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &linesPerNotch, 0);
    if (linesPerNotch == 0)
    {
        return true;
    }

    const UINT dpi     = GetDpiForWindow(host);
    const int lineStep = std::max(1, ThemedControls::ScaleDip(dpi, 24));

    int scrollDelta = 0;
    if (linesPerNotch == WHEEL_PAGESCROLL)
    {
        SCROLLINFO si{};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_PAGE;
        GetScrollInfo(host, SB_VERT, &si);
        scrollDelta = steps * static_cast<int>(si.nPage);
    }
    else
    {
        scrollDelta = steps * lineStep * static_cast<int>(linesPerNotch);
    }

    const int newPos = state.pageScrollY - scrollDelta;
    PrefsPaneHost::ScrollTo(host, state, newPos);
    return true;
}

[[nodiscard]] int ColorLuma(COLORREF c) noexcept
{
    // Approximate ITU-R BT.601 luma in 0..255.
    const int r = static_cast<int>(GetRValue(c));
    const int g = static_cast<int>(GetGValue(c));
    const int b = static_cast<int>(GetBValue(c));
    return (299 * r + 587 * g + 114 * b) / 1000;
}

[[nodiscard]] COLORREF GetDisabledTextColor(const PreferencesDialogState& state, COLORREF background) noexcept
{
    COLORREF candidate = state.theme.menu.disabledText;
    if (state.theme.highContrast)
    {
        return candidate;
    }

    const COLORREF normal   = state.theme.menu.text;
    const int minBgDiff     = 80;
    const int minNormalDiff = 36;

    const auto isReadableAndDim = [&](COLORREF color) noexcept
    {
        const int bgDiff     = std::abs(ColorLuma(color) - ColorLuma(background));
        const int normalDiff = std::abs(ColorLuma(color) - ColorLuma(normal));
        return bgDiff >= minBgDiff && normalDiff >= minNormalDiff;
    };

    auto blended = ThemedControls::BlendColor(background, normal, state.theme.dark ? 140 : 90, 255);
    if (std::abs(ColorLuma(blended) - ColorLuma(background)) < minBgDiff)
    {
        blended = ThemedControls::BlendColor(background, normal, state.theme.dark ? 170 : 120, 255);
    }

    if (isReadableAndDim(candidate))
    {
        const int candNormalDiff  = std::abs(ColorLuma(candidate) - ColorLuma(normal));
        const int blendNormalDiff = std::abs(ColorLuma(blended) - ColorLuma(normal));
        if (candNormalDiff >= blendNormalDiff)
        {
            return candidate;
        }
    }

    return blended;
}

[[nodiscard]] const PrefsPluginConfigFieldControls* FindPluginDetailsToggleControls(const PreferencesDialogState& state, HWND toggle) noexcept
{
    if (! toggle)
    {
        return nullptr;
    }

    for (const PrefsPluginConfigFieldControls& controls : state.pluginsDetailsConfigFields)
    {
        if (controls.toggle.get() == toggle)
        {
            return &controls;
        }
    }

    return nullptr;
}

LRESULT CALLBACK PreferencesWheelRouteSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR uIdSubclass, DWORD_PTR refData) noexcept
{
    auto* state = reinterpret_cast<PreferencesDialogState*>(refData);
    if (! state)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_MOUSEWHEEL:
        {
            const HWND dlg = GetAncestor(hwnd, GA_ROOT);
            if (! dlg)
            {
                return 0;
            }

            POINT ptScreen{};
            ptScreen.x = GET_X_LPARAM(lp);
            ptScreen.y = GET_Y_LPARAM(lp);

            HWND target = FindWheelTargetFromPoint(dlg, ptScreen);
            if (! target)
            {
                // Don't scroll the dialog when the user is wheeling outside it.
                return 0;
            }

            if (target == hwnd)
            {
                break;
            }

            SendMessageW(target, msg, wp, lp);
            return 0;
        }
        case WM_NCDESTROY:
        {
            RemoveWindowSubclass(hwnd, PreferencesWheelRouteSubclassProc, uIdSubclass);
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

void InstallWheelRoutingSubclasses(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! dlg)
    {
        return;
    }

    const auto setSubclass = [](HWND hwnd, LPARAM lParam) noexcept -> BOOL
    {
        auto* state = reinterpret_cast<PreferencesDialogState*>(lParam);
        if (! state)
        {
            return TRUE;
        }

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowSubclass(hwnd, PreferencesWheelRouteSubclassProc, kPrefsWheelRouteSubclassId, reinterpret_cast<DWORD_PTR>(state));
#pragma warning(pop)
        return TRUE;
    };

    static_cast<void>(setSubclass(dlg, reinterpret_cast<LPARAM>(&state)));
    EnumChildWindows(dlg, setSubclass, reinterpret_cast<LPARAM>(&state));
}

void PaintPageHostBackgroundAndCards(HDC hdc, HWND host, const PreferencesDialogState& state) noexcept
{
    if (! hdc || ! host)
    {
        return;
    }

    RECT rc{};
    GetClientRect(host, &rc);

    HBRUSH brush = state.backgroundBrush ? state.backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(NULL_BRUSH));
    FillRect(hdc, &rc, brush);

    if (state.theme.systemHighContrast || state.pageSettingCards.empty())
    {
        return;
    }

    const UINT dpi         = GetDpiForWindow(host);
    const int radius       = ThemedControls::ScaleDip(dpi, 6);
    const COLORREF surface = ThemedControls::GetControlSurfaceColor(state.theme);
    const COLORREF border  = ThemedControls::BlendColor(surface, state.theme.menu.text, state.theme.dark ? 40 : 30, 255);

    wil::unique_hbrush cardBrush(CreateSolidBrush(surface));
    wil::unique_hpen cardPen(CreatePen(PS_SOLID, 1, border));
    if (! cardBrush || ! cardPen)
    {
        return;
    }

    [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc, cardBrush.get());
    [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc, cardPen.get());

    for (const RECT& baseCard : state.pageSettingCards)
    {
        RECT card = baseCard;
        OffsetRect(&card, 0, -state.pageScrollY);
        if (card.right <= card.left || card.bottom <= card.top)
        {
            continue;
        }
        if (card.bottom <= rc.top || card.top >= rc.bottom)
        {
            continue;
        }
        RoundRect(hdc, card.left, card.top, card.right, card.bottom, radius, radius);
    }
}

struct CategoryInfo
{
    PrefCategory id{};
    UINT labelId{};
    UINT descriptionId{};
};

constexpr std::array<CategoryInfo, 9> kCategories = {{
    {PrefCategory::General, IDS_PREFS_CAT_GENERAL, IDS_PREFS_CAT_GENERAL_DESC},
    {PrefCategory::Panes, IDS_PREFS_CAT_PANES, IDS_PREFS_CAT_PANES_DESC},
    {PrefCategory::Viewers, IDS_PREFS_CAT_VIEWERS, IDS_PREFS_CAT_VIEWERS_DESC},
    {PrefCategory::Editors, IDS_PREFS_CAT_EDITORS, IDS_PREFS_CAT_EDITORS_DESC},
    {PrefCategory::Keyboard, IDS_PREFS_CAT_KEYBOARD, IDS_PREFS_CAT_KEYBOARD_DESC},
    {PrefCategory::Mouse, IDS_PREFS_CAT_MOUSE, IDS_PREFS_CAT_MOUSE_DESC},
    {PrefCategory::Themes, IDS_PREFS_CAT_THEMES, IDS_PREFS_CAT_THEMES_DESC},
    {PrefCategory::Plugins, IDS_PREFS_CAT_PLUGINS, IDS_PREFS_CAT_PLUGINS_DESC},
    {PrefCategory::Advanced, IDS_PREFS_CAT_ADVANCED, IDS_PREFS_CAT_ADVANCED_DESC},
}};

wil::unique_hwnd g_preferencesDialog;

[[nodiscard]] const CategoryInfo* FindCategoryInfo(PrefCategory id) noexcept
{
    for (const auto& c : kCategories)
    {
        if (c.id == id)
        {
            return &c;
        }
    }
    return nullptr;
}

[[nodiscard]] PreferencesDialogState* GetState(HWND dlg) noexcept
{
    return reinterpret_cast<PreferencesDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
}

void SetState(HWND dlg, PreferencesDialogState* state) noexcept
{
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
}

void ShowDialogAlert(HWND dlg, HostAlertSeverity severity, const std::wstring& title, const std::wstring& message) noexcept
{
    if (! dlg || message.empty())
    {
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = severity;
    request.targetWindow = dlg;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(HostShowAlert(request));
}

[[nodiscard]] HFONT GetDialogFont(HWND hwnd) noexcept
{
    HFONT font = hwnd ? reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0)) : nullptr;
    if (! font)
    {
        font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }
    return font;
}

BOOL CALLBACK SetDialogChildFontProc(HWND child, LPARAM fontParam) noexcept
{
    const HFONT font = reinterpret_cast<HFONT>(fontParam);
    if (font)
    {
        SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }
    return TRUE;
}

void EnsureFonts(PreferencesDialogState& state, HFONT baseFont) noexcept
{
    if (! baseFont)
    {
        baseFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    }

    if (! state.italicFont)
    {
        LOGFONTW lf{};
        if (GetObjectW(baseFont, sizeof(lf), &lf) == sizeof(lf))
        {
            lf.lfItalic = TRUE;
            state.italicFont.reset(CreateFontIndirectW(&lf));
        }
    }

    if (! state.boldFont)
    {
        LOGFONTW lf{};
        if (GetObjectW(baseFont, sizeof(lf), &lf) == sizeof(lf))
        {
            lf.lfWeight = FW_SEMIBOLD;
            state.boldFont.reset(CreateFontIndirectW(&lf));
        }
    }

    if (! state.titleFont)
    {
        LOGFONTW lf{};
        if (GetObjectW(baseFont, sizeof(lf), &lf) == sizeof(lf))
        {
            lf.lfWeight = FW_SEMIBOLD;
            if (lf.lfHeight != 0)
            {
                lf.lfHeight *= 2;
            }
            else
            {
                lf.lfHeight = -24;
            }
            state.titleFont.reset(CreateFontIndirectW(&lf));
        }
    }
}

[[nodiscard]] Common::Settings::MainMenuState GetMainMenu(const Common::Settings::Settings& settings) noexcept
{
    if (settings.mainMenu.has_value())
    {
        return settings.mainMenu.value();
    }
    return {};
}

[[nodiscard]] const Common::Settings::StartupSettings& GetStartupSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::StartupSettings kDefaults{};
    if (settings.startup.has_value())
    {
        return settings.startup.value();
    }
    return kDefaults;
}

[[nodiscard]] const Common::Settings::MonitorSettings& GetMonitorSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::MonitorSettings kDefaults{};
    if (settings.monitor.has_value())
    {
        return settings.monitor.value();
    }
    return kDefaults;
}

[[nodiscard]] const Common::Settings::CacheSettings& GetCacheSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::CacheSettings kDefaults{};
    if (settings.cache.has_value())
    {
        return settings.cache.value();
    }
    return kDefaults;
}

[[nodiscard]] const Common::Settings::FileOperationsSettings& GetFileOperationsSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::FileOperationsSettings kDefaults{};
    if (settings.fileOperations.has_value())
    {
        return settings.fileOperations.value();
    }
    return kDefaults;
}

[[nodiscard]] bool AreEquivalentShortcutBindings(const std::vector<Common::Settings::ShortcutBinding>& a,
                                                 const std::vector<Common::Settings::ShortcutBinding>& b) noexcept
{
    using Key = std::tuple<uint32_t, uint32_t, std::wstring>;

    auto normalize = [](const std::vector<Common::Settings::ShortcutBinding>& bindings)
    {
        std::vector<Key> keys;
        keys.reserve(bindings.size());
        for (const auto& binding : bindings)
        {
            if (binding.commandId.empty())
            {
                continue;
            }
            keys.emplace_back(binding.vk, binding.modifiers & 0x7u, binding.commandId);
        }

        std::sort(keys.begin(), keys.end());
        keys.erase(std::unique(keys.begin(), keys.end()), keys.end());
        return keys;
    };

    return normalize(a) == normalize(b);
}

[[nodiscard]] bool AreEquivalentShortcuts(const std::optional<Common::Settings::ShortcutsSettings>& a,
                                          const std::optional<Common::Settings::ShortcutsSettings>& b) noexcept
{
    if (! a.has_value() && ! b.has_value())
    {
        return true;
    }

    if (a.has_value() && b.has_value())
    {
        return AreEquivalentShortcutBindings(a.value().functionBar, b.value().functionBar) &&
               AreEquivalentShortcutBindings(a.value().folderView, b.value().folderView);
    }

    const Common::Settings::ShortcutsSettings defaults = ShortcutDefaults::CreateDefaultShortcuts();
    const Common::Settings::ShortcutsSettings& aValue  = a.has_value() ? a.value() : defaults;
    const Common::Settings::ShortcutsSettings& bValue  = b.has_value() ? b.value() : defaults;

    return AreEquivalentShortcutBindings(aValue.functionBar, bValue.functionBar) && AreEquivalentShortcutBindings(aValue.folderView, bValue.folderView);
}

[[nodiscard]] bool AreEquivalentThemeDefinition(const Common::Settings::ThemeDefinition& a, const Common::Settings::ThemeDefinition& b) noexcept
{
    if (a.id != b.id || a.name != b.name || a.baseThemeId != b.baseThemeId)
    {
        return false;
    }

    if (a.colors.size() != b.colors.size())
    {
        return false;
    }

    for (const auto& [key, value] : a.colors)
    {
        const auto it = b.colors.find(key);
        if (it == b.colors.end() || it->second != value)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool AreEquivalentThemeSettings(const Common::Settings::ThemeSettings& a, const Common::Settings::ThemeSettings& b) noexcept
{
    if (a.currentThemeId != b.currentThemeId)
    {
        return false;
    }

    if (a.themes.size() != b.themes.size())
    {
        return false;
    }

    for (const auto& theme : a.themes)
    {
        const auto it =
            std::find_if(b.themes.begin(), b.themes.end(), [&](const Common::Settings::ThemeDefinition& other) noexcept { return other.id == theme.id; });
        if (it == b.themes.end() || ! AreEquivalentThemeDefinition(theme, *it))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool AreEquivalentJsonValue(const Common::Settings::JsonValue& a, const Common::Settings::JsonValue& b) noexcept
{
    return a.value == b.value;
}

[[nodiscard]] bool AreEquivalentPluginsDisabledIds(const std::vector<std::wstring>& a, const std::vector<std::wstring>& b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        const std::wstring_view id = a[i];
        if (id.empty())
        {
            return false;
        }

        for (size_t j = 0; j < i; ++j)
        {
            if (a[j] == a[i])
            {
                return false;
            }
        }

        if (std::find(b.begin(), b.end(), id) == b.end())
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool AreEquivalentPluginsSettings(const Common::Settings::PluginsSettings& a, const Common::Settings::PluginsSettings& b) noexcept
{
    if (a.currentFileSystemPluginId != b.currentFileSystemPluginId)
    {
        return false;
    }
    if (a.customPluginPaths != b.customPluginPaths)
    {
        return false;
    }
    if (! AreEquivalentPluginsDisabledIds(a.disabledPluginIds, b.disabledPluginIds))
    {
        return false;
    }

    if (a.configurationByPluginId.size() != b.configurationByPluginId.size())
    {
        return false;
    }

    for (const auto& [id, value] : a.configurationByPluginId)
    {
        const auto it = b.configurationByPluginId.find(id);
        if (it == b.configurationByPluginId.end())
        {
            return false;
        }
        if (! AreEquivalentJsonValue(value, it->second))
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool IsDirty(const PreferencesDialogState& state) noexcept
{
    const Common::Settings::MainMenuState baseline = GetMainMenu(state.baselineSettings);
    const Common::Settings::MainMenuState working  = GetMainMenu(state.workingSettings);
    if (baseline.menuBarVisible != working.menuBarVisible)
    {
        return true;
    }
    if (baseline.functionBarVisible != working.functionBarVisible)
    {
        return true;
    }
    {
        const auto& baselineStartup = GetStartupSettingsOrDefault(state.baselineSettings);
        const auto& workingStartup  = GetStartupSettingsOrDefault(state.workingSettings);
        if (baselineStartup.showSplash != workingStartup.showSplash)
        {
            return true;
        }
    }
    if (! AreEquivalentShortcuts(state.baselineSettings.shortcuts, state.workingSettings.shortcuts))
    {
        return true;
    }
    if (! AreEquivalentThemeSettings(state.baselineSettings.theme, state.workingSettings.theme))
    {
        return true;
    }
    if (! PrefsFolders::AreEquivalentFolderPreferences(state.baselineSettings, state.workingSettings))
    {
        return true;
    }
    {
        const auto& baselineMonitor = GetMonitorSettingsOrDefault(state.baselineSettings);
        const auto& workingMonitor  = GetMonitorSettingsOrDefault(state.workingSettings);
        if (baselineMonitor.menu.toolbarVisible != workingMonitor.menu.toolbarVisible ||
            baselineMonitor.menu.lineNumbersVisible != workingMonitor.menu.lineNumbersVisible ||
            baselineMonitor.menu.alwaysOnTop != workingMonitor.menu.alwaysOnTop || baselineMonitor.menu.showIds != workingMonitor.menu.showIds ||
            baselineMonitor.menu.autoScroll != workingMonitor.menu.autoScroll || baselineMonitor.filter.mask != workingMonitor.filter.mask ||
            baselineMonitor.filter.preset != workingMonitor.filter.preset)
        {
            return true;
        }
    }
    {
        const auto& baselineCache = GetCacheSettingsOrDefault(state.baselineSettings);
        const auto& workingCache  = GetCacheSettingsOrDefault(state.workingSettings);
        if (baselineCache.directoryInfo.maxBytes != workingCache.directoryInfo.maxBytes ||
            baselineCache.directoryInfo.maxWatchers != workingCache.directoryInfo.maxWatchers ||
            baselineCache.directoryInfo.mruWatched != workingCache.directoryInfo.mruWatched)
        {
            return true;
        }
    }
    {
        const auto& baselineFileOperations = GetFileOperationsSettingsOrDefault(state.baselineSettings);
        const auto& workingFileOperations  = GetFileOperationsSettingsOrDefault(state.workingSettings);
        if (baselineFileOperations.autoDismissSuccess != workingFileOperations.autoDismissSuccess ||
            baselineFileOperations.maxDiagnosticsLogFiles != workingFileOperations.maxDiagnosticsLogFiles ||
            baselineFileOperations.diagnosticsInfoEnabled != workingFileOperations.diagnosticsInfoEnabled ||
            baselineFileOperations.diagnosticsDebugEnabled != workingFileOperations.diagnosticsDebugEnabled ||
            baselineFileOperations.maxIssueReportFiles != workingFileOperations.maxIssueReportFiles ||
            baselineFileOperations.maxDiagnosticsInMemory != workingFileOperations.maxDiagnosticsInMemory ||
            baselineFileOperations.maxDiagnosticsPerFlush != workingFileOperations.maxDiagnosticsPerFlush ||
            baselineFileOperations.diagnosticsFlushIntervalMs != workingFileOperations.diagnosticsFlushIntervalMs ||
            baselineFileOperations.diagnosticsCleanupIntervalMs != workingFileOperations.diagnosticsCleanupIntervalMs)
        {
            return true;
        }
    }
    if (state.baselineSettings.extensions.openWithViewerByExtension != state.workingSettings.extensions.openWithViewerByExtension)
    {
        return true;
    }
    if (! AreEquivalentPluginsSettings(state.baselineSettings.plugins, state.workingSettings.plugins))
    {
        return true;
    }
    return false;
}

void UpdateApplyButton(HWND dlg, const PreferencesDialogState& state) noexcept
{
    HWND apply = dlg ? GetDlgItem(dlg, IDC_PREFS_APPLY) : nullptr;
    if (! apply)
    {
        return;
    }

    EnableWindow(apply, state.dirty ? TRUE : FALSE);
}

} // namespace

void SetDirty(HWND dlg, PreferencesDialogState& state) noexcept
{
    state.dirty = IsDirty(state);
    UpdateApplyButton(dlg, state);
}

namespace
{
[[nodiscard]] HRESULT SaveSettingsFromDialog(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (state.appId.empty())
    {
        return E_INVALIDARG;
    }

    if (state.owner && IsWindow(state.owner))
    {
        SendMessageW(state.owner, WndMsg::kPreferencesRequestSettingsSnapshot, 0, 0);
    }

    Common::Settings::Settings merged;
    merged = state.settings ? *state.settings : state.workingSettings;

    const Common::Settings::MainMenuState baselineMenu = GetMainMenu(state.baselineSettings);
    const Common::Settings::MainMenuState workingMenu  = GetMainMenu(state.workingSettings);
    // Always preserve mainMenu if it exists in working settings or if values differ from baseline
    // This ensures defaults are explicitly saved rather than relying on implicit defaults
    if (state.workingSettings.mainMenu.has_value() || baselineMenu.menuBarVisible != workingMenu.menuBarVisible ||
        baselineMenu.functionBarVisible != workingMenu.functionBarVisible)
    {
        merged.mainMenu = workingMenu;
    }

    {
        const auto& baselineStartup = GetStartupSettingsOrDefault(state.baselineSettings);
        const auto& workingStartup  = GetStartupSettingsOrDefault(state.workingSettings);
        if (baselineStartup.showSplash != workingStartup.showSplash)
        {
            merged.startup = workingStartup;
        }
    }

    if (! AreEquivalentShortcuts(state.baselineSettings.shortcuts, state.workingSettings.shortcuts))
    {
        merged.shortcuts = state.workingSettings.shortcuts;
    }
    if (! AreEquivalentThemeSettings(state.baselineSettings.theme, state.workingSettings.theme))
    {
        merged.theme = state.workingSettings.theme;
    }
    if (! PrefsFolders::AreEquivalentFolderPreferences(state.baselineSettings, state.workingSettings))
    {
        const PrefsFolders::FolderPanePreferences left  = PrefsFolders::GetFolderPanePreferences(state.workingSettings, PrefsFolders::kLeftPaneSlot);
        const PrefsFolders::FolderPanePreferences right = PrefsFolders::GetFolderPanePreferences(state.workingSettings, PrefsFolders::kRightPaneSlot);
        const uint32_t historyMax                       = PrefsFolders::GetFolderHistoryMax(state.workingSettings);

        auto* folders = PrefsFolders::EnsureWorkingFoldersSettings(merged);
        if (folders)
        {
            folders->historyMax = historyMax;

            if (auto* pane = PrefsFolders::EnsureWorkingFolderPane(merged, PrefsFolders::kLeftPaneSlot))
            {
                pane->view.display          = left.display;
                pane->view.sortBy           = left.sortBy;
                pane->view.sortDirection    = left.sortDirection;
                pane->view.statusBarVisible = left.statusBarVisible;
            }
            if (auto* pane = PrefsFolders::EnsureWorkingFolderPane(merged, PrefsFolders::kRightPaneSlot))
            {
                pane->view.display          = right.display;
                pane->view.sortBy           = right.sortBy;
                pane->view.sortDirection    = right.sortDirection;
                pane->view.statusBarVisible = right.statusBarVisible;
            }
        }
    }

    {
        const auto& baselineMonitor = GetMonitorSettingsOrDefault(state.baselineSettings);
        const auto& workingMonitor  = GetMonitorSettingsOrDefault(state.workingSettings);
        if (baselineMonitor.menu.toolbarVisible != workingMonitor.menu.toolbarVisible ||
            baselineMonitor.menu.lineNumbersVisible != workingMonitor.menu.lineNumbersVisible ||
            baselineMonitor.menu.alwaysOnTop != workingMonitor.menu.alwaysOnTop || baselineMonitor.menu.showIds != workingMonitor.menu.showIds ||
            baselineMonitor.menu.autoScroll != workingMonitor.menu.autoScroll || baselineMonitor.filter.mask != workingMonitor.filter.mask ||
            baselineMonitor.filter.preset != workingMonitor.filter.preset)
        {
            merged.monitor = workingMonitor;
        }
    }
    {
        const auto& baselineCache = GetCacheSettingsOrDefault(state.baselineSettings);
        const auto& workingCache  = GetCacheSettingsOrDefault(state.workingSettings);
        if (baselineCache.directoryInfo.maxBytes != workingCache.directoryInfo.maxBytes ||
            baselineCache.directoryInfo.maxWatchers != workingCache.directoryInfo.maxWatchers ||
            baselineCache.directoryInfo.mruWatched != workingCache.directoryInfo.mruWatched)
        {
            merged.cache = state.workingSettings.cache;
        }
    }
    {
        const auto& baselineFileOperations = GetFileOperationsSettingsOrDefault(state.baselineSettings);
        const auto& workingFileOperations  = GetFileOperationsSettingsOrDefault(state.workingSettings);
        if (baselineFileOperations.autoDismissSuccess != workingFileOperations.autoDismissSuccess ||
            baselineFileOperations.maxDiagnosticsLogFiles != workingFileOperations.maxDiagnosticsLogFiles ||
            baselineFileOperations.maxIssueReportFiles != workingFileOperations.maxIssueReportFiles ||
            baselineFileOperations.maxDiagnosticsInMemory != workingFileOperations.maxDiagnosticsInMemory ||
            baselineFileOperations.maxDiagnosticsPerFlush != workingFileOperations.maxDiagnosticsPerFlush ||
            baselineFileOperations.diagnosticsFlushIntervalMs != workingFileOperations.diagnosticsFlushIntervalMs ||
            baselineFileOperations.diagnosticsCleanupIntervalMs != workingFileOperations.diagnosticsCleanupIntervalMs)
        {
            merged.fileOperations = state.workingSettings.fileOperations;
        }
    }
    if (state.baselineSettings.extensions.openWithViewerByExtension != state.workingSettings.extensions.openWithViewerByExtension)
    {
        merged.extensions.openWithViewerByExtension = state.workingSettings.extensions.openWithViewerByExtension;
    }
    if (! AreEquivalentPluginsSettings(state.baselineSettings.plugins, state.workingSettings.plugins))
    {
        merged.plugins = state.workingSettings.plugins;
    }

    Common::Settings::Settings settingsToSave;
    settingsToSave = SettingsSave::PrepareForSave(merged);

    const HRESULT hr = Common::Settings::SaveSettings(state.appId, settingsToSave);
    if (FAILED(hr))
    {
        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(state.appId);
        const std::wstring title                 = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        const std::wstring message = FormatStringResource(nullptr, IDS_FMT_SETTINGS_SAVE_FAILED, settingsPath.wstring(), static_cast<unsigned long>(hr));
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, title, message);
        return hr;
    }

    const HRESULT schemaHr = SaveAggregatedSettingsSchema(state.appId, settingsToSave);
    if (FAILED(schemaHr))
    {
        Debug::Error(L"SaveAggregatedSettingsSchema failed (hr=0x{:08X})", static_cast<unsigned long>(schemaHr));
    }

    state.workingSettings = std::move(settingsToSave);

    return S_OK;
}

void RefreshPreferencesDialogTheme(HWND dlg, PreferencesDialogState& state) noexcept;

void CommitAndApply(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! dlg || ! state.settings)
    {
        return;
    }

    const HRESULT saveHr = SaveSettingsFromDialog(dlg, state);
    if (FAILED(saveHr))
    {
        return;
    }

    const bool pluginsChanged = ! AreEquivalentPluginsSettings(state.baselineSettings.plugins, state.workingSettings.plugins);

    *state.settings        = state.workingSettings;
    state.baselineSettings = state.workingSettings;
    state.previewApplied   = false;

    state.appliedOnce = true;
    SetDirty(dlg, state);

    if (state.owner)
    {
        PostMessageW(state.owner, WndMsg::kSettingsApplied, 0, 0);
    }
    if (pluginsChanged && state.owner)
    {
        PostMessageW(state.owner, WndMsg::kPluginsChanged, 0, 0);
    }

    RefreshPreferencesDialogTheme(dlg, state);
}

void LayoutPreferencesDialog(HWND dlg, PreferencesDialogState& state) noexcept;
void LayoutPreferencesPageHost(HWND host, PreferencesDialogState& state) noexcept;

void RefreshPreferencesDialogTheme(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! dlg || ! state.settings)
    {
        return;
    }

    ApplyThemeToPreferencesDialog(dlg, state, ResolveThemeFromSettingsForDialog(*state.settings));
    LayoutPreferencesDialog(dlg, state);
    if (state.pageHost)
    {
        LayoutPreferencesPageHost(state.pageHost, state);
        RedrawWindow(state.pageHost, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
    }
    RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN | RDW_UPDATENOW);
}

[[nodiscard]] int MeasurePageHostContentHeightPx(HWND host, const PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return 0;
    }

    const auto& hostState                 = static_cast<const PreferencesDialogHost&>(state);
    const std::array<HWND, 9> paneWindows = {
        hostState._generalPane.Hwnd(),
        hostState._panesPane.Hwnd(),
        hostState._viewersPane.Hwnd(),
        hostState._editorsPane.Hwnd(),
        hostState._keyboardPane.Hwnd(),
        hostState._mousePane.Hwnd(),
        hostState._themesPane.Hwnd(),
        hostState._pluginsPane.Hwnd(),
        hostState._advancedPane.Hwnd(),
    };

    const auto isPaneWindow = [&](HWND hwnd) noexcept
    {
        for (HWND pane : paneWindows)
        {
            if (pane && pane == hwnd)
            {
                return true;
            }
        }
        return false;
    };

    int maxBottomPx = 0;

    HWND current = GetWindow(host, GW_CHILD);
    while (current)
    {
        if (IsWindowVisible(current))
        {
            if (! isPaneWindow(current))
            {
                RECT rc{};
                if (GetWindowRect(current, &rc))
                {
                    MapWindowPoints(nullptr, host, reinterpret_cast<POINT*>(&rc), 2);
                    maxBottomPx = std::max(maxBottomPx, static_cast<int>(rc.bottom));
                }
            }
        }

        HWND next = GetWindow(current, GW_CHILD);
        if (next)
        {
            current = next;
            continue;
        }

        while (current)
        {
            next = GetWindow(current, GW_HWNDNEXT);
            if (next)
            {
                current = next;
                break;
            }

            current = GetParent(current);
            if (! current || current == host)
            {
                current = nullptr;
                break;
            }
        }
    }

    return std::max(0, maxBottomPx);
}

void UpdatePageHostScrollInfo(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    RECT client{};
    GetClientRect(host, &client);
    const int clientWidth  = std::max(0l, client.right - client.left);
    const int clientHeight = std::max(0l, client.bottom - client.top);

    const UINT dpi          = GetDpiForWindow(host);
    const int paddingBottom = ThemedControls::ScaleDip(dpi, 12);
    int contentHeight       = MeasurePageHostContentHeightPx(host, state);
    for (const RECT& card : state.pageSettingCards)
    {
        contentHeight = std::max(contentHeight, static_cast<int>(card.bottom));
    }
    contentHeight = std::max(0, contentHeight + paddingBottom);

    // The page host scrolls its child "pane" window(s). Those panes must be tall enough to contain
    // all laid-out controls; otherwise controls below the pane's client rect get clipped and appear
    // as "blank cards" after scrolling.
    if (const HWND pane = GetActivePrefsPaneWindow(state); pane && IsWindow(pane))
    {
        const int desiredHeight = std::max(clientHeight, contentHeight);
        SetWindowPos(pane, nullptr, 0, 0, clientWidth, desiredHeight, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    }

    const int maxScroll     = std::max(0, contentHeight - clientHeight);
    const bool wantsVScroll = maxScroll > 0;
    state.pageScrollMaxY    = maxScroll;
    state.pageScrollY       = std::clamp(state.pageScrollY, 0, maxScroll);

    const LONG_PTR styleNow = GetWindowLongPtrW(host, GWL_STYLE);
    LONG_PTR styleWanted    = styleNow;
    styleWanted &= ~WS_HSCROLL;
    if (wantsVScroll)
    {
        styleWanted |= WS_VSCROLL;
    }
    else
    {
        styleWanted &= ~WS_VSCROLL;
    }

    if (styleWanted != styleNow)
    {
        state.pageHostIgnoreSize = true;
        auto clearIgnore         = wil::scope_exit([&]() noexcept { state.pageHostIgnoreSize = false; });

        SetWindowLongPtrW(host, GWL_STYLE, styleWanted);
        SetWindowPos(host, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        SendMessageW(host, WM_THEMECHANGED, 0, 0);
        RedrawWindow(host, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = (contentHeight > 0) ? (contentHeight - 1) : 0;
    si.nPage  = static_cast<UINT>(clientHeight);
    si.nPos   = state.pageScrollY;
    SetScrollInfo(host, SB_VERT, &si, TRUE);
}

void ApplyPageHostScrollFromLayout(HWND host, const PreferencesDialogState& state) noexcept
{
    if (! host || state.pageScrollY == 0)
    {
        return;
    }

    PrefsPaneHost::ApplyScrollDelta(host, -state.pageScrollY);
    RedrawWindow(host, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
}

void FinalizePreferencesPageHostLayout(HWND host, PreferencesDialogState& state, int margin, int layoutWidth) noexcept
{
    if (! host)
    {
        return;
    }

    UpdatePageHostScrollInfo(host, state);
    ApplyPageHostScrollFromLayout(host, state);

    if (state.pageHostRelayoutInProgress)
    {
        return;
    }

    RECT client{};
    GetClientRect(host, &client);
    const int clientWidth = std::max(0l, client.right - client.left);
    const int widthNow    = std::max(0, clientWidth - 2 * margin);
    if (widthNow == layoutWidth)
    {
        return;
    }

    state.pageHostRelayoutInProgress = true;
    LayoutPreferencesPageHost(host, state);
    state.pageHostRelayoutInProgress = false;
}

[[nodiscard]] HWND FindFirstOrLastTabStopChild(HWND host, bool forward) noexcept
{
    if (! host)
    {
        return nullptr;
    }

    const HWND dlg = GetParent(host);
    if (! dlg)
    {
        return nullptr;
    }

    const BOOL previous = forward ? FALSE : TRUE;
    const HWND start    = GetNextDlgTabItem(dlg, nullptr, previous);
    if (! start)
    {
        return nullptr;
    }

    HWND item = start;
    do
    {
        if (IsChild(host, item) && IsWindowVisible(item) && IsWindowEnabled(item))
        {
            const LONG_PTR style = GetWindowLongPtrW(item, GWL_STYLE);
            if ((style & WS_TABSTOP) != 0)
            {
                return item;
            }
        }

        item = GetNextDlgTabItem(dlg, item, previous);
    } while (item && item != start);

    return nullptr;
}

void LayoutPreferencesDialog(HWND dlg, PreferencesDialogState& state) noexcept
{
    if (! dlg)
    {
        return;
    }

    HWND list   = state.categoryTree ? state.categoryTree : GetDlgItem(dlg, IDC_PREFS_CATEGORY_LIST);
    HWND host   = state.pageHost ? state.pageHost : GetDlgItem(dlg, IDC_PREFS_PAGE_HOST);
    HWND ok     = GetDlgItem(dlg, IDOK);
    HWND cancel = GetDlgItem(dlg, IDCANCEL);
    HWND apply  = GetDlgItem(dlg, IDC_PREFS_APPLY);
    if (! list || ! host || ! ok || ! cancel || ! apply)
    {
        return;
    }

    RECT client{};
    GetClientRect(dlg, &client);

    const UINT dpi   = GetDpiForWindow(dlg);
    const int margin = ThemedControls::ScaleDip(dpi, 8);
    const int gapX   = ThemedControls::ScaleDip(dpi, 8);

    RECT okRect{};
    RECT cancelRect{};
    RECT applyRect{};
    GetWindowRect(ok, &okRect);
    GetWindowRect(cancel, &cancelRect);
    GetWindowRect(apply, &applyRect);

    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&okRect), 2);
    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&cancelRect), 2);
    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&applyRect), 2);

    const int okWidthDesired     = std::max(0l, okRect.right - okRect.left);
    const int cancelWidthDesired = std::max(0l, cancelRect.right - cancelRect.left);
    const int applyWidthDesired  = std::max(0l, applyRect.right - applyRect.left);

    const int buttonHeight =
        std::max({std::max(0l, okRect.bottom - okRect.top), std::max(0l, cancelRect.bottom - cancelRect.top), std::max(0l, applyRect.bottom - applyRect.top)});

    const int buttonPadX = ThemedControls::ScaleDip(dpi, 12);
    const int minGapX    = ThemedControls::ScaleDip(dpi, 4);

    const auto measureButtonMinWidth = [&](HWND button) noexcept
    {
        if (! button)
        {
            return 0;
        }

        HFONT font = reinterpret_cast<HFONT>(SendMessageW(button, WM_GETFONT, 0, 0));
        if (! font)
        {
            font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        }

        const std::wstring text = PrefsUi::GetWindowTextString(button);
        const int textW         = ThemedControls::MeasureTextWidth(button, font, text);
        return std::max(ThemedControls::ScaleDip(dpi, 60), textW + 2 * buttonPadX);
    };

    const int okWidthMin     = measureButtonMinWidth(ok);
    const int cancelWidthMin = measureButtonMinWidth(cancel);
    const int applyWidthMin  = measureButtonMinWidth(apply);

    const int clientWidth         = std::max(0l, client.right - client.left);
    const int groupAvailableWidth = std::max(0, clientWidth - 2 * margin);
    int gapUsed                   = gapX;
    int minGroupWidth             = okWidthMin + cancelWidthMin + applyWidthMin + 2 * gapUsed;
    if (minGroupWidth > groupAvailableWidth)
    {
        gapUsed       = minGapX;
        minGroupWidth = okWidthMin + cancelWidthMin + applyWidthMin + 2 * gapUsed;
    }

    int okWidth     = okWidthDesired;
    int cancelWidth = cancelWidthDesired;
    int applyWidth  = applyWidthDesired;

    const int desiredGroupWidth = okWidth + cancelWidth + applyWidth + 2 * gapUsed;
    if (desiredGroupWidth > groupAvailableWidth)
    {
        okWidth     = okWidthMin;
        cancelWidth = cancelWidthMin;
        applyWidth  = applyWidthMin;

        int remaining   = std::max(0, groupAvailableWidth - 2 * gapUsed - (okWidth + cancelWidth + applyWidth));
        const auto grow = [&](int& width, int desired) noexcept
        {
            const int target = std::max(width, desired);
            const int add    = std::min(remaining, target - width);
            if (add > 0)
            {
                width += add;
                remaining -= add;
            }
        };

        grow(applyWidth, applyWidthDesired);
        grow(cancelWidth, cancelWidthDesired);
        grow(okWidth, okWidthDesired);
    }

    // Last-resort safety: avoid overlap if the window was resized smaller than the computed minimum.
    int finalGroupWidth = okWidth + cancelWidth + applyWidth + 2 * gapUsed;
    if (finalGroupWidth > groupAvailableWidth)
    {
        gapUsed                 = minGapX;
        int availableForButtons = std::max(0, groupAvailableWidth - 2 * gapUsed);
        if (availableForButtons < 3)
        {
            gapUsed             = 0;
            availableForButtons = std::max(0, groupAvailableWidth);
        }

        const int baseWidth = std::max(1, availableForButtons / 3);
        okWidth             = baseWidth;
        cancelWidth         = baseWidth;
        applyWidth          = baseWidth;

        const int remainder = std::max(0, availableForButtons - (baseWidth * 3));
        applyWidth += remainder;
        finalGroupWidth = okWidth + cancelWidth + applyWidth + 2 * gapUsed;
    }

    const int applyLeft  = static_cast<int>(client.right) - margin - applyWidth;
    const int cancelLeft = applyLeft - gapUsed - cancelWidth;
    const int okLeft     = cancelLeft - gapUsed - okWidth;
    const int buttonsTop = std::max(0, static_cast<int>(client.bottom) - margin - buttonHeight);

    const UINT moveFlags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS;

    const int contentTop    = margin;
    const int contentBottom = std::max(contentTop, buttonsTop - margin);
    const int contentHeight = std::max(0, contentBottom - contentTop);

    const int listDesiredWidth = state.categoryListWidthPx > 0 ? state.categoryListWidthPx : ThemedControls::ScaleDip(dpi, 120);
    const int listMinWidth     = ThemedControls::ScaleDip(dpi, 72);
    const int hostMinWidth     = ThemedControls::ScaleDip(dpi, 140);

    const int availableForList = std::max(0, groupAvailableWidth - gapX - hostMinWidth);
    const int listMaxWidth     = std::max(listMinWidth, availableForList);
    const int listWidth        = std::clamp(listDesiredWidth, listMinWidth, listMaxWidth);

    const int hostLeft  = std::max(0, margin + listWidth + gapX);
    const int hostWidth = std::max(0, static_cast<int>(client.right) - margin - hostLeft);

    const HFONT dialogFont = GetDialogFont(dlg);
    EnsureFonts(state, dialogFont);
    const HFONT titleFont = state.titleFont ? state.titleFont.get() : (state.boldFont ? state.boldFont.get() : dialogFont);

    const int headerMargin   = ThemedControls::ScaleDip(dpi, 12);
    const int headerGapY     = ThemedControls::ScaleDip(dpi, 6);
    const int headerSectionY = ThemedControls::ScaleDip(dpi, 14);

    const int headerX     = hostLeft + headerMargin;
    const int headerY     = contentTop + headerMargin;
    const int headerWidth = std::max(0, hostWidth - 2 * headerMargin);

    int headerContentY = headerY;
    int titleTop       = 0;
    int titleHeight    = 0;
    int descTop        = 0;
    int descHeightPx   = 0;
    if (state.pageTitle)
    {
        const std::wstring titleText  = PrefsUi::GetWindowTextString(state.pageTitle);
        titleTop                      = headerContentY;
        const int measuredTitleHeight = PrefsUi::MeasureStaticTextHeight(dlg, titleFont, headerWidth, titleText);
        titleHeight                   = std::max(ThemedControls::ScaleDip(dpi, 40), std::max(0, measuredTitleHeight));
        SendMessageW(state.pageTitle, WM_SETFONT, reinterpret_cast<WPARAM>(titleFont), TRUE);
        headerContentY += titleHeight + headerGapY;
    }

    if (state.pageDescription)
    {
        const std::wstring desc      = PrefsUi::GetWindowTextString(state.pageDescription);
        descTop                      = headerContentY;
        const int measuredDescHeight = PrefsUi::MeasureStaticTextHeight(dlg, dialogFont, headerWidth, desc);
        descHeightPx                 = std::max(0, measuredDescHeight);
        SendMessageW(state.pageDescription, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        headerContentY += descHeightPx + headerSectionY;
    }

    const int hostTop    = std::clamp(headerContentY - headerMargin, contentTop, contentBottom);
    const int hostHeight = std::max(0, contentBottom - hostTop);

    bool moved = false;
    if (HDWP hdwp = BeginDeferWindowPos(8))
    {
        auto deferWindow = [&](HWND hwnd, int left, int top, int width, int height) noexcept
        {
            if (! hwnd || ! hdwp)
            {
                return;
            }
            hdwp = DeferWindowPos(hdwp, hwnd, nullptr, left, top, width, height, moveFlags);
        };

        deferWindow(apply, applyLeft, buttonsTop, applyWidth, buttonHeight);
        deferWindow(cancel, cancelLeft, buttonsTop, cancelWidth, buttonHeight);
        deferWindow(ok, okLeft, buttonsTop, okWidth, buttonHeight);
        deferWindow(list, margin, contentTop, listWidth, contentHeight);
        if (state.pageTitle)
        {
            deferWindow(state.pageTitle, headerX, titleTop, headerWidth, titleHeight);
        }
        if (state.pageDescription)
        {
            deferWindow(state.pageDescription, headerX, descTop, headerWidth, descHeightPx);
        }
        deferWindow(host, hostLeft, hostTop, hostWidth, hostHeight);

        if (hdwp && EndDeferWindowPos(hdwp))
        {
            moved = true;
        }
    }

    if (! moved)
    {
        SetWindowPos(apply, nullptr, applyLeft, buttonsTop, applyWidth, buttonHeight, moveFlags);
        SetWindowPos(cancel, nullptr, cancelLeft, buttonsTop, cancelWidth, buttonHeight, moveFlags);
        SetWindowPos(ok, nullptr, okLeft, buttonsTop, okWidth, buttonHeight, moveFlags);
        SetWindowPos(list, nullptr, margin, contentTop, listWidth, contentHeight, moveFlags);
        if (state.pageTitle)
        {
            SetWindowPos(state.pageTitle, nullptr, headerX, titleTop, headerWidth, titleHeight, moveFlags);
        }
        if (state.pageDescription)
        {
            SetWindowPos(state.pageDescription, nullptr, headerX, descTop, headerWidth, descHeightPx, moveFlags);
        }
        SetWindowPos(host, nullptr, hostLeft, hostTop, hostWidth, hostHeight, moveFlags);
    }
}

void LayoutPreferencesPageHost(HWND host, PreferencesDialogState& state) noexcept
{
    if (! host)
    {
        return;
    }

    RECT client{};
    GetClientRect(host, &client);

    auto& hostState = static_cast<PreferencesDialogHost&>(state);
    hostState._generalPane.ResizeToHostClient(host);
    hostState._panesPane.ResizeToHostClient(host);
    hostState._viewersPane.ResizeToHostClient(host);
    hostState._editorsPane.ResizeToHostClient(host);
    hostState._keyboardPane.ResizeToHostClient(host);
    hostState._mousePane.ResizeToHostClient(host);
    hostState._themesPane.ResizeToHostClient(host);
    hostState._pluginsPane.ResizeToHostClient(host);
    hostState._advancedPane.ResizeToHostClient(host);

    const UINT dpi     = GetDpiForWindow(host);
    const int margin   = ThemedControls::ScaleDip(dpi, 12);
    const int gapY     = ThemedControls::ScaleDip(dpi, 6);
    const int sectionY = ThemedControls::ScaleDip(dpi, 14);

    const int width = std::max(0l, client.right - client.left - 2 * margin);
    int x           = margin;
    int y           = margin;

    const HWND dlg         = GetParent(host);
    const HFONT dialogFont = GetDialogFont(dlg ? dlg : host);
    EnsureFonts(state, dialogFont);

    const bool showGeneral           = state.currentCategory == PrefCategory::General;
    const bool showPanes             = state.currentCategory == PrefCategory::Panes;
    const bool showViewers           = state.currentCategory == PrefCategory::Viewers;
    const bool showEditors           = state.currentCategory == PrefCategory::Editors;
    const bool showKeyboard          = state.currentCategory == PrefCategory::Keyboard;
    const bool showMouse             = state.currentCategory == PrefCategory::Mouse;
    const bool showThemes            = state.currentCategory == PrefCategory::Themes;
    const bool showPlugins           = state.currentCategory == PrefCategory::Plugins;
    const bool showAdvanced          = state.currentCategory == PrefCategory::Advanced;
    const bool panesUseTwoStateCombo = state.theme.systemHighContrast;

    state.pageSettingCards.clear();

    auto setVisible = [&](const auto& hwndLike, bool visible) noexcept
    {
        HWND hwnd = nullptr;
        if constexpr (requires { hwndLike.get(); })
        {
            hwnd = hwndLike.get();
        }
        else
        {
            hwnd = hwndLike;
        }

        if (hwnd)
        {
            ShowWindow(hwnd, visible ? SW_SHOW : SW_HIDE);
        }
    };

    setVisible(hostState._generalPane.Hwnd(), showGeneral);
    setVisible(state.menuBarLabel, showGeneral);
    setVisible(state.menuBarToggle, showGeneral);
    setVisible(state.menuBarDescription, showGeneral);
    setVisible(state.functionBarLabel, showGeneral);
    setVisible(state.functionBarToggle, showGeneral);
    setVisible(state.functionBarDescription, showGeneral);
    setVisible(hostState._panesPane.Hwnd(), showPanes);
    setVisible(state.panesLeftHeader, showPanes);
    setVisible(state.panesLeftDisplayLabel, showPanes);
    setVisible(state.panesLeftDisplayFrame, showPanes && panesUseTwoStateCombo);
    setVisible(state.panesLeftDisplayCombo, showPanes && panesUseTwoStateCombo);
    setVisible(state.panesLeftDisplayToggle, showPanes && ! panesUseTwoStateCombo);
    setVisible(state.panesLeftSortByLabel, showPanes);
    setVisible(state.panesLeftSortByFrame, showPanes);
    setVisible(state.panesLeftSortByCombo, showPanes);
    setVisible(state.panesLeftSortDirLabel, showPanes);
    setVisible(state.panesLeftSortDirFrame, showPanes && panesUseTwoStateCombo);
    setVisible(state.panesLeftSortDirCombo, showPanes && panesUseTwoStateCombo);
    setVisible(state.panesLeftSortDirToggle, showPanes && ! panesUseTwoStateCombo);
    setVisible(state.panesLeftStatusBarLabel, showPanes);
    setVisible(state.panesLeftStatusBarToggle, showPanes);
    setVisible(state.panesLeftStatusBarDescription, showPanes);
    setVisible(state.panesRightHeader, showPanes);
    setVisible(state.panesRightDisplayLabel, showPanes);
    setVisible(state.panesRightDisplayFrame, showPanes && panesUseTwoStateCombo);
    setVisible(state.panesRightDisplayCombo, showPanes && panesUseTwoStateCombo);
    setVisible(state.panesRightDisplayToggle, showPanes && ! panesUseTwoStateCombo);
    setVisible(state.panesRightSortByLabel, showPanes);
    setVisible(state.panesRightSortByFrame, showPanes);
    setVisible(state.panesRightSortByCombo, showPanes);
    setVisible(state.panesRightSortDirLabel, showPanes);
    setVisible(state.panesRightSortDirFrame, showPanes && panesUseTwoStateCombo);
    setVisible(state.panesRightSortDirCombo, showPanes && panesUseTwoStateCombo);
    setVisible(state.panesRightSortDirToggle, showPanes && ! panesUseTwoStateCombo);
    setVisible(state.panesRightStatusBarLabel, showPanes);
    setVisible(state.panesRightStatusBarToggle, showPanes);
    setVisible(state.panesRightStatusBarDescription, showPanes);
    setVisible(state.panesHistoryLabel, showPanes);
    setVisible(state.panesHistoryFrame, showPanes);
    setVisible(state.panesHistoryEdit, showPanes);
    setVisible(state.panesHistoryDescription, showPanes);
    setVisible(hostState._viewersPane.Hwnd(), showViewers);
    setVisible(state.viewersSearchLabel, showViewers);
    setVisible(state.viewersSearchFrame, showViewers);
    setVisible(state.viewersSearchEdit, showViewers);
    setVisible(state.viewersList, showViewers);
    setVisible(state.viewersExtensionLabel, showViewers);
    setVisible(state.viewersExtensionFrame, showViewers);
    setVisible(state.viewersExtensionEdit, showViewers);
    setVisible(state.viewersViewerLabel, showViewers);
    setVisible(state.viewersViewerFrame, showViewers);
    setVisible(state.viewersViewerCombo, showViewers);
    setVisible(state.viewersSaveButton, showViewers);
    setVisible(state.viewersRemoveButton, showViewers);
    setVisible(state.viewersResetButton, showViewers);
    setVisible(state.viewersHint, showViewers);
    setVisible(hostState._editorsPane.Hwnd(), showEditors);
    setVisible(state.editorsNote, showEditors);
    setVisible(hostState._keyboardPane.Hwnd(), showKeyboard);
    setVisible(state.keyboardSearchLabel, showKeyboard);
    setVisible(state.keyboardSearchFrame, showKeyboard);
    setVisible(state.keyboardSearchEdit, showKeyboard);
    setVisible(state.keyboardScopeLabel, showKeyboard);
    setVisible(state.keyboardScopeFrame, showKeyboard);
    setVisible(state.keyboardScopeCombo, showKeyboard);
    setVisible(state.keyboardList, showKeyboard);
    setVisible(state.keyboardHint, showKeyboard);
    setVisible(state.keyboardAssign, showKeyboard);
    setVisible(state.keyboardRemove, showKeyboard);
    setVisible(state.keyboardReset, showKeyboard);
    setVisible(state.keyboardImport, showKeyboard);
    setVisible(state.keyboardExport, showKeyboard);
    setVisible(hostState._mousePane.Hwnd(), showMouse);
    setVisible(state.mouseNote, showMouse);
    setVisible(hostState._themesPane.Hwnd(), showThemes);
    setVisible(state.themesThemeLabel, showThemes);
    setVisible(state.themesThemeFrame, showThemes);
    setVisible(state.themesThemeCombo, showThemes);
    setVisible(state.themesNameLabel, showThemes);
    setVisible(state.themesNameFrame, showThemes);
    setVisible(state.themesNameEdit, showThemes);
    setVisible(state.themesBaseLabel, showThemes);
    setVisible(state.themesBaseFrame, showThemes);
    setVisible(state.themesBaseCombo, showThemes);
    setVisible(state.themesSearchLabel, showThemes);
    setVisible(state.themesSearchFrame, showThemes);
    setVisible(state.themesSearchEdit, showThemes);
    setVisible(state.themesColorsList, showThemes);
    setVisible(state.themesKeyLabel, showThemes);
    setVisible(state.themesKeyFrame, showThemes);
    setVisible(state.themesKeyEdit, showThemes);
    setVisible(state.themesColorLabel, showThemes);
    setVisible(state.themesColorSwatch, showThemes);
    setVisible(state.themesColorFrame, showThemes);
    setVisible(state.themesColorEdit, showThemes);
    setVisible(state.themesPickColor, showThemes);
    setVisible(state.themesSetOverride, showThemes);
    setVisible(state.themesRemoveOverride, showThemes);
    setVisible(state.themesLoadFromFile, showThemes);
    setVisible(state.themesDuplicateTheme, showThemes);
    setVisible(state.themesSaveTheme, showThemes);
    setVisible(state.themesApplyTemporarily, showThemes);
    setVisible(state.themesNote, showThemes);
    const bool showPluginsDetails = showPlugins && state.pluginsSelectedPlugin.has_value();
    const bool showPluginsList    = showPlugins && ! showPluginsDetails;
    setVisible(state.pluginsNote, showPluginsList);
    setVisible(state.pluginsList, showPluginsList);
    setVisible(state.pluginsDetailsHint, showPluginsDetails);
    setVisible(state.pluginsDetailsIdLabel, showPluginsDetails);
    setVisible(state.pluginsDetailsConfigLabel, showPluginsDetails);
    setVisible(state.pluginsDetailsConfigError, showPluginsDetails);
    setVisible(state.pluginsDetailsConfigFrame, showPluginsDetails);
    setVisible(state.pluginsDetailsConfigEdit, showPluginsDetails);
    setVisible(hostState._pluginsPane.Hwnd(), showPlugins);
    setVisible(hostState._advancedPane.Hwnd(), showAdvanced);
    setVisible(state.advancedConnectionsHelloHeader, showAdvanced);
    setVisible(state.advancedConnectionsBypassHelloLabel, showAdvanced);
    setVisible(state.advancedConnectionsBypassHelloToggle, showAdvanced);
    setVisible(state.advancedConnectionsBypassHelloDescription, showAdvanced);
    setVisible(state.advancedConnectionsHelloTimeoutLabel, showAdvanced);
    setVisible(state.advancedConnectionsHelloTimeoutFrame, showAdvanced);
    setVisible(state.advancedConnectionsHelloTimeoutEdit, showAdvanced);
    setVisible(state.advancedConnectionsHelloTimeoutDescription, showAdvanced);
    setVisible(state.advancedMonitorHeader, showAdvanced);
    setVisible(state.advancedMonitorToolbarLabel, showAdvanced);
    setVisible(state.advancedMonitorToolbarToggle, showAdvanced);
    setVisible(state.advancedMonitorToolbarDescription, showAdvanced);
    setVisible(state.advancedMonitorLineNumbersLabel, showAdvanced);
    setVisible(state.advancedMonitorLineNumbersToggle, showAdvanced);
    setVisible(state.advancedMonitorLineNumbersDescription, showAdvanced);
    setVisible(state.advancedMonitorAlwaysOnTopLabel, showAdvanced);
    setVisible(state.advancedMonitorAlwaysOnTopToggle, showAdvanced);
    setVisible(state.advancedMonitorAlwaysOnTopDescription, showAdvanced);
    setVisible(state.advancedMonitorShowIdsLabel, showAdvanced);
    setVisible(state.advancedMonitorShowIdsToggle, showAdvanced);
    setVisible(state.advancedMonitorShowIdsDescription, showAdvanced);
    setVisible(state.advancedMonitorAutoScrollLabel, showAdvanced);
    setVisible(state.advancedMonitorAutoScrollToggle, showAdvanced);
    setVisible(state.advancedMonitorAutoScrollDescription, showAdvanced);
    setVisible(state.advancedMonitorFilterPresetLabel, showAdvanced);
    setVisible(state.advancedMonitorFilterPresetFrame, showAdvanced);
    setVisible(state.advancedMonitorFilterPresetCombo, showAdvanced);
    setVisible(state.advancedMonitorFilterPresetDescription, showAdvanced);
    setVisible(state.advancedMonitorFilterMaskLabel, showAdvanced);
    setVisible(state.advancedMonitorFilterMaskFrame, showAdvanced);
    setVisible(state.advancedMonitorFilterMaskEdit, showAdvanced);
    setVisible(state.advancedMonitorFilterMaskDescription, showAdvanced);
    setVisible(state.advancedMonitorFilterTextLabel, showAdvanced);
    setVisible(state.advancedMonitorFilterTextToggle, showAdvanced);
    setVisible(state.advancedMonitorFilterTextDescription, showAdvanced);
    setVisible(state.advancedMonitorFilterErrorLabel, showAdvanced);
    setVisible(state.advancedMonitorFilterErrorToggle, showAdvanced);
    setVisible(state.advancedMonitorFilterErrorDescription, showAdvanced);
    setVisible(state.advancedMonitorFilterWarningLabel, showAdvanced);
    setVisible(state.advancedMonitorFilterWarningToggle, showAdvanced);
    setVisible(state.advancedMonitorFilterWarningDescription, showAdvanced);
    setVisible(state.advancedMonitorFilterInfoLabel, showAdvanced);
    setVisible(state.advancedMonitorFilterInfoToggle, showAdvanced);
    setVisible(state.advancedMonitorFilterInfoDescription, showAdvanced);
    setVisible(state.advancedMonitorFilterDebugLabel, showAdvanced);
    setVisible(state.advancedMonitorFilterDebugToggle, showAdvanced);
    setVisible(state.advancedMonitorFilterDebugDescription, showAdvanced);
    setVisible(state.advancedCacheHeader, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMaxBytesLabel, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMaxBytesFrame, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMaxBytesEdit, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMaxBytesDescription, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMaxWatchersLabel, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMaxWatchersFrame, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMaxWatchersEdit, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMaxWatchersDescription, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMruWatchedLabel, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMruWatchedFrame, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMruWatchedEdit, showAdvanced);
    setVisible(state.advancedCacheDirectoryInfoMruWatchedDescription, showAdvanced);

    if (showPanes)
    {
        PanesPane::LayoutControls(host, state, x, y, width, margin, gapY, sectionY, dialogFont);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showViewers)
    {
        ViewersPane::LayoutControls(host, state, x, y, width, margin, gapY, dialogFont);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showEditors)
    {
        EditorsPane::LayoutControls(host, state, x, y, width, margin, gapY, sectionY, dialogFont);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showMouse)
    {
        MousePane::LayoutControls(host, state, x, y, width, margin, gapY, sectionY, dialogFont);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showThemes)
    {
        ThemesPane::LayoutControls(host, state, x, y, width, margin, gapY, sectionY, dialogFont);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showPlugins)
    {
        PluginsPane::LayoutControls(host, state, x, y, width, margin, gapY, sectionY, dialogFont);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showAdvanced)
    {
        AdvancedPane::LayoutControls(host, state, x, y, width, margin, gapY, dialogFont);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showKeyboard)
    {
        KeyboardPane::LayoutControls(host, state, x, y, width, margin, gapY, sectionY, dialogFont);
        FinalizePreferencesPageHostLayout(host, state, margin, width);
        return;
    }

    if (showGeneral)
    {
        GeneralPane::LayoutControls(host, state, x, y, width, dialogFont);
    }

    FinalizePreferencesPageHostLayout(host, state, margin, width);
}

void RefreshAdvancedPage(HWND host, PreferencesDialogState& state) noexcept
{
    AdvancedPane::Refresh(host, state);
}

void UpdatePageText(HWND dlg, PreferencesDialogState& state, PrefCategory category) noexcept
{
    const bool categoryChanged = state.currentCategory != category;
    state.currentCategory      = category;

    const bool resetScroll = categoryChanged || category == PrefCategory::Plugins;
    if (resetScroll)
    {
        state.pageScrollY             = 0;
        state.pageScrollMaxY          = 0;
        state.pageWheelDeltaRemainder = 0;
    }

    const CategoryInfo* info = FindCategoryInfo(category);
    std::wstring title       = info ? LoadStringResource(nullptr, info->labelId) : std::wstring{};
    std::wstring description = info ? LoadStringResource(nullptr, info->descriptionId) : std::wstring{};

    if (category == PrefCategory::Plugins && state.pluginsSelectedPlugin.has_value())
    {
        const std::wstring_view pluginName = PrefsPlugins::GetDisplayName(state.pluginsSelectedPlugin.value());
        if (! pluginName.empty())
        {
            title = std::wstring(pluginName);
        }
        const std::wstring_view pluginDescription = PrefsPlugins::GetDescription(state.pluginsSelectedPlugin.value());
        if (! pluginDescription.empty())
        {
            description = std::wstring(pluginDescription);
        }
    }
    if (title.empty())
    {
        title = LoadStringResource(nullptr, IDS_PREFS_CAPTION);
    }

    if (state.pageTitle)
    {
        SetWindowTextW(state.pageTitle, title.c_str());
    }
    if (state.pageDescription)
    {
        SetWindowTextW(state.pageDescription, description.c_str());
    }

    if (dlg)
    {
        LayoutPreferencesDialog(dlg, state);
    }

    if (state.pageHost)
    {
        LayoutPreferencesPageHost(state.pageHost, state);
        RedrawWindow(state.pageHost, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
    }

    if (category == PrefCategory::General && state.pageHost)
    {
        GeneralPane::Refresh(state.pageHost, state);
    }

    if (category == PrefCategory::Keyboard && state.pageHost)
    {
        KeyboardPane::Refresh(state.pageHost, state);
    }
    if (category == PrefCategory::Panes && state.pageHost)
    {
        PanesPane::Refresh(state.pageHost, state);
    }
    if (category == PrefCategory::Viewers && state.pageHost)
    {
        ViewersPane::Refresh(state.pageHost, state);
    }
    if (category == PrefCategory::Themes && state.pageHost)
    {
        ThemesPane::Refresh(state.pageHost, state);
    }
    if (category == PrefCategory::Plugins && state.pageHost)
    {
        PluginsPane::Refresh(state.pageHost, state);
    }
    if (category == PrefCategory::Advanced && state.pageHost)
    {
        RefreshAdvancedPage(state.pageHost, state);
    }

    if (dlg && state.categoryTree)
    {
        InvalidateRect(state.categoryTree, nullptr, FALSE);
    }
}

void PopulateCategoryTree(HWND dlg, PreferencesDialogState& state) noexcept
{
    state.categoryTree = GetDlgItem(dlg, IDC_PREFS_CATEGORY_LIST);
    if (! state.categoryTree)
    {
        return;
    }

    TreeView_DeleteAllItems(state.categoryTree);
    state.categoryTreeItems.fill(nullptr);
    state.pluginsTreeRoot = nullptr;

    const UINT dpi         = GetDpiForWindow(dlg);
    const int itemHeightPx = std::max(1, ThemedControls::ScaleDip(dpi, 24));
    SendMessageW(state.categoryTree, TVM_SETITEMHEIGHT, static_cast<WPARAM>(itemHeightPx), 0);
    SendMessageW(state.categoryTree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER);

    RECT rc{};
    if (GetWindowRect(state.categoryTree, &rc))
    {
        state.categoryListWidthPx = std::max(0l, rc.right - rc.left);
    }

    for (const auto& c : kCategories)
    {
        std::wstring label = LoadStringResource(nullptr, c.labelId);
        if (label.empty())
        {
            label = L"?";
        }

        TVINSERTSTRUCTW ins{};
        ins.hParent      = TVI_ROOT;
        ins.hInsertAfter = TVI_LAST;
        ins.item.mask    = TVIF_TEXT | TVIF_PARAM;
        ins.item.pszText = label.data();
        ins.item.lParam  = static_cast<LPARAM>(c.id);

        const HTREEITEM inserted = reinterpret_cast<HTREEITEM>(SendMessageW(state.categoryTree, TVM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&ins)));
        const size_t index       = static_cast<size_t>(c.id);
        if (index < state.categoryTreeItems.size())
        {
            state.categoryTreeItems[index] = inserted;
        }
        if (c.id == PrefCategory::Plugins)
        {
            state.pluginsTreeRoot = inserted;
        }
    }

    if (state.pluginsTreeRoot)
    {
        std::vector<PrefsPluginListItem> plugins;
        PrefsPlugins::BuildListItems(plugins);
        for (const PrefsPluginListItem& plugin : plugins)
        {
            const std::wstring_view displayName = PrefsPlugins::GetDisplayName(plugin);
            if (displayName.empty())
            {
                continue;
            }

            std::wstring label(displayName);
            TVINSERTSTRUCTW child{};
            child.hParent      = state.pluginsTreeRoot;
            child.hInsertAfter = TVI_LAST;
            child.item.mask    = TVIF_TEXT | TVIF_PARAM;
            child.item.pszText = label.data();
            child.item.lParam  = PrefsNavTree::EncodePluginData(plugin.type, plugin.index);
            static_cast<void>(SendMessageW(state.categoryTree, TVM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&child)));
        }
    }

    if (state.pluginsTreeRoot)
    {
        TreeView_Expand(state.categoryTree, state.pluginsTreeRoot, TVE_EXPAND);
    }
}

void SelectCategory(HWND dlg, PreferencesDialogState& state, PrefCategory category) noexcept
{
    state.initialCategory = category;
    state.pluginsSelectedPlugin.reset();

    if (! dlg || ! state.categoryTree)
    {
        return;
    }

    const size_t index = static_cast<size_t>(category);
    if (index >= state.categoryTreeItems.size())
    {
        return;
    }

    const HTREEITEM item = state.categoryTreeItems[index];
    if (! item)
    {
        return;
    }

    TreeView_SelectItem(state.categoryTree, item);
    TreeView_EnsureVisible(state.categoryTree, item);
}

void CreatePageControls(HWND dlg, PreferencesDialogState& state) noexcept
{
    state.pageHost = GetDlgItem(dlg, IDC_PREFS_PAGE_HOST);
    if (! state.pageHost)
    {
        return;
    }

    LONG_PTR exStyle = GetWindowLongPtrW(state.pageHost, GWL_EXSTYLE);
    if ((exStyle & WS_EX_CONTROLPARENT) == 0)
    {
        exStyle |= WS_EX_CONTROLPARENT;
        SetWindowLongPtrW(state.pageHost, GWL_EXSTYLE, exStyle);
    }

    LONG_PTR style    = GetWindowLongPtrW(state.pageHost, GWL_STYLE);
    LONG_PTR newStyle = style;
    // Prevent the host from painting over its pane windows (avoids "blank until hover" artifacts).
    // Each pane paints its own themed background/cards.
    newStyle |= WS_CLIPCHILDREN;
    newStyle &= ~WS_HSCROLL;
    newStyle &= ~WS_VSCROLL;
    if (newStyle != style)
    {
        SetWindowLongPtrW(state.pageHost, GWL_STYLE, newStyle);
        SetWindowPos(state.pageHost, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    }

    auto& hostState = static_cast<PreferencesDialogHost&>(state);
    static_cast<void>(hostState._generalPane.EnsureCreated(state.pageHost));
    static_cast<void>(hostState._panesPane.EnsureCreated(state.pageHost));
    static_cast<void>(hostState._viewersPane.EnsureCreated(state.pageHost));
    static_cast<void>(hostState._editorsPane.EnsureCreated(state.pageHost));
    static_cast<void>(hostState._keyboardPane.EnsureCreated(state.pageHost));
    static_cast<void>(hostState._mousePane.EnsureCreated(state.pageHost));
    static_cast<void>(hostState._themesPane.EnsureCreated(state.pageHost));
    static_cast<void>(hostState._pluginsPane.EnsureCreated(state.pageHost));
    static_cast<void>(hostState._advancedPane.EnsureCreated(state.pageHost));

    const HWND generalParent  = hostState._generalPane.Hwnd() ? hostState._generalPane.Hwnd() : state.pageHost;
    const HWND panesParent    = hostState._panesPane.Hwnd() ? hostState._panesPane.Hwnd() : state.pageHost;
    const HWND viewersParent  = hostState._viewersPane.Hwnd() ? hostState._viewersPane.Hwnd() : state.pageHost;
    const HWND editorsParent  = hostState._editorsPane.Hwnd() ? hostState._editorsPane.Hwnd() : state.pageHost;
    const HWND keyboardParent = hostState._keyboardPane.Hwnd() ? hostState._keyboardPane.Hwnd() : state.pageHost;
    const HWND mouseParent    = hostState._mousePane.Hwnd() ? hostState._mousePane.Hwnd() : state.pageHost;
    const HWND themesParent   = hostState._themesPane.Hwnd() ? hostState._themesPane.Hwnd() : state.pageHost;
    const HWND pluginsParent  = hostState._pluginsPane.Hwnd() ? hostState._pluginsPane.Hwnd() : state.pageHost;
    const HWND advancedParent = hostState._advancedPane.Hwnd() ? hostState._advancedPane.Hwnd() : state.pageHost;

    const DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const DWORD wrapStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;

    state.pageTitle = CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, dlg, nullptr, GetModuleHandleW(nullptr), nullptr);

    state.pageDescription = CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, dlg, nullptr, GetModuleHandleW(nullptr), nullptr);

    const HFONT dialogFont = GetDialogFont(dlg);
    EnsureFonts(state, dialogFont);

    GeneralPane::CreateControls(generalParent, state);

    PanesPane::CreateControls(panesParent, state);

    auto populateEnumCombo = [&](const auto& comboLike, std::span<const std::pair<UINT, LPARAM>> options) noexcept
    {
        HWND combo = nullptr;
        if constexpr (requires { comboLike.get(); })
        {
            combo = comboLike.get();
        }
        else
        {
            combo = comboLike;
        }

        if (! combo)
        {
            return;
        }

        SendMessageW(combo, CB_RESETCONTENT, 0, 0);
        for (const auto& option : options)
        {
            const std::wstring text = LoadStringResource(nullptr, option.first);
            const LRESULT index     = SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(text.c_str()));
            if (index != CB_ERR && index != CB_ERRSPACE)
            {
                SendMessageW(combo, CB_SETITEMDATA, static_cast<WPARAM>(index), option.second);
            }
        }

        if (SendMessageW(combo, CB_GETCOUNT, 0, 0) > 0)
        {
            SendMessageW(combo, CB_SETCURSEL, 0, 0);
            PrefsUi::InvalidateComboBox(combo);
        }

        ThemedControls::ApplyThemeToComboBox(combo, state.theme);
    };

    const std::array<std::pair<UINT, LPARAM>, 2> displayOptions = {
        std::pair<UINT, LPARAM>{IDS_PREFS_PANES_OPTION_BRIEF, static_cast<LPARAM>(Common::Settings::FolderDisplayMode::Brief)},
        std::pair<UINT, LPARAM>{IDS_PREFS_PANES_OPTION_DETAILED, static_cast<LPARAM>(Common::Settings::FolderDisplayMode::Detailed)},
    };
    populateEnumCombo(state.panesLeftDisplayCombo, displayOptions);
    populateEnumCombo(state.panesRightDisplayCombo, displayOptions);

    const std::array<std::pair<UINT, LPARAM>, 6> sortByOptions = {
        std::pair<UINT, LPARAM>{IDS_PREFS_PANES_SORT_NAME, static_cast<LPARAM>(Common::Settings::FolderSortBy::Name)},
        std::pair<UINT, LPARAM>{IDS_PREFS_PANES_SORT_EXTENSION, static_cast<LPARAM>(Common::Settings::FolderSortBy::Extension)},
        std::pair<UINT, LPARAM>{IDS_PREFS_PANES_SORT_TIME, static_cast<LPARAM>(Common::Settings::FolderSortBy::Time)},
        std::pair<UINT, LPARAM>{IDS_PREFS_PANES_SORT_SIZE, static_cast<LPARAM>(Common::Settings::FolderSortBy::Size)},
        std::pair<UINT, LPARAM>{IDS_PREFS_PANES_SORT_ATTRIBUTES, static_cast<LPARAM>(Common::Settings::FolderSortBy::Attributes)},
        std::pair<UINT, LPARAM>{IDS_PREFS_PANES_SORT_NONE, static_cast<LPARAM>(Common::Settings::FolderSortBy::None)},
    };
    populateEnumCombo(state.panesLeftSortByCombo, sortByOptions);
    populateEnumCombo(state.panesRightSortByCombo, sortByOptions);

    const std::array<std::pair<UINT, LPARAM>, 2> sortDirOptions = {
        std::pair<UINT, LPARAM>{IDS_PREFS_PANES_OPTION_ASCENDING, static_cast<LPARAM>(Common::Settings::FolderSortDirection::Ascending)},
        std::pair<UINT, LPARAM>{IDS_PREFS_PANES_OPTION_DESCENDING, static_cast<LPARAM>(Common::Settings::FolderSortDirection::Descending)},
    };
    populateEnumCombo(state.panesLeftSortDirCombo, sortDirOptions);
    populateEnumCombo(state.panesRightSortDirCombo, sortDirOptions);

    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC  = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icc);

    // Viewers page (Phase 5).
    ViewersPane::CreateControls(viewersParent, state);
    // Editors page (placeholder).
    EditorsPane::CreateControls(editorsParent, state);
    // Keyboard page (Phase 5).
    KeyboardPane::CreateControls(keyboardParent, state);
    // Mouse page (placeholder).
    MousePane::CreateControls(mouseParent, state);

    // Themes page (Phase 4).
    ThemesPane::CreateControls(themesParent, state);

    // Plugins page (Phase 6 starter).
    PluginsPane::CreateControls(pluginsParent, state);

    // Advanced page (Phase 6 starter).
    AdvancedPane::CreateControls(advancedParent, state);

    if (state.themesThemeCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.themesThemeCombo, state.theme);
    }
    if (state.themesBaseCombo)
    {
        ThemedControls::ApplyThemeToComboBox(state.themesBaseCombo, state.theme);
    }

    if (state.themesColorsList)
    {
        ListView_SetExtendedListViewStyle(state.themesColorsList.get(), LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP);
        ListView_SetBkColor(state.themesColorsList.get(), state.theme.windowBackground);
        ListView_SetTextBkColor(state.themesColorsList.get(), state.theme.windowBackground);
        ListView_SetTextColor(state.themesColorsList.get(), state.theme.menu.text);

        if (! state.theme.systemHighContrast)
        {
            const bool darkBackground = ChooseContrastingTextColor(state.theme.windowBackground) == RGB(255, 255, 255);
            const wchar_t* listTheme  = darkBackground ? L"DarkMode_Explorer" : L"Explorer";
            SetWindowTheme(state.themesColorsList.get(), listTheme, nullptr);
            if (const HWND header = ListView_GetHeader(state.themesColorsList.get()))
            {
                SetWindowTheme(header, listTheme, nullptr);
                InvalidateRect(header, nullptr, TRUE);
            }
        }
        else
        {
            SetWindowTheme(state.themesColorsList.get(), L"", nullptr);
        }

        ThemedControls::EnsureListViewHeaderThemed(state.themesColorsList.get(), state.theme);
    }

    LayoutPreferencesPageHost(state.pageHost, state);
}

LRESULT CALLBACK PreferencesPageHostSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) noexcept
{
    auto* state = reinterpret_cast<PreferencesDialogState*>(dwRefData);
    if (! state)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_NCHITTEST:
        {
            // The page host is a custom control and uses WS_VSCROLL dynamically; ensure standard non-client
            // hit-testing is used so the scrollbar receives mouse interactions.
            return DefSubclassProc(hwnd, msg, wp, lp);
        }
        case WM_NCCALCSIZE:
        case WM_NCPAINT:
        case WM_NCLBUTTONDOWN:
        case WM_NCLBUTTONUP:
        case WM_NCLBUTTONDBLCLK:
        case WM_NCMOUSEMOVE:
        {
            // Ensure standard non-client handling runs (scrollbar sizing/painting/tracking) even though the
            // host uses custom client painting.
            return DefSubclassProc(hwnd, msg, wp, lp);
        }
        case WM_ERASEBKGND: return 1;
        case WM_SETFOCUS:
        {
            const bool forward = (GetKeyState(VK_SHIFT) & 0x8000) == 0;
            if (HWND target = FindFirstOrLastTabStopChild(hwnd, forward))
            {
                SetFocus(target);
                return 0;
            }
            break;
        }
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (! hdc)
            {
                return 0;
            }

            RECT client{};
            GetClientRect(hwnd, &client);
            const int width  = std::max(0l, client.right - client.left);
            const int height = std::max(0l, client.bottom - client.top);

            wil::unique_hdc memDc;
            wil::unique_hbitmap memBmp;
            if (width > 0 && height > 0)
            {
                memDc.reset(CreateCompatibleDC(hdc.get()));
                memBmp.reset(CreateCompatibleBitmap(hdc.get(), width, height));
            }

            if (memDc && memBmp)
            {
                [[maybe_unused]] auto oldBmp = wil::SelectObject(memDc.get(), memBmp.get());
                PaintPageHostBackgroundAndCards(memDc.get(), hwnd, *state);
                BitBlt(hdc.get(), 0, 0, width, height, memDc.get(), 0, 0, SRCCOPY);
            }
            else
            {
                PaintPageHostBackgroundAndCards(hdc.get(), hwnd, *state);
            }

            return 0;
        }
        case WM_PRINTCLIENT:
        {
            HDC hdc = reinterpret_cast<HDC>(wp);
            if (! hdc)
            {
                break;
            }

            PaintPageHostBackgroundAndCards(hdc, hwnd, *state);
            return 0;
        }
        case WndMsg::kPreferencesApplyComboThemeDeferred:
        {
            const HWND combo = reinterpret_cast<HWND>(wp);
            if (! combo || ! IsWindow(combo))
            {
                return 0;
            }

            const UINT comboNotify = static_cast<UINT>(lp);
            if (comboNotify == CBN_DROPDOWN)
            {
                ThemedControls::ApplyThemeToComboBoxDropDown(combo, state->theme);
            }
            else
            {
                ThemedControls::ApplyThemeToComboBox(combo, state->theme);
            }
            ThemedControls::EnsureComboBoxDroppedWidth(combo, GetDpiForWindow(combo));
            return 0;
        }
        case WM_CTLCOLORSTATIC:
        {
            HDC hdc      = reinterpret_cast<HDC>(wp);
            HWND control = reinterpret_cast<HWND>(lp);
            if (! hdc)
            {
                break;
            }

            bool enabled = true;
            if (control)
            {
                enabled = IsWindowEnabled(control) != FALSE;

                // Combo box selection fields sometimes paint via a child static control; match the input background.
                if (const HWND parent = GetParent(control))
                {
                    std::array<wchar_t, 32> className{};
                    const int len = GetClassNameW(parent, className.data(), static_cast<int>(className.size()));
                    if (len > 0 && (_wcsicmp(className.data(), L"ComboBox") == 0 || ThemedControls::IsModernComboBox(parent)))
                    {
                        const bool comboEnabled   = IsWindowEnabled(parent) != FALSE;
                        const bool focused        = comboEnabled && (GetFocus() == parent || SendMessageW(parent, CB_GETDROPPEDSTATE, 0, 0) != 0);
                        const bool themedInputs   = state->inputBrush.get() != nullptr;
                        const COLORREF background = themedInputs ? (comboEnabled ? (focused ? state->inputFocusedBackgroundColor : state->inputBackgroundColor)
                                                                                 : state->inputDisabledBackgroundColor)
                                                                 : state->theme.windowBackground;
                        HBRUSH brush              = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH));
                        if (themedInputs)
                        {
                            if (! comboEnabled)
                            {
                                brush = state->inputDisabledBrush ? state->inputDisabledBrush.get() : state->inputBrush.get();
                            }
                            else if (focused && state->inputFocusedBrush)
                            {
                                brush = state->inputFocusedBrush.get();
                            }
                            else
                            {
                                brush = state->inputBrush.get();
                            }
                        }

                        const COLORREF textColor = comboEnabled ? state->theme.menu.text : GetDisabledTextColor(*state, background);
                        SetBkMode(hdc, OPAQUE);
                        SetBkColor(hdc, background);
                        SetTextColor(hdc, textColor);
                        if (! state->backgroundBrush)
                        {
                            SetDCBrushColor(hdc, background);
                        }
                        return reinterpret_cast<LRESULT>(brush);
                    }
                }
            }

            const COLORREF windowBackground = state->theme.windowBackground;
            COLORREF background             = windowBackground;
            HBRUSH brush                    = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH));

            if (! state->theme.systemHighContrast && state->cardBrush && ! state->pageSettingCards.empty() && control)
            {
                RECT rcControl{};
                if (GetWindowRect(control, &rcControl))
                {
                    MapWindowPoints(nullptr, hwnd, reinterpret_cast<POINT*>(&rcControl), 2);
                    POINT center{};
                    center.x = (rcControl.left + rcControl.right) / 2;
                    center.y = (rcControl.top + rcControl.bottom) / 2;

                    for (const RECT& baseCard : state->pageSettingCards)
                    {
                        RECT card = baseCard;
                        OffsetRect(&card, 0, -state->pageScrollY);
                        if (PtInRect(&card, center) != FALSE)
                        {
                            background = state->cardBackgroundColor;
                            brush      = state->cardBrush.get();
                            break;
                        }
                    }
                }
            }

            const COLORREF textColor = enabled ? state->theme.menu.text : GetDisabledTextColor(*state, background);
            SetBkMode(hdc, OPAQUE);
            SetBkColor(hdc, background);
            SetTextColor(hdc, textColor);
            if (! state->backgroundBrush)
            {
                SetDCBrushColor(hdc, background);
            }
            return reinterpret_cast<LRESULT>(brush);
        }
        case WM_CTLCOLOREDIT:
        {
            HDC hdc = reinterpret_cast<HDC>(wp);
            if (! hdc)
            {
                break;
            }

            const HWND control      = reinterpret_cast<HWND>(lp);
            const bool enabled      = ! control || IsWindowEnabled(control) != FALSE;
            const bool focused      = enabled && control && GetFocus() == control;
            const bool themedInputs = state->inputBrush.get() != nullptr;
            const COLORREF background =
                themedInputs ? (enabled ? (focused ? state->inputFocusedBackgroundColor : state->inputBackgroundColor) : state->inputDisabledBackgroundColor)
                             : state->theme.windowBackground;
            const COLORREF textColor = enabled ? state->theme.menu.text : GetDisabledTextColor(*state, background);
            HBRUSH brush             = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH));
            if (themedInputs)
            {
                if (! enabled)
                {
                    brush = state->inputDisabledBrush ? state->inputDisabledBrush.get() : state->inputBrush.get();
                }
                else if (focused && state->inputFocusedBrush)
                {
                    brush = state->inputFocusedBrush.get();
                }
                else
                {
                    brush = state->inputBrush.get();
                }
            }
            SetBkMode(hdc, OPAQUE);
            SetBkColor(hdc, background);
            SetTextColor(hdc, textColor);
            if (! state->backgroundBrush)
            {
                SetDCBrushColor(hdc, background);
            }
            return reinterpret_cast<LRESULT>(brush);
        }
        case WM_CTLCOLORBTN:
        {
            HDC hdc      = reinterpret_cast<HDC>(wp);
            HWND control = reinterpret_cast<HWND>(lp);
            if (! hdc)
            {
                break;
            }

            const COLORREF windowBackground = state->theme.windowBackground;
            COLORREF background             = windowBackground;
            HBRUSH brush                    = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH));

            if (! state->theme.systemHighContrast && state->cardBrush && ! state->pageSettingCards.empty() && control)
            {
                RECT rcControl{};
                if (GetWindowRect(control, &rcControl))
                {
                    MapWindowPoints(nullptr, hwnd, reinterpret_cast<POINT*>(&rcControl), 2);
                    POINT center{};
                    center.x = (rcControl.left + rcControl.right) / 2;
                    center.y = (rcControl.top + rcControl.bottom) / 2;

                    for (const RECT& baseCard : state->pageSettingCards)
                    {
                        RECT card = baseCard;
                        OffsetRect(&card, 0, -state->pageScrollY);
                        if (PtInRect(&card, center) != FALSE)
                        {
                            background = state->cardBackgroundColor;
                            brush      = state->cardBrush.get();
                            break;
                        }
                    }
                }
            }

            SetBkMode(hdc, OPAQUE);
            SetBkColor(hdc, background);
            SetTextColor(hdc, state->theme.menu.text);
            if (! state->backgroundBrush)
            {
                SetDCBrushColor(hdc, background);
            }
            return reinterpret_cast<LRESULT>(brush);
        }
        case WM_CTLCOLORLISTBOX:
        {
            HDC hdc = reinterpret_cast<HDC>(wp);
            if (! hdc)
            {
                break;
            }

            const HWND control      = reinterpret_cast<HWND>(lp);
            const bool enabled      = ! control || IsWindowEnabled(control) != FALSE;
            const bool themedInputs = state->inputBrush.get() != nullptr;
            const COLORREF background =
                themedInputs ? (enabled ? state->inputBackgroundColor : state->inputDisabledBackgroundColor) : state->theme.windowBackground;
            const COLORREF textColor = enabled ? state->theme.menu.text : GetDisabledTextColor(*state, background);
            HBRUSH brush             = state->backgroundBrush ? state->backgroundBrush.get() : reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH));
            if (themedInputs)
            {
                brush = enabled ? state->inputBrush.get() : (state->inputDisabledBrush ? state->inputDisabledBrush.get() : state->inputBrush.get());
            }
            SetBkMode(hdc, OPAQUE);
            SetBkColor(hdc, background);
            SetTextColor(hdc, textColor);
            if (! state->backgroundBrush)
            {
                SetDCBrushColor(hdc, background);
            }
            return reinterpret_cast<LRESULT>(brush);
        }
        case WM_MEASUREITEM:
        {
            auto* mis = reinterpret_cast<MEASUREITEMSTRUCT*>(lp);
            if (mis)
            {
                const LRESULT handled = KeyboardPane::OnMeasureList(mis, *state);
                if (handled != 0)
                {
                    return handled;
                }
                const LRESULT handledViewers = ViewersPane::OnMeasureList(mis, *state);
                if (handledViewers != 0)
                {
                    return handledViewers;
                }
                const LRESULT handledThemes = ThemesPane::OnMeasureColorsList(mis, *state);
                if (handledThemes != 0)
                {
                    return handledThemes;
                }
            }
            break;
        }
        case WM_DRAWITEM:
        {
            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (! dis)
            {
                break;
            }

            {
                const LRESULT handled = ThemesPane::OnDrawColorSwatch(dis, *state);
                if (handled != 0)
                {
                    return handled;
                }
            }

            if (dis->CtlType == ODT_LISTVIEW && dis->CtlID == static_cast<UINT>(IDC_PREFS_KEYBOARD_LIST))
            {
                const LRESULT handled = KeyboardPane::OnDrawList(dis, *state);
                if (handled != 0)
                {
                    return handled;
                }
                break;
            }
            if (dis->CtlType == ODT_LISTVIEW && dis->CtlID == static_cast<UINT>(IDC_PREFS_VIEWERS_LIST))
            {
                const LRESULT handled = ViewersPane::OnDrawList(dis, *state);
                if (handled != 0)
                {
                    return handled;
                }
                break;
            }
            if (dis->CtlType == ODT_LISTVIEW && dis->CtlID == static_cast<UINT>(IDC_PREFS_THEMES_COLORS_LIST))
            {
                const LRESULT handled = ThemesPane::OnDrawColorsList(dis, *state);
                if (handled != 0)
                {
                    return handled;
                }
                break;
            }

            if (dis->CtlType != ODT_BUTTON)
            {
                break;
            }

            if (! dis->hwndItem || ! IsWindow(dis->hwndItem))
            {
                break;
            }

            const LONG_PTR style = GetWindowLongPtrW(dis->hwndItem, GWL_STYLE);
            if ((style & BS_TYPEMASK) != BS_OWNERDRAW)
            {
                break;
            }

            if (const auto* pluginControls = FindPluginDetailsToggleControls(*state, dis->hwndItem))
            {
                const bool toggledOn   = GetWindowLongPtrW(dis->hwndItem, GWLP_USERDATA) != 0;
                const COLORREF surface = ThemedControls::GetControlSurfaceColor(state->theme);
                const HFONT boldFont   = state->boldFont ? state->boldFont.get() : nullptr;

                const std::wstring onText  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
                const std::wstring offText = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
                std::wstring_view onLabel  = onText;
                std::wstring_view offLabel = offText;

                if (pluginControls->field.type == PrefsPluginConfigFieldType::Option && pluginControls->field.choices.size() >= 2)
                {
                    const auto& choices   = pluginControls->field.choices;
                    const size_t onIndex  = std::min(pluginControls->toggleOnChoiceIndex, choices.size() - 1);
                    const size_t offIndex = std::min(pluginControls->toggleOffChoiceIndex, choices.size() - 1);

                    onLabel  = choices[onIndex].label.empty() ? std::wstring_view(choices[onIndex].value) : std::wstring_view(choices[onIndex].label);
                    offLabel = choices[offIndex].label.empty() ? std::wstring_view(choices[offIndex].value) : std::wstring_view(choices[offIndex].label);
                }

                ThemedControls::DrawThemedSwitchToggle(*dis, state->theme, surface, boldFont, onLabel, offLabel, toggledOn);
                return TRUE;
            }

            if (dis->CtlID == IDC_PREFS_GENERAL_MENUBAR_TOGGLE || dis->CtlID == IDC_PREFS_GENERAL_FUNCTIONBAR_TOGGLE ||
                dis->CtlID == IDC_PREFS_GENERAL_SPLASH_TOGGLE || dis->CtlID == IDC_PREFS_PANES_LEFT_STATUSBAR_TOGGLE ||
                dis->CtlID == IDC_PREFS_PANES_RIGHT_STATUSBAR_TOGGLE || dis->CtlID == IDC_PREFS_ADV_CONNECTIONS_BYPASS_HELLO_TOGGLE ||
                dis->CtlID == IDC_PREFS_ADV_MONITOR_TOOLBAR_TOGGLE || dis->CtlID == IDC_PREFS_ADV_MONITOR_LINE_NUMBERS_TOGGLE ||
                dis->CtlID == IDC_PREFS_ADV_MONITOR_ALWAYS_ON_TOP_TOGGLE || dis->CtlID == IDC_PREFS_ADV_MONITOR_SHOW_IDS_TOGGLE ||
                dis->CtlID == IDC_PREFS_ADV_MONITOR_AUTO_SCROLL_TOGGLE || dis->CtlID == IDC_PREFS_ADV_MONITOR_FILTER_TEXT_TOGGLE ||
                dis->CtlID == IDC_PREFS_ADV_MONITOR_FILTER_ERROR_TOGGLE || dis->CtlID == IDC_PREFS_ADV_MONITOR_FILTER_WARNING_TOGGLE ||
                dis->CtlID == IDC_PREFS_ADV_MONITOR_FILTER_INFO_TOGGLE || dis->CtlID == IDC_PREFS_ADV_MONITOR_FILTER_DEBUG_TOGGLE ||
                dis->CtlID == IDC_PREFS_ADV_FILEOPS_DIAG_INFO_TOGGLE || dis->CtlID == IDC_PREFS_ADV_FILEOPS_DIAG_DEBUG_TOGGLE)
            {
                const bool toggledOn        = GetWindowLongPtrW(dis->hwndItem, GWLP_USERDATA) != 0;
                const COLORREF surface      = ThemedControls::GetControlSurfaceColor(state->theme);
                const HFONT boldFont        = state->boldFont ? state->boldFont.get() : nullptr;
                const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
                const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
                ThemedControls::DrawThemedSwitchToggle(*dis, state->theme, surface, boldFont, onLabel, offLabel, toggledOn);
                return TRUE;
            }

            if (dis->CtlID == IDC_PREFS_PANES_LEFT_DISPLAY_TOGGLE || dis->CtlID == IDC_PREFS_PANES_RIGHT_DISPLAY_TOGGLE)
            {
                const bool toggledOn             = GetWindowLongPtrW(dis->hwndItem, GWLP_USERDATA) != 0;
                const COLORREF surface           = ThemedControls::GetControlSurfaceColor(state->theme);
                const HFONT boldFont             = state->boldFont ? state->boldFont.get() : nullptr;
                const std::wstring briefLabel    = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_BRIEF);
                const std::wstring detailedLabel = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DETAILED);
                ThemedControls::DrawThemedSwitchToggle(*dis, state->theme, surface, boldFont, briefLabel, detailedLabel, toggledOn);
                return TRUE;
            }

            if (dis->CtlID == IDC_PREFS_PANES_LEFT_SORTDIR_TOGGLE || dis->CtlID == IDC_PREFS_PANES_RIGHT_SORTDIR_TOGGLE)
            {
                const bool toggledOn         = GetWindowLongPtrW(dis->hwndItem, GWLP_USERDATA) != 0;
                const COLORREF surface       = ThemedControls::GetControlSurfaceColor(state->theme);
                const HFONT boldFont         = state->boldFont ? state->boldFont.get() : nullptr;
                const std::wstring ascLabel  = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_ASCENDING);
                const std::wstring descLabel = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DESCENDING);
                ThemedControls::DrawThemedSwitchToggle(*dis, state->theme, surface, boldFont, ascLabel, descLabel, toggledOn);
                return TRUE;
            }

            ThemedControls::DrawThemedPushButton(*dis, state->theme);
            return TRUE;
        }
        case WM_COMMAND:
        {
            const UINT controlId = LOWORD(wp);
            const UINT notify    = HIWORD(wp);
            if (notify == BN_SETFOCUS || notify == EN_SETFOCUS || notify == CBN_SETFOCUS)
            {
                HWND control = reinterpret_cast<HWND>(lp);
                if (control)
                {
                    PrefsPaneHost::EnsureControlVisible(hwnd, *state, control);
                    InvalidateRect(control, nullptr, TRUE);
                }
            }
            if (notify == BN_KILLFOCUS || notify == EN_KILLFOCUS || notify == CBN_KILLFOCUS)
            {
                HWND control = reinterpret_cast<HWND>(lp);
                if (control)
                {
                    InvalidateRect(control, nullptr, TRUE);
                }
            }

            if (notify == CBN_DROPDOWN || notify == CBN_CLOSEUP)
            {
                HWND combo = reinterpret_cast<HWND>(lp);
                if (combo)
                {
                    PostMessageW(hwnd,
                                 WndMsg::kPreferencesApplyComboThemeDeferred,
                                 reinterpret_cast<WPARAM>(combo),
                                 static_cast<LPARAM>(notify));
                }
            }

            if (KeyboardPane::HandleCommand(hwnd, *state, controlId, notify, reinterpret_cast<HWND>(lp)))
            {
                return 0;
            }

            if (ViewersPane::HandleCommand(hwnd, *state, controlId, notify, reinterpret_cast<HWND>(lp)))
            {
                return 0;
            }

            if (ThemesPane::HandleCommand(hwnd, *state, controlId, notify, reinterpret_cast<HWND>(lp)))
            {
                return 0;
            }

            if (AdvancedPane::HandleCommand(hwnd, *state, controlId, notify, reinterpret_cast<HWND>(lp)))
            {
                return 0;
            }

            if (PanesPane::HandleCommand(hwnd, *state, controlId, notify, reinterpret_cast<HWND>(lp)))
            {
                return 0;
            }

            if (GeneralPane::HandleCommand(hwnd, *state, controlId, notify, reinterpret_cast<HWND>(lp)))
            {
                return 0;
            }

            if (PluginsPane::HandleCommand(hwnd, *state, controlId, notify, reinterpret_cast<HWND>(lp)))
            {
                return 0;
            }

            break;
        }
        case WM_NOTIFY:
        {
            auto* hdr = reinterpret_cast<NMHDR*>(lp);
            if (! hdr)
            {
                break;
            }

            LRESULT result = 0;
            if (ThemesPane::HandleNotify(hwnd, *state, hdr, result) || ViewersPane::HandleNotify(hwnd, *state, hdr, result) ||
                KeyboardPane::HandleNotify(hwnd, *state, hdr, result) || PluginsPane::HandleNotify(hwnd, *state, hdr, result))
            {
                return result;
            }

            break;
        }
        case WM_VSCROLL:
        {
            if (! state || state->pageScrollMaxY <= 0)
            {
                break;
            }

            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_ALL;
            if (! GetScrollInfo(hwnd, SB_VERT, &si))
            {
                break;
            }

            int newPos         = state->pageScrollY;
            const UINT dpi     = GetDpiForWindow(hwnd);
            const int lineStep = std::max(1, ThemedControls::ScaleDip(dpi, 24));

            switch (LOWORD(wp))
            {
                case SB_LINEUP: newPos -= lineStep; break;
                case SB_LINEDOWN: newPos += lineStep; break;
                case SB_PAGEUP: newPos -= static_cast<int>(si.nPage); break;
                case SB_PAGEDOWN: newPos += static_cast<int>(si.nPage); break;
                case SB_TOP: newPos = 0; break;
                case SB_BOTTOM: newPos = state->pageScrollMaxY; break;
                case SB_THUMBPOSITION:
                case SB_THUMBTRACK: newPos = si.nTrackPos; break;
                default: break;
            }

            PrefsPaneHost::ScrollTo(hwnd, *state, newPos);
            return 0;
        }
        case WM_MOUSEWHEEL:
        {
            if (! state)
            {
                break;
            }

            if (HandlePageHostMouseWheel(hwnd, *state, wp))
            {
                return 0;
            }
            break;
        }
        case WM_SIZE:
        {
            if (state->pageHostIgnoreSize)
            {
                return DefSubclassProc(hwnd, msg, wp, lp);
            }

            const LRESULT result = DefSubclassProc(hwnd, msg, wp, lp);
            SendMessageW(hwnd, WM_SETREDRAW, FALSE, 0);
            LayoutPreferencesPageHost(hwnd, *state);
            SendMessageW(hwnd, WM_SETREDRAW, TRUE, 0);
            RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
            return result;
        }
        case WM_NCDESTROY:
        {
            RemoveWindowSubclass(hwnd, PreferencesPageHostSubclassProc, uIdSubclass);
            break;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

INT_PTR OnInitDialog(HWND dlg, PreferencesDialogState* state)
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    SetState(dlg, state);

    SetWindowTextW(dlg, LoadStringResource(nullptr, IDS_PREFS_CAPTION).c_str());
    if (HWND ok = GetDlgItem(dlg, IDOK))
    {
        SetWindowTextW(ok, LoadStringResource(nullptr, IDS_BTN_OK).c_str());
    }
    if (HWND cancel = GetDlgItem(dlg, IDCANCEL))
    {
        SetWindowTextW(cancel, LoadStringResource(nullptr, IDS_BTN_CANCEL).c_str());
    }
    if (HWND apply = GetDlgItem(dlg, IDC_PREFS_APPLY))
    {
        SetWindowTextW(apply, LoadStringResource(nullptr, IDS_BTN_APPLY).c_str());
    }

    ApplyTitleBarTheme(dlg, state->theme, GetActiveWindow() == dlg);

    state->backgroundBrush.reset(CreateSolidBrush(state->theme.windowBackground));
    state->cardBackgroundColor = ThemedControls::GetControlSurfaceColor(state->theme);

    state->inputBackgroundColor = ThemedControls::BlendColor(state->cardBackgroundColor, state->theme.windowBackground, state->theme.dark ? 50 : 30, 255);
    state->inputFocusedBackgroundColor = ThemedControls::BlendColor(state->inputBackgroundColor, state->theme.menu.text, state->theme.dark ? 20 : 16, 255);
    state->inputDisabledBackgroundColor =
        ThemedControls::BlendColor(state->theme.windowBackground, state->inputBackgroundColor, state->theme.dark ? 70 : 40, 255);
    state->cardBrush.reset();
    state->inputBrush.reset();
    state->inputFocusedBrush.reset();
    state->inputDisabledBrush.reset();
    if (! state->theme.systemHighContrast)
    {
        state->cardBrush.reset(CreateSolidBrush(state->cardBackgroundColor));
        state->inputBrush.reset(CreateSolidBrush(state->inputBackgroundColor));
        state->inputFocusedBrush.reset(CreateSolidBrush(state->inputFocusedBackgroundColor));
        state->inputDisabledBrush.reset(CreateSolidBrush(state->inputDisabledBackgroundColor));
    }

    RECT initial{};
    if (GetWindowRect(dlg, &initial))
    {
        state->minTrackSizePx.cx = std::max(0l, initial.right - initial.left);
        state->minTrackSizePx.cy = std::max(0l, initial.bottom - initial.top);

        HWND ok     = GetDlgItem(dlg, IDOK);
        HWND cancel = GetDlgItem(dlg, IDCANCEL);
        HWND apply  = GetDlgItem(dlg, IDC_PREFS_APPLY);

        RECT client{};
        GetClientRect(dlg, &client);
        const int windowWidth     = std::max(0, static_cast<int>(state->minTrackSizePx.cx));
        const int clientWidth     = std::max(0l, client.right - client.left);
        const int windowHeight    = std::max(0, static_cast<int>(state->minTrackSizePx.cy));
        const int clientHeight    = std::max(0l, client.bottom - client.top);
        const int nonClientWidth  = std::max(0, windowWidth - clientWidth);
        const int nonClientHeight = std::max(0, windowHeight - clientHeight);

        if (ok && cancel && apply)
        {
            const UINT dpi       = GetDpiForWindow(dlg);
            const int margin     = ThemedControls::ScaleDip(dpi, 8);
            const int gapX       = ThemedControls::ScaleDip(dpi, 8);
            const int minGapX    = ThemedControls::ScaleDip(dpi, 4);
            const int buttonPadX = ThemedControls::ScaleDip(dpi, 12);

            const auto measureButtonMinWidth = [&](HWND button) noexcept
            {
                HFONT font = reinterpret_cast<HFONT>(SendMessageW(button, WM_GETFONT, 0, 0));
                if (! font)
                {
                    font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
                }

                const std::wstring text = PrefsUi::GetWindowTextString(button);
                const int textW         = ThemedControls::MeasureTextWidth(button, font, text);
                return std::max(ThemedControls::ScaleDip(dpi, 60), textW + 2 * buttonPadX);
            };

            const int okMin     = measureButtonMinWidth(ok);
            const int cancelMin = measureButtonMinWidth(cancel);
            const int applyMin  = measureButtonMinWidth(apply);

            const int minButtonsClientWidth = std::max(0, (2 * margin) + okMin + cancelMin + applyMin + (2 * minGapX));

            const int listMinWidth          = ThemedControls::ScaleDip(dpi, 72);
            const int hostMinWidth          = ThemedControls::ScaleDip(dpi, 140);
            const int minContentClientWidth = std::max(0, (2 * margin) + listMinWidth + gapX + hostMinWidth);

            const int minClientWidth = std::max(minButtonsClientWidth, minContentClientWidth);
            state->minTrackSizePx.cx = std::max(0, minClientWidth + nonClientWidth);

            RECT okRect{};
            RECT cancelRect{};
            RECT applyRect{};
            GetWindowRect(ok, &okRect);
            GetWindowRect(cancel, &cancelRect);
            GetWindowRect(apply, &applyRect);

            const int okHeight     = std::max(0l, okRect.bottom - okRect.top);
            const int cancelHeight = std::max(0l, cancelRect.bottom - cancelRect.top);
            const int applyHeight  = std::max(0l, applyRect.bottom - applyRect.top);
            int buttonHeight       = std::max({okHeight, cancelHeight, applyHeight});
            if (buttonHeight <= 0)
            {
                buttonHeight = ThemedControls::ScaleDip(dpi, 26);
            }

            // Content area = left list + page host (scrolls vertically). Keep the minimum height small enough
            // to allow the user to shrink the dialog while still keeping the buttons reachable.
            const int minContentClientHeight = ThemedControls::ScaleDip(dpi, 160);
            const int minClientHeight        = std::max(0, minContentClientHeight + buttonHeight + 3 * margin);
            state->minTrackSizePx.cy         = std::max(0, minClientHeight + nonClientHeight);
        }
    }

    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC  = ICC_TREEVIEW_CLASSES;
    InitCommonControlsEx(&icc);

    PopulateCategoryTree(dlg, *state);

    if (! state->theme.systemHighContrast)
    {
        ThemedControls::EnableOwnerDrawButton(dlg, IDOK);
        ThemedControls::EnableOwnerDrawButton(dlg, IDCANCEL);
        ThemedControls::EnableOwnerDrawButton(dlg, IDC_PREFS_APPLY);
    }

    UpdateApplyButton(dlg, *state);

    CreatePageControls(dlg, *state);
    ApplyThemeToPreferencesDialog(dlg, *state, state->theme);

    if (state->pageHost)
    {
        SetWindowSubclass(state->pageHost, PreferencesPageHostSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(state));
    }

    InstallWheelRoutingSubclasses(dlg, *state);

    LayoutPreferencesDialog(dlg, *state);

    SelectCategory(dlg, *state, state->initialCategory);
    return TRUE;
}

INT_PTR OnCtlColorDialog(PreferencesDialogState* state)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnCtlColorStatic(PreferencesDialogState* state, HDC hdc, HWND control)
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    COLORREF textColor = state->theme.menu.text;
    if (control)
    {
        if (IsWindowEnabled(control) == FALSE)
        {
            textColor = GetDisabledTextColor(*state, state->theme.windowBackground);
        }
    }

    if (! state->theme.systemHighContrast)
    {
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, textColor);
        return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
    }

    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, state->theme.windowBackground);
    SetTextColor(hdc, textColor);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnCtlColorListBox(PreferencesDialogState* state, HDC hdc, HWND listBox)
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    const bool isCategoryTree = listBox && state->categoryTree && listBox == state->categoryTree;
    const bool useInputBrush  = ! isCategoryTree && state->inputBrush && ! state->theme.systemHighContrast;

    const COLORREF background = useInputBrush ? state->inputBackgroundColor : state->theme.windowBackground;
    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, background);
    SetTextColor(hdc, state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(useInputBrush ? state->inputBrush.get() : state->backgroundBrush.get());
}

INT_PTR OnCommand(HWND dlg, PreferencesDialogState* state, UINT commandId, [[maybe_unused]] UINT notifyCode, [[maybe_unused]] HWND hwndCtl)
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    switch (commandId)
    {
        case IDOK:
            if (state->dirty)
            {
                CommitAndApply(dlg, *state);
                if (state->dirty)
                {
                    return TRUE;
                }
            }
            g_preferencesDialog.reset();
            return TRUE;
        case IDC_PREFS_APPLY:
            if (state->dirty)
            {
                CommitAndApply(dlg, *state);
            }
            return TRUE;
        case IDCANCEL:
            if (state->previewApplied && state->settings)
            {
                Common::Settings::Settings restored = *state->settings;
                restored.theme                      = state->baselineSettings.theme;
                *state->settings                    = std::move(restored);

                if (state->owner)
                {
                    PostMessageW(state->owner, WndMsg::kSettingsApplied, 0, 0);
                }
            }
            g_preferencesDialog.reset();
            return TRUE;
    }

    return FALSE;
}

INT_PTR CALLBACK PreferencesDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* state = GetState(dlg);

    switch (msg)
    {
        case WM_INITDIALOG: return OnInitDialog(dlg, reinterpret_cast<PreferencesDialogState*>(lp));
        case WM_CLOSE: return OnCommand(dlg, state, IDOK, 0, nullptr);
        case WM_ERASEBKGND:
            if (state && state->backgroundBrush && wp)
            {
                RECT rc{};
                if (GetClientRect(dlg, &rc))
                {
                    FillRect(reinterpret_cast<HDC>(wp), &rc, state->backgroundBrush.get());
                    return TRUE;
                }
            }
            break;
        case WM_CTLCOLORDLG: return OnCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnCtlColorStatic(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLORLISTBOX: return OnCtlColorListBox(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_NOTIFY:
        {
            if (! state)
            {
                break;
            }

            auto* hdr = reinterpret_cast<NMHDR*>(lp);
            if (! hdr)
            {
                break;
            }

            if (! state->categoryTree || hdr->hwndFrom != state->categoryTree)
            {
                break;
            }

            if (hdr->code == TVN_SELCHANGEDW)
            {
                auto* nmtv = reinterpret_cast<NMTREEVIEWW*>(lp);
                if (! nmtv || ! nmtv->itemNew.hItem)
                {
                    return TRUE;
                }

                TVITEMW item{};
                item.mask  = TVIF_PARAM;
                item.hItem = nmtv->itemNew.hItem;
                if (! TreeView_GetItem(state->categoryTree, &item))
                {
                    return TRUE;
                }

                PrefsPluginListItem pluginItem{};
                if (PrefsNavTree::TryDecodePluginData(item.lParam, pluginItem))
                {
                    state->pluginsSelectedPlugin = pluginItem;
                    UpdatePageText(dlg, *state, PrefCategory::Plugins);
                    return TRUE;
                }

                state->pluginsSelectedPlugin.reset();
                const auto category = static_cast<PrefCategory>(item.lParam);
                UpdatePageText(dlg, *state, category);
                return TRUE;
            }

            if (hdr->code == NM_CUSTOMDRAW)
            {
                auto* cd = reinterpret_cast<NMTVCUSTOMDRAW*>(lp);
                if (! cd)
                {
                    break;
                }

                switch (cd->nmcd.dwDrawStage)
                {
                    case CDDS_PREPAINT: return CDRF_NOTIFYITEMDRAW;
                    case CDDS_ITEMPREPAINT:
                    {
                        const bool selected    = (cd->nmcd.uItemState & CDIS_SELECTED) != 0;
                        const bool disabled    = (cd->nmcd.uItemState & CDIS_DISABLED) != 0;
                        const bool treeFocused = GetFocus() == state->categoryTree;

                        const HWND root         = GetAncestor(state->categoryTree, GA_ROOT);
                        const bool windowActive = root && GetActiveWindow() == root;

                        COLORREF bg   = state->theme.systemHighContrast ? GetSysColor(COLOR_WINDOW) : state->theme.windowBackground;
                        COLORREF text = state->theme.systemHighContrast ? GetSysColor(COLOR_WINDOWTEXT)
                                                                        : (disabled ? state->theme.menu.disabledText : state->theme.menu.text);

                        if (selected)
                        {
                            COLORREF selBg = state->theme.systemHighContrast ? GetSysColor(COLOR_HIGHLIGHT) : state->theme.menu.selectionBg;
                            std::array<wchar_t, 128> itemText{};
                            if (! state->theme.highContrast && state->theme.menu.rainbowMode)
                            {
                                TVITEMW tvi{};
                                tvi.mask       = TVIF_TEXT;
                                tvi.hItem      = reinterpret_cast<HTREEITEM>(cd->nmcd.dwItemSpec);
                                tvi.pszText    = itemText.data();
                                tvi.cchTextMax = static_cast<int>(itemText.size());
                                if (TreeView_GetItem(state->categoryTree, &tvi))
                                {
                                    const std::wstring_view seed(itemText.data());
                                    if (! seed.empty())
                                    {
                                        selBg = RainbowMenuSelectionColor(seed, state->theme.menu.darkBase);
                                    }
                                }
                            }

                            COLORREF selText = state->theme.systemHighContrast ? GetSysColor(COLOR_HIGHLIGHTTEXT) : state->theme.menu.selectionText;
                            if (! state->theme.highContrast && state->theme.menu.rainbowMode)
                            {
                                selText = ChooseContrastingTextColor(selBg);
                            }

                            if (windowActive && treeFocused)
                            {
                                bg   = selBg;
                                text = selText;
                            }
                            else if (! state->theme.highContrast)
                            {
                                const int denom = state->theme.menu.darkBase ? 2 : 3;
                                bg              = ThemedControls::BlendColor(state->theme.windowBackground, selBg, 1, denom);
                                text            = ChooseContrastingTextColor(bg);
                            }
                            else
                            {
                                bg   = selBg;
                                text = selText;
                            }
                        }

                        cd->clrTextBk = bg;
                        cd->clrText   = text;
                        return CDRF_DODEFAULT;
                    }
                }
            }

            break;
        }
        case WM_ACTIVATE:
            if (state)
            {
                if (state->categoryTree)
                {
                    InvalidateRect(state->categoryTree, nullptr, FALSE);
                }
                if (state->pageHost)
                {
                    RedrawWindow(state->pageHost, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
                }

                const auto invalidateList = [&](const auto& listLike) noexcept
                {
                    HWND list = nullptr;
                    if constexpr (requires { listLike.get(); })
                    {
                        list = listLike.get();
                    }
                    else
                    {
                        list = listLike;
                    }

                    if (! list)
                    {
                        return;
                    }
                    InvalidateRect(list, nullptr, FALSE);
                    if (const HWND header = ListView_GetHeader(list))
                    {
                        InvalidateRect(header, nullptr, TRUE);
                    }
                };

                invalidateList(state->keyboardList);
                invalidateList(state->viewersList);
                invalidateList(state->themesColorsList);
            }
            return FALSE;
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->theme, wp != FALSE);
            }
            return FALSE;
        case WM_GETMINMAXINFO:
        {
            auto* info = reinterpret_cast<MINMAXINFO*>(lp);
            if (! info)
            {
                break;
            }

            bool handled = false;
            if (state && state->minTrackSizePx.cx > 0 && state->minTrackSizePx.cy > 0)
            {
                info->ptMinTrackSize.x = state->minTrackSizePx.cx;
                info->ptMinTrackSize.y = state->minTrackSizePx.cy;
                handled                = true;
            }

            // Custom "maximize vertically": keep the current width, but expand to the monitor work-area height.
            MONITORINFO mi{};
            mi.cbSize              = sizeof(mi);
            const HMONITOR monitor = MonitorFromWindow(dlg, MONITOR_DEFAULTTONEAREST);
            if (monitor && GetMonitorInfoW(monitor, &mi))
            {
                RECT windowRc{};
                if (GetWindowRect(dlg, &windowRc))
                {
                    const int workWidth    = std::max(0l, mi.rcWork.right - mi.rcWork.left);
                    const int workHeight   = std::max(0l, mi.rcWork.bottom - mi.rcWork.top);
                    const int currentWidth = std::max(0l, windowRc.right - windowRc.left);
                    const int desiredWidth = std::clamp(currentWidth, 0, workWidth);
                    const int maxLeft      = static_cast<int>(mi.rcWork.right) - desiredWidth;
                    const int desiredLeft  = std::clamp(static_cast<int>(windowRc.left), static_cast<int>(mi.rcWork.left), maxLeft);

                    info->ptMaxSize.x     = desiredWidth;
                    info->ptMaxSize.y     = workHeight;
                    info->ptMaxPosition.x = desiredLeft - mi.rcMonitor.left;
                    info->ptMaxPosition.y = mi.rcWork.top - mi.rcMonitor.top;
                    handled               = true;
                }
            }

            if (handled)
            {
                return TRUE;
            }
            break;
        }
        case WM_DPICHANGED:
            if (state)
            {
                const UINT dpi              = static_cast<UINT>(HIWORD(wp));
                const RECT* const suggested = reinterpret_cast<const RECT*>(lp);
                if (suggested)
                {
                    const int width  = static_cast<int>(std::max(0l, suggested->right - suggested->left));
                    const int height = static_cast<int>(std::max(0l, suggested->bottom - suggested->top));
                    SetWindowPos(dlg, nullptr, suggested->left, suggested->top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
                }

                state->uiFont   = CreateMenuFontForDpi(dpi);
                HFONT fontToUse = state->uiFont.get();
                if (! fontToUse)
                {
                    fontToUse = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
                }

                SendMessageW(dlg, WM_SETFONT, reinterpret_cast<WPARAM>(fontToUse), TRUE);
                EnumChildWindows(dlg, SetDialogChildFontProc, reinterpret_cast<LPARAM>(fontToUse));

                state->italicFont.reset();
                state->boldFont.reset();
                state->titleFont.reset();

                if (state->categoryTree)
                {
                    const int itemHeightPx = std::max(1, ThemedControls::ScaleDip(dpi, 24));
                    SendMessageW(state->categoryTree, TVM_SETITEMHEIGHT, static_cast<WPARAM>(itemHeightPx), 0);
                }

                LayoutPreferencesDialog(dlg, *state);
                if (state->pageHost)
                {
                    LayoutPreferencesPageHost(state->pageHost, *state);
                }
                RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN | RDW_UPDATENOW);
            }
            return TRUE;
        case WM_SIZE:
            if (state)
            {
                LayoutPreferencesDialog(dlg, *state);
                InvalidateRect(dlg, nullptr, TRUE);
            }
            return TRUE;
        case WM_SETTINGCHANGE:
        case WM_THEMECHANGED:
        case WM_DWMCOLORIZATIONCOLORCHANGED:
        case WM_SYSCOLORCHANGE:
            if (state)
            {
                RefreshPreferencesDialogTheme(dlg, *state);
            }
            return TRUE;
        case WM_EXITSIZEMOVE:
            if (state)
            {
                LayoutPreferencesDialog(dlg, *state);
                RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            }
            return TRUE;
        case WM_DRAWITEM:
        {
            if (! state)
            {
                break;
            }

            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (! dis)
            {
                break;
            }

            if (dis->CtlType == ODT_BUTTON)
            {
                ThemedControls::DrawThemedPushButton(*dis, state->theme);
                return TRUE;
            }

            break;
        }
        case WM_COMMAND: return OnCommand(dlg, state, LOWORD(wp), HIWORD(wp), reinterpret_cast<HWND>(lp));
        case WM_NCDESTROY:
        {
            if (state)
            {
                std::unique_ptr<PreferencesDialogHost> stateOwner;
                stateOwner.reset(static_cast<PreferencesDialogHost*>(state));

                if (state->settings)
                {
                    WindowPlacementPersistence::Save(*state->settings, kPreferencesWindowId, dlg);

                    const Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(*state->settings);
                    const HRESULT saveHr                            = Common::Settings::SaveSettings(state->appId, settingsToSave);
                    if (FAILED(saveHr))
                    {
                        const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(state->appId);
                        Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
                    }
                }

                if (state->pageHost)
                {
                    RemoveWindowSubclass(state->pageHost, PreferencesPageHostSubclassProc, 1u);
                }

                RemoveWindowSubclass(dlg, PreferencesWheelRouteSubclassProc, kPrefsWheelRouteSubclassId);
                EnumChildWindows(
                    dlg,
                    [](HWND child, LPARAM) noexcept -> BOOL
                    {
                        RemoveWindowSubclass(child, PreferencesWheelRouteSubclassProc, kPrefsWheelRouteSubclassId);
                        return TRUE;
                    },
                    0);

                SetState(dlg, nullptr);
                if (g_preferencesDialog.get() == dlg)
                {
                    g_preferencesDialog.release();
                }
            }
            return FALSE;
        }
    }

    return FALSE;
}
} // namespace

[[nodiscard]] bool
PreferencesDialog::Show(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme, PrefCategory initialCategory) noexcept
{
    if (const HWND existing = g_preferencesDialog.get())
    {
        if (! IsWindow(existing))
        {
            g_preferencesDialog.release();
        }
        else
        {
            if (IsIconic(existing))
            {
                ShowWindow(existing, SW_RESTORE);
            }
            else
            {
                ShowWindow(existing, SW_SHOW);
            }
            SetForegroundWindow(existing);
            if (auto* state = GetState(existing))
            {
                SelectCategory(existing, *state, initialCategory);
            }
            return true;
        }
    }

    auto statePtr = std::make_unique<PreferencesDialogHost>();
    auto* state   = statePtr.get();

    HWND effectiveOwner = owner;
    if (effectiveOwner && IsWindow(effectiveOwner))
    {
        effectiveOwner = GetAncestor(effectiveOwner, GA_ROOT);
    }
    else
    {
        effectiveOwner = nullptr;
    }

    state->owner           = effectiveOwner;
    state->settings        = &settings;
    state->appId           = std::wstring(appId);
    state->theme           = theme;
    state->initialCategory = initialCategory;

    if (! EnsurePrefsPageHostClassRegistered())
    {
        return false;
    }

    state->baselineSettings = settings;
    state->workingSettings  = settings;

    // Ensure mainMenu is explicitly set with defaults if not present
    // This prevents function bar from being reset when applying preferences
    if (! state->workingSettings.mainMenu.has_value())
    {
        state->workingSettings.mainMenu = Common::Settings::MainMenuState{};
    }

    // Load schema fields for UI generation
    const std::wstring schemaPath = std::wstring(appId) + L".settings.schema.json";
    state->schemaFields           = SettingsSchemaParser::LoadAndParseSettingsSchema(schemaPath);

    SetDirty(nullptr, *state);

    const HWND dlg =
        CreateDialogParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_PREFERENCES), nullptr, PreferencesDialogProc, reinterpret_cast<LPARAM>(state));

    if (! dlg)
    {
        return false;
    }

    g_preferencesDialog.reset(dlg);
    static_cast<void>(statePtr.release());
    const int showCmd = WindowPlacementPersistence::Restore(settings, kPreferencesWindowId, dlg);
    static_cast<void>(ShowWindow(dlg, showCmd));
    static_cast<void>(SetForegroundWindow(dlg));
    return true;
}

HWND PreferencesDialog::GetHandle() noexcept
{
    if (const HWND dlg = g_preferencesDialog.get(); dlg && IsWindow(dlg))
    {
        return dlg;
    }
    return nullptr;
}
