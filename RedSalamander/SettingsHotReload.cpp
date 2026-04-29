#include "SettingsHotReload.h"

#include <algorithm>
#include <mutex>
#include <string>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"
#include "HostServices.h"
#include "SettingsSave.h"
#include "SettingsSchemaExport.h"
#include "WindowBackdropPolicy.h"
#include "WindowMessages.h"
#include "resource.h"

namespace
{
using unique_change_notification = wil::unique_any<HANDLE, decltype(&::FindCloseChangeNotification), ::FindCloseChangeNotification>;

struct HotReloadState
{
    std::mutex mutex;
    HWND targetWindow = nullptr;
    std::wstring appId;
    std::filesystem::path settingsPath;
    std::filesystem::path settingsDirectory;
    wil::unique_event_nothrow stopEvent;
    std::jthread watchThread;
    std::unordered_set<HWND> participants;
    std::optional<Common::Settings::SettingsFileStamp> lastAppliedStamp;
    std::optional<Common::Settings::SettingsFileStamp> lastRejectedStamp;
    bool invalidAlertVisible = false;

    HotReloadState()                                 = default;
    HotReloadState(const HotReloadState&)            = delete;
    HotReloadState& operator=(const HotReloadState&) = delete;
    HotReloadState(HotReloadState&&)                 = delete;
    HotReloadState& operator=(HotReloadState&&)      = delete;
};

HotReloadState g_state;
int g_invalidAlertCookieStorage = 0;

[[nodiscard]] void* InvalidAlertCookie() noexcept
{
    return &g_invalidAlertCookieStorage;
}

[[nodiscard]] bool IsWatchedAppId(std::wstring_view appId) noexcept
{
    std::scoped_lock lock(g_state.mutex);
    return ! g_state.appId.empty() && OrdinalString::EqualsNoCase(g_state.appId, appId);
}

[[nodiscard]] bool ShouldIgnoreStampLocked(const Common::Settings::SettingsFileStamp& stamp) noexcept
{
    return (g_state.lastAppliedStamp.has_value() && g_state.lastAppliedStamp.value() == stamp) ||
           (g_state.lastRejectedStamp.has_value() && g_state.lastRejectedStamp.value() == stamp);
}

void WatchSettingsDirectoryThread(HWND targetWindow, HANDLE stopEventHandle, std::filesystem::path directoryPath) noexcept
{
    if (! targetWindow || directoryPath.empty())
    {
        return;
    }

    const std::wstring directoryText = directoryPath.wstring();
    constexpr DWORD kNotifyFilter    = FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_LAST_WRITE | FILE_NOTIFY_CHANGE_SIZE;

    unique_change_notification changeNotification;

    for (;;)
    {
        if (! changeNotification)
        {
            HANDLE raw = FindFirstChangeNotificationW(directoryText.c_str(), FALSE, kNotifyFilter);
            if (raw == nullptr || raw == INVALID_HANDLE_VALUE)
            {
                static_cast<void>(Debug::ErrorWithLastError(L"SettingsHotReload: FindFirstChangeNotificationW failed for '{}'", directoryText));
                if (WaitForSingleObject(stopEventHandle, 1000) == WAIT_OBJECT_0)
                {
                    return;
                }
                continue;
            }

            changeNotification.reset(raw);
        }

        HANDLE waitHandles[2]  = {stopEventHandle, changeNotification.get()};
        const DWORD waitResult = WaitForMultipleObjects(2u, waitHandles, FALSE, INFINITE);
        if (waitResult == WAIT_OBJECT_0)
        {
            return;
        }

        if (waitResult == (WAIT_OBJECT_0 + 1))
        {
            auto payload       = std::make_unique<SettingsHotReload::SettingsFileChangedPayload>();
            payload->tickCount = GetTickCount64();
            if (! PostMessagePayload(targetWindow, WndMsg::kSettingsFileChanged, 0, std::move(payload)))
            {
                const DWORD lastError = GetLastError();
                if (lastError != ERROR_INVALID_WINDOW_HANDLE && lastError != ERROR_SUCCESS)
                {
                    Debug::Warning(L"SettingsHotReload: failed to post settings change notification (gle=0x{:08X})", lastError);
                }
            }

            if (! FindNextChangeNotification(changeNotification.get()))
            {
                static_cast<void>(Debug::ErrorWithLastError(L"SettingsHotReload: FindNextChangeNotification failed for '{}'", directoryText));
                changeNotification.reset();
            }

            continue;
        }

        if (waitResult == WAIT_FAILED)
        {
            static_cast<void>(Debug::ErrorWithLastError(L"SettingsHotReload: WaitForMultipleObjects failed"));
        }
        return;
    }
}

[[nodiscard]] HRESULT SavePreparedSettingsAndSchema(std::wstring_view appId,
                                                    Common::Settings::Settings& settings,
                                                    std::span<const PluginConfigurationSchemaSource> pluginSchemas) noexcept
{
    if (appId.empty())
    {
        return E_INVALIDARG;
    }

    const Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(settings);
    const HRESULT saveHr                            = Common::Settings::SaveSettings(appId, settingsToSave);
    if (FAILED(saveHr))
    {
        return saveHr;
    }

    const HRESULT schemaHr = pluginSchemas.empty() ? SaveAggregatedSettingsSchema(appId, settings) : SaveAggregatedSettingsSchema(appId, pluginSchemas);
    if (FAILED(schemaHr))
    {
        Debug::Error(L"SettingsHotReload: SaveAggregatedSettingsSchema failed (hr=0x{:08X})", static_cast<unsigned long>(schemaHr));
    }

    if (IsWatchedAppId(appId))
    {
        Common::Settings::SettingsFileStamp stamp{};
        const HRESULT stampHr = Common::Settings::TryGetSettingsFileStamp(appId, stamp);
        if (stampHr == S_OK)
        {
            SettingsHotReload::MarkAppliedStamp(stamp);
            SettingsHotReload::ClearInvalidReloadAlert();
        }
        else if (FAILED(stampHr))
        {
            Debug::Warning(L"SettingsHotReload: failed to refresh settings file stamp after save (hr=0x{:08X})", static_cast<unsigned long>(stampHr));
        }
    }

    return saveHr;
}

[[nodiscard]] std::wstring BuildExternalReloadConflictMessage(std::wstring_view editorName) noexcept
{
    return FormatStringResource(nullptr, IDS_FMT_SETTINGS_RELOAD_CONFLICT_KEEP_EDITING, editorName);
}

[[nodiscard]] std::wstring BuildStaleSaveConflictMessage(std::wstring_view editorName) noexcept
{
    return FormatStringResource(nullptr, IDS_FMT_SETTINGS_RELOAD_CONFLICT_STALE_SAVE, editorName);
}

[[nodiscard]] ThemeMode ThemeModeFromThemeId(std::wstring_view id) noexcept
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

[[nodiscard]] Common::Settings::UiSettings GetUiSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    if (settings.ui.has_value())
    {
        return settings.ui.value();
    }
    return {};
}

[[nodiscard]] AppBackdropType ToAppBackdropType(Common::WindowBackdrop::Kind kind) noexcept
{
    switch (kind)
    {
        case Common::WindowBackdrop::Kind::Mica: return AppBackdropType::Mica;
        case Common::WindowBackdrop::Kind::Acrylic: return AppBackdropType::Acrylic;
        case Common::WindowBackdrop::Kind::MicaAlt: return AppBackdropType::MicaAlt;
        case Common::WindowBackdrop::Kind::None:
        default: return AppBackdropType::None;
    }
}

void ApplyResolvedWindowBackdrop(Common::Settings::WindowBackdropMode mode, AppTheme& theme) noexcept
{
    theme.primaryWindowBackdrop = ToAppBackdropType(Common::WindowBackdrop::Resolve(mode, Common::WindowBackdrop::Target::Primary, false));
    theme.toolWindowBackdrop    = ToAppBackdropType(Common::WindowBackdrop::Resolve(mode, Common::WindowBackdrop::Target::Tool, false));

    if (! theme.highContrast && (theme.primaryWindowBackdrop != AppBackdropType::None || theme.toolWindowBackdrop != AppBackdropType::None))
    {
        theme.titleBar.captionColor.reset();
        theme.titleBar.borderColor.reset();
        theme.titleBar.textColor.reset();
    }
}

[[nodiscard]] COLORREF ColorRefFromArgb(uint32_t argb) noexcept
{
    const uint8_t r = static_cast<uint8_t>((argb >> 16) & 0xFFu);
    const uint8_t g = static_cast<uint8_t>((argb >> 8) & 0xFFu);
    const uint8_t b = static_cast<uint8_t>(argb & 0xFFu);
    return RGB(r, g, b);
}

[[nodiscard]] float AlphaFromArgb(uint32_t argb) noexcept
{
    const uint8_t a = static_cast<uint8_t>((argb >> 24) & 0xFFu);
    return static_cast<float>(a) / 255.0f;
}

[[nodiscard]] std::optional<uint32_t> FindColorOverride(const std::unordered_map<std::wstring, uint32_t>& colors, std::wstring_view key) noexcept
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
        const D2D1::ColorF composite  = D2D1::ColorF(overlay.r * alpha + background.r * (1.0f - alpha),
                                                     overlay.g * alpha + background.g * (1.0f - alpha),
                                                     overlay.b * alpha + background.b * (1.0f - alpha),
                                                     1.0f);

        const COLORREF contrastText           = ChooseContrastingTextColor(ColorToCOLORREF(composite));
        theme.folderView.textSelectedInactive = ColorFromCOLORREF(contrastText);
    }
}

[[nodiscard]] std::optional<D2D1::ColorF> FindAccentOverride(const std::unordered_map<std::wstring, uint32_t>& colors) noexcept
{
    const auto argb = FindColorOverride(colors, L"app.accent");
    if (! argb)
    {
        return std::nullopt;
    }
    const COLORREF rgb = ColorRefFromArgb(*argb);
    return ColorFromCOLORREF(rgb, AlphaFromArgb(*argb));
}

[[nodiscard]] bool IsRetryableSettingsLoadFailure(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) || hr == HRESULT_FROM_WIN32(ERROR_LOCK_VIOLATION) || hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
}

[[nodiscard]] bool IsInvalidSettingsLoadFailure(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}
} // namespace

namespace SettingsHotReload
{
void ApplyUiPreferencesToTheme(const Common::Settings::Settings& settings, AppTheme& theme) noexcept
{
    const Common::Settings::UiSettings ui = GetUiSettingsOrDefault(settings);
    theme.compactMode                     = ui.compactMode;

    switch (ui.reducedMotion)
    {
        case Common::Settings::ReducedMotionMode::On: theme.reducedMotionOverride = true; break;
        case Common::Settings::ReducedMotionMode::Off: theme.reducedMotionOverride = false; break;
        case Common::Settings::ReducedMotionMode::System: theme.reducedMotionOverride.reset(); break;
    }

    ApplyResolvedWindowBackdrop(ui.windowBackdrop, theme);
}

HRESULT Start(HWND targetWindow, std::wstring_view appId) noexcept
{
    Stop();

    if (! targetWindow || ! IsWindow(targetWindow) || appId.empty())
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(appId);
    if (settingsPath.empty())
    {
        return E_FAIL;
    }

    const std::filesystem::path settingsDirectory = settingsPath.parent_path();
    if (settingsDirectory.empty())
    {
        return E_FAIL;
    }

    wil::unique_event_nothrow stopEvent;
    stopEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! stopEvent)
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    std::optional<Common::Settings::SettingsFileStamp> initialStamp;
    Common::Settings::SettingsFileStamp stamp{};
    const HRESULT stampHr = Common::Settings::TryGetSettingsFileStamp(appId, stamp);
    if (stampHr == S_OK)
    {
        initialStamp = stamp;
    }
    else if (FAILED(stampHr))
    {
        Debug::Warning(L"SettingsHotReload: initial file stamp query failed (hr=0x{:08X})", static_cast<unsigned long>(stampHr));
    }

    {
        std::scoped_lock lock(g_state.mutex);
        g_state.targetWindow      = targetWindow;
        g_state.appId             = std::wstring(appId);
        g_state.settingsPath      = settingsPath;
        g_state.settingsDirectory = settingsDirectory;
        g_state.stopEvent         = std::move(stopEvent);
        g_state.lastAppliedStamp  = initialStamp;
        g_state.lastRejectedStamp.reset();
        g_state.invalidAlertVisible = false;
    }

    HANDLE stopHandle = nullptr;
    {
        std::scoped_lock lock(g_state.mutex);
        stopHandle = g_state.stopEvent.get();
    }

    g_state.watchThread = std::jthread([targetWindow, stopHandle, settingsDirectory](std::stop_token) noexcept
    { WatchSettingsDirectoryThread(targetWindow, stopHandle, settingsDirectory); });
    return S_OK;
}

void Stop() noexcept
{
    ClearInvalidReloadAlert();

    {
        std::scoped_lock lock(g_state.mutex);
        if (g_state.stopEvent)
        {
            static_cast<void>(SetEvent(g_state.stopEvent.get()));
        }
    }

    if (g_state.watchThread.joinable())
    {
        g_state.watchThread.join();
    }

    std::scoped_lock lock(g_state.mutex);
    g_state.stopEvent.reset();
    g_state.targetWindow = nullptr;
    g_state.appId.clear();
    g_state.settingsPath.clear();
    g_state.settingsDirectory.clear();
    g_state.participants.clear();
    g_state.lastAppliedStamp.reset();
    g_state.lastRejectedStamp.reset();
}

HRESULT SaveSettingsAndSchema(std::wstring_view appId, Common::Settings::Settings& settings) noexcept
{
    return SavePreparedSettingsAndSchema(appId, settings, {});
}

HRESULT SaveSettingsAndSchema(std::wstring_view appId,
                              Common::Settings::Settings& settings,
                              std::span<const PluginConfigurationSchemaSource> pluginSchemas) noexcept
{
    return SavePreparedSettingsAndSchema(appId, settings, pluginSchemas);
}

ChangedSettingsLoadResult TryLoadChangedSettings() noexcept
{
    ChangedSettingsLoadResult result{};

    std::wstring appId;
    {
        std::scoped_lock lock(g_state.mutex);
        appId = g_state.appId;
    }

    if (appId.empty())
    {
        result.status = ChangedSettingsStatus::Error;
        result.hr     = HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        return result;
    }

    Common::Settings::SettingsFileStamp stamp{};
    const HRESULT stampHr = Common::Settings::TryGetSettingsFileStamp(appId, stamp);
    if (stampHr == S_FALSE)
    {
        result.status = ChangedSettingsStatus::Missing;
        result.hr     = S_FALSE;
        return result;
    }
    if (FAILED(stampHr))
    {
        result.status = ChangedSettingsStatus::Error;
        result.hr     = stampHr;
        return result;
    }

    {
        std::scoped_lock lock(g_state.mutex);
        if (ShouldIgnoreStampLocked(stamp))
        {
            result.status = ChangedSettingsStatus::NoChange;
            result.hr     = S_OK;
            result.stamp  = stamp;
            return result;
        }
    }

    Common::Settings::Settings loaded;
    HRESULT loadHr = S_FALSE;
    for (int attempt = 0; attempt < 6; ++attempt)
    {
        loadHr = Common::Settings::TryLoadSettingsNoRecovery(appId, loaded);
        if (loadHr == S_OK)
        {
            break;
        }

        if (loadHr != S_FALSE && ! IsRetryableSettingsLoadFailure(loadHr))
        {
            break;
        }

        if (attempt + 1 < 6)
        {
            ::Sleep(25);
        }
    }
    if (loadHr == S_OK)
    {
        result.status   = ChangedSettingsStatus::Loaded;
        result.hr       = S_OK;
        result.settings = std::move(loaded);
        result.stamp    = stamp;
        return result;
    }

    if (loadHr == S_FALSE)
    {
        result.status = ChangedSettingsStatus::Missing;
        result.hr     = S_FALSE;
        return result;
    }

    result.status = IsInvalidSettingsLoadFailure(loadHr) ? ChangedSettingsStatus::Invalid : ChangedSettingsStatus::Error;
    result.hr     = loadHr;
    result.stamp  = stamp;
    return result;
}

void MarkAppliedStamp(const Common::Settings::SettingsFileStamp& stamp) noexcept
{
    std::scoped_lock lock(g_state.mutex);
    g_state.lastAppliedStamp = stamp;
    g_state.lastRejectedStamp.reset();
}

void MarkRejectedStamp(const Common::Settings::SettingsFileStamp& stamp) noexcept
{
    std::scoped_lock lock(g_state.mutex);
    g_state.lastRejectedStamp = stamp;
}

void RegisterParticipant(HWND hwnd) noexcept
{
    if (! hwnd || ! IsWindow(hwnd))
    {
        return;
    }

    std::scoped_lock lock(g_state.mutex);
    g_state.participants.insert(hwnd);
}

void UnregisterParticipant(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    std::scoped_lock lock(g_state.mutex);
    g_state.participants.erase(hwnd);
}

void NotifyParticipants() noexcept
{
    std::vector<HWND> participants;
    {
        std::scoped_lock lock(g_state.mutex);
        participants.reserve(g_state.participants.size());
        for (HWND hwnd : g_state.participants)
        {
            if (hwnd && IsWindow(hwnd))
            {
                participants.push_back(hwnd);
            }
        }
    }

    for (HWND hwnd : participants)
    {
        static_cast<void>(PostMessageW(hwnd, WndMsg::kSettingsReloadedFromDisk, 0, 0));
    }
}

HRESULT PromptExternalReloadConflict(HWND targetWindow, std::wstring_view editorName, ExternalReloadChoice& outChoice) noexcept
{
    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_SETTINGS_RELOADED_FROM_DISK);
    const std::wstring message = BuildExternalReloadConflictMessage(editorName);

    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_WARNING;
    prompt.buttons       = HOST_PROMPT_BUTTONS_YES_NO;
    prompt.targetWindow  = targetWindow;
    prompt.title         = title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_YES;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hrPrompt))
    {
        return hrPrompt;
    }

    outChoice = (promptResult == HOST_PROMPT_RESULT_YES) ? ExternalReloadChoice::ReloadFromDisk : ExternalReloadChoice::KeepEditing;
    return S_OK;
}

HRESULT PromptStaleSaveConflict(HWND targetWindow, std::wstring_view editorName, StaleSaveChoice& outChoice) noexcept
{
    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_SETTINGS_RELOADED_FROM_DISK);
    const std::wstring message = BuildStaleSaveConflictMessage(editorName);

    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_WARNING;
    prompt.buttons       = HOST_PROMPT_BUTTONS_YES_NO_CANCEL;
    prompt.targetWindow  = targetWindow;
    prompt.title         = title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_CANCEL;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hrPrompt))
    {
        return hrPrompt;
    }

    switch (promptResult)
    {
        case HOST_PROMPT_RESULT_OK:
        case HOST_PROMPT_RESULT_YES: outChoice = StaleSaveChoice::OverwriteCurrent; break;

        case HOST_PROMPT_RESULT_NO: outChoice = StaleSaveChoice::ReloadFromDisk; break;

        case HOST_PROMPT_RESULT_CANCEL:
        case HOST_PROMPT_RESULT_NONE:
        default: outChoice = StaleSaveChoice::Cancel; break;
    }

    return S_OK;
}

void ShowInvalidReloadAlert(const std::filesystem::path& settingsPath) noexcept
{
    bool showAlert = false;
    {
        std::scoped_lock lock(g_state.mutex);
        if (! g_state.invalidAlertVisible)
        {
            g_state.invalidAlertVisible = true;
            showAlert                   = true;
        }
    }

    if (! showAlert)
    {
        return;
    }

    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_SETTINGS_RELOAD_FAILED);
    const std::wstring message = FormatStringResource(nullptr, IDS_FMT_SETTINGS_RELOAD_FAILED_KEEP_CURRENT, settingsPath.wstring());

    HostAlertRequest request{};
    request.version   = 1;
    request.sizeBytes = sizeof(request);
    request.scope     = HOST_ALERT_SCOPE_APPLICATION;
    request.modality  = HOST_ALERT_MODELESS;
    request.severity  = HOST_ALERT_WARNING;
    request.title     = title.c_str();
    request.message   = message.c_str();
    request.closable  = TRUE;

    const HRESULT hrAlert = HostShowAlert(request, InvalidAlertCookie());
    if (FAILED(hrAlert))
    {
        std::scoped_lock lock(g_state.mutex);
        g_state.invalidAlertVisible = false;
        Debug::Warning(L"SettingsHotReload: failed to show invalid settings alert (hr=0x{:08X})", static_cast<unsigned long>(hrAlert));
    }
}

void ClearInvalidReloadAlert() noexcept
{
    bool clearAlert = false;
    {
        std::scoped_lock lock(g_state.mutex);
        if (g_state.invalidAlertVisible)
        {
            g_state.invalidAlertVisible = false;
            clearAlert                  = true;
        }
    }

    if (! clearAlert)
    {
        return;
    }

    const HRESULT hrClear = HostClearAlert(HOST_ALERT_SCOPE_APPLICATION, InvalidAlertCookie());
    if (FAILED(hrClear))
    {
        Debug::Warning(L"SettingsHotReload: failed to clear invalid settings alert (hr=0x{:08X})", static_cast<unsigned long>(hrClear));
    }
}

AppTheme ResolveDialogThemeFromSettings(const Common::Settings::Settings& settings) noexcept
{
    std::wstring_view themeId = settings.theme.currentThemeId;

    const Common::Settings::ThemeDefinition* custom = nullptr;
    if (themeId.rfind(L"user/", 0) == 0)
    {
        const auto it = std::find_if(settings.theme.themes.begin(), settings.theme.themes.end(), [&](const Common::Settings::ThemeDefinition& entry) noexcept {
            return entry.id == themeId;
        });
        if (it != settings.theme.themes.end())
        {
            custom = &*it;
        }
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

    ApplyUiPreferencesToTheme(settings, theme);

    return theme;
}

Common::Settings::Settings MergeDiskSettingsWithRuntimeSession(const Common::Settings::Settings& diskSettings,
                                                               const Common::Settings::Settings& runtimeSettings,
                                                               std::span<const std::wstring_view> runtimeWindowIds) noexcept
{
    Common::Settings::Settings merged = diskSettings;

    for (std::wstring_view windowId : runtimeWindowIds)
    {
        const auto runtimeIt = runtimeSettings.windows.find(std::wstring(windowId));
        if (runtimeIt != runtimeSettings.windows.end())
        {
            merged.windows[runtimeIt->first] = runtimeIt->second;
        }
    }

    if (! runtimeSettings.folders.has_value())
    {
        return merged;
    }

    if (! merged.folders.has_value())
    {
        merged.folders = Common::Settings::FoldersSettings{};
    }

    const auto& runtimeFolders = runtimeSettings.folders.value();
    auto& mergedFolders        = merged.folders.value();

    mergedFolders.active         = runtimeFolders.active;
    mergedFolders.layout         = runtimeFolders.layout;
    mergedFolders.history        = runtimeFolders.history;
    mergedFolders.historyFilters = runtimeFolders.historyFilters;

    for (const auto& runtimePane : runtimeFolders.items)
    {
        auto mergedIt = std::find_if(
            mergedFolders.items.begin(), mergedFolders.items.end(), [&](const Common::Settings::FolderPane& item) { return item.slot == runtimePane.slot; });
        if (mergedIt != mergedFolders.items.end())
        {
            mergedIt->current = runtimePane.current;
            continue;
        }

        Common::Settings::FolderPane pane;
        pane.slot    = runtimePane.slot;
        pane.current = runtimePane.current;
        mergedFolders.items.push_back(std::move(pane));
    }

    return merged;
}
} // namespace SettingsHotReload
