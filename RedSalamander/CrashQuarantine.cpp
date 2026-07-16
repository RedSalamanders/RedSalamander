#include "CrashQuarantine.h"

#include "Framework.h"

#include <filesystem>
#include <string>
#include <string_view>
#include <system_error>

#include "AppDataPaths.h"
#include "Helpers.h"
#include "HostServices.h"
#include "SessionState.h"
#include "SettingsStore.h"
#include "resource.h"

namespace CrashQuarantine
{
namespace
{
constexpr wchar_t kCompanyDirName[] = L"RedSalamander";
constexpr wchar_t kCrashDirName[]   = L"Crashes";
constexpr wchar_t kMarkerFileName[] = L"last_crash.txt";

[[nodiscard]] std::filesystem::path GetCrashMarkerPath() noexcept
{
    const std::filesystem::path base = AppDataPaths::GetLocalAppDataPath();
    if (base.empty())
    {
        return {};
    }

    return base / kCompanyDirName / kCrashDirName / kMarkerFileName;
}

[[nodiscard]] bool IsDisabledPluginId(std::wstring_view pluginId, const Common::Settings::Settings& settings) noexcept
{
    for (const std::wstring& disabled : settings.plugins.disabledPluginIds)
    {
        if (OrdinalString::EqualsNoCase(pluginId, disabled))
        {
            return true;
        }
    }
    return false;
}

void DisablePluginIdInSettings(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept
{
    if (pluginId.empty())
    {
        return;
    }

    if (! IsDisabledPluginId(pluginId, settings))
    {
        settings.plugins.disabledPluginIds.emplace_back(pluginId);
    }

    if (OrdinalString::EqualsNoCase(settings.plugins.currentFileSystemPluginId, pluginId))
    {
        settings.plugins.currentFileSystemPluginId.clear();
    }
}

// Pure decision core: filter active ids to non-empty ids not already disabled. No I/O, no UI.
[[nodiscard]] std::vector<std::wstring> ComputePluginsToDisable(bool markerExists,
                                                                const std::vector<std::wstring>& activeFileSystemPluginIds,
                                                                const Common::Settings::Settings& settings) noexcept
{
    std::vector<std::wstring> toDisable;
    if (! markerExists)
    {
        return toDisable;
    }

    toDisable.reserve(activeFileSystemPluginIds.size());
    for (const std::wstring& id : activeFileSystemPluginIds)
    {
        if (id.empty())
        {
            continue;
        }
        if (! IsDisabledPluginId(id, settings))
        {
            toDisable.push_back(id);
        }
    }
    return toDisable;
}
} // namespace

#ifdef ENABLE_TESTS
std::vector<std::wstring> SelectPluginsToDisable(bool markerExists,
                                                 const std::vector<std::wstring>& activeFileSystemPluginIds,
                                                 const Common::Settings::Settings& settings) noexcept
{
    return ComputePluginsToDisable(markerExists, activeFileSystemPluginIds, settings);
}

void DisablePluginIdInSettingsForTest(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept
{
    DisablePluginIdInSettings(pluginId, settings);
}
#endif

void OfferPluginDisableIfPreviousCrashDetected(Common::Settings::Settings& settings) noexcept
{
    const std::filesystem::path markerPath = GetCrashMarkerPath();
    if (markerPath.empty())
    {
        return;
    }

    std::error_code ec;
    if (! std::filesystem::exists(markerPath, ec))
    {
        return;
    }

    const auto stateOpt = SessionState::TryRead();
    if (! stateOpt.has_value() || stateOpt->activeFileSystemPluginIds.empty())
    {
        return;
    }

    const std::vector<std::wstring> toDisable = ComputePluginsToDisable(true, stateOpt->activeFileSystemPluginIds, settings);
    if (toDisable.empty())
    {
        return;
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_CRASH_QUARANTINE_TITLE);
    std::wstring pluginList;
    for (size_t i = 0; i < toDisable.size(); ++i)
    {
        if (i != 0)
        {
            pluginList.append(L"\r\n");
        }
        pluginList.append(toDisable[i]);
    }

    const UINT messageId       = (toDisable.size() > 1u) ? IDS_CRASH_QUARANTINE_MESSAGE_PLURAL_FMT : IDS_CRASH_QUARANTINE_MESSAGE_FMT;
    const std::wstring message = FormatStringResource(nullptr, messageId, pluginList);

    HostPromptRequest prompt{};
    prompt.version       = 1;
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_APPLICATION;
    prompt.severity      = HOST_ALERT_WARNING;
    prompt.buttons       = HOST_PROMPT_BUTTONS_YES_NO;
    prompt.targetWindow  = nullptr;
    prompt.title         = title.empty() ? nullptr : title.c_str();
    prompt.message       = message.empty() ? nullptr : message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_YES;

    HostPromptResult result = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt  = HostShowPrompt(prompt, nullptr, &result);
    if (FAILED(hrPrompt) || result != HOST_PROMPT_RESULT_YES)
    {
        return;
    }

    for (const std::wstring_view id : toDisable)
    {
        DisablePluginIdInSettings(id, settings);
    }
}
} // namespace CrashQuarantine
