#include "CrashQuarantine.h"

#include "Framework.h"

#include <exception>
#include <string>
#include <string_view>
#include <vector>

#include <commctrl.h>
#include <shellapi.h>

#include "CrashHandler.h"
#include "Helpers.h"
#include "SessionState.h"
#include "SettingsStore.h"
#include "resource.h"

#pragma comment(lib, "Comctl32.lib")

namespace CrashQuarantine
{
namespace
{
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

void OfferPreviousCrashRecovery(Common::Settings::Settings& settings) noexcept
{
    try
    {
        const auto report = CrashHandler::TakePendingReport();
        if (! report)
            return;

        const auto state         = SessionState::TryRead();
        const auto toDisable     = state ? ComputePluginsToDisable(true, state->activeFileSystemPluginIds, settings) : std::vector<std::wstring>{};
        const std::wstring title = LoadStringResource(nullptr, toDisable.empty() ? IDS_CRASH_DETECTED_TITLE : IDS_CRASH_QUARANTINE_TITLE);
        std::wstring content     = report->dumpPath.empty() ? LoadStringResource(nullptr, IDS_CRASH_DETECTED_MESSAGE)
                                                            : FormatStringResource(nullptr, IDS_CRASH_DETECTED_MESSAGE_FMT, report->dumpPath);
        if (! toDisable.empty())
        {
            std::wstring pluginList;
            for (const auto& id : toDisable)
            {
                if (! pluginList.empty())
                    pluginList.append(L"\r\n");
                pluginList.append(id);
            }
            const UINT warningId = toDisable.size() > 1 ? IDS_CRASH_QUARANTINE_MESSAGE_PLURAL_FMT : IDS_CRASH_QUARANTINE_MESSAGE_FMT;
            content.append(L"\r\n\r\n");
            content.append(FormatStringResource(nullptr, warningId, pluginList));
        }

        const std::wstring openFolderLabel = LoadStringResource(nullptr, IDS_CRASH_OPEN_FOLDER_OPTION);
        const std::wstring acceptLabel     = LoadStringResource(nullptr, toDisable.empty() ? IDS_BTN_OK : IDS_BTN_YES);
        const std::wstring keepLabel       = toDisable.empty() ? std::wstring{} : LoadStringResource(nullptr, IDS_BTN_NO);
        const TASKDIALOG_BUTTON buttons[]  = {{toDisable.empty() ? IDOK : IDYES, acceptLabel.c_str()}, {IDNO, keepLabel.c_str()}};
        TASKDIALOGCONFIG dialog{};
        dialog.cbSize              = sizeof(dialog);
        dialog.dwFlags             = TDF_ALLOW_DIALOG_CANCELLATION;
        dialog.cButtons            = toDisable.empty() ? 1u : 2u;
        dialog.pButtons            = buttons;
        dialog.nDefaultButton      = buttons[0].nButtonID;
        dialog.pszWindowTitle      = title.c_str();
        dialog.pszMainIcon         = TD_WARNING_ICON;
        dialog.pszContent          = content.c_str();
        dialog.pszVerificationText = openFolderLabel.c_str();

        int choice      = 0;
        BOOL openFolder = FALSE;
        if (FAILED(TaskDialogIndirect(&dialog, &choice, nullptr, &openFolder)) || choice == IDCANCEL)
            return;
        if (choice == IDYES)
        {
            for (const auto& id : toDisable)
                DisablePluginIdInSettings(id, settings);
        }
        if (openFolder)
            static_cast<void>(ShellExecuteW(nullptr, L"open", report->directory.c_str(), nullptr, nullptr, SW_SHOWNORMAL));
    }
    catch (const std::exception&)
    {
        // Recovery is a healthy-start action. A failure here leaves the local report intact.
    }
}
} // namespace CrashQuarantine
