// Preferences.Advanced.cpp

#include "Framework.h"

#include "Preferences.Advanced.h"

#include <algorithm>
#include <array>
#include <cwctype>
#include <limits>
#include <string>

#include "Helpers.h"
#include "ThemedControls.h"
#include "resource.h"

// Local convenience aliases for frequently-used shared utilities
namespace
{
using PrefsCache::EnsureWorkingCacheSettings;
using PrefsCache::FormatCacheBytes;
using PrefsCache::GetCacheSettingsOrDefault;
using PrefsCache::MaybeResetWorkingCacheSettingsIfEmpty;
using PrefsCache::TryParseCacheBytes;
using PrefsConnections::EnsureWorkingConnectionsSettings;
using PrefsConnections::GetConnectionsSettingsOrDefault;
using PrefsConnections::MaybeResetWorkingConnectionsSettingsIfEmpty;
using PrefsFileOperations::EnsureWorkingFileOperationsSettings;
using PrefsFileOperations::GetFileOperationsSettingsOrDefault;
using PrefsFileOperations::MaybeResetWorkingFileOperationsSettingsIfEmpty;
using PrefsMonitor::EnsureWorkingMonitorSettings;
using PrefsMonitor::GetMonitorSettingsOrDefault;
} // namespace

bool AdvancedPane::EnsureCreated(HWND pageHost) noexcept
{
    return PrefsPaneHost::EnsureCreated(pageHost, _hWnd);
}

void AdvancedPane::ResizeToHostClient(HWND pageHost) noexcept
{
    PrefsPaneHost::ResizeToHostClient(pageHost, _hWnd.get());
}

void AdvancedPane::Show(bool visible) noexcept
{
    PrefsPaneHost::Show(_hWnd.get(), visible);
}

void AdvancedPane::CreateControls(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    const DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const DWORD wrapStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;
    const bool customButtons    = ! state.theme.systemHighContrast;

    state.advancedConnectionsHelloHeader.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedConnectionsBypassHelloLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedConnectionsBypassHelloDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedConnectionsHelloTimeoutLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.advancedConnectionsHelloTimeoutFrame,
                                    state.advancedConnectionsHelloTimeoutEdit,
                                    IDC_PREFS_ADV_CONNECTIONS_HELLO_TIMEOUT_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL);
    if (state.advancedConnectionsHelloTimeoutEdit)
    {
        SendMessageW(state.advancedConnectionsHelloTimeoutEdit.get(), EM_SETLIMITTEXT, 10, 0);
    }
    state.advancedConnectionsHelloTimeoutDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    state.advancedMonitorHeader.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorToolbarLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorToolbarDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorLineNumbersLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorLineNumbersDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorAlwaysOnTopLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorAlwaysOnTopDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorShowIdsLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorShowIdsDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorAutoScrollLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorAutoScrollDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    state.advancedMonitorFilterPresetLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    PrefsInput::CreateFramedComboBox(
        state, parent, state.advancedMonitorFilterPresetFrame, state.advancedMonitorFilterPresetCombo, IDC_PREFS_ADV_MONITOR_FILTER_PRESET_COMBO);

    state.advancedMonitorFilterPresetDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    state.advancedMonitorFilterMaskLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.advancedMonitorFilterMaskFrame,
                                    state.advancedMonitorFilterMaskEdit,
                                    IDC_PREFS_ADV_MONITOR_FILTER_MASK_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL);
    if (state.advancedMonitorFilterMaskEdit)
    {
        SendMessageW(state.advancedMonitorFilterMaskEdit.get(), EM_SETLIMITTEXT, 2, 0);
    }
    state.advancedMonitorFilterMaskDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    state.advancedMonitorFilterTextLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorFilterTextDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorFilterErrorLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorFilterErrorDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorFilterWarningLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorFilterWarningDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorFilterInfoLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorFilterInfoDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorFilterDebugLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedMonitorFilterDebugDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    const DWORD monitorToggleStyle =
        customButtons ? (WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW) : (WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX);

    state.advancedConnectionsBypassHelloToggle.reset(
        CreateWindowExW(0,
                        L"Button",
                        customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_CONNECTIONS_BYPASS_HELLO).c_str(),
                        monitorToggleStyle,
                        0,
                        0,
                        10,
                        10,
                        parent,
                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_CONNECTIONS_BYPASS_HELLO_TOGGLE)),
                        GetModuleHandleW(nullptr),
                        nullptr));

    state.advancedMonitorToolbarToggle.reset(CreateWindowExW(0,
                                                             L"Button",
                                                             customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_SHOW_TOOLBAR).c_str(),
                                                             monitorToggleStyle,
                                                             0,
                                                             0,
                                                             10,
                                                             10,
                                                             parent,
                                                             reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_MONITOR_TOOLBAR_TOGGLE)),
                                                             GetModuleHandleW(nullptr),
                                                             nullptr));
    state.advancedMonitorLineNumbersToggle.reset(
        CreateWindowExW(0,
                        L"Button",
                        customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_SHOW_LINE_NUMBERS).c_str(),
                        monitorToggleStyle,
                        0,
                        0,
                        10,
                        10,
                        parent,
                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_MONITOR_LINE_NUMBERS_TOGGLE)),
                        GetModuleHandleW(nullptr),
                        nullptr));
    state.advancedMonitorAlwaysOnTopToggle.reset(CreateWindowExW(0,
                                                                 L"Button",
                                                                 customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_ALWAYS_ON_TOP).c_str(),
                                                                 monitorToggleStyle,
                                                                 0,
                                                                 0,
                                                                 10,
                                                                 10,
                                                                 parent,
                                                                 reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_MONITOR_ALWAYS_ON_TOP_TOGGLE)),
                                                                 GetModuleHandleW(nullptr),
                                                                 nullptr));
    state.advancedMonitorShowIdsToggle.reset(CreateWindowExW(0,
                                                             L"Button",
                                                             customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_SHOW_IDS).c_str(),
                                                             monitorToggleStyle,
                                                             0,
                                                             0,
                                                             10,
                                                             10,
                                                             parent,
                                                             reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_MONITOR_SHOW_IDS_TOGGLE)),
                                                             GetModuleHandleW(nullptr),
                                                             nullptr));
    state.advancedMonitorAutoScrollToggle.reset(CreateWindowExW(0,
                                                                L"Button",
                                                                customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_AUTO_SCROLL).c_str(),
                                                                monitorToggleStyle,
                                                                0,
                                                                0,
                                                                10,
                                                                10,
                                                                parent,
                                                                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_MONITOR_AUTO_SCROLL_TOGGLE)),
                                                                GetModuleHandleW(nullptr),
                                                                nullptr));

    state.advancedMonitorFilterTextToggle.reset(CreateWindowExW(0,
                                                                L"Button",
                                                                customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_FILTER_TEXT).c_str(),
                                                                monitorToggleStyle,
                                                                0,
                                                                0,
                                                                10,
                                                                10,
                                                                parent,
                                                                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_MONITOR_FILTER_TEXT_TOGGLE)),
                                                                GetModuleHandleW(nullptr),
                                                                nullptr));
    state.advancedMonitorFilterErrorToggle.reset(CreateWindowExW(0,
                                                                 L"Button",
                                                                 customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_FILTER_ERROR).c_str(),
                                                                 monitorToggleStyle,
                                                                 0,
                                                                 0,
                                                                 10,
                                                                 10,
                                                                 parent,
                                                                 reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_MONITOR_FILTER_ERROR_TOGGLE)),
                                                                 GetModuleHandleW(nullptr),
                                                                 nullptr));
    state.advancedMonitorFilterWarningToggle.reset(
        CreateWindowExW(0,
                        L"Button",
                        customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_FILTER_WARNING).c_str(),
                        monitorToggleStyle,
                        0,
                        0,
                        10,
                        10,
                        parent,
                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_MONITOR_FILTER_WARNING_TOGGLE)),
                        GetModuleHandleW(nullptr),
                        nullptr));
    state.advancedMonitorFilterInfoToggle.reset(CreateWindowExW(0,
                                                                L"Button",
                                                                customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_FILTER_INFO).c_str(),
                                                                monitorToggleStyle,
                                                                0,
                                                                0,
                                                                10,
                                                                10,
                                                                parent,
                                                                reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_MONITOR_FILTER_INFO_TOGGLE)),
                                                                GetModuleHandleW(nullptr),
                                                                nullptr));
    state.advancedMonitorFilterDebugToggle.reset(CreateWindowExW(0,
                                                                 L"Button",
                                                                 customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_FILTER_DEBUG).c_str(),
                                                                 monitorToggleStyle,
                                                                 0,
                                                                 0,
                                                                 10,
                                                                 10,
                                                                 parent,
                                                                 reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_MONITOR_FILTER_DEBUG_TOGGLE)),
                                                                 GetModuleHandleW(nullptr),
                                                                 nullptr));

    state.advancedFileOperationsDiagnosticsInfoToggle.reset(
        CreateWindowExW(0,
                        L"Button",
                        customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_FILEOPS_DIAG_INFO).c_str(),
                        monitorToggleStyle,
                        0,
                        0,
                        10,
                        10,
                        parent,
                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_FILEOPS_DIAG_INFO_TOGGLE)),
                        GetModuleHandleW(nullptr),
                        nullptr));
    state.advancedFileOperationsDiagnosticsDebugToggle.reset(
        CreateWindowExW(0,
                        L"Button",
                        customButtons ? L"" : LoadStringResource(nullptr, IDS_PREFS_ADV_CHECK_FILEOPS_DIAG_DEBUG).c_str(),
                        monitorToggleStyle,
                        0,
                        0,
                        10,
                        10,
                        parent,
                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_ADV_FILEOPS_DIAG_DEBUG_TOGGLE)),
                        GetModuleHandleW(nullptr),
                        nullptr));

    PrefsInput::EnableMouseWheelForwarding(state.advancedConnectionsBypassHelloToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedMonitorToolbarToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedMonitorLineNumbersToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedMonitorAlwaysOnTopToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedMonitorShowIdsToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedMonitorAutoScrollToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedMonitorFilterTextToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedMonitorFilterErrorToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedMonitorFilterWarningToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedMonitorFilterInfoToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedMonitorFilterDebugToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedFileOperationsDiagnosticsInfoToggle);
    PrefsInput::EnableMouseWheelForwarding(state.advancedFileOperationsDiagnosticsDebugToggle);

    state.advancedCacheHeader.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    state.advancedCacheDirectoryInfoMaxBytesLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.advancedCacheDirectoryInfoMaxBytesFrame,
                                    state.advancedCacheDirectoryInfoMaxBytesEdit,
                                    IDC_PREFS_ADV_CACHE_DIR_MAX_BYTES_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.advancedCacheDirectoryInfoMaxBytesEdit)
    {
        SendMessageW(state.advancedCacheDirectoryInfoMaxBytesEdit.get(), EM_SETLIMITTEXT, 24, 0);
    }
    state.advancedCacheDirectoryInfoMaxBytesDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    state.advancedCacheDirectoryInfoMaxWatchersLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.advancedCacheDirectoryInfoMaxWatchersFrame,
                                    state.advancedCacheDirectoryInfoMaxWatchersEdit,
                                    IDC_PREFS_ADV_CACHE_DIR_MAX_WATCHERS_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL);
    if (state.advancedCacheDirectoryInfoMaxWatchersEdit)
    {
        SendMessageW(state.advancedCacheDirectoryInfoMaxWatchersEdit.get(), EM_SETLIMITTEXT, 10, 0);
    }
    state.advancedCacheDirectoryInfoMaxWatchersDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    state.advancedCacheDirectoryInfoMruWatchedLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.advancedCacheDirectoryInfoMruWatchedFrame,
                                    state.advancedCacheDirectoryInfoMruWatchedEdit,
                                    IDC_PREFS_ADV_CACHE_DIR_MRU_WATCHED_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL);
    if (state.advancedCacheDirectoryInfoMruWatchedEdit)
    {
        SendMessageW(state.advancedCacheDirectoryInfoMruWatchedEdit.get(), EM_SETLIMITTEXT, 10, 0);
    }
    state.advancedCacheDirectoryInfoMruWatchedDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    state.advancedFileOperationsHeader.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    state.advancedFileOperationsMaxDiagnosticsLogFilesLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.advancedFileOperationsMaxDiagnosticsLogFilesFrame,
                                    state.advancedFileOperationsMaxDiagnosticsLogFilesEdit,
                                    IDC_PREFS_ADV_FILEOPS_MAX_DIAG_LOG_FILES_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL);
    if (state.advancedFileOperationsMaxDiagnosticsLogFilesEdit)
    {
        SendMessageW(state.advancedFileOperationsMaxDiagnosticsLogFilesEdit.get(), EM_SETLIMITTEXT, 10, 0);
    }
    state.advancedFileOperationsMaxDiagnosticsLogFilesDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    state.advancedFileOperationsDiagnosticsInfoLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedFileOperationsDiagnosticsInfoDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedFileOperationsDiagnosticsDebugLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
    state.advancedFileOperationsDiagnosticsDebugDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));

    if (state.advancedMonitorFilterPresetCombo)
    {
        SendMessageW(state.advancedMonitorFilterPresetCombo.get(), CB_RESETCONTENT, 0, 0);
        const std::array<std::pair<UINT, LPARAM>, 4> options = {
            std::pair<UINT, LPARAM>{IDS_PREFS_ADV_FILTER_CUSTOM, static_cast<LPARAM>(static_cast<int>(Common::Settings::MonitorFilterPreset::Custom))},
            std::pair<UINT, LPARAM>{IDS_PREFS_ADV_FILTER_ERRORS_ONLY, static_cast<LPARAM>(static_cast<int>(Common::Settings::MonitorFilterPreset::ErrorsOnly))},
            std::pair<UINT, LPARAM>{IDS_PREFS_ADV_FILTER_ERRORS_WARNINGS,
                                    static_cast<LPARAM>(static_cast<int>(Common::Settings::MonitorFilterPreset::ErrorsWarnings))},
            std::pair<UINT, LPARAM>{IDS_PREFS_ADV_FILTER_ALL_TYPES, static_cast<LPARAM>(static_cast<int>(Common::Settings::MonitorFilterPreset::AllTypes))},
        };

        for (const auto& option : options)
        {
            const std::wstring label = LoadStringResource(nullptr, option.first);
            const LRESULT index      = SendMessageW(state.advancedMonitorFilterPresetCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str()));
            if (index != CB_ERR && index != CB_ERRSPACE)
            {
                SendMessageW(state.advancedMonitorFilterPresetCombo.get(), CB_SETITEMDATA, static_cast<WPARAM>(index), option.second);
            }
        }

        ThemedControls::ApplyThemeToComboBox(state.advancedMonitorFilterPresetCombo, state.theme);
    }

    Refresh(parent, state);
}

void AdvancedPane::LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, HFONT dialogFont) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(margin);

    if (! host)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(host);

    const int rowHeight   = std::max(1, ThemedControls::ScaleDip(dpi, kRowHeightDip));
    const int titleHeight = std::max(1, ThemedControls::ScaleDip(dpi, kTitleHeightDip));

    const int cardPaddingX = ThemedControls::ScaleDip(dpi, kCardPaddingXDip);
    const int cardPaddingY = ThemedControls::ScaleDip(dpi, kCardPaddingYDip);
    const int cardGapY     = ThemedControls::ScaleDip(dpi, kCardGapYDip);
    const int cardGapX     = ThemedControls::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY = ThemedControls::ScaleDip(dpi, kCardSpacingYDip);

    const HFONT headerFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const HFONT infoFont   = state.italicFont ? state.italicFont.get() : dialogFont;
    const int headerHeight = std::max(1, ThemedControls::ScaleDip(dpi, kHeaderHeightDip));

    if (state.advancedConnectionsHelloHeader)
    {
        SetWindowTextW(state.advancedConnectionsHelloHeader.get(), LoadStringResource(nullptr, IDS_PREFS_ADV_HEADER_CONNECTIONS_HELLO).c_str());
        SetWindowPos(state.advancedConnectionsHelloHeader.get(), nullptr, x, y, width, headerHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.advancedConnectionsHelloHeader.get(), WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += headerHeight + gapY;
    }

    const int minToggleWidth    = ThemedControls::ScaleDip(dpi, kMinToggleWidthDip);
    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

    const HFONT toggleMeasureFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const int onWidth             = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, onLabel);
    const int offWidth            = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, offLabel);

    const int paddingX       = ThemedControls::ScaleDip(dpi, kTogglePaddingXDip);
    const int gapX           = ThemedControls::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth     = ThemedControls::ScaleDip(dpi, kToggleTrackWidthDip);
    const int stateTextWidth = std::max(onWidth, offWidth);

    const int measuredToggleWidth = std::max(minToggleWidth, (2 * paddingX) + stateTextWidth + gapX + trackWidth);
    const int toggleWidth         = std::min(std::max(0, width - 2 * cardPaddingX), measuredToggleWidth);

    auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

    auto layoutToggleCard = [&](HWND label, std::wstring_view labelText, HWND toggle, HWND descLabel, std::wstring_view descText) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight = descLabel ? PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, descText) : 0;

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        pushCard(card);

        if (label)
        {
            SetWindowTextW(label, labelText.data());
            SetWindowPos(label, nullptr, card.left + cardPaddingX, card.top + cardPaddingY, textWidth, titleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        if (descLabel)
        {
            SetWindowTextW(descLabel, descText.data());
            SetWindowPos(descLabel,
                         nullptr,
                         card.left + cardPaddingX,
                         card.top + cardPaddingY + titleHeight + cardGapY,
                         textWidth,
                         std::max(0, descHeight),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(descLabel, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
        }

        if (toggle)
        {
            SetWindowPos(toggle,
                         nullptr,
                         card.right - cardPaddingX - toggleWidth,
                         card.top + (cardHeight - rowHeight) / 2,
                         toggleWidth,
                         rowHeight,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        y += cardHeight + cardSpacingY;
    };

    auto layoutFramedComboCard = [&](HWND label, std::wstring_view labelText, HWND frame, HWND combo, HWND descLabel, std::wstring_view descText) noexcept
    {
        int desiredWidth          = combo ? ThemedControls::MeasureComboBoxPreferredWidth(combo, dpi) : 0;
        desiredWidth              = std::max(desiredWidth, ThemedControls::ScaleDip(dpi, kMinEditWidthDip));
        const int maxControlWidth = std::max(0, width - 2 * cardPaddingX);
        desiredWidth              = std::min(desiredWidth, std::min(maxControlWidth, ThemedControls::ScaleDip(dpi, kMaxEditWidthDip)));

        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - desiredWidth);
        const int descHeight = descLabel ? PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, descText) : 0;

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        pushCard(card);

        if (label)
        {
            SetWindowTextW(label, labelText.data());
            SetWindowPos(label, nullptr, card.left + cardPaddingX, card.top + cardPaddingY, textWidth, titleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        if (descLabel)
        {
            SetWindowTextW(descLabel, descText.data());
            SetWindowPos(descLabel,
                         nullptr,
                         card.left + cardPaddingX,
                         card.top + cardPaddingY + titleHeight + cardGapY,
                         textWidth,
                         std::max(0, descHeight),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(descLabel, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
        }

        const int inputX       = card.right - cardPaddingX - desiredWidth;
        const int inputY       = card.top + (cardHeight - rowHeight) / 2;
        const int framePadding = (frame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;

        if (frame)
        {
            SetWindowPos(frame, nullptr, inputX, inputY, desiredWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        if (combo)
        {
            SetWindowPos(combo,
                         nullptr,
                         inputX + framePadding,
                         inputY + framePadding,
                         std::max(1, desiredWidth - 2 * framePadding),
                         std::max(1, rowHeight - 2 * framePadding),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(combo, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            ThemedControls::EnsureComboBoxDroppedWidth(combo, dpi);
        }

        y += cardHeight + cardSpacingY;
    };

    auto layoutEditCard =
        [&](HWND label, std::wstring_view labelText, HWND frame, HWND edit, int desiredWidth, HWND descLabel, std::wstring_view descText) noexcept
    {
        desiredWidth         = std::min(desiredWidth, std::max(0, width - 2 * cardPaddingX));
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - desiredWidth);
        const int descHeight = descLabel ? PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, descText) : 0;

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        pushCard(card);

        if (label)
        {
            SetWindowTextW(label, labelText.data());
            SetWindowPos(label, nullptr, card.left + cardPaddingX, card.top + cardPaddingY, textWidth, titleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        if (descLabel)
        {
            SetWindowTextW(descLabel, descText.data());
            SetWindowPos(descLabel,
                         nullptr,
                         card.left + cardPaddingX,
                         card.top + cardPaddingY + titleHeight + cardGapY,
                         textWidth,
                         std::max(0, descHeight),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(descLabel, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
        }

        const int inputX       = card.right - cardPaddingX - desiredWidth;
        const int inputY       = card.top + (cardHeight - rowHeight) / 2;
        const int framePadding = (frame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;

        if (frame)
        {
            SetWindowPos(frame, nullptr, inputX, inputY, desiredWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        if (edit)
        {
            const int innerW = std::max(1, desiredWidth - 2 * framePadding);
            const int innerH = std::max(1, rowHeight - 2 * framePadding);
            SetWindowPos(edit, nullptr, inputX + framePadding, inputY + framePadding, innerW, innerH, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(edit, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        y += cardHeight + cardSpacingY;
    };

    const std::wstring labelBypassHelloText  = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CONNECTIONS_BYPASS_HELLO);
    const std::wstring labelHelloTimeoutText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CONNECTIONS_HELLO_TIMEOUT);

    const std::wstring descBypassHelloText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CONNECTIONS_BYPASS_HELLO);
    const std::wstring descHelloTimeoutText = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CONNECTIONS_HELLO_TIMEOUT);

    layoutToggleCard(state.advancedConnectionsBypassHelloLabel.get(),
                     labelBypassHelloText,
                     state.advancedConnectionsBypassHelloToggle.get(),
                     state.advancedConnectionsBypassHelloDescription.get(),
                     descBypassHelloText);
    layoutEditCard(state.advancedConnectionsHelloTimeoutLabel.get(),
                   labelHelloTimeoutText,
                   state.advancedConnectionsHelloTimeoutFrame.get(),
                   state.advancedConnectionsHelloTimeoutEdit.get(),
                   ThemedControls::ScaleDip(dpi, kMinToggleWidthDip),
                   state.advancedConnectionsHelloTimeoutDescription.get(),
                   descHelloTimeoutText);

    if (state.advancedMonitorHeader)
    {
        y += gapY;
        SetWindowTextW(state.advancedMonitorHeader.get(), LoadStringResource(nullptr, IDS_PREFS_ADV_HEADER_MONITOR).c_str());
        SetWindowPos(state.advancedMonitorHeader.get(), nullptr, x, y, width, headerHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.advancedMonitorHeader.get(), WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += headerHeight + gapY;
    }

    const std::wstring labelToolbarText      = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_TOOLBAR);
    const std::wstring labelLineNumbersText  = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_LINE_NUMBERS);
    const std::wstring labelAlwaysOnTopText  = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_ALWAYS_ON_TOP);
    const std::wstring labelShowIdsText      = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_SHOW_IDS);
    const std::wstring labelAutoScrollText   = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_AUTO_SCROLL);
    const std::wstring labelFilterPresetText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_PRESET);
    const std::wstring labelFilterMaskText   = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_MASK);
    const std::wstring labelFilterTextText   = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_TEXT);
    const std::wstring labelFilterErrorText  = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_ERROR);
    const std::wstring labelFilterWarnText   = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_WARNING);
    const std::wstring labelFilterInfoText   = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_INFO);
    const std::wstring labelFilterDebugText  = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_DEBUG);

    const std::wstring descToolbarText      = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_TOOLBAR);
    const std::wstring descLineNumbersText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_LINE_NUMBERS);
    const std::wstring descAlwaysOnTopText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_ALWAYS_ON_TOP);
    const std::wstring descShowIdsText      = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_SHOW_IDS);
    const std::wstring descAutoScrollText   = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_AUTO_SCROLL);
    const std::wstring descFilterPresetText = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILTER_PRESET);
    const std::wstring descFilterMaskText   = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILTER_MASK);
    const std::wstring descFilterTextText   = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILTER_TEXT);
    const std::wstring descFilterErrorText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILTER_ERROR);
    const std::wstring descFilterWarnText   = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILTER_WARNING);
    const std::wstring descFilterInfoText   = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILTER_INFO);
    const std::wstring descFilterDebugText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILTER_DEBUG);

    layoutToggleCard(state.advancedMonitorToolbarLabel.get(),
                     labelToolbarText,
                     state.advancedMonitorToolbarToggle.get(),
                     state.advancedMonitorToolbarDescription.get(),
                     descToolbarText);
    layoutToggleCard(state.advancedMonitorLineNumbersLabel.get(),
                     labelLineNumbersText,
                     state.advancedMonitorLineNumbersToggle.get(),
                     state.advancedMonitorLineNumbersDescription.get(),
                     descLineNumbersText);
    layoutToggleCard(state.advancedMonitorAlwaysOnTopLabel.get(),
                     labelAlwaysOnTopText,
                     state.advancedMonitorAlwaysOnTopToggle.get(),
                     state.advancedMonitorAlwaysOnTopDescription.get(),
                     descAlwaysOnTopText);
    layoutToggleCard(state.advancedMonitorShowIdsLabel.get(),
                     labelShowIdsText,
                     state.advancedMonitorShowIdsToggle.get(),
                     state.advancedMonitorShowIdsDescription.get(),
                     descShowIdsText);
    layoutToggleCard(state.advancedMonitorAutoScrollLabel.get(),
                     labelAutoScrollText,
                     state.advancedMonitorAutoScrollToggle.get(),
                     state.advancedMonitorAutoScrollDescription.get(),
                     descAutoScrollText);

    layoutFramedComboCard(state.advancedMonitorFilterPresetLabel.get(),
                          labelFilterPresetText,
                          state.advancedMonitorFilterPresetFrame.get(),
                          state.advancedMonitorFilterPresetCombo.get(),
                          state.advancedMonitorFilterPresetDescription.get(),
                          descFilterPresetText);

    layoutEditCard(state.advancedMonitorFilterMaskLabel.get(),
                   labelFilterMaskText,
                   state.advancedMonitorFilterMaskFrame.get(),
                   state.advancedMonitorFilterMaskEdit.get(),
                   ThemedControls::ScaleDip(dpi, kMinComboWidthDip),
                   state.advancedMonitorFilterMaskDescription.get(),
                   descFilterMaskText);

    layoutToggleCard(state.advancedMonitorFilterTextLabel.get(),
                     labelFilterTextText,
                     state.advancedMonitorFilterTextToggle.get(),
                     state.advancedMonitorFilterTextDescription.get(),
                     descFilterTextText);
    layoutToggleCard(state.advancedMonitorFilterErrorLabel.get(),
                     labelFilterErrorText,
                     state.advancedMonitorFilterErrorToggle.get(),
                     state.advancedMonitorFilterErrorDescription.get(),
                     descFilterErrorText);
    layoutToggleCard(state.advancedMonitorFilterWarningLabel.get(),
                     labelFilterWarnText,
                     state.advancedMonitorFilterWarningToggle.get(),
                     state.advancedMonitorFilterWarningDescription.get(),
                     descFilterWarnText);
    layoutToggleCard(state.advancedMonitorFilterInfoLabel.get(),
                     labelFilterInfoText,
                     state.advancedMonitorFilterInfoToggle.get(),
                     state.advancedMonitorFilterInfoDescription.get(),
                     descFilterInfoText);
    layoutToggleCard(state.advancedMonitorFilterDebugLabel.get(),
                     labelFilterDebugText,
                     state.advancedMonitorFilterDebugToggle.get(),
                     state.advancedMonitorFilterDebugDescription.get(),
                     descFilterDebugText);

    if (state.advancedCacheHeader)
    {
        y += gapY;
        SetWindowTextW(state.advancedCacheHeader.get(), LoadStringResource(nullptr, IDS_PREFS_ADV_HEADER_CACHE).c_str());
        SetWindowPos(state.advancedCacheHeader.get(), nullptr, x, y, width, headerHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.advancedCacheHeader.get(), WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += headerHeight + gapY;
    }

    const std::wstring labelCacheMaxBytesText    = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CACHE_DIR_MAX_BYTES);
    const std::wstring labelCacheMaxWatchersText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CACHE_DIR_MAX_WATCHERS);
    const std::wstring labelCacheMruWatchedText  = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CACHE_DIR_MRU_WATCHED);

    const std::wstring descCacheMaxBytesText    = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CACHE_DIR_MAX_BYTES);
    const std::wstring descCacheMaxWatchersText = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CACHE_DIR_MAX_WATCHERS);
    const std::wstring descCacheMruWatchedText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CACHE_DIR_MRU_WATCHED);

    layoutEditCard(state.advancedCacheDirectoryInfoMaxBytesLabel.get(),
                   labelCacheMaxBytesText,
                   state.advancedCacheDirectoryInfoMaxBytesFrame.get(),
                   state.advancedCacheDirectoryInfoMaxBytesEdit.get(),
                   ThemedControls::ScaleDip(dpi, kMediumComboWidthDip),
                   state.advancedCacheDirectoryInfoMaxBytesDescription.get(),
                   descCacheMaxBytesText);
    layoutEditCard(state.advancedCacheDirectoryInfoMaxWatchersLabel.get(),
                   labelCacheMaxWatchersText,
                   state.advancedCacheDirectoryInfoMaxWatchersFrame.get(),
                   state.advancedCacheDirectoryInfoMaxWatchersEdit.get(),
                   ThemedControls::ScaleDip(dpi, kMinToggleWidthDip),
                   state.advancedCacheDirectoryInfoMaxWatchersDescription.get(),
                   descCacheMaxWatchersText);
    layoutEditCard(state.advancedCacheDirectoryInfoMruWatchedLabel.get(),
                   labelCacheMruWatchedText,
                   state.advancedCacheDirectoryInfoMruWatchedFrame.get(),
                   state.advancedCacheDirectoryInfoMruWatchedEdit.get(),
                   ThemedControls::ScaleDip(dpi, kMinToggleWidthDip),
                   state.advancedCacheDirectoryInfoMruWatchedDescription.get(),
                   descCacheMruWatchedText);

    if (state.advancedFileOperationsHeader)
    {
        y += gapY;
        SetWindowTextW(state.advancedFileOperationsHeader.get(), LoadStringResource(nullptr, IDS_PREFS_ADV_HEADER_FILEOPS).c_str());
        SetWindowPos(state.advancedFileOperationsHeader.get(), nullptr, x, y, width, headerHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.advancedFileOperationsHeader.get(), WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += headerHeight + gapY;
    }

    const std::wstring labelMaxDiagnosticsLogFilesText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILEOPS_MAX_DIAG_LOG_FILES);
    const std::wstring descMaxDiagnosticsLogFilesText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILEOPS_MAX_DIAG_LOG_FILES);

    layoutEditCard(state.advancedFileOperationsMaxDiagnosticsLogFilesLabel.get(),
                   labelMaxDiagnosticsLogFilesText,
                   state.advancedFileOperationsMaxDiagnosticsLogFilesFrame.get(),
                   state.advancedFileOperationsMaxDiagnosticsLogFilesEdit.get(),
                   ThemedControls::ScaleDip(dpi, kMinToggleWidthDip),
                   state.advancedFileOperationsMaxDiagnosticsLogFilesDescription.get(),
                   descMaxDiagnosticsLogFilesText);

    const std::wstring labelDiagnosticsInfoText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILEOPS_DIAG_INFO);
    const std::wstring descDiagnosticsInfoText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILEOPS_DIAG_INFO);
    layoutToggleCard(state.advancedFileOperationsDiagnosticsInfoLabel.get(),
                     labelDiagnosticsInfoText,
                     state.advancedFileOperationsDiagnosticsInfoToggle.get(),
                     state.advancedFileOperationsDiagnosticsInfoDescription.get(),
                     descDiagnosticsInfoText);

    const std::wstring labelDiagnosticsDebugText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILEOPS_DIAG_DEBUG);
    const std::wstring descDiagnosticsDebugText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILEOPS_DIAG_DEBUG);
    layoutToggleCard(state.advancedFileOperationsDiagnosticsDebugLabel.get(),
                     labelDiagnosticsDebugText,
                     state.advancedFileOperationsDiagnosticsDebugToggle.get(),
                     state.advancedFileOperationsDiagnosticsDebugDescription.get(),
                     descDiagnosticsDebugText);
}

void AdvancedPane::Refresh(HWND /*host*/, PreferencesDialogState& state) noexcept
{
    const auto& connections = GetConnectionsSettingsOrDefault(state.workingSettings);
    PrefsUi::SetTwoStateToggleState(state.advancedConnectionsBypassHelloToggle, state.theme.systemHighContrast, connections.bypassWindowsHello);
    if (state.advancedConnectionsHelloTimeoutEdit)
    {
        std::wstring text;
        text = std::to_wstring(connections.windowsHelloReauthTimeoutMinute);
        SetWindowTextW(state.advancedConnectionsHelloTimeoutEdit.get(), text.c_str());
    }

    const auto& monitor           = GetMonitorSettingsOrDefault(state.workingSettings);
    const uint32_t mask           = monitor.filter.mask & 31u;
    const bool customFilter       = (monitor.filter.preset == Common::Settings::MonitorFilterPreset::Custom);
    const BOOL enableCustomFilter = customFilter ? TRUE : FALSE;

    const auto setEnabledAndInvalidate = [&](const auto& hwndLike, BOOL enabled) noexcept
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

        if (! hwnd)
        {
            return;
        }

        EnableWindow(hwnd, enabled);
        InvalidateRect(hwnd, nullptr, TRUE);
    };

    PrefsUi::SetTwoStateToggleState(state.advancedMonitorToolbarToggle, state.theme.systemHighContrast, monitor.menu.toolbarVisible);
    PrefsUi::SetTwoStateToggleState(state.advancedMonitorLineNumbersToggle, state.theme.systemHighContrast, monitor.menu.lineNumbersVisible);
    PrefsUi::SetTwoStateToggleState(state.advancedMonitorAlwaysOnTopToggle, state.theme.systemHighContrast, monitor.menu.alwaysOnTop);
    PrefsUi::SetTwoStateToggleState(state.advancedMonitorShowIdsToggle, state.theme.systemHighContrast, monitor.menu.showIds);
    PrefsUi::SetTwoStateToggleState(state.advancedMonitorAutoScrollToggle, state.theme.systemHighContrast, monitor.menu.autoScroll);
    PrefsUi::SelectComboItemByData(state.advancedMonitorFilterPresetCombo, static_cast<LPARAM>(static_cast<int>(monitor.filter.preset)));

    if (state.advancedMonitorFilterMaskEdit)
    {
        std::wstring text;
        text = std::to_wstring(mask);
        SetWindowTextW(state.advancedMonitorFilterMaskEdit.get(), text.c_str());
        setEnabledAndInvalidate(state.advancedMonitorFilterMaskEdit, enableCustomFilter);
    }
    setEnabledAndInvalidate(state.advancedMonitorFilterMaskLabel, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterMaskDescription, enableCustomFilter);

    PrefsUi::SetTwoStateToggleState(state.advancedMonitorFilterTextToggle, state.theme.systemHighContrast, HasFlag(mask, MonitorFilterBit::Text));
    PrefsUi::SetTwoStateToggleState(state.advancedMonitorFilterErrorToggle, state.theme.systemHighContrast, HasFlag(mask, MonitorFilterBit::Error));
    PrefsUi::SetTwoStateToggleState(state.advancedMonitorFilterWarningToggle, state.theme.systemHighContrast, HasFlag(mask, MonitorFilterBit::Warning));
    PrefsUi::SetTwoStateToggleState(state.advancedMonitorFilterInfoToggle, state.theme.systemHighContrast, HasFlag(mask, MonitorFilterBit::Info));
    PrefsUi::SetTwoStateToggleState(state.advancedMonitorFilterDebugToggle, state.theme.systemHighContrast, HasFlag(mask, MonitorFilterBit::Debug));

    setEnabledAndInvalidate(state.advancedMonitorFilterTextToggle, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterTextLabel, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterTextDescription, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterErrorToggle, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterErrorLabel, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterErrorDescription, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterWarningToggle, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterWarningLabel, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterWarningDescription, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterInfoToggle, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterInfoLabel, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterInfoDescription, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterDebugToggle, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterDebugLabel, enableCustomFilter);
    setEnabledAndInvalidate(state.advancedMonitorFilterDebugDescription, enableCustomFilter);

    const auto& cache = GetCacheSettingsOrDefault(state.workingSettings);

    if (state.advancedCacheDirectoryInfoMaxBytesEdit)
    {
        std::wstring text;
        if (cache.directoryInfo.maxBytes.has_value() && cache.directoryInfo.maxBytes.value() > 0)
        {
            text = FormatCacheBytes(cache.directoryInfo.maxBytes.value());
        }
        SetWindowTextW(state.advancedCacheDirectoryInfoMaxBytesEdit.get(), text.c_str());
    }

    if (state.advancedCacheDirectoryInfoMaxWatchersEdit)
    {
        std::wstring text;
        if (cache.directoryInfo.maxWatchers.has_value())
        {
            text = std::to_wstring(cache.directoryInfo.maxWatchers.value());
        }
        SetWindowTextW(state.advancedCacheDirectoryInfoMaxWatchersEdit.get(), text.c_str());
    }

    if (state.advancedCacheDirectoryInfoMruWatchedEdit)
    {
        std::wstring text;
        if (cache.directoryInfo.mruWatched.has_value())
        {
            text = std::to_wstring(cache.directoryInfo.mruWatched.value());
        }
        SetWindowTextW(state.advancedCacheDirectoryInfoMruWatchedEdit.get(), text.c_str());
    }

    const auto& fileOperations = GetFileOperationsSettingsOrDefault(state.workingSettings);
    PrefsUi::SetTwoStateToggleState(state.advancedFileOperationsDiagnosticsInfoToggle, state.theme.systemHighContrast, fileOperations.diagnosticsInfoEnabled);
    PrefsUi::SetTwoStateToggleState(state.advancedFileOperationsDiagnosticsDebugToggle, state.theme.systemHighContrast, fileOperations.diagnosticsDebugEnabled);
    if (state.advancedFileOperationsMaxDiagnosticsLogFilesEdit)
    {
        std::wstring text;
        text = std::to_wstring(fileOperations.maxDiagnosticsLogFiles);
        SetWindowTextW(state.advancedFileOperationsMaxDiagnosticsLogFilesEdit.get(), text.c_str());
    }
}

bool AdvancedPane::HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept
{
    if (commandId == IDC_PREFS_ADV_CONNECTIONS_HELLO_TIMEOUT_EDIT)
    {
        if (notifyCode == EN_CHANGE)
        {
            const std::wstring text = hwndCtl ? PrefsUi::GetWindowTextString(hwndCtl) : PrefsUi::GetWindowTextString(state.advancedConnectionsHelloTimeoutEdit);
            const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
            if (trimmed.empty())
            {
                return true;
            }

            const auto valueOpt = PrefsUi::TryParseUInt32(trimmed);
            if (! valueOpt.has_value())
            {
                return true;
            }

            const Common::Settings::ConnectionsSettings defaults{};
            const uint32_t value = valueOpt.value();
            if (! state.workingSettings.connections.has_value() && value == defaults.windowsHelloReauthTimeoutMinute)
            {
                return true;
            }

            auto* connections = EnsureWorkingConnectionsSettings(state.workingSettings);
            if (! connections)
            {
                return true;
            }

            if (connections->windowsHelloReauthTimeoutMinute != value)
            {
                connections->windowsHelloReauthTimeoutMinute = value;
                MaybeResetWorkingConnectionsSettingsIfEmpty(state.workingSettings);
                SetDirty(GetParent(host), state);
            }
            return true;
        }

        if (notifyCode == EN_KILLFOCUS)
        {
            const std::wstring text = hwndCtl ? PrefsUi::GetWindowTextString(hwndCtl) : PrefsUi::GetWindowTextString(state.advancedConnectionsHelloTimeoutEdit);
            const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);

            const Common::Settings::ConnectionsSettings defaults{};
            uint32_t value = defaults.windowsHelloReauthTimeoutMinute;
            if (const auto valueOpt = PrefsUi::TryParseUInt32(trimmed); valueOpt.has_value())
            {
                value = valueOpt.value();
            }

            if (! state.workingSettings.connections.has_value() && value == defaults.windowsHelloReauthTimeoutMinute)
            {
                Refresh(host, state);
                return true;
            }

            if (auto* connections = EnsureWorkingConnectionsSettings(state.workingSettings))
            {
                if (connections->windowsHelloReauthTimeoutMinute != value)
                {
                    connections->windowsHelloReauthTimeoutMinute = value;
                    MaybeResetWorkingConnectionsSettingsIfEmpty(state.workingSettings);
                    SetDirty(GetParent(host), state);
                }
            }

            Refresh(host, state);
            return true;
        }

        return false;
    }

    if (commandId == IDC_PREFS_ADV_MONITOR_FILTER_PRESET_COMBO && notifyCode == CBN_SELCHANGE)
    {
        const auto dataOpt = PrefsUi::TryGetSelectedComboItemData(state.advancedMonitorFilterPresetCombo);
        if (! dataOpt.has_value())
        {
            return true;
        }

        const int value = static_cast<int>(dataOpt.value());
        if (value < static_cast<int>(Common::Settings::MonitorFilterPreset::Custom) ||
            value > static_cast<int>(Common::Settings::MonitorFilterPreset::AllTypes))
        {
            return true;
        }

        auto* monitor = EnsureWorkingMonitorSettings(state.workingSettings);
        if (! monitor)
        {
            return true;
        }

        const auto preset      = static_cast<Common::Settings::MonitorFilterPreset>(value);
        monitor->filter.preset = preset;
        switch (preset)
        {
            case Common::Settings::MonitorFilterPreset::ErrorsOnly: monitor->filter.mask = static_cast<uint32_t>(MonitorFilterBit::Error); break;
            case Common::Settings::MonitorFilterPreset::ErrorsWarnings: monitor->filter.mask = MonitorFilterBit::Error | MonitorFilterBit::Warning; break;
            case Common::Settings::MonitorFilterPreset::AllTypes:
                monitor->filter.mask =
                    MonitorFilterBit::Text | MonitorFilterBit::Error | MonitorFilterBit::Warning | MonitorFilterBit::Info | MonitorFilterBit::Debug;
                break;
            case Common::Settings::MonitorFilterPreset::Custom:
            default: break;
        }
        SetDirty(GetParent(host), state);
        Refresh(host, state);
        return true;
    }

    if (commandId == IDC_PREFS_ADV_MONITOR_FILTER_MASK_EDIT)
    {
        if (notifyCode == EN_CHANGE)
        {
            const std::wstring text = hwndCtl ? PrefsUi::GetWindowTextString(hwndCtl) : PrefsUi::GetWindowTextString(state.advancedMonitorFilterMaskEdit);
            const auto valueOpt     = PrefsUi::TryParseUInt32(text);
            if (! valueOpt.has_value())
            {
                return true;
            }

            const uint32_t value = valueOpt.value();
            if (value > 31u)
            {
                return true;
            }

            auto* monitor = EnsureWorkingMonitorSettings(state.workingSettings);
            if (! monitor)
            {
                return true;
            }

            monitor->filter.mask = value;
            SetDirty(GetParent(host), state);
            return true;
        }

        if (notifyCode == EN_KILLFOCUS)
        {
            const std::wstring text = hwndCtl ? PrefsUi::GetWindowTextString(hwndCtl) : PrefsUi::GetWindowTextString(state.advancedMonitorFilterMaskEdit);
            const auto valueOpt     = PrefsUi::TryParseUInt32(text);
            if (valueOpt.has_value())
            {
                const uint32_t value = std::min(valueOpt.value(), 31u);
                if (auto* monitor = EnsureWorkingMonitorSettings(state.workingSettings))
                {
                    monitor->filter.mask = value;
                    SetDirty(GetParent(host), state);
                }
            }

            Refresh(host, state);
            return true;
        }

        return false;
    }

    if (commandId == IDC_PREFS_ADV_FILEOPS_MAX_DIAG_LOG_FILES_EDIT)
    {
        if (notifyCode == EN_CHANGE)
        {
            const std::wstring text =
                hwndCtl ? PrefsUi::GetWindowTextString(hwndCtl) : PrefsUi::GetWindowTextString(state.advancedFileOperationsMaxDiagnosticsLogFilesEdit);
            const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
            if (trimmed.empty())
            {
                return true;
            }

            const auto valueOpt = PrefsUi::TryParseUInt32(trimmed);
            if (! valueOpt.has_value() || valueOpt.value() == 0)
            {
                return true;
            }

            const Common::Settings::FileOperationsSettings defaults{};
            const uint32_t value = valueOpt.value();
            if (! state.workingSettings.fileOperations.has_value() && value == defaults.maxDiagnosticsLogFiles)
            {
                return true;
            }

            auto* fileOperations = EnsureWorkingFileOperationsSettings(state.workingSettings);
            if (! fileOperations)
            {
                return true;
            }

            if (fileOperations->maxDiagnosticsLogFiles != value)
            {
                fileOperations->maxDiagnosticsLogFiles = value;
                MaybeResetWorkingFileOperationsSettingsIfEmpty(state.workingSettings);
                SetDirty(GetParent(host), state);
            }
            return true;
        }

        if (notifyCode == EN_KILLFOCUS)
        {
            const std::wstring text =
                hwndCtl ? PrefsUi::GetWindowTextString(hwndCtl) : PrefsUi::GetWindowTextString(state.advancedFileOperationsMaxDiagnosticsLogFilesEdit);
            const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);

            const Common::Settings::FileOperationsSettings defaults{};
            uint32_t value = defaults.maxDiagnosticsLogFiles;
            if (const auto valueOpt = PrefsUi::TryParseUInt32(trimmed); valueOpt.has_value() && valueOpt.value() > 0)
            {
                value = valueOpt.value();
            }

            if (! state.workingSettings.fileOperations.has_value() && value == defaults.maxDiagnosticsLogFiles)
            {
                Refresh(host, state);
                return true;
            }

            if (auto* fileOperations = EnsureWorkingFileOperationsSettings(state.workingSettings))
            {
                if (fileOperations->maxDiagnosticsLogFiles != value)
                {
                    fileOperations->maxDiagnosticsLogFiles = value;
                    MaybeResetWorkingFileOperationsSettingsIfEmpty(state.workingSettings);
                    SetDirty(GetParent(host), state);
                }
            }

            Refresh(host, state);
            return true;
        }

        return false;
    }

    const bool isCacheEdit = (commandId == IDC_PREFS_ADV_CACHE_DIR_MAX_BYTES_EDIT || commandId == IDC_PREFS_ADV_CACHE_DIR_MAX_WATCHERS_EDIT ||
                              commandId == IDC_PREFS_ADV_CACHE_DIR_MRU_WATCHED_EDIT);
    if (isCacheEdit)
    {
        if (notifyCode == EN_CHANGE || notifyCode == EN_KILLFOCUS)
        {
            const std::wstring text =
                hwndCtl ? PrefsUi::GetWindowTextString(hwndCtl) : PrefsUi::GetWindowTextString(GetDlgItem(host, static_cast<int>(commandId)));
            const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);

            const bool commit = (notifyCode == EN_KILLFOCUS);

            if (commandId == IDC_PREFS_ADV_CACHE_DIR_MAX_BYTES_EDIT)
            {
                if (trimmed.empty())
                {
                    if (state.workingSettings.cache.has_value())
                    {
                        state.workingSettings.cache->directoryInfo.maxBytes.reset();
                        MaybeResetWorkingCacheSettingsIfEmpty(state.workingSettings);
                        SetDirty(GetParent(host), state);
                    }

                    if (commit)
                    {
                        Refresh(host, state);
                    }
                    return true;
                }

                const auto bytesOpt = TryParseCacheBytes(trimmed);
                if (! bytesOpt.has_value())
                {
                    if (commit)
                    {
                        Refresh(host, state);
                    }
                    return true;
                }

                auto* cache = EnsureWorkingCacheSettings(state.workingSettings);
                if (! cache)
                {
                    return true;
                }

                const uint64_t bytes = bytesOpt.value();
                if (bytes == 0)
                {
                    cache->directoryInfo.maxBytes.reset();
                }
                else
                {
                    cache->directoryInfo.maxBytes = bytes;
                }

                MaybeResetWorkingCacheSettingsIfEmpty(state.workingSettings);
                SetDirty(GetParent(host), state);

                if (commit)
                {
                    Refresh(host, state);
                }
                return true;
            }

            if (commandId == IDC_PREFS_ADV_CACHE_DIR_MAX_WATCHERS_EDIT)
            {
                if (trimmed.empty())
                {
                    if (state.workingSettings.cache.has_value() && state.workingSettings.cache->directoryInfo.maxWatchers.has_value())
                    {
                        state.workingSettings.cache->directoryInfo.maxWatchers.reset();
                        MaybeResetWorkingCacheSettingsIfEmpty(state.workingSettings);
                        SetDirty(GetParent(host), state);
                    }

                    if (commit)
                    {
                        Refresh(host, state);
                    }
                    return true;
                }

                const auto valueOpt = PrefsUi::TryParseUInt32(trimmed);
                if (! valueOpt.has_value())
                {
                    if (commit)
                    {
                        Refresh(host, state);
                    }
                    return true;
                }

                auto* cache = EnsureWorkingCacheSettings(state.workingSettings);
                if (! cache)
                {
                    return true;
                }

                const uint32_t value = valueOpt.value();
                if (! cache->directoryInfo.maxWatchers.has_value() || cache->directoryInfo.maxWatchers.value() != value)
                {
                    cache->directoryInfo.maxWatchers = value;
                    MaybeResetWorkingCacheSettingsIfEmpty(state.workingSettings);
                    SetDirty(GetParent(host), state);
                }

                if (commit)
                {
                    Refresh(host, state);
                }
                return true;
            }

            if (commandId == IDC_PREFS_ADV_CACHE_DIR_MRU_WATCHED_EDIT)
            {
                if (trimmed.empty())
                {
                    if (state.workingSettings.cache.has_value() && state.workingSettings.cache->directoryInfo.mruWatched.has_value())
                    {
                        state.workingSettings.cache->directoryInfo.mruWatched.reset();
                        MaybeResetWorkingCacheSettingsIfEmpty(state.workingSettings);
                        SetDirty(GetParent(host), state);
                    }

                    if (commit)
                    {
                        Refresh(host, state);
                    }
                    return true;
                }

                const auto valueOpt = PrefsUi::TryParseUInt32(trimmed);
                if (! valueOpt.has_value())
                {
                    if (commit)
                    {
                        Refresh(host, state);
                    }
                    return true;
                }

                auto* cache = EnsureWorkingCacheSettings(state.workingSettings);
                if (! cache)
                {
                    return true;
                }

                const uint32_t value = valueOpt.value();
                if (! cache->directoryInfo.mruWatched.has_value() || cache->directoryInfo.mruWatched.value() != value)
                {
                    cache->directoryInfo.mruWatched = value;
                    MaybeResetWorkingCacheSettingsIfEmpty(state.workingSettings);
                    SetDirty(GetParent(host), state);
                }

                if (commit)
                {
                    Refresh(host, state);
                }
                return true;
            }

            return false;
        }

        return false;
    }

    if (notifyCode == BN_CLICKED)
    {
        const bool isToggle = (commandId == IDC_PREFS_ADV_CONNECTIONS_BYPASS_HELLO_TOGGLE || commandId == IDC_PREFS_ADV_MONITOR_TOOLBAR_TOGGLE ||
                               commandId == IDC_PREFS_ADV_MONITOR_LINE_NUMBERS_TOGGLE || commandId == IDC_PREFS_ADV_MONITOR_ALWAYS_ON_TOP_TOGGLE ||
                               commandId == IDC_PREFS_ADV_MONITOR_SHOW_IDS_TOGGLE || commandId == IDC_PREFS_ADV_MONITOR_AUTO_SCROLL_TOGGLE ||
                               commandId == IDC_PREFS_ADV_MONITOR_FILTER_TEXT_TOGGLE || commandId == IDC_PREFS_ADV_MONITOR_FILTER_ERROR_TOGGLE ||
                               commandId == IDC_PREFS_ADV_MONITOR_FILTER_WARNING_TOGGLE || commandId == IDC_PREFS_ADV_MONITOR_FILTER_INFO_TOGGLE ||
                               commandId == IDC_PREFS_ADV_MONITOR_FILTER_DEBUG_TOGGLE || commandId == IDC_PREFS_ADV_FILEOPS_DIAG_INFO_TOGGLE ||
                               commandId == IDC_PREFS_ADV_FILEOPS_DIAG_DEBUG_TOGGLE);
        if (! isToggle)
        {
            return false;
        }

        if (! hwndCtl)
        {
            return true;
        }

        const bool ownerDraw = (GetWindowLongPtrW(hwndCtl, GWL_STYLE) & BS_TYPEMASK) == BS_OWNERDRAW;
        if (ownerDraw)
        {
            const bool current = PrefsUi::GetTwoStateToggleState(hwndCtl, false);
            PrefsUi::SetTwoStateToggleState(hwndCtl, false, ! current);
        }

        const bool toggledOn = PrefsUi::GetTwoStateToggleState(hwndCtl, state.theme.systemHighContrast);

        if (commandId == IDC_PREFS_ADV_CONNECTIONS_BYPASS_HELLO_TOGGLE)
        {
            auto* connections = EnsureWorkingConnectionsSettings(state.workingSettings);
            if (! connections)
            {
                return true;
            }

            connections->bypassWindowsHello = toggledOn;
            MaybeResetWorkingConnectionsSettingsIfEmpty(state.workingSettings);
            SetDirty(GetParent(host), state);
            Refresh(host, state);
            return true;
        }

        if (commandId == IDC_PREFS_ADV_FILEOPS_DIAG_INFO_TOGGLE || commandId == IDC_PREFS_ADV_FILEOPS_DIAG_DEBUG_TOGGLE)
        {
            auto* fileOperations = EnsureWorkingFileOperationsSettings(state.workingSettings);
            if (! fileOperations)
            {
                return true;
            }

            switch (commandId)
            {
                case IDC_PREFS_ADV_FILEOPS_DIAG_INFO_TOGGLE: fileOperations->diagnosticsInfoEnabled = toggledOn; break;
                case IDC_PREFS_ADV_FILEOPS_DIAG_DEBUG_TOGGLE: fileOperations->diagnosticsDebugEnabled = toggledOn; break;
                default: break;
            }

            MaybeResetWorkingFileOperationsSettingsIfEmpty(state.workingSettings);
            SetDirty(GetParent(host), state);
            Refresh(host, state);
            return true;
        }

        auto* monitor = EnsureWorkingMonitorSettings(state.workingSettings);
        if (! monitor)
        {
            return true;
        }

        const auto updateFilterBit = [&](uint32_t bit) noexcept
        {
            monitor->filter.preset = Common::Settings::MonitorFilterPreset::Custom;
            uint32_t mask          = monitor->filter.mask & 31u;
            if (toggledOn)
            {
                mask |= bit;
            }
            else
            {
                mask &= ~bit;
            }
            monitor->filter.mask = mask & 31u;
        };

        switch (commandId)
        {
            case IDC_PREFS_ADV_MONITOR_TOOLBAR_TOGGLE: monitor->menu.toolbarVisible = toggledOn; break;
            case IDC_PREFS_ADV_MONITOR_LINE_NUMBERS_TOGGLE: monitor->menu.lineNumbersVisible = toggledOn; break;
            case IDC_PREFS_ADV_MONITOR_ALWAYS_ON_TOP_TOGGLE: monitor->menu.alwaysOnTop = toggledOn; break;
            case IDC_PREFS_ADV_MONITOR_SHOW_IDS_TOGGLE: monitor->menu.showIds = toggledOn; break;
            case IDC_PREFS_ADV_MONITOR_AUTO_SCROLL_TOGGLE: monitor->menu.autoScroll = toggledOn; break;
            case IDC_PREFS_ADV_MONITOR_FILTER_TEXT_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Text)); break;
            case IDC_PREFS_ADV_MONITOR_FILTER_ERROR_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Error)); break;
            case IDC_PREFS_ADV_MONITOR_FILTER_WARNING_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Warning)); break;
            case IDC_PREFS_ADV_MONITOR_FILTER_INFO_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Info)); break;
            case IDC_PREFS_ADV_MONITOR_FILTER_DEBUG_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Debug)); break;
            default: break;
        }

        SetDirty(GetParent(host), state);
        Refresh(host, state);
        return true;
    }

    return false;
}
