// Preferences.Advanced.cpp

#include "Framework.h"

#include "Preferences.Advanced.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cwctype>
#include <limits>
#include <string>
#include <vector>

#include "DxUi/DxUi.h"
#include "Helpers.h"
#include "UiMetrics.h"
#include "WindowMessages.h"
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
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ButtonVariant;
using RedSalamander::DxUi::CardPanel;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;

constexpr std::array<UINT, 3> kAdvancedHeaderStringIds = {{
    IDS_PREFS_ADV_HEADER_CONNECTIONS_HELLO,
    IDS_PREFS_ADV_HEADER_CACHE,
    IDS_PREFS_ADV_HEADER_FILEOPS,
}};

constexpr std::array<UINT, 4> kAdvancedToggleLabelStringIds = {{
    IDS_PREFS_ADV_LABEL_CONNECTIONS_BYPASS_HELLO,
    IDS_PREFS_ADV_LABEL_CONNECTIONS_ALLOW_INSECURE_TLS_AUTOMATION,
    IDS_PREFS_ADV_LABEL_FILEOPS_DIAG_INFO,
    IDS_PREFS_ADV_LABEL_FILEOPS_DIAG_DEBUG,
}};

constexpr std::array<UINT, 4> kAdvancedToggleDescriptionStringIds = {{
    IDS_PREFS_ADV_DESC_CONNECTIONS_BYPASS_HELLO,
    IDS_PREFS_ADV_DESC_CONNECTIONS_ALLOW_INSECURE_TLS_AUTOMATION,
    IDS_PREFS_ADV_DESC_FILEOPS_DIAG_INFO,
    IDS_PREFS_ADV_DESC_FILEOPS_DIAG_DEBUG,
}};

constexpr std::array<UINT, 4> kAdvancedToggleCommandIds = {{
    IDC_PREFS_ADV_CONNECTIONS_BYPASS_HELLO_TOGGLE,
    IDC_PREFS_ADV_CONNECTIONS_ALLOW_INSECURE_TLS_AUTOMATION_TOGGLE,
    IDC_PREFS_ADV_FILEOPS_DIAG_INFO_TOGGLE,
    IDC_PREFS_ADV_FILEOPS_DIAG_DEBUG_TOGGLE,
}};

constexpr std::array<UINT, 5> kAdvancedInputLabelStringIds = {{
    IDS_PREFS_ADV_LABEL_CONNECTIONS_HELLO_TIMEOUT,
    IDS_PREFS_ADV_LABEL_CACHE_DIR_MAX_BYTES,
    IDS_PREFS_ADV_LABEL_CACHE_DIR_MAX_WATCHERS,
    IDS_PREFS_ADV_LABEL_CACHE_DIR_MRU_WATCHED,
    IDS_PREFS_ADV_LABEL_FILEOPS_MAX_DIAG_LOG_FILES,
}};

constexpr std::array<UINT, 5> kAdvancedInputDescriptionStringIds = {{
    IDS_PREFS_ADV_DESC_CONNECTIONS_HELLO_TIMEOUT,
    IDS_PREFS_ADV_DESC_CACHE_DIR_MAX_BYTES,
    IDS_PREFS_ADV_DESC_CACHE_DIR_MAX_WATCHERS,
    IDS_PREFS_ADV_DESC_CACHE_DIR_MRU_WATCHED,
    IDS_PREFS_ADV_DESC_FILEOPS_MAX_DIAG_LOG_FILES,
}};

constexpr std::array<UINT, 5> kAdvancedEditCommandIds = {{
    IDC_PREFS_ADV_CONNECTIONS_HELLO_TIMEOUT_EDIT,
    IDC_PREFS_ADV_CACHE_DIR_MAX_BYTES_EDIT,
    IDC_PREFS_ADV_CACHE_DIR_MAX_WATCHERS_EDIT,
    IDC_PREFS_ADV_CACHE_DIR_MRU_WATCHED_EDIT,
    IDC_PREFS_ADV_FILEOPS_MAX_DIAG_LOG_FILES_EDIT,
}};

constexpr std::array<size_t, 5> kAdvancedEditMaxChars = {{
    10u,
    24u,
    10u,
    10u,
    10u,
}};

constexpr std::array<bool, 5> kAdvancedEditDigitsOnly = {{
    true,
    false,
    true,
    true,
    true,
}};

[[nodiscard]] std::wstring FilterDigitsLimited(std::wstring_view text, size_t maxChars) noexcept
{
    std::wstring filtered;
    filtered.reserve(std::min(text.size(), maxChars));
    for (const wchar_t ch : text)
    {
        if (! std::iswdigit(static_cast<wint_t>(ch)))
        {
            continue;
        }

        filtered.push_back(ch);
        if (filtered.size() >= maxChars)
        {
            break;
        }
    }
    return filtered;
}

} // namespace

struct AdvancedToggleCardPageDx
{
    CardPanel* card    = nullptr;
    Label* label       = nullptr;
    Label* description = nullptr;
    Toggle* toggle     = nullptr;
};

struct AdvancedInputCardPageDx
{
    CardPanel* card    = nullptr;
    Label* label       = nullptr;
    Label* description = nullptr;
    TextField* edit    = nullptr;
};

struct AdvancedLinkCardPageDx
{
    CardPanel* card    = nullptr;
    Label* label       = nullptr;
    Label* description = nullptr;
    Button* link       = nullptr;
};

struct AdvancedDxPage
{
    AdvancedDxPage()                                 = default;
    AdvancedDxPage(const AdvancedDxPage&)            = delete;
    AdvancedDxPage& operator=(const AdvancedDxPage&) = delete;
    AdvancedDxPage(AdvancedDxPage&&)                 = delete;
    AdvancedDxPage& operator=(AdvancedDxPage&&)      = delete;

    std::array<Label*, kAdvancedHeaderStringIds.size()> headers{};
    std::array<AdvancedToggleCardPageDx, kAdvancedToggleLabelStringIds.size()> toggleCards{};
    std::array<AdvancedInputCardPageDx, kAdvancedInputLabelStringIds.size()> inputCards{};
    AdvancedLinkCardPageDx settingsFileCard{};

    void Detach() noexcept
    {
        headers.fill(nullptr);
        for (auto& toggleCard : toggleCards)
        {
            toggleCard = {};
        }
        for (auto& inputCard : inputCards)
        {
            inputCard = {};
        }
        settingsFileCard = {};
    }
};

struct AdvancedPane::DxState
{
    DxState()                          = default;
    DxState(const DxState&)            = delete;
    DxState& operator=(const DxState&) = delete;
    DxState(DxState&&)                 = delete;
    DxState& operator=(DxState&&)      = delete;

    AdvancedDxPage page;

    void Detach() noexcept
    {
        page.Detach();
    }
};

AdvancedPane::AdvancedPane() = default;

AdvancedPane::~AdvancedPane()
{
    DetachDxHosts();
}

void AdvancedPane::OnVisibilityChanged(bool visible) noexcept
{
    static_cast<void>(visible);
}

void AdvancedPane::Destroy(PreferencesDialogState& state) noexcept
{
    static_cast<void>(state);
    DetachDxHosts();
    _pageHost = nullptr;
}

bool AdvancedPane::EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept
{
    _pageHostDx      = state.pageHostDxHost;
    _pageContentRoot = state.pageHostDxContentRootControl;
    if (! _pageHostDx || ! _pageContentRoot)
    {
        return false;
    }

    if (_dxState && PrefsUi::HasRetainedDxChildren(_pageContentRoot))
    {
        ApplyDxTheme(state);
        SyncDxControlsFromState(state);
        return true;
    }

    auto dxState = std::make_unique<DxState>();
    _pageHostDx->ResetInteractionState();
    _pageContentRoot->ClearChildren();

    auto* root = _pageContentRoot;

    for (Label*& header : dxState->page.headers)
    {
        header = root->AddChild<Label>();
        header->SetFontRole(FontRole::Header);
    }

    for (size_t i = 0; i < dxState->page.toggleCards.size(); ++i)
    {
        auto& card = dxState->page.toggleCards[i];
        card.card  = root->AddChild<CardPanel>();
        card.label = root->AddChild<Label>();
        card.label->SetFontRole(FontRole::Body);
        card.description = root->AddChild<Label>();
        card.description->SetFontRole(FontRole::Small);
        card.description->SetMultiline(true);
        card.toggle = root->AddChild<Toggle>();
        card.toggle->SetStateLabels(LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
        card.toggle->SetOnToggled([this, host = parent, commandId = kAdvancedToggleCommandIds[i]](bool checked) noexcept
        {
            if (! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* dialogState = PrefsUi::GetDialogState(host);
            if (! dialogState)
            {
                return;
            }

            bool changed = false;

            switch (commandId)
            {
                case IDC_PREFS_ADV_CONNECTIONS_BYPASS_HELLO_TOGGLE:
                case IDC_PREFS_ADV_CONNECTIONS_ALLOW_INSECURE_TLS_AUTOMATION_TOGGLE:
                {
                    if (auto* connections = EnsureWorkingConnectionsSettings(dialogState->workingSettings))
                    {
                        if (commandId == IDC_PREFS_ADV_CONNECTIONS_BYPASS_HELLO_TOGGLE)
                        {
                            changed                         = connections->bypassWindowsHello != checked;
                            connections->bypassWindowsHello = checked;
                        }
                        else
                        {
                            changed                                   = connections->allowInsecureTlsInAutomation != checked;
                            connections->allowInsecureTlsInAutomation = checked;
                        }
                        MaybeResetWorkingConnectionsSettingsIfEmpty(dialogState->workingSettings);
                    }
                    break;
                }
                case IDC_PREFS_ADV_FILEOPS_DIAG_INFO_TOGGLE:
                case IDC_PREFS_ADV_FILEOPS_DIAG_DEBUG_TOGGLE:
                {
                    if (auto* fileOperations = EnsureWorkingFileOperationsSettings(dialogState->workingSettings))
                    {
                        if (commandId == IDC_PREFS_ADV_FILEOPS_DIAG_INFO_TOGGLE)
                        {
                            changed                                = fileOperations->diagnosticsInfoEnabled != checked;
                            fileOperations->diagnosticsInfoEnabled = checked;
                        }
                        else
                        {
                            changed                                 = fileOperations->diagnosticsDebugEnabled != checked;
                            fileOperations->diagnosticsDebugEnabled = checked;
                        }
                        MaybeResetWorkingFileOperationsSettingsIfEmpty(dialogState->workingSettings);
                    }
                    break;
                }
                default: break;
            }

            if (changed)
            {
                if (HWND dlg = GetParent(host))
                {
                    SetDirty(dlg, *dialogState);
                }
            }
            Refresh(host, *dialogState);
        });
    }

    std::array<bool*, kAdvancedEditCommandIds.size()> syncFlags = {
        &_syncingDxHelloTimeoutEdit,
        &_syncingDxCacheDirectoryInfoMaxBytesEdit,
        &_syncingDxCacheDirectoryInfoMaxWatchersEdit,
        &_syncingDxCacheDirectoryInfoMruWatchedEdit,
        &_syncingDxFileOperationsMaxDiagnosticsLogFilesEdit,
    };

    size_t editIndex = 0u;
    for (size_t i = 0; i < dxState->page.inputCards.size(); ++i)
    {
        auto& card = dxState->page.inputCards[i];
        card.card  = root->AddChild<CardPanel>();
        card.label = root->AddChild<Label>();
        card.label->SetFontRole(FontRole::Body);
        card.description = root->AddChild<Label>();
        card.description->SetFontRole(FontRole::Small);
        card.description->SetMultiline(true);

        {
            const size_t currentEditIndex = editIndex;
            card.edit                     = root->AddChild<TextField>();
            card.edit->SetOnTextChanged([host       = parent,
                                         &syncFlag  = *syncFlags[currentEditIndex],
                                         commandId  = kAdvancedEditCommandIds[currentEditIndex],
                                         field      = card.edit,
                                         digitsOnly = kAdvancedEditDigitsOnly[currentEditIndex],
                                         maxChars   = kAdvancedEditMaxChars[currentEditIndex]](std::wstring_view text) noexcept
            {
                if (syncFlag || ! host || ! field || IsWindow(host) == FALSE)
                {
                    return;
                }

                std::wstring normalized;
                if (digitsOnly)
                {
                    normalized = FilterDigitsLimited(text, maxChars);
                }
                else
                {
                    normalized.assign(text.substr(0, std::min(text.size(), maxChars)));
                }

                if (normalized != text)
                {
                    syncFlag = true;
                    field->SetText(normalized);
                    syncFlag = false;
                }

                auto* dialogState = PrefsUi::GetDialogState(host);
                if (! dialogState)
                {
                    return;
                }

                const std::wstring_view trimmed = PrefsUi::TrimWhitespace(normalized);
                bool changed                    = false;

                switch (commandId)
                {
                    case IDC_PREFS_ADV_CONNECTIONS_HELLO_TIMEOUT_EDIT:
                    {
                        if (trimmed.empty())
                        {
                            break;
                        }
                        const auto valueOpt = PrefsUi::TryParseUInt32(trimmed);
                        if (! valueOpt.has_value())
                        {
                            break;
                        }
                        const Common::Settings::ConnectionsSettings defaults{};
                        const uint32_t value = valueOpt.value();
                        if (! dialogState->workingSettings.connections.has_value() && value == defaults.windowsHelloReauthTimeoutMinute)
                        {
                            break;
                        }
                        if (auto* connections = EnsureWorkingConnectionsSettings(dialogState->workingSettings))
                        {
                            changed                                      = (connections->windowsHelloReauthTimeoutMinute != value);
                            connections->windowsHelloReauthTimeoutMinute = value;
                            MaybeResetWorkingConnectionsSettingsIfEmpty(dialogState->workingSettings);
                        }
                        break;
                    }
                    case IDC_PREFS_ADV_CACHE_DIR_MAX_BYTES_EDIT:
                    {
                        if (trimmed.empty())
                        {
                            if (dialogState->workingSettings.cache.has_value())
                            {
                                dialogState->workingSettings.cache->directoryInfo.maxBytes.reset();
                                MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                changed = true;
                            }
                            break;
                        }
                        const auto bytesOpt = TryParseCacheBytes(trimmed);
                        if (! bytesOpt.has_value())
                        {
                            break;
                        }
                        if (auto* cache = EnsureWorkingCacheSettings(dialogState->workingSettings))
                        {
                            const uint64_t bytes = bytesOpt.value();
                            if (bytes == 0)
                            {
                                cache->directoryInfo.maxBytes.reset();
                            }
                            else
                            {
                                cache->directoryInfo.maxBytes = bytes;
                            }
                            MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                            changed = true;
                        }
                        break;
                    }
                    case IDC_PREFS_ADV_CACHE_DIR_MAX_WATCHERS_EDIT:
                    {
                        if (trimmed.empty())
                        {
                            if (dialogState->workingSettings.cache.has_value() && dialogState->workingSettings.cache->directoryInfo.maxWatchers.has_value())
                            {
                                dialogState->workingSettings.cache->directoryInfo.maxWatchers.reset();
                                MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                changed = true;
                            }
                            break;
                        }
                        const auto valueOpt = PrefsUi::TryParseUInt32(trimmed);
                        if (! valueOpt.has_value())
                        {
                            break;
                        }
                        if (auto* cache = EnsureWorkingCacheSettings(dialogState->workingSettings))
                        {
                            const uint32_t value = valueOpt.value();
                            if (! cache->directoryInfo.maxWatchers.has_value() || cache->directoryInfo.maxWatchers.value() != value)
                            {
                                cache->directoryInfo.maxWatchers = value;
                                MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                changed = true;
                            }
                        }
                        break;
                    }
                    case IDC_PREFS_ADV_CACHE_DIR_MRU_WATCHED_EDIT:
                    {
                        if (trimmed.empty())
                        {
                            if (dialogState->workingSettings.cache.has_value() && dialogState->workingSettings.cache->directoryInfo.mruWatched.has_value())
                            {
                                dialogState->workingSettings.cache->directoryInfo.mruWatched.reset();
                                MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                changed = true;
                            }
                            break;
                        }
                        const auto valueOpt = PrefsUi::TryParseUInt32(trimmed);
                        if (! valueOpt.has_value())
                        {
                            break;
                        }
                        if (auto* cache = EnsureWorkingCacheSettings(dialogState->workingSettings))
                        {
                            const uint32_t value = valueOpt.value();
                            if (! cache->directoryInfo.mruWatched.has_value() || cache->directoryInfo.mruWatched.value() != value)
                            {
                                cache->directoryInfo.mruWatched = value;
                                MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                changed = true;
                            }
                        }
                        break;
                    }
                    case IDC_PREFS_ADV_FILEOPS_MAX_DIAG_LOG_FILES_EDIT:
                    {
                        if (trimmed.empty())
                        {
                            break;
                        }
                        const auto valueOpt = PrefsUi::TryParseUInt32(trimmed);
                        if (! valueOpt.has_value() || valueOpt.value() == 0)
                        {
                            break;
                        }
                        const Common::Settings::FileOperationsSettings defaults{};
                        const uint32_t value = valueOpt.value();
                        if (! dialogState->workingSettings.fileOperations.has_value() && value == defaults.maxDiagnosticsLogFiles)
                        {
                            break;
                        }
                        if (auto* fileOperations = EnsureWorkingFileOperationsSettings(dialogState->workingSettings))
                        {
                            changed                                = (fileOperations->maxDiagnosticsLogFiles != value);
                            fileOperations->maxDiagnosticsLogFiles = value;
                            MaybeResetWorkingFileOperationsSettingsIfEmpty(dialogState->workingSettings);
                        }
                        break;
                    }
                    default: break;
                }

                if (changed)
                {
                    if (HWND dlg = GetParent(host))
                    {
                        SetDirty(dlg, *dialogState);
                    }
                }
            });
            card.edit->SetOnBlur([this, host = parent, commandId = kAdvancedEditCommandIds[currentEditIndex], field = card.edit]() noexcept
            {
                if (! host || IsWindow(host) == FALSE)
                {
                    return;
                }

                auto* dialogState = PrefsUi::GetDialogState(host);
                if (! dialogState)
                {
                    return;
                }

                bool changed = false;

                switch (commandId)
                {
                    case IDC_PREFS_ADV_CONNECTIONS_HELLO_TIMEOUT_EDIT:
                    {
                        const std::wstring text{field ? field->GetText() : std::wstring_view{}};
                        const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
                        const Common::Settings::ConnectionsSettings defaults{};
                        uint32_t value = defaults.windowsHelloReauthTimeoutMinute;
                        if (const auto valueOpt = PrefsUi::TryParseUInt32(trimmed); valueOpt.has_value())
                        {
                            value = valueOpt.value();
                        }
                        if (! dialogState->workingSettings.connections.has_value() && value == defaults.windowsHelloReauthTimeoutMinute)
                        {
                            break;
                        }
                        if (auto* connections = EnsureWorkingConnectionsSettings(dialogState->workingSettings))
                        {
                            changed                                      = (connections->windowsHelloReauthTimeoutMinute != value);
                            connections->windowsHelloReauthTimeoutMinute = value;
                            MaybeResetWorkingConnectionsSettingsIfEmpty(dialogState->workingSettings);
                        }
                        break;
                    }
                    case IDC_PREFS_ADV_CACHE_DIR_MAX_BYTES_EDIT:
                    {
                        const std::wstring text{field ? field->GetText() : std::wstring_view{}};
                        const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
                        if (trimmed.empty())
                        {
                            if (dialogState->workingSettings.cache.has_value())
                            {
                                dialogState->workingSettings.cache->directoryInfo.maxBytes.reset();
                                MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                changed = true;
                            }
                            break;
                        }
                        const auto bytesOpt = TryParseCacheBytes(trimmed);
                        if (bytesOpt.has_value())
                        {
                            if (auto* cache = EnsureWorkingCacheSettings(dialogState->workingSettings))
                            {
                                const uint64_t bytes = bytesOpt.value();
                                if (bytes == 0)
                                {
                                    cache->directoryInfo.maxBytes.reset();
                                }
                                else
                                {
                                    cache->directoryInfo.maxBytes = bytes;
                                }
                                MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                changed = true;
                            }
                        }
                        break;
                    }
                    case IDC_PREFS_ADV_CACHE_DIR_MAX_WATCHERS_EDIT:
                    {
                        const std::wstring text{field ? field->GetText() : std::wstring_view{}};
                        const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
                        if (trimmed.empty())
                        {
                            if (dialogState->workingSettings.cache.has_value() && dialogState->workingSettings.cache->directoryInfo.maxWatchers.has_value())
                            {
                                dialogState->workingSettings.cache->directoryInfo.maxWatchers.reset();
                                MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                changed = true;
                            }
                            break;
                        }
                        const auto valueOpt = PrefsUi::TryParseUInt32(trimmed);
                        if (valueOpt.has_value())
                        {
                            if (auto* cache = EnsureWorkingCacheSettings(dialogState->workingSettings))
                            {
                                const uint32_t value = valueOpt.value();
                                if (! cache->directoryInfo.maxWatchers.has_value() || cache->directoryInfo.maxWatchers.value() != value)
                                {
                                    cache->directoryInfo.maxWatchers = value;
                                    MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                    changed = true;
                                }
                            }
                        }
                        break;
                    }
                    case IDC_PREFS_ADV_CACHE_DIR_MRU_WATCHED_EDIT:
                    {
                        const std::wstring text{field ? field->GetText() : std::wstring_view{}};
                        const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
                        if (trimmed.empty())
                        {
                            if (dialogState->workingSettings.cache.has_value() && dialogState->workingSettings.cache->directoryInfo.mruWatched.has_value())
                            {
                                dialogState->workingSettings.cache->directoryInfo.mruWatched.reset();
                                MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                changed = true;
                            }
                            break;
                        }
                        const auto valueOpt = PrefsUi::TryParseUInt32(trimmed);
                        if (valueOpt.has_value())
                        {
                            if (auto* cache = EnsureWorkingCacheSettings(dialogState->workingSettings))
                            {
                                const uint32_t value = valueOpt.value();
                                if (! cache->directoryInfo.mruWatched.has_value() || cache->directoryInfo.mruWatched.value() != value)
                                {
                                    cache->directoryInfo.mruWatched = value;
                                    MaybeResetWorkingCacheSettingsIfEmpty(dialogState->workingSettings);
                                    changed = true;
                                }
                            }
                        }
                        break;
                    }
                    case IDC_PREFS_ADV_FILEOPS_MAX_DIAG_LOG_FILES_EDIT:
                    {
                        const std::wstring text{field ? field->GetText() : std::wstring_view{}};
                        const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
                        const Common::Settings::FileOperationsSettings defaults{};
                        uint32_t value = defaults.maxDiagnosticsLogFiles;
                        if (const auto valueOpt = PrefsUi::TryParseUInt32(trimmed); valueOpt.has_value() && valueOpt.value() > 0)
                        {
                            value = valueOpt.value();
                        }
                        if (! dialogState->workingSettings.fileOperations.has_value() && value == defaults.maxDiagnosticsLogFiles)
                        {
                            break;
                        }
                        if (auto* fileOperations = EnsureWorkingFileOperationsSettings(dialogState->workingSettings))
                        {
                            changed                                = (fileOperations->maxDiagnosticsLogFiles != value);
                            fileOperations->maxDiagnosticsLogFiles = value;
                            MaybeResetWorkingFileOperationsSettingsIfEmpty(dialogState->workingSettings);
                        }
                        break;
                    }
                    default: break;
                }

                if (changed)
                {
                    if (HWND dlg = GetParent(host))
                    {
                        SetDirty(dlg, *dialogState);
                    }
                }

                Refresh(host, *dialogState);
            });
            ++editIndex;
        }
    }

    {
        auto& card = dxState->page.settingsFileCard;
        card.card  = root->AddChild<CardPanel>();
        card.label = root->AddChild<Label>();
        card.label->SetFontRole(FontRole::Body);
        card.description = root->AddChild<Label>();
        card.description->SetFontRole(FontRole::Small);
        card.description->SetMultiline(true);
        card.link = root->AddChild<Button>();
        card.link->SetVariant(ButtonVariant::Hyperlink);
        card.link->SetOnClick([host = parent]() noexcept
        {
            if (! host || IsWindow(host) == FALSE)
            {
                return;
            }

            if (HWND dlg = GetParent(host))
            {
                if (PostMessageW(dlg, WM_COMMAND, MAKEWPARAM(IDC_PREFS_ADV_OPEN_SETTINGS_FILE, 0), 0) == FALSE)
                {
                    static_cast<void>(Debug::ErrorWithLastError(L"Preferences.Advanced: PostMessageW failed for settings-file link"));
                }
            }
        });
        card.label->SetMnemonicTarget(card.link);
    }

    {
        AdvancedDxPage& page = dxState->page;
        std::vector<RedSalamander::DxUi::Control*> orderedChildren;
        orderedChildren.reserve(_pageContentRoot->DebugChildCount());

        const auto appendHeader = [&](const size_t index) noexcept { orderedChildren.push_back(page.headers[index]); };

        const auto appendToggleCard = [&](const size_t index) noexcept
        {
            orderedChildren.push_back(page.toggleCards[index].card);
            orderedChildren.push_back(page.toggleCards[index].label);
            orderedChildren.push_back(page.toggleCards[index].description);
            orderedChildren.push_back(page.toggleCards[index].toggle);
        };

        const auto appendInputCard = [&](const size_t index) noexcept
        {
            orderedChildren.push_back(page.inputCards[index].card);
            orderedChildren.push_back(page.inputCards[index].label);
            orderedChildren.push_back(page.inputCards[index].description);
            orderedChildren.push_back(page.inputCards[index].edit);
        };

        appendHeader(0u);
        appendToggleCard(0u);
        appendToggleCard(1u);
        appendInputCard(0u);

        appendHeader(1u);
        appendInputCard(1u);
        appendInputCard(2u);
        appendInputCard(3u);

        appendHeader(2u);
        appendInputCard(4u);
        appendToggleCard(2u);
        appendToggleCard(3u);
        orderedChildren.push_back(page.settingsFileCard.card);
        orderedChildren.push_back(page.settingsFileCard.label);
        orderedChildren.push_back(page.settingsFileCard.description);
        orderedChildren.push_back(page.settingsFileCard.link);

        PrefsReorderPanelChildren(_pageContentRoot, orderedChildren);
    }

    _dxState = std::move(dxState);
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    return true;
}

void AdvancedPane::DetachDxHosts() noexcept
{
    if (_pageContentRoot && _pageHostDx && _pageHost && IsWindow(_pageHost) != FALSE)
    {
        _pageHostDx->ResetInteractionState();
        _pageContentRoot->ClearChildren();
    }
    _pageHostDx      = nullptr;
    _pageContentRoot = nullptr;

    if (_dxState)
    {
        _dxState->Detach();
        _dxState.reset();
    }
}

void AdvancedPane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return;
    }

    const ThemePalette palette = MakeAppThemeDxPalette(state.theme);
    _pageHostDx->SetTheme(palette);
}

void AdvancedPane::SyncDxControlsFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState)
    {
        return;
    }

    const ThemePalette palette = MakeAppThemeDxPalette(state.theme);
    const auto& connections    = GetConnectionsSettingsOrDefault(state.workingSettings);
    const auto& cache          = GetCacheSettingsOrDefault(state.workingSettings);
    const auto& fileOperations = GetFileOperationsSettingsOrDefault(state.workingSettings);

    for (size_t i = 0; i < kAdvancedHeaderStringIds.size(); ++i)
    {
        if (_dxState->page.headers[i])
        {
            _dxState->page.headers[i]->SetText(LoadStringResource(nullptr, kAdvancedHeaderStringIds[i]));
            _dxState->page.headers[i]->SetTextColor(std::nullopt);
        }
    }

    const std::array<bool, kAdvancedToggleLabelStringIds.size()> toggleChecked = {{
        connections.bypassWindowsHello,
        connections.allowInsecureTlsInAutomation,
        fileOperations.diagnosticsInfoEnabled,
        fileOperations.diagnosticsDebugEnabled,
    }};

    const std::array<bool, kAdvancedToggleLabelStringIds.size()> toggleEnabled = {{
        true,
        true,
        true,
        true,
    }};

    for (size_t i = 0; i < kAdvancedToggleLabelStringIds.size(); ++i)
    {
        auto& toggleCard = _dxState->page.toggleCards[i];

        const bool enabled          = toggleEnabled[i];
        const auto labelColor       = enabled ? std::optional<D2D1_COLOR_F>{} : std::optional<D2D1_COLOR_F>(palette.disabledText);
        const auto descriptionColor = enabled ? std::optional<D2D1_COLOR_F>(palette.subduedText) : std::optional<D2D1_COLOR_F>(palette.disabledText);

        if (toggleCard.label)
        {
            toggleCard.label->SetText(LoadStringResource(nullptr, kAdvancedToggleLabelStringIds[i]));
            toggleCard.label->SetTextColor(labelColor);
        }
        if (toggleCard.description)
        {
            toggleCard.description->SetText(LoadStringResource(nullptr, kAdvancedToggleDescriptionStringIds[i]));
            toggleCard.description->SetTextColor(descriptionColor);
        }
        if (toggleCard.toggle)
        {
            toggleCard.toggle->SetEnabled(enabled);
            toggleCard.toggle->SetChecked(toggleChecked[i]);
        }
    }

    const std::array<bool, kAdvancedInputLabelStringIds.size()> inputEnabled = {{
        true,
        true,
        true,
        true,
        true,
    }};

    for (size_t i = 0; i < kAdvancedInputLabelStringIds.size(); ++i)
    {
        const bool enabled = inputEnabled[i];
        auto& inputCard    = _dxState->page.inputCards[i];
        if (inputCard.label)
        {
            inputCard.label->SetText(LoadStringResource(nullptr, kAdvancedInputLabelStringIds[i]));
            inputCard.label->SetTextColor(enabled ? std::optional<D2D1_COLOR_F>{} : std::optional<D2D1_COLOR_F>(palette.disabledText));
        }
        if (inputCard.description)
        {
            inputCard.description->SetText(LoadStringResource(nullptr, kAdvancedInputDescriptionStringIds[i]));
            inputCard.description->SetTextColor(enabled ? std::optional<D2D1_COLOR_F>(palette.subduedText) : std::optional<D2D1_COLOR_F>(palette.disabledText));
        }
    }

    const auto syncEditDirect = [](TextField* field, const std::wstring& text, bool enabled, bool& syncFlag) noexcept
    {
        if (! field)
        {
            return;
        }

        syncFlag = true;
        field->SetText(text);
        field->SetEnabled(enabled);
        syncFlag = false;
    };

    {
        const std::wstring helloTimeoutText = std::to_wstring(connections.windowsHelloReauthTimeoutMinute);
        syncEditDirect(_dxState->page.inputCards[0].edit, helloTimeoutText, true, _syncingDxHelloTimeoutEdit);
    }
    {
        std::wstring cacheBytesText;
        if (cache.directoryInfo.maxBytes.has_value() && cache.directoryInfo.maxBytes.value() > 0)
        {
            cacheBytesText = FormatCacheBytes(cache.directoryInfo.maxBytes.value());
        }
        syncEditDirect(_dxState->page.inputCards[1].edit, cacheBytesText, true, _syncingDxCacheDirectoryInfoMaxBytesEdit);
    }
    {
        std::wstring maxWatchersText;
        if (cache.directoryInfo.maxWatchers.has_value())
        {
            maxWatchersText = std::to_wstring(cache.directoryInfo.maxWatchers.value());
        }
        syncEditDirect(_dxState->page.inputCards[2].edit, maxWatchersText, true, _syncingDxCacheDirectoryInfoMaxWatchersEdit);
    }
    {
        std::wstring mruWatchedText;
        if (cache.directoryInfo.mruWatched.has_value())
        {
            mruWatchedText = std::to_wstring(cache.directoryInfo.mruWatched.value());
        }
        syncEditDirect(_dxState->page.inputCards[3].edit, mruWatchedText, true, _syncingDxCacheDirectoryInfoMruWatchedEdit);
    }
    {
        const std::wstring maxLogFilesText = std::to_wstring(fileOperations.maxDiagnosticsLogFiles);
        syncEditDirect(_dxState->page.inputCards[4].edit, maxLogFilesText, true, _syncingDxFileOperationsMaxDiagnosticsLogFilesEdit);
    }
    if (_dxState->page.settingsFileCard.label)
    {
        _dxState->page.settingsFileCard.label->SetText(LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_SETTINGS_FILE));
        _dxState->page.settingsFileCard.label->SetTextColor(std::nullopt);
    }
    if (_dxState->page.settingsFileCard.description)
    {
        _dxState->page.settingsFileCard.description->SetText(LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_SETTINGS_FILE));
        _dxState->page.settingsFileCard.description->SetTextColor(palette.subduedText);
    }
    if (_dxState->page.settingsFileCard.link)
    {
        _dxState->page.settingsFileCard.link->SetText(LoadStringResource(nullptr, IDS_PREFS_ADV_OPEN_SETTINGS_FILE_LINK));
        _dxState->page.settingsFileCard.link->SetEnabled(true);
    }
    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void AdvancedPane::LayoutDxHosts(const PreferencesDialogState&) noexcept
{
    if (_dxState && _pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void AdvancedPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost = parent;

    if (state.currentCategory != PrefCategory::Advanced)
    {
        return;
    }

    if (! EnsureDxHosts(parent, state))
    {
        Debug::Error(L"Preferences.Advanced: DxUi surface initialization failed; page will not render correctly.");
        DetachDxHosts();
        return;
    }

    Refresh(parent, state);
}

void AdvancedPane::LayoutDxPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(host);
    static_cast<void>(margin);

    Debug::Perf::Scope layoutPerf(L"preferences.ui.advanced_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(static_cast<uint64_t>(std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI)));

    const UINT dpi = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);

    const int rowHeight   = std::max(1, UiMetrics::ScaleDip(dpi, kRowHeightDip));
    const int titleHeight = std::max(1, UiMetrics::ScaleDip(dpi, kTitleHeightDip));

    const int cardPaddingX = UiMetrics::ScaleDip(dpi, kCardPaddingXDip);
    const int cardPaddingY = UiMetrics::ScaleDip(dpi, kCardPaddingYDip);
    const int cardGapY     = UiMetrics::ScaleDip(dpi, kCardGapYDip);
    const int cardGapX     = UiMetrics::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY = UiMetrics::ScaleDip(dpi, kCardSpacingYDip);

    if (! _dxState || ! _pageHostDx || ! _pageContentRoot)
    {
        return;
    }
    AdvancedDxPage& dxPage = _dxState->page;

    const auto pxToDip = [dpi](const int px) noexcept { return static_cast<float>(px) * 96.0f / static_cast<float>(std::max<UINT>(1u, dpi)); };

    const int minToggleWidth    = UiMetrics::ScaleDip(dpi, kMinToggleWidthDip);
    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

    const int onWidth  = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, onLabel);
    const int offWidth = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, offLabel);

    const int paddingX       = UiMetrics::ScaleDip(dpi, kTogglePaddingXDip);
    const int gapX           = UiMetrics::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth     = UiMetrics::ScaleDip(dpi, kToggleTrackWidthDip);
    const int stateTextWidth = std::max(onWidth, offWidth);

    const int measuredToggleWidth = std::max(minToggleWidth, (2 * paddingX) + stateTextWidth + gapX + trackWidth);
    const int toggleWidth =
        static_cast<int>(std::lround(RedSalamander::DxUi::ResolveConstrainedExtent({.minExtent       = static_cast<float>(minToggleWidth),
                                                                                    .preferredExtent = static_cast<float>(measuredToggleWidth),
                                                                                    .maxExtent       = static_cast<float>(measuredToggleWidth)},
                                                                                   static_cast<float>(std::max(0, width - 2 * cardPaddingX - cardGapX)))));

    auto layoutToggleCard =
        [&](std::wstring_view labelText, std::wstring_view descText, CardPanel* dxCard, Label* dxLabel, Label* dxDescription, Toggle* dxToggle) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        state.pageSettingCards.push_back(card);

        if (dxCard)
        {
            dxCard->SetBounds(D2D1::RectF(pxToDip(card.left), pxToDip(card.top), pxToDip(card.right), pxToDip(card.bottom)));
        }
        if (dxLabel)
        {
            dxLabel->SetText(std::wstring(labelText));
            dxLabel->SetMnemonicTarget(dxToggle);
            dxLabel->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                           pxToDip(card.top + cardPaddingY),
                                           pxToDip(card.left + cardPaddingX + textWidth),
                                           pxToDip(card.top + cardPaddingY + titleHeight)));
        }
        if (dxDescription)
        {
            dxDescription->SetText(std::wstring(descText));
            dxDescription->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                 pxToDip(card.top + cardPaddingY + titleHeight + cardGapY),
                                                 pxToDip(card.left + cardPaddingX + textWidth),
                                                 pxToDip(card.top + cardPaddingY + titleHeight + cardGapY + descHeight)));
        }
        if (dxToggle)
        {
            dxToggle->SetBounds(D2D1::RectF(pxToDip(card.right - cardPaddingX - toggleWidth),
                                            pxToDip(card.top + (cardHeight - rowHeight) / 2),
                                            pxToDip(card.right - cardPaddingX),
                                            pxToDip(card.top + (cardHeight - rowHeight) / 2 + rowHeight)));
        }

        y += cardHeight + cardSpacingY;
    };

    auto layoutEditCard = [&](std::wstring_view labelText,
                              int desiredWidth,
                              std::wstring_view descText,
                              CardPanel* dxCard,
                              Label* dxLabel,
                              Label* dxDescription,
                              TextField* dxEdit) noexcept
    {
        desiredWidth         = static_cast<int>(std::lround(RedSalamander::DxUi::ResolveConstrainedExtent(
            {.minExtent = 0.0f, .preferredExtent = static_cast<float>(desiredWidth), .maxExtent = static_cast<float>(desiredWidth)},
            static_cast<float>(std::max(0, width - 2 * cardPaddingX)))));
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - desiredWidth);
        const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        state.pageSettingCards.push_back(card);

        if (dxCard)
        {
            dxCard->SetBounds(D2D1::RectF(pxToDip(card.left), pxToDip(card.top), pxToDip(card.right), pxToDip(card.bottom)));
        }
        if (dxLabel)
        {
            dxLabel->SetText(std::wstring(labelText));
            dxLabel->SetMnemonicTarget(dxEdit);
            dxLabel->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                           pxToDip(card.top + cardPaddingY),
                                           pxToDip(card.left + cardPaddingX + textWidth),
                                           pxToDip(card.top + cardPaddingY + titleHeight)));
        }
        if (dxDescription)
        {
            dxDescription->SetText(std::wstring(descText));
            dxDescription->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                 pxToDip(card.top + cardPaddingY + titleHeight + cardGapY),
                                                 pxToDip(card.left + cardPaddingX + textWidth),
                                                 pxToDip(card.top + cardPaddingY + titleHeight + cardGapY + descHeight)));
        }
        if (dxEdit)
        {
            dxEdit->SetBounds(D2D1::RectF(pxToDip(card.right - cardPaddingX - desiredWidth),
                                          pxToDip(card.top + (cardHeight - rowHeight) / 2),
                                          pxToDip(card.right - cardPaddingX),
                                          pxToDip(card.top + (cardHeight - rowHeight) / 2 + rowHeight)));
        }

        y += cardHeight + cardSpacingY;
    };

    auto layoutLinkCard = [&](std::wstring_view labelText,
                              std::wstring_view descText,
                              std::wstring_view linkText,
                              CardPanel* dxCard,
                              Label* dxLabel,
                              Label* dxDescription,
                              Button* dxButton) noexcept
    {
        const int textWidth     = std::max(0, width - 2 * cardPaddingX);
        const int descHeight    = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);
        const int measuredWidth = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, linkText) + UiMetrics::ScaleDip(dpi, 24);
        const int linkWidth     = std::max(1, std::min(textWidth, measuredWidth));
        const int contentHeight = titleHeight + cardGapY + descHeight + cardGapY + rowHeight;
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);
        const int linkTop       = y + cardPaddingY + titleHeight + cardGapY + descHeight + cardGapY;

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        state.pageSettingCards.push_back(card);

        if (dxCard)
        {
            dxCard->SetBounds(D2D1::RectF(pxToDip(card.left), pxToDip(card.top), pxToDip(card.right), pxToDip(card.bottom)));
        }
        if (dxLabel)
        {
            dxLabel->SetText(std::wstring(labelText));
            dxLabel->SetMnemonicTarget(dxButton);
            dxLabel->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                           pxToDip(card.top + cardPaddingY),
                                           pxToDip(card.left + cardPaddingX + textWidth),
                                           pxToDip(card.top + cardPaddingY + titleHeight)));
        }
        if (dxDescription)
        {
            dxDescription->SetText(std::wstring(descText));
            dxDescription->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                 pxToDip(card.top + cardPaddingY + titleHeight + cardGapY),
                                                 pxToDip(card.left + cardPaddingX + textWidth),
                                                 pxToDip(card.top + cardPaddingY + titleHeight + cardGapY + descHeight)));
        }
        if (dxButton)
        {
            dxButton->SetText(std::wstring(linkText));
            dxButton->SetBounds(
                D2D1::RectF(pxToDip(card.left + cardPaddingX), pxToDip(linkTop), pxToDip(card.left + cardPaddingX + linkWidth), pxToDip(linkTop + rowHeight)));
        }

        y += cardHeight + cardSpacingY;
    };

    const std::wstring labelBypassHelloText                = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CONNECTIONS_BYPASS_HELLO);
    const std::wstring labelAllowInsecureTlsAutomationText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CONNECTIONS_ALLOW_INSECURE_TLS_AUTOMATION);
    const std::wstring labelHelloTimeoutText               = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CONNECTIONS_HELLO_TIMEOUT);

    const std::wstring descBypassHelloText                = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CONNECTIONS_BYPASS_HELLO);
    const std::wstring descAllowInsecureTlsAutomationText = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CONNECTIONS_ALLOW_INSECURE_TLS_AUTOMATION);
    const std::wstring descHelloTimeoutText               = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CONNECTIONS_HELLO_TIMEOUT);

    layoutToggleCard(labelBypassHelloText,
                     descBypassHelloText,
                     dxPage.toggleCards[0].card,
                     dxPage.toggleCards[0].label,
                     dxPage.toggleCards[0].description,
                     dxPage.toggleCards[0].toggle);
    layoutToggleCard(labelAllowInsecureTlsAutomationText,
                     descAllowInsecureTlsAutomationText,
                     dxPage.toggleCards[1].card,
                     dxPage.toggleCards[1].label,
                     dxPage.toggleCards[1].description,
                     dxPage.toggleCards[1].toggle);
    layoutEditCard(labelHelloTimeoutText,
                   UiMetrics::ScaleDip(dpi, kMinToggleWidthDip),
                   descHelloTimeoutText,
                   dxPage.inputCards[0].card,
                   dxPage.inputCards[0].label,
                   dxPage.inputCards[0].description,
                   dxPage.inputCards[0].edit);

    y += gapY;

    const std::wstring labelCacheMaxBytesText    = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CACHE_DIR_MAX_BYTES);
    const std::wstring labelCacheMaxWatchersText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CACHE_DIR_MAX_WATCHERS);
    const std::wstring labelCacheMruWatchedText  = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_CACHE_DIR_MRU_WATCHED);

    const std::wstring descCacheMaxBytesText    = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CACHE_DIR_MAX_BYTES);
    const std::wstring descCacheMaxWatchersText = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CACHE_DIR_MAX_WATCHERS);
    const std::wstring descCacheMruWatchedText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_CACHE_DIR_MRU_WATCHED);

    layoutEditCard(labelCacheMaxBytesText,
                   UiMetrics::ScaleDip(dpi, kMediumComboWidthDip),
                   descCacheMaxBytesText,
                   dxPage.inputCards[1].card,
                   dxPage.inputCards[1].label,
                   dxPage.inputCards[1].description,
                   dxPage.inputCards[1].edit);
    layoutEditCard(labelCacheMaxWatchersText,
                   UiMetrics::ScaleDip(dpi, kMinToggleWidthDip),
                   descCacheMaxWatchersText,
                   dxPage.inputCards[2].card,
                   dxPage.inputCards[2].label,
                   dxPage.inputCards[2].description,
                   dxPage.inputCards[2].edit);
    layoutEditCard(labelCacheMruWatchedText,
                   UiMetrics::ScaleDip(dpi, kMinToggleWidthDip),
                   descCacheMruWatchedText,
                   dxPage.inputCards[3].card,
                   dxPage.inputCards[3].label,
                   dxPage.inputCards[3].description,
                   dxPage.inputCards[3].edit);

    y += gapY;

    const std::wstring labelMaxDiagnosticsLogFilesText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILEOPS_MAX_DIAG_LOG_FILES);
    const std::wstring descMaxDiagnosticsLogFilesText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILEOPS_MAX_DIAG_LOG_FILES);

    layoutEditCard(labelMaxDiagnosticsLogFilesText,
                   UiMetrics::ScaleDip(dpi, kMinToggleWidthDip),
                   descMaxDiagnosticsLogFilesText,
                   dxPage.inputCards[4].card,
                   dxPage.inputCards[4].label,
                   dxPage.inputCards[4].description,
                   dxPage.inputCards[4].edit);

    const std::wstring labelDiagnosticsInfoText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILEOPS_DIAG_INFO);
    const std::wstring descDiagnosticsInfoText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILEOPS_DIAG_INFO);
    layoutToggleCard(labelDiagnosticsInfoText,
                     descDiagnosticsInfoText,
                     dxPage.toggleCards[2].card,
                     dxPage.toggleCards[2].label,
                     dxPage.toggleCards[2].description,
                     dxPage.toggleCards[2].toggle);

    const std::wstring labelDiagnosticsDebugText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILEOPS_DIAG_DEBUG);
    const std::wstring descDiagnosticsDebugText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILEOPS_DIAG_DEBUG);
    layoutToggleCard(labelDiagnosticsDebugText,
                     descDiagnosticsDebugText,
                     dxPage.toggleCards[3].card,
                     dxPage.toggleCards[3].label,
                     dxPage.toggleCards[3].description,
                     dxPage.toggleCards[3].toggle);

    y += gapY;

    const std::wstring labelSettingsFileText = LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_SETTINGS_FILE);
    const std::wstring descSettingsFileText  = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_SETTINGS_FILE);
    const std::wstring openSettingsFileText  = LoadStringResource(nullptr, IDS_PREFS_ADV_OPEN_SETTINGS_FILE_LINK);
    layoutLinkCard(labelSettingsFileText,
                   descSettingsFileText,
                   openSettingsFileText,
                   dxPage.settingsFileCard.card,
                   dxPage.settingsFileCard.label,
                   dxPage.settingsFileCard.description,
                   dxPage.settingsFileCard.link);

    SyncDxControlsFromState(state);
    _pageHostDx->Invalidate();
}

void AdvancedPane::LayoutPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    if (! host)
    {
        return;
    }

    if (EnsureDxHosts(_pageHost ? _pageHost : host, state))
    {
        LayoutDxPage(host, state, x, y, width, margin, gapY, typography);
        return;
    }

    Debug::Error(L"Preferences.Advanced: DxUi surface initialization failed; page will not render correctly.");
}

void AdvancedPane::Refresh(HWND /*host*/, PreferencesDialogState& state) noexcept
{
    if (state.currentCategory == PrefCategory::Advanced && ! _dxState)
    {
        static_cast<void>(EnsureDxHosts(_pageHost, state));
    }

    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
}

#ifdef ENABLE_TESTS
PreferencesAdvancedDebugFocusTarget AdvancedPane::DebugGetFocusTarget() const noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return PreferencesAdvancedDebugFocusTarget::None;
    }

    const auto* focused = _pageHostDx->GetFocusControl();
    if (! focused)
    {
        return PreferencesAdvancedDebugFocusTarget::None;
    }

    const auto& page = _dxState->page;
    if (page.toggleCards[0].toggle == focused)
    {
        return PreferencesAdvancedDebugFocusTarget::BypassHelloToggle;
    }
    if (page.toggleCards[1].toggle == focused)
    {
        return PreferencesAdvancedDebugFocusTarget::AllowInsecureTlsAutomationToggle;
    }
    if (page.inputCards[0].edit == focused)
    {
        return PreferencesAdvancedDebugFocusTarget::HelloTimeoutEdit;
    }
    if (page.inputCards[1].edit == focused)
    {
        return PreferencesAdvancedDebugFocusTarget::CacheMaxBytesEdit;
    }
    if (page.inputCards[2].edit == focused)
    {
        return PreferencesAdvancedDebugFocusTarget::CacheMaxWatchersEdit;
    }
    if (page.inputCards[3].edit == focused)
    {
        return PreferencesAdvancedDebugFocusTarget::CacheMruWatchedEdit;
    }
    if (page.inputCards[4].edit == focused)
    {
        return PreferencesAdvancedDebugFocusTarget::FileOpsMaxDiagnosticsLogFilesEdit;
    }
    if (page.toggleCards[2].toggle == focused)
    {
        return PreferencesAdvancedDebugFocusTarget::DiagnosticsInfoToggle;
    }
    if (page.toggleCards[3].toggle == focused)
    {
        return PreferencesAdvancedDebugFocusTarget::DiagnosticsDebugToggle;
    }
    if (page.settingsFileCard.link == focused)
    {
        return PreferencesAdvancedDebugFocusTarget::OpenSettingsFileLink;
    }

    return PreferencesAdvancedDebugFocusTarget::None;
}

bool AdvancedPane::DebugFocusBypassHelloToggle() noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return false;
    }

    auto* const toggle = _dxState->page.toggleCards[0].toggle;
    if (! toggle || ! toggle->IsVisible() || ! toggle->IsEnabled())
    {
        return false;
    }

    _pageHostDx->SetFocusControl(toggle);
    return _pageHostDx->GetFocusControl() == toggle;
}

#endif
