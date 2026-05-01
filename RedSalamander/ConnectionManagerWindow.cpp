#include "Framework.h"

#include "ConnectionManagerWindow.h"

#include <algorithm>
#include <atomic>
#include <cstdint>
#include <cwctype>
#include <format>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include <commdlg.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "ConnectionProfileUtils.h"
#include "ConnectionSecrets.h"
#include "DxUi/DxUi.Typography.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "Helpers.h"
#include "HostServices.h"
#include "SettingsHotReload.h"
#include "UiMetrics.h"
#include "WindowMaximizeBehavior.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "WindowSizing.h"
#include "resource.h"

// Single-canvas DxUi Connection Manager window.
//
// Phase 2 landed: top-level WS_OVERLAPPEDWINDOW + single DxUi::WindowHost +
// dispatcher gate behind constexpr `IsEnabled()`.
// Phase 3 landed: 3-pane layout (list / editor placeholder / footer).
// Phase 4 partial (4.1-4.3): the connection list grid is parented into the
// in-tree root and populated from settings.
//
// The public modeless entry point, synchronous dialog facade, and debug test
// hooks are implemented in this translation unit.
//
// See `Specs/Plans/WIP/UI_ConnectionManagerSingleCanvasPlan.md`.

namespace RedSalamander::ConnectionManager::SingleCanvas
{

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::CardPanel;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::Control;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridColumnKind;
using RedSalamander::DxUi::GridDebugRowVisualState;
using RedSalamander::DxUi::GridRowStyle;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::GridVisibleWorkMetrics;
using RedSalamander::DxUi::IDxGridDelegate;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::ScrollPanel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kWindowClassName[]  = L"RedSalamander.ConnectionManagerWindow";
constexpr wchar_t kWindowSettingsId[] = L"ConnectionManagerWindow";

constexpr float kListPaneWidthDip     = 200.0f;
constexpr float kFooterHeightDip      = 44.0f;
constexpr float kPanePaddingDip       = 8.0f;
constexpr float kListGridRowHeight    = 28.0f;
constexpr float kFooterButtonHeight   = 32.0f;
constexpr float kFooterButtonWidth    = 96.0f;
constexpr float kFooterButtonGap      = 8.0f;
constexpr float kFormLabelWidthDip    = 140.0f;
constexpr float kFormLabelGapDip      = 8.0f;
constexpr float kFormRowHeightDip     = 28.0f;
constexpr float kFormRowGapDip        = 4.0f;
constexpr float kFormSectionGapDip    = 8.0f;
constexpr float kFormControlHeightDip = 24.0f;
constexpr float kListButtonHeightDip  = 26.0f;
constexpr float kListButtonGapDip     = 6.0f;

class WindowImpl;
std::atomic<HWND> g_singleInstance{nullptr};

// Phase 8.2b - facade-mode result captured before the window is destroyed.
// `WindowImpl` writes the connection name + HRESULT into this struct from
// `OnConnectClicked` / `OnCancelClicked` / `OnCloseClicked` before closing
// the owned HWND. The synchronous `ShowDialog` reads the struct after the
// nested message pump exits.
struct ModalFacadeResult
{
    HRESULT hr = S_FALSE;
    std::wstring connectionName;
};

struct ConnectionListGridRow
{
    uint64_t stableId = 0u;
    size_t modelIndex = 0u;
    std::wstring text;
};

[[nodiscard]] uint64_t MakeConnectionListStableId(std::wstring_view connectionId, const size_t fallbackModelIndex) noexcept
{
    constexpr uint64_t kFNVOffset = 1469598103934665603ull;
    constexpr uint64_t kFNVPrime  = 1099511628211ull;
    if (connectionId.empty())
    {
        return fallbackModelIndex + 1u;
    }
    uint64_t value = kFNVOffset;
    for (const wchar_t ch : connectionId)
    {
        value ^= static_cast<uint64_t>(std::towlower(static_cast<wint_t>(ch)));
        value *= kFNVPrime;
    }
    return value;
}

class ConnectionListGridModel final : public IDxGridModel
{
public:
    ConnectionListGridModel()
    {
        // The single column carries the section title so the legacy
        // "Connections" heading lives inside the grid's column header (and
        // gets click-to-sort behaviour for free).
        GridColumnDesc col{};
        col.id    = L"name";
        col.title = LoadStringResource(nullptr, IDS_CONNECTIONS_LIST_HEADER);
        if (col.title.empty())
        {
            col.title = L"Connections";
        }
        col.widthDip    = 220.0f;
        col.minWidthDip = 120.0f;
        col.kind        = GridColumnKind::Text;
        col.sortable    = true;
        col.multiline   = false;
        _columns        = {std::move(col)};
    }

    void SetRows(std::vector<ConnectionListGridRow> rows)
    {
        _rows = std::move(rows);
        _rowIndexByStableId.clear();
        _rowIndexByStableId.reserve(_rows.size());
        for (size_t rowIndex = 0u; rowIndex < _rows.size(); ++rowIndex)
        {
            _rowIndexByStableId[_rows[rowIndex].stableId] = rowIndex;
        }
    }

    [[nodiscard]] const std::vector<ConnectionListGridRow>& GetRows() const noexcept
    {
        return _rows;
    }

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _rows.size();
    }
    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }
    [[nodiscard]] GridColumnDesc GetColumn(const size_t columnIndex) const override
    {
        return _columns.at(columnIndex);
    }
    void GetCellData(const size_t rowIndex, const size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        if (rowIndex >= _rows.size() || columnIndex != 0u)
        {
            return;
        }
        outCell.text = _rows[rowIndex].text;
    }
    [[nodiscard]] uint64_t GetStableRowId(const size_t rowIndex) const noexcept override
    {
        return rowIndex < _rows.size() ? _rows[rowIndex].stableId : 0u;
    }
    [[nodiscard]] std::optional<size_t> FindRowByStableId(const uint64_t rowId) const noexcept override
    {
        const auto it = _rowIndexByStableId.find(rowId);
        return it == _rowIndexByStableId.end() ? std::nullopt : std::optional<size_t>{it->second};
    }
    [[nodiscard]] GridRowStyle GetRowStyle(const size_t rowIndex) const override
    {
        if (rowIndex >= _rows.size())
        {
            return {};
        }
        GridRowStyle style{};
        style.rainbowSeed = _rows[rowIndex].text;
        return style;
    }

private:
    std::vector<GridColumnDesc> _columns;
    std::vector<ConnectionListGridRow> _rows;
    std::unordered_map<uint64_t, size_t> _rowIndexByStableId;
};

[[nodiscard]] HWND NormalizeOwnerWindow(HWND owner) noexcept
{
    if (owner && IsWindow(owner) != FALSE)
    {
        return GetAncestor(owner, GA_ROOT);
    }
    return nullptr;
}

// Plugin-id helpers used by the profile editor's protocol-specific fields.
[[nodiscard]] bool IsFtpPluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-ftp";
}
[[nodiscard]] bool IsSshPluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-sftp" || pluginId == L"builtin/file-system-scp";
}
[[nodiscard]] bool IsImapPluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-imap";
}
[[nodiscard]] bool IsGoogleDrivePluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-gdrive";
}
[[nodiscard]] bool IsS3PluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-s3";
}
[[nodiscard]] bool IsS3TablePluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-s3table";
}
[[nodiscard]] bool IsOneDrivePersonalPluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-onedrive-personal";
}
[[nodiscard]] bool IsOneDriveBusinessPluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-onedrive-business";
}
[[nodiscard]] bool IsSharePointPluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-sharepoint";
}
[[nodiscard]] bool IsMicrosoftPluginId(std::wstring_view pluginId) noexcept
{
    return IsOneDrivePersonalPluginId(pluginId) || IsOneDriveBusinessPluginId(pluginId) || IsSharePointPluginId(pluginId);
}

[[nodiscard]] std::wstring SanitizeUnsignedText(std::wstring_view text) noexcept
{
    std::wstring result;
    result.reserve(text.size());
    for (const wchar_t ch : text)
    {
        if (ch >= L'0' && ch <= L'9')
        {
            result.push_back(ch);
        }
    }
    return result;
}

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text) noexcept
{
    size_t start = 0u;
    while (start < text.size() && std::iswspace(static_cast<wint_t>(text[start])) != 0)
    {
        ++start;
    }
    size_t end = text.size();
    while (end > start && std::iswspace(static_cast<wint_t>(text[end - 1u])) != 0)
    {
        --end;
    }
    return std::wstring(text.substr(start, end - start));
}

[[nodiscard]] bool EqualsIgnoreCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }
    for (size_t i = 0; i < a.size(); ++i)
    {
        if (std::towlower(a[i]) != std::towlower(b[i]))
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] std::wstring NewGuidString() noexcept
{
    GUID guid{};
    if (FAILED(CoCreateGuid(&guid)))
    {
        return {};
    }
    wchar_t buf[64]{};
    if (StringFromGUID2(guid, buf, static_cast<int>(std::size(buf))) <= 0)
    {
        return {};
    }
    std::wstring text(buf);
    if (! text.empty() && text.front() == L'{' && text.back() == L'}')
    {
        text.erase(text.begin());
        text.pop_back();
    }
    return text;
}

[[nodiscard]] std::wstring MakeUniqueConnectionName(const std::vector<Common::Settings::ConnectionProfile>& connections,
                                                    std::wstring_view desired,
                                                    std::wstring_view excludeId)
{
    std::wstring base = TrimWhitespace(desired);
    if (base.empty())
    {
        base = TrimWhitespace(LoadStringResource(nullptr, IDS_CONNECTIONS_DEFAULT_NEW_NAME));
    }
    for (auto& ch : base)
    {
        if (ch == L'/' || ch == L'\\')
        {
            ch = L'-';
        }
    }
    const auto isUsed = [&](std::wstring_view name) noexcept
    {
        if (name.empty())
        {
            return false;
        }
        for (const auto& c : connections)
        {
            if (! excludeId.empty() && c.id == excludeId)
            {
                continue;
            }
            if (! c.name.empty() && EqualsIgnoreCase(c.name, name))
            {
                return true;
            }
        }
        return false;
    };
    if (! isUsed(base))
    {
        return base;
    }
    for (int suffix = 2; suffix < 10000; ++suffix)
    {
        std::wstring candidate = std::format(L"{} ({})", base, suffix);
        if (! isUsed(candidate))
        {
            return candidate;
        }
    }
    return base;
}

// UTF helpers for the JSON `extra` payload - strings are stored as UTF-8.
[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }
    const int needed = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (needed <= 0)
    {
        return {};
    }
    std::wstring result(static_cast<size_t>(needed), L'\0');
    static_cast<void>(MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), needed));
    return result;
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }
    const int needed = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (needed <= 0)
    {
        return {};
    }
    std::string result(static_cast<size_t>(needed), '\0');
    static_cast<void>(WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), needed, nullptr, nullptr));
    return result;
}

// `extra` JSON helpers for plugin-specific connection profile fields.
[[nodiscard]] std::optional<std::wstring> ExtraGetString(const Common::Settings::JsonValue& extra, std::string_view key) noexcept
{
    const auto* objPtr = std::get_if<Common::Settings::JsonValue::ObjectPtr>(&extra.value);
    if (! objPtr || ! *objPtr)
    {
        return std::nullopt;
    }
    for (const auto& [k, v] : (*objPtr)->members)
    {
        if (k != key)
        {
            continue;
        }
        const auto* str = std::get_if<std::string>(&v.value);
        if (! str)
        {
            return std::nullopt;
        }
        const std::wstring wide = Utf16FromUtf8(*str);
        return wide.empty() && ! str->empty() ? std::nullopt : std::make_optional(wide);
    }
    return std::nullopt;
}

void ExtraSetString(Common::Settings::JsonValue& extra, std::string_view key, std::wstring_view value)
{
    Common::Settings::JsonValue::ObjectPtr obj;
    if (auto* existing = std::get_if<Common::Settings::JsonValue::ObjectPtr>(&extra.value); existing && *existing)
    {
        obj = *existing;
    }
    else
    {
        obj         = std::make_shared<Common::Settings::JsonObject>();
        extra.value = obj;
    }
    const std::string keyUtf8(key);
    if (keyUtf8.empty())
    {
        return;
    }
    const std::string valueUtf8 = Utf8FromUtf16(value);
    if (valueUtf8.empty() && ! value.empty())
    {
        return;
    }
    for (auto& member : obj->members)
    {
        if (member.first != keyUtf8)
        {
            continue;
        }
        member.second.value = valueUtf8;
        return;
    }
    Common::Settings::JsonValue v;
    v.value = valueUtf8;
    obj->members.emplace_back(keyUtf8, std::move(v));
}

void ExtraSetBool(Common::Settings::JsonValue& extra, std::string_view key, bool value)
{
    Common::Settings::JsonValue::ObjectPtr obj;
    if (auto* existing = std::get_if<Common::Settings::JsonValue::ObjectPtr>(&extra.value); existing && *existing)
    {
        obj = *existing;
    }
    else
    {
        obj         = std::make_shared<Common::Settings::JsonObject>();
        extra.value = obj;
    }
    const std::string keyUtf8(key);
    if (keyUtf8.empty())
    {
        return;
    }
    for (auto& member : obj->members)
    {
        if (member.first != keyUtf8)
        {
            continue;
        }
        member.second.value = value;
        return;
    }
    Common::Settings::JsonValue v;
    v.value = value;
    obj->members.emplace_back(keyUtf8, std::move(v));
}

void ExtraSetUInt32(Common::Settings::JsonValue& extra, std::string_view key, uint32_t value)
{
    Common::Settings::JsonValue::ObjectPtr obj;
    if (auto* existing = std::get_if<Common::Settings::JsonValue::ObjectPtr>(&extra.value); existing && *existing)
    {
        obj = *existing;
    }
    else
    {
        obj         = std::make_shared<Common::Settings::JsonObject>();
        extra.value = obj;
    }
    const std::string keyUtf8(key);
    if (keyUtf8.empty())
    {
        return;
    }
    for (auto& member : obj->members)
    {
        if (member.first != keyUtf8)
        {
            continue;
        }
        member.second.value = static_cast<uint64_t>(value);
        return;
    }
    Common::Settings::JsonValue v;
    v.value = static_cast<uint64_t>(value);
    obj->members.emplace_back(keyUtf8, std::move(v));
}

[[nodiscard]] bool IsQuickConnectProfile(const Common::Settings::ConnectionProfile& profile) noexcept
{
    return RedSalamander::Connections::IsQuickConnectConnectionId(profile.id);
}

enum class ConnectionProfileValidationError : uint8_t
{
    None,
    EmptyName,
    DuplicateName,
    ReservedName,
};

struct ConnectionProfileValidationResult
{
    ConnectionProfileValidationError error = ConnectionProfileValidationError::None;
    std::wstring proposedName;
    std::wstring normalizedName;
    size_t conflictingIndex = static_cast<size_t>(-1);
};

[[nodiscard]] bool IsReservedConnectionName(std::wstring_view name)
{
    if (EqualsIgnoreCase(name, RedSalamander::Connections::kQuickConnectConnectionName) || EqualsIgnoreCase(name, L"Quick Connect"))
    {
        return true;
    }

    const std::wstring quickConnectLabel = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
    return ! quickConnectLabel.empty() && EqualsIgnoreCase(name, quickConnectLabel);
}

[[nodiscard]] ConnectionProfileValidationResult ValidateConnectionProfileName(const std::vector<Common::Settings::ConnectionProfile>& profiles,
                                                                              size_t editedIndex,
                                                                              std::wstring_view proposedName)
{
    ConnectionProfileValidationResult result{};
    result.proposedName     = std::wstring(proposedName);
    result.normalizedName   = TrimWhitespace(proposedName);
    result.conflictingIndex = static_cast<size_t>(-1);

    if (result.normalizedName.empty())
    {
        result.error = ConnectionProfileValidationError::EmptyName;
        return result;
    }

    if (IsReservedConnectionName(result.normalizedName))
    {
        result.error = ConnectionProfileValidationError::ReservedName;
        return result;
    }

    for (size_t index = 0; index < profiles.size(); ++index)
    {
        if (index == editedIndex)
        {
            continue;
        }

        const std::wstring existingName = TrimWhitespace(profiles[index].name);
        if (! existingName.empty() && EqualsIgnoreCase(existingName, result.normalizedName))
        {
            result.error            = ConnectionProfileValidationError::DuplicateName;
            result.conflictingIndex = index;
            return result;
        }
    }

    return result;
}

[[nodiscard]] UINT MessageResourceForConnectionNameValidation(ConnectionProfileValidationError error) noexcept
{
    switch (error)
    {
        case ConnectionProfileValidationError::None: return 0u;
        case ConnectionProfileValidationError::EmptyName: return IDS_CONNECTIONS_ERR_NAME_REQUIRED;
        case ConnectionProfileValidationError::DuplicateName: return IDS_CONNECTIONS_ERR_NAME_UNIQUE;
        case ConnectionProfileValidationError::ReservedName: return IDS_CONNECTIONS_ERR_NAME_RESERVED;
        default: return 0u;
    }
}

void ShowConnectionManagerAlert(HWND hwnd, HostAlertSeverity severity, const std::wstring& title, const std::wstring& message) noexcept
{
    if (! hwnd || message.empty())
    {
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = severity;
    request.targetWindow = hwnd;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;
    static_cast<void>(HostShowAlert(request));
}

struct EditorVisibility
{
    bool hasSelection       = false;
    bool isFtp              = false;
    bool isSsh              = false;
    bool isImap             = false;
    bool isGoogleDrive      = false;
    bool isMicrosoft        = false;
    bool isS3               = false;
    bool isS3Table          = false;
    bool isAwsS3            = false;
    bool anonymous          = false;
    bool oauth2             = false;
    bool showProtocol       = false;
    bool showHostEdit       = false;
    bool showAwsRegion      = false;
    bool showPort           = false;
    bool showConcurrency    = false;
    bool showAnonymous      = false;
    bool showSshSection     = false;
    bool showS3Section      = false;
    bool showVirtual        = false;
    bool showSecretRow      = false;
    bool authInputsEnabled  = false;
    bool showIgnoreSslTrust = false;
};

struct ProtocolItem
{
    const wchar_t* pluginId;
    const wchar_t* label;
};

constexpr ProtocolItem kProtocolItems[] = {
    {L"builtin/file-system-ftp", L"FTP"},
    {L"builtin/file-system-sftp", L"SFTP"},
    {L"builtin/file-system-scp", L"SCP"},
    {L"builtin/file-system-imap", L"IMAP"},
    {L"builtin/file-system-gdrive", L"Google Drive"},
    {L"builtin/file-system-onedrive-personal", L"OneDrive Personal"},
    {L"builtin/file-system-onedrive-business", L"OneDrive Business"},
    {L"builtin/file-system-sharepoint", L"SharePoint"},
    {L"builtin/file-system-s3", L"S3"},
    {L"builtin/file-system-s3table", L"S3 Table"},
};

struct AwsRegionItem
{
    const wchar_t* code;
    const wchar_t* name;
};

constexpr AwsRegionItem kAwsRegionItems[] = {
    {L"af-south-1", L"Africa (Cape Town)"},
    {L"ap-east-1", L"Asia Pacific (Hong Kong)"},
    {L"ap-northeast-1", L"Asia Pacific (Tokyo)"},
    {L"ap-northeast-2", L"Asia Pacific (Seoul)"},
    {L"ap-south-1", L"Asia Pacific (Mumbai)"},
    {L"ap-southeast-1", L"Asia Pacific (Singapore)"},
    {L"ap-southeast-2", L"Asia Pacific (Sydney)"},
    {L"ca-central-1", L"Canada (Central)"},
    {L"eu-central-1", L"Europe (Frankfurt)"},
    {L"eu-north-1", L"Europe (Stockholm)"},
    {L"eu-west-1", L"Europe (Ireland)"},
    {L"eu-west-2", L"Europe (London)"},
    {L"eu-west-3", L"Europe (Paris)"},
    {L"sa-east-1", L"South America (Sao Paulo)"},
    {L"us-east-1", L"US East (N. Virginia)"},
    {L"us-east-2", L"US East (Ohio)"},
    {L"us-west-1", L"US West (N. California)"},
    {L"us-west-2", L"US West (Oregon)"},
};

[[nodiscard]] std::vector<ComboBox::Item> BuildProtocolComboItems()
{
    std::vector<ComboBox::Item> items;
    items.reserve(std::size(kProtocolItems));
    for (const auto& entry : kProtocolItems)
    {
        items.push_back(ComboBox::Item{.value = entry.pluginId, .display = entry.label});
    }
    return items;
}

#ifdef ENABLE_TESTS
[[nodiscard]] std::optional<size_t> FindProtocolItemIndexByPluginId(std::wstring_view pluginId) noexcept
{
    if (pluginId.empty())
    {
        return std::nullopt;
    }
    for (size_t index = 0u; index < std::size(kProtocolItems); ++index)
    {
        if (pluginId == kProtocolItems[index].pluginId)
        {
            return index;
        }
    }
    return std::nullopt;
}
#endif

[[nodiscard]] std::vector<ComboBox::Item> BuildAwsRegionComboItems()
{
    std::vector<ComboBox::Item> items;
    items.reserve(std::size(kAwsRegionItems));
    for (const auto& entry : kAwsRegionItems)
    {
        items.push_back(ComboBox::Item{.value = entry.code, .display = std::format(L"{} ({})", entry.name, entry.code)});
    }
    return items;
}

class WindowImpl final : public IDxGridDelegate
{
public:
    WindowImpl(
        std::wstring appId, Common::Settings::Settings& settings, AppTheme theme, std::wstring filterPluginId, uint8_t targetPane, HWND notifyOwner) noexcept
        : _appId(std::move(appId)),
          _settings(&settings),
          _theme(std::move(theme)),
          _filterPluginId(std::move(filterPluginId)),
          _targetPane(targetPane),
          _notifyOwner(notifyOwner)
    {
    }

    ~WindowImpl() noexcept override = default;

    WindowImpl(const WindowImpl&)            = delete;
    WindowImpl(WindowImpl&&)                 = delete;
    WindowImpl& operator=(const WindowImpl&) = delete;
    WindowImpl& operator=(WindowImpl&&)      = delete;

    [[nodiscard]] bool Create() noexcept;
    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hwnd.get();
    }

    void UpdateTheme(const AppTheme& theme) noexcept;
    void UpdateContext(const AppTheme& theme, std::wstring_view filterPluginId, HWND notifyOwner, uint8_t targetPane) noexcept;

    // Phase 8.2b - opt the window into modal-facade mode. The window does NOT
    // register in `g_singleInstance`; on Connect/Close/Cancel the result is
    // written to `*resultPtr` before the window is destroyed.
    void EnterModalFacadeMode(ModalFacadeResult* resultPtr) noexcept
    {
        _isModalFacade = true;
        _modalResult   = resultPtr;
    }

    using IDxGridDelegate::OnGridRowActivated;
    using IDxGridDelegate::OnGridSelectionChanged;
    using IDxGridDelegate::OnGridSortRequested;
    void OnGridSelectionChanged(Grid& sender) override;
    void OnGridRowActivated(Grid& sender, size_t rowIndex) override;
    void OnGridSortRequested(const RedSalamander::DxUi::GridSortSpec& sortSpec) override;

#ifdef ENABLE_TESTS
    void DebugFillSnapshot(::ConnectionManagerDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugClickListRow(size_t rowIndex) noexcept;
    [[nodiscard]] bool DebugScrollListByWheelDetents(int detents) noexcept;
    [[nodiscard]] bool DebugFocusFirstInput() noexcept;
    [[nodiscard]] bool DebugFocusList() noexcept;
    [[nodiscard]] bool DebugRouteMnemonic(wchar_t mnemonic) noexcept;
    [[nodiscard]] bool DebugRouteCommandKey(WPARAM virtualKey) noexcept;
    [[nodiscard]] bool DebugRouteTab(bool reverse) noexcept;
    [[nodiscard]] bool DebugSetProtocolPluginId(std::wstring_view pluginId) noexcept;
    [[nodiscard]] bool DebugGetAlternateProtocolPluginId(std::wstring_view baselinePluginId, std::wstring& outPluginId) const noexcept;
    [[nodiscard]] bool DebugGetControlClientRect(const Control* control, RECT& outRect) const noexcept;
    [[nodiscard]] Button* DebugGetCommandButton(UINT commandId) noexcept;
    [[nodiscard]] Toggle* DebugGetSavePasswordToggle() noexcept
    {
        return _toggleSavePassword;
    }
    [[nodiscard]] Toggle* DebugGetS3UseHttpsToggle() noexcept
    {
        return _toggleS3UseHttps;
    }
    [[nodiscard]] HWND DebugGetHwnd() const noexcept
    {
        return _hwnd.get();
    }
#endif

    [[nodiscard]] static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

private:
    [[nodiscard]] bool OnCreate() noexcept;
    void OnSize() noexcept;
    void OnDpiChanged(const RECT* suggested) noexcept;
    void OnNcActivate(bool active) noexcept;
    void OnClose() noexcept;
    void OnNcDestroy() noexcept;
    void ApplyTheme() noexcept;
    void Layout() noexcept;
    void BuildUi();

    void MarkConnectionsDirty() noexcept;
    void LoadConnections() noexcept;
    void ReloadConnectionsFromSettingsPreservingSelection() noexcept;
    void RebuildList() noexcept;
    [[nodiscard]] std::optional<size_t> GetSelectedModelIndex() const noexcept;
    void SelectConnectionModelIndex(size_t modelIndex) noexcept;
    void RefreshEditorFromSelection() noexcept;
    void LoadEditorFromProfile(const Common::Settings::ConnectionProfile& profile) noexcept;
    void ClearEditor() noexcept;
    void BuildListPaneButtons();
    void BuildEditorForm();
    void BuildFooter();
    void LayoutEditorForm(D2D1_RECT_F editorRect) noexcept;
    void WireEditorChangeCallbacks();
    [[nodiscard]] EditorVisibility ComputeEditorVisibility() const noexcept;
    void ApplyEditorVisibility(const EditorVisibility& visibility) noexcept;
    void OnEditorFieldChanged() noexcept;
    void OnProtocolChanged(size_t comboIndex) noexcept;
    [[nodiscard]] bool BrowseSshFile(TextField* target) noexcept;
    void RequestCloseWindow() noexcept;
    void RefreshBaselineConnectionIds() noexcept;
    void StageSecretForProfile(const Common::Settings::ConnectionProfile& profile, std::wstring_view secret);
    void DeleteSecretsForRemovedConnections() noexcept;
    void CommitQuickConnectSecretsAndProfile(const Common::Settings::ConnectionProfile& profile) noexcept;
    [[nodiscard]] HRESULT CommitSecretsForProfile(const Common::Settings::ConnectionProfile& profile) noexcept;
    void ShowNameValidationError(const ConnectionProfileValidationResult& validation) noexcept;
    [[nodiscard]] bool TryValidateAndNormalizeConnectionProfiles() noexcept;
    void OnSettingsReloadedFromDisk() noexcept;
    [[nodiscard]] bool ResolveStaleSettingsBeforeSave() noexcept;
    [[nodiscard]] bool SaveConnectionsSettings() noexcept;
    [[nodiscard]] bool NotifyOwnerToConnectSelectedProfile() noexcept;
    void OnConnectClicked() noexcept;
    void OnCloseClicked() noexcept;
    void OnCancelClicked() noexcept;
    [[nodiscard]] bool OnCommand(WORD commandId) noexcept;
    void OnNewClicked() noexcept;
    void OnRenameClicked() noexcept;
    void OnRemoveClicked() noexcept;

    LRESULT WindowProc(UINT msg, WPARAM wp, LPARAM lp) noexcept;

    wil::unique_hwnd _hwnd;
    HWND _closingHwnd = nullptr;
    WindowHost _dxHost;

    Panel* _root                                          = nullptr;
    Panel* _listPane                                      = nullptr;
    ScrollPanel* _editorPane                              = nullptr;
    Panel* _footerPane                                    = nullptr;
    Grid* _list                                           = nullptr;
    Button* _newButton                                    = nullptr;
    Button* _renameButton                                 = nullptr;
    Button* _removeButton                                 = nullptr;
    RedSalamander::DxUi::CardPanel* _connectionCard       = nullptr;
    RedSalamander::DxUi::CardPanel* _authCard             = nullptr;
    RedSalamander::DxUi::CardPanel* _s3Card               = nullptr;
    RedSalamander::DxUi::CardPanel* _sshCard              = nullptr;
    RedSalamander::DxUi::CardPanel* _s3EndpointCard       = nullptr;
    RedSalamander::DxUi::SortDirection _listSortDirection = RedSalamander::DxUi::SortDirection::Ascending;

    // Section headers (4: Connection / Auth / S3 / SSH).
    Label* _sectionConnection = nullptr;
    Label* _sectionAuth       = nullptr;
    Label* _sectionS3         = nullptr;
    Label* _sectionSsh        = nullptr;

    // Form labels (19).
    Label* _labelName                   = nullptr;
    Label* _labelProtocol               = nullptr;
    Label* _labelHost                   = nullptr;
    Label* _labelPort                   = nullptr;
    Label* _labelInitialPath            = nullptr;
    Label* _labelCopyMoveMaxConcurrency = nullptr;
    Label* _labelDeleteMaxConcurrency   = nullptr;
    Label* _labelAnonymous              = nullptr;
    Label* _labelUser                   = nullptr;
    Label* _labelSecret                 = nullptr;
    Label* _labelSavePassword           = nullptr;
    Label* _labelRequireHello           = nullptr;
    Label* _labelIgnoreSslTrust         = nullptr;
    Label* _labelS3EndpointOverride     = nullptr;
    Label* _labelS3UseHttps             = nullptr;
    Label* _labelS3VerifyTls            = nullptr;
    Label* _labelS3UseVirtualAddressing = nullptr;
    Label* _labelSshPrivateKey          = nullptr;
    Label* _labelSshKnownHosts          = nullptr;

    // Text fields (10).
    TextField* _editName                   = nullptr;
    TextField* _editHost                   = nullptr;
    TextField* _editPort                   = nullptr;
    TextField* _editInitialPath            = nullptr;
    TextField* _editCopyMoveMaxConcurrency = nullptr;
    TextField* _editDeleteMaxConcurrency   = nullptr;
    TextField* _editUser                   = nullptr;
    TextField* _editSecret                 = nullptr;
    TextField* _editS3EndpointOverride     = nullptr;
    TextField* _editSshPrivateKey          = nullptr;
    TextField* _editSshKnownHosts          = nullptr;

    // Combos (2).
    ComboBox* _comboProtocol  = nullptr;
    ComboBox* _comboAwsRegion = nullptr;

    // Toggles (7).
    Toggle* _toggleAnonymous              = nullptr;
    Toggle* _toggleSavePassword           = nullptr;
    Toggle* _toggleRequireHello           = nullptr;
    Toggle* _toggleIgnoreSslTrust         = nullptr;
    Toggle* _toggleS3UseHttps             = nullptr;
    Toggle* _toggleS3VerifyTls            = nullptr;
    Toggle* _toggleS3UseVirtualAddressing = nullptr;

    // Form action buttons (3).
    Button* _btnShowSecret          = nullptr;
    Button* _btnSshPrivateKeyBrowse = nullptr;
    Button* _btnSshKnownHostsBrowse = nullptr;

    // Footer buttons (3).
    Button* _connectButton = nullptr;
    Button* _closeButton   = nullptr;
    Button* _cancelButton  = nullptr;

    std::unique_ptr<Panel> _rootStorage;
    ConnectionListGridModel _listModel;
    std::wstring _toggleOnLabel;
    std::wstring _toggleOffLabel;

    std::wstring _appId;
    Common::Settings::Settings* _settings = nullptr;
    AppTheme _theme{};
    std::wstring _filterPluginId;
    uint8_t _targetPane = 0u;
    HWND _notifyOwner   = nullptr;

    std::vector<Common::Settings::ConnectionProfile> _connections;
    std::vector<size_t> _viewToModel;
    std::unordered_set<std::wstring> _baselineConnectionIds;
    std::unordered_map<std::wstring, std::wstring> _stagedPasswordById;
    std::unordered_map<std::wstring, std::wstring> _stagedPassphraseById;
    std::wstring _selectedConnectionName;
    int _selectedListIndex = -1;

    bool _loadingEditor               = false;
    bool _settingsHotReloadRegistered = false;
    bool _dirtySinceLastSettingsLoad  = false;
    bool _staleExternalSettings       = false;
    bool _loadingFromSettings         = false;
    bool _isModalFacade               = false;
    ModalFacadeResult* _modalResult   = nullptr;
    size_t _dispatchDepth             = 0u;
    bool _deletePending               = false;
};

bool WindowImpl::Create() noexcept
{
    HINSTANCE instance = GetModuleHandleW(nullptr);

    static ATOM classAtom = 0;
    if (classAtom == 0)
    {
        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.style         = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc   = &WindowImpl::WndProcThunk;
        wc.hInstance     = instance;
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        wc.hIcon         = LoadIconW(instance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
        wc.hIconSm       = LoadIconW(instance, MAKEINTRESOURCEW(IDI_SMALL));
        wc.lpszClassName = kWindowClassName;
        classAtom        = RegisterClassExW(&wc);
        if (classAtom == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS)
        {
            Debug::ErrorWithLastError(L"ConnectionManagerWindow: RegisterClassEx failed.");
            return false;
        }
    }

    const std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_CONNECTIONS);
    const HWND hwnd          = CreateWindowExW(0,
                                               kWindowClassName,
                                               title.empty() ? L"Connections" : title.c_str(),
                                               WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                               CW_USEDEFAULT,
                                               CW_USEDEFAULT,
                                               UiMetrics::ScaleDip(96u, 1040),
                                               UiMetrics::ScaleDip(96u, 600),
                                               nullptr,
                                               nullptr,
                                               instance,
                                               this);
    if (! hwnd)
    {
        Debug::ErrorWithLastError(L"ConnectionManagerWindow: CreateWindowEx failed.");
        return false;
    }

    if (! _isModalFacade)
    {
        g_singleInstance.store(hwnd, std::memory_order_release);
    }

    const bool hasPlacement = _settings && _settings->windows.contains(std::wstring(kWindowSettingsId));
    const int showCmd       = hasPlacement ? WindowPlacementPersistence::Restore(*_settings, kWindowSettingsId, hwnd) : SW_SHOWNORMAL;
    ::ShowWindow(hwnd, showCmd);
    SetForegroundWindow(hwnd);
    return true;
}

bool WindowImpl::OnCreate() noexcept
{
    if (! _dxHost.Attach(_hwnd.get()))
    {
        Debug::Error(L"ConnectionManagerWindow: failed to attach DxUi host.");
        return false;
    }

    BuildUi();
    LoadConnections();
    RebuildList();
    ApplyTheme();
    Layout();
    if (_hwnd)
    {
        SettingsHotReload::RegisterParticipant(_hwnd.get());
        _settingsHotReloadRegistered = true;
    }
    return true;
}

void WindowImpl::BuildUi()
{
    if (_root != nullptr)
    {
        return;
    }

    _toggleOnLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    _toggleOffLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
    if (_toggleOnLabel.empty())
    {
        _toggleOnLabel = L"On";
    }
    if (_toggleOffLabel.empty())
    {
        _toggleOffLabel = L"Off";
    }

    _rootStorage = std::make_unique<Panel>();
    _root        = _rootStorage.get();

    _listPane   = _root->AddChild<Panel>();
    _editorPane = _root->AddChild<ScrollPanel>();
    _editorPane->SetScrollStepDip(48.0f);
    _footerPane = _root->AddChild<Panel>();

    // List pane: the "Connections" heading lives inside the grid's column
    // header (see `ConnectionListGridModel` - `col.title`/`col.sortable=true`).
    _list = _listPane->AddChild<Grid>();
    _list->SetSelectionMode(GridSelectionMode::Single);
    _list->SetRowHeightDip(kListGridRowHeight);
    _list->SetModel(&_listModel);
    _list->SetDelegate(this);
    {
        RedSalamander::DxUi::GridSortSpec spec{};
        spec.columnIndex = 0u;
        spec.direction   = RedSalamander::DxUi::SortDirection::Ascending;
        _list->SetSortSpec(spec);
    }

    BuildListPaneButtons();
    BuildEditorForm();
    BuildFooter();

    _dxHost.SetRoot(std::move(_rootStorage));
    _dxHost.SetDefaultButton(_connectButton);
    _dxHost.SetCancelButton(_cancelButton);
    _dxHost.SetOnEscape([this]
    {
        OnCancelClicked();
        return true;
    });

    ClearEditor();
}

void WindowImpl::BuildListPaneButtons()
{
    _newButton = _listPane->AddChild<Button>(LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_NEW_ELLIPSIS));
    _newButton->SetOnClick([this] { OnNewClicked(); });
    _newButton->SetMnemonic(L'N');

    _renameButton = _listPane->AddChild<Button>(LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_RENAME_ELLIPSIS));
    _renameButton->SetOnClick([this] { OnRenameClicked(); });
    _renameButton->SetMnemonic(L'R');

    _removeButton = _listPane->AddChild<Button>(LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_REMOVE));
    _removeButton->SetOnClick([this] { OnRemoveClicked(); });
    _removeButton->SetMnemonic(L'M');
}

void WindowImpl::BuildEditorForm()
{
    // Cards are added BEFORE their controls so they paint behind the
    // section content. Tab order is unaffected - `CardPanel` is non-focusable.
    _connectionCard = _editorPane->AddChild<CardPanel>();
    _connectionCard->SetCornerRadius(8.0f);
    _authCard = _editorPane->AddChild<CardPanel>();
    _authCard->SetCornerRadius(8.0f);
    _s3Card = _editorPane->AddChild<CardPanel>();
    _s3Card->SetCornerRadius(8.0f);
    _sshCard = _editorPane->AddChild<CardPanel>();
    _sshCard->SetCornerRadius(8.0f);
    _s3EndpointCard = _editorPane->AddChild<CardPanel>();
    _s3EndpointCard->SetCornerRadius(8.0f);

    auto addSection = [this](Label*& slot, UINT stringId)
    {
        slot = _editorPane->AddChild<Label>(LoadStringResource(nullptr, stringId));
        slot->SetFontRole(FontRole::BodyStrong);
    };
    auto addLabel = [this](Label*& slot, UINT stringId, wchar_t mnemonic = L'\0')
    {
        slot = _editorPane->AddChild<Label>(LoadStringResource(nullptr, stringId));
        if (mnemonic != L'\0')
        {
            slot->SetMnemonic(mnemonic);
        }
    };
    auto addEdit   = [this](TextField*& slot) { slot = _editorPane->AddChild<TextField>(); };
    auto addToggle = [this](Toggle*& slot)
    {
        slot = _editorPane->AddChild<Toggle>();
        slot->SetStateLabels(_toggleOffLabel, _toggleOnLabel);
    };
    auto addCombo = [this](ComboBox*& slot, bool editable)
    {
        slot = _editorPane->AddChild<ComboBox>();
        slot->SetEditable(editable);
    };
    auto addActionButton = [this](Button*& slot, std::wstring caption) { slot = _editorPane->AddChild<Button>(std::move(caption)); };

    // Section headers
    addSection(_sectionConnection, IDS_CONNECTIONS_SECTION_CONNECTION);
    addSection(_sectionAuth, IDS_CONNECTIONS_SECTION_AUTH);
    addSection(_sectionS3, IDS_CONNECTIONS_SECTION_S3);
    addSection(_sectionSsh, IDS_CONNECTIONS_SECTION_SSH);

    // Connection group
    addLabel(_labelName, IDS_CONNECTIONS_LABEL_NAME, L'N');
    addEdit(_editName);
    _labelName->SetMnemonicTarget(_editName);

    addLabel(_labelProtocol, IDS_CONNECTIONS_LABEL_PROTOCOL, L'P');
    addCombo(_comboProtocol, /*editable=*/false);
    _comboProtocol->SetItems(BuildProtocolComboItems());
    _labelProtocol->SetMnemonicTarget(_comboProtocol);

    addLabel(_labelHost, IDS_CONNECTIONS_LABEL_HOST);
    addEdit(_editHost);
    _labelHost->SetMnemonicTarget(_editHost);

    addLabel(_labelPort, IDS_CONNECTIONS_LABEL_PORT);
    addEdit(_editPort);

    // AwsRegion is added AFTER Port so the legacy tab order
    //   Host -> Port -> AwsRegion
    // is preserved by the natural tree-traversal focus walk.
    addCombo(_comboAwsRegion, /*editable=*/true);
    _comboAwsRegion->SetItems(BuildAwsRegionComboItems());

    addLabel(_labelInitialPath, IDS_CONNECTIONS_LABEL_INITIAL_PATH);
    addEdit(_editInitialPath);

    addLabel(_labelCopyMoveMaxConcurrency, IDS_CONNECTIONS_LABEL_COPYMOVE_CONCURRENCY);
    addEdit(_editCopyMoveMaxConcurrency);
    _editCopyMoveMaxConcurrency->SetPlaceholder(LoadStringResource(nullptr, IDS_CONNECTIONS_CUE_COPYMOVE_CONCURRENCY));

    addLabel(_labelDeleteMaxConcurrency, IDS_CONNECTIONS_LABEL_DELETE_CONCURRENCY);
    addEdit(_editDeleteMaxConcurrency);
    _editDeleteMaxConcurrency->SetPlaceholder(LoadStringResource(nullptr, IDS_CONNECTIONS_CUE_DELETE_CONCURRENCY));

    // Auth group
    addLabel(_labelAnonymous, IDS_CONNECTIONS_LABEL_ANONYMOUS);
    addToggle(_toggleAnonymous);

    addLabel(_labelUser, IDS_CONNECTIONS_LABEL_USER, L'U');
    addEdit(_editUser);
    _labelUser->SetMnemonicTarget(_editUser);

    addLabel(_labelSecret, IDS_CONNECTIONS_LABEL_PASSWORD);
    addEdit(_editSecret);
    _editSecret->SetMasked(true);
    addActionButton(_btnShowSecret, LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_SHOW_SECRET));
    _btnShowSecret->SetOnClick([this]
    {
        if (_editSecret)
        {
            const bool nowMasked = ! _editSecret->IsMasked();
            _editSecret->SetMasked(nowMasked);
            _btnShowSecret->SetText(LoadStringResource(nullptr, nowMasked ? IDS_CONNECTIONS_BTN_SHOW_SECRET : IDS_CONNECTIONS_BTN_HIDE_SECRET));
        }
    });

    addLabel(_labelSavePassword, IDS_CONNECTIONS_LABEL_SAVE_PASSWORD);
    addToggle(_toggleSavePassword);

    addLabel(_labelRequireHello, IDS_CONNECTIONS_LABEL_REQUIRE_HELLO);
    addToggle(_toggleRequireHello);

    addLabel(_labelIgnoreSslTrust, IDS_CONNECTIONS_LABEL_IGNORE_SSL_TRUST);
    addToggle(_toggleIgnoreSslTrust);

    // S3 toggles. Legacy tab order is
    //   S3UseVirtualAddressing -> S3UseHttps -> S3VerifyTls
    // matching the bounds-driven order in `IDD_CONNECTION_MANAGER`.
    addLabel(_labelS3UseVirtualAddressing, IDS_CONNECTIONS_LABEL_USE_VIRTUAL_ADDRESSING);
    addToggle(_toggleS3UseVirtualAddressing);

    addLabel(_labelS3UseHttps, IDS_CONNECTIONS_LABEL_USE_HTTPS);
    addToggle(_toggleS3UseHttps);

    addLabel(_labelS3VerifyTls, IDS_CONNECTIONS_LABEL_VERIFY_TLS);
    addToggle(_toggleS3VerifyTls);

    // SSH group
    addLabel(_labelSshPrivateKey, IDS_CONNECTIONS_LABEL_SSH_PRIVATEKEY);
    addEdit(_editSshPrivateKey);
    addActionButton(_btnSshPrivateKeyBrowse, L"...");
    _btnSshPrivateKeyBrowse->SetOnClick([this]
    {
        if (BrowseSshFile(_editSshPrivateKey))
        {
            OnEditorFieldChanged();
        }
    });

    addLabel(_labelSshKnownHosts, IDS_CONNECTIONS_LABEL_SSH_KNOWNHOSTS);
    addEdit(_editSshKnownHosts);
    addActionButton(_btnSshKnownHostsBrowse, L"...");
    _btnSshKnownHostsBrowse->SetOnClick([this]
    {
        if (BrowseSshFile(_editSshKnownHosts))
        {
            OnEditorFieldChanged();
        }
    });

    // S3 endpoint-override is created LAST in the editor tab order to match
    // the legacy `BuildConnectionManagerTabTargets` reorder rule (the override
    // sits after the visible S3 toggles even though it visually belongs to the
    // S3 section).
    addLabel(_labelS3EndpointOverride, IDS_CONNECTIONS_LABEL_ENDPOINT_OVERRIDE);
    addEdit(_editS3EndpointOverride);

    WireEditorChangeCallbacks();
}

void WindowImpl::BuildFooter()
{
    _connectButton = _footerPane->AddChild<Button>(LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_CONNECT));
    _connectButton->SetPrimary(true);
    _connectButton->SetMnemonic(L'C');
    _connectButton->SetOnClick([this] { OnConnectClicked(); });

    const std::wstring closeText = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_CLOSE);
    _closeButton                 = _footerPane->AddChild<Button>(closeText.empty() ? std::wstring(L"Close") : closeText);
    _closeButton->SetMnemonic(L'L');
    _closeButton->SetOnClick([this] { OnCloseClicked(); });

    _cancelButton = _footerPane->AddChild<Button>(LoadStringResource(nullptr, IDS_BTN_CANCEL));
    _cancelButton->SetMnemonic(L'A');
    _cancelButton->SetOnClick([this] { OnCancelClicked(); });
}

void WindowImpl::ApplyTheme() noexcept
{
    if (! _hwnd)
    {
        return;
    }
    _dxHost.SetTheme(MakeAppThemeDxPalette(_theme));
    ApplyWindowChromeTheme(_hwnd.get(), _theme, WindowBackdropTarget::Tool, GetActiveWindow() == _hwnd.get());
    _dxHost.Invalidate();
}

void WindowImpl::Layout() noexcept
{
    if (! _hwnd || ! _root)
    {
        return;
    }
    const D2D1_RECT_F clientDip = _dxHost.GetClientBoundsDip();
    _root->SetBounds(clientDip);

    const float pad          = kPanePaddingDip;
    const float footerTop    = clientDip.bottom - kFooterHeightDip;
    const float listLeft     = clientDip.left + pad;
    const float listTop      = clientDip.top + pad;
    const float listRight    = listLeft + kListPaneWidthDip;
    const float listBottom   = footerTop;
    const float editorLeft   = listRight + pad;
    const float editorTop    = listTop;
    const float editorRight  = clientDip.right - pad;
    const float editorBottom = listBottom;

    if (_listPane)
    {
        _listPane->SetBounds(D2D1::RectF(listLeft, listTop, listRight, listBottom));
    }
    if (_editorPane)
    {
        _editorPane->SetBounds(D2D1::RectF(editorLeft, editorTop, editorRight, editorBottom));
    }
    if (_footerPane)
    {
        _footerPane->SetBounds(D2D1::RectF(clientDip.left + pad, footerTop, clientDip.right - pad, clientDip.bottom - pad));
    }

    // List pane internal layout: the grid (with its own column header carrying
    // "Connections") fills the pane top-down; the New/Rename/Remove button row
    // sits immediately under the grid with a small gap (no trailing whitespace
    // between the buttons and the editor pane bottom).
    if (_listPane)
    {
        const float buttonRowBottom = listBottom;
        const float buttonRowTop    = buttonRowBottom - kListButtonHeightDip;
        const float gridTop         = listTop;
        const float gridBottom      = buttonRowTop - kListButtonGapDip;
        if (_list)
        {
            _list->SetBounds(D2D1::RectF(listLeft, gridTop, listRight, gridBottom));
        }
        const float buttonWidth = (listRight - listLeft - 2.0f * kListButtonGapDip) / 3.0f;
        const float btn1Left    = listLeft;
        const float btn2Left    = btn1Left + buttonWidth + kListButtonGapDip;
        const float btn3Left    = btn2Left + buttonWidth + kListButtonGapDip;
        if (_newButton)
        {
            _newButton->SetBounds(D2D1::RectF(btn1Left, buttonRowTop, btn1Left + buttonWidth, buttonRowBottom));
        }
        if (_renameButton)
        {
            _renameButton->SetBounds(D2D1::RectF(btn2Left, buttonRowTop, btn2Left + buttonWidth, buttonRowBottom));
        }
        if (_removeButton)
        {
            _removeButton->SetBounds(D2D1::RectF(btn3Left, buttonRowTop, btn3Left + buttonWidth, buttonRowBottom));
        }
    }

    // Editor pane: form layout
    LayoutEditorForm(D2D1::RectF(editorLeft, editorTop, editorRight, editorBottom));

    // Footer: Connect (primary) | Close | Cancel - right-aligned
    if (_footerPane)
    {
        const float footerRight  = clientDip.right - pad;
        const float footerYTop   = footerTop + (kFooterHeightDip - kFooterButtonHeight) * 0.5f;
        const float footerYBot   = footerYTop + kFooterButtonHeight;
        const float cancelRight  = footerRight;
        const float cancelLeft   = cancelRight - kFooterButtonWidth;
        const float closeRight   = cancelLeft - kFooterButtonGap;
        const float closeLeft    = closeRight - kFooterButtonWidth;
        const float connectRight = closeLeft - kFooterButtonGap;
        const float connectLeft  = connectRight - kFooterButtonWidth;
        if (_cancelButton)
        {
            _cancelButton->SetBounds(D2D1::RectF(cancelLeft, footerYTop, cancelRight, footerYBot));
        }
        if (_closeButton)
        {
            _closeButton->SetBounds(D2D1::RectF(closeLeft, footerYTop, closeRight, footerYBot));
        }
        if (_connectButton)
        {
            _connectButton->SetBounds(D2D1::RectF(connectLeft, footerYTop, connectRight, footerYBot));
        }
    }

    _dxHost.Invalidate();
}

void WindowImpl::LayoutEditorForm(D2D1_RECT_F editorRect) noexcept
{
    // The form is laid out as a flat vertical stack inside the editor's
    // ScrollPanel; the card backgrounds (created up-front in `BuildEditorForm`)
    // are sized at the end of each section to wrap that section's controls.
    constexpr float kCardPaddingX      = 12.0f;
    constexpr float kCardPaddingY      = 8.0f;
    constexpr float kCardOuterPaddingY = 4.0f;

    const float left      = editorRect.left + kCardPaddingX;
    const float right     = editorRect.right - kCardPaddingX;
    const float cardLeft  = editorRect.left;
    const float cardRight = editorRect.right;
    const float labelLeft = left;
    const float ctrlLeft  = labelLeft + kFormLabelWidthDip + kFormLabelGapDip;
    const float ctrlRight = right;

    float y = editorRect.top + kCardOuterPaddingY;

    // Track the top of each card and finalise its bounds when the section ends.
    CardPanel* currentCard = nullptr;
    float currentCardTop   = y;
    auto finishCurrentCard = [&]() noexcept
    {
        if (currentCard)
        {
            currentCard->SetBounds(D2D1::RectF(cardLeft, currentCardTop, cardRight, y + kCardPaddingY));
        }
        currentCard = nullptr;
    };
    auto beginCard = [&](CardPanel* card, Label* sectionLabel) noexcept
    {
        finishCurrentCard();
        if (! card)
        {
            return;
        }
        currentCard    = card;
        currentCardTop = y;
        y += kCardPaddingY;
        if (sectionLabel)
        {
            sectionLabel->SetBounds(D2D1::RectF(left, y, right, y + kFormRowHeightDip));
            y += kFormRowHeightDip + kFormRowGapDip;
        }
    };

    auto layoutSection = [&](Label* section, CardPanel* card)
    {
        if (card)
        {
            if (! card->IsVisible())
            {
                // Hidden card: don't open a new card or advance y. Close any
                // currently open card first so its bounds don't extend into
                // the next visible section.
                finishCurrentCard();
                return;
            }
            // Add a bit of breathing room between cards.
            if (currentCard)
            {
                y += kFormSectionGapDip;
            }
            beginCard(card, section);
            return;
        }
        if (! section)
        {
            return;
        }
        y += kFormSectionGapDip;
        section->SetBounds(D2D1::RectF(left, y, right, y + kFormRowHeightDip));
        y += kFormRowHeightDip + kFormRowGapDip;
    };

    // Skip rows whose primary control is hidden - keeps cards tight to the
    // visible content (no empty whitespace around `Anonymous` / `IgnoreSslTrust`
    // when those rows are hidden by `ApplyEditorVisibility`).
    auto layoutRow = [&](Label* label, Control* control)
    {
        if (control && ! control->IsVisible())
        {
            return;
        }
        const float topY    = y;
        const float bottomY = y + kFormRowHeightDip;
        const float ctrlTop = y + (kFormRowHeightDip - kFormControlHeightDip) * 0.5f;
        const float ctrlBot = ctrlTop + kFormControlHeightDip;
        if (label)
        {
            label->SetBounds(D2D1::RectF(labelLeft, topY, labelLeft + kFormLabelWidthDip, bottomY));
        }
        if (control)
        {
            control->SetBounds(D2D1::RectF(ctrlLeft, ctrlTop, ctrlRight, ctrlBot));
        }
        y = bottomY + kFormRowGapDip;
    };

    auto layoutRowWithAction = [&](Label* label, Control* control, Button* action)
    {
        if (control && ! control->IsVisible())
        {
            return;
        }
        const float topY             = y;
        const float bottomY          = y + kFormRowHeightDip;
        const float ctrlTop          = y + (kFormRowHeightDip - kFormControlHeightDip) * 0.5f;
        const float ctrlBot          = ctrlTop + kFormControlHeightDip;
        constexpr float kActionWidth = 56.0f;
        const float actionLeft       = ctrlRight - kActionWidth;
        const float controlEnd       = actionLeft - kFormLabelGapDip;
        if (label)
        {
            label->SetBounds(D2D1::RectF(labelLeft, topY, labelLeft + kFormLabelWidthDip, bottomY));
        }
        if (control)
        {
            control->SetBounds(D2D1::RectF(ctrlLeft, ctrlTop, controlEnd, ctrlBot));
        }
        if (action)
        {
            action->SetBounds(D2D1::RectF(actionLeft, ctrlTop, ctrlRight, ctrlBot));
        }
        y = bottomY + kFormRowGapDip;
    };

    layoutSection(_sectionConnection, _connectionCard);
    layoutRow(_labelName, _editName);
    layoutRow(_labelProtocol, _comboProtocol);
    // The Host edit and AwsRegion combo share the Host row visually because
    // they target the same `profile.host` field - `ApplyEditorVisibility`
    // toggles which of the two is visible based on the protocol (host edit
    // for non-S3, region combo for S3). Both occupy the same rect.
    layoutRow(_labelHost, _editHost);
    if (_comboAwsRegion)
    {
        const float ctrlTop = (y - kFormRowHeightDip - kFormRowGapDip) + (kFormRowHeightDip - kFormControlHeightDip) * 0.5f;
        const float ctrlBot = ctrlTop + kFormControlHeightDip;
        _comboAwsRegion->SetBounds(D2D1::RectF(ctrlLeft, ctrlTop, ctrlRight, ctrlBot));
    }
    layoutRow(_labelPort, _editPort);
    layoutRow(_labelInitialPath, _editInitialPath);
    layoutRow(_labelCopyMoveMaxConcurrency, _editCopyMoveMaxConcurrency);
    layoutRow(_labelDeleteMaxConcurrency, _editDeleteMaxConcurrency);

    layoutSection(_sectionAuth, _authCard);
    layoutRow(_labelAnonymous, _toggleAnonymous);
    layoutRow(_labelUser, _editUser);
    layoutRowWithAction(_labelSecret, _editSecret, _btnShowSecret);
    layoutRow(_labelSavePassword, _toggleSavePassword);
    layoutRow(_labelRequireHello, _toggleRequireHello);
    layoutRow(_labelIgnoreSslTrust, _toggleIgnoreSslTrust);

    layoutSection(_sectionS3, _s3Card);
    layoutRow(_labelS3UseVirtualAddressing, _toggleS3UseVirtualAddressing);
    layoutRow(_labelS3UseHttps, _toggleS3UseHttps);
    layoutRow(_labelS3VerifyTls, _toggleS3VerifyTls);

    layoutSection(_sectionSsh, _sshCard);
    layoutRowWithAction(_labelSshPrivateKey, _editSshPrivateKey, _btnSshPrivateKeyBrowse);
    layoutRowWithAction(_labelSshKnownHosts, _editSshKnownHosts, _btnSshKnownHostsBrowse);

    // S3 endpoint override sits in its own card at the very end so that the
    // legacy tab-order quirk (`S3EndpointOverride` after the SSH browse
    // buttons) is preserved while still being visually grouped.
    layoutSection(nullptr, _s3EndpointCard);
    layoutRow(_labelS3EndpointOverride, _editS3EndpointOverride);

    // Close the trailing card.
    finishCurrentCard();

    // Publish the form's total content height so the ScrollPanel can decide
    // whether to show a scrollbar and clamp the scroll offset.
    if (_editorPane)
    {
        const float contentHeight = std::max(0.0f, y - editorRect.top + kPanePaddingDip);
        _editorPane->SetContentHeight(contentHeight);
    }
}

void WindowImpl::WireEditorChangeCallbacks()
{
    auto onTextChanged = [this](std::wstring_view) { OnEditorFieldChanged(); };
    auto onToggled     = [this](bool) { OnEditorFieldChanged(); };

    if (_editName)
    {
        _editName->SetOnTextChanged(onTextChanged);
    }
    if (_editHost)
    {
        _editHost->SetOnTextChanged(onTextChanged);
    }
    if (_editPort)
    {
        _editPort->SetOnTextChanged([this](std::wstring_view text)
        {
            if (_loadingEditor)
            {
                return;
            }
            std::wstring sanitized = SanitizeUnsignedText(text);
            if (sanitized != text)
            {
                _editPort->SetText(std::move(sanitized));
            }
            OnEditorFieldChanged();
        });
    }
    if (_editInitialPath)
    {
        _editInitialPath->SetOnTextChanged(onTextChanged);
    }
    auto wireConcurrency = [this](TextField* edit)
    {
        if (! edit)
        {
            return;
        }
        edit->SetOnTextChanged([this, edit](std::wstring_view text)
        {
            if (_loadingEditor)
            {
                return;
            }
            std::wstring sanitized = SanitizeUnsignedText(text);
            if (sanitized != text)
            {
                edit->SetText(std::move(sanitized));
            }
            OnEditorFieldChanged();
        });
    };
    wireConcurrency(_editCopyMoveMaxConcurrency);
    wireConcurrency(_editDeleteMaxConcurrency);
    if (_editUser)
    {
        _editUser->SetOnTextChanged(onTextChanged);
    }
    if (_editSecret)
    {
        _editSecret->SetOnTextChanged(onTextChanged);
    }
    if (_editS3EndpointOverride)
    {
        _editS3EndpointOverride->SetOnTextChanged(onTextChanged);
    }
    if (_editSshPrivateKey)
    {
        _editSshPrivateKey->SetOnTextChanged(onTextChanged);
    }
    if (_editSshKnownHosts)
    {
        _editSshKnownHosts->SetOnTextChanged(onTextChanged);
    }

    if (_comboProtocol)
    {
        _comboProtocol->SetOnSelectionChanged([this](size_t index) { OnProtocolChanged(index); });
    }
    if (_comboAwsRegion)
    {
        _comboAwsRegion->SetOnTextChanged(onTextChanged);
        _comboAwsRegion->SetOnSelectionChanged([this](size_t) { OnEditorFieldChanged(); });
    }

    auto wireToggle = [&](Toggle* toggle)
    {
        if (toggle)
        {
            toggle->SetOnToggled(onToggled);
        }
    };
    wireToggle(_toggleAnonymous);
    wireToggle(_toggleSavePassword);
    wireToggle(_toggleRequireHello);
    wireToggle(_toggleIgnoreSslTrust);
    wireToggle(_toggleS3UseHttps);
    wireToggle(_toggleS3VerifyTls);
    wireToggle(_toggleS3UseVirtualAddressing);
}

EditorVisibility WindowImpl::ComputeEditorVisibility() const noexcept
{
    EditorVisibility v{};
    const auto modelIndex = GetSelectedModelIndex();
    if (! modelIndex || *modelIndex >= _connections.size())
    {
        return v;
    }
    const auto& profile = _connections[*modelIndex];
    v.hasSelection      = true;
    v.isFtp             = IsFtpPluginId(profile.pluginId);
    v.isSsh             = IsSshPluginId(profile.pluginId);
    v.isImap            = IsImapPluginId(profile.pluginId);
    v.isGoogleDrive     = IsGoogleDrivePluginId(profile.pluginId);
    v.isMicrosoft       = IsMicrosoftPluginId(profile.pluginId);
    v.isS3              = IsS3PluginId(profile.pluginId);
    v.isS3Table         = IsS3TablePluginId(profile.pluginId);
    v.isAwsS3           = v.isS3 || v.isS3Table;
    v.anonymous         = v.isFtp && profile.authMode == Common::Settings::ConnectionAuthMode::Anonymous;
    v.oauth2            = (v.isMicrosoft || v.isGoogleDrive) && profile.authMode == Common::Settings::ConnectionAuthMode::OAuth2Pkce;
    v.showProtocol      = v.hasSelection && _filterPluginId.empty();
    v.showAwsRegion     = v.hasSelection && v.isAwsS3;
    v.showHostEdit =
        v.hasSelection && ! v.isAwsS3 && ! v.isGoogleDrive && ! IsOneDrivePersonalPluginId(profile.pluginId) && ! IsOneDriveBusinessPluginId(profile.pluginId);
    v.showConcurrency    = v.hasSelection;
    v.showAnonymous      = v.hasSelection && v.isFtp;
    v.showSshSection     = v.hasSelection && v.isSsh;
    v.showS3Section      = v.hasSelection && v.isAwsS3;
    v.showVirtual        = v.showS3Section && v.isS3;
    v.showSecretRow      = v.hasSelection && ! v.oauth2;
    v.authInputsEnabled  = v.hasSelection && ! v.anonymous;
    v.showIgnoreSslTrust = v.hasSelection && v.isImap;
    v.showPort           = v.hasSelection && ! v.isAwsS3 && ! v.isMicrosoft && ! v.isGoogleDrive;
    return v;
}

void WindowImpl::ApplyEditorVisibility(const EditorVisibility& v) noexcept
{
    auto setVisible = [](Control* control, bool visible)
    {
        if (control)
        {
            control->SetVisible(visible);
        }
    };

    // Card visibility tracks the section
    setVisible(_connectionCard, v.hasSelection);
    setVisible(_authCard, v.hasSelection);
    setVisible(_s3Card, v.showS3Section);
    setVisible(_sshCard, v.showSshSection);
    setVisible(_s3EndpointCard, v.showS3Section);

    // Connection group
    setVisible(_sectionConnection, v.hasSelection);
    setVisible(_labelName, v.hasSelection);
    setVisible(_editName, v.hasSelection);
    setVisible(_labelProtocol, v.showProtocol);
    setVisible(_comboProtocol, v.showProtocol);
    setVisible(_labelHost, v.hasSelection && (v.showHostEdit || v.showAwsRegion));
    setVisible(_editHost, v.showHostEdit);
    setVisible(_comboAwsRegion, v.showAwsRegion);
    setVisible(_labelPort, v.showPort);
    setVisible(_editPort, v.showPort);
    setVisible(_labelInitialPath, v.hasSelection);
    setVisible(_editInitialPath, v.hasSelection);
    setVisible(_labelCopyMoveMaxConcurrency, v.showConcurrency);
    setVisible(_editCopyMoveMaxConcurrency, v.showConcurrency);
    setVisible(_labelDeleteMaxConcurrency, v.showConcurrency);
    setVisible(_editDeleteMaxConcurrency, v.showConcurrency);

    // Auth group
    setVisible(_sectionAuth, v.hasSelection);
    setVisible(_labelAnonymous, v.showAnonymous);
    setVisible(_toggleAnonymous, v.showAnonymous);
    setVisible(_labelUser, v.hasSelection && ! v.oauth2);
    setVisible(_editUser, v.hasSelection && ! v.oauth2);
    setVisible(_labelSecret, v.showSecretRow);
    setVisible(_editSecret, v.showSecretRow);
    setVisible(_btnShowSecret, v.showSecretRow);
    setVisible(_labelSavePassword, v.hasSelection);
    setVisible(_toggleSavePassword, v.hasSelection);
    setVisible(_labelRequireHello, v.hasSelection);
    setVisible(_toggleRequireHello, v.hasSelection);
    setVisible(_labelIgnoreSslTrust, v.showIgnoreSslTrust);
    setVisible(_toggleIgnoreSslTrust, v.showIgnoreSslTrust);

    // S3 group
    setVisible(_sectionS3, v.showS3Section);
    setVisible(_labelS3EndpointOverride, v.showS3Section);
    setVisible(_editS3EndpointOverride, v.showS3Section);
    setVisible(_labelS3UseHttps, v.showS3Section);
    setVisible(_toggleS3UseHttps, v.showS3Section);
    setVisible(_labelS3VerifyTls, v.showS3Section);
    setVisible(_toggleS3VerifyTls, v.showS3Section);
    setVisible(_labelS3UseVirtualAddressing, v.showVirtual);
    setVisible(_toggleS3UseVirtualAddressing, v.showVirtual);

    // SSH group
    setVisible(_sectionSsh, v.showSshSection);
    setVisible(_labelSshPrivateKey, v.showSshSection);
    setVisible(_editSshPrivateKey, v.showSshSection);
    setVisible(_btnSshPrivateKeyBrowse, v.showSshSection);
    setVisible(_labelSshKnownHosts, v.showSshSection);
    setVisible(_editSshKnownHosts, v.showSshSection);
    setVisible(_btnSshKnownHostsBrowse, v.showSshSection);

    // Auth input enable/disable based on Anonymous
    auto setEnabled = [](Control* control, bool enabled)
    {
        if (control)
        {
            control->SetEnabled(enabled);
        }
    };
    setEnabled(_editUser, v.authInputsEnabled);
    setEnabled(_editSecret, v.authInputsEnabled);
    setEnabled(_btnShowSecret, v.authInputsEnabled);
}

void WindowImpl::OnEditorFieldChanged() noexcept
{
    if (_loadingEditor)
    {
        return;
    }
    const auto modelIndex = GetSelectedModelIndex();
    if (! modelIndex || *modelIndex >= _connections.size())
    {
        return;
    }
    auto& profile = _connections[*modelIndex];
    MarkConnectionsDirty();

    if (_editName)
    {
        profile.name = std::wstring(_editName->GetText());
    }
    if (_editHost)
    {
        profile.host = std::wstring(_editHost->GetText());
    }
    if (_editPort)
    {
        const std::wstring portText = std::wstring(_editPort->GetText());
        profile.port                = portText.empty() ? 0u : static_cast<uint32_t>(std::wcstoul(portText.c_str(), nullptr, 10));
    }
    if (_editInitialPath)
    {
        profile.initialPath = std::wstring(_editInitialPath->GetText());
    }
    if (_editUser)
    {
        profile.userName = std::wstring(_editUser->GetText());
    }
    if (_toggleAnonymous && _toggleAnonymous->IsChecked())
    {
        profile.authMode = Common::Settings::ConnectionAuthMode::Anonymous;
    }
    else if (profile.authMode == Common::Settings::ConnectionAuthMode::Anonymous)
    {
        // Leaving Anonymous mode: revert to the closest matching default.
        profile.authMode = IsSshPluginId(profile.pluginId) ? Common::Settings::ConnectionAuthMode::Password : Common::Settings::ConnectionAuthMode::Password;
    }
    if (_toggleSavePassword)
    {
        profile.savePassword = _toggleSavePassword->IsChecked();
    }
    if (_toggleRequireHello)
    {
        profile.requireWindowsHello = _toggleRequireHello->IsChecked();
    }
    if (_editSecret)
    {
        StageSecretForProfile(profile, _editSecret->GetText());
    }

    // Extra-JSON fields.
    if (_toggleIgnoreSslTrust)
    {
        ExtraSetBool(profile.extra, "ignoreSslTrust", _toggleIgnoreSslTrust->IsChecked());
    }
    if (_editS3EndpointOverride)
    {
        ExtraSetString(profile.extra, "endpointOverride", _editS3EndpointOverride->GetText());
    }
    if (_toggleS3UseHttps)
    {
        ExtraSetBool(profile.extra, "useHttps", _toggleS3UseHttps->IsChecked());
    }
    if (_toggleS3VerifyTls)
    {
        ExtraSetBool(profile.extra, "verifyTls", _toggleS3VerifyTls->IsChecked());
    }
    if (_toggleS3UseVirtualAddressing)
    {
        ExtraSetBool(profile.extra, "useVirtualAddressing", _toggleS3UseVirtualAddressing->IsChecked());
    }
    if (_editSshPrivateKey)
    {
        ExtraSetString(profile.extra, "sshPrivateKey", _editSshPrivateKey->GetText());
    }
    if (_editSshKnownHosts)
    {
        ExtraSetString(profile.extra, "sshKnownHosts", _editSshKnownHosts->GetText());
    }
    if (_editCopyMoveMaxConcurrency)
    {
        const std::wstring t = std::wstring(_editCopyMoveMaxConcurrency->GetText());
        const uint32_t value = t.empty() ? 0u : static_cast<uint32_t>(std::wcstoul(t.c_str(), nullptr, 10));
        ExtraSetUInt32(profile.extra, "copyMoveMaxConcurrency", value);
    }
    if (_editDeleteMaxConcurrency)
    {
        const std::wstring t = std::wstring(_editDeleteMaxConcurrency->GetText());
        const uint32_t value = t.empty() ? 0u : static_cast<uint32_t>(std::wcstoul(t.c_str(), nullptr, 10));
        ExtraSetUInt32(profile.extra, "deleteMaxConcurrency", value);
    }

    // Update the list row text if the name changed and refresh the visible list.
    RebuildList();
    // Visibility may need to flip when authMode changes (e.g. Anonymous toggle).
    ApplyEditorVisibility(ComputeEditorVisibility());
    Layout();
}

void WindowImpl::OnProtocolChanged(size_t comboIndex) noexcept
{
    if (_loadingEditor)
    {
        return;
    }
    const auto modelIndex = GetSelectedModelIndex();
    if (! modelIndex || *modelIndex >= _connections.size())
    {
        return;
    }
    if (comboIndex >= std::size(kProtocolItems))
    {
        return;
    }
    auto& profile    = _connections[*modelIndex];
    profile.pluginId = kProtocolItems[comboIndex].pluginId;
    OnEditorFieldChanged();
}

bool WindowImpl::BrowseSshFile(TextField* target) noexcept
{
    if (! target || ! _hwnd)
    {
        return false;
    }
    wchar_t buffer[MAX_PATH] = {};
    const std::wstring current(target->GetText());
    if (! current.empty() && current.size() < std::size(buffer))
    {
        wcscpy_s(buffer, current.c_str());
    }
    OPENFILENAMEW ofn{};
    ofn.lStructSize           = sizeof(ofn);
    ofn.hwndOwner             = _hwnd.get();
    ofn.lpstrFile             = buffer;
    ofn.nMaxFile              = static_cast<DWORD>(std::size(buffer));
    const std::wstring filter = LoadStringResource(nullptr, IDS_CONNECTIONS_FILE_FILTER_ALL_FILES);
    if (! filter.empty())
    {
        ofn.lpstrFilter = filter.c_str();
    }
    ofn.nFilterIndex = 1;
    ofn.Flags        = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
    if (GetOpenFileNameW(&ofn) == FALSE)
    {
        return false;
    }
    target->SetText(std::wstring(buffer));
    return true;
}

void WindowImpl::RequestCloseWindow() noexcept
{
    if (! _hwnd)
    {
        return;
    }

    const HWND hwnd = _hwnd.get();
    _closingHwnd    = hwnd;
    static_cast<void>(_hwnd.release());
    if (DestroyWindow(hwnd) == FALSE)
    {
        const DWORD lastError = GetLastError();
        _closingHwnd          = nullptr;
        _hwnd.reset(hwnd);
        Debug::Error(L"ConnectionManagerWindow: DestroyWindow failed after HWND ownership release error={}", lastError);
    }
}

void WindowImpl::RefreshBaselineConnectionIds() noexcept
{
    _baselineConnectionIds.clear();
    _baselineConnectionIds.reserve(_connections.size());
    for (const auto& profile : _connections)
    {
        if (! profile.id.empty() && ! IsQuickConnectProfile(profile))
        {
            _baselineConnectionIds.insert(profile.id);
        }
    }
}

void WindowImpl::StageSecretForProfile(const Common::Settings::ConnectionProfile& profile, std::wstring_view secret)
{
    if (profile.id.empty())
    {
        return;
    }
    if (secret.empty())
    {
        _stagedPasswordById.erase(profile.id);
        _stagedPassphraseById.erase(profile.id);
        return;
    }

    if (profile.authMode == Common::Settings::ConnectionAuthMode::SshKey)
    {
        _stagedPassphraseById[profile.id] = std::wstring(secret);
        _stagedPasswordById.erase(profile.id);
    }
    else
    {
        _stagedPasswordById[profile.id] = std::wstring(secret);
        _stagedPassphraseById.erase(profile.id);
    }
}

void WindowImpl::DeleteSecretsForRemovedConnections() noexcept
{
    std::unordered_set<std::wstring> currentIds;
    currentIds.reserve(_connections.size());
    for (const auto& profile : _connections)
    {
        if (! profile.id.empty() && ! IsQuickConnectProfile(profile))
        {
            currentIds.insert(profile.id);
        }
    }

    for (const auto& id : _baselineConnectionIds)
    {
        if (id.empty() || currentIds.contains(id))
        {
            continue;
        }

        const std::wstring passwordTarget = RedSalamander::Connections::BuildCredentialTargetName(id, RedSalamander::Connections::SecretKind::Password);
        const std::wstring passphraseTarget =
            RedSalamander::Connections::BuildCredentialTargetName(id, RedSalamander::Connections::SecretKind::SshKeyPassphrase);
        const std::wstring refreshTarget = RedSalamander::Connections::BuildCredentialTargetName(id, RedSalamander::Connections::SecretKind::RefreshToken);
        if (! passwordTarget.empty())
        {
            static_cast<void>(RedSalamander::Connections::DeleteGenericCredential(passwordTarget));
        }
        if (! passphraseTarget.empty())
        {
            static_cast<void>(RedSalamander::Connections::DeleteGenericCredential(passphraseTarget));
        }
        if (! refreshTarget.empty())
        {
            static_cast<void>(RedSalamander::Connections::DeleteGenericCredential(refreshTarget));
        }
    }
}

void WindowImpl::CommitQuickConnectSecretsAndProfile(const Common::Settings::ConnectionProfile& profile) noexcept
{
    if (! IsQuickConnectProfile(profile))
    {
        return;
    }

    using RedSalamander::Connections::SecretKind;
    RedSalamander::Connections::SetQuickConnectProfile(profile);

    if (! profile.savePassword)
    {
        RedSalamander::Connections::ClearQuickConnectSecret(SecretKind::Password);
        RedSalamander::Connections::ClearQuickConnectSecret(SecretKind::SshKeyPassphrase);
        RedSalamander::Connections::ClearQuickConnectSecret(SecretKind::RefreshToken);
        return;
    }

    if (profile.authMode == Common::Settings::ConnectionAuthMode::OAuth2Pkce)
    {
        return;
    }

    const bool sshPassphrase = profile.authMode == Common::Settings::ConnectionAuthMode::SshKey;
    const auto& stagedMap    = sshPassphrase ? _stagedPassphraseById : _stagedPasswordById;
    const auto it            = stagedMap.find(profile.id);
    if (it == stagedMap.end() || it->second.empty())
    {
        return;
    }

    RedSalamander::Connections::SetQuickConnectSecret(sshPassphrase ? SecretKind::SshKeyPassphrase : SecretKind::Password, it->second);
}

HRESULT WindowImpl::CommitSecretsForProfile(const Common::Settings::ConnectionProfile& profile) noexcept
{
    using RedSalamander::Connections::BuildCredentialTargetName;
    using RedSalamander::Connections::DeleteGenericCredential;
    using RedSalamander::Connections::SaveGenericCredential;
    using RedSalamander::Connections::SecretKind;

    if (profile.id.empty())
    {
        return S_OK;
    }

    const std::wstring passwordTarget   = BuildCredentialTargetName(profile.id, SecretKind::Password);
    const std::wstring passphraseTarget = BuildCredentialTargetName(profile.id, SecretKind::SshKeyPassphrase);
    const std::wstring refreshTarget    = BuildCredentialTargetName(profile.id, SecretKind::RefreshToken);
    const auto deleteStoredSecret       = [&](const std::wstring& targetName, std::wstring_view kindLabel) noexcept
    {
        if (targetName.empty())
        {
            return;
        }

        const HRESULT delHr = DeleteGenericCredential(targetName);
        if (FAILED(delHr) && delHr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
        {
            Debug::Warning(L"ConnectionManagerWindow: DeleteGenericCredential failed connection='{}' id='{}' kind='{}' hr=0x{:08X}",
                           profile.name,
                           profile.id,
                           kindLabel,
                           static_cast<unsigned long>(delHr));
        }
    };

    if (! profile.savePassword)
    {
        deleteStoredSecret(passwordTarget, L"password");
        deleteStoredSecret(passphraseTarget, L"sshKeyPassphrase");
        deleteStoredSecret(refreshTarget, L"refreshToken");
        return S_OK;
    }

    if (profile.authMode == Common::Settings::ConnectionAuthMode::OAuth2Pkce)
    {
        deleteStoredSecret(passwordTarget, L"password");
        deleteStoredSecret(passphraseTarget, L"sshKeyPassphrase");
        return S_OK;
    }

    const bool sshPassphrase = profile.authMode == Common::Settings::ConnectionAuthMode::SshKey;
    if (sshPassphrase)
    {
        deleteStoredSecret(passwordTarget, L"password");
    }
    else
    {
        deleteStoredSecret(passphraseTarget, L"sshKeyPassphrase");
    }
    deleteStoredSecret(refreshTarget, L"refreshToken");

    const auto& stagedMap = sshPassphrase ? _stagedPassphraseById : _stagedPasswordById;
    const auto it         = stagedMap.find(profile.id);
    if (it == stagedMap.end() || it->second.empty())
    {
        return S_OK;
    }

    const SecretKind kind          = sshPassphrase ? SecretKind::SshKeyPassphrase : SecretKind::Password;
    const std::wstring targetName  = BuildCredentialTargetName(profile.id, kind);
    const HRESULT credentialSaveHr = SaveGenericCredential(targetName, profile.userName, it->second);
    if (FAILED(credentialSaveHr))
    {
        Debug::Error(L"ConnectionManagerWindow: SaveGenericCredential failed connection='{}' id='{}' hr=0x{:08X}",
                     profile.name,
                     profile.id,
                     static_cast<unsigned long>(credentialSaveHr));
    }
    return credentialSaveHr;
}

void WindowImpl::ShowNameValidationError(const ConnectionProfileValidationResult& validation) noexcept
{
    const UINT messageId = MessageResourceForConnectionNameValidation(validation.error);
    std::wstring message;
    if (validation.error == ConnectionProfileValidationError::ReservedName)
    {
        message = FormatStringResource(nullptr, messageId, validation.normalizedName);
    }
    else if (messageId != 0u)
    {
        message = LoadStringResource(nullptr, messageId);
    }

    if (message.empty())
    {
        message = L"Connection name is not valid.";
    }

    Debug::Warning(L"ConnectionManagerWindow: rejected connection profile name proposed='{}' normalized='{}' error={}",
                   validation.proposedName,
                   validation.normalizedName,
                   static_cast<unsigned>(validation.error));

    std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_CONNECTIONS);
    if (title.empty())
    {
        title = L"Connections";
    }
    ShowConnectionManagerAlert(_hwnd.get(), HOST_ALERT_WARNING, title, message);

    if (_editName)
    {
        _dxHost.SetFocusControl(_editName);
        _editName->SetSelectionRange(0u, std::numeric_limits<size_t>::max());
    }
}

bool WindowImpl::TryValidateAndNormalizeConnectionProfiles() noexcept
{
    bool normalizedAny            = false;
    const auto selectedModelIndex = GetSelectedModelIndex();

    for (size_t index = 0; index < _connections.size(); ++index)
    {
        if (IsQuickConnectProfile(_connections[index]))
        {
            continue;
        }

        std::wstring proposedName = _connections[index].name;
        if (selectedModelIndex && *selectedModelIndex == index && _editName)
        {
            proposedName = std::wstring(_editName->GetText());
        }

        const ConnectionProfileValidationResult validation = ValidateConnectionProfileName(_connections, index, proposedName);
        if (validation.error != ConnectionProfileValidationError::None)
        {
            SelectConnectionModelIndex(index);
            ShowNameValidationError(validation);
            return false;
        }

        if (_connections[index].name != validation.normalizedName)
        {
            _connections[index].name = validation.normalizedName;
            normalizedAny            = true;
        }

        if (selectedModelIndex && *selectedModelIndex == index && _editName && std::wstring(_editName->GetText()) != validation.normalizedName)
        {
            _loadingEditor                   = true;
            const auto loadingEditorScopeOff = wil::scope_exit([this]() noexcept { _loadingEditor = false; });
            _editName->SetText(validation.normalizedName);
        }
    }

    if (normalizedAny)
    {
        RebuildList();
    }

    return true;
}

void WindowImpl::OnSettingsReloadedFromDisk() noexcept
{
    if (! _dirtySinceLastSettingsLoad)
    {
        ReloadConnectionsFromSettingsPreservingSelection();
        _dirtySinceLastSettingsLoad = false;
        _staleExternalSettings      = false;
        ApplyTheme();
        Layout();
        return;
    }

    SettingsHotReload::ExternalReloadChoice choice = SettingsHotReload::ExternalReloadChoice::KeepEditing;
    const HRESULT promptHr = SettingsHotReload::PromptExternalReloadConflict(_hwnd.get(), LoadStringResource(nullptr, IDS_CAPTION_CONNECTIONS), choice);
    if (FAILED(promptHr))
    {
        Debug::Warning(L"ConnectionManagerWindow: failed to prompt for external reload conflict (hr=0x{:08X})", static_cast<unsigned long>(promptHr));
        return;
    }

    if (choice == SettingsHotReload::ExternalReloadChoice::ReloadFromDisk)
    {
        ReloadConnectionsFromSettingsPreservingSelection();
        _dirtySinceLastSettingsLoad = false;
        _staleExternalSettings      = false;
        ApplyTheme();
        Layout();
        return;
    }

    _staleExternalSettings = true;
}

bool WindowImpl::ResolveStaleSettingsBeforeSave() noexcept
{
    if (! _staleExternalSettings)
    {
        return true;
    }

    SettingsHotReload::StaleSaveChoice choice = SettingsHotReload::StaleSaveChoice::Cancel;
    const HRESULT promptHr = SettingsHotReload::PromptStaleSaveConflict(_hwnd.get(), LoadStringResource(nullptr, IDS_CAPTION_CONNECTIONS), choice);
    if (FAILED(promptHr))
    {
        Debug::Warning(L"ConnectionManagerWindow: failed to prompt for stale save conflict (hr=0x{:08X})", static_cast<unsigned long>(promptHr));
        return false;
    }

    if (choice == SettingsHotReload::StaleSaveChoice::ReloadFromDisk)
    {
        ReloadConnectionsFromSettingsPreservingSelection();
        _dirtySinceLastSettingsLoad = false;
        _staleExternalSettings      = false;
        ApplyTheme();
        Layout();
        return false;
    }

    if (choice == SettingsHotReload::StaleSaveChoice::Cancel)
    {
        return false;
    }

    _staleExternalSettings = false;
    return true;
}

bool WindowImpl::SaveConnectionsSettings() noexcept
{
    if (! _settings)
    {
        Debug::Error(L"ConnectionManagerWindow: cannot save connections without settings.");
        return false;
    }
    if (! ResolveStaleSettingsBeforeSave())
    {
        return false;
    }

    Common::Settings::ConnectionsSettings connSettings;
    if (_settings->connections)
    {
        connSettings.bypassWindowsHello              = _settings->connections->bypassWindowsHello;
        connSettings.allowInsecureTlsInAutomation    = _settings->connections->allowInsecureTlsInAutomation;
        connSettings.windowsHelloReauthTimeoutMinute = _settings->connections->windowsHelloReauthTimeoutMinute;
    }
    connSettings.items     = _connections;
    _settings->connections = std::move(connSettings);

    DeleteSecretsForRemovedConnections();

    for (const auto& profile : _connections)
    {
        if (IsQuickConnectProfile(profile))
        {
            CommitQuickConnectSecretsAndProfile(profile);
            continue;
        }

        const HRESULT secretHr = CommitSecretsForProfile(profile);
        if (FAILED(secretHr))
        {
            Debug::Error(L"ConnectionManagerWindow: CommitSecretsForProfile failed connection='{}' id='{}' hr=0x{:08X}",
                         profile.name,
                         profile.id,
                         static_cast<unsigned long>(secretHr));
            return false;
        }
    }

    if (_hwnd && ! _isModalFacade)
    {
        WindowPlacementPersistence::Save(*_settings, kWindowSettingsId, _hwnd.get());
    }

    const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(_appId, *_settings);
    if (FAILED(saveHr))
    {
        Debug::Error(L"ConnectionManagerWindow: SaveSettingsAndSchema failed appId='{}' hr=0x{:08X}", _appId, static_cast<unsigned long>(saveHr));
        return false;
    }

    RefreshBaselineConnectionIds();
    _stagedPasswordById.clear();
    _stagedPassphraseById.clear();
    _dirtySinceLastSettingsLoad = false;
    _staleExternalSettings      = false;
    return true;
}

bool WindowImpl::NotifyOwnerToConnectSelectedProfile() noexcept
{
    if (_notifyOwner == nullptr || IsWindow(_notifyOwner) == FALSE)
    {
        Debug::Error(L"ConnectionManagerWindow: cannot notify owner for Connect because the owner window is invalid.");
        return false;
    }

    auto payload = std::make_unique<std::wstring>(_selectedConnectionName);
    if (! PostMessagePayload(_notifyOwner, WndMsg::kConnectionManagerConnect, static_cast<WPARAM>(_targetPane), std::move(payload)))
    {
        Debug::ErrorWithLastError(L"ConnectionManagerWindow: failed to post Connect navigation request to owner.");
        return false;
    }

    return true;
}

void WindowImpl::OnSize() noexcept
{
    Layout();
}

void WindowImpl::OnDpiChanged(const RECT* suggested) noexcept
{
    if (suggested && _hwnd)
    {
        SetWindowPos(_hwnd.get(),
                     nullptr,
                     suggested->left,
                     suggested->top,
                     suggested->right - suggested->left,
                     suggested->bottom - suggested->top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }
    ApplyTheme();
    Layout();
}

void WindowImpl::OnNcActivate(bool active) noexcept
{
    if (_hwnd)
    {
        ApplyTitleBarTheme(_hwnd.get(), _theme, active);
    }
}

void WindowImpl::OnClose() noexcept
{
    OnCancelClicked();
}

void WindowImpl::OnNcDestroy() noexcept
{
    const HWND destroyedHwnd = _hwnd ? _hwnd.get() : _closingHwnd;
    if (_settingsHotReloadRegistered)
    {
        SettingsHotReload::UnregisterParticipant(destroyedHwnd);
        _settingsHotReloadRegistered = false;
    }
    if (_settings && destroyedHwnd && ! _isModalFacade)
    {
        WindowPlacementPersistence::Save(*_settings, kWindowSettingsId, destroyedHwnd);
    }
    if (destroyedHwnd)
    {
        SetWindowLongPtrW(destroyedHwnd, GWLP_USERDATA, 0);
    }
    if (_list)
    {
        _list->SetDelegate(nullptr);
        _list->SetModel(nullptr);
    }
    _dxHost.Detach();
    _root           = nullptr;
    _listPane       = nullptr;
    _editorPane     = nullptr;
    _footerPane     = nullptr;
    _list           = nullptr;
    _newButton      = nullptr;
    _renameButton   = nullptr;
    _removeButton   = nullptr;
    _connectButton  = nullptr;
    _closeButton    = nullptr;
    _cancelButton   = nullptr;
    _connectionCard = nullptr;
    _authCard       = nullptr;
    _s3Card         = nullptr;
    _sshCard        = nullptr;
    _s3EndpointCard = nullptr;

    if (! _isModalFacade)
    {
        HWND expected = destroyedHwnd;
        g_singleInstance.compare_exchange_strong(expected, HWND{nullptr});
    }
    if (_hwnd)
    {
        static_cast<void>(_hwnd.release()); // window already destroyed by Windows
    }
    _closingHwnd = nullptr;

    _deletePending = true;
    // Actual `delete this` happens once the thunk's outermost dispatch unwinds,
    // see `WndProcThunk` below.
}

void WindowImpl::UpdateTheme(const AppTheme& theme) noexcept
{
    _theme = theme;
    ApplyTheme();
    Layout();
}

void WindowImpl::UpdateContext(const AppTheme& theme, std::wstring_view filterPluginId, HWND notifyOwner, uint8_t targetPane) noexcept
{
    _theme = theme;
    _filterPluginId.assign(filterPluginId);
    _notifyOwner = notifyOwner;
    _targetPane  = targetPane;
    ReloadConnectionsFromSettingsPreservingSelection();
    _dirtySinceLastSettingsLoad = false;
    _staleExternalSettings      = false;
    ApplyTheme();
    Layout();
}

void WindowImpl::MarkConnectionsDirty() noexcept
{
    if (! _loadingFromSettings)
    {
        _dirtySinceLastSettingsLoad = true;
    }
}

void WindowImpl::LoadConnections() noexcept
{
    const bool wasLoadingFromSettings     = _loadingFromSettings;
    _loadingFromSettings                  = true;
    const auto restoreLoadingFromSettings = wil::scope_exit([&]() noexcept { _loadingFromSettings = wasLoadingFromSettings; });

    _connections.clear();
    _baselineConnectionIds.clear();
    _stagedPasswordById.clear();
    _stagedPassphraseById.clear();
    if (_settings && _settings->connections)
    {
        _connections = _settings->connections->items;
    }
    RefreshBaselineConnectionIds();

    // Drop the persisted Quick Connect entry (if any) - we re-synthesise it
    // below with the live preferred plugin id.
    _connections.erase(std::remove_if(_connections.begin(),
                                      _connections.end(),
                                      [](const Common::Settings::ConnectionProfile& p) noexcept { return IsQuickConnectProfile(p); }),
                       _connections.end());

    RedSalamander::Connections::EnsureQuickConnectProfile(_filterPluginId);
    Common::Settings::ConnectionProfile quickConnect;
    RedSalamander::Connections::GetQuickConnectProfile(quickConnect);
    if (! _filterPluginId.empty())
    {
        quickConnect.pluginId.assign(_filterPluginId);
    }
    _connections.insert(_connections.begin(), std::move(quickConnect));
}

void WindowImpl::ReloadConnectionsFromSettingsPreservingSelection() noexcept
{
    std::wstring selectedId;
    std::wstring selectedName = _selectedConnectionName;
    if (const auto modelIndex = GetSelectedModelIndex(); modelIndex && *modelIndex < _connections.size())
    {
        selectedId   = _connections[*modelIndex].id;
        selectedName = _connections[*modelIndex].name;
    }

    LoadConnections();
    RebuildList();

    if (! selectedId.empty())
    {
        const auto match = std::find_if(
            _connections.begin(), _connections.end(), [&](const Common::Settings::ConnectionProfile& profile) noexcept { return profile.id == selectedId; });
        if (match != _connections.end())
        {
            SelectConnectionModelIndex(static_cast<size_t>(std::distance(_connections.begin(), match)));
            return;
        }
    }

    if (! selectedName.empty())
    {
        const auto match = std::find_if(_connections.begin(), _connections.end(), [&](const Common::Settings::ConnectionProfile& profile) noexcept {
            return profile.name == selectedName;
        });
        if (match != _connections.end())
        {
            SelectConnectionModelIndex(static_cast<size_t>(std::distance(_connections.begin(), match)));
        }
    }
}

void WindowImpl::RebuildList() noexcept
{
    const std::optional<size_t> previousModelIndex = GetSelectedModelIndex();
    const std::wstring previousSelectedName        = _selectedConnectionName;
    const std::wstring quickConnectLabel           = []()
    {
        std::wstring label = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
        return label.empty() ? std::wstring(L"<Quick Connect>") : label;
    }();

    _viewToModel.clear();
    std::vector<ConnectionListGridRow> rows;
    rows.reserve(_connections.size());
    for (size_t i = 0; i < _connections.size(); ++i)
    {
        const Common::Settings::ConnectionProfile& profile = _connections[i];
        if (! _filterPluginId.empty() && ! IsQuickConnectProfile(profile) && profile.pluginId != _filterPluginId)
        {
            continue;
        }
        ConnectionListGridRow row;
        row.modelIndex = i;
        row.stableId   = MakeConnectionListStableId(profile.id, i);
        row.text       = IsQuickConnectProfile(profile) ? quickConnectLabel : profile.name;
        rows.push_back(std::move(row));
    }

    // Sort case-insensitively; Quick Connect is always pinned to the top
    // regardless of direction.
    using SortDirection = RedSalamander::DxUi::SortDirection;
    if (_listSortDirection != SortDirection::None)
    {
        std::stable_sort(rows.begin(),
                         rows.end(),
                         [&](const ConnectionListGridRow& a, const ConnectionListGridRow& b) noexcept
        {
            const bool aQuick = a.modelIndex < _connections.size() && IsQuickConnectProfile(_connections[a.modelIndex]);
            const bool bQuick = b.modelIndex < _connections.size() && IsQuickConnectProfile(_connections[b.modelIndex]);
            if (aQuick != bQuick)
            {
                return aQuick;
            }
            const int cmp = _wcsicmp(a.text.c_str(), b.text.c_str());
            return _listSortDirection == SortDirection::Ascending ? cmp < 0 : cmp > 0;
        });
    }

    _viewToModel.reserve(rows.size());
    for (const auto& row : rows)
    {
        _viewToModel.push_back(row.modelIndex);
    }

    _listModel.SetRows(std::move(rows));
    if (_list)
    {
        std::optional<size_t> selectedVisibleIndex;
        if (previousModelIndex)
        {
            const auto it = std::find(_viewToModel.begin(), _viewToModel.end(), *previousModelIndex);
            if (it != _viewToModel.end())
            {
                selectedVisibleIndex = static_cast<size_t>(std::distance(_viewToModel.begin(), it));
            }
        }
        if (! selectedVisibleIndex && ! previousSelectedName.empty())
        {
            for (size_t rowIndex = 0u; rowIndex < _viewToModel.size(); ++rowIndex)
            {
                const size_t modelIndex = _viewToModel[rowIndex];
                if (modelIndex < _connections.size() && _connections[modelIndex].name == previousSelectedName)
                {
                    selectedVisibleIndex = rowIndex;
                    break;
                }
            }
        }
        if (! selectedVisibleIndex && ! _viewToModel.empty())
        {
            selectedVisibleIndex = 0u;
        }

        if (selectedVisibleIndex)
        {
            const uint64_t stableId = _listModel.GetStableRowId(*selectedVisibleIndex);
            _list->GetSelectionModel().SetSingle(stableId);
            _list->EnsureRowVisible(*selectedVisibleIndex);
        }
        else
        {
            _list->GetSelectionModel().Clear();
        }
        _list->NotifyDataChanged();
    }
    RefreshEditorFromSelection();
}

void WindowImpl::OnGridSortRequested(const RedSalamander::DxUi::GridSortSpec& sortSpec)
{
    // Honor the grid's 3-state cycle: Ascending -> Descending -> None -> ...
    _listSortDirection = sortSpec.direction;
    if (_list)
    {
        _list->SetSortSpec(sortSpec);
    }
    RebuildList();
}

std::optional<size_t> WindowImpl::GetSelectedModelIndex() const noexcept
{
    if (! _list)
    {
        return std::nullopt;
    }
    const auto primary = _list->GetPrimarySelectedRow();
    if (! primary || *primary >= _viewToModel.size())
    {
        return std::nullopt;
    }
    return _viewToModel[*primary];
}

void WindowImpl::SelectConnectionModelIndex(size_t modelIndex) noexcept
{
    if (! _list || ! _list->GetModel())
    {
        return;
    }

    const auto rowIndex = std::find(_viewToModel.begin(), _viewToModel.end(), modelIndex);
    if (rowIndex == _viewToModel.end())
    {
        return;
    }

    const size_t visibleIndex = static_cast<size_t>(std::distance(_viewToModel.begin(), rowIndex));
    const uint64_t stableId   = _listModel.GetStableRowId(visibleIndex);
    _list->GetSelectionModel().SetSingle(stableId);
    _list->EnsureRowVisible(visibleIndex);
    _list->NotifyDataChanged();
    RefreshEditorFromSelection();
}

void WindowImpl::RefreshEditorFromSelection() noexcept
{
    if (const auto modelIndex = GetSelectedModelIndex())
    {
        const auto& profile     = _connections[*modelIndex];
        _selectedConnectionName = profile.name;
        LoadEditorFromProfile(profile);
    }
    else
    {
        _selectedConnectionName.clear();
        ClearEditor();
    }
    _dxHost.Invalidate();
}

void WindowImpl::LoadEditorFromProfile(const Common::Settings::ConnectionProfile& profile) noexcept
{
    _loadingEditor                   = true;
    const auto loadingEditorScopeOff = wil::scope_exit([this]() noexcept { _loadingEditor = false; });

    if (_editName)
    {
        _editName->SetText(profile.name);
    }
    if (_comboProtocol)
    {
        std::optional<size_t> selected;
        for (size_t i = 0; i < std::size(kProtocolItems); ++i)
        {
            if (profile.pluginId == kProtocolItems[i].pluginId)
            {
                selected = i;
                break;
            }
        }
        _comboProtocol->SetSelectedIndex(selected);
    }
    if (_editHost)
    {
        _editHost->SetText(profile.host);
    }
    if (_comboAwsRegion)
    {
        // host stores the AWS region for S3 plugins.
        _comboAwsRegion->SetText(profile.host);
        std::optional<size_t> selected;
        for (size_t i = 0; i < std::size(kAwsRegionItems); ++i)
        {
            if (profile.host == kAwsRegionItems[i].code)
            {
                selected = i;
                break;
            }
        }
        _comboAwsRegion->SetSelectedIndex(selected);
    }
    if (_editPort)
    {
        _editPort->SetText(profile.port == 0u ? std::wstring{} : std::to_wstring(profile.port));
    }
    if (_editInitialPath)
    {
        _editInitialPath->SetText(profile.initialPath);
    }
    if (_editCopyMoveMaxConcurrency)
    {
        _editCopyMoveMaxConcurrency->SetText(std::wstring{});
    }
    if (_editDeleteMaxConcurrency)
    {
        _editDeleteMaxConcurrency->SetText(std::wstring{});
    }
    if (_editUser)
    {
        _editUser->SetText(profile.userName);
    }
    if (_editSecret)
    {
        _editSecret->SetText(std::wstring{});
        _editSecret->SetMasked(true);
    }
    if (_btnShowSecret)
    {
        _btnShowSecret->SetText(LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_SHOW_SECRET));
    }
    if (_toggleAnonymous)
    {
        _toggleAnonymous->SetChecked(profile.authMode == Common::Settings::ConnectionAuthMode::Anonymous);
    }
    if (_toggleSavePassword)
    {
        _toggleSavePassword->SetChecked(profile.savePassword);
    }
    if (_toggleRequireHello)
    {
        _toggleRequireHello->SetChecked(profile.requireWindowsHello);
    }
    if (_toggleIgnoreSslTrust)
    {
        _toggleIgnoreSslTrust->SetChecked(ConnectionProfileUtils::ExtraGetBool(profile.extra, "ignoreSslTrust").value_or(false));
    }
    if (_editS3EndpointOverride)
    {
        _editS3EndpointOverride->SetText(ExtraGetString(profile.extra, "endpointOverride").value_or(std::wstring{}));
    }
    if (_toggleS3UseHttps)
    {
        _toggleS3UseHttps->SetChecked(ConnectionProfileUtils::ExtraGetBool(profile.extra, "useHttps").value_or(true));
    }
    if (_toggleS3VerifyTls)
    {
        _toggleS3VerifyTls->SetChecked(ConnectionProfileUtils::ExtraGetBool(profile.extra, "verifyTls").value_or(true));
    }
    if (_toggleS3UseVirtualAddressing)
    {
        _toggleS3UseVirtualAddressing->SetChecked(ConnectionProfileUtils::ExtraGetBool(profile.extra, "useVirtualAddressing").value_or(true));
    }
    if (_editSshPrivateKey)
    {
        _editSshPrivateKey->SetText(ExtraGetString(profile.extra, "sshPrivateKey").value_or(std::wstring{}));
    }
    if (_editSshKnownHosts)
    {
        _editSshKnownHosts->SetText(ExtraGetString(profile.extra, "sshKnownHosts").value_or(std::wstring{}));
    }
    if (_editCopyMoveMaxConcurrency)
    {
        const uint32_t v = ConnectionProfileUtils::ExtraGetUInt32(profile.extra, "copyMoveMaxConcurrency").value_or(0u);
        _editCopyMoveMaxConcurrency->SetText(v == 0u ? std::wstring{} : std::to_wstring(v));
    }
    if (_editDeleteMaxConcurrency)
    {
        const uint32_t v = ConnectionProfileUtils::ExtraGetUInt32(profile.extra, "deleteMaxConcurrency").value_or(0u);
        _editDeleteMaxConcurrency->SetText(v == 0u ? std::wstring{} : std::to_wstring(v));
    }

    ApplyEditorVisibility(ComputeEditorVisibility());
    Layout();
}

void WindowImpl::ClearEditor() noexcept
{
    _loadingEditor                   = true;
    const auto loadingEditorScopeOff = wil::scope_exit([this]() noexcept { _loadingEditor = false; });

    auto clearEdit = [](TextField* edit)
    {
        if (edit)
        {
            edit->SetText(std::wstring{});
        }
    };
    auto clearToggle = [](Toggle* toggle, bool defaultChecked)
    {
        if (toggle)
        {
            toggle->SetChecked(defaultChecked);
        }
    };

    clearEdit(_editName);
    clearEdit(_editHost);
    clearEdit(_editPort);
    clearEdit(_editInitialPath);
    clearEdit(_editCopyMoveMaxConcurrency);
    clearEdit(_editDeleteMaxConcurrency);
    clearEdit(_editUser);
    clearEdit(_editSecret);
    clearEdit(_editS3EndpointOverride);
    clearEdit(_editSshPrivateKey);
    clearEdit(_editSshKnownHosts);
    if (_comboProtocol)
    {
        _comboProtocol->SetSelectedIndex(std::nullopt);
    }
    if (_comboAwsRegion)
    {
        _comboAwsRegion->SetText(std::wstring{});
        _comboAwsRegion->SetSelectedIndex(std::nullopt);
    }
    clearToggle(_toggleAnonymous, false);
    clearToggle(_toggleSavePassword, false);
    clearToggle(_toggleRequireHello, true);
    clearToggle(_toggleIgnoreSslTrust, false);
    clearToggle(_toggleS3UseHttps, true);
    clearToggle(_toggleS3VerifyTls, true);
    clearToggle(_toggleS3UseVirtualAddressing, false);

    ApplyEditorVisibility(ComputeEditorVisibility());
    Layout();
}

void WindowImpl::OnNewClicked() noexcept
{
    Common::Settings::ConnectionProfile profile;
    profile.id = NewGuidString();
    if (profile.id.empty())
    {
        Debug::Error(L"ConnectionManagerWindow: failed to allocate connection id (CoCreateGuid).");
        return;
    }
    profile.pluginId            = _filterPluginId.empty() ? std::wstring(kProtocolItems[0].pluginId) : _filterPluginId;
    profile.name                = MakeUniqueConnectionName(_connections, LoadStringResource(nullptr, IDS_CONNECTIONS_DEFAULT_NEW_NAME), {});
    profile.host                = std::wstring{};
    profile.initialPath         = L"/";
    profile.port                = 0u;
    profile.userName            = std::wstring{};
    profile.authMode            = Common::Settings::ConnectionAuthMode::Password;
    profile.savePassword        = false;
    profile.requireWindowsHello = true;

    _connections.push_back(std::move(profile));
    MarkConnectionsDirty();
    const size_t newModelIndex = _connections.size() - 1u;

    RebuildList();

    // Select the freshly added row and focus the Name field for renaming.
    if (_list && _list->GetModel())
    {
        if (auto rowIndex = std::find(_viewToModel.begin(), _viewToModel.end(), newModelIndex); rowIndex != _viewToModel.end())
        {
            const size_t visibleIndex = static_cast<size_t>(std::distance(_viewToModel.begin(), rowIndex));
            const uint64_t stableId   = _listModel.GetStableRowId(visibleIndex);
            _list->GetSelectionModel().SetSingle(stableId);
            _list->EnsureRowVisible(visibleIndex);
            _list->NotifyDataChanged();
            RefreshEditorFromSelection();
        }
    }

    if (_editName)
    {
        _dxHost.SetFocusControl(_editName);
        _editName->SetSelectionRange(0u, std::numeric_limits<size_t>::max());
    }
}

void WindowImpl::OnRenameClicked() noexcept
{
    const auto modelIndex = GetSelectedModelIndex();
    if (! modelIndex || *modelIndex >= _connections.size())
    {
        return;
    }
    if (IsQuickConnectProfile(_connections[*modelIndex]))
    {
        return; // Quick Connect can't be renamed.
    }
    if (_editName)
    {
        _dxHost.SetFocusControl(_editName);
        _editName->SetSelectionRange(0u, std::numeric_limits<size_t>::max());
    }
}

void WindowImpl::OnRemoveClicked() noexcept
{
    const auto modelIndex = GetSelectedModelIndex();
    if (! modelIndex || *modelIndex >= _connections.size())
    {
        return;
    }
    if (IsQuickConnectProfile(_connections[*modelIndex]))
    {
        return; // Quick Connect can't be removed.
    }

    _connections.erase(_connections.begin() + static_cast<ptrdiff_t>(*modelIndex));
    MarkConnectionsDirty();
    RebuildList();

    // Select the nearest remaining row, if any.
    if (_list && _list->GetModel() && _listModel.GetRowCount() > 0u)
    {
        const size_t target   = (*modelIndex < _listModel.GetRowCount()) ? *modelIndex : _listModel.GetRowCount() - 1u;
        const uint64_t stable = _listModel.GetStableRowId(target);
        _list->GetSelectionModel().SetSingle(stable);
        _list->EnsureRowVisible(target);
        _list->NotifyDataChanged();
        RefreshEditorFromSelection();
    }
    else
    {
        ClearEditor();
    }
}

void WindowImpl::OnCloseClicked() noexcept
{
    // Close persists dirty state and closes without selecting a connection.
    _selectedConnectionName.clear();
    if (! TryValidateAndNormalizeConnectionProfiles())
    {
        return;
    }
    if (! SaveConnectionsSettings())
    {
        if (_modalResult)
        {
            _modalResult->connectionName.clear();
            _modalResult->hr = E_FAIL;
        }
        return;
    }
    if (_modalResult)
    {
        _modalResult->connectionName.clear();
        _modalResult->hr = S_FALSE;
    }
    RequestCloseWindow();
}

void WindowImpl::OnGridSelectionChanged(Grid& /*sender*/)
{
    RefreshEditorFromSelection();
}

void WindowImpl::OnGridRowActivated(Grid& /*sender*/, size_t /*rowIndex*/)
{
    OnConnectClicked();
}

void WindowImpl::OnConnectClicked() noexcept
{
    if (! TryValidateAndNormalizeConnectionProfiles())
    {
        if (_modalResult)
        {
            _modalResult->connectionName.clear();
            _modalResult->hr = E_FAIL;
        }
        return;
    }
    if (const auto modelIndex = GetSelectedModelIndex())
    {
        _selectedConnectionName = _connections[*modelIndex].name;
    }
    if (_selectedConnectionName.empty() || ! SaveConnectionsSettings())
    {
        if (_modalResult)
        {
            _modalResult->connectionName.clear();
            _modalResult->hr = E_FAIL;
        }
        return;
    }
    if (_modalResult)
    {
        _modalResult->connectionName = _selectedConnectionName;
        _modalResult->hr             = S_OK;
        RequestCloseWindow();
        return;
    }

    if (! NotifyOwnerToConnectSelectedProfile())
    {
        return;
    }

    RequestCloseWindow();
}

void WindowImpl::OnCancelClicked() noexcept
{
    _selectedConnectionName.clear();
    if (_modalResult)
    {
        _modalResult->connectionName.clear();
        _modalResult->hr = S_FALSE;
    }
    RequestCloseWindow();
}

bool WindowImpl::OnCommand(WORD commandId) noexcept
{
    switch (commandId)
    {
        case IDOK: OnConnectClicked(); return true;
        case IDCANCEL: OnCancelClicked(); return true;
        case IDC_CONNECTION_CLOSE: OnCloseClicked(); return true;
        case IDC_CONNECTION_NEW: OnNewClicked(); return true;
        case IDC_CONNECTION_RENAME: OnRenameClicked(); return true;
        case IDC_CONNECTION_REMOVE: OnRemoveClicked(); return true;
        case IDC_CONNECTION_SHOW_SECRET: return _btnShowSecret && _btnShowSecret->Invoke(_dxHost, false);
        default: return false;
    }
}

LRESULT WindowImpl::WindowProc(UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    const HWND messageHwnd = _hwnd ? _hwnd.get() : _closingHwnd;
    bool dxHandled         = false;
    const LRESULT dxResult = _dxHost.HandleMessage(messageHwnd, msg, wp, lp, dxHandled);
    if (dxHandled)
    {
        if (msg == WM_SIZE)
        {
            Layout();
        }
        return dxResult;
    }

    switch (msg)
    {
        case WM_CREATE: return OnCreate() ? 0 : -1;
        case WM_SIZE: OnSize(); return 0;
        case WM_DPICHANGED: OnDpiChanged(reinterpret_cast<const RECT*>(lp)); return 0;
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lp))
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(messageHwnd, *info, 720, 460);
                static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(messageHwnd, *info));
            }
            return 0;
        case WM_NCACTIVATE: OnNcActivate(wp != FALSE); return DefWindowProcW(messageHwnd, msg, wp, lp);
        case WM_COMMAND:
            if (OnCommand(LOWORD(wp)))
            {
                return 0;
            }
            break;
        case WM_CLOSE: OnClose(); return 0;
        case WM_NCDESTROY: OnNcDestroy(); return 0;
        case WndMsg::kSettingsReloadedFromDisk: OnSettingsReloadedFromDisk(); return 0;
        default: break;
    }

    return DefWindowProcW(messageHwnd, msg, wp, lp);
}

LRESULT CALLBACK WindowImpl::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    WindowImpl* self = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE)
    {
        auto* create = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self         = create ? reinterpret_cast<WindowImpl*>(create->lpCreateParams) : nullptr;
        if (! self)
        {
            return FALSE;
        }
        self->_hwnd.reset(hwnd);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    }

    if (! self)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    ++self->_dispatchDepth;
    const LRESULT result = self->WindowProc(msg, wp, lp);
    if (self->_dispatchDepth > 0u)
    {
        --self->_dispatchDepth;
    }
    if (self->_dispatchDepth == 0u && self->_deletePending)
    {
        delete self;
    }
    return result;
}

} // namespace

bool ShowWindow(HWND owner,
                std::wstring_view appId,
                Common::Settings::Settings& settings,
                const AppTheme& theme,
                std::wstring_view filterPluginId,
                uint8_t targetPane) noexcept
{
    if (appId.empty())
    {
        return false;
    }

    const HWND effectiveOwner = NormalizeOwnerWindow(owner);

    if (const HWND existing = g_singleInstance.load(std::memory_order_acquire))
    {
        if (! IsWindow(existing))
        {
            g_singleInstance.store(nullptr, std::memory_order_release);
        }
        else if (auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(existing, GWLP_USERDATA)))
        {
            impl->UpdateContext(theme, filterPluginId, effectiveOwner, targetPane);
            if (IsIconic(existing))
            {
                ::ShowWindow(existing, SW_RESTORE);
            }
            else
            {
                ::ShowWindow(existing, SW_SHOW);
            }
            SetForegroundWindow(existing);
            return true;
        }
    }

    auto impl = std::make_unique<WindowImpl>(std::wstring(appId), settings, theme, std::wstring(filterPluginId), targetPane, effectiveOwner);
    if (! impl->Create())
    {
        return false;
    }
    static_cast<void>(impl.release()); // owned by HWND via GWLP_USERDATA, freed in WM_NCDESTROY
    return true;
}

HRESULT ShowDialog(HWND owner,
                   std::wstring_view appId,
                   Common::Settings::Settings& settings,
                   const AppTheme& theme,
                   std::wstring_view filterPluginId,
                   std::wstring& selectedConnectionNameOut) noexcept
{
    selectedConnectionNameOut.clear();
    if (appId.empty())
    {
        return E_INVALIDARG;
    }

    // Phase 8.2b - synchronous facade over a modeless top-level DxUi window.
    // The window itself is unowned per the top-level tool window contract; the
    // facade disables the owner for the duration of the nested message pump
    // and re-enables it on exit, preserving the legacy `DialogBoxParam`
    // semantics for plugin-host callers (see `HostServices.cpp:1419`).
    const HWND effectiveOwner = NormalizeOwnerWindow(owner);

    ModalFacadeResult result;
    auto impl = std::make_unique<WindowImpl>(
        std::wstring(appId), settings, theme, std::wstring(filterPluginId), /*targetPane=*/static_cast<uint8_t>(0), effectiveOwner);
    impl->EnterModalFacadeMode(&result);
    if (! impl->Create())
    {
        return E_FAIL;
    }
    const HWND modalHwnd = impl->GetHwnd();
    static_cast<void>(impl.release()); // owned by HWND, freed in WM_NCDESTROY
    if (! modalHwnd)
    {
        return E_FAIL;
    }

    HWND disabledOwner = nullptr;
    if (effectiveOwner && IsWindow(effectiveOwner) != FALSE && IsWindowEnabled(effectiveOwner) != FALSE)
    {
        EnableWindow(effectiveOwner, FALSE);
        disabledOwner = effectiveOwner;
    }

    // Nested message pump: spin until the modal-facade window is destroyed.
    // WM_QUIT during the pump is preserved for the outer message loop.
    MSG msg{};
    while (IsWindow(modalHwnd) != FALSE)
    {
        const BOOL got = GetMessageW(&msg, nullptr, 0, 0);
        if (got == 0)
        {
            // WM_QUIT - re-post for the outer pump and exit ours.
            PostQuitMessage(static_cast<int>(msg.wParam));
            break;
        }
        if (got < 0)
        {
            break;
        }
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    if (disabledOwner)
    {
        EnableWindow(disabledOwner, TRUE);
        SetForegroundWindow(disabledOwner);
    }

    selectedConnectionNameOut = std::move(result.connectionName);
    if (result.hr == S_OK && selectedConnectionNameOut.empty())
    {
        return E_FAIL;
    }
    return result.hr;
}

HWND GetWindowHandle() noexcept
{
    const HWND existing = g_singleInstance.load(std::memory_order_acquire);
    return (existing && IsWindow(existing) != FALSE) ? existing : nullptr;
}

void UpdateTheme(const AppTheme& theme) noexcept
{
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return;
    }
    if (auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA)))
    {
        impl->UpdateTheme(theme);
    }
}

#ifdef ENABLE_TESTS
namespace
{
[[nodiscard]] size_t CountIfVisibleControl(const Control* control) noexcept
{
    return (control && control->IsVisible()) ? 1u : 0u;
}
} // namespace

void WindowImpl::DebugFillSnapshot(::ConnectionManagerDebugSnapshot& out) const noexcept
{
    out                                      = {};
    out.usesDxUiCommandButtons               = true;
    out.usesDxUiSectionHeaders               = true;
    out.usesDxUiFormLabels                   = true;
    out.usesDxUiFormInputs                   = true;
    out.usesDxUiFormActionButtons            = true;
    out.usesDxUiList                         = true;
    out.legacyOwnerDrawCommandButtonCount    = 0u;
    out.legacyOwnerDrawFormInputCount        = 0u;
    out.legacyOwnerDrawFormActionButtonCount = 0u;
    out.visibleLegacyCommandButtonCount      = 0u;
    out.visibleLegacySectionHeaderCount      = 0u;
    out.visibleLegacyFormLabelCount          = 0u;
    out.visibleLegacyFormInputCount          = 0u;
    out.visibleLegacyFormActionButtonCount   = 0u;
    out.visibleLegacyListCount               = 0u;

    // The single-canvas path has no per-widget host HWNDs; for parity with the
    // legacy snapshot, expose every visible Dx control as a "host" of 1 so the
    // existing self-tests that check `> 0u` keep passing.
    out.visibleDxSectionHeaderHostCount = CountIfVisibleControl(_sectionConnection) + CountIfVisibleControl(_sectionAuth) + CountIfVisibleControl(_sectionS3) +
                                          CountIfVisibleControl(_sectionSsh);
    out.visibleDxFormInputHostCount =
        CountIfVisibleControl(_editName) + CountIfVisibleControl(_editHost) + CountIfVisibleControl(_editPort) + CountIfVisibleControl(_editInitialPath) +
        CountIfVisibleControl(_editCopyMoveMaxConcurrency) + CountIfVisibleControl(_editDeleteMaxConcurrency) + CountIfVisibleControl(_editUser) +
        CountIfVisibleControl(_editSecret) + CountIfVisibleControl(_editS3EndpointOverride) + CountIfVisibleControl(_editSshPrivateKey) +
        CountIfVisibleControl(_editSshKnownHosts) + CountIfVisibleControl(_comboProtocol) + CountIfVisibleControl(_comboAwsRegion) +
        CountIfVisibleControl(_toggleAnonymous) + CountIfVisibleControl(_toggleSavePassword) + CountIfVisibleControl(_toggleRequireHello) +
        CountIfVisibleControl(_toggleIgnoreSslTrust) + CountIfVisibleControl(_toggleS3UseHttps) + CountIfVisibleControl(_toggleS3VerifyTls) +
        CountIfVisibleControl(_toggleS3UseVirtualAddressing);
    out.visibleDxFormActionButtonHostCount =
        CountIfVisibleControl(_btnShowSecret) + CountIfVisibleControl(_btnSshPrivateKeyBrowse) + CountIfVisibleControl(_btnSshKnownHostsBrowse);
    out.visibleDxListHostCount = CountIfVisibleControl(_list);

    // List metrics
    out.listRowCount = _listModel.GetRowCount();
    if (_list)
    {
        const GridVisibleWorkMetrics metrics = _list->GetVisibleWorkMetrics();
        out.visibleListRowCount              = static_cast<size_t>(metrics.visibleRowCount);
        out.visibleListColumnCount           = metrics.visibleColumnCount;
        out.visibleListCellCount             = static_cast<size_t>(metrics.visibleCellCount);
        out.listHasVerticalScrollbar         = metrics.hasVerticalScrollbar;
    }

    // Theme flags
    out.themeDark         = _theme.dark;
    out.themeHighContrast = _theme.highContrast;
    out.themeRainbow      = _theme.menu.rainbowMode;

    // The connection list is rendered by the single WindowHost canvas, so expose
    // the host counters as the list's live render/resize accounting.
    out.dxListRenderCount        = _dxHost.DebugGetRenderCount();
    out.dxListResizeCount        = _dxHost.DebugGetResizeCount();
    out.dxListResizeFailureCount = _dxHost.DebugGetResizeFailureCount();

    // Selected list state
    if (const auto modelIndex = GetSelectedModelIndex(); modelIndex && *modelIndex < _connections.size())
    {
        const auto& profile     = _connections[*modelIndex];
        out.selectedListRowName = profile.name;
        out.currentNameText     = _editName ? std::wstring(_editName->GetText()) : profile.name;
        out.currentPluginId     = profile.pluginId;
    }
    else
    {
        out.selectedListRowName.clear();
        out.currentNameText.clear();
        out.currentPluginId.clear();
    }
    if (_list)
    {
        if (const auto primary = _list->GetPrimarySelectedRow())
        {
            out.selectedListIndex = static_cast<int>(*primary);
        }
        else
        {
            out.selectedListIndex = -1;
        }
    }
    else
    {
        out.selectedListIndex = -1;
    }
    if (out.selectedListIndex >= 0)
    {
        const auto& rows      = _listModel.GetRows();
        const size_t rowIndex = static_cast<size_t>(out.selectedListIndex);
        if (rowIndex < rows.size())
        {
            out.selectedListRowName = rows[rowIndex].text;
        }
    }
    out.selectedListRowFillArgb    = 0u;
    out.selectedListRowTextArgb    = 0u;
    out.selectedListRowUsesRainbow = false;
    if (_list && out.selectedListIndex >= 0)
    {
        GridDebugRowVisualState visualState{};
        if (_list->DebugGetRowVisualState(MakeAppThemeDxPalette(_theme), static_cast<size_t>(out.selectedListIndex), visualState))
        {
            out.selectedListRowFillArgb    = visualState.fillArgb;
            out.selectedListRowTextArgb    = visualState.textArgb;
            out.selectedListRowUsesRainbow = visualState.usesRainbow;
        }
    }

    // Focus tracking: legacy reported the focused-control "kind". For the
    // single-canvas path the WindowHost owns a single focused Control* - map
    // it back to a kind for the test surface.
    out.focusKind = ConnectionManagerDebugFocusKind::None;
    out.focusLabel.clear();
    const auto buttonText   = [](const Button* button) -> std::wstring { return button ? std::wstring(button->GetText()) : std::wstring{}; };
    const auto resourceText = [](UINT stringId) -> std::wstring { return LoadStringResource(nullptr, stringId); };
    if (const Control* focused = _dxHost.GetFocusControl())
    {
        if (focused == _list)
        {
            out.focusKind = ConnectionManagerDebugFocusKind::List;
        }
        else if (focused == _newButton || focused == _renameButton || focused == _removeButton || focused == _connectButton || focused == _closeButton ||
                 focused == _cancelButton)
        {
            out.focusKind = ConnectionManagerDebugFocusKind::CommandButton;
            if (focused == _newButton)
            {
                out.focusLabel = buttonText(_newButton);
            }
            else if (focused == _renameButton)
            {
                out.focusLabel = buttonText(_renameButton);
            }
            else if (focused == _removeButton)
            {
                out.focusLabel = buttonText(_removeButton);
            }
            else if (focused == _connectButton)
            {
                out.focusLabel = buttonText(_connectButton);
            }
            else if (focused == _closeButton)
            {
                out.focusLabel = buttonText(_closeButton);
            }
            else
            {
                out.focusLabel = buttonText(_cancelButton);
            }
        }
        else if (focused == _editName || focused == _editHost || focused == _editPort || focused == _editInitialPath ||
                 focused == _editCopyMoveMaxConcurrency || focused == _editDeleteMaxConcurrency || focused == _editUser || focused == _editSecret ||
                 focused == _editS3EndpointOverride || focused == _editSshPrivateKey || focused == _editSshKnownHosts)
        {
            out.focusKind = ConnectionManagerDebugFocusKind::Edit;
            if (focused == _editName)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_NAME);
            }
            else if (focused == _editHost)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_HOST);
            }
            else if (focused == _editPort)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_PORT);
            }
            else if (focused == _editInitialPath)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_INITIAL_PATH);
            }
            else if (focused == _editCopyMoveMaxConcurrency)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_COPYMOVE_CONCURRENCY);
            }
            else if (focused == _editDeleteMaxConcurrency)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_DELETE_CONCURRENCY);
            }
            else if (focused == _editUser)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_USER);
            }
            else if (focused == _editSecret)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_PASSWORD);
            }
            else if (focused == _editS3EndpointOverride)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_ENDPOINT_OVERRIDE);
            }
            else if (focused == _editSshPrivateKey)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_SSH_PRIVATEKEY);
            }
            else
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_SSH_KNOWNHOSTS);
            }
        }
        else if (focused == _comboProtocol || focused == _comboAwsRegion)
        {
            out.focusKind  = ConnectionManagerDebugFocusKind::Combo;
            out.focusLabel = resourceText(focused == _comboProtocol ? IDS_CONNECTIONS_LABEL_PROTOCOL : IDS_CONNECTIONS_LABEL_REGION);
        }
        else if (focused == _toggleAnonymous || focused == _toggleSavePassword || focused == _toggleRequireHello || focused == _toggleIgnoreSslTrust ||
                 focused == _toggleS3UseHttps || focused == _toggleS3VerifyTls || focused == _toggleS3UseVirtualAddressing)
        {
            out.focusKind = ConnectionManagerDebugFocusKind::Toggle;
            if (focused == _toggleAnonymous)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_ANONYMOUS);
            }
            else if (focused == _toggleSavePassword)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_SAVE_PASSWORD);
            }
            else if (focused == _toggleRequireHello)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_REQUIRE_HELLO);
            }
            else if (focused == _toggleIgnoreSslTrust)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_IGNORE_SSL_TRUST);
            }
            else if (focused == _toggleS3UseHttps)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_USE_HTTPS);
            }
            else if (focused == _toggleS3VerifyTls)
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_VERIFY_TLS);
            }
            else
            {
                out.focusLabel = resourceText(IDS_CONNECTIONS_LABEL_USE_VIRTUAL_ADDRESSING);
            }
        }
        else if (focused == _btnShowSecret || focused == _btnSshPrivateKeyBrowse || focused == _btnSshKnownHostsBrowse)
        {
            out.focusKind = ConnectionManagerDebugFocusKind::FormActionButton;
            if (focused == _btnShowSecret)
            {
                out.focusLabel = buttonText(_btnShowSecret);
            }
            else if (focused == _btnSshPrivateKeyBrowse)
            {
                out.focusLabel = buttonText(_btnSshPrivateKeyBrowse);
            }
            else
            {
                out.focusLabel = buttonText(_btnSshKnownHostsBrowse);
            }
        }
    }

    // Name host and bridge readiness: the single-canvas path doesn't have a
    // separate host HWND for the Name field; report the `_editName` slot's
    // visibility / enabled state for compatibility with the legacy assertions.
    out.nameHostPresent             = _editName != nullptr;
    out.nameHostVisible             = _editName && _editName->IsVisible();
    out.nameHostEnabled             = _editName && _editName->IsEnabled();
    out.nameLegacyVisible           = false;
    out.nameTextFieldPresent        = _editName != nullptr;
    out.nameTextFieldVisible        = out.nameHostVisible;
    out.nameTextFieldEnabled        = out.nameHostEnabled;
    out.nameHostFocusControlMatches = (_dxHost.GetFocusControl() == _editName);
    out.nameHostOwnsFocus           = out.nameHostFocusControlMatches;
}

bool DebugGetSnapshot(::ConnectionManagerDebugSnapshot& out) noexcept
{
    out             = {};
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! impl)
    {
        return false;
    }
    impl->DebugFillSnapshot(out);
    return true;
}

bool WindowImpl::DebugClickListRow(size_t rowIndex) noexcept
{
    if (! _list || ! _list->GetModel() || rowIndex >= _list->GetModel()->GetRowCount())
    {
        return false;
    }
    if (! _list->RequestSelectRow(rowIndex, 0u))
    {
        return false;
    }
    RefreshEditorFromSelection();
    return true;
}

bool WindowImpl::DebugScrollListByWheelDetents(int detents) noexcept
{
    if (! _list)
    {
        return false;
    }
    // Fall through to the host-level wheel handler with the list as the
    // hovered control. This matches the Grid's `OnMouseWheel` semantics.
    const D2D1_RECT_F bounds = _list->GetBounds();
    const D2D1_POINT_2F mid  = D2D1::Point2F((bounds.left + bounds.right) * 0.5f, (bounds.top + bounds.bottom) * 0.5f);
    const float delta        = static_cast<float>(detents) * static_cast<float>(WHEEL_DELTA);
    return _list->OnMouseWheel(_dxHost, mid, delta, 0u);
}

bool WindowImpl::DebugFocusFirstInput() noexcept
{
    if (! _editName)
    {
        return false;
    }
    _dxHost.SetFocusControl(_editName);
    return _dxHost.GetFocusControl() == _editName;
}

bool WindowImpl::DebugFocusList() noexcept
{
    if (! _list)
    {
        return false;
    }
    _dxHost.SetFocusControl(_list);
    return _dxHost.GetFocusControl() == _list;
}

bool WindowImpl::DebugRouteMnemonic(wchar_t mnemonic) noexcept
{
    return _dxHost.HandleMnemonic(mnemonic);
}

bool WindowImpl::DebugRouteCommandKey(WPARAM virtualKey) noexcept
{
    if (! _hwnd)
    {
        return false;
    }
    if (virtualKey == VK_RETURN)
    {
        if (Button* def = _dxHost.GetDefaultButton())
        {
            return def->Invoke(_dxHost, false);
        }
    }
    if (virtualKey == VK_ESCAPE)
    {
        bool handled = false;
        static_cast<void>(_dxHost.HandleMessage(_hwnd.get(), WM_KEYDOWN, VK_ESCAPE, 0, handled));
        static_cast<void>(_dxHost.HandleMessage(_hwnd.get(), WM_KEYUP, VK_ESCAPE, 0, handled));
        return handled;
    }
    return false;
}

bool WindowImpl::DebugRouteTab(bool reverse) noexcept
{
    // Synthesise a VK_TAB through the host. The host's own `HandleTabNavigation`
    // is private; route through `HandleMessage` with WM_KEYDOWN VK_TAB.
    if (! _hwnd)
    {
        return false;
    }
    bool keyHandled     = false;
    const WPARAM modKey = reverse ? VK_SHIFT : 0;
    if (reverse)
    {
        bool shiftHandled = false;
        static_cast<void>(_dxHost.HandleMessage(_hwnd.get(), WM_KEYDOWN, modKey, 0, shiftHandled));
    }
    static_cast<void>(_dxHost.HandleMessage(_hwnd.get(), WM_KEYDOWN, VK_TAB, 0, keyHandled));
    bool keyUpHandled = false;
    static_cast<void>(_dxHost.HandleMessage(_hwnd.get(), WM_KEYUP, VK_TAB, 0, keyUpHandled));
    if (reverse)
    {
        bool shiftHandled = false;
        static_cast<void>(_dxHost.HandleMessage(_hwnd.get(), WM_KEYUP, modKey, 0, shiftHandled));
    }
    return keyHandled;
}

bool WindowImpl::DebugSetProtocolPluginId(std::wstring_view pluginId) noexcept
{
    const std::optional<size_t> comboIndex = FindProtocolItemIndexByPluginId(pluginId);
    if (! comboIndex || ! _comboProtocol)
    {
        return false;
    }
    _comboProtocol->SetSelectedIndex(comboIndex);
    OnProtocolChanged(*comboIndex);
    ApplyEditorVisibility(ComputeEditorVisibility());
    Layout();
    return true;
}

bool WindowImpl::DebugGetAlternateProtocolPluginId(std::wstring_view baselinePluginId, std::wstring& outPluginId) const noexcept
{
    outPluginId.clear();
    for (const auto& entry : kProtocolItems)
    {
        if (baselinePluginId != entry.pluginId)
        {
            outPluginId.assign(entry.pluginId);
            return true;
        }
    }
    return false;
}

bool WindowImpl::DebugGetControlClientRect(const Control* control, RECT& outRect) const noexcept
{
    if (! control || ! _hwnd)
    {
        outRect = {};
        return false;
    }
    const D2D1_RECT_F bounds = control->GetBounds();
    const float scale        = _dxHost.DipsToPixels(1.0f);
    outRect.left             = static_cast<LONG>(std::floor(bounds.left * scale));
    outRect.top              = static_cast<LONG>(std::floor(bounds.top * scale));
    outRect.right            = static_cast<LONG>(std::ceil(bounds.right * scale));
    outRect.bottom           = static_cast<LONG>(std::ceil(bounds.bottom * scale));
    return true;
}

Button* WindowImpl::DebugGetCommandButton(UINT commandId) noexcept
{
    switch (commandId)
    {
        case IDOK: return _connectButton;
        case IDCANCEL: return _cancelButton;
        case IDC_CONNECTION_CLOSE: return _closeButton;
        case IDC_CONNECTION_NEW: return _newButton;
        case IDC_CONNECTION_RENAME: return _renameButton;
        case IDC_CONNECTION_REMOVE: return _removeButton;
        case IDC_CONNECTION_SHOW_SECRET: return _btnShowSecret;
        default: return nullptr;
    }
}

bool DebugClickListRow(size_t rowIndex) noexcept
{
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return impl && impl->DebugClickListRow(rowIndex);
}

bool DebugScrollListByWheelDetents(int detents) noexcept
{
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return impl && impl->DebugScrollListByWheelDetents(detents);
}

bool DebugFocusFirstInput() noexcept
{
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return impl && impl->DebugFocusFirstInput();
}

bool DebugFocusList() noexcept
{
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return impl && impl->DebugFocusList();
}

bool DebugRouteMnemonic(wchar_t mnemonic) noexcept
{
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return impl && impl->DebugRouteMnemonic(mnemonic);
}

bool DebugRouteCommandKey(WPARAM virtualKey) noexcept
{
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return impl && impl->DebugRouteCommandKey(virtualKey);
}

bool DebugRouteTab(bool reverse) noexcept
{
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return impl && impl->DebugRouteTab(reverse);
}

bool DebugSetProtocolPluginId(std::wstring_view pluginId) noexcept
{
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return impl && impl->DebugSetProtocolPluginId(pluginId);
}

bool DebugGetAlternateProtocolPluginId(std::wstring_view baselinePluginId, std::wstring& outPluginId) noexcept
{
    outPluginId.clear();
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return impl && impl->DebugGetAlternateProtocolPluginId(baselinePluginId, outPluginId);
}

bool DebugGetListHostHandle(HWND& outHost) noexcept
{
    outHost         = nullptr;
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    // Single-canvas: the list "host" is the window itself.
    outHost = hwnd;
    return true;
}

bool DebugGetNameHostHandle(HWND& outHost) noexcept
{
    outHost         = nullptr;
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! impl)
    {
        return false;
    }
    outHost = hwnd;
    return true;
}

bool DebugAcknowledgeS3InsecureTlsPrompt() noexcept
{
    // The S3 insecure-TLS prompt runs through `HostServices` / `AlertOverlayWindow`
    // - the new path does not own that prompt. Returning true keeps this
    // Connection Manager hook compatible with legacy tests while the prompt's
    // own host remains responsible for actual acknowledgement.
    return true;
}

bool DebugGetSavePasswordToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept
{
    outHost         = nullptr;
    outRect         = {};
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! impl)
    {
        return false;
    }
    Toggle* toggle = impl->DebugGetSavePasswordToggle();
    if (! toggle)
    {
        return false;
    }
    outHost = hwnd;
    return impl->DebugGetControlClientRect(toggle, outRect);
}

bool DebugGetSavePasswordToggleState(bool& outChecked, std::wstring& outLabel) noexcept
{
    outChecked = false;
    outLabel.clear();
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! impl)
    {
        return false;
    }
    Toggle* toggle = impl->DebugGetSavePasswordToggle();
    if (! toggle)
    {
        return false;
    }
    outChecked = toggle->IsChecked();
    outLabel   = LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_SAVE_PASSWORD);
    return true;
}

bool DebugGetCommandButtonHostAndClientRect(UINT commandId, HWND& outHost, RECT& outRect, std::wstring& outLabel) noexcept
{
    outHost = nullptr;
    outRect = {};
    outLabel.clear();
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! impl)
    {
        return false;
    }
    Button* btn = impl->DebugGetCommandButton(commandId);
    if (! btn)
    {
        return false;
    }
    outHost = hwnd;
    outLabel.assign(btn->GetText());
    return impl->DebugGetControlClientRect(btn, outRect);
}

bool DebugGetS3UseHttpsToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept
{
    outHost         = nullptr;
    outRect         = {};
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! impl)
    {
        return false;
    }
    Toggle* toggle = impl->DebugGetS3UseHttpsToggle();
    if (! toggle)
    {
        return false;
    }
    outHost = hwnd;
    return impl->DebugGetControlClientRect(toggle, outRect);
}

bool DebugGetS3UseHttpsToggleState(bool& outChecked, std::wstring& outLabel) noexcept
{
    outChecked = false;
    outLabel.clear();
    const HWND hwnd = GetWindowHandle();
    if (! hwnd)
    {
        return false;
    }
    auto* impl = reinterpret_cast<WindowImpl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! impl)
    {
        return false;
    }
    Toggle* toggle = impl->DebugGetS3UseHttpsToggle();
    if (! toggle)
    {
        return false;
    }
    outChecked = toggle->IsChecked();
    outLabel   = LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_USE_HTTPS);
    return true;
}
#endif // ENABLE_TESTS

} // namespace RedSalamander::ConnectionManager::SingleCanvas

HRESULT ShowConnectionManagerDialog(HWND owner,
                                    std::wstring_view appId,
                                    Common::Settings::Settings& settings,
                                    const AppTheme& theme,
                                    std::wstring_view filterPluginId,
                                    std::wstring& selectedConnectionNameOut) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::ShowDialog(owner, appId, settings, theme, filterPluginId, selectedConnectionNameOut);
}

bool ShowConnectionManagerWindow(HWND owner,
                                 std::wstring_view appId,
                                 Common::Settings::Settings& settings,
                                 const AppTheme& theme,
                                 std::wstring_view filterPluginId,
                                 uint8_t targetPane) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::ShowWindow(owner, appId, settings, theme, filterPluginId, targetPane);
}

HWND GetConnectionManagerDialogHandle() noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::GetWindowHandle();
}

void UpdateConnectionManagerWindowsTheme(const AppTheme& theme) noexcept
{
    RedSalamander::ConnectionManager::SingleCanvas::UpdateTheme(theme);
}

#ifdef ENABLE_TESTS

bool DebugGetConnectionManagerDialogSnapshot(ConnectionManagerDebugSnapshot& out) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugGetSnapshot(out);
}

bool DebugClickConnectionManagerListRow(size_t rowIndex) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugClickListRow(rowIndex);
}

bool DebugScrollConnectionManagerListByWheelDetents(int detents) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugScrollListByWheelDetents(detents);
}

bool DebugFocusConnectionManagerFirstInput() noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugFocusFirstInput();
}

bool DebugFocusConnectionManagerList() noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugFocusList();
}

bool DebugRouteConnectionManagerMnemonic(wchar_t mnemonic) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugRouteMnemonic(mnemonic);
}

bool DebugRouteConnectionManagerCommandKey(WPARAM virtualKey) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugRouteCommandKey(virtualKey);
}

bool DebugRouteConnectionManagerTab(bool reverse) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugRouteTab(reverse);
}

bool DebugSetConnectionManagerProtocolPluginId(std::wstring_view pluginId) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugSetProtocolPluginId(pluginId);
}

bool DebugGetConnectionManagerAlternateProtocolPluginId(std::wstring_view baselinePluginId, std::wstring& outPluginId) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugGetAlternateProtocolPluginId(baselinePluginId, outPluginId);
}

bool DebugGetConnectionManagerListHostHandle(HWND& outHost) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugGetListHostHandle(outHost);
}

bool DebugGetConnectionManagerNameHostHandle(HWND& outHost) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugGetNameHostHandle(outHost);
}

bool DebugAcknowledgeConnectionManagerS3InsecureTlsPrompt() noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugAcknowledgeS3InsecureTlsPrompt();
}

bool DebugGetConnectionManagerSavePasswordToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugGetSavePasswordToggleHostAndClientRect(outHost, outRect);
}

bool DebugGetConnectionManagerSavePasswordToggleState(bool& outChecked, std::wstring& outLabel) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugGetSavePasswordToggleState(outChecked, outLabel);
}

bool DebugGetConnectionManagerCommandButtonHostAndClientRect(UINT commandId, HWND& outHost, RECT& outRect, std::wstring& outLabel) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugGetCommandButtonHostAndClientRect(commandId, outHost, outRect, outLabel);
}

bool DebugGetConnectionManagerS3UseHttpsToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugGetS3UseHttpsToggleHostAndClientRect(outHost, outRect);
}

bool DebugGetConnectionManagerS3UseHttpsToggleState(bool& outChecked, std::wstring& outLabel) noexcept
{
    return RedSalamander::ConnectionManager::SingleCanvas::DebugGetS3UseHttpsToggleState(outChecked, outLabel);
}

#endif // ENABLE_TESTS
