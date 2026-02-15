#include "Framework.h"

#include "ConnectionManagerDialog.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <cwctype>
#include <filesystem>
#include <format>
#include <functional>
#include <limits>
#include <memory>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "ConnectionCredentialPromptDialog.h"
#include "ConnectionSecrets.h"
#include "Helpers.h"
#include "HostServices.h"
#include "SettingsSave.h"
#include "SettingsSchemaExport.h"
#include "ThemedControls.h"
#include "ThemedInputFrames.h"
#include "WindowMaximizeBehavior.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "WindowsHello.h"
#include "resource.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <commctrl.h>
#include <commdlg.h>
#include <objbase.h>
#include <uxtheme.h>
#include <wincred.h>

namespace
{
using RedSalamander::Connections::BuildCredentialTargetName;
using RedSalamander::Connections::DeleteGenericCredential;
using RedSalamander::Connections::SaveGenericCredential;
using RedSalamander::Connections::SecretKind;

using ThemedControls::ApplyModernEditStyle;
using ThemedControls::ApplyThemeToComboBox;
using ThemedControls::ApplyThemeToListView;
using ThemedControls::BlendColor;
using ThemedControls::DrawThemedPushButton;
using ThemedControls::DrawThemedSwitchToggle;
using ThemedControls::EnableOwnerDrawButton;
using ThemedControls::GetControlSurfaceColor;
using ThemedControls::ScaleDip;

struct ProtocolEntry
{
    const wchar_t* pluginId = nullptr;
    const wchar_t* label    = nullptr;
};

constexpr wchar_t kConnectionManagerWindowId[] = L"ConnectionManagerWindow";

constexpr ProtocolEntry kProtocols[] = {
    {L"builtin/file-system-ftp", L"FTP"},
    {L"builtin/file-system-sftp", L"SFTP"},
    {L"builtin/file-system-scp", L"SCP"},
    {L"builtin/file-system-imap", L"IMAP"},
    {L"builtin/file-system-s3", L"S3"},
    {L"builtin/file-system-s3table", L"S3 Table"},
};

struct AwsRegionEntry
{
    const wchar_t* code = nullptr;
    const wchar_t* name = nullptr;
};

constexpr AwsRegionEntry kAwsRegions[] = {
    {L"af-south-1", L"Africa (Cape Town)"},
    {L"ap-east-1", L"Asia Pacific (Hong Kong)"},
    {L"ap-east-2", L"Asia Pacific (Taipei)"},
    {L"ap-northeast-1", L"Asia Pacific (Tokyo)"},
    {L"ap-northeast-2", L"Asia Pacific (Seoul)"},
    {L"ap-northeast-3", L"Asia Pacific (Osaka)"},
    {L"ap-south-1", L"Asia Pacific (Mumbai)"},
    {L"ap-south-2", L"Asia Pacific (Hyderabad)"},
    {L"ap-southeast-1", L"Asia Pacific (Singapore)"},
    {L"ap-southeast-2", L"Asia Pacific (Sydney)"},
    {L"ap-southeast-3", L"Asia Pacific (Jakarta)"},
    {L"ap-southeast-4", L"Asia Pacific (Melbourne)"},
    {L"ap-southeast-5", L"Asia Pacific (Malaysia)"},
    {L"ap-southeast-6", L"Asia Pacific (New Zealand)"},
    {L"ap-southeast-7", L"Asia Pacific (Thailand)"},
    {L"ca-central-1", L"Canada (Central)"},
    {L"ca-west-1", L"Canada West (Calgary)"},
    {L"eu-central-1", L"Europe (Frankfurt)"},
    {L"eu-central-2", L"Europe (Zurich)"},
    {L"eu-north-1", L"Europe (Stockholm)"},
    {L"eu-south-1", L"Europe (Milan)"},
    {L"eu-south-2", L"Europe (Spain)"},
    {L"eu-west-1", L"Europe (Ireland)"},
    {L"eu-west-2", L"Europe (London)"},
    {L"eu-west-3", L"Europe (Paris)"},
    {L"il-central-1", L"Israel (Tel Aviv)"},
    {L"me-central-1", L"Middle East (UAE)"},
    {L"me-south-1", L"Middle East (Bahrain)"},
    {L"mx-central-1", L"Mexico (Central)"},
    {L"sa-east-1", L"South America (Sao Paulo)"},
    {L"us-east-1", L"US East (N. Virginia)"},
    {L"us-east-2", L"US East (Ohio)"},
    {L"us-west-1", L"US West (N. California)"},
    {L"us-west-2", L"US West (Oregon)"},
};

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

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text) noexcept
{
    size_t start = 0;
    while (start < text.size() && std::iswspace(static_cast<wint_t>(text[start])) != 0)
    {
        ++start;
    }

    size_t end = text.size();
    while (end > start && std::iswspace(static_cast<wint_t>(text[end - 1])) != 0)
    {
        --end;
    }

    return std::wstring(text.substr(start, end - start));
}

[[nodiscard]] std::wstring
MakeUniqueConnectionName(const std::vector<Common::Settings::ConnectionProfile>& connections, std::wstring_view desired, std::wstring_view excludeId) noexcept
{
    std::wstring base = TrimWhitespace(desired);
    if (base.empty())
    {
        base = LoadStringResource(nullptr, IDS_CONNECTIONS_DEFAULT_NEW_NAME);
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

    for (int suffix = 2; suffix < 10'000; ++suffix)
    {
        std::wstring candidate;
        candidate = std::format(L"{} ({})", base, suffix);

        if (! isUsed(candidate))
        {
            return candidate;
        }
    }

    return base;
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::wstring GetWindowTextString(HWND h)
{
    const int len = GetWindowTextLengthW(h);
    std::wstring s;
    s.resize(static_cast<size_t>(std::max(0, len)));
    if (len > 0)
    {
        GetWindowTextW(h, s.data(), len + 1);
    }
    return s;
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

void SetTwoStateToggleState(HWND toggle, const AppTheme& theme, bool toggledOn) noexcept
{
    if (! toggle)
    {
        return;
    }

    if (theme.highContrast)
    {
        SendMessageW(toggle, BM_SETCHECK, toggledOn ? BST_CHECKED : BST_UNCHECKED, 0);
        return;
    }

    SetWindowLongPtrW(toggle, GWLP_USERDATA, toggledOn ? 1 : 0);
    InvalidateRect(toggle, nullptr, TRUE);
}

[[nodiscard]] bool GetTwoStateToggleState(HWND toggle, const AppTheme& theme) noexcept
{
    if (! toggle)
    {
        return false;
    }

    if (theme.highContrast)
    {
        return SendMessageW(toggle, BM_GETCHECK, 0, 0) == BST_CHECKED;
    }

    return GetWindowLongPtrW(toggle, GWLP_USERDATA) != 0;
}

void PrepareFlatControl(HWND control) noexcept
{
    if (! control)
    {
        return;
    }

    const LONG_PTR exStyle = GetWindowLongPtrW(control, GWL_EXSTYLE);
    if ((exStyle & WS_EX_CLIENTEDGE) == 0)
    {
        return;
    }

    SetWindowLongPtrW(control, GWL_EXSTYLE, exStyle & ~WS_EX_CLIENTEDGE);
    SetWindowPos(control, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    InvalidateRect(control, nullptr, TRUE);
}

void PrepareEditMargins(HWND edit) noexcept
{
    if (! edit)
    {
        return;
    }

    const UINT dpi       = GetDpiForWindow(edit);
    const int textMargin = ThemedControls::ScaleDip(dpi, 6);
    SendMessageW(edit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(textMargin, textMargin));
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

[[nodiscard]] int FindProtocolComboIndex(std::wstring_view pluginId) noexcept
{
    for (int i = 0; i < static_cast<int>(std::size(kProtocols)); ++i)
    {
        if (kProtocols[i].pluginId && pluginId == kProtocols[i].pluginId)
        {
            return i;
        }
    }
    return 0;
}

[[nodiscard]] std::wstring_view PluginIdFromProtocolComboIndex(int index) noexcept
{
    if (index < 0 || index >= static_cast<int>(std::size(kProtocols)))
    {
        return kProtocols[0].pluginId;
    }
    return kProtocols[index].pluginId ? kProtocols[index].pluginId : kProtocols[0].pluginId;
}

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

[[nodiscard]] bool IsS3PluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-s3";
}

[[nodiscard]] bool IsS3TablePluginId(std::wstring_view pluginId) noexcept
{
    return pluginId == L"builtin/file-system-s3table";
}

[[nodiscard]] bool IsAwsS3PluginId(std::wstring_view pluginId) noexcept
{
    return IsS3PluginId(pluginId) || IsS3TablePluginId(pluginId);
}

[[nodiscard]] std::wstring BuildConnectionDisplayUrl(const Common::Settings::ConnectionProfile& profile) noexcept
{
    const wchar_t* scheme = nullptr;
    if (IsFtpPluginId(profile.pluginId))
    {
        scheme = L"ftp";
    }
    else if (profile.pluginId == L"builtin/file-system-sftp")
    {
        scheme = L"sftp";
    }
    else if (profile.pluginId == L"builtin/file-system-scp")
    {
        scheme = L"scp";
    }
    else if (IsImapPluginId(profile.pluginId))
    {
        scheme = L"imap";
    }
    else if (IsS3PluginId(profile.pluginId))
    {
        scheme = L"s3";
    }
    else if (IsS3TablePluginId(profile.pluginId))
    {
        scheme = L"s3table";
    }

    if (! scheme || profile.host.empty())
    {
        return {};
    }

    std::wstring authority = profile.host;
    if (profile.port != 0)
    {
        authority = std::format(L"{}:{}", profile.host, profile.port);
    }

    std::wstring user;
    if (profile.authMode == Common::Settings::ConnectionAuthMode::Anonymous)
    {
        user = L"anonymous";
    }
    else if (! profile.userName.empty())
    {
        user = profile.userName;
    }

    const bool hideAnonymous = IsFtpPluginId(profile.pluginId) && (user == L"anonymous");
    const bool showUser      = ! user.empty() && ! hideAnonymous;
    if (showUser)
    {
        return std::format(L"{}://{}@{}", scheme, user, authority);
    }
    return std::format(L"{}://{}", scheme, authority);
}

[[nodiscard]] bool TryParsePort(std::wstring_view text, uint32_t& out) noexcept
{
    out = 0;
    if (text.empty())
    {
        return true;
    }

    uint32_t value = 0;
    for (wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }
        const uint32_t digit = static_cast<uint32_t>(ch - L'0');
        if (value > (std::numeric_limits<uint32_t>::max() - digit) / 10u)
        {
            return false;
        }
        value = (value * 10u) + digit;
    }

    if (value > 65535u)
    {
        return false;
    }

    out = value;
    return true;
}

[[nodiscard]] bool HasCredential(std::wstring_view targetName) noexcept
{
    if (targetName.empty())
    {
        return false;
    }

    std::wstring targetCopy;
    targetCopy.assign(targetName);

    PCREDENTIALW raw = nullptr;
    if (! CredReadW(targetCopy.c_str(), CRED_TYPE_GENERIC, 0, &raw))
    {
        return false;
    }

    wil::unique_any<PCREDENTIALW, decltype(&::CredFree), ::CredFree> cred(raw);

    const BYTE* blobBytes = cred.get()->CredentialBlob;
    const DWORD byteCount = cred.get()->CredentialBlobSize;
    if (! blobBytes || byteCount < sizeof(wchar_t) || (byteCount % sizeof(wchar_t)) != 0)
    {
        return false;
    }

    const size_t charCount = static_cast<size_t>(byteCount / sizeof(wchar_t));
    const wchar_t* blob    = reinterpret_cast<const wchar_t*>(blobBytes);
    if (blob[charCount - 1u] != L'\0')
    {
        return false;
    }

    return blob[0] != L'\0';
}

[[nodiscard]] bool IsQuickConnectProfile(const Common::Settings::ConnectionProfile& profile) noexcept
{
    return RedSalamander::Connections::IsQuickConnectConnectionId(profile.id);
}

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

[[nodiscard]] std::optional<bool> ExtraGetBool(const Common::Settings::JsonValue& extra, std::string_view key) noexcept
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

        const auto* b = std::get_if<bool>(&v.value);
        if (! b)
        {
            return std::nullopt;
        }

        return *b;
    }

    return std::nullopt;
}

[[nodiscard]] std::optional<uint32_t> ExtraGetUInt32(const Common::Settings::JsonValue& extra, std::string_view key) noexcept
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

        if (const auto* n = std::get_if<uint64_t>(&v.value))
        {
            if (*n <= std::numeric_limits<uint32_t>::max())
            {
                return static_cast<uint32_t>(*n);
            }
            return std::nullopt;
        }

        if (const auto* n = std::get_if<int64_t>(&v.value))
        {
            if (*n >= 0 && *n <= static_cast<int64_t>(std::numeric_limits<uint32_t>::max()))
            {
                return static_cast<uint32_t>(*n);
            }
            return std::nullopt;
        }

        return std::nullopt;
    }

    return std::nullopt;
}

[[nodiscard]] std::wstring MakeSavedSecretPlaceholder(std::wstring_view connectionId) noexcept
{
    uint64_t seed = static_cast<uint64_t>(GetTickCount64());
    seed ^= static_cast<uint64_t>(std::hash<std::wstring_view>{}(connectionId));

    const size_t length = 8u + static_cast<size_t>(seed % 9u); // 8-16 dots
    return std::wstring(length, L'\u2022');
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

struct DialogState
{
    DialogState()                              = default;
    DialogState(const DialogState&)            = delete;
    DialogState& operator=(const DialogState&) = delete;

    bool modeless             = false;
    HWND connectNotifyWindow  = nullptr;
    uint8_t connectTargetPane = 0; // app-defined: 0=Left, 1=Right

    Common::Settings::Settings* baselineSettings = nullptr;
    std::wstring appId;
    AppTheme theme{};
    std::wstring filterPluginId;

    std::vector<Common::Settings::ConnectionProfile> connections;
    std::vector<size_t> viewToModel;
    std::unordered_set<std::wstring> baselineConnectionIds;
    std::unordered_map<std::wstring, bool> baselineSavePasswordById;

    std::unordered_map<std::wstring, std::wstring> stagedPasswordById;
    std::unordered_map<std::wstring, std::wstring> stagedPassphraseById;
    std::unordered_map<std::wstring, std::wstring> secretPlaceholderById;
    std::unordered_set<std::wstring> secretDirtyIds;
    std::unordered_map<std::wstring, uint64_t> lastHelloVerificationTickByConnectionId;

    std::wstring selectedConnectionName;

    wil::unique_hbrush backgroundBrush;
    wil::unique_hbrush cardBrush;
    wil::unique_hbrush inputBrush;
    wil::unique_hbrush inputFocusedBrush;
    wil::unique_hbrush inputDisabledBrush;
    COLORREF cardBackgroundColor          = RGB(255, 255, 255);
    COLORREF inputBackgroundColor         = RGB(255, 255, 255);
    COLORREF inputFocusedBackgroundColor  = RGB(255, 255, 255);
    COLORREF inputDisabledBackgroundColor = RGB(255, 255, 255);

    ThemedInputFrames::FrameStyle inputFrameStyle{};

    wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> boldFont;
    wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> titleFont;

    std::vector<RECT> cards;
    std::wstring toggleOnLabel;
    std::wstring toggleOffLabel;
    std::wstring quickConnectLabel;

    wil::unique_hwnd nameFrame;
    wil::unique_hwnd protocolFrame;
    wil::unique_hwnd hostFrame;
    wil::unique_hwnd awsRegionFrame;
    wil::unique_hwnd portFrame;
    wil::unique_hwnd initialPathFrame;
    wil::unique_hwnd userFrame;
    wil::unique_hwnd secretFrame;
    wil::unique_hwnd s3EndpointOverrideFrame;
    wil::unique_hwnd sshPrivateKeyFrame;
    wil::unique_hwnd sshKnownHostsFrame;

    HWND sectionConnection           = nullptr;
    HWND sectionAuth                 = nullptr;
    HWND sectionS3                   = nullptr;
    HWND sectionSsh                  = nullptr;
    HWND nameLabel                   = nullptr;
    HWND protocolLabel               = nullptr;
    HWND hostLabel                   = nullptr;
    HWND portLabel                   = nullptr;
    HWND initialPathLabel            = nullptr;
    HWND anonymousLabel              = nullptr;
    HWND userLabel                   = nullptr;
    HWND secretLabel                 = nullptr;
    HWND savePasswordLabel           = nullptr;
    HWND requireHelloLabel           = nullptr;
    HWND ignoreSslTrustLabel         = nullptr;
    HWND s3EndpointOverrideLabel     = nullptr;
    HWND s3UseHttpsLabel             = nullptr;
    HWND s3VerifyTlsLabel            = nullptr;
    HWND s3UseVirtualAddressingLabel = nullptr;
    HWND sshPrivateKeyLabel          = nullptr;
    HWND sshKnownHostsLabel          = nullptr;
    HWND listTitle                   = nullptr;
    HWND btnNew                      = nullptr;
    HWND btnRename                   = nullptr;
    HWND btnRemove                   = nullptr;
    HWND btnConnect                  = nullptr;
    HWND btnClose                    = nullptr;
    HWND btnCancel                   = nullptr;
    HWND settingsHost                = nullptr;

    HWND list                         = nullptr;
    HWND nameEdit                     = nullptr;
    HWND protocolCombo                = nullptr;
    HWND hostEdit                     = nullptr;
    HWND awsRegionCombo               = nullptr;
    HWND portEdit                     = nullptr;
    HWND initialPathEdit              = nullptr;
    HWND anonymousToggle              = nullptr;
    HWND userEdit                     = nullptr;
    HWND secretEdit                   = nullptr;
    HWND showSecretBtn                = nullptr;
    HWND savePasswordToggle           = nullptr;
    HWND requireHelloToggle           = nullptr;
    HWND ignoreSslTrustToggle         = nullptr;
    HWND s3EndpointOverrideEdit       = nullptr;
    HWND s3UseHttpsToggle             = nullptr;
    HWND s3VerifyTlsToggle            = nullptr;
    HWND s3UseVirtualAddressingToggle = nullptr;
    HWND sshPrivateKeyEdit            = nullptr;
    HWND sshPrivateKeyBrowseBtn       = nullptr;
    HWND sshKnownHostsEdit            = nullptr;
    HWND sshKnownHostsBrowseBtn       = nullptr;

    int selectedListIndex = -1;
    bool loadingControls  = false;
    bool secretVisible    = false;

    RECT settingsViewport{}; // In dialog coordinates (client): the host client viewport where cards are painted.
    int settingsScrollOffset = 0;
    int settingsScrollMax    = 0;
};

wil::unique_hwnd g_connectionManagerDialog;

[[nodiscard]] HWND NormalizeOwnerWindow(HWND owner) noexcept
{
    if (owner && IsWindow(owner))
    {
        return GetAncestor(owner, GA_ROOT);
    }
    return nullptr;
}

void PopulateStateFromSettings(DialogState& state, const Common::Settings::Settings& settings, std::wstring_view filterPluginId) noexcept
{
    state.connections.clear();
    state.viewToModel.clear();
    state.baselineConnectionIds.clear();
    state.baselineSavePasswordById.clear();
    state.stagedPasswordById.clear();
    state.stagedPassphraseById.clear();
    state.secretPlaceholderById.clear();
    state.secretDirtyIds.clear();
    state.lastHelloVerificationTickByConnectionId.clear();
    state.selectedConnectionName.clear();

    if (settings.connections)
    {
        state.connections = settings.connections->items;
        for (const auto& c : settings.connections->items)
        {
            if (IsQuickConnectProfile(c))
            {
                continue;
            }
            if (! c.id.empty())
            {
                state.baselineConnectionIds.insert(c.id);
                state.baselineSavePasswordById.emplace(c.id, c.savePassword);
            }
        }
    }

    state.connections.erase(std::remove_if(state.connections.begin(), state.connections.end(), IsQuickConnectProfile), state.connections.end());

    RedSalamander::Connections::EnsureQuickConnectProfile(filterPluginId);
    Common::Settings::ConnectionProfile quickConnect;
    RedSalamander::Connections::GetQuickConnectProfile(quickConnect);
    if (! filterPluginId.empty())
    {
        quickConnect.pluginId.assign(filterPluginId);
    }
    state.connections.insert(state.connections.begin(), std::move(quickConnect));
}

void CloseConnectionManagerWindow(HWND dlg, const DialogState& state, INT_PTR result) noexcept
{
    if (! dlg)
    {
        return;
    }

    if (state.modeless)
    {
        DestroyWindow(dlg);
        return;
    }

    EndDialog(dlg, result);
}

void NotifyConnectSelection(const DialogState& state, std::wstring_view connectionName) noexcept
{
    if (! state.modeless || ! state.connectNotifyWindow || connectionName.empty())
    {
        return;
    }

    auto owned = std::make_unique<std::wstring>(connectionName);
    static_cast<void>(
        PostMessagePayload(state.connectNotifyWindow, WndMsg::kConnectionManagerConnect, static_cast<WPARAM>(state.connectTargetPane), std::move(owned)));
}

void EnsureFonts(DialogState& state, HFONT baseFont) noexcept
{
    if (! baseFont)
    {
        baseFont = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
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

void PersistSettings(HWND owner, Common::Settings::Settings& settings, std::wstring_view appId) noexcept
{
    if (appId.empty())
    {
        return;
    }

    const Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(settings);
    const HRESULT hr                                = Common::Settings::SaveSettings(appId, settingsToSave);
    if (SUCCEEDED(hr))
    {
        const HRESULT schemaHr = SaveAggregatedSettingsSchema(appId, settings);
        if (FAILED(schemaHr))
        {
            Debug::Error(L"Failed to write aggregated settings schema (hr=0x{:08X})", static_cast<unsigned long>(schemaHr));
        }
        return;
    }

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(appId);
    Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(hr), settingsPath.wstring());

    if (! owner)
    {
        return;
    }

    const std::wstring message = FormatStringResource(nullptr, IDS_FMT_SETTINGS_SAVE_FAILED, settingsPath.wstring(), static_cast<unsigned long>(hr));
    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
    ShowDialogAlert(owner, HOST_ALERT_ERROR, title, message);
}

void EnsureControls(DialogState& state, HWND dlg) noexcept
{
    state.listTitle                    = GetDlgItem(dlg, IDC_CONNECTION_LIST_TITLE);
    state.list                         = GetDlgItem(dlg, IDC_CONNECTION_LIST);
    state.sectionConnection            = GetDlgItem(dlg, IDC_CONNECTION_SECTION_CONNECTION);
    state.sectionAuth                  = GetDlgItem(dlg, IDC_CONNECTION_SECTION_AUTH);
    state.sectionS3                    = GetDlgItem(dlg, IDC_CONNECTION_SECTION_S3);
    state.sectionSsh                   = GetDlgItem(dlg, IDC_CONNECTION_SECTION_SSH);
    state.nameLabel                    = GetDlgItem(dlg, IDC_CONNECTION_LABEL_NAME);
    state.protocolLabel                = GetDlgItem(dlg, IDC_CONNECTION_LABEL_PROTOCOL);
    state.hostLabel                    = GetDlgItem(dlg, IDC_CONNECTION_LABEL_HOST);
    state.portLabel                    = GetDlgItem(dlg, IDC_CONNECTION_LABEL_PORT);
    state.initialPathLabel             = GetDlgItem(dlg, IDC_CONNECTION_LABEL_INITIAL_PATH);
    state.anonymousLabel               = GetDlgItem(dlg, IDC_CONNECTION_LABEL_ANONYMOUS);
    state.userLabel                    = GetDlgItem(dlg, IDC_CONNECTION_LABEL_USER);
    state.secretLabel                  = GetDlgItem(dlg, IDC_CONNECTION_LABEL_SECRET);
    state.savePasswordLabel            = GetDlgItem(dlg, IDC_CONNECTION_LABEL_SAVE_PASSWORD);
    state.requireHelloLabel            = GetDlgItem(dlg, IDC_CONNECTION_LABEL_REQUIRE_HELLO);
    state.ignoreSslTrustLabel          = GetDlgItem(dlg, IDC_CONNECTION_LABEL_IGNORE_SSL_TRUST);
    state.s3EndpointOverrideLabel      = GetDlgItem(dlg, IDC_CONNECTION_LABEL_S3_ENDPOINT_OVERRIDE);
    state.s3UseHttpsLabel              = GetDlgItem(dlg, IDC_CONNECTION_LABEL_S3_USE_HTTPS);
    state.s3VerifyTlsLabel             = GetDlgItem(dlg, IDC_CONNECTION_LABEL_S3_VERIFY_TLS);
    state.s3UseVirtualAddressingLabel  = GetDlgItem(dlg, IDC_CONNECTION_LABEL_S3_USE_VIRTUAL_ADDRESSING);
    state.sshPrivateKeyLabel           = GetDlgItem(dlg, IDC_CONNECTION_LABEL_SSH_PRIVATEKEY);
    state.sshKnownHostsLabel           = GetDlgItem(dlg, IDC_CONNECTION_LABEL_SSH_KNOWNHOSTS);
    state.nameEdit                     = GetDlgItem(dlg, IDC_CONNECTION_NAME);
    state.protocolCombo                = GetDlgItem(dlg, IDC_CONNECTION_PROTOCOL);
    state.hostEdit                     = GetDlgItem(dlg, IDC_CONNECTION_HOST);
    state.portEdit                     = GetDlgItem(dlg, IDC_CONNECTION_PORT);
    state.initialPathEdit              = GetDlgItem(dlg, IDC_CONNECTION_INITIAL_PATH);
    state.anonymousToggle              = GetDlgItem(dlg, IDC_CONNECTION_ANONYMOUS);
    state.userEdit                     = GetDlgItem(dlg, IDC_CONNECTION_USER);
    state.secretEdit                   = GetDlgItem(dlg, IDC_CONNECTION_PASSWORD);
    state.showSecretBtn                = GetDlgItem(dlg, IDC_CONNECTION_SHOW_SECRET);
    state.savePasswordToggle           = GetDlgItem(dlg, IDC_CONNECTION_SAVE_PASSWORD);
    state.requireHelloToggle           = GetDlgItem(dlg, IDC_CONNECTION_REQUIRE_HELLO);
    state.ignoreSslTrustToggle         = GetDlgItem(dlg, IDC_CONNECTION_IGNORE_SSL_TRUST);
    state.s3EndpointOverrideEdit       = GetDlgItem(dlg, IDC_CONNECTION_S3_ENDPOINT_OVERRIDE);
    state.s3UseHttpsToggle             = GetDlgItem(dlg, IDC_CONNECTION_S3_USE_HTTPS);
    state.s3VerifyTlsToggle            = GetDlgItem(dlg, IDC_CONNECTION_S3_VERIFY_TLS);
    state.s3UseVirtualAddressingToggle = GetDlgItem(dlg, IDC_CONNECTION_S3_USE_VIRTUAL_ADDRESSING);
    state.sshPrivateKeyEdit            = GetDlgItem(dlg, IDC_CONNECTION_SSH_PRIVATEKEY);
    state.sshPrivateKeyBrowseBtn       = GetDlgItem(dlg, IDC_CONNECTION_SSH_PRIVATEKEY_BROWSE);
    state.sshKnownHostsEdit            = GetDlgItem(dlg, IDC_CONNECTION_SSH_KNOWNHOSTS);
    state.sshKnownHostsBrowseBtn       = GetDlgItem(dlg, IDC_CONNECTION_SSH_KNOWNHOSTS_BROWSE);
    state.btnNew                       = GetDlgItem(dlg, IDC_CONNECTION_NEW);
    state.btnRename                    = GetDlgItem(dlg, IDC_CONNECTION_RENAME);
    state.btnRemove                    = GetDlgItem(dlg, IDC_CONNECTION_REMOVE);
    state.btnConnect                   = GetDlgItem(dlg, IDOK);
    state.btnClose                     = GetDlgItem(dlg, IDC_CONNECTION_CLOSE);
    state.btnCancel                    = GetDlgItem(dlg, IDCANCEL);
}

void UpdateSecretVisibility(DialogState& state) noexcept
{
    if (! state.secretEdit)
    {
        return;
    }

    DWORD selStart = 0;
    DWORD selEnd   = 0;
    SendMessageW(state.secretEdit, EM_GETSEL, reinterpret_cast<WPARAM>(&selStart), reinterpret_cast<LPARAM>(&selEnd));

    LONG_PTR style = GetWindowLongPtrW(state.secretEdit, GWL_STYLE);
    if (state.secretVisible)
    {
        style &= ~static_cast<LONG_PTR>(ES_PASSWORD);
        SetWindowLongPtrW(state.secretEdit, GWL_STYLE, style);
        SendMessageW(state.secretEdit, EM_SETPASSWORDCHAR, 0, 0);
    }
    else
    {
        style |= ES_PASSWORD;
        SetWindowLongPtrW(state.secretEdit, GWL_STYLE, style);
        SendMessageW(state.secretEdit, EM_SETPASSWORDCHAR, static_cast<WPARAM>(L'\u2022'), 0);
    }

    SetWindowPos(state.secretEdit, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    SendMessageW(state.secretEdit, EM_SETSEL, selStart, selEnd);
    InvalidateRect(state.secretEdit, nullptr, TRUE);

    if (state.showSecretBtn)
    {
        const UINT labelId = state.secretVisible ? IDS_CONNECTIONS_BTN_HIDE_SECRET : IDS_CONNECTIONS_BTN_SHOW_SECRET;
        SetWindowTextW(state.showSecretBtn, LoadStringResource(nullptr, labelId).c_str());
    }
}

[[nodiscard]] HRESULT PromptWindowsHelloIfRequired(HWND owner, DialogState& state, const Common::Settings::ConnectionProfile& profile) noexcept
{
    if (! profile.requireWindowsHello)
    {
        return S_OK;
    }

    const Common::Settings::ConnectionsSettings defaults{};
    bool bypassWindowsHello                  = false;
    uint32_t windowsHelloReauthTimeoutMinute = defaults.windowsHelloReauthTimeoutMinute;
    if (state.baselineSettings && state.baselineSettings->connections)
    {
        bypassWindowsHello              = state.baselineSettings->connections->bypassWindowsHello;
        windowsHelloReauthTimeoutMinute = state.baselineSettings->connections->windowsHelloReauthTimeoutMinute;
    }

    if (bypassWindowsHello)
    {
        return S_OK;
    }

    const uint64_t windowsHelloReauthTimeoutMs = static_cast<uint64_t>(windowsHelloReauthTimeoutMinute) * 60'000ull;

    bool shouldPrompt = true;
    if (windowsHelloReauthTimeoutMs != 0 && ! profile.id.empty())
    {
        const uint64_t now = GetTickCount64();
        if (const auto it = state.lastHelloVerificationTickByConnectionId.find(profile.id); it != state.lastHelloVerificationTickByConnectionId.end())
        {
            const uint64_t elapsed = now - it->second;
            if (elapsed < windowsHelloReauthTimeoutMs)
            {
                shouldPrompt = false;
            }
        }
    }

    if (! shouldPrompt)
    {
        return S_OK;
    }

    const HRESULT helloHr = RedSalamander::Security::VerifyWindowsHelloForWindow(owner, LoadStringResource(nullptr, IDS_CONNECTIONS_HELLO_PROMPT_CREDENTIAL));
    if (FAILED(helloHr))
    {
        Debug::Warning(L"ConnectionManager: Windows Hello verification failed for connection '{}' (id={}) hr=0x{:08X}",
                       profile.name,
                       profile.id,
                       static_cast<unsigned long>(helloHr));
        return helloHr;
    }

    if (windowsHelloReauthTimeoutMs != 0 && ! profile.id.empty())
    {
        state.lastHelloVerificationTickByConnectionId[profile.id] = GetTickCount64();
    }

    return S_OK;
}

[[nodiscard]] HRESULT
LoadStoredSecretForProfile(HWND owner, DialogState& state, const Common::Settings::ConnectionProfile& profile, std::wstring& secretOut) noexcept
{
    secretOut.clear();

    if (profile.id.empty())
    {
        return E_INVALIDARG;
    }
    if (! profile.savePassword)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }
    if (profile.authMode == Common::Settings::ConnectionAuthMode::Anonymous)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const bool sshPassphrase = profile.authMode == Common::Settings::ConnectionAuthMode::SshKey;
    const SecretKind kind    = sshPassphrase ? SecretKind::SshKeyPassphrase : SecretKind::Password;

    const HRESULT helloHr = PromptWindowsHelloIfRequired(owner, state, profile);
    if (FAILED(helloHr))
    {
        return helloHr;
    }

    if (IsQuickConnectProfile(profile))
    {
        std::wstring secret;
        const HRESULT loadHr = RedSalamander::Connections::LoadQuickConnectSecret(kind, secret);
        if (FAILED(loadHr))
        {
            Debug::Error(L"ConnectionManager: LoadQuickConnectSecret failed connection='{}' id='{}' kind='{}' hr=0x{:08X}",
                         profile.name,
                         profile.id,
                         sshPassphrase ? L"sshKeyPassphrase" : L"password",
                         static_cast<unsigned long>(loadHr));
            return loadHr;
        }

        secretOut = std::move(secret);
        return S_OK;
    }

    const std::wstring targetName = BuildCredentialTargetName(profile.id, kind);
    if (targetName.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    std::wstring userName;
    std::wstring secret;
    const HRESULT loadHr = RedSalamander::Connections::LoadGenericCredential(targetName, userName, secret);
    if (FAILED(loadHr))
    {
        Debug::Error(L"ConnectionManager: LoadGenericCredential failed connection='{}' id='{}' kind='{}' hr=0x{:08X}",
                     profile.name,
                     profile.id,
                     sshPassphrase ? L"sshKeyPassphrase" : L"password",
                     static_cast<unsigned long>(loadHr));
        return loadHr;
    }

    secretOut = std::move(secret);
    return S_OK;
}

[[nodiscard]] bool ShouldCommitSecretsForProfile(const DialogState& state, const Common::Settings::ConnectionProfile& profile) noexcept
{
    if (profile.id.empty())
    {
        return false;
    }

    const bool sshPassphrase = profile.authMode == Common::Settings::ConnectionAuthMode::SshKey;
    const auto& stagedMap    = sshPassphrase ? state.stagedPassphraseById : state.stagedPasswordById;
    if (const auto it = stagedMap.find(profile.id); it != stagedMap.end() && ! it->second.empty())
    {
        return true;
    }

    const auto baselineIt = state.baselineSavePasswordById.find(profile.id);
    if (baselineIt == state.baselineSavePasswordById.end())
    {
        return false;
    }

    return baselineIt->second != profile.savePassword;
}

[[nodiscard]] std::wstring FormatHresultForUi(HRESULT hr) noexcept
{
    wil::unique_hlocal_string message;
    if (SUCCEEDED(::FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS,
                                   nullptr,
                                   static_cast<DWORD>(hr),
                                   MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                                   reinterpret_cast<LPWSTR>(message.addressof()),
                                   0,
                                   nullptr)) &&
        message.get())
    {
        std::wstring text(message.get());
        while (! text.empty() && (text.back() == L'\r' || text.back() == L'\n'))
        {
            text.pop_back();
        }
        return std::format(L"0x{:08X}: {}", static_cast<unsigned long>(hr), text);
    }

    return std::format(L"0x{:08X}", static_cast<unsigned long>(hr));
}

void ApplyPluginDefaultsToNewProfile(const DialogState& state, Common::Settings::ConnectionProfile& profile) noexcept
{
    if (IsAwsS3PluginId(profile.pluginId))
    {
        if (profile.host.empty())
        {
            profile.host = L"us-east-1";
        }
        profile.port = 0;
    }

    if (profile.pluginId.empty() || ! state.baselineSettings)
    {
        return;
    }

    const auto it = state.baselineSettings->plugins.configurationByPluginId.find(profile.pluginId);
    if (it == state.baselineSettings->plugins.configurationByPluginId.end())
    {
        return;
    }

    const Common::Settings::JsonValue& config = it->second;

    if (IsAwsS3PluginId(profile.pluginId))
    {
        if (profile.host.empty())
        {
            profile.host = L"us-east-1";
        }
        if (const auto region = ExtraGetString(config, "defaultRegion"); region.has_value() && ! region->empty())
        {
            profile.host = *region;
        }

        profile.port = 0;

        if (const auto endpoint = ExtraGetString(config, "defaultEndpointOverride"); endpoint.has_value())
        {
            ExtraSetString(profile.extra, "endpointOverride", *endpoint);
        }
        if (const auto useHttps = ExtraGetBool(config, "useHttps"); useHttps.has_value())
        {
            ExtraSetBool(profile.extra, "useHttps", useHttps.value());
        }
        if (const auto verifyTls = ExtraGetBool(config, "verifyTls"); verifyTls.has_value())
        {
            ExtraSetBool(profile.extra, "verifyTls", verifyTls.value());
        }
        if (IsS3PluginId(profile.pluginId))
        {
            if (const auto virtualHost = ExtraGetBool(config, "useVirtualAddressing"); virtualHost.has_value())
            {
                ExtraSetBool(profile.extra, "useVirtualAddressing", virtualHost.value());
            }
        }

        return;
    }

    if (profile.host.empty())
    {
        if (const auto host = ExtraGetString(config, "defaultHost"); host.has_value() && ! host->empty())
        {
            profile.host = *host;
        }
    }

    if (profile.port == 0)
    {
        if (const auto port = ExtraGetUInt32(config, "defaultPort"); port.has_value() && *port <= 65535u)
        {
            profile.port = *port;
        }
    }

    if (profile.initialPath.empty() || profile.initialPath == L"/")
    {
        if (const auto basePath = ExtraGetString(config, "defaultBasePath"); basePath.has_value() && ! basePath->empty())
        {
            profile.initialPath = *basePath;
        }
    }
    if (! profile.initialPath.empty() && profile.initialPath.front() != L'/')
    {
        profile.initialPath.insert(profile.initialPath.begin(), L'/');
    }

    if (profile.userName.empty())
    {
        if (const auto user = ExtraGetString(config, "defaultUser"); user.has_value())
        {
            profile.userName = *user;
        }
    }

    if (IsFtpPluginId(profile.pluginId))
    {
        // Anonymous login is always opt-in.
        profile.authMode = Common::Settings::ConnectionAuthMode::Password;
        if (profile.userName.empty() || EqualsIgnoreCase(profile.userName, L"anonymous"))
        {
            profile.userName.clear();
        }
    }
}

void PopulateProtocolCombo(HWND combo)
{
    if (! combo)
    {
        return;
    }

    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    for (const auto& p : kProtocols)
    {
        if (! p.label)
        {
            continue;
        }
        const int index = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(p.label)));
        if (index >= 0)
        {
            SendMessageW(combo, CB_SETITEMDATA, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(p.pluginId));
        }
    }

    SendMessageW(combo, CB_SETCURSEL, 0, 0);
}

void PopulateAwsRegionCombo(HWND combo)
{
    if (! combo)
    {
        return;
    }

    SendMessageW(combo, CB_RESETCONTENT, 0, 0);

    for (const auto& region : kAwsRegions)
    {
        if (! region.code || ! region.name)
        {
            continue;
        }

        std::wstring display;
        display = std::format(L"{} ({})", region.name, region.code);

        const int index = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(display.c_str())));
        if (index >= 0)
        {
            SendMessageW(combo, CB_SETITEMDATA, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(region.code));
        }
    }

    SendMessageW(combo, CB_SETCURSEL, static_cast<WPARAM>(-1), 0);
}

void SetupListViewColumns(HWND list)
{
    if (! list)
    {
        return;
    }

    ListView_DeleteAllItems(list);
    while (ListView_DeleteColumn(list, 0))
    {
    }

    LVCOLUMNW col{};
    col.mask = LVCF_WIDTH;
    col.cx   = 200;
    ListView_InsertColumn(list, 0, &col);
}

void RebuildList(HWND dlg, DialogState& state)
{
    UNREFERENCED_PARAMETER(dlg);

    if (! state.list)
    {
        return;
    }

    const int prevSel = state.selectedListIndex;

    state.viewToModel.clear();
    ListView_DeleteAllItems(state.list);

    for (size_t modelIndex = 0; modelIndex < state.connections.size(); ++modelIndex)
    {
        const auto& profile = state.connections[modelIndex];
        if (! state.filterPluginId.empty() && profile.pluginId != state.filterPluginId)
        {
            continue;
        }

        LVITEMW item{};
        item.mask          = LVIF_TEXT | LVIF_PARAM;
        item.iItem         = static_cast<int>(state.viewToModel.size());
        item.pszText       = const_cast<wchar_t*>((IsQuickConnectProfile(profile) && ! state.quickConnectLabel.empty()) ? state.quickConnectLabel.c_str()
                                                                                                                  : profile.name.c_str());
        item.lParam        = static_cast<LPARAM>(modelIndex);
        const int inserted = ListView_InsertItem(state.list, &item);
        if (inserted >= 0)
        {
            state.viewToModel.push_back(modelIndex);
        }
    }

    state.selectedListIndex = -1;

    if (prevSel >= 0 && prevSel < ListView_GetItemCount(state.list))
    {
        ListView_SetItemState(state.list, prevSel, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        state.selectedListIndex = prevSel;
        return;
    }

    if (ListView_GetItemCount(state.list) > 0)
    {
        ListView_SetItemState(state.list, 0, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_SetSelectionMark(state.list, 0);
        state.selectedListIndex = 0;
    }
}

void EnsureListSelection(DialogState& state) noexcept
{
    if (! state.list)
    {
        return;
    }

    const int count = ListView_GetItemCount(state.list);
    if (count <= 0)
    {
        state.selectedListIndex = -1;
        return;
    }

    const int sel = ListView_GetNextItem(state.list, -1, LVNI_SELECTED);
    if (sel >= 0)
    {
        state.selectedListIndex = sel;
        return;
    }

    int desired = state.selectedListIndex;
    if (desired < 0 || desired >= count)
    {
        desired = 0;
    }

    ListView_SetItemState(state.list, desired, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetSelectionMark(state.list, desired);
    state.selectedListIndex = desired;
}

[[nodiscard]] std::optional<size_t> GetSelectedModelIndex(const DialogState& state) noexcept
{
    if (! state.list)
    {
        return std::nullopt;
    }

    int sel = ListView_GetNextItem(state.list, -1, LVNI_SELECTED);
    if (sel < 0)
    {
        sel = ListView_GetNextItem(state.list, -1, LVNI_FOCUSED);
    }
    if (sel < 0)
    {
        const int count = ListView_GetItemCount(state.list);
        if (count <= 0)
        {
            return std::nullopt;
        }

        sel = state.selectedListIndex;
        if (sel < 0 || sel >= count)
        {
            sel = 0;
        }
    }
    if (sel < 0)
    {
        return std::nullopt;
    }

    LVITEMW item{};
    item.mask  = LVIF_PARAM;
    item.iItem = sel;
    if (! ListView_GetItem(state.list, &item))
    {
        return std::nullopt;
    }

    if (item.lParam < 0)
    {
        return std::nullopt;
    }

    const size_t modelIndex = static_cast<size_t>(item.lParam);
    if (modelIndex >= state.connections.size())
    {
        return std::nullopt;
    }

    return modelIndex;
}

void UpdateControlEnabledState(DialogState& state) noexcept
{
    const auto modelIndex     = GetSelectedModelIndex(state);
    const bool hasSelection   = modelIndex.has_value();
    const auto* profile       = hasSelection ? &state.connections[modelIndex.value()] : nullptr;
    const bool isFtp          = profile && IsFtpPluginId(profile->pluginId);
    const bool isSsh          = profile && IsSshPluginId(profile->pluginId);
    const bool isImap         = profile && IsImapPluginId(profile->pluginId);
    const bool isS3           = profile && IsS3PluginId(profile->pluginId);
    const bool isS3Table      = profile && IsS3TablePluginId(profile->pluginId);
    const bool isAwsS3        = isS3 || isS3Table;
    const bool isQuickConnect = profile && IsQuickConnectProfile(*profile);
    const bool anonymous      = isFtp && profile && profile->authMode == Common::Settings::ConnectionAuthMode::Anonymous;
    const bool sshKey         = isSsh && profile && profile->authMode == Common::Settings::ConnectionAuthMode::SshKey;

    const auto show = [&](HWND hwnd, bool visible) noexcept
    {
        if (hwnd)
        {
            ShowWindow(hwnd, visible ? SW_SHOW : SW_HIDE);
        }
    };

    const auto showFrame = [&](const wil::unique_hwnd& frame, bool visible) noexcept { show(frame.get(), visible); };

    const bool showProtocol = hasSelection && state.filterPluginId.empty();
    show(state.protocolLabel, showProtocol);
    show(state.protocolCombo, showProtocol);
    showFrame(state.protocolFrame, showProtocol);

    const bool showAwsRegionCombo = hasSelection && isAwsS3;
    show(state.awsRegionCombo, showAwsRegionCombo);
    showFrame(state.awsRegionFrame, showAwsRegionCombo);

    const bool showHostEdit = hasSelection && ! isAwsS3;
    show(state.hostEdit, showHostEdit);
    showFrame(state.hostFrame, showHostEdit);

    const bool showAnonymous = hasSelection && isFtp;
    show(state.anonymousLabel, showAnonymous);
    show(state.anonymousToggle, showAnonymous);

    const bool showSshSection = hasSelection && isSsh;
    show(state.sectionSsh, showSshSection);
    show(state.sshPrivateKeyLabel, showSshSection);
    show(state.sshPrivateKeyEdit, showSshSection);
    showFrame(state.sshPrivateKeyFrame, showSshSection);
    show(state.sshPrivateKeyBrowseBtn, showSshSection);
    show(state.sshKnownHostsLabel, showSshSection);
    show(state.sshKnownHostsEdit, showSshSection);
    showFrame(state.sshKnownHostsFrame, showSshSection);
    show(state.sshKnownHostsBrowseBtn, showSshSection);

    const bool showS3Section = hasSelection && isAwsS3;
    show(state.sectionS3, showS3Section);
    show(state.s3EndpointOverrideLabel, showS3Section);
    show(state.s3EndpointOverrideEdit, showS3Section);
    showFrame(state.s3EndpointOverrideFrame, showS3Section);
    show(state.s3UseHttpsLabel, showS3Section);
    show(state.s3UseHttpsToggle, showS3Section);
    show(state.s3VerifyTlsLabel, showS3Section);
    show(state.s3VerifyTlsToggle, showS3Section);
    const bool showVirtual = showS3Section && isS3;
    show(state.s3UseVirtualAddressingLabel, showVirtual);
    show(state.s3UseVirtualAddressingToggle, showVirtual);

    if (state.secretLabel)
    {
        const UINT id =
            static_cast<UINT>(isAwsS3 ? IDS_CONNECTIONS_LABEL_SECRET_ACCESS_KEY : (sshKey ? IDS_CONNECTIONS_LABEL_PASSPHRASE : IDS_CONNECTIONS_LABEL_PASSWORD));
        SetWindowTextW(state.secretLabel, LoadStringResource(nullptr, id).c_str());
    }

    EnableWindow(state.nameEdit, hasSelection && ! isQuickConnect);
    EnableWindow(state.hostEdit, hasSelection && ! isAwsS3);
    EnableWindow(state.awsRegionCombo, hasSelection && isAwsS3);
    EnableWindow(state.portEdit, hasSelection && ! isAwsS3);
    EnableWindow(state.initialPathEdit, hasSelection);
    EnableWindow(state.anonymousToggle, showAnonymous);
    EnableWindow(state.btnRename, hasSelection && ! isQuickConnect);
    EnableWindow(state.btnRemove, hasSelection && ! isQuickConnect);

    const bool authInputsEnabled = hasSelection && ! anonymous;
    EnableWindow(state.userEdit, authInputsEnabled);
    EnableWindow(state.secretEdit, authInputsEnabled);
    EnableWindow(state.showSecretBtn, authInputsEnabled);
    EnableWindow(state.s3EndpointOverrideEdit, showS3Section);
    EnableWindow(state.s3UseHttpsToggle, showS3Section);
    EnableWindow(state.s3VerifyTlsToggle, showS3Section);
    EnableWindow(state.s3UseVirtualAddressingToggle, showVirtual);

    EnableWindow(state.savePasswordToggle, hasSelection && ! anonymous);

    const bool showIgnoreSslTrust = hasSelection && isImap;
    show(state.ignoreSslTrustLabel, showIgnoreSslTrust);
    show(state.ignoreSslTrustToggle, showIgnoreSslTrust);
    EnableWindow(state.ignoreSslTrustToggle, showIgnoreSslTrust);

    const bool showPort = hasSelection && ! isAwsS3;
    show(state.portLabel, showPort);
    show(state.portEdit, showPort);
    showFrame(state.portFrame, showPort);

    // Hidden expert setting; editable via Settings Store JSON only.
    show(state.requireHelloLabel, false);
    show(state.requireHelloToggle, false);
    EnableWindow(state.requireHelloToggle, FALSE);
}

void LoadEditorFromProfile(DialogState& state, const Common::Settings::ConnectionProfile& profile) noexcept
{
    state.loadingControls = true;

    const wchar_t* nameText = (IsQuickConnectProfile(profile) && ! state.quickConnectLabel.empty()) ? state.quickConnectLabel.c_str() : profile.name.c_str();
    SetWindowTextW(state.nameEdit, nameText);

    const int protocolIndex = FindProtocolComboIndex(profile.pluginId);
    SendMessageW(state.protocolCombo, CB_SETCURSEL, static_cast<WPARAM>(protocolIndex), 0);

    SetWindowTextW(state.hostEdit, profile.host.c_str());
    if (state.awsRegionCombo)
    {
        SetWindowTextW(state.awsRegionCombo, profile.host.c_str());
    }

    if (profile.port != 0)
    {
        SetWindowTextW(state.portEdit, std::to_wstring(profile.port).c_str());
    }
    else
    {
        SetWindowTextW(state.portEdit, L"");
    }

    const std::wstring initialPath = profile.initialPath.empty() ? L"/" : profile.initialPath;
    SetWindowTextW(state.initialPathEdit, initialPath.c_str());

    const bool anonymous = profile.authMode == Common::Settings::ConnectionAuthMode::Anonymous;
    SetTwoStateToggleState(state.anonymousToggle, state.theme, anonymous);

    SetWindowTextW(state.userEdit, profile.userName.c_str());

    state.secretPlaceholderById.erase(profile.id);
    state.secretDirtyIds.erase(profile.id);

    std::wstring secretText;
    const bool sshKeyAuth    = profile.authMode == Common::Settings::ConnectionAuthMode::SshKey;
    const bool anonymousAuth = profile.authMode == Common::Settings::ConnectionAuthMode::Anonymous;

    if (! profile.id.empty() && profile.savePassword && ! anonymousAuth)
    {
        const auto& stagedMap = sshKeyAuth ? state.stagedPassphraseById : state.stagedPasswordById;
        if (const auto it = stagedMap.find(profile.id); it != stagedMap.end() && ! it->second.empty())
        {
            secretText = it->second;
        }
        else
        {
            const SecretKind kind = sshKeyAuth ? SecretKind::SshKeyPassphrase : SecretKind::Password;
            const bool hasStored  = IsQuickConnectProfile(profile) ? RedSalamander::Connections::HasQuickConnectSecret(kind)
                                                                   : HasCredential(BuildCredentialTargetName(profile.id, kind));
            if (hasStored)
            {
                secretText = MakeSavedSecretPlaceholder(profile.id);
                if (! secretText.empty())
                {
                    state.secretPlaceholderById[profile.id] = secretText;
                }
            }
        }
    }

    SetWindowTextW(state.secretEdit, secretText.c_str());

    state.secretVisible = false;
    UpdateSecretVisibility(state);

    SetTwoStateToggleState(state.savePasswordToggle, state.theme, profile.savePassword);
    SetTwoStateToggleState(state.requireHelloToggle, state.theme, profile.savePassword && profile.requireWindowsHello);
    SetTwoStateToggleState(state.ignoreSslTrustToggle, state.theme, ExtraGetBool(profile.extra, "ignoreSslTrust").value_or(false));

    if (const auto v = ExtraGetString(profile.extra, "sshPrivateKey"); v.has_value())
    {
        SetWindowTextW(state.sshPrivateKeyEdit, v->c_str());
    }
    else
    {
        SetWindowTextW(state.sshPrivateKeyEdit, L"");
    }

    if (const auto v = ExtraGetString(profile.extra, "sshKnownHosts"); v.has_value())
    {
        SetWindowTextW(state.sshKnownHostsEdit, v->c_str());
    }
    else
    {
        SetWindowTextW(state.sshKnownHostsEdit, L"");
    }

    if (const auto v = ExtraGetString(profile.extra, "endpointOverride"); v.has_value())
    {
        SetWindowTextW(state.s3EndpointOverrideEdit, v->c_str());
    }
    else
    {
        SetWindowTextW(state.s3EndpointOverrideEdit, L"");
    }

    SetTwoStateToggleState(state.s3UseHttpsToggle, state.theme, ExtraGetBool(profile.extra, "useHttps").value_or(true));
    SetTwoStateToggleState(state.s3VerifyTlsToggle, state.theme, ExtraGetBool(profile.extra, "verifyTls").value_or(true));
    SetTwoStateToggleState(state.s3UseVirtualAddressingToggle, state.theme, ExtraGetBool(profile.extra, "useVirtualAddressing").value_or(true));

    state.loadingControls = false;
}

void StageSecretsFromEditor(DialogState& state, const Common::Settings::ConnectionProfile& profile)
{
    if (profile.id.empty())
    {
        return;
    }

    if (! profile.savePassword)
    {
        state.stagedPasswordById.erase(profile.id);
        state.stagedPassphraseById.erase(profile.id);
        state.secretPlaceholderById.erase(profile.id);
        state.secretDirtyIds.erase(profile.id);
        return;
    }

    if (! state.secretDirtyIds.contains(profile.id))
    {
        return;
    }

    const std::wstring secret = GetWindowTextString(state.secretEdit);
    if (const auto it = state.secretPlaceholderById.find(profile.id); it != state.secretPlaceholderById.end())
    {
        if (secret == it->second)
        {
            state.secretDirtyIds.erase(profile.id);
            return;
        }
    }

    state.stagedPasswordById.erase(profile.id);
    state.stagedPassphraseById.erase(profile.id);
    state.secretPlaceholderById.erase(profile.id);

    if (secret.empty())
    {
        return;
    }

    if (profile.authMode == Common::Settings::ConnectionAuthMode::SshKey)
    {
        state.stagedPassphraseById[profile.id] = secret;
        return;
    }

    if (profile.authMode == Common::Settings::ConnectionAuthMode::Password)
    {
        state.stagedPasswordById[profile.id] = secret;
    }
}

void CommitEditorToProfile(DialogState& state, Common::Settings::ConnectionProfile& profile)
{
    if (! IsQuickConnectProfile(profile))
    {
        const std::wstring rawName        = GetWindowTextString(state.nameEdit);
        const std::wstring normalizedName = TrimWhitespace(rawName);
        const std::wstring uniqueName     = MakeUniqueConnectionName(state.connections, normalizedName, profile.id);
        profile.name                      = uniqueName;

        if (state.nameEdit && ! state.loadingControls && rawName != uniqueName)
        {
            SetWindowTextW(state.nameEdit, uniqueName.c_str());
        }

        if (state.list && state.selectedListIndex >= 0 && state.selectedListIndex < ListView_GetItemCount(state.list))
        {
            ListView_SetItemText(state.list, state.selectedListIndex, 0, const_cast<wchar_t*>(profile.name.c_str()));
        }
    }

    if (state.filterPluginId.empty())
    {
        const int sel                    = static_cast<int>(SendMessageW(state.protocolCombo, CB_GETCURSEL, 0, 0));
        const std::wstring_view pluginId = PluginIdFromProtocolComboIndex(sel);
        if (! pluginId.empty())
        {
            profile.pluginId.assign(pluginId);
        }
    }
    else
    {
        profile.pluginId = state.filterPluginId;
    }

    const HWND hostControl            = IsAwsS3PluginId(profile.pluginId) && state.awsRegionCombo ? state.awsRegionCombo : state.hostEdit;
    const std::wstring rawHost        = GetWindowTextString(hostControl);
    const std::wstring normalizedHost = TrimWhitespace(rawHost);
    profile.host                      = normalizedHost;
    if (hostControl && ! state.loadingControls && rawHost != normalizedHost)
    {
        SetWindowTextW(hostControl, normalizedHost.c_str());
    }

    if (IsAwsS3PluginId(profile.pluginId))
    {
        // S3/S3 Tables connections are region-based; any port value is ignored.
        profile.port = 0;
    }
    else
    {
        uint32_t port               = 0;
        const std::wstring portText = GetWindowTextString(state.portEdit);
        if (TryParsePort(portText, port))
        {
            profile.port = port;
        }
    }

    profile.initialPath = GetWindowTextString(state.initialPathEdit);
    if (profile.initialPath.empty())
    {
        profile.initialPath = L"/";
    }
    if (! profile.initialPath.empty() && profile.initialPath.front() != L'/')
    {
        profile.initialPath.insert(profile.initialPath.begin(), L'/');
    }

    const bool anonymous = GetTwoStateToggleState(state.anonymousToggle, state.theme);
    if (anonymous && IsFtpPluginId(profile.pluginId))
    {
        profile.authMode = Common::Settings::ConnectionAuthMode::Anonymous;
        profile.userName = L"anonymous";
    }
    else
    {
        if (profile.authMode == Common::Settings::ConnectionAuthMode::Anonymous)
        {
            profile.authMode = Common::Settings::ConnectionAuthMode::Password;
        }
        const std::wstring rawUser        = GetWindowTextString(state.userEdit);
        const std::wstring normalizedUser = TrimWhitespace(rawUser);
        profile.userName                  = normalizedUser;
        if (state.userEdit && ! state.loadingControls && rawUser != normalizedUser)
        {
            SetWindowTextW(state.userEdit, normalizedUser.c_str());
        }
    }

    profile.savePassword = GetTwoStateToggleState(state.savePasswordToggle, state.theme);

    if (IsImapPluginId(profile.pluginId))
    {
        ExtraSetBool(profile.extra, "ignoreSslTrust", GetTwoStateToggleState(state.ignoreSslTrustToggle, state.theme));
    }

    if (IsAwsS3PluginId(profile.pluginId))
    {
        ExtraSetString(profile.extra, "endpointOverride", TrimWhitespace(GetWindowTextString(state.s3EndpointOverrideEdit)));
        ExtraSetBool(profile.extra, "useHttps", GetTwoStateToggleState(state.s3UseHttpsToggle, state.theme));
        ExtraSetBool(profile.extra, "verifyTls", GetTwoStateToggleState(state.s3VerifyTlsToggle, state.theme));
        if (IsS3PluginId(profile.pluginId))
        {
            ExtraSetBool(profile.extra, "useVirtualAddressing", GetTwoStateToggleState(state.s3UseVirtualAddressingToggle, state.theme));
        }
    }

    const std::wstring sshPrivateKey = GetWindowTextString(state.sshPrivateKeyEdit);
    ExtraSetString(profile.extra, "sshPrivateKey", sshPrivateKey);
    ExtraSetString(profile.extra, "sshKnownHosts", GetWindowTextString(state.sshKnownHostsEdit));

    if (IsSshPluginId(profile.pluginId))
    {
        profile.authMode = sshPrivateKey.empty() ? Common::Settings::ConnectionAuthMode::Password : Common::Settings::ConnectionAuthMode::SshKey;
    }

    StageSecretsFromEditor(state, profile);
}

[[nodiscard]] bool HasDuplicateConnectionName(const std::vector<Common::Settings::ConnectionProfile>& connections) noexcept
{
    for (size_t i = 0; i < connections.size(); ++i)
    {
        if (connections[i].name.empty())
        {
            continue;
        }
        for (size_t j = i + 1; j < connections.size(); ++j)
        {
            if (connections[j].name.empty())
            {
                continue;
            }
            if (EqualsIgnoreCase(connections[i].name, connections[j].name))
            {
                return true;
            }
        }
    }
    return false;
}

[[nodiscard]] HRESULT ValidateProfileForConnect(HWND dlg, const DialogState& state, const Common::Settings::ConnectionProfile& profile) noexcept
{
    if (profile.name.empty())
    {
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_NAME_REQUIRED));
        return E_INVALIDARG;
    }
    if (profile.host.empty() && ! IsAwsS3PluginId(profile.pluginId))
    {
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_HOST_REQUIRED));
        return E_INVALIDARG;
    }
    if (profile.pluginId.empty())
    {
        ShowDialogAlert(
            dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_PROTOCOL_REQUIRED));
        return E_INVALIDARG;
    }

    if (HasDuplicateConnectionName(state.connections))
    {
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_NAME_UNIQUE));
        return HRESULT_FROM_WIN32(ERROR_DUP_NAME);
    }

    if (IsFtpPluginId(profile.pluginId) && profile.authMode == Common::Settings::ConnectionAuthMode::Password && profile.userName.empty())
    {
        ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_USER_REQUIRED));
        return E_INVALIDARG;
    }

    if (profile.savePassword)
    {
        const bool quickConnect = IsQuickConnectProfile(profile);
        if (profile.authMode == Common::Settings::ConnectionAuthMode::Password)
        {
            const bool hasExisting = quickConnect ? RedSalamander::Connections::HasQuickConnectSecret(SecretKind::Password)
                                                  : HasCredential(BuildCredentialTargetName(profile.id, SecretKind::Password));
            const bool hasStaged   = state.stagedPasswordById.contains(profile.id);

            if (! hasExisting && ! hasStaged)
            {
                ShowDialogAlert(dlg,
                                HOST_ALERT_ERROR,
                                LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_PASSWORD_REQUIRED_TO_SAVE));
                return E_INVALIDARG;
            }
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT PromptAndStageMissingPasswordForConnect(HWND dlg, DialogState& state, Common::Settings::ConnectionProfile& profile) noexcept
{
    if (profile.id.empty() || ! profile.savePassword || profile.authMode != Common::Settings::ConnectionAuthMode::Password)
    {
        return S_OK;
    }

    bool hasStaged = false;
    if (const auto it = state.stagedPasswordById.find(profile.id); it != state.stagedPasswordById.end() && ! it->second.empty())
    {
        hasStaged = true;
    }

    const bool quickConnect = IsQuickConnectProfile(profile);
    const bool hasExisting  = quickConnect ? RedSalamander::Connections::HasQuickConnectSecret(SecretKind::Password)
                                           : HasCredential(BuildCredentialTargetName(profile.id, SecretKind::Password));

    if (hasExisting || hasStaged)
    {
        return S_OK;
    }

    const std::wstring caption          = LoadStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_CAPTION);
    const std::wstring_view displayName = (IsQuickConnectProfile(profile) && ! state.quickConnectLabel.empty())
                                              ? std::wstring_view(state.quickConnectLabel)
                                              : (profile.name.empty() ? std::wstring_view(L"(unnamed)") : std::wstring_view(profile.name));
    std::wstring message                = FormatStringResource(nullptr, IDS_CONNECTIONS_PROMPT_PASSWORD_MESSAGE_FMT, displayName);
    const std::wstring secretLabel      = LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PASSWORD);

    if (const std::wstring url = BuildConnectionDisplayUrl(profile); ! url.empty())
    {
        message = std::format(L"{}\n{}", message, url);
    }

    std::wstring userName;
    std::wstring secret;
    const HRESULT promptHr = profile.userName.empty() ? PromptForConnectionUserAndPassword(dlg, state.theme, caption, message, {}, userName, secret)
                                                      : PromptForConnectionSecret(dlg, state.theme, caption, message, secretLabel, false, secret);
    if (FAILED(promptHr) || promptHr == S_FALSE)
    {
        return promptHr;
    }

    if (! userName.empty())
    {
        profile.userName = userName;
        if (state.userEdit)
        {
            SetWindowTextW(state.userEdit, userName.c_str());
        }
    }
    state.stagedPasswordById[profile.id] = std::move(secret);

    return S_OK;
}

HRESULT CommitSecretsForProfile(const DialogState& state, const Common::Settings::ConnectionProfile& profile) noexcept
{
    const std::wstring passwordTarget   = BuildCredentialTargetName(profile.id, SecretKind::Password);
    const std::wstring passphraseTarget = BuildCredentialTargetName(profile.id, SecretKind::SshKeyPassphrase);

    if (! profile.savePassword)
    {
        Debug::Info(L"ConnectionManager: clearing stored secrets connection='{}' id='{}'", profile.name, profile.id);
        if (! passwordTarget.empty())
        {
            const HRESULT delHr = DeleteGenericCredential(passwordTarget);
            if (FAILED(delHr) && delHr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
            {
                Debug::Warning(L"ConnectionManager: DeleteGenericCredential failed connection='{}' id='{}' kind='password' hr=0x{:08X}",
                               profile.name,
                               profile.id,
                               static_cast<unsigned long>(delHr));
            }
        }
        if (! passphraseTarget.empty())
        {
            const HRESULT delHr = DeleteGenericCredential(passphraseTarget);
            if (FAILED(delHr) && delHr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
            {
                Debug::Warning(L"ConnectionManager: DeleteGenericCredential failed connection='{}' id='{}' kind='sshKeyPassphrase' hr=0x{:08X}",
                               profile.name,
                               profile.id,
                               static_cast<unsigned long>(delHr));
            }
        }
        return S_OK;
    }

    const bool sshPassphrase = profile.authMode == Common::Settings::ConnectionAuthMode::SshKey;
    const SecretKind kind    = sshPassphrase ? SecretKind::SshKeyPassphrase : SecretKind::Password;
    const auto& stagedMap    = sshPassphrase ? state.stagedPassphraseById : state.stagedPasswordById;
    const auto it            = stagedMap.find(profile.id);
    if (it == stagedMap.end() || it->second.empty())
    {
        return S_OK; // keep existing
    }

    const std::wstring targetName = BuildCredentialTargetName(profile.id, kind);
    Debug::Info(
        L"ConnectionManager: saving credential connection='{}' id='{}' kind='{}'", profile.name, profile.id, sshPassphrase ? L"sshKeyPassphrase" : L"password");
    return SaveGenericCredential(targetName, profile.userName, it->second);
}

void DeleteSecretsForRemovedConnections(const DialogState& state) noexcept;

void CommitQuickConnectSecretsAndProfile(const DialogState& state, const Common::Settings::ConnectionProfile& profile) noexcept
{
    if (! IsQuickConnectProfile(profile))
    {
        return;
    }

    RedSalamander::Connections::SetQuickConnectProfile(profile);

    if (! profile.savePassword)
    {
        RedSalamander::Connections::ClearQuickConnectSecret(SecretKind::Password);
        RedSalamander::Connections::ClearQuickConnectSecret(SecretKind::SshKeyPassphrase);
        return;
    }

    const bool sshPassphrase = profile.authMode == Common::Settings::ConnectionAuthMode::SshKey;
    const SecretKind kind    = sshPassphrase ? SecretKind::SshKeyPassphrase : SecretKind::Password;
    const auto& stagedMap    = sshPassphrase ? state.stagedPassphraseById : state.stagedPasswordById;
    const auto it            = stagedMap.find(profile.id);
    if (it == stagedMap.end() || it->second.empty())
    {
        return; // keep existing
    }

    RedSalamander::Connections::SetQuickConnectSecret(kind, it->second);
}

[[nodiscard]] bool SaveConnectionsSettings(HWND dlg, DialogState& state) noexcept
{
    if (! state.baselineSettings)
    {
        return true;
    }

    Common::Settings::ConnectionsSettings connSettings;
    if (state.baselineSettings->connections)
    {
        connSettings.bypassWindowsHello              = state.baselineSettings->connections->bypassWindowsHello;
        connSettings.windowsHelloReauthTimeoutMinute = state.baselineSettings->connections->windowsHelloReauthTimeoutMinute;
    }
    connSettings.items                  = state.connections;
    state.baselineSettings->connections = std::move(connSettings);

    DeleteSecretsForRemovedConnections(state);

    for (const auto& c : state.connections)
    {
        if (IsQuickConnectProfile(c))
        {
            CommitQuickConnectSecretsAndProfile(state, c);
            continue;
        }

        if (! ShouldCommitSecretsForProfile(state, c))
        {
            continue;
        }

        const HRESULT secretHr = CommitSecretsForProfile(state, c);
        if (FAILED(secretHr))
        {
            Debug::Error(L"CommitSecretsForProfile failed connection='{}' id='{}' hr=0x{:08X}", c.name, c.id, static_cast<unsigned long>(secretHr));

            const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
            const std::wstring details = FormatHresultForUi(secretHr);
            const std::wstring message =
                FormatStringResource(nullptr, IDS_CONNECTIONS_ERR_SAVE_CREDENTIAL_FAILED_FMT, c.name.empty() ? L"(unnamed)" : c.name, details);
            ShowDialogAlert(dlg, HOST_ALERT_ERROR, title, message);
            return false;
        }
    }

    PersistSettings(dlg, *state.baselineSettings, state.appId);
    return true;
}

void DeleteSecretsForRemovedConnections(const DialogState& state) noexcept
{
    std::unordered_set<std::wstring> currentIds;
    currentIds.reserve(state.connections.size());
    for (const auto& c : state.connections)
    {
        if (! c.id.empty())
        {
            currentIds.insert(c.id);
        }
    }

    for (const auto& id : state.baselineConnectionIds)
    {
        if (id.empty())
        {
            continue;
        }
        if (currentIds.contains(id))
        {
            continue;
        }

        const std::wstring passwordTarget   = BuildCredentialTargetName(id, SecretKind::Password);
        const std::wstring passphraseTarget = BuildCredentialTargetName(id, SecretKind::SshKeyPassphrase);
        if (! passwordTarget.empty())
        {
            static_cast<void>(DeleteGenericCredential(passwordTarget));
        }
        if (! passphraseTarget.empty())
        {
            static_cast<void>(DeleteGenericCredential(passphraseTarget));
        }
    }
}

std::optional<std::filesystem::path> BrowseForFile(HWND owner, const wchar_t* title)
{
    std::array<wchar_t, 2048> fileBuffer{};

    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = owner;
    ofn.lpstrFile   = fileBuffer.data();
    ofn.nMaxFile    = static_cast<DWORD>(fileBuffer.size());
    ofn.lpstrTitle  = title;
    ofn.Flags       = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_EXPLORER;

    if (! GetOpenFileNameW(&ofn))
    {
        return std::nullopt;
    }

    return std::filesystem::path(fileBuffer.data());
}

INT_PTR OnCtlColorDialog(DialogState* state) noexcept
{
    if (! state)
    {
        return FALSE;
    }
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnCtlColorStatic(DialogState* state, HDC hdc, HWND control) noexcept
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    COLORREF textColor = state->theme.menu.text;
    if (control && IsWindowEnabled(control) == FALSE)
    {
        textColor = state->theme.menu.disabledText;
    }

    if (! state->theme.highContrast)
    {
        // Combo box selection fields often paint via a child static window; match the input background.
        HWND parent = control ? GetParent(control) : nullptr;
        if (parent)
        {
            std::array<wchar_t, 32> className{};
            const int len = GetClassNameW(parent, className.data(), static_cast<int>(className.size()));
            if (len > 0 && (_wcsicmp(className.data(), L"ComboBox") == 0 || ThemedControls::IsModernComboBox(parent)))
            {
                const bool enabled = IsWindowEnabled(parent) != FALSE;
                const bool focused = enabled && (GetFocus() == parent || SendMessageW(parent, CB_GETDROPPEDSTATE, 0, 0) != 0);

                const COLORREF background =
                    enabled ? (focused ? state->inputFocusedBackgroundColor : state->inputBackgroundColor) : state->inputDisabledBackgroundColor;

                HBRUSH brush = state->backgroundBrush.get();
                if (enabled && focused && state->inputFocusedBrush)
                {
                    brush = state->inputFocusedBrush.get();
                }
                else if (enabled && state->inputBrush)
                {
                    brush = state->inputBrush.get();
                }
                else if (! enabled && state->inputDisabledBrush)
                {
                    brush = state->inputDisabledBrush.get();
                }

                const COLORREF comboText = enabled ? state->theme.menu.text : state->theme.menu.disabledText;
                SetBkMode(hdc, OPAQUE);
                SetBkColor(hdc, background);
                SetTextColor(hdc, comboText);
                return reinterpret_cast<INT_PTR>(brush);
            }
        }

        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, textColor);

        HBRUSH brush = state->backgroundBrush.get();
        if (control && state->cardBrush && ! state->cards.empty())
        {
            RECT rc{};
            if (GetWindowRect(control, &rc))
            {
                const HWND root = GetAncestor(control, GA_ROOT);
                if (root)
                {
                    MapWindowPoints(nullptr, root, reinterpret_cast<POINT*>(&rc), 2);
                    for (const RECT& card : state->cards)
                    {
                        RECT intersect{};
                        if (IntersectRect(&intersect, &card, &rc) != FALSE)
                        {
                            brush = state->cardBrush.get();
                            break;
                        }
                    }
                }
            }
        }

        return reinterpret_cast<INT_PTR>(brush);
    }

    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, state->theme.windowBackground);
    SetTextColor(hdc, textColor);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnCtlColorEdit(DialogState* state, HDC hdc, HWND control) noexcept
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    const bool enabled = ! control || IsWindowEnabled(control) != FALSE;
    const bool focused = enabled && control && GetFocus() == control;
    const COLORREF bg  = enabled ? (focused ? state->inputFocusedBackgroundColor : state->inputBackgroundColor) : state->inputDisabledBackgroundColor;
    SetBkColor(hdc, bg);
    SetTextColor(hdc, enabled ? state->theme.menu.text : state->theme.menu.disabledText);

    if (state->theme.highContrast)
    {
        return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
    }

    if (! enabled)
    {
        return reinterpret_cast<INT_PTR>(state->inputDisabledBrush.get());
    }
    return reinterpret_cast<INT_PTR>(focused && state->inputFocusedBrush ? state->inputFocusedBrush.get() : state->inputBrush.get());
}

INT_PTR OnCtlColorButton(DialogState* state, HDC hdc, HWND control) noexcept
{
    if (! state || ! hdc)
    {
        return FALSE;
    }

    const bool enabled              = ! control || IsWindowEnabled(control) != FALSE;
    const COLORREF windowBackground = state->theme.windowBackground;
    COLORREF background             = windowBackground;
    HBRUSH brush                    = state->backgroundBrush.get();

    if (! state->theme.highContrast && control && state->cardBrush && ! state->cards.empty())
    {
        RECT rc{};
        if (GetWindowRect(control, &rc))
        {
            const HWND root = GetAncestor(control, GA_ROOT);
            if (root)
            {
                MapWindowPoints(nullptr, root, reinterpret_cast<POINT*>(&rc), 2);
                for (const RECT& card : state->cards)
                {
                    RECT intersect{};
                    if (IntersectRect(&intersect, &card, &rc) != FALSE)
                    {
                        background = state->cardBackgroundColor;
                        brush      = state->cardBrush.get();
                        break;
                    }
                }
            }
        }
    }

    SetBkMode(hdc, OPAQUE);
    SetBkColor(hdc, background);
    SetTextColor(hdc, enabled ? state->theme.menu.text : state->theme.menu.disabledText);
    return reinterpret_cast<INT_PTR>(brush);
}

INT_PTR OnListCustomDraw(DialogState* state, NMLVCUSTOMDRAW* cd) noexcept
{
    if (! state || ! cd)
    {
        return CDRF_DODEFAULT;
    }

    if (cd->nmcd.dwDrawStage == CDDS_PREPAINT)
    {
        return CDRF_NOTIFYITEMDRAW;
    }

    if (cd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
    {
        const bool selected = (cd->nmcd.uItemState & CDIS_SELECTED) != 0;
        cd->clrText         = selected ? state->theme.menu.selectionText : state->theme.menu.text;
        cd->clrTextBk       = selected ? state->theme.menu.selectionBg : state->theme.windowBackground;
        return CDRF_DODEFAULT;
    }

    return CDRF_DODEFAULT;
}

void PaintDialogBackgroundAndCards(HDC hdc, HWND dlg, const DialogState& state) noexcept
{
    if (! hdc || ! dlg || ! state.backgroundBrush)
    {
        return;
    }

    RECT rc{};
    if (! GetClientRect(dlg, &rc))
    {
        return;
    }

    FillRect(hdc, &rc, state.backgroundBrush.get());

    if (state.theme.highContrast || state.cards.empty())
    {
        return;
    }

    const UINT dpi         = GetDpiForWindow(dlg);
    const int radius       = ThemedControls::ScaleDip(dpi, 6);
    const COLORREF surface = state.cardBackgroundColor;
    const COLORREF border  = BlendColor(surface, state.theme.menu.text, state.theme.dark ? 40 : 30, 255);

    wil::unique_hbrush cardBrush(CreateSolidBrush(surface));
    wil::unique_hpen cardPen(CreatePen(PS_SOLID, 1, border));
    if (! cardBrush || ! cardPen)
    {
        return;
    }

    [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc, cardBrush.get());
    [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc, cardPen.get());

    const int savedDc = SaveDC(hdc);
    if (state.settingsViewport.right > state.settingsViewport.left && state.settingsViewport.bottom > state.settingsViewport.top)
    {
        static_cast<void>(
            IntersectClipRect(hdc, state.settingsViewport.left, state.settingsViewport.top, state.settingsViewport.right, state.settingsViewport.bottom));
    }

    for (const RECT& card : state.cards)
    {
        if (card.right <= card.left || card.bottom <= card.top)
        {
            continue;
        }

        RoundRect(hdc, card.left, card.top, card.right, card.bottom, radius, radius);
    }

    if (savedDc != 0)
    {
        RestoreDC(hdc, savedDc);
    }
}

constexpr wchar_t kConnectionsSettingsHostClassName[] = L"RedSalamanderConnectionsSettingsHost";

void LayoutDialog(HWND dlg, DialogState& state) noexcept;
LRESULT CALLBACK ConnectionsSettingsHostProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

[[nodiscard]] bool EnsureConnectionsSettingsHostClassRegistered() noexcept
{
    const HINSTANCE instance = GetModuleHandleW(nullptr);

    WNDCLASSEXW existing{};
    existing.cbSize = sizeof(existing);
    if (GetClassInfoExW(instance, kConnectionsSettingsHostClassName, &existing) != 0)
    {
        return true;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_DBLCLKS;
    wc.lpfnWndProc   = ConnectionsSettingsHostProc;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.lpszClassName = kConnectionsSettingsHostClassName;

    const ATOM atom = RegisterClassExW(&wc);
    return atom != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

[[nodiscard]] HWND FindFirstOrLastTabStopChild(HWND root, bool forward) noexcept
{
    if (! root)
    {
        return nullptr;
    }

    const HWND dlg = GetAncestor(root, GA_ROOT);
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
        if (IsChild(root, item) && IsWindowVisible(item) && IsWindowEnabled(item))
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

void PaintSettingsHostBackgroundAndCards(HDC hdc, HWND host, const DialogState& state) noexcept
{
    if (! hdc || ! host || ! state.backgroundBrush)
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(host, &client))
    {
        return;
    }

    FillRect(hdc, &client, state.backgroundBrush.get());

    if (state.theme.highContrast || state.cards.empty())
    {
        return;
    }

    const UINT dpi         = GetDpiForWindow(host);
    const int radius       = ThemedControls::ScaleDip(dpi, 6);
    const COLORREF surface = state.cardBackgroundColor;
    const COLORREF border  = BlendColor(surface, state.theme.menu.text, state.theme.dark ? 40 : 30, 255);

    wil::unique_hbrush cardBrush(CreateSolidBrush(surface));
    wil::unique_hpen cardPen(CreatePen(PS_SOLID, 1, border));
    if (! cardBrush || ! cardPen)
    {
        return;
    }

    [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc, cardBrush.get());
    [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc, cardPen.get());

    for (const RECT& cardInDialog : state.cards)
    {
        RECT card = cardInDialog;
        OffsetRect(&card, -state.settingsViewport.left, -state.settingsViewport.top);
        if (card.right <= card.left || card.bottom <= card.top)
        {
            continue;
        }

        RoundRect(hdc, card.left, card.top, card.right, card.bottom, radius, radius);
    }
}

LRESULT CALLBACK ConnectionsSettingsHostProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* state    = reinterpret_cast<DialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    const HWND dlg = GetParent(hwnd);

    switch (msg)
    {
        case WM_ERASEBKGND:
            if (state && wp)
            {
                PaintSettingsHostBackgroundAndCards(reinterpret_cast<HDC>(wp), hwnd, *state);
                return 1;
            }
            return 1;
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
                if (state)
                {
                    PaintSettingsHostBackgroundAndCards(memDc.get(), hwnd, *state);
                }
                BitBlt(hdc.get(), 0, 0, width, height, memDc.get(), 0, 0, SRCCOPY);
            }
            else if (state)
            {
                PaintSettingsHostBackgroundAndCards(hdc.get(), hwnd, *state);
            }

            return 0;
        }
        case WM_VSCROLL:
        {
            if (! state || state->settingsScrollMax <= 0)
            {
                break;
            }

            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_ALL;
            GetScrollInfo(hwnd, SB_VERT, &si);

            const UINT dpi  = GetDpiForWindow(hwnd);
            const int lineY = ScaleDip(dpi, 24);

            int newPos = state->settingsScrollOffset;
            switch (LOWORD(wp))
            {
                case SB_TOP: newPos = 0; break;
                case SB_BOTTOM: newPos = state->settingsScrollMax; break;
                case SB_LINEUP: newPos -= lineY; break;
                case SB_LINEDOWN: newPos += lineY; break;
                case SB_PAGEUP: newPos -= static_cast<int>(si.nPage); break;
                case SB_PAGEDOWN: newPos += static_cast<int>(si.nPage); break;
                case SB_THUMBTRACK: newPos = si.nTrackPos; break;
                case SB_THUMBPOSITION: newPos = si.nPos; break;
                default: break;
            }

            newPos = std::clamp(newPos, 0, state->settingsScrollMax);
            if (newPos != state->settingsScrollOffset && dlg)
            {
                state->settingsScrollOffset = newPos;
                LayoutDialog(dlg, *state);
            }
            return 0;
        }
        case WM_MOUSEWHEEL:
        {
            if (! state || state->settingsScrollMax <= 0 || ! dlg)
            {
                break;
            }

            const int delta = GET_WHEEL_DELTA_WPARAM(wp);
            if (delta == 0)
            {
                return 0;
            }

            UINT linesPerNotch = 3;
            SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &linesPerNotch, 0);
            if (linesPerNotch == 0)
            {
                return 0;
            }

            const UINT dpi  = GetDpiForWindow(hwnd);
            const int lineY = ScaleDip(dpi, 32);

            int scrollDelta = 0;
            if (linesPerNotch == WHEEL_PAGESCROLL)
            {
                SCROLLINFO si{};
                si.cbSize = sizeof(si);
                si.fMask  = SIF_PAGE;
                GetScrollInfo(hwnd, SB_VERT, &si);
                scrollDelta = (delta / WHEEL_DELTA) * static_cast<int>(si.nPage);
            }
            else
            {
                scrollDelta = (delta / WHEEL_DELTA) * lineY * static_cast<int>(linesPerNotch);
            }

            const int newPos = std::clamp(state->settingsScrollOffset - scrollDelta, 0, state->settingsScrollMax);
            if (newPos != state->settingsScrollOffset)
            {
                state->settingsScrollOffset = newPos;
                LayoutDialog(dlg, *state);
            }
            return 0;
        }
        case WM_COMMAND:
        case WM_NOTIFY:
        case WM_DRAWITEM:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORBTN:
            if (dlg)
            {
                return SendMessageW(dlg, msg, wp, lp);
            }
            break;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void LayoutDialog(HWND dlg, DialogState& state) noexcept
{
    if (! dlg)
    {
        return;
    }

    RECT rc{};
    if (! GetClientRect(dlg, &rc))
    {
        return;
    }

    const UINT dpi         = GetDpiForWindow(dlg);
    const int margin       = ScaleDip(dpi, 12);
    const int gapX         = ScaleDip(dpi, 12);
    const int gapY         = ScaleDip(dpi, 10);
    const int rowHeight    = ScaleDip(dpi, 28);
    const int headerHeight = ScaleDip(dpi, 18);
    const int sectionGapY  = ScaleDip(dpi, 6);
    const int cardPaddingX = ScaleDip(dpi, 12);
    const int cardPaddingY = ScaleDip(dpi, 10);
    const int cardSpacingY = ScaleDip(dpi, 12);
    const int framePadding = ScaleDip(dpi, 2);
    const int labelWidth   = ScaleDip(dpi, 100);

    const int clientW = std::max(0l, rc.right - rc.left);
    const int clientH = std::max(0l, rc.bottom - rc.top);

    HFONT dialogFont       = GetDialogFont(dlg);
    const HFONT headerFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const HFONT titleFont  = state.titleFont ? state.titleFont.get() : headerFont;

    const std::wstring connectText = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_CONNECT);
    const std::wstring closeText   = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_CLOSE);
    const std::wstring cancelText  = LoadStringResource(nullptr, IDS_BTN_CANCEL);
    const std::wstring newText     = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_NEW_ELLIPSIS);
    const std::wstring renameText  = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_RENAME_ELLIPSIS);
    const std::wstring removeText  = LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_REMOVE);
    const int buttonPadX           = ScaleDip(dpi, 16);
    const int connectW             = std::max(ScaleDip(dpi, 90), buttonPadX + ThemedControls::MeasureTextWidth(dlg, dialogFont, connectText) + buttonPadX);
    const int closeW               = std::max(ScaleDip(dpi, 90), buttonPadX + ThemedControls::MeasureTextWidth(dlg, dialogFont, closeText) + buttonPadX);
    const int cancelW              = std::max(ScaleDip(dpi, 90), buttonPadX + ThemedControls::MeasureTextWidth(dlg, dialogFont, cancelText) + buttonPadX);

    const int bottomButtonsY = std::max(0, clientH - margin - rowHeight);
    const int cancelX        = std::max(0, clientW - margin - cancelW);
    const int closeX         = std::max(0, cancelX - gapX - closeW);
    const int connectX       = std::max(0, closeX - gapX - connectW);

    if (const HWND ok = GetDlgItem(dlg, IDOK))
    {
        SetWindowPos(ok, nullptr, connectX, bottomButtonsY, connectW, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(ok, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    if (const HWND closeBtn = GetDlgItem(dlg, IDC_CONNECTION_CLOSE))
    {
        SetWindowPos(closeBtn, nullptr, closeX, bottomButtonsY, closeW, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(closeBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    if (const HWND cancel = GetDlgItem(dlg, IDCANCEL))
    {
        SetWindowPos(cancel, nullptr, cancelX, bottomButtonsY, cancelW, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(cancel, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    const int listButtonY = bottomButtonsY;

    const auto measureButtonWidth = [&](std::wstring_view text) noexcept
    {
        const int textW = ThemedControls::MeasureTextWidth(dlg, dialogFont, text);
        return std::max(ScaleDip(dpi, 80), (2 * buttonPadX) + textW);
    };

    const int listBtnMinW = std::max(measureButtonWidth(newText), std::max(measureButtonWidth(renameText), measureButtonWidth(removeText)));
    const int listMinW    = ScaleDip(dpi, 180);

    int listWidth = std::max(listMinW, (3 * listBtnMinW) + (2 * gapX));

    const int portWidth    = ScaleDip(dpi, 90);
    const int portLabelW   = ScaleDip(dpi, 40);
    const int minHostWidth = ScaleDip(dpi, 140);
    const int minRightW    = (2 * cardPaddingX) + labelWidth + gapX + minHostWidth + gapX + portLabelW + gapX + portWidth;
    const int maxListW     = std::max(listMinW, std::max(0, clientW - (2 * margin) - gapX - minRightW));
    listWidth              = std::min(listWidth, maxListW);

    const int listTitleHeight = ScaleDip(dpi, 40);
    const int listTitleGapY   = ScaleDip(dpi, 8);
    const int listTop         = margin + listTitleHeight + listTitleGapY;
    const int listHeight      = std::max(0, listButtonY - gapY - listTop);

    if (state.listTitle)
    {
        SetWindowPos(state.listTitle, nullptr, margin, margin, listWidth, listTitleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.listTitle, WM_SETFONT, reinterpret_cast<WPARAM>(titleFont), TRUE);
    }

    if (state.list)
    {
        SetWindowPos(state.list, nullptr, margin, listTop, listWidth, listHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.list, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        RECT listClient{};
        if (GetClientRect(state.list, &listClient))
        {
            const int colWidth = std::max(0l, listClient.right - listClient.left - ScaleDip(dpi, 2));
            ListView_SetColumnWidth(state.list, 0, colWidth);
        }
    }

    const int listBtnW = std::max(1, (listWidth - 2 * gapX) / 3);
    if (const HWND btnNew = GetDlgItem(dlg, IDC_CONNECTION_NEW))
    {
        SetWindowPos(btnNew, nullptr, margin, listButtonY, listBtnW, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(btnNew, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    if (const HWND btnRename = GetDlgItem(dlg, IDC_CONNECTION_RENAME))
    {
        SetWindowPos(btnRename, nullptr, margin + listBtnW + gapX, listButtonY, listBtnW, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(btnRename, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }
    if (const HWND btnRemove = GetDlgItem(dlg, IDC_CONNECTION_REMOVE))
    {
        const int x = margin + 2 * (listBtnW + gapX);
        const int w = std::max(1, listWidth - (x - margin));
        SetWindowPos(btnRemove, nullptr, x, listButtonY, w, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(btnRemove, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
    }

    const std::wstring_view selectedPluginId = [&]() noexcept -> std::wstring_view
    {
        const auto modelIndex = GetSelectedModelIndex(state);
        if (! modelIndex.has_value() || modelIndex.value() >= state.connections.size())
        {
            return {};
        }

        return state.connections[modelIndex.value()].pluginId;
    }();

    const bool isAwsS3Selection = IsAwsS3PluginId(selectedPluginId);
    const bool isS3Selection    = IsS3PluginId(selectedPluginId);
    const bool showSshCard      = IsSshPluginId(selectedPluginId);

    const int viewportTop    = margin;
    const int viewportBottom = std::max(viewportTop, listButtonY - gapY);
    const int viewportHeight = std::max(0, viewportBottom - viewportTop);

    const auto estimateCardBlockHeight = [&](int rows) noexcept
    {
        if (rows <= 0)
        {
            return 0;
        }

        const int cardHeight = (2 * cardPaddingY) + (rows * rowHeight) + ((rows - 1) * gapY);
        return headerHeight + sectionGapY + cardHeight + cardSpacingY;
    };

    const int connectionRows      = state.filterPluginId.empty() ? 4 : 3;
    const int authRowsForEstimate = 5;
    const int s3Rows              = 1 /* endpoint */ + 2 /* useHttps + verifyTls */ + (isS3Selection ? 1 : 0);
    const int sshRows             = 2;

    int estimatedContentHeight = 0;
    estimatedContentHeight += estimateCardBlockHeight(connectionRows);
    estimatedContentHeight += estimateCardBlockHeight(authRowsForEstimate);
    if (isAwsS3Selection)
    {
        estimatedContentHeight += estimateCardBlockHeight(s3Rows);
    }
    if (showSshCard)
    {
        estimatedContentHeight += estimateCardBlockHeight(sshRows);
    }

    const int rightX         = margin + listWidth + gapX;
    const int rightWidthFull = std::max(0, clientW - rightX - margin);

    if (state.settingsHost)
    {
        const bool wantsVScroll = viewportHeight > 0 && estimatedContentHeight > viewportHeight;

        LONG_PTR exStyle = GetWindowLongPtrW(state.settingsHost, GWL_EXSTYLE);
        if ((exStyle & WS_EX_CONTROLPARENT) == 0)
        {
            exStyle |= WS_EX_CONTROLPARENT;
            SetWindowLongPtrW(state.settingsHost, GWL_EXSTYLE, exStyle);
        }

        LONG_PTR styleNow    = GetWindowLongPtrW(state.settingsHost, GWL_STYLE);
        LONG_PTR styleWanted = styleNow;
        styleWanted |= WS_TABSTOP;
        styleWanted &= ~WS_HSCROLL;
        if (wantsVScroll)
        {
            styleWanted |= WS_VSCROLL;
        }
        else
        {
            styleWanted &= ~WS_VSCROLL;
        }

        const bool styleChanged = (styleWanted != styleNow);
        if (styleChanged)
        {
            SetWindowLongPtrW(state.settingsHost, GWL_STYLE, styleWanted);
        }

        SetWindowPos(state.settingsHost,
                     nullptr,
                     rightX,
                     viewportTop,
                     rightWidthFull,
                     viewportHeight,
                     SWP_NOZORDER | SWP_NOACTIVATE | (styleChanged ? SWP_FRAMECHANGED : 0u));

        if (styleChanged)
        {
            if (state.theme.highContrast)
            {
                SetWindowTheme(state.settingsHost, L"", nullptr);
            }
            else
            {
                const wchar_t* hostTheme = state.theme.dark ? L"DarkMode_Explorer" : L"Explorer";
                SetWindowTheme(state.settingsHost, hostTheme, nullptr);
            }
            SendMessageW(state.settingsHost, WM_THEMECHANGED, 0, 0);
            RedrawWindow(state.settingsHost, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);
        }
    }

    RECT hostClient{};
    if (! state.settingsHost || ! GetClientRect(state.settingsHost, &hostClient))
    {
        return;
    }

    const int rightWidth = std::max(0l, hostClient.right - hostClient.left);
    int cardY            = margin;

    state.settingsViewport.left   = rightX;
    state.settingsViewport.top    = viewportTop;
    state.settingsViewport.right  = rightX + rightWidth;
    state.settingsViewport.bottom = viewportBottom;

    state.settingsScrollMax    = (viewportHeight > 0) ? std::max(0, estimatedContentHeight - viewportHeight) : 0;
    state.settingsScrollOffset = std::clamp(state.settingsScrollOffset, 0, state.settingsScrollMax);
    if (state.settingsScrollMax <= 0)
    {
        state.settingsScrollOffset = 0;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = std::max(0, estimatedContentHeight - 1);
    si.nPage  = (viewportHeight > 0) ? static_cast<UINT>(viewportHeight) : 0u;
    si.nPos   = state.settingsScrollOffset;
    SetScrollInfo(state.settingsHost, SB_VERT, &si, TRUE);

    const int scrollOffset = state.settingsScrollOffset;

    state.cards.clear();

    const auto positionScrollable = [&](HWND hwnd, int x, int y, int w, int h) noexcept
    {
        if (! hwnd)
        {
            return;
        }

        const int relX = x - rightX;
        const int relY = (y - viewportTop) - scrollOffset;
        SetWindowPos(hwnd, nullptr, relX, relY, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
    };

    auto pushCard = [&](int cardHeight) noexcept
    {
        RECT card{};
        card.left   = rightX;
        card.top    = cardY;
        card.right  = rightX + rightWidth;
        card.bottom = cardY + cardHeight;

        RECT paint = card;
        paint.top -= scrollOffset;
        paint.bottom -= scrollOffset;

        state.cards.push_back(paint);

        cardY += cardHeight + cardSpacingY;
        return card;
    };

    const auto positionLabel = [&](HWND label, HFONT font, int x, int y, int w, int h, std::wstring_view text) noexcept
    {
        if (! label)
        {
            return;
        }

        if (! text.empty())
        {
            SetWindowTextW(label, std::wstring(text).c_str());
        }

        positionScrollable(label, x, y, w, h);
        SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    };

    const auto positionFramed = [&](const wil::unique_hwnd& frame, HWND input, int x, int y, int w) noexcept
    {
        if (frame)
        {
            positionScrollable(frame.get(), x, y, w, rowHeight);
        }
        if (input)
        {
            positionScrollable(input, x + framePadding, y + framePadding, std::max(1, w - 2 * framePadding), std::max(1, rowHeight - 2 * framePadding));
            SendMessageW(input, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }
    };

    const auto positionToggleRow = [&](HWND label, HWND toggle, std::wstring_view labelText, const RECT& card, int& y) noexcept
    {
        if (! label || ! toggle)
        {
            return;
        }

        const int minToggleWidth      = ScaleDip(dpi, 90);
        const int paddingX            = ScaleDip(dpi, 6);
        const int toggleGapX          = ScaleDip(dpi, 8);
        const int trackWidth          = ScaleDip(dpi, 34);
        const int stateTextWidth      = std::max(ThemedControls::MeasureTextWidth(dlg, headerFont, state.toggleOnLabel),
                                            ThemedControls::MeasureTextWidth(dlg, headerFont, state.toggleOffLabel));
        const int measuredToggleWidth = std::max(minToggleWidth, (2 * paddingX) + stateTextWidth + toggleGapX + trackWidth);
        const int cardWidth           = std::max(0l, card.right - card.left);
        const int toggleWidth         = std::min(std::max(0, cardWidth - 2 * cardPaddingX), measuredToggleWidth);

        positionLabel(label,
                      dialogFont,
                      card.left + cardPaddingX,
                      y + (rowHeight - headerHeight) / 2,
                      std::max(0, cardWidth - 2 * cardPaddingX - toggleWidth - gapX),
                      headerHeight,
                      labelText);

        positionScrollable(toggle, card.right - cardPaddingX - toggleWidth, y, toggleWidth, rowHeight);
        SendMessageW(toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);

        y += rowHeight + gapY;
    };

    const auto pushSectionHeader = [&](HWND label, std::wstring_view text) noexcept
    {
        positionLabel(label, headerFont, rightX + cardPaddingX, cardY, std::max(0, rightWidth - 2 * cardPaddingX), headerHeight, text);
        cardY += headerHeight + sectionGapY;
    };

    const std::wstring hostLabelText = LoadStringResource(nullptr, isAwsS3Selection ? IDS_CONNECTIONS_LABEL_REGION : IDS_CONNECTIONS_LABEL_HOST);
    const std::wstring userLabelText = LoadStringResource(nullptr, isAwsS3Selection ? IDS_CONNECTIONS_LABEL_ACCESS_KEY_ID : IDS_CONNECTIONS_LABEL_USER);

    // Connection card
    {
        pushSectionHeader(state.sectionConnection, LoadStringResource(nullptr, IDS_CONNECTIONS_SECTION_CONNECTION));

        const int rows       = state.filterPluginId.empty() ? 4 : 3;
        const int cardHeight = (2 * cardPaddingY) + (rows * rowHeight) + ((rows - 1) * gapY);
        const RECT card      = pushCard(cardHeight);

        int cy = card.top + cardPaddingY;

        positionLabel(state.nameLabel,
                      dialogFont,
                      card.left + cardPaddingX,
                      cy + (rowHeight - headerHeight) / 2,
                      labelWidth,
                      headerHeight,
                      LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_NAME));
        positionFramed(
            state.nameFrame, state.nameEdit, card.left + cardPaddingX + labelWidth + gapX, cy, std::max(0, rightWidth - 2 * cardPaddingX - labelWidth - gapX));
        cy += rowHeight + gapY;

        if (state.filterPluginId.empty())
        {
            positionLabel(state.protocolLabel,
                          dialogFont,
                          card.left + cardPaddingX,
                          cy + (rowHeight - headerHeight) / 2,
                          labelWidth,
                          headerHeight,
                          LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PROTOCOL));
            positionFramed(state.protocolFrame, state.protocolCombo, card.left + cardPaddingX + labelWidth + gapX, cy, std::max(0, ScaleDip(dpi, 180)));
            cy += rowHeight + gapY;
        }

        positionLabel(state.hostLabel, dialogFont, card.left + cardPaddingX, cy + (rowHeight - headerHeight) / 2, labelWidth, headerHeight, hostLabelText);

        const int hostWidth = isAwsS3Selection ? std::max(0, rightWidth - 2 * cardPaddingX - labelWidth - gapX)
                                               : std::max(0, rightWidth - 2 * cardPaddingX - labelWidth - gapX - gapX - portLabelW - gapX - portWidth);
        const int hostX     = card.left + cardPaddingX + labelWidth + gapX;
        if (isAwsS3Selection)
        {
            positionFramed(state.awsRegionFrame, state.awsRegionCombo, hostX, cy, hostWidth);
            if (state.awsRegionCombo && ! ThemedControls::IsModernComboBox(state.awsRegionCombo))
            {
                const int editW      = std::max(1, hostWidth - 2 * framePadding);
                const int dropHeight = ScaleDip(dpi, 260);
                SetWindowPos(
                    state.awsRegionCombo, nullptr, 0, 0, editW, std::max(dropHeight, rowHeight - 2 * framePadding), SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOMOVE);
            }
        }
        else
        {
            positionFramed(state.hostFrame, state.hostEdit, hostX, cy, hostWidth);
        }

        if (! isAwsS3Selection)
        {
            positionLabel(state.portLabel,
                          dialogFont,
                          hostX + hostWidth + gapX,
                          cy + (rowHeight - headerHeight) / 2,
                          portLabelW,
                          headerHeight,
                          LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_PORT));
            positionFramed(state.portFrame, state.portEdit, hostX + hostWidth + gapX + portLabelW + gapX, cy, portWidth);
        }
        cy += rowHeight + gapY;

        positionLabel(state.initialPathLabel,
                      dialogFont,
                      card.left + cardPaddingX,
                      cy + (rowHeight - headerHeight) / 2,
                      labelWidth,
                      headerHeight,
                      LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_INITIAL_PATH));
        positionFramed(state.initialPathFrame,
                       state.initialPathEdit,
                       card.left + cardPaddingX + labelWidth + gapX,
                       cy,
                       std::max(0, rightWidth - 2 * cardPaddingX - labelWidth - gapX));
    }

    // Auth card
    {
        pushSectionHeader(state.sectionAuth, LoadStringResource(nullptr, IDS_CONNECTIONS_SECTION_AUTH));

        const int baseRows   = 4;            // user, secret, save, protocol-specific bool
        const int authRows   = baseRows + 1; // include anonymous row (hidden when not ftp)
        const int cardHeight = (2 * cardPaddingY) + (authRows * rowHeight) + ((authRows - 1) * gapY);
        const RECT card      = pushCard(cardHeight);

        int cy = card.top + cardPaddingY;

        positionToggleRow(state.anonymousLabel, state.anonymousToggle, LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_ANONYMOUS), card, cy);

        positionLabel(state.userLabel, dialogFont, card.left + cardPaddingX, cy + (rowHeight - headerHeight) / 2, labelWidth, headerHeight, userLabelText);
        positionFramed(
            state.userFrame, state.userEdit, card.left + cardPaddingX + labelWidth + gapX, cy, std::max(0, rightWidth - 2 * cardPaddingX - labelWidth - gapX));
        cy += rowHeight + gapY;

        // secret label text updated in UpdateControlEnabledState
        positionLabel(state.secretLabel, dialogFont, card.left + cardPaddingX, cy + (rowHeight - headerHeight) / 2, labelWidth, headerHeight, {});
        const int showSecretW = ScaleDip(dpi, 60);
        const int secretEditW = std::max(0, rightWidth - 2 * cardPaddingX - labelWidth - gapX - gapX - showSecretW);
        const int secretEditX = card.left + cardPaddingX + labelWidth + gapX;
        positionFramed(state.secretFrame, state.secretEdit, secretEditX, cy, secretEditW);
        if (state.showSecretBtn)
        {
            positionScrollable(state.showSecretBtn, secretEditX + secretEditW + gapX, cy, showSecretW, rowHeight);
            SendMessageW(state.showSecretBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }
        cy += rowHeight + gapY;

        positionToggleRow(state.savePasswordLabel, state.savePasswordToggle, LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_SAVE_PASSWORD), card, cy);
        positionToggleRow(state.ignoreSslTrustLabel, state.ignoreSslTrustToggle, LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_IGNORE_SSL_TRUST), card, cy);
    }

    if (isAwsS3Selection)
    {
        // S3 card
        pushSectionHeader(state.sectionS3, LoadStringResource(nullptr, IDS_CONNECTIONS_SECTION_S3));

        const int rows       = 1 /* endpoint */ + 2 /* useHttps + verifyTls */ + (isS3Selection ? 1 : 0);
        const int cardHeight = (2 * cardPaddingY) + (rows * rowHeight) + ((rows - 1) * gapY);
        const RECT card      = pushCard(cardHeight);

        int cy = card.top + cardPaddingY;

        positionLabel(state.s3EndpointOverrideLabel,
                      dialogFont,
                      card.left + cardPaddingX,
                      cy + (rowHeight - headerHeight) / 2,
                      labelWidth,
                      headerHeight,
                      LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_ENDPOINT_OVERRIDE));

        positionFramed(state.s3EndpointOverrideFrame,
                       state.s3EndpointOverrideEdit,
                       card.left + cardPaddingX + labelWidth + gapX,
                       cy,
                       std::max(0, rightWidth - 2 * cardPaddingX - labelWidth - gapX));
        cy += rowHeight + gapY;

        positionToggleRow(state.s3UseHttpsLabel, state.s3UseHttpsToggle, LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_USE_HTTPS), card, cy);
        positionToggleRow(state.s3VerifyTlsLabel, state.s3VerifyTlsToggle, LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_VERIFY_TLS), card, cy);
        if (isS3Selection)
        {
            positionToggleRow(state.s3UseVirtualAddressingLabel,
                              state.s3UseVirtualAddressingToggle,
                              LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_USE_VIRTUAL_ADDRESSING),
                              card,
                              cy);
        }
    }

    if (showSshCard)
    {
        // SSH card
        pushSectionHeader(state.sectionSsh, LoadStringResource(nullptr, IDS_CONNECTIONS_SECTION_SSH));

        const int rows       = 2;
        const int cardHeight = (2 * cardPaddingY) + (rows * rowHeight) + ((rows - 1) * gapY);
        const RECT card      = pushCard(cardHeight);

        int cy = card.top + cardPaddingY;

        const int browseW = rowHeight;
        const int editW   = std::max(0, rightWidth - 2 * cardPaddingX - labelWidth - gapX - gapX - browseW);
        const int editX   = card.left + cardPaddingX + labelWidth + gapX;

        positionLabel(state.sshPrivateKeyLabel,
                      dialogFont,
                      card.left + cardPaddingX,
                      cy + (rowHeight - headerHeight) / 2,
                      labelWidth,
                      headerHeight,
                      LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_SSH_PRIVATEKEY));
        positionFramed(state.sshPrivateKeyFrame, state.sshPrivateKeyEdit, editX, cy, editW);
        if (state.sshPrivateKeyBrowseBtn)
        {
            positionScrollable(state.sshPrivateKeyBrowseBtn, editX + editW + gapX, cy, browseW, rowHeight);
            SendMessageW(state.sshPrivateKeyBrowseBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }
        cy += rowHeight + gapY;

        positionLabel(state.sshKnownHostsLabel,
                      dialogFont,
                      card.left + cardPaddingX,
                      cy + (rowHeight - headerHeight) / 2,
                      labelWidth,
                      headerHeight,
                      LoadStringResource(nullptr, IDS_CONNECTIONS_LABEL_SSH_KNOWNHOSTS));
        positionFramed(state.sshKnownHostsFrame, state.sshKnownHostsEdit, editX, cy, editW);
        if (state.sshKnownHostsBrowseBtn)
        {
            positionScrollable(state.sshKnownHostsBrowseBtn, editX + editW + gapX, cy, browseW, rowHeight);
            SendMessageW(state.sshKnownHostsBrowseBtn, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }
    }

    if (state.settingsHost)
    {
        InvalidateRect(state.settingsHost, nullptr, TRUE);
    }
    InvalidateRect(dlg, nullptr, TRUE);
}

INT_PTR OnCommand(HWND dlg, DialogState& state, int controlId) noexcept;

INT_PTR OnInitDialog(HWND dlg, DialogState* init) noexcept
{
    if (! init)
    {
        return FALSE;
    }

    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(init));

    EnsureControls(*init, dlg);
    UpdateSecretVisibility(*init);

    if (! init->settingsHost)
    {
        static_cast<void>(EnsureConnectionsSettingsHostClassRegistered());

        init->settingsHost = CreateWindowExW(WS_EX_CONTROLPARENT,
                                             kConnectionsSettingsHostClassName,
                                             nullptr,
                                             WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                                             0,
                                             0,
                                             0,
                                             0,
                                             dlg,
                                             reinterpret_cast<HMENU>(static_cast<intptr_t>(IDC_CONNECTION_SETTINGS_SCROLL)),
                                             GetModuleHandleW(nullptr),
                                             nullptr);
        if (init->settingsHost)
        {
            SetWindowLongPtrW(init->settingsHost, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(init));

            if (init->theme.highContrast)
            {
                SetWindowTheme(init->settingsHost, L"", nullptr);
            }
            else
            {
                const wchar_t* hostTheme = init->theme.dark ? L"DarkMode_Explorer" : L"Explorer";
                SetWindowTheme(init->settingsHost, hostTheme, nullptr);
            }
            SendMessageW(init->settingsHost, WM_THEMECHANGED, 0, 0);
        }
    }

    if (init->settingsHost)
    {
        const auto reparent = [&](HWND hwnd) noexcept
        {
            if (hwnd)
            {
                SetParent(hwnd, init->settingsHost);
            }
        };

        // Right pane (scrollable settings editor) controls.
        reparent(init->sectionConnection);
        reparent(init->sectionAuth);
        reparent(init->sectionS3);
        reparent(init->sectionSsh);

        reparent(init->nameLabel);
        reparent(init->protocolLabel);
        reparent(init->hostLabel);
        reparent(init->portLabel);
        reparent(init->initialPathLabel);
        reparent(init->anonymousLabel);
        reparent(init->userLabel);
        reparent(init->secretLabel);
        reparent(init->savePasswordLabel);
        reparent(init->requireHelloLabel);
        reparent(init->ignoreSslTrustLabel);
        reparent(init->s3EndpointOverrideLabel);
        reparent(init->s3UseHttpsLabel);
        reparent(init->s3VerifyTlsLabel);
        reparent(init->s3UseVirtualAddressingLabel);
        reparent(init->sshPrivateKeyLabel);
        reparent(init->sshKnownHostsLabel);

        reparent(init->nameEdit);
        reparent(init->protocolCombo);
        reparent(init->hostEdit);
        reparent(init->portEdit);
        reparent(init->initialPathEdit);
        reparent(init->anonymousToggle);
        reparent(init->userEdit);
        reparent(init->secretEdit);
        reparent(init->showSecretBtn);
        reparent(init->savePasswordToggle);
        reparent(init->requireHelloToggle);
        reparent(init->ignoreSslTrustToggle);
        reparent(init->s3EndpointOverrideEdit);
        reparent(init->s3UseHttpsToggle);
        reparent(init->s3VerifyTlsToggle);
        reparent(init->s3UseVirtualAddressingToggle);
        reparent(init->sshPrivateKeyEdit);
        reparent(init->sshPrivateKeyBrowseBtn);
        reparent(init->sshKnownHostsEdit);
        reparent(init->sshKnownHostsBrowseBtn);
    }

    SetWindowTextW(dlg, LoadStringResource(nullptr, IDS_CAPTION_CONNECTIONS).c_str());
    if (init->listTitle)
    {
        SetWindowTextW(init->listTitle, LoadStringResource(nullptr, IDS_CAPTION_CONNECTIONS).c_str());
    }
    if (const HWND ok = GetDlgItem(dlg, IDOK))
    {
        SetWindowTextW(ok, LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_CONNECT).c_str());
    }
    if (const HWND closeBtn = GetDlgItem(dlg, IDC_CONNECTION_CLOSE))
    {
        SetWindowTextW(closeBtn, LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_CLOSE).c_str());
    }
    if (const HWND cancel = GetDlgItem(dlg, IDCANCEL))
    {
        SetWindowTextW(cancel, LoadStringResource(nullptr, IDS_BTN_CANCEL).c_str());
    }
    if (const HWND btnNew = GetDlgItem(dlg, IDC_CONNECTION_NEW))
    {
        SetWindowTextW(btnNew, LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_NEW_ELLIPSIS).c_str());
    }
    if (const HWND btnRename = GetDlgItem(dlg, IDC_CONNECTION_RENAME))
    {
        SetWindowTextW(btnRename, LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_RENAME_ELLIPSIS).c_str());
    }
    if (const HWND btnRemove = GetDlgItem(dlg, IDC_CONNECTION_REMOVE))
    {
        SetWindowTextW(btnRemove, LoadStringResource(nullptr, IDS_CONNECTIONS_BTN_REMOVE).c_str());
    }

    ApplyTitleBarTheme(dlg, init->theme, GetActiveWindow() == dlg);

    init->backgroundBrush.reset(CreateSolidBrush(init->theme.windowBackground));

    init->cardBackgroundColor = GetControlSurfaceColor(init->theme);

    init->inputBackgroundColor         = BlendColor(init->cardBackgroundColor, init->theme.windowBackground, init->theme.dark ? 50 : 30, 255);
    init->inputFocusedBackgroundColor  = BlendColor(init->inputBackgroundColor, init->theme.menu.text, init->theme.dark ? 20 : 16, 255);
    init->inputDisabledBackgroundColor = BlendColor(init->theme.windowBackground, init->inputBackgroundColor, init->theme.dark ? 70 : 40, 255);

    init->cardBrush.reset();
    init->inputBrush.reset();
    init->inputFocusedBrush.reset();
    init->inputDisabledBrush.reset();
    if (! init->theme.highContrast)
    {
        init->cardBrush.reset(CreateSolidBrush(init->cardBackgroundColor));
        init->inputBrush.reset(CreateSolidBrush(init->inputBackgroundColor));
        init->inputFocusedBrush.reset(CreateSolidBrush(init->inputFocusedBackgroundColor));
        init->inputDisabledBrush.reset(CreateSolidBrush(init->inputDisabledBackgroundColor));
    }

    init->inputFrameStyle.theme                        = &init->theme;
    init->inputFrameStyle.backdropBrush                = init->cardBrush ? init->cardBrush.get() : init->backgroundBrush.get();
    init->inputFrameStyle.inputBackgroundColor         = init->inputBackgroundColor;
    init->inputFrameStyle.inputFocusedBackgroundColor  = init->inputFocusedBackgroundColor;
    init->inputFrameStyle.inputDisabledBackgroundColor = init->inputDisabledBackgroundColor;

    const HFONT dialogFont = GetDialogFont(dlg);
    EnsureFonts(*init, dialogFont);

    init->toggleOnLabel     = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    init->toggleOffLabel    = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
    init->quickConnectLabel = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
    if (init->quickConnectLabel.empty())
    {
        init->quickConnectLabel = L"<Quick Connect>";
    }

    const HWND settingsParent = init->settingsHost ? init->settingsHost : dlg;

    if (! init->awsRegionCombo)
    {
        init->awsRegionCombo = CreateWindowExW(0,
                                               L"ComboBox",
                                               L"",
                                               WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL,
                                               0,
                                               0,
                                               10,
                                               10,
                                               settingsParent,
                                               reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_CONNECTION_AWS_REGION_COMBO)),
                                               GetModuleHandleW(nullptr),
                                               nullptr);
        if (init->awsRegionCombo)
        {
            SetWindowPos(init->awsRegionCombo, init->hostEdit, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            PrepareFlatControl(init->awsRegionCombo);
            ApplyThemeToComboBox(init->awsRegionCombo, init->theme);
            PopulateAwsRegionCombo(init->awsRegionCombo);

            COMBOBOXINFO cbi{};
            cbi.cbSize = sizeof(cbi);
            if (GetComboBoxInfo(init->awsRegionCombo, &cbi) && cbi.hwndItem)
            {
                PrepareEditMargins(cbi.hwndItem);
            }

            ShowWindow(init->awsRegionCombo, SW_HIDE);
        }
    }

    const auto setupToggleStyle = [&](HWND toggle, int controlId) noexcept
    {
        if (! toggle)
        {
            return;
        }

        SetWindowTextW(toggle, L"");

        if (init->theme.highContrast)
        {
            LONG_PTR style = GetWindowLongPtrW(toggle, GWL_STYLE);
            style &= ~BS_TYPEMASK;
            style |= BS_AUTOCHECKBOX;
            SetWindowLongPtrW(toggle, GWL_STYLE, style);
            SetWindowPos(toggle, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
            return;
        }

        EnableOwnerDrawButton(settingsParent, controlId);
    };

    if (! init->theme.highContrast)
    {
        EnableOwnerDrawButton(dlg, IDOK);
        EnableOwnerDrawButton(dlg, IDC_CONNECTION_CLOSE);
        EnableOwnerDrawButton(dlg, IDCANCEL);
        EnableOwnerDrawButton(dlg, IDC_CONNECTION_NEW);
        EnableOwnerDrawButton(settingsParent, IDC_CONNECTION_SHOW_SECRET);
        EnableOwnerDrawButton(dlg, IDC_CONNECTION_RENAME);
        EnableOwnerDrawButton(dlg, IDC_CONNECTION_REMOVE);
        EnableOwnerDrawButton(settingsParent, IDC_CONNECTION_SSH_PRIVATEKEY_BROWSE);
        EnableOwnerDrawButton(settingsParent, IDC_CONNECTION_SSH_KNOWNHOSTS_BROWSE);
    }

    setupToggleStyle(init->anonymousToggle, IDC_CONNECTION_ANONYMOUS);
    setupToggleStyle(init->savePasswordToggle, IDC_CONNECTION_SAVE_PASSWORD);
    setupToggleStyle(init->requireHelloToggle, IDC_CONNECTION_REQUIRE_HELLO);
    setupToggleStyle(init->ignoreSslTrustToggle, IDC_CONNECTION_IGNORE_SSL_TRUST);
    setupToggleStyle(init->s3UseHttpsToggle, IDC_CONNECTION_S3_USE_HTTPS);
    setupToggleStyle(init->s3VerifyTlsToggle, IDC_CONNECTION_S3_VERIFY_TLS);
    setupToggleStyle(init->s3UseVirtualAddressingToggle, IDC_CONNECTION_S3_USE_VIRTUAL_ADDRESSING);

    if (! init->theme.highContrast && init->protocolCombo && ! ThemedControls::IsModernComboBox(init->protocolCombo))
    {
        RECT rc{};
        if (GetWindowRect(init->protocolCombo, &rc))
        {
            MapWindowPoints(nullptr, settingsParent, reinterpret_cast<POINT*>(&rc), 2);
            const int width  = std::max(0l, rc.right - rc.left);
            const int height = std::max(0l, rc.bottom - rc.top);

            const HWND oldCombo    = init->protocolCombo;
            const HWND modernCombo = ThemedControls::CreateModernComboBox(settingsParent, IDC_CONNECTION_PROTOCOL, &init->theme);
            if (modernCombo)
            {
                SetWindowPos(modernCombo, oldCombo, rc.left, rc.top, width, height, SWP_NOACTIVATE);
                DestroyWindow(oldCombo);
                init->protocolCombo = modernCombo;
            }
        }
    }

    if (init->protocolCombo)
    {
        PopulateProtocolCombo(init->protocolCombo);
        ApplyThemeToComboBox(init->protocolCombo, init->theme);
        PrepareFlatControl(init->protocolCombo);
    }

    if (init->list)
    {
        ListView_SetExtendedListViewStyle(init->list, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);
        LONG_PTR style = GetWindowLongPtrW(init->list, GWL_STYLE);
        style |= static_cast<LONG_PTR>(LVS_NOCOLUMNHEADER);
        SetWindowLongPtrW(init->list, GWL_STYLE, style);
        SetWindowPos(init->list, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);

        if (const HWND header = ListView_GetHeader(init->list))
        {
            ShowWindow(header, SW_HIDE);
        }

        SetupListViewColumns(init->list);
        ApplyThemeToListView(init->list, init->theme);
    }

    const std::array<HWND, 8> edits = {
        init->nameEdit,
        init->hostEdit,
        init->portEdit,
        init->initialPathEdit,
        init->userEdit,
        init->secretEdit,
        init->s3EndpointOverrideEdit,
        init->sshPrivateKeyEdit,
    };

    for (HWND edit : edits)
    {
        if (! edit)
        {
            continue;
        }

        PrepareFlatControl(edit);
        PrepareEditMargins(edit);
    }

    if (init->sshKnownHostsEdit)
    {
        PrepareFlatControl(init->sshKnownHostsEdit);
        PrepareEditMargins(init->sshKnownHostsEdit);
    }

    if (! init->theme.highContrast)
    {
        auto createFrame = [&](wil::unique_hwnd& frameOut, HWND input) noexcept
        {
            if (! input)
            {
                return;
            }

            const HWND parent = init->settingsHost ? init->settingsHost : dlg;

            frameOut.reset(
                CreateWindowExW(0, L"Static", L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr));
            if (! frameOut)
            {
                return;
            }

            SetWindowPos(frameOut.get(), input, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            ThemedInputFrames::InstallFrame(frameOut.get(), input, &init->inputFrameStyle);
        };

        createFrame(init->nameFrame, init->nameEdit);
        createFrame(init->protocolFrame, init->protocolCombo);
        createFrame(init->hostFrame, init->hostEdit);
        createFrame(init->awsRegionFrame, init->awsRegionCombo);
        createFrame(init->portFrame, init->portEdit);
        createFrame(init->initialPathFrame, init->initialPathEdit);
        createFrame(init->userFrame, init->userEdit);
        createFrame(init->secretFrame, init->secretEdit);
        createFrame(init->s3EndpointOverrideFrame, init->s3EndpointOverrideEdit);
        createFrame(init->sshPrivateKeyFrame, init->sshPrivateKeyEdit);
        createFrame(init->sshKnownHostsFrame, init->sshKnownHostsEdit);
    }

    int restoreShowCmd = SW_SHOWNORMAL;
    if (! init->modeless && init->baselineSettings)
    {
        restoreShowCmd = WindowPlacementPersistence::Restore(*init->baselineSettings, kConnectionManagerWindowId, dlg);
    }

    LayoutDialog(dlg, *init);

    RebuildList(dlg, *init);
    EnsureListSelection(*init);
    if (init->list && ListView_GetItemCount(init->list) == 0)
    {
        static_cast<void>(OnCommand(dlg, *init, IDC_CONNECTION_NEW));
        return FALSE;
    }
    UpdateControlEnabledState(*init);

    if (const auto model = GetSelectedModelIndex(*init); model.has_value())
    {
        LoadEditorFromProfile(*init, init->connections[model.value()]);
        UpdateControlEnabledState(*init);
    }

    if (! init->modeless && restoreShowCmd == SW_MAXIMIZE)
    {
        ShowWindow(dlg, SW_MAXIMIZE);
    }

    return TRUE;
}

INT_PTR OnCommand(HWND dlg, DialogState& state, int controlId) noexcept
{
    if (controlId == IDC_CONNECTION_SHOW_SECRET)
    {
        const bool show = ! state.secretVisible;

        if (show)
        {
            const auto model = GetSelectedModelIndex(state);
            if (model.has_value())
            {
                const auto& profile = state.connections[model.value()];
                if (! profile.id.empty() && profile.savePassword && profile.authMode != Common::Settings::ConnectionAuthMode::Anonymous)
                {
                    const std::wstring current = GetWindowTextString(state.secretEdit);
                    const auto placeholderIt   = state.secretPlaceholderById.find(profile.id);
                    if (placeholderIt != state.secretPlaceholderById.end() && ! state.secretDirtyIds.contains(profile.id) && current == placeholderIt->second)
                    {
                        std::wstring loaded;
                        const HRESULT loadHr = LoadStoredSecretForProfile(dlg, state, profile, loaded);
                        if (FAILED(loadHr))
                        {
                            if (loadHr != HRESULT_FROM_WIN32(ERROR_CANCELLED))
                            {
                                const std::wstring title = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
                                ShowDialogAlert(dlg, HOST_ALERT_ERROR, title, FormatHresultForUi(loadHr));
                            }
                            return TRUE;
                        }

                        state.loadingControls = true;
                        SetWindowTextW(state.secretEdit, loaded.c_str());
                        state.loadingControls = false;
                    }
                }
            }
        }

        state.secretVisible = show;
        UpdateSecretVisibility(state);
        if (state.secretEdit)
        {
            SetFocus(state.secretEdit);
        }
        return TRUE;
    }

    if (controlId == IDC_CONNECTION_NEW)
    {
        if (const auto current = GetSelectedModelIndex(state); current.has_value())
        {
            CommitEditorToProfile(state, state.connections[current.value()]);
        }

        Common::Settings::ConnectionProfile profile;
        profile.id = NewGuidString();
        if (profile.id.empty())
        {
            ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_CREATE_ID));
            return TRUE;
        }

        profile.pluginId = state.filterPluginId.empty() ? std::wstring(kProtocols[0].pluginId) : state.filterPluginId;

        profile.name                = MakeUniqueConnectionName(state.connections, LoadStringResource(nullptr, IDS_CONNECTIONS_DEFAULT_NEW_NAME), {});
        profile.host                = L"";
        profile.initialPath         = L"/";
        profile.port                = 0;
        profile.userName            = L"";
        profile.authMode            = Common::Settings::ConnectionAuthMode::Password;
        profile.savePassword        = false;
        profile.requireWindowsHello = true;

        ApplyPluginDefaultsToNewProfile(state, profile);

        state.connections.push_back(std::move(profile));
        RebuildList(dlg, state);

        const int count = ListView_GetItemCount(state.list);
        if (count > 0)
        {
            ListView_SetItemState(state.list, count - 1, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
            SetFocus(state.nameEdit);
            SendMessageW(state.nameEdit, EM_SETSEL, 0, -1);
        }

        return TRUE;
    }

    if (controlId == IDC_CONNECTION_RENAME)
    {
        if (const auto model = GetSelectedModelIndex(state); model.has_value() && IsQuickConnectProfile(state.connections[model.value()]))
        {
            return TRUE;
        }

        if (state.nameEdit)
        {
            SetFocus(state.nameEdit);
            SendMessageW(state.nameEdit, EM_SETSEL, 0, -1);
        }
        return TRUE;
    }

    if (controlId == IDC_CONNECTION_REMOVE)
    {
        const auto model = GetSelectedModelIndex(state);
        if (! model.has_value())
        {
            return TRUE;
        }

        if (IsQuickConnectProfile(state.connections[model.value()]))
        {
            return TRUE;
        }

        state.connections.erase(state.connections.begin() + static_cast<ptrdiff_t>(model.value()));
        RebuildList(dlg, state);
        if (const auto newModel = GetSelectedModelIndex(state); newModel.has_value())
        {
            LoadEditorFromProfile(state, state.connections[newModel.value()]);
        }
        else
        {
            state.loadingControls = true;
            SetWindowTextW(state.nameEdit, L"");
            SetWindowTextW(state.hostEdit, L"");
            if (state.awsRegionCombo)
            {
                SetWindowTextW(state.awsRegionCombo, L"");
            }
            SetWindowTextW(state.portEdit, L"");
            SetWindowTextW(state.initialPathEdit, L"");
            SetWindowTextW(state.userEdit, L"");
            SetWindowTextW(state.secretEdit, L"");
            SetWindowTextW(state.sshPrivateKeyEdit, L"");
            SetWindowTextW(state.sshKnownHostsEdit, L"");
            state.loadingControls = false;
        }
        UpdateControlEnabledState(state);
        return TRUE;
    }

    if (controlId == IDC_CONNECTION_SSH_PRIVATEKEY_BROWSE)
    {
        const auto selected = BrowseForFile(dlg, LoadStringResource(nullptr, IDS_CONNECTIONS_BROWSE_PRIVATE_KEY).c_str());
        if (selected.has_value())
        {
            SetWindowTextW(state.sshPrivateKeyEdit, selected->wstring().c_str());
        }
        return TRUE;
    }

    if (controlId == IDC_CONNECTION_SSH_KNOWNHOSTS_BROWSE)
    {
        const auto selected = BrowseForFile(dlg, LoadStringResource(nullptr, IDS_CONNECTIONS_BROWSE_KNOWN_HOSTS).c_str());
        if (selected.has_value())
        {
            SetWindowTextW(state.sshKnownHostsEdit, selected->wstring().c_str());
        }
        return TRUE;
    }

    if (controlId == IDOK)
    {
        const auto model = GetSelectedModelIndex(state);
        if (! model.has_value())
        {
            ShowDialogAlert(
                dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), LoadStringResource(nullptr, IDS_CONNECTIONS_ERR_SELECT_CONNECTION));
            return TRUE;
        }

        CommitEditorToProfile(state, state.connections[model.value()]);

        const HRESULT promptHr = PromptAndStageMissingPasswordForConnect(dlg, state, state.connections[model.value()]);
        if (promptHr == S_FALSE)
        {
            return TRUE; // user cancelled
        }
        if (FAILED(promptHr))
        {
            ShowDialogAlert(dlg, HOST_ALERT_ERROR, LoadStringResource(nullptr, IDS_CAPTION_ERROR), FormatHresultForUi(promptHr));
            return TRUE;
        }

        const HRESULT validateHr = ValidateProfileForConnect(dlg, state, state.connections[model.value()]);
        if (FAILED(validateHr))
        {
            return TRUE;
        }

        if (! SaveConnectionsSettings(dlg, state))
        {
            return TRUE;
        }

        const std::wstring_view selectedName = state.connections[model.value()].name;
        if (state.modeless)
        {
            NotifyConnectSelection(state, selectedName);
        }
        else
        {
            state.selectedConnectionName.assign(selectedName);
        }

        CloseConnectionManagerWindow(dlg, state, IDOK);
        return TRUE;
    }

    if (controlId == IDC_CONNECTION_CLOSE)
    {
        if (const auto model = GetSelectedModelIndex(state); model.has_value())
        {
            CommitEditorToProfile(state, state.connections[model.value()]);
        }

        if (! SaveConnectionsSettings(dlg, state))
        {
            return TRUE;
        }

        CloseConnectionManagerWindow(dlg, state, IDC_CONNECTION_CLOSE);
        return TRUE;
    }

    if (controlId == IDCANCEL)
    {
        CloseConnectionManagerWindow(dlg, state, IDCANCEL);
        return TRUE;
    }

    return FALSE;
}

INT_PTR CALLBACK ConnectionManagerDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<DialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

    switch (msg)
    {
        case WM_INITDIALOG: return OnInitDialog(dlg, reinterpret_cast<DialogState*>(lp));
        case WM_CLOSE:
            if (state)
            {
                return OnCommand(dlg, *state, IDC_CONNECTION_CLOSE);
            }
            EndDialog(dlg, IDC_CONNECTION_CLOSE);
            return TRUE;
        case WM_NCDESTROY:
        {
            std::unique_ptr<DialogState> stateOwner;

            if (state && state->modeless)
            {
                stateOwner.reset(state);
            }

            if (state && state->baselineSettings)
            {
                WindowPlacementPersistence::Save(*state->baselineSettings, kConnectionManagerWindowId, dlg);

                const Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(*state->baselineSettings);
                const HRESULT saveHr                            = Common::Settings::SaveSettings(state->appId, settingsToSave);
                if (FAILED(saveHr))
                {
                    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(state->appId);
                    Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
                }
            }

            if (stateOwner)
            {
                SetWindowLongPtrW(dlg, DWLP_USER, 0);
                if (g_connectionManagerDialog.get() == dlg)
                {
                    g_connectionManagerDialog.release();
                }
            }
        }
        break;
        case WM_ERASEBKGND:
            if (state && wp)
            {
                PaintDialogBackgroundAndCards(reinterpret_cast<HDC>(wp), dlg, *state);
                return TRUE;
            }
            break;
        case WM_CTLCOLORDLG: return OnCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnCtlColorStatic(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLOREDIT: return OnCtlColorEdit(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLORBTN: return OnCtlColorButton(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
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

            if (WindowMaximizeBehavior::ApplyVerticalMaximize(dlg, *info))
            {
                return TRUE;
            }
            break;
        }
        case WM_DRAWITEM:
        {
            if (! state || state->theme.highContrast)
            {
                break;
            }

            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lp);
            if (! dis || dis->CtlType != ODT_BUTTON)
            {
                break;
            }

            if (dis->hwndItem == state->anonymousToggle || dis->hwndItem == state->savePasswordToggle || dis->hwndItem == state->requireHelloToggle ||
                dis->hwndItem == state->ignoreSslTrustToggle || dis->hwndItem == state->s3UseHttpsToggle || dis->hwndItem == state->s3VerifyTlsToggle ||
                dis->hwndItem == state->s3UseVirtualAddressingToggle)
            {
                const bool toggledOn   = GetTwoStateToggleState(dis->hwndItem, state->theme);
                const COLORREF surface = state->cardBrush ? state->cardBackgroundColor : GetControlSurfaceColor(state->theme);
                const HFONT boldFont   = state->boldFont ? state->boldFont.get() : nullptr;
                DrawThemedSwitchToggle(*dis, state->theme, surface, boldFont, state->toggleOnLabel, state->toggleOffLabel, toggledOn);
                return TRUE;
            }

            DrawThemedPushButton(*dis, state->theme);
            return TRUE;
        }
        case WM_NOTIFY:
        {
            if (! state)
            {
                break;
            }

            const auto* hdr = reinterpret_cast<const NMHDR*>(lp);
            if (! hdr)
            {
                break;
            }

            if (hdr->idFrom == IDC_CONNECTION_LIST && hdr->code == NM_CUSTOMDRAW)
            {
                return OnListCustomDraw(state, reinterpret_cast<NMLVCUSTOMDRAW*>(lp));
            }

            if (hdr->idFrom == IDC_CONNECTION_LIST && hdr->code == LVN_ITEMCHANGED)
            {
                const auto* change = reinterpret_cast<const NMLISTVIEW*>(lp);
                if (! change)
                {
                    break;
                }

                const bool nowSelected = (change->uNewState & LVIS_SELECTED) != 0;
                if (! nowSelected)
                {
                    break;
                }

                const int newSel = change->iItem;
                if (state->selectedListIndex == newSel)
                {
                    break;
                }

                const int oldSel = state->selectedListIndex;
                if (oldSel >= 0 && oldSel < static_cast<int>(state->viewToModel.size()) && ! state->loadingControls)
                {
                    const size_t oldModel = state->viewToModel[static_cast<size_t>(oldSel)];
                    if (oldModel < state->connections.size())
                    {
                        CommitEditorToProfile(*state, state->connections[oldModel]);
                    }
                }

                state->selectedListIndex = newSel;

                if (const auto model = GetSelectedModelIndex(*state); model.has_value())
                {
                    LoadEditorFromProfile(*state, state->connections[model.value()]);
                }

                UpdateControlEnabledState(*state);
                state->settingsScrollOffset = 0;
                LayoutDialog(dlg, *state);
                return TRUE;
            }
            break;
        }
        case WM_SIZE:
        {
            if (state)
            {
                LayoutDialog(dlg, *state);
            }
            break;
        }
        case WM_COMMAND:
        {
            if (! state)
            {
                break;
            }

            const int id = LOWORD(wp);

            if (id == IDC_CONNECTION_PROTOCOL && HIWORD(wp) == CBN_SELCHANGE)
            {
                if (const auto model = GetSelectedModelIndex(*state); model.has_value())
                {
                    CommitEditorToProfile(*state, state->connections[model.value()]);
                    UpdateControlEnabledState(*state);
                    LayoutDialog(dlg, *state);
                }
                return TRUE;
            }

            if (id == IDC_CONNECTION_AWS_REGION_COMBO && HIWORD(wp) == CBN_SELCHANGE)
            {
                if (! state->awsRegionCombo)
                {
                    return TRUE;
                }

                const int sel = static_cast<int>(SendMessageW(state->awsRegionCombo, CB_GETCURSEL, 0, 0));
                if (sel >= 0)
                {
                    const LRESULT data = SendMessageW(state->awsRegionCombo, CB_GETITEMDATA, static_cast<WPARAM>(sel), 0);
                    if (data != CB_ERR)
                    {
                        const auto* regionCode = reinterpret_cast<const wchar_t*>(data);
                        state->loadingControls = true;
                        SetWindowTextW(state->awsRegionCombo, regionCode ? regionCode : L"");
                        state->loadingControls = false;
                        SendMessageW(state->awsRegionCombo, CB_SETEDITSEL, 0, MAKELPARAM(0, -1));
                    }
                }
                return TRUE;
            }

            if (id == IDC_CONNECTION_PASSWORD && (HIWORD(wp) == EN_SETFOCUS || HIWORD(wp) == EN_CHANGE))
            {
                if (state->loadingControls)
                {
                    return TRUE;
                }

                const auto model = GetSelectedModelIndex(*state);
                if (! model.has_value())
                {
                    return TRUE;
                }

                const auto& profile = state->connections[model.value()];
                if (profile.id.empty())
                {
                    return TRUE;
                }

                if (HIWORD(wp) == EN_SETFOCUS)
                {
                    if (state->secretEdit && state->secretPlaceholderById.contains(profile.id) && ! state->secretDirtyIds.contains(profile.id))
                    {
                        SendMessageW(state->secretEdit, EM_SETSEL, 0, -1);
                    }
                    return TRUE;
                }

                state->secretDirtyIds.insert(profile.id);
                return TRUE;
            }

            if (id == IDC_CONNECTION_SSH_PRIVATEKEY && HIWORD(wp) == EN_CHANGE)
            {
                if (! state->loadingControls)
                {
                    if (const auto model = GetSelectedModelIndex(*state); model.has_value())
                    {
                        CommitEditorToProfile(*state, state->connections[model.value()]);
                        UpdateControlEnabledState(*state);
                    }
                }
                return TRUE;
            }

            if (id == IDC_CONNECTION_ANONYMOUS || id == IDC_CONNECTION_SAVE_PASSWORD || id == IDC_CONNECTION_IGNORE_SSL_TRUST ||
                id == IDC_CONNECTION_S3_USE_HTTPS || id == IDC_CONNECTION_S3_VERIFY_TLS || id == IDC_CONNECTION_S3_USE_VIRTUAL_ADDRESSING)
            {
                if (HIWORD(wp) == BN_CLICKED && ! state->theme.highContrast)
                {
                    const HWND button = reinterpret_cast<HWND>(lp);
                    if (button)
                    {
                        const bool toggledOn = GetTwoStateToggleState(button, state->theme);
                        SetTwoStateToggleState(button, state->theme, ! toggledOn);
                    }
                }

                if (const auto model = GetSelectedModelIndex(*state); model.has_value())
                {
                    CommitEditorToProfile(*state, state->connections[model.value()]);
                    SetTwoStateToggleState(state->requireHelloToggle, state->theme, state->connections[model.value()].requireWindowsHello);
                    UpdateControlEnabledState(*state);
                }
                return TRUE;
            }

            return OnCommand(dlg, *state, id);
        }
    }

    return FALSE;
}

} // namespace

HRESULT ShowConnectionManagerDialog(HWND owner,
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

    owner = NormalizeOwnerWindow(owner);

    DialogState state;
    state.baselineSettings = &settings;
    state.appId            = std::wstring(appId);
    state.theme            = theme;
    state.filterPluginId   = std::wstring(filterPluginId);
    PopulateStateFromSettings(state, settings, filterPluginId);

#pragma warning(push)
#pragma warning(disable : 5039) // pointer/reference to potentially throwing function passed to extern "C"
    const INT_PTR dlgResult = DialogBoxParamW(
        GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_CONNECTION_MANAGER), owner, ConnectionManagerDialogProc, reinterpret_cast<LPARAM>(&state));
#pragma warning(pop)

    if (dlgResult == IDOK)
    {
        selectedConnectionNameOut = state.selectedConnectionName;
        return selectedConnectionNameOut.empty() ? E_FAIL : S_OK;
    }

    return S_FALSE;
}

bool ShowConnectionManagerWindow(HWND owner,
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

    if (const HWND existing = g_connectionManagerDialog.get())
    {
        if (! IsWindow(existing))
        {
            g_connectionManagerDialog.release();
        }
        else
        {
            auto* state = reinterpret_cast<DialogState*>(GetWindowLongPtrW(existing, DWLP_USER));
            if (state)
            {
                state->baselineSettings    = &settings;
                state->theme               = theme;
                state->connectNotifyWindow = effectiveOwner;
                state->connectTargetPane   = targetPane;

                const std::wstring newFilter(filterPluginId);
                if (state->filterPluginId != newFilter)
                {
                    state->filterPluginId = newFilter;
                    RebuildList(existing, *state);
                    if (const auto model = GetSelectedModelIndex(*state); model.has_value())
                    {
                        LoadEditorFromProfile(*state, state->connections[model.value()]);
                    }
                    UpdateControlEnabledState(*state);
                    LayoutDialog(existing, *state);
                    RedrawWindow(existing, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
                }
            }

            if (IsIconic(existing))
            {
                ShowWindow(existing, SW_RESTORE);
            }
            else
            {
                ShowWindow(existing, SW_SHOW);
            }
            SetForegroundWindow(existing);
            return true;
        }
    }

    auto statePtr              = std::make_unique<DialogState>();
    auto* state                = statePtr.get();
    state->modeless            = true;
    state->connectNotifyWindow = effectiveOwner;
    state->connectTargetPane   = targetPane;
    state->baselineSettings    = &settings;
    state->appId               = std::wstring(appId);
    state->theme               = theme;
    state->filterPluginId      = std::wstring(filterPluginId);

    PopulateStateFromSettings(*state, settings, filterPluginId);

#pragma warning(push)
#pragma warning(disable : 5039) // pointer/reference to potentially throwing function passed to extern "C"
    const HWND dlg = CreateDialogParamW(
        GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_CONNECTION_MANAGER), nullptr, ConnectionManagerDialogProc, reinterpret_cast<LPARAM>(state));
#pragma warning(pop)
    if (! dlg)
    {
        return false;
    }

    g_connectionManagerDialog.reset(dlg);
    static_cast<void>(statePtr.release());
    const int showCmd = WindowPlacementPersistence::Restore(settings, kConnectionManagerWindowId, dlg);
    ShowWindow(dlg, showCmd);
    SetForegroundWindow(dlg);
    return true;
}

HWND GetConnectionManagerDialogHandle() noexcept
{
    if (const HWND dlg = g_connectionManagerDialog.get(); dlg && IsWindow(dlg))
    {
        return dlg;
    }
    return nullptr;
}
