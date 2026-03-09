#include "ConnectionProfileUtils.h"

#include "Framework.h"

#include <format>
#include <limits>

#include "Helpers.h"
#include "SettingsStore.h"

namespace ConnectionProfileUtils
{
namespace
{
[[nodiscard]] const wchar_t* PluginIdToScheme(std::wstring_view pluginId) noexcept
{
    if (OrdinalString::EqualsNoCase(pluginId, L"builtin/file-system-ftp"))
    {
        return L"ftp";
    }
    if (OrdinalString::EqualsNoCase(pluginId, L"builtin/file-system-sftp"))
    {
        return L"sftp";
    }
    if (OrdinalString::EqualsNoCase(pluginId, L"builtin/file-system-scp"))
    {
        return L"scp";
    }
    if (OrdinalString::EqualsNoCase(pluginId, L"builtin/file-system-imap"))
    {
        return L"imap";
    }
    if (OrdinalString::EqualsNoCase(pluginId, L"builtin/file-system-gdrive"))
    {
        return L"gdrive";
    }
    if (OrdinalString::EqualsNoCase(pluginId, L"builtin/file-system-onedrive-personal"))
    {
        return L"onedrive";
    }
    if (OrdinalString::EqualsNoCase(pluginId, L"builtin/file-system-onedrive-business"))
    {
        return L"onedrive-pro";
    }
    if (OrdinalString::EqualsNoCase(pluginId, L"builtin/file-system-sharepoint"))
    {
        return L"sharepoint";
    }
    if (OrdinalString::EqualsNoCase(pluginId, L"builtin/file-system-s3"))
    {
        return L"s3";
    }
    if (OrdinalString::EqualsNoCase(pluginId, L"builtin/file-system-s3table"))
    {
        return L"s3table";
    }

    return nullptr;
}
} // namespace

std::optional<bool> ExtraGetBool(const Common::Settings::JsonValue& extra, std::string_view key) noexcept
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

std::optional<uint32_t> ExtraGetUInt32(const Common::Settings::JsonValue& extra, std::string_view key) noexcept
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

const Common::Settings::ConnectionProfile* FindConnectionProfileByName(const Common::Settings::Settings* settings, std::wstring_view connectionName) noexcept
{
    if (! settings || ! settings->connections)
    {
        return nullptr;
    }

    for (const Common::Settings::ConnectionProfile& profile : settings->connections->items)
    {
        if (OrdinalString::EqualsNoCase(profile.name, connectionName))
        {
            return &profile;
        }
    }

    return nullptr;
}

std::optional<std::wstring> TryParseConnNameFromPluginPath(std::wstring_view pluginPath) noexcept
{
    if (pluginPath.empty())
    {
        return std::nullopt;
    }

    // Accept both raw plugin paths ("/@conn:<name>/...") and prefixed variants ("ftp:/@conn:<name>/...").
    std::wstring_view view(pluginPath);
    const size_t colon = view.find(L':');
    if (colon != std::wstring_view::npos && colon > 1u && (colon + 1u) < view.size())
    {
        const wchar_t after = view[colon + 1u];
        if (after == L'/' || after == L'\\')
        {
            view.remove_prefix(colon + 1u);
        }
    }

    constexpr std::wstring_view kConnPrefixForward = L"/@conn:";
    constexpr std::wstring_view kConnPrefixBack    = L"\\@conn:";

    std::wstring_view rest;
    if (OrdinalString::StartsWithNoCase(view, kConnPrefixForward))
    {
        rest = view.substr(kConnPrefixForward.size());
    }
    else if (OrdinalString::StartsWithNoCase(view, kConnPrefixBack))
    {
        rest = view.substr(kConnPrefixBack.size());
    }
    else
    {
        return std::nullopt;
    }

    const size_t slash           = rest.find_first_of(L"/\\");
    const std::wstring_view name = slash == std::wstring_view::npos ? rest : rest.substr(0, slash);
    if (name.empty())
    {
        return std::nullopt;
    }

    return std::wstring(name);
}

std::optional<std::wstring> TryParseConnNameFromPluginPath(const std::optional<std::filesystem::path>& pluginPath) noexcept
{
    if (! pluginPath.has_value())
    {
        return std::nullopt;
    }

    return TryParseConnNameFromPluginPath(std::wstring_view(pluginPath->native()));
}

bool ConnectionProfileUsesInsecureTls(const Common::Settings::ConnectionProfile& profile) noexcept
{
    // IMAP: allow skipping certificate validation.
    if (OrdinalString::EqualsNoCase(profile.pluginId, L"builtin/file-system-imap"))
    {
        if (ExtraGetBool(profile.extra, "ignoreSslTrust").value_or(false))
        {
            return true;
        }
    }

    // S3/S3 Tables: allow disabling HTTPS or TLS verification.
    if (OrdinalString::EqualsNoCase(profile.pluginId, L"builtin/file-system-s3") ||
        OrdinalString::EqualsNoCase(profile.pluginId, L"builtin/file-system-s3table"))
    {
        const bool useHttps  = ExtraGetBool(profile.extra, "useHttps").value_or(true);
        const bool verifyTls = ExtraGetBool(profile.extra, "verifyTls").value_or(true);
        if (! useHttps || ! verifyTls)
        {
            return true;
        }
    }

    return false;
}

std::wstring BuildConnectionDisplayUrl(const Common::Settings::ConnectionProfile& profile) noexcept
{
    const wchar_t* scheme = PluginIdToScheme(profile.pluginId);
    if (! scheme)
    {
        return {};
    }

    std::wstring authority = profile.host;
    if (! authority.empty() && profile.port != 0)
    {
        authority.push_back(L':');
        authority.append(std::to_wstring(profile.port));
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

    const bool hideAnonymous = OrdinalString::EqualsNoCase(profile.pluginId, L"builtin/file-system-ftp") && (user == L"anonymous");
    const bool showUser      = ! user.empty() && ! hideAnonymous;
    std::wstring result(scheme);
    result.append(L"://");
    if (authority.empty())
    {
        return result;
    }
    if (showUser)
    {
        result.append(user);
        result.push_back(L'@');
    }
    result.append(authority);
    return result;
}
} // namespace ConnectionProfileUtils
