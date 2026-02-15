#include "ConnectionSecrets.h"

#include <cwctype>
#include <format>
#include <limits>
#include <mutex>
#include <optional>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <wincred.h>
#include <windows.h>

#pragma comment(lib, "Advapi32.lib")

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/resource.h>
#pragma warning(pop)

#include "SettingsStore.h"

namespace RedSalamander::Connections
{
namespace
{
constexpr std::wstring_view kTargetPrefix = L"RedSalamander/Connections/";

[[nodiscard]] std::wstring_view SecretKindSuffix(SecretKind kind) noexcept
{
    switch (kind)
    {
        case SecretKind::Password: return L"password";
        case SecretKind::SshKeyPassphrase: return L"sshKeyPassphrase";
    }
    return L"password";
}
} // namespace

std::wstring BuildCredentialTargetName(std::wstring_view connectionId, SecretKind kind)
{
    if (connectionId.empty())
    {
        return {};
    }

    return std::format(L"{}{}/{}", kTargetPrefix, connectionId, SecretKindSuffix(kind));
}

HRESULT SaveGenericCredential(std::wstring_view targetName, std::wstring_view userName, std::wstring_view secret) noexcept
{
    if (targetName.empty())
    {
        return E_INVALIDARG;
    }

    if (secret.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    std::wstring targetNameCopy;
    std::wstring userNameCopy;
    std::wstring secretCopy;
    targetNameCopy.assign(targetName);
    if (! userName.empty())
    {
        userNameCopy.assign(userName);
    }
    secretCopy.assign(secret);

    CREDENTIALW cred{};
    cred.Type       = CRED_TYPE_GENERIC;
    cred.TargetName = targetNameCopy.data();
    cred.UserName   = userNameCopy.empty() ? nullptr : userNameCopy.data();

    cred.Persist = CRED_PERSIST_LOCAL_MACHINE;

    const size_t bytesToStore = (secretCopy.size() + 1u) * sizeof(wchar_t);
    if (bytesToStore > std::numeric_limits<DWORD>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
    }

    cred.CredentialBlobSize = static_cast<DWORD>(bytesToStore);
    cred.CredentialBlob     = reinterpret_cast<BYTE*>(const_cast<wchar_t*>(secretCopy.c_str()));

    if (! CredWriteW(&cred, 0))
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    return S_OK;
}

HRESULT LoadGenericCredential(std::wstring_view targetName, std::wstring& userNameOut, std::wstring& secretOut) noexcept
{
    userNameOut.clear();
    secretOut.clear();

    if (targetName.empty())
    {
        return E_INVALIDARG;
    }

    std::wstring targetNameCopy;
    targetNameCopy.assign(targetName);

    PCREDENTIALW raw = nullptr;
    if (! CredReadW(targetNameCopy.c_str(), CRED_TYPE_GENERIC, 0, &raw))
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    wil::unique_any<PCREDENTIALW, decltype(&::CredFree), ::CredFree> cred(raw);

    if (cred.get()->UserName)
    {
        userNameOut.assign(cred.get()->UserName);
    }

    const BYTE* blobBytes = cred.get()->CredentialBlob;
    const DWORD byteCount = cred.get()->CredentialBlobSize;
    if (! blobBytes || byteCount < sizeof(wchar_t) || (byteCount % sizeof(wchar_t)) != 0)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
    }

    const size_t charCount = static_cast<size_t>(byteCount / sizeof(wchar_t));
    const wchar_t* blob    = reinterpret_cast<const wchar_t*>(blobBytes);
    if (blob[charCount - 1u] != L'\0')
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
    }

    size_t len = 0;
    while (len < charCount && blob[len] != L'\0')
    {
        ++len;
    }

    secretOut.assign(blob, len);

    if (secretOut.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
    }

    return S_OK;
}

HRESULT DeleteGenericCredential(std::wstring_view targetName) noexcept
{
    if (targetName.empty())
    {
        return E_INVALIDARG;
    }

    std::wstring targetNameCopy;
    targetNameCopy.assign(targetName);

    if (! CredDeleteW(targetNameCopy.c_str(), CRED_TYPE_GENERIC, 0))
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    return S_OK;
}

namespace
{
std::mutex g_quickConnectMutex;
std::optional<Common::Settings::ConnectionProfile> g_quickConnectProfile;
std::wstring g_quickConnectPassword;
std::wstring g_quickConnectPassphrase;
bool g_quickConnectHasPassword   = false;
bool g_quickConnectHasPassphrase = false;

void TryAssign(std::wstring& target, std::wstring_view value) noexcept
{
    target.assign(value);
}

void ClearProfile(Common::Settings::ConnectionProfile& profile) noexcept
{
    profile.id.clear();
    profile.name.clear();
    profile.pluginId.clear();
    profile.host.clear();
    profile.port = 0;
    profile.initialPath.clear();
    profile.userName.clear();
    profile.authMode            = Common::Settings::ConnectionAuthMode::Password;
    profile.savePassword        = false;
    profile.requireWindowsHello = true;
    profile.extra.value         = std::monostate{};
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
} // namespace

bool IsQuickConnectConnectionId(std::wstring_view connectionId) noexcept
{
    return connectionId == kQuickConnectConnectionId;
}

bool IsQuickConnectConnectionName(std::wstring_view connectionName) noexcept
{
    return EqualsIgnoreCase(connectionName, kQuickConnectConnectionName);
}

void EnsureQuickConnectProfile(std::wstring_view preferredPluginId) noexcept
{
    std::scoped_lock lock(g_quickConnectMutex);
    if (g_quickConnectProfile.has_value())
    {
        return;
    }

    Common::Settings::ConnectionProfile profile;
    TryAssign(profile.id, kQuickConnectConnectionId);
    TryAssign(profile.name, kQuickConnectConnectionName);
    TryAssign(profile.pluginId, preferredPluginId.empty() ? L"builtin/file-system-ftp" : preferredPluginId);
    profile.host                = L"";
    profile.port                = 0;
    profile.userName            = L"";
    profile.authMode            = Common::Settings::ConnectionAuthMode::Password;
    profile.savePassword        = true;
    profile.requireWindowsHello = true;

    g_quickConnectProfile.emplace(std::move(profile));
}

void GetQuickConnectProfile(Common::Settings::ConnectionProfile& out) noexcept
{
    EnsureQuickConnectProfile({});
    std::scoped_lock lock(g_quickConnectMutex);
    if (! g_quickConnectProfile.has_value())
    {
        ClearProfile(out);
        return;
    }

    out = g_quickConnectProfile.value();
}

void SetQuickConnectProfile(const Common::Settings::ConnectionProfile& profile) noexcept
{
    EnsureQuickConnectProfile({});

    std::scoped_lock lock(g_quickConnectMutex);

    Common::Settings::ConnectionProfile copy = profile;
    TryAssign(copy.id, kQuickConnectConnectionId);
    TryAssign(copy.name, kQuickConnectConnectionName);
    g_quickConnectProfile.emplace(std::move(copy));
}

bool HasQuickConnectSecret(SecretKind kind) noexcept
{
    std::scoped_lock lock(g_quickConnectMutex);
    switch (kind)
    {
        case SecretKind::Password: return g_quickConnectHasPassword;
        case SecretKind::SshKeyPassphrase: return g_quickConnectHasPassphrase;
    }
    return false;
}

HRESULT LoadQuickConnectSecret(SecretKind kind, std::wstring& secretOut) noexcept
{
    secretOut.clear();
    std::scoped_lock lock(g_quickConnectMutex);

    switch (kind)
    {
        case SecretKind::Password:
            if (! g_quickConnectHasPassword || g_quickConnectPassword.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
            }

            secretOut = g_quickConnectPassword;
            return S_OK;
        case SecretKind::SshKeyPassphrase:
            if (! g_quickConnectHasPassphrase || g_quickConnectPassphrase.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
            }

            secretOut = g_quickConnectPassphrase;
            return S_OK;
    }

    return E_INVALIDARG;
}

void SetQuickConnectSecret(SecretKind kind, std::wstring_view secret) noexcept
{
    std::scoped_lock lock(g_quickConnectMutex);
    if (secret.empty())
    {
        switch (kind)
        {
            case SecretKind::Password:
                g_quickConnectHasPassword = false;
                g_quickConnectPassword.clear();
                return;
            case SecretKind::SshKeyPassphrase:
                g_quickConnectHasPassphrase = false;
                g_quickConnectPassphrase.clear();
                return;
        }
        return;
    }

    switch (kind)
    {
        case SecretKind::Password:
            g_quickConnectPassword.assign(secret);
            g_quickConnectHasPassword = true;
            return;
        case SecretKind::SshKeyPassphrase:
            g_quickConnectPassphrase.assign(secret);
            g_quickConnectHasPassphrase = true;
            return;
    }
}

void ClearQuickConnectSecret(SecretKind kind) noexcept
{
    SetQuickConnectSecret(kind, {});
}
} // namespace RedSalamander::Connections
