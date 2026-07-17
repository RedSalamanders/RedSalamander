#include "ConnectionSecrets.h"

#include <cwctype>
#include <atomic>
#include <format>
#include <limits>
#include <mutex>
#include <optional>
#include <unordered_map>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <wincred.h>
#include <windows.h>

#pragma comment(lib, "Advapi32.lib")

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"
#include "SettingsStore.h"

namespace RedSalamander::Connections
{
namespace
{
constexpr std::wstring_view kTargetPrefix = L"RedSalamander/Connections/";
#ifdef ENABLE_TESTS
std::atomic<CredentialPersistenceFault> g_credentialPersistenceFault{CredentialPersistenceFault::None};
#endif

[[nodiscard]] std::wstring_view SecretKindSuffix(SecretKind kind) noexcept
{
    switch (kind)
    {
        case SecretKind::Password: return L"password";
        case SecretKind::SshKeyPassphrase: return L"sshKeyPassphrase";
        case SecretKind::RefreshToken: return L"refreshToken";
    }
    return L"password";
}
} // namespace

std::wstring BuildCredentialTargetName(std::wstring_view connectionId, SecretKind kind)
{
    std::wstring canonicalId;
    if (FAILED(Common::Settings::NormalizeConnectionProfileId(connectionId, canonicalId)) || canonicalId != connectionId)
    {
        return {};
    }

    return std::format(L"{}{}/{}", kTargetPrefix, canonicalId, SecretKindSuffix(kind));
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
#ifdef ENABLE_TESTS
    if (g_credentialPersistenceFault.exchange(CredentialPersistenceFault::None, std::memory_order_acq_rel) == CredentialPersistenceFault::SaveOnce)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
#endif

    std::wstring targetNameCopy;
    std::wstring userNameCopy;
    std::wstring secretCopy;
    targetNameCopy.assign(targetName);
    if (! userName.empty())
    {
        userNameCopy.assign(userName);
    }
    secretCopy.assign(secret);
    auto clearSecretCopy = wil::scope_exit([&]() noexcept { SecureWipe::SecureClear(secretCopy); });

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
    SecureWipe::SecureClear(secretOut);

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

    // The credential blob holds the plaintext secret; scrub it before CredFree releases the buffer so the
    // cleartext does not linger in the credential-manager heap block after this function returns. Declared
    // after cred so it runs first on scope exit (wipe, then free).
    auto wipeBlob = wil::scope_exit([&cred]() noexcept
    {
        if (cred && cred.get()->CredentialBlob != nullptr && cred.get()->CredentialBlobSize != 0)
        {
            SecureZeroMemory(cred.get()->CredentialBlob, cred.get()->CredentialBlobSize);
        }
    });

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
#ifdef ENABLE_TESTS
    if (g_credentialPersistenceFault.exchange(CredentialPersistenceFault::None, std::memory_order_acq_rel) == CredentialPersistenceFault::DeleteOnce)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
#endif

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
std::wstring g_quickConnectRefreshToken;
bool g_quickConnectHasPassword     = false;
bool g_quickConnectHasPassphrase   = false;
bool g_quickConnectHasRefreshToken = false;

std::mutex g_secretAccessAuthorizationMutex;

struct SecretAccessAuthorizationKey
{
    std::wstring connectionId;
    SecretKind kind                 = SecretKind::Password;
    SecretAccessPurpose purpose     = SecretAccessPurpose::Interactive;

    bool operator==(const SecretAccessAuthorizationKey&) const noexcept = default;
};

struct SecretAccessAuthorizationKeyHash
{
    [[nodiscard]] size_t operator()(const SecretAccessAuthorizationKey& key) const noexcept
    {
        size_t hash = std::hash<std::wstring>{}(key.connectionId);
        hash        = hash * 131u + static_cast<size_t>(key.kind);
        return hash * 131u + static_cast<size_t>(key.purpose);
    }
};

std::unordered_map<SecretAccessAuthorizationKey, uint64_t, SecretAccessAuthorizationKeyHash> g_secretAccessAuthorizationTicks;

[[nodiscard]] bool TryGetCanonicalConnectionId(std::wstring_view connectionId, std::wstring& canonicalId) noexcept
{
    return SUCCEEDED(Common::Settings::NormalizeConnectionProfileId(connectionId, canonicalId)) && canonicalId == connectionId;
}

[[nodiscard]] bool IsAuthorizationFresh(uint64_t now, uint64_t authorizedAt, uint64_t timeoutMs) noexcept
{
    return timeoutMs != 0u && now - authorizedAt < timeoutMs;
}

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

#ifdef ENABLE_TESTS
void SetCredentialPersistenceFaultForTesting(CredentialPersistenceFault fault) noexcept
{
    g_credentialPersistenceFault.store(fault, std::memory_order_release);
}
#endif

bool IsQuickConnectConnectionId(std::wstring_view connectionId) noexcept
{
    return connectionId == Common::Settings::kQuickConnectConnectionId;
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
    TryAssign(profile.id, Common::Settings::kQuickConnectConnectionId);
    TryAssign(profile.name, kQuickConnectConnectionName);
    TryAssign(profile.pluginId, preferredPluginId.empty() ? L"builtin/file-system-ftp" : preferredPluginId);
    profile.host         = L"";
    profile.port         = 0;
    profile.userName     = L"";
    profile.authMode     = (profile.pluginId == L"builtin/file-system-onedrive-personal" || profile.pluginId == L"builtin/file-system-onedrive-business" ||
                            profile.pluginId == L"builtin/file-system-sharepoint" || profile.pluginId == L"builtin/file-system-gdrive")
                               ? Common::Settings::ConnectionAuthMode::OAuth2Pkce
                               : Common::Settings::ConnectionAuthMode::Password;
    profile.savePassword = true;
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
    TryAssign(copy.id, Common::Settings::kQuickConnectConnectionId);
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
        case SecretKind::RefreshToken: return g_quickConnectHasRefreshToken;
    }
    return false;
}

HRESULT LoadQuickConnectSecret(SecretKind kind, std::wstring& secretOut) noexcept
{
    SecureWipe::SecureClear(secretOut);
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
        case SecretKind::RefreshToken:
            if (! g_quickConnectHasRefreshToken || g_quickConnectRefreshToken.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
            }

            secretOut = g_quickConnectRefreshToken;
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
                SecureWipe::SecureClear(g_quickConnectPassword);
                return;
            case SecretKind::SshKeyPassphrase:
                g_quickConnectHasPassphrase = false;
                SecureWipe::SecureClear(g_quickConnectPassphrase);
                return;
            case SecretKind::RefreshToken:
                g_quickConnectHasRefreshToken = false;
                SecureWipe::SecureClear(g_quickConnectRefreshToken);
                return;
        }
        return;
    }

    switch (kind)
    {
        case SecretKind::Password:
            SecureWipe::SecureClear(g_quickConnectPassword);
            g_quickConnectPassword.assign(secret);
            g_quickConnectHasPassword = true;
            return;
        case SecretKind::SshKeyPassphrase:
            SecureWipe::SecureClear(g_quickConnectPassphrase);
            g_quickConnectPassphrase.assign(secret);
            g_quickConnectHasPassphrase = true;
            return;
        case SecretKind::RefreshToken:
            SecureWipe::SecureClear(g_quickConnectRefreshToken);
            g_quickConnectRefreshToken.assign(secret);
            g_quickConnectHasRefreshToken = true;
            return;
    }
}

void ClearQuickConnectSecret(SecretKind kind) noexcept
{
    SetQuickConnectSecret(kind, {});
}

void NoteSecretAccessAuthorized(std::wstring_view connectionId, SecretKind kind) noexcept
{
    std::wstring canonicalId;
    if (! TryGetCanonicalConnectionId(connectionId, canonicalId))
    {
        return;
    }

    const uint64_t now = GetTickCount64();
    std::scoped_lock lock(g_secretAccessAuthorizationMutex);
    g_secretAccessAuthorizationTicks.insert_or_assign(
        SecretAccessAuthorizationKey{canonicalId, kind, SecretAccessPurpose::Interactive}, now);
    g_secretAccessAuthorizationTicks.insert_or_assign(
        SecretAccessAuthorizationKey{std::move(canonicalId), kind, SecretAccessPurpose::Background}, now);
}

bool IsSecretAccessAuthorized(std::wstring_view connectionId,
                              SecretKind kind,
                              SecretAccessPurpose purpose,
                              uint64_t reauthTimeoutMs) noexcept
{
    std::wstring canonicalId;
    if (! TryGetCanonicalConnectionId(connectionId, canonicalId))
    {
        return false;
    }

    std::scoped_lock lock(g_secretAccessAuthorizationMutex);
    const auto it = g_secretAccessAuthorizationTicks.find(SecretAccessAuthorizationKey{std::move(canonicalId), kind, purpose});
    if (it == g_secretAccessAuthorizationTicks.end())
    {
        return false;
    }

    return purpose == SecretAccessPurpose::Background || IsAuthorizationFresh(GetTickCount64(), it->second, reauthTimeoutMs);
}

void ClearSecretAccessAuthorization(std::wstring_view connectionId, SecretKind kind) noexcept
{
    std::wstring canonicalId;
    if (! TryGetCanonicalConnectionId(connectionId, canonicalId))
    {
        return;
    }

    std::scoped_lock lock(g_secretAccessAuthorizationMutex);
    g_secretAccessAuthorizationTicks.erase(SecretAccessAuthorizationKey{canonicalId, kind, SecretAccessPurpose::Interactive});
    g_secretAccessAuthorizationTicks.erase(SecretAccessAuthorizationKey{std::move(canonicalId), kind, SecretAccessPurpose::Background});
}

void ClearSecretAccessAuthorization(std::wstring_view connectionId) noexcept
{
    std::wstring canonicalId;
    if (! TryGetCanonicalConnectionId(connectionId, canonicalId))
    {
        return;
    }

    std::scoped_lock lock(g_secretAccessAuthorizationMutex);
    std::erase_if(g_secretAccessAuthorizationTicks,
                  [&](const auto& entry) noexcept { return entry.first.connectionId == canonicalId; });
}

void ClearAllSecretAccessAuthorizations() noexcept
{
    std::scoped_lock lock(g_secretAccessAuthorizationMutex);
    g_secretAccessAuthorizationTicks.clear();
}

#ifdef ENABLE_TESTS
void SetSecretAccessAuthorizationTickForTesting(std::wstring_view connectionId,
                                                SecretKind kind,
                                                SecretAccessPurpose purpose,
                                                uint64_t tick) noexcept
{
    std::wstring canonicalId;
    if (! TryGetCanonicalConnectionId(connectionId, canonicalId))
    {
        return;
    }

    std::scoped_lock lock(g_secretAccessAuthorizationMutex);
    g_secretAccessAuthorizationTicks.insert_or_assign(SecretAccessAuthorizationKey{std::move(canonicalId), kind, purpose}, tick);
}

bool IsSecretAccessAuthorizationFreshForTesting(uint64_t now, uint64_t authorizedAt, uint64_t timeoutMs) noexcept
{
    return IsAuthorizationFresh(now, authorizedAt, timeoutMs);
}
#endif
} // namespace RedSalamander::Connections
