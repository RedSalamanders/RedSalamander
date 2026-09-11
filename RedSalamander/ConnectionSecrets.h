#pragma once

#include <string>
#include <string_view>

#include <cstdint>
#include <optional>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

namespace Common::Settings
{
struct ConnectionProfile;
}

namespace RedSalamander::Connections
{
enum class SecretKind : uint8_t
{
    Password,
    SshKeyPassphrase,
    RefreshToken,
};

enum class SecretAccessPurpose : uint8_t
{
    Interactive,
    Background,
};

[[nodiscard]] std::wstring BuildCredentialTargetName(std::wstring_view connectionId, SecretKind kind);

HRESULT SaveGenericCredential(std::wstring_view targetName, std::wstring_view userName, std::wstring_view secret) noexcept;
HRESULT LoadGenericCredential(std::wstring_view targetName, std::wstring& userNameOut, std::wstring& secretOut) noexcept;
HRESULT DeleteGenericCredential(std::wstring_view targetName) noexcept;

inline constexpr std::wstring_view kQuickConnectConnectionName = L"@quick";

[[nodiscard]] bool IsQuickConnectConnectionId(std::wstring_view connectionId) noexcept;
[[nodiscard]] bool IsQuickConnectConnectionName(std::wstring_view connectionName) noexcept;

// Quick Connect is an in-memory-only connection profile (never serialized to disk).
// The host stores its secret material in memory as well (not in WinCred).
// `preferredPluginId` is used only when Quick Connect has not been initialized yet.
void EnsureQuickConnectProfile(std::wstring_view preferredPluginId) noexcept;
void GetQuickConnectProfile(Common::Settings::ConnectionProfile& out) noexcept;
void SetQuickConnectProfile(const Common::Settings::ConnectionProfile& profile) noexcept;

[[nodiscard]] bool HasQuickConnectSecret(SecretKind kind) noexcept;
[[nodiscard]] HRESULT LoadQuickConnectSecret(SecretKind kind, std::wstring& secretOut) noexcept;
void SetQuickConnectSecret(SecretKind kind, std::wstring_view secret) noexcept;
void ClearQuickConnectSecret(SecretKind kind) noexcept;

// Tracks whether the user approved secret access for one profile and secret kind.
//
// Approval is recorded when the user:
// - completed Windows Hello verification, or
// - manually entered a password/passphrase to connect.
//
// Interactive access is valid only inside `reauthTimeoutMs`; zero always prompts.
// Background access uses the explicit app-run grant created by the same approval.
void NoteSecretAccessAuthorized(std::wstring_view connectionId, SecretKind kind) noexcept;
[[nodiscard]] bool IsSecretAccessAuthorized(std::wstring_view connectionId,
                                            SecretKind kind,
                                            SecretAccessPurpose purpose,
                                            uint64_t reauthTimeoutMs) noexcept;
void ClearSecretAccessAuthorization(std::wstring_view connectionId, SecretKind kind) noexcept;
void ClearSecretAccessAuthorization(std::wstring_view connectionId) noexcept;
void ClearAllSecretAccessAuthorizations() noexcept;

#ifdef ENABLE_TESTS
enum class CredentialPersistenceFault : uint8_t
{
    None,
    SaveOnce,
    DeleteOnce,
};

void SetCredentialPersistenceFaultForTesting(CredentialPersistenceFault fault) noexcept;
// Test hook: allows selftests to simulate an expired authorization timestamp without sleeping.
void SetSecretAccessAuthorizationTickForTesting(std::wstring_view connectionId,
                                                SecretKind kind,
                                                SecretAccessPurpose purpose,
                                                uint64_t tick) noexcept;
[[nodiscard]] bool IsSecretAccessAuthorizationFreshForTesting(uint64_t now, uint64_t authorizedAt, uint64_t timeoutMs) noexcept;
#endif
} // namespace RedSalamander::Connections
