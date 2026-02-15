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
};

[[nodiscard]] std::wstring BuildCredentialTargetName(std::wstring_view connectionId, SecretKind kind);

HRESULT SaveGenericCredential(std::wstring_view targetName, std::wstring_view userName, std::wstring_view secret) noexcept;
HRESULT LoadGenericCredential(std::wstring_view targetName, std::wstring& userNameOut, std::wstring& secretOut) noexcept;
HRESULT DeleteGenericCredential(std::wstring_view targetName) noexcept;

inline constexpr std::wstring_view kQuickConnectConnectionId   = L"00000000-0000-0000-0000-000000000001";
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
} // namespace RedSalamander::Connections
