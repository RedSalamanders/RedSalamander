#pragma once

#include <cstdint>
#include <filesystem>
#include <optional>
#include <string>
#include <string_view>

namespace Common::Settings
{
struct ConnectionProfile;
struct JsonValue;
struct Settings;
} // namespace Common::Settings

namespace ConnectionProfileUtils
{
[[nodiscard]] std::optional<bool> ExtraGetBool(const Common::Settings::JsonValue& extra, std::string_view key) noexcept;
[[nodiscard]] std::optional<uint32_t> ExtraGetUInt32(const Common::Settings::JsonValue& extra, std::string_view key) noexcept;

[[nodiscard]] const Common::Settings::ConnectionProfile* FindConnectionProfileByName(const Common::Settings::Settings* settings,
                                                                                     std::wstring_view connectionName) noexcept;

[[nodiscard]] std::optional<std::wstring> TryParseConnNameFromPluginPath(std::wstring_view pluginPath) noexcept;
[[nodiscard]] std::optional<std::wstring> TryParseConnNameFromPluginPath(const std::optional<std::filesystem::path>& pluginPath) noexcept;

[[nodiscard]] bool ConnectionProfileUsesInsecureTls(const Common::Settings::ConnectionProfile& profile) noexcept;
[[nodiscard]] std::wstring BuildConnectionDisplayUrl(const Common::Settings::ConnectionProfile& profile) noexcept;
} // namespace ConnectionProfileUtils
