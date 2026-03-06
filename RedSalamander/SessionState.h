#pragma once

#include <cstdint>
#include <filesystem>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace SessionState
{
enum class OperationKind : uint8_t
{
    Unknown = 0,
    Browse,
    Copy,
    Compare,
};

struct State
{
    // One or more FileSystem plugin ids (e.g. browse uses one; compare/copy may involve two).
    std::vector<std::wstring> activeFileSystemPluginIds;
    OperationKind lastOperation = OperationKind::Unknown;
};

// Best-effort: removes the marker file so crash-quarantine does not trigger on clean shutdown.
void Clear() noexcept;

// Best-effort: writes a small marker file for crash-quarantine.
void UpdateActiveFileSystemPluginIdsAndOperation(std::initializer_list<std::wstring_view> pluginIds, OperationKind operation) noexcept;

[[nodiscard]] std::optional<State> TryRead() noexcept;

// Exposed for diagnostics and crash-quarantine.
[[nodiscard]] std::filesystem::path GetSessionStatePath() noexcept;
} // namespace SessionState
