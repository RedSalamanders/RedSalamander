#pragma once

#include <cstdint>
#include <filesystem>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace RedConfigure
{
struct PageDefinition
{
    std::wstring_view id;
    uint32_t titleResourceId       = 0u;
    uint32_t descriptionResourceId = 0u;
};

struct ThemePreviewHitCandidate
{
    std::wstring key;
    float left   = 0.0f;
    float top    = 0.0f;
    float right  = 0.0f;
    float bottom = 0.0f;
};

[[nodiscard]] std::span<const PageDefinition> GetPageDefinitions() noexcept;
[[nodiscard]] std::filesystem::path ResolveWorkspaceRootForLaunchPath(const std::filesystem::path& startPath);
[[nodiscard]] std::vector<std::wstring> FilterThemeColorKeys(std::span<const std::wstring> keys, std::wstring_view filterText);
[[nodiscard]] std::wstring SelectThemePreviewHitKey(std::span<const ThemePreviewHitCandidate> candidates, float x, float y, std::wstring_view previousKey);
[[nodiscard]] std::vector<std::wstring> BuildThemeColorSuggestions(std::wstring_view selectedKey,
                                                                   std::wstring_view previousKey,
                                                                   std::optional<uint32_t> currentColor);
} // namespace RedConfigure
