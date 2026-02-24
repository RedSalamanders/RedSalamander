#pragma once

#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

namespace MaskSyntax
{
struct WildcardMask final
{
    std::vector<std::wstring> includePatterns;
    std::vector<std::wstring> excludePatterns;
};

constexpr size_t kWildcardMaskHistoryMaxItems = 10u;

[[nodiscard]] WildcardMask ParseWildcardMask(std::wstring_view rawText);
[[nodiscard]] bool MatchesWildcardMask(std::wstring_view text, const WildcardMask& mask) noexcept;

void NormalizeWildcardMaskHistory(std::vector<std::wstring>& history, size_t maxItems = kWildcardMaskHistoryMaxItems);
void AddToWildcardMaskHistory(std::vector<std::wstring>& history, size_t maxItems, std::wstring_view entry);
} // namespace MaskSyntax

