#pragma once

#include <string>
#include <string_view>
#include <vector>

namespace RedConfigure::Localization
{
enum class PlaceholderStatus
{
    Ok,
    BarePlaceholder,
    UnindexedFormatSpec,
    PrintfPlaceholder,
    PlaceholderMismatch,
};

struct PlaceholderValidationResult
{
    PlaceholderStatus status = PlaceholderStatus::Ok;
    std::vector<std::wstring> sourcePlaceholders;
    std::vector<std::wstring> targetPlaceholders;
};

[[nodiscard]] PlaceholderValidationResult ValidatePlaceholders(std::wstring_view source, std::wstring_view target);
} // namespace RedConfigure::Localization
