#pragma once

#include <cstddef>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "DxUi/DxUi.h"

namespace BatchRenameMenus
{
enum class HelperMenuKind : uint8_t
{
    Template,
    RegexSearch,
    Replacement,
};

struct HelperInsertionResult final
{
    std::wstring text;
    size_t selectionStart = 0u;
    size_t selectionEnd   = 0u;
};

struct HelperCommandInsertion final
{
    std::wstring insertionText;
    size_t selectionStart = 0u;
    size_t selectionEnd   = 0u;
};

[[nodiscard]] std::vector<RedSalamander::DxUi::MenuFlyoutItem> BuildTemplateHelperMenuItems();
[[nodiscard]] std::vector<RedSalamander::DxUi::MenuFlyoutItem> BuildRegexSearchHelperMenuItems();
[[nodiscard]] std::vector<RedSalamander::DxUi::MenuFlyoutItem> BuildReplacementHelperMenuItems();
[[nodiscard]] std::vector<RedSalamander::DxUi::MenuFlyoutItem> BuildHelperMenuItems(HelperMenuKind kind);

[[nodiscard]] int RegexEscapedLiteralHelperCommandId() noexcept;
[[nodiscard]] int ReplacementCustomSubexpressionHelperCommandId() noexcept;
[[nodiscard]] std::optional<std::wstring_view> TryGetHelperInsertionText(int commandId) noexcept;
[[nodiscard]] std::optional<HelperCommandInsertion> TryBuildDynamicHelperInsertion(int commandId, std::wstring_view selectedText);
[[nodiscard]] HelperInsertionResult ApplyHelperInsertion(std::wstring_view text, size_t selectionStart, size_t selectionEnd, std::wstring_view insertion);
} // namespace BatchRenameMenus
