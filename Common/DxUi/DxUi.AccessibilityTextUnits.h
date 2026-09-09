#pragma once

#include <cstddef>
#include <string_view>

#include <UIAutomation.h>
#include <objbase.h>

namespace RedSalamander::DxUi
{
struct TextRangeUnitMoveResult final
{
    size_t position = 0u;
    int moved       = 0;
};

struct AccessibilityTextUnitSpan final
{
    size_t start = 0u;
    size_t end   = 0u;
};

// Pure UIA boundary policy shared by DxUi controls and plugin-owned providers.
// Unsupported units advance to the next larger supported unit as required by
// UI Automation: Format -> Word, Paragraph -> Line, Page -> Document.
[[nodiscard]] TextUnit NormalizeAccessibilityTextUnit(TextUnit unit) noexcept;
[[nodiscard]] TextRangeUnitMoveResult MoveAccessibilityTextPositionByUnit(std::wstring_view text, size_t position, TextUnit unit, int count) noexcept;
[[nodiscard]] AccessibilityTextUnitSpan GetEnclosingAccessibilityTextUnitSpan(std::wstring_view text, size_t position, TextUnit unit) noexcept;
} // namespace RedSalamander::DxUi
