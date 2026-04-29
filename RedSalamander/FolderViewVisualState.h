#pragma once

#include <algorithm>

namespace FolderViewVisualState
{
inline constexpr float kFocusBorderOpacityUnfocused = 0.60f;
inline constexpr float kUnfocusedPaneTextOpacity    = 0.58f;
inline constexpr float kUnfocusedPaneIconOpacity    = 0.55f;

[[nodiscard]] constexpr float ClampOpacity(float opacity) noexcept
{
    return std::clamp(opacity, 0.0f, 1.0f);
}

[[nodiscard]] constexpr float ApplyOpacityScale(float opacity, float scale) noexcept
{
    return ClampOpacity(opacity * scale);
}

[[nodiscard]] constexpr float ResolveNormalTextAlpha(float baseAlpha, bool paneFocused, bool selected) noexcept
{
    return (! paneFocused && ! selected) ? ApplyOpacityScale(baseAlpha, kUnfocusedPaneTextOpacity) : ClampOpacity(baseAlpha);
}

[[nodiscard]] constexpr float ResolveFocusBorderAlpha(float baseAlpha, bool paneFocused) noexcept
{
    return paneFocused ? ClampOpacity(baseAlpha) : ApplyOpacityScale(baseAlpha, kFocusBorderOpacityUnfocused);
}

[[nodiscard]] constexpr float ResolveNormalIconOpacity(float baseOpacity, bool paneFocused) noexcept
{
    return paneFocused ? ClampOpacity(baseOpacity) : ApplyOpacityScale(baseOpacity, kUnfocusedPaneIconOpacity);
}

[[nodiscard]] constexpr float ResolvePlaceholderIconOpacity(bool paneFocused) noexcept
{
    return ResolveNormalIconOpacity(0.4f, paneFocused);
}
} // namespace FolderViewVisualState
