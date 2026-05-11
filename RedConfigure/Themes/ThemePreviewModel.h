#pragma once

#include "SettingsStore.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

namespace RedConfigure::Themes
{
enum class ThemeColorExpressionKind : uint8_t
{
    Reference,
    Lighten,
    Darken,
    Alpha,
    Blend,
    Contrast,
};

struct ThemeColorExpression
{
    ThemeColorExpressionKind kind = ThemeColorExpressionKind::Reference;
    std::wstring firstKey;
    std::wstring secondKey;
    double amount = 0.0;
};

class ThemePreviewModel final
{
public:
    void SetTheme(const Common::Settings::ThemeDefinition& theme);

    [[nodiscard]] const Common::Settings::ThemeDefinition& GetTheme() const noexcept;
    [[nodiscard]] Common::Settings::ThemeDefinition BuildFlattenedTheme() const;
    [[nodiscard]] std::optional<uint32_t> GetEffectiveColor(std::wstring_view key) const;
    [[nodiscard]] std::wstring GetAuthoredColorText(std::wstring_view key) const;
    [[nodiscard]] bool TryEditOverride(std::wstring_view key, std::wstring_view colorText);
    void ResetOverride(std::wstring_view key);

private:
    [[nodiscard]] std::optional<uint32_t> ResolveColor(std::wstring_view key, std::vector<std::wstring>& stack) const;
    [[nodiscard]] std::optional<uint32_t> ResolveExpression(const ThemeColorExpression& expression, std::vector<std::wstring>& stack) const;

    Common::Settings::ThemeDefinition _theme;
    std::unordered_map<std::wstring, ThemeColorExpression> _expressions;
};
} // namespace RedConfigure::Themes
