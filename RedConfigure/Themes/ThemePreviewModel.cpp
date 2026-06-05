#include "ThemePreviewModel.h"

#include "ThemeDefinitionIo.h"

#include <algorithm>
#include <array>
#include <cwchar>
#include <cwctype>
#include <string>
#include <utility>
#include <vector>

namespace
{
struct DefaultColor
{
    std::wstring_view key;
    uint32_t darkValue  = 0u;
    uint32_t lightValue = 0u;
};

constexpr std::array<DefaultColor, 26> kDefaultColors = {{
    {L"app.accent", 0xFF0078D4u, 0xFF005FB8u},
    {L"window.background", 0xFF202020u, 0xFFFFFFFFu},
    {L"window.text", 0xFFF3F3F3u, 0xFF111111u},
    {L"window.subduedText", 0xFFB8B8B8u, 0xFF5B5B5Bu},
    {L"navigation.background", 0xFF252525u, 0xFFEAF2FEu},
    {L"navigation.text", 0xFFF3F3F3u, 0xFF1F2937u},
    {L"navigation.accent", 0xFF0078D4u, 0xFF005FB8u},
    {L"menu.background", 0xFF2B2B2Bu, 0xFFFFFFFFu},
    {L"menu.text", 0xFFF4F4F4u, 0xFF111111u},
    {L"menu.selectionBackground", 0xFF3B3B3Bu, 0xFFE8F1FFu},
    {L"menu.selectionText", 0xFFFFFFFFu, 0xFF0F172Au},
    {L"menu.border", 0xFF505050u, 0xFFD8D8D8u},
    {L"folderView.background", 0xFF1E1E1Eu, 0xFFFFFFFFu},
    {L"folderView.itemForeground", 0xFFEDEDEDu, 0xFF111111u},
    {L"folderView.itemBackgroundHovered", 0xFF333333u, 0xFFF0F6FFu},
    {L"folderView.itemBackgroundSelected", 0xFF264F78u, 0xFFCFE8FFu},
    {L"folderView.itemForegroundSelected", 0xFFFFFFFFu, 0xFF0F172Au},
    {L"folderView.warningForeground", 0xFFFFD166u, 0xFF8A4B00u},
    {L"dialog.background", 0xFF2D2D2Du, 0xFFF7F7F7u},
    {L"dialog.text", 0xFFF3F3F3u, 0xFF111111u},
    {L"dialog.buttonBackground", 0xFF3A3A3Au, 0xFFFFFFFFu},
    {L"dialog.buttonText", 0xFFF3F3F3u, 0xFF111111u},
    {L"progress.background", 0xFF3A3A3Au, 0xFFE5E7EBu},
    {L"progress.fill", 0xFF0078D4u, 0xFF2563EBu},
    {L"diff.addedBackground", 0xFF173B24u, 0xFFEAF8EFu},
    {L"diff.removedBackground", 0xFF4A1F25u, 0xFFFFECEFu},
}};

[[nodiscard]] bool IsLightBase(std::wstring_view baseThemeId) noexcept
{
    return baseThemeId == L"builtin/light" || baseThemeId == L"builtin/system";
}

[[nodiscard]] std::wstring TrimCopy(std::wstring_view text)
{
    const auto isSpace = [](wchar_t ch) noexcept { return ch == L' ' || ch == L'\t' || ch == L'\r' || ch == L'\n'; };

    while (! text.empty() && isSpace(text.front()))
    {
        text.remove_prefix(1u);
    }
    while (! text.empty() && isSpace(text.back()))
    {
        text.remove_suffix(1u);
    }

    return std::wstring(text);
}

[[nodiscard]] std::wstring ToLowerCopy(std::wstring_view text)
{
    std::wstring result(text);
    std::transform(result.begin(), result.end(), result.begin(), [](wchar_t ch) noexcept { return static_cast<wchar_t>(::towlower(ch)); });
    return result;
}

[[nodiscard]] std::vector<std::wstring> SplitArguments(std::wstring_view text)
{
    std::vector<std::wstring> args;
    size_t start = 0u;
    while (start <= text.size())
    {
        const size_t comma = text.find(L',', start);
        const size_t end   = comma == std::wstring_view::npos ? text.size() : comma;
        args.push_back(TrimCopy(text.substr(start, end - start)));
        if (comma == std::wstring_view::npos)
        {
            break;
        }
        start = comma + 1u;
    }
    return args;
}

[[nodiscard]] bool ParseAmount(std::wstring_view text, double& outAmount) noexcept
{
    std::wstring trimmed = TrimCopy(text);
    if (trimmed.empty())
    {
        return false;
    }

    bool percent = false;
    if (trimmed.back() == L'%')
    {
        percent = true;
        trimmed.pop_back();
        trimmed = TrimCopy(trimmed);
        if (trimmed.empty())
        {
            return false;
        }
    }

    wchar_t* end        = nullptr;
    const double parsed = std::wcstod(trimmed.c_str(), &end);
    if (! end || *end != L'\0')
    {
        return false;
    }

    outAmount = percent ? parsed / 100.0 : parsed;
    return outAmount >= 0.0 && outAmount <= 1.0;
}

[[nodiscard]] bool ParseExpressionText(std::wstring_view text, RedConfigure::Themes::ThemeColorExpression& outExpression)
{
    const std::wstring trimmed = TrimCopy(text);
    const size_t open          = trimmed.find(L'(');
    const size_t close         = trimmed.rfind(L')');
    if (open == std::wstring::npos || close == std::wstring::npos || close <= open || close != trimmed.size() - 1u)
    {
        return false;
    }

    const std::wstring function          = ToLowerCopy(TrimCopy(std::wstring_view(trimmed).substr(0u, open)));
    const std::vector<std::wstring> args = SplitArguments(std::wstring_view(trimmed).substr(open + 1u, close - open - 1u));

    using RedConfigure::Themes::ThemeColorExpressionKind;
    RedConfigure::Themes::ThemeColorExpression expression;
    if (function == L"ref")
    {
        if (args.size() != 1u || ! Common::Settings::IsValidThemeColorKey(args[0]))
        {
            return false;
        }
        expression.kind     = ThemeColorExpressionKind::Reference;
        expression.firstKey = args[0];
    }
    else if (function == L"lighten" || function == L"darken" || function == L"alpha")
    {
        if (args.size() != 2u || ! Common::Settings::IsValidThemeColorKey(args[0]) || ! ParseAmount(args[1], expression.amount))
        {
            return false;
        }
        expression.kind     = (function == L"lighten")  ? ThemeColorExpressionKind::Lighten
                              : (function == L"darken") ? ThemeColorExpressionKind::Darken
                                                        : ThemeColorExpressionKind::Alpha;
        expression.firstKey = args[0];
    }
    else if (function == L"blend")
    {
        if (args.size() != 3u || ! Common::Settings::IsValidThemeColorKey(args[0]) || ! Common::Settings::IsValidThemeColorKey(args[1]) ||
            ! ParseAmount(args[2], expression.amount))
        {
            return false;
        }
        expression.kind      = ThemeColorExpressionKind::Blend;
        expression.firstKey  = args[0];
        expression.secondKey = args[1];
    }
    else if (function == L"contrast")
    {
        if (args.size() != 1u || ! Common::Settings::IsValidThemeColorKey(args[0]))
        {
            return false;
        }
        expression.kind     = ThemeColorExpressionKind::Contrast;
        expression.firstKey = args[0];
    }
    else
    {
        return false;
    }

    outExpression = std::move(expression);
    return true;
}

[[nodiscard]] uint32_t Channel(uint32_t argb, uint32_t shift) noexcept
{
    return (argb >> shift) & 0xFFu;
}

[[nodiscard]] uint32_t PackArgb(uint32_t a, uint32_t r, uint32_t g, uint32_t b) noexcept
{
    return ((a & 0xFFu) << 24u) | ((r & 0xFFu) << 16u) | ((g & 0xFFu) << 8u) | (b & 0xFFu);
}

[[nodiscard]] uint32_t MixChannel(uint32_t from, uint32_t to, double amount) noexcept
{
    const double mixed   = static_cast<double>(from) + ((static_cast<double>(to) - static_cast<double>(from)) * amount);
    const double clamped = std::clamp(mixed, 0.0, 255.0);
    return static_cast<uint32_t>(clamped + 0.5);
}

[[nodiscard]] uint32_t MixColors(uint32_t from, uint32_t to, double amount) noexcept
{
    return PackArgb(MixChannel(Channel(from, 24u), Channel(to, 24u), amount),
                    MixChannel(Channel(from, 16u), Channel(to, 16u), amount),
                    MixChannel(Channel(from, 8u), Channel(to, 8u), amount),
                    MixChannel(Channel(from, 0u), Channel(to, 0u), amount));
}

[[nodiscard]] uint32_t SetAlpha(uint32_t argb, double amount) noexcept
{
    return PackArgb(MixChannel(0u, 255u, amount), Channel(argb, 16u), Channel(argb, 8u), Channel(argb, 0u));
}

[[nodiscard]] uint32_t RelativeBrightness(uint32_t argb) noexcept
{
    return ((Channel(argb, 16u) * 299u) + (Channel(argb, 8u) * 587u) + (Channel(argb, 0u) * 114u)) / 1000u;
}

[[nodiscard]] std::wstring FormatAmount(double amount)
{
    const uint32_t percent = static_cast<uint32_t>((amount * 100.0) + 0.5);
    return std::to_wstring(percent) + L"%";
}

[[nodiscard]] std::wstring FormatExpression(const RedConfigure::Themes::ThemeColorExpression& expression)
{
    using RedConfigure::Themes::ThemeColorExpressionKind;
    switch (expression.kind)
    {
        case ThemeColorExpressionKind::Reference: return L"ref(" + expression.firstKey + L")";
        case ThemeColorExpressionKind::Lighten: return L"lighten(" + expression.firstKey + L"," + FormatAmount(expression.amount) + L")";
        case ThemeColorExpressionKind::Darken: return L"darken(" + expression.firstKey + L"," + FormatAmount(expression.amount) + L")";
        case ThemeColorExpressionKind::Alpha: return L"alpha(" + expression.firstKey + L"," + FormatAmount(expression.amount) + L")";
        case ThemeColorExpressionKind::Blend:
            return L"blend(" + expression.firstKey + L"," + expression.secondKey + L"," + FormatAmount(expression.amount) + L")";
        case ThemeColorExpressionKind::Contrast: return L"contrast(" + expression.firstKey + L")";
        default: return {};
    }
}
} // namespace

namespace RedConfigure::Themes
{
void ThemePreviewModel::SetTheme(const Common::Settings::ThemeDefinition& theme)
{
    _theme = theme;
    _expressions.clear();
}

const Common::Settings::ThemeDefinition& ThemePreviewModel::GetTheme() const noexcept
{
    return _theme;
}

Common::Settings::ThemeDefinition ThemePreviewModel::BuildFlattenedTheme() const
{
    Common::Settings::ThemeDefinition flattened = _theme;
    for (const auto& [key, _] : _expressions)
    {
        if (const std::optional<uint32_t> color = GetEffectiveColor(key))
        {
            flattened.colors[key] = color.value();
        }
    }
    return flattened;
}

std::optional<uint32_t> ThemePreviewModel::GetEffectiveColor(std::wstring_view key) const
{
    std::vector<std::wstring> stack;
    return ResolveColor(key, stack);
}

std::wstring ThemePreviewModel::GetAuthoredColorText(std::wstring_view key) const
{
    if (const auto expression = _expressions.find(std::wstring(key)); expression != _expressions.end())
    {
        return FormatExpression(expression->second);
    }
    if (const auto color = _theme.colors.find(std::wstring(key)); color != _theme.colors.end())
    {
        return Common::Settings::FormatColor(color->second);
    }
    return {};
}

std::optional<uint32_t> ThemePreviewModel::ResolveColor(std::wstring_view key, std::vector<std::wstring>& stack) const
{
    const std::wstring keyText(key);
    if (std::find(stack.begin(), stack.end(), keyText) != stack.end())
    {
        return std::nullopt;
    }

    if (const auto expression = _expressions.find(keyText); expression != _expressions.end())
    {
        stack.push_back(keyText);
        const std::optional<uint32_t> color = ResolveExpression(expression->second, stack);
        stack.pop_back();
        return color;
    }

    if (const auto it = _theme.colors.find(std::wstring(key)); it != _theme.colors.end())
    {
        return it->second;
    }

    const bool light = IsLightBase(_theme.baseThemeId);
    for (const DefaultColor& color : kDefaultColors)
    {
        if (color.key == key)
        {
            return light ? color.lightValue : color.darkValue;
        }
    }

    return std::nullopt;
}

std::optional<uint32_t> ThemePreviewModel::ResolveExpression(const ThemeColorExpression& expression, std::vector<std::wstring>& stack) const
{
    const std::optional<uint32_t> first = ResolveColor(expression.firstKey, stack);
    if (! first)
    {
        return std::nullopt;
    }

    switch (expression.kind)
    {
        case ThemeColorExpressionKind::Reference: return first;
        case ThemeColorExpressionKind::Lighten: return MixColors(first.value(), 0xFFFFFFFFu, expression.amount);
        case ThemeColorExpressionKind::Darken: return MixColors(first.value(), 0xFF000000u, expression.amount);
        case ThemeColorExpressionKind::Alpha: return SetAlpha(first.value(), expression.amount);
        case ThemeColorExpressionKind::Blend:
        {
            const std::optional<uint32_t> second = ResolveColor(expression.secondKey, stack);
            if (! second)
            {
                return std::nullopt;
            }
            return MixColors(first.value(), second.value(), expression.amount);
        }
        case ThemeColorExpressionKind::Contrast:
            return RelativeBrightness(first.value()) >= 128u ? std::optional<uint32_t>(0xFF000000u) : std::optional<uint32_t>(0xFFFFFFFFu);
        default: return std::nullopt;
    }
}

bool ThemePreviewModel::TryEditOverride(std::wstring_view key, std::wstring_view colorText)
{
    if (! Common::Settings::IsValidThemeColorKey(key))
    {
        return false;
    }

    const std::wstring authoredText = TrimCopy(colorText);
    uint32_t argb                   = 0u;
    if (! Common::Settings::TryParseColor(authoredText, argb))
    {
        ThemeColorExpression expression;
        if (! ParseExpressionText(authoredText, expression))
        {
            return false;
        }

        const std::wstring keyText(key);
        auto previousExpressions = _expressions;
        auto previousColors      = _theme.colors;
        _theme.colors.erase(keyText);
        _expressions[keyText] = std::move(expression);
        if (! GetEffectiveColor(key))
        {
            _expressions  = std::move(previousExpressions);
            _theme.colors = std::move(previousColors);
            return false;
        }
        return true;
    }

    const std::wstring keyText(key);
    _expressions.erase(keyText);
    _theme.colors[keyText] = argb;
    return true;
}

void ThemePreviewModel::ResetOverride(std::wstring_view key)
{
    const std::wstring keyText(key);
    _expressions.erase(keyText);
    _theme.colors.erase(keyText);
}
} // namespace RedConfigure::Themes
