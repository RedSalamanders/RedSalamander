#pragma once

#include <array>
#include <cstddef>
#include <cstdint>
#include <functional>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#ifndef COMMON_API
#ifdef COMMON_EXPORTS
#define COMMON_API __declspec(dllexport)
#else
#define COMMON_API __declspec(dllimport)
#endif
#endif

namespace Common::Settings
{
enum class ThemeColorSourceKind : uint8_t
{
    Direct,
    Reference,
    Lighten,
    Darken,
    Alpha,
    Blend,
    Contrast,
    PerceptualTone,
    EnsureContrast,
    Harmonize,
    SystemColor,
    Tone,
    SeededRainbow,
    SeededChoice,
};

enum class ThemeSystemColorRole : uint8_t
{
    Accent,
    AccentLight,
    AccentDark,
    Window,
    WindowText,
    Highlight,
    HighlightText,
    Count,
};

struct ThemeColorSource
{
    ThemeColorSourceKind kind = ThemeColorSourceKind::Direct;
    uint32_t directArgb       = 0xFF000000u;
    std::vector<std::wstring> references;
    std::array<double, 4> parameters{};
    ThemeSystemColorRole systemRole   = ThemeSystemColorRole::Accent;
    bool preserveSystemAccentSpelling = false;

    ThemeColorSource() = default;
    ThemeColorSource(uint32_t argb) noexcept : directArgb(argb)
    {
    }

    bool operator==(const ThemeColorSource&) const = default;
};

enum class CompiledThemeColorKind : uint8_t
{
    SeededRainbow,
    SeededChoice,
};

struct CompiledThemeColor
{
    CompiledThemeColorKind kind = CompiledThemeColorKind::SeededRainbow;
    std::array<double, 4> parameters{};
    std::vector<uint32_t> candidates;
    uint32_t fallbackArgb = 0xFF000000u;

    bool operator==(const CompiledThemeColor&) const = default;
};

struct ResolvedThemeColors
{
    std::unordered_map<std::wstring, uint32_t> paletteColors;
    std::unordered_map<std::wstring, uint32_t> colors;
    std::unordered_map<std::wstring, CompiledThemeColor> dynamicColors;
    std::unordered_map<std::wstring, std::vector<std::wstring>> dependencies;
    std::unordered_map<std::wstring, std::vector<std::wstring>> affected;
};

enum class ThemeColorEvaluationPhase : uint8_t
{
    Load,
    Event,
    Paint,
};

enum class ThemeDynamicContextKind : uint8_t
{
    None,
    StableItemHash32,
};

using ThemeBaseColorLookup = std::function<std::optional<uint32_t>(std::wstring_view)>;

struct ThemeResolutionContext
{
    ThemeBaseColorLookup baseColor;
    bool effectiveDark = false;
    bool highContrast  = false;
    std::array<uint32_t, static_cast<size_t>(ThemeSystemColorRole::Count)> systemColors{};
};

struct ThemeRuntimeContext
{
    uint32_t seedHash32 = 0u;
    bool highContrast   = false;
};

struct ThemeDefinition;

COMMON_API ThemeResolutionContext MakeSystemThemeResolutionContext(bool effectiveDark) noexcept;
COMMON_API HRESULT ParseThemeColorSource(std::wstring_view text, ThemeColorSource& outSource, std::wstring* outMessage = nullptr) noexcept;
COMMON_API std::wstring FormatThemeColorSource(const ThemeColorSource& source);
COMMON_API bool IsPaintTimeThemeColorSource(const ThemeColorSource& source) noexcept;
COMMON_API bool IsEventTimeThemeColorSource(const ThemeColorSource& source) noexcept;
COMMON_API ThemeColorEvaluationPhase GetThemeColorEvaluationPhase(const ThemeColorSource& source) noexcept;
COMMON_API ThemeDynamicContextKind GetDynamicThemeColorContext(std::wstring_view semanticKey) noexcept;
COMMON_API HRESULT ResolveThemeDefinition(const ThemeDefinition& theme,
                                          const ThemeResolutionContext& context,
                                          ResolvedThemeColors& outResolved,
                                          std::wstring* outMessage = nullptr) noexcept;
COMMON_API uint32_t EvaluateDynamicThemeColor(const CompiledThemeColor& program, const ThemeRuntimeContext& context) noexcept;
} // namespace Common::Settings
