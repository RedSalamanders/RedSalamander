#include "ThemeExpression.h"

#include "Helpers.h"
#include "SettingsStore.h"
#include "ThemeDefinitionIo.h"

#include <algorithm>
#include <array>
#include <charconv>
#include <chrono>
#include <cmath>
#include <cwctype>
#include <format>
#include <limits>
#include <numbers>
#include <ranges>
#include <span>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>
#include <dwmapi.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026, C5027
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#pragma comment(lib, "dwmapi.lib")

namespace
{
using Common::Settings::CompiledThemeColor;
using Common::Settings::CompiledThemeColorKind;
using Common::Settings::ResolvedThemeColors;
using Common::Settings::ThemeColorSource;
using Common::Settings::ThemeColorSourceKind;
using Common::Settings::ThemeDefinition;
using Common::Settings::ThemeResolutionContext;
using Common::Settings::ThemeSystemColorRole;

[[nodiscard]] std::wstring_view Trim(std::wstring_view text) noexcept
{
    while (! text.empty() && std::iswspace(text.front()) != 0)
    {
        text.remove_prefix(1u);
    }
    while (! text.empty() && std::iswspace(text.back()) != 0)
    {
        text.remove_suffix(1u);
    }
    return text;
}

[[nodiscard]] std::wstring ToLower(std::wstring_view text)
{
    std::wstring result(text);
    std::ranges::transform(result, result.begin(), [](wchar_t ch) noexcept { return static_cast<wchar_t>(std::towlower(ch)); });
    return result;
}

[[nodiscard]] HRESULT Invalid(std::wstring* outMessage, std::wstring message) noexcept
{
    if (outMessage)
    {
        *outMessage = std::move(message);
    }
    return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

[[nodiscard]] bool ParseDouble(std::wstring_view text, double& out) noexcept
{
    text                                  = Trim(text);
    constexpr size_t kMaxNumberCharacters = 64u;
    if (text.empty() || text.size() > kMaxNumberCharacters)
    {
        return false;
    }

    std::array<char, kMaxNumberCharacters> ascii{};
    bool mantissaDigit = false;
    bool exponent      = false;
    bool exponentDigit = false;
    bool decimalPoint  = false;
    for (size_t index = 0u; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch >= L'0' && ch <= L'9')
        {
            ascii[index] = static_cast<char>(ch);
            if (exponent)
            {
                exponentDigit = true;
            }
            else
            {
                mantissaDigit = true;
            }
            continue;
        }
        if ((ch == L'+' || ch == L'-') && (index == 0u || (exponent && (text[index - 1u] == L'e' || text[index - 1u] == L'E'))))
        {
            ascii[index] = static_cast<char>(ch);
            continue;
        }
        if (ch == L'.' && ! decimalPoint && ! exponent)
        {
            ascii[index] = '.';
            decimalPoint = true;
            continue;
        }
        if ((ch == L'e' || ch == L'E') && mantissaDigit && ! exponent)
        {
            ascii[index] = static_cast<char>(ch);
            exponent     = true;
            continue;
        }
        return false;
    }

    if (! mantissaDigit || (exponent && ! exponentDigit))
    {
        return false;
    }

    const char* first = ascii.data();
    const char* last  = first + text.size();
    if (*first == '+')
    {
        ++first;
    }

    double value            = 0.0;
    const auto [end, error] = std::from_chars(first, last, value, std::chars_format::general);
    if (error != std::errc{} || end != last || ! std::isfinite(value))
    {
        return false;
    }
    out = value;
    return true;
}

[[nodiscard]] bool ParseAmount(std::wstring_view text, double& out) noexcept
{
    text         = Trim(text);
    bool percent = false;
    if (! text.empty() && text.back() == L'%')
    {
        text.remove_suffix(1u);
        percent = true;
    }
    if (! ParseDouble(text, out))
    {
        return false;
    }
    if (percent)
    {
        out /= 100.0;
    }
    return out >= 0.0 && out <= 1.0;
}

[[nodiscard]] std::vector<std::wstring_view> SplitArguments(std::wstring_view text)
{
    std::vector<std::wstring_view> result;
    size_t start = 0u;
    for (size_t index = 0u; index <= text.size(); ++index)
    {
        if (index != text.size() && text[index] != L',')
        {
            if (text[index] == L'(' || text[index] == L')')
            {
                return {};
            }
            continue;
        }
        const std::wstring_view part = Trim(text.substr(start, index - start));
        if (part.empty())
        {
            return {};
        }
        result.push_back(part);
        start = index + 1u;
    }
    return result;
}

[[nodiscard]] bool IsValidReference(std::wstring_view reference) noexcept
{
    if (reference.empty() || reference.size() > 96u)
    {
        return false;
    }
    if (reference.rfind(L"palette.", 0u) == 0u)
    {
        reference.remove_prefix(8u);
        if (reference.empty() || reference.find(L'.') != std::wstring_view::npos)
        {
            return false;
        }
    }
    return Common::Settings::IsValidThemeColorKey(reference);
}

[[nodiscard]] std::optional<ThemeSystemColorRole> ParseSystemRole(std::wstring_view text) noexcept
{
    const std::wstring lower = ToLower(Trim(text));
    if (lower == L"accent")
        return ThemeSystemColorRole::Accent;
    if (lower == L"accentlight")
        return ThemeSystemColorRole::AccentLight;
    if (lower == L"accentdark")
        return ThemeSystemColorRole::AccentDark;
    if (lower == L"window")
        return ThemeSystemColorRole::Window;
    if (lower == L"windowtext")
        return ThemeSystemColorRole::WindowText;
    if (lower == L"highlight")
        return ThemeSystemColorRole::Highlight;
    if (lower == L"highlighttext")
        return ThemeSystemColorRole::HighlightText;
    return std::nullopt;
}

[[nodiscard]] std::wstring SystemRoleText(ThemeSystemColorRole role)
{
    switch (role)
    {
        case ThemeSystemColorRole::Accent: return L"accent";
        case ThemeSystemColorRole::AccentLight: return L"accentLight";
        case ThemeSystemColorRole::AccentDark: return L"accentDark";
        case ThemeSystemColorRole::Window: return L"window";
        case ThemeSystemColorRole::WindowText: return L"windowText";
        case ThemeSystemColorRole::Highlight: return L"highlight";
        case ThemeSystemColorRole::HighlightText: return L"highlightText";
        case ThemeSystemColorRole::Count: break;
    }
    return L"accent";
}

[[nodiscard]] std::wstring NumberText(double value)
{
    return std::format(L"{:.6g}", value);
}

[[nodiscard]] uint8_t Channel(uint32_t argb, uint32_t shift) noexcept
{
    return static_cast<uint8_t>((argb >> shift) & 0xFFu);
}

[[nodiscard]] uint32_t Pack(uint8_t a, uint8_t r, uint8_t g, uint8_t b) noexcept
{
    return (static_cast<uint32_t>(a) << 24u) | (static_cast<uint32_t>(r) << 16u) | (static_cast<uint32_t>(g) << 8u) | static_cast<uint32_t>(b);
}

[[nodiscard]] uint8_t LerpChannel(uint8_t first, uint8_t second, double amount) noexcept
{
    return static_cast<uint8_t>(std::clamp(std::lround(static_cast<double>(first) + (static_cast<double>(second) - first) * amount), 0l, 255l));
}

[[nodiscard]] uint32_t Blend(uint32_t first, uint32_t second, double amount) noexcept
{
    return Pack(LerpChannel(Channel(first, 24u), Channel(second, 24u), amount),
                LerpChannel(Channel(first, 16u), Channel(second, 16u), amount),
                LerpChannel(Channel(first, 8u), Channel(second, 8u), amount),
                LerpChannel(Channel(first, 0u), Channel(second, 0u), amount));
}

[[nodiscard]] double SrgbToLinear(double channel) noexcept
{
    return channel <= 0.04045 ? channel / 12.92 : std::pow((channel + 0.055) / 1.055, 2.4);
}

[[nodiscard]] double LinearToSrgb(double channel) noexcept
{
    return channel <= 0.0031308 ? channel * 12.92 : 1.055 * std::pow(channel, 1.0 / 2.4) - 0.055;
}

struct Oklch
{
    double lightness = 0.0;
    double chroma    = 0.0;
    double hue       = 0.0;
    uint8_t alpha    = 0xFFu;
};

[[nodiscard]] Oklch ToOklch(uint32_t argb) noexcept
{
    const double r = SrgbToLinear(Channel(argb, 16u) / 255.0);
    const double g = SrgbToLinear(Channel(argb, 8u) / 255.0);
    const double b = SrgbToLinear(Channel(argb, 0u) / 255.0);

    const double l    = std::cbrt(0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * b);
    const double m    = std::cbrt(0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * b);
    const double s    = std::cbrt(0.0883024619 * r + 0.2817188376 * g + 0.6299787005 * b);
    const double labL = 0.2104542553 * l + 0.7936177850 * m - 0.0040720468 * s;
    const double labA = 1.9779984951 * l - 2.4285922050 * m + 0.4505937099 * s;
    const double labB = 0.0259040371 * l + 0.7827717662 * m - 0.8086757660 * s;
    double hue        = std::atan2(labB, labA) * 180.0 / std::numbers::pi;
    if (hue < 0.0)
    {
        hue += 360.0;
    }
    return {.lightness = labL, .chroma = std::hypot(labA, labB), .hue = hue, .alpha = Channel(argb, 24u)};
}

struct LinearRgb
{
    double r = 0.0;
    double g = 0.0;
    double b = 0.0;
};

[[nodiscard]] LinearRgb OklchToLinearRgb(const Oklch& color, double chroma) noexcept
{
    const double radians = color.hue * std::numbers::pi / 180.0;
    const double labA    = chroma * std::cos(radians);
    const double labB    = chroma * std::sin(radians);
    const double lPrime  = color.lightness + 0.3963377774 * labA + 0.2158037573 * labB;
    const double mPrime  = color.lightness - 0.1055613458 * labA - 0.0638541728 * labB;
    const double sPrime  = color.lightness - 0.0894841775 * labA - 1.2914855480 * labB;
    const double l       = lPrime * lPrime * lPrime;
    const double m       = mPrime * mPrime * mPrime;
    const double s       = sPrime * sPrime * sPrime;
    return {
        .r = +4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s,
        .g = -1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s,
        .b = -0.0041960863 * l - 0.7034186147 * m + 1.7076147010 * s,
    };
}

[[nodiscard]] bool InGamut(const LinearRgb& rgb) noexcept
{
    return rgb.r >= 0.0 && rgb.r <= 1.0 && rgb.g >= 0.0 && rgb.g <= 1.0 && rgb.b >= 0.0 && rgb.b <= 1.0;
}

[[nodiscard]] uint32_t FromOklch(Oklch color) noexcept
{
    color.lightness = std::clamp(color.lightness, 0.0, 1.0);
    color.chroma    = std::max(color.chroma, 0.0);
    double low      = 0.0;
    double high     = color.chroma;
    LinearRgb rgb   = OklchToLinearRgb(color, high);
    if (! InGamut(rgb))
    {
        for (int iteration = 0; iteration < 24; ++iteration)
        {
            const double middle       = (low + high) * 0.5;
            const LinearRgb candidate = OklchToLinearRgb(color, middle);
            if (InGamut(candidate))
            {
                low = middle;
                rgb = candidate;
            }
            else
            {
                high = middle;
            }
        }
    }
    const auto encode = [](double value) noexcept
    { return static_cast<uint8_t>(std::clamp(std::lround(LinearToSrgb(std::clamp(value, 0.0, 1.0)) * 255.0), 0l, 255l)); };
    return Pack(color.alpha, encode(rgb.r), encode(rgb.g), encode(rgb.b));
}

[[nodiscard]] double RelativeLuminance(uint32_t argb) noexcept
{
    return Common::Colors::RelativeLuminanceFromArgb(argb);
}

[[nodiscard]] double ContrastRatio(uint32_t first, uint32_t second) noexcept
{
    const double a = RelativeLuminance(first);
    const double b = RelativeLuminance(second);
    return Common::Colors::ContrastRatioFromRelativeLuminance(a, b);
}

[[nodiscard]] std::optional<double> RenderedContrastRatio(uint32_t foreground, uint32_t background) noexcept
{
    if (Channel(background, 24u) != 0xFFu)
    {
        return std::nullopt;
    }
    return ContrastRatio(Common::Colors::CompositeArgbOverOpaqueBackground(foreground, background), background);
}

[[nodiscard]] std::optional<uint32_t> EnsureContrast(uint32_t foreground, uint32_t background, double ratio) noexcept
{
    const std::optional<double> initialRatio = RenderedContrastRatio(foreground, background);
    if (! initialRatio.has_value())
    {
        return std::nullopt;
    }
    if (initialRatio.value() >= ratio)
    {
        return foreground;
    }
    const Oklch source = ToOklch(foreground);
    std::optional<uint32_t> best;
    double bestDistance = std::numeric_limits<double>::max();
    for (int tone = 0; tone <= 100; ++tone)
    {
        Oklch candidate                            = source;
        candidate.lightness                        = static_cast<double>(tone) / 100.0;
        const uint32_t argb                        = FromOklch(candidate);
        const std::optional<double> candidateRatio = RenderedContrastRatio(argb, background);
        if (! candidateRatio.has_value() || candidateRatio.value() < ratio)
        {
            continue;
        }
        const double distance = std::abs(candidate.lightness - source.lightness);
        if (distance < bestDistance)
        {
            bestDistance = distance;
            best         = argb;
        }
    }
    return best;
}

[[nodiscard]] uint32_t HsvToArgb(double hue, double saturation, double value, double alpha) noexcept
{
    hue = std::fmod(hue, 360.0);
    if (hue < 0.0)
        hue += 360.0;
    const double chroma = value * saturation;
    const double x      = chroma * (1.0 - std::abs(std::fmod(hue / 60.0, 2.0) - 1.0));
    const double m      = value - chroma;
    double r            = 0.0;
    double g            = 0.0;
    double b            = 0.0;
    if (hue < 60.0)
    {
        r = chroma;
        g = x;
    }
    else if (hue < 120.0)
    {
        r = x;
        g = chroma;
    }
    else if (hue < 180.0)
    {
        g = chroma;
        b = x;
    }
    else if (hue < 240.0)
    {
        g = x;
        b = chroma;
    }
    else if (hue < 300.0)
    {
        r = x;
        b = chroma;
    }
    else
    {
        r = chroma;
        b = x;
    }
    const auto channel = [m](double component) noexcept { return static_cast<uint8_t>(std::clamp(std::lround((component + m) * 255.0), 0l, 255l)); };
    return Pack(static_cast<uint8_t>(std::clamp(std::lround(alpha * 255.0), 0l, 255l)), channel(r), channel(g), channel(b));
}

struct Resolver
{
    enum class State : uint8_t
    {
        Unvisited,
        Visiting,
        Resolved,
    };

    const ThemeDefinition& theme;
    const ThemeResolutionContext& context;
    ResolvedThemeColors& output;
    std::unordered_map<std::wstring, uint32_t> paletteCache;
    std::unordered_map<std::wstring, State> states;
    std::vector<std::wstring> stack;
    std::wstring* message = nullptr;

    Resolver(const ThemeDefinition& themeValue,
             const ThemeResolutionContext& contextValue,
             ResolvedThemeColors& outputValue,
             std::wstring* messageValue) noexcept
        : theme(themeValue),
          context(contextValue),
          output(outputValue),
          message(messageValue)
    {
    }

    Resolver(const Resolver&)            = delete;
    Resolver& operator=(const Resolver&) = delete;
    Resolver(Resolver&&)                 = delete;
    Resolver& operator=(Resolver&&)      = delete;

    [[nodiscard]] HRESULT ResolveReference(std::wstring_view reference, uint32_t& out)
    {
        if (reference.rfind(L"palette.", 0u) == 0u)
        {
            const std::wstring name(reference.substr(8u));
            if (const auto cached = paletteCache.find(name); cached != paletteCache.end())
            {
                out = cached->second;
                return S_OK;
            }
            const auto source = theme.palette.find(name);
            if (source == theme.palette.end())
            {
                return Invalid(message, std::format(L"Missing palette reference '{}'.", reference));
            }
            return ResolveNamed(std::wstring(reference), source->second, true, out);
        }

        const std::wstring key(reference);
        if (output.dynamicColors.contains(key))
        {
            return Invalid(message, std::format(L"Paint-time source '{}' cannot be referenced.", reference));
        }
        if (const auto cached = output.colors.find(key); cached != output.colors.end())
        {
            out = cached->second;
            return S_OK;
        }
        if (const auto source = theme.colors.find(key); source != theme.colors.end())
        {
            if (Common::Settings::IsPaintTimeThemeColorSource(source->second))
            {
                return Invalid(message, std::format(L"Paint-time source '{}' cannot be referenced.", reference));
            }
            return ResolveNamed(key, source->second, false, out);
        }
        if (context.baseColor)
        {
            const std::optional<uint32_t> base = context.baseColor(reference);
            if (base.has_value())
            {
                out = base.value();
                return S_OK;
            }
        }
        return Invalid(message, std::format(L"Missing semantic reference '{}'.", reference));
    }

    [[nodiscard]] HRESULT ResolveNamed(const std::wstring& name, const ThemeColorSource& source, bool palette, uint32_t& out)
    {
        const auto state = states.find(name);
        if (state != states.end() && state->second == State::Resolved)
        {
            if (palette)
            {
                const auto cached = paletteCache.find(name.substr(8u));
                if (cached != paletteCache.end())
                {
                    out = cached->second;
                    return S_OK;
                }
            }
            else if (const auto cached = output.colors.find(name); cached != output.colors.end())
            {
                out = cached->second;
                return S_OK;
            }
        }
        if (state != states.end() && state->second == State::Visiting)
        {
            std::wstring path;
            for (const std::wstring& part : stack)
            {
                if (! path.empty())
                    path += L" -> ";
                path += part;
            }
            if (! path.empty())
                path += L" -> ";
            path += name;
            return Invalid(message, std::format(L"Theme color dependency cycle: {}.", path));
        }
        if (stack.size() >= 32u)
        {
            return Invalid(message, std::format(L"Theme color dependency depth exceeds 32 while resolving '{}'.", name));
        }
        states[name] = State::Visiting;
        stack.push_back(name);
        bool resolved             = false;
        const auto pop            = wil::scope_exit([&]
        {
            stack.pop_back();
            states[name] = resolved ? State::Resolved : State::Unvisited;
        });
        output.dependencies[name] = source.references;
        for (const std::wstring& reference : source.references)
        {
            output.affected[reference].push_back(name);
        }
        if (Common::Settings::IsPaintTimeThemeColorSource(source))
        {
            if (palette)
            {
                return Invalid(message, std::format(L"Palette entry '{}' cannot use a paint-time source.", name));
            }
            if (Common::Settings::GetDynamicThemeColorContext(name) == Common::Settings::ThemeDynamicContextKind::None)
            {
                return Invalid(message, std::format(L"Theme color '{}' does not accept a paint-time source.", name));
            }
            CompiledThemeColor compiled;
            compiled.fallbackArgb = context.baseColor ? context.baseColor(name).value_or(0xFF000000u) : 0xFF000000u;
            compiled.parameters   = source.parameters;
            if (source.kind == ThemeColorSourceKind::SeededRainbow)
            {
                compiled.kind = CompiledThemeColorKind::SeededRainbow;
            }
            else
            {
                compiled.kind = CompiledThemeColorKind::SeededChoice;
                for (const std::wstring& reference : source.references)
                {
                    uint32_t candidate = 0u;
                    if (const HRESULT hr = ResolveReference(reference, candidate); FAILED(hr))
                        return hr;
                    compiled.candidates.push_back(candidate);
                }
            }
            output.dynamicColors[name] = std::move(compiled);
            out                        = output.dynamicColors[name].fallbackArgb;
            output.colors[name]        = out;
            resolved                   = true;
            return S_OK;
        }
        if (palette && Common::Settings::IsEventTimeThemeColorSource(source))
        {
            return Invalid(message, std::format(L"Palette entry '{}' cannot use an event-time source.", name));
        }

        if (const HRESULT hr = ResolveSource(source, out); FAILED(hr))
        {
            return hr;
        }
        if (palette)
        {
            paletteCache[name.substr(8u)]         = out;
            output.paletteColors[name.substr(8u)] = out;
        }
        else
            output.colors[name] = out;
        resolved = true;
        return S_OK;
    }

    [[nodiscard]] HRESULT ResolveSource(const ThemeColorSource& source, uint32_t& out)
    {
        auto ref = [&](size_t index, uint32_t& value) -> HRESULT
        {
            return index < source.references.size() ? ResolveReference(source.references[index], value)
                                                    : Invalid(message, L"Theme color source has an invalid reference arity.");
        };

        uint32_t first  = 0u;
        uint32_t second = 0u;
        uint32_t third  = 0u;
        switch (source.kind)
        {
            case ThemeColorSourceKind::Direct: out = source.directArgb; return S_OK;
            case ThemeColorSourceKind::Reference: return ref(0u, out);
            case ThemeColorSourceKind::Lighten:
                if (const HRESULT hr = ref(0u, first); FAILED(hr))
                    return hr;
                out = Blend(first, 0xFFFFFFFFu, source.parameters[0]);
                return S_OK;
            case ThemeColorSourceKind::Darken:
                if (const HRESULT hr = ref(0u, first); FAILED(hr))
                    return hr;
                out = Blend(first, 0xFF000000u, source.parameters[0]);
                return S_OK;
            case ThemeColorSourceKind::Alpha:
                if (const HRESULT hr = ref(0u, first); FAILED(hr))
                    return hr;
                out = (static_cast<uint32_t>(std::clamp(std::lround(source.parameters[0] * 255.0), 0l, 255l)) << 24u) | (first & 0x00FFFFFFu);
                return S_OK;
            case ThemeColorSourceKind::Blend:
                if (const HRESULT hr = ref(0u, first); FAILED(hr))
                    return hr;
                if (const HRESULT hr = ref(1u, second); FAILED(hr))
                    return hr;
                out = Blend(first, second, source.parameters[0]);
                return S_OK;
            case ThemeColorSourceKind::Contrast:
                if (const HRESULT hr = ref(0u, first); FAILED(hr))
                    return hr;
                if (source.references.size() == 1u)
                {
                    out = ContrastRatio(first, 0xFFFFFFFFu) >= ContrastRatio(first, 0xFF000000u) ? 0xFFFFFFFFu : 0xFF000000u;
                    return S_OK;
                }
                if (const HRESULT hr = ref(1u, second); FAILED(hr))
                    return hr;
                if (const HRESULT hr = ref(2u, third); FAILED(hr))
                    return hr;
                out = ContrastRatio(first, second) >= ContrastRatio(first, third) ? second : third;
                return S_OK;
            case ThemeColorSourceKind::PerceptualTone:
            {
                if (const HRESULT hr = ref(0u, first); FAILED(hr))
                    return hr;
                Oklch color     = ToOklch(first);
                color.lightness = source.parameters[0] / 100.0;
                out             = FromOklch(color);
                return S_OK;
            }
            case ThemeColorSourceKind::EnsureContrast:
                if (const HRESULT hr = ref(0u, first); FAILED(hr))
                    return hr;
                if (const HRESULT hr = ref(1u, second); FAILED(hr))
                    return hr;
                if (const std::optional<uint32_t> adjusted = EnsureContrast(first, second, source.parameters[0]); adjusted.has_value())
                {
                    out = adjusted.value();
                    return S_OK;
                }
                return Invalid(message, std::format(L"Contrast target {} is unattainable.", source.parameters[0]));
            case ThemeColorSourceKind::Harmonize:
            {
                if (const HRESULT hr = ref(0u, first); FAILED(hr))
                    return hr;
                if (const HRESULT hr = ref(1u, second); FAILED(hr))
                    return hr;
                Oklch color        = ToOklch(first);
                const Oklch target = ToOklch(second);
                double delta       = std::fmod(target.hue - color.hue + 540.0, 360.0) - 180.0;
                color.hue          = std::fmod(color.hue + delta * source.parameters[0] + 360.0, 360.0);
                out                = FromOklch(color);
                return S_OK;
            }
            case ThemeColorSourceKind::SystemColor: out = context.systemColors[static_cast<size_t>(source.systemRole)]; return S_OK;
            case ThemeColorSourceKind::Tone: return ref(context.effectiveDark ? 1u : 0u, out);
            case ThemeColorSourceKind::SeededRainbow:
            case ThemeColorSourceKind::SeededChoice: return Invalid(message, L"Paint-time source reached static resolution.");
        }
        return Invalid(message, L"Unsupported theme color source.");
    }
};
} // namespace

namespace Common::Settings
{
ThemeResolutionContext MakeSystemThemeResolutionContext(bool effectiveDark) noexcept
{
    const auto argbFromSystem = [](int index) noexcept
    {
        const COLORREF color = ::GetSysColor(index);
        return 0xFF000000u | (static_cast<uint32_t>(GetRValue(color)) << 16u) | (static_cast<uint32_t>(GetGValue(color)) << 8u) |
               static_cast<uint32_t>(GetBValue(color));
    };
    ThemeResolutionContext context;
    context.effectiveDark   = effectiveDark;
    DWORD colorizationColor = 0u;
    BOOL opaque             = FALSE;
    const uint32_t accent =
        SUCCEEDED(DwmGetColorizationColor(&colorizationColor, &opaque)) ? (0xFF000000u | (colorizationColor & 0x00FFFFFFu)) : argbFromSystem(COLOR_HIGHLIGHT);
    context.systemColors[static_cast<size_t>(ThemeSystemColorRole::Accent)]        = accent;
    context.systemColors[static_cast<size_t>(ThemeSystemColorRole::AccentLight)]   = Blend(accent, 0xFFFFFFFFu, 0.25);
    context.systemColors[static_cast<size_t>(ThemeSystemColorRole::AccentDark)]    = Blend(accent, 0xFF000000u, 0.25);
    context.systemColors[static_cast<size_t>(ThemeSystemColorRole::Window)]        = argbFromSystem(COLOR_WINDOW);
    context.systemColors[static_cast<size_t>(ThemeSystemColorRole::WindowText)]    = argbFromSystem(COLOR_WINDOWTEXT);
    context.systemColors[static_cast<size_t>(ThemeSystemColorRole::Highlight)]     = accent;
    context.systemColors[static_cast<size_t>(ThemeSystemColorRole::HighlightText)] = argbFromSystem(COLOR_HIGHLIGHTTEXT);
    return context;
}

HRESULT ParseThemeColorSource(std::wstring_view text, ThemeColorSource& outSource, std::wstring* outMessage) noexcept
{
    outSource = {};
    if (outMessage)
        outMessage->clear();
    text = Trim(text);
    if (text.empty() || text.size() > 256u)
    {
        return Invalid(outMessage, L"Theme color source is empty or exceeds 256 UTF-16 code units.");
    }

    uint32_t direct = 0u;
    if (TryParseColor(text, direct))
    {
        outSource.directArgb = direct;
        return S_OK;
    }

    const size_t open = text.find(L'(');
    if (open == std::wstring_view::npos || text.back() != L')' || text.find(L'(', open + 1u) != std::wstring_view::npos)
    {
        return Invalid(outMessage, std::format(L"Invalid theme color expression '{}'.", text));
    }
    const std::wstring function               = ToLower(Trim(text.substr(0u, open)));
    const std::wstring_view argumentsText     = Trim(text.substr(open + 1u, text.size() - open - 2u));
    const std::vector<std::wstring_view> args = argumentsText.empty() ? std::vector<std::wstring_view>{} : SplitArguments(argumentsText);
    if (! argumentsText.empty() && args.empty())
    {
        return Invalid(outMessage, std::format(L"Invalid arguments for '{}'.", function));
    }
    const auto addReference = [&](std::wstring_view reference) -> bool
    {
        if (! IsValidReference(reference))
            return false;
        outSource.references.emplace_back(reference);
        return true;
    };
    const auto requireAmount = [&](std::wstring_view value, size_t index) -> bool { return ParseAmount(value, outSource.parameters[index]); };

    if (function == L"ref" && args.size() == 1u && addReference(args[0]))
        outSource.kind = ThemeColorSourceKind::Reference;
    else if (function == L"lighten" && args.size() == 2u && addReference(args[0]) && requireAmount(args[1], 0u))
        outSource.kind = ThemeColorSourceKind::Lighten;
    else if (function == L"darken" && args.size() == 2u && addReference(args[0]) && requireAmount(args[1], 0u))
        outSource.kind = ThemeColorSourceKind::Darken;
    else if (function == L"alpha" && args.size() == 2u && addReference(args[0]) && requireAmount(args[1], 0u))
        outSource.kind = ThemeColorSourceKind::Alpha;
    else if (function == L"blend" && args.size() == 3u && addReference(args[0]) && addReference(args[1]) && requireAmount(args[2], 0u))
        outSource.kind = ThemeColorSourceKind::Blend;
    else if (function == L"contrast" && (args.size() == 1u || args.size() == 3u) && std::ranges::all_of(args, addReference))
        outSource.kind = ThemeColorSourceKind::Contrast;
    else if (function == L"perceptualtone" && args.size() == 2u && addReference(args[0]) && ParseDouble(args[1], outSource.parameters[0]) &&
             outSource.parameters[0] >= 0.0 && outSource.parameters[0] <= 100.0)
        outSource.kind = ThemeColorSourceKind::PerceptualTone;
    else if (function == L"ensurecontrast" && args.size() == 3u && addReference(args[0]) && addReference(args[1]) &&
             ParseDouble(args[2], outSource.parameters[0]) && outSource.parameters[0] >= 1.0 && outSource.parameters[0] <= 21.0)
        outSource.kind = ThemeColorSourceKind::EnsureContrast;
    else if (function == L"harmonize" && args.size() == 3u && addReference(args[0]) && addReference(args[1]) && requireAmount(args[2], 0u))
        outSource.kind = ThemeColorSourceKind::Harmonize;
    else if (function == L"systemaccent" && args.empty())
    {
        outSource.kind                         = ThemeColorSourceKind::SystemColor;
        outSource.systemRole                   = ThemeSystemColorRole::Accent;
        outSource.preserveSystemAccentSpelling = true;
    }
    else if (function == L"systemcolor" && args.size() == 1u)
    {
        const std::optional<ThemeSystemColorRole> role = ParseSystemRole(args[0]);
        if (! role.has_value())
            return Invalid(outMessage, std::format(L"Unknown system color role '{}'.", args[0]));
        outSource.kind       = ThemeColorSourceKind::SystemColor;
        outSource.systemRole = role.value();
    }
    else if (function == L"tone" && args.size() == 2u && addReference(args[0]) && addReference(args[1]))
        outSource.kind = ThemeColorSourceKind::Tone;
    else if (function == L"seededrainbow" && args.size() == 5u && ToLower(args[0]) == L"runtime.seed" && requireAmount(args[1], 0u) &&
             requireAmount(args[2], 1u) && requireAmount(args[3], 2u) && ParseDouble(args[4], outSource.parameters[3]) && outSource.parameters[3] >= 0.0 &&
             outSource.parameters[3] <= 360.0)
        outSource.kind = ThemeColorSourceKind::SeededRainbow;
    else if (function == L"seededchoice" && args.size() >= 3u && args.size() <= 9u && ToLower(args[0]) == L"runtime.seed")
    {
        outSource.kind = ThemeColorSourceKind::SeededChoice;
        for (const std::wstring_view reference : std::span<const std::wstring_view>(args).subspan(1u))
        {
            if (! addReference(reference))
                return Invalid(outMessage, std::format(L"Invalid seededChoice reference '{}'.", reference));
        }
    }
    else
    {
        return Invalid(outMessage, std::format(L"Unsupported or invalid theme color function '{}'.", function));
    }
    return S_OK;
}

std::wstring FormatThemeColorSource(const ThemeColorSource& source)
{
    const auto ref = [&](size_t index) -> std::wstring { return index < source.references.size() ? source.references[index] : L"invalid"; };
    switch (source.kind)
    {
        case ThemeColorSourceKind::Direct: return FormatColor(source.directArgb);
        case ThemeColorSourceKind::Reference: return std::format(L"ref({})", ref(0u));
        case ThemeColorSourceKind::Lighten: return std::format(L"lighten({},{})", ref(0u), NumberText(source.parameters[0]));
        case ThemeColorSourceKind::Darken: return std::format(L"darken({},{})", ref(0u), NumberText(source.parameters[0]));
        case ThemeColorSourceKind::Alpha: return std::format(L"alpha({},{})", ref(0u), NumberText(source.parameters[0]));
        case ThemeColorSourceKind::Blend: return std::format(L"blend({},{},{})", ref(0u), ref(1u), NumberText(source.parameters[0]));
        case ThemeColorSourceKind::Contrast:
            return source.references.size() == 1u ? std::format(L"contrast({})", ref(0u)) : std::format(L"contrast({},{},{})", ref(0u), ref(1u), ref(2u));
        case ThemeColorSourceKind::PerceptualTone: return std::format(L"perceptualTone({},{})", ref(0u), NumberText(source.parameters[0]));
        case ThemeColorSourceKind::EnsureContrast: return std::format(L"ensureContrast({},{},{})", ref(0u), ref(1u), NumberText(source.parameters[0]));
        case ThemeColorSourceKind::Harmonize: return std::format(L"harmonize({},{},{})", ref(0u), ref(1u), NumberText(source.parameters[0]));
        case ThemeColorSourceKind::SystemColor:
            return source.preserveSystemAccentSpelling && source.systemRole == ThemeSystemColorRole::Accent
                       ? L"systemAccent()"
                       : std::format(L"systemColor({})", SystemRoleText(source.systemRole));
        case ThemeColorSourceKind::Tone: return std::format(L"tone({},{})", ref(0u), ref(1u));
        case ThemeColorSourceKind::SeededRainbow:
            return std::format(L"seededRainbow(runtime.seed,{},{},{},{})",
                               NumberText(source.parameters[0]),
                               NumberText(source.parameters[1]),
                               NumberText(source.parameters[2]),
                               NumberText(source.parameters[3]));
        case ThemeColorSourceKind::SeededChoice:
        {
            std::wstring result = L"seededChoice(runtime.seed";
            for (const std::wstring& reference : source.references)
                result += std::format(L",{}", reference);
            result += L")";
            return result;
        }
    }
    return {};
}

bool IsPaintTimeThemeColorSource(const ThemeColorSource& source) noexcept
{
    return source.kind == ThemeColorSourceKind::SeededRainbow || source.kind == ThemeColorSourceKind::SeededChoice;
}

bool IsEventTimeThemeColorSource(const ThemeColorSource& source) noexcept
{
    return source.kind == ThemeColorSourceKind::SystemColor || source.kind == ThemeColorSourceKind::Tone;
}

ThemeColorEvaluationPhase GetThemeColorEvaluationPhase(const ThemeColorSource& source) noexcept
{
    if (IsPaintTimeThemeColorSource(source))
        return ThemeColorEvaluationPhase::Paint;
    if (IsEventTimeThemeColorSource(source))
        return ThemeColorEvaluationPhase::Event;
    return ThemeColorEvaluationPhase::Load;
}

ThemeDynamicContextKind GetDynamicThemeColorContext(std::wstring_view semanticKey) noexcept
{
    return semanticKey == L"folderView.itemBackgroundSelected" ? ThemeDynamicContextKind::StableItemHash32 : ThemeDynamicContextKind::None;
}

HRESULT ResolveThemeDefinition(const ThemeDefinition& theme,
                               const ThemeResolutionContext& context,
                               ResolvedThemeColors& outResolved,
                               std::wstring* outMessage) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    HRESULT result       = S_OK;
    const auto emitPerf  = wil::scope_exit([&]
    {
        size_t edgeCount = 0u;
        for (const auto& [_, dependencies] : outResolved.dependencies)
            edgeCount += dependencies.size();
        Debug::Perf::EmitDurationUs(L"theme.resolve_us",
                                    Debug::Perf::ElapsedUs(startedAt),
                                    static_cast<uint64_t>(outResolved.dependencies.size()),
                                    static_cast<uint64_t>(edgeCount),
                                    result);
        Debug::Perf::EmitValue(L"theme.resolve.node_count", static_cast<uint64_t>(outResolved.dependencies.size()), result);
        Debug::Perf::EmitValue(L"theme.resolve.edge_count", static_cast<uint64_t>(edgeCount), result);
    });
    outResolved          = {};
    if (outMessage)
        outMessage->clear();

    std::unordered_map<std::wstring, uint8_t> depthStates;
    std::unordered_map<std::wstring, size_t> depthCache;
    HRESULT depthHr         = S_OK;
    const auto measureDepth = [&](auto&& self, const std::wstring& name, const ThemeColorSource& source) -> size_t
    {
        if (FAILED(depthHr))
            return 0u;
        if (const auto cached = depthCache.find(name); cached != depthCache.end())
            return cached->second;
        if (depthStates[name] == 1u)
            return 0u; // The resolver reports the complete cycle path after this bounded pass.
        depthStates[name] = 1u;
        size_t depth      = 1u;
        for (const std::wstring& reference : source.references)
        {
            const ThemeColorSource* dependency = nullptr;
            std::wstring dependencyName        = reference;
            if (reference.rfind(L"palette.", 0u) == 0u)
            {
                const auto found = theme.palette.find(reference.substr(8u));
                if (found != theme.palette.end())
                    dependency = &found->second;
            }
            else
            {
                const auto found = theme.colors.find(reference);
                if (found != theme.colors.end())
                    dependency = &found->second;
            }
            if (dependency)
                depth = std::max(depth, self(self, dependencyName, *dependency) + 1u);
        }
        depthStates[name] = 2u;
        depthCache[name]  = depth;
        if (depth > 32u && SUCCEEDED(depthHr))
        {
            depthHr = Invalid(outMessage, std::format(L"Theme color dependency depth exceeds 32 while resolving '{}'.", name));
        }
        return depth;
    };
    for (const auto& [name, source] : theme.palette)
    {
        static_cast<void>(measureDepth(measureDepth, L"palette." + name, source));
        if (FAILED(depthHr))
        {
            result = depthHr;
            return result;
        }
    }
    for (const auto& [name, source] : theme.colors)
    {
        static_cast<void>(measureDepth(measureDepth, name, source));
        if (FAILED(depthHr))
        {
            result = depthHr;
            return result;
        }
    }

    Resolver resolver(theme, context, outResolved, outMessage);
    for (const auto& [name, source] : theme.palette)
    {
        uint32_t ignored = 0u;
        if (const HRESULT hr = resolver.ResolveNamed(L"palette." + name, source, true, ignored); FAILED(hr))
        {
            result = hr;
            return result;
        }
    }
    for (const auto& [key, source] : theme.colors)
    {
        uint32_t ignored = 0u;
        if (const HRESULT hr = resolver.ResolveNamed(key, source, false, ignored); FAILED(hr))
        {
            result = hr;
            return result;
        }
    }
    return result;
}

uint32_t EvaluateDynamicThemeColor(const CompiledThemeColor& program, const ThemeRuntimeContext& context) noexcept
{
    if (context.highContrast)
    {
        return program.fallbackArgb;
    }
    if (program.kind == CompiledThemeColorKind::SeededChoice)
    {
        return program.candidates.empty() ? program.fallbackArgb : program.candidates[context.seedHash32 % program.candidates.size()];
    }
    const double hue = static_cast<double>(context.seedHash32 % 360u) + program.parameters[3];
    return HsvToArgb(hue, program.parameters[0], program.parameters[1], program.parameters[2]);
}
} // namespace Common::Settings
