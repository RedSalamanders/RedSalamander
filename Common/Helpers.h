#pragma once

#include "StringConversion.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <concepts>
#include <condition_variable>
#include <cstdint>
#include <cstring>
#include <cwctype>
#include <exception>
#include <filesystem>
#include <format>
#include <limits>
#include <locale>
#include <mutex>
#include <span>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include "Win32CallbackHelpers.h"
#include <windows.h>

#include <bcrypt.h>
#include <evntrace.h>

#pragma comment(lib, "bcrypt.lib")

#pragma warning(push)
// WIL and TraceLogging: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted),
// C5027 (move assign deleted), C4820 (padding)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <TraceLoggingProvider.h>
#include <wil/resource.h>

// MSVC can emit C4625/C4626/C5026/C5027 for WIL move-only templates at their first instantiation site.
// Force the instantiations used by this header while warnings are disabled.
namespace WilWarningSilenceDetail
{
struct ForceWilTemplateInstantiations_Helpers
{
    wil::unique_hbrush brush;
    wil::unique_hdc_paint hdcPaint;
    wil::unique_hlocal_string localString;
    wil::unique_hmodule module;
    wil::unique_hrgn region;
};
} // namespace WilWarningSilenceDetail
#pragma warning(pop)

// DLL export/import macro for Common.dll. COMMON_EXPORTS is defined by the Common project.
#ifndef COMMON_API
#ifdef COMMON_EXPORTS
#define COMMON_API __declspec(dllexport)
#else
#define COMMON_API __declspec(dllimport)
#endif
#endif

#include <Helpers.h>
#include <LocalizationManager.h>

#ifdef min
#undef min
#endif
#ifdef max
#undef max
#endif

#pragma warning(push)
#pragma warning(disable : 4514) // unreferenced inline function has been removed

namespace Debug
{
// predefine a TraceLogging provider for use in other modules
template <typename... Args> inline void Error(std::wformat_string<Args...> format, Args&&... args) noexcept;
inline void Error(const std::wstring& message) noexcept;
template <typename... Args> inline DWORD ErrorWithLastError(std::wformat_string<Args...> format, Args&&... args) noexcept;
inline DWORD ErrorWithLastError(const std::wstring& message) noexcept;
}; // namespace Debug

namespace LocaleFormatting
{
inline void InvalidateFormatLocaleCache() noexcept;
inline const std::locale& GetFormatLocale() noexcept;
} // namespace LocaleFormatting

namespace Common::Colors
{
// Drops the packed ARGB alpha channel and converts the remaining RGB bytes to Win32 COLORREF order.
[[nodiscard]] inline COLORREF ColorRefFromArgb(uint32_t argb) noexcept
{
    return RGB(static_cast<BYTE>((argb >> 16u) & 0xFFu), static_cast<BYTE>((argb >> 8u) & 0xFFu), static_cast<BYTE>(argb & 0xFFu));
}

// Interpolates COLORREF channels with an explicit integer denominator. Invalid denominators preserve
// the base color, the overlay weight clamps to the denominator, and division intentionally truncates.
[[nodiscard]] inline COLORREF BlendColorRefWeightedTruncate(COLORREF base, COLORREF overlay, int overlayWeight, int denominator) noexcept
{
    if (denominator <= 0)
    {
        return base;
    }

    overlayWeight        = std::clamp(overlayWeight, 0, denominator);
    const int baseWeight = denominator - overlayWeight;
    const auto channel   = [baseWeight, overlayWeight, denominator](BYTE baseChannel, BYTE overlayChannel) noexcept
    { return static_cast<BYTE>((static_cast<int>(baseChannel) * baseWeight + static_cast<int>(overlayChannel) * overlayWeight) / denominator); };

    return RGB(channel(GetRValue(base), GetRValue(overlay)), channel(GetGValue(base), GetGValue(overlay)), channel(GetBValue(base), GetBValue(overlay)));
}

// Convenience form for the repeated 8-bit overlay-alpha policy.
[[nodiscard]] inline COLORREF BlendColorRefTruncate(COLORREF under, COLORREF over, uint8_t overlayAlpha) noexcept
{
    return BlendColorRefWeightedTruncate(under, over, overlayAlpha, 255);
}

// Version 1 hashes each UTF-16 code unit as one FNV-1a input value. This is intentionally different
// from byte-wise UTF-16 hashing used by persisted AppTheme identities.
[[nodiscard]] inline uint32_t StableVisualHash32Utf16V1(std::wstring_view text) noexcept
{
    uint32_t hash = 2166136261u;
    for (const wchar_t codeUnit : text)
    {
        hash ^= static_cast<uint32_t>(codeUnit);
        hash *= 16777619u;
    }
    return hash;
}

namespace Detail
{
[[nodiscard]] inline COLORREF ColorRefFromNormalizedHsv(float hueDegrees, float saturation, float value) noexcept
{
    const float chroma = value * saturation;
    const float x      = chroma * (1.0f - std::fabs(std::fmod(hueDegrees / 60.0f, 2.0f) - 1.0f));
    const float m      = value - chroma;

    float red   = 0.0f;
    float green = 0.0f;
    float blue  = 0.0f;
    if (hueDegrees < 60.0f)
    {
        red   = chroma;
        green = x;
    }
    else if (hueDegrees < 120.0f)
    {
        red   = x;
        green = chroma;
    }
    else if (hueDegrees < 180.0f)
    {
        green = chroma;
        blue  = x;
    }
    else if (hueDegrees < 240.0f)
    {
        green = x;
        blue  = chroma;
    }
    else if (hueDegrees < 300.0f)
    {
        red  = x;
        blue = chroma;
    }
    else
    {
        red  = chroma;
        blue = x;
    }

    const auto toByte = [](float channel) noexcept { return static_cast<BYTE>(std::lround(std::clamp(channel * 255.0f, 0.0f, 255.0f))); };
    return RGB(toByte(red + m), toByte(green + m), toByte(blue + m));
}
} // namespace Detail

// Matches the majority viewer policy: negative hue clamps to red, hue wraps above 360 degrees,
// and saturation/value clamp to [0, 1].
[[nodiscard]] inline COLORREF ColorRefFromHsvClampedNegativeHueToZero(float hueDegrees, float saturation, float value) noexcept
{
    const float normalizedHue = std::fmod(std::max(0.0f, hueDegrees), 360.0f);
    return Detail::ColorRefFromNormalizedHsv(normalizedHue, std::clamp(saturation, 0.0f, 1.0f), std::clamp(value, 0.0f, 1.0f));
}

// Matches ViewerWeb's browser-style hue policy: negative hue wraps around the wheel. Saturation and
// value remain unbounded until the final RGB channels are clamped, preserving its existing behavior.
[[nodiscard]] inline COLORREF ColorRefFromHsvWrappedHue(float hueDegrees, float saturation, float value) noexcept
{
    const float normalizedHue = std::fmod(std::fmod(hueDegrees, 360.0f) + 360.0f, 360.0f);
    return Detail::ColorRefFromNormalizedHsv(normalizedHue, saturation, value);
}

[[nodiscard]] inline double SrgbChannelToLinear(double channel) noexcept
{
    return channel <= 0.04045 ? channel / 12.92 : std::pow((channel + 0.055) / 1.055, 2.4);
}

[[nodiscard]] inline double RelativeLuminanceFromSrgb(double red, double green, double blue) noexcept
{
    return 0.2126 * SrgbChannelToLinear(red) + 0.7152 * SrgbChannelToLinear(green) + 0.0722 * SrgbChannelToLinear(blue);
}

[[nodiscard]] inline double RelativeLuminanceFromArgb(uint32_t argb) noexcept
{
    return RelativeLuminanceFromSrgb(
        static_cast<double>((argb >> 16u) & 0xFFu) / 255.0, static_cast<double>((argb >> 8u) & 0xFFu) / 255.0, static_cast<double>(argb & 0xFFu) / 255.0);
}

[[nodiscard]] inline uint32_t CompositeArgbOverOpaqueBackground(uint32_t foregroundArgb, uint32_t backgroundArgb) noexcept
{
    const uint32_t alpha        = (foregroundArgb >> 24u) & 0xFFu;
    const uint32_t inverseAlpha = 0xFFu - alpha;
    const auto compositeChannel = [&](uint32_t shift) noexcept
    {
        const uint32_t foreground = (foregroundArgb >> shift) & 0xFFu;
        const uint32_t background = (backgroundArgb >> shift) & 0xFFu;
        return ((foreground * alpha) + (background * inverseAlpha) + 0x7Fu) / 0xFFu;
    };
    return 0xFF000000u | (compositeChannel(16u) << 16u) | (compositeChannel(8u) << 8u) | compositeChannel(0u);
}

[[nodiscard]] inline double RelativeLuminanceFromColorRef(COLORREF color) noexcept
{
    return RelativeLuminanceFromSrgb(
        static_cast<double>(GetRValue(color)) / 255.0, static_cast<double>(GetGValue(color)) / 255.0, static_cast<double>(GetBValue(color)) / 255.0);
}

// Some product policies intentionally apply their threshold to encoded sRGB channels without linearizing.
// Keep that math named separately from WCAG relative luminance.
[[nodiscard]] inline double WeightedSrgbLuminanceWithoutLinearization(double red, double green, double blue) noexcept
{
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue;
}

[[nodiscard]] inline double ContrastRatioFromRelativeLuminance(double first, double second) noexcept
{
    return (std::max(first, second) + 0.05) / (std::min(first, second) + 0.05);
}
} // namespace Common::Colors

namespace Common::Crypto
{
[[nodiscard]] inline HRESULT GenerateRandomBytes(std::span<std::byte> bytes) noexcept
{
    if (bytes.empty())
    {
        return S_OK;
    }
    if (bytes.size() > static_cast<size_t>((std::numeric_limits<ULONG>::max)()))
    {
        return E_INVALIDARG;
    }

    const NTSTATUS status = BCryptGenRandom(nullptr, reinterpret_cast<PUCHAR>(bytes.data()), static_cast<ULONG>(bytes.size()), BCRYPT_USE_SYSTEM_PREFERRED_RNG);
    return BCRYPT_SUCCESS(status) ? S_OK : HRESULT_FROM_NT(status);
}

template <typename CharT> inline void AppendHexToken(std::basic_string<CharT>& value, std::span<const std::byte> bytes)
{
    constexpr std::basic_string_view<CharT> hex = []
    {
        if constexpr (std::same_as<CharT, char>)
        {
            return std::basic_string_view<CharT>("0123456789abcdef");
        }
        else
        {
            return std::basic_string_view<CharT>(L"0123456789abcdef");
        }
    }();

    value.reserve(value.size() + bytes.size() * 2u);
    for (const std::byte byte : bytes)
    {
        const unsigned int item = std::to_integer<unsigned int>(byte);
        value.push_back(hex[(item >> 4u) & 0x0Fu]);
        value.push_back(hex[item & 0x0Fu]);
    }
}
} // namespace Common::Crypto

namespace Common::Paths
{
// External programs must not be resolved through the process working directory or PATH.
// Accept explicit drive, UNC, and Win32 extended paths; reject rooted, drive-relative,
// device-namespace, relative, and bare executable names.
[[nodiscard]] inline bool IsExplicitAbsoluteExecutablePath(std::wstring_view path) noexcept
{
    const auto isSeparator   = [](const wchar_t ch) noexcept { return ch == L'\\' || ch == L'/'; };
    const auto isDriveLetter = [](const wchar_t ch) noexcept { return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z'); };

    if (path.size() >= 3u && isDriveLetter(path[0]) && path[1] == L':' && isSeparator(path[2]))
    {
        return path.size() > 3u;
    }

    if (path.size() < 5u || ! isSeparator(path[0]) || ! isSeparator(path[1]))
    {
        return false;
    }

    // Win32 device paths (for example, \\.\PhysicalDrive0) are not executable paths.
    if (path[2] == L'.' && isSeparator(path[3]))
    {
        return false;
    }

    size_t componentStart = 2u;
    if (path[2] == L'?' && isSeparator(path[3]))
    {
        if (path.size() >= 7u && isDriveLetter(path[4]) && path[5] == L':' && isSeparator(path[6]))
        {
            return path.size() > 7u;
        }

        const auto equalsAsciiNoCase = [](const wchar_t lhs, const wchar_t rhs) noexcept
        {
            const wchar_t folded = (lhs >= L'a' && lhs <= L'z') ? static_cast<wchar_t>(lhs - (L'a' - L'A')) : lhs;
            return folded == rhs;
        };
        if (path.size() <= 8u || ! equalsAsciiNoCase(path[4], L'U') || ! equalsAsciiNoCase(path[5], L'N') || ! equalsAsciiNoCase(path[6], L'C') ||
            ! isSeparator(path[7]))
        {
            return false;
        }
        componentStart = 8u;
    }

    const size_t serverEnd = path.find_first_of(L"\\/", componentStart);
    if (serverEnd == std::wstring_view::npos || serverEnd == componentStart)
    {
        return false;
    }
    const size_t shareStart = serverEnd + 1u;
    const size_t shareEnd   = path.find_first_of(L"\\/", shareStart);
    return shareEnd != std::wstring_view::npos && shareEnd > shareStart && shareEnd + 1u < path.size();
}

// Builds `<readable-prefix><marker><128-bit-hex><suffix>`, truncating only the readable prefix
// when a provider imposes a leaf-length cap. The marker/token/suffix identity is always preserved.
template <typename CharT>
[[nodiscard]] inline HRESULT BuildUniqueSiblingName(std::basic_string_view<CharT> readablePrefix,
                                                    std::basic_string_view<CharT> marker,
                                                    std::basic_string_view<CharT> suffix,
                                                    size_t maxLength,
                                                    std::basic_string<CharT>& nameOut) noexcept
{
    nameOut.clear();
    constexpr size_t kEntropyBytes = 16u;
    constexpr size_t kHexChars     = kEntropyBytes * 2u;
    if (marker.size() > maxLength || suffix.size() > maxLength - marker.size() || kHexChars > maxLength - marker.size() - suffix.size())
    {
        return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
    }

    std::array<std::byte, kEntropyBytes> entropy{};
    const HRESULT randomHr = Common::Crypto::GenerateRandomBytes(entropy);
    if (FAILED(randomHr))
    {
        return randomHr;
    }

    const size_t prefixLimit = maxLength - marker.size() - kHexChars - suffix.size();
    nameOut.reserve((std::min)(readablePrefix.size(), prefixLimit) + marker.size() + kHexChars + suffix.size());
    nameOut.append(readablePrefix.substr(0u, prefixLimit));
    nameOut.append(marker);
    Common::Crypto::AppendHexToken(nameOut, entropy);
    nameOut.append(suffix);
    return S_OK;
}
} // namespace Common::Paths

#if defined(_DEBUG)
namespace Common::DebugSelfTest
{
struct Check final
{
    std::wstring_view component;

    bool operator()(bool condition, const wchar_t* message, unsigned int& passed, unsigned int& failed) const noexcept
    {
        if (condition)
        {
            ++passed;
            return true;
        }

        ++failed;
        Debug::Error(L"{} debug selftest failed: {}", component, message);
        return false;
    }
};
} // namespace Common::DebugSelfTest
#endif

namespace OrdinalString
{
inline int Compare(std::wstring_view a, std::wstring_view b, bool ignoreCase) noexcept
{
    if (a.size() > static_cast<size_t>(std::numeric_limits<int>::max()) || b.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        const int fallback = a.compare(b);
        return (fallback < 0) ? -1 : ((fallback > 0) ? 1 : 0);
    }

    const int aLen   = static_cast<int>(a.size());
    const int bLen   = static_cast<int>(b.size());
    const int result = CompareStringOrdinal(a.data(), aLen, b.data(), bLen, ignoreCase ? TRUE : FALSE);

    if (result == CSTR_LESS_THAN)
    {
        return -1;
    }
    if (result == CSTR_GREATER_THAN)
    {
        return 1;
    }
    if (result == CSTR_EQUAL)
    {
        return 0;
    }

    const int fallback = a.compare(b);
    return (fallback < 0) ? -1 : ((fallback > 0) ? 1 : 0);
}

inline bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    return Compare(a, b, true) == 0;
}

[[nodiscard]] inline bool StartsWithNoCase(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (prefix.empty())
    {
        return true;
    }

    if (text.size() < prefix.size())
    {
        return false;
    }

    if (prefix.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    return EqualsNoCase(text.substr(0, prefix.size()), prefix);
}

inline bool EqualsNoCasePath(const std::filesystem::path& a, const std::filesystem::path& b) noexcept
{
    return EqualsNoCase(a.native(), b.native());
}

inline bool EqualsNoCasePath(const std::filesystem::path& a, std::wstring_view b) noexcept
{
    return EqualsNoCase(a.native(), b);
}

inline bool EqualsNoCasePath(std::wstring_view a, const std::filesystem::path& b) noexcept
{
    return EqualsNoCase(a, b.native());
}

inline bool LessNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    const int cmp = Compare(a, b, true);
    if (cmp != 0)
    {
        return cmp < 0;
    }

    const int caseCmp = Compare(a, b, false);
    if (caseCmp != 0)
    {
        return caseCmp < 0;
    }

    return false;
}

[[nodiscard]] inline std::wstring FoldCaseInvariant(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    std::wstring fallback(text);
    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        static_cast<void>(::CharLowerBuffW(fallback.data(), static_cast<DWORD>(fallback.size())));
        return fallback;
    }

    const int sourceLength   = static_cast<int>(text.size());
    const int requiredLength = ::LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, text.data(), sourceLength, nullptr, 0, nullptr, nullptr, 0);
    if (requiredLength <= 0)
    {
        static_cast<void>(::CharLowerBuffW(fallback.data(), static_cast<DWORD>(fallback.size())));
        return fallback;
    }

    std::wstring folded(static_cast<size_t>(requiredLength), L'\0');
    const int written = ::LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, text.data(), sourceLength, folded.data(), requiredLength, nullptr, nullptr, 0);
    if (written <= 0)
    {
        static_cast<void>(::CharLowerBuffW(fallback.data(), static_cast<DWORD>(fallback.size())));
        return fallback;
    }

    folded.resize(static_cast<size_t>(written));
    if (! folded.empty() && folded.back() == L'\0')
    {
        folded.pop_back();
    }

    return folded;
}

[[nodiscard]] inline bool EqualsFoldedInvariant(std::wstring_view left, std::wstring_view right) noexcept
{
    return FoldCaseInvariant(left) == FoldCaseInvariant(right);
}

[[nodiscard]] inline bool StartsWithFoldedInvariant(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (prefix.empty())
    {
        return true;
    }

    if (text.size() < prefix.size())
    {
        return false;
    }

    return EqualsFoldedInvariant(text.substr(0u, prefix.size()), prefix);
}

[[nodiscard]] inline bool StartsWithPreFoldedInvariant(std::wstring_view text, std::wstring_view foldedPrefix) noexcept
{
    if (foldedPrefix.empty())
    {
        return true;
    }

    if (text.size() < foldedPrefix.size())
    {
        return false;
    }

    return FoldCaseInvariant(text.substr(0u, foldedPrefix.size())) == foldedPrefix;
}

[[nodiscard]] inline bool FindContainsFoldedInvariant(std::wstring_view text, std::wstring_view query, size_t& outOffset) noexcept
{
    outOffset = std::wstring_view::npos;
    if (query.empty() || text.size() < query.size())
    {
        return false;
    }

    const std::wstring foldedQuery = FoldCaseInvariant(query);
    const size_t querySize         = query.size();
    const size_t lastStartPosition = text.size() - querySize;
    for (size_t startPosition = 0u; startPosition <= lastStartPosition; ++startPosition)
    {
        if (FoldCaseInvariant(text.substr(startPosition, querySize)) == foldedQuery)
        {
            outOffset = startPosition;
            return true;
        }
    }

    return false;
}
} // namespace OrdinalString

namespace EnvironmentVariables
{
[[nodiscard]] inline std::optional<std::wstring> Read(std::wstring_view name) noexcept
{
    if (name.empty())
    {
        return std::nullopt;
    }

    const std::wstring key(name);
    for (unsigned int attempt = 0u; attempt < 4u; ++attempt)
    {
        ::SetLastError(ERROR_SUCCESS);
        const DWORD required = ::GetEnvironmentVariableW(key.c_str(), nullptr, 0u);
        if (required == 0u)
        {
            return ::GetLastError() == ERROR_ENVVAR_NOT_FOUND ? std::nullopt : std::optional<std::wstring>{std::wstring{}};
        }

        std::wstring value(static_cast<size_t>(required), L'\0');
        ::SetLastError(ERROR_SUCCESS);
        const DWORD written = ::GetEnvironmentVariableW(key.c_str(), value.data(), required);
        if (written == 0u)
        {
            return ::GetLastError() == ERROR_ENVVAR_NOT_FOUND ? std::nullopt : std::optional<std::wstring>{std::wstring{}};
        }
        if (written < required)
        {
            value.resize(static_cast<size_t>(written));
            return value;
        }
    }
    return std::nullopt;
}

[[nodiscard]] inline bool IsTruthyFlagSet(const wchar_t* name) noexcept
{
    if (name == nullptr)
    {
        return false;
    }
    const std::optional<std::wstring> value = Read(name);
    if (! value.has_value() || value->empty())
    {
        return false;
    }
    return value.value()[0] == L'1' || value.value()[0] == L'y' || value.value()[0] == L'Y' || value.value()[0] == L't' || value.value()[0] == L'T';
}
} // namespace EnvironmentVariables

namespace StringUtils
{
[[nodiscard]] inline std::wstring_view TrimWhitespace(std::wstring_view text) noexcept
{
    size_t start = 0;
    size_t end   = text.size();

    while (start < end && std::iswspace(text[start]) != 0)
    {
        ++start;
    }

    while (end > start && std::iswspace(text[end - 1]) != 0)
    {
        --end;
    }

    return text.substr(start, end - start);
}

[[nodiscard]] inline std::wstring TrimWhitespaceCopy(std::wstring_view text)
{
    return std::wstring(TrimWhitespace(text));
}
} // namespace StringUtils

namespace Redaction
{
[[nodiscard]] inline bool EqualsNoCaseAscii(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        const wchar_t ca = a[i];
        const wchar_t cb = b[i];

        const wchar_t la = (ca >= L'A' && ca <= L'Z') ? static_cast<wchar_t>(ca - L'A' + L'a') : ca;
        const wchar_t lb = (cb >= L'A' && cb <= L'Z') ? static_cast<wchar_t>(cb - L'A' + L'a') : cb;
        if (la != lb)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] inline bool IsSensitiveQueryKey(std::wstring_view key) noexcept
{
    constexpr std::wstring_view kKeys[] = {
        L"password",
        L"pass",
        L"pwd",
        L"token",
        L"access_key",
        L"accesskey",
        L"secret",
        L"secret_key",
        L"apikey",
        L"api_key",
        L"signature",
        L"sig",
        // S3-style presigned URL keys.
        L"x-amz-credential",
        L"x-amz-security-token",
        L"x-amz-signature",
    };

    for (const std::wstring_view k : kKeys)
    {
        if (EqualsNoCaseAscii(key, k))
        {
            return true;
        }
    }

    return false;
}

inline void RedactUserInfoPassword(std::wstring& text) noexcept
{
    const size_t schemeSep = text.find(L"://");
    if (schemeSep == std::wstring::npos)
    {
        return;
    }

    const size_t authorityStart = schemeSep + 3u;
    if (authorityStart >= text.size())
    {
        return;
    }

    const size_t authorityEnd = text.find_first_of(L"/?#", authorityStart);
    const size_t end          = (authorityEnd == std::wstring::npos) ? text.size() : authorityEnd;
    if (end <= authorityStart)
    {
        return;
    }

    const size_t at = text.rfind(L'@', end - 1u);
    if (at == std::wstring::npos || at <= authorityStart)
    {
        return;
    }

    const size_t colon = text.find(L':', authorityStart);
    if (colon == std::wstring::npos || colon >= at)
    {
        return;
    }

    constexpr std::wstring_view kRedacted = L"<redacted>";
    text.replace(colon + 1u, at - (colon + 1u), kRedacted);
}

inline void RedactSensitiveQueryParams(std::wstring& text) noexcept
{
    const size_t q = text.find(L'?');
    if (q == std::wstring::npos || (q + 1u) >= text.size())
    {
        return;
    }

    const size_t frag = text.find(L'#', q + 1u);
    const size_t end  = (frag == std::wstring::npos) ? text.size() : frag;

    const std::wstring_view query = std::wstring_view(text).substr(q + 1u, end - (q + 1u));
    if (query.empty())
    {
        return;
    }

    std::wstring redacted;
    redacted.reserve(query.size());

    size_t pos = 0;
    while (pos < query.size())
    {
        const size_t delimPos = query.find_first_of(L"&;", pos);
        const size_t segEnd   = (delimPos == std::wstring_view::npos) ? query.size() : delimPos;
        const std::wstring_view seg(query.substr(pos, segEnd - pos));

        const size_t eq = seg.find(L'=');
        if (eq != std::wstring_view::npos)
        {
            const std::wstring_view key = seg.substr(0, eq);
            if (IsSensitiveQueryKey(key))
            {
                redacted.append(key);
                redacted.push_back(L'=');
                redacted.append(L"<redacted>");
            }
            else
            {
                redacted.append(seg);
            }
        }
        else
        {
            redacted.append(seg);
        }

        if (delimPos != std::wstring_view::npos)
        {
            redacted.push_back(query[delimPos]);
            pos = delimPos + 1u;
        }
        else
        {
            pos = query.size();
        }
    }

    std::wstring rebuilt;
    rebuilt.reserve(text.size() - query.size() + redacted.size());
    rebuilt.append(text, 0, q + 1u);
    rebuilt.append(redacted);
    rebuilt.append(text, end, std::wstring::npos);
    text.swap(rebuilt);
}

// Best-effort redaction for log/diagnostics strings (URLs, endpoints, etc.).
// Intended to strip passwords/userinfo and common secret-bearing query parameters.
[[nodiscard]] inline std::wstring ForLog(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    std::wstring out(text);
    RedactUserInfoPassword(out);
    RedactSensitiveQueryParams(out);
    return out;
}
} // namespace Redaction

namespace SecureWipe
{
// Best-effort wipe; reallocation may leave copies.
inline void SecureClear(std::string& text) noexcept
{
    if (! text.empty())
    {
        SecureZeroMemory(text.data(), text.size());
        text.clear();
    }
}

// Best-effort wipe; reallocation may leave copies.
inline void SecureClear(std::wstring& text) noexcept
{
    if (! text.empty())
    {
        SecureZeroMemory(text.data(), text.size() * sizeof(wchar_t));
        text.clear();
    }
}
} // namespace SecureWipe

namespace Win32Text
{
[[nodiscard]] inline std::wstring GetWindowTextString(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return {};
    }

    const int length = GetWindowTextLengthW(hwnd);
    if (length <= 0)
    {
        return {};
    }

    std::wstring text(static_cast<size_t>(length) + 1u, L'\0');
    const int copied = GetWindowTextW(hwnd, text.data(), length + 1);
    if (copied <= 0)
    {
        return {};
    }

    text.resize(static_cast<size_t>(copied));
    return text;
}

[[nodiscard]] inline std::wstring GetDlgItemTextString(HWND dlg, int controlId) noexcept
{
    return GetWindowTextString(dlg ? GetDlgItem(dlg, controlId) : nullptr);
}
} // namespace Win32Text

// Loads directly from the embedded module, bypassing localization satellites.
template <typename string_type> int LoadEmbeddedStringResource(_In_opt_ HINSTANCE hInstance, _In_ UINT uID, string_type& result) WI_NOEXCEPT
{
    const HINSTANCE instance = hInstance ? hInstance : GetModuleHandleW(nullptr);
    if (! instance)
    {
        result.clear();
        return 0;
    }

    // LoadStringW supports returning a pointer directly to the resource string when cchBufferMax == 0.
    // This avoids guessing the required buffer size and supports embedded NULs (e.g. file dialog filters).
    PCWSTR ptr       = nullptr;
    const int length = ::LoadStringW(instance, uID, reinterpret_cast<LPWSTR>(&ptr), 0);
    if (length <= 0 || ! ptr)
    {
        result.clear();
        return 0;
    }

    result.assign(ptr, static_cast<size_t>(length));
    return length;
}

inline std::wstring LoadEmbeddedStringResource(_In_opt_ HINSTANCE hInstance, _In_ UINT uID) WI_NOEXCEPT
{
    std::wstring result;
    LoadEmbeddedStringResource(hInstance, uID, result);
    return result;
}

// LoadString from resource ID
template <typename string_type, size_t stackBufferLength = 256>
int LoadStringResource(_In_opt_ HINSTANCE hInstance, _In_ UINT uID, string_type& result) WI_NOEXCEPT
{
    static_assert(stackBufferLength <= INT_MAX, "stackBufferLength must fit in int");
    const HINSTANCE instance = hInstance ? hInstance : GetModuleHandleW(nullptr);
#if defined(COMMON_EXPORTS) || defined(REDSAL_USE_COMMON_LOCALIZATION)
    return Localization::LoadString(instance, uID, result);
#else
    return LoadEmbeddedStringResource(instance, uID, result);
#endif
}

// Convenience overload returning std::wstring.
inline std::wstring LoadStringResource(_In_opt_ HINSTANCE hInstance, _In_ UINT uID) WI_NOEXCEPT
{
    std::wstring result;
    LoadStringResource(hInstance, uID, result);
    return result;
}

// Loads a resource string and formats it using std::format-style placeholders.
// Uses std::vformat since resource strings are runtime values (not compile-time format strings).
template <typename... Args> std::wstring FormatLoadedStringResource(_In_ UINT uID, std::wstring_view fmt, Args... args)
{
    if (fmt.empty())
    {
        return {};
    }

    try
    {
        // std::make_wformat_args requires non-const lvalue references; take args by value so we can safely pass lvalues.
        return std::vformat(LocaleFormatting::GetFormatLocale(), fmt, std::make_wformat_args(args...));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::format_error& e)
    {
        std::wstring detail;
        const std::string_view what(e.what());
        detail.reserve(what.size());
        for (const char ch : what)
        {
            const auto byte = static_cast<unsigned char>(ch);
            detail.push_back(byte < 0x80u ? static_cast<wchar_t>(byte) : L'?');
        }

        Debug::Error(L"FormatLoadedStringResource: invalid format string for IDS={} ({}); using fallback text.", static_cast<unsigned int>(uID), detail);
        return std::wstring(fmt);
    }
}

template <typename... Args> std::wstring FormatStringResource(_In_opt_ HINSTANCE hInstance, _In_ UINT uID, Args... args)
{
    std::wstring fmt;
    LoadStringResource(hInstance, uID, fmt);
    return FormatLoadedStringResource(uID, std::wstring_view(fmt), args...);
}

template <typename... Args> std::wstring FormatEmbeddedStringResource(_In_opt_ HINSTANCE hInstance, _In_ UINT uID, Args... args)
{
    std::wstring fmt;
    LoadEmbeddedStringResource(hInstance, uID, fmt);
    return FormatLoadedStringResource(uID, std::wstring_view(fmt), args...);
}

namespace LocaleFormatting
{
inline std::atomic_uint32_t g_formatLocaleGeneration{1u};
inline std::mutex g_cachedFormatLocaleMutex;
inline std::locale g_cachedFormatLocale        = std::locale::classic();
inline uint32_t g_cachedFormatLocaleGeneration = 0u;

inline thread_local uint32_t g_threadFormatLocaleGeneration = 0u;
inline thread_local std::locale g_threadFormatLocale        = std::locale::classic();

inline void InvalidateFormatLocaleCache() noexcept
{
    g_formatLocaleGeneration.fetch_add(1u, std::memory_order_acq_rel);
}

inline const std::locale& GetFormatLocale() noexcept
{
    const uint32_t currentGeneration = g_formatLocaleGeneration.load(std::memory_order_acquire);
    if (g_threadFormatLocaleGeneration == currentGeneration)
    {
        return g_threadFormatLocale;
    }

    std::scoped_lock lock(g_cachedFormatLocaleMutex);
    if (g_cachedFormatLocaleGeneration != currentGeneration)
    {
        g_cachedFormatLocale           = std::locale("");
        g_cachedFormatLocaleGeneration = currentGeneration;
    }

    g_threadFormatLocale           = g_cachedFormatLocale;
    g_threadFormatLocaleGeneration = currentGeneration;
    return g_threadFormatLocale;
}
} // namespace LocaleFormatting

// Formats byte sizes as "B/KB/MB/GB/TB" with compact significant digits:
// - 1-digit integer part: 2 decimals (e.g. 4.60 MB)
// - 2-digit integer part: 1 decimal (e.g. 12.3 MB)
// - 3+ digit integer part: no decimals (e.g. 156 GB)
inline std::wstring FormatBytesCompact(uint64_t bytes)
{
    static constexpr std::array<std::wstring_view, 5> suffixes = {
        std::wstring_view(L"B"),
        std::wstring_view(L"KB"),
        std::wstring_view(L"MB"),
        std::wstring_view(L"GB"),
        std::wstring_view(L"TB"),
    };

    double value       = static_cast<double>(bytes);
    size_t suffixIndex = 0;
    while (value >= 1024.0 && (suffixIndex + 1) < suffixes.size())
    {
        value /= 1024.0;
        suffixIndex += 1;
    }

    if (suffixIndex == 0)
    {
        return std::format(LocaleFormatting::GetFormatLocale(), L"{:L} {}", bytes, suffixes[suffixIndex]);
    }

    int decimals = 0;
    if (value < 10.0)
    {
        decimals = (value >= 9.995) ? 1 : 2;
    }
    else if (value < 100.0)
    {
        decimals = (value >= 99.95) ? 0 : 1;
    }
    else
    {
        decimals = 0;
    }

    return std::format(LocaleFormatting::GetFormatLocale(), L"{:.{}Lf} {}", value, decimals, suffixes[suffixIndex]);
}

inline int MessageBoxThemedImpl(_In_opt_ HWND owner, PCWSTR text, PCWSTR caption, _In_ UINT type, bool centerOnOwner);

inline int MessageBoxResource(_In_opt_ HWND owner, _In_opt_ HINSTANCE hInstance, _In_ UINT textId, _In_ UINT captionId, _In_ UINT type)
{
    const std::wstring text    = LoadStringResource(hInstance, textId);
    const std::wstring caption = LoadStringResource(hInstance, captionId);
    return MessageBoxThemedImpl(owner, text.c_str(), caption.c_str(), type, false);
}

// Thread-local storage for MessageBox centering hook
namespace MessageBoxCenteringDetail
{
inline thread_local HWND g_centerOnWindow      = nullptr;
inline thread_local HHOOK g_hook               = nullptr;
inline thread_local WNDPROC g_msgBoxWndProc    = nullptr;
inline thread_local bool g_themeEnabled        = false;
inline thread_local bool g_themeUseDarkMode    = false;
inline thread_local COLORREF g_themeBackground = RGB(255, 255, 255);
inline thread_local COLORREF g_themeText       = RGB(0, 0, 0);
inline thread_local wil::unique_hbrush g_themeBrush;

inline std::atomic_bool g_defaultThemeEnabled{false};
inline std::atomic_bool g_defaultThemeUseDarkMode{false};
inline std::atomic_bool g_defaultThemeHighContrast{false};
inline std::atomic<DWORD> g_defaultThemeBackground{RGB(255, 255, 255)};
inline std::atomic<DWORD> g_defaultThemeText{RGB(0, 0, 0)};

inline void ApplyImmersiveDarkMode(HWND hwnd, bool enabled) noexcept
{
    if (! hwnd)
    {
        return;
    }

    using DwmSetWindowAttributeFunc          = HRESULT(WINAPI*)(HWND, DWORD, LPCVOID, DWORD);
    static DwmSetWindowAttributeFunc setAttr = []() noexcept -> DwmSetWindowAttributeFunc
    {
        HMODULE dwm = LoadLibrary(L"dwmapi.dll");
        if (! dwm)
        {
            Debug::ErrorWithLastError(L"Failed to load dwmapi.dll for ApplyImmersiveDarkMode.");
            return nullptr;
        }
#pragma warning(push)
#pragma warning(disable : 4191) // C4191: 'reinterpret_cast': unsafe conversion from 'FARPROC'
        return reinterpret_cast<DwmSetWindowAttributeFunc>(GetProcAddress(dwm, "DwmSetWindowAttribute"));
#pragma warning(pop)
    }();

    if (! setAttr)
    {
        return;
    }

    static constexpr DWORD kDwmwaUseImmersiveDarkMode19 = 19u;
    static constexpr DWORD kDwmwaUseImmersiveDarkMode20 = 20u;

    const BOOL darkMode = enabled ? TRUE : FALSE;
    setAttr(hwnd, kDwmwaUseImmersiveDarkMode20, &darkMode, sizeof(darkMode));
    setAttr(hwnd, kDwmwaUseImmersiveDarkMode19, &darkMode, sizeof(darkMode));
}

inline void ApplyWindowTheme(HWND hwnd, bool darkMode) noexcept
{
    if (! hwnd)
    {
        return;
    }

    using SetWindowThemeFunc           = HRESULT(WINAPI*)(HWND, LPCWSTR, LPCWSTR);
    static SetWindowThemeFunc setTheme = []() noexcept -> SetWindowThemeFunc
    {
        HMODULE uxTheme = LoadLibrary(L"uxtheme.dll");
        if (! uxTheme)
        {
            Debug::ErrorWithLastError(L"Failed to load uxtheme.dll for ApplyWindowTheme.");
            return nullptr;
        }
#pragma warning(push)
#pragma warning(disable : 4191) // C4191: 'reinterpret_cast': unsafe conversion from 'FARPROC'
        return reinterpret_cast<SetWindowThemeFunc>(GetProcAddress(uxTheme, "SetWindowTheme"));
#pragma warning(pop)
    }();

    if (! setTheme)
    {
        return;
    }

    setTheme(hwnd, darkMode ? L"DarkMode_Explorer" : L"Explorer", nullptr);
}

inline LRESULT OnThemedMessageBoxPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
    if (hdc)
    {
        FillRect(hdc.get(), &ps.rcPaint, g_themeBrush.get());
    }
    return 0;
}

inline bool TryHandleThemedMessageBoxEraseBkgnd(HWND hwnd, HDC hdc) noexcept
{
    RECT client{};
    if (GetClientRect(hwnd, &client))
    {
        FillRect(hdc, &client, g_themeBrush.get());
        return true;
    }
    return false;
}

inline LRESULT OnThemedMessageBoxCtlColorText(HDC hdc) noexcept
{
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, g_themeText);
    return reinterpret_cast<LRESULT>(g_themeBrush.get());
}

inline LRESULT CALLBACK ThemedMessageBoxWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (g_themeEnabled && g_themeBrush)
    {
        switch (msg)
        {
            case WM_PAINT: return OnThemedMessageBoxPaint(hwnd);
            case WM_ERASEBKGND:
                if (TryHandleThemedMessageBoxEraseBkgnd(hwnd, reinterpret_cast<HDC>(wp)))
                {
                    return 1;
                }
                break;
            case WM_CTLCOLORDLG: return reinterpret_cast<LRESULT>(g_themeBrush.get());
            case WM_CTLCOLORSTATIC: return OnThemedMessageBoxCtlColorText(reinterpret_cast<HDC>(wp));
            case WM_CTLCOLORBTN: return OnThemedMessageBoxCtlColorText(reinterpret_cast<HDC>(wp));
        }
    }

    if (msg == WM_NCDESTROY && g_msgBoxWndProc)
    {
        WNDPROC original = g_msgBoxWndProc;
        g_msgBoxWndProc  = nullptr;
        static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(original)));
        const LRESULT result = RedSalamander::Win32Callback::CallWindowProcNoThrow(original, hwnd, msg, wp, lp);
        return result;
    }

    if (g_msgBoxWndProc)
    {
        const LRESULT result = RedSalamander::Win32Callback::CallWindowProcNoThrow(g_msgBoxWndProc, hwnd, msg, wp, lp);
        return result;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

inline BOOL CALLBACK ApplyThemeToChildWindowsProc(HWND hwnd, LPARAM /*lp*/) noexcept
{
    ApplyWindowTheme(hwnd, g_themeUseDarkMode);
    SendMessageW(hwnd, WM_THEMECHANGED, 0, 0);
    return TRUE;
}

inline LRESULT CALLBACK CenteringHookProc(int nCode, WPARAM wParam, LPARAM lParam) noexcept
{
    if (nCode == HCBT_ACTIVATE && (g_centerOnWindow || g_themeEnabled))
    {
        HWND msgBox = reinterpret_cast<HWND>(wParam);

        if (g_themeEnabled && g_themeBrush)
        {
            ApplyImmersiveDarkMode(msgBox, g_themeUseDarkMode);
            ApplyWindowTheme(msgBox, g_themeUseDarkMode);
            EnumChildWindows(msgBox, ApplyThemeToChildWindowsProc, 0);
            SendMessageW(msgBox, WM_THEMECHANGED, 0, 0);

            if (! g_msgBoxWndProc)
            {
                g_msgBoxWndProc = reinterpret_cast<WNDPROC>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(
                    msgBox, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(MessageBoxCenteringDetail::ThemedMessageBoxWndProc)));
            }
        }

        if (g_centerOnWindow)
        {
            HWND owner = g_centerOnWindow;

            RECT ownerRc{};
            RECT msgRc{};
            if (GetWindowRect(owner, &ownerRc) && GetWindowRect(msgBox, &msgRc))
            {
                const int ownerW = ownerRc.right - ownerRc.left;
                const int ownerH = ownerRc.bottom - ownerRc.top;
                const int msgW   = msgRc.right - msgRc.left;
                const int msgH   = msgRc.bottom - msgRc.top;

                const int x = ownerRc.left + (ownerW - msgW) / 2;
                const int y = ownerRc.top + (ownerH - msgH) / 2;

                SetWindowPos(msgBox, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
            }
        }

        // Unhook after first activation
        if (g_hook)
        {
            UnhookWindowsHookEx(g_hook);
            g_hook = nullptr;
        }
        g_centerOnWindow = nullptr;
    }

    return CallNextHookEx(g_hook, nCode, wParam, lParam);
}
} // namespace MessageBoxCenteringDetail

struct MessageBoxTheme
{
    bool enabled        = false;
    bool useDarkMode    = false;
    bool highContrast   = false;
    COLORREF background = RGB(255, 255, 255);
    COLORREF text       = RGB(0, 0, 0);
};

inline void SetDefaultMessageBoxTheme(const MessageBoxTheme& theme) noexcept
{
    MessageBoxCenteringDetail::g_defaultThemeEnabled.store(theme.enabled, std::memory_order_relaxed);
    MessageBoxCenteringDetail::g_defaultThemeUseDarkMode.store(theme.useDarkMode, std::memory_order_relaxed);
    MessageBoxCenteringDetail::g_defaultThemeHighContrast.store(theme.highContrast, std::memory_order_relaxed);
    MessageBoxCenteringDetail::g_defaultThemeBackground.store(static_cast<DWORD>(theme.background), std::memory_order_relaxed);
    MessageBoxCenteringDetail::g_defaultThemeText.store(static_cast<DWORD>(theme.text), std::memory_order_relaxed);
}

inline void ClearDefaultMessageBoxTheme() noexcept
{
    MessageBoxTheme theme{};
    SetDefaultMessageBoxTheme(theme);
}

inline int MessageBoxThemedImpl(_In_opt_ HWND owner, PCWSTR text, PCWSTR caption, _In_ UINT type, bool centerOnOwner)
{
    const bool themeEnabled = MessageBoxCenteringDetail::g_defaultThemeEnabled.load(std::memory_order_relaxed) &&
                              ! MessageBoxCenteringDetail::g_defaultThemeHighContrast.load(std::memory_order_relaxed);

    if (centerOnOwner && owner && IsWindow(owner))
    {
        MessageBoxCenteringDetail::g_centerOnWindow = owner;
    }

    if (themeEnabled)
    {
        MessageBoxCenteringDetail::g_themeEnabled     = true;
        MessageBoxCenteringDetail::g_themeUseDarkMode = MessageBoxCenteringDetail::g_defaultThemeUseDarkMode.load(std::memory_order_relaxed);
        MessageBoxCenteringDetail::g_themeBackground =
            static_cast<COLORREF>(MessageBoxCenteringDetail::g_defaultThemeBackground.load(std::memory_order_relaxed));
        MessageBoxCenteringDetail::g_themeText = static_cast<COLORREF>(MessageBoxCenteringDetail::g_defaultThemeText.load(std::memory_order_relaxed));
        MessageBoxCenteringDetail::g_themeBrush.reset(CreateSolidBrush(MessageBoxCenteringDetail::g_themeBackground));
    }

    if ((MessageBoxCenteringDetail::g_centerOnWindow || themeEnabled) && ! MessageBoxCenteringDetail::g_hook)
    {
        MessageBoxCenteringDetail::g_hook = SetWindowsHookExW(WH_CBT, MessageBoxCenteringDetail::CenteringHookProc, nullptr, GetCurrentThreadId());
    }

    const int result = MessageBoxW(owner, text, caption, type);

    if (MessageBoxCenteringDetail::g_hook)
    {
        UnhookWindowsHookEx(MessageBoxCenteringDetail::g_hook);
        MessageBoxCenteringDetail::g_hook = nullptr;
    }

    MessageBoxCenteringDetail::g_centerOnWindow   = nullptr;
    MessageBoxCenteringDetail::g_msgBoxWndProc    = nullptr;
    MessageBoxCenteringDetail::g_themeEnabled     = false;
    MessageBoxCenteringDetail::g_themeUseDarkMode = false;
    MessageBoxCenteringDetail::g_themeBrush.reset();

    return result;
}

// MessageBox that is centered on the owner window
inline int MessageBoxCentered(_In_ HWND owner, _In_opt_ HINSTANCE hInstance, _In_ UINT textId, _In_ UINT captionId, _In_ UINT type)
{
    const std::wstring text    = LoadStringResource(hInstance, textId);
    const std::wstring caption = LoadStringResource(hInstance, captionId);
    return MessageBoxThemedImpl(owner, text.c_str(), caption.c_str(), type, true);
}

// MessageBox with caller-provided text, centered on the owner window.
inline int MessageBoxCenteredText(_In_opt_ HWND owner, const std::wstring& text, const std::wstring& caption, _In_ UINT type)
{
    return MessageBoxThemedImpl(owner, text.c_str(), caption.c_str(), type, true);
}

// class name for the RedSalamander Monitor window
constexpr auto g_redSalamanderMonitor          = L"RedSalamander Monitor";
constexpr auto g_redSalamanderMonitorClassName = L"RedSalamanderMonitor Window";

//////////////////////////////////////////////////////////////////////////////////
// DEBUG helpers
namespace Debug
{
struct TransportStats
{
    uint64_t etwWritten = 0;
    uint64_t etwFailed  = 0;
};

struct InfoParam // the real data/string will be after this structure
{
    enum Type : uint32_t
    {
        Text    = 0x0,
        Error   = 0x1,
        Warning = 0x2,
        Info    = 0x4,
        Perf    = 0x8,
        Debug   = 0x10,
        All     = 0x3F // Bitmask for all visible message types (bits 0-5)
    };
    FILETIME time; // More efficient: 8 bytes vs 16 bytes for SYSTEMTIME
    DWORD processID;
    DWORD threadID;
    Type type; // 0 - text, 1 - error, 2 - warning, 3 - info, 4 - debug

    // Helper method to get SYSTEMTIME when needed for display
    SYSTEMTIME GetLocalTime() const noexcept
    {
        SYSTEMTIME st{};
        FILETIME localFileTime{};
        if (FileTimeToLocalFileTime(&time, &localFileTime))
        {
            FileTimeToSystemTime(&localFileTime, &st);
        }
        return st;
    }

    // Helper method to get formatted time string
    std::wstring GetTimeString() const noexcept
    {
        SYSTEMTIME st = GetLocalTime();
        return std::format(L"{:02d}:{:02d}:{:02d}.{:03d}", st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);
    }
};

inline constexpr std::array<InfoParam::Type, 6> kFilterableInfoTypes{{
    InfoParam::Type::Text,
    InfoParam::Type::Error,
    InfoParam::Type::Warning,
    InfoParam::Type::Info,
    InfoParam::Type::Perf,
    InfoParam::Type::Debug,
}};

[[nodiscard]] inline constexpr uint32_t FilterBitForType(InfoParam::Type type) noexcept
{
    switch (type)
    {
        case InfoParam::Type::Text: return 0x01u;
        case InfoParam::Type::Error: return 0x02u;
        case InfoParam::Type::Warning: return 0x04u;
        case InfoParam::Type::Info: return 0x08u;
        case InfoParam::Type::Perf: return 0x10u;
        case InfoParam::Type::Debug: return 0x20u;
        case InfoParam::Type::All:
        default: return InfoParam::Type::All;
    }
}

// TraceLogging provider declaration
// Each module (EXE/DLL) must define its own provider instance using the same GUID
// to avoid cross-module provider handle issues. TraceLogging does NOT support
// sharing provider handles across DLL boundaries.
//
// Usage:
//   - In ONE .cpp file per module: #define REDSAL_DEFINE_TRACE_PROVIDER before including Helpers.h
//   - In all other files: Just include Helpers.h (uses TRACELOGGING_DECLARE_PROVIDER)

#if ! defined(REDSAL_DEFINE_TRACE_PROVIDER)
// Declaration only - provider will be defined elsewhere in this module
TRACELOGGING_DECLARE_PROVIDER(g_RedSalamanderProvider);
#else
// Definition - creates the actual provider storage in this module
TRACELOGGING_DEFINE_PROVIDER(g_RedSalamanderProvider,
                             "RedSalamanderMonitor",
                             // {440c70f6-6c6b-4ff7-9a3f-0b7db411b31a}
                             (0x440c70f6, 0x6c6b, 0x4ff7, 0x9a, 0x3f, 0x0b, 0x7d, 0xb4, 0x11, 0xb3, 0x1a));
#endif

namespace detail
{
inline std::atomic<uint64_t> g_etwWritten{0};
inline std::atomic<uint64_t> g_etwFailed{0};
inline std::once_flag g_traceLoggingRegisterOnce;
inline std::atomic<bool> g_etwRegistered{false};

struct IndentationState
{
    int level = 0;
    std::wstring prefix;
};

inline thread_local IndentationState g_indentation{};

inline void UpdateIndentationPrefix() noexcept
{
    constexpr int kMaxIndentLevel       = 64;
    constexpr int kIndentSpacesPerLevel = 2;
    const int boundedLevel              = std::clamp(g_indentation.level, 0, kMaxIndentLevel);

    if (boundedLevel <= 0)
    {
        g_indentation.prefix.clear();
        return;
    }

    const size_t spaceCount = static_cast<size_t>(kIndentSpacesPerLevel) * static_cast<size_t>(boundedLevel) - 1u;
    g_indentation.prefix.assign(spaceCount, L' ');
    g_indentation.prefix.append(L" - ");
}

inline void Indent() noexcept
{
    if (g_indentation.level < std::numeric_limits<int>::max())
    {
        ++g_indentation.level;
    }
    UpdateIndentationPrefix();
}

inline void Unindent() noexcept
{
    if (g_indentation.level > 0)
    {
        --g_indentation.level;
    }
    UpdateIndentationPrefix();
}

inline std::wstring_view GetIndentationPrefix() noexcept
{
    return g_indentation.prefix;
}

inline void PrependIndentation(std::wstring& message) noexcept
{
    const std::wstring_view prefix = GetIndentationPrefix();
    if (prefix.empty())
    {
        return;
    }

    const size_t newlineCount = static_cast<size_t>(std::count(message.begin(), message.end(), L'\n'));
    if (newlineCount == 0)
    {
        message.insert(0, prefix);
        return;
    }

    std::wstring indented;
    indented.reserve(message.size() + ((newlineCount + 1u) * prefix.size()));
    indented.append(prefix);
    for (size_t i = 0; i < message.size(); ++i)
    {
        const wchar_t ch = message[i];
        indented.push_back(ch);
        if (ch == L'\n' && (i + 1u) < message.size())
        {
            indented.append(prefix);
        }
    }

    message.swap(indented);
}

inline bool EnsureTraceLoggingRegistered() noexcept
{
    std::call_once(g_traceLoggingRegisterOnce,
                   []() noexcept
    {
        const HRESULT hr   = TraceLoggingRegister(g_RedSalamanderProvider);
        const bool success = SUCCEEDED(hr);
        g_etwRegistered.store(success, std::memory_order_release);

#ifdef _DEBUG
        // Output detailed registration result for debugging
        wchar_t msg[256]{};
        const size_t msgMax = (sizeof(msg) / sizeof(msg[0])) - 1;
        if (success)
        {
            const auto r                      = std::format_to_n(msg, msgMax, L"ETW TraceLoggingRegister succeeded: 0x{:08X}\n", static_cast<unsigned>(hr));
            using SizeType                    = decltype(r.size);
            const SizeType cap                = static_cast<SizeType>(msgMax);
            const SizeType written            = (r.size < 0) ? 0 : ((r.size > cap) ? cap : r.size);
            msg[static_cast<size_t>(written)] = L'\0';
        }
        else
        {
            const wchar_t* const hrText = hr == static_cast<HRESULT>(0x80070005)   ? L"E_ACCESSDENIED"
                                          : hr == E_INVALIDARG                     ? L"E_INVALIDARG"
                                          : hr == static_cast<HRESULT>(0x800700B7) ? L"ERROR_ALREADY_EXISTS"
                                                                                   : L"Unknown Error";

            const auto r           = std::format_to_n(msg, msgMax, L"ETW TraceLoggingRegister FAILED: 0x{:08X} ({})\n", static_cast<unsigned>(hr), hrText);
            using SizeType         = decltype(r.size);
            const SizeType cap     = static_cast<SizeType>(msgMax);
            const SizeType written = (r.size < 0) ? 0 : ((r.size > cap) ? cap : r.size);
            msg[static_cast<size_t>(written)] = L'\0';
        }
        OutputDebugStringW(msg);
#endif
    });
    return g_etwRegistered.load(std::memory_order_acquire);
}

inline bool IsEtwRegistered() noexcept
{
    return g_etwRegistered.load(std::memory_order_acquire);
}

constexpr ULONGLONG kDebugKeyword = 0x0000000000000001ull;
constexpr ULONGLONG kPerfKeyword  = 0x0000000000000002ull;

[[nodiscard]] inline constexpr bool IsMonitorDiagnosticsBuild() noexcept
{
#if defined(RS_DIAGNOSTICS_RUNTIME_OPT_IN)
    return false;
#elif defined(_DEBUG) || defined(RS_ASAN_DEBUG_BUILD)
    return true;
#else
    return false;
#endif
}

inline std::atomic<bool> g_runtimeMonitorDiagnosticsEnabled{false};

inline void SetRuntimeMonitorDiagnosticsEnabled(bool enabled) noexcept
{
    g_runtimeMonitorDiagnosticsEnabled.store(enabled, std::memory_order_release);
}

[[nodiscard]] inline bool IsRuntimeMonitorDiagnosticsEnabled() noexcept
{
    if (g_runtimeMonitorDiagnosticsEnabled.load(std::memory_order_acquire))
    {
        return true;
    }

    wchar_t value[8]{};
    const DWORD chars = ::GetEnvironmentVariableW(L"REDSALAMANDER_DIAGNOSTICS_ETW", value, static_cast<DWORD>(_countof(value)));
    return chars > 0 && value[0] != L'\0' && value[0] != L'0';
}

[[nodiscard]] inline constexpr bool ShouldEmitMonitorDiagnosticMessageTypeByDefault(InfoParam::Type type) noexcept
{
    if (IsMonitorDiagnosticsBuild())
    {
        return true;
    }

    switch (type)
    {
        case InfoParam::Type::Error:
        case InfoParam::Type::Warning: return true;

        case InfoParam::Type::Text:
        case InfoParam::Type::Info:
        case InfoParam::Type::Perf:
        case InfoParam::Type::Debug:
        case InfoParam::Type::All:
        default: return false;
    }
}

[[nodiscard]] inline bool ShouldEmitMonitorDiagnosticMessageType(InfoParam::Type type) noexcept
{
    return ShouldEmitMonitorDiagnosticMessageTypeByDefault(type) || IsRuntimeMonitorDiagnosticsEnabled();
}

inline bool IsEtwEnabled(ULONGLONG keyword) noexcept
{
    if (! EnsureTraceLoggingRegistered())
    {
        return false;
    }

    return TraceLoggingProviderEnabled(g_RedSalamanderProvider, TRACE_LEVEL_INFORMATION, keyword) != 0;
}

inline bool IsDebugEtwEnabled() noexcept
{
    return IsEtwEnabled(kDebugKeyword);
}

inline InfoParam BuildInfoParam(InfoParam::Type type) noexcept
{
    InfoParam dbg{};
    GetSystemTimeAsFileTime(&dbg.time);
    dbg.processID = GetCurrentProcessId();
    dbg.threadID  = GetCurrentThreadId();
    dbg.type      = type;
    return dbg;
}

inline bool EmitEtwEvent(const InfoParam& info, std::wstring_view message) noexcept
{
    if (! EnsureTraceLoggingRegistered())
    {
        g_etwFailed.fetch_add(1, std::memory_order_relaxed);
        return false;
    }

    if (TraceLoggingProviderEnabled(g_RedSalamanderProvider, TRACE_LEVEL_INFORMATION, kDebugKeyword) == 0)
    {
        return false;
    }

    ULARGE_INTEGER fileTime{};
    fileTime.LowPart  = info.time.dwLowDateTime;
    fileTime.HighPart = info.time.dwHighDateTime;
    USHORT length     = static_cast<USHORT>(std::min<size_t>(message.size(), std::numeric_limits<USHORT>::max()));

    // TraceLoggingWrite is a macro that doesn't return a value we can easily capture.
    // Once registration succeeds, write failures are extremely rare (only if provider disabled).
    // We count successful writes; failures would show as missing events in consumer.
    TraceLoggingWrite(g_RedSalamanderProvider,
                      "DebugMessage",
                      TraceLoggingLevel(TRACE_LEVEL_INFORMATION),
                      TraceLoggingKeyword(kDebugKeyword),
                      TraceLoggingUInt32(static_cast<UINT32>(info.type), "Type"),
                      TraceLoggingUInt32(info.processID, "ProcessId"),
                      TraceLoggingUInt32(info.threadID, "ThreadId"),
                      TraceLoggingUInt64(fileTime.QuadPart, "FileTime"),
                      TraceLoggingCountedWideString(message.data(), length, "Message"));

    g_etwWritten.fetch_add(1, std::memory_order_relaxed);
    return true;
}

inline void Publish(const InfoParam& dbg, std::wstring_view payload) noexcept
{
    EmitEtwEvent(dbg, payload);
}

inline void PublishString(std::wstring_view payload) noexcept
{
    const InfoParam dbg = BuildInfoParam(InfoParam::Type::Text);
    EmitEtwEvent(dbg, payload);
}

// Perf JSONL sink -- implementation lives in Common.dll (PerfJsonl.cpp).
COMMON_API void WritePerfJsonl(std::wstring_view metric, std::wstring_view detail, uint64_t durationUs, uint64_t value0, uint64_t value1, HRESULT hr) noexcept;

COMMON_API bool HasPerfJsonlOutput() noexcept;
} // namespace detail

inline TransportStats GetTransportStats() noexcept
{
    TransportStats stats{};
    stats.etwWritten = detail::g_etwWritten.load(std::memory_order_relaxed);
    stats.etwFailed  = detail::g_etwFailed.load(std::memory_order_relaxed);
    return stats;
}

namespace Perf
{
inline bool IsEnabled() noexcept
{
    return (detail::IsMonitorDiagnosticsBuild() || detail::IsRuntimeMonitorDiagnosticsEnabled()) && detail::IsEtwEnabled(detail::kPerfKeyword);
}

inline bool HasJsonlOutput() noexcept
{
    return detail::HasPerfJsonlOutput();
}

inline bool IsCaptureEnabled() noexcept
{
    return IsEnabled() || HasJsonlOutput();
}

[[nodiscard]] inline uint64_t ElapsedUs(std::chrono::steady_clock::time_point startedAt) noexcept
{
    const auto elapsed = std::chrono::steady_clock::now() - startedAt;
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count());
}

inline void Emit(std::wstring_view name, std::wstring_view detail, uint64_t durationUs, uint64_t value0 = 0, uint64_t value1 = 0, HRESULT hr = S_OK) noexcept
{
    detail::WritePerfJsonl(name, detail, durationUs, value0, value1, hr);

    if (! IsEnabled())
    {
        return;
    }

    const USHORT nameLen   = static_cast<USHORT>(std::min<size_t>(name.size(), std::numeric_limits<USHORT>::max()));
    const USHORT detailLen = static_cast<USHORT>(std::min<size_t>(detail.size(), std::numeric_limits<USHORT>::max()));

    TraceLoggingWrite(g_RedSalamanderProvider,
                      "PerfScope",
                      TraceLoggingLevel(TRACE_LEVEL_INFORMATION),
                      TraceLoggingKeyword(detail::kPerfKeyword),
                      TraceLoggingCountedWideString(name.data() ? name.data() : L"", nameLen, "Name"),
                      TraceLoggingCountedWideString(detail.data() ? detail.data() : L"", detailLen, "Detail"),
                      TraceLoggingUInt64(durationUs, "DurationUs"),
                      TraceLoggingUInt64(value0, "Value0"),
                      TraceLoggingUInt64(value1, "Value1"),
                      TraceLoggingUInt32(static_cast<uint32_t>(hr), "Hr"));
}

inline void EmitCounter(std::wstring_view name, uint64_t value = 1, HRESULT hr = S_OK) noexcept
{
    Emit(name, L"counter", 0, value, 0, hr);
}

inline void EmitValue(std::wstring_view name, uint64_t value, HRESULT hr = S_OK) noexcept
{
    Emit(name, L"value", 0, value, 0, hr);
}

inline void EmitDurationUs(std::wstring_view name, uint64_t durationUs, uint64_t value0 = 0, uint64_t value1 = 0, HRESULT hr = S_OK) noexcept
{
    Emit(name, L"duration", durationUs, value0, value1, hr);
}

COMMON_API void ConfigureJsonlOutput(const std::filesystem::path& path,
                                     std::wstring_view scenario    = {},
                                     std::wstring_view build       = {},
                                     std::wstring_view branch      = {},
                                     std::wstring_view commit      = {},
                                     std::wstring_view machineHash = {},
                                     std::wstring_view runId       = {}) noexcept;

COMMON_API void ClearJsonlOutput() noexcept;

class Scope final
{
public:
    explicit Scope(std::wstring_view name) noexcept
        : _etwEnabled(IsEnabled()),
          _jsonlEnabled(HasJsonlOutput()),
          _start((_etwEnabled || _jsonlEnabled) ? std::chrono::steady_clock::now() : std::chrono::steady_clock::time_point{})
    {
        if (_etwEnabled || _jsonlEnabled)
        {
            _name.assign(name);
        }
    }

    Scope(const Scope&)            = delete;
    Scope& operator=(const Scope&) = delete;
    Scope(Scope&&)                 = delete;
    Scope& operator=(Scope&&)      = delete;

    ~Scope() noexcept
    {
        if (! _etwEnabled && ! _jsonlEnabled)
        {
            return;
        }

        const auto elapsed        = std::chrono::steady_clock::now() - _start;
        const uint64_t durationUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count());

        if (_jsonlEnabled)
        {
            detail::WritePerfJsonl(_name, _detail, durationUs, _value0, _value1, static_cast<HRESULT>(_hr));
        }

        if (! _etwEnabled)
        {
            return;
        }

        const USHORT nameLen     = static_cast<USHORT>(std::min<size_t>(_name.size(), std::numeric_limits<USHORT>::max()));
        const USHORT detailLen   = static_cast<USHORT>(std::min<size_t>(_detail.size(), std::numeric_limits<USHORT>::max()));
        const wchar_t* namePtr   = _name.data() ? _name.data() : L"";
        const wchar_t* detailPtr = _detail.data() ? _detail.data() : L"";

        TraceLoggingWrite(g_RedSalamanderProvider,
                          "PerfScope",
                          TraceLoggingLevel(TRACE_LEVEL_INFORMATION),
                          TraceLoggingKeyword(detail::kPerfKeyword),
                          TraceLoggingCountedWideString(namePtr, nameLen, "Name"),
                          TraceLoggingCountedWideString(detailPtr, detailLen, "Detail"),
                          TraceLoggingUInt64(durationUs, "DurationUs"),
                          TraceLoggingUInt64(_value0, "Value0"),
                          TraceLoggingUInt64(_value1, "Value1"),
                          TraceLoggingUInt32(_hr, "Hr"));
    }

    void SetDetail(std::wstring_view detail) noexcept
    {
        if (_etwEnabled || _jsonlEnabled)
        {
            _detail.assign(detail);
        }
    }

    void SetValue0(uint64_t value) noexcept
    {
        _value0 = value;
    }

    void SetValue1(uint64_t value) noexcept
    {
        _value1 = value;
    }

    void SetHr(HRESULT hr) noexcept
    {
        _hr = static_cast<uint32_t>(hr);
    }

private:
    bool _etwEnabled   = false;
    bool _jsonlEnabled = false;
    std::wstring _name;
    std::wstring _detail;
    std::chrono::steady_clock::time_point _start;
    uint64_t _value0 = 0;
    uint64_t _value1 = 0;
    uint32_t _hr     = static_cast<uint32_t>(S_OK);
};
} // namespace Perf

inline void Out(PCWSTR p) noexcept
{
    if (! p)
    {
        return;
    }

    if (! detail::ShouldEmitMonitorDiagnosticMessageType(InfoParam::Type::Text))
    {
        return;
    }

    if (! detail::IsDebugEtwEnabled())
    {
        return;
    }

    const std::wstring_view prefix = detail::GetIndentationPrefix();
    if (prefix.empty())
    {
        detail::PublishString(p);
        return;
    }

    std::wstring message{p};
    detail::PrependIndentation(message);
    detail::PublishString(message);
}

template <typename... Args> inline void Out(InfoParam::Type type, std::wformat_string<Args...> format, Args&&... args) noexcept
{
    if (! detail::ShouldEmitMonitorDiagnosticMessageType(type))
    {
        return;
    }

    if (! detail::IsDebugEtwEnabled())
    {
        return;
    }

    // Mandatory: noexcept boundary. Formatting can throw; keep debug output best-effort.
    try
    {
        std::wstring formattedString = std::vformat(format.get(), std::make_wformat_args(args...));
        const InfoParam dbg          = detail::BuildInfoParam(type);

        detail::PrependIndentation(formattedString);

#ifdef _DEBUG
        if (type & InfoParam::Type::Error)
        {
            OutputDebugStringW(formattedString.c_str());
        }
#endif

        detail::Publish(dbg, formattedString);
    }
    catch (const std::bad_alloc&)
    {
        // Out-of-memory is treated as fatal. Fail-fast so the crash pipeline can capture a dump.
        std::terminate();
    }
    catch (const std::format_error&)
    {
        // Fallback for format string / argument mismatches.
        Debug::Out(L"[Formatting Error in DbgOut]");
    }
    catch (const std::exception&)
    {
        // Fallback for unexpected failures inside debug output.
        Debug::Out(L"[Unexpected Error in DbgOut]");
    }
}

// returns the last error code
template <typename... Args> inline DWORD LastError(InfoParam::Type type, std::wformat_string<Args...> format, Args&&... args) noexcept
{
    const DWORD lastError = ::GetLastError();

    if (! detail::ShouldEmitMonitorDiagnosticMessageType(type))
    {
        return lastError;
    }

    if (! detail::IsDebugEtwEnabled())
    {
        return lastError;
    }

    // Mandatory: noexcept boundary. Formatting can throw; keep debug output best-effort.
    try
    {
        std::wstring formattedString = std::vformat(format.get(), std::make_wformat_args(args...));
        if (lastError == 0)
        {
            formattedString.append(L" --> (NO ERROR)");
            Debug::Out(type, L"{}", formattedString);
            return 0;
        }

        wil::unique_hlocal_string message;
        const DWORD result = ::FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                                             nullptr,
                                             lastError,
                                             MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                                             reinterpret_cast<PWSTR>(message.addressof()),
                                             0,
                                             nullptr);

        if (result > 0 && message)
        {
            // Remove trailing newlines from system message
            std::wstring_view messageView{message.get()};
            while (! messageView.empty() && (messageView.back() == L'\r' || messageView.back() == L'\n'))
            {
                messageView.remove_suffix(1);
            }

            formattedString.append(std::format(L" --> ({}) {}", lastError, messageView));
        }
        else
        {
            formattedString.append(std::format(L" --> ({}) Unknown error", lastError));
        }

        Debug::Out(type, L"{}", formattedString);
        return lastError;
    }
    catch (const std::bad_alloc&)
    {
        // Out-of-memory is treated as fatal. Fail-fast so the crash pipeline can capture a dump.
        std::terminate();
    }
    catch (const std::format_error&)
    {
        // Best-effort: avoid throwing from this noexcept boundary.
        Debug::Out(type, L"[Formatting Error in Debug::OutLastError] LastError: {}", lastError);
        return lastError;
    }
    catch (const std::exception&)
    {
        // Best-effort: avoid throwing from this noexcept boundary.
        Debug::Out(type, L"[Unexpected Error in Debug::OutLastError] LastError: {}", lastError);
        return lastError;
    }
}

// Additional utility functions for common debug scenarios

template <typename... Args> inline void Info(std::wformat_string<Args...> format, Args&&... args) noexcept
{
    Debug::Out(InfoParam::Type::Info, format, std::forward<Args>(args)...);
}
inline void Info(const std::wstring& message) noexcept
{
    Debug::Out(InfoParam::Type::Info, L"{}", message);
}

template <typename... Args> inline void Warning(std::wformat_string<Args...> format, Args&&... args) noexcept
{
    Debug::Out(InfoParam::Type::Warning, format, std::forward<Args>(args)...);
}
inline void Warning(const std::wstring& message) noexcept
{
    Debug::Out(InfoParam::Type::Warning, L"{}", message);
}

template <typename... Args> inline void Error(std::wformat_string<Args...> format, Args&&... args) noexcept
{
    Debug::Out(InfoParam::Type::Error, format, std::forward<Args>(args)...);
}
inline void Error(const std::wstring& message) noexcept
{
    Debug::Out(InfoParam::Type::Error, L"{}", message);
}

template <typename... Args> inline DWORD ErrorWithLastError(std::wformat_string<Args...> format, Args&&... args) noexcept
{
    return Debug::LastError(InfoParam::Type::Error, format, std::forward<Args>(args)...);
}
inline DWORD ErrorWithLastError(const std::wstring& message) noexcept
{
    return Debug::LastError(InfoParam::Type::Error, L"{}", message);
}

} // namespace Debug

// ============================================================================
// HRESULT / system-message helpers
// ============================================================================
[[nodiscard]] inline std::wstring TrimTrailingSystemMessageWhitespace(std::wstring text) noexcept
{
    while (! text.empty())
    {
        const wchar_t ch = text.back();
        if (ch != L'\r' && ch != L'\n' && ch != L' ' && ch != L'\t')
        {
            break;
        }
        text.pop_back();
    }

    return text;
}

[[nodiscard]] inline DWORD GetSystemMessageIdFromHResult(HRESULT hr) noexcept
{
    DWORD messageId = static_cast<DWORD>(hr);
    if (HRESULT_FACILITY(hr) == FACILITY_WIN32)
    {
        const DWORD code = HRESULT_CODE(messageId);
        if (code != 0)
        {
            messageId = code;
        }
    }

    return messageId;
}

[[nodiscard]] inline std::wstring TryFormatSystemMessage(DWORD messageId) noexcept
{
    wil::unique_hlocal_string message;
    const DWORD written = ::FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS,
                                           nullptr,
                                           messageId,
                                           MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                                           reinterpret_cast<LPWSTR>(message.addressof()),
                                           0,
                                           nullptr);
    if (written == 0 || message.get() == nullptr)
    {
        return {};
    }

    return TrimTrailingSystemMessageWhitespace(std::wstring(message.get(), written));
}

[[nodiscard]] inline std::wstring FormatHResultMessage(HRESULT hr) noexcept
{
    std::wstring text = TryFormatSystemMessage(GetSystemMessageIdFromHResult(hr));
    if (! text.empty())
    {
        return text;
    }

    return std::format(L"HRESULT 0x{:08X}", static_cast<unsigned long>(hr));
}

[[nodiscard]] inline std::wstring FormatHResultMessageWithCode(HRESULT hr) noexcept
{
    std::wstring text = TryFormatSystemMessage(GetSystemMessageIdFromHResult(hr));
    if (! text.empty())
    {
        return std::format(L"0x{:08X}: {}", static_cast<unsigned long>(hr), text);
    }

    return std::format(L"0x{:08X}", static_cast<unsigned long>(hr));
}

template <typename TCallback> class RegistrationCallbackState
{
public:
    RegistrationCallbackState()                                            = default;
    RegistrationCallbackState(const RegistrationCallbackState&)            = delete;
    RegistrationCallbackState(RegistrationCallbackState&&)                 = delete;
    RegistrationCallbackState& operator=(const RegistrationCallbackState&) = delete;
    RegistrationCallbackState& operator=(RegistrationCallbackState&&)      = delete;
    ~RegistrationCallbackState()                                           = default;

    struct Snapshot
    {
        TCallback* callback = nullptr;
        void* cookie        = nullptr;
        uint64_t generation = 0;
    };

    void Set(TCallback* callback, void* cookie) noexcept
    {
        std::unique_lock lock(_mutex);
        ++_generation;
        _callback = callback;
        _cookie   = callback ? cookie : nullptr;
        if (! callback)
        {
            _drainCv.wait(lock, [this]() noexcept { return _inFlight == 0; });
        }
    }

    [[nodiscard]] bool TryCapture(Snapshot& snapshot) noexcept
    {
        std::scoped_lock lock(_mutex);
        if (! _callback)
        {
            snapshot = {};
            return false;
        }

        snapshot.callback   = _callback;
        snapshot.cookie     = _cookie;
        snapshot.generation = _generation;
        return true;
    }

    [[nodiscard]] bool TryEnter(const Snapshot& snapshot, TCallback*& callback, void*& cookie) noexcept
    {
        std::unique_lock lock(_mutex);
        if (_generation != snapshot.generation || _callback != snapshot.callback)
        {
            callback = nullptr;
            cookie   = nullptr;
            return false;
        }

        ++_inFlight;
        callback = snapshot.callback;
        cookie   = snapshot.cookie;
        return true;
    }

    void FinishInvoke() noexcept
    {
        {
            std::scoped_lock lock(_mutex);
            if (_inFlight > 0)
            {
                --_inFlight;
            }
        }

        _drainCv.notify_all();
    }

private:
    std::mutex _mutex;
    std::condition_variable _drainCv;
    TCallback* _callback = nullptr;
    void* _cookie        = nullptr;
    uint64_t _generation = 0;
    size_t _inFlight     = 0;
};

template <typename T> [[nodiscard]] inline bool TrySubmitUniqueToThreadpool(std::unique_ptr<T>& payload) noexcept
{
    if (! payload)
    {
        return true;
    }

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept { std::unique_ptr<T> owned(static_cast<T*>(context)); }, payload.get(), nullptr);
    if (queued == 0)
    {
        return false;
    }

    static_cast<void>(payload.release());
    return true;
}

template <typename T> [[nodiscard]] inline bool SubmitOwnedThreadpoolCallback(std::unique_ptr<T>& payload) noexcept
{
    if (! payload)
    {
        return true;
    }

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
    {
        std::unique_ptr<T> owned(static_cast<T*>(context));
        if (owned)
        {
            owned->Execute();
        }
    },
        payload.get(),
        nullptr);
    if (queued == 0)
    {
        return false;
    }

    static_cast<void>(payload.release());
    return true;
}

template <typename T> [[nodiscard]] inline bool SubmitOwnedThreadpoolCallbackWithInstance(std::unique_ptr<T>& payload) noexcept
{
    if (! payload)
    {
        return true;
    }

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE instance, void* context) noexcept
    {
        std::unique_ptr<T> owned(static_cast<T*>(context));
        if (owned)
        {
            owned->Execute(instance);
        }
    },
        payload.get(),
        nullptr);
    if (queued == 0)
    {
        return false;
    }

    static_cast<void>(payload.release());
    return true;
}

// ============================================================================
// Module Lifetime Helpers
// ============================================================================
/// Returns an owning module handle for the module that contains `address`.
/// This increments the module reference count so the module cannot be unloaded while the returned handle is alive.
/// Returns an empty handle on failure.
/// A threadpool callback must not merely destroy the last returned handle at the
/// end of its callback body: that can unmap the callback's own code before the
/// callback returns. Transfer that handle with
/// `TransferModulePinToCallbackReturn(instance, handle)` at callback entry.
[[nodiscard]] inline wil::unique_hmodule AcquireModuleReferenceFromAddress(const void* address) noexcept
{
    if (! address)
    {
        return {};
    }

    HMODULE module = nullptr;
    const BOOL ok  = GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, reinterpret_cast<LPCWSTR>(address), &module);
    if (ok == 0 || module == nullptr)
    {
        return {};
    }

    wil::unique_hmodule owned;
    owned.reset(module);
    return owned;
}

/// Transfers a pinned module reference to the Windows threadpool so it is
/// released only after the callback has returned to system code.
inline void TransferModulePinToCallbackReturn(PTP_CALLBACK_INSTANCE instance, wil::unique_hmodule& modulePin) noexcept
{
    if (instance && modulePin)
    {
        FreeLibraryWhenCallbackReturns(instance, modulePin.release());
    }
}

// ============================================================================
// PostMessage Payload RAII Helpers
// ============================================================================
// These helpers provide safe ownership transfer for payloads sent via PostMessageW/SendMessageW.
// They eliminate raw new/delete by using std::unique_ptr for automatic cleanup.
//
// Usage pattern:
//   Sender:
//     auto payload = std::make_unique<MyPayload>();
//     // ... fill payload ...
//     // Capture any payload fields needed by the other arguments before moving it;
//     // function arguments may evaluate the move before a sibling dereference.
//     if (!PostMessagePayload(hwnd, WM_MYMSG, 0, std::move(payload))) { /* handle error */ }
//
//   Receiver (WndProc):
//     auto payload = TakeMessagePayload<MyPayload>(lParam);
//     // NOTE: A nonzero LPARAM is an opaque registry token, never a payload pointer. Receiver MUST use
//     // TakeMessagePayload<T>, check for null, and ignore stale/drained/type-mismatched nonzero tokens.
//     // LPARAM zero is reserved for an explicitly posted payload-less fallback when that message defines one.
//     // ... use payload ...
//     // payload automatically deleted when scope exits
//
// Window teardown:
// - If an `HWND` is destroyed while messages are still queued, Windows may discard those messages without delivering them.
//   If those messages carry heap payload pointers, the payloads become unreachable (leak).
// - To prevent that, windows that receive payload messages should call `DrainPostedPayloadsForWindow(hwnd)` in `WM_NCDESTROY`.
//   It closes the target, invalidates every registered token, and then deletes registered payloads. Any stale queued token is harmless:
//   TakeMessagePayload<T> returns null and never adopts it as storage.
// - Call `InitPostedPayloadWindow(hwnd)` during create (`WM_NCCREATE`/`WM_CREATE`) to handle potential HWND reuse.

namespace detail
{
using MessagePayloadDeleter = void (*)(void*) noexcept;

struct PostedMessagePayloadEntry final
{
    void* payload             = nullptr;
    HWND hwnd                 = nullptr;
    UINT msg                  = 0;
    MessagePayloadDeleter del = nullptr;
};

struct PostedMessagePayloadRegistry final
{
    PostedMessagePayloadRegistry()                                               = default;
    PostedMessagePayloadRegistry(const PostedMessagePayloadRegistry&)            = delete;
    PostedMessagePayloadRegistry& operator=(const PostedMessagePayloadRegistry&) = delete;
    PostedMessagePayloadRegistry(PostedMessagePayloadRegistry&&)                 = delete;
    PostedMessagePayloadRegistry& operator=(PostedMessagePayloadRegistry&&)      = delete;

    std::mutex mutex;
    std::unordered_map<LPARAM, PostedMessagePayloadEntry> entriesByToken;
    std::unordered_map<HWND, std::unordered_set<LPARAM>> tokensByHwnd;
    std::unordered_set<HWND> closedHwnds;
    LPARAM nextToken = 1;
};

[[nodiscard]] inline PostedMessagePayloadRegistry& GetPostedMessagePayloadRegistry() noexcept
{
    // Intentionally leaked to avoid shutdown UAF from static destruction order issues.
    static PostedMessagePayloadRegistry* registry = new PostedMessagePayloadRegistry();
    return *registry;
}

inline void InitPostedPayloadWindow(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    auto& registry = GetPostedMessagePayloadRegistry();
    std::lock_guard lock(registry.mutex);
    static_cast<void>(registry.closedHwnds.erase(hwnd));
}

template <typename T> inline void DeletePostedMessagePayload(void* payload) noexcept
{
    delete static_cast<T*>(payload);
}

[[nodiscard]] inline LPARAM AllocatePostedMessagePayloadTokenLocked(PostedMessagePayloadRegistry& registry) noexcept
{
    const size_t maxAttempts = registry.entriesByToken.size() + 1u;
    for (size_t attempt = 0u; attempt < maxAttempts; ++attempt)
    {
        const LPARAM token = registry.nextToken;
        registry.nextToken = token >= std::numeric_limits<LPARAM>::max() ? 1 : token + 1;
        if (token != 0 && ! registry.entriesByToken.contains(token))
        {
            return token;
        }
    }
    return 0;
}

[[nodiscard]] inline bool ErasePostedMessagePayloadEntryLocked(PostedMessagePayloadRegistry& registry,
                                                               LPARAM token,
                                                               HWND expectedHwnd,
                                                               UINT expectedMessage,
                                                               bool validateDestination,
                                                               PostedMessagePayloadEntry& removed) noexcept
{
    const auto it = registry.entriesByToken.find(token);
    if (it == registry.entriesByToken.end())
    {
        return false;
    }
    if (validateDestination && (it->second.hwnd != expectedHwnd || it->second.msg != expectedMessage))
    {
        return false;
    }

    removed         = it->second;
    const HWND hwnd = removed.hwnd;
    registry.entriesByToken.erase(it);

    const auto hwIt = registry.tokensByHwnd.find(hwnd);
    if (hwIt != registry.tokensByHwnd.end())
    {
        hwIt->second.erase(token);
        if (hwIt->second.empty())
        {
            registry.tokensByHwnd.erase(hwIt);
        }
    }
    return true;
}

[[nodiscard]] inline bool TakeRegisteredPostedMessagePayload(
    LPARAM token, HWND expectedHwnd, UINT expectedMessage, bool validateDestination, PostedMessagePayloadEntry& removed) noexcept
{
    if (token == 0)
    {
        return false;
    }

    auto& registry = GetPostedMessagePayloadRegistry();
    std::lock_guard lock(registry.mutex);
    return ErasePostedMessagePayloadEntryLocked(registry, token, expectedHwnd, expectedMessage, validateDestination, removed);
}
} // namespace detail

// Call during window creation (WM_NCCREATE/WM_CREATE) for any window that can receive payload messages.
// This clears any previous drained state in case the HWND value is reused.
inline void InitPostedPayloadWindow(HWND hwnd) noexcept
{
    detail::InitPostedPayloadWindow(hwnd);
}

[[nodiscard]] inline size_t DrainPostedPayloadsForWindow(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return 0;
    }

    std::vector<detail::PostedMessagePayloadEntry> toDelete;
    {
        auto& registry = detail::GetPostedMessagePayloadRegistry();
        std::lock_guard lock(registry.mutex);

        static_cast<void>(registry.closedHwnds.insert(hwnd));

        const auto hwIt = registry.tokensByHwnd.find(hwnd);
        if (hwIt == registry.tokensByHwnd.end())
        {
            return 0;
        }

        toDelete.reserve(hwIt->second.size());
        const std::vector<LPARAM> tokens(hwIt->second.begin(), hwIt->second.end());
        for (const LPARAM token : tokens)
        {
            detail::PostedMessagePayloadEntry removed{};
            if (detail::ErasePostedMessagePayloadEntryLocked(registry, token, hwnd, 0, false, removed))
            {
                toDelete.push_back(removed);
            }
        }
    }

    for (const detail::PostedMessagePayloadEntry& entry : toDelete)
    {
        if (entry.del)
        {
            entry.del(entry.payload);
        }
    }

    return toDelete.size();
}

/// Posts a message with a unique_ptr payload behind an opaque, process-unique LPARAM token. If PostMessageW
/// fails, the payload is automatically deleted. Returns true on success, false on failure.
template <typename T> [[nodiscard]] inline bool PostMessagePayload(HWND hwnd, UINT msg, WPARAM wParam, std::unique_ptr<T> payload) noexcept
{
    T* raw = payload.release();
    if (raw == nullptr)
    {
        return PostMessageW(hwnd, msg, wParam, 0) != 0;
    }

    constexpr detail::MessagePayloadDeleter deleter = &detail::DeletePostedMessagePayload<T>;

    auto& registry = detail::GetPostedMessagePayloadRegistry();
    std::unique_lock lock(registry.mutex);

    if (! hwnd || registry.closedHwnds.contains(hwnd))
    {
        lock.unlock();
        deleter(raw);
        return false;
    }

    const LPARAM token = detail::AllocatePostedMessagePayloadTokenLocked(registry);
    if (token == 0)
    {
        lock.unlock();
        deleter(raw);
        return false;
    }

    static_cast<void>(registry.entriesByToken.emplace(token, detail::PostedMessagePayloadEntry{.payload = raw, .hwnd = hwnd, .msg = msg, .del = deleter}));
    static_cast<void>(registry.tokensByHwnd[hwnd].insert(token));

    if (! PostMessageW(hwnd, msg, wParam, token))
    {
        detail::PostedMessagePayloadEntry removed{};
        static_cast<void>(detail::ErasePostedMessagePayloadEntryLocked(registry, token, hwnd, msg, true, removed));

        lock.unlock();
        deleter(removed.payload);
        return false;
    }

    return true;
}

/// Takes ownership of a registered message payload identified by the opaque LPARAM token.
/// Returns null for a zero LPARAM or a stale, drained, unregistered, or differently typed token.
template <typename T> [[nodiscard]] inline std::unique_ptr<T> TakeMessagePayload(LPARAM lParam) noexcept
{
    detail::PostedMessagePayloadEntry entry{};
    if (! detail::TakeRegisteredPostedMessagePayload(lParam, nullptr, 0, false, entry))
    {
        return {};
    }

    constexpr detail::MessagePayloadDeleter expectedDeleter = &detail::DeletePostedMessagePayload<T>;
    if (entry.del != expectedDeleter)
    {
        entry.del(entry.payload);
        return {};
    }
    return std::unique_ptr<T>(static_cast<T*>(entry.payload));
}

template <typename T> struct ContiguousPostedPayloadDrainResult final
{
    ContiguousPostedPayloadDrainResult()                                                         = default;
    ContiguousPostedPayloadDrainResult(const ContiguousPostedPayloadDrainResult&)                = delete;
    ContiguousPostedPayloadDrainResult& operator=(const ContiguousPostedPayloadDrainResult&)     = delete;
    ContiguousPostedPayloadDrainResult(ContiguousPostedPayloadDrainResult&&) noexcept            = default;
    ContiguousPostedPayloadDrainResult& operator=(ContiguousPostedPayloadDrainResult&&) noexcept = default;

    std::unique_ptr<T> payload;
    uint64_t drainedPayloadCount = 0u;
    bool stoppedAtQueuedMessage  = false;
    MSG queuedMessage{};
};

/// Takes the payload from the message currently being dispatched and coalesces only immediately contiguous
/// queued payload messages with the same HWND, message ID, and operation key in wParam. The unfiltered thread
/// queue head is inspected before the matching message is removed, so unrelated input, completion, and other
/// operation messages retain their ordering. `canDrain` is called before inspecting the next message and may
/// enforce cancellation or caller-specific budgets. `reduce` owns the payload merge/replacement policy. Both
/// callbacks must be noexcept.
template <typename T, typename CanDrain, typename Reducer>
[[nodiscard]] inline ContiguousPostedPayloadDrainResult<T> TakeAndCoalesceContiguousPostedPayloads(
    HWND hwnd, UINT message, WPARAM operationKey, LPARAM currentLParam, CanDrain&& canDrain, Reducer&& reduce) noexcept
{
    ContiguousPostedPayloadDrainResult<T> result;
    result.payload = TakeMessagePayload<T>(currentLParam);
    if (! result.payload || ! hwnd)
    {
        return result;
    }

    auto&& canDrainRef = canDrain;
    auto&& reduceRef   = reduce;
    while (canDrainRef(*result.payload, result.drainedPayloadCount))
    {
        MSG queuedMessage{};
        if (PeekMessageW(&queuedMessage, nullptr, 0, 0, PM_NOREMOVE) == 0)
        {
            break;
        }
        if (queuedMessage.hwnd != hwnd || queuedMessage.message != message || queuedMessage.wParam != operationKey)
        {
            result.stoppedAtQueuedMessage = true;
            result.queuedMessage          = queuedMessage;
            break;
        }

        if (PeekMessageW(&queuedMessage, hwnd, message, message, PM_REMOVE) == 0)
        {
            break;
        }

        std::unique_ptr<T> newerPayload = TakeMessagePayload<T>(queuedMessage.lParam);
        if (! newerPayload)
        {
            continue;
        }

        reduceRef(result.payload, std::move(newerPayload));
        ++result.drainedPayloadCount;
        if (! result.payload)
        {
            break;
        }
    }
    return result;
}

// Macro Helpers for debug output
#ifdef _DEBUG
#define DBGOUT_INFO Debug::Info
#define DBGOUT_WARNING Debug::Warning
#define DBGOUT_ERROR Debug::Error
#define DBGOUT_ERROR_LASTERROR Debug::ErrorWithLastError
#else
#define DBGOUT_INFO __noop
#define DBGOUT_WARNING __noop
#define DBGOUT_ERROR __noop
#define DBGOUT_ERROR_LASTERROR __noop
#endif

// CallTracer: hierarchical indentation + performance measurement for Debug::Out messages on the current thread.
//
// Default behavior:
// - TRACER / TRACER_CTX: only logs the Exiting message (indentation still applies to all nested logs)
// - TRACER_INOUT / TRACER_INOUT_CTX: logs both Entering and Exiting messages
//
// Indentation is shared with Debug::Info/Warning/Error/Out on the same thread.

class CallTracer
{
public:
    enum class Mode : uint8_t
    {
        ExitOnly,
        EnterExit,
    };

    explicit CallTracer(const wchar_t* functionName, Mode mode = Mode::ExitOnly) noexcept : CallTracer(functionName, nullptr, mode)
    {
    }

    CallTracer(const wchar_t* functionName, const wchar_t* context, Mode mode = Mode::ExitOnly) noexcept
        : _enabled(Debug::detail::IsDebugEtwEnabled()),
          _functionName(functionName),
          _context(context),
          _mode(mode)
    {
        if (! _enabled)
        {
            return;
        }

        if (_mode == Mode::EnterExit)
        {
            if (_context)
            {
                Debug::Info(L"{} ({}) Entering", _functionName, _context);
            }
            else
            {
                Debug::Info(L"{} Entering", _functionName);
            }
        }

        Debug::detail::Indent();
        QueryPerformanceCounter(&_start);
    }

    ~CallTracer() noexcept
    {
        if (! _enabled)
        {
            return;
        }

        LARGE_INTEGER elapse{};
        QueryPerformanceCounter(&elapse);
        const double frequency = GetQpcFrequency();
        const double elapsedMs = (frequency > 0.0) ? (static_cast<double>(elapse.QuadPart - _start.QuadPart) * 1000.0 / frequency) : 0.0;

        Debug::detail::Unindent();

        if (_context)
        {
            Debug::Info(L"{} ({}) Exiting ({:.3f}ms)", _functionName, _context, elapsedMs);
        }
        else
        {
            Debug::Info(L"{} Exiting ({:.3f}ms)", _functionName, elapsedMs);
        }
    }

private:
    static double GetQpcFrequency() noexcept
    {
        static const double cached = []() noexcept
        {
            LARGE_INTEGER freq{};
            if (! QueryPerformanceFrequency(&freq) || freq.QuadPart <= 0)
            {
                return 0.0;
            }
            return static_cast<double>(freq.QuadPart);
        }();
        return cached;
    }

    bool _enabled                = false;
    const wchar_t* _functionName = nullptr;
    const wchar_t* _context      = nullptr;
    Mode _mode                   = Mode::ExitOnly;
    LARGE_INTEGER _start{};
};

// Helper macros for proper token concatenation
#define REDSAL_TRACER_CONCAT_IMPL(a, b) a##b
#define REDSAL_TRACER_CONCAT(a, b) REDSAL_TRACER_CONCAT_IMPL(a, b)

// Convert __FUNCTION__ (narrow string) to wide string at compile time
#define REDSAL_TRACER_WIDEN_IMPL(x) L##x
#define REDSAL_TRACER_WIDEN(x) REDSAL_TRACER_WIDEN_IMPL(x)

// Main tracing macros - use __FUNCTIONW__ (wide version) or convert __FUNCTION__
#if defined(__FUNCTIONW__)
#define TRACER [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(__FUNCTIONW__, CallTracer::Mode::ExitOnly)
#define TRACER_CTX(ctx) [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(__FUNCTIONW__, ctx, CallTracer::Mode::ExitOnly)
#define TRACER_INOUT [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(__FUNCTIONW__, CallTracer::Mode::EnterExit)
#define TRACER_INOUT_CTX(ctx) [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(__FUNCTIONW__, ctx, CallTracer::Mode::EnterExit)
#else
#define TRACER [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(REDSAL_TRACER_WIDEN(__FUNCTION__), CallTracer::Mode::ExitOnly)
#define TRACER_CTX(ctx)                                                                                                                                        \
    [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(REDSAL_TRACER_WIDEN(__FUNCTION__), ctx, CallTracer::Mode::ExitOnly)
#define TRACER_INOUT [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(REDSAL_TRACER_WIDEN(__FUNCTION__), CallTracer::Mode::EnterExit)
#define TRACER_INOUT_CTX(ctx)                                                                                                                                  \
    [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(REDSAL_TRACER_WIDEN(__FUNCTION__), ctx, CallTracer::Mode::EnterExit)
#endif

#define TRACER_CTW(ctx) TRACER_CTX(ctx)
#define TRACER_INOUT_CTW(ctx) TRACER_INOUT_CTX(ctx)

#pragma warning(pop)
