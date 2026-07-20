#pragma once

#include <cstddef>
#include <cstdint>
#include <format>
#include <limits>
#include <string>
#include <string_view>

namespace Common::Parsing
{
enum class ThroughputBoundaryWhitespacePolicy
{
    AsciiWhitespace,
    ControlCharactersThroughSpace,
};

namespace Detail
{
[[nodiscard]] inline bool IsBoundaryWhitespace(wchar_t ch, ThroughputBoundaryWhitespacePolicy policy) noexcept
{
    if (policy == ThroughputBoundaryWhitespacePolicy::ControlCharactersThroughSpace)
    {
        return ch <= L' ';
    }

    return ch == L' ' || ch == L'\t' || ch == L'\n' || ch == L'\r' || ch == L'\f' || ch == L'\v';
}

[[nodiscard]] inline std::wstring_view TrimBoundaryWhitespace(std::wstring_view text, ThroughputBoundaryWhitespacePolicy policy) noexcept
{
    while (! text.empty() && IsBoundaryWhitespace(text.front(), policy))
    {
        text.remove_prefix(1u);
    }
    while (! text.empty() && IsBoundaryWhitespace(text.back(), policy))
    {
        text.remove_suffix(1u);
    }
    return text;
}

[[nodiscard]] inline bool EqualsAsciiNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    if (left.size() != right.size())
    {
        return false;
    }

    for (size_t index = 0u; index < left.size(); ++index)
    {
        wchar_t lhs = left[index];
        wchar_t rhs = right[index];
        if (lhs >= L'A' && lhs <= L'Z')
        {
            lhs = static_cast<wchar_t>(lhs + (L'a' - L'A'));
        }
        if (rhs >= L'A' && rhs <= L'Z')
        {
            rhs = static_cast<wchar_t>(rhs + (L'a' - L'A'));
        }
        if (lhs != rhs)
        {
            return false;
        }
    }
    return true;
}
} // namespace Detail

// Parses the established File Operations edit grammar. Unit spellings are binary even when the optional
// "i" is omitted, and a bare number means KiB/s. Empty input means unlimited (zero).
[[nodiscard]] inline bool TryParseBinaryThroughputText(std::wstring_view text,
                                                       ThroughputBoundaryWhitespacePolicy whitespacePolicy,
                                                       uint64_t& outBytesPerSecond) noexcept
{
    constexpr uint64_t kKiB = 1024ull;
    constexpr uint64_t kMiB = 1024ull * 1024ull;
    constexpr uint64_t kGiB = 1024ull * 1024ull * 1024ull;
    constexpr uint64_t kTiB = 1024ull * 1024ull * 1024ull * 1024ull;
    constexpr uint64_t kPiB = 1024ull * 1024ull * 1024ull * 1024ull * 1024ull;

    outBytesPerSecond = 0u;
    text              = Detail::TrimBoundaryWhitespace(text, whitespacePolicy);
    if (text.empty())
    {
        return true;
    }

    bool sawDigit          = false;
    bool sawDecimal        = false;
    double number          = 0.0;
    double fractionalScale = 0.1;
    size_t index           = 0u;
    for (; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch >= L'0' && ch <= L'9')
        {
            sawDigit                 = true;
            const unsigned int digit = static_cast<unsigned int>(ch - L'0');
            if (! sawDecimal)
            {
                number = (number * 10.0) + static_cast<double>(digit);
            }
            else
            {
                number += static_cast<double>(digit) * fractionalScale;
                fractionalScale *= 0.1;
            }
            continue;
        }

        if ((ch == L'.' || ch == L',') && ! sawDecimal)
        {
            sawDecimal = true;
            continue;
        }
        break;
    }

    if (! sawDigit)
    {
        return false;
    }

    std::wstring_view unit = Detail::TrimBoundaryWhitespace(text.substr(index), whitespacePolicy);
    if (unit.size() >= 2u)
    {
        const wchar_t penultimate = unit[unit.size() - 2u];
        const wchar_t last        = unit.back();
        if (penultimate == L'/' && (last == L's' || last == L'S'))
        {
            unit.remove_suffix(2u);
            unit = Detail::TrimBoundaryWhitespace(unit, whitespacePolicy);
        }
    }

    uint64_t multiplier = 0u;
    if (unit.empty() || Detail::EqualsAsciiNoCase(unit, L"kb") || Detail::EqualsAsciiNoCase(unit, L"k") || Detail::EqualsAsciiNoCase(unit, L"kib"))
    {
        multiplier = kKiB;
    }
    else if (Detail::EqualsAsciiNoCase(unit, L"b"))
    {
        multiplier = 1u;
    }
    else if (Detail::EqualsAsciiNoCase(unit, L"mb") || Detail::EqualsAsciiNoCase(unit, L"m") || Detail::EqualsAsciiNoCase(unit, L"mib"))
    {
        multiplier = kMiB;
    }
    else if (Detail::EqualsAsciiNoCase(unit, L"gb") || Detail::EqualsAsciiNoCase(unit, L"g") || Detail::EqualsAsciiNoCase(unit, L"gib"))
    {
        multiplier = kGiB;
    }
    else if (Detail::EqualsAsciiNoCase(unit, L"tb") || Detail::EqualsAsciiNoCase(unit, L"t") || Detail::EqualsAsciiNoCase(unit, L"tib"))
    {
        multiplier = kTiB;
    }
    else if (Detail::EqualsAsciiNoCase(unit, L"pb") || Detail::EqualsAsciiNoCase(unit, L"p") || Detail::EqualsAsciiNoCase(unit, L"pib"))
    {
        multiplier = kPiB;
    }
    else
    {
        return false;
    }

    const double result = number * static_cast<double>(multiplier);
    if (result <= 0.0)
    {
        outBytesPerSecond = 0u;
        return true;
    }

    constexpr double maxValue = static_cast<double>((std::numeric_limits<uint64_t>::max)());
    if (result >= maxValue)
    {
        outBytesPerSecond = (std::numeric_limits<uint64_t>::max)();
        return true;
    }

    outBytesPerSecond = static_cast<uint64_t>(result + 0.5);
    return true;
}

// Formats the nonlocalized edit grammar consumed by TryParseBinaryThroughputText.
[[nodiscard]] inline std::wstring FormatBinaryThroughputText(uint64_t bytesPerSecond) noexcept
{
    constexpr uint64_t kKiB = 1024ull;
    constexpr uint64_t kMiB = 1024ull * 1024ull;
    constexpr uint64_t kGiB = 1024ull * 1024ull * 1024ull;

    if (bytesPerSecond == 0u)
    {
        return {};
    }
    if ((bytesPerSecond % kGiB) == 0u)
    {
        return std::format(L"{} GiB/s", bytesPerSecond / kGiB);
    }
    if ((bytesPerSecond % kMiB) == 0u)
    {
        return std::format(L"{} MiB/s", bytesPerSecond / kMiB);
    }
    if ((bytesPerSecond % kKiB) == 0u)
    {
        return std::format(L"{} KiB/s", bytesPerSecond / kKiB);
    }
    return std::format(L"{} B/s", bytesPerSecond);
}
} // namespace Common::Parsing
