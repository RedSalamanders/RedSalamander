#pragma once

#include <limits>
#include <optional>
#include <string>
#include <string_view>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

namespace Common::Strings
{
// Converts UTF-8 using the Windows replacement-character policy (flags 0): malformed byte
// sequences are represented by U+FFFD rather than rejecting the entire value.
[[nodiscard]] inline std::wstring Utf16FromUtf8ReplacingInvalid(std::string_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), required);
    return written == required ? result : std::wstring{};
}

// Strict protocol/settings conversion. An engaged empty value represents valid empty input;
// nullopt represents malformed UTF-8, an oversized input, or a conversion failure.
[[nodiscard]] inline std::optional<std::wstring> TryUtf16FromUtf8Strict(std::string_view text) noexcept
{
    if (text.empty())
    {
        return std::wstring{};
    }
    if (text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return std::nullopt;
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return std::nullopt;
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return std::nullopt;
    }
    return result;
}

// Compatibility shape for existing strict parsers that historically represented conversion failure as empty.
[[nodiscard]] inline std::wstring Utf16FromUtf8StrictOrEmpty(std::string_view text) noexcept
{
    return TryUtf16FromUtf8Strict(text).value_or(std::wstring{});
}

// Converts UTF-16 using the Windows replacement-character policy (flags 0).
[[nodiscard]] inline std::string Utf8FromUtf16ReplacingInvalid(std::wstring_view text) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    return written == required ? result : std::string{};
}

// Strict protocol/settings conversion. An engaged empty value represents valid empty input;
// nullopt represents malformed UTF-16, an oversized input, or a conversion failure.
[[nodiscard]] inline std::optional<std::string> TryUtf8FromUtf16Strict(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return std::string{};
    }
    if (text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return std::nullopt;
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return std::nullopt;
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return std::nullopt;
    }
    return result;
}

// Compatibility shape for existing strict serializers that historically represented conversion failure as empty.
[[nodiscard]] inline std::string Utf8FromUtf16StrictOrEmpty(std::wstring_view text) noexcept
{
    return TryUtf8FromUtf16Strict(text).value_or(std::string{});
}
} // namespace Common::Strings
