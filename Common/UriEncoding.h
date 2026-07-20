#pragma once

#include "StringConversion.h"

#include <string>
#include <string_view>

namespace Common::Uri
{
enum class SlashPolicy
{
    Encode,
    Preserve,
};

[[nodiscard]] inline bool IsRfc3986UnreservedByte(unsigned char ch) noexcept
{
    return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9') || ch == '-' || ch == '.' || ch == '_' || ch == '~';
}

[[nodiscard]] inline std::string PercentEncodeBytes(std::string_view bytes, SlashPolicy slashPolicy = SlashPolicy::Encode)
{
    std::string encoded;
    encoded.reserve(bytes.size() * 3u);
    constexpr char kHex[] = "0123456789ABCDEF";
    for (const char value : bytes)
    {
        const unsigned char ch = static_cast<unsigned char>(value);
        if (IsRfc3986UnreservedByte(ch) || (slashPolicy == SlashPolicy::Preserve && ch == '/'))
        {
            encoded.push_back(static_cast<char>(ch));
            continue;
        }
        encoded.push_back('%');
        encoded.push_back(kHex[(ch >> 4u) & 0x0Fu]);
        encoded.push_back(kHex[ch & 0x0Fu]);
    }
    return encoded;
}

[[nodiscard]] inline bool TryPercentEncodeUtf8(std::wstring_view text,
                                               SlashPolicy slashPolicy,
                                               std::string& encodedOut) noexcept
{
    encodedOut.clear();
    const std::optional<std::string> utf8 = Common::Strings::TryUtf8FromUtf16Strict(text);
    if (! utf8.has_value())
    {
        return false;
    }
    encodedOut = PercentEncodeBytes(utf8.value(), slashPolicy);
    return true;
}

[[nodiscard]] inline bool TryPercentEncodeUtf8ToWide(std::wstring_view text,
                                                     SlashPolicy slashPolicy,
                                                     std::wstring& encodedOut) noexcept
{
    encodedOut.clear();
    std::string encoded;
    if (! TryPercentEncodeUtf8(text, slashPolicy, encoded))
    {
        return false;
    }
    encodedOut.assign(encoded.begin(), encoded.end());
    return true;
}
} // namespace Common::Uri
