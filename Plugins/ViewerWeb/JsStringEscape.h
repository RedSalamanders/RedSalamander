#pragma once

#include <array>
#include <cstddef>
#include <cstdint>
#include <format>
#include <string>
#include <string_view>

// JavaScript-string escapers shared by ViewerWeb.cpp (the JSON / JSONL / Markdown document
// templates) and the ViewerWeb self-tests.
//
// SECURITY: the escaped text is concatenated into INLINE <script> string literals. The HTML
// tokenizer terminates a <script> element on the byte sequence "</script" regardless of JS
// string-literal quoting, so any '<' that survives escaping lets untrusted file content close the
// script and inject arbitrary HTML/JS. We therefore neutralize '<' (as \x3C) in addition to the
// usual quote/backslash/control escaping, and we neutralize the JS line/paragraph separators
// U+2028 / U+2029, which terminate a string literal even though they are not '\n'/'\r'.
namespace ViewerWebDetail
{

[[nodiscard]] inline bool TryAppendWithinLimit(std::string& output, std::string_view text, size_t limit) noexcept
{
    if (output.size() > limit || text.size() > limit - output.size())
    {
        return false;
    }
    output.append(text);
    return true;
}

// Appends directly to the publication buffer so JSONL rendering never keeps a
// second, fully escaped copy of attacker-controlled fields beside the HTML.
[[nodiscard]] inline bool TryAppendEscapedJavaScriptStringUtf8(
    std::string_view text, std::string& output, size_t limit) noexcept
{
    constexpr char kHex[] = "0123456789ABCDEF";
    const auto append = [&](std::string_view value) noexcept { return TryAppendWithinLimit(output, value, limit); };

    const size_t size = text.size();
    for (size_t i = 0u; i < size;)
    {
        const char ch = text[i];
        const auto u  = static_cast<uint8_t>(ch);
        if (u == 0xE2u && i + 2u < size && static_cast<uint8_t>(text[i + 1u]) == 0x80u)
        {
            const auto trail = static_cast<uint8_t>(text[i + 2u]);
            if (trail == 0xA8u || trail == 0xA9u)
            {
                if (! append(trail == 0xA8u ? "\\u2028" : "\\u2029"))
                {
                    return false;
                }
                i += 3u;
                continue;
            }
        }

        std::string_view replacement;
        switch (ch)
        {
            case '\\': replacement = "\\\\"; break;
            case '\'': replacement = "\\'"; break;
            case '\"': replacement = "\\\""; break;
            case '\r': replacement = "\\r"; break;
            case '\n': replacement = "\\n"; break;
            case '\t': replacement = "\\t"; break;
            case '<': replacement = "\\x3C"; break;
            default: break;
        }

        if (! replacement.empty())
        {
            if (! append(replacement))
            {
                return false;
            }
        }
        else if (u < 0x20u)
        {
            const std::array<char, 4u> escapedControl{'\\', 'x', kHex[(u >> 4u) & 0x0Fu], kHex[u & 0x0Fu]};
            if (! append(std::string_view(escapedControl.data(), escapedControl.size())))
            {
                return false;
            }
        }
        else if (output.size() >= limit)
        {
            return false;
        }
        else
        {
            output.push_back(ch);
        }
        ++i;
    }
    return true;
}

// Wide (UTF-16) variant. U+2028 / U+2029 are single code units in the BMP, so a simple per-unit
// switch suffices.
[[nodiscard]] inline std::wstring EscapeJavaScriptString(std::wstring_view text) noexcept
{
    std::wstring out;
    out.reserve(text.size() + 16);
    for (wchar_t ch : text)
    {
        switch (ch)
        {
            case L'\\': out += L"\\\\"; break;
            case L'\'': out += L"\\'"; break;
            case L'\"': out += L"\\\""; break;
            case L'\r': out += L"\\r"; break;
            case L'\n': out += L"\\n"; break;
            case L'\t': out += L"\\t"; break;
            case L'<': out += L"\\x3C"; break;     // neutralizes </script>, <!--, <![CDATA[
            case 0x2028: out += L"\\u2028"; break; // JS line separator
            case 0x2029: out += L"\\u2029"; break; // JS paragraph separator
            default:
                if (ch < 0x20)
                {
                    out += std::format(L"\\x{:02X}", static_cast<unsigned int>(ch));
                }
                else
                {
                    out.push_back(ch);
                }
                break;
        }
    }
    return out;
}

// UTF-8 (byte) variant. U+2028 / U+2029 arrive as the 3-byte sequences E2 80 A8 / E2 80 A9, so the
// loop is index-based with bounded look-ahead to detect and neutralize them.
[[nodiscard]] inline std::string EscapeJavaScriptStringUtf8(std::string_view text) noexcept
{
    std::string out;
    out.reserve(text.size() + text.size() / 8);
    const size_t size = text.size();
    for (size_t i = 0; i < size;)
    {
        const char ch = text[i];
        const auto u  = static_cast<uint8_t>(ch);

        // U+2028 (E2 80 A8) / U+2029 (E2 80 A9) terminate a JS string literal; emit  / .
        if (u == 0xE2u && i + 2 < size && static_cast<uint8_t>(text[i + 1]) == 0x80u)
        {
            const auto trail = static_cast<uint8_t>(text[i + 2]);
            if (trail == 0xA8u || trail == 0xA9u)
            {
                out += (trail == 0xA8u) ? "\\u2028" : "\\u2029";
                i += 3;
                continue;
            }
        }

        switch (ch)
        {
            case '\\': out += "\\\\"; break;
            case '\'': out += "\\'"; break;
            case '\"': out += "\\\""; break;
            case '\r': out += "\\r"; break;
            case '\n': out += "\\n"; break;
            case '\t': out += "\\t"; break;
            case '<': out += "\\x3C"; break; // neutralizes </script>, <!--, <![CDATA[
            default:
                if (u < 0x20u)
                {
                    out += std::format("\\x{:02X}", static_cast<unsigned int>(u));
                }
                else
                {
                    out.push_back(ch);
                }
                break;
        }
        ++i;
    }
    return out;
}

} // namespace ViewerWebDetail
