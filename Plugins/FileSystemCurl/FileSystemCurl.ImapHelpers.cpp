#include "FileSystemCurl.ImapHelpers.h"
#include "Helpers.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#include <algorithm>
#include <format>
#include <limits>

namespace FileSystemCurlInternal
{
namespace
{
[[nodiscard]] constexpr char AsciiLower(char ch) noexcept
{
    if (ch >= 'A' && ch <= 'Z')
    {
        return static_cast<char>(ch - 'A' + 'a');
    }
    return ch;
}

[[nodiscard]] constexpr wchar_t AsciiLower(wchar_t ch) noexcept
{
    if (ch >= L'A' && ch <= L'Z')
    {
        return static_cast<wchar_t>(ch - L'A' + L'a');
    }
    return ch;
}

[[nodiscard]] bool EndsWithNoCase(std::wstring_view text, std::wstring_view suffix) noexcept
{
    if (text.size() < suffix.size())
    {
        return false;
    }

    const std::wstring_view tail = text.substr(text.size() - suffix.size());
    for (size_t i = 0; i < suffix.size(); ++i)
    {
        if (AsciiLower(tail[i]) != AsciiLower(suffix[i]))
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] bool EqualsAsciiNoCase(std::string_view left, std::string_view right) noexcept
{
    if (left.size() != right.size())
    {
        return false;
    }

    for (size_t index = 0; index < left.size(); ++index)
    {
        if (AsciiLower(left[index]) != AsciiLower(right[index]))
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] bool TryParseAsciiUint64(std::wstring_view text, uint64_t& out) noexcept
{
    out = 0;
    if (text.empty())
    {
        return false;
    }

    uint64_t value = 0;
    for (const wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }

        const uint64_t digit = static_cast<uint64_t>(ch - L'0');
        if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10u)
        {
            return false;
        }
        value = (value * 10u) + digit;
    }

    out = value;
    return true;
}

[[nodiscard]] std::wstring SanitizeImapMessageNamePart(std::wstring_view text) noexcept
{
    std::wstring out;
    out.reserve(text.size());

    for (const wchar_t ch : text)
    {
        if (ch < 0x20)
        {
            out.push_back(L'_');
            continue;
        }

        switch (ch)
        {
            case L'<':
            case L'>':
            case L':':
            case L'"':
            case L'/':
            case L'\\':
            case L'|':
            case L'?':
            case L'*': out.push_back(L'_'); break;
            default: out.push_back(ch); break;
        }
    }

    while (! out.empty() && (out.back() == L' ' || out.back() == L'.'))
    {
        out.pop_back();
    }

    return out;
}

[[nodiscard]] std::wstring TruncateSubjectForLeafName(std::wstring_view text, size_t maxChars) noexcept
{
    if (text.size() <= maxChars)
    {
        return std::wstring(text);
    }

    if (maxChars <= 3u)
    {
        return std::wstring(maxChars, L'.');
    }

    std::wstring out(text.substr(0, maxChars - 3u));
    out.append(L"...");
    return out;
}

[[nodiscard]] bool IsAsciiWhitespace(char ch) noexcept
{
    return ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n';
}

[[nodiscard]] bool IsAsciiHexDigit(char ch) noexcept
{
    return (ch >= '0' && ch <= '9') || (ch >= 'a' && ch <= 'f') || (ch >= 'A' && ch <= 'F');
}

[[nodiscard]] uint8_t AsciiHexValue(char ch) noexcept
{
    if (ch >= '0' && ch <= '9')
    {
        return static_cast<uint8_t>(ch - '0');
    }
    if (ch >= 'a' && ch <= 'f')
    {
        return static_cast<uint8_t>(10 + (ch - 'a'));
    }
    if (ch >= 'A' && ch <= 'F')
    {
        return static_cast<uint8_t>(10 + (ch - 'A'));
    }
    return 0;
}

[[nodiscard]] size_t FindAsciiNoCase(std::string_view haystack, std::string_view needle, size_t start) noexcept
{
    if (needle.empty())
    {
        return start <= haystack.size() ? start : std::string_view::npos;
    }

    for (size_t i = start; i + needle.size() <= haystack.size(); ++i)
    {
        bool match = true;
        for (size_t j = 0; j < needle.size(); ++j)
        {
            if (AsciiLower(haystack[i + j]) != AsciiLower(needle[j]))
            {
                match = false;
                break;
            }
        }
        if (match)
        {
            return i;
        }
    }

    return std::string_view::npos;
}

[[nodiscard]] bool TryParseAsciiUint64(std::string_view text, size_t& pos, uint64_t& out) noexcept
{
    out            = 0;
    uint64_t value = 0;
    size_t digits  = 0;
    while (pos < text.size() && text[pos] >= '0' && text[pos] <= '9')
    {
        const uint64_t digit = static_cast<uint64_t>(text[pos] - '0');
        if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10u)
        {
            return false;
        }
        value = (value * 10u) + digit;
        ++digits;
        ++pos;
    }

    if (digits == 0)
    {
        return false;
    }

    out = value;
    return true;
}

[[nodiscard]] bool TryDecodeRfc2047Q(std::string_view encodedText, std::string& outBytes) noexcept
{
    outBytes.clear();
    outBytes.reserve(encodedText.size());

    for (size_t i = 0; i < encodedText.size(); ++i)
    {
        const char ch = encodedText[i];
        if (ch == '_')
        {
            outBytes.push_back(' ');
            continue;
        }

        if (ch == '=' && i + 2u < encodedText.size() && IsAsciiHexDigit(encodedText[i + 1u]) && IsAsciiHexDigit(encodedText[i + 2u]))
        {
            const uint8_t value = static_cast<uint8_t>((AsciiHexValue(encodedText[i + 1u]) << 4) | AsciiHexValue(encodedText[i + 2u]));
            outBytes.push_back(static_cast<char>(value));
            i += 2u;
            continue;
        }

        if (ch == '=' && i + 1u == encodedText.size())
        {
            continue;
        }

        outBytes.push_back(ch);
    }

    return true;
}

[[nodiscard]] int Base64Value(char ch) noexcept
{
    if (ch >= 'A' && ch <= 'Z')
    {
        return ch - 'A';
    }
    if (ch >= 'a' && ch <= 'z')
    {
        return 26 + (ch - 'a');
    }
    if (ch >= '0' && ch <= '9')
    {
        return 52 + (ch - '0');
    }
    if (ch == '+')
    {
        return 62;
    }
    if (ch == '/')
    {
        return 63;
    }
    return -1;
}

[[nodiscard]] bool TryDecodeRfc2047B(std::string_view encodedText, std::string& outBytes) noexcept
{
    outBytes.clear();
    outBytes.reserve((encodedText.size() * 3u) / 4u);

    uint32_t acc = 0;
    int bits     = 0;

    for (const char ch : encodedText)
    {
        if (IsAsciiWhitespace(ch))
        {
            continue;
        }

        if (ch == '=')
        {
            break;
        }

        const int v = Base64Value(ch);
        if (v < 0)
        {
            return false;
        }

        acc = (acc << 6) | static_cast<uint32_t>(v);
        bits += 6;

        if (bits >= 8)
        {
            bits -= 8;
            const uint8_t byte = static_cast<uint8_t>((acc >> bits) & 0xFFu);
            outBytes.push_back(static_cast<char>(byte));
        }
    }

    return true;
}

[[nodiscard]] std::wstring Utf16FromCodePage(std::string_view text, UINT codePage, DWORD flags = 0) noexcept
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int required = MultiByteToWideChar(codePage, flags, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring out(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(codePage, flags, text.data(), static_cast<int>(text.size()), out.data(), required);
    if (written != required)
    {
        return {};
    }

    return out;
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

[[nodiscard]] std::wstring Utf16FromHeaderLiteral(std::string_view text) noexcept
{
    std::wstring out = Utf16FromUtf8(text);
    if (! out.empty() || text.empty())
    {
        return out;
    }

    out = Utf16FromCodePage(text, CP_ACP);
    if (! out.empty())
    {
        return out;
    }

    out.reserve(text.size());
    for (const char ch : text)
    {
        out.push_back(static_cast<wchar_t>(static_cast<unsigned char>(ch)));
    }
    return out;
}

[[nodiscard]] std::string NormalizeMimeCharsetLabel(std::string_view charset) noexcept
{
    std::string normalized;
    normalized.reserve(charset.size());
    for (const char ch : charset)
    {
        if (ch >= 'A' && ch <= 'Z')
        {
            normalized.push_back(static_cast<char>(ch - 'A' + 'a'));
        }
        else if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9'))
        {
            normalized.push_back(ch);
        }
    }
    return normalized;
}

[[nodiscard]] std::optional<UINT> CodePageFromMimeCharset(std::string_view charset) noexcept
{
    const std::string label = NormalizeMimeCharsetLabel(charset);
    if (label.empty())
    {
        return std::nullopt;
    }

    if (label == "utf8" || label == "unicode11utf8")
    {
        return CP_UTF8;
    }
    if (label == "usascii" || label == "ascii" || label == "ansi_x341968" || label == "iso646us")
    {
        return 20127u;
    }
    if (label == "utf16" || label == "utf16le")
    {
        return 1200u;
    }
    if (label == "utf16be")
    {
        return 1201u;
    }
    if (label == "shiftjis" || label == "shiftjisx0213" || label == "sjis" || label == "mskanji" || label == "csshiftjis")
    {
        return 932u;
    }
    if (label == "eucjp" || label == "xeucjp")
    {
        return 51932u;
    }
    if (label == "iso2022jp" || label == "csiso2022jp")
    {
        return 50220u;
    }
    if (label == "gb18030")
    {
        return 54936u;
    }
    if (label == "gb2312" || label == "gbk" || label == "chinesesimplified" || label == "cp936")
    {
        return 936u;
    }
    if (label == "big5" || label == "big5hkscs" || label == "cnbig5")
    {
        return 950u;
    }
    if (label == "euckr" || label == "ksc5601" || label == "ksc56011987" || label == "cp949")
    {
        return 949u;
    }
    if (label == "koi8r" || label == "cskoi8r")
    {
        return 20866u;
    }
    if (label == "koi8u")
    {
        return 21866u;
    }
    if (label == "tis620" || label == "windows874" || label == "iso885911")
    {
        return 874u;
    }
    if (label == "macintosh" || label == "macroman")
    {
        return 10000u;
    }

    if (label.starts_with("windows") && label.size() > 7u)
    {
        uint64_t value = 0;
        size_t pos     = 0;
        const std::string_view digits(std::string_view(label).substr(7u));
        if (TryParseAsciiUint64(digits, pos, value) && pos == digits.size() && value <= (std::numeric_limits<UINT>::max)())
        {
            return static_cast<UINT>(value);
        }
    }
    if (label.starts_with("cp") && label.size() > 2u)
    {
        uint64_t value = 0;
        size_t pos     = 0;
        const std::string_view digits(std::string_view(label).substr(2u));
        if (TryParseAsciiUint64(digits, pos, value) && pos == digits.size() && value <= (std::numeric_limits<UINT>::max)())
        {
            return static_cast<UINT>(value);
        }
    }

    if (label.starts_with("iso8859"))
    {
        uint64_t part = 0;
        size_t pos    = 0;
        const std::string_view digits(std::string_view(label).substr(7u));
        if (TryParseAsciiUint64(digits, pos, part) && pos == digits.size())
        {
            switch (part)
            {
                case 1: return 28591u;
                case 2: return 28592u;
                case 3: return 28593u;
                case 4: return 28594u;
                case 5: return 28595u;
                case 6: return 28596u;
                case 7: return 28597u;
                case 8: return 28598u;
                case 9: return 28599u;
                case 13: return 28603u;
                case 15: return 28605u;
                default: break;
            }
        }
    }

    return std::nullopt;
}

[[nodiscard]] std::wstring DecodeBytesForMimeCharset(std::string_view charset, std::string_view bytes) noexcept
{
    if (bytes.empty())
    {
        return {};
    }

    if (const std::optional<UINT> codePage = CodePageFromMimeCharset(charset); codePage.has_value())
    {
        const DWORD flags = codePage.value() == CP_UTF8 ? MB_ERR_INVALID_CHARS : 0;
        std::wstring out  = Utf16FromCodePage(bytes, codePage.value(), flags);
        if (! out.empty())
        {
            return out;
        }
    }

    std::wstring out = Utf16FromUtf8(bytes);
    if (! out.empty())
    {
        return out;
    }

    out = Utf16FromCodePage(bytes, CP_ACP);
    if (! out.empty())
    {
        return out;
    }

    return Utf16FromCodePage(bytes, 1252u);
}

void NormalizeDecodedMimeSubject(std::wstring& text) noexcept
{
    for (wchar_t& ch : text)
    {
        if (ch == static_cast<wchar_t>(0x00AD))
        {
            ch = L'-';
        }
    }
}

struct EncodedWord
{
    std::string_view charset;
    std::string_view encoding;
    std::string_view encodedText;
    size_t nextPos = 0;
};

[[nodiscard]] bool TryParseStandardEncodedWord(std::string_view text, size_t marker, EncodedWord& out) noexcept
{
    if (marker + 2u > text.size() || text.substr(marker, 2u) != "=?")
    {
        return false;
    }

    const size_t q1  = text.find('?', marker + 2u);
    const size_t q2  = (q1 != std::string_view::npos) ? text.find('?', q1 + 1u) : std::string_view::npos;
    const size_t end = (q2 != std::string_view::npos) ? text.find("?=", q2 + 1u) : std::string_view::npos;
    if (q1 == std::string_view::npos || q2 == std::string_view::npos || end == std::string_view::npos)
    {
        return false;
    }

    out.charset     = text.substr(marker + 2u, q1 - (marker + 2u));
    out.encoding    = text.substr(q1 + 1u, q2 - (q1 + 1u));
    out.encodedText = text.substr(q2 + 1u, end - (q2 + 1u));
    out.nextPos     = end + 2u;
    return ! out.charset.empty() && ! out.encoding.empty();
}

[[nodiscard]] bool LooksLikeEncodedWordStart(std::string_view text, size_t pos) noexcept
{
    return pos + 2u < text.size() && text[pos] == '=' && (text[pos + 1u] == '?' || text[pos + 1u] == '_');
}

[[nodiscard]] bool TryParseUnderscoreEncodedWord(std::string_view text, size_t marker, EncodedWord& out) noexcept
{
    if (marker + 2u > text.size() || text.substr(marker, 2u) != "=_")
    {
        return false;
    }

    const size_t c1 = text.find('_', marker + 2u);
    const size_t c2 = (c1 != std::string_view::npos) ? text.find('_', c1 + 1u) : std::string_view::npos;
    if (c1 == std::string_view::npos || c2 == std::string_view::npos || c2 + 1u >= text.size())
    {
        return false;
    }

    size_t encodedEnd = text.size();
    size_t nextPos    = text.size();
    for (size_t i = c2 + 1u; i < text.size(); ++i)
    {
        if (i + 1u < text.size() && text[i] == '_' && text[i + 1u] == '=' &&
            (i + 2u >= text.size() || IsAsciiWhitespace(text[i + 2u]) || LooksLikeEncodedWordStart(text, i + 2u)))
        {
            encodedEnd = i;
            nextPos    = i + 2u;
            break;
        }

        if (text[i] == '=' && i + 1u < text.size() && IsAsciiWhitespace(text[i + 1u]))
        {
            size_t ws = i + 1u;
            while (ws < text.size() && IsAsciiWhitespace(text[ws]))
            {
                ++ws;
            }
            if (LooksLikeEncodedWordStart(text, ws))
            {
                encodedEnd = i;
                nextPos    = i + 1u;
                break;
            }
        }

        if (IsAsciiWhitespace(text[i]))
        {
            size_t ws = i;
            while (ws < text.size() && IsAsciiWhitespace(text[ws]))
            {
                ++ws;
            }
            if (LooksLikeEncodedWordStart(text, ws))
            {
                encodedEnd = i;
                nextPos    = i;
                break;
            }
        }
    }

    out.charset     = text.substr(marker + 2u, c1 - (marker + 2u));
    out.encoding    = text.substr(c1 + 1u, c2 - (c1 + 1u));
    out.encodedText = text.substr(c2 + 1u, encodedEnd - (c2 + 1u));
    out.nextPos     = nextPos;
    return ! out.charset.empty() && ! out.encoding.empty() && ! out.encodedText.empty();
}

[[nodiscard]] size_t FindNextEncodedWordMarker(std::string_view text, size_t pos) noexcept
{
    const size_t standard   = text.find("=?", pos);
    const size_t underscore = text.find("=_", pos);
    return std::min(standard, underscore);
}

[[nodiscard]] bool TryDecodeEncodedWord(const EncodedWord& word, std::wstring& out) noexcept
{
    out.clear();

    std::string bytes;
    bool bytesOk = false;
    if (word.encoding.size() == 1u && (word.encoding[0] == 'Q' || word.encoding[0] == 'q'))
    {
        bytesOk = TryDecodeRfc2047Q(word.encodedText, bytes);
    }
    else if (word.encoding.size() == 1u && (word.encoding[0] == 'B' || word.encoding[0] == 'b'))
    {
        bytesOk = TryDecodeRfc2047B(word.encodedText, bytes);
    }

    if (! bytesOk)
    {
        return false;
    }

    out = DecodeBytesForMimeCharset(word.charset, bytes);
    NormalizeDecodedMimeSubject(out);
    return ! out.empty() || bytes.empty();
}
} // namespace

[[nodiscard]] bool TryParseImapUidFromLeafName(std::wstring_view leafName, uint64_t& outUid) noexcept
{
    outUid = 0;

    ImapMessageIdentity identity;
    if (TryParseImapMessageIdentityFromLeafName(leafName, identity))
    {
        outUid = identity.uid;
        return true;
    }

    constexpr std::wstring_view kExt = L".eml";
    if (! EndsWithNoCase(leafName, kExt))
    {
        return false;
    }

    const std::wstring_view base = leafName.substr(0, leafName.size() - kExt.size());
    if (base.empty())
    {
        return false;
    }

    if (TryParseAsciiUint64(base, outUid))
    {
        return true;
    }

    if (base.back() == L']')
    {
        const size_t open = base.rfind(L'[');
        if (open != std::wstring_view::npos && open + 1u < base.size() - 1u)
        {
            return TryParseAsciiUint64(base.substr(open + 1u, base.size() - open - 2u), outUid);
        }
        return false;
    }

    return false;
}

[[nodiscard]] bool TryParseImapMessageIdentityFromLeafName(std::wstring_view leafName, ImapMessageIdentity& outIdentity) noexcept
{
    outIdentity = {};

    constexpr std::wstring_view kExt = L".eml";
    if (! EndsWithNoCase(leafName, kExt))
    {
        return false;
    }

    const std::wstring_view base = leafName.substr(0, leafName.size() - kExt.size());
    if (base.empty() || base.back() != L']')
    {
        return false;
    }

    const size_t open = base.rfind(L'[');
    if (open == std::wstring_view::npos || open + 1u >= base.size() - 1u)
    {
        return false;
    }

    const std::wstring_view encodedIdentity = base.substr(open + 1u, base.size() - open - 2u);
    const size_t separator                  = encodedIdentity.find(L'-');
    if (separator == std::wstring_view::npos || separator == 0u || separator + 1u >= encodedIdentity.size() ||
        encodedIdentity.find(L'-', separator + 1u) != std::wstring_view::npos)
    {
        return false;
    }

    if (! TryParseAsciiUint64(encodedIdentity.substr(0, separator), outIdentity.uidValidity) ||
        ! TryParseAsciiUint64(encodedIdentity.substr(separator + 1u), outIdentity.uid) || outIdentity.uidValidity == 0u || outIdentity.uid == 0u)
    {
        outIdentity = {};
        return false;
    }
    return true;
}

[[nodiscard]] std::wstring DecodeRfc2047EncodedWordsToUtf16(std::string_view headerValue) noexcept
{
    if (headerValue.empty())
    {
        return {};
    }

    std::wstring out;
    bool appendedAnything = false;

    size_t pos = 0;
    while (pos < headerValue.size())
    {
        const size_t marker = FindNextEncodedWordMarker(headerValue, pos);
        if (marker == std::string_view::npos)
        {
            std::wstring tail = Utf16FromHeaderLiteral(headerValue.substr(pos));
            NormalizeDecodedMimeSubject(tail);
            out.append(tail);
            appendedAnything = appendedAnything || ! tail.empty();
            break;
        }

        if (marker > pos)
        {
            std::wstring plain = Utf16FromHeaderLiteral(headerValue.substr(pos, marker - pos));
            NormalizeDecodedMimeSubject(plain);
            if (! plain.empty())
            {
                out.append(plain);
                appendedAnything = true;
            }
        }

        EncodedWord word;
        const bool parsed = TryParseStandardEncodedWord(headerValue, marker, word) || TryParseUnderscoreEncodedWord(headerValue, marker, word);
        if (! parsed)
        {
            out.append(Utf16FromHeaderLiteral(headerValue.substr(marker, 2u)));
            appendedAnything = true;
            pos              = marker + 2u;
            continue;
        }

        std::wstring decoded;
        if (TryDecodeEncodedWord(word, decoded))
        {
            out.append(decoded);
            appendedAnything = appendedAnything || ! decoded.empty();
        }
        else
        {
            std::wstring literal = Utf16FromHeaderLiteral(headerValue.substr(marker, word.nextPos - marker));
            NormalizeDecodedMimeSubject(literal);
            out.append(literal);
            appendedAnything = appendedAnything || ! literal.empty();
        }

        pos = word.nextPos;

        // RFC 2047 says whitespace between adjacent encoded-words is not display text.
        size_t ws = pos;
        while (ws < headerValue.size() && IsAsciiWhitespace(headerValue[ws]))
        {
            ++ws;
        }
        if (ws > pos && LooksLikeEncodedWordStart(headerValue, ws))
        {
            pos = ws;
        }
    }

    if (! appendedAnything)
    {
        std::wstring fallback = Utf16FromHeaderLiteral(headerValue);
        NormalizeDecodedMimeSubject(fallback);
        return fallback;
    }

    return out;
}

[[nodiscard]] std::wstring BuildImapMessageLeafName(std::wstring_view subject, std::wstring_view from, uint64_t uidValidity, uint64_t uid) noexcept
{
    static_cast<void>(from);

    std::wstring safeSubject = SanitizeImapMessageNamePart(subject.empty() ? L"(no subject)" : subject);
    safeSubject              = TruncateSubjectForLeafName(safeSubject, 96u);

    const std::wstring_view subjectPart = safeSubject.empty() ? std::wstring_view(L"message") : std::wstring_view(safeSubject);
    const std::wstring identityText     = std::format(L"{}-{}", uidValidity, uid);

    std::wstring out;
    out.reserve(subjectPart.size() + identityText.size() + 8u);
    out.append(subjectPart);
    out.append(L" [");
    out.append(identityText);
    out.append(L"].eml");
    return out;
}

[[nodiscard]] bool TryParseImapMailboxStatus(std::string_view response, ImapMailboxStatus& out) noexcept
{
    out = {};

    const size_t statusPos = FindAsciiNoCase(response, "STATUS", 0);
    if (statusPos == std::string_view::npos)
    {
        return false;
    }

    const size_t open = response.find('(', statusPos);
    if (open == std::string_view::npos)
    {
        return false;
    }

    const size_t close = response.find(')', open + 1u);
    if (close == std::string_view::npos || close <= open)
    {
        return false;
    }

    const std::string_view body = response.substr(open + 1u, close - open - 1u);
    bool recognized             = false;
    size_t pos                  = 0;
    while (pos < body.size())
    {
        while (pos < body.size() && IsAsciiWhitespace(body[pos]))
        {
            ++pos;
        }

        const size_t keyStart = pos;
        while (pos < body.size() && ! IsAsciiWhitespace(body[pos]))
        {
            ++pos;
        }

        const std::string_view key = body.substr(keyStart, pos - keyStart);
        while (pos < body.size() && IsAsciiWhitespace(body[pos]))
        {
            ++pos;
        }

        uint64_t value = 0;
        if (key.empty() || ! TryParseAsciiUint64(body, pos, value))
        {
            while (pos < body.size() && ! IsAsciiWhitespace(body[pos]))
            {
                ++pos;
            }
            continue;
        }

        if (key.size() == 8u && FindAsciiNoCase(key, "MESSAGES", 0) == 0)
        {
            out.messages = value;
            recognized   = true;
        }
        else if (key.size() == 6u && FindAsciiNoCase(key, "RECENT", 0) == 0)
        {
            out.recent = value;
            recognized = true;
        }
        else if (key.size() == 7u && FindAsciiNoCase(key, "UIDNEXT", 0) == 0)
        {
            out.uidNext = value;
            recognized  = true;
        }
        else if (key.size() == 11u && FindAsciiNoCase(key, "UIDVALIDITY", 0) == 0)
        {
            out.uidValidity = value;
            recognized      = true;
        }
        else if (key.size() == 6u && FindAsciiNoCase(key, "UNSEEN", 0) == 0)
        {
            out.unseen = value;
            recognized = true;
        }
    }

    return recognized;
}

[[nodiscard]] bool TryParseImapCapabilities(std::string_view response, ImapCapabilities& out) noexcept
{
    out = {};

    bool foundCapabilityLine = false;
    size_t lineStart         = 0;
    while (lineStart < response.size())
    {
        size_t lineEnd = response.find('\n', lineStart);
        if (lineEnd == std::string_view::npos)
        {
            lineEnd = response.size();
        }

        std::string_view line = response.substr(lineStart, lineEnd - lineStart);
        if (! line.empty() && line.back() == '\r')
        {
            line.remove_suffix(1u);
        }
        lineStart = lineEnd + 1u;

        size_t pos = 0;
        while (pos < line.size() && IsAsciiWhitespace(line[pos]))
        {
            ++pos;
        }
        if (pos >= line.size() || line[pos] != '*')
        {
            continue;
        }
        ++pos;
        while (pos < line.size() && IsAsciiWhitespace(line[pos]))
        {
            ++pos;
        }

        const size_t commandStart = pos;
        while (pos < line.size() && ! IsAsciiWhitespace(line[pos]))
        {
            ++pos;
        }
        if (! EqualsAsciiNoCase(line.substr(commandStart, pos - commandStart), "CAPABILITY"))
        {
            continue;
        }

        foundCapabilityLine = true;
        while (pos < line.size())
        {
            while (pos < line.size() && IsAsciiWhitespace(line[pos]))
            {
                ++pos;
            }
            const size_t tokenStart = pos;
            while (pos < line.size() && ! IsAsciiWhitespace(line[pos]))
            {
                ++pos;
            }
            if (EqualsAsciiNoCase(line.substr(tokenStart, pos - tokenStart), "UIDPLUS"))
            {
                out.uidPlus = true;
            }
        }
    }

    return foundCapabilityLine;
}

[[nodiscard]] HRESULT ValidateImapMessageUidValidity(uint64_t expectedUidValidity, const std::optional<uint64_t>& observedUidValidity) noexcept
{
    if (expectedUidValidity == 0u)
    {
        return E_INVALIDARG;
    }
    if (! observedUidValidity.has_value() || observedUidValidity.value() == 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (observedUidValidity.value() != expectedUidValidity)
    {
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
    return S_OK;
}

[[nodiscard]] HRESULT ExecuteImapSingleMessageDelete(bool uidPlusAvailable,
                                                     uint64_t uid,
                                                     ImapDeleteCommandExecutor executor,
                                                     void* context,
                                                     ImapDeleteOutcome& outOutcome) noexcept
{
    outOutcome = {};
    if (uid == 0u || executor == nullptr)
    {
        return E_INVALIDARG;
    }
    if (! uidPlusAvailable)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT hr = executor(context, ImapDeleteCommand::AddDeletedFlag, uid);
    if (FAILED(hr))
    {
        return hr;
    }
    outOutcome.targetMarkedDeleted = true;

    hr = executor(context, ImapDeleteCommand::UidExpunge, uid);
    if (SUCCEEDED(hr))
    {
        outOutcome.targetMarkedDeleted = false;
        return S_OK;
    }

    const HRESULT expungeHr       = hr;
    outOutcome.rollbackAttempted  = true;
    outOutcome.rollbackHr         = executor(context, ImapDeleteCommand::RemoveDeletedFlag, uid);
    if (SUCCEEDED(outOutcome.rollbackHr))
    {
        outOutcome.targetMarkedDeleted = false;
    }
    return expungeHr;
}

[[nodiscard]] std::vector<ImapUidBatchRange> BuildImapUidBatchRanges(size_t uidCount, size_t maxBatchSize)
{
    std::vector<ImapUidBatchRange> ranges;
    if (uidCount == 0 || maxBatchSize == 0)
    {
        return ranges;
    }

    ranges.reserve((uidCount + maxBatchSize - 1u) / maxBatchSize);
    for (size_t offset = 0; offset < uidCount; offset += maxBatchSize)
    {
        ranges.push_back(ImapUidBatchRange{.offset = offset, .count = std::min(maxBatchSize, uidCount - offset)});
    }

    return ranges;
}

[[nodiscard]] size_t ResolveImapSummaryRepairFetchBudget(size_t requestedUidCount) noexcept
{
    if (requestedUidCount == 0u)
    {
        return 0u;
    }

    constexpr size_t kPrimaryFetchChunkSize = 200u;
    constexpr size_t kMinRepairFetchBudget  = 4u;
    constexpr size_t kMaxRepairFetchBudget  = 64u;
    constexpr size_t kRepairFetchesPerChunk = 4u;

    const size_t chunkCount = (requestedUidCount + kPrimaryFetchChunkSize - 1u) / kPrimaryFetchChunkSize;
    const size_t scaledBudget =
        chunkCount > ((std::numeric_limits<size_t>::max)() / kRepairFetchesPerChunk) ? kMaxRepairFetchBudget : chunkCount * kRepairFetchesPerChunk;
    return (std::min)(kMaxRepairFetchBudget, (std::max)(kMinRepairFetchBudget, scaledBudget));
}
} // namespace FileSystemCurlInternal
