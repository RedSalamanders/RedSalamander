#include "SearchTextHelpers.h"

#include "Helpers.h"

#include <algorithm>
#include <array>
#include <limits>

namespace SearchTextHelpers
{
namespace
{
constexpr HRESULT kFileTooLargeHr = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);

[[nodiscard]] std::wstring FoldWideText(std::wstring_view text) noexcept
{
    return OrdinalString::FoldCaseInvariant(text);
}

struct RegexQuantifier final
{
    bool isQuantifier = false;
    bool unbounded    = false;
    bool zeroMinimum  = false;
    size_t end        = 0u;
};

struct RegexGroupState final
{
    size_t contentStart                  = 0u;
    bool containsUnboundedRepetition     = false;
    bool containsZeroMinimumRepetition   = false;
    bool hasOverlappingTopLevelAlternate = false;
};

[[nodiscard]] bool IsRegexMetaCharacter(const wchar_t ch) noexcept
{
    switch (ch)
    {
        case L'.':
        case L'^':
        case L'$':
        case L'(':
        case L')':
        case L'[':
        case L']':
        case L'{':
        case L'}':
        case L'*':
        case L'+':
        case L'?':
        case L'|': return true;
        default: return false;
    }
}

[[nodiscard]] RegexQuantifier ParseRegexQuantifier(std::wstring_view pattern, size_t index) noexcept
{
    RegexQuantifier result{};
    if (index >= pattern.size())
    {
        return result;
    }

    const wchar_t ch = pattern[index];
    if (ch == L'*')
    {
        result.isQuantifier = true;
        result.unbounded    = true;
        result.zeroMinimum  = true;
        result.end          = index;
        return result;
    }

    if (ch == L'+')
    {
        result.isQuantifier = true;
        result.unbounded    = true;
        result.end          = index;
        return result;
    }

    if (ch == L'?')
    {
        result.isQuantifier = true;
        result.zeroMinimum  = true;
        result.end          = index;
        return result;
    }

    if (ch != L'{')
    {
        return result;
    }

    size_t cursor = index + 1u;
    size_t minimum = 0u;
    bool hasMinimum = false;
    while (cursor < pattern.size() && pattern[cursor] >= L'0' && pattern[cursor] <= L'9')
    {
        hasMinimum = true;
        minimum    = minimum * 10u + static_cast<size_t>(pattern[cursor] - L'0');
        ++cursor;
    }

    if (! hasMinimum || cursor >= pattern.size())
    {
        return result;
    }

    if (pattern[cursor] == L'}')
    {
        result.isQuantifier = true;
        result.zeroMinimum  = minimum == 0u;
        result.end          = cursor;
        return result;
    }

    if (pattern[cursor] != L',')
    {
        return result;
    }

    ++cursor;
    if (cursor < pattern.size() && pattern[cursor] == L'}')
    {
        result.isQuantifier = true;
        result.unbounded    = true;
        result.zeroMinimum  = minimum == 0u;
        result.end          = cursor;
        return result;
    }

    bool hasMaximum = false;
    while (cursor < pattern.size() && pattern[cursor] >= L'0' && pattern[cursor] <= L'9')
    {
        hasMaximum = true;
        ++cursor;
    }

    if (! hasMaximum || cursor >= pattern.size() || pattern[cursor] != L'}')
    {
        return result;
    }

    result.isQuantifier = true;
    result.zeroMinimum  = minimum == 0u;
    result.end          = cursor;
    return result;
}

[[nodiscard]] std::wstring ExtractLeadingLiteralPrefix(std::wstring_view text) noexcept
{
    std::wstring prefix;
    bool escaped = false;
    bool inClass = false;

    for (const wchar_t ch : text)
    {
        if (escaped)
        {
            prefix.push_back(ch);
            escaped = false;
            continue;
        }

        if (ch == L'\\')
        {
            escaped = true;
            continue;
        }

        if (inClass)
        {
            if (ch == L']')
            {
                inClass = false;
            }
            break;
        }

        if (ch == L'[')
        {
            inClass = true;
            break;
        }

        if (IsRegexMetaCharacter(ch))
        {
            break;
        }

        prefix.push_back(ch);
    }

    return prefix;
}

void AppendTopLevelAlternativePrefix(std::vector<std::wstring>& prefixes, std::wstring_view pattern, size_t start, size_t end)
{
    if (start <= end && start < pattern.size())
    {
        prefixes.push_back(ExtractLeadingLiteralPrefix(pattern.substr(start, end - start)));
    }
    else
    {
        prefixes.emplace_back();
    }
}

[[nodiscard]] bool HasOverlappingTopLevelAlternatives(std::wstring_view pattern, size_t start, size_t end) noexcept
{
    if (start >= end || end > pattern.size())
    {
        return false;
    }

    std::vector<std::wstring> prefixes;
    size_t alternativeStart = start;
    size_t depth            = 0u;
    bool escaped            = false;
    bool inClass            = false;
    bool sawAlternative     = false;

    for (size_t i = start; i < end; ++i)
    {
        const wchar_t ch = pattern[i];
        if (escaped)
        {
            escaped = false;
            continue;
        }

        if (ch == L'\\')
        {
            escaped = true;
            continue;
        }

        if (inClass)
        {
            if (ch == L']')
            {
                inClass = false;
            }
            continue;
        }

        if (ch == L'[')
        {
            inClass = true;
            continue;
        }

        if (ch == L'(')
        {
            ++depth;
            continue;
        }

        if (ch == L')')
        {
            if (depth > 0u)
            {
                --depth;
            }
            continue;
        }

        if (ch == L'|' && depth == 0u)
        {
            sawAlternative = true;
            AppendTopLevelAlternativePrefix(prefixes, pattern, alternativeStart, i);
            alternativeStart = i + 1u;
        }
    }

    if (! sawAlternative)
    {
        return false;
    }

    AppendTopLevelAlternativePrefix(prefixes, pattern, alternativeStart, end);

    for (size_t left = 0u; left < prefixes.size(); ++left)
    {
        if (prefixes[left].empty())
        {
            return true;
        }

        for (size_t right = left + 1u; right < prefixes.size(); ++right)
        {
            if (prefixes[right].empty())
            {
                return true;
            }

            const std::wstring& shorter = prefixes[left].size() <= prefixes[right].size() ? prefixes[left] : prefixes[right];
            const std::wstring& longer  = prefixes[left].size() <= prefixes[right].size() ? prefixes[right] : prefixes[left];
            if (longer.rfind(shorter, 0u) == 0u)
            {
                return true;
            }
        }
    }

    return false;
}

[[nodiscard]] std::wstring DecodeMultiByte(UINT codePage, DWORD flags, const char* bytes, size_t size) noexcept
{
    if (bytes == nullptr || size == 0u || size > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int charCount = ::MultiByteToWideChar(codePage, flags, bytes, static_cast<int>(size), nullptr, 0);
    if (charCount <= 0)
    {
        return {};
    }

    std::wstring result;
    result.resize(static_cast<size_t>(charCount));
    const int converted = ::MultiByteToWideChar(codePage, flags, bytes, static_cast<int>(size), result.data(), charCount);
    if (converted != charCount)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::wstring DecodeUtf16(std::span<const std::byte> bytes, size_t offset, bool bigEndian) noexcept
{
    if (bytes.size() <= offset)
    {
        return {};
    }

    const size_t pairCount = (bytes.size() - offset) / 2u;
    std::wstring result;
    result.resize(pairCount);

    for (size_t i = 0; i < pairCount; ++i)
    {
        const uint16_t first  = std::to_integer<uint8_t>(bytes[offset + i * 2u]);
        const uint16_t second = std::to_integer<uint8_t>(bytes[offset + i * 2u + 1u]);
        const uint16_t value  = bigEndian ? static_cast<uint16_t>((first << 8u) | second) : static_cast<uint16_t>(first | (second << 8u));
        result[i]             = static_cast<wchar_t>(value);
    }

    return result;
}

void AppendUtf32CodePoint(std::wstring& text, uint32_t codePoint) noexcept
{
    if (codePoint > 0x10FFFFu)
    {
        codePoint = 0xFFFDu;
    }

    if (codePoint <= 0xFFFFu)
    {
        if (codePoint >= 0xD800u && codePoint <= 0xDFFFu)
        {
            codePoint = 0xFFFDu;
        }

        text.push_back(static_cast<wchar_t>(codePoint));
        return;
    }

    codePoint -= 0x10000u;
    const wchar_t high = static_cast<wchar_t>(0xD800u + ((codePoint >> 10u) & 0x3FFu));
    const wchar_t low  = static_cast<wchar_t>(0xDC00u + (codePoint & 0x3FFu));
    text.push_back(high);
    text.push_back(low);
}

[[nodiscard]] std::wstring DecodeUtf32(std::span<const std::byte> bytes, size_t offset, bool bigEndian) noexcept
{
    if (bytes.size() <= offset)
    {
        return {};
    }

    const size_t wordCount = (bytes.size() - offset) / 4u;
    std::wstring result;
    result.reserve(wordCount);

    for (size_t i = 0; i < wordCount; ++i)
    {
        const uint32_t b0 = std::to_integer<uint8_t>(bytes[offset + i * 4u + 0u]);
        const uint32_t b1 = std::to_integer<uint8_t>(bytes[offset + i * 4u + 1u]);
        const uint32_t b2 = std::to_integer<uint8_t>(bytes[offset + i * 4u + 2u]);
        const uint32_t b3 = std::to_integer<uint8_t>(bytes[offset + i * 4u + 3u]);

        uint32_t codePoint = 0u;
        if (bigEndian)
        {
            codePoint = (b0 << 24u) | (b1 << 16u) | (b2 << 8u) | b3;
        }
        else
        {
            codePoint = b0 | (b1 << 8u) | (b2 << 16u) | (b3 << 24u);
        }

        AppendUtf32CodePoint(result, codePoint);
    }

    return result;
}
} // namespace

bool LooksBinary(std::span<const std::byte> bytes) noexcept
{
    const size_t sampleBytes = (std::min)(bytes.size(), static_cast<size_t>(4096u));
    for (size_t i = 0; i < sampleBytes; ++i)
    {
        if (std::to_integer<uint8_t>(bytes[i]) == 0u)
        {
            return true;
        }
    }

    return false;
}

bool TryDecodeSearchableText(std::span<const std::byte> bytes, UINT fallbackCodePage, DecodedTextResult& result) noexcept
{
    result = {};

    if (bytes.empty())
    {
        return true;
    }

    if (bytes.size() >= 4u)
    {
        const uint8_t b0 = std::to_integer<uint8_t>(bytes[0]);
        const uint8_t b1 = std::to_integer<uint8_t>(bytes[1]);
        const uint8_t b2 = std::to_integer<uint8_t>(bytes[2]);
        const uint8_t b3 = std::to_integer<uint8_t>(bytes[3]);
        if (b0 == 0xFFu && b1 == 0xFEu && b2 == 0x00u && b3 == 0x00u)
        {
            result.text     = DecodeUtf32(bytes, 4u, false);
            result.encoding = DecodedTextEncoding::Utf32Le;
            return true;
        }
        if (b0 == 0x00u && b1 == 0x00u && b2 == 0xFEu && b3 == 0xFFu)
        {
            result.text     = DecodeUtf32(bytes, 4u, true);
            result.encoding = DecodedTextEncoding::Utf32Be;
            return true;
        }
    }

    if (bytes.size() >= 2u)
    {
        const uint8_t b0 = std::to_integer<uint8_t>(bytes[0]);
        const uint8_t b1 = std::to_integer<uint8_t>(bytes[1]);
        if (b0 == 0xFFu && b1 == 0xFEu)
        {
            result.text     = DecodeUtf16(bytes, 2u, false);
            result.encoding = DecodedTextEncoding::Utf16Le;
            return true;
        }
        if (b0 == 0xFEu && b1 == 0xFFu)
        {
            result.text     = DecodeUtf16(bytes, 2u, true);
            result.encoding = DecodedTextEncoding::Utf16Be;
            return true;
        }
    }

    if (bytes.size() >= 3u && std::to_integer<uint8_t>(bytes[0]) == 0xEFu && std::to_integer<uint8_t>(bytes[1]) == 0xBBu &&
        std::to_integer<uint8_t>(bytes[2]) == 0xBFu)
    {
        result.text          = DecodeMultiByte(CP_UTF8, MB_ERR_INVALID_CHARS, reinterpret_cast<const char*>(bytes.data() + 3u), bytes.size() - 3u);
        result.encoding      = DecodedTextEncoding::Utf8;
        result.utf8Validated = ! result.text.empty() || bytes.size() == 3u;
        return true;
    }

    if (LooksBinary(bytes))
    {
        result.binary = true;
        return false;
    }

    result.text = DecodeMultiByte(CP_UTF8, MB_ERR_INVALID_CHARS, reinterpret_cast<const char*>(bytes.data()), bytes.size());
    if (! result.text.empty())
    {
        result.encoding      = DecodedTextEncoding::Utf8;
        result.utf8Validated = true;
        return true;
    }

    const UINT resolvedFallback = fallbackCodePage != 0u ? fallbackCodePage : CP_ACP;
    result.text                 = DecodeMultiByte(resolvedFallback, 0u, reinterpret_cast<const char*>(bytes.data()), bytes.size());
    if (! result.text.empty())
    {
        result.encoding             = DecodedTextEncoding::Ansi;
        result.usedFallbackCodePage = true;
        return true;
    }

    return false;
}

[[nodiscard]] bool IsHighSurrogate(wchar_t ch) noexcept
{
    return ch >= static_cast<wchar_t>(0xD800) && ch <= static_cast<wchar_t>(0xDBFF);
}

[[nodiscard]] bool IsLowSurrogate(wchar_t ch) noexcept
{
    return ch >= static_cast<wchar_t>(0xDC00) && ch <= static_cast<wchar_t>(0xDFFF);
}

[[nodiscard]] size_t ExpandStartToCodePointBoundary(std::wstring_view text, size_t start) noexcept
{
    if (start > 0u && start < text.size() && IsLowSurrogate(text[start]) && IsHighSurrogate(text[start - 1u]))
    {
        return start - 1u;
    }

    return start;
}

[[nodiscard]] size_t ExpandEndToCodePointBoundary(std::wstring_view text, size_t end) noexcept
{
    if (end > 0u && end < text.size() && IsHighSurrogate(text[end - 1u]) && IsLowSurrogate(text[end]))
    {
        return end + 1u;
    }

    return end;
}

std::wstring BuildSnippet(std::wstring_view text, size_t matchPosition, size_t matchLength, uint32_t maxSnippetCharacters) noexcept
{
    const size_t maxChars = (std::max)(static_cast<size_t>(maxSnippetCharacters), static_cast<size_t>(matchLength == 0u ? 1u : matchLength));
    if (text.size() <= maxChars)
    {
        return std::wstring(text);
    }

    const size_t desiredPrefix = maxChars / 3u;
    size_t start               = matchPosition > desiredPrefix ? matchPosition - desiredPrefix : 0u;
    size_t end                 = (std::min)(text.size(), start + maxChars);
    if (end - start < maxChars && end == text.size() && end > maxChars)
    {
        start = end - maxChars;
    }

    start = ExpandStartToCodePointBoundary(text, start);
    end   = ExpandEndToCodePointBoundary(text, end);

    std::wstring snippet;
    if (start > 0u)
    {
        snippet.append(L"...");
    }
    snippet.append(text.substr(start, end - start));
    if (end < text.size())
    {
        snippet.append(L"...");
    }

    return snippet;
}

bool FindLiteralWithChunkOverlap(std::wstring_view haystack, std::wstring_view needle, bool caseSensitive, size_t chunkCharacters, size_t& outPosition) noexcept
{
    outPosition = std::wstring_view::npos;

    if (needle.empty())
    {
        outPosition = 0u;
        return true;
    }

    if (haystack.size() < needle.size())
    {
        return false;
    }

    if (chunkCharacters == 0u)
    {
        chunkCharacters = haystack.size();
    }

    const size_t chunkSize          = (std::max)(chunkCharacters, static_cast<size_t>(1u));
    const size_t overlap            = needle.size() > 1u ? needle.size() - 1u : 0u;
    const std::wstring foldedNeedle = caseSensitive ? std::wstring() : FoldWideText(needle);

    for (size_t start = 0u; start < haystack.size();)
    {
        const size_t windowLength      = (std::min)(haystack.size() - start, chunkSize + overlap);
        const std::wstring_view window = haystack.substr(start, windowLength);

        size_t localPos = std::wstring_view::npos;
        if (caseSensitive)
        {
            localPos = window.find(needle);
        }
        else
        {
            const std::wstring foldedWindow = FoldWideText(window);
            localPos                        = foldedWindow.find(foldedNeedle);
        }

        if (localPos != std::wstring_view::npos)
        {
            outPosition = start + localPos;
            return true;
        }

        if (start + windowLength >= haystack.size())
        {
            break;
        }

        start += chunkSize;
    }

    return false;
}

bool ValidateRegexPatternSafety(std::wstring_view pattern, std::wstring& outReason) noexcept
{
    outReason.clear();
    if (pattern.size() > kMaxRegexPatternLength)
    {
        outReason = L"Regex pattern exceeds maximum length.";
        return false;
    }

    std::array<RegexGroupState, kMaxRegexGroupDepth + 1u> groups{};
    size_t depth = 0u;
    bool inCharClass = false;
    bool lastWasEscape = false;
    bool prevWasOpenParen = false;
    bool lastWasGroupClose = false;
    RegexGroupState lastClosedGroup{};

    for (size_t i = 0u; i < pattern.size(); ++i)
    {
        const wchar_t ch = pattern[i];

        if (lastWasEscape)
        {
            lastWasEscape    = false;
            lastWasGroupClose = false;
            prevWasOpenParen = false;
            continue;
        }

        if (ch == L'\\')
        {
            lastWasEscape    = true;
            lastWasGroupClose = false;
            prevWasOpenParen = false;
            continue;
        }

        if (inCharClass)
        {
            if (ch == L']')
            {
                inCharClass = false;
            }
            continue;
        }

        if (ch == L'[')
        {
            inCharClass      = true;
            lastWasGroupClose = false;
            prevWasOpenParen = false;
            continue;
        }

        if (ch == L'(')
        {
            ++depth;
            if (depth > kMaxRegexGroupDepth)
            {
                outReason = L"Regex group nesting too deep.";
                return false;
            }

            size_t contentStart = i + 1u;
            if (contentStart + 1u < pattern.size() && pattern[contentStart] == L'?' &&
                (pattern[contentStart + 1u] == L':' || pattern[contentStart + 1u] == L'=' || pattern[contentStart + 1u] == L'!'))
            {
                contentStart += 2u;
            }

            groups[depth] = RegexGroupState{.contentStart = contentStart};
            lastWasGroupClose = false;
            prevWasOpenParen = true;
            continue;
        }

        if (ch == L')')
        {
            prevWasOpenParen = false;
            if (depth == 0u)
            {
                lastWasGroupClose = false;
                continue;
            }

            groups[depth].hasOverlappingTopLevelAlternate =
                groups[depth].hasOverlappingTopLevelAlternate || HasOverlappingTopLevelAlternatives(pattern, groups[depth].contentStart, i);
            lastClosedGroup = groups[depth];
            --depth;

            groups[depth].containsUnboundedRepetition =
                groups[depth].containsUnboundedRepetition || lastClosedGroup.containsUnboundedRepetition;
            groups[depth].containsZeroMinimumRepetition =
                groups[depth].containsZeroMinimumRepetition || lastClosedGroup.containsZeroMinimumRepetition;
            groups[depth].hasOverlappingTopLevelAlternate =
                groups[depth].hasOverlappingTopLevelAlternate || lastClosedGroup.hasOverlappingTopLevelAlternate;

            lastWasGroupClose = true;
            continue;
        }

        // ? immediately after ( is group syntax (?:, (?=, etc.), not a quantifier.
        if (ch == L'?' && prevWasOpenParen)
        {
            prevWasOpenParen = false;
            continue;
        }

        const RegexQuantifier quantifier = ParseRegexQuantifier(pattern, i);
        if (quantifier.isQuantifier)
        {
            if (quantifier.unbounded && lastWasGroupClose &&
                (lastClosedGroup.containsUnboundedRepetition || lastClosedGroup.containsZeroMinimumRepetition ||
                 lastClosedGroup.hasOverlappingTopLevelAlternate))
            {
                outReason = L"Nested or ambiguous repetition in regex pattern (potential ReDoS).";
                return false;
            }

            if (quantifier.unbounded)
            {
                groups[depth].containsUnboundedRepetition = true;
            }
            if (quantifier.zeroMinimum)
            {
                groups[depth].containsZeroMinimumRepetition = true;
            }

            lastWasGroupClose = false;
            prevWasOpenParen = false;
            i = quantifier.end;
            continue;
        }

        lastWasGroupClose = false;
        prevWasOpenParen = false;
    }

    return true;
}

bool MatchDecodedText(
    const DecodedTextResult& decoded, const TextSearchPattern& pattern, uint32_t maxSnippetCharacters, bool wantSnippets, TextSearchResult& result) noexcept
{
    result                      = {};
    result.encoding             = decoded.encoding;
    result.usedFallbackCodePage = decoded.usedFallbackCodePage;
    result.binarySkipped        = decoded.binary;

    size_t matchPosition = std::wstring_view::npos;
    size_t matchLength   = 0u;

    switch (pattern.mode)
    {
        case FILESYSTEM_SEARCH_CONTENT_DISABLED: result.matched = true; return true;

        case FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL:
            result.matched = FindLiteralWithChunkOverlap(decoded.text, pattern.pattern, pattern.caseSensitive, pattern.literalChunkCharacters, matchPosition);
            matchLength    = pattern.pattern.size();
            break;

        case FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX:
        {
            if (pattern.compiledRegex == nullptr)
            {
                return false;
            }

            if (decoded.text.size() > kMaxRegexContentCharacters)
            {
                result.overflowSkipped = true;
                return false;
            }

            // noexcept boundary: std::regex_search may throw regex_error on
            // implementation-defined complexity/stack limits.
            try
            {
                std::wsmatch match;
                result.matched = std::regex_search(decoded.text, match, *pattern.compiledRegex);
                if (result.matched)
                {
                    matchPosition = static_cast<size_t>(match.position());
                    matchLength   = static_cast<size_t>(match.length());
                }
            }
            catch (const std::regex_error&)
            {
                result.matched = false;
            }
            break;
        }
    }

    if (! result.matched)
    {
        return false;
    }

    if (matchPosition != std::wstring_view::npos)
    {
        result.matchOffset = static_cast<uint64_t>(matchPosition);
        result.matchLength = matchLength > static_cast<size_t>((std::numeric_limits<uint32_t>::max)()) ? (std::numeric_limits<uint32_t>::max)()
                                                                                                       : static_cast<uint32_t>(matchLength);
        if (wantSnippets)
        {
            result.previewText = BuildSnippet(decoded.text, matchPosition, matchLength, maxSnippetCharacters);
        }
    }

    return true;
}

HRESULT ReadReaderBytes(
    IFileReader* reader, uint64_t maxBytes, CancelCheck cancelCheck, void* cancelCookie, size_t chunkBytes, std::vector<std::byte>& bytes) noexcept
{
    if (reader == nullptr)
    {
        return E_POINTER;
    }

    uint64_t sizeBytes = 0u;
    HRESULT hr         = reader->GetSize(&sizeBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    if (sizeBytes > maxBytes)
    {
        return kFileTooLargeHr;
    }

    if (sizeBytes > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
    {
        return E_OUTOFMEMORY;
    }

    bytes.assign(static_cast<size_t>(sizeBytes), std::byte{0});

    uint64_t newPosition = 0u;
    hr                   = reader->Seek(0, FILE_BEGIN, &newPosition);
    if (FAILED(hr))
    {
        return hr;
    }

    const size_t resolvedChunkBytes = chunkBytes != 0u ? chunkBytes : kDefaultReadChunkBytes;
    size_t offset                   = 0u;
    while (offset < bytes.size())
    {
        if (cancelCheck != nullptr)
        {
            hr = cancelCheck(cancelCookie);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        const size_t remaining       = bytes.size() - offset;
        const unsigned long toRead   = static_cast<unsigned long>((std::min)(remaining, resolvedChunkBytes));
        unsigned long bytesReadChunk = 0u;
        hr                           = reader->Read(bytes.data() + offset, toRead, &bytesReadChunk);
        if (FAILED(hr))
        {
            return hr;
        }

        if (bytesReadChunk == 0u)
        {
            break;
        }

        offset += bytesReadChunk;
    }

    bytes.resize(offset);
    return S_OK;
}

HRESULT SearchFileReaderText(IFileReader* reader,
                             const TextSearchPattern& pattern,
                             uint64_t maxBytes,
                             UINT fallbackCodePage,
                             uint32_t maxSnippetCharacters,
                             bool wantSnippets,
                             CancelCheck cancelCheck,
                             void* cancelCookie,
                             TextSearchResult& result) noexcept
{
    result = {};

    std::vector<std::byte> bytes;
    HRESULT hr = ReadReaderBytes(reader, maxBytes, cancelCheck, cancelCookie, kDefaultReadChunkBytes, bytes);
    if (FAILED(hr))
    {
        return hr;
    }

    DecodedTextResult decoded{};
    if (! TryDecodeSearchableText(bytes, fallbackCodePage, decoded))
    {
        result.binarySkipped        = decoded.binary;
        result.encoding             = decoded.encoding;
        result.usedFallbackCodePage = decoded.usedFallbackCodePage;
        return S_OK;
    }

    if (pattern.mode == FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX && cancelCheck != nullptr)
    {
        hr = cancelCheck(cancelCookie);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    static_cast<void>(MatchDecodedText(decoded, pattern, maxSnippetCharacters, wantSnippets, result));

    if (pattern.mode == FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX && cancelCheck != nullptr)
    {
        hr = cancelCheck(cancelCookie);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return S_OK;
}
} // namespace SearchTextHelpers
