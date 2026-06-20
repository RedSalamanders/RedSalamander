#pragma once

#include "PlugInterfaces/FileSystem.h"

#include <cstddef>
#include <cstdint>
#include <regex>
#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace SearchTextHelpers
{
inline constexpr uint64_t kDefaultContentBytesPerFile = 64ull * 1024ull * 1024ull;
inline constexpr uint32_t kDefaultSnippetCharacters   = 160u;
inline constexpr size_t kDefaultReadChunkBytes        = 256u * 1024u;
inline constexpr size_t kDefaultLiteralChunkChars     = 32u * 1024u;
inline constexpr size_t kMaxRegexContentCharacters    = 5u * 1024u * 1024u;
inline constexpr size_t kMaxRegexPatternLength         = 1000u;
inline constexpr size_t kMaxRegexGroupDepth            = 20u;

enum class DecodedTextEncoding : uint32_t
{
    None = 0,
    Utf8,
    Utf16Le,
    Utf16Be,
    Utf32Le,
    Utf32Be,
    Ansi,
};

struct DecodedTextResult
{
    std::wstring text;
    DecodedTextEncoding encoding = DecodedTextEncoding::None;
    bool binary                  = false;
    bool utf8Validated           = false;
    bool usedFallbackCodePage    = false;
};

struct TextSearchPattern
{
    FileSystemSearchContentMode mode = FILESYSTEM_SEARCH_CONTENT_DISABLED;
    std::wstring_view pattern;
    const std::wregex* compiledRegex = nullptr;
    bool caseSensitive               = false;
    size_t literalChunkCharacters    = kDefaultLiteralChunkChars;
};

struct TextSearchResult
{
    bool matched         = false;
    uint64_t matchOffset = 0;
    uint32_t matchLength = 0;
    std::wstring previewText;
    bool binarySkipped           = false;
    bool overflowSkipped         = false;
    DecodedTextEncoding encoding = DecodedTextEncoding::None;
    bool usedFallbackCodePage    = false;
};

using CancelCheck = HRESULT(STDMETHODCALLTYPE*)(void* cookie) noexcept;

[[nodiscard]] bool LooksBinary(std::span<const std::byte> bytes) noexcept;

[[nodiscard]] bool TryDecodeSearchableText(std::span<const std::byte> bytes, UINT fallbackCodePage, DecodedTextResult& result) noexcept;

[[nodiscard]] std::wstring BuildSnippet(std::wstring_view text, size_t matchPosition, size_t matchLength, uint32_t maxSnippetCharacters) noexcept;

[[nodiscard]] bool FindLiteralWithChunkOverlap(
    std::wstring_view haystack, std::wstring_view needle, bool caseSensitive, size_t chunkCharacters, size_t& outPosition) noexcept;

[[nodiscard]] bool ValidateRegexPatternSafety(std::wstring_view pattern, std::wstring& outReason) noexcept;

[[nodiscard]] bool MatchDecodedText(
    const DecodedTextResult& decoded, const TextSearchPattern& pattern, uint32_t maxSnippetCharacters, bool wantSnippets, TextSearchResult& result) noexcept;

HRESULT ReadReaderBytes(
    IFileReader* reader, uint64_t maxBytes, CancelCheck cancelCheck, void* cancelCookie, size_t chunkBytes, std::vector<std::byte>& bytes) noexcept;

HRESULT SearchFileReaderText(IFileReader* reader,
                             const TextSearchPattern& pattern,
                             uint64_t maxBytes,
                             UINT fallbackCodePage,
                             uint32_t maxSnippetCharacters,
                             bool wantSnippets,
                             CancelCheck cancelCheck,
                             void* cancelCookie,
                             TextSearchResult& result) noexcept;
} // namespace SearchTextHelpers
