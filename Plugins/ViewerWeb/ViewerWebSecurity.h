#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>

namespace ViewerWebSecurity
{
inline constexpr uint32_t kDefaultMaxDocumentMiB = 32u;
inline constexpr uint32_t kMaximumDocumentMiB = 64u;
inline constexpr bool kDefaultAllowExternalNavigation = false;
inline constexpr uint64_t kGeneratedOutputFixedAllowanceBytes = 1ull * 1024ull * 1024ull;
inline constexpr uint64_t kMaximumGeneratedOutputBytes = 128ull * 1024ull * 1024ull;

inline constexpr std::wstring_view kInternalDocumentOrigin = L"https://viewer.redsalamander.invalid";
inline constexpr std::wstring_view kInternalDocumentFilter = L"https://viewer.redsalamander.invalid/*";

// Raw HTML is an untrusted document. The response is sandboxed, has no script or
// network capability, and can only request a user-activated top-level navigation;
// that request is still mediated by EvaluateNavigation below.
inline constexpr wchar_t kRawHtmlResponseHeaders[] =
    L"Content-Type: text/html\r\n"
    L"Content-Security-Policy: sandbox allow-top-navigation-by-user-activation; default-src 'none'; script-src 'none'; connect-src 'none'; "
    L"frame-src 'none'; child-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'; style-src 'unsafe-inline'; img-src data:\r\n"
    L"X-Content-Type-Options: nosniff\r\n"
    L"Cache-Control: no-store\r\n";

// Generated JSON/Markdown pages need their bundled inline renderer, but retain a
// closed network/resource policy and cannot create frames, forms, or objects.
inline constexpr wchar_t kGeneratedDocumentResponseHeaders[] =
    L"Content-Type: text/html; charset=utf-8\r\n"
    L"Content-Security-Policy: default-src 'none'; script-src 'unsafe-inline'; connect-src 'none'; frame-src 'none'; child-src 'none'; "
    L"object-src 'none'; base-uri 'none'; form-action 'none'; style-src 'unsafe-inline'; img-src data:\r\n"
    L"X-Content-Type-Options: nosniff\r\n"
    L"Cache-Control: no-store\r\n";

enum class DocumentRoute : uint8_t
{
    None,
    RawHtmlPrivateOrigin,
    GeneratedPrivateOrigin,
    StagedPdf,
};

enum class NavigationSurface : uint8_t
{
    TopLevel,
    Frame,
    NewWindow,
};

enum class NavigationAction : uint8_t
{
    AllowInViewer,
    Block,
    OpenExternal,
};

[[nodiscard]] constexpr wchar_t AsciiLower(wchar_t value) noexcept
{
    return value >= L'A' && value <= L'Z' ? static_cast<wchar_t>(value - L'A' + L'a') : value;
}

[[nodiscard]] constexpr bool EqualsNoCase(std::wstring_view left, std::wstring_view right) noexcept
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

[[nodiscard]] constexpr bool StartsWithNoCase(std::wstring_view value, std::wstring_view prefix) noexcept
{
    return value.size() >= prefix.size() && EqualsNoCase(value.substr(0u, prefix.size()), prefix);
}

[[nodiscard]] constexpr bool IsHttpOrHttps(std::wstring_view uri) noexcept
{
    return StartsWithNoCase(uri, L"http://") || StartsWithNoCase(uri, L"https://");
}

[[nodiscard]] constexpr NavigationAction EvaluateNavigation(std::wstring_view uri,
                                                            NavigationSurface surface,
                                                            bool userInitiated,
                                                            bool allowExternalNavigation,
                                                            std::wstring_view allowedTopLevelDocument) noexcept
{
    const bool exactDocument = ! allowedTopLevelDocument.empty() && EqualsNoCase(uri, allowedTopLevelDocument);
    const bool documentFragment = ! allowedTopLevelDocument.empty() && uri.size() > allowedTopLevelDocument.size() &&
                                  StartsWithNoCase(uri, allowedTopLevelDocument) && uri[allowedTopLevelDocument.size()] == L'#';
    if (surface == NavigationSurface::TopLevel && (exactDocument || documentFragment))
    {
        return NavigationAction::AllowInViewer;
    }

    if ((! allowedTopLevelDocument.empty() && StartsWithNoCase(uri, allowedTopLevelDocument)) || StartsWithNoCase(uri, kInternalDocumentOrigin))
    {
        return NavigationAction::Block;
    }

    if (! IsHttpOrHttps(uri))
    {
        return NavigationAction::Block;
    }

    if (surface != NavigationSurface::Frame && userInitiated && allowExternalNavigation)
    {
        return NavigationAction::OpenExternal;
    }

    return NavigationAction::Block;
}

[[nodiscard]] constexpr uint64_t GeneratedOutputLimit(uint64_t configuredInputLimitBytes) noexcept
{
    constexpr uint64_t kExpansionFactor = 2u;
    if (configuredInputLimitBytes > (kMaximumGeneratedOutputBytes - kGeneratedOutputFixedAllowanceBytes) / kExpansionFactor)
    {
        return kMaximumGeneratedOutputBytes;
    }
    return kGeneratedOutputFixedAllowanceBytes + configuredInputLimitBytes * kExpansionFactor;
}

[[nodiscard]] constexpr bool IsGeneratedOutputWithinLimit(uint64_t outputBytes, uint64_t configuredInputLimitBytes) noexcept
{
    return outputBytes <= GeneratedOutputLimit(configuredInputLimitBytes);
}

[[nodiscard]] constexpr bool TryAccumulateWithinLimit(uint64_t current, uint64_t increment, uint64_t limit, uint64_t& next) noexcept
{
    if (current > limit || increment > limit - current)
    {
        return false;
    }
    next = current + increment;
    return true;
}

[[nodiscard]] constexpr bool IsProviderReadCountValid(size_t bufferBytes, unsigned long returnedBytes) noexcept
{
    return static_cast<uint64_t>(returnedBytes) <= static_cast<uint64_t>(bufferBytes);
}

enum class NormalizeTextResult : uint8_t
{
    Ok,
    TooLarge,
    InvalidEncoding,
};

// Converts UTF-16 BOM input in small chunks and never materializes a full UTF-16
// copy beside the raw and UTF-8 buffers. UTF-8/no-BOM input is copied directly.
[[nodiscard]] inline NormalizeTextResult NormalizeTextUtf8Bounded(std::string_view bytes, size_t maxOutputBytes, std::string& output) noexcept
{
    output.clear();

    if (bytes.size() >= 3u && static_cast<uint8_t>(bytes[0]) == 0xEFu && static_cast<uint8_t>(bytes[1]) == 0xBBu &&
        static_cast<uint8_t>(bytes[2]) == 0xBFu)
    {
        bytes.remove_prefix(3u);
    }

    const bool utf16Le = bytes.size() >= 2u && static_cast<uint8_t>(bytes[0]) == 0xFFu && static_cast<uint8_t>(bytes[1]) == 0xFEu;
    const bool utf16Be = bytes.size() >= 2u && static_cast<uint8_t>(bytes[0]) == 0xFEu && static_cast<uint8_t>(bytes[1]) == 0xFFu;
    if (! utf16Le && ! utf16Be)
    {
        if (bytes.size() > maxOutputBytes)
        {
            return NormalizeTextResult::TooLarge;
        }
        output.assign(bytes.data(), bytes.size());
        return NormalizeTextResult::Ok;
    }

    bytes.remove_prefix(2u);
    if ((bytes.size() % 2u) != 0u)
    {
        return NormalizeTextResult::InvalidEncoding;
    }

    constexpr size_t kChunkCodeUnits = 4096u;
    std::array<wchar_t, kChunkCodeUnits> chunk{};
    size_t byteOffset = 0u;
    while (byteOffset < bytes.size())
    {
        size_t codeUnits = (std::min)(kChunkCodeUnits, (bytes.size() - byteOffset) / 2u);
        for (size_t index = 0u; index < codeUnits; ++index)
        {
            const uint8_t first = static_cast<uint8_t>(bytes[byteOffset + index * 2u]);
            const uint8_t second = static_cast<uint8_t>(bytes[byteOffset + index * 2u + 1u]);
            chunk[index] = utf16Le ? static_cast<wchar_t>(static_cast<uint16_t>(first) | (static_cast<uint16_t>(second) << 8u))
                                   : static_cast<wchar_t>((static_cast<uint16_t>(first) << 8u) | static_cast<uint16_t>(second));
        }

        const bool moreInput = byteOffset + codeUnits * 2u < bytes.size();
        if (moreInput && codeUnits > 0u && chunk[codeUnits - 1u] >= 0xD800 && chunk[codeUnits - 1u] <= 0xDBFF)
        {
            --codeUnits;
        }
        if (codeUnits == 0u)
        {
            return NormalizeTextResult::InvalidEncoding;
        }

        const int codeUnitCount = static_cast<int>(codeUnits);
        const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, chunk.data(), codeUnitCount, nullptr, 0, nullptr, nullptr);
        if (required <= 0)
        {
            return NormalizeTextResult::InvalidEncoding;
        }

        uint64_t nextSize = 0u;
        if (! TryAccumulateWithinLimit(static_cast<uint64_t>(output.size()), static_cast<uint64_t>(required), static_cast<uint64_t>(maxOutputBytes), nextSize) ||
            nextSize > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()))
        {
            return NormalizeTextResult::TooLarge;
        }

        const size_t oldSize = output.size();
        output.resize(static_cast<size_t>(nextSize));
        const int written = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, chunk.data(), codeUnitCount, output.data() + oldSize, required, nullptr, nullptr);
        if (written != required)
        {
            output.resize(oldSize);
            return NormalizeTextResult::InvalidEncoding;
        }

        byteOffset += codeUnits * 2u;
    }

    return NormalizeTextResult::Ok;
}

struct DebugSnapshot
{
    uint32_t sizeBytes = sizeof(DebugSnapshot);
    DocumentRoute route = DocumentRoute::None;
    BOOL scriptsEnabled = FALSE;
    BOOL privateOrigin = FALSE;
    BOOL stagedFileTracked = FALSE;
    BOOL navigationCompleted = FALSE;
    BOOL navigationSucceeded = FALSE;
    BOOL generatedOutputRejected = FALSE;
    uint64_t loadedSourceBytes = 0u;
    uint64_t pendingCleanupCount = 0u;
    uint64_t generatedOutputBytes = 0u;
    uint64_t generatedOutputLimit = 0u;
    uint64_t asyncLoadPostFailures = 0u;
    uint64_t asyncSavePostFailures = 0u;
    BOOL loadPostFailureTerminal = FALSE;
    BOOL saveInProgress = FALSE;
    std::array<wchar_t, 512u> allowedDocumentUrl{};
    std::array<wchar_t, 512u> webViewSourceUrl{};
};

enum class DebugControlAction : WPARAM
{
    FailNextAsyncLoadCompletionPost = 1u,
    FailNextAsyncSaveCompletionPost = 2u,
    SaveAsToPath = 3u,
};

struct DebugSaveAsRequest
{
    uint32_t sizeBytes = sizeof(DebugSaveAsRequest);
    const wchar_t* destinationPath = nullptr;
    uint32_t faultMask = 0u;
    HRESULT submissionHr = E_FAIL;
};

enum DebugSaveFault : uint32_t
{
    DebugSaveFaultNone = 0u,
    DebugSaveFaultWrite = 1u << 0u,
    DebugSaveFaultFlush = 1u << 1u,
    DebugSaveFaultCommit = 1u << 2u,
};

[[nodiscard]] inline UINT GetDebugSnapshotMessage() noexcept
{
    static const UINT message = RegisterWindowMessageW(L"RedSalamander.ViewerWeb.DebugSnapshot.1");
    return message;
}

[[nodiscard]] inline UINT GetDebugControlMessage() noexcept
{
    static const UINT message = RegisterWindowMessageW(L"RedSalamander.ViewerWeb.DebugControl.1");
    return message;
}
} // namespace ViewerWebSecurity
