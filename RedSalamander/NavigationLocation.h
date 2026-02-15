#pragma once

#include <algorithm>
#include <cstdint>
#include <cwctype>
#include <filesystem>
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

namespace NavigationLocation
{
struct Location
{
    std::wstring pluginShortId;
    std::wstring instanceContext;
    std::filesystem::path pluginPath;
};

enum class EmptyPathPolicy
{
    ReturnEmpty,
    Root,
};

enum class LeadingSlashPolicy
{
    Preserve,
    Ensure,
};

enum class TrailingSlashPolicy
{
    Preserve,
    Trim,
    Ensure,
};

[[nodiscard]] inline bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        const wint_t left  = std::towlower(static_cast<wint_t>(a[i]));
        const wint_t right = std::towlower(static_cast<wint_t>(b[i]));
        if (left != right)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] inline bool IsFilePluginShortId(std::wstring_view pluginShortId) noexcept
{
    return pluginShortId.empty() || EqualsNoCase(pluginShortId, L"file");
}

[[nodiscard]] inline bool LooksLikeWindowsDrivePath(std::wstring_view text) noexcept
{
    if (text.size() < 2u)
    {
        return false;
    }

    const wchar_t first = text[0];
    if (! ((first >= L'A' && first <= L'Z') || (first >= L'a' && first <= L'z')))
    {
        return false;
    }

    return text[1] == L':';
}

[[nodiscard]] inline bool LooksLikeExtendedPath(std::wstring_view text) noexcept;

[[nodiscard]] inline std::optional<wchar_t> TryGetWindowsDriveLetter(std::wstring_view text) noexcept
{
    if (LooksLikeWindowsDrivePath(text))
    {
        return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(text[0])));
    }

    if (LooksLikeExtendedPath(text) && text.size() >= 4u)
    {
        text.remove_prefix(4u);
        if (LooksLikeWindowsDrivePath(text))
        {
            return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(text[0])));
        }
    }

    return std::nullopt;
}

[[nodiscard]] inline std::optional<wchar_t> TryGetWindowsDriveLetter(const std::filesystem::path& path) noexcept
{
    return TryGetWindowsDriveLetter(std::wstring_view(path.native()));
}

[[nodiscard]] inline bool DriveMaskContainsLetter(uint32_t unitmask, wchar_t driveLetter) noexcept
{
    const wchar_t upper = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(driveLetter)));
    if (upper < L'A' || upper > L'Z')
    {
        return false;
    }

    const uint32_t bit = 1u << static_cast<uint32_t>(upper - L'A');
    return (unitmask & bit) != 0;
}

[[nodiscard]] inline bool LooksLikeUncPath(std::wstring_view text) noexcept
{
    return text.rfind(L"\\\\", 0) == 0 || text.rfind(L"//", 0) == 0;
}

[[nodiscard]] inline bool LooksLikeExtendedPath(std::wstring_view text) noexcept
{
    return text.rfind(L"\\\\?\\", 0) == 0 || text.rfind(L"\\\\.\\", 0) == 0;
}

[[nodiscard]] inline bool LooksLikeWindowsAbsolutePath(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return false;
    }

    if (LooksLikeExtendedPath(text))
    {
        return true;
    }

    if (LooksLikeUncPath(text))
    {
        return true;
    }

    return LooksLikeWindowsDrivePath(text);
}

[[nodiscard]] inline bool IsValidPluginShortId(std::wstring_view prefix) noexcept
{
    if (prefix.empty())
    {
        return false;
    }

    for (wchar_t ch : prefix)
    {
        if (std::iswalnum(static_cast<wint_t>(ch)) == 0)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] inline bool TryParsePluginPrefix(std::wstring_view text, std::wstring_view& outPrefix, std::wstring_view& outRemainder) noexcept
{
    outPrefix    = {};
    outRemainder = {};

    if (text.empty())
    {
        return false;
    }

    const size_t colon = text.find(L':');
    if (colon == std::wstring_view::npos || colon < 1u)
    {
        return false;
    }

    if (colon == 1u && std::iswalpha(static_cast<wint_t>(text[0])) != 0)
    {
        return false;
    }

    const size_t sep = text.find_first_of(L"\\/");
    if (sep != std::wstring_view::npos && sep < colon)
    {
        return false;
    }

    const std::wstring_view prefix = text.substr(0, colon);
    if (! IsValidPluginShortId(prefix))
    {
        return false;
    }

    outPrefix    = prefix;
    outRemainder = text.substr(colon + 1u);
    return true;
}

[[nodiscard]] inline std::wstring NormalizePluginPathText(std::wstring_view rawPath,
                                                          EmptyPathPolicy emptyPolicy        = EmptyPathPolicy::Root,
                                                          LeadingSlashPolicy leadingPolicy   = LeadingSlashPolicy::Ensure,
                                                          TrailingSlashPolicy trailingPolicy = TrailingSlashPolicy::Preserve) noexcept
{
    if (rawPath.empty())
    {
        if (emptyPolicy == EmptyPathPolicy::ReturnEmpty)
        {
            return {};
        }

        rawPath = L"/";
    }

    std::wstring pathText(rawPath);
    std::replace(pathText.begin(), pathText.end(), L'\\', L'/');
    if (leadingPolicy == LeadingSlashPolicy::Ensure && ! pathText.empty() && pathText.front() != L'/')
    {
        pathText.insert(pathText.begin(), L'/');
    }

    if (trailingPolicy == TrailingSlashPolicy::Trim)
    {
        while (pathText.size() > 1u && pathText.back() == L'/')
        {
            pathText.pop_back();
        }
    }
    else if (trailingPolicy == TrailingSlashPolicy::Ensure)
    {
        if (! pathText.empty() && pathText.back() != L'/')
        {
            pathText.push_back(L'/');
        }
    }

    return pathText;
}

[[nodiscard]] inline std::filesystem::path NormalizePluginPath(std::wstring_view rawPath) noexcept
{
    return std::filesystem::path(NormalizePluginPathText(rawPath, EmptyPathPolicy::Root, LeadingSlashPolicy::Ensure, TrailingSlashPolicy::Preserve));
}

[[nodiscard]] inline bool TrySplitPluginPathIntoFolderAndLeaf(std::wstring_view rawPath,
                                                              std::filesystem::path& outFolder,
                                                              std::wstring& outLeaf,
                                                              EmptyPathPolicy emptyPolicy = EmptyPathPolicy::Root) noexcept
{
    outFolder.clear();
    outLeaf.clear();

    const std::wstring normalized = NormalizePluginPathText(rawPath, emptyPolicy, LeadingSlashPolicy::Ensure, TrailingSlashPolicy::Preserve);
    if (normalized.empty())
    {
        return false;
    }

    const size_t lastSlash = normalized.find_last_of(L'/');
    if (lastSlash == std::wstring::npos)
    {
        return false;
    }

    outFolder = std::filesystem::path(normalized.substr(0, lastSlash + 1u));
    outLeaf   = normalized.substr(lastSlash + 1u);
    return true;
}

[[nodiscard]] inline bool TryPercentDecodeUtf8(std::wstring_view input, std::wstring& decoded) noexcept
{
    decoded.clear();
    if (input.empty())
    {
        return true;
    }

    if (input.size() > decoded.max_size())
    {
        return false;
    }
    decoded.reserve(input.size());

    std::string bytes;
    bytes.reserve(64);

    const auto flushBytes = [&]() noexcept
    {
        if (bytes.empty())
        {
            return true;
        }

        const int byteLen = static_cast<int>(std::min<size_t>(bytes.size(), static_cast<size_t>(std::numeric_limits<int>::max())));
        if (byteLen <= 0)
        {
            bytes.clear();
            return true;
        }

        const int needed = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), byteLen, nullptr, 0);
        if (needed <= 0)
        {
            // Be forgiving: fall back to a byte-wise widening.
            for (const char ch : bytes)
            {
                decoded.push_back(static_cast<wchar_t>(static_cast<unsigned char>(ch)));
            }
            bytes.clear();
            return true;
        }

        std::wstring temp;
        const size_t neededChars = static_cast<size_t>(needed);
        if (neededChars > temp.max_size())
        {
            return false;
        }
        temp.resize(neededChars);

        const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), byteLen, temp.data(), needed);
        if (written <= 0)
        {
            return false;
        }

        temp.resize(static_cast<size_t>(written));
        decoded.append(temp);
        bytes.clear();
        return true;
    };

    const auto isHex = [](wchar_t ch) noexcept -> int
    {
        if (ch >= L'0' && ch <= L'9')
        {
            return static_cast<int>(ch - L'0');
        }
        if (ch >= L'a' && ch <= L'f')
        {
            return 10 + static_cast<int>(ch - L'a');
        }
        if (ch >= L'A' && ch <= L'F')
        {
            return 10 + static_cast<int>(ch - L'A');
        }
        return -1;
    };

    for (size_t i = 0; i < input.size(); ++i)
    {
        const wchar_t ch = input[i];
        if (ch == L'%' && i + 2u < input.size())
        {
            const int hi = isHex(input[i + 1u]);
            const int lo = isHex(input[i + 2u]);
            if (hi >= 0 && lo >= 0)
            {
                bytes.push_back(static_cast<char>((hi << 4) | lo));
                i += 2u;
                continue;
            }
        }

        if (! flushBytes())
        {
            return false;
        }
        decoded.push_back(ch);
    }

    return flushBytes();
}

[[nodiscard]] inline bool TryParseFileUriRemainder(std::wstring_view uriRemainder, std::filesystem::path& outPath) noexcept
{
    outPath.clear();

    std::wstring_view remainderView = uriRemainder;
    if (remainderView.empty())
    {
        return false;
    }

    std::wstring_view authority;
    std::wstring_view pathPart;

    const auto startsWithTwoSlashes = [&](std::wstring_view value) noexcept
    { return value.size() >= 2u && (value[0] == L'/' || value[0] == L'\\') && (value[1] == L'/' || value[1] == L'\\'); };

    if (startsWithTwoSlashes(remainderView))
    {
        std::wstring_view after = remainderView.substr(2u);
        const size_t slashPos   = after.find_first_of(L"/\\");
        authority               = slashPos == std::wstring_view::npos ? after : after.substr(0, slashPos);
        pathPart                = slashPos == std::wstring_view::npos ? std::wstring_view{} : after.substr(slashPos);
    }
    else
    {
        pathPart = remainderView;
    }

    std::wstring decodedPath;
    if (! TryPercentDecodeUtf8(pathPart, decodedPath))
    {
        return false;
    }

    auto isAlpha = [](wchar_t ch) noexcept { return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z'); };

    const auto looksLikeDrive = [&](std::wstring_view maybeDrive) noexcept
    { return maybeDrive.size() == 2u && isAlpha(maybeDrive[0]) && maybeDrive[1] == L':'; };

    const bool authorityIsLocalhost = ! authority.empty() && EqualsNoCase(authority, L"localhost");

    // `file:////server/share/...` (authority empty, UNC encoded in path)
    if (authority.empty() && startsWithTwoSlashes(decodedPath))
    {
        std::wstring_view unc(decodedPath);
        while (startsWithTwoSlashes(unc))
        {
            unc.remove_prefix(2u);
        }
        const size_t shareSlash = unc.find_first_of(L"/\\");
        if (shareSlash == std::wstring_view::npos || shareSlash == 0u)
        {
            return false;
        }

        const std::wstring_view server = unc.substr(0, shareSlash);
        std::wstring_view shareAndRest = unc.substr(shareSlash + 1u);
        const size_t restSlash         = shareAndRest.find_first_of(L"/\\");
        const std::wstring_view share  = restSlash == std::wstring_view::npos ? shareAndRest : shareAndRest.substr(0, restSlash);
        const std::wstring_view rest   = restSlash == std::wstring_view::npos ? std::wstring_view{} : shareAndRest.substr(restSlash + 1u);
        if (share.empty())
        {
            return false;
        }

        std::wstring win;
        win.reserve(4u + server.size() + share.size() + rest.size());

        win.append(L"\\\\");
        win.append(server);
        win.push_back(L'\\');
        win.append(share);
        if (! rest.empty())
        {
            win.push_back(L'\\');
            win.append(rest);
        }

        std::replace(win.begin(), win.end(), L'/', L'\\');
        outPath = std::filesystem::path(std::move(win));
        return true;
    }

    // `file://C:/path` (nonstandard but common): authority is a drive like "C:".
    if (! authority.empty() && ! authorityIsLocalhost && looksLikeDrive(authority))
    {
        std::wstring drivePath;
        drivePath.reserve(authority.size() + decodedPath.size());
        drivePath.append(authority);
        drivePath.append(decodedPath);
        decodedPath.swap(drivePath);
        authority = {};
    }

    if (authority.empty() || authorityIsLocalhost)
    {
        std::wstring win = std::move(decodedPath);
        std::replace(win.begin(), win.end(), L'/', L'\\');

        if (win.size() >= 3u && win[0] == L'\\' && isAlpha(win[1]) && win[2] == L':')
        {
            win.erase(win.begin());
        }

        outPath = std::filesystem::path(std::move(win));
        return true;
    }

    // `file://server/share/path` -> `\\server\share\path`
    std::wstring_view shareAndRest(decodedPath);
    while (! shareAndRest.empty() && (shareAndRest.front() == L'/' || shareAndRest.front() == L'\\'))
    {
        shareAndRest.remove_prefix(1u);
    }
    if (shareAndRest.empty())
    {
        return false;
    }

    const size_t shareSlash       = shareAndRest.find_first_of(L"/\\");
    const std::wstring_view share = shareSlash == std::wstring_view::npos ? shareAndRest : shareAndRest.substr(0, shareSlash);
    const std::wstring_view rest  = shareSlash == std::wstring_view::npos ? std::wstring_view{} : shareAndRest.substr(shareSlash + 1u);
    if (share.empty())
    {
        return false;
    }

    std::wstring win;
    win.reserve(4u + authority.size() + share.size() + rest.size());

    win.append(L"\\\\");
    win.append(authority);
    win.push_back(L'\\');
    win.append(share);
    if (! rest.empty())
    {
        win.push_back(L'\\');
        win.append(rest);
    }

    std::replace(win.begin(), win.end(), L'/', L'\\');
    outPath = std::filesystem::path(std::move(win));
    return true;
}

[[nodiscard]] inline bool TryParseLocation(std::wstring_view text, Location& out) noexcept
{
    out.pluginShortId.clear();
    out.instanceContext.clear();
    out.pluginPath.clear();

    if (text.empty())
    {
        return false;
    }

    if (LooksLikeWindowsAbsolutePath(text))
    {
        out.pluginPath = std::filesystem::path(text);
        return true;
    }

    std::wstring_view prefix;
    std::wstring_view remainder;
    if (! TryParsePluginPrefix(text, prefix, remainder))
    {
        out.pluginPath = std::filesystem::path(text);
        return true;
    }

    if (IsFilePluginShortId(prefix))
    {
        std::filesystem::path filePath;
        if (TryParseFileUriRemainder(remainder, filePath))
        {
            out.pluginPath = std::move(filePath);
            return true;
        }

        std::wstring normalized(remainder);
        std::replace(normalized.begin(), normalized.end(), L'/', L'\\');
        out.pluginPath = std::filesystem::path(std::move(normalized));
        return true;
    }

    out.pluginShortId.assign(prefix);

    const size_t bar = remainder.find(L'|');
    std::wstring_view pathPart;
    if (bar != std::wstring_view::npos)
    {
        out.instanceContext.assign(remainder.substr(0, bar));
        pathPart = remainder.substr(bar + 1u);
    }
    else
    {
        pathPart = remainder;
    }

    out.pluginPath = NormalizePluginPath(pathPart);
    return true;
}

[[nodiscard]] inline std::filesystem::path
FormatHistoryPath(std::wstring_view pluginShortId, std::wstring_view instanceContext, const std::filesystem::path& pluginPath) noexcept
{
    if (IsFilePluginShortId(pluginShortId))
    {
        return pluginPath;
    }

    const std::filesystem::path normalized = NormalizePluginPath(pluginPath.wstring());
    const std::wstring pathText            = normalized.wstring();

    std::wstring result;
    if (instanceContext.empty())
    {
        result.reserve(pluginShortId.size() + 1u + pathText.size());
        result.append(pluginShortId);
        result.push_back(L':');
        result.append(pathText);
        return std::filesystem::path(result);
    }

    result.reserve(pluginShortId.size() + 1u + instanceContext.size() + 1u + pathText.size());
    result.append(pluginShortId);
    result.push_back(L':');
    result.append(instanceContext);
    result.push_back(L'|');
    result.append(pathText);
    return std::filesystem::path(result);
}

[[nodiscard]] inline std::filesystem::path FormatEditPath(std::wstring_view pluginShortId, const std::filesystem::path& pluginPath) noexcept
{
    if (IsFilePluginShortId(pluginShortId))
    {
        return pluginPath;
    }

    const std::filesystem::path normalized = NormalizePluginPath(pluginPath.wstring());
    const std::wstring pathText            = normalized.wstring();

    std::wstring result;
    result.reserve(pluginShortId.size() + 1u + pathText.size());
    result.append(pluginShortId);
    result.push_back(L':');
    result.append(pathText);
    return std::filesystem::path(result);
}
} // namespace NavigationLocation
