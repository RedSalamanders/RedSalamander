#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <cstddef>
#include <limits>
#include <optional>
#include <string_view>

namespace FolderViewIncrementalSearch
{
[[nodiscard]] inline bool EqualsNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    if (left.size() != right.size())
    {
        return false;
    }
    if (left.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }
    if (left.empty())
    {
        return true;
    }

    const int result = CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE);
    return result == CSTR_EQUAL;
}

[[nodiscard]] inline bool StartsWithNoCase(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (prefix.empty() || text.size() < prefix.size())
    {
        return false;
    }

    return EqualsNoCase(text.substr(0u, prefix.size()), prefix);
}

[[nodiscard]] inline std::optional<UINT32> FindContainsOffsetNoCase(std::wstring_view text, std::wstring_view query) noexcept
{
    if (query.empty() || text.size() < query.size())
    {
        return std::nullopt;
    }
    if (query.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return std::nullopt;
    }

    const size_t querySize         = query.size();
    const size_t lastStartPosition = text.size() - querySize;
    for (size_t startPosition = 0u; startPosition <= lastStartPosition; ++startPosition)
    {
        const std::wstring_view window(text.data() + startPosition, querySize);
        if (! EqualsNoCase(window, query))
        {
            continue;
        }

        if (startPosition > static_cast<size_t>((std::numeric_limits<UINT32>::max)()))
        {
            return std::nullopt;
        }

        return static_cast<UINT32>(startPosition);
    }

    return std::nullopt;
}

template <typename DisplayNameAt>
[[nodiscard]] std::optional<size_t> FindNextPrefixMatchIndex(
    size_t itemCount, size_t startIndex, bool forward, DisplayNameAt displayNameAt, std::wstring_view query) noexcept
{
    if (itemCount == 0u || query.empty())
    {
        return std::nullopt;
    }

    for (size_t offset = 0u; offset < itemCount; ++offset)
    {
        const size_t index = forward ? ((startIndex + offset) % itemCount) : ((startIndex + itemCount - offset) % itemCount);
        if (StartsWithNoCase(displayNameAt(index), query))
        {
            return index;
        }
    }

    return std::nullopt;
}
} // namespace FolderViewIncrementalSearch
