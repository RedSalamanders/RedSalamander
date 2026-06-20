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

#include "Helpers.h"

namespace FolderViewIncrementalSearch
{
[[nodiscard]] inline bool EqualsNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    return OrdinalString::EqualsFoldedInvariant(left, right);
}

[[nodiscard]] inline bool StartsWithNoCase(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (prefix.empty() || text.size() < prefix.size())
    {
        return false;
    }

    return OrdinalString::StartsWithFoldedInvariant(text, prefix);
}

[[nodiscard]] inline std::optional<UINT32> FindContainsOffsetNoCase(std::wstring_view text, std::wstring_view query) noexcept
{
    if (query.empty() || text.size() < query.size())
    {
        return std::nullopt;
    }
    size_t startPosition = std::wstring_view::npos;
    if (! OrdinalString::FindContainsFoldedInvariant(text, query, startPosition))
    {
        return std::nullopt;
    }

    if (startPosition > static_cast<size_t>((std::numeric_limits<UINT32>::max)()))
    {
        return std::nullopt;
    }

    return static_cast<UINT32>(startPosition);
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
