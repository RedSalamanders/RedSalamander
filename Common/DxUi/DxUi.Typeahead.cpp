#include "DxUi.Internal.h"

#include <climits>

namespace RedSalamander::DxUi
{
bool StartsWithInsensitive(std::wstring_view text, std::wstring_view prefix) noexcept
{
    if (prefix.size() > text.size())
    {
        return false;
    }

    if (prefix.size() > static_cast<size_t>(INT_MAX))
    {
        return false;
    }

    return CompareStringOrdinal(text.data(), static_cast<int>(prefix.size()), prefix.data(), static_cast<int>(prefix.size()), TRUE) == CSTR_EQUAL;
}
} // namespace RedSalamander::DxUi
