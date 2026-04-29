#include "DxUi.Internal.h"

#include <climits>
#include <cwctype>

namespace RedSalamander::DxUi
{
wchar_t NormalizeTypeaheadChar(wchar_t ch) noexcept
{
    wchar_t mapped[2] = {};
    if (LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_UPPERCASE | LCMAP_LINGUISTIC_CASING, &ch, 1, mapped, 2, nullptr, nullptr, 0) > 0)
    {
        return mapped[0];
    }

    return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch)));
}

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
