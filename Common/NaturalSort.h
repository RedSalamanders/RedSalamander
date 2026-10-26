#pragma once

#include <array>
#include <cstring>
#include <string>
#include <string_view>

#include "Helpers.h"
#include <shlwapi.h>

#pragma comment(lib, "shlwapi.lib")

// Display ordering only. Never use natural equality for provider identity,
// collision detection, operation dependencies or persisted canonical ordering.
namespace NaturalSort
{
namespace Detail
{
class TerminatedView final
{
public:
    TerminatedView(const TerminatedView&)            = delete;
    TerminatedView& operator=(const TerminatedView&) = delete;

    explicit TerminatedView(std::wstring_view text) noexcept
    {
        if (text.size() < _buffer.size())
        {
            // Copy the proven byte range in bulk. Checked STL iterators add
            // per-character Debug work to every comparison of provider slices.
            if (! text.empty())
            {
                std::memcpy(_buffer.data(), text.data(), text.size() * sizeof(wchar_t));
            }
            _buffer[text.size()] = L'\0';
        }
        else
        {
            // Provider names and root-relative path slices need not be terminated.
            // Long inputs are copied in full; allocation failure is fatal, never a
            // comparator-dependent fallback that could violate strict weak ordering.
            _longText = wil::make_process_heap_string_failfast(text.data(), text.size());
        }
    }

    [[nodiscard]] PCWSTR Get() const noexcept
    {
        return _longText ? _longText.get() : _buffer.data();
    }

private:
    std::array<wchar_t, MAX_PATH> _buffer;
    wil::unique_process_heap_string _longText;
};
} // namespace Detail

// File/folder display text has no embedded NUL. Ordinary length-delimited names
// use bounded stack scratch; owned strings use the zero-copy overload below.
[[nodiscard]] inline int Compare(std::wstring_view left, std::wstring_view right) noexcept
{
    // Extension/path groups commonly repeat an identical key. Equality needs no
    // terminated copy, allocation or collation call, including empty provider slices.
    if (left == right)
    {
        return 0;
    }
    const Detail::TerminatedView leftText(left);
    const Detail::TerminatedView rightText(right);
    return StrCmpLogicalW(leftText.Get(), rightText.Get());
}

[[nodiscard]] inline int Compare(const std::wstring& left, const std::wstring& right) noexcept
{
    if (left == right)
    {
        return 0;
    }
    return StrCmpLogicalW(left.c_str(), right.c_str());
}

// Bounded menus need a deterministic retained subset. Stable result grids instead
// use Compare and retain their existing source/arrival order for natural ties.
[[nodiscard]] inline bool Less(std::wstring_view left, std::wstring_view right) noexcept
{
    const int comparison = Compare(left, right);
    return comparison != 0 ? comparison < 0 : left < right;
}

[[nodiscard]] inline bool Less(const std::wstring& left, const std::wstring& right) noexcept
{
    const int comparison = Compare(left, right);
    return comparison != 0 ? comparison < 0 : left < right;
}
} // namespace NaturalSort
