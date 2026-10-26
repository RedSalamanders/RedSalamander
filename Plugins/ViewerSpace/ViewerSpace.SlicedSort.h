#pragma once

#include <algorithm>
#include <cstddef>
#include <vector>

namespace ViewerSpaceDetail
{
// UI preparation needs a resumable stable merge, unlike std::sort's indivisible callback.
// This consumer-local helper carries the UI work cursor; it does not change shared sorting policy.
template <typename T> struct SlicedSort final
{
    size_t width  = 1u;
    size_t begin  = 0u;
    size_t left   = 0u;
    size_t middle = 0u;
    size_t right  = 0u;
    size_t end    = 0u;
    size_t output = 0u;
    bool merging  = false;
    std::vector<T> scratch;

    template <typename Less> bool Step(std::vector<T>& values, Less less)
    {
        const size_t count = values.size();
        if (width >= count)
        {
            return true;
        }
        if (scratch.size() != count)
        {
            scratch.resize(count);
        }
        if (! merging)
        {
            if (begin >= count)
            {
                values.swap(scratch);
                width = width > count / 2u ? count : width * 2u;
                begin = 0u;
                return width >= count;
            }
            left    = begin;
            middle  = std::min(count, begin + width);
            right   = middle;
            end     = std::min(count, middle + width);
            output  = begin;
            merging = true;
        }
        if (left < middle && (right >= end || ! less(values[right], values[left])))
        {
            scratch[output++] = values[left++];
        }
        else
        {
            scratch[output++] = values[right++];
        }
        if (output == end)
        {
            begin   = end;
            merging = false;
        }
        return false;
    }
};
} // namespace ViewerSpaceDetail
