#pragma once

#include <cstddef>

namespace FolderViewSortPolicy
{
inline constexpr size_t kParallelSortThreshold = 20000u;

[[nodiscard]] constexpr bool ShouldUseParallelSort(size_t itemCount) noexcept
{
    return itemCount >= kParallelSortThreshold;
}
} // namespace FolderViewSortPolicy
