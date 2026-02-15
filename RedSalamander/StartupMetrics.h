#pragma once

#include <cstdint>
#include <string_view>

namespace StartupMetrics
{
void Initialize() noexcept;

void MarkFirstWindowCreated(std::wstring_view windowId) noexcept;
void MarkFirstPaint(std::wstring_view windowId) noexcept;
void MarkInputReady(std::wstring_view windowId) noexcept;
void MarkFirstPanePopulated(std::wstring_view detail, uint64_t itemCount) noexcept;
} // namespace StartupMetrics
