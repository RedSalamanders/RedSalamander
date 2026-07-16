#pragma once

#include <cstdint>
#include <optional>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

namespace RedSalamander::DxUi
{
enum class PointerInputKind : uint8_t
{
    Move,
    Leave,
    LeftDown,
    LeftUp,
    LeftDoubleClick,
    RightDown,
    RightUp,
    Wheel,
    Unknown,
};

struct PointerInputEvent
{
    PointerInputKind kind = PointerInputKind::Unknown;
    HWND targetHwnd       = nullptr;
    HWND rootHwnd         = nullptr;
    HWND captureHwnd      = nullptr;
    UINT message          = 0;
    WPARAM wParam         = 0;
    LPARAM lParam         = 0;
    DWORD messageTime     = 0;
    POINT clientPointPx{};
    POINT screenPointPx{};
    int wheelDelta      = 0;
    bool hasClientPoint = false;
    bool hasScreenPoint = false;
};

[[nodiscard]] std::optional<PointerInputKind> PointerInputKindFromMessage(UINT message) noexcept;
[[nodiscard]] std::optional<PointerInputEvent> TryBuildPointerInputEvent(HWND targetHwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
} // namespace RedSalamander::DxUi
