#pragma once

#include <cstdint>
#include <optional>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

namespace RedSalamander::DxUi
{
enum class PointerInputSource : uint8_t
{
    WindowProc,
    ModalLoopMessage,
    PopupWindowProc,
    ForwardedChild,
    DiagnosticOnly,
};

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

struct InputGeneration
{
    uint64_t value = 0;
};

struct PointerInputEvent
{
    PointerInputSource source = PointerInputSource::WindowProc;
    PointerInputKind kind     = PointerInputKind::Unknown;
    HWND targetHwnd           = nullptr;
    HWND rootHwnd             = nullptr;
    HWND captureHwnd          = nullptr;
    UINT message              = 0;
    WPARAM wParam             = 0;
    LPARAM lParam             = 0;
    DWORD messageTime         = 0;
    POINT clientPointPx{};
    POINT screenPointPx{};
    InputGeneration generation{};
    int wheelDelta     = 0;
    bool hasClientPoint = false;
    bool hasScreenPoint = false;
};

[[nodiscard]] std::optional<PointerInputKind> PointerInputKindFromMessage(UINT message) noexcept;
[[nodiscard]] std::optional<PointerInputEvent> TryBuildPointerInputEvent(
    HWND targetHwnd,
    UINT message,
    WPARAM wParam,
    LPARAM lParam,
    PointerInputSource source,
    InputGeneration generation = {}) noexcept;
[[nodiscard]] std::optional<PointerInputEvent> TryBuildPointerInputEventFromMsg(
    const MSG& message,
    PointerInputSource source,
    InputGeneration generation = {}) noexcept;
} // namespace RedSalamander::DxUi
