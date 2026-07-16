#include "DxUi.PointerInput.h"

#include <windowsx.h>

namespace RedSalamander::DxUi
{
namespace
{
[[nodiscard]] bool IsClientPointMouseMessage(UINT message) noexcept
{
    switch (message)
    {
        case WM_MOUSEMOVE:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP: return true;
        default: return false;
    }
}

void FillClientPoint(PointerInputEvent& event, HWND targetHwnd, LPARAM lParam) noexcept
{
    event.clientPointPx  = POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    event.hasClientPoint = true;

    POINT screenPoint = event.clientPointPx;
    if (targetHwnd && ClientToScreen(targetHwnd, &screenPoint) != FALSE)
    {
        event.screenPointPx  = screenPoint;
        event.hasScreenPoint = true;
    }
}

void FillScreenPoint(PointerInputEvent& event, HWND targetHwnd, LPARAM lParam) noexcept
{
    event.screenPointPx  = POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    event.hasScreenPoint = true;

    POINT clientPoint = event.screenPointPx;
    if (targetHwnd && ScreenToClient(targetHwnd, &clientPoint) != FALSE)
    {
        event.clientPointPx  = clientPoint;
        event.hasClientPoint = true;
    }
}

} // namespace

std::optional<PointerInputKind> PointerInputKindFromMessage(UINT message) noexcept
{
    switch (message)
    {
        case WM_MOUSEMOVE: return PointerInputKind::Move;
        case WM_MOUSELEAVE: return PointerInputKind::Leave;
        case WM_LBUTTONDOWN: return PointerInputKind::LeftDown;
        case WM_LBUTTONUP: return PointerInputKind::LeftUp;
        case WM_LBUTTONDBLCLK: return PointerInputKind::LeftDoubleClick;
        case WM_RBUTTONDOWN: return PointerInputKind::RightDown;
        case WM_RBUTTONUP: return PointerInputKind::RightUp;
        case WM_MOUSEWHEEL: return PointerInputKind::Wheel;
        default: return std::nullopt;
    }
}

std::optional<PointerInputEvent> TryBuildPointerInputEvent(HWND targetHwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    const std::optional<PointerInputKind> kind = PointerInputKindFromMessage(message);
    if (! kind.has_value())
    {
        return std::nullopt;
    }

    PointerInputEvent event{};
    event.kind        = kind.value();
    event.targetHwnd  = targetHwnd;
    event.rootHwnd    = targetHwnd ? GetAncestor(targetHwnd, GA_ROOT) : nullptr;
    event.captureHwnd = GetCapture();
    event.message     = message;
    event.wParam      = wParam;
    event.lParam      = lParam;
    event.messageTime = static_cast<DWORD>(GetMessageTime());
    event.wheelDelta  = (message == WM_MOUSEWHEEL) ? GET_WHEEL_DELTA_WPARAM(wParam) : 0;
    if (IsClientPointMouseMessage(message))
    {
        FillClientPoint(event, targetHwnd, lParam);
    }
    else if (message == WM_MOUSEWHEEL)
    {
        FillScreenPoint(event, targetHwnd, lParam);
    }

    return event;
}
} // namespace RedSalamander::DxUi
