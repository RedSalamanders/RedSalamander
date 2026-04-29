#pragma once

#include "Framework.h"

#include <d2d1.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#pragma warning(pop)

namespace D2DHdcPaint
{
[[nodiscard]] D2D1_COLOR_F Color(COLORREF color, float alpha = 1.0f) noexcept;

class Session final
{
public:
    Session() noexcept = default;
    ~Session() noexcept;

    Session(const Session&)            = delete;
    Session& operator=(const Session&) = delete;

    [[nodiscard]] bool Begin(HDC hdc, const RECT& boundsPx) noexcept;
    void End() noexcept;

    void FillRectangle(const RECT& rectPx, COLORREF color) noexcept;
    void FillRoundedRectangle(const RECT& rectPx, float radiusPx, COLORREF fill, COLORREF border, float strokeWidthPx = 1.0f) noexcept;
    void DrawLine(float x1, float y1, float x2, float y2, COLORREF color, float strokeWidthPx = 1.0f) noexcept;

private:
    [[nodiscard]] bool SetBrush(COLORREF color, float alpha = 1.0f) noexcept;
    [[nodiscard]] D2D1_RECT_F ToLocalRect(const RECT& rectPx) const noexcept;
    [[nodiscard]] D2D1_POINT_2F ToLocalPoint(float x, float y) const noexcept;

    wil::com_ptr<ID2D1DCRenderTarget> _target;
    wil::com_ptr<ID2D1SolidColorBrush> _brush;
    COLORREF _brushColor = CLR_INVALID;
    float _brushAlpha    = -1.0f;
    LONG _originX        = 0;
    LONG _originY        = 0;
    bool _drawing        = false;
};
} // namespace D2DHdcPaint
