#include "Framework.h"

#include "ThemedInputFrames.h"

#include <algorithm>
#include <array>

#include "ThemedControls.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include <commctrl.h>

namespace
{
[[nodiscard]] bool IsComboBoxWindow(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    std::array<wchar_t, 64> className{};
    const int len = GetClassNameW(hwnd, className.data(), static_cast<int>(className.size()));
    if (len <= 0)
    {
        return false;
    }

    return _wcsicmp(className.data(), L"ComboBox") == 0 || ThemedControls::IsModernComboBox(hwnd);
}

void TryApplyRoundedComboRegion(HWND combo, UINT dpi) noexcept
{
    if (! combo)
    {
        return;
    }

    if (! IsComboBoxWindow(combo))
    {
        return;
    }

    RECT rc{};
    GetClientRect(combo, &rc);
    const int width  = std::max(0l, rc.right - rc.left);
    const int height = std::max(0l, rc.bottom - rc.top);
    if (width <= 0 || height <= 0)
    {
        return;
    }

    const int inset  = 1;
    const int baseR  = ThemedControls::ScaleDip(dpi, 4);
    const int radius = std::max(1, baseR - 2);
    const int right  = std::max(inset + 1, width - inset);
    const int bottom = std::max(inset + 1, height - inset);

    wil::unique_hrgn rgn(CreateRoundRectRgn(inset, inset, right + 1, bottom + 1, radius, radius));
    if (! rgn)
    {
        return;
    }

    SetWindowRgn(combo, rgn.release(), TRUE);
}

[[nodiscard]] HWND FindScrollableAncestor(HWND hwnd) noexcept
{
    HWND target = hwnd ? GetParent(hwnd) : nullptr;
    while (target)
    {
        const LONG_PTR style = GetWindowLongPtrW(target, GWL_STYLE);
        if ((style & WS_VSCROLL) != 0)
        {
            return target;
        }
        target = GetParent(target);
    }
    return nullptr;
}
} // namespace

namespace ThemedInputFrames
{
void InstallFrame(HWND frame, HWND input, FrameStyle* style) noexcept
{
    if (! frame || ! input || ! style)
    {
        return;
    }

    SetWindowLongPtrW(frame, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(input));

#pragma warning(push)
#pragma warning(disable : 5039) // passing potentially-throwing callback to extern "C" Win32 API under -EHc
    SetWindowSubclass(frame, InputFrameSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(style));
    SetWindowSubclass(input, InputControlSubclassProc, 1u, reinterpret_cast<DWORD_PTR>(frame));
#pragma warning(pop)
}

void InvalidateComboBox(HWND combo) noexcept
{
    if (! combo)
    {
        return;
    }

    InvalidateRect(combo, nullptr, TRUE);

    COMBOBOXINFO cbi{};
    cbi.cbSize = sizeof(cbi);
    if (GetComboBoxInfo(combo, &cbi) && cbi.hwndItem)
    {
        InvalidateRect(cbi.hwndItem, nullptr, TRUE);
    }
}

LRESULT CALLBACK InputControlSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR /*subclassId*/, DWORD_PTR refData) noexcept
{
    HWND frame = reinterpret_cast<HWND>(refData);

    switch (msg)
    {
        case WM_LBUTTONDOWN:
        case WM_RBUTTONDOWN:
        case WM_MBUTTONDOWN:
            SetPropW(hwnd, L"FocusViaMouse", reinterpret_cast<HANDLE>(1));
            if (frame)
            {
                InvalidateRect(frame, nullptr, TRUE);
            }
            break;
        case WM_SIZE:
        {
            if (frame && IsComboBoxWindow(hwnd))
            {
                const UINT dpi = GetDpiForWindow(hwnd);
                TryApplyRoundedComboRegion(hwnd, dpi);
                InvalidateComboBox(hwnd);
            }
            break;
        }
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORLISTBOX:
        case WM_CTLCOLORBTN:
        {
            const HWND parent = GetParent(hwnd);
            if (parent)
            {
                return SendMessageW(parent, msg, wp, lp);
            }
            break;
        }
        case WM_MOUSEWHEEL:
        {
            HWND target = FindScrollableAncestor(hwnd);
            if (! target)
            {
                break;
            }

            if (SendMessageW(hwnd, CB_GETDROPPEDSTATE, 0, 0) != 0)
            {
                break;
            }

            SendMessageW(target, msg, wp, lp);
            return 0;
        }
        case WM_SETFOCUS:
            if (GetPropW(hwnd, L"FocusViaMouse"))
            {
                const SHORT tabState = GetAsyncKeyState(VK_TAB);
                if ((tabState & 0x8000) != 0)
                {
                    RemovePropW(hwnd, L"FocusViaMouse");
                }
            }
            if (frame)
            {
                InvalidateRect(frame, nullptr, TRUE);
            }
            InvalidateRect(hwnd, nullptr, TRUE);
            break;
        case WM_KILLFOCUS:
            RemovePropW(hwnd, L"FocusViaMouse");
            if (frame)
            {
                InvalidateRect(frame, nullptr, TRUE);
            }
            InvalidateRect(hwnd, nullptr, TRUE);
            break;
        case WM_ENABLE:
            if (frame)
            {
                InvalidateRect(frame, nullptr, TRUE);
            }
            InvalidateRect(hwnd, nullptr, TRUE);
            break;
        case WM_KEYDOWN:
            RemovePropW(hwnd, L"FocusViaMouse");
            if (frame)
            {
                InvalidateRect(frame, nullptr, TRUE);
            }
            break;
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

LRESULT CALLBACK InputFrameSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR /*subclassId*/, DWORD_PTR refData) noexcept
{
    auto* style = reinterpret_cast<FrameStyle*>(refData);
    if (! style || ! style->theme)
    {
        return DefSubclassProc(hwnd, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_ERASEBKGND: return 1;
        case WM_LBUTTONDOWN:
        {
            HWND input = reinterpret_cast<HWND>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
            if (input)
            {
                SetPropW(input, L"FocusViaMouse", reinterpret_cast<HANDLE>(1));
                SetFocus(input);
            }
            return 0;
        }
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (! hdc)
            {
                return 0;
            }

            RECT rc{};
            GetClientRect(hwnd, &rc);

            if (style->backdropBrush)
            {
                FillRect(hdc.get(), &rc, style->backdropBrush);
            }
            else
            {
                const COLORREF bg = style->theme->windowBackground;
                wil::unique_hbrush bgBrush(CreateSolidBrush(bg));
                if (bgBrush)
                {
                    FillRect(hdc.get(), &rc, bgBrush.get());
                }
            }

            const UINT dpi           = GetDpiForWindow(hwnd);
            const int cornerDiameter = ThemedControls::ScaleDip(dpi, 8);
            const int cornerInset    = std::max(1, cornerDiameter / 2);

            HWND input              = reinterpret_cast<HWND>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
            const bool dropDownOpen = input && SendMessageW(input, CB_GETDROPPEDSTATE, 0, 0) != 0;
            const bool hasFocus     = input && (GetFocus() == input || dropDownOpen);
            const bool enabled      = input ? (IsWindowEnabled(input) != FALSE) : true;

            const bool isCombo = IsComboBoxWindow(input);

            const COLORREF surface = ThemedControls::GetControlSurfaceColor(*style->theme);
            COLORREF border        = ThemedControls::BlendColor(surface, style->theme->menu.text, style->theme->dark ? 60 : 40, 255);
            if (isCombo && hasFocus && enabled && ! style->theme->highContrast)
            {
                border = ThemedControls::BlendColor(surface, style->theme->menu.text, style->theme->dark ? 110 : 80, 255);
            }

            COLORREF fill = enabled ? style->inputBackgroundColor : style->inputDisabledBackgroundColor;
            if (hasFocus && enabled && ! style->theme->highContrast)
            {
                fill = style->inputFocusedBackgroundColor;
            }

            wil::unique_hbrush brush(CreateSolidBrush(fill));
            wil::unique_hpen pen(CreatePen(PS_SOLID, 1, border));
            if (brush && pen)
            {
                [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc.get(), brush.get());
                [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc.get(), pen.get());

                RoundRect(hdc.get(), rc.left, rc.top, rc.right, rc.bottom, cornerDiameter, cornerDiameter);
            }

            if (hasFocus && enabled && ! style->theme->highContrast)
            {
                if (isCombo)
                {
                    RECT bar            = rc;
                    const int barInsetX = std::max(1, ThemedControls::ScaleDip(dpi, 8));
                    const int barInsetY = std::max(1, ThemedControls::ScaleDip(dpi, 6));
                    const int barWidth  = std::max(1, ThemedControls::ScaleDip(dpi, 3));
                    bar.left            = std::min(bar.right, bar.left + barInsetX);
                    bar.right           = std::min(bar.right, bar.left + barWidth);
                    bar.top             = std::min(bar.bottom, bar.top + barInsetY);
                    bar.bottom          = std::max(bar.top, bar.bottom - barInsetY);

                    wil::unique_hbrush accentBrush(CreateSolidBrush(style->theme->menu.selectionBg));
                    if (accentBrush)
                    {
                        [[maybe_unused]] auto oldBrush = wil::SelectObject(hdc.get(), accentBrush.get());
                        [[maybe_unused]] auto oldPen   = wil::SelectObject(hdc.get(), GetStockObject(NULL_PEN));
                        const int barRadius            = ThemedControls::ScaleDip(dpi, 4);
                        RoundRect(hdc.get(), bar.left, bar.top, bar.right, bar.bottom, barRadius, barRadius);
                    }
                }
                else
                {
                    const int underline = std::max(1, ThemedControls::ScaleDip(dpi, 1));
                    RECT line{};
                    line.left   = rc.left + cornerInset;
                    line.right  = rc.right - cornerInset;
                    line.top    = rc.bottom - underline;
                    line.bottom = rc.bottom;

                    wil::unique_hbrush accentBrush(CreateSolidBrush(style->theme->menu.selectionBg));
                    if (accentBrush)
                    {
                        FillRect(hdc.get(), &line, accentBrush.get());
                    }
                }
            }

            return 0;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}
} // namespace ThemedInputFrames
