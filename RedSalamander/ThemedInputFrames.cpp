#include "Framework.h"

#include "ThemedInputFrames.h"

#include <algorithm>
#include <array>
#include <cwchar>
#include <cwctype>
#include <string>

#include "D2DHdcPaint.h"
#include "UiMetrics.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

namespace
{
constexpr wchar_t kInputFrameOriginalWndProcProp[]   = L"RS.ThemedInputFrameOriginalWndProc";
constexpr wchar_t kInputControlOriginalWndProcProp[] = L"RS.ThemedInputControlOriginalWndProc";
constexpr wchar_t kInputFrameStyleProp[]             = L"RS.ThemedInputFrameStyle";
constexpr wchar_t kInputControlFrameProp[]           = L"RS.ThemedInputControlFrame";
constexpr wchar_t kCtrlBackspaceCharProp[]           = L"Win32UiCtrlBackspaceChar";

[[nodiscard]] WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* propName) noexcept
{
    return RedSalamander::Win32Callback::GetStoredWndProc(hwnd, propName);
}

bool InstallWndProcHook(HWND hwnd, const wchar_t* propName, WNDPROC newProc) noexcept
{
    if (! hwnd || ! propName || ! newProc)
    {
        return false;
    }

    if (GetStoredWndProc(hwnd, propName))
    {
        return true;
    }

    const auto original = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(hwnd, GWLP_WNDPROC));
    if (! original)
    {
        return false;
    }

    if (! RedSalamander::Win32Callback::SetPropNoThrow(hwnd, propName, reinterpret_cast<HANDLE>(original)))
    {
        return false;
    }

    static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(newProc)));
    return true;
}

void RestoreWndProcHook(HWND hwnd, const wchar_t* propName) noexcept
{
    const auto original = GetStoredWndProc(hwnd, propName);
    if (! original)
    {
        return;
    }

    static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(original)));
    RemovePropW(hwnd, propName);
}

[[nodiscard]] LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* propName, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    const auto original = GetStoredWndProc(hwnd, propName);
    if (! original)
    {
        return DefWindowProcW(hwnd, msg, wp, lp);
    }
    return RedSalamander::Win32Callback::CallWindowProcNoThrow(original, hwnd, msg, wp, lp);
}

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

    return _wcsicmp(className.data(), L"ComboBox") == 0;
}

[[nodiscard]] bool IsEditWindow(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    std::array<wchar_t, 16> className{};
    const int length = GetClassNameW(hwnd, className.data(), static_cast<int>(className.size()));
    if (length <= 0)
    {
        return false;
    }

    return _wcsicmp(className.data(), L"Edit") == 0;
}

[[nodiscard]] bool IsWordCharacter(wchar_t ch) noexcept
{
    return std::iswalnum(static_cast<wint_t>(ch)) != 0 || ch == L'_';
}

[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

[[nodiscard]] bool HandleEditCtrlBackspaceKeyDown(HWND edit, WPARAM key) noexcept
{
    if (! edit)
    {
        return false;
    }

    RemovePropW(edit, kCtrlBackspaceCharProp);

    if (! IsEditWindow(edit) || key != VK_BACK)
    {
        return false;
    }

    const LONG_PTR style = GetWindowLongPtrW(edit, GWL_STYLE);
    if ((style & ES_READONLY) != 0)
    {
        return false;
    }

    const bool ctrlDown = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool altDown  = (GetKeyState(VK_MENU) & 0x8000) != 0;
    if (! ctrlDown || altDown)
    {
        return false;
    }

    DWORD selectionStart = 0;
    DWORD selectionEnd   = 0;
    SendMessageW(edit, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));

    if (selectionStart != selectionEnd)
    {
        SendMessageW(edit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
        SetPropW(edit, kCtrlBackspaceCharProp, reinterpret_cast<HANDLE>(1));
        return true;
    }

    const int length = GetWindowTextLengthW(edit);
    std::wstring text;
    text.resize(static_cast<size_t>(std::max(0, length)) + 1u);
    GetWindowTextW(edit, text.data(), static_cast<int>(text.size()));
    text.resize(wcsnlen(text.c_str(), text.size()));

    const size_t caret = std::min(static_cast<size_t>(selectionEnd), text.size());
    if (caret == 0u)
    {
        SetPropW(edit, kCtrlBackspaceCharProp, reinterpret_cast<HANDLE>(1));
        return true;
    }

    size_t eraseFrom = caret;
    while (eraseFrom > 0u && std::iswspace(static_cast<wint_t>(text[eraseFrom - 1u])) != 0)
    {
        --eraseFrom;
    }

    if (eraseFrom > 0u)
    {
        const wchar_t previous = text[eraseFrom - 1u];
        if (IsPathSeparator(previous))
        {
            while (eraseFrom > 0u && IsPathSeparator(text[eraseFrom - 1u]))
            {
                --eraseFrom;
            }
        }
        else if (IsWordCharacter(previous))
        {
            while (eraseFrom > 0u && IsWordCharacter(text[eraseFrom - 1u]))
            {
                --eraseFrom;
            }
        }
        else
        {
            while (eraseFrom > 0u)
            {
                const wchar_t current = text[eraseFrom - 1u];
                if (std::iswspace(static_cast<wint_t>(current)) != 0 || IsPathSeparator(current) || IsWordCharacter(current))
                {
                    break;
                }
                --eraseFrom;
            }
        }
    }

    if (eraseFrom == caret)
    {
        eraseFrom = caret > 0u ? (caret - 1u) : 0u;
    }

    SendMessageW(edit, EM_SETSEL, static_cast<WPARAM>(eraseFrom), static_cast<LPARAM>(caret));
    SendMessageW(edit, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
    SetPropW(edit, kCtrlBackspaceCharProp, reinterpret_cast<HANDLE>(1));
    return true;
}

[[nodiscard]] bool HandleEditCtrlBackspaceChar(HWND edit, WPARAM key) noexcept
{
    if (! IsEditWindow(edit) || key != 0x7Fu)
    {
        return false;
    }

    if (! GetPropW(edit, kCtrlBackspaceCharProp))
    {
        return false;
    }

    RemovePropW(edit, kCtrlBackspaceCharProp);
    return true;
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
    const int baseR  = UiMetrics::ScaleDip(dpi, 4);
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
LRESULT CALLBACK InputControlWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
LRESULT CALLBACK InputFrameWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

void InstallFrame(HWND frame, HWND input, FrameStyle* style) noexcept
{
    if (! frame || ! input || ! style)
    {
        return;
    }

    SetWindowLongPtrW(frame, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(input));
    SetPropW(frame, kInputFrameStyleProp, reinterpret_cast<HANDLE>(style));
    SetPropW(input, kInputControlFrameProp, reinterpret_cast<HANDLE>(frame));
    InstallWndProcHook(frame, kInputFrameOriginalWndProcProp, InputFrameWndProc);
    InstallWndProcHook(input, kInputControlOriginalWndProcProp, InputControlWndProc);
}

void InstallControl(HWND input, HWND frame) noexcept
{
    if (! input)
    {
        return;
    }

    if (frame)
    {
        SetPropW(input, kInputControlFrameProp, reinterpret_cast<HANDLE>(frame));
    }
    else
    {
        RemovePropW(input, kInputControlFrameProp);
    }

    InstallWndProcHook(input, kInputControlOriginalWndProcProp, InputControlWndProc);
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

LRESULT HandleInputControlMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, HWND frame, bool& handled) noexcept
{
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

            if (IsComboBoxWindow(hwnd) && SendMessageW(hwnd, CB_GETDROPPEDSTATE, 0, 0) != 0)
            {
                break;
            }

            handled = true;
            SendMessageW(target, msg, wp, lp);
            return 0;
        }
        case WM_CHAR:
            if (HandleEditCtrlBackspaceChar(hwnd, wp))
            {
                handled = true;
                return 0;
            }
            break;
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
            if (HandleEditCtrlBackspaceKeyDown(hwnd, wp))
            {
                handled = true;
                return 0;
            }
            break;
        case WM_NCDESTROY:
            RemovePropW(hwnd, L"FocusViaMouse");
            RemovePropW(hwnd, kInputControlFrameProp);
            RestoreWndProcHook(hwnd, kInputControlOriginalWndProcProp);
            break;
    }

    handled = false;
    return 0;
}

LRESULT HandleInputFrameMessage(HWND hwnd, UINT msg, [[maybe_unused]] WPARAM wp, [[maybe_unused]] LPARAM lp, FrameStyle* style, bool& handled) noexcept
{
    if (! style || ! style->theme)
    {
        handled = false;
        return 0;
    }

    switch (msg)
    {
        case WM_ERASEBKGND: handled = true; return 1;
        case WM_LBUTTONDOWN:
        {
            HWND input = reinterpret_cast<HWND>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
            if (input)
            {
                SetPropW(input, L"FocusViaMouse", reinterpret_cast<HANDLE>(1));
                SetFocus(input);
            }
            handled = true;
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

            // Prevent the frame from painting over the input control (frame and input are siblings and overlap).
            // This fixes cases where the input text appears and then disappears while typing due to frame repaints.
            if (HWND input = reinterpret_cast<HWND>(GetWindowLongPtrW(hwnd, GWLP_USERDATA)))
            {
                RECT inputRc{};
                if (GetWindowRect(input, &inputRc))
                {
                    MapWindowPoints(nullptr, hwnd, reinterpret_cast<POINT*>(&inputRc), 2);

                    wil::unique_hrgn inputRgn(CreateRectRgn(0, 0, 0, 0));
                    const int rgnType = inputRgn ? GetWindowRgn(input, inputRgn.get()) : ERROR;
                    if (rgnType != ERROR && rgnType != NULLREGION)
                    {
                        OffsetRgn(inputRgn.get(), inputRc.left, inputRc.top);
                        ExtSelectClipRgn(hdc.get(), inputRgn.get(), RGN_DIFF);
                    }
                    else
                    {
                        ExcludeClipRect(hdc.get(), inputRc.left, inputRc.top, inputRc.right, inputRc.bottom);
                    }
                }
            }

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
            const int cornerDiameter = UiMetrics::ScaleDip(dpi, 8);
            const int cornerInset    = std::max(1, cornerDiameter / 2);

            HWND input              = reinterpret_cast<HWND>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
            const bool dropDownOpen = input && SendMessageW(input, CB_GETDROPPEDSTATE, 0, 0) != 0;
            const bool hasFocus     = input && (GetFocus() == input || dropDownOpen);
            const bool enabled      = input ? (IsWindowEnabled(input) != FALSE) : true;

            const bool isCombo = IsComboBoxWindow(input);

            const COLORREF surface = UiMetrics::GetControlSurfaceColor(*style->theme);
            COLORREF border        = UiMetrics::BlendColorRefWeightedTruncate(surface, style->theme->menu.text, style->theme->dark ? 60 : 40, 255);
            if (isCombo && hasFocus && enabled && ! style->theme->highContrast)
            {
                border = UiMetrics::BlendColorRefWeightedTruncate(surface, style->theme->menu.text, style->theme->dark ? 110 : 80, 255);
            }

            COLORREF fill = enabled ? style->inputBackgroundColor : style->inputDisabledBackgroundColor;
            if (hasFocus && enabled && ! style->theme->highContrast)
            {
                fill = style->inputFocusedBackgroundColor;
            }

            D2DHdcPaint::Session framePaint;
            const bool framePaintReady = framePaint.Begin(hdc.get(), rc);
            if (framePaintReady)
            {
                framePaint.FillRoundedRectangle(rc, static_cast<float>(cornerDiameter), fill, border);
            }

            if (hasFocus && enabled && ! style->theme->highContrast)
            {
                if (isCombo)
                {
                    RECT bar            = rc;
                    const int barInsetX = std::max(1, UiMetrics::ScaleDip(dpi, 8));
                    const int barInsetY = std::max(1, UiMetrics::ScaleDip(dpi, 6));
                    const int barWidth  = std::max(1, UiMetrics::ScaleDip(dpi, 3));
                    bar.left            = std::min(bar.right, bar.left + barInsetX);
                    bar.right           = std::min(bar.right, bar.left + barWidth);
                    bar.top             = std::min(bar.bottom, bar.top + barInsetY);
                    bar.bottom          = std::max(bar.top, bar.bottom - barInsetY);

                    if (framePaintReady)
                    {
                        const int barRadius = UiMetrics::ScaleDip(dpi, 4);
                        framePaint.FillRoundedRectangle(bar, static_cast<float>(barRadius), style->theme->menu.selectionBg, style->theme->menu.selectionBg);
                    }
                }
                else
                {
                    const int underline = std::max(1, UiMetrics::ScaleDip(dpi, 1));
                    RECT line{};
                    line.left   = rc.left + cornerInset;
                    line.right  = rc.right - cornerInset;
                    line.top    = rc.bottom - underline;
                    line.bottom = rc.bottom;

                    if (framePaintReady)
                    {
                        framePaint.FillRectangle(line, style->theme->menu.selectionBg);
                    }
                    else
                    {
                        wil::unique_hbrush accentBrush(CreateSolidBrush(style->theme->menu.selectionBg));
                        if (accentBrush)
                        {
                            FillRect(hdc.get(), &line, accentBrush.get());
                        }
                    }
                }
            }

            handled = true;
            return 0;
        }
        case WM_NCDESTROY:
            RemovePropW(hwnd, kInputFrameStyleProp);
            RestoreWndProcHook(hwnd, kInputFrameOriginalWndProcProp);
            break;
    }

    handled = false;
    return 0;
}

LRESULT CALLBACK InputControlWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    bool handled      = false;
    const auto result = HandleInputControlMessage(hwnd, msg, wp, lp, reinterpret_cast<HWND>(GetPropW(hwnd, kInputControlFrameProp)), handled);
    if (handled)
    {
        return result;
    }
    return CallStoredWndProc(hwnd, kInputControlOriginalWndProcProp, msg, wp, lp);
}

LRESULT CALLBACK InputFrameWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    bool handled      = false;
    const auto result = HandleInputFrameMessage(hwnd, msg, wp, lp, reinterpret_cast<FrameStyle*>(GetPropW(hwnd, kInputFrameStyleProp)), handled);
    if (handled)
    {
        return result;
    }
    return CallStoredWndProc(hwnd, kInputFrameOriginalWndProcProp, msg, wp, lp);
}
} // namespace ThemedInputFrames
