#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <type_traits>
#include <utility>

#include "Win32CallbackHelpers.h"

namespace RedSalamander::ViewerFileComboHost
{
inline constexpr int kStandaloneComboHeightDip        = 28;
inline constexpr int kStandaloneComboChromePaddingDip = 2;
inline constexpr int kStandaloneComboAccentHeightDip  = 1;
inline constexpr int kStandaloneComboAccentGapDip     = 1;

template <typename ViewerT>
[[nodiscard]] LRESULT DispatchFileComboHostWndProc(HWND hwnd,
                                                    UINT msg,
                                                    WPARAM wp,
                                                    LPARAM lp,
                                                    const wchar_t* stateProp,
                                                    const wchar_t* originalWndProcProp,
                                                    WNDPROC hookWndProc) noexcept
{
    auto* self = reinterpret_cast<ViewerT*>(::GetPropW(hwnd, stateProp));
    if (! self)
    {
        return Win32Callback::CallStoredWndProc(hwnd, originalWndProcProp, msg, wp, lp);
    }

    if (msg == WM_NCDESTROY)
    {
        const auto originalWndProc = Win32Callback::GetStoredWndProc(hwnd, originalWndProcProp);
        ::RemovePropW(hwnd, stateProp);
        Win32Callback::RestoreWndProcHook(hwnd, originalWndProcProp, hookWndProc);

        bool handled = false;
        static_cast<void>(self->HandleFileComboHostMessage(hwnd, msg, wp, lp, handled));

        return originalWndProc ? Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : ::DefWindowProcW(hwnd, msg, wp, lp);
    }

    const bool escapeKeyDown = msg == WM_KEYDOWN && wp == VK_ESCAPE;

    bool handled           = false;
    const LRESULT dxResult = self->HandleFileComboHostMessage(hwnd, msg, wp, lp, handled);
    if (handled)
    {
        if (escapeKeyDown)
        {
            self->FocusMainSurfaceFromFileCombo(::GetAncestor(hwnd, GA_ROOT));
        }
        return dxResult;
    }

    if (escapeKeyDown)
    {
        self->FocusMainSurfaceFromFileCombo(::GetAncestor(hwnd, GA_ROOT));
        return 0;
    }

    return Win32Callback::CallStoredWndProc(hwnd, originalWndProcProp, msg, wp, lp);
}

template <typename ViewerT>
[[nodiscard]] bool InstallFileComboHostWindow(
    HWND hwnd, ViewerT* self, const wchar_t* stateProp, const wchar_t* originalWndProcProp, WNDPROC hookWndProc) noexcept
{
    if (! hwnd || ! self || ! stateProp || ! originalWndProcProp || ! hookWndProc)
    {
        return false;
    }

    if (! Win32Callback::SetPropNoThrow(hwnd, stateProp, reinterpret_cast<HANDLE>(self)))
    {
        return false;
    }

    if (! Win32Callback::InstallWndProcHook(hwnd, originalWndProcProp, hookWndProc))
    {
        ::RemovePropW(hwnd, stateProp);
        return false;
    }

    return true;
}

inline void UnhookFileComboHostWindow(HWND hwnd, const wchar_t* stateProp, const wchar_t* originalWndProcProp, WNDPROC hookWndProc) noexcept
{
    if (! hwnd || ::IsWindow(hwnd) == FALSE)
    {
        return;
    }

    ::RemovePropW(hwnd, stateProp);
    Win32Callback::RestoreWndProcHook(hwnd, originalWndProcProp, hookWndProc);
}

template <typename WindowHostT, typename FocusMainSurfaceFn>
void ConfigureFileComboKeyboard(WindowHostT& host, FocusMainSurfaceFn&& focusMainSurface)
{
    using FocusCallback = std::decay_t<FocusMainSurfaceFn>;
    FocusCallback focusForTab(std::forward<FocusMainSurfaceFn>(focusMainSurface));
    FocusCallback focusForEscape(focusForTab);

    host.SetOnTabBoundary([focus = std::move(focusForTab)](bool) mutable noexcept
    {
        focus();
        return true;
    });
    host.SetOnEscape([focus = std::move(focusForEscape)]() mutable noexcept
    {
        focus();
        return true;
    });
}
} // namespace RedSalamander::ViewerFileComboHost
