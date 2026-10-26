#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace RedSalamander::ViewerMenuHost
{
enum class AttachOutcome
{
    Hosted,
    NativeFallback,
    DetachedFallback,
    ReplacedMenu,
    InvalidInput,
};

struct AttachResult
{
    AttachOutcome outcome = AttachOutcome::InvalidInput;
    DWORD fallbackError   = ERROR_SUCCESS;
    HMENU inputMenu       = nullptr;
    HMENU menuAfterAttach = nullptr;
    bool attempted        = false;
    bool attached         = false;
};

// Run on the owner UI thread, after CreateWindowEx returns and its temporary menu
// owner has handed the menu to the HWND. Input is either that borrowed menu or the
// existing detached RAII model. Never adopt a menu while it remains HWND-owned.
template <typename Host>
[[nodiscard]] AttachResult FinishAttach(
    HWND hwnd, HMENU menu, bool attached, Host& host, wil::unique_hmenu& detachedOwner, BOOL(WINAPI* restoreMenu)(HWND, HMENU) = ::SetMenu) noexcept
{
    if (! menu || ! hwnd || IsWindow(hwnd) == FALSE || IsMenu(menu) == FALSE || (detachedOwner && detachedOwner.get() != menu))
    {
        host.Detach();
        return {};
    }

    if (attached && GetMenu(hwnd) == nullptr)
    {
        if (! detachedOwner)
            detachedOwner.reset(menu);
        return {AttachOutcome::Hosted};
    }

    // Attach may fail during sync, after it detached the native menu. Discard the
    // partial host before restoring a functional native fallback.
    host.Detach();
    HMENU nativeMenu = GetMenu(hwnd);
    if (nativeMenu == menu)
    {
        if (detachedOwner.get() == menu)
            static_cast<void>(detachedOwner.release());
        return {AttachOutcome::NativeFallback};
    }

    if (! detachedOwner)
        detachedOwner.reset(menu);
    if (nativeMenu)
    {
        detachedOwner.reset();                // The discarded host no longer borrows the original.
        return {AttachOutcome::ReplacedMenu}; // Preserve the HWND-owned replacement.
    }

    // SetMenu can reenter window teardown after the HWND acquires the handle.
    // Give up the RAII claim before that call, then reclaim only a still-valid,
    // detached original. This also preserves an unexpected replacement menu.
    static_cast<void>(detachedOwner.release());
    const BOOL restored = restoreMenu(hwnd, menu);
    const DWORD error   = restored != FALSE ? ERROR_SUCCESS : GetLastError();
    if (IsMenu(menu) == FALSE)
        return {AttachOutcome::InvalidInput, error != ERROR_SUCCESS ? error : ERROR_INVALID_STATE};
    if (IsWindow(hwnd) == FALSE || GetMenu(hwnd) != menu)
    {
        detachedOwner.reset(menu);
        if (IsWindow(hwnd) != FALSE && GetMenu(hwnd) != nullptr)
        {
            detachedOwner.reset();
            return {AttachOutcome::ReplacedMenu, error != ERROR_SUCCESS ? error : ERROR_INVALID_STATE};
        }
        return {AttachOutcome::DetachedFallback, error != ERROR_SUCCESS ? error : ERROR_INVALID_STATE};
    }

    if (IsWindow(hwnd) != FALSE)
    {
        static_cast<void>(DrawMenuBar(hwnd));
        static_cast<void>(SetWindowPos(hwnd, nullptr, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE));
    }
    return {AttachOutcome::NativeFallback, error};
}

template <typename Host> [[nodiscard]] AttachResult Attach(HINSTANCE instance, HWND hwnd, Host& host, wil::unique_hmenu& detachedOwner) noexcept
{
    const HMENU menu            = detachedOwner ? detachedOwner.get() : GetMenu(hwnd);
    const bool attached         = menu && host.Attach(instance, hwnd, menu);
    const HMENU menuAfterAttach = GetMenu(hwnd);
    AttachResult result         = FinishAttach(hwnd, menu, attached, host, detachedOwner);
    result.inputMenu            = menu;
    result.menuAfterAttach      = menuAfterAttach;
    result.attempted            = menu != nullptr;
    result.attached             = attached;
    return result;
}
} // namespace RedSalamander::ViewerMenuHost
