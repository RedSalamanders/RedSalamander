#pragma once

#include "Win32CallbackHelpers.h"

namespace RedSalamander::DxUi
{
[[nodiscard]] inline WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* propName) noexcept
{
    return RedSalamander::Win32Callback::GetStoredWndProc(hwnd, propName);
}

[[nodiscard]] inline bool InstallWndProcHook(HWND hwnd, const wchar_t* propName, WNDPROC hookWndProc) noexcept
{
    return RedSalamander::Win32Callback::InstallWndProcHook(hwnd, propName, hookWndProc);
}

inline void RestoreWndProcHook(HWND hwnd, const wchar_t* propName) noexcept
{
    RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, propName);
}

[[nodiscard]] inline LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* propName, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    return RedSalamander::Win32Callback::CallStoredWndProc(hwnd, propName, msg, wp, lp);
}
} // namespace RedSalamander::DxUi
