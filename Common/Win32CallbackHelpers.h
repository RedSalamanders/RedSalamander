#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <cstddef>
#include <vector>
#include <windows.h>

#include "LocalizationManager.h"

namespace RedSalamander::Win32Callback
{
template <typename Callback> inline auto InvokeC5039Suppressed(Callback&& callback) noexcept(noexcept(callback())) -> decltype(callback())
{
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback or callback data to extern "C" Win32 API under -EHc
    return callback();
#pragma warning(pop)
}

[[nodiscard]] inline BOOL SetPropNoThrow(HWND hwnd, LPCWSTR string, HANDLE data) noexcept
{
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback data to extern "C" Win32 API under -EHc
    return ::SetPropW(hwnd, string, data);
#pragma warning(pop)
}

[[nodiscard]] inline LONG_PTR SetWindowLongPtrNoThrow(HWND hwnd, int index, LONG_PTR newLong) noexcept
{
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback data to extern "C" Win32 API under -EHc
    return ::SetWindowLongPtrW(hwnd, index, newLong);
#pragma warning(pop)
}

[[nodiscard]] inline LRESULT CallWindowProcNoThrow(WNDPROC wndProc, HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
    return ::CallWindowProcW(wndProc, hwnd, msg, wp, lp);
#pragma warning(pop)
}

[[nodiscard]] inline INT_PTR DialogBoxParamNoThrow(HINSTANCE instance, LPCWSTR templateName, HWND owner, DLGPROC dialogProc, LPARAM initParam) noexcept
{
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
    return ::DialogBoxParamW(instance, templateName, owner, dialogProc, initParam);
#pragma warning(pop)
}

[[nodiscard]] inline HWND CreateDialogParamNoThrow(HINSTANCE instance, LPCWSTR templateName, HWND parent, DLGPROC dialogProc, LPARAM initParam) noexcept
{
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
    return ::CreateDialogParamW(instance, templateName, parent, dialogProc, initParam);
#pragma warning(pop)
}

struct DialogTemplateResource final
{
    HINSTANCE creationInstance = nullptr;
    std::vector<std::byte> bytes;
};

[[nodiscard]] inline bool LoadDialogTemplateResource(HINSTANCE instance, LPCWSTR templateName, DialogTemplateResource& result) noexcept
{
    result.creationInstance = nullptr;
    result.bytes.clear();

    const Localization::ResourceLookupResult found = Localization::FindLocalizedResourceHandle(instance, templateName, RT_DIALOG);
    if (! found)
    {
        return false;
    }

    // Template bytes may come from a satellite DLL, but dialog class lookup must use the embedded owner module.
    const HINSTANCE creationInstance = instance ? instance : ::GetModuleHandleW(nullptr);
    if (! creationInstance)
    {
        return false;
    }

    const DWORD size = ::SizeofResource(found.instance, found.resource);
    if (size == 0)
    {
        return false;
    }

    HGLOBAL loaded = ::LoadResource(found.instance, found.resource);
    if (! loaded)
    {
        return false;
    }

    const void* bytes = ::LockResource(loaded);
    if (! bytes)
    {
        return false;
    }

    const auto* first       = static_cast<const std::byte*>(bytes);
    result.creationInstance = creationInstance;
    result.bytes.assign(first, first + size);
    return true;
}

[[nodiscard]] inline INT_PTR DialogBoxParamResourceNoThrow(HINSTANCE instance, LPCWSTR templateName, HWND owner, DLGPROC dialogProc, LPARAM initParam) noexcept
{
    DialogTemplateResource dialogTemplate;
    if (! LoadDialogTemplateResource(instance, templateName, dialogTemplate))
    {
        return -1;
    }

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
    return ::DialogBoxIndirectParamW(
        dialogTemplate.creationInstance, reinterpret_cast<LPCDLGTEMPLATEW>(dialogTemplate.bytes.data()), owner, dialogProc, initParam);
#pragma warning(pop)
}

[[nodiscard]] inline HWND CreateDialogParamResourceNoThrow(HINSTANCE instance, LPCWSTR templateName, HWND parent, DLGPROC dialogProc, LPARAM initParam) noexcept
{
    DialogTemplateResource dialogTemplate;
    if (! LoadDialogTemplateResource(instance, templateName, dialogTemplate))
    {
        return nullptr;
    }

#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
    return ::CreateDialogIndirectParamW(
        dialogTemplate.creationInstance, reinterpret_cast<LPCDLGTEMPLATEW>(dialogTemplate.bytes.data()), parent, dialogProc, initParam);
#pragma warning(pop)
}

[[nodiscard]] inline WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* propName) noexcept
{
    return reinterpret_cast<WNDPROC>(::GetPropW(hwnd, propName));
}

[[nodiscard]] inline bool InstallWndProcHook(HWND hwnd, const wchar_t* propName, WNDPROC hookWndProc) noexcept
{
    if (! hwnd || ! propName || ! hookWndProc)
    {
        return false;
    }

    if (GetStoredWndProc(hwnd, propName))
    {
        return true;
    }

    ::SetLastError(ERROR_SUCCESS);
    const auto originalWndProc = reinterpret_cast<WNDPROC>(::GetWindowLongPtrW(hwnd, GWLP_WNDPROC));
    if (! originalWndProc && ::GetLastError() != ERROR_SUCCESS)
    {
        return false;
    }

    if (! SetPropNoThrow(hwnd, propName, reinterpret_cast<HANDLE>(originalWndProc)))
    {
        return false;
    }

    ::SetLastError(ERROR_SUCCESS);
    const auto previousWndProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(hookWndProc)));
    if (! previousWndProc && ::GetLastError() != ERROR_SUCCESS)
    {
        ::RemovePropW(hwnd, propName);
        return false;
    }

    return true;
}

inline void RestoreWndProcHook(HWND hwnd, const wchar_t* propName) noexcept
{
    if (! hwnd || ::IsWindow(hwnd) == FALSE)
    {
        return;
    }

    if (const auto originalWndProc = GetStoredWndProc(hwnd, propName))
    {
        ::RemovePropW(hwnd, propName);
        static_cast<void>(SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(originalWndProc)));
    }
}

inline void RestoreWndProcHook(HWND hwnd, const wchar_t* propName, WNDPROC currentWndProc) noexcept
{
    if (! hwnd || ! propName)
    {
        return;
    }

    const auto originalWndProc = GetStoredWndProc(hwnd, propName);
    if (! originalWndProc)
    {
        return;
    }

    const auto currentWndProcValue = reinterpret_cast<WNDPROC>(::GetWindowLongPtrW(hwnd, GWLP_WNDPROC));
    if (currentWndProcValue == currentWndProc)
    {
        static_cast<void>(SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(originalWndProc)));
    }

    ::RemovePropW(hwnd, propName);
}

[[nodiscard]] inline LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* propName, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (const auto originalWndProc = GetStoredWndProc(hwnd, propName))
    {
        return CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp);
    }

    return ::DefWindowProcW(hwnd, msg, wp, lp);
}
} // namespace RedSalamander::Win32Callback
