#pragma once

#include <unknwn.h>
#include <wchar.h>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
#pragma warning(disable : 4820) // padding in data structure
enum NavigationMenuItemFlags : uint32_t
{
    NAV_MENU_ITEM_FLAG_NONE      = 0,
    NAV_MENU_ITEM_FLAG_SEPARATOR = 0x1,
    NAV_MENU_ITEM_FLAG_DISABLED  = 0x2,
    NAV_MENU_ITEM_FLAG_HEADER    = 0x4,
};

struct NavigationMenuItem
{
    NavigationMenuItemFlags flags;
    // Display label (UTF-16). nullptr/empty for separators.
    const wchar_t* label;
    // Navigation target path (UTF-16). nullptr when not applicable.
    const wchar_t* path;
    // Optional path used for icon resolution (UTF-16). nullptr when not applicable.
    const wchar_t* iconPath;
    // Optional command identifier; 0 when not applicable.
    // Notes:
    // - This value is plugin-defined and is passed back to ExecuteMenuCommand / ExecuteDriveMenuCommand unchanged.
    // - This is NOT the Win32 WM_COMMAND identifier; the host assigns its own temporary menu IDs.
    unsigned int commandId;
};
#pragma warning(pop)

// Host callback for plugin-driven navigation requests.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The host must call INavigationMenu::SetCallback(nullptr, nullptr) before releasing/unloading the plugin.
// - The cookie is provided by the host at registration time and must be passed back verbatim by the plugin.
interface __declspec(novtable) INavigationMenuCallback
{
    // Requests the host to navigate to `path` (plugin path for the active file system).
    virtual HRESULT STDMETHODCALLTYPE NavigationMenuRequestNavigate(const wchar_t* path, void* cookie) noexcept = 0;
};

// Plugin-provided navigation menu entries.
// Notes:
// - Returned pointers are owned by the plugin and remain valid until the next call
//   to the same method or until the object is released.
interface __declspec(uuid("a7c7d693-5ba9-4f4d-8e90-0a2d9d7e49e4")) __declspec(novtable) INavigationMenu : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE ExecuteMenuCommand(unsigned int commandId) noexcept                          = 0;
    virtual HRESULT STDMETHODCALLTYPE SetCallback(INavigationMenuCallback * callback, void* cookie) noexcept       = 0;
};
