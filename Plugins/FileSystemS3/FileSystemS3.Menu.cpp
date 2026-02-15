#include "FileSystemS3.Internal.h"

HRESULT STDMETHODCALLTYPE FileSystemS3::GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (items == nullptr || count == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    _menuEntries.clear();
    _menuEntryView.clear();

    MenuEntry header;
    header.flags = NAV_MENU_ITEM_FLAG_HEADER;
    header.label = _metaData.name ? _metaData.name : L"";
    _menuEntries.push_back(std::move(header));

    MenuEntry separator;
    separator.flags = NAV_MENU_ITEM_FLAG_SEPARATOR;
    _menuEntries.push_back(std::move(separator));

    MenuEntry root;
    root.label = L"/";
    root.path  = L"/";
    _menuEntries.push_back(std::move(root));

    _menuEntryView.reserve(_menuEntries.size());
    for (const auto& e : _menuEntries)
    {
        NavigationMenuItem item{};
        item.flags     = e.flags;
        item.label     = e.label.empty() ? nullptr : e.label.c_str();
        item.path      = e.path.empty() ? nullptr : e.path.c_str();
        item.iconPath  = e.iconPath.empty() ? nullptr : e.iconPath.c_str();
        item.commandId = e.commandId;
        _menuEntryView.push_back(item);
    }

    *items = _menuEntryView.empty() ? nullptr : _menuEntryView.data();
    *count = static_cast<unsigned int>(_menuEntryView.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::ExecuteMenuCommand([[maybe_unused]] unsigned int commandId) noexcept
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::SetCallback(INavigationMenuCallback* callback, void* cookie) noexcept
{
    std::lock_guard lock(_stateMutex);
    _navigationMenuCallback       = callback;
    _navigationMenuCallbackCookie = callback != nullptr ? cookie : nullptr;
    return S_OK;
}
