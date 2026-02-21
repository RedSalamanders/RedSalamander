#include "NavigationViewInternal.h"

#include <windowsx.h>

#include <commctrl.h>

#include "ConnectionSecrets.h"
#include "DirectoryInfoCache.h"
#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "HostServices.h"
#include "IconCache.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "SettingsStore.h"
#include "ThemedControls.h"
#include "resource.h"

ATOM NavigationView::RegisterEditSuggestPopupWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = EditSuggestPopupWndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kEditSuggestPopupClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

LRESULT CALLBACK NavigationView::EditSuggestPopupWndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    NavigationView* self = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self    = reinterpret_cast<NavigationView*>(cs->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    }
    else
    {
        self = reinterpret_cast<NavigationView*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (self)
    {
        return self->EditSuggestPopupWndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT NavigationView::OnCtlColorEdit(HDC hdc, HWND hwndControl)
{
    if (_pathEdit && hwndControl == _pathEdit.get())
    {
        SetTextColor(hdc, ColorToCOLORREF(_theme.text));
        SetBkColor(hdc, _theme.gdiBackground);
        return reinterpret_cast<LRESULT>(_backgroundBrush.get());
    }

    if (! _hWnd)
    {
        return 0;
    }

    return DefWindowProcW(_hWnd.get(), WM_CTLCOLOREDIT, reinterpret_cast<WPARAM>(hdc), reinterpret_cast<LPARAM>(hwndControl));
}

LRESULT NavigationView::OnEditSuggestResults(std::unique_ptr<EditSuggestResultsPayload> owned)
{
    if (! owned)
    {
        return 0;
    }
    if (owned->requestId != _editSuggestRequestId.load(std::memory_order_acquire))
    {
        return 0;
    }

    const size_t count = std::min(owned->displayItems.size(), owned->insertItems.size());

    std::vector<EditSuggestItem> merged;
    merged.reserve(kEditSuggestMaxItems);

    const size_t maxWithoutEllipsis = (owned->hasMore && kEditSuggestMaxItems > 0u) ? (kEditSuggestMaxItems - 1u) : kEditSuggestMaxItems;

    if (_editSuggestAdditionalRequestId == owned->requestId && ! _editSuggestAdditionalItems.empty())
    {
        for (auto& item : _editSuggestAdditionalItems)
        {
            if (merged.size() >= maxWithoutEllipsis)
            {
                break;
            }
            merged.push_back(std::move(item));
        }
        _editSuggestAdditionalItems.clear();
        _editSuggestAdditionalRequestId = 0;
    }

    for (size_t i = 0; i < count && merged.size() < maxWithoutEllipsis; ++i)
    {
        EditSuggestItem item{};
        item.display            = std::move(owned->displayItems[i]);
        item.insertText         = std::move(owned->insertItems[i]);
        item.directorySeparator = owned->directorySeparator;
        merged.push_back(std::move(item));
    }

    if (owned->hasMore && merged.size() < kEditSuggestMaxItems)
    {
        EditSuggestItem item{};
        item.display            = std::wstring(kEllipsisText);
        item.enabled            = false;
        item.directorySeparator = L'\0';
        merged.push_back(std::move(item));
    }

    _editSuggestHoveredIndex  = -1;
    _editSuggestSelectedIndex = -1;
    _editSuggestHighlightText = std::move(owned->highlightText);
    _editSuggestItems         = std::move(merged);

    if (_editSuggestItems.empty())
    {
        CloseEditSuggestPopup();
    }
    else
    {
        UpdateEditSuggestPopupWindow();
    }

    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupCreate()
{
    _editSuggestHoveredIndex  = -1;
    _editSuggestSelectedIndex = -1;
    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupNcDestroy()
{
    DiscardEditSuggestPopupD2DResources();
    _editSuggestPopup.release();
    _editSuggestPopupClientSize  = {0, 0};
    _editSuggestPopupRowHeightPx = 0;
    _editSuggestItems.clear();
    _editSuggestHoveredIndex  = -1;
    _editSuggestSelectedIndex = -1;
    _editSuggestHighlightText.clear();
    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupSize(HWND hwnd, UINT width, UINT height)
{
    _editSuggestPopupClientSize.cx = static_cast<LONG>(width);
    _editSuggestPopupClientSize.cy = static_cast<LONG>(height);

    if (_editSuggestPopupTarget)
    {
        _editSuggestPopupTarget->Resize(D2D1::SizeU(static_cast<UINT32>(_editSuggestPopupClientSize.cx), static_cast<UINT32>(_editSuggestPopupClientSize.cy)));
    }

    InvalidateRect(hwnd, nullptr, FALSE);
    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupMouseMove(HWND hwnd, POINT pt)
{
    TRACKMOUSEEVENT tme{};
    tme.cbSize    = sizeof(tme);
    tme.dwFlags   = TME_LEAVE;
    tme.hwndTrack = hwnd;
    TrackMouseEvent(&tme);

    const int itemHeight =
        std::max(1, _editSuggestPopupRowHeightPx > 0 ? _editSuggestPopupRowHeightPx : static_cast<int>(_sectionPathRect.bottom - _sectionPathRect.top));
    const int index = itemHeight > 0 ? (pt.y / itemHeight) : -1;

    int newHovered = -1;
    if (index >= 0 && static_cast<size_t>(index) < _editSuggestItems.size() && _editSuggestItems[static_cast<size_t>(index)].enabled)
    {
        newHovered = index;
    }

    if (newHovered != _editSuggestHoveredIndex)
    {
        _editSuggestHoveredIndex = newHovered;
        if (newHovered >= 0 && newHovered != _editSuggestSelectedIndex)
        {
            _editSuggestSelectedIndex = newHovered;
        }
        InvalidateRect(hwnd, nullptr, FALSE);
    }

    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupMouseLeave(HWND hwnd)
{
    if (_editSuggestHoveredIndex != -1)
    {
        _editSuggestHoveredIndex = -1;
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return 0;
}

LRESULT NavigationView::OnEditSuggestPopupLButtonDown(HWND /*hwnd*/, POINT pt)
{
    const int itemHeight =
        std::max(1, _editSuggestPopupRowHeightPx > 0 ? _editSuggestPopupRowHeightPx : static_cast<int>(_sectionPathRect.bottom - _sectionPathRect.top));
    const int index = itemHeight > 0 ? (pt.y / itemHeight) : -1;
    if (index >= 0 && static_cast<size_t>(index) < _editSuggestItems.size() && _editSuggestItems[static_cast<size_t>(index)].enabled)
    {
        ApplyEditSuggestIndex(static_cast<size_t>(index));
    }
    return 0;
}

LRESULT NavigationView::EditSuggestPopupWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_CREATE: return OnEditSuggestPopupCreate();
        case WM_NCDESTROY: return OnEditSuggestPopupNcDestroy();
        case WM_ERASEBKGND: return 1;
        case WM_MOUSEACTIVATE: return MA_NOACTIVATE;
        case WM_PAINT: RenderEditSuggestPopup(); return 0;
        case WM_SIZE: return OnEditSuggestPopupSize(hwnd, LOWORD(lp), HIWORD(lp));
        case WM_MOUSEMOVE: return OnEditSuggestPopupMouseMove(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
        case WM_MOUSELEAVE: return OnEditSuggestPopupMouseLeave(hwnd);
        case WM_LBUTTONDOWN: return OnEditSuggestPopupLButtonDown(hwnd, {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)});
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void NavigationView::EnterEditMode()
{
    if (_editMode || ! _currentPath)
        return;
    _editMode         = true;
    _renderMode       = RenderMode::Edit;
    _editCloseHovered = false;
    _editSuggestItems.clear();
    _editSuggestHighlightText.clear();
    CloseEditSuggestPopup();

    const std::filesystem::path& currentPath = _currentEditPath.has_value() ? _currentEditPath.value() : _currentPath.value();

    // Create or show Edit control overlay
    if (! _pathEdit)
    {
        int x      = _sectionPathRect.left;
        int y      = _sectionPathRect.top;
        int width  = _sectionPathRect.right - _sectionPathRect.left;
        int height = _sectionPathRect.bottom - _sectionPathRect.top;

        _pathEdit.reset(CreateWindowExW(0,
                                        L"EDIT",
                                        currentPath.c_str(),
                                        WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOHSCROLL | ES_LEFT,
                                        x,
                                        y,
                                        width,
                                        height,
                                        _hWnd.get(),
                                        reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_PATH_EDIT)),
                                        _hInstance,
                                        nullptr));

        SendMessageW(_pathEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_pathFont.get()), TRUE);
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
        SetWindowSubclass(_pathEdit.get(), NavigationView::EditSubclassProc, EDIT_SUBCLASS_ID, reinterpret_cast<DWORD_PTR>(this));
#pragma warning(pop)
    }
    else
    {
        SetWindowTextW(_pathEdit.get(), currentPath.c_str());
        ShowWindow(_pathEdit.get(), SW_SHOW);
    }

    if (_pathEdit)
    {
        const auto chrome = ComputeEditChromeRects(_sectionPathRect, _dpi);
        LayoutSingleLineEditInRect(_pathEdit.get(), chrome.editRect);
    }

    // Select all text
    SendMessageW(_pathEdit.get(), EM_SETSEL, 0, -1);
    SetFocus(_pathEdit.get());

    const std::wstring_view currentPathText = currentPath.native();
    const bool endsWithSeparator            = ! currentPathText.empty() && (currentPathText.back() == L'\\' || currentPathText.back() == L'/');
    if (endsWithSeparator)
    {
        UpdateEditSuggest();
    }

    UpdateHoverTimerState();
}

void NavigationView::ExitEditMode(bool accept)
{
    if (! _editMode)
        return;

    CloseEditSuggestPopup();
    static_cast<void>(_editSuggestRequestId.fetch_add(1, std::memory_order_acq_rel));
    {
        std::lock_guard lock(_editSuggestMutex);
        _editSuggestPendingQuery.reset();
    }
    _editSuggestMountedInstance.reset();

    _editMode = false;

    if (accept && _pathEdit)
    {
        wchar_t buffer[MAX_PATH];
        GetWindowTextW(_pathEdit.get(), buffer, MAX_PATH);

        if (ValidatePath(buffer))
        {
            std::filesystem::path newPath(buffer);

            const bool isFilePlugin = _pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file");
            const std::wstring_view typedText(buffer);
            if (! isFilePlugin && ! _currentInstanceContext.empty() && typedText.find(L'|') != std::wstring_view::npos)
            {
                std::wstring_view typedPrefix;
                std::wstring_view typedRemainder;
                if (! TryParsePluginPrefix(typedText, typedPrefix, typedRemainder))
                {
                    std::wstring canonical;
                    canonical.reserve(_pluginShortId.size() + 1u + typedText.size());
                    canonical.append(_pluginShortId);
                    canonical.push_back(L':');
                    canonical.append(typedText);
                    newPath = std::filesystem::path(std::move(canonical));
                }
            }
            RequestPathChange(newPath);
        }
        else
        {
            const std::wstring message = FormatStringResource(nullptr, IDS_FMT_INVALID_PATH, buffer);
            const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_INVALID_PATH);

            EDITBALLOONTIP tip{};
            tip.cbStruct = sizeof(tip);
            tip.pszTitle = title.c_str();
            tip.pszText  = message.c_str();
            tip.ttiIcon  = TTI_WARNING;
            SendMessageW(_pathEdit.get(), EM_SHOWBALLOONTIP, 0, reinterpret_cast<LPARAM>(&tip));
            // Keep edit mode active
            _editMode = true;
            UpdateHoverTimerState();
            return;
        }
    }

    if (_pathEdit)
    {
        ShowWindow(_pathEdit.get(), SW_HIDE);
    }
    _renderMode = RenderMode::Breadcrumb;
    InvalidateRect(_hWnd.get(), nullptr, FALSE);

    UpdateHoverTimerState();
}

void NavigationView::UpdateEditSuggest()
{
    if (! _editMode || ! _pathEdit)
    {
        _editSuggestItems.clear();
        _editSuggestHighlightText.clear();
        CloseEditSuggestPopup();
        return;
    }

    const int length = GetWindowTextLengthW(_pathEdit.get());
    std::wstring text;
    text.resize(static_cast<size_t>(std::max(0, length)) + 1u);
    GetWindowTextW(_pathEdit.get(), text.data(), static_cast<int>(text.size()));
    text.resize(wcsnlen(text.c_str(), text.size()));

    const uint64_t requestId        = _editSuggestRequestId.fetch_add(1, std::memory_order_acq_rel) + 1u;
    _editSuggestAdditionalRequestId = 0;
    _editSuggestAdditionalItems.clear();

    std::wstring normalizedInput = TrimWhitespace(text);
    if (normalizedInput.size() >= 2u && normalizedInput.front() == L'"' && normalizedInput.back() == L'"')
    {
        normalizedInput = normalizedInput.substr(1, normalizedInput.size() - 2u);
        normalizedInput = TrimWhitespace(normalizedInput);
    }

    const auto startsWithNoCase = [](std::wstring_view value, std::wstring_view prefix) noexcept
    {
        if (prefix.empty() || value.size() < prefix.size() || prefix.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
        {
            return false;
        }

        const int len = static_cast<int>(prefix.size());
        return CompareStringOrdinal(value.data(), len, prefix.data(), len, TRUE) == CSTR_EQUAL;
    };

    auto showStaticSuggestions = [&](std::vector<EditSuggestItem>&& items, std::wstring&& highlightText)
    {
        _editSuggestItems         = std::move(items);
        _editSuggestHoveredIndex  = -1;
        _editSuggestSelectedIndex = -1;
        _editSuggestHighlightText = std::move(highlightText);

        if (_editSuggestItems.empty())
        {
            CloseEditSuggestPopup();
        }
        else
        {
            UpdateEditSuggestPopupWindow();
        }
    };

    const auto buildProtocolAndDriveSuggestions = [&](std::wstring_view filterText) -> std::vector<EditSuggestItem>
    {
        std::vector<EditSuggestItem> items;
        if (filterText.empty())
        {
            return items;
        }

        // Host-level reserved prefix to route to Connection Manager profiles.
        if (filterText.front() == L'@' && startsWithNoCase(L"@conn:", filterText))
        {
            EditSuggestItem item{};
            item.display            = L"@conn:";
            item.insertText         = L"@conn:";
            item.directorySeparator = L'\0';
            items.push_back(std::move(item));
        }

        // File system plugins (shortId:)
        FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
        const auto& plugins                    = pluginManager.GetPlugins();
        for (const auto& entry : plugins)
        {
            if (entry.shortId.empty() || ! entry.loadable || entry.disabled)
            {
                continue;
            }

            if (! startsWithNoCase(entry.shortId, filterText))
            {
                continue;
            }

            EditSuggestItem item{};
            item.display            = entry.shortId + L":";
            item.insertText         = item.display;
            item.directorySeparator = L'\0';
            items.push_back(std::move(item));
        }

        // Drive roots (C:\)
        const DWORD drives = GetLogicalDrives();
        if (drives != 0)
        {
            const auto isAlpha = [](wchar_t ch) noexcept { return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z'); };

            const bool driveQuery =
                ! filterText.empty() && isAlpha(filterText.front()) && (filterText.size() == 1u || (filterText.size() == 2u && filterText[1] == L':'));
            if (driveQuery)
            {
                const wchar_t wanted = static_cast<wchar_t>(std::towupper(static_cast<wint_t>(filterText.front())));
                for (int i = 0; i < 26; ++i)
                {
                    if ((drives & (static_cast<DWORD>(1u) << i)) == 0)
                    {
                        continue;
                    }

                    const wchar_t driveLetter = static_cast<wchar_t>(L'A' + i);
                    if (driveLetter != wanted)
                    {
                        continue;
                    }

                    std::wstring root;
                    root.push_back(driveLetter);
                    root.append(L":\\");

                    EditSuggestItem item{};
                    item.display            = root;
                    item.insertText         = std::move(root);
                    item.directorySeparator = L'\\';
                    items.push_back(std::move(item));
                }
            }
        }

        std::sort(
            items.begin(), items.end(), [](const EditSuggestItem& a, const EditSuggestItem& b) { return _wcsicmp(a.display.c_str(), b.display.c_str()) < 0; });

        if (items.size() > kEditSuggestMaxItems)
        {
            items.resize(kEditSuggestMaxItems);
        }

        return items;
    };

    auto buildConnectionSuggestions = [&](std::wstring_view filterText,
                                          std::wstring_view insertPrefix,
                                          std::wstring_view filterPluginId,
                                          bool usePluginFilter,
                                          wchar_t directorySeparator) -> std::vector<EditSuggestItem>
    {
        std::vector<EditSuggestItem> items;
        if (! _settings)
        {
            return items;
        }

        struct Candidate
        {
            std::wstring sortKey;
            std::wstring display;
            std::wstring name;
        };

        std::vector<Candidate> candidates;

        const auto& plugins                 = FileSystemPluginManager::GetInstance().GetPlugins();
        const auto tryGetShortIdForPluginId = [&](std::wstring_view pluginId) noexcept -> std::wstring_view
        {
            for (const auto& entry : plugins)
            {
                if (! entry.id.empty() && EqualsNoCase(entry.id, pluginId) && ! entry.shortId.empty())
                {
                    return entry.shortId;
                }
            }
            return {};
        };

        const auto buildPreview = [&](const Common::Settings::ConnectionProfile& profile) -> std::wstring
        {
            const std::wstring_view shortId = tryGetShortIdForPluginId(profile.pluginId);
            if (shortId.empty() || profile.host.empty())
            {
                return {};
            }

            std::wstring host = profile.host;
            if (profile.port != 0u)
            {
                host = std::format(L"{}:{}", profile.host, profile.port);
            }

            if (! profile.userName.empty())
            {
                return std::format(L"{}://{}@{}", shortId, profile.userName, host);
            }

            return std::format(L"{}://{}", shortId, host);
        };

        const auto tryAddProfile = [&](std::wstring_view name, const Common::Settings::ConnectionProfile& profile, std::wstring_view labelOverride)
        {
            if (name.empty())
            {
                return;
            }

            if (usePluginFilter && ! filterPluginId.empty() && ! EqualsNoCase(profile.pluginId, filterPluginId))
            {
                return;
            }

            const std::wstring_view labelView = labelOverride.empty() ? std::wstring_view(profile.name) : labelOverride;

            if (! filterText.empty() && ! ContainsInsensitive(name, filterText) && ! ContainsInsensitive(labelView, filterText))
            {
                return;
            }

            Candidate c{};
            c.sortKey = std::wstring(name);
            c.name    = std::wstring(name);

            const std::wstring preview = buildPreview(profile);
            if (! labelOverride.empty())
            {
                c.display = preview.empty() ? std::format(L"{} — {}", name, labelOverride) : std::format(L"{} — {} — {}", name, labelOverride, preview);
            }
            else
            {
                c.display = preview.empty() ? std::wstring(name) : std::format(L"{} — {}", name, preview);
            }

            candidates.push_back(std::move(c));
        };

        // Quick Connect (session-only)
        {
            Common::Settings::ConnectionProfile quick{};
            const std::wstring_view preferredPluginId =
                usePluginFilter && ! filterPluginId.empty() ? filterPluginId : FileSystemPluginManager::GetInstance().GetActivePluginId();
            RedSalamander::Connections::EnsureQuickConnectProfile(preferredPluginId);
            RedSalamander::Connections::GetQuickConnectProfile(quick);

            const std::wstring quickLabel = LoadStringResource(nullptr, IDS_CONNECTIONS_QUICK_CONNECT);
            tryAddProfile(RedSalamander::Connections::kQuickConnectConnectionName, quick, quickLabel);
        }

        // Persisted profiles
        if (_settings->connections)
        {
            for (const auto& profile : _settings->connections->items)
            {
                if (profile.name.empty() || profile.pluginId.empty())
                {
                    continue;
                }
                tryAddProfile(profile.name, profile, {});
            }
        }

        std::sort(
            candidates.begin(), candidates.end(), [](const Candidate& a, const Candidate& b) { return _wcsicmp(a.sortKey.c_str(), b.sortKey.c_str()) < 0; });

        const size_t maxVisible = std::min(kEditSuggestMaxItems, static_cast<size_t>(10u));
        for (const auto& c : candidates)
        {
            if (items.size() >= maxVisible)
            {
                break;
            }
            EditSuggestItem item{};
            item.display            = c.display;
            item.insertText         = std::format(L"{}{}", insertPrefix, c.name);
            item.directorySeparator = directorySeparator;
            items.push_back(std::move(item));
        }

        if (candidates.size() > items.size() && items.size() < kEditSuggestMaxItems)
        {
            EditSuggestItem item{};
            item.display            = std::wstring(kEllipsisText);
            item.enabled            = false;
            item.directorySeparator = L'\0';
            items.push_back(std::move(item));
        }

        return items;
    };

    // `nav:` / `nav://` (Connection Manager routing)
    if (startsWithNoCase(normalizedInput, L"nav:"))
    {
        std::wstring rest = TrimWhitespace(std::wstring_view(normalizedInput).substr(4u));
        if (rest.size() >= 2u && rest[0] == L'/' && rest[1] == L'/')
        {
            rest.erase(0, 2);
        }

        std::wstring highlight = rest;
        std::wstring prefix    = startsWithNoCase(normalizedInput, L"nav://") ? L"nav://" : L"nav:";
        showStaticSuggestions(buildConnectionSuggestions(rest, prefix, {}, false, L'\0'), std::move(highlight));
        return;
    }

    // `@conn:` (Connection Manager routing alias)
    if (startsWithNoCase(normalizedInput, L"@conn:"))
    {
        std::wstring rest = TrimWhitespace(std::wstring_view(normalizedInput).substr(6u));
        showStaticSuggestions(buildConnectionSuggestions(rest, L"@conn:", {}, false, L'\0'), std::move(rest));
        return;
    }

    // Protocol-local Connection Manager prefix (ex: `ftp:/@conn:`)
    {
        std::wstring_view typedPrefix;
        std::wstring_view typedRemainder;
        if (TryParsePluginPrefix(normalizedInput, typedPrefix, typedRemainder))
        {
            const bool supportsConnections = EqualsNoCase(typedPrefix, L"ftp") || EqualsNoCase(typedPrefix, L"sftp") || EqualsNoCase(typedPrefix, L"scp") ||
                                             EqualsNoCase(typedPrefix, L"imap");

            if (supportsConnections && ! typedRemainder.empty() && typedRemainder.find(L'|') == std::wstring_view::npos)
            {
                std::wstring rem(typedRemainder);
                for (wchar_t& ch : rem)
                {
                    if (ch == L'\\')
                    {
                        ch = L'/';
                    }
                }

                if (! rem.empty() && rem.front() == L'@')
                {
                    rem.insert(rem.begin(), L'/');
                }

                std::wstring_view remView(rem);
                if (startsWithNoCase(remView, L"/@conn:"))
                {
                    std::wstring_view after(remView);
                    after.remove_prefix(7u); // "/@conn:"

                    const size_t nextSlash = after.find(L'/');
                    if (nextSlash == std::wstring_view::npos)
                    {
                        std::wstring insertPrefix = std::wstring(typedPrefix) + L":/@conn:";
                        std::wstring_view pluginIdFilter;
                        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
                        for (const auto& entry : plugins)
                        {
                            if (! entry.shortId.empty() && EqualsNoCase(entry.shortId, typedPrefix) && ! entry.id.empty())
                            {
                                pluginIdFilter = entry.id;
                                break;
                            }
                        }

                        std::wstring highlight = std::wstring(after);
                        showStaticSuggestions(buildConnectionSuggestions(after, insertPrefix, pluginIdFilter, true, L'/'), std::move(highlight));
                        return;
                    }
                }
                else if (startsWithNoCase(remView, L"/@") && remView.find(L'/') == 0u)
                {
                    // Complete `/@` to the reserved Connection Manager prefix.
                    std::wstring_view after(remView);
                    after.remove_prefix(2u); // "/@"

                    if (startsWithNoCase(L"conn:", after) || startsWithNoCase(L"conn", after) || after.empty())
                    {
                        EditSuggestItem item{};
                        item.display            = L"@conn:";
                        item.insertText         = std::wstring(typedPrefix) + L":/@conn:";
                        item.directorySeparator = L'\0';

                        std::vector<EditSuggestItem> items;
                        items.push_back(std::move(item));
                        showStaticSuggestions(std::move(items), std::wstring(after));
                        return;
                    }
                }
            }
        }
    }

    EditSuggestParseResult parseResult{};
    if (! TryParseEditSuggestQuery(normalizedInput, _pluginShortId, _currentEditPath, parseResult))
    {
        auto items = buildProtocolAndDriveSuggestions(normalizedInput);
        showStaticSuggestions(std::move(items), std::move(normalizedInput));
        return;
    }

    const auto isFileShortId = [](std::wstring_view shortId) noexcept { return shortId.empty() || EqualsNoCase(shortId, L"file"); };

    wil::com_ptr<IFileSystem> fileSystem = nullptr;
    std::shared_ptr<EditSuggestFileSystemInstance> keepAlive;

    const bool needsInstanceContext =
        parseResult.instanceContextSpecified && ! parseResult.instanceContext.empty() && ! isFileShortId(parseResult.enumerationShortId);

    if (! needsInstanceContext && _fileSystemPlugin &&
        (EqualsNoCase(parseResult.enumerationShortId, _pluginShortId) || (isFileShortId(parseResult.enumerationShortId) && isFileShortId(_pluginShortId))))
    {
        fileSystem = _fileSystemPlugin;
    }
    else if (needsInstanceContext && _fileSystemPlugin && EqualsNoCase(parseResult.enumerationShortId, _pluginShortId) &&
             EqualsNoCase(parseResult.instanceContext, _currentInstanceContext))
    {
        fileSystem = _fileSystemPlugin;
    }
    else
    {
        FileSystemPluginManager& plugins = FileSystemPluginManager::GetInstance();
        const auto& allPlugins           = plugins.GetPlugins();

        const FileSystemPluginManager::PluginEntry* entry = nullptr;
        for (const auto& candidate : allPlugins)
        {
            if (candidate.shortId.empty())
            {
                continue;
            }

            if (EqualsNoCase(candidate.shortId, parseResult.enumerationShortId))
            {
                entry = &candidate;
                break;
            }
        }

        if (entry != nullptr)
        {
            if (needsInstanceContext)
            {
                if (_editSuggestMountedInstance && EqualsNoCase(_editSuggestMountedInstance->pluginShortId, parseResult.enumerationShortId) &&
                    EqualsNoCase(_editSuggestMountedInstance->instanceContext, parseResult.instanceContext))
                {
                    keepAlive  = _editSuggestMountedInstance;
                    fileSystem = keepAlive->fileSystem;
                }
                else
                {
                    if (! entry->path.empty())
                    {
                        wil::unique_hmodule module(LoadLibraryExW(entry->path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));

                        if (module)
                        {
                            using CreateFactoryFunc   = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, void**);
                            using CreateFactoryExFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);

#pragma warning(push)
#pragma warning(disable : 4191) // C4191: unsafe conversion from FARPROC
                            const auto createFactory   = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(module.get(), "RedSalamanderCreate"));
                            const auto createFactoryEx = reinterpret_cast<CreateFactoryExFunc>(GetProcAddress(module.get(), "RedSalamanderCreateEx"));
#pragma warning(pop)

                            if (createFactory)
                            {
                                FactoryOptions options{};
                                options.debugLevel = DEBUG_LEVEL_NONE;

                                wil::com_ptr<IFileSystem> created;
                                HRESULT createHr = E_FAIL;
                                if (entry->factoryPluginId.empty())
                                {
                                    createHr = createFactory(__uuidof(IFileSystem), &options, GetHostServices(), created.put_void());
                                }
                                else if (createFactoryEx)
                                {
                                    createHr =
                                        createFactoryEx(__uuidof(IFileSystem), &options, GetHostServices(), entry->factoryPluginId.c_str(), created.put_void());
                                }
                                else
                                {
                                    createHr = HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
                                }
                                if (SUCCEEDED(createHr) && created)
                                {
                                    std::string configurationJsonUtf8;
                                    if (entry->informations)
                                    {
                                        const char* configuration = nullptr;
                                        if (SUCCEEDED(entry->informations->GetConfiguration(&configuration)) && configuration != nullptr &&
                                            configuration[0] != '\0')
                                        {
                                            configurationJsonUtf8 = configuration;
                                        }
                                    }

                                    if (! configurationJsonUtf8.empty())
                                    {
                                        wil::com_ptr<IInformations> createdInfos;
                                        if (SUCCEEDED(created->QueryInterface(__uuidof(IInformations), createdInfos.put_void())) && createdInfos)
                                        {
                                            static_cast<void>(createdInfos->SetConfiguration(configurationJsonUtf8.c_str()));
                                        }
                                    }

                                    wil::com_ptr<IFileSystemInitialize> initializer;
                                    const HRESULT initQi = created->QueryInterface(__uuidof(IFileSystemInitialize), initializer.put_void());
                                    if (SUCCEEDED(initQi) && initializer && SUCCEEDED(initializer->Initialize(parseResult.instanceContext.c_str(), nullptr)))
                                    {
                                        auto instance        = std::make_shared<EditSuggestFileSystemInstance>();
                                        instance->module     = std::move(module);
                                        instance->fileSystem = created;
                                        instance->pluginShortId.assign(parseResult.enumerationShortId);
                                        instance->instanceContext = parseResult.instanceContext;

                                        _editSuggestMountedInstance = instance;
                                        keepAlive                   = std::move(instance);
                                        fileSystem                  = created;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                fileSystem = entry->fileSystem;
            }
        }
    }

    if (! fileSystem)
    {
        _editSuggestItems.clear();
        _editSuggestHighlightText.clear();
        CloseEditSuggestPopup();
        return;
    }

    std::vector<EditSuggestItem> additionalItems;
    {
        const std::wstring_view view(normalizedInput);
        const bool hasSeparator = view.find_first_of(L"\\/") != std::wstring_view::npos;
        const size_t colonPos   = view.find(L':');

        const auto isAlpha = [](wchar_t ch) noexcept { return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z'); };

        const bool driveLike = view.size() <= 2u && ! view.empty() && isAlpha(view.front()) && (view.size() == 1u || view[1] == L':');

        if (! hasSeparator && (! view.empty()) && (view.front() == L'@' || colonPos == std::wstring_view::npos || driveLike))
        {
            additionalItems = buildProtocolAndDriveSuggestions(view);
        }
    }

    std::vector<std::wstring> names;
    bool usedCache = false;

    if (fileSystem)
    {
        auto borrowed =
            DirectoryInfoCache::GetInstance().BorrowDirectoryInfo(fileSystem.get(), parseResult.pluginFolder, DirectoryInfoCache::BorrowMode::CacheOnly);
        IFilesInformation* info = borrowed.Get();
        if (borrowed.Status() == S_OK && info)
        {
            usedCache = true;
            AppendMatchingDirectoryNamesFromFilesInformation(info, parseResult.filter, names);
        }
    }

    if (usedCache)
    {
        const bool hasMore = SortAndTrimEditSuggestNames(names);

        std::vector<std::wstring> displayItems;
        std::vector<std::wstring> insertItems;
        BuildEditSuggestLists(parseResult.displayFolder, names, parseResult.directorySeparator, displayItems, insertItems);

        const size_t count = std::min(displayItems.size(), insertItems.size());

        std::vector<EditSuggestItem> merged;
        merged.reserve(kEditSuggestMaxItems);

        const size_t maxWithoutEllipsis = (hasMore && kEditSuggestMaxItems > 0u) ? (kEditSuggestMaxItems - 1u) : kEditSuggestMaxItems;

        for (auto& item : additionalItems)
        {
            if (merged.size() >= maxWithoutEllipsis)
            {
                break;
            }
            merged.push_back(std::move(item));
        }

        for (size_t i = 0; i < count && merged.size() < maxWithoutEllipsis; ++i)
        {
            EditSuggestItem item{};
            item.display            = std::move(displayItems[i]);
            item.insertText         = std::move(insertItems[i]);
            item.directorySeparator = parseResult.directorySeparator;
            merged.push_back(std::move(item));
        }

        if (hasMore && merged.size() < kEditSuggestMaxItems)
        {
            EditSuggestItem item{};
            item.display            = std::wstring(kEllipsisText);
            item.enabled            = false;
            item.directorySeparator = L'\0';
            merged.push_back(std::move(item));
        }

        _editSuggestHoveredIndex  = -1;
        _editSuggestSelectedIndex = -1;
        _editSuggestHighlightText = parseResult.filter;
        _editSuggestItems         = std::move(merged);
        UpdateEditSuggestPopupWindow();
        return;
    }

    EnsureEditSuggestWorker();
    {
        std::lock_guard lock(_editSuggestMutex);
        EditSuggestQuery query{};
        query.requestId          = requestId;
        query.fileSystem         = fileSystem;
        query.displayFolder      = parseResult.displayFolder;
        query.pluginFolder       = parseResult.pluginFolder;
        query.prefix             = std::move(parseResult.filter);
        query.directorySeparator = parseResult.directorySeparator;
        query.keepAlive          = keepAlive;
        _editSuggestPendingQuery = std::move(query);
    }
    _editSuggestCv.notify_one();

    if (! additionalItems.empty())
    {
        _editSuggestAdditionalRequestId = requestId;
        _editSuggestAdditionalItems     = additionalItems;

        _editSuggestItems         = std::move(additionalItems);
        _editSuggestHoveredIndex  = -1;
        _editSuggestSelectedIndex = -1;
        _editSuggestHighlightText = parseResult.filter;
        UpdateEditSuggestPopupWindow();
    }
    else
    {
        _editSuggestItems.clear();
        _editSuggestHoveredIndex  = -1;
        _editSuggestSelectedIndex = -1;
        _editSuggestHighlightText.clear();
        CloseEditSuggestPopup();
    }
}

void NavigationView::UpdateEditSuggestPopupWindow()
{
    if (! _hWnd || ! _editMode || ! _pathEdit || _editSuggestItems.empty())
    {
        CloseEditSuggestPopup();
        return;
    }

    EnsureD2DResources();
    if (! _d2dFactory || ! _dwriteFactory || ! _pathFormat)
    {
        return;
    }

    if (! RegisterEditSuggestPopupWndClass(_hInstance))
    {
        return;
    }

    const auto chrome = ComputeEditChromeRects(_sectionPathRect, _dpi);

    const int navHeightPx         = std::max(1, static_cast<int>(_sectionPathRect.bottom - _sectionPathRect.top));
    const int minRowHeightPx      = std::max(1, DipsToPixelsInt(40, _dpi));
    const int itemHeight          = std::max(navHeightPx, minRowHeightPx);
    _editSuggestPopupRowHeightPx  = itemHeight;
    const int desiredClientWidth  = std::max(1l, chrome.editRect.right - chrome.editRect.left);
    const size_t itemCount        = std::min(kEditSuggestMaxItems, _editSuggestItems.size());
    const int desiredClientHeight = std::max(1, static_cast<int>(itemCount) * itemHeight);

    const DWORD style   = WS_POPUP;
    const DWORD exStyle = WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;

    RECT windowRect = {0, 0, desiredClientWidth, desiredClientHeight};
    if (! AdjustWindowRectExForDpi(&windowRect, style, FALSE, exStyle, _dpi))
    {
        AdjustWindowRectEx(&windowRect, style, FALSE, exStyle);
    }

    const int winWidth  = windowRect.right - windowRect.left;
    const int winHeight = windowRect.bottom - windowRect.top;

    POINT anchor = {chrome.editRect.left, _sectionPathRect.bottom};
    ClientToScreen(_hWnd.get(), &anchor);

    HMONITOR hMon = MonitorFromPoint(anchor, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (! GetMonitorInfoW(hMon, &mi))
    {
        return;
    }

    const RECT work = mi.rcWork;

    int x = anchor.x;
    int y = anchor.y;

    if (y + winHeight > work.bottom)
    {
        const int aboveY = anchor.y - winHeight;
        if (aboveY >= work.top)
        {
            y = aboveY;
        }
        else
        {
            y = std::max(static_cast<int>(work.top), static_cast<int>(work.bottom - winHeight));
        }
    }

    if (x + winWidth > work.right)
    {
        x = std::max(static_cast<int>(work.left), static_cast<int>(work.right - winWidth));
    }

    x = std::clamp(x, static_cast<int>(work.left), static_cast<int>(work.right - winWidth));
    y = std::clamp(y, static_cast<int>(work.top), static_cast<int>(work.bottom - winHeight));

    if (! _editSuggestPopup)
    {
        HWND popup = CreateWindowExW(exStyle, kEditSuggestPopupClassName, L"", style, x, y, winWidth, winHeight, _hWnd.get(), nullptr, _hInstance, this);
        if (! popup)
        {
            return;
        }

        _editSuggestPopup.reset(popup);
    }
    else
    {
        SetWindowPos(_editSuggestPopup.get(), HWND_TOP, x, y, winWidth, winHeight, SWP_NOACTIVATE);
    }

    RECT clientRect{};
    GetClientRect(_editSuggestPopup.get(), &clientRect);
    _editSuggestPopupClientSize.cx = clientRect.right - clientRect.left;
    _editSuggestPopupClientSize.cy = clientRect.bottom - clientRect.top;

    ShowWindow(_editSuggestPopup.get(), SW_SHOWNOACTIVATE);
    InvalidateRect(_editSuggestPopup.get(), nullptr, FALSE);
}

void NavigationView::CloseEditSuggestPopup()
{
    if (_editSuggestPopup)
    {
        _editSuggestPopup.reset();
        return;
    }

    _editSuggestItems.clear();
    _editSuggestHoveredIndex  = -1;
    _editSuggestSelectedIndex = -1;
    _editSuggestHighlightText.clear();
}

void NavigationView::EnsureEditSuggestPopupD2DResources()
{
    if (! _editSuggestPopup)
    {
        return;
    }

    EnsureD2DResources();
    if (! _d2dFactory)
    {
        return;
    }

    if (! _editSuggestPopupTarget)
    {
        D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
        props.dpiX                          = 96.0f;
        props.dpiY                          = 96.0f;

        D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(
            _editSuggestPopup.get(), D2D1::SizeU(static_cast<UINT32>(_editSuggestPopupClientSize.cx), static_cast<UINT32>(_editSuggestPopupClientSize.cy)));

        wil::com_ptr<ID2D1HwndRenderTarget> target;
        if (FAILED(_d2dFactory->CreateHwndRenderTarget(props, hwndProps, target.addressof())))
        {
            return;
        }

        target->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
        _editSuggestPopupTarget = std::move(target);
    }

    if (_editSuggestPopupTarget)
    {
        if (! _editSuggestPopupBackgroundBrush)
        {
            const COLORREF surface = _appTheme.systemHighContrast ? GetSysColor(COLOR_WINDOW) : _appTheme.menu.background;
            _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(surface), _editSuggestPopupBackgroundBrush.addressof());
        }
        if (! _editSuggestPopupTextBrush)
        {
            _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(_appTheme.menu.text), _editSuggestPopupTextBrush.addressof());
        }
        if (! _editSuggestPopupDisabledTextBrush)
        {
            _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(_appTheme.menu.disabledText), _editSuggestPopupDisabledTextBrush.addressof());
        }
        if (! _editSuggestPopupHighlightBrush)
        {
            _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(_appTheme.menu.selectionBg), _editSuggestPopupHighlightBrush.addressof());
        }
        if (! _editSuggestPopupHoverBrush)
        {
            const COLORREF surface    = _appTheme.systemHighContrast ? GetSysColor(COLOR_WINDOW) : _appTheme.menu.background;
            const int highlightWeight = _appTheme.dark ? 30 : 18;
            const COLORREF highlightColor =
                _appTheme.systemHighContrast ? GetSysColor(COLOR_HIGHLIGHT) : ThemedControls::BlendColor(surface, _appTheme.menu.text, highlightWeight, 255);
            _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(highlightColor), _editSuggestPopupHoverBrush.addressof());
        }
        if (! _editSuggestPopupBorderBrush)
        {
            if (! _appTheme.systemHighContrast)
            {
                const COLORREF surface = _appTheme.menu.background;
                const COLORREF border  = ThemedControls::BlendColor(surface, _appTheme.menu.text, _appTheme.dark ? 60 : 40, 255);
                _editSuggestPopupTarget->CreateSolidColorBrush(ColorFromCOLORREF(border), _editSuggestPopupBorderBrush.addressof());
            }
        }
    }
}

void NavigationView::DiscardEditSuggestPopupD2DResources()
{
    _editSuggestPopupBorderBrush       = nullptr;
    _editSuggestPopupBackgroundBrush   = nullptr;
    _editSuggestPopupHoverBrush        = nullptr;
    _editSuggestPopupHighlightBrush    = nullptr;
    _editSuggestPopupDisabledTextBrush = nullptr;
    _editSuggestPopupTextBrush         = nullptr;
    _editSuggestPopupTarget            = nullptr;
}

void NavigationView::RenderEditSuggestPopup()
{
    if (! _editSuggestPopup)
    {
        return;
    }

    PAINTSTRUCT ps;
    wil::unique_hdc_paint hdc = wil::BeginPaint(_editSuggestPopup.get(), &ps);
    static_cast<void>(hdc);

    EnsureEditSuggestPopupD2DResources();
    if (! _editSuggestPopupTarget || ! _dwriteFactory || ! _pathFormat || ! _editSuggestPopupBackgroundBrush || ! _editSuggestPopupTextBrush)
    {
        return;
    }

    _editSuggestPopupTarget->BeginDraw();
    auto endDraw = wil::scope_exit(
        [&]
        {
            const HRESULT hr = _editSuggestPopupTarget->EndDraw();
            if (hr == D2DERR_RECREATE_TARGET)
            {
                DiscardEditSuggestPopupD2DResources();
            }
        });

    const float width            = static_cast<float>(_editSuggestPopupClientSize.cx);
    const float height           = static_cast<float>(_editSuggestPopupClientSize.cy);
    const D2D1_RECT_F clientRect = D2D1::RectF(0.0f, 0.0f, width, height);

    _editSuggestPopupTarget->FillRectangle(clientRect, _editSuggestPopupBackgroundBrush.get());

    const float rowHeight = static_cast<float>(
        std::max(1, _editSuggestPopupRowHeightPx > 0 ? _editSuggestPopupRowHeightPx : static_cast<int>(_sectionPathRect.bottom - _sectionPathRect.top)));

    const float highlightInsetX = DipsToPixels(6.0f, _dpi);
    const float highlightInsetY = DipsToPixels(2.0f, _dpi);
    const float highlightRadius = DipsToPixels(8.0f, _dpi);

    const float barWidth  = DipsToPixels(5.0f, _dpi);
    const float barInsetX = DipsToPixels(4.0f, _dpi);
    const float barInsetY = DipsToPixels(4.0f, _dpi);
    const float barRadius = DipsToPixels(4.0f, _dpi);

    const float textInsetX       = DipsToPixels(22.0f, _dpi);
    const float textPaddingRight = DipsToPixels(22.0f, _dpi);

    const int activeIndex = _editSuggestSelectedIndex >= 0 ? _editSuggestSelectedIndex : _editSuggestHoveredIndex;

    const size_t count = std::min(kEditSuggestMaxItems, _editSuggestItems.size());
    for (size_t i = 0; i < count; ++i)
    {
        const float top     = rowHeight * static_cast<float>(i);
        D2D1_RECT_F rowRect = D2D1::RectF(0.0f, top, width, top + rowHeight);

        const auto& item    = _editSuggestItems[i];
        const bool enabled  = item.enabled;
        const bool selected = enabled && (static_cast<int>(i) == activeIndex);
        if (selected && _editSuggestPopupHoverBrush)
        {
            const D2D1_RECT_F highlightRect = InsetRectF(rowRect, highlightInsetX, highlightInsetY);
            _editSuggestPopupTarget->FillRoundedRectangle(RoundedRect(highlightRect, highlightRadius), _editSuggestPopupHoverBrush.get());

            if (_editSuggestPopupHighlightBrush)
            {
                D2D1_RECT_F barRect = highlightRect;
                barRect.left        = std::min(barRect.right, barRect.left + barInsetX);
                barRect.right       = std::min(barRect.right, barRect.left + barWidth);
                barRect.top         = std::min(barRect.bottom, barRect.top + barInsetY);
                barRect.bottom      = std::max(barRect.top, barRect.bottom - barInsetY);

                _editSuggestPopupTarget->FillRoundedRectangle(RoundedRect(barRect, barRadius), _editSuggestPopupHighlightBrush.get());
            }
        }

        D2D1_RECT_F textRect = rowRect;
        textRect.left        = std::min(textRect.right, textRect.left + textInsetX);
        textRect.right       = std::max(textRect.left, textRect.right - textPaddingRight);

        const auto& text = item.display;
        if (text.empty())
        {
            continue;
        }

        wil::com_ptr<IDWriteTextLayout> layout;
        const float layoutWidth  = std::max(1.0f, textRect.right - textRect.left);
        const float layoutHeight = std::max(1.0f, rowHeight);
        if (SUCCEEDED(_dwriteFactory->CreateTextLayout(
                text.data(), static_cast<UINT32>(text.size()), _pathFormat.get(), layoutWidth, layoutHeight, layout.addressof())) &&
            layout)
        {
            layout->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
            layout->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

            if (enabled && ! _editSuggestHighlightText.empty() && _editSuggestPopupHighlightBrush)
            {
                size_t searchStart = 0;
                while (searchStart < text.size())
                {
                    const size_t remaining = text.size() - searchStart;
                    const size_t maxInt    = static_cast<size_t>(std::numeric_limits<int>::max());
                    const int sourceLen    = static_cast<int>(std::min(remaining, maxInt));
                    const int valueLen     = static_cast<int>(std::min(_editSuggestHighlightText.size(), maxInt));
                    if (valueLen <= 0 || sourceLen < valueLen)
                    {
                        break;
                    }

                    const int foundAt = FindStringOrdinal(0, text.data() + searchStart, sourceLen, _editSuggestHighlightText.data(), valueLen, TRUE);
                    if (foundAt < 0)
                    {
                        break;
                    }

                    const size_t matchStart  = searchStart + static_cast<size_t>(foundAt);
                    const size_t matchLength = std::min(_editSuggestHighlightText.size(), text.size() - matchStart);
                    if (matchLength == 0u)
                    {
                        break;
                    }

                    const size_t maxUInt32 = static_cast<size_t>(std::numeric_limits<UINT32>::max());
                    const UINT32 startPos  = static_cast<UINT32>(std::min(matchStart, maxUInt32));
                    const UINT32 len       = static_cast<UINT32>(std::min(matchLength, maxUInt32));
                    const DWRITE_TEXT_RANGE range{startPos, len};

                    static_cast<void>(layout->SetDrawingEffect(_editSuggestPopupHighlightBrush.get(), range));
                    static_cast<void>(layout->SetFontWeight(DWRITE_FONT_WEIGHT_SEMI_BOLD, range));

                    searchStart = matchStart + matchLength;
                }
            }

            ID2D1SolidColorBrush* brush = _editSuggestPopupTextBrush.get();
            if (! enabled && _editSuggestPopupDisabledTextBrush)
            {
                brush = _editSuggestPopupDisabledTextBrush.get();
            }

            _editSuggestPopupTarget->DrawTextLayout(
                D2D1::Point2F(textRect.left, rowRect.top), layout.get(), brush, D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);
        }
    }

    if (_editSuggestPopupBorderBrush)
    {
        const D2D1_RECT_F borderRect = InsetRectF(clientRect, 0.5f, 0.5f);
        _editSuggestPopupTarget->DrawRoundedRectangle(RoundedRect(borderRect, highlightRadius), _editSuggestPopupBorderBrush.get(), 1.0f);
    }
}

void NavigationView::ApplyEditSuggestIndex(size_t index)
{
    if (! _editMode || ! _pathEdit || index >= _editSuggestItems.size())
    {
        return;
    }

    const auto& item = _editSuggestItems[index];
    if (! item.enabled || item.insertText.empty())
    {
        return;
    }

    std::wstring text = item.insertText;
    if (! text.empty() && text.back() != L'\\' && text.back() != L'/')
    {
        if (item.directorySeparator != L'\0')
        {
            text.push_back(item.directorySeparator);
        }
    }

    SetWindowTextW(_pathEdit.get(), text.c_str());
    const LONG caret = static_cast<LONG>(std::min<size_t>(text.size(), static_cast<size_t>(std::numeric_limits<LONG>::max())));
    SendMessageW(_pathEdit.get(), EM_SETSEL, static_cast<WPARAM>(caret), static_cast<LPARAM>(caret));
    SetFocus(_pathEdit.get());
    UpdateEditSuggest();
}

void NavigationView::EnsureEditSuggestWorker()
{
    if (_editSuggestThread.joinable())
    {
        return;
    }

    _editSuggestThread = std::jthread([this](std::stop_token stopToken) { EditSuggestWorker(stopToken); });
}

void NavigationView::EnsureSiblingPrefetchWorker()
{
    if (_siblingPrefetchThread.joinable())
    {
        return;
    }

    _siblingPrefetchThread = std::jthread([this](std::stop_token stopToken) { SiblingPrefetchWorker(stopToken); });
}

void NavigationView::QueueSiblingPrefetchForPath(const std::filesystem::path& displayPath)
{
    if (! _fileSystemPlugin)
    {
        return;
    }

    // /@conn: is a host-reserved prefix used by connection manager routing.
    // Prefetching parents like "/@conn:" or "/" triggers invalid enumerations for curl-backed protocols
    // (they require either an authority //host/... or a concrete /@conn:<name>/...), and can also cause
    // redundant remote calls right after Connect.
    const std::wstring_view displayText(displayPath.native());
    if (displayText.starts_with(L"/@conn:"))
    {
        return;
    }

    constexpr size_t kMaxFolders = 16u;

    const auto parts = SplitPathComponents(displayPath);
    if (parts.size() < 2u)
    {
        return;
    }

    std::vector<std::filesystem::path> folders;
    folders.reserve(std::min(parts.size(), kMaxFolders));

    for (size_t index = parts.size() - 1; index > 0; --index)
    {
        const std::filesystem::path normalized = NormalizeDirectoryPath(parts[index].fullPath);
        const std::filesystem::path parent     = normalized.parent_path();
        if (parent.empty())
        {
            continue;
        }

        const std::filesystem::path pluginParent = ToPluginPath(parent);
        if (pluginParent.empty())
        {
            continue;
        }

        const std::wstring_view pluginText(pluginParent.native());
        bool alreadyQueued = false;
        for (const auto& existing : folders)
        {
            if (EqualsNoCase(existing.native(), pluginText))
            {
                alreadyQueued = true;
                break;
            }
        }
        if (alreadyQueued)
        {
            continue;
        }

        folders.push_back(pluginParent);
        if (folders.size() >= kMaxFolders)
        {
            break;
        }
    }

    if (folders.empty())
    {
        return;
    }

    EnsureSiblingPrefetchWorker();
    const uint64_t requestId = _siblingPrefetchRequestId.fetch_add(1, std::memory_order_acq_rel) + 1u;

    {
        std::lock_guard lock(_siblingPrefetchMutex);
        SiblingPrefetchQuery query{};
        query.requestId              = requestId;
        query.fileSystem             = _fileSystemPlugin;
        query.folders                = std::move(folders);
        _siblingPrefetchPendingQuery = std::move(query);
    }

    _siblingPrefetchCv.notify_one();
}

void NavigationView::QueueSiblingPrefetchForParent(const std::filesystem::path& parentPath)
{
    if (! _fileSystemPlugin)
    {
        return;
    }

    const std::filesystem::path pluginParent = ToPluginPath(parentPath);
    if (pluginParent.empty())
    {
        return;
    }

    std::vector<std::filesystem::path> folders;
    folders.push_back(pluginParent);

    EnsureSiblingPrefetchWorker();
    const uint64_t requestId = _siblingPrefetchRequestId.fetch_add(1, std::memory_order_acq_rel) + 1u;

    {
        std::lock_guard lock(_siblingPrefetchMutex);
        SiblingPrefetchQuery query{};
        query.requestId              = requestId;
        query.fileSystem             = _fileSystemPlugin;
        query.folders                = std::move(folders);
        _siblingPrefetchPendingQuery = std::move(query);
    }

    _siblingPrefetchCv.notify_one();
}

void NavigationView::SiblingPrefetchWorker(std::stop_token stopToken)
{
    std::stop_callback stopCallback(stopToken, [this] { _siblingPrefetchCv.notify_all(); });

    for (;;)
    {
        if (stopToken.stop_requested())
        {
            return;
        }

        SiblingPrefetchQuery query{};
        {
            std::unique_lock lock(_siblingPrefetchMutex);
            _siblingPrefetchCv.wait(lock, [&] { return stopToken.stop_requested() || _siblingPrefetchPendingQuery.has_value(); });
            if (stopToken.stop_requested())
            {
                return;
            }

            query = std::move(_siblingPrefetchPendingQuery.value());
            _siblingPrefetchPendingQuery.reset();
        }

        if (! query.fileSystem)
        {
            continue;
        }

        for (const auto& folder : query.folders)
        {
            if (stopToken.stop_requested())
            {
                return;
            }

            const uint64_t latest = _siblingPrefetchRequestId.load(std::memory_order_acquire);
            if (query.requestId != latest)
            {
                break;
            }

            auto borrowed =
                DirectoryInfoCache::GetInstance().BorrowDirectoryInfo(query.fileSystem.get(), folder, DirectoryInfoCache::BorrowMode::AllowEnumerate);
            static_cast<void>(borrowed);
        }
    }
}

void NavigationView::EditSuggestWorker(std::stop_token stopToken)
{
    std::stop_callback stopCallback(stopToken, [this] { _editSuggestCv.notify_all(); });

    for (;;)
    {
        if (stopToken.stop_requested())
        {
            return;
        }

        EditSuggestQuery query{};
        {
            std::unique_lock lock(_editSuggestMutex);
            _editSuggestCv.wait(lock, [&] { return stopToken.stop_requested() || _editSuggestPendingQuery.has_value(); });
            if (stopToken.stop_requested())
            {
                return;
            }

            query = std::move(_editSuggestPendingQuery.value());
            _editSuggestPendingQuery.reset();
        }

        std::vector<std::wstring> names;

        if (query.fileSystem)
        {
            auto borrowed = DirectoryInfoCache::GetInstance().BorrowDirectoryInfo(
                query.fileSystem.get(), query.pluginFolder, DirectoryInfoCache::BorrowMode::AllowEnumerate);
            IFilesInformation* info = borrowed.Get();
            if (borrowed.Status() == S_OK && info)
            {
                AppendMatchingDirectoryNamesFromFilesInformation(info, query.prefix, names);
            }
        }

        const bool hasMore = SortAndTrimEditSuggestNames(names);

        std::vector<std::wstring> displayItems;
        std::vector<std::wstring> insertItems;
        BuildEditSuggestLists(query.displayFolder, names, query.directorySeparator, displayItems, insertItems);

        if (stopToken.stop_requested())
        {
            return;
        }

        PostEditSuggestResults(query.requestId, hasMore, query.directorySeparator, std::move(query.prefix), std::move(displayItems), std::move(insertItems));
    }
}

void NavigationView::PostEditSuggestResults(uint64_t requestId,
                                            bool hasMore,
                                            wchar_t directorySeparator,
                                            std::wstring&& highlightText,
                                            std::vector<std::wstring>&& displayItems,
                                            std::vector<std::wstring>&& insertItems)
{
    if (! _hWnd)
    {
        return;
    }

    auto payload                = std::make_unique<EditSuggestResultsPayload>();
    payload->requestId          = requestId;
    payload->hasMore            = hasMore;
    payload->directorySeparator = directorySeparator;
    payload->highlightText      = std::move(highlightText);
    payload->displayItems       = std::move(displayItems);
    payload->insertItems        = std::move(insertItems);
    static_cast<void>(PostMessagePayload(_hWnd.get(), WndMsg::kEditSuggestResults, 0, std::move(payload)));
}

bool NavigationView::ValidatePath(const std::wstring& pathStr)
{
    const std::wstring_view text(pathStr);

    if (text.size() >= 6u)
    {
        constexpr std::wstring_view kConnPrefix = L"@conn:";
        const int prefixLen                     = static_cast<int>(kConnPrefix.size());
        if (CompareStringOrdinal(text.data(), prefixLen, kConnPrefix.data(), prefixLen, TRUE) == CSTR_EQUAL)
        {
            return true;
        }
    }

    const size_t colon = text.find(L':');
    if (colon != std::wstring_view::npos && colon >= 2)
    {
        const size_t sep = text.find_first_of(L"\\/");
        if (sep == std::wstring_view::npos || sep > colon)
        {
            bool ok = true;
            for (size_t i = 0; i < colon; ++i)
            {
                if (std::iswalnum(text[i]) == 0)
                {
                    ok = false;
                    break;
                }
            }

            if (ok)
            {
                return true;
            }
        }
    }

    if (! EqualsNoCase(_pluginShortId, L"file") && LooksLikeWindowsAbsolutePath(text))
    {
        // Allow switching to the file plugin; validation will happen during plugin enumeration.
        return true;
    }

    if (! _pluginShortId.empty() && ! EqualsNoCase(_pluginShortId, L"file"))
    {
        return false;
    }

    if (! _fileSystemIo)
    {
        return false;
    }

    unsigned long attrs = 0;
    const HRESULT hr    = _fileSystemIo->GetAttributes(pathStr.c_str(), &attrs);
    if (FAILED(hr))
    {
        return false;
    }

    return (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

LRESULT NavigationView::OnEditSubclassNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp, UINT_PTR subclassId) noexcept
{
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: pointer or reference to potentially throwing function passed to 'extern "C"' function
    RemoveWindowSubclass(hwnd, NavigationView::EditSubclassProc, subclassId);
#pragma warning(pop)
    return DefSubclassProc(hwnd, WM_NCDESTROY, wp, lp);
}

bool NavigationView::HandleEditSubclassKeyDown(HWND editHwnd, WPARAM key)
{
    _suppressCtrlBackspaceCharHwnd = nullptr;

    const bool isPopupEdit = _fullPathPopupEdit && editHwnd == _fullPathPopupEdit.get();
    if (key == VK_BACK)
    {
        const bool ctrlDown = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool altDown  = (GetKeyState(VK_MENU) & 0x8000) != 0;
        if (ctrlDown && ! altDown)
        {
            DWORD selectionStart = 0;
            DWORD selectionEnd   = 0;
            SendMessageW(editHwnd, EM_GETSEL, reinterpret_cast<WPARAM>(&selectionStart), reinterpret_cast<LPARAM>(&selectionEnd));

            if (selectionStart != selectionEnd)
            {
                SendMessageW(editHwnd, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
                _suppressCtrlBackspaceCharHwnd = editHwnd;
                return true;
            }

            const int length = GetWindowTextLengthW(editHwnd);
            std::wstring text;
            text.resize(static_cast<size_t>(std::max(0, length)) + 1u);
            GetWindowTextW(editHwnd, text.data(), static_cast<int>(text.size()));
            text.resize(wcsnlen(text.c_str(), text.size()));

            const size_t caret = std::min(static_cast<size_t>(selectionEnd), text.size());
            if (caret == 0u)
            {
                _suppressCtrlBackspaceCharHwnd = editHwnd;
                return true;
            }

            auto isSeparator = [](wchar_t ch) noexcept { return ch == L'\\' || ch == L'/'; };

            size_t eraseFrom = caret;
            while (eraseFrom > 0u && std::iswspace(text[eraseFrom - 1u]) != 0)
            {
                --eraseFrom;
            }
            while (eraseFrom > 0u && isSeparator(text[eraseFrom - 1u]))
            {
                --eraseFrom;
            }
            while (eraseFrom > 0u)
            {
                const wchar_t ch = text[eraseFrom - 1u];
                if (std::iswspace(ch) != 0 || isSeparator(ch))
                {
                    break;
                }
                --eraseFrom;
            }

            if (eraseFrom == caret)
            {
                eraseFrom = caret > 0u ? (caret - 1u) : 0u;
            }

            SendMessageW(editHwnd, EM_SETSEL, static_cast<WPARAM>(eraseFrom), static_cast<LPARAM>(caret));
            SendMessageW(editHwnd, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L""));
            _suppressCtrlBackspaceCharHwnd = editHwnd;
            return true;
        }
    }

    if (key == VK_RETURN)
    {
        if (! isPopupEdit && _editSuggestSelectedIndex >= 0 && static_cast<size_t>(_editSuggestSelectedIndex) < _editSuggestItems.size())
        {
            ApplyEditSuggestIndex(static_cast<size_t>(_editSuggestSelectedIndex));
        }
        else if (isPopupEdit)
        {
            ExitFullPathPopupEditMode(true);
        }
        else
        {
            ExitEditMode(true);
            if (! _editMode && _requestFolderViewFocusCallback)
            {
                _requestFolderViewFocusCallback();
            }
        }
        return true;
    }

    if (key == VK_ESCAPE)
    {
        if (! isPopupEdit && _editSuggestPopup)
        {
            CloseEditSuggestPopup();
            return true;
        }

        if (isPopupEdit)
        {
            ExitFullPathPopupEditMode(false);
        }
        else
        {
            ExitEditMode(false);
            if (_requestFolderViewFocusCallback)
            {
                _requestFolderViewFocusCallback();
            }
        }
        return true;
    }

    if (! isPopupEdit && (key == VK_DOWN || key == VK_UP) && _editSuggestPopup && ! _editSuggestItems.empty())
    {
        const int count = static_cast<int>(_editSuggestItems.size());
        if (count > 0)
        {
            int next = _editSuggestSelectedIndex;
            if (key == VK_DOWN)
            {
                next = (next < 0) ? 0 : std::min(next + 1, count - 1);
                while (next < count && ! _editSuggestItems[static_cast<size_t>(next)].enabled)
                {
                    ++next;
                }
                if (next >= count)
                {
                    next = _editSuggestSelectedIndex;
                }
            }
            else
            {
                next = (next < 0) ? (count - 1) : std::max(next - 1, 0);
                while (next >= 0 && ! _editSuggestItems[static_cast<size_t>(next)].enabled)
                {
                    --next;
                }
                if (next < 0)
                {
                    next = _editSuggestSelectedIndex;
                }
            }

            if (next != _editSuggestSelectedIndex)
            {
                _editSuggestSelectedIndex = next;
                InvalidateRect(_editSuggestPopup.get(), nullptr, FALSE);
            }
        }
        return true;
    }

    if (key == VK_TAB)
    {
        const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

        if (isPopupEdit)
        {
            ExitFullPathPopupEditMode(false);
        }
        else
        {
            ExitEditMode(false);
        }
        if (_hWnd)
        {
            SetFocus(_hWnd.get());
        }
        MoveFocus(! shift);
        return true;
    }

    return false;
}

bool NavigationView::HandleEditSubclassChar(HWND editHwnd, WPARAM key)
{
    if (_suppressCtrlBackspaceCharHwnd && _suppressCtrlBackspaceCharHwnd == editHwnd && key == 0x7Fu)
    {
        _suppressCtrlBackspaceCharHwnd = nullptr;
        return true;
    }
    return key == L'\r' || key == L'\n';
}

bool NavigationView::HandleEditSubclassPaste(HWND editHwnd)
{
    if (OpenClipboard(editHwnd) == 0)
    {
        return false;
    }

    auto closeClipboard = wil::scope_exit([] { CloseClipboard(); });

    HANDLE hText = GetClipboardData(CF_UNICODETEXT);
    if (! hText)
    {
        return false;
    }

    const auto* raw = static_cast<const wchar_t*>(GlobalLock(hText));
    if (! raw)
    {
        return false;
    }

    auto unlock = wil::scope_exit([&] { GlobalUnlock(hText); });

    std::wstring text(raw);
    text.erase(std::remove_if(text.begin(), text.end(), [](wchar_t ch) { return ch == L'\r' || ch == L'\n'; }), text.end());
    SendMessageW(editHwnd, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(text.c_str()));
    return true;
}

namespace
{
void NotifyPaneFocusChangedForEdit(HWND editHwnd) noexcept
{
    const HWND navigationView = GetParent(editHwnd);
    if (! navigationView)
    {
        return;
    }

    const HWND paneWindow = GetParent(navigationView);
    if (! paneWindow)
    {
        return;
    }

    PostMessageW(paneWindow, WndMsg::kPaneFocusChanged, 0, 0);
}
} // namespace

LRESULT CALLBACK NavigationView::EditSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData)
{
    auto* self = reinterpret_cast<NavigationView*>(refData);

    switch (msg)
    {
        case WM_SETFOCUS: NotifyPaneFocusChangedForEdit(hwnd); break;
        case WM_KILLFOCUS: NotifyPaneFocusChangedForEdit(hwnd); break;
        case WM_KEYDOWN:
            if (self && self->HandleEditSubclassKeyDown(hwnd, wp))
            {
                return 0;
            }
            break;
        case WM_CHAR:
            if (self && self->HandleEditSubclassChar(hwnd, wp))
            {
                return 0;
            }
            break;
        case WM_PASTE:
            if (self && self->HandleEditSubclassPaste(hwnd))
            {
                return 0;
            }
            break;
        case WM_NCDESTROY: return OnEditSubclassNcDestroy(hwnd, wp, lp, subclassId);
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}
