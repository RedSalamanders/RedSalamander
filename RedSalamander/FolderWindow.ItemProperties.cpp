#include "FolderWindow.h"

#include "AppTheme.h"
#include "Resource.h"
#include "ThemedControls.h"
#include "WindowMaximizeBehavior.h"
#include "WindowPlacementPersistence.h"

#include "Helpers.h"
#include "SettingsSave.h"

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/resource.h>
#pragma warning(pop)

#include <CommCtrl.h>
#include <ShlObj.h>
#include <Uxtheme.h>
#include <yyjson.h>

#include <algorithm>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace
{
struct ItemPropertiesField
{
    std::wstring key;
    std::wstring value;
};

struct ItemPropertiesSection
{
    std::wstring title;
    std::vector<ItemPropertiesField> fields;
};

struct ItemPropertiesDocument
{
    std::wstring title;
    std::vector<ItemPropertiesSection> sections;
};

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    const int required = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::optional<ItemPropertiesDocument> TryParseItemPropertiesJson(std::string_view jsonUtf8) noexcept
{
    if (jsonUtf8.empty())
    {
        return std::nullopt;
    }

    // yyjson may modify the input buffer; it requires a mutable char*.
    std::string jsonCopy(jsonUtf8);
    yyjson_read_err err{};
    yyjson_doc* doc = yyjson_read_opts(jsonCopy.data(), jsonCopy.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
    if (! doc)
    {
        return std::nullopt;
    }
    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return std::nullopt;
    }

    yyjson_val* versionVal = yyjson_obj_get(root, "version");
    if (! versionVal || ! yyjson_is_int(versionVal) || yyjson_get_int(versionVal) != 1)
    {
        return std::nullopt;
    }

    ItemPropertiesDocument out{};

    if (yyjson_val* titleVal = yyjson_obj_get(root, "title"); titleVal && yyjson_is_str(titleVal))
    {
        if (const char* titleUtf8 = yyjson_get_str(titleVal); titleUtf8 && titleUtf8[0] != '\0')
        {
            out.title = Utf16FromUtf8(titleUtf8);
        }
    }

    yyjson_val* sectionsVal = yyjson_obj_get(root, "sections");
    if (! sectionsVal || ! yyjson_is_arr(sectionsVal))
    {
        return out;
    }

    const size_t sectionCount = yyjson_arr_size(sectionsVal);
    out.sections.reserve(sectionCount);

    for (size_t i = 0; i < sectionCount; ++i)
    {
        yyjson_val* sectionVal = yyjson_arr_get(sectionsVal, i);
        if (! sectionVal || ! yyjson_is_obj(sectionVal))
        {
            continue;
        }

        ItemPropertiesSection section{};
        if (yyjson_val* sectionTitleVal = yyjson_obj_get(sectionVal, "title"); sectionTitleVal && yyjson_is_str(sectionTitleVal))
        {
            if (const char* titleUtf8 = yyjson_get_str(sectionTitleVal); titleUtf8 && titleUtf8[0] != '\0')
            {
                section.title = Utf16FromUtf8(titleUtf8);
            }
        }

        if (yyjson_val* fieldsVal = yyjson_obj_get(sectionVal, "fields"); fieldsVal && yyjson_is_arr(fieldsVal))
        {
            const size_t fieldCount = yyjson_arr_size(fieldsVal);
            section.fields.reserve(fieldCount);

            for (size_t f = 0; f < fieldCount; ++f)
            {
                yyjson_val* fieldVal = yyjson_arr_get(fieldsVal, f);
                if (! fieldVal || ! yyjson_is_obj(fieldVal))
                {
                    continue;
                }

                yyjson_val* keyVal   = yyjson_obj_get(fieldVal, "key");
                yyjson_val* valueVal = yyjson_obj_get(fieldVal, "value");
                if (! keyVal || ! valueVal || ! yyjson_is_str(keyVal) || ! yyjson_is_str(valueVal))
                {
                    continue;
                }

                const char* keyUtf8   = yyjson_get_str(keyVal);
                const char* valueUtf8 = yyjson_get_str(valueVal);
                if (! keyUtf8 || keyUtf8[0] == '\0' || ! valueUtf8)
                {
                    continue;
                }

                ItemPropertiesField field{};
                field.key   = Utf16FromUtf8(keyUtf8);
                field.value = Utf16FromUtf8(valueUtf8);
                if (! field.key.empty())
                {
                    section.fields.emplace_back(std::move(field));
                }
            }
        }

        out.sections.emplace_back(std::move(section));
    }

    return out;
}

constexpr wchar_t kItemPropertiesWindowClass[] = L"RedSalamander.ItemPropertiesWindow";
constexpr int kItemPropertiesListId            = 1001;
constexpr int kItemPropertiesCloseId           = 1002;
constexpr wchar_t kItemPropertiesWindowId[]    = L"ItemPropertiesWindow";
constexpr wchar_t kSettingsAppId[]             = L"RedSalamander";

struct ItemPropertiesWindowState
{
    explicit ItemPropertiesWindowState(const AppTheme& themeIn, ItemPropertiesDocument docIn) noexcept : theme(themeIn), doc(std::move(docIn))
    {
    }

    ItemPropertiesWindowState(const ItemPropertiesWindowState&)            = delete;
    ItemPropertiesWindowState& operator=(const ItemPropertiesWindowState&) = delete;
    ItemPropertiesWindowState(ItemPropertiesWindowState&&)                 = delete;
    ItemPropertiesWindowState& operator=(ItemPropertiesWindowState&&)      = delete;

    AppTheme theme;
    ItemPropertiesDocument doc;
    Common::Settings::Settings* settings = nullptr;
    UINT dpi                             = USER_DEFAULT_SCREEN_DPI;

    wil::unique_hwnd list;
    wil::unique_hwnd closeButton;

    wil::unique_hbrush backgroundBrush;
};

void SetWindowThemeForMode(HWND hwnd, const AppTheme& theme) noexcept
{
    if (theme.highContrast)
    {
        ::SetWindowTheme(hwnd, L"", nullptr);
        return;
    }

    ::SetWindowTheme(hwnd, theme.dark ? L"DarkMode_Explorer" : L"Explorer", nullptr);
}

void LayoutItemPropertiesWindow(HWND hwnd, ItemPropertiesWindowState& state) noexcept
{
    RECT rc{};
    if (::GetClientRect(hwnd, &rc) == 0)
    {
        return;
    }

    const int width  = std::max(0l, rc.right - rc.left);
    const int height = std::max(0l, rc.bottom - rc.top);

    const int margin  = ThemedControls::ScaleDip(state.dpi, 10);
    const int buttonH = ThemedControls::ScaleDip(state.dpi, 28);
    const int buttonW = std::max(ThemedControls::ScaleDip(state.dpi, 90), buttonH * 3);

    const int buttonTop  = std::max(margin, height - margin - buttonH);
    const int buttonLeft = std::max(margin, width - margin - buttonW);

    if (state.closeButton)
    {
        ::SetWindowPos(state.closeButton.get(), nullptr, buttonLeft, buttonTop, buttonW, buttonH, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    if (state.list)
    {
        const int listTop    = margin;
        const int listHeight = std::max(0, buttonTop - margin - listTop);

        ::SetWindowPos(state.list.get(), nullptr, margin, listTop, std::max(0, width - margin * 2), listHeight, SWP_NOZORDER | SWP_NOACTIVATE);

        RECT listRc{};
        if (::GetClientRect(state.list.get(), &listRc) != 0)
        {
            const int listW = std::max(0l, listRc.right - listRc.left);

            const int keyW   = std::clamp(ThemedControls::ScaleDip(state.dpi, 180), 80, std::max(80, listW / 2));
            const int valueW = std::max(80, listW - keyW - ThemedControls::ScaleDip(state.dpi, 16));

            ListView_SetColumnWidth(state.list.get(), 0, keyW);
            ListView_SetColumnWidth(state.list.get(), 1, valueW);
        }
    }
}

void PopulateItemPropertiesList(HWND list, const ItemPropertiesDocument& doc) noexcept
{
    if (! list)
    {
        return;
    }

    ListView_DeleteAllItems(list);

    // Enable groups (best-effort).
    ListView_EnableGroupView(list, TRUE);

    int nextGroupId = 1;
    for (const auto& section : doc.sections)
    {
        const std::wstring header = section.title.empty() ? L"" : section.title;

        LVGROUP group{};
        group.cbSize    = sizeof(group);
        group.mask      = LVGF_GROUPID | LVGF_HEADER;
        group.iGroupId  = nextGroupId;
        group.pszHeader = const_cast<wchar_t*>(header.c_str());
        ListView_InsertGroup(list, -1, &group);

        int itemIndex = ListView_GetItemCount(list);
        for (const auto& field : section.fields)
        {
            LVITEM item{};
            item.mask          = LVIF_TEXT | LVIF_GROUPID;
            item.iItem         = itemIndex++;
            item.iSubItem      = 0;
            item.iGroupId      = nextGroupId;
            item.pszText       = const_cast<wchar_t*>(field.key.c_str());
            const int inserted = ListView_InsertItem(list, &item);
            if (inserted >= 0)
            {
                ListView_SetItemText(list, inserted, 1, const_cast<wchar_t*>(field.value.c_str()));
            }
        }

        ++nextGroupId;
    }
}

LRESULT CALLBACK ItemPropertiesWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) noexcept
{
    auto* state = reinterpret_cast<ItemPropertiesWindowState*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (msg)
    {
        case WM_NCCREATE:
        {
            const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            auto* newState = static_cast<ItemPropertiesWindowState*>(cs->lpCreateParams);
            ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(newState));
            return TRUE;
        }
        case WM_CREATE:
        {
            if (! state)
            {
                return -1;
            }

            state->dpi = ::GetDpiForWindow(hwnd);
            state->backgroundBrush.reset(::CreateSolidBrush(state->theme.windowBackground));

            const std::wstring closeText = LoadStringResource(nullptr, IDS_PROPERTIES_BTN_CLOSE);
            const std::wstring keyText   = LoadStringResource(nullptr, IDS_PROPERTIES_COL_KEY);
            const std::wstring valueText = LoadStringResource(nullptr, IDS_PROPERTIES_COL_VALUE);

            state->list.reset(::CreateWindowExW(WS_EX_CLIENTEDGE,
                                                WC_LISTVIEWW,
                                                L"",
                                                WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SHOWSELALWAYS,
                                                0,
                                                0,
                                                10,
                                                10,
                                                hwnd,
                                                reinterpret_cast<HMENU>(static_cast<INT_PTR>(kItemPropertiesListId)),
                                                ::GetModuleHandleW(nullptr),
                                                nullptr));
            if (! state->list)
            {
                return -1;
            }

            state->closeButton.reset(::CreateWindowExW(0,
                                                       L"Button",
                                                       closeText.c_str(),
                                                       WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
                                                       0,
                                                       0,
                                                       10,
                                                       10,
                                                       hwnd,
                                                       reinterpret_cast<HMENU>(static_cast<INT_PTR>(kItemPropertiesCloseId)),
                                                       ::GetModuleHandleW(nullptr),
                                                       nullptr));
            if (! state->closeButton)
            {
                return -1;
            }

            if (! state->theme.highContrast)
            {
                ThemedControls::EnableOwnerDrawButton(hwnd, kItemPropertiesCloseId);
            }

            SetWindowThemeForMode(state->list.get(), state->theme);
            SetWindowThemeForMode(state->closeButton.get(), state->theme);

            ListView_SetExtendedListViewStyle(state->list.get(), LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP | LVS_EX_INFOTIP);

            LVCOLUMNW col0{};
            col0.mask    = LVCF_TEXT | LVCF_WIDTH;
            col0.pszText = const_cast<wchar_t*>(keyText.c_str());
            col0.cx      = ThemedControls::ScaleDip(state->dpi, 180);
            ListView_InsertColumn(state->list.get(), 0, &col0);

            LVCOLUMNW col1{};
            col1.mask    = LVCF_TEXT | LVCF_WIDTH;
            col1.pszText = const_cast<wchar_t*>(valueText.c_str());
            col1.cx      = ThemedControls::ScaleDip(state->dpi, 420);
            ListView_InsertColumn(state->list.get(), 1, &col1);

            ThemedControls::ApplyThemeToListView(state->list.get(), state->theme);
            ThemedControls::EnsureListViewHeaderThemed(state->list.get(), state->theme);

            PopulateItemPropertiesList(state->list.get(), state->doc);
            LayoutItemPropertiesWindow(hwnd, *state);

            ApplyTitleBarTheme(hwnd, state->theme, ::GetActiveWindow() == hwnd);
            return 0;
        }
        case WM_SIZE:
        {
            if (state)
            {
                LayoutItemPropertiesWindow(hwnd, *state);
            }
            return 0;
        }
        case WM_DPICHANGED:
        {
            if (state)
            {
                state->dpi = static_cast<UINT>(wParam);
                LayoutItemPropertiesWindow(hwnd, *state);
            }

            if (const auto* rc = reinterpret_cast<const RECT*>(lParam))
            {
                ::SetWindowPos(
                    hwnd, nullptr, rc->left, rc->top, std::max(0l, rc->right - rc->left), std::max(0l, rc->bottom - rc->top), SWP_NOZORDER | SWP_NOACTIVATE);
            }
            return 0;
        }
        case WM_ACTIVATE:
        {
            if (state)
            {
                ApplyTitleBarTheme(hwnd, state->theme, wParam != FALSE);
            }
            return 0;
        }
        case WM_GETMINMAXINFO:
        {
            auto* info = reinterpret_cast<MINMAXINFO*>(lParam);
            if (info)
            {
                static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(hwnd, *info));
            }
            return 0;
        }
        case WM_ERASEBKGND:
        {
            if (state && state->backgroundBrush)
            {
                RECT rc{};
                if (::GetClientRect(hwnd, &rc) != 0)
                {
                    ::FillRect(reinterpret_cast<HDC>(wParam), &rc, state->backgroundBrush.get());
                    return TRUE;
                }
            }
            break;
        }
        case WM_DRAWITEM:
        {
            if (state)
            {
                const auto* dis = reinterpret_cast<const DRAWITEMSTRUCT*>(lParam);
                if (dis && dis->CtlType == ODT_BUTTON)
                {
                    ThemedControls::DrawThemedPushButton(*dis, state->theme);
                    return TRUE;
                }
            }
            break;
        }
        case WM_COMMAND:
        {
            if (LOWORD(wParam) == kItemPropertiesCloseId)
            {
                ::DestroyWindow(hwnd);
                return 0;
            }
            break;
        }
        case WM_CLOSE:
        {
            ::DestroyWindow(hwnd);
            return 0;
        }
        case WM_NCDESTROY:
        {
            auto* toDelete = state;
            ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            if (toDelete && toDelete->settings)
            {
                WindowPlacementPersistence::Save(*toDelete->settings, kItemPropertiesWindowId, hwnd);

                const Common::Settings::Settings settingsToSave = SettingsSave::PrepareForSave(*toDelete->settings);
                const HRESULT saveHr                            = Common::Settings::SaveSettings(kSettingsAppId, settingsToSave);
                if (FAILED(saveHr))
                {
                    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kSettingsAppId);
                    Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
                }
            }
            delete toDelete;
            return 0;
        }
    }

    return ::DefWindowProcW(hwnd, msg, wParam, lParam);
}

[[nodiscard]] bool EnsureItemPropertiesWindowClassRegistered() noexcept
{
    static const bool registered = []
    {
        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = ItemPropertiesWndProc;
        wc.hInstance     = ::GetModuleHandleW(nullptr);
        wc.hCursor       = ::LoadCursorW(nullptr, IDC_ARROW);
        wc.hIcon         = ::LoadIconW(wc.hInstance, MAKEINTRESOURCEW(IDI_REDSALAMANDER));
        wc.hIconSm       = ::LoadIconW(wc.hInstance, MAKEINTRESOURCEW(IDI_SMALL));
        wc.lpszClassName = kItemPropertiesWindowClass;

        return ::RegisterClassExW(&wc) != 0;
    }();

    return registered;
}

HRESULT ShowItemPropertiesWindow(HWND owner, Common::Settings::Settings* settings, const AppTheme& theme, ItemPropertiesDocument doc) noexcept
{
    if (! EnsureItemPropertiesWindowClassRegistered())
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_PROPERTIES);

    auto state      = std::make_unique<ItemPropertiesWindowState>(theme, std::move(doc));
    state->settings = settings;
    auto* raw       = state.get();

    const UINT dpi = owner ? ::GetDpiForWindow(owner) : USER_DEFAULT_SCREEN_DPI;
    const int w    = ThemedControls::ScaleDip(dpi, 720);
    const int h    = ThemedControls::ScaleDip(dpi, 520);

    RECT ownerRc{};
    if (owner && ::GetWindowRect(owner, &ownerRc) == 0)
    {
        owner = nullptr;
    }

    int x = CW_USEDEFAULT;
    int y = CW_USEDEFAULT;
    if (owner)
    {
        const int ownerW = std::max(0l, ownerRc.right - ownerRc.left);
        const int ownerH = std::max(0l, ownerRc.bottom - ownerRc.top);
        x                = ownerRc.left + std::max(0, (ownerW - w) / 2);
        y                = ownerRc.top + std::max(0, (ownerH - h) / 2);
    }

    HWND hwnd =
        ::CreateWindowExW(0, kItemPropertiesWindowClass, caption.c_str(), WS_OVERLAPPEDWINDOW, x, y, w, h, nullptr, nullptr, ::GetModuleHandleW(nullptr), raw);
    if (! hwnd)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    static_cast<void>(state.release());
    const int showCmd = settings ? WindowPlacementPersistence::Restore(*settings, kItemPropertiesWindowId, hwnd) : SW_SHOWNORMAL;
    ShowWindow(hwnd, showCmd);
    return S_OK;
}
} // namespace

HRESULT FolderWindow::ShowItemPropertiesFromFolderView(Pane pane, std::filesystem::path path) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.fileSystem)
    {
        return E_POINTER;
    }

    if (path.empty())
    {
        return E_INVALIDARG;
    }

    // Win32 filesystem: use the shell property sheet for maximum detail.
    if (CompareStringOrdinal(state.pluginId.c_str(), -1, L"builtin/file-system", -1, TRUE) == CSTR_EQUAL)
    {
        ::SHObjectProperties(_hWnd.get(), SHOP_FILEPATH, path.c_str(), nullptr);
        return S_OK;
    }

    wil::com_ptr<IFileSystemIO> io;
    const HRESULT hrQI = state.fileSystem->QueryInterface(IID_PPV_ARGS(io.addressof()));
    if (FAILED(hrQI) || ! io)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const char* jsonUtf8  = nullptr;
    const HRESULT hrProps = io->GetItemProperties(path.c_str(), &jsonUtf8);
    if (FAILED(hrProps))
    {
        return hrProps;
    }

    if (! jsonUtf8 || jsonUtf8[0] == '\0')
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const std::optional<ItemPropertiesDocument> doc = TryParseItemPropertiesJson(std::string_view(jsonUtf8));
    if (! doc.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    return ShowItemPropertiesWindow(_hWnd.get(), _settings, _theme, doc.value());
}
