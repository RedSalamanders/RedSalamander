#include "FolderWindow.h"

#include "AppTheme.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "Resource.h"
#include "WindowMaximizeBehavior.h"
#include "WindowPlacementPersistence.h"

#include "Helpers.h"
#include "SettingsHotReload.h"

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/resource.h>
#pragma warning(pop)

#pragma warning(push)
#pragma warning(disable : 6297 28182) // yyjson warnings
#include <yyjson.h>
#pragma warning(pop)

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
constexpr wchar_t kItemPropertiesWindowId[]    = L"ItemPropertiesWindow";
constexpr wchar_t kSettingsAppId[]             = L"RedSalamander";

[[nodiscard]] std::wstring BuildItemPropertiesText(const ItemPropertiesDocument& doc) noexcept
{
    std::wstring text;
    if (! doc.title.empty())
    {
        text.append(doc.title);
        text.append(L"\r\n\r\n");
    }

    for (size_t sectionIndex = 0; sectionIndex < doc.sections.size(); ++sectionIndex)
    {
        const auto& section = doc.sections[sectionIndex];
        if (! section.title.empty())
        {
            text.append(section.title);
            text.append(L"\r\n");
        }

        for (const auto& field : section.fields)
        {
            text.append(field.key);
            text.append(L": ");
            text.append(field.value);
            text.append(L"\r\n");
        }

        if (sectionIndex + 1u < doc.sections.size())
        {
            text.append(L"\r\n");
        }
    }

    return text;
}

[[nodiscard]] size_t CountItemPropertiesFields(const ItemPropertiesDocument& doc) noexcept
{
    size_t count = 0u;
    for (const auto& section : doc.sections)
    {
        count += section.fields.size();
    }
    return count;
}

#ifdef ENABLE_TESTS
[[nodiscard]] size_t CountVisibleItemPropertiesChildWindows(HWND hwnd) noexcept
{
    if (! hwnd || ::IsWindow(hwnd) == FALSE)
    {
        return 0u;
    }

    size_t count = 0u;
    ::EnumChildWindows(hwnd,
                       [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& countRef = *reinterpret_cast<size_t*>(lParam);
        if (::IsWindowVisible(child) != FALSE)
        {
            ++countRef;
        }
        return TRUE;
    },
                       reinterpret_cast<LPARAM>(&count));
    return count;
}
#endif

[[nodiscard]] bool EnsureItemPropertiesWindowClassRegistered() noexcept;

class ItemPropertiesWindow final
{
public:
    ItemPropertiesWindow(Common::Settings::Settings* settings, const AppTheme& theme, ItemPropertiesDocument doc) noexcept
        : _settings(settings),
          _theme(theme),
          _doc(std::move(doc)),
          _contentText(BuildItemPropertiesText(_doc)),
          _fieldCount(CountItemPropertiesFields(_doc))
    {
    }

    ItemPropertiesWindow(const ItemPropertiesWindow&)            = delete;
    ItemPropertiesWindow& operator=(const ItemPropertiesWindow&) = delete;
    ItemPropertiesWindow(ItemPropertiesWindow&&)                 = delete;
    ItemPropertiesWindow& operator=(ItemPropertiesWindow&&)      = delete;

    [[nodiscard]] HRESULT CreateAndShow(HWND owner) noexcept
    {
        if (! EnsureItemPropertiesWindowClassRegistered())
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        _ownerWindow        = (owner && ::IsWindow(owner) != FALSE) ? owner : nullptr;
        _restoreFocusWindow = nullptr;
        if (_ownerWindow)
        {
            const HWND focused = ::GetFocus();
            if (focused && ::IsWindow(focused) != FALSE && (focused == _ownerWindow || ::IsChild(_ownerWindow, focused) != FALSE))
            {
                _restoreFocusWindow = focused;
            }
        }

        const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_PROPERTIES);
        const UINT dpi             = owner ? ::GetDpiForWindow(owner) : USER_DEFAULT_SCREEN_DPI;
        const int w                = UiMetrics::ScaleDip(dpi, 720);
        const int h                = UiMetrics::ScaleDip(dpi, 520);

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

        const HWND hwnd = ::CreateWindowExW(0,
                                            kItemPropertiesWindowClass,
                                            caption.c_str(),
                                            WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                            x,
                                            y,
                                            w,
                                            h,
                                            owner,
                                            nullptr,
                                            ::GetModuleHandleW(nullptr),
                                            this);
        if (! hwnd)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        const int showCmd = _settings ? WindowPlacementPersistence::Restore(*_settings, kItemPropertiesWindowId, hwnd) : SW_SHOWNORMAL;
        ::ShowWindow(hwnd, showCmd);
        return S_OK;
    }

    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hWnd.get();
    }

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) noexcept
    {
        ItemPropertiesWindow* self = reinterpret_cast<ItemPropertiesWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (msg == WM_NCCREATE)
        {
            const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
            self               = create ? static_cast<ItemPropertiesWindow*>(create->lpCreateParams) : nullptr;
            ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            if (self)
            {
                self->_hWnd.reset(hwnd);
            }
        }

        if (! self)
        {
            return ::DefWindowProcW(hwnd, msg, wParam, lParam);
        }

        ++self->_dispatchDepth;
        const auto finishDispatch = wil::scope_exit([self]() noexcept
        {
            if (self->_dispatchDepth > 0u)
            {
                --self->_dispatchDepth;
            }
            if (self->_dispatchDepth == 0u && self->_deletePending)
            {
                delete self;
            }
        });

        return self->WindowProc(hwnd, msg, wParam, lParam);
    }

#ifdef ENABLE_TESTS
    [[nodiscard]] bool DebugGetSnapshot(ItemPropertiesWindowDebugSnapshot& out) const noexcept
    {
        const HWND hwnd             = _hWnd.get();
        out.usesDxUiHost            = hwnd && ::IsWindow(hwnd) != FALSE;
        out.visibleChildWindowCount = hwnd ? CountVisibleItemPropertiesChildWindows(hwnd) : 0u;
        out.sectionCount            = _doc.sections.size();
        out.fieldCount              = _fieldCount;
        RedSalamander::DxUi::TextFieldDebugMultilineState multilineState{};
        if (_body && _body->DebugGetMultilineState(_dxHost, multilineState))
        {
            out.bodyFirstVisibleLine    = multilineState.firstVisibleLine;
            out.bodyVisibleLineCount    = multilineState.visibleLineCount;
            out.bodyTotalLineCount      = multilineState.totalLineCount;
            out.bodyCanScrollVertically = multilineState.canScrollVertically;
        }
        out.renderCount        = _dxHost.DebugGetRenderCount();
        out.resizeCount        = _dxHost.DebugGetResizeCount();
        out.resizeFailureCount = _dxHost.DebugGetResizeFailureCount();
        out.contentText        = _contentText;
        return out.usesDxUiHost;
    }

    [[nodiscard]] bool DebugScrollByWheelDetents(int detents) noexcept
    {
        const HWND hwnd = _hWnd.get();
        if (! hwnd || ::IsWindow(hwnd) == FALSE || ! _body)
        {
            return false;
        }

        if (::IsIconic(hwnd))
        {
            ::ShowWindow(hwnd, SW_RESTORE);
        }

        const float wheelDelta = detents > 0 ? static_cast<float>(WHEEL_DELTA) : -static_cast<float>(WHEEL_DELTA);
        const int stepCount    = detents > 0 ? detents : -detents;
        for (int remaining = stepCount; remaining > 0; --remaining)
        {
            if (! _body->OnMouseWheel(_dxHost, D2D1::Point2F(0.0f, 0.0f), wheelDelta, 0u))
            {
                return false;
            }
        }
        return true;
    }
#endif

private:
    [[nodiscard]] LRESULT WindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) noexcept
    {
        bool handled           = false;
        const LRESULT dxResult = _dxHost.HandleMessage(hwnd, msg, wParam, lParam, handled);
        if (msg == WM_NCDESTROY)
        {
            OnNcDestroy(hwnd);
            if (handled)
            {
                return dxResult;
            }
        }
        if (handled)
        {
            return dxResult;
        }

        switch (msg)
        {
            case WM_CREATE: return OnCreate(hwnd);
            case WM_SIZE: return 0;
            case WM_WINDOWPOSCHANGED:
            {
                const auto* windowPos = reinterpret_cast<const WINDOWPOS*>(lParam);
                if (windowPos && (windowPos->flags & SWP_NOSIZE) == 0)
                {
                    LayoutControls();
                }
                return ::DefWindowProcW(hwnd, msg, wParam, lParam);
            }
            case WM_DPICHANGED: return OnDpiChanged(hwnd, wParam, lParam);
            case WM_ACTIVATE: ApplyTitleBarTheme(hwnd, _theme, wParam != FALSE); return 0;
            case WM_GETMINMAXINFO:
            {
                if (auto* info = reinterpret_cast<MINMAXINFO*>(lParam))
                {
                    static_cast<void>(WindowMaximizeBehavior::ApplyVerticalMaximize(hwnd, *info));
                }
                return 0;
            }
            case WM_ERASEBKGND:
            {
                RECT rc{};
                if (::GetClientRect(hwnd, &rc) != 0)
                {
                    ::FillRect(reinterpret_cast<HDC>(wParam), &rc, _backgroundBrush.get());
                    return TRUE;
                }
                return FALSE;
            }
            case WM_CLOSE: ::DestroyWindow(hwnd); return 0;
        }

        return ::DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    [[nodiscard]] LRESULT OnCreate(HWND hwnd) noexcept
    {
        _dpi = ::GetDpiForWindow(hwnd);
        _backgroundBrush.reset(::CreateSolidBrush(_theme.windowBackground));
        if (! _dxHost.Attach(hwnd))
        {
            Debug::Error(L"ItemProperties: failed to attach DxUi host.");
            return -1;
        }
        _dxHost.SetTheme(MakeAppThemeDxPalette(_theme, _theme.windowBackground));
        BuildUi();
        LayoutControls();
        ApplyWindowChromeTheme(hwnd, _theme, WindowBackdropTarget::Tool, ::GetActiveWindow() == hwnd);
        return 0;
    }

    [[nodiscard]] LRESULT OnDpiChanged(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
    {
        _dpi = static_cast<UINT>(wParam);
        LayoutControls();
        if (const auto* rc = reinterpret_cast<const RECT*>(lParam))
        {
            ::SetWindowPos(
                hwnd, nullptr, rc->left, rc->top, std::max(0l, rc->right - rc->left), std::max(0l, rc->bottom - rc->top), SWP_NOZORDER | SWP_NOACTIVATE);
        }
        return 0;
    }

    void BuildUi()
    {
        auto root = std::make_unique<RedSalamander::DxUi::Panel>();
        _root     = root.get();

        _body = root->AddChild<RedSalamander::DxUi::TextField>(_contentText);
        _body->SetMultiline(true);
        _body->SetReadOnly(true);

        _closeButton = root->AddChild<RedSalamander::DxUi::Button>(LoadStringResource(nullptr, IDS_PROPERTIES_BTN_CLOSE));
        _closeButton->SetPrimary(true);
        _closeButton->SetMnemonic(L'o');
        _closeButton->SetOnClick([this]()
        {
            if (const HWND hwnd = _hWnd.get(); hwnd && ::IsWindow(hwnd) != FALSE)
            {
                ::PostMessageW(hwnd, WM_CLOSE, 0, 0);
            }
        });

        _dxHost.SetRoot(std::move(root));
        _dxHost.SetDefaultButton(_closeButton);
        _dxHost.SetCancelButton(_closeButton);
        // Warm the hidden native text bridge once so the read-only multiline body
        // exposes readable UIA text even though final keyboard focus stays on Close.
        _dxHost.SetFocusControl(_body);
        _dxHost.SetFocusControl(_closeButton);
    }

    void LayoutControls() noexcept
    {
        if (! _root || ! _body || ! _closeButton)
        {
            return;
        }

        const D2D1_RECT_F client = _dxHost.GetClientBoundsDip();
        const float margin       = static_cast<float>(UiMetrics::ScaleDip(_dpi, 12));
        const float buttonH      = static_cast<float>(UiMetrics::ScaleDip(_dpi, 30));
        const float buttonW      = static_cast<float>((std::max)(UiMetrics::ScaleDip(_dpi, 100), UiMetrics::ScaleDip(_dpi, 92)));

        _root->SetBounds(client);

        const float buttonTop = (std::max)(margin, client.bottom - margin - buttonH);
        _closeButton->SetBounds(D2D1::RectF(client.right - margin - buttonW, buttonTop, client.right - margin, buttonTop + buttonH));
        _body->SetBounds(D2D1::RectF(client.left + margin, client.top + margin, client.right - margin, buttonTop - margin));
        _dxHost.Invalidate();
    }

    void OnNcDestroy(HWND hwnd) noexcept
    {
        ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);

        if (_settings)
        {
            WindowPlacementPersistence::Save(*_settings, kItemPropertiesWindowId, hwnd);
            const HRESULT saveHr = SettingsHotReload::SaveSettingsAndSchema(kSettingsAppId, *_settings);
            if (FAILED(saveHr))
            {
                const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(kSettingsAppId);
                Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(saveHr), settingsPath.wstring());
            }
        }

        if (_ownerWindow && ::IsWindow(_ownerWindow) != FALSE)
        {
            static_cast<void>(::SetActiveWindow(_ownerWindow));

            const HWND restoreFocus = (_restoreFocusWindow && ::IsWindow(_restoreFocusWindow) != FALSE &&
                                       (_restoreFocusWindow == _ownerWindow || ::IsChild(_ownerWindow, _restoreFocusWindow) != FALSE))
                                          ? _restoreFocusWindow
                                          : _ownerWindow;
            static_cast<void>(::SetFocus(restoreFocus));
        }

        _hWnd.release();
        _dxHost.Detach();
        _deletePending = true;
        if (_dispatchDepth == 0u)
        {
            delete this;
        }
    }

    wil::unique_hwnd _hWnd;
    HWND _ownerWindow                     = nullptr;
    HWND _restoreFocusWindow              = nullptr;
    Common::Settings::Settings* _settings = nullptr;
    AppTheme _theme{};
    ItemPropertiesDocument _doc;
    std::wstring _contentText;
    size_t _fieldCount    = 0u;
    size_t _dispatchDepth = 0u;
    bool _deletePending   = false;
    UINT _dpi             = USER_DEFAULT_SCREEN_DPI;
    wil::unique_hbrush _backgroundBrush;
    RedSalamander::DxUi::WindowHost _dxHost;
    RedSalamander::DxUi::Panel* _root         = nullptr;
    RedSalamander::DxUi::TextField* _body     = nullptr;
    RedSalamander::DxUi::Button* _closeButton = nullptr;
};

[[nodiscard]] bool EnsureItemPropertiesWindowClassRegistered() noexcept
{
    static const bool registered = []
    {
        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpfnWndProc   = ItemPropertiesWindow::WndProc;
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
    auto* window = new (std::nothrow) ItemPropertiesWindow(settings, theme, std::move(doc));
    if (! window)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = window->CreateAndShow(owner);
    if (FAILED(hr))
    {
        delete window;
    }
    return hr;
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

#ifdef ENABLE_TESTS
HWND GetItemPropertiesWindowHandle() noexcept
{
    const HWND hwnd = ::FindWindowW(kItemPropertiesWindowClass, nullptr);
    return (hwnd && ::IsWindow(hwnd) != FALSE) ? hwnd : nullptr;
}

bool DebugGetItemPropertiesWindowSnapshot(ItemPropertiesWindowDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetItemPropertiesWindowHandle();
    if (! hwnd)
    {
        return false;
    }

    auto* window = reinterpret_cast<ItemPropertiesWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugGetSnapshot(out) : false;
}

bool DebugScrollItemPropertiesWindowByWheelDetents(int detents) noexcept
{
    const HWND hwnd = GetItemPropertiesWindowHandle();
    if (! hwnd)
    {
        return false;
    }

    auto* window = reinterpret_cast<ItemPropertiesWindow*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    return window ? window->DebugScrollByWheelDetents(detents) : false;
}
#endif
