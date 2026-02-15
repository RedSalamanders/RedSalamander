#include "ViewerWeb.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <cwctype>
#include <filesystem>
#include <format>
#include <iterator>
#include <limits>
#include <memory>
#include <new>
#include <optional>
#include <string_view>
#include <thread>
#include <type_traits>
#include <utility>
#include <vector>

#include <commctrl.h>
#include <commdlg.h>
#include <dwmapi.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <uxtheme.h>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#pragma comment(lib, "comctl32")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "uxtheme")
#pragma comment(lib, "WebView2Loader.dll.lib")

#include "Helpers.h"
#include "WindowMessages.h"

#include "FluentIcons.h"

#include "resource.h"

extern HINSTANCE g_hInstance;

namespace
{
constexpr UINT kAsyncLoadCompleteMessage = WndMsg::kViewerWebAsyncLoadComplete;
constexpr int kHeaderHeightDip           = 28;

static const int kViewerWebModuleAnchor = 0;

constexpr UINT_PTR kFileComboEscCloseSubclassId = 1u;

LRESULT CALLBACK
FileComboEscCloseSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, [[maybe_unused]] UINT_PTR subclassId, [[maybe_unused]] DWORD_PTR refData) noexcept
{
    if (msg == WM_KEYDOWN && wp == VK_ESCAPE)
    {
        const bool dropped = SendMessageW(hwnd, CB_GETDROPPEDSTATE, 0, 0) != 0;
        if (! dropped)
        {
            const HWND root = GetAncestor(hwnd, GA_ROOT);
            if (root)
            {
                PostMessageW(root, WM_CLOSE, 0, 0);
            }
            return 0;
        }
    }

    return DefSubclassProc(hwnd, msg, wp, lp);
}

void InstallFileComboEscClose(HWND combo) noexcept
{
    if (! combo)
    {
        return;
    }

    static_cast<void>(SetWindowSubclass(combo, FileComboEscCloseSubclassProc, kFileComboEscCloseSubclassId, 0));
}

using unique_yyjson_doc     = wil::unique_any<yyjson_doc*, decltype(&yyjson_doc_free), yyjson_doc_free>;
using unique_yyjson_mut_doc = wil::unique_any<yyjson_mut_doc*, decltype(&yyjson_mut_doc_free), yyjson_mut_doc_free>;
using unique_malloc_string  = wil::unique_any<char*, decltype(&::free), ::free>;

template <typename TInterface, typename TFn, typename... TInvokeArgs> class ComCallback final : public TInterface
{
public:
    explicit ComCallback(TFn fn) noexcept : _fn(std::move(fn))
    {
    }

    ComCallback(const ComCallback&)            = delete;
    ComCallback& operator=(const ComCallback&) = delete;
    ComCallback(ComCallback&&)                 = delete;
    ComCallback& operator=(ComCallback&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;

        if (riid == __uuidof(IUnknown) || riid == __uuidof(TInterface))
        {
            *ppvObject = static_cast<TInterface*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG refs = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (refs == 0)
        {
            delete this;
        }
        return refs;
    }

    HRESULT STDMETHODCALLTYPE Invoke(TInvokeArgs... args) noexcept override
    {
        return _fn(args...);
    }

private:
    std::atomic<ULONG> _refCount{1};
    TFn _fn;
};

template <typename TInterface, typename... TInvokeArgs, typename TFn> wil::com_ptr<TInterface> MakeComCallback(TFn&& fn) noexcept
{
    using FnT      = std::decay_t<TFn>;
    auto* callback = new (std::nothrow) ComCallback<TInterface, FnT, TInvokeArgs...>(FnT(std::forward<TFn>(fn)));

    wil::com_ptr<TInterface> out;
    if (callback)
    {
        out.attach(callback);
    }
    return out;
}

int PxFromDip(int dip, UINT dpi) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
}

COLORREF ColorRefFromArgb(uint32_t argb) noexcept
{
    const BYTE r = static_cast<BYTE>((argb >> 16) & 0xFFu);
    const BYTE g = static_cast<BYTE>((argb >> 8) & 0xFFu);
    const BYTE b = static_cast<BYTE>(argb & 0xFFu);
    return RGB(r, g, b);
}

COLORREF BlendColor(COLORREF under, COLORREF over, uint8_t alpha) noexcept
{
    const uint32_t inv = static_cast<uint32_t>(255u - alpha);

    const uint32_t ur = static_cast<uint32_t>(GetRValue(under));
    const uint32_t ug = static_cast<uint32_t>(GetGValue(under));
    const uint32_t ub = static_cast<uint32_t>(GetBValue(under));

    const uint32_t or_ = static_cast<uint32_t>(GetRValue(over));
    const uint32_t og  = static_cast<uint32_t>(GetGValue(over));
    const uint32_t ob  = static_cast<uint32_t>(GetBValue(over));

    const uint8_t r = static_cast<uint8_t>((ur * inv + or_ * static_cast<uint32_t>(alpha)) / 255u);
    const uint8_t g = static_cast<uint8_t>((ug * inv + og * static_cast<uint32_t>(alpha)) / 255u);
    const uint8_t b = static_cast<uint8_t>((ub * inv + ob * static_cast<uint32_t>(alpha)) / 255u);
    return RGB(r, g, b);
}

COLORREF ContrastingTextColor(COLORREF background) noexcept
{
    const uint32_t r    = static_cast<uint32_t>(GetRValue(background));
    const uint32_t g    = static_cast<uint32_t>(GetGValue(background));
    const uint32_t b    = static_cast<uint32_t>(GetBValue(background));
    const uint32_t luma = (r * 299u + g * 587u + b * 114u) / 1000u;
    return luma < 128u ? RGB(255, 255, 255) : RGB(0, 0, 0);
}

COLORREF ColorFromHSV(float hDegrees, float s, float v) noexcept
{
    const float h = std::fmod(std::fmod(hDegrees, 360.0f) + 360.0f, 360.0f);
    const float c = v * s;
    const float x = c * (1.0f - std::fabs(std::fmod(h / 60.0f, 2.0f) - 1.0f));
    const float m = v - c;

    float rf = 0.0f;
    float gf = 0.0f;
    float bf = 0.0f;
    if (h < 60.0f)
    {
        rf = c;
        gf = x;
        bf = 0.0f;
    }
    else if (h < 120.0f)
    {
        rf = x;
        gf = c;
        bf = 0.0f;
    }
    else if (h < 180.0f)
    {
        rf = 0.0f;
        gf = c;
        bf = x;
    }
    else if (h < 240.0f)
    {
        rf = 0.0f;
        gf = x;
        bf = c;
    }
    else if (h < 300.0f)
    {
        rf = x;
        gf = 0.0f;
        bf = c;
    }
    else
    {
        rf = c;
        gf = 0.0f;
        bf = x;
    }

    const auto toByte = [](float v01) noexcept
    {
        const float scaled = std::clamp(v01 * 255.0f, 0.0f, 255.0f);
        return static_cast<BYTE>(std::lround(scaled));
    };

    const BYTE r = toByte(rf + m);
    const BYTE g = toByte(gf + m);
    const BYTE b = toByte(bf + m);
    return RGB(r, g, b);
}

uint32_t StableHash32(std::wstring_view text) noexcept
{
    // FNV-1a
    uint32_t hash = 2166136261u;
    for (const wchar_t ch : text)
    {
        hash ^= static_cast<uint32_t>(ch);
        hash *= 16777619u;
    }
    return hash;
}

COLORREF ResolveAccentColor(const ViewerTheme& theme, std::wstring_view seed) noexcept
{
    if (theme.rainbowMode)
    {
        const uint32_t h = StableHash32(seed);
        const float hue  = static_cast<float>(h % 360u);
        const float sat  = theme.darkBase ? 0.70f : 0.55f;
        const float val  = theme.darkBase ? 0.95f : 0.85f;
        return ColorFromHSV(hue, sat, val);
    }

    return ColorRefFromArgb(theme.accentArgb);
}

std::wstring LeafNameFromPath(std::wstring_view path)
{
    if (path.empty())
    {
        return {};
    }

    const size_t slash = path.find_last_of(L"\\/");
    if (slash == std::wstring_view::npos)
    {
        return std::wstring(path);
    }

    return std::wstring(path.substr(slash + 1));
}

std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return result;
}

bool StartsWithNoCase(std::wstring_view value, std::wstring_view prefix) noexcept
{
    if (value.size() < prefix.size())
    {
        return false;
    }

    const int len = static_cast<int>(prefix.size());
    return CompareStringOrdinal(value.data(), len, prefix.data(), len, TRUE) == CSTR_EQUAL;
}

// Forward declarations for file-scope helpers defined later in this file.
bool CopyUnicodeTextToClipboard(HWND hwnd, std::wstring_view text) noexcept;
[[nodiscard]] bool IsProbablyWin32Path(std::wstring_view path) noexcept;
std::optional<std::filesystem::path> ShowSaveAsDialog(HWND hwnd, std::wstring_view suggestedFileName) noexcept;
[[nodiscard]] std::wstring EscapeJavaScriptString(std::wstring_view text) noexcept;

struct ViewerWebClassBackgroundBrushState
{
    ViewerWebClassBackgroundBrushState()                                                     = default;
    ViewerWebClassBackgroundBrushState(const ViewerWebClassBackgroundBrushState&)            = delete;
    ViewerWebClassBackgroundBrushState& operator=(const ViewerWebClassBackgroundBrushState&) = delete;
    ViewerWebClassBackgroundBrushState(ViewerWebClassBackgroundBrushState&&)                 = delete;
    ViewerWebClassBackgroundBrushState& operator=(ViewerWebClassBackgroundBrushState&&)      = delete;

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> activeBrush;
    COLORREF activeColor = CLR_INVALID;

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> pendingBrush;
    COLORREF pendingColor = CLR_INVALID;

    bool classRegistered = false;
};

ViewerWebClassBackgroundBrushState g_viewerWebClassBackgroundBrush;

HBRUSH GetActiveViewerWebClassBackgroundBrush() noexcept
{
    if (! g_viewerWebClassBackgroundBrush.activeBrush)
    {
        const COLORREF sys = GetSysColor(COLOR_WINDOW);
        g_viewerWebClassBackgroundBrush.activeBrush.reset(CreateSolidBrush(sys));
        g_viewerWebClassBackgroundBrush.activeColor = sys;
    }
    return g_viewerWebClassBackgroundBrush.activeBrush.get();
}

void RequestViewerWebClassBackgroundColor(COLORREF color) noexcept
{
    if (color == CLR_INVALID)
    {
        return;
    }

    if (g_viewerWebClassBackgroundBrush.activeBrush && g_viewerWebClassBackgroundBrush.activeColor == color)
    {
        return;
    }

    if (g_viewerWebClassBackgroundBrush.pendingBrush && g_viewerWebClassBackgroundBrush.pendingColor == color)
    {
        return;
    }

    g_viewerWebClassBackgroundBrush.pendingBrush.reset(CreateSolidBrush(color));
    g_viewerWebClassBackgroundBrush.pendingColor = color;
}

void ApplyPendingViewerWebClassBackgroundBrush(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (! g_viewerWebClassBackgroundBrush.classRegistered)
    {
        return;
    }

    if (! g_viewerWebClassBackgroundBrush.pendingBrush)
    {
        return;
    }

    g_viewerWebClassBackgroundBrush.activeBrush  = std::move(g_viewerWebClassBackgroundBrush.pendingBrush);
    g_viewerWebClassBackgroundBrush.activeColor  = g_viewerWebClassBackgroundBrush.pendingColor;
    g_viewerWebClassBackgroundBrush.pendingColor = CLR_INVALID;

    SetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND, reinterpret_cast<LONG_PTR>(g_viewerWebClassBackgroundBrush.activeBrush.get()));
}

std::string ResourceBytesToString(HINSTANCE hinst, UINT id)
{
    HRSRC res = FindResourceW(hinst, MAKEINTRESOURCEW(id), RT_RCDATA);
    if (! res)
    {
        return {};
    }

    const DWORD size = SizeofResource(hinst, res);
    if (size == 0)
    {
        return {};
    }

    HGLOBAL loaded = LoadResource(hinst, res);
    if (! loaded)
    {
        return {};
    }

    const void* bytes = LockResource(loaded);
    if (! bytes)
    {
        return {};
    }

    return std::string(reinterpret_cast<const char*>(bytes), reinterpret_cast<const char*>(bytes) + size);
}

std::wstring UrlFromFilePath(std::wstring_view path)
{
    if (path.empty())
    {
        return {};
    }

    const std::wstring pathCopy(path);
    DWORD capacity = static_cast<DWORD>(pathCopy.size() * 3 + 64);
    capacity       = std::max<DWORD>(capacity, 256u);

    for (int attempt = 0; attempt < 4; ++attempt)
    {
        std::wstring url(static_cast<size_t>(capacity), L'\0');
        DWORD written = capacity;

        const HRESULT hr = UrlCreateFromPathW(pathCopy.c_str(), url.data(), &written, 0);
        if (SUCCEEDED(hr))
        {
            if (! url.empty())
            {
                const size_t len = wcsnlen(url.c_str(), url.size());
                url.resize(len);
            }
            return url;
        }

        if ((hr == E_POINTER || hr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER)) && written > capacity)
        {
            capacity = written;
            continue;
        }

        break;
    }

    return {};
}

constexpr char kViewerWebSchemaJson[] = R"json({
    "version": 1,
    "title": "Web Viewer",
    "fields": [
        {
            "key": "allowExternalNavigation",
            "type": "option",
            "label": "External navigation",
            "description": "Allow navigating to http/https links (Web/Markdown).",
            "default": "1",
            "options": [
                { "value": "0", "label": "Block" },
                { "value": "1", "label": "Allow" }
            ]
        },
        {
            "key": "devToolsEnabled",
            "type": "option",
            "label": "DevTools",
            "description": "Allow opening DevTools for the viewer WebView2 instance.",
            "default": "0",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        }
    ]
})json";

constexpr char kViewerJsonSchemaJson[] = R"json({
    "version": 1,
    "title": "JSON Viewer",
    "fields": [
        {
            "key": "maxDocumentMiB",
            "type": "value",
            "label": "Max document size (MiB)",
            "description": "Maximum size for in-memory loads.",
            "default": 32,
            "min": 1,
            "max": 512
        },
        {
            "key": "viewMode",
            "type": "option",
            "label": "View mode",
            "description": "Pretty highlighted text or interactive tree view.",
            "default": "pretty",
            "options": [
                { "value": "pretty", "label": "Pretty" },
                { "value": "tree", "label": "Tree" }
            ]
        },
        {
            "key": "devToolsEnabled",
            "type": "option",
            "label": "DevTools",
            "description": "Allow opening DevTools for the viewer WebView2 instance.",
            "default": "0",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        }
    ]
})json";

constexpr char kViewerMarkdownSchemaJson[] = R"json({
    "version": 1,
    "title": "Markdown Viewer",
    "fields": [
        {
            "key": "maxDocumentMiB",
            "type": "value",
            "label": "Max document size (MiB)",
            "description": "Maximum size for in-memory loads.",
            "default": 32,
            "min": 1,
            "max": 512
        },
        {
            "key": "allowExternalNavigation",
            "type": "option",
            "label": "External navigation",
            "description": "Allow navigating to http/https links.",
            "default": "1",
            "options": [
                { "value": "0", "label": "Block" },
                { "value": "1", "label": "Allow" }
            ]
        },
        {
            "key": "devToolsEnabled",
            "type": "option",
            "label": "DevTools",
            "description": "Allow opening DevTools for the viewer WebView2 instance.",
            "default": "0",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        }
    ]
})json";
} // namespace

void ViewerWeb::OnCreate(HWND hwnd)
{
    const UINT dpi       = GetDpiForWindow(hwnd);
    const int uiHeightPx = -MulDiv(9, static_cast<int>(dpi), 72);

    _uiFont.reset(CreateFontW(uiHeightPx,
                              0,
                              0,
                              0,
                              FW_NORMAL,
                              FALSE,
                              FALSE,
                              FALSE,
                              DEFAULT_CHARSET,
                              OUT_DEFAULT_PRECIS,
                              CLIP_DEFAULT_PRECIS,
                              CLEARTYPE_QUALITY,
                              DEFAULT_PITCH | FF_DONTCARE,
                              L"Segoe UI"));
    if (! _uiFont)
    {
        Debug::ErrorWithLastError(L"ViewerWeb: CreateFontW failed for UI font.");
    }

    const DWORD comboStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS;
    _hFileCombo.reset(CreateWindowExW(
        0, L"COMBOBOX", nullptr, comboStyle, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_VIEWERWEB_FILE_COMBO)), g_hInstance, nullptr));
    if (! _hFileCombo)
    {
        Debug::ErrorWithLastError(L"ViewerWeb: CreateWindowExW failed for file combo.");
    }

    if (_hFileCombo && _uiFont)
    {
        SendMessageW(_hFileCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_uiFont.get()), TRUE);
    }
    if (_hFileCombo)
    {
        InstallFileComboEscClose(_hFileCombo.get());
    }

    if (_hFileCombo)
    {
        int itemHeight = PxFromDip(24, dpi);
        auto hdc       = wil::GetDC(hwnd);
        if (hdc)
        {
            HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
            static_cast<void>(oldFont);

            TEXTMETRICW tm{};
            if (GetTextMetricsW(hdc.get(), &tm) != 0)
            {
                itemHeight = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, dpi);
            }
        }

        itemHeight = std::max(itemHeight, 1);
        SendMessageW(_hFileCombo.get(), CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), static_cast<LPARAM>(itemHeight));
        SendMessageW(_hFileCombo.get(), CB_SETITEMHEIGHT, 0, static_cast<LPARAM>(itemHeight));
    }

    if (_hFileCombo)
    {
        COMBOBOXINFO info{};
        info.cbSize = sizeof(info);
        if (GetComboBoxInfo(_hFileCombo.get(), &info) != 0)
        {
            _hFileComboList = info.hwndList;
            _hFileComboItem = info.hwndItem;
        }
    }

    ApplyTheme(hwnd);
    RefreshFileCombo(hwnd);
    Layout(hwnd);
    static_cast<void>(EnsureWebView2(hwnd));
}

void ViewerWeb::OnDestroy() noexcept
{
    _hFindDialog.reset();
    DiscardWebView2();

    if (_tempExtractedPath.has_value() && ! _tempExtractedPath->empty())
    {
        std::error_code ec;
        static_cast<void>(std::filesystem::remove(*_tempExtractedPath, ec));
        _tempExtractedPath.reset();
    }

    IViewerCallback* callback = _callback;
    void* cookie              = _callbackCookie;
    if (callback)
    {
        AddRef();
        static_cast<void>(callback->ViewerClosed(cookie));
        Release();
    }
}

void ViewerWeb::OnSize(UINT /*width*/, UINT /*height*/) noexcept
{
    if (_hWnd)
    {
        Layout(_hWnd.get());
    }
}

void ViewerWeb::OnCommand(HWND hwnd, UINT commandId, UINT code, HWND /*control*/) noexcept
{
    if (commandId == IDC_VIEWERWEB_FILE_COMBO && code == CBN_SELCHANGE && _hFileCombo)
    {
        const LRESULT sel = SendMessageW(_hFileCombo.get(), CB_GETCURSEL, 0, 0);
        if (sel != CB_ERR)
        {
            const size_t index = static_cast<size_t>(sel);
            if (index < _otherFiles.size())
            {
                _otherIndex = index;
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
            }
        }
        return;
    }

    switch (commandId)
    {
        case IDM_VIEWERWEB_FILE_SAVE_AS: static_cast<void>(CommandSaveAs(hwnd)); break;
        case IDM_VIEWERWEB_FILE_REFRESH: static_cast<void>(OpenPath(hwnd, _currentPath, false)); break;
        case IDM_VIEWERWEB_FILE_EXIT: DestroyWindow(hwnd); break;

        case IDM_VIEWERWEB_OTHER_NEXT:
            if (_otherFiles.size() > 1)
            {
                _otherIndex = (_otherIndex + 1) % _otherFiles.size();
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
                RefreshFileCombo(hwnd);
            }
            break;
        case IDM_VIEWERWEB_OTHER_PREVIOUS:
            if (_otherFiles.size() > 1)
            {
                _otherIndex = (_otherIndex + _otherFiles.size() - 1) % _otherFiles.size();
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
                RefreshFileCombo(hwnd);
            }
            break;
        case IDM_VIEWERWEB_OTHER_FIRST:
            if (_otherFiles.size() > 1)
            {
                _otherIndex = 0;
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
                RefreshFileCombo(hwnd);
            }
            break;
        case IDM_VIEWERWEB_OTHER_LAST:
            if (_otherFiles.size() > 1)
            {
                _otherIndex = _otherFiles.size() - 1;
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
                RefreshFileCombo(hwnd);
            }
            break;

        case IDM_VIEWERWEB_SEARCH_FIND: CommandFind(hwnd); break;
        case IDM_VIEWERWEB_SEARCH_FIND_NEXT: CommandFindNext(hwnd); break;
        case IDM_VIEWERWEB_SEARCH_FIND_PREVIOUS: CommandFindPrevious(hwnd); break;

        case IDM_VIEWERWEB_VIEW_ZOOM_IN: CommandZoomIn(); break;
        case IDM_VIEWERWEB_VIEW_ZOOM_OUT: CommandZoomOut(); break;
        case IDM_VIEWERWEB_VIEW_ZOOM_RESET: CommandZoomReset(); break;
        case IDM_VIEWERWEB_VIEW_DEVTOOLS: CommandToggleDevTools(); break;

        case IDM_VIEWERWEB_TOOLS_COPY_URL: CommandCopyUrl(hwnd); break;
        case IDM_VIEWERWEB_TOOLS_OPEN_EXTERNAL: CommandOpenExternal(hwnd); break;
        case IDM_VIEWERWEB_TOOLS_JSON_EXPAND_ALL: CommandJsonExpandAll(); break;
        case IDM_VIEWERWEB_TOOLS_JSON_COLLAPSE_ALL: CommandJsonCollapseAll(); break;
        case IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE: CommandMarkdownToggleSource(); break;

        default: break;
    }
}

void ViewerWeb::OnKeyDown(HWND hwnd, UINT vk) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

    if (vk == VK_ESCAPE)
    {
        DestroyWindow(hwnd);
        return;
    }

    if (vk == VK_F5)
    {
        static_cast<void>(OpenPath(hwnd, _currentPath, false));
        return;
    }

    if (vk == VK_F12)
    {
        CommandToggleDevTools();
        return;
    }

    if (vk == VK_F3)
    {
        shift ? CommandFindPrevious(hwnd) : CommandFindNext(hwnd);
        return;
    }

    if (ctrl && (vk == 'F' || vk == 'f'))
    {
        CommandFind(hwnd);
        return;
    }

    if (ctrl && (vk == 'S' || vk == 's'))
    {
        static_cast<void>(CommandSaveAs(hwnd));
        return;
    }

    if (ctrl && (vk == 'L' || vk == 'l'))
    {
        CommandCopyUrl(hwnd);
        return;
    }

    if (ctrl && vk == VK_RETURN)
    {
        CommandOpenExternal(hwnd);
        return;
    }

    if (ctrl && (vk == VK_OEM_PLUS || vk == VK_ADD || vk == '='))
    {
        CommandZoomIn();
        return;
    }

    if (ctrl && (vk == VK_OEM_MINUS || vk == VK_SUBTRACT || vk == '-'))
    {
        CommandZoomOut();
        return;
    }

    if (ctrl && (vk == '0'))
    {
        CommandZoomReset();
        return;
    }

    if (ctrl && vk == VK_OEM_3)
    {
        CommandMarkdownToggleSource();
        return;
    }

    if (ctrl && vk == VK_UP)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_PREVIOUS, 0);
        return;
    }

    if (ctrl && vk == VK_DOWN)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_NEXT, 0);
        return;
    }

    if (ctrl && vk == VK_HOME)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_FIRST, 0);
        return;
    }

    if (ctrl && vk == VK_END)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_LAST, 0);
        return;
    }

    if (_kind != ViewerWebKind::Web && vk == VK_SPACE)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_NEXT, 0);
        return;
    }

    if (_kind != ViewerWebKind::Web && vk == VK_BACK)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_PREVIOUS, 0);
        return;
    }
}

void ViewerWeb::OnPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);

    FillRect(hdc.get(), &ps.rcPaint, GetActiveViewerWebClassBackgroundBrush());

    if (_headerBrush)
    {
        FillRect(hdc.get(), &_headerRect, _headerBrush.get());
    }

    if (_hasTheme)
    {
        const UINT dpi          = GetDpiForWindow(hwnd);
        const COLORREF accent   = ResolveAccentColor(_theme, _currentPath.empty() ? _metaId : _currentPath);
        const int lineThickness = std::max(1, PxFromDip(1, dpi));
        RECT line               = _headerRect;
        line.top                = std::max(line.top, line.bottom - lineThickness);
        line.bottom             = std::max(line.bottom, line.top);
        wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> brush(CreateSolidBrush(accent));
        FillRect(hdc.get(), &line, brush.get());
    }

    if (! _statusMessage.empty())
    {
        const UINT dpi    = GetDpiForWindow(hwnd);
        const int padding = PxFromDip(8, dpi);
        RECT rc           = _headerRect;
        rc.left           = std::min(rc.right, rc.left + padding);
        rc.right          = std::max(rc.left, rc.right - padding);

        SetBkMode(hdc.get(), TRANSPARENT);
        SetTextColor(hdc.get(), _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT));

        HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
        static_cast<void>(oldFont);

        DrawTextW(hdc.get(), _statusMessage.c_str(), static_cast<int>(_statusMessage.size()), &rc, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    }
}

LRESULT ViewerWeb::OnEraseBkgnd(HWND /*hwnd*/, HDC /*hdc*/) noexcept
{
    return 1;
}

void ViewerWeb::OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggested) noexcept
{
    if (suggested)
    {
        SetWindowPos(hwnd,
                     nullptr,
                     suggested->left,
                     suggested->top,
                     std::max(1L, suggested->right - suggested->left),
                     std::max(1L, suggested->bottom - suggested->top),
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

    const int uiHeightPx = -MulDiv(9, static_cast<int>(newDpi), 72);
    _uiFont.reset(CreateFontW(uiHeightPx,
                              0,
                              0,
                              0,
                              FW_NORMAL,
                              FALSE,
                              FALSE,
                              FALSE,
                              DEFAULT_CHARSET,
                              OUT_DEFAULT_PRECIS,
                              CLIP_DEFAULT_PRECIS,
                              CLEARTYPE_QUALITY,
                              DEFAULT_PITCH | FF_DONTCARE,
                              L"Segoe UI"));
    if (! _uiFont)
    {
        Debug::ErrorWithLastError(L"ViewerWeb: CreateFontW failed for UI font on DPI change.");
    }

    if (_hFileCombo && _uiFont)
    {
        SendMessageW(_hFileCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_uiFont.get()), TRUE);
    }

    ApplyTheme(hwnd);
    Layout(hwnd);
}

LRESULT ViewerWeb::OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    OnDestroy();
    static_cast<void>(DrainPostedPayloadsForWindow(hwnd));

    _hFileCombo.release();
    _hWnd.release();
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);

    Release();
    return DefWindowProcW(hwnd, WM_NCDESTROY, wp, lp);
}

void ViewerWeb::OnFindMessage(const FINDREPLACEW* findReplace) noexcept
{
    if (! findReplace)
    {
        return;
    }

    if ((findReplace->Flags & FR_DIALOGTERM) != 0)
    {
        _hFindDialog.release();
        return;
    }

    if ((findReplace->Flags & FR_FINDNEXT) == 0)
    {
        return;
    }

    if (findReplace->lpstrFindWhat)
    {
        _findQuery.assign(findReplace->lpstrFindWhat);
    }

    if (_findQuery.empty() || ! _webView)
    {
        return;
    }

    const bool backwards        = (findReplace->Flags & FR_DOWN) == 0;
    const std::wstring queryEsc = EscapeJavaScriptString(_findQuery);
    const std::wstring script   = std::format(L"(function(){{try{{return window.find('{}',false,{},true,false,true,false);}}catch(e){{return false;}}}})();",
                                            queryEsc,
                                            backwards ? L"true" : L"false");
    static_cast<void>(_webView->ExecuteScript(script.c_str(), nullptr));
}

LRESULT ViewerWeb::OnMeasureItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept
{
    if (! measure)
    {
        return FALSE;
    }

    if (measure->CtlType == ODT_MENU)
    {
        const size_t index = static_cast<size_t>(measure->itemData);
        if (index >= _menuThemeItems.size())
        {
            return TRUE;
        }

        const MenuItemData& data = _menuThemeItems[index];
        const UINT dpi           = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;

        if (data.separator)
        {
            measure->itemWidth  = 1;
            measure->itemHeight = static_cast<UINT>(MulDiv(8, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));
            return TRUE;
        }

        const UINT heightDip = data.topLevel ? 20u : 24u;
        measure->itemHeight  = static_cast<UINT>(MulDiv(static_cast<int>(heightDip), static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));

        auto hdc = wil::GetDC(hwnd);
        if (! hdc)
        {
            measure->itemWidth = 120;
            return TRUE;
        }

        HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
        static_cast<void>(oldFont);

        SIZE textSize{};
        if (! data.text.empty())
        {
            GetTextExtentPoint32W(hdc.get(), data.text.c_str(), static_cast<int>(data.text.size()), &textSize);
        }

        SIZE shortcutSize{};
        if (! data.shortcut.empty())
        {
            GetTextExtentPoint32W(hdc.get(), data.shortcut.c_str(), static_cast<int>(data.shortcut.size()), &shortcutSize);
        }

        const int dpiInt           = static_cast<int>(dpi);
        const int paddingX         = MulDiv(8, dpiInt, USER_DEFAULT_SCREEN_DPI);
        const int shortcutGap      = MulDiv(20, dpiInt, USER_DEFAULT_SCREEN_DPI);
        const int subMenuAreaWidth = data.hasSubMenu && ! data.topLevel ? MulDiv(18, dpiInt, USER_DEFAULT_SCREEN_DPI) : 0;
        const int checkAreaWidth   = data.topLevel ? 0 : MulDiv(20, dpiInt, USER_DEFAULT_SCREEN_DPI);
        const int checkGap         = data.topLevel ? 0 : MulDiv(4, dpiInt, USER_DEFAULT_SCREEN_DPI);

        int width = paddingX + checkAreaWidth + checkGap + textSize.cx + paddingX;
        if (! data.shortcut.empty())
        {
            width += shortcutGap + shortcutSize.cx;
        }
        width += subMenuAreaWidth;

        measure->itemWidth = static_cast<UINT>(std::max(width, 60));
        return TRUE;
    }

    if (measure->CtlType == ODT_COMBOBOX && measure->CtlID == IDC_VIEWERWEB_FILE_COMBO)
    {
        const UINT dpi = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;

        int height = PxFromDip(24, dpi);
        auto hdc   = wil::GetDC(hwnd);
        if (hdc)
        {
            HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
            static_cast<void>(oldFont);

            TEXTMETRICW tm{};
            if (GetTextMetricsW(hdc.get(), &tm) != 0)
            {
                height = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, dpi);
            }
        }

        measure->itemHeight = static_cast<UINT>(std::max(height, 1));
        return TRUE;
    }

    return FALSE;
}

LRESULT ViewerWeb::OnDrawItem(HWND /*hwnd*/, DRAWITEMSTRUCT* draw) noexcept
{
    if (! draw || ! draw->hDC)
    {
        return FALSE;
    }

    if (draw->CtlType == ODT_MENU)
    {
        const size_t index = static_cast<size_t>(draw->itemData);
        if (index >= _menuThemeItems.size())
        {
            return TRUE;
        }

        const MenuItemData& data = _menuThemeItems[index];
        const bool selected      = (draw->itemState & ODS_SELECTED) != 0;
        const bool disabled      = (draw->itemState & ODS_DISABLED) != 0;
        const bool checked       = (draw->itemState & ODS_CHECKED) != 0;

        const COLORREF bg             = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_MENU);
        const COLORREF fg             = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_MENUTEXT);
        const COLORREF selBg          = _hasTheme ? ColorRefFromArgb(_theme.selectionBackgroundArgb) : GetSysColor(COLOR_HIGHLIGHT);
        const COLORREF selFg          = _hasTheme ? ColorRefFromArgb(_theme.selectionTextArgb) : GetSysColor(COLOR_HIGHLIGHTTEXT);
        const COLORREF disabledFg     = _hasTheme ? BlendColor(bg, fg, 120u) : GetSysColor(COLOR_GRAYTEXT);
        const COLORREF separatorColor = _hasTheme ? BlendColor(bg, fg, 80u) : GetSysColor(COLOR_3DSHADOW);

        COLORREF fillColor = selected ? selBg : bg;
        COLORREF textColor = selected ? selFg : fg;
        if (disabled)
        {
            textColor = disabledFg;
        }

        wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> bgBrush(CreateSolidBrush(fillColor));
        FillRect(draw->hDC, &draw->rcItem, bgBrush.get());

        if (data.separator)
        {
            const int dpi      = GetDeviceCaps(draw->hDC, LOGPIXELSX);
            const int paddingX = MulDiv(6, dpi, USER_DEFAULT_SCREEN_DPI);
            const int y        = (draw->rcItem.top + draw->rcItem.bottom) / 2;
            wil::unique_any<HPEN, decltype(&::DeleteObject), ::DeleteObject> pen(CreatePen(PS_SOLID, 1, separatorColor));
            auto oldPen = wil::SelectObject(draw->hDC, pen.get());
            static_cast<void>(oldPen);
            MoveToEx(draw->hDC, draw->rcItem.left + paddingX, y, nullptr);
            LineTo(draw->hDC, draw->rcItem.right - paddingX, y);
            return TRUE;
        }

        HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        auto oldFont    = wil::SelectObject(draw->hDC, fontToUse);
        static_cast<void>(oldFont);

        SetBkMode(draw->hDC, TRANSPARENT);
        SetTextColor(draw->hDC, textColor);

        const int dpi              = GetDeviceCaps(draw->hDC, LOGPIXELSX);
        const int paddingX         = MulDiv(8, dpi, USER_DEFAULT_SCREEN_DPI);
        const int checkAreaWidth   = data.topLevel ? 0 : MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);
        const int subMenuAreaWidth = data.hasSubMenu && ! data.topLevel ? MulDiv(18, dpi, USER_DEFAULT_SCREEN_DPI) : 0;
        const int checkGap         = data.topLevel ? 0 : MulDiv(4, dpi, USER_DEFAULT_SCREEN_DPI);

        RECT textRect = draw->rcItem;
        textRect.left += paddingX + checkAreaWidth + checkGap;
        textRect.right -= paddingX + subMenuAreaWidth;

        DrawTextW(draw->hDC, data.text.c_str(), static_cast<int>(data.text.size()), &textRect, DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

        if (! data.shortcut.empty())
        {
            RECT sc  = draw->rcItem;
            sc.left  = std::min(sc.right, textRect.left + (textRect.right - textRect.left) / 2);
            sc.right = std::max(sc.left, draw->rcItem.right - paddingX - subMenuAreaWidth);
            DrawTextW(draw->hDC, data.shortcut.c_str(), static_cast<int>(data.shortcut.size()), &sc, DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_RIGHT);
        }

        if (! data.topLevel)
        {
            if (checked)
            {
                RECT checkRect = draw->rcItem;
                checkRect.left += paddingX;
                checkRect.right     = checkRect.left + checkAreaWidth;
                const wchar_t glyph = FluentIcons::kFallbackCheckMark;
                DrawTextW(draw->hDC, &glyph, 1, &checkRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            }

            if (data.hasSubMenu)
            {
                RECT arrowRect      = draw->rcItem;
                arrowRect.left      = std::max<LONG>(arrowRect.left, arrowRect.right - paddingX - subMenuAreaWidth);
                arrowRect.right     = std::max(arrowRect.right, arrowRect.left);
                const wchar_t glyph = FluentIcons::kFallbackChevronRight;
                DrawTextW(draw->hDC, &glyph, 1, &arrowRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            }
        }

        return TRUE;
    }

    if (draw->CtlType == ODT_COMBOBOX && _hFileCombo && draw->hwndItem == _hFileCombo.get())
    {
        const UINT dpi    = _hWnd ? GetDpiForWindow(_hWnd.get()) : USER_DEFAULT_SCREEN_DPI;
        const int padding = PxFromDip(6, dpi);

        const bool selected = (draw->itemState & ODS_SELECTED) != 0;
        const bool disabled = (draw->itemState & ODS_DISABLED) != 0;

        const COLORREF bgBase = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
        COLORREF headerBg     = bgBase;
        if (_hasTheme && _theme.darkMode)
        {
            headerBg = RGB(std::max(0, GetRValue(bgBase) - 10), std::max(0, GetGValue(bgBase) - 10), std::max(0, GetBValue(bgBase) - 10));
        }
        else
        {
            headerBg = RGB(std::max(0, GetRValue(bgBase) - 5), std::max(0, GetGValue(bgBase) - 5), std::max(0, GetBValue(bgBase) - 5));
        }

        const COLORREF baseFg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
        const COLORREF selBg  = _hasTheme ? ColorRefFromArgb(_theme.selectionBackgroundArgb) : GetSysColor(COLOR_HIGHLIGHT);
        const COLORREF selFg  = _hasTheme ? ColorRefFromArgb(_theme.selectionTextArgb) : GetSysColor(COLOR_HIGHLIGHTTEXT);

        COLORREF fill = selected ? selBg : headerBg;
        COLORREF text = selected ? selFg : baseFg;
        if (disabled)
        {
            text = BlendColor(fill, text, 160u);
        }

        wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> bgBrush(CreateSolidBrush(fill));
        FillRect(draw->hDC, &draw->rcItem, bgBrush.get());

        if (draw->itemID == static_cast<UINT>(-1))
        {
            return TRUE;
        }

        const int len = static_cast<int>(SendMessageW(draw->hwndItem, CB_GETLBTEXTLEN, draw->itemID, 0));
        if (len <= 0)
        {
            return TRUE;
        }

        std::wstring buf(static_cast<size_t>(len) + 1, L'\0');
        const LRESULT got = SendMessageW(draw->hwndItem, CB_GETLBTEXT, draw->itemID, reinterpret_cast<LPARAM>(buf.data()));
        if (got == CB_ERR)
        {
            return TRUE;
        }
        buf.resize(static_cast<size_t>(got));

        HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        auto oldFont    = wil::SelectObject(draw->hDC, fontToUse);
        static_cast<void>(oldFont);

        SetBkMode(draw->hDC, TRANSPARENT);
        SetTextColor(draw->hDC, text);

        RECT rc = draw->rcItem;
        rc.left += padding;
        rc.right -= padding;
        DrawTextW(draw->hDC, buf.c_str(), static_cast<int>(buf.size()), &rc, DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
        return TRUE;
    }

    return FALSE;
}

ViewerWeb::ViewerWeb(ViewerWebKind kind) noexcept : _kind(kind)
{
    switch (_kind)
    {
        case ViewerWebKind::Web:
            _metaId          = L"builtin/viewer-web";
            _metaShortId     = L"web";
            _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERWEB_NAME);
            _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERWEB_DESCRIPTION);
            break;
        case ViewerWebKind::Json:
            _metaId          = L"builtin/viewer-json";
            _metaShortId     = L"json";
            _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERJSON_NAME);
            _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERJSON_DESCRIPTION);
            break;
        case ViewerWebKind::Markdown:
            _metaId          = L"builtin/viewer-markdown";
            _metaShortId     = L"md";
            _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERMARKDOWN_NAME);
            _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERMARKDOWN_DESCRIPTION);
            break;
        default:
            _metaId          = L"builtin/viewer-web";
            _metaShortId     = L"web";
            _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERWEB_NAME);
            _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERWEB_DESCRIPTION);
            break;
    }
}

ViewerWeb::~ViewerWeb()
{
    // WebView2 cleanup is handled in OnDestroy() to allow async shutdown to complete
    // before the object is destroyed. Do not call DiscardWebView2() here.

    if (_tempExtractedPath.has_value() && ! _tempExtractedPath->empty())
    {
        std::error_code ec;
        static_cast<void>(std::filesystem::remove(*_tempExtractedPath, ec));
        _tempExtractedPath.reset();
    }
}

void ViewerWeb::SetHost(IHost* host) noexcept
{
    _host       = host;
    _hostAlerts = nullptr;

    if (! _host)
    {
        return;
    }

    wil::com_ptr<IHostAlerts> alerts;
    const HRESULT hr = _host->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
    if (SUCCEEDED(hr) && alerts)
    {
        _hostAlerts = std::move(alerts);
    }
}

HRESULT STDMETHODCALLTYPE ViewerWeb::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IViewer))
    {
        *ppvObject = static_cast<IViewer*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ViewerWeb::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE ViewerWeb::Release() noexcept
{
    const ULONG refs = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (refs == 0)
    {
        delete this;
    }
    return refs;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = nullptr;

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    switch (_kind)
    {
        case ViewerWebKind::Json: *schemaJsonUtf8 = kViewerJsonSchemaJson; break;
        case ViewerWebKind::Markdown: *schemaJsonUtf8 = kViewerMarkdownSchemaJson; break;
        case ViewerWebKind::Web:
        default: *schemaJsonUtf8 = kViewerWebSchemaJson; break;
    }
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    uint32_t maxDocumentMiB      = 32;
    bool allowExternalNavigation = true;
    bool devToolsEnabled         = false;
    JsonViewMode jsonViewMode    = JsonViewMode::Pretty;

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        const std::string_view utf8(configurationJsonUtf8);
        if (! utf8.empty())
        {
            std::string mutableJson(utf8);
            yyjson_read_err err{};
            unique_yyjson_doc doc(yyjson_read_opts(mutableJson.data(), mutableJson.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err));
            if (doc)
            {
                yyjson_val* root = yyjson_doc_get_root(doc.get());
                if (yyjson_is_obj(root))
                {
                    yyjson_val* maxDoc = yyjson_obj_get(root, "maxDocumentMiB");
                    if (yyjson_is_int(maxDoc))
                    {
                        const int64_t value = yyjson_get_int(maxDoc);
                        if (value >= 1)
                        {
                            maxDocumentMiB = static_cast<uint32_t>(std::min<int64_t>(value, 512));
                        }
                    }
                    else if (yyjson_is_uint(maxDoc))
                    {
                        maxDocumentMiB = static_cast<uint32_t>(std::min<uint64_t>(yyjson_get_uint(maxDoc), 512u));
                        maxDocumentMiB = std::max(maxDocumentMiB, 1u);
                    }

                    const auto readBool = [&](const char* key, bool defaultValue) -> bool
                    {
                        yyjson_val* v = yyjson_obj_get(root, key);
                        if (yyjson_is_bool(v))
                        {
                            return yyjson_get_bool(v);
                        }
                        if (yyjson_is_str(v))
                        {
                            const char* s = yyjson_get_str(v);
                            if (s && strcmp(s, "1") == 0)
                            {
                                return true;
                            }
                            if (s && strcmp(s, "0") == 0)
                            {
                                return false;
                            }
                        }
                        return defaultValue;
                    };

                    allowExternalNavigation = readBool("allowExternalNavigation", allowExternalNavigation);
                    devToolsEnabled         = readBool("devToolsEnabled", devToolsEnabled);

                    yyjson_val* modeVal = yyjson_obj_get(root, "viewMode");
                    if (yyjson_is_str(modeVal))
                    {
                        const char* s = yyjson_get_str(modeVal);
                        if (s && (strcmp(s, "tree") == 0 || strcmp(s, "1") == 0))
                        {
                            jsonViewMode = JsonViewMode::Tree;
                        }
                        else if (s && (strcmp(s, "pretty") == 0 || strcmp(s, "0") == 0))
                        {
                            jsonViewMode = JsonViewMode::Pretty;
                        }
                    }
                    else if (yyjson_is_int(modeVal))
                    {
                        jsonViewMode = (yyjson_get_int(modeVal) != 0) ? JsonViewMode::Tree : JsonViewMode::Pretty;
                    }
                    else if (yyjson_is_uint(modeVal))
                    {
                        jsonViewMode = (yyjson_get_uint(modeVal) != 0u) ? JsonViewMode::Tree : JsonViewMode::Pretty;
                    }
                }
            }
        }
    }

    _config.maxDocumentMiB          = maxDocumentMiB;
    _config.allowExternalNavigation = allowExternalNavigation;
    _config.devToolsEnabled         = devToolsEnabled;
    _config.jsonViewMode            = jsonViewMode;

    switch (_kind)
    {
        case ViewerWebKind::Json:
            _configurationJson = std::format(
                R"json({{
    "maxDocumentMiB": {},
    "viewMode": "{}",
    "devToolsEnabled": {}
}})json",
                _config.maxDocumentMiB,
                _config.jsonViewMode == JsonViewMode::Tree ? "tree" : "pretty",
                _config.devToolsEnabled ? "true" : "false");
            break;
        case ViewerWebKind::Markdown:
            _configurationJson = std::format(
                R"json({{
    "maxDocumentMiB": {},
    "allowExternalNavigation": {},
    "devToolsEnabled": {}
}})json",
                _config.maxDocumentMiB,
                _config.allowExternalNavigation ? "true" : "false",
                _config.devToolsEnabled ? "true" : "false");
            break;
        case ViewerWebKind::Web:
        default:
            _configurationJson = std::format(
                R"json({{
    "allowExternalNavigation": {},
    "devToolsEnabled": {}
}})json",
                _config.allowExternalNavigation ? "true" : "false",
                _config.devToolsEnabled ? "true" : "false");
            break;
    }

    if (_webView)
    {
        wil::com_ptr<ICoreWebView2Settings> settings;
        if (SUCCEEDED(_webView->get_Settings(settings.put())) && settings)
        {
            static_cast<void>(settings->put_AreDevToolsEnabled(_config.devToolsEnabled ? TRUE : FALSE));
        }
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    if (_configurationJson.empty())
    {
        *configurationJsonUtf8 = nullptr;
        return S_OK;
    }

    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    bool isDefault = false;
    switch (_kind)
    {
        case ViewerWebKind::Json: isDefault = _config.maxDocumentMiB == 32u && _config.jsonViewMode == JsonViewMode::Pretty && ! _config.devToolsEnabled; break;
        case ViewerWebKind::Markdown: isDefault = _config.maxDocumentMiB == 32u && _config.allowExternalNavigation && ! _config.devToolsEnabled; break;
        case ViewerWebKind::Web:
        default: isDefault = _config.allowExternalNavigation && ! _config.devToolsEnabled; break;
    }
    *pSomethingToSave = isDefault ? FALSE : TRUE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::Open(const ViewerOpenContext* context) noexcept
{
    if (! context || ! context->focusedPath || context->focusedPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (! context->fileSystem)
    {
        return E_INVALIDARG;
    }

    _fileSystem = context->fileSystem;
    _fileSystemName.assign(context->fileSystemName ? context->fileSystemName : L"");

    _otherFiles.clear();
    if (context->otherFiles && context->otherFileCount > 0)
    {
        _otherFiles.reserve(context->otherFileCount);
        for (unsigned long i = 0; i < context->otherFileCount; ++i)
        {
            const wchar_t* p = context->otherFiles[i];
            if (p && p[0] != L'\0')
            {
                _otherFiles.emplace_back(p);
            }
        }
    }
    if (_otherFiles.empty())
    {
        _otherFiles.emplace_back(context->focusedPath);
    }

    _otherIndex = 0;
    if (context->focusedOtherFileIndex < _otherFiles.size())
    {
        _otherIndex = static_cast<size_t>(context->focusedOtherFileIndex);
    }

    const std::wstring focusedPath(context->focusedPath);

    if (! _hWnd)
    {
        if (! RegisterWndClass(g_hInstance))
        {
            return E_FAIL;
        }

        HWND ownerWindow = context->ownerWindow;
        RECT ownerRect{};
        const bool hasOwnerRect = ownerWindow && GetWindowRect(ownerWindow, &ownerRect) != 0;

        wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(LoadMenuW(g_hInstance, MAKEINTRESOURCEW(IDR_VIEWERWEB_MENU)));

        const int x      = hasOwnerRect ? static_cast<int>(ownerRect.left) : CW_USEDEFAULT;
        const int y      = hasOwnerRect ? static_cast<int>(ownerRect.top) : CW_USEDEFAULT;
        const LONG wLong = ownerRect.right - ownerRect.left;
        const LONG hLong = ownerRect.bottom - ownerRect.top;
        const int w      = hasOwnerRect ? static_cast<int>(std::max<LONG>(1, wLong)) : 1000;
        const int h      = hasOwnerRect ? static_cast<int>(std::max<LONG>(1, hLong)) : 700;

        HWND window = CreateWindowExW(0, kClassName, L"", WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN, x, y, w, h, nullptr, menu.get(), g_hInstance, this);
        if (! window)
        {
            const DWORD lastError = Debug::ErrorWithLastError(L"ViewerWeb: CreateWindowExW failed.");
            return HRESULT_FROM_WIN32(lastError);
        }

        menu.release();
        _hWnd.reset(window);

        ApplyTheme(_hWnd.get());
        ApplyPendingViewerWebClassBackgroundBrush(_hWnd.get());

        AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
    }
    else
    {
        ApplyPendingViewerWebClassBackgroundBrush(_hWnd.get());
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
    }

    if (! _hWnd)
    {
        Debug::Error(L"ViewerWeb: Open failed because viewer window is missing after creation.");
        return E_FAIL;
    }

    RefreshFileCombo(_hWnd.get());
    return OpenPath(_hWnd.get(), focusedPath, false);
}

HRESULT STDMETHODCALLTYPE ViewerWeb::Close() noexcept
{
    _hWnd.reset();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::SetTheme(const ViewerTheme* theme) noexcept
{
    if (! theme || theme->version != 2)
    {
        return E_INVALIDARG;
    }

    _theme    = *theme;
    _hasTheme = true;

    RequestViewerWebClassBackgroundColor(ColorRefFromArgb(_theme.backgroundArgb));
    ApplyPendingViewerWebClassBackgroundBrush(_hWnd.get());

    if (_hWnd)
    {
        ApplyTheme(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::SetCallback(IViewerCallback* callback, void* cookie) noexcept
{
    _callback       = callback;
    _callbackCookie = cookie;
    return S_OK;
}

ATOM ViewerWeb::RegisterWndClass(HINSTANCE instance) noexcept
{
    if (g_viewerWebClassBackgroundBrush.classRegistered)
    {
        return 1;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProcThunk;
    wc.cbClsExtra    = 0;
    wc.cbWndExtra    = 0;
    wc.hInstance     = instance;
    wc.hIcon         = LoadIconW(nullptr, IDI_APPLICATION);
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = GetActiveViewerWebClassBackgroundBrush();
    wc.lpszMenuName  = nullptr;
    wc.lpszClassName = kClassName;
    wc.hIconSm       = wc.hIcon;

    const ATOM atom = RegisterClassExW(&wc);
    if (atom != 0)
    {
        g_viewerWebClassBackgroundBrush.classRegistered = true;
    }

    return atom;
}

LRESULT CALLBACK ViewerWeb::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ViewerWeb* self = reinterpret_cast<ViewerWeb*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self           = reinterpret_cast<ViewerWeb*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            InitPostedPayloadWindow(hwnd);
        }
    }

    if (self)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerWeb::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    static const UINT findMsg = RegisterWindowMessageW(FINDMSGSTRINGW);
    if (findMsg != 0 && msg == findMsg)
    {
        OnFindMessage(reinterpret_cast<const FINDREPLACEW*>(lp));
        return 0;
    }

    switch (msg)
    {
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_SIZE: OnSize(static_cast<UINT>(LOWORD(lp)), static_cast<UINT>(HIWORD(lp))); return 0;
        case WM_COMMAND: OnCommand(hwnd, static_cast<UINT>(LOWORD(wp)), static_cast<UINT>(HIWORD(wp)), reinterpret_cast<HWND>(lp)); return 0;
        case WM_KEYDOWN: OnKeyDown(hwnd, static_cast<UINT>(wp)); return 0;
        case WM_SYSKEYDOWN:
            if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
            {
                OnKeyDown(hwnd, static_cast<UINT>(wp));
                return 0;
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_PAINT: OnPaint(hwnd); return 0;
        case WM_ERASEBKGND: return OnEraseBkgnd(hwnd, reinterpret_cast<HDC>(wp));
        case WM_DPICHANGED: OnDpiChanged(hwnd, HIWORD(wp), reinterpret_cast<const RECT*>(lp)); return 0;
        case WM_MEASUREITEM: return OnMeasureItem(hwnd, reinterpret_cast<MEASUREITEMSTRUCT*>(lp));
        case WM_DRAWITEM: return OnDrawItem(hwnd, reinterpret_cast<DRAWITEMSTRUCT*>(lp));
        case kAsyncLoadCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncLoadResult>(lp);
            OnAsyncLoadComplete(std::move(result));
            return 0;
        }
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_NCACTIVATE:
        {
            const bool windowActive = wp != FALSE;
            ApplyTitleBarTheme(windowActive);
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case WM_NCDESTROY: return OnNcDestroy(hwnd, wp, lp);
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }
}

void ViewerWeb::OnAsyncLoadComplete(std::unique_ptr<AsyncLoadResult> result) noexcept
{
    if (! result || result->viewer != this)
    {
        return;
    }

    if (result->requestId != _openRequestId)
    {
        if (result->extractedWin32Path.has_value() && ! result->extractedWin32Path->empty())
        {
            std::error_code ec;
            static_cast<void>(std::filesystem::remove(*result->extractedWin32Path, ec));
        }
        return;
    }

    _statusMessage = result->statusMessage;

    if (FAILED(result->hr))
    {
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        }

        if (! _statusMessage.empty())
        {
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
        }
        return;
    }

    if (_hWnd && ! result->title.empty())
    {
        SetWindowTextW(_hWnd.get(), result->title.c_str());
    }

    _pendingPath.reset();
    _pendingWebContent.reset();

    if (_kind == ViewerWebKind::Web)
    {
        std::optional<std::filesystem::path> navPath;
        bool navIsTemp = false;

        if (result->extractedWin32Path.has_value() && ! result->extractedWin32Path->empty())
        {
            navPath   = result->extractedWin32Path.value();
            navIsTemp = true;
        }
        else if (IsProbablyWin32Path(result->path))
        {
            navPath   = std::filesystem::path(result->path);
            navIsTemp = false;
        }

        if (navPath.has_value())
        {
            if (navIsTemp)
            {
                if (_tempExtractedPath.has_value() && _tempExtractedPath != navPath)
                {
                    std::error_code ec;
                    static_cast<void>(std::filesystem::remove(*_tempExtractedPath, ec));
                }
                _tempExtractedPath = navPath;
            }
            else if (_tempExtractedPath.has_value())
            {
                std::error_code ec;
                static_cast<void>(std::filesystem::remove(*_tempExtractedPath, ec));
                _tempExtractedPath.reset();
            }

            const std::wstring url = UrlFromFilePath(navPath->wstring());
            if (url.empty())
            {
                _statusMessage = L"Failed to build file URL.";
                ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
                if (_hWnd)
                {
                    InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
                }
                return;
            }

            _pendingPath = url;
        }
    }
    else
    {
        const std::wstring html = Utf16FromUtf8(result->utf8);
        if (html.empty() && ! result->utf8.empty())
        {
            _statusMessage = L"Failed to build HTML document.";
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
            }
            return;
        }

        _pendingWebContent = html;
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
    }

    if (SUCCEEDED(EnsureWebView2(_hWnd.get())) && _webView)
    {
        if (_pendingWebContent.has_value())
        {
            const std::wstring html = std::move(_pendingWebContent.value());
            _pendingWebContent.reset();
            const HRESULT navHr = _webView->NavigateToString(html.c_str());
            if (FAILED(navHr))
            {
                ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, std::format(L"NavigateToString failed (hr=0x{:08X}).", static_cast<unsigned long>(navHr)));
            }
        }
        else if (_pendingPath.has_value())
        {
            const std::wstring url = std::move(_pendingPath.value());
            _pendingPath.reset();
            const HRESULT navHr = _webView->Navigate(url.c_str());
            if (FAILED(navHr))
            {
                ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, std::format(L"Navigate failed (hr=0x{:08X}).", static_cast<unsigned long>(navHr)));
            }
        }
    }
}

void ViewerWeb::Layout(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    ComputeLayoutRects(hwnd);

    const UINT dpi        = GetDpiForWindow(hwnd);
    const int edgeSizeY   = GetSystemMetricsForDpi(SM_CYEDGE, dpi);
    const int minPadding  = PxFromDip(3, dpi);
    const int accentH     = std::max(1, PxFromDip(1, dpi));
    const int accentGap   = std::max(1, PxFromDip(1, dpi));
    const int comboBorder = std::max(0, edgeSizeY) * 2;

    const int padding    = PxFromDip(8, dpi);
    const bool showCombo = _hFileCombo && _otherFiles.size() > 1;

    const int headerH = std::max(0L, _headerRect.bottom - _headerRect.top);

    int desiredComboHeight = 0;
    if (showCombo && _hFileCombo)
    {
        int comboItemHeight           = 0;
        const LRESULT selectionHeight = SendMessageW(_hFileCombo.get(), CB_GETITEMHEIGHT, static_cast<WPARAM>(-1), 0);
        if (selectionHeight != CB_ERR && selectionHeight > 0)
        {
            comboItemHeight = static_cast<int>(selectionHeight);
        }

        if (comboItemHeight <= 0)
        {
            comboItemHeight = PxFromDip(24, dpi);
            auto hdc        = wil::GetDC(hwnd);
            if (hdc)
            {
                HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
                auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
                static_cast<void>(oldFont);

                TEXTMETRICW tm{};
                if (GetTextMetricsW(hdc.get(), &tm) != 0)
                {
                    comboItemHeight = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, dpi);
                }
            }
        }

        const int comboChromePadding = std::max(PxFromDip(4, dpi), comboBorder);
        desiredComboHeight           = std::max(1, comboItemHeight + comboChromePadding);
    }

    RECT headerContentRect{};
    headerContentRect        = _headerRect;
    headerContentRect.top    = std::min(headerContentRect.bottom, headerContentRect.top + minPadding);
    headerContentRect.bottom = std::max(headerContentRect.top, headerContentRect.bottom - accentH - accentGap - minPadding);
    const int headerContentH = std::max(0L, headerContentRect.bottom - headerContentRect.top);

    if (_hFileCombo)
    {
        ShowWindow(_hFileCombo.get(), showCombo ? SW_SHOW : SW_HIDE);
        EnableWindow(_hFileCombo.get(), showCombo ? TRUE : FALSE);

        if (showCombo)
        {
            const int statusReserveW = _statusMessage.empty() ? 0 : PxFromDip(160, dpi);
            const int margin         = PxFromDip(10, dpi);

            const int comboX = headerContentRect.left + margin;
            int rightLimit   = std::max<LONG>(headerContentRect.left, headerContentRect.right) - margin;
            if (statusReserveW)
            {
                rightLimit = std::max(comboX, rightLimit - statusReserveW - margin);
            }
            const int comboW = std::max(0, rightLimit - comboX);

            int comboH = desiredComboHeight > 0 ? desiredComboHeight : std::max(1, headerH - 2 * padding);
            comboH     = std::clamp(comboH, 1, std::max(1, headerContentH));

            SetWindowPos(_hFileCombo.get(), nullptr, comboX, headerContentRect.top, comboW, comboH, SWP_NOZORDER | SWP_NOACTIVATE);

            RECT comboRc{};
            int actualComboH = comboH;
            if (GetWindowRect(_hFileCombo.get(), &comboRc) != 0)
            {
                actualComboH = std::max(0L, comboRc.bottom - comboRc.top);
            }

            int comboY = headerContentRect.top + std::max(0, (headerContentH - actualComboH) / 2);

            const int maxBottom = std::max(static_cast<int>(headerContentRect.top), static_cast<int>(headerContentRect.bottom));
            if (comboY + actualComboH > maxBottom)
            {
                comboY = std::max(static_cast<int>(headerContentRect.top), maxBottom - actualComboH);
            }

            SetWindowPos(_hFileCombo.get(), nullptr, comboX, comboY, 0, 0, SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSIZE);
        }
    }

    if (_webViewController)
    {
        RECT bounds   = _contentRect;
        bounds.right  = std::max(bounds.right, bounds.left);
        bounds.bottom = std::max(bounds.bottom, bounds.top);
        static_cast<void>(_webViewController->put_Bounds(bounds));
    }
}

void ViewerWeb::ComputeLayoutRects(HWND hwnd) noexcept
{
    RECT client{};
    if (! hwnd || GetClientRect(hwnd, &client) == 0)
    {
        _headerRect  = {};
        _contentRect = {};
        return;
    }

    const UINT dpi             = GetDpiForWindow(hwnd);
    const int edgeSizeY        = GetSystemMetricsForDpi(SM_CYEDGE, dpi);
    const int baseHeaderHeight = PxFromDip(kHeaderHeightDip, dpi);
    const int accentH          = std::max(1, PxFromDip(1, dpi));
    const int accentGap        = std::max(1, PxFromDip(1, dpi));
    const int minPadding       = PxFromDip(3, dpi);
    const int comboBorder      = std::max(0, edgeSizeY) * 2;

    const bool showCombo   = _hFileCombo && _otherFiles.size() > 1;
    int desiredComboHeight = 0;
    if (showCombo && _hFileCombo)
    {
        int comboItemHeight           = 0;
        const LRESULT selectionHeight = SendMessageW(_hFileCombo.get(), CB_GETITEMHEIGHT, static_cast<WPARAM>(-1), 0);
        if (selectionHeight != CB_ERR && selectionHeight > 0)
        {
            comboItemHeight = static_cast<int>(selectionHeight);
        }

        if (comboItemHeight <= 0)
        {
            comboItemHeight = PxFromDip(24, dpi);
            auto hdc        = wil::GetDC(hwnd);
            if (hdc)
            {
                HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
                auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
                static_cast<void>(oldFont);

                TEXTMETRICW tm{};
                if (GetTextMetricsW(hdc.get(), &tm) != 0)
                {
                    comboItemHeight = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, dpi);
                }
            }
        }

        const int comboChromePadding = std::max(PxFromDip(4, dpi), comboBorder);
        desiredComboHeight           = std::max(1, comboItemHeight + comboChromePadding);
    }

    const int minChromeHeight = PxFromDip(22, dpi) + accentH + accentGap + 2 * minPadding;
    int headerH               = std::max(baseHeaderHeight, minChromeHeight);
    if (showCombo && desiredComboHeight > 0)
    {
        headerH = std::max(headerH, desiredComboHeight + accentH + accentGap + 2 * minPadding);
    }

    _headerRect        = client;
    _headerRect.bottom = std::min(client.bottom, client.top + headerH);

    _contentRect     = client;
    _contentRect.top = _headerRect.bottom;

    _headerRect.left   = std::max<LONG>(0, _headerRect.left);
    _headerRect.top    = std::max<LONG>(0, _headerRect.top);
    _headerRect.right  = std::max(_headerRect.right, _headerRect.left);
    _headerRect.bottom = std::max(_headerRect.bottom, _headerRect.top);

    _contentRect.left   = std::max<LONG>(0, _contentRect.left);
    _contentRect.top    = std::max<LONG>(0, _contentRect.top);
    _contentRect.right  = std::max(_contentRect.right, _contentRect.left);
    _contentRect.bottom = std::max(_contentRect.bottom, _contentRect.top);
}

void ViewerWeb::ApplyTheme(HWND hwnd) noexcept
{
    const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);

    COLORREF headerBg = bg;
    if (_hasTheme && _theme.darkMode)
    {
        headerBg = RGB(std::max(0, GetRValue(bg) - 10), std::max(0, GetGValue(bg) - 10), std::max(0, GetBValue(bg) - 10));
    }
    else
    {
        headerBg = RGB(std::max(0, GetRValue(bg) - 5), std::max(0, GetGValue(bg) - 5), std::max(0, GetBValue(bg) - 5));
    }

    _headerBrush.reset(CreateSolidBrush(headerBg));

    if (_hasTheme && _hWnd)
    {
        const bool windowActive = GetActiveWindow() == _hWnd.get();
        ApplyTitleBarTheme(windowActive);
    }

    const wchar_t* winTheme = L"Explorer";
    if (_hasTheme && _theme.highContrast)
    {
        winTheme = L"";
    }
    else if (_hasTheme && _theme.darkMode)
    {
        winTheme = L"DarkMode_Explorer";
    }

    if (_hFileCombo)
    {
        SetWindowTheme(_hFileCombo.get(), winTheme, nullptr);
        SendMessageW(_hFileCombo.get(), WM_THEMECHANGED, 0, 0);
        if (_hFileComboList)
        {
            SetWindowTheme(_hFileComboList, winTheme, nullptr);
            SendMessageW(_hFileComboList, WM_THEMECHANGED, 0, 0);
        }
        if (_hFileComboItem)
        {
            SetWindowTheme(_hFileComboItem, winTheme, nullptr);
            SendMessageW(_hFileComboItem, WM_THEMECHANGED, 0, 0);
        }
    }

    ApplyMenuTheme(hwnd);

    HMENU menu = hwnd ? GetMenu(hwnd) : nullptr;
    if (menu)
    {
        const bool jsonTreeMode = _kind == ViewerWebKind::Json && _config.jsonViewMode == JsonViewMode::Tree;

        EnableMenuItem(menu, IDM_VIEWERWEB_VIEW_DEVTOOLS, static_cast<UINT>(MF_BYCOMMAND | (_config.devToolsEnabled ? MF_ENABLED : MF_GRAYED)));

        EnableMenuItem(menu, IDM_VIEWERWEB_TOOLS_JSON_EXPAND_ALL, static_cast<UINT>(MF_BYCOMMAND | (jsonTreeMode ? MF_ENABLED : MF_GRAYED)));
        EnableMenuItem(menu, IDM_VIEWERWEB_TOOLS_JSON_COLLAPSE_ALL, static_cast<UINT>(MF_BYCOMMAND | (jsonTreeMode ? MF_ENABLED : MF_GRAYED)));
        EnableMenuItem(
            menu, IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE, static_cast<UINT>(MF_BYCOMMAND | (_kind == ViewerWebKind::Markdown ? MF_ENABLED : MF_GRAYED)));

        CheckMenuItem(menu,
                      IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE,
                      static_cast<UINT>(MF_BYCOMMAND | ((_kind == ViewerWebKind::Markdown && _markdownShowSource) ? MF_CHECKED : MF_UNCHECKED)));

        DrawMenuBar(hwnd);
    }

    UpdateWebViewTheme();
}

void ViewerWeb::ApplyTitleBarTheme(bool windowActive) noexcept
{
    if (! _hasTheme || ! _hWnd)
    {
        return;
    }

    static constexpr DWORD kDwmwaUseImmersiveDarkMode19 = 19u;
    static constexpr DWORD kDwmwaUseImmersiveDarkMode20 = 20u;
    static constexpr DWORD kDwmwaBorderColor            = 34u;
    static constexpr DWORD kDwmwaCaptionColor           = 35u;
    static constexpr DWORD kDwmwaTextColor              = 36u;
    static constexpr DWORD kDwmColorDefault             = 0xFFFFFFFFu;

    const BOOL darkMode = (_theme.darkMode && ! _theme.highContrast) ? TRUE : FALSE;
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaUseImmersiveDarkMode20, &darkMode, sizeof(darkMode));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaUseImmersiveDarkMode19, &darkMode, sizeof(darkMode));

    DWORD borderValue  = kDwmColorDefault;
    DWORD captionValue = kDwmColorDefault;
    DWORD textValue    = kDwmColorDefault;
    if (! _theme.highContrast && _theme.rainbowMode)
    {
        COLORREF accent = ResolveAccentColor(_theme, L"title");
        if (! windowActive)
        {
            static constexpr uint8_t kInactiveTitleBlendAlpha = 223u; // ~7/8 toward background
            const COLORREF bg                                 = ColorRefFromArgb(_theme.backgroundArgb);
            accent                                            = BlendColor(accent, bg, kInactiveTitleBlendAlpha);
        }

        const COLORREF text = ContrastingTextColor(accent);
        borderValue         = static_cast<DWORD>(accent);
        captionValue        = static_cast<DWORD>(accent);
        textValue           = static_cast<DWORD>(text);
    }

    DwmSetWindowAttribute(_hWnd.get(), kDwmwaBorderColor, &borderValue, sizeof(borderValue));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaCaptionColor, &captionValue, sizeof(captionValue));
    DwmSetWindowAttribute(_hWnd.get(), kDwmwaTextColor, &textValue, sizeof(textValue));
}

void ViewerWeb::ApplyMenuTheme(HWND hwnd) noexcept
{
    HMENU menu = hwnd ? GetMenu(hwnd) : nullptr;
    if (! menu)
    {
        return;
    }

    if (_headerBrush)
    {
        MENUINFO mi{};
        mi.cbSize  = sizeof(mi);
        mi.fMask   = MIM_BACKGROUND | MIM_APPLYTOSUBMENUS;
        mi.hbrBack = _headerBrush.get();
        SetMenuInfo(menu, &mi);
    }

    _menuThemeItems.clear();
    PrepareMenuTheme(menu, true, _menuThemeItems);
    DrawMenuBar(hwnd);
}

void ViewerWeb::PrepareMenuTheme(HMENU menu, bool topLevel, std::vector<MenuItemData>& outItems) noexcept
{
    const int count = GetMenuItemCount(menu);
    if (count <= 0)
    {
        return;
    }

    for (UINT pos = 0; pos < static_cast<UINT>(count); ++pos)
    {
        MENUITEMINFOW info{};
        info.cbSize = sizeof(info);
        wchar_t textBuf[256]{};
        info.fMask      = MIIM_FTYPE | MIIM_STRING | MIIM_SUBMENU | MIIM_ID;
        info.dwTypeData = textBuf;
        info.cch        = static_cast<UINT>(std::size(textBuf) - 1);
        if (GetMenuItemInfoW(menu, pos, TRUE, &info) == 0)
        {
            continue;
        }

        MenuItemData data{};
        data.id         = info.wID;
        data.separator  = (info.fType & MFT_SEPARATOR) != 0;
        data.topLevel   = topLevel;
        data.hasSubMenu = info.hSubMenu != nullptr;

        if (! data.separator)
        {
            std::wstring text = textBuf;
            const size_t tab  = text.find(L'\t');
            if (tab != std::wstring::npos)
            {
                data.shortcut = text.substr(tab + 1);
                text.resize(tab);
            }
            data.text = std::move(text);
        }

        const size_t index = outItems.size();
        outItems.push_back(std::move(data));

        MENUITEMINFOW ownerDraw{};
        ownerDraw.cbSize     = sizeof(ownerDraw);
        ownerDraw.fMask      = MIIM_FTYPE | MIIM_DATA;
        ownerDraw.fType      = info.fType | MFT_OWNERDRAW;
        ownerDraw.dwItemData = static_cast<ULONG_PTR>(index);
        SetMenuItemInfoW(menu, pos, TRUE, &ownerDraw);

        if (info.hSubMenu)
        {
            PrepareMenuTheme(info.hSubMenu, false, outItems);
        }
    }
}

void ViewerWeb::ShowHostAlert(HWND targetWindow, HostAlertSeverity severity, const std::wstring& message) noexcept
{
    if (message.empty())
    {
        return;
    }

    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_TITLE);

    if (! _hostAlerts)
    {
        MessageBoxW(targetWindow, message.c_str(), title.empty() ? L"ViewerWeb" : title.c_str(), MB_OK | MB_ICONERROR);
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = (targetWindow && IsWindow(targetWindow)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = severity;
    request.targetWindow = request.scope == HOST_ALERT_SCOPE_WINDOW ? targetWindow : nullptr;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(_hostAlerts->ShowAlert(&request, targetWindow));
}

HRESULT ViewerWeb::EnsureWebView2(HWND hwnd) noexcept
{
    if (_webView)
    {
        return S_OK;
    }

    if (! hwnd)
    {
        return E_INVALIDARG;
    }

    if (_webViewInitInProgress)
    {
        return S_FALSE;
    }

    _webViewInitInProgress = true;

    AddRef();
    const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
        nullptr,
        nullptr,
        nullptr,
        MakeComCallback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler, HRESULT, ICoreWebView2Environment*>(
            [this, hwnd](HRESULT result, ICoreWebView2Environment* environment) -> HRESULT
            {
                if (FAILED(result) || ! environment)
                {
                    _webViewInitInProgress = false;

                    const UINT msgId = (result == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || result == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
                                           ? IDS_VIEWERWEB_ERROR_WEBVIEW2_RUNTIME_MISSING
                                           : IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED;
                    ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, msgId));
                    Release();
                    return S_OK;
                }

                _webViewEnvironment = environment;

                const HRESULT createControllerHr = environment->CreateCoreWebView2Controller(
                    hwnd,
                    MakeComCallback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler, HRESULT, ICoreWebView2Controller*>(
                        [this, hwnd](HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT
                        {
                            _webViewInitInProgress = false;
                            auto releaseSelf       = wil::scope_exit([&] { Release(); });

                            if (FAILED(controllerResult) || ! controller)
                            {
                                ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED));
                                return S_OK;
                            }

                            if (! _hWnd || _hWnd.get() != hwnd)
                            {
                                controller->Close();
                                return S_OK;
                            }

                            _webViewController = controller;

                            wil::com_ptr<ICoreWebView2> webView;
                            const HRESULT webViewHr = controller->get_CoreWebView2(webView.put());
                            if (FAILED(webViewHr) || ! webView)
                            {
                                ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED));
                                _webViewController = nullptr;
                                return S_OK;
                            }

                            _webView = std::move(webView);

                            static_cast<void>(_webView->add_NavigationStarting(
                                MakeComCallback<ICoreWebView2NavigationStartingEventHandler, ICoreWebView2*, ICoreWebView2NavigationStartingEventArgs*>(
                                    [this](ICoreWebView2* /*sender*/, ICoreWebView2NavigationStartingEventArgs* args) -> HRESULT
                                    {
                                        if (! args)
                                        {
                                            return S_OK;
                                        }

                                        wil::unique_cotaskmem_string uri;
                                        static_cast<void>(args->get_Uri(uri.put()));
                                        if (! uri)
                                        {
                                            return S_OK;
                                        }

                                        const std::wstring_view url(uri.get());
                                        const bool isHttp  = StartsWithNoCase(url, L"http://") || StartsWithNoCase(url, L"https://");
                                        const bool isAbout = StartsWithNoCase(url, L"about:");
                                        const bool isData  = StartsWithNoCase(url, L"data:");

                                        if (_kind == ViewerWebKind::Web)
                                        {
                                            if (isHttp && ! _config.allowExternalNavigation)
                                            {
                                                static_cast<void>(args->put_Cancel(TRUE));
                                            }
                                            return S_OK;
                                        }

                                        // JSON/Markdown: keep viewer content stable and open external links in the system browser.
                                        if (isHttp)
                                        {
                                            static_cast<void>(args->put_Cancel(TRUE));
                                            ShellExecuteW(nullptr, L"open", uri.get(), nullptr, nullptr, SW_SHOWNORMAL);
                                            return S_OK;
                                        }

                                        if (! isAbout && ! isData)
                                        {
                                            static_cast<void>(args->put_Cancel(TRUE));
                                        }

                                        return S_OK;
                                    })
                                    .get(),
                                &_navStartingToken));

                            static_cast<void>(_webView->add_NavigationCompleted(
                                MakeComCallback<ICoreWebView2NavigationCompletedEventHandler, ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs*>(
                                    [this](ICoreWebView2* /*sender*/, ICoreWebView2NavigationCompletedEventArgs* /*args*/) -> HRESULT
                                    {
                                        UpdateWebViewTheme();
                                        return S_OK;
                                    })
                                    .get(),
                                &_navCompletedToken));

                            static_cast<void>(_webViewController->add_AcceleratorKeyPressed(
                                MakeComCallback<ICoreWebView2AcceleratorKeyPressedEventHandler,
                                                ICoreWebView2Controller*,
                                                ICoreWebView2AcceleratorKeyPressedEventArgs*>(
                                    [this](ICoreWebView2Controller* /*sender*/, ICoreWebView2AcceleratorKeyPressedEventArgs* args) -> HRESULT
                                    {
                                        if (! args || ! _hWnd)
                                        {
                                            return S_OK;
                                        }

                                        COREWEBVIEW2_KEY_EVENT_KIND kind{};
                                        if (FAILED(args->get_KeyEventKind(&kind)))
                                        {
                                            return S_OK;
                                        }

                                        if (kind != COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN && kind != COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN)
                                        {
                                            return S_OK;
                                        }

                                        UINT vk = 0;
                                        static_cast<void>(args->get_VirtualKey(&vk));

                                        const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                                        const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

                                        const auto handle = [&](bool handled) noexcept
                                        {
                                            if (handled)
                                            {
                                                static_cast<void>(args->put_Handled(TRUE));
                                            }
                                        };

                                        if (vk == VK_ESCAPE)
                                        {
                                            _hWnd.reset();
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (vk == VK_F5)
                                        {
                                            static_cast<void>(OpenPath(_hWnd.get(), _currentPath, false));
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (vk == VK_F12)
                                        {
                                            CommandToggleDevTools();
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (vk == VK_F3)
                                        {
                                            shift ? CommandFindPrevious(_hWnd.get()) : CommandFindNext(_hWnd.get());
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && (vk == 'F' || vk == 'f'))
                                        {
                                            CommandFind(_hWnd.get());
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && (vk == 'S' || vk == 's'))
                                        {
                                            static_cast<void>(CommandSaveAs(_hWnd.get()));
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && (vk == 'L' || vk == 'l'))
                                        {
                                            CommandCopyUrl(_hWnd.get());
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && vk == VK_RETURN)
                                        {
                                            CommandOpenExternal(_hWnd.get());
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && (vk == VK_OEM_PLUS || vk == VK_ADD || vk == '='))
                                        {
                                            CommandZoomIn();
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && (vk == VK_OEM_MINUS || vk == VK_SUBTRACT || vk == '-'))
                                        {
                                            CommandZoomOut();
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && vk == '0')
                                        {
                                            CommandZoomReset();
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && vk == VK_OEM_3)
                                        {
                                            CommandMarkdownToggleSource();
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && vk == VK_UP)
                                        {
                                            SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_PREVIOUS, 0);
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && vk == VK_DOWN)
                                        {
                                            SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_NEXT, 0);
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && vk == VK_HOME)
                                        {
                                            SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_FIRST, 0);
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (ctrl && vk == VK_END)
                                        {
                                            SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_LAST, 0);
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (_kind != ViewerWebKind::Web && vk == VK_SPACE)
                                        {
                                            SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_NEXT, 0);
                                            handle(true);
                                            return S_OK;
                                        }

                                        if (_kind != ViewerWebKind::Web && vk == VK_BACK)
                                        {
                                            SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_PREVIOUS, 0);
                                            handle(true);
                                            return S_OK;
                                        }

                                        return S_OK;
                                    })
                                    .get(),
                                &_accelToken));

                            Layout(hwnd);
                            UpdateWebViewTheme();

                            if (_pendingWebContent.has_value())
                            {
                                const std::wstring html = std::move(_pendingWebContent.value());
                                _pendingWebContent.reset();
                                const HRESULT navHr = _webView->NavigateToString(html.c_str());
                                if (FAILED(navHr))
                                {
                                    ShowHostAlert(
                                        hwnd, HOST_ALERT_ERROR, std::format(L"NavigateToString failed (hr=0x{:08X}).", static_cast<unsigned long>(navHr)));
                                }
                            }
                            else if (_pendingPath.has_value())
                            {
                                const std::wstring url = std::move(_pendingPath.value());
                                _pendingPath.reset();
                                const HRESULT navHr = _webView->Navigate(url.c_str());
                                if (FAILED(navHr))
                                {
                                    ShowHostAlert(hwnd, HOST_ALERT_ERROR, std::format(L"Navigate failed (hr=0x{:08X}).", static_cast<unsigned long>(navHr)));
                                }
                            }

                            return S_OK;
                        })
                        .get());

                if (FAILED(createControllerHr))
                {
                    _webViewInitInProgress = false;
                    ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED));
                    Release();
                }

                return S_OK;
            })
            .get());

    if (FAILED(hr))
    {
        _webViewInitInProgress = false;
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED));
        Release();
        return hr;
    }

    return S_OK;
}

void ViewerWeb::DiscardWebView2() noexcept
{
    _webViewInitInProgress = false;

    // Unregister event handlers before closing
    if (_webViewController)
    {
        static_cast<void>(_webViewController->remove_AcceleratorKeyPressed(_accelToken));
    }
    if (_webView)
    {
        static_cast<void>(_webView->remove_NavigationStarting(_navStartingToken));
        static_cast<void>(_webView->remove_NavigationCompleted(_navCompletedToken));
    }

    _navStartingToken  = {};
    _navCompletedToken = {};
    _accelToken        = {};

    // Close the WebView2 controller. Note: Close() is asynchronous and may have
    // pending I/O operations that complete on thread pool threads. This is why
    // we must call DiscardWebView2() early in OnDestroy() rather than in the
    // destructor - to give WebView2 time to complete its shutdown before the
    // plugin DLL is unloaded.
    if (_webViewController)
    {
        _webViewController->Close();
    }

    // Release COM pointers after initiating close
    _webView.reset();
    _webViewController.reset();
    _webViewEnvironment.reset();
}

void ViewerWeb::UpdateWebViewTheme() noexcept
{
    if (_webViewController)
    {
        wil::com_ptr<ICoreWebView2Controller2> controller2;
        if (SUCCEEDED(_webViewController->QueryInterface(IID_PPV_ARGS(controller2.put()))) && controller2)
        {
            const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
            COREWEBVIEW2_COLOR color{};
            color.A = 255;
            color.R = GetRValue(bg);
            color.G = GetGValue(bg);
            color.B = GetBValue(bg);
            static_cast<void>(controller2->put_DefaultBackgroundColor(color));
        }
    }

    if (_webView)
    {
        wil::com_ptr<ICoreWebView2Settings> settings;
        if (SUCCEEDED(_webView->get_Settings(settings.put())) && settings)
        {
            static_cast<void>(settings->put_IsScriptEnabled(TRUE));
            static_cast<void>(settings->put_IsWebMessageEnabled(TRUE));
            static_cast<void>(settings->put_AreDefaultContextMenusEnabled(TRUE));
            static_cast<void>(settings->put_IsZoomControlEnabled(TRUE));
            static_cast<void>(settings->put_AreDevToolsEnabled(_config.devToolsEnabled ? TRUE : FALSE));

            wil::com_ptr<ICoreWebView2Settings3> settings3;
            if (SUCCEEDED(settings->QueryInterface(IID_PPV_ARGS(settings3.put()))) && settings3)
            {
                static_cast<void>(settings3->put_AreBrowserAcceleratorKeysEnabled(TRUE));
            }
        }

        if (_hasTheme)
        {
            const COLORREF bg     = ColorRefFromArgb(_theme.backgroundArgb);
            const COLORREF fg     = ColorRefFromArgb(_theme.textArgb);
            const COLORREF selBg  = ColorRefFromArgb(_theme.selectionBackgroundArgb);
            const COLORREF selFg  = ColorRefFromArgb(_theme.selectionTextArgb);
            const COLORREF accent = ResolveAccentColor(_theme, _currentPath.empty() ? _metaId : _currentPath);

            const auto cssRgb = [](COLORREF c) -> std::wstring { return std::format(L"rgb({},{},{})", GetRValue(c), GetGValue(c), GetBValue(c)); };

            const std::wstring script = std::format(L"(function(){{try{{if(window.RS&&window.RS.applyTheme){{window.RS.applyTheme({{bg:'{}',fg:'{}',selBg:'{}',"
                                                    L"selFg:'{}',accent:'{}'}});}}}}catch(e){{}}}})();",
                                                    cssRgb(bg),
                                                    cssRgb(fg),
                                                    cssRgb(selBg),
                                                    cssRgb(selFg),
                                                    cssRgb(accent));
            static_cast<void>(_webView->ExecuteScript(script.c_str(), nullptr));
        }
    }
}

HRESULT ViewerWeb::OpenPath(HWND hwnd, const std::wstring& path, bool updateOtherFiles) noexcept
{
    if (! hwnd)
    {
        return E_INVALIDARG;
    }

    if (path.empty())
    {
        Debug::Error(L"ViewerWeb: OpenPath called with an empty path.");
        return E_INVALIDARG;
    }

    if (! _fileSystem)
    {
        Debug::Error(L"ViewerWeb: OpenPath failed because file system is missing.");
        return E_FAIL;
    }

    _currentPath = path;

    if (updateOtherFiles)
    {
        _otherFiles.clear();
        _otherFiles.push_back(path);
        _otherIndex = 0;
        RefreshFileCombo(hwnd);
    }
    else if (! _otherFiles.empty())
    {
        for (size_t i = 0; i < _otherFiles.size(); ++i)
        {
            const std::wstring_view a(_otherFiles[i]);
            const std::wstring_view b(path);
            if (a.size() == b.size() && CompareStringOrdinal(a.data(), static_cast<int>(a.size()), b.data(), static_cast<int>(b.size()), TRUE) == CSTR_EQUAL)
            {
                _otherIndex = i;
                break;
            }
        }

        if (_hFileCombo)
        {
            SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(_otherIndex), 0);
        }
    }

    const std::wstring leaf = LeafNameFromPath(path);
    std::wstring title      = leaf;
    if (! _metaName.empty())
    {
        title = leaf.empty() ? _metaName : std::format(L"{} - {}", leaf, _metaName);
    }
    if (_hWnd && ! title.empty())
    {
        SetWindowTextW(_hWnd.get(), title.c_str());
    }

    _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_STATUS_LOADING);
    _pendingPath.reset();
    _pendingWebContent.reset();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
    }

    return StartAsyncLoad(hwnd, path);
}

void ViewerWeb::RefreshFileCombo(HWND hwnd) noexcept
{
    if (! _hFileCombo)
    {
        return;
    }

    SendMessageW(_hFileCombo.get(), CB_RESETCONTENT, 0, 0);

    for (const auto& path : _otherFiles)
    {
        std::wstring leaf = LeafNameFromPath(path);
        if (leaf.empty())
        {
            leaf = path;
        }

        SendMessageW(_hFileCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(leaf.c_str()));
    }

    if (_otherIndex < _otherFiles.size())
    {
        SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(_otherIndex), 0);
    }

    Layout(hwnd);
}

HRESULT ViewerWeb::StartAsyncLoad(HWND hwnd, const std::wstring& path) noexcept
{
    if (! hwnd)
    {
        return E_INVALIDARG;
    }

    if (path.empty())
    {
        return E_INVALIDARG;
    }

    if (! _fileSystem)
    {
        return E_FAIL;
    }

    _openRequestId += 1u;
    const uint64_t requestId = _openRequestId;

    std::unique_ptr<AsyncLoadResult> payload(new (std::nothrow) AsyncLoadResult{});
    if (! payload)
    {
        return E_OUTOFMEMORY;
    }

    payload->viewer    = this;
    payload->hwnd      = hwnd;
    payload->requestId = requestId;
    payload->path      = path;
    payload->hr        = E_FAIL;

    AddRef();

    struct AsyncLoadWorkItem final
    {
        AsyncLoadWorkItem()                                    = default;
        AsyncLoadWorkItem(const AsyncLoadWorkItem&)            = delete;
        AsyncLoadWorkItem& operator=(const AsyncLoadWorkItem&) = delete;

        std::unique_ptr<AsyncLoadResult> payload;
        wil::unique_hmodule moduleKeepAlive;
    };

    auto ctx = std::unique_ptr<AsyncLoadWorkItem>(new (std::nothrow) AsyncLoadWorkItem{});
    if (! ctx)
    {
        Release();
        return E_OUTOFMEMORY;
    }

    ctx->payload         = std::move(payload);
    ctx->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerWebModuleAnchor);

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
        {
            std::unique_ptr<AsyncLoadWorkItem> ctx(static_cast<AsyncLoadWorkItem*>(context));
            if (! ctx || ! ctx->payload)
            {
                return;
            }

            static_cast<void>(ctx->moduleKeepAlive);
            AsyncLoadProc(ctx->payload.release());
        },
        ctx.get(),
        nullptr);

    if (queued == 0)
    {
        Release();
        return E_FAIL;
    }

    ctx.release();
    return S_OK;
}

void ViewerWeb::AsyncLoadProc(AsyncLoadResult* payload) noexcept
{
    std::unique_ptr<AsyncLoadResult> result(payload);
    if (! result || ! result->viewer)
    {
        return;
    }

    ViewerWeb* self  = result->viewer;
    auto releaseSelf = wil::scope_exit([&] { self->Release(); });

    const ViewerWebKind kind      = self->_kind;
    const ViewerWebConfig config  = self->_config;
    const bool hasTheme           = self->_hasTheme;
    const ViewerTheme theme       = self->_theme;
    const bool markdownShowSource = self->_markdownShowSource;

    wil::com_ptr<IFileSystem> fileSystem = self->_fileSystem;

    const auto base64Encode = [](std::string_view bytes) noexcept -> std::string
    {
        static constexpr char kTable[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
        std::string out;
        out.reserve(((bytes.size() + 2) / 3) * 4);

        size_t i = 0;
        for (; i + 3 <= bytes.size(); i += 3)
        {
            const uint32_t n = (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i])) << 16) |
                               (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i + 1])) << 8) | static_cast<uint32_t>(static_cast<uint8_t>(bytes[i + 2]));
            out.push_back(kTable[(n >> 18) & 0x3F]);
            out.push_back(kTable[(n >> 12) & 0x3F]);
            out.push_back(kTable[(n >> 6) & 0x3F]);
            out.push_back(kTable[n & 0x3F]);
        }

        const size_t rem = bytes.size() - i;
        if (rem == 1)
        {
            const uint32_t n = (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i])) << 16);
            out.push_back(kTable[(n >> 18) & 0x3F]);
            out.push_back(kTable[(n >> 12) & 0x3F]);
            out.push_back('=');
            out.push_back('=');
        }
        else if (rem == 2)
        {
            const uint32_t n = (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i])) << 16) | (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i + 1])) << 8);
            out.push_back(kTable[(n >> 18) & 0x3F]);
            out.push_back(kTable[(n >> 12) & 0x3F]);
            out.push_back(kTable[(n >> 6) & 0x3F]);
            out.push_back('=');
        }

        return out;
    };

    const auto replaceAll = [](std::string& text, std::string_view needle, std::string_view replacement) noexcept
    {
        if (needle.empty())
        {
            return;
        }

        size_t pos = 0;
        while ((pos = text.find(needle, pos)) != std::string::npos)
        {
            text.replace(pos, needle.size(), replacement);
            pos += replacement.size();
        }
    };

    const auto normalizeTextUtf8 = [&](std::string_view bytes) noexcept -> std::string
    {
        const auto asU8 = [&](size_t offset) noexcept -> std::string { return std::string(bytes.data() + offset, bytes.data() + bytes.size()); };

        if (bytes.size() >= 3 && static_cast<uint8_t>(bytes[0]) == 0xEF && static_cast<uint8_t>(bytes[1]) == 0xBB && static_cast<uint8_t>(bytes[2]) == 0xBF)
        {
            return asU8(3);
        }

        if (bytes.size() >= 2 && static_cast<uint8_t>(bytes[0]) == 0xFF && static_cast<uint8_t>(bytes[1]) == 0xFE)
        {
            const size_t payloadBytes = bytes.size() - 2;
            const size_t wcharCount   = payloadBytes / 2;
            std::wstring w(static_cast<size_t>(wcharCount), L'\0');
            memcpy(w.data(), bytes.data() + 2, wcharCount * sizeof(wchar_t));
            return Utf8FromUtf16(w);
        }

        if (bytes.size() >= 2 && static_cast<uint8_t>(bytes[0]) == 0xFE && static_cast<uint8_t>(bytes[1]) == 0xFF)
        {
            const size_t payloadBytes = bytes.size() - 2;
            const size_t wcharCount   = payloadBytes / 2;
            std::wstring w(static_cast<size_t>(wcharCount), L'\0');
            for (size_t i = 0; i < wcharCount; ++i)
            {
                const uint8_t hi = static_cast<uint8_t>(bytes[2 + i * 2]);
                const uint8_t lo = static_cast<uint8_t>(bytes[2 + i * 2 + 1]);
                w[i]             = static_cast<wchar_t>((static_cast<uint16_t>(hi) << 8) | static_cast<uint16_t>(lo));
            }
            return Utf8FromUtf16(w);
        }

        return std::string(bytes);
    };

    const auto cssRgb = [](COLORREF c) { return std::format("rgb({},{},{})", GetRValue(c), GetGValue(c), GetBValue(c)); };

    const COLORREF bg    = hasTheme ? ColorRefFromArgb(theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg    = hasTheme ? ColorRefFromArgb(theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
    const COLORREF selBg = hasTheme ? ColorRefFromArgb(theme.selectionBackgroundArgb) : GetSysColor(COLOR_HIGHLIGHT);
    const COLORREF selFg = hasTheme ? ColorRefFromArgb(theme.selectionTextArgb) : GetSysColor(COLOR_HIGHLIGHTTEXT);
    const COLORREF accent =
        hasTheme ? ResolveAccentColor(theme, result->path.empty() ? self->_metaId : std::wstring_view(result->path)) : GetSysColor(COLOR_HIGHLIGHT);

    const std::string themeObj =
        std::format("{{bg:'{}',fg:'{}',selBg:'{}',selFg:'{}',accent:'{}'}}", cssRgb(bg), cssRgb(fg), cssRgb(selBg), cssRgb(selFg), cssRgb(accent));

    const std::wstring leafW  = LeafNameFromPath(result->path);
    const std::wstring titleW = leafW.empty() ? self->_metaName : std::format(L"{} - {}", leafW, self->_metaName);
    result->title             = titleW;

    result->statusMessage.clear();

    auto postBack = [&](bool cleanupTempOnFailure) noexcept
    {
        const HWND hwnd = result->hwnd;
        std::optional<std::filesystem::path> extractedPath;
        if (cleanupTempOnFailure)
        {
            extractedPath = result->extractedWin32Path;
        }

        if (! hwnd || ! PostMessagePayload(hwnd, kAsyncLoadCompleteMessage, 0, std::move(result)))
        {
            if (cleanupTempOnFailure && extractedPath.has_value() && ! extractedPath->empty())
            {
                std::error_code ec;
                static_cast<void>(std::filesystem::remove(*extractedPath, ec));
            }
            return;
        }
    };

    if (! fileSystem)
    {
        result->hr            = E_FAIL;
        result->statusMessage = L"File system unavailable.";
        postBack(false);
        return;
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    const HRESULT fileIoHr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (FAILED(fileIoHr) || ! fileIo)
    {
        result->hr            = FAILED(fileIoHr) ? fileIoHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        result->statusMessage = L"Active file system does not support file I/O.";
        postBack(false);
        return;
    }

    if (kind == ViewerWebKind::Web)
    {
        if (IsProbablyWin32Path(result->path))
        {
            result->hr = S_OK;
            postBack(false);
            return;
        }

        wil::com_ptr<IFileReader> reader;
        const HRESULT openHr = fileIo->CreateFileReader(result->path.c_str(), reader.put());
        if (FAILED(openHr) || ! reader)
        {
            result->hr            = FAILED(openHr) ? openHr : E_FAIL;
            result->statusMessage = L"Failed to open file for viewing.";
            postBack(false);
            return;
        }

        wchar_t tempDir[MAX_PATH]{};
        const DWORD tempDirLen = GetTempPathW(static_cast<DWORD>(std::size(tempDir)), tempDir);
        if (tempDirLen == 0 || tempDirLen >= std::size(tempDir))
        {
            result->hr            = HRESULT_FROM_WIN32(GetLastError());
            result->statusMessage = L"Failed to get temp folder.";
            postBack(false);
            return;
        }

        wchar_t tempName[MAX_PATH]{};
        if (GetTempFileNameW(tempDir, L"rsw", 0, tempName) == 0)
        {
            result->hr            = HRESULT_FROM_WIN32(GetLastError());
            result->statusMessage = L"Failed to create temp file.";
            postBack(false);
            return;
        }

        std::filesystem::path tempPath(tempName);
        const std::wstring ext = std::filesystem::path(result->path).extension().wstring();
        if (! ext.empty())
        {
            std::filesystem::path newPath = tempPath;
            newPath.replace_extension(ext);
            if (MoveFileExW(tempPath.c_str(), newPath.c_str(), MOVEFILE_REPLACE_EXISTING) != 0)
            {
                tempPath = std::move(newPath);
            }
        }

        wil::unique_handle outFile(CreateFileW(tempPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! outFile)
        {
            result->hr            = HRESULT_FROM_WIN32(GetLastError());
            result->statusMessage = L"Failed to write temp file.";
            std::error_code ec;
            static_cast<void>(std::filesystem::remove(tempPath, ec));
            postBack(false);
            return;
        }

        std::vector<uint8_t> buffer(256u * 1024u);
        for (;;)
        {
            unsigned long read   = 0;
            const HRESULT readHr = reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &read);
            if (FAILED(readHr))
            {
                result->hr            = readHr;
                result->statusMessage = L"Failed to read file.";
                std::error_code ec;
                static_cast<void>(std::filesystem::remove(tempPath, ec));
                postBack(false);
                return;
            }

            if (read == 0)
            {
                break;
            }

            DWORD written = 0;
            if (WriteFile(outFile.get(), buffer.data(), read, &written, nullptr) == 0 || written != read)
            {
                result->hr            = HRESULT_FROM_WIN32(GetLastError());
                result->statusMessage = L"Failed to write temp file.";
                std::error_code ec;
                static_cast<void>(std::filesystem::remove(tempPath, ec));
                postBack(false);
                return;
            }
        }

        result->extractedWin32Path = tempPath;
        result->hr                 = S_OK;
        postBack(true);
        return;
    }

    // JSON/Markdown: load file into memory (UTF-8/UTF-16 with BOM supported).
    wil::com_ptr<IFileReader> reader;
    const HRESULT openHr = fileIo->CreateFileReader(result->path.c_str(), reader.put());
    if (FAILED(openHr) || ! reader)
    {
        result->hr            = FAILED(openHr) ? openHr : E_FAIL;
        result->statusMessage = L"Failed to open file for viewing.";
        postBack(false);
        return;
    }

    unsigned __int64 sizeBytes = 0;
    const HRESULT sizeHr       = reader->GetSize(&sizeBytes);
    if (FAILED(sizeHr))
    {
        result->hr            = sizeHr;
        result->statusMessage = L"Failed to read file size.";
        postBack(false);
        return;
    }

    const uint64_t maxBytes = static_cast<uint64_t>(config.maxDocumentMiB) * 1024ull * 1024ull;
    if (static_cast<uint64_t>(sizeBytes) > maxBytes)
    {
        result->hr = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        result->statusMessage =
            std::format(L"File is too large ({}), limit is {}.", FormatBytesCompact(static_cast<uint64_t>(sizeBytes)), FormatBytesCompact(maxBytes));
        postBack(false);
        return;
    }

    std::string bytes;
    bytes.reserve(static_cast<size_t>(sizeBytes));
    std::vector<uint8_t> buffer(256u * 1024u);
    for (;;)
    {
        unsigned long read   = 0;
        const HRESULT readHr = reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &read);
        if (FAILED(readHr))
        {
            result->hr            = readHr;
            result->statusMessage = L"Failed to read file.";
            postBack(false);
            return;
        }

        if (read == 0)
        {
            break;
        }

        bytes.append(reinterpret_cast<const char*>(buffer.data()), reinterpret_cast<const char*>(buffer.data() + read));
        if (bytes.size() > maxBytes)
        {
            result->hr = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
            result->statusMessage =
                std::format(L"File is too large ({}), limit is {}.", FormatBytesCompact(static_cast<uint64_t>(bytes.size())), FormatBytesCompact(maxBytes));
            postBack(false);
            return;
        }
    }

    const std::string textUtf8 = normalizeTextUtf8(bytes);

    if (kind == ViewerWebKind::Json)
    {
        std::string jsonMutable(textUtf8);
        yyjson_read_err err{};
        unique_yyjson_doc doc(yyjson_read_opts(jsonMutable.data(), jsonMutable.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err));
        if (! doc)
        {
            result->hr            = E_FAIL;
            result->statusMessage = L"Failed to parse JSON/JSON5 document.";
            postBack(false);
            return;
        }

        size_t prettyLen = 0;
        unique_malloc_string pretty(yyjson_write_opts(doc.get(), YYJSON_WRITE_PRETTY | YYJSON_WRITE_ESCAPE_UNICODE, nullptr, &prettyLen, nullptr));
        if (! pretty)
        {
            result->hr            = E_OUTOFMEMORY;
            result->statusMessage = L"Failed to format JSON document.";
            postBack(false);
            return;
        }

        const std::string prettyJson(pretty.get(), prettyLen);
        const std::string jsonB64 = base64Encode(prettyJson);

        if (config.jsonViewMode == JsonViewMode::Pretty)
        {
            std::string highlightJs = ResourceBytesToString(g_hInstance, IDR_VIEWERWEB_HIGHLIGHT_JS);

            const COLORREF codeBg   = BlendColor(bg, fg, theme.darkMode ? 20u : 10u);
            const COLORREF border   = BlendColor(bg, fg, theme.darkMode ? 35u : 45u);
            const COLORREF mutedFg  = BlendColor(bg, fg, 140u);
            const COLORREF strColor = BlendColor(accent, fg, 60u);
            const COLORREF numColor = BlendColor(accent, fg, 90u);
            const COLORREF litColor = BlendColor(accent, fg, 120u);

            std::string html;
            html.reserve(highlightJs.size() + jsonB64.size() + 8192);
            html += "<!doctype html><html><head><meta charset=\"utf-8\">";
            html += "<style>";
            html += ":root{--rs-bg:" + cssRgb(bg) + ";--rs-fg:" + cssRgb(fg) + ";--rs-sel-bg:" + cssRgb(selBg) + ";--rs-sel-fg:" + cssRgb(selFg) +
                    ";--rs-accent:" + cssRgb(accent) + ";--rs-code-bg:" + cssRgb(codeBg) + ";--rs-border:" + cssRgb(border) +
                    ";--rs-muted-fg:" + cssRgb(mutedFg) + ";--rs-string:" + cssRgb(strColor) + ";--rs-number:" + cssRgb(numColor) +
                    ";--rs-literal:" + cssRgb(litColor) + ";}";
            html += "html,body{height:100%;margin:0;}body{background:var(--rs-bg);color:var(--rs-fg);font-family:Segoe UI,sans-serif;}";
            html += "::selection{background:var(--rs-sel-bg);color:var(--rs-sel-fg);}#app{height:100%;box-sizing:border-box;padding:12px;display:flex;}";
            html += "pre{flex:1;margin:0;background:var(--rs-code-bg);border:1px solid var(--rs-border);padding:12px;overflow:auto;border-radius:6px;}";
            html += "code{font-family:Consolas,ui-monospace,monospace;font-size:13px;line-height:1.45;}";
            html += ".hljs{background:transparent;}";
            html += ".hljs-attr{color:var(--rs-accent);} .hljs-string{color:var(--rs-string);} .hljs-number{color:var(--rs-number);} "
                    ".hljs-literal{color:var(--rs-literal);}";
            html += ".hljs-punctuation,.hljs-brace{color:var(--rs-muted-fg);} .hljs-comment{opacity:0.8;}";
            html += "</style></head><body><div id=\"app\"><pre><code id=\"code\" class=\"language-json\"></code></pre></div>";
            html += "<script>";
            html += highlightJs;
            html += "</script><script>";
            html += "(() => {";
            html += "const initialTheme=" + themeObj + ";";
            html +=
                "function parseRgb(s){const m=/rgb\\((\\d+),(\\d+),(\\d+)\\)/.exec(s.replace(/\\s+/g,''));return m?{r:+m[1],g:+m[2],b:+m[3]}:{r:0,g:0,b:0};}";
            html += "function rgb(c){return `rgb(${c.r},${c.g},${c.b})`;}";
            html += "function blend(u,o,a){const inv=255-a;return "
                    "{r:Math.round((u.r*inv+o.r*a)/255),g:Math.round((u.g*inv+o.g*a)/255),b:Math.round((u.b*inv+o.b*a)/255)};}";
            html += "function luma(c){return (c.r*299+c.g*587+c.b*114)/1000;}";
            html += "function applyTheme(t){const "
                    "r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);r.setProperty('--rs-sel-bg',t.selBg);r."
                    "setProperty('--rs-sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);const "
                    "bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const "
                    "dark=luma(bg)<128;r.setProperty('--rs-code-bg',rgb(blend(bg,fg,dark?20:10)));r.setProperty('--rs-border',rgb(blend(bg,fg,dark?35:45)));r."
                    "setProperty('--rs-muted-fg',rgb(blend(bg,fg,140)));r.setProperty('--rs-string',rgb(blend(acc,fg,60)));r.setProperty('--rs-number',rgb("
                    "blend(acc,fg,90)));r.setProperty('--rs-literal',rgb(blend(acc,fg,120)));}";
            html += "function decodeUtf8(b64){const bin=atob(b64);const bytes=new Uint8Array(bin.length);for(let "
                    "i=0;i<bin.length;i++){bytes[i]=bin.charCodeAt(i);}return new TextDecoder('utf-8').decode(bytes);}";
            html += "const code=document.getElementById('code');";
            html += "code.textContent=decodeUtf8('" + jsonB64 + "');";
            html += "window.RS={applyTheme:applyTheme};";
            html += "applyTheme(initialTheme);";
            html += "try{hljs.highlightElement(code);}catch(e){}";
            html += "})();";
            html += "</script></body></html>";

            result->utf8 = std::move(html);
            result->hr   = S_OK;
            postBack(false);
            return;
        }

        std::string jsonEditorJs   = ResourceBytesToString(g_hInstance, IDR_VIEWERWEB_JSONEDITOR_JS);
        std::string jsonEditorCss  = ResourceBytesToString(g_hInstance, IDR_VIEWERWEB_JSONEDITOR_CSS);
        const std::string iconsSvg = ResourceBytesToString(g_hInstance, IDR_VIEWERWEB_JSONEDITOR_ICONS_SVG);
        const std::string iconsB64 = base64Encode(iconsSvg);
        const std::string iconsUrl = std::string("data:image/svg+xml;base64,") + iconsB64;
        replaceAll(jsonEditorCss, "./img/jsoneditor-icons.svg", iconsUrl);
        replaceAll(jsonEditorCss, "img/jsoneditor-icons.svg", iconsUrl);

        const COLORREF border  = BlendColor(bg, fg, theme.darkMode ? 45u : 80u);
        const COLORREF mutedFg = BlendColor(bg, fg, 140u);

        std::string html;
        html.reserve(jsonEditorJs.size() + jsonEditorCss.size() + jsonB64.size() + 8192);
        html += "<!doctype html><html><head><meta charset=\"utf-8\">";
        html += "<style>";
        html += ":root{--rs-bg:" + cssRgb(bg) + ";--rs-fg:" + cssRgb(fg) + ";--rs-sel-bg:" + cssRgb(selBg) + ";--rs-sel-fg:" + cssRgb(selFg) +
                ";--rs-accent:" + cssRgb(accent) + ";--rs-border:" + cssRgb(border) + ";--rs-muted-fg:" + cssRgb(mutedFg) + ";}";
        html += "html,body{height:100%;margin:0;}body{background:var(--rs-bg);color:var(--rs-fg);font-family:Segoe UI,sans-serif;}#app{height:100%;}";
        html += jsonEditorCss;
        html += "html,body{background:var(--rs-bg)!important;color:var(--rs-fg)!important;}#app{height:100%!important;}";
        html += ".jsoneditor{border:none!important;height:100%!important;background:var(--rs-bg)!important;color:var(--rs-fg)!important;}";
        html += ".jsoneditor-frame{background:var(--rs-bg)!important;border:1px solid var(--rs-border)!important;}";
        html += ".jsoneditor-outer,.jsoneditor-inner,.jsoneditor-tree,.jsoneditor-tree-inner,.jsoneditor-text,.jsoneditor-text "
                "textarea{background:var(--rs-bg)!important;color:var(--rs-fg)!important;}";
        html += ".jsoneditor-field{color:var(--rs-fg)!important;}";
        html += ".jsoneditor-value.jsoneditor-object,.jsoneditor-value.jsoneditor-array,.jsoneditor-value.jsoneditor-null{color:var(--rs-muted-fg)!important;}";
        html += ".jsoneditor-selected,.jsoneditor-highlight-active{background-color:var(--rs-sel-bg)!important;color:var(--rs-sel-fg)!important;}";
        html += ".jsoneditor-highlight{background-color:var(--rs-sel-bg)!important;}";
        html += ".jsoneditor .autocomplete.dropdown{background:var(--rs-bg)!important;border:1px solid var(--rs-border)!important;}";
        html += ".jsoneditor .autocomplete.dropdown .item{color:var(--rs-fg)!important;}";
        html += ".jsoneditor .autocomplete.dropdown .item.hover{background-color:var(--rs-sel-bg)!important;color:var(--rs-sel-fg)!important;}";
        html += ".jsoneditor-contextmenu .jsoneditor-menu{background:var(--rs-bg)!important;border:1px solid var(--rs-border)!important;}";
        html += ".jsoneditor-contextmenu .jsoneditor-menu button{color:var(--rs-fg)!important;}";
        html += ".jsoneditor-contextmenu .jsoneditor-menu button:hover{background-color:var(--rs-sel-bg)!important;color:var(--rs-sel-fg)!important;}";
        html += ".jsoneditor-contextmenu .jsoneditor-separator{border-top:1px solid var(--rs-border)!important;}";
        html += ".jsoneditor-contextmenu .jsoneditor-menu button.jsoneditor-expand{border-left:1px solid var(--rs-border)!important;}";
        html += "</style></head><body><div id=\"app\"></div>";
        html += "<script>";
        html += jsonEditorJs;
        html += "</script><script>";
        html += "(() => {";
        html += "const initialTheme=" + themeObj + ";";
        html += "function parseRgb(s){const m=/rgb\\((\\d+),(\\d+),(\\d+)\\)/.exec(s.replace(/\\s+/g,''));return m?{r:+m[1],g:+m[2],b:+m[3]}:{r:0,g:0,b:0};}";
        html += "function rgb(c){return `rgb(${c.r},${c.g},${c.b})`;}";
        html += "function blend(u,o,a){const inv=255-a;return "
                "{r:Math.round((u.r*inv+o.r*a)/255),g:Math.round((u.g*inv+o.g*a)/255),b:Math.round((u.b*inv+o.b*a)/255)};}";
        html += "function luma(c){return (c.r*299+c.g*587+c.b*114)/1000;}";
        html +=
            "function applyTheme(t){const "
            "r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);r.setProperty('--rs-sel-bg',t.selBg);r.setProperty('-"
            "-rs-sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);const bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const "
            "dark=luma(bg)<128;r.setProperty('--rs-border',rgb(blend(bg,fg,dark?45:80)));r.setProperty('--rs-muted-fg',rgb(blend(bg,fg,140)));r.setProperty('--"
            "rs-string',rgb(blend(acc,fg,60)));r.setProperty('--rs-number',rgb(blend(acc,fg,90)));r.setProperty('--rs-literal',rgb(blend(acc,fg,120)));}";
        html += "function decodeUtf8(b64){const bin=atob(b64);const bytes=new Uint8Array(bin.length);for(let "
                "i=0;i<bin.length;i++){bytes[i]=bin.charCodeAt(i);}return new TextDecoder('utf-8').decode(bytes);}";
        html += "const jsonText=decodeUtf8('" + jsonB64 + "');";
        html += "const container=document.getElementById('app');";
        html += "const options={mode:'tree',modes:['tree','view'],onEditable:()=>false,mainMenuBar:false,navigationBar:false,statusBar:false};";
        html += "const editor=new JSONEditor(container,options);";
        html += "window.RS={applyTheme:applyTheme,expandAll:()=>editor.expandAll(),collapseAll:()=>editor.collapseAll()};";
        html += "applyTheme(initialTheme);";
        html += "try{editor.set(JSON.parse(jsonText));}catch(e){editor.set({error:String(e)});}";
        html += "})();";
        html += "</script></body></html>";

        result->utf8 = std::move(html);
        result->hr   = S_OK;
        postBack(false);
        return;
    }

    // Markdown
    const std::string markdownB64 = base64Encode(textUtf8);
    std::string markdownItJs      = ResourceBytesToString(g_hInstance, IDR_VIEWERWEB_MARKDOWNIT_JS);
    std::string highlightJs       = ResourceBytesToString(g_hInstance, IDR_VIEWERWEB_HIGHLIGHT_JS);

    const COLORREF codeBg      = BlendColor(bg, fg, theme.darkMode ? 20u : 10u);
    const COLORREF border      = BlendColor(bg, fg, theme.darkMode ? 35u : 45u);
    const COLORREF mutedFg     = BlendColor(bg, fg, 140u);
    const COLORREF stringColor = BlendColor(accent, fg, 60u);
    const COLORREF numberColor = BlendColor(accent, fg, 90u);

    std::string html;
    html.reserve(markdownItJs.size() + highlightJs.size() + markdownB64.size() + 8192);
    html += "<!doctype html><html><head><meta charset=\"utf-8\">";
    html += "<style>";
    html += ":root{--rs-bg:" + cssRgb(bg) + ";--rs-fg:" + cssRgb(fg) + ";--rs-sel-bg:" + cssRgb(selBg) + ";--rs-sel-fg:" + cssRgb(selFg) +
            ";--rs-accent:" + cssRgb(accent) + ";--rs-code-bg:" + cssRgb(codeBg) + ";--rs-border:" + cssRgb(border) + ";--rs-muted-fg:" + cssRgb(mutedFg) +
            ";--rs-string:" + cssRgb(stringColor) + ";--rs-number:" + cssRgb(numberColor) + ";}";
    html += "html,body{height:100%;margin:0;}body{background:var(--rs-bg);color:var(--rs-fg);font-family:Segoe UI,sans-serif;}";
    html += "#app{max-width:100%;padding:16px;box-sizing:border-box;}a{color:var(--rs-accent);}";
    html += "pre{background:var(--rs-code-bg);border:1px solid var(--rs-border);padding:12px;overflow:auto;border-radius:6px;}";
    html += "code{font-family:Consolas,ui-monospace,monospace;}";
    html += "table{border-collapse:collapse;}th,td{border:1px solid var(--rs-border);padding:6px 10px;}";
    html += ".rs-source{white-space:pre;overflow:auto;font-family:Consolas,ui-monospace,monospace;}";
    html += ".hljs-comment{opacity:0.8;}.hljs-keyword,.hljs-selector-tag{color:var(--rs-accent);}";
    html += ".hljs-string{color:var(--rs-string);}.hljs-number{color:var(--rs-number);}.hljs-punctuation,.hljs-brace{color:var(--rs-muted-fg);}";
    html += "</style></head><body><div id=\"app\"></div>";
    html += "<script>";
    html += markdownItJs;
    html += "</script><script>";
    html += highlightJs;
    html += "</script><script>";
    html += "(() => {";
    html += "const initialTheme=" + themeObj + ";";
    html += "function parseRgb(s){const m=/rgb\\((\\d+),(\\d+),(\\d+)\\)/.exec(s.replace(/\\s+/g,''));return m?{r:+m[1],g:+m[2],b:+m[3]}:{r:0,g:0,b:0};}";
    html += "function rgb(c){return `rgb(${c.r},${c.g},${c.b})`;}";
    html += "function blend(u,o,a){const inv=255-a;return "
            "{r:Math.round((u.r*inv+o.r*a)/255),g:Math.round((u.g*inv+o.g*a)/255),b:Math.round((u.b*inv+o.b*a)/255)};}";
    html += "function luma(c){return (c.r*299+c.g*587+c.b*114)/1000;}";
    html +=
        "function applyTheme(t){const "
        "r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);r.setProperty('--rs-sel-bg',t.selBg);r.setProperty('--rs-"
        "sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);const bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const "
        "dark=luma(bg)<128;r.setProperty('--rs-code-bg',rgb(blend(bg,fg,dark?20:10)));r.setProperty('--rs-border',rgb(blend(bg,fg,dark?35:45)));r.setProperty('"
        "--rs-muted-fg',rgb(blend(bg,fg,140)));r.setProperty('--rs-string',rgb(blend(acc,fg,60)));r.setProperty('--rs-number',rgb(blend(acc,fg,90)));}";
    html += "function decodeUtf8(b64){const bin=atob(b64);const bytes=new Uint8Array(bin.length);for(let "
            "i=0;i<bin.length;i++){bytes[i]=bin.charCodeAt(i);}return new TextDecoder('utf-8').decode(bytes);}";
    html += "const src=decodeUtf8('" + markdownB64 + "');";
    html += "const container=document.getElementById('app');";
    html += "let showSource=" + std::string(markdownShowSource ? "true" : "false") + ";";
    html += "const md=window.markdownit({html:false,linkify:true,typographer:true});";
    html += "function "
            "render(){if(showSource){container.className='rs-source';container.textContent=src;return;}container.className='';container.innerHTML=md.render("
            "src);document.querySelectorAll('pre code').forEach((el)=>{try{hljs.highlightElement(el);}catch(e){}});}";
    html += "window.RS={applyTheme:applyTheme,setShowSource:(v)=>{showSource=!!v;render();}};";
    html += "applyTheme(initialTheme);render();";
    html += "})();";
    html += "</script></body></html>";

    result->utf8 = std::move(html);
    result->hr   = S_OK;
    postBack(false);
}

HRESULT ViewerWeb::CommandSaveAs(HWND hwnd) noexcept
{
    if (_currentPath.empty() || ! _fileSystem)
    {
        return S_FALSE;
    }

    const std::wstring suggested = LeafNameFromPath(_currentPath);
    const auto dest              = ShowSaveAsDialog(hwnd, suggested);
    if (! dest.has_value())
    {
        return S_FALSE;
    }

    wil::unique_handle outFile(CreateFileW(dest->c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! outFile)
    {
        const DWORD lastError = GetLastError();
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, L"Save As failed.");
        return HRESULT_FROM_WIN32(lastError);
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    const HRESULT ioHr = _fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (FAILED(ioHr) || ! fileIo)
    {
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, L"Save As failed (file system I/O not supported).");
        return FAILED(ioHr) ? ioHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    wil::com_ptr<IFileReader> reader;
    const HRESULT openHr = fileIo->CreateFileReader(_currentPath.c_str(), reader.put());
    if (FAILED(openHr) || ! reader)
    {
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, L"Save As failed (unable to open file).");
        return FAILED(openHr) ? openHr : E_FAIL;
    }

    unsigned __int64 pos = 0;
    const HRESULT seekHr = reader->Seek(0, FILE_BEGIN, &pos);
    if (FAILED(seekHr))
    {
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, L"Save As failed (seek failed).");
        return seekHr;
    }

    std::vector<uint8_t> buffer(256u * 1024u);
    for (;;)
    {
        unsigned long read   = 0;
        const HRESULT readHr = reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &read);
        if (FAILED(readHr))
        {
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, L"Save As failed (read failed).");
            return readHr;
        }

        if (read == 0)
        {
            break;
        }

        DWORD written = 0;
        if (WriteFile(outFile.get(), buffer.data(), read, &written, nullptr) == 0 || written != read)
        {
            const DWORD lastError = GetLastError();
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, L"Save As failed (write failed).");
            return HRESULT_FROM_WIN32(lastError);
        }
    }

    return S_OK;
}

void ViewerWeb::CommandFind(HWND hwnd) noexcept
{
    if (_hFindDialog && IsWindow(_hFindDialog.get()))
    {
        ShowWindow(_hFindDialog.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hFindDialog.get()));
        return;
    }

    if (! _findQuery.empty())
    {
        wcsncpy_s(_findBuffer.data(), _findBuffer.size(), _findQuery.c_str(), _TRUNCATE);
    }
    else
    {
        _findBuffer[0] = L'\0';
    }

    _findReplace               = {};
    _findReplace.lStructSize   = sizeof(_findReplace);
    _findReplace.hwndOwner     = hwnd;
    _findReplace.lpstrFindWhat = _findBuffer.data();
    _findReplace.wFindWhatLen  = static_cast<WORD>(_findBuffer.size());
    _findReplace.Flags         = FR_DOWN;

    HWND dlg = FindTextW(&_findReplace);
    if (! dlg)
    {
        return;
    }

    _hFindDialog.reset(dlg);
}

void ViewerWeb::CommandFindNext(HWND hwnd) noexcept
{
    if (_findQuery.empty())
    {
        CommandFind(hwnd);
        return;
    }

    if (! _webView)
    {
        return;
    }

    const std::wstring queryEsc = EscapeJavaScriptString(_findQuery);
    const std::wstring script =
        std::format(L"(function(){{try{{return window.find('{}',false,false,true,false,true,false);}}catch(e){{return false;}}}})();", queryEsc);
    static_cast<void>(_webView->ExecuteScript(script.c_str(), nullptr));
}

void ViewerWeb::CommandFindPrevious(HWND hwnd) noexcept
{
    if (_findQuery.empty())
    {
        CommandFind(hwnd);
        return;
    }

    if (! _webView)
    {
        return;
    }

    const std::wstring queryEsc = EscapeJavaScriptString(_findQuery);
    const std::wstring script =
        std::format(L"(function(){{try{{return window.find('{}',false,true,true,false,true,false);}}catch(e){{return false;}}}})();", queryEsc);
    static_cast<void>(_webView->ExecuteScript(script.c_str(), nullptr));
}

void ViewerWeb::CommandCopyUrl(HWND hwnd) noexcept
{
    std::wstring toCopy;

    if (_webView)
    {
        wil::unique_cotaskmem_string source;
        if (SUCCEEDED(_webView->get_Source(source.put())) && source && source.get()[0] != L'\0')
        {
            const std::wstring_view src(source.get());
            if (! StartsWithNoCase(src, L"about:"))
            {
                toCopy.assign(src);
            }
        }
    }

    if (toCopy.empty())
    {
        if (_kind == ViewerWebKind::Web)
        {
            if (_tempExtractedPath.has_value() && ! _tempExtractedPath->empty())
            {
                toCopy = UrlFromFilePath(_tempExtractedPath->wstring());
            }
            else if (IsProbablyWin32Path(_currentPath))
            {
                toCopy = UrlFromFilePath(_currentPath);
            }
        }
    }

    if (toCopy.empty())
    {
        toCopy = _currentPath;
    }

    static_cast<void>(CopyUnicodeTextToClipboard(hwnd, toCopy));
}

void ViewerWeb::CommandOpenExternal(HWND hwnd) noexcept
{
    std::wstring url;

    if (_webView)
    {
        wil::unique_cotaskmem_string source;
        if (SUCCEEDED(_webView->get_Source(source.put())) && source && source.get()[0] != L'\0')
        {
            const std::wstring_view src(source.get());
            if (! StartsWithNoCase(src, L"about:"))
            {
                url.assign(src);
            }
        }
    }

    if (url.empty())
    {
        if (_kind == ViewerWebKind::Web)
        {
            if (_tempExtractedPath.has_value() && ! _tempExtractedPath->empty())
            {
                url = UrlFromFilePath(_tempExtractedPath->wstring());
            }
            else if (IsProbablyWin32Path(_currentPath))
            {
                url = UrlFromFilePath(_currentPath);
            }
        }
    }

    if (url.empty())
    {
        return;
    }

    const HINSTANCE res = ShellExecuteW(hwnd, L"open", url.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(res) <= 32)
    {
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, L"Failed to open in browser.");
    }
}

void ViewerWeb::CommandZoom(double factor) noexcept
{
    if (! _webViewController || factor <= 0.0)
    {
        return;
    }

    double current = 1.0;
    static_cast<void>(_webViewController->get_ZoomFactor(&current));

    const double next = std::clamp(current * factor, 0.25, 5.0);
    static_cast<void>(_webViewController->put_ZoomFactor(next));
}

void ViewerWeb::CommandZoomIn() noexcept
{
    CommandZoom(1.10);
}

void ViewerWeb::CommandZoomOut() noexcept
{
    CommandZoom(1.0 / 1.10);
}

void ViewerWeb::CommandZoomReset() noexcept
{
    if (_webViewController)
    {
        static_cast<void>(_webViewController->put_ZoomFactor(1.0));
    }
}

void ViewerWeb::CommandToggleDevTools() noexcept
{
    if (! _config.devToolsEnabled)
    {
        ShowHostAlert(_hWnd.get(), HOST_ALERT_WARNING, L"DevTools is disabled in plugin settings.");
        return;
    }

    if (_webView)
    {
        static_cast<void>(_webView->OpenDevToolsWindow());
    }
}

void ViewerWeb::CommandJsonExpandAll() noexcept
{
    if (_kind != ViewerWebKind::Json || _config.jsonViewMode != JsonViewMode::Tree || ! _webView)
    {
        return;
    }

    static_cast<void>(_webView->ExecuteScript(L"(function(){try{if(window.RS&&window.RS.expandAll){window.RS.expandAll();}}catch(e){}})();", nullptr));
}

void ViewerWeb::CommandJsonCollapseAll() noexcept
{
    if (_kind != ViewerWebKind::Json || _config.jsonViewMode != JsonViewMode::Tree || ! _webView)
    {
        return;
    }

    static_cast<void>(_webView->ExecuteScript(L"(function(){try{if(window.RS&&window.RS.collapseAll){window.RS.collapseAll();}}catch(e){}})();", nullptr));
}

void ViewerWeb::CommandMarkdownToggleSource() noexcept
{
    if (_kind != ViewerWebKind::Markdown)
    {
        return;
    }

    _markdownShowSource = ! _markdownShowSource;

    if (_hWnd)
    {
        HMENU menu = GetMenu(_hWnd.get());
        if (menu)
        {
            CheckMenuItem(
                menu, IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE, static_cast<UINT>(MF_BYCOMMAND | (_markdownShowSource ? MF_CHECKED : MF_UNCHECKED)));
            DrawMenuBar(_hWnd.get());
        }
    }

    if (_webView)
    {
        const std::wstring script = std::format(L"(function(){{try{{if(window.RS&&window.RS.setShowSource){{window.RS.setShowSource({});}}}}catch(e){{}}}})();",
                                                _markdownShowSource ? L"true" : L"false");
        static_cast<void>(_webView->ExecuteScript(script.c_str(), nullptr));
    }
}

namespace
{
bool CopyUnicodeTextToClipboard(HWND hwnd, std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return false;
    }

    if (OpenClipboard(hwnd) == 0)
    {
        return false;
    }
    auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });

    if (EmptyClipboard() == 0)
    {
        return false;
    }

    const size_t bytes = (text.size() + 1) * sizeof(wchar_t);
    wil::unique_hglobal storage(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! storage)
    {
        return false;
    }

    void* mem = GlobalLock(storage.get());
    if (! mem)
    {
        return false;
    }

    memcpy(mem, text.data(), text.size() * sizeof(wchar_t));
    static_cast<wchar_t*>(mem)[text.size()] = L'\0';
    GlobalUnlock(storage.get());

    if (SetClipboardData(CF_UNICODETEXT, storage.get()) == nullptr)
    {
        return false;
    }

    storage.release();
    return true;
}

[[nodiscard]] bool IsProbablyWin32Path(std::wstring_view path) noexcept
{
    if (path.size() >= 3 && ((path[0] >= L'A' && path[0] <= L'Z') || (path[0] >= L'a' && path[0] <= L'z')) && path[1] == L':' &&
        (path[2] == L'\\' || path[2] == L'/'))
    {
        return true;
    }

    if (StartsWithNoCase(path, L"\\\\") || StartsWithNoCase(path, L"//"))
    {
        return true;
    }

    return false;
}

std::optional<std::filesystem::path> ShowSaveAsDialog(HWND hwnd, std::wstring_view suggestedFileName) noexcept
{
    wil::com_ptr<IFileSaveDialog> dialog;
    const HRESULT hr = CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(hr) || ! dialog)
    {
        return std::nullopt;
    }

    DWORD options = 0;
    static_cast<void>(dialog->GetOptions(&options));
    options |= FOS_FORCEFILESYSTEM | FOS_OVERWRITEPROMPT | FOS_PATHMUSTEXIST;
    static_cast<void>(dialog->SetOptions(options));

    if (! suggestedFileName.empty())
    {
        static_cast<void>(dialog->SetFileName(std::wstring(suggestedFileName).c_str()));
    }

    const HRESULT showHr = dialog->Show(hwnd);
    if (showHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
    {
        return std::nullopt;
    }
    if (FAILED(showHr))
    {
        return std::nullopt;
    }

    wil::com_ptr<IShellItem> item;
    const HRESULT itemHr = dialog->GetResult(item.put());
    if (FAILED(itemHr) || ! item)
    {
        return std::nullopt;
    }

    wil::unique_cotaskmem_string path;
    const HRESULT nameHr = item->GetDisplayName(SIGDN_FILESYSPATH, path.put());
    if (FAILED(nameHr) || ! path)
    {
        return std::nullopt;
    }

    return std::filesystem::path(path.get());
}

[[nodiscard]] std::wstring EscapeJavaScriptString(std::wstring_view text) noexcept
{
    std::wstring out;
    out.reserve(text.size() + 16);
    for (wchar_t ch : text)
    {
        switch (ch)
        {
            case L'\\': out += L"\\\\"; break;
            case L'\'': out += L"\\'"; break;
            case L'\"': out += L"\\\""; break;
            case L'\r': out += L"\\r"; break;
            case L'\n': out += L"\\n"; break;
            case L'\t': out += L"\\t"; break;
            default:
                if (ch < 0x20)
                {
                    out += std::format(L"\\x{:02X}", static_cast<unsigned int>(ch));
                }
                else
                {
                    out.push_back(ch);
                }
                break;
        }
    }
    return out;
}
} // namespace
