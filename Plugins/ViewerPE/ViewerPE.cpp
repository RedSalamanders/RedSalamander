#include "ViewerPE.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <format>
#include <limits>
#include <memory>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include <commctrl.h>
#include <d2d1.h>
#include <d2d1helper.h>
#include <dwmapi.h>
#include <dwrite.h>
#include <shobjidl.h>
#include <uxtheme.h>

#define _PEPARSE_WINDOWS_CONFLICTS
#include <pe-parse/parse.h>

#pragma comment(lib, "comctl32")
#pragma comment(lib, "d2d1")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "dwrite")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "uxtheme")

#include "Helpers.h"
#include "WindowMessages.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

namespace
{
constexpr UINT kAsyncParseCompleteMessage = WndMsg::kViewerPeAsyncParseComplete;

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

constexpr float kOuterPaddingDip    = 12.0f;
constexpr float kCardRadiusDip      = 10.0f;
constexpr float kInnerPaddingDip    = 12.0f;
constexpr float kHeaderGapDip       = 10.0f;
constexpr float kScrollWheelStepDip = 48.0f;
constexpr float kScrollLineStepDip  = 24.0f;

constexpr char kViewerPESchemaJson[] = R"json({
  "version": 1,
  "title": "PE Viewer",
  "fields": []
})json";

constexpr std::array<std::wstring_view, 16> kDataDirectoryNames = {
    std::wstring_view(L"Export"),
    std::wstring_view(L"Import"),
    std::wstring_view(L"Resource"),
    std::wstring_view(L"Exception"),
    std::wstring_view(L"Security"),
    std::wstring_view(L"Base Relocation"),
    std::wstring_view(L"Debug"),
    std::wstring_view(L"Architecture"),
    std::wstring_view(L"GlobalPtr"),
    std::wstring_view(L"TLS"),
    std::wstring_view(L"Load Config"),
    std::wstring_view(L"Bound Import"),
    std::wstring_view(L"IAT"),
    std::wstring_view(L"Delay Import"),
    std::wstring_view(L"COM Descriptor"),
    std::wstring_view(L"Reserved"),
};

[[nodiscard]] COLORREF ColorRefFromArgb(uint32_t argb) noexcept
{
    const BYTE r = static_cast<BYTE>((argb >> 16) & 0xFFu);
    const BYTE g = static_cast<BYTE>((argb >> 8) & 0xFFu);
    const BYTE b = static_cast<BYTE>(argb & 0xFFu);
    return RGB(r, g, b);
}

[[nodiscard]] COLORREF BlendColor(COLORREF under, COLORREF over, uint8_t alpha) noexcept
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

[[nodiscard]] COLORREF ContrastingTextColor(COLORREF background) noexcept
{
    const uint32_t r    = static_cast<uint32_t>(GetRValue(background));
    const uint32_t g    = static_cast<uint32_t>(GetGValue(background));
    const uint32_t b    = static_cast<uint32_t>(GetBValue(background));
    const uint32_t luma = (r * 299u + g * 587u + b * 114u) / 1000u;
    return luma < 128u ? RGB(255, 255, 255) : RGB(0, 0, 0);
}

[[nodiscard]] uint32_t StableHash32(std::wstring_view text) noexcept
{
    uint32_t hash = 2166136261u;
    for (wchar_t ch : text)
    {
        hash ^= static_cast<uint32_t>(ch);
        hash *= 16777619u;
    }
    return hash;
}

[[nodiscard]] COLORREF ColorFromHSV(float hueDegrees, float saturation, float value) noexcept
{
    const float h = std::fmod(std::max(0.0f, hueDegrees), 360.0f);
    const float s = std::clamp(saturation, 0.0f, 1.0f);
    const float v = std::clamp(value, 0.0f, 1.0f);

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

[[nodiscard]] COLORREF ResolveAccentColor(const ViewerTheme& theme, std::wstring_view seed) noexcept
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

[[nodiscard]] D2D1_COLOR_F ColorFFromColorRef(COLORREF color, float alpha = 1.0f) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;
    return D2D1::ColorF(r, g, b, alpha);
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int needed = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (needed <= 0)
    {
        return {};
    }

    std::wstring out;
    out.resize(static_cast<size_t>(needed));

    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), out.data(), needed);
    if (written != needed)
    {
        return {};
    }

    return out;
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int needed = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (needed <= 0)
    {
        return {};
    }

    std::string out;
    out.resize(static_cast<size_t>(needed));

    const int written = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), out.data(), needed, nullptr, nullptr);
    if (written != needed)
    {
        return {};
    }

    return out;
}

[[nodiscard]] int PxFromDip(float dip, UINT dpi) noexcept
{
    const float px = dip * static_cast<float>(dpi) / 96.0f;
    return static_cast<int>(std::lround(px));
}

enum class ExportFormat
{
    Text,
    Markdown,
};

[[nodiscard]] std::optional<std::filesystem::path> ShowExportSaveDialog(HWND owner, const std::wstring& defaultFileName, ExportFormat format) noexcept
{
    if (! owner)
    {
        return std::nullopt;
    }

    [[maybe_unused]] auto coInit = wil::CoInitializeEx();

    wil::com_ptr<IFileSaveDialog> dialog;
    const HRESULT hr = CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(hr) || ! dialog)
    {
        return std::nullopt;
    }

    DWORD options = 0;
    static_cast<void>(dialog->GetOptions(&options));
    options |= FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_OVERWRITEPROMPT;
    static_cast<void>(dialog->SetOptions(options));

    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERPE_EXPORT_DIALOG_TITLE);
    if (! title.empty())
    {
        static_cast<void>(dialog->SetTitle(title.c_str()));
    }

    if (! defaultFileName.empty())
    {
        static_cast<void>(dialog->SetFileName(defaultFileName.c_str()));
    }

    COMDLG_FILTERSPEC spec{};
    if (format == ExportFormat::Markdown)
    {
        const std::wstring name = LoadStringResource(g_hInstance, IDS_VIEWERPE_EXPORT_FILTER_MARKDOWN);
        spec.pszName            = name.c_str();
        spec.pszSpec            = L"*.md";
        static_cast<void>(dialog->SetDefaultExtension(L"md"));
        static_cast<void>(dialog->SetFileTypes(1, &spec));
    }
    else
    {
        const std::wstring name = LoadStringResource(g_hInstance, IDS_VIEWERPE_EXPORT_FILTER_TEXT);
        spec.pszName            = name.c_str();
        spec.pszSpec            = L"*.txt";
        static_cast<void>(dialog->SetDefaultExtension(L"txt"));
        static_cast<void>(dialog->SetFileTypes(1, &spec));
    }

    const HRESULT showHr = dialog->Show(owner);
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

[[nodiscard]] HRESULT WriteUtf8FileWithBom(const std::filesystem::path& path, std::wstring_view content) noexcept
{
    const std::string utf8 = Utf8FromUtf16(content);
    if (utf8.empty() && ! content.empty())
    {
        return E_FAIL;
    }

    constexpr size_t kBomSize = 3;
    if (utf8.size() > static_cast<size_t>((std::numeric_limits<DWORD>::max)()) - kBomSize)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }

    wil::unique_handle outFile(CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! outFile)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    DWORD written                        = 0;
    static constexpr BYTE kBom[kBomSize] = {0xEF, 0xBB, 0xBF};
    if (WriteFile(outFile.get(), kBom, static_cast<DWORD>(kBomSize), &written, nullptr) == 0 || written != kBomSize)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if (! utf8.empty())
    {
        const DWORD want = static_cast<DWORD>(utf8.size());
        written          = 0;
        if (WriteFile(outFile.get(), utf8.data(), want, &written, nullptr) == 0 || written != want)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }

    return S_OK;
}

[[nodiscard]] float ClampScroll(float scrollDip, float contentHeightDip, float viewportHeightDip) noexcept
{
    const float maxScroll = std::max(0.0f, contentHeightDip - viewportHeightDip);
    return std::clamp(scrollDip, 0.0f, maxScroll);
}

struct ParsedPeDeleter
{
    void operator()(peparse::parsed_pe* pe) const noexcept
    {
        if (pe)
        {
            peparse::DestructParsedPE(pe);
        }
    }
};

[[nodiscard]] std::wstring MachineText(peparse::parsed_pe* pe) noexcept
{
    if (! pe)
    {
        return {};
    }
    if (const char* s = peparse::GetMachineAsString(pe); s && s[0] != '\0')
    {
        return Utf16FromUtf8(s);
    }
    return {};
}

[[nodiscard]] std::wstring SubsystemText(peparse::parsed_pe* pe) noexcept
{
    if (! pe)
    {
        return {};
    }
    if (const char* s = peparse::GetSubsystemAsString(pe); s && s[0] != '\0')
    {
        return Utf16FromUtf8(s);
    }
    return {};
}

[[nodiscard]] std::wstring_view PeKindName(std::uint16_t optionalMagic) noexcept
{
    switch (optionalMagic)
    {
        case 0x10Bu: return L"PE32";
        case 0x20Bu: return L"PE32+";
        default: return L"PE";
    }
}
} // namespace

ViewerPE::ViewerPE()
{
    _metaId          = L"builtin/viewer-pe";
    _metaShortId     = L"pe";
    _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERPE_NAME);
    _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERPE_DESCRIPTION);

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = nullptr;

    _configurationJson = "{}";
}

ViewerPE::~ViewerPE() = default;

void ViewerPE::SetHost(IHost* host) noexcept
{
    _hostAlerts = nullptr;

    if (! host)
    {
        return;
    }

    wil::com_ptr<IHostAlerts> alerts;
    const HRESULT hr = host->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
    if (SUCCEEDED(hr) && alerts)
    {
        _hostAlerts = std::move(alerts);
    }
}

HRESULT STDMETHODCALLTYPE ViewerPE::QueryInterface(REFIID riid, void** ppvObject) noexcept
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

ULONG STDMETHODCALLTYPE ViewerPE::AddRef() noexcept
{
    return _refCount.fetch_add(1) + 1;
}

ULONG STDMETHODCALLTYPE ViewerPE::Release() noexcept
{
    const ULONG remaining = _refCount.fetch_sub(1) - 1;
    if (remaining == 0)
    {
        delete this;
    }
    return remaining;
}

HRESULT STDMETHODCALLTYPE ViewerPE::GetMetaData(const PluginMetaData** metaData) noexcept
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

HRESULT STDMETHODCALLTYPE ViewerPE::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = kViewerPESchemaJson;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr || configurationJsonUtf8[0] == '\0')
    {
        _configurationJson = "{}";
        return S_OK;
    }

    _configurationJson = configurationJsonUtf8;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    *pSomethingToSave = FALSE;
    return S_OK;
}

ATOM ViewerPE::RegisterWndClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom != 0)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.hInstance     = instance;
    wc.lpszClassName = kClassName;
    wc.lpfnWndProc   = WndProcThunk;
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);

    atom = RegisterClassExW(&wc);
    return atom;
}

LRESULT CALLBACK ViewerPE::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        if (cs && cs->lpCreateParams)
        {
            auto* self = static_cast<ViewerPE*>(cs->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            InitPostedPayloadWindow(hwnd);
        }
    }

    auto* self = reinterpret_cast<ViewerPE*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerPE::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_DPICHANGED: OnDpiChanged(static_cast<UINT>(LOWORD(wp)), *reinterpret_cast<const RECT*>(lp)); return 0;
        case WM_COMMAND: OnCommand(hwnd, LOWORD(wp), HIWORD(wp), reinterpret_cast<HWND>(lp)); return 0;
        case WM_MEASUREITEM: return OnMeasureItem(reinterpret_cast<MEASUREITEMSTRUCT*>(lp));
        case WM_DRAWITEM: return OnDrawItem(reinterpret_cast<DRAWITEMSTRUCT*>(lp));
        case WM_PAINT: OnPaint(hwnd); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_MOUSEWHEEL: OnMouseWheel(GET_WHEEL_DELTA_WPARAM(wp)); return 0;
        case WM_VSCROLL: OnVScroll(LOWORD(wp), HIWORD(wp)); return 0;
        case WM_KEYDOWN: OnKeyDown(static_cast<UINT>(wp)); return 0;
        case WM_CTLCOLORLISTBOX:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC:
        case WM_CTLCOLORBTN: return OnCtlColor(msg, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_NCACTIVATE: ApplyTitleBarTheme(wp != FALSE); return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_NCDESTROY:
        {
            static_cast<void>(DrainPostedPayloadsForWindow(hwnd));
            ResetDeviceResources();
            _hFileCombo.release();
            _hFileComboList = nullptr;
            _hFileComboItem = nullptr;
            _hWnd.release();
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);

            if (_callback)
            {
                static_cast<void>(_callback->ViewerClosed(_callbackCookie));
            }

            Release();
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case kAsyncParseCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncParseResult>(lp);
            OnAsyncParseComplete(std::move(result));
            return 0;
        }
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }
}

void ViewerPE::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);

    const int uiHeightPx = -MulDiv(9, static_cast<int>(_dpi), 72);
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
        Debug::ErrorWithLastError(L"ViewerPE: CreateFontW failed for UI font.");
    }

    const DWORD comboStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS;
    _hFileCombo.reset(CreateWindowExW(
        0, L"COMBOBOX", nullptr, comboStyle, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_VIEWERPE_FILE_COMBO)), g_hInstance, nullptr));
    if (! _hFileCombo)
    {
        Debug::ErrorWithLastError(L"ViewerPE: CreateWindowExW failed for file combo.");
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
        int itemHeight = PxFromDip(24, _dpi);
        auto hdc       = wil::GetDC(hwnd);
        if (hdc)
        {
            HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
            static_cast<void>(oldFont);

            TEXTMETRICW tm{};
            if (GetTextMetricsW(hdc.get(), &tm) != 0)
            {
                itemHeight = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, _dpi);
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

    ApplyTitleBarTheme(true);
    RefreshFileCombo(hwnd);
    Layout(hwnd);
}

void ViewerPE::OnSize(UINT /*width*/, UINT /*height*/) noexcept
{
    if (_renderTarget)
    {
        RECT rc{};
        if (GetClientRect(_hWnd.get(), &rc) != 0)
        {
            const D2D1_SIZE_U size = D2D1::SizeU(static_cast<UINT>(std::max(1l, rc.right - rc.left)), static_cast<UINT>(std::max(1l, rc.bottom - rc.top)));
            _renderTarget->Resize(size);
        }
    }

    if (_hWnd)
    {
        Layout(_hWnd.get());
    }

    _textLayout.reset();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void ViewerPE::OnDpiChanged(UINT dpi, const RECT& suggestedRect) noexcept
{
    _dpi = dpi;

    if (_hWnd)
    {
        SetWindowPos(_hWnd.get(),
                     nullptr,
                     suggestedRect.left,
                     suggestedRect.top,
                     suggestedRect.right - suggestedRect.left,
                     suggestedRect.bottom - suggestedRect.top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

    const int uiHeightPx = -MulDiv(9, static_cast<int>(_dpi), 72);
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
    if (_hFileCombo && _uiFont)
    {
        SendMessageW(_hFileCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_uiFont.get()), TRUE);

        int itemHeight = PxFromDip(24, _dpi);
        auto hdc       = wil::GetDC(_hWnd.get());
        if (hdc)
        {
            HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
            static_cast<void>(oldFont);

            TEXTMETRICW tm{};
            if (GetTextMetricsW(hdc.get(), &tm) != 0)
            {
                itemHeight = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, _dpi);
            }
        }

        itemHeight = std::max(itemHeight, 1);
        SendMessageW(_hFileCombo.get(), CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), static_cast<LPARAM>(itemHeight));
        SendMessageW(_hFileCombo.get(), CB_SETITEMHEIGHT, 0, static_cast<LPARAM>(itemHeight));
    }

    _textLayout.reset();
    if (_renderTarget)
    {
        _renderTarget->SetDpi(static_cast<float>(dpi), static_cast<float>(dpi));
    }
    _headerBrush.reset();

    if (_hWnd)
    {
        Layout(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void ViewerPE::OnPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);

    EnsureDeviceResources(hwnd);
    if (! _renderTarget || ! _textBrush || ! _bgBrush || ! _cardBrush || ! _cardBorderBrush)
    {
        return;
    }

    const D2D1_SIZE_F size = _renderTarget->GetSize();

    const float cardLeft   = kOuterPaddingDip;
    const float cardTop    = kOuterPaddingDip;
    const float cardRight  = std::max(cardLeft + 1.0f, size.width - kOuterPaddingDip);
    const float cardBottom = std::max(cardTop + 1.0f, size.height - kOuterPaddingDip);
    const D2D1_ROUNDED_RECT card{D2D1::RectF(cardLeft, cardTop, cardRight, cardBottom), kCardRadiusDip, kCardRadiusDip};

    const float contentLeft       = cardLeft + kInnerPaddingDip;
    const float contentTop        = cardTop + kInnerPaddingDip + std::max(0.0f, _headerHeightDip);
    const float contentRight      = std::max(contentLeft + 1.0f, cardRight - kInnerPaddingDip);
    const float contentBottom     = std::max(contentTop + 1.0f, cardBottom - kInnerPaddingDip);
    const float contentWidthDip   = std::max(1.0f, contentRight - contentLeft);
    const float viewportHeightDip = std::max(1.0f, contentBottom - contentTop);

    EnsureTextLayout(contentWidthDip, viewportHeightDip);
    UpdateScrollBars(hwnd, viewportHeightDip);

    const COLORREF clearColor = (_hasTheme && ! _theme.highContrast) ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);

    _renderTarget->BeginDraw();
    _renderTarget->Clear(ColorFFromColorRef(clearColor));

    _renderTarget->FillRoundedRectangle(card, _cardBrush.get());
    _renderTarget->DrawRoundedRectangle(card, _cardBorderBrush.get(), 1.0f);

    const D2D1_RECT_F clipRc = D2D1::RectF(contentLeft, contentTop, contentRight, contentBottom);
    _renderTarget->PushAxisAlignedClip(clipRc, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

    if (_textLayout)
    {
        const D2D1_POINT_2F origin = D2D1::Point2F(contentLeft, contentTop - _scrollDip);
        _renderTarget->DrawTextLayout(origin, _textLayout.get(), _textBrush.get());
    }

    _renderTarget->PopAxisAlignedClip();

    const HRESULT hr = _renderTarget->EndDraw();
    if (hr == D2DERR_RECREATE_TARGET)
    {
        _textBrush.reset();
        _cardBorderBrush.reset();
        _cardBrush.reset();
        _bgBrush.reset();
        _renderTarget.reset();
    }
}

void ViewerPE::OnMouseWheel(short delta) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const float steps = static_cast<float>(delta) / static_cast<float>(WHEEL_DELTA);
    ScrollByDip(_hWnd.get(), -steps * kScrollWheelStepDip);
}

void ViewerPE::OnVScroll(WORD request, WORD /*position*/) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    if (GetScrollInfo(_hWnd.get(), SB_VERT, &si) == 0)
    {
        return;
    }

    int newPos = si.nPos;
    switch (request)
    {
        case SB_TOP: newPos = si.nMin; break;
        case SB_BOTTOM: newPos = si.nMax; break;
        case SB_LINEUP: newPos = si.nPos - static_cast<int>(MulDiv(static_cast<int>(std::lround(kScrollLineStepDip)), static_cast<int>(_dpi), 96)); break;
        case SB_LINEDOWN: newPos = si.nPos + static_cast<int>(MulDiv(static_cast<int>(std::lround(kScrollLineStepDip)), static_cast<int>(_dpi), 96)); break;
        case SB_PAGEUP: newPos = si.nPos - static_cast<int>(si.nPage); break;
        case SB_PAGEDOWN: newPos = si.nPos + static_cast<int>(si.nPage); break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: newPos = si.nTrackPos; break;
        default: return;
    }

    const float scrollDip = static_cast<float>(newPos) * 96.0f / static_cast<float>(_dpi);
    SetScrollDip(_hWnd.get(), scrollDip);
}

void ViewerPE::OnKeyDown(UINT vk) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

    if (vk == VK_ESCAPE)
    {
        CommandExit();
        return;
    }

    if (vk == VK_F5)
    {
        CommandRefresh(_hWnd.get());
        return;
    }

    if (vk == VK_SPACE)
    {
        CommandOtherNext(_hWnd.get());
        return;
    }

    if (vk == VK_BACK)
    {
        CommandOtherPrevious(_hWnd.get());
        return;
    }

    if (ctrl && vk == VK_UP)
    {
        CommandOtherPrevious(_hWnd.get());
        return;
    }

    if (ctrl && vk == VK_DOWN)
    {
        CommandOtherNext(_hWnd.get());
        return;
    }

    if (ctrl && vk == VK_HOME)
    {
        CommandOtherFirst(_hWnd.get());
        return;
    }

    if (ctrl && vk == VK_END)
    {
        CommandOtherLast(_hWnd.get());
        return;
    }

    if (ctrl && (vk == 'S' || vk == 's'))
    {
        const UINT cmd = shift ? IDM_VIEWERPE_FILE_EXPORT_MARKDOWN : IDM_VIEWERPE_FILE_EXPORT_TEXT;
        SendMessageW(_hWnd.get(), WM_COMMAND, static_cast<WPARAM>(cmd), 0);
        return;
    }

    switch (vk)
    {
        case VK_UP: ScrollByDip(_hWnd.get(), -kScrollLineStepDip); break;
        case VK_DOWN: ScrollByDip(_hWnd.get(), kScrollLineStepDip); break;
        case VK_PRIOR: ScrollByDip(_hWnd.get(), -std::max(1.0f, _viewportHeightDip * 0.9f)); break;
        case VK_NEXT: ScrollByDip(_hWnd.get(), std::max(1.0f, _viewportHeightDip * 0.9f)); break;
        case VK_HOME: SetScrollDip(_hWnd.get(), 0.0f); break;
        case VK_END: SetScrollDip(_hWnd.get(), std::numeric_limits<float>::infinity()); break;
        default: break;
    }
}

void ViewerPE::OnCommand(HWND hwnd, UINT commandId, UINT notifyCode, HWND control) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (_hFileCombo && control == _hFileCombo.get() && commandId == IDC_VIEWERPE_FILE_COMBO)
    {
        if (notifyCode == CBN_DROPDOWN)
        {
            COMBOBOXINFO info{};
            info.cbSize = sizeof(info);
            if (GetComboBoxInfo(_hFileCombo.get(), &info) != 0)
            {
                _hFileComboList = info.hwndList;
                _hFileComboItem = info.hwndItem;
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

            SetWindowTheme(_hFileCombo.get(), winTheme, nullptr);
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
            SendMessageW(_hFileCombo.get(), WM_THEMECHANGED, 0, 0);
            return;
        }

        if (notifyCode == CBN_SELCHANGE && ! _syncingFileCombo)
        {
            const LRESULT sel = SendMessageW(_hFileCombo.get(), CB_GETCURSEL, 0, 0);
            if (sel >= 0 && static_cast<size_t>(sel) < _otherFiles.size())
            {
                _otherIndex  = static_cast<size_t>(sel);
                _currentPath = _otherFiles[_otherIndex];
                StartAsyncParse(hwnd, _fileSystem, _currentPath);
                UpdateMenuState(hwnd);
                SetFocus(hwnd);
            }
        }

        return;
    }

    switch (commandId)
    {
        case IDM_VIEWERPE_FILE_EXPORT_TEXT: CommandExportText(hwnd); break;
        case IDM_VIEWERPE_FILE_EXPORT_MARKDOWN: CommandExportMarkdown(hwnd); break;
        case IDM_VIEWERPE_FILE_REFRESH: CommandRefresh(hwnd); break;
        case IDM_VIEWERPE_OTHER_NEXT: CommandOtherNext(hwnd); break;
        case IDM_VIEWERPE_OTHER_PREVIOUS: CommandOtherPrevious(hwnd); break;
        case IDM_VIEWERPE_OTHER_FIRST: CommandOtherFirst(hwnd); break;
        case IDM_VIEWERPE_OTHER_LAST: CommandOtherLast(hwnd); break;
        case IDM_VIEWERPE_VIEW_GOTO_TOP: SetScrollDip(hwnd, 0.0f); break;
        case IDM_VIEWERPE_VIEW_GOTO_BOTTOM: SetScrollDip(hwnd, std::numeric_limits<float>::infinity()); break;
        case IDM_VIEWERPE_FILE_EXIT: CommandExit(); break;
        default: break;
    }
}

LRESULT ViewerPE::OnMeasureItem(MEASUREITEMSTRUCT* measure) noexcept
{
    if (! measure)
    {
        return FALSE;
    }

    if (measure->CtlType == ODT_COMBOBOX && measure->CtlID == IDC_VIEWERPE_FILE_COMBO)
    {
        const UINT dpi = _hWnd ? GetDpiForWindow(_hWnd.get()) : USER_DEFAULT_SCREEN_DPI;

        int height = PxFromDip(24, dpi);
        auto hdc   = wil::GetDC(_hWnd.get());
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

LRESULT ViewerPE::OnDrawItem(DRAWITEMSTRUCT* draw) noexcept
{
    if (! draw)
    {
        return FALSE;
    }

    if (draw->CtlType != ODT_COMBOBOX || ! _hFileCombo || draw->hwndItem != _hFileCombo.get())
    {
        return FALSE;
    }

    if (! draw->hDC)
    {
        return TRUE;
    }

    const UINT dpi    = _hWnd ? GetDpiForWindow(_hWnd.get()) : USER_DEFAULT_SCREEN_DPI;
    const int padding = PxFromDip(6, dpi);

    const bool selected = (draw->itemState & ODS_SELECTED) != 0;
    const bool disabled = (draw->itemState & ODS_DISABLED) != 0;

    const bool themed = _hasTheme && ! _theme.highContrast;
    const COLORREF bg = themed ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg = themed ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
    COLORREF baseBg   = themed ? BlendColor(bg, fg, themed && _theme.darkMode ? 24u : 18u) : GetSysColor(COLOR_WINDOW);
    COLORREF baseFg   = fg;
    COLORREF selBg    = themed ? ResolveAccentColor(_theme, L"combo") : GetSysColor(COLOR_HIGHLIGHT);
    COLORREF selFg    = themed ? ContrastingTextColor(selBg) : GetSysColor(COLOR_HIGHLIGHTTEXT);

    if (_hasTheme && _theme.highContrast)
    {
        baseBg = GetSysColor(COLOR_WINDOW);
        baseFg = GetSysColor(COLOR_WINDOWTEXT);
        selBg  = GetSysColor(COLOR_HIGHLIGHT);
        selFg  = GetSysColor(COLOR_HIGHLIGHTTEXT);
    }

    COLORREF fillColor = selected ? selBg : baseBg;
    COLORREF textColor = selected ? selFg : baseFg;

    if (disabled)
    {
        textColor = BlendColor(fillColor, textColor, 120u);
    }

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> bgBrush(CreateSolidBrush(fillColor));
    FillRect(draw->hDC, &draw->rcItem, bgBrush.get());

    int itemId = static_cast<int>(draw->itemID);
    if (itemId < 0)
    {
        const LRESULT sel = SendMessageW(_hFileCombo.get(), CB_GETCURSEL, 0, 0);
        if (sel >= 0)
        {
            itemId = static_cast<int>(sel);
        }
    }

    std::wstring text;
    if (itemId >= 0)
    {
        const LRESULT lenRes = SendMessageW(_hFileCombo.get(), CB_GETLBTEXTLEN, static_cast<WPARAM>(itemId), 0);
        const int len        = (lenRes > 0) ? static_cast<int>(lenRes) : 0;
        if (len > 0)
        {
            text.resize(static_cast<size_t>(len) + 1);
            SendMessageW(_hFileCombo.get(), CB_GETLBTEXT, static_cast<WPARAM>(itemId), reinterpret_cast<LPARAM>(text.data()));
            text.resize(wcsnlen(text.c_str(), text.size()));
        }
    }

    HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(draw->hDC, fontToUse);
    static_cast<void>(oldFont);

    SetBkMode(draw->hDC, TRANSPARENT);
    SetTextColor(draw->hDC, textColor);

    RECT textRc = draw->rcItem;
    textRc.left += padding;
    textRc.right -= padding;
    DrawTextW(draw->hDC, text.c_str(), static_cast<int>(text.size()), &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    if ((draw->itemState & ODS_FOCUS) != 0)
    {
        DrawFocusRect(draw->hDC, &draw->rcItem);
    }

    return TRUE;
}

LRESULT ViewerPE::OnCtlColor(UINT /*msg*/, HDC hdc, HWND control) noexcept
{
    if (! hdc || ! control || ! _hasTheme || _theme.highContrast)
    {
        return 0;
    }

    if (_hFileCombo && (control == _hFileCombo.get() || (_hFileComboList != nullptr && control == _hFileComboList) ||
                        (_hFileComboItem != nullptr && control == _hFileComboItem)))
    {
        const COLORREF bg = BlendColor(ColorRefFromArgb(_theme.backgroundArgb), ColorRefFromArgb(_theme.textArgb), _theme.darkMode ? 24u : 18u);
        if (! _headerBrush)
        {
            _headerBrush.reset(CreateSolidBrush(bg));
        }
        if (! _headerBrush)
        {
            return 0;
        }

        SetBkMode(hdc, OPAQUE);
        SetTextColor(hdc, ColorRefFromArgb(_theme.textArgb));
        SetBkColor(hdc, bg);
        return reinterpret_cast<LRESULT>(_headerBrush.get());
    }

    return 0;
}

void ViewerPE::Layout(HWND hwnd) noexcept
{
    if (! hwnd || ! _hFileCombo)
    {
        _headerHeightDip = 0.0f;
        return;
    }

    RECT client{};
    GetClientRect(hwnd, &client);

    const UINT dpi           = GetDpiForWindow(hwnd);
    const bool showCombo     = (_otherFiles.size() > 1);
    const int outerPaddingPx = PxFromDip(kOuterPaddingDip, dpi);
    const int innerPaddingPx = PxFromDip(kInnerPaddingDip, dpi);

    const int cardLeft  = outerPaddingPx;
    const int cardTop   = outerPaddingPx;
    const int cardRight = std::max(cardLeft + 1, static_cast<int>(client.right) - outerPaddingPx);

    const int contentLeft  = cardLeft + innerPaddingPx;
    const int contentRight = std::max(contentLeft + 1, cardRight - innerPaddingPx);

    ShowWindow(_hFileCombo.get(), showCombo ? SW_SHOW : SW_HIDE);
    EnableWindow(_hFileCombo.get(), showCombo ? TRUE : FALSE);

    float newHeaderHeightDip = 0.0f;
    if (showCombo)
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
        }

        const int edgeSizeY     = GetSystemMetricsForDpi(SM_CYEDGE, dpi);
        const int comboBorder   = std::max(0, edgeSizeY) * 2;
        const int chromePadding = std::max(PxFromDip(4, dpi), comboBorder);
        const int comboHeight   = std::max(1, comboItemHeight + chromePadding);
        const int comboX        = contentLeft;
        const int comboW        = std::max(1, contentRight - contentLeft);
        const int comboY        = cardTop + innerPaddingPx;

        SetWindowPos(_hFileCombo.get(), nullptr, comboX, comboY, comboW, comboHeight, SWP_NOZORDER | SWP_NOACTIVATE);

        newHeaderHeightDip = static_cast<float>(comboHeight) * 96.0f / static_cast<float>(dpi) + kHeaderGapDip;
    }

    if (std::fabs(newHeaderHeightDip - _headerHeightDip) > 0.25f)
    {
        _headerHeightDip = newHeaderHeightDip;
        _textLayout.reset();
    }
}

void ViewerPE::RefreshFileCombo(HWND hwnd) noexcept
{
    if (! _hFileCombo)
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    SendMessageW(_hFileCombo.get(), CB_RESETCONTENT, 0, 0);

    if (_otherFiles.size() <= 1)
    {
        SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(-1), 0);
        if (hwnd)
        {
            Layout(hwnd);
            InvalidateRect(hwnd, nullptr, TRUE);
        }
        return;
    }

    for (const auto& path : _otherFiles)
    {
        std::wstring itemText;
        itemText = std::filesystem::path(path).filename().wstring();
        SendMessageW(_hFileCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(itemText.c_str()));
    }

    if (_otherIndex >= _otherFiles.size())
    {
        _otherIndex = 0;
    }

    SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(_otherIndex), 0);
    SendMessageW(_hFileCombo.get(), CB_SETMINVISIBLE, static_cast<WPARAM>(std::min<size_t>(_otherFiles.size(), 15)), 0);

    if (hwnd)
    {
        Layout(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);
    }
}

void ViewerPE::SyncFileComboSelection() noexcept
{
    if (! _hFileCombo)
    {
        return;
    }

    if (_otherFiles.size() <= 1)
    {
        return;
    }

    if (_otherIndex >= _otherFiles.size())
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(_otherIndex), 0);
}

void ViewerPE::UpdateMenuState([[maybe_unused]] HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    HMENU menu = GetMenu(hwnd);
    if (! menu)
    {
        return;
    }

    const bool hasOther   = _otherFiles.size() > 1;
    const UINT otherState = static_cast<UINT>(MF_BYCOMMAND | (hasOther ? MF_ENABLED : MF_GRAYED));

    EnableMenuItem(menu, IDM_VIEWERPE_OTHER_NEXT, otherState);
    EnableMenuItem(menu, IDM_VIEWERPE_OTHER_PREVIOUS, otherState);
    EnableMenuItem(menu, IDM_VIEWERPE_OTHER_FIRST, otherState);
    EnableMenuItem(menu, IDM_VIEWERPE_OTHER_LAST, otherState);

    const bool canRefresh = ! _currentPath.empty();
    EnableMenuItem(menu, IDM_VIEWERPE_FILE_REFRESH, static_cast<UINT>(MF_BYCOMMAND | (canRefresh ? MF_ENABLED : MF_GRAYED)));

    const bool canExport = ! _isLoading && (! _subtitleText.empty() || ! _bodyText.empty());
    EnableMenuItem(menu, IDM_VIEWERPE_FILE_EXPORT_TEXT, static_cast<UINT>(MF_BYCOMMAND | (canExport ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, IDM_VIEWERPE_FILE_EXPORT_MARKDOWN, static_cast<UINT>(MF_BYCOMMAND | (canExport ? MF_ENABLED : MF_GRAYED)));

    DrawMenuBar(hwnd);
}

void ViewerPE::CommandExit() noexcept
{
    static_cast<void>(Close());
}

void ViewerPE::CommandRefresh(HWND hwnd) noexcept
{
    if (! hwnd || _currentPath.empty())
    {
        return;
    }

    StartAsyncParse(hwnd, _fileSystem, _currentPath);
    UpdateMenuState(hwnd);
    SetFocus(hwnd);
}

void ViewerPE::CommandOtherNext(HWND hwnd) noexcept
{
    if (! hwnd || _otherFiles.size() <= 1)
    {
        return;
    }

    _otherIndex  = (_otherIndex + 1) % _otherFiles.size();
    _currentPath = _otherFiles[_otherIndex];

    SyncFileComboSelection();
    StartAsyncParse(hwnd, _fileSystem, _currentPath);
    UpdateMenuState(hwnd);
    SetFocus(hwnd);
}

void ViewerPE::CommandOtherPrevious(HWND hwnd) noexcept
{
    if (! hwnd || _otherFiles.size() <= 1)
    {
        return;
    }

    if (_otherIndex == 0)
    {
        _otherIndex = _otherFiles.size() - 1;
    }
    else
    {
        _otherIndex -= 1;
    }

    _currentPath = _otherFiles[_otherIndex];

    SyncFileComboSelection();
    StartAsyncParse(hwnd, _fileSystem, _currentPath);
    UpdateMenuState(hwnd);
    SetFocus(hwnd);
}

void ViewerPE::CommandOtherFirst(HWND hwnd) noexcept
{
    if (! hwnd || _otherFiles.empty())
    {
        return;
    }

    _otherIndex  = 0;
    _currentPath = _otherFiles[_otherIndex];

    SyncFileComboSelection();
    StartAsyncParse(hwnd, _fileSystem, _currentPath);
    UpdateMenuState(hwnd);
    SetFocus(hwnd);
}

void ViewerPE::CommandOtherLast(HWND hwnd) noexcept
{
    if (! hwnd || _otherFiles.empty())
    {
        return;
    }

    _otherIndex  = _otherFiles.size() - 1;
    _currentPath = _otherFiles[_otherIndex];

    SyncFileComboSelection();
    StartAsyncParse(hwnd, _fileSystem, _currentPath);
    UpdateMenuState(hwnd);
    SetFocus(hwnd);
}

void ViewerPE::CommandExportText(HWND hwnd) noexcept
{
    if (! hwnd || _isLoading)
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_NO_REPORT, IDS_VIEWERPE_NAME, MB_OK | MB_ICONINFORMATION);
        return;
    }

    std::wstring content;
    content = std::format(L"{}\n{}\n\n{}", _titleText, _subtitleText, _bodyText);
    if (content.empty())
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_NO_REPORT, IDS_VIEWERPE_NAME, MB_OK | MB_ICONINFORMATION);
        return;
    }

    std::wstring defaultName;
    defaultName     = _titleText.empty() ? L"pe-report.pe.txt" : (_titleText + L".pe.txt");
    const auto dest = ShowExportSaveDialog(hwnd, defaultName, ExportFormat::Text);
    if (! dest.has_value())
    {
        return;
    }

    const HRESULT hr = WriteUtf8FileWithBom(dest.value(), content);
    if (FAILED(hr))
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_FAILED, IDS_VIEWERPE_NAME, MB_OK | MB_ICONERROR);
    }
}

void ViewerPE::CommandExportMarkdown(HWND hwnd) noexcept
{
    if (! hwnd || _isLoading)
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_NO_REPORT, IDS_VIEWERPE_NAME, MB_OK | MB_ICONINFORMATION);
        return;
    }

    std::wstring content;
    content = _markdownText;
    if (content.empty())
    {
        const std::wstring mdTitle = _titleText.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERPE_NAME) : _titleText;
        content                    = std::format(L"# {}\n\n{}\n\n```text\n{}\n```\n", mdTitle, _subtitleText, _bodyText);
    }

    if (content.empty())
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_NO_REPORT, IDS_VIEWERPE_NAME, MB_OK | MB_ICONINFORMATION);
        return;
    }

    std::wstring defaultName;
    defaultName     = _titleText.empty() ? L"pe-report.pe.md" : (_titleText + L".pe.md");
    const auto dest = ShowExportSaveDialog(hwnd, defaultName, ExportFormat::Markdown);
    if (! dest.has_value())
    {
        return;
    }

    const HRESULT hr = WriteUtf8FileWithBom(dest.value(), content);
    if (FAILED(hr))
    {
        MessageBoxResource(hwnd, g_hInstance, IDS_VIEWERPE_EXPORT_ERROR_FAILED, IDS_VIEWERPE_NAME, MB_OK | MB_ICONERROR);
    }
}

void ViewerPE::ApplyTitleBarTheme(bool windowActive) noexcept
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
            static constexpr uint8_t kInactiveTitleBlendAlpha = 223u;
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

void ViewerPE::ResetDeviceResources() noexcept
{
    _worker = std::jthread();
    _textLayout.reset();
    _baseTextFormat.reset();
    _textBrush.reset();
    _cardBorderBrush.reset();
    _cardBrush.reset();
    _bgBrush.reset();
    _renderTarget.reset();
    _writeFactory.reset();
    _d2dFactory.reset();
}

void ViewerPE::EnsureDeviceResources(HWND /*hwnd*/) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const HWND hwnd = _hWnd.get();

    if (! _d2dFactory)
    {
        D2D1_FACTORY_OPTIONS options{};
        options.debugLevel = D2D1_DEBUG_LEVEL_NONE;
        const HRESULT hr   = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, options, _d2dFactory.put());
        if (FAILED(hr))
        {
            return;
        }
    }

    if (! _writeFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_writeFactory.put()));
        if (FAILED(hr))
        {
            return;
        }
    }

    if (! _renderTarget)
    {
        RECT rc{};
        GetClientRect(hwnd, &rc);
        const D2D1_SIZE_U size = D2D1::SizeU(static_cast<UINT>(std::max(1l, rc.right - rc.left)), static_cast<UINT>(std::max(1l, rc.bottom - rc.top)));

        D2D1_RENDER_TARGET_PROPERTIES rtProps        = D2D1::RenderTargetProperties(D2D1_RENDER_TARGET_TYPE_DEFAULT,
                                                                             D2D1::PixelFormat(DXGI_FORMAT_UNKNOWN, D2D1_ALPHA_MODE_IGNORE),
                                                                             static_cast<float>(_dpi),
                                                                             static_cast<float>(_dpi));
        D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, size);

        const HRESULT hr = _d2dFactory->CreateHwndRenderTarget(rtProps, hwndProps, _renderTarget.put());
        if (FAILED(hr))
        {
            return;
        }
    }

    const bool themed         = _hasTheme && ! _theme.highContrast;
    const COLORREF bg         = themed ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg         = themed ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
    const COLORREF accent     = themed ? ResolveAccentColor(_theme, L"content") : GetSysColor(COLOR_HIGHLIGHT);
    const COLORREF cardBg     = themed ? BlendColor(bg, fg, themed && _theme.darkMode ? 24u : 18u) : GetSysColor(COLOR_WINDOW);
    const COLORREF cardBorder = themed ? BlendColor(cardBg, accent, 92u) : GetSysColor(COLOR_WINDOWFRAME);

    if (! _bgBrush)
    {
        static_cast<void>(_renderTarget->CreateSolidColorBrush(ColorFFromColorRef(bg), _bgBrush.put()));
    }
    if (! _cardBrush)
    {
        static_cast<void>(_renderTarget->CreateSolidColorBrush(ColorFFromColorRef(cardBg), _cardBrush.put()));
    }
    if (! _cardBorderBrush)
    {
        static_cast<void>(_renderTarget->CreateSolidColorBrush(ColorFFromColorRef(cardBorder), _cardBorderBrush.put()));
    }
    if (! _textBrush)
    {
        static_cast<void>(_renderTarget->CreateSolidColorBrush(ColorFFromColorRef(fg), _textBrush.put()));
    }
}

void ViewerPE::EnsureTextLayout(float viewportWidthDip, float viewportHeightDip) noexcept
{
    if (! _writeFactory)
    {
        return;
    }

    _viewportHeightDip = std::max(1.0f, viewportHeightDip);

    if (_textLayout && std::fabs(_layoutWidthDip - viewportWidthDip) <= 0.5f)
    {
        return;
    }

    if (! _baseTextFormat)
    {
        wil::com_ptr<IDWriteTextFormat> fmt;
        HRESULT hr = _writeFactory->CreateTextFormat(
            L"Consolas", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 11.0f, L"en-us", fmt.put());
        if (FAILED(hr))
        {
            hr = _writeFactory->CreateTextFormat(
                L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 11.0f, L"en-us", fmt.put());
        }
        if (FAILED(hr))
        {
            return;
        }

        fmt->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
        _baseTextFormat = std::move(fmt);
    }

    std::wstring text;
    text = std::format(L"{}\n{}\n\n{}", _titleText, _subtitleText, _bodyText);

    wil::com_ptr<IDWriteTextLayout> layout;
    const HRESULT hr = _writeFactory->CreateTextLayout(
        text.c_str(), static_cast<UINT32>(text.size()), _baseTextFormat.get(), std::max(1.0f, viewportWidthDip), 1000000.0f, layout.put());
    if (FAILED(hr))
    {
        return;
    }

    const UINT32 titleLen      = static_cast<UINT32>(std::min<size_t>(_titleText.size(), std::numeric_limits<UINT32>::max()));
    const UINT32 subtitleStart = titleLen + 1u;
    const UINT32 subtitleLen = static_cast<UINT32>(std::min<size_t>(_subtitleText.size(), (subtitleStart <= text.size()) ? (text.size() - subtitleStart) : 0u));

    if (titleLen > 0)
    {
        const DWRITE_TEXT_RANGE r{0u, titleLen};
        layout->SetFontFamilyName(L"Segoe UI", r);
        layout->SetFontSize(20.0f, r);
        layout->SetFontWeight(DWRITE_FONT_WEIGHT_SEMI_BOLD, r);
    }
    if (subtitleLen > 0 && subtitleStart < text.size())
    {
        const DWRITE_TEXT_RANGE r{subtitleStart, subtitleLen};
        layout->SetFontFamilyName(L"Segoe UI", r);
        layout->SetFontSize(12.0f, r);
        layout->SetFontStyle(DWRITE_FONT_STYLE_ITALIC, r);
    }

    DWRITE_TEXT_METRICS metrics{};
    if (SUCCEEDED(layout->GetMetrics(&metrics)))
    {
        _contentHeightDip = metrics.height;
    }
    else
    {
        _contentHeightDip = 0.0f;
    }

    _layoutWidthDip = viewportWidthDip;
    _textLayout     = std::move(layout);
}

void ViewerPE::UpdateScrollBars(HWND hwnd, float viewportHeightDip) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const float maxScrollDip = std::max(0.0f, _contentHeightDip - viewportHeightDip);
    _scrollDip               = std::clamp(_scrollDip, 0.0f, maxScrollDip);

    const float pxPerDip = static_cast<float>(_dpi) / 96.0f;
    const int viewportPx = std::max(1, static_cast<int>(std::lround(viewportHeightDip * pxPerDip)));
    const int contentPx  = std::max(viewportPx, static_cast<int>(std::lround(_contentHeightDip * pxPerDip)));
    const int scrollPx   = std::clamp(static_cast<int>(std::lround(_scrollDip * pxPerDip)), 0, std::max(0, contentPx - viewportPx));

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = std::max(0, contentPx - 1);
    si.nPage  = static_cast<UINT>(viewportPx);
    si.nPos   = scrollPx;
    SetScrollInfo(hwnd, SB_VERT, &si, TRUE);
}

void ViewerPE::SetScrollDip(HWND hwnd, float scrollDip) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (! std::isfinite(scrollDip))
    {
        scrollDip = std::numeric_limits<float>::max();
    }

    const float clamped = ClampScroll(scrollDip, _contentHeightDip, _viewportHeightDip);
    if (std::fabs(clamped - _scrollDip) <= 0.25f)
    {
        return;
    }

    _scrollDip = clamped;
    UpdateScrollBars(hwnd, _viewportHeightDip);
    InvalidateRect(hwnd, nullptr, TRUE);
}

void ViewerPE::ScrollByDip(HWND hwnd, float deltaDip) noexcept
{
    SetScrollDip(hwnd, _scrollDip + deltaDip);
}

void ViewerPE::StartAsyncParse(HWND hwnd, wil::com_ptr<IFileSystem> fileSystem, std::wstring path) noexcept
{
    const uint64_t requestId = _parseRequestId.fetch_add(1) + 1u;
    _titleText               = std::filesystem::path(path).filename().wstring();

    _subtitleText = LoadStringResource(g_hInstance, IDS_VIEWERPE_STATUS_LOADING);
    _bodyText.clear();
    _markdownText.clear();
    _isLoading = true;
    _textLayout.reset();
    _scrollDip = 0.0f;
    InvalidateRect(hwnd, nullptr, TRUE);

    if (! fileSystem)
    {
        auto result = std::unique_ptr<AsyncParseResult>(new (std::nothrow) AsyncParseResult());
        if (! result)
        {
            return;
        }
        result->requestId = requestId;
        result->hr        = E_POINTER;
        result->title     = _titleText;
        result->subtitle  = {};
        result->body      = LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_OPEN_FAILED);
        static_cast<void>(PostMessagePayload(hwnd, kAsyncParseCompleteMessage, 0, std::move(result)));
        return;
    }

    _worker = std::jthread(
        [hwnd, requestId, fileSystem = std::move(fileSystem), path = std::move(path)](std::stop_token st) noexcept
        {
            auto postResult = [&](HRESULT hr, std::wstring title, std::wstring subtitle, std::wstring body, std::wstring markdown) noexcept
            {
                if (st.stop_requested())
                {
                    return;
                }

                if (! hwnd || IsWindow(hwnd) == 0)
                {
                    return;
                }

                auto result = std::unique_ptr<AsyncParseResult>(new (std::nothrow) AsyncParseResult());
                if (! result)
                {
                    return;
                }

                result->requestId = requestId;
                result->hr        = hr;
                result->title     = std::move(title);
                result->subtitle  = std::move(subtitle);
                result->body      = std::move(body);
                result->markdown  = std::move(markdown);
                static_cast<void>(PostMessagePayload(hwnd, kAsyncParseCompleteMessage, 0, std::move(result)));
            };

            wil::com_ptr<IFileSystemIO> fsio;
            HRESULT hr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fsio.put_void());
            if (FAILED(hr) || ! fsio)
            {
                postResult(E_NOINTERFACE, {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_NO_FILEIO), {});
                return;
            }

            wil::com_ptr<IFileReader> reader;
            hr = fsio->CreateFileReader(path.c_str(), reader.put());
            if (FAILED(hr) || ! reader)
            {
                postResult(hr, {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_OPEN_FAILED), {});
                return;
            }

            unsigned __int64 sizeBytes = 0;
            hr                         = reader->GetSize(&sizeBytes);
            if (FAILED(hr) || sizeBytes == 0)
            {
                postResult(hr, {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_OPEN_FAILED), {});
                return;
            }

            if (sizeBytes > static_cast<unsigned __int64>((std::numeric_limits<std::uint32_t>::max)()) ||
                sizeBytes > static_cast<unsigned __int64>((std::numeric_limits<size_t>::max)()))
            {
                postResult(HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE), {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_TOO_LARGE), {});
                return;
            }

            std::vector<std::uint8_t> bytes;
            bytes.resize(static_cast<size_t>(sizeBytes));

            size_t offset = 0;
            while (offset < bytes.size())
            {
                if (st.stop_requested())
                {
                    return;
                }

                const size_t remaining   = bytes.size() - offset;
                const unsigned long want = static_cast<unsigned long>((std::min)(remaining, static_cast<size_t>(16u * 1024u * 1024u)));
                unsigned long read       = 0;
                hr                       = reader->Read(bytes.data() + offset, want, &read);
                if (FAILED(hr))
                {
                    postResult(hr, {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_READ_FAILED), {});
                    return;
                }
                if (read == 0)
                {
                    break;
                }
                offset += static_cast<size_t>(read);
            }

            if (offset == 0)
            {
                postResult(E_FAIL, {}, {}, LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_READ_FAILED), {});
                return;
            }
            if (offset < bytes.size())
            {
                bytes.resize(offset);
            }

            std::unique_ptr<peparse::parsed_pe, ParsedPeDeleter> pe(peparse::ParsePEFromPointer(bytes.data(), static_cast<std::uint32_t>(bytes.size())));
            if (! pe)
            {
                std::wstring err          = LoadStringResource(g_hInstance, IDS_VIEWERPE_ERROR_PARSE_FAILED);
                const std::wstring detail = Utf16FromUtf8(peparse::GetPEErrString());
                const std::wstring loc    = Utf16FromUtf8(peparse::GetPEErrLoc());
                if (! detail.empty())
                {
                    err += L"\n\n";
                    err += detail;
                    if (! loc.empty())
                    {
                        err += L"\n";
                        err += loc;
                    }
                }

                postResult(E_FAIL, {}, {}, std::move(err), {});
                return;
            }

            const auto& nt               = pe->peHeader.nt;
            const std::uint16_t optMagic = nt.OptionalMagic;
            const bool is64              = optMagic == 0x20Bu;

            peparse::VA entryPoint   = 0;
            const bool hasEntryPoint = peparse::GetEntryPoint(pe.get(), entryPoint);

            const std::wstring machine       = MachineText(pe.get());
            const std::wstring subsystem     = SubsystemText(pe.get());
            const std::wstring_view kindName = PeKindName(optMagic);

            std::wstring subtitle;
            subtitle = std::format(L"{}  •  {}  •  {} section(s)", kindName, machine.empty() ? L"Unknown machine" : machine, nt.FileHeader.NumberOfSections);

            const auto& dos        = pe->peHeader.dos;
            const auto& fileHeader = nt.FileHeader;
            const auto& opt32      = nt.OptionalHeader;
            const auto& opt64      = nt.OptionalHeader64;

            const uint64_t imageBase = is64 ? opt64.ImageBase : opt32.ImageBase;

            std::wstring entryPointText;
            if (hasEntryPoint)
            {
                entryPointText = std::format(L"0x{:016X}", static_cast<uint64_t>(entryPoint));
            }
            else
            {
                entryPointText = L"(none)";
            }

            std::wstring body;
            body = std::format(L"Path: {}\nSize: {} ({})\n\nKind: {}\nMachine: {}\nSubsystem: {}\nTimestamp: 0x{:08X}\nCharacteristics: "
                               L"0x{:04X}\nImageBase: 0x{:016X}\nEntryPoint: {}\n\nDOS Header:\n  e_magic: 0x{:04X}\n  e_lfanew: 0x{:08X}\n\nSections:\n",
                               path,
                               sizeBytes,
                               FormatBytesCompact(sizeBytes),
                               kindName,
                               machine.empty() ? L"(unknown)" : machine,
                               subsystem.empty() ? L"(unknown)" : subsystem,
                               static_cast<std::uint32_t>(fileHeader.TimeDateStamp),
                               fileHeader.Characteristics,
                               imageBase,
                               entryPointText,
                               dos.e_magic,
                               dos.e_lfanew);

            body += std::format(L"{:<10} {:>10} {:>10} {:>10} {:>10} {:>10}\n", L"Name", L"RVA", L"VSize", L"RawPtr", L"RawSize", L"Chars");

            struct SectionRow
            {
                std::wstring name;
                std::uint32_t rva     = 0;
                std::uint32_t vsize   = 0;
                std::uint32_t rawPtr  = 0;
                std::uint32_t rawSize = 0;
                std::uint32_t chars   = 0;
            };

            std::vector<SectionRow> sections;
            peparse::IterSec(
                pe.get(),
                [](void* ctx,
                   const peparse::VA& /*va*/,
                   const std::string& name,
                   const peparse::image_section_header& hdr,
                   const peparse::bounded_buffer* /*buf*/) -> int
                {
                    auto* out = static_cast<std::vector<SectionRow>*>(ctx);
                    if (! out)
                    {
                        return 0;
                    }

                    SectionRow row{};
                    row.name.assign(name.begin(), name.end());
                    row.rva     = hdr.VirtualAddress;
                    row.vsize   = hdr.Misc.VirtualSize;
                    row.rawPtr  = hdr.PointerToRawData;
                    row.rawSize = hdr.SizeOfRawData;
                    row.chars   = hdr.Characteristics;

                    out->push_back(std::move(row));
                    return 0;
                },
                &sections);

            for (const auto& sec : sections)
            {
                if (st.stop_requested())
                {
                    return;
                }

                std::wstring line;
                line = std::format(L"{:<10} 0x{:08X} 0x{:08X} 0x{:08X} 0x{:08X} 0x{:08X}\n", sec.name, sec.rva, sec.vsize, sec.rawPtr, sec.rawSize, sec.chars);
                body += line;
            }

            const auto safeAppend = [&](std::wstring_view text) noexcept { body.append(text); };

            const auto safeAppendLine = [&](std::wstring_view text) noexcept
            {
                safeAppend(text);
                safeAppend(L"\n");
            };

            const auto safeAppendFormat = [&]<typename... Args>(std::wformat_string<Args...> fmt, Args&&... args) noexcept
            { safeAppend(std::format(fmt, std::forward<Args>(args)...)); };

            const auto safeAppendBlankLine = [&]() noexcept { safeAppend(L"\n"); };

            safeAppendBlankLine();
            safeAppendLine(L"Rich Header:");
            safeAppendFormat(L"Present: {}\n", pe->peHeader.rich.isPresent ? L"Yes" : L"No");
            safeAppendFormat(L"Valid: {}\n", pe->peHeader.rich.isValid ? L"Yes" : L"No");
            if (pe->peHeader.rich.isPresent)
            {
                safeAppendFormat(L"DecryptionKey: 0x{:08X}\n", pe->peHeader.rich.DecryptionKey);
                safeAppendFormat(L"Checksum: 0x{:08X}\n", pe->peHeader.rich.Checksum);
                safeAppendFormat(L"Entries: {:L}\n", pe->peHeader.rich.Entries.size());

                static constexpr size_t kMaxRichEntries = 256;
                if (! pe->peHeader.rich.Entries.empty())
                {
                    safeAppendBlankLine();
                    safeAppendLine(L"ProductId Build     Count      Product");
                    safeAppendLine(L"-------- -----     ---------  ------------------------------");

                    size_t shown = 0;
                    for (const auto& entry : pe->peHeader.rich.Entries)
                    {
                        if (st.stop_requested())
                        {
                            return;
                        }

                        if (shown >= kMaxRichEntries)
                        {
                            safeAppendFormat(L"... (truncated; showing first {:L} entries)\n", kMaxRichEntries);
                            break;
                        }

                        const std::string& prod = peparse::GetRichProductName(entry.BuildNumber);
                        const std::string& obj  = peparse::GetRichObjectType(entry.ProductId);
                        std::wstring label      = Utf16FromUtf8(prod);
                        if (! obj.empty())
                        {
                            label += L" ";
                            label += Utf16FromUtf8(obj);
                        }

                        safeAppendFormat(L"{:>8} {:>5}     {:>9}  {}\n", entry.ProductId, entry.BuildNumber, entry.Count, label);
                        shown += 1;
                    }
                }
            }

            safeAppendBlankLine();
            safeAppendLine(L"File Header:");
            safeAppendFormat(L"NumberOfSections: {:L}\n", fileHeader.NumberOfSections);
            safeAppendFormat(L"SizeOfOptionalHeader: {:L}\n", fileHeader.SizeOfOptionalHeader);
            safeAppendFormat(L"PointerToSymbolTable: 0x{:08X}\n", fileHeader.PointerToSymbolTable);
            safeAppendFormat(L"NumberOfSymbols: {:L}\n", fileHeader.NumberOfSymbols);

            safeAppendBlankLine();
            safeAppendLine(L"Optional Header:");
            if (is64)
            {
                safeAppendFormat(L"Magic: 0x{:04X}\n", opt64.Magic);
                safeAppendFormat(L"LinkerVersion: {}.{}\n", opt64.MajorLinkerVersion, opt64.MinorLinkerVersion);
                safeAppendFormat(L"SizeOfImage: 0x{:08X}\n", opt64.SizeOfImage);
                safeAppendFormat(L"SizeOfHeaders: 0x{:08X}\n", opt64.SizeOfHeaders);
                safeAppendFormat(L"CheckSum: 0x{:08X}\n", opt64.CheckSum);
                safeAppendFormat(L"DllCharacteristics: 0x{:04X}\n", opt64.DllCharacteristics);
                safeAppendFormat(L"SectionAlignment: 0x{:08X}\n", opt64.SectionAlignment);
                safeAppendFormat(L"FileAlignment: 0x{:08X}\n", opt64.FileAlignment);
                safeAppendFormat(L"OSVersion: {}.{}\n", opt64.MajorOperatingSystemVersion, opt64.MinorOperatingSystemVersion);
                safeAppendFormat(L"ImageVersion: {}.{}\n", opt64.MajorImageVersion, opt64.MinorImageVersion);
                safeAppendFormat(L"SubsystemVersion: {}.{}\n", opt64.MajorSubsystemVersion, opt64.MinorSubsystemVersion);
                safeAppendFormat(L"SizeOfStackReserve: 0x{:016X}\n", opt64.SizeOfStackReserve);
                safeAppendFormat(L"SizeOfStackCommit: 0x{:016X}\n", opt64.SizeOfStackCommit);
                safeAppendFormat(L"SizeOfHeapReserve: 0x{:016X}\n", opt64.SizeOfHeapReserve);
                safeAppendFormat(L"SizeOfHeapCommit: 0x{:016X}\n", opt64.SizeOfHeapCommit);
                safeAppendFormat(L"NumberOfRvaAndSizes: {:L}\n", opt64.NumberOfRvaAndSizes);
            }
            else
            {
                safeAppendFormat(L"Magic: 0x{:04X}\n", opt32.Magic);
                safeAppendFormat(L"LinkerVersion: {}.{}\n", opt32.MajorLinkerVersion, opt32.MinorLinkerVersion);
                safeAppendFormat(L"SizeOfImage: 0x{:08X}\n", opt32.SizeOfImage);
                safeAppendFormat(L"SizeOfHeaders: 0x{:08X}\n", opt32.SizeOfHeaders);
                safeAppendFormat(L"CheckSum: 0x{:08X}\n", opt32.CheckSum);
                safeAppendFormat(L"DllCharacteristics: 0x{:04X}\n", opt32.DllCharacteristics);
                safeAppendFormat(L"SectionAlignment: 0x{:08X}\n", opt32.SectionAlignment);
                safeAppendFormat(L"FileAlignment: 0x{:08X}\n", opt32.FileAlignment);
                safeAppendFormat(L"OSVersion: {}.{}\n", opt32.MajorOperatingSystemVersion, opt32.MinorOperatingSystemVersion);
                safeAppendFormat(L"ImageVersion: {}.{}\n", opt32.MajorImageVersion, opt32.MinorImageVersion);
                safeAppendFormat(L"SubsystemVersion: {}.{}\n", opt32.MajorSubsystemVersion, opt32.MinorSubsystemVersion);
                safeAppendFormat(L"SizeOfStackReserve: 0x{:08X}\n", opt32.SizeOfStackReserve);
                safeAppendFormat(L"SizeOfStackCommit: 0x{:08X}\n", opt32.SizeOfStackCommit);
                safeAppendFormat(L"SizeOfHeapReserve: 0x{:08X}\n", opt32.SizeOfHeapReserve);
                safeAppendFormat(L"SizeOfHeapCommit: 0x{:08X}\n", opt32.SizeOfHeapCommit);
                safeAppendFormat(L"NumberOfRvaAndSizes: {:L}\n", opt32.NumberOfRvaAndSizes);
            }

            safeAppendBlankLine();
            safeAppendLine(L"Data Directories:");
            safeAppendLine(L"Name                 RVA        Size");
            safeAppendLine(L"-------------------  ---------  ---------");
            for (size_t i = 0; i < kDataDirectoryNames.size(); ++i)
            {
                const peparse::data_directory dir = is64 ? opt64.DataDirectory[i] : opt32.DataDirectory[i];
                safeAppendFormat(L"{:<19}  0x{:08X}  0x{:08X}\n", kDataDirectoryNames[i], dir.VirtualAddress, dir.Size);
            }

            if (st.stop_requested())
            {
                return;
            }

            static constexpr size_t kMaxImportRows  = 600;
            static constexpr size_t kHardMaxImports = 6000;

            struct ImportRow
            {
                uint64_t va = 0;
                std::wstring module;
                std::wstring name;
            };

            std::vector<ImportRow> imports;
            imports.reserve(128);

            struct ImportCollect
            {
                std::vector<ImportRow>* out = nullptr;
                size_t maxRows              = 0;
                size_t hardMax              = 0;
                size_t seen                 = 0;
                bool truncated              = false;
            };

            ImportCollect importCollect{&imports, kMaxImportRows, kHardMaxImports, 0, false};
            peparse::IterImpVAString(
                pe.get(),
                [](void* ctx, const peparse::VA& va, const std::string& module, const std::string& name) -> int
                {
                    auto* c = static_cast<ImportCollect*>(ctx);
                    if (! c)
                    {
                        return 0;
                    }

                    c->seen += 1;
                    if (c->out && c->out->size() < c->maxRows)
                    {
                        ImportRow row{};
                        row.va     = static_cast<uint64_t>(va);
                        row.module = Utf16FromUtf8(module);
                        row.name   = Utf16FromUtf8(name);

                        c->out->push_back(std::move(row));
                    }

                    if (c->seen >= c->hardMax)
                    {
                        c->truncated = true;
                        return 1;
                    }

                    return 0;
                },
                &importCollect);

            safeAppendBlankLine();
            if (importCollect.seen == 0)
            {
                safeAppendLine(L"Imports: (none)");
            }
            else
            {
                safeAppendFormat(L"Imports: {:L}{}\n", importCollect.seen, importCollect.truncated ? L"+" : L"");
                safeAppendLine(L"VA                 Module                   Import");
                safeAppendLine(L"-----------------  -----------------------  ------------------------------");
                for (const auto& imp : imports)
                {
                    safeAppendFormat(L"0x{:016X} {:<23}  {}\n", imp.va, imp.module, imp.name);
                }
                if (importCollect.seen > imports.size())
                {
                    safeAppendFormat(L"... (truncated; showing first {:L} entries)\n", imports.size());
                }
            }

            if (st.stop_requested())
            {
                return;
            }

            static constexpr size_t kMaxExportRows  = 600;
            static constexpr size_t kHardMaxExports = 6000;

            struct ExportRow
            {
                uint64_t va       = 0;
                std::uint16_t ord = 0;
                std::wstring name;
                std::wstring forward;
            };

            std::vector<ExportRow> exports;
            exports.reserve(128);

            struct ExportCollect
            {
                std::vector<ExportRow>* out = nullptr;
                size_t maxRows              = 0;
                size_t hardMax              = 0;
                size_t seen                 = 0;
                bool truncated              = false;
            };

            ExportCollect exportCollect{&exports, kMaxExportRows, kHardMaxExports, 0, false};
            peparse::IterExpFull(
                pe.get(),
                [](void* ctx, const peparse::VA& va, std::uint16_t ord, const std::string& name, const std::string& /*module*/, const std::string& forwardStr)
                    -> int
                {
                    auto* c = static_cast<ExportCollect*>(ctx);
                    if (! c)
                    {
                        return 0;
                    }

                    c->seen += 1;
                    if (c->out && c->out->size() < c->maxRows)
                    {
                        ExportRow row{};
                        row.va      = static_cast<uint64_t>(va);
                        row.ord     = ord;
                        row.name    = Utf16FromUtf8(name);
                        row.forward = Utf16FromUtf8(forwardStr);

                        c->out->push_back(std::move(row));
                    }

                    if (c->seen >= c->hardMax)
                    {
                        c->truncated = true;
                        return 1;
                    }

                    return 0;
                },
                &exportCollect);

            safeAppendBlankLine();
            if (exportCollect.seen == 0)
            {
                safeAppendLine(L"Exports: (none)");
            }
            else
            {
                safeAppendFormat(L"Exports: {:L}{}\n", exportCollect.seen, exportCollect.truncated ? L"+" : L"");
                safeAppendLine(L"Ord   VA                 Name");
                safeAppendLine(L"----  -----------------  --------------------------------------------");
                for (const auto& exp : exports)
                {
                    if (exp.va != 0)
                    {
                        safeAppendFormat(L"{:>4} 0x{:016X}  {}\n", exp.ord, exp.va, exp.name);
                    }
                    else if (! exp.forward.empty())
                    {
                        safeAppendFormat(L"{:>4} (forwarded)        {} -> {}\n", exp.ord, exp.name, exp.forward);
                    }
                    else
                    {
                        safeAppendFormat(L"{:>4} (n/a)              {}\n", exp.ord, exp.name);
                    }
                }
                if (exportCollect.seen > exports.size())
                {
                    safeAppendFormat(L"... (truncated; showing first {:L} entries)\n", exports.size());
                }
            }

            if (st.stop_requested())
            {
                return;
            }

            static constexpr size_t kMaxResourceRows  = 600;
            static constexpr size_t kHardMaxResources = 6000;

            struct ResourceRow
            {
                std::wstring type;
                std::wstring name;
                std::wstring lang;
                std::uint32_t codepage = 0;
                std::uint32_t rva      = 0;
                std::uint32_t size     = 0;
            };

            std::vector<ResourceRow> resources;
            resources.reserve(128);

            struct ResourceCollect
            {
                std::vector<ResourceRow>* out = nullptr;
                size_t maxRows                = 0;
                size_t hardMax                = 0;
                size_t seen                   = 0;
                bool truncated                = false;
            };

            ResourceCollect rsrcCollect{&resources, kMaxResourceRows, kHardMaxResources, 0, false};
            peparse::IterRsrc(
                pe.get(),
                [](void* ctx, const peparse::resource& res) -> int
                {
                    auto* c = static_cast<ResourceCollect*>(ctx);
                    if (! c)
                    {
                        return 0;
                    }

                    c->seen += 1;
                    if (c->out && c->out->size() < c->maxRows)
                    {
                        ResourceRow row{};
                        row.type     = Utf16FromUtf8(res.type_str);
                        row.name     = Utf16FromUtf8(res.name_str);
                        row.lang     = Utf16FromUtf8(res.lang_str);
                        row.codepage = res.codepage;
                        row.rva      = res.RVA;
                        row.size     = res.size;

                        if (row.type.empty())
                        {
                            row.type = std::to_wstring(res.type);
                        }
                        if (row.name.empty())
                        {
                            row.name = std::to_wstring(res.name);
                        }
                        if (row.lang.empty())
                        {
                            row.lang = std::to_wstring(res.lang);
                        }

                        c->out->push_back(std::move(row));
                    }

                    if (c->seen >= c->hardMax)
                    {
                        c->truncated = true;
                        return 1;
                    }

                    return 0;
                },
                &rsrcCollect);

            safeAppendBlankLine();
            if (rsrcCollect.seen == 0)
            {
                safeAppendLine(L"Resources: (none)");
            }
            else
            {
                safeAppendFormat(L"Resources: {:L}{}\n", rsrcCollect.seen, rsrcCollect.truncated ? L"+" : L"");
                safeAppendLine(L"RVA        Size       CodePage  Type / Name / Lang");
                safeAppendLine(L"---------  ---------  --------  ---------------------------------------");
                for (const auto& res : resources)
                {
                    safeAppendFormat(L"0x{:08X} 0x{:08X} {:>8}  {}/{}/{}\n", res.rva, res.size, res.codepage, res.type, res.name, res.lang);
                }
                if (rsrcCollect.seen > resources.size())
                {
                    safeAppendFormat(L"... (truncated; showing first {:L} entries)\n", resources.size());
                }
            }

            if (st.stop_requested())
            {
                return;
            }

            static constexpr size_t kMaxRelocRows  = 600;
            static constexpr size_t kHardMaxRelocs = 20000;

            struct RelocRow
            {
                uint64_t va              = 0;
                peparse::reloc_type type = peparse::RELOC_ABSOLUTE;
            };

            std::vector<RelocRow> relocs;
            relocs.reserve(256);

            struct RelocCollect
            {
                std::vector<RelocRow>* out = nullptr;
                size_t maxRows             = 0;
                size_t hardMax             = 0;
                size_t seen                = 0;
                bool truncated             = false;
            };

            RelocCollect relocCollect{&relocs, kMaxRelocRows, kHardMaxRelocs, 0, false};
            peparse::IterRelocs(
                pe.get(),
                [](void* ctx, const peparse::VA& va, const peparse::reloc_type& type) -> int
                {
                    auto* c = static_cast<RelocCollect*>(ctx);
                    if (! c)
                    {
                        return 0;
                    }

                    c->seen += 1;
                    if (c->out && c->out->size() < c->maxRows)
                    {
                        RelocRow row{};
                        row.va   = static_cast<uint64_t>(va);
                        row.type = type;

                        c->out->push_back(row);
                    }

                    if (c->seen >= c->hardMax)
                    {
                        c->truncated = true;
                        return 1;
                    }

                    return 0;
                },
                &relocCollect);

            safeAppendBlankLine();
            if (relocCollect.seen == 0)
            {
                safeAppendLine(L"Relocations: (none)");
            }
            else
            {
                safeAppendFormat(L"Relocations: {:L}{}\n", relocCollect.seen, relocCollect.truncated ? L"+" : L"");
                safeAppendLine(L"VA                 Type");
                safeAppendLine(L"-----------------  ----");
                for (const auto& rel : relocs)
                {
                    safeAppendFormat(L"0x{:016X}  {:L}\n", rel.va, static_cast<unsigned long>(rel.type));
                }
                if (relocCollect.seen > relocs.size())
                {
                    safeAppendFormat(L"... (truncated; showing first {:L} entries)\n", relocs.size());
                }
            }

            if (st.stop_requested())
            {
                return;
            }

            static constexpr size_t kMaxDebugRows  = 256;
            static constexpr size_t kHardMaxDebugs = 2000;

            struct DebugRow
            {
                std::uint32_t type = 0;
                std::uint32_t size = 0;
            };

            std::vector<DebugRow> debugs;
            debugs.reserve(32);

            struct DebugCollect
            {
                std::vector<DebugRow>* out = nullptr;
                size_t maxRows             = 0;
                size_t hardMax             = 0;
                size_t seen                = 0;
                bool truncated             = false;
            };

            DebugCollect debugCollect{&debugs, kMaxDebugRows, kHardMaxDebugs, 0, false};
            peparse::IterDebugs(
                pe.get(),
                [](void* ctx, const std::uint32_t& type, const peparse::bounded_buffer* buf) -> int
                {
                    auto* c = static_cast<DebugCollect*>(ctx);
                    if (! c)
                    {
                        return 0;
                    }

                    c->seen += 1;
                    if (c->out && c->out->size() < c->maxRows)
                    {
                        DebugRow row{};
                        row.type = type;
                        row.size = buf ? buf->bufLen : 0;
                        c->out->push_back(row);
                    }

                    if (c->seen >= c->hardMax)
                    {
                        c->truncated = true;
                        return 1;
                    }

                    return 0;
                },
                &debugCollect);

            safeAppendBlankLine();
            if (debugCollect.seen == 0)
            {
                safeAppendLine(L"Debug Directories: (none)");
            }
            else
            {
                safeAppendFormat(L"Debug Directories: {:L}{}\n", debugCollect.seen, debugCollect.truncated ? L"+" : L"");
                safeAppendLine(L"Type       Size");
                safeAppendLine(L"---------  ---------");
                for (const auto& dbg : debugs)
                {
                    safeAppendFormat(L"{:>9} 0x{:08X}\n", dbg.type, dbg.size);
                }
                if (debugCollect.seen > debugs.size())
                {
                    safeAppendFormat(L"... (truncated; showing first {:L} entries)\n", debugs.size());
                }
            }

            if (st.stop_requested())
            {
                return;
            }

            static constexpr size_t kMaxSymbolRows  = 600;
            static constexpr size_t kHardMaxSymbols = 6000;

            struct SymbolRow
            {
                std::wstring name;
                std::uint32_t value  = 0;
                std::int16_t section = 0;
                std::uint16_t type   = 0;
                std::uint8_t storage = 0;
                std::uint8_t aux     = 0;
            };

            std::vector<SymbolRow> symbols;
            symbols.reserve(128);

            struct SymbolCollect
            {
                std::vector<SymbolRow>* out = nullptr;
                size_t maxRows              = 0;
                size_t hardMax              = 0;
                size_t seen                 = 0;
                bool truncated              = false;
            };

            SymbolCollect symbolCollect{&symbols, kMaxSymbolRows, kHardMaxSymbols, 0, false};
            peparse::IterSymbols(
                pe.get(),
                [](void* ctx,
                   const std::string& name,
                   const std::uint32_t& value,
                   const std::int16_t& section,
                   const std::uint16_t& type,
                   const std::uint8_t& storage,
                   const std::uint8_t& aux) -> int
                {
                    auto* c = static_cast<SymbolCollect*>(ctx);
                    if (! c)
                    {
                        return 0;
                    }

                    c->seen += 1;
                    if (c->out && c->out->size() < c->maxRows)
                    {
                        SymbolRow row{};
                        row.name    = Utf16FromUtf8(name);
                        row.value   = value;
                        row.section = section;
                        row.type    = type;
                        row.storage = storage;
                        row.aux     = aux;

                        c->out->push_back(std::move(row));
                    }

                    if (c->seen >= c->hardMax)
                    {
                        c->truncated = true;
                        return 1;
                    }

                    return 0;
                },
                &symbolCollect);

            safeAppendBlankLine();
            if (symbolCollect.seen == 0)
            {
                safeAppendLine(L"Symbols: (none)");
            }
            else
            {
                safeAppendFormat(L"Symbols: {:L}{}\n", symbolCollect.seen, symbolCollect.truncated ? L"+" : L"");
                safeAppendLine(L"Value      Sect Type  Stor Aux Name");
                safeAppendLine(L"---------  ---- ----  ---- --- --------------------------------");
                for (const auto& sym : symbols)
                {
                    safeAppendFormat(L"0x{:08X} {:>4} 0x{:04X} {:>4} {:>3} {}\n", sym.value, sym.section, sym.type, sym.storage, sym.aux, sym.name);
                }
                if (symbolCollect.seen > symbols.size())
                {
                    safeAppendFormat(L"... (truncated; showing first {:L} entries)\n", symbols.size());
                }
            }

            std::wstring title;
            title = std::filesystem::path(path).filename().wstring();

            std::wstring markdown;
            const std::wstring mdTitle = title.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERPE_NAME) : title;
            markdown                   = std::format(L"# {}\n\n{}\n\n```text\n{}\n```\n", mdTitle, subtitle, body);

            postResult(S_OK, std::move(title), std::move(subtitle), std::move(body), std::move(markdown));
        });
}

void ViewerPE::OnAsyncParseComplete(std::unique_ptr<AsyncParseResult> result) noexcept
{
    if (! result)
    {
        return;
    }

    if (result->requestId != _parseRequestId.load())
    {
        return;
    }

    _titleText    = std::move(result->title);
    _subtitleText = std::move(result->subtitle);
    _bodyText     = std::move(result->body);
    _markdownText = std::move(result->markdown);
    _isLoading    = false;

    if (_hWnd)
    {
        std::wstring title;
        const std::wstring viewerName = LoadStringResource(g_hInstance, IDS_VIEWERPE_NAME);
        title                         = _titleText.empty() ? viewerName : std::format(L"{} — {}", _titleText, viewerName);
        SetWindowTextW(_hWnd.get(), title.c_str());
    }

    _textLayout.reset();
    _scrollDip = 0.0f;
    if (_hWnd)
    {
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

HRESULT STDMETHODCALLTYPE ViewerPE::Open(const ViewerOpenContext* context) noexcept
{
    if (! context || ! context->focusedPath || ! context->fileSystem)
    {
        return E_INVALIDARG;
    }

    if (! RegisterWndClass(g_hInstance))
    {
        return E_FAIL;
    }

    const std::filesystem::path focused(context->focusedPath);
    const std::wstring fileName = focused.filename().wstring();

    if (! _hWnd)
    {
        RECT ownerRect{};
        if (context->ownerWindow && GetWindowRect(context->ownerWindow, &ownerRect) != 0)
        {
            const int w = std::max(1, static_cast<int>(ownerRect.right - ownerRect.left));
            const int h = std::max(1, static_cast<int>(ownerRect.bottom - ownerRect.top));

            wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(LoadMenuW(g_hInstance, MAKEINTRESOURCEW(IDR_VIEWERPE_MENU)));
            HWND window = CreateWindowExW(0,
                                          kClassName,
                                          L"",
                                          WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_VSCROLL,
                                          ownerRect.left,
                                          ownerRect.top,
                                          w,
                                          h,
                                          nullptr,
                                          menu.get(),
                                          g_hInstance,
                                          this);
            if (! window)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerPE: CreateWindowExW failed.");
                return HRESULT_FROM_WIN32(lastError);
            }

            menu.release();
            _hWnd.reset(window);
        }
        else
        {
            wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(LoadMenuW(g_hInstance, MAKEINTRESOURCEW(IDR_VIEWERPE_MENU)));
            HWND window = CreateWindowExW(0,
                                          kClassName,
                                          L"",
                                          WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_VSCROLL,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          900,
                                          700,
                                          nullptr,
                                          menu.get(),
                                          g_hInstance,
                                          this);
            if (! window)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerPE: CreateWindowExW failed.");
                return HRESULT_FROM_WIN32(lastError);
            }

            menu.release();
            _hWnd.reset(window);
        }

        ApplyTitleBarTheme(true);

        AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
    }
    else
    {
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
    }

    _fileSystem = context->fileSystem;

    _currentPath = context->focusedPath;
    _otherFiles.clear();
    if (context->otherFiles && context->otherFileCount > 0)
    {
        _otherFiles.reserve(static_cast<size_t>(context->otherFileCount));

        for (unsigned long i = 0; i < context->otherFileCount; ++i)
        {
            const wchar_t* s = context->otherFiles[i];
            if (! s)
            {
                continue;
            }

            _otherFiles.emplace_back(s);
        }
    }

    if (_otherFiles.size() > 1 && context->focusedOtherFileIndex < _otherFiles.size())
    {
        _otherIndex = static_cast<size_t>(context->focusedOtherFileIndex);
    }
    else
    {
        _otherIndex = 0;
    }

    if (_otherFiles.empty())
    {
        _otherFiles.push_back(_currentPath);
        _otherIndex = 0;
    }

    if (_hWnd)
    {
        RefreshFileCombo(_hWnd.get());
        UpdateMenuState(_hWnd.get());
    }

    _titleText    = fileName;
    _subtitleText = LoadStringResource(g_hInstance, IDS_VIEWERPE_STATUS_LOADING);
    _bodyText.clear();
    _markdownText.clear();
    _textLayout.reset();
    _scrollDip = 0.0f;

    StartAsyncParse(_hWnd.get(), _fileSystem, _currentPath);

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::Close() noexcept
{
    _hWnd.reset();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::SetTheme(const ViewerTheme* theme) noexcept
{
    if (! theme || theme->version != 2)
    {
        return E_INVALIDARG;
    }

    _theme    = *theme;
    _hasTheme = true;

    if (_hWnd)
    {
        const bool active = GetActiveWindow() == _hWnd.get();
        ApplyTitleBarTheme(active);
        _textBrush.reset();
        _cardBorderBrush.reset();
        _cardBrush.reset();
        _bgBrush.reset();
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerPE::SetCallback(IViewerCallback* callback, void* cookie) noexcept
{
    _callback       = callback;
    _callbackCookie = cookie;
    return S_OK;
}
