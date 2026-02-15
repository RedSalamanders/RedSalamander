#include "ViewerImgRaw.h"

#include "FluentIcons.h"
#include "ViewerImgRaw.Internal.h"
#include "ViewerImgRaw.ThemeHelpers.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstring>
#include <ctime>
#include <cwctype>
#include <format>
#include <limits>
#include <new>
#include <thread>

#include <commctrl.h>
#include <d2d1.h>
#include <d2d1helper.h>
#include <dwmapi.h>
#include <dwrite.h>
#include <dxgiformat.h>
#include <shobjidl_core.h>
#include <uxtheme.h>
#include <wincodec.h>

#include <libraw/libraw.h>
#include <turbojpeg.h>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#pragma comment(lib, "comctl32")
#pragma comment(lib, "d2d1")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "dwrite")
#pragma comment(lib, "uxtheme")
#pragma comment(lib, "windowscodecs")
#pragma comment(lib, "turbojpeg")
#if defined(_DEBUG)
#pragma comment(lib, "raw_rd")
#else
#pragma comment(lib, "raw_r")
#endif

#include "Helpers.h"
#include "resource.h"

namespace
{
constexpr int kHeaderHeightDip = 28;
constexpr int kStatusHeightDip = 22;

constexpr UINT_PTR kLoadingDelayTimerId  = 1;
constexpr UINT_PTR kLoadingAnimTimerId   = 2;
constexpr UINT kLoadingDelayMs           = 200u;
constexpr UINT kLoadingAnimIntervalMs    = 16u;
constexpr float kLoadingSpinnerDegPerSec = 90.0f;

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

wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> g_viewerImgRawMenuIconFont;
UINT g_viewerImgRawMenuIconFontDpi   = USER_DEFAULT_SCREEN_DPI;
bool g_viewerImgRawMenuIconFontValid = false;

[[nodiscard]] bool EnsureViewerImgRawMenuIconFont(HDC hdc, UINT dpi) noexcept
{
    if (! hdc)
    {
        return false;
    }

    if (dpi == 0)
    {
        dpi = USER_DEFAULT_SCREEN_DPI;
    }

    if (dpi != g_viewerImgRawMenuIconFontDpi || ! g_viewerImgRawMenuIconFont)
    {
        g_viewerImgRawMenuIconFont      = FluentIcons::CreateFontForDpi(dpi, FluentIcons::kDefaultSizeDip);
        g_viewerImgRawMenuIconFontDpi   = dpi;
        g_viewerImgRawMenuIconFontValid = false;

        if (g_viewerImgRawMenuIconFont)
        {
            g_viewerImgRawMenuIconFontValid = FluentIcons::FontHasGlyph(hdc, g_viewerImgRawMenuIconFont.get(), FluentIcons::kChevronRightSmall) &&
                                              FluentIcons::FontHasGlyph(hdc, g_viewerImgRawMenuIconFont.get(), FluentIcons::kCheckMark);
        }
    }

    return g_viewerImgRawMenuIconFontValid;
}

[[nodiscard]] std::wstring KeyGlyphFromVirtualKey(UINT vk, HKL keyboardLayout) noexcept
{
    if (! keyboardLayout)
    {
        return {};
    }

    const UINT scanCode = MapVirtualKeyExW(vk, MAPVK_VK_TO_VSC, keyboardLayout);
    if (scanCode == 0)
    {
        return {};
    }

    std::array<BYTE, 256> keyboardState{};
    std::array<wchar_t, 8> buffer{};
    const int result = ToUnicodeEx(vk, scanCode, keyboardState.data(), buffer.data(), static_cast<int>(buffer.size() - 1), 0, keyboardLayout);

    if (result > 0)
    {
        std::wstring out(buffer.data(), buffer.data() + result);
        if (! out.empty() && ! std::iswcntrl(out.front()))
        {
            return out;
        }
    }
    else if (result < 0)
    {
        std::array<wchar_t, 8> clearBuf{};
        static_cast<void>(ToUnicodeEx(vk, scanCode, keyboardState.data(), clearBuf.data(), static_cast<int>(clearBuf.size() - 1), 0, keyboardLayout));
    }

    wchar_t nameBuf[64]{};
    const LONG lParam = static_cast<LONG>(scanCode << 16);
    const int nameLen = GetKeyNameTextW(lParam, nameBuf, static_cast<int>(std::size(nameBuf)));
    if (nameLen > 0)
    {
        return std::wstring(nameBuf, nameBuf + nameLen);
    }

    return {};
}

[[nodiscard]] D2D1_MATRIX_3X2_F ExifOrientationTransform(uint16_t orientation, float widthDip, float heightDip) noexcept
{
    switch (orientation)
    {
        case 2: return D2D1::Matrix3x2F(-1.0f, 0.0f, 0.0f, 1.0f, widthDip, 0.0f);
        case 3: return D2D1::Matrix3x2F(-1.0f, 0.0f, 0.0f, -1.0f, widthDip, heightDip);
        case 4: return D2D1::Matrix3x2F(1.0f, 0.0f, 0.0f, -1.0f, 0.0f, heightDip);
        case 5: return D2D1::Matrix3x2F(0.0f, 1.0f, 1.0f, 0.0f, 0.0f, 0.0f);
        case 6: return D2D1::Matrix3x2F(0.0f, 1.0f, -1.0f, 0.0f, heightDip, 0.0f);
        case 7: return D2D1::Matrix3x2F(0.0f, -1.0f, -1.0f, 0.0f, heightDip, widthDip);
        case 8: return D2D1::Matrix3x2F(0.0f, -1.0f, 1.0f, 0.0f, 0.0f, widthDip);
        default: return D2D1::Matrix3x2F::Identity();
    }
}

constexpr char kViewerImgRawSchemaJson[] = R"json({
    "version": 1,
    "title": "Image Viewer",
    "fields": [
        {
            "key": "halfSize",
            "type": "bool",
            "label": "Half size",
            "description": "Decode at half resolution for faster loading and lower memory use.",
            "default": true
        },
        {
            "key": "preferThumbnail",
            "type": "bool",
            "label": "Prefer thumbnail",
            "description": "Open images in Thumbnail mode by default (uses sidecar JPEG when present, otherwise embedded thumbnail when available).",
            "default": true
        },
        {
            "key": "useCameraWb",
            "type": "bool",
            "label": "Use camera white balance",
            "default": true
        },
        {
            "key": "autoWb",
            "type": "bool",
            "label": "Auto white balance",
            "default": false
        },
        {
            "key": "zoomOnClickPercent",
            "type": "value",
            "label": "Zoom on click (%)",
            "description": "Temporary zoom level (percent) while the left mouse button is held down on the image.",
            "default": 50,
            "min": 1,
            "max": 6400
        },
        {
            "key": "prevCache",
            "type": "value",
            "label": "Keep previous",
            "description": "Number of previous images to keep decoded in memory.",
            "default": 1,
            "min": 0,
            "max": 8
        },
        {
            "key": "nextCache",
            "type": "value",
            "label": "Keep next",
            "description": "Number of next images to keep decoded in memory.",
            "default": 1,
            "min": 0,
            "max": 8
        },
        {
            "key": "exportJpegQualityPercent",
            "type": "value",
            "label": "Export JPEG quality (%)",
            "default": 90,
            "min": 1,
            "max": 100
        },
        {
            "key": "exportJpegSubsampling",
            "type": "value",
            "label": "Export JPEG subsampling",
            "description": "WICJpegYCrCbSubsamplingOption: 0=Default, 1=420, 2=422, 3=444, 4=440.",
            "default": 0,
            "min": 0,
            "max": 4
        },
        {
            "key": "exportPngFilter",
            "type": "value",
            "label": "Export PNG filter",
            "description": "WICPngFilterOption: 0=Unspecified, 1=None, 2=Sub, 3=Up, 4=Average, 5=Paeth, 6=Adaptive.",
            "default": 0,
            "min": 0,
            "max": 6
        },
        {
            "key": "exportPngInterlace",
            "type": "bool",
            "label": "Export PNG interlace",
            "default": false
        },
        {
            "key": "exportTiffCompression",
            "type": "value",
            "label": "Export TIFF compression",
            "description": "WICTiffCompressionOption: 0=DontCare, 1=None, 2=CCITT3, 3=CCITT4, 4=LZW, 5=RLE, 6=ZIP, 7=LZWHDifferencing.",
            "default": 0,
            "min": 0,
            "max": 7
        },
        {
            "key": "exportBmpUseV5Header32bppBGRA",
            "type": "bool",
            "label": "Export BMP V5 header (BGRA)",
            "default": true
        },
        {
            "key": "exportGifInterlace",
            "type": "bool",
            "label": "Export GIF interlace",
            "default": false
        },
        {
            "key": "exportWmpQualityPercent",
            "type": "value",
            "label": "Export JPEG XR quality (%)",
            "default": 90,
            "min": 1,
            "max": 100
        },
        {
            "key": "exportWmpLossless",
            "type": "bool",
            "label": "Export JPEG XR lossless",
            "default": false
        }
    ]
})json";

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

COLORREF ColorFromHSV(float hueDegrees, float saturation, float value) noexcept
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

int PxFromDip(int dip, UINT dpi) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
}

float DipsFromPixels(int px, UINT dpi) noexcept
{
    if (dpi == 0)
    {
        return static_cast<float>(px);
    }

    return static_cast<float>(px) * 96.0f / static_cast<float>(dpi);
}

float DipsFromPixelsF(float px, UINT dpi) noexcept
{
    if (dpi == 0)
    {
        return px;
    }

    return px * 96.0f / static_cast<float>(dpi);
}

D2D1_RECT_F RectFFromPixels(const RECT& rc, UINT dpi) noexcept
{
    const float left   = DipsFromPixels(static_cast<int>(rc.left), dpi);
    const float top    = DipsFromPixels(static_cast<int>(rc.top), dpi);
    const float right  = DipsFromPixels(static_cast<int>(rc.right), dpi);
    const float bottom = DipsFromPixels(static_cast<int>(rc.bottom), dpi);
    return D2D1::RectF(left, top, right, bottom);
}

D2D1_COLOR_F ColorFFromColorRef(COLORREF color, float alpha = 1.0f) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;
    return D2D1::ColorF(r, g, b, alpha);
}

struct ViewerImgRawClassBackgroundBrushState
{
    ViewerImgRawClassBackgroundBrushState()                                                        = default;
    ViewerImgRawClassBackgroundBrushState(const ViewerImgRawClassBackgroundBrushState&)            = delete;
    ViewerImgRawClassBackgroundBrushState& operator=(const ViewerImgRawClassBackgroundBrushState&) = delete;
    ViewerImgRawClassBackgroundBrushState(ViewerImgRawClassBackgroundBrushState&&)                 = delete;
    ViewerImgRawClassBackgroundBrushState& operator=(ViewerImgRawClassBackgroundBrushState&&)      = delete;

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> activeBrush;
    COLORREF activeColor = CLR_INVALID;

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> pendingBrush;
    COLORREF pendingColor = CLR_INVALID;

    bool classRegistered = false;
};

ViewerImgRawClassBackgroundBrushState g_viewerImgRawClassBackgroundBrush;

HBRUSH GetActiveViewerImgRawClassBackgroundBrush() noexcept
{
    if (g_viewerImgRawClassBackgroundBrush.pendingBrush)
    {
        return g_viewerImgRawClassBackgroundBrush.pendingBrush.get();
    }

    if (! g_viewerImgRawClassBackgroundBrush.activeBrush)
    {
        const COLORREF fallback                        = GetSysColor(COLOR_WINDOW);
        g_viewerImgRawClassBackgroundBrush.activeColor = fallback;
        g_viewerImgRawClassBackgroundBrush.activeBrush.reset(CreateSolidBrush(fallback));
    }

    return g_viewerImgRawClassBackgroundBrush.activeBrush.get();
}

void RequestViewerImgRawClassBackgroundColor(COLORREF color) noexcept
{
    if (color == CLR_INVALID)
    {
        return;
    }

    if (g_viewerImgRawClassBackgroundBrush.pendingBrush && g_viewerImgRawClassBackgroundBrush.pendingColor == color)
    {
        return;
    }

    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> brush(CreateSolidBrush(color));
    if (! brush)
    {
        return;
    }

    g_viewerImgRawClassBackgroundBrush.pendingColor = color;
    g_viewerImgRawClassBackgroundBrush.pendingBrush = std::move(brush);
}

void ApplyPendingViewerImgRawClassBackgroundBrush(HWND hwnd) noexcept
{
    if (! hwnd || ! g_viewerImgRawClassBackgroundBrush.pendingBrush || ! g_viewerImgRawClassBackgroundBrush.classRegistered)
    {
        return;
    }

    const LONG_PTR newBrush = reinterpret_cast<LONG_PTR>(g_viewerImgRawClassBackgroundBrush.pendingBrush.get());
    SetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND, newBrush);

    g_viewerImgRawClassBackgroundBrush.activeBrush  = std::move(g_viewerImgRawClassBackgroundBrush.pendingBrush);
    g_viewerImgRawClassBackgroundBrush.activeColor  = g_viewerImgRawClassBackgroundBrush.pendingColor;
    g_viewerImgRawClassBackgroundBrush.pendingColor = CLR_INVALID;
}

} // namespace

ViewerImgRaw::ViewerImgRaw()
{
    _metaId          = L"builtin/viewer-imgraw";
    _metaShortId     = L"viewimgraw";
    _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERRAW_NAME);
    _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERRAW_DESCRIPTION);

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = nullptr;

    static_cast<void>(SetConfiguration(nullptr));
}

ViewerImgRaw::~ViewerImgRaw() = default;

void ViewerImgRaw::SetHost(IHost* host) noexcept
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

HRESULT STDMETHODCALLTYPE ViewerImgRaw::QueryInterface(REFIID riid, void** ppvObject) noexcept
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

ULONG STDMETHODCALLTYPE ViewerImgRaw::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE ViewerImgRaw::Release() noexcept
{
    const ULONG refs = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (refs == 0)
    {
        delete this;
    }
    return refs;
}

HRESULT STDMETHODCALLTYPE ViewerImgRaw::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerImgRaw::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = kViewerImgRawSchemaJson;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerImgRaw::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    _config = {};

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        yyjson_doc* doc = yyjson_read(configurationJsonUtf8, strlen(configurationJsonUtf8), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (! doc)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

        yyjson_val* root = yyjson_doc_get_root(doc);
        if (! root || ! yyjson_is_obj(root))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        if (yyjson_val* v = yyjson_obj_get(root, "halfSize"); v && yyjson_is_bool(v))
        {
            _config.halfSize = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(root, "preferThumbnail"); v && yyjson_is_bool(v))
        {
            _config.preferThumbnail = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(root, "useCameraWb"); v && yyjson_is_bool(v))
        {
            _config.useCameraWb = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(root, "autoWb"); v && yyjson_is_bool(v))
        {
            _config.autoWb = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(root, "zoomOnClickPercent"); v && yyjson_is_int(v))
        {
            const int64_t value = yyjson_get_int(v);
            if (value > 0)
            {
                _config.zoomOnClickPercent = static_cast<uint32_t>(std::clamp<int64_t>(value, 1, 6400));
            }
        }
        if (yyjson_val* v = yyjson_obj_get(root, "prevCache"); v && yyjson_is_int(v))
        {
            const int64_t value = yyjson_get_int(v);
            if (value >= 0)
            {
                _config.prevCache = static_cast<uint32_t>(std::clamp<int64_t>(value, 0, 8));
            }
        }
        if (yyjson_val* v = yyjson_obj_get(root, "nextCache"); v && yyjson_is_int(v))
        {
            const int64_t value = yyjson_get_int(v);
            if (value >= 0)
            {
                _config.nextCache = static_cast<uint32_t>(std::clamp<int64_t>(value, 0, 8));
            }
        }

        if (yyjson_val* v = yyjson_obj_get(root, "exportJpegQualityPercent"); v && yyjson_is_int(v))
        {
            const int64_t value = yyjson_get_int(v);
            if (value > 0)
            {
                _config.exportJpegQualityPercent = static_cast<uint32_t>(std::clamp<int64_t>(value, 1, 100));
            }
        }
        if (yyjson_val* v = yyjson_obj_get(root, "exportJpegSubsampling"); v && yyjson_is_int(v))
        {
            const int64_t value = yyjson_get_int(v);
            if (value >= 0)
            {
                _config.exportJpegSubsampling = static_cast<uint32_t>(std::clamp<int64_t>(value, 0, 4));
            }
        }
        if (yyjson_val* v = yyjson_obj_get(root, "exportPngFilter"); v && yyjson_is_int(v))
        {
            const int64_t value = yyjson_get_int(v);
            if (value >= 0)
            {
                _config.exportPngFilter = static_cast<uint32_t>(std::clamp<int64_t>(value, 0, 6));
            }
        }
        if (yyjson_val* v = yyjson_obj_get(root, "exportPngInterlace"); v && yyjson_is_bool(v))
        {
            _config.exportPngInterlace = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(root, "exportTiffCompression"); v && yyjson_is_int(v))
        {
            const int64_t value = yyjson_get_int(v);
            if (value >= 0)
            {
                _config.exportTiffCompression = static_cast<uint32_t>(std::clamp<int64_t>(value, 0, 7));
            }
        }
        if (yyjson_val* v = yyjson_obj_get(root, "exportBmpUseV5Header32bppBGRA"); v && yyjson_is_bool(v))
        {
            _config.exportBmpUseV5Header32bppBGRA = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(root, "exportGifInterlace"); v && yyjson_is_bool(v))
        {
            _config.exportGifInterlace = yyjson_get_bool(v);
        }
        if (yyjson_val* v = yyjson_obj_get(root, "exportWmpQualityPercent"); v && yyjson_is_int(v))
        {
            const int64_t value = yyjson_get_int(v);
            if (value > 0)
            {
                _config.exportWmpQualityPercent = static_cast<uint32_t>(std::clamp<int64_t>(value, 1, 100));
            }
        }
        if (yyjson_val* v = yyjson_obj_get(root, "exportWmpLossless"); v && yyjson_is_bool(v))
        {
            _config.exportWmpLossless = yyjson_get_bool(v);
        }
    }

    _displayMode   = _config.preferThumbnail ? DisplayMode::Thumbnail : DisplayMode::Raw;
    _displayedMode = _displayMode;

    _configJson = std::format(
        R"json({{"halfSize":{},"preferThumbnail":{},"useCameraWb":{},"autoWb":{},"zoomOnClickPercent":{},"prevCache":{},"nextCache":{},"exportJpegQualityPercent":{},"exportJpegSubsampling":{},"exportPngFilter":{},"exportPngInterlace":{},"exportTiffCompression":{},"exportBmpUseV5Header32bppBGRA":{},"exportGifInterlace":{},"exportWmpQualityPercent":{},"exportWmpLossless":{}}})json",
        _config.halfSize,
        _config.preferThumbnail,
        _config.useCameraWb,
        _config.autoWb,
        _config.zoomOnClickPercent,
        _config.prevCache,
        _config.nextCache,
        _config.exportJpegQualityPercent,
        _config.exportJpegSubsampling,
        _config.exportPngFilter,
        _config.exportPngInterlace,
        _config.exportTiffCompression,
        _config.exportBmpUseV5Header32bppBGRA,
        _config.exportGifInterlace,
        _config.exportWmpQualityPercent,
        _config.exportWmpLossless);

    ClearImageCache();

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerImgRaw::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *configurationJsonUtf8 = _configJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerImgRaw::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    const bool isDefault = (_config == Config{});
    *pSomethingToSave    = isDefault ? FALSE : TRUE;
    return S_OK;
}

ATOM ViewerImgRaw::RegisterWndClass(HINSTANCE instance) noexcept
{
    if (! instance)
    {
        return 0;
    }

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);

    const ATOM existing = GetClassInfoExW(instance, kClassName, &wc) != 0 ? static_cast<ATOM>(1) : static_cast<ATOM>(0);
    if (existing != 0)
    {
        g_viewerImgRawClassBackgroundBrush.classRegistered = true;
        return existing;
    }

    wc               = {};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wc.lpfnWndProc   = &ViewerImgRaw::WndProcThunk;
    wc.cbClsExtra    = 0;
    wc.cbWndExtra    = 0;
    wc.hInstance     = instance;
    wc.hIcon         = LoadIconW(nullptr, IDI_APPLICATION);
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = GetActiveViewerImgRawClassBackgroundBrush();
    wc.lpszMenuName  = nullptr;
    wc.lpszClassName = kClassName;
    wc.hIconSm       = wc.hIcon;

    const ATOM atom = RegisterClassExW(&wc);
    if (atom != 0)
    {
        g_viewerImgRawClassBackgroundBrush.classRegistered = true;
    }

    return atom;
}

LRESULT CALLBACK ViewerImgRaw::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ViewerImgRaw* self = reinterpret_cast<ViewerImgRaw*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self           = reinterpret_cast<ViewerImgRaw*>(cs ? cs->lpCreateParams : nullptr);
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

LRESULT ViewerImgRaw::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    switch (msg)
    {
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_SIZE: OnSize(static_cast<UINT>(LOWORD(lp)), static_cast<UINT>(HIWORD(lp))); return 0;
        case WM_COMMAND: OnCommand(hwnd, static_cast<UINT>(LOWORD(wp)), static_cast<UINT>(HIWORD(wp)), reinterpret_cast<HWND>(lp)); return 0;
        case WM_LBUTTONDOWN: OnLButtonDown(hwnd, static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))); return 0;
        case WM_LBUTTONDBLCLK:
            OnLButtonDblClick(hwnd, static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp))));
            return 0;
        case WM_LBUTTONUP: OnLButtonUp(hwnd); return 0;
        case WM_MOUSEMOVE: OnMouseMove(hwnd, static_cast<int>(static_cast<short>(LOWORD(lp))), static_cast<int>(static_cast<short>(HIWORD(lp)))); return 0;
        case WM_CAPTURECHANGED: OnCaptureChanged(); return 0;
        case WM_TIMER: OnTimer(static_cast<UINT_PTR>(wp)); return 0;
        case WM_MEASUREITEM: return OnMeasureItem(hwnd, reinterpret_cast<MEASUREITEMSTRUCT*>(lp));
        case WM_DRAWITEM: return OnDrawItem(hwnd, reinterpret_cast<DRAWITEMSTRUCT*>(lp));
        case WM_KEYDOWN: OnKeyDown(hwnd, static_cast<UINT>(wp)); return 0;
        case WM_SYSKEYDOWN:
            if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
            {
                OnKeyDown(hwnd, static_cast<UINT>(wp));
                return 0;
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_HSCROLL: OnHScroll(hwnd, static_cast<UINT>(LOWORD(wp))); return 0;
        case WM_VSCROLL: OnVScroll(hwnd, static_cast<UINT>(LOWORD(wp))); return 0;
        case WM_MOUSEWHEEL:
            return OnMouseWheel(hwnd,
                                GET_WHEEL_DELTA_WPARAM(wp),
                                GET_KEYSTATE_WPARAM(wp),
                                static_cast<int>(static_cast<short>(LOWORD(lp))),
                                static_cast<int>(static_cast<short>(HIWORD(lp))));
        case WM_DPICHANGED: OnDpiChanged(hwnd, HIWORD(wp), reinterpret_cast<const RECT*>(lp)); return 0;
        case WM_INPUTLANGCHANGE: return OnInputLangChange(hwnd, wp, lp);
        case WM_PAINT: OnPaint(); return 0;
        case WM_ERASEBKGND: return OnEraseBkgnd(hwnd, reinterpret_cast<HDC>(wp));
        case kAsyncProgressMessage: OnAsyncProgress(static_cast<int>(wp), static_cast<int>(lp)); return 0;
        case kAsyncOpenCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncOpenResult>(lp);
            OnAsyncOpenComplete(std::move(result));
            return 0;
        }
        case kAsyncExportCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncExportResult>(lp);
            OnAsyncExportComplete(std::move(result));
            return 0;
        }
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_NCACTIVATE:
        {
            const bool windowActive = wp != FALSE;
            OnNcActivate(windowActive);
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case WM_NCDESTROY: return OnNcDestroy(hwnd, wp, lp);
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }
}

LRESULT ViewerImgRaw::OnEraseBkgnd(HWND hwnd, HDC hdc) noexcept
{
    if (_allowEraseBkgnd)
    {
        return DefWindowProcW(hwnd, WM_ERASEBKGND, reinterpret_cast<WPARAM>(hdc), 0);
    }
    return 1;
}

LRESULT ViewerImgRaw::OnMouseWheel(HWND hwnd, short delta, UINT keyState, int x, int y) noexcept
{
    if (! hwnd)
    {
        return 0;
    }

    if (_transientZoomActive || delta == 0 || ! HasDisplayImage() || ! _currentImage)
    {
        return 0;
    }

    const bool ctrl  = (keyState & MK_CONTROL) != 0;
    const bool shift = (keyState & MK_SHIFT) != 0;

    if (ctrl)
    {
        const int detents = static_cast<int>(delta / WHEEL_DELTA);
        if (detents != 0)
        {
            if (shift)
            {
                _contrast = std::clamp(_contrast + static_cast<float>(detents) * 0.05f, 0.10f, 3.00f);
            }
            else
            {
                _brightness = std::clamp(_brightness + static_cast<float>(detents) * 0.05f, -1.0f, 1.0f);
            }

            _imageBitmap.reset();
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
        }
        return 0;
    }

    POINT pt{x, y};
    if (ScreenToClient(hwnd, &pt) == FALSE)
    {
        return 0;
    }

    float displayedZoom = 1.0f;
    float drawX         = 0.0f;
    float drawY         = 0.0f;
    float drawW         = 0.0f;
    float drawH         = 0.0f;
    if (! ComputeImageLayoutPx(displayedZoom, drawX, drawY, drawW, drawH) || displayedZoom <= 0.0f)
    {
        return 0;
    }

    constexpr float kWheelStepPerDetent = 1.08f; // Smooth zoom (8% per wheel detent)
    constexpr float kMinWheelZoom       = 0.05f;
    constexpr float kMaxWheelZoom       = 16.0f;

    const float detents = static_cast<float>(delta) / static_cast<float>(WHEEL_DELTA);
    const float factor  = std::pow(kWheelStepPerDetent, detents);
    const float newZoom = std::clamp(displayedZoom * factor, kMinWheelZoom, kMaxWheelZoom);

    ApplyZoom(hwnd, newZoom, pt);
    return 0;
}

void ViewerImgRaw::OnHScroll(HWND hwnd, UINT code) noexcept
{
    if (! hwnd || _fitToWindow || _transientZoomActive || ! HasDisplayImage() || ! _hScrollVisible)
    {
        return;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    if (GetScrollInfo(hwnd, SB_HORZ, &si) == 0)
    {
        return;
    }

    const int page   = std::max(1, static_cast<int>(si.nPage));
    const int maxPos = std::max(si.nMin, si.nMax - page + 1);
    int newPos       = si.nPos;

    const UINT dpi = GetDpiForWindow(hwnd);
    const int line = std::max(1, PxFromDip(40, dpi));

    switch (code)
    {
        case SB_LINELEFT: newPos -= line; break;
        case SB_LINERIGHT: newPos += line; break;
        case SB_PAGELEFT: newPos -= page; break;
        case SB_PAGERIGHT: newPos += page; break;
        case SB_LEFT: newPos = si.nMin; break;
        case SB_RIGHT: newPos = maxPos; break;
        case SB_THUMBPOSITION:
        case SB_THUMBTRACK: newPos = si.nTrackPos; break;
        default: return;
    }

    newPos = std::clamp(newPos, si.nMin, maxPos);
    if (newPos == si.nPos)
    {
        return;
    }

    float zoom  = 1.0f;
    float drawX = 0.0f;
    float drawY = 0.0f;
    float drawW = 0.0f;
    float drawH = 0.0f;
    if (! ComputeImageLayoutPx(zoom, drawX, drawY, drawW, drawH) || zoom <= 0.0f)
    {
        return;
    }

    const float contentW = static_cast<float>(std::max(0L, _contentRect.right - _contentRect.left));
    const float delta    = drawW - contentW;
    if (delta <= 0.0f)
    {
        return;
    }

    _panOffsetXPx = delta * 0.5f - static_cast<float>(newPos);
    _panning      = false;

    float tmpZoom = 1.0f;
    float tmpX    = 0.0f;
    float tmpY    = 0.0f;
    float tmpW    = 0.0f;
    float tmpH    = 0.0f;
    static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));

    UpdateScrollBars(hwnd);
    InvalidateRect(hwnd, &_contentRect, FALSE);
    InvalidateRect(hwnd, &_statusRect, FALSE);
}

void ViewerImgRaw::OnVScroll(HWND hwnd, UINT code) noexcept
{
    if (! hwnd || _fitToWindow || _transientZoomActive || ! HasDisplayImage() || ! _vScrollVisible)
    {
        return;
    }

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_ALL;
    if (GetScrollInfo(hwnd, SB_VERT, &si) == 0)
    {
        return;
    }

    const int page   = std::max(1, static_cast<int>(si.nPage));
    const int maxPos = std::max(si.nMin, si.nMax - page + 1);
    int newPos       = si.nPos;

    const UINT dpi = GetDpiForWindow(hwnd);
    const int line = std::max(1, PxFromDip(40, dpi));

    switch (code)
    {
        case SB_LINEUP: newPos -= line; break;
        case SB_LINEDOWN: newPos += line; break;
        case SB_PAGEUP: newPos -= page; break;
        case SB_PAGEDOWN: newPos += page; break;
        case SB_TOP: newPos = si.nMin; break;
        case SB_BOTTOM: newPos = maxPos; break;
        case SB_THUMBPOSITION:
        case SB_THUMBTRACK: newPos = si.nTrackPos; break;
        default: return;
    }

    newPos = std::clamp(newPos, si.nMin, maxPos);
    if (newPos == si.nPos)
    {
        return;
    }

    float zoom  = 1.0f;
    float drawX = 0.0f;
    float drawY = 0.0f;
    float drawW = 0.0f;
    float drawH = 0.0f;
    if (! ComputeImageLayoutPx(zoom, drawX, drawY, drawW, drawH) || zoom <= 0.0f)
    {
        return;
    }

    const float contentH = static_cast<float>(std::max(0L, _contentRect.bottom - _contentRect.top));
    const float delta    = drawH - contentH;
    if (delta <= 0.0f)
    {
        return;
    }

    _panOffsetYPx = delta * 0.5f - static_cast<float>(newPos);
    _panning      = false;

    float tmpZoom = 1.0f;
    float tmpX    = 0.0f;
    float tmpY    = 0.0f;
    float tmpW    = 0.0f;
    float tmpH    = 0.0f;
    static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));

    UpdateScrollBars(hwnd);
    InvalidateRect(hwnd, &_contentRect, FALSE);
    InvalidateRect(hwnd, &_statusRect, FALSE);
}

void ViewerImgRaw::ApplyZoom(HWND hwnd, float newZoom, std::optional<POINT> anchorClientPt) noexcept
{
    if (! hwnd || _transientZoomActive || ! HasDisplayImage() || ! _currentImage)
    {
        return;
    }

    float displayedZoom = 1.0f;
    float drawX         = 0.0f;
    float drawY         = 0.0f;
    float drawW         = 0.0f;
    float drawH         = 0.0f;
    if (! ComputeImageLayoutPx(displayedZoom, drawX, drawY, drawW, drawH) || displayedZoom <= 0.0f)
    {
        return;
    }

    const CachedImage* image = _currentImage;
    const uint32_t imgWPx    = IsDisplayingThumbnail() ? image->thumbWidth : image->rawWidth;
    const uint32_t imgHPx    = IsDisplayingThumbnail() ? image->thumbHeight : image->rawHeight;
    if (imgWPx == 0 || imgHPx == 0)
    {
        return;
    }

    const uint16_t orientation = (_viewOrientation >= 1 && _viewOrientation <= 8) ? _viewOrientation : static_cast<uint16_t>(1);
    const bool swapAxes        = (orientation >= 5 && orientation <= 8);

    const float contentW = static_cast<float>(std::max(0L, _contentRect.right - _contentRect.left));
    const float contentH = static_cast<float>(std::max(0L, _contentRect.bottom - _contentRect.top));
    const float imgW     = static_cast<float>(swapAxes ? imgHPx : imgWPx);
    const float imgH     = static_cast<float>(swapAxes ? imgWPx : imgHPx);
    if (contentW <= 0.0f || contentH <= 0.0f || imgW <= 0.0f || imgH <= 0.0f)
    {
        return;
    }

    newZoom = std::clamp(newZoom, 0.01f, 64.0f);

    float anchorClientX = 0.0f;
    float anchorClientY = 0.0f;
    float anchorImgX    = 0.0f;
    float anchorImgY    = 0.0f;

    if (anchorClientPt)
    {
        anchorClientX = static_cast<float>(anchorClientPt->x);
        anchorClientY = static_cast<float>(anchorClientPt->y);
    }
    else
    {
        anchorClientX = static_cast<float>(_contentRect.left) + contentW * 0.5f;
        anchorClientY = static_cast<float>(_contentRect.top) + contentH * 0.5f;
    }

    const bool overImage = (anchorClientX >= drawX && anchorClientX < (drawX + drawW) && anchorClientY >= drawY && anchorClientY < (drawY + drawH));
    if (overImage)
    {
        anchorImgX = (anchorClientX - drawX) / displayedZoom;
        anchorImgY = (anchorClientY - drawY) / displayedZoom;
    }
    else
    {
        anchorClientX = static_cast<float>(_contentRect.left) + contentW * 0.5f;
        anchorClientY = static_cast<float>(_contentRect.top) + contentH * 0.5f;
        anchorImgX    = imgW * 0.5f;
        anchorImgY    = imgH * 0.5f;
    }

    const float newDrawW = imgW * newZoom;
    const float newDrawH = imgH * newZoom;
    const float baseX    = static_cast<float>(_contentRect.left) + (contentW - newDrawW) * 0.5f;
    const float baseY    = static_cast<float>(_contentRect.top) + (contentH - newDrawH) * 0.5f;

    const float desiredX = anchorClientX - anchorImgX * newZoom;
    const float desiredY = anchorClientY - anchorImgY * newZoom;

    _fitToWindow  = false;
    _manualZoom   = newZoom;
    _panOffsetXPx = desiredX - baseX;
    _panOffsetYPx = desiredY - baseY;

    float tmpZoom = 1.0f;
    float tmpX    = 0.0f;
    float tmpY    = 0.0f;
    float tmpW    = 0.0f;
    float tmpH    = 0.0f;
    static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));

    UpdateMenuChecks(hwnd);
    UpdateScrollBars(hwnd);
    InvalidateRect(hwnd, &_contentRect, FALSE);
    InvalidateRect(hwnd, &_statusRect, FALSE);
}

void ViewerImgRaw::OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggested) noexcept
{
    if (! hwnd || newDpi == 0)
    {
        return;
    }

    if (suggested)
    {
        const int width  = std::max(1, static_cast<int>(suggested->right - suggested->left));
        const int height = std::max(1, static_cast<int>(suggested->bottom - suggested->top));
        SetWindowPos(hwnd, nullptr, suggested->left, suggested->top, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
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

    if (_hFileCombo && _uiFont)
    {
        SendMessageW(_hFileCombo.get(), WM_SETFONT, reinterpret_cast<WPARAM>(_uiFont.get()), TRUE);
    }

    if (_hFileCombo)
    {
        int itemHeight = PxFromDip(24, newDpi);
        auto hdc       = wil::GetDC(hwnd);
        if (hdc)
        {
            HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            auto oldFont    = wil::SelectObject(hdc.get(), fontToUse);
            static_cast<void>(oldFont);

            TEXTMETRICW tm{};
            if (GetTextMetricsW(hdc.get(), &tm) != 0)
            {
                itemHeight = tm.tmHeight + tm.tmExternalLeading + PxFromDip(6, newDpi);
            }
        }

        itemHeight = std::max(itemHeight, 1);
        SendMessageW(_hFileCombo.get(), CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), static_cast<LPARAM>(itemHeight));
        SendMessageW(_hFileCombo.get(), CB_SETITEMHEIGHT, 0, static_cast<LPARAM>(itemHeight));
    }

    DiscardDirect2D();
    Layout(hwnd);
    InvalidateRect(hwnd, nullptr, TRUE);
}

void ViewerImgRaw::OnNcActivate(bool windowActive) noexcept
{
    ApplyTitleBarTheme(windowActive);
}

LRESULT ViewerImgRaw::OnInputLangChange(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    const LRESULT result = DefWindowProcW(hwnd, WM_INPUTLANGCHANGE, wp, lp);
    UpdateMenuShortcutTextForKeyboardLayout();
    DrawMenuBar(hwnd);
    return result;
}

LRESULT ViewerImgRaw::OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    OnDestroy();
    static_cast<void>(DrainPostedPayloadsForWindow(hwnd));

    _hFileCombo.release();
    _hWnd.release();
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);

    Release();
    return DefWindowProcW(hwnd, WM_NCDESTROY, wp, lp);
}

void ViewerImgRaw::OnCreate(HWND hwnd)
{
    const UINT dpi       = GetDpiForWindow(hwnd);
    const int uiHeightPx = -MulDiv(9, static_cast<int>(dpi), 72);

    _allowEraseBkgnd = true;

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
        Debug::ErrorWithLastError(L"ViewerImgRaw: CreateFontW failed for UI font.");
    }

    const DWORD comboStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS;
    _hFileCombo.reset(CreateWindowExW(
        0, L"COMBOBOX", nullptr, comboStyle, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_VIEWERRAW_FILE_COMBO)), g_hInstance, nullptr));
    if (! _hFileCombo)
    {
        Debug::ErrorWithLastError(L"ViewerImgRaw: CreateWindowExW failed for file combo.");
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
}

void ViewerImgRaw::OnDestroy()
{
    EndLoadingUi();
    DiscardDirect2D();
    ClearImageCache();

    IViewerCallback* callback = _callback;
    void* cookie              = _callbackCookie;
    if (callback)
    {
        AddRef();
        static_cast<void>(callback->ViewerClosed(cookie));
        Release();
    }
}

void ViewerImgRaw::OnTimer(UINT_PTR timerId) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    if (timerId == kLoadingDelayTimerId)
    {
        KillTimer(_hWnd.get(), kLoadingDelayTimerId);
        if (! _isLoading)
        {
            return;
        }

        _showLoadingOverlay       = true;
        _loadingSpinnerAngleDeg   = 0.0f;
        _loadingSpinnerLastTickMs = GetTickCount64();
        SetTimer(_hWnd.get(), kLoadingAnimTimerId, kLoadingAnimIntervalMs, nullptr);

        InvalidateRect(_hWnd.get(), &_contentRect, FALSE);
        return;
    }

    if (timerId == kLoadingAnimTimerId)
    {
        UpdateLoadingSpinner();
        return;
    }
}

void ViewerImgRaw::OnSize(UINT width, UINT height)
{
    if (width == 0 || height == 0)
    {
        return;
    }

    if (_d2dTarget)
    {
        const HRESULT hr = _d2dTarget->Resize(D2D1::SizeU(width, height));
        if (hr == D2DERR_RECREATE_TARGET)
        {
            DiscardDirect2D();
        }
    }

    if (_hWnd)
    {
        Layout(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerImgRaw::Layout(HWND hwnd) noexcept
{
    ComputeLayoutRects(hwnd);
    UpdateScrollBars(hwnd);
}

void ViewerImgRaw::ComputeLayoutRects(HWND hwnd) noexcept
{
    RECT client{};
    GetClientRect(hwnd, &client);

    const UINT dpi = GetDpiForWindow(hwnd);
    if (_vScrollVisible)
    {
        const int scrollW = GetSystemMetricsForDpi(SM_CXVSCROLL, dpi);
        client.right      = std::max(client.left, client.right - scrollW);
    }
    if (_hScrollVisible)
    {
        const int scrollH = GetSystemMetricsForDpi(SM_CYHSCROLL, dpi);
        client.bottom     = std::max(client.top, client.bottom - scrollH);
    }

    const int headerH = PxFromDip(kHeaderHeightDip, dpi);
    const int statusH = PxFromDip(kStatusHeightDip, dpi);

    _headerRect        = client;
    _headerRect.bottom = std::min(client.bottom, client.top + headerH);

    _statusRect     = client;
    _statusRect.top = std::max(client.top, client.bottom - statusH);

    _contentRect        = client;
    _contentRect.top    = _headerRect.bottom;
    _contentRect.bottom = _statusRect.top;

    if (_hFileCombo)
    {
        const bool showCombo = _otherItems.size() > 1;
        ShowWindow(_hFileCombo.get(), showCombo ? SW_SHOW : SW_HIDE);
        if (showCombo)
        {
            const int padding = PxFromDip(6, dpi);
            const int x       = _headerRect.left + padding;
            const int y       = _headerRect.top + padding / 2;
            const int w       = std::max(1L, (_headerRect.right - _headerRect.left) - 2 * padding);
            const int h       = std::max(1L, (_headerRect.bottom - _headerRect.top) - padding);
            SetWindowPos(_hFileCombo.get(), nullptr, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }
}

void ViewerImgRaw::UpdateScrollBars(HWND hwnd) noexcept
{
    if (! hwnd || _updatingScrollBars)
    {
        return;
    }

    _updatingScrollBars  = true;
    auto restoreUpdating = wil::scope_exit([&] { _updatingScrollBars = false; });

    auto computeDesiredVisibility = [&]() noexcept -> std::pair<bool, bool>
    {
        bool showH = false;
        bool showV = false;

        if (_fitToWindow || ! HasDisplayImage() || ! _currentImage)
        {
            return {false, false};
        }

        float zoom  = 1.0f;
        float x     = 0.0f;
        float y     = 0.0f;
        float drawW = 0.0f;
        float drawH = 0.0f;
        if (! ComputeImageLayoutPx(zoom, x, y, drawW, drawH))
        {
            return {false, false};
        }

        const float contentW = static_cast<float>(std::max(0L, _contentRect.right - _contentRect.left));
        const float contentH = static_cast<float>(std::max(0L, _contentRect.bottom - _contentRect.top));
        showH                = drawW > (contentW + 0.5f);
        showV                = drawH > (contentH + 0.5f);
        return {showH, showV};
    };

    auto applyVisibility = [&](bool showH, bool showV) noexcept -> bool
    {
        bool changed = false;

        if (showH != _hScrollVisible)
        {
            ShowScrollBar(hwnd, SB_HORZ, showH ? TRUE : FALSE);
            _hScrollVisible = showH;
            changed         = true;
        }

        if (showV != _vScrollVisible)
        {
            ShowScrollBar(hwnd, SB_VERT, showV ? TRUE : FALSE);
            _vScrollVisible = showV;
            changed         = true;
        }

        return changed;
    };

    auto [wantH, wantV] = computeDesiredVisibility();
    if (applyVisibility(wantH, wantV))
    {
        ComputeLayoutRects(hwnd);
        std::tie(wantH, wantV) = computeDesiredVisibility();
        if (applyVisibility(wantH, wantV))
        {
            ComputeLayoutRects(hwnd);
        }
    }

    float zoom  = 1.0f;
    float drawX = 0.0f;
    float drawY = 0.0f;
    float drawW = 0.0f;
    float drawH = 0.0f;
    static_cast<void>(ComputeImageLayoutPx(zoom, drawX, drawY, drawW, drawH));

    const int contentWPx = std::max(0L, _contentRect.right - _contentRect.left);
    const int contentHPx = std::max(0L, _contentRect.bottom - _contentRect.top);

    if (_hScrollVisible)
    {
        const int scaledW = std::max(1, static_cast<int>(std::lround(drawW)));
        const int page    = std::max(1, contentWPx);
        const int maxPos  = std::max(0, scaledW - page);

        int pos = static_cast<int>(std::lround(static_cast<float>(_contentRect.left) - drawX));
        pos     = std::clamp(pos, 0, maxPos);

        SCROLLINFO si{};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
        si.nMin   = 0;
        si.nMax   = scaledW - 1;
        si.nPage  = static_cast<UINT>(page);
        si.nPos   = pos;
        SetScrollInfo(hwnd, SB_HORZ, &si, TRUE);
    }
    else
    {
        SCROLLINFO si{};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
        si.nMin   = 0;
        si.nMax   = 0;
        si.nPage  = 0;
        si.nPos   = 0;
        SetScrollInfo(hwnd, SB_HORZ, &si, TRUE);
    }

    if (_vScrollVisible)
    {
        const int scaledH = std::max(1, static_cast<int>(std::lround(drawH)));
        const int page    = std::max(1, contentHPx);
        const int maxPos  = std::max(0, scaledH - page);

        int pos = static_cast<int>(std::lround(static_cast<float>(_contentRect.top) - drawY));
        pos     = std::clamp(pos, 0, maxPos);

        SCROLLINFO si{};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
        si.nMin   = 0;
        si.nMax   = scaledH - 1;
        si.nPage  = static_cast<UINT>(page);
        si.nPos   = pos;
        SetScrollInfo(hwnd, SB_VERT, &si, TRUE);
    }
    else
    {
        SCROLLINFO si{};
        si.cbSize = sizeof(si);
        si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
        si.nMin   = 0;
        si.nMax   = 0;
        si.nPage  = 0;
        si.nPos   = 0;
        SetScrollInfo(hwnd, SB_VERT, &si, TRUE);
    }
}

void ViewerImgRaw::ApplyTheme(HWND hwnd) noexcept
{
    const bool useDarkMode  = (_hasTheme && _theme.darkMode && ! _theme.highContrast);
    const wchar_t* winTheme = useDarkMode ? L"DarkMode_Explorer" : L"Explorer";

    _uiBg       = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    _uiText     = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
    _uiHeaderBg = _uiBg;
    _uiStatusBg = _uiBg;

    if (_hasTheme && ! _theme.highContrast)
    {
        const COLORREF accent           = ResolveAccentColor(_theme, L"header");
        static constexpr uint8_t kAlpha = 22u;
        _uiHeaderBg                     = BlendColor(_uiBg, accent, kAlpha);
        _uiStatusBg                     = BlendColor(_uiBg, accent, kAlpha);
    }

    _menuHeaderBrush.reset(CreateSolidBrush(_uiHeaderBg));
    if (! _menuHeaderBrush)
    {
        Debug::Warning(L"ViewerImgRaw: CreateSolidBrush failed for menu header brush.");
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

    UpdateMenuChecks(hwnd);
    ApplyMenuTheme(hwnd);
    ApplyTitleBarTheme(true);
    ApplyPendingViewerImgRawClassBackgroundBrush(hwnd);
}

void ViewerImgRaw::ApplyTitleBarTheme(bool windowActive) noexcept
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

void ViewerImgRaw::UpdateMenuChecks(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const HMENU menu = GetMenu(hwnd);
    if (! menu)
    {
        return;
    }

    CheckMenuItem(menu, IDM_VIEWERRAW_VIEW_FIT, static_cast<UINT>(MF_BYCOMMAND | (_fitToWindow ? MF_CHECKED : MF_UNCHECKED)));
    CheckMenuItem(menu,
                  IDM_VIEWERRAW_VIEW_ACTUAL_SIZE,
                  static_cast<UINT>(MF_BYCOMMAND | ((! _fitToWindow && std::fabs(_manualZoom - 1.0f) < 0.001f) ? MF_CHECKED : MF_UNCHECKED)));

    bool thumbSelectable     = ! _currentSidecarJpegPath.empty();
    const CachedImage* image = _currentImage;
    if (image && _currentImageKey == _currentPath)
    {
        thumbSelectable = thumbSelectable || image->thumbAvailable;
    }

    EnableMenuItem(menu, IDM_VIEWERRAW_VIEW_SOURCE_THUMBNAIL, static_cast<UINT>(MF_BYCOMMAND | (thumbSelectable ? MF_ENABLED : MF_GRAYED)));

    const DisplayMode effectiveSelection = thumbSelectable ? _displayMode : DisplayMode::Raw;
    const UINT selectedSource            = (effectiveSelection == DisplayMode::Thumbnail) ? IDM_VIEWERRAW_VIEW_SOURCE_THUMBNAIL : IDM_VIEWERRAW_VIEW_SOURCE_RAW;
    CheckMenuRadioItem(menu, IDM_VIEWERRAW_VIEW_SOURCE_RAW, IDM_VIEWERRAW_VIEW_SOURCE_THUMBNAIL, selectedSource, MF_BYCOMMAND);

    CheckMenuItem(menu, IDM_VIEWERRAW_VIEW_SHOW_EXIF_OVERLAY, static_cast<UINT>(MF_BYCOMMAND | (_showExifOverlay ? MF_CHECKED : MF_UNCHECKED)));

    CheckMenuItem(menu, IDM_VIEWERRAW_VIEW_TOGGLE_GRAYSCALE, static_cast<UINT>(MF_BYCOMMAND | (_grayscale ? MF_CHECKED : MF_UNCHECKED)));
    CheckMenuItem(menu, IDM_VIEWERRAW_VIEW_TOGGLE_NEGATIVE, static_cast<UINT>(MF_BYCOMMAND | (_negative ? MF_CHECKED : MF_UNCHECKED)));
}

void ViewerImgRaw::ApplyMenuTheme(HWND hwnd) noexcept
{
    HMENU menu = hwnd ? GetMenu(hwnd) : nullptr;
    if (! menu)
    {
        return;
    }

    if (_menuHeaderBrush)
    {
        MENUINFO mi{};
        mi.cbSize  = sizeof(mi);
        mi.fMask   = MIM_BACKGROUND | MIM_APPLYTOSUBMENUS;
        mi.hbrBack = _menuHeaderBrush.get();
        SetMenuInfo(menu, &mi);
    }

    _menuThemeItems.clear();
    PrepareMenuTheme(menu, true, _menuThemeItems);
    UpdateMenuShortcutTextForKeyboardLayout();
    DrawMenuBar(hwnd);
}

void ViewerImgRaw::UpdateMenuShortcutTextForKeyboardLayout() noexcept
{
    const HKL keyboardLayout = GetKeyboardLayout(0);
    if (! keyboardLayout)
    {
        return;
    }

    const std::wstring zoomInKey    = KeyGlyphFromVirtualKey(VK_OEM_PLUS, keyboardLayout);
    const std::wstring zoomOutKey   = KeyGlyphFromVirtualKey(VK_OEM_MINUS, keyboardLayout);
    const std::wstring zoomResetKey = KeyGlyphFromVirtualKey(static_cast<UINT>('0'), keyboardLayout);

    for (MenuItemData& item : _menuThemeItems)
    {
        switch (item.id)
        {
            case IDM_VIEWERRAW_VIEW_ZOOM_IN:
                if (! zoomInKey.empty())
                {
                    item.shortcut = zoomInKey;
                }
                break;
            case IDM_VIEWERRAW_VIEW_ZOOM_OUT:
                if (! zoomOutKey.empty())
                {
                    item.shortcut = zoomOutKey;
                }
                break;
            case IDM_VIEWERRAW_VIEW_ZOOM_RESET:
                if (! zoomResetKey.empty())
                {
                    item.shortcut = zoomResetKey;
                }
                break;
            default: break;
        }
    }
}

void ViewerImgRaw::PrepareMenuTheme(HMENU menu, bool topLevel, std::vector<MenuItemData>& outItems) noexcept
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

void ViewerImgRaw::OnMeasureMenuItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept
{
    if (! measure || measure->CtlType != ODT_MENU)
    {
        return;
    }

    const size_t index = static_cast<size_t>(measure->itemData);
    if (index >= _menuThemeItems.size())
    {
        return;
    }

    const MenuItemData& data = _menuThemeItems[index];
    const UINT dpi           = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;

    if (data.separator)
    {
        measure->itemWidth  = 1;
        measure->itemHeight = static_cast<UINT>(MulDiv(8, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));
        return;
    }

    const UINT heightDip = data.topLevel ? 20u : 24u;
    measure->itemHeight  = static_cast<UINT>(MulDiv(static_cast<int>(heightDip), static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI));

    auto hdc = wil::GetDC(hwnd);
    if (! hdc)
    {
        measure->itemWidth = 120;
        return;
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
}

void ViewerImgRaw::OnDrawMenuItem(DRAWITEMSTRUCT* draw) noexcept
{
    if (! draw || draw->CtlType != ODT_MENU || ! draw->hDC)
    {
        return;
    }

    const size_t index = static_cast<size_t>(draw->itemData);
    if (index >= _menuThemeItems.size())
    {
        return;
    }

    const MenuItemData& data = _menuThemeItems[index];
    const bool selected      = (draw->itemState & ODS_SELECTED) != 0;
    const bool disabled      = (draw->itemState & ODS_DISABLED) != 0;
    const bool checked       = (draw->itemState & ODS_CHECKED) != 0;

    const COLORREF bg             = _hasTheme ? (data.topLevel ? _uiHeaderBg : ColorRefFromArgb(_theme.backgroundArgb)) : GetSysColor(COLOR_MENU);
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

    RECT itemRect = draw->rcItem;
    wil::unique_any<HRGN, decltype(&::DeleteObject), ::DeleteObject> clipRgn(CreateRectRgnIndirect(&itemRect));
    if (clipRgn)
    {
        SelectClipRgn(draw->hDC, clipRgn.get());
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
        return;
    }

    HFONT fontToUse = _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    auto oldFont    = wil::SelectObject(draw->hDC, fontToUse);
    static_cast<void>(oldFont);

    const int dpi              = GetDeviceCaps(draw->hDC, LOGPIXELSX);
    const bool iconFontValid   = EnsureViewerImgRawMenuIconFont(draw->hDC, static_cast<UINT>(dpi));
    const int paddingX         = MulDiv(8, dpi, USER_DEFAULT_SCREEN_DPI);
    const int checkAreaWidth   = data.topLevel ? 0 : MulDiv(20, dpi, USER_DEFAULT_SCREEN_DPI);
    const int subMenuAreaWidth = data.hasSubMenu && ! data.topLevel ? MulDiv(18, dpi, USER_DEFAULT_SCREEN_DPI) : 0;
    const int checkGap         = data.topLevel ? 0 : MulDiv(4, dpi, USER_DEFAULT_SCREEN_DPI);

    RECT textRect = draw->rcItem;
    textRect.left += paddingX + checkAreaWidth + checkGap;
    textRect.right -= paddingX + subMenuAreaWidth;
    RECT shortcutRect = textRect;

    SetBkMode(draw->hDC, TRANSPARENT);
    SetTextColor(draw->hDC, textColor);

    if (checked && checkAreaWidth > 0)
    {
        RECT checkRect = draw->rcItem;
        checkRect.left += paddingX;
        checkRect.right     = checkRect.left + checkAreaWidth;
        const bool useIcons = iconFontValid && g_viewerImgRawMenuIconFont;
        const wchar_t glyph = useIcons ? FluentIcons::kCheckMark : FluentIcons::kFallbackCheckMark;
        wchar_t glyphText[2]{glyph, 0};

        HFONT glyphFont  = useIcons ? g_viewerImgRawMenuIconFont.get() : fontToUse;
        auto oldIconFont = wil::SelectObject(draw->hDC, glyphFont);
        DrawTextW(draw->hDC, glyphText, 1, &checkRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    const UINT drawFlags = DT_VCENTER | DT_SINGLELINE | DT_HIDEPREFIX;

    if (! data.text.empty())
    {
        DrawTextW(draw->hDC, data.text.c_str(), static_cast<int>(data.text.size()), &textRect, DT_LEFT | drawFlags);
    }

    if (! data.shortcut.empty())
    {
        DrawTextW(draw->hDC, data.shortcut.c_str(), static_cast<int>(data.shortcut.size()), &shortcutRect, DT_RIGHT | drawFlags);
    }

    if (data.hasSubMenu && ! data.topLevel)
    {
        RECT arrowRect = draw->rcItem;
        arrowRect.right -= paddingX;
        arrowRect.left = std::max(arrowRect.left, arrowRect.right - subMenuAreaWidth);

        const bool useIcons = iconFontValid && g_viewerImgRawMenuIconFont;
        const wchar_t glyph = useIcons ? FluentIcons::kChevronRightSmall : FluentIcons::kFallbackChevronRight;
        wchar_t glyphText[2]{glyph, 0};

        COLORREF arrowColor = textColor;
        if (! selected && ! disabled)
        {
            arrowColor = BlendColor(fillColor, textColor, 120u);
        }

        SetTextColor(draw->hDC, arrowColor);
        HFONT arrowFont  = useIcons ? g_viewerImgRawMenuIconFont.get() : fontToUse;
        auto oldIconFont = wil::SelectObject(draw->hDC, arrowFont);
        DrawTextW(draw->hDC, glyphText, 1, &arrowRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

        const int arrowExcludeWidth = std::max(subMenuAreaWidth, GetSystemMetricsForDpi(SM_CXMENUCHECK, static_cast<UINT>(dpi)));
        RECT arrowExcludeRect       = itemRect;
        arrowExcludeRect.left       = std::max(arrowExcludeRect.left, arrowExcludeRect.right - arrowExcludeWidth);
        ExcludeClipRect(draw->hDC, arrowExcludeRect.left, arrowExcludeRect.top, arrowExcludeRect.right, arrowExcludeRect.bottom);
    }
}

LRESULT ViewerImgRaw::OnMeasureItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept
{
    if (! measure)
    {
        return FALSE;
    }

    if (measure->CtlType == ODT_COMBOBOX && measure->CtlID == IDC_VIEWERRAW_FILE_COMBO)
    {
        const UINT dpi      = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;
        measure->itemHeight = static_cast<UINT>(std::max(1, PxFromDip(24, dpi)));
        return TRUE;
    }

    if (measure->CtlType == ODT_MENU)
    {
        OnMeasureMenuItem(hwnd, measure);
        return TRUE;
    }

    return FALSE;
}

LRESULT ViewerImgRaw::OnDrawItem(HWND hwnd, DRAWITEMSTRUCT* draw) noexcept
{
    if (! draw)
    {
        return FALSE;
    }

    if (draw->CtlType == ODT_COMBOBOX && draw->CtlID == IDC_VIEWERRAW_FILE_COMBO)
    {
        const HDC hdc = draw->hDC;
        RECT rc       = draw->rcItem;

        const bool selected = (draw->itemState & ODS_SELECTED) != 0;
        const bool disabled = (draw->itemState & ODS_DISABLED) != 0;

        COLORREF bg   = _uiHeaderBg;
        COLORREF text = _uiText;
        if (selected)
        {
            bg   = (_hasTheme && ! _theme.highContrast) ? ResolveAccentColor(_theme, L"combo") : GetSysColor(COLOR_HIGHLIGHT);
            text = (_hasTheme && ! _theme.highContrast) ? ContrastingTextColor(bg) : GetSysColor(COLOR_HIGHLIGHTTEXT);
        }
        if (disabled)
        {
            text = GetSysColor(COLOR_GRAYTEXT);
        }

        SetBkColor(hdc, bg);
        SetTextColor(hdc, text);
        SetDCBrushColor(hdc, bg);
        FillRect(hdc, &rc, static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));

        auto oldFont = wil::SelectObject(hdc, _uiFont ? _uiFont.get() : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
        static_cast<void>(oldFont);

        const UINT dpi = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;
        rc.left += PxFromDip(6, dpi);
        rc.right = std::max(rc.left, rc.right - PxFromDip(2, dpi));

        const UINT itemId = draw->itemID;
        if (itemId != static_cast<UINT>(-1) && itemId < _otherItems.size())
        {
            const OtherItem& item              = _otherItems[itemId];
            const std::wstring_view textToDraw = ! item.label.empty() ? std::wstring_view(item.label) : std::wstring_view(item.primaryPath);
            DrawTextW(hdc, textToDraw.data(), static_cast<int>(textToDraw.size()), &rc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
        }

        if ((draw->itemState & ODS_FOCUS) != 0)
        {
            DrawFocusRect(hdc, &draw->rcItem);
        }

        return TRUE;
    }

    if (draw->CtlType == ODT_MENU)
    {
        OnDrawMenuItem(draw);
        return TRUE;
    }

    return FALSE;
}

void ViewerImgRaw::BeginLoadingUi() noexcept
{
    if (_hostAlerts)
    {
        static_cast<void>(_hostAlerts->ClearAlert(HOST_ALERT_SCOPE_WINDOW, nullptr));
    }
    _alertVisible = false;

    _isLoading                = true;
    _showLoadingOverlay       = false;
    _loadingSpinnerAngleDeg   = 0.0f;
    _loadingSpinnerLastTickMs = GetTickCount64();

    if (! _hWnd)
    {
        return;
    }

    KillTimer(_hWnd.get(), kLoadingDelayTimerId);
    KillTimer(_hWnd.get(), kLoadingAnimTimerId);
    SetTimer(_hWnd.get(), kLoadingDelayTimerId, kLoadingDelayMs, nullptr);
}

void ViewerImgRaw::EndLoadingUi() noexcept
{
    if (_hWnd)
    {
        KillTimer(_hWnd.get(), kLoadingDelayTimerId);
        KillTimer(_hWnd.get(), kLoadingAnimTimerId);
    }

    _isLoading          = false;
    _showLoadingOverlay = false;
}

void ViewerImgRaw::UpdateLoadingSpinner() noexcept
{
    if (! _isLoading || ! _showLoadingOverlay)
    {
        return;
    }

    const ULONGLONG now       = GetTickCount64();
    const ULONGLONG last      = _loadingSpinnerLastTickMs;
    _loadingSpinnerLastTickMs = now;

    double deltaSec = 0.0;
    if (now > last)
    {
        deltaSec = static_cast<double>(now - last) / 1000.0;
    }

    _loadingSpinnerAngleDeg += static_cast<float>(deltaSec * static_cast<double>(kLoadingSpinnerDegPerSec));
    while (_loadingSpinnerAngleDeg >= 360.0f)
    {
        _loadingSpinnerAngleDeg -= 360.0f;
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_contentRect, FALSE);
    }
}

void ViewerImgRaw::DrawLoadingOverlay(ID2D1HwndRenderTarget* target, ID2D1SolidColorBrush* brush) noexcept
{
    if (! _isLoading || ! _showLoadingOverlay || ! target || ! brush || ! _hWnd)
    {
        return;
    }

    const UINT dpi            = GetDpiForWindow(_hWnd.get());
    const D2D1_RECT_F content = RectFFromPixels(_contentRect, dpi);

    const float widthDip  = std::max(0.0f, content.right - content.left);
    const float heightDip = std::max(0.0f, content.bottom - content.top);
    if (widthDip <= 0.0f || heightDip <= 0.0f)
    {
        return;
    }

    const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

    const std::wstring seed = _currentPath.empty() ? std::wstring(L"viewer") : LeafNameFromPath(_currentPath);
    const COLORREF accent   = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);

    if (! (_hasTheme && _theme.highContrast))
    {
        const bool hasPreviewImage = HasDisplayImage();
        const uint8_t tintAlpha    = hasPreviewImage ? ((_hasTheme && _theme.darkMode) ? 10u : 8u) : ((_hasTheme && _theme.darkMode) ? 28u : 18u);
        const COLORREF tint        = BlendColor(bg, accent, tintAlpha);
        const float overlayA       = hasPreviewImage ? ((_hasTheme && _theme.darkMode) ? 0.25f : 0.18f) : ((_hasTheme && _theme.darkMode) ? 0.85f : 0.75f);
        brush->SetColor(ColorFFromColorRef(tint, overlayA));
        target->FillRectangle(content, brush);
    }

    const float minDim = std::min(widthDip, heightDip);
    const float radius = std::clamp(minDim * 0.08f, 18.0f, 44.0f);
    const float stroke = std::clamp(radius * 0.20f, 3.0f, 6.0f);
    const float innerR = radius * 0.55f;
    const float outerR = radius;

    const float textHeightDip         = 34.0f;
    const float progressTextHeightDip = 18.0f;
    const float progressGapDip        = 6.0f;
    const float progressBarHeightDip  = 6.0f;
    const float spacingDip            = 14.0f;
    const float groupHeightDip        = outerR * 2.0f + spacingDip + textHeightDip + progressTextHeightDip + progressGapDip + progressBarHeightDip;
    const float groupTopDip           = content.top + std::max(0.0f, (heightDip - groupHeightDip) * 0.5f);

    const float cx = content.left + widthDip * 0.5f;
    const float cy = groupTopDip + outerR;

    constexpr int kSegments = 12;
    constexpr float kPi     = 3.14159265358979323846f;
    const float baseRad     = (_loadingSpinnerAngleDeg - 90.0f) * (kPi / 180.0f);

    const bool rainbowSpinner = _hasTheme && ! _theme.highContrast && _theme.rainbowMode;
    float rainbowHue          = 0.0f;
    float rainbowSat          = 0.0f;
    float rainbowVal          = 0.0f;
    if (rainbowSpinner)
    {
        const uint32_t h = StableHash32(seed);
        rainbowHue       = static_cast<float>(h % 360u);
        rainbowSat       = _theme.darkBase ? 0.70f : 0.55f;
        rainbowVal       = _theme.darkBase ? 0.95f : 0.85f;
    }

    for (int i = 0; i < kSegments; ++i)
    {
        const float t     = static_cast<float>(i) / static_cast<float>(kSegments);
        const float alpha = 0.15f + 0.85f * (1.0f - t);
        const float angle = baseRad + t * (2.0f * kPi);
        const float s     = std::sin(angle);
        const float c     = std::cos(angle);

        const D2D1_POINT_2F p1 = D2D1::Point2F(cx + c * innerR, cy + s * innerR);
        const D2D1_POINT_2F p2 = D2D1::Point2F(cx + c * outerR, cy + s * outerR);

        COLORREF segmentColor = accent;
        if (rainbowSpinner)
        {
            const float hueStep    = 360.0f / static_cast<float>(kSegments);
            const float hueDegrees = rainbowHue + static_cast<float>(i) * hueStep;
            segmentColor           = ColorFromHSV(hueDegrees, rainbowSat, rainbowVal);
        }

        brush->SetColor(ColorFFromColorRef(segmentColor, alpha));
        target->DrawLine(p1, p2, brush, stroke);
    }

    std::wstring loadingText = LoadStringResource(g_hInstance, IDS_VIEWERRAW_STATUS_LOADING);
    if (loadingText.empty())
    {
        return;
    }

    if (! _loadingOverlayFormat && _dwriteFactory)
    {
        wil::com_ptr<IDWriteTextFormat> format;
        const HRESULT hr = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 22.0f, L"", format.put());
        if (SUCCEEDED(hr) && format)
        {
            static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
            static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
            static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
            _loadingOverlayFormat = std::move(format);
        }
    }

    if (! _loadingOverlayFormat)
    {
        return;
    }

    const float textTopDip   = groupTopDip + outerR * 2.0f + spacingDip;
    const D2D1_RECT_F textRc = D2D1::RectF(content.left, textTopDip, content.right, std::min(content.bottom, textTopDip + textHeightDip));
    brush->SetColor(ColorFFromColorRef(fg, 0.90f));

    const UINT32 len = static_cast<UINT32>(std::min<size_t>(loadingText.size(), std::numeric_limits<UINT32>::max()));
    target->DrawTextW(loadingText.c_str(), len, _loadingOverlayFormat.get(), textRc, brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);

    std::wstring progressText;
    if (! _rawProgressStageText.empty() || _rawProgressPercent >= 0)
    {
        if (_rawProgressPercent >= 0 && ! _rawProgressStageText.empty())
        {
            progressText = std::format(L"{}% — {}", _rawProgressPercent, _rawProgressStageText);
        }
        else if (_rawProgressPercent >= 0)
        {
            progressText = std::format(L"{}%", _rawProgressPercent);
        }
        else
        {
            progressText = _rawProgressStageText;
        }
    }

    if (! progressText.empty())
    {
        if (! _loadingOverlaySubFormat && _dwriteFactory)
        {
            wil::com_ptr<IDWriteTextFormat> format;
            const HRESULT hr = _dwriteFactory->CreateTextFormat(
                L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 12.0f, L"", format.put());
            if (SUCCEEDED(hr) && format)
            {
                static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER));
                static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
                static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
                _loadingOverlaySubFormat = std::move(format);
            }
        }

        if (_loadingOverlaySubFormat)
        {
            const float progressTextTopDip = textTopDip + textHeightDip;
            const D2D1_RECT_F progressTextRc =
                D2D1::RectF(content.left, progressTextTopDip, content.right, std::min(content.bottom, progressTextTopDip + progressTextHeightDip));
            brush->SetColor(ColorFFromColorRef(fg, 0.85f));
            target->DrawTextW(progressText.c_str(),
                              static_cast<UINT32>(std::min<size_t>(progressText.size(), std::numeric_limits<UINT32>::max())),
                              _loadingOverlaySubFormat.get(),
                              progressTextRc,
                              brush,
                              D2D1_DRAW_TEXT_OPTIONS_CLIP);
        }
    }

    if (_rawProgressPercent >= 0)
    {
        const float progress = std::clamp(static_cast<float>(_rawProgressPercent) / 100.0f, 0.0f, 1.0f);
        const float barW     = std::clamp(widthDip * 0.38f, 160.0f, 280.0f);
        const float barX     = cx - barW * 0.5f;
        const float barY     = textTopDip + textHeightDip + progressTextHeightDip + progressGapDip;

        const D2D1_RECT_F trackRc = D2D1::RectF(barX, barY, barX + barW, barY + progressBarHeightDip);
        const float r             = progressBarHeightDip * 0.5f;

        if (_hasTheme && _theme.highContrast)
        {
            brush->SetColor(ColorFFromColorRef(fg, 1.0f));
            target->DrawRoundedRectangle(D2D1::RoundedRect(trackRc, r, r), brush, 1.0f);
        }
        else
        {
            const COLORREF track = _hasTheme ? BlendColor(bg, accent, (_theme.darkMode ? 92u : 72u)) : accent;
            brush->SetColor(ColorFFromColorRef(track, 0.55f));
            target->FillRoundedRectangle(D2D1::RoundedRect(trackRc, r, r), brush);
        }

        if (progress > 0.0f)
        {
            const float fillW        = barW * progress;
            const D2D1_RECT_F fillRc = D2D1::RectF(barX, barY, barX + fillW, barY + progressBarHeightDip);
            const float fillR        = std::min(r, fillW * 0.5f);
            brush->SetColor(ColorFFromColorRef(accent, 0.90f));
            target->FillRoundedRectangle(D2D1::RoundedRect(fillRc, fillR, fillR), brush);
        }
    }
    else
    {
        const float barW          = std::clamp(widthDip * 0.38f, 160.0f, 280.0f);
        const float barX          = cx - barW * 0.5f;
        const float barY          = textTopDip + textHeightDip + progressTextHeightDip + progressGapDip;
        const D2D1_RECT_F trackRc = D2D1::RectF(barX, barY, barX + barW, barY + progressBarHeightDip);
        const float r             = progressBarHeightDip * 0.5f;

        if (_hasTheme && _theme.highContrast)
        {
            brush->SetColor(ColorFFromColorRef(fg, 1.0f));
            target->DrawRoundedRectangle(D2D1::RoundedRect(trackRc, r, r), brush, 1.0f);
        }
        else
        {
            const COLORREF track = _hasTheme ? BlendColor(bg, accent, (_theme.darkMode ? 92u : 72u)) : accent;
            brush->SetColor(ColorFFromColorRef(track, 0.50f));
            target->FillRoundedRectangle(D2D1::RoundedRect(trackRc, r, r), brush);
        }

        const float t        = std::fmod(std::max(0.0f, _loadingSpinnerAngleDeg), 360.0f) / 360.0f;
        const float pingPong = (t <= 0.5f) ? (t * 2.0f) : (2.0f - t * 2.0f);
        const float segW     = std::max(40.0f, barW * 0.25f);
        const float segX     = barX + (barW - segW) * pingPong;

        const D2D1_RECT_F fillRc = D2D1::RectF(segX, barY, segX + segW, barY + progressBarHeightDip);
        const float fillR        = std::min(r, segW * 0.5f);
        brush->SetColor(ColorFFromColorRef(accent, 0.85f));
        target->FillRoundedRectangle(D2D1::RoundedRect(fillRc, fillR, fillR), brush);
    }
}

void ViewerImgRaw::DrawExifOverlay(ID2D1HwndRenderTarget* target, ID2D1SolidColorBrush* brush) noexcept
{
    if (! _showExifOverlay || _exifOverlayText.empty() || ! target || ! brush || ! _hWnd)
    {
        return;
    }

    const UINT dpi            = GetDpiForWindow(_hWnd.get());
    const D2D1_RECT_F content = RectFFromPixels(_contentRect, dpi);

    const float widthDip  = std::max(0.0f, content.right - content.left);
    const float heightDip = std::max(0.0f, content.bottom - content.top);
    if (widthDip <= 0.0f || heightDip <= 0.0f)
    {
        return;
    }

    if (! _exifOverlayFormat && _dwriteFactory)
    {
        wil::com_ptr<IDWriteTextFormat> format;
        const HRESULT hr = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 12.0f, L"", format.put());
        if (SUCCEEDED(hr) && format)
        {
            static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));
            static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
            static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP));
            _exifOverlayFormat = std::move(format);
        }
    }

    if (! _exifOverlayFormat || ! _dwriteFactory)
    {
        return;
    }

    const float maxWidthDip  = std::max(120.0f, widthDip * 0.60f);
    const float maxHeightDip = std::max(50.0f, heightDip * 0.60f);

    wil::com_ptr<IDWriteTextLayout> layout;
    const HRESULT layoutHr =
        _dwriteFactory->CreateTextLayout(_exifOverlayText.c_str(),
                                         static_cast<UINT32>(std::min<size_t>(_exifOverlayText.size(), std::numeric_limits<UINT32>::max())),
                                         _exifOverlayFormat.get(),
                                         maxWidthDip,
                                         maxHeightDip,
                                         layout.put());
    if (FAILED(layoutHr) || ! layout)
    {
        return;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return;
    }

    const float paddingDip = 10.0f;
    const float marginDip  = 12.0f;
    const float boxW       = std::clamp(metrics.widthIncludingTrailingWhitespace + paddingDip * 2.0f, 1.0f, widthDip);
    const float boxH       = std::clamp(metrics.height + paddingDip * 2.0f, 1.0f, heightDip);

    float x = content.right - marginDip - boxW;
    float y = content.bottom - marginDip - boxH;
    x       = std::max(content.left, x);
    y       = std::max(content.top, y);

    const D2D1_RECT_F box = D2D1::RectF(x, y, x + boxW, y + boxH);

    const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);

    const std::wstring seed = _currentPath.empty() ? std::wstring(L"viewer") : LeafNameFromPath(_currentPath);
    const COLORREF accent   = _hasTheme ? ResolveAccentColor(_theme, seed) : RGB(0, 120, 215);

    if (_hasTheme && _theme.highContrast)
    {
        brush->SetColor(ColorFFromColorRef(bg, 1.0f));
        target->FillRectangle(box, brush);
        brush->SetColor(ColorFFromColorRef(fg, 1.0f));
        target->DrawRectangle(box, brush, 1.0f);
        brush->SetColor(ColorFFromColorRef(fg, 1.0f));
    }
    else
    {
        const uint8_t tintAlpha = (_hasTheme && _theme.darkMode) ? 46u : 34u;
        const COLORREF tint     = BlendColor(bg, accent, tintAlpha);
        const float overlayA    = (_hasTheme && _theme.darkMode) ? 0.88f : 0.82f;
        brush->SetColor(ColorFFromColorRef(tint, overlayA));
        target->FillRectangle(box, brush);

        brush->SetColor(ColorFFromColorRef(accent, 0.85f));
        target->DrawRectangle(box, brush, 1.0f);

        brush->SetColor(ColorFFromColorRef(fg, 0.95f));
    }

    const D2D1_POINT_2F textPos = D2D1::Point2F(box.left + paddingDip, box.top + paddingDip);
    target->DrawTextLayout(textPos, layout.get(), brush, D2D1_DRAW_TEXT_OPTIONS_CLIP);
}

void ViewerImgRaw::OnCommand(HWND hwnd, UINT id, UINT code, HWND control)
{
    if (id == IDC_VIEWERRAW_FILE_COMBO && code == CBN_SELCHANGE && control == _hFileCombo.get())
    {
        if (_syncingFileCombo)
        {
            return;
        }

        const LRESULT sel = SendMessageW(_hFileCombo.get(), CB_GETCURSEL, 0, 0);
        if (sel != CB_ERR)
        {
            const size_t index = static_cast<size_t>(sel);
            if (index < _otherItems.size())
            {
                _otherIndex             = index;
                _currentSidecarJpegPath = _otherItems[_otherIndex].sidecarJpegPath;
                _currentLabel           = _otherItems[_otherIndex].label;
                StartAsyncOpen(hwnd, _otherItems[_otherIndex].primaryPath, false);
            }
        }
        return;
    }

    switch (id)
    {
        case IDM_VIEWERRAW_FILE_REFRESH:
            if (! _currentPath.empty())
            {
                StartAsyncOpen(hwnd, _currentPath, false);
            }
            return;
        case IDM_VIEWERRAW_FILE_EXPORT: BeginExport(hwnd); return;
        case IDM_VIEWERRAW_FILE_EXIT: DestroyWindow(hwnd); return;
        case IDM_VIEWERRAW_OTHER_NEXT:
            if (_otherItems.size() > 1)
            {
                _otherIndex = (_otherIndex + 1) % _otherItems.size();
                SyncFileComboSelection();
                _currentSidecarJpegPath = _otherItems[_otherIndex].sidecarJpegPath;
                _currentLabel           = _otherItems[_otherIndex].label;
                StartAsyncOpen(hwnd, _otherItems[_otherIndex].primaryPath, false);
            }
            return;
        case IDM_VIEWERRAW_OTHER_PREVIOUS:
            if (_otherItems.size() > 1)
            {
                _otherIndex = (_otherIndex + _otherItems.size() - 1) % _otherItems.size();
                SyncFileComboSelection();
                _currentSidecarJpegPath = _otherItems[_otherIndex].sidecarJpegPath;
                _currentLabel           = _otherItems[_otherIndex].label;
                StartAsyncOpen(hwnd, _otherItems[_otherIndex].primaryPath, false);
            }
            return;
        case IDM_VIEWERRAW_OTHER_FIRST:
            if (_otherItems.size() > 1)
            {
                _otherIndex = 0;
                SyncFileComboSelection();
                _currentSidecarJpegPath = _otherItems[_otherIndex].sidecarJpegPath;
                _currentLabel           = _otherItems[_otherIndex].label;
                StartAsyncOpen(hwnd, _otherItems[_otherIndex].primaryPath, false);
            }
            return;
        case IDM_VIEWERRAW_OTHER_LAST:
            if (_otherItems.size() > 1)
            {
                _otherIndex = _otherItems.size() - 1;
                SyncFileComboSelection();
                _currentSidecarJpegPath = _otherItems[_otherIndex].sidecarJpegPath;
                _currentLabel           = _otherItems[_otherIndex].label;
                StartAsyncOpen(hwnd, _otherItems[_otherIndex].primaryPath, false);
            }
            return;
        case IDM_VIEWERRAW_VIEW_FIT:
            _fitToWindow  = true;
            _panOffsetXPx = 0.0f;
            _panOffsetYPx = 0.0f;
            _panning      = false;
            UpdateMenuChecks(hwnd);
            UpdateScrollBars(hwnd);
            InvalidateRect(hwnd, nullptr, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_ACTUAL_SIZE: ApplyZoom(hwnd, 1.0f, std::nullopt); return;
        case IDM_VIEWERRAW_VIEW_ZOOM_IN:
        {
            float displayedZoom = 1.0f;
            float x             = 0.0f;
            float y             = 0.0f;
            float drawW         = 0.0f;
            float drawH         = 0.0f;
            if (! ComputeImageLayoutPx(displayedZoom, x, y, drawW, drawH) || displayedZoom <= 0.0f)
            {
                return;
            }
            ApplyZoom(hwnd, std::clamp(displayedZoom * 1.25f, 0.01f, 64.0f), std::nullopt);
            return;
        }
        case IDM_VIEWERRAW_VIEW_ZOOM_OUT:
        {
            float displayedZoom = 1.0f;
            float x             = 0.0f;
            float y             = 0.0f;
            float drawW         = 0.0f;
            float drawH         = 0.0f;
            if (! ComputeImageLayoutPx(displayedZoom, x, y, drawW, drawH) || displayedZoom <= 0.0f)
            {
                return;
            }
            ApplyZoom(hwnd, std::clamp(displayedZoom / 1.25f, 0.01f, 64.0f), std::nullopt);
            return;
        }
        case IDM_VIEWERRAW_VIEW_ZOOM_RESET: ApplyZoom(hwnd, 1.0f, std::nullopt); return;
        case IDM_VIEWERRAW_VIEW_TOGGLE_FIT_100:
            if (_fitToWindow)
            {
                _fitToWindow = false;
                _manualZoom  = 1.0f;
            }
            else
            {
                _fitToWindow = true;
            }
            _panOffsetXPx = 0.0f;
            _panOffsetYPx = 0.0f;
            _panning      = false;
            UpdateMenuChecks(hwnd);
            UpdateScrollBars(hwnd);
            InvalidateRect(hwnd, nullptr, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_ROTATE_CW:
            if (_transientZoomActive)
            {
                _fitToWindow         = _transientSavedFitToWindow;
                _manualZoom          = _transientSavedManualZoom;
                _panOffsetXPx        = _transientSavedPanOffsetXPx;
                _panOffsetYPx        = _transientSavedPanOffsetYPx;
                _transientZoomActive = false;
            }
            _userOrientation = ComposeExifOrientation(static_cast<uint16_t>(6), _userOrientation);
            UpdateOrientationState();
            _panOffsetXPx = 0.0f;
            _panOffsetYPx = 0.0f;
            _panning      = false;
            {
                float tmpZoom = 1.0f;
                float tmpX    = 0.0f;
                float tmpY    = 0.0f;
                float tmpW    = 0.0f;
                float tmpH    = 0.0f;
                static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));
            }
            UpdateScrollBars(hwnd);
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_ROTATE_CCW:
            if (_transientZoomActive)
            {
                _fitToWindow         = _transientSavedFitToWindow;
                _manualZoom          = _transientSavedManualZoom;
                _panOffsetXPx        = _transientSavedPanOffsetXPx;
                _panOffsetYPx        = _transientSavedPanOffsetYPx;
                _transientZoomActive = false;
            }
            _userOrientation = ComposeExifOrientation(static_cast<uint16_t>(8), _userOrientation);
            UpdateOrientationState();
            _panOffsetXPx = 0.0f;
            _panOffsetYPx = 0.0f;
            _panning      = false;
            {
                float tmpZoom = 1.0f;
                float tmpX    = 0.0f;
                float tmpY    = 0.0f;
                float tmpW    = 0.0f;
                float tmpH    = 0.0f;
                static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));
            }
            UpdateScrollBars(hwnd);
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_FLIP_HORIZONTAL:
            if (_transientZoomActive)
            {
                _fitToWindow         = _transientSavedFitToWindow;
                _manualZoom          = _transientSavedManualZoom;
                _panOffsetXPx        = _transientSavedPanOffsetXPx;
                _panOffsetYPx        = _transientSavedPanOffsetYPx;
                _transientZoomActive = false;
            }
            _userOrientation = ComposeExifOrientation(static_cast<uint16_t>(2), _userOrientation);
            UpdateOrientationState();
            _panOffsetXPx = 0.0f;
            _panOffsetYPx = 0.0f;
            _panning      = false;
            {
                float tmpZoom = 1.0f;
                float tmpX    = 0.0f;
                float tmpY    = 0.0f;
                float tmpW    = 0.0f;
                float tmpH    = 0.0f;
                static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));
            }
            UpdateScrollBars(hwnd);
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_FLIP_VERTICAL:
            if (_transientZoomActive)
            {
                _fitToWindow         = _transientSavedFitToWindow;
                _manualZoom          = _transientSavedManualZoom;
                _panOffsetXPx        = _transientSavedPanOffsetXPx;
                _panOffsetYPx        = _transientSavedPanOffsetYPx;
                _transientZoomActive = false;
            }
            _userOrientation = ComposeExifOrientation(static_cast<uint16_t>(4), _userOrientation);
            UpdateOrientationState();
            _panOffsetXPx = 0.0f;
            _panOffsetYPx = 0.0f;
            _panning      = false;
            {
                float tmpZoom = 1.0f;
                float tmpX    = 0.0f;
                float tmpY    = 0.0f;
                float tmpW    = 0.0f;
                float tmpH    = 0.0f;
                static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));
            }
            UpdateScrollBars(hwnd);
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_RESET_ORIENTATION:
            if (_transientZoomActive)
            {
                _fitToWindow         = _transientSavedFitToWindow;
                _manualZoom          = _transientSavedManualZoom;
                _panOffsetXPx        = _transientSavedPanOffsetXPx;
                _panOffsetYPx        = _transientSavedPanOffsetYPx;
                _transientZoomActive = false;
            }
            _userOrientation = 1;
            UpdateOrientationState();
            _panOffsetXPx = 0.0f;
            _panOffsetYPx = 0.0f;
            _panning      = false;
            {
                float tmpZoom = 1.0f;
                float tmpX    = 0.0f;
                float tmpY    = 0.0f;
                float tmpW    = 0.0f;
                float tmpH    = 0.0f;
                static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));
            }
            UpdateScrollBars(hwnd);
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_BRIGHTNESS_INCREASE:
            _brightness = std::clamp(_brightness + 0.05f, -1.0f, 1.0f);
            _imageBitmap.reset();
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_BRIGHTNESS_DECREASE:
            _brightness = std::clamp(_brightness - 0.05f, -1.0f, 1.0f);
            _imageBitmap.reset();
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_CONTRAST_INCREASE:
            _contrast = std::clamp(_contrast + 0.05f, 0.10f, 3.00f);
            _imageBitmap.reset();
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_CONTRAST_DECREASE:
            _contrast = std::clamp(_contrast - 0.05f, 0.10f, 3.00f);
            _imageBitmap.reset();
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_GAMMA_INCREASE:
            _gamma = std::clamp(_gamma + 0.05f, 0.10f, 5.00f);
            _imageBitmap.reset();
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_GAMMA_DECREASE:
            _gamma = std::clamp(_gamma - 0.05f, 0.10f, 5.00f);
            _imageBitmap.reset();
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_TOGGLE_GRAYSCALE:
            _grayscale = ! _grayscale;
            _imageBitmap.reset();
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_TOGGLE_NEGATIVE:
            _negative = ! _negative;
            _imageBitmap.reset();
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_SOURCE_RAW:
            SetDisplayMode(DisplayMode::Raw);
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, nullptr, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_SOURCE_THUMBNAIL:
            SetDisplayMode(DisplayMode::Thumbnail);
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, nullptr, FALSE);
            return;
        case IDM_VIEWERRAW_VIEW_SHOW_EXIF_OVERLAY:
            _showExifOverlay = ! _showExifOverlay;
            RebuildExifOverlayText();
            UpdateMenuChecks(hwnd);
            InvalidateRect(hwnd, nullptr, FALSE);
            return;
        default: return;
    }
}

void ViewerImgRaw::OnKeyDown(HWND hwnd, UINT vk) noexcept
{
    const bool ctrl  = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    const bool alt   = (GetKeyState(VK_MENU) & 0x8000) != 0;

    if (vk == VK_ESCAPE)
    {
        if (_alertVisible && _hostAlerts)
        {
            static_cast<void>(_hostAlerts->ClearAlert(HOST_ALERT_SCOPE_WINDOW, nullptr));
            _alertVisible = false;
            return;
        }

        DestroyWindow(hwnd);
        return;
    }

    if (vk == VK_F5)
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_FILE_REFRESH, 0), 0);
        return;
    }

    if (! ctrl && ! shift && (vk == VK_SPACE || vk == VK_RIGHT || vk == VK_NEXT))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_OTHER_NEXT, 0), 0);
        return;
    }

    if (! ctrl && ! shift && (vk == VK_BACK || vk == VK_LEFT || vk == VK_PRIOR))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_OTHER_PREVIOUS, 0), 0);
        return;
    }

    if (! ctrl && ! shift && vk == VK_HOME)
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_OTHER_FIRST, 0), 0);
        return;
    }

    if (! ctrl && ! shift && vk == VK_END)
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_OTHER_LAST, 0), 0);
        return;
    }

    if (ctrl && ! alt && ! shift && (vk == VK_LEFT || vk == VK_RIGHT || vk == VK_UP || vk == VK_DOWN))
    {
        if (! HasDisplayImage())
        {
            return;
        }

        const UINT dpi     = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;
        const float stepPx = static_cast<float>(PxFromDip(40, dpi));

        if (vk == VK_LEFT)
        {
            _panOffsetXPx += stepPx;
        }
        else if (vk == VK_RIGHT)
        {
            _panOffsetXPx -= stepPx;
        }
        else if (vk == VK_UP)
        {
            _panOffsetYPx += stepPx;
        }
        else if (vk == VK_DOWN)
        {
            _panOffsetYPx -= stepPx;
        }

        float tmpZoom = 1.0f;
        float tmpX    = 0.0f;
        float tmpY    = 0.0f;
        float tmpW    = 0.0f;
        float tmpH    = 0.0f;
        static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));

        UpdateScrollBars(hwnd);
        InvalidateRect(hwnd, &_contentRect, FALSE);
        InvalidateRect(hwnd, &_statusRect, FALSE);
        return;
    }

    if (ctrl && alt && (vk == VK_UP || vk == VK_DOWN || vk == VK_LEFT || vk == VK_RIGHT || vk == VK_PRIOR || vk == VK_NEXT))
    {
        UINT cmd = 0;
        switch (vk)
        {
            case VK_UP: cmd = IDM_VIEWERRAW_VIEW_BRIGHTNESS_INCREASE; break;
            case VK_DOWN: cmd = IDM_VIEWERRAW_VIEW_BRIGHTNESS_DECREASE; break;
            case VK_RIGHT: cmd = IDM_VIEWERRAW_VIEW_CONTRAST_INCREASE; break;
            case VK_LEFT: cmd = IDM_VIEWERRAW_VIEW_CONTRAST_DECREASE; break;
            case VK_PRIOR: cmd = IDM_VIEWERRAW_VIEW_GAMMA_INCREASE; break;
            case VK_NEXT: cmd = IDM_VIEWERRAW_VIEW_GAMMA_DECREASE; break;
            default: cmd = 0; break;
        }
        if (cmd != 0)
        {
            SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(cmd, 0), 0);
        }
        return;
    }

    if (ctrl && ! alt && ! shift && (vk == 'S' || vk == 's'))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_FILE_EXPORT, 0), 0);
        return;
    }

    if (! ctrl && ! alt && (vk == VK_ADD || vk == VK_OEM_PLUS))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_ZOOM_IN, 0), 0);
        return;
    }

    if (! ctrl && ! alt && (vk == VK_SUBTRACT || vk == VK_OEM_MINUS))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_ZOOM_OUT, 0), 0);
        return;
    }

    if (! ctrl && ! alt && vk == '0')
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_ZOOM_RESET, 0), 0);
        return;
    }

    if (ctrl && ! alt && ! shift && (vk == 'F' || vk == 'f'))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_FIT, 0), 0);
        return;
    }

    if (! ctrl && ! alt && (vk == 'F' || vk == 'f'))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_TOGGLE_FIT_100, 0), 0);
        return;
    }

    if (! ctrl && ! alt && ! shift && (vk == 'R' || vk == 'r'))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_ROTATE_CW, 0), 0);
        return;
    }

    if (! alt && (vk == 'R' || vk == 'r') && (ctrl || (! ctrl && shift)))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_ROTATE_CCW, 0), 0);
        return;
    }

    if (! ctrl && ! alt && (vk == 'H' || vk == 'h'))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_FLIP_HORIZONTAL, 0), 0);
        return;
    }

    if (! ctrl && ! alt && (vk == 'V' || vk == 'v'))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_FLIP_VERTICAL, 0), 0);
        return;
    }

    if (! ctrl && ! alt && (vk == 'O' || vk == 'o'))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_RESET_ORIENTATION, 0), 0);
        return;
    }

    if (! ctrl && ! alt && (vk == 'G' || vk == 'g'))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_TOGGLE_GRAYSCALE, 0), 0);
        return;
    }

    if (! ctrl && ! alt && (vk == 'N' || vk == 'n'))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_TOGGLE_NEGATIVE, 0), 0);
        return;
    }

    if (! ctrl && ! alt && vk == '1')
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_ACTUAL_SIZE, 0), 0);
        return;
    }

    if (! ctrl && ! alt && (vk == 'I' || vk == 'i'))
    {
        SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDM_VIEWERRAW_VIEW_SHOW_EXIF_OVERLAY, 0), 0);
        return;
    }
}

void ViewerImgRaw::OnLButtonDown(HWND hwnd, int x, int y) noexcept
{
    if (! hwnd || ! HasDisplayImage() || ! _currentImage)
    {
        return;
    }

    const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;

    if (_transientZoomActive)
    {
        _fitToWindow         = _transientSavedFitToWindow;
        _manualZoom          = _transientSavedManualZoom;
        _panOffsetXPx        = _transientSavedPanOffsetXPx;
        _panOffsetYPx        = _transientSavedPanOffsetYPx;
        _transientZoomActive = false;
    }

    const POINT pt{x, y};

    float displayedZoom = 1.0f;
    float imgX          = 0.0f;
    float imgY          = 0.0f;
    float drawX         = 0.0f;
    float drawY         = 0.0f;
    float drawW         = 0.0f;
    float drawH         = 0.0f;
    if (! ComputeImageLayoutPx(displayedZoom, drawX, drawY, drawW, drawH) || displayedZoom <= 0.0f || drawW <= 0.0f || drawH <= 0.0f)
    {
        return;
    }

    const float ptX = static_cast<float>(pt.x);
    const float ptY = static_cast<float>(pt.y);
    if (ptX < drawX || ptX >= drawX + drawW || ptY < drawY || ptY >= drawY + drawH)
    {
        return;
    }

    imgX = (ptX - drawX) / displayedZoom;
    imgY = (ptY - drawY) / displayedZoom;

    const CachedImage* image   = _currentImage;
    const uint32_t imgWPx      = IsDisplayingThumbnail() ? image->thumbWidth : image->rawWidth;
    const uint32_t imgHPx      = IsDisplayingThumbnail() ? image->thumbHeight : image->rawHeight;
    const uint16_t orientation = (_viewOrientation >= 1 && _viewOrientation <= 8) ? _viewOrientation : static_cast<uint16_t>(1);
    const bool swapAxes        = (orientation >= 5 && orientation <= 8);
    const float imgW           = static_cast<float>(swapAxes ? imgHPx : imgWPx);
    const float imgH           = static_cast<float>(swapAxes ? imgWPx : imgHPx);

    const float contentW = static_cast<float>(std::max(0L, _contentRect.right - _contentRect.left));
    const float contentH = static_cast<float>(std::max(0L, _contentRect.bottom - _contentRect.top));
    const bool canPan    = (drawW > contentW) || (drawH > contentH);

    if (ctrl && imgW > 0.0f && imgH > 0.0f)
    {
        const float newZoom = std::clamp(static_cast<float>(_config.zoomOnClickPercent) / 100.0f, 0.01f, 64.0f);

        _transientSavedFitToWindow  = _fitToWindow;
        _transientSavedManualZoom   = _manualZoom;
        _transientSavedPanOffsetXPx = _panOffsetXPx;
        _transientSavedPanOffsetYPx = _panOffsetYPx;
        _transientZoomActive        = true;

        const float newDrawW = imgW * newZoom;
        const float newDrawH = imgH * newZoom;

        const float baseX = static_cast<float>(_contentRect.left) + (contentW - newDrawW) / 2.0f;
        const float baseY = static_cast<float>(_contentRect.top) + (contentH - newDrawH) / 2.0f;

        const float desiredX = ptX - imgX * newZoom;
        const float desiredY = ptY - imgY * newZoom;

        _fitToWindow  = false;
        _manualZoom   = newZoom;
        _panOffsetXPx = desiredX - baseX;
        _panOffsetYPx = desiredY - baseY;

        float tmpZoom = 1.0f;
        float tmpX    = 0.0f;
        float tmpY    = 0.0f;
        float tmpW    = 0.0f;
        float tmpH    = 0.0f;
        static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));
        UpdateScrollBars(hwnd);
        InvalidateRect(hwnd, &_contentRect, FALSE);
        InvalidateRect(hwnd, &_statusRect, FALSE);
    }

    if (! ctrl && ! canPan)
    {
        return;
    }

    _panning           = true;
    _panStartPoint     = pt;
    _panStartOffsetXPx = _panOffsetXPx;
    _panStartOffsetYPx = _panOffsetYPx;
    SetCapture(hwnd);
}

void ViewerImgRaw::OnLButtonDblClick(HWND hwnd, int /*x*/, int /*y*/) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const UINT cmd  = ctrl ? IDM_VIEWERRAW_VIEW_ACTUAL_SIZE : IDM_VIEWERRAW_VIEW_FIT;
    SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(cmd, 0), 0);
}

void ViewerImgRaw::OnLButtonUp(HWND hwnd) noexcept
{
    if (! _panning)
    {
        if (_transientZoomActive)
        {
            _fitToWindow         = _transientSavedFitToWindow;
            _manualZoom          = _transientSavedManualZoom;
            _panOffsetXPx        = _transientSavedPanOffsetXPx;
            _panOffsetYPx        = _transientSavedPanOffsetYPx;
            _transientZoomActive = false;

            float tmpZoom = 1.0f;
            float tmpX    = 0.0f;
            float tmpY    = 0.0f;
            float tmpW    = 0.0f;
            float tmpH    = 0.0f;
            static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));

            if (hwnd)
            {
                UpdateScrollBars(hwnd);
                InvalidateRect(hwnd, &_contentRect, FALSE);
                InvalidateRect(hwnd, &_statusRect, FALSE);
            }
        }
        return;
    }

    _panning = false;
    if (hwnd && GetCapture() == hwnd)
    {
        ReleaseCapture();
    }

    if (_transientZoomActive)
    {
        _fitToWindow         = _transientSavedFitToWindow;
        _manualZoom          = _transientSavedManualZoom;
        _panOffsetXPx        = _transientSavedPanOffsetXPx;
        _panOffsetYPx        = _transientSavedPanOffsetYPx;
        _transientZoomActive = false;

        float tmpZoom = 1.0f;
        float tmpX    = 0.0f;
        float tmpY    = 0.0f;
        float tmpW    = 0.0f;
        float tmpH    = 0.0f;
        static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));

        if (hwnd)
        {
            UpdateScrollBars(hwnd);
            InvalidateRect(hwnd, &_contentRect, FALSE);
            InvalidateRect(hwnd, &_statusRect, FALSE);
        }
    }
}

void ViewerImgRaw::OnMouseMove(HWND hwnd, int x, int y) noexcept
{
    if (! _panning || ! hwnd || GetCapture() != hwnd)
    {
        return;
    }

    const float dx = static_cast<float>(x - _panStartPoint.x);
    const float dy = static_cast<float>(y - _panStartPoint.y);

    _panOffsetXPx = _panStartOffsetXPx + dx;
    _panOffsetYPx = _panStartOffsetYPx + dy;

    float tmpZoom = 1.0f;
    float tmpX    = 0.0f;
    float tmpY    = 0.0f;
    float tmpW    = 0.0f;
    float tmpH    = 0.0f;
    static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));

    UpdateScrollBars(hwnd);
    InvalidateRect(hwnd, &_contentRect, FALSE);
    InvalidateRect(hwnd, &_statusRect, FALSE);
}

void ViewerImgRaw::OnCaptureChanged() noexcept
{
    _panning = false;

    if (_transientZoomActive)
    {
        _fitToWindow         = _transientSavedFitToWindow;
        _manualZoom          = _transientSavedManualZoom;
        _panOffsetXPx        = _transientSavedPanOffsetXPx;
        _panOffsetYPx        = _transientSavedPanOffsetYPx;
        _transientZoomActive = false;

        float tmpZoom = 1.0f;
        float tmpX    = 0.0f;
        float tmpY    = 0.0f;
        float tmpW    = 0.0f;
        float tmpH    = 0.0f;
        static_cast<void>(ComputeImageLayoutPx(tmpZoom, tmpX, tmpY, tmpW, tmpH));

        if (_hWnd)
        {
            UpdateScrollBars(_hWnd.get());
            InvalidateRect(_hWnd.get(), &_contentRect, FALSE);
            InvalidateRect(_hWnd.get(), &_statusRect, FALSE);
        }
    }
}

void ViewerImgRaw::RefreshFileCombo(HWND hwnd) noexcept
{
    if (! _hFileCombo)
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    SendMessageW(_hFileCombo.get(), CB_RESETCONTENT, 0, 0);

    if (_otherItems.size() <= 1)
    {
        SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(-1), 0);
        if (hwnd)
        {
            Layout(hwnd);
            InvalidateRect(hwnd, nullptr, TRUE);
        }
        return;
    }

    for (const auto& item : _otherItems)
    {
        const wchar_t* text = item.label.empty() ? item.primaryPath.c_str() : item.label.c_str();
        SendMessageW(_hFileCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(text));
    }

    if (_otherIndex >= _otherItems.size())
    {
        _otherIndex = 0;
    }

    SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(_otherIndex), 0);
    SendMessageW(_hFileCombo.get(), CB_SETMINVISIBLE, static_cast<WPARAM>(std::min<size_t>(_otherItems.size(), 15)), 0);

    if (hwnd)
    {
        Layout(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);
    }
}

void ViewerImgRaw::SyncFileComboSelection() noexcept
{
    if (! _hFileCombo)
    {
        return;
    }

    if (_otherItems.size() <= 1)
    {
        return;
    }

    if (_otherIndex >= _otherItems.size())
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    SendMessageW(_hFileCombo.get(), CB_SETCURSEL, static_cast<WPARAM>(_otherIndex), 0);
}

bool ViewerImgRaw::EnsureDirect2D(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    const UINT dpi   = GetDpiForWindow(hwnd);
    const float dpiF = static_cast<float>(dpi);

    if (! _d2dFactory)
    {
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, _d2dFactory.put());
        if (FAILED(hr) || ! _d2dFactory)
        {
            _d2dFactory.reset();
            return false;
        }
    }

    if (! _dwriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwriteFactory.put()));
        if (FAILED(hr) || ! _dwriteFactory)
        {
            _dwriteFactory.reset();
            return false;
        }
    }

    if (! _d2dTarget)
    {
        RECT client{};
        GetClientRect(hwnd, &client);

        const UINT32 width     = static_cast<UINT32>(std::max<LONG>(0, client.right - client.left));
        const UINT32 height    = static_cast<UINT32>(std::max<LONG>(0, client.bottom - client.top));
        const D2D1_SIZE_U size = D2D1::SizeU(width, height);

        D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
        props.dpiX                          = dpiF;
        props.dpiY                          = dpiF;

        const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, size, D2D1_PRESENT_OPTIONS_RETAIN_CONTENTS);
        const HRESULT hr                                   = _d2dFactory->CreateHwndRenderTarget(props, hwndProps, _d2dTarget.put());
        if (FAILED(hr) || ! _d2dTarget)
        {
            _d2dTarget.reset();
            return false;
        }

        _d2dTarget->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    }
    else
    {
        _d2dTarget->SetDpi(dpiF, dpiF);
    }

    if (! _solidBrush && _d2dTarget)
    {
        const HRESULT hr = _d2dTarget->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::Black), _solidBrush.put());
        if (FAILED(hr))
        {
            _solidBrush.reset();
            return false;
        }
    }

    if (! _uiTextFormat && _dwriteFactory)
    {
        const float fontSizeDip = 12.0f;
        const HRESULT hr        = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, fontSizeDip, L"", _uiTextFormat.put());
        if (FAILED(hr))
        {
            _uiTextFormat.reset();
            return false;
        }

        _uiTextFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        _uiTextFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
        _uiTextFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
    }

    if (! _uiTextFormatRight && _dwriteFactory)
    {
        const float fontSizeDip = 12.0f;
        const HRESULT hr        = _dwriteFactory->CreateTextFormat(
            L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, fontSizeDip, L"", _uiTextFormatRight.put());
        if (FAILED(hr))
        {
            _uiTextFormatRight.reset();
            return false;
        }

        _uiTextFormatRight->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);
        _uiTextFormatRight->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
        _uiTextFormatRight->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
    }

    return true;
}

void ViewerImgRaw::DiscardDirect2D() noexcept
{
    _imageBitmap.reset();
    _solidBrush.reset();
    _uiTextFormat.reset();
    _uiTextFormatRight.reset();
    _loadingOverlayFormat.reset();
    _loadingOverlaySubFormat.reset();
    _d2dTarget.reset();
    _dwriteFactory.reset();
    _d2dFactory.reset();
}

bool ViewerImgRaw::EnsureImageBitmap() noexcept
{
    if (_imageBitmap || ! _d2dTarget)
    {
        return _imageBitmap != nullptr;
    }

    const CachedImage* image = _currentImage;
    if (! image)
    {
        return false;
    }

    const uint32_t w                 = IsDisplayingThumbnail() ? image->thumbWidth : image->rawWidth;
    const uint32_t h                 = IsDisplayingThumbnail() ? image->thumbHeight : image->rawHeight;
    const std::vector<uint8_t>& bgra = IsDisplayingThumbnail() ? image->thumbBgra : image->rawBgra;
    if (bgra.empty() || w == 0 || h == 0)
    {
        return false;
    }

    const bool needAdjust =
        std::fabs(_brightness) > 0.001f || std::fabs(_contrast - 1.0f) > 0.001f || std::fabs(_gamma - 1.0f) > 0.001f || _grayscale || _negative;

    const uint8_t* uploadData = bgra.data();
    if (needAdjust)
    {
        if (_adjustedBgra.size() != bgra.size())
        {
            _adjustedBgra.assign(bgra.size(), 0);
        }

        std::array<uint8_t, 256> map{};
        const float gamma      = std::clamp(_gamma, 0.10f, 5.00f);
        const float invGamma   = (std::fabs(gamma - 1.0f) < 0.0001f) ? 1.0f : (1.0f / gamma);
        const float contrast   = std::clamp(_contrast, 0.10f, 3.00f);
        const float brightness = std::clamp(_brightness, -1.0f, 1.0f);

        for (int i = 0; i < 256; ++i)
        {
            float x = static_cast<float>(i) / 255.0f;
            x       = (x - 0.5f) * contrast + 0.5f + brightness;
            x       = std::clamp(x, 0.0f, 1.0f);
            if (invGamma != 1.0f)
            {
                x = std::pow(x, invGamma);
            }
            const int out               = static_cast<int>(std::lround(x * 255.0f));
            map[static_cast<size_t>(i)] = static_cast<uint8_t>(std::clamp(out, 0, 255));
        }

        const uint8_t* src      = bgra.data();
        uint8_t* dst            = _adjustedBgra.data();
        const size_t pixelCount = static_cast<size_t>(w) * static_cast<size_t>(h);
        for (size_t i = 0; i < pixelCount; ++i)
        {
            const size_t di = i * 4u;
            uint8_t b       = src[di + 0];
            uint8_t g       = src[di + 1];
            uint8_t r       = src[di + 2];

            if (_negative)
            {
                b = static_cast<uint8_t>(255u - b);
                g = static_cast<uint8_t>(255u - g);
                r = static_cast<uint8_t>(255u - r);
            }

            if (_grayscale)
            {
                const uint16_t y = static_cast<uint16_t>((54u * r + 183u * g + 19u * b + 128u) >> 8);
                const uint8_t o  = map[static_cast<size_t>(std::min<uint16_t>(y, 255u))];
                dst[di + 0]      = o;
                dst[di + 1]      = o;
                dst[di + 2]      = o;
            }
            else
            {
                dst[di + 0] = map[static_cast<size_t>(b)];
                dst[di + 1] = map[static_cast<size_t>(g)];
                dst[di + 2] = map[static_cast<size_t>(r)];
            }

            dst[di + 3] = 255u;
        }

        uploadData = _adjustedBgra.data();
    }

    const UINT32 stride                = w * 4u;
    const D2D1_BITMAP_PROPERTIES props = D2D1::BitmapProperties(D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE));
    const HRESULT hr                   = _d2dTarget->CreateBitmap(D2D1::SizeU(w, h), uploadData, stride, props, _imageBitmap.put());
    if (FAILED(hr))
    {
        _imageBitmap.reset();
        return false;
    }

    return true;
}

bool ViewerImgRaw::ComputeImageLayoutPx(float& outZoom, float& outX, float& outY, float& outDrawW, float& outDrawH) noexcept
{
    outZoom  = _manualZoom;
    outX     = 0.0f;
    outY     = 0.0f;
    outDrawW = 0.0f;
    outDrawH = 0.0f;

    const CachedImage* image = _currentImage;
    if (! image)
    {
        return false;
    }

    const uint32_t imgWPx = IsDisplayingThumbnail() ? image->thumbWidth : image->rawWidth;
    const uint32_t imgHPx = IsDisplayingThumbnail() ? image->thumbHeight : image->rawHeight;
    if (imgWPx == 0 || imgHPx == 0)
    {
        return false;
    }

    const uint16_t orientation = (_viewOrientation >= 1 && _viewOrientation <= 8) ? _viewOrientation : static_cast<uint16_t>(1);
    const bool swapAxes        = (orientation >= 5 && orientation <= 8);

    const float contentW = static_cast<float>(std::max(0L, _contentRect.right - _contentRect.left));
    const float contentH = static_cast<float>(std::max(0L, _contentRect.bottom - _contentRect.top));
    const float imgW     = static_cast<float>(swapAxes ? imgHPx : imgWPx);
    const float imgH     = static_cast<float>(swapAxes ? imgWPx : imgHPx);

    float zoom = std::clamp(_manualZoom, 0.01f, 64.0f);
    if (_fitToWindow)
    {
        const float sx = imgW > 0.0f ? (contentW / imgW) : 1.0f;
        const float sy = imgH > 0.0f ? (contentH / imgH) : 1.0f;
        zoom           = std::min(sx, sy);
        _panOffsetXPx  = 0.0f;
        _panOffsetYPx  = 0.0f;
    }

    zoom    = std::clamp(zoom, 0.01f, 64.0f);
    outZoom = zoom;

    const float drawW = imgW * zoom;
    const float drawH = imgH * zoom;
    outDrawW          = drawW;
    outDrawH          = drawH;

    const float baseX = static_cast<float>(_contentRect.left) + (contentW - drawW) / 2.0f;
    const float baseY = static_cast<float>(_contentRect.top) + (contentH - drawH) / 2.0f;

    float x = baseX + _panOffsetXPx;
    float y = baseY + _panOffsetYPx;

    if (drawW <= contentW)
    {
        _panOffsetXPx = 0.0f;
        x             = baseX;
    }
    else
    {
        const float minX = static_cast<float>(_contentRect.left) + (contentW - drawW);
        const float maxX = static_cast<float>(_contentRect.left);
        x                = std::clamp(x, minX, maxX);
        _panOffsetXPx    = x - baseX;
    }

    if (drawH <= contentH)
    {
        _panOffsetYPx = 0.0f;
        y             = baseY;
    }
    else
    {
        const float minY = static_cast<float>(_contentRect.top) + (contentH - drawH);
        const float maxY = static_cast<float>(_contentRect.top);
        y                = std::clamp(y, minY, maxY);
        _panOffsetYPx    = y - baseY;
    }

    outX = x;
    outY = y;
    return true;
}

void ViewerImgRaw::OnPaint()
{
    const HWND hwnd = _hWnd.get();
    if (! hwnd)
    {
        return;
    }

    PAINTSTRUCT ps{};
    auto paintDc = wil::BeginPaint(hwnd, &ps);
    static_cast<void>(paintDc);

    const bool ok = EnsureDirect2D(hwnd);
    if (ok && _d2dTarget)
    {
        const UINT dpi = GetDpiForWindow(hwnd);

        _d2dTarget->BeginDraw();
        auto endDraw = wil::scope_exit(
            [&]
            {
                if (! _d2dTarget)
                {
                    return;
                }

                const HRESULT hrEnd = _d2dTarget->EndDraw();
                if (hrEnd == D2DERR_RECREATE_TARGET)
                {
                    DiscardDirect2D();
                }
            });

        const auto rectF = [&](const RECT& rc) noexcept { return RectFFromPixels(rc, dpi); };

        const auto colorF = [](COLORREF c) noexcept
        {
            return D2D1::ColorF(
                static_cast<float>(GetRValue(c)) / 255.0f, static_cast<float>(GetGValue(c)) / 255.0f, static_cast<float>(GetBValue(c)) / 255.0f, 1.0f);
        };

        const D2D1_RECT_F paintRc = rectF(ps.rcPaint);
        _d2dTarget->PushAxisAlignedClip(paintRc, D2D1_ANTIALIAS_MODE_ALIASED);

        if (_solidBrush)
        {
            _solidBrush->SetColor(colorF(_uiBg));
            _d2dTarget->FillRectangle(paintRc, _solidBrush.get());

            _solidBrush->SetColor(colorF(_uiHeaderBg));
            _d2dTarget->FillRectangle(rectF(_headerRect), _solidBrush.get());
            _solidBrush->SetColor(colorF(_uiStatusBg));
            _d2dTarget->FillRectangle(rectF(_statusRect), _solidBrush.get());
        }

        // Image
        const bool hasBitmap = EnsureImageBitmap();
        float displayedZoom  = _manualZoom;
        bool drewImage       = false;
        if (hasBitmap && _imageBitmap)
        {
            float x     = 0.0f;
            float y     = 0.0f;
            float drawW = 0.0f;
            float drawH = 0.0f;
            if (ComputeImageLayoutPx(displayedZoom, x, y, drawW, drawH))
            {
                const CachedImage* image = _currentImage;
                const uint32_t imgWPx    = image ? (IsDisplayingThumbnail() ? image->thumbWidth : image->rawWidth) : 0;
                const uint32_t imgHPx    = image ? (IsDisplayingThumbnail() ? image->thumbHeight : image->rawHeight) : 0;
                const float imgWDip      = DipsFromPixelsF(static_cast<float>(imgWPx), dpi);
                const float imgHDip      = DipsFromPixelsF(static_cast<float>(imgHPx), dpi);

                if (imgWDip > 0.0f && imgHDip > 0.0f)
                {
                    D2D1_MATRIX_3X2_F oldTransform{};
                    _d2dTarget->GetTransform(&oldTransform);
                    auto restoreTransform = wil::scope_exit(
                        [&]
                        {
                            if (_d2dTarget)
                            {
                                _d2dTarget->SetTransform(oldTransform);
                            }
                        });

                    const uint16_t orientation        = (_viewOrientation >= 1 && _viewOrientation <= 8) ? _viewOrientation : static_cast<uint16_t>(1);
                    const float xDip                  = DipsFromPixelsF(x, dpi);
                    const float yDip                  = DipsFromPixelsF(y, dpi);
                    const D2D1_MATRIX_3X2_F transform = ExifOrientationTransform(orientation, imgWDip, imgHDip) *
                                                        D2D1::Matrix3x2F::Scale(displayedZoom, displayedZoom) * D2D1::Matrix3x2F::Translation(xDip, yDip);
                    _d2dTarget->SetTransform(transform);

                    const D2D1_RECT_F dstLocal = D2D1::RectF(0.0f, 0.0f, imgWDip, imgHDip);
                    _d2dTarget->DrawBitmap(_imageBitmap.get(), dstLocal, 1.0f, D2D1_BITMAP_INTERPOLATION_MODE_LINEAR);
                    drewImage = true;
                }
            }
        }

        // Loading overlay
        DrawLoadingOverlay(_d2dTarget.get(), _solidBrush.get());
        DrawExifOverlay(_d2dTarget.get(), _solidBrush.get());

        // Status text
        if (_solidBrush && _uiTextFormat)
        {
            _solidBrush->SetColor(colorF(_uiText));
            const std::wstring leftText  = BuildStatusBarText(drewImage, displayedZoom);
            const std::wstring rightText = _uiTextFormatRight ? BuildStatusBarRightText(drewImage, displayedZoom) : std::wstring();

            D2D1_RECT_F statusTextRc = rectF(_statusRect);
            statusTextRc.left += 10.0f;
            statusTextRc.right -= 10.0f;
            if (statusTextRc.right > statusTextRc.left)
            {
                _d2dTarget->DrawTextW(
                    leftText.c_str(), static_cast<UINT32>(leftText.size()), _uiTextFormat.get(), statusTextRc, _solidBrush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                if (_uiTextFormatRight && ! rightText.empty())
                {
                    _d2dTarget->DrawTextW(rightText.c_str(),
                                          static_cast<UINT32>(rightText.size()),
                                          _uiTextFormatRight.get(),
                                          statusTextRc,
                                          _solidBrush.get(),
                                          D2D1_DRAW_TEXT_OPTIONS_CLIP);
                }
            }
        }

        _d2dTarget->PopAxisAlignedClip();
    }

    _allowEraseBkgnd = false;
}

HRESULT STDMETHODCALLTYPE ViewerImgRaw::Open(const ViewerOpenContext* context) noexcept
{
    if (context == nullptr || context->fileSystem == nullptr || context->focusedPath == nullptr || context->focusedPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    _fileSystem     = context->fileSystem;
    _fileSystemName = context->fileSystemName ? context->fileSystemName : L"";

    if (! _hWnd)
    {
        if (_hasTheme)
        {
            RequestViewerImgRawClassBackgroundColor(ColorRefFromArgb(_theme.backgroundArgb));
        }

        if (! RegisterWndClass(g_hInstance))
        {
            return E_FAIL;
        }

        HWND ownerWindow = context->ownerWindow;
        RECT ownerRc{};
        int x = CW_USEDEFAULT;
        int y = CW_USEDEFAULT;
        int w = 1000;
        int h = 700;
        if (ownerWindow && GetWindowRect(ownerWindow, &ownerRc) != 0)
        {
            x = ownerRc.left;
            y = ownerRc.top;
            w = std::max(1L, ownerRc.right - ownerRc.left);
            h = std::max(1L, ownerRc.bottom - ownerRc.top);
        }

        wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(LoadMenuW(g_hInstance, MAKEINTRESOURCEW(IDR_VIEWERRAW_MENU)));
        HWND window = CreateWindowExW(0,
                                      kClassName,
                                      _metaName.empty() ? L"" : _metaName.c_str(),
                                      WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN | WS_HSCROLL | WS_VSCROLL,
                                      x,
                                      y,
                                      w,
                                      h,
                                      nullptr,
                                      menu.get(),
                                      g_hInstance,
                                      this);
        if (! window)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        menu.release();
        _hWnd.reset(window);

        ApplyTheme(_hWnd.get());
        AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
    }
    else
    {
        ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hWnd.get()));
    }

    _otherItems.clear();

    std::vector<std::wstring> fileList;
    if (context->otherFiles && context->otherFileCount > 0)
    {
        fileList.reserve(static_cast<size_t>(context->otherFileCount));
        for (unsigned long i = 0; i < context->otherFileCount; ++i)
        {
            const wchar_t* p = context->otherFiles[i];
            if (p && p[0] != L'\0')
            {
                fileList.emplace_back(p);
            }
        }
    }

    if (fileList.empty())
    {
        fileList.emplace_back(context->focusedPath);
    }

    struct PairInfo final
    {
        std::wstring rawPath;
        std::wstring jpgPath;
    };

    std::unordered_map<std::wstring, PairInfo> pairsByBaseLower;
    pairsByBaseLower.reserve(fileList.size());

    for (const auto& path : fileList)
    {
        const std::wstring extLower = ToLowerCopy(PathExtensionView(path));
        if (! IsLikelyRawExtension(extLower) && ! IsJpegExtension(extLower))
        {
            continue;
        }

        const std::wstring baseLower = ToLowerCopy(PathWithoutExtensionView(path));
        if (baseLower.empty())
        {
            continue;
        }

        PairInfo& info = pairsByBaseLower[baseLower];
        if (IsLikelyRawExtension(extLower))
        {
            if (info.rawPath.empty())
            {
                info.rawPath = path;
            }
        }
        else if (IsJpegExtension(extLower))
        {
            if (info.jpgPath.empty())
            {
                info.jpgPath = path;
            }
        }
    }

    std::unordered_set<std::wstring> emittedBasesLower;
    emittedBasesLower.reserve(pairsByBaseLower.size());

    _otherItems.reserve(fileList.size());
    for (const auto& path : fileList)
    {
        const std::wstring extLower = ToLowerCopy(PathExtensionView(path));
        const bool isRawOrJpeg      = IsLikelyRawExtension(extLower) || IsJpegExtension(extLower);

        if (isRawOrJpeg)
        {
            const std::wstring baseLower = ToLowerCopy(PathWithoutExtensionView(path));
            if (! baseLower.empty())
            {
                if (emittedBasesLower.find(baseLower) != emittedBasesLower.end())
                {
                    continue;
                }
                emittedBasesLower.insert(baseLower);

                const auto it       = pairsByBaseLower.find(baseLower);
                const PairInfo info = (it != pairsByBaseLower.end()) ? it->second : PairInfo{};

                OtherItem item{};
                if (! info.rawPath.empty())
                {
                    item.primaryPath     = info.rawPath;
                    item.sidecarJpegPath = info.jpgPath;
                    item.isRaw           = true;
                }
                else
                {
                    item.primaryPath = ! info.jpgPath.empty() ? info.jpgPath : path;
                    item.isRaw       = false;
                }

                if (item.isRaw && ! item.sidecarJpegPath.empty())
                {
                    std::wstring left  = LeafNameFromPath(item.primaryPath);
                    std::wstring right = LeafNameFromPath(item.sidecarJpegPath);
                    if (left.empty())
                    {
                        left = item.primaryPath;
                    }
                    if (right.empty())
                    {
                        right = item.sidecarJpegPath;
                    }
                    item.label = std::format(L"{} | {}", left, right);
                }
                else
                {
                    item.label = LeafNameFromPath(item.primaryPath);
                    if (item.label.empty())
                    {
                        item.label = item.primaryPath;
                    }
                }

                _otherItems.push_back(std::move(item));
                continue;
            }
        }

        OtherItem item{};
        item.primaryPath = path;
        item.isRaw       = IsLikelyRawExtension(extLower) && ! IsWicImageExtension(extLower);
        item.label       = LeafNameFromPath(path);
        if (item.label.empty())
        {
            item.label = path;
        }
        _otherItems.push_back(std::move(item));
    }

    const wchar_t* focused         = context->focusedPath;
    const std::wstring focusedPath = focused ? focused : L"";

    _otherIndex = 0;
    if (! focusedPath.empty() && ! _otherItems.empty())
    {
        for (size_t i = 0; i < _otherItems.size(); ++i)
        {
            const OtherItem& item = _otherItems[i];
            if (item.primaryPath == focusedPath || EqualsIgnoreCase(item.primaryPath, focusedPath) ||
                (! item.sidecarJpegPath.empty() && (item.sidecarJpegPath == focusedPath || EqualsIgnoreCase(item.sidecarJpegPath, focusedPath))))
            {
                _otherIndex = i;
                break;
            }
        }
    }

    if (_otherIndex >= _otherItems.size())
    {
        _otherIndex = 0;
    }

    if (! _otherItems.empty())
    {
        _currentPath            = _otherItems[_otherIndex].primaryPath;
        _currentSidecarJpegPath = _otherItems[_otherIndex].sidecarJpegPath;
        _currentLabel           = _otherItems[_otherIndex].label;
    }
    else
    {
        _currentPath = focusedPath;
        _currentSidecarJpegPath.clear();
        _currentLabel = LeafNameFromPath(_currentPath);
        if (_currentLabel.empty())
        {
            _currentLabel = _currentPath;
        }
    }

    RefreshFileCombo(_hWnd.get());
    StartAsyncOpen(_hWnd.get(), _currentPath, false);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerImgRaw::Close() noexcept
{
    _hWnd.reset();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerImgRaw::SetTheme(const ViewerTheme* theme) noexcept
{
    if (theme == nullptr || theme->version != 2)
    {
        return E_INVALIDARG;
    }

    _theme    = *theme;
    _hasTheme = true;

    RequestViewerImgRawClassBackgroundColor(ColorRefFromArgb(_theme.backgroundArgb));
    ApplyPendingViewerImgRawClassBackgroundBrush(_hWnd.get());
    DiscardDirect2D();

    if (_hWnd)
    {
        ApplyTheme(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerImgRaw::SetCallback(IViewerCallback* callback, void* cookie) noexcept
{
    _callback       = callback;
    _callbackCookie = cookie;
    return S_OK;
}
