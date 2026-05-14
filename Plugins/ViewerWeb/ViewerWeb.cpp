#include "ViewerWeb.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cctype>
#include <cmath>
#include <cstddef>
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
#include <type_traits>
#include <utility>
#include <vector>

#include <commdlg.h>
#include <d2d1.h>
#include <dwmapi.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <uxtheme.h>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#pragma comment(lib, "d2d1")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "dwrite")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "uxtheme")
#pragma comment(lib, "WebView2Loader.dll.lib")

#include "DxUi/DxUi.Typography.h"
#include "Helpers.h"
#include "LocalizationManager.h"
#include "WindowMessages.h"
#include "WindowSizing.h"

#include "resource.h"

extern HINSTANCE g_hInstance;

namespace Typography = RedSalamander::DxUi::Typography;

using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::MakeDefaultThemePalette;
using RedSalamander::DxUi::MakeThemePaletteFromViewerTheme;

namespace
{
constexpr UINT kAsyncLoadCompleteMessage          = WndMsg::kViewerWebAsyncLoadComplete;
constexpr int kHeaderHeightDip                    = 28;
constexpr size_t kViewerComboPopupMaxVisibleItems = 8u;

static const int kViewerWebModuleAnchor = 0;

constexpr wchar_t kFileComboHostOriginalWndProcProp[] = L"RS.ViewerWeb.FileComboHostOriginalWndProc";
constexpr wchar_t kFileComboHostStateProp[]           = L"RS.ViewerWeb.FileComboHostState";

[[nodiscard]] WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* propName) noexcept
{
    return RedSalamander::Win32Callback::GetStoredWndProc(hwnd, propName);
}

[[nodiscard]] bool InstallWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp, WNDPROC hookWndProc) noexcept
{
    if (! hwnd || ! originalWndProcProp || ! hookWndProc)
    {
        return false;
    }

    if (GetStoredWndProc(hwnd, originalWndProcProp))
    {
        return true;
    }

    const auto originalWndProc = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(hwnd, GWLP_WNDPROC));
    if (! originalWndProc)
    {
        return false;
    }

    if (! RedSalamander::Win32Callback::SetPropNoThrow(hwnd, originalWndProcProp, reinterpret_cast<HANDLE>(originalWndProc)))
    {
        return false;
    }

    const auto previousWndProc =
        reinterpret_cast<WNDPROC>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(hookWndProc)));
    if (previousWndProc != originalWndProc)
    {
        RemovePropW(hwnd, originalWndProcProp);
        if (previousWndProc)
        {
            static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(previousWndProc)));
        }
        return false;
    }

    return true;
}

[[nodiscard]] LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* originalWndProcProp, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (const auto originalWndProc = GetStoredWndProc(hwnd, originalWndProcProp))
    {
        return RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK FileComboHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

void UnhookFileComboHostWindow(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    RemovePropW(hwnd, kFileComboHostStateProp);
    RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kFileComboHostOriginalWndProcProp, FileComboHostWndProc);
}

[[nodiscard]] bool MessageMayOpenWindowComboPopup(UINT msg, WPARAM wp) noexcept
{
    switch (msg)
    {
        case WM_LBUTTONDOWN:
        case WM_LBUTTONDBLCLK: return true;
        case WM_SYSKEYDOWN: return static_cast<UINT>(wp) == VK_DOWN || static_cast<UINT>(wp) == VK_UP;
        case WM_KEYDOWN:
        {
            const UINT vk = static_cast<UINT>(wp);
            return vk == VK_SPACE || vk == VK_RETURN || vk == VK_F4 || vk == VK_DOWN || vk == VK_UP;
        }
        default: return false;
    }
}

[[nodiscard]] int ComputeWindowComboPopupHeightPx(size_t itemCount, UINT dpi) noexcept
{
    const size_t visibleRows = std::max<size_t>(1u, std::min(itemCount, kViewerComboPopupMaxVisibleItems));
    const int popupHeightDip = 2 + 8 + (24 * static_cast<int>(visibleRows));
    return std::max(0, MulDiv(popupHeightDip, static_cast<int>(dpi), 96));
}

[[nodiscard]] size_t CountOwnerDrawMenuItems(HMENU menu) noexcept
{
    if (! menu)
    {
        return 0u;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return 0u;
    }

    size_t ownerDrawCount = 0u;
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_SUBMENU;
        if (GetMenuItemInfoW(menu, position, TRUE, &itemInfo) == 0)
        {
            continue;
        }

        if ((itemInfo.fType & MFT_OWNERDRAW) != 0)
        {
            ++ownerDrawCount;
        }

        if (itemInfo.hSubMenu)
        {
            ownerDrawCount += CountOwnerDrawMenuItems(itemInfo.hSubMenu);
        }
    }

    return ownerDrawCount;
}

[[nodiscard]] D2D1_COLOR_F D2DColor(COLORREF color) noexcept
{
    return D2D1::ColorF(
        static_cast<float>(GetRValue(color)) / 255.0f, static_cast<float>(GetGValue(color)) / 255.0f, static_cast<float>(GetBValue(color)) / 255.0f, 1.0f);
}

[[nodiscard]] float PixelsToDips(float pixels, UINT dpi) noexcept
{
    const UINT effectiveDpi = std::max<UINT>(dpi, USER_DEFAULT_SCREEN_DPI);
    return (pixels * static_cast<float>(USER_DEFAULT_SCREEN_DPI)) / static_cast<float>(effectiveDpi);
}

[[nodiscard]] ID2D1Factory* GetSharedStatusD2DFactory() noexcept
{
    static const auto resources = []() noexcept
    {
        wil::com_ptr<ID2D1Factory> factory;
        static_cast<void>(D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, factory.put()));
        return factory;
    }();

    return resources.get();
}

[[nodiscard]] bool DrawStatusMessageWithDirectWrite(HDC hdc, HWND hwnd, const RECT& rcPx, std::wstring_view text, COLORREF textColor) noexcept
{
    if (! hdc || ! hwnd || text.empty() || rcPx.right <= rcPx.left || rcPx.bottom <= rcPx.top ||
        text.size() > static_cast<size_t>((std::numeric_limits<UINT32>::max)()))
    {
        return false;
    }

    ID2D1Factory* const d2dFactory      = GetSharedStatusD2DFactory();
    IDWriteFactory* const dwriteFactory = Typography::GetSharedMeasurementFactory();
    if (! d2dFactory || ! dwriteFactory)
    {
        return false;
    }

    const UINT dpi = std::max<UINT>(GetDpiForWindow(hwnd), USER_DEFAULT_SCREEN_DPI);
    wil::com_ptr<IDWriteTextFormat> format;
    if (FAILED(Typography::CreateTextFormat(dwriteFactory, Typography::MakeUiTextSpec(12.0f), format.put())) || ! format)
    {
        return false;
    }
    static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING));
    static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));

    D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
    props.dpiX                          = static_cast<float>(dpi);
    props.dpiY                          = static_cast<float>(dpi);

    wil::com_ptr<ID2D1DCRenderTarget> target;
    if (FAILED(d2dFactory->CreateDCRenderTarget(&props, target.put())) || ! target)
    {
        return false;
    }
    if (FAILED(target->BindDC(hdc, &rcPx)))
    {
        return false;
    }

    wil::com_ptr<ID2D1SolidColorBrush> brush;
    if (FAILED(target->CreateSolidColorBrush(D2DColor(textColor), brush.put())) || ! brush)
    {
        return false;
    }

    const D2D1_RECT_F textRect =
        D2D1::RectF(0.0f, 0.0f, PixelsToDips(static_cast<float>(rcPx.right - rcPx.left), dpi), PixelsToDips(static_cast<float>(rcPx.bottom - rcPx.top), dpi));

    target->BeginDraw();
    target->DrawText(
        text.data(), static_cast<UINT32>(text.size()), format.get(), textRect, brush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP, DWRITE_MEASURING_MODE_NATURAL);
    const HRESULT endHr = target->EndDraw();
    if (FAILED(endHr))
    {
        Debug::Warning(L"ViewerWeb: DirectWrite status rendering EndDraw failed: 0x{:08X}", endHr);
        return false;
    }
    return true;
}

LRESULT CALLBACK FileComboHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* self = reinterpret_cast<ViewerWeb*>(GetPropW(hwnd, kFileComboHostStateProp));
    if (! self)
    {
        return CallStoredWndProc(hwnd, kFileComboHostOriginalWndProcProp, msg, wp, lp);
    }

    if (msg == WM_NCDESTROY)
    {
        const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kFileComboHostOriginalWndProcProp);
        RemovePropW(hwnd, kFileComboHostStateProp);
        RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kFileComboHostOriginalWndProcProp, FileComboHostWndProc);

        bool handled = false;
        static_cast<void>(self->HandleFileComboHostMessage(hwnd, msg, wp, lp, handled));

        return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    bool handled           = false;
    const LRESULT dxResult = self->HandleFileComboHostMessage(hwnd, msg, wp, lp, handled);
    if (handled)
    {
        return dxResult;
    }

    if (msg == WM_KEYDOWN && wp == VK_ESCAPE)
    {
        const HWND root = GetAncestor(hwnd, GA_ROOT);
        if (root)
        {
            PostMessageW(root, WM_CLOSE, 0, 0);
            return 0;
        }
    }

    return CallStoredWndProc(hwnd, kFileComboHostOriginalWndProcProp, msg, wp, lp);
}

[[maybe_unused]] [[nodiscard]] std::wstring KeyGlyphFromVirtualKey(UINT vk, HKL keyboardLayout) noexcept
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

[[nodiscard]] UINT VkFromScanCode(UINT scanCode, HKL keyboardLayout) noexcept
{
    if (! keyboardLayout)
    {
        return 0;
    }

    return MapVirtualKeyExW(scanCode, MAPVK_VSC_TO_VK_EX, keyboardLayout);
}

[[nodiscard]] UINT ZoomInVirtualKeyForLayout(HKL keyboardLayout) noexcept
{
    // Number row: key left of Backspace (US: =/+)
    constexpr UINT kScanCode = 0x0Du;
    if (const UINT vk = VkFromScanCode(kScanCode, keyboardLayout); vk != 0)
    {
        return vk;
    }

    return VK_OEM_PLUS;
}

[[nodiscard]] UINT ZoomOutVirtualKeyForLayout(HKL keyboardLayout) noexcept
{
    // Number row: key right of '0' (US: -/_)
    constexpr UINT kScanCode = 0x0Cu;
    if (const UINT vk = VkFromScanCode(kScanCode, keyboardLayout); vk != 0)
    {
        return vk;
    }

    return VK_OEM_MINUS;
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

struct HsvColor
{
    float hue = 0.0f;
    float sat = 0.0f;
    float val = 0.0f;
};

[[nodiscard]] HsvColor ColorToHSV(COLORREF color) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;

    const float maxValue = std::max({r, g, b});
    const float minValue = std::min({r, g, b});
    const float delta    = maxValue - minValue;

    HsvColor out{};
    out.val = maxValue;
    out.sat = maxValue <= 0.0f ? 0.0f : delta / maxValue;

    if (delta <= 0.0001f)
    {
        out.hue = 0.0f;
        return out;
    }

    if (maxValue == r)
    {
        out.hue = 60.0f * std::fmod(((g - b) / delta), 6.0f);
    }
    else if (maxValue == g)
    {
        out.hue = 60.0f * (((b - r) / delta) + 2.0f);
    }
    else
    {
        out.hue = 60.0f * (((r - g) / delta) + 4.0f);
    }

    if (out.hue < 0.0f)
    {
        out.hue += 360.0f;
    }

    return out;
}

struct JsonTokenColors
{
    COLORREF key          = RGB(0, 0, 0);
    COLORREF stringValue  = RGB(0, 0, 0);
    COLORREF numberValue  = RGB(0, 0, 0);
    COLORREF literalValue = RGB(0, 0, 0);
};

[[nodiscard]] COLORREF ThemeAwareSemanticColor(
    float hue, float saturation, float value, COLORREF accent, COLORREF fg, uint8_t accentAlpha, uint8_t fgAlpha) noexcept
{
    const COLORREF semantic   = ColorFromHSV(hue, saturation, value);
    const COLORREF harmonized = BlendColor(semantic, accent, accentAlpha);
    return BlendColor(harmonized, fg, fgAlpha);
}

[[nodiscard]] JsonTokenColors BuildJsonTokenColors(COLORREF accent, COLORREF fg, bool darkMode) noexcept
{
    const HsvColor accentHsv = ColorToHSV(accent);
    const float keyHue       = accentHsv.sat > 0.05f ? accentHsv.hue : 195.0f;
    const float keySat       = std::clamp(std::max(accentHsv.sat, darkMode ? 0.60f : 0.72f), 0.0f, 1.0f);
    const float keyVal       = darkMode ? std::max(accentHsv.val, 0.95f) : std::clamp(std::max(accentHsv.val * 0.62f, 0.46f), 0.46f, 0.70f);
    const COLORREF keyBase   = ColorFromHSV(keyHue, keySat, keyVal);

    return {
        .key = BlendColor(keyBase, fg, darkMode ? 12u : 18u),
        .stringValue =
            ThemeAwareSemanticColor(145.0f, darkMode ? 0.58f : 0.76f, darkMode ? 0.94f : 0.44f, accent, fg, darkMode ? 24u : 18u, darkMode ? 16u : 18u),
        .numberValue =
            ThemeAwareSemanticColor(32.0f, darkMode ? 0.78f : 0.82f, darkMode ? 0.98f : 0.42f, accent, fg, darkMode ? 20u : 14u, darkMode ? 12u : 14u),
        .literalValue =
            ThemeAwareSemanticColor(282.0f, darkMode ? 0.64f : 0.74f, darkMode ? 0.96f : 0.46f, accent, fg, darkMode ? 18u : 12u, darkMode ? 12u : 16u),
    };
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

// Forward declarations for file-scope helpers defined later in this file.
bool CopyUnicodeTextToClipboard(HWND hwnd, std::wstring_view text) noexcept;
[[nodiscard]] bool IsProbablyWin32Path(std::wstring_view path) noexcept;
std::optional<std::filesystem::path> ShowSaveAsDialog(HWND hwnd, std::wstring_view suggestedFileName) noexcept;
[[nodiscard]] std::wstring EscapeJavaScriptString(std::wstring_view text) noexcept;
[[nodiscard]] std::string EscapeJavaScriptStringUtf8(std::string_view text) noexcept;

struct ViewerWebClassBackgroundBrushState
{
    ViewerWebClassBackgroundBrushState()                                                     = default;
    ViewerWebClassBackgroundBrushState(const ViewerWebClassBackgroundBrushState&)            = delete;
    ViewerWebClassBackgroundBrushState& operator=(const ViewerWebClassBackgroundBrushState&) = delete;
    ViewerWebClassBackgroundBrushState(ViewerWebClassBackgroundBrushState&&)                 = delete;
    ViewerWebClassBackgroundBrushState& operator=(ViewerWebClassBackgroundBrushState&&)      = delete;

    wil::unique_hbrush activeBrush;
    COLORREF activeColor = CLR_INVALID;

    wil::unique_hbrush pendingBrush;
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

std::string ResourceBytesString(HINSTANCE hinst, UINT id) noexcept
{
    std::vector<std::byte> bytes;
    if (! Localization::LoadResourceBytes(hinst, MAKEINTRESOURCEW(id), RT_RCDATA, bytes))
    {
        return {};
    }

    return {reinterpret_cast<const char*>(bytes.data()), bytes.size()};
}

// ----- Shared WebView2 Environment -----
// A DLL-global environment shared across all viewer instances. Creating a
// WebView2 environment is the most expensive operation (spawns browser
// processes). Sharing it means only the first viewer pays the cost.

struct SharedEnvironmentState
{
    wil::com_ptr<ICoreWebView2Environment> environment;
    bool createInProgress = false;
    // Pending consumers waiting for the first environment creation to complete.
    struct PendingConsumer
    {
        ViewerWeb* viewer = nullptr;
        HWND hwnd         = nullptr;
    };
    std::vector<PendingConsumer> pendingConsumers;
};

SharedEnvironmentState g_sharedEnvironment;

std::wstring GetWebView2UserDataFolder() noexcept
{
    wchar_t appData[MAX_PATH]{};
    if (FAILED(SHGetFolderPathW(nullptr, CSIDL_LOCAL_APPDATA, nullptr, 0, appData)))
    {
        return {};
    }
    return std::format(L"{}\\RedSalamander\\WebView2UserData", appData);
}

void ResetSharedEnvironmentImpl() noexcept
{
    g_sharedEnvironment.pendingConsumers.clear();
    g_sharedEnvironment.createInProgress = false;
    g_sharedEnvironment.environment.reset();
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

constexpr std::wstring_view kInternalDocumentOrigin = L"https://viewer.redsalamander.invalid";
constexpr std::wstring_view kInternalDocumentFilter = L"https://viewer.redsalamander.invalid/*";

[[nodiscard]] std::wstring_view InternalDocumentKindSegment(ViewerWebKind kind) noexcept
{
    switch (kind)
    {
        case ViewerWebKind::Json: return L"json";
        case ViewerWebKind::Markdown: return L"markdown";
        case ViewerWebKind::Web:
        default: return L"web";
    }
}

[[nodiscard]] std::wstring BuildInternalDocumentUrl(ViewerWebKind kind, uint64_t requestId)
{
    return std::format(L"{}/{}/{}.html", kInternalDocumentOrigin, InternalDocumentKindSegment(kind), requestId);
}

[[nodiscard]] bool IsInternalDocumentUrl(std::wstring_view url) noexcept
{
    return OrdinalString::StartsWithNoCase(url, kInternalDocumentOrigin);
}

[[nodiscard]] std::string_view TrimAsciiWhitespace(std::string_view text) noexcept
{
    while (! text.empty() && std::isspace(static_cast<unsigned char>(text.front())) != 0)
    {
        text.remove_prefix(1);
    }

    while (! text.empty() && std::isspace(static_cast<unsigned char>(text.back())) != 0)
    {
        text.remove_suffix(1);
    }

    return text;
}

[[nodiscard]] bool IsJsonLinesPath(std::wstring_view path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    const std::wstring extension = std::filesystem::path(path).extension().wstring();
    return OrdinalString::EqualsNoCase(extension, L".jsonl") || OrdinalString::EqualsNoCase(extension, L".ndjson");
}

[[nodiscard]] std::string CssRgb(COLORREF c)
{
    return std::format("rgb({},{},{})", GetRValue(c), GetGValue(c), GetBValue(c));
}

struct ScrollbarColors
{
    COLORREF track      = RGB(0, 0, 0);
    COLORREF thumb      = RGB(0, 0, 0);
    COLORREF thumbHover = RGB(0, 0, 0);
    COLORREF corner     = RGB(0, 0, 0);
};

[[nodiscard]] ScrollbarColors BuildScrollbarColors(COLORREF bg, COLORREF fg, COLORREF accent, bool darkMode) noexcept
{
    const COLORREF thumbBase = BlendColor(bg, accent, darkMode ? 52u : 36u);
    return {
        .track      = BlendColor(bg, fg, darkMode ? 20u : 12u),
        .thumb      = BlendColor(thumbBase, fg, darkMode ? 92u : 120u),
        .thumbHover = BlendColor(thumbBase, fg, darkMode ? 128u : 152u),
        .corner     = BlendColor(bg, fg, darkMode ? 16u : 8u),
    };
}

struct JsonLinesEntry
{
    size_t lineNumber = 0;
    std::string timestampText;
    std::string levelText;
    std::string categoryText;
    std::string messageText;
    std::string summaryText;
    std::string prettyJson;
};

[[nodiscard]] std::string JsonStringFromScalar(yyjson_val* value)
{
    if (! value)
    {
        return {};
    }
    if (yyjson_is_str(value))
    {
        const char* s = yyjson_get_str(value);
        return s ? std::string(s) : std::string{};
    }
    if (yyjson_is_int(value))
    {
        return std::to_string(yyjson_get_int(value));
    }
    if (yyjson_is_uint(value))
    {
        return std::to_string(yyjson_get_uint(value));
    }
    if (yyjson_is_real(value))
    {
        return std::format("{}", yyjson_get_real(value));
    }
    if (yyjson_is_bool(value))
    {
        return yyjson_get_bool(value) ? "true" : "false";
    }
    if (yyjson_is_null(value))
    {
        return "null";
    }
    return {};
}

[[nodiscard]] std::string DescribeJsonValue(yyjson_val* value)
{
    if (! value)
    {
        return "JSON value";
    }
    if (yyjson_is_obj(value))
    {
        const size_t fieldCount = yyjson_obj_size(value);
        return std::format("Object with {} field{}", fieldCount, fieldCount == 1 ? "" : "s");
    }
    if (yyjson_is_arr(value))
    {
        const size_t itemCount = yyjson_arr_size(value);
        return std::format("Array with {} item{}", itemCount, itemCount == 1 ? "" : "s");
    }
    if (yyjson_is_str(value))
    {
        return "String value";
    }
    if (yyjson_is_bool(value))
    {
        return std::format("Boolean {}", yyjson_get_bool(value) ? "true" : "false");
    }
    if (yyjson_is_null(value))
    {
        return "Null value";
    }
    if (yyjson_is_num(value))
    {
        return std::format("Numeric {}", JsonStringFromScalar(value));
    }
    return "JSON value";
}

[[nodiscard]] std::string ReadJsonObjectSummaryValue(yyjson_val* value, std::initializer_list<const char*> keys)
{
    if (! yyjson_is_obj(value))
    {
        return {};
    }

    for (const char* key : keys)
    {
        const std::string text = JsonStringFromScalar(yyjson_obj_get(value, key));
        if (! text.empty())
        {
            return text;
        }
    }

    return {};
}

[[nodiscard]] bool TryParseJsonLinesEntries(std::string_view textUtf8, bool allowSingleEntry, std::vector<JsonLinesEntry>& outEntries)
{
    outEntries.clear();

    size_t lineNumber = 1;
    size_t offset     = 0;
    while (offset < textUtf8.size())
    {
        size_t lineEnd = textUtf8.find('\n', offset);
        if (lineEnd == std::string::npos)
        {
            lineEnd = textUtf8.size();
        }

        std::string_view line(textUtf8.data() + offset, lineEnd - offset);
        if (! line.empty() && line.back() == '\r')
        {
            line.remove_suffix(1);
        }
        line = TrimAsciiWhitespace(line);

        if (! line.empty())
        {
            std::string mutableLine(line);
            yyjson_read_err lineErr{};
            unique_yyjson_doc lineDoc(yyjson_read_opts(mutableLine.data(), mutableLine.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &lineErr));
            if (! lineDoc)
            {
                outEntries.clear();
                return false;
            }

            yyjson_val* root = yyjson_doc_get_root(lineDoc.get());
            if (! root)
            {
                outEntries.clear();
                return false;
            }

            size_t prettyLen = 0;
            unique_malloc_string pretty(yyjson_write_opts(lineDoc.get(), YYJSON_WRITE_PRETTY | YYJSON_WRITE_ESCAPE_UNICODE, nullptr, &prettyLen, nullptr));
            if (! pretty)
            {
                outEntries.clear();
                return false;
            }

            JsonLinesEntry entry{};
            entry.lineNumber = lineNumber;
            entry.prettyJson.assign(pretty.get(), prettyLen);
            entry.timestampText = ReadJsonObjectSummaryValue(root, {"ts", "timestamp", "@timestamp", "time"});
            entry.levelText     = ReadJsonObjectSummaryValue(root, {"level", "severity", "lvl", "type"});
            entry.categoryText  = ReadJsonObjectSummaryValue(root, {"category", "op", "operation", "event", "logger"});
            entry.messageText   = ReadJsonObjectSummaryValue(root, {"message", "msg", "text", "summary", "description"});
            entry.summaryText   = DescribeJsonValue(root);
            if (entry.messageText.empty())
            {
                const std::string nameText = ReadJsonObjectSummaryValue(root, {"name", "srcLeaf", "dstLeaf", "path"});
                if (! nameText.empty())
                {
                    entry.messageText = nameText;
                }
            }

            outEntries.push_back(std::move(entry));
        }

        offset = lineEnd < textUtf8.size() ? lineEnd + 1 : lineEnd;
        ++lineNumber;
    }

    const size_t minimumEntries = allowSingleEntry ? 1u : 2u;
    if (outEntries.size() < minimumEntries)
    {
        outEntries.clear();
        return false;
    }

    return true;
}

// Common theme helper JS shared by all internal HTML templates.
constexpr char kCommonThemeJs[] =
    "function parseRgb(s){const m=/rgb\\((\\d+),(\\d+),(\\d+)\\)/.exec(s.replace(/\\s+/g,''));return m?{r:+m[1],g:+m[2],b:+m[3]}:{r:0,g:0,b:0};}"
    "function rgb(c){return `rgb(${c.r},${c.g},${c.b})`;}"
    "function clamp(v,min,max){return Math.min(max,Math.max(min,v));}"
    "function blend(u,o,a){const inv=255-a;return {r:Math.round((u.r*inv+o.r*a)/255),g:Math.round((u.g*inv+o.g*a)/255),b:Math.round((u.b*inv+o.b*a)/255)};}"
    "function luma(c){return (c.r*299+c.g*587+c.b*114)/1000;}"
    "function hsv(h,s,v){const hue=((h%360)+360)%360;const c=v*s;const x=c*(1-Math.abs((hue/60)%2-1));const m=v-c;let rf=0,gf=0,bf=0;"
    "if(hue<60){rf=c;gf=x;}else if(hue<120){rf=x;gf=c;}else if(hue<180){gf=c;bf=x;}else if(hue<240){gf=x;bf=c;}else if(hue<300){rf=x;bf=c;}else{rf=c;bf=x;}"
    "return {r:Math.round((rf+m)*255),g:Math.round((gf+m)*255),b:Math.round((bf+m)*255)};}"
    "function rgbToHsv(c){const r=c.r/255,g=c.g/255,b=c.b/255;const max=Math.max(r,g,b),min=Math.min(r,g,b),delta=max-min;let h=0;"
    "if(delta>0){if(max===r){h=60*(((g-b)/delta)%6);}else if(max===g){h=60*(((b-r)/delta)+2);}else{h=60*(((r-g)/delta)+4);}if(h<0){h+=360;}}"
    "return {h,s:max<=0?0:delta/max,v:max};}"
    "function setScrollbarVars(r,bg,fg,acc){const dark=luma(bg)<128;const thumbBase=blend(bg,acc,dark?52:36);"
    "r.setProperty('--rs-scroll-track',rgb(blend(bg,fg,dark?20:12)));r.setProperty('--rs-scroll-thumb',rgb(blend(thumbBase,fg,dark?92:120)));"
    "r.setProperty('--rs-scroll-thumb-hover',rgb(blend(thumbBase,fg,dark?128:152)));r.setProperty('--rs-scroll-corner',rgb(blend(bg,fg,dark?16:8)));}"
    "function setJsonTokenVars(r,bg,fg,acc){const dark=luma(bg)<128;"
    "const accentHsv=rgbToHsv(acc);const keyHue=accentHsv.s>0.05?accentHsv.h:195;const keySat=clamp(Math.max(accentHsv.s,dark?0.60:0.72),0,1);"
    "const keyVal=dark?Math.max(accentHsv.v,0.95):clamp(Math.max(accentHsv.v*0.62,0.46),0.46,0.70);const key=blend(hsv(keyHue,keySat,keyVal),fg,dark?12:18);"
    "const str=blend(blend(hsv(145,dark?0.58:0.76,dark?0.94:0.44),acc,dark?24:18),fg,dark?16:18);"
    "const num=blend(blend(hsv(32,dark?0.78:0.82,dark?0.98:0.42),acc,dark?20:14),fg,dark?12:14);"
    "const lit=blend(blend(hsv(282,dark?0.64:0.74,dark?0.96:0.46),acc,dark?18:12),fg,dark?12:16);"
    "r.setProperty('--rs-key',rgb(key));r.setProperty('--rs-string',rgb(str));r.setProperty('--rs-number',rgb(num));r.setProperty('--rs-literal',rgb(lit));}";

constexpr char kCommonScrollbarCss[] =
    "html{scrollbar-color:var(--rs-scroll-thumb) var(--rs-scroll-track);}"
    "*{scrollbar-color:var(--rs-scroll-thumb) var(--rs-scroll-track);scrollbar-width:thin;}"
    "*::-webkit-scrollbar{width:12px;height:12px;background:var(--rs-scroll-track);}"
    "*::-webkit-scrollbar-track{background:var(--rs-scroll-track);}"
    "*::-webkit-scrollbar-thumb{background:var(--rs-scroll-thumb);border-radius:999px;border:3px solid var(--rs-scroll-track);}"
    "*::-webkit-scrollbar-thumb:hover{background:var(--rs-scroll-thumb-hover);}"
    "*::-webkit-scrollbar-button{display:none;width:0;height:0;}"
    "*::-webkit-scrollbar-corner{background:var(--rs-scroll-corner);}";

[[nodiscard]] std::string BuildJsonLinesHtml(const std::vector<JsonLinesEntry>& entries,
                                             std::string_view highlightJs,
                                             std::string_view themeObj,
                                             COLORREF bg,
                                             COLORREF fg,
                                             COLORREF selBg,
                                             COLORREF selFg,
                                             COLORREF accent,
                                             bool darkMode)
{
    const COLORREF codeBg             = BlendColor(bg, fg, darkMode ? 18u : 8u);
    const COLORREF cardBg             = BlendColor(bg, fg, darkMode ? 12u : 5u);
    const COLORREF cardBgOpen         = BlendColor(bg, fg, darkMode ? 18u : 10u);
    const COLORREF border             = BlendColor(bg, fg, darkMode ? 36u : 58u);
    const COLORREF mutedFg            = BlendColor(bg, fg, 140u);
    const JsonTokenColors tokenColors = BuildJsonTokenColors(accent, fg, darkMode);
    const ScrollbarColors scrollbar   = BuildScrollbarColors(bg, fg, accent, darkMode);

    std::string entriesJs;
    entriesJs.reserve(entries.size() * 256u);
    entriesJs.push_back('[');
    for (size_t i = 0; i < entries.size(); ++i)
    {
        const JsonLinesEntry& entry = entries[i];
        if (i != 0)
        {
            entriesJs.push_back(',');
        }

        entriesJs += "{line:";
        entriesJs += std::to_string(entry.lineNumber);
        entriesJs += ",ts:'";
        entriesJs += EscapeJavaScriptStringUtf8(entry.timestampText);
        entriesJs += "',level:'";
        entriesJs += EscapeJavaScriptStringUtf8(entry.levelText);
        entriesJs += "',category:'";
        entriesJs += EscapeJavaScriptStringUtf8(entry.categoryText);
        entriesJs += "',message:'";
        entriesJs += EscapeJavaScriptStringUtf8(entry.messageText);
        entriesJs += "',summary:'";
        entriesJs += EscapeJavaScriptStringUtf8(entry.summaryText);
        entriesJs += "',json:'";
        entriesJs += EscapeJavaScriptStringUtf8(entry.prettyJson);
        entriesJs += "'}";
    }
    entriesJs.push_back(']');

    std::string html;
    html.reserve(highlightJs.size() + entriesJs.size() + 16384);
    html += "<!doctype html><html><head><meta charset=\"utf-8\">";
    html += "<style>";
    html += ":root{--rs-bg:" + CssRgb(bg) + ";--rs-fg:" + CssRgb(fg) + ";--rs-sel-bg:" + CssRgb(selBg) + ";--rs-sel-fg:" + CssRgb(selFg) +
            ";--rs-accent:" + CssRgb(accent) + ";--rs-code-bg:" + CssRgb(codeBg) + ";--rs-card-bg:" + CssRgb(cardBg) +
            ";--rs-card-bg-open:" + CssRgb(cardBgOpen) + ";--rs-border:" + CssRgb(border) + ";--rs-muted-fg:" + CssRgb(mutedFg) +
            ";--rs-key:" + CssRgb(tokenColors.key) + ";--rs-string:" + CssRgb(tokenColors.stringValue) + ";--rs-number:" + CssRgb(tokenColors.numberValue) +
            ";--rs-literal:" + CssRgb(tokenColors.literalValue) + ";--rs-scroll-track:" + CssRgb(scrollbar.track) +
            ";--rs-scroll-thumb:" + CssRgb(scrollbar.thumb) + ";--rs-scroll-thumb-hover:" + CssRgb(scrollbar.thumbHover) +
            ";--rs-scroll-corner:" + CssRgb(scrollbar.corner) + ";}";
    html += "html,body{height:100%;margin:0;}*,*::before,*::after{box-sizing:border-box;}body{background:var(--rs-bg);color:var(--rs-fg);"
            "font-family:\"Segoe UI Variable Text\",\"Segoe UI Variable Small\",\"Segoe UI\",sans-serif;overflow:hidden;}";
    html += kCommonScrollbarCss;
    html += "::selection{background:var(--rs-sel-bg);color:var(--rs-sel-fg);}.rs-shell{height:100%;min-height:0;display:flex;flex-direction:column;}";
    html += ".rs-toolbar{flex:none;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 14px;"
            "background:linear-gradient(180deg,var(--rs-bg),var(--rs-card-bg));border-bottom:1px solid var(--rs-border);}";
    html += ".rs-toolbar-left{display:flex;align-items:center;gap:10px;min-width:0;}.rs-toolbar-title{font-size:13px;font-weight:600;letter-spacing:0.02em;}"
            ".rs-toolbar-meta{font-size:12px;color:var(--rs-muted-fg);white-space:nowrap;}.rs-toolbar-actions{display:flex;align-items:center;gap:8px;}";
    html += ".rs-btn,.rs-pill{border:1px solid var(--rs-border);background:var(--rs-card-bg);color:var(--rs-fg);border-radius:999px;}"
            ".rs-btn{padding:6px 10px;font:inherit;cursor:pointer;}.rs-btn:hover{background:var(--rs-card-bg-open);}."
            "rs-pill{padding:4px 9px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;}";
    html += ".rs-list{flex:1 1 auto;min-height:0;padding:14px;display:flex;flex-direction:column;gap:10px;overflow:auto;"
            "overscroll-behavior:contain;scrollbar-gutter:stable both-edges;}.rs-entry{flex:none;border:1px solid var(--rs-border);"
            "border-radius:10px;background:var(--rs-card-bg);overflow:hidden;}.rs-entry.is-open{background:var(--rs-card-bg-open);}";
    html += ".rs-entry-summary{width:100%;display:flex;align-items:center;gap:8px;padding:10px 12px;border:0;background:transparent;color:inherit;"
            "cursor:pointer;text-align:left;font:inherit;line-height:1.25;appearance:none;}.rs-entry-summary:hover{background:rgba(255,255,255,0.02);}"
            ".rs-entry-summary:focus-visible{outline:2px solid var(--rs-accent);outline-offset:-2px;}.rs-entry-summary>*{min-width:0;}"
            ".rs-entry-summary::before{content:'+';display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;"
            "border-radius:999px;background:var(--rs-code-bg);color:var(--rs-muted-fg);font-size:12px;font-weight:700;flex:none;}";
    html += ".rs-entry.is-open .rs-entry-summary::before{content:'-';color:var(--rs-accent);}.rs-line-pill{min-width:56px;text-align:center;}";
    html += ".rs-badge{padding:4px 8px;border-radius:999px;font-size:11px;font-weight:700;border:1px solid var(--rs-border);background:var(--rs-code-bg);"
            "color:var(--rs-muted-fg);text-transform:uppercase;letter-spacing:0.04em;}";
    html += ".rs-badge[data-kind='level'][data-level='error']{background:rgba(190,64,64,0.22);color:rgb(255,210,210);}."
            "rs-badge[data-kind='level'][data-level='warning']{background:rgba(190,140,40,0.24);color:rgb(255,230,170);}."
            "rs-badge[data-kind='level'][data-level='info']{background:rgba(64,116,190,0.22);color:rgb(210,225,255);}."
            "rs-badge[data-kind='level'][data-level='debug']{background:rgba(120,120,140,0.22);color:rgb(220,220,230);}";
    html += ".rs-summary-text{min-width:0;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:13px;}"
            ".rs-entry-body{padding:0 12px 12px 42px;}.rs-entry-body[hidden]{display:none;}.rs-entry-body pre{margin:0;max-width:100%;"
            "background:var(--rs-code-bg);border:1px solid var(--rs-border);padding:12px;"
            "overflow:auto;border-radius:8px;}.rs-entry-body code{font-family:Consolas,ui-monospace,monospace;font-size:13px;line-height:1.45;}";
    html += ".hljs{background:transparent;}.hljs-attr{color:var(--rs-key);} .hljs-string{color:var(--rs-string);} .hljs-number{color:var(--rs-number);} "
            ".hljs-literal{color:var(--rs-literal);} .hljs-punctuation,.hljs-brace{color:var(--rs-muted-fg);} .hljs-comment{opacity:0.8;}";
    html += "@media (max-width:720px){.rs-toolbar{flex-direction:column;align-items:stretch;}.rs-toolbar-actions{justify-content:flex-end;}.rs-entry-summary"
            "{flex-wrap:wrap;}.rs-entry-body{padding-left:12px;}}";
    html += "</style></head><body><div class=\"rs-shell\"><div class=\"rs-toolbar\">";
    html += "<div class=\"rs-toolbar-left\"><span class=\"rs-pill\">JSONL</span><span class=\"rs-toolbar-title\">Structured log view</span>"
            "<span id=\"summary\" class=\"rs-toolbar-meta\"></span></div>";
    html += "<div class=\"rs-toolbar-actions\"><button id=\"expandAll\" type=\"button\" class=\"rs-btn\">Expand all</button>"
            "<button id=\"collapseAll\" type=\"button\" class=\"rs-btn\">Collapse all</button></div>";
    html += "</div><div id=\"list\" class=\"rs-list\"></div></div>";
    html += "<script>";
    html.append(highlightJs);
    html += "</script><script>";
    html += "(() => {const initialTheme=" + std::string(themeObj) + ";";
    html += kCommonThemeJs;
    html += "function applyTheme(t){const r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);"
            "r.setProperty('--rs-sel-bg',t.selBg);r.setProperty('--rs-sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);"
            "const bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const dark=luma(bg)<128;"
            "r.setProperty('--rs-code-bg',rgb(blend(bg,fg,dark?18:8)));r.setProperty('--rs-card-bg',rgb(blend(bg,fg,dark?12:5)));"
            "r.setProperty('--rs-card-bg-open',rgb(blend(bg,fg,dark?18:10)));r.setProperty('--rs-border',rgb(blend(bg,fg,dark?36:58)));"
            "r.setProperty('--rs-muted-fg',rgb(blend(bg,fg,140)));setJsonTokenVars(r,bg,fg,acc);setScrollbarVars(r,bg,fg,acc);}";
    html += "const entries=" + entriesJs +
            ";const list=document.getElementById('list');"
            "document.getElementById('summary').textContent=`${entries.length} record${entries.length===1?'':'s'}`;";
    html += "function makeBadge(text,kind,level){if(!text){return null;}const el=document.createElement('span');el.className='rs-badge';"
            "el.textContent=text;el.dataset.kind=kind;if(level){el.dataset.level=String(level).toLowerCase();}return el;}";
    html += "function renderCode(code,text){code.textContent=text;try{hljs.highlightElement(code);}catch(e){}}";
    html += "function setEntryOpen(entryEl,open,scrollIntoView){if(!entryEl){return;}entryEl.classList.toggle('is-open',open);"
            "const header=entryEl.querySelector('.rs-entry-summary');const body=entryEl.querySelector('.rs-entry-body');"
            "if(header){header.setAttribute('aria-expanded',open?'true':'false');}if(body){body.hidden=!open;}"
            "if(open&&typeof entryEl._ensureRendered==='function'){entryEl._ensureRendered();if(scrollIntoView){entryEl.scrollIntoView({block:'nearest'});}}}";
    html += "function makeEntry(entry,index){const entryEl=document.createElement('article');entryEl.className='rs-entry';"
            "const "
            "summary=document.createElement('button');summary.type='button';summary.className='rs-entry-summary';summary.setAttribute('aria-expanded','false');"
            "const linePill=document.createElement('span');linePill.className='rs-pill "
            "rs-line-pill';linePill.textContent=`#${entry.line}`;summary.appendChild(linePill);"
            "if(entry.ts){const ts=document.createElement('span');ts.className='rs-toolbar-meta';ts.textContent=entry.ts;summary.appendChild(ts);}"
            "const levelBadge=makeBadge(entry.level,'level',entry.level);if(levelBadge){summary.appendChild(levelBadge);}const "
            "categoryBadge=makeBadge(entry.category,'category','');"
            "if(categoryBadge){summary.appendChild(categoryBadge);}const text=document.createElement('span');text.className='rs-summary-text';"
            "text.textContent=entry.message||entry.summary||'JSON value';summary.appendChild(text);entryEl.appendChild(summary);"
            "const body=document.createElement('div');body.className='rs-entry-body';body.hidden=true;const pre=document.createElement('pre');const "
            "code=document.createElement('code');"
            "code.className='language-json';pre.appendChild(code);body.appendChild(pre);entryEl.appendChild(body);let rendered=false;"
            "entryEl._ensureRendered=()=>{if(rendered){return;}renderCode(code,entry.json);rendered=true;};"
            "summary.addEventListener('click',()=>setEntryOpen(entryEl,!entryEl.classList.contains('is-open'),true));"
            "if(index<2){setEntryOpen(entryEl,true,false);}return entryEl;}";
    html += "function renderList(){const frag=document.createDocumentFragment();entries.forEach((entry,index)=>frag.appendChild(makeEntry(entry,index)));"
            "list.replaceChildren(frag);}function expandAll(){document.querySelectorAll('.rs-entry').forEach((entryEl)=>setEntryOpen(entryEl,true,false));}"
            "function collapseAll(){document.querySelectorAll('.rs-entry').forEach((entryEl)=>setEntryOpen(entryEl,false,false));}";
    html +=
        "document.getElementById('expandAll').addEventListener('click',expandAll);document.getElementById('collapseAll').addEventListener('click',collapseAll);"
        "window.RS={applyTheme:applyTheme,expandAll:expandAll,collapseAll:collapseAll};applyTheme(initialTheme);renderList();})();";
    html += "</script></body></html>";
    return html;
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
            "description": "Pretty highlighted text, interactive tree view, or JSONL log cards.",
            "default": "pretty",
            "options": [
                { "value": "pretty", "label": "Pretty" },
                { "value": "tree", "label": "Tree" },
                { "value": "jsonl", "label": "JSONL" }
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

[[nodiscard]] const char* GetViewerWebStaticConfigurationSchemaImpl(ViewerWebKind kind) noexcept
{
    switch (kind)
    {
        case ViewerWebKind::Json: return kViewerJsonSchemaJson;
        case ViewerWebKind::Markdown: return kViewerMarkdownSchemaJson;
        case ViewerWebKind::Web:
        default: return kViewerWebSchemaJson;
    }
}

// ----- Cached Resource Helpers -----
// JS/CSS resource data is immutable after load. The helpers below avoid
// repeated allocations by caching derived data (base64-encoded icons, CSS with
// icons inlined, and the common theme helper JS shared by all three HTML
// templates).

std::string_view GetHighlightJs() noexcept
{
    static const std::string s = ResourceBytesString(g_hInstance, IDR_VIEWERWEB_HIGHLIGHT_JS);
    return {s.data(), s.size()};
}

std::string_view GetMarkdownItJs() noexcept
{
    static const std::string s = ResourceBytesString(g_hInstance, IDR_VIEWERWEB_MARKDOWNIT_JS);
    return {s.data(), s.size()};
}

std::string_view GetJsonEditorJs() noexcept
{
    static const std::string s = ResourceBytesString(g_hInstance, IDR_VIEWERWEB_JSONEDITOR_JS);
    return {s.data(), s.size()};
}

// Returns the jsoneditor CSS with SVG icon URLs already inlined as data-URIs.
const std::string& GetJsonEditorCssWithIcons() noexcept
{
    static const std::string s = []
    {
        auto base64Encode = [](std::string_view bytes) noexcept -> std::string
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
                const uint32_t n =
                    (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i])) << 16) | (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i + 1])) << 8);
                out.push_back(kTable[(n >> 18) & 0x3F]);
                out.push_back(kTable[(n >> 12) & 0x3F]);
                out.push_back(kTable[(n >> 6) & 0x3F]);
                out.push_back('=');
            }
            return out;
        };

        const std::string iconsSvg = ResourceBytesString(g_hInstance, IDR_VIEWERWEB_JSONEDITOR_ICONS_SVG);
        const std::string iconsUrl = std::string("data:image/svg+xml;base64,") + base64Encode(iconsSvg);

        std::string css = ResourceBytesString(g_hInstance, IDR_VIEWERWEB_JSONEDITOR_CSS);

        auto replaceAll = [](std::string& text, std::string_view needle, std::string_view replacement) noexcept
        {
            if (needle.empty())
                return;
            size_t pos = 0;
            while ((pos = text.find(needle, pos)) != std::string::npos)
            {
                text.replace(pos, needle.size(), replacement);
                pos += replacement.size();
            }
        };

        replaceAll(css, "./img/jsoneditor-icons.svg", iconsUrl);
        replaceAll(css, "img/jsoneditor-icons.svg", iconsUrl);
        return css;
    }();
    return s;
}

} // namespace

const char* GetViewerWebStaticConfigurationSchema(ViewerWebKind kind) noexcept
{
    return GetViewerWebStaticConfigurationSchemaImpl(kind);
}

LRESULT ViewerWeb::HandleFileComboHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept
{
    const bool popupWasOpen      = _fileComboControl && _fileComboControl->DebugIsPopupOpen();
    const bool preExpandForPopup = ! popupWasOpen && _fileComboControl && MessageMayOpenWindowComboPopup(msg, wp);
    if (preExpandForPopup)
    {
        _fileComboHostPreExpandPopup = true;
        if (_hWnd)
        {
            Layout(_hWnd.get());
        }
    }

    const LRESULT dxResult = _fileComboHost.HandleMessage(hwnd, msg, wp, lp, handled);
    if (msg == WM_NCDESTROY)
    {
        handled = true;
        _fileComboHost.ReleaseMouseCapture();
        _fileComboControl            = nullptr;
        _fileComboHostPreExpandPopup = false;
        _hFileComboHost.release();
        return dxResult;
    }

    const bool popupIsOpen = _fileComboControl && _fileComboControl->DebugIsPopupOpen();
    if (popupIsOpen != popupWasOpen || (preExpandForPopup && ! popupIsOpen))
    {
        _fileComboHostPreExpandPopup = false;
        if (_hWnd)
        {
            Layout(_hWnd.get());
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    }
    return dxResult;
}

void ViewerWeb::OnCreate(HWND hwnd)
{
    const DWORD comboHostStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | SS_NOTIFY;
    _hFileComboHost.reset(CreateWindowExW(
        0, L"Static", L"", comboHostStyle, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_VIEWERWEB_FILE_COMBO)), g_hInstance, nullptr));
    if (! _hFileComboHost)
    {
        Debug::ErrorWithLastError(L"ViewerWeb: CreateWindowExW failed for DxUi file combo host.");
    }
    else if (! _fileComboHost.Attach(_hFileComboHost.get()))
    {
        Debug::Error(L"ViewerWeb: failed to attach DxUi host for file combo.");
        _hFileComboHost.reset();
    }
    else if (! SetPropW(_hFileComboHost.get(), kFileComboHostStateProp, reinterpret_cast<HANDLE>(this)) ||
             ! InstallWndProcHook(_hFileComboHost.get(), kFileComboHostOriginalWndProcProp, FileComboHostWndProc))
    {
        RemovePropW(_hFileComboHost.get(), kFileComboHostStateProp);
        Debug::ErrorWithLastError(L"ViewerWeb: failed to install WNDPROC hook for DxUi file combo host.");
        _fileComboHost.Detach();
        _hFileComboHost.reset();
    }
    else
    {
        auto combo        = std::make_unique<ComboBox>();
        _fileComboControl = combo.get();
        _fileComboControl->SetVariant(ComboBoxVariant::Window);
        _fileComboControl->SetOnSelectionChanged([this, hwnd](size_t selectedIndex)
        {
            if (_syncingFileCombo || selectedIndex >= _otherFiles.size())
            {
                return;
            }

            _otherIndex = selectedIndex;
            static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
            if (! _embeddedMode)
            {
                SetFocus(hwnd);
            }
        });
        _fileComboHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
        _fileComboHost.SetRoot(std::move(combo));
    }

    if (! _menuHandle)
    {
        _menuHandle.reset(GetMenu(hwnd));
    }
    if (_menuHandle)
    {
        _menuBarHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
        _menuBarHost.SetRefreshMenuStateCallback([this, hwnd] { UpdateMenuState(hwnd, false); });
        static_cast<void>(_menuBarHost.Attach(g_hInstance, hwnd, _menuHandle.get()));
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

    if (_tempExtractedPath.has_value() && ! _tempExtractedPath.value().empty())
    {
        std::error_code ec;
        static_cast<void>(std::filesystem::remove(_tempExtractedPath.value(), ec));
        _tempExtractedPath.reset();
    }

    _pendingPath.reset();
    _pendingWebContent.reset();
    _pendingDocumentUtf8.reset();
    _internalDocumentUrl.reset();

    NotifyViewerClosed();
}

void ViewerWeb::OnSize(UINT /*width*/, UINT /*height*/) noexcept
{
    if (_hWnd)
    {
        Layout(_hWnd.get());
    }
}

void ViewerWeb::OnCommand(HWND hwnd, UINT commandId, [[maybe_unused]] UINT code, HWND /*control*/) noexcept
{
    switch (commandId)
    {
        case IDM_VIEWERWEB_FILE_SAVE_AS: static_cast<void>(CommandSaveAs(hwnd)); break;
        case IDM_VIEWERWEB_FILE_REFRESH: static_cast<void>(OpenPath(hwnd, _currentPath, false)); break;
        case IDM_VIEWERWEB_FILE_EXIT: static_cast<void>(Close()); break;

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

void ViewerWeb::OnContextMenu(HWND hwnd, POINT screenPt) noexcept
{
    if (! _menuHandle)
    {
        _menuHandle.reset(Localization::LoadMenuResource(g_hInstance, IDR_VIEWERWEB_MENU));
    }

    HMENU menu = _menuHandle ? _menuHandle.get() : GetMenu(hwnd);
    if (! menu)
    {
        return;
    }

    UpdateMenuState(hwnd, false);
    static constexpr std::array<int, 5> kPreviewContextMenuExcludedCommandIds{{
        IDM_VIEWERWEB_FILE_EXIT,
        IDM_VIEWERWEB_OTHER_NEXT,
        IDM_VIEWERWEB_OTHER_PREVIOUS,
        IDM_VIEWERWEB_OTHER_FIRST,
        IDM_VIEWERWEB_OTHER_LAST,
    }};
    RedSalamander::DxUi::NativeMenuFlyoutOptions previewMenuOptions{};
    previewMenuOptions.includeAcceleratorText = false;
    previewMenuOptions.omitEmptySubmenus      = true;
    previewMenuOptions.trimSeparators         = true;
    previewMenuOptions.excludedCommandIds     = kPreviewContextMenuExcludedCommandIds;

    const auto result = RedSalamander::DxUi::ShowNativeHMenuContextMenu(
        hwnd, screenPt, menu, _hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false), previewMenuOptions);
    if (result.has_value() && result.value() > 0)
    {
        OnCommand(hwnd, static_cast<UINT>(result.value()), 0, nullptr);
    }
}

void ViewerWeb::OnKeyDown(HWND hwnd, UINT vk) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const bool ctrl          = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift         = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    const HKL keyboardLayout = GetKeyboardLayout(0);
    const UINT zoomInVk      = ZoomInVirtualKeyForLayout(keyboardLayout);
    const UINT zoomOutVk     = ZoomOutVirtualKeyForLayout(keyboardLayout);

    if (vk == VK_ESCAPE)
    {
        static_cast<void>(Close());
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

    if (ctrl && (vk == VK_ADD || vk == zoomInVk || vk == '='))
    {
        CommandZoomIn();
        return;
    }

    if (ctrl && (vk == VK_SUBTRACT || vk == zoomOutVk || vk == '-'))
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
        wil::unique_hbrush brush(CreateSolidBrush(accent));
        FillRect(hdc.get(), &line, brush.get());
    }

    if (! _statusMessage.empty())
    {
        const UINT dpi    = GetDpiForWindow(hwnd);
        const int padding = PxFromDip(8, dpi);
        RECT rc           = _headerRect;
        rc.left           = std::min(rc.right, rc.left + padding);
        rc.right          = std::max(rc.left, rc.right - padding);

        const COLORREF textColor = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
        if (! DrawStatusMessageWithDirectWrite(hdc.get(), hwnd, rc, _statusMessage, textColor))
        {
            static std::atomic<bool> loggedStatusRenderFailure{false};
            if (! loggedStatusRenderFailure.exchange(true, std::memory_order_relaxed))
            {
                Debug::Error(L"ViewerWeb: DirectWrite/D2D status rendering failed; status text is unavailable without a DX render path.");
            }
        }
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

    static_cast<void>(newDpi);
    ApplyTheme(hwnd);
    Layout(hwnd);
}

LRESULT ViewerWeb::OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    OnDestroy();
    static_cast<void>(DrainPostedPayloadsForWindow(hwnd));

    _menuBarHost.Detach();
    _menuHandle.reset();
    UnhookFileComboHostWindow(_hFileComboHost.get());
    _fileComboHost.Detach();
    _hFileComboHost.release();
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

    if (! _hFileComboHost)
    {
        return;
    }

    _fileComboControl            = nullptr;
    _fileComboHostPreExpandPopup = false;
    UnhookFileComboHostWindow(_hFileComboHost.get());
    _fileComboHost.Detach();
    _hFileComboHost.reset();

    if (_tempExtractedPath.has_value() && ! _tempExtractedPath.value().empty())
    {
        std::error_code ec;
        static_cast<void>(std::filesystem::remove(_tempExtractedPath.value(), ec));
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
    _metaData.version     = VERSINFO_PLUGIN_VERSION;

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = GetViewerWebStaticConfigurationSchema(_kind);
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
                        else if (s && (strcmp(s, "jsonl") == 0 || strcmp(s, "json-lines") == 0 || strcmp(s, "2") == 0))
                        {
                            jsonViewMode = JsonViewMode::JsonLines;
                        }
                        else if (s && (strcmp(s, "pretty") == 0 || strcmp(s, "0") == 0))
                        {
                            jsonViewMode = JsonViewMode::Pretty;
                        }
                    }
                    else if (yyjson_is_int(modeVal))
                    {
                        const int64_t value = yyjson_get_int(modeVal);
                        jsonViewMode        = value <= 0 ? JsonViewMode::Pretty : value == 1 ? JsonViewMode::Tree : JsonViewMode::JsonLines;
                    }
                    else if (yyjson_is_uint(modeVal))
                    {
                        const uint64_t value = yyjson_get_uint(modeVal);
                        jsonViewMode         = value == 0u ? JsonViewMode::Pretty : value == 1u ? JsonViewMode::Tree : JsonViewMode::JsonLines;
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
                _config.jsonViewMode == JsonViewMode::Tree        ? "tree"
                : _config.jsonViewMode == JsonViewMode::JsonLines ? "jsonl"
                                                                  : "pretty",
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

    const bool embeddedMode = IsEmbeddedOpen(*context);
    const HWND embeddedParent = embeddedMode ? context->ownerWindow : nullptr;
    if (embeddedMode && (embeddedParent == nullptr || IsWindow(embeddedParent) == FALSE))
    {
        Debug::Error(L"ViewerWeb: embedded Open requires a valid ownerWindow parent.");
        return E_INVALIDARG;
    }
    if (ShouldRecreateViewerWindow(embeddedMode, embeddedParent))
    {
        static_cast<void>(Close());
    }

    if (! _hWnd)
    {
        if (! RegisterWndClass(g_hInstance))
        {
            return E_FAIL;
        }

        RECT ownerRect{};
        int x = CW_USEDEFAULT;
        int y = CW_USEDEFAULT;
        int w = 1000;
        int h = 700;
        if (embeddedMode)
        {
            RECT client{};
            if (GetClientRect(embeddedParent, &client) == 0)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerWeb: GetClientRect failed for embedded preview parent.");
                return HRESULT_FROM_WIN32(lastError);
            }
            x = 0;
            y = 0;
            w = std::max(1L, client.right - client.left);
            h = std::max(1L, client.bottom - client.top);
        }
        else if (context->ownerWindow && GetWindowRect(context->ownerWindow, &ownerRect) != 0)
        {
            x = static_cast<int>(ownerRect.left);
            y = static_cast<int>(ownerRect.top);
            w = static_cast<int>(std::max<LONG>(1, ownerRect.right - ownerRect.left));
            h = static_cast<int>(std::max<LONG>(1, ownerRect.bottom - ownerRect.top));
        }

        wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(
            embeddedMode ? nullptr : Localization::LoadMenuResource(g_hInstance, IDR_VIEWERWEB_MENU));
        const DWORD style = embeddedMode ? (WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS) : (WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN);
        const std::wstring initialTitle = embeddedMode ? std::wstring{} : (_metaName.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERWEB_NAME) : _metaName);
        HWND window =
            CreateWindowExW(0, kClassName, initialTitle.c_str(), style, x, y, w, h, embeddedMode ? embeddedParent : nullptr, menu.get(), g_hInstance, this);
        if (! window)
        {
            const DWORD lastError = Debug::ErrorWithLastError(L"ViewerWeb: CreateWindowExW failed.");
            return HRESULT_FROM_WIN32(lastError);
        }

        if (! embeddedMode)
        {
            menu.release();
        }
        _hWnd.reset(window);
        _embeddedMode = embeddedMode;

        ApplyTheme(_hWnd.get());
        ApplyPendingViewerWebClassBackgroundBrush(_hWnd.get());

        AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
        ShowWindow(_hWnd.get(), embeddedMode ? SW_SHOWNA : SW_SHOWNORMAL);
        if (! embeddedMode)
        {
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }
    else
    {
        ApplyPendingViewerWebClassBackgroundBrush(_hWnd.get());
        if (_embeddedMode)
        {
            RECT client{};
            if (embeddedParent != nullptr && GetClientRect(embeddedParent, &client) != 0)
            {
                SetWindowPos(_hWnd.get(),
                             nullptr,
                             0,
                             0,
                             std::max(1L, client.right - client.left),
                             std::max(1L, client.bottom - client.top),
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            ShowWindow(_hWnd.get(), SW_SHOWNA);
        }
        else
        {
            ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
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
    AddRef();
    const auto releaseSelf = wil::scope_exit([&]() noexcept { Release(); });
    _hWnd.reset();
    _embeddedMode = false;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::SetTheme(const ViewerTheme* theme) noexcept
{
    if (! theme || theme->version < 2u || theme->version > 4u)
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
#ifdef ENABLE_TESTS
        case WndMsg::kViewerDebugGetNativeMenuModelSnapshot:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerNativeMenuModelDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot                    = {};
            snapshot->hasHiddenMenuModel = _menuHandle != nullptr;
            snapshot->ownerDrawItemCount = CountOwnerDrawMenuItems(_menuHandle.get());
            return TRUE;
        }
#endif
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_SIZE: OnSize(static_cast<UINT>(LOWORD(lp)), static_cast<UINT>(HIWORD(lp))); return 0;
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lp))
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hwnd, *info, 520, 320);
            }
            return 0;
        case WM_COMMAND: OnCommand(hwnd, static_cast<UINT>(LOWORD(wp)), static_cast<UINT>(HIWORD(wp)), reinterpret_cast<HWND>(lp)); return 0;
        case WM_CONTEXTMENU: OnContextMenu(hwnd, RedSalamander::DxUi::ResolveNativeContextMenuScreenPoint(hwnd, lp)); return 0;
        case WM_KEYDOWN: OnKeyDown(hwnd, static_cast<UINT>(wp)); return 0;
        case WM_SYSKEYDOWN:
            if ((wp == VK_F10 || wp == VK_MENU) && _menuBarHost.FocusFirstItem())
            {
                return 0;
            }
            if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
            {
                OnKeyDown(hwnd, static_cast<UINT>(wp));
                return 0;
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_SYSCHAR:
            if (wp >= 0x20u && _menuBarHost.ActivateMnemonic(static_cast<wchar_t>(wp)))
            {
                return 0;
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_PAINT: OnPaint(hwnd); return 0;
        case WM_ERASEBKGND: return OnEraseBkgnd(hwnd, reinterpret_cast<HDC>(wp));
        case WM_DPICHANGED: OnDpiChanged(hwnd, HIWORD(wp), reinterpret_cast<const RECT*>(lp)); return 0;
        case kAsyncLoadCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncLoadResult>(lp);
            OnAsyncLoadComplete(std::move(result));
            return 0;
        }
        case WM_CLOSE: static_cast<void>(Close()); return 0;
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
            static_cast<void>(std::filesystem::remove(result->extractedWin32Path.value(), ec));
        }
        return;
    }

    _statusMessage               = result->statusMessage;
    _jsonExpandCollapseAvailable = SUCCEEDED(result->hr) && result->jsonExpandCollapseAvailable;
    UpdateMenuState(_hWnd.get());

    if (FAILED(result->hr))
    {
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        }

        if (result->offerTextViewerFallback && OfferTextViewerFallbackPrompt())
        {
            return;
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
    _pendingDocumentUtf8.reset();
    _internalDocumentUrl.reset();

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
                    static_cast<void>(std::filesystem::remove(_tempExtractedPath.value(), ec));
                }
                _tempExtractedPath = navPath;
            }
            else if (_tempExtractedPath.has_value())
            {
                std::error_code ec;
                static_cast<void>(std::filesystem::remove(_tempExtractedPath.value(), ec));
                _tempExtractedPath.reset();
            }

            const std::wstring url = UrlFromFilePath(navPath.value().wstring());
            if (url.empty())
            {
                _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_BUILD_FILE_URL);
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
        if (result->utf8.empty())
        {
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_BUILD_HTML_DOCUMENT);
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
            }
            return;
        }

        _pendingDocumentUtf8 = std::move(result->utf8);
        _internalDocumentUrl = BuildInternalDocumentUrl(_kind, result->requestId);
        _pendingPath         = _internalDocumentUrl;
    }

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
    }

    if (SUCCEEDED(EnsureWebView2(_hWnd.get())) && _webView)
    {
        const HRESULT navHr = NavigatePendingContent(_hWnd.get());
        if (FAILED(navHr))
        {
            const std::wstring message = _statusMessage.empty()
                                             ? FormatStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_NAVIGATE_FAILED_FMT, static_cast<unsigned long>(navHr))
                                             : _statusMessage;
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, message);
        }
    }
}

void ViewerWeb::Layout(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    _menuBarHost.UpdateLayout();
    ComputeLayoutRects(hwnd);

    const UINT dpi       = GetDpiForWindow(hwnd);
    const int minPadding = MulDiv(3, static_cast<int>(dpi), 96);
    const int accentH    = std::max(1, MulDiv(1, static_cast<int>(dpi), 96));
    const int accentGap  = std::max(1, MulDiv(1, static_cast<int>(dpi), 96));
    const bool showCombo = _hFileComboHost && _otherFiles.size() > 1;

    RECT headerContentRect{};
    headerContentRect        = _headerRect;
    headerContentRect.top    = std::min(headerContentRect.bottom, headerContentRect.top + minPadding);
    headerContentRect.bottom = std::max(headerContentRect.top, headerContentRect.bottom - accentH - accentGap - minPadding);
    const int headerContentH = std::max(0L, headerContentRect.bottom - headerContentRect.top);

    if (_hFileComboHost)
    {
        ShowWindow(_hFileComboHost.get(), showCombo ? SW_SHOW : SW_HIDE);
        EnableWindow(_hFileComboHost.get(), showCombo ? TRUE : FALSE);
        if (_fileComboControl)
        {
            _fileComboControl->SetEnabled(showCombo);
            _fileComboControl->SetVisible(showCombo);
        }
        if (! showCombo)
        {
            _fileComboHostPreExpandPopup = false;
        }

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

            int comboH = std::max(1, MulDiv(32, static_cast<int>(dpi), 96));
            comboH     = std::clamp(comboH, 1, std::max(1, headerContentH));

            int comboY          = headerContentRect.top + std::max(0, (headerContentH - comboH) / 2);
            const int maxBottom = std::max(static_cast<int>(headerContentRect.top), static_cast<int>(headerContentRect.bottom));
            if (comboY + comboH > maxBottom)
            {
                comboY = std::max(static_cast<int>(headerContentRect.top), maxBottom - comboH);
            }

            const bool expandPopupHost = _fileComboHostPreExpandPopup || (_fileComboControl && _fileComboControl->DebugIsPopupOpen());
            const int hostHeight       = comboH + (expandPopupHost ? ComputeWindowComboPopupHeightPx(_otherFiles.size(), dpi) : 0);
            SetWindowPos(_hFileComboHost.get(), HWND_TOP, comboX, comboY, comboW, hostHeight, SWP_NOACTIVATE);
            if (_fileComboControl)
            {
                _fileComboControl->SetBounds(D2D1::RectF(
                    0.0f, 0.0f, static_cast<float>(comboW) * 96.0f / static_cast<float>(dpi), static_cast<float>(comboH) * 96.0f / static_cast<float>(dpi)));
                _fileComboHost.Invalidate();
            }
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

    client.top += _menuBarHost.GetHwnd() ? _menuBarHost.GetVisibleHeightPx() : 0;

    const UINT dpi               = GetDpiForWindow(hwnd);
    const int baseHeaderHeight   = MulDiv(kHeaderHeightDip, static_cast<int>(dpi), 96);
    const int accentH            = std::max(1, MulDiv(1, static_cast<int>(dpi), 96));
    const int accentGap          = std::max(1, MulDiv(1, static_cast<int>(dpi), 96));
    const int minPadding         = MulDiv(3, static_cast<int>(dpi), 96);
    const bool showCombo         = _hFileComboHost && _otherFiles.size() > 1;
    const int desiredComboHeight = std::max(1, MulDiv(32, static_cast<int>(dpi), 96));

    const int minChromeHeight = MulDiv(22, static_cast<int>(dpi), 96) + accentH + accentGap + 2 * minPadding;
    int headerH               = std::max(baseHeaderHeight, minChromeHeight);
    if (showCombo)
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

    _fileComboHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));

    ApplyMenuTheme(hwnd);
    UpdateMenuState(hwnd);
    ApplyWebViewThemeScript();
}

void ViewerWeb::UpdateMenuState(HWND hwnd, bool syncDxMenuBar) noexcept
{
    HMENU menu = _menuHandle ? _menuHandle.get() : (hwnd ? GetMenu(hwnd) : nullptr);
    if (! menu)
    {
        return;
    }

    const bool jsonInteractiveMode = _kind == ViewerWebKind::Json && _jsonExpandCollapseAvailable;

    EnableMenuItem(menu, IDM_VIEWERWEB_VIEW_DEVTOOLS, static_cast<UINT>(MF_BYCOMMAND | (_config.devToolsEnabled ? MF_ENABLED : MF_GRAYED)));

    EnableMenuItem(menu, IDM_VIEWERWEB_TOOLS_JSON_EXPAND_ALL, static_cast<UINT>(MF_BYCOMMAND | (jsonInteractiveMode ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, IDM_VIEWERWEB_TOOLS_JSON_COLLAPSE_ALL, static_cast<UINT>(MF_BYCOMMAND | (jsonInteractiveMode ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(
        menu, IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE, static_cast<UINT>(MF_BYCOMMAND | (_kind == ViewerWebKind::Markdown ? MF_ENABLED : MF_GRAYED)));

    CheckMenuItem(menu,
                  IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE,
                  static_cast<UINT>(MF_BYCOMMAND | ((_kind == ViewerWebKind::Markdown && _markdownShowSource) ? MF_CHECKED : MF_UNCHECKED)));

    if (syncDxMenuBar && _menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
    else if (hwnd)
    {
        DrawMenuBar(hwnd);
    }
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
    static_cast<void>(hwnd);
    _menuBarHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
    if (_menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
}

void ViewerWeb::ShowHostAlert(HWND targetWindow, HostAlertSeverity severity, const std::wstring& message) noexcept
{
    if (message.empty())
    {
        return;
    }

    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_TITLE);

    if (! _hostAlerts && _host)
    {
        wil::com_ptr<IHostAlerts> alerts;
        const HRESULT hr = _host->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
        if (SUCCEEDED(hr) && alerts)
        {
            _hostAlerts = std::move(alerts);
        }
    }

    if (! _hostAlerts)
    {
        const std::wstring_view normalizedTitle = title.empty() ? L"ViewerWeb" : std::wstring_view(title);
        switch (severity)
        {
            case HOST_ALERT_WARNING: Debug::Warning(L"ViewerWeb: Host alert dropped ({}): {}", normalizedTitle, message); break;
            case HOST_ALERT_INFO:
            case HOST_ALERT_BUSY: Debug::Info(L"ViewerWeb: Host alert dropped ({}): {}", normalizedTitle, message); break;
            case HOST_ALERT_ERROR:
            default: Debug::Error(L"ViewerWeb: Host alert dropped ({}): {}", normalizedTitle, message); break;
        }
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

bool ViewerWeb::OfferTextViewerFallbackPrompt() noexcept
{
    if (! _host || _statusMessage.empty())
    {
        return false;
    }

    wil::com_ptr<IHostPrompts> prompts;
    const HRESULT promptsHr = _host->QueryInterface(__uuidof(IHostPrompts), prompts.put_void());
    if (FAILED(promptsHr) || ! prompts)
    {
        return false;
    }

    const std::wstring title   = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_TITLE);
    std::wstring promptMessage = FormatStringResource(g_hInstance, IDS_VIEWERWEB_PROMPT_OPEN_TEXT_VIEWER_FMT, _statusMessage);
    if (promptMessage.empty())
    {
        promptMessage = _statusMessage;
    }

    HostPromptRequest request{};
    request.version       = 1;
    request.sizeBytes     = sizeof(request);
    request.scope         = (_hWnd && IsWindow(_hWnd.get())) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    request.severity      = HOST_ALERT_WARNING;
    request.buttons       = HOST_PROMPT_BUTTONS_YES_NO;
    request.targetWindow  = request.scope == HOST_ALERT_SCOPE_WINDOW ? _hWnd.get() : nullptr;
    request.title         = title.empty() ? nullptr : title.c_str();
    request.message       = promptMessage.c_str();
    request.defaultResult = HOST_PROMPT_RESULT_YES;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT promptHr        = prompts->ShowPrompt(&request, _hWnd.get(), &promptResult);
    if (FAILED(promptHr))
    {
        return false;
    }

    if (promptResult != HOST_PROMPT_RESULT_YES)
    {
        return true;
    }

    const HRESULT openHr = OpenCurrentDocumentInTextViewer();
    if (FAILED(openHr))
    {
        ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_OPEN_TEXT_VIEWER_FAILED));
        return true;
    }

    if (_hWnd)
    {
        PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0);
    }

    return true;
}

HRESULT ViewerWeb::OpenCurrentDocumentInTextViewer() noexcept
{
    if (! _host || ! _fileSystem || _currentPath.empty())
    {
        return E_UNEXPECTED;
    }

    wil::com_ptr<IHostViewers> hostViewers;
    const HRESULT viewersHr = _host->QueryInterface(__uuidof(IHostViewers), hostViewers.put_void());
    if (FAILED(viewersHr) || ! hostViewers)
    {
        return FAILED(viewersHr) ? viewersHr : E_NOINTERFACE;
    }

    std::vector<const wchar_t*> otherFilePointers;
    otherFilePointers.reserve(std::max<size_t>(_otherFiles.size(), 1u));
    for (const auto& otherPath : _otherFiles)
    {
        if (! otherPath.empty())
        {
            otherFilePointers.push_back(otherPath.c_str());
        }
    }
    if (otherFilePointers.empty())
    {
        otherFilePointers.push_back(_currentPath.c_str());
    }

    HostViewerOpenRequest request{};
    request.version               = 1;
    request.sizeBytes             = sizeof(request);
    request.pluginId              = L"builtin/viewer-text";
    request.ownerWindow           = _hWnd.get();
    request.fileSystem            = _fileSystem.get();
    request.fileSystemName        = _fileSystemName.empty() ? nullptr : _fileSystemName.c_str();
    request.focusedPath           = _currentPath.c_str();
    request.selectionPaths        = nullptr;
    request.selectionCount        = 0;
    request.otherFiles            = otherFilePointers.data();
    request.otherFileCount        = static_cast<unsigned long>(otherFilePointers.size());
    request.focusedOtherFileIndex = _otherIndex < otherFilePointers.size() ? static_cast<unsigned long>(_otherIndex) : 0;
    request.viewerFlags           = VIEWER_OPEN_FLAG_NONE;

    return hostViewers->OpenViewer(&request);
}

HRESULT ViewerWeb::CreateControllerFromEnvironment(HWND hwnd, ICoreWebView2Environment* environment) noexcept
{
    const HRESULT hr =
        environment->CreateCoreWebView2Controller(hwnd,
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

        // Configure settings once — these don't change between navigations.
        ConfigureWebViewSettings();

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
            const bool isHttp             = OrdinalString::StartsWithNoCase(url, L"http://") || OrdinalString::StartsWithNoCase(url, L"https://");
            const bool isAbout            = OrdinalString::StartsWithNoCase(url, L"about:");
            const bool isData             = OrdinalString::StartsWithNoCase(url, L"data:");
            const bool isInternalDocument = _internalDocumentUrl.has_value() && OrdinalString::EqualsNoCase(url, _internalDocumentUrl.value());

            if (_kind == ViewerWebKind::Web)
            {
                if (isHttp && ! _config.allowExternalNavigation)
                {
                    static_cast<void>(args->put_Cancel(TRUE));
                }
                return S_OK;
            }

            // JSON/Markdown: keep viewer content stable and open external links in the system browser.
            if (isInternalDocument)
            {
                return S_OK;
            }

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
        }).get(),
            &_navStartingToken));

        static_cast<void>(_webView->AddWebResourceRequestedFilter(kInternalDocumentFilter.data(), COREWEBVIEW2_WEB_RESOURCE_CONTEXT_DOCUMENT));
        static_cast<void>(_webView->add_WebResourceRequested(
            MakeComCallback<ICoreWebView2WebResourceRequestedEventHandler, ICoreWebView2*, ICoreWebView2WebResourceRequestedEventArgs*>(
                [this](ICoreWebView2* /*sender*/, ICoreWebView2WebResourceRequestedEventArgs* args) -> HRESULT
        {
            if (! args || ! g_sharedEnvironment.environment)
            {
                return S_OK;
            }

            wil::com_ptr<ICoreWebView2WebResourceRequest> request;
            if (FAILED(args->get_Request(request.put())) || ! request)
            {
                return S_OK;
            }

            wil::unique_cotaskmem_string uri;
            static_cast<void>(request->get_Uri(uri.put()));
            if (! uri)
            {
                return S_OK;
            }

            const std::wstring_view requestUrl(uri.get());
            if (! IsInternalDocumentUrl(requestUrl))
            {
                return S_OK;
            }

            wil::com_ptr<ICoreWebView2WebResourceResponse> response;
            if (_internalDocumentUrl.has_value() && OrdinalString::EqualsNoCase(requestUrl, _internalDocumentUrl.value()) && _pendingDocumentUtf8.has_value())
            {
                const auto* data = reinterpret_cast<const BYTE*>(_pendingDocumentUtf8->data());
                const UINT size  = static_cast<UINT>(std::min<size_t>(_pendingDocumentUtf8->size(), static_cast<size_t>(std::numeric_limits<UINT>::max())));
                wil::com_ptr<IStream> stream;
                stream.attach(SHCreateMemStream(data, size));
                if (! stream)
                {
                    return E_OUTOFMEMORY;
                }

                const HRESULT responseHr = g_sharedEnvironment.environment->CreateWebResourceResponse(
                    stream.get(), 200, L"OK", L"Content-Type: text/html; charset=utf-8\r\nCache-Control: no-store\r\n", response.put());
                if (FAILED(responseHr) || ! response)
                {
                    return responseHr;
                }
            }
            else
            {
                const HRESULT responseHr = g_sharedEnvironment.environment->CreateWebResourceResponse(
                    nullptr, 404, L"Not Found", L"Content-Type: text/plain; charset=utf-8\r\nCache-Control: no-store\r\n", response.put());
                if (FAILED(responseHr) || ! response)
                {
                    return responseHr;
                }
            }

            return args->put_Response(response.get());
        }).get(),
            &_webResourceRequestedToken));

        static_cast<void>(_webView->add_NavigationCompleted(
            MakeComCallback<ICoreWebView2NavigationCompletedEventHandler, ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs*>(
                [this](ICoreWebView2* /*sender*/, ICoreWebView2NavigationCompletedEventArgs* /*args*/) -> HRESULT
        {
            ApplyWebViewThemeScript();
            return S_OK;
        }).get(),
            &_navCompletedToken));

        static_cast<void>(_webViewController->add_AcceleratorKeyPressed(
            MakeComCallback<ICoreWebView2AcceleratorKeyPressedEventHandler, ICoreWebView2Controller*, ICoreWebView2AcceleratorKeyPressedEventArgs*>(
                [this](ICoreWebView2Controller* /*sender*/, ICoreWebView2AcceleratorKeyPressedEventArgs* args) -> HRESULT
        {
            if (! args || ! _hWnd)
            {
                return S_OK;
            }

            AddRef();
            auto releaseSelf = wil::scope_exit([&] { Release(); });

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

            const bool ctrl          = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
            const bool shift         = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
            const HKL keyboardLayout = GetKeyboardLayout(0);
            const UINT zoomInVk      = ZoomInVirtualKeyForLayout(keyboardLayout);
            const UINT zoomOutVk     = ZoomOutVirtualKeyForLayout(keyboardLayout);

            const auto handle = [&](bool handled) noexcept
            {
                if (handled)
                {
                    static_cast<void>(args->put_Handled(TRUE));
                }
            };

            if (vk == VK_ESCAPE)
            {
                const HWND hwndToClose = _hWnd.get();
                if (hwndToClose)
                {
                    PostMessageW(hwndToClose, WM_CLOSE, 0, 0);
                }
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

            if (ctrl && (vk == VK_ADD || vk == zoomInVk || vk == '='))
            {
                CommandZoomIn();
                handle(true);
                return S_OK;
            }

            if (ctrl && (vk == VK_SUBTRACT || vk == zoomOutVk || vk == '-'))
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
        }).get(),
            &_accelToken));

        Layout(hwnd);
        ApplyWebViewThemeScript();

        const HRESULT navHr = NavigatePendingContent(hwnd);
        if (FAILED(navHr))
        {
            const std::wstring message = _statusMessage.empty()
                                             ? FormatStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_NAVIGATE_FAILED_FMT, static_cast<unsigned long>(navHr))
                                             : _statusMessage;
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, message);
        }

        return S_OK;
    }).get());

    if (FAILED(hr))
    {
        _webViewInitInProgress = false;
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED));
        Release();
    }

    return hr;
}

HRESULT ViewerWeb::NavigatePendingContent(HWND /*hwnd*/) noexcept
{
    if (! _webView)
    {
        return E_UNEXPECTED;
    }

    if (_pendingPath.has_value())
    {
        const std::wstring url = std::move(_pendingPath.value());
        _pendingPath.reset();
        return _webView->Navigate(url.c_str());
    }

    if (_pendingWebContent.has_value())
    {
        const std::wstring html = std::move(_pendingWebContent.value());
        _pendingWebContent.reset();
        if (html.empty())
        {
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_BUILD_HTML_DOCUMENT);
            return E_INVALIDARG;
        }

        const HRESULT navHr = _webView->NavigateToString(html.c_str());
        if (FAILED(navHr))
        {
            _statusMessage = FormatStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_NAVIGATE_TO_STRING_FAILED_FMT, static_cast<unsigned long>(navHr));
        }
        return navHr;
    }

    if (_pendingDocumentUtf8.has_value() && _internalDocumentUrl.has_value())
    {
        return _webView->Navigate(_internalDocumentUrl->c_str());
    }

    return S_FALSE;
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

    // Fast path: shared environment already exists — create controller immediately.
    if (g_sharedEnvironment.environment)
    {
        return CreateControllerFromEnvironment(hwnd, g_sharedEnvironment.environment.get());
    }

    // Another instance is already creating the environment — queue ourselves.
    if (g_sharedEnvironment.createInProgress)
    {
        g_sharedEnvironment.pendingConsumers.push_back({this, hwnd});
        return S_FALSE;
    }

    // First caller — create the shared environment.
    g_sharedEnvironment.createInProgress = true;

    const std::wstring userDataFolder = GetWebView2UserDataFolder();

    const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
        nullptr,
        userDataFolder.empty() ? nullptr : userDataFolder.c_str(),
        nullptr,
        MakeComCallback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler, HRESULT, ICoreWebView2Environment*>(
            [this, hwnd](HRESULT result, ICoreWebView2Environment* environment) -> HRESULT
    {
        g_sharedEnvironment.createInProgress = false;

        if (FAILED(result) || ! environment)
        {
            _webViewInitInProgress = false;

            const UINT msgId = (result == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || result == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
                                   ? IDS_VIEWERWEB_ERROR_WEBVIEW2_RUNTIME_MISSING
                                   : IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED;
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, msgId));

            // Notify all pending consumers about the failure.
            for (auto& pending : g_sharedEnvironment.pendingConsumers)
            {
                if (pending.viewer)
                {
                    pending.viewer->_webViewInitInProgress = false;
                    pending.viewer->ShowHostAlert(pending.hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, msgId));
                    pending.viewer->Release();
                }
            }
            g_sharedEnvironment.pendingConsumers.clear();

            Release();
            return S_OK;
        }

        g_sharedEnvironment.environment = environment;

        // Create our own controller.
        static_cast<void>(CreateControllerFromEnvironment(hwnd, environment));

        // Drain pending consumers — each creates its own controller from the shared environment.
        auto pendingConsumers = std::move(g_sharedEnvironment.pendingConsumers);
        for (auto& pending : pendingConsumers)
        {
            if (pending.viewer && pending.hwnd)
            {
                static_cast<void>(pending.viewer->CreateControllerFromEnvironment(pending.hwnd, environment));
            }
        }

        return S_OK;
    }).get());

    if (FAILED(hr))
    {
        g_sharedEnvironment.createInProgress = false;
        _webViewInitInProgress               = false;
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
        static_cast<void>(_webView->remove_WebResourceRequested(_webResourceRequestedToken));
    }

    _navStartingToken          = {};
    _navCompletedToken         = {};
    _accelToken                = {};
    _webResourceRequestedToken = {};

    // Close the WebView2 controller. Note: Close() is asynchronous and may have
    // pending I/O operations that complete on thread pool threads. This is why
    // we must call DiscardWebView2() early in OnDestroy() rather than in the
    // destructor - to give WebView2 time to complete its shutdown before the
    // plugin DLL is unloaded.
    if (_webViewController)
    {
        _webViewController->Close();
    }

    // Release per-instance COM pointers after initiating close.
    // The shared environment is NOT released here — it lives for the DLL lifetime.
    _webView.reset();
    _webViewController.reset();
}

void ViewerWeb::ConfigureWebViewSettings() noexcept
{
    if (! _webView)
    {
        return;
    }

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
}

void ViewerWeb::ApplyWebViewThemeScript() noexcept
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
            if (OrdinalString::EqualsNoCase(a, b))
            {
                _otherIndex = i;
                break;
            }
        }

        if (_fileComboControl)
        {
            _syncingFileCombo = true;
            auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });
            _fileComboControl->SetSelectedIndex(_otherIndex);
            _fileComboHost.Invalidate();
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

    _statusMessage               = LoadStringResource(g_hInstance, IDS_VIEWERWEB_STATUS_LOADING);
    _jsonExpandCollapseAvailable = false;
    _pendingPath.reset();
    _pendingWebContent.reset();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        UpdateMenuState(_hWnd.get());
    }

    return StartAsyncLoad(hwnd, path);
}

void ViewerWeb::RefreshFileCombo(HWND hwnd) noexcept
{
    if (! _fileComboControl)
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    std::vector<ComboBox::Item> items;
    items.reserve(_otherFiles.size());
    for (const auto& path : _otherFiles)
    {
        std::wstring leaf = LeafNameFromPath(path);
        if (leaf.empty())
        {
            leaf = path;
        }

        items.push_back(ComboBox::Item{path, std::move(leaf)});
    }
    _fileComboControl->SetItems(std::move(items));

    if (_otherIndex < _otherFiles.size())
    {
        _fileComboControl->SetSelectedIndex(_otherIndex);
    }
    else
    {
        _fileComboControl->SetSelectedIndex(std::nullopt);
    }

    _fileComboHost.Invalidate();
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
                static_cast<void>(std::filesystem::remove(extractedPath.value(), ec));
            }
            return;
        }
    };

    if (! fileSystem)
    {
        result->hr            = E_FAIL;
        result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_FILE_SYSTEM_UNAVAILABLE);
        postBack(false);
        return;
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    const HRESULT fileIoHr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (FAILED(fileIoHr) || ! fileIo)
    {
        result->hr            = FAILED(fileIoHr) ? fileIoHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_NO_FILE_IO);
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
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_OPEN_FOR_VIEWING_FAILED);
            postBack(false);
            return;
        }

        wchar_t tempDir[MAX_PATH]{};
        const DWORD tempDirLen = GetTempPathW(static_cast<DWORD>(std::size(tempDir)), tempDir);
        if (tempDirLen == 0 || tempDirLen >= std::size(tempDir))
        {
            result->hr            = HRESULT_FROM_WIN32(GetLastError());
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_GET_TEMP_FOLDER_FAILED);
            postBack(false);
            return;
        }

        wchar_t tempName[MAX_PATH]{};
        if (GetTempFileNameW(tempDir, L"rsw", 0, tempName) == 0)
        {
            result->hr            = HRESULT_FROM_WIN32(GetLastError());
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_CREATE_TEMP_FILE_FAILED);
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
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WRITE_TEMP_FILE_FAILED);
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
                result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_READ_FILE_FAILED);
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
                result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WRITE_TEMP_FILE_FAILED);
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
        result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_OPEN_FOR_VIEWING_FAILED);
        postBack(false);
        return;
    }

    uint64_t sizeBytes   = 0;
    const HRESULT sizeHr = reader->GetSize(&sizeBytes);
    if (FAILED(sizeHr))
    {
        result->hr            = sizeHr;
        result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_READ_FILE_SIZE_FAILED);
        postBack(false);
        return;
    }

    const uint64_t maxBytes = static_cast<uint64_t>(config.maxDocumentMiB) * 1024ull * 1024ull;
    if (sizeBytes > maxBytes)
    {
        result->hr                      = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        result->offerTextViewerFallback = true;
        result->statusMessage =
            FormatStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_FILE_TOO_LARGE_FMT, FormatBytesCompact(sizeBytes), FormatBytesCompact(maxBytes));
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
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_READ_FILE_FAILED);
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
            result->hr                      = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
            result->offerTextViewerFallback = true;
            result->statusMessage           = FormatStringResource(
                g_hInstance, IDS_VIEWERWEB_ERROR_FILE_TOO_LARGE_FMT, FormatBytesCompact(static_cast<uint64_t>(bytes.size())), FormatBytesCompact(maxBytes));
            postBack(false);
            return;
        }
    }

    const std::string textUtf8 = normalizeTextUtf8(bytes);

    if (kind == ViewerWebKind::Json)
    {
        std::vector<JsonLinesEntry> jsonLinesEntries;
        const bool forceJsonLines     = config.jsonViewMode == JsonViewMode::JsonLines;
        const bool pathLooksJsonLines = IsJsonLinesPath(result->path);

        if (forceJsonLines)
        {
            if (! TryParseJsonLinesEntries(textUtf8, true, jsonLinesEntries))
            {
                result->hr            = E_FAIL;
                result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_PARSE_JSON_FAILED);
                postBack(false);
                return;
            }

            result->utf8 = BuildJsonLinesHtml(jsonLinesEntries, GetHighlightJs(), themeObj, bg, fg, selBg, selFg, accent, theme.darkMode != FALSE);
            result->jsonExpandCollapseAvailable = true;
            result->hr                          = S_OK;
            postBack(false);
            return;
        }

        if (pathLooksJsonLines && TryParseJsonLinesEntries(textUtf8, true, jsonLinesEntries))
        {
            result->utf8 = BuildJsonLinesHtml(jsonLinesEntries, GetHighlightJs(), themeObj, bg, fg, selBg, selFg, accent, theme.darkMode != FALSE);
            result->jsonExpandCollapseAvailable = true;
            result->hr                          = S_OK;
            postBack(false);
            return;
        }

        std::string jsonMutable(textUtf8);
        yyjson_read_err err{};
        unique_yyjson_doc doc(yyjson_read_opts(jsonMutable.data(), jsonMutable.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err));
        if (! doc)
        {
            if (TryParseJsonLinesEntries(textUtf8, false, jsonLinesEntries))
            {
                result->utf8 = BuildJsonLinesHtml(jsonLinesEntries, GetHighlightJs(), themeObj, bg, fg, selBg, selFg, accent, theme.darkMode != FALSE);
                result->jsonExpandCollapseAvailable = true;
                result->hr                          = S_OK;
                postBack(false);
                return;
            }

            result->hr            = E_FAIL;
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_PARSE_JSON_FAILED);
            postBack(false);
            return;
        }

        size_t prettyLen = 0;
        unique_malloc_string pretty(yyjson_write_opts(doc.get(), YYJSON_WRITE_PRETTY | YYJSON_WRITE_ESCAPE_UNICODE, nullptr, &prettyLen, nullptr));
        if (! pretty)
        {
            result->hr            = E_OUTOFMEMORY;
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_FORMAT_JSON_FAILED);
            postBack(false);
            return;
        }

        const std::string_view prettyJson(pretty.get(), prettyLen);
        const std::string escapedJson = EscapeJavaScriptStringUtf8(prettyJson);

        if (config.jsonViewMode == JsonViewMode::Pretty)
        {
            const std::string_view highlightJs = GetHighlightJs();

            const COLORREF codeBg             = BlendColor(bg, fg, theme.darkMode ? 20u : 10u);
            const COLORREF border             = BlendColor(bg, fg, theme.darkMode ? 35u : 45u);
            const COLORREF mutedFg            = BlendColor(bg, fg, 140u);
            const JsonTokenColors tokenColors = BuildJsonTokenColors(accent, fg, theme.darkMode != FALSE);
            const ScrollbarColors scrollbar   = BuildScrollbarColors(bg, fg, accent, theme.darkMode != FALSE);

            std::string html;
            html.reserve(highlightJs.size() + escapedJson.size() + 8192);
            html += "<!doctype html><html><head><meta charset=\"utf-8\">";
            html += "<style>";
            html += ":root{--rs-bg:" + cssRgb(bg) + ";--rs-fg:" + cssRgb(fg) + ";--rs-sel-bg:" + cssRgb(selBg) + ";--rs-sel-fg:" + cssRgb(selFg) +
                    ";--rs-accent:" + cssRgb(accent) + ";--rs-code-bg:" + cssRgb(codeBg) + ";--rs-border:" + cssRgb(border) +
                    ";--rs-muted-fg:" + cssRgb(mutedFg) + ";--rs-key:" + cssRgb(tokenColors.key) + ";--rs-string:" + cssRgb(tokenColors.stringValue) +
                    ";--rs-number:" + cssRgb(tokenColors.numberValue) + ";--rs-literal:" + cssRgb(tokenColors.literalValue) +
                    ";--rs-scroll-track:" + cssRgb(scrollbar.track) + ";--rs-scroll-thumb:" + cssRgb(scrollbar.thumb) +
                    ";--rs-scroll-thumb-hover:" + cssRgb(scrollbar.thumbHover) + ";--rs-scroll-corner:" + cssRgb(scrollbar.corner) + ";}";
            html += "html,body{height:100%;margin:0;}body{background:var(--rs-bg);color:var(--rs-fg);font-family:\"Segoe UI Variable Text\",\"Segoe UI "
                    "Variable Small\",\"Segoe UI\",sans-serif;}";
            html += kCommonScrollbarCss;
            html += "::selection{background:var(--rs-sel-bg);color:var(--rs-sel-fg);}#app{height:100%;box-sizing:border-box;padding:12px;display:flex;}";
            html += "pre{flex:1;margin:0;background:var(--rs-code-bg);border:1px solid var(--rs-border);padding:12px;overflow:auto;border-radius:6px;}";
            html += "code{font-family:Consolas,ui-monospace,monospace;font-size:13px;line-height:1.45;}";
            html += ".hljs{background:transparent;}";
            html += ".hljs-attr{color:var(--rs-key);} .hljs-string{color:var(--rs-string);} .hljs-number{color:var(--rs-number);} "
                    ".hljs-literal{color:var(--rs-literal);}";
            html += ".hljs-punctuation,.hljs-brace{color:var(--rs-muted-fg);} .hljs-comment{opacity:0.8;}";
            html += "</style></head><body><div id=\"app\"><pre><code id=\"code\" class=\"language-json\"></code></pre></div>";
            html += "<script>";
            html.append(highlightJs);
            html += "</script><script>";
            html += "(() => {";
            html += "const initialTheme=" + themeObj + ";";
            html += kCommonThemeJs;
            html += "function applyTheme(t){const "
                    "r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);r.setProperty('--rs-sel-bg',t.selBg);r."
                    "setProperty('--rs-sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);const "
                    "bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const "
                    "dark=luma(bg)<128;r.setProperty('--rs-code-bg',rgb(blend(bg,fg,dark?20:10)));r.setProperty('--rs-border',rgb(blend(bg,fg,dark?35:45)));r."
                    "setProperty('--rs-muted-fg',rgb(blend(bg,fg,140)));setJsonTokenVars(r,bg,fg,acc);setScrollbarVars(r,bg,fg,acc);}";
            html += "const code=document.getElementById('code');";
            html += "code.textContent='" + escapedJson + "';";
            html += "window.RS={applyTheme:applyTheme};";
            html += "applyTheme(initialTheme);";
            html += "try{hljs.highlightElement(code);}catch(e){}";
            html += "})();";
            html += "</script></body></html>";

            result->utf8                        = std::move(html);
            result->jsonExpandCollapseAvailable = false;
            result->hr                          = S_OK;
            postBack(false);
            return;
        }

        const std::string_view jsonEditorJs = GetJsonEditorJs();
        const std::string& jsonEditorCss    = GetJsonEditorCssWithIcons();

        const COLORREF border             = BlendColor(bg, fg, theme.darkMode ? 45u : 80u);
        const COLORREF mutedFg            = BlendColor(bg, fg, 140u);
        const JsonTokenColors tokenColors = BuildJsonTokenColors(accent, fg, theme.darkMode != FALSE);
        const ScrollbarColors scrollbar   = BuildScrollbarColors(bg, fg, accent, theme.darkMode != FALSE);

        std::string html;
        html.reserve(jsonEditorJs.size() + jsonEditorCss.size() + escapedJson.size() + 8192);
        html += "<!doctype html><html><head><meta charset=\"utf-8\">";
        html += "<style>";
        html += ":root{--rs-bg:" + cssRgb(bg) + ";--rs-fg:" + cssRgb(fg) + ";--rs-sel-bg:" + cssRgb(selBg) + ";--rs-sel-fg:" + cssRgb(selFg) +
                ";--rs-accent:" + cssRgb(accent) + ";--rs-border:" + cssRgb(border) + ";--rs-muted-fg:" + cssRgb(mutedFg) +
                ";--rs-scroll-track:" + cssRgb(scrollbar.track) + ";--rs-scroll-thumb:" + cssRgb(scrollbar.thumb) +
                ";--rs-scroll-thumb-hover:" + cssRgb(scrollbar.thumbHover) + ";--rs-scroll-corner:" + cssRgb(scrollbar.corner) +
                ";--rs-key:" + cssRgb(tokenColors.key) + ";--rs-string:" + cssRgb(tokenColors.stringValue) + ";--rs-number:" + cssRgb(tokenColors.numberValue) +
                ";--rs-literal:" + cssRgb(tokenColors.literalValue) + ";}";
        html += "html,body{height:100%;margin:0;}body{background:var(--rs-bg);color:var(--rs-fg);font-family:\"Segoe UI Variable Text\",\"Segoe UI Variable "
                "Small\",\"Segoe UI\",sans-serif;}#app{height:100%;}";
        html += kCommonScrollbarCss;
        html += jsonEditorCss;
        html += "html,body{background:var(--rs-bg)!important;color:var(--rs-fg)!important;}#app{height:100%!important;}";
        html += ".jsoneditor{border:none!important;height:100%!important;background:var(--rs-bg)!important;color:var(--rs-fg)!important;}";
        html += ".jsoneditor-frame{background:var(--rs-bg)!important;border:1px solid var(--rs-border)!important;}";
        html += ".jsoneditor-outer,.jsoneditor-inner,.jsoneditor-tree,.jsoneditor-tree-inner,.jsoneditor-text,.jsoneditor-text "
                "textarea{background:var(--rs-bg)!important;color:var(--rs-fg)!important;}";
        html += ".jsoneditor-field{color:var(--rs-key)!important;}";
        html += ".jsoneditor-value.jsoneditor-string{color:var(--rs-string)!important;}";
        html += ".jsoneditor-value.jsoneditor-number{color:var(--rs-number)!important;}";
        html += ".jsoneditor-value.jsoneditor-boolean,.jsoneditor-value.jsoneditor-null{color:var(--rs-literal)!important;}";
        html += ".jsoneditor-value.jsoneditor-object,.jsoneditor-value.jsoneditor-array{color:var(--rs-muted-fg)!important;}";
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
        html.append(jsonEditorJs);
        html += "</script><script>";
        html += "(() => {";
        html += "const initialTheme=" + themeObj + ";";
        html += kCommonThemeJs;
        html +=
            "function applyTheme(t){const "
            "r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);r.setProperty('--rs-sel-bg',t.selBg);r.setProperty('-"
            "-rs-sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);const bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const "
            "dark=luma(bg)<128;r.setProperty('--rs-border',rgb(blend(bg,fg,dark?45:80)));r.setProperty('--rs-muted-fg',rgb(blend(bg,fg,140)));"
            "setJsonTokenVars(r,bg,fg,acc);setScrollbarVars(r,bg,fg,acc);}";
        html += "const jsonText='" + escapedJson + "';";
        html += "const container=document.getElementById('app');";
        html += "const options={mode:'tree',modes:['tree','view'],onEditable:()=>false,mainMenuBar:false,navigationBar:false,statusBar:false};";
        html += "const editor=new JSONEditor(container,options);";
        html += "window.RS={applyTheme:applyTheme,expandAll:()=>editor.expandAll(),collapseAll:()=>editor.collapseAll()};";
        html += "applyTheme(initialTheme);";
        html += "try{editor.set(JSON.parse(jsonText));}catch(e){editor.set({error:String(e)});}";
        html += "})();";
        html += "</script></body></html>";

        result->utf8                        = std::move(html);
        result->jsonExpandCollapseAvailable = true;
        result->hr                          = S_OK;
        postBack(false);
        return;
    }

    // Markdown
    const std::string escapedMarkdown   = EscapeJavaScriptStringUtf8(textUtf8);
    const std::string_view markdownItJs = GetMarkdownItJs();
    const std::string_view highlightJs  = GetHighlightJs();

    const COLORREF codeBg           = BlendColor(bg, fg, theme.darkMode ? 20u : 10u);
    const COLORREF border           = BlendColor(bg, fg, theme.darkMode ? 35u : 45u);
    const COLORREF mutedFg          = BlendColor(bg, fg, 140u);
    const COLORREF stringColor      = BlendColor(accent, fg, 60u);
    const COLORREF numberColor      = BlendColor(accent, fg, 90u);
    const ScrollbarColors scrollbar = BuildScrollbarColors(bg, fg, accent, theme.darkMode != FALSE);

    std::string html;
    html.reserve(markdownItJs.size() + highlightJs.size() + escapedMarkdown.size() + 8192);
    html += "<!doctype html><html><head><meta charset=\"utf-8\">";
    html += "<style>";
    html += ":root{--rs-bg:" + cssRgb(bg) + ";--rs-fg:" + cssRgb(fg) + ";--rs-sel-bg:" + cssRgb(selBg) + ";--rs-sel-fg:" + cssRgb(selFg) +
            ";--rs-accent:" + cssRgb(accent) + ";--rs-code-bg:" + cssRgb(codeBg) + ";--rs-border:" + cssRgb(border) + ";--rs-muted-fg:" + cssRgb(mutedFg) +
            ";--rs-string:" + cssRgb(stringColor) + ";--rs-number:" + cssRgb(numberColor) + ";--rs-scroll-track:" + cssRgb(scrollbar.track) +
            ";--rs-scroll-thumb:" + cssRgb(scrollbar.thumb) + ";--rs-scroll-thumb-hover:" + cssRgb(scrollbar.thumbHover) +
            ";--rs-scroll-corner:" + cssRgb(scrollbar.corner) + ";}";
    html += "html,body{height:100%;margin:0;}body{background:var(--rs-bg);color:var(--rs-fg);font-family:\"Segoe UI Variable Text\",\"Segoe UI Variable "
            "Small\",\"Segoe UI\",sans-serif;}";
    html += kCommonScrollbarCss;
    html += "#app{max-width:100%;padding:16px;box-sizing:border-box;}a{color:var(--rs-accent);}";
    html += "pre{background:var(--rs-code-bg);border:1px solid var(--rs-border);padding:12px;overflow:auto;border-radius:6px;}";
    html += "code{font-family:Consolas,ui-monospace,monospace;}";
    html += "table{border-collapse:collapse;}th,td{border:1px solid var(--rs-border);padding:6px 10px;}";
    html += ".rs-source{white-space:pre;overflow:auto;font-family:Consolas,ui-monospace,monospace;}";
    html += ".hljs-comment{opacity:0.8;}.hljs-keyword,.hljs-selector-tag{color:var(--rs-accent);}";
    html += ".hljs-string{color:var(--rs-string);}.hljs-number{color:var(--rs-number);}.hljs-punctuation,.hljs-brace{color:var(--rs-muted-fg);}";
    html += "</style></head><body><div id=\"app\"></div>";
    html += "<script>";
    html.append(markdownItJs);
    html += "</script><script>";
    html.append(highlightJs);
    html += "</script><script>";
    html += "(() => {";
    html += "const initialTheme=" + themeObj + ";";
    html += kCommonThemeJs;
    html +=
        "function applyTheme(t){const "
        "r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);r.setProperty('--rs-sel-bg',t.selBg);r.setProperty('--rs-"
        "sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);const bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const "
        "dark=luma(bg)<128;r.setProperty('--rs-code-bg',rgb(blend(bg,fg,dark?20:10)));r.setProperty('--rs-border',rgb(blend(bg,fg,dark?35:45)));r.setProperty('"
        "--rs-muted-fg',rgb(blend(bg,fg,140)));r.setProperty('--rs-string',rgb(blend(acc,fg,60)));r.setProperty('--rs-number',rgb(blend(acc,fg,90)));"
        "setScrollbarVars(r,bg,fg,acc);}";
    html += "const src='" + escapedMarkdown + "';";
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
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_SAVE_AS_FAILED));
        return HRESULT_FROM_WIN32(lastError);
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    const HRESULT ioHr = _fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (FAILED(ioHr) || ! fileIo)
    {
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_SAVE_AS_NO_FILE_IO));
        return FAILED(ioHr) ? ioHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    wil::com_ptr<IFileReader> reader;
    const HRESULT openHr = fileIo->CreateFileReader(_currentPath.c_str(), reader.put());
    if (FAILED(openHr) || ! reader)
    {
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_SAVE_AS_OPEN_FAILED));
        return FAILED(openHr) ? openHr : E_FAIL;
    }

    uint64_t pos         = 0;
    const HRESULT seekHr = reader->Seek(0, FILE_BEGIN, &pos);
    if (FAILED(seekHr))
    {
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_SAVE_AS_SEEK_FAILED));
        return seekHr;
    }

    std::vector<uint8_t> buffer(256u * 1024u);
    for (;;)
    {
        unsigned long read   = 0;
        const HRESULT readHr = reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &read);
        if (FAILED(readHr))
        {
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_SAVE_AS_READ_FAILED));
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
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_SAVE_AS_WRITE_FAILED));
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
            if (! OrdinalString::StartsWithNoCase(src, L"about:") && ! IsInternalDocumentUrl(src))
            {
                toCopy.assign(src);
            }
        }
    }

    if (toCopy.empty())
    {
        if (_kind == ViewerWebKind::Web)
        {
            if (_tempExtractedPath.has_value() && ! _tempExtractedPath.value().empty())
            {
                toCopy = UrlFromFilePath(_tempExtractedPath.value().wstring());
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
            if (! OrdinalString::StartsWithNoCase(src, L"about:") && ! IsInternalDocumentUrl(src))
            {
                url.assign(src);
            }
        }
    }

    if (url.empty())
    {
        if (_kind == ViewerWebKind::Web)
        {
            if (_tempExtractedPath.has_value() && ! _tempExtractedPath.value().empty())
            {
                url = UrlFromFilePath(_tempExtractedPath.value().wstring());
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
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_OPEN_BROWSER_FAILED));
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
        ShowHostAlert(_hWnd.get(), HOST_ALERT_WARNING, LoadStringResource(g_hInstance, IDS_VIEWERWEB_WARNING_DEVTOOLS_DISABLED));
        return;
    }

    if (_webView)
    {
        static_cast<void>(_webView->OpenDevToolsWindow());
    }
}

void ViewerWeb::CommandJsonExpandAll() noexcept
{
    if (_kind != ViewerWebKind::Json || ! _jsonExpandCollapseAvailable || ! _webView)
    {
        return;
    }

    static_cast<void>(_webView->ExecuteScript(L"(function(){try{if(window.RS&&window.RS.expandAll){window.RS.expandAll();}}catch(e){}})();", nullptr));
}

void ViewerWeb::CommandJsonCollapseAll() noexcept
{
    if (_kind != ViewerWebKind::Json || ! _jsonExpandCollapseAvailable || ! _webView)
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
        HMENU menu = _menuHandle ? _menuHandle.get() : GetMenu(_hWnd.get());
        if (menu)
        {
            CheckMenuItem(
                menu, IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE, static_cast<UINT>(MF_BYCOMMAND | (_markdownShowSource ? MF_CHECKED : MF_UNCHECKED)));
            if (_menuBarHost.GetHwnd())
            {
                _menuBarHost.SyncMenuModel();
            }
            else
            {
                DrawMenuBar(_hWnd.get());
            }
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

    if (OrdinalString::StartsWithNoCase(path, L"\\\\") || OrdinalString::StartsWithNoCase(path, L"//"))
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

[[nodiscard]] std::string EscapeJavaScriptStringUtf8(std::string_view text) noexcept
{
    std::string out;
    out.reserve(text.size() + text.size() / 8);
    for (const char ch : text)
    {
        const auto u = static_cast<uint8_t>(ch);
        switch (ch)
        {
            case '\\': out += "\\\\"; break;
            case '\'': out += "\\'"; break;
            case '\"': out += "\\\""; break;
            case '\r': out += "\\r"; break;
            case '\n': out += "\\n"; break;
            case '\t': out += "\\t"; break;
            default:
                if (u < 0x20)
                {
                    out += std::format("\\x{:02X}", static_cast<unsigned int>(u));
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

void ResetSharedEnvironment() noexcept
{
    ResetSharedEnvironmentImpl();
}
