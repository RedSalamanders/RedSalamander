#include "FunctionBar.h"

#include <algorithm>
#include <cwchar>
#include <d2d1.h>
#include <mutex>
#include <string>
#include <unordered_set>
#include <utility>
#include <windowsx.h>

#include "CommandRegistry.h"
#include "DxUi/DxUi.Typography.h"
#include "Helpers.h"
#include "ShortcutManager.h"
#include "WindowMessages.h"
#include "resource.h"

namespace
{
constexpr int kFunctionKeyCount                     = 12;
constexpr int kZonePaddingXDip                      = 6;
constexpr int kZonePaddingYDip                      = 2;
constexpr int kGlyphPaddingXDip                     = 3;
constexpr int kGlyphPaddingYDip                     = 0;
constexpr int kGlyphCornerRadiusDip                 = 2;
constexpr int kMinZoneWidthForModifiers             = 50;
constexpr int kKeyFontHeightDip                     = 7;
constexpr int kTextFontHeightDip                    = 11;
constexpr int kModifiersGapDip                      = 6;
constexpr wchar_t kFunctionBarRenderResourcesProp[] = L"RedSalamander.FunctionBar.RenderResources";

struct FunctionBarRenderResources final
{
    UINT dpi = USER_DEFAULT_SCREEN_DPI;
    wil::com_ptr<ID2D1Factory> d2dFactory;
    wil::com_ptr<IDWriteFactory> dwriteFactory;
    wil::com_ptr<IDWriteInlineObject> ellipsisSign;
    wil::com_ptr<ID2D1HwndRenderTarget> target;
    wil::com_ptr<IDWriteTextFormat> keyFormat;
    wil::com_ptr<IDWriteTextFormat> labelFormat;
    wil::com_ptr<IDWriteTextFormat> modifierFormat;
    wil::com_ptr<ID2D1SolidColorBrush> solidBrush;
};

[[nodiscard]] float PixelToDip(float value, float dpi) noexcept
{
    return (value * static_cast<float>(USER_DEFAULT_SCREEN_DPI)) / std::max(1.0f, dpi);
}

[[nodiscard]] D2D1_RECT_F RectF(const RECT& rc, float dpi) noexcept
{
    return D2D1::RectF(PixelToDip(static_cast<float>(rc.left), dpi),
                       PixelToDip(static_cast<float>(rc.top), dpi),
                       PixelToDip(static_cast<float>(rc.right), dpi),
                       PixelToDip(static_cast<float>(rc.bottom), dpi));
}

[[nodiscard]] D2D1_COLOR_F D2DColor(COLORREF color) noexcept
{
    return D2D1::ColorF(
        static_cast<float>(GetRValue(color)) / 255.0f, static_cast<float>(GetGValue(color)) / 255.0f, static_cast<float>(GetBValue(color)) / 255.0f, 1.0f);
}

[[nodiscard]] FunctionBarRenderResources* GetFunctionBarRenderResources(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return nullptr;
    }

    return reinterpret_cast<FunctionBarRenderResources*>(GetPropW(hwnd, kFunctionBarRenderResourcesProp));
}

void ResetFunctionBarTarget(FunctionBarRenderResources& resources) noexcept
{
    resources.solidBrush.reset();
    resources.target.reset();
}

void ResetFunctionBarTypography(FunctionBarRenderResources& resources) noexcept
{
    resources.ellipsisSign.reset();
    resources.keyFormat.reset();
    resources.labelFormat.reset();
    resources.modifierFormat.reset();
    ResetFunctionBarTarget(resources);
}

void WarnMissingFunctionBarLabelOnce(std::wstring_view commandId)
{
    static std::mutex mutex;
    static std::unordered_set<std::wstring> warnedCommandIds;

    std::scoped_lock lock(mutex);
    if (warnedCommandIds.emplace(commandId).second)
    {
        Debug::Warning(L"FunctionBar: command '{}' has no short or fallback display label resource.", commandId);
    }
}

[[nodiscard]] FunctionBarRenderResources* EnsureFunctionBarRenderResources(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return nullptr;
    }

    if (auto* existing = GetFunctionBarRenderResources(hwnd))
    {
        return existing;
    }

    auto resources = std::make_unique<FunctionBarRenderResources>();
    if (! SetPropW(hwnd, kFunctionBarRenderResourcesProp, reinterpret_cast<HANDLE>(resources.get())))
    {
        return nullptr;
    }

    return resources.release();
}

void DestroyFunctionBarRenderResources(HWND hwnd) noexcept
{
    auto* resources = GetFunctionBarRenderResources(hwnd);
    if (! resources)
    {
        return;
    }

    RemovePropW(hwnd, kFunctionBarRenderResourcesProp);
    std::unique_ptr<FunctionBarRenderResources> cleanup(resources);
}

[[nodiscard]] bool EnsureFunctionBarFactories(FunctionBarRenderResources& resources) noexcept
{
    if (! resources.d2dFactory)
    {
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, resources.d2dFactory.addressof());
        if (FAILED(hr))
        {
            resources.d2dFactory.reset();
            return false;
        }
    }

    if (! resources.dwriteFactory)
    {
        const HRESULT hr =
            DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(resources.dwriteFactory.put()));
        if (FAILED(hr))
        {
            resources.dwriteFactory.reset();
            return false;
        }
    }

    return resources.d2dFactory && resources.dwriteFactory;
}

void ConfigureFunctionBarTextFormat(IDWriteTextFormat* format, DWRITE_TEXT_ALIGNMENT alignment) noexcept
{
    if (! format)
    {
        return;
    }

    static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    static_cast<void>(format->SetTextAlignment(alignment));
    static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
}

void ApplyFunctionBarTextTrimming(IDWriteTextFormat* format, IDWriteInlineObject* ellipsisSign) noexcept
{
    if (! format || ! ellipsisSign)
    {
        return;
    }

    DWRITE_TRIMMING trimming{};
    trimming.granularity = DWRITE_TRIMMING_GRANULARITY_CHARACTER;
    static_cast<void>(format->SetTrimming(&trimming, ellipsisSign));
}

[[nodiscard]] HRESULT CreateFunctionBarTextFormat(FunctionBarRenderResources& resources,
                                                  const RedSalamander::DxUi::Typography::TypographySpec& spec,
                                                  DWRITE_TEXT_ALIGNMENT alignment,
                                                  IDWriteTextFormat** outFormat) noexcept
{
    if (! resources.dwriteFactory || ! outFormat)
    {
        return E_INVALIDARG;
    }

    const HRESULT hr = RedSalamander::DxUi::Typography::CreateTextFormat(resources.dwriteFactory.get(), spec, outFormat, L"");
    if (FAILED(hr) || ! *outFormat)
    {
        return FAILED(hr) ? hr : E_FAIL;
    }

    ConfigureFunctionBarTextFormat(*outFormat, alignment);
    ApplyFunctionBarTextTrimming(*outFormat, resources.ellipsisSign.get());
    return S_OK;
}

[[nodiscard]] bool EnsureFunctionBarTextFormats(HWND hwnd, FunctionBarRenderResources& resources) noexcept
{
    const UINT dpi = RedSalamander::DxUi::Typography::GetEffectiveDpi(hwnd);
    if (resources.dpi != dpi)
    {
        resources.dpi = dpi;
        ResetFunctionBarTypography(resources);
    }

    if (! EnsureFunctionBarFactories(resources))
    {
        return false;
    }

    const auto keySpec  = RedSalamander::DxUi::Typography::MakeUiTextSpec(static_cast<float>(kKeyFontHeightDip));
    const auto textSpec = RedSalamander::DxUi::Typography::MakeUiTextSpec(static_cast<float>(kTextFontHeightDip));

    if (! resources.keyFormat)
    {
        if (FAILED(CreateFunctionBarTextFormat(resources, keySpec, DWRITE_TEXT_ALIGNMENT_CENTER, resources.keyFormat.put())))
        {
            resources.keyFormat.reset();
            return false;
        }
    }
    if (! resources.labelFormat)
    {
        if (FAILED(CreateFunctionBarTextFormat(resources, textSpec, DWRITE_TEXT_ALIGNMENT_LEADING, resources.labelFormat.put())))
        {
            resources.labelFormat.reset();
            return false;
        }
    }
    if (! resources.modifierFormat)
    {
        if (FAILED(CreateFunctionBarTextFormat(resources, textSpec, DWRITE_TEXT_ALIGNMENT_TRAILING, resources.modifierFormat.put())))
        {
            resources.modifierFormat.reset();
            return false;
        }
    }

    if (! resources.ellipsisSign)
    {
        const HRESULT hr = resources.dwriteFactory->CreateEllipsisTrimmingSign(resources.labelFormat.get(), resources.ellipsisSign.addressof());
        if (SUCCEEDED(hr) && resources.ellipsisSign)
        {
            ApplyFunctionBarTextTrimming(resources.labelFormat.get(), resources.ellipsisSign.get());
            ApplyFunctionBarTextTrimming(resources.modifierFormat.get(), resources.ellipsisSign.get());
        }
    }

    return resources.keyFormat && resources.labelFormat && resources.modifierFormat;
}

[[nodiscard]] bool EnsureFunctionBarTarget(HWND hwnd, FunctionBarRenderResources& resources) noexcept
{
    if (! EnsureFunctionBarFactories(resources))
    {
        return false;
    }

    RECT client{};
    if (! hwnd || ! GetClientRect(hwnd, &client))
    {
        return false;
    }

    const UINT width  = static_cast<UINT>(std::max(0L, client.right - client.left));
    const UINT height = static_cast<UINT>(std::max(0L, client.bottom - client.top));
    if (width == 0u || height == 0u)
    {
        return false;
    }

    if (! resources.target)
    {
        const D2D1_RENDER_TARGET_PROPERTIES props          = D2D1::RenderTargetProperties();
        const D2D1_HWND_RENDER_TARGET_PROPERTIES hwndProps = D2D1::HwndRenderTargetProperties(hwnd, D2D1::SizeU(width, height));

        wil::com_ptr<ID2D1HwndRenderTarget> target;
        const HRESULT hr = resources.d2dFactory->CreateHwndRenderTarget(props, hwndProps, target.addressof());
        if (FAILED(hr) || ! target)
        {
            resources.target.reset();
            return false;
        }

        target->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
        target->SetDpi(static_cast<float>(resources.dpi), static_cast<float>(resources.dpi));
        resources.target = std::move(target);
    }

    if (! resources.solidBrush)
    {
        const HRESULT brushHr = resources.target->CreateSolidColorBrush(D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f), resources.solidBrush.addressof());
        if (FAILED(brushHr))
        {
            resources.solidBrush.reset();
            return false;
        }
    }

    return resources.target && resources.solidBrush;
}

void DrawFunctionBarRect(FunctionBarRenderResources& resources, const RECT& rc, COLORREF color) noexcept
{
    if (! resources.target || ! resources.solidBrush || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return;
    }

    resources.solidBrush->SetColor(D2DColor(color));
    resources.target->FillRectangle(RectF(rc, static_cast<float>(resources.dpi)), resources.solidBrush.get());
}

void DrawFunctionBarRoundedRect(FunctionBarRenderResources& resources, const RECT& rc, float radiusPx, COLORREF color) noexcept
{
    if (! resources.target || ! resources.solidBrush || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return;
    }

    resources.solidBrush->SetColor(D2DColor(color));
    const float dpi                 = static_cast<float>(resources.dpi);
    const float radiusDip           = PixelToDip(radiusPx, dpi);
    const float strokeDip           = PixelToDip(1.0f, dpi);
    const D2D1_ROUNDED_RECT rounded = D2D1::RoundedRect(RectF(rc, dpi), radiusDip, radiusDip);
    resources.target->DrawRoundedRectangle(rounded, resources.solidBrush.get(), strokeDip);
}

void DrawFunctionBarSeparator(FunctionBarRenderResources& resources, float x, float top, float bottom, COLORREF color) noexcept
{
    if (! resources.target || ! resources.solidBrush || bottom <= top)
    {
        return;
    }

    resources.solidBrush->SetColor(D2DColor(color));
    const float dpi       = static_cast<float>(resources.dpi);
    const float strokeDip = PixelToDip(1.0f, dpi);
    resources.target->DrawLine(D2D1::Point2F(PixelToDip(x, dpi), PixelToDip(top, dpi)),
                               D2D1::Point2F(PixelToDip(x, dpi), PixelToDip(bottom, dpi)),
                               resources.solidBrush.get(),
                               strokeDip);
}

void DrawFunctionBarText(FunctionBarRenderResources& resources, std::wstring_view text, const RECT& rc, IDWriteTextFormat* format, COLORREF color) noexcept
{
    if (! resources.target || ! resources.solidBrush || ! format || text.empty() || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return;
    }

    resources.solidBrush->SetColor(D2DColor(color));
    resources.target->DrawText(text.data(),
                               static_cast<UINT32>(text.size()),
                               format,
                               RectF(rc, static_cast<float>(resources.dpi)),
                               resources.solidBrush.get(),
                               D2D1_DRAW_TEXT_OPTIONS_CLIP,
                               DWRITE_MEASURING_MODE_NATURAL);
}

[[nodiscard]] COLORREF BlendColor(COLORREF base, COLORREF overlay, int overlayWeight, int denom) noexcept
{
    if (denom <= 0)
    {
        return base;
    }
    overlayWeight        = std::clamp(overlayWeight, 0, denom);
    const int baseWeight = denom - overlayWeight;

    const int r = (static_cast<int>(GetRValue(base)) * baseWeight + static_cast<int>(GetRValue(overlay)) * overlayWeight) / denom;
    const int g = (static_cast<int>(GetGValue(base)) * baseWeight + static_cast<int>(GetGValue(overlay)) * overlayWeight) / denom;
    const int b = (static_cast<int>(GetBValue(base)) * baseWeight + static_cast<int>(GetBValue(overlay)) * overlayWeight) / denom;
    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}

[[nodiscard]] std::wstring BuildModifierText(uint32_t modifiers) noexcept
{
    std::wstring result;

    auto append = [&](UINT stringId) noexcept
    {
        std::wstring text;
        if (stringId == IDS_MOD_CTRL)
        {
            text = LoadEmbeddedStringResource(nullptr, stringId);
        }
        else
        {
            text = LoadStringResource(nullptr, stringId);
        }
        if (text.empty())
        {
            return;
        }

        if (! result.empty())
        {
            result.push_back(L'+');
        }

        result.append(text);
    };

    if ((modifiers & ShortcutManager::kModCtrl) != 0)
    {
        append(IDS_MOD_CTRL);
    }
    if ((modifiers & ShortcutManager::kModAlt) != 0)
    {
        append(IDS_MOD_ALT);
    }
    if ((modifiers & ShortcutManager::kModShift) != 0)
    {
        append(IDS_MOD_SHIFT);
    }

    return result;
}

[[nodiscard]] RedSalamander::DxUi::Typography::TextPixelMetrics MeasureFunctionBarTextMetrics(HWND hwnd,
                                                                                              const RedSalamander::DxUi::Typography::TypographySpec& spec,
                                                                                              std::wstring_view text) noexcept
{
    IDWriteFactory* dwriteFactory = RedSalamander::DxUi::Typography::GetSharedMeasurementFactory();
    if (! dwriteFactory)
    {
        return {};
    }

    wil::com_ptr<IDWriteTextFormat> textFormat;
    const HRESULT hr = RedSalamander::DxUi::Typography::CreateTextFormat(dwriteFactory, spec, textFormat.put(), L"");
    if (FAILED(hr) || ! textFormat)
    {
        return {};
    }

    ConfigureFunctionBarTextFormat(textFormat.get(), DWRITE_TEXT_ALIGNMENT_LEADING);
    return RedSalamander::DxUi::Typography::MeasureSingleLineTextMetrics(
        dwriteFactory, textFormat.get(), RedSalamander::DxUi::Typography::GetEffectiveDpi(hwnd), text);
}

} // namespace

FunctionBar::FunctionBar() = default;

FunctionBar::~FunctionBar()
{
    Destroy();
}

ATOM FunctionBar::RegisterWndClass(HINSTANCE instance)
{
    static ATOM atom = 0;
    if (atom)
    {
        return atom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProcThunk;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kClassName;

    atom = RegisterClassExW(&wc);
    return atom;
}

HWND FunctionBar::Create(HWND parent, int x, int y, int width, int height)
{
    _hInstance = GetModuleHandleW(nullptr);

    if (! RegisterWndClass(_hInstance))
    {
        return nullptr;
    }

    CreateWindowExW(0, kClassName, L"", WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS, x, y, width, height, parent, nullptr, _hInstance, this);

    return _hWnd.get();
}

void FunctionBar::Destroy() noexcept
{
    DestroyFunctionBarRenderResources(_hWnd.get());
    _hWnd.reset();
}

void FunctionBar::SetTheme(const AppTheme& theme)
{
    _theme = theme;

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void FunctionBar::SetShortcutManager(const ShortcutManager* shortcuts) noexcept
{
    _shortcuts = shortcuts;
    RecomputeLabels();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void FunctionBar::SetDpi(UINT dpi) noexcept
{
    _dpi = dpi;
    if (_hWnd)
    {
        if (auto* resources = GetFunctionBarRenderResources(_hWnd.get()))
        {
            ResetFunctionBarTypography(*resources);
        }
    }
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void FunctionBar::SetModifiers(uint32_t modifiers) noexcept
{
    modifiers = modifiers & 0x7u;
    if (_modifiers == modifiers)
    {
        return;
    }

    _modifiers = modifiers;
    RecomputeLabels();
    RecomputeModifierText();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }
}

void FunctionBar::SetPressedFunctionKey(std::optional<uint32_t> vk) noexcept
{
    if (_pressedKey == vk)
    {
        return;
    }

    _pressedKey = vk;
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

LRESULT CALLBACK FunctionBar::WndProcThunk(HWND hWindow, UINT msg, WPARAM wp, LPARAM lp)
{
    FunctionBar* self = nullptr;

    if (msg == WM_NCCREATE)
    {
        auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        self    = reinterpret_cast<FunctionBar*>(cs->lpCreateParams);
        SetWindowLongPtrW(hWindow, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        self->_hWnd.reset(hWindow);
    }
    else
    {
        self = reinterpret_cast<FunctionBar*>(GetWindowLongPtrW(hWindow, GWLP_USERDATA));
    }

    if (self)
    {
        return self->WndProc(hWindow, msg, wp, lp);
    }

    return DefWindowProcW(hWindow, msg, wp, lp);
}

LRESULT FunctionBar::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg)
    {
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_ERASEBKGND: return 1;
        case WM_PAINT: OnPaint(); return 0;
        case WM_DPICHANGED_AFTERPARENT: SetDpi(GetDpiForWindow(hwnd)); return 0;
        case WM_SIZE: OnSize(LOWORD(lp), HIWORD(lp)); return 0;
        case WM_KEYDOWN:
        case WM_SYSKEYDOWN:
            if (OnKeyDown(wp))
            {
                return 0;
            }
            break;
        case WM_MOUSEMOVE: OnMouseMove({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_LBUTTONUP: OnLButtonUp({GET_X_LPARAM(lp), GET_Y_LPARAM(lp)}); return 0;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void FunctionBar::OnCreate(HWND hwnd) noexcept
{
    _dpi = GetDpiForWindow(hwnd);
    RecomputeLabels();
    RecomputeModifierText();
}

void FunctionBar::OnDestroy() noexcept
{
    DestroyFunctionBarRenderResources(_hWnd.get());
}

void FunctionBar::OnSize(UINT width, UINT height) noexcept
{
    _clientSize = {static_cast<LONG>(width), static_cast<LONG>(height)};
    if (_hWnd)
    {
        if (auto* resources = GetFunctionBarRenderResources(_hWnd.get()))
        {
            ResetFunctionBarTarget(*resources);
        }
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool FunctionBar::OnKeyDown(WPARAM key) noexcept
{
    if (key != VK_ESCAPE || ! _hWnd)
    {
        return false;
    }

    const HWND parent = GetParent(_hWnd.get());
    if (parent)
    {
        SendMessageW(parent, WndMsg::kPaneRestoreFolderFocus, 0, 0);
    }
    return true;
}

void FunctionBar::OnMouseMove(POINT pt) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    if (! _trackingMouseLeave)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = _hWnd.get();
        if (TrackMouseEvent(&tme))
        {
            _trackingMouseLeave = true;
        }
    }

    const std::optional<uint32_t> key = HitTestFunctionKey(pt);
    if (_hoveredKey == key)
    {
        return;
    }

    _hoveredKey = key;
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void FunctionBar::OnMouseLeave() noexcept
{
    _trackingMouseLeave = false;
    if (! _hoveredKey.has_value() || ! _hWnd)
    {
        _hoveredKey.reset();
        return;
    }

    _hoveredKey.reset();
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void FunctionBar::OnLButtonUp(POINT pt) noexcept
{
    if (! _hWnd)
    {
        return;
    }

    const std::optional<uint32_t> vkOpt = HitTestFunctionKey(pt);
    if (! vkOpt.has_value())
    {
        return;
    }

    uint32_t modifiers = 0;
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModCtrl;
    }
    if ((GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModAlt;
    }
    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        modifiers |= ShortcutManager::kModShift;
    }

    const HWND owner = GetAncestor(_hWnd.get(), GA_ROOT);
    if (! owner)
    {
        return;
    }

    SendMessageW(owner, WndMsg::kFunctionBarInvoke, static_cast<WPARAM>(vkOpt.value()), static_cast<LPARAM>(modifiers & 0x7u));
}

std::optional<uint32_t> FunctionBar::HitTestFunctionKey(POINT pt) const noexcept
{
    if (! _hWnd)
    {
        return std::nullopt;
    }

    const int width  = static_cast<int>(_clientSize.cx);
    const int height = static_cast<int>(_clientSize.cy);
    if (width <= 0 || height <= 0)
    {
        return std::nullopt;
    }

    if (pt.x < 0 || pt.y < 0 || pt.x >= width || pt.y >= height)
    {
        return std::nullopt;
    }

    const int paddingX          = PxFromDip(kZonePaddingXDip);
    const int modifiersGap      = std::max(1, PxFromDip(kModifiersGapDip) / 2);
    const int modifiersPaddingX = std::max(1, paddingX / 2);

    int reservedModifiersWidth = 0;
    {
        const std::wstring maxModifiersText = BuildModifierText(ShortcutManager::kModCtrl | ShortcutManager::kModAlt | ShortcutManager::kModShift);
        if (! maxModifiersText.empty())
        {
            const auto modMetrics = MeasureFunctionBarTextMetrics(
                _hWnd.get(), RedSalamander::DxUi::Typography::MakeUiTextSpec(static_cast<float>(kTextFontHeightDip)), maxModifiersText);
            reservedModifiersWidth = modMetrics.widthPx + (modifiersPaddingX * 2) + modifiersGap;

            const int minZoneWidthPx = PxFromDip(kMinZoneWidthForModifiers);
            if ((width - reservedModifiersWidth) / kFunctionKeyCount < minZoneWidthPx)
            {
                reservedModifiersWidth = 0;
            }
        }
    }

    const int zonesWidth = std::max(0, width - reservedModifiersWidth);
    if (zonesWidth <= 0 || pt.x >= zonesWidth)
    {
        return std::nullopt;
    }

    const int index = std::clamp((static_cast<int>(pt.x) * kFunctionKeyCount) / zonesWidth, 0, kFunctionKeyCount - 1);
    return static_cast<uint32_t>(VK_F1 + index);
}

void FunctionBar::OnPaint() noexcept
{
    if (! _hWnd)
    {
        return;
    }

    Debug::Perf::Scope paintPerf(L"render.function_bar.paint_us");

    PAINTSTRUCT ps;
    wil::unique_hdc_paint paint = wil::BeginPaint(_hWnd.get(), &ps);
    if (! paint)
    {
        return;
    }

    PaintToHdc(paint.get());
}

void FunctionBar::PaintToHdc(HDC hdc) noexcept
{
    if (! _hWnd || ! hdc)
    {
        return;
    }

    RECT client{};
    if (! GetClientRect(_hWnd.get(), &client))
    {
        return;
    }

    const int width  = std::max(0l, client.right - client.left);
    const int height = std::max(0l, client.bottom - client.top);
    if (width <= 0 || height <= 0)
    {
        return;
    }

    auto* resources = EnsureFunctionBarRenderResources(_hWnd.get());
    if (! resources || ! EnsureFunctionBarTextFormats(_hWnd.get(), *resources) || ! EnsureFunctionBarTarget(_hWnd.get(), *resources))
    {
        return;
    }

    resources->target->BeginDraw();
    auto endDraw = wil::scope_exit([&]() noexcept
    {
        const HRESULT endHr = resources->target->EndDraw();
        if (FAILED(endHr))
        {
            if (endHr == D2DERR_RECREATE_TARGET)
            {
                ResetFunctionBarTarget(*resources);
            }
        }
    });

    DrawFunctionBarRect(*resources, client, _theme.menu.background);

    const int paddingX          = PxFromDip(kZonePaddingXDip);
    const int paddingY          = PxFromDip(kZonePaddingYDip);
    const int glyphPadX         = PxFromDip(kGlyphPaddingXDip);
    const int glyphPadY         = PxFromDip(kGlyphPaddingYDip);
    const int modifiersGap      = std::max(1, PxFromDip(kModifiersGapDip) / 2);
    const int modifiersPaddingX = std::max(1, paddingX / 2);

    int reservedModifiersWidth = 0;
    {
        const std::wstring maxModifiersText = BuildModifierText(ShortcutManager::kModCtrl | ShortcutManager::kModAlt | ShortcutManager::kModShift);
        if (! maxModifiersText.empty())
        {
            const auto modMetrics = RedSalamander::DxUi::Typography::MeasureSingleLineTextMetrics(
                resources->dwriteFactory.get(), resources->modifierFormat.get(), resources->dpi, maxModifiersText);
            reservedModifiersWidth = modMetrics.widthPx + (modifiersPaddingX * 2) + modifiersGap;

            const int minZoneWidthPx = PxFromDip(kMinZoneWidthForModifiers);
            if ((width - reservedModifiersWidth) / kFunctionKeyCount < minZoneWidthPx)
            {
                reservedModifiersWidth = 0;
            }
        }
    }

    const bool showModifiers = reservedModifiersWidth > 0 && ! _modifierText.empty();
    const int zonesWidth     = std::max(0, width - reservedModifiersWidth);

    const COLORREF textColorNormal  = _theme.menu.text;
    const COLORREF textColorPressed = _theme.menu.selectionText;

    for (int i = 0; i < kFunctionKeyCount; ++i)
    {
        const int left  = (i * zonesWidth) / kFunctionKeyCount;
        const int right = ((i + 1) * zonesWidth) / kFunctionKeyCount;

        RECT zone{};
        zone.left   = left;
        zone.top    = 0;
        zone.right  = right;
        zone.bottom = height;

        const uint32_t vk  = static_cast<uint32_t>(VK_F1 + i);
        const bool pressed = _pressedKey.has_value() && _pressedKey.value() == vk;
        const bool hovered = _hoveredKey.has_value() && _hoveredKey.value() == vk;
        if (pressed)
        {
            DrawFunctionBarRect(*resources, zone, _theme.menu.selectionBg);
        }
        else if (hovered)
        {
            const COLORREF hoverColor = _theme.highContrast ? _theme.menu.selectionBg : BlendColor(_theme.menu.background, _theme.menu.selectionBg, 1, 3);
            DrawFunctionBarRect(*resources, zone, hoverColor);
        }

        if (i > 0)
        {
            DrawFunctionBarSeparator(
                *resources, static_cast<float>(zone.left), static_cast<float>(zone.top), static_cast<float>(zone.bottom), _theme.menu.separator);
        }

        wchar_t keyLabel[4]{};
        keyLabel[0]           = L'F';
        const unsigned number = static_cast<unsigned>(i + 1);
        if (number < 10)
        {
            keyLabel[1] = static_cast<wchar_t>(L'0' + number);
            keyLabel[2] = L'\0';
        }
        else
        {
            keyLabel[1] = static_cast<wchar_t>(L'0' + (number / 10));
            keyLabel[2] = static_cast<wchar_t>(L'0' + (number % 10));
            keyLabel[3] = L'\0';
        }
        const int keyLabelLength = static_cast<int>(std::wcslen(keyLabel));
        const auto keyMetrics    = RedSalamander::DxUi::Typography::MeasureSingleLineTextMetrics(
            resources->dwriteFactory.get(), resources->keyFormat.get(), resources->dpi, std::wstring_view(keyLabel, static_cast<size_t>(keyLabelLength)));
        const int keyWidthPx  = std::max(1, keyMetrics.widthPx);
        const int keyHeightPx = std::max(1, keyMetrics.lineHeightPx);

        const int availableHeight    = std::max(1, height - (paddingY * 2));
        const int desiredGlyphHeight = keyHeightPx + (glyphPadY * 2);
        const int glyphHeight        = std::clamp(desiredGlyphHeight, 1, availableHeight);
        const int zoneWidthPx        = std::max(0, static_cast<int>(zone.right - zone.left));
        const int glyphWidth         = std::min(std::max(1, zoneWidthPx - (paddingX * 2)), keyWidthPx + (glyphPadX * 2));
        const int glyphTop           = zone.top + (height - glyphHeight) / 2;

        RECT glyph{};
        glyph.left   = zone.left + paddingX;
        glyph.top    = glyphTop;
        glyph.right  = glyph.left + std::max(1, glyphWidth);
        glyph.bottom = glyph.top + glyphHeight;

        const float radiusPx = static_cast<float>(std::clamp(PxFromDip(kGlyphCornerRadiusDip), 1, std::max(1, glyphHeight / 2)));
        DrawFunctionBarRoundedRect(*resources, glyph, radiusPx, _theme.menu.separator);
        RECT keyTextRect  = glyph;
        keyTextRect.left  = std::min(keyTextRect.right, keyTextRect.left + glyphPadX);
        keyTextRect.right = std::max(keyTextRect.left, keyTextRect.right - glyphPadX);
        DrawFunctionBarText(*resources,
                            std::wstring_view(keyLabel, static_cast<size_t>(keyLabelLength)),
                            keyTextRect,
                            resources->keyFormat.get(),
                            pressed ? textColorPressed : textColorNormal);

        RECT textRect{};
        textRect.left   = std::min(zone.right, glyph.right + paddingX);
        textRect.top    = zone.top;
        textRect.right  = std::max(textRect.left, zone.right - paddingX);
        textRect.bottom = zone.bottom;

        if (i >= 0 && i < static_cast<int>(_labels.size()))
        {
            DrawFunctionBarText(*resources, _labels[static_cast<size_t>(i)], textRect, resources->labelFormat.get(), textColorNormal);
        }
    }

    if (showModifiers)
    {
        RECT modRect{};
        modRect.left   = zonesWidth;
        modRect.top    = 0;
        modRect.right  = width;
        modRect.bottom = height;

        RECT modTextRect  = modRect;
        modTextRect.left  = std::min(modTextRect.right, modTextRect.left + modifiersPaddingX);
        modTextRect.right = std::max(modTextRect.left, modTextRect.right - modifiersPaddingX);
        DrawFunctionBarText(*resources, _modifierText, modTextRect, resources->modifierFormat.get(), textColorNormal);
    }
}

void FunctionBar::RecomputeLabels()
{
    for (auto& label : _labels)
    {
        label.clear();
    }

    if (! _shortcuts)
    {
        return;
    }

    for (int i = 0; i < kFunctionKeyCount; ++i)
    {
        const uint32_t vk                                 = static_cast<uint32_t>(VK_F1 + i);
        const std::optional<std::wstring_view> commandOpt = _shortcuts->FindFunctionBarCommand(vk, _modifiers);
        if (! commandOpt.has_value())
        {
            continue;
        }
        if (ShortcutIds::IsUnassignedCommandId(CanonicalizeCommandId(commandOpt.value())))
        {
            continue;
        }

        std::wstring label;
        if (const std::optional<unsigned int> shortDisplayNameId = TryGetCommandShortDisplayNameStringId(commandOpt.value()))
        {
            label = LoadStringResource(nullptr, shortDisplayNameId.value());
        }
        if (label.empty())
        {
            if (const std::optional<unsigned int> displayNameId = TryGetCommandDisplayNameStringId(commandOpt.value()))
            {
                label = LoadStringResource(nullptr, displayNameId.value());
            }
        }
        if (label.empty())
        {
            WarnMissingFunctionBarLabelOnce(commandOpt.value());
            label = L"?";
        }
        _labels[static_cast<size_t>(i)] = std::move(label);
    }
}

void FunctionBar::RecomputeModifierText()
{
    _modifierText = BuildModifierText(_modifiers);
}

int FunctionBar::PxFromDip(int dip) const noexcept
{
    return MulDiv(dip, static_cast<int>(_dpi), USER_DEFAULT_SCREEN_DPI);
}

#ifdef ENABLE_TESTS
bool FunctionBar::DebugUsesDirectWriteTextMetrics() const noexcept
{
    if (! _hWnd)
    {
        return false;
    }

    auto* resources = EnsureFunctionBarRenderResources(_hWnd.get());
    return resources && EnsureFunctionBarTextFormats(_hWnd.get(), *resources) && resources->dwriteFactory && resources->keyFormat && resources->labelFormat &&
           resources->modifierFormat;
}
#endif
