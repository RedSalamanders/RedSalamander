#include "FolderWindowInternal.h"

#include "ConnectionProfileUtils.h"
#include "DxUi/DxUi.Typography.h"
#include "FluentIcons.h"
#include "SettingsStore.h"

#include <d2d1.h>
#include <windowsx.h>

namespace
{
constexpr int kStatusBarFocusLineHeightDip        = 2;
constexpr int kStatusBarSecurityMinPartWidthDip   = 90;
constexpr int kStatusBarSecurityMaxPartWidthDip   = 240;
constexpr wchar_t kStatusBarRenderResourcesProp[] = L"RedSalamander.StatusBar.RenderResources";

struct StatusBarParts final
{
    RECT selection{};
    RECT security{};
    RECT sort{};
};

struct StatusBarRenderResources final
{
    UINT dpi                 = USER_DEFAULT_SCREEN_DPI;
    bool fluentIconAvailable = false;
    wil::com_ptr<ID2D1Factory> d2dFactory;
    wil::com_ptr<IDWriteFactory> dwriteFactory;
    wil::com_ptr<IDWriteInlineObject> ellipsisSign;
    wil::com_ptr<ID2D1HwndRenderTarget> target;
    wil::com_ptr<IDWriteTextFormat> selectionFormat;
    wil::com_ptr<IDWriteTextFormat> securityFormat;
    wil::com_ptr<IDWriteTextFormat> sortTextFormat;
    wil::com_ptr<IDWriteTextFormat> arrowFormat;
    wil::com_ptr<IDWriteTextFormat> iconLeadingFormat;
    wil::com_ptr<IDWriteTextFormat> iconCenteredFormat;
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

[[nodiscard]] StatusBarRenderResources* GetStatusBarRenderResources(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return nullptr;
    }

    return reinterpret_cast<StatusBarRenderResources*>(GetPropW(hwnd, kStatusBarRenderResourcesProp));
}

void ResetStatusBarTarget(StatusBarRenderResources& resources) noexcept
{
    resources.solidBrush.reset();
    resources.target.reset();
}

void ResetStatusBarTypography(StatusBarRenderResources& resources) noexcept
{
    resources.ellipsisSign.reset();
    resources.selectionFormat.reset();
    resources.securityFormat.reset();
    resources.sortTextFormat.reset();
    resources.arrowFormat.reset();
    resources.iconLeadingFormat.reset();
    resources.iconCenteredFormat.reset();
    ResetStatusBarTarget(resources);
}

[[nodiscard]] StatusBarRenderResources* EnsureStatusBarRenderResources(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return nullptr;
    }

    if (auto* existing = GetStatusBarRenderResources(hwnd))
    {
        return existing;
    }

    auto resources = std::make_unique<StatusBarRenderResources>();
    if (! SetPropW(hwnd, kStatusBarRenderResourcesProp, reinterpret_cast<HANDLE>(resources.get())))
    {
        return nullptr;
    }

    return resources.release();
}

void DestroyStatusBarRenderResources(HWND hwnd) noexcept
{
    auto* resources = GetStatusBarRenderResources(hwnd);
    if (! resources)
    {
        return;
    }

    RemovePropW(hwnd, kStatusBarRenderResourcesProp);
    std::unique_ptr<StatusBarRenderResources> cleanup(resources);
}

[[nodiscard]] bool EnsureStatusBarFactories(StatusBarRenderResources& resources) noexcept
{
    if (! resources.d2dFactory)
    {
        const HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, resources.d2dFactory.addressof());
        if (FAILED(hr))
        {
            resources.d2dFactory.reset();
        }
    }

    if (! resources.dwriteFactory)
    {
        const HRESULT hr =
            DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(resources.dwriteFactory.addressof()));
        if (FAILED(hr))
        {
            resources.dwriteFactory.reset();
        }
    }

    resources.fluentIconAvailable = resources.dwriteFactory && RedSalamander::DxUi::Typography::IsFontFamilyAvailable(
                                                                   resources.dwriteFactory.get(), RedSalamander::DxUi::Typography::kSegoeFluentIconsFamily);
    return resources.d2dFactory && resources.dwriteFactory;
}

void ConfigureStatusBarTextFormat(IDWriteTextFormat* format, DWRITE_TEXT_ALIGNMENT alignment) noexcept
{
    if (! format)
    {
        return;
    }

    static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    static_cast<void>(format->SetTextAlignment(alignment));
    static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
}

void ApplyStatusBarTextTrimming(IDWriteTextFormat* format, IDWriteInlineObject* ellipsisSign) noexcept
{
    if (! format || ! ellipsisSign)
    {
        return;
    }

    DWRITE_TRIMMING trimming{};
    trimming.granularity = DWRITE_TRIMMING_GRANULARITY_CHARACTER;
    static_cast<void>(format->SetTrimming(&trimming, ellipsisSign));
}

[[nodiscard]] HRESULT CreateStatusBarTextFormat(StatusBarRenderResources& resources,
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

    ConfigureStatusBarTextFormat(*outFormat, alignment);
    ApplyStatusBarTextTrimming(*outFormat, resources.ellipsisSign.get());
    return S_OK;
}

[[nodiscard]] bool EnsureStatusBarTextFormats(HWND hwnd, StatusBarRenderResources& resources) noexcept
{
    const UINT dpi = RedSalamander::DxUi::Typography::GetEffectiveDpi(hwnd);
    if (resources.dpi != dpi)
    {
        resources.dpi = dpi;
        ResetStatusBarTypography(resources);
    }

    if (! EnsureStatusBarFactories(resources))
    {
        return false;
    }

    const auto textSpec = RedSalamander::DxUi::Typography::MakeUiTextSpec(kStatusBarTextSizeDip);
    const auto iconSpec = RedSalamander::DxUi::Typography::MakeUiIconSpec(FluentIcons::kDefaultSizeDip);

    if (! resources.selectionFormat)
    {
        if (FAILED(CreateStatusBarTextFormat(resources, textSpec, DWRITE_TEXT_ALIGNMENT_LEADING, resources.selectionFormat.put())))
        {
            resources.selectionFormat.reset();
            return false;
        }
    }
    if (! resources.securityFormat)
    {
        if (FAILED(CreateStatusBarTextFormat(resources, textSpec, DWRITE_TEXT_ALIGNMENT_CENTER, resources.securityFormat.put())))
        {
            resources.securityFormat.reset();
            return false;
        }
    }
    if (! resources.sortTextFormat)
    {
        if (FAILED(CreateStatusBarTextFormat(resources, textSpec, DWRITE_TEXT_ALIGNMENT_TRAILING, resources.sortTextFormat.put())))
        {
            resources.sortTextFormat.reset();
            return false;
        }
    }
    if (! resources.arrowFormat)
    {
        if (FAILED(CreateStatusBarTextFormat(resources, textSpec, DWRITE_TEXT_ALIGNMENT_TRAILING, resources.arrowFormat.put())))
        {
            resources.arrowFormat.reset();
            return false;
        }
    }

    if (! resources.ellipsisSign)
    {
        const HRESULT hr = resources.dwriteFactory->CreateEllipsisTrimmingSign(resources.selectionFormat.get(), resources.ellipsisSign.addressof());
        if (SUCCEEDED(hr) && resources.ellipsisSign)
        {
            ApplyStatusBarTextTrimming(resources.selectionFormat.get(), resources.ellipsisSign.get());
            ApplyStatusBarTextTrimming(resources.securityFormat.get(), resources.ellipsisSign.get());
            ApplyStatusBarTextTrimming(resources.sortTextFormat.get(), resources.ellipsisSign.get());
        }
    }

    if (resources.fluentIconAvailable && ! resources.iconLeadingFormat)
    {
        if (FAILED(CreateStatusBarTextFormat(resources, iconSpec, DWRITE_TEXT_ALIGNMENT_LEADING, resources.iconLeadingFormat.put())))
        {
            resources.iconLeadingFormat.reset();
        }
    }
    if (resources.fluentIconAvailable && ! resources.iconCenteredFormat)
    {
        if (FAILED(CreateStatusBarTextFormat(resources, iconSpec, DWRITE_TEXT_ALIGNMENT_CENTER, resources.iconCenteredFormat.put())))
        {
            resources.iconCenteredFormat.reset();
        }
    }

    if (resources.fluentIconAvailable && (! resources.iconLeadingFormat || ! resources.iconCenteredFormat))
    {
        resources.fluentIconAvailable = false;
        resources.iconLeadingFormat.reset();
        resources.iconCenteredFormat.reset();
    }

    return resources.selectionFormat && resources.securityFormat && resources.sortTextFormat && resources.arrowFormat;
}

[[nodiscard]] int MeasureStatusBarSecurityTextWidthPx(HWND hwnd, std::wstring_view text) noexcept
{
    if (! hwnd || text.empty())
    {
        return 0;
    }

    auto* resources = EnsureStatusBarRenderResources(hwnd);
    if (! resources || ! EnsureStatusBarTextFormats(hwnd, *resources) || ! resources->dwriteFactory || ! resources->securityFormat)
    {
        return 0;
    }

    return RedSalamander::DxUi::Typography::MeasureSingleLineTextMetrics(resources->dwriteFactory.get(), resources->securityFormat.get(), resources->dpi, text)
        .widthPx;
}

[[nodiscard]] bool EnsureStatusBarTarget(HWND hwnd, StatusBarRenderResources& resources) noexcept
{
    if (! EnsureStatusBarFactories(resources))
    {
        return false;
    }

    RECT client{};
    if (! GetClientRect(hwnd, &client))
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
        const HRESULT hr = resources.target->CreateSolidColorBrush(D2D1::ColorF(0.0f, 0.0f, 0.0f, 1.0f), resources.solidBrush.addressof());
        if (FAILED(hr))
        {
            resources.solidBrush.reset();
            return false;
        }
    }

    return resources.target && resources.solidBrush;
}

void InvalidateStatusBarDeviceResources(HWND hwnd, bool resetTypography) noexcept
{
    if (auto* resources = GetStatusBarRenderResources(hwnd))
    {
        if (resetTypography)
        {
            ResetStatusBarTypography(*resources);
        }
        else
        {
            ResetStatusBarTarget(*resources);
        }
    }
}

[[nodiscard]] bool StatusBarSupportsFluentIcons(HWND hwnd) noexcept
{
    auto* resources = EnsureStatusBarRenderResources(hwnd);
    if (! resources || ! EnsureStatusBarTextFormats(hwnd, *resources))
    {
        return false;
    }

    return resources->fluentIconAvailable && resources->iconLeadingFormat && resources->iconCenteredFormat;
}

[[nodiscard]] bool ComputeStatusBarParts(HWND hwnd, StatusBarParts& parts, RECT* outClientRect = nullptr) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    RECT client{};
    if (! GetClientRect(hwnd, &client))
    {
        return false;
    }

    const int width = std::max(0L, client.right - client.left);
    const int dpi   = std::max(1, static_cast<int>(GetDpiForWindow(hwnd)));

    const int minSortPartWidth = MulDiv(kStatusBarSortMinPartWidthDip, dpi, USER_DEFAULT_SCREEN_DPI);
    const int sortWidth        = std::clamp(minSortPartWidth, 0, width);

    int securityWidth        = 0;
    const auto* securityText = reinterpret_cast<const std::wstring*>(GetPropW(hwnd, kStatusBarSecurityTextProp));
    if (securityText && ! securityText->empty() && width > sortWidth)
    {
        const int paddingX = std::max(1, MulDiv(kStatusBarSortPaddingXDip, dpi, USER_DEFAULT_SCREEN_DPI));
        const int textPx   = MeasureStatusBarSecurityTextWidthPx(hwnd, *securityText);

        const int desired  = textPx + 2 * paddingX;
        const int minWidth = MulDiv(kStatusBarSecurityMinPartWidthDip, dpi, USER_DEFAULT_SCREEN_DPI);
        const int maxWidth = MulDiv(kStatusBarSecurityMaxPartWidthDip, dpi, USER_DEFAULT_SCREEN_DPI);
        securityWidth      = std::clamp(desired, minWidth, maxWidth);
        securityWidth      = std::clamp(securityWidth, 0, width - sortWidth);
    }

    const LONG selectionRight = static_cast<LONG>(std::max(0, width - sortWidth - securityWidth));
    const LONG securityRight  = static_cast<LONG>(std::max<int>(selectionRight, width - sortWidth));

    parts.selection = {client.left, client.top, selectionRight, client.bottom};
    parts.security  = {selectionRight, client.top, securityRight, client.bottom};
    parts.sort      = {securityRight, client.top, client.right, client.bottom};

    if (outClientRect)
    {
        *outClientRect = client;
    }

    return true;
}

void UpdateStatusBarParts(HWND hwnd) noexcept
{
    if (hwnd && IsWindow(hwnd) != FALSE)
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
}

[[nodiscard]] bool ContainsPrivateUseAreaGlyph(std::wstring_view text) noexcept
{
    for (wchar_t ch : text)
    {
        if (ch >= 0xE000 && ch <= 0xF8FF)
        {
            return true;
        }
    }
    return false;
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

void FillStatusBarRect(StatusBarRenderResources& resources, const RECT& rc, COLORREF color) noexcept
{
    if (! resources.target || ! resources.solidBrush || rc.right <= rc.left || rc.bottom <= rc.top)
    {
        return;
    }

    resources.solidBrush->SetColor(D2DColor(color));
    resources.target->FillRectangle(RectF(rc, static_cast<float>(resources.dpi)), resources.solidBrush.get());
}

void DrawStatusBarText(StatusBarRenderResources& resources, std::wstring_view text, const RECT& rc, IDWriteTextFormat* format, COLORREF color) noexcept
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

void PaintSortIndicatorGlyph(StatusBarRenderResources& resources, const RECT& rc, COLORREF color, std::wstring_view sortText) noexcept
{
    if (! resources.target || sortText.empty())
    {
        return;
    }

    const wchar_t icon    = sortText.back();
    const bool hasArrow   = sortText.size() >= 2;
    const wchar_t arrowCh = hasArrow ? sortText.front() : L'\0';

    RECT indicatorRect = rc;
    const int width    = std::max(0L, indicatorRect.right - indicatorRect.left);
    if (width <= 0)
    {
        return;
    }

    const int indicatorPx = std::max(1, MulDiv(kStatusBarSortMinPartWidthDip, static_cast<int>(resources.dpi), USER_DEFAULT_SCREEN_DPI));
    indicatorRect.left    = std::max(indicatorRect.left, indicatorRect.right - std::min(width, indicatorPx));

    RECT iconRect = indicatorRect;
    if (hasArrow && arrowCh != 0 && resources.arrowFormat)
    {
        RECT arrowRect       = indicatorRect;
        const LONG widthPx   = static_cast<LONG>(width);
        const LONG arrowArea = std::clamp(static_cast<LONG>(MulDiv(12, static_cast<int>(resources.dpi), USER_DEFAULT_SCREEN_DPI)), 0L, widthPx);
        const LONG split     = indicatorRect.left + arrowArea;
        const LONG gap       = std::max(1L, static_cast<LONG>(MulDiv(1, static_cast<int>(resources.dpi), USER_DEFAULT_SCREEN_DPI)));

        arrowRect.right = std::max(arrowRect.left, split - gap);
        iconRect.left   = std::min(iconRect.right, split + gap);
        DrawStatusBarText(resources, std::wstring_view{&arrowCh, 1u}, arrowRect, resources.arrowFormat.get(), color);
    }

    IDWriteTextFormat* iconFormat = (hasArrow && arrowCh != 0) ? resources.iconLeadingFormat.get() : resources.iconCenteredFormat.get();
    DrawStatusBarText(resources, std::wstring_view{&icon, 1u}, iconRect, iconFormat, color);
}

[[nodiscard]] bool IsStatusBarActivePane(HWND statusBar, const FolderWindow& owner) noexcept
{
    const int id = GetDlgCtrlID(statusBar);
    if (id == static_cast<int>(kLeftStatusBarId))
    {
        return owner.GetActivePane() == FolderWindow::Pane::Left;
    }
    if (id == static_cast<int>(kRightStatusBarId))
    {
        return owner.GetActivePane() == FolderWindow::Pane::Right;
    }
    return false;
}

[[nodiscard]] int GetStatusBarFocusLineHeightPx(int dpi, const RECT& clientRect) noexcept
{
    const int clientHeight = std::max(0l, clientRect.bottom - clientRect.top);
    if (clientHeight <= 0)
    {
        return 0;
    }

    const int desired = MulDiv(kStatusBarFocusLineHeightDip, dpi, USER_DEFAULT_SCREEN_DPI);
    return std::clamp(desired, 1, clientHeight);
}

[[nodiscard]] COLORREF StatusBarFocusLineColor(const AppTheme& theme, bool activePane, const uint32_t* hueDegrees) noexcept
{
    if (! activePane)
    {
        return theme.menu.separator;
    }

    if (! theme.menu.rainbowMode)
    {
        return theme.menu.selectionBg;
    }

    const uint32_t hueDegreesValue = hueDegrees ? *hueDegrees : 0u;
    const float hue                = static_cast<float>(hueDegreesValue % 360u);
    const float saturation         = 0.85f;
    const float value              = theme.menu.darkBase ? 0.80f : 0.90f;
    return ColorToCOLORREF(ColorFromHSV(hue, saturation, value, 1.0f));
}

[[nodiscard]] COLORREF StatusBarTextColor(const AppTheme& theme, bool activePane) noexcept
{
    if (activePane || theme.highContrast)
    {
        return theme.menu.text;
    }

    return BlendColor(theme.menu.background, theme.menu.text, 9, 20);
}

[[nodiscard]] bool PaintStatusBarDirectWrite(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    auto* owner = reinterpret_cast<FolderWindow*>(GetPropW(hwnd, kStatusBarOwnerProp));
    if (! owner)
    {
        return false;
    }

    const auto* selectionText = reinterpret_cast<const std::wstring*>(GetPropW(hwnd, kStatusBarSelectionTextProp));
    const auto* securityText  = reinterpret_cast<const std::wstring*>(GetPropW(hwnd, kStatusBarSecurityTextProp));
    const auto* sortText      = reinterpret_cast<const std::wstring*>(GetPropW(hwnd, kStatusBarSortTextProp));
    if (! selectionText || ! securityText || ! sortText)
    {
        return false;
    }

    StatusBarParts parts{};
    RECT client{};
    if (! ComputeStatusBarParts(hwnd, parts, &client))
    {
        return false;
    }

    const AppTheme& theme       = owner->GetTheme();
    const bool activePane       = IsStatusBarActivePane(hwnd, *owner);
    const COLORREF textColor    = StatusBarTextColor(theme, activePane);
    const auto* focusHueDegrees = reinterpret_cast<const uint32_t*>(GetPropW(hwnd, kStatusBarFocusHueProp));

    const RECT& part0 = parts.selection;
    const RECT& part1 = parts.security;
    const RECT& part2 = parts.sort;
    const bool has0   = part0.right > part0.left;
    const bool has1   = part1.right > part1.left;
    const bool has2   = part2.right > part2.left;

    auto* resources = EnsureStatusBarRenderResources(hwnd);
    if (! resources || ! EnsureStatusBarTextFormats(hwnd, *resources) || ! EnsureStatusBarTarget(hwnd, *resources))
    {
        return false;
    }

    Debug::Perf::Scope paintPerf(L"render.status_bar.paint_us");
    paintPerf.SetValue0(static_cast<uint64_t>(std::max(0L, client.right - client.left)));
    paintPerf.SetValue1(static_cast<uint64_t>(GetDlgCtrlID(hwnd)));

    resources->target->BeginDraw();
    auto endDraw = wil::scope_exit([&]() noexcept
    {
        const HRESULT hr = resources->target->EndDraw();
        if (hr == D2DERR_RECREATE_TARGET)
        {
            ResetStatusBarTarget(*resources);
        }
    });

    resources->target->Clear(D2DColor(theme.menu.background));

    const bool hot = GetPropW(hwnd, kStatusBarSortHotProp) != nullptr;
    if (hot && has2)
    {
        const COLORREF hotBg = BlendColor(theme.menu.background, theme.menu.selectionBg, 1, 2);
        FillStatusBarRect(*resources, part2, hotBg);

        RECT frame              = part2;
        const int focusLineSize = GetStatusBarFocusLineHeightPx(static_cast<int>(resources->dpi), client);
        frame.top               = std::min(frame.bottom, frame.top + focusLineSize);
        frame.left              = std::min(frame.right, frame.left + 1);

        RECT leftEdge{frame.left, frame.top, std::min(frame.right, frame.left + 1), frame.bottom};
        RECT rightEdge{std::max(frame.left, frame.right - 1), frame.top, frame.right, frame.bottom};
        RECT bottomEdge{frame.left, std::max(frame.top, frame.bottom - 1), frame.right, frame.bottom};
        FillStatusBarRect(*resources, leftEdge, theme.menu.separator);
        FillStatusBarRect(*resources, rightEdge, theme.menu.separator);
        FillStatusBarRect(*resources, bottomEdge, theme.menu.separator);
    }

    const int focusLinePx = GetStatusBarFocusLineHeightPx(static_cast<int>(resources->dpi), client);
    const RECT topLine    = {client.left, client.top, client.right, std::min(client.bottom, client.top + focusLinePx)};
    if (topLine.bottom > topLine.top)
    {
        const COLORREF lineColor = StatusBarFocusLineColor(theme, activePane, focusHueDegrees);
        FillStatusBarRect(*resources, topLine, lineColor);
    }

    const bool showSecurity = ! securityText->empty();
    if (showSecurity && has1)
    {
        RECT securityRect = part1;
        securityRect.top  = std::min(securityRect.bottom, securityRect.top + focusLinePx);

        if (securityRect.bottom > securityRect.top)
        {
            const COLORREF securityBg = BlendColor(theme.menu.background, theme.menu.selectionBg, 1, 5);
            FillStatusBarRect(*resources, securityRect, securityBg);
        }
    }

    if (has0)
    {
        RECT sepRect  = part0;
        sepRect.left  = std::max(part0.left, part0.right - 1);
        sepRect.right = part0.right;
        sepRect.top   = std::min(part0.bottom, part0.top + focusLinePx);
        if (sepRect.right > sepRect.left && sepRect.bottom > sepRect.top)
        {
            FillStatusBarRect(*resources, sepRect, theme.menu.separator);
        }
    }
    if (has1)
    {
        RECT sepRect  = part1;
        sepRect.left  = std::max(part1.left, part1.right - 1);
        sepRect.right = part1.right;
        sepRect.top   = std::min(part1.bottom, part1.top + focusLinePx);
        if (sepRect.right > sepRect.left && sepRect.bottom > sepRect.top)
        {
            FillStatusBarRect(*resources, sepRect, theme.menu.separator);
        }
    }

    const int paddingX     = std::max(1, MulDiv(kStatusBarPaddingXDip, static_cast<int>(resources->dpi), USER_DEFAULT_SCREEN_DPI));
    const int sortPaddingX = std::max(1, MulDiv(kStatusBarSortPaddingXDip, static_cast<int>(resources->dpi), USER_DEFAULT_SCREEN_DPI));

    RECT rc0  = has0 ? part0 : client;
    rc0.left  = std::min(rc0.right, rc0.left + paddingX);
    rc0.right = std::max(rc0.left, rc0.right - paddingX);
    rc0.top   = std::min(rc0.bottom, rc0.top + focusLinePx);

    RECT rc1  = has1 ? part1 : client;
    rc1.left  = std::min(rc1.right, rc1.left + sortPaddingX);
    rc1.right = std::max(rc1.left, rc1.right - sortPaddingX);
    rc1.top   = std::min(rc1.bottom, rc1.top + focusLinePx);

    RECT rc2  = has2 ? part2 : client;
    rc2.left  = std::min(rc2.right, rc2.left + sortPaddingX);
    rc2.right = std::max(rc2.left, rc2.right - sortPaddingX);
    rc2.top   = std::min(rc2.bottom, rc2.top + focusLinePx);

    DrawStatusBarText(*resources, *selectionText, rc0, resources->selectionFormat.get(), textColor);

    if (showSecurity)
    {
        DrawStatusBarText(*resources, *securityText, rc1, resources->securityFormat.get(), textColor);
    }

    const bool sortUsesIcons =
        resources->fluentIconAvailable && resources->iconLeadingFormat && resources->iconCenteredFormat && ContainsPrivateUseAreaGlyph(*sortText);
    const COLORREF sortColor = hot ? theme.menu.selectionText : textColor;
    if (sortUsesIcons)
    {
        PaintSortIndicatorGlyph(*resources, rc2, sortColor, *sortText);
    }
    else
    {
        DrawStatusBarText(*resources, *sortText, rc2, resources->sortTextFormat.get(), sortColor);
    }

    return true;
}

bool StatusBarCanCustomPaint(HWND hwnd) noexcept
{
    return GetPropW(hwnd, kStatusBarOwnerProp) != nullptr && GetPropW(hwnd, kStatusBarSelectionTextProp) != nullptr &&
           GetPropW(hwnd, kStatusBarSecurityTextProp) != nullptr && GetPropW(hwnd, kStatusBarSortTextProp) != nullptr;
}

[[nodiscard]] bool IsPointInStatusBarPart(HWND statusBar, int part, POINT clientPt) noexcept
{
    RECT rc{};
    if (! GetStatusBarPartRect(statusBar, part, rc))
    {
        return false;
    }

    return PtInRect(&rc, clientPt) != 0;
}

void UpdateStatusBarSortHot(HWND hwnd, POINT pt) noexcept
{
    const bool hotNow = IsPointInStatusBarPart(hwnd, kStatusBarPartSort, pt);
    const bool hotWas = GetPropW(hwnd, kStatusBarSortHotProp) != nullptr;
    if (hotNow == hotWas)
    {
        return;
    }

    if (hotNow)
    {
        SetPropW(hwnd, kStatusBarSortHotProp, reinterpret_cast<HANDLE>(1));
    }
    else
    {
        RemovePropW(hwnd, kStatusBarSortHotProp);
    }
    InvalidateRect(hwnd, nullptr, FALSE);
}

void TrackStatusBarMouseLeave(HWND hwnd) noexcept
{
    TRACKMOUSEEVENT tme{};
    tme.cbSize    = sizeof(tme);
    tme.dwFlags   = TME_LEAVE;
    tme.hwndTrack = hwnd;
    TrackMouseEvent(&tme);
}

LRESULT StatusBarOnEraseBkgnd(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    if (StatusBarCanCustomPaint(hwnd))
    {
        return 1;
    }
    return DefWindowProcW(hwnd, WM_ERASEBKGND, wParam, lParam);
}

LRESULT StatusBarOnPaint(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    if (! StatusBarCanCustomPaint(hwnd))
    {
        return DefWindowProcW(hwnd, WM_PAINT, wParam, lParam);
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint paintDc = wil::BeginPaint(hwnd, &ps);
    static_cast<void>(paintDc);
    static_cast<void>(PaintStatusBarDirectWrite(hwnd));
    return 0;
}

LRESULT StatusBarOnSetCursor(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    const LPARAM messagePos = static_cast<LPARAM>(GetMessagePos());
    POINT screenPt{GET_X_LPARAM(messagePos), GET_Y_LPARAM(messagePos)};
    if (ScreenToClient(hwnd, &screenPt) != FALSE)
    {
        if (IsPointInStatusBarPart(hwnd, kStatusBarPartSort, screenPt))
        {
            SetCursor(LoadCursorW(nullptr, IDC_HAND));
            return TRUE;
        }
    }
    return DefWindowProcW(hwnd, WM_SETCURSOR, wParam, lParam);
}

LRESULT StatusBarOnMouseMove(HWND hwnd, [[maybe_unused]] WPARAM wParam, LPARAM lParam) noexcept
{
    POINT pt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    UpdateStatusBarSortHot(hwnd, pt);
    TrackStatusBarMouseLeave(hwnd);
    return 0;
}

LRESULT StatusBarOnMouseLeave(HWND hwnd, [[maybe_unused]] WPARAM wParam, [[maybe_unused]] LPARAM lParam) noexcept
{
    if (RemovePropW(hwnd, kStatusBarSortHotProp))
    {
        InvalidateRect(hwnd, nullptr, FALSE);
    }
    return 0;
}

LRESULT StatusBarOnSize(HWND hwnd, [[maybe_unused]] WPARAM wParam, [[maybe_unused]] LPARAM lParam) noexcept
{
    if (StatusBarCanCustomPaint(hwnd))
    {
        InvalidateStatusBarDeviceResources(hwnd, false);
        UpdateStatusBarParts(hwnd);
    }

    InvalidateRect(hwnd, nullptr, FALSE);
    return 0;
}

LRESULT StatusBarOnDpiChangedAfterParent(HWND hwnd, [[maybe_unused]] WPARAM wParam, [[maybe_unused]] LPARAM lParam) noexcept
{
    if (StatusBarCanCustomPaint(hwnd))
    {
        InvalidateStatusBarDeviceResources(hwnd, true);
        UpdateStatusBarParts(hwnd);
    }

    InvalidateRect(hwnd, nullptr, FALSE);
    return 0;
}

LRESULT StatusBarOnLButtonUp(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    if (! StatusBarCanCustomPaint(hwnd))
    {
        return DefWindowProcW(hwnd, WM_LBUTTONUP, wParam, lParam);
    }

    const POINT clientPt{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
    if (! IsPointInStatusBarPart(hwnd, kStatusBarPartSort, clientPt))
    {
        return 0;
    }

    StatusBarSortClickNotification mouse{};
    mouse.hdr.hwndFrom = hwnd;
    mouse.hdr.idFrom   = static_cast<UINT_PTR>(GetDlgCtrlID(hwnd));
    mouse.hdr.code     = kStatusBarSortClickNotification;
    mouse.part         = kStatusBarPartSort;
    mouse.clientPoint  = clientPt;

    if (const HWND parent = GetParent(hwnd))
    {
        SendMessageW(parent, WM_NOTIFY, mouse.hdr.idFrom, reinterpret_cast<LPARAM>(&mouse));
    }

    return 0;
}

LRESULT StatusBarOnNcDestroy(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept
{
    RemovePropW(hwnd, kStatusBarSortHotProp);
    RemovePropW(hwnd, kStatusBarOwnerProp);
    RemovePropW(hwnd, kStatusBarSelectionTextProp);
    RemovePropW(hwnd, kStatusBarSecurityTextProp);
    RemovePropW(hwnd, kStatusBarSortTextProp);
    RemovePropW(hwnd, kStatusBarFocusHueProp);
    DestroyStatusBarRenderResources(hwnd);
    return DefWindowProcW(hwnd, WM_NCDESTROY, wParam, lParam);
}
} // namespace

HRESULT EnsureFolderWindowStatusBarClass(HINSTANCE instance) noexcept
{
    static ATOM atom = 0;
    if (atom != 0)
    {
        return S_OK;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = StatusBarWndProc;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kFolderWindowStatusBarClassName;
    wc.style         = CS_HREDRAW | CS_VREDRAW;

    atom = RegisterClassExW(&wc);
    if (atom != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS)
    {
        return S_OK;
    }

    return HRESULT_FROM_WIN32(GetLastError());
}

bool GetStatusBarPartRect(HWND hwnd, int part, RECT& rect) noexcept
{
    StatusBarParts parts{};
    if (! ComputeStatusBarParts(hwnd, parts))
    {
        return false;
    }

    switch (part)
    {
        case kStatusBarPartSelection: rect = parts.selection; break;
        case kStatusBarPartSecurity: rect = parts.security; break;
        case kStatusBarPartSort: rect = parts.sort; break;
        default: return false;
    }

    return rect.right > rect.left;
}

LRESULT CALLBACK StatusBarWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
        case WM_ERASEBKGND: return StatusBarOnEraseBkgnd(hwnd, wParam, lParam);
        case WM_PAINT: return StatusBarOnPaint(hwnd, wParam, lParam);
        case WM_DPICHANGED_AFTERPARENT: return StatusBarOnDpiChangedAfterParent(hwnd, wParam, lParam);
        case WM_SETCURSOR: return StatusBarOnSetCursor(hwnd, wParam, lParam);
        case WM_MOUSEMOVE: return StatusBarOnMouseMove(hwnd, wParam, lParam);
        case WM_MOUSELEAVE: return StatusBarOnMouseLeave(hwnd, wParam, lParam);
        case WM_SIZE: return StatusBarOnSize(hwnd, wParam, lParam);
        case WM_LBUTTONUP: return StatusBarOnLButtonUp(hwnd, wParam, lParam);
        case WM_NCDESTROY: return StatusBarOnNcDestroy(hwnd, wParam, lParam);
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

namespace
{

std::wstring FormatLocalTime(int64_t fileTime)
{
    if (fileTime <= 0)
    {
        return {};
    }

    ULARGE_INTEGER uli{};
    uli.QuadPart = static_cast<ULONGLONG>(fileTime);

    FILETIME ft{};
    ft.dwLowDateTime  = uli.LowPart;
    ft.dwHighDateTime = uli.HighPart;

    FILETIME local{};
    SYSTEMTIME st{};
    if (! FileTimeToLocalFileTime(&ft, &local) || ! FileTimeToSystemTime(&local, &st))
    {
        return {};
    }

    return std::format(L"{:04d}-{:02d}-{:02d} {:02d}:{:02d}", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute);
}

std::wstring FormatFileAttributes(DWORD attrs)
{
    std::wstring result;
    result.reserve(10);

    auto add = [&](DWORD flag, wchar_t ch)
    {
        if ((attrs & flag) != 0)
        {
            result.push_back(ch);
        }
    };

    add(FILE_ATTRIBUTE_READONLY, L'R');
    add(FILE_ATTRIBUTE_HIDDEN, L'H');
    add(FILE_ATTRIBUTE_SYSTEM, L'S');
    add(FILE_ATTRIBUTE_ARCHIVE, L'A');
    add(FILE_ATTRIBUTE_COMPRESSED, L'C');
    add(FILE_ATTRIBUTE_ENCRYPTED, L'E');
    add(FILE_ATTRIBUTE_TEMPORARY, L'T');
    add(FILE_ATTRIBUTE_OFFLINE, L'O');
    add(FILE_ATTRIBUTE_REPARSE_POINT, L'P');

    if (result.empty())
    {
        result = L"-";
    }

    return result;
}

std::wstring BuildSingleItemSummaryText(const FolderView::SelectionStats::SelectedItemDetails& details, std::wstring_view directorySizeText)
{
    const std::wstring timeText  = FormatLocalTime(details.lastWriteTime);
    const std::wstring attrsText = FormatFileAttributes(details.fileAttributes);

    if (details.isDirectory)
    {
        std::wstring sizeText{directorySizeText};
        if (sizeText.empty())
        {
            sizeText = LoadEmbeddedStringResource(nullptr, IDS_STATUS_SIZE_UNKNOWN);
        }

        if (! timeText.empty())
        {
            return FormatEmbeddedStringResource(nullptr, IDS_FMT_STATUS_SELECTED_SINGLE_DIR_TIME_ATTRS, sizeText, timeText, attrsText);
        }
        return FormatEmbeddedStringResource(nullptr, IDS_FMT_STATUS_SELECTED_SINGLE_DIR_ATTRS, sizeText, attrsText);
    }

    const std::wstring sizeText = FormatBytesCompact(details.sizeBytes);
    if (! timeText.empty())
    {
        return FormatEmbeddedStringResource(nullptr, IDS_FMT_STATUS_SELECTED_SINGLE_FILE_SIZE_TIME_ATTRS, sizeText, timeText, attrsText);
    }
    return FormatEmbeddedStringResource(nullptr, IDS_FMT_STATUS_SELECTED_SINGLE_FILE_SIZE_ATTRS, sizeText, attrsText);
}

std::wstring BuildSelectionSummaryText(const FolderView::SelectionStats& stats,
                                       std::wstring_view selectionSizeText,
                                       const std::optional<FolderView::SelectionStats::SelectedItemDetails>& focusedItemDetails)
{
    if (stats.selectedFiles == 0 && stats.selectedFolders == 0)
    {
        if (focusedItemDetails.has_value())
        {
            return BuildSingleItemSummaryText(focusedItemDetails.value(), {});
        }

        return LoadStringResource(nullptr, IDS_STATUS_NO_SELECTION);
    }

    if (stats.singleItem.has_value())
    {
        return BuildSingleItemSummaryText(stats.singleItem.value(), selectionSizeText);
    }

    const std::wstring_view folderSuffix = stats.selectedFolders == 1 ? std::wstring_view{} : std::wstring_view{L"s"};
    const std::wstring_view fileSuffix   = stats.selectedFiles == 1 ? std::wstring_view{} : std::wstring_view{L"s"};

    if (stats.selectedFiles > 0 && stats.selectedFolders > 0)
    {
        return FormatStringResource(
            nullptr, IDS_FMT_STATUS_SELECTED_FOLDERS_FILES, stats.selectedFolders, folderSuffix, stats.selectedFiles, fileSuffix, selectionSizeText);
    }

    if (stats.selectedFiles > 0)
    {
        return FormatStringResource(nullptr, IDS_FMT_STATUS_SELECTED_FILES, stats.selectedFiles, fileSuffix, selectionSizeText);
    }

    return FormatStringResource(nullptr, IDS_FMT_STATUS_SELECTED_FOLDERS, stats.selectedFolders, folderSuffix, selectionSizeText);
}

std::wstring BuildSortIndicatorText(FolderView::SortBy sortBy, FolderView::SortDirection direction, bool useFluentIcons)
{
    if (sortBy == FolderView::SortBy::None)
    {
        if (useFluentIcons)
        {
            return std::wstring(1, FluentIcons::kSort);
        }

        std::wstring placeholder = LoadEmbeddedStringResource(nullptr, IDS_STATUS_SORT_INDICATOR);
        if (placeholder.empty())
        {
            placeholder.assign(1, FluentIcons::kFallbackSort);
        }
        return placeholder;
    }

    // Asc/Desc should use arrows (not chevrons). The status bar paint draws the arrow + glyph.
    const wchar_t arrow = direction == FolderView::SortDirection::Ascending ? L'\u2191' : L'\u2193';

    const wchar_t icon = [&]() noexcept -> wchar_t
    {
        if (useFluentIcons)
        {
            switch (sortBy)
            {
                case FolderView::SortBy::Name: return FluentIcons::kFont;
                case FolderView::SortBy::Extension: return FluentIcons::kDocument;
                case FolderView::SortBy::Time: return FluentIcons::kCalendar;
                case FolderView::SortBy::Size: return FluentIcons::kHardDrive;
                case FolderView::SortBy::Attributes: return FluentIcons::kTag;
                case FolderView::SortBy::None: break;
            }
        }
        else
        {
            switch (sortBy)
            {
                case FolderView::SortBy::Name: return L'\u2263';
                case FolderView::SortBy::Extension: return L'\u24D4';
                case FolderView::SortBy::Time: return L'\u23F1';
                case FolderView::SortBy::Size: return direction == FolderView::SortDirection::Ascending ? L'\u25F0' : L'\u25F2';
                case FolderView::SortBy::Attributes: return L'\u24B6';
                case FolderView::SortBy::None: break;
            }
        }

        return 0;
    }();

    std::wstring result;
    result.reserve(icon != 0 ? 2u : 1u);
    result.push_back(arrow);
    if (icon != 0)
    {
        result.push_back(icon);
    }
    return result;
}
} // namespace

void FolderWindow::UpdatePaneStatusBar(Pane pane)
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.hStatusBar)
    {
        return;
    }

    const RECT& rect = pane == Pane::Left ? _leftStatusBarRect : _rightStatusBarRect;
    const int width  = std::max(0L, rect.right - rect.left);
    const int height = std::max(0L, rect.bottom - rect.top);

    const bool visible = state.statusBarVisible && width > 0 && height > 0;
    ShowWindow(state.hStatusBar.get(), visible ? SW_SHOWNA : SW_HIDE);
    if (! visible)
    {
        return;
    }

    std::wstring selectionSizeText;
    if (state.selectionStats.selectedFiles > 0 || state.selectionStats.selectedFolders > 0)
    {
        if (state.selectionStats.selectedFolders > 0)
        {
            if (state.selectionFolderBytesPending)
            {
                uint64_t totalBytesSoFar = state.selectionStats.selectedFileBytes;
                totalBytesSoFar += state.selectionFolderBytes;
                const std::wstring sizeText = FormatBytesCompact(totalBytesSoFar);
                selectionSizeText           = FormatStringResource(nullptr, IDS_FMT_STATUS_CALCULATING_SIZE_WITH_BYTES, sizeText);
                if (selectionSizeText.empty())
                {
                    selectionSizeText = LoadStringResource(nullptr, IDS_STATUS_CALCULATING_SIZE);
                }
            }
            else if (! state.selectionFolderBytesValid)
            {
                selectionSizeText = LoadEmbeddedStringResource(nullptr, IDS_STATUS_SIZE_UNKNOWN);
            }
            else
            {
                uint64_t totalBytes = state.selectionStats.selectedFileBytes;
                totalBytes += state.selectionFolderBytes;
                selectionSizeText = FormatBytesCompact(totalBytes);
            }
        }
        else
        {
            selectionSizeText = FormatBytesCompact(state.selectionStats.selectedFileBytes);
        }
    }

    if (state.folderView.IsIncrementalSearchActive())
    {
        const std::wstring queryText = std::wstring{state.folderView.GetIncrementalSearchQuery()};
        state.statusSelectionText    = FormatStringResource(nullptr, IDS_FMT_STATUS_INCREMENTAL_SEARCH, queryText);
        if (state.statusSelectionText.empty())
        {
            state.statusSelectionText = queryText;
        }
    }
    else
    {
        std::optional<FolderView::SelectionStats::SelectedItemDetails> focusedItemDetails;
        if (state.selectionStats.selectedFiles == 0 && state.selectionStats.selectedFolders == 0)
        {
            focusedItemDetails = state.folderView.GetFocusedItemDetails();
        }
        state.statusSelectionText = BuildSelectionSummaryText(state.selectionStats, selectionSizeText, focusedItemDetails);
    }
    const bool useFluentIcons = StatusBarSupportsFluentIcons(state.hStatusBar.get());
    state.statusSortText      = BuildSortIndicatorText(state.folderView.GetSortBy(), state.folderView.GetSortDirection(), useFluentIcons);

    state.statusSecurityText.clear();
    if (_settings)
    {
        if (const auto connName = ConnectionProfileUtils::TryParseConnNameFromPluginPath(state.currentPath); connName.has_value())
        {
            if (const Common::Settings::ConnectionProfile* profile = ConnectionProfileUtils::FindConnectionProfileByName(_settings, *connName);
                profile && ConnectionProfileUtils::ConnectionProfileUsesInsecureTls(*profile))
            {
                state.statusSecurityText = LoadStringResource(nullptr, IDS_STATUS_INSECURE_TLS);
            }
        }
    }

    UpdateStatusBarParts(state.hStatusBar.get());
}

#ifdef ENABLE_TESTS
bool FolderWindow::DebugGetPaneStatusBarSnapshot(Pane pane, FolderWindowPaneStatusBarDebugSnapshot& out) const noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.hStatusBar || IsWindow(state.hStatusBar.get()) == FALSE)
    {
        return false;
    }

    out                              = {};
    out.visible                      = IsWindowVisible(state.hStatusBar.get()) != FALSE;
    out.usesDirectWriteTextRendering = false;
    out.hasNativeFont                = false;
    out.activePane                   = pane == _activePane;
    out.selectionTextDimmed          = ! out.activePane && ! _theme.highContrast;
    out.textSizeDip                  = kStatusBarTextSizeDip;
    out.selectionText                = state.statusSelectionText;
    out.securityText                 = state.statusSecurityText;
    out.sortText                     = state.statusSortText;

    wchar_t className[128]{};
    const int classLength = GetClassNameW(state.hStatusBar.get(), className, static_cast<int>(std::size(className)));
    if (classLength > 0)
    {
        out.className.assign(className, static_cast<size_t>(classLength));
        out.usesNativeStatusBarClass = _wcsicmp(out.className.c_str(), L"msctls_statusbar32") == 0;
    }

    if (auto* resources = EnsureStatusBarRenderResources(state.hStatusBar.get()))
    {
        out.usesDirectWriteTextRendering = EnsureStatusBarTextFormats(state.hStatusBar.get(), *resources) && resources->selectionFormat &&
                                           resources->securityFormat && resources->sortTextFormat && resources->arrowFormat;
    }

    return true;
}

bool FolderWindow::DebugClickPaneStatusBarSort(Pane pane) noexcept
{
    const PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.hStatusBar || IsWindow(state.hStatusBar.get()) == FALSE)
    {
        return false;
    }

    RECT sortRect{};
    if (! GetStatusBarPartRect(state.hStatusBar.get(), kStatusBarPartSort, sortRect))
    {
        return false;
    }

    const int x = std::max(sortRect.left, sortRect.right - 4);
    const int y = sortRect.top + ((sortRect.bottom - sortRect.top) / 2);

    PostMessageW(state.hStatusBar.get(), WM_MOUSEMOVE, 0, MAKELPARAM(x, y));
    PostMessageW(state.hStatusBar.get(), WM_LBUTTONUP, MK_LBUTTON, MAKELPARAM(x, y));
    return true;
}
#endif
