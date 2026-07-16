#include "SplashScreen.h"

#include "Framework.h"

#include "DxUi/DxUi.h"

#include <wincodec.h>

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <exception>
#include <mutex>
#include <stop_token>
#include <string>
#include <thread>

#pragma warning(push)
// WIL headers: deleted copy/move and unused inline Helpers
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"
#include "LocalizationManager.h"
#include "Version.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#include "resource.h"

namespace SplashScreen
{
namespace
{
// Constexpr RGB helper to avoid macro cast warnings with constexpr
constexpr COLORREF MakeRGB(BYTE r, BYTE g, BYTE b) noexcept
{
    return static_cast<COLORREF>(r) | (static_cast<COLORREF>(g) << 8) | (static_cast<COLORREF>(b) << 16);
}

constexpr BYTE ColorRefR(COLORREF value) noexcept
{
    return static_cast<BYTE>(value & 0xFFu);
}

constexpr BYTE ColorRefG(COLORREF value) noexcept
{
    return static_cast<BYTE>((value >> 8) & 0xFFu);
}

constexpr BYTE ColorRefB(COLORREF value) noexcept
{
    return static_cast<BYTE>((value >> 16) & 0xFFu);
}

std::atomic<bool> g_threadStarted{false};
wil::unique_event_nothrow g_closeEvent;
std::atomic<HWND> g_hwnd{nullptr};
std::atomic<HWND> g_owner{nullptr};

#ifdef ENABLE_TESTS
std::atomic<unsigned long> g_debugStage{0};
std::atomic<unsigned long> g_debugLastError{0};
std::atomic<long> g_debugComHr{0};
#endif

std::mutex g_textMutex;
std::wstring g_statusText;

// g_workerThread must be declared last: it is destroyed first at process teardown, so
// ~jthread joins the worker before the globals above (mutex, status text, close event) die.
std::jthread g_workerThread;

constexpr int kSplashLogoDesignDip         = 162;
constexpr int kSplashContentOffsetDip      = 14;
constexpr wchar_t kSplashWindowClassName[] = L"RedSalamander.SplashWindow";

// Option 6: Moonlight Metal
constexpr COLORREF kBgStart          = MakeRGB(15, 20, 27);
constexpr COLORREF kBgEnd            = MakeRGB(34, 42, 51);
constexpr COLORREF kFallbackRing     = MakeRGB(255, 206, 130);
constexpr COLORREF kFallbackFrame    = MakeRGB(98, 108, 128);
constexpr COLORREF kSeparator        = MakeRGB(182, 123, 50);
constexpr COLORREF kBorder           = MakeRGB(74, 89, 104);
constexpr COLORREF kTitleText        = MakeRGB(243, 247, 255);
constexpr COLORREF kStatusText       = MakeRGB(207, 177, 137);
constexpr COLORREF kSecondaryText    = MakeRGB(182, 150, 108);
constexpr COLORREF kLogoFallbackText = MakeRGB(251, 243, 232);

struct SplashWindowState
{
    SplashWindowState()                                    = default;
    SplashWindowState(const SplashWindowState&)            = delete;
    SplashWindowState& operator=(const SplashWindowState&) = delete;
    SplashWindowState(SplashWindowState&&)                 = delete;
    SplashWindowState& operator=(SplashWindowState&&)      = delete;

    RedSalamander::DxUi::WindowHost dxHost;
    class SplashRootControl* root              = nullptr;
    RedSalamander::DxUi::Label* titleLabel     = nullptr;
    RedSalamander::DxUi::Label* versionLabel   = nullptr;
    RedSalamander::DxUi::Label* copyrightLabel = nullptr;
    RedSalamander::DxUi::Label* statusLabel    = nullptr;
    wil::com_ptr<IWICImagingFactory> wicFactory;
    wil::com_ptr<ID2D1Bitmap1> splashLogoBitmap;
    wil::com_ptr<ID2D1Device> splashLogoDevice;
    float splashLogoDpiX = 0.0f;
    float splashLogoDpiY = 0.0f;
};

LRESULT CALLBACK SplashWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) noexcept;
void StartSplashDrag(HWND hwnd) noexcept;
void CenterOverOwner(HWND hwnd, HWND owner) noexcept;
[[nodiscard]] std::wstring GetStatusText() noexcept;
[[nodiscard]] bool EnsureSplashLogoBitmap(SplashWindowState& state, RedSalamander::DxUi::WindowHost& host) noexcept;

[[nodiscard]] bool EnsureSplashWindowClassRegistered(HINSTANCE instance) noexcept
{
    WNDCLASSEXW existing{};
    if (GetClassInfoExW(instance, kSplashWindowClassName, &existing) != FALSE)
    {
        return true;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = SplashWindowProc;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kSplashWindowClassName;
    return RegisterClassExW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

class SplashRootControl final : public RedSalamander::DxUi::Panel
{
public:
    explicit SplashRootControl(SplashWindowState* state) noexcept : _state(state)
    {
    }

    void Paint(RedSalamander::DxUi::WindowHost& host) const override
    {
        if (auto* dc = host.GetDeviceContext())
        {
            const auto bounds = host.GetClientBoundsDip();
            const D2D1_ROUNDED_RECT backgroundRect{
                bounds,
                14.0f,
                14.0f,
            };

            const D2D1_GRADIENT_STOP gradientStops[] = {
                {0.0f, D2D1::ColorF(ColorRefR(kBgStart) / 255.0f, ColorRefG(kBgStart) / 255.0f, ColorRefB(kBgStart) / 255.0f, 1.0f)},
                {1.0f, D2D1::ColorF(ColorRefR(kBgEnd) / 255.0f, ColorRefG(kBgEnd) / 255.0f, ColorRefB(kBgEnd) / 255.0f, 1.0f)},
            };

            wil::com_ptr<ID2D1GradientStopCollection> stopCollection;
            if (SUCCEEDED(dc->CreateGradientStopCollection(gradientStops, static_cast<UINT32>(std::size(gradientStops)), stopCollection.put())))
            {
                wil::com_ptr<ID2D1LinearGradientBrush> backgroundBrush;
                const D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES brushProps{
                    D2D1::Point2F(bounds.left, bounds.top),
                    D2D1::Point2F(bounds.right, bounds.bottom),
                };
                if (SUCCEEDED(dc->CreateLinearGradientBrush(brushProps, stopCollection.get(), backgroundBrush.put())))
                {
                    dc->FillRoundedRectangle(backgroundRect, backgroundBrush.get());
                }
            }

            const float logoSize = static_cast<float>(kSplashLogoDesignDip);
            const float logoX    = -38.0f;
            const float logoY    = std::max(0.0f, (bounds.bottom - bounds.top - logoSize) / 2.0f);
            const auto separatorBrush =
                host.GetSolidBrush(D2D1::ColorF(ColorRefR(kSeparator) / 255.0f, ColorRefG(kSeparator) / 255.0f, ColorRefB(kSeparator) / 255.0f, 1.0f));
            const auto borderBrush =
                host.GetSolidBrush(D2D1::ColorF(ColorRefR(kBorder) / 255.0f, ColorRefG(kBorder) / 255.0f, ColorRefB(kBorder) / 255.0f, 1.0f));

            const bool hasBitmap = _state && EnsureSplashLogoBitmap(*_state, host) && _state->splashLogoBitmap;
            if (hasBitmap)
            {
                const D2D1_RECT_F logoRect = D2D1::RectF(logoX, logoY, logoX + logoSize, logoY + logoSize);
                dc->DrawBitmap(_state->splashLogoBitmap.get(), &logoRect, 1.0f, D2D1_INTERPOLATION_MODE_LINEAR);
            }
            else
            {
                const auto accentBrush = host.GetSolidBrush(
                    D2D1::ColorF(ColorRefR(kFallbackRing) / 255.0f, ColorRefG(kFallbackRing) / 255.0f, ColorRefB(kFallbackRing) / 255.0f, 1.0f));
                const auto frameBrush = host.GetSolidBrush(
                    D2D1::ColorF(ColorRefR(kFallbackFrame) / 255.0f, ColorRefG(kFallbackFrame) / 255.0f, ColorRefB(kFallbackFrame) / 255.0f, 1.0f));
                const auto fillBrush = host.GetSolidBrush(
                    D2D1::ColorF(ColorRefR(kLogoFallbackText) / 255.0f, ColorRefG(kLogoFallbackText) / 255.0f, ColorRefB(kLogoFallbackText) / 255.0f, 0.12f));
                if (fillBrush && frameBrush && accentBrush)
                {
                    const D2D1_ELLIPSE outer{
                        D2D1::Point2F(logoX + (logoSize * 0.52f), logoY + (logoSize * 0.5f)),
                        logoSize * 0.42f,
                        logoSize * 0.42f,
                    };
                    const D2D1_ELLIPSE inner{
                        D2D1::Point2F(logoX + (logoSize * 0.58f), logoY + (logoSize * 0.44f)),
                        logoSize * 0.19f,
                        logoSize * 0.19f,
                    };
                    dc->FillEllipse(outer, fillBrush);
                    dc->DrawEllipse(outer, frameBrush, 2.0f);
                    dc->DrawEllipse(inner, accentBrush, 3.0f);
                }
            }

            if (separatorBrush && borderBrush)
            {
                const float separatorX = logoX + logoSize + 8.0f + static_cast<float>(kSplashContentOffsetDip);
                dc->DrawLine(D2D1::Point2F(separatorX, 18.0f), D2D1::Point2F(separatorX, std::max(18.0f, bounds.bottom - 18.0f)), separatorBrush, 1.0f);

                dc->DrawRoundedRectangle(backgroundRect, borderBrush, 1.0f);
            }
        }

        Panel::Paint(host);
    }

    bool OnMouseDown(RedSalamander::DxUi::WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override
    {
        if (! rightButton)
        {
            static_cast<void>(point);
            static_cast<void>(modifiers);
            StartSplashDrag(host.GetHwnd());
            return true;
        }
        return Panel::OnMouseDown(host, point, rightButton, modifiers);
    }

private:
    SplashWindowState* _state = nullptr;
};

void UpdateSplashWindowSize(HWND hwnd) noexcept
{
    const UINT dpi     = GetDpiForWindow(hwnd);
    const int widthPx  = Common::WindowSizing::DipToPixelRounded(dpi, 560);
    const int heightPx = Common::WindowSizing::DipToPixelRounded(dpi, 220);
    SetWindowPos(hwnd, nullptr, 0, 0, widthPx, heightPx, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

    const int radius = Common::WindowSizing::DipToPixelRounded(dpi, 14);
    wil::unique_hrgn rgn(CreateRoundRectRgn(0, 0, widthPx + 1, heightPx + 1, radius, radius));
    if (rgn)
    {
        if (SetWindowRgn(hwnd, rgn.get(), TRUE) != 0)
        {
            rgn.release();
        }
    }
}

void LayoutSplashUi(SplashWindowState& state) noexcept
{
    const auto bounds = state.dxHost.GetClientBoundsDip();
    if (! state.root)
    {
        return;
    }

    state.root->SetBounds(bounds);

    const float paddingX     = 20.0f;
    const float paddingY     = 18.0f;
    const float gapX         = 16.0f;
    const float logoSize     = static_cast<float>(kSplashLogoDesignDip);
    const float textX        = std::max(paddingX, -38.0f + logoSize + gapX + static_cast<float>(kSplashContentOffsetDip) + 8.0f);
    const float textWidth    = std::max(1.0f, bounds.right - textX - paddingX);
    const float titleHeight  = 30.0f;
    const float metaHeight   = 18.0f;
    const float statusHeight = 18.0f;
    const float lineGap      = 6.0f;

    float y = paddingY;
    if (state.titleLabel)
    {
        state.titleLabel->SetBounds(D2D1::RectF(textX, y, textX + textWidth, y + titleHeight));
    }
    y += titleHeight + lineGap;

    if (state.versionLabel)
    {
        state.versionLabel->SetBounds(D2D1::RectF(textX, y, textX + textWidth, y + metaHeight));
    }
    y += metaHeight + 2.0f;

    if (state.copyrightLabel)
    {
        state.copyrightLabel->SetBounds(D2D1::RectF(textX, y, textX + textWidth, y + metaHeight));
    }

    float statusY = bounds.bottom - paddingY - statusHeight;
    statusY       = std::max(statusY, y + lineGap);
    if (state.statusLabel)
    {
        state.statusLabel->SetBounds(D2D1::RectF(textX, statusY, textX + textWidth, statusY + statusHeight));
    }
}

void UpdateSplashLabels(SplashWindowState& state) noexcept
{
    const std::wstring appTitle = LoadEmbeddedStringResource(GetModuleHandleW(nullptr), IDS_APP_TITLE);
    if (state.titleLabel)
    {
        state.titleLabel->SetText(! appTitle.empty() ? appTitle : std::wstring(L"RedSalamander"));
    }
    if (state.versionLabel)
    {
        state.versionLabel->SetText(std::format(L"Version {}", VERSINFO_VERSION));
    }
    if (state.copyrightLabel)
    {
        state.copyrightLabel->SetText(VERSINFO_COPYRIGHT);
    }
    if (state.statusLabel)
    {
        state.statusLabel->SetText(GetStatusText());
    }
}

void BuildSplashUi(SplashWindowState& state) noexcept
{
    auto root  = std::make_unique<SplashRootControl>(&state);
    state.root = root.get();

    state.titleLabel = root->AddChild<RedSalamander::DxUi::Label>(L"");
    state.titleLabel->SetFontRole(RedSalamander::DxUi::FontRole::Header);
    state.titleLabel->SetTextColor(D2D1::ColorF(ColorRefR(kTitleText) / 255.0f, ColorRefG(kTitleText) / 255.0f, ColorRefB(kTitleText) / 255.0f, 1.0f));

    state.versionLabel = root->AddChild<RedSalamander::DxUi::Label>(L"");
    state.versionLabel->SetTextColor(
        D2D1::ColorF(ColorRefR(kSecondaryText) / 255.0f, ColorRefG(kSecondaryText) / 255.0f, ColorRefB(kSecondaryText) / 255.0f, 1.0f));

    state.copyrightLabel = root->AddChild<RedSalamander::DxUi::Label>(L"");
    state.copyrightLabel->SetTextColor(
        D2D1::ColorF(ColorRefR(kSecondaryText) / 255.0f, ColorRefG(kSecondaryText) / 255.0f, ColorRefB(kSecondaryText) / 255.0f, 1.0f));

    state.statusLabel = root->AddChild<RedSalamander::DxUi::Label>(L"");
    state.statusLabel->SetTextColor(D2D1::ColorF(ColorRefR(kStatusText) / 255.0f, ColorRefG(kStatusText) / 255.0f, ColorRefB(kStatusText) / 255.0f, 1.0f));

    state.dxHost.SetTheme(RedSalamander::DxUi::MakeDefaultThemePalette(true));
    state.dxHost.SetRoot(std::move(root));
    UpdateSplashLabels(state);
    LayoutSplashUi(state);
}

[[nodiscard]] bool EnsureSplashLogoBitmap(SplashWindowState& state, RedSalamander::DxUi::WindowHost& host) noexcept
{
    auto* dc = host.GetDeviceContext();
    if (! dc)
    {
        return false;
    }

    float dpiX = 96.0f;
    float dpiY = 96.0f;
    dc->GetDpi(&dpiX, &dpiY);

    wil::com_ptr<ID2D1Device> currentDevice;
    dc->GetDevice(currentDevice.put());

    const bool needsReload = ! state.splashLogoBitmap || ! state.splashLogoDevice || state.splashLogoDevice.get() != currentDevice.get() ||
                             std::fabs(state.splashLogoDpiX - dpiX) > 0.01f || std::fabs(state.splashLogoDpiY - dpiY) > 0.01f;
    if (! needsReload)
    {
        return true;
    }

    if (! state.wicFactory)
    {
        const HRESULT hrFactory = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(state.wicFactory.put()));
        if (FAILED(hrFactory) || ! state.wicFactory)
        {
            Debug::Warning(L"SplashScreen: Failed to create WIC factory: 0x{:08X}", static_cast<unsigned int>(hrFactory));
            return false;
        }
    }

    const HINSTANCE instance = GetModuleHandleW(nullptr);
    if (! instance)
    {
        return false;
    }

    const int targetPixels = std::max(1, static_cast<int>(std::lround((static_cast<float>(kSplashLogoDesignDip) * dpiX) / 96.0f)));
    wil::unique_hicon splashIcon(static_cast<HICON>(
        Localization::LoadImageResource(instance, MAKEINTRESOURCEW(IDI_SPLASH_LOGO_ICON), IMAGE_ICON, targetPixels, targetPixels, LR_DEFAULTCOLOR)));
    if (! splashIcon)
    {
        Debug::Warning(L"SplashScreen: Failed to load splash logo icon resource.");
        return false;
    }

    wil::com_ptr<IWICBitmap> wicBitmap;
    HRESULT hr = state.wicFactory->CreateBitmapFromHICON(splashIcon.get(), wicBitmap.put());
    if (FAILED(hr) || ! wicBitmap)
    {
        Debug::Warning(L"SplashScreen: Failed to create WIC bitmap from splash logo icon: 0x{:08X}", static_cast<unsigned int>(hr));
        return false;
    }

    wil::com_ptr<IWICFormatConverter> converter;
    hr = state.wicFactory->CreateFormatConverter(converter.put());
    if (FAILED(hr) || ! converter)
    {
        Debug::Warning(L"SplashScreen: Failed to create WIC format converter for splash logo: 0x{:08X}", static_cast<unsigned int>(hr));
        return false;
    }

    hr = converter->Initialize(wicBitmap.get(), GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, nullptr, 0.0f, WICBitmapPaletteTypeCustom);
    if (FAILED(hr))
    {
        Debug::Warning(L"SplashScreen: Failed to initialize WIC format converter for splash logo: 0x{:08X}", static_cast<unsigned int>(hr));
        return false;
    }

    D2D1_BITMAP_PROPERTIES1 bitmapProps{};
    bitmapProps.pixelFormat.format    = DXGI_FORMAT_B8G8R8A8_UNORM;
    bitmapProps.pixelFormat.alphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED;
    bitmapProps.dpiX                  = dpiX;
    bitmapProps.dpiY                  = dpiY;
    bitmapProps.bitmapOptions         = D2D1_BITMAP_OPTIONS_NONE;

    wil::com_ptr<ID2D1Bitmap1> bitmap;
    hr = dc->CreateBitmapFromWicBitmap(converter.get(), &bitmapProps, bitmap.put());
    if (FAILED(hr) || ! bitmap)
    {
        Debug::Warning(L"SplashScreen: Failed to create D2D bitmap for splash logo: 0x{:08X}", static_cast<unsigned int>(hr));
        return false;
    }

    state.splashLogoBitmap = std::move(bitmap);
    state.splashLogoDevice = std::move(currentDevice);
    state.splashLogoDpiX   = dpiX;
    state.splashLogoDpiY   = dpiY;
    return true;
}

LRESULT CALLBACK SplashWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) noexcept
{
    auto* state = reinterpret_cast<SplashWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    if (msg == WM_NCCREATE)
    {
        auto* create      = reinterpret_cast<CREATESTRUCTW*>(lParam);
        auto* createState = create ? reinterpret_cast<SplashWindowState*>(create->lpCreateParams) : nullptr;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(createState));
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    if (state)
    {
        bool handled         = false;
        const LRESULT result = state->dxHost.HandleMessage(hwnd, msg, wParam, lParam, handled);
        if (handled)
        {
            if (msg == WM_NCDESTROY)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                g_hwnd.store(nullptr, std::memory_order_release);
            }
            return result;
        }
    }

    switch (msg)
    {
        case WM_CREATE:
            if (! state || ! state->dxHost.Attach(hwnd))
            {
                return -1;
            }
            UpdateSplashWindowSize(hwnd);
            BuildSplashUi(*state);
            CenterOverOwner(hwnd, g_owner.load(std::memory_order_acquire));
            return 0;
        case WM_SIZE:
            if (state)
            {
                LayoutSplashUi(*state);
                state->dxHost.Invalidate();
            }
            return 0;
        case WM_DPICHANGED:
        {
            const auto* suggested = reinterpret_cast<const RECT*>(lParam);
            if (suggested)
            {
                SetWindowPos(hwnd,
                             nullptr,
                             suggested->left,
                             suggested->top,
                             suggested->right - suggested->left,
                             suggested->bottom - suggested->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            UpdateSplashWindowSize(hwnd);
            if (state)
            {
                LayoutSplashUi(*state);
                state->dxHost.Invalidate();
            }
            return 0;
        }
        case WM_ERASEBKGND: return 1;
        case WM_LBUTTONDOWN: StartSplashDrag(hwnd); return 0;
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_DESTROY:
            g_hwnd.store(nullptr, std::memory_order_release);
            PostQuitMessage(0);
            return 0;
        case WndMsg::kSplashScreenSetText:
            if (state)
            {
                UpdateSplashLabels(*state);
                state->dxHost.Invalidate();
            }
            return 0;
        case WndMsg::kSplashScreenRecenter: CenterOverOwner(hwnd, g_owner.load(std::memory_order_acquire)); return 0;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

void StartSplashDrag(HWND hwnd) noexcept
{
    const HWND dragTarget = GetAncestor(hwnd, GA_ROOT);
    if (! dragTarget)
    {
        return;
    }

    ReleaseCapture();
    static_cast<void>(SendMessageW(dragTarget, WM_NCLBUTTONDOWN, HTCAPTION, 0));
}

[[nodiscard]] RECT GetWorkAreaForOwner(HWND owner) noexcept
{
    HMONITOR monitor = nullptr;
    if (owner && IsWindow(owner))
    {
        monitor = MonitorFromWindow(owner, MONITOR_DEFAULTTONEAREST);
    }
    if (! monitor)
    {
        const HWND foreground = GetForegroundWindow();
        if (foreground && IsWindow(foreground))
        {
            monitor = MonitorFromWindow(foreground, MONITOR_DEFAULTTONEAREST);
        }
    }
    if (! monitor)
    {
        monitor = MonitorFromPoint(POINT{}, MONITOR_DEFAULTTOPRIMARY);
    }

    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (monitor && GetMonitorInfoW(monitor, &mi))
    {
        return mi.rcWork;
    }

    RECT fallback{};
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &fallback, 0);
    return fallback;
}

void CenterOverOwner(HWND hwnd, HWND owner) noexcept
{
    if (! hwnd)
    {
        return;
    }

    RECT rc{};
    if (! GetWindowRect(hwnd, &rc))
    {
        return;
    }

    const int width  = std::max(1, static_cast<int>(rc.right - rc.left));
    const int height = std::max(1, static_cast<int>(rc.bottom - rc.top));

    const RECT workArea = GetWorkAreaForOwner(owner);

    int targetCenterX = (workArea.left + workArea.right) / 2;
    int targetCenterY = (workArea.top + workArea.bottom) / 2;
    if (owner && IsWindow(owner))
    {
        RECT ownerRc{};
        if (GetWindowRect(owner, &ownerRc))
        {
            targetCenterX = (ownerRc.left + ownerRc.right) / 2;
            targetCenterY = (ownerRc.top + ownerRc.bottom) / 2;
        }
    }

    int left = targetCenterX - (width / 2);
    int top  = targetCenterY - (height / 2);

    const int maxLeft = std::max(static_cast<int>(workArea.left), static_cast<int>(workArea.right) - width);
    const int maxTop  = std::max(static_cast<int>(workArea.top), static_cast<int>(workArea.bottom) - height);
    left              = std::clamp(left, static_cast<int>(workArea.left), maxLeft);
    top               = std::clamp(top, static_cast<int>(workArea.top), maxTop);

    SetWindowPos(hwnd, HWND_TOPMOST, left, top, 0, 0, SWP_NOSIZE | SWP_NOACTIVATE);
}

[[nodiscard]] std::wstring GetStatusText() noexcept
{
    std::scoped_lock lock(g_textMutex);
    return g_statusText;
}

void SetStatusText(std::wstring_view text) noexcept
{
    std::scoped_lock lock(g_textMutex);
    g_statusText.assign(text);
}

void ThreadMain(std::stop_token stopToken, std::chrono::milliseconds delay, HINSTANCE instance) noexcept
{
#ifdef ENABLE_TESTS
    g_debugStage.store(1, std::memory_order_release);
    g_debugLastError.store(0, std::memory_order_release);
    g_debugComHr.store(S_OK, std::memory_order_release);
#endif
    const auto resetState = wil::scope_exit([]() noexcept
    {
        g_hwnd.store(nullptr, std::memory_order_release);
        // g_closeEvent is owned by the main thread (BeginDelayedOpen/CloseIfExist); resetting it
        // here would race a concurrent SetEvent from RequestCloseIfExist.
        g_threadStarted.store(false, std::memory_order_release);
    });

    if (! g_closeEvent)
    {
        return;
    }

    const DWORD delayMs = delay.count() < 0 ? 0u : static_cast<DWORD>(delay.count());
    const DWORD wait    = WaitForSingleObject(g_closeEvent.get(), delayMs);
    if (wait == WAIT_OBJECT_0)
    {
#ifdef ENABLE_TESTS
        g_debugStage.store(2, std::memory_order_release);
#endif
        return;
    }

    if (Detail::ShouldAbortPendingOpen(stopToken, g_closeEvent.get()))
    {
#ifdef ENABLE_TESTS
        g_debugStage.store(3, std::memory_order_release);
#endif
        return;
    }

    const HRESULT comHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
#ifdef ENABLE_TESTS
    g_debugComHr.store(static_cast<long>(comHr), std::memory_order_release);
    g_debugStage.store(4, std::memory_order_release);
#endif
    const auto comCleanup = wil::scope_exit([&]
    {
        RedSalamander::DxUi::ShutdownNativeTextInputForCurrentThread();
        if (SUCCEEDED(comHr))
        {
            CoUninitialize();
        }
    });

    if (! EnsureSplashWindowClassRegistered(instance))
    {
#ifdef ENABLE_TESTS
        g_debugLastError.store(GetLastError(), std::memory_order_release);
        g_debugStage.store(5, std::memory_order_release);
#endif
        return;
    }
#ifdef ENABLE_TESTS
    g_debugStage.store(6, std::memory_order_release);
#endif

    if (Detail::ShouldAbortPendingOpen(stopToken, g_closeEvent.get()))
    {
#ifdef ENABLE_TESTS
        g_debugStage.store(7, std::memory_order_release);
#endif
        return;
    }

    // Keep splash state ownership on the worker thread for the entire window lifetime.
    // The window proc only borrows the raw pointer through GWLP_USERDATA and clears it
    // during WM_NCDESTROY; that avoids delete-from-callback lifetime traps and keeps the
    // CreateWindowEx failure path single-owner too.
    auto splashState = std::make_unique<SplashWindowState>();
    const HWND owner = g_owner.load(std::memory_order_acquire);
    wil::unique_hwnd hwnd(CreateWindowExW(WS_EX_TOOLWINDOW,
                                          kSplashWindowClassName,
                                          L"",
                                          WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                          CW_USEDEFAULT,
                                          CW_USEDEFAULT,
                                          Common::WindowSizing::DipToPixelRounded(USER_DEFAULT_SCREEN_DPI, 560),
                                          Common::WindowSizing::DipToPixelRounded(USER_DEFAULT_SCREEN_DPI, 220),
                                          owner,
                                          nullptr,
                                          instance,
                                          splashState.get()));
    if (! hwnd)
    {
#ifdef ENABLE_TESTS
        g_debugLastError.store(GetLastError(), std::memory_order_release);
        g_debugStage.store(8, std::memory_order_release);
#endif
        return;
    }
#ifdef ENABLE_TESTS
    g_debugStage.store(9, std::memory_order_release);
#endif

    if (Detail::ShouldAbortPendingOpen(stopToken, g_closeEvent.get()))
    {
#ifdef ENABLE_TESTS
        g_debugStage.store(10, std::memory_order_release);
#endif
        hwnd.reset();
        return;
    }

    ShowWindow(hwnd.get(), SW_SHOWNOACTIVATE);
    SetWindowPos(hwnd.get(), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    UpdateWindow(hwnd.get());
#ifdef ENABLE_TESTS
    g_debugStage.store(11, std::memory_order_release);
#endif
    g_hwnd.store(hwnd.get(), std::memory_order_release);
    // Re-check after publishing the handle: a close requested between the abort check
    // above and the store would otherwise be lost (the closer only posts WM_CLOSE when it
    // sees a published hwnd) and GetMessageW below would block forever.
    if (Detail::ShouldAbortPendingOpen(stopToken, g_closeEvent.get()))
    {
#ifdef ENABLE_TESTS
        g_debugStage.store(12, std::memory_order_release);
#endif
        hwnd.reset();
        return;
    }

    MSG msg{};
    while (GetMessageW(&msg, nullptr, 0, 0) > 0)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    if (IsWindow(hwnd.get()) != FALSE)
    {
        hwnd.reset();
    }
    else
    {
        // WM_CLOSE/WM_DESTROY already owns the splash lifetime in the normal path.
        // Releasing here avoids a second DestroyWindow on a handle that has already been torn down.
        hwnd.release();
    }
}
} // namespace

void BeginDelayedOpen(std::chrono::milliseconds delay, HINSTANCE instance) noexcept
{
#ifdef ENABLE_TESTS
    g_debugStage.store(20, std::memory_order_release);
    g_debugLastError.store(0, std::memory_order_release);
#endif
    const bool alreadyStarted = g_threadStarted.exchange(true, std::memory_order_acq_rel);
    if (alreadyStarted)
    {
#ifdef ENABLE_TESTS
        g_debugStage.store(21, std::memory_order_release);
#endif
        return;
    }

    // A previous worker may have finished (clearing g_threadStarted) without being joined;
    // join it before replacing the close event it might still reference.
    if (g_workerThread.joinable())
    {
        g_workerThread.join();
    }

    g_closeEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! g_closeEvent)
    {
#ifdef ENABLE_TESTS
        g_debugLastError.store(GetLastError(), std::memory_order_release);
        g_debugStage.store(22, std::memory_order_release);
#endif
        g_threadStarted.store(false, std::memory_order_release);
        return;
    }

    // Mandatory: `noexcept` boundary. Splash is best-effort; thread creation can throw.
    try
    {
        g_workerThread = std::jthread([delay, instance](std::stop_token stopToken) noexcept
        {
            std::stop_callback stopCallback(stopToken,
                                            []() noexcept
            {
                if (g_closeEvent)
                {
                    static_cast<void>(SetEvent(g_closeEvent.get()));
                }

                const HWND hwnd = g_hwnd.load(std::memory_order_acquire);
                if (hwnd)
                {
                    static_cast<void>(PostMessageW(hwnd, WM_CLOSE, 0, 0));
                }
            });

            ThreadMain(stopToken, delay, instance);
        });
    }
    catch (const std::bad_alloc&)
    {
        // Out-of-memory is treated as fatal. Fail-fast so the crash pipeline can capture a dump.
        std::terminate();
    }
    catch (const std::exception&)
    {
#ifdef ENABLE_TESTS
        g_debugStage.store(23, std::memory_order_release);
#endif
        g_closeEvent.reset();
        g_threadStarted.store(false, std::memory_order_release);
    }
}

void RequestCloseIfExist() noexcept
{
    static_cast<void>(g_workerThread.request_stop());
    if (g_closeEvent)
    {
        static_cast<void>(SetEvent(g_closeEvent.get()));
    }

    const HWND hwnd = g_hwnd.load(std::memory_order_acquire);
    if (hwnd)
    {
        static_cast<void>(PostMessageW(hwnd, WM_CLOSE, 0, 0));
    }
}

void CloseIfExist() noexcept
{
    RequestCloseIfExist();

    if (g_workerThread.joinable() && g_workerThread.get_id() != std::this_thread::get_id())
    {
        g_workerThread.join();
        // The main thread owns the close-event lifetime; release it only after the worker
        // (including its stop callback) can no longer reference the handle.
        g_closeEvent.reset();
    }
}

bool Exist() noexcept
{
    return g_hwnd.load(std::memory_order_acquire) != nullptr;
}

HWND GetHwnd() noexcept
{
    return g_hwnd.load(std::memory_order_acquire);
}

void SetOwner(HWND owner) noexcept
{
    g_owner.store(owner, std::memory_order_release);

    const HWND hwnd = g_hwnd.load(std::memory_order_acquire);
    if (hwnd)
    {
        static_cast<void>(PostMessageW(hwnd, WndMsg::kSplashScreenRecenter, 0, 0));
    }
}

void IfExistSetText(std::wstring_view text) noexcept
{
    SetStatusText(text);

    const HWND hwnd = g_hwnd.load(std::memory_order_acquire);
    if (hwnd)
    {
        static_cast<void>(PostMessageW(hwnd, WndMsg::kSplashScreenSetText, 0, 0));
    }
}

#ifdef ENABLE_TESTS
DebugSnapshot DebugGetSnapshot() noexcept
{
    DebugSnapshot snapshot{};
    snapshot.threadStarted = g_threadStarted.load(std::memory_order_acquire);
    snapshot.hasHwnd       = g_hwnd.load(std::memory_order_acquire) != nullptr;
    snapshot.stage         = g_debugStage.load(std::memory_order_acquire);
    snapshot.lastError     = g_debugLastError.load(std::memory_order_acquire);
    snapshot.comHr         = g_debugComHr.load(std::memory_order_acquire);
    return snapshot;
}
#endif

} // namespace SplashScreen
