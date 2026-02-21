#include "SplashScreen.h"

#include "Framework.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <exception>
#include <cmath>
#include <mutex>
#include <stop_token>
#include <string>
#include <thread>

#pragma warning(push)
// WIL headers: deleted copy/move and unused inline Helpers
#pragma warning(disable: 4625 4626 5026 5027 4514 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include <commctrl.h>

#include "Helpers.h"
#include "Version.h"
#include "WindowMessages.h"
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
std::jthread g_workerThread;

std::mutex g_textMutex;
std::wstring g_statusText;

wil::unique_hicon g_logoIcon;
wil::unique_hfont g_titleFont;
constexpr UINT_PTR kSplashDragChildSubclassId = 0xA200u;
constexpr int kSplashLogoResourcePx           = 256;
constexpr int kSplashLogoDesignDip            = 162;
constexpr int kSplashContentOffsetDip         = 14;

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

LRESULT CALLBACK SplashChildDragProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR) noexcept
{
    static_cast<void>(wParam);
    static_cast<void>(lParam);
    if (msg == WM_LBUTTONDOWN)
    {
        StartSplashDrag(hwnd);
        return 0;
    }

    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

void InstallSplashChildDrag(HWND splashWnd) noexcept
{
    constexpr std::array<int, 4> dragControlIds{{
        IDC_SPLASH_REDSALAMANDER,
        IDC_SPLASH_VERSION,
        IDC_SPLASH_COPYRIGHT,
        IDC_SPLASH_STATUS,
    }};

    for (const int id : dragControlIds)
    {
        if (HWND control = GetDlgItem(splashWnd, id))
        {
            static_cast<void>(SetWindowSubclass(control, SplashChildDragProc, kSplashDragChildSubclassId, 0));
        }
    }
}

void RemoveSplashChildDrag(HWND splashWnd) noexcept
{
    constexpr std::array<int, 4> dragControlIds{{
        IDC_SPLASH_REDSALAMANDER,
        IDC_SPLASH_VERSION,
        IDC_SPLASH_COPYRIGHT,
        IDC_SPLASH_STATUS,
    }};

    for (const int id : dragControlIds)
    {
        if (HWND control = GetDlgItem(splashWnd, id))
        {
            static_cast<void>(RemoveWindowSubclass(control, SplashChildDragProc, kSplashDragChildSubclassId));
        }
    }
}

[[nodiscard]] int ScaleDip(int dip, UINT dpi) noexcept
{
    return MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
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
        POINT pt{};
        GetCursorPos(&pt);
        monitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
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

[[nodiscard]] wil::unique_hicon TryLoadLogoIcon(HINSTANCE instance) noexcept
{
    const HINSTANCE currentInstance = instance ? instance : GetModuleHandleW(nullptr);

    const auto tryLoadSizedIcon = [](HINSTANCE module, int resourceId) noexcept -> wil::unique_hicon
    {
        if (HICON icon = reinterpret_cast<HICON>(
                LoadImageW(module, MAKEINTRESOURCEW(resourceId), IMAGE_ICON, kSplashLogoResourcePx, kSplashLogoResourcePx, LR_DEFAULTCOLOR)))
        {
            return wil::unique_hicon{icon};
        }

        return nullptr;
    };

    if (auto splashIcon = tryLoadSizedIcon(currentInstance, IDI_SPLASH_LOGO_ICON); splashIcon)
    {
        return splashIcon;
    }

    if (auto fallback = tryLoadSizedIcon(currentInstance, IDI_REDSALAMANDER); fallback)
    {
        return fallback;
    }

    if (auto fallbackFromMain = tryLoadSizedIcon(GetModuleHandleW(nullptr), IDI_REDSALAMANDER); fallbackFromMain)
    {
        return fallbackFromMain;
    }

    if (auto fallbackSmall = tryLoadSizedIcon(currentInstance, IDI_SMALL); fallbackSmall)
    {
        return fallbackSmall;
    }

    return nullptr;
}

void PaintSplash(HWND hwnd, HDC hdc) noexcept
{
    if (! hwnd || ! hdc)
    {
        return;
    }

    RECT client{};
    GetClientRect(hwnd, &client);

    const int width  = std::max(1, static_cast<int>(client.right - client.left));
    const int height = std::max(1, static_cast<int>(client.bottom - client.top));
    const UINT dpi   = GetDpiForWindow(hwnd);

    // Background: warm dark vertical gradient.
    auto oldPen   = wil::SelectObject(hdc, GetStockObject(DC_PEN));
    auto oldBrush = wil::SelectObject(hdc, GetStockObject(DC_BRUSH));
    for (int y = 0; y < height; ++y)
    {
        const double t = (height > 1) ? (static_cast<double>(y) / static_cast<double>(height - 1)) : 0.0;
        const auto blendChannel = [](BYTE start, BYTE end, double weight) noexcept
        {
            const double value      = static_cast<double>(start) + (static_cast<double>(end) - static_cast<double>(start)) * weight;
            const int roundedValue  = static_cast<int>(std::lround(value));
            const int clampedValue  = std::clamp(roundedValue, 0, 255);
            return static_cast<BYTE>(clampedValue);
        };
        const BYTE r = blendChannel(ColorRefR(kBgStart), ColorRefR(kBgEnd), t);
        const BYTE g = blendChannel(ColorRefG(kBgStart), ColorRefG(kBgEnd), t);
        const BYTE b = blendChannel(ColorRefB(kBgStart), ColorRefB(kBgEnd), t);

        SetDCPenColor(hdc, MakeRGB(r, g, b));
        MoveToEx(hdc, 0, y, nullptr);
        LineTo(hdc, width, y);
    }

    // Moonlight Metal accents: faint diagonal panel.
    {
        const int panelRight = ScaleDip(340, dpi);
        const std::array<POINT, 4> panelPts{
            POINT{0, 0},
            POINT{panelRight, 0},
            POINT{ScaleDip(460, dpi), height},
            POINT{0, height},
        };
        constexpr COLORREF panelColor = MakeRGB(38, 50, 66);
        auto panelBrush           = wil::unique_hbrush{CreateSolidBrush(panelColor)};
        if (panelBrush)
        {
            auto oldFillBrush = wil::SelectObject(hdc, panelBrush.get());
            auto oldLinePen   = wil::SelectObject(hdc, GetStockObject(DC_PEN));
            SetDCPenColor(hdc, panelColor);
            Polygon(hdc, panelPts.data(), static_cast<int>(panelPts.size()));
        }
    }

    const int gapX     = ScaleDip(16, dpi);
    const int logoSize = std::clamp(ScaleDip(kSplashLogoDesignDip, dpi), ScaleDip(96, dpi), 256);

    const int logoX = -ScaleDip(38, dpi);
    const int logoY = std::max(0, (height - logoSize) / 2);

    if (g_logoIcon)
    {
        DrawIconEx(hdc, logoX, logoY, g_logoIcon.get(), logoSize, logoSize, 0, nullptr, DI_NORMAL);
    }
    else
    {
        const int pad  = std::max(1, ScaleDip(8, dpi));
        auto oldPen3   = wil::SelectObject(hdc, GetStockObject(DC_PEN));
        auto oldBrush3 = wil::SelectObject(hdc, GetStockObject(NULL_BRUSH));
        auto oldFont3  = wil::SelectObject(hdc, g_titleFont.get() ? g_titleFont.get() : GetStockObject(DEFAULT_GUI_FONT));
        SetDCPenColor(hdc, kFallbackRing);
        const int baseX = std::max(0, logoX);
        const int baseY = std::max(0, logoY);
        Ellipse(hdc, baseX, baseY, baseX + logoSize, baseY + logoSize);

        SetDCPenColor(hdc, kFallbackFrame);
        Rectangle(hdc, baseX + pad, baseY + pad, baseX + logoSize - pad, baseY + logoSize - pad);

        if (logoSize > pad * 4)
        {
            SetTextColor(hdc, kLogoFallbackText);
            const int oldBkMode = SetBkMode(hdc, TRANSPARENT);
            TextOutW(hdc, baseX + pad, baseY + pad, L"RS", 2);
            SetBkMode(hdc, oldBkMode);
        }
    }

    // Subtle separator between the logo and the text area.
    {
        const int separatorX = logoX + logoSize + (gapX / 2) + ScaleDip(kSplashContentOffsetDip, dpi);
        const int insetY     = ScaleDip(18, dpi);
        SetDCPenColor(hdc, kSeparator);
        MoveToEx(hdc, separatorX, insetY, nullptr);
        LineTo(hdc, separatorX, std::max(insetY, height - insetY));
    }

    // Subtle 1px border.
    {
        auto oldBorderBrush = wil::SelectObject(hdc, GetStockObject(HOLLOW_BRUSH));
        SetDCPenColor(hdc, kBorder);
        Rectangle(hdc, client.left, client.top, client.right, client.bottom);
    }
}

INT_PTR CALLBACK SplashDialogProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    static_cast<void>(lParam);
    static_cast<void>(wParam);

    switch (msg)
    {
        case WM_INITDIALOG:
        {
            // Size the window explicitly (DPI-aware) to keep the splash visually consistent.
            const UINT dpi     = GetDpiForWindow(hwnd);
            const int widthPx  = ScaleDip(560, dpi);
            const int heightPx = ScaleDip(220, dpi);
            SetWindowPos(hwnd, nullptr, 0, 0, widthPx, heightPx, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

            // Rounded corners.
            {
                const int radius = ScaleDip(14, dpi);
                wil::unique_hrgn rgn(CreateRoundRectRgn(0, 0, widthPx + 1, heightPx + 1, radius, radius));
                if (rgn)
                {
                    if (SetWindowRgn(hwnd, rgn.get(), TRUE) != 0)
                    {
                        rgn.release(); // owned by the window once applied successfully
                    }
                }
            }

            // Layout: position text controls to the right of the logo.
            {
                const int paddingX     = ScaleDip(20, dpi);
                const int paddingY     = ScaleDip(18, dpi);
                const int gapX         = ScaleDip(16, dpi);
                const int logoSize     = ScaleDip(162, dpi);
                const int textX        = std::max(paddingX, -ScaleDip(38, dpi) + logoSize + gapX + ScaleDip(kSplashContentOffsetDip, dpi));
                const int textWidth    = std::max(1, widthPx - textX - paddingX);
                const int titleHeight  = ScaleDip(30, dpi);
                const int metaHeight   = ScaleDip(18, dpi);
                const int statusHeight = ScaleDip(18, dpi);
                const int lineGap      = ScaleDip(6, dpi);

                int y = paddingY;
                if (HWND title = GetDlgItem(hwnd, IDC_SPLASH_REDSALAMANDER))
                {
                    SetWindowPos(title, nullptr, textX, y, textWidth, titleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                }
                y += titleHeight + lineGap;

                if (HWND version = GetDlgItem(hwnd, IDC_SPLASH_VERSION))
                {
                    SetWindowPos(version, nullptr, textX, y, textWidth, metaHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                }
                y += metaHeight + ScaleDip(2, dpi);

                if (HWND copyright = GetDlgItem(hwnd, IDC_SPLASH_COPYRIGHT))
                {
                    SetWindowPos(copyright, nullptr, textX, y, textWidth, metaHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                }

                int statusY = heightPx - paddingY - statusHeight;
                statusY     = std::max(statusY, y + lineGap);
                if (HWND status = GetDlgItem(hwnd, IDC_SPLASH_STATUS))
                {
                    SetWindowPos(status, nullptr, textX, statusY, textWidth, statusHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                }
            }

            wchar_t appTitle[128]{};
            const int titleLen = LoadStringW(GetModuleHandleW(nullptr), IDS_APP_TITLE, appTitle, static_cast<int>(sizeof(appTitle) / sizeof(appTitle[0])));
            if (titleLen > 0)
            {
                SetDlgItemTextW(hwnd, IDC_SPLASH_REDSALAMANDER, appTitle);
            }

            const std::wstring version = std::format(L"Version {}", VERSINFO_VERSION);
            SetDlgItemTextW(hwnd, IDC_SPLASH_VERSION, version.c_str());
            SetDlgItemTextW(hwnd, IDC_SPLASH_COPYRIGHT, VERSINFO_COPYRIGHT);

            const std::wstring status = GetStatusText();
            SetDlgItemTextW(hwnd, IDC_SPLASH_STATUS, status.c_str());

            // Typography: bold title.
            if (! g_titleFont)
            {
                LOGFONTW lf{};
                NONCLIENTMETRICSW ncm{};
                ncm.cbSize = sizeof(ncm);
                if (SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0))
                {
                    lf = ncm.lfMessageFont;
                }
                else
                {
                    lf.lfHeight = -MulDiv(14, static_cast<int>(dpi), 72);
                    wcscpy_s(lf.lfFaceName, L"Segoe UI");
                }

                lf.lfWeight = FW_SEMIBOLD;
                lf.lfHeight = -MulDiv(18, static_cast<int>(dpi), 72);
                g_titleFont.reset(CreateFontIndirectW(&lf));
            }

            if (g_titleFont)
            {
                if (HWND title = GetDlgItem(hwnd, IDC_SPLASH_REDSALAMANDER))
                {
                    SendMessageW(title, WM_SETFONT, reinterpret_cast<WPARAM>(g_titleFont.get()), TRUE);
                }
            }

            InstallSplashChildDrag(hwnd);
            CenterOverOwner(hwnd, g_owner.load(std::memory_order_acquire));
            return static_cast<INT_PTR>(TRUE);
        }
        case WM_ERASEBKGND:
            // We paint a full custom background in WM_PAINT.
            return static_cast<INT_PTR>(TRUE);
        case WM_LBUTTONDOWN: StartSplashDrag(hwnd); return static_cast<INT_PTR>(TRUE);
        case WM_PAINT:
        {
            wil::unique_hdc_paint dc = wil::BeginPaint(hwnd);
            if (dc)
            {
                PaintSplash(hwnd, dc.get());
            }
            return static_cast<INT_PTR>(TRUE);
        }
        case WM_CTLCOLORDLG:
        case WM_CTLCOLORSTATIC:
        {
            const HDC hdc = reinterpret_cast<HDC>(wParam);
            if (hdc)
            {
                SetBkMode(hdc, TRANSPARENT);
                const HWND ctl = reinterpret_cast<HWND>(lParam);

                if (ctl == GetDlgItem(hwnd, IDC_SPLASH_REDSALAMANDER))
                {
                    SetTextColor(hdc, kTitleText);
                }
                else if (ctl == GetDlgItem(hwnd, IDC_SPLASH_STATUS))
                {
                    SetTextColor(hdc, kStatusText);
                }
                else
                {
                    SetTextColor(hdc, kSecondaryText);
                }
            }
            return reinterpret_cast<INT_PTR>(GetStockObject(NULL_BRUSH));
        }
        case WM_CLOSE: DestroyWindow(hwnd); return static_cast<INT_PTR>(TRUE);
        case WM_DESTROY:
            RemoveSplashChildDrag(hwnd);
            g_hwnd.store(nullptr, std::memory_order_release);
            PostQuitMessage(0);
            return static_cast<INT_PTR>(TRUE);
        case WndMsg::kSplashScreenSetText:
        {
            const std::wstring status = GetStatusText();
            SetDlgItemTextW(hwnd, IDC_SPLASH_STATUS, status.c_str());
            InvalidateRect(hwnd, nullptr, FALSE);
            return static_cast<INT_PTR>(TRUE);
        }
        case WndMsg::kSplashScreenRecenter: CenterOverOwner(hwnd, g_owner.load(std::memory_order_acquire)); return static_cast<INT_PTR>(TRUE);
    }

    return static_cast<INT_PTR>(FALSE);
}

void ThreadMain(std::chrono::milliseconds delay, HINSTANCE instance) noexcept
{
    if (! g_closeEvent)
    {
        return;
    }

    const DWORD delayMs = delay.count() < 0 ? 0u : static_cast<DWORD>(delay.count());
    const DWORD wait    = WaitForSingleObject(g_closeEvent.get(), delayMs);
    if (wait == WAIT_OBJECT_0)
    {
        return;
    }

    const HRESULT comHr   = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const auto comCleanup = wil::scope_exit(
        [&]
        {
            if (SUCCEEDED(comHr))
            {
                CoUninitialize();
            }
        });

    if (! g_logoIcon)
    {
        g_logoIcon = TryLoadLogoIcon(instance);
    }

    // The dialog proc is written to be exception-free; suppress at call site for MSVC (-EHc).
#pragma warning(push)
#pragma warning(disable : 5039)
    wil::unique_hwnd hwnd(CreateDialogParamW(instance, MAKEINTRESOURCEW(IDD_SPLASH), nullptr, SplashDialogProc, 0));
#pragma warning(pop)
    if (! hwnd)
    {
        return;
    }

    g_hwnd.store(hwnd.get(), std::memory_order_release);

    CenterOverOwner(hwnd.get(), g_owner.load(std::memory_order_acquire));
    ShowWindow(hwnd.get(), SW_SHOWNOACTIVATE);
    SetWindowPos(hwnd.get(), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    UpdateWindow(hwnd.get());

    MSG msg{};
    HANDLE waitHandles[] = {g_closeEvent.get()};
    for (;;)
    {
        const DWORD msgWait = MsgWaitForMultipleObjects(1, waitHandles, FALSE, INFINITE, QS_ALLINPUT);
        if (msgWait == WAIT_OBJECT_0)
        {
            hwnd.reset();
        }

        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) != 0)
        {
            if (msg.message == WM_QUIT)
            {
                hwnd.reset();
                return;
            }

            if (! IsDialogMessageW(hwnd.get(), &msg))
            {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
    }
}
} // namespace

void BeginDelayedOpen(std::chrono::milliseconds delay, HINSTANCE instance) noexcept
{
    const bool alreadyStarted = g_threadStarted.exchange(true, std::memory_order_acq_rel);
    if (alreadyStarted)
    {
        return;
    }

    g_closeEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! g_closeEvent)
    {
        g_threadStarted.store(false, std::memory_order_release);
        return;
    }

    // Mandatory: `noexcept` boundary. Splash is best-effort; thread creation can throw.
    try
    {
        g_workerThread = std::jthread(
            [delay, instance](std::stop_token stopToken) noexcept
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

                ThreadMain(delay, instance);
            });
    }
    catch (const std::bad_alloc&)
    {
        // Out-of-memory is treated as fatal. Fail-fast so the crash pipeline can capture a dump.
        std::terminate();
    }
    catch (const std::exception&)
    {
        g_closeEvent.reset();
        g_threadStarted.store(false, std::memory_order_release);
    }
}

void CloseIfExist() noexcept
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

} // namespace SplashScreen
