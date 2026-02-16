#include "SplashScreen.h"

#include "Framework.h"

#include <algorithm>
#include <atomic>
#include <chrono>
#include <exception>
#include <mutex>
#include <stop_token>
#include <string>
#include <thread>

#include <wil/com.h>
#include <wil/resource.h>

#include <wincodec.h>

#include "Helpers.h"
#include "Version.h"
#include "WindowMessages.h"
#include "resource.h"

namespace SplashScreen
{
namespace
{
std::atomic<bool> g_threadStarted{false};
wil::unique_event_nothrow g_closeEvent;
std::atomic<HWND> g_hwnd{nullptr};
std::atomic<HWND> g_owner{nullptr};
std::jthread g_workerThread;

std::mutex g_textMutex;
std::wstring g_statusText;

wil::unique_hbitmap g_logoBitmap;
SIZE g_logoSize{};
wil::unique_hfont g_titleFont;

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

// AlphaBlend replacement (software, straight alpha)
[[nodiscard]] BOOL BlitAlphaBlend(HDC hdcDest,
                                  int xoriginDest,
                                  int yoriginDest,
                                  int wDest,
                                  int hDest,
                                  HDC hdcSrc,
                                  int xoriginSrc,
                                  int yoriginSrc,
                                  int wSrc,
                                  int hSrc,
                                  BLENDFUNCTION ftn) noexcept
{
    if (! hdcDest || ! hdcSrc || wDest <= 0 || hDest <= 0 || wSrc <= 0 || hSrc <= 0)
    {
        return TRUE;
    }
    if (ftn.BlendOp != AC_SRC_OVER)
    {
        return FALSE;
    }

    const bool useSrcAlpha     = (ftn.AlphaFormat & AC_SRC_ALPHA) != 0;
    const uint32_t globalAlpha = ftn.SourceConstantAlpha;

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = wDest;
    bmi.bmiHeader.biHeight      = -hDest; // top-down
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* destBits = nullptr;
    wil::unique_hbitmap destDib(CreateDIBSection(hdcDest, &bmi, DIB_RGB_COLORS, &destBits, nullptr, 0));
    if (! destDib || ! destBits)
    {
        return FALSE;
    }

    wil::unique_hdc destMem(CreateCompatibleDC(hdcDest));
    if (! destMem)
    {
        return FALSE;
    }
    auto oldDestBmp = wil::SelectObject(destMem.get(), destDib.get());
    if (! BitBlt(destMem.get(), 0, 0, wDest, hDest, hdcDest, xoriginDest, yoriginDest, SRCCOPY))
    {
        return FALSE;
    }

    void* srcBits = nullptr;
    wil::unique_hbitmap srcDib(CreateDIBSection(hdcDest, &bmi, DIB_RGB_COLORS, &srcBits, nullptr, 0));
    if (! srcDib || ! srcBits)
    {
        return FALSE;
    }

    wil::unique_hdc srcMem(CreateCompatibleDC(hdcDest));
    if (! srcMem)
    {
        return FALSE;
    }
    auto oldSrcBmp = wil::SelectObject(srcMem.get(), srcDib.get());

    const int prevStretch = SetStretchBltMode(srcMem.get(), HALFTONE);
    const bool copyOk     = (wSrc == wDest && hSrc == hDest) ? BitBlt(srcMem.get(), 0, 0, wDest, hDest, hdcSrc, xoriginSrc, yoriginSrc, SRCCOPY)
                                                             : StretchBlt(srcMem.get(), 0, 0, wDest, hDest, hdcSrc, xoriginSrc, yoriginSrc, wSrc, hSrc, SRCCOPY);
    SetStretchBltMode(srcMem.get(), prevStretch);
    if (! copyOk)
    {
        return FALSE;
    }

    auto* dst = static_cast<uint32_t*>(destBits);
    auto* src = static_cast<uint32_t*>(srcBits);

    for (int y = 0; y < hDest; ++y)
    {
        const auto rowOffset = static_cast<size_t>(y) * static_cast<size_t>(wDest);
        for (int x = 0; x < wDest; ++x)
        {
            const uint32_t s     = src[rowOffset + static_cast<size_t>(x)];
            const uint8_t srcA   = useSrcAlpha ? static_cast<uint8_t>(s >> 24) : 255u;
            const uint32_t alpha = (static_cast<uint32_t>(srcA) * globalAlpha + 127u) / 255u;
            if (alpha == 0)
            {
                continue;
            }

            const uint32_t invA = 255u - alpha;

            const uint8_t srcB = static_cast<uint8_t>(s);
            const uint8_t srcG = static_cast<uint8_t>(s >> 8);
            const uint8_t srcR = static_cast<uint8_t>(s >> 16);

            const uint32_t d   = dst[rowOffset + static_cast<size_t>(x)];
            const uint8_t dstB = static_cast<uint8_t>(d);
            const uint8_t dstG = static_cast<uint8_t>(d >> 8);
            const uint8_t dstR = static_cast<uint8_t>(d >> 16);

            const uint8_t outB = static_cast<uint8_t>((static_cast<uint32_t>(srcB) * alpha + static_cast<uint32_t>(dstB) * invA + 127u) / 255u);
            const uint8_t outG = static_cast<uint8_t>((static_cast<uint32_t>(srcG) * alpha + static_cast<uint32_t>(dstG) * invA + 127u) / 255u);
            const uint8_t outR = static_cast<uint8_t>((static_cast<uint32_t>(srcR) * alpha + static_cast<uint32_t>(dstR) * invA + 127u) / 255u);

            dst[rowOffset + static_cast<size_t>(x)] = (static_cast<uint32_t>(outR) << 16) | (static_cast<uint32_t>(outG) << 8) | outB | 0xFF000000u;
        }
    }

    return BitBlt(hdcDest, xoriginDest, yoriginDest, wDest, hDest, destMem.get(), 0, 0, SRCCOPY) != 0;
}

[[nodiscard]] wil::unique_hbitmap TryLoadLogoPng(HINSTANCE instance, SIZE& outSize) noexcept
{
    outSize = {};

    HRSRC resInfo = FindResourceW(instance, MAKEINTRESOURCEW(IDR_SPLASH_LOGO_PNG), L"PNG");
    if (! resInfo)
    {
        return nullptr;
    }

    HGLOBAL resData = LoadResource(instance, resInfo);
    if (! resData)
    {
        return nullptr;
    }

    void* bytes           = LockResource(resData);
    const DWORD byteCount = SizeofResource(instance, resInfo);
    if (! bytes || byteCount == 0)
    {
        return nullptr;
    }

    wil::com_ptr<IWICImagingFactory> factory;
    HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory));
    if (FAILED(hr) || ! factory)
    {
        return nullptr;
    }

    wil::com_ptr<IWICStream> stream;
    hr = factory->CreateStream(&stream);
    if (FAILED(hr) || ! stream)
    {
        return nullptr;
    }

    hr = stream->InitializeFromMemory(static_cast<BYTE*>(bytes), byteCount);
    if (FAILED(hr))
    {
        return nullptr;
    }

    wil::com_ptr<IWICBitmapDecoder> decoder;
    hr = factory->CreateDecoderFromStream(stream.get(), nullptr, WICDecodeMetadataCacheOnLoad, &decoder);
    if (FAILED(hr) || ! decoder)
    {
        return nullptr;
    }

    wil::com_ptr<IWICBitmapFrameDecode> frame;
    hr = decoder->GetFrame(0, &frame);
    if (FAILED(hr) || ! frame)
    {
        return nullptr;
    }

    UINT width  = 0;
    UINT height = 0;
    hr          = frame->GetSize(&width, &height);
    if (FAILED(hr) || width == 0 || height == 0)
    {
        return nullptr;
    }

    wil::com_ptr<IWICFormatConverter> converter;
    hr = factory->CreateFormatConverter(&converter);
    if (FAILED(hr) || ! converter)
    {
        return nullptr;
    }

    hr = converter->Initialize(frame.get(), GUID_WICPixelFormat32bppBGRA, WICBitmapDitherTypeNone, nullptr, 0.0, WICBitmapPaletteTypeCustom);
    if (FAILED(hr))
    {
        return nullptr;
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = static_cast<LONG>(width);
    bmi.bmiHeader.biHeight      = -static_cast<LONG>(height); // top-down
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* dibBits = nullptr;
    wil::unique_hdc_window screen{GetDC(nullptr)};
    wil::unique_hbitmap dib{CreateDIBSection(screen.get(), &bmi, DIB_RGB_COLORS, &dibBits, nullptr, 0)};
    if (! dib || ! dibBits)
    {
        return nullptr;
    }

    const UINT stride      = width * 4u;
    const UINT bytesToCopy = stride * height;
    WICRect rect{0, 0, static_cast<INT>(width), static_cast<INT>(height)};
    hr = converter->CopyPixels(&rect, stride, bytesToCopy, static_cast<BYTE*>(dibBits));
    if (FAILED(hr))
    {
        return nullptr;
    }

    outSize.cx = static_cast<LONG>(width);
    outSize.cy = static_cast<LONG>(height);
    return dib;
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

    // Background: subtle dark vertical gradient (fast GDI line fill).
    auto oldPen   = wil::SelectObject(hdc, GetStockObject(DC_PEN));
    auto oldBrush = wil::SelectObject(hdc, GetStockObject(DC_BRUSH));
    for (int y = 0; y < height; ++y)
    {
        const double t = (height > 1) ? (static_cast<double>(y) / static_cast<double>(height - 1)) : 0.0;
        const int r    = static_cast<int>(16.0 + (10.0 - 16.0) * t);
        const int g    = static_cast<int>(16.0 + (10.0 - 16.0) * t);
        const int b    = static_cast<int>(20.0 + (12.0 - 20.0) * t);

        SetDCPenColor(hdc, RGB(r, g, b));
        MoveToEx(hdc, 0, y, nullptr);
        LineTo(hdc, width, y);
    }

    const UINT dpi     = GetDpiForWindow(hwnd);
    const int paddingX = ScaleDip(20, dpi);
    const int gapX     = ScaleDip(16, dpi);
    const int logoSize = ScaleDip(96, dpi);

    const int logoX   = paddingX;
    const int logoY   = std::max(0, (height - logoSize) / 2);
    const int centerX = logoX + (logoSize / 2);
    const int centerY = logoY + (logoSize / 2);

    // Warm radial highlight behind the logo (concentric circles).
    {
        auto oldNoPen = wil::SelectObject(hdc, GetStockObject(NULL_PEN));

        const int maxRadius = static_cast<int>(static_cast<double>(logoSize) * 0.9);
        const int minRadius = static_cast<int>(static_cast<double>(logoSize) * 0.35);
        const int steps     = 10;
        for (int i = 0; i < steps; ++i)
        {
            const double t   = static_cast<double>(i) / static_cast<double>(std::max(1, steps - 1)); // 0..1
            const double inv = 1.0 - t;
            const int radius = static_cast<int>(static_cast<double>(minRadius) + inv * static_cast<double>(maxRadius - minRadius));

            const double intensity = inv * inv;
            const int rr           = static_cast<int>(18.0 + intensity * 110.0);
            const int gg           = static_cast<int>(12.0 + intensity * 45.0);
            const int bb           = static_cast<int>(12.0 + intensity * 10.0);

            SetDCBrushColor(hdc, RGB(rr, gg, bb));
            Ellipse(hdc, centerX - radius, centerY - radius, centerX + radius, centerY + radius);
        }
    }

    // Logo (alpha blended).
    if (g_logoBitmap && g_logoSize.cx > 0 && g_logoSize.cy > 0)
    {
        wil::unique_hdc mem(CreateCompatibleDC(hdc));
        if (mem)
        {
            auto oldBmp = wil::SelectObject(mem.get(), g_logoBitmap.get());
            BLENDFUNCTION blend{};
            blend.BlendOp             = AC_SRC_OVER;
            blend.SourceConstantAlpha = 255;
            blend.AlphaFormat         = AC_SRC_ALPHA;
            static_cast<void>(BlitAlphaBlend(hdc, logoX, logoY, logoSize, logoSize, mem.get(), 0, 0, g_logoSize.cx, g_logoSize.cy, blend));
        }
    }

    // Subtle separator between the logo and the text area.
    {
        const int separatorX = logoX + logoSize + (gapX / 2);
        const int insetY     = ScaleDip(18, dpi);
        SetDCPenColor(hdc, RGB(40, 40, 48));
        MoveToEx(hdc, separatorX, insetY, nullptr);
        LineTo(hdc, separatorX, std::max(insetY, height - insetY));
    }

    // Subtle 1px border.
    {
        auto oldBorderBrush = wil::SelectObject(hdc, GetStockObject(HOLLOW_BRUSH));
        SetDCPenColor(hdc, RGB(48, 48, 56));
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
            const int widthPx  = ScaleDip(480, dpi);
            const int heightPx = ScaleDip(200, dpi);
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
                const int logoSize     = ScaleDip(96, dpi);
                const int textX        = paddingX + logoSize + gapX;
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

            CenterOverOwner(hwnd, g_owner.load(std::memory_order_acquire));
            return static_cast<INT_PTR>(TRUE);
        }
        case WM_ERASEBKGND:
            // We paint a full custom background in WM_PAINT.
            return static_cast<INT_PTR>(TRUE);
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
                    SetTextColor(hdc, RGB(250, 250, 252));
                }
                else if (ctl == GetDlgItem(hwnd, IDC_SPLASH_STATUS))
                {
                    SetTextColor(hdc, RGB(230, 200, 160));
                }
                else
                {
                    SetTextColor(hdc, RGB(200, 200, 208));
                }
            }
            return reinterpret_cast<INT_PTR>(GetStockObject(NULL_BRUSH));
        }
        case WM_CLOSE: DestroyWindow(hwnd); return static_cast<INT_PTR>(TRUE);
        case WM_DESTROY:
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

    if (! g_logoBitmap)
    {
        SIZE size{};
        if (auto bmp = TryLoadLogoPng(instance, size))
        {
            g_logoSize   = size;
            g_logoBitmap = std::move(bmp);
        }
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
