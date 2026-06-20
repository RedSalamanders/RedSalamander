#include "RedConfigureSplashScreen.h"

#include "DxUi.h"
#include "Helpers.h"
#include "WindowMessages.h"
#include "resource.h"

#include <algorithm>
#include <atomic>
#include <chrono>
#include <cmath>
#include <exception>
#include <memory>
#include <mutex>
#include <string>
#include <thread>

#pragma warning(push)
// WIL headers: deleted copy/move and unused inline helpers
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace RedConfigure::SplashScreen
{
namespace
{
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::MakeDefaultThemePalette;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::ProgressBar;
using RedSalamander::DxUi::WindowHost;

constexpr wchar_t kSplashWindowClassName[] = L"RedConfigure.SplashWindow";
constexpr int kSplashWidthDip              = 520;
constexpr int kSplashHeightDip             = 190;
constexpr int kSplashRadiusDip             = 16;

std::atomic<bool> g_threadStarted{false};
wil::unique_event_nothrow g_closeEvent;
std::atomic<HWND> g_hwnd{nullptr};
std::atomic<HWND> g_owner{nullptr};
std::mutex g_textMutex;
std::wstring g_statusText;

#ifdef ENABLE_TESTS
std::atomic<unsigned long> g_debugStage{0};
std::atomic<unsigned long> g_debugLastError{0};
std::atomic<long> g_debugComHr{0};
#endif

// g_workerThread must be declared last: it is destroyed first at process teardown, so
// ~jthread joins the worker before the globals above (mutex, status text, close event) die.
std::jthread g_workerThread;

[[nodiscard]] D2D1_COLOR_F Rgba(float red, float green, float blue, float alpha = 1.0f) noexcept
{
    return D2D1::ColorF(red / 255.0f, green / 255.0f, blue / 255.0f, alpha);
}

[[nodiscard]] int ScaleDip(int dip, UINT dpi) noexcept
{
    return ::MulDiv(dip, static_cast<int>(dpi), USER_DEFAULT_SCREEN_DPI);
}

[[nodiscard]] std::wstring LoadAppString(HINSTANCE instance, UINT resourceId)
{
    wchar_t buffer[1024]{};
    const int length = ::LoadStringW(instance, resourceId, buffer, static_cast<int>(std::size(buffer)));
    return length > 0 ? std::wstring(buffer, static_cast<size_t>(length)) : std::wstring{};
}

[[nodiscard]] std::wstring GetStatusText()
{
    std::scoped_lock lock(g_textMutex);
    return g_statusText;
}

void SetStatusText(std::wstring_view text)
{
    std::scoped_lock lock(g_textMutex);
    g_statusText.assign(text);
}

struct SplashWindowState
{
    SplashWindowState()                                    = default;
    SplashWindowState(const SplashWindowState&)            = delete;
    SplashWindowState& operator=(const SplashWindowState&) = delete;
    SplashWindowState(SplashWindowState&&)                 = delete;
    SplashWindowState& operator=(SplashWindowState&&)      = delete;

    HINSTANCE instance = nullptr;
    WindowHost dxHost;
    class SplashRootControl* root = nullptr;
    Label* titleLabel             = nullptr;
    Label* statusLabel            = nullptr;
    Label* detailLabel            = nullptr;
    ProgressBar* progressBar      = nullptr;
};

LRESULT CALLBACK SplashWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) noexcept;
void CenterOverOwner(HWND hwnd, HWND owner) noexcept;
void LayoutSplashUi(SplashWindowState& state) noexcept;
void UpdateSplashLabels(SplashWindowState& state);
void StartSplashDrag(HWND hwnd) noexcept;

class SplashRootControl final : public Panel
{
public:
    bool Tick(WindowHost& host, uint64_t nowTickMs) override
    {
        static_cast<void>(Panel::Tick(host, nowTickMs));
        if (_lastTickMs == 0u)
        {
            _lastTickMs = nowTickMs;
        }

        const uint64_t elapsedMs = nowTickMs - _lastTickMs;
        _lastTickMs              = nowTickMs;
        _phase += static_cast<float>(elapsedMs) / 1600.0f;
        if (_phase >= 1.0f)
        {
            _phase -= static_cast<float>(static_cast<int>(_phase));
        }
        Invalidate(host);
        return true;
    }

    void Paint(WindowHost& host) const override
    {
        if (auto* dc = host.GetDeviceContext())
        {
            const D2D1_RECT_F bounds = host.GetClientBoundsDip();
            const D2D1_ROUNDED_RECT backgroundRect{bounds, 16.0f, 16.0f};

            const D2D1_GRADIENT_STOP stops[] = {
                {0.0f, Rgba(18.0f, 25.0f, 31.0f)},
                {0.58f, Rgba(27.0f, 39.0f, 49.0f)},
                {1.0f, Rgba(42.0f, 48.0f, 56.0f)},
            };

            wil::com_ptr<ID2D1GradientStopCollection> stopCollection;
            if (SUCCEEDED(dc->CreateGradientStopCollection(stops, static_cast<UINT32>(std::size(stops)), stopCollection.put())))
            {
                wil::com_ptr<ID2D1LinearGradientBrush> brush;
                const D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES props{D2D1::Point2F(bounds.left, bounds.top), D2D1::Point2F(bounds.right, bounds.bottom)};
                if (SUCCEEDED(dc->CreateLinearGradientBrush(props, stopCollection.get(), brush.put())))
                {
                    dc->FillRoundedRectangle(backgroundRect, brush.get());
                }
            }

            const auto borderBrush = host.GetSolidBrush(Rgba(86.0f, 101.0f, 116.0f));
            const auto accentBrush = host.GetSolidBrush(Rgba(255.0f, 190.0f, 96.0f));
            const auto mutedBrush  = host.GetSolidBrush(Rgba(130.0f, 147.0f, 162.0f, 0.58f));
            const auto glowBrush   = host.GetSolidBrush(Rgba(255.0f, 222.0f, 154.0f, 0.18f));
            if (borderBrush)
            {
                dc->DrawRoundedRectangle(backgroundRect, borderBrush, 1.0f);
            }

            const float centerX = 82.0f;
            const float centerY = (bounds.bottom - bounds.top) * 0.5f;
            const float radius  = 42.0f;
            if (mutedBrush && accentBrush && glowBrush)
            {
                const D2D1_ELLIPSE glow{D2D1::Point2F(centerX, centerY), radius + 14.0f, radius + 14.0f};
                dc->FillEllipse(glow, glowBrush);

                const D2D1_ELLIPSE orbit{D2D1::Point2F(centerX, centerY), radius, radius};
                dc->DrawEllipse(orbit, mutedBrush, 2.0f);

                constexpr float twoPi = 6.28318530718f;
                for (int index = 0; index < 3; ++index)
                {
                    const float angle = (_phase + (static_cast<float>(index) / 3.0f)) * twoPi;
                    const float dotX  = centerX + (std::cos(angle) * radius);
                    const float dotY  = centerY + (std::sin(angle) * radius);
                    const float dotR  = index == 0 ? 5.5f : 3.8f;
                    dc->FillEllipse(D2D1::Ellipse(D2D1::Point2F(dotX, dotY), dotR, dotR), accentBrush);
                }

                dc->DrawLine(D2D1::Point2F(centerX + 70.0f, 28.0f), D2D1::Point2F(centerX + 70.0f, bounds.bottom - 28.0f), accentBrush, 1.0f);
            }
        }

        Panel::Paint(host);
    }

    bool OnMouseDown(WindowHost& host, D2D1_POINT_2F point, bool rightButton, UINT modifiers) override
    {
        static_cast<void>(point);
        static_cast<void>(modifiers);
        if (rightButton)
        {
            return false;
        }

        StartSplashDrag(host.GetHwnd());
        return true;
    }

private:
    uint64_t _lastTickMs = 0u;
    float _phase         = 0.0f;
};

[[nodiscard]] bool EnsureSplashWindowClassRegistered(HINSTANCE instance) noexcept
{
    WNDCLASSEXW existing{};
    if (::GetClassInfoExW(instance, kSplashWindowClassName, &existing) != FALSE)
    {
        return true;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = SplashWindowProc;
    wc.hInstance     = instance;
    wc.hCursor       = ::LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kSplashWindowClassName;
    return ::RegisterClassExW(&wc) != 0 || ::GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

void UpdateSplashWindowSize(HWND hwnd) noexcept
{
    const UINT dpi     = ::GetDpiForWindow(hwnd);
    const int widthPx  = ScaleDip(kSplashWidthDip, dpi);
    const int heightPx = ScaleDip(kSplashHeightDip, dpi);
    ::SetWindowPos(hwnd, nullptr, 0, 0, widthPx, heightPx, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

    const int radius = ScaleDip(kSplashRadiusDip, dpi);
    wil::unique_hrgn region(::CreateRoundRectRgn(0, 0, widthPx + 1, heightPx + 1, radius, radius));
    if (region && ::SetWindowRgn(hwnd, region.get(), TRUE) != 0)
    {
        region.release();
    }
}

void LayoutSplashUi(SplashWindowState& state) noexcept
{
    if (! state.root)
    {
        return;
    }

    const D2D1_RECT_F bounds = state.dxHost.GetClientBoundsDip();
    state.root->SetBounds(bounds);

    const float textLeft  = 174.0f;
    const float textRight = std::max(textLeft + 120.0f, bounds.right - 26.0f);
    if (state.titleLabel)
    {
        state.titleLabel->SetBounds(D2D1::RectF(textLeft, 34.0f, textRight, 66.0f));
    }
    if (state.detailLabel)
    {
        state.detailLabel->SetBounds(D2D1::RectF(textLeft, 72.0f, textRight, 94.0f));
    }
    if (state.statusLabel)
    {
        state.statusLabel->SetBounds(D2D1::RectF(textLeft, 114.0f, textRight, 138.0f));
    }
    if (state.progressBar)
    {
        state.progressBar->SetBounds(D2D1::RectF(textLeft, 148.0f, textRight, 158.0f));
    }
}

void UpdateSplashLabels(SplashWindowState& state)
{
    if (state.titleLabel)
    {
        state.titleLabel->SetText(LoadAppString(state.instance, IDS_REDCONFIGURE_APP_TITLE));
    }
    if (state.detailLabel)
    {
        state.detailLabel->SetText(LoadAppString(state.instance, IDS_REDCONFIGURE_SPLASH_DETAIL));
    }
    if (state.statusLabel)
    {
        state.statusLabel->SetText(GetStatusText());
    }
}

void BuildSplashUi(SplashWindowState& state)
{
    auto root  = std::make_unique<SplashRootControl>();
    state.root = root.get();

    state.titleLabel = root->AddChild<Label>();
    state.titleLabel->SetFontRole(FontRole::Subtitle);
    state.titleLabel->SetTextColor(Rgba(246.0f, 248.0f, 252.0f));

    state.detailLabel = root->AddChild<Label>();
    state.detailLabel->SetFontRole(FontRole::Small);
    state.detailLabel->SetTextColor(Rgba(178.0f, 191.0f, 205.0f));

    state.statusLabel = root->AddChild<Label>();
    state.statusLabel->SetFontRole(FontRole::Small);
    state.statusLabel->SetTextColor(Rgba(255.0f, 202.0f, 125.0f));

    state.progressBar = root->AddChild<ProgressBar>();
    state.progressBar->SetIndeterminate(true);

    state.dxHost.SetTheme(MakeDefaultThemePalette(true));
    state.dxHost.SetRoot(std::move(root));
    UpdateSplashLabels(state);
    LayoutSplashUi(state);
    state.dxHost.RequestAnimation();
}

[[nodiscard]] RECT GetWorkAreaForOwner(HWND owner) noexcept
{
    HMONITOR monitor = nullptr;
    if (owner && ::IsWindow(owner))
    {
        monitor = ::MonitorFromWindow(owner, MONITOR_DEFAULTTONEAREST);
    }
    if (! monitor)
    {
        const HWND foreground = ::GetForegroundWindow();
        if (foreground && ::IsWindow(foreground))
        {
            monitor = ::MonitorFromWindow(foreground, MONITOR_DEFAULTTONEAREST);
        }
    }
    if (! monitor)
    {
        monitor = ::MonitorFromPoint(POINT{}, MONITOR_DEFAULTTOPRIMARY);
    }

    MONITORINFO monitorInfo{};
    monitorInfo.cbSize = sizeof(monitorInfo);
    if (monitor && ::GetMonitorInfoW(monitor, &monitorInfo))
    {
        return monitorInfo.rcWork;
    }

    RECT fallback{};
    static_cast<void>(::SystemParametersInfoW(SPI_GETWORKAREA, 0, &fallback, 0));
    return fallback;
}

void CenterOverOwner(HWND hwnd, HWND owner) noexcept
{
    RECT rect{};
    if (! hwnd || ! ::GetWindowRect(hwnd, &rect))
    {
        return;
    }

    const int width       = std::max(1, static_cast<int>(rect.right - rect.left));
    const int height      = std::max(1, static_cast<int>(rect.bottom - rect.top));
    const RECT workArea   = GetWorkAreaForOwner(owner);
    const int workLeft    = static_cast<int>(workArea.left);
    const int workTop     = static_cast<int>(workArea.top);
    const int workRight   = static_cast<int>(workArea.right);
    const int workBottom  = static_cast<int>(workArea.bottom);
    int targetCenterX     = (workLeft + workRight) / 2;
    int targetCenterY     = (workTop + workBottom) / 2;
    RECT ownerRect{};
    if (owner && ::IsWindow(owner) && ::GetWindowRect(owner, &ownerRect))
    {
        targetCenterX = (ownerRect.left + ownerRect.right) / 2;
        targetCenterY = (ownerRect.top + ownerRect.bottom) / 2;
    }

    const int maxLeft = std::max(workLeft, workRight - width);
    const int maxTop  = std::max(workTop, workBottom - height);
    const int left    = std::clamp(targetCenterX - (width / 2), workLeft, maxLeft);
    const int top     = std::clamp(targetCenterY - (height / 2), workTop, maxTop);
    ::SetWindowPos(hwnd, HWND_TOPMOST, left, top, 0, 0, SWP_NOSIZE | SWP_NOACTIVATE);
}

void StartSplashDrag(HWND hwnd) noexcept
{
    const HWND dragTarget = ::GetAncestor(hwnd, GA_ROOT);
    if (! dragTarget)
    {
        return;
    }

    ::ReleaseCapture();
    static_cast<void>(::SendMessageW(dragTarget, WM_NCLBUTTONDOWN, HTCAPTION, 0));
}

LRESULT CALLBACK SplashWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) noexcept
{
    auto* state = reinterpret_cast<SplashWindowState*>(::GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    if (msg == WM_NCCREATE)
    {
        const auto* create = reinterpret_cast<const CREATESTRUCTW*>(lParam);
        auto* createState  = create ? static_cast<SplashWindowState*>(create->lpCreateParams) : nullptr;
        ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(createState));
        return ::DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    if (state)
    {
        bool handled         = false;
        const LRESULT result = state->dxHost.HandleMessage(hwnd, msg, wParam, lParam, handled);
        if (handled)
        {
            if (msg == WM_NCDESTROY)
            {
                ::SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
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
                ::SetWindowPos(hwnd,
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
        case WM_CLOSE: ::DestroyWindow(hwnd); return 0;
        case WM_DESTROY:
            g_hwnd.store(nullptr, std::memory_order_release);
            ::PostQuitMessage(0);
            return 0;
        case WndMsg::kSplashScreenSetText:
            if (state)
            {
                UpdateSplashLabels(*state);
                state->dxHost.Invalidate();
            }
            return 0;
        case WndMsg::kSplashScreenRecenter: CenterOverOwner(hwnd, g_owner.load(std::memory_order_acquire)); return 0;
        default: break;
    }

    return ::DefWindowProcW(hwnd, msg, wParam, lParam);
}

void ThreadMain(std::stop_token stopToken, std::chrono::milliseconds delay, HINSTANCE instance) noexcept
{
#ifdef ENABLE_TESTS
    g_debugStage.store(1, std::memory_order_release);
    g_debugLastError.store(0, std::memory_order_release);
    g_debugComHr.store(S_OK, std::memory_order_release);
#endif
    const auto startedAt = std::chrono::steady_clock::now();
    const auto resetState = wil::scope_exit([]() noexcept
    {
        g_hwnd.store(nullptr, std::memory_order_release);
        // g_closeEvent is owned by the main thread (BeginOpen/CloseIfExist); resetting it
        // here would race a concurrent SetEvent from RequestCloseIfExist.
        g_threadStarted.store(false, std::memory_order_release);
    });

    const DWORD delayMs = delay.count() < 0 ? 0u : static_cast<DWORD>(delay.count());
    if (g_closeEvent && ::WaitForSingleObject(g_closeEvent.get(), delayMs) == WAIT_OBJECT_0)
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

    const HRESULT comHr = ::CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
#ifdef ENABLE_TESTS
    g_debugComHr.store(static_cast<long>(comHr), std::memory_order_release);
    g_debugStage.store(4, std::memory_order_release);
#endif
    const auto comCleanup = wil::scope_exit([&]() noexcept
    {
        if (SUCCEEDED(comHr))
        {
            ::CoUninitialize();
        }
    });

    if (! EnsureSplashWindowClassRegistered(instance))
    {
#ifdef ENABLE_TESTS
        g_debugLastError.store(::GetLastError(), std::memory_order_release);
        g_debugStage.store(5, std::memory_order_release);
#endif
        return;
    }

    if (Detail::ShouldAbortPendingOpen(stopToken, g_closeEvent.get()))
    {
#ifdef ENABLE_TESTS
        g_debugStage.store(6, std::memory_order_release);
#endif
        return;
    }

    auto splashState      = std::make_unique<SplashWindowState>();
    splashState->instance = instance;
    const HWND owner      = g_owner.load(std::memory_order_acquire);
    wil::unique_hwnd hwnd(::CreateWindowExW(WS_EX_TOOLWINDOW,
                                            kSplashWindowClassName,
                                            L"",
                                            WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                            CW_USEDEFAULT,
                                            CW_USEDEFAULT,
                                            ScaleDip(kSplashWidthDip, USER_DEFAULT_SCREEN_DPI),
                                            ScaleDip(kSplashHeightDip, USER_DEFAULT_SCREEN_DPI),
                                            owner,
                                            nullptr,
                                            instance,
                                            splashState.get()));
    if (! hwnd)
    {
#ifdef ENABLE_TESTS
        g_debugLastError.store(::GetLastError(), std::memory_order_release);
        g_debugStage.store(7, std::memory_order_release);
#endif
        return;
    }

    if (Detail::ShouldAbortPendingOpen(stopToken, g_closeEvent.get()))
    {
#ifdef ENABLE_TESTS
        g_debugStage.store(8, std::memory_order_release);
#endif
        hwnd.reset();
        return;
    }

    ::ShowWindow(hwnd.get(), SW_SHOWNOACTIVATE);
    ::SetWindowPos(hwnd.get(), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    ::UpdateWindow(hwnd.get());
    g_hwnd.store(hwnd.get(), std::memory_order_release);
    // Re-check after publishing the handle: a close requested between the abort check
    // above and the store would otherwise be lost (the closer only posts WM_CLOSE when it
    // sees a published hwnd) and GetMessageW below would block forever.
    if (Detail::ShouldAbortPendingOpen(stopToken, g_closeEvent.get()))
    {
#ifdef ENABLE_TESTS
        g_debugStage.store(10, std::memory_order_release);
#endif
        hwnd.reset();
        return;
    }
    Debug::Perf::Emit(L"redconfigure.startup.splash.visible_us", L"worker-ui", Debug::Perf::ElapsedUs(startedAt), 1u, 0u, S_OK);
#ifdef ENABLE_TESTS
    g_debugStage.store(9, std::memory_order_release);
#endif

    MSG msg{};
    while (::GetMessageW(&msg, nullptr, 0, 0) > 0)
    {
        ::TranslateMessage(&msg);
        ::DispatchMessageW(&msg);
    }

    if (::IsWindow(hwnd.get()) != FALSE)
    {
        hwnd.reset();
    }
    else
    {
        hwnd.release();
    }
}
} // namespace

void BeginOpen(std::chrono::milliseconds delay, HINSTANCE instance, std::wstring_view initialStatus) noexcept
{
    SetStatusText(initialStatus);

    const bool alreadyStarted = g_threadStarted.exchange(true, std::memory_order_acq_rel);
    if (alreadyStarted)
    {
        IfExistSetText(initialStatus);
        return;
    }

    // A previous worker may have finished (clearing g_threadStarted) without being joined;
    // join it before replacing the close event it might still reference.
    if (g_workerThread.joinable())
    {
        g_workerThread.join();
    }

    g_closeEvent.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! g_closeEvent)
    {
#ifdef ENABLE_TESTS
        g_debugLastError.store(::GetLastError(), std::memory_order_release);
        g_debugStage.store(20, std::memory_order_release);
#endif
        g_threadStarted.store(false, std::memory_order_release);
        return;
    }

    try
    {
        g_workerThread = std::jthread([delay, instance](std::stop_token stopToken) noexcept
        {
            std::stop_callback stopCallback(stopToken,
                                            []() noexcept
            {
                if (g_closeEvent)
                {
                    static_cast<void>(::SetEvent(g_closeEvent.get()));
                }

                const HWND hwnd = g_hwnd.load(std::memory_order_acquire);
                if (hwnd)
                {
                    static_cast<void>(::PostMessageW(hwnd, WM_CLOSE, 0, 0));
                }
            });

            ThreadMain(stopToken, delay, instance);
        });
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Splash is best-effort; thread creation failures must not block RedConfigure startup.
#ifdef ENABLE_TESTS
        g_debugStage.store(21, std::memory_order_release);
#endif
        g_closeEvent.reset();
        g_threadStarted.store(false, std::memory_order_release);
    }
}

void BeginImmediateOpen(HINSTANCE instance, std::wstring_view initialStatus) noexcept
{
    BeginOpen(std::chrono::milliseconds(0), instance, initialStatus);
}

void RequestCloseIfExist() noexcept
{
    static_cast<void>(g_workerThread.request_stop());
    if (g_closeEvent)
    {
        static_cast<void>(::SetEvent(g_closeEvent.get()));
    }

    const HWND hwnd = g_hwnd.load(std::memory_order_acquire);
    if (hwnd)
    {
        static_cast<void>(::PostMessageW(hwnd, WM_CLOSE, 0, 0));
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
        static_cast<void>(::PostMessageW(hwnd, WndMsg::kSplashScreenRecenter, 0, 0));
    }
}

void IfExistSetText(std::wstring_view text) noexcept
{
    SetStatusText(text);

    const HWND hwnd = g_hwnd.load(std::memory_order_acquire);
    if (hwnd)
    {
        static_cast<void>(::PostMessageW(hwnd, WndMsg::kSplashScreenSetText, 0, 0));
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
} // namespace RedConfigure::SplashScreen
