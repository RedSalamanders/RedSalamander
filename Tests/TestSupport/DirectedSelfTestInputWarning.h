#pragma once

#include <algorithm>
#include <atomic>
#include <chrono>
#include <thread>
#include <utility>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace RedSalamander::TestSupport
{
// Shared version of the original Commands warning. An outer suite guard owns
// the surface; nested directed-input guards borrow it and never add a window.
// The suite must join its input workers before its guard leaves scope.
class DirectedSelfTestInputWarning final
{
public:
    explicit DirectedSelfTestInputWarning(HWND owner = nullptr, bool enabled = true) noexcept
    {
        if (! enabled)
        {
            return;
        }
        _borrowed = _activeWindow.load(std::memory_order_acquire);
        if (_borrowed && ::IsWindow(_borrowed))
        {
            return;
        }
        _owner = owner;
        _hwnd.reset(CreateWindowExW(WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_LAYERED | WS_EX_TRANSPARENT,
                                    L"STATIC",
                                    L"\nTests need keyboard and mouse\nPause input until this warning closes.",
                                    WS_POPUP | WS_BORDER | SS_CENTER,
                                    0,
                                    0,
                                    800,
                                    160,
                                    nullptr,
                                    nullptr,
                                    GetModuleHandleW(nullptr),
                                    nullptr));
        if (! _hwnd)
        {
            return;
        }
        HWND expected = nullptr;
        if (! _activeWindow.compare_exchange_strong(expected, _hwnd.get(), std::memory_order_acq_rel))
        {
            _borrowed = expected;
            _hwnd.reset();
            return;
        }
        SetWindowLongPtrW(_hwnd.get(), GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));
        if (! SetLayeredWindowAttributes(_hwnd.get(), 0, 240, LWA_ALPHA))
        {
            _activeWindow.store(nullptr, std::memory_order_release);
            _hwnd.reset();
            return;
        }
        Reposition();
        // Window-owned timers are destroyed with the RAII-owned HWND.
        if (! SetTimer(_hwnd.get(), 1u, 200u, TimerProc))
        {
            _activeWindow.store(nullptr, std::memory_order_release);
            _hwnd.reset();
            return;
        }
        ShowWindow(_hwnd.get(), SW_SHOWNOACTIVATE);
        UpdateWindow(_hwnd.get());
        const ULONGLONG deadline = GetTickCount64() + 3000u;
        do
        {
            MSG message{};
            for (size_t count = 0; count < 512u && PeekMessageW(&message, nullptr, 0, 0, PM_REMOVE); ++count)
            {
                TranslateMessage(&message);
                DispatchMessageW(&message);
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
        } while (GetTickCount64() < deadline);
    }

    ~DirectedSelfTestInputWarning() noexcept
    {
        if (_hwnd)
        {
            _activeWindow.store(nullptr, std::memory_order_release);
            _hwnd.reset();
        }
    }
    DirectedSelfTestInputWarning(const DirectedSelfTestInputWarning&)            = delete;
    DirectedSelfTestInputWarning& operator=(const DirectedSelfTestInputWarning&) = delete;
    DirectedSelfTestInputWarning(DirectedSelfTestInputWarning&&)                 = delete;
    DirectedSelfTestInputWarning& operator=(DirectedSelfTestInputWarning&&)      = delete;

    [[nodiscard]] bool IsVisible() const noexcept
    {
        const HWND window = _hwnd ? _hwnd.get() : _borrowed;
        return window && ::IsWindow(window) && IsWindowVisible(window);
    }

    [[nodiscard]] static bool IsWarningWindow(HWND window) noexcept
    {
        return window && window == _activeWindow.load(std::memory_order_acquire);
    }

private:
    static void CALLBACK TimerProc(HWND window, UINT, UINT_PTR, DWORD) noexcept
    {
        if (IsWarningWindow(window))
        {
            auto* warning = reinterpret_cast<DirectedSelfTestInputWarning*>(GetWindowLongPtrW(window, GWLP_USERDATA));
            if (warning)
            {
                warning->Reposition();
            }
        }
    }

    void Reposition() noexcept
    {
        HWND target           = _owner;
        const HWND foreground = GetForegroundWindow();
        DWORD processId       = 0;
        GetWindowThreadProcessId(foreground, &processId);
        if (processId == GetCurrentProcessId() && foreground != _hwnd.get())
        {
            target = foreground;
            _owner = target;
        }
        MONITORINFO monitor{sizeof(MONITORINFO)};
        if (! GetMonitorInfoW(MonitorFromWindow(target, MONITOR_DEFAULTTOPRIMARY), &monitor))
        {
            return;
        }
        RECT anchor = monitor.rcWork;
        RECT bounds{};
        if (target && GetWindowRect(target, &bounds) && bounds.right > bounds.left && bounds.bottom > bounds.top)
        {
            anchor = bounds;
        }
        const UINT dpi          = target ? GetDpiForWindow(target) : 96u;
        const UINT effectiveDpi = dpi ? dpi : 96u;
        if (effectiveDpi != _dpi)
        {
            wil::unique_hfont font(CreateFontW(-MulDiv(32, static_cast<int>(effectiveDpi), 96),
                                               0,
                                               0,
                                               0,
                                               FW_SEMIBOLD,
                                               FALSE,
                                               FALSE,
                                               FALSE,
                                               DEFAULT_CHARSET,
                                               OUT_DEFAULT_PRECIS,
                                               CLIP_DEFAULT_PRECIS,
                                               CLEARTYPE_QUALITY,
                                               DEFAULT_PITCH | FF_SWISS,
                                               L"Segoe UI"));
            if (font)
            {
                SendMessageW(_hwnd.get(), WM_SETFONT, reinterpret_cast<WPARAM>(font.get()), TRUE);
                _font = std::move(font);
                _dpi  = effectiveDpi;
            }
        }
        const int width  = (std::min)(MulDiv(800, static_cast<int>(effectiveDpi), 96), static_cast<int>(monitor.rcWork.right - monitor.rcWork.left));
        const int height = (std::min)(MulDiv(160, static_cast<int>(effectiveDpi), 96), static_cast<int>(monitor.rcWork.bottom - monitor.rcWork.top));
        const int left   = std::clamp(static_cast<int>(anchor.left + (anchor.right - anchor.left - width) / 2),
                                      static_cast<int>(monitor.rcWork.left),
                                      static_cast<int>(monitor.rcWork.right) - width);
        const int top    = std::clamp(static_cast<int>(anchor.top + (anchor.bottom - anchor.top - height) / 2),
                                      static_cast<int>(monitor.rcWork.top),
                                      static_cast<int>(monitor.rcWork.bottom) - height);
        RECT current{};
        if (! GetWindowRect(_hwnd.get(), &current) || current.left != left || current.top != top || current.right - current.left != width ||
            current.bottom - current.top != height)
        {
            SetWindowPos(_hwnd.get(), HWND_TOPMOST, left, top, width, height, SWP_NOACTIVATE);
        }
    }

    inline static std::atomic<HWND> _activeWindow{nullptr};
    wil::unique_hfont _font;
    wil::unique_hwnd _hwnd;
    HWND _owner    = nullptr;
    HWND _borrowed = nullptr;
    UINT _dpi      = 0;
};
} // namespace RedSalamander::TestSupport
