#pragma once

#include <algorithm>
#include <array>
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

#include "Common.Tests.WindowActivationGuard.h"
#include "Win32CallbackHelpers.h"
#include "WindowActivation.h"

namespace RedSalamander::TestSupport
{
// Shared version of the original Commands warning. An outer suite guard owns
// the surface; nested directed-input guards borrow it and never add a window.
// The suite must join its input workers before its guard leaves scope.
class DirectedSelfTestInputWarning final
{
public:
    struct ForegroundAdmissionDiagnostics final
    {
        const wchar_t* stage = L"not-started";
        DWORD win32Error     = ERROR_SUCCESS;
        int heldKey          = 0;
        UINT sentInputCount  = 0u;
    };

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

    // A registered hotkey grants its receiving thread foreground permission.
    // Never join a foreign input queue or alter system foreground-lock policy.
    // The explicit hotkey route also gives native controls a deterministic witness.
    [[nodiscard]] static bool TryActivateOwnedWindow(HWND target, bool requireHotkey = false, ForegroundAdmissionDiagnostics* diagnostics = nullptr) noexcept
    {
        ForegroundAdmissionDiagnostics localDiagnostics{};
        auto& observed       = diagnostics ? *diagnostics : localDiagnostics;
        observed             = {};
        observed.stage       = L"ownership-warning-or-activation-guard";
        const HWND warning   = _activeWindow.load(std::memory_order_acquire);
        DWORD processId      = 0u;
        const DWORD threadId = GetWindowThreadProcessId(target, &processId);
        if (ScopedWindowActivationBlocker::IsActiveForCurrentThread() || ! warning || ! IsWindowVisible(warning) || ! target ||
            processId != GetCurrentProcessId() || threadId != GetCurrentThreadId())
            return false;
        ForegroundHotkeyRequest request{.target = target, .warning = warning, .threadId = threadId};
        if (! requireHotkey && request.Activate())
        {
            observed.stage = L"direct-activation";
            return true;
        }

        // This nontext key is consumed by the registration on the owned target.
        // Do not release or override a key/modifier held by the person.
        for (const int key : {VK_F24, VK_SHIFT, VK_CONTROL, VK_MENU, VK_LWIN, VK_RWIN})
            if ((GetAsyncKeyState(key) & 0x8000) != 0)
            {
                observed.stage   = L"held-input";
                observed.heldKey = key;
                return false;
            }
        observed.stage = L"hotkey-atom-or-warning-thread";
        wil::unique_any<ATOM, decltype(&GlobalDeleteAtom), GlobalDeleteAtom> hotkeyId(
            GlobalAddAtomW(L"RedSalamander.TestForeground.{56157152-5E78-422C-A88F-73D02F93AA7A}"));
        if (! hotkeyId || GetWindowThreadProcessId(warning, nullptr) != threadId)
            return false;
        request.id     = hotkeyId.get();
        observed.stage = L"warning-hook";
        if (GetPropW(warning, kForegroundRequestProp) || Win32Callback::GetStoredWndProc(warning, kForegroundOriginalProcProp))
            return false;
        if (! Win32Callback::SetPropNoThrow(warning, kForegroundRequestProp, reinterpret_cast<HANDLE>(&request)))
        {
            observed.win32Error = GetLastError();
            return false;
        }
        const auto restoreHook = wil::scope_exit([&]() noexcept
        {
            Win32Callback::RestoreWndProcHook(warning, kForegroundOriginalProcProp, ForegroundHotkeyProc);
            RemovePropW(warning, kForegroundRequestProp);
        });
        if (! Win32Callback::InstallWndProcHook(warning, kForegroundOriginalProcProp, ForegroundHotkeyProc))
        {
            observed.win32Error = GetLastError();
            return false;
        }
        observed.stage = L"hotkey-registration";
        if (! RegisterHotKey(warning, static_cast<int>(request.id), MOD_NOREPEAT, VK_F24))
        {
            observed.win32Error = GetLastError();
            return false;
        }
        const auto unregister = wil::scope_exit([&]() noexcept { UnregisterHotKey(warning, static_cast<int>(request.id)); });
        std::array<INPUT, 2u> input{};
        input[0].type           = INPUT_KEYBOARD;
        input[0].ki.wVk         = VK_F24;
        input[1]                = input[0];
        input[1].ki.dwFlags     = KEYEVENTF_KEYUP;
        const UINT sent         = SendInput(static_cast<UINT>(input.size()), input.data(), sizeof(INPUT));
        observed.sentInputCount = sent;
        observed.stage          = L"hotkey-input";
        if (sent != input.size())
        {
            observed.win32Error = GetLastError();
            if (sent == 1u)
                static_cast<void>(SendInput(1u, &input[1], sizeof(INPUT)));
            return false;
        }
        const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds{500};
        observed.stage      = L"hotkey-dispatch-timeout";
        while (std::chrono::steady_clock::now() < deadline)
        {
            MSG message{};
            if (PeekMessageW(&message, warning, WM_HOTKEY, WM_HOTKEY, PM_REMOVE))
            {
                // Foreground permission belongs to handling the real hotkey.
                // Merely removing it and activating outside its window procedure
                // does not establish that dispatch boundary.
                DispatchMessageW(&message);
                if (request.handled)
                {
                    observed.stage = request.activated ? L"hotkey-activation" : L"hotkey-activation-refused";
                    return request.activated;
                }
            }
            static_cast<void>(MsgWaitForMultipleObjectsEx(0u, nullptr, 10u, QS_HOTKEY, MWMO_INPUTAVAILABLE));
        }
        return false;
    }

    [[nodiscard]] static bool IsWarningWindow(HWND window) noexcept
    {
        return window && window == _activeWindow.load(std::memory_order_acquire);
    }

private:
    static constexpr wchar_t kForegroundRequestProp[]      = L"RedSalamander.TestForeground.Request";
    static constexpr wchar_t kForegroundOriginalProcProp[] = L"RedSalamander.TestForeground.OriginalWndProc";

    struct ForegroundHotkeyRequest final
    {
        HWND target    = nullptr;
        HWND warning   = nullptr;
        DWORD threadId = 0u;
        UINT_PTR id    = 0u;
        bool handled   = false;
        bool activated = false;

        [[nodiscard]] bool Activate() const noexcept
        {
            DWORD currentOwner = 0u;
            if (GetWindowThreadProcessId(target, &currentOwner) != threadId || currentOwner != GetCurrentProcessId() || ! IsWindowVisible(warning))
                return false;
            ShowWindow(target, SW_SHOWNORMAL);
            BringWindowToTop(target);
            SetActiveWindow(target);
            static_cast<void>(Common::RequestForegroundWindow(target));
            SetFocus(target);
            SetWindowPos(target, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            SetWindowPos(target, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            UpdateWindow(target);
            return GetForegroundWindow() == target;
        }
    };

    static LRESULT CALLBACK ForegroundHotkeyProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam) noexcept
    {
        auto* const request = reinterpret_cast<ForegroundHotkeyRequest*>(GetPropW(window, kForegroundRequestProp));
        if (message == WM_NCDESTROY)
        {
            const auto original = Win32Callback::GetStoredWndProc(window, kForegroundOriginalProcProp);
            RemovePropW(window, kForegroundRequestProp);
            Win32Callback::RestoreWndProcHook(window, kForegroundOriginalProcProp, ForegroundHotkeyProc);
            return original ? Win32Callback::CallWindowProcNoThrow(original, window, message, wParam, lParam) : DefWindowProcW(window, message, wParam, lParam);
        }
        if (message == WM_HOTKEY && request && window == request->warning && wParam == request->id)
        {
            request->handled   = true;
            request->activated = request->Activate();
            return 0;
        }
        return Win32Callback::CallStoredWndProc(window, kForegroundOriginalProcProp, message, wParam, lParam);
    }

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
