#pragma once

#include <Windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include <cstdint>

namespace RedSalamander::TestSupport
{
// Test-only guard for GUI coverage whose contract does not include real desktop
// focus. A thread CBT hook adds WS_EX_NOACTIVATE to every top-level window and
// rejects activation/focus transfers before production ShowWindow or
// SetForegroundWindow calls can disturb the user's foreground application.
class ScopedWindowActivationBlocker final
{
public:
    ScopedWindowActivationBlocker()                                                = default;
    ScopedWindowActivationBlocker(const ScopedWindowActivationBlocker&)            = delete;
    ScopedWindowActivationBlocker& operator=(const ScopedWindowActivationBlocker&) = delete;
    ScopedWindowActivationBlocker(ScopedWindowActivationBlocker&&)                 = delete;
    ScopedWindowActivationBlocker& operator=(ScopedWindowActivationBlocker&&)      = delete;

    ~ScopedWindowActivationBlocker() noexcept
    {
        Stop();
    }

    [[nodiscard]] bool Start() noexcept
    {
        if (_hook || s_activeBlocker != nullptr)
        {
            return false;
        }

        _ownerThreadId = GetCurrentThreadId();
        _hook.reset(SetWindowsHookExW(WH_CBT, HookProc, nullptr, _ownerThreadId));
        if (! _hook)
        {
            _ownerThreadId = 0u;
            return false;
        }

        s_activeBlocker = this;
        return true;
    }

    void Stop() noexcept
    {
        if (_ownerThreadId == GetCurrentThreadId() && s_activeBlocker == this)
        {
            s_activeBlocker = nullptr;
        }
        _hook.reset();
        _ownerThreadId = 0u;
    }

    [[nodiscard]] bool IsActive() const noexcept
    {
        return static_cast<bool>(_hook) && s_activeBlocker == this;
    }

    [[nodiscard]] uint64_t BlockedActivationCount() const noexcept
    {
        return _blockedActivationCount;
    }

    [[nodiscard]] uint64_t BlockedFocusCount() const noexcept
    {
        return _blockedFocusCount;
    }

    [[nodiscard]] uint64_t ProtectedTopLevelWindowCount() const noexcept
    {
        return _protectedTopLevelWindowCount;
    }

    [[nodiscard]] static bool IsActiveForCurrentThread() noexcept
    {
        return s_activeBlocker != nullptr && s_activeBlocker->_ownerThreadId == GetCurrentThreadId();
    }

private:
    static LRESULT CALLBACK HookProc(int code, WPARAM wParam, LPARAM lParam) noexcept
    {
        ScopedWindowActivationBlocker* const blocker = s_activeBlocker;
        if (code < 0 || blocker == nullptr)
        {
            return CallNextHookEx(nullptr, code, wParam, lParam);
        }

        if (code == HCBT_CREATEWND)
        {
            auto* const create = reinterpret_cast<CBT_CREATEWNDW*>(lParam);
            if (create != nullptr && create->lpcs != nullptr && (create->lpcs->style & WS_CHILD) == 0)
            {
                create->lpcs->dwExStyle |= WS_EX_NOACTIVATE;
                const HWND window = reinterpret_cast<HWND>(wParam);
                if (window != nullptr)
                {
                    const LONG_PTR exStyle = GetWindowLongPtrW(window, GWL_EXSTYLE);
                    static_cast<void>(SetWindowLongPtrW(window, GWL_EXSTYLE, exStyle | WS_EX_NOACTIVATE));
                }
                ++blocker->_protectedTopLevelWindowCount;
            }
            return CallNextHookEx(nullptr, code, wParam, lParam);
        }

        if (code == HCBT_ACTIVATE)
        {
            ++blocker->_blockedActivationCount;
            static_cast<void>(CallNextHookEx(nullptr, code, wParam, lParam));
            return 1;
        }

        if (code == HCBT_SETFOCUS)
        {
            ++blocker->_blockedFocusCount;
            static_cast<void>(CallNextHookEx(nullptr, code, wParam, lParam));
            return 1;
        }

        return CallNextHookEx(nullptr, code, wParam, lParam);
    }

    inline static thread_local ScopedWindowActivationBlocker* s_activeBlocker = nullptr;

    wil::unique_hhook _hook;
    DWORD _ownerThreadId                   = 0u;
    uint64_t _blockedActivationCount       = 0u;
    uint64_t _blockedFocusCount            = 0u;
    uint64_t _protectedTopLevelWindowCount = 0u;
};
} // namespace RedSalamander::TestSupport
