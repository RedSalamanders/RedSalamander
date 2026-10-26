#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <chrono>
#include <cstdint>

#include "Common.Tests.WindowActivationGuard.h"
#include "TestSupport.Tests.Runner.h"

namespace RedSalamander::TestSupport
{
// Read-only desktop evidence. Last-input ticks are diagnostic, never proof of
// which process produced input (injected input can also change them).
struct DesktopStateSnapshot
{
    HWND foregroundWindow = nullptr;
    POINT cursor{};
    DWORD foregroundProcessId = 0u;
    DWORD foregroundThreadId  = 0u;
    DWORD lastInputTick       = 0u;
    BOOL lastInputAvailable   = FALSE;
};

[[nodiscard]] inline bool CaptureDesktopState(void*, DesktopStateSnapshot& state) noexcept
{
    state                  = {};
    state.foregroundWindow = GetForegroundWindow();
    if (state.foregroundWindow)
    {
        state.foregroundThreadId = GetWindowThreadProcessId(state.foregroundWindow, &state.foregroundProcessId);
        if (! state.foregroundThreadId)
        {
            return false;
        }
    }
    if (GetCursorPos(&state.cursor) == FALSE)
    {
        return false;
    }
    LASTINPUTINFO input{sizeof(input)};
    state.lastInputAvailable = GetLastInputInfo(&input);
    state.lastInputTick      = input.dwTime;
    return true;
}

// Tests may supply observations without changing the physical desktop.
struct DesktopSnapshotSource
{
    bool (*capture)(void*, DesktopStateSnapshot&) noexcept = CaptureDesktopState;
    void* context                                          = nullptr;
};

enum class DesktopObservationFailure
{
    None,
    GuardUnavailable,
    ObservationUnavailable,
    OwnedForeground
};

struct DesktopForegroundEvents
{
    uint64_t transitions = 0u;

    void Record(DWORD event) noexcept
    {
        if (event == EVENT_SYSTEM_FOREGROUND)
        {
            ++transitions;
        }
    }
};

[[nodiscard]] inline DesktopObservationFailure EvaluateDesktopObservation(
    bool guardActive, bool observationAvailable, DWORD ownerProcessId, DWORD foregroundProcessId, const DesktopForegroundEvents& events) noexcept
{
    if (! guardActive)
    {
        return DesktopObservationFailure::GuardUnavailable;
    }
    if (! observationAvailable || ! ownerProcessId)
    {
        return DesktopObservationFailure::ObservationUnavailable;
    }
    if (foregroundProcessId == ownerProcessId || events.transitions)
    {
        return DesktopObservationFailure::OwnedForeground;
    }
    return DesktopObservationFailure::None;
}

// Test-only, UI-thread-affine evidence for no-activation UIA work. Foreign
// desktop activity is observed, never restored or attributed to this process.
// A process-filtered WinEvent stream retains transient owned activation even
// when the final foreground window belongs to another application.
class ScopedNonActivatingDesktopObserver final
{
public:
    struct Statistics
    {
        uint64_t samples           = 0u;
        uint64_t cursorChanges     = 0u;
        uint64_t foregroundChanges = 0u;
        uint64_t inputChanges      = 0u;
        uint64_t barrierUs         = 0u;
    };

    ScopedNonActivatingDesktopObserver()                                                     = default;
    ScopedNonActivatingDesktopObserver(const ScopedNonActivatingDesktopObserver&)            = delete;
    ScopedNonActivatingDesktopObserver& operator=(const ScopedNonActivatingDesktopObserver&) = delete;
    ScopedNonActivatingDesktopObserver(ScopedNonActivatingDesktopObserver&&)                 = delete;
    ScopedNonActivatingDesktopObserver& operator=(ScopedNonActivatingDesktopObserver&&)      = delete;

    ~ScopedNonActivatingDesktopObserver() noexcept
    {
        Stop();
    }

    [[nodiscard]] bool Start(DesktopSnapshotSource source = {}) noexcept
    {
        if (_hook || s_active || ! ScopedWindowActivationBlocker::IsActiveForCurrentThread())
        {
            _failure = DesktopObservationFailure::GuardUnavailable;
            return false;
        }
        _source                 = source;
        _statistics             = {};
        _events                 = {};
        _barrierRequests        = 0u;
        _failure                = DesktopObservationFailure::None;
        _ownerThread            = GetCurrentThreadId();
        _startTick              = GetTickCount();
        _lastForegroundWindow   = nullptr;
        _lastForegroundThread   = 0u;
        _lastForegroundTick     = 0u;
        _lastForegroundClass[0] = L'\0';
        _sentinel.reset(CreateWindowExW(WS_EX_NOACTIVATE, L"STATIC", L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, GetModuleHandleW(nullptr), nullptr));
        if (! _sentinel)
        {
            _failure = DesktopObservationFailure::ObservationUnavailable;
            return false;
        }
        // Name-change on our hidden message-only window is a FIFO drain marker.
        // It is an accessibility notification, never a foreground/input write.
        _hook.reset(SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_OBJECT_NAMECHANGE, nullptr, OnWinEvent, GetCurrentProcessId(), 0, WINEVENT_OUTOFCONTEXT));
        if (! _hook)
        {
            _failure = DesktopObservationFailure::ObservationUnavailable;
            _sentinel.reset();
            return false;
        }
        s_active = this;
        if (! Sample())
        {
            Stop();
            return false;
        }
        return true;
    }

    void Stop() noexcept
    {
        if (s_active == this)
        {
            s_active = nullptr;
        }
        _hook.reset();
        _sentinel.reset();
        _ownerThread = 0u;
    }

    [[nodiscard]] bool Sample() noexcept
    {
        DesktopStateSnapshot current{};
        const bool guardActive = ScopedWindowActivationBlocker::IsActiveForCurrentThread();
        const bool available =
            _hook && s_active == this && _ownerThread == GetCurrentThreadId() && _source.capture && _source.capture(_source.context, current);
        _lastObservation          = current;
        _lastObservationAvailable = available;
        _lastObservationError     = available ? ERROR_SUCCESS : GetLastError();
        const auto failure        = EvaluateDesktopObservation(guardActive, available, GetCurrentProcessId(), current.foregroundProcessId, _events);
        if (_failure == DesktopObservationFailure::None)
        {
            _failure = failure;
        }
        if (_failure != DesktopObservationFailure::None)
        {
            return false;
        }
        if (_statistics.samples)
        {
            _statistics.cursorChanges += current.cursor.x != _previous.cursor.x || current.cursor.y != _previous.cursor.y ? 1u : 0u;
            _statistics.foregroundChanges +=
                current.foregroundWindow != _previous.foregroundWindow || current.foregroundProcessId != _previous.foregroundProcessId ? 1u : 0u;
            _statistics.inputChanges +=
                current.lastInputAvailable && _previous.lastInputAvailable && current.lastInputTick != _previous.lastInputTick ? 1u : 0u;
        }
        _previous = current;
        ++_statistics.samples;
        return true;
    }

    [[nodiscard]] bool Finish() noexcept
    {
        if (! Sample())
        {
            return false;
        }
        _barrierSeen = false;
        ++_barrierRequests;
        const auto started = std::chrono::steady_clock::now();
        NotifyWinEvent(EVENT_OBJECT_NAMECHANGE, _sentinel.get(), OBJID_WINDOW, CHILDID_SELF);
        MessagePumpWaitOptions options;
        options.timeout       = std::chrono::milliseconds(250);
        options.pollInterval  = std::chrono::milliseconds(1);
        options.operationName = L"Nonactivating desktop observation event drain";
        const auto drained    = PumpMessagesUntil([&]() noexcept { return _barrierSeen != FALSE; }, options);
        _statistics.barrierUs =
            static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count());
        if (! drained.conditionMet)
        {
            _failure = DesktopObservationFailure::ObservationUnavailable;
            return false;
        }
        return Sample();
    }

    [[nodiscard]] DesktopObservationFailure Failure() const noexcept
    {
        return _failure;
    }
    [[nodiscard]] const Statistics& Stats() const noexcept
    {
        return _statistics;
    }
    [[nodiscard]] uint64_t OwnedForegroundTransitions() const noexcept
    {
        return _events.transitions;
    }
    // Borrowed identities for failure diagnostics only. The HWND may already be
    // destroyed when an asynchronous WinEvent is delivered; never act on it.
    [[nodiscard]] HWND LastOwnedForegroundWindow() const noexcept
    {
        return _lastForegroundWindow;
    }
    [[nodiscard]] DWORD LastOwnedForegroundThread() const noexcept
    {
        return _lastForegroundThread;
    }
    [[nodiscard]] DWORD LastOwnedForegroundTick() const noexcept
    {
        return _lastForegroundTick;
    }
    [[nodiscard]] DWORD StartTick() const noexcept
    {
        return _startTick;
    }
    [[nodiscard]] const wchar_t* LastOwnedForegroundClass() const noexcept
    {
        return _lastForegroundClass;
    }
    [[nodiscard]] const DesktopStateSnapshot& LastObservation() const noexcept
    {
        return _lastObservation;
    }
    [[nodiscard]] bool LastObservationAvailable() const noexcept
    {
        return _lastObservationAvailable;
    }
    [[nodiscard]] DWORD LastObservationError() const noexcept
    {
        return _lastObservationError;
    }
    [[nodiscard]] DWORD BarrierCount() const noexcept
    {
        return _barrierRequests;
    }

private:
    static void CALLBACK OnWinEvent(HWINEVENTHOOK hook, DWORD event, HWND window, LONG object, LONG child, DWORD eventThread, DWORD eventTick) noexcept
    {
        auto* observer = s_active;
        if (! observer || hook != observer->_hook.get())
        {
            return;
        }
        observer->_events.Record(event);
        if (event == EVENT_SYSTEM_FOREGROUND)
        {
            observer->_lastForegroundWindow   = window;
            observer->_lastForegroundThread   = eventThread;
            observer->_lastForegroundTick     = eventTick;
            observer->_lastForegroundClass[0] = L'\0';
            static_cast<void>(
                GetClassNameW(window, observer->_lastForegroundClass, static_cast<int>(sizeof(observer->_lastForegroundClass) / sizeof(wchar_t))));
        }
        if (event == EVENT_OBJECT_NAMECHANGE && window == observer->_sentinel.get() && object == OBJID_WINDOW && child == CHILDID_SELF)
        {
            observer->_barrierSeen = true;
        }
    }

    inline static thread_local ScopedNonActivatingDesktopObserver* s_active = nullptr;
    using UniqueWinEventHook                                                = wil::unique_any<HWINEVENTHOOK, decltype(&::UnhookWinEvent), ::UnhookWinEvent>;
    DesktopSnapshotSource _source;
    DesktopStateSnapshot _previous;
    DesktopStateSnapshot _lastObservation;
    bool _lastObservationAvailable = false;
    DWORD _lastObservationError    = ERROR_SUCCESS;
    DesktopForegroundEvents _events;
    Statistics _statistics;
    wil::unique_hwnd _sentinel;
    UniqueWinEventHook _hook;
    DWORD _ownerThread                 = 0u;
    DesktopObservationFailure _failure = DesktopObservationFailure::None;
    DWORD _barrierRequests             = 0u;
    BOOL _barrierSeen                  = FALSE;
    HWND _lastForegroundWindow         = nullptr;
    DWORD _lastForegroundThread        = 0u;
    DWORD _lastForegroundTick          = 0u;
    DWORD _startTick                   = 0u;
    wchar_t _lastForegroundClass[128]{};
};
} // namespace RedSalamander::TestSupport
