#pragma once

#include <algorithm>
#include <cstdint>
#include <limits>
#include <vector>

#pragma warning(push)
// Windows headers: C4710 (not inlined), C4711 (auto inline), C4514 (unreferenced inline)
#pragma warning(disable : 4710 4711 4514)
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#pragma warning(pop)

#include "DxUi/DxUi.FrameRuntime.h"
#include "Helpers.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace RedSalamander::Ui
{
class AnimationDispatcher final
{
private:
    struct TickTiming
    {
        uint64_t rawDeltaUs      = 0u;
        uint64_t callbackDeltaUs = 0u;
        uint64_t legacyGapMs     = 0u;
        bool legacyOverrun       = false;
    };

public:
    using TickCallback = bool (*)(void* context, uint64_t nowTickMs) noexcept;

    AnimationDispatcher() noexcept = default;

    AnimationDispatcher(const AnimationDispatcher&)            = delete;
    AnimationDispatcher(AnimationDispatcher&&)                 = delete;
    AnimationDispatcher& operator=(const AnimationDispatcher&) = delete;
    AnimationDispatcher& operator=(AnimationDispatcher&&)      = delete;

    [[nodiscard]] static AnimationDispatcher& GetInstance() noexcept
    {
        // Intentionally leaked to avoid shutdown UAF from static destruction order issues.
        static AnimationDispatcher* instance = new AnimationDispatcher();
        return *instance;
    }

    void Shutdown() noexcept
    {
        StopTimer();
        _subscriptions.clear();
        _pendingAdds.clear();
        _inTick             = false;
        _timerRunning       = false;
        _nextSubscriptionId = 1;
        _lastTickMs         = 0u;
        _tickEpochMs        = 0u;
        _elapsedTickUs      = 0u;
        _lastTickTimestamp  = {};
        _hasTickTimestamp   = false;
        _hwnd.reset();
    }

    [[nodiscard]] uint64_t Subscribe(TickCallback callback, void* context) noexcept
    {
        if (! callback)
        {
            return 0;
        }

        EnsureWindow();
        if (! _hwnd)
        {
            return 0;
        }

        Subscription entry{};
        entry.id            = _nextSubscriptionId;
        entry.callback      = callback;
        entry.context       = context;
        entry.pendingRemove = false;

        _nextSubscriptionId += 1u;

        if (_inTick)
        {
            _pendingAdds.push_back(entry);
        }
        else
        {
            _subscriptions.push_back(entry);
        }

        EnsureTimerRunning();
        return entry.id;
    }

    void Unsubscribe(uint64_t id) noexcept
    {
        if (id == 0)
        {
            return;
        }

        MarkPendingRemove(_subscriptions, id);
        MarkPendingRemove(_pendingAdds, id);

        if (! _inTick)
        {
            GarbageCollect();
            EnsureTimerState();
        }
    }

#if defined(ENABLE_TESTS)
    struct DebugTimingForTest
    {
        uint64_t callbackDeltaUs = 0u;
        uint64_t legacyGapMs     = 0u;
        bool legacyOverrun       = false;
    };

    [[nodiscard]] uint64_t DebugGetTargetFrameUsForTest() const noexcept
    {
        return _frameBudget.refreshPeriodUs;
    }

    [[nodiscard]] uint64_t DebugGetHitchClampUsForTest() const noexcept
    {
        return _frameBudget.hitchClampUs;
    }

    [[nodiscard]] DebugTimingForTest DebugComputeTimingForTest(uint64_t rawDeltaUs) noexcept
    {
        const TickTiming timing = ComputeTickTiming(rawDeltaUs);
        return DebugTimingForTest{
            timing.callbackDeltaUs,
            timing.legacyGapMs,
            timing.legacyOverrun,
        };
    }
#endif

private:
    static constexpr wchar_t kWindowClassName[] = L"RedSalamander.AnimationDispatcher";
    static constexpr UINT_PTR kTimerId          = 1;
    static constexpr UINT kFrameIntervalMs      = 8u;
    static constexpr uint64_t kTargetFrameUs    = 1'000'000u / 120u;
    static constexpr uint64_t kHitchClampUs     = 50'000u;

    struct Subscription
    {
        uint64_t id           = 0;
        TickCallback callback = nullptr;
        void* context         = nullptr;
        bool pendingRemove    = false;
    };

    static void MarkPendingRemove(std::vector<Subscription>& list, uint64_t id) noexcept
    {
        for (auto& entry : list)
        {
            if (entry.id == id)
            {
                entry.pendingRemove = true;
                return;
            }
        }
    }

    [[nodiscard]] static uint64_t AbsoluteDifference(uint64_t lhs, uint64_t rhs) noexcept
    {
        return lhs >= rhs ? lhs - rhs : rhs - lhs;
    }

    [[nodiscard]] static size_t CountActiveSubscriptions(const std::vector<Subscription>& list) noexcept
    {
        size_t count = 0u;
        for (const auto& entry : list)
        {
            if (! entry.pendingRemove && entry.callback != nullptr)
            {
                ++count;
            }
        }

        return count;
    }

    [[nodiscard]] uint64_t AddElapsedTickUs(uint64_t deltaUs) noexcept
    {
        if (deltaUs > std::numeric_limits<uint64_t>::max() - _elapsedTickUs)
        {
            _elapsedTickUs = std::numeric_limits<uint64_t>::max();
        }
        else
        {
            _elapsedTickUs += deltaUs;
        }

        const uint64_t elapsedMs = _elapsedTickUs / 1'000u;
        if (elapsedMs > std::numeric_limits<uint64_t>::max() - _tickEpochMs)
        {
            return std::numeric_limits<uint64_t>::max();
        }

        return _tickEpochMs + elapsedMs;
    }

    [[nodiscard]] TickTiming ComputeTickTiming(uint64_t rawDeltaUs) noexcept
    {
        const uint64_t callbackDeltaUs = _clock.SmoothDeltaUs(rawDeltaUs, _frameBudget);
        const uint64_t legacyGapMs     = rawDeltaUs / 1'000u;
        return TickTiming{
            rawDeltaUs,
            callbackDeltaUs,
            legacyGapMs,
            legacyGapMs > 20u,
        };
    }

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
    {
        if (msg == WM_NCCREATE)
        {
            const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
            auto* self     = static_cast<AnimationDispatcher*>(cs->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            return DefWindowProcW(hwnd, msg, wp, lp);
        }

        auto* self = reinterpret_cast<AnimationDispatcher*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (self)
        {
            return self->WndProc(hwnd, msg, wp, lp);
        }

        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
    {
        switch (msg)
        {
            case WM_TIMER:
                if (static_cast<UINT_PTR>(wp) == kTimerId)
                {
                    OnTimerTick();
                    return 0;
                }
                break;
            case WM_NCDESTROY:
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                _timerRunning = false;
                _hwnd.release();
                return DefWindowProcW(hwnd, msg, wp, lp);
            }
        }

        return DefWindowProcW(hwnd, msg, wp, lp);
    }

    void EnsureWindow() noexcept
    {
        if (_hwnd && IsWindow(_hwnd.get()))
        {
            return;
        }

        static ATOM atom = 0;
        if (atom == 0)
        {
            WNDCLASSEXW wc{};
            wc.cbSize        = sizeof(wc);
            wc.lpfnWndProc   = &AnimationDispatcher::WndProcThunk;
            wc.hInstance     = GetModuleHandleW(nullptr);
            wc.lpszClassName = kWindowClassName;
            atom             = RegisterClassExW(&wc);
        }

        if (atom == 0)
        {
            return;
        }

        HWND hwnd = CreateWindowExW(0, kWindowClassName, L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, GetModuleHandleW(nullptr), this);

        if (! hwnd)
        {
            return;
        }

        _hwnd.reset(hwnd);
        _timerRunning = false;
        EnsureTimerState();
    }

    void EnsureTimerRunning() noexcept
    {
        if (_timerRunning)
        {
            return;
        }

        if (! _hwnd)
        {
            return;
        }

        if (_subscriptions.empty() && _pendingAdds.empty())
        {
            return;
        }

        const UINT_PTR timer = SetTimer(_hwnd.get(), kTimerId, kFrameIntervalMs, nullptr);
        if (timer != 0)
        {
            _timerRunning = true;
        }
    }

    void StopTimer() noexcept
    {
        if (! _timerRunning || ! _hwnd)
        {
            _timerRunning      = false;
            _lastTickMs        = 0u;
            _tickEpochMs       = 0u;
            _elapsedTickUs     = 0u;
            _lastTickTimestamp = {};
            _hasTickTimestamp  = false;
            return;
        }

        KillTimer(_hwnd.get(), kTimerId);
        _timerRunning      = false;
        _lastTickMs        = 0u;
        _tickEpochMs       = 0u;
        _elapsedTickUs     = 0u;
        _lastTickTimestamp = {};
        _hasTickTimestamp  = false;
    }

    void EnsureTimerState() noexcept
    {
        if (_subscriptions.empty() && _pendingAdds.empty())
        {
            StopTimer();
            return;
        }

        EnsureTimerRunning();
    }

    void GarbageCollect() noexcept
    {
        std::erase_if(_subscriptions, [](const Subscription& entry) noexcept { return entry.pendingRemove || entry.callback == nullptr; });
        std::erase_if(_pendingAdds, [](const Subscription& entry) noexcept { return entry.pendingRemove || entry.callback == nullptr; });
    }

    void AppendPendingAdds() noexcept
    {
        if (_pendingAdds.empty())
        {
            return;
        }

        _subscriptions.reserve(_subscriptions.size() + _pendingAdds.size());
        for (const auto& entry : _pendingAdds)
        {
            if (! entry.pendingRemove && entry.callback != nullptr)
            {
                _subscriptions.push_back(entry);
            }
        }

        _pendingAdds.clear();
    }

    void OnTimerTick() noexcept
    {
        if (_subscriptions.empty())
        {
            GarbageCollect();
            AppendPendingAdds();
            EnsureTimerState();
            return;
        }

        const DxUi::FrameTimestamp timestamp = _clock.Now();
        const bool hadPreviousTick           = _hasTickTimestamp;
        if (! _hasTickTimestamp)
        {
            _tickEpochMs       = GetTickCount64();
            _lastTickTimestamp = timestamp;
            _hasTickTimestamp  = true;
            _elapsedTickUs     = 0u;
            _lastTickMs        = _tickEpochMs;
        }

        uint64_t rawDeltaUs = kTargetFrameUs;
        if (_lastTickTimestamp.qpc != timestamp.qpc)
        {
            rawDeltaUs = _clock.ElapsedUs(_lastTickTimestamp, timestamp);
        }

        const TickTiming timing = ComputeTickTiming(rawDeltaUs);
        const uint64_t now      = AddElapsedTickUs(timing.callbackDeltaUs);
        Debug::Perf::EmitDurationUs(L"dxui.animation.tick_delta_us", timing.rawDeltaUs);
        Debug::Perf::EmitDurationUs(L"dxui.animation.jitter_us", AbsoluteDifference(timing.rawDeltaUs, _frameBudget.refreshPeriodUs));
        if (hadPreviousTick)
        {
            Debug::Perf::EmitValue(L"dxui.animation.tick_gap_ms", timing.legacyGapMs);
            if (timing.legacyOverrun)
            {
                Debug::Perf::EmitCounter(L"dxui.animation.tick_overrun");
            }
        }
        _lastTickMs        = now;
        _lastTickTimestamp = timestamp;

        _inTick = true;
        for (auto& entry : _subscriptions)
        {
            if (entry.pendingRemove || entry.callback == nullptr)
            {
                continue;
            }

            const bool keep = entry.callback(entry.context, now);
            if (! keep)
            {
                entry.pendingRemove = true;
            }
        }
        _inTick = false;

        GarbageCollect();
        AppendPendingAdds();
        Debug::Perf::EmitValue(L"dxui.animation.active_count", static_cast<uint64_t>(CountActiveSubscriptions(_subscriptions)));
        EnsureTimerState();
    }

private:
    wil::unique_hwnd _hwnd;
    bool _timerRunning           = false;
    bool _inTick                 = false;
    uint64_t _nextSubscriptionId = 1;
    uint64_t _lastTickMs         = 0u;
    uint64_t _tickEpochMs        = 0u;
    uint64_t _elapsedTickUs      = 0u;
    DxUi::FrameClock _clock;
    DxUi::FrameBudget _frameBudget{kTargetFrameUs, kHitchClampUs};
    DxUi::FrameTimestamp _lastTickTimestamp{};
    bool _hasTickTimestamp = false;
    std::vector<Subscription> _subscriptions;
    std::vector<Subscription> _pendingAdds;
};
} // namespace RedSalamander::Ui
