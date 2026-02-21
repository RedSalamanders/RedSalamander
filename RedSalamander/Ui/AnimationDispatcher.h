#pragma once

#include <algorithm>
#include <cstdint>
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

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace RedSalamander::Ui
{
class AnimationDispatcher final
{
public:
    using TickCallback = bool (*)(void* context, uint64_t nowTickMs) noexcept;

    AnimationDispatcher() noexcept = default;

    AnimationDispatcher(const AnimationDispatcher&)            = delete;
    AnimationDispatcher(AnimationDispatcher&&)                 = delete;
    AnimationDispatcher& operator=(const AnimationDispatcher&) = delete;
    AnimationDispatcher& operator=(AnimationDispatcher&&)      = delete;

    [[nodiscard]] static AnimationDispatcher& GetInstance() noexcept
    {
        static AnimationDispatcher* instance = new AnimationDispatcher();
        return *instance;
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

private:
    static constexpr wchar_t kWindowClassName[] = L"RedSalamander.AnimationDispatcher";
    static constexpr UINT_PTR kTimerId          = 1;
    static constexpr UINT kFrameIntervalMs      = 16u;

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
            _timerRunning = false;
            return;
        }

        KillTimer(_hwnd.get(), kTimerId);
        _timerRunning = false;
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

        const uint64_t now = GetTickCount64();

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
        EnsureTimerState();
    }

private:
    wil::unique_hwnd _hwnd;
    bool _timerRunning           = false;
    bool _inTick                 = false;
    uint64_t _nextSubscriptionId = 1;
    std::vector<Subscription> _subscriptions;
    std::vector<Subscription> _pendingAdds;
};
} // namespace RedSalamander::Ui
