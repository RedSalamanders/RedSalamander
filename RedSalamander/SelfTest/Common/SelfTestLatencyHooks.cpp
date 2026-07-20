#include "SelfTestLatencyHooks.h"

#ifdef ENABLE_TESTS

#include <algorithm>
#include <array>
#include <atomic>
#include <thread>

namespace
{
constexpr size_t kPointCount          = 5;
constexpr int64_t kMaxSelfTestDelayMs = 5000;
constexpr std::chrono::milliseconds kSlice{5};

std::array<std::atomic<int64_t>, kPointCount> g_nextDelayMs{};
std::array<std::atomic<HRESULT>, kPointCount> g_nextFailureHr{};
std::array<std::atomic<uint64_t>, kPointCount> g_consumeCounts{};

[[nodiscard]] size_t PointIndex(SelfTestLatency::Point point) noexcept
{
    switch (point)
    {
        case SelfTestLatency::Point::ShellThumbnailProviderAllowed: return 0;
        case SelfTestLatency::Point::IconExtractSystemIcon: return 1;
        case SelfTestLatency::Point::IconPathLiveLookup: return 2;
        case SelfTestLatency::Point::PasteShortcutSave: return 3;
        case SelfTestLatency::Point::PasteShortcutAfterSlotProbe: return 4;
    }

    return 0;
}
} // namespace

namespace SelfTestLatency
{
void SetNextDelay(Point point, std::chrono::milliseconds delay) noexcept
{
    const int64_t clamped = std::clamp<int64_t>(delay.count(), 0, kMaxSelfTestDelayMs);
    g_nextDelayMs[PointIndex(point)].store(clamped, std::memory_order_release);
}

void SetNextFailure(Point point, HRESULT hr) noexcept
{
    g_nextFailureHr[PointIndex(point)].store(FAILED(hr) ? hr : S_OK, std::memory_order_release);
}

void ClearAll() noexcept
{
    for (std::atomic<int64_t>& delay : g_nextDelayMs)
    {
        delay.store(0, std::memory_order_release);
    }

    for (std::atomic<HRESULT>& failure : g_nextFailureHr)
    {
        failure.store(S_OK, std::memory_order_release);
    }

    for (std::atomic<uint64_t>& count : g_consumeCounts)
    {
        count.store(0, std::memory_order_release);
    }
}

void Consume(Point point, std::stop_token stopToken) noexcept
{
    const size_t index = PointIndex(point);
    g_consumeCounts[index].fetch_add(1, std::memory_order_acq_rel);

    const int64_t delayMs = g_nextDelayMs[index].exchange(0, std::memory_order_acq_rel);
    if (delayMs <= 0)
    {
        return;
    }

    const auto deadline = std::chrono::steady_clock::now() + std::chrono::milliseconds{delayMs};
    while (! stopToken.stop_requested())
    {
        const auto now = std::chrono::steady_clock::now();
        if (now >= deadline)
        {
            break;
        }

        const auto remaining = std::chrono::duration_cast<std::chrono::milliseconds>(deadline - now);
        std::this_thread::sleep_for((std::min)(remaining, kSlice));
    }
}

HRESULT ConsumeFailure(Point point) noexcept
{
    return g_nextFailureHr[PointIndex(point)].exchange(S_OK, std::memory_order_acq_rel);
}

uint64_t ConsumeCount(Point point) noexcept
{
    return g_consumeCounts[PointIndex(point)].load(std::memory_order_acquire);
}
} // namespace SelfTestLatency

#else

namespace SelfTestLatency
{
void SetNextDelay(Point /*point*/, std::chrono::milliseconds /*delay*/) noexcept {}
void SetNextFailure(Point /*point*/, HRESULT /*hr*/) noexcept {}
void ClearAll() noexcept {}
void Consume(Point /*point*/, std::stop_token /*stopToken*/) noexcept {}
HRESULT ConsumeFailure(Point /*point*/) noexcept
{
    return S_OK;
}
uint64_t ConsumeCount(Point /*point*/) noexcept
{
    return 0;
}
} // namespace SelfTestLatency

#endif
