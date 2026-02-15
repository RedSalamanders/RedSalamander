#include "StartupMetrics.h"

#include "Helpers.h"

#include <atomic>
#include <chrono>
#include <mutex>

namespace StartupMetrics
{
namespace
{
std::once_flag g_once;
std::chrono::steady_clock::time_point g_start;

std::atomic_flag g_firstWindowEmitted        = ATOMIC_FLAG_INIT;
std::atomic_flag g_firstPaintEmitted         = ATOMIC_FLAG_INIT;
std::atomic_flag g_inputReadyEmitted         = ATOMIC_FLAG_INIT;
std::atomic_flag g_firstPanePopulatedEmitted = ATOMIC_FLAG_INIT;

[[nodiscard]] uint64_t ElapsedUs() noexcept
{
    const auto now     = std::chrono::steady_clock::now();
    const auto elapsed = now - g_start;
    return static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count());
}

void EmitOnce(std::atomic_flag& flag, std::wstring_view name, std::wstring_view detail, uint64_t value0 = 0, uint64_t value1 = 0, HRESULT hr = S_OK) noexcept
{
    if (flag.test_and_set(std::memory_order_acq_rel))
    {
        return;
    }

    Debug::Perf::Emit(name, detail, ElapsedUs(), value0, value1, hr);
}
} // namespace

void Initialize() noexcept
{
    std::call_once(g_once, [] { g_start = std::chrono::steady_clock::now(); });
}

void MarkFirstWindowCreated(std::wstring_view windowId) noexcept
{
    Initialize();
    EmitOnce(g_firstWindowEmitted, L"App.Startup.Metric.TimeToFirstWindow", windowId);
}

void MarkFirstPaint(std::wstring_view windowId) noexcept
{
    Initialize();
    EmitOnce(g_firstPaintEmitted, L"App.Startup.Metric.TimeToFirstPaint", windowId);
}

void MarkInputReady(std::wstring_view windowId) noexcept
{
    Initialize();
    EmitOnce(g_inputReadyEmitted, L"App.Startup.Metric.TimeToInputReady", windowId);
}

void MarkFirstPanePopulated(std::wstring_view detail, uint64_t itemCount) noexcept
{
    Initialize();
    EmitOnce(g_firstPanePopulatedEmitted, L"App.Startup.Metric.TimeToFirstPanePopulated", detail, itemCount, 0);
}
} // namespace StartupMetrics
