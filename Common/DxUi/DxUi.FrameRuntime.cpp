#include "DxUi.FrameRuntime.h"

#include "Helpers.h"

#include <limits>

#include <windows.h>

namespace RedSalamander::DxUi
{
namespace
{
constexpr uint64_t kMicrosecondsPerSecond = 1'000'000u;

thread_local FrameStage g_currentDebugFrameStage = FrameStage::Idle;

[[nodiscard]] uint64_t ScaleQuotientFloor(uint64_t numerator, uint64_t multiplier, uint64_t denominator) noexcept
{
    if (denominator == 0u || multiplier == 0u)
    {
        return 0u;
    }

    uint64_t quotient         = 0u;
    uint64_t partialNumerator = 0u;
    uint64_t addQuotient      = numerator / denominator;
    uint64_t addNumerator     = numerator % denominator;

    while (multiplier != 0u)
    {
        if ((multiplier & 1u) != 0u)
        {
            if (addQuotient > std::numeric_limits<uint64_t>::max() - quotient)
            {
                return std::numeric_limits<uint64_t>::max();
            }
            quotient += addQuotient;

            if (addNumerator != 0u)
            {
                if (partialNumerator >= denominator - addNumerator)
                {
                    if (quotient == std::numeric_limits<uint64_t>::max())
                    {
                        return std::numeric_limits<uint64_t>::max();
                    }
                    ++quotient;
                    partialNumerator -= denominator - addNumerator;
                }
                else
                {
                    partialNumerator += addNumerator;
                }
            }
        }

        multiplier >>= 1u;
        if (multiplier == 0u)
        {
            break;
        }

        const bool numeratorCarry = addNumerator >= denominator - addNumerator;
        addNumerator              = numeratorCarry ? addNumerator - (denominator - addNumerator) : addNumerator + addNumerator;
        if (addQuotient > (std::numeric_limits<uint64_t>::max() - (numeratorCarry ? 1u : 0u)) / 2u)
        {
            return std::numeric_limits<uint64_t>::max();
        }
        addQuotient = (addQuotient * 2u) + (numeratorCarry ? 1u : 0u);
    }

    return quotient;
}
} // namespace

FrameClock::FrameClock() noexcept
{
    LARGE_INTEGER frequency{};
    if (QueryPerformanceFrequency(&frequency) == 0 || frequency.QuadPart <= 0)
    {
        _frequency = 0;
        return;
    }

    _frequency = frequency.QuadPart;
}

FrameTimestamp FrameClock::Now() const noexcept
{
    LARGE_INTEGER counter{};
    if (QueryPerformanceCounter(&counter) == 0)
    {
        return {};
    }

    return FrameTimestamp{counter.QuadPart};
}

uint64_t FrameClock::ElapsedUs(FrameTimestamp start, FrameTimestamp end) const noexcept
{
    if (_frequency <= 0 || end.qpc < start.qpc)
    {
        return 0;
    }

    const uint64_t deltaQpc  = static_cast<uint64_t>(end.qpc) - static_cast<uint64_t>(start.qpc);
    const uint64_t frequency = static_cast<uint64_t>(_frequency);

    const uint64_t wholeSeconds = deltaQpc / frequency;
    if (wholeSeconds > std::numeric_limits<uint64_t>::max() / kMicrosecondsPerSecond)
    {
        return std::numeric_limits<uint64_t>::max();
    }

    const uint64_t wholeUs     = wholeSeconds * kMicrosecondsPerSecond;
    const uint64_t remainderUs = ScaleQuotientFloor(deltaQpc % frequency, kMicrosecondsPerSecond, frequency);
    if (remainderUs > std::numeric_limits<uint64_t>::max() - wholeUs)
    {
        return std::numeric_limits<uint64_t>::max();
    }

    return wholeUs + remainderUs;
}

uint64_t FrameClock::SmoothDeltaUs(uint64_t rawDeltaUs, const FrameBudget& budget) noexcept
{
    uint64_t smoothedDeltaUs = rawDeltaUs;
    if (budget.hitchClampUs != 0u && smoothedDeltaUs > budget.hitchClampUs)
    {
        smoothedDeltaUs = budget.hitchClampUs;
    }

    _lastSmoothedDeltaUs = smoothedDeltaUs;
    return smoothedDeltaUs;
}

FrameStageScope::FrameStageScope(FrameStage& currentStage, FrameStage nextStage) noexcept
    : _currentStage(currentStage),
      _previousStage(currentStage),
      _previousDebugStage(g_currentDebugFrameStage)
{
    _currentStage            = nextStage;
    g_currentDebugFrameStage = nextStage;
}

FrameStageScope::~FrameStageScope() noexcept
{
    _currentStage            = _previousStage;
    g_currentDebugFrameStage = _previousDebugStage;
}

bool MotionPolicy::ShouldAnimate() const noexcept
{
    return ! reducedMotion;
}

float MotionPolicy::ResolveProgress(float animatedProgress, float targetProgress) const noexcept
{
    return reducedMotion ? targetProgress : animatedProgress;
}

void EmitFrameMetric(std::wstring_view metric, uint64_t valueUs) noexcept
{
    Debug::Perf::EmitDurationUs(metric, valueUs);
}

bool IsDxUiRenderStageActiveForDebug() noexcept
{
#if defined(_DEBUG) || defined(ENABLE_TESTS)
    return g_currentDebugFrameStage == FrameStage::Render;
#else
    return false;
#endif
}

void EmitDxUiRenderMutationBlockedForDebug() noexcept
{
    Debug::Perf::EmitCounter(L"dxui.frame.render_layout_mutation_blocked");
}
} // namespace RedSalamander::DxUi
