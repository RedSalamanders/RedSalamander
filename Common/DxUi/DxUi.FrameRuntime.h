#pragma once

#include <cstdint>
#include <string_view>

namespace RedSalamander::DxUi
{
enum class FrameStage : uint8_t
{
    Idle,
    Input,
    Update,
    Layout,
    Render,
    Present,
};

enum class FrameMode : uint8_t
{
    Idle,
    EventDriven,
    ActiveAnimation,
    Occluded,
};

struct FrameBudget
{
    uint64_t refreshPeriodUs = 16667;
    uint64_t hitchClampUs    = 50000;
};

struct FrameTimestamp
{
    int64_t qpc = 0;
};

class FrameClock final
{
public:
    FrameClock() noexcept;
    [[nodiscard]] FrameTimestamp Now() const noexcept;
    [[nodiscard]] uint64_t ElapsedUs(FrameTimestamp start, FrameTimestamp end) const noexcept;
    [[nodiscard]] uint64_t SmoothDeltaUs(uint64_t rawDeltaUs, const FrameBudget& budget) noexcept;

private:
    int64_t _frequency            = 1;
    uint64_t _lastSmoothedDeltaUs = 16667;
};

class FrameStageScope final
{
public:
    FrameStageScope(FrameStage& currentStage, FrameStage nextStage) noexcept;
    ~FrameStageScope() noexcept;
    FrameStageScope(const FrameStageScope&)            = delete;
    FrameStageScope& operator=(const FrameStageScope&) = delete;

private:
    FrameStage& _currentStage;
    FrameStage _previousStage      = FrameStage::Idle;
    FrameStage _previousDebugStage = FrameStage::Idle;
};

struct MotionPolicy
{
    bool reducedMotion = false;
    [[nodiscard]] bool ShouldAnimate() const noexcept;
    [[nodiscard]] float ResolveProgress(float animatedProgress, float targetProgress) const noexcept;
};

void EmitFrameMetric(std::wstring_view metric, uint64_t valueUs) noexcept;
} // namespace RedSalamander::DxUi
