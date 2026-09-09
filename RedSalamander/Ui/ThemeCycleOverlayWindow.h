#pragma once

#include <chrono>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <span>
#include <string>

#include "DxUi/DxUi.h"

namespace RedSalamander::Ui
{
enum class ThemeCycleDirection : int8_t
{
    Previous = -1,
    Direct   = 0,
    Next     = 1,
};

enum class CommandInvocationSource : uint8_t
{
    KeyboardShortcut,
    ThemeMenu,
    FunctionBarPointer,
    OtherWmCommand,
    Programmatic,
    SelfTest,
};

enum class ThemeCycleOverlayPhase : uint8_t
{
    Hidden,
    Appearing,
    ContentTransition,
    FullyVisible,
    Disappearing,
};

inline constexpr uint64_t kThemeCycleAppearDurationMs             = 140u;
inline constexpr uint64_t kThemeCycleAdjacentTransitionDurationMs = 160u;
inline constexpr uint64_t kThemeCycleDirectTransitionDurationMs   = 140u;
inline constexpr uint64_t kThemeCycleDisappearDelayMs             = 900u;
inline constexpr uint64_t kThemeCycleDisappearDurationMs          = 160u;
inline constexpr uint64_t kThemeCycleExitCancelRecoveryMs         = 80u;

struct ThemeCycleOverlaySnapshot final
{
    uint64_t generation = 0u;
    ThemeCycleDirection direction = ThemeCycleDirection::Next;
    std::wstring previousThemeId;
    std::wstring previousDisplayName;
    std::wstring currentThemeId;
    std::wstring currentDisplayName;
    std::wstring nextThemeId;
    std::wstring nextDisplayName;
    std::chrono::steady_clock::time_point inputAcceptedAt{};
};

struct ThemeCycleOverlayTheme final
{
    std::wstring themeId;
    std::wstring displayName;
};

struct ThemeCycleOverlayPlacement final
{
    bool visible = false;
    RECT windowRectPx{};
};

[[nodiscard]] ThemeCycleOverlayPlacement ComputeThemeCycleOverlayPlacement(
    const RECT& ownerClientScreenRectPx,
    const RECT& monitorWorkAreaPx,
    UINT dpi) noexcept;

[[nodiscard]] ThemeCycleOverlaySnapshot BuildThemeCycleOverlaySnapshot(
    std::span<const ThemeCycleOverlayTheme> ring,
    size_t selectedIndex,
    ThemeCycleDirection direction,
    uint64_t generation,
    std::chrono::steady_clock::time_point inputAcceptedAt = {});

struct ThemeCycleOverlayAccessibility final
{
    std::wstring notification;
    std::wstring previousName;
    std::wstring currentName;
    std::wstring nextName;
    std::wstring dismissAction;
    std::wstring dismissHelp;
};

struct ThemeCycleOverlayDebugSnapshot final
{
    bool created                    = false;
    bool visible                    = false;
    bool dismissalTimerArmed       = false;
    bool fallbackDeadlineArmed     = false;
    bool animationActive           = false;
    bool pressed                   = false;
    bool captured                  = false;
    bool reducedMotion             = false;
    bool highContrast              = false;
    bool backdropCaptured          = false;
    ThemeCycleOverlayPhase phase   = ThemeCycleOverlayPhase::Hidden;
    ThemeCycleDirection direction  = ThemeCycleDirection::Next;
    uint64_t generation            = 0u;
    uint64_t phaseStartTickMs      = 0u;
    uint64_t steadyVisibleTickMs   = 0u;
    uint64_t disappearDeadlineMs   = 0u;
    float progress                 = 1.0f;
    float surfaceOpacity           = 0.0f;
    float surfaceScale             = 1.0f;
    float surfaceTranslateYDip     = 0.0f;
    RECT windowRectPx{};
    RECT surfaceRectPx{};
    D2D1_RECT_F previousRectDip{};
    D2D1_RECT_F currentRectDip{};
    D2D1_RECT_F nextRectDip{};
    UINT dpi                       = 96u;
    std::wstring previousThemeId;
    std::wstring previousDisplayName;
    std::wstring currentThemeId;
    std::wstring currentDisplayName;
    std::wstring nextThemeId;
    std::wstring nextDisplayName;
    uint64_t paintCount            = 0u;
    uint64_t textLayoutBuildCount  = 0u;
    uint64_t textLayoutCreateCount = 0u;
    uint64_t explicitDismissCount  = 0u;
    uint64_t windowCreateCount     = 0u;
    uint64_t windowReuseCount      = 0u;
    uint64_t dismissTimerArmCount  = 0u;
    uint64_t staleTimerIgnoredCount = 0u;
    uint64_t droppedGenerationCount = 0u;
    uint64_t backdropCaptureCount   = 0u;
    uint64_t backdropCaptureFailureCount = 0u;
    UINT backdropWidthPx            = 0u;
    UINT backdropHeightPx           = 0u;
};

class ThemeCycleOverlayWindow final
{
public:
    ThemeCycleOverlayWindow() noexcept;
    ~ThemeCycleOverlayWindow();

    ThemeCycleOverlayWindow(const ThemeCycleOverlayWindow&)            = delete;
    ThemeCycleOverlayWindow(ThemeCycleOverlayWindow&&)                 = delete;
    ThemeCycleOverlayWindow& operator=(const ThemeCycleOverlayWindow&) = delete;
    ThemeCycleOverlayWindow& operator=(ThemeCycleOverlayWindow&&)      = delete;

    HRESULT Show(HWND owner,
                 const DxUi::ThemePalette& palette,
                 ThemeCycleOverlaySnapshot snapshot,
                 ThemeCycleOverlayAccessibility accessibility) noexcept;
    void Hide() noexcept;
    void OnOwnerGeometryChanged() noexcept;
    void OnOwnerAvailabilityChanged() noexcept;
    [[nodiscard]] bool IsVisible() const noexcept;
    [[nodiscard]] HWND GetHwnd() const noexcept;

#if defined(ENABLE_TESTS)
    [[nodiscard]] ThemeCycleOverlayDebugSnapshot DebugGetSnapshot() const noexcept;
    [[nodiscard]] IRawElementProviderFragmentRoot* DebugCreateAccessibilityProvider() const noexcept;
    [[nodiscard]] bool DebugCaptureBitmap(DxUi::WindowHostBitmapCapture& capture) noexcept;
    void DebugAdvanceTo(uint64_t nowTickMs) noexcept;
    void DebugDismissImmediately() noexcept;
    void DebugSetForceDismissTimerFailure(bool force) noexcept;
    void DebugSimulateDeviceLoss() noexcept;
#endif

private:
    class Impl;
    std::unique_ptr<Impl> _impl;
};
} // namespace RedSalamander::Ui
