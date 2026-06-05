#pragma once

#include <cstdint>
#include <optional>

#include "AlertOverlay.h"

namespace RedSalamander::Ui
{
class AlertOverlayUiaProvider;

#if defined(ENABLE_TESTS)
struct AlertOverlayWindowDebugSnapshot
{
    bool visible = false;
    bool hasLayout = false;
    bool hasBackdropBitmap = false;
    uint64_t paintCount = 0;
    float lastDrawOpacity = 0.0f;
    float lastDrawScrimOpacity = 0.0f;
    float minimumDrawOpacity = 1.0f;
    uint64_t backdropCaptureCount = 0;
    SIZE clientSizePx{};
    SIZE backdropSizePx{};
    RECT closeRectPx{};
    uint64_t mouseDownCount = 0;
    uint64_t mouseUpCount = 0;
    uint64_t dismissCount = 0;
    POINT lastMouseDownPointPx{};
    POINT lastMouseUpPointPx{};
    int lastMouseDownHitPart = -1;
    int lastMouseUpHitPart = -1;
};

[[nodiscard]] bool DebugGetAlertOverlayWindowSnapshot(HWND hwnd, AlertOverlayWindowDebugSnapshot& out) noexcept;
#endif

struct AlertOverlayWindowCallbacks
{
    void* context                                               = nullptr;
    void (*onButton)(void* context, uint32_t buttonId) noexcept = nullptr;
    void (*onDismissed)(void* context) noexcept                 = nullptr;
};

class AlertOverlayWindow final
{
    friend class AlertOverlayUiaProvider;
#if defined(ENABLE_TESTS)
    friend bool DebugGetAlertOverlayWindowSnapshot(HWND hwnd, AlertOverlayWindowDebugSnapshot& out) noexcept;
#endif

public:
    AlertOverlayWindow() noexcept = default;

    AlertOverlayWindow(const AlertOverlayWindow&)            = delete;
    AlertOverlayWindow(AlertOverlayWindow&&)                 = delete;
    AlertOverlayWindow& operator=(const AlertOverlayWindow&) = delete;
    AlertOverlayWindow& operator=(AlertOverlayWindow&&)      = delete;

    ~AlertOverlayWindow();

    HRESULT ShowForParentClient(HWND parent, const AlertTheme& theme, AlertModel model, bool blocksInput) noexcept;
    HRESULT ShowForAnchor(HWND anchor, const AlertTheme& theme, AlertModel model, bool blocksInput) noexcept;

    void Hide() noexcept;

    void SetCallbacks(AlertOverlayWindowCallbacks callbacks) noexcept;
    void ClearCallbacks() noexcept;

    void SetKeyBindings(std::optional<uint32_t> primaryButtonId, std::optional<uint32_t> escapeButtonId) noexcept;

    [[nodiscard]] bool IsVisible() const noexcept
    {
        return _visible;
    }

private:
    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnPaint() noexcept;
    void OnSize(UINT width, UINT height) noexcept;
    void OnMouseMove(POINT pt) noexcept;
    void OnMouseLeave() noexcept;
    void OnLButtonDown(POINT pt) noexcept;
    void OnLButtonUp(POINT pt) noexcept;
    void OnCaptureChanged(HWND newCapture) noexcept;
    void OnKeyDown(WPARAM key) noexcept;
    [[nodiscard]] bool OnSysChar(WPARAM key) noexcept;
    LRESULT OnSetCursor(HWND cursorWindow, UINT hitTest, UINT mouseMsg) noexcept;

    void InvokeButton(uint32_t buttonId) noexcept;
    void InvokeDismiss() noexcept;

    void StartAnimationTimer() noexcept;
    void StopAnimationTimer() noexcept;
    static bool AnimationTickThunk(void* context, uint64_t nowTickMs) noexcept;
    bool OnAnimationTimer(uint64_t nowTickMs) noexcept;

    HRESULT TransitionVisibility(bool show, const AlertTheme* theme, AlertModel* model, bool blocksInput) noexcept;
    HRESULT EnsureCreated(HWND hostParent) noexcept;
    void Destroy() noexcept;

    void ApplyAttachmentState(HWND hostParent, HWND anchor, bool trackHostParent, bool trackAnchor) noexcept;
    void AttachToParentClient(HWND parent) noexcept;
    void AttachToAnchor(HWND anchor) noexcept;

    void UpdatePlacement() noexcept;
    [[nodiscard]] bool EnsureOverlayLayoutForInput() noexcept;

    static LRESULT CALLBACK ParentWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    static LRESULT CALLBACK AnchorWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void EnsureD2DResources() noexcept;
    void DiscardD2DResources() noexcept;
    void CaptureBackdrop(bool preserveExistingOnFailure = false) noexcept;

    void ApplyRegionFromOverlay() noexcept;
    void ClearRegion() noexcept;

    [[nodiscard]] float DipFromPx(int px) const noexcept;
    [[nodiscard]] int PxFromDipFloor(float dip) const noexcept;
    [[nodiscard]] int PxFromDipCeil(float dip) const noexcept;
    [[nodiscard]] int PxFromDipRound(float dip) const noexcept;

private:
    wil::unique_hwnd _hwnd;
    HWND _hostParent = nullptr;
    HWND _anchor     = nullptr;

    bool _visible            = false;
    bool _blocksInput        = true;
    bool _trackingMouseLeave = false;
    bool _alwaysAnimate      = false;

    uint64_t _animationSubscriptionId = 0;

    bool _hostParentSubclassed   = false;
    bool _anchorSubclassed       = false;
    uint64_t _animateUntilTickMs = 0;
    uint64_t _startTickMs        = 0;

    HWND _restoreFocus = nullptr;

    UINT _dpi = 96;
    SIZE _clientSizePx{};

    wil::com_ptr<ID2D1Factory> _d2dFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> _target;
    wil::com_ptr<ID2D1Bitmap> _backdropBitmap;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;

    AlertOverlay _overlay;
    AlertHitTest _pressedHit{};

    std::optional<RECT> _panelRegionPx;

    AlertOverlayWindowCallbacks _callbacks{};
    std::optional<uint32_t> _primaryButtonId;
    std::optional<uint32_t> _escapeButtonId;
#if defined(ENABLE_TESTS)
    uint64_t _debugPaintCount = 0;
    float _debugMinimumDrawOpacity = 1.0f;
    uint64_t _debugBackdropCaptureCount = 0;
    SIZE _debugBackdropSizePx{};
    uint64_t _debugMouseDownCount = 0;
    uint64_t _debugMouseUpCount = 0;
    uint64_t _debugDismissCount = 0;
    POINT _debugLastMouseDownPointPx{};
    POINT _debugLastMouseUpPointPx{};
    int _debugLastMouseDownHitPart = -1;
    int _debugLastMouseUpHitPart = -1;
#endif
};
} // namespace RedSalamander::Ui
