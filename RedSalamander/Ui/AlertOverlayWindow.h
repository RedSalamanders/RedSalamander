#pragma once

#include <cstdint>
#include <optional>

#include "AlertOverlay.h"

#pragma warning(push)
// Windows headers: C4710 (not inlined), C4711 (auto inline), C4514 (unreferenced inline)
#pragma warning(disable : 4710 4711 4514)
#include <commctrl.h>
#pragma warning(pop)

namespace RedSalamander::Ui
{
struct AlertOverlayWindowCallbacks
{
    void* context                                               = nullptr;
    void (*onButton)(void* context, uint32_t buttonId) noexcept = nullptr;
    void (*onDismissed)(void* context) noexcept                 = nullptr;
};

class AlertOverlayWindow final
{
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
    void OnKeyDown(WPARAM key) noexcept;
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

    static LRESULT CALLBACK ParentSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR id, DWORD_PTR refData) noexcept;
    static LRESULT CALLBACK AnchorSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR id, DWORD_PTR refData) noexcept;

    void EnsureD2DResources() noexcept;
    void DiscardD2DResources() noexcept;

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

    bool _hostParentSubclassed = false;
    bool _anchorSubclassed     = false;
    UINT_PTR _subclassId       = 0;

    uint64_t _animateUntilTickMs = 0;
    uint64_t _startTickMs        = 0;

    HWND _restoreFocus = nullptr;

    UINT _dpi = 96;
    SIZE _clientSizePx{};

    wil::com_ptr<ID2D1Factory> _d2dFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> _target;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;

    AlertOverlay _overlay;

    std::optional<RECT> _panelRegionPx;

    AlertOverlayWindowCallbacks _callbacks{};
    std::optional<uint32_t> _primaryButtonId;
    std::optional<uint32_t> _escapeButtonId;
};
} // namespace RedSalamander::Ui
