#pragma once

#include <cstdint>
#include <memory>
#include <optional>
#include <vector>

#include "AlertOverlay.h"

namespace RedSalamander::Ui
{
class AlertOverlayUiaProvider;

#if defined(ENABLE_TESTS)
struct AlertOverlayWindowDebugSnapshot
{
    bool visible                     = false;
    bool hasLayout                   = false;
    bool hasBackdropBitmap           = false;
    bool usesCompositionPresentation = false;
    bool hasPreparedFrame            = false;
    uint64_t paintCount              = 0;
    HRESULT lastEndDrawResult        = E_PENDING;
    uint32_t renderTargetWindowState = 0u;
    float lastDrawOpacity            = 0.0f;
    float lastDrawScrimOpacity       = 0.0f;
    float minimumDrawOpacity         = 1.0f;
    uint64_t backdropCaptureCount    = 0;
    SIZE clientSizePx{};
    SIZE backdropSizePx{};
    RECT closeRectPx{};
    bool usesSharedCloseChrome  = false;
    bool usesSharedButtonChrome = false;
    uint64_t mouseDownCount     = 0;
    uint64_t mouseUpCount       = 0;
    uint64_t dismissCount       = 0;
    POINT lastMouseDownPointPx{};
    POINT lastMouseUpPointPx{};
    int lastMouseDownHitPart                   = -1;
    int lastMouseUpHitPart                     = -1;
    AlertPresentation presentation             = AlertPresentation::Severity;
    AlertPresentation renderedIconPresentation = AlertPresentation::Severity;
    wchar_t renderedIconGlyph                  = L'\0';
    wchar_t renderedCloseGlyph                 = L'\0';
    std::wstring title;
    std::wstring primaryButtonLabel;
    std::vector<uint64_t> optionValues;
    float scrollOffsetDip  = 0.0f;
    float scrollMaximumDip = 0.0f;
    D2D1_RECT_F bodyViewport{};
    AlertHitTest focusedControl{};
    std::vector<uint32_t> optionIds;
    std::vector<D2D1_RECT_F> visibleOptionRects;
};

struct AlertOverlayIdentitySelfTestResult
{
    bool staleRootRejected                             = false;
    bool staleButtonRejected                           = false;
    bool staleOptionRejected                           = false;
    bool staleCloseRejected                            = false;
    bool queuedButtonIgnoredAfterReplacement           = false;
    bool queuedOptionIgnoredAfterReplacement           = false;
    bool queuedCloseIgnoredAfterReplacement            = false;
    bool currentButtonInvoked                          = false;
    bool currentOptionInvoked                          = false;
    bool currentProviderSurvivesSamePresentationUpdate = false;
    bool runtimeIdChanged                              = false;
    bool reusedForeignWindowRejected                   = false;
    bool destroyedProviderRejectedAfterReplacement     = false;
    bool drainedCommandIgnoredAfterReplacement         = false;
};

struct AlertOverlayTrackingSelfTestResult
{
    bool siblingDetachKeepsTracking      = false;
    bool laterSubclassKept               = false;
    bool originalProcedureRestored       = false;
    bool ownerDestroyedClearsVisibility  = false;
    bool closeHidesWithoutEscapeBinding  = false;
    bool targetDestroyedCallbackReported = false;
    uint64_t foregroundRequests          = 0u;
};

struct AlertOverlayUxAccessibilitySelfTestResult
{
    // Option activation in a scrolling alert (pointer/UIA, arrow, Space) keeps the content in place.
    bool optionFixtureScrolls         = false;
    bool pointerActivationKeepsScroll = false;
    bool arrowActivationKeepsScroll   = false;
    bool spaceActivationKeepsScroll   = false;
    float optionScrollBeforeDip       = 0.0f;
    float optionScrollAfterDip        = 0.0f;
    // The prompt icon sits in the visible body of a scrolling alert.
    bool iconFixtureScrolls    = false;
    bool iconInsideVisibleBody = false;
    D2D1_RECT_F iconRect{};
    D2D1_RECT_F bodyViewport{};
    // UI Automation events for focus moves inside the alert and for a replaced model.
    bool tabRaisesFocusEvent              = false;
    bool arrowRaisesFocusEvent            = false;
    bool uiaSetFocusRaisesFocusEvent      = false;
    bool optionCycleRaisesNameChange      = false;
    bool unfocusedMoveRaisesNoFocusEvent  = false;
    bool replacementRaisesStructureChange = false;
    bool replacementAnnouncesText         = false;
    bool replacementRaisesFocusEvent      = false;
    // A re-shown blocking alert keeps the focus it captured; the overlay is never that focus.
    bool reshowKeepsRestoreFocus            = false;
    bool overlayNeverRecordedAsRestoreFocus = false;
    // Alt mnemonics come only from '&' markers.
    bool unmarkedLabelsHaveNoMnemonic   = false;
    bool markedMnemonicInvokesItsButton = false;
    bool ambiguousMnemonicIgnored       = false;
    bool markersHiddenFromName          = false;
    // UIA names of an untitled alert and of its close button, in the current UI language.
    std::wstring untitledRootName;
    std::wstring closeName;
    // The UI Automation marshal: a message carrying a raw call pointer runs nothing, a provider called
    // from another thread still reaches the UI thread, and an untaken synchronous payload is reclaimed.
    bool forgedUiaCallIgnored        = false;
    bool crossThreadUiaCallMarshals  = false;
    bool untakenSyncPayloadReclaimed = false;
    uint64_t foregroundRequests      = 0u;
};

// I35 R8.4/R8.12/R8.13: the modeless decision overlay (the File Operations decision mirror).
struct AlertOverlayModelessSelfTestResult
{
    // Shown while the anchor had the focus: takes it, and gives it back to that window when it hides.
    bool takeFocusCapturesAnchor = false;
    bool takeFocusMovesFocus     = false;
    bool hideRestoresFocus       = false;
    // Not shown with the focus: an Escape routed to every visible alert does nothing, and Enter is inert.
    bool unfocusedEscapeIgnored = false;
    bool unfocusedEnterIgnored  = false;
    // With the focus, Escape runs only the escape binding; the close glyph runs onClose, not the binding.
    bool focusedEscapeRunsBinding   = false;
    bool closeRunsOnCloseNotBinding = false;
    // A first show announces the given text through a UI Automation notification.
    bool firstShowAnnounces = false;
    // The overlay follows a move of the anchor's top-level window, and hides with a hidden anchor.
    bool followsRootMove       = false;
    bool hidesWithHiddenAnchor = false;
    bool showsWithShownAnchor  = false;
    // The anchor (a pane-like child window) is destroyed: the overlay hides and its owner is told.
    bool destroyedAnchorHides    = false;
    bool destroyedAnchorReported = false;
    // A focus the overlay took by itself is guarded (I35 S5 red-team): Tab, arrows, Space and mnemonics are
    // inert and the focus stays on the default; an auto-repeated key, a key typed within the typeahead
    // window and a key still held from before the show act on nothing; Enter submits only the primary
    // binding and Escape only the escape binding; a click on the overlay lifts the guard.
    bool guardKeepsFocusOnDefault   = false;
    bool guardIgnoresRepeat         = false;
    bool guardIgnoresTypeahead      = false;
    bool guardIgnoresHeldKey        = false;
    bool guardEnterSubmitsDefault   = false;
    bool clickLiftsGuard            = false;
    bool unguardedTabReachesClose   = false;
    bool repeatIgnoredWhenUnguarded = false;
    // Losing the focus to another window forgets the focus-return target; its own owned menu does not.
    bool focusLossForgetsReturn = false;
    bool ownedMenuKeepsReturn   = false;
    // An option change reaches the owner; the owner's value can be set without a new presentation.
    bool optionChangeReported  = false;
    bool optionValueSetInPlace = false;
    // A blocking alert whose focus-return target was hidden gives the focus to its anchor instead: the
    // helper, and the blocking hide that uses it.
    bool hiddenReturnFallsBackToAnchor = false;
    bool blockingHideFallsBackToAnchor = false;
    // Alt+F4 (WM_CLOSE) on a modeless overlay closes its owner's window, never the overlay window.
    bool closeMessageForwardedToOwner = false;
    // The keyboard way out of the guard (I35 S5 second red-team): F6 or Shift+F6 hands the focus back to the
    // pane without answering, and Ctrl+Tab lifts the guard with the focus ring on the default; both act
    // only fresh (an auto-repeat, a key typed within the typeahead window and a key, the chord's Ctrl
    // included, held since the guard was armed act on nothing).
    bool f6IgnoresRepeatTypeaheadAndHeld      = false;
    bool f6ReturnsFocusToPane                 = false;
    bool shiftF6ReturnsFocusToPane            = false;
    bool ctrlTabIgnoresRepeatTypeaheadAndHeld = false;
    bool ctrlTabLiftsGuardToDefault           = false;
    bool tabAfterCtrlTabLeavesDefault         = false;
    // A focus the overlay gets later without a click or a UI Automation request arms the guard again; a
    // focus coming back from its own menu keeps the guard as it was.
    bool focusGainArmsGuard             = false;
    bool ownedMenuFocusReturnKeepsGuard = false;
    // An overlay that does not require the focus for keys (a non-modal plugin alert) keeps acting on a
    // repeated key and lets Alt+F4 close its own window, as before the decision overlays.
    bool plainOverlayKeepsRepeatAndClose = false;
    // A move of the tracked top-level window keeps a modeless overlay's place in the z-order and raises a
    // blocking one.
    bool trackingMoveKeepsModelessZOrder = false;
    bool trackingMoveRaisesBlocking      = false;
    uint64_t foregroundRequests          = 0u;
};

[[nodiscard]] bool DebugGetAlertOverlayWindowSnapshot(HWND hwnd, AlertOverlayWindowDebugSnapshot& out) noexcept;
// Diagnostic full-frame replay only; documentation still requires the presented-window capture.
[[nodiscard]] bool DebugCaptureAlertOverlayWindowBitmap(HWND hwnd, DxUi::WindowHostBitmapCapture& out) noexcept;
[[nodiscard]] bool DebugRunAlertOverlayIdentitySelfTest(HWND parent, AlertOverlayIdentitySelfTestResult& out) noexcept;
// Overlays sharing one tracked window, a later foreign subclass, an owner destroyed without the
// tracking hook, and modeless hides (which must never request the foreground). Creates and destroys
// its own unowned anchor window placed at `placement`'s origin.
[[nodiscard]] bool DebugRunAlertOverlayTrackingSelfTest(HWND placement, AlertOverlayTrackingSelfTestResult& out) noexcept;
// Option activation and icon placement in a scrolling alert, UIA focus/structure/notification events,
// the focus a re-shown blocking alert restores, '&' mnemonics and the localized UIA names. Stands in
// the thread focus and captures UIA events, so it needs no activation; creates its own anchor window.
[[nodiscard]] bool DebugRunAlertOverlayUxAccessibilitySelfTest(HWND placement, AlertOverlayUxAccessibilitySelfTestResult& out) noexcept;
// The modeless decision overlay: focus capture and return, keys only while focused, the key guard and its
// keyboard way out (F6, Ctrl+Tab), focus gains, close versus Escape, first-show announcement, root and
// visibility tracking with the z-order kept, a destroyed child anchor, and the blocking-hide focus fallback.
// Stands in the thread focus, the foreground window and the key message time and captures UIA events, so it
// needs no activation; creates its own windows near `placement`.
[[nodiscard]] bool DebugRunAlertOverlayModelessSelfTest(HWND placement, AlertOverlayModelessSelfTestResult& out) noexcept;
// Process-wide count of the foreground requests made when an alert hides.
[[nodiscard]] uint64_t DebugGetAlertOverlayForegroundRequestCount() noexcept;
// Stands in the calling thread's focus window for the alert overlays (and for owners that probe the focus
// through AlertOverlayFocusWindow) until it is cleared with nullopt. A modeless overlay that takes or gives
// back the focus then moves this stand-in instead of the real focus.
void DebugSetAlertOverlayFocusOverride(std::optional<HWND> focus) noexcept;
// Stands in the message time (GetMessageTime) the key guard of a modeless overlay sees for the key messages
// the calling thread delivers, until it is cleared with nullopt; a self-test sends keys with SendMessage,
// which carries no message time of its own.
void DebugSetAlertOverlayKeyMessageTimeOverride(std::optional<DWORD> messageTime) noexcept;
// Stands in the foreground window a hiding blocking alert checks before it hands the focus back, until it
// is cleared with nullopt.
void DebugSetAlertOverlayForegroundOverride(std::optional<HWND> foreground) noexcept;
// The real input routes of an overlay reached through its window: a button press, the close affordance
// and an option activation, exactly as a click or UI Automation Invoke dispatches them.
[[nodiscard]] bool DebugInvokeAlertOverlayButton(HWND hwnd, uint32_t buttonId) noexcept;
[[nodiscard]] bool DebugInvokeAlertOverlayClose(HWND hwnd) noexcept;
[[nodiscard]] bool DebugInvokeAlertOverlayOption(HWND hwnd, uint32_t optionId) noexcept;
// A modeless overlay's model, key bindings, focus-return target and the text its last first show
// announced (empty when none).
struct AlertOverlayModelessDebugSnapshot
{
    bool visible          = false;
    bool windowVisible    = false;
    bool keysRequireFocus = false;
    bool restoresFocus    = false;
    HWND restoreFocus     = nullptr;
    HWND anchor           = nullptr;
    std::wstring title;
    std::wstring message;
    std::wstring lastAnnouncement;
    std::wstring rootUiaName;
    std::vector<uint32_t> buttonIds;
    std::vector<std::wstring> buttonLabels;
    std::optional<uint32_t> primaryButtonId;
    std::optional<uint32_t> escapeButtonId;
    std::optional<uint32_t> focusedButtonId;
    bool focusedClose = false;
    std::vector<uint32_t> optionIds;
    std::vector<uint64_t> optionValues;
    // The key guard: whether the overlay guards every focus it gets without a deliberate interaction, whether
    // the guard is armed now, and the tick of the show or focus gain that armed it.
    bool guardsKeyboardFocus = false;
    bool keysGuarded         = false;
    DWORD guardArmedTick     = 0u;
    // Advances with every presentation (a show or a re-show of a new model), not with an in-place update.
    uint64_t presentationGeneration = 0u;
};
[[nodiscard]] bool DebugGetAlertOverlayModelessSnapshot(HWND hwnd, AlertOverlayModelessDebugSnapshot& out) noexcept;
#endif

// The window that has the keyboard focus on the calling thread, as the alert overlays see it (a self-test
// may stand in for it). Owners that decide whether a modeless overlay takes the focus probe it here.
[[nodiscard]] HWND AlertOverlayFocusWindow() noexcept;

struct AlertOverlayWindowCallbacks
{
    void* context                                               = nullptr;
    void (*onButton)(void* context, uint32_t buttonId) noexcept = nullptr;
    // When set, the close affordance (the close glyph, a focused Close with Enter or Space, UI Automation
    // Invoke on Close) runs this instead of the escape-button binding; Escape still runs the binding.
    void (*onClose)(void* context) noexcept = nullptr;
    // When set, runs after the user changed an option (click, keyboard or UI Automation) with its new value,
    // so the owner can keep its own state as the single truth.
    void (*onOptionChanged)(void* context, uint32_t optionId, uint64_t value) noexcept = nullptr;
    // When set, runs after F6 or Shift+F6 handed the keyboard focus back to the window the overlay covers
    // (nothing was answered; the overlay stays shown).
    void (*onFocusReturned)(void* context) noexcept = nullptr;
};

// Keys typed this soon after a guarded modeless overlay got the focus were meant for the window the focus
// came from (typeahead); the guard drops them.
inline constexpr DWORD kModelessKeyGuardTypeaheadMs = 500u;

// A modeless overlay that hosts a decision (the File Operations decision mirror, I35 R8.4/R8.12/R8.13).
// The options apply to one Show; Hide clears them together with the callbacks and the key bindings.
struct AlertModelessShowOptions
{
    // Take the keyboard focus when shown (the owner decided the anchor had it) and give it back to the
    // window that had it when the overlay hides while it still has the focus. Never requests the foreground.
    // Losing the focus to a window outside the overlay (other than its own menu) forgets that return target.
    bool takeFocus = false;
    // Guard every keyboard focus the overlay gets without a deliberate interaction: the focus takeFocus takes
    // at the show and a focus handed to it later (an activation that returns to it, a prompt that gives the
    // focus back); not a click on it or a UI Automation focus request, and not a focus coming back from its
    // own menu. While guarded, keys the user meant for the pane cannot reach a destructive button: Tab,
    // arrows, Space, paging keys and mnemonics are inert and the focus ring stays on the default; only
    // Enter (the primary binding), Escape (the escape binding), F6 or Shift+F6 (the focus back to the pane,
    // as on an unguarded overlay) and Ctrl+Tab (lifts the guard and puts the focus ring on the default)
    // act, and only fresh: not auto-repeated, not typed before or within kModelessKeyGuardTypeaheadMs of
    // the arming, and not held (Ctrl of the chord included) since the arming. A click on the overlay or a
    // UI Automation focus request lifts the guard as well.
    bool guardKeyboardFocus = false;
    // Keyboard input, including an Escape routed to every visible alert, acts only while the overlay
    // itself has the keyboard focus. Such a decision overlay also ignores every auto-repeated key, hands
    // the focus back to the window it covers on F6 or Shift+F6 (onFocusReturned), reaches its close glyph
    // with Tab and passes Alt+F4 (WM_CLOSE) to its owner's top-level window.
    bool keysRequireFocus = false;
    // The overlay window is hidden while its anchor is hidden and shown again with it.
    bool followAnchorVisibility = false;
    // Raised as a UI Automation notification once shown (an unfocused first show raises no system focus
    // event that would announce it). Empty: nothing is announced.
    std::wstring announcement;
};

class AlertOverlayWindow final
{
    friend class AlertOverlayUiaProvider;
#if defined(ENABLE_TESTS)
    friend bool DebugGetAlertOverlayWindowSnapshot(HWND hwnd, AlertOverlayWindowDebugSnapshot& out) noexcept;
    friend bool DebugCaptureAlertOverlayWindowBitmap(HWND hwnd, DxUi::WindowHostBitmapCapture& out) noexcept;
    friend bool DebugRunAlertOverlayIdentitySelfTest(HWND parent, AlertOverlayIdentitySelfTestResult& out) noexcept;
    friend bool DebugRunAlertOverlayTrackingSelfTest(HWND placement, AlertOverlayTrackingSelfTestResult& out) noexcept;
    friend bool DebugRunAlertOverlayUxAccessibilitySelfTest(HWND placement, AlertOverlayUxAccessibilitySelfTestResult& out) noexcept;
    friend bool DebugRunAlertOverlayModelessSelfTest(HWND placement, AlertOverlayModelessSelfTestResult& out) noexcept;
    friend bool DebugInvokeAlertOverlayButton(HWND hwnd, uint32_t buttonId) noexcept;
    friend bool DebugInvokeAlertOverlayClose(HWND hwnd) noexcept;
    friend bool DebugInvokeAlertOverlayOption(HWND hwnd, uint32_t optionId) noexcept;
    friend bool DebugGetAlertOverlayModelessSnapshot(HWND hwnd, AlertOverlayModelessDebugSnapshot& out) noexcept;
#endif

public:
    AlertOverlayWindow() noexcept;

    AlertOverlayWindow(const AlertOverlayWindow&)            = delete;
    AlertOverlayWindow(AlertOverlayWindow&&)                 = delete;
    AlertOverlayWindow& operator=(const AlertOverlayWindow&) = delete;
    AlertOverlayWindow& operator=(AlertOverlayWindow&&)      = delete;

    ~AlertOverlayWindow();

    HRESULT ShowForParentClient(HWND parent, const AlertTheme& theme, AlertModel model, bool blocksInput) noexcept;
    HRESULT ShowForAnchor(HWND anchor, const AlertTheme& theme, AlertModel model, bool blocksInput) noexcept;
    // A modeless (non-blocking) overlay over `anchor` that hosts a decision; see AlertModelessShowOptions.
    // Showing a new model while visible keeps the focus-return target the first show captured.
    HRESULT ShowModelessForAnchor(HWND anchor, const AlertTheme& theme, AlertModel model, AlertModelessShowOptions options) noexcept;

    void Hide() noexcept;
    // Hides without handing the focus or the activation back (the owner is being torn down).
    void HideWithoutRestoringFocus() noexcept;

    [[nodiscard]] HWND GetHwnd() const noexcept
    {
        return _hwnd.get();
    }

    // Whether the overlay window has the keyboard focus, as the overlay sees it (AlertOverlayFocusWindow).
    [[nodiscard]] bool HasKeyboardFocus() const noexcept;
    // The screen rectangle of a visible button, for anchoring a menu it opens; nullopt when not visible.
    [[nodiscard]] std::optional<RECT> GetButtonScreenRect(uint32_t buttonId) noexcept;

    void SetCallbacks(AlertOverlayWindowCallbacks callbacks) noexcept;
    void ClearCallbacks() noexcept;

    // Owner notification that survives Hide: the window this overlay tracks (its anchor or host
    // parent) is being destroyed. It is the last thing that path does, so the owner may delete the
    // overlay inside it.
    void SetTrackedWindowDestroyedCallback(void* context, void (*callback)(void* context, AlertOverlayWindow& overlay) noexcept) noexcept;

    void SetKeyBindings(std::optional<uint32_t> primaryButtonId, std::optional<uint32_t> escapeButtonId) noexcept;

    [[nodiscard]] bool IsVisible() const noexcept
    {
        return _visible;
    }

    [[nodiscard]] std::optional<uint64_t> GetOptionValue(uint32_t optionId) const noexcept
    {
        return _overlay.GetOptionValue(optionId);
    }

    // Shows the owner's current value of an option without a new presentation (no focus or text change).
    void SetOptionValue(uint32_t optionId, uint64_t value) noexcept;

private:
    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnPaint() noexcept;
    void OnSize(UINT width, UINT height) noexcept;
    void OnMouseMove(POINT pt) noexcept;
    void OnMouseWheel(int delta, POINT screenPoint) noexcept;
    LRESULT OnFocusControl(AlertHitTest::Part part, uint32_t id) noexcept;
    void OnMouseLeave() noexcept;
    void OnLButtonDown(POINT pt) noexcept;
    void OnLButtonUp(POINT pt) noexcept;
    void OnCaptureChanged(HWND newCapture) noexcept;
    void OnKeyDown(WPARAM key, LPARAM flags) noexcept;
    void OnKeyUp(WPARAM key) noexcept;
    void OnSetFocus(HWND previousFocus) noexcept;
    void OnKillFocus(HWND newFocus) noexcept;
    // Alt+F4 or another close request: a decision overlay passes it to its owner's top-level window.
    [[nodiscard]] bool OnCloseRequest() noexcept;
    [[nodiscard]] bool OnSysChar(WPARAM key) noexcept;
    LRESULT OnSetCursor(HWND cursorWindow, UINT hitTest, UINT mouseMsg) noexcept;

    enum class DismissSource : uint8_t
    {
        Escape, // the Escape key: the escape binding wins
        Close,  // the close glyph, a focused Close, UI Automation Invoke on Close: onClose wins
    };

    void InvokeButton(uint32_t buttonId) noexcept;
    void InvokeOption(uint32_t optionId, bool reverse = false) noexcept;
    void InvokeDismiss(DismissSource source = DismissSource::Close) noexcept;
    void OnInvokeMessage(UINT message, LPARAM parameter) noexcept;
    // Keyboard input acts unless the overlay requires the focus for it and does not have it.
    [[nodiscard]] bool AcceptsKeyboardInput() const noexcept;
    // The key guard (AlertModelessShowOptions::guardKeyboardFocus): arms it from now with the keys held now.
    void ArmKeyGuard() noexcept;
    // Whether this key-down is fresh enough to act on a guarded overlay (not within the typeahead window, not
    // held since the arming; `controlChord`: the Ctrl of the chord too); the caller limits the keys.
    [[nodiscard]] bool PassesKeyGuard(WPARAM key, bool controlChord = false) const noexcept;
    // The user interacted with the overlay deliberately (a click, a UI Automation focus request, Ctrl+Tab).
    void ReleaseKeyGuard() noexcept;
    // Ctrl+Tab on a guarded overlay: lifts the guard and puts the focus ring on the default.
    void LiftKeyGuardToDefault() noexcept;
    // Puts the focus ring on the primary binding's button (a guarded overlay's Enter submits that one).
    void FocusDefaultButton() noexcept;
    // F6 on a decision overlay: the focus goes back to the window it covers (the control that had it while
    // that is still inside the anchor and visible, else the anchor), without answering.
    void ReturnFocusToAnchor() noexcept;
    // The window a hide hands the focus back to: `target` while it is visible, otherwise this alert's own
    // anchor (or host parent) when that is visible; nullptr when neither can take it.
    [[nodiscard]] HWND ResolveFocusReturnTarget(HWND target) const noexcept;
    // Moves the keyboard focus (a self-test's stand-in focus when one is active).
    static void MoveFocusTo(HWND hwnd) noexcept;
    // UI Automation events; raised on the UI thread once the overlay state is consistent.
    bool RaiseUiaFocusChanged() noexcept;
    void RaiseUiaOptionValueChanged(uint32_t optionId) noexcept;
    void RaiseUiaPresentationReplaced(bool textChanged) noexcept;
    void RaiseUiaAnnouncement(std::wstring_view text) noexcept;
    void AdvancePresentationGeneration() noexcept;
    [[nodiscard]] bool IsCurrentPresentation(const std::weak_ptr<const uint64_t>& instanceToken, uint64_t generation) const noexcept;

    void StartAnimationTimer() noexcept;
    void StopAnimationTimer() noexcept;
    static bool AnimationTickThunk(void* context, uint64_t nowTickMs) noexcept;
    bool OnAnimationTimer(uint64_t nowTickMs) noexcept;

    HRESULT TransitionVisibility(
        bool show, const AlertTheme* theme, AlertModel* model, bool blocksInput, AlertModelessShowOptions* modeless = nullptr) noexcept;
    // A modeless overlay that follows its anchor's visibility hides or shows its window with the anchor.
    void ApplyAnchorVisibility() noexcept;
    HRESULT EnsureCreated(HWND hostParent) noexcept;
    void Destroy() noexcept;

    void ApplyAttachmentState(HWND hostParent, HWND anchor, bool trackHostParent, bool trackAnchor) noexcept;
    void AttachToParentClient(HWND parent) noexcept;
    void AttachToAnchor(HWND anchor) noexcept;

    // keepZOrder: a tracking move (the anchor or its top-level window moved) never raises a non-blocking
    // overlay; a blocking one is raised again as on every placement.
    void UpdatePlacement(bool keepZOrder = false) noexcept;
    [[nodiscard]] bool EnsureOverlayLayoutForInput() noexcept;

    static LRESULT CALLBACK TrackedWindowProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    void OnTrackedWindowMessage(HWND hwnd, UINT msg, bool anchorRole) noexcept;

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
    const std::shared_ptr<const uint64_t> _instanceToken;
    uint64_t _presentationGeneration = 0u;
    HWND _hostParent                 = nullptr;
    HWND _anchor                     = nullptr;
    // The anchor's top-level window when it is neither the anchor nor its parent: an owned popup does
    // not move with its owner, and a move of the top-level window reaches neither the anchor nor its parent.
    HWND _root = nullptr;

    bool _visible            = false;
    bool _blocksInput        = true;
    bool _trackingMouseLeave = false;
    bool _alwaysAnimate      = false;

    uint64_t _animationSubscriptionId = 0;

    // Registration ids in the tracked windows' shared hook; zero when not tracking.
    uint64_t _hostParentTrackingId = 0;
    uint64_t _anchorTrackingId     = 0;
    uint64_t _rootTrackingId       = 0;
    uint64_t _animateUntilTickMs   = 0;
    uint64_t _startTickMs          = 0;

    HWND _restoreFocus = nullptr;
    // Modeless decision overlay (AlertModelessShowOptions); all cleared on Hide.
    bool _modelessRestoresFocus      = false;
    bool _keysRequireFocus           = false;
    bool _followAnchorVisibility     = false;
    bool _suppressFocusRestoreOnHide = false;
    // The key guard (AlertModelessShowOptions::guardKeyboardFocus): whether this presentation guards its focus
    // gains, whether the guard is armed, the tick it counts the typeahead window from, and the guard keys
    // (Enter, Escape, F6, Tab, Ctrl) that were down at the arming and have not been released since.
    bool _guardKeyboardFocus = false;
    bool _keysGuarded        = false;
    DWORD _guardArmedTick    = 0u;
    uint8_t _guardHeldKeys   = 0u;
#if defined(ENABLE_TESTS)
    std::wstring _debugLastAnnouncement;
#endif

    UINT _dpi = 96;
    SIZE _clientSizePx{};

    // Presentation only: the existing alert owns input, focus and UI Automation.
    DxUi::WindowHost _presentationHost;
    wil::com_ptr<ID2D1DeviceContext> _target;
    wil::com_ptr<ID2D1CommandList> _preparedFrame;
    wil::com_ptr<ID2D1Bitmap> _backdropBitmap;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;

    bool _scrollDragging = false;
    AlertOverlay _overlay;
    AlertHitTest _pressedHit{};

    std::optional<RECT> _panelRegionPx;

    AlertOverlayWindowCallbacks _callbacks{};
    void* _trackedWindowDestroyedContext                                                   = nullptr;
    void (*_onTrackedWindowDestroyed)(void* context, AlertOverlayWindow& overlay) noexcept = nullptr;
    std::optional<uint32_t> _primaryButtonId;
    std::optional<uint32_t> _escapeButtonId;
#if defined(ENABLE_TESTS)
    uint64_t _debugPaintCount           = 0;
    HRESULT _debugLastEndDrawResult     = E_PENDING;
    float _debugMinimumDrawOpacity      = 1.0f;
    uint64_t _debugBackdropCaptureCount = 0;
    SIZE _debugBackdropSizePx{};
    uint64_t _debugMouseDownCount = 0;
    uint64_t _debugMouseUpCount   = 0;
    uint64_t _debugDismissCount   = 0;
    POINT _debugLastMouseDownPointPx{};
    POINT _debugLastMouseUpPointPx{};
    int _debugLastMouseDownHitPart = -1;
    int _debugLastMouseUpHitPart   = -1;
#endif
};
} // namespace RedSalamander::Ui
