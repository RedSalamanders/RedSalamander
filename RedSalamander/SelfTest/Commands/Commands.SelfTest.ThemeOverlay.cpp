// Commands.SelfTest.ThemeOverlay.cpp
// Included from Commands.SelfTest.cpp — NOT compiled standalone.
// Theme-cycle overlay behavior, lifetime, accessibility, and performance coverage.

namespace
{
[[nodiscard]] uint8_t GetThemeOverlayCaptureAlpha(const DxUi::WindowHostBitmapCapture& capture, LONG x, LONG y) noexcept
{
    if (x < 0 || y < 0 || static_cast<UINT>(x) >= capture.widthPx || static_cast<UINT>(y) >= capture.heightPx)
    {
        return 0u;
    }
    const size_t offset = ((static_cast<size_t>(y) * static_cast<size_t>(capture.widthPx)) + static_cast<size_t>(x)) * 4u + 3u;
    return offset < capture.bgraPixels.size() ? capture.bgraPixels[offset] : 0u;
}

[[nodiscard]] bool WaitForThemeOverlayHidden(std::chrono::milliseconds timeout) noexcept
{
    const auto deadline = std::chrono::steady_clock::now() + timeout;
    while (std::chrono::steady_clock::now() < deadline)
    {
        PumpPendingMessages();
        if (! DebugGetThemeCycleOverlaySnapshot().visible)
        {
            return true;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    }
    PumpPendingMessages();
    return ! DebugGetThemeCycleOverlaySnapshot().visible;
}

void SetThemeWithoutOverlay(HWND mainWindow, std::wstring_view themeId, CaseState& state) noexcept
{
    const std::wstring commandId = std::format(L"cmd/app/theme/select/{0}", themeId);
    state.Require(DebugDispatchShortcutCommand(mainWindow, commandId), L"Direct test theme setup command failed.");
    DebugHideThemeCycleOverlay();
    PumpPendingMessages();
}

[[nodiscard]] bool TestThemeCycleOverlayPureStateAndGeometry(HWND, CaseState& state) noexcept
{
    using namespace RedSalamander::Ui;

    const ThemeCycleOverlaySnapshot empty = BuildThemeCycleOverlaySnapshot({}, 0u, ThemeCycleDirection::Direct, 41u);
    state.Require(empty.generation == 41u && empty.direction == ThemeCycleDirection::Direct && empty.currentThemeId.empty(),
                  L"An empty ring should retain generation/direction and expose no theme strings.");

    std::array<ThemeCycleOverlayTheme, 1u> one{{ThemeCycleOverlayTheme{L"user/one", L"One"}}};
    const ThemeCycleOverlaySnapshot only = BuildThemeCycleOverlaySnapshot(one, 99u, ThemeCycleDirection::Next, 42u);
    state.Require(only.currentThemeId == L"user/one" && only.currentDisplayName == L"One", L"A one-entry ring should clamp selection to its only entry.");
    state.Require(only.previousThemeId.empty() && only.nextThemeId.empty(), L"A one-entry ring should omit duplicate previous/next semantics.");

    std::array<ThemeCycleOverlayTheme, 2u> two{{
        ThemeCycleOverlayTheme{L"user/a", L"A"},
        ThemeCycleOverlayTheme{L"user/b", L"B"},
    }};
    const ThemeCycleOverlaySnapshot pair = BuildThemeCycleOverlaySnapshot(two, 0u, ThemeCycleDirection::Previous, 43u);
    state.Require(pair.previousThemeId == L"user/b" && pair.nextThemeId == L"user/b",
                  L"A two-entry ring should intentionally expose the same opposite entry as both neighbors.");

    std::array<ThemeCycleOverlayTheme, 5u> ring{{
        ThemeCycleOverlayTheme{L"builtin/system", L"System"},
        ThemeCycleOverlayTheme{L"builtin/light", L"Light"},
        ThemeCycleOverlayTheme{L"builtin/dark", L"Dark"},
        ThemeCycleOverlayTheme{L"builtin/rainbow", L"Rainbow"},
        ThemeCycleOverlayTheme{L"builtin/highContrast", L"High Contrast (App)"},
    }};
    for (size_t index = 0u; index < ring.size(); ++index)
    {
        const ThemeCycleOverlaySnapshot item = BuildThemeCycleOverlaySnapshot(ring, index, ThemeCycleDirection::Next, 100u + index);
        state.Require(item.currentThemeId == ring[index].themeId && item.previousThemeId == ring[(index + ring.size() - 1u) % ring.size()].themeId &&
                          item.nextThemeId == ring[(index + 1u) % ring.size()].themeId,
                      L"Every built-in ring position should wrap to its exact previous/current/next IDs.");
    }

    std::array<ThemeCycleOverlayTheme, 3u> unicode{{
        ThemeCycleOverlayTheme{L"user/latin", L"Creme\u0301 brulee\u0301"},
        ThemeCycleOverlayTheme{L"user/japanese", L"\u591C\u660E\u3051 \U0001F308"},
        ThemeCycleOverlayTheme{L"user/rtl", L"\u0645\u0648\u0636\u0648\u0639 \U0001F319"},
    }};
    const ThemeCycleOverlaySnapshot owned = BuildThemeCycleOverlaySnapshot(unicode, 1u, ThemeCycleDirection::Direct, 51u);
    unicode[0].displayName                = L"mutated previous";
    unicode[1].themeId                    = L"mutated/current";
    unicode[1].displayName                = L"mutated current";
    unicode[2].displayName                = L"mutated next";
    state.Require(owned.previousDisplayName == L"Creme\u0301 brulee\u0301" && owned.currentThemeId == L"user/japanese" &&
                      owned.currentDisplayName == L"\u591C\u660E\u3051 \U0001F308" && owned.nextDisplayName == L"\u0645\u0648\u0636\u0648\u0639 \U0001F319",
                  L"Snapshots must own and preserve combining, Japanese, supplementary, and RTL UTF-16 names.");

    constexpr std::array<UINT, 4u> kDpis{{96u, 144u, 192u, 240u}};
    for (const UINT dpi : kDpis)
    {
        const LONG widthPx  = static_cast<LONG>(1000u * dpi / 96u);
        const LONG heightPx = static_cast<LONG>(700u * dpi / 96u);
        const RECT client{100, 80, 100 + widthPx, 80 + heightPx};
        const RECT work{0, 0, 8000, 8000};
        const ThemeCycleOverlayPlacement placement = ComputeThemeCycleOverlayPlacement(client, work, dpi);
        state.Require(placement.visible, L"A normal owner should produce overlay placement at every protected DPI.");
        const LONG ownerCenterX = client.left + ((client.right - client.left) / 2);
        const LONG ownerCenterY = client.top + ((client.bottom - client.top) / 2);
        const LONG popupCenterX = placement.windowRectPx.left + ((placement.windowRectPx.right - placement.windowRectPx.left) / 2);
        const LONG popupCenterY = placement.windowRectPx.top + ((placement.windowRectPx.bottom - placement.windowRectPx.top) / 2);
        state.Require(std::abs(ownerCenterX - popupCenterX) <= 1 && std::abs(ownerCenterY - popupCenterY) <= 1,
                      L"Overlay geometry should remain centered within one physical pixel across DPI values.");
    }

    const ThemeCycleOverlayPlacement tooSmall = ComputeThemeCycleOverlayPlacement(RECT{0, 0, 191, 127}, RECT{0, 0, 1920, 1080}, 96u);
    state.Require(! tooSmall.visible, L"Owners below the defensive minimum should suppress the overlay.");
    const ThemeCycleOverlayPlacement clamped = ComputeThemeCycleOverlayPlacement(RECT{-700, -500, 300, 200}, RECT{0, 0, 1920, 1040}, 96u);
    state.Require(clamped.visible && clamped.windowRectPx.left >= 0 && clamped.windowRectPx.top >= 0 && clamped.windowRectPx.right <= 1920 &&
                      clamped.windowRectPx.bottom <= 1040,
                  L"Defensive placement should clamp a displaced owner to the monitor work area.");
    return state.failure.empty();
}

[[nodiscard]] bool TestThemeCycleOverlayKeyboardTiming(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    const std::wstring originalThemeId = g_settings.theme.currentThemeId;
    const auto restore                 = wil::scope_exit([&]() noexcept
    {
        DebugHideThemeCycleOverlay();
        g_settings.theme.currentThemeId = originalThemeId;
        SendMessageW(mainWindow, WM_THEMECHANGED, 0, 0);
        PumpPendingMessages();
    });

    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    const HWND focusBefore  = GetFocus();
    const HWND activeBefore = GetActiveWindow();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Keyboard-style Next Theme dispatch failed.");

    RedSalamander::Ui::ThemeCycleOverlayDebugSnapshot snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.created && snapshot.visible, L"Next Theme should create and show the overlay.");
    state.Require(snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::Appearing, L"First show should begin in Appearing.");
    state.Require(snapshot.direction == RedSalamander::Ui::ThemeCycleDirection::Next, L"Next Theme should expose Next motion.");
    state.Require(snapshot.surfaceOpacity <= 0.01f && snapshot.surfaceScale >= 0.939f && snapshot.surfaceScale <= 0.941f,
                  L"First frame should begin at the specified opacity/scale endpoint.");
    const HWND overlayWindow = DebugGetThemeCycleOverlayWindowHandle();
    state.Require(overlayWindow != nullptr && IsWindowVisible(overlayWindow) != FALSE, L"Overlay HWND should be visible.");
    if (overlayWindow)
    {
        const LONG_PTR exStyle = GetWindowLongPtrW(overlayWindow, GWL_EXSTYLE);
        state.Require((exStyle & WS_EX_NOACTIVATE) != 0 && (exStyle & WS_EX_TOOLWINDOW) != 0, L"Overlay must be a no-activate tool window.");
        state.Require(GetWindow(overlayWindow, GW_OWNER) == mainWindow, L"Overlay must be owned by the main window.");
    }
    state.Require(GetFocus() == focusBefore && GetActiveWindow() == activeBefore, L"Showing the overlay must preserve focus and activation.");

    DebugAdvanceThemeCycleOverlayTo(snapshot.phaseStartTickMs + RedSalamander::Ui::kThemeCycleAppearDurationMs - 1u);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::Appearing, L"Appearance must not settle before 140 ms.");

    DebugAdvanceThemeCycleOverlayTo(snapshot.phaseStartTickMs + RedSalamander::Ui::kThemeCycleAppearDurationMs);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::FullyVisible && snapshot.dismissalTimerArmed,
                  L"Appearance must settle and arm one dismissal timer at 140 ms.");
    state.Require(snapshot.disappearDeadlineMs == snapshot.steadyVisibleTickMs + RedSalamander::Ui::kThemeCycleDisappearDelayMs,
                  L"The disappearance deadline must be exactly 900 ms after steady visibility begins.");

    DebugAdvanceThemeCycleOverlayTo(snapshot.disappearDeadlineMs - 1u);
    state.Require(DebugGetThemeCycleOverlaySnapshot().phase == RedSalamander::Ui::ThemeCycleOverlayPhase::FullyVisible,
                  L"The popup must remain fully visible for the complete 900 ms delay.");
    DebugAdvanceThemeCycleOverlayTo(snapshot.disappearDeadlineMs);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::Disappearing && ! snapshot.dismissalTimerArmed,
                  L"The valid deadline must begin the 160 ms disappearance animation.");
    DebugAdvanceThemeCycleOverlayTo(snapshot.phaseStartTickMs + RedSalamander::Ui::kThemeCycleDisappearDurationMs);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::Hidden && ! snapshot.visible,
                  L"The popup must be hidden at the 160 ms disappearance endpoint.");
    return state.failure.empty();
}

[[nodiscard]] bool TestThemeCycleOverlayMenuNoOpAndReuse(HWND mainWindow, CaseState& state) noexcept
{
    const std::wstring originalThemeId = g_settings.theme.currentThemeId;
    const auto restore                 = wil::scope_exit([&]() noexcept
    {
        DebugHideThemeCycleOverlay();
        g_settings.theme.currentThemeId = originalThemeId;
        SendMessageW(mainWindow, WM_THEMECHANGED, 0, 0);
        PumpPendingMessages();
    });

    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_THEME_LIGHT, 0), 0);
    auto first               = DebugGetThemeCycleOverlaySnapshot();
    const HWND overlayWindow = DebugGetThemeCycleOverlayWindowHandle();
    state.Require(first.visible && first.direction == RedSalamander::Ui::ThemeCycleDirection::Direct,
                  L"A direct View -> Theme choice should use neutral Direct motion.");
    state.Require(g_settings.theme.currentThemeId == L"builtin/light", L"The direct menu choice should apply Light.");

    DebugAdvanceThemeCycleOverlayTo(first.phaseStartTickMs + RedSalamander::Ui::kThemeCycleAppearDurationMs);
    first = DebugGetThemeCycleOverlaySnapshot();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_THEME_LIGHT, 0), 0);
    const auto noOp = DebugGetThemeCycleOverlaySnapshot();
    state.Require(noOp.generation == first.generation && noOp.disappearDeadlineMs == first.disappearDeadlineMs,
                  L"Selecting the already active direct theme must not restart or mutate the overlay.");

    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_THEME_DARK, 0), 0);
    const auto changed = DebugGetThemeCycleOverlaySnapshot();
    state.Require(changed.generation > first.generation && changed.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::ContentTransition,
                  L"A different direct menu theme should replace content through a transition.");
    state.Require(changed.direction == RedSalamander::Ui::ThemeCycleDirection::Direct,
                  L"Direct menu replacement must remain neutral rather than imply adjacency.");
    state.Require(DebugGetThemeCycleOverlayWindowHandle() == overlayWindow, L"Direct replacements must reuse the same popup HWND.");
    state.Require(! changed.dismissalTimerArmed, L"A new transition must cancel the old disappearance timer.");
    return state.failure.empty();
}

[[nodiscard]] bool TestThemeCycleOverlaySourcesAndCustomThemes(HWND mainWindow, CaseState& state) noexcept
{
    const std::wstring originalThemeId                                  = g_settings.theme.currentThemeId;
    const std::vector<Common::Settings::ThemeDefinition> originalThemes = g_settings.theme.themes;
    const auto restore                                                  = wil::scope_exit([&]() noexcept
    {
        DebugSetForceThemeCycleDismissTimerFailure(false);
        DebugHideThemeCycleOverlay();
        g_settings.theme.themes         = originalThemes;
        g_settings.theme.currentThemeId = originalThemeId;
        SendMessageW(mainWindow, WM_THEMECHANGED, 0, 0);
        PumpPendingMessages();
    });

    Common::Settings::ThemeDefinition emptyName{};
    emptyName.id          = L"user/aurora-empty";
    emptyName.baseThemeId = L"builtin/light";
    Common::Settings::ThemeDefinition richName{};
    richName.id             = L"user/aurora-unicode";
    richName.name           = L"Ze\u0301phyr - \u65E5\u672C\u8A9E - \u0645\u0648\u0636\u0648\u0639 - \U0001F308";
    richName.baseThemeId    = L"builtin/dark";
    g_settings.theme.themes = {emptyName, richName};

    const std::vector<RedSalamander::Ui::ThemeCycleOverlayTheme> ring = DebugBuildThemeCycleRing();
    constexpr std::array<std::wstring_view, 5u> expectedBuiltIns{{
        L"builtin/system",
        L"builtin/light",
        L"builtin/dark",
        L"builtin/rainbow",
        L"builtin/highContrast",
    }};
    state.Require(ring.size() >= expectedBuiltIns.size() + 2u, L"The live cycle ring should contain built-ins and both inline custom themes.");
    if (ring.size() >= expectedBuiltIns.size() + 2u)
    {
        for (size_t index = 0u; index < expectedBuiltIns.size(); ++index)
        {
            state.Require(ring[index].themeId == expectedBuiltIns[index], L"The live ring should preserve the canonical built-in order.");
        }
        state.Require(ring[ring.size() - 2u].themeId == emptyName.id && ring.back().themeId == richName.id,
                      L"Settings themes should be appended after file themes in deterministic name/ID order.");
        state.Require(ring[ring.size() - 2u].displayName == emptyName.id && ring.back().displayName == richName.name,
                      L"Custom display names should preserve full Unicode and fall back to the ID only when empty.");
        state.Require(ring[0].displayName == LoadStringResource(nullptr, IDS_PREFS_THEMES_BASE_SYSTEM) &&
                          ring[4].displayName == LoadStringResource(nullptr, IDS_PREFS_THEMES_BASE_HIGH_CONTRAST),
                      L"Built-in ring names should resolve from the active localized resources.");
    }

    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    DebugSelectThemeTarget(mainWindow, richName.id, RedSalamander::Ui::CommandInvocationSource::ThemeMenu);
    auto snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.visible && snapshot.direction == RedSalamander::Ui::ThemeCycleDirection::Direct && snapshot.currentThemeId == richName.id &&
                      snapshot.currentDisplayName == richName.name,
                  L"A custom View -> Theme choice should show its exact Unicode name with neutral motion.");
    state.Require(snapshot.previousThemeId == emptyName.id && snapshot.nextThemeId == L"builtin/system",
                  L"A selected final settings theme should expose its exact ring neighbors and wrap forward.");
    const uint64_t customGeneration = snapshot.generation;
    DebugSelectThemeTarget(mainWindow, richName.id, RedSalamander::Ui::CommandInvocationSource::ThemeMenu);
    state.Require(DebugGetThemeCycleOverlaySnapshot().generation == customGeneration,
                  L"Selecting the active custom menu theme should remain a true overlay no-op.");

    DebugSelectThemeTarget(mainWindow, emptyName.id, RedSalamander::Ui::CommandInvocationSource::FunctionBarPointer);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.visible && snapshot.direction == RedSalamander::Ui::ThemeCycleDirection::Direct && snapshot.currentDisplayName == emptyName.id &&
                      snapshot.nextThemeId == richName.id,
                  L"A Function Bar direct custom theme should show neutral feedback and the empty-name ID fallback.");

    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_THEME_PREV, 0), 0);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.visible && snapshot.direction == RedSalamander::Ui::ThemeCycleDirection::Previous && snapshot.currentThemeId == richName.id,
                  L"View -> Theme -> Previous should show directional feedback and wrap to the final custom theme.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_THEME_NEXT, 0), 0);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.direction == RedSalamander::Ui::ThemeCycleDirection::Next && snapshot.currentThemeId == L"builtin/system",
                  L"View -> Theme -> Next should use the same directional ring and wrap back to System.");

    DebugHideThemeCycleOverlay();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/select/builtin/dark"), L"Parameterized ordinary shortcut setup should dispatch.");
    state.Require(g_settings.theme.currentThemeId == L"builtin/dark" && ! DebugGetThemeCycleOverlaySnapshot().visible,
                  L"A parameterized ordinary shortcut should apply its theme without showing the overlay.");

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Programmatic suppression warmup should show the overlay.");
    DebugSelectThemeTarget(mainWindow, L"builtin/light", RedSalamander::Ui::CommandInvocationSource::Programmatic);
    state.Require(g_settings.theme.currentThemeId == L"builtin/light" && ! DebugGetThemeCycleOverlaySnapshot().visible,
                  L"A programmatic theme change should suppress any existing overlay.");

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Other-WM_COMMAND suppression warmup should show the overlay.");
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_THEME_SYSTEM, 1), reinterpret_cast<LPARAM>(mainWindow));
    state.Require(g_settings.theme.currentThemeId == L"builtin/system" && ! DebugGetThemeCycleOverlaySnapshot().visible,
                  L"A non-menu WM_COMMAND theme change should apply and hide rather than masquerade as a menu click.");
    return state.failure.empty();
}

[[nodiscard]] bool TestThemeCycleOverlayFunctionBarAndDismiss(HWND mainWindow, CaseState& state) noexcept
{
    const std::wstring originalThemeId = g_settings.theme.currentThemeId;
    const auto restore                 = wil::scope_exit([&]() noexcept
    {
        DebugHideThemeCycleOverlay();
        g_settings.theme.currentThemeId = originalThemeId;
        SendMessageW(mainWindow, WM_THEMECHANGED, 0, 0);
        PumpPendingMessages();
    });

    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    SendMessageW(mainWindow, WndMsg::kFunctionBarInvoke, VK_F12, ShortcutManager::kModShift);
    auto snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.visible && snapshot.direction == RedSalamander::Ui::ThemeCycleDirection::Next,
                  L"A clicked Function Bar Next Theme command should show directional feedback.");
    state.Require(g_settings.theme.currentThemeId == L"builtin/light", L"The Function Bar command should apply the next theme.");

    const uint64_t dismissCount = snapshot.explicitDismissCount;
    const HWND focusBefore      = GetFocus();
    const HWND activeBefore     = GetActiveWindow();
    const LONG clickX           = (snapshot.surfaceRectPx.left + snapshot.surfaceRectPx.right) / 2;
    const LONG clickY           = (snapshot.surfaceRectPx.top + snapshot.surfaceRectPx.bottom) / 2;
    const LPARAM clickPoint     = MAKELPARAM(static_cast<WORD>(clickX), static_cast<WORD>(clickY));
    const HWND overlayWindow    = DebugGetThemeCycleOverlayWindowHandle();
    SendMessageW(overlayWindow, WM_LBUTTONDOWN, MK_LBUTTON, clickPoint);
    state.Require(DebugGetThemeCycleOverlaySnapshot().pressed && GetCapture() == overlayWindow,
                  L"A primary down inside the popup should consume and capture the full click.");
    SendMessageW(overlayWindow, WM_LBUTTONUP, 0, clickPoint);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(! snapshot.visible && snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::Hidden,
                  L"A completed primary click should dismiss immediately without an exit animation.");
    state.Require(snapshot.explicitDismissCount == dismissCount + 1u, L"Explicit dismissal should share one counted cleanup path.");
    state.Require(GetCapture() == nullptr, L"Explicit dismissal must release mouse capture.");
    state.Require(GetFocus() == focusBefore && GetActiveWindow() == activeBefore, L"Click dismissal must not change focus or activation.");
    return state.failure.empty();
}

[[nodiscard]] bool TestThemeCycleOverlayRenderingGeometryAndRecovery(HWND mainWindow, CaseState& state) noexcept
{
    const std::wstring originalThemeId                           = g_settings.theme.currentThemeId;
    const std::optional<Common::Settings::UiSettings> originalUi = g_settings.ui;
    const LONG_PTR originalOwnerExStyle                          = GetWindowLongPtrW(mainWindow, GWL_EXSTYLE);
    const auto restore                                           = wil::scope_exit([&]() noexcept
    {
        DebugHideThemeCycleOverlay();
        SetWindowLongPtrW(mainWindow, GWL_EXSTYLE, originalOwnerExStyle);
        g_settings.theme.currentThemeId = originalThemeId;
        g_settings.ui                   = originalUi;
        SendMessageW(mainWindow, WM_THEMECHANGED, 0, 0);
        PumpPendingMessages();
    });

    Common::Settings::UiSettings ui = g_settings.ui.value_or(Common::Settings::UiSettings{});
    ui.windowBackdrop               = Common::Settings::WindowBackdropMode::Acrylic;
    g_settings.ui                   = ui;
    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    const auto countersBefore   = DebugGetThemeCycleOverlaySnapshot();
    const HWND focusBefore      = GetFocus();
    const HWND activeBefore     = GetActiveWindow();
    const HWND foregroundBefore = GetForegroundWindow();
    SendMessageW(mainWindow, WM_COMMAND, MAKEWPARAM(IDM_VIEW_THEME_DARK, 0), 0);
    auto snapshot            = DebugGetThemeCycleOverlaySnapshot();
    const HWND overlayWindow = DebugGetThemeCycleOverlayWindowHandle();
    state.Require(snapshot.visible && overlayWindow != nullptr && snapshot.paintCount >= 1u && snapshot.textLayoutCreateCount >= 3u,
                  L"The first show should synchronously prime a complete painted snapshot and its three text layouts.");
    const LONG transformedWidthPx       = snapshot.surfaceRectPx.right - snapshot.surfaceRectPx.left;
    const LONG transformedHeightPx      = snapshot.surfaceRectPx.bottom - snapshot.surfaceRectPx.top;
    const float surfaceScale            = std::max(0.01f, snapshot.surfaceScale);
    const LONG expectedBackdropWidthPx  = static_cast<LONG>(std::lround(static_cast<float>(transformedWidthPx) / surfaceScale));
    const LONG expectedBackdropHeightPx = static_cast<LONG>(std::lround(static_cast<float>(transformedHeightPx) / surfaceScale));
    const auto withinCaptureRounding    = [](UINT actual, LONG expected) noexcept
    { return std::abs(static_cast<int64_t>(actual) - static_cast<int64_t>(expected)) <= 2; };
    state.Require(snapshot.backdropCaptured && snapshot.backdropCaptureCount == countersBefore.backdropCaptureCount + 1u &&
                      snapshot.backdropCaptureFailureCount == countersBefore.backdropCaptureFailureCount &&
                      withinCaptureRounding(snapshot.backdropWidthPx, expectedBackdropWidthPx) &&
                      withinCaptureRounding(snapshot.backdropHeightPx, expectedBackdropHeightPx),
                  std::format(L"An Acrylic theme-cycle popup should capture exactly one stable app-rendered backdrop for its rounded surface "
                              L"(captured={} count={} failures={} size={}x{} expected={}x{} scale={:.3f}).",
                              snapshot.backdropCaptured,
                              snapshot.backdropCaptureCount - countersBefore.backdropCaptureCount,
                              snapshot.backdropCaptureFailureCount - countersBefore.backdropCaptureFailureCount,
                              snapshot.backdropWidthPx,
                              snapshot.backdropHeightPx,
                              expectedBackdropWidthPx,
                              expectedBackdropHeightPx,
                              snapshot.surfaceScale));
    state.Require(GetFocus() == focusBefore && GetActiveWindow() == activeBefore && GetForegroundWindow() == foregroundBefore,
                  L"Showing a rendered popup must preserve focus, activation, and foreground ownership.");

    if (! overlayWindow)
    {
        return false;
    }
    const LONG_PTR style   = GetWindowLongPtrW(overlayWindow, GWL_STYLE);
    const LONG_PTR exStyle = GetWindowLongPtrW(overlayWindow, GWL_EXSTYLE);
    state.Require((style & (WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS)) == (WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS),
                  L"The popup should use the required clipped borderless styles.");
    state.Require((exStyle & (WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_NOREDIRECTIONBITMAP)) ==
                          (WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_NOREDIRECTIONBITMAP) &&
                      (exStyle & WS_EX_APPWINDOW) == 0,
                  L"The popup should be a non-activating, non-taskbar composition tool window.");
    const std::optional<Common::WindowBackdrop::Kind> appliedBackdrop = Common::WindowBackdrop::TryGetAppliedWindowBackdropKind(overlayWindow);
    state.Require(appliedBackdrop.has_value() && appliedBackdrop.value() == Common::WindowBackdrop::Kind::None,
                  L"The theme-cycle popup HWND must keep DWM system backdrop disabled; material is app-rendered inside the rounded surface.");
    state.Require(IsWindowEnabled(mainWindow) != FALSE, L"Passive theme feedback must never disable its owner.");

    RECT ownerClient{};
    GetClientRect(mainWindow, &ownerClient);
    POINT ownerOrigin{ownerClient.left, ownerClient.top};
    ClientToScreen(mainWindow, &ownerOrigin);
    ownerClient =
        RECT{ownerOrigin.x, ownerOrigin.y, ownerOrigin.x + (ownerClient.right - ownerClient.left), ownerOrigin.y + (ownerClient.bottom - ownerClient.top)};
    const LONG ownerCenterX = ownerClient.left + ((ownerClient.right - ownerClient.left) / 2);
    const LONG ownerCenterY = ownerClient.top + ((ownerClient.bottom - ownerClient.top) / 2);
    const LONG popupCenterX = snapshot.windowRectPx.left + ((snapshot.windowRectPx.right - snapshot.windowRectPx.left) / 2);
    const LONG popupCenterY = snapshot.windowRectPx.top + ((snapshot.windowRectPx.bottom - snapshot.windowRectPx.top) / 2);
    state.Require(std::abs(ownerCenterX - popupCenterX) <= 1 && std::abs(ownerCenterY - popupCenterY) <= 1,
                  L"The live popup should center over the full main client within one pixel.");
    MONITORINFO monitorInfo{sizeof(monitorInfo)};
    const HMONITOR monitor = MonitorFromWindow(mainWindow, MONITOR_DEFAULTTONEAREST);
    if (monitor && GetMonitorInfoW(monitor, &monitorInfo) != FALSE)
    {
        state.Require(snapshot.windowRectPx.left >= monitorInfo.rcWork.left && snapshot.windowRectPx.top >= monitorInfo.rcWork.top &&
                          snapshot.windowRectPx.right <= monitorInfo.rcWork.right && snapshot.windowRectPx.bottom <= monitorInfo.rcWork.bottom,
                      L"The live popup should remain inside the monitor work area.");
    }

    POINT centerClient{(snapshot.surfaceRectPx.left + snapshot.surfaceRectPx.right) / 2, (snapshot.surfaceRectPx.top + snapshot.surfaceRectPx.bottom) / 2};
    POINT centerScreen = centerClient;
    ClientToScreen(overlayWindow, &centerScreen);
    const LPARAM centerScreenParam = MAKELPARAM(static_cast<WORD>(centerScreen.x), static_cast<WORD>(centerScreen.y));
    const LPARAM gutterScreenParam = MAKELPARAM(static_cast<WORD>(snapshot.windowRectPx.left), static_cast<WORD>(snapshot.windowRectPx.top));
    state.Require(SendMessageW(overlayWindow, WM_NCHITTEST, 0, centerScreenParam) == HTCLIENT,
                  L"Hit testing should expose the visible surface as client content.");
    state.Require(SendMessageW(overlayWindow, WM_NCHITTEST, 0, gutterScreenParam) == HTTRANSPARENT,
                  L"Hit testing should pass transparent popup gutters/corners through.");
    state.Require(SendMessageW(overlayWindow, WM_MOUSEACTIVATE, reinterpret_cast<WPARAM>(mainWindow), 0) == MA_NOACTIVATE,
                  L"Pointer activation should remain explicitly non-activating.");

    const uint64_t initialLayouts = snapshot.textLayoutCreateCount;
    DebugAdvanceThemeCycleOverlayTo(snapshot.phaseStartTickMs + (RedSalamander::Ui::kThemeCycleAppearDurationMs / 2u));
    DxUi::WindowHostBitmapCapture midCapture{};
    state.Require(DebugCaptureThemeCycleOverlayBitmap(midCapture) && midCapture.widthPx > 0u && midCapture.heightPx > 0u &&
                      midCapture.bgraPixels.size() == static_cast<size_t>(midCapture.widthPx) * static_cast<size_t>(midCapture.heightPx) * 4u,
                  L"A frozen appearance frame should produce a complete composition bitmap.");
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.textLayoutCreateCount == initialLayouts, L"Rendering an animation-only appearance frame must not recreate DirectWrite layouts.");
    state.Require(GetThemeOverlayCaptureAlpha(midCapture, centerClient.x, centerClient.y) > 0u,
                  L"A visible appearance frame should contain nontransparent material at its center.");
    const POINT midInsideClient{snapshot.surfaceRectPx.right - 2, (snapshot.surfaceRectPx.top + snapshot.surfaceRectPx.bottom) / 2};
    const POINT midOutsideClient{snapshot.surfaceRectPx.right + 2, midInsideClient.y};
    POINT midInsideScreen  = midInsideClient;
    POINT midOutsideScreen = midOutsideClient;
    ClientToScreen(overlayWindow, &midInsideScreen);
    ClientToScreen(overlayWindow, &midOutsideScreen);
    state.Require(SendMessageW(overlayWindow, WM_NCHITTEST, 0, MAKELPARAM(static_cast<WORD>(midInsideScreen.x), static_cast<WORD>(midInsideScreen.y))) ==
                          HTCLIENT &&
                      SendMessageW(overlayWindow, WM_NCHITTEST, 0, MAKELPARAM(static_cast<WORD>(midOutsideScreen.x), static_cast<WORD>(midOutsideScreen.y))) ==
                          HTTRANSPARENT,
                  L"Mid-animation hit testing must follow the same transformed surface edge that paint uses.");

    DebugAdvanceThemeCycleOverlayTo(snapshot.phaseStartTickMs + RedSalamander::Ui::kThemeCycleAppearDurationMs);
    DxUi::WindowHostBitmapCapture stableCapture{};
    state.Require(DebugCaptureThemeCycleOverlayBitmap(stableCapture), L"The fully visible overlay bitmap capture should succeed.");
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::FullyVisible && snapshot.textLayoutCreateCount == initialLayouts,
                  L"Settling appearance should preserve cached layouts and enter the steady phase.");
    state.Require(GetThemeOverlayCaptureAlpha(stableCapture, centerClient.x, centerClient.y) > 0u && GetThemeOverlayCaptureAlpha(stableCapture, 0, 0) == 0u,
                  L"The stable composition capture should contain an opaque surface and transparent outer gutter.");
    state.Require((snapshot.currentRectDip.bottom - snapshot.currentRectDip.top) > (snapshot.previousRectDip.bottom - snapshot.previousRectDip.top) &&
                      (snapshot.currentRectDip.bottom - snapshot.currentRectDip.top) > (snapshot.nextRectDip.bottom - snapshot.nextRectDip.top),
                  L"The semantic current-theme region should be visibly larger than both neighbor regions.");

    const uint64_t steadyPaintCount = snapshot.paintCount;
    DebugAdvanceThemeCycleOverlayTo(snapshot.disappearDeadlineMs - 1u);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::FullyVisible && snapshot.paintCount == steadyPaintCount,
                  L"The 900 ms steady delay should request and render no animation frames.");

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Repeat/recovery setup dispatch failed.");
    auto replacement = DebugGetThemeCycleOverlaySnapshot();
    state.Require(replacement.backdropCaptured && replacement.backdropCaptureCount == snapshot.backdropCaptureCount,
                  L"A visible theme-cycle replacement should reuse its clean backdrop snapshot instead of capturing the popup into itself.");
    const uint64_t staleBefore = replacement.staleTimerIgnoredCount;
    SendMessageW(overlayWindow, WM_TIMER, 1u, 0);
    replacement = DebugGetThemeCycleOverlaySnapshot();
    state.Require(replacement.visible && replacement.generation > snapshot.generation && replacement.staleTimerIgnoredCount == staleBefore + 1u,
                  L"A delivered old dismissal timer should be rejected without hiding the newest generation.");
    const uint64_t replacementLayouts = replacement.textLayoutCreateCount;
    DebugAdvanceThemeCycleOverlayTo(replacement.phaseStartTickMs + (RedSalamander::Ui::kThemeCycleAdjacentTransitionDurationMs / 2u));
    DxUi::WindowHostBitmapCapture transitionCapture{};
    state.Require(DebugCaptureThemeCycleOverlayBitmap(transitionCapture), L"A frozen adjacent transition capture should succeed.");
    state.Require(DebugGetThemeCycleOverlaySnapshot().textLayoutCreateCount == replacementLayouts,
                  L"An adjacent animation frame must not create additional text layouts.");
    DebugAdvanceThemeCycleOverlayTo(replacement.phaseStartTickMs + RedSalamander::Ui::kThemeCycleAdjacentTransitionDurationMs);
    replacement = DebugGetThemeCycleOverlaySnapshot();
    DebugAdvanceThemeCycleOverlayTo(replacement.disappearDeadlineMs);
    auto disappearing = DebugGetThemeCycleOverlaySnapshot();
    DebugAdvanceThemeCycleOverlayTo(disappearing.phaseStartTickMs + (RedSalamander::Ui::kThemeCycleDisappearDurationMs / 2u));
    disappearing                   = DebugGetThemeCycleOverlaySnapshot();
    const float interruptedOpacity = disappearing.surfaceOpacity;
    const float interruptedScale   = disappearing.surfaceScale;
    const float interruptedY       = disappearing.surfaceTranslateYDip;
    state.Require(interruptedOpacity > 0.0f && interruptedOpacity < 1.0f, L"The frozen exit midpoint should expose an intermediate composed opacity.");

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Exit-cancel replacement dispatch failed.");
    auto recovered = DebugGetThemeCycleOverlaySnapshot();
    state.Require(recovered.visible && recovered.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::ContentTransition &&
                      std::abs(recovered.surfaceOpacity - interruptedOpacity) < 0.01f && std::abs(recovered.surfaceScale - interruptedScale) < 0.01f &&
                      std::abs(recovered.surfaceTranslateYDip - interruptedY) < 0.01f,
                  L"Input during exit should preserve the sampled surface opacity/scale/translation without a flash.");
    DebugAdvanceThemeCycleOverlayTo(recovered.phaseStartTickMs + (RedSalamander::Ui::kThemeCycleExitCancelRecoveryMs / 2u));
    const auto recoveryMid = DebugGetThemeCycleOverlaySnapshot();
    state.Require(recoveryMid.surfaceOpacity > interruptedOpacity && recoveryMid.surfaceOpacity < 1.0f,
                  L"Exit cancellation should animate continuously back toward full visibility.");
    DebugAdvanceThemeCycleOverlayTo(recovered.phaseStartTickMs + RedSalamander::Ui::kThemeCycleExitCancelRecoveryMs);
    recovered = DebugGetThemeCycleOverlaySnapshot();
    state.Require(std::abs(recovered.surfaceOpacity - 1.0f) < 0.001f && std::abs(recovered.surfaceScale - 1.0f) < 0.001f &&
                      std::abs(recovered.surfaceTranslateYDip) < 0.001f,
                  L"Exit cancellation should restore the canonical surface endpoint within 80 ms.");

    SetWindowLongPtrW(mainWindow, GWL_EXSTYLE, originalOwnerExStyle | WS_EX_LAYOUTRTL);
    DebugSelectThemeTarget(mainWindow, L"builtin/light", RedSalamander::Ui::CommandInvocationSource::ThemeMenu);
    auto rtl                   = DebugGetThemeCycleOverlaySnapshot();
    const float currentCenterX = (rtl.currentRectDip.left + rtl.currentRectDip.right) * 0.5f;
    state.Require(rtl.previousRectDip.left > currentCenterX && rtl.nextRectDip.right < currentCenterX,
                  L"RTL flow should mirror logical previous to top-leading and next to bottom-trailing.");
    SetWindowLongPtrW(mainWindow, GWL_EXSTYLE, originalOwnerExStyle);

    DebugSelectThemeTarget(mainWindow, L"builtin/highContrast", RedSalamander::Ui::CommandInvocationSource::ThemeMenu);
    auto highContrast = DebugGetThemeCycleOverlaySnapshot();
    DebugAdvanceThemeCycleOverlayTo(highContrast.phaseStartTickMs + RedSalamander::Ui::kThemeCycleDirectTransitionDurationMs);
    DxUi::WindowHostBitmapCapture highContrastCapture{};
    state.Require(DebugCaptureThemeCycleOverlayBitmap(highContrastCapture), L"The app High Contrast popup capture should succeed.");
    highContrast = DebugGetThemeCycleOverlaySnapshot();
    state.Require(highContrast.highContrast && highContrast.surfaceOpacity == 1.0f && ! highContrast.backdropCaptured &&
                      GetThemeOverlayCaptureAlpha(highContrastCapture,
                                                  (highContrast.surfaceRectPx.left + highContrast.surfaceRectPx.right) / 2,
                                                  (highContrast.surfaceRectPx.top + highContrast.surfaceRectPx.bottom) / 2) > 0u &&
                      GetThemeOverlayCaptureAlpha(highContrastCapture, 0, 0) == 0u,
                  L"High Contrast should use a fully opaque solid surface while preserving transparent gutters.");

    const uint64_t paintBeforeDeviceLoss = highContrast.paintCount;
    DebugSimulateThemeCycleOverlayDeviceLoss();
    DxUi::WindowHostBitmapCapture recoveredCapture{};
    state.Require(DebugCaptureThemeCycleOverlayBitmap(recoveredCapture), L"The overlay should recreate its device resources after simulated loss.");
    state.Require(DebugGetThemeCycleOverlaySnapshot().paintCount > paintBeforeDeviceLoss,
                  L"Device-loss recovery should repaint the current generation without changing the theme.");

    highContrast         = DebugGetThemeCycleOverlaySnapshot();
    const LONG insideX   = (highContrast.surfaceRectPx.left + highContrast.surfaceRectPx.right) / 2;
    const LONG insideY   = (highContrast.surfaceRectPx.top + highContrast.surfaceRectPx.bottom) / 2;
    const LPARAM inside  = MAKELPARAM(static_cast<WORD>(insideX), static_cast<WORD>(insideY));
    const LPARAM outside = MAKELPARAM(static_cast<WORD>(highContrast.surfaceRectPx.right + 1), static_cast<WORD>(insideY));
    SendMessageW(overlayWindow, WM_LBUTTONDOWN, MK_LBUTTON, inside);
    SendMessageW(overlayWindow, WM_LBUTTONUP, 0, outside);
    state.Require(DebugGetThemeCycleOverlaySnapshot().visible && ! DebugGetThemeCycleOverlaySnapshot().pressed && GetCapture() == nullptr,
                  L"A primary release outside the surface should cancel capture without dismissing or clicking through.");
    SendMessageW(overlayWindow, WM_LBUTTONDOWN, MK_LBUTTON, inside);
    SendMessageW(overlayWindow, WM_CANCELMODE, 0, 0);
    state.Require(DebugGetThemeCycleOverlaySnapshot().visible && ! DebugGetThemeCycleOverlaySnapshot().pressed && GetCapture() == nullptr,
                  L"WM_CANCELMODE should clear the pressed/capture state without dismissal.");
    state.Require(SendMessageW(overlayWindow, WM_RBUTTONDOWN, MK_RBUTTON, inside) == 0 &&
                      SendMessageW(overlayWindow, WM_MOUSEWHEEL, MAKEWPARAM(0, WHEEL_DELTA), inside) == 0 && DebugGetThemeCycleOverlaySnapshot().visible,
                  L"Secondary and wheel input over the popup should be consumed without invoking or dismissing it.");
    return state.failure.empty();
}

[[nodiscard]] bool TestThemeCycleOverlayReducedMotion(HWND mainWindow, CaseState& state) noexcept
{
    const std::wstring originalThemeId                           = g_settings.theme.currentThemeId;
    const std::optional<Common::Settings::UiSettings> originalUi = g_settings.ui;
    const auto restore                                           = wil::scope_exit([&]() noexcept
    {
        DebugHideThemeCycleOverlay();
        g_settings.ui                   = originalUi;
        g_settings.theme.currentThemeId = originalThemeId;
        SendMessageW(mainWindow, WM_THEMECHANGED, 0, 0);
        PumpPendingMessages();
    });

    if (! g_settings.ui.has_value())
    {
        g_settings.ui = Common::Settings::UiSettings{};
    }
    g_settings.ui.value().reducedMotion = Common::Settings::ReducedMotionMode::On;
    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Reduced-motion Next Theme dispatch failed.");
    auto snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.reducedMotion && snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::FullyVisible,
                  L"Reduced motion should present the complete snapshot atomically.");
    state.Require(! snapshot.animationActive && snapshot.dismissalTimerArmed && snapshot.surfaceOpacity == 1.0f && snapshot.surfaceScale == 1.0f,
                  L"Reduced motion should skip intermediate animation while retaining the one-shot timer.");
    DebugAdvanceThemeCycleOverlayTo(snapshot.disappearDeadlineMs - 1u);
    state.Require(DebugGetThemeCycleOverlaySnapshot().visible, L"Reduced motion must preserve the full 900 ms readable delay.");
    DebugAdvanceThemeCycleOverlayTo(snapshot.disappearDeadlineMs);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(! snapshot.visible && snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::Hidden,
                  L"Reduced motion should hide immediately at the valid deadline.");
    return state.failure.empty();
}

[[nodiscard]] bool TestThemeCycleOverlayTimerFallback(HWND mainWindow, CaseState& state) noexcept
{
    const std::wstring originalThemeId = g_settings.theme.currentThemeId;
    const auto restore                 = wil::scope_exit([&]() noexcept
    {
        DebugSetForceThemeCycleDismissTimerFailure(false);
        DebugHideThemeCycleOverlay();
        g_settings.theme.currentThemeId = originalThemeId;
        SendMessageW(mainWindow, WM_THEMECHANGED, 0, 0);
        PumpPendingMessages();
    });

    // Ensure the lazy implementation exists before applying its test-only failure switch.
    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Fallback warmup dispatch failed.");
    DebugHideThemeCycleOverlay();
    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    DebugSetForceThemeCycleDismissTimerFailure(true);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Forced timer-failure dispatch failed.");
    auto snapshot = DebugGetThemeCycleOverlaySnapshot();
    DebugAdvanceThemeCycleOverlayTo(snapshot.phaseStartTickMs + RedSalamander::Ui::kThemeCycleAppearDurationMs);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.visible && snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::FullyVisible && ! snapshot.dismissalTimerArmed &&
                      snapshot.fallbackDeadlineArmed,
                  L"A forced SetTimer failure should arm the physical thread-pool deadline fallback.");

    const HWND overlayWindow   = DebugGetThemeCycleOverlayWindowHandle();
    const uint64_t staleBefore = snapshot.staleTimerIgnoredCount;
    SendMessageW(overlayWindow, WndMsg::kThemeCycleOverlayFallbackDeadline, 0u, 0);
    snapshot = DebugGetThemeCycleOverlaySnapshot();
    state.Require(snapshot.visible && snapshot.fallbackDeadlineArmed && snapshot.staleTimerIgnoredCount == staleBefore + 1u,
                  L"A stale fallback cookie should be ignored without canceling the active deadline.");

    const uint64_t steadyStart = snapshot.steadyVisibleTickMs;
    state.Require(WaitForThemeOverlayHidden(std::chrono::milliseconds(3'000)),
                  L"The real thread-pool fallback should post, disappear, and hide within the bounded test timeout.");
    const uint64_t observedDelay = GetTickCount64() - steadyStart;
    snapshot                     = DebugGetThemeCycleOverlaySnapshot();
    state.Require(observedDelay >= RedSalamander::Ui::kThemeCycleDisappearDelayMs,
                  std::format(L"The physical fallback must never begin an early disappearance; observed={0}ms required={1}ms steadyStart={2}ms.",
                              observedDelay,
                              RedSalamander::Ui::kThemeCycleDisappearDelayMs,
                              steadyStart));
    state.Require(! snapshot.visible && snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::Hidden && ! snapshot.dismissalTimerArmed &&
                      ! snapshot.fallbackDeadlineArmed && ! snapshot.animationActive,
                  L"Fallback completion should converge to hidden with no timer or animation subscription retained.");
    return state.failure.empty();
}

[[nodiscard]] bool TestThemeCycleOverlayAccessibility(HWND mainWindow, CaseState& state) noexcept
{
    const std::wstring originalThemeId = g_settings.theme.currentThemeId;
    const auto restore                 = wil::scope_exit([&]() noexcept
    {
        DebugHideThemeCycleOverlay();
        g_settings.theme.currentThemeId = originalThemeId;
        SendMessageW(mainWindow, WM_THEMECHANGED, 0, 0);
        PumpPendingMessages();
    });

    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Accessibility setup dispatch failed.");

    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> root;
    root.attach(DebugCreateThemeCycleOverlayAccessibilityProvider());
    state.Require(root != nullptr, L"Theme overlay should expose a WindowHost accessibility provider.");
    wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
    if (root)
    {
        state.Require(SUCCEEDED(root.query_to(simple.put())) && simple != nullptr, L"Theme overlay root should expose provider-simple semantics.");
    }
    if (! simple)
    {
        return false;
    }

    VARIANT value{};
    VariantInit(&value);
    state.Require(SUCCEEDED(simple->GetPropertyValue(UIA_ControlTypePropertyId, &value)) && value.vt == VT_I4 && value.lVal == UIA_StatusBarControlTypeId,
                  L"Theme overlay root should expose the Status control type.");
    VariantClear(&value);
    VariantInit(&value);
    state.Require(SUCCEEDED(simple->GetPropertyValue(UIA_AutomationIdPropertyId, &value)) && value.vt == VT_BSTR &&
                      std::wstring_view(value.bstrVal ? value.bstrVal : L"") == L"ThemeCycleOverlay",
                  L"Theme overlay root should expose its stable automation ID.");
    VariantClear(&value);
    VariantInit(&value);
    state.Require(SUCCEEDED(simple->GetPropertyValue(UIA_IsKeyboardFocusablePropertyId, &value)) && value.vt == VT_BOOL && value.boolVal == VARIANT_FALSE,
                  L"Theme overlay Status should remain outside keyboard focus order.");
    VariantClear(&value);

    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Accessibility most-recent generation replacement failed.");
    const auto latest = DebugGetThemeCycleOverlaySnapshot();
    VariantInit(&value);
    const std::wstring expectedRootName =
        FormatStringResource(nullptr, IDS_THEME_CYCLE_OVERLAY_UIA_FULL, latest.currentDisplayName, latest.previousDisplayName, latest.nextDisplayName);
    state.Require(SUCCEEDED(simple->GetPropertyValue(UIA_NamePropertyId, &value)) && value.vt == VT_BSTR &&
                      std::wstring_view(value.bstrVal ? value.bstrVal : L"") == expectedRootName,
                  L"The retained root provider should expose the newest generation's full live-region text.");
    VariantClear(&value);

    const std::array<std::wstring_view, 3u> expectedAutomationIds{{
        L"ThemeCycleOverlay.Previous",
        L"ThemeCycleOverlay.Current",
        L"ThemeCycleOverlay.Next",
    }};
    const std::array<std::wstring, 3u> expectedNames{{
        FormatStringResource(nullptr, IDS_THEME_CYCLE_OVERLAY_PREVIOUS_NAME, latest.previousDisplayName),
        FormatStringResource(nullptr, IDS_THEME_CYCLE_OVERLAY_CURRENT_NAME, latest.currentDisplayName),
        FormatStringResource(nullptr, IDS_THEME_CYCLE_OVERLAY_NEXT_NAME, latest.nextDisplayName),
    }};
    wil::com_ptr_nothrow<IRawElementProviderFragment> rootFragment;
    state.Require(SUCCEEDED(root.query_to(rootFragment.put())) && rootFragment != nullptr, L"The Status root should also expose fragment navigation.");
    wil::com_ptr_nothrow<IRawElementProviderFragment> child;
    state.Require(rootFragment && SUCCEEDED(rootFragment->Navigate(NavigateDirection_FirstChild, child.put())) && child != nullptr,
                  L"The Status root should expose its first semantic text child.");
    for (size_t index = 0u; index < expectedAutomationIds.size() && child; ++index)
    {
        wil::com_ptr_nothrow<IRawElementProviderSimple> childSimple;
        state.Require(SUCCEEDED(child.query_to(childSimple.put())) && childSimple != nullptr,
                      L"Each semantic text fragment should expose provider-simple properties.");
        if (! childSimple)
        {
            break;
        }

        VariantInit(&value);
        state.Require(SUCCEEDED(childSimple->GetPropertyValue(UIA_AutomationIdPropertyId, &value)) && value.vt == VT_BSTR &&
                          std::wstring_view(value.bstrVal ? value.bstrVal : L"") == expectedAutomationIds[index],
                      L"Each semantic theme child should retain its stable positional automation ID.");
        VariantClear(&value);
        VariantInit(&value);
        state.Require(SUCCEEDED(childSimple->GetPropertyValue(UIA_NamePropertyId, &value)) && value.vt == VT_BSTR &&
                          std::wstring_view(value.bstrVal ? value.bstrVal : L"") == expectedNames[index],
                      L"Each semantic theme child should expose the full unellipsized localized name.");
        VariantClear(&value);
        VariantInit(&value);
        state.Require(SUCCEEDED(childSimple->GetPropertyValue(UIA_IsKeyboardFocusablePropertyId, &value)) && value.vt == VT_BOOL &&
                          value.boolVal == VARIANT_FALSE,
                      L"Semantic theme text children should not enter keyboard focus order.");
        VariantClear(&value);
        wil::com_ptr_nothrow<IUnknown> childInvoke;
        state.Require(SUCCEEDED(childSimple->GetPatternProvider(UIA_InvokePatternId, childInvoke.put())) && childInvoke == nullptr,
                      L"Only the Status root should expose the dismiss Invoke pattern.");

        wil::com_ptr_nothrow<IRawElementProviderFragment> next;
        state.Require(SUCCEEDED(child->Navigate(NavigateDirection_NextSibling, next.put())), L"Semantic child sibling navigation should succeed.");
        child = std::move(next);
    }
    state.Require(child == nullptr, L"The Status root should expose exactly Previous, Current, and Next semantic children.");

    wil::com_ptr_nothrow<IUnknown> pattern;
    state.Require(SUCCEEDED(simple->GetPatternProvider(UIA_InvokePatternId, pattern.put())) && pattern != nullptr,
                  L"Theme overlay root should expose Invoke for dismissal.");
    wil::com_ptr_nothrow<IInvokeProvider> invoke;
    if (pattern)
    {
        state.Require(SUCCEEDED(pattern.query_to(invoke.put())) && invoke != nullptr, L"Invoke pattern should be queryable.");
    }
    if (invoke)
    {
        state.Require(SUCCEEDED(invoke->Invoke()), L"UIA Invoke should dismiss successfully.");
        state.Require(! DebugGetThemeCycleOverlaySnapshot().visible, L"UIA Invoke should use immediate dismissal.");
    }
    return state.failure.empty();
}

[[nodiscard]] bool TestThemeCycleOverlayRetention(HWND mainWindow, CaseState& state) noexcept
{
    const std::wstring originalThemeId = g_settings.theme.currentThemeId;
    const auto restore                 = wil::scope_exit([&]() noexcept
    {
        DebugHideThemeCycleOverlay();
        g_settings.theme.currentThemeId = originalThemeId;
        SendMessageW(mainWindow, WM_THEMECHANGED, 0, 0);
        PumpPendingMessages();
    });

    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Retention warmup dispatch failed.");
    const HWND stableWindow   = DebugGetThemeCycleOverlayWindowHandle();
    const auto countersBefore = DebugGetThemeCycleOverlaySnapshot();
    const DWORD gdiBefore     = GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS);
    const DWORD userBefore    = GetGuiResources(GetCurrentProcess(), GR_USEROBJECTS);
    PROCESS_MEMORY_COUNTERS_EX memoryBefore{};
    memoryBefore.cb = sizeof(memoryBefore);
    static_cast<void>(GetProcessMemoryInfo(GetCurrentProcess(), reinterpret_cast<PROCESS_MEMORY_COUNTERS*>(&memoryBefore), sizeof(memoryBefore)));

    constexpr uint64_t kIterationCount = 1'000u;
    for (uint64_t iteration = 0u; iteration < kIterationCount; ++iteration)
    {
        state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Retention burst dispatch failed.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    PROCESS_MEMORY_COUNTERS_EX memoryAfter{};
    memoryAfter.cb = sizeof(memoryAfter);
    static_cast<void>(GetProcessMemoryInfo(GetCurrentProcess(), reinterpret_cast<PROCESS_MEMORY_COUNTERS*>(&memoryAfter), sizeof(memoryAfter)));
    const DWORD gdiAfter  = GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS);
    const DWORD userAfter = GetGuiResources(GetCurrentProcess(), GR_USEROBJECTS);
    const auto snapshot   = DebugGetThemeCycleOverlaySnapshot();
    state.Require(DebugGetThemeCycleOverlayWindowHandle() == stableWindow, L"A 1,000-command burst must reuse one overlay HWND.");
    state.Require(gdiAfter <= gdiBefore + 2u && userAfter <= userBefore + 2u, L"A 1,000-command burst must not retain GDI/USER handles.");
    state.Require(memoryAfter.PrivateUsage <= memoryBefore.PrivateUsage + (8u * 1024u * 1024u),
                  L"A 1,000-command burst must keep private-memory growth bounded.");
    state.Require(snapshot.phase == RedSalamander::Ui::ThemeCycleOverlayPhase::ContentTransition && ! snapshot.dismissalTimerArmed,
                  L"The burst should converge on one latest transition without accumulating timers.");
    state.Require(snapshot.windowCreateCount == countersBefore.windowCreateCount &&
                      snapshot.windowReuseCount == countersBefore.windowReuseCount + kIterationCount,
                  L"Every warm burst command should reuse the single retained popup HWND.");
    state.Require(snapshot.droppedGenerationCount == countersBefore.droppedGenerationCount,
                  L"Synchronous matching-generation presentation should drop no accepted burst generation.");
    DebugHideThemeCycleOverlay();
    const auto hidden = DebugGetThemeCycleOverlaySnapshot();
    state.Require(! hidden.visible && ! hidden.animationActive && ! hidden.dismissalTimerArmed && ! hidden.fallbackDeadlineArmed && ! hidden.captured &&
                      ! hidden.pressed,
                  L"Retention teardown should leave no animation, timer, pointer capture, or pressed state.");

    const uint64_t createCountBeforeExternalDestroy = hidden.windowCreateCount;
    state.Require(stableWindow && DestroyWindow(stableWindow) != FALSE, L"The retained overlay HWND should support owner-equivalent external teardown.");
    const auto externallyDestroyed = DebugGetThemeCycleOverlaySnapshot();
    state.Require(! externallyDestroyed.created && ! externallyDestroyed.visible && ! externallyDestroyed.animationActive &&
                      ! externallyDestroyed.dismissalTimerArmed && ! externallyDestroyed.fallbackDeadlineArmed,
                  L"WM_NCDESTROY should clear the retained control view, HWND ownership, animation, and deadlines.");
    DebugHideThemeCycleOverlay();
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"The overlay should recreate cleanly after external HWND teardown.");
    const auto recreated = DebugGetThemeCycleOverlaySnapshot();
    state.Require(recreated.created && recreated.visible && recreated.windowCreateCount == createCountBeforeExternalDestroy + 1u,
                  L"A post-teardown theme change should create exactly one fresh overlay window.");
    DebugHideThemeCycleOverlay();
    Debug::Perf::EmitCounter(L"theme.cycle.overlay.retention_command_count", kIterationCount);
    return state.failure.empty();
}

[[nodiscard]] bool TestThemeCycleOverlayPerfBaseline(HWND mainWindow, CaseState& state) noexcept
{
    if (! mainWindow || IsWindow(mainWindow) == FALSE)
    {
        state.Require(false, L"Main window handle invalid.");
        return false;
    }

    if (! PrepareMainWindowForIsolatedUiCase(mainWindow, state, L"theme-cycle overlay performance validation"))
    {
        return false;
    }

    DebugHideThemeCycleOverlay();
    if (const HWND retainedOverlay = DebugGetThemeCycleOverlayWindowHandle(); retainedOverlay)
    {
        state.Require(DestroyWindow(retainedOverlay) != FALSE, L"Failed to reset the retained theme-cycle overlay HWND before performance validation.");
    }
    PumpPendingMessages();
    if (! state.failure.empty())
    {
        return false;
    }

    constexpr uint64_t kIterationCount                           = 240u;
    const std::wstring originalThemeId                           = g_settings.theme.currentThemeId;
    const std::optional<Common::Settings::UiSettings> originalUi = g_settings.ui;
    const auto restore                                           = wil::scope_exit([&]() noexcept
    {
        DebugHideThemeCycleOverlay();
        g_settings.theme.currentThemeId = originalThemeId;
        g_settings.ui                   = originalUi;
        SendMessageW(mainWindow, WM_THEMECHANGED, 0, 0);
        PumpPendingMessages();
    });

    Common::Settings::UiSettings ui = g_settings.ui.value_or(Common::Settings::UiSettings{});
    ui.windowBackdrop               = Common::Settings::WindowBackdropMode::Acrylic;
    g_settings.ui                   = ui;
    // Exercise the retained hidden HWND path used by the broad Commands run so
    // the fixture is self-contained when selected as an exact retry.
    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    state.Require(DebugDispatchShortcutCommand(mainWindow, L"cmd/app/theme/selectNext"), L"Theme-cycle perf retained-window warmup failed.");
    DebugHideThemeCycleOverlay();
    SetThemeWithoutOverlay(mainWindow, L"builtin/system", state);
    const auto before = DebugGetThemeCycleOverlaySnapshot();
    for (uint64_t iteration = 0u; iteration < kIterationCount; ++iteration)
    {
        const std::wstring_view commandId =
            (iteration % 16u) < 12u ? std::wstring_view{L"cmd/app/theme/selectNext"} : std::wstring_view{L"cmd/app/theme/selectPrev"};
        state.Require(DebugDispatchShortcutCommand(mainWindow, commandId), L"Theme-cycle perf command dispatch failed.");
        if (! state.failure.empty())
        {
            return false;
        }
    }

    auto after                                 = DebugGetThemeCycleOverlaySnapshot();
    const uint64_t createDelta                 = after.windowCreateCount - before.windowCreateCount;
    const uint64_t reuseDelta                  = after.windowReuseCount - before.windowReuseCount;
    const uint64_t renderDelta                 = after.paintCount - before.paintCount;
    const uint64_t layoutDelta                 = after.textLayoutCreateCount - before.textLayoutCreateCount;
    const uint64_t backdropCaptureDelta        = after.backdropCaptureCount - before.backdropCaptureCount;
    const uint64_t backdropCaptureFailureDelta = after.backdropCaptureFailureCount - before.backdropCaptureFailureCount;
    state.Require(createDelta <= 1u && createDelta + reuseDelta == kIterationCount,
                  L"The perf burst should create at most one lazy HWND and account for every later update as reuse.");
    state.Require(backdropCaptureDelta == 1u && backdropCaptureFailureDelta == 0u,
                  std::format(L"The Acrylic perf burst should capture one clean app backdrop and reuse it for every visible update; "
                              L"captures={0} failures={1} beforeCaptured={2} afterCaptured={3} size={4}x{5}.",
                              backdropCaptureDelta,
                              backdropCaptureFailureDelta,
                              before.backdropCaptured,
                              after.backdropCaptured,
                              after.backdropWidthPx,
                              after.backdropHeightPx));
    state.Require(renderDelta >= kIterationCount, L"Every accepted perf generation should synchronously render a matching complete frame.");
    state.Require(after.droppedGenerationCount == before.droppedGenerationCount,
                  L"The synchronous perf burst should not drop an accepted generation before presentation.");

    DebugAdvanceThemeCycleOverlayTo(after.phaseStartTickMs + RedSalamander::Ui::kThemeCycleAdjacentTransitionDurationMs);
    after                           = DebugGetThemeCycleOverlaySnapshot();
    const uint64_t steadyPaintCount = after.paintCount;
    DebugAdvanceThemeCycleOverlayTo(after.disappearDeadlineMs - 1u);
    state.Require(DebugGetThemeCycleOverlaySnapshot().paintCount == steadyPaintCount,
                  L"The perf fixture should prove that the 900 ms steady interval renders no frames.");
    DebugAdvanceThemeCycleOverlayTo(after.disappearDeadlineMs);
    auto exit = DebugGetThemeCycleOverlaySnapshot();
    DebugAdvanceThemeCycleOverlayTo(exit.phaseStartTickMs + RedSalamander::Ui::kThemeCycleDisappearDurationMs);
    after = DebugGetThemeCycleOverlaySnapshot();
    state.Require(! after.visible && ! after.animationActive && ! after.dismissalTimerArmed && ! after.fallbackDeadlineArmed,
                  L"The perf fixture should close one complete lifetime without retaining animation or deadline state.");

    Debug::Perf::EmitCounter(L"theme.cycle.perf.accepted_command_count", kIterationCount);
    Debug::Perf::Emit(L"theme.cycle.overlay.window_create_count", L"case-aggregate", 0u, createDelta);
    Debug::Perf::Emit(L"theme.cycle.overlay.reuse_count", L"case-aggregate", 0u, reuseDelta);
    Debug::Perf::Emit(L"theme.cycle.overlay.render_count", L"case-aggregate", 0u, renderDelta);
    Debug::Perf::Emit(L"theme.cycle.overlay.layout_create_count", L"case-aggregate", 0u, layoutDelta);
    Debug::Perf::Emit(L"theme.cycle.overlay.backdrop_capture_count", L"case-aggregate", 0u, backdropCaptureDelta);
    Debug::Perf::Emit(L"theme.cycle.overlay.backdrop_capture_failure_count", L"case-aggregate", 0u, backdropCaptureFailureDelta);
    Debug::Perf::Emit(L"theme.cycle.overlay.dismiss_timer_arm_count", L"case-aggregate", 0u, after.dismissTimerArmCount - before.dismissTimerArmCount);
    Debug::Perf::Emit(L"theme.cycle.overlay.stale_timer_ignored_count", L"case-aggregate", 0u, after.staleTimerIgnoredCount - before.staleTimerIgnoredCount);
    Debug::Perf::Emit(L"theme.cycle.overlay.dropped_generation_count", L"case-aggregate", 0u, after.droppedGenerationCount - before.droppedGenerationCount);
    state.Require(g_settings.theme.currentThemeId == L"builtin/system", L"The deterministic theme-cycle burst should end on the built-in system theme.");
    return state.failure.empty();
}
} // namespace

void RunThemeCycleOverlayCommandsSelfTestCases(HWND mainWindow, const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite) noexcept
{
    SelfTest::RunCase(options, suite, L"theme_cycle_overlay_pure_state_geometry", [=](CaseState& state) noexcept {
        return TestThemeCycleOverlayPureStateAndGeometry(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"theme_cycle_overlay_keyboard_timing", [=](CaseState& state) noexcept {
        return TestThemeCycleOverlayKeyboardTiming(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"theme_cycle_overlay_menu_noop_reuse", [=](CaseState& state) noexcept {
        return TestThemeCycleOverlayMenuNoOpAndReuse(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"theme_cycle_overlay_sources_custom_themes", [=](CaseState& state) noexcept {
        return TestThemeCycleOverlaySourcesAndCustomThemes(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"theme_cycle_overlay_function_bar_dismiss", [=](CaseState& state) noexcept {
        return TestThemeCycleOverlayFunctionBarAndDismiss(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"theme_cycle_overlay_rendering_geometry_recovery", [=](CaseState& state) noexcept {
        return TestThemeCycleOverlayRenderingGeometryAndRecovery(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"theme_cycle_overlay_accessibility", [=](CaseState& state) noexcept { return TestThemeCycleOverlayAccessibility(mainWindow, state); });
    SelfTest::RunCase(options, suite, L"theme_cycle_overlay_reduced_motion", [=](CaseState& state) noexcept {
        return TestThemeCycleOverlayReducedMotion(mainWindow, state);
    });
    SelfTest::RunCase(options, suite, L"theme_cycle_overlay_timer_fallback", [=](CaseState& state) noexcept {
        return TestThemeCycleOverlayTimerFallback(mainWindow, state);
    });
    SelfTest::RunCase(
        options, suite, L"theme_cycle_overlay_retention", [=](CaseState& state) noexcept { return TestThemeCycleOverlayRetention(mainWindow, state); });
    SelfTest::RunCase(
        options, suite, L"theme_cycle_overlay_perf", [=](CaseState& state) noexcept { return TestThemeCycleOverlayPerfBaseline(mainWindow, state); });
}
