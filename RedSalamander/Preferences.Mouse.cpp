// Preferences.Mouse.cpp

#include "Framework.h"

#include "Preferences.Mouse.h"

#include "DxUiThemePalette.h"
#include "Helpers.h"
#include "UiMetrics.h"

#include "resource.h"

namespace
{
using RedSalamander::DxUi::CardPanel;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;

[[nodiscard]] Common::Settings::MouseSettings GetMouseSettings(const Common::Settings::Settings& settings) noexcept
{
    return settings.mouse.value_or(Common::Settings::MouseSettings{});
}

void SetMouseSettings(PreferencesDialogState& state, const Common::Settings::MouseSettings& mouse) noexcept
{
    if (mouse == Common::Settings::MouseSettings{})
    {
        state.workingSettings.mouse.reset();
    }
    else
    {
        state.workingSettings.mouse = mouse;
    }
}

struct MouseDxPage
{
    Label* paneFocusHeader                       = nullptr;
    CardPanel* focusFollowsPointerCard           = nullptr;
    Label* focusFollowsPointerTitle              = nullptr;
    Label* focusFollowsPointerDescription        = nullptr;
    Toggle* focusFollowsPointerToggle            = nullptr;
    CardPanel* terminalOpenFocusCard             = nullptr;
    Label* terminalOpenFocusTitle                = nullptr;
    Label* terminalOpenFocusDescription          = nullptr;
    Toggle* terminalOpenFocusToggle              = nullptr;

    void Detach() noexcept
    {
        paneFocusHeader                       = nullptr;
        focusFollowsPointerCard               = nullptr;
        focusFollowsPointerTitle              = nullptr;
        focusFollowsPointerDescription        = nullptr;
        focusFollowsPointerToggle             = nullptr;
        terminalOpenFocusCard                 = nullptr;
        terminalOpenFocusTitle                = nullptr;
        terminalOpenFocusDescription          = nullptr;
        terminalOpenFocusToggle               = nullptr;
    }
};
} // namespace

struct MousePane::DxCardState
{
    MouseDxPage page;

    void Detach() noexcept
    {
        page.Detach();
    }
};

MousePane::MousePane()  = default;
MousePane::~MousePane() = default;

void MousePane::OnVisibilityChanged(bool visible) noexcept
{
    static_cast<void>(visible);
}

void MousePane::Destroy(PreferencesDialogState& state) noexcept
{
    DetachDxCardHosts();
    static_cast<void>(state);
    _pageHost = nullptr;
}

void MousePane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost = parent;
    static_cast<void>(EnsureDxCardHosts(parent, state));
}

bool MousePane::EnsureDxCardHosts(HWND parent, PreferencesDialogState& state) noexcept
{
    _pageHostDx      = state.pageHostDxHost;
    _pageContentRoot = state.pageHostDxContentRootControl;
    if (! _pageHostDx || ! _pageContentRoot)
    {
        return false;
    }

    if (_dxCardState && PrefsUi::HasRetainedDxChildren(_pageContentRoot))
    {
        ApplyDxTheme(state);
        SyncDxControlsFromState(state);
        return true;
    }

    auto dxState = std::make_unique<DxCardState>();
    _pageHostDx->ResetInteractionState();
    _pageContentRoot->ClearChildren();

    Panel* const root = _pageContentRoot;
    dxState->page.paneFocusHeader = root->AddChild<Label>();
    dxState->page.paneFocusHeader->SetFontRole(FontRole::Header);

    const auto addToggleCard = [&](CardPanel*& card, Label*& title, Label*& description, Toggle*& toggle, bool terminalOpenOnly) noexcept
    {
        card = root->AddChild<CardPanel>();
        title = root->AddChild<Label>();
        title->SetFontRole(FontRole::Body);
        description = root->AddChild<Label>();
        description->SetFontRole(FontRole::Small);
        description->SetMultiline(true);
        toggle = root->AddChild<Toggle>();
        toggle->SetStateLabels(LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
        toggle->SetOnToggled([this, host = parent, terminalOpenOnly](bool checked) noexcept
        {
            if (_syncingToggles || ! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* dialogState = PrefsUi::GetDialogState(host);
            if (! dialogState)
            {
                return;
            }

            Common::Settings::MouseSettings mouse = GetMouseSettings(dialogState->workingSettings);
            if (terminalOpenOnly)
            {
                mouse.focusFollowsPointerWhenTerminalOpen = checked;
            }
            else
            {
                mouse.focusFollowsPointer = checked;
            }
            SetMouseSettings(*dialogState, mouse);
            if (const HWND dialog = GetAncestor(host, GA_ROOT); dialog && IsWindow(dialog) != FALSE)
            {
                SetDirty(dialog, *dialogState);
            }
            SyncDxControlsFromState(*dialogState);
        });
    };

    addToggleCard(dxState->page.focusFollowsPointerCard,
                  dxState->page.focusFollowsPointerTitle,
                  dxState->page.focusFollowsPointerDescription,
                  dxState->page.focusFollowsPointerToggle,
                  false);
    addToggleCard(dxState->page.terminalOpenFocusCard,
                  dxState->page.terminalOpenFocusTitle,
                  dxState->page.terminalOpenFocusDescription,
                  dxState->page.terminalOpenFocusToggle,
                  true);

    _dxCardState = std::move(dxState);
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    return true;
}

void MousePane::DetachDxCardHosts() noexcept
{
    if (_pageContentRoot && _pageHostDx && _pageHost && IsWindow(_pageHost) != FALSE)
    {
        _pageHostDx->ResetInteractionState();
        _pageContentRoot->ClearChildren();
    }
    _pageHostDx      = nullptr;
    _pageContentRoot = nullptr;

    if (_dxCardState)
    {
        _dxCardState->Detach();
        _dxCardState.reset();
    }
}

void MousePane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (_dxCardState && _pageHostDx)
    {
        _pageHostDx->SetTheme(MakeAppThemeDxPalette(state.theme));
    }
}

void MousePane::SyncDxControlsFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxCardState)
    {
        return;
    }

    MouseDxPage& page = _dxCardState->page;
    const std::wstring paneFocusHeader     = LoadStringResource(nullptr, IDS_PREFS_MOUSE_SECTION_PANE_FOCUS);
    const std::wstring focusTitle          = LoadStringResource(nullptr, IDS_PREFS_MOUSE_LABEL_FOCUS_FOLLOWS_POINTER);
    const std::wstring focusDescription    = LoadStringResource(nullptr, IDS_PREFS_MOUSE_DESC_FOCUS_FOLLOWS_POINTER);
    const std::wstring terminalTitle       = LoadStringResource(nullptr, IDS_PREFS_MOUSE_LABEL_FOCUS_FOLLOWS_POINTER_TERMINAL_OPEN);
    const std::wstring terminalDescription = LoadStringResource(nullptr, IDS_PREFS_MOUSE_DESC_FOCUS_FOLLOWS_POINTER_TERMINAL_OPEN);

    page.paneFocusHeader->SetText(paneFocusHeader);
    page.focusFollowsPointerTitle->SetText(focusTitle);
    page.focusFollowsPointerDescription->SetText(focusDescription);
    page.focusFollowsPointerToggle->SetAccessibleName(focusTitle);
    page.focusFollowsPointerToggle->SetAccessibleHelpText(focusDescription);
    page.terminalOpenFocusTitle->SetText(terminalTitle);
    page.terminalOpenFocusDescription->SetText(terminalDescription);
    page.terminalOpenFocusToggle->SetAccessibleName(terminalTitle);
    page.terminalOpenFocusToggle->SetAccessibleHelpText(terminalDescription);

    const Common::Settings::MouseSettings mouse = GetMouseSettings(state.workingSettings);
    _syncingToggles = true;
    page.focusFollowsPointerToggle->SetChecked(mouse.focusFollowsPointer);
    page.focusFollowsPointerToggle->SetEnabled(true);
    page.terminalOpenFocusToggle->SetChecked(mouse.focusFollowsPointerWhenTerminalOpen);
    page.terminalOpenFocusToggle->SetEnabled(true);
    _syncingToggles = false;

    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void MousePane::LayoutPage(HWND host,
                           PreferencesDialogState& state,
                           int x,
                           int& y,
                           int width,
                           int margin,
                           int gapY,
                           int sectionY,
                           const PreferencesTypographyContext& typography) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(margin);
    static_cast<void>(gapY);
    static_cast<void>(sectionY);

    if (! host || ! EnsureDxCardHosts(_pageHost ? _pageHost : host, state))
    {
        Debug::Error(L"Preferences.Mouse: DxUi surface initialization failed; page will not render correctly.");
        return;
    }

    Debug::Perf::Scope layoutPerf(L"preferences.ui.mouse_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(typography.dpi);

    const Common::Settings::UiSettings ui = state.workingSettings.ui.value_or(Common::Settings::UiSettings{});
    const bool compactMode                = ui.compactMode;
    const UINT dpi                        = (std::max<UINT>)(typography.dpi, USER_DEFAULT_SCREEN_DPI);
    const int rowHeight                   = UiMetrics::ScaleDip(dpi, compactMode ? 22 : kRowHeightDip);
    const int titleHeight                 = UiMetrics::ScaleDip(dpi, compactMode ? 16 : kTitleHeightDip);
    const int headerHeight                = UiMetrics::ScaleDip(dpi, compactMode ? 18 : kHeaderHeightDip);
    const int cardPaddingX                = UiMetrics::ScaleDip(dpi, compactMode ? 10 : kCardPaddingXDip);
    const int cardPaddingY                = UiMetrics::ScaleDip(dpi, compactMode ? 6 : kCardPaddingYDip);
    const int cardGapY                    = UiMetrics::ScaleDip(dpi, compactMode ? 1 : kCardGapYDip);
    const int cardGapX                    = UiMetrics::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY                = UiMetrics::ScaleDip(dpi, compactMode ? 6 : kCardSpacingYDip);
    const int minToggleWidth              = UiMetrics::ScaleDip(dpi, compactMode ? 82 : kMinToggleWidthDip);
    const int paddingX                    = UiMetrics::ScaleDip(dpi, kTogglePaddingXDip);
    const int toggleGapX                  = UiMetrics::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth                  = UiMetrics::ScaleDip(dpi, kToggleTrackWidthDip);
    const int onWidth = PrefsUi::MeasureSingleLineTextWidthPx(
        typography, typography.strong, LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
    const int offWidth = PrefsUi::MeasureSingleLineTextWidthPx(
        typography, typography.strong, LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF));
    const int toggleWidth = std::max(minToggleWidth, (2 * paddingX) + std::max(onWidth, offWidth) + toggleGapX + trackWidth);
    const auto pxToDip = [dpi](int pixels) noexcept { return static_cast<float>(pixels) * 96.0f / static_cast<float>(dpi); };

    MouseDxPage& page = _dxCardState->page;
    page.paneFocusHeader->SetVisible(true);
    page.paneFocusHeader->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + headerHeight)));
    y += headerHeight + cardSpacingY;

    const auto layoutToggleCard = [&](CardPanel* card, Label* title, Label* description, Toggle* toggle, const std::wstring& descriptionText) noexcept
    {
        const int textWidth        = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descriptionHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descriptionText);
        const int contentHeight    = titleHeight + cardGapY + descriptionHeight;
        const int cardHeight       = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);
        const int textLeft         = x + cardPaddingX;
        const int titleTop         = y + cardPaddingY;
        const int descriptionTop   = titleTop + titleHeight + cardGapY;
        const int toggleLeft       = x + width - cardPaddingX - toggleWidth;
        const int toggleTop        = y + (cardHeight - rowHeight) / 2;

        card->SetVisible(true);
        card->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + cardHeight)));
        title->SetVisible(true);
        title->SetBounds(D2D1::RectF(pxToDip(textLeft), pxToDip(titleTop), pxToDip(textLeft + textWidth), pxToDip(titleTop + titleHeight)));
        description->SetVisible(true);
        description->SetBounds(
            D2D1::RectF(pxToDip(textLeft), pxToDip(descriptionTop), pxToDip(textLeft + textWidth), pxToDip(descriptionTop + descriptionHeight)));
        toggle->SetVisible(true);
        toggle->SetBounds(D2D1::RectF(pxToDip(toggleLeft), pxToDip(toggleTop), pxToDip(toggleLeft + toggleWidth), pxToDip(toggleTop + rowHeight)));
        y += cardHeight + cardSpacingY;
    };

    layoutToggleCard(page.focusFollowsPointerCard,
                     page.focusFollowsPointerTitle,
                     page.focusFollowsPointerDescription,
                     page.focusFollowsPointerToggle,
                     LoadStringResource(nullptr, IDS_PREFS_MOUSE_DESC_FOCUS_FOLLOWS_POINTER));
    layoutToggleCard(page.terminalOpenFocusCard,
                     page.terminalOpenFocusTitle,
                     page.terminalOpenFocusDescription,
                     page.terminalOpenFocusToggle,
                     LoadStringResource(nullptr, IDS_PREFS_MOUSE_DESC_FOCUS_FOLLOWS_POINTER_TERMINAL_OPEN));
    _pageHostDx->Invalidate();
}

#ifdef ENABLE_TESTS
bool MousePane::DebugSetFocusFollowsPointerSettings(bool always, bool whenTerminalOpen) noexcept
{
    if (! _dxCardState || ! _pageHost || IsWindow(_pageHost) == FALSE)
    {
        return false;
    }
    auto* state = PrefsUi::GetDialogState(_pageHost);
    if (! state)
    {
        return false;
    }

    SetMouseSettings(*state,
                     Common::Settings::MouseSettings{
                         .focusFollowsPointer = always,
                         .focusFollowsPointerWhenTerminalOpen = whenTerminalOpen,
                     });
    if (const HWND dialog = GetAncestor(_pageHost, GA_ROOT); dialog && IsWindow(dialog) != FALSE)
    {
        SetDirty(dialog, *state);
    }
    SyncDxControlsFromState(*state);
    return _dxCardState->page.focusFollowsPointerToggle->IsChecked() == always &&
        _dxCardState->page.terminalOpenFocusToggle->IsChecked() == whenTerminalOpen;
}

bool MousePane::DebugGetFocusFollowsPointerSettings(Common::Settings::MouseSettings& outSettings) const noexcept
{
    if (! _pageHost || IsWindow(_pageHost) == FALSE)
    {
        return false;
    }
    const auto* state = PrefsUi::GetDialogState(_pageHost);
    if (! state)
    {
        return false;
    }
    outSettings = GetMouseSettings(state->workingSettings);
    return true;
}
#endif
