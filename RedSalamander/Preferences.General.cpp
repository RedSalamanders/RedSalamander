// Preferences.General.cpp

#include "Framework.h"

#include "Preferences.General.h"

#include "Helpers.h"
#include "SettingsSchemaParser.h"
#include "ThemedControls.h"

#include "resource.h"

namespace
{
[[nodiscard]] Common::Settings::MainMenuState GetMainMenuState(const Common::Settings::Settings& settings) noexcept
{
    if (settings.mainMenu.has_value())
    {
        return settings.mainMenu.value();
    }
    return {};
}

[[nodiscard]] Common::Settings::StartupSettings GetStartupSettings(const Common::Settings::Settings& settings) noexcept
{
    if (settings.startup.has_value())
    {
        return settings.startup.value();
    }
    return {};
}

void UpdateMainMenuFromToggle(PreferencesDialogState& state, bool menuBarVisible, bool functionBarVisible) noexcept
{
    Common::Settings::MainMenuState menu = GetMainMenuState(state.workingSettings);
    menu.menuBarVisible                  = menuBarVisible;
    menu.functionBarVisible              = functionBarVisible;
    state.workingSettings.mainMenu       = menu;
}

void UpdateStartupFromToggle(PreferencesDialogState& state, bool showSplashScreen) noexcept
{
    Common::Settings::StartupSettings startup = GetStartupSettings(state.workingSettings);
    startup.showSplash                        = showSplashScreen;
    state.workingSettings.startup             = startup;
}
} // namespace

bool GeneralPane::EnsureCreated(HWND pageHost) noexcept
{
    return PrefsPaneHost::EnsureCreated(pageHost, _hWnd);
}

void GeneralPane::ResizeToHostClient(HWND pageHost) noexcept
{
    PrefsPaneHost::ResizeToHostClient(pageHost, _hWnd.get());
}

void GeneralPane::Show(bool visible) noexcept
{
    PrefsPaneHost::Show(_hWnd.get(), visible);
}

void GeneralPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    static_cast<void>(host);

    const Common::Settings::MainMenuState menu      = GetMainMenuState(state.workingSettings);
    const Common::Settings::StartupSettings startup = GetStartupSettings(state.workingSettings);
    PrefsUi::SetTwoStateToggleState(state.menuBarToggle.get(), state.theme.systemHighContrast, menu.menuBarVisible);
    PrefsUi::SetTwoStateToggleState(state.functionBarToggle.get(), state.theme.systemHighContrast, menu.functionBarVisible);
    PrefsUi::SetTwoStateToggleState(state.splashScreenToggle.get(), state.theme.systemHighContrast, startup.showSplash);
}

bool GeneralPane::HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept
{
    if (! host)
    {
        return false;
    }

    if (commandId != IDC_PREFS_GENERAL_MENUBAR_TOGGLE && commandId != IDC_PREFS_GENERAL_FUNCTIONBAR_TOGGLE && commandId != IDC_PREFS_GENERAL_SPLASH_TOGGLE)
    {
        return false;
    }

    if (notifyCode != BN_CLICKED)
    {
        return true;
    }

    if (! state.menuBarToggle || ! state.functionBarToggle || ! state.splashScreenToggle)
    {
        return true;
    }

    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return true;
    }

    if (! hwndCtl)
    {
        return true;
    }

    const bool ownerDraw = (GetWindowLongPtrW(hwndCtl, GWL_STYLE) & BS_TYPEMASK) == BS_OWNERDRAW;
    if (ownerDraw)
    {
        const bool toggledOn = PrefsUi::GetTwoStateToggleState(hwndCtl, false);
        PrefsUi::SetTwoStateToggleState(hwndCtl, false, ! toggledOn);
    }

    const bool menuBarVisible     = PrefsUi::GetTwoStateToggleState(state.menuBarToggle.get(), state.theme.systemHighContrast);
    const bool functionBarVisible = PrefsUi::GetTwoStateToggleState(state.functionBarToggle.get(), state.theme.systemHighContrast);
    const bool showSplashScreen   = PrefsUi::GetTwoStateToggleState(state.splashScreenToggle.get(), state.theme.systemHighContrast);
    UpdateMainMenuFromToggle(state, menuBarVisible, functionBarVisible);
    UpdateStartupFromToggle(state, showSplashScreen);
    SetDirty(dlg, state);
    return true;
}

void GeneralPane::LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, HFONT dialogFont) noexcept
{
    using namespace PrefsLayoutConstants;

    if (! host)
    {
        return;
    }

    const UINT dpi           = GetDpiForWindow(host);
    const int rowHeight      = std::max(1, ThemedControls::ScaleDip(dpi, kRowHeightDip));
    const int titleHeight    = std::max(1, ThemedControls::ScaleDip(dpi, kTitleHeightDip));
    const int minToggleWidth = ThemedControls::ScaleDip(dpi, kMinToggleWidthDip);

    const int cardPaddingX = ThemedControls::ScaleDip(dpi, kCardPaddingXDip);
    const int cardPaddingY = ThemedControls::ScaleDip(dpi, kCardPaddingYDip);
    const int cardGapY     = ThemedControls::ScaleDip(dpi, kCardGapYDip);
    const int cardGapX     = ThemedControls::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY = ThemedControls::ScaleDip(dpi, kCardSpacingYDip);

    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

    const HFONT toggleMeasureFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const int onWidth             = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, onLabel);
    const int offWidth            = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, offLabel);

    const int paddingX       = ThemedControls::ScaleDip(dpi, kTogglePaddingXDip);
    const int gapX           = ThemedControls::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth     = ThemedControls::ScaleDip(dpi, kToggleTrackWidthDip);
    const int stateTextWidth = std::max(onWidth, offWidth);

    const int measuredToggleWidth = std::max(minToggleWidth, (2 * paddingX) + stateTextWidth + gapX + trackWidth);
    const int toggleWidth         = std::min(std::max(0, width - 2 * cardPaddingX), measuredToggleWidth);

    const HFONT infoFont = state.italicFont ? state.italicFont.get() : dialogFont;

    auto layoutToggleCard = [&](HWND title, HWND toggle, HWND descLabel, const std::wstring& descText) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight = descLabel ? PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, descText) : 0;

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;

        state.pageSettingCards.push_back(card);

        if (title)
        {
            SetWindowPos(title, nullptr, card.left + cardPaddingX, card.top + cardPaddingY, textWidth, titleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(title, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        if (descLabel)
        {
            SetWindowTextW(descLabel, descText.c_str());
            SetWindowPos(descLabel,
                         nullptr,
                         card.left + cardPaddingX,
                         card.top + cardPaddingY + titleHeight + cardGapY,
                         textWidth,
                         std::max(0, descHeight),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(descLabel, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
        }

        if (toggle)
        {
            SetWindowPos(toggle,
                         nullptr,
                         card.right - cardPaddingX - toggleWidth,
                         card.top + (cardHeight - rowHeight) / 2,
                         toggleWidth,
                         rowHeight,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        y += cardHeight + cardSpacingY;
    };

    const std::wstring menuBarDesc     = LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_MENU_BAR);
    const std::wstring functionBarDesc = LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_FUNCTION_BAR);
    const std::wstring splashDesc      = LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_SPLASH_SCREEN);

    layoutToggleCard(state.menuBarLabel.get(), state.menuBarToggle.get(), state.menuBarDescription.get(), menuBarDesc);
    layoutToggleCard(state.functionBarLabel.get(), state.functionBarToggle.get(), state.functionBarDescription.get(), functionBarDesc);
    layoutToggleCard(state.splashScreenLabel.get(), state.splashScreenToggle.get(), state.splashScreenDescription.get(), splashDesc);

    // DEMO: Schema-driven UI generation (auto-generates controls from SettingsStore.schema.json)
    // This demonstrates the hybrid approach: existing handcrafted controls above,
    // schema-driven auto-generated controls below
    const auto generalFields = SettingsSchemaParser::GetNonCustomFieldsForPane(state.schemaFields, L"General");
    if (! generalFields.empty())
    {
        const int margin = ThemedControls::ScaleDip(dpi, 16);
        const int gapY   = ThemedControls::ScaleDip(dpi, 12);

        for (const auto& field : generalFields)
        {
            static_cast<void>(PrefsUi::CreateSchemaControl(host, field, state, x, y, width, margin, gapY, dialogFont));
        }
    }
}

void GeneralPane::CreateControls(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    const DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const DWORD wrapStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;

    state.menuBarLabel.reset(CreateWindowExW(0,
                                             L"Static",
                                             LoadStringResource(nullptr, IDS_PREFS_GENERAL_LABEL_MENU_BAR).c_str(),
                                             baseStaticStyle,
                                             0,
                                             0,
                                             10,
                                             10,
                                             parent,
                                             nullptr,
                                             GetModuleHandleW(nullptr),
                                             nullptr));

    state.menuBarDescription.reset(CreateWindowExW(0,
                                                   L"Static",
                                                   LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_MENU_BAR).c_str(),
                                                   wrapStaticStyle,
                                                   0,
                                                   0,
                                                   10,
                                                   10,
                                                   parent,
                                                   nullptr,
                                                   GetModuleHandleW(nullptr),
                                                   nullptr));

    state.functionBarLabel.reset(CreateWindowExW(0,
                                                 L"Static",
                                                 LoadStringResource(nullptr, IDS_PREFS_GENERAL_LABEL_FUNCTION_BAR).c_str(),
                                                 baseStaticStyle,
                                                 0,
                                                 0,
                                                 10,
                                                 10,
                                                 parent,
                                                 nullptr,
                                                 GetModuleHandleW(nullptr),
                                                 nullptr));

    state.functionBarDescription.reset(CreateWindowExW(0,
                                                       L"Static",
                                                       LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_FUNCTION_BAR).c_str(),
                                                       wrapStaticStyle,
                                                       0,
                                                       0,
                                                       10,
                                                       10,
                                                       parent,
                                                       nullptr,
                                                       GetModuleHandleW(nullptr),
                                                       nullptr));

    state.splashScreenLabel.reset(CreateWindowExW(0,
                                                  L"Static",
                                                  LoadStringResource(nullptr, IDS_PREFS_GENERAL_LABEL_SPLASH_SCREEN).c_str(),
                                                  baseStaticStyle,
                                                  0,
                                                  0,
                                                  10,
                                                  10,
                                                  parent,
                                                  nullptr,
                                                  GetModuleHandleW(nullptr),
                                                  nullptr));

    state.splashScreenDescription.reset(CreateWindowExW(0,
                                                        L"Static",
                                                        LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_SPLASH_SCREEN).c_str(),
                                                        wrapStaticStyle,
                                                        0,
                                                        0,
                                                        10,
                                                        10,
                                                        parent,
                                                        nullptr,
                                                        GetModuleHandleW(nullptr),
                                                        nullptr));

    const bool customButtons = ! state.theme.systemHighContrast;
    if (customButtons)
    {
        state.menuBarToggle.reset(CreateWindowExW(0,
                                                  L"Button",
                                                  L"",
                                                  WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                                  0,
                                                  0,
                                                  10,
                                                  10,
                                                  parent,
                                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_GENERAL_MENUBAR_TOGGLE)),
                                                  GetModuleHandleW(nullptr),
                                                  nullptr));
        state.functionBarToggle.reset(CreateWindowExW(0,
                                                      L"Button",
                                                      L"",
                                                      WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                                      0,
                                                      0,
                                                      10,
                                                      10,
                                                      parent,
                                                      reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_GENERAL_FUNCTIONBAR_TOGGLE)),
                                                      GetModuleHandleW(nullptr),
                                                      nullptr));
        state.splashScreenToggle.reset(CreateWindowExW(0,
                                                       L"Button",
                                                       L"",
                                                       WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                                       0,
                                                       0,
                                                       10,
                                                       10,
                                                       parent,
                                                       reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_GENERAL_SPLASH_TOGGLE)),
                                                       GetModuleHandleW(nullptr),
                                                       nullptr));
    }
    else
    {
        state.menuBarToggle.reset(CreateWindowExW(0,
                                                  L"Button",
                                                  LoadStringResource(nullptr, IDS_PREFS_GENERAL_CHECK_SHOW_MENU_BAR).c_str(),
                                                  WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                                                  0,
                                                  0,
                                                  10,
                                                  10,
                                                  parent,
                                                  reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_GENERAL_MENUBAR_TOGGLE)),
                                                  GetModuleHandleW(nullptr),
                                                  nullptr));
        state.functionBarToggle.reset(CreateWindowExW(0,
                                                      L"Button",
                                                      LoadStringResource(nullptr, IDS_PREFS_GENERAL_CHECK_SHOW_FUNCTION_BAR).c_str(),
                                                      WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                                                      0,
                                                      0,
                                                      10,
                                                      10,
                                                      parent,
                                                      reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_GENERAL_FUNCTIONBAR_TOGGLE)),
                                                      GetModuleHandleW(nullptr),
                                                      nullptr));
        state.splashScreenToggle.reset(CreateWindowExW(0,
                                                       L"Button",
                                                       LoadStringResource(nullptr, IDS_PREFS_GENERAL_CHECK_SHOW_SPLASH_SCREEN).c_str(),
                                                       WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                                                       0,
                                                       0,
                                                       10,
                                                       10,
                                                       parent,
                                                       reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_GENERAL_SPLASH_TOGGLE)),
                                                       GetModuleHandleW(nullptr),
                                                       nullptr));
    }

    PrefsInput::EnableMouseWheelForwarding(state.menuBarToggle.get());
    PrefsInput::EnableMouseWheelForwarding(state.functionBarToggle.get());
    PrefsInput::EnableMouseWheelForwarding(state.splashScreenToggle.get());
}
