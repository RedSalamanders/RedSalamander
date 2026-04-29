// Preferences.General.cpp

#include "Framework.h"

#include "Preferences.General.h"

#include "DxUi/DxUi.h"
#include "Helpers.h"
#include "LocalizationManager.h"
#include "SettingsHotReload.h"
#include "UiMetrics.h"

#include "resource.h"

namespace
{
using RedSalamander::DxUi::CardPanel;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;
using RedSalamander::DxUi::WindowHost;

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

[[nodiscard]] Common::Settings::UiSettings GetUiSettings(const Common::Settings::Settings& settings) noexcept
{
    if (settings.ui.has_value())
    {
        return settings.ui.value();
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

void UpdateUiFromReducedMotionSelection(PreferencesDialogState& state, Common::Settings::ReducedMotionMode reducedMotion) noexcept
{
    Common::Settings::UiSettings ui = GetUiSettings(state.workingSettings);
    ui.reducedMotion                = reducedMotion;
    state.workingSettings.ui        = ui;
}

void UpdateUiFromCompactModeToggle(PreferencesDialogState& state, bool compactMode) noexcept
{
    Common::Settings::UiSettings ui = GetUiSettings(state.workingSettings);
    ui.compactMode                  = compactMode;
    state.workingSettings.ui        = ui;
}

void UpdateUiFromLanguageSelection(PreferencesDialogState& state, std::wstring language) noexcept
{
    Common::Settings::UiSettings ui = GetUiSettings(state.workingSettings);
    ui.language                     = language.empty() ? L"system" : std::move(language);
    state.workingSettings.ui        = std::move(ui);
}

void UpdateUiFromWindowBackdropSelection(PreferencesDialogState& state, Common::Settings::WindowBackdropMode windowBackdrop) noexcept
{
    Common::Settings::UiSettings ui = GetUiSettings(state.workingSettings);
    ui.windowBackdrop               = windowBackdrop;
    state.workingSettings.ui        = ui;
}

[[nodiscard]] std::vector<ComboBox::Item> BuildLanguageComboItems()
{
    std::vector<ComboBox::Item> items;
    items.push_back({L"system", LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_LANGUAGE_SYSTEM)});
    items.push_back({L"en", LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_LANGUAGE_ENGLISH)});

    for (const std::wstring& culture : Localization::DiscoverAvailableCultures())
    {
        if (culture.empty() || culture == L"en")
        {
            continue;
        }

        const bool alreadyAdded =
            std::find_if(items.begin(), items.end(), [&](const ComboBox::Item& item) noexcept { return item.value == culture; }) != items.end();
        if (! alreadyAdded)
        {
            items.push_back({culture, culture});
        }
    }

    return items;
}

struct GeneralDxPage
{
    GeneralDxPage()                                = default;
    GeneralDxPage(const GeneralDxPage&)            = delete;
    GeneralDxPage& operator=(const GeneralDxPage&) = delete;
    GeneralDxPage(GeneralDxPage&&)                 = delete;
    GeneralDxPage& operator=(GeneralDxPage&&)      = delete;

    Label* displayHeader             = nullptr;
    CardPanel* menuBarCard           = nullptr;
    Label* menuBarTitle              = nullptr;
    Label* menuBarDescription        = nullptr;
    Toggle* menuBarToggle            = nullptr;
    CardPanel* functionBarCard       = nullptr;
    Label* functionBarTitle          = nullptr;
    Label* functionBarDescription    = nullptr;
    Toggle* functionBarToggle        = nullptr;
    CardPanel* languageCard          = nullptr;
    Label* languageTitle             = nullptr;
    Label* languageDescription       = nullptr;
    ComboBox* languageCombo          = nullptr;
    Label* dxUiHeader                = nullptr;
    CardPanel* compactModeCard       = nullptr;
    Label* compactModeTitle          = nullptr;
    Label* compactModeDescription    = nullptr;
    Toggle* compactModeToggle        = nullptr;
    CardPanel* reducedMotionCard     = nullptr;
    Label* reducedMotionTitle        = nullptr;
    Label* reducedMotionDescription  = nullptr;
    ComboBox* reducedMotionCombo     = nullptr;
    CardPanel* windowBackdropCard    = nullptr;
    Label* windowBackdropTitle       = nullptr;
    Label* windowBackdropDescription = nullptr;
    ComboBox* windowBackdropCombo    = nullptr;
    Label* startupHeader             = nullptr;
    CardPanel* splashScreenCard      = nullptr;
    Label* splashScreenTitle         = nullptr;
    Label* splashScreenDescription   = nullptr;
    Toggle* splashScreenToggle       = nullptr;

    void Detach() noexcept
    {
        displayHeader             = nullptr;
        menuBarCard               = nullptr;
        menuBarTitle              = nullptr;
        menuBarDescription        = nullptr;
        menuBarToggle             = nullptr;
        functionBarCard           = nullptr;
        functionBarTitle          = nullptr;
        functionBarDescription    = nullptr;
        functionBarToggle         = nullptr;
        languageCard              = nullptr;
        languageTitle             = nullptr;
        languageDescription       = nullptr;
        languageCombo             = nullptr;
        dxUiHeader                = nullptr;
        compactModeCard           = nullptr;
        compactModeTitle          = nullptr;
        compactModeDescription    = nullptr;
        compactModeToggle         = nullptr;
        reducedMotionCard         = nullptr;
        reducedMotionTitle        = nullptr;
        reducedMotionDescription  = nullptr;
        reducedMotionCombo        = nullptr;
        windowBackdropCard        = nullptr;
        windowBackdropTitle       = nullptr;
        windowBackdropDescription = nullptr;
        windowBackdropCombo       = nullptr;
        startupHeader             = nullptr;
        splashScreenCard          = nullptr;
        splashScreenTitle         = nullptr;
        splashScreenDescription   = nullptr;
        splashScreenToggle        = nullptr;
    }
};

} // namespace

void RefreshPreferencesDialogTheme(HWND dlg, PreferencesDialogState& state) noexcept;

struct GeneralPane::DxCardState
{
    DxCardState()                              = default;
    DxCardState(const DxCardState&)            = delete;
    DxCardState& operator=(const DxCardState&) = delete;
    DxCardState(DxCardState&&)                 = delete;
    DxCardState& operator=(DxCardState&&)      = delete;

    GeneralDxPage page;

    void Detach() noexcept
    {
        page.Detach();
    }
};

GeneralPane::GeneralPane()  = default;
GeneralPane::~GeneralPane() = default;

void GeneralPane::OnVisibilityChanged(bool visible) noexcept
{
    static_cast<void>(visible);
}

void GeneralPane::Destroy(PreferencesDialogState& state) noexcept
{
    DetachDxCardHosts();
    static_cast<void>(state);
    _pageHost = nullptr;
}

bool GeneralPane::EnsureDxCardHosts(HWND parent, PreferencesDialogState& state) noexcept
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

    auto* root = _pageContentRoot;

    const auto addHeader = [&](Label*& outHeader) noexcept
    {
        outHeader = root->AddChild<Label>();
        outHeader->SetFontRole(FontRole::Header);
    };

    const auto addToggleCard = [&](CardPanel*& outCard, Label*& outTitle, Label*& outDescription) noexcept
    {
        outCard  = root->AddChild<CardPanel>();
        outTitle = root->AddChild<Label>();
        outTitle->SetFontRole(FontRole::Body);
        outDescription = root->AddChild<Label>();
        outDescription->SetFontRole(FontRole::Small);
        outDescription->SetMultiline(true);
    };

    const auto addToggle = [&](Toggle*& outToggle, const UINT commandId) noexcept
    {
        outToggle = root->AddChild<Toggle>();
        outToggle->SetStateLabels(LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
        outToggle->SetOnToggled([this, host = parent, commandId](bool checked) noexcept
        {
            if (! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* state = PrefsUi::GetDialogState(host);
            if (! state)
            {
                return;
            }

            switch (commandId)
            {
                case IDC_PREFS_GENERAL_MENUBAR_TOGGLE:
                {
                    const Common::Settings::MainMenuState menu = GetMainMenuState(state->workingSettings);
                    UpdateMainMenuFromToggle(*state, checked, menu.functionBarVisible);
                    break;
                }
                case IDC_PREFS_GENERAL_FUNCTIONBAR_TOGGLE:
                {
                    const Common::Settings::MainMenuState menu = GetMainMenuState(state->workingSettings);
                    UpdateMainMenuFromToggle(*state, menu.menuBarVisible, checked);
                    break;
                }
                case IDC_PREFS_GENERAL_COMPACT_MODE_TOGGLE: UpdateUiFromCompactModeToggle(*state, checked); break;
                case IDC_PREFS_GENERAL_SPLASH_TOGGLE: UpdateStartupFromToggle(*state, checked); break;
                default: return;
            }

            if (const HWND dlg = GetParent(host); dlg && IsWindow(dlg) != FALSE)
            {
                SetDirty(dlg, *state);
                if (commandId == IDC_PREFS_GENERAL_COMPACT_MODE_TOGGLE)
                {
                    SettingsHotReload::ApplyUiPreferencesToTheme(state->workingSettings, state->theme);
                    RefreshPreferencesDialogTheme(dlg, *state);
                }
            }
            else if (commandId == IDC_PREFS_GENERAL_COMPACT_MODE_TOGGLE)
            {
                SettingsHotReload::ApplyUiPreferencesToTheme(state->workingSettings, state->theme);
            }

            SyncDxControlsFromState(*state);
        });
    };

    const auto addCombo = [&](ComboBox*& outCombo, bool& syncFlag, const std::vector<ComboBox::Item>& items, const UINT commandId) noexcept
    {
        outCombo = root->AddChild<ComboBox>();
        outCombo->SetVariant(ComboBoxVariant::Window);
        outCombo->SetItems(items);
        outCombo->SetOnSelectionChanged([this, host = parent, &syncFlag, commandId](size_t itemIndex) noexcept
        {
            if (syncFlag || ! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* state = PrefsUi::GetDialogState(host);
            if (! state)
            {
                return;
            }

            switch (commandId)
            {
                case IDC_PREFS_GENERAL_REDUCED_MOTION_COMBO:
                    switch (itemIndex)
                    {
                        case 0: UpdateUiFromReducedMotionSelection(*state, Common::Settings::ReducedMotionMode::System); break;
                        case 1: UpdateUiFromReducedMotionSelection(*state, Common::Settings::ReducedMotionMode::Off); break;
                        case 2: UpdateUiFromReducedMotionSelection(*state, Common::Settings::ReducedMotionMode::On); break;
                        default: return;
                    }
                    break;
                case IDC_PREFS_GENERAL_WINDOW_BACKDROP_COMBO:
                    switch (itemIndex)
                    {
                        case 0: UpdateUiFromWindowBackdropSelection(*state, Common::Settings::WindowBackdropMode::Default); break;
                        case 1: UpdateUiFromWindowBackdropSelection(*state, Common::Settings::WindowBackdropMode::None); break;
                        case 2: UpdateUiFromWindowBackdropSelection(*state, Common::Settings::WindowBackdropMode::Mica); break;
                        case 3: UpdateUiFromWindowBackdropSelection(*state, Common::Settings::WindowBackdropMode::MicaAlt); break;
                        case 4: UpdateUiFromWindowBackdropSelection(*state, Common::Settings::WindowBackdropMode::Acrylic); break;
                        default: return;
                    }
                    break;
                case IDC_PREFS_GENERAL_LANGUAGE_COMBO:
                {
                    ComboBox* const combo = _dxCardState ? _dxCardState->page.languageCombo : nullptr;
                    if (! combo)
                    {
                        return;
                    }
                    const auto& items = combo->GetItems();
                    if (itemIndex >= items.size())
                    {
                        return;
                    }
                    UpdateUiFromLanguageSelection(*state, items[itemIndex].value);
                    break;
                }
                default: return;
            }

            if (const HWND dlg = GetParent(host); dlg && IsWindow(dlg) != FALSE)
            {
                SetDirty(dlg, *state);
                if (commandId == IDC_PREFS_GENERAL_LANGUAGE_COMBO)
                {
                    SyncDxControlsFromState(*state);
                }
                else
                {
                    SettingsHotReload::ApplyUiPreferencesToTheme(state->workingSettings, state->theme);
                    RefreshPreferencesDialogTheme(dlg, *state);
                }
            }
            else
            {
                SyncDxControlsFromState(*state);
            }
        });
    };

    addHeader(dxState->page.displayHeader);
    addToggleCard(dxState->page.menuBarCard, dxState->page.menuBarTitle, dxState->page.menuBarDescription);
    addToggle(dxState->page.menuBarToggle, IDC_PREFS_GENERAL_MENUBAR_TOGGLE);
    addToggleCard(dxState->page.functionBarCard, dxState->page.functionBarTitle, dxState->page.functionBarDescription);
    addToggle(dxState->page.functionBarToggle, IDC_PREFS_GENERAL_FUNCTIONBAR_TOGGLE);
    addToggleCard(dxState->page.languageCard, dxState->page.languageTitle, dxState->page.languageDescription);
    addCombo(dxState->page.languageCombo, _syncingLanguageCombo, BuildLanguageComboItems(), IDC_PREFS_GENERAL_LANGUAGE_COMBO);

    addHeader(dxState->page.dxUiHeader);
    addToggleCard(dxState->page.compactModeCard, dxState->page.compactModeTitle, dxState->page.compactModeDescription);
    addToggle(dxState->page.compactModeToggle, IDC_PREFS_GENERAL_COMPACT_MODE_TOGGLE);
    addToggleCard(dxState->page.reducedMotionCard, dxState->page.reducedMotionTitle, dxState->page.reducedMotionDescription);
    addCombo(dxState->page.reducedMotionCombo,
             _syncingReducedMotionCombo,
             {
                 {L"system", LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_REDUCED_MOTION_SYSTEM)},
                 {L"on", LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_REDUCED_MOTION_ON)},
                 {L"off", LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_REDUCED_MOTION_OFF)},
             },
             IDC_PREFS_GENERAL_REDUCED_MOTION_COMBO);
    addToggleCard(dxState->page.windowBackdropCard, dxState->page.windowBackdropTitle, dxState->page.windowBackdropDescription);
    addCombo(dxState->page.windowBackdropCombo,
             _syncingWindowBackdropCombo,
             {
                 {L"default", LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_DEFAULT)},
                 {L"none", LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_NONE)},
                 {L"mica", LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_MICA)},
                 {L"micaAlt", LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_MICA_ALT)},
                 {L"acrylic", LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_ACRYLIC)},
             },
             IDC_PREFS_GENERAL_WINDOW_BACKDROP_COMBO);

    addHeader(dxState->page.startupHeader);
    addToggleCard(dxState->page.splashScreenCard, dxState->page.splashScreenTitle, dxState->page.splashScreenDescription);
    addToggle(dxState->page.splashScreenToggle, IDC_PREFS_GENERAL_SPLASH_TOGGLE);

    _dxCardState = std::move(dxState);
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    return true;
}

void GeneralPane::DetachDxCardHosts() noexcept
{
    _usesDxUiTypographyContext = false;
    _usesDxUiTypographyMetrics = false;

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

void GeneralPane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (! _dxCardState || ! _pageHostDx)
    {
        return;
    }

    const ThemePalette palette = PrefsUi::MakeDxPalette(state.theme);
    _pageHostDx->SetTheme(palette);
}

void GeneralPane::SyncDxControlsFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxCardState)
    {
        return;
    }

    auto& page = _dxCardState->page;

    if (page.displayHeader)
    {
        page.displayHeader->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_SECTION_DISPLAY));
    }
    if (page.dxUiHeader)
    {
        page.dxUiHeader->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_SECTION_DXUI));
    }
    if (page.startupHeader)
    {
        page.startupHeader->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_SECTION_STARTUP));
    }

    if (page.menuBarTitle)
    {
        page.menuBarTitle->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_LABEL_MENU_BAR));
    }
    if (page.menuBarDescription)
    {
        page.menuBarDescription->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_MENU_BAR));
    }
    if (page.functionBarTitle)
    {
        page.functionBarTitle->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_LABEL_FUNCTION_BAR));
    }
    if (page.functionBarDescription)
    {
        page.functionBarDescription->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_FUNCTION_BAR));
    }
    if (page.languageTitle)
    {
        page.languageTitle->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_LABEL_LANGUAGE));
    }
    if (page.languageDescription)
    {
        page.languageDescription->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_LANGUAGE));
    }
    if (page.splashScreenTitle)
    {
        page.splashScreenTitle->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_LABEL_SPLASH_SCREEN));
    }
    if (page.splashScreenDescription)
    {
        page.splashScreenDescription->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_SPLASH_SCREEN));
    }
    if (page.reducedMotionTitle)
    {
        page.reducedMotionTitle->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_LABEL_REDUCED_MOTION));
    }
    if (page.reducedMotionDescription)
    {
        page.reducedMotionDescription->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_REDUCED_MOTION));
    }
    if (page.windowBackdropTitle)
    {
        page.windowBackdropTitle->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_LABEL_WINDOW_BACKDROP));
    }
    if (page.windowBackdropDescription)
    {
        page.windowBackdropDescription->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_WINDOW_BACKDROP));
    }
    if (page.compactModeTitle)
    {
        page.compactModeTitle->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_LABEL_COMPACT_MODE));
    }
    if (page.compactModeDescription)
    {
        page.compactModeDescription->SetText(LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_COMPACT_MODE));
    }

    const Common::Settings::MainMenuState menu      = GetMainMenuState(state.workingSettings);
    const Common::Settings::StartupSettings startup = GetStartupSettings(state.workingSettings);
    const Common::Settings::UiSettings ui           = GetUiSettings(state.workingSettings);
    if (page.menuBarToggle)
    {
        page.menuBarToggle->SetChecked(menu.menuBarVisible);
        page.menuBarToggle->SetEnabled(true);
    }
    if (page.functionBarToggle)
    {
        page.functionBarToggle->SetChecked(menu.functionBarVisible);
        page.functionBarToggle->SetEnabled(true);
    }
    if (page.compactModeToggle)
    {
        page.compactModeToggle->SetChecked(ui.compactMode);
        page.compactModeToggle->SetEnabled(true);
    }
    if (page.splashScreenToggle)
    {
        page.splashScreenToggle->SetChecked(startup.showSplash);
        page.splashScreenToggle->SetEnabled(true);
    }

    const auto syncCombo = [](ComboBox* combo, std::wstring_view targetValue, bool& syncFlag) noexcept
    {
        if (! combo)
        {
            return;
        }

        syncFlag          = true;
        const auto& items = combo->GetItems();
        std::optional<size_t> index{};
        for (size_t i = 0; i < items.size(); ++i)
        {
            if (items[i].value == targetValue)
            {
                index = i;
                break;
            }
        }
        combo->SetSelectedIndex(index);
        combo->SetEnabled(true);
        syncFlag = false;
    };

    syncCombo(page.languageCombo, ui.language.empty() ? L"system" : ui.language, _syncingLanguageCombo);

    switch (ui.reducedMotion)
    {
        case Common::Settings::ReducedMotionMode::On: syncCombo(page.reducedMotionCombo, L"off", _syncingReducedMotionCombo); break;
        case Common::Settings::ReducedMotionMode::Off: syncCombo(page.reducedMotionCombo, L"on", _syncingReducedMotionCombo); break;
        case Common::Settings::ReducedMotionMode::System:
        default: syncCombo(page.reducedMotionCombo, L"system", _syncingReducedMotionCombo); break;
    }

    switch (ui.windowBackdrop)
    {
        case Common::Settings::WindowBackdropMode::None: syncCombo(page.windowBackdropCombo, L"none", _syncingWindowBackdropCombo); break;
        case Common::Settings::WindowBackdropMode::Mica: syncCombo(page.windowBackdropCombo, L"mica", _syncingWindowBackdropCombo); break;
        case Common::Settings::WindowBackdropMode::MicaAlt: syncCombo(page.windowBackdropCombo, L"micaAlt", _syncingWindowBackdropCombo); break;
        case Common::Settings::WindowBackdropMode::Acrylic: syncCombo(page.windowBackdropCombo, L"acrylic", _syncingWindowBackdropCombo); break;
        case Common::Settings::WindowBackdropMode::Default:
        default: syncCombo(page.windowBackdropCombo, L"default", _syncingWindowBackdropCombo); break;
    }

    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void GeneralPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    static_cast<void>(host);
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
}

void GeneralPane::LayoutDxPage(HWND host, PreferencesDialogState& state, int x, int& y, int width, const PreferencesTypographyContext& typography) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(host);

    Debug::Perf::Scope layoutPerf(L"preferences.ui.general_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(typography.dpi);

    _usesDxUiTypographyContext = true;
    _usesDxUiTypographyMetrics = false;

    const Common::Settings::UiSettings ui = GetUiSettings(state.workingSettings);
    const bool compactMode                = ui.compactMode;
    const UINT dpi                        = (std::max<UINT>)(typography.dpi, USER_DEFAULT_SCREEN_DPI);
    const int rowHeightDip                = compactMode ? 22 : kRowHeightDip;
    const int titleHeightDip              = compactMode ? 16 : kTitleHeightDip;
    const int headerHeightDip             = compactMode ? 18 : kHeaderHeightDip;
    const int minToggleWidthDip           = compactMode ? 82 : kMinToggleWidthDip;
    const int sectionSpacingYDip          = compactMode ? 12 : kSectionSpacingYDip;
    const int cardPaddingXDip             = compactMode ? 10 : kCardPaddingXDip;
    const int cardPaddingYDip             = compactMode ? 6 : kCardPaddingYDip;
    const int cardGapYDip                 = compactMode ? 1 : kCardGapYDip;
    const int cardSpacingYDip             = compactMode ? 6 : kCardSpacingYDip;
    const int rowHeight                   = std::max(1, UiMetrics::ScaleDip(dpi, rowHeightDip));
    const int titleHeight                 = std::max(1, UiMetrics::ScaleDip(dpi, titleHeightDip));
    const int headerHeight                = std::max(1, UiMetrics::ScaleDip(dpi, headerHeightDip));
    const int minToggleWidth              = UiMetrics::ScaleDip(dpi, minToggleWidthDip);
    const int sectionSpacingY             = UiMetrics::ScaleDip(dpi, sectionSpacingYDip);

    const int cardPaddingX  = UiMetrics::ScaleDip(dpi, cardPaddingXDip);
    const int cardPaddingY  = UiMetrics::ScaleDip(dpi, cardPaddingYDip);
    const int cardGapY      = UiMetrics::ScaleDip(dpi, cardGapYDip);
    const int cardGapX      = UiMetrics::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY  = UiMetrics::ScaleDip(dpi, cardSpacingYDip);
    const int minComboWidth = UiMetrics::ScaleDip(dpi, kMinEditWidthDip + 10);
    const int maxComboWidth = UiMetrics::ScaleDip(dpi, kMaxEditWidthDip);

    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

    const int onWidth          = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, onLabel);
    const int offWidth         = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, offLabel);
    _usesDxUiTypographyMetrics = onWidth > 0 || offWidth > 0;

    const int paddingX       = UiMetrics::ScaleDip(dpi, kTogglePaddingXDip);
    const int gapX           = UiMetrics::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth     = UiMetrics::ScaleDip(dpi, kToggleTrackWidthDip);
    const int stateTextWidth = std::max(onWidth, offWidth);

    const int measuredToggleWidth = std::max(minToggleWidth, (2 * paddingX) + stateTextWidth + gapX + trackWidth);
    const int toggleWidth         = static_cast<int>(RedSalamander::DxUi::ResolveConstrainedExtent(
        {.minExtent = static_cast<float>(minToggleWidth), .preferredExtent = static_cast<float>(measuredToggleWidth)},
        static_cast<float>(std::max(0, width - 2 * cardPaddingX))));
    const int comboWidth          = std::min(std::max(minComboWidth, UiMetrics::ScaleDip(dpi, 140)), std::max(minComboWidth, maxComboWidth));

    DxCardState* dxState  = _dxCardState.get();
    GeneralDxPage* dxPage = (dxState && _pageHostDx && _pageContentRoot) ? &dxState->page : nullptr;
    const auto pxToDip    = [dpi](const int pixels) noexcept { return (static_cast<float>(pixels) * 96.0f) / static_cast<float>(dpi); };

    const auto layoutHeader = [&](Label* header, const std::wstring& text) noexcept
    {
        if (! header)
        {
            return;
        }

        header->SetVisible(true);
        header->SetText(text);
        header->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + headerHeight)));
        y += headerHeight + cardSpacingY;
    };

    const auto layoutToggleCard = [&](const std::wstring& descText, CardPanel* dxCard, Label* dxTitle, Label* dxDescription, Toggle* dxToggle) noexcept
    {
        const int textWidth        = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight       = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);
        _usesDxUiTypographyMetrics = _usesDxUiTypographyMetrics || descHeight > 0;
        const int contentHeight    = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight       = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        const int textLeft   = x + cardPaddingX;
        const int titleTop   = y + cardPaddingY;
        const int descTop    = y + cardPaddingY + titleHeight + cardGapY;
        const int toggleLeft = x + width - cardPaddingX - toggleWidth;
        const int toggleTop  = y + (cardHeight - rowHeight) / 2;
        if (dxCard)
        {
            dxCard->SetVisible(true);
            dxCard->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + cardHeight)));
        }
        if (dxTitle)
        {
            dxTitle->SetVisible(true);
            dxTitle->SetBounds(D2D1::RectF(pxToDip(textLeft), pxToDip(titleTop), pxToDip(textLeft + textWidth), pxToDip(titleTop + titleHeight)));
        }
        if (dxDescription)
        {
            dxDescription->SetVisible(true);
            dxDescription->SetBounds(D2D1::RectF(pxToDip(textLeft), pxToDip(descTop), pxToDip(textLeft + textWidth), pxToDip(descTop + descHeight)));
        }
        if (dxToggle)
        {
            dxToggle->SetVisible(true);
            dxToggle->SetBounds(D2D1::RectF(pxToDip(toggleLeft), pxToDip(toggleTop), pxToDip(toggleLeft + toggleWidth), pxToDip(toggleTop + rowHeight)));
        }

        y += cardHeight + cardSpacingY;
    };

    const auto layoutComboCard = [&](const std::wstring& descText, CardPanel* dxCard, Label* dxTitle, Label* dxDescription, ComboBox* dxCombo) noexcept
    {
        const int actualComboWidth = std::min(comboWidth, std::max(0, width - 2 * cardPaddingX));
        const int textWidth        = std::max(0, width - 2 * cardPaddingX - cardGapX - actualComboWidth);
        const int descHeight       = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);
        _usesDxUiTypographyMetrics = _usesDxUiTypographyMetrics || descHeight > 0;
        const int contentHeight    = std::max(rowHeight, titleHeight + cardGapY + descHeight);
        const int cardHeight       = contentHeight + (2 * cardPaddingY);
        const int titleTop         = y + cardPaddingY;
        const int descTop          = y + cardPaddingY + titleHeight + cardGapY;
        const int comboLeft        = x + width - cardPaddingX - actualComboWidth;
        const int comboTop         = y + (cardHeight - rowHeight) / 2;

        if (dxCard)
        {
            dxCard->SetVisible(true);
            dxCard->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + cardHeight)));
        }
        if (dxTitle)
        {
            dxTitle->SetVisible(true);
            dxTitle->SetBounds(
                D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(titleTop), pxToDip(x + cardPaddingX + textWidth), pxToDip(titleTop + titleHeight)));
        }
        if (dxDescription)
        {
            dxDescription->SetVisible(true);
            dxDescription->SetBounds(
                D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(descTop), pxToDip(x + cardPaddingX + textWidth), pxToDip(descTop + descHeight)));
        }
        if (dxCombo)
        {
            dxCombo->SetVisible(true);
            dxCombo->SetBounds(D2D1::RectF(pxToDip(comboLeft), pxToDip(comboTop), pxToDip(comboLeft + actualComboWidth), pxToDip(comboTop + rowHeight)));
        }

        y += cardHeight + cardSpacingY;
    };

    const std::wstring menuBarDesc        = LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_MENU_BAR);
    const std::wstring functionBarDesc    = LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_FUNCTION_BAR);
    const std::wstring languageDesc       = LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_LANGUAGE);
    const std::wstring compactModeDesc    = LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_COMPACT_MODE);
    const std::wstring reducedMotionDesc  = LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_REDUCED_MOTION);
    const std::wstring windowBackdropDesc = LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_WINDOW_BACKDROP);
    const std::wstring splashDesc         = LoadStringResource(nullptr, IDS_PREFS_GENERAL_DESC_SPLASH_SCREEN);

    layoutHeader(dxPage ? dxPage->displayHeader : nullptr, LoadStringResource(nullptr, IDS_PREFS_GENERAL_SECTION_DISPLAY));
    layoutToggleCard(menuBarDesc,
                     dxPage ? dxPage->menuBarCard : nullptr,
                     dxPage ? dxPage->menuBarTitle : nullptr,
                     dxPage ? dxPage->menuBarDescription : nullptr,
                     dxPage ? dxPage->menuBarToggle : nullptr);
    layoutToggleCard(functionBarDesc,
                     dxPage ? dxPage->functionBarCard : nullptr,
                     dxPage ? dxPage->functionBarTitle : nullptr,
                     dxPage ? dxPage->functionBarDescription : nullptr,
                     dxPage ? dxPage->functionBarToggle : nullptr);
    layoutComboCard(languageDesc,
                    dxPage ? dxPage->languageCard : nullptr,
                    dxPage ? dxPage->languageTitle : nullptr,
                    dxPage ? dxPage->languageDescription : nullptr,
                    dxPage ? dxPage->languageCombo : nullptr);
    y += sectionSpacingY;
    layoutHeader(dxPage ? dxPage->dxUiHeader : nullptr, LoadStringResource(nullptr, IDS_PREFS_GENERAL_SECTION_DXUI));
    layoutToggleCard(compactModeDesc,
                     dxPage ? dxPage->compactModeCard : nullptr,
                     dxPage ? dxPage->compactModeTitle : nullptr,
                     dxPage ? dxPage->compactModeDescription : nullptr,
                     dxPage ? dxPage->compactModeToggle : nullptr);
    layoutComboCard(reducedMotionDesc,
                    dxPage ? dxPage->reducedMotionCard : nullptr,
                    dxPage ? dxPage->reducedMotionTitle : nullptr,
                    dxPage ? dxPage->reducedMotionDescription : nullptr,
                    dxPage ? dxPage->reducedMotionCombo : nullptr);
    layoutComboCard(windowBackdropDesc,
                    dxPage ? dxPage->windowBackdropCard : nullptr,
                    dxPage ? dxPage->windowBackdropTitle : nullptr,
                    dxPage ? dxPage->windowBackdropDescription : nullptr,
                    dxPage ? dxPage->windowBackdropCombo : nullptr);
    y += sectionSpacingY;
    layoutHeader(dxPage ? dxPage->startupHeader : nullptr, LoadStringResource(nullptr, IDS_PREFS_GENERAL_SECTION_STARTUP));
    layoutToggleCard(splashDesc,
                     dxPage ? dxPage->splashScreenCard : nullptr,
                     dxPage ? dxPage->splashScreenTitle : nullptr,
                     dxPage ? dxPage->splashScreenDescription : nullptr,
                     dxPage ? dxPage->splashScreenToggle : nullptr);

    if (dxPage && _pageHostDx)
    {
        _pageHostDx->Invalidate();
    }

    static_cast<void>(state);
}

void GeneralPane::LayoutPage(HWND host, PreferencesDialogState& state, int x, int& y, int width, const PreferencesTypographyContext& typography) noexcept
{
    _usesDxUiTypographyContext = false;
    _usesDxUiTypographyMetrics = false;

    if (! host)
    {
        return;
    }

    if (EnsureDxCardHosts(_pageHost ? _pageHost : host, state))
    {
        LayoutDxPage(host, state, x, y, width, typography);
        return;
    }

    Debug::Error(L"Preferences.General: DxUi surface initialization failed; page will not render correctly.");
}

void GeneralPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost = parent;

    static_cast<void>(EnsureDxCardHosts(parent, state));
}

#ifdef ENABLE_TESTS
PreferencesGeneralDebugFocusTarget GeneralPane::DebugGetFocusTarget() const noexcept
{
    if (! _dxCardState || ! _pageHostDx)
    {
        return PreferencesGeneralDebugFocusTarget::None;
    }

    const auto& page                                   = _dxCardState->page;
    RedSalamander::DxUi::Control* const focusedControl = _pageHostDx->GetFocusControl();
    if (! focusedControl)
    {
        return PreferencesGeneralDebugFocusTarget::None;
    }

    if (focusedControl == page.menuBarToggle)
        return PreferencesGeneralDebugFocusTarget::MenuBarToggle;
    if (focusedControl == page.functionBarToggle)
        return PreferencesGeneralDebugFocusTarget::FunctionBarToggle;
    if (focusedControl == page.languageCombo)
        return PreferencesGeneralDebugFocusTarget::LanguageCombo;
    if (focusedControl == page.compactModeToggle)
        return PreferencesGeneralDebugFocusTarget::CompactModeToggle;
    if (focusedControl == page.reducedMotionCombo)
        return PreferencesGeneralDebugFocusTarget::ReducedMotionCombo;
    if (focusedControl == page.windowBackdropCombo)
        return PreferencesGeneralDebugFocusTarget::WindowBackdropCombo;
    if (focusedControl == page.splashScreenToggle)
        return PreferencesGeneralDebugFocusTarget::SplashScreenToggle;
    return PreferencesGeneralDebugFocusTarget::None;
}

bool GeneralPane::DebugUsesDxUiTypographyContext() const noexcept
{
    return _usesDxUiTypographyContext;
}

bool GeneralPane::DebugUsesDxUiTypographyMetrics() const noexcept
{
    return _usesDxUiTypographyMetrics;
}

bool GeneralPane::DebugFocusMenuBarToggle() noexcept
{
    if (! _dxCardState || ! _pageHostDx || ! _dxCardState->page.menuBarToggle)
    {
        return false;
    }

    _pageHostDx->SetFocusControl(_dxCardState->page.menuBarToggle);
    return true;
}

bool GeneralPane::DebugGetMenuBarToggleChecked(bool& outChecked) const noexcept
{
    if (! _dxCardState || ! _dxCardState->page.menuBarToggle)
    {
        return false;
    }

    outChecked = _dxCardState->page.menuBarToggle->IsChecked();
    return true;
}

bool GeneralPane::DebugGetCompactModeToggleHeightDip(float& outHeightDip) const noexcept
{
    outHeightDip = 0.0f;
    if (! _dxCardState || ! _dxCardState->page.compactModeToggle)
    {
        return false;
    }

    const D2D1_RECT_F bounds = _dxCardState->page.compactModeToggle->GetBounds();
    outHeightDip             = (std::max)(0.0f, bounds.bottom - bounds.top);
    return outHeightDip > 0.0f;
}

bool GeneralPane::DebugSetCompactMode(bool checked) noexcept
{
    if (! _dxCardState || ! _pageHostDx || ! _pageHost || IsWindow(_pageHost) == FALSE)
    {
        return false;
    }

    auto* state = PrefsUi::GetDialogState(_pageHost);
    if (! state)
    {
        return false;
    }

    UpdateUiFromCompactModeToggle(*state, checked);

    if (const HWND dlg = GetParent(_pageHost); dlg && IsWindow(dlg) != FALSE)
    {
        SetDirty(dlg, *state);
        SettingsHotReload::ApplyUiPreferencesToTheme(state->workingSettings, state->theme);
        RefreshPreferencesDialogTheme(dlg, *state);
    }

    return _dxCardState->page.compactModeToggle && _dxCardState->page.compactModeToggle->IsChecked() == checked;
}

bool GeneralPane::DebugSelectLanguageByText(std::wstring_view displayText) noexcept
{
    if (! _dxCardState || ! _pageHostDx || ! _pageHost || IsWindow(_pageHost) == FALSE)
    {
        return false;
    }

    auto* state = PrefsUi::GetDialogState(_pageHost);
    if (! state)
    {
        return false;
    }

    auto* combo = _dxCardState->page.languageCombo;
    if (! combo)
    {
        return false;
    }

    const auto& items = combo->GetItems();
    const auto it     = std::find_if(
        items.begin(), items.end(), [displayText](const ComboBox::Item& item) noexcept { return item.display == displayText || item.value == displayText; });
    if (it == items.end())
    {
        return false;
    }

    UpdateUiFromLanguageSelection(*state, it->value);

    if (const HWND dlg = GetParent(_pageHost); dlg && IsWindow(dlg) != FALSE)
    {
        SetDirty(dlg, *state);
        SyncDxControlsFromState(*state);
    }

    return combo->GetDisplayedText() == displayText || combo->GetDisplayedText() == it->display;
}

bool GeneralPane::DebugSelectReducedMotionByText(std::wstring_view displayText) noexcept
{
    if (! _dxCardState || ! _pageHostDx || ! _pageHost || IsWindow(_pageHost) == FALSE)
    {
        return false;
    }

    auto* state = PrefsUi::GetDialogState(_pageHost);
    if (! state)
    {
        return false;
    }

    if (displayText == LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_REDUCED_MOTION_SYSTEM))
    {
        UpdateUiFromReducedMotionSelection(*state, Common::Settings::ReducedMotionMode::System);
    }
    else if (displayText == LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_REDUCED_MOTION_ON))
    {
        UpdateUiFromReducedMotionSelection(*state, Common::Settings::ReducedMotionMode::Off);
    }
    else if (displayText == LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_REDUCED_MOTION_OFF))
    {
        UpdateUiFromReducedMotionSelection(*state, Common::Settings::ReducedMotionMode::On);
    }
    else
    {
        return false;
    }

    if (const HWND dlg = GetParent(_pageHost); dlg && IsWindow(dlg) != FALSE)
    {
        SetDirty(dlg, *state);
        SettingsHotReload::ApplyUiPreferencesToTheme(state->workingSettings, state->theme);
        RefreshPreferencesDialogTheme(dlg, *state);
    }

    auto* combo = _dxCardState->page.reducedMotionCombo;
    return combo && combo->GetDisplayedText() == displayText;
}

bool GeneralPane::DebugSelectWindowBackdropByText(std::wstring_view displayText) noexcept
{
    if (! _dxCardState || ! _pageHostDx || ! _pageHost || IsWindow(_pageHost) == FALSE)
    {
        return false;
    }

    auto* state = PrefsUi::GetDialogState(_pageHost);
    if (! state)
    {
        return false;
    }

    if (displayText == LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_DEFAULT))
    {
        UpdateUiFromWindowBackdropSelection(*state, Common::Settings::WindowBackdropMode::Default);
    }
    else if (displayText == LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_NONE))
    {
        UpdateUiFromWindowBackdropSelection(*state, Common::Settings::WindowBackdropMode::None);
    }
    else if (displayText == LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_MICA))
    {
        UpdateUiFromWindowBackdropSelection(*state, Common::Settings::WindowBackdropMode::Mica);
    }
    else if (displayText == LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_MICA_ALT))
    {
        UpdateUiFromWindowBackdropSelection(*state, Common::Settings::WindowBackdropMode::MicaAlt);
    }
    else if (displayText == LoadStringResource(nullptr, IDS_PREFS_GENERAL_OPTION_WINDOW_BACKDROP_ACRYLIC))
    {
        UpdateUiFromWindowBackdropSelection(*state, Common::Settings::WindowBackdropMode::Acrylic);
    }
    else
    {
        return false;
    }

    if (const HWND dlg = GetParent(_pageHost); dlg && IsWindow(dlg) != FALSE)
    {
        SetDirty(dlg, *state);
        SettingsHotReload::ApplyUiPreferencesToTheme(state->workingSettings, state->theme);
        RefreshPreferencesDialogTheme(dlg, *state);
    }

    auto* combo = _dxCardState->page.windowBackdropCombo;
    return combo && combo->GetDisplayedText() == displayText;
}
#endif
