// Preferences.Monitor.cpp

#include "Framework.h"

#include "Preferences.Monitor.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <string>
#include <vector>

#include "DxUi/DxUi.h"
#include "Helpers.h"
#include "UiMetrics.h"
#include "resource.h"

namespace
{
using PrefsMonitor::EnsureWorkingMonitorSettings;
using PrefsMonitor::GetMonitorSettingsOrDefault;
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ButtonVariant;
using RedSalamander::DxUi::CardPanel;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;

constexpr std::array<UINT, 11> kMonitorToggleLabelStringIds = {{
    IDS_PREFS_ADV_LABEL_TOOLBAR,
    IDS_PREFS_ADV_LABEL_LINE_NUMBERS,
    IDS_PREFS_ADV_LABEL_ALWAYS_ON_TOP,
    IDS_PREFS_ADV_LABEL_SHOW_IDS,
    IDS_PREFS_ADV_LABEL_AUTO_SCROLL,
    IDS_PREFS_ADV_LABEL_FILTER_TEXT,
    IDS_PREFS_ADV_LABEL_FILTER_ERROR,
    IDS_PREFS_ADV_LABEL_FILTER_WARNING,
    IDS_PREFS_ADV_LABEL_FILTER_INFO,
    IDS_PREFS_ADV_LABEL_FILTER_PERF,
    IDS_PREFS_ADV_LABEL_FILTER_DEBUG,
}};

constexpr std::array<UINT, 11> kMonitorToggleDescriptionStringIds = {{
    IDS_PREFS_ADV_DESC_TOOLBAR,
    IDS_PREFS_ADV_DESC_LINE_NUMBERS,
    IDS_PREFS_ADV_DESC_ALWAYS_ON_TOP,
    IDS_PREFS_ADV_DESC_SHOW_IDS,
    IDS_PREFS_ADV_DESC_AUTO_SCROLL,
    IDS_PREFS_ADV_DESC_FILTER_TEXT,
    IDS_PREFS_ADV_DESC_FILTER_ERROR,
    IDS_PREFS_ADV_DESC_FILTER_WARNING,
    IDS_PREFS_ADV_DESC_FILTER_INFO,
    IDS_PREFS_ADV_DESC_FILTER_PERF,
    IDS_PREFS_ADV_DESC_FILTER_DEBUG,
}};

constexpr std::array<UINT, 11> kMonitorToggleCommandIds = {{
    IDC_PREFS_MONITOR_TOOLBAR_TOGGLE,
    IDC_PREFS_MONITOR_LINE_NUMBERS_TOGGLE,
    IDC_PREFS_MONITOR_ALWAYS_ON_TOP_TOGGLE,
    IDC_PREFS_MONITOR_SHOW_IDS_TOGGLE,
    IDC_PREFS_MONITOR_AUTO_SCROLL_TOGGLE,
    IDC_PREFS_MONITOR_FILTER_TEXT_TOGGLE,
    IDC_PREFS_MONITOR_FILTER_ERROR_TOGGLE,
    IDC_PREFS_MONITOR_FILTER_WARNING_TOGGLE,
    IDC_PREFS_MONITOR_FILTER_INFO_TOGGLE,
    IDC_PREFS_MONITOR_FILTER_PERF_TOGGLE,
    IDC_PREFS_MONITOR_FILTER_DEBUG_TOGGLE,
}};

constexpr size_t kMonitorDisplayToggleCount = 5u;
constexpr size_t kMonitorFilterToggleCount  = kMonitorToggleLabelStringIds.size() - kMonitorDisplayToggleCount;

[[nodiscard]] bool ApplyMonitorFilterPresetSelection(Common::Settings::MonitorSettings& monitor,
                                                     const Common::Settings::MonitorFilterPreset preset) noexcept
{
    bool changed          = monitor.filter.preset != preset;
    monitor.filter.preset = preset;

    const uint32_t previousMask = monitor.filter.mask;
    switch (preset)
    {
        case Common::Settings::MonitorFilterPreset::ErrorsOnly: monitor.filter.mask = static_cast<uint32_t>(MonitorFilterBit::Error); break;
        case Common::Settings::MonitorFilterPreset::ErrorsWarnings: monitor.filter.mask = MonitorFilterBit::Error | MonitorFilterBit::Warning; break;
        case Common::Settings::MonitorFilterPreset::AllTypes:
            monitor.filter.mask =
                MonitorFilterBit::Text | MonitorFilterBit::Error | MonitorFilterBit::Warning | MonitorFilterBit::Info | MonitorFilterBit::Perf |
                MonitorFilterBit::Debug;
            break;
        case Common::Settings::MonitorFilterPreset::Custom:
        default: break;
    }

    changed = changed || previousMask != monitor.filter.mask;
    return changed;
}

void ReorderPanelChildren(Panel* root, std::span<RedSalamander::DxUi::Control* const> orderedControls)
{
    if (! root)
    {
        return;
    }

    auto children = root->GetChildren();
    if (children.empty())
    {
        return;
    }

    std::vector<std::unique_ptr<RedSalamander::DxUi::Control>> reordered;
    reordered.reserve(children.size());

    auto moveChild = [&](RedSalamander::DxUi::Control* wanted) noexcept
    {
        if (! wanted)
        {
            return;
        }

        for (auto& child : children)
        {
            if (child && child.get() == wanted)
            {
                reordered.push_back(std::move(child));
                return;
            }
        }
    };

    for (auto* control : orderedControls)
    {
        moveChild(control);
    }

    for (auto& child : children)
    {
        if (child)
        {
            reordered.push_back(std::move(child));
        }
    }

    if (reordered.size() != children.size())
    {
        return;
    }

    for (size_t index = 0; index < children.size(); ++index)
    {
        children[index] = std::move(reordered[index]);
    }
}
} // namespace

struct MonitorToggleCardPageDx
{
    CardPanel* card    = nullptr;
    Label* label       = nullptr;
    Label* description = nullptr;
    Toggle* toggle     = nullptr;
};

struct MonitorFilterCardPageDx
{
    CardPanel* card           = nullptr;
    Label* presetLabel        = nullptr;
    Label* presetDescription  = nullptr;
    ComboBox* presetCombo     = nullptr;
    std::array<Label*, kMonitorFilterToggleCount> toggleLabels{};
    std::array<Label*, kMonitorFilterToggleCount> toggleDescriptions{};
    std::array<Toggle*, kMonitorFilterToggleCount> toggles{};
};

struct MonitorLinkCardPageDx
{
    CardPanel* card    = nullptr;
    Label* label       = nullptr;
    Label* description = nullptr;
    Button* link       = nullptr;
};

struct MonitorDxPage
{
    MonitorDxPage()                                = default;
    MonitorDxPage(const MonitorDxPage&)            = delete;
    MonitorDxPage& operator=(const MonitorDxPage&) = delete;
    MonitorDxPage(MonitorDxPage&&)                 = delete;
    MonitorDxPage& operator=(MonitorDxPage&&)      = delete;

    std::array<MonitorToggleCardPageDx, kMonitorDisplayToggleCount> displayToggleCards{};
    MonitorLinkCardPageDx settingsFileCard{};
    MonitorFilterCardPageDx filterCard{};

    void Detach() noexcept
    {
        for (auto& toggleCard : displayToggleCards)
        {
            toggleCard = {};
        }
        settingsFileCard = {};
        filterCard       = {};
    }
};

struct MonitorPane::DxState
{
    DxState()                          = default;
    DxState(const DxState&)            = delete;
    DxState& operator=(const DxState&) = delete;
    DxState(DxState&&)                 = delete;
    DxState& operator=(DxState&&)      = delete;

    MonitorDxPage page;

    void Detach() noexcept
    {
        page.Detach();
    }
};

MonitorPane::MonitorPane() = default;

MonitorPane::~MonitorPane()
{
    DetachDxHosts();
}

void MonitorPane::OnVisibilityChanged(bool visible) noexcept
{
    static_cast<void>(visible);
}

void MonitorPane::Destroy(PreferencesDialogState& state) noexcept
{
    static_cast<void>(state);
    DetachDxHosts();
    _pageHost = nullptr;
}

bool MonitorPane::EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept
{
    _pageHostDx      = state.pageHostDxHost;
    _pageContentRoot = state.pageHostDxContentRootControl;
    if (! _pageHostDx || ! _pageContentRoot)
    {
        return false;
    }

    if (_dxState && PrefsUi::HasRetainedDxChildren(_pageContentRoot))
    {
        ApplyDxTheme(state);
        SyncDxControlsFromState(state);
        return true;
    }

    auto dxState = std::make_unique<DxState>();
    _pageHostDx->ResetInteractionState();
    _pageContentRoot->ClearChildren();

    auto* root = _pageContentRoot;

    const auto attachToggleHandler = [this, parent](Toggle* toggle, const UINT commandId) noexcept
    {
        if (! toggle)
        {
            return;
        }

        toggle->SetStateLabels(LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
        toggle->SetOnToggled([this, host = parent, commandId](bool checked) noexcept
        {
            if (! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* dialogState = PrefsUi::GetDialogState(host);
            if (! dialogState)
            {
                return;
            }

            auto* monitor = EnsureWorkingMonitorSettings(dialogState->workingMonitorSettings);
            if (! monitor)
            {
                return;
            }

            bool changed = false;
            const auto updateFilterBit = [&](uint32_t bit) noexcept
            {
                monitor->filter.preset = Common::Settings::MonitorFilterPreset::Custom;
                uint32_t mask          = monitor->filter.mask & 63u;
                const uint32_t updated = checked ? (mask | bit) : (mask & ~bit);
                changed                = (mask != updated);
                monitor->filter.mask   = updated & 63u;
            };

            switch (commandId)
            {
                case IDC_PREFS_MONITOR_TOOLBAR_TOGGLE:
                    changed                      = monitor->menu.toolbarVisible != checked;
                    monitor->menu.toolbarVisible = checked;
                    break;
                case IDC_PREFS_MONITOR_LINE_NUMBERS_TOGGLE:
                    changed                          = monitor->menu.lineNumbersVisible != checked;
                    monitor->menu.lineNumbersVisible = checked;
                    break;
                case IDC_PREFS_MONITOR_ALWAYS_ON_TOP_TOGGLE:
                    changed                   = monitor->menu.alwaysOnTop != checked;
                    monitor->menu.alwaysOnTop = checked;
                    break;
                case IDC_PREFS_MONITOR_SHOW_IDS_TOGGLE:
                    changed               = monitor->menu.showIds != checked;
                    monitor->menu.showIds = checked;
                    break;
                case IDC_PREFS_MONITOR_AUTO_SCROLL_TOGGLE:
                    changed                  = monitor->menu.autoScroll != checked;
                    monitor->menu.autoScroll = checked;
                    break;
                case IDC_PREFS_MONITOR_FILTER_TEXT_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Text)); break;
                case IDC_PREFS_MONITOR_FILTER_ERROR_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Error)); break;
                case IDC_PREFS_MONITOR_FILTER_WARNING_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Warning)); break;
                case IDC_PREFS_MONITOR_FILTER_INFO_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Info)); break;
                case IDC_PREFS_MONITOR_FILTER_PERF_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Perf)); break;
                case IDC_PREFS_MONITOR_FILTER_DEBUG_TOGGLE: updateFilterBit(static_cast<uint32_t>(MonitorFilterBit::Debug)); break;
                default: break;
            }

            if (changed)
            {
                SetDirty(GetParent(host), *dialogState);
            }
            Refresh(host, *dialogState);
        });
    };

    for (size_t i = 0; i < dxState->page.displayToggleCards.size(); ++i)
    {
        auto& card = dxState->page.displayToggleCards[i];
        card.card  = root->AddChild<CardPanel>();
        card.label = root->AddChild<Label>();
        card.label->SetFontRole(FontRole::Body);
        card.description = root->AddChild<Label>();
        card.description->SetFontRole(FontRole::Small);
        card.description->SetMultiline(true);
        card.toggle = root->AddChild<Toggle>();
        attachToggleHandler(card.toggle, kMonitorToggleCommandIds[i]);
    }

    {
        auto& card             = dxState->page.filterCard;
        card.card              = root->AddChild<CardPanel>();
        card.presetLabel       = root->AddChild<Label>();
        card.presetDescription = root->AddChild<Label>();
        card.presetCombo       = root->AddChild<ComboBox>();

        card.presetLabel->SetFontRole(FontRole::Body);
        card.presetDescription->SetFontRole(FontRole::Small);
        card.presetDescription->SetMultiline(true);
        card.presetCombo->SetVariant(ComboBoxVariant::Window);
        card.presetCombo->SetOnSelectionChanged([this, host = parent](size_t itemIndex) noexcept
        {
            if (_syncingDxFilterPresetCombo || ! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* dialogState = PrefsUi::GetDialogState(host);
            if (! dialogState)
            {
                return;
            }

            auto* monitor = EnsureWorkingMonitorSettings(dialogState->workingMonitorSettings);
            if (! monitor)
            {
                return;
            }

            constexpr std::array<Common::Settings::MonitorFilterPreset, 4> kPresetOrder = {{
                Common::Settings::MonitorFilterPreset::Custom,
                Common::Settings::MonitorFilterPreset::ErrorsOnly,
                Common::Settings::MonitorFilterPreset::ErrorsWarnings,
                Common::Settings::MonitorFilterPreset::AllTypes,
            }};
            if (itemIndex >= kPresetOrder.size())
            {
                return;
            }

            const bool changed = ApplyMonitorFilterPresetSelection(*monitor, kPresetOrder[itemIndex]);
            if (changed)
            {
                SetDirty(GetParent(host), *dialogState);
            }
            Refresh(host, *dialogState);
        });

        for (size_t i = 0; i < card.toggles.size(); ++i)
        {
            card.toggleLabels[i] = root->AddChild<Label>();
            card.toggleLabels[i]->SetFontRole(FontRole::Body);
            card.toggleDescriptions[i] = root->AddChild<Label>();
            card.toggleDescriptions[i]->SetFontRole(FontRole::Small);
            card.toggleDescriptions[i]->SetMultiline(true);
            card.toggles[i] = root->AddChild<Toggle>();
            attachToggleHandler(card.toggles[i], kMonitorToggleCommandIds[kMonitorDisplayToggleCount + i]);
        }
    }

    {
        auto& card = dxState->page.settingsFileCard;
        card.card  = root->AddChild<CardPanel>();
        card.label = root->AddChild<Label>();
        card.label->SetFontRole(FontRole::Body);
        card.description = root->AddChild<Label>();
        card.description->SetFontRole(FontRole::Small);
        card.description->SetMultiline(true);
        card.link = root->AddChild<Button>();
        card.link->SetVariant(ButtonVariant::Hyperlink);
        card.link->SetOnClick([host = parent]() noexcept
        {
            if (! host || IsWindow(host) == FALSE)
            {
                return;
            }

            if (HWND dlg = GetParent(host))
            {
                if (PostMessageW(dlg, WM_COMMAND, MAKEWPARAM(IDC_PREFS_MONITOR_OPEN_SETTINGS_FILE, 0), 0) == FALSE)
                {
                    static_cast<void>(Debug::ErrorWithLastError(L"Preferences.Monitor: PostMessageW failed for settings-file link"));
                }
            }
        });
        card.label->SetMnemonicTarget(card.link);
    }

    {
        MonitorDxPage& page = dxState->page;
        std::vector<RedSalamander::DxUi::Control*> orderedChildren;
        orderedChildren.reserve(_pageContentRoot->DebugChildCount());

        const auto appendDisplayToggleCard = [&](const size_t index) noexcept
        {
            orderedChildren.push_back(page.displayToggleCards[index].card);
            orderedChildren.push_back(page.displayToggleCards[index].label);
            orderedChildren.push_back(page.displayToggleCards[index].description);
            orderedChildren.push_back(page.displayToggleCards[index].toggle);
        };

        const auto appendSettingsFileCard = [&]() noexcept
        {
            orderedChildren.push_back(page.settingsFileCard.card);
            orderedChildren.push_back(page.settingsFileCard.label);
            orderedChildren.push_back(page.settingsFileCard.description);
            orderedChildren.push_back(page.settingsFileCard.link);
        };

        const auto appendFilterCard = [&]() noexcept
        {
            orderedChildren.push_back(page.filterCard.card);
            orderedChildren.push_back(page.filterCard.presetLabel);
            orderedChildren.push_back(page.filterCard.presetDescription);
            orderedChildren.push_back(page.filterCard.presetCombo);
            for (size_t i = 0; i < page.filterCard.toggles.size(); ++i)
            {
                orderedChildren.push_back(page.filterCard.toggleLabels[i]);
                orderedChildren.push_back(page.filterCard.toggleDescriptions[i]);
                orderedChildren.push_back(page.filterCard.toggles[i]);
            }
        };

        for (size_t i = 0; i < page.displayToggleCards.size(); ++i)
        {
            appendDisplayToggleCard(i);
        }
        appendFilterCard();
        appendSettingsFileCard();

        ReorderPanelChildren(_pageContentRoot, orderedChildren);
    }

    _dxState = std::move(dxState);
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    return true;
}

void MonitorPane::DetachDxHosts() noexcept
{
    if (_pageContentRoot && _pageHostDx && _pageHost && IsWindow(_pageHost) != FALSE)
    {
        _pageHostDx->ResetInteractionState();
        _pageContentRoot->ClearChildren();
    }
    _pageHostDx      = nullptr;
    _pageContentRoot = nullptr;

    if (_dxState)
    {
        _dxState->Detach();
        _dxState.reset();
    }
}

void MonitorPane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return;
    }

    const ThemePalette palette = PrefsUi::MakeDxPalette(state.theme);
    _pageHostDx->SetTheme(palette);
}

void MonitorPane::SyncDxControlsFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState)
    {
        return;
    }

    const ThemePalette palette = PrefsUi::MakeDxPalette(state.theme);
    const auto& monitor        = GetMonitorSettingsOrDefault(state.workingMonitorSettings);
    const uint32_t mask        = monitor.filter.mask & 63u;
    const bool customFilter    = (monitor.filter.preset == Common::Settings::MonitorFilterPreset::Custom);

    const std::array<bool, kMonitorDisplayToggleCount> displayToggleChecked = {{
        monitor.menu.toolbarVisible,
        monitor.menu.lineNumbersVisible,
        monitor.menu.alwaysOnTop,
        monitor.menu.showIds,
        monitor.menu.autoScroll,
    }};

    const std::array<bool, kMonitorFilterToggleCount> filterToggleChecked = {{
        HasFlag(mask, MonitorFilterBit::Text),
        HasFlag(mask, MonitorFilterBit::Error),
        HasFlag(mask, MonitorFilterBit::Warning),
        HasFlag(mask, MonitorFilterBit::Info),
        HasFlag(mask, MonitorFilterBit::Perf),
        HasFlag(mask, MonitorFilterBit::Debug),
    }};

    for (size_t i = 0; i < _dxState->page.displayToggleCards.size(); ++i)
    {
        auto& toggleCard = _dxState->page.displayToggleCards[i];
        if (toggleCard.label)
        {
            toggleCard.label->SetText(LoadStringResource(nullptr, kMonitorToggleLabelStringIds[i]));
            toggleCard.label->SetTextColor(std::nullopt);
        }
        if (toggleCard.description)
        {
            toggleCard.description->SetText(LoadStringResource(nullptr, kMonitorToggleDescriptionStringIds[i]));
            toggleCard.description->SetTextColor(palette.subduedText);
        }
        if (toggleCard.toggle)
        {
            toggleCard.toggle->SetEnabled(true);
            toggleCard.toggle->SetChecked(displayToggleChecked[i]);
        }
    }

    auto& filterCard = _dxState->page.filterCard;
    if (filterCard.presetLabel)
    {
        filterCard.presetLabel->SetText(LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_PRESET));
        filterCard.presetLabel->SetTextColor(std::nullopt);
    }
    if (filterCard.presetDescription)
    {
        filterCard.presetDescription->SetText(LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILTER_PRESET));
        filterCard.presetDescription->SetTextColor(palette.subduedText);
    }

    if (filterCard.presetCombo)
    {
        constexpr std::array<UINT, 4> kPresetStringIds = {{
            IDS_PREFS_ADV_FILTER_CUSTOM,
            IDS_PREFS_ADV_FILTER_ERRORS_ONLY,
            IDS_PREFS_ADV_FILTER_ERRORS_WARNINGS,
            IDS_PREFS_ADV_FILTER_ALL_TYPES,
        }};

        std::vector<ComboBox::Item> items;
        items.reserve(kPresetStringIds.size());
        for (const auto stringId : kPresetStringIds)
        {
            const std::wstring label = LoadStringResource(nullptr, stringId);
            items.push_back(ComboBox::Item{label, label});
        }

        constexpr std::array<Common::Settings::MonitorFilterPreset, 4> kPresetOrder = {{
            Common::Settings::MonitorFilterPreset::Custom,
            Common::Settings::MonitorFilterPreset::ErrorsOnly,
            Common::Settings::MonitorFilterPreset::ErrorsWarnings,
            Common::Settings::MonitorFilterPreset::AllTypes,
        }};

        std::optional<size_t> selectedIndex;
        for (size_t i = 0; i < kPresetOrder.size(); ++i)
        {
            if (kPresetOrder[i] == monitor.filter.preset)
            {
                selectedIndex = i;
                break;
            }
        }

        _syncingDxFilterPresetCombo = true;
        filterCard.presetCombo->SetItems(std::move(items));
        filterCard.presetCombo->SetSelectedIndex(selectedIndex);
        filterCard.presetCombo->SetEnabled(true);
        _syncingDxFilterPresetCombo = false;
    }

    for (size_t i = 0; i < filterCard.toggles.size(); ++i)
    {
        const size_t stringIndex = kMonitorDisplayToggleCount + i;
        if (filterCard.toggleLabels[i])
        {
            filterCard.toggleLabels[i]->SetText(LoadStringResource(nullptr, kMonitorToggleLabelStringIds[stringIndex]));
            filterCard.toggleLabels[i]->SetTextColor(customFilter ? std::optional<D2D1_COLOR_F>{} : std::optional<D2D1_COLOR_F>(palette.disabledText));
        }
        if (filterCard.toggleDescriptions[i])
        {
            filterCard.toggleDescriptions[i]->SetText(LoadStringResource(nullptr, kMonitorToggleDescriptionStringIds[stringIndex]));
            filterCard.toggleDescriptions[i]->SetTextColor(customFilter ? std::optional<D2D1_COLOR_F>(palette.subduedText)
                                                                        : std::optional<D2D1_COLOR_F>(palette.disabledText));
        }
        if (filterCard.toggles[i])
        {
            filterCard.toggles[i]->SetEnabled(customFilter);
            filterCard.toggles[i]->SetChecked(filterToggleChecked[i]);
        }
    }

    if (_dxState->page.settingsFileCard.label)
    {
        _dxState->page.settingsFileCard.label->SetText(LoadStringResource(nullptr, IDS_PREFS_MONITOR_LABEL_SETTINGS_FILE));
        _dxState->page.settingsFileCard.label->SetTextColor(std::nullopt);
    }
    if (_dxState->page.settingsFileCard.description)
    {
        _dxState->page.settingsFileCard.description->SetText(LoadStringResource(nullptr, IDS_PREFS_MONITOR_DESC_SETTINGS_FILE));
        _dxState->page.settingsFileCard.description->SetTextColor(palette.subduedText);
    }
    if (_dxState->page.settingsFileCard.link)
    {
        _dxState->page.settingsFileCard.link->SetText(LoadStringResource(nullptr, IDS_PREFS_MONITOR_OPEN_SETTINGS_FILE_LINK));
        _dxState->page.settingsFileCard.link->SetEnabled(true);
    }
    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void MonitorPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost = parent;

    if (state.currentCategory != PrefCategory::Monitor)
    {
        return;
    }

    if (! EnsureDxHosts(parent, state))
    {
        Debug::Error(L"Preferences.Monitor: DxUi surface initialization failed; page will not render correctly.");
        DetachDxHosts();
        return;
    }

    Refresh(parent, state);
}

void MonitorPane::LayoutDxPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(host);
    static_cast<void>(margin);

    Debug::Perf::Scope layoutPerf(L"preferences.ui.monitor_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(static_cast<uint64_t>(std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI)));

    const UINT dpi = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);

    const int rowHeight   = std::max(1, UiMetrics::ScaleDip(dpi, kRowHeightDip));
    const int titleHeight = std::max(1, UiMetrics::ScaleDip(dpi, kTitleHeightDip));

    const int cardPaddingX = UiMetrics::ScaleDip(dpi, kCardPaddingXDip);
    const int cardPaddingY = UiMetrics::ScaleDip(dpi, kCardPaddingYDip);
    const int cardGapY     = UiMetrics::ScaleDip(dpi, kCardGapYDip);
    const int cardGapX     = UiMetrics::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY = UiMetrics::ScaleDip(dpi, kCardSpacingYDip);

    if (! _dxState || ! _pageHostDx || ! _pageContentRoot)
    {
        return;
    }
    MonitorDxPage& dxPage = _dxState->page;

    const auto pxToDip = [dpi](const int px) noexcept { return static_cast<float>(px) * 96.0f / static_cast<float>(std::max<UINT>(1u, dpi)); };

    const int minToggleWidth    = UiMetrics::ScaleDip(dpi, kMinToggleWidthDip);
    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

    const int onWidth  = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, onLabel);
    const int offWidth = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, offLabel);

    const int paddingX       = UiMetrics::ScaleDip(dpi, kTogglePaddingXDip);
    const int gapX           = UiMetrics::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth     = UiMetrics::ScaleDip(dpi, kToggleTrackWidthDip);
    const int stateTextWidth = std::max(onWidth, offWidth);

    const int measuredToggleWidth = std::max(minToggleWidth, (2 * paddingX) + stateTextWidth + gapX + trackWidth);
    const int toggleWidth =
        static_cast<int>(std::lround(RedSalamander::DxUi::ResolveConstrainedExtent({.minExtent       = static_cast<float>(minToggleWidth),
                                                                                    .preferredExtent = static_cast<float>(measuredToggleWidth),
                                                                                    .maxExtent       = static_cast<float>(measuredToggleWidth)},
                                                                                   static_cast<float>(std::max(0, width - 2 * cardPaddingX - cardGapX)))));

    auto layoutToggleCard =
        [&](std::wstring_view labelText, std::wstring_view descText, CardPanel* dxCard, Label* dxLabel, Label* dxDescription, Toggle* dxToggle) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        state.pageSettingCards.push_back(card);

        if (dxCard)
        {
            dxCard->SetBounds(D2D1::RectF(pxToDip(card.left), pxToDip(card.top), pxToDip(card.right), pxToDip(card.bottom)));
        }
        if (dxLabel)
        {
            dxLabel->SetText(std::wstring(labelText));
            dxLabel->SetMnemonicTarget(dxToggle);
            dxLabel->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                           pxToDip(card.top + cardPaddingY),
                                           pxToDip(card.left + cardPaddingX + textWidth),
                                           pxToDip(card.top + cardPaddingY + titleHeight)));
        }
        if (dxDescription)
        {
            dxDescription->SetText(std::wstring(descText));
            dxDescription->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                 pxToDip(card.top + cardPaddingY + titleHeight + cardGapY),
                                                 pxToDip(card.left + cardPaddingX + textWidth),
                                                 pxToDip(card.top + cardPaddingY + titleHeight + cardGapY + descHeight)));
        }
        if (dxToggle)
        {
            dxToggle->SetBounds(D2D1::RectF(pxToDip(card.right - cardPaddingX - toggleWidth),
                                            pxToDip(card.top + (cardHeight - rowHeight) / 2),
                                            pxToDip(card.right - cardPaddingX),
                                            pxToDip(card.top + (cardHeight - rowHeight) / 2 + rowHeight)));
        }

        y += cardHeight + cardSpacingY;
    };

    auto layoutFilterCard = [&](MonitorFilterCardPageDx& dxCard) noexcept
    {
        const int comboWidth = static_cast<int>(
            std::lround(RedSalamander::DxUi::ResolveConstrainedExtent({.minExtent       = static_cast<float>(UiMetrics::ScaleDip(dpi, kMinEditWidthDip)),
                                                                       .preferredExtent = static_cast<float>(UiMetrics::ScaleDip(dpi, kMinEditWidthDip)),
                                                                       .maxExtent       = static_cast<float>(UiMetrics::ScaleDip(dpi, kMaxEditWidthDip))},
                                                                      static_cast<float>(std::max(0, width - 2 * cardPaddingX)))));
        const int rowGapY = cardSpacingY;

        const auto measureRowHeight = [&](const int controlWidth, std::wstring_view descText) noexcept
        {
            const int textWidth     = std::max(0, width - 2 * cardPaddingX - cardGapX - controlWidth);
            const int descHeight    = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);
            const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
            return std::max(rowHeight, contentHeight);
        };

        const std::wstring presetDesc = LoadStringResource(nullptr, IDS_PREFS_ADV_DESC_FILTER_PRESET);
        const int presetRowHeight     = measureRowHeight(comboWidth, presetDesc);

        std::array<int, kMonitorFilterToggleCount> filterRowHeights{};
        int contentHeight = presetRowHeight;
        for (size_t i = 0; i < filterRowHeights.size(); ++i)
        {
            const size_t stringIndex = kMonitorDisplayToggleCount + i;
            const std::wstring desc  = LoadStringResource(nullptr, kMonitorToggleDescriptionStringIds[stringIndex]);
            filterRowHeights[i]      = measureRowHeight(toggleWidth, desc);
            contentHeight += rowGapY + filterRowHeights[i];
        }

        const int cardHeight = contentHeight + (2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        state.pageSettingCards.push_back(card);

        if (dxCard.card)
        {
            dxCard.card->SetBounds(D2D1::RectF(pxToDip(card.left), pxToDip(card.top), pxToDip(card.right), pxToDip(card.bottom)));
        }

        int rowTop = card.top + cardPaddingY;
        {
            const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - comboWidth);
            const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, presetDesc);
            if (dxCard.presetLabel)
            {
                dxCard.presetLabel->SetText(LoadStringResource(nullptr, IDS_PREFS_ADV_LABEL_FILTER_PRESET));
                dxCard.presetLabel->SetMnemonicTarget(dxCard.presetCombo);
                dxCard.presetLabel->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                          pxToDip(rowTop),
                                                          pxToDip(card.left + cardPaddingX + textWidth),
                                                          pxToDip(rowTop + titleHeight)));
            }
            if (dxCard.presetDescription)
            {
                dxCard.presetDescription->SetText(presetDesc);
                dxCard.presetDescription->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                                pxToDip(rowTop + titleHeight + cardGapY),
                                                                pxToDip(card.left + cardPaddingX + textWidth),
                                                                pxToDip(rowTop + titleHeight + cardGapY + descHeight)));
            }
            if (dxCard.presetCombo)
            {
                const int comboTop = rowTop + ((presetRowHeight - rowHeight) / 2);
                dxCard.presetCombo->SetBounds(D2D1::RectF(pxToDip(card.right - cardPaddingX - comboWidth),
                                                          pxToDip(comboTop),
                                                          pxToDip(card.right - cardPaddingX),
                                                          pxToDip(comboTop + rowHeight)));
            }
        }

        rowTop += presetRowHeight + rowGapY;
        for (size_t i = 0; i < dxCard.toggles.size(); ++i)
        {
            const size_t stringIndex  = kMonitorDisplayToggleCount + i;
            const std::wstring label  = LoadStringResource(nullptr, kMonitorToggleLabelStringIds[stringIndex]);
            const std::wstring desc   = LoadStringResource(nullptr, kMonitorToggleDescriptionStringIds[stringIndex]);
            const int rowBlockHeight  = filterRowHeights[i];
            const int textWidth       = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
            const int descHeight      = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, desc);

            if (dxCard.toggleLabels[i])
            {
                dxCard.toggleLabels[i]->SetText(label);
                dxCard.toggleLabels[i]->SetMnemonicTarget(dxCard.toggles[i]);
                dxCard.toggleLabels[i]->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                              pxToDip(rowTop),
                                                              pxToDip(card.left + cardPaddingX + textWidth),
                                                              pxToDip(rowTop + titleHeight)));
            }
            if (dxCard.toggleDescriptions[i])
            {
                dxCard.toggleDescriptions[i]->SetText(desc);
                dxCard.toggleDescriptions[i]->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                                    pxToDip(rowTop + titleHeight + cardGapY),
                                                                    pxToDip(card.left + cardPaddingX + textWidth),
                                                                    pxToDip(rowTop + titleHeight + cardGapY + descHeight)));
            }
            if (dxCard.toggles[i])
            {
                const int toggleTop = rowTop + ((rowBlockHeight - rowHeight) / 2);
                dxCard.toggles[i]->SetBounds(D2D1::RectF(pxToDip(card.right - cardPaddingX - toggleWidth),
                                                         pxToDip(toggleTop),
                                                         pxToDip(card.right - cardPaddingX),
                                                         pxToDip(toggleTop + rowHeight)));
            }

            rowTop += rowBlockHeight + rowGapY;
        }

        y += cardHeight + cardSpacingY;
    };

    auto layoutLinkCard = [&](std::wstring_view labelText,
                              std::wstring_view descText,
                              std::wstring_view linkText,
                              CardPanel* dxCard,
                              Label* dxLabel,
                              Label* dxDescription,
                              Button* dxButton) noexcept
    {
        const int textWidth       = std::max(0, width - 2 * cardPaddingX);
        const int descHeight      = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);
        const int measuredWidth   = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, linkText) + UiMetrics::ScaleDip(dpi, 16);
        const int linkWidth       = std::max(1, std::min(textWidth, measuredWidth));
        const int contentHeight   = titleHeight + cardGapY + descHeight + cardGapY + rowHeight;
        const int cardHeight      = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);
        const int linkTop         = y + cardPaddingY + titleHeight + cardGapY + descHeight + cardGapY;

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        state.pageSettingCards.push_back(card);

        if (dxCard)
        {
            dxCard->SetBounds(D2D1::RectF(pxToDip(card.left), pxToDip(card.top), pxToDip(card.right), pxToDip(card.bottom)));
        }
        if (dxLabel)
        {
            dxLabel->SetText(std::wstring(labelText));
            dxLabel->SetMnemonicTarget(dxButton);
            dxLabel->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                           pxToDip(card.top + cardPaddingY),
                                           pxToDip(card.left + cardPaddingX + textWidth),
                                           pxToDip(card.top + cardPaddingY + titleHeight)));
        }
        if (dxDescription)
        {
            dxDescription->SetText(std::wstring(descText));
            dxDescription->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                 pxToDip(card.top + cardPaddingY + titleHeight + cardGapY),
                                                 pxToDip(card.left + cardPaddingX + textWidth),
                                                 pxToDip(card.top + cardPaddingY + titleHeight + cardGapY + descHeight)));
        }
        if (dxButton)
        {
            dxButton->SetText(std::wstring(linkText));
            dxButton->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                            pxToDip(linkTop),
                                            pxToDip(card.left + cardPaddingX + linkWidth),
                                            pxToDip(linkTop + rowHeight)));
        }

        y += cardHeight + cardSpacingY;
    };

    for (size_t i = 0; i < dxPage.displayToggleCards.size(); ++i)
    {
        layoutToggleCard(LoadStringResource(nullptr, kMonitorToggleLabelStringIds[i]),
                         LoadStringResource(nullptr, kMonitorToggleDescriptionStringIds[i]),
                         dxPage.displayToggleCards[i].card,
                         dxPage.displayToggleCards[i].label,
                         dxPage.displayToggleCards[i].description,
                         dxPage.displayToggleCards[i].toggle);
    }

    y += gapY;

    layoutFilterCard(dxPage.filterCard);

    y += gapY;

    layoutLinkCard(LoadStringResource(nullptr, IDS_PREFS_MONITOR_LABEL_SETTINGS_FILE),
                   LoadStringResource(nullptr, IDS_PREFS_MONITOR_DESC_SETTINGS_FILE),
                   LoadStringResource(nullptr, IDS_PREFS_MONITOR_OPEN_SETTINGS_FILE_LINK),
                   dxPage.settingsFileCard.card,
                   dxPage.settingsFileCard.label,
                   dxPage.settingsFileCard.description,
                   dxPage.settingsFileCard.link);

    SyncDxControlsFromState(state);
    _pageHostDx->Invalidate();
}

void MonitorPane::LayoutPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    if (! host)
    {
        return;
    }

    if (EnsureDxHosts(_pageHost ? _pageHost : host, state))
    {
        LayoutDxPage(host, state, x, y, width, margin, gapY, typography);
        return;
    }

    Debug::Error(L"Preferences.Monitor: DxUi surface initialization failed; page will not render correctly.");
}

void MonitorPane::Refresh(HWND /*host*/, PreferencesDialogState& state) noexcept
{
    if (state.currentCategory == PrefCategory::Monitor && ! _dxState)
    {
        static_cast<void>(EnsureDxHosts(_pageHost, state));
    }

    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
}

#ifdef ENABLE_TESTS
PreferencesMonitorDebugFocusTarget MonitorPane::DebugGetFocusTarget() const noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return PreferencesMonitorDebugFocusTarget::None;
    }

    const auto* focused = _pageHostDx->GetFocusControl();
    if (! focused)
    {
        return PreferencesMonitorDebugFocusTarget::None;
    }

    const auto& page = _dxState->page;
    if (page.displayToggleCards[0].toggle == focused)
        return PreferencesMonitorDebugFocusTarget::ToolbarToggle;
    if (page.displayToggleCards[1].toggle == focused)
        return PreferencesMonitorDebugFocusTarget::LineNumbersToggle;
    if (page.displayToggleCards[2].toggle == focused)
        return PreferencesMonitorDebugFocusTarget::AlwaysOnTopToggle;
    if (page.displayToggleCards[3].toggle == focused)
        return PreferencesMonitorDebugFocusTarget::ShowIdsToggle;
    if (page.displayToggleCards[4].toggle == focused)
        return PreferencesMonitorDebugFocusTarget::AutoScrollToggle;
    if (page.filterCard.presetCombo == focused)
        return PreferencesMonitorDebugFocusTarget::FilterPresetCombo;
    if (page.filterCard.toggles[0] == focused)
        return PreferencesMonitorDebugFocusTarget::FilterTextToggle;
    if (page.filterCard.toggles[1] == focused)
        return PreferencesMonitorDebugFocusTarget::FilterErrorToggle;
    if (page.filterCard.toggles[2] == focused)
        return PreferencesMonitorDebugFocusTarget::FilterWarningToggle;
    if (page.filterCard.toggles[3] == focused)
        return PreferencesMonitorDebugFocusTarget::FilterInfoToggle;
    if (page.filterCard.toggles[4] == focused)
        return PreferencesMonitorDebugFocusTarget::FilterPerfToggle;
    if (page.filterCard.toggles[5] == focused)
        return PreferencesMonitorDebugFocusTarget::FilterDebugToggle;
    if (page.settingsFileCard.link == focused)
        return PreferencesMonitorDebugFocusTarget::OpenSettingsFileLink;

    return PreferencesMonitorDebugFocusTarget::None;
}

bool MonitorPane::DebugFocusToolbarToggle() noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return false;
    }

    auto* const toggle = _dxState->page.displayToggleCards[0].toggle;
    if (! toggle || ! toggle->IsVisible() || ! toggle->IsEnabled())
    {
        return false;
    }

    _pageHostDx->SetFocusControl(toggle);
    return _pageHostDx->GetFocusControl() == toggle;
}

bool MonitorPane::DebugSelectFilterPresetByText(std::wstring_view displayText) noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return false;
    }

    auto* const combo = _dxState->page.filterCard.presetCombo;
    if (! combo || ! combo->IsVisible() || ! combo->IsEnabled())
    {
        return false;
    }

    const auto items = combo->GetItems();
    const auto it    = std::find_if(items.begin(), items.end(), [displayText](const ComboBox::Item& item) noexcept { return item.display == displayText; });
    if (it == items.end())
    {
        return false;
    }

    const size_t itemIndex = static_cast<size_t>(std::distance(items.begin(), it));
    _pageHostDx->SetFocusControl(combo);
    combo->SetSelectedIndex(itemIndex);

    if (_pageHost && IsWindow(_pageHost) != FALSE)
    {
        auto* dialogState = PrefsUi::GetDialogState(_pageHost);
        if (dialogState)
        {
            constexpr std::array<Common::Settings::MonitorFilterPreset, 4> kPresetOrder = {{
                Common::Settings::MonitorFilterPreset::Custom,
                Common::Settings::MonitorFilterPreset::ErrorsOnly,
                Common::Settings::MonitorFilterPreset::ErrorsWarnings,
                Common::Settings::MonitorFilterPreset::AllTypes,
            }};
            if (itemIndex < kPresetOrder.size())
            {
                if (auto* monitor = EnsureWorkingMonitorSettings(dialogState->workingMonitorSettings))
                {
                    const bool changed = ApplyMonitorFilterPresetSelection(*monitor, kPresetOrder[itemIndex]);
                    if (changed)
                    {
                        SetDirty(GetParent(_pageHost), *dialogState);
                    }
                    Refresh(_pageHost, *dialogState);
                }
            }
        }
    }

    _pageHostDx->Invalidate();
    return combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == itemIndex && combo->GetDisplayedText() == displayText;
}

bool MonitorPane::DebugSettingsFileCardIsLast() const noexcept
{
    if (! _pageContentRoot || ! _dxState || ! _dxState->page.settingsFileCard.link)
    {
        return false;
    }

    const auto children = _pageContentRoot->GetChildren();
    size_t filterCardIndex = children.size();
    size_t linkIndex       = children.size();
    for (size_t i = 0; i < children.size(); ++i)
    {
        const auto* const child = children[i].get();
        if (child == _dxState->page.filterCard.card)
        {
            filterCardIndex = i;
        }
        if (child == _dxState->page.settingsFileCard.link)
        {
            linkIndex = i;
        }
    }

    return filterCardIndex < children.size() && linkIndex + 1u == children.size() && filterCardIndex < linkIndex;
}
#endif
