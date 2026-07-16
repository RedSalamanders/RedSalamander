// Preferences.Panes.cpp

#include "Framework.h"

#include "Preferences.Panes.h"

#include <algorithm>
#include <array>
#include <cwctype>
#include <optional>
#include <string>

#include "DxUi/DxUi.h"
#include "Helpers.h"
#include "UiMetrics.h"
#include "WindowMessages.h"

#include "resource.h"

namespace
{
using RedSalamander::DxUi::CardPanel;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;
// Re-landed on the stabilized one-host page pattern.

enum class PanesComboSlot : size_t
{
    LeftDisplay = 0,
    LeftSortBy,
    LeftSortDir,
    RightDisplay,
    RightSortBy,
    RightSortDir,
};

struct PanesDxCard
{
    CardPanel* card    = nullptr;
    Label* title       = nullptr;
    Label* description = nullptr;
    Toggle* toggle     = nullptr;
    ComboBox* combo    = nullptr;
    TextField* edit    = nullptr;
};

} // namespace

struct PanesPane::DxState
{
    DxState()                          = default;
    DxState(const DxState&)            = delete;
    DxState& operator=(const DxState&) = delete;
    DxState(DxState&&)                 = delete;
    DxState& operator=(DxState&&)      = delete;

    Label* leftHeader    = nullptr;
    Label* rightHeader   = nullptr;
    Label* generalHeader = nullptr;

    PanesDxCard leftDisplay;
    PanesDxCard leftSortBy;
    PanesDxCard leftSortDir;
    PanesDxCard leftStatusBar;
    PanesDxCard rightDisplay;
    PanesDxCard rightSortBy;
    PanesDxCard rightSortDir;
    PanesDxCard rightStatusBar;
    PanesDxCard showHiddenFiles;
    PanesDxCard showSystemFiles;
    PanesDxCard history;

    void Detach() noexcept
    {
        leftHeader      = nullptr;
        rightHeader     = nullptr;
        generalHeader   = nullptr;
        leftDisplay     = {};
        leftSortBy      = {};
        leftSortDir     = {};
        leftStatusBar   = {};
        rightDisplay    = {};
        rightSortBy     = {};
        rightSortDir    = {};
        rightStatusBar  = {};
        showHiddenFiles = {};
        showSystemFiles = {};
        history         = {};
    }
};

PanesPane::PanesPane()  = default;
PanesPane::~PanesPane() = default;

void PanesPane::OnVisibilityChanged(bool visible) noexcept
{
    static_cast<void>(visible);
}

void PanesPane::Destroy(PreferencesDialogState& state) noexcept
{
    DetachDxHosts();
    UNREFERENCED_PARAMETER(state);
    _pageHost = nullptr;
}

bool PanesPane::EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept
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

    const auto addHeader = [&](Label*& outLabel) noexcept
    {
        outLabel = root->AddChild<Label>();
        outLabel->SetFontRole(FontRole::Header);
    };

    const auto addCard = [&](PanesDxCard& card, const bool withDescription) noexcept
    {
        card.card  = root->AddChild<CardPanel>();
        card.title = root->AddChild<Label>();
        card.title->SetFontRole(FontRole::Body);
        if (withDescription)
        {
            card.description = root->AddChild<Label>();
            card.description->SetFontRole(FontRole::Small);
            card.description->SetMultiline(true);
        }
    };

    const auto addToggle = [&](Toggle*& outToggle, const UINT commandId) noexcept
    {
        outToggle = root->AddChild<Toggle>();
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

            bool changed = false;

            switch (commandId)
            {
                case IDC_PREFS_PANES_LEFT_DISPLAY_TOGGLE:
                case IDC_PREFS_PANES_RIGHT_DISPLAY_TOGGLE:
                {
                    const std::wstring_view slot =
                        commandId == IDC_PREFS_PANES_LEFT_DISPLAY_TOGGLE ? PrefsFolders::kLeftPaneSlot : PrefsFolders::kRightPaneSlot;
                    if (auto* pane = PrefsFolders::EnsureWorkingFolderPane(state->workingSettings, slot))
                    {
                        const auto value   = checked ? Common::Settings::FolderDisplayMode::Brief : Common::Settings::FolderDisplayMode::Detailed;
                        changed            = pane->view.display != value;
                        pane->view.display = value;
                    }
                    break;
                }
                case IDC_PREFS_PANES_LEFT_SORTDIR_TOGGLE:
                case IDC_PREFS_PANES_RIGHT_SORTDIR_TOGGLE:
                {
                    const std::wstring_view slot =
                        commandId == IDC_PREFS_PANES_LEFT_SORTDIR_TOGGLE ? PrefsFolders::kLeftPaneSlot : PrefsFolders::kRightPaneSlot;
                    if (auto* pane = PrefsFolders::EnsureWorkingFolderPane(state->workingSettings, slot))
                    {
                        const auto value = checked ? Common::Settings::FolderSortDirection::Ascending : Common::Settings::FolderSortDirection::Descending;
                        changed          = pane->view.sortDirection != value;
                        pane->view.sortDirection = value;
                    }
                    break;
                }
                case IDC_PREFS_PANES_LEFT_STATUSBAR_TOGGLE:
                case IDC_PREFS_PANES_RIGHT_STATUSBAR_TOGGLE:
                {
                    const std::wstring_view slot =
                        commandId == IDC_PREFS_PANES_LEFT_STATUSBAR_TOGGLE ? PrefsFolders::kLeftPaneSlot : PrefsFolders::kRightPaneSlot;
                    if (auto* pane = PrefsFolders::EnsureWorkingFolderPane(state->workingSettings, slot))
                    {
                        changed                     = pane->view.statusBarVisible != checked;
                        pane->view.statusBarVisible = checked;
                    }
                    break;
                }
                case IDC_PREFS_PANES_SHOW_HIDDEN_TOGGLE:
                    if (auto* folders = PrefsFolders::EnsureWorkingFoldersSettings(state->workingSettings))
                    {
                        changed                  = folders->showHiddenFiles != checked;
                        folders->showHiddenFiles = checked;
                    }
                    break;
                case IDC_PREFS_PANES_SHOW_SYSTEM_TOGGLE:
                    if (auto* folders = PrefsFolders::EnsureWorkingFoldersSettings(state->workingSettings))
                    {
                        changed                  = folders->showSystemFiles != checked;
                        folders->showSystemFiles = checked;
                    }
                    break;
            }

            if (! changed)
            {
                return;
            }

            if (HWND dlg = GetParent(host))
            {
                SetDirty(dlg, *state);
            }
            Refresh(host, *state);
        });
    };

    const auto populateDisplayCombo = [](ComboBox* combo) noexcept
    {
        if (! combo)
        {
            return;
        }
        std::vector<ComboBox::Item> items;
        items.push_back(
            {std::to_wstring(static_cast<int>(Common::Settings::FolderDisplayMode::Brief)), LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_BRIEF)});
        items.push_back(
            {std::to_wstring(static_cast<int>(Common::Settings::FolderDisplayMode::Detailed)), LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DETAILED)});
        items.push_back({std::to_wstring(static_cast<int>(Common::Settings::FolderDisplayMode::ExtraDetailed)),
                         LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_EXTRA_DETAILED)});
        items.push_back({std::to_wstring(static_cast<int>(Common::Settings::FolderDisplayMode::Thumbnails)),
                         LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_THUMBNAILS)});
        combo->SetItems(std::move(items));
    };

    const auto populateSortByCombo = [](ComboBox* combo) noexcept
    {
        if (! combo)
        {
            return;
        }
        std::vector<ComboBox::Item> items;
        items.push_back({std::to_wstring(static_cast<int>(Common::Settings::FolderSortBy::Name)), LoadStringResource(nullptr, IDS_PREFS_PANES_SORT_NAME)});
        items.push_back(
            {std::to_wstring(static_cast<int>(Common::Settings::FolderSortBy::Extension)), LoadStringResource(nullptr, IDS_PREFS_PANES_SORT_EXTENSION)});
        items.push_back({std::to_wstring(static_cast<int>(Common::Settings::FolderSortBy::Time)), LoadStringResource(nullptr, IDS_PREFS_PANES_SORT_TIME)});
        items.push_back({std::to_wstring(static_cast<int>(Common::Settings::FolderSortBy::Size)), LoadStringResource(nullptr, IDS_PREFS_PANES_SORT_SIZE)});
        items.push_back(
            {std::to_wstring(static_cast<int>(Common::Settings::FolderSortBy::Attributes)), LoadStringResource(nullptr, IDS_PREFS_PANES_SORT_ATTRIBUTES)});
        items.push_back({std::to_wstring(static_cast<int>(Common::Settings::FolderSortBy::None)), LoadStringResource(nullptr, IDS_PREFS_PANES_SORT_NONE)});
        combo->SetItems(std::move(items));
    };

    const auto populateSortDirCombo = [](ComboBox* combo) noexcept
    {
        if (! combo)
        {
            return;
        }
        std::vector<ComboBox::Item> items;
        items.push_back({std::to_wstring(static_cast<int>(Common::Settings::FolderSortDirection::Ascending)),
                         LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_ASCENDING)});
        items.push_back({std::to_wstring(static_cast<int>(Common::Settings::FolderSortDirection::Descending)),
                         LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DESCENDING)});
        combo->SetItems(std::move(items));
    };

    const auto addDisplayCombo = [&](ComboBox*& outCombo, bool& syncFlag, const std::wstring_view slot) noexcept
    {
        outCombo = root->AddChild<ComboBox>();
        outCombo->SetVariant(ComboBoxVariant::Window);
        populateDisplayCombo(outCombo);
        outCombo->SetOnSelectionChanged([host = parent, &syncFlag, slot](size_t itemIndex) noexcept
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

            static constexpr Common::Settings::FolderDisplayMode kModes[] = {
                Common::Settings::FolderDisplayMode::Brief,
                Common::Settings::FolderDisplayMode::Detailed,
                Common::Settings::FolderDisplayMode::ExtraDetailed,
                Common::Settings::FolderDisplayMode::Thumbnails,
            };
            if (itemIndex >= std::size(kModes))
            {
                return;
            }

            auto* pane = PrefsFolders::EnsureWorkingFolderPane(state->workingSettings, slot);
            if (! pane)
            {
                return;
            }

            pane->view.display = kModes[itemIndex];
            if (HWND dlg = GetParent(host))
            {
                SetDirty(dlg, *state);
            }
        });
    };

    const auto addSortByCombo = [&](ComboBox*& outCombo, bool& syncFlag, const std::wstring_view slot) noexcept
    {
        outCombo = root->AddChild<ComboBox>();
        outCombo->SetVariant(ComboBoxVariant::Window);
        populateSortByCombo(outCombo);
        outCombo->SetOnSelectionChanged([this, host = parent, &syncFlag, slot](size_t itemIndex) noexcept
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

            static constexpr Common::Settings::FolderSortBy kValues[] = {
                Common::Settings::FolderSortBy::Name,
                Common::Settings::FolderSortBy::Extension,
                Common::Settings::FolderSortBy::Time,
                Common::Settings::FolderSortBy::Size,
                Common::Settings::FolderSortBy::Attributes,
                Common::Settings::FolderSortBy::None,
            };
            if (itemIndex >= std::size(kValues))
            {
                return;
            }

            auto* pane = PrefsFolders::EnsureWorkingFolderPane(state->workingSettings, slot);
            if (! pane)
            {
                return;
            }

            pane->view.sortBy        = kValues[itemIndex];
            pane->view.sortDirection = PrefsFolders::DefaultFolderSortDirection(kValues[itemIndex]);
            if (HWND dlg = GetParent(host))
            {
                SetDirty(dlg, *state);
            }
            Refresh(host, *state);
        });
    };

    const auto addSortDirCombo = [&](ComboBox*& outCombo, bool& syncFlag, const std::wstring_view slot) noexcept
    {
        outCombo = root->AddChild<ComboBox>();
        outCombo->SetVariant(ComboBoxVariant::Window);
        populateSortDirCombo(outCombo);
        outCombo->SetOnSelectionChanged([host = parent, &syncFlag, slot](size_t itemIndex) noexcept
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

            static constexpr Common::Settings::FolderSortDirection kValues[] = {
                Common::Settings::FolderSortDirection::Ascending,
                Common::Settings::FolderSortDirection::Descending,
            };
            if (itemIndex >= std::size(kValues))
            {
                return;
            }

            auto* pane = PrefsFolders::EnsureWorkingFolderPane(state->workingSettings, slot);
            if (! pane)
            {
                return;
            }

            pane->view.sortDirection = kValues[itemIndex];
            if (HWND dlg = GetParent(host))
            {
                SetDirty(dlg, *state);
            }
        });
    };

    const auto addEdit = [&](TextField*& outEdit, bool& syncFlag) noexcept
    {
        outEdit = root->AddChild<TextField>();
        outEdit->SetOnTextChanged([this, host = parent, &syncFlag, field = outEdit](std::wstring_view text) noexcept
        {
            if (syncFlag || ! host || ! field || IsWindow(host) == FALSE)
            {
                return;
            }

            std::wstring normalized;
            normalized.reserve(std::min<size_t>(text.size(), 2u));
            for (const wchar_t ch : text)
            {
                if (! std::iswdigit(static_cast<wint_t>(ch)))
                {
                    continue;
                }
                normalized.push_back(ch);
                if (normalized.size() >= 2u)
                {
                    break;
                }
            }

            if (normalized != text)
            {
                _syncingDxHistoryEdit = true;
                field->SetText(normalized);
                _syncingDxHistoryEdit = false;
            }

            auto* state = PrefsUi::GetDialogState(host);
            if (! state)
            {
                return;
            }

            const auto valueOpt = PrefsUi::TryParseUInt32(normalized);
            if (! valueOpt.has_value())
            {
                return;
            }

            const uint32_t value = valueOpt.value();
            if (value < 1u || value > 50u)
            {
                return;
            }

            auto* folders = PrefsFolders::EnsureWorkingFoldersSettings(state->workingSettings);
            if (! folders)
            {
                return;
            }

            folders->historyMax = value;
            if (HWND dlg = GetParent(host))
            {
                SetDirty(dlg, *state);
            }
        });
        outEdit->SetOnBlur([this, host = parent]() noexcept
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

            const auto text     = _dxState->history.edit->GetText();
            const auto valueOpt = PrefsUi::TryParseUInt32(std::wstring(text));
            if (valueOpt.has_value())
            {
                const uint32_t value = std::clamp(valueOpt.value(), 1u, 50u);
                auto* folders        = PrefsFolders::EnsureWorkingFoldersSettings(state->workingSettings);
                if (folders)
                {
                    folders->historyMax = value;
                    if (HWND dlg = GetParent(host))
                    {
                        SetDirty(dlg, *state);
                    }
                }
            }

            Refresh(host, *state);
        });
    };

    addHeader(dxState->leftHeader);
    addCard(dxState->leftDisplay, true);
    addToggle(dxState->leftDisplay.toggle, IDC_PREFS_PANES_LEFT_DISPLAY_TOGGLE);
    addDisplayCombo(dxState->leftDisplay.combo, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::LeftDisplay)], PrefsFolders::kLeftPaneSlot);
    addCard(dxState->leftSortBy, true);
    addSortByCombo(dxState->leftSortBy.combo, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::LeftSortBy)], PrefsFolders::kLeftPaneSlot);
    addCard(dxState->leftSortDir, false);
    addToggle(dxState->leftSortDir.toggle, IDC_PREFS_PANES_LEFT_SORTDIR_TOGGLE);
    addSortDirCombo(dxState->leftSortDir.combo, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::LeftSortDir)], PrefsFolders::kLeftPaneSlot);
    addCard(dxState->leftStatusBar, true);
    addToggle(dxState->leftStatusBar.toggle, IDC_PREFS_PANES_LEFT_STATUSBAR_TOGGLE);

    addHeader(dxState->rightHeader);
    addCard(dxState->rightDisplay, true);
    addToggle(dxState->rightDisplay.toggle, IDC_PREFS_PANES_RIGHT_DISPLAY_TOGGLE);
    addDisplayCombo(dxState->rightDisplay.combo, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::RightDisplay)], PrefsFolders::kRightPaneSlot);
    addCard(dxState->rightSortBy, true);
    addSortByCombo(dxState->rightSortBy.combo, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::RightSortBy)], PrefsFolders::kRightPaneSlot);
    addCard(dxState->rightSortDir, false);
    addToggle(dxState->rightSortDir.toggle, IDC_PREFS_PANES_RIGHT_SORTDIR_TOGGLE);
    addSortDirCombo(dxState->rightSortDir.combo, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::RightSortDir)], PrefsFolders::kRightPaneSlot);
    addCard(dxState->rightStatusBar, true);
    addToggle(dxState->rightStatusBar.toggle, IDC_PREFS_PANES_RIGHT_STATUSBAR_TOGGLE);

    addHeader(dxState->generalHeader);
    addCard(dxState->showHiddenFiles, false);
    addToggle(dxState->showHiddenFiles.toggle, IDC_PREFS_PANES_SHOW_HIDDEN_TOGGLE);
    addCard(dxState->showSystemFiles, false);
    addToggle(dxState->showSystemFiles.toggle, IDC_PREFS_PANES_SHOW_SYSTEM_TOGGLE);
    addCard(dxState->history, true);
    addEdit(dxState->history.edit, _syncingDxHistoryEdit);

    _dxState = std::move(dxState);
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    return true;
}

void PanesPane::DetachDxHosts() noexcept
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

    _useDxUiTwoStateCombos     = false;
    _usesDxUiTypographyContext = false;
    _usesDxUiTypographyMetrics = false;
    _syncingDxCombos.fill(false);
    _syncingDxHistoryEdit = false;
}

void PanesPane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return;
    }

    _useDxUiTwoStateCombos = state.theme.systemHighContrast;
    _pageHostDx->SetTheme(MakeAppThemeDxPalette(state.theme));
}

void PanesPane::SyncDxControlsFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState)
    {
        return;
    }

    const PrefsFolders::FolderPanePreferences left  = PrefsFolders::GetFolderPanePreferences(state.workingSettings, PrefsFolders::kLeftPaneSlot);
    const PrefsFolders::FolderPanePreferences right = PrefsFolders::GetFolderPanePreferences(state.workingSettings, PrefsFolders::kRightPaneSlot);
    const uint32_t historyMax                       = PrefsFolders::GetFolderHistoryMax(state.workingSettings);
    const bool showHiddenFiles                      = PrefsFolders::GetFolderShowHiddenFiles(state.workingSettings);
    const bool showSystemFiles                      = PrefsFolders::GetFolderShowSystemFiles(state.workingSettings);

    const auto syncLabel = [](Label* label, const UINT stringId) noexcept
    {
        if (label)
        {
            label->SetText(LoadStringResource(nullptr, stringId));
        }
    };

    const auto syncToggle = [](PanesDxCard& card, bool checked, std::wstring uncheckedText, std::wstring checkedText) noexcept
    {
        if (! card.toggle)
        {
            return;
        }

        card.toggle->SetStateLabels(std::move(uncheckedText), std::move(checkedText));
        card.toggle->SetChecked(checked);
        card.toggle->SetEnabled(true);
    };

    syncLabel(_dxState->leftHeader, IDS_PREFS_PANES_HEADER_LEFT);
    syncLabel(_dxState->leftDisplay.title, IDS_PREFS_PANES_LABEL_DISPLAY);
    syncLabel(_dxState->leftDisplay.description, IDS_PREFS_PANES_DESC_DISPLAY);
    syncLabel(_dxState->leftSortBy.title, IDS_PREFS_PANES_LABEL_SORT_BY);
    syncLabel(_dxState->leftSortBy.description, IDS_PREFS_PANES_DESC_SORT);
    syncLabel(_dxState->leftSortDir.title, IDS_PREFS_PANES_LABEL_DIRECTION);
    syncLabel(_dxState->leftStatusBar.title, IDS_PREFS_PANES_LABEL_STATUS_BAR);
    syncLabel(_dxState->leftStatusBar.description, IDS_PREFS_PANES_DESC_STATUS_BAR);
    syncLabel(_dxState->rightHeader, IDS_PREFS_PANES_HEADER_RIGHT);
    syncLabel(_dxState->rightDisplay.title, IDS_PREFS_PANES_LABEL_DISPLAY);
    syncLabel(_dxState->rightDisplay.description, IDS_PREFS_PANES_DESC_DISPLAY);
    syncLabel(_dxState->rightSortBy.title, IDS_PREFS_PANES_LABEL_SORT_BY);
    syncLabel(_dxState->rightSortBy.description, IDS_PREFS_PANES_DESC_SORT);
    syncLabel(_dxState->rightSortDir.title, IDS_PREFS_PANES_LABEL_DIRECTION);
    syncLabel(_dxState->rightStatusBar.title, IDS_PREFS_PANES_LABEL_STATUS_BAR);
    syncLabel(_dxState->rightStatusBar.description, IDS_PREFS_PANES_DESC_STATUS_BAR);
    syncLabel(_dxState->generalHeader, IDS_PREFS_PANES_HEADER_GENERAL);
    syncLabel(_dxState->showHiddenFiles.title, IDS_PREFS_PANES_LABEL_SHOW_HIDDEN_FILES);
    syncLabel(_dxState->showSystemFiles.title, IDS_PREFS_PANES_LABEL_SHOW_SYSTEM_FILES);
    syncLabel(_dxState->history.title, IDS_PREFS_PANES_LABEL_HISTORY_SIZE);
    syncLabel(_dxState->history.description, IDS_PREFS_PANES_DESC_HISTORY_SIZE);

    syncToggle(_dxState->leftDisplay,
               left.display == Common::Settings::FolderDisplayMode::Brief,
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DETAILED),
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_BRIEF));
    syncToggle(_dxState->leftSortDir,
               left.sortDirection == Common::Settings::FolderSortDirection::Ascending,
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DESCENDING),
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_ASCENDING));
    syncToggle(_dxState->leftStatusBar,
               left.statusBarVisible,
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_HIDE),
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_SHOW));
    syncToggle(_dxState->rightDisplay,
               right.display == Common::Settings::FolderDisplayMode::Brief,
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DETAILED),
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_BRIEF));
    syncToggle(_dxState->rightSortDir,
               right.sortDirection == Common::Settings::FolderSortDirection::Ascending,
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DESCENDING),
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_ASCENDING));
    syncToggle(_dxState->rightStatusBar,
               right.statusBarVisible,
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_HIDE),
               LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_SHOW));
    syncToggle(_dxState->showHiddenFiles, showHiddenFiles, LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
    syncToggle(_dxState->showSystemFiles, showSystemFiles, LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));

    const auto syncCombo = [](PanesDxCard& card, auto enumValue, bool& syncFlag) noexcept
    {
        if (! card.combo)
        {
            return;
        }

        syncFlag               = true;
        const auto& items      = card.combo->GetItems();
        const auto targetValue = std::to_wstring(static_cast<int>(enumValue));
        std::optional<size_t> selectedIndex;
        for (size_t i = 0; i < items.size(); ++i)
        {
            if (items[i].value == targetValue)
            {
                selectedIndex = i;
                break;
            }
        }
        card.combo->SetSelectedIndex(selectedIndex);
        card.combo->SetEnabled(true);
        syncFlag = false;
    };

    const auto syncEdit = [](PanesDxCard& card, uint32_t value, bool& syncFlag) noexcept
    {
        if (! card.edit)
        {
            return;
        }

        syncFlag = true;
        card.edit->SetText(std::to_wstring(value));
        card.edit->SetEnabled(true);
        syncFlag = false;
    };

    syncCombo(_dxState->leftDisplay, left.display, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::LeftDisplay)]);
    syncCombo(_dxState->leftSortBy, left.sortBy, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::LeftSortBy)]);
    syncCombo(_dxState->leftSortDir, left.sortDirection, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::LeftSortDir)]);
    syncCombo(_dxState->rightDisplay, right.display, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::RightDisplay)]);
    syncCombo(_dxState->rightSortBy, right.sortBy, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::RightSortBy)]);
    syncCombo(_dxState->rightSortDir, right.sortDirection, _syncingDxCombos[static_cast<size_t>(PanesComboSlot::RightSortDir)]);
    syncEdit(_dxState->history, historyMax, _syncingDxHistoryEdit);
    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void PanesPane::LayoutDxPage(
    PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, int sectionY, const PreferencesTypographyContext& typography) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(margin);

    Debug::Perf::Scope layoutPerf(L"preferences.ui.panes_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(typography.dpi);

    _usesDxUiTypographyContext = true;
    _usesDxUiTypographyMetrics = false;

    const UINT dpi = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);

    const int rowHeight    = std::max(1, UiMetrics::ScaleDip(dpi, kRowHeightDip));
    const int headerHeight = std::max(1, UiMetrics::ScaleDip(dpi, kHeaderHeightDip));

    const std::wstring leftHeaderText     = LoadStringResource(nullptr, IDS_PREFS_PANES_HEADER_LEFT);
    const std::wstring rightHeaderText    = LoadStringResource(nullptr, IDS_PREFS_PANES_HEADER_RIGHT);
    const std::wstring generalHeaderText  = LoadStringResource(nullptr, IDS_PREFS_PANES_HEADER_GENERAL);
    const std::wstring displayLabelText   = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_DISPLAY);
    const std::wstring displayDescText    = LoadStringResource(nullptr, IDS_PREFS_PANES_DESC_DISPLAY);
    const std::wstring sortByLabelText    = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_SORT_BY);
    const std::wstring sortDescText       = LoadStringResource(nullptr, IDS_PREFS_PANES_DESC_SORT);
    const std::wstring directionLabelText = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_DIRECTION);
    const std::wstring statusBarLabelText = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_STATUS_BAR);
    const std::wstring statusBarDescText  = LoadStringResource(nullptr, IDS_PREFS_PANES_DESC_STATUS_BAR);
    const std::wstring showHiddenText     = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_SHOW_HIDDEN_FILES);
    const std::wstring showSystemText     = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_SHOW_SYSTEM_FILES);
    const std::wstring historyLabelText   = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_HISTORY_SIZE);
    const std::wstring historyDescText    = LoadStringResource(nullptr, IDS_PREFS_PANES_DESC_HISTORY_SIZE);
    const std::wstring briefText          = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_BRIEF);
    const std::wstring detailedText       = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DETAILED);
    const std::wstring ascendingText      = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_ASCENDING);
    const std::wstring descendingText     = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DESCENDING);
    const std::wstring hideText           = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_HIDE);
    const std::wstring showText           = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_SHOW);

    DxState* dxState = _dxState.get();
    if (dxState && _pageHostDx && _pageContentRoot)
    {
        const int titleHeight  = std::max(1, UiMetrics::ScaleDip(dpi, kTitleHeightDip));
        const int cardPaddingX = UiMetrics::ScaleDip(dpi, kCardPaddingXDip);
        const int cardPaddingY = UiMetrics::ScaleDip(dpi, kCardPaddingYDip);
        const int cardGapY     = UiMetrics::ScaleDip(dpi, kCardGapYDip);
        const int cardGapX     = UiMetrics::ScaleDip(dpi, kCardGapXDip);
        const int cardSpacingY = UiMetrics::ScaleDip(dpi, kCardSpacingYDip);
        int contentBottom      = y;

        const auto pxToDip = [dpi](const int px) noexcept { return static_cast<float>(px) * 96.0f / static_cast<float>(std::max<UINT>(1u, dpi)); };

        const auto hideCardControls = [](PanesDxCard& card) noexcept
        {
            if (card.card)
            {
                card.card->SetVisible(false);
            }
            if (card.title)
            {
                card.title->SetVisible(false);
            }
            if (card.description)
            {
                card.description->SetVisible(false);
            }
            if (card.toggle)
            {
                card.toggle->SetVisible(false);
            }
            if (card.combo)
            {
                card.combo->SetVisible(false);
            }
            if (card.edit)
            {
                card.edit->SetVisible(false);
            }
        };

        hideCardControls(dxState->leftDisplay);
        hideCardControls(dxState->leftSortBy);
        hideCardControls(dxState->leftSortDir);
        hideCardControls(dxState->leftStatusBar);
        hideCardControls(dxState->rightDisplay);
        hideCardControls(dxState->rightSortBy);
        hideCardControls(dxState->rightSortDir);
        hideCardControls(dxState->rightStatusBar);
        hideCardControls(dxState->showHiddenFiles);
        hideCardControls(dxState->showSystemFiles);
        hideCardControls(dxState->history);

        const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
        const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
        const int onWidth           = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, onLabel);
        const int offWidth          = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, offLabel);
        const int briefWidth        = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, briefText);
        const int detailedWidth     = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, detailedText);
        const int ascendingWidth    = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, ascendingText);
        const int descendingWidth   = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, descendingText);
        const int paddingX          = UiMetrics::ScaleDip(dpi, kTogglePaddingXDip);
        const int stateGapX         = UiMetrics::ScaleDip(dpi, kToggleGapXDip);
        const int trackWidth        = UiMetrics::ScaleDip(dpi, kToggleTrackWidthDip);
        const int stateTextWidth    = (std::max)({onWidth, offWidth, briefWidth, detailedWidth, ascendingWidth, descendingWidth});
        _usesDxUiTypographyMetrics  = stateTextWidth > 0;
        const int toggleSlackWidth  = UiMetrics::ScaleDip(dpi, 12);
        const int measuredSwitchWidth =
            std::max(UiMetrics::ScaleDip(dpi, kMinToggleWidthDip), (2 * paddingX) + stateTextWidth + stateGapX + trackWidth + toggleSlackWidth);
        const int maxControlWidth    = std::max(0, width - 2 * cardPaddingX);
        const int compactSwitchWidth = static_cast<int>(RedSalamander::DxUi::ResolveConstrainedExtent(
            {.preferredExtent = static_cast<float>(std::max(UiMetrics::ScaleDip(dpi, 44), trackWidth + (2 * paddingX)))}, static_cast<float>(maxControlWidth)));
        const int switchWidth        = static_cast<int>(RedSalamander::DxUi::ResolveConstrainedExtent(
            {.minExtent = static_cast<float>(UiMetrics::ScaleDip(dpi, kMinToggleWidthDip)), .preferredExtent = static_cast<float>(measuredSwitchWidth)},
            static_cast<float>(maxControlWidth)));

        const auto measureDescriptionHeight = [&](const int textWidth, const std::wstring& text) noexcept
        {
            const int height           = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, text);
            _usesDxUiTypographyMetrics = _usesDxUiTypographyMetrics || height > 0;
            return height;
        };

        const auto layoutHeaderDx = [&](Label* header, const std::wstring& text) noexcept
        {
            if (! header)
            {
                return;
            }

            header->SetVisible(true);
            header->SetText(text);
            header->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + headerHeight)));
            contentBottom = std::max(contentBottom, y + headerHeight);
            y += headerHeight + gapY;
        };

        const auto layoutToggleCardDx =
            [&](PanesDxCard& card, const std::wstring& titleText, const std::wstring& descText, std::wstring uncheckedLabel, std::wstring checkedLabel) noexcept
        {
            const bool compactSwitch    = uncheckedLabel.empty() && checkedLabel.empty();
            const int actualSwitchWidth = compactSwitch ? std::min(compactSwitchWidth, maxControlWidth) : switchWidth;
            const bool hasDesc          = card.description && ! descText.empty();
            const int textWidth         = std::max(0, width - 2 * cardPaddingX - cardGapX - actualSwitchWidth);
            const int descHeight        = hasDesc ? measureDescriptionHeight(textWidth, descText) : 0;
            const int contentHeight     = hasDesc ? (titleHeight + cardGapY + descHeight) : titleHeight;
            const int cardHeight        = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);
            const int localCardTop      = y;
            const int titleTop          = hasDesc ? (localCardTop + cardPaddingY) : (localCardTop + (cardHeight - titleHeight) / 2);

            if (card.card)
            {
                card.card->SetVisible(true);
                card.card->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(localCardTop), pxToDip(x + width), pxToDip(localCardTop + cardHeight)));
            }
            if (card.title)
            {
                card.title->SetVisible(true);
                card.title->SetText(titleText);
                card.title->SetMnemonicTarget(card.toggle);
                card.title->SetBounds(
                    D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(titleTop), pxToDip(x + cardPaddingX + textWidth), pxToDip(titleTop + titleHeight)));
            }
            if (card.description)
            {
                card.description->SetVisible(hasDesc);
                if (hasDesc)
                {
                    card.description->SetText(descText);
                    card.description->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                            pxToDip(localCardTop + cardPaddingY + titleHeight + cardGapY),
                                                            pxToDip(x + cardPaddingX + textWidth),
                                                            pxToDip(localCardTop + cardPaddingY + titleHeight + cardGapY + descHeight)));
                }
            }
            if (card.toggle)
            {
                card.toggle->SetVisible(true);
                card.toggle->SetStateLabels(std::move(uncheckedLabel), std::move(checkedLabel));
                card.toggle->SetBounds(D2D1::RectF(pxToDip(x + width - cardPaddingX - actualSwitchWidth),
                                                   pxToDip(localCardTop + (cardHeight - rowHeight) / 2),
                                                   pxToDip(x + width - cardPaddingX),
                                                   pxToDip(localCardTop + (cardHeight - rowHeight) / 2 + rowHeight)));
            }
            if (card.combo)
            {
                card.combo->SetVisible(false);
            }
            if (card.edit)
            {
                card.edit->SetVisible(false);
            }

            contentBottom = std::max(contentBottom, y + cardHeight);
            y += cardHeight + cardSpacingY;
        };

        const auto layoutComboCardDx = [&](PanesDxCard& card, const std::wstring& titleText, const std::wstring& descText) noexcept
        {
            int desiredWidth        = UiMetrics::ScaleDip(dpi, kMinEditWidthDip + 10);
            desiredWidth            = std::min(desiredWidth, std::min(maxControlWidth, UiMetrics::ScaleDip(dpi, kMaxEditWidthDip)));
            const bool hasDesc      = card.description && ! descText.empty();
            const int textWidth     = std::max(0, width - 2 * cardPaddingX - cardGapX - desiredWidth);
            const int descHeight    = hasDesc ? measureDescriptionHeight(textWidth, descText) : 0;
            const int contentHeight = hasDesc ? std::max(rowHeight, titleHeight + cardGapY + descHeight) : rowHeight;
            const int cardHeight    = contentHeight + 2 * cardPaddingY;
            const int localCardTop  = y;
            const int titleY        = hasDesc ? (localCardTop + cardPaddingY) : (localCardTop + cardPaddingY + (rowHeight - titleHeight) / 2);
            const int comboY        = localCardTop + (cardHeight - rowHeight) / 2;

            if (card.card)
            {
                card.card->SetVisible(true);
                card.card->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(localCardTop), pxToDip(x + width), pxToDip(localCardTop + cardHeight)));
            }
            if (card.title)
            {
                card.title->SetVisible(true);
                card.title->SetText(titleText);
                card.title->SetMnemonicTarget(card.combo);
                card.title->SetBounds(
                    D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(titleY), pxToDip(x + cardPaddingX + textWidth), pxToDip(titleY + titleHeight)));
            }
            if (card.description)
            {
                card.description->SetVisible(hasDesc);
                if (hasDesc)
                {
                    card.description->SetText(descText);
                    card.description->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                            pxToDip(localCardTop + cardPaddingY + titleHeight + cardGapY),
                                                            pxToDip(x + cardPaddingX + textWidth),
                                                            pxToDip(localCardTop + cardPaddingY + titleHeight + cardGapY + descHeight)));
                }
            }
            if (card.toggle)
            {
                card.toggle->SetVisible(false);
            }
            if (card.combo)
            {
                card.combo->SetVisible(true);
                card.combo->SetBounds(D2D1::RectF(
                    pxToDip(x + width - cardPaddingX - desiredWidth), pxToDip(comboY), pxToDip(x + width - cardPaddingX), pxToDip(comboY + rowHeight)));
            }
            if (card.edit)
            {
                card.edit->SetVisible(false);
            }

            contentBottom = std::max(contentBottom, y + cardHeight);
            y += cardHeight + cardSpacingY;
        };

        const auto layoutSortCardDx = [&](PanesDxCard& sortCard, PanesDxCard& directionCard) noexcept
        {
            int comboWidth                  = UiMetrics::ScaleDip(dpi, kMinEditWidthDip + 10);
            comboWidth                      = std::min(comboWidth, std::min(maxControlWidth, UiMetrics::ScaleDip(dpi, kMaxEditWidthDip)));
            const int directionControlWidth = state.theme.systemHighContrast ? comboWidth : switchWidth;
            const int controlWidth          = std::max(comboWidth, directionControlWidth);
            const bool hasDesc              = sortCard.description && ! sortDescText.empty();
            const int textWidth             = std::max(0, width - 2 * cardPaddingX - cardGapX - controlWidth);
            const int descHeight            = hasDesc ? measureDescriptionHeight(textWidth, sortDescText) : 0;
            const int topBlockHeight        = hasDesc ? std::max(rowHeight, titleHeight + cardGapY + descHeight) : rowHeight;
            const int contentHeight         = topBlockHeight + cardGapY + rowHeight;
            const int cardHeight            = contentHeight + 2 * cardPaddingY;
            const int localCardTop          = y;
            const int topRowY               = localCardTop + cardPaddingY + (topBlockHeight - rowHeight) / 2;
            const int titleY                = hasDesc ? (localCardTop + cardPaddingY) : (localCardTop + cardPaddingY + (rowHeight - titleHeight) / 2);
            const int directionRowY         = localCardTop + cardPaddingY + topBlockHeight + cardGapY;

            if (sortCard.card)
            {
                sortCard.card->SetVisible(true);
                sortCard.card->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(localCardTop), pxToDip(x + width), pxToDip(localCardTop + cardHeight)));
            }
            if (sortCard.title)
            {
                sortCard.title->SetVisible(true);
                sortCard.title->SetText(sortByLabelText);
                sortCard.title->SetMnemonicTarget(sortCard.combo);
                sortCard.title->SetBounds(
                    D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(titleY), pxToDip(x + cardPaddingX + textWidth), pxToDip(titleY + titleHeight)));
            }
            if (sortCard.description)
            {
                sortCard.description->SetVisible(hasDesc);
                if (hasDesc)
                {
                    sortCard.description->SetText(sortDescText);
                    sortCard.description->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                                pxToDip(localCardTop + cardPaddingY + titleHeight + cardGapY),
                                                                pxToDip(x + cardPaddingX + textWidth),
                                                                pxToDip(localCardTop + cardPaddingY + titleHeight + cardGapY + descHeight)));
                }
            }
            if (sortCard.combo)
            {
                sortCard.combo->SetVisible(true);
                sortCard.combo->SetBounds(D2D1::RectF(
                    pxToDip(x + width - cardPaddingX - comboWidth), pxToDip(topRowY), pxToDip(x + width - cardPaddingX), pxToDip(topRowY + rowHeight)));
            }
            if (sortCard.toggle)
            {
                sortCard.toggle->SetVisible(false);
            }
            if (sortCard.edit)
            {
                sortCard.edit->SetVisible(false);
            }

            if (directionCard.card)
            {
                directionCard.card->SetVisible(false);
            }
            if (directionCard.title)
            {
                directionCard.title->SetVisible(true);
                directionCard.title->SetText(directionLabelText);
                directionCard.title->SetMnemonicTarget(state.theme.systemHighContrast ? static_cast<RedSalamander::DxUi::Control*>(directionCard.combo)
                                                                                      : static_cast<RedSalamander::DxUi::Control*>(directionCard.toggle));
                directionCard.title->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                           pxToDip(directionRowY + (rowHeight - titleHeight) / 2),
                                                           pxToDip(x + cardPaddingX + textWidth),
                                                           pxToDip(directionRowY + (rowHeight - titleHeight) / 2 + titleHeight)));
            }
            if (directionCard.description)
            {
                directionCard.description->SetVisible(false);
            }
            if (directionCard.toggle)
            {
                directionCard.toggle->SetVisible(! state.theme.systemHighContrast);
                if (! state.theme.systemHighContrast)
                {
                    directionCard.toggle->SetStateLabels(descendingText, ascendingText);
                    directionCard.toggle->SetBounds(D2D1::RectF(pxToDip(x + width - cardPaddingX - directionControlWidth),
                                                                pxToDip(directionRowY),
                                                                pxToDip(x + width - cardPaddingX),
                                                                pxToDip(directionRowY + rowHeight)));
                }
            }
            if (directionCard.combo)
            {
                directionCard.combo->SetVisible(state.theme.systemHighContrast);
                if (state.theme.systemHighContrast)
                {
                    directionCard.combo->SetBounds(D2D1::RectF(pxToDip(x + width - cardPaddingX - directionControlWidth),
                                                               pxToDip(directionRowY),
                                                               pxToDip(x + width - cardPaddingX),
                                                               pxToDip(directionRowY + rowHeight)));
                }
            }
            if (directionCard.edit)
            {
                directionCard.edit->SetVisible(false);
            }

            contentBottom = std::max(contentBottom, y + cardHeight);
            y += cardHeight + cardSpacingY;
        };

        const auto layoutEditCardDx = [&](PanesDxCard& card, const std::wstring& titleText, const std::wstring& descText) noexcept
        {
            const int desiredWidth  = std::min(maxControlWidth, UiMetrics::ScaleDip(dpi, kMinComboWidthDip));
            const int textWidth     = std::max(0, width - 2 * cardPaddingX - cardGapX - desiredWidth);
            const int descHeight    = card.description ? measureDescriptionHeight(textWidth, descText) : 0;
            const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
            const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);
            const int localCardTop  = y;

            if (card.card)
            {
                card.card->SetVisible(true);
                card.card->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(localCardTop), pxToDip(x + width), pxToDip(localCardTop + cardHeight)));
            }
            if (card.title)
            {
                card.title->SetVisible(true);
                card.title->SetText(titleText);
                card.title->SetMnemonicTarget(card.edit);
                card.title->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                  pxToDip(localCardTop + cardPaddingY),
                                                  pxToDip(x + cardPaddingX + textWidth),
                                                  pxToDip(localCardTop + cardPaddingY + titleHeight)));
            }
            if (card.description)
            {
                card.description->SetVisible(true);
                card.description->SetText(descText);
                card.description->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                        pxToDip(localCardTop + cardPaddingY + titleHeight + cardGapY),
                                                        pxToDip(x + cardPaddingX + textWidth),
                                                        pxToDip(localCardTop + cardPaddingY + titleHeight + cardGapY + descHeight)));
            }
            if (card.toggle)
            {
                card.toggle->SetVisible(false);
            }
            if (card.combo)
            {
                card.combo->SetVisible(false);
            }
            if (card.edit)
            {
                card.edit->SetVisible(true);
                card.edit->SetBounds(D2D1::RectF(pxToDip(x + width - cardPaddingX - desiredWidth),
                                                 pxToDip(localCardTop + (cardHeight - rowHeight) / 2),
                                                 pxToDip(x + width - cardPaddingX),
                                                 pxToDip(localCardTop + (cardHeight - rowHeight) / 2 + rowHeight)));
            }

            contentBottom = std::max(contentBottom, y + cardHeight);
            y += cardHeight + cardSpacingY;
        };

        layoutHeaderDx(dxState->leftHeader, leftHeaderText);
        if (state.theme.systemHighContrast)
        {
            layoutComboCardDx(dxState->leftDisplay, displayLabelText, displayDescText);
        }
        else
        {
            layoutToggleCardDx(dxState->leftDisplay, displayLabelText, displayDescText, detailedText, briefText);
        }
        layoutSortCardDx(dxState->leftSortBy, dxState->leftSortDir);
        layoutToggleCardDx(dxState->leftStatusBar, statusBarLabelText, statusBarDescText, hideText, showText);

        y += std::max(0, sectionY - cardSpacingY);

        layoutHeaderDx(dxState->rightHeader, rightHeaderText);
        if (state.theme.systemHighContrast)
        {
            layoutComboCardDx(dxState->rightDisplay, displayLabelText, displayDescText);
        }
        else
        {
            layoutToggleCardDx(dxState->rightDisplay, displayLabelText, displayDescText, detailedText, briefText);
        }
        layoutSortCardDx(dxState->rightSortBy, dxState->rightSortDir);
        layoutToggleCardDx(dxState->rightStatusBar, statusBarLabelText, statusBarDescText, hideText, showText);

        y += std::max(0, sectionY - cardSpacingY);

        layoutHeaderDx(dxState->generalHeader, generalHeaderText);
        layoutToggleCardDx(dxState->showHiddenFiles, showHiddenText, {}, offLabel, onLabel);
        layoutToggleCardDx(dxState->showSystemFiles, showSystemText, {}, offLabel, onLabel);
        layoutEditCardDx(dxState->history, historyLabelText, historyDescText);

        _pageHostDx->Invalidate();
        return;
    }
}

void PanesPane::LayoutPage(HWND host,
                           PreferencesDialogState& state,
                           int x,
                           int& y,
                           int width,
                           int margin,
                           int gapY,
                           int sectionY,
                           const PreferencesTypographyContext& typography) noexcept
{
    if (! host)
    {
        return;
    }

    if (EnsureDxHosts(_pageHost ? _pageHost : host, state))
    {
        LayoutDxPage(state, x, y, width, margin, gapY, sectionY, typography);
        return;
    }

    Debug::Error(L"Preferences.Panes: DxUi surface initialization failed; page will not render correctly.");
}

void PanesPane::Refresh(HWND /*host*/, PreferencesDialogState& state) noexcept
{
    if (state.currentCategory == PrefCategory::Panes && ! _dxState)
    {
        static_cast<void>(EnsureDxHosts(_pageHost, state));
    }

    state.refreshingPanesPage = true;
    const auto reset          = wil::scope_exit([&] { state.refreshingPanesPage = false; });

    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
}

void PanesPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost = parent;

    if (state.currentCategory != PrefCategory::Panes)
    {
        return;
    }

    if (! EnsureDxHosts(parent, state))
    {
        Debug::Error(L"Preferences.Panes: DxUi surface initialization failed; page will not render correctly.");
        DetachDxHosts();
        return;
    }
}

#ifdef ENABLE_TESTS
PreferencesPanesDebugFocusTarget PanesPane::DebugGetFocusTarget() const noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return PreferencesPanesDebugFocusTarget::None;
    }

    const RedSalamander::DxUi::Control* const focusedControl = _pageHostDx->GetFocusControl();
    if (! focusedControl)
    {
        return PreferencesPanesDebugFocusTarget::None;
    }

    const auto& dxState = *_dxState;
    if (focusedControl == dxState.leftDisplay.toggle)
        return PreferencesPanesDebugFocusTarget::LeftDisplayToggle;
    if (focusedControl == dxState.leftDisplay.combo)
        return PreferencesPanesDebugFocusTarget::LeftDisplayCombo;
    if (focusedControl == dxState.leftSortBy.combo)
        return PreferencesPanesDebugFocusTarget::LeftSortByCombo;
    if (focusedControl == dxState.leftSortDir.toggle)
        return PreferencesPanesDebugFocusTarget::LeftSortDirToggle;
    if (focusedControl == dxState.leftSortDir.combo)
        return PreferencesPanesDebugFocusTarget::LeftSortDirCombo;
    if (focusedControl == dxState.leftStatusBar.toggle)
        return PreferencesPanesDebugFocusTarget::LeftStatusBarToggle;
    if (focusedControl == dxState.rightDisplay.toggle)
        return PreferencesPanesDebugFocusTarget::RightDisplayToggle;
    if (focusedControl == dxState.rightDisplay.combo)
        return PreferencesPanesDebugFocusTarget::RightDisplayCombo;
    if (focusedControl == dxState.rightSortBy.combo)
        return PreferencesPanesDebugFocusTarget::RightSortByCombo;
    if (focusedControl == dxState.rightSortDir.toggle)
        return PreferencesPanesDebugFocusTarget::RightSortDirToggle;
    if (focusedControl == dxState.rightSortDir.combo)
        return PreferencesPanesDebugFocusTarget::RightSortDirCombo;
    if (focusedControl == dxState.rightStatusBar.toggle)
        return PreferencesPanesDebugFocusTarget::RightStatusBarToggle;
    if (focusedControl == dxState.showHiddenFiles.toggle)
        return PreferencesPanesDebugFocusTarget::ShowHiddenFilesToggle;
    if (focusedControl == dxState.showSystemFiles.toggle)
        return PreferencesPanesDebugFocusTarget::ShowSystemFilesToggle;
    if (focusedControl == dxState.history.edit)
        return PreferencesPanesDebugFocusTarget::HistoryField;
    return PreferencesPanesDebugFocusTarget::None;
}

bool PanesPane::DebugUsesDxUiTypographyContext() const noexcept
{
    return _usesDxUiTypographyContext;
}

bool PanesPane::DebugUsesDxUiTypographyMetrics() const noexcept
{
    return _usesDxUiTypographyMetrics;
}

bool PanesPane::DebugFocusLeftDisplayToggle() noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return false;
    }

    RedSalamander::DxUi::Control* focusTarget = nullptr;
    if (_dxState->leftDisplay.toggle && _dxState->leftDisplay.toggle->IsVisible())
    {
        focusTarget = _dxState->leftDisplay.toggle;
    }
    else if (_dxState->leftDisplay.combo && _dxState->leftDisplay.combo->IsVisible())
    {
        focusTarget = _dxState->leftDisplay.combo;
    }
    if (! focusTarget)
    {
        return false;
    }

    _pageHostDx->SetFocusControl(focusTarget);
    return true;
}

bool PanesPane::DebugSelectLeftDisplayByText(std::wstring_view displayText) noexcept
{
    if (! _dxState || ! _pageHostDx || ! _pageHost || IsWindow(_pageHost) == FALSE)
    {
        return false;
    }

    auto* state = PrefsUi::GetDialogState(_pageHost);
    if (! state)
    {
        return false;
    }

    Common::Settings::FolderDisplayMode targetMode{};
    if (displayText == LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_BRIEF))
    {
        targetMode = Common::Settings::FolderDisplayMode::Brief;
    }
    else if (displayText == LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DETAILED))
    {
        targetMode = Common::Settings::FolderDisplayMode::Detailed;
    }
    else if (displayText == LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_EXTRA_DETAILED))
    {
        targetMode = Common::Settings::FolderDisplayMode::ExtraDetailed;
    }
    else if (displayText == LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_THUMBNAILS))
    {
        targetMode = Common::Settings::FolderDisplayMode::Thumbnails;
    }
    else
    {
        return false;
    }

    auto* pane = PrefsFolders::EnsureWorkingFolderPane(state->workingSettings, PrefsFolders::kLeftPaneSlot);
    if (! pane)
    {
        return false;
    }

    pane->view.display = targetMode;
    if (HWND dlg = GetParent(_pageHost))
    {
        SetDirty(dlg, *state);
    }

    Refresh(_pageHost, *state);

    auto* const combo = _dxState->leftDisplay.combo;
    return pane->view.display == targetMode && (! combo || combo->GetDisplayedText() == displayText);
}

bool PanesPane::DebugFocusLeftStatusBarToggle() noexcept
{
    if (! _dxState || ! _pageHostDx || ! _dxState->leftStatusBar.toggle || ! _dxState->leftStatusBar.toggle->IsVisible())
    {
        return false;
    }

    _pageHostDx->SetFocusControl(_dxState->leftStatusBar.toggle);
    return true;
}

bool PanesPane::DebugGetLeftStatusBarToggleChecked(bool& outChecked) const noexcept
{
    if (! _dxState || ! _dxState->leftStatusBar.toggle || ! _dxState->leftStatusBar.toggle->IsVisible())
    {
        return false;
    }

    outChecked = _dxState->leftStatusBar.toggle->IsChecked();
    return true;
}
#endif
