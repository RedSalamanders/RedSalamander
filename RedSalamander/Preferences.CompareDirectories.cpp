// Preferences.CompareDirectories.cpp

#include "Framework.h"

#include "Preferences.CompareDirectories.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <span>
#include <string>
#include <vector>

#include "DxUi/DxUi.h"
#include "Helpers.h"
#include "UiMetrics.h"
#include "resource.h"

// Local convenience aliases for frequently-used shared utilities
namespace
{
using PrefsCompareDirectories::EnsureWorkingCompareDirectoriesSettings;
using PrefsCompareDirectories::GetCompareDirectoriesSettingsOrDefault;
using PrefsCompareDirectories::MaybeResetWorkingCompareDirectoriesSettingsIfEmpty;

using RedSalamander::DxUi::CardPanel;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;
constexpr std::array<UINT, 5> kCompareHeaderStringIds = {{
    IDS_PREFS_ADV_HEADER_COMPARE_DIRECTORIES,
    IDS_COMPARE_OPTIONS_SECTION_SUBDIRS,
    IDS_COMPARE_OPTIONS_SECTION_COMPARE,
    IDS_COMPARE_OPTIONS_SECTION_ADVANCED,
    IDS_COMPARE_OPTIONS_SECTION_IGNORE,
}};

constexpr std::array<UINT, 11> kCompareToggleLabelStringIds = {{
    IDS_COMPARE_OPTIONS_SUBDIRS_TITLE,
    IDS_COMPARE_OPTIONS_SIZE_TITLE,
    IDS_COMPARE_OPTIONS_DATETIME_TITLE,
    IDS_COMPARE_OPTIONS_ATTRIBUTES_TITLE,
    IDS_COMPARE_OPTIONS_CONTENT_TITLE,
    IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_TITLE,
    IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_TITLE,
    IDS_PREFS_COMPARE_KEEP_IDENTICAL_TITLE,
    IDS_PREFS_COMPARE_SHOW_IDENTICAL_TITLE,
    IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE,
    IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_TITLE,
}};

constexpr std::array<UINT, 11> kCompareToggleDescriptionStringIds = {{
    IDS_COMPARE_OPTIONS_SUBDIRS_DESC,
    IDS_COMPARE_OPTIONS_SIZE_DESC,
    IDS_COMPARE_OPTIONS_DATETIME_DESC,
    IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC,
    IDS_COMPARE_OPTIONS_CONTENT_DESC,
    IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC,
    IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC,
    IDS_PREFS_COMPARE_KEEP_IDENTICAL_DESC,
    IDS_PREFS_COMPARE_SHOW_IDENTICAL_DESC,
    IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC,
    IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC,
}};

constexpr std::array<UINT, 11> kCompareToggleCommandIds = {{
    IDC_PREFS_ADV_COMPARE_SUBDIRS_TOGGLE,
    IDC_PREFS_ADV_COMPARE_SIZE_TOGGLE,
    IDC_PREFS_ADV_COMPARE_DATETIME_TOGGLE,
    IDC_PREFS_ADV_COMPARE_ATTRIBUTES_TOGGLE,
    IDC_PREFS_ADV_COMPARE_CONTENT_TOGGLE,
    IDC_PREFS_ADV_COMPARE_SUBDIR_ATTRIBUTES_TOGGLE,
    IDC_PREFS_ADV_COMPARE_SELECT_SUBDIRS_ONE_PANE_TOGGLE,
    IDC_PREFS_ADV_COMPARE_KEEP_IDENTICAL_TOGGLE,
    IDC_PREFS_ADV_COMPARE_SHOW_IDENTICAL_TOGGLE,
    IDC_PREFS_ADV_COMPARE_IGNORE_FILES_TOGGLE,
    IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_TOGGLE,
}};

constexpr size_t kCompareIgnoreFilesIndex       = 9u;
constexpr size_t kCompareIgnoreDirectoriesIndex = 10u;

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

    for (RedSalamander::DxUi::Control* const control : orderedControls)
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

struct CompareToggleCardPageDx
{
    CardPanel* card    = nullptr;
    Label* title       = nullptr;
    Label* description = nullptr;
    Toggle* toggle     = nullptr;
    TextField* edit    = nullptr;
};

struct CompareComboCardPageDx
{
    CardPanel* card    = nullptr;
    Label* title       = nullptr;
    Label* description = nullptr;
    ComboBox* combo    = nullptr;
};

struct CompareDxPage
{
    CompareDxPage()                                = default;
    CompareDxPage(const CompareDxPage&)            = delete;
    CompareDxPage& operator=(const CompareDxPage&) = delete;
    CompareDxPage(CompareDxPage&&)                 = delete;
    CompareDxPage& operator=(CompareDxPage&&)      = delete;

    std::array<Label*, kCompareHeaderStringIds.size()> headers{};
    std::array<CompareToggleCardPageDx, kCompareToggleLabelStringIds.size()> toggleCards{};
    CompareComboCardPageDx contentWorkers;

    void Detach() noexcept
    {
        headers.fill(nullptr);
        for (auto& card : toggleCards)
        {
            card = {};
        }
        contentWorkers = {};
    }
};

struct CompareDirectoriesPane::DxState
{
    DxState()                          = default;
    DxState(const DxState&)            = delete;
    DxState& operator=(const DxState&) = delete;
    DxState(DxState&&)                 = delete;
    DxState& operator=(DxState&&)      = delete;

    CompareDxPage page;

    void Detach() noexcept
    {
        page.Detach();
    }
};

CompareDirectoriesPane::CompareDirectoriesPane() = default;

CompareDirectoriesPane::~CompareDirectoriesPane()
{
    DetachDxHosts();
}

void CompareDirectoriesPane::OnVisibilityChanged(bool visible) noexcept
{
    static_cast<void>(visible);
}

void CompareDirectoriesPane::Destroy(PreferencesDialogState& state) noexcept
{
    static_cast<void>(state);
    DetachDxHosts();
    _state    = nullptr;
    _pageHost = nullptr;
}

bool CompareDirectoriesPane::EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept
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

    for (Label*& header : dxState->page.headers)
    {
        header = root->AddChild<Label>();
        header->SetFontRole(FontRole::Header);
    }

    for (size_t i = 0; i < dxState->page.toggleCards.size(); ++i)
    {
        CompareToggleCardPageDx& card = dxState->page.toggleCards[i];
        const UINT commandId          = kCompareToggleCommandIds[i];

        card.card  = root->AddChild<CardPanel>();
        card.title = root->AddChild<Label>();
        card.title->SetFontRole(FontRole::Body);
        card.description = root->AddChild<Label>();
        card.description->SetFontRole(FontRole::Small);
        card.description->SetMultiline(true);
        card.toggle = root->AddChild<Toggle>();
        card.toggle->SetStateLabels(LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
        card.toggle->SetOnToggled([this, host = parent, commandId](bool checked) noexcept
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

            auto* compare = EnsureWorkingCompareDirectoriesSettings(dialogState->workingSettings);
            if (! compare)
            {
                return;
            }

            DeferredFocusTarget focusAfterLayout = DeferredFocusTarget::None;
            if (_pageHostDx && _dxState)
            {
                const RedSalamander::DxUi::Control* const focused = _pageHostDx->GetFocusControl();
                if (commandId == IDC_PREFS_ADV_COMPARE_IGNORE_FILES_TOGGLE && focused == _dxState->page.toggleCards[kCompareIgnoreFilesIndex].toggle)
                {
                    focusAfterLayout = DeferredFocusTarget::IgnoreFilesToggle;
                }
                else if (commandId == IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_TOGGLE &&
                         focused == _dxState->page.toggleCards[kCompareIgnoreDirectoriesIndex].toggle)
                {
                    focusAfterLayout = DeferredFocusTarget::IgnoreDirectoriesToggle;
                }
            }

            switch (commandId)
            {
                case IDC_PREFS_ADV_COMPARE_SIZE_TOGGLE: compare->compareSize = checked; break;
                case IDC_PREFS_ADV_COMPARE_DATETIME_TOGGLE: compare->compareDateTime = checked; break;
                case IDC_PREFS_ADV_COMPARE_ATTRIBUTES_TOGGLE: compare->compareAttributes = checked; break;
                case IDC_PREFS_ADV_COMPARE_CONTENT_TOGGLE: compare->compareContent = checked; break;
                case IDC_PREFS_ADV_COMPARE_SUBDIRS_TOGGLE: compare->compareSubdirectories = checked; break;
                case IDC_PREFS_ADV_COMPARE_SUBDIR_ATTRIBUTES_TOGGLE: compare->compareSubdirectoryAttributes = checked; break;
                case IDC_PREFS_ADV_COMPARE_SELECT_SUBDIRS_ONE_PANE_TOGGLE: compare->selectSubdirsOnlyInOnePane = checked; break;
                case IDC_PREFS_ADV_COMPARE_KEEP_IDENTICAL_TOGGLE:
                    compare->keepIdenticalItems = checked;
                    if (! checked)
                    {
                        compare->showIdenticalItems = false;
                    }
                    break;
                case IDC_PREFS_ADV_COMPARE_SHOW_IDENTICAL_TOGGLE:
                    compare->showIdenticalItems = checked;
                    if (checked)
                    {
                        compare->keepIdenticalItems = true;
                    }
                    break;
                case IDC_PREFS_ADV_COMPARE_IGNORE_FILES_TOGGLE: compare->ignoreFiles = checked; break;
                case IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_TOGGLE: compare->ignoreDirectories = checked; break;
                default: return;
            }

            if (focusAfterLayout != DeferredFocusTarget::None)
            {
                _deferredFocusAfterLayout = focusAfterLayout;
            }

            MaybeResetWorkingCompareDirectoriesSettingsIfEmpty(dialogState->workingSettings);
            SetDirty(GetParent(host), *dialogState);
            Refresh(host, *dialogState);
            RestoreDeferredFocusTarget(focusAfterLayout);

            if (commandId == IDC_PREFS_ADV_COMPARE_IGNORE_FILES_TOGGLE || commandId == IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_TOGGLE)
            {
                // Post a deferred relayout so LayoutPreferencesPageHost shows/hides the
                // pattern edit field.  SendMessageW(host, WM_SIZE) does NOT work: the
                // DxUi host intercepts WM_SIZE before LayoutPreferencesPageHost is reached.
                static_cast<void>(PrefsUi::PostDeferredAction(host, PreferencesDeferredActionKind::CompareDirectoriesIgnoreToggleChanged));
            }
        });

        if (i == kCompareIgnoreFilesIndex || i == kCompareIgnoreDirectoriesIndex)
        {
            TextField* edit = root->AddChild<TextField>();
            card.edit       = edit;

            bool* syncFlag = (i == kCompareIgnoreFilesIndex) ? &_syncingDxIgnoreFilesEdit : &_syncingDxIgnoreDirectoriesEdit;

            edit->SetOnTextChanged([this, host = parent, syncFlag, isFiles = (i == kCompareIgnoreFilesIndex)](std::wstring_view text) noexcept
            {
                if (! syncFlag || *syncFlag || ! _state)
                {
                    return;
                }

                auto* compare = EnsureWorkingCompareDirectoriesSettings(_state->workingSettings);
                if (! compare)
                {
                    return;
                }

                const std::wstring newValue(text);
                bool changed = false;
                if (isFiles)
                {
                    if (compare->ignoreFilesPatterns != newValue)
                    {
                        compare->ignoreFilesPatterns = newValue;
                        changed                      = true;
                    }
                }
                else
                {
                    if (compare->ignoreDirectoriesPatterns != newValue)
                    {
                        compare->ignoreDirectoriesPatterns = newValue;
                        changed                            = true;
                    }
                }

                if (changed)
                {
                    MaybeResetWorkingCompareDirectoriesSettingsIfEmpty(_state->workingSettings);
                    SetDirty(GetParent(host), *_state);
                }
            });
            edit->SetOnBlur([this, host = parent]() noexcept
            {
                if (! _state)
                {
                    return;
                }
                Refresh(host, *_state);
            });
        }
    }

    dxState->page.contentWorkers.card  = root->AddChild<CardPanel>();
    dxState->page.contentWorkers.title = root->AddChild<Label>();
    dxState->page.contentWorkers.title->SetFontRole(FontRole::Body);
    dxState->page.contentWorkers.description = root->AddChild<Label>();
    dxState->page.contentWorkers.description->SetFontRole(FontRole::Small);
    dxState->page.contentWorkers.description->SetMultiline(true);
    dxState->page.contentWorkers.combo = root->AddChild<ComboBox>();
    dxState->page.contentWorkers.combo->SetVariant(ComboBoxVariant::Window);

    {
        std::vector<ComboBox::Item> items;
        items.push_back({L"0", LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_WORKERS_AUTO)});
        for (uint32_t i = 1; i <= 4; ++i)
        {
            auto label = std::to_wstring(i);
            items.push_back({label, label});
        }
        dxState->page.contentWorkers.combo->SetItems(std::move(items));
    }

    dxState->page.contentWorkers.combo->SetOnSelectionChanged([this, host = parent](size_t itemIndex) noexcept
    {
        if (_syncingDxContentWorkersCombo || ! _state)
        {
            return;
        }

        const uint32_t newValue = std::clamp(static_cast<uint32_t>(itemIndex), 0u, 4u);

        auto* compare = EnsureWorkingCompareDirectoriesSettings(_state->workingSettings);
        if (! compare)
        {
            return;
        }

        if (compare->contentCompareWorkerCount != newValue)
        {
            compare->contentCompareWorkerCount = newValue;
            MaybeResetWorkingCompareDirectoriesSettingsIfEmpty(_state->workingSettings);
            SetDirty(GetParent(host), *_state);
            Refresh(host, *_state);
        }
    });

    {
        CompareDxPage& page = dxState->page;
        std::vector<RedSalamander::DxUi::Control*> orderedChildren;
        orderedChildren.reserve(_pageContentRoot->DebugChildCount());

        const auto appendHeader = [&](const size_t index) noexcept { orderedChildren.push_back(page.headers[index]); };

        const auto appendToggleCard = [&](const size_t index, const bool includeEdit) noexcept
        {
            orderedChildren.push_back(page.toggleCards[index].card);
            orderedChildren.push_back(page.toggleCards[index].title);
            orderedChildren.push_back(page.toggleCards[index].description);
            orderedChildren.push_back(page.toggleCards[index].toggle);
            if (includeEdit)
            {
                orderedChildren.push_back(page.toggleCards[index].edit);
            }
        };

        const auto appendComboCard = [&]() noexcept
        {
            orderedChildren.push_back(page.contentWorkers.card);
            orderedChildren.push_back(page.contentWorkers.title);
            orderedChildren.push_back(page.contentWorkers.description);
            orderedChildren.push_back(page.contentWorkers.combo);
        };

        appendHeader(0u);
        appendHeader(1u);
        appendToggleCard(0u, false);

        appendHeader(2u);
        appendToggleCard(1u, false);
        appendToggleCard(2u, false);
        appendToggleCard(3u, false);
        appendToggleCard(4u, false);
        appendComboCard();

        appendHeader(3u);
        appendToggleCard(5u, false);
        appendToggleCard(6u, false);
        appendToggleCard(7u, false);
        appendToggleCard(8u, false);

        appendHeader(4u);
        appendToggleCard(9u, true);
        appendToggleCard(10u, true);

        ReorderPanelChildren(_pageContentRoot, orderedChildren);
    }

    _dxState = std::move(dxState);
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    return true;
}

void CompareDirectoriesPane::DetachDxHosts() noexcept
{
    _deferredFocusAfterLayout = DeferredFocusTarget::None;
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

void CompareDirectoriesPane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return;
    }

    _pageHostDx->SetTheme(MakeAppThemeDxPalette(state.theme));
}

void CompareDirectoriesPane::SyncDxControlsFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState)
    {
        return;
    }

    const auto& compare = GetCompareDirectoriesSettingsOrDefault(state.workingSettings);
    CompareDxPage& page = _dxState->page;

    for (size_t i = 0; i < kCompareHeaderStringIds.size(); ++i)
    {
        Label* header = page.headers[i];
        if (! header)
        {
            continue;
        }

        header->SetText(LoadStringResource(nullptr, kCompareHeaderStringIds[i]));
        header->SetEnabled(true);
    }

    for (size_t i = 0; i < page.toggleCards.size(); ++i)
    {
        CompareToggleCardPageDx& card = page.toggleCards[i];
        if (card.title)
        {
            card.title->SetText(LoadStringResource(nullptr, kCompareToggleLabelStringIds[i]));
        }
        if (card.description)
        {
            card.description->SetText(LoadStringResource(nullptr, kCompareToggleDescriptionStringIds[i]));
        }
        if (card.toggle)
        {
            bool checked = false;
            bool enabled = true;
            switch (kCompareToggleCommandIds[i])
            {
                case IDC_PREFS_ADV_COMPARE_SUBDIRS_TOGGLE: checked = compare.compareSubdirectories; break;
                case IDC_PREFS_ADV_COMPARE_SIZE_TOGGLE: checked = compare.compareSize; break;
                case IDC_PREFS_ADV_COMPARE_DATETIME_TOGGLE: checked = compare.compareDateTime; break;
                case IDC_PREFS_ADV_COMPARE_ATTRIBUTES_TOGGLE: checked = compare.compareAttributes; break;
                case IDC_PREFS_ADV_COMPARE_CONTENT_TOGGLE: checked = compare.compareContent; break;
                case IDC_PREFS_ADV_COMPARE_SUBDIR_ATTRIBUTES_TOGGLE: checked = compare.compareSubdirectoryAttributes; break;
                case IDC_PREFS_ADV_COMPARE_SELECT_SUBDIRS_ONE_PANE_TOGGLE: checked = compare.selectSubdirsOnlyInOnePane; break;
                case IDC_PREFS_ADV_COMPARE_KEEP_IDENTICAL_TOGGLE: checked = compare.keepIdenticalItems; break;
                case IDC_PREFS_ADV_COMPARE_SHOW_IDENTICAL_TOGGLE:
                    checked = compare.keepIdenticalItems && compare.showIdenticalItems;
                    enabled = compare.keepIdenticalItems;
                    break;
                case IDC_PREFS_ADV_COMPARE_IGNORE_FILES_TOGGLE: checked = compare.ignoreFiles; break;
                case IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_TOGGLE: checked = compare.ignoreDirectories; break;
                default: break;
            }
            card.toggle->SetChecked(checked);
            card.toggle->SetEnabled(enabled);
        }

        if (card.edit)
        {
            bool* syncFlag = nullptr;
            std::wstring text;

            if (i == kCompareIgnoreFilesIndex)
            {
                syncFlag = &_syncingDxIgnoreFilesEdit;
                text     = compare.ignoreFilesPatterns;
            }
            else if (i == kCompareIgnoreDirectoriesIndex)
            {
                syncFlag = &_syncingDxIgnoreDirectoriesEdit;
                text     = compare.ignoreDirectoriesPatterns;
            }

            if (syncFlag)
            {
                *syncFlag = true;
                card.edit->SetText(std::move(text));
                const bool enabled = (i == kCompareIgnoreFilesIndex) ? compare.ignoreFiles : compare.ignoreDirectories;
                card.edit->SetEnabled(enabled);
                *syncFlag = false;
            }
        }
    }

    if (page.contentWorkers.title)
    {
        page.contentWorkers.title->SetText(LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_WORKERS_TITLE));
        page.contentWorkers.title->SetEnabled(true);
    }
    if (page.contentWorkers.description)
    {
        page.contentWorkers.description->SetText(LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_WORKERS_DESC));
        page.contentWorkers.description->SetEnabled(true);
    }
    if (page.contentWorkers.combo)
    {
        _syncingDxContentWorkersCombo = true;
        const size_t selectedIndex    = std::min(static_cast<size_t>(compare.contentCompareWorkerCount), size_t{4});
        page.contentWorkers.combo->SetSelectedIndex(selectedIndex);
        page.contentWorkers.combo->SetEnabled(true);
        _syncingDxContentWorkersCombo = false;
    }

    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void CompareDirectoriesPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost = parent;
    _state    = &state;

    if (state.currentCategory == PrefCategory::CompareDirectories)
    {
        if (! EnsureDxHosts(parent, state))
        {
            Debug::Error(L"Preferences.CompareDirectories: DxUi surface initialization failed; page will not render correctly.");
            DetachDxHosts();
        }
    }

    Refresh(parent, state);
}

void CompareDirectoriesPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    if (state.currentCategory == PrefCategory::CompareDirectories && ! _dxState)
    {
        const HWND parent = _pageHost;
        if (parent)
        {
            static_cast<void>(EnsureDxHosts(parent, state));
        }
    }

    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    if (host)
    {
        InvalidateRect(host, nullptr, FALSE);
    }
}

void CompareDirectoriesPane::LayoutDxPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(host);
    static_cast<void>(margin);

    Debug::Perf::Scope layoutPerf(L"preferences.ui.compare_directories_layout_us");
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
    const int headerHeight = UiMetrics::ScaleDip(dpi, kHeaderHeightDip);

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
                                                                                   static_cast<float>(std::max(0, width - 2 * cardPaddingX)))));

    if (! _dxState || ! _pageHostDx || ! _pageContentRoot)
    {
        return;
    }

    const auto& compare   = GetCompareDirectoriesSettingsOrDefault(state.workingSettings);
    CompareDxPage& dxPage = _dxState->page;
    const auto pxToDip    = [dpi](const int px) noexcept { return static_cast<float>(px) * 96.0f / static_cast<float>(std::max<UINT>(1u, dpi)); };

    for (Label* header : dxPage.headers)
    {
        if (header)
        {
            header->SetVisible(false);
        }
    }
    for (CompareToggleCardPageDx& card : dxPage.toggleCards)
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
        if (card.edit)
        {
            card.edit->SetVisible(false);
        }
    }
    if (dxPage.contentWorkers.card)
    {
        dxPage.contentWorkers.card->SetVisible(false);
    }
    if (dxPage.contentWorkers.title)
    {
        dxPage.contentWorkers.title->SetVisible(false);
    }
    if (dxPage.contentWorkers.description)
    {
        dxPage.contentWorkers.description->SetVisible(false);
    }
    if (dxPage.contentWorkers.combo)
    {
        dxPage.contentWorkers.combo->SetVisible(false);
    }

    auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

    auto layoutHeader = [&](Label* header) noexcept
    {
        if (! header)
        {
            return;
        }

        header->SetVisible(true);
        header->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + headerHeight)));
        y += headerHeight + std::max(2, gapY / 2);
    };

    auto layoutToggleCard = [&](std::wstring_view labelText, std::wstring_view descText, const size_t index) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);
        const int cardHeight = std::max(rowHeight + 2 * cardPaddingY, titleHeight + cardGapY + descHeight + 2 * cardPaddingY);
        const RECT card{x, y, x + width, y + cardHeight};
        pushCard(card);

        if (index < dxPage.toggleCards.size())
        {
            CompareToggleCardPageDx& dxCard = dxPage.toggleCards[index];
            if (dxCard.card)
            {
                dxCard.card->SetVisible(true);
                dxCard.card->SetBounds(D2D1::RectF(pxToDip(card.left), pxToDip(card.top), pxToDip(card.right), pxToDip(card.bottom)));
            }
            if (dxCard.title)
            {
                dxCard.title->SetVisible(true);
                dxCard.title->SetText(std::wstring(labelText));
                dxCard.title->SetMnemonicTarget(dxCard.toggle);
                dxCard.title->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                    pxToDip(card.top + cardPaddingY),
                                                    pxToDip(card.left + cardPaddingX + textWidth),
                                                    pxToDip(card.top + cardPaddingY + titleHeight)));
            }
            if (dxCard.description)
            {
                dxCard.description->SetVisible(true);
                dxCard.description->SetText(std::wstring(descText));
                dxCard.description->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                          pxToDip(card.top + cardPaddingY + titleHeight + cardGapY),
                                                          pxToDip(card.left + cardPaddingX + textWidth),
                                                          pxToDip(card.top + cardPaddingY + titleHeight + cardGapY + descHeight)));
            }
            if (dxCard.toggle)
            {
                dxCard.toggle->SetVisible(true);
                dxCard.toggle->SetBounds(D2D1::RectF(pxToDip(card.right - cardPaddingX - toggleWidth),
                                                     pxToDip(card.top + (cardHeight - rowHeight) / 2),
                                                     pxToDip(card.right - cardPaddingX),
                                                     pxToDip(card.top + (cardHeight - rowHeight) / 2 + rowHeight)));
            }
            if (dxCard.edit)
            {
                dxCard.edit->SetVisible(false);
            }
        }

        y += cardHeight + cardSpacingY;
    };

    auto layoutComboCard = [&](std::wstring_view labelText, std::wstring_view descText) noexcept
    {
        const int desiredWidth = static_cast<int>(
            std::lround(RedSalamander::DxUi::ResolveConstrainedExtent({.minExtent       = static_cast<float>(UiMetrics::ScaleDip(dpi, kMinEditWidthDip)),
                                                                       .preferredExtent = static_cast<float>(UiMetrics::ScaleDip(dpi, kMinEditWidthDip)),
                                                                       .maxExtent       = static_cast<float>(UiMetrics::ScaleDip(dpi, kMaxEditWidthDip))},
                                                                      static_cast<float>(std::max(0, width - 2 * cardPaddingX)))));
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - desiredWidth);
        const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);
        const int cardHeight = std::max(rowHeight + 2 * cardPaddingY, titleHeight + cardGapY + descHeight + 2 * cardPaddingY);
        const RECT card{x, y, x + width, y + cardHeight};
        pushCard(card);

        CompareComboCardPageDx& dxCard = dxPage.contentWorkers;
        if (dxCard.card)
        {
            dxCard.card->SetVisible(true);
            dxCard.card->SetBounds(D2D1::RectF(pxToDip(card.left), pxToDip(card.top), pxToDip(card.right), pxToDip(card.bottom)));
        }
        if (dxCard.title)
        {
            dxCard.title->SetVisible(true);
            dxCard.title->SetText(std::wstring(labelText));
            dxCard.title->SetMnemonicTarget(dxCard.combo);
            dxCard.title->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                pxToDip(card.top + cardPaddingY),
                                                pxToDip(card.left + cardPaddingX + textWidth),
                                                pxToDip(card.top + cardPaddingY + titleHeight)));
        }
        if (dxCard.description)
        {
            dxCard.description->SetVisible(true);
            dxCard.description->SetText(std::wstring(descText));
            dxCard.description->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                      pxToDip(card.top + cardPaddingY + titleHeight + cardGapY),
                                                      pxToDip(card.left + cardPaddingX + textWidth),
                                                      pxToDip(card.top + cardPaddingY + titleHeight + cardGapY + descHeight)));
        }
        if (dxCard.combo)
        {
            dxCard.combo->SetVisible(true);
            dxCard.combo->SetBounds(D2D1::RectF(pxToDip(card.right - cardPaddingX - desiredWidth),
                                                pxToDip(card.top + (cardHeight - rowHeight) / 2),
                                                pxToDip(card.right - cardPaddingX),
                                                pxToDip(card.top + (cardHeight - rowHeight) / 2 + rowHeight)));
        }

        y += cardHeight + cardSpacingY;
    };

    auto layoutIgnoreCard = [&](std::wstring_view labelText, std::wstring_view descText, const bool showEdit, const size_t index) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, descText);
        int contentHeight    = titleHeight + cardGapY + descHeight;
        if (showEdit)
        {
            contentHeight += cardGapY + rowHeight;
        }
        const int cardHeight = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);
        const RECT card{x, y, x + width, y + cardHeight};
        pushCard(card);

        if (index < dxPage.toggleCards.size())
        {
            CompareToggleCardPageDx& dxCard = dxPage.toggleCards[index];
            const int editX                 = card.left + cardPaddingX;
            const int editW                 = std::max(0, width - 2 * cardPaddingX);
            const int editTop               = card.top + cardPaddingY + titleHeight + cardGapY + descHeight + cardGapY;
            if (dxCard.card)
            {
                dxCard.card->SetVisible(true);
                dxCard.card->SetBounds(D2D1::RectF(pxToDip(card.left), pxToDip(card.top), pxToDip(card.right), pxToDip(card.bottom)));
            }
            if (dxCard.title)
            {
                dxCard.title->SetVisible(true);
                dxCard.title->SetText(std::wstring(labelText));
                dxCard.title->SetMnemonicTarget(showEdit ? static_cast<RedSalamander::DxUi::Control*>(dxCard.edit)
                                                         : static_cast<RedSalamander::DxUi::Control*>(dxCard.toggle));
                dxCard.title->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                    pxToDip(card.top + cardPaddingY),
                                                    pxToDip(card.left + cardPaddingX + textWidth),
                                                    pxToDip(card.top + cardPaddingY + titleHeight)));
            }
            if (dxCard.description)
            {
                dxCard.description->SetVisible(true);
                dxCard.description->SetText(std::wstring(descText));
                dxCard.description->SetBounds(D2D1::RectF(pxToDip(card.left + cardPaddingX),
                                                          pxToDip(card.top + cardPaddingY + titleHeight + cardGapY),
                                                          pxToDip(card.left + cardPaddingX + textWidth),
                                                          pxToDip(card.top + cardPaddingY + titleHeight + cardGapY + descHeight)));
            }
            if (dxCard.toggle)
            {
                dxCard.toggle->SetVisible(true);
                dxCard.toggle->SetBounds(D2D1::RectF(pxToDip(card.right - cardPaddingX - toggleWidth),
                                                     pxToDip(card.top + cardPaddingY),
                                                     pxToDip(card.right - cardPaddingX),
                                                     pxToDip(card.top + cardPaddingY + rowHeight)));
            }
            if (dxCard.edit)
            {
                dxCard.edit->SetVisible(showEdit);
                if (showEdit)
                {
                    dxCard.edit->SetBounds(D2D1::RectF(pxToDip(editX), pxToDip(editTop), pxToDip(editX + editW), pxToDip(editTop + rowHeight)));
                }
            }
        }

        y += cardHeight + cardSpacingY;
    };

    layoutHeader(dxPage.headers[1]);
    layoutToggleCard(LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIRS_TITLE), LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIRS_DESC), 0u);
    y += gapY;

    layoutHeader(dxPage.headers[2]);
    layoutToggleCard(LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SIZE_TITLE), LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SIZE_DESC), 1u);
    layoutToggleCard(LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_DATETIME_TITLE), LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_DATETIME_DESC), 2u);
    layoutToggleCard(LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_ATTRIBUTES_TITLE), LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC), 3u);
    layoutToggleCard(LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_TITLE), LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_DESC), 4u);
    layoutComboCard(LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_WORKERS_TITLE),
                    LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_WORKERS_DESC));
    y += gapY;

    layoutHeader(dxPage.headers[3]);
    layoutToggleCard(
        LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_TITLE), LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC), 5u);
    layoutToggleCard(
        LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_TITLE), LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC), 6u);
    layoutToggleCard(
        LoadStringResource(nullptr, IDS_PREFS_COMPARE_KEEP_IDENTICAL_TITLE), LoadStringResource(nullptr, IDS_PREFS_COMPARE_KEEP_IDENTICAL_DESC), 7u);
    layoutToggleCard(
        LoadStringResource(nullptr, IDS_PREFS_COMPARE_SHOW_IDENTICAL_TITLE), LoadStringResource(nullptr, IDS_PREFS_COMPARE_SHOW_IDENTICAL_DESC), 8u);
    y += gapY;

    layoutHeader(dxPage.headers[4]);
    layoutIgnoreCard(LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC),
                     compare.ignoreFiles,
                     kCompareIgnoreFilesIndex);
    layoutIgnoreCard(LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_TITLE),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC),
                     compare.ignoreDirectories,
                     kCompareIgnoreDirectoriesIndex);

    SyncDxControlsFromState(state);
    _pageHostDx->Invalidate();
}

void CompareDirectoriesPane::LayoutPage(
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

    Debug::Error(L"Preferences.CompareDirectories: DxUi surface initialization failed; page will not render correctly.");
}

bool CompareDirectoriesPane::HandleDeferredAction(HWND host, PreferencesDialogState& /*state*/, PreferencesDeferredActionKind action) noexcept
{
    if (action != PreferencesDeferredActionKind::CompareDirectoriesIgnoreToggleChanged)
    {
        return false;
    }

    if (! host)
    {
        return false;
    }

    if (_deferredFocusAfterLayout == DeferredFocusTarget::None && _pageHostDx && _dxState)
    {
        const RedSalamander::DxUi::Control* const focused = _pageHostDx->GetFocusControl();
        if (focused == _dxState->page.toggleCards[kCompareIgnoreFilesIndex].toggle)
        {
            _deferredFocusAfterLayout = DeferredFocusTarget::IgnoreFilesToggle;
        }
        else if (focused == _dxState->page.toggleCards[kCompareIgnoreDirectoriesIndex].toggle)
        {
            _deferredFocusAfterLayout = DeferredFocusTarget::IgnoreDirectoriesToggle;
        }
    }

    // Relayout is performed by the caller (HandleDeferredPaneAction) via
    // LayoutPreferencesPageHost so scroll, background cards, and DxUi are all
    // updated through the proper channel.
    return true;
}

void CompareDirectoriesPane::RestoreDeferredFocusAfterLayout() noexcept
{
    const DeferredFocusTarget target = _deferredFocusAfterLayout;
    _deferredFocusAfterLayout        = DeferredFocusTarget::None;
    RestoreDeferredFocusTarget(target);
}

void CompareDirectoriesPane::RestoreDeferredFocusTarget(const DeferredFocusTarget target) noexcept
{
    if (target == DeferredFocusTarget::None || ! _pageHostDx || ! _dxState)
    {
        return;
    }

    RedSalamander::DxUi::Control* focusControl = nullptr;
    switch (target)
    {
        case DeferredFocusTarget::IgnoreFilesToggle: focusControl = _dxState->page.toggleCards[kCompareIgnoreFilesIndex].toggle; break;
        case DeferredFocusTarget::IgnoreDirectoriesToggle: focusControl = _dxState->page.toggleCards[kCompareIgnoreDirectoriesIndex].toggle; break;
        case DeferredFocusTarget::None: return;
    }

    if (focusControl && focusControl->IsVisible() && focusControl->IsEnabled())
    {
        _pageHostDx->SetFocusControl(focusControl);
    }
}

#ifdef ENABLE_TESTS
PreferencesCompareDirectoriesDebugFocusTarget CompareDirectoriesPane::DebugGetFocusTarget() const noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return PreferencesCompareDirectoriesDebugFocusTarget::None;
    }

    const auto* focused = _pageHostDx->GetFocusControl();
    if (! focused)
    {
        return PreferencesCompareDirectoriesDebugFocusTarget::None;
    }

    const auto& page = _dxState->page;
    if (page.toggleCards[0].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle;
    if (page.toggleCards[1].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::CompareSizeToggle;
    if (page.toggleCards[2].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::CompareDateTimeToggle;
    if (page.toggleCards[3].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::CompareAttributesToggle;
    if (page.toggleCards[4].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::CompareContentToggle;
    if (page.contentWorkers.combo == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::ContentWorkersCombo;
    if (page.toggleCards[5].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirAttributesToggle;
    if (page.toggleCards[6].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::SelectSubdirsOnlyInOnePaneToggle;
    if (page.toggleCards[7].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle;
    if (page.toggleCards[8].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::ShowIdenticalItemsToggle;
    if (page.toggleCards[9].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesToggle;
    if (page.toggleCards[9].edit == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesEdit;
    if (page.toggleCards[10].toggle == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesToggle;
    if (page.toggleCards[10].edit == focused)
        return PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesEdit;
    return PreferencesCompareDirectoriesDebugFocusTarget::None;
}

bool CompareDirectoriesPane::DebugFocusCompareSubdirectoriesToggle() noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return false;
    }

    auto* const toggle = _dxState->page.toggleCards[0].toggle;
    if (! toggle || ! toggle->IsVisible() || ! toggle->IsEnabled())
    {
        return false;
    }

    _pageHostDx->SetFocusControl(toggle);
    return _pageHostDx->GetFocusControl() == toggle;
}

bool CompareDirectoriesPane::DebugFocusTarget(const PreferencesCompareDirectoriesDebugFocusTarget target) noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return false;
    }

    RedSalamander::DxUi::Control* focusControl = nullptr;
    auto& page                                 = _dxState->page;
    switch (target)
    {
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle: focusControl = page.toggleCards[0].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareSizeToggle: focusControl = page.toggleCards[1].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareDateTimeToggle: focusControl = page.toggleCards[2].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareAttributesToggle: focusControl = page.toggleCards[3].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareContentToggle: focusControl = page.toggleCards[4].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::ContentWorkersCombo: focusControl = page.contentWorkers.combo; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirAttributesToggle: focusControl = page.toggleCards[5].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::SelectSubdirsOnlyInOnePaneToggle: focusControl = page.toggleCards[6].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle: focusControl = page.toggleCards[7].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::ShowIdenticalItemsToggle: focusControl = page.toggleCards[8].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesToggle: focusControl = page.toggleCards[9].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesEdit: focusControl = page.toggleCards[9].edit; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesToggle: focusControl = page.toggleCards[10].toggle; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesEdit: focusControl = page.toggleCards[10].edit; break;
        case PreferencesCompareDirectoriesDebugFocusTarget::None: return false;
    }

    if (! focusControl || ! focusControl->IsVisible() || ! focusControl->IsEnabled())
    {
        return false;
    }

    _pageHostDx->SetFocusControl(focusControl);
    return _pageHostDx->GetFocusControl() == focusControl;
}

bool CompareDirectoriesPane::DebugGetToggleChecked(const PreferencesCompareDirectoriesDebugFocusTarget target, bool& outChecked) const noexcept
{
    if (! _state)
    {
        return false;
    }

    const auto compare = _state->workingSettings.compareDirectories.value_or(Common::Settings::CompareDirectoriesSettings{});
    switch (target)
    {
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirectoriesToggle: outChecked = compare.compareSubdirectories; return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareSizeToggle: outChecked = compare.compareSize; return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareDateTimeToggle: outChecked = compare.compareDateTime; return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareAttributesToggle: outChecked = compare.compareAttributes; return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareContentToggle: outChecked = compare.compareContent; return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::CompareSubdirAttributesToggle: outChecked = compare.compareSubdirectoryAttributes; return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::SelectSubdirsOnlyInOnePaneToggle: outChecked = compare.selectSubdirsOnlyInOnePane; return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::KeepIdenticalItemsToggle: outChecked = compare.keepIdenticalItems; return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::ShowIdenticalItemsToggle:
            outChecked = compare.keepIdenticalItems && compare.showIdenticalItems;
            return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesToggle: outChecked = compare.ignoreFiles; return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesToggle: outChecked = compare.ignoreDirectories; return true;
        case PreferencesCompareDirectoriesDebugFocusTarget::ContentWorkersCombo:
        case PreferencesCompareDirectoriesDebugFocusTarget::IgnoreFilesEdit:
        case PreferencesCompareDirectoriesDebugFocusTarget::IgnoreDirectoriesEdit:
        case PreferencesCompareDirectoriesDebugFocusTarget::None: return false;
    }

    return false;
}

bool CompareDirectoriesPane::DebugSelectContentWorkersByText(std::wstring_view displayText) noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return false;
    }

    auto* const combo = _dxState->page.contentWorkers.combo;
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
    _pageHostDx->Invalidate();
    return combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == itemIndex && combo->GetDisplayedText() == displayText;
}

size_t CompareDirectoriesPane::DebugVisibleSectionHeaderCount() const noexcept
{
    if (! _dxState)
    {
        return 0u;
    }

    size_t visible = 0u;
    for (size_t i = 1u; i < _dxState->page.headers.size(); ++i)
    {
        const Label* header = _dxState->page.headers[i];
        if (header && header->IsVisible())
        {
            ++visible;
        }
    }

    return visible;
}
#endif
