#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "RedConfigureRoot.h"

#include "Helpers.h"
#include "RedConfigureApp.h"
#include "RedConfigureGridModels.h"
#include "RedConfigureLocalizationExampleControl.h"
#include "RedConfigurePagePresenters.h"
#include "RedConfigureSession.h"
#include "RedConfigureThemeExampleControl.h"
#include "RedConfigureUiHelpers.h"
#include "RedConfigureWorkflow.h"
#include "SettingsStore.h"
#include "resource.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <filesystem>
#include <format>
#include <functional>
#include <iterator>
#include <limits>
#include <memory>
#include <optional>
#include <set>
#include <span>
#include <string>
#include <string_view>
#include <system_error>
#include <vector>

namespace RedConfigure::Ui
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::ButtonVariant;
using RedSalamander::DxUi::ColorSwatch;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::Control;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Grid;
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridColumnKind;
using RedSalamander::DxUi::GridRowStyle;
using RedSalamander::DxUi::GridRowTone;
using RedSalamander::DxUi::GridSelectionMode;
using RedSalamander::DxUi::GridSortSpec;
using RedSalamander::DxUi::GridVisualMode;
using RedSalamander::DxUi::IDxGridDelegate;
using RedSalamander::DxUi::IDxGridModel;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::MakeDefaultThemePalette;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::ScrollPanel;
using RedSalamander::DxUi::Slider;
using RedSalamander::DxUi::SortDirection;
using RedSalamander::DxUi::TagPicker;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::WindowHost;

constexpr float kThemePreviewContentHeightDip                      = 560.0f;
constexpr float kLocalizationGridRowHeightDip                      = 52.0f;
constexpr float kLocalizationEditorMinFieldHeightDip               = 32.0f;
constexpr float kLocalizationEditorAdditionalLineHeightDip         = 28.0f;
constexpr float kLocalizationPageMinContentHeightDip               = 640.0f;
constexpr std::array<std::wstring_view, 4> kPageIconGlyphs         = {{L"\xE80F", L"\xE774", L"\xE790", L"\xE74E"}};
constexpr std::array<std::wstring_view, 26> kThemePreviewColorKeys = {{
    L"app.accent",
    L"window.background",
    L"navigation.background",
    L"navigation.text",
    L"navigation.accent",
    L"navigation.progressOk",
    L"navigation.progressBackground",
    L"menu.background",
    L"menu.text",
    L"menu.disabledText",
    L"menu.selectionBg",
    L"menu.selectionText",
    L"menu.border",
    L"folderView.background",
    L"folderView.textNormal",
    L"folderView.itemBackgroundHovered",
    L"folderView.itemBackgroundSelected",
    L"folderView.textSelected",
    L"folderView.warningBackground",
    L"folderView.warningText",
    L"fileOps.progressBackground",
    L"fileOps.progressTotal",
    L"viewer.diff.addedBackground",
    L"viewer.diff.removedBackground",
    L"monitor.textView.bg",
    L"monitor.textView.fg",
}};

[[nodiscard]] size_t ExplicitTextLineCount(std::wstring_view text) noexcept
{
    size_t lineCount = 1u;
    for (size_t index = 0u; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch == L'\n' || (ch == L'\r' && (index + 1u >= text.size() || text[index + 1u] != L'\n')))
        {
            ++lineCount;
        }
    }
    return lineCount;
}

[[nodiscard]] size_t EditorTextLineCount(std::wstring_view text, float fieldWidthDip) noexcept
{
    if (fieldWidthDip <= 0.0f)
    {
        return ExplicitTextLineCount(text);
    }

    const float usableWidthDip = std::max(1.0f, fieldWidthDip - 14.0f);
    const size_t charsPerLine  = std::max<size_t>(1u, static_cast<size_t>(std::floor(usableWidthDip / 10.0f)));
    size_t visualLineCount     = 0u;
    size_t segmentLength       = 0u;

    auto finishSegment = [&]()
    {
        visualLineCount += std::max<size_t>(1u, (segmentLength + charsPerLine - 1u) / charsPerLine);
        segmentLength = 0u;
    };

    for (size_t index = 0u; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch == L'\r')
        {
            finishSegment();
            if (index + 1u < text.size() && text[index + 1u] == L'\n')
            {
                ++index;
            }
        }
        else if (ch == L'\n')
        {
            finishSegment();
        }
        else
        {
            ++segmentLength;
        }
    }
    finishSegment();
    return std::max<size_t>(1u, visualLineCount);
}

[[nodiscard]] float EditorFieldHeightForText(std::wstring_view text, float fieldWidthDip = 0.0f) noexcept
{
    const size_t lineCount = EditorTextLineCount(text, fieldWidthDip);
    return kLocalizationEditorMinFieldHeightDip + (static_cast<float>(lineCount - 1u) * kLocalizationEditorAdditionalLineHeightDip);
}

[[nodiscard]] float TargetEditorLabelWidth(float widthDip) noexcept
{
    return std::clamp(widthDip * 0.18f, 58.0f, 86.0f);
}

[[nodiscard]] std::wstring ThemeKeyGroup(std::wstring_view key)
{
    const size_t dot = key.find(L'.');
    return std::wstring(dot == std::wstring_view::npos ? key : key.substr(0u, dot));
}

BOOL CALLBACK CollectLocaleNameProc(LPWSTR localeName, DWORD, LPARAM userData) noexcept
{
    auto* names = reinterpret_cast<std::set<std::wstring>*>(userData);
    if (names && localeName && localeName[0] != L'\0')
    {
        names->insert(localeName);
    }
    return TRUE;
}

[[nodiscard]] std::vector<std::wstring> EnumerateOfficialCultureNames()
{
    std::set<std::wstring> names;
    static_cast<void>(::EnumSystemLocalesEx(CollectLocaleNameProc, LOCALE_ALL, reinterpret_cast<LPARAM>(&names), nullptr));
    if (names.empty())
    {
        names.insert(L"en-US");
        names.insert(L"fr-FR");
        names.insert(L"de-DE");
        names.insert(L"es-ES");
        names.insert(L"it-IT");
    }
    return std::vector<std::wstring>(names.begin(), names.end());
}

[[nodiscard]] std::wstring LocaleDisplayName(std::wstring_view culture)
{
    const std::wstring cultureText(culture);
    wchar_t buffer[256]{};
    int length = ::GetLocaleInfoEx(cultureText.c_str(), LOCALE_SENGLISHDISPLAYNAME, buffer, static_cast<int>(std::size(buffer)));
    if (length <= 1)
    {
        length = ::GetLocaleInfoEx(cultureText.c_str(), LOCALE_SLOCALIZEDDISPLAYNAME, buffer, static_cast<int>(std::size(buffer)));
    }
    if (length <= 1)
    {
        return cultureText;
    }
    return std::wstring(buffer, static_cast<size_t>(length - 1));
}

[[nodiscard]] std::wstring CultureDisplayText(std::wstring_view culture, std::wstring_view suffix)
{
    const std::wstring cultureText(culture);
    const std::wstring displayName = LocaleDisplayName(culture);
    if (displayName.empty() || displayName == cultureText)
    {
        return cultureText + L" (" + std::wstring(suffix) + L")";
    }
    return cultureText + L" - " + displayName + L" (" + std::wstring(suffix) + L")";
}

[[nodiscard]] bool IsEnUsCulture(std::wstring_view culture) noexcept
{
    return ::CompareStringOrdinal(culture.data(), static_cast<int>(culture.size()), L"en-US", 5, TRUE) == CSTR_EQUAL;
}

[[nodiscard]] std::set<std::wstring> DiscoverExistingCultures(const RedConfigure::Workspace::WorkspaceScanResult& workspace)
{
    std::set<std::wstring> cultures;
    for (const auto& owner : workspace.resourceOwners)
    {
        for (const std::filesystem::path& path : owner.satelliteResourcePaths)
        {
            const std::filesystem::path cultureFolder = path.parent_path();
            if (! cultureFolder.empty() && cultureFolder.parent_path().filename() == L"Lang")
            {
                cultures.insert(cultureFolder.filename().wstring());
            }
        }
    }
    return cultures;
}

[[nodiscard]] std::wstring ValidationCategoryText(HINSTANCE instance, RedConfigure::Workflow::ValidationCategory category)
{
    return LoadAppString(instance, static_cast<UINT>(ReviewExportPagePresenter::GetCategoryResourceId(category)));
}

[[nodiscard]] std::wstring ValidationMessageText(HINSTANCE instance, const RedConfigure::Workflow::ValidationIssue& issue)
{
    const UINT resourceId = static_cast<UINT>(ReviewExportPagePresenter::GetMessageResourceId(issue.code));
    if (issue.code == RedConfigure::Workflow::ValidationCode::DuplicateAccelerator)
    {
        return FormatStringResource(instance, resourceId, issue.arguments.empty() ? std::wstring{} : issue.arguments.front());
    }
    return LoadAppString(instance, resourceId);
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text)
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

class RedConfigureRoot final : public Panel, public RedConfigureRootController, public IDxGridDelegate
{
public:
    RedConfigureRoot(HINSTANCE instance, RedConfigure::RedConfigureSession& session, std::filesystem::path initialRoot)
        : _instance(instance),
          _session(session),
          _workspaceRoot(std::move(initialRoot)),
          _inventoryModel(instance, session),
          _localizationReviewModel(instance, session),
          _themeColorModel(instance, session)
    {
        BuildControls();
        SetPage(0u);
    }

    RedConfigureRoot(const RedConfigureRoot&)            = delete;
    RedConfigureRoot& operator=(const RedConfigureRoot&) = delete;
    RedConfigureRoot(RedConfigureRoot&&)                 = delete;
    RedConfigureRoot& operator=(RedConfigureRoot&&)      = delete;

    void ReloadWorkspaceFromFields()
    {
        InvalidatePendingApprovals();
        _workspaceRoot       = RedConfigure::ResolveWorkspaceRootForLaunchPath(std::filesystem::path(std::wstring(_workspaceRootField->GetText())));
        std::wstring culture = ! _localizationReviewViewOptions.visibleCultureNames.empty()
                                   ? _localizationReviewViewOptions.visibleCultureNames.front()
                                   : (_cultureCombo ? std::wstring(_cultureCombo->GetSelectedValue()) : std::wstring(_session.GetCultureName()));
        if (culture.empty())
        {
            culture = L"fr-FR";
        }
        const std::wstring previousOwner(_session.GetActiveResourceOwnerName());
        const HRESULT hr = _session.LoadWorkspace(_workspaceRoot, culture);
        _localizationReviewViewOptions.visibleOwnerNames.clear();
        _localizationReviewViewOptions.visibleCultureNames.clear();
        _localizationReviewFiltersInitialized = false;
        _selectedReviewCulture.clear();
        if (SUCCEEDED(hr) && ! previousOwner.empty())
        {
            const auto& owners = _session.GetWorkspace().resourceOwners;
            for (size_t index = 0u; index < owners.size(); ++index)
            {
                if (owners[index].name == previousOwner)
                {
                    static_cast<void>(_session.SetActiveResourceOwner(index));
                    break;
                }
            }
        }
        _lastLoadSucceeded = SUCCEEDED(hr);
        _statusLabel->SetText(
            LoadAppString(_instance, _lastLoadSucceeded ? IDS_REDCONFIGURE_STATUS_WORKSPACE_READY : IDS_REDCONFIGURE_STATUS_WORKSPACE_FAILED));
        SyncFromSession();
    }

    void SelectPageForTest(size_t pageIndex) override
    {
        SetPage(pageIndex);
    }

    void OnGridSelectionChanged(Grid& sender) override
    {
        if (&sender == _themeColorGrid)
        {
            if (_syncing)
            {
                return;
            }

            const std::optional<size_t> row = sender.GetPrimarySelectedRow();
            if (! row)
            {
                return;
            }

            if (const std::optional<std::wstring> key = _themeColorModel.GetKeyAt(row.value()))
            {
                SelectThemeColorKey(key.value());
            }
            return;
        }

        if (&sender != _translationGrid)
        {
            return;
        }

        const std::optional<size_t> row = sender.GetPrimarySelectedRow();
        if (row)
        {
            const std::optional<size_t> sessionRow = _localizationReviewModel.ResolveSessionRow(row.value());
            if (! sessionRow)
            {
                return;
            }
            _selectedReviewRow = sessionRow.value();
            SyncTranslationEditor();
        }
    }

    void OnGridSelectionChanged() override
    {
    }

    void OnGridRowActivated(Grid& sender, size_t rowIndex) override
    {
        if (&sender != _translationGrid)
        {
            return;
        }
        const std::optional<size_t> sessionRow = _localizationReviewModel.ResolveSessionRow(rowIndex);
        if (! sessionRow.has_value())
        {
            return;
        }
        _selectedReviewRow = sessionRow.value();
        SyncTranslationEditor();
        if (WindowHost* host = GetHost(); host && ! _targetEditors.empty() && _targetEditors.front().field)
        {
            host->SetFocusControl(_targetEditors.front().field);
        }
    }

    void OnGridRowActivated(size_t) override
    {
    }

    void OnGridSortRequested(const GridSortSpec& sortSpec) override
    {
        _localizationReviewViewOptions.sortColumn      = ReviewColumnFromGridSort(sortSpec.columnIndex, _localizationReviewModel.GetColumnCount());
        _localizationReviewViewOptions.sortCultureName = ReviewSortCultureFromGridSort(sortSpec.columnIndex);
        _localizationReviewViewOptions.sortDirection   = SortDirectionFromGrid(sortSpec.direction);
        RebuildTranslationView(true);
    }

protected:
    void OnBoundsChanged() noexcept override
    {
        LayoutControls();
    }

private:
    struct TargetEditor
    {
        std::wstring cultureName;
        Label* label     = nullptr;
        TextField* field = nullptr;
    };

    [[nodiscard]] static RedConfigure::LocalizationViewColumn ReviewColumnFromGridSort(size_t columnIndex, size_t columnCount) noexcept
    {
        if (columnIndex == 0u)
        {
            return RedConfigure::LocalizationViewColumn::Owner;
        }
        if (columnIndex == 1u)
        {
            return RedConfigure::LocalizationViewColumn::Id;
        }
        if (columnIndex == 2u)
        {
            return RedConfigure::LocalizationViewColumn::Source;
        }
        if (columnCount > 0u && columnIndex + 1u == columnCount)
        {
            return RedConfigure::LocalizationViewColumn::Status;
        }
        return RedConfigure::LocalizationViewColumn::Target;
    }

    [[nodiscard]] std::wstring ReviewSortCultureFromGridSort(size_t columnIndex) const
    {
        return _localizationReviewModel.ResolveTargetCulture(columnIndex).value_or(std::wstring{});
    }

    [[nodiscard]] static RedConfigure::LocalizationSortDirection SortDirectionFromGrid(SortDirection direction) noexcept
    {
        switch (direction)
        {
            case SortDirection::Ascending: return RedConfigure::LocalizationSortDirection::Ascending;
            case SortDirection::Descending: return RedConfigure::LocalizationSortDirection::Descending;
            case SortDirection::None: return RedConfigure::LocalizationSortDirection::None;
            default: return RedConfigure::LocalizationSortDirection::None;
        }
    }

    [[nodiscard]] static bool ContainsText(std::span<const std::wstring> values, std::wstring_view value) noexcept
    {
        for (const std::wstring& candidate : values)
        {
            if (candidate == value)
            {
                return true;
            }
        }
        return false;
    }

    [[nodiscard]] std::vector<std::wstring> BuildUniqueOwnerNames() const
    {
        std::vector<std::wstring> owners;
        owners.reserve(_session.GetWorkspace().resourceOwners.size());
        for (const auto& owner : _session.GetWorkspace().resourceOwners)
        {
            if (! ContainsText(owners, owner.name))
            {
                owners.push_back(owner.name);
            }
        }
        return owners;
    }

    [[nodiscard]] std::vector<std::wstring> BuildReviewCultureNames() const
    {
        const auto cultures = _session.GetLocalizationReviewCultures();
        return std::vector<std::wstring>(cultures.begin(), cultures.end());
    }

    void BuildControls()
    {
        const auto pages = RedConfigure::GetPageDefinitions();
        _navButtons.reserve(pages.size());
        for (size_t index = 0u; index < pages.size(); ++index)
        {
            auto* button = AddChild<Button>(LoadAppString(_instance, pages[index].titleResourceId));
            button->SetOnClick([this, index] { SetPage(index); });
            _navButtons.push_back(button);
        }

        _titleLabel = AddChild<Label>();
        _titleLabel->SetFontRole(FontRole::Subtitle);
        _descriptionLabel = AddChild<Label>();
        _descriptionLabel->SetMultiline(true);
        _descriptionLabel->SetFontRole(FontRole::Small);
        _descriptionLabel->SetVisible(false);
        _scopeLabel = AddChild<Label>();
        _scopeLabel->SetFontRole(FontRole::Small);
        _scopeLabel->SetMultiline(false);
        _statusLabel = AddChild<Label>();
        _statusLabel->SetFontRole(FontRole::Small);
        _statusLabel->SetMultiline(true);

        _commandSearchField = AddChild<TextField>();
        _commandSearchField->SetPlaceholder(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_COMMAND_SEARCH));
        _commandSearchField->SetOnTextChanged([this](std::wstring_view text)
        {
            if (_syncing) return;
            if (_selectedPage == 1u && _localizationSearchField)
            {
                _localizationSearchField->SetText(std::wstring(text));
            }
            else if (_selectedPage == 2u && _colorKeyFilterField)
            {
                _colorKeyFilterField->SetText(std::wstring(text));
            }
        });
        _validateButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_VALIDATE));
        _validateButton->SetOnClick([this] { RunValidation(true); });
        _undoButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_UNDO));
        _undoButton->SetOnClick([this]
        {
            if (_session.Undo()) SyncFromSession();
        });
        _redoButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_REDO));
        _redoButton->SetOnClick([this]
        {
            if (_session.Redo()) SyncFromSession();
        });
        _reviewExportButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_REVIEW_EXPORT));
        _reviewExportButton->SetPrimary(true);
        _reviewExportButton->SetOnClick([this] { SetPage(3u); });
        _validationDrawerButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_VALIDATION_DRAWER));
        _validationDrawerButton->SetOnClick([this]
        {
            _validationDrawerExpanded = ! _validationDrawerExpanded;
            RunValidation(false);
            SyncVisibility();
            LayoutControls();
        });
        _validationDrawerLabel = AddChild<Label>();
        _validationDrawerLabel->SetMultiline(true);
        _validationDrawerLabel->SetFontRole(FontRole::Small);

        _workspaceRootLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_WORKSPACE_ROOT));
        _workspaceRootField = AddChild<TextField>(_workspaceRoot.wstring());
        _cultureLabel       = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_TARGET_LANGUAGE));
        _cultureCombo       = AddChild<ComboBox>();
        _cultureCombo->SetOnSelectionChanged([this](size_t)
        {
            if (_syncing)
            {
                return;
            }
            AddVisibleReviewCulture(std::wstring(_cultureCombo->GetSelectedValue()));
        });
        _reloadButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_RELOAD));
        _reloadButton->SetPrimary(true);
        _reloadButton->SetOnClick([this] { ReloadWorkspaceFromFields(); });
        _ownerCountLabel     = AddChild<Label>();
        _themeFileCountLabel = AddChild<Label>();
        _scanErrorCountLabel = AddChild<Label>();
        _ownerSelectorLabel  = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_ACTIVE_OWNER));
        _ownerCombo          = AddChild<ComboBox>();
        _ownerCombo->SetOnSelectionChanged([this](size_t selectedIndex)
        {
            if (_syncing)
            {
                return;
            }
            if (SUCCEEDED(_session.SetActiveResourceOwner(selectedIndex)))
            {
                _selectedReviewRow = 0u;
                SyncFromSession();
            }
        });
        _localizationPageScroll = AddChild<ScrollPanel>();
        _localizationPageScroll->SetScrollStepDip(64.0f);
        _localizationPageContent = _localizationPageScroll->AddChild<Panel>();

        _ownerFilterLabel  = _localizationPageContent->AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_OWNERS));
        _ownerFilterPicker = _localizationPageContent->AddChild<TagPicker>();
        _ownerFilterPicker->SetOnSelectionChanged([this](std::span<const std::wstring> values)
        {
            if (_syncing)
            {
                return;
            }

            _localizationReviewViewOptions.visibleOwnerNames.assign(values.begin(), values.end());
            _localizationReviewFiltersInitialized = true;
            RebuildTranslationView(true);
            SyncScopeLabel();
            if (auto* host = GetHost())
            {
                host->Invalidate();
            }
        });
        _languageFilterLabel  = _localizationPageContent->AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_LANGUAGES));
        _languageFilterPicker = _localizationPageContent->AddChild<TagPicker>();
        _languageFilterPicker->SetOnSelectionChanged([this](std::span<const std::wstring> values)
        {
            if (_syncing)
            {
                return;
            }

            _localizationReviewViewOptions.visibleCultureNames.clear();
            _localizationReviewViewOptions.visibleCultureNames.reserve(values.size());
            for (const std::wstring& culture : values)
            {
                if (_session.EnsureLocalizationReviewCulture(culture))
                {
                    _localizationReviewViewOptions.visibleCultureNames.push_back(culture);
                }
            }
            _languageColumns.Set(_localizationReviewViewOptions.visibleCultureNames);
            _localizationReviewFiltersInitialized = true;
            if (std::find(_localizationReviewViewOptions.visibleCultureNames.begin(),
                          _localizationReviewViewOptions.visibleCultureNames.end(),
                          _selectedReviewCulture) == _localizationReviewViewOptions.visibleCultureNames.end())
            {
                _selectedReviewCulture =
                    _localizationReviewViewOptions.visibleCultureNames.empty() ? std::wstring{} : _localizationReviewViewOptions.visibleCultureNames.front();
            }
            RebuildTranslationView(true);
            SyncExportPathLabels();
            SyncExportPreviews();
            SyncScopeLabel();
            if (auto* host = GetHost())
            {
                host->Invalidate();
            }
        });
        _activeOwnerLabel = _localizationPageContent->AddChild<Label>();
        _activeOwnerLabel->SetFontRole(FontRole::Small);
        _translationCountLabel   = _localizationPageContent->AddChild<Label>();
        _localizationSearchLabel = _localizationPageContent->AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SEARCH));
        _localizationSearchField = _localizationPageContent->AddChild<TextField>();
        _localizationSearchField->SetOnTextChanged([this](std::wstring_view text)
        {
            if (_syncing)
            {
                return;
            }
            _localizationReviewViewOptions.searchText.assign(text);
            RebuildTranslationView(true);
        });
        _localizationIdFilterLabel = _localizationPageContent->AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_ID_FILTER));
        _localizationIdFilterField = _localizationPageContent->AddChild<TextField>();
        _localizationIdFilterField->SetOnTextChanged([this](std::wstring_view text)
        {
            if (_syncing)
            {
                return;
            }
            _localizationReviewViewOptions.idFilterText.assign(text);
            RebuildTranslationView(true);
        });
        _localizationStatusFilterLabel = _localizationPageContent->AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_COL_STATUS));
        _localizationStatusFilterCombo = _localizationPageContent->AddChild<ComboBox>();
        std::vector<ComboBox::Item> statusFilterItems;
        statusFilterItems.push_back(ComboBox::Item{.value = L"all", .display = LoadAppString(_instance, IDS_REDCONFIGURE_FILTER_ALL)});
        statusFilterItems.push_back(ComboBox::Item{.value = L"ok", .display = LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_OK)});
        statusFilterItems.push_back(ComboBox::Item{.value = L"problems", .display = LoadAppString(_instance, IDS_REDCONFIGURE_FILTER_PROBLEMS)});
        _localizationStatusFilterCombo->SetItems(std::move(statusFilterItems));
        _localizationStatusFilterCombo->SetSelectedIndex(0u);
        _localizationStatusFilterCombo->SetOnSelectionChanged([this](size_t selectedIndex)
        {
            if (_syncing)
            {
                return;
            }
            _localizationReviewViewOptions.statusFilter = selectedIndex == 1u   ? RedConfigure::LocalizationStatusFilter::Ok
                                                          : selectedIndex == 2u ? RedConfigure::LocalizationStatusFilter::Problems
                                                                                : RedConfigure::LocalizationStatusFilter::All;
            RebuildTranslationView(true);
        });

        _inventoryGrid = AddChild<Grid>();
        _inventoryGrid->SetModel(&_inventoryModel);
        _inventoryGrid->SetSelectionMode(GridSelectionMode::Single);
        _inventoryGrid->SetRowHeightDip(30.0f);
        _inventoryGrid->SetHeaderHeightDip(32.0f);
        _inventoryGrid->SetLineClamp(1u);
        _inventoryGrid->SetEmptyStateText(LoadAppString(_instance, IDS_REDCONFIGURE_LOCALIZATION_EMPTY));

        _translationGrid = _localizationPageContent->AddChild<Grid>();
        _translationGrid->SetModel(&_localizationReviewModel);
        _translationGrid->SetDelegate(this);
        _translationGrid->SetSelectionMode(GridSelectionMode::Single);
        _translationGrid->SetRowHeightDip(kLocalizationGridRowHeightDip);
        _translationGrid->SetHeaderHeightDip(32.0f);
        _translationGrid->SetLineClamp(2u);
        _translationGrid->SetEmptyStateText(LoadAppString(_instance, IDS_REDCONFIGURE_LOCALIZATION_REVIEW_EMPTY));

        _sourceTextLabel = _localizationPageContent->AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_ENGLISH_SOURCE));
        _sourceTextField = _localizationPageContent->AddChild<TextField>();
        _sourceTextField->SetReadOnly(true);
        _sourceTextField->SetMultiline(true);
        _targetTextLabel    = _localizationPageContent->AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_TARGET_TEXT));
        _targetEditorsPanel = _localizationPageContent->AddChild<ScrollPanel>();
        _targetEditorsPanel->SetScrollStepDip(kLocalizationGridRowHeightDip);
        _validationLabel = _localizationPageContent->AddChild<Label>();
        _validationLabel->SetFontRole(FontRole::BodyStrong);
        _previousProblemButton = _localizationPageContent->AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_PREVIOUS_PROBLEM));
        _previousProblemButton->SetOnClick([this] { SelectAdjacentProblem(false); });
        _nextProblemButton = _localizationPageContent->AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_NEXT_PROBLEM));
        _nextProblemButton->SetOnClick([this] { SelectAdjacentProblem(true); });
        _pasteMatrixButton = _localizationPageContent->AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_PASTE_MATRIX));
        _pasteMatrixButton->SetOnClick([this] { PasteLocalizationMatrix(); });
        _pinLanguageButton = _localizationPageContent->AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_PIN_LANGUAGE));
        _pinLanguageButton->SetOnClick([this] { PinSelectedLanguage(); });
        _moveLanguageLeftButton = _localizationPageContent->AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_MOVE_LEFT));
        _moveLanguageLeftButton->SetOnClick([this] { MoveSelectedLanguage(false); });
        _moveLanguageRightButton = _localizationPageContent->AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_MOVE_RIGHT));
        _moveLanguageRightButton->SetOnClick([this] { MoveSelectedLanguage(true); });
        _removeLanguageButton = _localizationPageContent->AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_REMOVE_LANGUAGE));
        _removeLanguageButton->SetOnClick([this] { RemoveSelectedLanguage(); });
        _localizationBatchCombo = _localizationPageContent->AddChild<ComboBox>();
        _localizationBatchCombo->SetItems({
            {.value = L"copyEnglish", .display = LoadAppString(_instance, IDS_REDCONFIGURE_BATCH_COPY_ENGLISH)},
            {.value = L"copyCulture", .display = LoadAppString(_instance, IDS_REDCONFIGURE_BATCH_COPY_CULTURE)},
            {.value = L"clear", .display = LoadAppString(_instance, IDS_REDCONFIGURE_BATCH_CLEAR)},
            {.value = L"findReplace", .display = LoadAppString(_instance, IDS_REDCONFIGURE_BATCH_FIND_REPLACE)},
            {.value = L"normalize", .display = LoadAppString(_instance, IDS_REDCONFIGURE_BATCH_NORMALIZE)},
            {.value = L"accelerators", .display = LoadAppString(_instance, IDS_REDCONFIGURE_BATCH_ACCELERATORS)},
            {.value = L"reviewed", .display = LoadAppString(_instance, IDS_REDCONFIGURE_BATCH_MARK_REVIEWED)},
        });
        _localizationBatchCombo->SetSelectedIndex(0u);
        _applyLocalizationBatchButton = _localizationPageContent->AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_APPLY_BATCH));
        _applyLocalizationBatchButton->SetOnClick([this] { PreviewOrApplyLocalizationBatch(); });
        _localizationExampleLabel = _localizationPageContent->AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_LIVE_EXAMPLE));
        _localizationExample = _localizationPageContent->AddChild<LocalizationExampleControl>();
        _exportRcButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_EXPORT_CHANGED_RC));
        _exportRcButton->SetOnClick([this] { ExportLocalization(); });

        _themeSelectorLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_NAME));
        _themeCombo         = AddChild<ComboBox>();
        _themeCombo->SetOnSelectionChanged([this](size_t selectedIndex)
        {
            if (_syncing)
            {
                return;
            }
            if (_session.SetActiveTheme(selectedIndex))
            {
                SyncThemeLibraryLabels();
                SyncThemeColorKeyCombo({});
                SyncThemeColorEditor();
                SyncExportPathLabels();
                SyncExportPreviews();
                SyncScopeLabel();
                if (auto* host = GetHost())
                {
                    host->Invalidate();
                }
            }
        });
        _themeNameLabel  = AddChild<Label>();
        _themePathLabel  = AddChild<Label>();
        _themeErrorLabel = AddChild<Label>();
        _themeImportPathField = AddChild<TextField>();
        _themeImportPathField->SetPlaceholder(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_PATH));
        _themeImportButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_IMPORT_THEME));
        _themeImportButton->SetOnClick([this] { ImportThemeFromField(); });
        _themeDuplicateButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_DUPLICATE_THEME));
        _themeDuplicateButton->SetOnClick([this] { DuplicateActiveTheme(); });
        _themeResetButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_RESET_THEME));
        _themeResetButton->SetOnClick([this]
        {
            if (_session.ResetActiveTheme()) SyncFromSession();
        });

        _colorKeyFilterLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_KEY_FILTER));
        _colorKeyFilterField = AddChild<TextField>();
        _colorKeyFilterField->SetOnTextChanged([this](std::wstring_view)
        {
            if (_syncing)
            {
                return;
            }

            SyncThemeColorKeyCombo(std::wstring(_colorKeyCombo->GetSelectedValue()));
            SyncThemeColorEditor();
            if (auto* host = GetHost())
            {
                host->Invalidate();
            }
        });
        _colorKeyLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_COLOR_KEY));
        _colorKeyCombo = AddChild<ComboBox>();
        _colorKeyCombo->SetVisible(false);
        _themeColorGrid = AddChild<Grid>();
        _themeColorGrid->SetModel(&_themeColorModel);
        _themeColorGrid->SetDelegate(this);
        _themeColorGrid->SetSelectionMode(GridSelectionMode::Single);
        _themeColorGrid->SetRowHeightDip(28.0f);
        _themeColorGrid->SetHeaderHeightDip(30.0f);
        _themeColorGrid->SetLineClamp(1u);
        _themeColorGrid->SetEmptyStateText(LoadAppString(_instance, IDS_REDCONFIGURE_THEME_KEYS_EMPTY));
        SyncThemeColorKeyCombo({});
        _colorKeyCombo->SetOnSelectionChanged([this](size_t) { SyncThemeColorEditor(); });
        _previousColorLabel = AddChild<Label>();
        _previousColorLabel->SetFontRole(FontRole::Small);
        _previousColorSwatch = AddChild<ColorSwatch>();
        _colorValueLabel     = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_COLOR_VALUE));
        _colorSwatch         = AddChild<ColorSwatch>();
        _colorValueCombo     = AddChild<ComboBox>();
        _colorValueCombo->SetEditable(true);
        _colorValueCombo->SetAutoOpenOnTextInput(true);
        _colorValueCombo->SetMaxVisibleItems(8u);
        _colorValueCombo->SetOnTextChanged([this](std::wstring_view text) { OnThemeColorTextChanged(text); });
        _themeColorStatusLabel = AddChild<Label>();
        _themeColorStatusLabel->SetFontRole(FontRole::Small);
        _themeColorStatusLabel->SetMultiline(true);
        _paletteNameLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_PALETTE_NAME));
        _paletteNameField = AddChild<TextField>();
        _addPaletteButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_ADD_PALETTE));
        _addPaletteButton->SetOnClick([this] { AddPaletteEntry(); });
        _renamePaletteButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_RENAME_PALETTE));
        _renamePaletteButton->SetOnClick([this] { RenameSelectedPaletteEntry(); });
        _previewSeedLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_PREVIEW_SEED));
        _previewSeedCombo = AddChild<ComboBox>();
        _previewSeedCombo->SetItems({{.value = L"0", .display = L"0"},
                                     {.value = L"1", .display = L"1"},
                                     {.value = L"42", .display = L"42"},
                                     {.value = L"1391283949", .display = L"0x52ED5EED"}});
        _previewSeedCombo->SetSelectedIndex(3u);
        _previewSeedCombo->SetOnSelectionChanged([this](size_t index)
        {
            constexpr std::array<uint32_t, 4> kPreviewSeeds{{0u, 1u, 42u, 0x52ED5EEDu}};
            if (index >= kPreviewSeeds.size()) return;
            _session.GetThemePreviewModel().SetPreviewSeed(kPreviewSeeds[index]);
            _themeColorGrid->NotifyDataChanged();
            SyncThemeColorEditor();
            _themePreview->Refresh();
        });
        _themeExpressionHelpLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_EXPRESSION_EXAMPLES));
        _themeExpressionHelpLabel->SetFontRole(FontRole::Small);
        _themeExpressionHelpLabel->SetMultiline(true);
        _themePreviewScroll = AddChild<ScrollPanel>();
        _themePreviewScroll->SetScrollStepDip(56.0f);
        _themePreview = _themePreviewScroll->AddChild<ThemeExampleControl>(_instance);
        _themePreview->SetModel(&_session.GetThemePreviewModel());
        _themePreview->SetOnTokenSelected([this](std::wstring_view key) { SelectThemeColorKey(key); });
        _themeSceneLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SCENE));
        _themeSceneCombo = AddChild<ComboBox>();
        _themeSceneCombo->SetItems({
            {.value = L"app", .display = LoadAppString(_instance, IDS_REDCONFIGURE_SCENE_APP_SHELL)},
            {.value = L"folderView", .display = LoadAppString(_instance, IDS_REDCONFIGURE_SCENE_FOLDER_VIEW)},
            {.value = L"menu", .display = LoadAppString(_instance, IDS_REDCONFIGURE_SCENE_MENU_POPUP)},
            {.value = L"window", .display = LoadAppString(_instance, IDS_REDCONFIGURE_SCENE_DIALOGS)},
            {.value = L"fileOps", .display = LoadAppString(_instance, IDS_REDCONFIGURE_SCENE_FILE_OPERATIONS)},
            {.value = L"monitor", .display = LoadAppString(_instance, IDS_REDCONFIGURE_SCENE_MONITOR_LOG)},
            {.value = L"viewer", .display = LoadAppString(_instance, IDS_REDCONFIGURE_SCENE_VIEWER_DIFF)},
        });
        _themeSceneCombo->SetSelectedIndex(0u);
        _themeSceneCombo->SetOnSelectionChanged([this](size_t)
        {
            if (_syncing) return;
            _colorKeyFilterField->SetText(std::wstring(_themeSceneCombo->GetSelectedValue()));
        });
        _themeRecipeLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_RECIPE));
        _themeRecipeCombo = AddChild<ComboBox>();
        _themeRecipeCombo->SetItems({
            {.value = L"dark", .display = LoadAppString(_instance, IDS_REDCONFIGURE_RECIPE_DARK)},
            {.value = L"light", .display = LoadAppString(_instance, IDS_REDCONFIGURE_RECIPE_LIGHT)},
            {.value = L"accent", .display = LoadAppString(_instance, IDS_REDCONFIGURE_RECIPE_ACCENT)},
            {.value = L"soft", .display = LoadAppString(_instance, IDS_REDCONFIGURE_RECIPE_SOFT_SELECTION)},
            {.value = L"contrast", .display = LoadAppString(_instance, IDS_REDCONFIGURE_RECIPE_CONTRAST)},
            {.value = L"semantic", .display = LoadAppString(_instance, IDS_REDCONFIGURE_RECIPE_SEMANTIC)},
            {.value = L"alpha", .display = LoadAppString(_instance, IDS_REDCONFIGURE_RECIPE_SET_ALPHA)},
            {.value = L"replace", .display = LoadAppString(_instance, IDS_REDCONFIGURE_RECIPE_REPLACE_REFERENCE)},
            {.value = L"convert", .display = LoadAppString(_instance, IDS_REDCONFIGURE_RECIPE_CONVERT_REFERENCES)},
            {.value = L"remove", .display = LoadAppString(_instance, IDS_REDCONFIGURE_RECIPE_REMOVE_OVERRIDES)},
        });
        _themeRecipeCombo->SetSelectedIndex(0u);
        _applyThemeRecipeButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_APPLY_BATCH));
        _applyThemeRecipeButton->SetOnClick([this] { PreviewOrApplyThemeRecipe(); });
        _themeAlphaSlider = AddChild<Slider>();
        _themeAlphaSlider->SetMinimum(0.0);
        _themeAlphaSlider->SetMaximum(100.0);
        _themeAlphaSlider->SetValue(80.0);
        _themeAlphaSlider->SetStep(5.0);
        _copyEffectiveButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_COPY_EFFECTIVE));
        _copyEffectiveButton->SetOnClick([this] { CopySelectedThemeColor(false); });
        _copyOverrideButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_COPY_OVERRIDE));
        _copyOverrideButton->SetOnClick([this] { CopySelectedThemeColor(true); });
        _darkenGroupButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_DARKEN_GROUP));
        _darkenGroupButton->SetOnClick([this] { ApplyThemeGroupTransform(ThemeGroupTransform::Darken); });
        _blendAccentGroupButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_BLEND_ACCENT_GROUP));
        _blendAccentGroupButton->SetOnClick([this] { ApplyThemeGroupTransform(ThemeGroupTransform::BlendAccent); });
        _resetColorButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_RESET_COLOR));
        _resetColorButton->SetOnClick([this] { ResetSelectedThemeColor(); });
        _exportThemeButton = AddChild<Button>(LoadAppString(_instance, IDS_REDCONFIGURE_BUTTON_EXPORT_THEME));
        _exportThemeButton->SetOnClick([this] { ExportTheme(); });

        _defaultRcPathLabel    = AddChild<Label>();
        _defaultThemePathLabel = AddChild<Label>();
        _reviewOutputScroll    = AddChild<ScrollPanel>();
    }

    void SetPage(size_t pageIndex)
    {
        const auto pages = RedConfigure::GetPageDefinitions();
        if (pageIndex >= pages.size())
        {
            return;
        }

        _selectedPage = pageIndex;
        if (_themePreview)
        {
            _themePreview->SetModel(pageIndex == 2u ? &_session.GetThemePreviewModel() : nullptr);
        }
        _titleLabel->SetText(LoadAppString(_instance, pages[pageIndex].titleResourceId));
        _descriptionLabel->SetText(LoadAppString(_instance, pages[pageIndex].descriptionResourceId));
        for (size_t index = 0u; index < _navButtons.size(); ++index)
        {
            _navButtons[index]->SetPrimary(index == pageIndex);
        }
        SyncScopeLabel();
        SyncVisibility();
        LayoutControls();
        if (auto* host = GetHost())
        {
            host->Invalidate();
        }
    }

    void SyncFromSession()
    {
        _syncing = true;
        const StartPageSummary startSummary = StartPagePresenter::Build(_session);
        _workspaceRootField->SetText(_workspaceRoot.wstring());
        _ownerCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_RESOURCE_OWNERS) + L": " +
                                  std::to_wstring(startSummary.resourceOwnerCount));
        _themeFileCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_FILES) + L": " +
                                      std::to_wstring(startSummary.themeFileCount));
        _scanErrorCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SCAN_ERRORS) + L": " +
                                      std::to_wstring(startSummary.scanErrorCount));
        _activeOwnerLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_ACTIVE_OWNER) + L": " + std::wstring(_session.GetActiveResourceOwnerName()));
        _translationCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_TRANSLATION_COUNT) + L": " +
                                        std::to_wstring(_session.GetLocalizationReviewRows().size()));
        SyncScopeLabel();

        SyncCultureCombo();
        SyncOwnerCombo();
        SyncLocalizationReviewFilterDefaults();
        SyncThemeCombo();
        SyncThemeLibraryLabels();
        SyncThemeColorKeyCombo({});
        SyncExportPathLabels();
        SyncExportPreviews();
        _inventoryGrid->NotifyDataChanged();
        RebuildTranslationView(true);
        SyncThemeColorEditor();
        _undoButton->SetEnabled(_session.CanUndo());
        _redoButton->SetEnabled(_session.CanRedo());
        RunValidation(false);
        _syncing = false;

        if (auto* host = GetHost())
        {
            host->Invalidate();
        }
    }

    void InvalidatePendingApprovals() noexcept
    {
        _localizationPresenter.Invalidate();
        _themesPresenter.Invalidate();
    }

    void RebuildTranslationView(bool preserveSelection)
    {
        InvalidatePendingApprovals();
        SyncLocalizationReviewFilterDefaults();
        _localizationReviewViewRows = RedConfigure::BuildLocalizationReviewView(_session.GetLocalizationReviewRows(), _localizationReviewViewOptions);
        _localizationReviewModel.SetVisibleCultures(_localizationReviewViewOptions.visibleCultureNames);
        _localizationReviewModel.SetViewRows(_localizationReviewViewRows);
        _translationGrid->NotifyDataChanged();
        const size_t totalTranslations = _session.GetLocalizationReviewRows().size();
        std::wstring countText =
            LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_TRANSLATION_COUNT) + L": " + std::to_wstring(_localizationReviewViewRows.size());
        if (_localizationReviewViewRows.size() != totalTranslations)
        {
            countText += L" / " + std::to_wstring(totalTranslations);
        }
        _translationCountLabel->SetText(std::move(countText));

        if (_localizationReviewViewRows.empty())
        {
            _selectedReviewRow = 0u;
            SyncTranslationEditor();
            return;
        }

        size_t viewRow = 0u;
        if (preserveSelection)
        {
            const auto it = std::find(_localizationReviewViewRows.begin(), _localizationReviewViewRows.end(), _selectedReviewRow);
            if (it != _localizationReviewViewRows.end())
            {
                viewRow = static_cast<size_t>(std::distance(_localizationReviewViewRows.begin(), it));
            }
            else
            {
                _selectedReviewRow = _localizationReviewViewRows.front();
            }
        }
        else
        {
            _selectedReviewRow = _localizationReviewViewRows.front();
        }

        static_cast<void>(_translationGrid->RequestSelectRow(viewRow, 0u));
        SyncTranslationEditor();
    }

    void AddVisibleReviewCulture(std::wstring cultureName)
    {
        if (cultureName.empty() || ! _session.EnsureLocalizationReviewCulture(cultureName))
        {
            return;
        }

        if (std::find(_localizationReviewViewOptions.visibleCultureNames.begin(), _localizationReviewViewOptions.visibleCultureNames.end(), cultureName) ==
            _localizationReviewViewOptions.visibleCultureNames.end())
        {
            _localizationReviewViewOptions.visibleCultureNames.push_back(cultureName);
        }

        _selectedReviewCulture                = std::move(cultureName);
        _localizationReviewFiltersInitialized = true;
        SyncLocalizationTagPickers();
        RebuildTranslationView(true);
        SyncExportPathLabels();
        SyncExportPreviews();
        SyncScopeLabel();
    }

    void SyncLocalizationReviewFilterDefaults()
    {
        if (! _localizationReviewFiltersInitialized)
        {
            _localizationReviewViewOptions.visibleOwnerNames = BuildUniqueOwnerNames();

            const std::wstring currentCulture(_session.GetCultureName());
            if (! currentCulture.empty())
            {
                static_cast<void>(_session.EnsureLocalizationReviewCulture(currentCulture));
            }

            const auto cultures = _session.GetLocalizationReviewCultures();
            _localizationReviewViewOptions.visibleCultureNames.clear();
            _localizationReviewViewOptions.visibleCultureNames.assign(cultures.begin(), cultures.end());
            _languageColumns.Set(_localizationReviewViewOptions.visibleCultureNames);
            _localizationReviewFiltersInitialized = true;
        }

        if (_selectedReviewCulture.empty() || std::find(_localizationReviewViewOptions.visibleCultureNames.begin(),
                                                        _localizationReviewViewOptions.visibleCultureNames.end(),
                                                        _selectedReviewCulture) == _localizationReviewViewOptions.visibleCultureNames.end())
        {
            _selectedReviewCulture =
                _localizationReviewViewOptions.visibleCultureNames.empty() ? std::wstring{} : _localizationReviewViewOptions.visibleCultureNames.front();
        }
        SyncLocalizationTagPickers();
    }

    void SyncLocalizationTagPickers()
    {
        if (_ownerFilterPicker && _languageFilterPicker)
        {
            _ownerFilterPicker->SetOptions(LoadAppString(_instance, IDS_REDCONFIGURE_TOGGLE_ALL_OWNERS), BuildUniqueOwnerNames());
            _ownerFilterPicker->SetSelectedValues(_localizationReviewViewOptions.visibleOwnerNames);

            _languageFilterPicker->SetOptions(LoadAppString(_instance, IDS_REDCONFIGURE_TOGGLE_ALL_LANGUAGES), BuildReviewCultureNames());
            _languageFilterPicker->SetSelectedValues(_localizationReviewViewOptions.visibleCultureNames);
        }
    }

    void SyncThemeLibraryLabels()
    {
        const auto& catalog   = _session.GetThemeCatalog();
        std::wstring nameText = LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_NAME) + L": ";
        std::wstring pathText = LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_PATH) + L": ";
        if (! catalog.themes.empty())
        {
            const size_t themeIndex = std::min(_session.GetActiveThemeIndex(), catalog.themes.size() - 1u);
            nameText += catalog.themes[themeIndex].definition.name;
            pathText += catalog.themes[themeIndex].path.wstring();
        }
        else
        {
            nameText += _session.GetThemePreviewModel().GetTheme().name;
        }

        _themeNameLabel->SetText(std::move(nameText));
        _themePathLabel->SetText(std::move(pathText));
        _themeErrorLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_ERRORS) + L": " + std::to_wstring(catalog.errors.size()));
    }

    void SyncCultureCombo()
    {
        std::wstring current(_session.GetCultureName());
        if (current.empty())
        {
            current = L"fr-FR";
        }

        const std::wstring existingSuffix = LoadAppString(_instance, IDS_REDCONFIGURE_CULTURE_EXISTING_SUFFIX);
        const std::wstring createSuffix   = LoadAppString(_instance, IDS_REDCONFIGURE_CULTURE_CREATE_SUFFIX);

        const std::set<std::wstring> existingCultures = DiscoverExistingCultures(_session.GetWorkspace());

        // en-US is the embedded source language; offering it would only produce a dead
        // option because EnsureLocalizationReviewCulture rejects it.
        std::vector<ComboBox::Item> cultureItems;
        cultureItems.reserve(existingCultures.size() + 128u);
        std::set<std::wstring> added;
        for (const std::wstring& culture : existingCultures)
        {
            if (IsEnUsCulture(culture))
            {
                continue;
            }
            cultureItems.push_back(ComboBox::Item{.value = culture, .display = CultureDisplayText(culture, existingSuffix)});
            added.insert(culture);
        }

        for (const std::wstring& culture : EnumerateOfficialCultureNames())
        {
            if (! IsEnUsCulture(culture) && added.insert(culture).second)
            {
                cultureItems.push_back(ComboBox::Item{.value = culture, .display = CultureDisplayText(culture, createSuffix)});
            }
        }
        if (! IsEnUsCulture(current) && added.insert(current).second)
        {
            cultureItems.push_back(ComboBox::Item{.value = current, .display = CultureDisplayText(current, createSuffix)});
        }

        size_t selectedIndex = 0u;
        for (size_t index = 0u; index < cultureItems.size(); ++index)
        {
            if (cultureItems[index].value == current)
            {
                selectedIndex = index;
                break;
            }
        }

        _cultureCombo->SetItems(std::move(cultureItems));
        _cultureCombo->SetSelectedIndex(selectedIndex);
    }

    void SyncOwnerCombo()
    {
        std::vector<ComboBox::Item> ownerItems;
        ownerItems.reserve(_session.GetWorkspace().resourceOwners.size());
        for (const auto& owner : _session.GetWorkspace().resourceOwners)
        {
            ownerItems.push_back(ComboBox::Item{.value = owner.name, .display = owner.name});
        }

        _ownerCombo->SetItems(std::move(ownerItems));
        if (_session.GetWorkspace().resourceOwners.empty())
        {
            _ownerCombo->SetSelectedIndex(std::nullopt);
        }
        else
        {
            _ownerCombo->SetSelectedIndex(std::min(_session.GetActiveResourceOwnerIndex(), _session.GetWorkspace().resourceOwners.size() - 1u));
        }
    }

    void SyncThemeCombo()
    {
        std::vector<ComboBox::Item> themeItems;
        const auto& catalog = _session.GetThemeCatalog();
        if (catalog.themes.empty())
        {
            const auto& fallback = _session.GetThemePreviewModel().GetTheme();
            themeItems.push_back(
                ComboBox::Item{.value = fallback.id, .display = LoadAppString(_instance, IDS_REDCONFIGURE_THEME_ORIGIN_BUILTIN) + fallback.name});
        }
        else
        {
            themeItems.reserve(catalog.themes.size());
            for (const auto& theme : catalog.themes)
            {
                const UINT originId = static_cast<UINT>(ThemesPagePresenter::GetOriginResourceId(theme.origin));
                const std::wstring section = LoadAppString(_instance, originId);
                themeItems.push_back(ComboBox::Item{.value = theme.definition.id, .display = section + theme.definition.name});
            }
        }

        _themeCombo->SetItems(std::move(themeItems));
        _themeCombo->SetSelectedIndex(catalog.themes.empty() ? std::optional<size_t>(0u)
                                                             : std::optional<size_t>(std::min(_session.GetActiveThemeIndex(), catalog.themes.size() - 1u)));
    }

    void SyncThemeColorKeyCombo(std::wstring_view preferredKey)
    {
        if (! _colorKeyCombo)
        {
            return;
        }

        std::wstring selectedKey(preferredKey);
        if (selectedKey.empty())
        {
            selectedKey = std::wstring(_colorKeyCombo->GetSelectedValue());
        }

        std::set<std::wstring> added;
        std::vector<std::wstring> keys;
        keys.reserve(kThemePreviewColorKeys.size() + _session.GetThemePreviewModel().GetTheme().palette.size() +
                     _session.GetThemePreviewModel().GetTheme().colors.size());
        const auto addKey = [&added, &keys](std::wstring key)
        {
            if (added.insert(key).second)
            {
                keys.push_back(std::move(key));
            }
        };

        for (std::wstring_view key : kThemePreviewColorKeys)
        {
            addKey(std::wstring(key));
        }

        std::vector<std::wstring> paletteKeys;
        paletteKeys.reserve(_session.GetThemePreviewModel().GetTheme().palette.size());
        for (const auto& [name, _] : _session.GetThemePreviewModel().GetTheme().palette)
        {
            paletteKeys.push_back(std::wstring(L"palette.") + name);
        }
        std::sort(paletteKeys.begin(), paletteKeys.end());
        for (std::wstring& key : paletteKeys)
        {
            addKey(std::move(key));
        }

        std::vector<std::wstring> authoredKeys;
        authoredKeys.reserve(_session.GetThemePreviewModel().GetTheme().colors.size());
        for (const auto& [key, _] : _session.GetThemePreviewModel().GetTheme().colors)
        {
            authoredKeys.push_back(key);
        }
        std::sort(authoredKeys.begin(), authoredKeys.end());
        for (std::wstring& key : authoredKeys)
        {
            addKey(std::move(key));
        }

        _themeColorKeys = std::move(keys);

        const std::wstring filterText = _colorKeyFilterField ? std::wstring(_colorKeyFilterField->GetText()) : std::wstring{};
        _filteredThemeColorKeys       = RedConfigure::FilterThemeColorKeys(_themeColorKeys, filterText);

        const bool previousSyncing = _syncing;
        _syncing                   = true;
        _themeColorModel.SetKeys(_filteredThemeColorKeys);
        if (_themeColorGrid)
        {
            _themeColorGrid->NotifyDataChanged();
        }

        std::vector<ComboBox::Item> colorItems;
        colorItems.reserve(_filteredThemeColorKeys.size());
        for (const std::wstring& key : _filteredThemeColorKeys)
        {
            colorItems.push_back(ComboBox::Item{.value = key, .display = key});
        }

        size_t selectedIndex = 0u;
        for (size_t index = 0u; index < _filteredThemeColorKeys.size(); ++index)
        {
            if (! selectedKey.empty() && _filteredThemeColorKeys[index] == selectedKey)
            {
                selectedIndex = index;
                break;
            }
        }

        _colorKeyCombo->SetItems(std::move(colorItems));
        _colorKeyCombo->SetSelectedIndex(_filteredThemeColorKeys.empty() ? std::optional<size_t>{} : std::optional<size_t>(selectedIndex));
        if (_themeColorGrid && ! _filteredThemeColorKeys.empty())
        {
            static_cast<void>(_themeColorGrid->RequestSelectRow(selectedIndex, 0u));
        }
        _syncing = previousSyncing;
    }

    void SyncExportPathLabels()
    {
        const std::filesystem::path localizationPath = _session.GetLocalizationReviewRows().empty() ? _session.GetDefaultLocalizationExportPath()
                                                                                                    : _session.GetDefaultLocalizationExportPath().parent_path();
        _defaultRcPathLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_DEFAULT_RC_PATH) + L": " + localizationPath.wstring());
        _defaultThemePathLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_DEFAULT_THEME_PATH) + L": " +
                                        _session.GetDefaultThemeExportPath().wstring());
    }

    void SyncExportPreviews()
    {
        InvalidatePendingApprovals();
        if (! _reviewOutputScroll)
        {
            return;
        }

        _reviewOutputCards.clear();
        _reviewOutputScroll->ClearChildren();

        const auto addCard = [this](const std::filesystem::path& path, std::wstring text) {
            auto* label = _reviewOutputScroll->AddChild<Label>(path.wstring());
            label->SetFontRole(FontRole::Small);
            auto* field = _reviewOutputScroll->AddChild<TextField>();
            field->SetReadOnly(true);
            field->SetMultiline(true);
            field->SetText(std::move(text));
            _reviewOutputCards.push_back(ReviewOutputCard{.label = label, .field = field});
        };

        std::vector<RedConfigure::LocalizationExportPreview> reviewPreviews;
        if (! _session.GetLocalizationReviewRows().empty() && SUCCEEDED(_session.BuildLocalizationReviewExportPreviews(reviewPreviews)))
        {
            for (const RedConfigure::LocalizationExportPreview& preview : reviewPreviews)
            {
                addCard(preview.path, preview.text);
            }
        }
        else
        {
            std::wstring rcPreview;
            if (SUCCEEDED(_session.BuildLocalizationExportText(rcPreview)))
            {
                addCard(_session.GetDefaultLocalizationExportPath(), std::move(rcPreview));
            }
        }

        std::string themePreview;
        if (SUCCEEDED(_session.BuildThemeExportText(themePreview)))
        {
            addCard(_session.GetDefaultThemeExportPath(), Utf16FromUtf8(themePreview));
        }

        LayoutExportPage(_reviewOutputScroll->GetBounds().left,
                         _defaultRcPathLabel->GetBounds().top,
                         _reviewOutputScroll->GetBounds().right,
                         _reviewOutputScroll->GetBounds().bottom);
    }

    [[nodiscard]] bool TargetEditorCulturesMatch() const noexcept
    {
        if (_targetEditors.size() != _localizationReviewViewOptions.visibleCultureNames.size())
        {
            return false;
        }

        for (size_t index = 0u; index < _targetEditors.size(); ++index)
        {
            if (_targetEditors[index].cultureName != _localizationReviewViewOptions.visibleCultureNames[index])
            {
                return false;
            }
        }
        return true;
    }

    void EnsureTargetEditorControls()
    {
        if (! _targetEditorsPanel || TargetEditorCulturesMatch())
        {
            return;
        }

        _targetEditors.clear();
        _targetEditorsPanel->ClearChildren();
        _targetEditors.reserve(_localizationReviewViewOptions.visibleCultureNames.size());
        for (const std::wstring& culture : _localizationReviewViewOptions.visibleCultureNames)
        {
            auto* label = _targetEditorsPanel->AddChild<Label>(culture);
            label->SetFontRole(FontRole::Small);
            auto* field = _targetEditorsPanel->AddChild<TextField>();
            field->SetMultiline(true);
            field->SetOnTextChanged([this, culture](std::wstring_view text) { OnTargetTextChanged(culture, text); });
            _targetEditors.push_back(TargetEditor{.cultureName = culture, .label = label, .field = field});
        }

        LayoutTargetEditorsPanel(_targetEditorsPanel->GetBounds());
    }

    void SyncTranslationEditor()
    {
        const auto rows = _session.GetLocalizationReviewRows();
        _syncing        = true;
        EnsureTargetEditorControls();
        _targetTextLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_TARGET_TEXT));
        RedConfigure::Localization::PlaceholderStatus summaryStatus = RedConfigure::Localization::PlaceholderStatus::Ok;
        if (_selectedReviewRow < rows.size())
        {
            const RedConfigure::LocalizationReviewRow& row = rows[_selectedReviewRow];
            _activeOwnerLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SELECTED_CELL) + L": " + row.ownerName + L" / " + row.id);
            if (_sourceTextField->GetText() != row.sourceText)
            {
                _sourceTextField->SetText(row.sourceText);
            }
            for (TargetEditor& editor : _targetEditors)
            {
                if (! editor.field)
                {
                    continue;
                }

                if (const RedConfigure::LocalizationTargetCell* cell = FindLocalizationReviewTargetCell(row, editor.cultureName))
                {
                    editor.field->SetReadOnly(false);
                    // Skip redundant SetText: per-keystroke rebuilds would otherwise reset
                    // the focused editor's scroll position and selection.
                    if (editor.field->GetText() != cell->targetText)
                    {
                        editor.field->SetText(cell->targetText);
                    }
                    if (summaryStatus == RedConfigure::Localization::PlaceholderStatus::Ok &&
                        cell->validation.status != RedConfigure::Localization::PlaceholderStatus::Ok)
                    {
                        summaryStatus = cell->validation.status;
                    }
                }
                else
                {
                    editor.field->SetReadOnly(true);
                    if (! editor.field->GetText().empty())
                    {
                        editor.field->SetText({});
                    }
                }
            }
            SetValidationStatus(summaryStatus);
            const auto sourceValidation = RedConfigure::Localization::ValidatePlaceholders(row.sourceText, row.sourceText);
            std::wstring placeholders;
            for (const std::wstring& placeholder : sourceValidation.sourcePlaceholders)
            {
                if (! placeholders.empty()) placeholders += L", ";
                placeholders += placeholder;
            }
            if (placeholders.empty()) placeholders = L"—";
            std::wstring accelerator = L"—";
            for (size_t index = 0u; index + 1u < row.sourceText.size(); ++index)
            {
                if (row.sourceText[index] == L'&' && row.sourceText[index + 1u] != L'&')
                {
                    accelerator.assign(1u, row.sourceText[index + 1u]);
                    break;
                }
            }
            _validationLabel->SetText(FormatStringResource(_instance,
                                                           IDS_REDCONFIGURE_FMT_LOCALIZATION_INSPECTOR,
                                                           PlaceholderStatusText(_instance, summaryStatus),
                                                           placeholders,
                                                           accelerator));
            std::wstring target;
            if (const RedConfigure::LocalizationTargetCell* cell = FindLocalizationReviewTargetCell(row, _selectedReviewCulture))
            {
                target = cell->targetText;
            }
            LocalizationExampleControl::Kind kind = row.sourceText.find(L"{0}") != std::wstring::npos
                                                        ? LocalizationExampleControl::Kind::FormattedString
                                                        : row.id.find(L"MENU") != std::wstring::npos ? LocalizationExampleControl::Kind::Menu
                                                        : row.id.find(L"CMD") != std::wstring::npos ? LocalizationExampleControl::Kind::CommandLabel
                                                                                                    : LocalizationExampleControl::Kind::String;
            _localizationExample->SetExample(kind, row.sourceText, target);
        }
        else
        {
            _activeOwnerLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SELECTED_CELL) + L":");
            _sourceTextField->SetText({});
            for (TargetEditor& editor : _targetEditors)
            {
                if (editor.field)
                {
                    editor.field->SetReadOnly(true);
                    editor.field->SetText({});
                }
            }
            SetValidationStatus(RedConfigure::Localization::PlaceholderStatus::Ok);
            _localizationExample->SetExample(LocalizationExampleControl::Kind::String, {}, {});
        }
        _syncing = false;
        LayoutControls();
        if (auto* host = GetHost())
        {
            host->Invalidate();
        }
    }

    void SyncThemeColorEditor()
    {
        if (! _colorKeyCombo)
        {
            return;
        }

        const std::wstring key = std::wstring(_colorKeyCombo->GetSelectedValue());
        if (key.empty())
        {
            _colorSwatch->SetSwatchValue(std::nullopt);
            _previousColorSwatch->SetSwatchValue(std::nullopt);
            _previousColorLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_PREVIOUS_COLOR) + L":");
            _colorValueCombo->SetItems({});
            _colorValueCombo->SetText({});
            _themeColorStatusLabel->SetText({});
            _themePreview->SetSelectedToken({});
            return;
        }

        const std::wstring previousKey      = (! _activeThemeColorKey.empty() && _activeThemeColorKey != key) ? _activeThemeColorKey : _previousThemeColorKey;
        _syncing                            = true;
        const std::optional<uint32_t> color = _session.GetThemePreviewModel().GetEffectiveColor(key);
        const std::wstring authoredText     = _session.GetThemePreviewModel().GetAuthoredColorText(key);
        std::vector<std::wstring> suggestions = RedConfigure::BuildThemeColorSuggestions(key, previousKey, color);
        for (const std::wstring& recent : _recentThemeColors)
        {
            if (std::find(suggestions.begin(), suggestions.end(), recent) == suggestions.end()) suggestions.push_back(recent);
        }
        std::vector<ComboBox::Item> suggestionItems;
        suggestionItems.reserve(suggestions.size());
        for (const std::wstring& suggestion : suggestions)
        {
            suggestionItems.push_back(ComboBox::Item{.value = suggestion, .display = suggestion});
        }

        if (! previousKey.empty())
        {
            _previousColorSwatch->SetSwatchValue(_session.GetThemePreviewModel().GetEffectiveColor(previousKey));
            _previousColorLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_PREVIOUS_COLOR) + L": " + previousKey);
        }
        else
        {
            _previousColorSwatch->SetSwatchValue(std::nullopt);
            _previousColorLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_PREVIOUS_COLOR) + L":");
        }

        _colorSwatch->SetSwatchValue(color);
        _colorValueCombo->SetItems(std::move(suggestionItems));
        _colorValueCombo->SetText(! authoredText.empty() ? authoredText : (color ? Common::Settings::FormatColor(color.value()) : std::wstring{}));
        const std::optional<std::wstring_view> paletteName = key.starts_with(L"palette.") ? std::optional<std::wstring_view>(std::wstring_view(key).substr(8u))
                                                                                          : std::nullopt;
        _paletteNameField->SetText(paletteName.has_value() ? std::wstring(paletteName.value()) : std::wstring{});
        _renamePaletteButton->SetEnabled(paletteName.has_value());
        SyncThemeSourceInspector(key);
        _themePreview->SetSelectedToken(key);
        _themePreview->Refresh();
        if (_activeThemeColorKey != key)
        {
            _previousThemeColorKey = _activeThemeColorKey;
            _activeThemeColorKey   = key;
        }
        _syncing = false;
    }

    void OnTargetTextChanged(std::wstring_view cultureName, std::wstring_view text)
    {
        if (_syncing)
        {
            return;
        }

        const auto rows = _session.GetLocalizationReviewRows();
        if (_selectedReviewRow >= rows.size() || cultureName.empty())
        {
            return;
        }

        const auto validation = RedConfigure::Localization::ValidatePlaceholders(rows[_selectedReviewRow].sourceText, text);
        SetValidationStatus(validation.status);
        if (_session.UpdateLocalizationReviewTarget(_selectedReviewRow, cultureName, text))
        {
            RebuildTranslationView(true);
            SyncExportPathLabels();
            SyncExportPreviews();
        }
    }

    void OnThemeColorTextChanged(std::wstring_view text)
    {
        if (_syncing)
        {
            return;
        }

        const std::wstring key = std::wstring(_colorKeyCombo->GetSelectedValue());
        if (key.empty())
        {
            return;
        }

        if (_session.UpdateThemeColor(key, text))
        {
            const std::wstring recent(text);
            _recentThemeColors.erase(std::remove(_recentThemeColors.begin(), _recentThemeColors.end(), recent), _recentThemeColors.end());
            _recentThemeColors.insert(_recentThemeColors.begin(), recent);
            if (_recentThemeColors.size() > 8u) _recentThemeColors.resize(8u);
            const std::optional<uint32_t> color = _session.GetThemePreviewModel().GetEffectiveColor(key);
            _colorSwatch->SetSwatchValue(color);
            SyncThemeSourceInspector(key);
            _themeColorStatusLabel->SetTextColor(ResolvePalette().subduedText);
            _themeColorGrid->NotifyDataChanged();
            _themePreview->Refresh();
            SyncExportPathLabels();
            SyncExportPreviews();
        }
        else
        {
            _themeColorStatusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_THEME_COLOR_INVALID));
            _themeColorStatusLabel->SetTextColor(ResolvePalette().errorText);
        }
    }

    enum class ThemeGroupTransform : uint8_t
    {
        Darken,
        BlendAccent,
    };

    void SyncThemeSourceInspector(std::wstring_view key)
    {
        const auto& model = _session.GetThemePreviewModel();
        std::wstring sourceLabel;
        const std::optional<Common::Settings::ThemeColorSourceKind> kind = model.GetSourceKind(key);
        if (! kind.has_value())
        {
            sourceLabel = model.GetTheme().baseThemeId == L"builtin/rainbow" && key == L"folderView.itemBackgroundSelected"
                              ? LoadAppString(_instance, IDS_REDCONFIGURE_SOURCE_INHERITED_RAINBOW)
                              : LoadAppString(_instance, IDS_REDCONFIGURE_SOURCE_BASE);
        }
        else if (kind.value() == Common::Settings::ThemeColorSourceKind::Direct)
        {
            sourceLabel = LoadAppString(_instance, IDS_REDCONFIGURE_SOURCE_LITERAL);
        }
        else if (kind.value() == Common::Settings::ThemeColorSourceKind::Reference)
        {
            const std::vector<std::wstring> dependencies = model.GetDependencies(key);
            sourceLabel = ! dependencies.empty() && dependencies.front().starts_with(L"palette.")
                              ? LoadAppString(_instance, IDS_REDCONFIGURE_SOURCE_PALETTE_REFERENCE)
                              : LoadAppString(_instance, IDS_REDCONFIGURE_SOURCE_TOKEN_REFERENCE);
        }
        else if (model.GetEvaluationPhase(key) == Common::Settings::ThemeColorEvaluationPhase::Paint)
        {
            sourceLabel = LoadAppString(_instance, IDS_REDCONFIGURE_SOURCE_DYNAMIC);
        }
        else
        {
            sourceLabel = LoadAppString(_instance, IDS_REDCONFIGURE_SOURCE_FUNCTION);
        }

        const auto phase = model.GetEvaluationPhase(key);
        const UINT phaseId = static_cast<UINT>(phase == Common::Settings::ThemeColorEvaluationPhase::Paint
                                                   ? IDS_REDCONFIGURE_PHASE_PAINT
                                                   : phase == Common::Settings::ThemeColorEvaluationPhase::Event ? IDS_REDCONFIGURE_PHASE_EVENT
                                                                                                                  : IDS_REDCONFIGURE_PHASE_LOAD);
        const auto join = [](const std::vector<std::wstring>& values)
        {
            if (values.empty()) return std::wstring(L"—");
            std::wstring text;
            for (const std::wstring& value : values)
            {
                if (! text.empty()) text += L", ";
                text += value;
            }
            return text;
        };
        _themeColorStatusLabel->SetText(FormatStringResource(
            _instance, IDS_REDCONFIGURE_FMT_THEME_SOURCE_INSPECTOR, sourceLabel, LoadAppString(_instance, phaseId), join(model.GetDependencies(key)), join(model.GetAffected(key))));
    }

    void AddPaletteEntry()
    {
        const std::wstring name(_paletteNameField->GetText());
        std::wstring source(_colorValueCombo->GetText());
        if (source.empty())
        {
            source = _colorSwatch->GetSwatchValue().has_value() ? Common::Settings::FormatColor(_colorSwatch->GetSwatchValue().value()) : L"#000000";
        }
        const std::wstring key = std::wstring(L"palette.") + name;
        if (! _session.GetThemePreviewModel().CreatePaletteEntry(name, source, true))
        {
            _themeColorStatusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_THEME_PALETTE_INVALID));
            _themeColorStatusLabel->SetTextColor(ResolvePalette().errorText);
            return;
        }
        SyncThemeColorKeyCombo(key);
        SyncThemeColorEditor();
        SyncExportPreviews();
        _themePreview->Refresh();
    }

    void RenameSelectedPaletteEntry()
    {
        const std::wstring selectedKey(_colorKeyCombo->GetSelectedValue());
        if (! selectedKey.starts_with(L"palette.")) return;
        const std::wstring oldName = selectedKey.substr(8u);
        const std::wstring newName(_paletteNameField->GetText());
        if (! _session.GetThemePreviewModel().RenamePaletteEntry(oldName, newName))
        {
            _themeColorStatusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_THEME_PALETTE_INVALID));
            _themeColorStatusLabel->SetTextColor(ResolvePalette().errorText);
            return;
        }
        const std::wstring newKey = std::wstring(L"palette.") + newName;
        SyncThemeColorKeyCombo(newKey);
        SyncThemeColorEditor();
        SyncExportPreviews();
        _themePreview->Refresh();
    }

    void SelectThemeColorKey(std::wstring_view key)
    {
        if (! _colorKeyCombo)
        {
            return;
        }

        for (size_t index = 0u; index < _filteredThemeColorKeys.size(); ++index)
        {
            if (_filteredThemeColorKeys[index] == key)
            {
                const bool previousSyncing = _syncing;
                _syncing                   = true;
                _colorKeyCombo->SetSelectedIndex(index);
                if (_themeColorGrid)
                {
                    static_cast<void>(_themeColorGrid->RequestSelectRow(index, 0u));
                }
                _syncing = previousSyncing;
                SyncThemeColorEditor();
                return;
            }
        }

        for (const std::wstring& availableKey : _themeColorKeys)
        {
            if (availableKey == key)
            {
                _syncing = true;
                _colorKeyFilterField->SetText({});
                _syncing = false;
                SyncThemeColorKeyCombo(key);
                SyncThemeColorEditor();
                return;
            }
        }
    }

    void ApplyThemeGroupTransform(ThemeGroupTransform transform)
    {
        const std::wstring selectedKey(_colorKeyCombo->GetSelectedValue());
        const std::wstring selectedGroup = ThemeKeyGroup(selectedKey);
        if (selectedGroup.empty())
        {
            return;
        }

        bool changed = false;
        for (const std::wstring& key : _themeColorKeys)
        {
            if (ThemeKeyGroup(key) != selectedGroup)
            {
                continue;
            }

            const RedConfigure::Themes::ThemeSourceTransform sourceTransform = transform == ThemeGroupTransform::Darken
                                                                                   ? RedConfigure::Themes::ThemeSourceTransform::Darken10
                                                                                   : RedConfigure::Themes::ThemeSourceTransform::BlendAccent16;
            changed = _session.GetThemePreviewModel().WrapSourceWithTransform(key, sourceTransform) || changed;
        }

        if (changed)
        {
            SyncThemeColorEditor();
            _themeColorGrid->NotifyDataChanged();
            SyncExportPathLabels();
            SyncExportPreviews();
            _themePreview->Refresh();
        }
    }

    void ResetSelectedThemeColor()
    {
        const std::wstring selectedKey(_colorKeyCombo->GetSelectedValue());
        if (selectedKey.empty())
        {
            return;
        }

        if (! _session.GetThemePreviewModel().ResetOverride(selectedKey))
        {
            const size_t affectedCount = _session.GetThemePreviewModel().GetAffected(selectedKey).size();
            _themeColorStatusLabel->SetText(affectedCount == 0u ? LoadAppString(_instance, IDS_REDCONFIGURE_THEME_COLOR_INVALID)
                                                                : FormatStringResource(_instance, IDS_REDCONFIGURE_FMT_THEME_PALETTE_REFERENCED, affectedCount));
            _themeColorStatusLabel->SetTextColor(ResolvePalette().errorText);
            return;
        }
        SyncThemeColorEditor();
        _themeColorGrid->NotifyDataChanged();
        SyncExportPathLabels();
        SyncExportPreviews();
        _themePreview->Refresh();
    }

    void SetValidationStatus(RedConfigure::Localization::PlaceholderStatus status)
    {
        if (status == RedConfigure::Localization::PlaceholderStatus::Ok)
        {
            _validationLabel->SetText({});
            _validationLabel->SetTextColor(std::nullopt);
        }
        else
        {
            _validationLabel->SetText(PlaceholderStatusText(_instance, status));
            _validationLabel->SetTextColor(ResolvePalette().errorText);
        }
        SyncScopeLabel();
    }

    void SyncScopeLabel()
    {
        if (! _scopeLabel)
        {
            return;
        }

        std::wstring text = LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SCOPE) + L": ";
        text += _workspaceRoot.wstring();
        text += L" | " + FormatStringResource(_instance,
                                               IDS_REDCONFIGURE_FMT_DIRTY_SCOPE,
                                               _session.GetDirtyLocalizationCellCount(),
                                               LoadAppString(_instance,
                                                             _session.IsThemeDirty() ? IDS_REDCONFIGURE_VALUE_YES : IDS_REDCONFIGURE_VALUE_NO)) +
                L" | ";
        if (_selectedPage == 0u)
        {
            text += LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_RESOURCE_OWNERS) + L" " + std::to_wstring(_session.GetWorkspace().resourceOwners.size());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_FILES) + L" " + std::to_wstring(_session.GetThemeCatalog().themes.size());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SCAN_ERRORS) + L" " + std::to_wstring(_session.GetWorkspace().errors.size());
        }
        else if (_selectedPage == 1u)
        {
            text += LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_OWNERS) + L" " + std::to_wstring(_localizationReviewViewOptions.visibleOwnerNames.size()) +
                    L"/" + std::to_wstring(_session.GetWorkspace().resourceOwners.size());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_LANGUAGES) + L" " +
                    std::to_wstring(_localizationReviewViewOptions.visibleCultureNames.size()) + L"/" +
                    std::to_wstring(_session.GetLocalizationReviewCultures().size());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_NAME) + L" " + _session.GetThemePreviewModel().GetTheme().name;
        }
        else
        {
            text += LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_CULTURE) + L" " + std::wstring(_session.GetCultureName());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_ACTIVE_OWNER) + L" " + std::wstring(_session.GetActiveResourceOwnerName());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_NAME) + L" " + _session.GetThemePreviewModel().GetTheme().name;
        }
        _scopeLabel->SetText(std::move(text));
    }

    void RunValidation(bool expandDrawer)
    {
        _validationSummary = _session.Validate();
        _warningsAcknowledged = false;
        if (expandDrawer)
        {
            _validationDrawerExpanded = true;
        }
        std::wstring text = _validationSummary.issues.empty()
                                ? LoadAppString(_instance, IDS_REDCONFIGURE_VALIDATION_OK)
                                : FormatStringResource(_instance,
                                                       IDS_REDCONFIGURE_FMT_VALIDATION_SUMMARY,
                                                       _validationSummary.errorCount,
                                                       _validationSummary.warningCount);
        const size_t visibleIssueCount = std::min<size_t>(8u, _validationSummary.issues.size());
        for (size_t index = 0u; index < visibleIssueCount; ++index)
        {
            const auto& issue = _validationSummary.issues[index];
            text += L"\n" + ValidationCategoryText(_instance, issue.category) + L": " + ValidationMessageText(_instance, issue);
            if (! issue.resourceId.empty()) text += L" [" + issue.resourceId + L"]";
            if (! issue.cultureName.empty()) text += L" (" + issue.cultureName + L")";
        }
        _validationDrawerLabel->SetText(text);
        _validationDrawerLabel->SetTextColor(_validationSummary.errorCount > 0u ? std::optional<D2D1_COLOR_F>(ResolvePalette().errorText)
                                                                                : std::optional<D2D1_COLOR_F>(ResolvePalette().subduedText));
        _statusLabel->SetText(_validationSummary.issues.empty()
                                  ? LoadAppString(_instance, IDS_REDCONFIGURE_VALIDATION_OK)
                                  : FormatStringResource(_instance,
                                                         IDS_REDCONFIGURE_FMT_VALIDATION_SUMMARY,
                                                         _validationSummary.errorCount,
                                                         _validationSummary.warningCount));
        SyncVisibility();
        LayoutControls();
    }

    [[nodiscard]] bool ConfirmExportAllowed()
    {
        const bool warningsAcknowledged = _warningsAcknowledged;
        RunValidation(false);
        _warningsAcknowledged = warningsAcknowledged;
        if (! _validationSummary.CanExport())
        {
            _validationDrawerExpanded = true;
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_EXPORT_BLOCKED));
            SyncVisibility();
            LayoutControls();
            return false;
        }
        if (_validationSummary.warningCount > 0u && ! _warningsAcknowledged)
        {
            _warningsAcknowledged = true;
            _validationDrawerExpanded = true;
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_EXPORT_WARNINGS_CONFIRM));
            SyncVisibility();
            LayoutControls();
            return false;
        }
        return true;
    }

    void SelectAdjacentProblem(bool next)
    {
        const auto rows = _session.GetLocalizationReviewRows();
        if (rows.empty()) return;
        for (size_t offset = 1u; offset <= rows.size(); ++offset)
        {
            const size_t candidate = next ? (_selectedReviewRow + offset) % rows.size()
                                          : (_selectedReviewRow + rows.size() - (offset % rows.size())) % rows.size();
            const bool problem = std::ranges::any_of(rows[candidate].targets,
                                                     [](const auto& cell)
            {
                return cell.validation.status != RedConfigure::Localization::PlaceholderStatus::Ok ||
                       (! cell.hasExistingTranslation && ! cell.dirty);
            });
            if (problem)
            {
                _selectedReviewRow = candidate;
                RebuildTranslationView(true);
                return;
            }
        }
    }

    void PinSelectedLanguage()
    {
        if (_selectedReviewCulture.empty()) return;
        const auto columns = _languageColumns.Get();
        const auto it = std::ranges::find_if(columns, [this](const auto& column) { return column.cultureName == _selectedReviewCulture; });
        if (it == columns.end()) return;
        static_cast<void>(_languageColumns.SetPinned(_selectedReviewCulture, ! it->pinned));
        _localizationReviewViewOptions.visibleCultureNames = _languageColumns.GetOrderedCultures();
        SyncLocalizationTagPickers();
        RebuildTranslationView(true);
    }

    void MoveSelectedLanguage(bool right)
    {
        auto cultures = _languageColumns.GetOrderedCultures();
        const auto it = std::ranges::find(cultures, _selectedReviewCulture);
        if (it == cultures.end()) return;
        const size_t index = static_cast<size_t>(std::distance(cultures.begin(), it));
        const size_t newIndex = right ? std::min(index + 1u, cultures.size() - 1u) : (index == 0u ? 0u : index - 1u);
        if (_languageColumns.Move(_selectedReviewCulture, newIndex))
        {
            _localizationReviewViewOptions.visibleCultureNames = _languageColumns.GetOrderedCultures();
            SyncLocalizationTagPickers();
            RebuildTranslationView(true);
        }
    }

    void RemoveSelectedLanguage()
    {
        if (_selectedReviewCulture.empty() || ! _languageColumns.Remove(_selectedReviewCulture)) return;
        _localizationReviewViewOptions.visibleCultureNames = _languageColumns.GetOrderedCultures();
        _selectedReviewCulture = _localizationReviewViewOptions.visibleCultureNames.empty() ? std::wstring{}
                                                                                             : _localizationReviewViewOptions.visibleCultureNames.front();
        SyncLocalizationTagPickers();
        RebuildTranslationView(true);
    }

    void PreviewOrApplyLocalizationBatch()
    {
        if (_selectedReviewCulture.empty()) return;
        RedConfigure::Workflow::LocalizationBatchKind kind = RedConfigure::Workflow::LocalizationBatchKind::CopyEnglish;
        const std::wstring value(_localizationBatchCombo->GetSelectedValue());
        if (value == L"copyCulture") kind = RedConfigure::Workflow::LocalizationBatchKind::CopyCulture;
        else if (value == L"clear") kind = RedConfigure::Workflow::LocalizationBatchKind::Clear;
        else if (value == L"findReplace") kind = RedConfigure::Workflow::LocalizationBatchKind::FindReplace;
        else if (value == L"normalize") kind = RedConfigure::Workflow::LocalizationBatchKind::NormalizePlaceholderWhitespace;
        else if (value == L"accelerators") kind = RedConfigure::Workflow::LocalizationBatchKind::PreserveAccelerators;
        else if (value == L"reviewed") kind = RedConfigure::Workflow::LocalizationBatchKind::MarkReviewed;
        RedConfigure::Workflow::LocalizationBatchRequest request{.kind = kind,
                                                                 .targetCulture = _selectedReviewCulture,
                                                                 .findText = std::wstring(_commandSearchField->GetText())};
        if (kind == RedConfigure::Workflow::LocalizationBatchKind::FindReplace)
        {
            const size_t equals = request.findText.find(L'=');
            if (equals != std::wstring::npos)
            {
                request.replaceText = request.findText.substr(equals + 1u);
                request.findText.resize(equals);
            }
        }
        for (const std::wstring& culture : _localizationReviewViewOptions.visibleCultureNames)
        {
            if (culture != _selectedReviewCulture)
            {
                request.sourceCulture = culture;
                break;
            }
        }
        const BatchInteraction interaction = _localizationPresenter.Execute(_session, request);
        if (interaction.phase == BatchInteractionPhase::Apply)
        {
            if (interaction.result == RedConfigure::Workflow::BatchApprovalResult::Applied)
            {
                RebuildTranslationView(true);
                SyncExportPreviews();
                RunValidation(false);
            }
            else
            {
                _statusLabel->SetText(LoadAppString(_instance,
                                                    interaction.result == RedConfigure::Workflow::BatchApprovalResult::Stale
                                                        ? IDS_REDCONFIGURE_STATUS_BATCH_STALE
                                                        : IDS_REDCONFIGURE_STATUS_BATCH_FAILED));
            }
            return;
        }
        if (interaction.result == RedConfigure::Workflow::BatchApprovalResult::Invalid)
        {
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_BATCH_INVALID));
            return;
        }
        if (interaction.result == RedConfigure::Workflow::BatchApprovalResult::NoChanges)
        {
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_BATCH_NO_CHANGES));
            return;
        }
        std::wstring previewText = FormatStringResource(_instance, IDS_REDCONFIGURE_FMT_BATCH_PREVIEW, interaction.changeCount);
        if (interaction.firstChange.has_value())
        {
            const BatchChangeSummary& first = interaction.firstChange.value();
            previewText += L"\n" + FormatStringResource(_instance, IDS_REDCONFIGURE_FMT_BEFORE_AFTER, first.identity, first.before, first.after);
        }
        _statusLabel->SetText(previewText);
    }

    void PasteLocalizationMatrix()
    {
        if (! GetHost()) return;
        const std::optional<std::wstring> text = GetHost()->ReadTextFromClipboard();
        if (! text.has_value()) return;
        const auto rows = _session.GetLocalizationReviewRows();
        if (_selectedReviewRow >= rows.size()) return;
        size_t cultureIndex = 0u;
        for (size_t index = 0u; index < rows[_selectedReviewRow].targets.size(); ++index)
        {
            if (rows[_selectedReviewRow].targets[index].cultureName == _selectedReviewCulture)
            {
                cultureIndex = index;
                break;
            }
        }
        if (_session.ApplyClipboardMatrix(_selectedReviewRow, cultureIndex, text.value()))
        {
            RebuildTranslationView(true);
            SyncExportPreviews();
            RunValidation(false);
        }
    }

    void PreviewOrApplyThemeRecipe()
    {
        RedConfigure::Workflow::ThemeRecipe recipe = RedConfigure::Workflow::ThemeRecipe::DarkVariant;
        const std::wstring value(_themeRecipeCombo->GetSelectedValue());
        if (value == L"light") recipe = RedConfigure::Workflow::ThemeRecipe::LightVariant;
        else if (value == L"accent") recipe = RedConfigure::Workflow::ThemeRecipe::AccentRecolor;
        else if (value == L"soft") recipe = RedConfigure::Workflow::ThemeRecipe::SoftenedSelections;
        else if (value == L"contrast") recipe = RedConfigure::Workflow::ThemeRecipe::IncreasedContrast;
        else if (value == L"semantic") recipe = RedConfigure::Workflow::ThemeRecipe::SemanticStatusColors;
        else if (value == L"alpha") recipe = RedConfigure::Workflow::ThemeRecipe::SetAlpha;
        else if (value == L"replace") recipe = RedConfigure::Workflow::ThemeRecipe::ReplaceReference;
        else if (value == L"convert") recipe = RedConfigure::Workflow::ThemeRecipe::ConvertSolidsToReferences;
        else if (value == L"remove") recipe = RedConfigure::Workflow::ThemeRecipe::RemoveOverrides;
        const std::wstring group = ThemeKeyGroup(_colorKeyCombo->GetSelectedValue());
        RedConfigure::Workflow::ThemeMassRequest request{.recipe = recipe,
                                                         .argument = std::wstring(_commandSearchField->GetText()),
                                                         .alphaPercent = static_cast<uint32_t>(std::clamp(_themeAlphaSlider->GetValue(), 0.0, 100.0))};
        for (const std::wstring& key : _themeColorKeys)
        {
            if (ThemeKeyGroup(key) == group) request.keys.push_back(key);
        }
        const BatchInteraction interaction = _themesPresenter.Execute(_session, request);
        if (interaction.phase == BatchInteractionPhase::Apply)
        {
            if (interaction.result == RedConfigure::Workflow::BatchApprovalResult::Applied)
            {
                SyncThemeColorKeyCombo({});
                SyncThemeColorEditor();
                SyncExportPreviews();
                _themePreview->Refresh();
            }
            else
            {
                _statusLabel->SetText(LoadAppString(_instance,
                                                    interaction.result == RedConfigure::Workflow::BatchApprovalResult::Stale
                                                        ? IDS_REDCONFIGURE_STATUS_BATCH_STALE
                                                        : IDS_REDCONFIGURE_STATUS_BATCH_FAILED));
            }
            return;
        }
        if (interaction.result == RedConfigure::Workflow::BatchApprovalResult::Invalid)
        {
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_BATCH_INVALID));
            return;
        }
        if (interaction.result == RedConfigure::Workflow::BatchApprovalResult::NoChanges)
        {
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_BATCH_NO_CHANGES));
            return;
        }
        std::wstring previewText = FormatStringResource(_instance, IDS_REDCONFIGURE_FMT_BATCH_PREVIEW, interaction.changeCount);
        if (interaction.firstChange.has_value())
        {
            const BatchChangeSummary& first = interaction.firstChange.value();
            previewText += L"\n" + FormatStringResource(_instance, IDS_REDCONFIGURE_FMT_BEFORE_AFTER, first.identity, first.before, first.after);
        }
        _statusLabel->SetText(previewText);
    }

    void CopySelectedThemeColor(bool authored)
    {
        const std::wstring key(_colorKeyCombo->GetSelectedValue());
        if (key.empty() || ! GetHost()) return;
        std::wstring text = authored ? _session.GetThemePreviewModel().GetAuthoredColorText(key) : std::wstring{};
        if (! authored)
        {
            if (const std::optional<uint32_t> color = _session.GetThemePreviewModel().GetEffectiveColor(key))
            {
                text = Common::Settings::FormatColor(color.value());
            }
        }
        if (! text.empty())
        {
            static_cast<void>(GetHost()->CopyTextToClipboard(text));
            _recentThemeColors.erase(std::remove(_recentThemeColors.begin(), _recentThemeColors.end(), text), _recentThemeColors.end());
            _recentThemeColors.insert(_recentThemeColors.begin(), text);
            if (_recentThemeColors.size() > 8u) _recentThemeColors.resize(8u);
            SyncThemeColorEditor();
        }
    }

    void ImportThemeFromField()
    {
        const std::filesystem::path path(std::wstring(_themeImportPathField->GetText()));
        if (SUCCEEDED(_session.ImportTheme(path)))
        {
            SyncFromSession();
        }
        else
        {
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_EXPORT_FAILED) + L": " + path.wstring());
        }
    }

    void DuplicateActiveTheme()
    {
        const auto& active = _session.GetThemePreviewModel().GetTheme();
        const std::wstring copyLabel = LoadAppString(_instance, IDS_REDCONFIGURE_THEME_COPY_LABEL);
        for (uint32_t sequence = 1u; sequence <= 100u; ++sequence)
        {
            const auto candidate = RedConfigure::Workflow::BuildDuplicateThemeCandidate(active.id, active.name, copyLabel, sequence);
            if (! candidate.has_value())
            {
                _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_DUPLICATE_INVALID));
                return;
            }
            const RedConfigure::DuplicateThemeResult result = _session.DuplicateActiveTheme(candidate->id, candidate->name);
            if (result == RedConfigure::DuplicateThemeResult::Created)
            {
                SyncFromSession();
                return;
            }
            if (result != RedConfigure::DuplicateThemeResult::Collision)
            {
                _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_DUPLICATE_INVALID));
                return;
            }
        }
        _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_DUPLICATE_EXHAUSTED));
    }

    [[nodiscard]] ThemePalette ResolvePalette() const noexcept
    {
        return GetHost() ? GetHost()->GetTheme() : MakeDefaultThemePalette(false);
    }

    void ExportLocalization()
    {
        if (! ConfirmExportAllowed()) return;
        if (_session.GetLocalizationReviewRows().empty())
        {
            const std::filesystem::path path = _session.GetDefaultLocalizationExportPath();
            const HRESULT hr                 = _session.ExportLocalization(path);
            _statusLabel->SetText(LoadAppString(_instance, SUCCEEDED(hr) ? IDS_REDCONFIGURE_STATUS_EXPORT_RC_DONE : IDS_REDCONFIGURE_STATUS_EXPORT_FAILED) +
                                  L": " + path.wstring());
            if (SUCCEEDED(hr))
            {
                RunValidation(false);
                _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_EXPORTED_VALIDATED) + L": " + path.wstring());
            }
            return;
        }

        size_t exportedFileCount               = 0u;
        const std::filesystem::path outputRoot = _session.GetDefaultLocalizationExportPath().parent_path();
        const HRESULT hr                       = _session.ExportLocalizationReview(outputRoot, &exportedFileCount);

        if (SUCCEEDED(hr))
        {
            _statusLabel->SetText(FormatStringResource(_instance, IDS_REDCONFIGURE_STATUS_REVIEW_EXPORT_DONE, exportedFileCount) + L": " +
                                  outputRoot.wstring());
            RunValidation(false);
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_EXPORTED_VALIDATED) + L": " + outputRoot.wstring());
        }
        else
        {
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_EXPORT_FAILED) + L": " + outputRoot.wstring());
        }
    }

    void ExportTheme()
    {
        if (! ConfirmExportAllowed()) return;
        const std::filesystem::path path = _session.GetDefaultThemeExportPath();
        const HRESULT hr                 = _session.ExportTheme(path);
        _statusLabel->SetText(LoadAppString(_instance, SUCCEEDED(hr) ? IDS_REDCONFIGURE_STATUS_EXPORT_THEME_DONE : IDS_REDCONFIGURE_STATUS_EXPORT_FAILED) +
                              L": " + path.wstring());
        if (SUCCEEDED(hr))
        {
            RunValidation(false);
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_EXPORTED_VALIDATED) + L": " + path.wstring());
        }
    }

    void SyncVisibility() noexcept
    {
        auto setVisible = [](std::span<Control* const> controls, bool visible) noexcept
        {
            for (Control* control : controls)
            {
                if (control)
                {
                    control->SetVisible(visible);
                }
            }
        };

        Control* workspaceControls[] = {_workspaceRootLabel, _workspaceRootField, _reloadButton, _ownerCountLabel, _themeFileCountLabel, _scanErrorCountLabel};
        Control* retiredLocalizationControls[] = {_ownerSelectorLabel, _ownerCombo};
        Control* localizationControls[]        = {_localizationPageScroll,
                                                  _localizationPageContent,
                                                  _ownerFilterLabel,
                                                  _ownerFilterPicker,
                                                  _languageFilterLabel,
                                                  _languageFilterPicker,
                                                  _translationCountLabel,
                                                  _activeOwnerLabel,
                                                  _localizationSearchLabel,
                                                  _localizationSearchField,
                                                  _localizationIdFilterLabel,
                                                  _localizationIdFilterField,
                                                  _localizationStatusFilterLabel,
                                                  _localizationStatusFilterCombo,
                                                  _translationGrid,
                                                  _sourceTextLabel,
                                                  _sourceTextField,
                                                  _targetTextLabel,
                                                  _targetEditorsPanel,
                                                  _validationLabel,
                                                  _cultureLabel,
                                                  _cultureCombo,
                                                  _previousProblemButton,
                                                  _nextProblemButton,
                                                  _pasteMatrixButton,
                                                  _pinLanguageButton,
                                                  _moveLanguageLeftButton,
                                                  _moveLanguageRightButton,
                                                  _removeLanguageButton,
                                                  _localizationBatchCombo,
                                                  _applyLocalizationBatchButton,
                                                  _localizationExampleLabel,
                                                  _localizationExample};
        Control* themeDesignerControls[]       = {
            _themeSelectorLabel,  _themeCombo,         _themeNameLabel,        _themePathLabel,           _themeErrorLabel,    _colorKeyFilterLabel,
            _colorKeyFilterField, _colorKeyLabel,      _themeColorGrid,        _colorValueLabel,          _previousColorLabel, _previousColorSwatch,
            _colorSwatch,         _colorValueCombo,    _themeColorStatusLabel, _paletteNameLabel,          _paletteNameField,  _addPaletteButton,
            _renamePaletteButton, _previewSeedLabel,   _previewSeedCombo,      _themeExpressionHelpLabel,  _darkenGroupButton, _blendAccentGroupButton,
            _resetColorButton,    _themePreviewScroll, _themePreview, _themeImportPathField, _themeImportButton, _themeDuplicateButton,
            _themeResetButton, _themeSceneLabel, _themeSceneCombo, _themeRecipeLabel, _themeRecipeCombo, _applyThemeRecipeButton,
            _themeAlphaSlider, _copyEffectiveButton, _copyOverrideButton};
        Control* exportControls[] = {_defaultRcPathLabel,
                                     _defaultThemePathLabel,
                                     _exportRcButton,
                                     _exportThemeButton,
                                     _reviewOutputScroll};

        setVisible(workspaceControls, false);
        setVisible(retiredLocalizationControls, false);
        setVisible(localizationControls, false);
        setVisible(themeDesignerControls, false);
        setVisible(exportControls, false);

        switch (_selectedPage)
        {
            case 0u: setVisible(workspaceControls, true); break;
            case 1u: setVisible(localizationControls, true); break;
            case 2u: setVisible(themeDesignerControls, true); break;
            case 3u: setVisible(exportControls, true); break;
            default: break;
        }

        if (_colorKeyCombo)
        {
            _colorKeyCombo->SetVisible(false);
        }
        if (_validationDrawerLabel)
        {
            _validationDrawerLabel->SetVisible(_validationDrawerExpanded);
        }
    }

    [[nodiscard]] float TargetEditorsContentHeight(float widthDip) const noexcept
    {
        if (_targetEditors.empty())
        {
            return 0.0f;
        }

        const float labelW     = TargetEditorLabelWidth(widthDip);
        const float fieldWidth = std::max(0.0f, widthDip - labelW - 6.0f);
        const float rowGap     = _targetEditors.size() > 1u ? 4.0f : 0.0f;
        float height           = rowGap * static_cast<float>(_targetEditors.size() - 1u);
        for (const TargetEditor& editor : _targetEditors)
        {
            height += EditorFieldHeightForText(editor.field ? editor.field->GetText() : std::wstring_view{}, fieldWidth);
        }
        return height;
    }

    void LayoutControls() noexcept
    {
        const D2D1_RECT_F bounds = GetBounds();
        const float width        = std::max(0.0f, bounds.right - bounds.left);
        const float height       = std::max(0.0f, bounds.bottom - bounds.top);
        const float margin       = 18.0f;
        const float gap          = 12.0f;
        const bool collapsedNav  = width < 980.0f;
        const float navWidth     = collapsedNav ? 44.0f : 220.0f;
        const float navButtonH   = collapsedNav ? 40.0f : 36.0f;
        const float contentLeft  = margin + navWidth + (collapsedNav ? 10.0f : 18.0f);
        const float contentRight = std::max(contentLeft, width - margin);
        const float contentW     = std::max(0.0f, contentRight - contentLeft);
        const float contentTop   = margin;
        const float bodyTop      = contentTop + 108.0f;
        const float statusTop    = std::max(bodyTop, height - margin - 42.0f);
        const float drawerTop    = _validationDrawerExpanded ? std::max(bodyTop, statusTop - 112.0f) : statusTop;
        const float bodyBottom   = std::max(bodyTop, drawerTop - gap);

        float navY       = margin;
        const auto pages = RedConfigure::GetPageDefinitions();
        for (size_t index = 0u; index < _navButtons.size(); ++index)
        {
            Button* button = _navButtons[index];
            if (! button)
            {
                continue;
            }

            const std::wstring title = index < pages.size() ? LoadAppString(_instance, pages[index].titleResourceId) : std::wstring{};
            if (collapsedNav)
            {
                const std::wstring_view icon = index < kPageIconGlyphs.size() ? kPageIconGlyphs[index] : L"\xE10F";
                if (button->GetText() != icon)
                {
                    button->SetText(std::wstring(icon));
                }
                if (button->GetVariant() != ButtonVariant::IconOnly)
                {
                    button->SetVariant(ButtonVariant::IconOnly);
                }
                if (button->GetTooltipText() != title)
                {
                    button->SetTooltipText(title);
                }
            }
            else
            {
                if (button->GetText() != title)
                {
                    button->SetText(title);
                }
                if (button->GetVariant() != ButtonVariant::Standard)
                {
                    button->SetVariant(ButtonVariant::Standard);
                }
                if (! button->GetTooltipText().empty())
                {
                    button->SetTooltipText({});
                }
            }
            button->SetBounds(D2D1::RectF(margin, navY, margin + navWidth, navY + navButtonH));
            navY += navButtonH + 8.0f;
        }

        _titleLabel->SetBounds(D2D1::RectF(contentLeft, contentTop, contentRight, contentTop + 34.0f));
        _descriptionLabel->SetBounds(D2D1::RectF(contentLeft, contentTop + 34.0f, contentLeft, contentTop + 34.0f));
        _scopeLabel->SetBounds(D2D1::RectF(contentLeft, contentTop + 38.0f, contentRight, contentTop + 62.0f));
        const float commandTop = contentTop + 68.0f;
        const float commandH = 32.0f;
        const float commandGap = 6.0f;
        const float actionW = 82.0f;
        const float reviewW = 124.0f;
        const float actionsWidth = (actionW * 4.0f) + reviewW + (commandGap * 5.0f);
        const float searchRight = std::max(contentLeft + 120.0f, contentRight - actionsWidth);
        _commandSearchField->SetBounds(D2D1::RectF(contentLeft, commandTop, searchRight, commandTop + commandH));
        float commandX = searchRight + commandGap;
        _validateButton->SetBounds(D2D1::RectF(commandX, commandTop, commandX + actionW, commandTop + commandH));
        commandX += actionW + commandGap;
        _undoButton->SetBounds(D2D1::RectF(commandX, commandTop, commandX + actionW, commandTop + commandH));
        commandX += actionW + commandGap;
        _redoButton->SetBounds(D2D1::RectF(commandX, commandTop, commandX + actionW, commandTop + commandH));
        commandX += actionW + commandGap;
        _validationDrawerButton->SetBounds(D2D1::RectF(commandX, commandTop, commandX + actionW, commandTop + commandH));
        commandX += actionW + commandGap;
        _reviewExportButton->SetBounds(D2D1::RectF(commandX, commandTop, contentRight, commandTop + commandH));
        if (_validationDrawerExpanded)
        {
            _validationDrawerLabel->SetBounds(D2D1::RectF(contentLeft, drawerTop, contentRight, statusTop - 4.0f));
        }
        _statusLabel->SetBounds(D2D1::RectF(contentLeft, statusTop, contentRight, height - margin));

        switch (_selectedPage)
        {
            case 0u: LayoutWorkspacePage(contentLeft, bodyTop, contentRight, bodyBottom); break;
            case 1u: LayoutLocalizationWorkbenchPage(contentLeft, bodyTop, contentRight, bodyBottom); break;
            case 2u: LayoutThemeWorkbenchPage(contentLeft, bodyTop, contentRight, bodyBottom); break;
            case 3u: LayoutExportPage(contentLeft, bodyTop, contentRight, bodyBottom); break;
            default: break;
        }

        (void)contentW;
    }

    void LayoutWorkspacePage(float left, float top, float right, float bottom) noexcept
    {
        const float labelW  = 140.0f;
        const float rowH    = 32.0f;
        const float buttonW = 120.0f;
        const float gapX    = 12.0f;
        float y             = top;
        _workspaceRootLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
        _workspaceRootField->SetBounds(D2D1::RectF(left + labelW, y, std::max(left + labelW, right - buttonW - gapX), y + rowH));
        _reloadButton->SetBounds(D2D1::RectF(right - 120.0f, y, right, y + rowH));
        y += rowH + 30.0f;
        _ownerCountLabel->SetBounds(D2D1::RectF(left, y, right, y + 28.0f));
        _themeFileCountLabel->SetBounds(D2D1::RectF(left, y + 34.0f, right, y + 62.0f));
        _scanErrorCountLabel->SetBounds(D2D1::RectF(left, y + 68.0f, right, std::min(bottom, y + 96.0f)));
    }

    void LayoutLocalizationWorkbenchPage(float left, float top, float right, float bottom) noexcept
    {
        const float viewportH = std::max(0.0f, bottom - top);
        if (_localizationPageScroll)
        {
            _localizationPageScroll->SetBounds(D2D1::RectF(left, top, right, bottom));
        }

        const float scrollbarReserveW = _localizationPageScroll ? (_localizationPageScroll->GetScrollbarThickness() + 2.0f) : 0.0f;
        const float contentRight      = std::max(left, right - scrollbarReserveW);
        const float rowH              = 32.0f;
        const float rowGap            = 10.0f;
        const float labelH            = 22.0f;
        const float contentW          = std::max(0.0f, contentRight - left);
        const bool compact            = contentW < 760.0f;
        const auto tagPickerHeight = [](const TagPicker* picker, float widthDip) noexcept { return picker ? picker->GetPreferredHeightDip(widthDip) : 32.0f; };
        float y                    = top;

        if (_translationGrid)
        {
            _translationGrid->SetVisualMode(compact ? GridVisualMode::FolderView : GridVisualMode::Standard);
            _translationGrid->SetRowHeightDip(kLocalizationGridRowHeightDip);
            _translationGrid->SetLineClamp(2u);
        }

        if (compact)
        {
            _ownerFilterLabel->SetBounds(D2D1::RectF(left, y, contentRight, y + labelH));
            y += labelH + 2.0f;
            const float ownerPickerH = tagPickerHeight(_ownerFilterPicker, contentW);
            _ownerFilterPicker->SetBounds(D2D1::RectF(left, y, contentRight, y + ownerPickerH));
            y += ownerPickerH + 8.0f;

            _languageFilterLabel->SetBounds(D2D1::RectF(left, y, contentRight, y + labelH));
            y += labelH + 2.0f;
            const float languagePickerH = tagPickerHeight(_languageFilterPicker, contentW);
            _languageFilterPicker->SetBounds(D2D1::RectF(left, y, contentRight, y + languagePickerH));
            y += languagePickerH + 8.0f;

            _translationCountLabel->SetBounds(D2D1::RectF(left, y, contentRight, y + 22.0f));
            y += 30.0f;

            _localizationSearchLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + 70.0f, y + rowH));
            _localizationSearchField->SetBounds(D2D1::RectF(left + 72.0f, y, contentRight, y + rowH));
            y += rowH + 8.0f;

            const float mid = left + ((contentW - rowGap) / 2.0f);
            _localizationIdFilterLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + 42.0f, y + rowH));
            _localizationIdFilterField->SetBounds(D2D1::RectF(left + 44.0f, y, mid, y + rowH));
            _localizationStatusFilterLabel->SetBounds(D2D1::RectF(mid + rowGap, y + 6.0f, mid + rowGap + 72.0f, y + rowH));
            _localizationStatusFilterCombo->SetBounds(D2D1::RectF(mid + rowGap + 74.0f, y, contentRight, y + rowH));
            y += rowH + 10.0f;
        }
        else
        {
            const float halfW     = (contentW - rowGap) / 2.0f;
            const float rightColL = left + halfW + rowGap;

            _ownerFilterLabel->SetBounds(D2D1::RectF(left, y, left + halfW, y + labelH));
            _languageFilterLabel->SetBounds(D2D1::RectF(rightColL, y, contentRight, y + labelH));
            y += labelH + 2.0f;
            const float ownerPickerH    = tagPickerHeight(_ownerFilterPicker, halfW);
            const float languagePickerH = tagPickerHeight(_languageFilterPicker, std::max(0.0f, contentRight - rightColL));
            _ownerFilterPicker->SetBounds(D2D1::RectF(left, y, left + halfW, y + ownerPickerH));
            _languageFilterPicker->SetBounds(D2D1::RectF(rightColL, y, contentRight, y + languagePickerH));
            y += std::max(ownerPickerH, languagePickerH) + 8.0f;

            _translationCountLabel->SetBounds(D2D1::RectF(left, y, contentRight, y + 22.0f));
            y += 30.0f;

            const float searchW = std::clamp(contentW * 0.36f, 220.0f, 500.0f);
            const float idW     = std::clamp(contentW * 0.22f, 170.0f, 300.0f);
            const float statusW = 168.0f;
            float x             = left;
            _localizationSearchLabel->SetBounds(D2D1::RectF(x, y + 6.0f, x + 70.0f, y + rowH));
            _localizationSearchField->SetBounds(D2D1::RectF(x + 72.0f, y, x + searchW, y + rowH));
            x += searchW + rowGap;
            _localizationIdFilterLabel->SetBounds(D2D1::RectF(x, y + 6.0f, x + 42.0f, y + rowH));
            _localizationIdFilterField->SetBounds(D2D1::RectF(x + 44.0f, y, x + idW, y + rowH));
            x += idW + rowGap;
            _localizationStatusFilterLabel->SetBounds(D2D1::RectF(x, y + 6.0f, x + 72.0f, y + rowH));
            _localizationStatusFilterCombo->SetBounds(D2D1::RectF(x + 74.0f, y, std::min(contentRight, x + 74.0f + statusW), y + rowH));
            y += rowH + 10.0f;
        }

        const float cultureLabelW = 88.0f;
        const float smallButtonW = 82.0f;
        _cultureLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + cultureLabelW, y + rowH));
        const float cultureRight = std::min(contentRight, left + std::max(210.0f, contentW * 0.30f));
        _cultureCombo->SetBounds(D2D1::RectF(left + cultureLabelW, y, cultureRight, y + rowH));
        float languageX = cultureRight + rowGap;
        _pinLanguageButton->SetBounds(D2D1::RectF(languageX, y, languageX + smallButtonW, y + rowH));
        languageX += smallButtonW + 4.0f;
        _moveLanguageLeftButton->SetBounds(D2D1::RectF(languageX, y, languageX + smallButtonW, y + rowH));
        languageX += smallButtonW + 4.0f;
        _moveLanguageRightButton->SetBounds(D2D1::RectF(languageX, y, languageX + smallButtonW, y + rowH));
        languageX += smallButtonW + 4.0f;
        _removeLanguageButton->SetBounds(D2D1::RectF(languageX, y, std::min(contentRight, languageX + smallButtonW), y + rowH));
        y += rowH + 8.0f;

        const float navButtonW = 118.0f;
        _previousProblemButton->SetBounds(D2D1::RectF(left, y, left + navButtonW, y + rowH));
        _nextProblemButton->SetBounds(D2D1::RectF(left + navButtonW + 4.0f, y, left + (navButtonW * 2.0f) + 4.0f, y + rowH));
        _pasteMatrixButton->SetBounds(D2D1::RectF(left + (navButtonW * 2.0f) + 8.0f, y, left + (navButtonW * 3.0f) + 8.0f, y + rowH));
        const float batchLeft = left + (navButtonW * 3.0f) + 16.0f;
        const float applyW = 112.0f;
        _localizationBatchCombo->SetBounds(D2D1::RectF(batchLeft, y, std::max(batchLeft, contentRight - applyW - 6.0f), y + rowH));
        _applyLocalizationBatchButton->SetBounds(D2D1::RectF(std::max(batchLeft, contentRight - applyW), y, contentRight, y + rowH));
        y += rowH + 10.0f;

        const float editorLabelW             = TargetEditorLabelWidth(contentW);
        const float editorFieldLeft          = std::min(contentRight, left + editorLabelW + 6.0f);
        const float editorFieldW             = std::max(0.0f, contentRight - editorFieldLeft);
        const float sourceFieldH             = EditorFieldHeightForText(_sourceTextField ? _sourceTextField->GetText() : std::wstring_view{}, editorFieldW);
        const float targetContentH           = std::max(kLocalizationEditorMinFieldHeightDip, TargetEditorsContentHeight(contentW));
        const float targetViewportPreferredH = targetContentH;
        const float targetViewportMinH       = std::min(targetContentH, kLocalizationEditorMinFieldHeightDip);
        const float sourceBlockPreferredH    = labelH + 2.0f + sourceFieldH;
        const float targetBlockPreferredH    = labelH + 2.0f + targetViewportPreferredH;
        const float targetBlockMinH          = labelH + 2.0f + targetViewportMinH;
        const float editorGap                = 8.0f;
        const float editorPreferredH         = sourceBlockPreferredH + editorGap + targetBlockPreferredH;
        const float editorMinH               = sourceBlockPreferredH + editorGap + targetBlockMinH;
        const float minGridH                 = 32.0f + (2.0f * kLocalizationGridRowHeightDip);
        constexpr float exampleBlockH        = 122.0f;
        const float minContentBottom         = y + minGridH + 8.0f + editorPreferredH + exampleBlockH + 12.0f;
        float layoutBottom                   = std::max(bottom, top + kLocalizationPageMinContentHeightDip);
        layoutBottom                         = std::max(layoutBottom, minContentBottom);
        const float availableH               = std::max(0.0f, layoutBottom - y);
        const float gridPreferredH           = availableH - exampleBlockH - editorPreferredH - 8.0f;
        const float gridMaxH                 = std::max(minGridH, availableH - exampleBlockH - editorMinH - 8.0f);
        const float gridMinH                 = minGridH;
        const float gridH                    = std::clamp(gridPreferredH, gridMinH, gridMaxH);

        const float gridTop      = y;
        const float gridBottom   = gridTop + gridH;
        const float editorTop    = gridBottom + 8.0f;
        const float editorBottom = std::max(editorTop, layoutBottom - exampleBlockH - 12.0f);

        _translationGrid->SetBounds(D2D1::RectF(left, gridTop, contentRight, gridBottom));
        const float sourceLabelW      = std::min(112.0f, contentW);
        const float validationW       = std::clamp(contentW * 0.30f, 168.0f, 320.0f);
        const float validationLeft    = std::max(left + sourceLabelW, contentRight - validationW);
        const float selectedCellLeft  = std::min(contentRight, left + sourceLabelW + 8.0f);
        const float selectedCellRight = std::max(selectedCellLeft, validationLeft - 8.0f);
        _sourceTextLabel->SetBounds(D2D1::RectF(left, editorTop, left + sourceLabelW, editorTop + labelH));
        _activeOwnerLabel->SetBounds(D2D1::RectF(selectedCellLeft, editorTop, selectedCellRight, editorTop + labelH));
        _validationLabel->SetBounds(D2D1::RectF(validationLeft, editorTop, contentRight, editorTop + labelH));

        const float sourceFieldTop    = editorTop + labelH + 2.0f;
        const float sourceFieldBottom = sourceFieldTop + sourceFieldH;
        const float targetLabelTop    = sourceFieldBottom + editorGap;
        const float targetPanelTop    = targetLabelTop + labelH + 2.0f;
        const float targetPanelBottom = std::max(targetPanelTop + targetViewportPreferredH, editorBottom);
        _sourceTextField->SetBounds(D2D1::RectF(editorFieldLeft, sourceFieldTop, contentRight, sourceFieldBottom));
        _targetTextLabel->SetBounds(D2D1::RectF(left, targetLabelTop, contentRight, targetLabelTop + labelH));
        _targetEditorsPanel->SetBounds(D2D1::RectF(left, targetPanelTop, contentRight, targetPanelBottom));
        LayoutTargetEditorsPanel(_targetEditorsPanel->GetBounds());

        const float exampleLabelTop = targetPanelBottom + 8.0f;
        _localizationExampleLabel->SetBounds(D2D1::RectF(left, exampleLabelTop, contentRight, exampleLabelTop + 22.0f));
        _localizationExample->SetBounds(D2D1::RectF(left, exampleLabelTop + 24.0f, contentRight, exampleLabelTop + 112.0f));

        const float contentHeight = std::max(viewportH, exampleLabelTop + 120.0f - top);
        if (_localizationPageContent)
        {
            _localizationPageContent->SetBounds(D2D1::RectF(left, top, contentRight, top + contentHeight));
        }
        if (_localizationPageScroll)
        {
            _localizationPageScroll->SetContentHeight(contentHeight);
        }
    }

    void LayoutTargetEditorsPanel(const D2D1_RECT_F& bounds) noexcept
    {
        if (_targetEditors.empty())
        {
            if (_targetEditorsPanel)
            {
                _targetEditorsPanel->SetContentHeight(0.0f);
            }
            return;
        }

        const float width  = std::max(0.0f, bounds.right - bounds.left);
        const float height = std::max(0.0f, bounds.bottom - bounds.top);
        if (width <= 0.0f || height <= 0.0f)
        {
            for (TargetEditor& editor : _targetEditors)
            {
                if (editor.label)
                {
                    editor.label->SetBounds(D2D1::RectF(bounds.left, bounds.top, bounds.left, bounds.top));
                }
                if (editor.field)
                {
                    editor.field->SetBounds(D2D1::RectF(bounds.left, bounds.top, bounds.left, bounds.top));
                }
            }
            if (_targetEditorsPanel)
            {
                _targetEditorsPanel->SetContentHeight(0.0f);
            }
            return;
        }

        const float estimatedContentHeight = TargetEditorsContentHeight(width);
        const bool needsScrollbar          = _targetEditorsPanel && estimatedContentHeight > height;
        const float scrollbarW             = needsScrollbar ? _targetEditorsPanel->GetScrollbarThickness() + 2.0f : 0.0f;
        const float contentRight           = std::max(bounds.left, bounds.right - scrollbarW);
        const float rowGap                 = _targetEditors.size() > 1u ? 4.0f : 0.0f;
        const float contentWidth           = std::max(0.0f, contentRight - bounds.left);
        const float labelW                 = TargetEditorLabelWidth(contentWidth);
        const float fieldLeft              = std::min(contentRight, bounds.left + labelW + 6.0f);
        const float fieldWidth             = std::max(0.0f, contentRight - fieldLeft);
        float y                            = bounds.top;
        for (TargetEditor& editor : _targetEditors)
        {
            const float rowH      = EditorFieldHeightForText(editor.field ? editor.field->GetText() : std::wstring_view{}, fieldWidth);
            const float rowBottom = y + rowH;
            const float labelTop  = rowH >= 12.0f ? y + 5.0f : y;
            if (editor.label)
            {
                editor.label->SetBounds(D2D1::RectF(bounds.left, labelTop, std::min(contentRight, bounds.left + labelW), rowBottom));
            }
            if (editor.field)
            {
                editor.field->SetBounds(D2D1::RectF(fieldLeft, y, contentRight, rowBottom));
            }
            y = rowBottom + rowGap;
        }
        if (_targetEditorsPanel)
        {
            _targetEditorsPanel->SetContentHeight(std::max(0.0f, y - rowGap - bounds.top));
        }
    }

    void LayoutTranslationEditorPage(float left, float top, float right, float bottom) noexcept
    {
        const float editorTop = std::max(top + 180.0f, bottom - 190.0f);
        const float halfW     = (right - left - 12.0f) / 2.0f;
        _translationGrid->SetBounds(D2D1::RectF(left, top, right, editorTop - 12.0f));
        _sourceTextLabel->SetBounds(D2D1::RectF(left, editorTop, left + halfW, editorTop + 24.0f));
        _targetTextLabel->SetBounds(D2D1::RectF(left + halfW + 12.0f, editorTop, right, editorTop + 24.0f));
        _sourceTextField->SetBounds(D2D1::RectF(left, editorTop + 28.0f, left + halfW, bottom - 42.0f));
        _targetEditorsPanel->SetBounds(D2D1::RectF(left + halfW + 12.0f, editorTop + 28.0f, right, bottom - 42.0f));
        LayoutTargetEditorsPanel(_targetEditorsPanel->GetBounds());
        _validationLabel->SetBounds(D2D1::RectF(left, bottom - 32.0f, right - 140.0f, bottom));
        _exportRcButton->SetBounds(D2D1::RectF(right - 120.0f, bottom - 34.0f, right, bottom - 2.0f));
    }

    void LayoutThemeLibraryPage(float left, float top, float right, float) noexcept
    {
        _themeSelectorLabel->SetBounds(D2D1::RectF(left, top + 6.0f, left + 100.0f, top + 32.0f));
        _themeCombo->SetBounds(D2D1::RectF(left + 100.0f, top, right, top + 32.0f));
        _themeNameLabel->SetBounds(D2D1::RectF(left, top + 48.0f, right, top + 74.0f));
        _themePathLabel->SetBounds(D2D1::RectF(left, top + 82.0f, right, top + 108.0f));
        _themeFileCountLabel->SetBounds(D2D1::RectF(left, top + 116.0f, right, top + 142.0f));
        _themeErrorLabel->SetBounds(D2D1::RectF(left, top + 150.0f, right, top + 176.0f));
    }

    void LayoutThemeWorkbenchPage(float left, float top, float right, float bottom) noexcept
    {
        const float labelW   = 108.0f;
        const float rowH     = 32.0f;
        const float rowGap   = 10.0f;
        const float contentW = std::max(0.0f, right - left);
        _themeSelectorLabel->SetBounds(D2D1::RectF(left, top + 6.0f, left + labelW, top + rowH));
        _themeCombo->SetBounds(D2D1::RectF(left + labelW, top, right, top + rowH));
        _themeNameLabel->SetBounds(D2D1::RectF(left, top + 42.0f, right, top + 68.0f));
        _themePathLabel->SetBounds(D2D1::RectF(left, top + 68.0f, right, top + 92.0f));
        _themeErrorLabel->SetBounds(D2D1::RectF(left, top + 94.0f, right, top + 118.0f));

        const float libraryButtonW = 84.0f;
        _themeImportPathField->SetBounds(D2D1::RectF(left, top + 122.0f, std::max(left, right - (libraryButtonW * 3.0f) - 18.0f), top + 154.0f));
        float libraryX = std::max(left, right - (libraryButtonW * 3.0f) - 12.0f);
        _themeImportButton->SetBounds(D2D1::RectF(libraryX, top + 122.0f, libraryX + libraryButtonW, top + 154.0f));
        libraryX += libraryButtonW + 6.0f;
        _themeDuplicateButton->SetBounds(D2D1::RectF(libraryX, top + 122.0f, libraryX + libraryButtonW, top + 154.0f));
        libraryX += libraryButtonW + 6.0f;
        _themeResetButton->SetBounds(D2D1::RectF(libraryX, top + 122.0f, right, top + 154.0f));

        const bool sideBySide   = contentW >= 820.0f;
        const float editorTop   = top + 164.0f;
        const float editorRight = sideBySide ? std::min(left + 460.0f, left + std::max(390.0f, contentW * 0.46f)) : right;
        const float previewLeft = sideBySide ? editorRight + 18.0f : left;
        float previewTop        = sideBySide ? editorTop : bottom;

        float y = editorTop;
        _colorKeyFilterLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
        _colorKeyFilterField->SetBounds(D2D1::RectF(left + labelW, y, editorRight, y + rowH));
        y += rowH + rowGap;

        _colorKeyLabel->SetBounds(D2D1::RectF(left, y, editorRight, y + 24.0f));
        y += 26.0f;
        const float keyGridH = sideBySide ? 176.0f : 132.0f;
        _themeColorGrid->SetBounds(D2D1::RectF(left, y, editorRight, y + keyGridH));
        y += keyGridH + rowGap;

        _previousColorSwatch->SetBounds(D2D1::RectF(left + labelW, y + 4.0f, left + labelW + 24.0f, y + 28.0f));
        _previousColorLabel->SetBounds(D2D1::RectF(left + labelW + 34.0f, y + 6.0f, editorRight, y + rowH));
        y += rowH + rowGap;

        _colorValueLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
        _colorSwatch->SetBounds(D2D1::RectF(left + labelW, y, left + labelW + rowH, y + rowH));
        _colorValueCombo->SetBounds(D2D1::RectF(left + labelW + rowH + 10.0f, y, editorRight, y + rowH));
        y += rowH + 4.0f;
        _themeColorStatusLabel->SetBounds(D2D1::RectF(left + labelW, y, editorRight, y + 42.0f));
        y += 44.0f;

        const float paletteButtonGap = 6.0f;
        const float paletteButtonW   = 72.0f;
        _paletteNameLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
        _paletteNameField->SetBounds(D2D1::RectF(left + labelW, y, editorRight - (paletteButtonW * 2.0f) - (paletteButtonGap * 2.0f), y + rowH));
        _addPaletteButton->SetBounds(
            D2D1::RectF(editorRight - (paletteButtonW * 2.0f) - paletteButtonGap, y, editorRight - paletteButtonW - paletteButtonGap, y + rowH));
        _renamePaletteButton->SetBounds(D2D1::RectF(editorRight - paletteButtonW, y, editorRight, y + rowH));
        y += rowH + 4.0f;

        _previewSeedLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
        _previewSeedCombo->SetBounds(D2D1::RectF(left + labelW, y, editorRight, y + rowH));
        y += rowH + rowGap;

        _themeRecipeLabel->SetBounds(D2D1::RectF(left, y + 6.0f, left + labelW, y + rowH));
        _themeRecipeCombo->SetBounds(D2D1::RectF(left + labelW, y, editorRight - 112.0f, y + rowH));
        _applyThemeRecipeButton->SetBounds(D2D1::RectF(editorRight - 106.0f, y, editorRight, y + rowH));
        y += rowH + 6.0f;

        _themeAlphaSlider->SetBounds(D2D1::RectF(left + labelW, y, editorRight, y + 28.0f));
        y += 34.0f;

        const float copyW = std::max(100.0f, (editorRight - (left + labelW) - 6.0f) / 2.0f);
        _copyEffectiveButton->SetBounds(D2D1::RectF(left + labelW, y, left + labelW + copyW, y + rowH));
        _copyOverrideButton->SetBounds(D2D1::RectF(left + labelW + copyW + 6.0f, y, editorRight, y + rowH));
        y += rowH + rowGap;

        const float buttonGap = 8.0f;
        const float buttonW   = std::max(88.0f, (editorRight - (left + labelW) - (buttonGap * 2.0f)) / 3.0f);
        const float buttonX   = left + labelW;
        _darkenGroupButton->SetBounds(D2D1::RectF(buttonX, y, buttonX + buttonW, y + rowH));
        _blendAccentGroupButton->SetBounds(D2D1::RectF(buttonX + buttonW + buttonGap, y, buttonX + (buttonW * 2.0f) + buttonGap, y + rowH));
        _resetColorButton->SetBounds(D2D1::RectF(buttonX + (buttonW * 2.0f) + (buttonGap * 2.0f), y, editorRight, y + rowH));
        y += rowH + rowGap;

        const float helpBottom = sideBySide ? std::min(bottom, y + 82.0f) : std::min(bottom, y + 64.0f);
        _themeExpressionHelpLabel->SetBounds(D2D1::RectF(left, y, editorRight, std::max(y + 28.0f, helpBottom)));
        if (! sideBySide)
        {
            previewTop = std::min(bottom, std::max(y + 36.0f, helpBottom + 8.0f));
        }

        _themeSceneLabel->SetBounds(D2D1::RectF(previewLeft, previewTop + 6.0f, previewLeft + 100.0f, previewTop + 34.0f));
        _themeSceneCombo->SetBounds(D2D1::RectF(previewLeft + 102.0f, previewTop, right, previewTop + 32.0f));
        previewTop += 42.0f;
        const D2D1_RECT_F previewBounds = D2D1::RectF(previewLeft, previewTop, right, bottom);
        _themePreviewScroll->SetBounds(previewBounds);
        const float previewContentHeight = std::max(kThemePreviewContentHeightDip, previewBounds.bottom - previewBounds.top);
        _themePreviewScroll->SetContentHeight(previewContentHeight);
        const float previewContentRight = std::max(previewBounds.left + 120.0f, previewBounds.right - _themePreviewScroll->GetScrollbarThickness() - 2.0f);
        _themePreview->SetBounds(D2D1::RectF(previewBounds.left, previewBounds.top, previewContentRight, previewBounds.top + previewContentHeight));
    }

    void LayoutExportPage(float left, float top, float right, float bottom) noexcept
    {
        _defaultRcPathLabel->SetBounds(D2D1::RectF(left, top, right, top + 42.0f));
        _exportRcButton->SetBounds(D2D1::RectF(left, top + 48.0f, left + 140.0f, top + 80.0f));
        _defaultThemePathLabel->SetBounds(D2D1::RectF(left + 170.0f, top, right, top + 42.0f));
        _exportThemeButton->SetBounds(D2D1::RectF(left + 170.0f, top + 48.0f, left + 310.0f, top + 80.0f));

        const float previewTop    = top + 108.0f;
        const float viewportBottom = std::max(previewTop + 80.0f, bottom);
        _reviewOutputScroll->SetBounds(D2D1::RectF(left, previewTop, right, viewportBottom));

        const float scrollbarW = _reviewOutputScroll->GetScrollbarThickness() + 2.0f;
        const float cardRight  = std::max(left, right - scrollbarW);
        constexpr float labelH = 24.0f;
        constexpr float fieldH = 184.0f;
        constexpr float gap    = 12.0f;
        float y                = previewTop;
        for (ReviewOutputCard& card : _reviewOutputCards)
        {
            card.label->SetBounds(D2D1::RectF(left, y, cardRight, y + labelH));
            y += labelH + 4.0f;
            card.field->SetBounds(D2D1::RectF(left, y, cardRight, y + fieldH));
            y += fieldH + gap;
        }
        _reviewOutputScroll->SetContentHeight(std::max(0.0f, y - gap - previewTop));
    }

    HINSTANCE _instance = nullptr;
    RedConfigure::RedConfigureSession& _session;
    std::filesystem::path _workspaceRoot;
    InventoryGridModel _inventoryModel;
    LocalizationReviewGridModel _localizationReviewModel;
    ThemeColorGridModel _themeColorModel;
    bool _syncing                              = false;
    bool _lastLoadSucceeded                    = false;
    size_t _selectedPage                       = 0u;
    size_t _selectedReviewRow                  = 0u;
    bool _localizationReviewFiltersInitialized = false;
    RedConfigure::LocalizationReviewViewOptions _localizationReviewViewOptions;
    std::vector<size_t> _localizationReviewViewRows;
    std::wstring _selectedReviewCulture;
    std::vector<std::wstring> _themeColorKeys;
    std::vector<std::wstring> _filteredThemeColorKeys;
    std::wstring _activeThemeColorKey;
    std::wstring _previousThemeColorKey;
    std::vector<std::wstring> _recentThemeColors;
    RedConfigure::Workflow::LanguageColumnModel _languageColumns;
    RedConfigure::Workflow::ValidationSummary _validationSummary;
    LocalizationPagePresenter _localizationPresenter;
    ThemesPagePresenter _themesPresenter;
    bool _validationDrawerExpanded = false;
    bool _warningsAcknowledged = false;

    std::vector<Button*> _navButtons;
    Label* _titleLabel       = nullptr;
    Label* _descriptionLabel = nullptr;
    Label* _scopeLabel       = nullptr;
    Label* _statusLabel      = nullptr;
    TextField* _commandSearchField = nullptr;
    Button* _validateButton = nullptr;
    Button* _undoButton = nullptr;
    Button* _redoButton = nullptr;
    Button* _reviewExportButton = nullptr;
    Button* _validationDrawerButton = nullptr;
    Label* _validationDrawerLabel = nullptr;

    Label* _workspaceRootLabel               = nullptr;
    TextField* _workspaceRootField           = nullptr;
    Label* _cultureLabel                     = nullptr;
    ComboBox* _cultureCombo                  = nullptr;
    Button* _reloadButton                    = nullptr;
    Label* _ownerCountLabel                  = nullptr;
    Label* _themeFileCountLabel              = nullptr;
    Label* _scanErrorCountLabel              = nullptr;
    Label* _ownerSelectorLabel               = nullptr;
    ComboBox* _ownerCombo                    = nullptr;
    Label* _ownerFilterLabel                 = nullptr;
    TagPicker* _ownerFilterPicker            = nullptr;
    Label* _languageFilterLabel              = nullptr;
    TagPicker* _languageFilterPicker         = nullptr;
    Label* _activeOwnerLabel                 = nullptr;
    Label* _translationCountLabel            = nullptr;
    Label* _localizationSearchLabel          = nullptr;
    TextField* _localizationSearchField      = nullptr;
    Label* _localizationIdFilterLabel        = nullptr;
    TextField* _localizationIdFilterField    = nullptr;
    Label* _localizationStatusFilterLabel    = nullptr;
    ComboBox* _localizationStatusFilterCombo = nullptr;

    ScrollPanel* _localizationPageScroll = nullptr;
    Panel* _localizationPageContent      = nullptr;

    Grid* _inventoryGrid             = nullptr;
    Grid* _translationGrid           = nullptr;
    Label* _sourceTextLabel          = nullptr;
    TextField* _sourceTextField      = nullptr;
    Label* _targetTextLabel          = nullptr;
    ScrollPanel* _targetEditorsPanel = nullptr;
    std::vector<TargetEditor> _targetEditors;
    Label* _validationLabel = nullptr;
    Button* _previousProblemButton = nullptr;
    Button* _nextProblemButton = nullptr;
    Button* _pasteMatrixButton = nullptr;
    Button* _pinLanguageButton = nullptr;
    Button* _moveLanguageLeftButton = nullptr;
    Button* _moveLanguageRightButton = nullptr;
    Button* _removeLanguageButton = nullptr;
    ComboBox* _localizationBatchCombo = nullptr;
    Button* _applyLocalizationBatchButton = nullptr;
    Label* _localizationExampleLabel = nullptr;
    LocalizationExampleControl* _localizationExample = nullptr;
    Button* _exportRcButton = nullptr;

    Label* _themeSelectorLabel = nullptr;
    ComboBox* _themeCombo      = nullptr;
    Label* _themeNameLabel     = nullptr;
    Label* _themePathLabel     = nullptr;
    Label* _themeErrorLabel    = nullptr;
    TextField* _themeImportPathField = nullptr;
    Button* _themeImportButton = nullptr;
    Button* _themeDuplicateButton = nullptr;
    Button* _themeResetButton = nullptr;

    Label* _colorKeyFilterLabel        = nullptr;
    TextField* _colorKeyFilterField    = nullptr;
    Label* _colorKeyLabel              = nullptr;
    ComboBox* _colorKeyCombo           = nullptr;
    Grid* _themeColorGrid              = nullptr;
    Label* _previousColorLabel         = nullptr;
    ColorSwatch* _previousColorSwatch  = nullptr;
    Label* _colorValueLabel            = nullptr;
    ColorSwatch* _colorSwatch          = nullptr;
    ComboBox* _colorValueCombo         = nullptr;
    Label* _themeColorStatusLabel      = nullptr;
    Label* _themeExpressionHelpLabel   = nullptr;
    Label* _paletteNameLabel           = nullptr;
    TextField* _paletteNameField       = nullptr;
    Button* _addPaletteButton          = nullptr;
    Button* _renamePaletteButton       = nullptr;
    Label* _previewSeedLabel           = nullptr;
    ComboBox* _previewSeedCombo        = nullptr;
    Button* _darkenGroupButton         = nullptr;
    Button* _blendAccentGroupButton    = nullptr;
    Button* _resetColorButton          = nullptr;
    ScrollPanel* _themePreviewScroll   = nullptr;
    ThemeExampleControl* _themePreview = nullptr;
    Label* _themeSceneLabel = nullptr;
    ComboBox* _themeSceneCombo = nullptr;
    Label* _themeRecipeLabel = nullptr;
    ComboBox* _themeRecipeCombo = nullptr;
    Button* _applyThemeRecipeButton = nullptr;
    Slider* _themeAlphaSlider = nullptr;
    Button* _copyEffectiveButton = nullptr;
    Button* _copyOverrideButton = nullptr;
    Button* _exportThemeButton         = nullptr;

    struct ReviewOutputCard
    {
        Label* label     = nullptr;
        TextField* field = nullptr;
    };

    Label* _defaultRcPathLabel        = nullptr;
    Label* _defaultThemePathLabel     = nullptr;
    ScrollPanel* _reviewOutputScroll  = nullptr;
    std::vector<ReviewOutputCard> _reviewOutputCards;
};

RedConfigureRootCreateResult CreateRedConfigureRoot(HINSTANCE instance, RedConfigureSession& session, std::filesystem::path initialRoot)
{
    auto root                      = std::make_unique<RedConfigureRoot>(instance, session, std::move(initialRoot));
    auto* controller               = static_cast<RedConfigureRootController*>(root.get());
    std::unique_ptr<Panel> control = std::move(root);
    RedConfigureRootCreateResult result;
    result.control    = std::move(control);
    result.controller = controller;
    return result;
}

} // namespace RedConfigure::Ui
