#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "RedConfigureRoot.h"

#include "Helpers.h"
#include "RedConfigureApp.h"
#include "RedConfigureGridModels.h"
#include "RedConfigureSession.h"
#include "RedConfigureThemeExampleControl.h"
#include "RedConfigureUiHelpers.h"
#include "SettingsStore.h"
#include "resource.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <filesystem>
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
    L"window.text",
    L"window.subduedText",
    L"navigation.background",
    L"navigation.text",
    L"navigation.accent",
    L"menu.background",
    L"menu.text",
    L"menu.selectionBackground",
    L"menu.selectionText",
    L"menu.border",
    L"folderView.background",
    L"folderView.itemForeground",
    L"folderView.itemBackgroundHovered",
    L"folderView.itemBackgroundSelected",
    L"folderView.itemForegroundSelected",
    L"folderView.warningForeground",
    L"dialog.background",
    L"dialog.text",
    L"dialog.buttonBackground",
    L"dialog.buttonText",
    L"progress.background",
    L"progress.fill",
    L"diff.addedBackground",
    L"diff.removedBackground",
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

[[nodiscard]] uint32_t Channel(uint32_t argb, uint32_t shift) noexcept
{
    return (argb >> shift) & 0xFFu;
}

[[nodiscard]] uint32_t PackArgb(uint32_t a, uint32_t r, uint32_t g, uint32_t b) noexcept
{
    return ((a & 0xFFu) << 24u) | ((r & 0xFFu) << 16u) | ((g & 0xFFu) << 8u) | (b & 0xFFu);
}

[[nodiscard]] uint32_t MixChannel(uint32_t from, uint32_t to, double amount) noexcept
{
    const double mixed = static_cast<double>(from) + ((static_cast<double>(to) - static_cast<double>(from)) * amount);
    return static_cast<uint32_t>(std::clamp(mixed, 0.0, 255.0) + 0.5);
}

[[nodiscard]] uint32_t MixArgb(uint32_t from, uint32_t to, double amount) noexcept
{
    return PackArgb(MixChannel(Channel(from, 24u), Channel(to, 24u), amount),
                    MixChannel(Channel(from, 16u), Channel(to, 16u), amount),
                    MixChannel(Channel(from, 8u), Channel(to, 8u), amount),
                    MixChannel(Channel(from, 0u), Channel(to, 0u), amount));
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

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text)
{
    if (text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return {};
    }

    const int required = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
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
        _workspaceRoot       = RedConfigure::ResolveWorkspaceRootForLaunchPath(std::filesystem::path(std::wstring(_workspaceRootField->GetText())));
        std::wstring culture = !_localizationReviewViewOptions.visibleCultureNames.empty()
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
            _localizationReviewFiltersInitialized = true;
            if (std::find(_localizationReviewViewOptions.visibleCultureNames.begin(),
                          _localizationReviewViewOptions.visibleCultureNames.end(),
                          _selectedReviewCulture) == _localizationReviewViewOptions.visibleCultureNames.end())
            {
                _selectedReviewCulture = _localizationReviewViewOptions.visibleCultureNames.empty() ? std::wstring{}
                                                                                                    : _localizationReviewViewOptions.visibleCultureNames.front();
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
        _activeOwnerLabel        = _localizationPageContent->AddChild<Label>();
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
        _targetTextLabel = _localizationPageContent->AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_TARGET_TEXT));
        _targetEditorsPanel = _localizationPageContent->AddChild<ScrollPanel>();
        _targetEditorsPanel->SetScrollStepDip(kLocalizationGridRowHeightDip);
        _validationLabel = _localizationPageContent->AddChild<Label>();
        _validationLabel->SetFontRole(FontRole::BodyStrong);
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
        _themeExpressionHelpLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_EXPRESSION_EXAMPLES));
        _themeExpressionHelpLabel->SetFontRole(FontRole::Small);
        _themeExpressionHelpLabel->SetMultiline(true);
        _themePreviewScroll = AddChild<ScrollPanel>();
        _themePreviewScroll->SetScrollStepDip(56.0f);
        _themePreview = _themePreviewScroll->AddChild<ThemeExampleControl>(_instance);
        _themePreview->SetModel(&_session.GetThemePreviewModel());
        _themePreview->SetOnTokenSelected([this](std::wstring_view key) { SelectThemeColorKey(key); });
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
        _rcPreviewLabel        = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_RC_PREVIEW));
        _rcPreviewField        = AddChild<TextField>();
        _rcPreviewField->SetReadOnly(true);
        _rcPreviewField->SetMultiline(true);
        _themeJsonPreviewLabel = AddChild<Label>(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_PREVIEW));
        _themeJsonPreviewField = AddChild<TextField>();
        _themeJsonPreviewField->SetReadOnly(true);
        _themeJsonPreviewField->SetMultiline(true);
    }

    void SetPage(size_t pageIndex)
    {
        const auto pages = RedConfigure::GetPageDefinitions();
        if (pageIndex >= pages.size())
        {
            return;
        }

        _selectedPage = pageIndex;
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
        _workspaceRootField->SetText(_workspaceRoot.wstring());
        _ownerCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_RESOURCE_OWNERS) + L": " +
                                  std::to_wstring(_session.GetWorkspace().resourceOwners.size()));
        _themeFileCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_FILES) + L": " +
                                      std::to_wstring(_session.GetThemeCatalog().themes.size()));
        _scanErrorCountLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SCAN_ERRORS) + L": " +
                                      std::to_wstring(_session.GetWorkspace().errors.size()));
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
        _syncing = false;

        if (auto* host = GetHost())
        {
            host->Invalidate();
        }
    }

    void RebuildTranslationView(bool preserveSelection)
    {
        SyncLocalizationReviewFilterDefaults();
        _localizationReviewViewRows = RedConfigure::BuildLocalizationReviewView(_session.GetLocalizationReviewRows(), _localizationReviewViewOptions);
        _localizationReviewModel.SetVisibleCultures(_localizationReviewViewOptions.visibleCultureNames);
        _localizationReviewModel.SetViewRows(_localizationReviewViewRows);
        _translationGrid->NotifyDataChanged();
        const size_t totalTranslations = _session.GetLocalizationReviewRows().size();
        std::wstring countText = LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_TRANSLATION_COUNT) + L": " + std::to_wstring(_localizationReviewViewRows.size());
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

        if (std::find(_localizationReviewViewOptions.visibleCultureNames.begin(),
                      _localizationReviewViewOptions.visibleCultureNames.end(),
                      cultureName) == _localizationReviewViewOptions.visibleCultureNames.end())
        {
            _localizationReviewViewOptions.visibleCultureNames.push_back(cultureName);
        }

        _selectedReviewCulture                 = std::move(cultureName);
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
            _localizationReviewFiltersInitialized = true;
        }

        if (_selectedReviewCulture.empty() ||
            std::find(_localizationReviewViewOptions.visibleCultureNames.begin(), _localizationReviewViewOptions.visibleCultureNames.end(), _selectedReviewCulture) ==
                _localizationReviewViewOptions.visibleCultureNames.end())
        {
            _selectedReviewCulture = _localizationReviewViewOptions.visibleCultureNames.empty() ? std::wstring{} : _localizationReviewViewOptions.visibleCultureNames.front();
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
            themeItems.push_back(ComboBox::Item{.value = fallback.id, .display = fallback.name});
        }
        else
        {
            themeItems.reserve(catalog.themes.size());
            for (const auto& theme : catalog.themes)
            {
                themeItems.push_back(ComboBox::Item{.value = theme.definition.id, .display = theme.definition.name});
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
        keys.reserve(kThemePreviewColorKeys.size() + _session.GetThemePreviewModel().GetTheme().colors.size());
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
        std::wstring rcPreview;
        std::vector<RedConfigure::LocalizationExportPreview> reviewPreviews;
        if (! _session.GetLocalizationReviewRows().empty() && SUCCEEDED(_session.BuildLocalizationReviewExportPreviews(reviewPreviews)))
        {
            for (const RedConfigure::LocalizationExportPreview& preview : reviewPreviews)
            {
                if (! rcPreview.empty())
                {
                    rcPreview += L"\r\n\r\n";
                }
                rcPreview += L"// " + preview.path.wstring() + L"\r\n";
                rcPreview += preview.text;
            }
            _rcPreviewField->SetText(std::move(rcPreview));
        }
        else if (SUCCEEDED(_session.BuildLocalizationExportText(rcPreview)))
        {
            _rcPreviewField->SetText(std::move(rcPreview));
        }
        else
        {
            _rcPreviewField->SetText({});
        }

        std::string themePreview;
        if (SUCCEEDED(_session.BuildThemeExportText(themePreview)))
        {
            _themeJsonPreviewField->SetText(Utf16FromUtf8(themePreview));
        }
        else
        {
            _themeJsonPreviewField->SetText({});
        }
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
        const std::vector<std::wstring> suggestions = RedConfigure::BuildThemeColorSuggestions(key, previousKey, color);
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
        _themeColorStatusLabel->SetText({});
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
            const std::optional<uint32_t> color = _session.GetThemePreviewModel().GetEffectiveColor(key);
            _colorSwatch->SetSwatchValue(color);
            _themeColorStatusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_THEME_COLOR_VALID));
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

        const std::optional<uint32_t> accent = _session.GetThemePreviewModel().GetEffectiveColor(L"app.accent");
        bool changed                         = false;
        for (const std::wstring& key : _themeColorKeys)
        {
            if (ThemeKeyGroup(key) != selectedGroup)
            {
                continue;
            }

            const std::optional<uint32_t> color = _session.GetThemePreviewModel().GetEffectiveColor(key);
            if (! color)
            {
                continue;
            }

            uint32_t transformed = color.value();
            switch (transform)
            {
                case ThemeGroupTransform::Darken: transformed = MixArgb(color.value(), 0xFF000000u, 0.10); break;
                case ThemeGroupTransform::BlendAccent:
                    if (! accent || key == L"app.accent")
                    {
                        continue;
                    }
                    transformed = MixArgb(color.value(), accent.value(), 0.16);
                    break;
                default: break;
            }

            changed = _session.UpdateThemeColor(key, Common::Settings::FormatColor(transformed)) || changed;
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

        _session.GetThemePreviewModel().ResetOverride(selectedKey);
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
        if (_selectedPage == 0u)
        {
            text += LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_RESOURCE_OWNERS) + L" " + std::to_wstring(_session.GetWorkspace().resourceOwners.size());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_THEME_FILES) + L" " + std::to_wstring(_session.GetThemeCatalog().themes.size());
            text += L" | " + LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_SCAN_ERRORS) + L" " + std::to_wstring(_session.GetWorkspace().errors.size());
        }
        else if (_selectedPage == 1u)
        {
            text += LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_OWNERS) + L" " +
                    std::to_wstring(_localizationReviewViewOptions.visibleOwnerNames.size()) + L"/" +
                    std::to_wstring(_session.GetWorkspace().resourceOwners.size());
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

    [[nodiscard]] ThemePalette ResolvePalette() const noexcept
    {
        return GetHost() ? GetHost()->GetTheme() : MakeDefaultThemePalette(false);
    }

    void ExportLocalization()
    {
        if (_session.GetLocalizationReviewRows().empty())
        {
            const std::filesystem::path path = _session.GetDefaultLocalizationExportPath();
            const HRESULT hr                 = _session.ExportLocalization(path);
            _statusLabel->SetText(LoadAppString(_instance, SUCCEEDED(hr) ? IDS_REDCONFIGURE_STATUS_EXPORT_RC_DONE : IDS_REDCONFIGURE_STATUS_EXPORT_FAILED) +
                                  L": " + path.wstring());
            return;
        }

        size_t exportedFileCount               = 0u;
        const std::filesystem::path outputRoot = _session.GetDefaultLocalizationExportPath().parent_path();
        const HRESULT hr                       = _session.ExportLocalizationReview(outputRoot, &exportedFileCount);

        if (SUCCEEDED(hr))
        {
            _statusLabel->SetText(FormatStringResource(_instance, IDS_REDCONFIGURE_STATUS_REVIEW_EXPORT_DONE, exportedFileCount) + L": " + outputRoot.wstring());
        }
        else
        {
            _statusLabel->SetText(LoadAppString(_instance, IDS_REDCONFIGURE_STATUS_EXPORT_FAILED) + L": " + outputRoot.wstring());
        }
    }

    void ExportTheme()
    {
        const std::filesystem::path path = _session.GetDefaultThemeExportPath();
        const HRESULT hr                 = _session.ExportTheme(path);
        _statusLabel->SetText(LoadAppString(_instance, SUCCEEDED(hr) ? IDS_REDCONFIGURE_STATUS_EXPORT_THEME_DONE : IDS_REDCONFIGURE_STATUS_EXPORT_FAILED) +
                              L": " + path.wstring());
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
        Control* retiredLocalizationControls[] = {_ownerSelectorLabel, _ownerCombo, _cultureLabel, _cultureCombo};
        Control* localizationControls[]  = {_localizationPageScroll,
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
                                            _validationLabel};
        Control* themeDesignerControls[] = {
            _themeSelectorLabel,  _themeCombo,         _themeNameLabel,        _themePathLabel,           _themeErrorLabel,    _colorKeyFilterLabel,
            _colorKeyFilterField, _colorKeyLabel,      _themeColorGrid,        _colorValueLabel,          _previousColorLabel, _previousColorSwatch,
            _colorSwatch,         _colorValueCombo,    _themeColorStatusLabel, _themeExpressionHelpLabel, _darkenGroupButton,  _blendAccentGroupButton,
            _resetColorButton,    _themePreviewScroll, _themePreview};
        Control* exportControls[] = {_defaultRcPathLabel,
                                     _defaultThemePathLabel,
                                     _exportRcButton,
                                     _exportThemeButton,
                                     _rcPreviewLabel,
                                     _rcPreviewField,
                                     _themeJsonPreviewLabel,
                                     _themeJsonPreviewField};

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
        const float bodyTop      = contentTop + 68.0f;
        const float statusTop    = std::max(bodyTop, height - margin - 42.0f);
        const float bodyBottom   = std::max(bodyTop, statusTop - gap);

        float navY = margin;
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
        const auto tagPickerHeight = [](const TagPicker* picker, float widthDip) noexcept
        {
            return picker ? picker->GetPreferredHeightDip(widthDip) : 32.0f;
        };
        float y                       = top;

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
        const float minContentBottom         = y + minGridH + 8.0f + editorPreferredH + 12.0f;
        float layoutBottom                   = std::max(bottom, top + kLocalizationPageMinContentHeightDip);
        layoutBottom                         = std::max(layoutBottom, minContentBottom);
        const float availableH               = std::max(0.0f, layoutBottom - y);
        const float gridPreferredH           = availableH - editorPreferredH - 8.0f;
        const float gridMaxH                 = std::max(minGridH, availableH - editorMinH - 8.0f);
        const float gridMinH                 = minGridH;
        const float gridH                    = std::clamp(gridPreferredH, gridMinH, gridMaxH);

        const float gridTop      = y;
        const float gridBottom   = gridTop + gridH;
        const float editorTop    = gridBottom + 8.0f;
        const float editorBottom = std::max(editorTop, layoutBottom - 12.0f);

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

        const float contentHeight = std::max(viewportH, targetPanelBottom + 12.0f - top);
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
        const float scrollbarW    = needsScrollbar ? _targetEditorsPanel->GetScrollbarThickness() + 2.0f : 0.0f;
        const float contentRight  = std::max(bounds.left, bounds.right - scrollbarW);
        const float rowGap        = _targetEditors.size() > 1u ? 4.0f : 0.0f;
        const float contentWidth  = std::max(0.0f, contentRight - bounds.left);
        const float labelW        = TargetEditorLabelWidth(contentWidth);
        const float fieldLeft     = std::min(contentRight, bounds.left + labelW + 6.0f);
        const float fieldWidth    = std::max(0.0f, contentRight - fieldLeft);
        float y                   = bounds.top;
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

        const bool sideBySide   = contentW >= 820.0f;
        const float editorTop   = top + 128.0f;
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
        _themeColorStatusLabel->SetBounds(D2D1::RectF(left + labelW, y, editorRight, y + 22.0f));
        y += 24.0f;

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

        const D2D1_RECT_F previewBounds = D2D1::RectF(previewLeft, previewTop, right, bottom);
        _themePreviewScroll->SetBounds(previewBounds);
        const float previewContentHeight = std::max(kThemePreviewContentHeightDip, previewBounds.bottom - previewBounds.top);
        _themePreviewScroll->SetContentHeight(previewContentHeight);
        const float previewContentRight = std::max(previewBounds.left + 120.0f, previewBounds.right - _themePreviewScroll->GetScrollbarThickness() - 2.0f);
        _themePreview->SetBounds(D2D1::RectF(previewBounds.left, previewBounds.top, previewContentRight, previewBounds.top + previewContentHeight));
    }

    void LayoutExportPage(float left, float top, float right, float) noexcept
    {
        _defaultRcPathLabel->SetBounds(D2D1::RectF(left, top, right, top + 42.0f));
        _exportRcButton->SetBounds(D2D1::RectF(left, top + 48.0f, left + 140.0f, top + 80.0f));
        _defaultThemePathLabel->SetBounds(D2D1::RectF(left + 170.0f, top, right, top + 42.0f));
        _exportThemeButton->SetBounds(D2D1::RectF(left + 170.0f, top + 48.0f, left + 310.0f, top + 80.0f));

        const float previewTop  = top + 108.0f;
        const float gap         = 12.0f;
        const float mid         = left + ((right - left - gap) / 2.0f);
        const float fieldBottom = std::max(previewTop + 58.0f, GetBounds().bottom - 18.0f);
        _rcPreviewLabel->SetBounds(D2D1::RectF(left, previewTop, mid, previewTop + 24.0f));
        _rcPreviewField->SetBounds(D2D1::RectF(left, previewTop + 28.0f, mid, fieldBottom));
        _themeJsonPreviewLabel->SetBounds(D2D1::RectF(mid + gap, previewTop, right, previewTop + 24.0f));
        _themeJsonPreviewField->SetBounds(D2D1::RectF(mid + gap, previewTop + 28.0f, right, fieldBottom));
    }

    HINSTANCE _instance = nullptr;
    RedConfigure::RedConfigureSession& _session;
    std::filesystem::path _workspaceRoot;
    InventoryGridModel _inventoryModel;
    LocalizationReviewGridModel _localizationReviewModel;
    ThemeColorGridModel _themeColorModel;
    bool _syncing               = false;
    bool _lastLoadSucceeded     = false;
    size_t _selectedPage        = 0u;
    size_t _selectedReviewRow   = 0u;
    bool _localizationReviewFiltersInitialized = false;
    RedConfigure::LocalizationReviewViewOptions _localizationReviewViewOptions;
    std::vector<size_t> _localizationReviewViewRows;
    std::wstring _selectedReviewCulture;
    std::vector<std::wstring> _themeColorKeys;
    std::vector<std::wstring> _filteredThemeColorKeys;
    std::wstring _activeThemeColorKey;
    std::wstring _previousThemeColorKey;

    std::vector<Button*> _navButtons;
    Label* _titleLabel       = nullptr;
    Label* _descriptionLabel = nullptr;
    Label* _scopeLabel       = nullptr;
    Label* _statusLabel      = nullptr;

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

    Grid* _inventoryGrid              = nullptr;
    Grid* _translationGrid            = nullptr;
    Label* _sourceTextLabel           = nullptr;
    TextField* _sourceTextField       = nullptr;
    Label* _targetTextLabel           = nullptr;
    ScrollPanel* _targetEditorsPanel  = nullptr;
    std::vector<TargetEditor> _targetEditors;
    Label* _validationLabel           = nullptr;
    Button* _exportRcButton           = nullptr;

    Label* _themeSelectorLabel = nullptr;
    ComboBox* _themeCombo      = nullptr;
    Label* _themeNameLabel     = nullptr;
    Label* _themePathLabel     = nullptr;
    Label* _themeErrorLabel    = nullptr;

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
    Button* _darkenGroupButton         = nullptr;
    Button* _blendAccentGroupButton    = nullptr;
    Button* _resetColorButton          = nullptr;
    ScrollPanel* _themePreviewScroll   = nullptr;
    ThemeExampleControl* _themePreview = nullptr;
    Button* _exportThemeButton         = nullptr;

    Label* _defaultRcPathLabel        = nullptr;
    Label* _defaultThemePathLabel     = nullptr;
    Label* _rcPreviewLabel            = nullptr;
    TextField* _rcPreviewField        = nullptr;
    Label* _themeJsonPreviewLabel     = nullptr;
    TextField* _themeJsonPreviewField = nullptr;
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
