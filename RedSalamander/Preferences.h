#pragma once

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "AppTheme.h"
#include "SettingsStore.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

enum class PrefCategory : int;

[[nodiscard]] bool ShowPreferencesDialog(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme);
[[nodiscard]] bool ShowPreferencesDialogPlugins(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme);
[[nodiscard]] bool ShowPreferencesDialogHotPaths(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme);
[[nodiscard]] bool ShowPreferencesDialogUserMenu(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme);

[[nodiscard]] HWND GetPreferencesDialogHandle() noexcept;
void UpdatePreferencesWindowsTheme(const AppTheme& theme) noexcept;

#ifdef ENABLE_TESTS
enum class PreferencesKeyboardDebugFocusTarget : uint8_t
{
    None = 0u,
    SearchField,
    ScopeCombo,
    ShortcutsGrid,
    AssignButton,
    RemoveButton,
    ResetButton,
    ImportButton,
    ExportButton,
};

enum class PreferencesPluginsDebugFocusTarget : uint8_t
{
    None = 0u,
    SearchField,
    MainList,
    ConfigureButton,
    TestButton,
    TestAllButton,
    CustomPathsList,
    CustomPathsAddButton,
    CustomPathsRemoveButton,
};

enum class PreferencesViewersDebugFocusTarget : uint8_t
{
    None = 0u,
    SearchField,
    MappingsGrid,
    ExtensionField,
    ViewerCombo,
    SaveButton,
    RemoveButton,
    ResetButton,
    ActionsGrid,
    ActionIdField,
    TestFileField,
};

enum class PreferencesGeneralDebugFocusTarget : uint8_t
{
    None = 0u,
    MenuBarToggle,
    FunctionBarToggle,
    LanguageCombo,
    CompactModeToggle,
    ReducedMotionCombo,
    WindowBackdropCombo,
    SplashScreenToggle,
};

enum class PreferencesPanesDebugFocusTarget : uint8_t
{
    None = 0u,
    LeftDisplayToggle,
    LeftDisplayCombo,
    LeftSortByCombo,
    LeftSortDirToggle,
    LeftSortDirCombo,
    LeftStatusBarToggle,
    RightDisplayToggle,
    RightDisplayCombo,
    RightSortByCombo,
    RightSortDirToggle,
    RightSortDirCombo,
    RightStatusBarToggle,
    ShowHiddenFilesToggle,
    ShowSystemFilesToggle,
    HistoryField,
};

enum class PreferencesHotPathsDebugFocusTarget : uint8_t
{
    None = 0u,
    FirstPathField,
    FirstBrowseButton,
    FirstLabelField,
    FirstShowInMenuToggle,
    SecondPathField,
    OpenPrefsToggle,
};

enum class PreferencesAdvancedDebugFocusTarget : uint8_t
{
    None = 0u,
    BypassHelloToggle,
    AllowInsecureTlsAutomationToggle,
    HelloTimeoutEdit,
    ToolbarToggle,
    LineNumbersToggle,
    AlwaysOnTopToggle,
    ShowIdsToggle,
    AutoScrollToggle,
    FilterPresetCombo,
    FilterMaskEdit,
    FilterTextToggle,
    DiagnosticsDebugToggle,
};

enum class PreferencesCompareDirectoriesDebugFocusTarget : uint8_t
{
    None = 0u,
    CompareSubdirectoriesToggle,
    CompareSizeToggle,
    CompareDateTimeToggle,
    CompareAttributesToggle,
    CompareContentToggle,
    ContentWorkersCombo,
    CompareSubdirAttributesToggle,
    SelectSubdirsOnlyInOnePaneToggle,
    KeepIdenticalItemsToggle,
    ShowIdenticalItemsToggle,
    IgnoreFilesToggle,
    IgnoreFilesEdit,
    IgnoreDirectoriesToggle,
    IgnoreDirectoriesEdit,
};

enum class PreferencesFileOperationsDebugFocusTarget : uint8_t
{
    None = 0u,
    PreCalcEnabledToggle,
    PreCalcWorkersCombo,
    BandwidthPresetCombo,
    CustomBandwidthEdit,
    BridgeBufferEdit,
};

enum class PreferencesThemesDebugFocusTarget : uint8_t
{
    None = 0u,
    ThemeCombo,
    NameField,
    BaseCombo,
    LoadFromFileButton,
    DuplicateButton,
    ResetButton,
    SaveButton,
    ApplyTemporarilyButton,
    SearchField,
    ColorsGrid,
    KeyField,
    ColorField,
    PickButton,
    SetButton,
    ClearButton,
};

enum class PreferencesShellDebugFocusTarget : uint8_t
{
    None = 0u,
    ResetAllButton,
    OkButton,
    CancelButton,
    ApplyButton,
};

struct PreferencesGridPointerDebugState
{
    uint64_t headerResizeDownCount = 0u;
    uint64_t resizeMoveCount       = 0u;
    bool resizeActive              = false;
    float lastResizeDeltaDip       = 0.0f;
    float lastResizeWidthDip       = 0.0f;
};

struct PreferencesKeyboardDebugSnapshot
{
    PrefCategory currentCategory          = static_cast<PrefCategory>(0);
    size_t keyboardListRowCount           = 0u;
    size_t keyboardListVisibleRowCount    = 0u;
    size_t keyboardListVisibleColumnCount = 0u;
    size_t keyboardListVisibleCellCount   = 0u;
    std::wstring keyboardSearchText;
    std::wstring keyboardSelectedCommandIdText;
    std::wstring keyboardSelectedChordText;
    PreferencesKeyboardDebugFocusTarget keyboardFocusTarget = PreferencesKeyboardDebugFocusTarget::None;
    bool keyboardCaptureActive                              = false;
    size_t visibleCurrentPageChildWindowCount               = 0u;
    uint64_t currentPageDxHostResizeFailureCount            = 0u;
};

struct PreferencesDebugSnapshot
{
    bool categoryTreeUsesDxUiHost         = false;
    bool shellUsesDxUiHost                = false;
    bool pageHostUsesDxUiHost             = false;
    size_t themesListRowCount             = 0u;
    size_t themesListVisibleRowCount      = 0u;
    size_t themesListVisibleColumnCount   = 0u;
    size_t themesListVisibleCellCount     = 0u;
    bool themesListHasVerticalScrollbar   = false;
    float themesListVerticalScrollDip     = 0.0f;
    uint64_t themesListRenderCount        = 0u;
    uint64_t themesListResizeCount        = 0u;
    uint64_t themesListResizeFailureCount = 0u;
    std::wstring themesSearchText;
    std::wstring themesSelectedThemeIdText;
    std::wstring themesSelectedColorKeyText;
    std::wstring themesColorText;
    bool themesSelectedColorOverrideActive                                      = false;
    PreferencesGeneralDebugFocusTarget generalFocusTarget                       = PreferencesGeneralDebugFocusTarget::None;
    PreferencesPanesDebugFocusTarget panesFocusTarget                           = PreferencesPanesDebugFocusTarget::None;
    PreferencesHotPathsDebugFocusTarget hotPathsFocusTarget                     = PreferencesHotPathsDebugFocusTarget::None;
    PreferencesAdvancedDebugFocusTarget advancedFocusTarget                     = PreferencesAdvancedDebugFocusTarget::None;
    PreferencesCompareDirectoriesDebugFocusTarget compareDirectoriesFocusTarget = PreferencesCompareDirectoriesDebugFocusTarget::None;
    PreferencesFileOperationsDebugFocusTarget fileOperationsFocusTarget         = PreferencesFileOperationsDebugFocusTarget::None;
    PreferencesThemesDebugFocusTarget themesFocusTarget                         = PreferencesThemesDebugFocusTarget::None;
    PreferencesShellDebugFocusTarget shellFocusTarget                           = PreferencesShellDebugFocusTarget::None;
    bool previewApplied                                                         = false;
    bool themeDark                                                              = false;
    bool themeHighContrast                                                      = false;
    bool themeRainbow                                                           = false;
    bool themeCompactMode                                                       = false;
    bool themeReducedMotion                                                     = false;
    std::wstring generalUiLanguage;
    bool generalUsesDxUiTypographyContext  = false;
    bool generalUsesDxUiTypographyMetrics  = false;
    bool panesUsesDxUiTypographyContext    = false;
    bool panesUsesDxUiTypographyMetrics    = false;
    bool viewersUsesDxUiTypographyContext  = false;
    bool viewersUsesDxUiTypographyMetrics  = false;
    uint32_t themeOverlayBackgroundArgb    = 0u;
    AppBackdropType themePrimaryBackdrop   = AppBackdropType::None;
    AppBackdropType themeToolBackdrop      = AppBackdropType::None;
    float generalCompactToggleHeightDip    = 0.0f;
    size_t viewersListRowCount             = 0u;
    size_t viewersListVisibleRowCount      = 0u;
    size_t viewersListVisibleColumnCount   = 0u;
    size_t viewersListVisibleCellCount     = 0u;
    bool viewersListHasVerticalScrollbar   = false;
    float viewersListVerticalScrollDip     = 0.0f;
    uint64_t viewersListRenderCount        = 0u;
    uint64_t viewersListResizeCount        = 0u;
    uint64_t viewersListResizeFailureCount = 0u;
    size_t viewersActionCount              = 0u;
    size_t viewersActionRowCount           = 0u;
    std::wstring viewersPrimaryActionIdText;
    std::wstring viewersAlternateActionIdText;
    std::wstring viewersPreviewActionIdText;
    std::wstring viewersPreviewReasonText;
    std::wstring viewersSearchText;
    std::wstring viewersSelectedExtensionText;
    PreferencesViewersDebugFocusTarget viewersFocusTarget = PreferencesViewersDebugFocusTarget::None;
    size_t editorsActionCount                            = 0u;
    size_t editorsAssociationRowCount                    = 0u;
    size_t editorsActionRowCount                         = 0u;
    std::wstring editorsPrimaryActionIdText;
    std::wstring editorsAlternateActionIdText;
    std::wstring editorsEditNewActionIdText;
    std::wstring editorsPreviewActionIdText;
    std::wstring editorsPreviewReasonText;
    size_t userMenuActionCount                           = 0u;
    size_t keyboardListRowCount                           = 0u;
    size_t keyboardListVisibleRowCount                    = 0u;
    size_t keyboardListVisibleColumnCount                 = 0u;
    size_t keyboardListVisibleCellCount                   = 0u;
    bool keyboardListHasVerticalScrollbar                 = false;
    float keyboardListVerticalScrollDip                   = 0.0f;
    uint64_t keyboardListRenderCount                      = 0u;
    uint64_t keyboardListResizeCount                      = 0u;
    uint64_t keyboardListResizeFailureCount               = 0u;
    std::wstring keyboardSearchText;
    std::wstring keyboardHintText;
    PreferencesKeyboardDebugFocusTarget keyboardFocusTarget = PreferencesKeyboardDebugFocusTarget::None;
    bool keyboardCaptureActive                              = false;
    size_t pluginsMainListRowCount                          = 0u;
    size_t pluginsMainListVisibleRowCount                   = 0u;
    size_t pluginsMainListVisibleColumnCount                = 0u;
    size_t pluginsMainListVisibleCellCount                  = 0u;
    bool pluginsMainListHasVerticalScrollbar                = false;
    float pluginsMainListVerticalScrollDip                  = 0.0f;
    uint64_t pluginsMainListRenderCount                     = 0u;
    uint64_t pluginsMainListResizeCount                     = 0u;
    uint64_t pluginsMainListResizeFailureCount              = 0u;
    size_t pluginsCustomPathsListRowCount                   = 0u;
    size_t pluginsCustomPathsListVisibleRowCount            = 0u;
    size_t pluginsCustomPathsListVisibleColumnCount         = 0u;
    size_t pluginsCustomPathsListVisibleCellCount           = 0u;
    bool pluginsCustomPathsListHasVerticalScrollbar         = false;
    float pluginsCustomPathsListVerticalScrollDip           = 0.0f;
    uint64_t pluginsCustomPathsListRenderCount              = 0u;
    uint64_t pluginsCustomPathsListResizeCount              = 0u;
    uint64_t pluginsCustomPathsListResizeFailureCount       = 0u;
    bool pluginsCustomPathsEmptyPlaceholderVisible          = false;
    std::wstring pluginsSearchText;
    std::wstring pluginsSelectedPluginIdText;
    std::wstring pluginsSelectedCustomPathText;
    std::wstring pluginsStatusTitleText;
    std::wstring pluginsStatusBodyText;
    std::wstring pluginsDetailsConfigErrorText;
    std::wstring pluginsDetailsConfigEmptyStateText;
    size_t pluginsDetailsConfigFieldCount                 = 0u;
    size_t pluginsDetailsVisibleConfigFieldCount          = 0u;
    size_t pluginsDetailsConfigDxChildCount               = 0u;
    bool pluginsDetailsConfigDxPanelVisible               = false;
    PreferencesPluginsDebugFocusTarget pluginsFocusTarget = PreferencesPluginsDebugFocusTarget::None;
    bool categoryTreeFocused                              = false;
    bool pluginItemSelected                               = false;
    bool pluginsDetailsActive                             = false;
    bool pluginsExpanded                                  = false;
    bool generalPaneVisible                               = false;
    bool pluginsPaneVisible                               = false;
    size_t pluginsTreeChildCount                          = 0u;
    size_t visibleChildWindowCount                        = 0u;
    size_t createdPaneWindowCount                         = 0u;
    size_t visiblePaneWindowCount                         = 0u;
    size_t visibleCurrentPageChildWindowCount             = 0u;
    size_t currentPageCardCount                           = 0u;
    size_t currentPageRenderedDxHostCount                 = 0u;
    size_t currentPageDxHostResizeFailureCount            = 0u;
    uint64_t currentPageDxHostRenderCountTotal            = 0u;
    size_t visibleShellRenderedDxHostCount                = 0u;
    size_t shellDxHostResizeFailureCount                  = 0u;
    uint64_t pageHostDxHostRenderCount                    = 0u;
    uint64_t categoryTreeDxHostRenderCount                = 0u;
    bool categoryTreeDxHostHasResizeFailures              = false;
    bool categoryTreeHasVerticalScrollbar                 = false;
    bool categoryTreeHasSelectedItem                      = false;
    size_t categoryTreeFirstVisibleIndex                  = 0u;
    size_t categoryTreeSelectedVisibleIndex               = 0u;
    float categoryTreeVerticalScrollDip                   = 0.0f;
    bool pageHostShowsVerticalScroll                      = false;
    int pageScrollY                                       = 0;
    int pageScrollMaxY                                    = 0;
    size_t visibleLegacyTreeViewCount                     = 0u;
    size_t visibleLegacyShellStaticCount                  = 0u;
    size_t visibleLegacyFooterButtonCount                 = 0u;
    size_t createdLegacyPluginsListBridgeCount            = 0u;
    size_t createdLegacyPluginsButtonBridgeCount          = 0u;
    size_t createdLegacyPluginsInputBridgeCount           = 0u;
    size_t createdLegacyPluginsConfigStaticBridgeCount    = 0u;
    size_t createdLegacyPluginsConfigInputBridgeCount     = 0u;
    PrefCategory currentCategory                          = static_cast<PrefCategory>(0);
    std::wstring pageTitle;
    std::wstring pageDescription;
};

[[nodiscard]] bool DebugGetPreferencesDialogSnapshot(PreferencesDebugSnapshot& out) noexcept;
[[nodiscard]] HWND DebugGetPreferencesActivePageHandle() noexcept;
[[nodiscard]] HWND DebugGetPreferencesActivePageDxHostHandle() noexcept;
[[nodiscard]] HWND DebugGetPreferencesShellHostHandle() noexcept;
[[nodiscard]] bool DebugSelectPreferencesCategory(PrefCategory category) noexcept;
[[nodiscard]] bool DebugSelectPreferencesPluginsTreeChild(size_t childIndex) noexcept;
[[nodiscard]] bool DebugScrollPreferencesCategoryTreeByWheelDelta(int wheelDelta) noexcept;
[[nodiscard]] bool DebugScrollPreferencesCategoryTreeByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugSelectPreferencesPluginsMainListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugClickPreferencesPluginsMainListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugFindPreferencesPluginsToggleableMainListRow(size_t& outRowIndex, bool& outEnabled) noexcept;
[[nodiscard]] bool DebugFindPreferencesPluginsLoadableMainListRow(size_t& outRowIndex) noexcept;
[[nodiscard]] bool DebugTogglePreferencesPluginsMainListCheckbox(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugGetPreferencesPluginsMainListRowEnabled(size_t rowIndex, bool& outEnabled) noexcept;
[[nodiscard]] bool DebugGetPreferencesPluginsMainListCheckboxClientRect(size_t rowIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetPreferencesPluginsMainListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugSelectPreferencesPluginsCustomPathsListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugClearPreferencesPluginsCustomPaths() noexcept;
[[nodiscard]] bool DebugGetPreferencesPluginsCustomPathsListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugFocusPreferencesPluginsMainList() noexcept;
[[nodiscard]] bool DebugFocusPreferencesPluginsSearchField() noexcept;
[[nodiscard]] bool DebugSetPreferencesPluginsNextCustomPathBrowsePath(std::wstring_view path) noexcept;
[[nodiscard]] bool DebugCancelPreferencesPluginsNextCustomPathBrowse() noexcept;
[[nodiscard]] bool DebugSetPreferencesHotPathsNextBrowsePath(std::wstring_view path) noexcept;
[[nodiscard]] bool DebugCancelPreferencesHotPathsNextBrowse() noexcept;
[[nodiscard]] bool DebugSelectPreferencesKeyboardListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugFindPreferencesKeyboardListRowByCommandId(std::wstring_view commandId, size_t& outRowIndex) noexcept;
[[nodiscard]] bool DebugGetPreferencesKeyboardVisibleRowChordByCommandId(std::wstring_view commandId, std::wstring& outChordText) noexcept;
[[nodiscard]] bool DebugGetPreferencesKeyboardListRowClientRect(size_t rowIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetPreferencesKeyboardListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugHitTestPreferencesKeyboardListClientPoint(
    POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) noexcept;
[[nodiscard]] bool DebugGetPreferencesKeyboardListPointerState(PreferencesGridPointerDebugState& outState) noexcept;
[[nodiscard]] bool DebugSelectPreferencesViewersListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugGetPreferencesViewersListRowClientRect(size_t rowIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetPreferencesViewersListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugSelectPreferencesThemesListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugGetPreferencesThemesListRowClientRect(size_t rowIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetPreferencesThemesListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugSetPreferencesViewersSearchText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSelectPreferencesViewersDefaultAction(bool alternate, std::wstring_view actionId) noexcept;
[[nodiscard]] bool DebugSelectPreferencesEditorsDefaultAction(bool alternate, std::wstring_view actionId) noexcept;
[[nodiscard]] bool DebugSelectPreferencesEditorsDefaultEditNewAction(std::wstring_view actionId) noexcept;
[[nodiscard]] bool DebugSetPreferencesKeyboardSearchText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetPreferencesKeyboardFunctionBarScope() noexcept;
[[nodiscard]] bool DebugCapturePreferencesKeyboardShortcut(uint32_t vk, uint32_t modifiers = 0) noexcept;
[[nodiscard]] bool DebugGetPreferencesKeyboardSnapshot(PreferencesKeyboardDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugFocusPreferencesViewersSearchField() noexcept;
[[nodiscard]] bool DebugFocusPreferencesKeyboardSearchField() noexcept;
[[nodiscard]] bool DebugFocusPreferencesGeneralMenuBarToggle() noexcept;
[[nodiscard]] bool DebugGetPreferencesGeneralMenuBarToggleChecked(bool& outChecked) noexcept;
[[nodiscard]] bool DebugSetPreferencesGeneralCompactMode(bool checked) noexcept;
[[nodiscard]] bool DebugSelectPreferencesGeneralLanguage(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugSelectPreferencesGeneralReducedMotion(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugSelectPreferencesGeneralWindowBackdrop(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugFocusPreferencesPanesLeftDisplayToggle() noexcept;
[[nodiscard]] bool DebugSelectPreferencesPanesLeftDisplay(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugFocusPreferencesPanesLeftStatusBarToggle() noexcept;
[[nodiscard]] bool DebugGetPreferencesPanesLeftStatusBarToggleChecked(bool& outChecked) noexcept;
[[nodiscard]] bool DebugFocusPreferencesHotPathsFirstPathField() noexcept;
[[nodiscard]] bool DebugGetPreferencesHotPathsFirstPathText(std::wstring& outText) noexcept;
[[nodiscard]] bool DebugSetPreferencesHotPathsFirstPathText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugFocusPreferencesHotPathsOpenPrefsToggle() noexcept;
[[nodiscard]] bool DebugGetPreferencesHotPathsOpenPrefsToggleChecked(bool& outChecked) noexcept;
[[nodiscard]] bool DebugFocusPreferencesAdvancedBypassHelloToggle() noexcept;
[[nodiscard]] bool DebugSelectPreferencesAdvancedFilterPreset(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugFocusPreferencesFileOperationsPreCalcEnabledToggle() noexcept;
[[nodiscard]] bool DebugGetPreferencesFileOperationsPreCalcEnabledToggleChecked(bool& outChecked) noexcept;
[[nodiscard]] bool DebugSelectPreferencesFileOperationsBandwidthPreset(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugSelectPreferencesCompareDirectoriesContentWorkers(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugSetPreferencesFileOperationsBridgeBufferText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugCancelPreferencesDialog() noexcept;
[[nodiscard]] bool DebugResetPreferencesToDefaults(bool confirm) noexcept;
[[nodiscard]] bool DebugFocusPreferencesCompareDirectoriesSubdirectoriesToggle() noexcept;
[[nodiscard]] bool DebugFocusPreferencesCompareDirectoriesTarget(PreferencesCompareDirectoriesDebugFocusTarget target) noexcept;
[[nodiscard]] bool DebugGetPreferencesCompareDirectoriesToggleChecked(PreferencesCompareDirectoriesDebugFocusTarget target, bool& outChecked) noexcept;
[[nodiscard]] bool DebugFocusPreferencesThemesSearchField() noexcept;
[[nodiscard]] bool DebugSetPreferencesKeyboardNextBrowsePath(std::wstring_view path) noexcept;
[[nodiscard]] bool DebugCancelPreferencesKeyboardNextBrowse() noexcept;
[[nodiscard]] bool DebugSetPreferencesPluginsSearchText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetPreferencesThemesSearchText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetPreferencesThemesNextBrowsePath(std::wstring_view path) noexcept;
[[nodiscard]] bool DebugCancelPreferencesThemesNextBrowse() noexcept;
[[nodiscard]] bool DebugScrollPreferencesPluginsMainListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugScrollPreferencesPluginsCustomPathsListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugScrollPreferencesKeyboardListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugScrollPreferencesViewersListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugScrollPreferencesThemesListByWheelDetents(int detents) noexcept;
#endif
