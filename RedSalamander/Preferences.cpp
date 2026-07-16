#include "Preferences.h"

#include "Preferences.Dialog.h"
#include "Preferences.Internal.h"

bool ShowPreferencesDialog(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme)
{
    return PreferencesDialog::Show(owner, appId, settings, theme, PrefCategory::General);
}

bool ShowPreferencesDialogPlugins(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme)
{
    return PreferencesDialog::Show(owner, appId, settings, theme, PrefCategory::Plugins);
}

bool ShowPreferencesDialogHotPaths(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme)
{
    return PreferencesDialog::Show(owner, appId, settings, theme, PrefCategory::HotPaths);
}

bool ShowPreferencesDialogUserMenu(HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme)
{
    return PreferencesDialog::Show(owner, appId, settings, theme, PrefCategory::UserMenu);
}

HWND GetPreferencesDialogHandle() noexcept
{
    return PreferencesDialog::GetHandle();
}

#ifdef ENABLE_TESTS
bool DebugGetPreferencesDialogSnapshot(PreferencesDebugSnapshot& out) noexcept
{
    return PreferencesDialog::DebugGetSnapshot(out);
}

HWND DebugGetPreferencesActivePageHandle() noexcept
{
    return PreferencesDialog::DebugGetActivePageHandle();
}

HWND DebugGetPreferencesActivePageDxHostHandle() noexcept
{
    return PreferencesDialog::DebugGetActivePageDxHostHandle();
}

HWND DebugGetPreferencesShellHostHandle() noexcept
{
    return PreferencesDialog::DebugGetShellHostHandle();
}

bool DebugSelectPreferencesCategory(const PrefCategory category) noexcept
{
    return PreferencesDialog::DebugSelectCategory(category);
}

bool DebugSelectPreferencesPluginsTreeChild(const size_t childIndex) noexcept
{
    return PreferencesDialog::DebugSelectPluginsTreeChild(childIndex);
}

bool DebugFocusPreferencesCategoryTree() noexcept
{
    return PreferencesDialog::DebugFocusCategoryTree();
}

bool DebugSendPreferencesCategoryTreeKey(const UINT virtualKey) noexcept
{
    return PreferencesDialog::DebugSendCategoryTreeKey(virtualKey);
}

bool DebugScrollPreferencesCategoryTreeByWheelDelta(const int wheelDelta) noexcept
{
    return PreferencesDialog::DebugScrollCategoryTreeByWheelDelta(wheelDelta);
}

bool DebugScrollPreferencesCategoryTreeByWheelDetents(const int detents) noexcept
{
    return PreferencesDialog::DebugScrollCategoryTreeByWheelDetents(detents);
}

bool DebugDragPreferencesPageHostDxScrollbarThumb(const int distancePx, const int moveCount) noexcept
{
    return PreferencesDialog::DebugDragPageHostDxScrollbarThumb(distancePx, moveCount);
}

bool DebugSelectPreferencesPluginsMainListRow(const size_t rowIndex) noexcept
{
    return PreferencesDialog::DebugSelectPluginsMainListRow(rowIndex);
}

bool DebugClickPreferencesPluginsMainListRow(const size_t rowIndex) noexcept
{
    return PreferencesDialog::DebugClickPluginsMainListRow(rowIndex);
}

bool DebugFindPreferencesPluginsToggleableMainListRow(size_t& outRowIndex, bool& outEnabled) noexcept
{
    return PreferencesDialog::DebugFindToggleablePluginsMainListRow(outRowIndex, outEnabled);
}

bool DebugFindPreferencesPluginsLoadableMainListRow(size_t& outRowIndex) noexcept
{
    return PreferencesDialog::DebugFindLoadablePluginsMainListRow(outRowIndex);
}

bool DebugTogglePreferencesPluginsMainListCheckbox(const size_t rowIndex) noexcept
{
    return PreferencesDialog::DebugTogglePluginsMainListCheckbox(rowIndex);
}

bool DebugGetPreferencesPluginsMainListRowEnabled(const size_t rowIndex, bool& outEnabled) noexcept
{
    return PreferencesDialog::DebugGetPluginsMainListRowEnabled(rowIndex, outEnabled);
}

bool DebugGetPreferencesPluginsMainListCheckboxClientRect(const size_t rowIndex, RECT& outRect) noexcept
{
    return PreferencesDialog::DebugGetPluginsMainListCheckboxClientRect(rowIndex, outRect);
}

bool DebugGetPreferencesPluginsMainListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    return PreferencesDialog::DebugGetPluginsMainListHeaderClientRect(columnIndex, outRect);
}

bool DebugSelectPreferencesPluginsCustomPathsListRow(const size_t rowIndex) noexcept
{
    return PreferencesDialog::DebugSelectPluginsCustomPathsListRow(rowIndex);
}

bool DebugClearPreferencesPluginsCustomPaths() noexcept
{
    return PreferencesDialog::DebugClearPluginsCustomPaths();
}

bool DebugGetPreferencesPluginsCustomPathsListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    return PreferencesDialog::DebugGetPluginsCustomPathsListHeaderClientRect(columnIndex, outRect);
}

bool DebugFocusPreferencesPluginsSearchField() noexcept
{
    return PreferencesDialog::DebugFocusPluginsSearchField();
}

bool DebugFocusPreferencesPluginsMainList() noexcept
{
    return PreferencesDialog::DebugFocusPluginsMainList();
}

bool DebugSelectPreferencesKeyboardListRow(const size_t rowIndex) noexcept
{
    return PreferencesDialog::DebugSelectKeyboardListRow(rowIndex);
}

bool DebugFindPreferencesKeyboardListRowByCommandId(std::wstring_view commandId, size_t& outRowIndex) noexcept
{
    return PreferencesDialog::DebugFindKeyboardListRowByCommandId(commandId, outRowIndex);
}

bool DebugGetPreferencesKeyboardVisibleRowChordByCommandId(std::wstring_view commandId, std::wstring& outChordText) noexcept
{
    return PreferencesDialog::DebugGetKeyboardVisibleRowChordByCommandId(commandId, outChordText);
}

bool DebugGetPreferencesKeyboardListRowClientRect(const size_t rowIndex, RECT& outRect) noexcept
{
    return PreferencesDialog::DebugGetKeyboardListRowClientRect(rowIndex, outRect);
}

bool DebugGetPreferencesKeyboardListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    return PreferencesDialog::DebugGetKeyboardListHeaderClientRect(columnIndex, outRect);
}

bool DebugHitTestPreferencesKeyboardListClientPoint(
    const POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) noexcept
{
    return PreferencesDialog::DebugHitTestKeyboardListClientPoint(clientPoint, outZone, outColumnIndex, outHeaderResize, outHostHitsList);
}

bool DebugGetPreferencesKeyboardListPointerState(PreferencesGridPointerDebugState& outState) noexcept
{
    return PreferencesDialog::DebugGetKeyboardListPointerState(outState);
}

bool DebugSelectPreferencesViewersListRow(const size_t rowIndex) noexcept
{
    return PreferencesDialog::DebugSelectViewersListRow(rowIndex);
}

bool DebugGetPreferencesViewersListRowClientRect(const size_t rowIndex, RECT& outRect) noexcept
{
    return PreferencesDialog::DebugGetViewersListRowClientRect(rowIndex, outRect);
}

bool DebugGetPreferencesViewersListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    return PreferencesDialog::DebugGetViewersListHeaderClientRect(columnIndex, outRect);
}

bool DebugHitTestPreferencesViewersListClientPoint(
    const POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) noexcept
{
    return PreferencesDialog::DebugHitTestViewersListClientPoint(clientPoint, outZone, outColumnIndex, outHeaderResize, outHostHitsList);
}

bool DebugGetPreferencesViewersListPointerState(PreferencesGridPointerDebugState& outState) noexcept
{
    return PreferencesDialog::DebugGetViewersListPointerState(outState);
}

bool DebugGetPreferencesViewersTabClientRect(const size_t tabIndex, RECT& outRect) noexcept
{
    return PreferencesDialog::DebugGetViewersTabClientRect(tabIndex, outRect);
}

bool DebugGetPreferencesViewersSelectedTabIndex(size_t& outIndex) noexcept
{
    return PreferencesDialog::DebugGetViewersSelectedTabIndex(outIndex);
}

bool DebugSelectPreferencesThemesListRow(const size_t rowIndex) noexcept
{
    return PreferencesDialog::DebugSelectThemesListRow(rowIndex);
}

bool DebugGetPreferencesThemesListRowClientRect(const size_t rowIndex, RECT& outRect) noexcept
{
    return PreferencesDialog::DebugGetThemesListRowClientRect(rowIndex, outRect);
}

bool DebugGetPreferencesThemesListHeaderClientRect(const size_t columnIndex, RECT& outRect) noexcept
{
    return PreferencesDialog::DebugGetThemesListHeaderClientRect(columnIndex, outRect);
}

bool DebugSetPreferencesPluginsSearchText(std::wstring_view text) noexcept
{
    return PreferencesDialog::DebugSetPluginsSearchText(text);
}

bool DebugSetPreferencesViewersSearchText(std::wstring_view text) noexcept
{
    return PreferencesDialog::DebugSetViewersSearchText(text);
}

bool DebugSelectPreferencesViewersDefaultAction(const bool alternate, std::wstring_view actionId) noexcept
{
    return PreferencesDialog::DebugSelectViewersDefaultAction(alternate, actionId);
}

bool DebugSelectPreferencesEditorsDefaultAction(const bool alternate, std::wstring_view actionId) noexcept
{
    return PreferencesDialog::DebugSelectEditorsDefaultAction(alternate, actionId);
}

bool DebugSelectPreferencesEditorsDefaultEditNewAction(std::wstring_view actionId) noexcept
{
    return PreferencesDialog::DebugSelectEditorsDefaultEditNewAction(actionId);
}

bool DebugSetPreferencesKeyboardSearchText(std::wstring_view text) noexcept
{
    return PreferencesDialog::DebugSetKeyboardSearchText(text);
}

bool DebugSetPreferencesKeyboardFunctionBarScope() noexcept
{
    return PreferencesDialog::DebugSetKeyboardFunctionBarScope();
}

bool DebugCapturePreferencesKeyboardShortcut(const uint32_t vk, const uint32_t modifiers) noexcept
{
    return PreferencesDialog::DebugCaptureKeyboardShortcut(vk, modifiers);
}

bool DebugGetPreferencesKeyboardSnapshot(PreferencesKeyboardDebugSnapshot& out) noexcept
{
    return PreferencesDialog::DebugGetKeyboardSnapshot(out);
}

bool DebugFocusPreferencesViewersSearchField() noexcept
{
    return PreferencesDialog::DebugFocusViewersSearchField();
}

bool DebugFocusPreferencesKeyboardSearchField() noexcept
{
    return PreferencesDialog::DebugFocusKeyboardSearchField();
}

bool DebugFocusPreferencesGeneralMenuBarToggle() noexcept
{
    return PreferencesDialog::DebugFocusGeneralMenuBarToggle();
}

bool DebugGetPreferencesGeneralMenuBarToggleChecked(bool& outChecked) noexcept
{
    return PreferencesDialog::DebugGetGeneralMenuBarToggleChecked(outChecked);
}

bool DebugSetPreferencesGeneralCompactMode(bool checked) noexcept
{
    return PreferencesDialog::DebugSetGeneralCompactMode(checked);
}

bool DebugSelectPreferencesGeneralLanguage(std::wstring_view displayText) noexcept
{
    return PreferencesDialog::DebugSelectGeneralLanguage(displayText);
}

bool DebugSelectPreferencesGeneralReducedMotion(std::wstring_view displayText) noexcept
{
    return PreferencesDialog::DebugSelectGeneralReducedMotion(displayText);
}

bool DebugSelectPreferencesGeneralWindowBackdrop(std::wstring_view displayText) noexcept
{
    return PreferencesDialog::DebugSelectGeneralWindowBackdrop(displayText);
}

bool DebugFocusPreferencesPanesLeftDisplayToggle() noexcept
{
    return PreferencesDialog::DebugFocusPanesLeftDisplayToggle();
}

bool DebugSelectPreferencesPanesLeftDisplay(std::wstring_view displayText) noexcept
{
    return PreferencesDialog::DebugSelectPanesLeftDisplay(displayText);
}

bool DebugFocusPreferencesPanesLeftStatusBarToggle() noexcept
{
    return PreferencesDialog::DebugFocusPanesLeftStatusBarToggle();
}

bool DebugGetPreferencesPanesLeftStatusBarToggleChecked(bool& outChecked) noexcept
{
    return PreferencesDialog::DebugGetPanesLeftStatusBarToggleChecked(outChecked);
}

bool DebugFocusPreferencesHotPathsFirstPathField() noexcept
{
    return PreferencesDialog::DebugFocusHotPathsFirstPathField();
}

bool DebugGetPreferencesHotPathsFirstPathText(std::wstring& outText) noexcept
{
    return PreferencesDialog::DebugGetHotPathsFirstPathText(outText);
}

bool DebugSetPreferencesHotPathsFirstPathText(std::wstring_view text) noexcept
{
    return PreferencesDialog::DebugSetHotPathsFirstPathText(text);
}

bool DebugFocusPreferencesHotPathsOpenPrefsToggle() noexcept
{
    return PreferencesDialog::DebugFocusHotPathsOpenPrefsToggle();
}

bool DebugGetPreferencesHotPathsOpenPrefsToggleChecked(bool& outChecked) noexcept
{
    return PreferencesDialog::DebugGetHotPathsOpenPrefsToggleChecked(outChecked);
}

bool DebugFocusPreferencesAdvancedBypassHelloToggle() noexcept
{
    return PreferencesDialog::DebugFocusAdvancedBypassHelloToggle();
}

bool DebugFocusPreferencesMonitorToolbarToggle() noexcept
{
    return PreferencesDialog::DebugFocusMonitorToolbarToggle();
}

bool DebugSelectPreferencesMonitorFilterPreset(std::wstring_view displayText) noexcept
{
    return PreferencesDialog::DebugSelectMonitorFilterPreset(displayText);
}

void DebugSetPreferencesSettingsFileOpenCapture(const bool capture) noexcept
{
    PreferencesDialog::DebugSetSettingsFileOpenCapture(capture);
}

void DebugClearPreferencesLastSettingsFileOpen() noexcept
{
    PreferencesDialog::DebugClearLastSettingsFileOpen();
}

bool DebugGetPreferencesLastSettingsFileOpen(std::filesystem::path& outPath, HRESULT& outHr) noexcept
{
    return PreferencesDialog::DebugGetLastSettingsFileOpen(outPath, outHr);
}

bool DebugFocusPreferencesFileOperationsPreCalcEnabledToggle() noexcept
{
    return PreferencesDialog::DebugFocusFileOperationsPreCalcEnabledToggle();
}

bool DebugGetPreferencesFileOperationsPreCalcEnabledToggleChecked(bool& outChecked) noexcept
{
    return PreferencesDialog::DebugGetFileOperationsPreCalcEnabledToggleChecked(outChecked);
}

bool DebugSelectPreferencesFileOperationsBandwidthPreset(std::wstring_view displayText) noexcept
{
    return PreferencesDialog::DebugSelectFileOperationsBandwidthPreset(displayText);
}

bool DebugSelectPreferencesCompareDirectoriesContentWorkers(std::wstring_view displayText) noexcept
{
    return PreferencesDialog::DebugSelectCompareDirectoriesContentWorkers(displayText);
}

bool DebugSetPreferencesFileOperationsBridgeBufferText(std::wstring_view text) noexcept
{
    return PreferencesDialog::DebugSetFileOperationsBridgeBufferText(text);
}

bool DebugCancelPreferencesDialog() noexcept
{
    return PreferencesDialog::DebugCancelDialog();
}

bool DebugResetPreferencesToDefaults(const bool confirm) noexcept
{
    return PreferencesDialog::DebugResetAllToDefaults(confirm);
}

bool DebugFocusPreferencesCompareDirectoriesSubdirectoriesToggle() noexcept
{
    return PreferencesDialog::DebugFocusCompareDirectoriesSubdirectoriesToggle();
}

bool DebugFocusPreferencesCompareDirectoriesTarget(const PreferencesCompareDirectoriesDebugFocusTarget target) noexcept
{
    return PreferencesDialog::DebugFocusCompareDirectoriesTarget(target);
}

bool DebugGetPreferencesCompareDirectoriesToggleChecked(const PreferencesCompareDirectoriesDebugFocusTarget target, bool& outChecked) noexcept
{
    return PreferencesDialog::DebugGetCompareDirectoriesToggleChecked(target, outChecked);
}

bool DebugSetPreferencesThemesSearchText(std::wstring_view text) noexcept
{
    return PreferencesDialog::DebugSetThemesSearchText(text);
}

bool DebugFocusPreferencesThemesSearchField() noexcept
{
    return PreferencesDialog::DebugFocusThemesSearchField();
}

bool DebugScrollPreferencesPluginsMainListByWheelDetents(const int detents) noexcept
{
    return PreferencesDialog::DebugScrollPluginsMainListByWheelDetents(detents);
}

bool DebugScrollPreferencesPluginsCustomPathsListByWheelDetents(const int detents) noexcept
{
    return PreferencesDialog::DebugScrollPluginsCustomPathsListByWheelDetents(detents);
}

bool DebugScrollPreferencesKeyboardListByWheelDetents(const int detents) noexcept
{
    return PreferencesDialog::DebugScrollKeyboardListByWheelDetents(detents);
}

bool DebugScrollPreferencesViewersListByWheelDetents(const int detents) noexcept
{
    return PreferencesDialog::DebugScrollViewersListByWheelDetents(detents);
}

bool DebugScrollPreferencesThemesListByWheelDetents(const int detents) noexcept
{
    return PreferencesDialog::DebugScrollThemesListByWheelDetents(detents);
}
#endif
