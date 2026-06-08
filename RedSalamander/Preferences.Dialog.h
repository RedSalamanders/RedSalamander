#pragma once

#include <filesystem>
#include <string_view>

#include "AppTheme.h"
#include "SettingsStore.h"

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

enum class PrefCategory : int;
struct PreferencesDebugSnapshot;
struct PreferencesKeyboardDebugSnapshot;

namespace PreferencesDialog
{
[[nodiscard]] bool Show(
    HWND owner, std::wstring_view appId, Common::Settings::Settings& settings, const AppTheme& theme, PrefCategory initialCategory) noexcept;

[[nodiscard]] HWND GetHandle() noexcept;

#ifdef ENABLE_TESTS
[[nodiscard]] bool DebugGetSnapshot(::PreferencesDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugGetKeyboardSnapshot(::PreferencesKeyboardDebugSnapshot& out) noexcept;
[[nodiscard]] HWND DebugGetActivePageHandle() noexcept;
[[nodiscard]] HWND DebugGetActivePageDxHostHandle() noexcept;
[[nodiscard]] HWND DebugGetShellHostHandle() noexcept;
[[nodiscard]] bool DebugSelectCategory(PrefCategory category) noexcept;
[[nodiscard]] bool DebugSelectPluginsTreeChild(size_t childIndex) noexcept;
[[nodiscard]] bool DebugScrollCategoryTreeByWheelDelta(int wheelDelta) noexcept;
[[nodiscard]] bool DebugScrollCategoryTreeByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugDragPageHostDxScrollbarThumb(int distancePx, int moveCount) noexcept;
[[nodiscard]] bool DebugSelectPluginsMainListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugClickPluginsMainListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugFindToggleablePluginsMainListRow(size_t& outRowIndex, bool& outEnabled) noexcept;
[[nodiscard]] bool DebugFindLoadablePluginsMainListRow(size_t& outRowIndex) noexcept;
[[nodiscard]] bool DebugTogglePluginsMainListCheckbox(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugGetPluginsMainListRowEnabled(size_t rowIndex, bool& outEnabled) noexcept;
[[nodiscard]] bool DebugGetPluginsMainListCheckboxClientRect(size_t rowIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetPluginsMainListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugSelectPluginsCustomPathsListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugClearPluginsCustomPaths() noexcept;
[[nodiscard]] bool DebugGetPluginsCustomPathsListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugFocusPluginsMainList() noexcept;
[[nodiscard]] bool DebugFocusPluginsSearchField() noexcept;
[[nodiscard]] bool DebugSelectKeyboardListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugFindKeyboardListRowByCommandId(std::wstring_view commandId, size_t& outRowIndex) noexcept;
[[nodiscard]] bool DebugGetKeyboardVisibleRowChordByCommandId(std::wstring_view commandId, std::wstring& outChordText) noexcept;
[[nodiscard]] bool DebugGetKeyboardListRowClientRect(size_t rowIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetKeyboardListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugHitTestKeyboardListClientPoint(
    POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) noexcept;
[[nodiscard]] bool DebugGetKeyboardListPointerState(::PreferencesGridPointerDebugState& outState) noexcept;
[[nodiscard]] bool DebugSelectViewersListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugGetViewersListRowClientRect(size_t rowIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetViewersListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugHitTestViewersListClientPoint(
    POINT clientPoint, uint32_t& outZone, size_t& outColumnIndex, bool& outHeaderResize, bool& outHostHitsList) noexcept;
[[nodiscard]] bool DebugGetViewersListPointerState(::PreferencesGridPointerDebugState& outState) noexcept;
[[nodiscard]] bool DebugGetViewersTabClientRect(size_t tabIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetViewersSelectedTabIndex(size_t& outIndex) noexcept;
[[nodiscard]] bool DebugSelectThemesListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugGetThemesListRowClientRect(size_t rowIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetThemesListHeaderClientRect(size_t columnIndex, RECT& outRect) noexcept;
[[nodiscard]] bool DebugSetViewersSearchText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSelectViewersDefaultAction(bool alternate, std::wstring_view actionId) noexcept;
[[nodiscard]] bool DebugSelectEditorsDefaultAction(bool alternate, std::wstring_view actionId) noexcept;
[[nodiscard]] bool DebugSelectEditorsDefaultEditNewAction(std::wstring_view actionId) noexcept;
[[nodiscard]] bool DebugSetKeyboardSearchText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetKeyboardFunctionBarScope() noexcept;
[[nodiscard]] bool DebugCaptureKeyboardShortcut(uint32_t vk, uint32_t modifiers) noexcept;
[[nodiscard]] bool DebugFocusViewersSearchField() noexcept;
[[nodiscard]] bool DebugFocusKeyboardSearchField() noexcept;
[[nodiscard]] bool DebugFocusGeneralMenuBarToggle() noexcept;
[[nodiscard]] bool DebugGetGeneralMenuBarToggleChecked(bool& outChecked) noexcept;
[[nodiscard]] bool DebugSetGeneralCompactMode(bool checked) noexcept;
[[nodiscard]] bool DebugSelectGeneralLanguage(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugSelectGeneralReducedMotion(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugSelectGeneralWindowBackdrop(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugFocusPanesLeftDisplayToggle() noexcept;
[[nodiscard]] bool DebugSelectPanesLeftDisplay(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugFocusPanesLeftStatusBarToggle() noexcept;
[[nodiscard]] bool DebugGetPanesLeftStatusBarToggleChecked(bool& outChecked) noexcept;
[[nodiscard]] bool DebugFocusHotPathsFirstPathField() noexcept;
[[nodiscard]] bool DebugGetHotPathsFirstPathText(std::wstring& outText) noexcept;
[[nodiscard]] bool DebugSetHotPathsFirstPathText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugFocusHotPathsOpenPrefsToggle() noexcept;
[[nodiscard]] bool DebugGetHotPathsOpenPrefsToggleChecked(bool& outChecked) noexcept;
[[nodiscard]] bool DebugFocusAdvancedBypassHelloToggle() noexcept;
[[nodiscard]] bool DebugFocusMonitorToolbarToggle() noexcept;
[[nodiscard]] bool DebugSelectMonitorFilterPreset(std::wstring_view displayText) noexcept;
void DebugSetSettingsFileOpenCapture(bool capture) noexcept;
void DebugClearLastSettingsFileOpen() noexcept;
[[nodiscard]] bool DebugGetLastSettingsFileOpen(std::filesystem::path& outPath, HRESULT& outHr) noexcept;
[[nodiscard]] bool DebugFocusFileOperationsPreCalcEnabledToggle() noexcept;
[[nodiscard]] bool DebugGetFileOperationsPreCalcEnabledToggleChecked(bool& outChecked) noexcept;
[[nodiscard]] bool DebugSelectFileOperationsBandwidthPreset(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugSelectCompareDirectoriesContentWorkers(std::wstring_view displayText) noexcept;
[[nodiscard]] bool DebugSetFileOperationsBridgeBufferText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugCancelDialog() noexcept;
[[nodiscard]] bool DebugResetAllToDefaults(bool confirm) noexcept;
[[nodiscard]] bool DebugFocusCompareDirectoriesSubdirectoriesToggle() noexcept;
[[nodiscard]] bool DebugFocusCompareDirectoriesTarget(PreferencesCompareDirectoriesDebugFocusTarget target) noexcept;
[[nodiscard]] bool DebugGetCompareDirectoriesToggleChecked(PreferencesCompareDirectoriesDebugFocusTarget target, bool& outChecked) noexcept;
[[nodiscard]] bool DebugSetPluginsSearchText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetThemesSearchText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugFocusThemesSearchField() noexcept;
[[nodiscard]] bool DebugScrollPluginsMainListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugScrollPluginsCustomPathsListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugScrollKeyboardListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugScrollViewersListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugScrollThemesListByWheelDetents(int detents) noexcept;
#endif
} // namespace PreferencesDialog
