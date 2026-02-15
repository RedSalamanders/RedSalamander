#pragma once

// Internal types shared across Preferences dialog implementation files.
// Keep this header private to Preferences translation units.

#include "framework.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include <commctrl.h>

#include "AppTheme.h"
#include "SettingsSchemaParser.h"
#include "SettingsStore.h"

enum class PrefCategory : int
{
    General = 0,
    Panes,
    Viewers,
    Editors,
    Keyboard,
    Mouse,
    Themes,
    Plugins,
    Advanced,
};

enum class ShortcutScope : uint8_t
{
    FunctionBar,
    FolderView,
};

enum class ThemeSchemaSource : uint8_t
{
    Builtin,
    Settings,
    File,
    New,
};

// Layout constants (DPI-independent values in logical units)
namespace PrefsLayoutConstants
{
inline constexpr int kRowHeightDip        = 26;
inline constexpr int kTitleHeightDip      = 18;
inline constexpr int kCardPaddingXDip     = 12;
inline constexpr int kCardPaddingYDip     = 8;
inline constexpr int kCardGapYDip         = 2;
inline constexpr int kCardGapXDip         = 12;
inline constexpr int kCardSpacingYDip     = 8;
inline constexpr int kSectionSpacingYDip  = 16;
inline constexpr int kCornerRadiusDip     = 6;
inline constexpr int kMinToggleWidthDip   = 90;
inline constexpr int kTogglePaddingXDip   = 6;
inline constexpr int kToggleGapXDip       = 8;
inline constexpr int kToggleTrackWidthDip = 34;
inline constexpr int kEditHeightDip       = 28;
inline constexpr int kComboHeightDip      = 28;
inline constexpr int kButtonHeightDip     = 28;
inline constexpr int kMarginDip           = 16;
inline constexpr int kGapYDip             = 12;
inline constexpr int kHeaderHeightDip     = 20;
inline constexpr int kFramePaddingDip     = 2;
inline constexpr int kMinEditWidthDip     = 100;
inline constexpr int kMaxEditWidthDip     = 220;
inline constexpr int kMinComboWidthDip    = 80;
inline constexpr int kMediumComboWidthDip = 140;
inline constexpr int kLargeComboWidthDip  = 180;
} // namespace PrefsLayoutConstants

// Monitor filter mask bits for Advanced pane
enum class MonitorFilterBit : uint32_t
{
    Text    = 0x01u,
    Error   = 0x02u,
    Warning = 0x04u,
    Info    = 0x08u,
    Debug   = 0x10u,
};

[[nodiscard]] inline constexpr uint32_t operator|(MonitorFilterBit a, MonitorFilterBit b) noexcept
{
    return static_cast<uint32_t>(a) | static_cast<uint32_t>(b);
}

[[nodiscard]] inline constexpr uint32_t operator|(uint32_t a, MonitorFilterBit b) noexcept
{
    return a | static_cast<uint32_t>(b);
}

[[nodiscard]] inline constexpr bool HasFlag(uint32_t mask, MonitorFilterBit bit) noexcept
{
    return (mask & static_cast<uint32_t>(bit)) != 0;
}

struct ThemeComboItem
{
    std::wstring id;
    std::wstring displayName;
    ThemeSchemaSource source = ThemeSchemaSource::Builtin;
};

struct ViewerPluginOption
{
    std::wstring id;
    std::wstring displayName;
};

enum class PrefsPluginType : uint8_t
{
    FileSystem,
    Viewer,
};

struct PrefsPluginListItem
{
    PrefsPluginType type = PrefsPluginType::FileSystem;
    size_t index         = 0;
};

enum class PrefsPluginConfigFieldType : uint8_t
{
    Text,
    Value,
    Bool,
    Option,
    Selection,
};

struct PrefsPluginConfigChoice
{
    std::wstring value;
    std::wstring label;
};

struct PrefsPluginConfigField
{
    PrefsPluginConfigFieldType type = PrefsPluginConfigFieldType::Text;
    std::wstring key;
    std::wstring label;
    std::wstring description;
    bool browseFolder = false;

    bool hasMin = false;
    bool hasMax = false;
    int64_t min = 0;
    int64_t max = 0;

    std::wstring defaultText;
    int64_t defaultInt = 0;
    bool defaultBool   = false;
    std::wstring defaultOption;
    std::vector<std::wstring> defaultSelection;
    std::vector<PrefsPluginConfigChoice> choices;
};

struct PrefsPluginConfigFieldControls
{
    PrefsPluginConfigFieldControls() = default;

    PrefsPluginConfigFieldControls(const PrefsPluginConfigFieldControls&)            = delete;
    PrefsPluginConfigFieldControls& operator=(const PrefsPluginConfigFieldControls&) = delete;

    PrefsPluginConfigFieldControls(PrefsPluginConfigFieldControls&&)            = default;
    PrefsPluginConfigFieldControls& operator=(PrefsPluginConfigFieldControls&&) = default;

    PrefsPluginConfigField field;
    std::wstring schemaDefaultOption;
    wil::unique_hwnd label;
    wil::unique_hwnd description;
    wil::unique_hwnd editFrame;
    wil::unique_hwnd edit;
    wil::unique_hwnd browseButton;
    wil::unique_hwnd comboFrame;
    wil::unique_hwnd combo;
    wil::unique_hwnd toggle;
    size_t toggleOnChoiceIndex  = 0;
    size_t toggleOffChoiceIndex = 0;
    std::vector<wil::unique_hwnd> choiceButtons;
};

namespace PrefsNavTree
{
inline constexpr uintptr_t kPluginTag      = uintptr_t(1) << ((sizeof(uintptr_t) * 8u) - 1u);
inline constexpr uintptr_t kPluginTypeMask = 0xFFu;
inline constexpr int kPluginIndexShift     = 8;

[[nodiscard]] constexpr LPARAM EncodePluginData(PrefsPluginType type, size_t index) noexcept
{
    uintptr_t value = kPluginTag;
    value |= (static_cast<uintptr_t>(type) & kPluginTypeMask);
    value |= (static_cast<uintptr_t>(index) << kPluginIndexShift);
    return static_cast<LPARAM>(value);
}

[[nodiscard]] inline bool TryDecodePluginData(LPARAM data, PrefsPluginListItem& out) noexcept
{
    const uintptr_t value = static_cast<uintptr_t>(data);
    if ((value & kPluginTag) == 0)
    {
        return false;
    }

    const uintptr_t payload = (value & ~kPluginTag);
    out.type                = static_cast<PrefsPluginType>(payload & kPluginTypeMask);
    out.index               = static_cast<size_t>(payload >> kPluginIndexShift);
    return true;
}
} // namespace PrefsNavTree

struct KeyboardShortcutRow
{
    ShortcutScope scope = ShortcutScope::FunctionBar;
    std::wstring commandId;
    std::wstring commandDisplayName;
    std::wstring chordText;
    std::optional<size_t> bindingIndex;
    uint32_t vk        = 0;
    uint32_t modifiers = 0;
    bool placeholder   = false;
    bool hasConflict   = false;
};

struct PreferencesDialogState
{
    PreferencesDialogState()                                         = default;
    PreferencesDialogState(const PreferencesDialogState&)            = delete;
    PreferencesDialogState& operator=(const PreferencesDialogState&) = delete;

    // Dialog Ownership and Settings
    HWND owner                           = nullptr;
    Common::Settings::Settings* settings = nullptr;
    std::wstring appId;
    AppTheme theme{};

    // Settings Management
    Common::Settings::Settings baselineSettings;
    Common::Settings::Settings workingSettings;

    // Schema-driven UI support
    std::vector<SettingsSchemaParser::SettingField> schemaFields;

    bool dirty       = false;
    bool appliedOnce = false;

    // Navigation State
    PrefCategory currentCategory = PrefCategory::General;
    PrefCategory initialCategory = PrefCategory::General;
    std::optional<PrefsPluginListItem> pluginsSelectedPlugin;

    // Layout and Sizing
    int categoryListWidthPx = 0;
    SIZE minTrackSizePx{};

    int pageScrollY             = 0;
    int pageScrollMaxY          = 0;
    int pageWheelDeltaRemainder = 0;
    std::vector<RECT> pageSettingCards;
    bool pageHostRelayoutInProgress = false;
    bool pageHostIgnoreSize         = false;

    // Theme Resources (RAII-managed)
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> backgroundBrush;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> cardBrush;
    COLORREF cardBackgroundColor = RGB(255, 255, 255);
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> inputBrush;
    COLORREF inputBackgroundColor = RGB(255, 255, 255);
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> inputFocusedBrush;
    COLORREF inputFocusedBackgroundColor = RGB(255, 255, 255);
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> inputDisabledBrush;
    COLORREF inputDisabledBackgroundColor = RGB(255, 255, 255);
    wil::unique_hfont italicFont;
    wil::unique_hfont boldFont;
    wil::unique_hfont titleFont;
    wil::unique_hfont uiFont;

    // Dialog Structure Controls
    HWND categoryTree = nullptr;
    std::array<HTREEITEM, 9> categoryTreeItems{};
    HTREEITEM pluginsTreeRoot = nullptr;
    HWND pageHost             = nullptr;
    HWND pageTitle            = nullptr;
    HWND pageDescription      = nullptr;

    // General Page Controls (RAII-managed)
    wil::unique_hwnd menuBarLabel;
    wil::unique_hwnd menuBarToggle;
    wil::unique_hwnd menuBarDescription;
    wil::unique_hwnd functionBarLabel;
    wil::unique_hwnd functionBarToggle;
    wil::unique_hwnd functionBarDescription;
    wil::unique_hwnd splashScreenLabel;
    wil::unique_hwnd splashScreenToggle;
    wil::unique_hwnd splashScreenDescription;

    // Panes Page Controls (RAII-managed)
    wil::unique_hwnd panesLeftHeader;
    wil::unique_hwnd panesLeftDisplayLabel;
    wil::unique_hwnd panesLeftDisplayFrame;
    wil::unique_hwnd panesLeftDisplayCombo;
    wil::unique_hwnd panesLeftDisplayToggle;
    wil::unique_hwnd panesLeftSortByLabel;
    wil::unique_hwnd panesLeftSortByFrame;
    wil::unique_hwnd panesLeftSortByCombo;
    wil::unique_hwnd panesLeftSortDirLabel;
    wil::unique_hwnd panesLeftSortDirFrame;
    wil::unique_hwnd panesLeftSortDirCombo;
    wil::unique_hwnd panesLeftSortDirToggle;
    wil::unique_hwnd panesLeftStatusBarLabel;
    wil::unique_hwnd panesLeftStatusBarToggle;
    wil::unique_hwnd panesLeftStatusBarDescription;

    wil::unique_hwnd panesRightHeader;
    wil::unique_hwnd panesRightDisplayLabel;
    wil::unique_hwnd panesRightDisplayFrame;
    wil::unique_hwnd panesRightDisplayCombo;
    wil::unique_hwnd panesRightDisplayToggle;
    wil::unique_hwnd panesRightSortByLabel;
    wil::unique_hwnd panesRightSortByFrame;
    wil::unique_hwnd panesRightSortByCombo;
    wil::unique_hwnd panesRightSortDirLabel;
    wil::unique_hwnd panesRightSortDirFrame;
    wil::unique_hwnd panesRightSortDirCombo;
    wil::unique_hwnd panesRightSortDirToggle;
    wil::unique_hwnd panesRightStatusBarLabel;
    wil::unique_hwnd panesRightStatusBarToggle;
    wil::unique_hwnd panesRightStatusBarDescription;

    wil::unique_hwnd panesHistoryLabel;
    wil::unique_hwnd panesHistoryFrame;
    wil::unique_hwnd panesHistoryEdit;
    wil::unique_hwnd panesHistoryDescription;

    // Viewers Page Controls (RAII-managed)
    wil::unique_hwnd viewersSearchLabel;
    wil::unique_hwnd viewersSearchFrame;
    wil::unique_hwnd viewersSearchEdit;
    wil::unique_hwnd viewersList;
    wil::unique_hwnd viewersExtensionLabel;
    wil::unique_hwnd viewersExtensionFrame;
    wil::unique_hwnd viewersExtensionEdit;
    wil::unique_hwnd viewersViewerLabel;
    wil::unique_hwnd viewersViewerFrame;
    wil::unique_hwnd viewersViewerCombo;
    wil::unique_hwnd viewersSaveButton;
    wil::unique_hwnd viewersRemoveButton;
    wil::unique_hwnd viewersResetButton;
    wil::unique_hwnd viewersHint;

    std::vector<std::wstring> viewersExtensionKeys;
    std::vector<ViewerPluginOption> viewersPluginOptions;

    // Editors Page Controls (RAII-managed)
    wil::unique_hwnd editorsNote;

    // Keyboard Page Controls (RAII-managed)
    wil::unique_hwnd keyboardSearchLabel;
    wil::unique_hwnd keyboardSearchFrame;
    wil::unique_hwnd keyboardSearchEdit;
    wil::unique_hwnd keyboardScopeLabel;
    wil::unique_hwnd keyboardScopeFrame;
    wil::unique_hwnd keyboardScopeCombo;
    wil::unique_hwnd keyboardList;
    wil::unique_hwnd keyboardHint;
    wil::unique_hwnd keyboardAssign;
    wil::unique_hwnd keyboardRemove;
    wil::unique_hwnd keyboardReset;
    wil::unique_hwnd keyboardImport;
    wil::unique_hwnd keyboardExport;

    wil::unique_any<HIMAGELIST, decltype(&::ImageList_Destroy), ::ImageList_Destroy> keyboardImageList;

    bool keyboardCaptureActive         = false;
    ShortcutScope keyboardCaptureScope = ShortcutScope::FunctionBar;
    std::wstring keyboardCaptureCommandId;
    std::optional<size_t> keyboardCaptureBindingIndex;
    std::optional<uint32_t> keyboardCapturePendingVk;
    uint32_t keyboardCapturePendingModifiers = 0;
    std::wstring keyboardCaptureConflictCommandId;
    std::optional<size_t> keyboardCaptureConflictBindingIndex;
    bool keyboardCaptureConflictMultiple = false;

    std::vector<KeyboardShortcutRow> keyboardRows;

    // Mouse Page Controls (RAII-managed)
    wil::unique_hwnd mouseNote;

    // Themes Page Controls (RAII-managed)
    wil::unique_hwnd themesThemeLabel;
    wil::unique_hwnd themesThemeFrame;
    wil::unique_hwnd themesThemeCombo;
    wil::unique_hwnd themesNameLabel;
    wil::unique_hwnd themesNameFrame;
    wil::unique_hwnd themesNameEdit;
    wil::unique_hwnd themesBaseLabel;
    wil::unique_hwnd themesBaseFrame;
    wil::unique_hwnd themesBaseCombo;
    wil::unique_hwnd themesSearchLabel;
    wil::unique_hwnd themesSearchFrame;
    wil::unique_hwnd themesSearchEdit;
    wil::unique_hwnd themesColorsList;
    wil::unique_hwnd themesKeyLabel;
    wil::unique_hwnd themesKeyFrame;
    wil::unique_hwnd themesKeyEdit;
    wil::unique_hwnd themesColorLabel;
    wil::unique_hwnd themesColorSwatch;
    wil::unique_hwnd themesColorFrame;
    wil::unique_hwnd themesColorEdit;
    wil::unique_hwnd themesPickColor;
    wil::unique_hwnd themesSetOverride;
    wil::unique_hwnd themesRemoveOverride;
    wil::unique_hwnd themesLoadFromFile;
    wil::unique_hwnd themesDuplicateTheme;
    wil::unique_hwnd themesSaveTheme;
    wil::unique_hwnd themesApplyTemporarily;
    wil::unique_hwnd themesNote;

    std::vector<ThemeComboItem> themeComboItems;
    std::vector<Common::Settings::ThemeDefinition> themeFileThemes;

    // Plugins Page Controls (RAII-managed)
    wil::unique_hwnd pluginsConfigureButton;
    wil::unique_hwnd pluginsTestButton;
    wil::unique_hwnd pluginsTestAllButton;
    wil::unique_hwnd pluginsNote;
    wil::unique_hwnd pluginsSearchLabel;
    wil::unique_hwnd pluginsSearchFrame;
    wil::unique_hwnd pluginsSearchEdit;
    wil::unique_hwnd pluginsList;
    wil::unique_hwnd pluginsCustomPathsHeader;
    wil::unique_hwnd pluginsCustomPathsNote;
    wil::unique_hwnd pluginsCustomPathsList;
    wil::unique_hwnd pluginsCustomPathsAddButton;
    wil::unique_hwnd pluginsCustomPathsRemoveButton;

    // Plugins details subpage (when a plugin tree child is selected). (RAII-managed)
    wil::unique_hwnd pluginsDetailsHint;
    wil::unique_hwnd pluginsDetailsIdLabel;
    wil::unique_hwnd pluginsDetailsConfigLabel;
    wil::unique_hwnd pluginsDetailsConfigFrame;
    wil::unique_hwnd pluginsDetailsConfigEdit;
    wil::unique_hwnd pluginsDetailsConfigError;
    std::wstring pluginsDetailsConfigPluginId;
    std::vector<PrefsPluginConfigFieldControls> pluginsDetailsConfigFields;

    std::vector<PrefsPluginListItem> pluginsListItems;

    // Advanced Page Controls (RAII-managed)
    wil::unique_hwnd advancedConnectionsHelloHeader;
    wil::unique_hwnd advancedConnectionsBypassHelloLabel;
    wil::unique_hwnd advancedConnectionsBypassHelloToggle;
    wil::unique_hwnd advancedConnectionsBypassHelloDescription;
    wil::unique_hwnd advancedConnectionsHelloTimeoutLabel;
    wil::unique_hwnd advancedConnectionsHelloTimeoutFrame;
    wil::unique_hwnd advancedConnectionsHelloTimeoutEdit;
    wil::unique_hwnd advancedConnectionsHelloTimeoutDescription;

    wil::unique_hwnd advancedMonitorHeader;
    wil::unique_hwnd advancedMonitorToolbarLabel;
    wil::unique_hwnd advancedMonitorToolbarToggle;
    wil::unique_hwnd advancedMonitorToolbarDescription;
    wil::unique_hwnd advancedMonitorLineNumbersLabel;
    wil::unique_hwnd advancedMonitorLineNumbersToggle;
    wil::unique_hwnd advancedMonitorLineNumbersDescription;
    wil::unique_hwnd advancedMonitorAlwaysOnTopLabel;
    wil::unique_hwnd advancedMonitorAlwaysOnTopToggle;
    wil::unique_hwnd advancedMonitorAlwaysOnTopDescription;
    wil::unique_hwnd advancedMonitorShowIdsLabel;
    wil::unique_hwnd advancedMonitorShowIdsToggle;
    wil::unique_hwnd advancedMonitorShowIdsDescription;
    wil::unique_hwnd advancedMonitorAutoScrollLabel;
    wil::unique_hwnd advancedMonitorAutoScrollToggle;
    wil::unique_hwnd advancedMonitorAutoScrollDescription;
    wil::unique_hwnd advancedMonitorFilterPresetLabel;
    wil::unique_hwnd advancedMonitorFilterPresetFrame;
    wil::unique_hwnd advancedMonitorFilterPresetCombo;
    wil::unique_hwnd advancedMonitorFilterPresetDescription;
    wil::unique_hwnd advancedMonitorFilterMaskLabel;
    wil::unique_hwnd advancedMonitorFilterMaskFrame;
    wil::unique_hwnd advancedMonitorFilterMaskEdit;
    wil::unique_hwnd advancedMonitorFilterMaskDescription;

    wil::unique_hwnd advancedMonitorFilterTextLabel;
    wil::unique_hwnd advancedMonitorFilterTextToggle;
    wil::unique_hwnd advancedMonitorFilterTextDescription;
    wil::unique_hwnd advancedMonitorFilterErrorLabel;
    wil::unique_hwnd advancedMonitorFilterErrorToggle;
    wil::unique_hwnd advancedMonitorFilterErrorDescription;
    wil::unique_hwnd advancedMonitorFilterWarningLabel;
    wil::unique_hwnd advancedMonitorFilterWarningToggle;
    wil::unique_hwnd advancedMonitorFilterWarningDescription;
    wil::unique_hwnd advancedMonitorFilterInfoLabel;
    wil::unique_hwnd advancedMonitorFilterInfoToggle;
    wil::unique_hwnd advancedMonitorFilterInfoDescription;
    wil::unique_hwnd advancedMonitorFilterDebugLabel;
    wil::unique_hwnd advancedMonitorFilterDebugToggle;
    wil::unique_hwnd advancedMonitorFilterDebugDescription;

    wil::unique_hwnd advancedCacheHeader;
    wil::unique_hwnd advancedCacheDirectoryInfoMaxBytesLabel;
    wil::unique_hwnd advancedCacheDirectoryInfoMaxBytesFrame;
    wil::unique_hwnd advancedCacheDirectoryInfoMaxBytesEdit;
    wil::unique_hwnd advancedCacheDirectoryInfoMaxBytesDescription;
    wil::unique_hwnd advancedCacheDirectoryInfoMaxWatchersLabel;
    wil::unique_hwnd advancedCacheDirectoryInfoMaxWatchersFrame;
    wil::unique_hwnd advancedCacheDirectoryInfoMaxWatchersEdit;
    wil::unique_hwnd advancedCacheDirectoryInfoMaxWatchersDescription;
    wil::unique_hwnd advancedCacheDirectoryInfoMruWatchedLabel;
    wil::unique_hwnd advancedCacheDirectoryInfoMruWatchedFrame;
    wil::unique_hwnd advancedCacheDirectoryInfoMruWatchedEdit;
    wil::unique_hwnd advancedCacheDirectoryInfoMruWatchedDescription;

    wil::unique_hwnd advancedFileOperationsHeader;
    wil::unique_hwnd advancedFileOperationsMaxDiagnosticsLogFilesLabel;
    wil::unique_hwnd advancedFileOperationsMaxDiagnosticsLogFilesFrame;
    wil::unique_hwnd advancedFileOperationsMaxDiagnosticsLogFilesEdit;
    wil::unique_hwnd advancedFileOperationsMaxDiagnosticsLogFilesDescription;
    wil::unique_hwnd advancedFileOperationsDiagnosticsInfoLabel;
    wil::unique_hwnd advancedFileOperationsDiagnosticsInfoToggle;
    wil::unique_hwnd advancedFileOperationsDiagnosticsInfoDescription;
    wil::unique_hwnd advancedFileOperationsDiagnosticsDebugLabel;
    wil::unique_hwnd advancedFileOperationsDiagnosticsDebugToggle;
    wil::unique_hwnd advancedFileOperationsDiagnosticsDebugDescription;

    // Refresh State Flags
    bool previewApplied        = false;
    bool refreshingPanesPage   = false;
    bool refreshingThemesPage  = false;
    bool refreshingPluginsPage = false;
};

void SetDirty(HWND dlg, PreferencesDialogState& state) noexcept;

namespace PrefsUi
{
[[nodiscard]] std::wstring GetWindowTextString(HWND hwnd) noexcept;
[[nodiscard]] inline std::wstring GetWindowTextString(const wil::unique_hwnd& hwnd) noexcept
{
    return GetWindowTextString(hwnd.get());
}
[[nodiscard]] int MeasureStaticTextHeight(HWND referenceWindow, HFONT font, int width, std::wstring_view text) noexcept;
[[nodiscard]] inline int MeasureStaticTextHeight(const wil::unique_hwnd& referenceWindow, HFONT font, int width, std::wstring_view text) noexcept
{
    return MeasureStaticTextHeight(referenceWindow.get(), font, width, text);
}
[[nodiscard]] std::wstring_view TrimWhitespace(std::wstring_view text) noexcept;
[[nodiscard]] bool ContainsCaseInsensitive(std::wstring_view haystack, std::wstring_view needle) noexcept;
void InvalidateComboBox(HWND combo) noexcept;
inline void InvalidateComboBox(const wil::unique_hwnd& combo) noexcept
{
    InvalidateComboBox(combo.get());
}
void SelectComboItemByData(HWND combo, LPARAM data) noexcept;
inline void SelectComboItemByData(const wil::unique_hwnd& combo, LPARAM data) noexcept
{
    SelectComboItemByData(combo.get(), data);
}
[[nodiscard]] std::optional<LPARAM> TryGetSelectedComboItemData(HWND combo) noexcept;
[[nodiscard]] inline std::optional<LPARAM> TryGetSelectedComboItemData(const wil::unique_hwnd& combo) noexcept
{
    return TryGetSelectedComboItemData(combo.get());
}
void SetTwoStateToggleState(HWND toggle, bool highContrast, bool toggledOn) noexcept;
inline void SetTwoStateToggleState(const wil::unique_hwnd& toggle, bool highContrast, bool toggledOn) noexcept
{
    SetTwoStateToggleState(toggle.get(), highContrast, toggledOn);
}
[[nodiscard]] bool GetTwoStateToggleState(HWND toggle, bool highContrast) noexcept;
[[nodiscard]] inline bool GetTwoStateToggleState(const wil::unique_hwnd& toggle, bool highContrast) noexcept
{
    return GetTwoStateToggleState(toggle.get(), highContrast);
}
[[nodiscard]] std::optional<uint32_t> TryParseUInt32(std::wstring_view text) noexcept;
[[nodiscard]] std::optional<uint64_t> TryParseUInt64(std::wstring_view text) noexcept;
[[nodiscard]] bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept;

// Layout helper functions
void PositionControl(HWND hwnd, int x, int y, int width, int height) noexcept;
void PositionAndSetFont(HWND hwnd, HFONT font, int x, int y, int width, int height) noexcept;
void SetControlText(HWND hwnd, std::wstring_view text) noexcept;
[[nodiscard]] int CalculateCardHeight(int rowHeight, int titleHeight, int cardPaddingY, int cardGapY, int descHeight) noexcept;
void TryPushCard(std::vector<RECT>& cards, const RECT& card) noexcept;

// Schema-driven UI generation
[[nodiscard]] HWND CreateSchemaControl(HWND parent,
                                       const SettingsSchemaParser::SettingField& field,
                                       PreferencesDialogState& state,
                                       int x,
                                       int& y,
                                       int width,
                                       int margin,
                                       int gapY,
                                       HFONT font) noexcept;
} // namespace PrefsUi

namespace PrefsFile
{
[[nodiscard]] bool TryReadFileToString(const std::filesystem::path& path, std::string& out) noexcept;
[[nodiscard]] bool TryWriteFileFromString(const std::filesystem::path& path, std::string_view text) noexcept;
} // namespace PrefsFile

namespace PrefsListView
{
[[nodiscard]] int GetSingleLineRowHeightPx(HWND list, HDC hdc) noexcept;
[[nodiscard]] LRESULT
DrawThemedTwoColumnListRow(DRAWITEMSTRUCT* dis, PreferencesDialogState& state, HWND list, UINT expectedCtlId, bool secondColumnRightAlign) noexcept;
} // namespace PrefsListView

namespace PrefsPaneHost
{
[[nodiscard]] bool EnsureCreated(HWND pageHost, wil::unique_hwnd& paneHwnd) noexcept;
void ResizeToHostClient(HWND pageHost, HWND paneHwnd) noexcept;
void Show(HWND paneHwnd, bool visible) noexcept;
void ApplyScrollDelta(HWND pageHost, int dy) noexcept;
void ScrollTo(HWND pageHost, PreferencesDialogState& state, int newScrollY) noexcept;
void EnsureControlVisible(HWND pageHost, PreferencesDialogState& state, HWND control) noexcept;
} // namespace PrefsPaneHost

namespace PrefsInput
{
void CreateFramedComboBox(PreferencesDialogState& state, HWND parent, HWND& outFrame, HWND& outCombo, int controlId) noexcept;
void CreateFramedComboBox(PreferencesDialogState& state, HWND parent, wil::unique_hwnd& outFrame, wil::unique_hwnd& outCombo, int controlId) noexcept;
void CreateFramedEditBox(PreferencesDialogState& state, HWND parent, HWND& outFrame, HWND& outEdit, int controlId, DWORD style) noexcept;
void CreateFramedEditBox(
    PreferencesDialogState& state, HWND parent, wil::unique_hwnd& outFrame, wil::unique_hwnd& outEdit, int controlId, DWORD style) noexcept;
void EnableMouseWheelForwarding(HWND control) noexcept;
void EnableMouseWheelForwarding(const wil::unique_hwnd& control) noexcept;
} // namespace PrefsInput

namespace PrefsPlugins
{
void BuildListItems(std::vector<PrefsPluginListItem>& out) noexcept;
[[nodiscard]] std::wstring_view GetId(const PrefsPluginListItem& item) noexcept;
[[nodiscard]] std::wstring_view GetDisplayName(const PrefsPluginListItem& item) noexcept;
[[nodiscard]] std::wstring_view GetDescription(const PrefsPluginListItem& item) noexcept;
[[nodiscard]] std::wstring_view GetShortIdOrId(const PrefsPluginListItem& item) noexcept;
[[nodiscard]] bool IsLoadable(const PrefsPluginListItem& item) noexcept;
[[nodiscard]] int GetOriginOrder(const PrefsPluginListItem& item) noexcept;
} // namespace PrefsPlugins

namespace PrefsFolders
{
inline constexpr std::wstring_view kLeftPaneSlot  = L"left";
inline constexpr std::wstring_view kRightPaneSlot = L"right";

struct FolderPanePreferences
{
    Common::Settings::FolderDisplayMode display         = Common::Settings::FolderDisplayMode::Brief;
    Common::Settings::FolderSortBy sortBy               = Common::Settings::FolderSortBy::Name;
    Common::Settings::FolderSortDirection sortDirection = Common::Settings::FolderSortDirection::Ascending;
    bool statusBarVisible                               = true;
};

[[nodiscard]] FolderPanePreferences GetFolderPanePreferences(const Common::Settings::Settings& settings, std::wstring_view slot) noexcept;
[[nodiscard]] uint32_t GetFolderHistoryMax(const Common::Settings::Settings& settings) noexcept;
[[nodiscard]] bool AreEquivalentFolderPreferences(const Common::Settings::Settings& a, const Common::Settings::Settings& b) noexcept;
[[nodiscard]] Common::Settings::FolderSortDirection DefaultFolderSortDirection(Common::Settings::FolderSortBy sortBy) noexcept;

[[nodiscard]] Common::Settings::FoldersSettings* EnsureWorkingFoldersSettings(Common::Settings::Settings& settings) noexcept;
[[nodiscard]] Common::Settings::FolderPane* EnsureWorkingFolderPane(Common::Settings::Settings& settings, std::wstring_view slot) noexcept;
} // namespace PrefsFolders

namespace PrefsMonitor
{
[[nodiscard]] const Common::Settings::MonitorSettings& GetMonitorSettingsOrDefault(const Common::Settings::Settings& settings) noexcept;
[[nodiscard]] Common::Settings::MonitorSettings* EnsureWorkingMonitorSettings(Common::Settings::Settings& settings) noexcept;
} // namespace PrefsMonitor

namespace PrefsCache
{
[[nodiscard]] const Common::Settings::CacheSettings& GetCacheSettingsOrDefault(const Common::Settings::Settings& settings) noexcept;
[[nodiscard]] Common::Settings::CacheSettings* EnsureWorkingCacheSettings(Common::Settings::Settings& settings) noexcept;
void MaybeResetWorkingCacheSettingsIfEmpty(Common::Settings::Settings& settings) noexcept;
[[nodiscard]] std::optional<uint64_t> TryParseCacheBytes(std::wstring_view text) noexcept;
[[nodiscard]] std::wstring FormatCacheBytes(uint64_t bytes) noexcept;
} // namespace PrefsCache

namespace PrefsConnections
{
[[nodiscard]] const Common::Settings::ConnectionsSettings& GetConnectionsSettingsOrDefault(const Common::Settings::Settings& settings) noexcept;
[[nodiscard]] Common::Settings::ConnectionsSettings* EnsureWorkingConnectionsSettings(Common::Settings::Settings& settings) noexcept;
void MaybeResetWorkingConnectionsSettingsIfEmpty(Common::Settings::Settings& settings) noexcept;
} // namespace PrefsConnections

namespace PrefsFileOperations
{
[[nodiscard]] const Common::Settings::FileOperationsSettings& GetFileOperationsSettingsOrDefault(const Common::Settings::Settings& settings) noexcept;
[[nodiscard]] Common::Settings::FileOperationsSettings* EnsureWorkingFileOperationsSettings(Common::Settings::Settings& settings) noexcept;
void MaybeResetWorkingFileOperationsSettingsIfEmpty(Common::Settings::Settings& settings) noexcept;
} // namespace PrefsFileOperations
