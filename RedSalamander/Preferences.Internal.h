#pragma once

// Internal types shared across Preferences dialog implementation files.
// Keep this header private to Preferences translation units.

#include "framework.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <memory>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <vector>

#include "AppTheme.h"
#include "DxUi/DxUi.Typography.h"
#include "DxUi/DxUi.h"
#include "DxUiThemePalette.h"
#include "Helpers.h"
#include "PluginConfiguration.h"
#include "SettingsStore.h"
#include "UiMetrics.h"

// Window props used by Preferences UI controls.
inline constexpr wchar_t kPrefsVisuallyDisabledProp[] = L"RedSalamander.Preferences.VisuallyDisabled";

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
    CompareDirectories,
    HotPaths,
    FileOperations,
    UserMenu,
    Monitor,
};

inline constexpr size_t kPrefCategoryCount                  = static_cast<size_t>(PrefCategory::Monitor) + 1u;
inline constexpr std::wstring_view kPreferencesMonitorAppId = L"RedSalamanderMonitor";

[[nodiscard]] constexpr size_t PrefCategoryIndex(const PrefCategory category) noexcept
{
    return static_cast<size_t>(category);
}

class KeyboardPane;
struct PreferencesDialogState;

struct PreferencesTypographyContext
{
    UINT dpi = USER_DEFAULT_SCREEN_DPI;
    RedSalamander::DxUi::Typography::TypographySpec body;
    RedSalamander::DxUi::Typography::TypographySpec caption;
    RedSalamander::DxUi::Typography::TypographySpec title;
    RedSalamander::DxUi::Typography::TypographySpec strong;
};

[[nodiscard]] bool PrefsKeyboardCaptureWantsAllKeys(const PreferencesDialogState* state) noexcept;
[[nodiscard]] bool PrefsHandleKeyboardCaptureMessage(HWND hostHwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

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

enum class PreferencesDeferredActionKind : uint8_t
{
    ViewersSearchChanged,
    KeyboardSearchChanged,
    KeyboardScopeChanged,
    KeyboardAssign,
    KeyboardRemove,
    KeyboardReset,
    KeyboardImport,
    KeyboardExport,
    ThemesThemeChanged,
    ThemesBaseChanged,
    ThemesNameBlur,
    ThemesSearchChanged,
    PluginsSearchChanged,
    PluginsConfigure,
    PluginsTest,
    PluginsTestAll,
    FileOperationsBandwidthPresetChanged,
    CompareDirectoriesIgnoreToggleChanged,
};

struct PreferencesDeferredActionPayload
{
    PreferencesDeferredActionKind kind = PreferencesDeferredActionKind::ViewersSearchChanged;
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

void PrefsReorderPanelChildren(RedSalamander::DxUi::Panel* root, std::span<RedSalamander::DxUi::Control* const> orderedControls);

// Monitor filter mask bits for the Monitor Preferences page.
enum class MonitorFilterBit : uint32_t
{
    Text    = 0x01u,
    Error   = 0x02u,
    Warning = 0x04u,
    Info    = 0x08u,
    Perf    = 0x10u,
    Debug   = 0x20u,
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

struct PreferencesEmptyStateSpec
{
    wchar_t iconGlyph         = L'\0';
    wchar_t fallbackIconGlyph = L'\0';
    enum class Tone : uint8_t
    {
        Neutral,
        Accent,
        Warning,
        Error,
        Info,
    } tone = Tone::Neutral;
    std::wstring title;
    std::wstring body;
    std::wstring caption;
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

enum class PrefsPluginDetailsMessageKind : uint8_t
{
    None,
    Error,
    EmptyState,
};

enum class PrefsInlineMessageSeverity : uint8_t
{
    Info,
    Warning,
    Error,
};

using PrefsPluginConfigFieldType = Common::PluginConfiguration::FieldType;
using PrefsPluginConfigChoice    = Common::PluginConfiguration::Choice;
using PrefsPluginConfigField     = Common::PluginConfiguration::Field;

struct PrefsPluginConfigChoiceDxControl
{
    PrefsPluginConfigChoiceDxControl() = default;

    PrefsPluginConfigChoiceDxControl(const PrefsPluginConfigChoiceDxControl&)            = delete;
    PrefsPluginConfigChoiceDxControl& operator=(const PrefsPluginConfigChoiceDxControl&) = delete;

    PrefsPluginConfigChoiceDxControl(PrefsPluginConfigChoiceDxControl&&)            = default;
    PrefsPluginConfigChoiceDxControl& operator=(PrefsPluginConfigChoiceDxControl&&) = default;

    RedSalamander::DxUi::Checkbox* checkbox = nullptr;
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
    std::wstring retainedText;
    bool retainedToggleValue = false;
    std::wstring retainedOptionValue;
    std::vector<std::wstring> retainedSelectionValues;
    RedSalamander::DxUi::Label* dxLabelControl         = nullptr;
    RedSalamander::DxUi::Label* dxDescriptionControl   = nullptr;
    RedSalamander::DxUi::TextField* dxEditControl      = nullptr;
    RedSalamander::DxUi::Button* dxBrowseButtonControl = nullptr;
    RedSalamander::DxUi::ComboBox* dxComboControl      = nullptr;
    RedSalamander::DxUi::Toggle* dxToggleControl       = nullptr;
    size_t toggleOnChoiceIndex                         = 0;
    size_t toggleOffChoiceIndex                        = 0;
    std::vector<PrefsPluginConfigChoiceDxControl> dxChoiceControls;
};

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
    Common::Settings::Settings monitorBaselineSettings;
    Common::Settings::Settings workingMonitorSettings;

    bool dirty                   = false;
    bool appliedOnce             = false;
    bool staleFromExternalReload = false;

    // Navigation State
    PrefCategory currentCategory = PrefCategory::General;
    PrefCategory initialCategory = PrefCategory::General;
    std::optional<PrefsPluginListItem> pluginsSelectedPlugin;
    std::wstring pluginsSelectedPluginId;
    std::wstring pluginsRetainedSelectedPluginId;
    bool pluginsDetailsActive = false;
    std::wstring viewersSearchText;
    std::wstring viewersSelectedExtensionText;
    RedSalamander::DxUi::GridSortSpec viewersListSortSpec{};
    std::wstring keyboardSearchText;
    std::wstring pluginsSearchText;
    std::wstring pluginsSelectedCustomPathText;
    std::wstring themesSearchText;
    std::wstring themesSelectedColorKey;

    // Layout and Sizing
    int categoryListWidthPx = 0;
    SIZE minTrackSizePx{};
    SIZE restoreMinSizePx{};

    int pageScrollY             = 0;
    int pageScrollMaxY          = 0;
    int pageWheelDeltaRemainder = 0;
#ifdef ENABLE_TESTS
    uint64_t pageHostScrollRequestCount             = 0u;
    uint64_t pageHostScrollCoalescedRequestCount    = 0u;
    uint64_t pageHostScrollApplyCount               = 0u;
    uint64_t pageHostScrollMovedChildCountTotal     = 0u;
    uint64_t pageHostDxScrollMovedControlCountTotal = 0u;
    uint64_t pageHostDxScrollLastMovedControlCount  = 0u;
    uint64_t pageHostScrollLastApplyUs              = 0u;
#endif
    int pageHostPendingScrollY      = 0;
    bool pageHostScrollApplyPending = false;
    bool pageHostSyncingScrollPanel = false;
    std::array<int, kPrefCategoryCount> retainedPageScrollYByCategory{};
    int pageHostDirectContentBottomPx = 0;
    std::vector<RECT> pageSettingCards;
    bool pageHostRelayoutInProgress = false;
    bool pageHostIgnoreSize         = false;
    bool updatingPageText           = false;
    std::array<bool, kPrefCategoryCount> paneFirstCreateDone{};
    std::array<RedSalamander::DxUi::Panel*, kPrefCategoryCount> paneWrapperPanels{};

#ifdef ENABLE_TESTS
    bool debugLastWheelRouteSeen                      = false;
    bool debugLastWheelRouteForwarded                 = false;
    bool debugLastWheelRouteTargetWasPageHost         = false;
    bool debugLastWheelRouteTargetWasCategoryTree     = false;
    bool debugLastWheelRouteTargetHadVerticalScroll   = false;
    bool debugLastWheelWindowFromPointWasPageHost     = false;
    bool debugLastWheelWindowFromPointWasCategoryTree = false;
    bool debugLastWheelWndProcSeen                    = false;
    bool debugLastWheelDxHandled                      = false;
    bool debugLastWheelFallbackCalled                 = false;
    bool debugLastWheelFallbackHandled                = false;
    int debugLastWheelDelta                           = 0;
    int debugLastWheelClientX                         = 0;
    int debugLastWheelClientY                         = 0;
    int debugLastWheelBeforeY                         = 0;
    int debugLastWheelBeforeMaxY                      = 0;
    int debugLastWheelAfterY                          = 0;
    int debugLastWheelAfterMaxY                       = 0;
#endif

    // Theme Resources (RAII-managed)
    wil::unique_hbrush backgroundBrush;
    wil::unique_hbrush cardBrush;
    COLORREF cardBackgroundColor = RGB(255, 255, 255);
    wil::unique_hbrush inputBrush;
    COLORREF inputBackgroundColor = RGB(255, 255, 255);
    wil::unique_hbrush inputFocusedBrush;
    COLORREF inputFocusedBackgroundColor = RGB(255, 255, 255);
    wil::unique_hbrush inputDisabledBrush;
    COLORREF inputDisabledBackgroundColor = RGB(255, 255, 255);
    // Dialog Structure Controls
    HWND categoryTreeWindow                                        = nullptr;
    bool categoryTreeUsesDxUi                                      = false;
    bool pageHostUsesDxUi                                          = false;
    RedSalamander::DxUi::WindowHost* pageHostDxHost                = nullptr;
    RedSalamander::DxUi::Panel* pageHostDxRootControl              = nullptr;
    RedSalamander::DxUi::ScrollPanel* pageHostDxScrollPanelControl = nullptr;
    RedSalamander::DxUi::Panel* pageHostDxContentRootControl       = nullptr;
    RedSalamander::DxUi::Control* pageHostDxNoteControl            = nullptr;
    HWND pageHostWindow                                            = nullptr;

    std::vector<std::wstring> viewersExtensionKeys;
    std::vector<ViewerPluginOption> viewersPluginOptions;

    // Keyboard Page Controls
    KeyboardPane* keyboardPaneOwner = nullptr;
    std::wstring keyboardHintText;

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
    ShortcutScope keyboardSelectedScope = ShortcutScope::FunctionBar;
    std::wstring keyboardSelectedCommandId;
    std::optional<size_t> keyboardSelectedBindingIndex;

    // Themes Page Controls (RAII-managed)
    std::wstring themesNoteText;
    std::wstring themesNameText;
    std::wstring themesKeyText;
    std::wstring themesColorText;

    std::vector<ThemeComboItem> themeComboItems;
    std::vector<Common::Settings::ThemeDefinition> themeFileThemes;

    // Plugins details subpage (when a plugin tree child is selected). (RAII-managed)
    std::wstring pluginsDetailsIdText;
    std::wstring pluginsDetailsConfigErrorText;
    std::wstring pluginsDetailsConfigEmptyStateText;
    PrefsPluginDetailsMessageKind pluginsDetailsMessageKind = PrefsPluginDetailsMessageKind::None;
    std::wstring pluginsStatusTitleText;
    std::wstring pluginsStatusBodyText;
    PrefsInlineMessageSeverity pluginsStatusSeverity = PrefsInlineMessageSeverity::Info;
    std::wstring pluginsDetailsConfigPluginId;
    std::string pluginsDetailsConfigSourceJsonUtf8;
    std::vector<PrefsPluginConfigFieldControls> pluginsDetailsConfigFields;
    RedSalamander::DxUi::Panel* pluginsDetailsConfigDxPanel = nullptr;

    std::vector<PrefsPluginListItem> pluginsListItems;

    // Hot Paths page: no Win32 intermediary controls; DxUi writes directly to workingSettings.

    // Refresh State Flags
    bool previewApplied        = false;
    bool refreshingPanesPage   = false;
    bool refreshingThemesPage  = false;
    bool refreshingPluginsPage = false;
};

void SetDirty(HWND dlg, PreferencesDialogState& state) noexcept;

// Shared debug diagnostics helpers for DxUi host surfaces.
[[nodiscard]] inline UINT GetDxHostDebugWidthPx(const RedSalamander::DxUi::WindowHost* host) noexcept
{
#ifdef ENABLE_TESTS
    return host ? host->DebugGetWidthPx() : 0u;
#else
    static_cast<void>(host);
    return 0u;
#endif
}

[[nodiscard]] inline UINT GetDxHostDebugHeightPx(const RedSalamander::DxUi::WindowHost* host) noexcept
{
#ifdef ENABLE_TESTS
    return host ? host->DebugGetHeightPx() : 0u;
#else
    static_cast<void>(host);
    return 0u;
#endif
}

[[nodiscard]] inline uint64_t GetDxHostDebugRenderCount(const RedSalamander::DxUi::WindowHost* host) noexcept
{
#ifdef ENABLE_TESTS
    return host ? host->DebugGetRenderCount() : 0u;
#else
    static_cast<void>(host);
    return 0u;
#endif
}

[[nodiscard]] inline uint64_t GetDxHostDebugResizeCount(const RedSalamander::DxUi::WindowHost* host) noexcept
{
#ifdef ENABLE_TESTS
    return host ? host->DebugGetResizeCount() : 0u;
#else
    static_cast<void>(host);
    return 0u;
#endif
}

[[nodiscard]] inline uint64_t GetDxHostDebugResizeFailureCount(const RedSalamander::DxUi::WindowHost* host) noexcept
{
#ifdef ENABLE_TESTS
    return host ? host->DebugGetResizeFailureCount() : 0u;
#else
    static_cast<void>(host);
    return 0u;
#endif
}

// Reference overloads for non-nullable host references.
[[nodiscard]] inline UINT GetDxHostDebugWidthPx(const RedSalamander::DxUi::WindowHost& host) noexcept
{
    return GetDxHostDebugWidthPx(&host);
}
[[nodiscard]] inline UINT GetDxHostDebugHeightPx(const RedSalamander::DxUi::WindowHost& host) noexcept
{
    return GetDxHostDebugHeightPx(&host);
}
[[nodiscard]] inline uint64_t GetDxHostDebugRenderCount(const RedSalamander::DxUi::WindowHost& host) noexcept
{
    return GetDxHostDebugRenderCount(&host);
}
[[nodiscard]] inline uint64_t GetDxHostDebugResizeCount(const RedSalamander::DxUi::WindowHost& host) noexcept
{
    return GetDxHostDebugResizeCount(&host);
}
[[nodiscard]] inline uint64_t GetDxHostDebugResizeFailureCount(const RedSalamander::DxUi::WindowHost& host) noexcept
{
    return GetDxHostDebugResizeFailureCount(&host);
}

namespace PrefsUi
{
using Win32Text::GetWindowTextString;
[[nodiscard]] PreferencesTypographyContext MakeTypographyContext(HWND hwnd) noexcept;
[[nodiscard]] int MeasureSingleLineTextWidthPx(const PreferencesTypographyContext& typography,
                                               const RedSalamander::DxUi::Typography::TypographySpec& spec,
                                               std::wstring_view text) noexcept;
[[nodiscard]] int MeasureWrappedTextHeightPx(const PreferencesTypographyContext& typography,
                                             const RedSalamander::DxUi::Typography::TypographySpec& spec,
                                             int width,
                                             std::wstring_view text) noexcept;
[[nodiscard]] std::wstring_view TrimWhitespace(std::wstring_view text) noexcept;
[[nodiscard]] bool ContainsCaseInsensitive(std::wstring_view haystack, std::wstring_view needle) noexcept;
void InvalidateComboBox(HWND combo) noexcept;
[[nodiscard]] PreferencesDialogState* GetDialogState(HWND childOrDialog) noexcept;
[[nodiscard]] bool PostDeferredAction(HWND hwnd, PreferencesDeferredActionKind kind) noexcept;
[[nodiscard]] std::optional<uint32_t> TryParseUInt32(std::wstring_view text) noexcept;
[[nodiscard]] std::optional<uint64_t> TryParseUInt64(std::wstring_view text) noexcept;
[[nodiscard]] bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept;

inline constexpr wchar_t kPrefsTreeRedrawBlockProp[] = L"RedSalamander.Preferences.TreeRedrawBlock";

void TryPushCard(std::vector<RECT>& cards, const RECT& card) noexcept;
void HideSharedPageEmptyState(PreferencesDialogState& state) noexcept;
[[nodiscard]] int ShowSharedPageEmptyState(HWND host,
                                           PreferencesDialogState& state,
                                           const PreferencesEmptyStateSpec& spec,
                                           int x,
                                           int y,
                                           int width,
                                           const PreferencesTypographyContext& typography) noexcept;

[[nodiscard]] bool IsActuallyVisibleChildWindow(HWND hwnd) noexcept;
[[nodiscard]] inline bool HasRetainedDxChildren(const RedSalamander::DxUi::Panel* root) noexcept
{
    return root != nullptr && ! root->GetChildren().empty();
}
} // namespace PrefsUi

namespace PrefsDxHost
{
[[nodiscard]] bool Attach(HWND hwnd, RedSalamander::DxUi::WindowHost& host) noexcept;
void ResetOwnedHostWindow(wil::unique_hwnd& hwnd) noexcept;
#ifdef ENABLE_TESTS
[[nodiscard]] size_t CountVisibleRenderedHosts(HWND parent) noexcept;
[[nodiscard]] size_t CountVisibleHostsWithResizeFailures(HWND parent) noexcept;
[[nodiscard]] uint64_t SumVisibleRenderedHostRenderCounts(HWND parent) noexcept;
[[nodiscard]] bool TryGetDirectHostMetrics(HWND hwnd, size_t& visibleHostCount, size_t& resizeFailureCount, uint64_t& renderCountTotal) noexcept;
#endif
} // namespace PrefsDxHost

namespace PrefsFile
{
[[nodiscard]] bool TryReadFileToString(const std::filesystem::path& path, std::string& out) noexcept;
[[nodiscard]] bool TryWriteFileFromString(const std::filesystem::path& path, std::string_view text) noexcept;
} // namespace PrefsFile

namespace PrefsPageHost
{
void ApplyScrollDelta(HWND pageHostWindow, int dy, bool syncScrollPanel = true) noexcept;
void ScrollTo(HWND pageHostWindow, PreferencesDialogState& state, int newScrollY) noexcept;
void RequestScrollTo(HWND pageHostWindow, PreferencesDialogState& state, int newScrollY) noexcept;
void FlushPendingScroll(HWND pageHostWindow, PreferencesDialogState& state) noexcept;
void EnsureControlVisible(HWND pageHostWindow, PreferencesDialogState& state, HWND control) noexcept;
} // namespace PrefsPageHost

namespace PrefsPlugins
{
void BuildListItems(std::vector<PrefsPluginListItem>& out) noexcept;
[[nodiscard]] std::optional<PrefsPluginListItem> FindItemById(std::wstring_view pluginId) noexcept;
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
[[nodiscard]] bool GetFolderShowHiddenFiles(const Common::Settings::Settings& settings) noexcept;
[[nodiscard]] bool GetFolderShowSystemFiles(const Common::Settings::Settings& settings) noexcept;
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

namespace PrefsCompareDirectories
{
[[nodiscard]] const Common::Settings::CompareDirectoriesSettings& GetCompareDirectoriesSettingsOrDefault(const Common::Settings::Settings& settings) noexcept;
[[nodiscard]] Common::Settings::CompareDirectoriesSettings* EnsureWorkingCompareDirectoriesSettings(Common::Settings::Settings& settings) noexcept;
void MaybeResetWorkingCompareDirectoriesSettingsIfEmpty(Common::Settings::Settings& settings) noexcept;
} // namespace PrefsCompareDirectories

namespace PrefsHotPaths
{
[[nodiscard]] const Common::Settings::HotPathsSettings& GetHotPathsSettingsOrDefault(const Common::Settings::Settings& settings) noexcept;
[[nodiscard]] Common::Settings::HotPathsSettings* EnsureWorkingHotPathsSettings(Common::Settings::Settings& settings) noexcept;
void MaybeResetWorkingHotPathsSettingsIfEmpty(Common::Settings::Settings& settings) noexcept;
} // namespace PrefsHotPaths
