#pragma once

#include <filesystem>
#include <string>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#pragma warning(pop)

#include "AppTheme.h"
#include "SettingsStore.h"

struct IFileSystem;
class FolderWindow;
class ShortcutManager;

#ifdef ENABLE_TESTS
enum class CompareDirectoriesOptionsDebugFocusTarget
{
    None,
    CompareSubdirectoriesToggle,
    CompareSizeToggle,
    CompareDateTimeToggle,
    CompareAttributesToggle,
    CompareContentToggle,
    CompareSubdirAttributesToggle,
    SelectSubdirsOnlyInOnePaneToggle,
    KeepIdenticalItemsToggle,
    IgnoreFilesToggle,
    IgnoreFilesEdit,
    IgnoreDirectoriesToggle,
    IgnoreDirectoriesEdit,
    OkButton,
    CancelButton,
};

struct CompareDirectoriesOptionsDebugSnapshot
{
    bool optionsDialogVisible                             = false;
    bool optionsUsesDxUiStatics                           = false;
    bool optionsUsesDxUiButtons                           = false;
    bool optionsUsesDxUiToggles                           = false;
    bool optionsUsesDxUiEdits                             = false;
    bool usesDxUiTypographyMetrics                        = false;
    bool optionsReloadParticipantRegistered               = false;
    bool optionsStaleFromExternalReload                   = false;
    bool optionsDialogDirty                               = false;
    bool compareSubdirectoriesChecked                     = false;
    bool themeDark                                        = false;
    bool themeHighContrast                                = false;
    bool themeRainbow                                     = false;
    size_t visibleLegacyStaticCount                       = 0u;
    size_t visibleLegacyFooterButtonCount                 = 0u;
    size_t visibleLegacyToggleCount                       = 0u;
    size_t visibleLegacyEditCount                         = 0u;
    size_t visibleNativeBodyControlCount                  = 0u;
    size_t hiddenLegacyOwnerDrawFooterButtonCount         = 0u;
    size_t hiddenLegacyOwnerDrawToggleCount               = 0u;
    size_t visibleBodyRenderedDxHostCount                 = 0u;
    size_t visibleDxBodyHeaderCount                       = 0u;
    size_t visibleDxBodyCardCount                         = 0u;
    size_t dxFooterButtonHostCount                        = 0u;
    size_t visibleDxFooterButtonHostCount                 = 0u;
    size_t bodyDxHostResizeFailureCount                   = 0u;
    size_t bodyDxHostPresentFailureCount                  = 0u;
    int bodyDxHostWidth                                   = 0;
    int bodyDxHostHeight                                  = 0;
    int bodyContentHeight                                 = 0;
    int bodyScrollOffset                                  = 0;
    int bodyScrollMax                                     = 0;
    int okFooterAttachFailureStage                        = 0;
    int cancelFooterAttachFailureStage                    = 0;
    int maxVisibleDxToggleHostWidth                       = 0;
    int maxDxToggleTextOverlapWidth                       = 0;
    bool bodyUsesTwoColumns                               = false;
    CompareDirectoriesOptionsDebugFocusTarget focusTarget = CompareDirectoriesOptionsDebugFocusTarget::None;
};

struct CompareDirectoriesRunDebugSnapshot
{
    bool windowVisible                    = false;
    bool optionsDialogVisible             = false;
    bool compareStarted                   = false;
    bool compareActive                    = false;
    bool compareRunPending                = false;
    bool compareRunSawScanProgress        = false;
    bool leaveScopePromptPending          = false;
    bool usesDxUiMenuBar                  = false;
    bool usesDxUiBannerButtons            = false;
    bool usesDxUiBannerText               = false;
    bool hasNativeUiFontState             = false;
    bool nativeMenuAttached               = false;
    bool bannerOptionsEnabled             = false;
    bool bannerRescanEnabled              = false;
    uint32_t scanActiveScans              = 0;
    uint64_t scanFolderCount              = 0;
    uint64_t scanEntryCount               = 0;
    uint64_t contentPendingCompares       = 0;
    uint64_t contentCompletedCompares     = 0;
    uint64_t contentTotalCompares         = 0;
    uint64_t dxMenuBarRenderCount         = 0;
    size_t menuBarItemCount               = 0u;
    size_t visibleDxMenuBarHostCount      = 0u;
    size_t visibleDxBannerButtonHostCount = 0u;
    size_t visibleDxBannerTextHostCount   = 0u;
    size_t visibleLegacyBannerButtonCount = 0u;
    size_t visibleLegacyBannerTextCount   = 0u;
    size_t leftPaneItemCount              = 0u;
    size_t rightPaneItemCount             = 0u;
    size_t leftPaneSelectedCount          = 0u;
    size_t rightPaneSelectedCount         = 0u;
    std::wstring leftPanePluginPath;
    std::wstring rightPanePluginPath;
    std::wstring leftPaneEmptyStateMessage;
    std::wstring rightPaneEmptyStateMessage;
};
#endif

struct CompareDirectoriesPaneContext
{
    std::wstring pluginId;
    std::wstring instanceContext;
    std::filesystem::path rootPluginPath;
};

[[nodiscard]] bool ShowCompareDirectoriesWindow(HWND owner,
                                                FolderWindow& applicationFolderWindow,
                                                Common::Settings::Settings& settings,
                                                const AppTheme& theme,
                                                const ShortcutManager* shortcuts,
                                                CompareDirectoriesPaneContext left,
                                                CompareDirectoriesPaneContext right) noexcept;

void UpdateCompareDirectoriesWindowsTheme(const AppTheme& theme) noexcept;

[[nodiscard]] HWND GetCompareDirectoriesWindowHandle() noexcept;

#ifdef ENABLE_TESTS
[[nodiscard]] bool DebugGetCompareDirectoriesOptionsSnapshot(CompareDirectoriesOptionsDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugGetCompareDirectoriesOptionsSnapshotForWindow(HWND compareWindow, CompareDirectoriesOptionsDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugGetCompareDirectoriesRunSnapshot(CompareDirectoriesRunDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugGetCompareDirectoriesRunSnapshotForWindow(HWND compareWindow, CompareDirectoriesRunDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetCompareDirectoriesPanePathForWindow(HWND compareWindow, bool leftPane, const std::filesystem::path& path) noexcept;
[[nodiscard]] bool DebugSetCompareDirectoriesRunPendingForWindow(HWND compareWindow, bool pending) noexcept;
void DebugFailNextCompareDirectoriesWindowCreate() noexcept;
[[nodiscard]] bool DebugGetCompareDirectoriesMenuBarItemLabel(size_t index, std::wstring& outText) noexcept;
[[nodiscard]] bool DebugGetCompareDirectoriesMenuBarItemScreenRect(size_t index, RECT& outRect) noexcept;
[[nodiscard]] bool DebugFocusCompareDirectoriesOptionsFirstControl() noexcept;
[[nodiscard]] bool DebugFocusCompareDirectoriesOptionsTarget(CompareDirectoriesOptionsDebugFocusTarget target) noexcept;
[[nodiscard]] bool DebugFocusCompareDirectoriesOptionsTargetForWindow(HWND compareWindow, CompareDirectoriesOptionsDebugFocusTarget target) noexcept;
[[nodiscard]] bool DebugSetCompareDirectoriesOptionsIgnoreFilesEnabled(bool enabled) noexcept;
[[nodiscard]] bool DebugSetCompareDirectoriesOptionsIgnoreDirectoriesEnabled(bool enabled) noexcept;
[[nodiscard]] HWND DebugGetCompareDirectoriesOptionsDialogHandle() noexcept;
[[nodiscard]] bool DebugScrollCompareDirectoriesOptionsBodyPages(int pageDelta) noexcept;
[[nodiscard]] bool DebugScrollCompareDirectoriesOptionsBodyPagesForWindow(HWND compareWindow, int pageDelta) noexcept;
[[nodiscard]] bool DebugGetCompareDirectoriesOptionsTargetHostAndClientRect(CompareDirectoriesOptionsDebugFocusTarget target,
                                                                            HWND& outHost,
                                                                            RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetCompareDirectoriesOptionsTargetHostAndClientRectForWindow(HWND compareWindow,
                                                                                     CompareDirectoriesOptionsDebugFocusTarget target,
                                                                                     HWND& outHost,
                                                                                     RECT& outRect) noexcept;
#endif
