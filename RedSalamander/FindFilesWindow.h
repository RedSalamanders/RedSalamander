#pragma once

#include <cstdint>
#include <filesystem>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#pragma warning(pop)

#include "AppTheme.h"
#include "SettingsStore.h"

struct IFileSystem;

struct FindFilesPaneContext
{
    wil::com_ptr<IFileSystem> fileSystem;
    std::wstring pluginId;
    std::wstring pluginShortId;
    std::wstring instanceContext;
    std::filesystem::path rootPluginPath;
};

[[nodiscard]] bool ShowFindFilesWindow(HWND owner, Common::Settings::Settings& settings, const AppTheme& theme, FindFilesPaneContext context) noexcept;

void UpdateFindFilesWindowsTheme(const AppTheme& theme) noexcept;

[[nodiscard]] HWND GetFindFilesWindowHandle() noexcept;
[[nodiscard]] bool IsFindFilesWindowHandle(HWND hwnd) noexcept;

#ifdef ENABLE_TESTS
[[nodiscard]] size_t DebugGetFindFilesWindowCount() noexcept;
[[nodiscard]] std::wstring DebugMakeFindFilesResultKeyForTests(std::wstring_view pluginId,
                                                               std::wstring_view instanceContext,
                                                               std::wstring_view fullPath) noexcept;

struct FindFilesDebugSourceOutcome final
{
    size_t sourceIndex = 0;
    HRESULT status     = E_PENDING;
};
[[nodiscard]] std::vector<size_t> DebugSelectKnownCompletedFindFilesSourceIndicesForTests(
    size_t sourceCount, std::span<const FindFilesDebugSourceOutcome> outcomes, HRESULT overallStatus);

enum class FindFilesDebugOperation : uint8_t
{
    Find,
    Append,
    Intersect,
    Subtract,
};

enum class FindFilesDebugFocusTarget : uint8_t
{
    None,
    RootCombo,
    NameCombo,
    NameModeCombo,
    ContentCombo,
    ContentModeCombo,
    RecursiveCheck,
    IncludeFilesCheck,
    IncludeDirectoriesCheck,
    FollowSymlinksCheck,
    MatchCaseNameCheck,
    MatchCaseContentCheck,
    PreferIndexCheck,
    WantSnippetsCheck,
    FindButton,
    AppendButton,
    IntersectButton,
    SubtractButton,
    CancelButton,
    OpenButton,
    ParentButton,
    HelpButton,
    ResultsGrid,
};

enum class FindFilesDebugOpenDisposition : uint8_t
{
    None,
    NavigateToResult,
    NavigateToParent,
    NavigateToParentAndOpen,
    DefaultOpenFile,
};

struct FindFilesDebugSnapshot
{
    bool searchActive                                      = false;
    bool usesDxUiHost                                      = false;
    bool findButtonEnabled                                 = false;
    bool findButtonPressed                                 = false;
    bool appendButtonEnabled                               = false;
    bool intersectButtonEnabled                            = false;
    bool subtractButtonEnabled                             = false;
    bool cancelButtonEnabled                               = false;
    bool openButtonEnabled                                 = false;
    bool parentButtonEnabled                               = false;
    bool helpButtonEnabled                                 = false;
    bool rootComboEnabled                                  = false;
    bool nameComboEnabled                                  = false;
    bool nameModeComboEnabled                              = false;
    bool contentComboEnabled                               = false;
    bool contentModeComboEnabled                           = false;
    bool matchCaseContentEnabled                           = false;
    bool preferIndexEnabled                                = false;
    bool preferIndexChecked                                = false;
    bool wantSnippetsEnabled                               = false;
    bool recursiveChecked                                  = false;
    size_t resultCount                                     = 0;
    size_t selectedResultCount                             = 0;
    size_t visibleChildWindowCount                         = 0;
    bool hasStatusStrip                                    = false;
    bool statusStripVisible                                = false;
    bool statusStripBlendsWithWindowBackground             = false;
    uint32_t statusStripSectionCount                       = 0u;
    float statusStripHeightDip                             = 0.0f;
    bool rootNavigationVisible                             = false;
    D2D1_RECT_F rootNavigationRect                         = D2D1::RectF();
    bool rootNavigationEmbedded                            = false;
    bool rootNavigationEditMode                            = false;
    bool rootNavigationHasWin32Focus                       = false;
    HWND rootNavigationHwnd                                = nullptr;
    HWND rootNavigationEditHostHwnd                        = nullptr;
    HWND rootNavigationEditInputHwnd                       = nullptr;
    size_t rootNavigationHistoryCount                      = 0u;
    bool destinationNavigationVisible                      = false;
    D2D1_RECT_F destinationNavigationRect                  = D2D1::RectF();
    bool destinationNavigationEmbedded                     = false;
    bool destinationNavigationEditMode                     = false;
    bool destinationNavigationMenuHovered                  = false;
    bool destinationNavigationHistoryHovered               = false;
    bool destinationNavigationDiskHovered                  = false;
    int destinationNavigationHoveredSegmentIndex           = -1;
    int destinationNavigationHoveredSeparatorIndex         = -1;
    size_t destinationNavigationHistoryCount               = 0u;
    bool destinationNavigationHistoryDropdownVisible       = false;
    uint64_t destinationNavigationHistoryDropdownOpenCount = 0u;
    RECT destinationNavigationHistoryRect                  = {};
    bool rootPopupOpen                                     = false;
    bool nameModePopupOpen                                 = false;
    bool contentModePopupOpen                              = false;
    bool hasWin32Focus                                     = false;
    bool isForegroundWindow                                = false;
    std::optional<size_t> nameModeSelectedIndex;
    std::optional<size_t> contentModeSelectedIndex;
    FindFilesDebugFocusTarget focusTarget              = FindFilesDebugFocusTarget::None;
    HRESULT lastStatusHint                             = S_OK;
    uint32_t warningFlags                              = 0;
    uint32_t backend                                   = 0u;
    uint32_t phase                                     = 0u;
    bool hasServiceStatus                              = false;
    uint32_t resultListFullRebuildCount                = 0;
    uint32_t incrementalResultRefreshCount             = 0;
    uint32_t incrementalVisibleResultRefreshCount      = 0;
    uint64_t debugResultActionFocusRestoreRequestCount = 0u;
    size_t visibleResultRowCount                       = 0u;
    size_t visibleResultColumnCount                    = 0u;
    size_t visibleResultCellCount                      = 0u;
    uint64_t visibleResultIconCellCount                = 0u;
    bool resultListHasVerticalScrollbar                = false;
    bool themeCompactMode                              = false;
    bool themeDark                                     = false;
    bool themeHighContrast                             = false;
    bool themeRainbow                                  = false;
    uint64_t dxRenderCount                             = 0u;
    uint64_t resultGridPaintCount                      = 0u;
    uint64_t dxResizeCount                             = 0u;
    uint64_t dxResizeFailureCount                      = 0u;
    float debugResizeBeforeWidthDip                    = 0.0f;
    float debugResizeTargetWidthDip                    = 0.0f;
    float debugResizeObservedWidthDip                  = 0.0f;
    float debugSettingsFirstWidthDip                   = 0.0f;
    bool debugResizeSucceeded                          = false;
    FindFilesDebugFocusTarget debugLastSetComboTarget  = FindFilesDebugFocusTarget::None;
    D2D1_RECT_F firstResultHeaderRect                  = D2D1::RectF();
    D2D1_RECT_F secondResultHeaderRect                 = D2D1::RectF();
    D2D1_RECT_F selectedResultRowRect                  = D2D1::RectF();
    uint32_t selectedResultRowFillArgb                 = 0u;
    uint32_t selectedResultRowTextArgb                 = 0u;
    bool selectedResultRowUsesRainbow                  = false;
    bool resultsGridFolderViewMode                     = false;
    std::vector<std::wstring> resultColumnIds;
    std::vector<float> resultColumnWidthsDip;
    std::vector<std::wstring> fullPaths;
    std::vector<std::wstring> resultPathTexts;
    std::vector<int> resultIconIndices;
    std::wstring selectedResultFullPath;
    std::wstring rootText;
    std::wstring namePatternText;
    std::wstring contentPatternText;
    std::wstring builtRootText;
    std::wstring builtNamePatternText;
    std::wstring builtContentPatternText;
    std::wstring beginRootText;
    std::wstring beginNamePatternText;
    std::wstring beginContentPatternText;
    std::wstring debugStartRootText;
    std::wstring debugStartNamePatternText;
    std::wstring debugStartContentPatternText;
    std::wstring debugLastSetComboRequestedText;
    std::wstring debugLastSetComboObservedText;
    std::wstring submittedRootText;
    std::wstring submittedNamePatternText;
    std::wstring submittedContentPatternText;
    std::wstring statusText;
    std::wstring destinationStatusText;
    std::wstring rootNavigationText;
    std::wstring destinationNavigationText;
    std::wstring backendStatusText;
};

struct FindFilesDebugGridHit
{
    uint32_t zone       = 0u;
    size_t columnIndex  = 0u;
    D2D1_RECT_F rectDip = D2D1::RectF();
    bool isHeaderResize = false;
};

[[nodiscard]] bool DebugConfigureFindFilesWindow(std::wstring rootPath,
                                                 std::wstring namePattern,
                                                 std::wstring contentPattern,
                                                 Common::Settings::SearchNameMode nameMode,
                                                 Common::Settings::SearchContentMode contentMode) noexcept;

[[nodiscard]] bool DebugSetFindFilesWindowOptions(bool recursive, bool includeFiles, bool includeDirectories, bool preferIndex, bool wantSnippets) noexcept;

[[nodiscard]] bool DebugSetFindFilesWindowComboText(FindFilesDebugFocusTarget target, std::wstring text) noexcept;
[[nodiscard]] bool DebugSetFindFilesWindowDestinationPath(std::wstring path) noexcept;

[[nodiscard]] bool DebugStartFindFilesWindowSearch(FindFilesDebugOperation operation) noexcept;

[[nodiscard]] bool DebugCancelFindFilesWindowSearch() noexcept;
[[nodiscard]] bool DebugPostFindFilesWindowStaleSearchPayloads(std::wstring fullPath) noexcept;

[[nodiscard]] bool DebugGetFindFilesWindowSnapshot(FindFilesDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugHitTestFindFilesWindowResultsGrid(float xDip, float yDip, FindFilesDebugGridHit& out) noexcept;

[[nodiscard]] bool DebugFocusFindFilesWindowTarget(FindFilesDebugFocusTarget target) noexcept;
[[nodiscard]] bool DebugGetFindFilesWindowTargetClientRect(FindFilesDebugFocusTarget target, RECT& outRect) noexcept;

[[nodiscard]] bool DebugSetFindFilesWindowResultSort(size_t columnIndex, bool descending) noexcept;
[[nodiscard]] bool DebugReorderFindFilesWindowVisibleResultColumn(size_t fromVisibleIndex, size_t targetVisibleIndex) noexcept;
[[nodiscard]] bool DebugResizeFindFilesWindowVisibleResultColumn(size_t visibleIndex, float deltaDip) noexcept;
[[nodiscard]] bool DebugApplyFindFilesWindowResultsLayoutFromSettings() noexcept;

[[nodiscard]] bool DebugSelectFindFilesWindowResult(std::wstring fullPath) noexcept;
[[nodiscard]] bool DebugSelectFindFilesWindowResults(std::vector<std::wstring> fullPaths) noexcept;

[[nodiscard]] bool DebugActivateSelectedFindFilesWindowResult() noexcept;

[[nodiscard]] bool DebugOpenSelectedFindFilesWindowResultParent() noexcept;
[[nodiscard]] bool DebugGetSelectedFindFilesWindowOpenDisposition(bool parentOnly, FindFilesDebugOpenDisposition& out) noexcept;

[[nodiscard]] bool DebugWaitForFindFilesWindowIdle(uint32_t timeoutMs) noexcept;

[[nodiscard]] bool DebugScrollFindFilesWindowResultsByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugFindFilesIsNextQueuedMessage(HWND targetHwnd, UINT targetMessage) noexcept;
void DebugConfigureFindFilesWindowSearchRunBlocker(bool enabled) noexcept;
void DebugReleaseFindFilesWindowSearchRunBlocker() noexcept;
[[nodiscard]] bool DebugWaitForFindFilesWindowSearchRunBlocked(uint32_t timeoutMs) noexcept;
#endif
