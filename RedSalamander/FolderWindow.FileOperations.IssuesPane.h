#pragma once

#include "FolderWindow.h"

#include <array>
#include <cstdint>
#include <limits>
#include <memory>
#include <string_view>

#include <DxUi/DxUi.h>

#ifdef ENABLE_TESTS
#include "SelfTest/FileOperations/FileOperations.Tests.FileOperationsSurfaceResources.h"
#endif

class FileOperationsIssuesPane final
{
public:
    static HWND Create(FolderWindow::FileOperationState* fileOps, FolderWindow* folderWindow, HWND ownerWindow, std::weak_ptr<void> hostLifetime) noexcept;
    [[nodiscard]] static bool SelectTask(HWND hwnd, uint64_t taskId, uint64_t promptGeneration = 0u) noexcept;

#ifdef ENABLE_TESTS
    struct SelfTestSnapshot
    {
        FileOperationsSelfTest::SurfaceResources surfaceResources{};
        bool detailPresent                   = false;
        bool technicalExpanded               = false;
        float detailScrollOffset             = 0.0f;
        D2D1_RECT_F detailViewport           = D2D1::RectF();
        D2D1_RECT_F firstRowBounds           = D2D1::RectF();
        D2D1_RECT_F primarySelectedRowBounds = D2D1::RectF();
        D2D1_RECT_F copyDetailBounds         = D2D1::RectF();
        D2D1_RECT_F technicalBounds          = D2D1::RectF();
        std::array<std::wstring, 9> detailTexts{};
        std::array<D2D1_RECT_F, 9> detailTextBounds{};
        std::array<float, 9> detailRequiredHeights{};
        std::array<wchar_t, 9> detailIconGlyphs{};
        std::array<D2D1_RECT_F, 9> detailIconBounds{};
        std::array<D2D1_COLOR_F, 9> detailIconColors{};
        bool detailIconsUseFluentFont = false;
        std::wstring detailCopyText;
        std::wstring hiddenSortText;
        std::wstring copyCaption;
        std::wstring technicalCaption;
        std::wstring emptyStateText;
        std::wstring detailAccessibleName;
        std::vector<std::wstring> columnTitles;
        std::vector<std::wstring> columnIds;
        std::vector<std::wstring> selectedCells;
        std::vector<D2D1_RECT_F> columnBounds;
        size_t rowCount                 = 0;
        size_t selectionCount           = 0;
        uint64_t primarySelectedRowId   = 0;
        uint64_t primarySelectedTaskId  = 0;
        uint64_t refreshGeneration      = 0;
        uint64_t projectionBuildCount   = 0u;
        uint64_t projectionSortCount    = 0u;
        uint64_t issuesSourceRevision   = 0u;
        bool refreshTimerRunning        = false;
        uint32_t refreshTimerIntervalMs = 0u;
        bool primaryHasAcceptedDecision = false;
        bool primaryHasOriginalProblem  = false;
        bool primaryHasConflictResult   = false;
        bool primaryResultIsAmbiguous   = false;
        std::wstring primaryResultText;
        bool primaryDecisionIsCached   = false;
        uint32_t primaryDecisionAction = 0u;
        DxUi::GridVisibleWorkMetrics visibleWork{};
        bool themeDark                    = false;
        bool themeHighContrast            = false;
        bool themeRainbow                 = false;
        uint64_t dxRenderCount            = 0u;
        uint64_t dxResizeCount            = 0u;
        uint64_t dxResizeFailureCount     = 0u;
        bool detailFocused                = false;
        bool technicalFocused             = false;
        bool copyDetailFocused            = false;
        bool gridFocused                  = false;
        uint64_t firstVisibleTaskId       = 0u;
        D2D1_RECT_F taskHeaderRect        = D2D1::RectF();
        D2D1_RECT_F operationHeaderRect   = D2D1::RectF();
        D2D1_RECT_F messageHeaderRect     = D2D1::RectF();
        bool hasActiveSort                = false;
        size_t sortColumnIndex            = (std::numeric_limits<size_t>::max)();
        bool sortDescending               = false;
        uint32_t selectedIssueRowFillArgb = 0u;
        uint32_t selectedIssueRowTextArgb = 0u;
        bool selectedIssueRowUsesRainbow  = false;
        // Issues pane UX (issues#1, #4, #5, x-perf#6): the row the detail shows, the projected row
        // order, the Copy action's state and the grid/presentation and formatting work counters.
        uint64_t detailRowId = 0u;
        std::vector<uint64_t> rowStableIds;
        std::vector<uint64_t> rowTaskIds;
        std::wstring primaryProblemText;
        bool copyDetailEnabled                 = false;
        bool gridSingleSelection               = false;
        bool gridNarrow                        = false;
        uint64_t presentationApplyCount        = 0u;
        uint64_t presentationRebuildCount      = 0u;
        uint64_t gridModelAssignmentGeneration = 0u;
        uint64_t systemMessageLookupCount      = 0u;
    };

    struct SelfTestGridHit
    {
        uint32_t zone         = 0u;
        size_t rowIndex       = 0u;
        size_t groupIndex     = 0u;
        size_t columnIndex    = 0u;
        D2D1_RECT_F rectDip   = D2D1::RectF();
        bool onScrollbarThumb = false;
        bool isHeaderResize   = false;
    };

    // Debug-only test helpers: these observe and drive the live pane without owning it.
    static bool TryGetSelfTestSnapshot(HWND hwnd, SelfTestSnapshot& outSnapshot, bool includeDetail = false) noexcept;
    static bool SelfTestHitTestGridPoint(HWND hwnd, float xDip, float yDip, SelfTestGridHit& outHit) noexcept;
    static bool SelfTestSelectTask(HWND hwnd, uint64_t taskId, uint64_t promptGeneration = 0u) noexcept;
    static bool SelfTestSelectDiagnostic(HWND hwnd, uint64_t taskId, std::wstring_view category, HRESULT status) noexcept;
    static bool SelfTestFocusGrid(HWND hwnd) noexcept;
    // Selects a grid row as a click with these modifiers would (MK_SHIFT, MK_CONTROL), without focus.
    static bool SelfTestRequestSelectRow(HWND hwnd, size_t rowIndex, UINT modifiers) noexcept;
    static bool SelfTestDetailKey(HWND hwnd, UINT key) noexcept;
    static uint64_t SelfTestLiveStateCount() noexcept;
    static void SelfTestFailNextCreation(bool fail) noexcept;
    static bool SelfTestSetTechnicalExpanded(HWND hwnd, bool expanded) noexcept;
    static bool SelfTestRefresh(HWND hwnd, bool force) noexcept;
    static bool SelfTestScrollByWheelDetents(HWND hwnd, int detents) noexcept;
#endif

private:
    FileOperationsIssuesPane() = delete;
};
