#pragma once

#include "FolderWindow.h"

#include <cstdint>
#include <limits>
#include <memory>

#include "DxUi/DxUi.h"

class FileOperationsIssuesPane final
{
public:
    static HWND Create(FolderWindow::FileOperationState* fileOps, FolderWindow* folderWindow, HWND ownerWindow, std::weak_ptr<void> hostLifetime) noexcept;

#ifdef ENABLE_TESTS
    struct SelfTestSnapshot
    {
        size_t rowCount                = 0;
        size_t selectionCount          = 0;
        uint64_t primarySelectedRowId  = 0;
        uint64_t primarySelectedTaskId = 0;
        uint64_t refreshGeneration     = 0;
        RedSalamander::DxUi::GridVisibleWorkMetrics visibleWork{};
        bool themeDark                    = false;
        bool themeHighContrast            = false;
        bool themeRainbow                 = false;
        uint64_t dxRenderCount            = 0u;
        uint64_t dxResizeCount            = 0u;
        uint64_t dxResizeFailureCount     = 0u;
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
    static bool TryGetSelfTestSnapshot(HWND hwnd, SelfTestSnapshot& outSnapshot) noexcept;
    static bool SelfTestHitTestGridPoint(HWND hwnd, float xDip, float yDip, SelfTestGridHit& outHit) noexcept;
    static bool SelfTestSelectTask(HWND hwnd, uint64_t taskId) noexcept;
    static bool SelfTestFocusGrid(HWND hwnd) noexcept;
    static bool SelfTestRefresh(HWND hwnd, bool force) noexcept;
    static bool SelfTestScrollByWheelDetents(HWND hwnd, int detents) noexcept;
#endif

private:
    FileOperationsIssuesPane() = delete;
};
