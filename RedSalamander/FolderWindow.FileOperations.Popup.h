#pragma once

#include "FolderWindow.h"

#include <array>
#include <filesystem>
#include <optional>
#include <string>
#include <unordered_map>
#include <vector>

namespace FileOperationsPopupInternal
{
struct PopupHitTest
{
    enum class Kind : uint8_t
    {
        None,
        FooterCancelAll,
        FooterQueueMode,
        FooterAutoDismissSuccess,
        TaskToggleCollapse,
        TaskPause,
        TaskCancel,
        TaskSkip,
        TaskDestination,
        TaskSpeedLimit,
        TaskShowLog,
        TaskExportIssues,
        TaskConflictToggleApplyToAll,
        TaskConflictAction,
        TaskDismiss,
    };

    Kind kind       = Kind::None;
    uint64_t taskId = 0;
    uint32_t data   = 0;
};

#ifdef _DEBUG
struct PopupSelfTestInvoke
{
    PopupHitTest::Kind kind = PopupHitTest::Kind::None;
    uint64_t taskId         = 0;
    uint32_t data           = 0;
};
#endif

struct PopupButton
{
    D2D1_RECT_F bounds{};
    PopupHitTest hit{};
};

struct TaskSnapshot
{
    static constexpr size_t kMaxInFlightFiles   = 8u;
    static constexpr size_t kMaxConflictActions = 8u;

    struct InFlightFileSnapshot
    {
        std::wstring sourcePath;
        unsigned __int64 totalBytes     = 0;
        unsigned __int64 completedBytes = 0;
        ULONGLONG lastUpdateTick        = 0;
    };

    uint64_t taskId               = 0;
    FileSystemOperation operation = FILESYSTEM_COPY;

    unsigned long totalItems            = 0;
    unsigned long completedItems        = 0;
    unsigned __int64 totalBytes         = 0;
    unsigned __int64 completedBytes     = 0;
    unsigned __int64 itemTotalBytes     = 0;
    unsigned __int64 itemCompletedBytes = 0;

    std::wstring currentSourcePath;
    std::wstring currentDestinationPath;

    std::array<InFlightFileSnapshot, kMaxInFlightFiles> inFlightFiles{};
    size_t inFlightFileCount = 0;

    struct ConflictPromptSnapshot
    {
        bool active    = false;
        uint8_t bucket = 0;
        HRESULT status = S_OK;
        std::wstring sourcePath;
        std::wstring destinationPath;
        std::array<uint8_t, kMaxConflictActions> actions{};
        size_t actionCount     = 0;
        bool applyToAllChecked = false;
        bool retryFailed       = false;
    };

    ConflictPromptSnapshot conflict{};

    unsigned __int64 desiredSpeedLimitBytesPerSecond   = 0;
    unsigned __int64 effectiveSpeedLimitBytesPerSecond = 0;

    bool finished              = false;
    HRESULT resultHr           = S_OK;
    unsigned long warningCount = 0;
    unsigned long errorCount   = 0;
    std::wstring lastDiagnosticMessage;

    bool started                 = false;
    bool paused                  = false;
    bool hasProgressCallbacks    = false;
    ULONGLONG operationStartTick = 0;

    bool waitingForOthers = false;
    bool waitingInQueue   = false;
    bool queuePaused      = false;

    // Pre-calculation state
    bool preCalcInProgress              = false;
    bool preCalcSkipped                 = false;
    bool preCalcCompleted               = false;
    unsigned __int64 preCalcTotalBytes  = 0;
    unsigned long preCalcFileCount      = 0;
    unsigned long preCalcDirectoryCount = 0;
    ULONGLONG preCalcElapsedMs          = 0;

    unsigned long plannedItems = 0;
    std::filesystem::path destinationFolder;
    std::optional<FolderWindow::Pane> destinationPane;
};

struct RateSnapshot
{
    uint64_t taskId               = 0;
    FileSystemOperation operation = FILESYSTEM_COPY;

    unsigned long completedItems    = 0;
    unsigned __int64 completedBytes = 0;
    std::wstring currentSourcePath;
    bool started          = false;
    bool paused           = false;
    bool waitingForOthers = false;
    bool waitingInQueue   = false;
    bool queuePaused      = false;
};

struct RateHistory
{
    static constexpr size_t kMaxSamples = 180u; // ~18s @ 100ms

    std::array<float, kMaxSamples> samples{};
    std::array<float, kMaxSamples> hues{}; // Per-sample hue (0-360) for rainbow mode
    size_t count      = 0;
    size_t writeIndex = 0;

    ULONGLONG lastTick         = 0;
    unsigned __int64 lastBytes = 0;
    unsigned long lastItems    = 0;

    float smoothedBytesPerSec = 0.0f;
    float smoothedItemsPerSec = 0.0f;
};

class FileOperationsPopupState final
{
public:
    FolderWindow::FileOperationState* fileOps = nullptr;
    FolderWindow* folderWindow                = nullptr;

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

private:
    enum class CaptionStatus : uint8_t
    {
        None,
        Ok,
        Warning,
        Error,
    };

    void ApplyScrollBarTheme(HWND hwnd) const noexcept;

    bool IsTaskCollapsed(uint64_t taskId) const noexcept;
    void ToggleTaskCollapsed(uint64_t taskId) noexcept;
    void CleanupCollapsedTasks(const std::vector<TaskSnapshot>& snapshot) noexcept;

    void DiscardDeviceResources() noexcept;
    void EnsureFactories() noexcept;
    void EnsureTextFormats() noexcept;
    void EnsureTarget(HWND hwnd) noexcept;
    void EnsureBrushes() noexcept;

    std::vector<TaskSnapshot> BuildSnapshot() const;
    std::vector<RateSnapshot> BuildRateSnapshot() const;
    void UpdateRates() noexcept;

    void LayoutChrome(float width, float height) noexcept;
    void UpdateScrollBar(HWND hwnd, float viewH, float contentH) noexcept;
    void AutoResizeWindow(HWND hwnd, float desiredContentHeight, size_t taskCount) noexcept;

    void DrawButton(const PopupButton& button, IDWriteTextFormat* format, std::wstring_view text) noexcept;
    void DrawMenuButton(const PopupButton& button, IDWriteTextFormat* format, std::wstring_view text) noexcept;
    void DrawCheckboxBox(const D2D1_RECT_F& rect, bool checked) noexcept;
    void DrawCollapseChevron(const D2D1_RECT_F& rc, bool collapsed) noexcept;
    void DrawBandwidthGraph(const D2D1_RECT_F& rect,
                            const RateHistory& history,
                            unsigned __int64 limitBytesPerSecond,
                            std::wstring_view overlayText,
                            bool showAnimation,
                            bool rainbowMode,
                            ULONGLONG tick) noexcept;
    void Render(HWND hwnd) noexcept;
    void UpdateLastPopupRect(HWND hwnd) noexcept;
    void UpdateCaptionStatus(HWND hwnd, const std::vector<TaskSnapshot>& snapshot) noexcept;
    void PaintCaptionStatusGlyph(HWND hwnd) const noexcept;

    PopupHitTest HitTest(float x, float y) const noexcept;
    void Invalidate(HWND hwnd) const noexcept;

    bool ConfirmCancelAll(HWND hwnd) noexcept;
    void ShowSpeedLimitMenu(HWND hwnd, uint64_t taskId) noexcept;
    void ShowDestinationMenu(HWND hwnd, uint64_t taskId) noexcept;

    LRESULT OnCreate(HWND hwnd) noexcept;
    LRESULT OnThemeChanged(HWND hwnd) noexcept;
    LRESULT OnNcDestroy(HWND hwnd) noexcept;
    LRESULT OnSize(HWND hwnd, UINT width, UINT height) noexcept;
    LRESULT OnDpiChanged(HWND hwnd, UINT newDpi, const RECT& suggested) noexcept;
    LRESULT OnGetMinMaxInfo(HWND hwnd, MINMAXINFO* info) noexcept;
    LRESULT OnMove(HWND hwnd) noexcept;
    LRESULT OnTimer(HWND hwnd, UINT_PTR timerId) noexcept;
    LRESULT OnEnterSizeMove(HWND hwnd) noexcept;
    LRESULT OnExitSizeMove(HWND hwnd) noexcept;
    LRESULT OnVScroll(HWND hwnd, UINT request) noexcept;
    LRESULT OnMouseMove(HWND hwnd, POINT pt) noexcept;
    LRESULT OnMouseLeave(HWND hwnd) noexcept;
    LRESULT OnLButtonDown(HWND hwnd, POINT pt) noexcept;
    LRESULT OnLButtonUp(HWND hwnd, POINT pt) noexcept;
    LRESULT OnMouseWheel(HWND hwnd, int delta) noexcept;
    LRESULT OnClose(HWND hwnd) noexcept;
    LRESULT OnNcPaint(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept;
    LRESULT OnNcActivate(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept;
#ifdef _DEBUG
    LRESULT OnSelfTestInvoke(HWND hwnd, const PopupSelfTestInvoke* payload) noexcept;
#endif

    UINT _dpi = USER_DEFAULT_SCREEN_DPI;
    SIZE _clientSize{};

    bool _trackingMouse = false;
    bool _inSizeMove    = false;
    bool _inThemeChange = false;

    CaptionStatus _captionStatus = CaptionStatus::None;

    float _scrollY                    = 0.0f;
    float _contentHeight              = 0.0f;
    float _lastAutoSizedContentHeight = 0.0f; // For auto-resize tracking
    size_t _lastTaskCount             = 0;    // For auto-resize tracking
    int _maxAutoSizedWindowHeight     = 0;    // Sticky max window height (prevents resize "dancing")
    int _scrollPos                    = 0;
    bool _scrollBarVisible            = false;

    D2D1_RECT_F _footerCancelAllRect{};
    D2D1_RECT_F _footerQueueModeRect{};
    D2D1_RECT_F _listViewportRect{};

    std::vector<PopupButton> _buttons;
    PopupHitTest _hotHit{};
    PopupHitTest _pressedHit{};

    std::unordered_map<uint64_t, RateHistory> _rates;
    std::unordered_map<uint64_t, bool> _collapsedTasks;

    wil::com_ptr<ID2D1Factory> _d2dFactory;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> _target;

    wil::com_ptr<IDWriteTextFormat> _headerFormat;
    wil::com_ptr<IDWriteTextFormat> _bodyFormat;
    wil::com_ptr<IDWriteTextFormat> _smallFormat;
    wil::com_ptr<IDWriteTextFormat> _buttonFormat;
    wil::com_ptr<IDWriteTextFormat> _buttonSmallFormat;
    wil::com_ptr<IDWriteTextFormat> _graphOverlayFormat;
    wil::com_ptr<IDWriteTextFormat> _statusIconFormat;
    wil::com_ptr<IDWriteTextFormat> _statusIconFallbackFormat;

    wil::com_ptr<ID2D1SolidColorBrush> _bgBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _textBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _subTextBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _borderBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _progressBgBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _progressGlobalBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _progressItemBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _checkboxFillBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _checkboxCheckBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _statusOkBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _statusWarningBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _statusErrorBrush;
    D2D1::ColorF _progressItemBaseColor = D2D1::ColorF(D2D1::ColorF::Black);
    wil::com_ptr<ID2D1SolidColorBrush> _graphBgBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _graphGridBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _graphLimitBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _graphLineBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _graphFillBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _graphDynamicBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _graphTextShadowBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _buttonBgBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _buttonHoverBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _buttonPressedBrush;
    D2D1::ColorF _graphFillBaseColor = D2D1::ColorF(D2D1::ColorF::Black);

    int _mouseWheelRemainder = 0;
};
} // namespace FileOperationsPopupInternal

class FileOperationsPopup final
{
public:
    static HWND Create(FolderWindow::FileOperationState* fileOps, FolderWindow* folderWindow, HWND ownerWindow) noexcept;

private:
    FileOperationsPopup() = delete;
};
