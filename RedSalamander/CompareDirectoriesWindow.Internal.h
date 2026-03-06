#pragma once

#include "CompareDirectoriesWindow.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cmath>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include <commctrl.h>
#include <uxtheme.h>
#include <windowsx.h>

#include "CommandRegistry.h"
#include "CompareDirectoriesEngine.h"
#include "FileSystemPluginManager.h"
#include "FluentIcons.h"
#include "FolderView.h"
#include "FolderWindow.h"
#include "Helpers.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "SessionState.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/Informations.h"
#include "ShortcutManager.h"
#include "ThemedControls.h"
#include "ThemedInputFrames.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "resource.h"

namespace CompareDirectoriesWindowInternal
{
[[nodiscard]] std::wstring_view LoadStringResourceView(_In_opt_ HINSTANCE hInstance, _In_ UINT uID) noexcept;

void LogComparePerfStats(std::wstring_view reason, const std::shared_ptr<CompareDirectoriesSession>& session, HRESULT resultHr) noexcept;

constexpr wchar_t kCompareDirectoriesWindowClassName[] = L"RedSalamander.CompareDirectoriesWindow";
constexpr wchar_t kCompareDirectoriesWindowId[]        = L"CompareDirectoriesWindow";

constexpr UINT_PTR kScanProgressTextId                = 1003;
constexpr UINT_PTR kScanProgressBarId                 = 1004;
constexpr UINT_PTR kCompareTaskAutoDismissTimerId     = 1005;
constexpr UINT kCompareTaskAutoDismissDelayMs         = 5000;
constexpr UINT_PTR kCompareBannerSpinnerTimerId       = 1006;
constexpr UINT kCompareBannerSpinnerTimerIntervalMs   = 16;
constexpr UINT_PTR kCompareDecisionRefreshTimerId     = 1007;
constexpr UINT kCompareDecisionRefreshTimerIntervalMs = 200;
constexpr UINT_PTR kCompareProgressSpinnerSubclassId  = 3u;

constexpr int kScanStatusHeightDip    = 22;
constexpr int kScanStatusPaddingXDip  = 6;
constexpr int kSplitterGripDotSizeDip = 2;
constexpr int kSplitterGripDotGapDip  = 2;
constexpr int kSplitterGripDotCount   = 3;
constexpr float kMinSplitRatio        = 0.0f;
constexpr float kMaxSplitRatio        = 1.0f;

struct CompareMenuItemData
{
    bool separator  = false;
    bool topLevel   = false;
    bool hasSubMenu = false;
    std::wstring text;
    std::wstring shortcut;
};

LRESULT CALLBACK CompareOptionsHostSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept;
LRESULT CALLBACK CompareOptionsWheelRouteSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept;
LRESULT CALLBACK CompareProgressSpinnerSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept;

class CompareDirectoriesWindow final
{
public:
    CompareDirectoriesWindow(Common::Settings::Settings& settings,
                             AppTheme theme,
                             const ShortcutManager* shortcuts,
                             CompareDirectoriesPaneContext left,
                             CompareDirectoriesPaneContext right) noexcept;

    [[nodiscard]] bool Create(HWND owner) noexcept;
    void UpdateTheme(const AppTheme& theme) noexcept;

    CompareDirectoriesWindow(const CompareDirectoriesWindow&)            = delete;
    CompareDirectoriesWindow& operator=(const CompareDirectoriesWindow&) = delete;
    CompareDirectoriesWindow(CompareDirectoriesWindow&&)                 = delete;
    CompareDirectoriesWindow& operator=(CompareDirectoriesWindow&&)      = delete;

private:
    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    static INT_PTR CALLBACK OptionsDlgProc(HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam) noexcept;

    friend LRESULT CALLBACK CompareOptionsHostSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept;
    friend LRESULT CALLBACK CompareOptionsWheelRouteSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept;
    friend LRESULT CALLBACK CompareProgressSpinnerSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR subclassId, DWORD_PTR refData) noexcept;

    bool OnCreate(HWND hwnd) noexcept;
    void OnDestroy() noexcept;
    void OnNcDestroy() noexcept;
    void OnSize() noexcept;
    void OnDpiChanged(UINT newDpi, const RECT* newRect) noexcept;
    void OnCommand(UINT id) noexcept;
    LRESULT OnFunctionBarInvoke(WPARAM wParam, LPARAM lParam) noexcept;
    void OnPaint() noexcept;
    LRESULT OnCtlColorStatic(HDC hdc, HWND control) noexcept;
    void PrepareThemedMenu() noexcept;
    void PrepareThemedMenuRecursive(HMENU menu, bool topLevel, std::vector<std::unique_ptr<CompareMenuItemData>>& itemData) noexcept;
    void UpdateViewMenuChecks() noexcept;
    void OnMeasureItem(MEASUREITEMSTRUCT* mis) noexcept;
    void OnDrawItem(DRAWITEMSTRUCT* dis) noexcept;
    void ShowSortMenuPopup(FolderWindow::Pane pane, POINT screenPoint) noexcept;

    void OnLButtonDown(POINT pt) noexcept;
    void OnLButtonDblClk(POINT pt) noexcept;
    void OnLButtonUp() noexcept;
    void OnMouseMove(POINT pt) noexcept;
    void OnCaptureChanged() noexcept;
    [[nodiscard]] bool OnSetCursor(POINT pt) noexcept;
    void SetSplitRatio(float ratio) noexcept;

    void ApplyTheme() noexcept;
    void ApplyOptionsDialogTheme() noexcept;
    [[nodiscard]] INT_PTR OnOptionsInitDialog(HWND dlg) noexcept;
    [[nodiscard]] INT_PTR OnOptionsEraseBkgnd(HWND dlg, HDC hdc) noexcept;
    [[nodiscard]] INT_PTR OnOptionsCommand(HWND dlg, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] INT_PTR OnOptionsDrawItem(const DRAWITEMSTRUCT* dis) noexcept;
    [[nodiscard]] INT_PTR OnOptionsCtlColorEdit(HDC hdc, HWND control) noexcept;
    [[nodiscard]] INT_PTR OnOptionsCtlColorDlg(HDC hdc) noexcept;
    [[nodiscard]] INT_PTR OnOptionsCtlColorStatic(HDC hdc, HWND control) noexcept;
    [[nodiscard]] INT_PTR OnOptionsCtlColorBtn(HDC hdc, HWND control) noexcept;
    void CreateChildWindows(HWND hwnd) noexcept;
    void EnsureOptionsControlsCreated(HWND dlg) noexcept;
    void LayoutOptionsControls() noexcept;
    void PaintOptionsHostBackgroundAndCards(HDC hdc, HWND host) noexcept;
    void Layout() noexcept;

    void EnsureCompareSession() noexcept;
    void StartCompare() noexcept;
    void BeginOrRescanCompare() noexcept;
    void ScheduleBeginOrRescanCompare() noexcept;
    void CancelCompareMode() noexcept;

    struct PreparedCompareRun
    {
        uint64_t runId     = 0;
        bool startedBefore = false;
    };

    [[nodiscard]] std::optional<PreparedCompareRun> PrepareCompareRun() noexcept;
    void ExecutePreparedCompareRun(uint64_t runId, bool startedBefore) noexcept;
    LRESULT OnDeferredBeginOrRescanCompare(WPARAM wp) noexcept;
    void SetSessionCallbacksForRun(uint64_t runId) noexcept;
    void UpdateCompareRootsFromCurrentPanes() noexcept;
    void ShowOptionsPanel(bool show) noexcept;

    void OnPanePathChanged(ComparePane pane, const std::optional<std::filesystem::path>& newPath) noexcept;
    void SyncOtherPanePath(ComparePane changedPane,
                           const std::optional<std::filesystem::path>& previousPath,
                           const std::optional<std::filesystem::path>& newPath) noexcept;
    void ApplySelectionForFolder(ComparePane pane, const std::filesystem::path& folder) noexcept;
    void UpdateEmptyStateForFolder(ComparePane pane, const std::filesystem::path& folder) noexcept;
    [[nodiscard]] std::wstring BuildDetailsTextForCompareItem(ComparePane pane,
                                                              const std::filesystem::path& folder,
                                                              std::wstring_view displayName,
                                                              bool isDirectory,
                                                              uint64_t sizeBytes,
                                                              int64_t lastWriteTime,
                                                              DWORD fileAttributes) noexcept;
    [[nodiscard]] std::wstring BuildMetadataTextForCompareItem(ComparePane pane,
                                                               const std::filesystem::path& folder,
                                                               std::wstring_view displayName,
                                                               bool isDirectory,
                                                               uint64_t sizeBytes,
                                                               int64_t lastWriteTime,
                                                               DWORD fileAttributes) noexcept;
    void OnFolderWindowFileOperationCompleted(const FolderWindow::FileOperationCompletedEvent& e) noexcept;
    LRESULT OnScanProgress(LPARAM lp) noexcept;
    LRESULT OnContentProgress(LPARAM lp) noexcept;
    void UpdateProgressControls() noexcept;
    void OnProgressSpinnerTimer() noexcept;
    void InvalidateSpinnerPens() noexcept;
    void EnsureSpinnerPens(COLORREF background, COLORREF accent, bool rainbowSpinner, uint32_t rainbowSeedHash, int stroke) noexcept;
    void DrawProgressSpinner(HDC hdc, const RECT& bounds) noexcept;
    void UpdateCompareWatermark() noexcept;
    void UpdateRescanButtonText() noexcept;
    void UpdateCompareTaskCard(bool finished) noexcept;
    void MaybeCompleteCompareRun() noexcept;
    void DismissCompareTaskCard() noexcept;
    LRESULT OnExecuteShortcutCommand(LPARAM lp) noexcept;
    void ExecuteShortcutCommand(std::wstring_view commandId) noexcept;

    Common::Settings::CompareDirectoriesSettings GetEffectiveCompareSettings() const noexcept;
    void LoadOptionsControlsFromSettings() noexcept;
    void SaveOptionsControlsToSettings() noexcept;
    void UpdateOptionsVisibility() noexcept;
    void RefreshBothPanes() noexcept;
    void ScheduleDecisionRefresh() noexcept;
    void OnDecisionRefreshTimer() noexcept;

    wil::unique_hwnd _hWnd;
    wil::unique_hwnd _optionsDlg;
    wil::unique_hwnd _scanProgressText;
    wil::unique_hwnd _scanProgressBar;
    wil::unique_hwnd _bannerTitle;
    wil::unique_hwnd _bannerOptionsButton;
    wil::unique_hwnd _bannerRescanButton;

    struct BannerProgressState
    {
        static constexpr size_t kMaxContentInFlightSlots = 8u;

        struct ContentInFlightEntry
        {
            std::filesystem::path relativePath;
            uint64_t totalBytes      = 0;
            uint64_t completedBytes  = 0;
            ULONGLONG lastUpdateTick = 0;
        };

        uint32_t scanActiveScans                = 0;
        uint64_t scanFolderCount                = 0;
        uint64_t scanEntryCount                 = 0;
        uint64_t scanContentCandidateFileCount  = 0;
        uint64_t scanContentCandidateTotalBytes = 0;
        std::filesystem::path scanRelativeFolder;
        std::wstring scanEntryName;

        uint64_t contentPendingCompares       = 0;
        uint64_t contentTotalCompares         = 0;
        uint64_t contentCompletedCompares     = 0;
        uint64_t contentOverallTotalBytes     = 0;
        uint64_t contentOverallCompletedBytes = 0;
        uint64_t contentFileTotalBytes        = 0;
        uint64_t contentFileCompletedBytes    = 0;
        std::filesystem::path contentRelativeFolder;
        std::wstring contentEntryName;

        std::array<ContentInFlightEntry, kMaxContentInFlightSlots> contentInFlight{};
    };

    BannerProgressState _progress{};

    uint64_t _scanStartTickMs = 0;

    float _progressSpinnerAngleDeg               = 0.0f;
    ULONGLONG _progressSpinnerLastTickMs         = 0;
    bool _progressSpinnerTimerActive             = false;
    ULONGLONG _paneWatermarkLastInvalidateTickMs = 0;

    static constexpr int kProgressSpinnerSegments = 12;

    struct ProgressSpinnerPenKey
    {
        COLORREF background = 0;
        COLORREF accent     = 0;
        bool rainbow        = false;
        bool darkBase       = false;
        uint32_t seedHash   = 0;
        int strokeWidthPx   = 0;

        bool operator==(const ProgressSpinnerPenKey&) const noexcept = default;
    };

    ProgressSpinnerPenKey _progressSpinnerPenKey{};
    bool _progressSpinnerPenKeyValid = false;
    std::array<wil::unique_hpen, kProgressSpinnerSegments> _progressSpinnerPens{};

    bool _decisionRefreshPending     = false;
    bool _decisionRefreshTimerActive = false;

    std::optional<std::filesystem::path> _lastRefreshedLeftRelativeFolder;
    std::optional<std::filesystem::path> _lastRefreshedRightRelativeFolder;
    std::shared_ptr<const CompareDirectoriesFolderDecision> _lastRefreshedLeftDecision;
    std::shared_ptr<const CompareDirectoriesFolderDecision> _lastRefreshedRightDecision;

    uint64_t _contentEtaLastTickMs         = 0;
    uint64_t _contentEtaLastCompletedBytes = 0;
    double _contentEtaSmoothedBytesPerSec  = 0.0;
    std::optional<uint64_t> _contentEtaSeconds;

    struct OptionsToggleCard
    {
        HWND title       = nullptr;
        HWND description = nullptr;
        HWND toggle      = nullptr;
    };

    struct OptionsIgnoreCard
    {
        HWND title       = nullptr;
        HWND description = nullptr;
        HWND toggle      = nullptr;
        HWND frame       = nullptr;
        HWND edit        = nullptr;
    };

    struct OptionsUi
    {
        HWND host = nullptr;

        HWND headerCompare  = nullptr;
        HWND headerSubdirs  = nullptr;
        HWND headerAdvanced = nullptr;
        HWND headerIgnore   = nullptr;

        OptionsToggleCard compareSize;
        OptionsToggleCard compareDateTime;
        OptionsToggleCard compareAttributes;
        OptionsToggleCard compareContent;
        OptionsToggleCard compareSubdirectories;

        OptionsToggleCard compareSubdirAttributes;
        OptionsToggleCard selectSubdirsOnlyInOnePane;
        OptionsToggleCard keepIdenticalItems;

        OptionsIgnoreCard ignoreFiles;
        OptionsIgnoreCard ignoreDirectories;
    };

    OptionsUi _optionsUi{};
    std::vector<RECT> _optionsCards;
    int _optionsScrollOffset        = 0;
    int _optionsScrollMax           = 0;
    int _optionsWheelRemainder      = 0;
    bool _optionsUseTwoColumns      = false;
    int _optionsTwoColumnSeparatorX = -1;

    Common::Settings::Settings* _settings = nullptr;
    AppTheme _theme{};
    const ShortcutManager* _shortcuts = nullptr;
    CompareDirectoriesPaneContext _leftContext;
    CompareDirectoriesPaneContext _rightContext;

    wil::unique_hmodule _leftBaseModule;
    wil::unique_hmodule _rightBaseModule;
    wil::com_ptr<IFileSystem> _leftBaseFs;
    wil::com_ptr<IFileSystem> _rightBaseFs;
    std::wstring _leftPluginShortId;
    std::wstring _rightPluginShortId;

    std::shared_ptr<CompareDirectoriesSession> _session;
    wil::com_ptr<IFileSystem> _fsLeft;
    wil::com_ptr<IFileSystem> _fsRight;

    FolderWindow _folderWindow;

    struct DetailsDecisionCache
    {
        std::filesystem::path folder;
        uint64_t sessionUiVersion = 0;
        std::shared_ptr<const CompareDirectoriesFolderDecision> decision;
    };

    DetailsDecisionCache _detailsCacheLeft;
    DetailsDecisionCache _detailsCacheRight;

    FolderView::DisplayMode _compareDisplayMode = FolderView::DisplayMode::Detailed;

    // Layout
    SIZE _clientSize{};
    RECT _splitterRect{};
    float _splitRatio         = 0.5f;
    bool _draggingSplitter    = false;
    int _splitterDragOffsetPx = 0;

    wil::unique_hfont _uiFont;
    wil::unique_hfont _uiBoldFont;
    wil::unique_hfont _uiItalicFont;
    wil::unique_hfont _bannerTitleFont;
    wil::unique_hbrush _backgroundBrush;
    wil::unique_hbrush _splitterBrush;
    wil::unique_hbrush _splitterGripBrush;
    wil::unique_hbrush _menuBackgroundBrush;
    wil::unique_hbrush _optionsBackgroundBrush;
    wil::unique_hbrush _optionsCardBrush;
    wil::unique_hbrush _optionsInputBrush;
    wil::unique_hbrush _optionsInputFocusedBrush;
    wil::unique_hbrush _optionsInputDisabledBrush;

    COLORREF _optionsInputBackgroundColor         = RGB(255, 255, 255);
    COLORREF _optionsInputFocusedBackgroundColor  = RGB(255, 255, 255);
    COLORREF _optionsInputDisabledBackgroundColor = RGB(255, 255, 255);
    ThemedInputFrames::FrameStyle _optionsFrameStyle{};

    std::vector<std::unique_ptr<CompareMenuItemData>> _menuItemData;
    std::vector<std::unique_ptr<CompareMenuItemData>> _popupMenuItemData;

    bool _compareStarted            = false;
    bool _compareActive             = false;
    bool _compareRunPending         = false;
    bool _compareRunSawScanProgress = false;
    bool _bannerRescanIsCancel      = false;
    bool _syncingPaths              = false;
    uint64_t _compareRunId          = 0;
    uint64_t _compareTaskId         = 0;
    HRESULT _compareRunResultHr     = S_OK;

    enum class CompareWatermarkState
    {
        Hidden,
        InProgress,
        Cancelled,
    };

    CompareWatermarkState _watermarkState = CompareWatermarkState::Hidden;

    enum class DeferredStartPhase
    {
        None,
        Scheduled,
        Prepared,
    };

    DeferredStartPhase _deferredCompareStartPhase = DeferredStartPhase::None;
    uint64_t _deferredCompareStartRunId           = 0;
    bool _deferredCompareStartStartedBefore       = false;

    std::optional<std::filesystem::path> _lastLeftPluginPath;
    std::optional<std::filesystem::path> _lastRightPluginPath;
    UINT _dpi               = USER_DEFAULT_SCREEN_DPI;
    int _restoreShowCmd     = SW_SHOWNORMAL;
    bool _hasSavedPlacement = false;
};

} // namespace CompareDirectoriesWindowInternal
