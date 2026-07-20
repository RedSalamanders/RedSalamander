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

#include <uxtheme.h>
#include <windowsx.h>

#include "CommandRegistry.h"
#include "CompareDirectoriesEngine.h"
#include "DxUi/DxUi.h"
#include "FileSystemPluginManager.h"
#include "FluentIcons.h"
#include "FolderView.h"
#include "FolderWindow.h"
#include "Helpers.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/Informations.h"
#include "SessionState.h"
#include "ShortcutManager.h"
#include "UiMetrics.h"
#include "WindowMessages.h"
#include "WindowPlacementPersistence.h"
#include "WindowSizing.h"
#include "resource.h"

namespace CompareDirectoriesWindowInternal
{
void LogComparePerfStats(std::wstring_view reason, const std::shared_ptr<CompareDirectoriesSession>& session, HRESULT resultHr) noexcept;

constexpr wchar_t kCompareDirectoriesWindowClassName[] = L"RedSalamander.CompareDirectoriesWindow";
constexpr wchar_t kCompareDirectoriesWindowId[]        = L"CompareDirectoriesWindow";

constexpr UINT_PTR kScanProgressTextId                  = 1003;
constexpr UINT_PTR kScanProgressBarId                   = 1004;
constexpr UINT_PTR kCompareTaskAutoDismissTimerId       = 1005;
constexpr UINT kCompareTaskAutoDismissDelayMs           = 5000;
constexpr UINT_PTR kCompareProgressPulseTimerId         = 1006;
constexpr UINT kCompareProgressPulseTimerIntervalMs     = 16;
constexpr UINT_PTR kCompareDecisionRefreshTimerId       = 1007;
constexpr UINT kCompareDecisionRefreshTimerIntervalMs   = 200;
constexpr UINT_PTR kCompareTaskCardTrailingFlushTimerId = 1008;
constexpr int kScanStatusHeightDip                      = 22;
constexpr int kScanStatusPaddingXDip                    = 6;
constexpr int kSplitterGripDotSizeDip                   = 2;
constexpr int kSplitterGripDotGapDip                    = 2;
constexpr int kSplitterGripDotCount                     = 3;
constexpr float kMinSplitRatio                          = 0.0f;
constexpr float kMaxSplitRatio                          = 1.0f;

LRESULT CALLBACK CompareOptionsHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
LRESULT CALLBACK CompareOptionsWheelRouteWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
LRESULT CALLBACK CompareOptionsDxHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
LRESULT CALLBACK CompareProgressSpinnerWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
LRESULT CALLBACK CompareDxChromeHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

class CompareDirectoriesWindow final
{
public:
    CompareDirectoriesWindow(FolderWindow& applicationFolderWindow,
                             Common::Settings::Settings& settings,
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

    friend LRESULT CALLBACK CompareOptionsHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    friend LRESULT CALLBACK CompareOptionsWheelRouteWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    friend LRESULT CALLBACK CompareOptionsDxHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    friend LRESULT CALLBACK CompareProgressSpinnerWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    friend LRESULT CALLBACK CompareDxChromeHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    bool OnCreate(HWND hwnd) noexcept;
    void OnDestroy() noexcept;
    void OnNcDestroy() noexcept;
    [[nodiscard]] HWND ResolveRestoreFolderViewWindow() const noexcept;
    void UpdateOwnerWindow(HWND owner) noexcept;
    void RestoreOwnerFocusAfterClose() noexcept;
    void OnSize() noexcept;
    void OnDpiChanged(UINT newDpi, const RECT* newRect) noexcept;
    void OnCommand(UINT id) noexcept;
    LRESULT OnFunctionBarInvoke(WPARAM wParam, LPARAM lParam) noexcept;
    void OnPaint() noexcept;
    LRESULT OnCtlColorStatic(HDC hdc, HWND control) noexcept;
    void UpdateViewMenuChecks() noexcept;
    void ShowSortMenuPopup(FolderWindow::Pane pane, POINT screenPoint) noexcept;
    [[nodiscard]] bool EnsureDxChromeHosts() noexcept;
    void DetachDxChromeHosts() noexcept;
    void ApplyDxChromeTheme() noexcept;
    void SyncDxMenuBar() noexcept;
    void SyncDxBannerButtons() noexcept;
    void SyncDxBannerText() noexcept;
    void SyncDxChrome() noexcept;
    [[nodiscard]] int GetDxMenuBarVisibleHeightPx() const noexcept;
    [[nodiscard]] bool FocusFirstDxMenuBarItem() noexcept;
    [[nodiscard]] bool ActivateDxMenuBarMnemonic(wchar_t mnemonic) noexcept;
    [[nodiscard]] std::optional<size_t> HitTestDxMenuBarScreenPoint(POINT screenPoint) const noexcept;
    [[nodiscard]] std::optional<POINT> GetDxMenuBarItemAnchorScreenPoint(size_t index) const noexcept;
    [[nodiscard]] std::optional<size_t> FindNextEnabledDxMenuBarItem(size_t currentIndex, bool forward) const noexcept;
    [[nodiscard]] std::optional<RedSalamander::DxUi::ContextMenuRootSwitchRequest> BuildDxMenuBarRootSwitchRequest(size_t index) noexcept;
    void CaptureDxMenuBarFocusRestoreTarget() noexcept;
    void RestoreDxMenuBarFocus() noexcept;
    void OpenDxMenuBarPopup(size_t index, POINT screenPoint, bool keyboardInvocation) noexcept;
    [[nodiscard]] LRESULT HandleDxChromeHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept;

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
    [[nodiscard]] INT_PTR OnOptionsSettingsReloadedFromDisk(HWND dlg) noexcept;
    [[nodiscard]] INT_PTR OnOptionsEraseBkgnd(HWND dlg, HDC hdc) noexcept;
    [[nodiscard]] INT_PTR OnOptionsCommand(HWND dlg, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] INT_PTR OnOptionsCtlColorEdit(HDC hdc, HWND control) noexcept;
    [[nodiscard]] INT_PTR OnOptionsCtlColorDlg(HDC hdc) noexcept;
    [[nodiscard]] INT_PTR OnOptionsCtlColorStatic(HDC hdc, HWND control) noexcept;
    [[nodiscard]] INT_PTR OnOptionsCtlColorBtn(HDC hdc, HWND control) noexcept;
    [[nodiscard]] bool HandleOptionsDxMnemonic(wchar_t mnemonic) noexcept;
    void RegisterOptionsReloadParticipant() noexcept;
    void UnregisterOptionsReloadParticipant() noexcept;
    void CreateChildWindows(HWND hwnd) noexcept;
    void EnsureOptionsControlsCreated(HWND dlg) noexcept;
    [[nodiscard]] bool EnsureOptionsDxStaticHosts() noexcept;
    void DetachOptionsDxStaticHosts() noexcept;
    [[nodiscard]] bool EnsureOptionsDxButtonHosts() noexcept;
    void DetachOptionsDxButtonHosts() noexcept;
    [[nodiscard]] bool EnsureOptionsDxToggleHosts() noexcept;
    void DetachOptionsDxToggleHosts() noexcept;
    [[nodiscard]] bool EnsureOptionsDxEditHosts() noexcept;
    void DetachOptionsDxEditHosts() noexcept;
    void ApplyOptionsDxStaticTheme() noexcept;
    void ApplyOptionsDxButtonTheme() noexcept;
    void ApplyOptionsDxToggleTheme() noexcept;
    void ApplyOptionsDxEditTheme() noexcept;
    void SyncOptionsDxStatics() noexcept;
    void SyncOptionsDxButtons() noexcept;
    void SyncOptionsDxToggles() noexcept;
    void SyncOptionsDxEdits() noexcept;
    [[nodiscard]] bool EnsureOptionsDxBodyControlVisible(RedSalamander::DxUi::Control* control) noexcept;
    [[nodiscard]] LRESULT HandleOptionsDxHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept;
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

    struct LeaveScopePromptRequest
    {
        ComparePane changedPane = ComparePane::Left;
        uint64_t runId          = 0;
        std::filesystem::path attemptedPath;
        std::filesystem::path revertPath;
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
    void SetComparePanePathSilently(ComparePane pane, const std::filesystem::path& path) noexcept;
    void QueueLeaveScopePrompt(ComparePane changedPane, const std::filesystem::path& attemptedPath, const std::filesystem::path& revertPath) noexcept;
    LRESULT OnLeaveScopePrompt() noexcept;
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
    uint64_t _fileOperationCompletedCallbackToken = 0;
    LRESULT OnScanProgress(WPARAM operationKey, LPARAM lp) noexcept;
    LRESULT OnContentProgress(WPARAM operationKey, LPARAM lp) noexcept;
    void UpdateProgressControls() noexcept;
    void EnsureProgressPulseTimer() noexcept;
    void StopProgressPulseTimerIfIdle() noexcept;
    void OnProgressPulseTimer() noexcept;
    void InvalidateCompareWatermarkPanesIfDue(ULONGLONG now) noexcept;
    void DrawProgressSpinner(HDC hdc, const RECT& bounds) noexcept;
    void UpdateCompareWatermark() noexcept;
    void UpdateRescanButtonText() noexcept;
    void UpdateCompareTaskCard(bool finished, bool force = false) noexcept;
    void ScheduleCompareTaskCardTrailingFlush(ULONGLONG delayMs) noexcept;
    void CancelCompareTaskCardTrailingFlush() noexcept;
    void OnCompareTaskCardTrailingFlushTimer() noexcept;
    void MaybeCompleteCompareRun() noexcept;
    void DismissCompareTaskCard() noexcept;
    LRESULT OnExecuteShortcutCommand(LPARAM lp) noexcept;
    void ExecuteShortcutCommand(std::wstring_view commandId) noexcept;

    Common::Settings::CompareDirectoriesSettings GetEffectiveCompareSettings() const noexcept;
    Common::Settings::CompareDirectoriesSettings ReadOptionsControlsToSettings() const noexcept;
    [[nodiscard]] bool IsOptionsDialogDirty() const noexcept;
    void ApplyOptionsDraftInterlocks() noexcept;
    void UpdateOptionsDraftToggle(UINT controlId, bool checked) noexcept;
    void UpdateOptionsDraftText(UINT controlId, std::wstring_view text) noexcept;
    void LoadOptionsControlsFromSettings() noexcept;
    void SaveOptionsControlsToSettings() noexcept;
    [[nodiscard]] bool ResolveOptionsStaleSaveConflict(HWND dlg) noexcept;
    void ReloadOptionsDialogFromDisk() noexcept;
    void UpdateOptionsVisibility() noexcept;
    void RefreshBothPanes() noexcept;
    void ScheduleDecisionRefresh() noexcept;
    void OnDecisionRefreshTimer() noexcept;
    [[nodiscard]] bool StartCompareSyncToOtherPane(FolderWindow::Pane sourcePane, FileSystemOperation operation) noexcept;

    wil::unique_hwnd _hWnd;

    struct ChromeController
    {
        ChromeController()                                       = default;
        ~ChromeController()                                      = default;
        ChromeController(const ChromeController&)                = delete;
        ChromeController& operator=(const ChromeController&)     = delete;
        ChromeController(ChromeController&&) noexcept            = delete;
        ChromeController& operator=(ChromeController&&) noexcept = delete;

        wil::unique_hwnd bannerTitle;
        wil::unique_hwnd bannerOptionsButton;
        wil::unique_hwnd bannerRescanButton;
        wil::unique_hmenu menuHandle;
        wil::unique_hwnd menuBarHostHwnd;
        wil::unique_hwnd bannerOptionsHostHwnd;
        wil::unique_hwnd bannerRescanHostHwnd;
        wil::unique_hwnd bannerTitleHostHwnd;
        RedSalamander::DxUi::WindowHost menuBarHost;
        RedSalamander::DxUi::WindowHost bannerOptionsHost;
        RedSalamander::DxUi::WindowHost bannerRescanHost;
        RedSalamander::DxUi::WindowHost bannerTitleHost;
        RedSalamander::DxUi::MenuBar* menuBar              = nullptr;
        RedSalamander::DxUi::Button* bannerOptionsButtonDx = nullptr;
        RedSalamander::DxUi::Button* bannerRescanButtonDx  = nullptr;
        RedSalamander::DxUi::Label* bannerTitleLabel       = nullptr;
        std::atomic<int> menuBarSelectedIndexSnapshot{-1};
        HWND menuBarFocusRestoreHwnd = nullptr;
        bool usesMenuBar             = false;
        bool usesBannerButtons       = false;
        bool usesBannerText          = false;
        bool rescanIsCancel          = false;
    };

    ChromeController _chrome{};

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

    enum class CompareWatermarkState
    {
        Hidden,
        InProgress,
        Cancelled,
    };

    struct ProgressController
    {
        ProgressController()                                         = default;
        ~ProgressController()                                        = default;
        ProgressController(const ProgressController&)                = delete;
        ProgressController& operator=(const ProgressController&)     = delete;
        ProgressController(ProgressController&&) noexcept            = delete;
        ProgressController& operator=(ProgressController&&) noexcept = delete;

        wil::unique_hwnd scanProgressText;
        wil::unique_hwnd scanProgressBar;
        wil::unique_hwnd scanProgressTextHostHwnd;
        RedSalamander::DxUi::WindowHost scanProgressTextHost;
        RedSalamander::DxUi::Label* scanProgressTextLabel = nullptr;
        BannerProgressState banner{};
        uint64_t scanStartTickMs                    = 0;
        float spinnerAngleDeg                       = 0.0f;
        ULONGLONG spinnerLastTickMs                 = 0;
        bool pulseTimerActive                       = false;
        ULONGLONG paneWatermarkLastInvalidateTickMs = 0;
        uint64_t contentEtaLastTickMs               = 0;
        uint64_t contentEtaLastCompletedBytes       = 0;
        double contentEtaSmoothedBytesPerSec        = 0.0;
        std::optional<uint64_t> contentEtaSeconds;
        bool controlsVisible = false;
        std::wstring lastMessage;
        ULONGLONG lastTaskCardUpdateTickMs   = 0;
        bool taskCardTrailingFlushPending    = false;
        bool compareRunSawScanProgress       = false;
        uint64_t compareTaskId               = 0;
        HRESULT compareRunResultHr           = S_OK;
        CompareWatermarkState watermarkState = CompareWatermarkState::Hidden;
    };

    ProgressController _progress{};

    static constexpr int kProgressSpinnerSegments = 12;

    bool _decisionRefreshPending         = false;
    bool _decisionRefreshTimerActive     = false;
    bool _decisionRefreshFallbackPending = false;

    std::optional<std::filesystem::path> _lastRefreshedLeftRelativeFolder;
    std::optional<std::filesystem::path> _lastRefreshedRightRelativeFolder;
    std::shared_ptr<const CompareDirectoriesFolderDecision> _lastRefreshedLeftDecision;
    std::shared_ptr<const CompareDirectoriesFolderDecision> _lastRefreshedRightDecision;

    struct OptionsUi
    {
        HWND host = nullptr;
    };

    struct OptionsToggleCardDx
    {
        RedSalamander::DxUi::CardPanel* card    = nullptr;
        RedSalamander::DxUi::Label* title       = nullptr;
        RedSalamander::DxUi::Label* description = nullptr;
        RedSalamander::DxUi::Toggle* toggle     = nullptr;
    };

    struct OptionsIgnoreCardDx
    {
        RedSalamander::DxUi::CardPanel* card    = nullptr;
        RedSalamander::DxUi::Label* title       = nullptr;
        RedSalamander::DxUi::Label* description = nullptr;
        RedSalamander::DxUi::Toggle* toggle     = nullptr;
        RedSalamander::DxUi::TextField* edit    = nullptr;
    };

    struct OptionsBodyDx
    {
        OptionsBodyDx() = default;
        ~OptionsBodyDx() noexcept
        {
            Detach();
        }
        OptionsBodyDx(const OptionsBodyDx&)                = delete;
        OptionsBodyDx& operator=(const OptionsBodyDx&)     = delete;
        OptionsBodyDx(OptionsBodyDx&&) noexcept            = delete;
        OptionsBodyDx& operator=(OptionsBodyDx&&) noexcept = delete;

        wil::unique_hwnd hostHwnd;
        RedSalamander::DxUi::WindowHost host;
        RedSalamander::DxUi::Label* headerCompare  = nullptr;
        RedSalamander::DxUi::Label* headerSubdirs  = nullptr;
        RedSalamander::DxUi::Label* headerAdvanced = nullptr;
        RedSalamander::DxUi::Label* headerIgnore   = nullptr;
        OptionsToggleCardDx compareSize;
        OptionsToggleCardDx compareDateTime;
        OptionsToggleCardDx compareAttributes;
        OptionsToggleCardDx compareContent;
        OptionsToggleCardDx compareSubdirectories;
        OptionsToggleCardDx compareSubdirAttributes;
        OptionsToggleCardDx selectSubdirsOnlyInOnePane;
        OptionsToggleCardDx keepIdenticalItems;
        OptionsIgnoreCardDx ignoreFiles;
        OptionsIgnoreCardDx ignoreDirectories;
        RedSalamander::DxUi::Control* lastFooterReturnTarget = nullptr;

        void Detach() noexcept
        {
            host.Detach();
            if (hostHwnd)
            {
                if (IsWindow(hostHwnd.get()) != FALSE)
                {
                    hostHwnd.reset();
                }
                else
                {
                    static_cast<void>(hostHwnd.release());
                }
            }
            headerCompare              = nullptr;
            headerSubdirs              = nullptr;
            headerAdvanced             = nullptr;
            headerIgnore               = nullptr;
            compareSize                = {};
            compareDateTime            = {};
            compareAttributes          = {};
            compareContent             = {};
            compareSubdirectories      = {};
            compareSubdirAttributes    = {};
            selectSubdirsOnlyInOnePane = {};
            keepIdenticalItems         = {};
            ignoreFiles                = {};
            ignoreDirectories          = {};
            lastFooterReturnTarget     = nullptr;
        }
    };

    struct OptionsButtonDx
    {
        OptionsButtonDx() = default;
        ~OptionsButtonDx() noexcept
        {
            Detach();
        }
        OptionsButtonDx(const OptionsButtonDx&)                = delete;
        OptionsButtonDx& operator=(const OptionsButtonDx&)     = delete;
        OptionsButtonDx(OptionsButtonDx&&) noexcept            = delete;
        OptionsButtonDx& operator=(OptionsButtonDx&&) noexcept = delete;

        wil::unique_hwnd hostHwnd;
        RedSalamander::DxUi::WindowHost host;
        RedSalamander::DxUi::Button* button = nullptr;
        int attachFailureStage              = 0;

        void Detach() noexcept
        {
            host.Detach();
            if (hostHwnd)
            {
                if (IsWindow(hostHwnd.get()) != FALSE)
                {
                    hostHwnd.reset();
                }
                else
                {
                    static_cast<void>(hostHwnd.release());
                }
            }
            button             = nullptr;
            attachFailureStage = 0;
        }
    };

    struct OptionsDxUiState
    {
        OptionsDxUiState()                                       = default;
        ~OptionsDxUiState()                                      = default;
        OptionsDxUiState(const OptionsDxUiState&)                = delete;
        OptionsDxUiState& operator=(const OptionsDxUiState&)     = delete;
        OptionsDxUiState(OptionsDxUiState&&) noexcept            = delete;
        OptionsDxUiState& operator=(OptionsDxUiState&&) noexcept = delete;

        OptionsBodyDx body;
        OptionsButtonDx okButton;
        OptionsButtonDx cancelButton;
        bool usesDxUiStatics = false;
        bool usesDxUiButtons = false;
        bool usesDxUiToggles = false;
        bool usesDxUiEdits   = false;

        void Detach() noexcept
        {
            body.Detach();
            okButton.Detach();
            cancelButton.Detach();
            usesDxUiStatics = false;
            usesDxUiButtons = false;
            usesDxUiToggles = false;
            usesDxUiEdits   = false;
        }
    };

    struct OptionsPanelController
    {
        OptionsPanelController()                                             = default;
        ~OptionsPanelController()                                            = default;
        OptionsPanelController(const OptionsPanelController&)                = delete;
        OptionsPanelController& operator=(const OptionsPanelController&)     = delete;
        OptionsPanelController(OptionsPanelController&&) noexcept            = delete;
        OptionsPanelController& operator=(OptionsPanelController&&) noexcept = delete;

        wil::unique_hwnd dlg;
        OptionsUi ui{};
        std::unique_ptr<OptionsDxUiState> dxUi;
        std::vector<RECT> cards;
        int scrollOffset                 = 0;
        int scrollMax                    = 0;
        int bodyContentHeight            = 0;
        int wheelRemainder               = 0;
        bool useTwoColumns               = false;
        bool usesDxUiTypographyMetrics   = false;
        int twoColumnSeparatorX          = -1;
        bool staleFromExternalReload     = false;
        bool reloadParticipantRegistered = false;
        bool syncingDxDraft              = false;
        Common::Settings::CompareDirectoriesSettings draft{};
        bool draftLoaded = false;
        wil::unique_hbrush backgroundBrush;
        wil::unique_hbrush cardBrush;
        wil::unique_hbrush inputBrush;
        wil::unique_hbrush inputFocusedBrush;
        wil::unique_hbrush inputDisabledBrush;
        COLORREF inputBackgroundColor         = RGB(255, 255, 255);
        COLORREF inputFocusedBackgroundColor  = RGB(255, 255, 255);
        COLORREF inputDisabledBackgroundColor = RGB(255, 255, 255);
    };

    OptionsPanelController _optionsPanel{};

    Common::Settings::Settings* _settings  = nullptr;
    FolderWindow* _applicationFolderWindow = nullptr;
    AppTheme _theme{};
    const ShortcutManager* _shortcuts = nullptr;
    HWND _ownerWindow                 = nullptr;
    HWND _restoreFolderViewWindow     = nullptr;
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

    wil::unique_hbrush _backgroundBrush;
    wil::unique_hbrush _splitterBrush;
    wil::unique_hbrush _splitterGripBrush;
    wil::unique_hbrush _menuBackgroundBrush;

    bool _compareStarted    = false;
    bool _compareActive     = false;
    bool _compareRunPending = false;
    bool _syncingPaths      = false;
    uint64_t _compareRunId  = 0;

    enum class CompareSelectionMode : uint8_t
    {
        Default,
        Inverted,
    };

    CompareSelectionMode _selectionMode = CompareSelectionMode::Default;

    enum class DeferredStartPhase
    {
        None,
        Scheduled,
        Prepared,
    };

    DeferredStartPhase _deferredCompareStartPhase = DeferredStartPhase::None;
    uint64_t _deferredCompareStartRunId           = 0;
    bool _deferredCompareStartStartedBefore       = false;
    std::optional<LeaveScopePromptRequest> _leaveScopePromptRequest;
    bool _leaveScopePromptPending = false;

    std::optional<std::filesystem::path> _lastLeftPluginPath;
    std::optional<std::filesystem::path> _lastRightPluginPath;
    UINT _dpi               = USER_DEFAULT_SCREEN_DPI;
    int _restoreShowCmd     = SW_SHOWNORMAL;
    bool _hasSavedPlacement = false;
    size_t _dispatchDepth   = 0u;
    bool _createInProgress  = false;
    bool _deletePending     = false;

#ifdef ENABLE_TESTS
public:
    [[nodiscard]] bool DebugGetOptionsSnapshot(::CompareDirectoriesOptionsDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugGetRunSnapshot(::CompareDirectoriesRunDebugSnapshot& out) const noexcept;
    [[nodiscard]] bool DebugSetPanePath(bool leftPane, const std::filesystem::path& path) noexcept;
    [[nodiscard]] bool DebugSetRunPending(bool pending) noexcept;
    [[nodiscard]] bool DebugGetMenuBarItemLabel(size_t index, std::wstring& outText) const noexcept;
    [[nodiscard]] bool DebugGetMenuBarItemScreenRect(size_t index, RECT& outRect) const noexcept;
    [[nodiscard]] bool DebugFocusOptionsFirstControl() noexcept;
    [[nodiscard]] bool DebugFocusOptionsTarget(::CompareDirectoriesOptionsDebugFocusTarget target) noexcept;
    [[nodiscard]] bool DebugSetOptionsIgnoreFilesEnabled(bool enabled) noexcept;
    [[nodiscard]] bool DebugSetOptionsIgnoreDirectoriesEnabled(bool enabled) noexcept;
    [[nodiscard]] HWND DebugGetOptionsDialogHandle() const noexcept;
    [[nodiscard]] bool DebugScrollOptionsBodyPages(int pageDelta) noexcept;
    [[nodiscard]] bool DebugGetOptionsTargetHostAndClientRect(::CompareDirectoriesOptionsDebugFocusTarget target, HWND& outHost, RECT& outRect) const noexcept;
#endif
};

} // namespace CompareDirectoriesWindowInternal
