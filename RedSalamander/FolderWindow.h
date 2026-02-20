#pragma once
#include <array>
#include <condition_variable>
#include <cstdint>
#include <filesystem>
#include <functional>
#include <memory>
#include <mutex>
#include <optional>
#include <stop_token>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#include "AppTheme.h"
#include "FolderView.h"
#include "Framework.h"
#include "FunctionBar.h"
#include "NavigationView.h"
#include "PlugInterfaces/Viewer.h"

namespace Common::Settings
{
struct Settings;
}

class ShortcutManager;

class FolderWindow
{
public:
    FolderWindow();
    ~FolderWindow();

    // Disable copy and move
    FolderWindow(const FolderWindow&)            = delete;
    FolderWindow& operator=(const FolderWindow&) = delete;
    FolderWindow(FolderWindow&&)                 = delete;
    FolderWindow& operator=(FolderWindow&&)      = delete;

    // Window management
    HWND Create(HWND parent, int x, int y, int width, int height);
    void Destroy();
    [[maybe_unused]] HWND GetHwnd() const noexcept
    {
        return _hWnd.get();
    }

    enum class Pane : uint8_t
    {
        Left,
        Right,
    };

    // Navigation
    void SetFolderPath(const std::filesystem::path& path);
    std::optional<std::filesystem::path> GetCurrentPath() const;
    std::optional<std::filesystem::path> GetCurrentPluginPath() const;
    std::vector<std::filesystem::path> GetFolderHistory() const;
    void SetFolderHistory(const std::vector<std::filesystem::path>& history);
    uint32_t GetFolderHistoryMax() const noexcept;
    void SetFolderHistoryMax(uint32_t maxItems);

    void SetActivePane(Pane pane) noexcept;
    [[maybe_unused]] Pane GetActivePane() const noexcept
    {
        return _activePane;
    }
    [[maybe_unused]] Pane GetFocusedPane() const noexcept;
    [[nodiscard]] HWND GetFocusedFolderViewHwnd() const noexcept;
    [[nodiscard]] HWND GetFolderViewHwnd(Pane pane) const noexcept;

    HRESULT ExecuteInActivePane(const std::filesystem::path& folderPath,
                                std::wstring_view focusItemDisplayName,
                                unsigned int folderViewCommandId,
                                bool activateWindow) noexcept;

    void SetFolderPath(Pane pane, const std::filesystem::path& path);
    std::optional<std::filesystem::path> GetCurrentPath(Pane pane) const;
    std::optional<std::filesystem::path> GetCurrentPluginPath(Pane pane) const;
    std::vector<std::filesystem::path> GetFolderHistory(Pane pane) const;
    void SetFolderHistory(Pane pane, const std::vector<std::filesystem::path>& history);

    void SetDisplayMode(Pane pane, FolderView::DisplayMode mode);
    FolderView::DisplayMode GetDisplayMode(Pane pane) const noexcept;

    void SetSort(Pane pane, FolderView::SortBy sortBy, FolderView::SortDirection direction);
    void CycleSortBy(Pane pane, FolderView::SortBy sortBy);
    FolderView::SortBy GetSortBy(Pane pane) const noexcept;
    FolderView::SortDirection GetSortDirection(Pane pane) const noexcept;

    void SetStatusBarVisible(Pane pane, bool visible);
    bool GetStatusBarVisible(Pane pane) const noexcept;

    void CommandRename(Pane pane);
    void CommandView(Pane pane);
    void CommandViewSpace(Pane pane);
    void CommandDelete(Pane pane);
    void CommandPermanentDelete(Pane pane);
    void CommandPermanentDeleteWithValidation(Pane pane);
    void CommandCopyToOtherPane(Pane sourcePane);
    void CommandMoveToOtherPane(Pane sourcePane);
    void CommandToggleFileOperationsIssuesPane();
    bool IsFileOperationsIssuesPaneVisible() noexcept;
    void CommandCreateDirectory(Pane pane);
    void CommandChangeDirectory(Pane pane);
    void CommandFocusAddressBar(Pane pane);
    void CommandOpenDriveMenu(Pane pane);
    void CommandShowFolderHistory(Pane pane);
    void CommandRefresh(Pane pane);
    void CommandCalculateDirectorySizes(Pane pane);
    void CommandChangeCase(Pane pane);
    void CommandOpenCommandShell(Pane pane);
    void PrepareForNetworkDriveDisconnect(Pane pane);
    void SwapPanes();

    bool ConfirmCancelAllFileOperations(HWND ownerWindow) noexcept;
    void CloseAllViewers() noexcept;

    using ShowSortMenuCallback = std::function<void(Pane pane, POINT screenPoint)>;
    void SetShowSortMenuCallback(ShowSortMenuCallback callback);

    float GetSplitRatio() const noexcept
    {
        return _splitRatio;
    }
    void SetSplitRatio(float ratio);
    void BeginViewWidthAdjust() noexcept;
    void CommitViewWidthAdjust() noexcept;
    void CancelViewWidthAdjust() noexcept;
    [[nodiscard]] bool IsViewWidthAdjustActive() const noexcept
    {
        return _viewWidthAdjustActive;
    }
    [[nodiscard]] bool HandleViewWidthAdjustKey(uint32_t vk) noexcept;

#ifdef _DEBUG
    [[nodiscard]] bool DebugIsViewWidthAdjustActive() const noexcept
    {
        return _viewWidthAdjustActive;
    }
#endif
    void ToggleZoomPanel(Pane pane);
    [[nodiscard]] std::optional<Pane> GetZoomedPane() const noexcept
    {
        return _zoomedPane;
    }
    [[nodiscard]] std::optional<float> GetZoomRestoreSplitRatio() const noexcept
    {
        return _zoomRestoreSplitRatio;
    }
    void SetZoomState(std::optional<Pane> zoomedPane, std::optional<float> restoreSplitRatio);

    void SetSettings(Common::Settings::Settings* settings) noexcept;
    void SetShortcutManager(const ShortcutManager* shortcuts) noexcept;
    void SetFunctionBarModifiers(uint32_t modifiers) noexcept;
    void SetFunctionBarPressedKey(std::optional<uint32_t> vk) noexcept;
    void SetFunctionBarVisible(bool visible) noexcept;
    [[nodiscard]] bool GetFunctionBarVisible() const noexcept
    {
        return _functionBarVisible;
    }

    // Extension points (used by scoped folder windows like Compare).
    using PanePathChangedCallback = std::function<void(Pane pane, const std::optional<std::filesystem::path>& pluginPath)>;
    void SetPanePathChangedCallback(PanePathChangedCallback callback);
    void SetPaneEnumerationCompletedCallback(Pane pane, FolderView::EnumerationCompletedCallback callback);
    void SetPaneDetailsTextProvider(Pane pane, FolderView::DetailsTextProvider provider);
    void SetPaneMetadataTextProvider(Pane pane, FolderView::MetadataTextProvider provider);
    void SetPaneEmptyStateMessage(Pane pane, std::wstring message);
    void RefreshPaneDetailsText(Pane pane);
    void
    SetPaneSelectionByDisplayNamePredicate(Pane pane, const std::function<bool(std::wstring_view)>& shouldSelect, bool clearExistingSelection = true) noexcept;

    struct FileOperationCompletedEvent
    {
        FileSystemOperation operation = static_cast<FileSystemOperation>(0);
        Pane sourcePane               = Pane::Left;
        std::optional<Pane> destinationPane;
        std::vector<std::filesystem::path> sourcePaths;
        std::optional<std::filesystem::path> destinationFolder;
        HRESULT hr = S_OK;
    };
    using FileOperationCompletedCallback = std::function<void(const FileOperationCompletedEvent& e)>;
    void SetFileOperationCompletedCallback(FileOperationCompletedCallback callback);

    struct InformationalTaskUpdate final
    {
        static constexpr size_t kMaxContentInFlightFiles = 8u;

         enum class Kind : uint8_t
         {
             CompareDirectories,
            ChangeCase,
         };

         Kind kind       = Kind::CompareDirectories;
         uint64_t taskId = 0;
         std::wstring title;

        // Compare Directories payload (Kind::CompareDirectories)
        std::filesystem::path leftRoot;
        std::filesystem::path rightRoot;

        bool scanActive = false;
        std::filesystem::path scanCurrentRelative;
        uint64_t scanFolderCount         = 0;
        uint64_t scanEntryCount          = 0;
        uint64_t scanCandidateFileCount  = 0;
        uint64_t scanCandidateTotalBytes = 0;
        std::optional<uint64_t> scanElapsedSeconds;

        bool contentActive = false;
        std::filesystem::path contentCurrentRelative;
        uint64_t contentCurrentTotalBytes     = 0;
        uint64_t contentCurrentCompletedBytes = 0;
        uint64_t contentTotalBytes            = 0;
        uint64_t contentCompletedBytes        = 0;
        uint64_t contentPendingCount          = 0;
        uint64_t contentCompletedCount        = 0;
        std::optional<uint64_t> contentEtaSeconds;

        struct ContentInFlightFile final
        {
            std::filesystem::path relativePath;
            uint64_t totalBytes      = 0;
            uint64_t completedBytes  = 0;
            ULONGLONG lastUpdateTick = 0;
        };

         std::array<ContentInFlightFile, kMaxContentInFlightFiles> contentInFlight{};
         size_t contentInFlightCount = 0;

        // Change Case payload (Kind::ChangeCase)
        bool changeCaseEnumerating = false;
        bool changeCaseRenaming    = false;
        std::filesystem::path changeCaseCurrentPath;
        uint64_t changeCaseScannedFolders   = 0;
        uint64_t changeCaseScannedEntries   = 0;
        uint64_t changeCasePlannedRenames   = 0;
        uint64_t changeCaseCompletedRenames = 0;

         bool finished    = false;
         HRESULT resultHr = S_OK;
         std::wstring doneSummary;
     };

    // Informational tasks are read-only task cards displayed in the File Operations popup for background work
    // that isn't a file operation (e.g., Compare Directories scan/content progress).
    [[nodiscard]] uint64_t CreateOrUpdateInformationalTask(const InformationalTaskUpdate& update) noexcept;
    void DismissInformationalTask(uint64_t taskId) noexcept;

    // DPI handling
    void OnDpiChanged(float newDpi);

    void ApplyTheme(const AppTheme& theme);
    [[maybe_unused]] const AppTheme& GetTheme() const noexcept
    {
        return _theme;
    }

    HRESULT ReloadFileSystemPlugins() noexcept;
    HRESULT SetFileSystemPluginForPane(Pane pane, std::wstring_view pluginId) noexcept;
    HRESULT SetFileSystemInstanceForPane(
        Pane pane, wil::com_ptr<IFileSystem> fileSystem, std::wstring pluginId, std::wstring pluginShortId, std::wstring instanceContext) noexcept;
    [[maybe_unused]] std::wstring_view GetFileSystemPluginId(Pane pane) const noexcept;

    void DebugShowOverlaySample(Pane pane, FolderView::OverlaySeverity severity);
    void DebugShowOverlaySampleNonModal(Pane pane, FolderView::OverlaySeverity severity);
    void DebugShowOverlaySampleBusyWithCancel(Pane pane);
    void DebugShowOverlaySampleCanceled(Pane pane);
    void DebugHideOverlaySample(Pane pane);

    struct FileOperationState;

#ifdef _DEBUG
    // Debug/testing hook: access the file-operations state for automation/self-tests.
    // This will initialize file operations if they are not yet created.
    FileOperationState* DebugGetFileOperationState() noexcept;

    [[nodiscard]] size_t DebugGetViewerInstanceCount() const noexcept;
    [[nodiscard]] bool DebugHasViewerPluginId(std::wstring_view viewerPluginId) const noexcept;
    [[nodiscard]] uint64_t DebugGetForceRefreshCount(Pane pane) const noexcept;
#endif

    void ShowPaneAlertOverlay(Pane pane,
                              FolderView::ErrorOverlayKind kind,
                              FolderView::OverlaySeverity severity,
                              std::wstring title,
                              std::wstring message,
                              HRESULT hr       = S_OK,
                              bool closable    = true,
                              bool blocksInput = true);

    void DismissPaneAlertOverlay(Pane pane);

    struct FileOperationStateDeleter
    {
        void operator()(FileOperationState* state) const noexcept;
    };

private:
    // Class registration
    static ATOM RegisterWndClass(HINSTANCE instance);
    static constexpr PCWSTR kClassName = L"RedSalamander.FolderWindow";

    // Window procedure
    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);

    // Message handlers
    bool OnCreate(HWND hwnd) noexcept;
    void OnDestroy();
    void OnSize(UINT width, UINT height);
    void OnSetFocus();
    void OnPaint();
    void OnLButtonDown(POINT pt);
    void OnLButtonDblClk(POINT pt);
    void OnLButtonUp();
    void OnMouseMove(POINT pt);
    void OnCaptureChanged();
    LRESULT OnDrawItem(DRAWITEMSTRUCT* dis);
    LRESULT OnNotify(const NMHDR* header);
    LRESULT OnSetCursor(HWND cursorWindow, UINT hitTest, UINT mouseMsg);
    bool OnSetCursor(POINT pt);
    void OnParentNotify(UINT eventMsg, UINT childId);
    LRESULT OnDeviceChange(UINT event, LPARAM data) noexcept;
    void OnNetworkConnectivityChanged() noexcept;
    LRESULT OnPaneSelectionSizeComputed(LPARAM lp) noexcept;
    LRESULT OnPaneSelectionSizeProgress(LPARAM lp) noexcept;
    LRESULT OnFileOperationCompleted(LPARAM lp) noexcept;
    LRESULT OnChangeCaseTaskUpdate(LPARAM lp) noexcept;
    LRESULT OnChangeCaseCompleted(LPARAM lp) noexcept;

    // File operations (internal implementation in FolderWindow.FileOperations.cpp)
    void EnsureFileOperations();
    HRESULT StartFileOperationFromFolderView(Pane pane, FolderView::FileOperationRequest request) noexcept;
    HRESULT ShowItemPropertiesFromFolderView(Pane pane, std::filesystem::path path) noexcept;
    void ShutdownFileOperations() noexcept;
    void ApplyFileOperationsTheme() noexcept;

    // Layout
    void CalculateLayout();
    void AdjustChildWindows();
    void UpdatePaneStatusBar(Pane pane);
    void UpdatePaneFocusStates() noexcept;
    void StartSelectionSizeWorker(Pane pane) noexcept;
    void CancelSelectionSizeComputation(Pane pane) noexcept;
    void RequestSelectionSizeComputation(Pane pane);
    void SelectionSizeWorkerMain(Pane pane, std::stop_token stopToken) noexcept;

    // Path synchronization
    void OnNavigationPathChanged(Pane pane, const std::optional<std::filesystem::path>& path);
    void OnFolderViewPathChanged(Pane pane, const std::optional<std::filesystem::path>& path);
    void OnFolderViewNavigateUpFromRoot(Pane pane) noexcept;
    HRESULT EnsurePaneFileSystem(Pane pane, std::wstring_view pluginId) noexcept;
    Pane GetPaneFromChild(HWND child) const noexcept;
    bool TryOpenFileAsVirtualFileSystem(Pane pane, const std::filesystem::path& path) noexcept;
    bool TryViewFileWithViewer(Pane pane, const FolderView::ViewFileRequest& request) noexcept;
    bool TryViewSpaceWithViewer(Pane pane, const std::filesystem::path& folderPath) noexcept;

    struct ViewerInstance final
    {
        std::wstring viewerPluginId;
        wil::com_ptr<IViewer> viewer;
        ViewerOpenContext openContext{};
        wil::com_ptr<IFileSystem> fileSystem;
        std::wstring fileSystemName;
        std::wstring focusedPath;
        std::vector<std::wstring> selectionStorage;
        std::vector<const wchar_t*> selectionPointers;
        std::vector<std::wstring> otherFilesStorage;
        std::vector<const wchar_t*> otherFilePointers;
    };

    struct ViewerCallbackState final : public IViewerCallback
    {
        FolderWindow* owner = nullptr;
        HRESULT STDMETHODCALLTYPE ViewerClosed(void* cookie) noexcept override;
    };

    void ShutdownViewers() noexcept;
    void ApplyViewerTheme() noexcept;
    ViewerTheme BuildViewerTheme() const noexcept;
    HRESULT OnViewerClosed(ViewerInstance* instance) noexcept;

private:
    class NetworkChangeSubscription;

    wil::unique_hwnd _hWnd;
    HINSTANCE _hInstance = nullptr;
    UINT _dpi            = USER_DEFAULT_SCREEN_DPI;

    // Child components
    struct PaneState
    {
        PaneState() = default;
        // Explicitly delete copy/move operations
        PaneState(const PaneState&)            = delete;
        PaneState(PaneState&&)                 = delete;
        PaneState& operator=(const PaneState&) = delete;
        PaneState& operator=(PaneState&&)      = delete;

        NavigationView navigationView;
        FolderView folderView;
        wil::unique_hwnd hNavigationView;
        wil::unique_hwnd hFolderView;
        wil::unique_hwnd hStatusBar;
        bool statusBarVisible = true;
        FolderView::SelectionStats selectionStats{};
        uint64_t selectionSizeGeneration = 0;
        std::jthread selectionSizeThread;
        std::mutex selectionSizeMutex;
        std::condition_variable selectionSizeCv;
        bool selectionSizeWorkPending        = false;
        uint64_t selectionSizeWorkGeneration = 0;
        std::vector<std::filesystem::path> selectionSizeWorkFolders;
        wil::com_ptr<IFileSystem> selectionSizeWorkFileSystem;
        std::shared_ptr<std::stop_source> selectionSizeWorkStopSource;

        std::jthread changeCaseThread;
        bool selectionFolderBytesPending = false;
        bool selectionFolderBytesValid   = false;
        uint64_t selectionFolderBytes    = 0;
        std::wstring statusSelectionText;
        std::wstring statusSortText;
        uint32_t statusFocusHueDegrees = 0;
        bool sortIndicatorHot          = false;

        wil::unique_hmodule fileSystemModule;
        wil::com_ptr<IFileSystem> fileSystem;
        std::wstring pluginId;
        std::wstring pluginShortId;
        std::wstring instanceContext;

        std::optional<std::filesystem::path> currentPath;
        bool updatingPath = false;
    };
    bool SanityCheckBothPanes(PaneState& src, PaneState& dest, FileSystemOperation operation);

    PaneState _leftPane;
    PaneState _rightPane;
    Pane _activePane = Pane::Left;
    FunctionBar _functionBar;
    bool _functionBarVisible                = true;
    const ShortcutManager* _shortcutManager = nullptr;

    // Layout
    SIZE _clientSize{};
    RECT _leftPaneRect{};
    RECT _rightPaneRect{};
    RECT _splitterRect{};
    RECT _leftNavigationRect{};
    RECT _leftFolderViewRect{};
    RECT _leftStatusBarRect{};
    RECT _rightNavigationRect{};
    RECT _rightFolderViewRect{};
    RECT _rightStatusBarRect{};
    RECT _functionBarRect{};
    float _splitRatio = 0.5f;
    bool _viewWidthAdjustActive       = false;
    float _viewWidthAdjustRestoreRatio = 0.5f;
    std::optional<float> _zoomRestoreSplitRatio;
    std::optional<Pane> _zoomedPane;
    bool _draggingSplitter    = false;
    int _splitterDragOffsetPx = 0;
    wil::unique_hbrush _backgroundBrush;
    wil::unique_hbrush _splitterBrush;
    wil::unique_hbrush _splitterGripBrush;

    AppTheme _theme;
    uint32_t _statusBarRainbowHueDegrees = 0;
    ShowSortMenuCallback _showSortMenuCallback;
    PanePathChangedCallback _panePathChangedCallback;
    FileOperationCompletedCallback _fileOperationCompletedCallback;

    std::unique_ptr<FileOperationState, FileOperationStateDeleter> _fileOperations;
    Common::Settings::Settings* _settings = nullptr;
    uint32_t _folderHistoryMax            = 20u;
    std::vector<std::filesystem::path> _folderHistory;

    ViewerCallbackState _viewerCallback;
    std::vector<std::unique_ptr<ViewerInstance>> _viewerInstances;

    std::unique_ptr<NetworkChangeSubscription> _networkChangeSubscription;
    uint64_t _lastNetworkConnectivityRefreshTick = 0;
};
