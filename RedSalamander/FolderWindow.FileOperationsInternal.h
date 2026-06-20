#pragma once

// Internal implementation header for FolderWindow file operations.
// Keep this header private to the FolderWindow file-operation translation units.

#include "FolderWindowInternal.h"

struct FolderWindow::FileOperationState
{
    enum class DiagnosticSeverity : unsigned char
    {
        Debug,
        Info,
        Warning,
        Error,
    };

    enum class ExecutionMode : unsigned char
    {
        BulkItems,
        PerItem,
    };

    struct TaskCompletedPayload
    {
        uint64_t taskId             = 0;
        HRESULT hr                  = S_OK;
        unsigned long warningCount  = 0;
        unsigned long errorCount    = 0;
    };

    struct TaskDiagnosticEntry
    {
        SYSTEMTIME localTime{};
        uint64_t taskId                 = 0;
        FileSystemOperation operation   = FILESYSTEM_COPY;
        DiagnosticSeverity severity     = DiagnosticSeverity::Info;
        HRESULT status                  = S_OK;
        uint64_t processWorkingSetBytes = 0;
        uint64_t processPrivateBytes    = 0;
        std::wstring category;
        std::wstring message;
        std::wstring sourcePath;
        std::wstring destinationPath;
        std::wstring concurrencyMode;
        std::wstring storageType;
        std::wstring destinationStorageType;
        unsigned long autoTunedConcurrency       = 0;
        unsigned long effectiveConcurrencyBudget = 0;
    };

    struct CompletedTaskSummary
    {
        uint64_t taskId               = 0;
        FileSystemOperation operation = FILESYSTEM_COPY;
        FolderWindow::Pane sourcePane = FolderWindow::Pane::Left;
        std::optional<FolderWindow::Pane> destinationPane;
        std::filesystem::path destinationFolder;
        std::filesystem::path diagnosticsLogPath;

        HRESULT resultHr             = S_OK;
        unsigned long totalItems     = 0;
        unsigned long completedItems = 0;
        uint64_t totalBytes          = 0;
        uint64_t completedBytes      = 0;

        // When pre-calc is skipped, totals may be unknown; keep a best-effort top-level type breakdown for UI.
        bool preCalcSkipped            = false;
        unsigned long completedFiles   = 0;
        unsigned long completedFolders = 0;
        std::wstring sourcePath;
        std::wstring destinationPath;

        bool autoConcurrencyUsed                       = false;
        uint32_t autoConcurrencyStorageKind            = FILESYSTEM_STORAGE_UNKNOWN;
        uint32_t autoConcurrencyDestinationStorageKind = FILESYSTEM_STORAGE_UNKNOWN;
        unsigned int autoTunedConcurrency              = 0;
        unsigned int effectiveConcurrencyBudget        = 0;

        unsigned long warningCount = 0;
        unsigned long errorCount   = 0;
        std::wstring lastDiagnosticMessage;
        std::vector<TaskDiagnosticEntry> issueDiagnostics;

        ULONGLONG lastProgressCallbackTick = 0;
        ULONGLONG completedTick            = 0;
    };

    struct Task final : public IFileSystemCallback, public IFileSystemDirectorySizeCallback
    {
        // Maximum number of in-flight file lines the popup can display for a single task.
        // This should be >= the Copy/Move worker concurrency cap so parallel file copies can be represented.
        static constexpr size_t kMaxInFlightFiles = 16u;

        enum class ConflictBucket : uint8_t
        {
            Exists,
            ReadOnly,
            AccessDenied,
            SharingViolation,
            DiskFull,
            PathTooLong,
            RecycleBinFailed,
            NetworkOffline,
            UnsupportedReparse,
            Unknown,
            Count,
        };

        enum class ConflictAction : uint8_t
        {
            None,
            Overwrite,
            ReplaceReadOnly,
            PermanentDelete,
            Retry,
            Skip,
            SkipAll,
            Cancel,
        };

        struct PerItemCallbackCookie
        {
            size_t itemIndex = 0;
            std::wstring lastProgressSourcePath;
            std::wstring lastProgressDestinationPath;
            std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> issueRetryCounts{};
        };

        struct ConflictPromptState
        {
            static constexpr size_t kMaxActions = 8u;

            bool active           = false;
            ConflictBucket bucket = ConflictBucket::Unknown;
            HRESULT status        = S_OK;
            std::wstring sourcePath;
            std::wstring destinationPath;
            std::array<ConflictAction, kMaxActions> actions{};
            size_t actionCount     = 0;
            bool applyToAllChecked = false;
            bool retryFailed       = false;
        };

        struct ConflictArbiter
        {
            ConflictArbiter()                                  = default;
            ConflictArbiter(const ConflictArbiter&)            = delete;
            ConflictArbiter(ConflictArbiter&&)                 = delete;
            ConflictArbiter& operator=(const ConflictArbiter&) = delete;
            ConflictArbiter& operator=(ConflictArbiter&&)      = delete;
            ~ConflictArbiter()                                 = default;

            std::mutex mutex;
            std::condition_variable cv;
            std::array<std::optional<ConflictAction>, static_cast<size_t>(ConflictBucket::Count)> decisionCache{};
            ConflictPromptState prompt{};
            DWORD ownerThreadId = 0;
            std::optional<ConflictAction> decisionAction;
            bool decisionApplyToAll = false;
            wil::unique_event_nothrow decisionEvent;
        };

        struct InFlightFileProgress
        {
            const void* cookieKey     = nullptr;
            uint64_t progressStreamId = 0;
            std::wstring sourcePath;
            uint64_t totalBytes      = 0;
            uint64_t completedBytes  = 0;
            ULONGLONG lastUpdateTick = 0;
        };

        struct ProgressStreamPerf
        {
            const void* cookieKey           = nullptr;
            uint64_t progressStreamId       = 0;
            uint64_t callbackCount          = 0;
            uint64_t callbackUs             = 0;
            uint64_t lockWaitUs             = 0;
            uint64_t callbackGapCount       = 0;
            uint64_t callbackGapMs          = 0;
            uint64_t callbackGapBytes       = 0;
            uint64_t maxCallbackGapMs       = 0;
            uint64_t maxCallbackGapBytes    = 0;
            uint64_t lastItemCompletedBytes = 0;
            ULONGLONG firstUpdateTick       = 0;
            ULONGLONG lastUpdateTick        = 0;
        };

        struct ConflictWorkerPerf
        {
            const void* cookieKey    = nullptr;
            uint64_t promptCount     = 0;
            uint64_t waitUs          = 0;
            ULONGLONG lastUpdateTick = 0;
        };

        struct DiagnosticPathSnapshot
        {
            std::wstring progressSourcePath;
            std::wstring progressDestinationPath;
            std::wstring lastProgressCallbackSourcePath;
            std::wstring lastProgressCallbackDestinationPath;
        };

        struct PerfStats
        {
            PerfStats()                            = default;
            PerfStats(const PerfStats&)            = delete;
            PerfStats(PerfStats&&)                 = delete;
            PerfStats& operator=(const PerfStats&) = delete;
            PerfStats& operator=(PerfStats&&)      = delete;

            uint64_t queueWaitUs                         = 0;
            uint64_t schedulerWaitUs                     = 0;
            uint64_t schedulerWaitForWorkUs              = 0;
            uint64_t schedulerProcessIndexUs             = 0;
            uint64_t schedulerDequeueAttempts            = 0;
            uint64_t schedulerDequeueSuccess             = 0;
            uint64_t bridgeCopyUs                        = 0;
            uint64_t bridgeReaderWaitUs                  = 0;
            uint64_t bridgeWriterWaitUs                  = 0;
            uint64_t bridgeReadUs                        = 0;
            uint64_t bridgeWriteUs                       = 0;
            uint64_t bridgeDirectoryEnsureCount          = 0;
            uint64_t bridgeFileAdmissionCount            = 0;
            uint64_t bridgeFileStartedBeforeProducerDone = 0;
            uint64_t bridgeAdmissionMaxQueueDepth        = 0;
            std::atomic<uint64_t> preCalcUs{0};
            std::atomic<uint64_t> preCalcCallbackCount{0};
            std::atomic<uint64_t> preCalcCallbackUs{0};
            std::atomic<uint64_t> preCalcLockWaitUs{0};
            uint64_t progressCallbackUs                  = 0;
            uint64_t progressFirstCallbackDelayMs        = 0;
            uint64_t progressLockWaitUs                  = 0;
            uint64_t progressLockHoldUs                  = 0;
            uint64_t progressLockContentionCount         = 0;
            uint64_t progressPathUpdateBytes             = 0;
            uint64_t progressPathUpdateAppliedCount      = 0;
            uint64_t progressPathUpdateSkippedCount      = 0;
            uint64_t progressPathUpdateThrottledCount    = 0;
            uint64_t progressInFlightEvictions           = 0;
            uint64_t perItemInFlightEvictions            = 0;
            uint64_t pauseWaitUs                         = 0;
            uint64_t conflictWaitUs                      = 0;
            uint64_t conflictConvergenceWaitUs           = 0;
            uint64_t conflictPromptCount                 = 0;
            uint64_t queueEnterCount                     = 0;
            uint64_t queueNotifyAllCount                 = 0;
            uint64_t queueCancelWhileWaiting             = 0;
            uint64_t queueDepthOnEnter                   = 0;
            uint64_t queueActiveOperations               = 0;
            uint64_t itemCompletedCallbackUs             = 0;
            uint64_t itemCompletedLockWaitUs             = 0;
            uint64_t itemCompletedLockHoldUs             = 0;
            uint64_t itemCompletedLockContentionCount    = 0;
            uint64_t itemCompletedPathUpdateBytes        = 0;
            uint64_t itemCompletedPathUpdateAppliedCount = 0;
            uint64_t itemCompletedPathUpdateSkippedCount = 0;
        };

        explicit Task(FileOperationState& state) noexcept;

        Task(const Task&)            = delete;
        Task(Task&&)                 = delete;
        Task& operator=(const Task&) = delete;
        Task& operator=(Task&&)      = delete;
        ~Task()                      = default;

        // IFileSystemCallback
        HRESULT STDMETHODCALLTYPE FileSystemProgress(FileSystemOperation operationType,
                                                     unsigned long totalItems,
                                                     unsigned long completedItems,
                                                     uint64_t totalBytes,
                                                     uint64_t completedBytes,
                                                     const wchar_t* currentSourcePath,
                                                     const wchar_t* currentDestinationPath,
                                                     uint64_t currentItemTotalBytes,
                                                     uint64_t currentItemCompletedBytes,
                                                     FileSystemOptions* options,
                                                     uint64_t progressStreamId,
                                                     void* cookie) noexcept override;

        HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(FileSystemOperation operationType,
                                                          unsigned long itemIndex,
                                                          const wchar_t* sourcePath,
                                                          const wchar_t* destinationPath,
                                                          HRESULT status,
                                                          FileSystemOptions* options,
                                                          void* cookie) noexcept override;

        HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, void* cookie) noexcept override;

        HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation operationType,
                                                  const wchar_t* sourcePath,
                                                  const wchar_t* destinationPath,
                                                  HRESULT status,
                                                  FileSystemIssueAction* action,
                                                  FileSystemOptions* options,
                                                  void* cookie) noexcept override;

        // IFileSystemDirectorySizeCallback
        HRESULT STDMETHODCALLTYPE DirectorySizeProgress(uint64_t scannedEntries,
                                                        uint64_t totalBytes,
                                                        uint64_t fileCount,
                                                        uint64_t directoryCount,
                                                        const wchar_t* currentPath,
                                                        void* cookie) noexcept override;

        HRESULT STDMETHODCALLTYPE DirectorySizeShouldCancel(BOOL* pCancel, void* cookie) noexcept override;

        void ThreadMain(std::stop_token stopToken) noexcept;
        void RunPreCalculation() noexcept;
        void SkipPreCalculation() noexcept;
        void RequestCancel() noexcept;
        void TogglePause() noexcept;
        void SetDesiredSpeedLimit(uint64_t bytesPerSecond) noexcept;
        void SetWaitForOthers(bool wait) noexcept;
        void SetWaitingInQueue(bool waiting) noexcept;
        void SetQueuePaused(bool paused) noexcept;
        void ToggleConflictApplyToAllChecked() noexcept;
        void SubmitConflictDecision(ConflictAction action, bool applyToAllChecked) noexcept;

        bool HasStarted() const noexcept;
        bool HasEnteredOperation() const noexcept;
        ULONGLONG GetEnteredOperationTick() const noexcept;
        bool IsPaused() const noexcept;
        bool IsWaitingForOthers() const noexcept;
        bool IsWaitingInQueue() const noexcept;
        bool IsQueuePaused() const noexcept;

        void SetDestinationFolder(const std::filesystem::path& folder);
        std::filesystem::path GetDestinationFolder() const;

        unsigned long GetPlannedItemCount() const noexcept;

        uint64_t GetId() const noexcept;
        HRESULT GetResult() const noexcept;

        FileSystemOperation GetOperation() const noexcept;
        FolderWindow::Pane GetSourcePane() const noexcept;
        std::optional<FolderWindow::Pane> GetDestinationPane() const noexcept;

        void WaitWhilePaused() noexcept;
        void WaitWhilePreCalcPaused() noexcept;
        void MarkRateSamplingStateChanged() noexcept;

        HRESULT ExecuteOperation() noexcept;
        HRESULT TryStartPreCalculationThread(std::jthread& preCalcThread) noexcept;
        void InitializeFileSystemOptions(FileSystemOptions& options) const noexcept;
        void LogDiagnostic(DiagnosticSeverity severity,
                           HRESULT status,
                           std::wstring_view category,
                           std::wstring_view message,
                           std::wstring_view sourcePath      = {},
                           std::wstring_view destinationPath = {}) noexcept;

        static HRESULT BuildPathArrayArena(const std::vector<std::filesystem::path>& paths,
                                           FileSystemArenaOwner& arenaOwner,
                                           const wchar_t*** outPaths,
                                           unsigned long* outCount) noexcept;

        FileOperationState* _state     = nullptr;
        FolderWindow* _folderWindow    = nullptr;
        uint64_t _taskId               = 0;
        FileSystemOperation _operation = FILESYSTEM_COPY;
        ExecutionMode _executionMode   = ExecutionMode::BulkItems;
        FolderWindow::Pane _sourcePane = FolderWindow::Pane::Left;
        std::optional<FolderWindow::Pane> _destinationPane;
        wil::com_ptr<IFileSystem> _fileSystem;
        wil::com_ptr<IFileSystem> _destinationFileSystem;
        std::vector<std::filesystem::path> _sourcePaths;
        std::vector<DWORD> _sourcePathAttributesHint;
        mutable std::mutex _operationMutex;
        std::filesystem::path _destinationFolder;
        FileSystemFlags _flags                  = FILESYSTEM_FLAG_NONE;
        bool _enablePreCalc                     = true;
        unsigned int _preCalcMaxWorkers         = 4u;
        unsigned long _crossFsBridgeBufferBytes = 4096u * 1024u;
        std::atomic<unsigned long> _resolvedCrossFsBridgeBufferBytes{0};
        std::atomic<unsigned int> _preCalcWorkerCountUsed{0};
        std::atomic<bool> _autoConcurrencyUsed{false};
        std::atomic<uint32_t> _autoConcurrencyStorageKind{FILESYSTEM_STORAGE_UNKNOWN};
        std::atomic<uint32_t> _autoConcurrencyDestinationStorageKind{FILESYSTEM_STORAGE_UNKNOWN};
        std::atomic<unsigned int> _autoTunedConcurrency{0};
        std::atomic<unsigned int> _effectiveConcurrencyBudget{0};

        unsigned long _perItemTotalItems          = 0;
        unsigned int _perItemMaxConcurrencyBudget = 1;
        unsigned int _perItemMaxConcurrency       = 1;
        unsigned long _perItemCompletedItems      = 0;
        uint64_t _perItemCompletedEntryCount      = 0;
        uint64_t _perItemTotalEntryCount          = 0;
        uint64_t _perItemCompletedBytes           = 0;

        struct PerItemInFlightCall
        {
            const void* cookie           = nullptr;
            unsigned long completedItems = 0;
            uint64_t completedBytes      = 0;
            unsigned long totalItems     = 0;
            ULONGLONG lastUpdateTick     = 0;
        };

        std::mutex _perItemInFlightCallsMutex;
        std::array<PerItemInFlightCall, kMaxInFlightFiles> _perItemInFlightCalls{};
        size_t _perItemInFlightCallCount        = 0;
        uint64_t _perItemInFlightCompletedBytes = 0;
        uint64_t _perItemInFlightCompletedItems = 0;
        uint64_t _perItemInFlightTotalItems     = 0;

        enum class TopLevelItemKind : uint8_t
        {
            Unknown,
            File,
            Folder,
        };
        std::mutex _topLevelCompletionMutex;
        std::vector<TopLevelItemKind> _topLevelItemKinds;
        std::vector<uint8_t> _topLevelItemCompleted;
        unsigned long _plannedTopLevelFiles     = 0;
        unsigned long _plannedTopLevelFolders   = 0;
        unsigned long _completedTopLevelFiles   = 0;
        unsigned long _completedTopLevelFolders = 0;

        std::atomic<bool> _waitForOthers{false};
        std::atomic<bool> _waitingInQueue{false};
        std::atomic<bool> _enteredOperation{false};
        std::atomic<ULONGLONG> _enteredOperationTick{0};
        std::atomic<bool> _cancelled{false};
        std::atomic<ULONGLONG> _cancelRequestedTick{0};
        std::atomic<bool> _paused{false};
        std::atomic<bool> _queuePaused{false};
        std::atomic<ULONGLONG> _rateSamplingStateChangeTick{0};
        std::atomic<bool> _started{false};
        std::atomic<ULONGLONG> _operationStartTick{0};
        std::atomic<uint64_t> _desiredSpeedLimitBytesPerSecond{0};
        std::atomic<uint64_t> _appliedSpeedLimitBytesPerSecond{0};
        std::atomic<uint64_t> _effectiveSpeedLimitBytesPerSecond{0};
        std::atomic<HRESULT> _resultHr{S_OK};
        std::atomic<bool> _taskFinished{false};
        std::atomic<bool> _observedSkipAction{false};

        ConflictArbiter _conflictArbiter;

        // Pre-calculation state
        std::atomic<bool> _preCalcInProgress{false};
        std::atomic<bool> _preCalcSkipped{false};
        std::atomic<bool> _preCalcCompleted{false};
        std::atomic<ULONGLONG> _preCalcStartTick{0};
        std::atomic<uint64_t> _preCalcTotalBytes{0};
        std::atomic<unsigned long> _preCalcFileCount{0};
        std::atomic<unsigned long> _preCalcDirectoryCount{0};
        std::vector<uint64_t> _preCalcSourceBytes;
        // 5F early admission: latched true the first time a transfer progress callback fires while
        // pre-calc has not yet completed (i.e. bytes moved before the recursive scan finished).
        // Impossible in the old serial pre-calc-then-execute model; the deterministic proof that
        // the transfer overlapped pre-calc.
        std::atomic<bool> _transferStartedBeforePreCalcComplete{false};

        std::stop_token _stopToken{};
        std::mutex _pauseMutex;
        std::condition_variable _pauseCv;

        std::mutex _progressMutex;
        unsigned long _progressTotalItems     = 0;
        unsigned long _progressCompletedItems = 0;
        uint64_t _progressTotalBytes          = 0;
        uint64_t _progressCompletedBytes      = 0;
        uint64_t _progressItemTotalBytes      = 0;
        uint64_t _progressItemCompletedBytes  = 0;
        std::atomic<unsigned long> _publishedProgressTotalItems{0};
        std::atomic<unsigned long> _publishedProgressCompletedItems{0};
        std::atomic<uint64_t> _publishedProgressTotalBytes{0};
        std::atomic<uint64_t> _publishedProgressCompletedBytes{0};
        std::atomic<uint64_t> _publishedProgressItemTotalBytes{0};
        std::atomic<uint64_t> _publishedProgressItemCompletedBytes{0};
        std::atomic<unsigned long> _publishedCompletedTopLevelFiles{0};
        std::atomic<unsigned long> _publishedCompletedTopLevelFolders{0};
        std::mutex _progressPathMutex;
        std::wstring _progressSourcePath;
        std::wstring _progressDestinationPath;
        std::wstring _lastProgressCallbackSourcePath;
        std::wstring _lastProgressCallbackDestinationPath;
        std::atomic<std::shared_ptr<const DiagnosticPathSnapshot>> _publishedDiagnosticPathSnapshot{};
        ULONGLONG _lastVisibleProgressPathUpdateTick = 0;
        ULONGLONG _lastProgressCallbackTick          = 0;
        unsigned long _lastItemIndex                 = 0;
        HRESULT _lastItemHr                          = S_OK;
        std::atomic<uint64_t> _progressCallbackCount{0};
        std::atomic<uint64_t> _itemCompletedCallbackCount{0};

        std::mutex _inFlightFilesMutex;
        std::array<InFlightFileProgress, kMaxInFlightFiles> _inFlightFiles{};
        size_t _inFlightFileCount = 0;
        std::mutex _progressStreamPerfMutex;
        std::array<ProgressStreamPerf, kMaxInFlightFiles> _progressStreamPerf{};
        size_t _progressStreamPerfCount = 0;
        std::array<ConflictWorkerPerf, kMaxInFlightFiles> _conflictWorkerPerf{};
        size_t _conflictWorkerPerfCount = 0;

        PerfStats _perf{};
        std::atomic<uint64_t> _bridgeDirectoryEnsureCount{0};
        std::atomic<uint64_t> _bridgeFileAdmissionCount{0};
        std::atomic<uint64_t> _bridgeFileStartedBeforeProducerDone{0};
        std::atomic<uint64_t> _bridgeAdmissionMaxQueueDepth{0};

#ifdef ENABLE_TESTS
        unsigned int _dbgConfiguredMaxConcurrency      = 1;
        ULONGLONG _dbgSingleInFlightStartTick          = 0;
        ULONGLONG _dbgLastSingleInFlightWarnTick       = 0;
        bool _dbgObservedMultipleInFlightFiles         = false;
        ULONGLONG _dbgLastPerItemInFlightEvictWarnTick = 0;
        std::atomic_uint32_t _dbgCallbackActiveScopeCount{0};
#endif

        std::jthread _thread;
    };

    explicit FileOperationState(FolderWindow& owner);

    FileOperationState(const FileOperationState&)            = delete;
    FileOperationState(FileOperationState&&)                 = delete;
    FileOperationState& operator=(const FileOperationState&) = delete;
    FileOperationState& operator=(FileOperationState&&)      = delete;

    ~FileOperationState();

    HRESULT StartOperation(FileSystemOperation operation,
                           FolderWindow::Pane sourcePane,
                           std::optional<FolderWindow::Pane> destinationPane,
                           const wil::com_ptr<IFileSystem>& fileSystem,
                           std::vector<std::filesystem::path> sourcePaths,
                           std::filesystem::path destinationFolder,
                           FileSystemFlags flags,
                           bool waitForOthers,
                           uint64_t initialSpeedLimitBytesPerSecond        = 0,
                           ExecutionMode executionMode                     = ExecutionMode::BulkItems,
                           bool requireConfirmation                        = false,
                           wil::com_ptr<IFileSystem> destinationFileSystem = nullptr,
                           uint64_t* taskIdOut                             = nullptr);

    void ApplyTheme(const AppTheme& theme);
    void Shutdown() noexcept;
    void NotifyQueueChanged();
    bool HasActiveOperations() noexcept;
    bool ShouldQueueNewTask() noexcept;
    void SetQueueNewTasks(bool queue) noexcept;
    bool GetQueueNewTasks() const noexcept;
    void ApplyQueueMode(bool queue) noexcept;
    void CancelAll() noexcept;
    void CollectTasks(std::vector<Task*>& outTasks) noexcept;
    void CollectInformationalTasks(std::vector<FolderWindow::InformationalTaskUpdate>& outTasks) noexcept;
    void CollectCompletedTasks(std::vector<CompletedTaskSummary>& outTasks) noexcept;
    void CollectDiagnostics(std::vector<TaskDiagnosticEntry>& outEntries) noexcept;
    void CollectTaskDiagnosticSnapshot(uint64_t taskId,
                                       unsigned long& warningCount,
                                       unsigned long& errorCount,
                                       std::wstring& lastDiagnosticMessage) noexcept;
    void DismissCompletedTask(uint64_t taskId) noexcept;
    uint64_t CreateOrUpdateInformationalTask(const FolderWindow::InformationalTaskUpdate& update) noexcept;
    void DismissInformationalTask(uint64_t taskId) noexcept;
    bool GetAutoDismissSuccess() const noexcept;
    void SetAutoDismissSuccess(bool enabled) noexcept;
    bool OpenDiagnosticsLogForTask(uint64_t taskId) noexcept;
    bool ExportTaskIssuesReport(uint64_t taskId, std::filesystem::path* reportPathOut = nullptr, bool openAfterExport = true) noexcept;
    void ToggleIssuesPane() noexcept;
    bool IsIssuesPaneVisible() noexcept;
    bool TryGetIssuesPanePlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept;
    void SaveIssuesPanePlacement(HWND hwnd) noexcept;
    bool TryGetIssuesPaneViewState(std::wstring& outSortColumnId,
                                   bool& outSortDescending,
                                   std::vector<Common::Settings::GridColumnLayoutEntry>& outGridLayout) const noexcept;
    void SaveIssuesPaneViewState(std::wstring_view sortColumnId,
                                 bool sortDescending,
                                 const std::vector<Common::Settings::GridColumnLayoutEntry>& gridLayout) noexcept;
    bool TryGetPopupPlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept;
    void SavePopupPlacement(HWND hwnd) noexcept;
    void OnPopupDestroyed(HWND hwnd) noexcept;
    void OnIssuesPaneDestroyed(HWND hwnd) noexcept;
    void UpdateLastPopupRect(const RECT& rect) noexcept;
    std::optional<RECT> GetLastPopupRect() noexcept;
#ifdef ENABLE_TESTS
    HWND GetPopupHwndForSelfTest() noexcept;
    HWND GetIssuesPaneHwndForSelfTest() noexcept;
    void DebugResetIssuesPaneForSelfTest() noexcept;
    void DebugClearDiagnosticsForSelfTest() noexcept;
    void DebugRemoveDiagnosticsForTask(uint64_t taskId) noexcept;
    void DebugAppendCompletedTaskForSelfTest(CompletedTaskSummary summary) noexcept;
#endif

    void RecordTaskDiagnostic(uint64_t taskId,
                              FileSystemOperation operation,
                              DiagnosticSeverity severity,
                              HRESULT status,
                              std::wstring_view category,
                              std::wstring_view message,
                              std::wstring_view sourcePath,
                              std::wstring_view destinationPath) noexcept;
    void EnqueueTaskDiagnostic(TaskDiagnosticEntry entry) noexcept;

    bool EnterOperation(Task& task, std::stop_token stopToken) noexcept;
    void LeaveOperation() noexcept;
    void PostCompleted(Task& task) noexcept;

    Task* FindTask(uint64_t taskId) noexcept;
    void RemoveTask(uint64_t taskId) noexcept;

private:
    void EnsurePopupVisible() noexcept;
    CompletedTaskSummary RecordCompletedTask(Task& task) noexcept;
    void FlushDiagnostics(bool force) noexcept;
    static std::filesystem::path GetDiagnosticsLogDirectory() noexcept;
    static std::filesystem::path GetDiagnosticsLogPathForDate(const SYSTEMTIME& localTime) noexcept;
    std::filesystem::path GetLatestDiagnosticsLogPathUnlocked() const noexcept;
    void RemoveFromQueue(uint64_t taskId) noexcept;
    void UpdateQueuePausedTasks() noexcept;

    FolderWindow& _owner;
    std::mutex _mutex;
    std::shared_ptr<void> _uiLifetime;
    std::vector<std::unique_ptr<Task>> _tasks;
    std::vector<FolderWindow::InformationalTaskUpdate> _informationalTasks;
    std::deque<CompletedTaskSummary> _completedTasks;
    uint64_t _nextTaskId = 1;

    wil::unique_hwnd _popup;
    wil::unique_hwnd _issuesPane;
    std::optional<RECT> _lastPopupRect;

    std::mutex _queueMutex;
    std::condition_variable _queueCv;
    std::deque<uint64_t> _queue;
    unsigned long _activeOperations = 0;

    std::mutex _diagnosticsMutex;
    std::deque<TaskDiagnosticEntry> _diagnosticsInMemory;
    std::vector<TaskDiagnosticEntry> _diagnosticsPendingFlush;
    std::unordered_map<uint64_t, std::pair<unsigned long, unsigned long>> _taskDiagnosticCounts;
    std::unordered_map<uint64_t, std::wstring> _taskLastDiagnosticMessage;
    std::unordered_map<uint64_t, std::deque<TaskDiagnosticEntry>> _taskIssueDiagnostics;
    ULONGLONG _lastDiagnosticsFlushTick   = 0;
    ULONGLONG _lastDiagnosticsCleanupTick = 0;

    std::mutex _followTargetsWarningMutex;
    bool _followTargetsWarningPromptActive = false;
    bool _followTargetsWarningAccepted     = false;

    std::atomic<bool> _queueNewTasks{true};
};

// pluginId is diagnostic-only: it identifies the provider in the one-time capabilities-contract-violation error.
[[nodiscard]] bool CanSameFileSystemOperation(const wil::com_ptr<IFileSystem>& fileSystem,
                                              FileSystemOperation operation,
                                              std::wstring_view pluginId = {}) noexcept;
[[nodiscard]] bool IsAutoDismissableFileOperationCompletion(HRESULT resultHr, unsigned long warningCount, unsigned long errorCount) noexcept;

#ifdef ENABLE_TESTS
enum class FileOpsBridgePipelineMode : unsigned char
{
    Default,
    Disabled,
    Enabled,
};

void SetFileOpsBridgePipelineModeForSelfTest(FileOpsBridgePipelineMode mode) noexcept;
FileOpsBridgePipelineMode GetFileOpsBridgePipelineModeForSelfTest() noexcept;
void SetFileOpsBridgeProducerDelayForSelfTest(unsigned int delayMs) noexcept;
unsigned int GetFileOpsBridgeProducerDelayForSelfTest() noexcept;
void SetFileOpsBridgeFailNextFileCopiesForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest() noexcept;
void SetFileOpsPreCalcThreadStartFailureForSelfTest(bool enabled) noexcept;
unsigned long TakeFileOpsPreCalcThreadStartAttemptsForSelfTest() noexcept;
void SetFileOpsAutoConcurrencyOverrideForSelfTest(bool enabled, unsigned int preferredConcurrency, uint32_t storageKind) noexcept;
void SetFileOpsPostFinishedCompletionPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsPostFinishedCompletionPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsPostFinishedCompletionPauseForSelfTest() noexcept;
bool RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsBridgeDirectoryBufferValidationSelfTestForSelfTest() noexcept;
#endif
