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
        uint64_t taskId = 0;
        HRESULT hr      = S_OK;
    };

    struct TaskDiagnosticEntry
    {
        SYSTEMTIME localTime{};
        uint64_t taskId                         = 0;
        FileSystemOperation operation           = FILESYSTEM_COPY;
        DiagnosticSeverity severity             = DiagnosticSeverity::Info;
        HRESULT status                          = S_OK;
        unsigned __int64 processWorkingSetBytes = 0;
        unsigned __int64 processPrivateBytes    = 0;
        std::wstring category;
        std::wstring message;
        std::wstring sourcePath;
        std::wstring destinationPath;
    };

    struct CompletedTaskSummary
    {
        uint64_t taskId               = 0;
        FileSystemOperation operation = FILESYSTEM_COPY;
        FolderWindow::Pane sourcePane = FolderWindow::Pane::Left;
        std::optional<FolderWindow::Pane> destinationPane;
        std::filesystem::path destinationFolder;
        std::filesystem::path diagnosticsLogPath;

        HRESULT resultHr                = S_OK;
        unsigned long totalItems        = 0;
        unsigned long completedItems    = 0;
        unsigned __int64 totalBytes     = 0;
        unsigned __int64 completedBytes = 0;
        std::wstring sourcePath;
        std::wstring destinationPath;

        unsigned long warningCount = 0;
        unsigned long errorCount   = 0;
        std::wstring lastDiagnosticMessage;
        std::vector<TaskDiagnosticEntry> issueDiagnostics;

        ULONGLONG completedTick = 0;
    };

    struct Task final : public IFileSystemCallback, public IFileSystemDirectorySizeCallback
    {
        // Maximum number of in-flight file lines the popup can display for a single task.
        // This should be >= the Copy/Move worker concurrency cap so parallel file copies can be represented.
        static constexpr size_t kMaxInFlightFiles = 8u;

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

        struct InFlightFileProgress
        {
            const void* cookieKey             = nullptr;
            unsigned __int64 progressStreamId = 0;
            std::wstring sourcePath;
            unsigned __int64 totalBytes     = 0;
            unsigned __int64 completedBytes = 0;
            ULONGLONG lastUpdateTick        = 0;
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
                                                     unsigned __int64 totalBytes,
                                                     unsigned __int64 completedBytes,
                                                     const wchar_t* currentSourcePath,
                                                     const wchar_t* currentDestinationPath,
                                                     unsigned __int64 currentItemTotalBytes,
                                                     unsigned __int64 currentItemCompletedBytes,
                                                     FileSystemOptions* options,
                                                     unsigned __int64 progressStreamId,
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
        HRESULT STDMETHODCALLTYPE DirectorySizeProgress(unsigned __int64 scannedEntries,
                                                        unsigned __int64 totalBytes,
                                                        unsigned __int64 fileCount,
                                                        unsigned __int64 directoryCount,
                                                        const wchar_t* currentPath,
                                                        void* cookie) noexcept override;

        HRESULT STDMETHODCALLTYPE DirectorySizeShouldCancel(BOOL* pCancel, void* cookie) noexcept override;

        void ThreadMain(std::stop_token stopToken) noexcept;
        void RunPreCalculation() noexcept;
        void SkipPreCalculation() noexcept;
        void RequestCancel() noexcept;
        void TogglePause() noexcept;
        void SetDesiredSpeedLimit(unsigned __int64 bytesPerSecond) noexcept;
        void SetWaitForOthers(bool wait) noexcept;
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

        HRESULT ExecuteOperation() noexcept;
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
        FileSystemFlags _flags = FILESYSTEM_FLAG_NONE;
        bool _enablePreCalc    = true;

        unsigned long _perItemTotalItems             = 0;
        unsigned int _perItemMaxConcurrency          = 1;
        unsigned long _perItemCompletedItems         = 0;
        unsigned __int64 _perItemCompletedEntryCount = 0;
        unsigned __int64 _perItemTotalEntryCount     = 0;
        unsigned __int64 _perItemCompletedBytes      = 0;

        struct PerItemInFlightCall
        {
            const void* cookie              = nullptr;
            unsigned long completedItems    = 0;
            unsigned __int64 completedBytes = 0;
            unsigned long totalItems        = 0;
        };

        std::array<PerItemInFlightCall, kMaxInFlightFiles> _perItemInFlightCalls{};
        size_t _perItemInFlightCallCount = 0;

        std::atomic<bool> _waitForOthers{false};
        std::atomic<bool> _waitingInQueue{false};
        std::atomic<bool> _enteredOperation{false};
        std::atomic<ULONGLONG> _enteredOperationTick{0};
        std::atomic<bool> _cancelled{false};
        std::atomic<ULONGLONG> _cancelRequestedTick{0};
        std::atomic<bool> _paused{false};
        std::atomic<bool> _queuePaused{false};
        std::atomic<bool> _started{false};
        std::atomic<ULONGLONG> _operationStartTick{0};
        std::atomic<unsigned __int64> _desiredSpeedLimitBytesPerSecond{0};
        std::atomic<unsigned __int64> _appliedSpeedLimitBytesPerSecond{0};
        std::atomic<unsigned __int64> _effectiveSpeedLimitBytesPerSecond{0};
        std::atomic<HRESULT> _resultHr{S_OK};
        std::atomic<bool> _observedSkipAction{false};

        std::mutex _conflictMutex;
        std::condition_variable _conflictCv;
        std::array<std::optional<ConflictAction>, static_cast<size_t>(ConflictBucket::Count)> _conflictDecisionCache{};
        ConflictPromptState _conflictPrompt{};
        std::optional<ConflictAction> _conflictDecisionAction;
        bool _conflictDecisionApplyToAll = false;
        wil::unique_event_nothrow _conflictDecisionEvent;

        // Pre-calculation state
        std::atomic<bool> _preCalcInProgress{false};
        std::atomic<bool> _preCalcSkipped{false};
        std::atomic<bool> _preCalcCompleted{false};
        std::atomic<ULONGLONG> _preCalcStartTick{0};
        std::atomic<unsigned __int64> _preCalcTotalBytes{0};
        std::atomic<unsigned long> _preCalcFileCount{0};
        std::atomic<unsigned long> _preCalcDirectoryCount{0};
        std::vector<unsigned __int64> _preCalcSourceBytes;

        std::stop_token _stopToken{};
        std::mutex _pauseMutex;
        std::condition_variable _pauseCv;

        std::mutex _progressMutex;
        unsigned long _progressTotalItems            = 0;
        unsigned long _progressCompletedItems        = 0;
        unsigned __int64 _progressTotalBytes         = 0;
        unsigned __int64 _progressCompletedBytes     = 0;
        unsigned __int64 _progressItemTotalBytes     = 0;
        unsigned __int64 _progressItemCompletedBytes = 0;
        std::wstring _progressSourcePath;
        std::wstring _progressDestinationPath;
        std::wstring _lastProgressCallbackSourcePath;
        std::wstring _lastProgressCallbackDestinationPath;
        unsigned long _lastItemIndex                 = 0;
        HRESULT _lastItemHr                          = S_OK;
        unsigned __int64 _progressCallbackCount      = 0;
        unsigned __int64 _itemCompletedCallbackCount = 0;

        std::array<InFlightFileProgress, kMaxInFlightFiles> _inFlightFiles{};
        size_t _inFlightFileCount = 0;

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
                           unsigned __int64 initialSpeedLimitBytesPerSecond = 0,
                           ExecutionMode executionMode                      = ExecutionMode::BulkItems,
                           bool requireConfirmation                         = false,
                           wil::com_ptr<IFileSystem> destinationFileSystem  = nullptr);

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
    void CollectCompletedTasks(std::vector<CompletedTaskSummary>& outTasks) noexcept;
    void DismissCompletedTask(uint64_t taskId) noexcept;
    bool GetAutoDismissSuccess() const noexcept;
    void SetAutoDismissSuccess(bool enabled) noexcept;
    bool OpenDiagnosticsLogForTask(uint64_t taskId) noexcept;
    bool ExportTaskIssuesReport(uint64_t taskId, std::filesystem::path* reportPathOut = nullptr, bool openAfterExport = true) noexcept;
    void ToggleIssuesPane() noexcept;
    bool IsIssuesPaneVisible() noexcept;
    bool TryGetIssuesPanePlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept;
    void SaveIssuesPanePlacement(HWND hwnd) noexcept;
    bool TryGetPopupPlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept;
    void SavePopupPlacement(HWND hwnd) noexcept;
    void OnPopupDestroyed(HWND hwnd) noexcept;
    void OnIssuesPaneDestroyed(HWND hwnd) noexcept;
    void UpdateLastPopupRect(const RECT& rect) noexcept;
    std::optional<RECT> GetLastPopupRect() noexcept;
#ifdef _DEBUG
    HWND GetPopupHwndForSelfTest() noexcept;
#endif

    void RecordTaskDiagnostic(uint64_t taskId,
                              FileSystemOperation operation,
                              DiagnosticSeverity severity,
                              HRESULT status,
                              std::wstring_view category,
                              std::wstring_view message,
                              std::wstring_view sourcePath,
                              std::wstring_view destinationPath) noexcept;

    bool EnterOperation(Task& task, std::stop_token stopToken) noexcept;
    void LeaveOperation() noexcept;
    void PostCompleted(Task& task) noexcept;

    Task* FindTask(uint64_t taskId) noexcept;
    void RemoveTask(uint64_t taskId) noexcept;

private:
    void CreateProgressDialog(Task& task) noexcept;
    void RecordCompletedTask(Task& task) noexcept;
    void FlushDiagnostics(bool force) noexcept;
    static std::filesystem::path GetDiagnosticsLogDirectory() noexcept;
    static std::filesystem::path GetDiagnosticsLogPathForDate(const SYSTEMTIME& localTime) noexcept;
    std::filesystem::path GetLatestDiagnosticsLogPathUnlocked() const noexcept;
    void RemoveFromQueue(uint64_t taskId) noexcept;
    void UpdateQueuePausedTasks() noexcept;

    FolderWindow& _owner;
    std::mutex _mutex;
    std::vector<std::unique_ptr<Task>> _tasks;
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
