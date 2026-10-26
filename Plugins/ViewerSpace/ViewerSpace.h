#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <array>
#include <atomic>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <memory>
#include <memory_resource>
#include <mutex>
#include <optional>
#include <span>
#include <stop_token>
#include <string>
#include <string_view>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include <d2d1_1.h>
#include <d3d11.h>
#include <dwrite.h>
#include <dxgi1_2.h>
#include <dxgiformat.h>

#pragma comment(lib, "d2d1")
#pragma comment(lib, "dwrite")

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182) // WIL: deleted copy/move operators and padding
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "EmbeddedViewerBase.h"
#include "FileSystemPathIdentity.h"
#include "Helpers.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"
#include "ViewerMenuHost.h"
#include "ViewerSpace.AnalysisHost.h"
#include "ViewerSpace.Memory.h"
#include "ViewerSpace.NavigationHost.h"
#include "ViewerSpace.ScanPolicy.h"
#include "WindowMessages.h"
#include <DxUi/NativeMenuInterop.h>

[[nodiscard]] const char* GetViewerSpaceStaticConfigurationSchema() noexcept;
void ShutdownViewerSpaceModuleState() noexcept;
[[nodiscard]] bool CanUnloadViewerSpaceModuleNow() noexcept;
[[nodiscard]] bool ViewerSpaceWorkspaceExportsIdle() noexcept;

class ViewerSpace final : public EmbeddedViewerBase<ViewerSpace>, public IInformations
{
public:
    ViewerSpace();
    ~ViewerSpace();

    void SetHost(IHost* host) noexcept;

    ViewerSpace(const ViewerSpace&)            = delete;
    ViewerSpace(ViewerSpace&&)                 = delete;
    ViewerSpace& operator=(const ViewerSpace&) = delete;
    ViewerSpace& operator=(ViewerSpace&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override;

    HRESULT STDMETHODCALLTYPE Open(const ViewerOpenContext* context) noexcept override;
    HRESULT STDMETHODCALLTYPE Close() noexcept override;
    HRESULT STDMETHODCALLTYPE SetTheme(const ViewerTheme* theme) noexcept override;

private:
    struct LayoutBuildState;
    struct RetiredDisplayStorage;
    struct RetiredD3DDevice;

    enum class ScanState : uint8_t
    {
        NotStarted,
        Queued,
        Scanning,
        Done,
        Error,
        Canceled,
    };

    enum class HeaderHit : uint8_t
    {
        None,
        Cancel,
    };

    enum class ViewMode : uint8_t
    {
        Explorer,
        Analyst
    };

    struct Config final
    {
        uint32_t topFilesPerDirectory        = 96;
        uint32_t scanThreads                 = 1;
        uint32_t maxConcurrentScansPerVolume = 1;
        bool cacheEnabled                    = true;
        uint32_t cacheTtlSeconds             = 60;
        uint32_t cacheMaxEntries             = 1;
        ViewMode viewMode                    = ViewMode::Explorer;
        bool analysisPanelVisible            = true;
        float analysisPanelRatio             = 0.33f;
        float analysisBottomRatio            = 0.62f;
        bool inspectorVisible                = true;
        float inspectorRatio                 = 0.38f;
        bool compactRows                     = false;
        uint32_t colorMode                   = 0u; // Folder groups, type, own Modified, newest retained file.
        bool fileCountWeight                 = false;
    };

    struct FileSummaryItem final
    {
        uint32_t nodeId = 0;
        std::wstring name;
        uint64_t bytes        = 0;
        int64_t lastWriteTime = 0;
    };

    struct Node final
    {
        uint32_t id             = 0;
        uint32_t parentId       = 0;
        bool isDirectory        = false;
        bool isSynthetic        = false;
        bool contentNotMeasured = false;
        ScanState scanState     = ScanState::NotStarted;

        std::wstring_view name;

        uint64_t totalBytes           = 0;
        int64_t lastWriteTime         = 0;
        uint64_t observedOwnFileBytes = 0;
        uint64_t observedChildBytes   = 0;
        uint32_t childrenStart        = 0;
        uint32_t childrenCount        = 0;
        uint32_t childrenCapacity     = 0;
        uint64_t aggregateBytes       = 0;
        uint64_t aggregateFolders     = 0;
        uint64_t aggregateFiles       = 0;
    };

    static constexpr uint32_t kFileRecordIdBase    = 0x40000000u;
    static constexpr uint32_t kSyntheticNodeIdBase = 0x80000000u;
    static constexpr uint32_t kTreemapItemIdMask   = 0x3FFFFFFFu;

    struct FileRecord final
    {
        uint32_t id       = 0;
        uint32_t parentId = 0;
        std::wstring_view name;
        uint64_t bytes        = 0;
        int64_t lastWriteTime = 0;
    };

    struct ResolvedItem final
    {
        uint32_t id             = 0;
        uint32_t parentId       = 0;
        bool isDirectory        = false;
        bool isSynthetic        = false;
        bool isFileRecord       = false;
        bool contentNotMeasured = false;
        ScanState scanState     = ScanState::Done;
        std::wstring_view name;
        uint64_t totalBytes       = 0;
        int64_t lastWriteTime     = 0;
        uint64_t aggregateFolders = 0;
        uint64_t aggregateFiles   = 0;
    };

    struct DrawItem final
    {
        enum class Lod : uint8_t
        {
            Culled,
            Tiny,
            Small,
            Medium,
            Large,
            Hero,
        };

        uint32_t nodeId      = 0;
        uint8_t depth        = 0;
        Lod lod              = Lod::Culled;
        float areaDip2       = 0.0f;
        float labelHeightDip = 0.0f;
        D2D1_RECT_F targetRect{};
        D2D1_RECT_F currentRect{};
        D2D1_RECT_F startRect{};
        double animationStartSeconds = 0.0;
    };

    struct LayoutItem final
    {
        uint32_t nodeId = 0;
        double weight   = 0.0;
        uint64_t bytes  = 0;
    };

    struct LayoutExpandTask final
    {
        uint32_t nodeId = 0;
        D2D1_RECT_F bounds{};
        float area    = 0.0f;
        uint8_t depth = 0;
    };

    struct LayoutWorkspace final
    {
        std::vector<LayoutItem> items;
        std::vector<LayoutItem> topItems;
        std::vector<LayoutItem> forcedItems;
        std::vector<LayoutItem> row;
        std::vector<uint32_t> forcedChildIds;
        std::vector<LayoutExpandTask> expandTasks;
    };

    struct LayoutCandidateCacheEntry final
    {
        uint64_t signature = 0u;
        uint32_t maxItems  = 0u;
        std::vector<LayoutItem> items;
        uint64_t otherBytes   = 0u;
        double otherWeight    = 0.0;
        size_t otherCount     = 0u;
        uint64_t otherFolders = 0u;
        uint64_t otherFiles   = 0u;
    };

    struct PendingUpdate final
    {
        enum class Kind : uint8_t
        {
            AddChild,
            UpdateSize,
            UpdateState,
            UpdateMetadata,
            DirectorySummaryBegin,
            DirectorySummaryChunk,
            DirectorySummaryEnd,
            DirectorySummaryReject,
            Progress,
            ResourceSummary,
            PathIdentity,
        };

        Kind kind             = Kind::UpdateSize;
        uint32_t generation   = 0;
        uint32_t nodeId       = 0;
        uint32_t parentId     = 0;
        uint64_t bytes        = 0;
        int64_t lastWriteTime = 0;
        ScanState state       = ScanState::NotStarted;
        std::wstring name;
        std::optional<FileSystemPathIdentity> pathIdentity;
        bool isDirectory        = false;
        bool isSynthetic        = false;
        bool contentNotMeasured = false;

        uint64_t scannedFolders           = 0;
        uint64_t scannedFiles             = 0;
        uint64_t failedFolders            = 0;
        uint32_t firstPreviewPriorityJobs = 0u;

        uint64_t otherBytes            = 0;
        uint64_t otherFolders          = 0;
        uint64_t otherFiles            = 0;
        uint64_t summaryRevision       = 0u;
        uint64_t summaryNameBytes      = 0u;
        uint32_t summarySequence       = 0u;
        uint32_t summaryExpectedChunks = 0u;
        uint32_t summaryExpectedFiles  = 0u;
        std::vector<FileSummaryItem> topFiles;

        uint64_t acceptedEntries                         = 0;
        uint64_t rejectedEntries                         = 0;
        uint64_t cappedDirectories                       = 0;
        uint64_t cappedFiles                             = 0;
        uint64_t retainedNameBytes                       = 0;
        uint64_t retainedChildReferences                 = 0;
        uint64_t traversedDirectories                    = 0;
        ViewerSpaceScan::ValidationError validationError = ViewerSpaceScan::ValidationError::None;
    };

    struct DirectorySummaryStageKey final
    {
        uint32_t generation                                                           = 0u;
        uint32_t nodeId                                                               = 0u;
        uint64_t revision                                                             = 0u;
        [[nodiscard]] bool operator==(const DirectorySummaryStageKey&) const noexcept = default;
    };

    struct DirectorySummaryStageKeyHash final
    {
        [[nodiscard]] size_t operator()(const DirectorySummaryStageKey& key) const noexcept
        {
            uint64_t value = key.revision ^ (static_cast<uint64_t>(key.generation) << 32u) ^ static_cast<uint64_t>(key.nodeId);
            value ^= value >> 30u;
            value *= 0xBF58476D1CE4E5B9ull;
            value ^= value >> 27u;
            value *= 0x94D049BB133111EBull;
            value ^= value >> 31u;
            return static_cast<size_t>(value);
        }
    };

    struct DirectorySummaryStage final
    {
        uint64_t otherBytes            = 0u;
        uint64_t stagedFileBytes       = 0u;
        uint64_t otherFolders          = 0u;
        uint64_t otherFiles            = 0u;
        uint64_t expectedNameBytes     = 0u;
        uint64_t stagedNameBytes       = 0u;
        size_t fileRecordStart         = 0u;
        uint32_t expectedChunks        = 0u;
        uint32_t expectedFiles         = 0u;
        uint32_t receivedChunks        = 0u;
        uint32_t receivedFiles         = 0u;
        uint32_t childArenaStart       = 0u;
        uint32_t childArenaCapacity    = 0u;
        uint32_t previousChildStart    = 0u;
        uint32_t previousChildCount    = 0u;
        uint32_t previousChildCapacity = 0u;
        uint32_t copiedChildren        = 0u;
        uint64_t accountedBytes        = 0u;
        bool hasChildArena             = false;
        bool endReceived               = false;
    };

    struct FileRecordFreeBlock final
    {
        size_t start = 0u;
        size_t count = 0u;
    };

    struct ChildArenaFreeBlock final
    {
        uint32_t start    = 0u;
        uint32_t capacity = 0u;
    };

    // A generation keeps every borrowed name view dominated by the PMR resources that own it.
    // The stable owner address is adopted as a unit; its PMR containers are never swapped across allocators.
    struct ModelGeneration final
    {
        ViewerSpaceDetail::CountingMemoryResource memory;
        std::pmr::monotonic_buffer_resource nameArena{&memory};
        std::pmr::monotonic_buffer_resource layoutNameArena{&memory};
        // These contiguous buffers grow geometrically. Return replaced allocations immediately;
        // MSVC's default pool also retains large size-class chunks until generation destruction.
        std::pmr::vector<Node> nodes{&memory};
        std::pmr::vector<FileRecord> fileRecords{&memory};
        std::pmr::vector<uint32_t> childrenArena{&memory};
        std::pmr::unordered_map<DirectorySummaryStageKey, DirectorySummaryStage, DirectorySummaryStageKeyHash> directorySummaryStages{&memory};
        std::pmr::deque<DirectorySummaryStageKey> directorySummaryCopyQueue{&memory};
        std::pmr::unordered_map<uint32_t, ScanState> deferredDirectoryDoneStates{&memory};
        std::pmr::unordered_set<uint32_t> rejectedDirectorySummaryNodes{&memory};
        std::pmr::vector<FileRecordFreeBlock> fileRecordFreeBlocks{&memory};
        std::vector<ChildArenaFreeBlock> childrenArenaFreeBlocks;
        uint32_t generation = 0u;
        uint32_t rootNodeId = 0u;
        uint32_t viewNodeId = 0u;
        std::wstring scanRootPath;
        std::optional<FileSystemPathIdentity> workspacePathIdentity;
        std::optional<std::wstring> scanRootParentPath;
        uint64_t directorySummaryStageBytes = 0u;

        ModelGeneration()                                  = default;
        ModelGeneration(const ModelGeneration&)            = delete;
        ModelGeneration& operator=(const ModelGeneration&) = delete;
        [[nodiscard]] uint64_t AllocatedBytes() const noexcept;
    };

    struct RefreshDisplayState final
    {
        ScanState overallState             = ScanState::NotStarted;
        double completedSinceSeconds       = 0.0;
        uint64_t progressBytes             = 0u;
        uint64_t progressFolders           = 0u;
        uint64_t progressFiles             = 0u;
        uint64_t failedFolders             = 0u;
        uint32_t effectiveThreads          = 1u;
        uint32_t topFilesPerDirectory      = 0u;
        uint32_t cacheLastStoredGeneration = 0u;
        uint64_t cacheSnapshotBytes        = 0u;
        uint64_t cacheRestoreBytes         = 0u;
        uint64_t cacheResidentBytes        = 0u;
        uint64_t cacheRestoreUs            = 0u;
        uint32_t processingNodeId          = 0u;
        std::wstring processingFolderName;
        uint64_t acceptedEntries                         = 0u;
        uint64_t rejectedEntries                         = 0u;
        uint64_t cappedDirectories                       = 0u;
        uint64_t cappedFiles                             = 0u;
        uint64_t retainedNameBytes                       = 0u;
        uint64_t retainedChildReferences                 = 0u;
        uint64_t traversedDirectories                    = 0u;
        uint64_t summaryCommittedCount                   = 0u;
        uint64_t summaryRejectedCount                    = 0u;
        uint64_t summaryChunkMaxBytes                    = 0u;
        ViewerSpaceScan::ValidationError validationError = ViewerSpaceScan::ValidationError::None;
        uint32_t viewNodeId                              = 0u;
        uint32_t selectedItemId                          = 0u;
        bool selectionFromAnalysis                       = false;
        std::wstring viewPathText;
        std::vector<uint32_t> navStack;
        std::vector<uint32_t> navForwardStack;
        std::vector<std::wstring> selectedRelativePath;
        size_t selectionPathPart     = 0u;
        uint32_t selectionPathNodeId = 0u;
        bool selectionPathIsFile     = false;
    };

    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static void UnregisterWndClassIfIdle() noexcept;
    static constexpr wchar_t kClassName[] = L"RedSalamander.ViewerSpace";

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnCreate(HWND hwnd);
    void OnDestroy();
    void OnSize(UINT width, UINT height) noexcept;
    void OnPaint();
    enum class InputPresentationKind : uint8_t
    {
        Hover,
        Selection,
        Navigation,
        Count
    };
    void BeginInputPresentation(InputPresentationKind kind, std::chrono::steady_clock::time_point startedAt) noexcept;
    void CompletePendingInputPresentation(HRESULT result, uint64_t visibleTileCount) noexcept;
    std::array<std::chrono::steady_clock::time_point, static_cast<size_t>(InputPresentationKind::Count)> _pendingInputPresentation{};
    uint64_t _inputPresentationCount  = 0u;
    uint64_t _lastInputPresentationUs = 0u;
    uint32_t _lastRenderPresentHr     = static_cast<uint32_t>(S_OK);
    void OnCommand(HWND hwnd, UINT commandId) noexcept;
    void OnKeyDown(WPARAM vk, bool alt) noexcept;
    void OnMouseMove(int x, int y) noexcept;
    void OnMouseLeave() noexcept;
    void OnLButtonDown(int x, int y) noexcept;
    void OnLButtonDblClk(int x, int y) noexcept;
    void OnContextMenu(HWND hwnd, POINT screenPt) noexcept;
    void OnDpiChanged(HWND hwnd, UINT dpi, const RECT* suggestedRect) noexcept;
    void OnNcActivate(HWND hwnd, bool windowActive) noexcept;
    LRESULT OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept;
    void OnTimer(UINT_PTR timerId) noexcept;
    void CreateAnalysisHost(HWND hwnd);
    void LayoutAnalysisHost() noexcept;
    void CreateNavigationHost(HWND hwnd);
    void LayoutNavigationHost() noexcept;
    void UpdateNavigationHostState() noexcept;
    [[nodiscard]] std::vector<ViewerSpaceNavigation::NavigationHost::Crumb> GetNavigationFolders(uint64_t token) const;
#ifdef ENABLE_TESTS
    LRESULT DebugNavigationFolder(WndMsg::ViewerSpaceNavigationFolderDebugQuery* query);
#endif
    void ContinueAnalysisBuild() noexcept;
    [[nodiscard]] uint64_t AnalysisOwnedBytes() const noexcept;
    void InvalidateAnalysis() noexcept;
    void SelectItem(uint32_t itemId, bool syncGrid = true) noexcept;
    void UpdateInspector() noexcept;
    [[nodiscard]] bool AnalysisFilterMatches(uint32_t id) const noexcept;
    void SetViewMode(ViewMode mode) noexcept;
    void ToggleAnalysisPanel() noexcept;
    void ResizeAnalysisPanel(DxUi::SplitterChange change) noexcept;
    void UpdateConfigurationJson();
    HRESULT ApplyConfiguration(const char* configurationJsonUtf8, bool publishSharedPolicy) noexcept;
    [[nodiscard]] uint64_t AnalysisRowToken(size_t index) const noexcept;
    [[nodiscard]] std::optional<size_t> FindAnalysisRow(uint64_t token) const noexcept;
    void GetAnalysisCell(size_t row, size_t column, DxUi::GridCellData& cell) const;
    [[nodiscard]] std::optional<ResolvedItem> ResolveAnalysisItem(uint32_t id) const noexcept;
    void ActivateItem(uint32_t id) noexcept;
    void FocusItemInPane(uint32_t id) noexcept;
    void CopyItemPath() noexcept;
    void ToggleExcludedItem(uint32_t id) noexcept;
    void RestoreExcludedItems() noexcept;
    void ContinueWorkspaceBindings() noexcept;
    void WorkspaceChanged() noexcept;
    [[nodiscard]] bool IsItemExcluded(uint32_t id) const noexcept;
    [[nodiscard]] uint64_t ShownBytes(uint32_t id, uint64_t original) const noexcept;
    [[nodiscard]] uint64_t WorkspaceOwnedBytes() const noexcept;
    [[nodiscard]] bool CanAdmitWorkspaceBytes(uint64_t additional) const noexcept;
    void ShowWorkspaceMessage(UINT stringId) noexcept;
    void ExportCurrentView(bool csv) noexcept;
    [[nodiscard]] bool QueueWorkspaceExport(std::wstring_view outputPath, bool csv) noexcept;
#ifdef ENABLE_TESTS
    LRESULT DebugQueueWorkspaceExport(const WndMsg::ViewerSpaceExportDebugRequest* request) noexcept;
#endif
    struct WorkspaceExportBuild;
    void ContinueWorkspaceExport() noexcept;
    void SubmitWorkspaceExport(bool save) noexcept;
    void HandleWorkspaceExportCompletion(HWND hwnd, LPARAM token) noexcept;
    static void CALLBACK WorkspaceExportCallback(PTP_CALLBACK_INSTANCE instance, void* context) noexcept;
    static constexpr UINT kWorkspaceExportCompleted = WM_APP + 0x660;
    void ShowScanProblems() noexcept;
    void ToggleCleanupMark(uint32_t id, bool notify = true) noexcept;
    void MarkMatchingItems() noexcept;
    void ReviewCleanup() noexcept;
    void SendReviewedCleanup(bool recycle) noexcept;
    void OpenFolderInOtherPane() noexcept;
    void UpdateWorkspaceScope() noexcept;
    void ContinueWorkspaceFacts() noexcept;
    [[nodiscard]] uint64_t MapWeight(uint32_t id, uint64_t logicalBytes) const noexcept;
    [[nodiscard]] uint64_t ShownFileCount(uint32_t id) const noexcept;
    [[nodiscard]] uint32_t ItemCategory(const ResolvedItem& item) const noexcept;
    [[nodiscard]] int64_t ItemAgeDate(const ResolvedItem& item) const noexcept;
    [[nodiscard]] bool MatchesWorkspaceFilter(const ResolvedItem& item) const noexcept;
    [[nodiscard]] D2D1_COLOR_F CategoryColor(uint32_t category) const noexcept;
    void ShowWorkspaceFilters(bool breakdown) noexcept;
    void NavigateHistory(bool forward) noexcept;

    bool EnsureDirect2D(HWND hwnd) noexcept;
    void DiscardDirect2D() noexcept;
    void InvalidateStaticTreemapCache() noexcept;
    bool CanReplayStaticTreemapCache(bool scanActive, bool completionOverlayActive) const noexcept;
    bool CanRecordStaticTreemapCache(bool scanActive, bool completionOverlayActive, double nowSeconds) const noexcept;
    bool HasActiveLayoutAnimation(double nowSeconds) const noexcept;
    HRESULT EnsureStaticTreemapCacheBitmap() noexcept;
    HRESULT RecordStaticTreemapCacheFromCurrentTarget() noexcept;
    void PaintHoverOverlay(const D2D1::ColorF& accentColor) noexcept;
    void ApplyThemeToWindow(HWND hwnd) noexcept;
    void ApplyTitleBarTheme(HWND hwnd, bool windowActive) noexcept;
    void UpdateWindowTitle(HWND hwnd) noexcept;
    void ApplyMenuTheme(HWND hwnd) noexcept;
    void UpdateTooltipForHit(uint32_t nodeId) noexcept;
    void RefreshTooltipText() noexcept;
    void OnSettingChange(const wchar_t* area) noexcept;
    void UpdateTooltipPosition(int x, int y) noexcept;
    void PaintTooltipOverlay() noexcept;
    std::wstring BuildTooltipText(uint32_t nodeId) const;

    void UpdateHeaderTextCache() noexcept;

    void StartScan(std::wstring_view rootPath, bool allowCache = true, bool retainPrevious = false);
    void RebuildCurrent() noexcept;
    void AdoptWorkingModel() noexcept;
    void DiscardWorkingModel(bool refused = false) noexcept;
    void ContinueRefreshAdoption() noexcept;
    void ContinueDeferredStartScan() noexcept;
    void ContinueDeferredModelRetirement() noexcept;
    void ContinueCancelledLayoutBuildRetirement() noexcept;
    void CancelDeferredStartScan() noexcept;
    [[nodiscard]] bool HasDeferredDisposalCapacity(bool closing = false) noexcept;
    void CancelScan() noexcept;
    void CancelScanByUser() noexcept;
    void AbandonScanWorkers() noexcept;
    void ReapFinishedScanWorkers() noexcept;
    void ContinueDeferredDisposalSubmissions() noexcept;
    void ReapFinishedDeferredDisposals() noexcept;
    void AbandonDeferredDisposals() noexcept;
    [[nodiscard]] bool SubmitDeferredDisposal(size_t index) noexcept;
    void RetireDisplayedStorage(std::unique_ptr<ModelGeneration> displayModel,
                                std::unique_ptr<ModelGeneration> workingModel    = {},
                                bool includeDerived                              = true,
                                bool closing                                     = false,
                                std::unique_ptr<RetiredD3DDevice> rendererDevice = {}) noexcept;
    [[nodiscard]] std::unique_ptr<RetiredD3DDevice> CaptureClosingRendererDevice() noexcept;
    [[nodiscard]] static uint32_t RetiredRendererDeviceCount() noexcept;
    void RetireLayoutStorage(std::shared_ptr<LayoutBuildState> layoutBuild,
                             std::shared_ptr<LayoutBuildState> publishedLayoutOwner,
                             bool cancellation = false) noexcept;
    void QueueDeferredDisposal(std::unique_ptr<RetiredDisplayStorage> storage, uint64_t bytes, bool closing) noexcept;
    static VOID CALLBACK DeferredDisposalCallback(PTP_CALLBACK_INSTANCE instance, void* context) noexcept;
    void ScanMain(std::stop_token stopToken,
                  uint32_t generation,
                  wil::com_ptr<IFileSystem> fileSystem,
                  bool fileSystemIsWin32,
                  std::wstring rootPath,
                  uint32_t rootNodeId,
                  uint32_t nextNodeId,
                  size_t topFilesPerDirectory,
                  uint32_t scanThreads,
                  ViewerSpaceScan::ResourcePolicy resourcePolicy) noexcept;

    void PostUpdate(PendingUpdate&& update) noexcept;
    void DrainUpdates() noexcept;
    void EmitQueueDiagnostics(bool terminalSnapshot, bool force) noexcept;
    bool IsQueueTerminalReady() noexcept;
    static size_t EstimatePendingUpdateBytes(const PendingUpdate& update) noexcept;
    void ContinueScanCacheBuild() noexcept;
    void CancelScanCacheBuild() noexcept;
    void ContinueScanCacheRestore() noexcept;
    void QueueScanCacheRestoreContinuation(uint32_t generation) noexcept;
    void CancelScanCacheRestore() noexcept;
    const Node* TryGetRealNode(uint32_t nodeId) const noexcept;
    Node* TryGetRealNode(uint32_t nodeId) noexcept;
    const FileRecord* TryGetFileRecord(uint32_t itemId) const noexcept;
    std::optional<ResolvedItem> ResolveTreemapItem(uint32_t itemId) const noexcept;
    static bool IsFileRecordId(uint32_t itemId) noexcept;
    static bool IsSyntheticNodeId(uint32_t itemId) noexcept;
    static uint32_t FileRecordIdFromIndex(size_t index) noexcept;
    static uint32_t SyntheticOtherBucketIdForParent(uint32_t parentId) noexcept;
    static uint64_t SubtractAtomicFloor(std::atomic_uint64_t& value, uint64_t delta) noexcept;
    std::span<const uint32_t> GetRealNodeChildren(const Node& node) const noexcept;
    [[nodiscard]] bool AddRealNodeChild(Node& parent, uint32_t childItemId) noexcept;
    // UI-thread model projection: preserve monotonic provider totals and propagate only new child lower-bound bytes.
    [[nodiscard]] bool RaiseNodeObservedBytes(uint32_t nodeId, uint64_t reportedBytes) noexcept;
    [[nodiscard]] std::optional<uint32_t> AllocateChildArenaBlock(uint32_t capacity) noexcept;
    void ReleaseChildArenaBlock(uint32_t start, uint32_t capacity) noexcept;
    std::wstring BuildNodePathText(uint32_t nodeId) const;
    std::wstring BuildItemPathText(uint32_t itemId) const;
    void UpdateViewPathText() noexcept;

    void EnsureLayoutForView() noexcept;
    void MaybeRebuildLayout() noexcept;
    [[nodiscard]] bool CanBeginScanProjection(bool hasPublishedProjection) const noexcept;
    void RebuildLayout() noexcept;
    void CancelLayoutBuild() noexcept;
    [[nodiscard]] D2D1_RECT_F GetTreemapBounds() const noexcept;
    [[nodiscard]] uint64_t ViewerOwnedBytes() const noexcept;
    [[nodiscard]] uint64_t LayoutOwnedBytes() const noexcept;
    [[nodiscard]] uint64_t LayoutBuildOwnedBytes(const LayoutBuildState& build) const noexcept;
    static DrawItem::Lod ClassifyDrawItemLod(const D2D1_RECT_F& itemRc, uint8_t depth) noexcept;
    void BuildHitGrid(const D2D1_RECT_F& bounds) noexcept;
    std::optional<uint32_t> HitTestTreemap(float xDip, float yDip) const noexcept;
    std::optional<uint32_t> HitTestTreemapLinear(float xDip, float yDip, uint32_t& candidatesChecked) const noexcept;

    void NavigateTo(uint32_t nodeId) noexcept;
    void NavigateUp() noexcept;
    void RefreshCurrent() noexcept;
    bool CanNavigateUp() const noexcept;
    void UpdateMenuState(HWND hwnd, bool syncDxMenuBar = true) noexcept;
    float GetMenuBarHeightDip() const noexcept;
    float GetHeaderTopDip() const noexcept;
    float GetHeaderBottomDip() const noexcept;
    uint32_t ComputeAdaptiveFileCandidateBudget() const noexcept;

    float DipFromPx(int px) const noexcept;
    int PxFromDip(float dip) const noexcept;
    double NowSeconds() const noexcept;

private:
    std::atomic_ulong _refCount{1};

    PluginMetaData _metaData{};
    std::wstring _metaId;
    std::wstring _metaShortId;
    std::wstring _metaName;
    std::wstring _metaDescription;

    std::string _configurationJson;
    Config _config{};

    wil::com_ptr<IHostPaneExecute> _hostPaneExecute;
    wil::com_ptr<IHostAlerts> _hostAlerts;
    wil::com_ptr<IHostViewerActions> _hostViewerActions;

    wil::com_ptr<IFileSystem> _fileSystem;
    std::wstring _fileSystemName;
    std::wstring _fileSystemShortId;
    bool _fileSystemIsWin32 = true;

    bool _allowEraseBkgnd = true;
    wil::unique_hmenu _menuHandle;
    RedSalamander::ViewerMenuHost::AttachResult _menuAttachResult;
    DxUi::NativeMenuBarHost _menuBarHost;
    ViewerSpaceAnalysis::AnalysisHost _analysisHost;
    ViewerSpaceNavigation::NavigationHost _navigationHost;
    D2D1_RECT_F _treemapBounds{};
    D2D1_RECT_F _analysisBounds{};
    bool _analysisSidePanel = false;
    std::optional<float> _analysisResizeStartRatio;
    struct AnalysisBuildState;
    std::shared_ptr<AnalysisBuildState> _analysisBuild;
    std::shared_ptr<AnalysisBuildState> _publishedAnalysisOwner;
    std::vector<uint32_t> _analysisRows;
    uint32_t _analysisGeneration    = 0u;
    uint32_t _analysisViewId        = 0u;
    uint32_t _selectedItemId        = 0u;
    uint32_t _selectedGeneration    = 0u;
    bool _selectionFromAnalysis     = false;
    uint32_t _contextItemIdOverride = 0u;
    std::wstring _analysisFilter;
    double _analysisFilterDueSeconds     = 0.0;
    double _analysisLastPublishedSeconds = 0.0;
    bool _analysisDirty                  = true;
    size_t _analysisSortColumn           = 1u;
    bool _analysisSortAscending          = false;
    uint64_t _analysisCellRequests       = 0u;
    uint64_t _analysisMaxTurnUs          = 0u;

    float _dpi = static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    SIZE _clientSize{};

    struct RendererResources final
    {
        wil::com_ptr<ID3D11Device> d3dDevice;
        wil::com_ptr<ID3D11DeviceContext> d3dContext;
        wil::com_ptr<IDXGIFactory2> dxgiFactory;
        wil::com_ptr<ID2D1Device> d2dDevice;
        wil::com_ptr<ID2D1DeviceContext> d2dContext;
        wil::com_ptr<IDXGISwapChain1> swapChain;
        wil::com_ptr<ID2D1Bitmap1> targetBitmap;
        UINT swapChainWidthPx          = 0;
        UINT swapChainHeightPx         = 0;
        D3D_FEATURE_LEVEL featureLevel = D3D_FEATURE_LEVEL_11_0;
    };

    wil::com_ptr<ID2D1Factory1> _d2dFactory;
    RendererResources _renderer;
    wil::com_ptr<ID2D1RenderTarget> _renderTarget;
    wil::com_ptr<ID2D1SolidColorBrush> _brushBackground;
    wil::com_ptr<ID2D1SolidColorBrush> _brushText;
    wil::com_ptr<ID2D1SolidColorBrush> _brushOutline;
    wil::com_ptr<ID2D1SolidColorBrush> _brushAccent;
    wil::com_ptr<ID2D1SolidColorBrush> _brushWatermark;
    wil::com_ptr<ID2D1LinearGradientBrush> _brushShading;
    wil::com_ptr<ID2D1GradientStopCollection> _shadingStops;
    wil::com_ptr<ID2D1StrokeStyle> _otherStrokeStyle;
    wil::com_ptr<ID2D1PathGeometry> _dogEarFlapGeometry;

    struct StaticTreemapCache final
    {
        wil::com_ptr<ID2D1Bitmap1> bitmap;
        UINT widthPx          = 0;
        UINT heightPx         = 0;
        uint64_t generation   = 0u;
        uint64_t hits         = 0u;
        uint64_t misses       = 0u;
        uint64_t recordCount  = 0u;
        uint64_t lastRecordUs = 0u;
        uint64_t bytes        = 0u;
    };
    StaticTreemapCache _staticTreemapCache;

    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    wil::com_ptr<IDWriteTextFormat> _textFormat;
    wil::com_ptr<IDWriteTextFormat> _headerFormat;
    wil::com_ptr<IDWriteTextFormat> _headerStatusFormatRight;
    wil::com_ptr<IDWriteTextFormat> _headerInfoFormat;
    wil::com_ptr<IDWriteTextFormat> _headerInfoFormatRight;
    wil::com_ptr<IDWriteTextFormat> _headerIconFormat;
    wil::com_ptr<IDWriteTextFormat> _watermarkFormat;
    wil::com_ptr<IDWriteTextFormat> _tooltipFormat;

    struct ScanWorker final
    {
        ScanWorker()                        = default;
        ~ScanWorker()                       = default;
        ScanWorker(ScanWorker&&)            = default;
        ScanWorker& operator=(ScanWorker&&) = default;

        ScanWorker(const ScanWorker&)            = delete;
        ScanWorker& operator=(const ScanWorker&) = delete;

        [[nodiscard]] bool IsComplete() const noexcept
        {
            return done != nullptr ? done->load(std::memory_order_acquire) : ! thread.joinable();
        }

        std::jthread thread;
        std::shared_ptr<std::atomic_bool> done;
    };

    std::unique_ptr<ModelGeneration> _displayModel;
    std::unique_ptr<ModelGeneration> _workingModel;
    // Normally aliases _displayModel. DrainUpdates temporarily binds it to _workingModel on the UI thread.
    ModelGeneration* _model                  = nullptr;
    uint64_t _directorySummaryCommittedCount = 0u;
    uint64_t _directorySummaryRejectedCount  = 0u;
    std::atomic_uint64_t _directorySummaryChunkMaxBytes{0u};
    uint32_t _displayGeneration  = 0u;
    bool _refreshPending         = false;
    bool _refreshRefused         = false;
    bool _refreshFailed          = false;
    bool _refreshAdoptionPending = false;
    std::optional<RefreshDisplayState> _refreshDisplayState;
    std::vector<uint32_t> _refreshPathCandidates;
    std::vector<uint32_t> _refreshNextPathCandidates;
    size_t _refreshPathCandidateIndex    = 0u;
    size_t _refreshPathChildIndex        = 0u;
    uint32_t _refreshResolvedSelectionId = 0u;
    bool _refreshPathLookupStarted       = false;
    uint32_t _workingScanGeneration      = 0u;
    struct DeferredStartScan final
    {
        std::wstring rootPath;
        uint64_t requestGeneration = 0u;
        bool allowCache            = true;
        bool retainPrevious        = false;
    };
    std::optional<DeferredStartScan> _deferredStartScan;
    uint64_t _deferredStartScanGeneration = 0u;
    bool _workingModelRetirementPending   = false;
    uint64_t ChildArenaFreeSlots() const noexcept;
    [[nodiscard]] std::optional<size_t> AllocateFileRecordRange(size_t count) noexcept;
    void ReleaseFileRecordRange(size_t start, size_t count) noexcept;
    [[nodiscard]] bool BeginDirectorySummary(const PendingUpdate& update) noexcept;
    [[nodiscard]] bool ApplyDirectorySummaryChunk(const PendingUpdate& update) noexcept;
    [[nodiscard]] bool EndDirectorySummary(const PendingUpdate& update) noexcept;
    [[nodiscard]] bool CommitDirectorySummary(const DirectorySummaryStageKey& key) noexcept;
    [[nodiscard]] bool RejectDirectorySummary(const DirectorySummaryStageKey& key, bool markRejected) noexcept;
    [[nodiscard]] size_t AdvanceDirectorySummaryStages(size_t copyBudget) noexcept;
    [[nodiscard]] bool HasDirectorySummaryStage(uint32_t generation, uint32_t nodeId) const noexcept;
    [[nodiscard]] bool HasPendingDirectorySummaryChildBlock(uint32_t generation, uint32_t nodeId) const noexcept;
    void ApplyDeferredDirectoryDoneStates() noexcept;
    void DiscardDirectorySummaryStages() noexcept;
    ViewerSpaceScan::ResourcePolicy _scanResourcePolicy = ViewerSpaceScan::kProductionResourcePolicy;

    ScanWorker _scanWorker;
    std::vector<ScanWorker> _retiredScanWorkers;
    std::atomic_uint32_t _scanGeneration{0};
    std::atomic_bool _scanActive{false};

    std::mutex _updateMutex;
    std::deque<PendingUpdate> _pendingUpdates;
#ifdef ENABLE_TESTS
    std::atomic<HANDLE> _debugPostUpdateEnteredEvent{nullptr};
    std::atomic<HANDLE> _debugPostUpdateReleaseEvent{nullptr};
    std::atomic_bool _debugPauseNextPostUpdate{false};
    std::atomic_uint64_t _debugPostUpdateInnerGenerationRejects{0u};
    uint64_t _debugLocaleInvalidations = 0u;
#endif

    std::shared_ptr<void> _scanCacheBuildSnapshot;
    std::wstring _scanCacheBuildRootKey;
    std::shared_ptr<const void> _scanCacheRestoreSnapshot;
    std::wstring _scanCacheRestoreRootKey;
    uint32_t _scanTopFilesPerDirectory             = 0;
    uint32_t _scanCacheBuildTopFilesPerDirectory   = 0;
    uint32_t _scanCacheBuildGeneration             = 0;
    uint32_t _scanCacheRestoreTopFilesPerDirectory = 0;
    uint32_t _scanCacheRestoreGeneration           = 0;
    uint32_t _scanCacheLastStoredGeneration        = 0;
    uint8_t _scanCacheRestorePhase                 = 0;
    bool _scanCacheRestorePaintSliceStarted        = false;
    struct ScanCacheRestoreContinuation final
    {
        uint32_t generation = 0u;
    };
    bool _scanCacheRestoreContinuationAttempted = false;
    bool _scanCacheRestoreContinuationPending   = false;
    size_t _scanCacheBuildPreflightNodesNext    = 0;
    size_t _scanCacheBuildPreflightFilesNext    = 0;
    size_t _scanCacheBuildNameChars             = 0;
    size_t _scanCacheBuildChildrenNext          = 0;
    size_t _scanCacheBuildNodesNext             = 0;
    size_t _scanCacheBuildFileRecordsNext       = 0;
    size_t _scanCacheRestoreChildrenNext        = 0;
    size_t _scanCacheRestoreNodesNext           = 0;
    size_t _scanCacheRestoreFileRecordsNext     = 0;
    size_t _scanCacheRestoreValidationNext      = 0;
    size_t _scanCacheRestoreValidationChildNext = 0;
    std::vector<uint64_t> _scanCacheRestoreSeenChildren;
    double _scanCacheRestoreStartedAtSeconds                   = 0.0;
    uint64_t _lastScanCacheSnapshotBytes                       = 0u;
    uint64_t _lastScanCacheRestoreBytes                        = 0u;
    uint64_t _lastScanCacheResidentBytes                       = 0u;
    uint64_t _lastScanCacheRestoreUs                           = 0u;
    uint64_t _lastModelAcceptedEntries                         = 0u;
    uint64_t _lastModelRejectedEntries                         = 0u;
    uint64_t _lastModelCappedDirectories                       = 0u;
    uint64_t _lastModelCappedFiles                             = 0u;
    uint64_t _lastModelRetainedNameBytes                       = 0u;
    uint64_t _lastModelRetainedChildReferences                 = 0u;
    uint64_t _lastModelTraversedDirectories                    = 0u;
    ViewerSpaceScan::ValidationError _lastModelValidationError = ViewerSpaceScan::ValidationError::None;

    std::unordered_map<uint32_t, Node> _syntheticNodes;
    std::unordered_map<uint32_t, uint32_t> _layoutMaxItemsByNode;
    std::unordered_set<uint32_t> _autoExpandedOtherByNode;
    std::unordered_map<uint32_t, LayoutCandidateCacheEntry> _layoutCandidateCache;
    uint64_t _layoutCandidateCacheHits   = 0u;
    uint64_t _layoutCandidateCacheMisses = 0u;
    std::wstring _viewPathText;
    std::vector<uint32_t> _navStack;
    std::vector<uint32_t> _navForwardStack;
    bool _lastLButtonDownNavigated = false;
    bool _animationTimerArmed      = false;
    void EnsureAnimationTimer() noexcept;
    [[nodiscard]] bool CanStopAnimationTimer(double nowSeconds) noexcept;

    ScanState _overallState           = ScanState::NotStarted;
    double _scanCompletedSinceSeconds = 0.0;

    uint64_t _scanProgressBytes       = 0;
    uint32_t _scanPreviewPriorityJobs = 0u;
    uint64_t _scanProgressFolders     = 0;
    uint64_t _scanProgressFiles       = 0;
    uint64_t _scanFailedFolders       = 0;
    uint32_t _scanEffectiveThreads    = 1;
    uint32_t _scanProcessingNodeId    = 0;
    std::wstring _scanProcessingFolderName;
    UINT _headerStatusId                = 0;
    uint64_t _headerStatusFailedFolders = 0;
    std::wstring _headerStatusText;
    std::wstring _headerCountsText;
    std::wstring _headerSizeText;
    std::wstring _headerViewText;
    std::wstring _headerProcessingText;
    std::wstring _scanInProgressWatermarkText;
    std::wstring _scanIncompleteWatermarkText;
    std::wstring _headerPathSourceText;
    std::wstring _headerPathDisplayText;
    float _headerPathDisplayMaxWidthDip = 0.0f;

    std::vector<DrawItem> _drawItems;
    std::shared_ptr<LayoutBuildState> _layoutBuild;
    // Exactly one canceled candidate may wait here for bounded disposal capacity.
    std::shared_ptr<LayoutBuildState> _cancelledLayoutBuild;
    // Published synthetic-name views are dominated by this owner until the next coherent layout adoption.
    std::shared_ptr<LayoutBuildState> _publishedLayoutOwner;
    uint64_t _modelRevision = 0u;
    // Session-only view intent. IDs and ancestor deductions are rebound to each display generation.
    // 128 paths, 512 ancestors/path and 32K UTF-16 units/path bound memory and lookup work.
    struct ExcludedPath final
    {
        std::wstring path;
        bool directory = false;
        uint32_t id    = 0u;
    };
    std::vector<ExcludedPath> _excludedPaths;
    std::unordered_map<uint32_t, std::vector<uint32_t>> _excludedBelow;
    std::unordered_set<uint32_t> _excludedRoots;
    std::unordered_map<uint32_t, std::vector<uint32_t>> _workspaceNextExcludedBelow;
    std::unordered_set<uint32_t> _workspaceNextExcludedRoots;
    uint32_t _workspaceBindingGeneration  = 0u;
    uint64_t _workspaceBindingRevision    = 0u;
    size_t _workspaceNextAncestorLinks    = 0u;
    bool _workspaceResetBindings          = false;
    uint32_t _workspaceGeneration         = 0u;
    uint64_t _workspaceBoundRevision      = 0u;
    uint64_t _workspaceProjectionRevision = 0u;
    size_t _workspaceBindingIndex         = 0u;
    uint32_t _workspaceBindingParent      = 0u;
    size_t _workspaceBindingChild         = 0u;
    size_t _workspaceBindingPart          = 0u;
    uint32_t _workspaceBindingMatch       = 0u;
    bool _workspaceBindingAmbiguous       = false;
    bool _workspaceRebinding              = false;
    std::vector<std::wstring> _workspaceBindingParts;
    size_t _workspaceUnresolved    = 0u;
    bool _workspaceViewExcluded    = false;
    size_t _workspaceAncestorLinks = 0u;
    std::array<uint64_t, 8u> _typeBreakdownBytes{};
    std::array<uint64_t, 8u> _typeBreakdownCounts{};
    std::vector<uint32_t> _cleanupMarks;
    uint32_t _cleanupGeneration = 0u;
    bool _reviewingCleanup      = false;
    bool _showingProblems       = false;
    std::vector<uint32_t> _problemRows;
    size_t _problemCursor = 0u;
    struct WorkspaceFacts final
    {
        uint64_t files       = 0u;
        int64_t newest       = 0;
        uint8_t types        = 0u;
        bool unknown         = false;
        bool unknownTypes    = false;
        bool countIncomplete = false;
    };
    std::vector<WorkspaceFacts> _workspaceFacts;
    std::vector<WorkspaceFacts> _workspaceFactsNext;
    uint32_t _workspaceFactsGeneration  = 0u;
    uint64_t _workspaceFactsRevision    = 0u;
    size_t _workspaceFactsFileCursor    = 0u;
    size_t _workspaceFactsNodeCursor    = 0u;
    bool _workspaceFactsReady           = false;
    bool _workspaceFactsBuilding        = false;
    uint32_t _typeFilter                = 0u;
    uint32_t _ageFilter                 = 0u;
    int64_t _workspaceNow               = 0;
    double _workspaceScanStartedSeconds = 0.0;
    D2D1_RECT_F _headerProblemsBounds{};
    std::shared_ptr<WorkspaceExportBuild> _workspaceExport;
    bool _exportProjectionRequested = false;
    bool _exportDialogOpen          = false;
    std::shared_ptr<const uint8_t> _workspaceLifetime{std::make_shared<const uint8_t>(uint8_t{0})};
    uint64_t _layoutTurnCount        = 0u;
    uint64_t _layoutMaxTurnUs        = 0u;
    uint32_t _layoutWorkUnitsPerTurn = 32768u;
    std::vector<LayoutWorkspace> _layoutWorkspaceByDepth;

    struct HitGrid final
    {
        D2D1_RECT_F bounds{};
        uint32_t columns              = 0u;
        uint32_t rows                 = 0u;
        uint32_t maxCandidatesPerCell = 0u;
        std::vector<std::vector<uint32_t>> cells;

        void Clear() noexcept
        {
            bounds               = {};
            columns              = 0u;
            rows                 = 0u;
            maxCandidatesPerCell = 0u;
            cells.clear();
        }
    };

    // UI-owned references are moved here at generation replacement. Storage is released on a
    // module-pinned callback so large PMR/index/vector destruction cannot monopolize the UI.
    struct RetiredD3DDevice final
    {
        explicit RetiredD3DDevice(wil::com_ptr<ID3D11Device> ownedDevice) noexcept : device(std::move(ownedDevice))
        {
        }
        ~RetiredD3DDevice() noexcept;
        RetiredD3DDevice(const RetiredD3DDevice&)            = delete;
        RetiredD3DDevice& operator=(const RetiredD3DDevice&) = delete;
        RetiredD3DDevice(RetiredD3DDevice&&)                 = delete;
        RetiredD3DDevice& operator=(RetiredD3DDevice&&)      = delete;
        wil::com_ptr<ID3D11Device> device;
    };

    struct RetiredDisplayStorage final
    {
        RetiredDisplayStorage()                                        = default;
        ~RetiredDisplayStorage()                                       = default;
        RetiredDisplayStorage(const RetiredDisplayStorage&)            = delete;
        RetiredDisplayStorage& operator=(const RetiredDisplayStorage&) = delete;
        RetiredDisplayStorage(RetiredDisplayStorage&&)                 = default;
        RetiredDisplayStorage& operator=(RetiredDisplayStorage&&)      = default;

        std::unique_ptr<ModelGeneration> displayModel;
        std::unique_ptr<ModelGeneration> workingModel;
        std::unique_ptr<RetiredD3DDevice> rendererDevice;
        bool layoutOnly              = false;
        bool layoutCancellation      = false;
        uint64_t layoutGeometryCount = 0u;
        uint32_t layoutGeneration    = 0u;
        std::shared_ptr<LayoutBuildState> layoutBuild;
        std::shared_ptr<LayoutBuildState> cancelledLayoutBuild;
        std::shared_ptr<LayoutBuildState> publishedLayoutOwner;
        std::shared_ptr<AnalysisBuildState> analysisBuild;
        std::shared_ptr<AnalysisBuildState> publishedAnalysisOwner;
        std::vector<uint32_t> analysisRows;
        std::unordered_map<uint32_t, Node> syntheticNodes;
        std::unordered_map<uint32_t, uint32_t> layoutMaxItemsByNode;
        std::unordered_set<uint32_t> autoExpandedOtherByNode;
        std::unordered_map<uint32_t, LayoutCandidateCacheEntry> layoutCandidateCache;
        std::vector<DrawItem> drawItems;
        std::vector<LayoutWorkspace> layoutWorkspaceByDepth;
        HitGrid hitGrid;
        std::unordered_map<uint32_t, D2D1_RECT_F> lastRectsByNode;

        void ReleaseStorage() noexcept;
    };

    struct DeferredDisposalWorker final
    {
        std::shared_ptr<RetiredDisplayStorage> storage;
        std::shared_ptr<std::atomic_bool> done;
        uint64_t bytes           = 0u;
        double retryAfterSeconds = 0.0;
        bool submitted           = false;
    };

    struct DeferredDisposalContext final
    {
        DeferredDisposalContext()                                          = default;
        ~DeferredDisposalContext()                                         = default;
        DeferredDisposalContext(const DeferredDisposalContext&)            = delete;
        DeferredDisposalContext& operator=(const DeferredDisposalContext&) = delete;
        DeferredDisposalContext(DeferredDisposalContext&&)                 = delete;
        DeferredDisposalContext& operator=(DeferredDisposalContext&&)      = delete;

        wil::unique_hmodule modulePin;
        wil::com_ptr<ViewerSpace> owner;
        std::shared_ptr<RetiredDisplayStorage> storage;
        std::shared_ptr<std::atomic_bool> done;
        uint64_t bytes = 0u;
    };

    HitGrid _hitGrid;
    std::unordered_map<uint32_t, D2D1_RECT_F> _lastRectsByNode;
    uint32_t _hoverNodeId = 0;

    std::wstring _tooltipText;
    uint32_t _tooltipNodeId              = 0;
    uint32_t _tooltipCandidateNodeId     = 0;
    double _tooltipCandidateSinceSeconds = 0.0;
    D2D1_POINT_2F _tooltipAnchorDip      = D2D1::Point2F(0.0f, 0.0f);
    bool _hasLastMouseMoveClientPoint    = false;
    POINT _lastMouseMoveClientPoint{};
    POINT _lastMouseMoveScreenPoint{};
#ifdef ENABLE_TESTS
    uint64_t _tooltipPaintCount              = 0u;
    D2D1_RECT_F _lastTooltipPaintedBoundsDip = {};
    bool _debugPublicationPhaseTiming        = false;
    bool _debugHasLastContextMenuScreenPoint = false;
    POINT _debugLastContextMenuScreenPoint{};
    uint32_t _debugLastContextMenuHitNodeId = 0;
    uint64_t _lastPaintUs                   = 0u;
    uint64_t _lastLayoutUs                  = 0u;
    uint64_t _lastDrainUs                   = 0u;
    uint64_t _lastWorkingSetBytes           = 0u;
    uint64_t _lastTileDrawCount             = 0u;
    uint64_t _lastTextDrawCount             = 0u;
    uint64_t _layoutGeneration              = 0u;
    uint64_t _presentedLayoutGeneration     = 0u;
    uint32_t _rendererDeviceCreateCount     = 0u;
    uint32_t _swapChainResizeCount          = 0u;
    uint32_t _rendererBrushCreateCount      = 0u;
    uint32_t _rendererTextFormatCreateCount = 0u;
    uint32_t _rendererFailureStage          = 0u;
    uint32_t _rendererFailureHr             = 0u;
    uint32_t _debugForcedRendererFault      = 0u;
    std::atomic<HANDLE> _debugValidationScratchEnteredEvent{nullptr};
    std::atomic<HANDLE> _debugValidationScratchReleaseEvent{nullptr};
    std::atomic_bool _debugPauseAfterValidationScratch{false};
    std::atomic<HANDLE> _debugWorkerStagingEnteredEvent{nullptr};
    std::atomic<HANDLE> _debugWorkerStagingReleaseEvent{nullptr};
    std::atomic_bool _debugPauseAfterWorkerStaging{false};
    mutable uint64_t _lastHitTestUs                = 0u;
    mutable uint32_t _lastHitTestCandidatesChecked = 0u;
    mutable uint32_t _lastHitTestPath              = 0u;
    uint32_t _lastLButtonDownHitNodeId             = 0u;
#endif

    std::atomic_uint64_t _pendingUpdatePostedCount{0};
    std::atomic_uint64_t _pendingUpdateCoalescedCount{0};
    std::atomic_uint64_t _pendingUpdateBytes{0};
    std::atomic_uint64_t _pendingUpdateAcceptedBytes{0};
    std::atomic_uint64_t _pendingUpdateLockHoldCount{0};
    std::atomic_uint64_t _pendingUpdateLockHoldTotalUs{0};
    std::atomic_uint64_t _pendingUpdateLockHoldMaxUs{0};
    std::atomic_uint64_t _pendingUpdateDrainedCount{0};
    std::atomic_uint64_t _pendingUpdateDrainCount{0};
    std::atomic_uint64_t _pendingUpdateDrainTotalUs{0};
    std::atomic_uint64_t _pendingUpdateDrainMaxUs{0};
    std::chrono::steady_clock::time_point _lastQueueDiagnosticsEmitAt{};
    bool _queueDiagnosticsTerminalEmitted = false;
    std::atomic_uint64_t _pendingUpdatePeakBytes{0};
    std::atomic_uint64_t _providerBufferCurrentBytes{0};
    std::atomic_uint64_t _providerBufferPeakBytes{0};
    std::atomic_uint64_t _validationScratchCurrentBytes{0};
    std::atomic_uint64_t _workerStagingCurrentBytes{0};
    std::atomic_uint64_t _workerStagingPeakBytes{0};
    std::atomic_uint64_t _workerStagingRejectedCount{0};
    std::atomic_uint64_t _scannerBookkeepingCurrentBytes{0};
    std::atomic_uint64_t _scannerBookkeepingPeakBytes{0};
    std::atomic_uint64_t _validationScratchPeakBytes{0};
    std::atomic_uint64_t _validationScratchRejectedCount{0};
    std::atomic_uint64_t _activePathCurrentBytes{0};
    std::atomic_uint64_t _activePathPeakBytes{0};
    std::atomic_uint64_t _activePathRejectedCount{0};
    std::atomic_uint64_t _deferredDisposalBytes{0};
    std::atomic_uint64_t _deferredDisposalMaxUs{0};
    std::atomic_uint64_t _deferredDisposalSubmitMaxUs{0};
    // Submission aggregates are UI-thread owned and emitted once at window teardown.
    uint64_t _deferredDisposalSubmitCount     = 0u;
    uint64_t _deferredDisposalSubmitTotalUs   = 0u;
    uint64_t _deferredDisposalSubmitBytes     = 0u;
    uint64_t _deferredDisposalSubmitOwnerPeak = 0u;
    std::atomic_uint64_t _refreshAdoptionMaxUs{0};
    std::vector<DeferredDisposalWorker> _deferredDisposalWorkers;
    std::chrono::steady_clock::time_point _openStartedAt{};
    bool _openToFirstPaintEmitted = false;
    std::chrono::steady_clock::time_point _viewerOpenStartedAt{};
    uint64_t _firstMeaningfulPaintUs = 0u;

    bool _trackingMouse       = false;
    bool _layoutDirty         = true;
    HeaderHit _hoverHeaderHit = HeaderHit::None;

    double _lastLayoutRebuildSeconds       = 0.0;
    bool _firstPositiveProjectionAttempted = false;
    bool _hasPublishedPositiveGeometry     = false;
    double _lastScanInvalidateSeconds      = 0.0;
    double _animationStartSeconds          = 0.0;
};
