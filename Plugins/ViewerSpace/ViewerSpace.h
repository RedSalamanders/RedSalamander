#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
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
#include <dwrite.h>
#include <dxgiformat.h>

#pragma comment(lib, "d2d1")
#pragma comment(lib, "dwrite")

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820) // WIL: deleted copy/move operators and padding
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"

class ViewerSpace final : public IViewer, public IInformations
{
public:
    ViewerSpace();
    ~ViewerSpace() = default;

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
    HRESULT STDMETHODCALLTYPE SetCallback(IViewerCallback* callback, void* cookie) noexcept override;

private:
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
        Up,
        Cancel,
    };

    struct MenuItemData final
    {
        std::wstring text;
        std::wstring shortcut;
        bool separator  = false;
        bool topLevel   = false;
        bool hasSubMenu = false;
    };

    struct Config final
    {
        uint32_t topFilesPerDirectory        = 96;
        uint32_t scanThreads                 = 1;
        uint32_t maxConcurrentScansPerVolume = 1;
        bool cacheEnabled                    = true;
        uint32_t cacheTtlSeconds             = 60;
        uint32_t cacheMaxEntries             = 1;
    };

    struct FileSummaryItem final
    {
        uint32_t nodeId = 0;
        std::wstring name;
        uint64_t bytes = 0;
    };

    struct Node final
    {
        uint32_t id         = 0;
        uint32_t parentId   = 0;
        bool isDirectory    = false;
        bool isSynthetic    = false;
        ScanState scanState = ScanState::NotStarted;

        std::wstring_view name;

        uint64_t totalBytes       = 0;
        uint32_t childrenStart    = 0;
        uint32_t childrenCount    = 0;
        uint32_t childrenCapacity = 0;
        uint32_t aggregateFolders = 0;
        uint32_t aggregateFiles   = 0;
    };

    struct DrawItem final
    {
        uint32_t nodeId      = 0;
        uint8_t depth        = 0;
        float labelHeightDip = 0.0f;
        D2D1_RECT_F targetRect{};
        D2D1_RECT_F currentRect{};
        D2D1_RECT_F startRect{};
        double animationStartSeconds = 0.0;
    };

    struct PendingUpdate final
    {
        enum class Kind : uint8_t
        {
            AddChild,
            UpdateSize,
            UpdateState,
            DirectoryFilesSummary,
            Progress,
        };

        Kind kind           = Kind::UpdateSize;
        uint32_t generation = 0;
        uint32_t nodeId     = 0;
        uint32_t parentId   = 0;
        uint64_t bytes      = 0;
        ScanState state     = ScanState::NotStarted;
        std::wstring name;
        bool isDirectory = false;
        bool isSynthetic = false;

        uint32_t scannedFolders = 0;
        uint32_t scannedFiles   = 0;

        uint64_t otherBytes  = 0;
        uint32_t otherCount  = 0;
        uint32_t otherNodeId = 0;
        std::vector<FileSummaryItem> topFiles;
    };

    static ATOM RegisterWndClass(HINSTANCE instance) noexcept;
    static constexpr wchar_t kClassName[] = L"RedSalamander.ViewerSpace";

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

    void OnCreate(HWND hwnd);
    void OnDestroy();
    void OnSize(UINT width, UINT height) noexcept;
    void OnPaint();
    void OnCommand(HWND hwnd, UINT commandId) noexcept;
    void OnKeyDown(WPARAM vk, bool alt) noexcept;
    void OnMouseMove(int x, int y) noexcept;
    void OnMouseLeave() noexcept;
    void OnLButtonDown(int x, int y) noexcept;
    void OnLButtonDblClk(int x, int y) noexcept;
    void OnContextMenu(HWND hwnd, POINT screenPt) noexcept;
    LRESULT OnMeasureItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept;
    LRESULT OnDrawItem(DRAWITEMSTRUCT* draw) noexcept;
    LRESULT OnNotify(WPARAM wp, LPARAM lp) noexcept;
    void OnNcActivate(HWND hwnd, bool windowActive) noexcept;
    LRESULT OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept;
    void OnTimer(UINT_PTR timerId) noexcept;

    bool EnsureDirect2D(HWND hwnd) noexcept;
    void DiscardDirect2D() noexcept;
    void ApplyThemeToWindow(HWND hwnd) noexcept;
    void ApplyTitleBarTheme(HWND hwnd, bool windowActive) noexcept;
    void UpdateWindowTitle(HWND hwnd) noexcept;
    void ApplyMenuTheme(HWND hwnd) noexcept;
    void PrepareMenuTheme(HMENU menu, bool topLevel, std::vector<MenuItemData>& outItems) noexcept;
    void OnMeasureMenuItem(HWND hwnd, MEASUREITEMSTRUCT* measure) noexcept;
    void OnDrawMenuItem(DRAWITEMSTRUCT* draw) noexcept;
    void EnsureTooltip(HWND hwnd) noexcept;
    void ApplyThemeToTooltip() noexcept;
    void UpdateTooltipForHit(uint32_t nodeId) noexcept;
    void UpdateTooltipPosition(int x, int y) noexcept;
    std::wstring BuildTooltipText(uint32_t nodeId) const;

    void UpdateHeaderTextCache() noexcept;

    void StartScan(std::wstring_view rootPath, bool allowCache = true);
    void CancelScan() noexcept;
    void CancelScanByUser() noexcept;
    void CancelScanAndWait() noexcept;
    void ReapFinishedScanWorkers(bool wait) noexcept;
    void ScanMain(std::stop_token stopToken,
                  uint32_t generation,
                  wil::com_ptr<IFileSystem> fileSystem,
                  bool fileSystemIsWin32,
                  std::wstring rootPath,
                  uint32_t rootNodeId,
                  uint32_t nextNodeId,
                  size_t topFilesPerDirectory,
                  uint32_t scanThreads) noexcept;

    void PostUpdate(PendingUpdate&& update) noexcept;
    void DrainUpdates() noexcept;
    void ContinueScanCacheBuild() noexcept;
    void CancelScanCacheBuild() noexcept;
    const Node* TryGetRealNode(uint32_t nodeId) const noexcept;
    Node* TryGetRealNode(uint32_t nodeId) noexcept;
    std::span<const uint32_t> GetRealNodeChildren(const Node& node) const noexcept;
    void AddRealNodeChild(Node& parent, uint32_t childNodeId) noexcept;
    std::wstring BuildNodePathText(uint32_t nodeId) const;
    void UpdateViewPathText() noexcept;

    void EnsureLayoutForView() noexcept;
    void MaybeRebuildLayout() noexcept;
    void RebuildLayout() noexcept;
    std::optional<uint32_t> HitTestTreemap(float xDip, float yDip) const noexcept;

    void NavigateTo(uint32_t nodeId) noexcept;
    void NavigateUp() noexcept;
    void RefreshCurrent() noexcept;
    bool CanNavigateUp() const noexcept;
    void UpdateMenuState(HWND hwnd) noexcept;

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

    IViewerCallback* _callback = nullptr;
    void* _callbackCookie      = nullptr;

    wil::com_ptr<IHostPaneExecute> _hostPaneExecute;

    wil::com_ptr<IFileSystem> _fileSystem;
    std::wstring _fileSystemName;
    std::wstring _fileSystemShortId;
    bool _fileSystemIsWin32 = true;

    ViewerTheme _theme{};
    bool _hasTheme        = false;
    bool _allowEraseBkgnd = true;

    wil::unique_hwnd _hWnd;
    wil::unique_any<HBRUSH, decltype(&::DeleteObject), ::DeleteObject> _menuBackgroundBrush;
    wil::unique_any<HFONT, decltype(&::DeleteObject), ::DeleteObject> _menuFont;
    std::vector<MenuItemData> _menuThemeItems;

    float _dpi = static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    SIZE _clientSize{};

    wil::com_ptr<ID2D1Factory> _d2dFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> _renderTarget;
    wil::com_ptr<ID2D1SolidColorBrush> _brushBackground;
    wil::com_ptr<ID2D1SolidColorBrush> _brushText;
    wil::com_ptr<ID2D1SolidColorBrush> _brushOutline;
    wil::com_ptr<ID2D1SolidColorBrush> _brushAccent;
    wil::com_ptr<ID2D1SolidColorBrush> _brushWatermark;
    wil::com_ptr<ID2D1LinearGradientBrush> _brushShading;
    wil::com_ptr<ID2D1GradientStopCollection> _shadingStops;
    wil::com_ptr<ID2D1StrokeStyle> _otherStrokeStyle;
    wil::com_ptr<ID2D1PathGeometry> _dogEarFlapGeometry;

    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    wil::com_ptr<IDWriteTextFormat> _textFormat;
    wil::com_ptr<IDWriteTextFormat> _headerFormat;
    wil::com_ptr<IDWriteTextFormat> _headerStatusFormatRight;
    wil::com_ptr<IDWriteTextFormat> _headerInfoFormat;
    wil::com_ptr<IDWriteTextFormat> _headerInfoFormatRight;
    wil::com_ptr<IDWriteTextFormat> _headerIconFormat;
    wil::com_ptr<IDWriteTextFormat> _watermarkFormat;

    struct ScanWorker final
    {
        ScanWorker()                        = default;
        ~ScanWorker()                       = default;
        ScanWorker(ScanWorker&&)            = default;
        ScanWorker& operator=(ScanWorker&&) = default;

        ScanWorker(const ScanWorker&)            = delete;
        ScanWorker& operator=(const ScanWorker&) = delete;

        std::jthread thread;
        std::shared_ptr<std::atomic_bool> done;
    };

    std::pmr::unsynchronized_pool_resource _nodePool;
    std::pmr::monotonic_buffer_resource _nameArena;
    std::pmr::monotonic_buffer_resource _layoutNameArena;
    std::pmr::vector<Node> _nodes             = std::pmr::vector<Node>(&_nodePool);
    std::pmr::vector<uint32_t> _childrenArena = std::pmr::vector<uint32_t>(&_nodePool);

    ScanWorker _scanWorker;
    std::vector<ScanWorker> _retiredScanWorkers;
    std::atomic_uint32_t _scanGeneration{0};
    std::atomic_bool _scanActive{false};

    std::mutex _updateMutex;
    std::deque<PendingUpdate> _pendingUpdates;

    std::shared_ptr<void> _scanCacheBuildSnapshot;
    std::wstring _scanCacheBuildRootKey;
    uint32_t _scanCacheBuildTopFilesPerDirectory = 0;
    uint32_t _scanCacheBuildGeneration           = 0;
    uint32_t _scanCacheLastStoredGeneration      = 0;
    size_t _scanCacheBuildChildrenNext           = 0;
    size_t _scanCacheBuildNodesNext              = 0;

    std::unordered_map<uint32_t, Node> _syntheticNodes;
    std::unordered_map<uint32_t, uint32_t> _otherBucketIdsByParent;
    std::unordered_map<uint32_t, uint32_t> _layoutMaxItemsByNode;
    std::unordered_set<uint32_t> _autoExpandedOtherByNode;
    uint32_t _nextSyntheticNodeId = 0x80000000u;
    uint32_t _rootNodeId          = 0;
    uint32_t _viewNodeId          = 0;
    std::wstring _scanRootPath;
    std::optional<std::wstring> _scanRootParentPath;
    std::wstring _viewPathText;
    std::vector<uint32_t> _navStack;

    ScanState _overallState = ScanState::NotStarted;
    double _scanCompletedSinceSeconds = 0.0;

    uint64_t _scanProgressBytes    = 0;
    uint32_t _scanProgressFolders  = 0;
    uint32_t _scanProgressFiles    = 0;
    uint32_t _scanProcessingNodeId = 0;
    std::wstring _scanProcessingFolderName;
    UINT _headerStatusId = 0;
    std::wstring _headerStatusText;
    std::wstring _headerCountsText;
    std::wstring _headerSizeText;
    std::wstring _headerProcessingText;
    std::wstring _scanInProgressWatermarkText;
    std::wstring _scanIncompleteWatermarkText;
    std::wstring _headerPathSourceText;
    std::wstring _headerPathDisplayText;
    float _headerPathDisplayMaxWidthDip = 0.0f;

    std::vector<DrawItem> _drawItems;
    uint32_t _hoverNodeId = 0;

    wil::unique_hwnd _hTooltip;
    std::wstring _tooltipText;
    uint32_t _tooltipNodeId              = 0;
    uint32_t _tooltipCandidateNodeId     = 0;
    double _tooltipCandidateSinceSeconds = 0.0;

    bool _trackingMouse       = false;
    bool _layoutDirty         = true;
    HeaderHit _hoverHeaderHit = HeaderHit::None;

    double _lastLayoutRebuildSeconds  = 0.0;
    double _lastScanInvalidateSeconds = 0.0;
    double _animationStartSeconds     = 0.0;
};
