#include "ViewerSpace.h"
#include "LocalizationManager.h"
#include "ViewerSpace.SlicedSort.h"
#include "resource.h"

#include <algorithm>
#include <cmath>
#include <exception>
#include <limits>

extern HINSTANCE g_hInstance;

namespace
{
constexpr float kPadding           = 8.0f;
constexpr float kGap               = 1.0f;
constexpr double kAnimationSeconds = 0.18;
float Area(const D2D1_RECT_F& rect) noexcept
{
    return std::max(0.0f, rect.right - rect.left) * std::max(0.0f, rect.bottom - rect.top);
}
double Worst(double sum, double smallest, double largest, double side, double scale) noexcept
{
    if (sum <= 0.0 || smallest <= 0.0 || side <= 0.0)
    {
        return std::numeric_limits<double>::infinity();
    }
    const double area = sum * scale;
    return std::max(side * side * largest * scale / (area * area), area * area / (side * side * smallest * scale));
}
} // namespace

struct ViewerSpace::LayoutBuildState final
{
    enum class Phase : uint8_t
    {
        Capture,
        Next,
        Score,
        Sort,
        Row,
        Place,
        Finalize,
        Grid,
        Publish
    };
    struct Task final
    {
        uint32_t id = 0u;
        D2D1_RECT_F bounds{};
        uint8_t depth     = 0u;
        size_t ownerIndex = std::numeric_limits<size_t>::max();
    };
    uint32_t generation = 0u;
    uint64_t revision   = 0u;
    uint32_t view       = 0u;
    float dpi           = 96.0f;
    D2D1_RECT_F bounds{};
    bool scanning        = false;
    Phase phase          = Phase::Capture;
    size_t cursor        = 0u;
    size_t maxDraw       = 0u;
    size_t maxCandidates = 0u;
    std::unordered_map<uint32_t, D2D1_RECT_F> previous;
    std::vector<DrawItem> draw;
    std::unordered_map<uint32_t, Node> synthetic;
    // deque keeps string storage stable when a later synthetic name is appended.
    std::deque<std::wstring> names;
    std::vector<Task> tasks;
    Task task{};
    std::vector<LayoutItem> items;
    ViewerSpaceDetail::SlicedSort<LayoutItem> sort;
    uint64_t otherBytes   = 0u;
    uint64_t otherFolders = 0u;
    uint64_t otherFiles   = 0u;
    double otherWeight    = 0.0;
    double totalWeight    = 0.0;
    D2D1_RECT_F freeRect{};
    double scale         = 0.0;
    size_t rowStart      = 0u;
    size_t rowEnd        = 0u;
    size_t placing       = 0u;
    double rowWeight     = 0.0;
    double rowMin        = 0.0;
    double rowMax        = 0.0;
    bool horizontal      = false;
    float thickness      = 0.0f;
    float coordinate     = 0.0f;
    double animationTime = 0.0;
    HitGrid grid;
    bool gridItemActive           = false;
    uint32_t gridLeft             = 0u;
    uint32_t gridRight            = 0u;
    uint32_t gridTop              = 0u;
    uint32_t gridBottom           = 0u;
    uint32_t gridX                = 0u;
    uint32_t gridY                = 0u;
    uint64_t nestedDirectories    = 0u;
    uint64_t nestedFiles          = 0u;
    uint64_t expandableUnexpanded = 0u;
    bool hasPositiveGeometry      = false;
};

void ViewerSpace::RetireLayoutStorage(std::shared_ptr<LayoutBuildState> layoutBuild,
                                      std::shared_ptr<LayoutBuildState> publishedLayoutOwner,
                                      bool cancellation) noexcept
{
    auto storage                = std::make_unique<RetiredDisplayStorage>();
    storage->layoutOnly         = true;
    storage->layoutCancellation = cancellation;
    if (cancellation && layoutBuild != nullptr)
    {
        storage->layoutGeometryCount = layoutBuild->draw.size();
        storage->layoutGeneration    = layoutBuild->generation;
    }
    uint64_t bytes = layoutBuild == nullptr ? 0u : LayoutBuildOwnedBytes(*layoutBuild);
    if (publishedLayoutOwner != nullptr)
    {
        const uint64_t ownerBytes = LayoutBuildOwnedBytes(*publishedLayoutOwner);
        bytes                     = ownerBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + ownerBytes;
    }
    storage->layoutBuild          = std::move(layoutBuild);
    storage->publishedLayoutOwner = std::move(publishedLayoutOwner);
    QueueDeferredDisposal(std::move(storage), bytes, false);
}

void ViewerSpace::CancelLayoutBuild() noexcept
{
    if (! _layoutBuild)
    {
        return;
    }
    if (_cancelledLayoutBuild)
    {
        Debug::Error(L"ViewerSpace: a second layout candidate reached cancellation before the first retired.");
        std::terminate();
    }

    _layoutDirty          = true;
    _cancelledLayoutBuild = std::move(_layoutBuild);
    ContinueCancelledLayoutBuildRetirement();
}

void ViewerSpace::ContinueCancelledLayoutBuildRetirement() noexcept
{
    if (! _cancelledLayoutBuild || ! HasDeferredDisposalCapacity())
    {
        return;
    }

    auto cancelled = std::move(_cancelledLayoutBuild);
    RetireLayoutStorage(std::move(cancelled), {}, true);
}

uint64_t ViewerSpace::LayoutBuildOwnedBytes(const LayoutBuildState& build) const noexcept
{
    uint64_t bytes = build.draw.capacity() * sizeof(DrawItem) + build.items.capacity() * sizeof(LayoutItem) +
                     build.sort.scratch.capacity() * sizeof(LayoutItem) + build.tasks.capacity() * sizeof(LayoutBuildState::Task) +
                     build.grid.cells.capacity() * sizeof(std::vector<uint32_t>);
    for (const auto& cell : build.grid.cells)
    {
        bytes += cell.capacity() * sizeof(uint32_t);
    }
    // Node-based containers have implementation overhead; reserve a conservative node allowance.
    bytes += build.previous.bucket_count() * sizeof(void*) + build.previous.size() * (sizeof(uint32_t) + sizeof(D2D1_RECT_F) + 3u * sizeof(void*));
    bytes += build.synthetic.bucket_count() * sizeof(void*) + build.synthetic.size() * (sizeof(uint32_t) + sizeof(Node) + 3u * sizeof(void*));
    bytes += build.names.size() * sizeof(std::wstring);
    for (const auto& name : build.names)
    {
        bytes += name.capacity() * sizeof(wchar_t);
    }
    return bytes;
}

uint64_t ViewerSpace::LayoutOwnedBytes() const noexcept
{
    auto gridBytes = [](const HitGrid& grid) noexcept
    {
        uint64_t bytes = grid.cells.capacity() * sizeof(std::vector<uint32_t>);
        for (const auto& cell : grid.cells)
        {
            bytes += cell.capacity() * sizeof(uint32_t);
        }
        return bytes;
    };
    uint64_t bytes = _drawItems.capacity() * sizeof(DrawItem) + gridBytes(_hitGrid);
    // Charge retained node-based layout indexes conservatively (bucket array + node links/value storage).
    bytes += _syntheticNodes.bucket_count() * sizeof(void*) + _syntheticNodes.size() * (sizeof(uint32_t) + sizeof(Node) + 3u * sizeof(void*));
    bytes += _layoutMaxItemsByNode.bucket_count() * sizeof(void*) + _layoutMaxItemsByNode.size() * (2u * sizeof(uint32_t) + 3u * sizeof(void*));
    bytes += _autoExpandedOtherByNode.bucket_count() * sizeof(void*) + _autoExpandedOtherByNode.size() * (sizeof(uint32_t) + 3u * sizeof(void*));
    bytes += _lastRectsByNode.bucket_count() * sizeof(void*) + _lastRectsByNode.size() * (sizeof(uint32_t) + sizeof(D2D1_RECT_F) + 3u * sizeof(void*));
    bytes += _layoutCandidateCache.bucket_count() * sizeof(void*) +
             _layoutCandidateCache.size() * (sizeof(uint32_t) + sizeof(LayoutCandidateCacheEntry) + 3u * sizeof(void*));
    for (const auto& [key, entry] : _layoutCandidateCache)
    {
        static_cast<void>(key);
        bytes += entry.items.capacity() * sizeof(LayoutItem);
    }
    bytes += _layoutWorkspaceByDepth.capacity() * sizeof(LayoutWorkspace);
    for (const LayoutWorkspace& workspace : _layoutWorkspaceByDepth)
    {
        bytes += workspace.items.capacity() * sizeof(LayoutItem) + workspace.topItems.capacity() * sizeof(LayoutItem) +
                 workspace.forcedItems.capacity() * sizeof(LayoutItem) + workspace.row.capacity() * sizeof(LayoutItem) +
                 workspace.forcedChildIds.capacity() * sizeof(uint32_t) + workspace.expandTasks.capacity() * sizeof(LayoutExpandTask);
    }
    if (_layoutBuild)
    {
        bytes += LayoutBuildOwnedBytes(*_layoutBuild);
    }
    if (_cancelledLayoutBuild)
    {
        bytes += LayoutBuildOwnedBytes(*_cancelledLayoutBuild);
    }
    if (_publishedLayoutOwner)
    {
        bytes += LayoutBuildOwnedBytes(*_publishedLayoutOwner);
    }
    return bytes;
}

uint64_t ViewerSpace::ViewerOwnedBytes() const noexcept
{
    uint64_t bytes = _displayModel == nullptr ? 0u : _displayModel->AllocatedBytes();
    if (_workingModel != nullptr)
    {
        const uint64_t candidateBytes = _workingModel->AllocatedBytes();
        bytes = candidateBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + candidateBytes;
    }
    const uint64_t pendingBytes  = _pendingUpdateBytes.load(std::memory_order_relaxed);
    bytes                        = pendingBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + pendingBytes;
    const uint64_t livePathBytes = _activePathCurrentBytes.load(std::memory_order_relaxed);
    bytes                        = livePathBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + livePathBytes;
    const uint64_t liveValidationBytes = _validationScratchCurrentBytes.load(std::memory_order_relaxed);
    bytes = liveValidationBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + liveValidationBytes;
    const uint64_t liveStagingBytes = _workerStagingCurrentBytes.load(std::memory_order_relaxed);
    bytes = liveStagingBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + liveStagingBytes;
    const uint64_t bookkeepingBytes = _scannerBookkeepingCurrentBytes.load(std::memory_order_relaxed);
    bytes = bookkeepingBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + bookkeepingBytes;
    const uint64_t cacheValidationBytes = static_cast<uint64_t>(_scanCacheRestoreSeenChildren.capacity()) * sizeof(uint64_t);
    bytes = cacheValidationBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + cacheValidationBytes;
    if (_scanCacheRestoreContinuationPending)
    {
        bytes = ViewerSpaceScan::SaturatingAdd(bytes, sizeof(ScanCacheRestoreContinuation));
    }
    const uint64_t layoutBytes   = LayoutOwnedBytes();
    bytes                        = layoutBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + layoutBytes;
    const uint64_t analysisBytes = AnalysisOwnedBytes();
    bytes                        = analysisBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + analysisBytes;
    const uint64_t deferredDisposalBytes = _deferredDisposalBytes.load(std::memory_order_relaxed);
    bytes = deferredDisposalBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + deferredDisposalBytes;
    if (_deferredStartScan.has_value())
    {
        const uint64_t deferredPathBytes = (static_cast<uint64_t>(_deferredStartScan.value().rootPath.capacity()) + 1u) * sizeof(wchar_t);
        bytes = deferredPathBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + deferredPathBytes;
    }
    uint64_t refreshBytes = static_cast<uint64_t>(_refreshPathCandidates.capacity() + _refreshNextPathCandidates.capacity()) * sizeof(uint32_t);
    if (_refreshDisplayState.has_value())
    {
        const auto& state = _refreshDisplayState.value();
        refreshBytes      = ViewerSpaceScan::SaturatingAdd(
            refreshBytes, static_cast<uint64_t>(state.processingFolderName.capacity() + state.viewPathText.capacity()) * sizeof(wchar_t));
        refreshBytes = ViewerSpaceScan::SaturatingAdd(refreshBytes,
                                                      static_cast<uint64_t>(state.navStack.capacity() + state.navForwardStack.capacity()) * sizeof(uint32_t));
        refreshBytes = ViewerSpaceScan::SaturatingAdd(refreshBytes, static_cast<uint64_t>(state.selectedRelativePath.capacity()) * sizeof(std::wstring));
        for (const auto& part : state.selectedRelativePath)
        {
            refreshBytes = ViewerSpaceScan::SaturatingAdd(refreshBytes, static_cast<uint64_t>(part.capacity()) * sizeof(wchar_t));
        }
    }
    bytes = ViewerSpaceScan::SaturatingAdd(bytes, refreshBytes);
    bytes = ViewerSpaceScan::SaturatingAdd(bytes, WorkspaceOwnedBytes());
    return bytes;
}

D2D1_RECT_F ViewerSpace::GetTreemapBounds() const noexcept
{
    if (_treemapBounds.right > _treemapBounds.left && _treemapBounds.bottom > _treemapBounds.top)
    {
        return _treemapBounds;
    }
    const float top = GetHeaderBottomDip() + kPadding;
    return D2D1::RectF(kPadding, top, std::max(kPadding, DipFromPx(_clientSize.cx) - kPadding), std::max(top, DipFromPx(_clientSize.cy) - kPadding));
}

void ViewerSpace::MaybeRebuildLayout() noexcept
{
    ContinueCancelledLayoutBuildRetirement();
    if (_cancelledLayoutBuild)
    {
        return;
    }

    // A cache restore is an uncommitted generation. Its child/name/file arrays are populated in separate slices.
    if (_scanCacheRestoreSnapshot)
    {
        CancelLayoutBuild();
        return;
    }
    if (! _layoutDirty && ! _layoutBuild)
    {
        return;
    }
    const double now = NowSeconds();
    const Node* view = TryGetRealNode(_model->viewNodeId);
    const bool viewHasCommittedPositiveBytes =
        view != nullptr && (view->observedOwnFileBytes != 0u || view->observedChildBytes != 0u || view->aggregateBytes != 0u);
    const bool firstPositiveProjection = ! _hasPublishedPositiveGeometry && viewHasCommittedPositiveBytes && ! _firstPositiveProjectionAttempted;
    const bool scanCadenceThrottled    = _scanActive.load() && _lastLayoutRebuildSeconds > 0.0 && now - _lastLayoutRebuildSeconds < 0.25;
    if (! _layoutBuild && (! CanBeginScanProjection(! _drawItems.empty() && ! firstPositiveProjection) || (scanCadenceThrottled && ! firstPositiveProjection)))
    {
        return;
    }
    if (! _layoutBuild && _scanActive.load() && firstPositiveProjection)
    {
        // Count this first positive-preview attempt against cadence if a later model revision cancels it.
        _firstPositiveProjectionAttempted = true;
        _lastLayoutRebuildSeconds         = now;
    }
    RebuildLayout();
}

void ViewerSpace::RebuildLayout() noexcept
{
    if (_cancelledLayoutBuild)
    {
        return;
    }

    const auto started        = std::chrono::steady_clock::now();
    const D2D1_RECT_F bounds  = GetTreemapBounds();
    const uint32_t generation = _displayGeneration;
    if (_layoutBuild && (_layoutBuild->generation != generation || _layoutBuild->revision != _modelRevision || _layoutBuild->view != _model->viewNodeId ||
                         _layoutBuild->dpi != _dpi || _layoutBuild->bounds.left != bounds.left || _layoutBuild->bounds.top != bounds.top ||
                         _layoutBuild->bounds.right != bounds.right || _layoutBuild->bounds.bottom != bounds.bottom))
    {
        CancelLayoutBuild();
        if (_cancelledLayoutBuild)
        {
            return;
        }
    }
    if (! _layoutBuild)
    {
        _layoutBuild     = std::make_shared<LayoutBuildState>();
        auto& build      = *_layoutBuild;
        build.generation = generation;
        build.revision   = _modelRevision;
        build.view       = _model->viewNodeId;
        build.dpi        = _dpi;
        build.bounds     = bounds;
        build.scanning   = _scanActive.load();
        build.maxDraw    = build.scanning ? 20000u : 50000u;
        build.tasks.push_back({build.view, bounds, 0u});
    }
    auto& build     = *_layoutBuild;
    auto heapLess   = [](const LayoutItem& a, const LayoutItem& b) noexcept { return a.weight > b.weight || (a.weight == b.weight && a.nodeId < b.nodeId); };
    auto sortedLess = [](const LayoutItem& a, const LayoutItem& b) noexcept { return a.bytes > b.bytes || (a.bytes == b.bytes && a.nodeId < b.nodeId); };
    auto taskLess   = [](const LayoutBuildState::Task& a, const LayoutBuildState::Task& b) noexcept
    { return Area(a.bounds) < Area(b.bounds) || (Area(a.bounds) == Area(b.bounds) && a.id > b.id); };
    auto addOther = [&](const LayoutItem& item) noexcept
    {
        build.otherBytes = ViewerSpaceScan::SaturatingAdd(build.otherBytes, item.bytes);
        build.otherWeight += item.weight;
        const auto resolved = ResolveTreemapItem(item.nodeId);
        if (resolved.has_value())
        {
            const auto& value  = resolved.value();
            build.otherFolders = ViewerSpaceScan::SaturatingAdd(build.otherFolders, value.isSynthetic ? value.aggregateFolders : value.isDirectory ? 1u : 0u);
            build.otherFiles   = ViewerSpaceScan::SaturatingAdd(build.otherFiles, value.isSynthetic ? value.aggregateFiles : value.isDirectory ? 0u : 1u);
        }
    };
    auto beginPlacement = [&]() noexcept
    {
        build.horizontal = (build.freeRect.bottom - build.freeRect.top) > (build.freeRect.right - build.freeRect.left);
        const float side = build.horizontal ? build.freeRect.right - build.freeRect.left : build.freeRect.bottom - build.freeRect.top;
        build.thickness  = static_cast<float>(build.rowWeight * build.scale / std::max(0.001f, side));
        build.coordinate = build.horizontal ? build.freeRect.left : build.freeRect.top;
        build.placing    = build.rowStart;
        build.phase      = LayoutBuildState::Phase::Place;
    };

    size_t work                  = 0u;
    bool published               = false;
    bool disposalCapacityBlocked = false;
    while (work++ < _layoutWorkUnitsPerTurn)
    {
        if ((work & 63u) == 0u && Debug::Perf::ElapsedUs(started) >= 1800u)
        {
            break;
        }
        switch (build.phase)
        {
            case LayoutBuildState::Phase::Capture:
            {
                if (build.cursor < _drawItems.size())
                {
                    const auto& item = _drawItems[build.cursor++];
                    build.previous.emplace(item.nodeId, item.currentRect);
                    break;
                }
                build.phase = LayoutBuildState::Phase::Next;
                break;
            }
            case LayoutBuildState::Phase::Next:
            {
                if (build.tasks.empty() || build.draw.size() >= build.maxDraw)
                {
                    build.cursor        = 0u;
                    build.animationTime = NowSeconds();
                    build.phase         = LayoutBuildState::Phase::Finalize;
                    break;
                }
                std::pop_heap(build.tasks.begin(), build.tasks.end(), taskLess);
                build.task = build.tasks.back();
                build.tasks.pop_back();
                const Node* parent = TryGetRealNode(build.task.id);
                if (! parent || Area(build.task.bounds) <= 1.0f)
                {
                    break;
                }
                build.items.clear();
                build.sort         = {};
                build.cursor       = 0u;
                build.otherBytes   = parent->aggregateBytes;
                build.otherFolders = parent->aggregateFolders;
                build.otherFiles   = parent->aggregateFiles;
                build.otherWeight =
                    _workspaceViewExcluded
                        ? 0.0
                        : static_cast<double>(_config.fileCountWeight
                                                  ? (_workspaceFactsReady && _workspaceFactsGeneration == _displayGeneration ? parent->aggregateFiles : 0u)
                                                  : parent->aggregateBytes);
                const size_t pixelLimit  = static_cast<size_t>(std::clamp(std::ceil(static_cast<double>(Area(build.task.bounds)) / 64.0), 4096.0, 50000.0));
                const auto overrideLimit = _layoutMaxItemsByNode.find(parent->id);
                const size_t requested =
                    overrideLimit == _layoutMaxItemsByNode.end() ? pixelLimit : std::max(pixelLimit, static_cast<size_t>(overrideLimit->second));
                // Exact-fit real children need no synthetic slot. Reserve it after scoring only
                // when a positive remainder actually exists; zero-byte rows have no map area.
                build.maxCandidates = std::min(requested, build.maxDraw - build.draw.size());
                build.items.reserve(build.maxCandidates + 1u);
                build.phase = LayoutBuildState::Phase::Score;
                break;
            }
            case LayoutBuildState::Phase::Score:
            {
                const Node* parent  = TryGetRealNode(build.task.id);
                const auto children = parent ? GetRealNodeChildren(*parent) : std::span<const uint32_t>{};
                if (build.cursor < children.size())
                {
                    const uint32_t id   = children[build.cursor++];
                    const auto resolved = ResolveTreemapItem(id);
                    if (! resolved.has_value())
                    {
                        break;
                    }
                    const uint64_t bytes = ShownBytes(id, resolved.value().totalBytes);
                    // A zero-byte item stays available in the linked analysis table, but has no
                    // proportional treemap area and must not acquire synthetic weight.
                    const uint64_t weight = MapWeight(id, resolved.value().totalBytes);
                    if (weight == 0u)
                    {
                        break;
                    }
                    LayoutItem item{id, static_cast<double>(weight), bytes};
                    if (build.items.size() < build.maxCandidates)
                    {
                        build.items.push_back(item);
                        std::push_heap(build.items.begin(), build.items.end(), heapLess);
                    }
                    else if (! build.items.empty() && sortedLess(item, build.items.front()))
                    {
                        std::pop_heap(build.items.begin(), build.items.end(), heapLess);
                        addOther(build.items.back());
                        build.items.back() = item;
                        std::push_heap(build.items.begin(), build.items.end(), heapLess);
                    }
                    else
                    {
                        addOther(item);
                    }
                    break;
                }
                build.phase = LayoutBuildState::Phase::Sort;
                break;
            }
            case LayoutBuildState::Phase::Sort:
            {
                if (build.otherWeight > 0.0 && ! build.items.empty() && build.items.size() == build.maxDraw - build.draw.size())
                {
                    // Scoring left a min-heap. Fold its smallest candidate into the exact remainder
                    // once, before sliced sorting, to make room for the eventual Other tile.
                    std::pop_heap(build.items.begin(), build.items.end(), heapLess);
                    addOther(build.items.back());
                    build.items.pop_back();
                }
                if (! build.sort.Step(build.items, sortedLess))
                {
                    break;
                }
                build.totalWeight = build.otherWeight;
                // Accumulation is its own resumable pass, also allowing a synthetic final item.
                build.cursor   = 0u;
                build.phase    = LayoutBuildState::Phase::Row;
                build.rowStart = std::numeric_limits<size_t>::max();
                break;
            }
            case LayoutBuildState::Phase::Row:
            {
                if (build.rowStart == std::numeric_limits<size_t>::max())
                {
                    if (build.cursor < build.items.size())
                    {
                        build.totalWeight += build.items[build.cursor++].weight;
                        break;
                    }
                    if (build.otherWeight > 0.0)
                    {
                        const uint32_t id = SyntheticOtherBucketIdForParent(build.task.id);
                        build.names.push_back(FormatStringResource(
                            g_hInstance, IDS_VIEWERSPACE_OTHER_BUCKET_FORMAT, ViewerSpaceScan::SaturatingAdd(build.otherFolders, build.otherFiles)));
                        Node other;
                        other.id                 = id;
                        other.parentId           = build.task.id;
                        other.isSynthetic        = true;
                        const Node* parent       = TryGetRealNode(build.task.id);
                        other.contentNotMeasured = parent != nullptr && parent->contentNotMeasured;
                        other.scanState          = ScanState::Done;
                        other.name               = build.names.back();
                        other.totalBytes         = build.otherBytes;
                        other.aggregateBytes     = build.otherBytes;
                        other.aggregateFolders   = build.otherFolders;
                        other.aggregateFiles     = build.otherFiles;
                        build.synthetic.emplace(id, other);
                        build.items.push_back({id, build.otherWeight, build.otherBytes});
                    }
                    build.freeRect  = build.task.bounds;
                    build.scale     = build.totalWeight > 0.0 ? static_cast<double>(Area(build.task.bounds)) / build.totalWeight : 0.0;
                    build.rowStart  = 0u;
                    build.rowEnd    = 0u;
                    build.rowWeight = 0.0;
                }
                if (build.rowEnd == build.items.size())
                {
                    if (build.rowEnd > build.rowStart)
                    {
                        beginPlacement();
                    }
                    else
                    {
                        build.phase = LayoutBuildState::Phase::Next;
                    }
                    break;
                }
                const auto& next  = build.items[build.rowEnd];
                const double side = std::min(build.freeRect.right - build.freeRect.left, build.freeRect.bottom - build.freeRect.top);
                if (build.rowStart == build.rowEnd ||
                    Worst(build.rowWeight + next.weight, std::min(build.rowMin, next.weight), std::max(build.rowMax, next.weight), side, build.scale) <=
                        Worst(build.rowWeight, build.rowMin, build.rowMax, side, build.scale))
                {
                    if (build.rowStart == build.rowEnd)
                    {
                        build.rowMin = next.weight;
                        build.rowMax = next.weight;
                    }
                    else
                    {
                        build.rowMin = std::min(build.rowMin, next.weight);
                        build.rowMax = std::max(build.rowMax, next.weight);
                    }
                    build.rowWeight += next.weight;
                    ++build.rowEnd;
                }
                else
                {
                    beginPlacement();
                }
                break;
            }
            case LayoutBuildState::Phase::Place:
            {
                if (build.placing == build.rowEnd)
                {
                    if (build.horizontal)
                    {
                        build.freeRect.top += build.thickness;
                    }
                    else
                    {
                        build.freeRect.left += build.thickness;
                    }
                    build.rowStart  = build.rowEnd;
                    build.rowWeight = 0.0;
                    build.phase     = LayoutBuildState::Phase::Row;
                    break;
                }
                const auto& item   = build.items[build.placing++];
                const float length = static_cast<float>(item.weight * build.scale / std::max(0.001f, build.thickness));
                D2D1_RECT_F rect   = build.horizontal
                                         ? D2D1::RectF(build.coordinate, build.freeRect.top, build.coordinate + length, build.freeRect.top + build.thickness)
                                         : D2D1::RectF(build.freeRect.left, build.coordinate, build.freeRect.left + build.thickness, build.coordinate + length);
                build.coordinate += length;
                // Exact last edge avoids accumulated float seams while preserving every child's coverage.
                if (build.placing == build.rowEnd)
                {
                    if (build.horizontal)
                    {
                        rect.right = build.freeRect.right;
                    }
                    else
                    {
                        rect.bottom = build.freeRect.bottom;
                    }
                }
                DrawItem draw;
                draw.nodeId            = item.nodeId;
                draw.depth             = build.task.depth;
                draw.targetRect        = rect;
                draw.lod               = ClassifyDrawItemLod(rect, draw.depth);
                draw.areaDip2          = Area(rect);
                const Node* node       = TryGetRealNode(item.nodeId);
                const bool hasChildren = node && (! GetRealNodeChildren(*node).empty() || node->aggregateFolders || node->aggregateFiles);
                const float label      = node && node->scanState != ScanState::Done ? 34.0f : 22.0f;
                D2D1_RECT_F child      = D2D1::RectF(rect.left + 2.0f, rect.top + label + 2.0f, rect.right - 2.0f, rect.bottom - 2.0f);
                const bool fits        = child.right - child.left >= 16.0f && child.bottom - child.top >= 16.0f && Area(child) >= 512.0f;
                if (hasChildren && fits && draw.depth < (build.scanning ? 8u : 10u) && build.draw.size() + 1u < build.maxDraw)
                {
                    draw.labelHeightDip = label;
                    build.tasks.push_back({item.nodeId, child, static_cast<uint8_t>(draw.depth + 1u), build.draw.size()});
                    std::push_heap(build.tasks.begin(), build.tasks.end(), taskLess);
                }
                else if (hasChildren && fits)
                {
                    ++build.expandableUnexpanded;
                }
                if (draw.depth > 0u)
                {
                    if (node)
                    {
                        ++build.nestedDirectories;
                    }
                    else if (IsFileRecordId(draw.nodeId))
                    {
                        ++build.nestedFiles;
                    }
                }
                build.hasPositiveGeometry = build.hasPositiveGeometry || (item.bytes != 0u && draw.lod != DrawItem::Lod::Culled && Area(rect) > 0.0f);
                build.draw.push_back(draw);
                break;
            }
            case LayoutBuildState::Phase::Finalize:
            {
                if (build.cursor < build.draw.size())
                {
                    auto& item          = build.draw[build.cursor++];
                    const auto previous = build.previous.find(item.nodeId);
                    const auto& target  = item.targetRect;
                    const float cx      = (target.left + target.right) * 0.5f;
                    const float cy      = (target.top + target.bottom) * 0.5f;
                    item.startRect      = previous == build.previous.end() ? D2D1::RectF(cx, cy, cx, cy) : previous->second;
                    const bool moved    = std::fabs(item.startRect.left - target.left) > 0.5f || std::fabs(item.startRect.top - target.top) > 0.5f ||
                                          std::fabs(item.startRect.right - target.right) > 0.5f || std::fabs(item.startRect.bottom - target.bottom) > 0.5f;
                    item.animationStartSeconds = build.animationTime - (moved ? 0.0 : kAnimationSeconds);
                    item.currentRect           = moved ? item.startRect : target;
                    break;
                }
                build.cursor = 0u;
                if (build.draw.size() >= 256u && Area(bounds) > 1.0f)
                {
                    build.grid.bounds  = bounds;
                    build.grid.columns = static_cast<uint32_t>(std::clamp(std::ceil((bounds.right - bounds.left) / 96.0), 8.0, 64.0));
                    build.grid.rows    = static_cast<uint32_t>(std::clamp(std::ceil((bounds.bottom - bounds.top) / 96.0), 8.0, 64.0));
                    build.grid.cells.resize(static_cast<size_t>(build.grid.columns) * build.grid.rows);
                    build.phase = LayoutBuildState::Phase::Grid;
                }
                else
                {
                    build.phase = LayoutBuildState::Phase::Publish;
                }
                break;
            }
            case LayoutBuildState::Phase::Grid:
            {
                if (build.cursor == build.draw.size())
                {
                    build.phase = LayoutBuildState::Phase::Publish;
                    break;
                }
                if (! build.gridItemActive)
                {
                    const auto& item = build.draw[build.cursor];
                    auto rect        = item.targetRect;
                    const float gap  = std::clamp(kGap - static_cast<float>(item.depth) * 0.15f, 0.5f, kGap);
                    rect.left += gap;
                    rect.top += gap;
                    rect.right -= gap;
                    rect.bottom -= gap;
                    if (item.lod == DrawItem::Lod::Culled || Area(rect) < 16.0f || rect.right <= rect.left || rect.bottom <= rect.top)
                    {
                        ++build.cursor;
                        break;
                    }
                    const auto cell = [](float value, float origin, float extent, uint32_t count) noexcept
                    {
                        return static_cast<uint32_t>(
                            std::clamp(std::floor(static_cast<double>((value - origin) / extent) * count), 0.0, static_cast<double>(count - 1u)));
                    };
                    build.gridLeft       = cell(rect.left, bounds.left, bounds.right - bounds.left, build.grid.columns);
                    build.gridRight      = cell(rect.right, bounds.left, bounds.right - bounds.left, build.grid.columns);
                    build.gridTop        = cell(rect.top, bounds.top, bounds.bottom - bounds.top, build.grid.rows);
                    build.gridBottom     = cell(rect.bottom, bounds.top, bounds.bottom - bounds.top, build.grid.rows);
                    build.gridX          = build.gridLeft;
                    build.gridY          = build.gridTop;
                    build.gridItemActive = true;
                }
                auto& cell = build.grid.cells[static_cast<size_t>(build.gridY) * build.grid.columns + build.gridX];
                cell.push_back(static_cast<uint32_t>(build.cursor));
                build.grid.maxCandidatesPerCell = std::max(build.grid.maxCandidatesPerCell, static_cast<uint32_t>(cell.size()));
                if (++build.gridX > build.gridRight)
                {
                    build.gridX = build.gridLeft;
                    if (++build.gridY > build.gridBottom)
                    {
                        build.gridItemActive = false;
                        ++build.cursor;
                    }
                }
                break;
            }
            case LayoutBuildState::Phase::Publish:
            {
                if (! HasDeferredDisposalCapacity())
                {
                    disposalCapacityBlocked = true;
                    break;
                }

                // Measure the coherent adoption and deferred handoff of the previous geometry/index/name owner.
                const bool capturePublishTiming = Debug::Perf::IsCaptureEnabled();
#ifdef ENABLE_TESTS
                const bool capturePublishPhases = _debugPublicationPhaseTiming && capturePublishTiming;
#else
                const bool capturePublishPhases = false;
#endif
                const auto publishStarted             = capturePublishTiming ? std::chrono::steady_clock::now() : std::chrono::steady_clock::time_point{};
                auto phaseStarted                     = capturePublishPhases ? std::chrono::steady_clock::now() : std::chrono::steady_clock::time_point{};
                uint64_t cacheInvalidateUs            = 0u;
                uint64_t retireAndHandoffUs           = 0u;
                const uint64_t publishedGeometryCount = build.draw.size();
                const uint32_t publishedGeneration    = build.generation;
                const uint64_t nestedDirectories      = build.nestedDirectories;
                const uint64_t nestedFiles            = build.nestedFiles;
                const uint64_t expandableUnexpanded   = build.expandableUnexpanded;
                auto retiredPublishedLayoutOwner      = std::move(_publishedLayoutOwner);
                // Geometry, synthetic names and hit index become visible as one UI-thread adoption.
                if (capturePublishPhases)
                {
                    phaseStarted = std::chrono::steady_clock::now();
                }
                InvalidateStaticTreemapCache();
                if (capturePublishPhases)
                {
                    cacheInvalidateUs = Debug::Perf::ElapsedUs(phaseStarted);
                }
                _hasPublishedPositiveGeometry = build.hasPositiveGeometry;
                _drawItems.swap(build.draw);
                _syntheticNodes.swap(build.synthetic);
                std::swap(_hitGrid, build.grid);
                // Move only stable synthetic-name storage into the new published owner. The old
                // geometry, index, scratch and prior names stay together until worker disposal.
                auto publishedLayoutOwner = std::make_shared<LayoutBuildState>();
                publishedLayoutOwner->names.swap(build.names);
                _publishedLayoutOwner     = std::move(publishedLayoutOwner);
                auto retiredLayoutBuild   = std::move(_layoutBuild);
                _layoutDirty              = false;
                _lastLayoutRebuildSeconds = NowSeconds();
#ifdef ENABLE_TESTS
                ++_layoutGeneration;
#endif
                Debug::Perf::EmitValue(L"viewer.space.layout.draw_items", _drawItems.size());
                Debug::Perf::EmitValue(L"viewer.space.layout.nested_directory_tiles", nestedDirectories);
                Debug::Perf::EmitValue(L"viewer.space.layout.nested_file_tiles", nestedFiles);
                Debug::Perf::EmitValue(L"viewer.space.layout.expandable_unexpanded_folders", expandableUnexpanded);
                if (_hWnd)
                {
                    InvalidateRect(_hWnd.get(), nullptr, FALSE);
                }
                if (capturePublishPhases)
                {
                    phaseStarted = std::chrono::steady_clock::now();
                }
                RetireLayoutStorage(std::move(retiredLayoutBuild), std::move(retiredPublishedLayoutOwner));
                if (capturePublishPhases)
                {
                    retireAndHandoffUs = Debug::Perf::ElapsedUs(phaseStarted);
                }
                published = true;
                if (capturePublishTiming)
                {
                    const uint64_t publishElapsedUs = Debug::Perf::ElapsedUs(publishStarted);
                    Debug::Perf::Emit(L"viewer.space.layout.publish_us",
                                      L"commit-and-handoff-old-layout",
                                      publishElapsedUs,
                                      publishedGeometryCount,
                                      publishedGeneration,
                                      S_OK);
                    if (capturePublishPhases)
                    {
                        // Endpoints precede phase rows; their observer writes remain in layout.turn_us.
                        Debug::Perf::Emit(
                            L"viewer.space.layout.publish_phase_us", L"cache-invalidate", cacheInvalidateUs, publishedGeometryCount, publishedGeneration, S_OK);
                        Debug::Perf::Emit(L"viewer.space.layout.publish_phase_us",
                                          L"retire-and-handoff",
                                          retireAndHandoffUs,
                                          publishedGeometryCount,
                                          publishedGeneration,
                                          S_OK);
                    }
                }
                break;
            }
        }
        if (published || disposalCapacityBlocked)
        {
            break;
        }
    }
    const uint64_t elapsed = Debug::Perf::ElapsedUs(started);
    ++_layoutTurnCount;
    _layoutMaxTurnUs = std::max(_layoutMaxTurnUs, elapsed);
#ifdef ENABLE_TESTS
    _lastLayoutUs = elapsed;
#endif
    Debug::Perf::Emit(L"viewer.space.layout.turn_us", published ? L"publish" : L"continue", elapsed, work, _drawItems.size(), S_OK);
    // Retain the existing metric family for source-bound baseline comparison.
    Debug::Perf::Emit(L"viewer.space.layout.rebuild_us", published ? L"publish" : L"continue", elapsed, _drawItems.size(), 0u, S_OK);
}
