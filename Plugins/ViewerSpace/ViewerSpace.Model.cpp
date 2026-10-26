#include "ViewerSpace.h"

#include <array>
#include <exception>
#include <limits>

namespace
{
// Independent of each viewer's bounded CPU owners: at most three final device
// releases may be deferred across the module. Capacity fallback stays on the UI.
std::atomic<uint32_t> g_retiredRendererDevices{0u};
constexpr uint32_t kMaxRetiredRendererDevices = 3u;
} // namespace

ViewerSpace::RetiredD3DDevice::~RetiredD3DDevice() noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    device.reset();
    Debug::Perf::EmitDurationUs(L"viewer.space.render.retired_device.release_us", Debug::Perf::ElapsedUs(startedAt));
    g_retiredRendererDevices.fetch_sub(1u, std::memory_order_release);
}

uint32_t ViewerSpace::RetiredRendererDeviceCount() noexcept
{
    return g_retiredRendererDevices.load(std::memory_order_acquire);
}

std::unique_ptr<ViewerSpace::RetiredD3DDevice> ViewerSpace::CaptureClosingRendererDevice() noexcept
{
    if (! _renderer.d3dDevice || (_renderer.d3dDevice->GetCreationFlags() & D3D11_CREATE_DEVICE_SINGLETHREADED) != 0u)
    {
        return {};
    }
    auto observed = g_retiredRendererDevices.load(std::memory_order_relaxed);
    while (observed < kMaxRetiredRendererDevices)
    {
        if (g_retiredRendererDevices.compare_exchange_weak(observed, observed + 1u, std::memory_order_acq_rel, std::memory_order_relaxed))
        {
            auto retired = std::make_unique<RetiredD3DDevice>(std::move(_renderer.d3dDevice));
            Debug::Perf::EmitCounter(L"viewer.space.render.retired_device.admitted");
            return retired;
        }
    }
    Debug::Perf::EmitCounter(L"viewer.space.render.retired_device.capacity_fallback");
    return {};
}

void ViewerSpace::RetiredDisplayStorage::ReleaseStorage() noexcept
{
    rendererDevice.reset();
    layoutBuild.reset();
    cancelledLayoutBuild.reset();
    publishedLayoutOwner.reset();
    analysisBuild.reset();
    publishedAnalysisOwner.reset();
    std::vector<uint32_t>().swap(analysisRows);
    std::unordered_map<uint32_t, Node>().swap(syntheticNodes);
    std::unordered_map<uint32_t, uint32_t>().swap(layoutMaxItemsByNode);
    std::unordered_set<uint32_t>().swap(autoExpandedOtherByNode);
    std::unordered_map<uint32_t, LayoutCandidateCacheEntry>().swap(layoutCandidateCache);
    std::vector<DrawItem>().swap(drawItems);
    std::vector<LayoutWorkspace>().swap(layoutWorkspaceByDepth);
    std::vector<std::vector<uint32_t>>().swap(hitGrid.cells);
    std::unordered_map<uint32_t, D2D1_RECT_F>().swap(lastRectsByNode);
    displayModel.reset();
    workingModel.reset();
}

void ViewerSpace::RetireDisplayedStorage(std::unique_ptr<ModelGeneration> displayModel,
                                         std::unique_ptr<ModelGeneration> workingModel,
                                         bool includeDerived,
                                         bool closing,
                                         std::unique_ptr<RetiredD3DDevice> rendererDevice) noexcept
{
    auto storage            = std::make_unique<RetiredDisplayStorage>();
    storage->rendererDevice = std::move(rendererDevice);
    uint64_t bytes          = displayModel == nullptr ? 0u : displayModel->AllocatedBytes();
    if (workingModel != nullptr)
    {
        const uint64_t candidateBytes = workingModel->AllocatedBytes();
        bytes = candidateBytes > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + candidateBytes;
    }
    const auto add = [&bytes](uint64_t amount) noexcept
    { bytes = amount > std::numeric_limits<uint64_t>::max() - bytes ? std::numeric_limits<uint64_t>::max() : bytes + amount; };
    if (includeDerived)
    {
        add(LayoutOwnedBytes());
        add(AnalysisOwnedBytes());
    }

    storage->displayModel = std::move(displayModel);
    storage->workingModel = std::move(workingModel);
    if (includeDerived)
    {
        storage->layoutBuild             = std::move(_layoutBuild);
        storage->cancelledLayoutBuild    = std::move(_cancelledLayoutBuild);
        storage->publishedLayoutOwner    = std::move(_publishedLayoutOwner);
        storage->analysisBuild           = std::move(_analysisBuild);
        storage->publishedAnalysisOwner  = std::move(_publishedAnalysisOwner);
        storage->analysisRows            = std::move(_analysisRows);
        storage->syntheticNodes          = std::move(_syntheticNodes);
        storage->layoutMaxItemsByNode    = std::move(_layoutMaxItemsByNode);
        storage->autoExpandedOtherByNode = std::move(_autoExpandedOtherByNode);
        storage->layoutCandidateCache    = std::move(_layoutCandidateCache);
        storage->drawItems               = std::move(_drawItems);
        storage->layoutWorkspaceByDepth  = std::move(_layoutWorkspaceByDepth);
        storage->hitGrid                 = std::move(_hitGrid);
        storage->lastRectsByNode         = std::move(_lastRectsByNode);
    }

    QueueDeferredDisposal(std::move(storage), bytes, closing);
}

void ViewerSpace::QueueDeferredDisposal(std::unique_ptr<RetiredDisplayStorage> storage, uint64_t bytes, bool closing) noexcept
{
    static constexpr size_t kMaxNormalDeferredDisposalOwners  = 2u;
    static constexpr size_t kMaxClosingDeferredDisposalOwners = 3u;
    ReapFinishedDeferredDisposals();
    const size_t ownerLimit = closing ? kMaxClosingDeferredDisposalOwners : kMaxNormalDeferredDisposalOwners;
    if (_deferredDisposalWorkers.size() >= ownerLimit)
    {
        Debug::Error(L"ViewerSpace: deferred disposal capacity invariant violated (closing={})", closing);
        std::terminate();
    }

    auto observed = _deferredDisposalBytes.load(std::memory_order_relaxed);
    while (bytes > std::numeric_limits<uint64_t>::max() - observed ||
           ! _deferredDisposalBytes.compare_exchange_weak(observed, observed + bytes, std::memory_order_release, std::memory_order_relaxed))
    {
        if (bytes > std::numeric_limits<uint64_t>::max() - observed)
        {
            bytes = std::numeric_limits<uint64_t>::max() - observed;
        }
    }

    DeferredDisposalWorker worker;
    worker.storage = std::shared_ptr<RetiredDisplayStorage>(std::move(storage));
    worker.done    = std::make_shared<std::atomic_bool>(false);
    worker.bytes   = bytes;
    _deferredDisposalWorkers.emplace_back(std::move(worker));
    const size_t index = _deferredDisposalWorkers.size() - 1u;
    static_cast<void>(SubmitDeferredDisposal(index));
}

bool ViewerSpace::HasDeferredDisposalCapacity(bool closing) noexcept
{
    ReapFinishedDeferredDisposals();
#ifdef ENABLE_TESTS
    if (! closing)
    {
        std::array<wchar_t, 2> forcedFull{};
        if (GetEnvironmentVariableW(L"REDSALAMANDER_VIEWERSPACE_TEST_FORCE_DISPOSAL_SATURATION", forcedFull.data(), static_cast<DWORD>(forcedFull.size())) > 0u)
        {
            return false;
        }
    }
#endif
    const size_t ownerLimit = closing ? 3u : 2u;
    return _deferredDisposalWorkers.size() < ownerLimit;
}

void ViewerSpace::CancelDeferredStartScan() noexcept
{
    ++_deferredStartScanGeneration;
    _deferredStartScan.reset();
    UpdateHeaderTextCache();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::ContinueDeferredStartScan() noexcept
{
    if (! _deferredStartScan.has_value() || ! HasDeferredDisposalCapacity())
    {
        return;
    }

    const DeferredStartScan& pending = _deferredStartScan.value();
    const size_t ownersNeeded        = static_cast<size_t>(_workingModelRetirementPending) +
                                       static_cast<size_t>(! pending.retainPrevious && (_displayModel != nullptr || _workingModel != nullptr));
    if (ownersNeeded > 1u && _deferredDisposalWorkers.size() + ownersNeeded > 2u)
    {
        return;
    }

    DeferredStartScan request = std::move(_deferredStartScan.value());
    _deferredStartScan.reset();
    if (request.requestGeneration != _deferredStartScanGeneration)
    {
        return;
    }
    StartScan(request.rootPath, request.allowCache, request.retainPrevious);
    Debug::Perf::EmitCounter(L"viewer.space.refresh.start_resumed");
}

void ViewerSpace::ContinueDeferredModelRetirement() noexcept
{
    if (! _workingModelRetirementPending || _deferredStartScan.has_value() || ! HasDeferredDisposalCapacity())
    {
        return;
    }

    _workingModelRetirementPending = false;
    RetireDisplayedStorage({}, std::move(_workingModel), false);
}

uint64_t ViewerSpace::ModelGeneration::AllocatedBytes() const noexcept
{
    uint64_t bytes = memory.AllocatedBytes();
    const auto add = [&bytes](uint64_t value) noexcept
    {
        if (value > std::numeric_limits<uint64_t>::max() - bytes)
        {
            bytes = std::numeric_limits<uint64_t>::max();
        }
        else
        {
            bytes += value;
        }
    };

    add(static_cast<uint64_t>(childrenArenaFreeBlocks.capacity()) * static_cast<uint64_t>(sizeof(ChildArenaFreeBlock)));
    add(static_cast<uint64_t>(scanRootPath.capacity()) * static_cast<uint64_t>(sizeof(wchar_t)));
    if (workspacePathIdentity.has_value())
        add(static_cast<uint64_t>(workspacePathIdentity.value().acceptedSeparators.capacity()) * sizeof(wchar_t));
    if (scanRootParentPath.has_value())
    {
        add(static_cast<uint64_t>(scanRootParentPath.value().capacity()) * static_cast<uint64_t>(sizeof(wchar_t)));
    }
    return bytes;
}

void ViewerSpace::RefreshCurrent() noexcept
{
    if (_refreshPending || _refreshAdoptionPending)
    {
        return;
    }
    const std::wstring rootPath = BuildNodePathText(_model->viewNodeId);
    const bool retainPrevious   = ! _scanActive.load(std::memory_order_acquire) && _overallState != ScanState::Queued && _overallState != ScanState::Scanning;
    if (! rootPath.empty())
    {
        StartScan(rootPath, false, retainPrevious);
    }
}

void ViewerSpace::RebuildCurrent() noexcept
{
    const std::wstring rootPath = BuildNodePathText(_model->viewNodeId);
    if (! rootPath.empty())
    {
        StartScan(rootPath, false, false);
    }
}

void ViewerSpace::DiscardWorkingModel(bool refused) noexcept
{
    if (_workingModel == nullptr)
    {
        _refreshRefused = refused;
        return;
    }

    _scanGeneration.fetch_add(1u, std::memory_order_acq_rel);
    CancelScan();
    CancelScanCacheBuild();
    CancelScanCacheRestore();
    {
        std::scoped_lock lock(_updateMutex);
        _pendingUpdates.clear();
        _pendingUpdateBytes.store(0u, std::memory_order_relaxed);
    }

    _model = _displayModel.get();
    if (HasDeferredDisposalCapacity())
    {
        RetireDisplayedStorage({}, std::move(_workingModel), false);
    }
    else
    {
        _workingModelRetirementPending = true;
    }
    _workingScanGeneration    = 0u;
    _refreshPending           = false;
    _refreshRefused           = refused;
    _refreshAdoptionPending   = false;
    _refreshPathLookupStarted = false;
    _refreshPathCandidates.clear();
    _refreshNextPathCandidates.clear();
    _refreshPathCandidateIndex  = 0u;
    _refreshPathChildIndex      = 0u;
    _refreshResolvedSelectionId = 0u;
    _scanActive.store(false, std::memory_order_release);

    if (_refreshDisplayState.has_value())
    {
        const RefreshDisplayState& state  = _refreshDisplayState.value();
        _overallState                     = state.overallState;
        _scanCompletedSinceSeconds        = state.completedSinceSeconds;
        _scanProgressBytes                = state.progressBytes;
        _scanProgressFolders              = state.progressFolders;
        _scanProgressFiles                = state.progressFiles;
        _scanFailedFolders                = state.failedFolders;
        _scanEffectiveThreads             = state.effectiveThreads;
        _scanTopFilesPerDirectory         = state.topFilesPerDirectory;
        _scanCacheLastStoredGeneration    = state.cacheLastStoredGeneration;
        _lastScanCacheSnapshotBytes       = state.cacheSnapshotBytes;
        _lastScanCacheRestoreBytes        = state.cacheRestoreBytes;
        _lastScanCacheResidentBytes       = state.cacheResidentBytes;
        _lastScanCacheRestoreUs           = state.cacheRestoreUs;
        _scanProcessingNodeId             = state.processingNodeId;
        _scanProcessingFolderName         = state.processingFolderName;
        _lastModelAcceptedEntries         = state.acceptedEntries;
        _lastModelRejectedEntries         = state.rejectedEntries;
        _lastModelCappedDirectories       = state.cappedDirectories;
        _lastModelCappedFiles             = state.cappedFiles;
        _lastModelRetainedNameBytes       = state.retainedNameBytes;
        _lastModelRetainedChildReferences = state.retainedChildReferences;
        _lastModelTraversedDirectories    = state.traversedDirectories;
        _lastModelValidationError         = state.validationError;
        _directorySummaryCommittedCount   = state.summaryCommittedCount;
        _directorySummaryRejectedCount    = state.summaryRejectedCount;
        _directorySummaryChunkMaxBytes.store(state.summaryChunkMaxBytes, std::memory_order_relaxed);
        _model->viewNodeId     = state.viewNodeId;
        _viewPathText          = state.viewPathText;
        _navStack              = state.navStack;
        _navForwardStack       = state.navForwardStack;
        _selectedItemId        = state.selectedItemId;
        _selectedGeneration    = _displayGeneration;
        _selectionFromAnalysis = state.selectionFromAnalysis;
        _refreshDisplayState.reset();
    }

    UpdateHeaderTextCache();
    UpdateInspector();
    if (_hWnd)
    {
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

// This UI-thread-only projection helper is local to ViewerSpace because its inputs are generation-owned
// retained files, committed Other bytes, and the exact real-directory parent graph.
bool ViewerSpace::RaiseNodeObservedBytes(uint32_t nodeId, uint64_t reportedBytes) noexcept
{
    bool changed = false;
    for (uint32_t depth = 0u; nodeId != 0u && depth <= 512u; ++depth)
    {
        Node* node = TryGetRealNode(nodeId);
        if (node == nullptr)
        {
            return changed;
        }

        const uint64_t ownAndOther = ViewerSpaceScan::SaturatingAdd(node->observedOwnFileBytes, node->aggregateBytes);
        const uint64_t derived     = ViewerSpaceScan::SaturatingAdd(ownAndOther, node->observedChildBytes);
        const uint64_t next        = std::max(node->totalBytes, std::max(reportedBytes, derived));
        const uint64_t delta       = next - node->totalBytes;
        if (delta == 0u)
        {
            return changed;
        }

        node->totalBytes = next;
        changed          = true;
        if (nodeId == _model->rootNodeId || node->parentId == 0u)
        {
            return changed;
        }

        Node* parent = TryGetRealNode(node->parentId);
        if (parent == nullptr)
        {
            return changed;
        }
        parent->observedChildBytes = ViewerSpaceScan::SaturatingAdd(parent->observedChildBytes, delta);
        nodeId                     = parent->id;
        reportedBytes              = 0u;
    }

    if (nodeId != 0u)
    {
        Debug::Error(L"ViewerSpace: observed byte propagation exceeded the 512-level scan depth bound");
    }
    return changed;
}

void ViewerSpace::AdoptWorkingModel() noexcept
{
    if (_workingModel == nullptr || ! _refreshAdoptionPending)
    {
        return;
    }

    if (! HasDeferredDisposalCapacity())
    {
        EnsureAnimationTimer();
        return;
    }

    _analysisDirty = true;
    InvalidateStaticTreemapCache();

    std::unique_ptr<ModelGeneration> previousModel = std::move(_displayModel);
    _displayModel                                  = std::move(_workingModel);
    _model                                         = _displayModel.get();
    RetireDisplayedStorage(std::move(previousModel));
    _displayGeneration = _displayModel->generation;
    if (_displayGeneration == 0u)
    {
        _displayGeneration = _workingScanGeneration;
    }
    _model->viewNodeId = _model->rootNodeId;
    _analysisHost.NotifyRowsChanged(std::nullopt);
    _navStack.clear();
    _navForwardStack.clear();
    _refreshPending           = false;
    _refreshRefused           = false;
    _refreshAdoptionPending   = false;
    _refreshPathLookupStarted = false;
    _workingScanGeneration    = 0u;

    _selectedItemId             = _refreshResolvedSelectionId;
    _selectedGeneration         = _displayGeneration;
    _selectionFromAnalysis      = false;
    _refreshResolvedSelectionId = 0u;
    _hoverNodeId                = 0u;
    _tooltipNodeId              = 0u;
    _refreshDisplayState.reset();
    _layoutDirty                      = true;
    _lastLayoutRebuildSeconds         = 0.0;
    _firstPositiveProjectionAttempted = false;
    _hasPublishedPositiveGeometry     = false;
    ++_modelRevision;
    UpdateViewPathText();
    InvalidateAnalysis();
    UpdateHeaderTextCache();
    UpdateInspector();
    if (_hWnd)
    {
        UpdateWindowTitle(_hWnd.get());
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::ContinueRefreshAdoption() noexcept
{
    if (! _refreshAdoptionPending || _workingModel == nullptr || ! _refreshDisplayState.has_value())
    {
        return;
    }

    const auto adoptionTurnStartedAt               = std::chrono::steady_clock::now();
    [[maybe_unused]] const auto recordAdoptionTurn = wil::scope_exit([this, adoptionTurnStartedAt]() noexcept
    {
        const uint64_t elapsedUs = Debug::Perf::ElapsedUs(adoptionTurnStartedAt);
        uint64_t observed        = _refreshAdoptionMaxUs.load(std::memory_order_relaxed);
        while (elapsedUs > observed && ! _refreshAdoptionMaxUs.compare_exchange_weak(observed, elapsedUs, std::memory_order_relaxed))
        {
        }
        Debug::Perf::EmitDurationUs(
            L"viewer.space.refresh.adoption_turn_us", elapsedUs, _refreshPathChildIndex, static_cast<uint64_t>(_deferredDisposalWorkers.size()), S_OK);
        Debug::Perf::EmitValue(L"viewer.space.refresh.adoption_turn_max_us", _refreshAdoptionMaxUs.load(std::memory_order_relaxed));
    });

    RefreshDisplayState& state = _refreshDisplayState.value();
    if (! _refreshPathLookupStarted)
    {
        _refreshPathLookupStarted = true;
        _refreshPathCandidates.clear();
        _refreshNextPathCandidates.clear();
        _refreshPathCandidateIndex  = 0u;
        _refreshPathChildIndex      = 0u;
        state.selectedItemId        = _selectedItemId;
        state.selectionFromAnalysis = _selectionFromAnalysis;
        state.selectedRelativePath.clear();
        state.selectionPathIsFile = false;
        if (_selectedItemId != 0u && _selectedItemId != _model->viewNodeId)
        {
            if (const FileRecord* file = TryGetFileRecord(_selectedItemId))
            {
                state.selectionPathIsFile = true;
                state.selectedRelativePath.emplace_back(file->name);
                uint32_t parentId = file->parentId;
                while (parentId != 0u && parentId != _model->viewNodeId)
                {
                    const Node* parent = TryGetRealNode(parentId);
                    if (parent == nullptr || parent->isSynthetic || parent->name.empty())
                    {
                        state.selectedRelativePath.clear();
                        break;
                    }
                    state.selectedRelativePath.emplace_back(parent->name);
                    parentId = parent->parentId;
                }
                if (parentId != _model->viewNodeId)
                {
                    state.selectedRelativePath.clear();
                }
            }
            else if (const Node* node = TryGetRealNode(_selectedItemId); node != nullptr && ! node->isSynthetic)
            {
                uint32_t currentId = node->id;
                while (currentId != 0u && currentId != _model->viewNodeId)
                {
                    const Node* current = TryGetRealNode(currentId);
                    if (current == nullptr || current->name.empty())
                    {
                        state.selectedRelativePath.clear();
                        break;
                    }
                    state.selectedRelativePath.emplace_back(current->name);
                    currentId = current->parentId;
                }
                if (currentId != _model->viewNodeId)
                {
                    state.selectedRelativePath.clear();
                }
            }
            std::reverse(state.selectedRelativePath.begin(), state.selectedRelativePath.end());
        }
        if (state.selectedItemId == state.viewNodeId)
        {
            _refreshResolvedSelectionId = _workingModel->rootNodeId;
            AdoptWorkingModel();
            return;
        }
        if (state.selectedRelativePath.empty())
        {
            _refreshResolvedSelectionId = 0u;
            AdoptWorkingModel();
            return;
        }
        _refreshPathCandidates.push_back(_workingModel->rootNodeId);
    }

    ModelGeneration* const previousModel = _model;
    _model                               = _workingModel.get();
    auto restoreModel                    = wil::scope_exit([this, previousModel]() noexcept { _model = previousModel; });
    const auto started                   = std::chrono::steady_clock::now();
    size_t checked                       = 0u;
    while (state.selectionPathPart < state.selectedRelativePath.size())
    {
        if (_refreshPathCandidateIndex >= _refreshPathCandidates.size())
        {
            _refreshPathCandidates.swap(_refreshNextPathCandidates);
            _refreshNextPathCandidates.clear();
            _refreshPathCandidateIndex = 0u;
            _refreshPathChildIndex     = 0u;
            state.selectionPathPart += 1u;
            if (_refreshPathCandidates.empty())
            {
                _refreshResolvedSelectionId = 0u;
                _model                      = previousModel;
                restoreModel.release();
                AdoptWorkingModel();
                return;
            }
            continue;
        }

        const uint32_t parentId = _refreshPathCandidates[_refreshPathCandidateIndex];
        const Node* parent      = TryGetRealNode(parentId);
        if (parent == nullptr || ! parent->isDirectory)
        {
            _refreshPathCandidateIndex += 1u;
            _refreshPathChildIndex = 0u;
            continue;
        }
        const uint64_t childEnd = static_cast<uint64_t>(parent->childrenStart) + parent->childrenCount;
        if (childEnd > _model->childrenArena.size())
        {
            _refreshPathCandidateIndex += 1u;
            _refreshPathChildIndex = 0u;
            continue;
        }
        if (_refreshPathChildIndex >= parent->childrenCount)
        {
            _refreshPathCandidateIndex += 1u;
            _refreshPathChildIndex = 0u;
            continue;
        }

        const uint32_t childId = _model->childrenArena[static_cast<size_t>(parent->childrenStart) + _refreshPathChildIndex++];
        const bool lastPart    = state.selectionPathPart + 1u == state.selectedRelativePath.size();
        bool nameMatches       = false;
        if (childId >= kFileRecordIdBase)
        {
            if (lastPart && state.selectionPathIsFile)
            {
                const FileRecord* file = TryGetFileRecord(childId);
                nameMatches            = file != nullptr && file->parentId == parentId && file->name == state.selectedRelativePath[state.selectionPathPart];
            }
        }
        else if (const Node* node = TryGetRealNode(childId); node != nullptr)
        {
            nameMatches = node->name == state.selectedRelativePath[state.selectionPathPart] && (! lastPart || ! state.selectionPathIsFile);
        }
        if (nameMatches && _refreshNextPathCandidates.size() < 16u)
        {
            _refreshNextPathCandidates.push_back(childId);
        }

        ++checked;
        if (checked >= 128u && std::chrono::duration<double>(std::chrono::steady_clock::now() - started).count() >= 0.0012)
        {
            return;
        }
    }

    _refreshResolvedSelectionId = _refreshPathCandidates.size() == 1u ? _refreshPathCandidates.front() : 0u;
    _model                      = previousModel;
    restoreModel.release();
    AdoptWorkingModel();
}
