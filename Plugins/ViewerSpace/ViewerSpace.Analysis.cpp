#include "ViewerSpace.h"
#include "FileMetadataFormatting.h"
#include "LocalizationManager.h"
#include "NaturalSort.h"
#include "ViewerDxUiTheme.h"
#include "ViewerSpace.ListingValidation.h"
#include "ViewerSpace.SlicedSort.h"
#include "resource.h"

#include <algorithm>
#include <cmath>
#include <utility>

extern HINSTANCE g_hInstance;

struct ViewerSpace::AnalysisBuildState final
{
    uint32_t generation   = 0u;
    uint64_t revision     = 0u;
    uint32_t view         = 0u;
    size_t cursor         = 0u;
    size_t rowIndexCursor = 0u;
    bool sorting          = false;
    bool indexing         = false;
    std::vector<uint32_t> rows;
    std::array<uint64_t, 8u> typeBytes{};
    std::array<uint64_t, 8u> typeCounts{};
    ViewerSpaceDetail::SlicedSort<uint32_t> sort;
    std::vector<std::pair<uint32_t, size_t>> rowIndex;
    ViewerSpaceDetail::SlicedSort<std::pair<uint32_t, size_t>> rowIndexSort;
};

void ViewerSpace::CreateAnalysisHost(HWND hwnd)
{
    using Host = ViewerSpaceAnalysis::AnalysisHost;
    auto text  = [](UINT id) { return LoadStringResource(g_hInstance, id); };
    Host::Captions captions;
    captions.filterLabel        = text(IDS_VIEWERSPACE_ANALYSIS_FILTER_LABEL);
    captions.filterPlaceholder  = text(IDS_VIEWERSPACE_ANALYSIS_FILTER_PLACEHOLDER);
    captions.filterScope        = text(IDS_VIEWERSPACE_ANALYSIS_FILTER_SCOPE);
    captions.gridAccessibleName = text(IDS_VIEWERSPACE_ANALYSIS_GRID_ACCESSIBLE_NAME);
    captions.emptyState         = text(IDS_VIEWERSPACE_ANALYSIS_EMPTY_STATE);
    captions.columns            = {text(IDS_VIEWERSPACE_ANALYSIS_COLUMN_NAME),
                                   text(IDS_VIEWERSPACE_ANALYSIS_COLUMN_LOGICAL_BYTES),
                                   text(IDS_VIEWERSPACE_ANALYSIS_COLUMN_PARENT_SHARE),
                                   text(IDS_VIEWERSPACE_CLEANUP_COLUMN),
                                   text(IDS_VIEWERSPACE_ANALYSIS_COLUMN_STATE)};
    captions.back               = text(IDS_VIEWERSPACE_NAV_BACK);
    captions.reviewCleanup      = text(IDS_VIEWERSPACE_REVIEW_CLEANUP);
    captions.recycle            = text(IDS_VIEWERSPACE_REVIEW_RECYCLE);
    captions.permanentDelete    = text(IDS_VIEWERSPACE_REVIEW_DELETE);
    captions.inspectorTitle     = text(IDS_VIEWERSPACE_ANALYSIS_INSPECTOR_TITLE);
    captions.inspectorFields    = {text(IDS_VIEWERSPACE_ANALYSIS_FIELD_PATH),
                                   text(IDS_VIEWERSPACE_ANALYSIS_FIELD_NAME),
                                   text(IDS_VIEWERSPACE_ANALYSIS_FIELD_KIND),
                                   text(IDS_VIEWERSPACE_ANALYSIS_FIELD_LOGICAL_BYTES),
                                   text(IDS_VIEWERSPACE_ANALYSIS_FIELD_PARENT_SHARE),
                                   text(IDS_VIEWERSPACE_ANALYSIS_FIELD_MODIFIED),
                                   text(IDS_VIEWERSPACE_ANALYSIS_FIELD_STATE)};
    Host::Callbacks callbacks;
    callbacks.rowExcluded = [this](size_t row) { return row < _analysisRows.size() && IsItemExcluded(_analysisRows[row]); };
    callbacks.rowCount    = [this] { return _analysisRows.size(); };
    callbacks.rowId       = [this](size_t row) { return AnalysisRowToken(row); };
    callbacks.findRow     = [this](uint64_t token) { return FindAnalysisRow(token); };
    callbacks.cellText    = [this](size_t row, size_t column, DxUi::GridCellData& cell)
    {
        ++_analysisCellRequests;
        GetAnalysisCell(row, column, cell);
    };
    callbacks.selectionChanged = [this](std::optional<uint64_t> token)
    {
        if (! token.has_value())
        {
            SelectItem(0u, false);
            return;
        }
        if (FindAnalysisRow(token.value()).has_value())
        {
            SelectItem(static_cast<uint32_t>(token.value()), false);
        }
    };
    callbacks.rowActivated = [this](uint64_t token)
    {
        if (FindAnalysisRow(token).has_value())
        {
            ActivateItem(static_cast<uint32_t>(token));
        }
    };
    callbacks.rowContextMenu = [this](uint64_t token, POINT point)
    {
        if (! FindAnalysisRow(token).has_value())
        {
            return;
        }
        SelectItem(static_cast<uint32_t>(token), false);
        _contextItemIdOverride = _selectedItemId;
        OnContextMenu(_hWnd.get(), point);
        _contextItemIdOverride = 0u;
    };
    callbacks.sortRequested = [this](const DxUi::GridSortSpec& sort)
    {
        _analysisSortColumn    = sort.columnIndex;
        _analysisSortAscending = sort.direction == DxUi::SortDirection::Ascending;
        InvalidateAnalysis();
    };
    callbacks.exclusionToggled = [this](uint64_t token)
    {
        if (FindAnalysisRow(token).has_value())
            ToggleExcludedItem(static_cast<uint32_t>(token));
    };
    callbacks.cleanupToggled = [this](uint64_t token)
    {
        if (FindAnalysisRow(token).has_value())
            ToggleCleanupMark(static_cast<uint32_t>(token));
    };
    callbacks.workspaceAction = [this](UINT action) { OnCommand(_hWnd.get(), action); };
    callbacks.filterChanged   = [this](std::wstring_view filter)
    {
        _analysisFilter.assign(filter);
        _analysisFilterDueSeconds = NowSeconds() + 0.15;
        InvalidateAnalysis();
    };
    callbacks.panelClose             = [this] { ToggleAnalysisPanel(); };
    callbacks.panelResize            = [this](DxUi::SplitterChange change) { ResizeAnalysisPanel(change); };
    callbacks.inspectorLayoutChanged = [this](bool visible, float ratio)
    {
        _config.inspectorVisible = visible;
        _config.inspectorRatio   = ratio;
        UpdateConfigurationJson();
    };
    callbacks.commandKey    = [this](UINT virtualKey, UINT modifiers) { OnKeyDown(virtualKey, (modifiers & MOD_ALT) != 0u); };
    callbacks.focusBoundary = [this](bool reverse)
    {
        if (reverse && _navigationHost.Focus(true))
        {
            return true;
        }
        if (_hWnd)
        {
            SetFocus(_hWnd.get());
        }
        return true;
    };
    const HRESULT result = _analysisHost.Attach(hwnd, captions, std::move(callbacks));
    if (FAILED(result))
    {
        Debug::Error(L"ViewerSpace analysis host failed: 0x{:08X}", static_cast<unsigned long>(result));
    }
    _analysisHost.SetCompactRows(_config.compactRows);
    _analysisHost.ApplyTheme(_hasTheme ? RedSalamander::MakeThemePaletteFromViewerTheme(_theme) : DxUi::MakeDefaultThemePalette(false));
    _analysisHost.SetInspectorLayout(_config.inspectorVisible, _config.inspectorRatio);
    _analysisHost.SetSortSpec({_analysisSortColumn, DxUi::SortDirection::Descending});
    LayoutAnalysisHost();
}

void ViewerSpace::LayoutAnalysisHost() noexcept
{
    const float width     = DipFromPx(_clientSize.cx);
    const float height    = DipFromPx(_clientSize.cy);
    const float top       = GetHeaderBottomDip();
    const float available = std::max(0.0f, height - top);
    _treemapBounds        = D2D1::RectF(8.0f, top + 8.0f, std::max(8.0f, width - 8.0f), std::max(top + 8.0f, height - 8.0f));
    _analysisBounds       = {};
    _analysisSidePanel    = false;
    _navigationHost.SetAnalysisVisible(_config.analysisPanelVisible);
    _navigationHost.SetMode(_config.viewMode == ViewMode::Analyst);
    if (_config.analysisPanelVisible && width >= 504.0f && available >= 462.0f)
    {
        if (_config.viewMode == ViewMode::Explorer && width >= 1100.0f)
        {
            const float panelWidth = std::clamp(width * _config.analysisPanelRatio, 280.0f, width - 256.0f);
            _analysisBounds        = D2D1::RectF(width - panelWidth - 6.0f, top, width, height);
            _analysisSidePanel     = true;
            _treemapBounds.right   = std::max(_treemapBounds.left, _analysisBounds.left - 8.0f);
        }
        else
        {
            // Keep a 120-DIP drawable treemap plus its 8-DIP top and bottom insets, then size the analysis panel.
            const float panelHeight = std::min(available - 142.0f, std::max(320.0f, available * _config.analysisBottomRatio));
            _analysisBounds         = D2D1::RectF(0.0f, height - panelHeight - 6.0f, width, height);
            _treemapBounds.bottom   = std::max(_treemapBounds.top, _analysisBounds.top - 8.0f);
        }
    }
    if (const HWND child = _analysisHost.GetHwnd())
    {
        const bool visible = _analysisBounds.right > _analysisBounds.left && _analysisBounds.bottom > _analysisBounds.top;
        if (! visible && _analysisHost.OwnsFocus() && _hWnd)
        {
            SetFocus(_hWnd.get());
        }
        ShowWindow(child, visible ? SW_SHOWNOACTIVATE : SW_HIDE);
        if (visible)
        {
            _analysisHost.SetPanelSplitter(
                _analysisSidePanel, _analysisSidePanel ? width : available, _analysisSidePanel ? _analysisBounds.left : _analysisBounds.top - top);
            _analysisHost.SetBounds(
                {PxFromDip(_analysisBounds.left), PxFromDip(_analysisBounds.top), PxFromDip(_analysisBounds.right), PxFromDip(_analysisBounds.bottom)});
        }
    }
}

void ViewerSpace::ResizeAnalysisPanel(DxUi::SplitterChange change) noexcept
{
    const float extent = _analysisSidePanel ? DipFromPx(_clientSize.cx) : DipFromPx(_clientSize.cy) - GetHeaderBottomDip();
    if (extent <= 0.0f)
        return;
    float& ratio = _analysisSidePanel ? _config.analysisPanelRatio : _config.analysisBottomRatio;
    if (change.phase == DxUi::SplitterChangePhase::Preview && ! _analysisResizeStartRatio.has_value())
        _analysisResizeStartRatio = ratio;
    ratio = std::clamp((extent - change.position - 6.0f) / extent, 0.20f, _analysisSidePanel ? 0.65f : 0.80f);
    if (change.phase == DxUi::SplitterChangePhase::Cancel && _analysisResizeStartRatio.has_value())
        ratio = _analysisResizeStartRatio.value();
    if (change.phase != DxUi::SplitterChangePhase::Preview)
    {
        _analysisResizeStartRatio.reset();
        UpdateConfigurationJson();
    }
    LayoutAnalysisHost();
    CancelLayoutBuild();
    _layoutDirty = true;
    InvalidateStaticTreemapCache();
    if (_hWnd)
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void ViewerSpace::InvalidateAnalysis() noexcept
{
    _analysisDirty = true;
    _analysisBuild.reset();
    EnsureAnimationTimer();
}

uint64_t ViewerSpace::AnalysisOwnedBytes() const noexcept
{
    uint64_t bytes    = WorkspaceOwnedBytes() + _analysisRows.capacity() * sizeof(uint32_t) + _analysisFilter.capacity() * sizeof(wchar_t);
    const auto charge = [](const std::shared_ptr<AnalysisBuildState>& state) noexcept -> uint64_t
    {
        if (! state)
        {
            return 0u;
        }
        return sizeof(AnalysisBuildState) + (state->rows.capacity() + state->sort.scratch.capacity()) * sizeof(uint32_t) +
               (state->rowIndex.capacity() + state->rowIndexSort.scratch.capacity()) * sizeof(std::pair<uint32_t, size_t>);
    };
    bytes += charge(_analysisBuild);
    if (_publishedAnalysisOwner != _analysisBuild)
    {
        bytes += charge(_publishedAnalysisOwner);
    }
    return bytes;
}

std::optional<ViewerSpace::ResolvedItem> ViewerSpace::ResolveAnalysisItem(uint32_t id) const noexcept
{
    if (IsSyntheticNodeId(id))
    {
        const uint32_t parentId = id & kTreemapItemIdMask;
        const Node* parent      = TryGetRealNode(parentId);
        if (! parent || (! parent->aggregateFolders && ! parent->aggregateFiles))
        {
            return std::nullopt;
        }
        ResolvedItem item;
        item.id                 = id;
        item.parentId           = parentId;
        item.isSynthetic        = true;
        item.contentNotMeasured = parent->contentNotMeasured;
        item.totalBytes         = parent->aggregateBytes;
        item.aggregateFolders   = parent->aggregateFolders;
        item.aggregateFiles     = parent->aggregateFiles;
        return item;
    }
    return ResolveTreemapItem(id);
}

uint64_t ViewerSpace::AnalysisRowToken(size_t index) const noexcept
{
    if (index >= _analysisRows.size())
    {
        return 0u;
    }
    return (static_cast<uint64_t>(_analysisGeneration) << 32u) | _analysisRows[index];
}

std::optional<size_t> ViewerSpace::FindAnalysisRow(uint64_t token) const noexcept
{
    if (token == 0u || _analysisGeneration != _displayGeneration || static_cast<uint32_t>(token >> 32u) != _analysisGeneration || ! _publishedAnalysisOwner ||
        _publishedAnalysisOwner->generation != _analysisGeneration)
    {
        return std::nullopt;
    }
    const uint32_t id = static_cast<uint32_t>(token);
    const auto& index = _publishedAnalysisOwner->rowIndex;
    const auto found  = std::lower_bound(index.begin(), index.end(), id, [](const auto& entry, uint32_t key) { return entry.first < key; });
    if (found == index.end() || found->first != id || found->second >= _analysisRows.size())
    {
        return std::nullopt;
    }
    return found->second;
}

void ViewerSpace::ContinueAnalysisBuild() noexcept
{
    if (! _config.analysisPanelVisible && _analysisFilter.empty() && ! _typeFilter && ! _ageFilter && ! _exportProjectionRequested)
    {
        // Hidden, unfiltered panels have no consumer for the projection. These arrays hold
        // only trivial IDs/index entries; releasing them requires no per-item model teardown.
        // Keep selection/filter state independent so reopening restores the same inspector.
        const bool hadRows = ! _analysisRows.empty();
        _analysisBuild.reset();
        _publishedAnalysisOwner.reset();
        std::vector<uint32_t>().swap(_analysisRows);
        _analysisDirty = false;
        if (hadRows)
        {
            _analysisGeneration = 0u;
            _analysisHost.NotifyRowsChanged(std::nullopt);
        }
        return;
    }
    if (_scanCacheRestoreSnapshot)
    {
        _analysisBuild.reset();
        _analysisDirty = true;
        return;
    }
    if (! _analysisDirty && ! _analysisBuild)
    {
        return;
    }
    const double now = NowSeconds();
    if (_analysisFilterDueSeconds > now ||
        (! _analysisBuild && (! CanBeginScanProjection(! _analysisRows.empty()) || (_scanActive.load() && now - _analysisLastPublishedSeconds < 0.25))))
    {
        return;
    }
    const auto started = std::chrono::steady_clock::now();
#ifdef ENABLE_TESTS
    const bool capturePublishPhases = _debugPublicationPhaseTiming && Debug::Perf::IsCaptureEnabled();
#else
    const bool capturePublishPhases = false;
#endif
    bool published               = false;
    uint64_t notifyRowsChangedUs = 0u;
    uint64_t inspectorUpdateUs   = 0u;
    uint64_t publishElapsedUs    = 0u;
    auto publishStarted          = std::chrono::steady_clock::time_point{};
    auto phaseStarted            = std::chrono::steady_clock::time_point{};
    const uint32_t generation    = _displayGeneration;
    if (_analysisBuild &&
        (_analysisBuild->generation != generation || _analysisBuild->revision != _modelRevision || _analysisBuild->view != _model->viewNodeId))
    {
        _analysisBuild.reset();
    }
    if (! _analysisBuild)
    {
        _analysisBuild             = std::make_shared<AnalysisBuildState>();
        _analysisBuild->generation = generation;
        _analysisBuild->revision   = _modelRevision;
        _analysisBuild->view       = _model->viewNodeId;
    }
    auto& build         = *_analysisBuild;
    const Node* parent  = TryGetRealNode(build.view);
    const auto children = _reviewingCleanup  ? std::span<const uint32_t>(_cleanupMarks)
                          : _showingProblems ? std::span<const uint32_t>(_problemRows)
                          : parent           ? GetRealNodeChildren(*parent)
                                             : std::span<const uint32_t>{};
    auto less           = [this](uint32_t left, uint32_t right) noexcept
    {
        const auto l = ResolveAnalysisItem(left);
        const auto r = ResolveAnalysisItem(right);
        if (! l.has_value() || ! r.has_value())
        {
            return left < right;
        }
        const auto& a         = l.value();
        const auto& b         = r.value();
        int primaryComparison = 0;
        if (_analysisSortColumn == 2u && MapWeight(a.id, a.totalBytes) != MapWeight(b.id, b.totalBytes))
        {
            primaryComparison = MapWeight(a.id, a.totalBytes) < MapWeight(b.id, b.totalBytes) ? -1 : 1;
        }
        else if (_analysisSortColumn == 1u && a.totalBytes != b.totalBytes)
        {
            primaryComparison = a.totalBytes < b.totalBytes ? -1 : 1;
        }
        else if (_analysisSortColumn == 3u && a.isDirectory != b.isDirectory)
        {
            primaryComparison = a.isDirectory ? -1 : 1;
        }
        else if (_analysisSortColumn == 4u && a.scanState != b.scanState)
        {
            primaryComparison = a.scanState < b.scanState ? -1 : 1;
        }
        else if (_analysisSortColumn == 0u)
        {
            if (NaturalSort::Less(a.name, b.name))
            {
                primaryComparison = -1;
            }
            else if (NaturalSort::Less(b.name, a.name))
            {
                primaryComparison = 1;
            }
        }
        if (primaryComparison != 0)
        {
            return _analysisSortAscending ? primaryComparison < 0 : primaryComparison > 0;
        }

        // Ties keep a deterministic ascending logical-name order, including descending byte sorts.
        if (NaturalSort::Less(a.name, b.name))
        {
            return true;
        }
        if (NaturalSort::Less(b.name, a.name))
        {
            return false;
        }
        return left < right;
    };
    size_t work = 0u;
    while (work++ < 32768u)
    {
        if ((work & 63u) == 0u && Debug::Perf::ElapsedUs(started) >= 1800u)
        {
            break;
        }
        if (! build.sorting)
        {
            if (build.cursor < children.size())
            {
                const uint32_t id = children[build.cursor++];
                const auto item   = ResolveAnalysisItem(id);
                if (! item.has_value())
                {
                    continue;
                }
                if (! _reviewingCleanup && ! _showingProblems)
                {
                    const uint32_t category   = item.value().isDirectory ? 7u : ItemCategory(item.value());
                    build.typeBytes[category] = ViewerSpaceScan::SaturatingAdd(build.typeBytes[category], item.value().totalBytes);
                    ++build.typeCounts[category];
                }
                const auto name = item.value().name;
                const bool match =
                    _reviewingCleanup || _showingProblems ||
                    ((_analysisFilter.empty() ||
                      FindStringOrdinal(
                          FIND_FROMSTART, name.data(), static_cast<int>(name.size()), _analysisFilter.data(), static_cast<int>(_analysisFilter.size()), TRUE) >=
                          0) &&
                     MatchesWorkspaceFilter(item.value()));
                if (match)
                {
                    build.rows.push_back(id);
                }
                continue;
            }
            if (! _reviewingCleanup && ! _showingProblems && parent && (parent->aggregateFolders || parent->aggregateFiles))
            {
                // Unknown names remain explicit Other even when the retained-name filter is active.
                build.rows.push_back(SyntheticOtherBucketIdForParent(parent->id));
                build.typeBytes[6u]  = ViewerSpaceScan::SaturatingAdd(build.typeBytes[6u], parent->aggregateBytes);
                build.typeCounts[6u] = ViewerSpaceScan::SaturatingAdd(build.typeCounts[6u], parent->aggregateFiles + parent->aggregateFolders);
            }
            build.sorting = true;
        }
        if (! build.sort.Step(build.rows, less))
        {
            continue;
        }
        if (! build.indexing)
        {
            build.indexing       = true;
            build.rowIndexCursor = 0u;
        }
        while (build.rowIndexCursor < build.rows.size() && work < 32768u)
        {
            const size_t displayIndex = build.rowIndexCursor++;
            build.rowIndex.emplace_back(build.rows[displayIndex], displayIndex);
            ++work;
            if ((work & 63u) == 0u && Debug::Perf::ElapsedUs(started) >= 1800u)
            {
                break;
            }
        }
        if (build.rowIndexCursor < build.rows.size())
        {
            continue;
        }
        const auto indexLess = [](const auto& left, const auto& right) noexcept
        { return left.first < right.first || (left.first == right.first && left.second < right.second); };
        if (! build.rowIndexSort.Step(build.rowIndex, indexLess))
        {
            continue;
        }
        if (capturePublishPhases)
        {
            publishStarted = std::chrono::steady_clock::now();
        }
        if (! _reviewingCleanup && ! _showingProblems)
        {
            _typeBreakdownBytes  = build.typeBytes;
            _typeBreakdownCounts = build.typeCounts;
        }
        _analysisRows.swap(build.rows);
        _analysisGeneration = generation;
        _analysisViewId     = build.view;
        // The published owner retains only the ID index. Release both sort workspaces and the
        // prior row vector so publication does not pin two full generations of temporary arrays.
        std::vector<uint32_t>().swap(build.rows);
        std::vector<uint32_t>().swap(build.sort.scratch);
        std::vector<std::pair<uint32_t, size_t>>().swap(build.rowIndexSort.scratch);
        _publishedAnalysisOwner = _analysisBuild;
        _analysisBuild.reset();
        _analysisDirty                = false;
        _analysisLastPublishedSeconds = now;
        const std::optional<uint64_t> selected =
            _selectedItemId ? std::optional<uint64_t>((static_cast<uint64_t>(generation) << 32u) | _selectedItemId) : std::nullopt;
        if (capturePublishPhases)
        {
            phaseStarted = std::chrono::steady_clock::now();
        }
        _analysisHost.NotifyRowsChanged(selected);
        if (capturePublishPhases)
        {
            notifyRowsChangedUs = Debug::Perf::ElapsedUs(phaseStarted);
            phaseStarted        = std::chrono::steady_clock::now();
        }
        UpdateInspector();
        if (capturePublishPhases)
        {
            inspectorUpdateUs = Debug::Perf::ElapsedUs(phaseStarted);
        }
        Debug::Perf::EmitValue(L"viewer.space.analysis.retained_rows", _analysisRows.size());
        if (capturePublishPhases)
        {
            publishElapsedUs = Debug::Perf::ElapsedUs(publishStarted);
        }
        published = true;
        break;
    }
    if (capturePublishPhases && published)
    {
        // Endpoints precede phase rows; analysis.turn_us below includes their observer-write cost.
        Debug::Perf::Emit(L"viewer.space.analysis.publish_phase_us", L"notify-rows-changed", notifyRowsChangedUs, _analysisRows.size(), generation, S_OK);
        Debug::Perf::Emit(L"viewer.space.analysis.publish_phase_us", L"inspector-update", inspectorUpdateUs, _analysisRows.size(), generation, S_OK);
        Debug::Perf::Emit(L"viewer.space.analysis.publish_us", L"commit-notify-inspector", publishElapsedUs, _analysisRows.size(), generation, S_OK);
    }
    const uint64_t elapsed = Debug::Perf::ElapsedUs(started);
    _analysisMaxTurnUs     = std::max(_analysisMaxTurnUs, elapsed);
    Debug::Perf::Emit(L"viewer.space.analysis.turn_us", L"projection", elapsed, work, _analysisRows.size(), S_OK);
}

void ViewerSpace::GetAnalysisCell(size_t row, size_t column, DxUi::GridCellData& cell) const
{
    cell = {};
    if (row >= _analysisRows.size())
    {
        return;
    }
    const auto item = ResolveAnalysisItem(_analysisRows[row]);
    if (! item.has_value())
    {
        return;
    }
    const auto& value = item.value();
    switch (column)
    {
        case 0u:
            cell.kind         = DxUi::GridCellKind::IconText;
            cell.iconText     = value.isSynthetic ? L"•" : value.isDirectory ? L"\uE8B7" : L"\uE8A5";
            cell.iconFontRole = value.isSynthetic ? DxUi::FontRole::Body : DxUi::FontRole::Icon;
            cell.enabled      = ! IsItemExcluded(value.id);
            cell.badgeText    = IsItemExcluded(value.id) ? LoadStringResource(g_hInstance, IDS_VIEWERSPACE_EXCLUDED) : std::wstring{};
            cell.tooltipText  = value.isSynthetic
                                    ? LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_DETAIL_LIMITATION)
                                    : LoadStringResource(g_hInstance, IsItemExcluded(value.id) ? IDS_VIEWERSPACE_RESTORE_ITEM : IDS_VIEWERSPACE_EXCLUDE);
            if (! IsItemExcluded(value.id))
            {
                if (value.scanState == ScanState::Error)
                    cell.badgeText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_STATE_UNREADABLE);
                else if (value.contentNotMeasured)
                    cell.badgeText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_UNMEASURED_SHORT);
                else if (value.isSynthetic || _showingProblems)
                    cell.badgeText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_SUMMARIZED_SHORT);
                else if (value.scanState != ScanState::Done)
                    cell.badgeText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_STATE_PARTIAL);
            }
            cell.text = value.isSynthetic                       ? LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_KIND_OTHER)
                        : _reviewingCleanup || _showingProblems ? BuildItemPathText(value.id)
                                                                : std::wstring(value.name);
            break;
        case 1u:
        {
            cell.text = value.scanState == ScanState::Error || value.contentNotMeasured
                            ? LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_STATE_UNAVAILABLE)
                            : FormatBytesCompact(value.totalBytes);
            if (! IsItemExcluded(value.id) && ShownBytes(value.id, value.totalBytes) < value.totalBytes)
            {
                const auto excluded = _excludedBelow.find(value.id);
                cell.tooltipText    = FormatStringResource(g_hInstance,
                                                           IDS_VIEWERSPACE_EXCLUSION_SUMMARY,
                                                           FormatBytesCompact(ShownBytes(value.id, value.totalBytes)),
                                                           excluded == _excludedBelow.end() ? 0u : excluded->second.size());
            }
            break;
        }
        case 2u:
        {
            const Node* parent = TryGetRealNode(_analysisViewId);
            cell.text          = LoadEmbeddedStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_SHARE_UNKNOWN);
            if (parent && MapWeight(parent->id, parent->totalBytes) && value.scanState != ScanState::Error && ! value.contentNotMeasured &&
                ! (_config.fileCountWeight && value.isSynthetic))
            {
                cell.text = FormatStringResource(g_hInstance,
                                                 IDS_VIEWERSPACE_TOOLTIP_PERCENT_FORMAT,
                                                 100.0 * static_cast<double>(MapWeight(value.id, value.totalBytes)) /
                                                     static_cast<double>(MapWeight(parent->id, parent->totalBytes)));
            }
            break;
        }
        case 3u:
            cell.kind    = DxUi::GridCellKind::Checkbox;
            cell.checked = _cleanupGeneration == _displayGeneration && std::find(_cleanupMarks.begin(), _cleanupMarks.end(), value.id) != _cleanupMarks.end();
            cell.enabled = ! value.isSynthetic && (cell.checked || ! IsItemExcluded(value.id));
            cell.tooltipText = LoadStringResource(g_hInstance, cell.checked ? IDS_VIEWERSPACE_UNMARK_CLEANUP : IDS_VIEWERSPACE_MARK_CLEANUP);
            break;
        case 4u:
        {
            UINT state = IDS_VIEWERSPACE_STATUS_DONE;
            if (value.contentNotMeasured)
            {
                state = IDS_VIEWERSPACE_ANALYSIS_CONTENTS_NOT_MEASURED;
            }
            else if (value.scanState == ScanState::Error)
            {
                state = IDS_VIEWERSPACE_ANALYSIS_STATE_UNREADABLE;
            }
            else if (value.scanState == ScanState::Queued || value.scanState == ScanState::NotStarted)
            {
                state = IDS_VIEWERSPACE_STATUS_QUEUED;
            }
            else if (value.scanState == ScanState::Scanning)
            {
                state = IDS_VIEWERSPACE_STATUS_SCANNING;
            }
            else if (value.scanState == ScanState::Canceled)
            {
                state = IDS_VIEWERSPACE_STATUS_CANCELED;
            }
            cell.text = LoadStringResource(g_hInstance, state);
            break;
        }
        default: break;
    }
}

void ViewerSpace::SelectItem(uint32_t id, bool syncGrid) noexcept
{
    _selectedItemId        = id;
    _selectedGeneration    = _displayGeneration;
    _selectionFromAnalysis = ! syncGrid;
    UpdateInspector();
    if (syncGrid)
    {
        const uint64_t token = (static_cast<uint64_t>(_analysisGeneration) << 32u) | id;
        _analysisHost.NotifyRowsChanged(id && FindAnalysisRow(token).has_value() ? std::optional<uint64_t>(token) : std::nullopt);
    }
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool ViewerSpace::AnalysisFilterMatches(uint32_t id) const noexcept
{
    // The published table's compact ID index is the filter result. Paint never rescans names.
    return ! _reviewingCleanup && ! _showingProblems && (! _analysisFilter.empty() || _typeFilter || _ageFilter) && _analysisGeneration == _displayGeneration &&
           FindAnalysisRow((static_cast<uint64_t>(_analysisGeneration) << 32u) | id).has_value();
}

void ViewerSpace::UpdateInspector() noexcept
{
    const auto item = _selectionFromAnalysis ? ResolveAnalysisItem(_selectedItemId) : ResolveTreemapItem(_selectedItemId);
    ViewerSpaceAnalysis::AnalysisHost::InspectorText text;
    if (item.has_value())
    {
        const auto& value = item.value();
        text.name         = value.isSynthetic ? LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_KIND_OTHER) : std::wstring(value.name);
        text.path         = BuildNodePathText(value.parentId);
        text.fullPath     = value.isSynthetic ? text.path : BuildItemPathText(value.id);
        text.icon         = value.isSynthetic ? L"\uE712" : value.isDirectory ? L"\uE8B7" : L"\uE8A5";
        text.kind         = LoadStringResource(g_hInstance,
                                               static_cast<UINT>(value.isSynthetic   ? IDS_VIEWERSPACE_ANALYSIS_KIND_OTHER
                                                                 : value.isDirectory ? IDS_VIEWERSPACE_ANALYSIS_KIND_FOLDER
                                                                                     : IDS_VIEWERSPACE_ANALYSIS_KIND_FILE));
        text.logicalBytes = FormatBytesCompact(value.totalBytes);
        text.exactBytes   = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_HEADER_SIZE_FORMAT, text.logicalBytes, value.totalBytes);
        text.modified = Common::FileMetadata::FormatDisplayFields({value.lastWriteTime, 0u}, Common::FileMetadata::DisplayProfile::CompactDetails).localTime;
        if (text.modified.empty())
        {
            text.modified = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_STATE_UNAVAILABLE);
        }
        const Node* parent = TryGetRealNode(value.parentId);
        text.share         = LoadEmbeddedStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_SHARE_UNKNOWN);
        if (parent && MapWeight(parent->id, parent->totalBytes) && value.scanState != ScanState::Error && ! value.contentNotMeasured &&
            ! (_config.fileCountWeight && value.isSynthetic))
        {
            text.share = FormatStringResource(g_hInstance,
                                              IDS_VIEWERSPACE_TOOLTIP_PERCENT_FORMAT,
                                              100.0 * static_cast<double>(MapWeight(value.id, value.totalBytes)) /
                                                  static_cast<double>(MapWeight(parent->id, parent->totalBytes)));
        }
        if (value.scanState == ScanState::Error || value.contentNotMeasured)
            text.logicalBytes = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_STATE_UNAVAILABLE);
        text.state = value.scanState == ScanState::Done && ! value.isSynthetic && ! value.contentNotMeasured
                         ? std::wstring{}
                         : LoadStringResource(g_hInstance,
                                              static_cast<UINT>(value.contentNotMeasured              ? IDS_VIEWERSPACE_ANALYSIS_CONTENTS_NOT_MEASURED
                                                                : value.isSynthetic                   ? IDS_VIEWERSPACE_ANALYSIS_DETAIL_LIMITATION
                                                                : value.scanState == ScanState::Error ? IDS_VIEWERSPACE_ANALYSIS_STATE_UNREADABLE
                                                                : value.scanState == ScanState::Done  ? IDS_VIEWERSPACE_STATUS_DONE
                                                                                                      : IDS_VIEWERSPACE_ANALYSIS_STATE_PARTIAL));
    }
    else
    {
        text.state = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_ANALYSIS_SELECTION_HELP);
    }
    if (item.has_value() && IsItemExcluded(item.value().id))
        text.state = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_EXCLUDED_STATE);
    else if (item.has_value() && ShownBytes(item.value().id, item.value().totalBytes) < item.value().totalBytes)
    {
        if (! text.state.empty())
            text.state.append(L" · ");
        const auto excluded = _excludedBelow.find(item.value().id);
        text.state.append(FormatStringResource(g_hInstance,
                                               IDS_VIEWERSPACE_EXCLUSION_SUMMARY,
                                               FormatBytesCompact(ShownBytes(item.value().id, item.value().totalBytes)),
                                               excluded == _excludedBelow.end() ? 0u : excluded->second.size()));
    }
    _analysisHost.SetInspector(text);
}

void ViewerSpace::ActivateItem(uint32_t id) noexcept
{
    const auto item = ResolveTreemapItem(id);
    if (! item.has_value())
    {
        return;
    }
    if (item.value().isDirectory)
    {
        NavigateTo(id);
    }
    else if (! item.value().isSynthetic)
    {
        FocusItemInPane(id);
    }
}

void ViewerSpace::FocusItemInPane(uint32_t id) noexcept
{
    const auto item = ResolveTreemapItem(id);
    if (! item.has_value() || item.value().isSynthetic)
    {
        return;
    }
    wil::com_ptr<IViewer> keepAlive = static_cast<IViewer*>(this);
    const bool hasParent            = item.value().parentId != 0u;
    const std::wstring folder       = hasParent ? BuildNodePathText(item.value().parentId) : BuildItemPathText(id);
    const std::wstring name         = hasParent ? std::wstring(item.value().name) : std::wstring{};
    if (folder.empty())
        return;
    if (_hostViewerActions)
    {
        const wchar_t* path = folder.c_str();
        HostViewerActionRequest action{};
        action.sizeBytes            = sizeof(action);
        action.action               = HOST_VIEWER_OPEN_FOCUS_PANE;
        action.fileSystem           = _fileSystem.get();
        action.paths                = &path;
        action.pathCount            = 1u;
        action.focusItemDisplayName = name.empty() ? nullptr : name.c_str();
        if (FAILED(_hostViewerActions->ExecuteViewerAction(&action, nullptr)))
            ShowWorkspaceMessage(IDS_VIEWERSPACE_TARGET_UNAVAILABLE);
        return;
    }
    // The legacy service has no provider identity; only local paths are safe to pass to it.
    if (_fileSystemShortId != L"file" || ! _hostPaneExecute)
        return;
    HostPaneExecuteRequest request{};
    request.sizeBytes            = sizeof(request);
    request.flags                = HOST_PANE_EXECUTE_FLAG_ACTIVATE_WINDOW;
    request.folderPath           = folder.c_str();
    request.focusItemDisplayName = name.empty() ? nullptr : name.c_str();
    static_cast<void>(_hostPaneExecute->ExecuteInActivePane(&request));
}

void ViewerSpace::SetViewMode(ViewMode mode) noexcept
{
    _analysisHost.CancelResize();
    _config.viewMode             = mode;
    _config.analysisPanelVisible = true;
    _navigationHost.SetMode(mode == ViewMode::Analyst);
    UpdateConfigurationJson();
    LayoutAnalysisHost();
    InvalidateAnalysis();
    CancelLayoutBuild();
    _layoutDirty = true;
    InvalidateStaticTreemapCache();
    if (_hWnd)
    {
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::ToggleAnalysisPanel() noexcept
{
    _analysisHost.CancelResize();
    _config.analysisPanelVisible = ! _config.analysisPanelVisible;
    UpdateConfigurationJson();
    LayoutAnalysisHost();
    InvalidateAnalysis();
    CancelLayoutBuild();
    _layoutDirty = true;
    InvalidateStaticTreemapCache();
    if (_hWnd)
    {
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::NavigateHistory(bool forward) noexcept
{
    CancelDeferredStartScan();
    const auto& initialSource = forward ? _navForwardStack : _navStack;
    if (initialSource.empty())
    {
        return;
    }
    const uint32_t id = initialSource.back();
    if (! TryGetRealNode(id))
    {
        return;
    }
    if (_workingModel != nullptr)
    {
        DiscardWorkingModel(false);
        _refreshFailed             = false;
        _refreshRefused            = false;
        const auto& restoredSource = forward ? _navForwardStack : _navStack;
        if (restoredSource.empty() || restoredSource.back() != id || ! TryGetRealNode(id))
        {
            UpdateNavigationHostState();
            if (_hWnd)
            {
                UpdateMenuState(_hWnd.get());
            }
            return;
        }
    }
    auto& source = forward ? _navForwardStack : _navStack;
    auto& target = forward ? _navStack : _navForwardStack;
    if (source.empty() || source.back() != id || ! TryGetRealNode(id))
    {
        return;
    }
    _refreshFailed  = false;
    _refreshRefused = false;
    source.pop_back();
    if (target.size() >= 64u)
    {
        target.erase(target.begin());
    }
    target.push_back(_model->viewNodeId);
    _model->viewNodeId                = id;
    _firstPositiveProjectionAttempted = false;
    _hasPublishedPositiveGeometry     = false;
    UpdateViewPathText();
    CancelLayoutBuild();
    _layoutDirty = true;
    InvalidateAnalysis();
    InvalidateStaticTreemapCache();
    if (_hWnd)
    {
        UpdateWindowTitle(_hWnd.get());
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::UpdateConfigurationJson()
{
    _configurationJson =
        std::format("{{\"topFilesPerDirectory\":{},\"scanThreads\":{},\"maxConcurrentScansPerVolume\":{},\"cacheEnabled\":\"{}\","
                    "\"cacheTtlSeconds\":{},\"cacheMaxEntries\":{},\"viewMode\":\"{}\",\"analysisPanelVisible\":{},\"analysisPanelRatio\":{},"
                    "\"analysisBottomRatio\":{},\"inspectorVisible\":{},\"inspectorRatio\":{},\"compactRows\":{},\"fileCountWeight\":{},\"colorMode\":\"{}\"}}",
                    _config.topFilesPerDirectory,
                    _config.scanThreads,
                    _config.maxConcurrentScansPerVolume,
                    _config.cacheEnabled ? "1" : "0",
                    _config.cacheTtlSeconds,
                    _config.cacheMaxEntries,
                    _config.viewMode == ViewMode::Analyst ? "analyst" : "explorer",
                    _config.analysisPanelVisible,
                    _config.analysisPanelRatio,
                    _config.analysisBottomRatio,
                    _config.inspectorVisible,
                    _config.inspectorRatio,
                    _config.compactRows,
                    _config.fileCountWeight,
                    std::array<const char*, 4u>{{"folder", "type", "age", "newest"}}[_config.colorMode]);
}
