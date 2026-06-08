#include "FolderViewInternal.h"

void FolderView::UpdateEstimatedMetrics()
{
    // Compute estimated character width and height from actual font metrics
    // This ensures estimates are accurate across different DPI settings
    if (_estimatedMetricsValid || ! _dwriteFactory || ! _labelFormat)
    {
        return;
    }

    // Measure a representative sample string to get average character width
    // Using alphanumeric chars that represent typical filename characters
    constexpr wchar_t kSampleText[] = L"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    constexpr size_t kSampleLength  = std::size(kSampleText) - 1; // Exclude null terminator

    wil::com_ptr<IDWriteTextLayout> sampleLayout;
    HRESULT hr = _dwriteFactory->CreateTextLayout(kSampleText,
                                                  static_cast<UINT32>(kSampleLength),
                                                  _labelFormat.get(),
                                                  10000.0f, // Large max width
                                                  1000.0f,  // Large max height
                                                  sampleLayout.addressof());

    if (SUCCEEDED(hr) && sampleLayout)
    {
        DWRITE_TEXT_METRICS metrics{};
        if (SUCCEEDED(sampleLayout->GetMetrics(&metrics)))
        {
            // Average width per character
            _estimatedCharWidthDip   = metrics.widthIncludingTrailingWhitespace / static_cast<float>(kSampleLength);
            _estimatedLabelHeightDip = metrics.height;

            Debug::Info(L"FolderView: Updated estimated metrics - charWidth={:.2f}, labelHeight={:.2f} (DPI={:.0f})",
                        _estimatedCharWidthDip,
                        _estimatedLabelHeightDip,
                        _dpi);
        }
    }

    // Also measure details format if available
    if (_detailsFormat)
    {
        wil::com_ptr<IDWriteTextLayout> detailsLayout;
        hr = _dwriteFactory->CreateTextLayout(
            kSampleText, static_cast<UINT32>(kSampleLength), _detailsFormat.get(), 10000.0f, 1000.0f, detailsLayout.addressof());

        if (SUCCEEDED(hr) && detailsLayout)
        {
            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(detailsLayout->GetMetrics(&metrics)))
            {
                _estimatedDetailsHeightDip  = metrics.height;
                _estimatedMetadataHeightDip = metrics.height;
            }
        }
    }

    _estimatedMetricsValid = true;
}

void FolderView::LayoutItems()
{
    Debug::Perf::Scope layoutPerf(L"render.layout_items_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(_items.size()));

    EnsureDeviceIndependentResources();

    // Ensure estimated metrics are computed from actual font (DPI-aware)
    UpdateEstimatedMetrics();

    const float clientWidthDip  = std::max(0.0f, DipFromPx(_clientSize.cx));
    const float clientHeightDip = std::max(0.0f, DipFromPx(_clientSize.cy));
    const bool includeDetailsLine =
        _displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    const bool includeMetadataLine = _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;

    _columnCounts.clear();
    _columnPrefixSums.clear();
    _columnLayout.clear();

    if (_items.empty() || clientWidthDip <= 0.0f)
    {
        if (_items.empty())
        {
            _tileWidthDip   = 0.0f;
            _tileHeightDip  = 0.0f;
            _labelHeightDip = 0.0f;

            if (clientWidthDip > 0.0f)
            {
                std::wstring parentRowLabel = LoadStringResource(nullptr, IDS_EMPTY_FOLDER_PARENT_ROW);
                if (parentRowLabel.empty())
                {
                    parentRowLabel = LoadStringResource(nullptr, IDS_EMPTY_FOLDER_TITLE);
                }

                const FolderViewEmptyStateLayout::PlaceholderItemMetrics metrics =
                    FolderViewEmptyStateLayout::ResolvePlaceholderItemMetrics(FolderViewEmptyStateLayout::PlaceholderItemMetricsInput{
                        .clientWidthDip          = clientWidthDip,
                        .clientHeightDip         = clientHeightDip,
                        .iconSizeDip             = _iconSizeDip,
                        .estimatedCharWidthDip   = _estimatedCharWidthDip,
                        .estimatedLabelHeightDip = _estimatedLabelHeightDip,
                        .detailsLineHeightDip    = _detailsLineHeightDip > 0.0f ? _detailsLineHeightDip : _estimatedDetailsHeightDip,
                        .metadataLineHeightDip   = _metadataLineHeightDip > 0.0f ? _metadataLineHeightDip : _estimatedMetadataHeightDip,
                        .titleLength             = parentRowLabel.size(),
                        .includeDetailsLine      = includeDetailsLine,
                        .includeMetadataLine     = includeMetadataLine && static_cast<bool>(_metadataTextProvider),
                    });

                _tileWidthDip   = metrics.tileWidthDip;
                _tileHeightDip  = metrics.tileHeightDip;
                _labelHeightDip = metrics.labelHeightDip;
            }
        }

        _columns          = 1;
        _rowsPerColumn    = 0;
        _contentHeight    = std::max(clientHeightDip, 0.0f);
        _contentWidth     = std::max(clientWidthDip, 0.0f);
        _horizontalOffset = 0.0f;
        layoutPerf.SetValue1(0);
        return;
    }

    float maxLabelWidth    = 0.0f;
    float maxLabelHeight   = 0.0f;
    float maxDetailsWidth  = 0.0f;
    float maxMetadataWidth = 0.0f;

    // Use estimated metrics for initial layout to avoid blocking UI thread
    // Text layouts are created lazily when items are rendered (visible items only)
    if (! _itemMetricsCached)
    {
        TRACER_CTX(L"EstimateMetrics");

        if (includeDetailsLine)
        {
            size_t sizeSlotChars = 0;
            for (const auto& item : _items)
            {
                if (item.isDirectory)
                {
                    continue;
                }
                const std::wstring sizeText = FormatBytesCompact(item.sizeBytes);
                sizeSlotChars               = std::max(sizeSlotChars, sizeText.size());
            }

            if (sizeSlotChars == 0)
            {
                const std::wstring sizeText = FormatBytesCompact(0);
                sizeSlotChars               = sizeText.size();
            }

            constexpr size_t kMaxSizeSlotChars = 12;
            _detailsSizeSlotChars              = std::min(sizeSlotChars, kMaxSizeSlotChars);
        }
        else
        {
            _detailsSizeSlotChars = 0;
        }

        // Use estimated metrics based on character count instead of creating layouts
        // This avoids O(N) DirectWrite calls for large directories
        for (auto& item : _items)
        {
            if (item.displayName.empty())
            {
                continue;
            }

            // Estimate label width based on character count
            const float estimatedWidth                         = static_cast<float>(item.displayName.length()) * _estimatedCharWidthDip;
            item.labelMetrics.width                            = estimatedWidth;
            item.labelMetrics.widthIncludingTrailingWhitespace = estimatedWidth;
            item.labelMetrics.height                           = _estimatedLabelHeightDip;

            maxLabelWidth  = std::max(maxLabelWidth, estimatedWidth);
            maxLabelHeight = std::max(maxLabelHeight, _estimatedLabelHeightDip);

            // Clear any existing layout - will be created lazily on render
            item.labelLayout.reset();

            if (includeDetailsLine)
            {
                if (item.detailsText.empty())
                {
                    if (_detailsTextProvider)
                    {
                        item.detailsText =
                            _detailsTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
                    }
                    else
                    {
                        item.detailsText = BuildDetailsText(item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes, _detailsSizeSlotChars);
                    }
                }

                // Estimate details width
                const float estimatedDetailsWidth                    = static_cast<float>(item.detailsText.length()) * _estimatedCharWidthDip * 0.85f;
                item.detailsMetrics.width                            = estimatedDetailsWidth;
                item.detailsMetrics.widthIncludingTrailingWhitespace = estimatedDetailsWidth;
                item.detailsMetrics.height                           = _estimatedDetailsHeightDip;

                maxDetailsWidth = std::max(maxDetailsWidth, estimatedDetailsWidth);

                // Clear any existing layout - will be created lazily on render
                item.detailsLayout.reset();

                if (includeMetadataLine)
                {
                    if (item.metadataText.empty() && _metadataTextProvider)
                    {
                        item.metadataText =
                            _metadataTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
                    }

                    const float estimatedMetadataWidth                    = static_cast<float>(item.metadataText.length()) * _estimatedCharWidthDip * 0.85f;
                    item.metadataMetrics.width                            = estimatedMetadataWidth;
                    item.metadataMetrics.widthIncludingTrailingWhitespace = estimatedMetadataWidth;
                    item.metadataMetrics.height                           = _estimatedMetadataHeightDip;
                    maxMetadataWidth                                      = std::max(maxMetadataWidth, estimatedMetadataWidth);

                    item.metadataLayout.reset();
                }
                else
                {
                    item.metadataLayout.reset();
                    item.metadataMetrics = {};
                }
            }
        }

        _cachedMaxLabelWidth    = maxLabelWidth;
        _cachedMaxLabelHeight   = maxLabelHeight;
        _cachedMaxDetailsWidth  = maxDetailsWidth;
        _cachedMaxMetadataWidth = maxMetadataWidth;
        _itemMetricsCached      = true;

        Debug::Info(L"FolderView::LayoutItems estimated {} items, max width={:.1f}, max height={:.1f}", _items.size(), maxLabelWidth, maxLabelHeight);
    }
    else
    {
        // Reuse cached measurements
        maxLabelWidth    = _cachedMaxLabelWidth;
        maxLabelHeight   = _cachedMaxLabelHeight;
        maxDetailsWidth  = _cachedMaxDetailsWidth;
        maxMetadataWidth = _cachedMaxMetadataWidth;
    }

    if (maxLabelHeight <= 0.0f)
    {
        maxLabelHeight = 16.0f;
    }

    const float textWidthSafety = std::max(_estimatedCharWidthDip, 8.0f);

    _labelHeightDip = maxLabelHeight + kLabelVerticalPaddingDip * 2.0f;
    if (includeDetailsLine)
    {
        const float detailsHeight = _detailsLineHeightDip > 0.0f ? _detailsLineHeightDip : 12.0f;
        float textBlockHeight     = maxLabelHeight + kDetailsGapDip + detailsHeight;
        if (includeMetadataLine && _metadataTextProvider && maxMetadataWidth > 0.0f)
        {
            const float metadataHeight = _metadataLineHeightDip > 0.0f ? _metadataLineHeightDip : detailsHeight;
            textBlockHeight += kDetailsGapDip + metadataHeight;
        }
        _tileHeightDip = std::max(_iconSizeDip, textBlockHeight) + kLabelVerticalPaddingDip * 2.0f;
    }
    else
    {
        _tileHeightDip = std::max(_iconSizeDip, maxLabelHeight) + kLabelVerticalPaddingDip * 2.0f;
    }

    const float rowSpacingDip = GetFolderViewRowSpacingDip(_appTheme);
    std::vector<FolderViewColumnLayout::ItemTextMetrics> columnTextMetrics;
    columnTextMetrics.reserve(_items.size());
    for (const auto& item : _items)
    {
        const std::wstring_view labelText = GetVisualDisplayName(item);
        FolderViewColumnLayout::ItemTextMetrics metrics{
            .labelWidthDip = static_cast<float>(labelText.size()) * _estimatedCharWidthDip,
        };

        if (includeDetailsLine)
        {
            metrics.detailsWidthDip = static_cast<float>(item.detailsText.size()) * _estimatedCharWidthDip * 0.85f;
        }
        if (includeMetadataLine)
        {
            metrics.metadataWidthDip = static_cast<float>(item.metadataText.size()) * _estimatedCharWidthDip * 0.85f;
        }

        columnTextMetrics.push_back(metrics);
    }

    FolderViewColumnLayout::Result columnLayout = FolderViewColumnLayout::Resolve(FolderViewColumnLayout::Input{
        .clientWidthDip       = std::max(1.0f, clientWidthDip),
        .clientHeightDip      = clientHeightDip,
        .tileHeightDip        = _tileHeightDip,
        .rowSpacingDip        = rowSpacingDip,
        .iconSizeDip          = _iconSizeDip,
        .iconTextGapDip       = kIconTextGapDip,
        .horizontalPaddingDip = kLabelHorizontalPaddingDip * 2.0f,
        .columnSpacingDip     = kColumnSpacingDip,
        .textWidthSafetyDip   = textWidthSafety,
        .includeDetailsLine   = includeDetailsLine,
        .includeMetadataLine  = includeMetadataLine,
        .items                = columnTextMetrics,
    });

    _columnLayout  = std::move(columnLayout.columns);
    _tileWidthDip  = columnLayout.maxColumnWidthDip;
    _rowsPerColumn = std::max(1, columnLayout.rowsPerColumn);
    _columns       = static_cast<int>(_columnLayout.size());
    if (_columns < 1)
    {
        _columns = 1;
    }

    _columnCounts.reserve(_columnLayout.size());
    for (const auto& column : _columnLayout)
    {
        _columnCounts.push_back(static_cast<int>(column.itemCount));
    }

    // Build prefix sums for O(1) keyboard navigation: _columnPrefixSums[c] = items before column c
    _columnPrefixSums.clear();
    _columnPrefixSums.reserve(_columnCounts.size() + 1);
    size_t prefixSum = 0;
    for (int count : _columnCounts)
    {
        _columnPrefixSums.push_back(prefixSum);
        prefixSum += static_cast<size_t>(count);
    }
    _columnPrefixSums.push_back(prefixSum); // Sentinel for bounds checking

    float maxBottom       = 0.0f;
    float maxRight        = 0.0f;
    const float rowStride = _tileHeightDip + rowSpacingDip;

    for (int columnIndex = 0; columnIndex < static_cast<int>(_columnLayout.size()); ++columnIndex)
    {
        const auto& column = _columnLayout[static_cast<size_t>(columnIndex)];
        float y            = rowSpacingDip;
        for (size_t row = 0; row < column.itemCount && column.startIndex + row < _items.size(); ++row)
        {
            auto& item  = _items[column.startIndex + row];
            item.column = columnIndex;
            item.row    = static_cast<int>(row);
            item.bounds = D2D1::RectF(column.leftDip, y, column.leftDip + column.widthDip, y + _tileHeightDip);
            y += rowStride;
            maxBottom = std::max(maxBottom, item.bounds.bottom);
            maxRight  = std::max(maxRight, item.bounds.right);
        }
    }

    _lastLayoutWidth = _tileWidthDip;
    UpdateItemTextLayouts();

    _contentHeight                  = clientHeightDip;
    _contentWidth                   = std::max({columnLayout.contentWidthDip, maxRight + kColumnSpacingDip, clientWidthDip});
    _scrollOffset                   = 0.0f;
    const float viewWidthDip        = std::max(clientWidthDip, 0.0f);
    const float maxHorizontalOffset = std::max(0.0f, _contentWidth - viewWidthDip);
    _horizontalOffset               = std::clamp(_horizontalOffset, 0.0f, maxHorizontalOffset);
    layoutPerf.SetValue1(static_cast<uint64_t>(_columnLayout.size()));
}

void FolderView::UpdateScrollMetrics()
{
    if (! _hWnd)
    {
        return;
    }

    ShowScrollBar(_hWnd.get(), SB_VERT, FALSE);

    const int contentWidthPx = PxFromDip(_contentWidth);
    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = std::max(0, contentWidthPx);
    si.nPage  = static_cast<UINT>(_clientSize.cx);
    si.nPos   = PxFromDip(_horizontalOffset);
    SetScrollInfo(_hWnd.get(), SB_HORZ, &si, TRUE);

    const bool needHorizontal = contentWidthPx > _clientSize.cx;
    ShowScrollBar(_hWnd.get(), SB_HORZ, needHorizontal ? TRUE : FALSE);
}

float FolderView::GetItemTextLayoutWidth(const FolderItem& item) const noexcept
{
    const float itemWidth = std::max(0.0f, item.bounds.right - item.bounds.left);
    return std::max(0.0f, itemWidth - (kLabelHorizontalPaddingDip * 2.0f) - _iconSizeDip - kIconTextGapDip);
}

void FolderView::UpdateItemTextLayouts()
{
    if (! _dwriteFactory || ! _labelFormat)
    {
        return;
    }

    // Only update layouts for visible items to avoid O(N) DirectWrite operations
    const auto [startIndex, endIndex] = GetVisibleItemRange();
    if (startIndex >= _items.size())
    {
        return;
    }

    const bool includeDetailsLine =
        _displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    const bool includeMetadataLine        = _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    const float constrainedHeight         = std::max(_labelHeightDip, 1.0f);
    const float constrainedDetailsHeight  = std::max(_detailsLineHeightDip, 1.0f);
    const float constrainedMetadataHeight = std::max(_metadataLineHeightDip, 1.0f);

    // Track scroll direction for predictive pre-loading
    if (_horizontalOffset != _lastHorizontalOffset)
    {
        _scrollDirectionX     = (_horizontalOffset > _lastHorizontalOffset) ? 1 : -1;
        _lastHorizontalOffset = _horizontalOffset;
    }
    if (_scrollOffset != _lastScrollOffset)
    {
        _scrollDirectionY = (_scrollOffset > _lastScrollOffset) ? 1 : -1;
        _lastScrollOffset = _scrollOffset;
    }

    // Process visible items + biased buffer based on scroll direction
    // Pre-load more items in the direction of scroll for smoother experience
    constexpr size_t kBufferItems   = 10;
    constexpr size_t kPredictBuffer = 30; // Extra items in scroll direction
    const size_t bufferBefore       = (_scrollDirectionX < 0) ? kPredictBuffer : kBufferItems;
    const size_t bufferAfter        = (_scrollDirectionX > 0) ? kPredictBuffer : kBufferItems;
    const size_t rangeStart         = (startIndex > bufferBefore) ? startIndex - bufferBefore : 0;
    const size_t rangeEnd           = std::min(endIndex + bufferAfter, _items.size());

    for (size_t i = rangeStart; i < rangeEnd; ++i)
    {
        auto& item                   = _items[i];
        const float constrainedWidth = std::max(GetItemTextLayoutWidth(item), 1.0f);

        if (item.displayName.empty())
        {
            item.labelLayout.reset();
            item.detailsLayout.reset();
            item.detailsMetrics = {};
            item.metadataLayout.reset();
            item.metadataMetrics = {};
            continue;
        }

        // Create label layout lazily if needed
        if (! item.labelLayout)
        {
            wil::com_ptr<IDWriteTextLayout> layout;
            HRESULT hr = _dwriteFactory->CreateTextLayout(item.displayName.data(),
                                                          static_cast<UINT32>(item.displayName.length()),
                                                          _labelFormat.get(),
                                                          constrainedWidth,
                                                          constrainedHeight,
                                                          layout.addressof());
            if (FAILED(hr))
            {
                continue;
            }

            ConfigureLabelLayout(layout.get(), _ellipsisSign.get());

            // Update metrics with actual measured values
            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(layout->GetMetrics(&metrics)))
            {
                item.labelMetrics = metrics;
            }

            item.labelLayout = std::move(layout);
        }

        if (item.labelLayout)
        {
            item.labelLayout->SetMaxWidth(constrainedWidth);
            item.labelLayout->SetMaxHeight(constrainedHeight);
        }

        if (! includeDetailsLine)
        {
            item.detailsLayout.reset();
            item.detailsMetrics = {};
            item.metadataLayout.reset();
            item.metadataMetrics = {};
            continue;
        }

        if (! _detailsFormat)
        {
            continue;
        }

        if (item.detailsText.empty())
        {
            if (_detailsTextProvider)
            {
                item.detailsText =
                    _detailsTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
            }
            else
            {
                item.detailsText = BuildDetailsText(item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes, _detailsSizeSlotChars);
            }
        }

        if (! item.detailsLayout)
        {
            wil::com_ptr<IDWriteTextLayout> layout;
            const HRESULT hr = _dwriteFactory->CreateTextLayout(item.detailsText.c_str(),
                                                                static_cast<UINT32>(item.detailsText.length()),
                                                                _detailsFormat.get(),
                                                                constrainedWidth,
                                                                constrainedDetailsHeight,
                                                                layout.addressof());
            if (FAILED(hr))
            {
                continue;
            }

            ConfigureLabelLayout(layout.get(), _detailsEllipsisSign.get(), false);

            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(layout->GetMetrics(&metrics)))
            {
                item.detailsMetrics = metrics;
            }

            item.detailsLayout = std::move(layout);
        }

        if (item.detailsLayout)
        {
            item.detailsLayout->SetMaxWidth(constrainedWidth);
            item.detailsLayout->SetMaxHeight(constrainedDetailsHeight);
        }

        if (! includeMetadataLine)
        {
            item.metadataLayout.reset();
            item.metadataMetrics = {};
            continue;
        }

        if (item.metadataText.empty() && _metadataTextProvider)
        {
            item.metadataText =
                _metadataTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
        }

        if (! item.metadataLayout && ! item.metadataText.empty())
        {
            wil::com_ptr<IDWriteTextLayout> layout;
            const HRESULT hr = _dwriteFactory->CreateTextLayout(item.metadataText.c_str(),
                                                                static_cast<UINT32>(item.metadataText.length()),
                                                                _detailsFormat.get(),
                                                                constrainedWidth,
                                                                constrainedMetadataHeight,
                                                                layout.addressof());
            if (FAILED(hr))
            {
                continue;
            }

            ConfigureLabelLayout(layout.get(), _detailsEllipsisSign.get(), false);

            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(layout->GetMetrics(&metrics)))
            {
                item.metadataMetrics = metrics;
            }

            item.metadataLayout = std::move(layout);
        }

        if (item.metadataLayout)
        {
            item.metadataLayout->SetMaxWidth(constrainedWidth);
            item.metadataLayout->SetMaxHeight(constrainedMetadataHeight);
        }
    }

    // For large directories, release rendering state for distant items to bound memory
    // Temporarily disabled - may interfere with icon loading
    // ReleaseDistantRenderingState();
}

std::pair<size_t, size_t> FolderView::GetVisibleItemRange() const
{
    if (_items.empty() || _columnLayout.empty() || _tileHeightDip <= 0.0f)
    {
        return {0, _items.size()};
    }

    const float viewWidthDip  = std::max(0.0f, DipFromPx(_clientSize.cx));
    const float viewHeightDip = std::max(0.0f, DipFromPx(_clientSize.cy));

    if (viewWidthDip <= 0.0f || viewHeightDip <= 0.0f)
    {
        return {0, _items.size()};
    }

    const float layoutLeft  = _horizontalOffset;
    const float layoutRight = _horizontalOffset + viewWidthDip;

    size_t firstVisibleColumn = _columnLayout.size();
    size_t lastVisibleColumn  = 0;
    for (size_t columnIndex = 0; columnIndex < _columnLayout.size(); ++columnIndex)
    {
        const auto& column = _columnLayout[columnIndex];
        if (column.RightDip() < layoutLeft)
        {
            continue;
        }
        if (column.leftDip > layoutRight)
        {
            break;
        }
        if (firstVisibleColumn == _columnLayout.size())
        {
            firstVisibleColumn = columnIndex;
        }
        lastVisibleColumn = columnIndex;
    }

    if (firstVisibleColumn == _columnLayout.size())
    {
        return {0, 0};
    }

    const size_t startIndex = std::min(_columnLayout[firstVisibleColumn].startIndex, _items.size());
    const size_t endIndex   = std::min(_columnLayout[lastVisibleColumn].startIndex + _columnLayout[lastVisibleColumn].itemCount, _items.size());
    return {startIndex, std::min(endIndex, _items.size())};
}

void FolderView::ReleaseDistantRenderingState()
{
    // For large directories, release rendering resources (layouts, icons) for items
    // far from the visible range to bound memory usage
    constexpr size_t kMinItemsForSparseMode = 10000; // Only apply to large directories
    constexpr size_t kKeepAroundVisible     = 2000;  // Keep this many items around visible range

    if (_items.size() < kMinItemsForSparseMode)
    {
        return; // Small directory, keep all rendering state
    }

    const auto [visStart, visEnd] = GetVisibleItemRange();

    // Calculate the range of items to keep
    const size_t keepStart = (visStart > kKeepAroundVisible) ? (visStart - kKeepAroundVisible) : 0;
    const size_t keepEnd   = std::min(visEnd + kKeepAroundVisible, _items.size());

    size_t released = 0;

    // Release items before the keep range
    for (size_t i = 0; i < keepStart && i < _items.size(); ++i)
    {
        auto& item = _items[i];
        if (item.labelLayout || item.detailsLayout || item.metadataLayout || item.icon || item.thumbnail)
        {
            item.labelLayout.reset();
            item.labelMetrics = {};
            item.detailsLayout.reset();
            item.detailsMetrics = {};
            item.detailsText.clear();
            item.detailsText.shrink_to_fit();
            item.metadataLayout.reset();
            item.metadataMetrics = {};
            item.metadataText.clear();
            item.metadataText.shrink_to_fit();
            item.icon.reset();
            item.thumbnail.reset();
            item.thumbnailFallbackResolved = false;
            item.thumbnailFallbackTargetPx = 0u;
            ++released;
        }
    }

    // Release items after the keep range
    for (size_t i = keepEnd; i < _items.size(); ++i)
    {
        auto& item = _items[i];
        if (item.labelLayout || item.detailsLayout || item.metadataLayout || item.icon || item.thumbnail)
        {
            item.labelLayout.reset();
            item.labelMetrics = {};
            item.detailsLayout.reset();
            item.detailsMetrics = {};
            item.detailsText.clear();
            item.detailsText.shrink_to_fit();
            item.metadataLayout.reset();
            item.metadataMetrics = {};
            item.metadataText.clear();
            item.metadataText.shrink_to_fit();
            item.icon.reset();
            item.thumbnail.reset();
            item.thumbnailFallbackResolved = false;
            item.thumbnailFallbackTargetPx = 0u;
            ++released;
        }
    }

    if (released > 0)
    {
        Debug::Info(L"FolderView: Released rendering state for {} distant items (visible: {}-{}, keep: {}-{})", released, visStart, visEnd, keepStart, keepEnd);
    }
}

std::wstring_view FolderView::GetVisualDisplayName(const FolderItem& item) const noexcept
{
    if (_fileExtensionsVisible || item.isDirectory)
    {
        return item.displayName;
    }

    return item.GetNameWithoutExtension();
}

void FolderView::EnsureItemTextLayout(FolderItem& item, float labelWidth)
{
    if (! _dwriteFactory || ! _labelFormat)
    {
        return;
    }

    const std::wstring_view labelText = GetVisualDisplayName(item);
    if (labelText.empty())
    {
        return;
    }

    const bool includeDetailsLine =
        _displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    const bool includeMetadataLine        = _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    const float constrainedWidth          = std::max(labelWidth, 1.0f);
    const float constrainedHeight         = std::max(_labelHeightDip, 1.0f);
    const float constrainedDetailsHeight  = std::max(_detailsLineHeightDip, 1.0f);
    const float constrainedMetadataHeight = std::max(_metadataLineHeightDip, 1.0f);

    // Create label layout if not yet created
    if (! item.labelLayout)
    {
        wil::com_ptr<IDWriteTextLayout> layout;
        HRESULT hr = _dwriteFactory->CreateTextLayout(
            labelText.data(), static_cast<UINT32>(labelText.length()), _labelFormat.get(), constrainedWidth, constrainedHeight, layout.addressof());
        if (SUCCEEDED(hr))
        {
            ConfigureLabelLayout(layout.get(), _ellipsisSign.get());

            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(layout->GetMetrics(&metrics)))
            {
                item.labelMetrics = metrics;
            }

            item.labelLayout = std::move(layout);
        }
    }
    else
    {
        item.labelLayout->SetMaxWidth(constrainedWidth);
        item.labelLayout->SetMaxHeight(constrainedHeight);
    }

    // Create details layout if the display mode includes secondary lines.
    if (includeDetailsLine && _detailsFormat)
    {
        if (item.detailsText.empty())
        {
            if (_detailsTextProvider)
            {
                item.detailsText =
                    _detailsTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
            }
            else
            {
                item.detailsText = BuildDetailsText(item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes, _detailsSizeSlotChars);
            }
        }

        if (! item.detailsLayout && ! item.detailsText.empty())
        {
            wil::com_ptr<IDWriteTextLayout> layout;
            const HRESULT hr = _dwriteFactory->CreateTextLayout(item.detailsText.c_str(),
                                                                static_cast<UINT32>(item.detailsText.length()),
                                                                _detailsFormat.get(),
                                                                constrainedWidth,
                                                                constrainedDetailsHeight,
                                                                layout.addressof());
            if (SUCCEEDED(hr))
            {
                ConfigureLabelLayout(layout.get(), _detailsEllipsisSign.get(), false);

                DWRITE_TEXT_METRICS metrics{};
                if (SUCCEEDED(layout->GetMetrics(&metrics)))
                {
                    item.detailsMetrics = metrics;
                }

                item.detailsLayout = std::move(layout);
            }
        }
        else if (item.detailsLayout)
        {
            item.detailsLayout->SetMaxWidth(constrainedWidth);
            item.detailsLayout->SetMaxHeight(constrainedDetailsHeight);
        }

        if (includeMetadataLine)
        {
            if (item.metadataText.empty() && _metadataTextProvider)
            {
                item.metadataText =
                    _metadataTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
            }

            if (! item.metadataLayout && ! item.metadataText.empty())
            {
                wil::com_ptr<IDWriteTextLayout> layout;
                const HRESULT hr = _dwriteFactory->CreateTextLayout(item.metadataText.c_str(),
                                                                    static_cast<UINT32>(item.metadataText.length()),
                                                                    _detailsFormat.get(),
                                                                    constrainedWidth,
                                                                    constrainedMetadataHeight,
                                                                    layout.addressof());
                if (SUCCEEDED(hr))
                {
                    ConfigureLabelLayout(layout.get(), _detailsEllipsisSign.get(), false);

                    DWRITE_TEXT_METRICS metrics{};
                    if (SUCCEEDED(layout->GetMetrics(&metrics)))
                    {
                        item.metadataMetrics = metrics;
                    }

                    item.metadataLayout = std::move(layout);
                }
            }
            else if (item.metadataLayout)
            {
                item.metadataLayout->SetMaxWidth(constrainedWidth);
                item.metadataLayout->SetMaxHeight(constrainedMetadataHeight);
            }
        }
        else
        {
            item.metadataLayout.reset();
            item.metadataMetrics = {};
        }
    }
}

void FolderView::ScheduleIdleLayoutCreation()
{
    // Don't schedule if already running or no items need processing
    if (_idleLayoutTimer != 0 || _items.empty())
    {
        return;
    }

    // Reset index to start from visible items and work outward
    const auto [startIndex, endIndex] = GetVisibleItemRange();
    _idleLayoutNextIndex              = endIndex; // Start from just after visible items

    // Only schedule if there are items without layouts
    bool hasUnprocessedItems = false;
    for (size_t i = _idleLayoutNextIndex; i < _items.size(); ++i)
    {
        if (! _items[i].labelLayout && ! _items[i].displayName.empty())
        {
            hasUnprocessedItems = true;
            break;
        }
    }

    if (! hasUnprocessedItems)
    {
        // Check items before visible range too
        for (size_t i = 0; i < startIndex && i < _items.size(); ++i)
        {
            if (! _items[i].labelLayout && ! _items[i].displayName.empty())
            {
                hasUnprocessedItems  = true;
                _idleLayoutNextIndex = i;
                break;
            }
        }
    }

    if (hasUnprocessedItems && _hWnd)
    {
        _idleLayoutTimer = SetTimer(_hWnd.get(), kIdleLayoutTimerId, kIdleLayoutIntervalMs, nullptr);
    }
}

void FolderView::ProcessIdleLayoutBatch()
{
    if (! _dwriteFactory || ! _labelFormat || _items.empty())
    {
        if (_idleLayoutTimer != 0 && _hWnd)
        {
            KillTimer(_hWnd.get(), kIdleLayoutTimerId);
            _idleLayoutTimer = 0;
        }
        return;
    }

    const bool includeDetailsLine =
        _displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    const bool includeMetadataLine        = _displayMode == DisplayMode::ExtraDetailed || _displayMode == DisplayMode::Thumbnails;
    const float constrainedHeight         = std::max(_labelHeightDip, 1.0f);
    const float constrainedDetailsHeight  = std::max(_detailsLineHeightDip, 1.0f);
    const float constrainedMetadataHeight = std::max(_metadataLineHeightDip, 1.0f);

    size_t processed      = 0;
    const size_t startIdx = _idleLayoutNextIndex;

    // Process a batch of items
    while (processed < kIdleLayoutBatchSize && _idleLayoutNextIndex < _items.size())
    {
        auto& item = _items[_idleLayoutNextIndex];
        ++_idleLayoutNextIndex;

        if (item.displayName.empty() || item.labelLayout)
        {
            continue; // Skip empty names or already processed items
        }

        const float constrainedWidth = std::max(GetItemTextLayoutWidth(item), 1.0f);

        // Create label layout
        wil::com_ptr<IDWriteTextLayout> layout;
        HRESULT hr = _dwriteFactory->CreateTextLayout(item.displayName.data(),
                                                      static_cast<UINT32>(item.displayName.length()),
                                                      _labelFormat.get(),
                                                      constrainedWidth,
                                                      constrainedHeight,
                                                      layout.addressof());
        if (SUCCEEDED(hr))
        {
            ConfigureLabelLayout(layout.get(), _ellipsisSign.get());
            DWRITE_TEXT_METRICS metrics{};
            if (SUCCEEDED(layout->GetMetrics(&metrics)))
            {
                item.labelMetrics = metrics;
            }
            item.labelLayout = std::move(layout);
        }

        // Create details layout if needed
        if (includeDetailsLine && _detailsFormat)
        {
            if (item.detailsText.empty())
            {
                if (_detailsTextProvider)
                {
                    item.detailsText =
                        _detailsTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
                }
                else
                {
                    item.detailsText = BuildDetailsText(item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes, _detailsSizeSlotChars);
                }
            }

            if (! item.detailsLayout && ! item.detailsText.empty())
            {
                wil::com_ptr<IDWriteTextLayout> detailsLayout;
                hr = _dwriteFactory->CreateTextLayout(item.detailsText.c_str(),
                                                      static_cast<UINT32>(item.detailsText.length()),
                                                      _detailsFormat.get(),
                                                      constrainedWidth,
                                                      constrainedDetailsHeight,
                                                      detailsLayout.addressof());
                if (SUCCEEDED(hr))
                {
                    ConfigureLabelLayout(detailsLayout.get(), _detailsEllipsisSign.get(), false);
                    DWRITE_TEXT_METRICS metrics{};
                    if (SUCCEEDED(detailsLayout->GetMetrics(&metrics)))
                    {
                        item.detailsMetrics = metrics;
                    }
                    item.detailsLayout = std::move(detailsLayout);
                }
            }

            if (includeMetadataLine)
            {
                if (item.metadataText.empty() && _metadataTextProvider)
                {
                    item.metadataText =
                        _metadataTextProvider(_itemsFolder, item.displayName, item.isDirectory, item.sizeBytes, item.lastWriteTime, item.fileAttributes);
                }

                if (! item.metadataLayout && ! item.metadataText.empty())
                {
                    wil::com_ptr<IDWriteTextLayout> metaLayout;
                    hr = _dwriteFactory->CreateTextLayout(item.metadataText.c_str(),
                                                          static_cast<UINT32>(item.metadataText.length()),
                                                          _detailsFormat.get(),
                                                          constrainedWidth,
                                                          constrainedMetadataHeight,
                                                          metaLayout.addressof());
                    if (SUCCEEDED(hr))
                    {
                        ConfigureLabelLayout(metaLayout.get(), _detailsEllipsisSign.get(), false);
                        DWRITE_TEXT_METRICS metaMetrics{};
                        if (SUCCEEDED(metaLayout->GetMetrics(&metaMetrics)))
                        {
                            item.metadataMetrics = metaMetrics;
                        }
                        item.metadataLayout = std::move(metaLayout);
                    }
                }
            }
            else
            {
                item.metadataLayout.reset();
                item.metadataMetrics = {};
            }
        }

        ++processed;
    }

    // Check if we're done
    if (_idleLayoutNextIndex >= _items.size())
    {
        // Wrap around to process items before the visible range
        const auto [visStart, visEnd] = GetVisibleItemRange();
        if (startIdx > 0 && visStart > 0)
        {
            _idleLayoutNextIndex = 0;
        }
        else
        {
            // All items processed, stop the timer
            if (_idleLayoutTimer != 0 && _hWnd)
            {
                KillTimer(_hWnd.get(), kIdleLayoutTimerId);
                _idleLayoutTimer = 0;
                Debug::Info(L"FolderView: Idle layout pre-creation complete for {} items", _items.size());
            }
        }
    }
}

std::optional<size_t> FolderView::HitTest(POINT clientPt) const
{
    float x = DipFromPx(clientPt.x) + _horizontalOffset;
    float y = DipFromPx(clientPt.y) + _scrollOffset;
    if (_columnLayout.empty() || _tileHeightDip <= 0.0f)
    {
        for (size_t i = 0; i < _items.size(); ++i)
        {
            const auto& item = _items[i];
            if (x >= item.bounds.left && x <= item.bounds.right && y >= item.bounds.top && y <= item.bounds.bottom)
            {
                return i;
            }
        }
        return std::nullopt;
    }

    const float rowSpacingDip = GetFolderViewRowSpacingDip(_appTheme);
    const float rowStride     = _tileHeightDip + rowSpacingDip;
    if (rowStride <= 0.0f)
    {
        return std::nullopt;
    }

    const float firstRowTop = rowSpacingDip;
    if (y < firstRowTop)
    {
        return std::nullopt;
    }

    const std::optional<size_t> columnIndex = FolderViewColumnLayout::ResolveHitColumnIndex(x, _columnLayout);
    if (! columnIndex.has_value())
    {
        return std::nullopt;
    }
    const FolderViewColumnLayout::Column& column = _columnLayout[columnIndex.value()];

    int row = static_cast<int>(std::floor((y - firstRowTop) / rowStride));
    if (row < 0 || row >= static_cast<int>(column.itemCount))
    {
        return std::nullopt;
    }

    const float rowTop = firstRowTop + static_cast<float>(row) * rowStride;
    if (y > rowTop + _tileHeightDip)
    {
        return std::nullopt;
    }

    const size_t index = column.startIndex + static_cast<size_t>(row);
    if (index >= _items.size())
    {
        return std::nullopt;
    }
    return index;
}

POINT FolderView::ScreenToClientPoint(POINT screenPt) const
{
    POINT pt = screenPt;
    ScreenToClient(_hWnd.get(), &pt);
    return pt;
}

void FolderView::EnsureVisible(size_t index)
{
    if (index >= _items.size())
        return;

    const auto& item         = _items[index];
    const auto& bounds       = item.bounds;
    const float viewWidthDip = std::max(0.0f, DipFromPx(_clientSize.cx));
    float columnLeft         = bounds.left;
    float columnWidth        = std::max(0.0f, bounds.right - bounds.left);
    if (item.column >= 0 && item.column < static_cast<int>(_columnLayout.size()))
    {
        const auto& column = _columnLayout[static_cast<size_t>(item.column)];
        columnLeft         = column.leftDip;
        columnWidth        = column.widthDip;
    }

    if (bounds.left < _horizontalOffset)
    {
        // Item is to the left - scroll to show its column aligned on left
        _horizontalOffset = columnLeft;
    }
    else if (bounds.right > _horizontalOffset + viewWidthDip)
    {
        // Item is to the right - scroll to show its column
        // Try to align column on left edge if possible
        _horizontalOffset = columnLeft;

        // If the column is wider than the viewport, keep the target item's right edge visible.
        if (columnWidth > viewWidthDip)
        {
            _horizontalOffset = bounds.right - viewWidthDip;
        }
    }

    _horizontalOffset = std::clamp(_horizontalOffset, 0.0f, std::max(0.0f, _contentWidth - viewWidthDip));
    UpdateScrollMetrics();
    QueueMissingVisibleThumbnails();
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}
