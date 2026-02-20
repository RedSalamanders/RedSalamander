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
    EnsureDeviceIndependentResources();

    // Ensure estimated metrics are computed from actual font (DPI-aware)
    UpdateEstimatedMetrics();

    const float clientWidthDip  = std::max(0.0f, DipFromPx(_clientSize.cx));
    const float clientHeightDip = std::max(0.0f, DipFromPx(_clientSize.cy));

    _columnCounts.clear();
    _columnPrefixSums.clear();

    if (_items.empty() || clientWidthDip <= 0.0f)
    {
        _columns          = 1;
        _rowsPerColumn    = 0;
        _contentHeight    = std::max(clientHeightDip, 0.0f);
        _contentWidth     = std::max(clientWidthDip, 0.0f);
        _horizontalOffset = 0.0f;
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

        if (_displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed)
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

            if (_displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed)
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

                if (_displayMode == DisplayMode::ExtraDetailed)
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

    float textWidthForLayout = maxLabelWidth;
    if (_displayMode == DisplayMode::Detailed)
    {
        textWidthForLayout = std::max(maxLabelWidth, maxDetailsWidth);
    }
    else if (_displayMode == DisplayMode::ExtraDetailed)
    {
        textWidthForLayout = std::max(std::max(maxLabelWidth, maxDetailsWidth), maxMetadataWidth);
    }

    const float minColumnWidth     = _iconSizeDip + kIconTextGapDip + kLabelHorizontalPaddingDip * 2.0f;
    const float textWidthSafety    = std::max(_estimatedCharWidthDip, 8.0f);
    const float desiredColumnWidth = _iconSizeDip + kIconTextGapDip + textWidthForLayout + kLabelHorizontalPaddingDip * 2.0f + textWidthSafety;
    const float targetColumnWidth  = std::max(minColumnWidth, desiredColumnWidth);
    const float maxAllowedWidth    = std::max(1.0f, clientWidthDip);
    _tileWidthDip                  = std::min(targetColumnWidth, maxAllowedWidth);

    _labelHeightDip = maxLabelHeight + kLabelVerticalPaddingDip * 2.0f;
    if (_displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed)
    {
        const float detailsHeight = _detailsLineHeightDip > 0.0f ? _detailsLineHeightDip : 12.0f;
        float textBlockHeight     = maxLabelHeight + kDetailsGapDip + detailsHeight;
        if (_displayMode == DisplayMode::ExtraDetailed && _metadataTextProvider && maxMetadataWidth > 0.0f)
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

    const float columnStride = _tileWidthDip + kColumnSpacingDip;
    const float rowStride    = _tileHeightDip + kRowSpacingDip;

    const int maxRowsPerColumn = std::max(1, static_cast<int>(std::floor((clientHeightDip + kRowSpacingDip) / rowStride)));
    _rowsPerColumn             = std::max(1, maxRowsPerColumn);
    const int requiredColumns  = std::max(1, static_cast<int>(std::ceil(static_cast<float>(_items.size()) / static_cast<float>(_rowsPerColumn))));
    _columns                   = requiredColumns;

    _columnCounts.reserve(static_cast<size_t>(_columns));
    size_t remaining = _items.size();
    for (int column = 0; column < _columns && remaining > 0; ++column)
    {
        const int count = static_cast<int>(std::min<size_t>(static_cast<size_t>(_rowsPerColumn), remaining));
        _columnCounts.push_back(count);
        remaining -= count;
    }
    _columns = static_cast<int>(_columnCounts.size());
    if (_columns < 1)
    {
        _columns = 1;
    }

    // Build prefix sums for O(1) hit testing: _columnPrefixSums[c] = items before column c
    _columnPrefixSums.clear();
    _columnPrefixSums.reserve(_columnCounts.size() + 1);
    size_t prefixSum = 0;
    for (int count : _columnCounts)
    {
        _columnPrefixSums.push_back(prefixSum);
        prefixSum += static_cast<size_t>(count);
    }
    _columnPrefixSums.push_back(prefixSum); // Sentinel for bounds checking

    size_t index    = 0;
    float x         = kColumnSpacingDip;
    float maxBottom = 0.0f;
    float maxRight  = 0.0f;

    for (int column = 0; column < static_cast<int>(_columnCounts.size()) && index < _items.size(); ++column)
    {
        const int itemsInColumn = _columnCounts[static_cast<size_t>(column)];
        float y                 = kRowSpacingDip;
        for (int row = 0; row < itemsInColumn && index < _items.size(); ++row, ++index)
        {
            auto& item  = _items[index];
            item.column = column;
            item.row    = row;
            item.bounds = D2D1::RectF(x, y, x + _tileWidthDip, y + _tileHeightDip);
            y += rowStride;
            maxBottom = std::max(maxBottom, item.bounds.bottom);
            maxRight  = std::max(maxRight, item.bounds.right);
        }
        x += columnStride;
    }

    const float labelWidth = std::max(0.0f, _tileWidthDip - (kLabelHorizontalPaddingDip * 2.0f) - _iconSizeDip - kIconTextGapDip);

    // Track width changes to avoid unnecessary layout work
    constexpr float kWidthChangeThreshold = 1.0f; // DIPs
    if (std::abs(labelWidth - _lastLayoutWidth) > kWidthChangeThreshold)
    {
        _lastLayoutWidth = labelWidth;
    }

    UpdateItemTextLayouts(labelWidth);

    _contentHeight                  = clientHeightDip;
    _contentWidth                   = std::max(maxRight + kColumnSpacingDip, clientWidthDip);
    _scrollOffset                   = 0.0f;
    const float viewWidthDip        = std::max(clientWidthDip, 0.0f);
    const float maxHorizontalOffset = std::max(0.0f, _contentWidth - viewWidthDip);
    _horizontalOffset               = std::clamp(_horizontalOffset, 0.0f, maxHorizontalOffset);
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

void FolderView::UpdateItemTextLayouts(float labelWidth)
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

    const float constrainedWidth          = std::max(labelWidth, 1.0f);
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
        auto& item = _items[i];

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

        if (_displayMode == DisplayMode::Brief)
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

        if (_displayMode != DisplayMode::ExtraDetailed)
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
    if (_items.empty() || _columnCounts.empty() || _tileWidthDip <= 0.0f || _tileHeightDip <= 0.0f)
    {
        return {0, _items.size()};
    }

    const float viewWidthDip  = std::max(0.0f, DipFromPx(_clientSize.cx));
    const float viewHeightDip = std::max(0.0f, DipFromPx(_clientSize.cy));

    if (viewWidthDip <= 0.0f || viewHeightDip <= 0.0f)
    {
        return {0, _items.size()};
    }

    const float columnStride = _tileWidthDip + kColumnSpacingDip;
    const float rowStride    = _tileHeightDip + kRowSpacingDip;

    if (columnStride <= 0.0f || rowStride <= 0.0f)
    {
        return {0, _items.size()};
    }

    // Calculate visible column range
    const float layoutLeft  = _horizontalOffset;
    const float layoutRight = _horizontalOffset + viewWidthDip;

    int firstVisibleColumn = static_cast<int>(std::floor((layoutLeft - kColumnSpacingDip) / columnStride));
    int lastVisibleColumn  = static_cast<int>(std::ceil((layoutRight - kColumnSpacingDip) / columnStride));

    firstVisibleColumn = std::max(0, firstVisibleColumn);
    lastVisibleColumn  = std::min(lastVisibleColumn, static_cast<int>(_columnCounts.size()) - 1);

    if (firstVisibleColumn > lastVisibleColumn)
    {
        return {0, 0};
    }

    // Use prefix sums for O(1) index calculation
    const size_t firstCol = static_cast<size_t>(firstVisibleColumn);
    const size_t lastCol  = static_cast<size_t>(lastVisibleColumn);

    size_t startIndex = 0;
    size_t endIndex   = 0;
    if (! _columnPrefixSums.empty() && firstCol < _columnPrefixSums.size())
    {
        startIndex = _columnPrefixSums[firstCol];
        endIndex   = (lastCol + 1 < _columnPrefixSums.size()) ? _columnPrefixSums[lastCol + 1] : _items.size();
    }
    else
    {
        // Fallback to loop (shouldn't happen if prefix sums are properly maintained)
        for (int col = 0; col < firstVisibleColumn && col < static_cast<int>(_columnCounts.size()); ++col)
        {
            startIndex += static_cast<size_t>(_columnCounts[static_cast<size_t>(col)]);
        }
        endIndex = startIndex;
        for (int col = firstVisibleColumn; col <= lastVisibleColumn && col < static_cast<int>(_columnCounts.size()); ++col)
        {
            endIndex += static_cast<size_t>(_columnCounts[static_cast<size_t>(col)]);
        }
    }

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
        if (item.labelLayout || item.detailsLayout || item.metadataLayout || item.icon)
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
            ++released;
        }
    }

    // Release items after the keep range
    for (size_t i = keepEnd; i < _items.size(); ++i)
    {
        auto& item = _items[i];
        if (item.labelLayout || item.detailsLayout || item.metadataLayout || item.icon)
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
            ++released;
        }
    }

    if (released > 0)
    {
        Debug::Info(L"FolderView: Released rendering state for {} distant items (visible: {}-{}, keep: {}-{})", released, visStart, visEnd, keepStart, keepEnd);
    }
}

void FolderView::EnsureItemTextLayout(FolderItem& item, float labelWidth)
{
    if (! _dwriteFactory || ! _labelFormat)
    {
        return;
    }

    if (item.displayName.empty())
    {
        return;
    }

    const float constrainedWidth          = std::max(labelWidth, 1.0f);
    const float constrainedHeight         = std::max(_labelHeightDip, 1.0f);
    const float constrainedDetailsHeight  = std::max(_detailsLineHeightDip, 1.0f);
    const float constrainedMetadataHeight = std::max(_metadataLineHeightDip, 1.0f);

    // Create label layout if not yet created
    if (! item.labelLayout)
    {
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
    }
    else
    {
        item.labelLayout->SetMaxWidth(constrainedWidth);
        item.labelLayout->SetMaxHeight(constrainedHeight);
    }

    // Create details layout if in detailed/extra detailed mode and not yet created
    if ((_displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed) && _detailsFormat)
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

        if (_displayMode == DisplayMode::ExtraDetailed)
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

    const float labelWidth                = std::max(0.0f, _tileWidthDip - (kLabelHorizontalPaddingDip * 2.0f) - _iconSizeDip - kIconTextGapDip);
    const float constrainedWidth          = std::max(labelWidth, 1.0f);
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
        if ((_displayMode == DisplayMode::Detailed || _displayMode == DisplayMode::ExtraDetailed) && _detailsFormat)
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

            if (_displayMode == DisplayMode::ExtraDetailed)
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
    if (_columnCounts.empty() || _tileWidthDip <= 0.0f || _tileHeightDip <= 0.0f)
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

    const float columnStride = _tileWidthDip + kColumnSpacingDip;
    const float rowStride    = _tileHeightDip + kRowSpacingDip;
    if (columnStride <= 0.0f || rowStride <= 0.0f)
    {
        return std::nullopt;
    }

    const float firstColumnLeft = kColumnSpacingDip;
    const float firstRowTop     = kRowSpacingDip;
    if (x < firstColumnLeft || y < firstRowTop)
    {
        return std::nullopt;
    }

    int column = static_cast<int>(std::floor((x - firstColumnLeft) / columnStride));
    if (column < 0 || column >= static_cast<int>(_columnCounts.size()))
    {
        return std::nullopt;
    }

    const float columnLeft = firstColumnLeft + static_cast<float>(column) * columnStride;
    if (x > columnLeft + _tileWidthDip)
    {
        return std::nullopt;
    }

    int row = static_cast<int>(std::floor((y - firstRowTop) / rowStride));
    if (row < 0 || row >= _columnCounts[static_cast<size_t>(column)])
    {
        return std::nullopt;
    }

    const float rowTop = firstRowTop + static_cast<float>(row) * rowStride;
    if (y > rowTop + _tileHeightDip)
    {
        return std::nullopt;
    }

    // O(1) index calculation using prefix sums
    const size_t columnIndex = static_cast<size_t>(column);
    if (columnIndex >= _columnPrefixSums.size())
    {
        return std::nullopt;
    }
    const size_t index = _columnPrefixSums[columnIndex] + static_cast<size_t>(row);
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
    const float columnStride = _tileWidthDip + kColumnSpacingDip;

    // Calculate the column's left edge (snap to column boundary)
    const float columnLeft = kColumnSpacingDip + (static_cast<float>(item.column) * columnStride);

    if (columnLeft < _horizontalOffset)
    {
        // Item is to the left - scroll to show its column aligned on left
        _horizontalOffset = columnLeft;
    }
    else if (bounds.right > _horizontalOffset + viewWidthDip)
    {
        // Item is to the right - scroll to show its column
        // Try to align column on left edge if possible
        _horizontalOffset = columnLeft;

        // If that would scroll too far, just ensure item is visible
        if (_horizontalOffset > bounds.right - viewWidthDip)
        {
            _horizontalOffset = bounds.right - viewWidthDip;
            // Snap to nearest column boundary
            const float columnIndex = std::round((_horizontalOffset - kColumnSpacingDip) / columnStride);
            _horizontalOffset       = kColumnSpacingDip + (columnIndex * columnStride);
        }
    }

    _horizontalOffset = std::clamp(_horizontalOffset, 0.0f, std::max(0.0f, _contentWidth - viewWidthDip));
    UpdateScrollMetrics();
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}
