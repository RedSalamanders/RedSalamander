#include "NavigationViewInternal.h"

#include <windowsx.h>

#include <commctrl.h>

#include "DirectoryInfoCache.h"
#include "Helpers.h"
#include "IconCache.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "Ui/AnimationDispatcher.h"
#include "resource.h"

void NavigationView::RenderPathSection()
{
    // TRACER_CTX(L"[NavigationView] RenderPathSection called");

    // Ensure D2D resources are initialized before rendering
    EnsureD2DResources();

    if (_clientSize.cx == 0 || _clientSize.cy == 0)
    {
        return;
    }

    if (! _d2dContext || ! _d2dTarget)
    {
        return;
    }

    // Allow rendering background even without path
    _d2dContext->BeginDraw();
    auto endDraw = wil::scope_exit(
        [&]
        {
            const HRESULT hrEnd = _d2dContext->EndDraw();
            if (FAILED(hrEnd))
            {
                if (hrEnd == D2DERR_RECREATE_TARGET)
                {
                    DiscardD2DResources();
                    return;
                }

                Debug::Error(L"[NavigationView] EndDraw failed (hr=0x{:08X})", static_cast<unsigned long>(hrEnd));
                return;
            }

            RECT dirtyRect = _sectionPathRect;
            Present(&dirtyRect);
        });
    _d2dContext->SetTarget(_d2dTarget.get());

    D2D1_RECT_F section2Rect = D2D1::RectF(static_cast<float>(_sectionPathRect.left),
                                           static_cast<float>(_sectionPathRect.top),
                                           static_cast<float>(_sectionPathRect.right),
                                           static_cast<float>(_sectionPathRect.bottom));
    if (_backgroundBrushD2D)
    {
        _d2dContext->FillRectangle(section2Rect, _backgroundBrushD2D.get());
    }

    if (_editMode)
    {
        const auto chrome           = ComputeEditChromeRects(_sectionPathRect, _dpi);
        const D2D1_RECT_F closeRect = D2D1::RectF(static_cast<float>(chrome.closeRect.left),
                                                  static_cast<float>(chrome.closeRect.top),
                                                  static_cast<float>(chrome.closeRect.right),
                                                  static_cast<float>(chrome.closeRect.bottom));

        const D2D1_RECT_F underlineRect = D2D1::RectF(static_cast<float>(chrome.underlineRect.left),
                                                      static_cast<float>(chrome.underlineRect.top),
                                                      static_cast<float>(chrome.underlineRect.right),
                                                      static_cast<float>(chrome.underlineRect.bottom));

        if (_accentBrush)
        {
            _d2dContext->FillRectangle(underlineRect, _accentBrush.get());
        }

        const float hoverInset        = DipsToPixels(kBreadcrumbHoverInsetDip, _dpi);
        const float hoverCornerRadius = DipsToPixels(kBreadcrumbHoverCornerRadiusDip, _dpi);

        if (_editCloseHovered && _hoverBrush)
        {
            const D2D1_RECT_F hoverRect = InsetRectF(closeRect, hoverInset, hoverInset);
            _d2dContext->FillRoundedRectangle(RoundedRect(hoverRect, hoverCornerRadius), _hoverBrush.get());
        }

        ID2D1SolidColorBrush* closeBrush = _textBrush.get();
        if (closeBrush)
        {
            const float iconStroke = std::max(1.0f, DipsToPixels(kEditCloseIconStrokeDip, _dpi));

            const float closeWidth  = std::max(0.0f, closeRect.right - closeRect.left);
            const float closeHeight = std::max(0.0f, closeRect.bottom - closeRect.top);
            const float maxHalf     = std::min(closeWidth, closeHeight) * 0.5f;
            const float iconHalf    = std::min(DipsToPixels(kEditCloseIconHalfDip, _dpi), maxHalf);

            const float cx = (closeRect.left + closeRect.right) * 0.5f;
            const float cy = (closeRect.top + closeRect.bottom) * 0.5f;

            _d2dContext->DrawLine(D2D1::Point2F(cx - iconHalf, cy - iconHalf), D2D1::Point2F(cx + iconHalf, cy + iconHalf), closeBrush, iconStroke);
            _d2dContext->DrawLine(D2D1::Point2F(cx - iconHalf, cy + iconHalf), D2D1::Point2F(cx + iconHalf, cy - iconHalf), closeBrush, iconStroke);
        }
    }
    else
    {
        RenderBreadcrumbs();
    }

    if (! _editMode && _accentBrush && _hWnd && GetFocus() == _hWnd.get() && _focusedRegion == FocusRegion::Path)
    {
        constexpr float inset       = 1.0f;
        const D2D1_RECT_F focusRect = D2D1::RectF(section2Rect.left + inset, section2Rect.top + inset, section2Rect.right - inset, section2Rect.bottom - inset);
        const float cornerRadius    = DipsToPixels(kFocusRingCornerRadiusDip, _dpi);
        const D2D1_ROUNDED_RECT rounded = RoundedRect(focusRect, cornerRadius);
        _d2dContext->DrawRoundedRectangle(rounded, _accentBrush.get(), 2.0f);
    }
    // Reset transform
}

void NavigationView::InvalidateBreadcrumbLayoutCache() noexcept
{
    _breadcrumbTextLayoutCache.clear();
    _breadcrumbTextLayoutCacheFactory = nullptr;
    _breadcrumbTextLayoutCacheFormat  = nullptr;
    _breadcrumbTextLayoutCacheHeight  = 0.0f;

    _breadcrumbLayoutCacheValid           = false;
    _breadcrumbLayoutCachePath            = std::filesystem::path{};
    _breadcrumbLayoutCacheDpi             = USER_DEFAULT_SCREEN_DPI;
    _breadcrumbLayoutCacheAvailableWidth  = 0.0f;
    _breadcrumbLayoutCacheSectionHeight   = 0.0f;
    _breadcrumbLayoutCacheFactory         = nullptr;
    _breadcrumbLayoutCachePathFormat      = nullptr;
    _breadcrumbLayoutCacheSeparatorFormat = nullptr;
}

void NavigationView::EnsureBreadcrumbTextLayoutCache(float height) noexcept
{
    if (_breadcrumbTextLayoutCacheFactory == _dwriteFactory.get() && _breadcrumbTextLayoutCacheFormat == _pathFormat.get() &&
        _breadcrumbTextLayoutCacheHeight == height)
    {
        return;
    }

    _breadcrumbTextLayoutCache.clear();
    _breadcrumbTextLayoutCacheFactory = _dwriteFactory.get();
    _breadcrumbTextLayoutCacheFormat  = _pathFormat.get();
    _breadcrumbTextLayoutCacheHeight  = height;
}

void NavigationView::GetBreadcrumbTextLayoutAndWidth(std::wstring_view text, float height, wil::com_ptr<IDWriteTextLayout>& layout, float& width) noexcept
{
    layout.reset();
    width = 0.0f;

    if (! _dwriteFactory || ! _pathFormat || text.empty())
    {
        return;
    }

    EnsureBreadcrumbTextLayoutCache(height);

    const auto it = _breadcrumbTextLayoutCache.find(text);
    if (it != _breadcrumbTextLayoutCache.end())
    {
        layout = it->second.layout;
        width  = it->second.width;
        return;
    }

    BreadcrumbTextLayoutCacheEntry entry{};
    CreateTextLayoutAndWidth(_dwriteFactory.get(), _pathFormat.get(), text, kIntrinsicTextLayoutMaxWidth, height, entry.layout, entry.width);
    if (! entry.layout)
    {
        return;
    }

    if (_breadcrumbTextLayoutCache.size() >= kMaxBreadcrumbTextLayoutCacheEntries)
    {
        _breadcrumbTextLayoutCache.clear();
    }

    const auto inserted = _breadcrumbTextLayoutCache.emplace(std::wstring(text), std::move(entry));
    layout              = inserted.first->second.layout;
    width               = inserted.first->second.width;
}

void NavigationView::UpdateBreadcrumbLayout()
{
    if (! _currentPluginPath.has_value())
    {
        return;
    }

    if (! _pathFormat || ! _separatorFormat || ! _dwriteFactory)
    {
        EnsureD2DResources();
        if (! _pathFormat || ! _separatorFormat || ! _dwriteFactory)
        {
            return;
        }
    }

    const float paddingX       = DipsToPixels(kPathPaddingDip, _dpi);
    const float separatorWidth = DipsToPixels(kPathSeparatorWidthDip, _dpi);
    const float spacing        = DipsToPixels(kPathSpacingDip, _dpi);
    const float availableWidth = static_cast<float>(_sectionPathRect.right - _sectionPathRect.left) - paddingX * 2.0f;
    const float sectionHeight  = static_cast<float>(_sectionPathRect.bottom - _sectionPathRect.top);

    if (_breadcrumbLayoutCacheValid && _breadcrumbLayoutCachePath == _currentPluginPath.value() && _breadcrumbLayoutCacheDpi == _dpi &&
        _breadcrumbLayoutCacheAvailableWidth == availableWidth && _breadcrumbLayoutCacheSectionHeight == sectionHeight &&
        _breadcrumbLayoutCacheFactory == _dwriteFactory.get() && _breadcrumbLayoutCachePathFormat == _pathFormat.get() &&
        _breadcrumbLayoutCacheSeparatorFormat == _separatorFormat.get() && ! _segments.empty())
    {
        return;
    }

    auto parts = SplitPathComponents(_currentPluginPath.value());
    _segments.clear();
    _separators.clear();

    if (parts.empty())
    {
        _breadcrumbLayoutCacheValid = false;
        Debug::Warning(L"[NavigationView] No path components found");
        return;
    }

    const size_t partCount = parts.size();
    std::vector<float> partWidths;
    partWidths.reserve(partCount);
    std::vector<wil::com_ptr<IDWriteTextLayout>> partLayouts;
    partLayouts.reserve(partCount);

    for (const auto& part : parts)
    {
        float width = 0.0f;
        wil::com_ptr<IDWriteTextLayout> layout;
        GetBreadcrumbTextLayoutAndWidth(part.text, sectionHeight, layout, width);
        partWidths.push_back(width);
        partLayouts.push_back(std::move(layout));
    }

    constexpr std::wstring_view ellipsisText = kEllipsisText;
    float ellipsisWidth                      = 0.0f;
    wil::com_ptr<IDWriteTextLayout> ellipsisLayout;
    GetBreadcrumbTextLayoutAndWidth(ellipsisText, sectionHeight, ellipsisLayout, ellipsisWidth);

    std::vector<float> prefixSums;
    prefixSums.resize(partCount + 1, 0.0f);
    for (size_t i = 0; i < partCount; ++i)
    {
        prefixSums[i + 1] = prefixSums[i] + partWidths[i];
    }

    auto sumFirst = [&](size_t count) -> float { return prefixSums[std::min(count, partCount)]; };

    auto sumLast = [&](size_t count) -> float
    {
        if (count == 0)
        {
            return 0.0f;
        }
        const size_t clamped = std::min(count, partCount);
        return prefixSums[partCount] - prefixSums[partCount - clamped];
    };

    auto sequenceWidth = [&](float sumWidths, size_t segmentCount) -> float
    {
        if (segmentCount == 0)
        {
            return 0.0f;
        }
        return sumWidths + spacing * static_cast<float>(segmentCount) + separatorWidth * static_cast<float>(segmentCount - 1);
    };

    struct CollapsePlan
    {
        size_t prefixCount   = 0;
        size_t suffixCount   = 0;
        bool showEllipsis    = false;
        bool ellipsisAtStart = false;
        bool truncateFirst   = false;
        bool truncateLast    = false;
        std::wstring firstText;
        std::wstring lastText;
    };

    CollapsePlan plan{};

    const float fullWidth = sequenceWidth(prefixSums[partCount], partCount);
    if (fullWidth <= availableWidth)
    {
        plan.prefixCount  = partCount;
        plan.suffixCount  = 0;
        plan.showEllipsis = false;
    }
    else if (partCount == 1)
    {
        plan.prefixCount   = 1;
        plan.showEllipsis  = false;
        plan.truncateFirst = true;
        plan.firstText     = parts[0].text;
    }
    else
    {
        // Choose the widest-fitting collapsed form that keeps the end visible.
        bool found         = false;
        size_t bestShown   = 0;
        size_t bestPrefix  = 0;
        size_t bestSuffix  = 0;
        size_t bestBalance = 0;

        for (size_t prefixCount = 1; prefixCount < partCount; ++prefixCount)
        {
            for (size_t suffixCount = 1; suffixCount < partCount; ++suffixCount)
            {
                if (prefixCount + suffixCount >= partCount)
                {
                    continue;
                }

                const size_t segmentCount = prefixCount + 1 + suffixCount;
                const float sumWidths     = sumFirst(prefixCount) + ellipsisWidth + sumLast(suffixCount);
                const float w             = sequenceWidth(sumWidths, segmentCount);
                if (w > availableWidth)
                {
                    continue;
                }

                const size_t shown   = prefixCount + suffixCount;
                const size_t balance = prefixCount > suffixCount ? (prefixCount - suffixCount) : (suffixCount - prefixCount);
                if (! found || shown > bestShown || (shown == bestShown && balance < bestBalance) ||
                    (shown == bestShown && balance == bestBalance && suffixCount > bestSuffix) ||
                    (shown == bestShown && balance == bestBalance && suffixCount == bestSuffix && prefixCount > bestPrefix))
                {
                    found       = true;
                    bestShown   = shown;
                    bestPrefix  = prefixCount;
                    bestSuffix  = suffixCount;
                    bestBalance = balance;
                }
            }
        }

        if (found)
        {
            plan.prefixCount     = bestPrefix;
            plan.suffixCount     = bestSuffix;
            plan.showEllipsis    = true;
            plan.ellipsisAtStart = false;
        }
        else
        {
            // Try dropping the prefix entirely and keep as much suffix context as possible: "... > tail"
            bool foundSuffix = false;
            size_t bestTail  = 0;
            for (size_t suffixCount = 1; suffixCount < partCount; ++suffixCount)
            {
                const size_t segmentCount = 1 + suffixCount;
                const float sumWidths     = ellipsisWidth + sumLast(suffixCount);
                const float w             = sequenceWidth(sumWidths, segmentCount);
                if (w > availableWidth)
                {
                    continue;
                }

                if (! foundSuffix || suffixCount > bestTail)
                {
                    foundSuffix = true;
                    bestTail    = suffixCount;
                }
            }

            if (foundSuffix)
            {
                plan.prefixCount     = 0;
                plan.suffixCount     = bestTail;
                plan.showEllipsis    = true;
                plan.ellipsisAtStart = true;
            }
            else
            {
                // Fallback: "first > ... > last" with truncation before dropping the first segment.
                const float lastWidth = partWidths.back();
                const float fixed     = ellipsisWidth + lastWidth + spacing * 3.0f + separatorWidth * 2.0f;
                if (fixed < availableWidth)
                {
                    plan.prefixCount   = 1;
                    plan.suffixCount   = 1;
                    plan.showEllipsis  = true;
                    plan.truncateFirst = true;
                    plan.firstText     = parts.front().text;
                }
                else
                {
                    // Fallback: "... > last" with truncation of last if needed.
                    plan.prefixCount     = 0;
                    plan.suffixCount     = 1;
                    plan.showEllipsis    = true;
                    plan.ellipsisAtStart = true;
                    plan.truncateLast    = true;
                    plan.lastText        = parts.back().text;
                }
            }
        }
    }

    auto truncateToWidth = [&](std::wstring_view text, float maxWidth) -> std::wstring
    { return TruncateTextToWidth(_dwriteFactory.get(), _pathFormat.get(), text, maxWidth, sectionHeight, ellipsisText); };

    // Apply truncation decisions now that we know the plan.
    if (plan.truncateFirst && plan.prefixCount > 0)
    {
        size_t segmentCount = plan.prefixCount + (plan.showEllipsis ? 1 : 0) + plan.suffixCount;
        float fixedSum      = sumFirst(plan.prefixCount) - partWidths.front();
        if (plan.showEllipsis)
        {
            fixedSum += ellipsisWidth;
        }
        fixedSum += sumLast(plan.suffixCount);
        const float base          = sequenceWidth(fixedSum, segmentCount);
        const float maxFirstWidth = std::max(0.0f, availableWidth - base);
        plan.firstText            = truncateToWidth(plan.firstText, maxFirstWidth);
        if (plan.firstText == ellipsisText)
        {
            plan.prefixCount     = 0;
            plan.truncateFirst   = false;
            plan.showEllipsis    = true;
            plan.ellipsisAtStart = true;
            plan.suffixCount     = std::min<size_t>(1, partCount);
            plan.truncateLast    = true;
            plan.lastText        = parts.back().text;
        }
    }

    if (plan.truncateLast && plan.suffixCount > 0)
    {
        const size_t segmentCount = (plan.showEllipsis ? 1 : 0) + plan.suffixCount + plan.prefixCount;
        float fixedSum            = sumFirst(plan.prefixCount);
        if (plan.showEllipsis)
        {
            fixedSum += ellipsisWidth;
        }
        fixedSum += sumLast(plan.suffixCount) - partWidths.back();
        const float base         = sequenceWidth(fixedSum, segmentCount);
        const float maxLastWidth = std::max(0.0f, availableWidth - base);
        plan.lastText            = truncateToWidth(plan.lastText, maxLastWidth);
        if (plan.lastText == ellipsisText)
        {
            plan.prefixCount     = 0;
            plan.suffixCount     = 0;
            plan.showEllipsis    = true;
            plan.ellipsisAtStart = true;
            plan.truncateLast    = false;
        }
    }

    struct DisplaySegment
    {
        bool isEllipsis  = false;
        size_t partIndex = 0;
        std::wstring displayText;
    };

    std::vector<DisplaySegment> displaySegments;
    displaySegments.reserve(partCount + 1);

    if (! plan.showEllipsis)
    {
        for (size_t i = 0; i < plan.prefixCount; ++i)
        {
            DisplaySegment ds;
            ds.partIndex = i;
            if (plan.truncateFirst && i == 0)
            {
                ds.displayText = plan.firstText;
            }
            displaySegments.push_back(std::move(ds));
        }
    }
    else
    {
        if (! plan.ellipsisAtStart)
        {
            for (size_t i = 0; i < plan.prefixCount; ++i)
            {
                DisplaySegment ds;
                ds.partIndex = i;
                if (plan.truncateFirst && i == 0)
                {
                    ds.displayText = plan.firstText;
                }
                displaySegments.push_back(std::move(ds));
            }
        }

        DisplaySegment ellipsisSeg;
        ellipsisSeg.isEllipsis = true;
        displaySegments.push_back(std::move(ellipsisSeg));

        const size_t tailStart = partCount - plan.suffixCount;
        for (size_t i = tailStart; i < partCount; ++i)
        {
            DisplaySegment ds;
            ds.partIndex = i;
            if (plan.truncateLast && i == partCount - 1)
            {
                ds.displayText = plan.lastText;
            }
            displaySegments.push_back(std::move(ds));
        }
    }

    float x = paddingX;
    for (size_t displayIndex = 0; displayIndex < displaySegments.size(); ++displayIndex)
    {
        const auto& ds = displaySegments[displayIndex];

        PathSegment segment;
        float segmentWidth = 0.0f;

        if (ds.isEllipsis)
        {
            segment.text       = std::wstring(ellipsisText);
            segment.fullPath   = std::filesystem::path{};
            segment.isEllipsis = true;
            segment.layout     = ellipsisLayout;
            segmentWidth       = ellipsisWidth;
        }
        else
        {
            segment.fullPath   = parts[ds.partIndex].fullPath;
            segment.isEllipsis = false;

            if (! ds.displayText.empty())
            {
                segment.text = ds.displayText;
                GetBreadcrumbTextLayoutAndWidth(segment.text, sectionHeight, segment.layout, segmentWidth);
            }
            else
            {
                segment.text   = parts[ds.partIndex].text;
                segment.layout = partLayouts[ds.partIndex];
                segmentWidth   = partWidths[ds.partIndex];
            }
        }

        segment.bounds = D2D1::RectF(x - spacing / 2.0f, 0.0f, x + segmentWidth + spacing / 2.0f, sectionHeight);
        _segments.push_back(std::move(segment));
        x += segmentWidth + spacing;

        if (displayIndex + 1 < displaySegments.size())
        {
            BreadcrumbSeparator sep;
            sep.bounds            = D2D1::RectF(x, 0.0f, x + separatorWidth, sectionHeight);
            sep.leftSegmentIndex  = _segments.size() - 1;
            sep.rightSegmentIndex = _segments.size();
            _separators.push_back(sep);
            x += separatorWidth;
        }
    }

    // Initialize rotation angles for separators
    _separatorRotationAngles.resize(_separators.size(), 0.0f);
    _separatorTargetAngles.resize(_separators.size(), 0.0f);

    _breadcrumbLayoutCacheValid           = true;
    _breadcrumbLayoutCachePath            = _currentPluginPath.value();
    _breadcrumbLayoutCacheDpi             = _dpi;
    _breadcrumbLayoutCacheAvailableWidth  = availableWidth;
    _breadcrumbLayoutCacheSectionHeight   = sectionHeight;
    _breadcrumbLayoutCacheFactory         = _dwriteFactory.get();
    _breadcrumbLayoutCachePathFormat      = _pathFormat.get();
    _breadcrumbLayoutCacheSeparatorFormat = _separatorFormat.get();

    // OutputDebugStringW(std::format(L"[NavigationView] Layout complete: {} segments, {} separators\n", _segments.size(), _separators.size()).c_str());
}

void NavigationView::RenderBreadcrumbs()
{
    // Set viewport transform to Section Path Coordinates
    _d2dContext->SetTransform(D2D1::Matrix3x2F::Translation(static_cast<float>(_sectionPathRect.left), static_cast<float>(_sectionPathRect.top)));
    auto transformBack = wil::scope_exit([&] { _d2dContext->SetTransform(D2D1::Matrix3x2F::Identity()); });

    if (! _currentPath || ! _pathFormat || ! _separatorFormat)
    {
        return;
    }

    // Render segments from cached layout
    const float textInsetX        = DipsToPixels(kPathTextInsetDip, _dpi);
    const float hoverInset        = DipsToPixels(kBreadcrumbHoverInsetDip, _dpi);
    const float hoverCornerRadius = DipsToPixels(kBreadcrumbHoverCornerRadiusDip, _dpi);
    for (size_t i = 0; i < _segments.size(); ++i)
    {
        const auto& segment = _segments[i];

        // Draw hover background if this segment is hovered
        if (_hoveredSegmentIndex == static_cast<int>(i) && _hoverBrush)
        {
            const D2D1_RECT_F hoverRect = InsetRectF(segment.bounds, hoverInset, hoverInset);
            _d2dContext->FillRoundedRectangle(RoundedRect(hoverRect, hoverCornerRadius), _hoverBrush.get());
        }

        const bool lastSegment = i == (_segments.size() - 1);

        ID2D1SolidColorBrush* textBrush = (! segment.isEllipsis && lastSegment) ? _accentBrush.get() : _textBrush.get();

        if (! segment.isEllipsis && _theme.rainbowMode && _rainbowBrush)
        {
            const uint32_t hash  = StableHash32(std::wstring_view(segment.fullPath.native()));
            const float hue      = static_cast<float>(hash % 360u);
            const float sat      = 0.85f;
            const float val      = _theme.darkBase ? 0.90f : 0.75f;
            D2D1::ColorF rainbow = ColorFromHSV(hue, sat, val, 1.0f);
            if (! _paneFocused)
            {
                const float rainbowBlend = _theme.darkBase ? 0.50f : 0.40f;
                rainbow                  = BlendColorF(rainbow, _theme.background, rainbowBlend);
            }
            _rainbowBrush->SetColor(rainbow);

            D2D1_RECT_F underline = segment.bounds;
            underline.top         = std::max(underline.top, underline.bottom - 2.0f);
            _d2dContext->FillRectangle(underline, _rainbowBrush.get());

            if (lastSegment)
            {
                textBrush = _rainbowBrush.get();
            }
        }

        if (segment.layout)
        {
            _d2dContext->DrawTextLayout(D2D1::Point2F(segment.bounds.left + textInsetX, segment.bounds.top), segment.layout.get(), textBrush);
        }
    }

    // Render separators from cached layout
    for (size_t i = 0; i < _separators.size(); ++i)
    {
        const auto& bounds = _separators[i].bounds;

        // Draw hover/pressed background
        if (static_cast<int>(i) == _hoveredSeparatorIndex && _hoverBrush)
        {
            const D2D1_RECT_F hoverRect = InsetRectF(bounds, hoverInset, hoverInset);
            _d2dContext->FillRoundedRectangle(RoundedRect(hoverRect, hoverCornerRadius), _hoverBrush.get());
        }
        else if (static_cast<int>(i) == _activeSeparatorIndex)
        {
            if (_pressedBrush)
            {
                const D2D1_RECT_F pressedRect = InsetRectF(bounds, hoverInset, hoverInset);
                _d2dContext->FillRoundedRectangle(RoundedRect(pressedRect, hoverCornerRadius), _pressedBrush.get());
            }
        }

        // Apply rotation transform if animating
        float rotationAngle = 0.0f;
        if (i < _separatorRotationAngles.size())
        {
            rotationAngle = _separatorRotationAngles[i];
        }

        if (rotationAngle > 0.1f)
        {
            // Calculate center point of separator rect for rotation
            D2D1_POINT_2F center = D2D1::Point2F((bounds.left + bounds.right) / 2.0f, (bounds.top + bounds.bottom) / 2.0f);

            // Save current transform
            D2D1::Matrix3x2F oldTransform;
            _d2dContext->GetTransform(&oldTransform);

            // Apply rotation around center
            D2D1::Matrix3x2F rotation = D2D1::Matrix3x2F::Rotation(rotationAngle, center);
            _d2dContext->SetTransform(rotation * oldTransform);

            _d2dContext->DrawText(&_breadcrumbSeparatorGlyph, 1, _separatorFormat.get(), bounds, _separatorBrush.get());

            // Restore old transform
            _d2dContext->SetTransform(oldTransform);
        }
        else
        {
            _d2dContext->DrawText(&_breadcrumbSeparatorGlyph, 1, _separatorFormat.get(), bounds, _separatorBrush.get());
        }
    }
}

std::vector<NavigationView::PathSegment> NavigationView::SplitPathComponents(const std::filesystem::path& path)
{
    std::vector<PathSegment> result;

    const bool isFilePlugin = _pluginShortId.empty() || EqualsNoCase(_pluginShortId, L"file");
    if (isFilePlugin)
    {
        std::filesystem::path accumulated;

        if (path.has_root_path())
        {
            accumulated = path.root_path();

            PathSegment root;
            root.text = path.root_name().wstring();
            if (root.text.empty())
            {
                root.text = path.root_path().wstring();
            }

            root.fullPath = accumulated;
            result.push_back(std::move(root));
        }

        for (const auto& part : path)
        {
            if (part == path.root_name() || part == path.root_directory())
            {
                continue;
            }

            accumulated /= part;

            PathSegment segment;
            segment.text     = part.wstring();
            segment.fullPath = accumulated;
            result.push_back(std::move(segment));
        }

        return result;
    }

    std::wstring text = NavigationLocation::NormalizePluginPathText(path.wstring(),
                                                                    NavigationLocation::EmptyPathPolicy::Root,
                                                                    NavigationLocation::LeadingSlashPolicy::Ensure,
                                                                    NavigationLocation::TrailingSlashPolicy::Trim);

    constexpr std::wstring_view kConnPrefix = L"/@conn:";

    std::wstring accumulated;
    size_t start = 1u;

    if (text.starts_with(kConnPrefix))
    {
        const size_t nextSlash = text.find(L'/', 1u);
        const size_t end       = nextSlash == std::wstring::npos ? text.size() : nextSlash;
        const std::wstring_view rootView(text.data() + 1u, end > 1u ? (end - 1u) : 0u);

        PathSegment root;
        root.text = std::wstring(rootView);

        accumulated.reserve(1u + root.text.size());
        accumulated.push_back(L'/');
        accumulated.append(root.text);

        root.fullPath = std::filesystem::path(accumulated);
        result.push_back(std::move(root));

        start = nextSlash == std::wstring::npos ? text.size() : (nextSlash + 1u);
    }
    else
    {
        PathSegment root;
        root.text     = L"/";
        root.fullPath = std::filesystem::path(L"/");
        result.push_back(std::move(root));

        accumulated = L"/";
        start       = 1u;
    }

    while (start < text.size())
    {
        size_t next = text.find(L'/', start);
        if (next == std::wstring::npos)
        {
            next = text.size();
        }

        if (next > start)
        {
            const std::wstring_view partView(text.data() + start, next - start);
            std::wstring part(partView);

            PathSegment segment;
            segment.text = part;

            if (accumulated.size() > 1u)
            {
                accumulated.push_back(L'/');
            }
            accumulated.append(part);

            segment.fullPath = std::filesystem::path(accumulated);
            result.push_back(std::move(segment));
        }

        start = next + 1u;
    }

    return result;
}

void NavigationView::StartSeparatorAnimation(size_t separatorIndex, float targetAngle)
{
    // Ensure vectors are large enough
    if (separatorIndex >= _separatorRotationAngles.size())
    {
        _separatorRotationAngles.resize(separatorIndex + 1, 0.0f);
        _separatorTargetAngles.resize(separatorIndex + 1, 0.0f);
    }

    _separatorTargetAngles[separatorIndex] = targetAngle;

    if (_separatorAnimationSubscriptionId != 0 || ! _hWnd)
    {
        return;
    }

    static constexpr float kAngleEpsilon = 0.01f;
    if (std::abs(_separatorRotationAngles[separatorIndex] - targetAngle) <= kAngleEpsilon)
    {
        if (_separatorRotationAngles[separatorIndex] != targetAngle)
        {
            _separatorRotationAngles[separatorIndex] = targetAngle;
            RenderPathSection();
        }
        return;
    }

    _separatorAnimationLastTickMs     = GetTickCount64();
    _separatorAnimationSubscriptionId = RedSalamander::Ui::AnimationDispatcher::GetInstance().Subscribe(&NavigationView::SeparatorAnimationTickThunk, this);
}

bool NavigationView::SeparatorAnimationTickThunk(void* context, uint64_t nowTickMs) noexcept
{
    auto* self = static_cast<NavigationView*>(context);
    if (! self)
    {
        return false;
    }

    return self->UpdateSeparatorAnimations(nowTickMs);
}

void NavigationView::StopSeparatorAnimation() noexcept
{
    if (_separatorAnimationSubscriptionId == 0)
    {
        _separatorAnimationLastTickMs = 0;
        return;
    }

    RedSalamander::Ui::AnimationDispatcher::GetInstance().Unsubscribe(_separatorAnimationSubscriptionId);
    _separatorAnimationSubscriptionId = 0;
    _separatorAnimationLastTickMs     = 0;
}

bool NavigationView::UpdateSeparatorAnimations(uint64_t nowTickMs) noexcept
{
    if (_separatorAnimationSubscriptionId == 0 || ! _hWnd)
    {
        StopSeparatorAnimation();
        return false;
    }

    float dtSeconds = 0.0f;
    if (_separatorAnimationLastTickMs == 0 || nowTickMs <= _separatorAnimationLastTickMs)
    {
        dtSeconds = 1.0f / 60.0f;
    }
    else
    {
        const uint64_t deltaMs = nowTickMs - _separatorAnimationLastTickMs;
        dtSeconds              = static_cast<float>(deltaMs) / 1000.0f;
        dtSeconds              = std::clamp(dtSeconds, 0.0f, 0.05f);
    }

    _separatorAnimationLastTickMs = nowTickMs;

    static constexpr float kAngleEpsilon = 0.01f;
    const float deltaAngle               = ROTATION_SPEED * dtSeconds;

    bool anyAnimating = false;
    bool anyChanged   = false;

    for (size_t i = 0; i < _separatorRotationAngles.size(); ++i)
    {
        float& current     = _separatorRotationAngles[i];
        const float target = _separatorTargetAngles[i];
        const float diff   = target - current;

        if (std::abs(diff) <= kAngleEpsilon)
        {
            if (current != target)
            {
                current    = target;
                anyChanged = true;
            }
            continue;
        }

        anyAnimating = true;

        const float before = current;
        if (diff > 0.0f)
        {
            current = std::min(current + deltaAngle, target);
        }
        else
        {
            current = std::max(current - deltaAngle, target);
        }

        if (current != before)
        {
            anyChanged = true;
        }
    }

    if (! anyAnimating)
    {
        StopSeparatorAnimation();
    }

    if (anyAnimating || anyChanged)
    {
        RenderPathSection();
    }

    return anyAnimating;
}
