#include "Document.h"

#include <algorithm>
#include <array>
#include <format>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

// Helper function to strip carriage returns from text
static void StripCarriageReturns(std::wstring& text)
{
    text.erase(std::remove(text.begin(), text.end(), L'\r'), text.end());
}

// Helper function for emoji display
static std::wstring_view EmojiForType(Debug::InfoParam::Type t)
{
    switch (t)
    {
        case Debug::InfoParam::Type::Error: return L"🛑 ";
        case Debug::InfoParam::Type::Warning: return L"⚠️ ";
        case Debug::InfoParam::Type::Info: return L"ℹ️ ";
        case Debug::InfoParam::Type::Debug: return L"🐞 ";
        case Debug::InfoParam::Type::Text: return L"📝 ";
        case Debug::InfoParam::Type::All: // Not a message type
        default: return L"";
    }
}

// --- Document Implementation ---

void Document::InvalidateCaches(CacheInvalidationReason reason)
{
    switch (reason)
    {
        case CacheInvalidationReason::ShowIdsChanged:
            // Prefix and display strings change; max line chars includes prefix so it's stale too
            for (auto& line : _lines)
            {
                line.cachedPrefix.clear();
                line.cachedPrefixLenValid = false;
                line.cachedDisplay.clear();
                line.cachedDisplayValid = false;
            }
            _totalLengthValid  = false;
            _offsetsValid      = false;
            _maxLineCharsValid = false;
            _maxLineChars      = 0;
            _maxLineIndex      = 0;
            break;

        case CacheInvalidationReason::FontChanged:
            // Width measurements invalid, but cached strings still valid
            _maxLineCharsValid = false;
            _maxLineChars      = 0;
            _maxLineIndex      = 0;
            ResetDirtyRange();
            // Display text unchanged, only layout widths affected
            break;

        case CacheInvalidationReason::ThemeChanged:
            // Only colors changed - no text cache invalidation needed
            // Nothing to do for document caches
            break;

        case CacheInvalidationReason::FilterChanged:
            // Visibility changed - visibleLines already rebuilt by SetFilterMask
            // No additional cache invalidation needed
            break;

        case CacheInvalidationReason::FullInvalidation:
        default:
            // Full invalidation (backward compatibility with old behavior)
            _totalLengthValid  = false;
            _offsetsValid      = false;
            _maxLineCharsValid = false;
            _maxLineChars      = 0;
            _maxLineIndex      = 0;
            ResetDirtyRange();
            for (auto& line : _lines)
            {
                line.cachedPrefix.clear();
                line.cachedDisplay.clear();
                line.cachedDisplayValid   = false;
                line.cachedPrefixLenValid = false;
            }
            break;
    }
}

void Document::EnsureOffsetsValid() const
{
    if (_offsetsValid)
        return;

    _lineOffsets.clear();
    _lineOffsets.reserve(_lines.size());
    UINT32 offset = 0;

    for (const auto& line : _lines)
    {
        _lineOffsets.push_back(offset);
        const UINT32 plen = PrefixLength(line);
        offset += plen + static_cast<UINT32>(line.text.size()) + 1; // +1 for '\n' separator between logical lines
    }
    _offsetsValid = true;
}

void Document::ResetDirtyRange()
{
    _dirtyRangeValid = false;
    _dirtyRangeFirst = 0;
    _dirtyRangeLast  = 0;
}

void Document::UpdateDirtyRange(size_t first, size_t last)
{
    if (_lines.empty())
    {
        ResetDirtyRange();
        return;
    }

    const size_t clampedFirst = std::min(first, last);
    const size_t clampedLast  = std::max(first, last);

    if (! _dirtyRangeValid)
    {
        _dirtyRangeValid = true;
        _dirtyRangeFirst = clampedFirst;
        _dirtyRangeLast  = clampedLast;
    }
    else
    {
        _dirtyRangeFirst = std::min(_dirtyRangeFirst, clampedFirst);
        _dirtyRangeLast  = std::max(_dirtyRangeLast, clampedLast);
    }
}

void Document::OnLineLengthChanged(size_t index, [[maybe_unused]] size_t oldLen, size_t newLen) const
{
    if (! _maxLineCharsValid)
    {
        if (newLen > _maxLineChars)
        {
            _maxLineChars = newLen;
            _maxLineIndex = index;
        }
        return;
    }

    if (newLen >= _maxLineChars)
    {
        _maxLineChars = newLen;
        _maxLineIndex = index;
        return;
    }

    if (index == _maxLineIndex && newLen < _maxLineChars)
    {
        _maxLineCharsValid = false;
    }
}

void Document::EnsureTotalLengthValid() const
{
    if (_totalLengthValid)
        return;

    _cachedTotalLength = 0;
    for (size_t i = 0; i < _lines.size(); i++)
    {
        _cachedTotalLength += PrefixLength(_lines[i]);
        _cachedTotalLength += static_cast<UINT32>(_lines[i].text.size());
        if (i + 1 < _lines.size())
            _cachedTotalLength += 1; // '\n' separator
    }
    _totalLengthValid = true;
}

void Document::SetText(const std::wstring& text)
{
    std::unique_lock lock(_rwMutex); // Write operation
    _lines.clear();
    _visibleLines.clear(); // Clear visible lines when replacing all text
    size_t start = 0;
    size_t end   = 0;
    while (end != std::wstring::npos)
    {
        end = text.find(L'\n', start);
        Line line;
        if (end == std::wstring::npos)
            line.text = text.substr(start);
        else
            line.text = text.substr(start, end - start);
        StripCarriageReturns(line.text);
        line.newlineCount       = static_cast<UINT32>(std::count(line.text.begin(), line.text.end(), L'\n'));
        line.cachedDisplayValid = false;
        _lines.push_back(std::move(line));
        start = (end == std::wstring::npos) ? end : end + 1;
    }
    InvalidateCaches();
    MarkAllDirtyUnsafe();  // Already holding lock
    RebuildVisibleLines(); // Rebuild visible lines after replacing all text
}

void Document::AppendText(const std::wstring& more)
{
    std::unique_lock lock(_rwMutex); // Write operation
    if (more.empty())
        return;

    // Optimization - Reserve vector capacity proactively to reduce reallocations
    const size_t currentSize     = _lines.size();
    const size_t currentCapacity = _lines.capacity();
    if (currentCapacity - currentSize < 100)
    {
        const size_t newCapacity = currentCapacity + (currentCapacity / 2) + 100;
        _lines.reserve(newCapacity);
    }

    if (_lines.empty())
        _lines.push_back(Line{});
    const size_t prevLineCount = _lines.size();

    size_t currentIndex           = _lines.size() - 1;
    const wchar_t* data           = more.c_str();
    const size_t length           = more.size();
    size_t segmentStart           = 0;
    size_t totalCharsAppended     = 0;
    size_t newlineSeparatorsAdded = 0;

    auto appendSegment = [&](size_t start, size_t end)
    {
        if (end <= start)
            return; // most of the time return here when you've got \r\n and allready add the segment on \r
        auto& line          = _lines[currentIndex];
        const UINT32 prefix = PrefixLength(line);
        const size_t oldLen = static_cast<size_t>(prefix) + line.text.size();
        const size_t count  = end - start;
        line.text.append(data + start, count);
        line.cachedDisplayValid = false;
        totalCharsAppended += count;
        const size_t newLen = static_cast<size_t>(prefix) + line.text.size();
        OnLineLengthChanged(currentIndex, oldLen, newLen);
    };

    for (size_t i = 0; i < length; ++i)
    {
        const wchar_t ch = data[i];
        if (ch == L'\r')
        {
            appendSegment(segmentStart, i);
            segmentStart = i + 1;
            continue;
        }
        if (ch == L'\n')
        {
            appendSegment(segmentStart, i);
            ++newlineSeparatorsAdded;
            _lines.push_back(Line{});
            currentIndex = _lines.size() - 1;
            segmentStart = i + 1;
            continue;
        }
    }
    appendSegment(segmentStart, length);

    if (_totalLengthValid)
    {
        _cachedTotalLength += static_cast<UINT32>(totalCharsAppended);
        _cachedTotalLength += static_cast<UINT32>(newlineSeparatorsAdded);
    }

    if (_offsetsValid)
    {
        if (_lineOffsets.size() != prevLineCount)
        {
            _offsetsValid = false;
        }
        else
        {
            UINT32 offset = 0;
            if (prevLineCount > 0)
            {
                const auto& tail = _lines[prevLineCount - 1];
                offset           = _lineOffsets.back() + PrefixLength(tail) + static_cast<UINT32>(tail.text.size()) + 1;
            }
            for (size_t idx = prevLineCount; idx < _lines.size() && _offsetsValid; ++idx)
            {
                if (_lineOffsets.size() != idx)
                {
                    _offsetsValid = false;
                    break;
                }
                _lineOffsets.push_back(offset);
                const auto& newLine = _lines[idx];
                offset += PrefixLength(newLine) + static_cast<UINT32>(newLine.text.size()) + 1;
            }
        }
    }

    if (! _lines.empty())
    {
        const size_t lastIndex  = _lines.size() - 1;
        const size_t firstDirty = prevLineCount ? prevLineCount - 1 : 0;
        UpdateDirtyRange(firstDirty, lastIndex);
    }

    // Rebuild visibleLines after bulk append since multiple lines may have been added
    RebuildVisibleLines();
}

void Document::AppendInfoLine(const std::wstring& text, const Debug::InfoParam& info)
{
    std::unique_lock lock(_rwMutex); // Write operation

    // Optimization - Reserve vector capacity proactively to reduce reallocations
    const size_t currentSize     = _lines.size();
    const size_t currentCapacity = _lines.capacity();
    if (currentCapacity - currentSize < 100)
    {
        const size_t newCapacity = currentCapacity + (currentCapacity / 2) + 100;
        _lines.reserve(newCapacity);
    }

    Line line;
    line.text = text;
    StripCarriageReturns(line.text);
    line.newlineCount       = static_cast<UINT32>(std::count(line.text.begin(), line.text.end(), L'\n'));
    line.cachedDisplayValid = false;
    line.hasMeta            = true;
    line.meta               = info;

#ifdef _DEBUG
    if (line.newlineCount > 0)
    {
        auto msg = std::format("AppendInfoLine: line {} has newlineCount={} (embedded newlines in text)\n", _lines.size(), line.newlineCount);
        OutputDebugStringA(msg.c_str());
    }
#endif

    _lines.push_back(std::move(line));

    const size_t newIndex = _lines.size() - 1;
    const size_t newLen   = static_cast<size_t>(PrefixLength(_lines[newIndex])) + _lines[newIndex].text.size();
    OnLineLengthChanged(newIndex, 0, newLen);

    if (_totalLengthValid)
    {
        const auto& back  = _lines.back();
        const UINT32 plen = PrefixLength(back);
        _cachedTotalLength += plen + static_cast<UINT32>(back.text.size());
        if (_lines.size() > 1)
            _cachedTotalLength += 1; // newline separator before this line
    }

    if (_offsetsValid)
    {
        if (_lineOffsets.size() != _lines.size() - 1)
        {
            _offsetsValid = false;
        }
        else
        {
            UINT32 offset = 0;
            if (! _lineOffsets.empty())
            {
                const auto& prev = _lines[_lines.size() - 2];
                offset           = _lineOffsets.back() + PrefixLength(prev) + static_cast<UINT32>(prev.text.size()) + 1;
            }
            _lineOffsets.push_back(offset);
        }
    }

    // Update visibleLines incrementally if the new line is visible
    // Use unsafe version since we already hold unique_lock
    if (IsLineVisibleUnsafe(newIndex))
    {
        UINT32 displayRow = 0;
        if (! _visibleLines.empty())
        {
            // Calculate display row based on last visible line
            const auto& lastVisible = _visibleLines.back();
            displayRow              = lastVisible.displayRowStart + _lines[lastVisible.sourceIndex].newlineCount + 1u;
        }
        _visibleLines.push_back({newIndex, displayRow});
    }

    UpdateDirtyRange(newIndex, newIndex);
}

void Document::Clear()
{
    std::unique_lock lock(_rwMutex); // Write operation
    _lines.clear();
    _visibleLines.clear(); // Clear visible lines when document is cleared
    InvalidateCaches();
}

size_t Document::TotalLength() const
{
    std::shared_lock lock(_rwMutex); // Read operation
    EnsureTotalLengthValid();
    return _cachedTotalLength;
}

size_t Document::LongestLineChars() const
{
    std::shared_lock lock(_rwMutex); // Read operation
    if (! _maxLineCharsValid)
    {
        _maxLineChars = 0;
        _maxLineIndex = 0;
        for (size_t i = 0; i < _lines.size(); ++i)
        {
            const size_t len = static_cast<size_t>(PrefixLength(_lines[i])) + _lines[i].text.size();
            if (len > _maxLineChars)
            {
                _maxLineChars = len;
                _maxLineIndex = i;
            }
        }
        _maxLineCharsValid = true;
    }
    return _maxLineChars;
}

size_t Document::VisibleLineCount() const
{
    std::shared_lock lock(_rwMutex);
    return _visibleLines.size();
}

size_t Document::TotalLineCount() const
{
    std::shared_lock lock(_rwMutex);
    return _lines.size();
}

void Document::SetFilterMask(uint32_t mask)
{
    std::unique_lock lock(_rwMutex);
    if (_filterMask == mask)
        return;

#ifdef _DEBUG
    auto msg = std::format("SetFilterMask: 0x{:02X} -> 0x{:02X} (lineCount={})\n", _filterMask, mask, _lines.size());
    OutputDebugStringA(msg.c_str());
#endif

    _filterMask = mask;
    RebuildVisibleLines(); // Rebuild visible lines vector for new filter
    InvalidateCaches(CacheInvalidationReason::FilterChanged);
    MarkAllDirtyUnsafe();
}

void Document::RebuildVisibleLines()
{
    // No lock needed - called from methods that already hold unique_lock
    _visibleLines.clear();

    if (_filterMask == Debug::InfoParam::Type::All)
    {
        // No filtering - all lines visible
        _visibleLines.reserve(_lines.size());
        UINT32 displayRow = 0;
        for (size_t i = 0; i < _lines.size(); ++i)
        {
            _visibleLines.push_back({i, displayRow});
            displayRow += _lines[i].newlineCount + 1u;
        }
    }
    else
    {
        // Filtering active - only visible lines
        size_t estimatedVisible = _lines.size() / 2; // Rough estimate for initial capacity
        _visibleLines.reserve(estimatedVisible);

        UINT32 displayRow = 0;
        for (size_t i = 0; i < _lines.size(); ++i)
        {
            const auto& line = _lines[i];

            // Lines without metadata are always visible
            bool visible = ! line.hasMeta;
            if (line.hasMeta)
            {
                // Check filter mask for this line's type
                uint32_t bitPos = 0;
                switch (line.meta.type)
                {
                    case Debug::InfoParam::Type::Text: bitPos = 0; break;
                    case Debug::InfoParam::Type::Error: bitPos = 1; break;
                    case Debug::InfoParam::Type::Warning: bitPos = 2; break;
                    case Debug::InfoParam::Type::Info: bitPos = 3; break;
                    case Debug::InfoParam::Type::Debug: bitPos = 4; break;
                    case Debug::InfoParam::Type::All: visible = true; break;
                }
                if (line.meta.type != Debug::InfoParam::Type::All)
                {
                    visible = (_filterMask & (1u << bitPos)) != 0;
                }
            }

            if (visible)
            {
                _visibleLines.push_back({i, displayRow});
                displayRow += line.newlineCount + 1u;
            }
        }
    }

#ifdef _DEBUG
    auto msg = std::format("RebuildVisibleLines: {} visible of {} total lines, {} display rows\n",
                           _visibleLines.size(),
                           _lines.size(),
                           _visibleLines.empty() ? 0u : _visibleLines.back().displayRowStart + _lines[_visibleLines.back().sourceIndex].newlineCount + 1u);
    OutputDebugStringA(msg.c_str());
#endif
}

bool Document::IsLineVisibleUnsafe(size_t sourceIndex) const
{
    // Caller must hold lock
    if (_filterMask == Debug::InfoParam::Type::All)
        return true;

    if (sourceIndex >= _lines.size())
        return false;

    const auto& line = _lines[sourceIndex];
    if (! line.hasMeta)
        return true; // Lines without metadata always visible

    // Convert InfoParam::Type enum value to bit position
    uint32_t bitPos = 0;
    switch (line.meta.type)
    {
        case Debug::InfoParam::Type::Text: bitPos = 0; break;
        case Debug::InfoParam::Type::Error: bitPos = 1; break;
        case Debug::InfoParam::Type::Warning: bitPos = 2; break;
        case Debug::InfoParam::Type::Info: bitPos = 3; break;
        case Debug::InfoParam::Type::Debug: bitPos = 4; break;
        case Debug::InfoParam::Type::All: return true;
    }

    return (_filterMask & (1u << bitPos)) != 0;
}

bool Document::IsLineVisible(size_t sourceIndex) const
{
    std::shared_lock lock(_rwMutex);
    return IsLineVisibleUnsafe(sourceIndex);
}

const Line& Document::GetVisibleLine(size_t visibleIndex) const
{
    std::shared_lock lock(_rwMutex);
    static const Line emptyLine{};

    if (visibleIndex >= _visibleLines.size())
        return emptyLine;

    const size_t sourceIndex = _visibleLines[visibleIndex].sourceIndex;
    if (sourceIndex >= _lines.size())
        return emptyLine;

    return _lines[sourceIndex];
}

const Line& Document::GetSourceLine(size_t sourceIndex) const
{
    std::shared_lock lock(_rwMutex);
    static const Line emptyLine{};

    if (sourceIndex >= _lines.size())
        return emptyLine;

    return _lines[sourceIndex];
}

const std::vector<Line>& Document::Lines() const
{
    return _lines;
}

const std::vector<VisibleLine>& Document::VisibleLines() const
{
    return _visibleLines;
}

UINT32 Document::DisplayRowForVisible(size_t visibleIndex) const
{
    std::shared_lock lock(_rwMutex);

    if (visibleIndex >= _visibleLines.size())
    {
        if (_visibleLines.empty())
            return 0;
        const auto& lastVisible = _visibleLines.back();
        const auto& lastLine    = _lines[lastVisible.sourceIndex];
        return lastVisible.displayRowStart + lastLine.newlineCount + 1u;
    }

    return _visibleLines[visibleIndex].displayRowStart;
}

size_t Document::VisibleIndexFromDisplayRow(UINT32 displayRow) const
{
    std::shared_lock lock(_rwMutex);

    if (_visibleLines.empty())
        return 0;

    // Binary search: find first visible line where displayRowStart > displayRow
    auto it =
        std::upper_bound(_visibleLines.begin(), _visibleLines.end(), displayRow, [](UINT32 row, const VisibleLine& vl) { return row < vl.displayRowStart; });

    // upper_bound returns iterator to first element > displayRow, so back up one
    if (it == _visibleLines.begin())
        return 0;

    return static_cast<size_t>(std::distance(_visibleLines.begin(), it - 1));
}

UINT32 Document::TotalDisplayRows() const
{
    std::shared_lock lock(_rwMutex);

    if (_visibleLines.empty())
        return 0;

    const auto& lastVisible = _visibleLines.back();
    const auto& lastLine    = _lines[lastVisible.sourceIndex];
    return lastVisible.displayRowStart + lastLine.newlineCount + 1u;
}

UINT32 Document::DisplayRowForSource(size_t sourceIndex) const
{
    std::shared_lock lock(_rwMutex);

    if (sourceIndex >= _lines.size())
    {
        if (_visibleLines.empty())
            return 0;
        const auto& lastVisible = _visibleLines.back();
        const auto& lastLine    = _lines[lastVisible.sourceIndex];
        return lastVisible.displayRowStart + lastLine.newlineCount + 1u;
    }

    if (_visibleLines.empty())
        return 0;

    // O(log n): visibleLines is sorted by sourceIndex.
    auto it = std::lower_bound(_visibleLines.begin(), _visibleLines.end(), sourceIndex, [](const VisibleLine& vl, size_t src) { return vl.sourceIndex < src; });

    if (it != _visibleLines.end())
        return it->displayRowStart; // exact match or next visible line (filtered source line)

    const auto& lastVisible = _visibleLines.back();
    const auto& lastLine    = _lines[lastVisible.sourceIndex];
    return lastVisible.displayRowStart + lastLine.newlineCount + 1u;
}

UINT32 Document::GetLineStartOffset(size_t sourceIndex) const
{
    std::shared_lock lock(_rwMutex);

    if (sourceIndex >= _lines.size())
        return 0;

    EnsureOffsetsValid();
    if (! _offsetsValid || _lineOffsets.size() != _lines.size())
        return 0;

    return _lineOffsets[sourceIndex];
}

std::pair<size_t, UINT32> Document::GetLineAndOffsetUnsafe(UINT32 position) const
{
    // Caller must hold lock
    if (_lines.empty())
        return {0, 0};

    const size_t lastIdx = _lines.size() - 1;
    EnsureOffsetsValid();
    if (! _offsetsValid || _lineOffsets.size() != _lines.size())
        return {lastIdx, 0};

    const UINT32 lastStart = _lineOffsets[lastIdx];
    const UINT32 lastLen   = PrefixLength(_lines[lastIdx]) + static_cast<UINT32>(_lines[lastIdx].text.size());
    const UINT32 totalLen  = lastStart + lastLen; // no trailing separator after last line

    if (position >= totalLen)
        return {lastIdx, lastLen};

    // Find the line such that lineStart <= position < nextLineStart.
    auto it                = std::upper_bound(_lineOffsets.begin(), _lineOffsets.end(), position);
    const size_t idx       = (it == _lineOffsets.begin()) ? 0u : static_cast<size_t>((it - _lineOffsets.begin()) - 1);
    const UINT32 lineStart = _lineOffsets[idx];
    const UINT32 off       = position - lineStart;
    const UINT32 lineLen   = PrefixLength(_lines[idx]) + static_cast<UINT32>(_lines[idx].text.size());
    return {idx, std::min(off, lineLen)};
}

std::pair<size_t, UINT32> Document::GetLineAndOffset(UINT32 position) const
{
    std::shared_lock lock(_rwMutex);
    return GetLineAndOffsetUnsafe(position);
}

const std::wstring& Document::GetDisplayTextRef(size_t visibleIndex) const
{
    std::shared_lock lock(_rwMutex);
    static const std::wstring emptyString{};

    if (visibleIndex >= _visibleLines.size())
        return emptyString;

    const size_t sourceIndex = _visibleLines[visibleIndex].sourceIndex;
    if (sourceIndex >= _lines.size())
        return emptyString;

    auto& line = _lines[sourceIndex];
    if (! line.cachedDisplayValid)
    {
        const auto& prefix = BuildPrefix(line);
        line.cachedDisplay.clear();
        line.cachedDisplay.reserve(prefix.size() + line.text.size());
        line.cachedDisplay.append(prefix);
        line.cachedDisplay.append(line.text);
        line.cachedDisplay.erase(std::remove(line.cachedDisplay.begin(), line.cachedDisplay.end(), L'\r'), line.cachedDisplay.end());
        line.cachedDisplayValid = true;
    }
    return line.cachedDisplay;
}

const std::wstring& Document::GetDisplayTextRefAll(size_t sourceIndex) const
{
    std::shared_lock lock(_rwMutex);
    static const std::wstring emptyString{};
    if (sourceIndex >= _lines.size())
        return emptyString;

    auto& line = _lines[sourceIndex];
    if (! line.cachedDisplayValid)
    {
        const auto& prefix = BuildPrefix(line);
        line.cachedDisplay.clear();
        line.cachedDisplay.reserve(prefix.size() + line.text.size());
        line.cachedDisplay.append(prefix);
        line.cachedDisplay.append(line.text);
        line.cachedDisplay.erase(std::remove(line.cachedDisplay.begin(), line.cachedDisplay.end(), L'\r'), line.cachedDisplay.end());
        line.cachedDisplayValid = true;
    }
    return line.cachedDisplay;
}

Document::DisplayTextBatch Document::GetDisplayTextBatch(size_t firstVisible, size_t lastVisible) const
{
    DisplayTextBatch batch;
    batch.lock = std::shared_lock(_rwMutex);
    batch.texts.reserve(lastVisible >= firstVisible ? (lastVisible - firstVisible + 1) : 0);

    for (size_t visIdx = firstVisible; visIdx <= lastVisible && visIdx < _visibleLines.size(); ++visIdx)
    {
        const size_t sourceIdx = _visibleLines[visIdx].sourceIndex;
        if (sourceIdx >= _lines.size())
            break;

        auto& line = _lines[sourceIdx];
        if (! line.cachedDisplayValid)
        {
            const auto& prefix = BuildPrefix(line);
            line.cachedDisplay.clear();
            line.cachedDisplay.reserve(prefix.size() + line.text.size());
            line.cachedDisplay.append(prefix);
            line.cachedDisplay.append(line.text);
            line.cachedDisplay.erase(std::remove(line.cachedDisplay.begin(), line.cachedDisplay.end(), L'\r'), line.cachedDisplay.end());
            line.cachedDisplayValid = true;
        }
        batch.texts.emplace_back(std::cref(line.cachedDisplay));
    }

    // The lock is now owned by the batch object, which will release it when destructed.
    return batch;
}

Document::DisplayTextBatch Document::GetDisplayTextBatchAll(size_t firstAll, size_t lastAll) const
{
    DisplayTextBatch batch;
    batch.lock = std::shared_lock(_rwMutex);
    batch.texts.reserve(lastAll >= firstAll ? (lastAll - firstAll + 1) : 0);

    for (size_t i = firstAll; i <= lastAll && i < _lines.size(); ++i)
    {
        auto& line = _lines[i];
        if (! line.cachedDisplayValid)
        {
            const auto& prefix = BuildPrefix(line);
            line.cachedDisplay.clear();
            line.cachedDisplay.reserve(prefix.size() + line.text.size());
            line.cachedDisplay.append(prefix);
            line.cachedDisplay.append(line.text);
            line.cachedDisplay.erase(std::remove(line.cachedDisplay.begin(), line.cachedDisplay.end(), L'\r'), line.cachedDisplay.end());
            line.cachedDisplayValid = true;
        }
        batch.texts.emplace_back(std::cref(line.cachedDisplay));
    }

    return batch;
}

Document::FilteredTailResult Document::BuildFilteredTailText(size_t firstAll, size_t lastAll) const
{
    FilteredTailResult result;
    std::shared_lock lock(_rwMutex); // Single lock for entire operation

    if (firstAll >= _lines.size())
        return result;

    lastAll = std::min(lastAll, _lines.size() - 1);
    result.lines.reserve(lastAll - firstAll + 1);

    for (size_t i = firstAll; i <= lastAll; ++i)
    {
        if (! IsLineVisibleUnsafe(i))
            continue;

        ++result.visibleCount;

        auto& line = _lines[i];
        if (! line.cachedDisplayValid)
        {
            const auto& prefix = BuildPrefix(line);
            line.cachedDisplay.clear();
            line.cachedDisplay.reserve(prefix.size() + line.text.size());
            line.cachedDisplay.append(prefix);
            line.cachedDisplay.append(line.text);
            line.cachedDisplay.erase(std::remove(line.cachedDisplay.begin(), line.cachedDisplay.end(), L'\r'), line.cachedDisplay.end());
            line.cachedDisplayValid = true;
        }

        const UINT32 prefixLen = PrefixLength(line);
        result.lines.push_back({i, prefixLen, static_cast<UINT32>(line.text.size()), line.hasMeta, line.meta.type});
        result.text.append(line.cachedDisplay);
        result.text.append(L"\n");
    }

    // Remove trailing newline
    if (! result.text.empty())
        result.text.pop_back();

    return result;
}

bool Document::SaveTextToFile(const std::wstring& path) const
{
    std::shared_lock lock(_rwMutex); // Read operation
    std::ofstream file(path, std::ios::binary);
    if (! file)
        return false;
    // Write BOM for UTF-8
    const unsigned char bom[] = {0xEF, 0xBB, 0xBF};
    file.write(reinterpret_cast<const char*>(bom), sizeof(bom));

    std::string strTo;
    strTo.reserve(200); // just to start with some capacity
    for (size_t i = 0; i < _lines.size(); ++i)
    {
        // Convert wchar_t to UTF-8
        const wchar_t* pData = _lines[i].text.data();
        size_t size          = _lines[i].text.size();
        size_t size_needed   = static_cast<size_t>(WideCharToMultiByte(CP_UTF8, 0, pData, static_cast<int>(size), NULL, 0, NULL, NULL));
        if (size_needed <= 0)
            continue;
        if (strTo.capacity() < size_needed)
            strTo.resize(size_needed);

        int len = WideCharToMultiByte(CP_UTF8, 0, pData, static_cast<int>(size), strTo.data(), static_cast<int>(size_needed), NULL, NULL);
        if (len > 0)
        {
            file.write(strTo.data(), len);
            file.put('\n');
        }
    }
    return file.good();
}

std::wstring Document::GetTextRange(UINT32 start, UINT32 length) const
{
    std::shared_lock lock(_rwMutex);
    if (length == 0)
        return L"";

    auto [startLine, startOffset] = GetLineAndOffsetUnsafe(start);
    auto [endLine, endOffset]     = GetLineAndOffsetUnsafe(start + length - 1);

    std::wstring result;
    result.reserve(length);

    // Helper to append a slice within one logical line from a unified (prefix+text) offset
    auto appendSlice = [&](const Line& line, UINT32 from, UINT32 count)
    {
        const UINT32 plen = PrefixLength(line);
        if (count == 0)
            return;
        // from lies in [0, plen + text.size()]
        if (from < plen)
        {
            const auto prefix      = BuildPrefix(line);
            const UINT32 firstPart = std::min(count, plen - from);
            result.append(prefix, from, firstPart);
            if (count > firstPart)
            {
                const UINT32 rem   = count - firstPart;
                const UINT32 tcopy = std::min<UINT32>(rem, static_cast<UINT32>(line.text.size()));
                result.append(line.text, 0, tcopy);
            }
        }
        else
        {
            const UINT32 off   = from - plen;
            const UINT32 tcopy = std::min<UINT32>(count, static_cast<UINT32>(line.text.size()) - std::min<UINT32>(off, static_cast<UINT32>(line.text.size())));
            if (off < line.text.size() && tcopy)
                result.append(line.text, off, tcopy);
        }
    };

    if (startLine == endLine)
    {
        if (startLine < _lines.size())
        {
            appendSlice(_lines[startLine], startOffset, length);
        }
        return result;
    }

    // First line tail
    if (startLine < _lines.size())
    {
        const auto& fl       = _lines[startLine];
        const UINT32 flTotal = PrefixLength(fl) + static_cast<UINT32>(fl.text.size());
        if (startOffset < flTotal)
            appendSlice(fl, startOffset, flTotal - startOffset);
        result += L'\n';
    }
    // Middle full lines
    for (size_t i = startLine + 1; i < endLine && i < _lines.size(); ++i)
    {
        const auto& ml = _lines[i];
        result += BuildPrefix(ml);
        result += ml.text;
        result += L'\n';
    }
    // Last line head
    if (endLine < _lines.size())
    {
        const auto& ll    = _lines[endLine];
        const UINT32 upto = endOffset + 1; // inclusive end
        appendSlice(ll, 0, std::min<UINT32>(upto, PrefixLength(ll) + static_cast<UINT32>(ll.text.size())));
    }
    return result;
}

void Document::AddColorRange(UINT32 start, UINT32 length, const D2D1_COLOR_F& color)
{
    std::unique_lock lock(_rwMutex);

    // Map flat range to lines using GetLineAndOffsetUnsafe
    lock.unlock(); // Release our write lock temporarily
    auto [startLine, startOffset] = GetLineAndOffset(start);
    auto [endLine, endOffset]     = GetLineAndOffset(start + length - 1);
    lock.lock(); // Re-acquire write lock

    for (size_t lineIdx = startLine; lineIdx <= endLine && lineIdx < _lines.size(); ++lineIdx)
    {
        auto& line = _lines[lineIdx];

        const UINT32 plen     = PrefixLength(line);
        UINT32 localStartFull = (lineIdx == startLine) ? startOffset : 0;
        UINT32 localEndFull   = (lineIdx == endLine) ? endOffset : (plen + static_cast<UINT32>(line.text.size()) - 1);
        // Map to text-only coordinates, skipping prefix
        UINT32 localStart = (localStartFull > plen) ? (localStartFull - plen) : 0;
        UINT32 localEnd   = (localEndFull > plen) ? (localEndFull - plen) : 0;
        UINT32 localLen   = localEnd - localStart + 1;

        if (localLen > 0 && localStart < line.text.size())
        {
            localLen = std::min(localLen, static_cast<UINT32>(line.text.size()) - localStart);
            line.spans.push_back({localStart, localLen, color});
        }
    }
}

void Document::ClearColoring()
{
    std::unique_lock lock(_rwMutex); // Write operation
    for (auto& line : _lines)
        line.spans.clear();
}

std::optional<std::pair<size_t, size_t>> Document::ExtractDirtyLineRange()
{
    std::unique_lock lock(_rwMutex); // Write operation (modifies _dirtyRangeValid)
    if (! _dirtyRangeValid)
        return std::nullopt;
    std::pair<size_t, size_t> range{_dirtyRangeFirst, _dirtyRangeLast};
    ResetDirtyRange();
    return range;
}

void Document::MarkAllDirty()
{
    std::unique_lock lock(_rwMutex); // Write operation
    MarkAllDirtyUnsafe();
}

void Document::MarkAllDirtyUnsafe()
{
    // Assumes lock already held by caller
    if (_lines.empty())
    {
        ResetDirtyRange();
        return;
    }
    _dirtyRangeValid = true;
    _dirtyRangeFirst = 0;
    _dirtyRangeLast  = _lines.size() - 1;
}

void Document::EnableShowIds(bool enable)
{
    std::unique_lock lock(_rwMutex);
    _showIds = enable;
    InvalidateCaches(CacheInvalidationReason::ShowIdsChanged);
    MarkAllDirtyUnsafe(); // Already holding lock
}

const std::wstring& Document::BuildPrefix(const Line& line) const
{
    if (! line.hasMeta)
    {
        static const std::wstring emptyString;
        return emptyString;
    }

    // Cache the built prefix to avoid repeated allocations and conversions
    if (! line.cachedPrefix.empty())
        return line.cachedPrefix;

    line.cachedDisplayValid = false;
    line.cachedPrefix.clear();

    // Pre-allocate buffer for formatting to avoid multiple allocations
    std::array<wchar_t, 128> buffer{};
    int offset = 0;

    // Add emoji
    const auto emoji   = EmojiForType(line.meta.type);
    const int emojiLen = static_cast<int>(emoji.size());
    if (emojiLen > 0 && offset + emojiLen < 128)
    {
        wmemcpy(buffer.data() + offset, emoji.data(), static_cast<size_t>(emojiLen));
        offset += emojiLen;
    }

    // Format time using GetTimeString (HH:MM:SS.mmm)
    const std::wstring ts = line.meta.GetTimeString();
    const int tsLen       = static_cast<int>(ts.size());
    if (tsLen > 0 && offset + tsLen < 128)
    {
        wmemcpy(buffer.data() + offset, ts.data(), static_cast<size_t>(tsLen));
        offset += tsLen;
    }

    if (_showIds)
    {
        // Append process-thread ids if available (non-zero)
        if (line.meta.processID || line.meta.threadID)
        {
            // Format: " PID:TID "
            const int remaining = 128 - offset;
            if (remaining > 20) // Ensure space for IDs
            {
                wchar_t* const outBegin = buffer.data() + offset;
                const auto result       = std::format_to_n(outBegin, static_cast<size_t>(remaining - 1), L" {}:{}", line.meta.processID, line.meta.threadID);
                offset += static_cast<int>(result.out - outBegin);
            }
        }
    }

    // Add trailing space
    if (offset < 127)
    {
        buffer[static_cast<size_t>(offset)] = L' ';
        ++offset;
    }

    line.cachedPrefix.assign(buffer.data(), static_cast<size_t>(offset));
    return line.cachedPrefix;
}

UINT32 Document::PrefixLength(const Line& line) const
{
    if (! line.hasMeta)
        return 0;

    // Use cached length if valid (Optimization - avoids repeated .length() calls)
    if (line.cachedPrefixLenValid)
        return line.cachedPrefixLen;

    // BuildPrefix() caches the string; compute and cache the length
    const auto& prefix        = BuildPrefix(line);
    line.cachedPrefixLen      = static_cast<UINT32>(prefix.length());
    line.cachedPrefixLenValid = true;
    return line.cachedPrefixLen;
}
