#pragma once

#include <algorithm>
#include <cstdint>
#include <fstream>
#include <optional>
#include <shared_mutex>
#include <string>
#include <utility>
#include <vector>

#include <d2d1.h>

#include "Helpers.h"

#pragma warning(push)
#pragma warning(disable : 4820) // bytes padding added after data member

// Line: Represents a single logical line in the document with optional metadata and color spans
struct Line
{
    struct ColorSpan
    {
        UINT32 start  = 0;
        UINT32 length = 0;
        D2D1_COLOR_F color{};
    };

    std::wstring text;                  // message text (may include '\n')
    std::vector<ColorSpan> spans;       // optional text coloring
    bool hasMeta = false;               // whether metadata exists for this logical line
    Debug::InfoParam meta{};            // metadata (time/pid/tid/type)
    mutable std::wstring cachedPrefix;  // Cached prefix to avoid rebuilding time/ids string repeatedly
    mutable std::wstring cachedDisplay; // Cached full display string (prefix+text) to avoid rebuilding
    mutable bool cachedDisplayValid   = false;
    mutable UINT32 cachedPrefixLen    = 0;     // Cached prefix length for fast access
    mutable bool cachedPrefixLenValid = false; // Whether cached length is valid
    UINT32 newlineCount               = 0;     // Cached count of embedded '\n' characters for display-row math
};

// VisibleLine: Lightweight index mapping visible lines to source lines with display row offsets
// Only 12 bytes per visible line, rebuilt on filter changes (O(n) acceptable for user actions)
struct VisibleLine
{
    size_t sourceIndex;     // Index into Document::Lines() vector (source of truth)
    UINT32 displayRowStart; // Display row where this visible line starts (accumulated from newlineCount)
};

// Document: Manages document content with filtering support and display row mapping
// Thread-safe with reader-writer lock for concurrent access
class Document
{
public:
    Document() = default;

    // Explicitly delete copy/move operations due to mutex member
    Document(const Document&)            = delete;
    Document(Document&&)                 = delete;
    Document& operator=(const Document&) = delete;
    Document& operator=(Document&&)      = delete;

    // Content mutation
    void SetText(const std::wstring& text);
    void AppendText(const std::wstring& more);
    void AppendInfoLine(const std::wstring& text, const Debug::InfoParam& info);
    void Clear();

    // Content queries
    size_t TotalLength() const;
    size_t LongestLineChars() const;
    bool SaveTextToFile(const std::wstring& path) const;

    // Line count methods - explicit visible vs total
    size_t VisibleLineCount() const; // Returns visible lines (filtered)
    size_t TotalLineCount() const;   // Returns all lines (unfiltered)

    // Filter methods
    void SetFilterMask(uint32_t mask);
    [[maybe_unused]] uint32_t GetFilterMask() const
    {
        return _filterMask;
    }
    bool IsLineVisible(size_t sourceIndex) const;

    // Line access methods
    const Line& GetVisibleLine(size_t visibleIndex) const; // Access by visible index
    const Line& GetSourceLine(size_t sourceIndex) const;   // Access by source index
    const std::vector<Line>& Lines() const;                // Direct access to lines vector (read-only)
    const std::vector<VisibleLine>& VisibleLines() const;  // Direct access to visibleLines vector (read-only)

    // Display row mapping (VisibleLine architecture)
    UINT32 DisplayRowForVisible(size_t visibleIndex) const;
    size_t VisibleIndexFromDisplayRow(UINT32 displayRow) const;
    UINT32 TotalDisplayRows() const;
    UINT32 DisplayRowForSource(size_t sourceIndex) const; // Helper: map source line to display row

    // Character position mapping (source document - orthogonal to filtering)
    UINT32 GetLineStartOffset(size_t sourceIndex) const;
    std::pair<size_t, UINT32> GetLineAndOffset(UINT32 position) const;

    // Get text slice without full copy (use source indices)
    std::wstring GetTextRange(UINT32 start, UINT32 length) const;
    const std::wstring& GetDisplayTextRef(size_t visibleIndex) const;
    const std::wstring& GetDisplayTextRefAll(size_t sourceIndex) const;

    // Optimization - Batch API for range access with single lock
    struct DisplayTextBatch
    {
        DisplayTextBatch()                              = default;
        DisplayTextBatch(DisplayTextBatch&&)            = default;
        DisplayTextBatch& operator=(DisplayTextBatch&&) = default;

        DisplayTextBatch(const DisplayTextBatch&)            = delete;
        DisplayTextBatch& operator=(const DisplayTextBatch&) = delete;

        std::vector<std::reference_wrapper<const std::wstring>> texts;
        std::shared_lock<std::shared_mutex> lock;
    };
    DisplayTextBatch GetDisplayTextBatch(size_t firstVisible, size_t lastVisible) const;
    DisplayTextBatch GetDisplayTextBatchAll(size_t firstAll, size_t lastAll) const;

    // Coloring
    void AddColorRange(UINT32 start, UINT32 length, const D2D1_COLOR_F& color);
    void ClearColoring();

    // Display Helpers
    void EnableShowIds(bool enable);
    [[maybe_unused]] bool IsShowIds() const
    {
        return _showIds;
    }
    UINT32 PrefixLength(const Line& line) const;
    std::optional<std::pair<size_t, size_t>> ExtractDirtyLineRange();
    void MarkAllDirty();

private:
    // Selective cache invalidation reasons
    enum class CacheInvalidationReason : uint8_t
    {
        ShowIdsChanged,  // Only affects prefix
        FontChanged,     // Only affects width measurements
        ThemeChanged,    // Only colors changed, no text
        FilterChanged,   // Only visibility changed
        FullInvalidation // Invalidate everything
    };

    void RebuildVisibleLines(); // Rebuild visibleLines vector from current filter mask
    // Compute display prefix (emoji + time + ids) based on metadata and settings
    const std::wstring& BuildPrefix(const Line& line) const;

    void InvalidateCaches(CacheInvalidationReason reason = CacheInvalidationReason::FullInvalidation);
    void EnsureOffsetsValid() const;
    void EnsureTotalLengthValid() const;
    void OnLineLengthChanged(size_t index, size_t oldLen, size_t newLen) const;
    void UpdateDirtyRange(size_t first, size_t last);
    void ResetDirtyRange();

    // Unsafe methods - assume caller already holds appropriate lock
    void MarkAllDirtyUnsafe();
    std::pair<size_t, UINT32> GetLineAndOffsetUnsafe(UINT32 position) const;
    bool IsLineVisibleUnsafe(size_t sourceIndex) const;

    // Document content
    std::vector<Line> _lines;               // Source of truth: all lines (append-only)
    std::vector<VisibleLine> _visibleLines; // Computed view: maps visible index -> source index + display row
    mutable std::shared_mutex _rwMutex;     // Reader-writer lock for better concurrency

    // Cache for performance
    mutable bool _totalLengthValid    = false;
    mutable UINT32 _cachedTotalLength = 0;
    mutable bool _offsetsValid        = false;
    mutable std::vector<UINT32> _lineOffsets; // Cumulative offsets for each line (for character positions)
    mutable bool _maxLineCharsValid = false;
    mutable size_t _maxLineChars    = 0;
    mutable size_t _maxLineIndex    = 0;

    // Settings for how to display metadata
    bool _showIds            = true;
    bool _lineNumbersEnabled = false;

    // Dirty tracking
    bool _dirtyRangeValid   = false;
    size_t _dirtyRangeFirst = 0;
    size_t _dirtyRangeLast  = 0;

    // Filter state
    uint32_t _filterMask = 0x1F; // All 5 types enabled by default (bits 0-4)
};

#pragma warning(pop)
