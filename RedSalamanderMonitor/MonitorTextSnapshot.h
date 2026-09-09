#pragma once

#include <cstddef>
#include <cstdint>
#include <deque>
#include <memory>
#include <string>
#include <string_view>
#include <utility>

namespace RedSalamanderMonitor
{
// Records are immutable immediately, so capture has no mutable tail to freeze or copy.
inline constexpr uint64_t kMonitorSnapshotMaxActiveTailBytes = 0u;

// One immutable logical-record block. Document and Save As snapshots share the
// same allocation; replacing or extending a record publishes a new block.
class MonitorTextBlock final
{
public:
    MonitorTextBlock() noexcept = default;
    MonitorTextBlock(std::wstring text) : _text(std::make_shared<const std::wstring>(std::move(text)))
    {
    }

    MonitorTextBlock& operator=(std::wstring text)
    {
        _text = std::make_shared<const std::wstring>(std::move(text));
        return *this;
    }

    [[nodiscard]] const std::wstring& Text() const noexcept
    {
        static const std::wstring empty;
        return _text ? *_text : empty;
    }

    [[nodiscard]] bool empty() const noexcept
    {
        return Text().empty();
    }

    [[nodiscard]] size_t size() const noexcept
    {
        return Text().size();
    }

    [[nodiscard]] wchar_t operator[](size_t index) const noexcept
    {
        return Text()[index];
    }

    [[nodiscard]] auto begin() const noexcept
    {
        return Text().begin();
    }

    [[nodiscard]] auto end() const noexcept
    {
        return Text().end();
    }

    [[nodiscard]] size_t find(std::wstring_view value, size_t offset = 0u) const noexcept
    {
        return Text().find(value, offset);
    }

    void Append(const wchar_t* text, size_t count)
    {
        std::wstring replacement(Text());
        replacement.append(text, count);
        *this = std::move(replacement);
    }

    operator const std::wstring&() const noexcept
    {
        return Text();
    }

    friend bool operator==(const MonitorTextBlock& block, std::wstring_view text) noexcept
    {
        return block.Text() == text;
    }

    friend bool operator==(std::wstring_view text, const MonitorTextBlock& block) noexcept
    {
        return block == text;
    }

private:
    std::shared_ptr<const std::wstring> _text;
};

// Immutable exact-start Save As payload. Text blocks are shared with Document;
// only this handle deque is newly allocated during capture.
struct MonitorTextSnapshot final
{
    std::deque<MonitorTextBlock> lines;
    uint64_t retainedTextBytes           = 0u;
    uint64_t copiedTextBytes             = 0u;
    uint64_t sharedBlockCount            = 0u;
    uint64_t sharedBlockBytes            = 0u;
    uint64_t activeTailCopiedBytes       = 0u;
    uint64_t peakAdditionalSnapshotBytes = 0u;
    uint64_t lockHoldUs                  = 0u;
};
} // namespace RedSalamanderMonitor
