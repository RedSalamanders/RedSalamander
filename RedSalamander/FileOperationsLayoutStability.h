#pragma once

#include <cstddef>
#include <cstdint>

// I63 slice F: the layout-stability rules of the File Operations popup ("Layout stability" in
// Specs/UI/UI_FileOperationsPopup.md). Small state machines over a monotonic millisecond tick. The
// popup owns one set per card, feeds each the value the current frame needs, and lays out with the
// value it returns. Nothing here measures or paints, so the rules are checked at compile time below.
namespace FileOperationsLayoutStability
{
inline constexpr uint64_t kReservationHoldMs = 2'000u; // a reserved row, a stacked pair or a note outlives its need by this long
inline constexpr uint64_t kNumericTextHoldMs = 500u;   // rate and ETA text change at most twice per second
inline constexpr uint64_t kFileNameHoldMs    = 250u;   // a displayed file name stays at least this long
inline constexpr float kUnstackSpareDip      = 16.0f;  // a stacked pair returns to one line only with this much room to spare

[[nodiscard]] constexpr bool Elapsed(uint64_t sinceMs, uint64_t nowMs, uint64_t durationMs) noexcept
{
    return nowMs >= sinceMs && nowMs - sinceMs >= durationMs;
}

// The rows a card reserves for a varying row count (active stream rows): the high-water mark of the
// rows it needed, released one row at a time once the top row has gone unused for the hold.
class ReservedRows final
{
public:
    [[nodiscard]] constexpr size_t Update(size_t needed, uint64_t nowMs) noexcept
    {
        if (needed >= _reserved)
        {
            _reserved     = needed;
            _lastNeededMs = nowMs;
        }
        else if (Elapsed(_lastNeededMs, nowMs, kReservationHoldMs))
        {
            --_reserved;
            _lastNeededMs = nowMs;
        }
        return _reserved;
    }

    constexpr void Reset() noexcept
    {
        _reserved     = 0u;
        _lastNeededMs = 0u;
    }

private:
    size_t _reserved       = 0u;
    uint64_t _lastNeededMs = 0u;
};

// A layout condition that holds from the moment it is wanted until it has gone unwanted for the
// hold: a stacked fact pair, a conditional note row.
class HeldCondition final
{
public:
    [[nodiscard]] constexpr bool Update(bool wanted, uint64_t nowMs) noexcept
    {
        if (wanted)
        {
            _held         = true;
            _lastWantedMs = nowMs;
        }
        else if (_held && Elapsed(_lastWantedMs, nowMs, kReservationHoldMs))
        {
            _held = false;
        }
        return _held;
    }

    constexpr void Reset() noexcept
    {
        _held         = false;
        _lastWantedMs = 0u;
    }

private:
    bool _held             = false;
    uint64_t _lastWantedMs = 0u;
};

// Whether a two-fact row needs two lines. It stacks as soon as the facts do not fit; once stacked,
// it counts as fitting again only with spare room, so a width that hovers at the boundary cannot
// flip the row on every frame. Feed the result to a HeldCondition.
[[nodiscard]] constexpr bool PairWantsStack(float neededWidthDip, float availableWidthDip, bool currentlyStacked) noexcept
{
    return neededWidthDip + (currentlyStacked ? kUnstackSpareDip : 0.0f) > availableWidthDip;
}

// When a displayed text may be replaced by a newer value: at once when nothing is shown, otherwise
// only after the shown text has been displayed for the minimum time. The caller keeps the shown
// text and the tick it appeared at.
[[nodiscard]] constexpr bool MayReplaceShownText(bool nothingShown, uint64_t shownSinceMs, uint64_t nowMs, uint64_t minimumMs) noexcept
{
    return nothingShown || Elapsed(shownSinceMs, nowMs, minimumMs);
}

// The card list's scrollbar is decided for the viewport the window is about to have. While an
// automatic window resize is pending or running, the height it will add (negative: remove) counts as
// already there: a bar that the resize is about to make unnecessary never appears, and a window that
// cannot grow enough shows the bar at once. Scrolling is never held back by a resize: a scroll request
// made while the content overflows today's viewport (the wheel, a scrollbar message, revealing the
// focused control) shows the bar for as long as the content overflows.
[[nodiscard]] constexpr bool ScrollBarNeeded(float contentHeight, float viewportHeight, float pendingResizeDelta, bool scrollRequested) noexcept
{
    return contentHeight > viewportHeight + pendingResizeDelta || (scrollRequested && contentHeight > viewportHeight);
}

// I63 F1: stream rows. A transfer stream keeps one row slot for its lifetime, and a new stream takes the lowest slot no
// other stream holds. A stream holds its slot while it runs and, once it ended, for the linger in which its finished row
// is still painted. The engine assigns the slots with these two rules; the popup paints each stream in its slot and
// decides which streams are visible with StreamHoldsRow, so both sides agree on when a row is free.
inline constexpr uint64_t kStreamRowLingerMs = 300u;
inline constexpr size_t kMaxStreamSlots      = 32u; // the width of the occupancy mask below; the engine uses sixteen slots

// A stream with an explicit lifecycle runs until its terminal report. A stream known only from byte reports runs while its
// size is unknown or its bytes are incomplete. Either one lingers for kStreamRowLingerMs after its last report once it ended;
// a tick earlier than that report does not linger.
[[nodiscard]] constexpr bool StreamHoldsRow(
    bool explicitLifecycle, bool terminal, uint64_t totalBytes, uint64_t completedBytes, uint64_t lastUpdateMs, uint64_t nowMs) noexcept
{
    const bool lingering = lastUpdateMs != 0u && nowMs >= lastUpdateMs && nowMs - lastUpdateMs <= kStreamRowLingerMs;
    if (explicitLifecycle)
        return ! terminal || lingering;
    return totalBytes == 0u || completedBytes < totalBytes || lingering;
}

// The lowest slot below slotCount whose bit is clear in occupied; slotCount when every slot is taken.
[[nodiscard]] constexpr size_t LowestFreeSlot(uint32_t occupied, size_t slotCount) noexcept
{
    for (size_t slot = 0u; slot < slotCount && slot < kMaxStreamSlots; ++slot)
        if ((occupied & (uint32_t{1u} << slot)) == 0u)
            return slot;
    return slotCount;
}

// I63 F4: what FileOps.Popup.CardHeightChanges counts. A card's first measurement is not a change; every new target
// height after it is one, whether it is then eased or snapped.
[[nodiscard]] constexpr bool CardHeightChanged(bool measuredBefore, double previousTarget, double nextTarget) noexcept
{
    return measuredBefore && previousTarget != nextTarget;
}

namespace Contract
{
// Stream rows: growth is immediate, release is one row per hold, and renewed need restarts the hold.
[[nodiscard]] constexpr bool ReservedRowsHold() noexcept
{
    ReservedRows rows;
    bool ok = rows.Update(1u, 1'000u) == 1u;
    ok      = ok && rows.Update(4u, 1'100u) == 4u; // grows at once
    ok      = ok && rows.Update(1u, 1'200u) == 4u; // keeps the high-water mark
    ok      = ok && rows.Update(1u, 3'099u) == 4u; // until the hold has passed
    ok      = ok && rows.Update(1u, 3'100u) == 3u; // then one row
    ok      = ok && rows.Update(1u, 3'200u) == 3u; // and the next row waits its own hold
    ok      = ok && rows.Update(3u, 5'000u) == 3u; // a renewed need keeps the row
    ok      = ok && rows.Update(1u, 6'999u) == 3u;
    ok      = ok && rows.Update(1u, 7'000u) == 2u;
    ok      = ok && rows.Update(1u, 9'000u) == 1u;
    ok      = ok && rows.Update(1u, 20'000u) == 1u; // never below the need
    ok      = ok && rows.Update(0u, 20'001u) == 1u;
    return ok;
}
static_assert(ReservedRowsHold());

[[nodiscard]] constexpr bool HeldConditionHolds() noexcept
{
    HeldCondition held;
    bool ok = ! held.Update(false, 100u);
    ok      = ok && held.Update(true, 200u);
    ok      = ok && held.Update(false, 300u);   // stays through a short gap
    ok      = ok && held.Update(true, 2'000u);  // renewed
    ok      = ok && held.Update(false, 3'999u); // the hold counts from the last time it was wanted
    ok      = ok && ! held.Update(false, 4'000u);
    ok      = ok && ! held.Update(false, 4'001u);
    return ok;
}
static_assert(HeldConditionHolds());

// A pair at the boundary stacks at once and needs spare room to return to one line.
static_assert(! PairWantsStack(200.0f, 200.0f, false));
static_assert(PairWantsStack(200.5f, 200.0f, false));
static_assert(PairWantsStack(190.0f, 200.0f, true));
static_assert(! PairWantsStack(184.0f, 200.0f, true));

static_assert(MayReplaceShownText(true, 0u, 0u, kNumericTextHoldMs));
static_assert(! MayReplaceShownText(false, 1'000u, 1'499u, kNumericTextHoldMs));
static_assert(MayReplaceShownText(false, 1'000u, 1'500u, kNumericTextHoldMs));
static_assert(! MayReplaceShownText(false, 1'000u, 900u, kNumericTextHoldMs)); // a tick that went backwards replaces nothing

static_assert(! ScrollBarNeeded(400.0f, 400.0f, 0.0f, false));
static_assert(ScrollBarNeeded(401.0f, 400.0f, 0.0f, false));
static_assert(! ScrollBarNeeded(436.0f, 400.0f, 36.0f, false)); // the window is about to grow by the row that was added: no bar
static_assert(ScrollBarNeeded(436.0f, 400.0f, 36.0f, true));    // unless something asked to scroll meanwhile
static_assert(! ScrollBarNeeded(400.0f, 400.0f, 36.0f, true));  // a request cannot show a bar for content that fits
static_assert(ScrollBarNeeded(900.0f, 400.0f, 200.0f, false));  // it cannot grow enough: the bar is needed now
static_assert(ScrollBarNeeded(400.0f, 400.0f, -36.0f, false));  // it is about to shrink below the content: the bar stays

// Stream rows: a slot is held while its stream runs and for the linger after it ended.
static_assert(StreamHoldsRow(false, false, 0u, 0u, 1'000u, 50'000u));    // size unknown: running, however quiet
static_assert(StreamHoldsRow(false, false, 100u, 10u, 1'000u, 50'000u)); // bytes incomplete: running
static_assert(StreamHoldsRow(false, false, 100u, 100u, 1'000u, 1'300u)); // complete: lingers 300 ms
static_assert(! StreamHoldsRow(false, false, 100u, 100u, 1'000u, 1'301u));
static_assert(StreamHoldsRow(true, false, 100u, 100u, 1'000u, 50'000u)); // an explicit lifecycle runs until its terminal report
static_assert(StreamHoldsRow(true, true, 0u, 0u, 1'000u, 1'300u));
static_assert(! StreamHoldsRow(true, true, 0u, 0u, 1'000u, 1'301u));
static_assert(! StreamHoldsRow(true, true, 0u, 0u, 1'000u, 900u)); // a tick that went backwards does not linger
static_assert(LowestFreeSlot(0u, 16u) == 0u);
static_assert(LowestFreeSlot(0b1011u, 16u) == 2u);  // a slot freed below live streams is reused first
static_assert(LowestFreeSlot(0xFFFFu, 16u) == 16u); // every slot taken
static_assert(LowestFreeSlot(0xFFFFu, 17u) == 16u);
static_assert(LowestFreeSlot(0u, 0u) == 0u);

static_assert(CardHeightChanged(true, 100.0, 120.0));
static_assert(! CardHeightChanged(true, 100.0, 100.0));
static_assert(! CardHeightChanged(false, 0.0, 120.0)); // the first measurement
} // namespace Contract
} // namespace FileOperationsLayoutStability
