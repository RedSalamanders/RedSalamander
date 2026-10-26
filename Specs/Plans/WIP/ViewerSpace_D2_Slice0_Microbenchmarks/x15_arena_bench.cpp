// X15 microbenchmark: ViewerSpace child-arena free list, old (sorted vector) vs new (map + set).
// Ported from Plugins/ViewerSpace/ViewerSpace.cpp at 1dcd28c (old) and 3fa79cf (new); the arena contents are
// replaced by a size counter because only block placement matters. The growth policy mirrors AddRealNodeChild.
#include <algorithm>
#include <chrono>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cmath>
#include <limits>
#include <type_traits>
#include <map>
#include <optional>
#include <random>
#include <set>
#include <vector>

namespace
{
constexpr uint32_t kMaxChildArenaSlots = 1'200'000u; // production ResourcePolicy::maxChildArenaSlots

struct OldArena
{
    struct Block
    {
        uint32_t start;
        uint32_t capacity;
    };
    size_t size = 0;
    std::vector<Block> free;

    std::optional<uint32_t> Allocate(uint32_t capacity)
    {
        size_t bestIndex      = free.size();
        uint32_t bestCapacity = std::numeric_limits<uint32_t>::max();
        for (size_t index = 0; index < free.size(); ++index)
        {
            if (free[index].capacity >= capacity && free[index].capacity < bestCapacity)
            {
                bestIndex    = index;
                bestCapacity = free[index].capacity;
            }
        }
        if (bestIndex < free.size())
        {
            Block& block         = free[bestIndex];
            const uint32_t start = block.start;
            if (block.capacity == capacity)
            {
                free.erase(free.begin() + static_cast<std::ptrdiff_t>(bestIndex));
            }
            else
            {
                block.start += capacity;
                block.capacity -= capacity;
            }
            return start;
        }
        if (size > kMaxChildArenaSlots || capacity > kMaxChildArenaSlots - size)
        {
            return std::nullopt;
        }
        const size_t previous = size;
        size += capacity;
        return static_cast<uint32_t>(previous);
    }

    void Release(uint32_t start, uint32_t capacity)
    {
        free.push_back(Block{start, capacity});
        std::sort(free.begin(), free.end(), [](const Block& l, const Block& r) { return l.start < r.start; });
        size_t write = 0;
        for (const Block& block : free)
        {
            if (write == 0)
            {
                free[write++] = block;
                continue;
            }
            Block& previous = free[write - 1];
            if (static_cast<uint64_t>(previous.start) + previous.capacity == block.start)
            {
                previous.capacity += block.capacity;
            }
            else
            {
                free[write++] = block;
            }
        }
        free.resize(write);
        while (! free.empty() && static_cast<uint64_t>(free.back().start) + free.back().capacity == size)
        {
            size = free.back().start;
            free.pop_back();
        }
    }
};

struct NewArena
{
    size_t size = 0;
    std::map<uint32_t, uint32_t> byStart;
    std::set<std::pair<uint32_t, uint32_t>> bySize;

    std::optional<uint32_t> Allocate(uint32_t capacity)
    {
        const auto bestFit = bySize.lower_bound({capacity, 0u});
        if (bestFit != bySize.end())
        {
            const uint32_t blockCapacity = bestFit->first;
            const uint32_t start         = bestFit->second;
            bySize.erase(bestFit);
            byStart.erase(start);
            if (blockCapacity > capacity)
            {
                byStart.emplace(start + capacity, blockCapacity - capacity);
                bySize.emplace(blockCapacity - capacity, start + capacity);
            }
            return start;
        }
        if (size > kMaxChildArenaSlots || capacity > kMaxChildArenaSlots - size)
        {
            return std::nullopt;
        }
        const size_t previous = size;
        size += capacity;
        return static_cast<uint32_t>(previous);
    }

    void Release(uint32_t start, uint32_t capacity)
    {
        uint32_t freeStart    = start;
        uint64_t freeCapacity = capacity;
        auto next             = byStart.lower_bound(start);
        if (next != byStart.end() && static_cast<uint64_t>(start) + capacity == next->first)
        {
            freeCapacity += next->second;
            bySize.erase({next->second, next->first});
            next = byStart.erase(next);
        }
        if (next != byStart.begin())
        {
            const auto previous = std::prev(next);
            if (static_cast<uint64_t>(previous->first) + previous->second == start)
            {
                freeStart = previous->first;
                freeCapacity += previous->second;
                bySize.erase({previous->second, previous->first});
                byStart.erase(previous);
            }
        }
        if (static_cast<uint64_t>(freeStart) + freeCapacity == size)
        {
            size = freeStart;
            return;
        }
        byStart.emplace(freeStart, static_cast<uint32_t>(freeCapacity));
        bySize.emplace(static_cast<uint32_t>(freeCapacity), freeStart);
    }
};

struct Parent
{
    uint32_t start    = 0;
    uint32_t count    = 0;
    uint32_t capacity = 0;
};

// Mirrors ViewerSpace::AddRealNodeChild. Returns false when the arena ceiling is reached.
template <typename Arena> bool AddChild(Arena& arena, Parent& parent, std::vector<uint32_t>& trace)
{
    constexpr uint32_t kInitialCapacity = 2;
    if (parent.capacity == 0)
    {
        const auto start = arena.Allocate(kInitialCapacity);
        if (! start)
        {
            return false;
        }
        trace.push_back(*start);
        parent.start    = *start;
        parent.capacity = kInitialCapacity;
    }
    if (parent.count >= parent.capacity)
    {
        const uint32_t next = std::min(parent.capacity * 2u, kMaxChildArenaSlots);
        if (static_cast<uint64_t>(parent.start) + parent.capacity == arena.size && next - parent.capacity <= kMaxChildArenaSlots - arena.size)
        {
            arena.size += next - parent.capacity; // grow the tail block in place
            parent.capacity = next;
        }
        else
        {
            const auto start = arena.Allocate(next);
            if (! start)
            {
                return false;
            }
            trace.push_back(*start);
            arena.Release(parent.start, parent.capacity);
            parent.start    = *start;
            parent.capacity = next;
        }
    }
    parent.count += 1;
    return true;
}

struct Workload
{
    size_t parents;
    std::vector<uint32_t> order; // parent index for each child, interleaved as a parallel scan would publish them
};

Workload MakeWorkload(size_t parents, size_t children, uint32_t seed)
{
    std::mt19937 rng(seed);
    // Heavy-tailed fan-out: most folders are small, a few are large (as on a real volume).
    std::vector<double> weights(parents);
    std::uniform_real_distribution<double> u(0.0001, 1.0);
    for (double& w : weights)
    {
        w = 1.0 / std::pow(u(rng), 1.3);
    }
    std::discrete_distribution<size_t> pick(weights.begin(), weights.end());
    Workload workload{parents, {}};
    workload.order.reserve(children);
    for (size_t i = 0; i < children; ++i)
    {
        workload.order.push_back(static_cast<uint32_t>(pick(rng)));
    }
    return workload;
}

template <typename Arena> double Run(const Workload& workload, std::vector<uint32_t>& trace, size_t& added, size_t& peakFree)
{
    Arena arena;
    std::vector<Parent> parents(workload.parents);
    trace.clear();
    added    = 0;
    peakFree = 0;
    const auto started = std::chrono::steady_clock::now();
    for (uint32_t parentIndex : workload.order)
    {
        if (! AddChild(arena, parents[parentIndex], trace))
        {
            break;
        }
        ++added;
        if constexpr (std::is_same_v<Arena, OldArena>)
        {
            peakFree = std::max(peakFree, arena.free.size());
        }
        else
        {
            peakFree = std::max(peakFree, arena.byStart.size());
        }
    }
    return std::chrono::duration<double, std::milli>(std::chrono::steady_clock::now() - started).count();
}
} // namespace

int main()
{
    struct Case
    {
        size_t parents;
        size_t children;
    };
    // The largest case uses the production retained-child-reference ceiling (300,000).
    const Case cases[] = {{2'000, 10'000}, {10'000, 50'000}, {30'000, 150'000}, {100'000, 300'000}};
    const bool roundRobin = std::getenv("X15_ROUND_ROBIN") != nullptr;
    std::printf("parents,children,added,old_ms_median,new_ms_median,speedup,peak_free_blocks,identical_placement\n");
    for (const Case& c : cases)
    {
        Workload workload = MakeWorkload(c.parents, c.children, 20260927u);
        if (std::getenv("X15_FRAGMENT") != nullptr)
        {
            // Pathological: every parent gets 3 children (round-robin), then only even parents keep growing, leaving an
            // isolated 4-slot hole beside every odd parent that later, larger allocations cannot reuse.
            workload.order.clear();
            for (int round = 0; round < 3; ++round)
                for (size_t p = 0; p < c.parents; ++p)
                    workload.order.push_back(static_cast<uint32_t>(p));
            while (workload.order.size() < c.children)
                for (size_t p = 0; p < c.parents && workload.order.size() < c.children; p += 2)
                    workload.order.push_back(static_cast<uint32_t>(p));
        }
        else if (roundRobin)
        {
            // Adversarial: every parent grows in lockstep, so released blocks are interleaved and rarely coalesce.
            for (size_t i = 0; i < workload.order.size(); ++i)
            {
                workload.order[i] = static_cast<uint32_t>(i % c.parents);
            }
        }
        std::vector<double> oldTimes, newTimes;
        std::vector<uint32_t> oldTrace, newTrace;
        size_t oldAdded = 0, newAdded = 0, oldPeak = 0, newPeak = 0;
        for (int rep = 0; rep < 5; ++rep)
        {
            oldTimes.push_back(Run<OldArena>(workload, oldTrace, oldAdded, oldPeak));
            newTimes.push_back(Run<NewArena>(workload, newTrace, newAdded, newPeak));
        }
        std::sort(oldTimes.begin(), oldTimes.end());
        std::sort(newTimes.begin(), newTimes.end());
        const bool identical = oldTrace == newTrace && oldAdded == newAdded && oldPeak == newPeak;
        std::printf("%zu,%zu,%zu,%.2f,%.2f,%.1fx,%zu,%s\n",
                    c.parents,
                    c.children,
                    oldAdded,
                    oldTimes[2],
                    newTimes[2],
                    oldTimes[2] / newTimes[2],
                    oldPeak,
                    identical ? "yes" : "NO");
    }
    return 0;
}
