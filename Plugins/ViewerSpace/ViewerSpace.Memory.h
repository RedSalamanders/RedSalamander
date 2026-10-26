#pragma once

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <memory_resource>

namespace ViewerSpaceDetail
{
// Counts actual upstream PMR block allocations, including pool/arena spare capacity.
// Single-owner observation/admission data, not an opaque-provider process cap. A model may move to
// the deferred-disposal callback only after UI inspection stops, so accounting has no concurrent readers.
class CountingMemoryResource final : public std::pmr::memory_resource
{
public:
    CountingMemoryResource()                                         = default;
    ~CountingMemoryResource() override                               = default;
    CountingMemoryResource(const CountingMemoryResource&)            = delete;
    CountingMemoryResource& operator=(const CountingMemoryResource&) = delete;
    CountingMemoryResource(CountingMemoryResource&&)                 = delete;
    CountingMemoryResource& operator=(CountingMemoryResource&&)      = delete;
    [[nodiscard]] uint64_t AllocatedBytes() const noexcept
    {
        return _bytes;
    }
    [[nodiscard]] uint64_t PeakBytes() const noexcept
    {
        return _peak;
    }

private:
    void* do_allocate(size_t bytes, size_t alignment) override
    {
        void* allocation = std::pmr::new_delete_resource()->allocate(bytes, alignment);
        _bytes += static_cast<uint64_t>(bytes);
        if (_bytes > _peak)
        {
            _peak = _bytes;
        }
        return allocation;
    }
    void do_deallocate(void* allocation, size_t bytes, size_t alignment) override
    {
        std::pmr::new_delete_resource()->deallocate(allocation, bytes, alignment);
        _bytes -= static_cast<uint64_t>(bytes);
    }
    bool do_is_equal(const std::pmr::memory_resource& other) const noexcept override
    {
        return this == &other;
    }
    uint64_t _bytes = 0u;
    uint64_t _peak  = 0u;
};

// Scanner containers are serialized by their frontier mutex, but the UI observes their aggregate
// allocation charge while retired generations can still run. Unlike model-local accounting above,
// this resource therefore uses instance-owned atomics and never resets them at generation start.
class ConcurrentCountingMemoryResource final : public std::pmr::memory_resource
{
public:
    ConcurrentCountingMemoryResource(std::atomic_uint64_t& bytes, std::atomic_uint64_t& peak) noexcept : _bytes(bytes), _peak(peak)
    {
    }
    ~ConcurrentCountingMemoryResource() override                                         = default;
    ConcurrentCountingMemoryResource(const ConcurrentCountingMemoryResource&)            = delete;
    ConcurrentCountingMemoryResource& operator=(const ConcurrentCountingMemoryResource&) = delete;
    ConcurrentCountingMemoryResource(ConcurrentCountingMemoryResource&&)                 = delete;
    ConcurrentCountingMemoryResource& operator=(ConcurrentCountingMemoryResource&&)      = delete;

private:
    void* do_allocate(size_t bytes, size_t alignment) override
    {
        void* allocation       = std::pmr::new_delete_resource()->allocate(bytes, alignment);
        const uint64_t current = _bytes.fetch_add(static_cast<uint64_t>(bytes), std::memory_order_relaxed) + static_cast<uint64_t>(bytes);
        uint64_t peak          = _peak.load(std::memory_order_relaxed);
        while (current > peak && ! _peak.compare_exchange_weak(peak, current, std::memory_order_relaxed))
        {
        }
        return allocation;
    }
    void do_deallocate(void* allocation, size_t bytes, size_t alignment) override
    {
        std::pmr::new_delete_resource()->deallocate(allocation, bytes, alignment);
        _bytes.fetch_sub(static_cast<uint64_t>(bytes), std::memory_order_relaxed);
    }
    bool do_is_equal(const std::pmr::memory_resource& other) const noexcept override
    {
        return this == &other;
    }
    std::atomic_uint64_t& _bytes;
    std::atomic_uint64_t& _peak;
};
} // namespace ViewerSpaceDetail
