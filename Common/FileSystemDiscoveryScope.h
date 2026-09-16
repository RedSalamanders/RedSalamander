#pragma once

#include <algorithm>
#include <cstdint>
#include <limits>
#include <mutex>

#include "PlugInterfaces/FileSystem.h"

namespace Common::FileOperations
{
// One governed provider call is one discovery scope. Every provider that owns a transfer or a
// removal answers the same three questions for its caller: how much work this call turned out to
// involve, whether that total is still growing, and when it stopped growing. The payload is defined
// by the ABI; what providers kept re-deriving was the bookkeeping around it, so it lives here.
//
// The rules this owns, and that a local copy would get wrong:
//
//   - Counters are cumulative for the whole governed call, never per item and never deltas.
//   - Exactly one report carries traversalClosed. Several closures in one call are a contract
//     violation, because the first would declare the caller's totals final while later work is
//     still being discovered.
//   - The closure happens on every exit, including error, partial and cancellation. A scope the
//     caller can never close is worse than one that closes on the totals the call actually reached.
//   - Reports are serialized. A provider that advances a total under a lock and emits it afterwards
//     lets two workers deliver their snapshots out of order, and the caller sees a cumulative
//     stream that goes backwards.
//   - An unknown size is not a zero size. A provider that cannot learn a byte total still reports
//     the item, and simply contributes no bytes.
//
// The caller owns what a "file", a "directory" and "queued work" mean for its own namespace, and
// owns cancellation, progress and bandwidth. This type never talks to the network or the disk.
class DiscoveryScope final
{
public:
    // options may be null, or carry no operation control; the scope then becomes inert and every
    // member is still safe to call. That keeps provider code free of null checks at each site.
    explicit DiscoveryScope(const FileSystemOptions* options) noexcept
        : _control(options != nullptr ? options->operationControl : nullptr),
          _cookie(options != nullptr ? options->operationControlCookie : nullptr)
    {
    }

    DiscoveryScope(const DiscoveryScope&)            = delete;
    DiscoveryScope(DiscoveryScope&&)                 = delete;
    DiscoveryScope& operator=(const DiscoveryScope&) = delete;
    DiscoveryScope& operator=(DiscoveryScope&&)      = delete;

    // Closing in the destructor is the point: a provider with many early returns cannot forget it.
    ~DiscoveryScope()
    {
        Close();
    }

    [[nodiscard]] bool IsActive() const noexcept
    {
        return _control != nullptr;
    }

    // One discovered file whose size is known. Reports the new cumulative totals.
    HRESULT AddFile(uint64_t sizeBytes, uint32_t queuedItems = 0u) noexcept
    {
        return Add(sizeBytes, 1u, 0u, queuedItems);
    }

    // One discovered file whose size the provider cannot learn without paying for another round
    // trip. The item still counts; inventing a zero byte total would be a lie the caller cannot
    // distinguish from a real empty file.
    HRESULT AddFileWithUnknownSize(uint32_t queuedItems = 0u) noexcept
    {
        return Add(0u, 1u, 0u, queuedItems);
    }

    // One discovered directory. A Move that relocates a whole subtree by renaming it discovers the
    // directory and none of its descendants, which is the honest answer for that call.
    HRESULT AddDirectory(uint32_t queuedItems = 0u) noexcept
    {
        return Add(0u, 0u, 1u, queuedItems);
    }

    // Adopt a total the provider computed in one pass, such as a plan built before any mutation.
    HRESULT AddPlan(uint64_t sizeBytes, uint64_t files, uint64_t directories, uint32_t queuedItems = 0u) noexcept
    {
        return Add(sizeBytes, files, directories, queuedItems);
    }

    // Re-publish the current totals without adding anything, for a provider that wants to refresh
    // the caller's queue depth while it works.
    HRESULT Republish(uint32_t queuedItems = 0u) noexcept
    {
        return Add(0u, 0u, 0u, queuedItems);
    }

    // The single closure. Idempotent, so an explicit early close and the destructor cannot both
    // emit one. Failures are deliberately swallowed: a caller that refuses a discovery report has
    // already been told by the report that returned the failure, and a destructor cannot report.
    void Close() noexcept
    {
        std::scoped_lock lock(_mutex);
        if (_control == nullptr || _closed)
        {
            return;
        }
        _closed = true;
        static_cast<void>(EmitLocked(true));
    }

private:
    [[nodiscard]] static uint64_t AddSaturating(uint64_t left, uint64_t right) noexcept
    {
        return left > (std::numeric_limits<uint64_t>::max)() - right ? (std::numeric_limits<uint64_t>::max)() : left + right;
    }

    HRESULT Add(uint64_t sizeBytes, uint64_t files, uint64_t directories, uint32_t queuedItems) noexcept
    {
        std::scoped_lock lock(_mutex);
        if (_control == nullptr || _closed)
        {
            return S_OK;
        }
        _bytes       = AddSaturating(_bytes, sizeBytes);
        _files       = AddSaturating(_files, files);
        _directories = AddSaturating(_directories, directories);
        _queuedItems = (std::max)(_queuedItems, queuedItems);
        return EmitLocked(false);
    }

    // Emitted under _mutex on purpose. Serializing the callback is what keeps the caller's
    // cumulative stream monotonic when several workers report concurrently.
    HRESULT EmitLocked(bool traversalClosed) noexcept
    {
        FileSystemDiscoveryProgress progress{};
        progress.sizeBytes             = sizeof(progress);
        progress.discoveredBytes       = _bytes;
        progress.discoveredFiles       = _files;
        progress.discoveredDirectories = _directories;
        progress.queuedItems           = _queuedItems;
        progress.traversalClosed       = traversalClosed ? TRUE : FALSE;
        return _control->FileSystemReportDiscoveryProgress(&progress, _cookie);
    }

    IFileSystemOperationControl* _control = nullptr;
    void* _cookie                         = nullptr;
    std::mutex _mutex;
    uint64_t _bytes       = 0;
    uint64_t _files       = 0;
    uint64_t _directories = 0;
    uint32_t _queuedItems = 0;
    bool _closed          = false;
};
} // namespace Common::FileOperations
