#include "FileSystem.Internal.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstring>
#include <deque>
#include <filesystem>
#include <functional>
#include <limits>
#include <memory>
#include <mutex>
#include <ranges>
#include <span>
#include <thread>
#include <utility>

#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <winioctl.h>

using namespace FileSystemInternal;

namespace FileSystemInternal
{
// Module anchor for AcquireModuleReferenceFromAddress — keeps the DLL loaded while worker threads are running.
const int kFileSystemModuleAnchor = 0;

class SharedFileOpsJobScheduler final
{
public:
    SharedFileOpsJobScheduler() = default;
    ~SharedFileOpsJobScheduler() noexcept
    {
        ShutdownAndJoin();
    }

    SharedFileOpsJobScheduler(const SharedFileOpsJobScheduler&)            = delete;
    SharedFileOpsJobScheduler(SharedFileOpsJobScheduler&&)                 = delete;
    SharedFileOpsJobScheduler& operator=(const SharedFileOpsJobScheduler&) = delete;
    SharedFileOpsJobScheduler& operator=(SharedFileOpsJobScheduler&&)      = delete;

    struct Job final
    {
        Job() noexcept = default;

        Job(const Job&)            = delete;
        Job(Job&&)                 = delete;
        Job& operator=(const Job&) = delete;
        Job& operator=(Job&&)      = delete;

        std::function<void(size_t, uint64_t)> processIndex;
        size_t totalItems           = 0;
        unsigned int maxConcurrency = 1;

        // Protected by the scheduler mutex.
        size_t nextIndex      = 0;
        unsigned int inFlight = 0;

        std::atomic<bool> done{false};
        std::mutex doneMutex;
        std::condition_variable doneCv;
    };

    using JobPtr = std::shared_ptr<Job>;

    JobPtr StartJob(unsigned int maxConcurrency, size_t totalItems, std::function<void(size_t, uint64_t)> processIndex)
    {
        auto job            = std::make_shared<Job>();
        job->totalItems     = totalItems;
        job->processIndex   = std::move(processIndex);
        job->maxConcurrency = std::max(1u, maxConcurrency);
        if (job->totalItems > 0)
        {
            job->maxConcurrency =
                std::min<unsigned int>(job->maxConcurrency, static_cast<unsigned int>((std::min)(job->totalItems, static_cast<size_t>(UINT_MAX))));
        }

        ensureWorkers();

        if (_workers.empty())
        {
            if (job->processIndex)
            {
                for (size_t i = 0; i < job->totalItems; ++i)
                {
                    job->processIndex(i, 0);
                }
            }

            finishJob(*job);
            return job;
        }

        {
            std::scoped_lock lock(_mutex);
            _jobs.push_back(job);
        }

        _cv.notify_all();
        return job;
    }

    void WaitJob(const JobPtr& job) noexcept
    {
        if (! job)
        {
            return;
        }

        if (IsWorkerThread())
        {
            // Avoid deadlocks when a file operation recursively starts parallel work from within a worker.
            while (! job->done.load(std::memory_order_acquire))
            {
                JobPtr dequeued;
                size_t index = 0;
                {
                    std::unique_lock lock(_mutex);
                    _cv.wait(lock, [&]() noexcept { return job->done.load(std::memory_order_acquire) || hasSchedulableWorkLocked(); });
                    if (job->done.load(std::memory_order_acquire))
                    {
                        break;
                    }

                    if (! tryDequeueWorkLocked(dequeued, index))
                    {
                        continue;
                    }
                }

                executeWorkItem(std::move(dequeued), index, tls_workerStreamId);
            }

            return;
        }

        std::unique_lock lock(job->doneMutex);
        job->doneCv.wait(lock, [&]() noexcept { return job->done.load(std::memory_order_acquire); });
    }

    void Shutdown() noexcept
    {
        {
            std::scoped_lock lock(_initMutex);
            if (! _initialized)
            {
                return;
            }

            for (std::jthread& worker : _workers)
            {
                worker.request_stop();
            }
        }

        // Ensure any thread blocked in WaitJob can proceed during teardown.
        {
            std::scoped_lock lock(_mutex);
            for (const JobPtr& job : _jobs)
            {
                if (job)
                {
                    finishJob(*job);
                }
            }
            _jobs.clear();
            _rrCursor = 0;
        }

        _cv.notify_all();
    }

    void ShutdownAndJoin() noexcept
    {
        std::vector<std::jthread> workers;
        {
            std::scoped_lock lock(_initMutex);
            if (_initialized)
            {
                for (std::jthread& worker : _workers)
                {
                    worker.request_stop();
                }

                workers      = std::move(_workers);
                _initialized = false;
            }
        }

        // Ensure any thread blocked in WaitJob can proceed during teardown.
        {
            std::scoped_lock lock(_mutex);
            for (const JobPtr& job : _jobs)
            {
                if (job)
                {
                    finishJob(*job);
                }
            }
            _jobs.clear();
            _rrCursor = 0;
        }

        _cv.notify_all();

        // 'workers' destructs here (joining the worker threads) outside any locks.
    }

    [[nodiscard]] bool IsWorkerThread() const noexcept
    {
        return tls_scheduler == this;
    }

    [[nodiscard]] uint64_t CurrentWorkerStreamId() const noexcept
    {
        return IsWorkerThread() ? tls_workerStreamId : 0;
    }

    [[nodiscard]] bool EnsureWorkersAvailable()
    {
        ensureWorkers();
        return ! _workers.empty();
    }

private:
    static inline thread_local const SharedFileOpsJobScheduler* tls_scheduler = nullptr;
    static inline thread_local uint64_t tls_workerStreamId                    = 0;

    void ensureWorkers()
    {
        std::scoped_lock lock(_initMutex);
        if (_initialized)
        {
            return;
        }

        unsigned int workerCount = std::thread::hardware_concurrency();
        if (workerCount == 0)
        {
            workerCount = 4;
        }

        constexpr unsigned int kMaxWorkers = 16u;
        workerCount                        = std::max(1u, std::min(workerCount, kMaxWorkers));

        _workers.reserve(workerCount);
        for (unsigned int i = 0; i < workerCount; ++i)
        {
            // Pin the module so the DLL cannot be unloaded while worker threads are running.
            wil::unique_hmodule modulePin = AcquireModuleReferenceFromAddress(&kFileSystemModuleAnchor);
            if (! modulePin)
            {
                Debug::Error(L"FileSystem: Failed to pin module for job scheduler worker thread {}", i);
                break;
            }

            try
            {
                _workers.emplace_back([this, i, pin = std::move(modulePin)](std::stop_token stopToken) noexcept
                {
                    static_cast<void>(pin); // prevent [[maybe_unused]] — released on thread exit
                    workerMain(stopToken, static_cast<uint64_t>(i));
                });
            }
            catch (const std::system_error&)
            {
                // Module pin released via RAII if thread creation fails.
                break;
            }
        }

        _initialized = true;
    }

    void finishJob(Job& job) noexcept
    {
        {
            std::scoped_lock lock(job.doneMutex);
            job.done.store(true, std::memory_order_release);
        }
        job.doneCv.notify_all();
    }

    void cleanupJobsLocked() noexcept
    {
        size_t write = 0;
        for (size_t read = 0; read < _jobs.size(); ++read)
        {
            const JobPtr& job = _jobs[read];
            if (! job)
            {
                continue;
            }

            const bool finished = job->nextIndex >= job->totalItems;
            if (finished && job->inFlight == 0)
            {
                finishJob(*job);
                continue;
            }

            if (write != read)
            {
                _jobs[write] = job;
            }
            ++write;
        }

        if (write < _jobs.size())
        {
            _jobs.resize(write);
        }

        if (_rrCursor >= _jobs.size())
        {
            _rrCursor = 0;
        }
    }

    [[nodiscard]] bool hasSchedulableWorkLocked() noexcept
    {
        cleanupJobsLocked();

        for (const JobPtr& job : _jobs)
        {
            if (! job)
            {
                continue;
            }

            if (job->inFlight >= job->maxConcurrency)
            {
                continue;
            }

            if (job->nextIndex >= job->totalItems)
            {
                continue;
            }

            return true;
        }

        return false;
    }

    [[nodiscard]] bool tryDequeueWorkLocked(JobPtr& outJob, size_t& outIndex) noexcept
    {
        const size_t jobCount = _jobs.size();
        if (jobCount == 0)
        {
            return false;
        }

        const size_t start = _rrCursor % jobCount;
        for (size_t attempt = 0; attempt < jobCount; ++attempt)
        {
            const size_t idx = (start + attempt) % jobCount;
            JobPtr& job      = _jobs[idx];
            if (! job)
            {
                continue;
            }

            if (job->inFlight >= job->maxConcurrency)
            {
                continue;
            }

            if (job->nextIndex >= job->totalItems)
            {
                continue;
            }

            outJob   = job;
            outIndex = job->nextIndex;
            job->nextIndex += 1;
            job->inFlight += 1;

            _rrCursor = (idx + 1u) % jobCount;
            return true;
        }

        return false;
    }

    void executeWorkItem(JobPtr job, size_t index, uint64_t streamId) noexcept
    {
        if (job && job->processIndex)
        {
            job->processIndex(index, streamId);
        }

        {
            std::scoped_lock lock(_mutex);
            if (job && job->inFlight > 0)
            {
                job->inFlight -= 1;
            }
            cleanupJobsLocked();
        }

        _cv.notify_all();
    }

    void workerMain(std::stop_token stopToken, uint64_t streamId) noexcept
    {
        tls_scheduler      = this;
        tls_workerStreamId = streamId;

        [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();

        while (! stopToken.stop_requested())
        {
            JobPtr job;
            size_t index = 0;

            {
                std::unique_lock lock(_mutex);
                _cv.wait(lock, [&]() noexcept { return stopToken.stop_requested() || hasSchedulableWorkLocked(); });
                if (stopToken.stop_requested())
                {
                    break;
                }

                cleanupJobsLocked();
                if (! tryDequeueWorkLocked(job, index))
                {
                    continue;
                }
            }

            executeWorkItem(std::move(job), index, streamId);
        }

        tls_scheduler      = nullptr;
        tls_workerStreamId = 0;
    }

private:
    std::mutex _mutex;
    std::condition_variable _cv;
    std::vector<JobPtr> _jobs;
    size_t _rrCursor = 0;

    std::mutex _initMutex;
    bool _initialized = false;
    std::vector<std::jthread> _workers;
};

SharedFileOpsJobScheduler& GetSharedFileOpsJobScheduler() noexcept
{
    static SharedFileOpsJobScheduler scheduler;
    return scheduler;
}

void ShutdownSharedFileOpsJobScheduler() noexcept
{
    GetSharedFileOpsJobScheduler().ShutdownAndJoin();
}
} // namespace FileSystemInternal

namespace
{
#pragma warning(push)
// C4625 (copy ctor deleted), C4626 (copy assign deleted)
#pragma warning(disable : 4625 4626)
constexpr size_t kMaxBandwidthThrottleWorkers = 16u;

struct ParallelOperationState
{
    ParallelOperationState() noexcept                                = default;
    ParallelOperationState(const ParallelOperationState&)            = delete;
    ParallelOperationState& operator=(const ParallelOperationState&) = delete;
    ParallelOperationState(ParallelOperationState&&)                 = delete;
    ParallelOperationState& operator=(ParallelOperationState&&)      = delete;

    std::atomic<uint64_t> completedBytes{0};
    std::atomic<unsigned long> completedItems{0};
    std::atomic<uint64_t> bandwidthLimitBytesPerSecond{0};
    std::atomic<unsigned int> copyMoveTransferLimit{0};

    std::mutex copyMoveTransferMutex;
    std::condition_variable copyMoveTransferCv;
    unsigned int activeCopyMoveTransfers = 0;

    struct BandwidthThrottleState
    {
        BandwidthThrottleState() noexcept                                = default;
        BandwidthThrottleState(const BandwidthThrottleState&)            = delete;
        BandwidthThrottleState& operator=(const BandwidthThrottleState&) = delete;
        BandwidthThrottleState(BandwidthThrottleState&&)                 = delete;
        BandwidthThrottleState& operator=(BandwidthThrottleState&&)      = delete;

        struct WorkerState
        {
            uint64_t configuredLimitBytesPerSecond = 0;
            int64_t availableBytes                 = 0;
            ULONGLONG lastRefillTick               = 0;
            ULONGLONG lastActiveTick               = 0;
        };

        std::mutex mutex;
        uint64_t configuredLimitBytesPerSecond = 0;
        int64_t availableBytes                 = 0;
        ULONGLONG lastRefillTick               = 0;
        std::array<WorkerState, kMaxBandwidthThrottleWorkers> workerStates{};
    } bandwidthThrottle;

    ULONGLONG startTick = 0;
    std::mutex callbackMutex;
    ULONGLONG lastProgressReportTick = 0;
    std::atomic<ULONGLONG> lastCancelCheckTick{0};

    std::atomic<bool> cancelRequested{false};
    std::atomic<bool> stopOnErrorRequested{false};
    std::atomic<HRESULT> firstError{S_OK};
    std::atomic<bool> hadFailure{false};
};

enum class BandwidthThrottleWorkerMode : uint8_t
{
    SharedOnly = 0,
    PerWorkerSubBudget,
};

struct OperationContext
{
    FileSystemOperation type      = FILESYSTEM_COPY;
    IFileSystemCallback* callback = nullptr;
    void* callbackCookie          = nullptr;
    uint64_t progressStreamId     = 0;
    FileSystemOptions optionsState{};
    FileSystemOptions* options           = nullptr;
    unsigned long totalItems             = 0;
    unsigned long completedItems         = 0;
    uint64_t totalBytes                  = 0;
    uint64_t completedBytes              = 0;
    bool continueOnError                 = false;
    bool allowOverwrite                  = false;
    bool allowReplaceReadonly            = false;
    bool recursive                       = false;
    bool useRecycleBin                   = false;
    unsigned int deleteConcurrencyBudget = 1;
    unsigned int recycleBinBatchSize     = 1;
    FileSystemArenaOwner itemArena;
    FileSystemArenaOwner progressArena;
    BandwidthThrottleWorkerMode bandwidthThrottleWorkerMode = BandwidthThrottleWorkerMode::PerWorkerSubBudget;
    const wchar_t* itemSource                               = nullptr;
    const wchar_t* itemDestination                          = nullptr;
    const wchar_t* progressSource                           = nullptr;
    const wchar_t* progressDestination                      = nullptr;

    ParallelOperationState* parallel = nullptr;

    ULONGLONG lastProgressReportTick = 0;

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    std::wstring reparseRootSourcePath;
    std::wstring reparseRootDestinationPath;
};
#pragma warning(pop)

struct CopyProgressContext
{
    CopyProgressContext() noexcept                             = default;
    CopyProgressContext(const CopyProgressContext&)            = delete;
    CopyProgressContext& operator=(const CopyProgressContext&) = delete;
    CopyProgressContext(CopyProgressContext&&)                 = delete;
    CopyProgressContext& operator=(CopyProgressContext&&)      = delete;

    OperationContext* context         = nullptr;
    uint64_t itemBaseBytes            = 0; // Used only for sequential operations.
    uint64_t lastItemBytesTransferred = 0; // Tracks callback deltas for both sequential and parallel operations.
    uint64_t lastItemTotalBytes       = 0; // Tracks the latest item total reported by the OS progress callback.
    uint64_t maxThrottleDeltaBytes    = 0; // Largest callback chunk observed for sequential rolling-window pacing.
    ULONGLONG startTick               = 0; // Legacy per-item timing; kept for diagnostics.
    uint64_t throttleCallbackCount    = 0; // Number of sequential throttle callbacks observed for the current item.
    uint64_t throttleMaxWindowBytes   = 0; // Largest rolling 1-second window observed at CopyFileEx callback timestamps.
    ULONGLONG firstThrottleTick       = 0; // First callback tick observed for sequential throttle diagnostics.
    ULONGLONG lastThrottleTick        = 0; // Last callback tick observed for sequential throttle diagnostics.
    uint64_t throttleMaxGapMs         = 0; // Largest gap between consecutive CopyFileEx callbacks.
    std::deque<std::pair<ULONGLONG, uint64_t>> throttleWindowSamples;
    std::mutex progressMutex; // Serializes sequential callback field access against concurrent reads.

    ParallelOperationState::BandwidthThrottleState bandwidthThrottle;
};

bool HasFlag(FileSystemFlags flags, FileSystemFlags flag) noexcept
{
    return (static_cast<unsigned long>(flags) & static_cast<unsigned long>(flag)) != 0u;
}

bool IsCancellationHr(HRESULT hr) noexcept
{
    return hr == E_ABORT || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED);
}

HRESULT NormalizeCancellation(HRESULT hr) noexcept
{
    if (IsCancellationHr(hr))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    return hr;
}

HRESULT RemovePathForOverwrite(OperationContext& context, const std::wstring& pathExtended) noexcept;

[[nodiscard]] bool IsReparsePoint(DWORD attributes) noexcept
{
    return (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
}

[[nodiscard]] bool IsDirectory(DWORD attributes) noexcept
{
    return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

struct ReparsePointHeader
{
    DWORD tag        = 0;
    USHORT dataBytes = 0;
    USHORT reserved  = 0;
};
static_assert(sizeof(ReparsePointHeader) == 8);

struct ReparsePointData
{
    DWORD tag       = 0;
    DWORD sizeBytes = 0;
    alignas(8) std::array<std::byte, MAXIMUM_REPARSE_DATA_BUFFER_SIZE> buffer{};
};

struct MountPointReparseHeader
{
    USHORT substituteOffset = 0;
    USHORT substituteLength = 0;
    USHORT printOffset      = 0;
    USHORT printLength      = 0;
};
static_assert(sizeof(MountPointReparseHeader) == 8);

struct SymbolicLinkReparseHeader
{
    USHORT substituteOffset = 0;
    USHORT substituteLength = 0;
    USHORT printOffset      = 0;
    USHORT printLength      = 0;
    ULONG flags             = 0;
};
static_assert(sizeof(SymbolicLinkReparseHeader) == 12);

constexpr ULONG kSymlinkRelativeFlag = 0x00000001u;

struct ParsedDirectoryReparsePoint
{
    DWORD tag       = 0;
    bool isRelative = false;
    std::wstring substitutePath;
    std::wstring printPath;
};

[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

void NormalizeSlashes(std::wstring& path) noexcept
{
    std::ranges::replace(path, L'/', L'\\');
}

[[nodiscard]] size_t GetRootLength(std::wstring_view path) noexcept
{
    if (path.size() >= 2 && path[1] == L':')
    {
        if (path.size() >= 3 && IsPathSeparator(path[2]))
        {
            return 3;
        }
        return 2;
    }

    if (path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\')
    {
        size_t firstSep = path.find(L'\\', 2);
        if (firstSep == std::wstring_view::npos)
        {
            firstSep = path.find(L'/', 2);
        }
        if (firstSep == std::wstring_view::npos)
        {
            return path.size();
        }

        size_t secondSep = path.find_first_of(L"\\/", firstSep + 1);
        if (secondSep == std::wstring_view::npos)
        {
            return path.size();
        }
        return secondSep + 1;
    }

    if (! path.empty() && IsPathSeparator(path.front()))
    {
        return 1;
    }

    return 0;
}

[[nodiscard]] std::wstring TrimTrailingSeparatorsPreserveRoot(std::wstring path) noexcept
{
    NormalizeSlashes(path);
    const size_t rootLength = GetRootLength(path);
    while (path.size() > rootLength && ! path.empty() && IsPathSeparator(path.back()))
    {
        path.pop_back();
    }
    return path;
}

[[nodiscard]] std::wstring_view TrimTrailingSeparatorsPreserveRootView(std::wstring_view path) noexcept
{
    const size_t rootLength = GetRootLength(path);
    while (path.size() > rootLength && ! path.empty() && IsPathSeparator(path.back()))
    {
        path.remove_suffix(1);
    }
    return path;
}

struct FileIdentity final
{
    DWORD volumeSerialNumber = 0;
    uint64_t fileIndex       = 0;
};

[[nodiscard]] HRESULT TryGetFileIdentity(const std::wstring& path, FileIdentity& identity) noexcept
{
    identity = {};

    const DWORD attributes = ::GetFileAttributesW(path.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    wil::unique_handle handle(::CreateFileW(path.c_str(),
                                            FILE_READ_ATTRIBUTES,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_BACKUP_SEMANTICS,
                                            nullptr));
    if (! handle)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    BY_HANDLE_FILE_INFORMATION info{};
    if (! ::GetFileInformationByHandle(handle.get(), &info))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    identity.volumeSerialNumber = info.dwVolumeSerialNumber;
    identity.fileIndex          = (static_cast<uint64_t>(info.nFileIndexHigh) << 32) | static_cast<uint64_t>(info.nFileIndexLow);
    return S_OK;
}

[[nodiscard]] HRESULT TryAreSameFile(const std::wstring& left, const std::wstring& right, bool& same) noexcept
{
    same = false;

    FileIdentity leftId{};
    HRESULT hr = TryGetFileIdentity(left, leftId);
    if (FAILED(hr))
    {
        return hr;
    }

    FileIdentity rightId{};
    hr = TryGetFileIdentity(right, rightId);
    if (FAILED(hr))
    {
        return hr;
    }

    same = leftId.volumeSerialNumber == rightId.volumeSerialNumber && leftId.fileIndex == rightId.fileIndex;
    return S_OK;
}

[[nodiscard]] bool IsPathWithinRoot(std::wstring_view path, std::wstring_view root) noexcept
{
    if (root.empty() || path.size() < root.size())
    {
        return false;
    }

    if (! OrdinalString::StartsWithNoCase(path, root))
    {
        return false;
    }

    if (path.size() == root.size())
    {
        return true;
    }

    return IsPathSeparator(path[root.size()]);
}

[[nodiscard]] std::wstring StripWin32ExtendedPrefix(std::wstring_view path)
{
    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        return std::wstring(L"\\\\") + std::wstring(path.substr(8));
    }
    if (path.rfind(L"\\\\?\\", 0) == 0)
    {
        return std::wstring(path.substr(4));
    }
    return std::wstring(path);
}

[[nodiscard]] std::wstring NtPathToWin32Path(std::wstring_view path)
{
    if (path.rfind(L"\\??\\UNC\\", 0) == 0)
    {
        return std::wstring(L"\\\\") + std::wstring(path.substr(8));
    }
    if (path.rfind(L"\\??\\", 0) == 0)
    {
        return std::wstring(path.substr(4));
    }
    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        return std::wstring(L"\\\\") + std::wstring(path.substr(8));
    }
    if (path.rfind(L"\\\\?\\", 0) == 0)
    {
        return std::wstring(path.substr(4));
    }
    return std::wstring(path);
}

[[nodiscard]] std::wstring Win32PathToNtPath(std::wstring_view path)
{
    if (path.rfind(L"\\??\\", 0) == 0)
    {
        return std::wstring(path);
    }

    if (path.rfind(L"\\\\", 0) == 0)
    {
        return std::wstring(L"\\??\\UNC\\") + std::wstring(path.substr(2));
    }

    return std::wstring(L"\\??\\") + std::wstring(path);
}

[[nodiscard]] bool ParseDirectoryReparsePoint(const ReparsePointData& data, ParsedDirectoryReparsePoint& out) noexcept
{
    out = {};

    if (data.sizeBytes < sizeof(ReparsePointHeader))
    {
        return false;
    }

    const auto* header = reinterpret_cast<const ReparsePointHeader*>(data.buffer.data());
    if (static_cast<size_t>(header->dataBytes) + sizeof(ReparsePointHeader) > static_cast<size_t>(data.sizeBytes))
    {
        return false;
    }

    out.tag = header->tag;

    const std::byte* payloadBase = data.buffer.data() + sizeof(ReparsePointHeader);
    const size_t payloadBytes    = header->dataBytes;

    auto readPathSlice = [&](USHORT offsetBytes, USHORT lengthBytes, size_t fixedHeaderBytes, std::wstring& target) noexcept -> bool
    {
        if ((offsetBytes % sizeof(wchar_t)) != 0u || (lengthBytes % sizeof(wchar_t)) != 0u)
        {
            return false;
        }
        if (payloadBytes < fixedHeaderBytes)
        {
            return false;
        }
        const size_t pathBufferBytes = payloadBytes - fixedHeaderBytes;
        if (offsetBytes > pathBufferBytes || lengthBytes > pathBufferBytes ||
            (static_cast<size_t>(offsetBytes) + static_cast<size_t>(lengthBytes)) > pathBufferBytes)
        {
            return false;
        }

        const auto* text = reinterpret_cast<const wchar_t*>(payloadBase + fixedHeaderBytes + offsetBytes);
        target.assign(text, text + (lengthBytes / sizeof(wchar_t)));
        return true;
    };

    if (out.tag == IO_REPARSE_TAG_MOUNT_POINT)
    {
        if (payloadBytes < sizeof(MountPointReparseHeader))
        {
            return false;
        }

        const auto* mount = reinterpret_cast<const MountPointReparseHeader*>(payloadBase);
        if (! readPathSlice(mount->substituteOffset, mount->substituteLength, sizeof(MountPointReparseHeader), out.substitutePath))
        {
            return false;
        }
        if (! readPathSlice(mount->printOffset, mount->printLength, sizeof(MountPointReparseHeader), out.printPath))
        {
            return false;
        }
        out.isRelative = false;
        return true;
    }

    if (out.tag == IO_REPARSE_TAG_SYMLINK)
    {
        if (payloadBytes < sizeof(SymbolicLinkReparseHeader))
        {
            return false;
        }

        const auto* symlink = reinterpret_cast<const SymbolicLinkReparseHeader*>(payloadBase);
        if (! readPathSlice(symlink->substituteOffset, symlink->substituteLength, sizeof(SymbolicLinkReparseHeader), out.substitutePath))
        {
            return false;
        }
        if (! readPathSlice(symlink->printOffset, symlink->printLength, sizeof(SymbolicLinkReparseHeader), out.printPath))
        {
            return false;
        }
        out.isRelative = (symlink->flags & kSymlinkRelativeFlag) != 0u;
        return true;
    }

    return false;
}

[[nodiscard]] std::wstring ResolveReparseTargetAbsolute(const PathInfo& source, const ParsedDirectoryReparsePoint& parsed) noexcept
{
    std::wstring rawTarget = parsed.substitutePath.empty() ? parsed.printPath : parsed.substitutePath;
    if (rawTarget.empty())
    {
        return {};
    }

    rawTarget = NtPathToWin32Path(rawTarget);
    NormalizeSlashes(rawTarget);

    if (parsed.isRelative)
    {
        std::filesystem::path parent   = std::filesystem::path(source.display).parent_path();
        std::filesystem::path combined = parent / std::filesystem::path(rawTarget);
        std::wstring absolute          = MakeAbsolutePath(combined.lexically_normal().wstring());
        absolute                       = StripWin32ExtendedPrefix(absolute);
        return TrimTrailingSeparatorsPreserveRoot(absolute);
    }

    std::wstring absolute = MakeAbsolutePath(rawTarget);
    absolute              = StripWin32ExtendedPrefix(absolute);
    return TrimTrailingSeparatorsPreserveRoot(absolute);
}

[[nodiscard]] bool TryRetargetPathIntoDestination(std::wstring_view absoluteTargetPath,
                                                  std::wstring_view sourceRootPath,
                                                  std::wstring_view destinationRootPath,
                                                  std::wstring& mappedOut) noexcept
{
    std::wstring normalizedTargetStorage;
    std::wstring normalizedSourceStorage;
    std::wstring normalizedDestStorage;

    const auto normalizeIfNeeded = [](std::wstring_view path, std::wstring& storage) noexcept -> std::wstring_view
    {
        if (path.find(L'/') == std::wstring_view::npos)
        {
            return path;
        }

        storage.assign(path);
        NormalizeSlashes(storage);
        return storage;
    };

    std::wstring_view normalizedTarget = normalizeIfNeeded(absoluteTargetPath, normalizedTargetStorage);
    std::wstring_view normalizedSource = normalizeIfNeeded(sourceRootPath, normalizedSourceStorage);
    std::wstring_view normalizedDest   = normalizeIfNeeded(destinationRootPath, normalizedDestStorage);

    normalizedTarget = TrimTrailingSeparatorsPreserveRootView(normalizedTarget);
    normalizedSource = TrimTrailingSeparatorsPreserveRootView(normalizedSource);
    normalizedDest   = TrimTrailingSeparatorsPreserveRootView(normalizedDest);

    if (normalizedTarget.empty() || normalizedSource.empty() || normalizedDest.empty())
    {
        return false;
    }

    if (! IsPathWithinRoot(normalizedTarget, normalizedSource))
    {
        return false;
    }

    std::wstring_view suffix;
    if (normalizedTarget.size() > normalizedSource.size())
    {
        suffix = normalizedTarget.substr(normalizedSource.size());
        while (! suffix.empty() && IsPathSeparator(suffix.front()))
        {
            suffix.remove_prefix(1);
        }
    }

    mappedOut.clear();
    mappedOut.reserve(normalizedDest.size() + (suffix.empty() ? 0 : 1 + suffix.size()));
    mappedOut.append(normalizedDest);
    if (! suffix.empty())
    {
        if (! mappedOut.empty() && ! IsPathSeparator(mappedOut.back()))
        {
            mappedOut.push_back(L'\\');
        }
        mappedOut.append(suffix);
    }

    mappedOut = TrimTrailingSeparatorsPreserveRoot(std::move(mappedOut));
    return true;
}

[[nodiscard]] bool EndsWithSeparator(std::wstring_view path) noexcept
{
    return ! path.empty() && IsPathSeparator(path.back());
}

HRESULT BuildMountPointReparseData(std::wstring targetPath, ReparsePointData& out) noexcept
{
    NormalizeSlashes(targetPath);
    if (! EndsWithSeparator(targetPath))
    {
        targetPath.push_back(L'\\');
    }

    std::wstring substitute = Win32PathToNtPath(targetPath);

    const size_t substituteBytes = substitute.size() * sizeof(wchar_t);
    const size_t printBytes      = targetPath.size() * sizeof(wchar_t);
    const size_t pathBufferBytes = substituteBytes + sizeof(wchar_t) + printBytes + sizeof(wchar_t);
    const size_t payloadBytes    = sizeof(MountPointReparseHeader) + pathBufferBytes;
    const size_t totalBytes      = sizeof(ReparsePointHeader) + payloadBytes;

    if (payloadBytes > static_cast<size_t>(std::numeric_limits<USHORT>::max()) || totalBytes > out.buffer.size())
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    out           = {};
    out.tag       = IO_REPARSE_TAG_MOUNT_POINT;
    out.sizeBytes = static_cast<DWORD>(totalBytes);

    auto* header      = reinterpret_cast<ReparsePointHeader*>(out.buffer.data());
    header->tag       = IO_REPARSE_TAG_MOUNT_POINT;
    header->dataBytes = static_cast<USHORT>(payloadBytes);
    header->reserved  = 0;

    auto* mountHeader             = reinterpret_cast<MountPointReparseHeader*>(out.buffer.data() + sizeof(ReparsePointHeader));
    mountHeader->substituteOffset = 0;
    mountHeader->substituteLength = static_cast<USHORT>(substituteBytes);
    mountHeader->printOffset      = static_cast<USHORT>(substituteBytes + sizeof(wchar_t));
    mountHeader->printLength      = static_cast<USHORT>(printBytes);

    std::byte* pathBuffer = out.buffer.data() + sizeof(ReparsePointHeader) + sizeof(MountPointReparseHeader);
    std::memcpy(pathBuffer, substitute.data(), substituteBytes);
    std::memset(pathBuffer + substituteBytes, 0, sizeof(wchar_t));
    std::memcpy(pathBuffer + substituteBytes + sizeof(wchar_t), targetPath.data(), printBytes);
    std::memset(pathBuffer + substituteBytes + sizeof(wchar_t) + printBytes, 0, sizeof(wchar_t));
    return S_OK;
}

HRESULT BuildSymlinkReparseData(std::wstring targetPath, bool relative, ReparsePointData& out) noexcept
{
    NormalizeSlashes(targetPath);
    std::wstring substitute = targetPath;
    std::wstring print      = targetPath;

    if (! relative)
    {
        substitute = Win32PathToNtPath(substitute);
    }

    const size_t substituteBytes = substitute.size() * sizeof(wchar_t);
    const size_t printBytes      = print.size() * sizeof(wchar_t);
    const size_t pathBufferBytes = substituteBytes + sizeof(wchar_t) + printBytes + sizeof(wchar_t);
    const size_t payloadBytes    = sizeof(SymbolicLinkReparseHeader) + pathBufferBytes;
    const size_t totalBytes      = sizeof(ReparsePointHeader) + payloadBytes;

    if (payloadBytes > static_cast<size_t>(std::numeric_limits<USHORT>::max()) || totalBytes > out.buffer.size())
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    out           = {};
    out.tag       = IO_REPARSE_TAG_SYMLINK;
    out.sizeBytes = static_cast<DWORD>(totalBytes);

    auto* header      = reinterpret_cast<ReparsePointHeader*>(out.buffer.data());
    header->tag       = IO_REPARSE_TAG_SYMLINK;
    header->dataBytes = static_cast<USHORT>(payloadBytes);
    header->reserved  = 0;

    auto* symlinkHeader             = reinterpret_cast<SymbolicLinkReparseHeader*>(out.buffer.data() + sizeof(ReparsePointHeader));
    symlinkHeader->substituteOffset = 0;
    symlinkHeader->substituteLength = static_cast<USHORT>(substituteBytes);
    symlinkHeader->printOffset      = static_cast<USHORT>(substituteBytes + sizeof(wchar_t));
    symlinkHeader->printLength      = static_cast<USHORT>(printBytes);
    symlinkHeader->flags            = relative ? kSymlinkRelativeFlag : 0u;

    std::byte* pathBuffer = out.buffer.data() + sizeof(ReparsePointHeader) + sizeof(SymbolicLinkReparseHeader);
    std::memcpy(pathBuffer, substitute.data(), substituteBytes);
    std::memset(pathBuffer + substituteBytes, 0, sizeof(wchar_t));
    std::memcpy(pathBuffer + substituteBytes + sizeof(wchar_t), print.data(), printBytes);
    std::memset(pathBuffer + substituteBytes + sizeof(wchar_t) + printBytes, 0, sizeof(wchar_t));
    return S_OK;
}

HRESULT ReadReparsePointData(const std::wstring& path, ReparsePointData& out) noexcept
{
    out = {};

    // Protected junctions (e.g. localized/system junctions) may deny "read data / list directory" access
    // but still allow querying reparse metadata. Keep access minimal so we can copy the link itself.
    wil::unique_handle handle(CreateFileW(path.c_str(),
                                          FILE_READ_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                          nullptr));
    if (! handle)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    DWORD bytesReturned = 0;
    if (! DeviceIoControl(handle.get(), FSCTL_GET_REPARSE_POINT, nullptr, 0, out.buffer.data(), static_cast<DWORD>(out.buffer.size()), &bytesReturned, nullptr))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if (bytesReturned < sizeof(ReparsePointHeader))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const auto* header = reinterpret_cast<const ReparsePointHeader*>(out.buffer.data());
    out.tag            = header->tag;
    out.sizeBytes      = bytesReturned;
    return S_OK;
}

HRESULT WriteReparsePointData(const std::wstring& path, const ReparsePointData& data) noexcept
{
    if (data.sizeBytes < sizeof(ReparsePointHeader) || data.sizeBytes > data.buffer.size())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    // Minimal access for setting reparse data on the destination link.
    wil::unique_handle handle(CreateFileW(path.c_str(),
                                          FILE_WRITE_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                          nullptr));
    if (! handle)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    DWORD bytesReturned = 0;
    if (! DeviceIoControl(
            handle.get(), FSCTL_SET_REPARSE_POINT, const_cast<std::byte*>(data.buffer.data()), data.sizeBytes, nullptr, 0, &bytesReturned, nullptr))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

void AddCompletedBytes(OperationContext& context, uint64_t bytes) noexcept
{
    if (bytes == 0)
    {
        return;
    }

    if (context.parallel)
    {
        context.parallel->completedBytes.fetch_add(bytes, std::memory_order_acq_rel);
        return;
    }

    if (std::numeric_limits<uint64_t>::max() - context.completedBytes < bytes)
    {
        context.completedBytes = std::numeric_limits<uint64_t>::max();
        return;
    }

    context.completedBytes += bytes;
}

void AddCompletedItems(OperationContext& context, unsigned long items) noexcept
{
    if (items == 0)
    {
        return;
    }

    if (context.parallel)
    {
        context.parallel->completedItems.fetch_add(items, std::memory_order_acq_rel);
        return;
    }

    constexpr uint64_t maxUlong = static_cast<uint64_t>(std::numeric_limits<unsigned long>::max());
    const uint64_t current      = static_cast<uint64_t>(context.completedItems);
    const uint64_t desired      = current + static_cast<uint64_t>(items);
    context.completedItems      = static_cast<unsigned long>(std::min(desired, maxUlong));
}

uint64_t GetBandwidthLimit(const FileSystemOptions* options) noexcept
{
    if (! options)
    {
        return 0;
    }
    return options->bandwidthLimitBytesPerSecond;
}

constexpr uint64_t kBandwidthThrottleBurstWindowMs        = 250ull;
constexpr uint64_t kBandwidthThrottleMinCapacityBytes     = 64ull * 1024ull;
constexpr uint64_t kBandwidthThrottleMaxCapacityBytes     = 4ull * 1024ull * 1024ull;
constexpr uint64_t kBandwidthThrottleWorkerActiveWindowMs = 500ull;
constexpr uint64_t kBandwidthThrottleBoundaryGuardMs      = 10ull;
constexpr DWORD kBandwidthThrottleCancelPollMs            = 10u;
#ifdef _DEBUG
constexpr std::wstring_view kBandwidthThrottleWorkerModeEnvVar = L"REDSALAMANDER_FILEOPS_BW_WORKER_MODE";
constexpr std::wstring_view kForceMoveCopyFallbackEnvVar       = L"REDSALAMANDER_FILEOPS_FORCE_MOVE_COPY_FALLBACK";
#endif

constexpr uint64_t SaturatingBytesForElapsedMs(uint64_t bytesPerSecond, uint64_t elapsedMs) noexcept
{
    if (bytesPerSecond == 0 || elapsedMs == 0)
    {
        return 0;
    }

    // Split: (bps * elapsed) / 1000 = bps * (elapsed/1000) + (bps * (elapsed%1000)) / 1000
    // This avoids overflow without platform-specific intrinsics.
    const uint64_t wholeSeconds = elapsedMs / 1000;
    const uint64_t remainderMs  = elapsedMs % 1000;

    if (wholeSeconds > 0 && bytesPerSecond > std::numeric_limits<uint64_t>::max() / wholeSeconds)
    {
        return std::numeric_limits<uint64_t>::max();
    }

    const uint64_t wholePart = bytesPerSecond * wholeSeconds;
    const uint64_t fracPart  = ((bytesPerSecond / 1000ull) * remainderMs) + (((bytesPerSecond % 1000ull) * remainderMs) / 1000ull);

    if (wholePart > std::numeric_limits<uint64_t>::max() - fracPart)
    {
        return std::numeric_limits<uint64_t>::max();
    }

    return wholePart + fracPart;
}

static_assert(SaturatingBytesForElapsedMs(1500ull, 500ull) == 750ull);
static_assert(SaturatingBytesForElapsedMs((std::numeric_limits<uint64_t>::max)(), 999ull) ==
              (((std::numeric_limits<uint64_t>::max)() / 1000ull) * 999ull) + ((((std::numeric_limits<uint64_t>::max)() % 1000ull) * 999ull) / 1000ull));
static_assert(SaturatingBytesForElapsedMs((std::numeric_limits<uint64_t>::max)(), (std::numeric_limits<uint64_t>::max)()) ==
              (std::numeric_limits<uint64_t>::max)());

uint64_t CalculateBandwidthThrottleCapacity(uint64_t bytesPerSecond) noexcept
{
    if (bytesPerSecond == 0)
    {
        return 0;
    }

    const uint64_t burstBytes = SaturatingBytesForElapsedMs(bytesPerSecond, kBandwidthThrottleBurstWindowMs);
    return std::clamp(burstBytes, kBandwidthThrottleMinCapacityBytes, kBandwidthThrottleMaxCapacityBytes);
}

uint64_t CalculateThrottleSleepMs(uint64_t debtBytes, uint64_t bytesPerSecond) noexcept
{
    if (debtBytes == 0 || bytesPerSecond == 0)
    {
        return 0;
    }

    constexpr uint64_t kScale = 1000ull;
    if (debtBytes > (std::numeric_limits<uint64_t>::max() - (bytesPerSecond - 1ull)) / kScale)
    {
        return std::numeric_limits<uint64_t>::max();
    }

    uint64_t sleepMs = ((debtBytes * kScale) + bytesPerSecond - 1ull) / bytesPerSecond;

    // Bias away from exact callback boundaries so a rolling 1-second window does not admit an
    // extra whole callback chunk when the copy API reports progress exactly on the rate limit.
    if (sleepMs <= (std::numeric_limits<uint64_t>::max)() - kBandwidthThrottleBoundaryGuardMs)
    {
        sleepMs += kBandwidthThrottleBoundaryGuardMs;
    }

    return sleepMs;
}

void CaptureSequentialThrottleSample(CopyProgressContext& progressContext, ULONGLONG nowTick, uint64_t itemCompleted) noexcept
{
    ++progressContext.throttleCallbackCount;
    if (progressContext.firstThrottleTick == 0 || nowTick < progressContext.firstThrottleTick)
    {
        progressContext.firstThrottleTick = nowTick;
    }

    if (progressContext.lastThrottleTick != 0 && nowTick >= progressContext.lastThrottleTick)
    {
        progressContext.throttleMaxGapMs = (std::max)(progressContext.throttleMaxGapMs, static_cast<uint64_t>(nowTick - progressContext.lastThrottleTick));
    }
    progressContext.lastThrottleTick = nowTick;

    if (! progressContext.throttleWindowSamples.empty())
    {
        const auto& lastSample = progressContext.throttleWindowSamples.back();
        if (lastSample.first == nowTick && lastSample.second == itemCompleted)
        {
            return;
        }
    }

    progressContext.throttleWindowSamples.emplace_back(nowTick, itemCompleted);
    while (! progressContext.throttleWindowSamples.empty() && progressContext.throttleWindowSamples.front().first + 1000ull < nowTick)
    {
        progressContext.throttleWindowSamples.pop_front();
    }

    if (! progressContext.throttleWindowSamples.empty() && itemCompleted >= progressContext.throttleWindowSamples.front().second)
    {
        progressContext.throttleMaxWindowBytes =
            (std::max)(progressContext.throttleMaxWindowBytes, itemCompleted - progressContext.throttleWindowSamples.front().second);
    }
}

void EmitSequentialThrottleSummary(const OperationContext& context, const CopyProgressContext& progressContext, uint64_t totalBytes, HRESULT hr) noexcept
{
    if (context.parallel || context.optionsState.bandwidthLimitBytesPerSecond == 0 || progressContext.throttleCallbackCount == 0)
    {
        return;
    }

    const uint64_t firstCallbackDelayUs = progressContext.firstThrottleTick >= progressContext.startTick
                                              ? static_cast<uint64_t>(progressContext.firstThrottleTick - progressContext.startTick) * 1000ull
                                              : 0ull;
    const uint64_t maxGapUs             = progressContext.throttleMaxGapMs * 1000ull;
    const std::wstring detail           = std::format(L"source={} destination={} limit={} totalBytes={} callbacks={} firstDelayUs={} maxGapUs={}",
                                                      context.progressSource ? context.progressSource : L"",
                                                      context.progressDestination ? context.progressDestination : L"",
                                                      context.optionsState.bandwidthLimitBytesPerSecond,
                                                      totalBytes,
                                                      progressContext.throttleCallbackCount,
                                                      firstCallbackDelayUs,
                                                      maxGapUs);

    Debug::Perf::Emit(L"FileOps.BandwidthThrottle.SequentialCallbackCount", detail, 0, progressContext.throttleCallbackCount, totalBytes, hr);
    Debug::Perf::Emit(L"FileOps.BandwidthThrottle.SequentialFirstCallbackDelayUs",
                      detail,
                      firstCallbackDelayUs,
                      progressContext.maxThrottleDeltaBytes,
                      context.optionsState.bandwidthLimitBytesPerSecond,
                      hr);
    Debug::Perf::Emit(L"FileOps.BandwidthThrottle.SequentialMaxGapUs",
                      detail,
                      maxGapUs,
                      progressContext.throttleCallbackCount,
                      context.optionsState.bandwidthLimitBytesPerSecond,
                      hr);
    Debug::Perf::Emit(L"FileOps.BandwidthThrottle.SequentialMaxDeltaBytes",
                      detail,
                      progressContext.maxThrottleDeltaBytes,
                      totalBytes,
                      context.optionsState.bandwidthLimitBytesPerSecond,
                      hr);
    Debug::Perf::Emit(L"FileOps.BandwidthThrottle.SequentialMaxWindowBytes",
                      detail,
                      progressContext.throttleMaxWindowBytes,
                      1000,
                      context.optionsState.bandwidthLimitBytesPerSecond,
                      hr);
}

template <typename TState> void ResetBandwidthThrottleState(TState& state, uint64_t bytesPerSecond, ULONGLONG nowTick) noexcept
{
    state.configuredLimitBytesPerSecond = bytesPerSecond;
    state.availableBytes                = 0;
    state.lastRefillTick                = nowTick;
}

template <typename TState> void RefillBandwidthThrottleState(TState& state, ULONGLONG nowTick) noexcept
{
    if (state.lastRefillTick == 0 || state.configuredLimitBytesPerSecond == 0)
    {
        state.lastRefillTick = nowTick;
        return;
    }

    if (nowTick < state.lastRefillTick)
    {
        state.lastRefillTick = nowTick;
        return;
    }

    const uint64_t elapsedMs = static_cast<uint64_t>(nowTick - state.lastRefillTick);
    if (elapsedMs == 0)
    {
        return;
    }

    const uint64_t refillBytes = SaturatingBytesForElapsedMs(state.configuredLimitBytesPerSecond, elapsedMs);
    const uint64_t capacity    = CalculateBandwidthThrottleCapacity(state.configuredLimitBytesPerSecond);
    const int64_t available    = state.availableBytes;
    const int64_t refill =
        refillBytes > static_cast<uint64_t>(std::numeric_limits<int64_t>::max()) ? std::numeric_limits<int64_t>::max() : static_cast<int64_t>(refillBytes);
    const int64_t maxAvailable = static_cast<int64_t>(capacity);
    if (available > std::numeric_limits<int64_t>::max() - refill)
    {
        state.availableBytes = maxAvailable;
    }
    else
    {
        state.availableBytes = (std::min)(available + refill, maxAvailable);
    }

    state.lastRefillTick = nowTick;
}

template <typename TState> void ReconfigureBandwidthThrottleState(TState& state, uint64_t bytesPerSecond, ULONGLONG nowTick) noexcept
{
    state.configuredLimitBytesPerSecond = bytesPerSecond;
    const int64_t maxAvailable          = static_cast<int64_t>(CalculateBandwidthThrottleCapacity(bytesPerSecond));
    if (state.availableBytes > maxAvailable)
    {
        state.availableBytes = maxAvailable;
    }
    state.lastRefillTick = nowTick;
}

template <typename TState> uint64_t ChargeBandwidthThrottleState(TState& state, uint64_t deltaBytes) noexcept
{
    if (deltaBytes > static_cast<uint64_t>(std::numeric_limits<int64_t>::max()))
    {
        state.availableBytes = (std::numeric_limits<int64_t>::min)() / 2;
    }
    else
    {
        const int64_t delta = static_cast<int64_t>(deltaBytes);
        if (state.availableBytes < (std::numeric_limits<int64_t>::min)() + delta)
        {
            state.availableBytes = (std::numeric_limits<int64_t>::min)() / 2;
        }
        else
        {
            state.availableBytes -= delta;
        }
    }

    if (state.availableBytes >= 0)
    {
        return 0;
    }

    if (state.availableBytes == (std::numeric_limits<int64_t>::min)())
    {
        return static_cast<uint64_t>((std::numeric_limits<int64_t>::max)());
    }

    return static_cast<uint64_t>(-state.availableBytes);
}

#ifdef _DEBUG
BandwidthThrottleWorkerMode GetBandwidthThrottleWorkerModeOverride() noexcept
{
    const DWORD required = ::GetEnvironmentVariableW(kBandwidthThrottleWorkerModeEnvVar.data(), nullptr, 0u);
    if (required == 0)
    {
        return BandwidthThrottleWorkerMode::PerWorkerSubBudget;
    }

    std::wstring value(required, L'\0');
    const DWORD written = ::GetEnvironmentVariableW(kBandwidthThrottleWorkerModeEnvVar.data(), value.data(), required);
    if (written == 0 || written >= required)
    {
        return BandwidthThrottleWorkerMode::PerWorkerSubBudget;
    }
    value.resize(written);
    if (_wcsicmp(value.c_str(), L"shared") == 0)
    {
        return BandwidthThrottleWorkerMode::SharedOnly;
    }
    if (_wcsicmp(value.c_str(), L"perworker") == 0 || _wcsicmp(value.c_str(), L"worker") == 0)
    {
        return BandwidthThrottleWorkerMode::PerWorkerSubBudget;
    }

    return BandwidthThrottleWorkerMode::PerWorkerSubBudget;
}
#else
BandwidthThrottleWorkerMode GetBandwidthThrottleWorkerModeOverride() noexcept
{
    return BandwidthThrottleWorkerMode::PerWorkerSubBudget;
}
#endif

#ifdef _DEBUG
bool ShouldForceMoveCopyFallbackForSelfTest() noexcept
{
    wchar_t value[8]{};
    const DWORD written = ::GetEnvironmentVariableW(kForceMoveCopyFallbackEnvVar.data(), value, static_cast<DWORD>(std::size(value)));
    if (written == 0 || written >= static_cast<DWORD>(std::size(value)))
    {
        return false;
    }

    return _wcsicmp(value, L"1") == 0 || _wcsicmp(value, L"true") == 0 || _wcsicmp(value, L"yes") == 0;
}
#endif

HRESULT CalculateStringBytes(const wchar_t* text, unsigned long* outBytes) noexcept
{
    if (! outBytes)
    {
        return E_POINTER;
    }

    if (! text)
    {
        *outBytes = 0;
        return S_OK;
    }

    const size_t length = ::wcslen(text);
    if (length > (std::numeric_limits<unsigned long>::max() / sizeof(wchar_t)) - 1u)
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    *outBytes = static_cast<unsigned long>((length + 1u) * sizeof(wchar_t));
    return S_OK;
}

HRESULT BuildArenaForPaths(
    FileSystemArenaOwner& arenaOwner, const wchar_t* source, const wchar_t* destination, const wchar_t** outSource, const wchar_t** outDestination) noexcept
{
    if (! outSource || ! outDestination)
    {
        return E_POINTER;
    }

    *outSource      = nullptr;
    *outDestination = nullptr;

    unsigned long sourceBytes = 0;
    HRESULT hr                = CalculateStringBytes(source, &sourceBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    unsigned long destinationBytes = 0;
    hr                             = CalculateStringBytes(destination, &destinationBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    unsigned long totalBytes = sourceBytes;
    if (destinationBytes > 0)
    {
        if (totalBytes > std::numeric_limits<unsigned long>::max() - destinationBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
        totalBytes += destinationBytes;
    }

    FileSystemArena* arena = arenaOwner.Get();
    if (! arena || arena->buffer == nullptr || arena->capacityBytes < totalBytes)
    {
        hr = arenaOwner.Initialize(totalBytes);
        if (FAILED(hr))
        {
            return hr;
        }
        arena = arenaOwner.Get();
    }

    if (arena && arena->buffer)
    {
        arena->usedBytes = 0;
    }

    if (sourceBytes > 0)
    {
        auto* sourceBuffer = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, sourceBytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! sourceBuffer)
        {
            return E_OUTOFMEMORY;
        }

        const size_t sourceLength = (sourceBytes / sizeof(wchar_t)) - 1u;
        if (sourceLength > 0)
        {
            ::CopyMemory(sourceBuffer, source, sourceLength * sizeof(wchar_t));
        }
        sourceBuffer[sourceLength] = L'\0';
        *outSource                 = sourceBuffer;
    }

    if (destinationBytes > 0)
    {
        auto* destinationBuffer = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, destinationBytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! destinationBuffer)
        {
            return E_OUTOFMEMORY;
        }

        const size_t destinationLength = (destinationBytes / sizeof(wchar_t)) - 1u;
        if (destinationLength > 0)
        {
            ::CopyMemory(destinationBuffer, destination, destinationLength * sizeof(wchar_t));
        }
        destinationBuffer[destinationLength] = L'\0';
        *outDestination                      = destinationBuffer;
    }

    return S_OK;
}

HRESULT SetItemPaths(OperationContext& context, const wchar_t* source, const wchar_t* destination) noexcept
{
    return BuildArenaForPaths(context.itemArena, source, destination, &context.itemSource, &context.itemDestination);
}

HRESULT SetProgressPaths(OperationContext& context, const wchar_t* source, const wchar_t* destination) noexcept
{
    return BuildArenaForPaths(context.progressArena, source, destination, &context.progressSource, &context.progressDestination);
}

HRESULT CheckCancelLocked(OperationContext& context) noexcept
{
    if (context.parallel)
    {
        const bool cancelRequested = context.parallel->cancelRequested.load(std::memory_order_acquire);
        const bool stopOnError     = context.parallel->stopOnErrorRequested.load(std::memory_order_acquire);
        if (cancelRequested || stopOnError)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    if (! context.callback)
    {
        return S_OK;
    }

    BOOL cancel = FALSE;
    HRESULT hr  = context.callback->FileSystemShouldCancel(&cancel, context.callbackCookie);
    hr          = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    if (cancel)
    {
        if (context.parallel)
        {
            context.parallel->cancelRequested.store(true, std::memory_order_release);
        }
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    return S_OK;
}

HRESULT CheckCancel(OperationContext& context) noexcept
{
    if (context.parallel)
    {
        const bool cancelRequested = context.parallel->cancelRequested.load(std::memory_order_acquire);
        const bool stopOnError     = context.parallel->stopOnErrorRequested.load(std::memory_order_acquire);
        if (cancelRequested || stopOnError)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        constexpr ULONGLONG kMinCancelCheckMs = 50ull;
        const ULONGLONG nowTick               = GetTickCount64();
        const ULONGLONG lastTick              = context.parallel->lastCancelCheckTick.load(std::memory_order_acquire);
        if (lastTick != 0 && nowTick >= lastTick && (nowTick - lastTick) < kMinCancelCheckMs)
        {
            return S_OK;
        }

        std::scoped_lock lock(context.parallel->callbackMutex);
        const HRESULT hr = CheckCancelLocked(context);
        context.parallel->lastCancelCheckTick.store(nowTick, std::memory_order_release);
        return hr;
    }

    return CheckCancelLocked(context);
}

HRESULT CheckCancelImmediate(OperationContext& context) noexcept
{
    if (context.parallel)
    {
        std::scoped_lock lock(context.parallel->callbackMutex);
        const HRESULT hr = CheckCancelLocked(context);
        context.parallel->lastCancelCheckTick.store(GetTickCount64(), std::memory_order_release);
        return hr;
    }

    return CheckCancelLocked(context);
}

HRESULT SleepForBandwidthThrottle(OperationContext& context, uint64_t waitMs) noexcept
{
    uint64_t remainingMs = waitMs;
    while (remainingMs > 0)
    {
        const DWORD sliceMs =
            remainingMs > static_cast<uint64_t>(kBandwidthThrottleCancelPollMs) ? kBandwidthThrottleCancelPollMs : static_cast<DWORD>(remainingMs);
        ::Sleep(sliceMs);
        remainingMs -= sliceMs;

        const HRESULT hr = CheckCancelImmediate(context);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    return S_OK;
}

HRESULT ApplyBandwidthThrottle(OperationContext& context, CopyProgressContext& progressContext, uint64_t deltaBytes) noexcept
{
    if (deltaBytes == 0)
    {
        return S_OK;
    }

    const uint64_t bytesPerSecond =
        context.parallel ? context.parallel->bandwidthLimitBytesPerSecond.load(std::memory_order_acquire) : GetBandwidthLimit(context.options);
    if (bytesPerSecond == 0)
    {
        return S_OK;
    }

    const ULONGLONG nowTick = GetTickCount64();
    if (! context.parallel)
    {
        if (progressContext.startTick == 0)
        {
            progressContext.startTick = nowTick;
        }

        progressContext.maxThrottleDeltaBytes = (std::max)(progressContext.maxThrottleDeltaBytes, deltaBytes);
        const uint64_t itemCompleted          = progressContext.lastItemBytesTransferred;
        CaptureSequentialThrottleSample(progressContext, nowTick, itemCompleted);
        const uint64_t reserveBytes     = progressContext.maxThrottleDeltaBytes;
        constexpr uint64_t maxSafeBytes = std::numeric_limits<uint64_t>::max() / 1000ull;

        uint64_t desiredBytes = itemCompleted;
        if (desiredBytes > std::numeric_limits<uint64_t>::max() - reserveBytes)
        {
            desiredBytes = std::numeric_limits<uint64_t>::max();
        }
        else
        {
            desiredBytes += reserveBytes;
        }

        uint64_t desiredMs = 0;
        if (desiredBytes > 0 && desiredBytes <= maxSafeBytes)
        {
            desiredMs = (desiredBytes * 1000ull) / bytesPerSecond;
        }
        else if (desiredBytes > maxSafeBytes)
        {
            desiredMs = std::numeric_limits<uint64_t>::max();
        }

        const uint64_t elapsedMs = nowTick >= progressContext.startTick ? static_cast<uint64_t>(nowTick - progressContext.startTick) : 0ull;
        if (desiredMs <= elapsedMs)
        {
            return S_OK;
        }

        return SleepForBandwidthThrottle(context, desiredMs - elapsedMs);
    }

    auto& throttleState = context.parallel ? context.parallel->bandwidthThrottle : progressContext.bandwidthThrottle;

    uint64_t sharedDebtBytes      = 0;
    uint64_t workerDebtBytes      = 0;
    uint64_t workerBytesPerSecond = 0;
    {
        std::scoped_lock lock(throttleState.mutex);

        if (throttleState.lastRefillTick == 0)
        {
            ResetBandwidthThrottleState(throttleState, bytesPerSecond, nowTick);
        }
        else
        {
            RefillBandwidthThrottleState(throttleState, nowTick);
            if (throttleState.configuredLimitBytesPerSecond != bytesPerSecond)
            {
                ReconfigureBandwidthThrottleState(throttleState, bytesPerSecond, nowTick);
            }
        }

        sharedDebtBytes = ChargeBandwidthThrottleState(throttleState, deltaBytes);

        const bool useWorkerSubBudget = context.parallel && context.bandwidthThrottleWorkerMode == BandwidthThrottleWorkerMode::PerWorkerSubBudget &&
                                        context.progressStreamId < throttleState.workerStates.size();
        if (useWorkerSubBudget)
        {
            ParallelOperationState::BandwidthThrottleState::WorkerState& workerState = throttleState.workerStates[context.progressStreamId];
            workerState.lastActiveTick                                               = nowTick;

            size_t activeWorkerCount = 0;
            for (const ParallelOperationState::BandwidthThrottleState::WorkerState& candidateState : throttleState.workerStates)
            {
                if (candidateState.lastActiveTick == 0 || nowTick < candidateState.lastActiveTick)
                {
                    continue;
                }

                if (static_cast<uint64_t>(nowTick - candidateState.lastActiveTick) > kBandwidthThrottleWorkerActiveWindowMs)
                {
                    continue;
                }

                ++activeWorkerCount;
            }

            activeWorkerCount    = (std::max)(activeWorkerCount, size_t{1});
            workerBytesPerSecond = bytesPerSecond / activeWorkerCount;
            if ((bytesPerSecond % activeWorkerCount) != 0)
            {
                ++workerBytesPerSecond;
            }

            if (workerState.lastRefillTick == 0)
            {
                ResetBandwidthThrottleState(workerState, workerBytesPerSecond, nowTick);
            }
            else
            {
                RefillBandwidthThrottleState(workerState, nowTick);
                if (workerState.configuredLimitBytesPerSecond != workerBytesPerSecond)
                {
                    ReconfigureBandwidthThrottleState(workerState, workerBytesPerSecond, nowTick);
                }
            }

            workerDebtBytes = ChargeBandwidthThrottleState(workerState, deltaBytes);
        }
    }

    const uint64_t sharedSleepMs = CalculateThrottleSleepMs(sharedDebtBytes, bytesPerSecond);
    const uint64_t workerSleepMs = workerBytesPerSecond == 0 ? 0 : CalculateThrottleSleepMs(workerDebtBytes, workerBytesPerSecond);
    const uint64_t sleepMs       = (std::max)(sharedSleepMs, workerSleepMs);
    if (sleepMs == 0)
    {
        return S_OK;
    }

    return SleepForBandwidthThrottle(context, sleepMs);
}

HRESULT ReportProgress(OperationContext& context, uint64_t currentItemTotalBytes, uint64_t currentItemCompletedBytes) noexcept
{
    if (context.parallel)
    {
        const bool cancelRequested = context.parallel->cancelRequested.load(std::memory_order_acquire);
        const bool stopOnError     = context.parallel->stopOnErrorRequested.load(std::memory_order_acquire);
        if (cancelRequested || stopOnError)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    if (! context.callback)
    {
        return S_OK;
    }

    constexpr ULONGLONG kMinProgressMsCopyMove = 50ull;
    constexpr ULONGLONG kMinProgressMsDelete   = 100ull;
    const ULONGLONG minProgressMs              = (context.type == FILESYSTEM_DELETE) ? kMinProgressMsDelete : kMinProgressMsCopyMove;

    const unsigned long completedItems = context.parallel ? context.parallel->completedItems.load(std::memory_order_acquire) : context.completedItems;
    const uint64_t completedBytes      = context.parallel ? context.parallel->completedBytes.load(std::memory_order_acquire) : context.completedBytes;

    const bool isFinalItem    = currentItemTotalBytes > 0 && currentItemCompletedBytes >= currentItemTotalBytes;
    const bool isFinalOverall = context.totalItems > 0 && completedItems >= context.totalItems;
    const bool isFinal        = isFinalItem || isFinalOverall;

    const ULONGLONG nowTick = GetTickCount64();
    if (! isFinal && context.lastProgressReportTick != 0 && nowTick >= context.lastProgressReportTick &&
        (nowTick - context.lastProgressReportTick) < minProgressMs)
    {
        return S_OK;
    }

    if (context.parallel)
    {
        std::scoped_lock lock(context.parallel->callbackMutex);

        if (context.type == FILESYSTEM_DELETE && ! isFinal && context.parallel->lastProgressReportTick != 0 &&
            nowTick >= context.parallel->lastProgressReportTick && (nowTick - context.parallel->lastProgressReportTick) < minProgressMs)
        {
            return S_OK;
        }

        HRESULT hr = context.callback->FileSystemProgress(context.type,
                                                          context.totalItems,
                                                          completedItems,
                                                          context.totalBytes,
                                                          completedBytes,
                                                          context.progressSource,
                                                          context.progressDestination,
                                                          currentItemTotalBytes,
                                                          currentItemCompletedBytes,
                                                          context.options,
                                                          context.progressStreamId,
                                                          context.callbackCookie);
        hr         = NormalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (context.options)
        {
            context.parallel->bandwidthLimitBytesPerSecond.store(context.options->bandwidthLimitBytesPerSecond, std::memory_order_release);
        }

        context.lastProgressReportTick           = nowTick;
        context.parallel->lastProgressReportTick = nowTick;

        return CheckCancelLocked(context);
    }

    HRESULT hr = context.callback->FileSystemProgress(context.type,
                                                      context.totalItems,
                                                      completedItems,
                                                      context.totalBytes,
                                                      completedBytes,
                                                      context.progressSource,
                                                      context.progressDestination,
                                                      currentItemTotalBytes,
                                                      currentItemCompletedBytes,
                                                      context.options,
                                                      context.progressStreamId,
                                                      context.callbackCookie);
    hr         = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    context.lastProgressReportTick = nowTick;

    return CheckCancel(context);
}

HRESULT ReportProgressForced(OperationContext& context, uint64_t currentItemTotalBytes, uint64_t currentItemCompletedBytes) noexcept
{
    if (context.parallel)
    {
        const bool cancelRequested = context.parallel->cancelRequested.load(std::memory_order_acquire);
        const bool stopOnError     = context.parallel->stopOnErrorRequested.load(std::memory_order_acquire);
        if (cancelRequested || stopOnError)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    if (! context.callback)
    {
        return S_OK;
    }

    const unsigned long completedItems = context.parallel ? context.parallel->completedItems.load(std::memory_order_acquire) : context.completedItems;
    const uint64_t completedBytes      = context.parallel ? context.parallel->completedBytes.load(std::memory_order_acquire) : context.completedBytes;

    const ULONGLONG nowTick = GetTickCount64();
    if (context.parallel)
    {
        std::scoped_lock lock(context.parallel->callbackMutex);

        HRESULT hr = context.callback->FileSystemProgress(context.type,
                                                          context.totalItems,
                                                          completedItems,
                                                          context.totalBytes,
                                                          completedBytes,
                                                          context.progressSource,
                                                          context.progressDestination,
                                                          currentItemTotalBytes,
                                                          currentItemCompletedBytes,
                                                          context.options,
                                                          context.progressStreamId,
                                                          context.callbackCookie);
        hr         = NormalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (context.options)
        {
            context.parallel->bandwidthLimitBytesPerSecond.store(context.options->bandwidthLimitBytesPerSecond, std::memory_order_release);
        }

        context.lastProgressReportTick           = nowTick;
        context.parallel->lastProgressReportTick = nowTick;
        return CheckCancelLocked(context);
    }

    HRESULT hr = context.callback->FileSystemProgress(context.type,
                                                      context.totalItems,
                                                      completedItems,
                                                      context.totalBytes,
                                                      completedBytes,
                                                      context.progressSource,
                                                      context.progressDestination,
                                                      currentItemTotalBytes,
                                                      currentItemCompletedBytes,
                                                      context.options,
                                                      context.progressStreamId,
                                                      context.callbackCookie);
    hr         = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    context.lastProgressReportTick = nowTick;
    return CheckCancel(context);
}

HRESULT ReportItemCompleted(OperationContext& context, unsigned long itemIndex, HRESULT status) noexcept
{
    if (! context.callback)
    {
        return S_OK;
    }

    if (context.parallel)
    {
        std::scoped_lock lock(context.parallel->callbackMutex);

        HRESULT hr = context.callback->FileSystemItemCompleted(
            context.type, itemIndex, context.itemSource, context.itemDestination, status, context.options, context.callbackCookie);
        hr = NormalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (context.options)
        {
            context.parallel->bandwidthLimitBytesPerSecond.store(context.options->bandwidthLimitBytesPerSecond, std::memory_order_release);
        }

        return CheckCancelLocked(context);
    }

    HRESULT hr = context.callback->FileSystemItemCompleted(
        context.type, itemIndex, context.itemSource, context.itemDestination, status, context.options, context.callbackCookie);
    hr = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    return CheckCancel(context);
}

HRESULT ReportIssue(OperationContext& context, HRESULT status, FileSystemIssueAction* action) noexcept
{
    if (! action)
    {
        return E_POINTER;
    }

    *action = FileSystemIssueAction::Cancel;

    if (! context.callback)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (context.parallel)
    {
        std::scoped_lock lock(context.parallel->callbackMutex);

        HRESULT hr = context.callback->FileSystemIssue(
            context.type, context.progressSource, context.progressDestination, status, action, context.options, context.callbackCookie);
        hr = NormalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }

        if (context.options)
        {
            context.parallel->bandwidthLimitBytesPerSecond.store(context.options->bandwidthLimitBytesPerSecond, std::memory_order_release);
        }

        return CheckCancelLocked(context);
    }

    HRESULT hr = context.callback->FileSystemIssue(
        context.type, context.progressSource, context.progressDestination, status, action, context.options, context.callbackCookie);
    hr = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    return CheckCancel(context);
}

void InitializeOperationContext(OperationContext& context,
                                FileSystemOperation type,
                                FileSystemFlags flags,
                                const FileSystemOptions* options,
                                IFileSystemCallback* callback,
                                void* cookie,
                                unsigned long totalItems,
                                FileSystemReparsePointPolicy reparsePointPolicy) noexcept
{
    context.type             = type;
    context.callback         = callback;
    context.callbackCookie   = callback != nullptr ? cookie : nullptr;
    context.progressStreamId = 0;
    context.optionsState     = {};
    if (options)
    {
        context.optionsState = *options;
    }
    context.optionsState.sizeBytes      = sizeof(FileSystemOptions);
    context.options                     = &context.optionsState;
    context.totalItems                  = totalItems;
    context.completedItems              = 0;
    context.totalBytes                  = 0;
    context.completedBytes              = 0;
    context.bandwidthThrottleWorkerMode = GetBandwidthThrottleWorkerModeOverride();
    context.continueOnError             = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    context.allowOverwrite              = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);
    context.allowReplaceReadonly        = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    context.recursive                   = HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE);
    context.useRecycleBin               = HasFlag(flags, FILESYSTEM_FLAG_USE_RECYCLE_BIN);
    context.deleteConcurrencyBudget     = 1;
    context.itemSource                  = nullptr;
    context.itemDestination             = nullptr;
    context.progressSource              = nullptr;
    context.progressDestination         = nullptr;
    context.reparsePointPolicy          = reparsePointPolicy;
    context.reparseRootSourcePath.clear();
    context.reparseRootDestinationPath.clear();
}

class ScopedCopyMoveTransferSlot final
{
public:
    ScopedCopyMoveTransferSlot() noexcept                                    = default;
    ScopedCopyMoveTransferSlot(const ScopedCopyMoveTransferSlot&)            = delete;
    ScopedCopyMoveTransferSlot(ScopedCopyMoveTransferSlot&&)                 = delete;
    ScopedCopyMoveTransferSlot& operator=(const ScopedCopyMoveTransferSlot&) = delete;
    ScopedCopyMoveTransferSlot& operator=(ScopedCopyMoveTransferSlot&&)      = delete;

    ~ScopedCopyMoveTransferSlot() noexcept
    {
        Release();
    }

    [[nodiscard]] HRESULT Acquire(OperationContext& context) noexcept
    {
        ParallelOperationState* parallel = context.parallel;
        if (! parallel)
        {
            return S_OK;
        }

        const unsigned int limit = parallel->copyMoveTransferLimit.load(std::memory_order_acquire);
        if (limit == 0)
        {
            return S_OK;
        }

        std::unique_lock lock(parallel->copyMoveTransferMutex);
        while (parallel->activeCopyMoveTransfers >= limit)
        {
            if (parallel->cancelRequested.load(std::memory_order_acquire) || parallel->stopOnErrorRequested.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            parallel->copyMoveTransferCv.wait_for(lock, std::chrono::milliseconds(25));
        }

        ++parallel->activeCopyMoveTransfers;
        _parallel = parallel;
        _acquired = true;
        return S_OK;
    }

    void Release() noexcept
    {
        if (! _acquired || ! _parallel)
        {
            return;
        }

        {
            std::scoped_lock lock(_parallel->copyMoveTransferMutex);
            if (_parallel->activeCopyMoveTransfers > 0)
            {
                --_parallel->activeCopyMoveTransfers;
            }
        }

        _parallel->copyMoveTransferCv.notify_one();
        _parallel = nullptr;
        _acquired = false;
    }

private:
    ParallelOperationState* _parallel = nullptr;
    bool _acquired                    = false;
};

HRESULT GetFileSizeBytes(const std::wstring& path, uint64_t* sizeBytes) noexcept
{
    if (! sizeBytes)
    {
        return E_POINTER;
    }

    *sizeBytes = 0;

    WIN32_FILE_ATTRIBUTE_DATA data{};
    if (! GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &data))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        return S_OK;
    }

    const uint64_t high = static_cast<uint64_t>(data.nFileSizeHigh);
    const uint64_t low  = static_cast<uint64_t>(data.nFileSizeLow);
    *sizeBytes          = (high << 32) | low;
    return S_OK;
}

HRESULT GetPathBasicInformation(const std::wstring& path, FILE_BASIC_INFO* info) noexcept
{
    if (! info)
    {
        return E_POINTER;
    }

    wil::unique_handle handle(::CreateFileW(path.c_str(),
                                            FILE_READ_ATTRIBUTES,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_BACKUP_SEMANTICS,
                                            nullptr));
    if (! handle)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    if (! ::GetFileInformationByHandleEx(handle.get(), FileBasicInfo, info, sizeof(*info)))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return S_OK;
}

FILETIME FileTimeFromQuadPart(LONGLONG value) noexcept
{
    ULARGE_INTEGER raw{};
    raw.QuadPart = static_cast<ULONGLONG>(value);

    FILETIME fileTime{};
    fileTime.dwLowDateTime  = raw.LowPart;
    fileTime.dwHighDateTime = raw.HighPart;
    return fileTime;
}

HRESULT SetPathBasicInformation(const std::wstring& path, const FILE_BASIC_INFO& info) noexcept
{
    wil::unique_handle handle(::CreateFileW(path.c_str(),
                                            FILE_WRITE_ATTRIBUTES,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_BACKUP_SEMANTICS,
                                            nullptr));
    if (! handle)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    FILE_BASIC_INFO mutableInfo = info;
    if (! ::SetFileInformationByHandle(handle.get(), FileBasicInfo, &mutableInfo, sizeof(mutableInfo)))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const FILETIME creationTime   = FileTimeFromQuadPart(info.CreationTime.QuadPart);
    const FILETIME lastAccessTime = FileTimeFromQuadPart(info.LastAccessTime.QuadPart);
    const FILETIME lastWriteTime  = FileTimeFromQuadPart(info.LastWriteTime.QuadPart);
    if (! ::SetFileTime(handle.get(), &creationTime, &lastAccessTime, &lastWriteTime))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return S_OK;
}

void CopyPathBasicInformationBestEffort(const std::wstring& sourcePath, const std::wstring& destinationPath) noexcept
{
    FILE_BASIC_INFO basic{};
    if (FAILED(GetPathBasicInformation(sourcePath, &basic)))
    {
        return;
    }

    static_cast<void>(SetPathBasicInformation(destinationPath, basic));
}

[[nodiscard]] bool TryRollbackCopiedDestination(const std::wstring& pathExtended) noexcept
{
    OperationContext cleanupContext{};
    cleanupContext.allowReplaceReadonly = true;
    return SUCCEEDED(RemovePathForOverwrite(cleanupContext, pathExtended));
}

HRESULT RemoveDirectoryRecursiveNoFollow(OperationContext& context, const std::wstring& directoryExtended) noexcept
{
    HRESULT hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring searchPattern = AppendPath(directoryExtended, L"*");
    WIN32_FIND_DATAW data{};
    wil::unique_hfind findHandle(FindFirstFileExW(searchPattern.c_str(), FindExInfoBasic, &data, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
    if (! findHandle)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND)
        {
            return S_OK;
        }
        return HRESULT_FROM_WIN32(error);
    }

    do
    {
        if (IsDotOrDotDot(data.cFileName))
        {
            continue;
        }

        const std::wstring child = AppendPath(directoryExtended, data.cFileName);
        const DWORD attributes   = data.dwFileAttributes;

        if (IsDirectory(attributes))
        {
            if (IsReparsePoint(attributes))
            {
                if (! RemoveDirectoryW(child.c_str()))
                {
                    return HRESULT_FROM_WIN32(GetLastError());
                }
            }
            else
            {
                hr = RemoveDirectoryRecursiveNoFollow(context, child);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
        }
        else
        {
            bool restoreChildAttributes      = false;
            auto restoreChildAttributesScope = wil::scope_exit([&]() noexcept
            {
                if (restoreChildAttributes)
                {
                    static_cast<void>(SetFileAttributesW(child.c_str(), attributes));
                }
            });

            if ((attributes & FILE_ATTRIBUTE_READONLY) != 0)
            {
                if (! context.allowReplaceReadonly)
                {
                    return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
                }

                const DWORD newAttributes = attributes & ~FILE_ATTRIBUTE_READONLY;
                if (! SetFileAttributesW(child.c_str(), newAttributes))
                {
                    return HRESULT_FROM_WIN32(GetLastError());
                }

                restoreChildAttributes = true;
            }

            if (! DeleteFileW(child.c_str()))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }

            restoreChildAttributes = false;
        }

        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }
    } while (FindNextFileW(findHandle.get(), &data));

    const DWORD error = GetLastError();
    if (error != ERROR_NO_MORE_FILES)
    {
        return HRESULT_FROM_WIN32(error);
    }

    DWORD dirAttributes = GetFileAttributesW(directoryExtended.c_str());
    if (dirAttributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const DWORD originalDirAttributes = dirAttributes;
    bool restoreDirAttributes         = false;
    auto restoreDirAttributesScope    = wil::scope_exit([&]() noexcept
    {
        if (restoreDirAttributes)
        {
            static_cast<void>(SetFileAttributesW(directoryExtended.c_str(), originalDirAttributes));
        }
    });

    if ((dirAttributes & FILE_ATTRIBUTE_READONLY) != 0)
    {
        if (! context.allowReplaceReadonly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        dirAttributes &= ~FILE_ATTRIBUTE_READONLY;
        if (! SetFileAttributesW(directoryExtended.c_str(), dirAttributes))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        restoreDirAttributes = true;
    }

    if (! RemoveDirectoryW(directoryExtended.c_str()))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    restoreDirAttributes = false;
    return S_OK;
}

HRESULT RemovePathForOverwrite(OperationContext& context, const std::wstring& pathExtended) noexcept
{
    HRESULT hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    DWORD attributes = GetFileAttributesW(pathExtended.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if (IsDirectory(attributes))
    {
        if (IsReparsePoint(attributes))
        {
            if (! RemoveDirectoryW(pathExtended.c_str()))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            return S_OK;
        }

        return RemoveDirectoryRecursiveNoFollow(context, pathExtended);
    }

    if ((attributes & FILE_ATTRIBUTE_READONLY) != 0)
    {
        bool restoreAttributes      = false;
        auto restoreAttributesScope = wil::scope_exit([&]() noexcept
        {
            if (restoreAttributes)
            {
                static_cast<void>(SetFileAttributesW(pathExtended.c_str(), attributes));
            }
        });

        if (! context.allowReplaceReadonly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        const DWORD newAttributes = attributes & ~FILE_ATTRIBUTE_READONLY;
        if (! SetFileAttributesW(pathExtended.c_str(), newAttributes))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        restoreAttributes = true;

        if (! DeleteFileW(pathExtended.c_str()))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        restoreAttributes = false;
        return S_OK;
    }

    if (! DeleteFileW(pathExtended.c_str()))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

DWORD CALLBACK CopyProgressRoutine(LARGE_INTEGER totalFileSize,
                                   LARGE_INTEGER totalBytesTransferred,
                                   [[maybe_unused]] LARGE_INTEGER streamSize,
                                   [[maybe_unused]] LARGE_INTEGER streamBytesTransferred,
                                   [[maybe_unused]] DWORD streamNumber,
                                   [[maybe_unused]] DWORD callbackReason,
                                   [[maybe_unused]] HANDLE sourceFile,
                                   [[maybe_unused]] HANDLE destinationFile,
                                   LPVOID context) noexcept
{
    auto* progressContext = static_cast<CopyProgressContext*>(context);
    if (! progressContext || ! progressContext->context)
    {
        return PROGRESS_CONTINUE;
    }

    OperationContext& opContext         = *progressContext->context;
    const uint64_t itemTotal            = static_cast<uint64_t>(totalFileSize.QuadPart);
    const uint64_t itemCompleted        = static_cast<uint64_t>(totalBytesTransferred.QuadPart);
    progressContext->lastItemTotalBytes = itemTotal;

    uint64_t deltaBytes = 0;
    if (itemCompleted >= progressContext->lastItemBytesTransferred)
    {
        deltaBytes                                = itemCompleted - progressContext->lastItemBytesTransferred;
        progressContext->lastItemBytesTransferred = itemCompleted;
    }
    else
    {
        // Defensive: restart delta tracking if the API reports a smaller value.
        progressContext->lastItemBytesTransferred = itemCompleted;
    }

    if (opContext.parallel)
    {
        const bool cancelRequested = opContext.parallel->cancelRequested.load(std::memory_order_acquire);
        const bool stopOnError     = opContext.parallel->stopOnErrorRequested.load(std::memory_order_acquire);
        if (cancelRequested || stopOnError)
        {
            return PROGRESS_CANCEL;
        }

        if (deltaBytes > 0)
        {
            opContext.parallel->completedBytes.fetch_add(deltaBytes, std::memory_order_acq_rel);
        }

        HRESULT hr = ApplyBandwidthThrottle(opContext, *progressContext, deltaBytes);
        if (FAILED(hr))
        {
            return PROGRESS_CANCEL;
        }

        hr = ReportProgress(opContext, itemTotal, itemCompleted);
        if (FAILED(hr))
        {
            return PROGRESS_CANCEL;
        }
    }
    else
    {
        std::scoped_lock lock(progressContext->progressMutex);
        opContext.completedBytes = progressContext->itemBaseBytes + itemCompleted;

        const HRESULT hr = ApplyBandwidthThrottle(opContext, *progressContext, deltaBytes);
        if (FAILED(hr))
        {
            return PROGRESS_CANCEL;
        }

        const HRESULT progressHr = ReportProgress(opContext, itemTotal, itemCompleted);
        if (FAILED(progressHr))
        {
            return PROGRESS_CANCEL;
        }
    }

    return PROGRESS_CONTINUE;
}

HRESULT CopyFileInternal(OperationContext& context, const PathInfo& source, const PathInfo& destination, uint64_t* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    HRESULT hr = SetProgressPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    const auto returnFailure = [&](HRESULT failure, uint64_t currentItemTotalBytes = 0, uint64_t currentItemCompletedBytes = 0) noexcept -> HRESULT
    {
        const HRESULT progressHr = ReportProgressForced(context, currentItemTotalBytes, currentItemCompletedBytes);
        if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return failure;
    };

    DWORD destinationAttributes          = GetFileAttributesW(destination.extended.c_str());
    const bool destinationExisted        = destinationAttributes != INVALID_FILE_ATTRIBUTES;
    bool restoreDestinationReadonly      = false;
    auto restoreDestinationReadonlyScope = wil::scope_exit([&]() noexcept
    {
        if (restoreDestinationReadonly)
        {
            static_cast<void>(SetFileAttributesW(destination.extended.c_str(), destinationAttributes));
        }
    });

    if (destinationExisted)
    {
        if (! context.allowOverwrite)
        {
            return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
        }

        if ((destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0)
        {
            if (! context.allowReplaceReadonly)
            {
                return returnFailure(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
            }

            const DWORD newAttributes = destinationAttributes & ~FILE_ATTRIBUTE_READONLY;
            if (! SetFileAttributesW(destination.extended.c_str(), newAttributes))
            {
                return returnFailure(HRESULT_FROM_WIN32(GetLastError()));
            }

            restoreDestinationReadonly = true;
        }
    }

    uint64_t fileBytes = 0;
    hr                 = GetFileSizeBytes(source.extended, &fileBytes);
    if (FAILED(hr))
    {
        return returnFailure(hr);
    }

    CopyProgressContext progress{};
    progress.context = &context;
    if (! context.parallel)
    {
        progress.itemBaseBytes = context.completedBytes;
        progress.startTick     = GetTickCount64();
        progress.throttleWindowSamples.emplace_back(progress.startTick, 0);
    }

    ScopedCopyMoveTransferSlot transferSlot{};
    hr = transferSlot.Acquire(context);
    if (FAILED(hr))
    {
        return returnFailure(hr, fileBytes, 0);
    }

    const DWORD copyFlags = context.allowOverwrite ? 0u : COPY_FILE_FAIL_IF_EXISTS;
    if (! CopyFileExW(source.extended.c_str(), destination.extended.c_str(), CopyProgressRoutine, &progress, nullptr, copyFlags))
    {
        const DWORD error = GetLastError();
        if (! destinationExisted)
        {
            static_cast<void>(TryRollbackCopiedDestination(destination.extended));
        }
        if (error == ERROR_REQUEST_ABORTED || error == ERROR_CANCELLED)
        {
            const HRESULT hrCancelled = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            EmitSequentialThrottleSummary(context, progress, fileBytes, hrCancelled);
            return hrCancelled;
        }
        const HRESULT hrFailure = returnFailure(HRESULT_FROM_WIN32(error), fileBytes, progress.lastItemBytesTransferred);
        EmitSequentialThrottleSummary(context, progress, fileBytes, hrFailure);
        return hrFailure;
    }

    restoreDestinationReadonly = false;
    *bytesCopied               = fileBytes;
    if (context.parallel)
    {
        if (fileBytes > progress.lastItemBytesTransferred)
        {
            context.parallel->completedBytes.fetch_add(fileBytes - progress.lastItemBytesTransferred, std::memory_order_acq_rel);
            progress.lastItemBytesTransferred = fileBytes;
        }
    }
    else
    {
        context.completedBytes = progress.itemBaseBytes + fileBytes;
    }

    const HRESULT progressHr = ReportProgressForced(context, fileBytes, fileBytes);
    if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    if (FAILED(progressHr))
    {
        EmitSequentialThrottleSummary(context, progress, fileBytes, progressHr);
        return progressHr;
    }
    EmitSequentialThrottleSummary(context, progress, fileBytes, S_OK);
    return S_OK;
}

HRESULT
CopyReparsePointInternal(OperationContext& context, const PathInfo& source, const PathInfo& destination, DWORD sourceAttributes, uint64_t* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    HRESULT hr = SetProgressPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    const auto returnFailure = [&](HRESULT failure, uint64_t currentItemTotalBytes = 0, uint64_t currentItemCompletedBytes = 0) noexcept -> HRESULT
    {
        const HRESULT progressHr = ReportProgressForced(context, currentItemTotalBytes, currentItemCompletedBytes);
        if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return failure;
    };

    const bool isDirectory = IsDirectory(sourceAttributes);
    if (! isDirectory)
    {
        // Copy file reparse points as links only. Never silently fall back to dereferencing data copy.
        uint64_t fileBytes   = 0;
        const HRESULT sizeHr = GetFileSizeBytes(source.extended, &fileBytes);
        if (FAILED(sizeHr))
        {
            return returnFailure(sizeHr);
        }

        CopyProgressContext progress{};
        progress.context = &context;
        if (! context.parallel)
        {
            progress.itemBaseBytes = context.completedBytes;
            progress.startTick     = GetTickCount64();
            progress.throttleWindowSamples.emplace_back(progress.startTick, 0);
        }

        DWORD destinationAttributes          = GetFileAttributesW(destination.extended.c_str());
        const bool destinationExisted        = destinationAttributes != INVALID_FILE_ATTRIBUTES;
        bool restoreDestinationReadonly      = false;
        auto restoreDestinationReadonlyScope = wil::scope_exit([&]() noexcept
        {
            if (restoreDestinationReadonly)
            {
                static_cast<void>(SetFileAttributesW(destination.extended.c_str(), destinationAttributes));
            }
        });

        if (destinationExisted && ! context.allowOverwrite)
        {
            return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
        }

        if (destinationExisted && (destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0)
        {
            if (! context.allowReplaceReadonly)
            {
                return returnFailure(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
            }

            const DWORD newAttributes = destinationAttributes & ~FILE_ATTRIBUTE_READONLY;
            if (! SetFileAttributesW(destination.extended.c_str(), newAttributes))
            {
                return returnFailure(HRESULT_FROM_WIN32(GetLastError()));
            }

            restoreDestinationReadonly = true;
        }

        const DWORD overwriteFlag = context.allowOverwrite ? 0u : COPY_FILE_FAIL_IF_EXISTS;
        const DWORD copyFlags     = overwriteFlag | COPY_FILE_COPY_SYMLINK;
        if (! CopyFileExW(source.extended.c_str(), destination.extended.c_str(), CopyProgressRoutine, &progress, nullptr, copyFlags))
        {
            const DWORD error = GetLastError();
            if (! destinationExisted)
            {
                static_cast<void>(TryRollbackCopiedDestination(destination.extended));
            }
            if (error == ERROR_REQUEST_ABORTED || error == ERROR_CANCELLED)
            {
                const HRESULT hrCancelled = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                EmitSequentialThrottleSummary(context, progress, fileBytes, hrCancelled);
                return hrCancelled;
            }
            if (error == ERROR_INVALID_PARAMETER)
            {
                const HRESULT hrFailure = returnFailure(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), fileBytes, progress.lastItemBytesTransferred);
                EmitSequentialThrottleSummary(context, progress, fileBytes, hrFailure);
                return hrFailure;
            }
            const HRESULT hrFailure = returnFailure(HRESULT_FROM_WIN32(error), fileBytes, progress.lastItemBytesTransferred);
            EmitSequentialThrottleSummary(context, progress, fileBytes, hrFailure);
            return hrFailure;
        }

        restoreDestinationReadonly = false;
        *bytesCopied               = fileBytes;
        if (context.parallel)
        {
            if (fileBytes > progress.lastItemBytesTransferred)
            {
                context.parallel->completedBytes.fetch_add(fileBytes - progress.lastItemBytesTransferred, std::memory_order_acq_rel);
                progress.lastItemBytesTransferred = fileBytes;
            }
        }
        else
        {
            context.completedBytes = progress.itemBaseBytes + fileBytes;
        }

        const HRESULT progressHr = ReportProgressForced(context, fileBytes, fileBytes);
        if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (FAILED(progressHr))
        {
            EmitSequentialThrottleSummary(context, progress, fileBytes, progressHr);
            return progressHr;
        }

        EmitSequentialThrottleSummary(context, progress, fileBytes, S_OK);
        return S_OK;
    }

    // Directory reparse points are handled explicitly to prevent recursive traversal (junction/symlink loops).
    ReparsePointData reparse{};
    hr = ReadReparsePointData(source.extended, reparse);
    if (FAILED(hr))
    {
        return returnFailure(hr);
    }

    if (reparse.tag != IO_REPARSE_TAG_SYMLINK && reparse.tag != IO_REPARSE_TAG_MOUNT_POINT)
    {
        return returnFailure(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    }

    ParsedDirectoryReparsePoint parsed{};
    if (! ParseDirectoryReparsePoint(reparse, parsed))
    {
        return returnFailure(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
    }

    const DWORD destinationAttributes = GetFileAttributesW(destination.extended.c_str());
    if (destinationAttributes != INVALID_FILE_ATTRIBUTES)
    {
        if (! context.allowOverwrite)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        hr = RemovePathForOverwrite(context, destination.extended);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    if (! CreateDirectoryW(destination.extended.c_str(), nullptr))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    bool created = true;
    auto cleanup = wil::scope_exit([&]
    {
        if (created)
        {
            static_cast<void>(RemoveDirectoryW(destination.extended.c_str()));
        }
    });

    std::wstring targetPath = ResolveReparseTargetAbsolute(source, parsed);
    if (targetPath.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const bool preserveTrailingSeparator = EndsWithSeparator(parsed.substitutePath) || EndsWithSeparator(parsed.printPath);
    if (preserveTrailingSeparator && ! EndsWithSeparator(targetPath))
    {
        targetPath.push_back(L'\\');
    }

    if (! context.reparseRootSourcePath.empty() && ! context.reparseRootDestinationPath.empty())
    {
        std::wstring mappedTargetPath;
        if (TryRetargetPathIntoDestination(targetPath, context.reparseRootSourcePath, context.reparseRootDestinationPath, mappedTargetPath))
        {
            targetPath = std::move(mappedTargetPath);
            if (preserveTrailingSeparator && ! EndsWithSeparator(targetPath))
            {
                targetPath.push_back(L'\\');
            }
        }
    }

    ReparsePointData rebuilt{};
    if (reparse.tag == IO_REPARSE_TAG_MOUNT_POINT)
    {
        hr = BuildMountPointReparseData(targetPath, rebuilt);
    }
    else
    {
        bool useRelative           = parsed.isRelative;
        std::wstring symlinkTarget = targetPath;
        if (parsed.isRelative)
        {
            const std::filesystem::path destinationParent = std::filesystem::path(destination.display).parent_path();
            const std::filesystem::path relativeTarget    = std::filesystem::path(targetPath).lexically_relative(destinationParent);
            if (relativeTarget.empty() || relativeTarget.native().empty())
            {
                useRelative = false;
            }
            else
            {
                symlinkTarget = relativeTarget.wstring();
            }
        }

        hr = BuildSymlinkReparseData(symlinkTarget, useRelative, rebuilt);
    }
    if (FAILED(hr))
    {
        return hr;
    }

    hr = WriteReparsePointData(destination.extended, rebuilt);
    if (FAILED(hr))
    {
        return hr;
    }

    CopyPathBasicInformationBestEffort(source.extended, destination.extended);
    created = false;
    return S_OK;
}

HRESULT CopyDirectoryInternal(OperationContext& context, const PathInfo& source, const PathInfo& destination, uint64_t* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    HRESULT hr = SetProgressPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    const auto returnFailure = [&](HRESULT failure) noexcept -> HRESULT
    {
        const HRESULT progressHr = ReportProgressForced(context, 0, 0);
        if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return failure;
    };

    DWORD destinationAttributes = GetFileAttributesW(destination.extended.c_str());
    bool createdDestination     = false;
    if (destinationAttributes == INVALID_FILE_ATTRIBUTES)
    {
        if (! CreateDirectoryW(destination.extended.c_str(), nullptr))
        {
            return returnFailure(HRESULT_FROM_WIN32(GetLastError()));
        }

        createdDestination = true;
    }
    else
    {
        if ((destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
        {
            return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
        }
        if (! context.allowOverwrite)
        {
            return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
        }
    }

    std::wstring searchPattern = AppendPath(source.extended, L"*");
    WIN32_FIND_DATAW data{};
    wil::unique_hfind findHandle(FindFirstFileExW(searchPattern.c_str(), FindExInfoBasic, &data, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
    if (! findHandle)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND)
        {
            if (createdDestination)
            {
                CopyPathBasicInformationBestEffort(source.extended, destination.extended);
            }
            return S_OK;
        }
        return returnFailure(HRESULT_FROM_WIN32(error));
    }

    bool hadFailure = false;
    bool hadSkipped = false;

    do
    {
        if (IsDotOrDotDot(data.cFileName))
        {
            continue;
        }

        PathInfo childSource{};
        childSource.display  = AppendPath(source.display, data.cFileName);
        childSource.extended = AppendPath(source.extended, data.cFileName);

        PathInfo childDestination{};
        childDestination.display  = AppendPath(destination.display, data.cFileName);
        childDestination.extended = AppendPath(destination.extended, data.cFileName);

        uint64_t childBytes = 0;
        HRESULT childHr     = S_OK;

        const DWORD childAttributes = data.dwFileAttributes;
        const bool childIsDirectory = IsDirectory(childAttributes);
        const bool childIsReparse   = IsReparsePoint(childAttributes);

        for (;;)
        {
            childBytes = 0;
            childHr    = S_OK;

            if (childIsDirectory)
            {
                if (childIsReparse && context.reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
                {
                    if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
                    {
                        hadSkipped = true;
                        childHr    = S_OK;
                    }
                    else
                    {
                        childHr = CopyReparsePointInternal(context, childSource, childDestination, childAttributes, &childBytes);
                    }
                }
                else if (! context.recursive)
                {
                    childHr = HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
                }
                else
                {
                    childHr = CopyDirectoryInternal(context, childSource, childDestination, &childBytes);
                }
            }
            else if (childIsReparse && context.reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
            {
                if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
                {
                    hadSkipped = true;
                    childHr    = S_OK;
                }
                else
                {
                    childHr = CopyReparsePointInternal(context, childSource, childDestination, childAttributes, &childBytes);
                }
            }
            else
            {
                childHr = CopyFileInternal(context, childSource, childDestination, &childBytes);
            }

            if (SUCCEEDED(childHr))
            {
                break;
            }

            childHr = NormalizeCancellation(childHr);
            if (IsCancellationHr(childHr))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            if (context.continueOnError)
            {
                hadFailure = true;
                childHr    = S_OK;
                break;
            }

            FileSystemIssueAction issueAction = FileSystemIssueAction::Cancel;
            const HRESULT issueHr             = ReportIssue(context, childHr, &issueAction);
            if (FAILED(issueHr))
            {
                return issueHr;
            }

            switch (issueAction)
            {
                case FileSystemIssueAction::Overwrite: context.allowOverwrite = true; continue;
                case FileSystemIssueAction::ReplaceReadOnly: context.allowReplaceReadonly = true; continue;
                case FileSystemIssueAction::PermanentDelete: context.useRecycleBin = false; continue;
                case FileSystemIssueAction::Retry: continue;
                case FileSystemIssueAction::Skip:
                    hadFailure = true;
                    childHr    = S_OK;
                    break;
                case FileSystemIssueAction::Cancel:
                case FileSystemIssueAction::None:
                default: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            break;
        }

        if (childBytes > 0)
        {
            if (std::numeric_limits<uint64_t>::max() - *bytesCopied < childBytes)
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            *bytesCopied += childBytes;
        }

        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }
    } while (FindNextFileW(findHandle.get(), &data));

    const DWORD error = GetLastError();
    if (error != ERROR_NO_MORE_FILES)
    {
        return returnFailure(HRESULT_FROM_WIN32(error));
    }

    findHandle.reset();

    if (createdDestination)
    {
        CopyPathBasicInformationBestEffort(source.extended, destination.extended);
    }

    if (hadFailure || hadSkipped)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT CopyPathInternal(OperationContext& context, const PathInfo& source, const PathInfo& destination, uint64_t* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    const DWORD attributes = GetFileAttributesW(source.extended.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        const DWORD error = GetLastError();
        static_cast<void>(SetProgressPaths(context, source.display.c_str(), destination.display.c_str()));
        static_cast<void>(ReportProgressForced(context, 0, 0));
        return HRESULT_FROM_WIN32(error);
    }

    const bool isReparse = IsReparsePoint(attributes);
    if (isReparse && context.reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
    {
        if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return CopyReparsePointInternal(context, source, destination, attributes, bytesCopied);
    }

    if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        if (! context.recursive)
        {
            return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
        }
        return CopyDirectoryInternal(context, source, destination, bytesCopied);
    }

    return CopyFileInternal(context, source, destination, bytesCopied);
}

enum class RecursiveCopyWorkKind : uint8_t
{
    Directory,
    File,
    ReparsePoint,
};

struct RecursiveCopyWorkItem
{
    RecursiveCopyWorkKind kind = RecursiveCopyWorkKind::File;
    PathInfo source;
    PathInfo destination;
    DWORD attributes = 0;
};

[[nodiscard]] HRESULT CopyDirectoryChildrenParallel(OperationContext& rootContext,
                                                    const PathInfo& source,
                                                    const PathInfo& destination,
                                                    FileSystemFlags flags,
                                                    FileSystemReparsePointPolicy reparsePointPolicy,
                                                    unsigned int maxConcurrency,
                                                    uint64_t* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    HRESULT hr = SetProgressPaths(rootContext, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(rootContext);
    if (FAILED(hr))
    {
        return hr;
    }

    const auto returnFailure = [&](HRESULT failure) noexcept -> HRESULT
    {
        const HRESULT progressHr = ReportProgressForced(rootContext, 0, 0);
        if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return failure;
    };

    const unsigned int requestedConcurrency = std::max(1u, maxConcurrency);
    if (requestedConcurrency <= 1u)
    {
        return CopyDirectoryInternal(rootContext, source, destination, bytesCopied);
    }

    const DWORD attributes = GetFileAttributesW(source.extended.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        return returnFailure(HRESULT_FROM_WIN32(GetLastError()));
    }
    if ((attributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
    {
        return CopyPathInternal(rootContext, source, destination, bytesCopied);
    }
    if (IsReparsePoint(attributes) && reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
    {
        return CopyPathInternal(rootContext, source, destination, bytesCopied);
    }

    constexpr unsigned int kMaxWorkers = 16u;
    const unsigned int concurrency     = std::max(2u, std::min<unsigned int>(requestedConcurrency, kMaxWorkers));
    if (! GetSharedFileOpsJobScheduler().EnsureWorkersAvailable())
    {
        Debug::Perf::EmitCounter(L"FileOps.CopyRecursiveParallel.SerialFallback.SchedulerUnavailable");
        return CopyDirectoryInternal(rootContext, source, destination, bytesCopied);
    }

    FileSystemOptions* sharedOptionsState = rootContext.options;

    ParallelOperationState parallelStorage{};
    ParallelOperationState& parallel = rootContext.parallel ? *rootContext.parallel : parallelStorage;
    if (! rootContext.parallel)
    {
        parallel.startTick = GetTickCount64();
        parallel.bandwidthLimitBytesPerSecond.store(sharedOptionsState ? sharedOptionsState->bandwidthLimitBytesPerSecond : 0ull, std::memory_order_release);
    }
    unsigned int currentTransferLimit = parallel.copyMoveTransferLimit.load(std::memory_order_acquire);
    while (currentTransferLimit == 0u && ! parallel.copyMoveTransferLimit.compare_exchange_weak(currentTransferLimit, concurrency, std::memory_order_acq_rel))
    {
    }

    std::atomic<bool> hadFailure{false};
    std::atomic<bool> hadSkipped{false};

    std::atomic<uint64_t> queuedFiles{0};
    std::atomic<uint64_t> queuedDirectories{0};
    std::atomic<uint64_t> queuedReparsePoints{0};
    std::atomic<uint64_t> processedFiles{0};
    std::atomic<uint64_t> processedDirectories{0};

    const std::wstring rootSource      = rootContext.reparseRootSourcePath;
    const std::wstring rootDestination = rootContext.reparseRootDestinationPath;

    struct RecursiveCopyQueue final
    {
        RecursiveCopyQueue()                                     = default;
        RecursiveCopyQueue(const RecursiveCopyQueue&)            = delete;
        RecursiveCopyQueue& operator=(const RecursiveCopyQueue&) = delete;
        RecursiveCopyQueue(RecursiveCopyQueue&&)                 = delete;
        RecursiveCopyQueue& operator=(RecursiveCopyQueue&&)      = delete;

        std::mutex mutex;
        std::condition_variable cv;
        std::deque<RecursiveCopyWorkItem> items;
        size_t activeItems = 0;
        bool done          = false;
    };

    RecursiveCopyQueue queue{};
    const size_t maxQueuedItems = std::max<size_t>(512u, static_cast<size_t>(concurrency) * 128u);

    struct CreatedDirectoryMetadataTarget final
    {
        std::wstring sourcePath;
        std::wstring destinationPath;
    };

    std::mutex createdDirectoriesMutex;
    std::vector<CreatedDirectoryMetadataTarget> createdDirectories;

    const auto rememberCreatedDirectory = [&](const PathInfo& directorySource, const PathInfo& directoryDestination) noexcept
    {
        std::scoped_lock lock(createdDirectoriesMutex);
        createdDirectories.push_back(
            CreatedDirectoryMetadataTarget{.sourcePath = directorySource.extended, .destinationPath = directoryDestination.extended});
    };

    const auto initializeChildContext = [&](OperationContext& context, uint64_t progressStreamId) noexcept
    {
        InitializeOperationContext(
            context, rootContext.type, flags, sharedOptionsState, rootContext.callback, rootContext.callbackCookie, 1, reparsePointPolicy);
        context.options                    = sharedOptionsState;
        context.parallel                   = &parallel;
        context.totalBytes                 = 0; // let the host provide totals via pre-calc
        context.progressStreamId           = progressStreamId;
        context.reparseRootSourcePath      = rootSource;
        context.reparseRootDestinationPath = rootDestination;
    };

    const auto finishActiveItem = [&]() noexcept
    {
        {
            std::scoped_lock lock(queue.mutex);
            if (queue.activeItems > 0)
            {
                --queue.activeItems;
            }
            if (queue.items.empty() && queue.activeItems == 0)
            {
                queue.done = true;
            }
        }
        queue.cv.notify_all();
    };

    const auto enqueueWork = [&](RecursiveCopyWorkItem item) noexcept -> bool
    {
        for (;;)
        {
            if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
            {
                return false;
            }

            {
                std::unique_lock lock(queue.mutex);
                if (queue.done)
                {
                    return false;
                }
                // The queue cap is advisory for scheduler workers. A worker may discover more child
                // work while it is also the consumer; bypassing the cap avoids worker-to-producer
                // self-deadlock while active transfer slots still bound concurrent I/O.
                if (queue.items.size() < maxQueuedItems || GetSharedFileOpsJobScheduler().IsWorkerThread())
                {
                    switch (item.kind)
                    {
                        case RecursiveCopyWorkKind::Directory: queuedDirectories.fetch_add(1, std::memory_order_relaxed); break;
                        case RecursiveCopyWorkKind::File: queuedFiles.fetch_add(1, std::memory_order_relaxed); break;
                        case RecursiveCopyWorkKind::ReparsePoint: queuedReparsePoints.fetch_add(1, std::memory_order_relaxed); break;
                    }

                    queue.items.push_back(std::move(item));
                    lock.unlock();
                    queue.cv.notify_one();
                    return true;
                }

                queue.cv.wait(lock,
                              [&]() noexcept
                {
                    return queue.done || parallel.cancelRequested.load(std::memory_order_acquire) ||
                           parallel.stopOnErrorRequested.load(std::memory_order_acquire) || queue.items.size() < maxQueuedItems;
                });
            }
        }
    };

    const auto completeWithError = [&](HRESULT failure) noexcept
    {
        failure = NormalizeCancellation(failure);
        if (IsCancellationHr(failure))
        {
            parallel.cancelRequested.store(true, std::memory_order_release);
        }
        else
        {
            HRESULT expected = S_OK;
            static_cast<void>(parallel.firstError.compare_exchange_strong(expected, failure, std::memory_order_acq_rel));
            parallel.stopOnErrorRequested.store(true, std::memory_order_release);
        }
        {
            std::scoped_lock lock(queue.mutex);
            queue.done = true;
        }
        queue.cv.notify_all();
    };

    const auto processDirectory = [&](OperationContext& context, const RecursiveCopyWorkItem& item) noexcept -> HRESULT
    {
        HRESULT directoryHr = SetProgressPaths(context, item.source.display.c_str(), item.destination.display.c_str());
        if (FAILED(directoryHr))
        {
            return directoryHr;
        }

        directoryHr = CheckCancel(context);
        if (FAILED(directoryHr))
        {
            return directoryHr;
        }

        DWORD destinationAttributes = GetFileAttributesW(item.destination.extended.c_str());
        bool createdDestination     = false;
        if (destinationAttributes == INVALID_FILE_ATTRIBUTES)
        {
            if (! CreateDirectoryW(item.destination.extended.c_str(), nullptr))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }

            createdDestination = true;
        }
        else
        {
            if ((destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
            {
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }
            if (! context.allowOverwrite)
            {
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }
        }

        std::wstring searchPattern = AppendPath(item.source.extended, L"*");
        WIN32_FIND_DATAW data{};
        wil::unique_hfind findHandle(
            FindFirstFileExW(searchPattern.c_str(), FindExInfoBasic, &data, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
        if (! findHandle)
        {
            const DWORD error = GetLastError();
            if (error == ERROR_FILE_NOT_FOUND)
            {
                if (createdDestination)
                {
                    rememberCreatedDirectory(item.source, item.destination);
                }
                processedDirectories.fetch_add(1, std::memory_order_relaxed);
                return S_OK;
            }
            return HRESULT_FROM_WIN32(error);
        }

        uint64_t cancelCheckCounter = 0;
        do
        {
            if (IsDotOrDotDot(data.cFileName))
            {
                continue;
            }

            PathInfo childSource{};
            childSource.display  = AppendPath(item.source.display, data.cFileName);
            childSource.extended = AppendPath(item.source.extended, data.cFileName);

            PathInfo childDestination{};
            childDestination.display  = AppendPath(item.destination.display, data.cFileName);
            childDestination.extended = AppendPath(item.destination.extended, data.cFileName);

            const DWORD childAttributes = data.dwFileAttributes;
            const bool childIsDirectory = IsDirectory(childAttributes);
            const bool childIsReparse   = IsReparsePoint(childAttributes);

            RecursiveCopyWorkItem childItem{};
            childItem.source      = std::move(childSource);
            childItem.destination = std::move(childDestination);
            childItem.attributes  = childAttributes;

            if (childIsDirectory)
            {
                if (childIsReparse && context.reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
                {
                    if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
                    {
                        hadSkipped.store(true, std::memory_order_release);
                        continue;
                    }

                    childItem.kind = RecursiveCopyWorkKind::ReparsePoint;
                }
                else if (! context.recursive)
                {
                    return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
                }
                else
                {
                    childItem.kind = RecursiveCopyWorkKind::Directory;
                }
            }
            else if (childIsReparse && context.reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
            {
                if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
                {
                    hadSkipped.store(true, std::memory_order_release);
                    continue;
                }

                childItem.kind = RecursiveCopyWorkKind::ReparsePoint;
            }
            else
            {
                childItem.kind = RecursiveCopyWorkKind::File;
            }

            if (! enqueueWork(std::move(childItem)))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            if ((++cancelCheckCounter % 64u) == 0u)
            {
                directoryHr = CheckCancel(context);
                if (FAILED(directoryHr))
                {
                    return directoryHr;
                }
            }
        } while (FindNextFileW(findHandle.get(), &data));

        const DWORD enumError = GetLastError();
        if (enumError != ERROR_NO_MORE_FILES)
        {
            return HRESULT_FROM_WIN32(enumError);
        }

        findHandle.reset();

        if (createdDestination)
        {
            rememberCreatedDirectory(item.source, item.destination);
        }

        processedDirectories.fetch_add(1, std::memory_order_relaxed);
        return S_OK;
    };

    const auto processWork = [&](OperationContext& context, const RecursiveCopyWorkItem& item) noexcept
    {
        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            return;
        }

        for (;;)
        {
            HRESULT itemHr     = S_OK;
            uint64_t itemBytes = 0;
            switch (item.kind)
            {
                case RecursiveCopyWorkKind::Directory: itemHr = processDirectory(context, item); break;
                case RecursiveCopyWorkKind::File:
                    itemHr = CopyFileInternal(context, item.source, item.destination, &itemBytes);
                    if (SUCCEEDED(itemHr))
                    {
                        processedFiles.fetch_add(1, std::memory_order_relaxed);
                    }
                    break;
                case RecursiveCopyWorkKind::ReparsePoint:
                    itemHr = CopyReparsePointInternal(context, item.source, item.destination, item.attributes, &itemBytes);
                    break;
            }

            if (SUCCEEDED(itemHr))
            {
                return;
            }

            itemHr = NormalizeCancellation(itemHr);
            if (IsCancellationHr(itemHr))
            {
                parallel.cancelRequested.store(true, std::memory_order_release);
                {
                    std::scoped_lock lock(queue.mutex);
                    queue.done = true;
                }
                queue.cv.notify_all();
                return;
            }

            if (itemHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
            {
                hadSkipped.store(true, std::memory_order_release);
                break;
            }

            if (context.continueOnError)
            {
                hadFailure.store(true, std::memory_order_release);
                break;
            }

            FileSystemIssueAction issueAction = FileSystemIssueAction::Cancel;
            const HRESULT issueHr             = ReportIssue(context, itemHr, &issueAction);
            if (FAILED(issueHr))
            {
                completeWithError(issueHr);
                return;
            }

            switch (issueAction)
            {
                case FileSystemIssueAction::Overwrite: context.allowOverwrite = true; continue;
                case FileSystemIssueAction::ReplaceReadOnly: context.allowReplaceReadonly = true; continue;
                case FileSystemIssueAction::PermanentDelete: context.useRecycleBin = false; continue;
                case FileSystemIssueAction::Retry: continue;
                case FileSystemIssueAction::Skip: hadFailure.store(true, std::memory_order_release); return;
                case FileSystemIssueAction::Cancel:
                case FileSystemIssueAction::None:
                default:
                    parallel.cancelRequested.store(true, std::memory_order_release);
                    {
                        std::scoped_lock lock(queue.mutex);
                        queue.done = true;
                    }
                    queue.cv.notify_all();
                    return;
            }
        }
    };

    RecursiveCopyWorkItem rootItem{};
    rootItem.kind        = RecursiveCopyWorkKind::Directory;
    rootItem.source      = source;
    rootItem.destination = destination;
    rootItem.attributes  = attributes;
    if (! enqueueWork(std::move(rootItem)))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    auto job = GetSharedFileOpsJobScheduler().StartJob(concurrency,
                                                       concurrency,
                                                       [&](size_t /*index*/, uint64_t schedulerStreamId) noexcept
    {
        OperationContext context{};
        initializeChildContext(context, schedulerStreamId);

        for (;;)
        {
            if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
            {
                return;
            }

            RecursiveCopyWorkItem item{};
            {
                std::unique_lock lock(queue.mutex);
                queue.cv.wait(lock,
                              [&]() noexcept
                {
                    return queue.done || parallel.cancelRequested.load(std::memory_order_acquire) ||
                           parallel.stopOnErrorRequested.load(std::memory_order_acquire) || ! queue.items.empty();
                });

                if (queue.done || parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                {
                    return;
                }

                item = std::move(queue.items.front());
                queue.items.pop_front();
                ++queue.activeItems;
            }

            queue.cv.notify_all();
            processWork(context, item);
            finishActiveItem();
        }
    });

    GetSharedFileOpsJobScheduler().WaitJob(job);

    *bytesCopied = parallel.completedBytes.load(std::memory_order_acquire);

    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.QueuedFiles",
                      L"",
                      queuedFiles.load(std::memory_order_acquire),
                      processedFiles.load(std::memory_order_acquire),
                      concurrency,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.QueuedDirectories",
                      L"",
                      queuedDirectories.load(std::memory_order_acquire),
                      processedDirectories.load(std::memory_order_acquire),
                      concurrency,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.QueuedReparsePoints", L"", queuedReparsePoints.load(std::memory_order_acquire), 0, concurrency, S_OK);

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
    {
        const HRESULT firstError = parallel.firstError.load(std::memory_order_acquire);
        if (FAILED(firstError))
        {
            return returnFailure(firstError);
        }
    }

    std::vector<CreatedDirectoryMetadataTarget> directoriesToRestore;
    {
        std::scoped_lock lock(createdDirectoriesMutex);
        directoriesToRestore = createdDirectories;
    }
    std::sort(directoriesToRestore.begin(),
              directoriesToRestore.end(),
              [](const CreatedDirectoryMetadataTarget& left, const CreatedDirectoryMetadataTarget& right) noexcept
    {
        return left.destinationPath.size() > right.destinationPath.size();
    });
    for (const CreatedDirectoryMetadataTarget& directory : directoriesToRestore)
    {
        CopyPathBasicInformationBestEffort(directory.sourcePath, directory.destinationPath);
    }

    if (hadFailure.load(std::memory_order_acquire) || hadSkipped.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

[[nodiscard]] HRESULT CopyPathInternalWithDirectoryParallelism(OperationContext& context,
                                                               const PathInfo& source,
                                                               const PathInfo& destination,
                                                               FileSystemFlags flags,
                                                               FileSystemReparsePointPolicy reparsePointPolicy,
                                                               unsigned int maxConcurrency,
                                                               uint64_t* bytesCopied) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    constexpr unsigned int kMaxRecursiveCopyConcurrency = 16u;
    const unsigned int effectiveConcurrency             = std::clamp(maxConcurrency, 1u, kMaxRecursiveCopyConcurrency);
    const DWORD attributes                              = GetFileAttributesW(source.extended.c_str());
    const bool canParallelizeDirectory                  = attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 &&
                                                          ! IsReparsePoint(attributes) && context.recursive && effectiveConcurrency > 1u;

    if (canParallelizeDirectory)
    {
        return CopyDirectoryChildrenParallel(context, source, destination, flags, reparsePointPolicy, effectiveConcurrency, bytesCopied);
    }

    if (effectiveConcurrency <= 1u && attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && context.recursive)
    {
        Debug::Perf::EmitCounter(L"FileOps.CopyRecursiveParallel.SerialFallback.MaxConcurrencyOne");
    }

    return CopyPathInternal(context, source, destination, bytesCopied);
}

[[nodiscard]] unsigned int CalculateNestedCopyMoveConcurrency(unsigned int maxConcurrency, unsigned int topLevelConcurrency) noexcept
{
    const unsigned int safeMax = std::max(1u, maxConcurrency);
    const unsigned int safeTop = std::max(1u, topLevelConcurrency);
    if (safeTop <= 1u)
    {
        return safeMax;
    }

    if (safeMax <= 1u)
    {
        return 1u;
    }

    const unsigned int spareBudget = safeMax > safeTop ? safeMax - safeTop : 0u;
    return std::clamp(spareBudget + 1u, 1u, safeMax);
}

HRESULT DeletePathInternal(OperationContext& context, const PathInfo& path) noexcept;

[[nodiscard]] HRESULT RenameCaseOnlyWithTemp(OperationContext& context,
                                             const std::wstring& sourceExtended,
                                             const std::wstring& destinationExtended,
                                             DWORD renameFlags) noexcept
{
    const std::wstring directory = GetPathDirectory(sourceExtended);
    if (directory.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    const DWORD pid      = GetCurrentProcessId();
    const DWORD tid      = GetCurrentThreadId();
    const ULONGLONG tick = GetTickCount64();

    constexpr unsigned int kMaxAttempts = 32;
    for (unsigned int attempt = 0; attempt < kMaxAttempts; ++attempt)
    {
        HRESULT hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }

        std::wstring leaf;
        leaf.reserve(96);
        leaf.append(L".rs_case_tmp_");
        leaf.append(std::to_wstring(pid));
        leaf.push_back(L'_');
        leaf.append(std::to_wstring(tid));
        leaf.push_back(L'_');
        leaf.append(std::to_wstring(tick));
        leaf.push_back(L'_');
        leaf.append(std::to_wstring(attempt));

        std::wstring tempPath = AppendPath(directory, leaf);
        if (tempPath.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const DWORD tempAttributes = ::GetFileAttributesW(tempPath.c_str());
        if (tempAttributes != INVALID_FILE_ATTRIBUTES)
        {
            continue;
        }

        if (! ::MoveFileExW(sourceExtended.c_str(), tempPath.c_str(), renameFlags))
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            const DWORD revertFlags = renameFlags & ~MOVEFILE_REPLACE_EXISTING;
            static_cast<void>(::MoveFileExW(tempPath.c_str(), sourceExtended.c_str(), revertFlags));
            return hr;
        }

        if (! ::MoveFileExW(tempPath.c_str(), destinationExtended.c_str(), renameFlags))
        {
            const DWORD error       = ::GetLastError();
            const DWORD revertFlags = renameFlags & ~MOVEFILE_REPLACE_EXISTING;
            static_cast<void>(::MoveFileExW(tempPath.c_str(), sourceExtended.c_str(), revertFlags));
            return HRESULT_FROM_WIN32(error);
        }

        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
}

HRESULT MovePathInternal(OperationContext& context,
                         const PathInfo& source,
                         const PathInfo& destination,
                         bool allowCopy,
                         FileSystemFlags flags                           = FILESYSTEM_FLAG_NONE,
                         FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse,
                         unsigned int maxCopyConcurrency                 = 1u) noexcept
{
    HRESULT hr = SetProgressPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    const DWORD sourceAttributes = GetFileAttributesW(source.extended.c_str());
    if (sourceAttributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const bool sourceIsDirectory = (sourceAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    const bool sourceIsReparse   = IsReparsePoint(sourceAttributes);

    bool caseOnlyRename         = false;
    DWORD destinationAttributes = GetFileAttributesW(destination.extended.c_str());
    if (destinationAttributes != INVALID_FILE_ATTRIBUTES)
    {
        if (source.extended != destination.extended && OrdinalString::EqualsNoCase(source.extended, destination.extended))
        {
            bool same            = false;
            const HRESULT sameHr = TryAreSameFile(source.extended, destination.extended, same);
            if (FAILED(sameHr))
            {
                return sameHr;
            }

            if (same)
            {
                caseOnlyRename = true;
            }
            else if (! context.allowOverwrite)
            {
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }
        }
        else if (! context.allowOverwrite)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        if (! caseOnlyRename && (destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0)
        {
            if (! context.allowReplaceReadonly)
            {
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }

            const DWORD newAttributes = destinationAttributes & ~FILE_ATTRIBUTE_READONLY;
            if (! SetFileAttributesW(destination.extended.c_str(), newAttributes))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
        }
    }

    DWORD moveFlags = 0;
    if (context.allowOverwrite)
    {
        moveFlags |= MOVEFILE_REPLACE_EXISTING;
    }
    if (allowCopy)
    {
        // Attempt a simple rename first; only fall back to copy+delete when required.
    }

    // Reparse-point policies apply to move operations, not rename.
    if (context.type == FILESYSTEM_MOVE && sourceIsReparse && context.reparsePointPolicy != FileSystemReparsePointPolicy::FollowTargets)
    {
        if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        uint64_t copiedBytes = 0;
        HRESULT copyHr       = CopyReparsePointInternal(context, source, destination, sourceAttributes, &copiedBytes);
        if (FAILED(copyHr))
        {
            if (destinationAttributes == INVALID_FILE_ATTRIBUTES)
            {
                static_cast<void>(TryRollbackCopiedDestination(destination.extended));
            }
            return copyHr;
        }

        if (sourceIsDirectory)
        {
            if (! RemoveDirectoryW(source.extended.c_str()))
            {
                const HRESULT deleteHr = HRESULT_FROM_WIN32(GetLastError());
                return TryRollbackCopiedDestination(destination.extended) ? deleteHr : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }
        }
        else
        {
            DWORD newAttributes = sourceAttributes;
            if ((newAttributes & FILE_ATTRIBUTE_READONLY) != 0)
            {
                if (! context.allowReplaceReadonly)
                {
                    return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
                }

                newAttributes &= ~FILE_ATTRIBUTE_READONLY;
                if (! SetFileAttributesW(source.extended.c_str(), newAttributes))
                {
                    return HRESULT_FROM_WIN32(GetLastError());
                }
            }

            if (! DeleteFileW(source.extended.c_str()))
            {
                const HRESULT deleteHr = HRESULT_FROM_WIN32(GetLastError());
                return TryRollbackCopiedDestination(destination.extended) ? deleteHr : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }
        }

        return S_OK;
    }

    const DWORD renameFlags = moveFlags;

    CopyProgressContext progress{};
    progress.context = &context;
    if (! context.parallel)
    {
        progress.itemBaseBytes = context.completedBytes;
        progress.startTick     = GetTickCount64();
        progress.throttleWindowSamples.emplace_back(progress.startTick, 0);
    }

#ifdef _DEBUG
    const bool forceCopyFallback = allowCopy && ShouldForceMoveCopyFallbackForSelfTest();
#else
    constexpr bool forceCopyFallback = false;
#endif

    if (! forceCopyFallback && MoveFileWithProgressW(source.extended.c_str(), destination.extended.c_str(), CopyProgressRoutine, &progress, renameFlags))
    {
        const uint64_t finalTotalBytes     = (std::max)(progress.lastItemTotalBytes, progress.lastItemBytesTransferred);
        const uint64_t finalCompletedBytes = finalTotalBytes;

        if (context.parallel)
        {
            if (finalCompletedBytes > progress.lastItemBytesTransferred)
            {
                context.parallel->completedBytes.fetch_add(finalCompletedBytes - progress.lastItemBytesTransferred, std::memory_order_acq_rel);
                progress.lastItemBytesTransferred = finalCompletedBytes;
            }
        }
        else
        {
            context.completedBytes = progress.itemBaseBytes + finalCompletedBytes;
        }

        const HRESULT progressHr = ReportProgressForced(context, finalTotalBytes, finalCompletedBytes);
        if (progressHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || progressHr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (FAILED(progressHr))
        {
            EmitSequentialThrottleSummary(context, progress, finalTotalBytes, progressHr);
            return progressHr;
        }
        EmitSequentialThrottleSummary(context, progress, finalTotalBytes, S_OK);
        return S_OK;
    }

    DWORD error = forceCopyFallback ? ERROR_NOT_SAME_DEVICE : GetLastError();
#ifdef _DEBUG
    if (forceCopyFallback)
    {
        Debug::Perf::EmitCounter(L"FileOps.Move.DebugForceCopyFallback");
    }
#endif
    if (error == ERROR_REQUEST_ABORTED || error == ERROR_CANCELLED)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (caseOnlyRename && (error == ERROR_ACCESS_DENIED || error == ERROR_ALREADY_EXISTS))
    {
        const HRESULT caseHr = RenameCaseOnlyWithTemp(context, source.extended, destination.extended, renameFlags);
        if (SUCCEEDED(caseHr))
        {
            return S_OK;
        }
        return caseHr;
    }

    if (! allowCopy || error != ERROR_NOT_SAME_DEVICE)
    {
        return HRESULT_FROM_WIN32(error);
    }

    // Cross-volume move fallback: copy with reparse policy applied, then best-effort delete.
    if (sourceIsDirectory && ! context.recursive)
    {
        return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
    }

    uint64_t bytesCopied = 0;
    const HRESULT copyHr = CopyPathInternalWithDirectoryParallelism(context, source, destination, flags, reparsePointPolicy, maxCopyConcurrency, &bytesCopied);
    if (FAILED(copyHr))
    {
        if (destinationAttributes == INVALID_FILE_ATTRIBUTES)
        {
            static_cast<void>(TryRollbackCopiedDestination(destination.extended));
        }
        return copyHr;
    }

    struct DeletePhaseCallback final : IFileSystemCallback
    {
        IFileSystemCallback* inner = nullptr;

        explicit DeletePhaseCallback(IFileSystemCallback* callback) noexcept : inner(callback)
        {
        }

        HRESULT STDMETHODCALLTYPE FileSystemProgress(FileSystemOperation /*operationType*/,
                                                     unsigned long /*totalItems*/,
                                                     unsigned long /*completedItems*/,
                                                     uint64_t /*totalBytes*/,
                                                     uint64_t /*completedBytes*/,
                                                     const wchar_t* /*currentSourcePath*/,
                                                     const wchar_t* /*currentDestinationPath*/,
                                                     uint64_t /*currentItemTotalBytes*/,
                                                     uint64_t /*currentItemCompletedBytes*/,
                                                     FileSystemOptions* /*options*/,
                                                     uint64_t /*progressStreamId*/,
                                                     void* /*cookie*/) noexcept override
        {
            // Suppress delete-phase progress reporting for move operations (move progress reflects copy bytes).
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(FileSystemOperation /*operationType*/,
                                                          unsigned long /*itemIndex*/,
                                                          const wchar_t* /*sourcePath*/,
                                                          const wchar_t* /*destinationPath*/,
                                                          HRESULT /*status*/,
                                                          FileSystemOptions* /*options*/,
                                                          void* /*cookie*/) noexcept override
        {
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, void* cookie) noexcept override
        {
            if (! inner)
            {
                if (pCancel)
                {
                    *pCancel = FALSE;
                }
                return S_OK;
            }
            return inner->FileSystemShouldCancel(pCancel, cookie);
        }

        HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation operationType,
                                                  const wchar_t* sourcePath,
                                                  const wchar_t* destinationPath,
                                                  HRESULT status,
                                                  FileSystemIssueAction* action,
                                                  FileSystemOptions* options,
                                                  void* cookie) noexcept override
        {
            if (! inner)
            {
                if (action)
                {
                    *action = FileSystemIssueAction::Cancel;
                }
                return S_OK;
            }
            return inner->FileSystemIssue(operationType, sourcePath, destinationPath, status, action, options, cookie);
        }
    };

    DeletePhaseCallback deleteCallback(context.callback);

    OperationContext deleteContext{};
    deleteContext.type                   = FILESYSTEM_DELETE;
    deleteContext.callback               = &deleteCallback;
    deleteContext.callbackCookie         = context.callbackCookie;
    deleteContext.options                = nullptr;
    deleteContext.totalItems             = 0;
    deleteContext.completedItems         = 0;
    deleteContext.totalBytes             = 0;
    deleteContext.completedBytes         = 0;
    deleteContext.continueOnError        = false;
    deleteContext.allowOverwrite         = false;
    deleteContext.allowReplaceReadonly   = context.allowReplaceReadonly;
    deleteContext.recursive              = true;
    deleteContext.useRecycleBin          = false;
    deleteContext.parallel               = nullptr;
    deleteContext.lastProgressReportTick = 0;

    const HRESULT deleteHr = DeletePathInternal(deleteContext, source);
    if (FAILED(deleteHr))
    {
        return TryRollbackCopiedDestination(destination.extended) ? deleteHr : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

class RecycleBinDeleteProgressSink final : public IFileOperationProgressSink
{
public:
    explicit RecycleBinDeleteProgressSink(OperationContext& context) noexcept : _context(&context)
    {
    }
    RecycleBinDeleteProgressSink(const RecycleBinDeleteProgressSink&)            = delete;
    RecycleBinDeleteProgressSink(RecycleBinDeleteProgressSink&&)                 = delete;
    RecycleBinDeleteProgressSink& operator=(const RecycleBinDeleteProgressSink&) = delete;
    RecycleBinDeleteProgressSink& operator=(RecycleBinDeleteProgressSink&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileOperationProgressSink))
        {
            *ppvObject = static_cast<IFileOperationProgressSink*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE StartOperations() noexcept override
    {
        if (_context != nullptr)
        {
            _baseCompletedItems = _context->parallel ? _context->parallel->completedItems.load(std::memory_order_acquire) : _context->completedItems;
            _baseTotalItems     = _context->totalItems;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FinishOperations([[maybe_unused]] HRESULT hrResult) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreRenameItem([[maybe_unused]] DWORD flags, [[maybe_unused]] IShellItem* item, [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostRenameItem([[maybe_unused]] DWORD flags,
                                             [[maybe_unused]] IShellItem* item,
                                             [[maybe_unused]] LPCWSTR newName,
                                             [[maybe_unused]] HRESULT hrRename,
                                             [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreMoveItem([[maybe_unused]] DWORD flags,
                                          [[maybe_unused]] IShellItem* item,
                                          [[maybe_unused]] IShellItem* destinationFolder,
                                          [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostMoveItem([[maybe_unused]] DWORD flags,
                                           [[maybe_unused]] IShellItem* item,
                                           [[maybe_unused]] IShellItem* destinationFolder,
                                           [[maybe_unused]] LPCWSTR newName,
                                           [[maybe_unused]] HRESULT hrMove,
                                           [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreCopyItem([[maybe_unused]] DWORD flags,
                                          [[maybe_unused]] IShellItem* item,
                                          [[maybe_unused]] IShellItem* destinationFolder,
                                          [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostCopyItem([[maybe_unused]] DWORD flags,
                                           [[maybe_unused]] IShellItem* item,
                                           [[maybe_unused]] IShellItem* destinationFolder,
                                           [[maybe_unused]] LPCWSTR newName,
                                           [[maybe_unused]] HRESULT hrCopy,
                                           [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreDeleteItem([[maybe_unused]] DWORD flags, [[maybe_unused]] IShellItem* item) noexcept override
    {
        const HRESULT hr = ReportItemPath(item, false);
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostDeleteItem([[maybe_unused]] DWORD flags,
                                             IShellItem* item,
                                             HRESULT hrDelete,
                                             [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        if (SUCCEEDED(hrDelete) && _context != nullptr)
        {
            if (! _workProgressAvailable)
            {
                AddCompletedItems(*_context, 1);
            }
            const HRESULT hr = ReportItemPath(item, false);
            if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
        }

        if (SUCCEEDED(hrDelete) || FAILED(_firstError))
        {
            return S_OK;
        }

        _firstError = hrDelete;
        if (item != nullptr)
        {
            wil::unique_cotaskmem_string path;
            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, path.put())) && path && path.get()[0] != L'\0')
            {
                _firstErrorPath.assign(path.get());
            }
            else
            {
                path.reset();
                if (SUCCEEDED(item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, path.put())) && path && path.get()[0] != L'\0')
                {
                    _firstErrorPath.assign(path.get());
                }
            }
        }

        static_cast<void>(ReportItemPath(item, true));
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreNewItem([[maybe_unused]] DWORD flags,
                                         [[maybe_unused]] IShellItem* destinationFolder,
                                         [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostNewItem([[maybe_unused]] DWORD flags,
                                          [[maybe_unused]] IShellItem* destinationFolder,
                                          [[maybe_unused]] LPCWSTR newName,
                                          [[maybe_unused]] LPCWSTR templateName,
                                          [[maybe_unused]] DWORD fileAttributes,
                                          [[maybe_unused]] HRESULT hrNew,
                                          [[maybe_unused]] IShellItem* newItem) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE UpdateProgress(UINT workTotal, UINT workSoFar) noexcept override
    {
        if (_context == nullptr)
        {
            return S_OK;
        }

        if (workTotal > 0 || workSoFar > 0)
        {
            _workProgressAvailable = true;
        }

        if (_workProgressAvailable)
        {
            constexpr uint64_t maxUlong       = static_cast<uint64_t>(std::numeric_limits<unsigned long>::max());
            const uint64_t desiredTotal64     = static_cast<uint64_t>(_baseCompletedItems) + static_cast<uint64_t>(workTotal);
            const uint64_t desiredCompleted64 = static_cast<uint64_t>(_baseCompletedItems) + static_cast<uint64_t>(workSoFar);

            const unsigned long desiredTotal     = static_cast<unsigned long>(std::min(desiredTotal64, maxUlong));
            const unsigned long desiredCompleted = static_cast<unsigned long>(std::min(desiredCompleted64, maxUlong));

            _context->totalItems = std::max(_context->totalItems, desiredTotal);
            if (_context->parallel)
            {
                unsigned long current = _context->parallel->completedItems.load(std::memory_order_acquire);
                while (current < desiredCompleted &&
                       ! _context->parallel->completedItems.compare_exchange_weak(current, desiredCompleted, std::memory_order_acq_rel))
                {
                }
            }
            else
            {
                _context->completedItems = std::max(_context->completedItems, desiredCompleted);
            }
        }

        const HRESULT hr = ReportProgress(*_context, 0, 0);
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || hr == E_ABORT)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ResetTimer() noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PauseTimer() noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ResumeTimer() noexcept override
    {
        return S_OK;
    }

    [[nodiscard]] HRESULT GetFirstError() const noexcept
    {
        return _firstError;
    }

    [[nodiscard]] const std::wstring& GetFirstErrorPath() const noexcept
    {
        return _firstErrorPath;
    }

private:
    [[nodiscard]] HRESULT ReportItemPath(IShellItem* item, bool force) noexcept
    {
        if (item == nullptr || _context == nullptr)
        {
            return S_OK;
        }

        wil::unique_cotaskmem_string path;
        if (FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, path.put())) || ! path || path.get()[0] == L'\0')
        {
            path.reset();
            static_cast<void>(item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, path.put()));
        }

        if (! path || path.get()[0] == L'\0')
        {
            return S_OK;
        }

        const HRESULT hrPaths = SetProgressPaths(*_context, path.get(), nullptr);
        if (FAILED(hrPaths))
        {
            return hrPaths;
        }

        return force ? ReportProgressForced(*_context, 0, 0) : ReportProgress(*_context, 0, 0);
    }

    ~RecycleBinDeleteProgressSink() = default;

    std::atomic_ulong _refCount{1};
    OperationContext* _context        = nullptr;
    unsigned long _baseCompletedItems = 0;
    unsigned long _baseTotalItems     = 0;
    bool _workProgressAvailable       = false;
    HRESULT _firstError               = S_OK;
    std::wstring _firstErrorPath;
};

struct RecycleBinBatchEntry
{
    unsigned long itemIndex = 0;
    PathInfo path{};
    HRESULT result = S_OK;
    bool observed  = false;
};

template <typename Operation> HRESULT RunShellOperationOnDedicatedStaThread(Operation operation) noexcept
{
    HRESULT result = E_FAIL;
    try
    {
        std::jthread worker([&result, operation = std::move(operation)]() mutable noexcept
        {
            wil::unique_hmodule modulePin = AcquireModuleReferenceFromAddress(&kFileSystemModuleAnchor);
            if (! modulePin)
            {
                DWORD lastError = GetLastError();
                if (lastError == ERROR_SUCCESS)
                {
                    lastError = ERROR_GEN_FAILURE;
                }
                result = HRESULT_FROM_WIN32(lastError);
                return;
            }

            const HRESULT coInitHr   = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
            const bool coInitialized = SUCCEEDED(coInitHr) || coInitHr == S_FALSE;
            auto coUninit            = wil::scope_exit([&]() noexcept
            {
                if (coInitialized)
                {
                    CoUninitialize();
                }
            });
            if (FAILED(coInitHr) && coInitHr != RPC_E_CHANGED_MODE)
            {
                result = coInitHr;
                return;
            }

            result = operation();
        });
        worker.join();
    }
    catch (const std::system_error& ex)
    {
        // Thread creation is required to run shell recycle-bin operations on an STA thread from MTA worker callbacks.
        const int code = ex.code().value();
        return code > 0 ? HRESULT_FROM_WIN32(static_cast<unsigned long>(code)) : E_FAIL;
    }

    return result;
}

[[nodiscard]] bool PathsEqualInsensitive(const std::wstring& lhs, const std::wstring& rhs) noexcept
{
    return _wcsicmp(lhs.c_str(), rhs.c_str()) == 0;
}

[[nodiscard]] HRESULT GetShellItemDisplayPath(IShellItem* item, std::wstring& path) noexcept
{
    path.clear();
    if (item == nullptr)
    {
        return E_POINTER;
    }

    wil::unique_cotaskmem_string rawPath;
    HRESULT hr = item->GetDisplayName(SIGDN_FILESYSPATH, rawPath.put());
    if (FAILED(hr) || ! rawPath || rawPath.get()[0] == L'\0')
    {
        rawPath.reset();
        hr = item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, rawPath.put());
    }

    if (FAILED(hr))
    {
        return hr;
    }

    if (! rawPath || rawPath.get()[0] == L'\0')
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    path.assign(rawPath.get());
    return S_OK;
}

class RecycleBinBatchProgressSink final : public IFileOperationProgressSink
{
public:
    RecycleBinBatchProgressSink(OperationContext& context, std::span<RecycleBinBatchEntry> items) noexcept : _context(&context), _items(items)
    {
    }

    RecycleBinBatchProgressSink(const RecycleBinBatchProgressSink&)            = delete;
    RecycleBinBatchProgressSink(RecycleBinBatchProgressSink&&)                 = delete;
    RecycleBinBatchProgressSink& operator=(const RecycleBinBatchProgressSink&) = delete;
    RecycleBinBatchProgressSink& operator=(RecycleBinBatchProgressSink&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileOperationProgressSink))
        {
            *ppvObject = static_cast<IFileOperationProgressSink*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE StartOperations() noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FinishOperations([[maybe_unused]] HRESULT hrResult) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreRenameItem([[maybe_unused]] DWORD flags, [[maybe_unused]] IShellItem* item, [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostRenameItem([[maybe_unused]] DWORD flags,
                                             [[maybe_unused]] IShellItem* item,
                                             [[maybe_unused]] LPCWSTR newName,
                                             [[maybe_unused]] HRESULT hrRename,
                                             [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreMoveItem([[maybe_unused]] DWORD flags,
                                          [[maybe_unused]] IShellItem* item,
                                          [[maybe_unused]] IShellItem* destinationFolder,
                                          [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostMoveItem([[maybe_unused]] DWORD flags,
                                           [[maybe_unused]] IShellItem* item,
                                           [[maybe_unused]] IShellItem* destinationFolder,
                                           [[maybe_unused]] LPCWSTR newName,
                                           [[maybe_unused]] HRESULT hrMove,
                                           [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreCopyItem([[maybe_unused]] DWORD flags,
                                          [[maybe_unused]] IShellItem* item,
                                          [[maybe_unused]] IShellItem* destinationFolder,
                                          [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostCopyItem([[maybe_unused]] DWORD flags,
                                           [[maybe_unused]] IShellItem* item,
                                           [[maybe_unused]] IShellItem* destinationFolder,
                                           [[maybe_unused]] LPCWSTR newName,
                                           [[maybe_unused]] HRESULT hrCopy,
                                           [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreDeleteItem([[maybe_unused]] DWORD flags, IShellItem* item) noexcept override
    {
        const HRESULT hr = ReportItemPath(item, false);
        if (FAILED(hr))
        {
            _callbackFailure = hr;
            return IsCancellationHr(hr) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : hr;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostDeleteItem([[maybe_unused]] DWORD flags,
                                             IShellItem* item,
                                             HRESULT hrDelete,
                                             [[maybe_unused]] IShellItem* newlyCreated) noexcept override
    {
        RecycleBinBatchEntry* entry = FindEntry(item);
        if (entry != nullptr)
        {
            entry->result   = hrDelete;
            entry->observed = true;
        }

        const HRESULT hr = ReportItemPath(item, FAILED(hrDelete));
        if (FAILED(hr))
        {
            _callbackFailure = hr;
            return IsCancellationHr(hr) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : hr;
        }

        if (FAILED(hrDelete) && _context != nullptr && ! _context->continueOnError)
        {
            return E_ABORT;
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PreNewItem([[maybe_unused]] DWORD flags,
                                         [[maybe_unused]] IShellItem* destinationFolder,
                                         [[maybe_unused]] LPCWSTR newName) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostNewItem([[maybe_unused]] DWORD flags,
                                          [[maybe_unused]] IShellItem* destinationFolder,
                                          [[maybe_unused]] LPCWSTR newName,
                                          [[maybe_unused]] LPCWSTR templateName,
                                          [[maybe_unused]] DWORD fileAttributes,
                                          [[maybe_unused]] HRESULT hrNew,
                                          [[maybe_unused]] IShellItem* newItem) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE UpdateProgress([[maybe_unused]] UINT workTotal, [[maybe_unused]] UINT workSoFar) noexcept override
    {
        if (_context == nullptr)
        {
            return S_OK;
        }

        const HRESULT hr = ReportProgress(*_context, 0, 0);
        if (FAILED(hr))
        {
            _callbackFailure = hr;
            return IsCancellationHr(hr) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : hr;
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ResetTimer() noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PauseTimer() noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE ResumeTimer() noexcept override
    {
        return S_OK;
    }

    [[nodiscard]] HRESULT GetCallbackFailure() const noexcept
    {
        return _callbackFailure;
    }

private:
    [[nodiscard]] RecycleBinBatchEntry* FindEntry(IShellItem* item) noexcept
    {
        std::wstring itemPath;
        if (FAILED(GetShellItemDisplayPath(item, itemPath)))
        {
            return nullptr;
        }

        const PathInfo itemInfo = MakePathInfo(itemPath);
        for (RecycleBinBatchEntry& entry : _items)
        {
            if (entry.observed)
            {
                continue;
            }

            if (PathsEqualInsensitive(entry.path.display, itemPath) || PathsEqualInsensitive(entry.path.extended, itemInfo.extended))
            {
                return &entry;
            }
        }

        return nullptr;
    }

    [[nodiscard]] HRESULT ReportItemPath(IShellItem* item, bool force) noexcept
    {
        if (_context == nullptr)
        {
            return S_OK;
        }

        std::wstring itemPath;
        const HRESULT hr = GetShellItemDisplayPath(item, itemPath);
        if (FAILED(hr))
        {
            return S_OK;
        }

        const HRESULT hrPaths = SetProgressPaths(*_context, itemPath.c_str(), nullptr);
        if (FAILED(hrPaths))
        {
            return hrPaths;
        }

        return force ? ReportProgressForced(*_context, 0, 0) : ReportProgress(*_context, 0, 0);
    }

    ~RecycleBinBatchProgressSink() = default;

    std::atomic_ulong _refCount{1};
    OperationContext* _context = nullptr;
    std::span<RecycleBinBatchEntry> _items;
    HRESULT _callbackFailure = S_OK;
};

HRESULT DeleteToRecycleBinCore(OperationContext& context, const PathInfo& path) noexcept
{
    if (path.display.empty())
    {
        return E_INVALIDARG;
    }

    wil::com_ptr<IFileOperation> fileOperation;
    HRESULT hr = CoCreateInstance(CLSID_FileOperation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(fileOperation.put()));
    if (FAILED(hr) || ! fileOperation)
    {
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }

    constexpr DWORD kOperationFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT | FOFX_EARLYFAILURE | FOFX_RECYCLEONDELETE;
    hr                              = fileOperation->SetOperationFlags(kOperationFlags);
    if (FAILED(hr))
    {
        return hr;
    }

    wil::com_ptr<IShellItem> item;
    hr = SHCreateItemFromParsingName(path.display.c_str(), nullptr, IID_PPV_ARGS(item.put()));
    if (FAILED(hr) || ! item)
    {
        return FAILED(hr) ? hr : E_INVALIDARG;
    }

    wil::com_ptr<IFileOperationProgressSink> progressSink;
    auto* progressSinkImpl = new (std::nothrow) RecycleBinDeleteProgressSink(context);
    if (! progressSinkImpl)
    {
        return E_OUTOFMEMORY;
    }
    progressSink.attach(progressSinkImpl);

    DWORD adviseCookie = 0;
    hr                 = fileOperation->Advise(progressSink.get(), &adviseCookie);
    if (FAILED(hr))
    {
        return hr;
    }
    auto unadvise = wil::scope_exit([&]() noexcept
    {
        if (adviseCookie != 0)
        {
            static_cast<void>(fileOperation->Unadvise(adviseCookie));
        }
    });

    hr = fileOperation->DeleteItem(item.get(), nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = fileOperation->PerformOperations();
    if (FAILED(hr))
    {
        const HRESULT itemError = progressSinkImpl->GetFirstError();
        if (FAILED(itemError))
        {
            const std::wstring& itemPath          = progressSinkImpl->GetFirstErrorPath();
            const std::wstring_view effectivePath = itemPath.empty() ? std::wstring_view(path.display) : std::wstring_view(itemPath);
            Debug::Warning(L"FileSystem: Recycle Bin delete failed for '{}' (hr={:#x})", effectivePath, static_cast<unsigned long>(itemError));
            return itemError;
        }

        return hr;
    }

    BOOL anyAborted = FALSE;
    hr              = fileOperation->GetAnyOperationsAborted(&anyAborted);
    if (FAILED(hr))
    {
        const HRESULT itemError = progressSinkImpl->GetFirstError();
        if (FAILED(itemError))
        {
            const std::wstring& itemPath          = progressSinkImpl->GetFirstErrorPath();
            const std::wstring_view effectivePath = itemPath.empty() ? std::wstring_view(path.display) : std::wstring_view(itemPath);
            Debug::Warning(L"FileSystem: Recycle Bin delete failed for '{}' (hr={:#x})", effectivePath, static_cast<unsigned long>(itemError));
            return itemError;
        }
        return hr;
    }

    if (anyAborted == TRUE)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    const HRESULT itemError = progressSinkImpl->GetFirstError();
    if (FAILED(itemError))
    {
        const std::wstring& itemPath          = progressSinkImpl->GetFirstErrorPath();
        const std::wstring_view effectivePath = itemPath.empty() ? std::wstring_view(path.display) : std::wstring_view(itemPath);
        Debug::Warning(L"FileSystem: Recycle Bin delete failed for '{}' (hr={:#x})", effectivePath, static_cast<unsigned long>(itemError));
        return itemError;
    }

    static_cast<void>(ReportProgressForced(context, 0, 0));
    return S_OK;
}

HRESULT DeleteToRecycleBin(OperationContext& context, const PathInfo& path) noexcept
{
    if (path.display.empty())
    {
        return E_INVALIDARG;
    }

    // IFileOperation is an STA-oriented shell API. File operation tasks can already be MTA,
    // so switch to a dedicated STA thread rather than running recycle-bin work in-place.
    const HRESULT coInitHr   = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    const bool coInitialized = SUCCEEDED(coInitHr) || coInitHr == S_FALSE;
    auto coUninit            = wil::scope_exit([&]() noexcept
    {
        if (coInitialized)
        {
            CoUninitialize();
        }
    });
    if (coInitHr == RPC_E_CHANGED_MODE)
    {
        return RunShellOperationOnDedicatedStaThread([&]() noexcept { return DeleteToRecycleBinCore(context, path); });
    }
    if (FAILED(coInitHr))
    {
        return coInitHr;
    }

    return DeleteToRecycleBinCore(context, path);
}

HRESULT DeleteToRecycleBinBatchedCore(OperationContext& context, std::span<RecycleBinBatchEntry> items, bool* batchStarted) noexcept
{
    if (batchStarted != nullptr)
    {
        *batchStarted = false;
    }

    if (items.size() < 2)
    {
        return E_INVALIDARG;
    }

    Debug::Perf::Scope batchPerf(L"FileOps.RecycleBin.BatchDeleteUs");
    batchPerf.SetValue0(items.size());

    wil::com_ptr<IFileOperation> fileOperation;
    HRESULT hr = CoCreateInstance(CLSID_FileOperation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(fileOperation.put()));
    if (FAILED(hr) || ! fileOperation)
    {
        hr = FAILED(hr) ? hr : E_NOINTERFACE;
        batchPerf.SetHr(hr);
        return hr;
    }

    constexpr DWORD kOperationFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT | FOFX_EARLYFAILURE | FOFX_RECYCLEONDELETE;
    hr                              = fileOperation->SetOperationFlags(kOperationFlags);
    if (FAILED(hr))
    {
        batchPerf.SetHr(hr);
        return hr;
    }

    wil::com_ptr<IFileOperationProgressSink> progressSink;
    auto* progressSinkImpl = new (std::nothrow) RecycleBinBatchProgressSink(context, items);
    if (progressSinkImpl == nullptr)
    {
        batchPerf.SetHr(E_OUTOFMEMORY);
        return E_OUTOFMEMORY;
    }
    progressSink.attach(progressSinkImpl);

    DWORD adviseCookie = 0;
    hr                 = fileOperation->Advise(progressSink.get(), &adviseCookie);
    if (FAILED(hr))
    {
        batchPerf.SetHr(hr);
        return hr;
    }
    auto unadvise = wil::scope_exit([&]() noexcept
    {
        if (adviseCookie != 0)
        {
            static_cast<void>(fileOperation->Unadvise(adviseCookie));
        }
    });

    wil::com_ptr<IShellItemArray> itemArray;
    {
        Debug::Perf::Scope buildPerf(L"FileOps.RecycleBin.BatchBuildUs");
        buildPerf.SetValue0(items.size());

        std::vector<PIDLIST_ABSOLUTE> pidlStorage;
        std::vector<PCIDLIST_ABSOLUTE> pidls;
        pidlStorage.reserve(items.size());
        pidls.reserve(items.size());
        auto freePidls = wil::scope_exit([&]() noexcept
        {
            for (PIDLIST_ABSOLUTE pidl : pidlStorage)
            {
                ::CoTaskMemFree(pidl);
            }
        });

        for (const RecycleBinBatchEntry& entry : items)
        {
            PIDLIST_ABSOLUTE pidl = ILCreateFromPathW(entry.path.display.c_str());
            if (pidl == nullptr)
            {
                hr = HRESULT_FROM_WIN32(GetLastError());
                if (FAILED(hr))
                {
                    buildPerf.SetHr(hr);
                    batchPerf.SetHr(hr);
                    return hr;
                }

                buildPerf.SetHr(E_INVALIDARG);
                batchPerf.SetHr(E_INVALIDARG);
                return E_INVALIDARG;
            }

            pidlStorage.push_back(pidl);
            pidls.push_back(pidlStorage.back());
        }

        hr = SHCreateShellItemArrayFromIDLists(static_cast<UINT>(pidls.size()), pidls.data(), itemArray.put());
        buildPerf.SetHr(hr);
        if (FAILED(hr) || ! itemArray)
        {
            hr = FAILED(hr) ? hr : E_NOINTERFACE;
            batchPerf.SetHr(hr);
            return hr;
        }
    }

    hr = fileOperation->DeleteItems(itemArray.get());
    if (FAILED(hr))
    {
        batchPerf.SetHr(hr);
        return hr;
    }

    if (batchStarted != nullptr)
    {
        *batchStarted = true;
    }

    {
        Debug::Perf::Scope performPerf(L"FileOps.RecycleBin.PerformOperationsUs");
        performPerf.SetValue0(items.size());
        hr = fileOperation->PerformOperations();
        performPerf.SetHr(hr);
    }

    const HRESULT callbackHr = progressSinkImpl->GetCallbackFailure();
    if (FAILED(callbackHr))
    {
        batchPerf.SetHr(callbackHr);
        return callbackHr;
    }

    BOOL anyAborted       = FALSE;
    const HRESULT abortHr = fileOperation->GetAnyOperationsAborted(&anyAborted);
    if (FAILED(abortHr))
    {
        batchPerf.SetHr(abortHr);
        return abortHr;
    }

    const uint64_t observedCount =
        static_cast<uint64_t>(std::ranges::count_if(items, [](const RecycleBinBatchEntry& entry) noexcept { return entry.observed; }));
    const uint64_t failedCount = static_cast<uint64_t>(std::ranges::count_if(
        items, [](const RecycleBinBatchEntry& entry) noexcept { return entry.observed && FAILED(entry.result) && ! IsCancellationHr(entry.result); }));
    Debug::Perf::EmitValue(L"FileOps.RecycleBin.BatchObservedItems", observedCount);
    Debug::Perf::EmitValue(L"FileOps.RecycleBin.BatchFailedItems", failedCount);

    const HRESULT result = (anyAborted == TRUE) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : hr;
    batchPerf.SetValue1(failedCount);
    batchPerf.SetHr(result);
    return result;
}

HRESULT DeleteToRecycleBinBatched(OperationContext& context, std::span<RecycleBinBatchEntry> items, bool* batchStarted) noexcept
{
    if (batchStarted != nullptr)
    {
        *batchStarted = false;
    }

    if (items.size() < 2)
    {
        return E_INVALIDARG;
    }

    const HRESULT coInitHr   = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    const bool coInitialized = SUCCEEDED(coInitHr) || coInitHr == S_FALSE;
    auto coUninit            = wil::scope_exit([&]() noexcept
    {
        if (coInitialized)
        {
            CoUninitialize();
        }
    });
    if (coInitHr == RPC_E_CHANGED_MODE)
    {
        return RunShellOperationOnDedicatedStaThread([&]() noexcept { return DeleteToRecycleBinBatchedCore(context, items, batchStarted); });
    }
    if (FAILED(coInitHr))
    {
        return coInitHr;
    }

    return DeleteToRecycleBinBatchedCore(context, items, batchStarted);
}

HRESULT DeleteDirectoryRecursive(OperationContext& context, const PathInfo& path) noexcept;
HRESULT DeleteDirectoryRecursiveSequential(OperationContext& context, const PathInfo& path) noexcept;
HRESULT DeleteDirectoryRecursiveParallel(OperationContext& context, const PathInfo& path, unsigned int requestedConcurrency) noexcept;

HRESULT DeletePathInternal(OperationContext& context, const PathInfo& path) noexcept
{
    HRESULT hr = SetProgressPaths(context, path.display.c_str(), nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReportProgress(context, 0, 0);
    if (FAILED(hr))
    {
        return hr;
    }

    if (context.useRecycleBin)
    {
        return DeleteToRecycleBin(context, path);
    }

    DWORD attributes = GetFileAttributesW(path.extended.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        // Never traverse directory reparse points during delete recursion (junction/symlink safety).
        if (IsReparsePoint(attributes))
        {
            if (! RemoveDirectoryW(path.extended.c_str()))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            AddCompletedItems(context, 1);
            return S_OK;
        }

        if (! context.recursive)
        {
            if (! RemoveDirectoryW(path.extended.c_str()))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            AddCompletedItems(context, 1);
            return S_OK;
        }

        return DeleteDirectoryRecursive(context, path);
    }

    uint64_t fileBytes = 0;
    static_cast<void>(GetFileSizeBytes(path.extended, &fileBytes)); // Best-effort only.

    if ((attributes & FILE_ATTRIBUTE_READONLY) != 0)
    {
        bool restoreAttributes      = false;
        auto restoreAttributesScope = wil::scope_exit([&]() noexcept
        {
            if (restoreAttributes)
            {
                static_cast<void>(SetFileAttributesW(path.extended.c_str(), attributes));
            }
        });

        if (! context.allowReplaceReadonly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        const DWORD newAttributes = attributes & ~FILE_ATTRIBUTE_READONLY;
        if (! SetFileAttributesW(path.extended.c_str(), newAttributes))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        restoreAttributes = true;

        if (! DeleteFileW(path.extended.c_str()))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        restoreAttributes = false;
        AddCompletedItems(context, 1);
        return S_OK;
    }

    if (! DeleteFileW(path.extended.c_str()))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    AddCompletedItems(context, 1);
    AddCompletedBytes(context, fileBytes);

    return S_OK;
}

struct DeleteFlattenFrame final
{
    PathInfo directory;
    bool enumerated = false;
};

[[nodiscard]] HRESULT FlattenDeleteDirectoryTree(OperationContext& context,
                                                 const PathInfo& root,
                                                 std::vector<PathInfo>& outFiles,
                                                 std::vector<PathInfo>& outDirectoriesPostOrder) noexcept
{
    constexpr size_t kMaxWorkItems = 200000u;

    std::vector<DeleteFlattenFrame> stack;
    stack.reserve(256);
    stack.push_back(DeleteFlattenFrame{root, false});

    while (! stack.empty())
    {
        if (! stack.back().enumerated)
        {
            stack.back().enumerated = true;
            const size_t frameIndex = stack.size() - 1;

            HRESULT hr = CheckCancel(context);
            if (FAILED(hr))
            {
                return hr;
            }

            // NOTE: We may push child frames onto the same vector while enumerating this directory.
            // Avoid holding references into `stack` when calling `push_back` (realloc can invalidate).
            std::wstring searchPattern = AppendPath(stack[frameIndex].directory.extended, L"*");
            WIN32_FIND_DATAW data{};
            wil::unique_hfind findHandle(
                FindFirstFileExW(searchPattern.c_str(), FindExInfoBasic, &data, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
            if (! findHandle)
            {
                const DWORD error = GetLastError();
                if (error == ERROR_FILE_NOT_FOUND)
                {
                    continue;
                }
                return HRESULT_FROM_WIN32(error);
            }

            do
            {
                if (IsDotOrDotDot(data.cFileName))
                {
                    continue;
                }

                PathInfo child{};
                child.display  = AppendPath(stack[frameIndex].directory.display, data.cFileName);
                child.extended = AppendPath(stack[frameIndex].directory.extended, data.cFileName);

                const bool isDirectory = (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (isDirectory)
                {
                    if (IsReparsePoint(data.dwFileAttributes))
                    {
                        outDirectoriesPostOrder.push_back(std::move(child));
                    }
                    else
                    {
                        stack.push_back(DeleteFlattenFrame{std::move(child), false});
                    }
                }
                else
                {
                    outFiles.push_back(std::move(child));
                }

                if (outFiles.size() + outDirectoriesPostOrder.size() >= kMaxWorkItems)
                {
                    return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
                }
            } while (FindNextFileW(findHandle.get(), &data));

            const DWORD error = GetLastError();
            if (error != ERROR_NO_MORE_FILES)
            {
                return HRESULT_FROM_WIN32(error);
            }
        }
        else
        {
            outDirectoriesPostOrder.push_back(std::move(stack.back().directory));
            stack.pop_back();
        }
    }

    return S_OK;
}

HRESULT DeleteDirectoryRecursiveParallel(OperationContext& rootContext, const PathInfo& path, unsigned int requestedConcurrency) noexcept
{
    constexpr unsigned int kMaxWorkers = 8u;
    const unsigned int concurrency     = std::clamp(requestedConcurrency, 1u, kMaxWorkers);

    std::vector<PathInfo> files;
    std::vector<PathInfo> directoriesPostOrder;
    HRESULT hr = FlattenDeleteDirectoryTree(rootContext, path, files, directoriesPostOrder);
    if (FAILED(hr))
    {
        return hr;
    }

    FileSystemOptions* sharedOptionsState = rootContext.options;

    ParallelOperationState parallel{};
    parallel.startTick = GetTickCount64();
    parallel.bandwidthLimitBytesPerSecond.store(sharedOptionsState ? sharedOptionsState->bandwidthLimitBytesPerSecond : 0ull, std::memory_order_release);

    const FileSystemFlags workerFlags =
        static_cast<FileSystemFlags>((rootContext.continueOnError ? FILESYSTEM_FLAG_CONTINUE_ON_ERROR : FILESYSTEM_FLAG_NONE) |
                                     (rootContext.allowReplaceReadonly ? FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY : FILESYSTEM_FLAG_NONE));

    const auto initializeWorkerContext = [&](OperationContext& context, uint64_t progressStreamId) noexcept
    {
        InitializeOperationContext(
            context, FILESYSTEM_DELETE, workerFlags, sharedOptionsState, rootContext.callback, rootContext.callbackCookie, 0, rootContext.reparsePointPolicy);
        context.options                 = sharedOptionsState;
        context.parallel                = &parallel;
        context.totalBytes              = 0; // let the host provide totals via pre-calc
        context.progressStreamId        = progressStreamId;
        context.recursive               = false;
        context.useRecycleBin           = false;
        context.deleteConcurrencyBudget = 1;
    };

    const auto processFile = [&](size_t index, uint64_t schedulerStreamId) noexcept
    {
        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            return;
        }

        if (index >= files.size())
        {
            return;
        }

        OperationContext context{};
        initializeWorkerContext(context, schedulerStreamId);

        HRESULT itemHr = DeletePathInternal(context, files[index]);
        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                if (! stopOnError)
                {
                    parallel.cancelRequested.store(true, std::memory_order_release);
                }
                return;
            }

            parallel.hadFailure.store(true, std::memory_order_release);
            if (! context.continueOnError)
            {
                parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                HRESULT expected = S_OK;
                static_cast<void>(parallel.firstError.compare_exchange_strong(expected, itemHr, std::memory_order_acq_rel));
                return;
            }
        }
    };

    if (files.size() >= 2 && concurrency > 1u && GetSharedFileOpsJobScheduler().EnsureWorkersAvailable())
    {
        auto job = GetSharedFileOpsJobScheduler().StartJob(concurrency, files.size(), processFile);
        GetSharedFileOpsJobScheduler().WaitJob(job);
    }
    else
    {
        for (size_t i = 0; i < files.size(); ++i)
        {
            processFile(i, 0);
            if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
            {
                break;
            }
        }
    }

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
    {
        const HRESULT first = parallel.firstError.load(std::memory_order_acquire);
        return FAILED(first) ? first : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    OperationContext dirContext{};
    initializeWorkerContext(dirContext, 0);

    for (const PathInfo& directory : directoriesPostOrder)
    {
        hr = DeletePathInternal(dirContext, directory);
        if (FAILED(hr))
        {
            if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return hr;
            }

            parallel.hadFailure.store(true, std::memory_order_release);
            if (! dirContext.continueOnError)
            {
                return hr;
            }
        }

        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            break;
        }
    }

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
    {
        const HRESULT first = parallel.firstError.load(std::memory_order_acquire);
        return FAILED(first) ? first : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.hadFailure.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT DeleteDirectoryRecursive(OperationContext& context, const PathInfo& path) noexcept
{
    const unsigned int requestedConcurrency = std::clamp(context.deleteConcurrencyBudget, 1u, 8u);
    if (context.recursive && ! context.useRecycleBin && context.parallel == nullptr && requestedConcurrency > 1u &&
        GetSharedFileOpsJobScheduler().EnsureWorkersAvailable())
    {
        const HRESULT parallelHr = DeleteDirectoryRecursiveParallel(context, path, requestedConcurrency);
        if (parallelHr != HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY))
        {
            return parallelHr;
        }
    }

    return DeleteDirectoryRecursiveSequential(context, path);
}

HRESULT DeleteDirectoryRecursiveSequential(OperationContext& context, const PathInfo& path) noexcept
{
    std::wstring searchPattern = AppendPath(path.extended, L"*");
    WIN32_FIND_DATAW data{};
    wil::unique_hfind findHandle(FindFirstFileExW(searchPattern.c_str(), FindExInfoBasic, &data, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
    if (! findHandle)
    {
        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND)
        {
            return S_OK;
        }
        return HRESULT_FROM_WIN32(error);
    }

    bool hadFailure = false;

    do
    {
        if (IsDotOrDotDot(data.cFileName))
        {
            continue;
        }

        PathInfo child{};
        child.display  = AppendPath(path.display, data.cFileName);
        child.extended = AppendPath(path.extended, data.cFileName);

        HRESULT childHr = DeletePathInternal(context, child);
        if (FAILED(childHr))
        {
            if (childHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return childHr;
            }

            hadFailure = true;
            if (! context.continueOnError)
            {
                return childHr;
            }
        }

        HRESULT hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }
    } while (FindNextFileW(findHandle.get(), &data));

    const DWORD error = GetLastError();
    if (error != ERROR_NO_MORE_FILES)
    {
        return HRESULT_FROM_WIN32(error);
    }

    if (! RemoveDirectoryW(path.extended.c_str()))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    AddCompletedItems(context, 1);

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}
} // namespace

HRESULT STDMETHODCALLTYPE FileSystem::CopyItem(const wchar_t* sourcePath,
                                               const wchar_t* destinationPath,
                                               FileSystemFlags flags,
                                               const FileSystemOptions* options,
                                               IFileSystemCallback* callback,
                                               void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int copyMoveMaxConcurrency             = 1u;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_COPY, flags, options, callback, cookie, 1, reparsePointPolicy);

    const PathInfo source      = MakePathInfo(sourcePath);
    const PathInfo destination = MakePathInfo(destinationPath);

    HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        Debug::Warning(L"FileSystem: CopyItem failed to set paths for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(hr));
        return hr;
    }

    context.reparseRootSourcePath      = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(source.display)));
    context.reparseRootDestinationPath = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(destination.display)));

    uint64_t bytesCopied = 0;
    HRESULT itemHr       = S_OK;

    const unsigned int maxConcurrency = std::clamp(copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    itemHr = CopyPathInternalWithDirectoryParallelism(context, source, destination, flags, reparsePointPolicy, maxConcurrency, &bytesCopied);
    if (FAILED(itemHr))
    {
        Debug::Warning(L"FileSystem: CopyItem failed for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(itemHr));
    }

    hr = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    context.completedItems = 1;
    return itemHr;
}

HRESULT STDMETHODCALLTYPE FileSystem::MoveItem(const wchar_t* sourcePath,
                                               const wchar_t* destinationPath,
                                               FileSystemFlags flags,
                                               const FileSystemOptions* options,
                                               IFileSystemCallback* callback,
                                               void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int copyMoveMaxConcurrency             = 1;
    unsigned int deleteMaxConcurrency               = 1;
    unsigned int deleteRecycleBinMaxConcurrency     = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy             = _reparsePointPolicy;
        copyMoveMaxConcurrency         = _copyMoveMaxConcurrency;
        deleteMaxConcurrency           = _deleteMaxConcurrency;
        deleteRecycleBinMaxConcurrency = _deleteRecycleBinMaxConcurrency;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_MOVE, flags, options, callback, cookie, 1, reparsePointPolicy);

    const PathInfo source      = MakePathInfo(sourcePath);
    const PathInfo destination = MakePathInfo(destinationPath);

    HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        Debug::Warning(L"FileSystem: MoveItem failed to set paths for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(hr));
        return hr;
    }

    const unsigned int maxCopyConcurrency = std::clamp(copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    HRESULT itemHr                        = MovePathInternal(context, source, destination, true, flags, reparsePointPolicy, maxCopyConcurrency);
    if (FAILED(itemHr))
    {
        Debug::Warning(L"FileSystem: MoveItem failed for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(itemHr));
    }

    hr = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    context.completedItems = 1;
    return itemHr;
}

HRESULT STDMETHODCALLTYPE
FileSystem::DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    if (! path)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int deleteMaxConcurrency               = 1;
    unsigned int deleteRecycleBinMaxConcurrency     = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy             = _reparsePointPolicy;
        deleteMaxConcurrency           = _deleteMaxConcurrency;
        deleteRecycleBinMaxConcurrency = _deleteRecycleBinMaxConcurrency;
    }

    OperationContext context{};
    // totalItems is 0 because the plugin does not know recursive totals; the host may provide totals via pre-calculation.
    InitializeOperationContext(context, FILESYSTEM_DELETE, flags, options, callback, cookie, 0, reparsePointPolicy);
    const bool useRecycleBin                    = HasFlag(flags, FILESYSTEM_FLAG_USE_RECYCLE_BIN);
    const unsigned int maxConcurrencyFast       = std::clamp(deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency);
    const unsigned int maxConcurrencyRecycleBin = std::clamp(deleteRecycleBinMaxConcurrency, 1u, kMaxDeleteRecycleBinMaxConcurrency);
    context.deleteConcurrencyBudget             = useRecycleBin ? maxConcurrencyRecycleBin : maxConcurrencyFast;

    const PathInfo target = MakePathInfo(path);

    HRESULT hr = SetItemPaths(context, target.display.c_str(), nullptr);
    if (FAILED(hr))
    {
        Debug::Warning(L"FileSystem: DeleteItem failed to set path for '{}' (hr={:#x})", target.display, static_cast<uint32_t>(hr));
        return hr;
    }

    HRESULT itemHr = DeletePathInternal(context, target);
    if (FAILED(itemHr))
    {
        Debug::Warning(L"FileSystem: DeleteItem failed for '{}' (hr={:#x})", target.display, static_cast<uint32_t>(itemHr));
    }

    hr = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReportProgressForced(context, 0, 0);
    if (FAILED(hr))
    {
        return hr;
    }
    return itemHr;
}

HRESULT STDMETHODCALLTYPE FileSystem::RenameItem(const wchar_t* sourcePath,
                                                 const wchar_t* destinationPath,
                                                 FileSystemFlags flags,
                                                 const FileSystemOptions* options,
                                                 IFileSystemCallback* callback,
                                                 void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy = _reparsePointPolicy;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_RENAME, flags, options, callback, cookie, 1, reparsePointPolicy);

    const PathInfo source      = MakePathInfo(sourcePath);
    const PathInfo destination = MakePathInfo(destinationPath);

    HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    HRESULT itemHr = MovePathInternal(context, source, destination, false);
    hr             = ReportItemCompleted(context, 0, itemHr);
    if (FAILED(hr))
    {
        return hr;
    }

    context.completedItems = 1;
    return itemHr;
}

HRESULT STDMETHODCALLTYPE FileSystem::CopyItems(const wchar_t* const* sourcePaths,
                                                unsigned long count,
                                                const wchar_t* destinationFolder,
                                                FileSystemFlags flags,
                                                const FileSystemOptions* options,
                                                IFileSystemCallback* callback,
                                                void* cookie) noexcept
{
    if (! sourcePaths && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (! destinationFolder)
    {
        return E_POINTER;
    }

    if (destinationFolder[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int copyMoveMaxConcurrency             = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    const PathInfo destinationRoot    = MakePathInfo(destinationFolder);
    const unsigned int maxConcurrency = std::clamp(copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    const unsigned int concurrency    = std::min<unsigned int>(maxConcurrency, count);

    if (concurrency <= 1u)
    {
        OperationContext context{};
        InitializeOperationContext(context, FILESYSTEM_COPY, flags, options, callback, cookie, count, reparsePointPolicy);

        bool hadFailure = false;

        for (unsigned long index = 0; index < count; ++index)
        {
            const wchar_t* sourcePath = sourcePaths[index];
            if (! sourcePath)
            {
                return E_POINTER;
            }

            if (sourcePath[0] == L'\0')
            {
                return E_INVALIDARG;
            }

            const std::wstring_view leaf = GetPathLeaf(sourcePath);
            if (leaf.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }

            const PathInfo source      = MakePathInfo(sourcePath);
            const PathInfo destination = {AppendPath(destinationRoot.display, leaf), AppendPath(destinationRoot.extended, leaf)};

            HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
            if (FAILED(hr))
            {
                return hr;
            }

            context.reparseRootSourcePath      = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(source.display)));
            context.reparseRootDestinationPath = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(destination.display)));

            uint64_t bytesCopied = 0;
            HRESULT itemHr = CopyPathInternalWithDirectoryParallelism(context, source, destination, flags, reparsePointPolicy, maxConcurrency, &bytesCopied);

            hr = ReportItemCompleted(context, index, itemHr);
            if (FAILED(hr))
            {
                return hr;
            }

            context.completedItems += 1;

            if (FAILED(itemHr))
            {
                if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
                {
                    return itemHr;
                }

                hadFailure = true;
                if (! context.continueOnError)
                {
                    return itemHr;
                }
            }
        }

        if (hadFailure)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* sourcePath = sourcePaths[index];
        if (! sourcePath)
        {
            return E_POINTER;
        }

        if (sourcePath[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const std::wstring_view leaf = GetPathLeaf(sourcePath);
        if (leaf.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
    }

    FileSystemOptions sharedOptionsState{};
    if (options)
    {
        sharedOptionsState = *options;
    }
    sharedOptionsState.sizeBytes = sizeof(FileSystemOptions);

    ParallelOperationState parallel{};
    parallel.startTick = GetTickCount64();
    parallel.bandwidthLimitBytesPerSecond.store(sharedOptionsState.bandwidthLimitBytesPerSecond, std::memory_order_release);
    parallel.copyMoveTransferLimit.store(maxConcurrency, std::memory_order_release);

    const unsigned int nestedConcurrency = CalculateNestedCopyMoveConcurrency(maxConcurrency, concurrency);
    Debug::Perf::Emit(L"FileOps.CopyItems.NestedConcurrencyBudget", L"", nestedConcurrency, maxConcurrency, concurrency, S_OK);

    auto job = GetSharedFileOpsJobScheduler().StartJob(concurrency,
                                                       static_cast<size_t>(count),
                                                       [&](size_t index, uint64_t schedulerStreamId) noexcept
    {
        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            return;
        }

        if (index >= static_cast<size_t>(count))
        {
            return;
        }

        OperationContext context{};
        InitializeOperationContext(context, FILESYSTEM_COPY, flags, &sharedOptionsState, callback, cookie, count, reparsePointPolicy);
        context.options          = &sharedOptionsState;
        context.parallel         = &parallel;
        context.totalBytes       = 0; // let the host provide totals via pre-calc
        context.progressStreamId = schedulerStreamId;

        const unsigned long itemIndex = static_cast<unsigned long>((std::min)(index, static_cast<size_t>(ULONG_MAX)));
        const wchar_t* sourcePath     = sourcePaths[itemIndex];
        const std::wstring_view leaf  = GetPathLeaf(sourcePath);

        HRESULT hr = CheckCancel(context);
        if (FAILED(hr))
        {
            const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
            if (! stopOnError)
            {
                parallel.cancelRequested.store(true, std::memory_order_release);
            }
            return;
        }

        const PathInfo source      = MakePathInfo(sourcePath);
        const PathInfo destination = {AppendPath(destinationRoot.display, leaf), AppendPath(destinationRoot.extended, leaf)};

        hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
        if (FAILED(hr))
        {
            parallel.stopOnErrorRequested.store(true, std::memory_order_release);
            HRESULT expected = S_OK;
            static_cast<void>(parallel.firstError.compare_exchange_strong(expected, hr, std::memory_order_acq_rel));
            return;
        }

        context.reparseRootSourcePath      = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(source.display)));
        context.reparseRootDestinationPath = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(destination.display)));

        uint64_t bytesCopied = 0;
        HRESULT itemHr = CopyPathInternalWithDirectoryParallelism(context, source, destination, flags, reparsePointPolicy, nestedConcurrency, &bytesCopied);

        hr = ReportItemCompleted(context, itemIndex, itemHr);
        if (FAILED(hr))
        {
            const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
            if (! stopOnError)
            {
                parallel.cancelRequested.store(true, std::memory_order_release);
            }
            return;
        }

        parallel.completedItems.fetch_add(1u, std::memory_order_acq_rel);

        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                if (! stopOnError)
                {
                    parallel.cancelRequested.store(true, std::memory_order_release);
                }
                return;
            }

            parallel.hadFailure.store(true, std::memory_order_release);
            if (! context.continueOnError)
            {
                parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                HRESULT expected = S_OK;
                static_cast<void>(parallel.firstError.compare_exchange_strong(expected, itemHr, std::memory_order_acq_rel));
                return;
            }
        }
    });

    GetSharedFileOpsJobScheduler().WaitJob(job);

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
    {
        const HRESULT hr = parallel.firstError.load(std::memory_order_acquire);
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.hadFailure.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::MoveItems(const wchar_t* const* sourcePaths,
                                                unsigned long count,
                                                const wchar_t* destinationFolder,
                                                FileSystemFlags flags,
                                                const FileSystemOptions* options,
                                                IFileSystemCallback* callback,
                                                void* cookie) noexcept
{
    if (! sourcePaths && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (! destinationFolder)
    {
        return E_POINTER;
    }

    if (destinationFolder[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int copyMoveMaxConcurrency             = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    const PathInfo destinationRoot    = MakePathInfo(destinationFolder);
    const unsigned int maxConcurrency = std::clamp(copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    const unsigned int concurrency    = std::min<unsigned int>(maxConcurrency, count);

    if (concurrency <= 1u)
    {
        OperationContext context{};
        InitializeOperationContext(context, FILESYSTEM_MOVE, flags, options, callback, cookie, count, reparsePointPolicy);

        bool hadFailure = false;

        for (unsigned long index = 0; index < count; ++index)
        {
            const wchar_t* sourcePath = sourcePaths[index];
            if (! sourcePath)
            {
                return E_POINTER;
            }

            if (sourcePath[0] == L'\0')
            {
                return E_INVALIDARG;
            }

            const std::wstring_view leaf = GetPathLeaf(sourcePath);
            if (leaf.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }

            const PathInfo source      = MakePathInfo(sourcePath);
            const PathInfo destination = {AppendPath(destinationRoot.display, leaf), AppendPath(destinationRoot.extended, leaf)};

            HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
            if (FAILED(hr))
            {
                return hr;
            }

            HRESULT itemHr = MovePathInternal(context, source, destination, true, flags, reparsePointPolicy, maxConcurrency);
            hr             = ReportItemCompleted(context, index, itemHr);
            if (FAILED(hr))
            {
                return hr;
            }

            context.completedItems += 1;

            if (FAILED(itemHr))
            {
                if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
                {
                    return itemHr;
                }

                hadFailure = true;
                if (! context.continueOnError)
                {
                    return itemHr;
                }
            }
        }

        if (hadFailure)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* sourcePath = sourcePaths[index];
        if (! sourcePath)
        {
            return E_POINTER;
        }

        if (sourcePath[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const std::wstring_view leaf = GetPathLeaf(sourcePath);
        if (leaf.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
    }

    FileSystemOptions sharedOptionsState{};
    if (options)
    {
        sharedOptionsState = *options;
    }
    sharedOptionsState.sizeBytes = sizeof(FileSystemOptions);

    ParallelOperationState parallel{};
    parallel.startTick = GetTickCount64();
    parallel.bandwidthLimitBytesPerSecond.store(sharedOptionsState.bandwidthLimitBytesPerSecond, std::memory_order_release);
    parallel.copyMoveTransferLimit.store(maxConcurrency, std::memory_order_release);

    const unsigned int nestedConcurrency = CalculateNestedCopyMoveConcurrency(maxConcurrency, concurrency);
    Debug::Perf::Emit(L"FileOps.MoveItems.NestedConcurrencyBudget", L"", nestedConcurrency, maxConcurrency, concurrency, S_OK);

    auto job = GetSharedFileOpsJobScheduler().StartJob(concurrency,
                                                       static_cast<size_t>(count),
                                                       [&](size_t index, uint64_t schedulerStreamId) noexcept
    {
        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            return;
        }

        if (index >= static_cast<size_t>(count))
        {
            return;
        }

        OperationContext context{};
        InitializeOperationContext(context, FILESYSTEM_MOVE, flags, &sharedOptionsState, callback, cookie, count, reparsePointPolicy);
        context.options          = &sharedOptionsState;
        context.parallel         = &parallel;
        context.totalBytes       = 0; // let the host provide totals via pre-calc
        context.progressStreamId = schedulerStreamId;

        const unsigned long itemIndex = static_cast<unsigned long>((std::min)(index, static_cast<size_t>(ULONG_MAX)));
        const wchar_t* sourcePath     = sourcePaths[itemIndex];
        const std::wstring_view leaf  = GetPathLeaf(sourcePath);

        HRESULT hr = CheckCancel(context);
        if (FAILED(hr))
        {
            const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
            if (! stopOnError)
            {
                parallel.cancelRequested.store(true, std::memory_order_release);
            }
            return;
        }

        const PathInfo source      = MakePathInfo(sourcePath);
        const PathInfo destination = {AppendPath(destinationRoot.display, leaf), AppendPath(destinationRoot.extended, leaf)};

        hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
        if (FAILED(hr))
        {
            parallel.stopOnErrorRequested.store(true, std::memory_order_release);
            HRESULT expected = S_OK;
            static_cast<void>(parallel.firstError.compare_exchange_strong(expected, hr, std::memory_order_acq_rel));
            return;
        }

        const HRESULT itemHr = MovePathInternal(context, source, destination, true, flags, reparsePointPolicy, nestedConcurrency);

        hr = ReportItemCompleted(context, itemIndex, itemHr);
        if (FAILED(hr))
        {
            const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
            if (! stopOnError)
            {
                parallel.cancelRequested.store(true, std::memory_order_release);
            }
            return;
        }

        parallel.completedItems.fetch_add(1u, std::memory_order_acq_rel);

        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                if (! stopOnError)
                {
                    parallel.cancelRequested.store(true, std::memory_order_release);
                }
                return;
            }

            parallel.hadFailure.store(true, std::memory_order_release);
            if (! context.continueOnError)
            {
                parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                HRESULT expected = S_OK;
                static_cast<void>(parallel.firstError.compare_exchange_strong(expected, itemHr, std::memory_order_acq_rel));
                return;
            }
        }
    });

    GetSharedFileOpsJobScheduler().WaitJob(job);

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
    {
        const HRESULT hr = parallel.firstError.load(std::memory_order_acquire);
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.hadFailure.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::DeleteItems(const wchar_t* const* paths,
                                                  unsigned long count,
                                                  FileSystemFlags flags,
                                                  const FileSystemOptions* options,
                                                  IFileSystemCallback* callback,
                                                  void* cookie) noexcept
{
    if (! paths && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int deleteMaxConcurrency               = 1;
    unsigned int deleteRecycleBinMaxConcurrency     = 1;
    unsigned int recycleBinBatchSize                = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy             = _reparsePointPolicy;
        deleteMaxConcurrency           = _deleteMaxConcurrency;
        deleteRecycleBinMaxConcurrency = _deleteRecycleBinMaxConcurrency;
        recycleBinBatchSize            = _recycleBinBatchSize;
    }

    const bool useRecycleBin = HasFlag(flags, FILESYSTEM_FLAG_USE_RECYCLE_BIN);

    const unsigned int maxConcurrencyFast            = std::clamp(deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency);
    const unsigned int maxConcurrencyRecycleBin      = std::clamp(deleteRecycleBinMaxConcurrency, 1u, kMaxDeleteRecycleBinMaxConcurrency);
    const unsigned int configuredRecycleBinBatchSize = std::clamp(recycleBinBatchSize, 1u, kMaxRecycleBinBatchSize);
    const unsigned int maxConcurrency                = useRecycleBin ? maxConcurrencyRecycleBin : maxConcurrencyFast;
    constexpr unsigned int kMaxSharedConcurrency     = 8u;
    const unsigned int concurrency                   = std::min<unsigned int>(std::min<unsigned int>(maxConcurrency, count), kMaxSharedConcurrency);

    if (concurrency > 1u || (useRecycleBin && count > 1u))
    {
        std::vector<std::wstring> extendedPaths;
        std::vector<std::wstring> parentDirectories;
        extendedPaths.reserve(count);
        if (useRecycleBin)
        {
            parentDirectories.reserve(count);
        }
        for (unsigned long index = 0; index < count; ++index)
        {
            const wchar_t* path = paths[index];
            if (! path)
            {
                return E_POINTER;
            }

            if (path[0] == L'\0')
            {
                return E_INVALIDARG;
            }

            const PathInfo target = MakePathInfo(path);
            extendedPaths.emplace_back(target.extended);
            if (useRecycleBin)
            {
                parentDirectories.emplace_back(GetPathDirectory(target.extended));
            }
        }

        const auto isPrefixPath = [](const std::wstring& prefix, const std::wstring& candidate) noexcept -> bool
        {
            if (prefix.empty() || candidate.empty())
            {
                return false;
            }

            if (prefix.size() > candidate.size())
            {
                return false;
            }

            if (_wcsnicmp(prefix.c_str(), candidate.c_str(), prefix.size()) != 0)
            {
                return false;
            }

            if (candidate.size() == prefix.size())
            {
                return true;
            }

            const wchar_t last = prefix.back();
            if (last == L'\\' || last == L'/')
            {
                return true;
            }

            const wchar_t next = candidate[prefix.size()];
            return next == L'\\' || next == L'/';
        };

        std::vector<size_t> order;
        order.reserve(extendedPaths.size());
        for (size_t i = 0; i < extendedPaths.size(); ++i)
        {
            order.emplace_back(i);
        }

        std::ranges::sort(order, [&](size_t a, size_t b) noexcept { return _wcsicmp(extendedPaths[a].c_str(), extendedPaths[b].c_str()) < 0; });

        // Build a dependency graph for overlapping inputs:
        // - If A is a prefix of B, we must delete B before A to avoid parent/child races.
        // We only depend on the *immediate* ancestor; transitive ordering falls out naturally.
        std::vector<unsigned long> remainingDeps(static_cast<size_t>(count), 0u);
        std::vector<std::vector<unsigned long>> dependents(static_cast<size_t>(count));

        std::vector<unsigned long> stack;
        stack.reserve(order.size());
        for (const size_t index : order)
        {
            const unsigned long cur = static_cast<unsigned long>(index);

            while (! stack.empty())
            {
                const unsigned long parent = stack.back();
                if (isPrefixPath(extendedPaths[parent], extendedPaths[cur]))
                {
                    break;
                }
                stack.pop_back();
            }

            if (! stack.empty())
            {
                const unsigned long parent = stack.back();
                ++remainingDeps[parent];
                dependents[cur].push_back(parent);
            }

            stack.push_back(cur);
        }

        std::deque<unsigned long> ready;
        for (unsigned long i = 0; i < count; ++i)
        {
            if (remainingDeps[i] == 0)
            {
                ready.push_back(i);
            }
        }

        FileSystemOptions sharedOptionsState{};
        if (options)
        {
            sharedOptionsState = *options;
        }
        sharedOptionsState.sizeBytes = sizeof(FileSystemOptions);

        ParallelOperationState parallel{};
        parallel.startTick = GetTickCount64();

        std::mutex scheduleMutex;
        std::condition_variable scheduleCv;
        unsigned long remainingWork = count;

        auto job = GetSharedFileOpsJobScheduler().StartJob(concurrency,
                                                           concurrency,
                                                           [&](size_t /*workerIndex*/, uint64_t streamId) noexcept
        {
            [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();

            OperationContext context{};
            // totalItems is 0 because the plugin does not know recursive totals; the host may provide totals via pre-calculation.
            InitializeOperationContext(context, FILESYSTEM_DELETE, flags, &sharedOptionsState, callback, cookie, 0, reparsePointPolicy);
            context.options             = &sharedOptionsState;
            context.parallel            = &parallel;
            context.totalBytes          = 0; // host pre-calc provides totals when available
            context.progressStreamId    = streamId;
            context.recycleBinBatchSize = configuredRecycleBinBatchSize;

            auto signalFatal = [&](HRESULT fatalHr) noexcept
            {
                parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                HRESULT expected = S_OK;
                static_cast<void>(parallel.firstError.compare_exchange_strong(expected, fatalHr, std::memory_order_acq_rel));
                scheduleCv.notify_all();
            };

            auto finalizeItem = [&](unsigned long completedIndex, const PathInfo& completedTarget, HRESULT itemHr) noexcept -> bool
            {
                HRESULT hr = SetItemPaths(context, completedTarget.display.c_str(), nullptr);
                if (FAILED(hr))
                {
                    signalFatal(hr);
                    return false;
                }

                hr = ReportItemCompleted(context, completedIndex, itemHr);
                if (FAILED(hr))
                {
                    const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                    if (! stopOnError)
                    {
                        parallel.cancelRequested.store(true, std::memory_order_release);
                    }
                    scheduleCv.notify_all();
                    return false;
                }

                parallel.completedItems.fetch_add(1u, std::memory_order_acq_rel);
                hr = ReportProgress(context, 0, 0);
                if (FAILED(hr))
                {
                    const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                    if (! stopOnError)
                    {
                        parallel.cancelRequested.store(true, std::memory_order_release);
                    }
                    scheduleCv.notify_all();
                    return false;
                }

                if (FAILED(itemHr))
                {
                    if (IsCancellationHr(itemHr))
                    {
                        const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                        if (! stopOnError)
                        {
                            parallel.cancelRequested.store(true, std::memory_order_release);
                        }
                        scheduleCv.notify_all();
                        return false;
                    }

                    parallel.hadFailure.store(true, std::memory_order_release);
                    if (! context.continueOnError)
                    {
                        parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                        HRESULT expected = S_OK;
                        static_cast<void>(parallel.firstError.compare_exchange_strong(expected, itemHr, std::memory_order_acq_rel));
                        scheduleCv.notify_all();
                        return false;
                    }
                }

                {
                    std::unique_lock lock(scheduleMutex);
                    for (const unsigned long dependent : dependents[completedIndex])
                    {
                        if (remainingDeps[dependent] > 0)
                        {
                            --remainingDeps[dependent];
                            if (remainingDeps[dependent] == 0)
                            {
                                ready.push_back(dependent);
                            }
                        }
                    }

                    if (remainingWork > 0)
                    {
                        --remainingWork;
                    }
                }

                scheduleCv.notify_all();
                return true;
            };

            for (;;)
            {
                if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                {
                    return;
                }

                std::vector<unsigned long> claimedIndices;
                const size_t recycleBinBatchLimit = useRecycleBin ? static_cast<size_t>(std::max(context.recycleBinBatchSize, 1u)) : 1u;
                claimedIndices.reserve(useRecycleBin ? (std::min)(static_cast<size_t>(count), recycleBinBatchLimit) : 1u);
                {
                    std::unique_lock lock(scheduleMutex);
                    scheduleCv.wait(lock,
                                    [&] noexcept
                    {
                        return parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire) ||
                               remainingWork == 0 || ! ready.empty();
                    });

                    if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                    {
                        return;
                    }

                    if (remainingWork == 0)
                    {
                        return;
                    }

                    if (ready.empty())
                    {
                        continue;
                    }

                    const unsigned long index = ready.front();
                    ready.pop_front();
                    claimedIndices.push_back(index);

                    if (useRecycleBin && ! parentDirectories.empty())
                    {
                        const std::wstring& batchParent = parentDirectories[index];
                        if (! batchParent.empty())
                        {
                            for (auto it = ready.begin(); it != ready.end() && claimedIndices.size() < recycleBinBatchLimit;)
                            {
                                const unsigned long candidate = *it;
                                if (candidate < parentDirectories.size() && PathsEqualInsensitive(batchParent, parentDirectories[candidate]))
                                {
                                    claimedIndices.push_back(candidate);
                                    it = ready.erase(it);
                                }
                                else
                                {
                                    ++it;
                                }
                            }
                        }
                    }
                }

                if (claimedIndices.empty())
                {
                    continue;
                }

                if (useRecycleBin && claimedIndices.size() > 1u)
                {
                    std::vector<RecycleBinBatchEntry> batchEntries;
                    batchEntries.reserve(claimedIndices.size());
                    for (const unsigned long batchIndex : claimedIndices)
                    {
                        const wchar_t* batchPath = paths[batchIndex];
                        if (! batchPath || batchPath[0] == L'\0')
                        {
                            signalFatal(batchPath ? E_INVALIDARG : E_POINTER);
                            return;
                        }

                        RecycleBinBatchEntry entry{};
                        entry.itemIndex = batchIndex;
                        entry.path      = MakePathInfo(batchPath);
                        batchEntries.emplace_back(std::move(entry));
                    }

                    bool batchStarted = false;
                    HRESULT batchHr   = DeleteToRecycleBinBatched(context, batchEntries, &batchStarted);
                    if (FAILED(batchHr) && ! batchStarted)
                    {
                        Debug::Perf::EmitCounter(L"FileOps.RecycleBin.BatchFallbackCount");
                        for (const RecycleBinBatchEntry& entry : batchEntries)
                        {
                            const HRESULT itemHr = DeletePathInternal(context, entry.path);
                            if (! finalizeItem(entry.itemIndex, entry.path, itemHr))
                            {
                                return;
                            }
                        }

                        continue;
                    }

                    size_t observedCount = 0;
                    bool observedFailure = false;
                    for (const RecycleBinBatchEntry& entry : batchEntries)
                    {
                        if (! entry.observed)
                        {
                            continue;
                        }

                        ++observedCount;
                        observedFailure = observedFailure || (FAILED(entry.result) && ! IsCancellationHr(entry.result));
                        if (! finalizeItem(entry.itemIndex, entry.path, entry.result))
                        {
                            return;
                        }
                    }

                    if (observedCount != batchEntries.size())
                    {
                        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                        {
                            return;
                        }

                        if (FAILED(batchHr) && IsCancellationHr(batchHr))
                        {
                            parallel.cancelRequested.store(true, std::memory_order_release);
                            scheduleCv.notify_all();
                            return;
                        }

                        signalFatal(FAILED(batchHr) ? batchHr : E_UNEXPECTED);
                        return;
                    }

                    if (FAILED(batchHr))
                    {
                        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                        {
                            return;
                        }

                        if (IsCancellationHr(batchHr))
                        {
                            parallel.cancelRequested.store(true, std::memory_order_release);
                            scheduleCv.notify_all();
                            return;
                        }

                        if (! observedFailure)
                        {
                            signalFatal(batchHr);
                            return;
                        }
                    }

                    continue;
                }

                const unsigned long index = claimedIndices.front();
                const wchar_t* path       = paths[index];
                if (! path || path[0] == L'\0')
                {
                    signalFatal(path ? E_INVALIDARG : E_POINTER);
                    return;
                }

                const PathInfo target = MakePathInfo(path);
                const HRESULT itemHr  = DeletePathInternal(context, target);
                if (! finalizeItem(index, target, itemHr))
                {
                    return;
                }
            }
        });

        GetSharedFileOpsJobScheduler().WaitJob(job);

        if (parallel.cancelRequested.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            const HRESULT hr = parallel.firstError.load(std::memory_order_acquire);
            return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (parallel.hadFailure.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    OperationContext context{};
    // totalItems is 0 because the plugin does not know recursive totals; the host may provide totals via pre-calculation.
    InitializeOperationContext(context, FILESYSTEM_DELETE, flags, options, callback, cookie, 0, reparsePointPolicy);
    context.deleteConcurrencyBudget = maxConcurrency;

    bool hadFailure = false;

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* path = paths[index];
        if (! path)
        {
            return E_POINTER;
        }

        if (path[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const PathInfo target = MakePathInfo(path);

        HRESULT hr = SetItemPaths(context, target.display.c_str(), nullptr);
        if (FAILED(hr))
        {
            return hr;
        }

        HRESULT itemHr = DeletePathInternal(context, target);
        hr             = ReportItemCompleted(context, index, itemHr);
        if (FAILED(hr))
        {
            return hr;
        }

        context.completedItems += 1;
        hr = ReportProgress(context, 0, 0);
        if (FAILED(hr))
        {
            return hr;
        }

        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return itemHr;
            }

            hadFailure = true;
            if (! context.continueOnError)
            {
                return itemHr;
            }
        }
    }

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::RenameItems(const FileSystemRenamePair* items,
                                                  unsigned long count,
                                                  FileSystemFlags flags,
                                                  const FileSystemOptions* options,
                                                  IFileSystemCallback* callback,
                                                  void* cookie) noexcept
{
    if (! items && count > 0)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        if (items[index].sizeBytes != sizeof(FileSystemRenamePair))
        {
            return E_INVALIDARG;
        }
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;
    unsigned int copyMoveMaxConcurrency             = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    const unsigned int maxConcurrency = std::clamp(copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    const unsigned int concurrency    = std::min<unsigned int>(maxConcurrency, count);

    if (concurrency > 1u)
    {
        FileSystemOptions sharedOptionsState{};
        if (options)
        {
            sharedOptionsState = *options;
        }
        sharedOptionsState.sizeBytes = sizeof(FileSystemOptions);

        ParallelOperationState parallel{};
        parallel.startTick = GetTickCount64();
        parallel.bandwidthLimitBytesPerSecond.store(sharedOptionsState.bandwidthLimitBytesPerSecond, std::memory_order_release);

        auto job = GetSharedFileOpsJobScheduler().StartJob(concurrency,
                                                           count,
                                                           [&](size_t taskIndex, uint64_t streamId) noexcept
        {
            if (taskIndex >= count)
            {
                return;
            }

            if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
            {
                return;
            }

            [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();

            OperationContext context{};
            InitializeOperationContext(context, FILESYSTEM_RENAME, flags, &sharedOptionsState, callback, cookie, count, reparsePointPolicy);
            context.options          = &sharedOptionsState;
            context.parallel         = &parallel;
            context.totalBytes       = 0;
            context.progressStreamId = streamId;

            const FileSystemRenamePair& item = items[taskIndex];
            HRESULT itemHr                   = S_OK;

            if (! item.sourcePath || ! item.newName)
            {
                itemHr = E_POINTER;
            }
            else if (item.sourcePath[0] == L'\0' || item.newName[0] == L'\0')
            {
                itemHr = E_INVALIDARG;
            }
            else
            {
                const std::wstring_view newName = item.newName;
                if (ContainsPathSeparator(newName))
                {
                    itemHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                }
                else
                {
                    const std::wstring directory = GetPathDirectory(item.sourcePath);
                    if (directory.empty())
                    {
                        itemHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                    }
                    else
                    {
                        const std::wstring destinationPath = AppendPath(directory, newName);
                        const PathInfo source              = MakePathInfo(item.sourcePath);
                        const PathInfo destination         = MakePathInfo(destinationPath);

                        HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
                        if (SUCCEEDED(hr))
                        {
                            itemHr = MovePathInternal(context, source, destination, false);
                            hr     = ReportItemCompleted(context, static_cast<unsigned long>(taskIndex), itemHr);
                            if (FAILED(hr))
                            {
                                const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                                if (! stopOnError)
                                {
                                    parallel.cancelRequested.store(true, std::memory_order_release);
                                }
                                return;
                            }
                        }
                        else
                        {
                            itemHr = hr;
                        }
                    }
                }
            }

            parallel.completedItems.fetch_add(1u, std::memory_order_acq_rel);
            HRESULT hr = ReportProgress(context, 0, 0);
            if (FAILED(hr))
            {
                const bool stopOnError = parallel.stopOnErrorRequested.load(std::memory_order_acquire);
                if (! stopOnError)
                {
                    parallel.cancelRequested.store(true, std::memory_order_release);
                }
                return;
            }

            if (FAILED(itemHr))
            {
                if (IsCancellationHr(itemHr))
                {
                    parallel.cancelRequested.store(true, std::memory_order_release);
                    return;
                }

                parallel.hadFailure.store(true, std::memory_order_release);
                if (! context.continueOnError)
                {
                    parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                    HRESULT expected = S_OK;
                    static_cast<void>(parallel.firstError.compare_exchange_strong(expected, itemHr, std::memory_order_acq_rel));
                    return;
                }
            }
        });

        GetSharedFileOpsJobScheduler().WaitJob(job);

        if (parallel.cancelRequested.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            const HRESULT hr = parallel.firstError.load(std::memory_order_acquire);
            return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (parallel.hadFailure.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_RENAME, flags, options, callback, cookie, count, reparsePointPolicy);

    bool hadFailure = false;

    for (unsigned long index = 0; index < count; ++index)
    {
        const FileSystemRenamePair& item = items[index];
        if (! item.sourcePath || ! item.newName)
        {
            return E_POINTER;
        }

        if (item.sourcePath[0] == L'\0' || item.newName[0] == L'\0')
        {
            return E_INVALIDARG;
        }

        const std::wstring_view newName = item.newName;
        if (ContainsPathSeparator(newName))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const std::wstring directory = GetPathDirectory(item.sourcePath);
        if (directory.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

        const std::wstring destinationPath = AppendPath(directory, newName);
        const PathInfo source              = MakePathInfo(item.sourcePath);
        const PathInfo destination         = MakePathInfo(destinationPath);

        HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
        if (FAILED(hr))
        {
            return hr;
        }

        HRESULT itemHr = MovePathInternal(context, source, destination, false);
        hr             = ReportItemCompleted(context, index, itemHr);
        if (FAILED(hr))
        {
            return hr;
        }

        context.completedItems += 1;
        hr = ReportProgress(context, 0, 0);
        if (FAILED(hr))
        {
            return hr;
        }

        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return itemHr;
            }

            hadFailure = true;
            if (! context.continueOnError)
            {
                return itemHr;
            }
        }
    }

    if (hadFailure)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return S_OK;
}
