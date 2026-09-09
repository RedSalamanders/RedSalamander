#include "FileOperationTraversalPolicy.h"
#include "FileSystem.Internal.h"
#include "FileSystemRouteContract.h"
#include "PathUtils.h"
#include "SynchronousIoCancelWatch.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstring>
#include <cwchar>
#include <deque>
#include <filesystem>
#include <functional>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <ranges>
#include <span>
#include <thread>
#include <unordered_map>
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
std::atomic<uint64_t> g_copyTempPathCounter{0};

// Result of one step of a dynamic (self-feeding) job. Dynamic jobs replace the old
// long-lived `for(;;)` consumer loops: a worker processes ONE unit per dispatch and
// returns to the pool between units, so concurrent jobs round-robin fairly instead of
// one operation pinning every scheduler worker for its entire duration.
enum class DynamicStep
{
    Processed, // Ran one unit; more work may remain. Scheduler re-dispatches promptly.
    Idle,      // No unit ready right now, but the job is not finished. Worker parks until poked.
    Finished,  // No more work will ever appear. Wind the job down.
};

// R0f: a scheduler worker blocked inside a synchronous Win32 call (typically a dead SMB share)
// returns once the job's operation control reports cancel or deadline; see SynchronousIoCancelWatch.h.
[[nodiscard]] bool OperationOptionsRequestCancel(void* context) noexcept
{
    return FAILED(FileSystemCheckOperationControl(static_cast<const FileSystemOptions*>(context)));
}

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
        // Operation control source for the cancel watch of every unit this job dispatches.
        const FileSystemOptions* operationOptions = nullptr;

        // Dynamic (self-feeding) jobs: when set, dispatch is gated by *readyCountPtr (a
        // count of immediately-runnable units owned by the caller) instead of nextIndex/
        // totalItems, and the job completes when dynamicFinished is set with no work in
        // flight. processDynamic is invoked once per dispatch and returns a DynamicStep.
        bool isDynamic                           = false;
        const std::atomic<size_t>* readyCountPtr = nullptr;
        std::function<DynamicStep(uint64_t)> processDynamic;
        std::atomic<bool> dynamicFinished{false};

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
        return StartJob(nullptr, maxConcurrency, totalItems, std::move(processIndex));
    }

    JobPtr StartJob(const FileSystemOptions* operationOptions,
                    unsigned int maxConcurrency,
                    size_t totalItems,
                    std::function<void(size_t, uint64_t)> processIndex)
    {
        auto job              = std::make_shared<Job>();
        job->operationOptions = operationOptions;
        job->totalItems       = totalItems;
        job->processIndex     = std::move(processIndex);
        job->maxConcurrency   = std::max(1u, maxConcurrency);
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

    // Starts a dynamic, self-feeding parallel job. `readyCountPtr` points to a
    // caller-owned atomic that reports how many units are immediately runnable (the
    // caller increments it when it enqueues work and decrements it when a unit is
    // taken). The scheduler dispatches `step` on up to `maxConcurrency` workers as
    // SHORT work items — each worker returns to the pool after one unit, so other jobs
    // stay fair. A worker is dispatched only while *readyCountPtr > 0; when a step
    // returns DynamicStep::Finished the job winds down once nothing is in flight.
    JobPtr StartDynamicJob(unsigned int maxConcurrency, const std::atomic<size_t>* readyCountPtr, std::function<DynamicStep(uint64_t)> step)
    {
        return StartDynamicJob(nullptr, maxConcurrency, readyCountPtr, std::move(step));
    }

    JobPtr StartDynamicJob(const FileSystemOptions* operationOptions,
                           unsigned int maxConcurrency,
                           const std::atomic<size_t>* readyCountPtr,
                           std::function<DynamicStep(uint64_t)> step)
    {
        auto job              = std::make_shared<Job>();
        job->operationOptions = operationOptions;
        job->isDynamic        = true;
        job->maxConcurrency   = std::max(1u, maxConcurrency);
        job->readyCountPtr    = readyCountPtr;
        job->processDynamic   = std::move(step);

        ensureWorkers();

        if (_workers.empty())
        {
            // No worker pool: drain on the calling thread. The job is self-feeding, so
            // a single drainer makes progress until the step reports Finished.
            if (job->processDynamic)
            {
                for (;;)
                {
                    const DynamicStep result = job->processDynamic(0);
                    if (result == DynamicStep::Finished)
                    {
                        break;
                    }
                    if (result == DynamicStep::Idle && (! readyCountPtr || readyCountPtr->load(std::memory_order_acquire) == 0))
                    {
                        break;
                    }
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
                    cleanupJobsLocked();
                    if (job->done.load(std::memory_order_acquire))
                    {
                        break;
                    }
                    if (! hasSchedulableWorkLocked())
                    {
                        _cv.wait(lock);
                        cleanupJobsLocked();
                        if (job->done.load(std::memory_order_acquire))
                        {
                            break;
                        }
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

    void NotifyDynamicWorkAvailable() noexcept
    {
        _cv.notify_all();
    }

    void Shutdown() noexcept
    {
        ShutdownAndJoin();
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

        _cv.notify_all();

        // Shutdown is called only after producers have stopped submitting jobs.
        // Join workers first so stack-backed operation contexts cannot unwind while callbacks are still active.
        workers.clear();

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

            const bool finished = job->isDynamic ? job->dynamicFinished.load(std::memory_order_acquire) : (job->nextIndex >= job->totalItems);
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

            if (job->isDynamic)
            {
                if (job->dynamicFinished.load(std::memory_order_acquire))
                {
                    continue;
                }
                if (! job->readyCountPtr || job->readyCountPtr->load(std::memory_order_acquire) == 0)
                {
                    continue;
                }
                return true;
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

            if (job->isDynamic)
            {
                if (job->dynamicFinished.load(std::memory_order_acquire))
                {
                    continue;
                }
                if (! job->readyCountPtr || job->readyCountPtr->load(std::memory_order_acquire) == 0)
                {
                    continue;
                }

                outJob   = job;
                outIndex = 0; // Dynamic steps pull their own unit; the index is unused.
                job->inFlight += 1;

                _rrCursor = (idx + 1u) % jobCount;
                return true;
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
        bool dynamicFinished = false;
        if (job)
        {
            const Common::SynchronousIoCancelWatch::Scope cancelWatch(job->operationOptions != nullptr ? &OperationOptionsRequestCancel : nullptr,
                                                                      const_cast<FileSystemOptions*>(job->operationOptions));
            if (job->isDynamic)
            {
                if (job->processDynamic)
                {
                    dynamicFinished = job->processDynamic(streamId) == DynamicStep::Finished;
                }
            }
            else if (job->processIndex)
            {
                job->processIndex(index, streamId);
            }
        }

        {
            std::scoped_lock lock(_mutex);
            if (job)
            {
                if (job->inFlight > 0)
                {
                    job->inFlight -= 1;
                }
                if (dynamicFinished)
                {
                    job->dynamicFinished.store(true, std::memory_order_release);
                }
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
                cleanupJobsLocked();
                if (! hasSchedulableWorkLocked())
                {
                    _cv.wait(lock);
                    if (stopToken.stop_requested())
                    {
                        break;
                    }
                    cleanupJobsLocked();
                }
                if (stopToken.stop_requested())
                {
                    break;
                }
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

[[nodiscard]] HRESULT NormalizeReparseCopyFailure(HRESULT failure, bool allowOverwrite) noexcept
{
    if (! allowOverwrite && failure == HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    return failure;
}

#if defined(_DEBUG)
void RunDebugReparseCopyErrorMappingSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept
    {
        if (condition)
        {
            ++passed;
            return;
        }
        ++failed;
        Debug::Error(L"FileSystem debug selftest failed: {}", message);
    };

    constexpr HRESULT invalidParameter = HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER);
    check(NormalizeReparseCopyFailure(invalidParameter, true) == invalidParameter,
          L"staged symlink overwrite must preserve ERROR_INVALID_PARAMETER from the failed promotion");
    check(NormalizeReparseCopyFailure(invalidParameter, false) == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
          L"non-overwrite symlink copy may map ERROR_INVALID_PARAMETER to ERROR_NOT_SUPPORTED");
}

void RunDebugSharedFileOpsSchedulerShutdownSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept -> bool
    {
        if (condition)
        {
            ++passed;
            return true;
        }

        ++failed;
        Debug::Error(L"FileSystem debug selftest failed: {}", message);
        return false;
    };

    SharedFileOpsJobScheduler scheduler;
    if (! check(scheduler.EnsureWorkersAvailable(), L"Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker should create workers"))
    {
        return;
    }

    struct ProbeState
    {
        ProbeState()                             = default;
        ProbeState(const ProbeState&)            = delete;
        ProbeState(ProbeState&&)                 = delete;
        ProbeState& operator=(const ProbeState&) = delete;
        ProbeState& operator=(ProbeState&&)      = delete;

        std::mutex mutex;
        std::condition_variable cv;
        bool workerEntered    = false;
        bool releaseWorker    = false;
        bool callbackActive   = false;
        bool workerExited     = false;
        bool waiterReturned   = false;
        bool shutdownReturned = false;
    } probe;

    const auto job = scheduler.StartJob(1u,
                                        1u,
                                        [&](size_t, uint64_t) noexcept
    {
        std::unique_lock lock(probe.mutex);
        probe.workerEntered  = true;
        probe.callbackActive = true;
        probe.cv.notify_all();
        probe.cv.wait(lock, [&]() noexcept { return probe.releaseWorker; });
        probe.callbackActive = false;
        probe.workerExited   = true;
        probe.cv.notify_all();
    });

    const auto waitFor = [&](auto predicate, std::chrono::milliseconds timeout) noexcept -> bool
    {
        std::unique_lock lock(probe.mutex);
        return probe.cv.wait_for(lock, timeout, predicate);
    };

    if (! check(waitFor([&]() noexcept { return probe.workerEntered; }, std::chrono::milliseconds(5000)),
                L"Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker should enter the blocked worker callback"))
    {
        {
            std::scoped_lock lock(probe.mutex);
            probe.releaseWorker = true;
        }
        probe.cv.notify_all();
        scheduler.ShutdownAndJoin();
        return;
    }

    std::jthread waiter([&]() noexcept
    {
        scheduler.WaitJob(job);
        {
            std::scoped_lock lock(probe.mutex);
            probe.waiterReturned = true;
        }
        probe.cv.notify_all();
    });

    std::jthread shutdownThread([&]() noexcept
    {
        scheduler.ShutdownAndJoin();
        {
            std::scoped_lock lock(probe.mutex);
            probe.shutdownReturned = true;
        }
        probe.cv.notify_all();
    });

    static_cast<void>(waitFor([&]() noexcept { return probe.waiterReturned || probe.shutdownReturned; }, std::chrono::milliseconds(250)));

    bool waiterReturnedEarly   = false;
    bool shutdownReturnedEarly = false;
    {
        std::scoped_lock lock(probe.mutex);
        waiterReturnedEarly   = probe.waiterReturned;
        shutdownReturnedEarly = probe.shutdownReturned;
        probe.releaseWorker   = true;
    }
    probe.cv.notify_all();

    check(! waiterReturnedEarly, L"Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker must not release WaitJob before the worker exits");
    check(! shutdownReturnedEarly, L"Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker must not return from shutdown before the worker exits");
    check(waitFor([&]() noexcept { return probe.workerExited && probe.waiterReturned && probe.shutdownReturned; }, std::chrono::milliseconds(5000)),
          L"Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker should drain the worker, waiter, and shutdown thread after release");
}
#endif
} // namespace FileSystemInternal

namespace
{
#pragma warning(push)
// C4625 (copy ctor deleted), C4626 (copy assign deleted)
#pragma warning(disable : 4625 4626)
constexpr size_t kMaxBandwidthThrottleWorkers = 16u;

#if defined(ENABLE_TESTS)
constexpr std::wstring_view kAbortOwnedStageUnknownPathEnvVar       = L"REDSALAMANDER_FILEOPS_ABORT_OWNED_STAGE_UNKNOWN_PATH";
constexpr std::wstring_view kAbortOwnedStageUnknownFiredEnvVar      = L"REDSALAMANDER_FILEOPS_ABORT_OWNED_STAGE_UNKNOWN_FIRED";
constexpr std::wstring_view kDirectFinalRollbackSwapPathEnvVar      = L"REDSALAMANDER_FILEOPS_DIRECT_FINAL_ROLLBACK_SWAP_PATH";
constexpr std::wstring_view kDirectFinalRollbackSwapMovedPathEnvVar = L"REDSALAMANDER_FILEOPS_DIRECT_FINAL_ROLLBACK_SWAP_MOVED_PATH";
constexpr std::wstring_view kDirectFinalRollbackSwapFiredEnvVar     = L"REDSALAMANDER_FILEOPS_DIRECT_FINAL_ROLLBACK_SWAP_FIRED";
constexpr std::wstring_view kDirectFinalAbortFailPathEnvVar         = L"REDSALAMANDER_FILEOPS_DIRECT_FINAL_ABORT_FAIL_PATH";
constexpr std::wstring_view kDirectFinalAbortFailFiredEnvVar        = L"REDSALAMANDER_FILEOPS_DIRECT_FINAL_ABORT_FAIL_FIRED";

#if defined(ENABLE_TESTS)
std::atomic<uint64_t> g_recycleBinBatchTestCalls{0};
std::atomic<uint64_t> g_recycleBinBatchTestRequestedItems{0};
std::atomic<uint64_t> g_recycleBinBatchTestObservedItems{0};
std::atomic<uint64_t> g_recycleBinBatchTestFailedItems{0};
std::atomic<uint64_t> g_recycleBinBatchTestFallbacks{0};
std::atomic<uint64_t> g_recycleBinBatchTestMaxBatchSize{0};

void RecordRecycleBinBatchSizeForTest(uint64_t batchSize) noexcept
{
    uint64_t observed = g_recycleBinBatchTestMaxBatchSize.load(std::memory_order_relaxed);
    while (observed < batchSize && ! g_recycleBinBatchTestMaxBatchSize.compare_exchange_weak(observed, batchSize, std::memory_order_relaxed))
    {
    }
}
#endif

[[nodiscard]] bool ConsumeAbortOwnedStageUnknownInjection(std::wstring_view destinationPath) noexcept
{
    const DWORD required = GetEnvironmentVariableW(kAbortOwnedStageUnknownPathEnvVar.data(), nullptr, 0u);
    if (required == 0u)
    {
        return false;
    }
    std::wstring configured(required, L'\0');
    const DWORD written = GetEnvironmentVariableW(kAbortOwnedStageUnknownPathEnvVar.data(), configured.data(), static_cast<DWORD>(configured.size()));
    if (written == 0u || written >= configured.size())
    {
        return false;
    }
    configured.resize(written);
    if (! OrdinalString::EqualsNoCase(configured, destinationPath))
    {
        return false;
    }
    static_cast<void>(SetEnvironmentVariableW(kAbortOwnedStageUnknownPathEnvVar.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kAbortOwnedStageUnknownFiredEnvVar.data(), L"1"));
    Debug::Perf::EmitCounter(L"FileOps.Local.AbortOwnedStageUnknownInjected");
    return true;
}

[[nodiscard]] std::wstring ReadDirectFinalSelfTestPath(std::wstring_view name) noexcept
{
    const DWORD required = GetEnvironmentVariableW(name.data(), nullptr, 0u);
    if (required == 0u)
    {
        return {};
    }

    std::wstring value(required, L'\0');
    const DWORD written = GetEnvironmentVariableW(name.data(), value.data(), static_cast<DWORD>(value.size()));
    if (written == 0u || written >= value.size())
    {
        return {};
    }
    value.resize(written);
    return value;
}

[[nodiscard]] bool ShouldFailDirectFinalCopyForSelfTest(std::wstring_view destinationPath) noexcept
{
    const std::wstring swapPath  = ReadDirectFinalSelfTestPath(kDirectFinalRollbackSwapPathEnvVar);
    const std::wstring abortPath = ReadDirectFinalSelfTestPath(kDirectFinalAbortFailPathEnvVar);
    return (! swapPath.empty() && OrdinalString::EqualsNoCase(swapPath, destinationPath)) ||
           (! abortPath.empty() && OrdinalString::EqualsNoCase(abortPath, destinationPath));
}

[[nodiscard]] bool ConsumeDirectFinalAbortFailureForSelfTest(std::wstring_view destinationPath) noexcept
{
    const std::wstring configured = ReadDirectFinalSelfTestPath(kDirectFinalAbortFailPathEnvVar);
    if (configured.empty() || ! OrdinalString::EqualsNoCase(configured, destinationPath))
    {
        return false;
    }

    static_cast<void>(SetEnvironmentVariableW(kDirectFinalAbortFailPathEnvVar.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kDirectFinalAbortFailFiredEnvVar.data(), L"1"));
    Debug::Perf::EmitCounter(L"FileOps.Local.DirectFinalAbortFailureInjected");
    return true;
}

void MaybeInjectDirectFinalRollbackSwapForSelfTest(std::wstring_view destinationPath) noexcept
{
    const std::wstring configured = ReadDirectFinalSelfTestPath(kDirectFinalRollbackSwapPathEnvVar);
    if (configured.empty() || ! OrdinalString::EqualsNoCase(configured, destinationPath))
    {
        return;
    }

    const std::wstring movedPath = ReadDirectFinalSelfTestPath(kDirectFinalRollbackSwapMovedPathEnvVar);
    if (movedPath.empty() || ! MoveFileExW(configured.c_str(), movedPath.c_str(), MOVEFILE_REPLACE_EXISTING))
    {
        return;
    }

    wil::unique_handle replacement(CreateFileW(
        configured.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! replacement)
    {
        return;
    }

    constexpr std::string_view kForeignPayload = "r0a-foreign-replacement";
    DWORD written                              = 0u;
    if (! WriteFile(replacement.get(), kForeignPayload.data(), static_cast<DWORD>(kForeignPayload.size()), &written, nullptr) ||
        written != kForeignPayload.size())
    {
        return;
    }

    static_cast<void>(SetEnvironmentVariableW(kDirectFinalRollbackSwapPathEnvVar.data(), nullptr));
    static_cast<void>(SetEnvironmentVariableW(kDirectFinalRollbackSwapFiredEnvVar.data(), L"1"));
    Debug::Perf::EmitCounter(L"FileOps.Local.DirectFinalRollbackSwapInjected");
}

#endif

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

struct DeleteDiscoveryState final
{
    DeleteDiscoveryState()                                       = default;
    DeleteDiscoveryState(const DeleteDiscoveryState&)            = delete;
    DeleteDiscoveryState(DeleteDiscoveryState&&)                 = delete;
    DeleteDiscoveryState& operator=(const DeleteDiscoveryState&) = delete;
    DeleteDiscoveryState& operator=(DeleteDiscoveryState&&)      = delete;
    ~DeleteDiscoveryState()                                      = default;

    std::mutex mutex;
    uint64_t discoveredBytes       = 0;
    uint64_t discoveredFiles       = 0;
    uint64_t discoveredDirectories = 0;
};

enum class TrackedPublicationTruth : uint8_t
{
    NotObserved,
    NotPublished,
    Published,
    Unknown,
};

struct OperationContext
{
    OperationContext()                                   = default;
    OperationContext(const OperationContext&)            = delete;
    OperationContext(OperationContext&&)                 = delete;
    OperationContext& operator=(const OperationContext&) = delete;
    OperationContext& operator=(OperationContext&&)      = delete;

    FileSystemOperation type      = FILESYSTEM_COPY;
    IFileSystemCallback* callback = nullptr;
    void* callbackCookie          = nullptr;
    uint64_t progressStreamId     = 0;
    FileSystemOptions optionsState{};
    FileSystemOptions* options   = nullptr;
    unsigned long totalItems     = 0;
    unsigned long completedItems = 0;
    uint64_t totalBytes          = 0;
    uint64_t completedBytes      = 0;
    bool continueOnError         = false;
    bool allowOverwrite          = false;
    bool allowReplaceReadonly    = false;
    bool allowReplaceLink        = false;
    bool recursive               = false;
    bool useRecycleBin           = false;
    // Per-conflict grants from a single answered FileSystemIssue prompt (no apply-to-all).
    // They authorize only the child whose conflict loop set them: cleared when that child's
    // retry loop exits and when recursion enters a child directory, so one answer can never
    // silently authorize overwrites the user was not asked about.
    bool oneShotAllowOverwrite       = false;
    bool oneShotAllowReplaceReadonly = false;
    bool oneShotAllowReplaceLink     = false;
    wil::com_ptr<IFileSystemBoundObject> oneShotExpectedDestination;
    wil::com_ptr<IFileSystemObjectBinding> objectBinding;
    unsigned int deleteConcurrencyBudget                = 1;
    unsigned int recycleBinBatchSize                    = 1;
    DeleteDiscoveryState* deleteDiscovery               = nullptr;
    uint64_t deleteTraversalDepth                       = 0;
    uint64_t deleteTraversalMaxDepth                    = 0;
    uint64_t deleteTraversalMaxBatchEntries             = 0;
    uint64_t deleteTraversalRetainedFailureCount        = 0;
    uint64_t deleteTraversalRetainedFailurePathBytes    = 0;
    uint64_t deleteTraversalMaxRetainedFailureCount     = 0;
    uint64_t deleteTraversalMaxRetainedFailurePathBytes = 0;
    FileSystemArenaOwner itemArena;
    FileSystemArenaOwner progressArena;
    BandwidthThrottleWorkerMode bandwidthThrottleWorkerMode = BandwidthThrottleWorkerMode::PerWorkerSubBudget;
    const wchar_t* itemSource                               = nullptr;
    const wchar_t* itemDestination                          = nullptr;
    const wchar_t* progressSource                           = nullptr;
    const wchar_t* progressDestination                      = nullptr;

    ParallelOperationState* parallel = nullptr;

    ULONGLONG lastProgressReportTick = 0;

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::Preserve;
    std::wstring reparseRootSourcePath;
    std::wstring reparseRootDestinationPath;
    bool trackTopLevelOwnedStageMutation = false;
    std::atomic<TrackedPublicationTruth> trackedPublication{TrackedPublicationTruth::NotObserved};
    std::atomic<FileSystemOwnedStageDisposition> trackedOwnedStageDisposition{FileSystemOwnedStageDisposition::NotApplicable};
    // Whole-item destination truth for the receipt when no exact top-level stage was
    // observed: a merge into an existing directory, a failure before any stage, or a
    // subtree whose children own their own stages. Every owned-stage site records here
    // unconditionally; only the exact top-level destination also updates the fields above.
    std::atomic<bool> anyDestinationStageCreated{false};
    std::atomic<bool> anyDestinationStageRetained{false};
    std::atomic<bool> anyDestinationIncompleteRetained{false};
    std::atomic<bool> anyDestinationPublicationUnknown{false};
    std::atomic<bool> anyDestinationPublished{false};
    // Native Move truth (MovePathInternal only): NotObserved until the first attempt,
    // NotPublished while every issued rename is proved not committed, Published after the
    // one native mutation committed, Unknown when a rename or revert outcome is unproved.
    // Monotonic: Unknown/Published survive a later attempt so a Retry can never erase them.
    std::atomic<TrackedPublicationTruth> trackedNativeMove{TrackedPublicationTruth::NotObserved};
};
#pragma warning(pop)

// The receipt for a Copy item. An exact tracked top-level stage wins; otherwise the
// whole-item aggregates decide. They are recorded at every owned-stage site regardless
// of top-level tracking, so a failure that never reached a stage is a proved no-commit.
[[nodiscard]] bool SnapshotTrackedItemMutation(const OperationContext& context, FileSystemItemMutationResult& result) noexcept
{
    const TrackedPublicationTruth publication = context.trackedPublication.load(std::memory_order_acquire);
    if (publication != TrackedPublicationTruth::NotObserved)
    {
        result = FileSystemItemMutationResult{
            sizeof(FileSystemItemMutationResult),
            publication == TrackedPublicationTruth::Unknown ? FALSE : TRUE,
            publication == TrackedPublicationTruth::Published ? TRUE : FALSE,
            TRUE,
            context.trackedOwnedStageDisposition.load(std::memory_order_acquire),
        };
        return true;
    }
    // Copy never removes its source, so originalStillPresent is TRUE on every axis below.
    if (context.anyDestinationPublicationUnknown.load(std::memory_order_acquire))
    {
        result = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE, FileSystemOwnedStageDisposition::Unknown};
        return true;
    }
    if (context.anyDestinationPublished.load(std::memory_order_acquire))
    {
        // At least one destination object under this item was published: the destination
        // changed, even when a later child failed or the user canceled.
        result = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, TRUE, TRUE, FileSystemOwnedStageDisposition::NotApplicable};
        return true;
    }
    if (context.anyDestinationIncompleteRetained.load(std::memory_order_acquire))
    {
        result = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE, FileSystemOwnedStageDisposition::RetainedIncomplete};
        return true;
    }
    if (context.anyDestinationStageRetained.load(std::memory_order_acquire))
    {
        result = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE, FileSystemOwnedStageDisposition::Retained};
        return true;
    }
    // No destination object was published and every created stage was removed (or none
    // was ever created): the attempt is a proved no-commit and may be retried.
    result = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult),
                                          TRUE,
                                          FALSE,
                                          TRUE,
                                          context.anyDestinationStageCreated.load(std::memory_order_acquire) ? FileSystemOwnedStageDisposition::Removed
                                                                                                             : FileSystemOwnedStageDisposition::NotCreated};
    return true;
}

[[nodiscard]] bool ShouldTrackOwnedStageMutation(const OperationContext& context, const PathInfo& destination) noexcept
{
    return context.trackTopLevelOwnedStageMutation && context.itemDestination != nullptr &&
           OrdinalString::EqualsNoCase(context.itemDestination, destination.display);
}

// Records that an owned stage (or direct final object) for `destination` reached a new
// publication state. Aggregates always learn about it; the exact top-level destination
// also updates the tracked receipt fields.
void RecordOwnedStagePublication(OperationContext& context,
                                 const PathInfo& destination,
                                 TrackedPublicationTruth publication,
                                 FileSystemOwnedStageDisposition disposition) noexcept
{
    switch (publication)
    {
        case TrackedPublicationTruth::Published: context.anyDestinationPublished.store(true, std::memory_order_release); break;
        case TrackedPublicationTruth::Unknown: context.anyDestinationPublicationUnknown.store(true, std::memory_order_release); break;
        case TrackedPublicationTruth::NotPublished: context.anyDestinationStageCreated.store(true, std::memory_order_release); break;
        case TrackedPublicationTruth::NotObserved:
        default: break;
    }
    if (ShouldTrackOwnedStageMutation(context, destination))
    {
        context.trackedPublication.store(publication, std::memory_order_release);
        context.trackedOwnedStageDisposition.store(disposition, std::memory_order_release);
    }
}

// Records the final disposition of an owned stage after its abort/cleanup ran.
void RecordOwnedStageDisposition(OperationContext& context, const PathInfo& destination, FileSystemOwnedStageDisposition disposition) noexcept
{
    switch (disposition)
    {
        case FileSystemOwnedStageDisposition::Retained:
        case FileSystemOwnedStageDisposition::Unknown: context.anyDestinationStageRetained.store(true, std::memory_order_release); break;
        case FileSystemOwnedStageDisposition::RetainedIncomplete: context.anyDestinationIncompleteRetained.store(true, std::memory_order_release); break;
        case FileSystemOwnedStageDisposition::Removed:
        case FileSystemOwnedStageDisposition::NotCreated:
        case FileSystemOwnedStageDisposition::Published:
        case FileSystemOwnedStageDisposition::NotApplicable:
        default: break;
    }
    if (ShouldTrackOwnedStageMutation(context, destination))
    {
        context.trackedOwnedStageDisposition.store(disposition, std::memory_order_release);
    }
}

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

[[nodiscard]] HRESULT ValidateFileSystemOptions(const FileSystemOptions* options, FileSystemOperation operation) noexcept
{
    if (options == nullptr)
    {
        return S_OK;
    }
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }

    if (operation == FILESYSTEM_MOVE)
    {
        return options->moveMode == FILESYSTEM_MOVE_DEFAULT || options->moveMode == FILESYSTEM_MOVE_NATIVE_ONLY ? S_OK : E_INVALIDARG;
    }
    return options->moveMode == FILESYSTEM_MOVE_DEFAULT ? S_OK : E_INVALIDARG;
}

[[nodiscard]] FileSystemReparsePointPolicy ResolveReparsePointPolicy(FileSystemReparsePointPolicy configuredDefault, const FileSystemOptions* options) noexcept
{
    if (options == nullptr)
    {
        return configuredDefault;
    }
    return options->linkPolicy == FILESYSTEM_LINK_SKIP ? FileSystemReparsePointPolicy::Skip : FileSystemReparsePointPolicy::Preserve;
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

[[nodiscard]] bool IsNameSurrogateReparseTag(DWORD tag) noexcept
{
    return IsReparseTagNameSurrogate(tag) != FALSE;
}

[[nodiscard]] HRESULT IsNameSurrogateReparsePoint(const std::wstring& pathExtended, DWORD attributes, bool& isNameSurrogate) noexcept
{
    isNameSurrogate = false;
    if (! IsReparsePoint(attributes))
    {
        return S_OK;
    }

    wil::unique_handle handle(::CreateFileW(pathExtended.c_str(),
                                            FILE_READ_ATTRIBUTES,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                            nullptr));
    if (! handle)
    {
        const DWORD error = ::GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }

    FILE_ATTRIBUTE_TAG_INFO tagInfo{};
    if (::GetFileInformationByHandleEx(handle.get(), FileAttributeTagInfo, &tagInfo, sizeof(tagInfo)) == FALSE)
    {
        const DWORD error = ::GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }

    isNameSurrogate = IsNameSurrogateReparseTag(tagInfo.ReparseTag);
    return S_OK;
}

[[nodiscard]] bool IsDirectory(DWORD attributes) noexcept
{
    return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

enum class LocalCopyPathKind : uint8_t
{
    RegularFile,
    RegularDirectory,
    SemanticLink,
};

[[nodiscard]] LocalCopyPathKind LocalCopyPathKindFromFacts(DWORD attributes, bool isNameSurrogate) noexcept
{
    if (isNameSurrogate)
    {
        return LocalCopyPathKind::SemanticLink;
    }
    return IsDirectory(attributes) ? LocalCopyPathKind::RegularDirectory : LocalCopyPathKind::RegularFile;
}

[[nodiscard]] HRESULT ClassifyLocalCopyPathKind(const std::wstring& pathExtended, DWORD attributes, LocalCopyPathKind& kind) noexcept
{
    bool isNameSurrogate = false;
    const HRESULT hr     = IsNameSurrogateReparsePoint(pathExtended, attributes, isNameSurrogate);
    if (FAILED(hr))
    {
        return hr;
    }
    kind = LocalCopyPathKindFromFacts(attributes, isNameSurrogate);
    return S_OK;
}

[[nodiscard]] bool HasOverwriteGrant(const OperationContext& context) noexcept
{
    return context.allowOverwrite || context.oneShotAllowOverwrite;
}

[[nodiscard]] bool HasReplaceReadonlyGrant(const OperationContext& context) noexcept
{
    return context.allowReplaceReadonly || context.oneShotAllowReplaceReadonly;
}

[[nodiscard]] bool HasReplaceLinkGrant(const OperationContext& context) noexcept
{
    return context.allowReplaceLink || context.oneShotAllowReplaceLink;
}

void ClearOneShotGrants(OperationContext& context) noexcept
{
    context.oneShotAllowOverwrite       = false;
    context.oneShotAllowReplaceReadonly = false;
    context.oneShotAllowReplaceLink     = false;
    context.oneShotExpectedDestination.reset();
}

[[nodiscard]] bool IssueActionRequiresExactDestination(FileSystemIssueAction action) noexcept
{
    return action == FileSystemIssueAction::Overwrite || action == FileSystemIssueAction::ReplaceReadOnly || action == FileSystemIssueAction::ReplaceLink;
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

HRESULT ReadReparsePointDataHandle(HANDLE handle, ReparsePointData& out) noexcept
{
    out = {};
    if (handle == nullptr || handle == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    DWORD bytesReturned = 0;
    if (! DeviceIoControl(handle, FSCTL_GET_REPARSE_POINT, nullptr, 0, out.buffer.data(), static_cast<DWORD>(out.buffer.size()), &bytesReturned, nullptr))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if (bytesReturned < sizeof(ReparsePointHeader))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const auto* header      = reinterpret_cast<const ReparsePointHeader*>(out.buffer.data());
    const size_t totalBytes = sizeof(ReparsePointHeader) + static_cast<size_t>(header->dataBytes);
    if (totalBytes > bytesReturned || totalBytes > out.buffer.size())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    out.tag       = header->tag;
    out.sizeBytes = static_cast<DWORD>(totalBytes);
    return S_OK;
}

HRESULT WriteReparsePointDataHandle(HANDLE handle, const ReparsePointData& data) noexcept
{
    if (data.sizeBytes < sizeof(ReparsePointHeader) || data.sizeBytes > data.buffer.size())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (handle == nullptr || handle == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    DWORD bytesReturned = 0;
    if (! DeviceIoControl(handle, FSCTL_SET_REPARSE_POINT, const_cast<std::byte*>(data.buffer.data()), data.sizeBytes, nullptr, 0, &bytesReturned, nullptr))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

#ifdef ENABLE_TESTS
HRESULT WriteReparsePointData(const std::wstring& path, const ReparsePointData& data) noexcept
{
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
    return WriteReparsePointDataHandle(handle.get(), data);
}
#endif

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

unsigned int ResolveCopyMoveConcurrencyLimit(unsigned int configuredConcurrency, const FileSystemOptions* options, unsigned int maxConcurrency) noexcept
{
    unsigned int effective = std::clamp(configuredConcurrency, 1u, maxConcurrency);
    if (options != nullptr && options->copyMoveMaxConcurrency != 0u)
    {
        effective = std::min(effective, std::clamp(options->copyMoveMaxConcurrency, 1u, maxConcurrency));
    }
    return effective;
}

constexpr uint64_t kBandwidthThrottleBurstWindowMs    = 250ull;
constexpr uint64_t kBandwidthThrottleMinCapacityBytes = 64ull * 1024ull;
constexpr uint64_t kBandwidthThrottleMaxCapacityBytes = 4ull * 1024ull * 1024ull;
constexpr uint64_t kBandwidthThrottleBoundaryGuardMs  = 10ull;
constexpr DWORD kBandwidthThrottleCancelPollMs        = 10u;
#ifdef ENABLE_TESTS
constexpr std::wstring_view kBandwidthThrottleWorkerModeEnvVar  = L"REDSALAMANDER_FILEOPS_BW_WORKER_MODE";
constexpr std::wstring_view kRecycleFailurePathEnvVar           = L"REDSALAMANDER_FILEOPS_RECYCLE_FAIL_PATH";
constexpr std::wstring_view kCaseRenameEntropyFailurePathEnvVar = L"REDSALAMANDER_FILEOPS_CASE_RENAME_ENTROPY_FAIL_PATH";
constexpr std::wstring_view kCaseRenameTempCollisionPathEnvVar  = L"REDSALAMANDER_FILEOPS_CASE_RENAME_TEMP_COLLISION_PATH";
constexpr std::wstring_view kCaseRenameTempCollisionFiredEnvVar = L"REDSALAMANDER_FILEOPS_CASE_RENAME_TEMP_COLLISION_FIRED";
#endif
#ifdef ENABLE_TESTS
constexpr std::wstring_view kDeleteToctouSwapPathEnvVar   = L"REDSALAMANDER_FILEOPS_DELETE_TOCTOU_SWAP_PATH";
constexpr std::wstring_view kDeleteToctouSwapTargetEnvVar = L"REDSALAMANDER_FILEOPS_DELETE_TOCTOU_SWAP_TARGET";
constexpr std::wstring_view kDeleteToctouSwapFiredEnvVar  = L"REDSALAMANDER_FILEOPS_DELETE_TOCTOU_SWAP_FIRED";
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

#ifdef ENABLE_TESTS
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

#ifdef ENABLE_TESTS
[[nodiscard]] std::wstring GetSelfTestEnvironmentString(std::wstring_view name) noexcept
{
    if (name.empty())
    {
        return {};
    }

    const DWORD required = ::GetEnvironmentVariableW(name.data(), nullptr, 0u);
    if (required == 0)
    {
        return {};
    }

    std::wstring value(required, L'\0');
    const DWORD written = ::GetEnvironmentVariableW(name.data(), value.data(), required);
    if (written == 0 || written >= required)
    {
        return {};
    }

    value.resize(written);
    return value;
}

[[nodiscard]] bool TryInjectRecycleFailureForSelfTest(const PathInfo& path, FileSystemItemMutationResult* mutationResult) noexcept
{
    const std::wstring failurePath = GetSelfTestEnvironmentString(kRecycleFailurePathEnvVar);
    if (failurePath.empty())
    {
        return false;
    }

    const PathInfo failure = MakePathInfo(failurePath);
    if (! OrdinalString::EqualsNoCase(path.extended, failure.extended))
    {
        return false;
    }

    static_cast<void>(::SetEnvironmentVariableW(kRecycleFailurePathEnvVar.data(), nullptr));
    if (mutationResult != nullptr)
    {
        *mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
    }
    return true;
}

#ifdef ENABLE_TESTS
void MaybeInjectDeleteToctouSwapForSelfTest(const std::wstring& candidateExtended) noexcept
{
    const std::wstring swapPath = GetSelfTestEnvironmentString(kDeleteToctouSwapPathEnvVar);
    if (swapPath.empty())
    {
        return;
    }

    const PathInfo swapInfo = MakePathInfo(swapPath);
    if (! OrdinalString::EqualsNoCase(candidateExtended, swapInfo.extended))
    {
        return;
    }

    static_cast<void>(::SetEnvironmentVariableW(kDeleteToctouSwapPathEnvVar.data(), nullptr));

    const std::wstring targetPath = GetSelfTestEnvironmentString(kDeleteToctouSwapTargetEnvVar);
    if (targetPath.empty())
    {
        return;
    }

    const PathInfo targetInfo = MakePathInfo(targetPath);
    if (! ::RemoveDirectoryW(candidateExtended.c_str()))
    {
        Debug::Warning(L"FileSystem: delete TOCTOU selftest hook could not remove victim directory '{}' (error={}).",
                       swapInfo.display,
                       static_cast<unsigned long>(::GetLastError()));
        return;
    }

    if (! ::CreateDirectoryW(candidateExtended.c_str(), nullptr))
    {
        Debug::Warning(L"FileSystem: delete TOCTOU selftest hook could not recreate victim directory '{}' (error={}).",
                       swapInfo.display,
                       static_cast<unsigned long>(::GetLastError()));
        return;
    }

    ReparsePointData reparse{};
    const std::wstring normalizedTarget = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(targetInfo.display)));
    const HRESULT buildHr               = BuildMountPointReparseData(normalizedTarget, reparse);
    if (FAILED(buildHr))
    {
        Debug::Warning(L"FileSystem: delete TOCTOU selftest hook could not build reparse data for '{}' (hr={:#x}).",
                       targetInfo.display,
                       static_cast<unsigned long>(buildHr));
        return;
    }

    const HRESULT writeHr = WriteReparsePointData(candidateExtended, reparse);
    if (FAILED(writeHr))
    {
        Debug::Warning(L"FileSystem: delete TOCTOU selftest hook could not write victim reparse point '{}' (hr={:#x}).",
                       swapInfo.display,
                       static_cast<unsigned long>(writeHr));
        return;
    }

    Debug::Perf::EmitCounter(L"FileOps.Delete.DebugToctouSwapInjected");
    static_cast<void>(::SetEnvironmentVariableW(kDeleteToctouSwapFiredEnvVar.data(), L"1"));
}
#endif

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

    const HRESULT controlHr = NormalizeCancellation(FileSystemCheckOperationControl(context.options));
    if (FAILED(controlHr))
    {
        if (controlHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) && context.parallel)
        {
            context.parallel->cancelRequested.store(true, std::memory_order_release);
        }
        return controlHr;
    }

    // ABI v2 operationControl is the canonical source. The legacy callback remains a fallback for
    // direct callers that did not supply operation control; do not query the same host task twice.
    if (context.options != nullptr && context.options->operationControl != nullptr)
    {
        return S_OK;
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

    auto& throttleState           = context.parallel ? context.parallel->bandwidthThrottle : progressContext.bandwidthThrottle;
    const bool useWorkerSubBudget = context.parallel && context.bandwidthThrottleWorkerMode == BandwidthThrottleWorkerMode::PerWorkerSubBudget &&
                                    context.progressStreamId < throttleState.workerStates.size();
    size_t activeWorkerCount      = 1;
    if (useWorkerSubBudget)
    {
        // A throttled CopyFileExW callback can sleep longer than the callback cadence. Count the
        // transfer slots that are actually held so sleeping workers retain their share and the
        // remaining workers immediately inherit bandwidth when a transfer finishes.
        std::scoped_lock lock(context.parallel->copyMoveTransferMutex);
        activeWorkerCount = (std::max)(static_cast<size_t>(context.parallel->activeCopyMoveTransfers), size_t{1});
    }

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

        if (useWorkerSubBudget)
        {
            ParallelOperationState::BandwidthThrottleState::WorkerState& workerState = throttleState.workerStates[context.progressStreamId];
            workerBytesPerSecond                                                     = bytesPerSecond / activeWorkerCount;
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

HRESULT ReportItemCompleted(OperationContext& context,
                            unsigned long itemIndex,
                            HRESULT status,
                            const FileSystemItemMutationResult* mutationResult = nullptr) noexcept
{
    if (! context.callback)
    {
        return S_OK;
    }

    if (context.parallel)
    {
        std::scoped_lock lock(context.parallel->callbackMutex);

        HRESULT hr = context.callback->FileSystemItemCompleted(
            context.type, itemIndex, context.itemSource, context.itemDestination, status, mutationResult, context.options, context.callbackCookie);
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
        context.type, itemIndex, context.itemSource, context.itemDestination, status, mutationResult, context.options, context.callbackCookie);
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

        wil::com_ptr<IFileSystemBoundObject> expectedDestination;
        HRESULT hr = context.callback->FileSystemIssue(context.type,
                                                       context.progressSource,
                                                       context.progressDestination,
                                                       status,
                                                       action,
                                                       expectedDestination.put(),
                                                       context.options,
                                                       context.callbackCookie);
        hr         = NormalizeCancellation(hr);
        if (FAILED(hr))
        {
            return hr;
        }
        if (IssueActionRequiresExactDestination(*action) && ! expectedDestination)
        {
            return E_UNEXPECTED;
        }
        context.oneShotExpectedDestination = std::move(expectedDestination);

        if (context.options)
        {
            context.parallel->bandwidthLimitBytesPerSecond.store(context.options->bandwidthLimitBytesPerSecond, std::memory_order_release);
        }

        return CheckCancelLocked(context);
    }

    wil::com_ptr<IFileSystemBoundObject> expectedDestination;
    HRESULT hr = context.callback->FileSystemIssue(
        context.type, context.progressSource, context.progressDestination, status, action, expectedDestination.put(), context.options, context.callbackCookie);
    hr = NormalizeCancellation(hr);
    if (FAILED(hr))
    {
        return hr;
    }

    if (IssueActionRequiresExactDestination(*action) && ! expectedDestination)
    {
        return E_UNEXPECTED;
    }
    context.oneShotExpectedDestination = std::move(expectedDestination);

    return CheckCancel(context);
}

[[nodiscard]] HRESULT SelectKeepBothPath(const PathInfo& destination, bool isDirectory, PathInfo& selected) noexcept
{
    selected                          = {};
    constexpr size_t kMaximumAttempts = 10'000u;
    for (size_t ordinal = 2u; ordinal < 2u + kMaximumAttempts; ++ordinal)
    {
        std::wstring candidate;
        const HRESULT candidateHr = Common::Paths::BuildUniqueSiblingPathCandidate(destination.display, isDirectory, ordinal, candidate);
        if (FAILED(candidateHr))
        {
            return candidateHr;
        }

        const std::wstring extendedCandidate = ToExtendedPath(candidate);
        if (GetFileAttributesW(extendedCandidate.c_str()) == INVALID_FILE_ATTRIBUTES)
        {
            const DWORD error = GetLastError();
            if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND)
            {
                selected.display  = std::move(candidate);
                selected.extended = std::move(extendedCandidate);
                return S_OK;
            }
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
    }
    return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
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
    context.optionsState.sizeBytes                     = sizeof(FileSystemOptions);
    context.options                                    = &context.optionsState;
    context.totalItems                                 = totalItems;
    context.completedItems                             = 0;
    context.totalBytes                                 = 0;
    context.completedBytes                             = 0;
    context.bandwidthThrottleWorkerMode                = GetBandwidthThrottleWorkerModeOverride();
    context.continueOnError                            = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    context.allowOverwrite                             = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);
    context.allowReplaceReadonly                       = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    context.allowReplaceLink                           = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_LINK);
    context.recursive                                  = HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE);
    context.useRecycleBin                              = HasFlag(flags, FILESYSTEM_FLAG_USE_RECYCLE_BIN);
    context.deleteConcurrencyBudget                    = 1;
    context.deleteDiscovery                            = nullptr;
    context.deleteTraversalDepth                       = 0;
    context.deleteTraversalMaxDepth                    = 0;
    context.deleteTraversalMaxBatchEntries             = 0;
    context.deleteTraversalRetainedFailureCount        = 0;
    context.deleteTraversalRetainedFailurePathBytes    = 0;
    context.deleteTraversalMaxRetainedFailureCount     = 0;
    context.deleteTraversalMaxRetainedFailurePathBytes = 0;
    context.itemSource                                 = nullptr;
    context.itemDestination                            = nullptr;
    context.progressSource                             = nullptr;
    context.progressDestination                        = nullptr;
    // A host task snapshots this option before admission. The provider configuration is only
    // the default for direct callers that omit FileSystemOptions.
    context.reparsePointPolicy = ResolveReparsePointPolicy(reparsePointPolicy, options);
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

DWORD CALLBACK CopyProgressRoutine(LARGE_INTEGER totalFileSize,
                                   LARGE_INTEGER totalBytesTransferred,
                                   LARGE_INTEGER streamSize,
                                   LARGE_INTEGER streamBytesTransferred,
                                   DWORD streamNumber,
                                   DWORD callbackReason,
                                   HANDLE sourceFile,
                                   HANDLE destinationFile,
                                   LPVOID context) noexcept;

[[nodiscard]] std::wstring MakeCopyTempSiblingPath(std::wstring_view finalPath) noexcept
{
    const uint64_t counter = g_copyTempPathCounter.fetch_add(1u, std::memory_order_acq_rel);
    return std::format(L"{}.rs_copy_tmp_{:08X}_{:08X}_{:016X}",
                       finalPath,
                       static_cast<unsigned int>(::GetCurrentProcessId()),
                       static_cast<unsigned int>(::GetCurrentThreadId()),
                       static_cast<unsigned long long>(counter));
}

void RollBackStagedProgress(OperationContext& context, CopyProgressContext& progress) noexcept
{
    if (progress.lastItemBytesTransferred == 0u)
    {
        return;
    }

    if (! context.parallel)
    {
        context.completedBytes            = progress.itemBaseBytes;
        progress.lastItemBytesTransferred = 0u;
        return;
    }

    const uint64_t rollbackBytes = progress.lastItemBytesTransferred;
    uint64_t current             = context.parallel->completedBytes.load(std::memory_order_acquire);
    while (current >= rollbackBytes &&
           ! context.parallel->completedBytes.compare_exchange_weak(current, current - rollbackBytes, std::memory_order_acq_rel, std::memory_order_acquire))
    {
    }

    progress.lastItemBytesTransferred = 0u;
}

[[nodiscard]] HRESULT CopyFileExWithStagedOverwrite(OperationContext& context,
                                                    const PathInfo& source,
                                                    const PathInfo& destination,
                                                    [[maybe_unused]] DWORD copyFlags,
                                                    CopyProgressContext& progress,
                                                    uint64_t expectedFileBytes,
                                                    bool verifyTempSize,
                                                    [[maybe_unused]] bool replacementIsReparsePoint) noexcept
{
    if (! context.objectBinding || ! context.options)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    wil::com_ptr<IFileSystemBoundObject> expectedDestination = context.oneShotExpectedDestination;
    if (! expectedDestination)
    {
        constexpr FileSystemBindFlags destinationFlags =
            static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_PUBLICATION);
        const HRESULT bindHr = context.objectBinding->BindObject(destination.display.c_str(), destinationFlags, expectedDestination.put());
        if (FAILED(bindHr))
        {
            return bindHr;
        }
        if (! expectedDestination)
        {
            return E_UNEXPECTED;
        }
    }
    Debug::Perf::EmitCounter(L"FileOps.Conflict.ExpectedDestinationBoundCount");

    constexpr FileSystemBindFlags sourceFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_CONTENT | FILESYSTEM_BIND_READ_METADATA);
    wil::com_ptr<IFileSystemBoundObject> sourceAuthority;
    HRESULT hr = context.objectBinding->BindObject(source.display.c_str(), sourceFlags, sourceAuthority.put());
    if (FAILED(hr) || ! sourceAuthority)
    {
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    wil::com_ptr<IFileReader> reader;
    hr = sourceAuthority->OpenReader(context.options, reader.put());
    if (FAILED(hr) || ! reader)
    {
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    uint64_t readerSize = 0u;
    hr                  = reader->GetSize(&readerSize);
    if (FAILED(hr))
    {
        return hr;
    }
    if (verifyTempSize && readerSize != expectedFileBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_INVALID);
    }

    std::wstring stagePath;
    wil::com_ptr<IFileWriter> writer;
    wil::com_ptr<IFileSystemBoundObject> ownedStage;
    constexpr unsigned int kMaxStageAttempts = 32u;
    for (unsigned int attempt = 0u; attempt < kMaxStageAttempts; ++attempt)
    {
        stagePath = MakeCopyTempSiblingPath(destination.display);
        if (stagePath.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
        }
        hr = context.objectBinding->CreateExclusiveWriter(stagePath.c_str(), context.options, writer.put(), ownedStage.put());
        if (SUCCEEDED(hr))
        {
            break;
        }
        writer.reset();
        ownedStage.reset();
        if (hr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && hr != HRESULT_FROM_WIN32(ERROR_FILE_EXISTS))
        {
            return hr;
        }
    }
    if (! writer || ! ownedStage)
    {
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }
    RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::NotPublished, FileSystemOwnedStageDisposition::Retained);

    bool publicationUnknown                = false;
    bool published                         = false;
    const auto finishWithOwnedStageCleanup = [&](HRESULT operationHr) noexcept -> HRESULT
    {
        if (published || publicationUnknown || ! ownedStage)
        {
            return operationHr;
        }
#if defined(ENABLE_TESTS)
        if (ConsumeAbortOwnedStageUnknownInjection(destination.display))
        {
            RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::Unknown);
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
#endif
        FileSystemConditionalMutationResult abortResult{};
        abortResult.sizeBytes                  = sizeof(abortResult);
        const FileSystemOptions cleanupOptions = MakeOwnedStageCleanupOptions(context.options);
        const HRESULT abortHr                  = ownedStage->AbortOwnedObject(&cleanupOptions, &abortResult);
        if (abortResult.outcomeKnown == FALSE || FAILED(abortHr) || abortResult.mutationCommitted == FALSE || abortResult.originalStillPresent != FALSE)
        {
            RecordOwnedStageDisposition(
                context, destination, abortResult.outcomeKnown == FALSE ? FileSystemOwnedStageDisposition::Unknown : FileSystemOwnedStageDisposition::Retained);
            Debug::Perf::EmitCounter(L"fileops.local.copy_stage_cleanup_unknown");
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::Removed);
        return operationHr;
    };

    constexpr unsigned long kBufferBytes = 1024u * 1024u;
    auto buffer                          = std::unique_ptr<std::byte[]>(new (std::nothrow) std::byte[kBufferBytes]);
    if (! buffer)
    {
        return finishWithOwnedStageCleanup(E_OUTOFMEMORY);
    }

    uint64_t copiedBytes = 0u;
    while (copiedBytes < readerSize)
    {
        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return finishWithOwnedStageCleanup(hr);
        }

        const unsigned long toRead = static_cast<unsigned long>((std::min)(static_cast<uint64_t>(kBufferBytes), readerSize - copiedBytes));
        unsigned long bytesRead    = 0u;
        hr                         = reader->Read(buffer.get(), toRead, &bytesRead);
        if (FAILED(hr) || bytesRead == 0u || bytesRead > toRead)
        {
            return finishWithOwnedStageCleanup(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_HANDLE_EOF));
        }

        unsigned long consumed = 0u;
        while (consumed < bytesRead)
        {
            unsigned long bytesWritten = 0u;
            hr                         = writer->Write(buffer.get() + consumed, bytesRead - consumed, &bytesWritten);
            if (FAILED(hr) || bytesWritten == 0u || bytesWritten > bytesRead - consumed)
            {
                return finishWithOwnedStageCleanup(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_WRITE_FAULT));
            }
            consumed += bytesWritten;
        }

        copiedBytes += bytesRead;
        LARGE_INTEGER total{};
        LARGE_INTEGER completed{};
        total.QuadPart     = static_cast<LONGLONG>(readerSize);
        completed.QuadPart = static_cast<LONGLONG>(copiedBytes);
        if (CopyProgressRoutine(total, completed, {}, {}, 0u, 0u, nullptr, nullptr, &progress) != PROGRESS_CONTINUE)
        {
            return finishWithOwnedStageCleanup(HRESULT_FROM_WIN32(ERROR_CANCELLED));
        }
    }

    hr = writer->Commit();
    if (FAILED(hr))
    {
        return finishWithOwnedStageCleanup(hr);
    }
    writer.reset();

    FileSystemBasicInformation sourceInfo{};
    sourceInfo.sizeBytes = sizeof(sourceInfo);
    hr                   = sourceAuthority->GetBasicInformation(&sourceInfo);
    if (SUCCEEDED(hr))
    {
        hr = ownedStage->SetBasicInformation(&sourceInfo);
    }
    if (FAILED(hr) && hr != E_NOTIMPL && hr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
    {
        return finishWithOwnedStageCleanup(hr);
    }

    FileSystemFlags publicationFlags = FILESYSTEM_FLAG_ALLOW_OVERWRITE;
    if (HasReplaceReadonlyGrant(context))
    {
        publicationFlags = static_cast<FileSystemFlags>(publicationFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
    }
    if (HasReplaceLinkGrant(context))
    {
        publicationFlags = static_cast<FileSystemFlags>(publicationFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_LINK);
    }

    FileSystemConditionalMutationResult mutationResult{};
    mutationResult.sizeBytes = sizeof(mutationResult);
    wil::com_ptr<IFileSystemBoundObject> publishedAuthority;
    hr = ownedStage->PublishAs(
        destination.display.c_str(), expectedDestination.get(), publicationFlags, context.options, &mutationResult, publishedAuthority.put());
    if (mutationResult.outcomeKnown == FALSE)
    {
        publicationUnknown = true;
        RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::Unknown, FileSystemOwnedStageDisposition::Unknown);
        return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
    }
    if (mutationResult.mutationCommitted == FALSE)
    {
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_INVALID) || hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            Debug::Perf::EmitCounter(L"FileOps.Conflict.ExpectedDestinationMismatchCount");
        }
        return finishWithOwnedStageCleanup(FAILED(hr) ? hr : E_UNEXPECTED);
    }
    published = true;
    RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::Published, FileSystemOwnedStageDisposition::Published);
    if (! publishedAuthority || mutationResult.originalStillPresent != FALSE)
    {
        return E_UNEXPECTED;
    }
    return hr;
}

struct DeleteHandleSnapshot final
{
    DeleteHandleSnapshot()                                       = default;
    DeleteHandleSnapshot(const DeleteHandleSnapshot&)            = delete;
    DeleteHandleSnapshot& operator=(const DeleteHandleSnapshot&) = delete;
    DeleteHandleSnapshot(DeleteHandleSnapshot&&)                 = default;
    DeleteHandleSnapshot& operator=(DeleteHandleSnapshot&&)      = default;

    wil::unique_handle handle;
    DWORD attributes   = 0;
    uint64_t fileBytes = 0;
};

HRESULT OpenPathForDeleteNoFollow(const std::wstring& pathExtended, DeleteHandleSnapshot& snapshot) noexcept
{
    snapshot = {};

    wil::unique_handle handle(::CreateFileW(pathExtended.c_str(),
                                            DELETE | FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
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

    snapshot.attributes = info.dwFileAttributes;
    if (! IsDirectory(snapshot.attributes))
    {
        LARGE_INTEGER size{};
        if (::GetFileSizeEx(handle.get(), &size) && size.QuadPart > 0)
        {
            snapshot.fileBytes = static_cast<uint64_t>(size.QuadPart);
        }
    }
    snapshot.handle = std::move(handle);
    return S_OK;
}

HRESULT SetHandleReadonlyAttribute(DeleteHandleSnapshot& snapshot, bool readonly) noexcept
{
    FILE_BASIC_INFO basic{};
    if (! ::GetFileInformationByHandleEx(snapshot.handle.get(), FileBasicInfo, &basic, sizeof(basic)))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    if (readonly)
    {
        basic.FileAttributes = snapshot.attributes | FILE_ATTRIBUTE_READONLY;
    }
    else
    {
        basic.FileAttributes = snapshot.attributes & ~FILE_ATTRIBUTE_READONLY;
        if (basic.FileAttributes == 0)
        {
            basic.FileAttributes = FILE_ATTRIBUTE_NORMAL;
        }
    }

    if (! ::SetFileInformationByHandle(snapshot.handle.get(), FileBasicInfo, &basic, sizeof(basic)))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return S_OK;
}

HRESULT MarkHandleForDeletion(DeleteHandleSnapshot& snapshot) noexcept
{
#ifdef FILE_DISPOSITION_FLAG_DELETE
    FILE_DISPOSITION_INFO_EX dispositionEx{};
    dispositionEx.Flags = FILE_DISPOSITION_FLAG_DELETE | FILE_DISPOSITION_FLAG_POSIX_SEMANTICS;
    if (::SetFileInformationByHandle(snapshot.handle.get(), FileDispositionInfoEx, &dispositionEx, sizeof(dispositionEx)))
    {
        snapshot.handle.reset();
        return S_OK;
    }

    const DWORD exError = ::GetLastError();
    if (exError != ERROR_INVALID_PARAMETER && exError != ERROR_NOT_SUPPORTED)
    {
        return HRESULT_FROM_WIN32(exError);
    }
#endif

    FILE_DISPOSITION_INFO disposition{};
    disposition.DeleteFile = TRUE;
    if (! ::SetFileInformationByHandle(snapshot.handle.get(), FileDispositionInfo, &disposition, sizeof(disposition)))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    snapshot.handle.reset();
    return S_OK;
}

HRESULT DeleteOpenedPathNoFollow(OperationContext& context, DeleteHandleSnapshot& snapshot, bool countCompletion) noexcept
{
    bool restoreReadonly      = false;
    auto restoreReadonlyScope = wil::scope_exit([&]() noexcept
    {
        if (restoreReadonly)
        {
            static_cast<void>(SetHandleReadonlyAttribute(snapshot, true));
        }
    });

    if ((snapshot.attributes & FILE_ATTRIBUTE_READONLY) != 0)
    {
        if (! context.allowReplaceReadonly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        const HRESULT clearHr = SetHandleReadonlyAttribute(snapshot, false);
        if (FAILED(clearHr))
        {
            return clearHr;
        }
        restoreReadonly = true;
    }

    const HRESULT deleteHr = MarkHandleForDeletion(snapshot);
    if (FAILED(deleteHr))
    {
        return deleteHr;
    }

    restoreReadonly = false;
    if (countCompletion)
    {
        AddCompletedItems(context, 1);
        if (! IsDirectory(snapshot.attributes))
        {
            AddCompletedBytes(context, snapshot.fileBytes);
        }
    }

    return S_OK;
}

// Directory frames are opened ONCE with list access and enumerated THROUGH the verified
// no-follow handle: there is no path re-walk between the reparse check and the listing, so a
// concurrent swap-to-junction cannot redirect the enumeration into the link target (TOCTOU).
struct DirectoryEnumHandle final
{
    wil::unique_handle handle;
    DWORD attributes = 0;

    DirectoryEnumHandle()                                          = default;
    DirectoryEnumHandle(const DirectoryEnumHandle&)                = delete;
    DirectoryEnumHandle& operator=(const DirectoryEnumHandle&)     = delete;
    DirectoryEnumHandle(DirectoryEnumHandle&&) noexcept            = default;
    DirectoryEnumHandle& operator=(DirectoryEnumHandle&&) noexcept = default;
};

[[nodiscard]] HRESULT OpenDirectoryForEnumerationNoFollow(const std::wstring& pathExtended, DirectoryEnumHandle& out) noexcept
{
    out = {};

    wil::unique_handle handle(::CreateFileW(pathExtended.c_str(),
                                            FILE_LIST_DIRECTORY | FILE_READ_ATTRIBUTES,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
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

    out.attributes = info.dwFileAttributes;
    out.handle     = std::move(handle);
    return S_OK;
}

// Streams FILE_FULL_DIR_INFO records through the verified handle. onEntry(name, attributes,
// sizeBytes) returning anything other than S_OK stops the walk (S_FALSE = benign early stop);
// "." and ".." are filtered here.
template <typename TOnEntry> [[nodiscard]] HRESULT EnumerateDirectoryByHandle(HANDLE directory, TOnEntry&& onEntry) noexcept
{
    alignas(8) std::array<std::byte, 64u * 1024u> buffer{};
    bool restart = true;
    for (;;)
    {
        if (! ::GetFileInformationByHandleEx(
                directory, restart ? FileFullDirectoryRestartInfo : FileFullDirectoryInfo, buffer.data(), static_cast<DWORD>(buffer.size())))
        {
            const DWORD error = ::GetLastError();
            if (error == ERROR_NO_MORE_FILES)
            {
                return S_OK;
            }
            return HRESULT_FROM_WIN32(error);
        }
        restart = false;

        const std::byte* cursor = buffer.data();
        for (;;)
        {
            const auto* info = reinterpret_cast<const FILE_FULL_DIR_INFO*>(cursor);
            const std::wstring_view name(info->FileName, info->FileNameLength / sizeof(wchar_t));
            if (name != L"." && name != L"..")
            {
                const HRESULT hr = onEntry(name, info->FileAttributes, static_cast<uint64_t>(info->EndOfFile.QuadPart));
                if (hr != S_OK)
                {
                    return hr;
                }
            }
            if (info->NextEntryOffset == 0)
            {
                break;
            }
            cursor += info->NextEntryOffset;
        }
    }
}

// Collects child names through a verified enumeration handle. If the path turned out not to be
// a real directory anymore (swapped to a file/reparse since the caller's check), returns S_FALSE
// without filling children — the caller deletes the path as what it is now, never traversing.
[[nodiscard]] HRESULT CollectDirectoryChildrenVerified(const std::wstring& directoryExtended, std::vector<std::wstring>& children) noexcept
{
    DirectoryEnumHandle frame{};
    HRESULT hr = OpenDirectoryForEnumerationNoFollow(directoryExtended, frame);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! IsDirectory(frame.attributes) || IsReparsePoint(frame.attributes))
    {
        return S_FALSE;
    }

    hr = EnumerateDirectoryByHandle(frame.handle.get(),
                                    [&](std::wstring_view name, DWORD /*childAttributes*/, uint64_t /*sizeBytes*/) noexcept -> HRESULT
    {
        children.emplace_back(name);
        return S_OK;
    });
    return FAILED(hr) ? hr : S_OK;
}

HRESULT RemoveDirectoryRecursiveNoFollow(OperationContext& context, const std::wstring& directoryExtended) noexcept
{
    HRESULT hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    std::vector<std::wstring> childNames;
    hr = CollectDirectoryChildrenVerified(directoryExtended, childNames);
    if (FAILED(hr))
    {
        if (HRESULT_CODE(hr) == ERROR_FILE_NOT_FOUND || HRESULT_CODE(hr) == ERROR_PATH_NOT_FOUND)
        {
            return S_OK;
        }
        return hr;
    }
    if (hr == S_FALSE)
    {
        // Changed identity since the caller's check: delete the path as what it is now.
        DeleteHandleSnapshot swappedSnapshot{};
        hr = OpenPathForDeleteNoFollow(directoryExtended, swappedSnapshot);
        if (FAILED(hr))
        {
            return hr;
        }
        return DeleteOpenedPathNoFollow(context, swappedSnapshot, false);
    }

    for (const std::wstring& childName : childNames)
    {
        const std::wstring child = AppendPath(directoryExtended, childName.c_str());
        hr                       = RemovePathForOverwrite(context, child);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    DeleteHandleSnapshot snapshot{};
    hr = OpenPathForDeleteNoFollow(directoryExtended, snapshot);
    if (FAILED(hr))
    {
        return hr;
    }
    return DeleteOpenedPathNoFollow(context, snapshot, false);
}

HRESULT RemovePathForOverwrite(OperationContext& context, const std::wstring& pathExtended) noexcept
{
    HRESULT hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }

    DeleteHandleSnapshot snapshot{};
    hr = OpenPathForDeleteNoFollow(pathExtended, snapshot);
    if (FAILED(hr))
    {
        return hr;
    }

    if (IsDirectory(snapshot.attributes))
    {
        if (IsReparsePoint(snapshot.attributes))
        {
            return DeleteOpenedPathNoFollow(context, snapshot, false);
        }

        if (snapshot.handle)
        {
            snapshot.handle.reset();
        }
        return RemoveDirectoryRecursiveNoFollow(context, pathExtended);
    }

    return DeleteOpenedPathNoFollow(context, snapshot, false);
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

[[nodiscard]] HRESULT CopyFileWithRetainedFinalObject(
    OperationContext& context, const PathInfo& source, const PathInfo& destination, CopyProgressContext& progress, uint64_t expectedFileBytes) noexcept
{
    constexpr unsigned long kBufferBytes = 1024u * 1024u;
    const auto metricStart               = std::chrono::steady_clock::now();
    HRESULT metricStatus                 = E_UNEXPECTED;
    const auto emitMetric                = wil::scope_exit([&]() noexcept
    {
        const uint64_t elapsedUs =
            static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - metricStart).count());
        Debug::Perf::Emit(L"FileOps.Local.DirectFinalCopyUs", L"shape=new-name-local-regular-file", elapsedUs, expectedFileBytes, kBufferBytes, metricStatus);
    });
    const auto finish                    = [&](HRESULT hr) noexcept -> HRESULT
    {
        metricStatus = hr;
        return hr;
    };

    if (! context.objectBinding || ! context.options)
    {
        return finish(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    }

    constexpr FileSystemBindFlags sourceFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_CONTENT | FILESYSTEM_BIND_READ_METADATA);
    wil::com_ptr<IFileSystemBoundObject> sourceAuthority;
    HRESULT hr = context.objectBinding->BindObject(source.display.c_str(), sourceFlags, sourceAuthority.put());
    if (FAILED(hr) || ! sourceAuthority)
    {
        return finish(FAILED(hr) ? hr : E_UNEXPECTED);
    }

    wil::com_ptr<IFileReader> reader;
    hr = sourceAuthority->OpenReader(context.options, reader.put());
    if (FAILED(hr) || ! reader)
    {
        return finish(FAILED(hr) ? hr : E_UNEXPECTED);
    }

    uint64_t readerSize = 0u;
    hr                  = reader->GetSize(&readerSize);
    if (FAILED(hr))
    {
        return finish(hr);
    }
    if (readerSize != expectedFileBytes)
    {
        return finish(HRESULT_FROM_WIN32(ERROR_FILE_INVALID));
    }

    wil::com_ptr<IFileSystemBoundMetadata> sourceMetadata;
    hr = sourceAuthority->QueryInterface(IID_PPV_ARGS(sourceMetadata.addressof()));
    if (FAILED(hr) || ! sourceMetadata)
    {
        return finish(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    }

    wil::com_ptr<IFileWriter> writer;
    wil::com_ptr<IFileSystemBoundObject> ownedFinal;
    hr = context.objectBinding->CreateExclusiveWriter(destination.display.c_str(), context.options, writer.put(), ownedFinal.put());
    if (FAILED(hr) || ! writer || ! ownedFinal)
    {
        return finish(FAILED(hr) ? hr : E_UNEXPECTED);
    }
    // The direct-final route creates the exact destination object itself; until content is
    // complete that object is an incomplete retained leaf, never a published copy.
    RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::NotPublished, FileSystemOwnedStageDisposition::RetainedIncomplete);

    bool finalCommitted             = false;
    const auto finishWithExactAbort = [&](HRESULT operationHr) noexcept -> HRESULT
    {
        if (finalCommitted || ! ownedFinal)
        {
            return operationHr;
        }

#if defined(ENABLE_TESTS)
        MaybeInjectDirectFinalRollbackSwapForSelfTest(destination.display);
        if (ConsumeDirectFinalAbortFailureForSelfTest(destination.display))
        {
            writer.reset();
            RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::RetainedIncomplete);
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        if (ConsumeAbortOwnedStageUnknownInjection(destination.display))
        {
            writer.reset();
            RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::Unknown);
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
#endif

        writer.reset();
        FileSystemConditionalMutationResult abortResult{};
        abortResult.sizeBytes = sizeof(abortResult);
        // The direct-final route owns one retained Local object and aborts it with one exact
        // handle mutation under the shared cleanup snapshot: primary cancellation must stop
        // content work without revoking that already-proved cleanup authority.
        const FileSystemOptions abortOptions = MakeOwnedStageCleanupOptions(context.options);
        const HRESULT abortHr                = ownedFinal->AbortOwnedObject(&abortOptions, &abortResult);
        if (abortResult.outcomeKnown == FALSE)
        {
            RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::Unknown);
            Debug::Perf::EmitCounter(L"FileOps.Local.DirectFinalAbortUnknown");
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        if (FAILED(abortHr) || abortResult.mutationCommitted == FALSE || abortResult.originalStillPresent != FALSE)
        {
            RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::RetainedIncomplete);
            Debug::Perf::EmitCounter(L"FileOps.Local.DirectFinalAbortRetainedIncomplete");
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::Removed);
        return operationHr;
    };

    FileSystemMetadataTransferResult prepareMetadata{};
    prepareMetadata.sizeBytes    = sizeof(prepareMetadata);
    prepareMetadata.firstFailure = S_OK;
    hr = sourceMetadata->TransferMetadataTo(ownedFinal.get(), FILESYSTEM_METADATA_TRANSFER_PREPARE_CONTENT, context.options, &prepareMetadata);
    if (FAILED(hr) || (prepareMetadata.lostFeatures & FILESYSTEM_METADATA_EFS) != 0u)
    {
        // The exact transfer itself failed, or an encrypted source cannot stay encrypted at the
        // destination. Like CopyFileExW, never publish plaintext for an EFS source.
        return finish(finishWithExactAbort(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_ENCRYPTION_FAILED)));
    }
    if (prepareMetadata.lostFeatures != 0u)
    {
        // Sparse state is best effort exactly as with CopyFileExW: a destination volume without
        // the feature receives plain content. The loss is recorded, never fatal. Compression is
        // never transferred; the destination inherits its folder's state like File Explorer.
        Debug::Perf::EmitValue(L"FileOps.Local.DirectFinalMetadataLost", prepareMetadata.lostFeatures, prepareMetadata.firstFailure);
    }

    auto buffer = std::unique_ptr<std::byte[]>(new (std::nothrow) std::byte[kBufferBytes]);
    if (! buffer)
    {
        return finish(finishWithExactAbort(E_OUTOFMEMORY));
    }

    uint64_t copiedBytes = 0u;
    while (copiedBytes < readerSize)
    {
        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return finish(finishWithExactAbort(hr));
        }

        const unsigned long toRead = static_cast<unsigned long>((std::min)(static_cast<uint64_t>(kBufferBytes), readerSize - copiedBytes));
        unsigned long bytesRead    = 0u;
        hr                         = reader->Read(buffer.get(), toRead, &bytesRead);
        if (FAILED(hr) || bytesRead == 0u || bytesRead > toRead)
        {
            return finish(finishWithExactAbort(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_HANDLE_EOF)));
        }

        unsigned long consumed = 0u;
        while (consumed < bytesRead)
        {
            unsigned long bytesWritten = 0u;
            hr                         = writer->Write(buffer.get() + consumed, bytesRead - consumed, &bytesWritten);
            if (FAILED(hr) || bytesWritten == 0u || bytesWritten > bytesRead - consumed)
            {
                return finish(finishWithExactAbort(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_WRITE_FAULT)));
            }
            consumed += bytesWritten;
        }

        copiedBytes += bytesRead;
        LARGE_INTEGER total{};
        LARGE_INTEGER completed{};
        total.QuadPart     = static_cast<LONGLONG>(readerSize);
        completed.QuadPart = static_cast<LONGLONG>(copiedBytes);
        if (CopyProgressRoutine(total, completed, {}, {}, 0u, 0u, nullptr, nullptr, &progress) != PROGRESS_CONTINUE)
        {
            return finish(finishWithExactAbort(HRESULT_FROM_WIN32(ERROR_CANCELLED)));
        }
#if defined(ENABLE_TESTS)
        if (ShouldFailDirectFinalCopyForSelfTest(destination.display))
        {
            return finish(finishWithExactAbort(HRESULT_FROM_WIN32(ERROR_CANCELLED)));
        }
#endif
    }

    hr = writer->Commit();
    if (FAILED(hr))
    {
        return finish(finishWithExactAbort(hr));
    }
    writer.reset();

    FileSystemMetadataTransferResult finalMetadata{};
    finalMetadata.sizeBytes    = sizeof(finalMetadata);
    finalMetadata.firstFailure = S_OK;
    hr                         = sourceMetadata->TransferMetadataTo(ownedFinal.get(), FILESYSTEM_METADATA_TRANSFER_FINALIZE, context.options, &finalMetadata);
    if (FAILED(hr))
    {
        return finish(finishWithExactAbort(hr));
    }
    if (finalMetadata.lostFeatures != 0u)
    {
        // Named streams, MOTW, extended attributes, or ordinary attributes the destination cannot
        // hold (for example FAT/exFAT) are dropped exactly as CopyFileExW drops them. The complete
        // content is published; the loss is recorded, never fatal.
        Debug::Perf::EmitValue(L"FileOps.Local.DirectFinalMetadataLost", finalMetadata.lostFeatures, finalMetadata.firstFailure);
    }

    finalCommitted = true;
    RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::Published, FileSystemOwnedStageDisposition::Published);
    return finish(S_OK);
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

    DWORD destinationAttributes   = GetFileAttributesW(destination.extended.c_str());
    const bool destinationExisted = destinationAttributes != INVALID_FILE_ATTRIBUTES;

    // An operation-wide compatibility flag is not a replacement receipt when the name is
    // currently absent. Keep the ordinary path exclusive so a destination that races into
    // existence is reported as a conflict instead of trying to bind a nonexistent object.
    const bool allowOverwriteEffective = destinationExisted && HasOverwriteGrant(context);
    if (destinationExisted)
    {
        if (! allowOverwriteEffective)
        {
            // Existing destination is always a user-visible collision. Content equality and hashes
            // may verify an explicit copy, but they never choose Skip or authorize Move cleanup.
            return returnFailure(HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
        }

        if ((destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0)
        {
            if (! HasReplaceReadonlyGrant(context))
            {
                return returnFailure(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
            }
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

    HRESULT copyHr = S_OK;
    if (allowOverwriteEffective)
    {
        copyHr = CopyFileExWithStagedOverwrite(context, source, destination, 0u, progress, fileBytes, true, false);
    }
    else
    {
        copyHr = CopyFileWithRetainedFinalObject(context, source, destination, progress, fileBytes);
    }

    if (FAILED(copyHr))
    {
        if (IsCancellationHr(copyHr))
        {
            const HRESULT hrCancelled = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            EmitSequentialThrottleSummary(context, progress, fileBytes, hrCancelled);
            return hrCancelled;
        }

        if (allowOverwriteEffective)
        {
            RollBackStagedProgress(context, progress);
        }
        const HRESULT hrFailure = returnFailure(copyHr, fileBytes, progress.lastItemBytesTransferred);
        EmitSequentialThrottleSummary(context, progress, fileBytes, hrFailure);
        return hrFailure;
    }

    *bytesCopied = fileBytes;
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

[[nodiscard]] HRESULT PublishSemanticLink(OperationContext& context, const PathInfo& source, const PathInfo& destination, DWORD sourceAttributes) noexcept
{
    if (! context.objectBinding || ! context.options)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    if (context.reparseRootSourcePath.empty() || context.reparseRootDestinationPath.empty())
    {
        return E_UNEXPECTED;
    }

    constexpr FileSystemBindFlags sourceFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    wil::com_ptr<IFileSystemBoundObject> sourceAuthority;
    HRESULT hr = context.objectBinding->BindObject(source.display.c_str(), sourceFlags, sourceAuthority.put());
    if (FAILED(hr) || ! sourceAuthority)
    {
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }

    FileSystemLinkTransform transform{};
    transform.sizeBytes           = sizeof(transform);
    transform.sourceLinkPath      = source.display.c_str();
    transform.destinationLinkPath = destination.display.c_str();
    transform.sourceRootPath      = context.reparseRootSourcePath.c_str();
    transform.destinationRootPath = context.reparseRootDestinationPath.c_str();

    FileSystemLinkInformation information{};
    information.sizeBytes = sizeof(information);
    hr                    = context.objectBinding->ReadBoundLink(sourceAuthority.get(), &transform, context.options, &information);
    if (hr != HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER))
    {
        return FAILED(hr) ? hr : E_UNEXPECTED;
    }
    if (information.targetLengthUtf16 == 0u || information.targetLengthUtf16 > 32'767u ||
        information.targetMapping != FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT || information.sourceRelativeTargetLengthUtf16 != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    std::vector<wchar_t> target(static_cast<size_t>(information.targetLengthUtf16) + 1u);
    information.targetBuffer        = target.data();
    information.targetCapacityUtf16 = static_cast<uint32_t>(target.size());
    hr                              = context.objectBinding->ReadBoundLink(sourceAuthority.get(), &transform, context.options, &information);
    if (FAILED(hr))
    {
        return hr;
    }
    const bool sourceIsDirectory = IsDirectory(sourceAttributes);
    const bool kindMatchesSource = sourceIsDirectory
                                       ? (information.kind == FILESYSTEM_LINK_KIND_SYMBOLIC_DIRECTORY || information.kind == FILESYSTEM_LINK_KIND_JUNCTION)
                                       : information.kind == FILESYSTEM_LINK_KIND_SYMBOLIC_FILE;
    if (! kindMatchesSource)
    {
        return E_UNEXPECTED;
    }

    wil::com_ptr<IFileSystemBoundObject> expectedDestination;
    const DWORD destinationAttributes = GetFileAttributesW(destination.extended.c_str());
    if (destinationAttributes != INVALID_FILE_ATTRIBUTES)
    {
        if (! IsReparsePoint(destinationAttributes) || sourceIsDirectory != IsDirectory(destinationAttributes))
        {
            return HRESULT_FROM_WIN32(ERROR_DATATYPE_MISMATCH);
        }
        if (! HasOverwriteGrant(context))
        {
            return HRESULT_FROM_WIN32(IsReparsePoint(destinationAttributes) ? ERROR_REPARSE_POINT_ENCOUNTERED : ERROR_ALREADY_EXISTS);
        }
        if ((destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0u && ! HasReplaceReadonlyGrant(context))
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        expectedDestination = context.oneShotExpectedDestination;
        if (! expectedDestination)
        {
            constexpr FileSystemBindFlags destinationFlags =
                static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_PUBLICATION);
            hr = context.objectBinding->BindObject(destination.display.c_str(), destinationFlags, expectedDestination.put());
            if (FAILED(hr) || ! expectedDestination)
            {
                return FAILED(hr) ? hr : E_UNEXPECTED;
            }
        }
        Debug::Perf::EmitCounter(L"FileOps.Conflict.ExpectedDestinationBoundCount");
    }
    else
    {
        const DWORD error = GetLastError();
        if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND)
        {
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
    }

    std::wstring stagePath;
    wil::com_ptr<IFileSystemBoundObject> ownedStage;
    constexpr unsigned int kMaxStageAttempts = 32u;
    for (unsigned int attempt = 0u; attempt < kMaxStageAttempts; ++attempt)
    {
        stagePath = MakeCopyTempSiblingPath(destination.display);
        if (stagePath.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
        }
        hr = context.objectBinding->CreateExclusiveLink(stagePath.c_str(), &information, context.options, ownedStage.put());
        if (SUCCEEDED(hr))
        {
            break;
        }
        ownedStage.reset();
        if (hr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && hr != HRESULT_FROM_WIN32(ERROR_FILE_EXISTS))
        {
            return hr;
        }
    }
    if (! ownedStage)
    {
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }
    RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::NotPublished, FileSystemOwnedStageDisposition::Retained);

    bool publicationUnknown                = false;
    bool published                         = false;
    const auto finishWithOwnedStageCleanup = [&](HRESULT operationHr) noexcept -> HRESULT
    {
        if (published || publicationUnknown || ! ownedStage)
        {
            return operationHr;
        }
#if defined(ENABLE_TESTS)
        if (ConsumeAbortOwnedStageUnknownInjection(destination.display))
        {
            RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::Unknown);
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
#endif
        FileSystemConditionalMutationResult abortResult{};
        abortResult.sizeBytes                  = sizeof(abortResult);
        const FileSystemOptions cleanupOptions = MakeOwnedStageCleanupOptions(context.options);
        const HRESULT abortHr                  = ownedStage->AbortOwnedObject(&cleanupOptions, &abortResult);
        if (abortResult.outcomeKnown == FALSE || FAILED(abortHr) || abortResult.mutationCommitted == FALSE || abortResult.originalStillPresent != FALSE)
        {
            RecordOwnedStageDisposition(
                context, destination, abortResult.outcomeKnown == FALSE ? FileSystemOwnedStageDisposition::Unknown : FileSystemOwnedStageDisposition::Retained);
            Debug::Perf::EmitCounter(L"fileops.local.link_stage_cleanup_unknown");
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::Removed);
        return operationHr;
    };

    FileSystemBasicInformation sourceInfo{};
    sourceInfo.sizeBytes = sizeof(sourceInfo);
    hr                   = sourceAuthority->GetBasicInformation(&sourceInfo);
    if (SUCCEEDED(hr))
    {
        hr = ownedStage->SetBasicInformation(&sourceInfo);
    }
    if (FAILED(hr) && hr != E_NOTIMPL && hr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
    {
        return finishWithOwnedStageCleanup(hr);
    }

    FileSystemFlags publicationFlags = FILESYSTEM_FLAG_NONE;
    if (expectedDestination)
    {
        publicationFlags = FILESYSTEM_FLAG_ALLOW_OVERWRITE;
        if (HasReplaceReadonlyGrant(context))
        {
            publicationFlags = static_cast<FileSystemFlags>(publicationFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        }
        if (HasReplaceLinkGrant(context))
        {
            publicationFlags = static_cast<FileSystemFlags>(publicationFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_LINK);
        }
        Debug::Perf::EmitCounter(L"FileOps.Link.ReplaceConditionalCount");
    }

    FileSystemConditionalMutationResult mutationResult{};
    mutationResult.sizeBytes = sizeof(mutationResult);
    wil::com_ptr<IFileSystemBoundObject> publishedAuthority;
    hr = ownedStage->PublishAs(
        destination.display.c_str(), expectedDestination.get(), publicationFlags, context.options, &mutationResult, publishedAuthority.put());
    if (mutationResult.outcomeKnown == FALSE)
    {
        publicationUnknown = true;
        RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::Unknown, FileSystemOwnedStageDisposition::Unknown);
        return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
    }
    if (mutationResult.mutationCommitted == FALSE)
    {
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_INVALID) || hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            Debug::Perf::EmitCounter(L"FileOps.Conflict.ExpectedDestinationMismatchCount");
        }
        return finishWithOwnedStageCleanup(FAILED(hr) ? hr : E_UNEXPECTED);
    }
    published = true;
    RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::Published, FileSystemOwnedStageDisposition::Published);
    if (! publishedAuthority || mutationResult.originalStillPresent != FALSE)
    {
        return E_UNEXPECTED;
    }
    return hr;
}

HRESULT
PreserveReparsePointInternal(
    OperationContext& context, const PathInfo& source, const PathInfo& destination, DWORD sourceAttributes, uint64_t* bytesCopied) noexcept
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

    // Literal Preserve: rebuild the link from its no-follow payload with the stored target text.
    // The exclusive-stage route covers file symlinks, directory symlinks, and junctions alike.
    return returnFailure(PublishSemanticLink(context, source, destination, sourceAttributes));
}

[[nodiscard]] FileSystemDiscoveryMode GetDiscoveryMode(OperationContext& context) noexcept
{
    if (context.options == nullptr || context.options->operationControl == nullptr)
    {
        return FILESYSTEM_DISCOVERY_AHEAD;
    }

    FileSystemDiscoveryMode mode = FILESYSTEM_DISCOVERY_AHEAD;
    const HRESULT hr             = context.options->operationControl->FileSystemGetDiscoveryMode(&mode, context.options->operationControlCookie);
    if (FAILED(hr) || (mode != FILESYSTEM_DISCOVERY_AHEAD && mode != FILESYSTEM_DISCOVERY_JUST_IN_TIME))
    {
        return FILESYSTEM_DISCOVERY_JUST_IN_TIME;
    }
    return mode;
}

HRESULT ReportDiscoveryProgress(OperationContext& context,
                                uint64_t discoveredBytes,
                                uint64_t discoveredFiles,
                                uint64_t discoveredDirectories,
                                uint32_t queuedItems,
                                bool traversalClosed) noexcept
{
    if (context.options == nullptr || context.options->operationControl == nullptr)
    {
        return S_OK;
    }

    FileSystemDiscoveryProgress progress{};
    progress.sizeBytes             = sizeof(progress);
    progress.discoveredBytes       = discoveredBytes;
    progress.discoveredFiles       = discoveredFiles;
    progress.discoveredDirectories = discoveredDirectories;
    progress.queuedItems           = queuedItems;
    progress.traversalClosed       = traversalClosed ? TRUE : FALSE;
    return context.options->operationControl->FileSystemReportDiscoveryProgress(&progress, context.options->operationControlCookie);
}

[[nodiscard]] HRESULT ReportTopLevelDiscovery(OperationContext& context, const std::wstring& sourcePath, bool recursiveDirectoryOwnsDiscovery) noexcept
{
    if (context.options == nullptr || context.options->operationControl == nullptr)
    {
        return S_OK;
    }

    WIN32_FILE_ATTRIBUTE_DATA data{};
    if (! ::GetFileAttributesExW(sourcePath.c_str(), GetFileExInfoStandard, &data))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const bool directory          = IsDirectory(data.dwFileAttributes);
    const bool recursiveDirectory = recursiveDirectoryOwnsDiscovery && directory && ! IsReparsePoint(data.dwFileAttributes);
    if (recursiveDirectory)
    {
        // The recursive copy walker owns the root and descendant cumulative totals.
        return S_OK;
    }

    const uint64_t fileBytes = directory ? 0u : (static_cast<uint64_t>(data.nFileSizeHigh) << 32u) | data.nFileSizeLow;
    return ReportDiscoveryProgress(context, fileBytes, directory ? 0u : 1u, directory ? 1u : 0u, 1u, true);
}

[[nodiscard]] HRESULT ReportDeleteDiscoveryObject(OperationContext& context, DWORD attributes, uint64_t sizeBytes, uint32_t queuedItems) noexcept
{
    DeleteDiscoveryState* const discovery = context.deleteDiscovery;
    if (discovery == nullptr)
    {
        return S_OK;
    }

    std::scoped_lock lock(discovery->mutex);
    if (IsDirectory(attributes))
    {
        if (discovery->discoveredDirectories != std::numeric_limits<uint64_t>::max())
        {
            ++discovery->discoveredDirectories;
        }
    }
    else
    {
        if (discovery->discoveredFiles != std::numeric_limits<uint64_t>::max())
        {
            ++discovery->discoveredFiles;
        }
        discovery->discoveredBytes = discovery->discoveredBytes > std::numeric_limits<uint64_t>::max() - sizeBytes ? std::numeric_limits<uint64_t>::max()
                                                                                                                   : discovery->discoveredBytes + sizeBytes;
    }

    return ReportDiscoveryProgress(context, discovery->discoveredBytes, discovery->discoveredFiles, discovery->discoveredDirectories, queuedItems, false);
}

[[nodiscard]] uint64_t AddRecursiveCopyRetainedBytes(uint64_t left, uint64_t right) noexcept
{
    return left > std::numeric_limits<uint64_t>::max() - right ? std::numeric_limits<uint64_t>::max() : left + right;
}

[[nodiscard]] uint64_t RecursiveCopyRetainedPathBytes(const PathInfo& source, const PathInfo& destination) noexcept
{
    const uint64_t sourceChars = AddRecursiveCopyRetainedBytes(static_cast<uint64_t>(source.display.size()), static_cast<uint64_t>(source.extended.size()));
    const uint64_t destinationChars =
        AddRecursiveCopyRetainedBytes(static_cast<uint64_t>(destination.display.size()), static_cast<uint64_t>(destination.extended.size()));
    const uint64_t totalChars = AddRecursiveCopyRetainedBytes(sourceChars, destinationChars);
    if (totalChars > std::numeric_limits<uint64_t>::max() / sizeof(wchar_t))
    {
        return std::numeric_limits<uint64_t>::max();
    }
    return AddRecursiveCopyRetainedBytes(totalChars * sizeof(wchar_t), Common::FileOperations::kTraversalRecordOverheadBytes);
}

struct RecursiveCopySerialBudget final
{
    uint64_t retainedAncestorMetadataBytes    = 0;
    uint64_t maxRetainedAncestorMetadataBytes = 0;
    uint64_t maxDepth                         = 0;
    uint64_t discoveredBytes                  = 0;
    uint64_t discoveredFiles                  = 0;
    uint64_t discoveredDirectories            = 0;
};

[[nodiscard]] HRESULT PublishLocalCopyDirectory(OperationContext& context,
                                                const PathInfo& destination,
                                                IFileSystemBoundObject* expectedDestination,
                                                wil::com_ptr<IFileSystemBoundObject>& publishedAuthority) noexcept
{
    publishedAuthority.reset();
    if (! context.objectBinding || ! context.options)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    std::wstring stagePath;
    wil::com_ptr<IFileSystemBoundObject> ownedStage;
    constexpr unsigned int kMaximumStageAttempts = 32u;
    HRESULT hr                                   = E_UNEXPECTED;
    for (unsigned int attempt = 0u; attempt < kMaximumStageAttempts; ++attempt)
    {
        stagePath = MakeCopyTempSiblingPath(destination.display);
        if (stagePath.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
        }
        hr = context.objectBinding->CreateExclusiveDirectory(stagePath.c_str(), context.options, ownedStage.put());
        if (SUCCEEDED(hr))
        {
            break;
        }
        ownedStage.reset();
        if (hr != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) && hr != HRESULT_FROM_WIN32(ERROR_FILE_EXISTS))
        {
            return hr;
        }
    }
    if (! ownedStage)
    {
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }
    RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::NotPublished, FileSystemOwnedStageDisposition::Retained);

    bool publicationUnknown                = false;
    bool published                         = false;
    const auto finishWithOwnedStageCleanup = [&](HRESULT operationHr) noexcept -> HRESULT
    {
        if (published || publicationUnknown || ! ownedStage)
        {
            return operationHr;
        }
#if defined(ENABLE_TESTS)
        if (ConsumeAbortOwnedStageUnknownInjection(destination.display))
        {
            RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::Unknown);
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
#endif
        FileSystemConditionalMutationResult abortResult{};
        abortResult.sizeBytes                  = sizeof(abortResult);
        const FileSystemOptions cleanupOptions = MakeOwnedStageCleanupOptions(context.options);
        const HRESULT abortHr                  = ownedStage->AbortOwnedObject(&cleanupOptions, &abortResult);
        if (abortResult.outcomeKnown == FALSE || FAILED(abortHr) || abortResult.mutationCommitted == FALSE || abortResult.originalStillPresent != FALSE)
        {
            RecordOwnedStageDisposition(
                context, destination, abortResult.outcomeKnown == FALSE ? FileSystemOwnedStageDisposition::Unknown : FileSystemOwnedStageDisposition::Retained);
            Debug::Perf::EmitCounter(L"fileops.local.directory_stage_cleanup_unknown");
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        RecordOwnedStageDisposition(context, destination, FileSystemOwnedStageDisposition::Removed);
        return operationHr;
    };

    FileSystemFlags publicationFlags = FILESYSTEM_FLAG_NONE;
    if (expectedDestination != nullptr)
    {
        publicationFlags = FILESYSTEM_FLAG_ALLOW_OVERWRITE;
        if (HasReplaceReadonlyGrant(context))
        {
            publicationFlags = static_cast<FileSystemFlags>(publicationFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        }
        if (HasReplaceLinkGrant(context))
        {
            publicationFlags = static_cast<FileSystemFlags>(publicationFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_LINK);
        }
    }

    FileSystemConditionalMutationResult result{};
    result.sizeBytes = sizeof(result);
    hr = ownedStage->PublishAs(destination.display.c_str(), expectedDestination, publicationFlags, context.options, &result, publishedAuthority.put());
    if (result.outcomeKnown == FALSE)
    {
        publicationUnknown = true;
        RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::Unknown, FileSystemOwnedStageDisposition::Unknown);
        return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
    }
    if (result.mutationCommitted == FALSE)
    {
        return finishWithOwnedStageCleanup(FAILED(hr) ? hr : E_UNEXPECTED);
    }
    published = true;
    RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::Published, FileSystemOwnedStageDisposition::Published);
    if (! publishedAuthority || result.originalStillPresent != FALSE)
    {
        return E_UNEXPECTED;
    }
    return hr;
}

[[nodiscard]] HRESULT EnsureLocalCopyDestinationDirectory(OperationContext& context,
                                                          const PathInfo& destination,
                                                          bool& createdDestination,
                                                          wil::com_ptr<IFileSystemBoundObject>& createdAuthority) noexcept
{
    createdDestination = false;
    createdAuthority.reset();
    const DWORD attributes = GetFileAttributesW(destination.extended.c_str());
    if (attributes != INVALID_FILE_ATTRIBUTES)
    {
        LocalCopyPathKind destinationKind = LocalCopyPathKind::RegularFile;
        const HRESULT kindHr              = ClassifyLocalCopyPathKind(destination.extended, attributes, destinationKind);
        if (FAILED(kindHr))
        {
            return kindHr;
        }
        if (destinationKind == LocalCopyPathKind::RegularDirectory)
        {
            return S_OK;
        }
        if (destinationKind != LocalCopyPathKind::SemanticLink)
        {
            return HRESULT_FROM_WIN32(ERROR_DATATYPE_MISMATCH);
        }
        if (! HasOverwriteGrant(context) || ! HasReplaceLinkGrant(context))
        {
            return HRESULT_FROM_WIN32(ERROR_REPARSE_POINT_ENCOUNTERED);
        }

        wil::com_ptr<IFileSystemBoundObject> expectedDestination = context.oneShotExpectedDestination;
        if (! expectedDestination)
        {
            constexpr FileSystemBindFlags bindFlags =
                static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_PUBLICATION);
            const HRESULT bindHr = context.objectBinding ? context.objectBinding->BindObject(destination.display.c_str(), bindFlags, expectedDestination.put())
                                                         : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            if (FAILED(bindHr) || ! expectedDestination)
            {
                return FAILED(bindHr) ? bindHr : E_UNEXPECTED;
            }
        }
        const HRESULT publishHr = PublishLocalCopyDirectory(context, destination, expectedDestination.get(), createdAuthority);
        if (SUCCEEDED(publishHr))
        {
            createdDestination = true;
            Debug::Perf::EmitCounter(L"FileOps.Link.ReplaceConditionalCount");
        }
        return publishHr;
    }

    const DWORD error = GetLastError();
    if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND)
    {
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    const HRESULT publishHr = PublishLocalCopyDirectory(context, destination, nullptr, createdAuthority);
    if (SUCCEEDED(publishHr))
    {
        createdDestination = true;
    }
    return publishHr;
}

void CopyLocalDirectoryBasicInformationBestEffort(OperationContext& context, const PathInfo& source, IFileSystemBoundObject* destinationAuthority) noexcept
{
    if (! context.objectBinding || destinationAuthority == nullptr)
    {
        return;
    }
    constexpr FileSystemBindFlags sourceFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    wil::com_ptr<IFileSystemBoundObject> sourceAuthority;
    if (FAILED(context.objectBinding->BindObject(source.display.c_str(), sourceFlags, sourceAuthority.put())) || ! sourceAuthority)
    {
        return;
    }
    FileSystemBasicInformation info{};
    info.sizeBytes = sizeof(info);
    if (SUCCEEDED(sourceAuthority->GetBasicInformation(&info)))
    {
        static_cast<void>(destinationAuthority->SetBasicInformation(&info));
    }
}

HRESULT CopyDirectoryInternal(OperationContext& context,
                              const PathInfo& source,
                              const PathInfo& destination,
                              uint64_t* bytesCopied,
                              uint64_t depth                          = 0u,
                              RecursiveCopySerialBudget* sharedBudget = nullptr) noexcept
{
    if (! bytesCopied)
    {
        return E_POINTER;
    }

    *bytesCopied = 0;

    RecursiveCopySerialBudget localBudget{};
    const bool ownsBudget             = sharedBudget == nullptr;
    RecursiveCopySerialBudget& budget = sharedBudget != nullptr ? *sharedBudget : localBudget;
    const auto emitBudgetOnExit       = wil::scope_exit([&]() noexcept
    {
        if (ownsBudget)
        {
            static_cast<void>(ReportDiscoveryProgress(context, budget.discoveredBytes, budget.discoveredFiles, budget.discoveredDirectories, 0u, true));
            Debug::Perf::Emit(L"FileOps.CopyRecursiveSerial.MaxTraversalDepth", L"", 0u, budget.maxDepth, Common::FileOperations::kTraversalMaxDepth, S_OK);
            Debug::Perf::Emit(L"FileOps.CopyRecursiveSerial.MaxAncestorMetadataBytes",
                              L"",
                              0u,
                              budget.maxRetainedAncestorMetadataBytes,
                              Common::FileOperations::kTraversalMaxMetadataBytes,
                              S_OK);
        }
    });
    ++budget.discoveredDirectories;
    budget.maxDepth = (std::max)(budget.maxDepth, depth);
    if (depth > Common::FileOperations::kTraversalMaxDepth)
    {
        Debug::Warning(L"FileSystem: serial recursive copy stopped at '{}' after reaching the walk-depth limit ({}).",
                       source.display,
                       Common::FileOperations::kTraversalMaxDepth);
        return HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW);
    }
    const uint64_t ancestorBytes = AddRecursiveCopyRetainedBytes(RecursiveCopyRetainedPathBytes(source, destination), sizeof(WIN32_FIND_DATAW));
    if (ancestorBytes > Common::FileOperations::kTraversalMaxMetadataBytes ||
        budget.retainedAncestorMetadataBytes > Common::FileOperations::kTraversalMaxMetadataBytes - ancestorBytes)
    {
        Debug::Warning(L"FileSystem: serial recursive copy stopped at '{}' after reaching the ancestor-metadata limit ({} bytes).",
                       source.display,
                       Common::FileOperations::kTraversalMaxMetadataBytes);
        return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
    }
    budget.retainedAncestorMetadataBytes += ancestorBytes;
    budget.maxRetainedAncestorMetadataBytes = (std::max)(budget.maxRetainedAncestorMetadataBytes, budget.retainedAncestorMetadataBytes);
    const auto releaseAncestorMetadata      = wil::scope_exit([&]() noexcept { budget.retainedAncestorMetadataBytes -= ancestorBytes; });

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

    bool createdDestination = false;
    wil::com_ptr<IFileSystemBoundObject> createdDestinationAuthority;
    hr = EnsureLocalCopyDestinationDirectory(context, destination, createdDestination, createdDestinationAuthority);
    if (FAILED(hr))
    {
        return returnFailure(hr);
    }

    // A grant that authorized this directory's own destination conflict must not leak into
    // the subtree: every child conflict below prompts on its own.
    ClearOneShotGrants(context);

    bool directoryMetadataRestored      = false;
    const auto restoreDirectoryMetadata = [&]() noexcept
    {
        if (createdDestination && ! directoryMetadataRestored)
        {
            CopyLocalDirectoryBasicInformationBestEffort(context, source, createdDestinationAuthority.get());
            directoryMetadataRestored = true;
        }
    };
    const auto restoreDirectoryMetadataOnExit = wil::scope_exit([&]() noexcept { restoreDirectoryMetadata(); });

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
                restoreDirectoryMetadata();
            }
            return S_OK;
        }
        return returnFailure(HRESULT_FROM_WIN32(error));
    }

    bool hadFailure           = false;
    bool hadSkipped           = false;
    bool hadIdenticalFileSkip = false;

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
        childDestination.display          = AppendPath(destination.display, data.cFileName);
        childDestination.extended         = AppendPath(destination.extended, data.cFileName);
        PathInfo selectedChildDestination = childDestination;

        uint64_t childBytes = 0;
        HRESULT childHr     = S_OK;

        const DWORD childAttributes       = data.dwFileAttributes;
        const bool childHasDirectoryShape = IsDirectory(childAttributes);
        LocalCopyPathKind childKind       = LocalCopyPathKind::RegularFile;
        hr                                = ClassifyLocalCopyPathKind(childSource.extended, childAttributes, childKind);
        if (FAILED(hr))
        {
            return hr;
        }
        if (childHasDirectoryShape)
        {
            if (childKind == LocalCopyPathKind::SemanticLink)
            {
                ++budget.discoveredDirectories;
            }
        }
        else
        {
            ++budget.discoveredFiles;
            const uint64_t fileBytes = (static_cast<uint64_t>(data.nFileSizeHigh) << 32u) | data.nFileSizeLow;
            budget.discoveredBytes   = AddRecursiveCopyRetainedBytes(budget.discoveredBytes, fileBytes);
        }
        hr = ReportDiscoveryProgress(context, budget.discoveredBytes, budget.discoveredFiles, budget.discoveredDirectories, 0u, false);
        if (FAILED(hr))
        {
            return hr;
        }

        for (;;)
        {
            childBytes = 0;
            childHr    = S_OK;

            if (childKind == LocalCopyPathKind::SemanticLink)
            {
                if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
                {
                    hadSkipped = true;
                    childHr    = S_OK;
                }
                else
                {
                    childHr = PreserveReparsePointInternal(context, childSource, selectedChildDestination, childAttributes, &childBytes);
                }
            }
            else if (childKind == LocalCopyPathKind::RegularDirectory)
            {
                if (! context.recursive)
                {
                    childHr = HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
                }
                else
                {
                    childHr = CopyDirectoryInternal(context, childSource, selectedChildDestination, &childBytes, depth + 1u, &budget);
                }
            }
            else
            {
                childHr = CopyFileInternal(context, childSource, selectedChildDestination, &childBytes);
            }

            if (childHr == S_FALSE)
            {
                hadIdenticalFileSkip = true;
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

            // Traversal ceilings are task-terminal. Continue-on-error is for item failures;
            // it must not turn a hard resource/depth bound into continued discovery.
            if (childHr == HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW) || childHr == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY))
            {
                return childHr;
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
                case FileSystemIssueAction::Overwrite: context.oneShotAllowOverwrite = true; continue;
                case FileSystemIssueAction::ReplaceLink:
                    context.oneShotAllowOverwrite   = true;
                    context.oneShotAllowReplaceLink = true;
                    continue;
                case FileSystemIssueAction::ReplaceReadOnly:
                    context.oneShotAllowOverwrite       = true;
                    context.oneShotAllowReplaceReadonly = true;
                    continue;
                case FileSystemIssueAction::PermanentDelete:
                case FileSystemIssueAction::Cancel: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                case FileSystemIssueAction::Retry: continue;
                case FileSystemIssueAction::KeepBoth:
                {
                    const HRESULT keepBothHr = SelectKeepBothPath(childDestination, childHasDirectoryShape, selectedChildDestination);
                    if (FAILED(keepBothHr))
                    {
                        return keepBothHr;
                    }
                    continue;
                }
                case FileSystemIssueAction::Skip:
                    hadFailure = true;
                    childHr    = S_OK;
                    break;
                case FileSystemIssueAction::None:
                default: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            break;
        }

        ClearOneShotGrants(context);

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
        restoreDirectoryMetadata();
    }

    if (hadFailure || hadSkipped)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return hadIdenticalFileSkip ? S_FALSE : S_OK;
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

    LocalCopyPathKind sourceKind = LocalCopyPathKind::RegularFile;
    const HRESULT kindHr         = ClassifyLocalCopyPathKind(source.extended, attributes, sourceKind);
    if (FAILED(kindHr))
    {
        return kindHr;
    }
    if (sourceKind == LocalCopyPathKind::SemanticLink)
    {
        if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return PreserveReparsePointInternal(context, source, destination, attributes, bytesCopied);
    }

    if (sourceKind == LocalCopyPathKind::RegularDirectory)
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
    File,
    ReparsePoint,
    Finish,
};

struct RecursiveCopyWorkItem
{
    RecursiveCopyWorkKind kind = RecursiveCopyWorkKind::File;
    PathInfo source;
    PathInfo destination;
    DWORD attributes           = 0;
    uint64_t retainedPathBytes = 0;
};

[[nodiscard]] uint64_t RecursiveCopyRetainedPathBytes(const RecursiveCopyWorkItem& item) noexcept
{
    return RecursiveCopyRetainedPathBytes(item.source, item.destination);
}

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
    LocalCopyPathKind sourceKind = LocalCopyPathKind::RegularFile;
    const HRESULT kindHr         = ClassifyLocalCopyPathKind(source.extended, attributes, sourceKind);
    if (FAILED(kindHr))
    {
        return returnFailure(kindHr);
    }
    if (sourceKind != LocalCopyPathKind::RegularDirectory)
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
    std::atomic<bool> hadIdenticalFileSkip{false};

    std::atomic<uint64_t> queuedFiles{0};
    std::atomic<uint64_t> queuedDirectories{0};
    std::atomic<uint64_t> queuedReparsePoints{0};
    std::atomic<uint64_t> processedFiles{0};
    std::atomic<uint64_t> processedDirectories{0};
    uint64_t discoveredBytes       = 0u;
    uint64_t discoveredFiles       = 0u;
    uint64_t discoveredDirectories = 1u;
    // Count of discrete scheduler dispatches that ran a queue item. With the dynamic
    // (per-item) job model each item is one short dispatch that returns the worker to the
    // pool, so this far exceeds `concurrency`; the old long-lived consumer loop would have
    // recorded exactly `concurrency`. Surfaced as a diagnostic to prove workers are no
    // longer pinned for the whole operation.
    std::atomic<uint64_t> workItemDispatches{0};

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
        size_t activeItems            = 0;
        size_t retainedEntries        = 0;
        uint64_t retainedPathBytes    = 0;
        uint64_t maxRetainedEntries   = 0;
        uint64_t maxRetainedPathBytes = 0;
        bool producerDone             = false;
        bool done                     = false;
        // Count of items immediately available in `items`, mirrored as an atomic so the
        // scheduler can gate dynamic-job dispatch without taking `mutex` (avoiding a lock-
        // order inversion against the scheduler lock). Kept in sync under `mutex`.
        std::atomic<size_t> ready{0};
    };

    RecursiveCopyQueue queue{};
    uint64_t maxTraversalDepth                = 0;
    uint64_t retainedAncestorMetadataBytes    = 0;
    uint64_t maxRetainedAncestorMetadataBytes = 0;

    const auto initializeChildContext = [&](OperationContext& context, uint64_t progressStreamId) noexcept
    {
        InitializeOperationContext(
            context, rootContext.type, flags, sharedOptionsState, rootContext.callback, rootContext.callbackCookie, 1, reparsePointPolicy);
        context.options                    = sharedOptionsState;
        context.parallel                   = &parallel;
        context.totalBytes                 = 0; // discovery totals are reported by this same traversal
        context.progressStreamId           = progressStreamId;
        context.reparseRootSourcePath      = rootSource;
        context.reparseRootDestinationPath = rootDestination;
        context.objectBinding              = rootContext.objectBinding;
    };

    const auto finishActiveItem = [&](uint64_t retainedPathBytes) noexcept
    {
        {
            std::scoped_lock lock(queue.mutex);
            if (queue.activeItems > 0)
            {
                --queue.activeItems;
            }
            if (queue.retainedEntries > 0u)
            {
                --queue.retainedEntries;
            }
            queue.retainedPathBytes = (queue.retainedPathBytes >= retainedPathBytes) ? (queue.retainedPathBytes - retainedPathBytes) : 0u;
            if (queue.producerDone && queue.items.empty() && queue.activeItems == 0)
            {
                queue.done = true;
            }
        }
        queue.cv.notify_all();
        GetSharedFileOpsJobScheduler().NotifyDynamicWorkAvailable();
    };

    const auto enqueueWork = [&](RecursiveCopyWorkItem item) noexcept -> HRESULT
    {
        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        item.retainedPathBytes = RecursiveCopyRetainedPathBytes(item);
        if (item.retainedPathBytes > Common::FileOperations::kTraversalMaxQueuedPathBytes)
        {
            Debug::Warning(L"FileSystem: recursive copy stopped because one queued item requires {} retained path bytes (limit={}).",
                           item.retainedPathBytes,
                           Common::FileOperations::kTraversalMaxQueuedPathBytes);
            return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
        }

        {
            std::unique_lock lock(queue.mutex);
            const auto currentQueueTarget = [&]() noexcept
            {
                return Common::FileOperations::DiscoveryQueueTarget(static_cast<size_t>(concurrency),
                                                                    GetDiscoveryMode(rootContext) == FILESYSTEM_DISCOVERY_AHEAD);
            };
            while (! queue.done && ! parallel.cancelRequested.load(std::memory_order_acquire) &&
                   ! parallel.stopOnErrorRequested.load(std::memory_order_acquire) &&
                   (queue.retainedEntries >= currentQueueTarget() || queue.retainedEntries >= Common::FileOperations::kTraversalMaxQueuedEntries ||
                    queue.retainedPathBytes > Common::FileOperations::kTraversalMaxQueuedPathBytes - item.retainedPathBytes))
            {
                queue.cv.wait_for(lock, std::chrono::milliseconds(50));
            }
            if (queue.done || parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            switch (item.kind)
            {
                case RecursiveCopyWorkKind::File: queuedFiles.fetch_add(1, std::memory_order_relaxed); break;
                case RecursiveCopyWorkKind::ReparsePoint: queuedReparsePoints.fetch_add(1, std::memory_order_relaxed); break;
                case RecursiveCopyWorkKind::Finish: break;
            }

            queue.items.push_back(std::move(item));
            ++queue.retainedEntries;
            queue.retainedPathBytes += queue.items.back().retainedPathBytes;
            queue.maxRetainedEntries   = (std::max)(queue.maxRetainedEntries, static_cast<uint64_t>(queue.retainedEntries));
            queue.maxRetainedPathBytes = (std::max)(queue.maxRetainedPathBytes, queue.retainedPathBytes);
            queue.ready.fetch_add(1, std::memory_order_release);
        }

        queue.cv.notify_one();
        GetSharedFileOpsJobScheduler().NotifyDynamicWorkAvailable();
        return S_OK;
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

    const auto waitForQueueDrain = [&]() noexcept -> HRESULT
    {
        std::unique_lock lock(queue.mutex);
        while ((! queue.items.empty() || queue.activeItems != 0u) && ! parallel.cancelRequested.load(std::memory_order_acquire) &&
               ! parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            queue.cv.wait_for(lock, std::chrono::milliseconds(50));
        }
        if (parallel.cancelRequested.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        const HRESULT failure = parallel.firstError.load(std::memory_order_acquire);
        return FAILED(failure) ? failure : S_OK;
    };

    const auto waitForActiveItemsQuiescence = [&]() noexcept
    {
        std::unique_lock lock(queue.mutex);
        while (queue.activeItems != 0u)
        {
            queue.cv.wait_for(lock, std::chrono::milliseconds(50));
        }
    };

    // Keep traversal-owned strings, find data, and COM authority off the recursive
    // producer's stack. The parallel path deliberately supports deep trees up to the
    // shared traversal ceiling; retaining these objects in every recursion frame can
    // exhaust the default Windows worker stack well before that ceiling is reached.
    struct RecursiveCopyDirectoryFrame final
    {
        RecursiveCopyDirectoryFrame()                                              = default;
        RecursiveCopyDirectoryFrame(const RecursiveCopyDirectoryFrame&)            = delete;
        RecursiveCopyDirectoryFrame& operator=(const RecursiveCopyDirectoryFrame&) = delete;
        RecursiveCopyDirectoryFrame(RecursiveCopyDirectoryFrame&&)                 = delete;
        RecursiveCopyDirectoryFrame& operator=(RecursiveCopyDirectoryFrame&&)      = delete;

        uint64_t ancestorBytes  = 0;
        bool createdDestination = false;
        wil::com_ptr<IFileSystemBoundObject> createdDestinationAuthority;
        bool directoryMetadataRestored = false;
        std::wstring searchPattern;
        WIN32_FIND_DATAW findData{};
        wil::unique_hfind findHandle;
        uint64_t cancelCheckCounter = 0;
        RecursiveCopyWorkItem childItem;
    };

    std::function<HRESULT(OperationContext&, const RecursiveCopyWorkItem&, uint64_t)> processDirectory;
    std::function<HRESULT(OperationContext&, const RecursiveCopyWorkItem&, uint64_t)> runDirectory;
    processDirectory = [&](OperationContext& context, const RecursiveCopyWorkItem& item, uint64_t depth) noexcept -> HRESULT
    {
        auto frame        = std::make_unique<RecursiveCopyDirectoryFrame>();
        maxTraversalDepth = (std::max)(maxTraversalDepth, depth);
        if (depth > Common::FileOperations::kTraversalMaxDepth)
        {
            Debug::Warning(L"FileSystem: recursive copy stopped at '{}' after reaching the walk-depth limit ({}).",
                           item.source.display,
                           Common::FileOperations::kTraversalMaxDepth);
            return HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW);
        }

        frame->ancestorBytes = AddRecursiveCopyRetainedBytes(RecursiveCopyRetainedPathBytes(item), sizeof(WIN32_FIND_DATAW));
        if (frame->ancestorBytes > Common::FileOperations::kTraversalMaxMetadataBytes ||
            retainedAncestorMetadataBytes > Common::FileOperations::kTraversalMaxMetadataBytes - frame->ancestorBytes)
        {
            Debug::Warning(L"FileSystem: recursive copy stopped at '{}' after reaching the ancestor-metadata limit ({} bytes).",
                           item.source.display,
                           Common::FileOperations::kTraversalMaxMetadataBytes);
            return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
        }
        retainedAncestorMetadataBytes += frame->ancestorBytes;
        maxRetainedAncestorMetadataBytes   = (std::max)(maxRetainedAncestorMetadataBytes, retainedAncestorMetadataBytes);
        const auto releaseAncestorMetadata = wil::scope_exit([&]() noexcept { retainedAncestorMetadataBytes -= frame->ancestorBytes; });

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

        directoryHr = EnsureLocalCopyDestinationDirectory(context, item.destination, frame->createdDestination, frame->createdDestinationAuthority);
        if (FAILED(directoryHr))
        {
            return directoryHr;
        }

        // A grant that authorized this directory's own destination conflict must not leak into
        // children enqueued below; they conflict and prompt on their own.
        ClearOneShotGrants(context);

        const auto restoreDirectoryMetadata = [&]() noexcept
        {
            if (frame->createdDestination && ! frame->directoryMetadataRestored)
            {
                // A fatal producer/worker result may leave queued work abandoned. Active
                // children still quiesce before metadata is applied to the partial directory.
                waitForActiveItemsQuiescence();
                CopyLocalDirectoryBasicInformationBestEffort(context, item.source, frame->createdDestinationAuthority.get());
                frame->directoryMetadataRestored = true;
            }
        };
        const auto restoreDirectoryMetadataOnExit = wil::scope_exit([&]() noexcept { restoreDirectoryMetadata(); });

        frame->searchPattern = AppendPath(item.source.extended, L"*");
        frame->findHandle.reset(
            FindFirstFileExW(frame->searchPattern.c_str(), FindExInfoBasic, &frame->findData, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
        if (! frame->findHandle)
        {
            const DWORD error = GetLastError();
            if (error == ERROR_FILE_NOT_FOUND)
            {
                directoryHr = waitForQueueDrain();
                if (FAILED(directoryHr))
                {
                    return directoryHr;
                }
                if (frame->createdDestination)
                {
                    restoreDirectoryMetadata();
                }
                processedDirectories.fetch_add(1, std::memory_order_relaxed);
                return S_OK;
            }
            return HRESULT_FROM_WIN32(error);
        }

        do
        {
            if (IsDotOrDotDot(frame->findData.cFileName))
            {
                continue;
            }

            frame->childItem                      = RecursiveCopyWorkItem{};
            frame->childItem.source.display       = AppendPath(item.source.display, frame->findData.cFileName);
            frame->childItem.source.extended      = AppendPath(item.source.extended, frame->findData.cFileName);
            frame->childItem.destination.display  = AppendPath(item.destination.display, frame->findData.cFileName);
            frame->childItem.destination.extended = AppendPath(item.destination.extended, frame->findData.cFileName);

            const DWORD childAttributes       = frame->findData.dwFileAttributes;
            const bool childHasDirectoryShape = IsDirectory(childAttributes);
            LocalCopyPathKind childKind       = LocalCopyPathKind::RegularFile;
            directoryHr                       = ClassifyLocalCopyPathKind(frame->childItem.source.extended, childAttributes, childKind);
            if (FAILED(directoryHr))
            {
                return directoryHr;
            }

            if (childHasDirectoryShape)
            {
                ++discoveredDirectories;
            }
            else
            {
                ++discoveredFiles;
                const uint64_t fileBytes = (static_cast<uint64_t>(frame->findData.nFileSizeHigh) << 32u) | frame->findData.nFileSizeLow;
                discoveredBytes          = AddRecursiveCopyRetainedBytes(discoveredBytes, fileBytes);
            }
            uint32_t queuedItemCount = 0u;
            {
                std::scoped_lock lock(queue.mutex);
                queuedItemCount = static_cast<uint32_t>((std::min)(queue.retainedEntries, static_cast<size_t>(std::numeric_limits<uint32_t>::max())));
            }
            directoryHr = ReportDiscoveryProgress(context, discoveredBytes, discoveredFiles, discoveredDirectories, queuedItemCount, false);
            if (FAILED(directoryHr))
            {
                return directoryHr;
            }

            frame->childItem.attributes = childAttributes;

            if (childKind == LocalCopyPathKind::SemanticLink)
            {
                if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
                {
                    hadSkipped.store(true, std::memory_order_release);
                    continue;
                }

                frame->childItem.kind = RecursiveCopyWorkKind::ReparsePoint;
            }
            else if (childKind == LocalCopyPathKind::RegularDirectory)
            {
                if (! context.recursive)
                {
                    return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
                }
                else
                {
                    queuedDirectories.fetch_add(1, std::memory_order_relaxed);
                    directoryHr = runDirectory(context, frame->childItem, depth + 1u);
                    if (FAILED(directoryHr))
                    {
                        return directoryHr;
                    }
                    continue;
                }
            }
            else
            {
                frame->childItem.kind = RecursiveCopyWorkKind::File;
            }

            directoryHr = enqueueWork(std::move(frame->childItem));
            if (FAILED(directoryHr))
            {
                return directoryHr;
            }

            if ((++frame->cancelCheckCounter % 64u) == 0u)
            {
                directoryHr = CheckCancel(context);
                if (FAILED(directoryHr))
                {
                    return directoryHr;
                }
            }
        } while (FindNextFileW(frame->findHandle.get(), &frame->findData));

        const DWORD enumError = GetLastError();
        if (enumError != ERROR_NO_MORE_FILES)
        {
            return HRESULT_FROM_WIN32(enumError);
        }

        frame->findHandle.reset();

        directoryHr = waitForQueueDrain();
        if (FAILED(directoryHr))
        {
            return directoryHr;
        }
        if (frame->createdDestination)
        {
            restoreDirectoryMetadata();
        }

        processedDirectories.fetch_add(1, std::memory_order_relaxed);
        return S_OK;
    };

    runDirectory = [&](OperationContext& context, const RecursiveCopyWorkItem& item, uint64_t depth) noexcept -> HRESULT
    {
        auto selectedItem = std::make_unique<RecursiveCopyWorkItem>(item);
        for (;;)
        {
            HRESULT itemHr = processDirectory(context, *selectedItem, depth);
            if (SUCCEEDED(itemHr))
            {
                return itemHr;
            }

            itemHr = NormalizeCancellation(itemHr);
            if (IsCancellationHr(itemHr))
            {
                completeWithError(itemHr);
                return itemHr;
            }
            if (itemHr == HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW) || itemHr == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY))
            {
                completeWithError(itemHr);
                return itemHr;
            }
            if (itemHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
            {
                hadSkipped.store(true, std::memory_order_release);
                return S_OK;
            }
            if (context.continueOnError)
            {
                hadFailure.store(true, std::memory_order_release);
                return S_OK;
            }

            FileSystemIssueAction issueAction = FileSystemIssueAction::Cancel;
            const HRESULT issueHr             = ReportIssue(context, itemHr, &issueAction);
            if (FAILED(issueHr))
            {
                completeWithError(issueHr);
                return issueHr;
            }

            switch (issueAction)
            {
                case FileSystemIssueAction::Overwrite: context.oneShotAllowOverwrite = true; continue;
                case FileSystemIssueAction::ReplaceLink:
                    context.oneShotAllowOverwrite   = true;
                    context.oneShotAllowReplaceLink = true;
                    continue;
                case FileSystemIssueAction::ReplaceReadOnly:
                    context.oneShotAllowOverwrite       = true;
                    context.oneShotAllowReplaceReadonly = true;
                    continue;
                case FileSystemIssueAction::PermanentDelete:
                case FileSystemIssueAction::Cancel: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                case FileSystemIssueAction::Retry: continue;
                case FileSystemIssueAction::KeepBoth:
                {
                    const HRESULT keepBothHr = SelectKeepBothPath(item.destination, true, selectedItem->destination);
                    if (FAILED(keepBothHr))
                    {
                        completeWithError(keepBothHr);
                        return keepBothHr;
                    }
                    continue;
                }
                case FileSystemIssueAction::Skip: hadFailure.store(true, std::memory_order_release); return S_OK;
                case FileSystemIssueAction::None:
                default: completeWithError(HRESULT_FROM_WIN32(ERROR_CANCELLED)); return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
        }
    };

    const auto processWork = [&](OperationContext& context, const RecursiveCopyWorkItem& item) noexcept
    {
        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            return;
        }

        // Worker contexts are reused across queue items; a per-conflict grant must never
        // survive into an unrelated item.
        auto clearGrantsOnExit = wil::scope_exit([&]() noexcept { ClearOneShotGrants(context); });

        PathInfo selectedDestination = item.destination;
        for (;;)
        {
            HRESULT itemHr     = S_OK;
            uint64_t itemBytes = 0;
            switch (item.kind)
            {
                case RecursiveCopyWorkKind::File:
                    itemHr = CopyFileInternal(context, item.source, selectedDestination, &itemBytes);
                    if (SUCCEEDED(itemHr))
                    {
                        processedFiles.fetch_add(1, std::memory_order_relaxed);
                    }
                    break;
                case RecursiveCopyWorkKind::ReparsePoint:
                    itemHr = PreserveReparsePointInternal(context, item.source, selectedDestination, item.attributes, &itemBytes);
                    break;
                case RecursiveCopyWorkKind::Finish: return;
            }

            if (itemHr == S_FALSE)
            {
                hadIdenticalFileSkip.store(true, std::memory_order_release);
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
                case FileSystemIssueAction::Overwrite: context.oneShotAllowOverwrite = true; continue;
                case FileSystemIssueAction::ReplaceLink:
                    context.oneShotAllowOverwrite   = true;
                    context.oneShotAllowReplaceLink = true;
                    continue;
                case FileSystemIssueAction::ReplaceReadOnly:
                    context.oneShotAllowOverwrite       = true;
                    context.oneShotAllowReplaceReadonly = true;
                    continue;
                case FileSystemIssueAction::Retry: continue;
                case FileSystemIssueAction::KeepBoth:
                {
                    const HRESULT keepBothHr = SelectKeepBothPath(item.destination, IsDirectory(item.attributes), selectedDestination);
                    if (FAILED(keepBothHr))
                    {
                        completeWithError(keepBothHr);
                        return;
                    }
                    continue;
                }
                case FileSystemIssueAction::Skip: hadFailure.store(true, std::memory_order_release); return;
                case FileSystemIssueAction::PermanentDelete:
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
    rootItem.source      = source;
    rootItem.destination = destination;
    rootItem.attributes  = attributes;
    queuedDirectories.fetch_add(1u, std::memory_order_relaxed);

    // Each scheduler worker stream reuses one OperationContext across the items it
    // happens to run, so progress-stream ids (and therefore bandwidth-graph colours)
    // stay stable per stream. A worker stream never runs two steps at once, so a slot
    // is only ever touched by one thread; the mutex just serialises lazy init.
    constexpr size_t kStreamContextSlots = kMaxWorkers;
    std::array<std::optional<OperationContext>, kStreamContextSlots> streamContexts{};
    std::mutex streamContextsMutex;
    const auto getStreamContext = [&](uint64_t streamId) -> OperationContext&
    {
        const size_t slot = streamId < streamContexts.size() ? static_cast<size_t>(streamId) : 0;
        std::scoped_lock lock(streamContextsMutex);
        if (! streamContexts[slot].has_value())
        {
            streamContexts[slot].emplace();
            initializeChildContext(*streamContexts[slot], streamId);
        }
        return *streamContexts[slot];
    };

    // Dynamic, self-feeding job: each dispatch processes exactly one queue item and
    // returns the worker to the pool, so concurrent operations round-robin fairly
    // instead of this copy pinning every scheduler worker for its whole duration.
    auto job = GetSharedFileOpsJobScheduler().StartDynamicJob(rootContext.options,
                                                              concurrency,
                                                              &queue.ready,
                                                              [&](uint64_t schedulerStreamId) noexcept -> DynamicStep
    {
        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            return DynamicStep::Finished;
        }

        RecursiveCopyWorkItem item{};
        {
            std::unique_lock lock(queue.mutex);
            if (queue.done)
            {
                return DynamicStep::Finished;
            }
            if (queue.items.empty())
            {
                // Transiently empty (another worker took the last item). Park; an in-flight
                // worker will enqueue more children or finish the tree and wake us.
                return DynamicStep::Idle;
            }

            const bool discoveryAhead = ! queue.producerDone && GetDiscoveryMode(rootContext) == FILESYSTEM_DISCOVERY_AHEAD;
            const size_t activeLimit  = Common::FileOperations::DiscoveryWorkerLimit(static_cast<size_t>(concurrency), queue.items.size(), discoveryAhead);
            if (queue.activeItems >= activeLimit)
            {
                return DynamicStep::Idle;
            }

            item = std::move(queue.items.front());
            queue.items.pop_front();
            queue.ready.fetch_sub(1, std::memory_order_release);
            ++queue.activeItems;
        }

        workItemDispatches.fetch_add(1, std::memory_order_relaxed);
        queue.cv.notify_all();
        if (item.kind == RecursiveCopyWorkKind::Finish)
        {
            finishActiveItem(item.retainedPathBytes);
            return DynamicStep::Finished;
        }
        processWork(getStreamContext(schedulerStreamId), item);
        finishActiveItem(item.retainedPathBytes);

        std::scoped_lock lock(queue.mutex);
        return queue.done ? DynamicStep::Finished : DynamicStep::Processed;
    });

    const HRESULT producerHr = runDirectory(rootContext, rootItem, 0u);
    if (SUCCEEDED(producerHr))
    {
        RecursiveCopyWorkItem finishItem{};
        finishItem.kind        = RecursiveCopyWorkKind::Finish;
        const HRESULT finishHr = enqueueWork(std::move(finishItem));
        if (FAILED(finishHr))
        {
            completeWithError(finishHr);
        }
    }
    else
    {
        completeWithError(producerHr);
    }
    {
        std::scoped_lock lock(queue.mutex);
        queue.producerDone = true;
        if (queue.items.empty() && queue.activeItems == 0u)
        {
            queue.done = true;
        }
    }
    uint32_t finalQueuedItems = 0u;
    {
        std::scoped_lock lock(queue.mutex);
        finalQueuedItems = static_cast<uint32_t>((std::min)(queue.retainedEntries, static_cast<size_t>(std::numeric_limits<uint32_t>::max())));
    }
    const HRESULT discoveryCloseHr = ReportDiscoveryProgress(rootContext, discoveredBytes, discoveredFiles, discoveredDirectories, finalQueuedItems, true);
    if (FAILED(discoveryCloseHr))
    {
        completeWithError(discoveryCloseHr);
    }
    queue.cv.notify_all();
    GetSharedFileOpsJobScheduler().NotifyDynamicWorkAvailable();

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
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.WorkItemDispatches",
                      L"",
                      workItemDispatches.load(std::memory_order_acquire),
                      processedFiles.load(std::memory_order_acquire) + processedDirectories.load(std::memory_order_acquire),
                      concurrency,
                      S_OK);
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.MaxTraversalDepth", L"", 0u, maxTraversalDepth, Common::FileOperations::kTraversalMaxDepth, producerHr);
    Debug::Perf::Emit(
        L"FileOps.CopyRecursiveParallel.MaxRetainedEntries", L"", 0u, queue.maxRetainedEntries, Common::FileOperations::kTraversalMaxQueuedEntries, producerHr);
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.MaxRetainedPathBytes",
                      L"",
                      0u,
                      queue.maxRetainedPathBytes,
                      Common::FileOperations::kTraversalMaxQueuedPathBytes,
                      producerHr);
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.MaxAncestorMetadataBytes",
                      L"",
                      0u,
                      maxRetainedAncestorMetadataBytes,
                      Common::FileOperations::kTraversalMaxMetadataBytes,
                      producerHr);

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

    if (hadFailure.load(std::memory_order_acquire) || hadSkipped.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    return hadIdenticalFileSkip.load(std::memory_order_acquire) ? S_FALSE : S_OK;
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
    LocalCopyPathKind sourceKind                        = LocalCopyPathKind::RegularFile;
    const HRESULT kindHr =
        attributes == INVALID_FILE_ATTRIBUTES ? HRESULT_FROM_WIN32(GetLastError()) : ClassifyLocalCopyPathKind(source.extended, attributes, sourceKind);
    if (FAILED(kindHr))
    {
        return kindHr;
    }
    const bool canParallelizeDirectory = sourceKind == LocalCopyPathKind::RegularDirectory && context.recursive && effectiveConcurrency > 1u;

    if (canParallelizeDirectory)
    {
        return CopyDirectoryChildrenParallel(context, source, destination, flags, reparsePointPolicy, effectiveConcurrency, bytesCopied);
    }

    if (effectiveConcurrency <= 1u && sourceKind == LocalCopyPathKind::RegularDirectory && context.recursive)
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

HRESULT DeletePathInternal(OperationContext& context,
                           const PathInfo& path,
                           bool discoveryAlreadyReported                = false,
                           FileSystemItemMutationResult* mutationResult = nullptr) noexcept;

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

    constexpr unsigned int kMaxAttempts = 32;
    for (unsigned int attempt = 0; attempt < kMaxAttempts; ++attempt)
    {
        HRESULT hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }

        std::wstring leaf;
#ifdef ENABLE_TESTS
        const std::wstring entropyFailurePath = GetSelfTestEnvironmentString(kCaseRenameEntropyFailurePathEnvVar);
        if (! entropyFailurePath.empty() && OrdinalString::EqualsNoCase(sourceExtended, MakePathInfo(entropyFailurePath).extended))
        {
            static_cast<void>(::SetEnvironmentVariableW(kCaseRenameEntropyFailurePathEnvVar.data(), nullptr));
            Debug::Perf::EmitCounter(L"FileOps.Rename.CaseTempEntropyFailureCount");
            return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
        }
#endif
        hr = Common::Paths::BuildUniqueSiblingName(std::wstring_view{}, std::wstring_view{L".rs_case_tmp_"}, std::wstring_view{}, 255u, leaf);
        if (FAILED(hr))
        {
            Debug::Error(L"FileSystem: cryptographic case-only rename staging name generation failed (hr=0x{0:08X}).", static_cast<unsigned long>(hr));
            Debug::Perf::EmitCounter(L"FileOps.Rename.CaseTempEntropyFailureCount");
            return hr;
        }

        std::wstring tempPath = AppendPath(directory, leaf);
        if (tempPath.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }

#ifdef ENABLE_TESTS
        const std::wstring collisionPath = GetSelfTestEnvironmentString(kCaseRenameTempCollisionPathEnvVar);
        if (! collisionPath.empty() && attempt == 0u && OrdinalString::EqualsNoCase(sourceExtended, MakePathInfo(collisionPath).extended))
        {
            const DWORD sourceAttributes            = ::GetFileAttributesW(sourceExtended.c_str());
            const bool sourceIsDirectory            = sourceAttributes != INVALID_FILE_ATTRIBUTES && (sourceAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
            const bool collisionCreated             = sourceIsDirectory ? ::CreateDirectoryW(tempPath.c_str(), nullptr) != FALSE : false;
            const std::wstring collisionPayloadPath = sourceIsDirectory ? AppendPath(tempPath, L"race.marker") : tempPath;
            wil::unique_hfile collisionFile(::CreateFileW(collisionPayloadPath.c_str(),
                                                          GENERIC_WRITE,
                                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                          nullptr,
                                                          CREATE_NEW,
                                                          FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_TEMPORARY,
                                                          nullptr));
            if (collisionFile && (collisionCreated || ! sourceIsDirectory))
            {
                constexpr char collisionPayload[] = "race";
                DWORD written                     = 0u;
                static_cast<void>(::WriteFile(collisionFile.get(), collisionPayload, static_cast<DWORD>(sizeof(collisionPayload) - 1u), &written, nullptr));
                collisionFile.reset();
                static_cast<void>(::SetEnvironmentVariableW(kCaseRenameTempCollisionPathEnvVar.data(), nullptr));
                static_cast<void>(::SetEnvironmentVariableW(kCaseRenameTempCollisionFiredEnvVar.data(), tempPath.c_str()));
            }
        }
#endif

        // The first hop is the exclusive claim: never replace a raced sibling, even when the
        // final rename carries an Overwrite receipt. ERROR_*_EXISTS simply chooses fresh entropy.
        const DWORD exclusiveTempFlags = renameFlags & ~MOVEFILE_REPLACE_EXISTING;
        if (! ::MoveFileExW(sourceExtended.c_str(), tempPath.c_str(), exclusiveTempFlags))
        {
            const DWORD error               = ::GetLastError();
            const bool destinationNowExists = ::GetFileAttributesW(tempPath.c_str()) != INVALID_FILE_ATTRIBUTES;
            if (error == ERROR_ALREADY_EXISTS || error == ERROR_FILE_EXISTS || (error == ERROR_ACCESS_DENIED && destinationNowExists))
            {
                Debug::Perf::EmitCounter(L"FileOps.Rename.CaseTempCollisionCount");
                continue;
            }
            return HRESULT_FROM_WIN32(error);
        }

        // The source now sits at the temp name. A failed revert leaves the object under a name
        // nobody requested: report that as unknown truth instead of a proved no-commit.
        const auto revertToSource = [&]() noexcept
        {
            const DWORD revertFlags = renameFlags & ~MOVEFILE_REPLACE_EXISTING;
            if (! ::MoveFileExW(tempPath.c_str(), sourceExtended.c_str(), revertFlags))
            {
                context.trackedNativeMove.store(TrackedPublicationTruth::Unknown, std::memory_order_release);
                Debug::Perf::EmitCounter(L"FileOps.Rename.CaseTempRevertFailureCount");
            }
        };

        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            revertToSource();
            return hr;
        }

        if (! ::MoveFileExW(tempPath.c_str(), destinationExtended.c_str(), renameFlags))
        {
            const DWORD error = ::GetLastError();
            revertToSource();
            return HRESULT_FROM_WIN32(error);
        }

        context.trackedNativeMove.store(TrackedPublicationTruth::Published, std::memory_order_release);
        Debug::Perf::EmitCounter(L"FileOps.Rename.CaseTempSuccessCount");
        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
}

HRESULT MovePathInternal(OperationContext& context, const PathInfo& source, const PathInfo& destination) noexcept
{
    // Every exit below either issued no rename or issued one whose failure is atomic, so
    // "not committed" is proved until a mutation site records Published or Unknown. A
    // previous attempt's Unknown/Published is never downgraded by this attempt.
    {
        TrackedPublicationTruth expected = TrackedPublicationTruth::NotObserved;
        static_cast<void>(context.trackedNativeMove.compare_exchange_strong(expected, TrackedPublicationTruth::NotPublished, std::memory_order_acq_rel));
    }

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
    bool sourceIsLink            = false;
    hr                           = IsNameSurrogateReparsePoint(source.extended, sourceAttributes, sourceIsLink);
    if (FAILED(hr))
    {
        return hr;
    }

    bool caseOnlyRename                     = false;
    DWORD destinationAttributes             = GetFileAttributesW(destination.extended.c_str());
    const bool destinationExistedBeforeCopy = destinationAttributes != INVALID_FILE_ATTRIBUTES;
    const bool destinationIsDirectory       = destinationExistedBeforeCopy && (destinationAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    bool destinationIsLink                  = false;
    if (destinationExistedBeforeCopy)
    {
        hr = IsNameSurrogateReparsePoint(destination.extended, destinationAttributes, destinationIsLink);
        if (FAILED(hr))
        {
            return hr;
        }
    }
    const bool destinationRegularDirectoryConflict = sourceIsDirectory && ! sourceIsLink && destinationIsDirectory && ! destinationIsLink;
    const bool allowOverwriteEffective             = HasOverwriteGrant(context);
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
        }

        if (! caseOnlyRename && destinationRegularDirectoryConflict)
        {
            // Native is one provider mutation only. Folder merge belongs to the host Managed route;
            // report a known non-commit before enumerating or relocating any child. A case-only
            // rename names the exact source object and must retain the provider's temp-rename path.
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        if (! caseOnlyRename && ((sourceIsLink && ! destinationIsLink) || sourceIsDirectory != destinationIsDirectory))
        {
            return HRESULT_FROM_WIN32(ERROR_DATATYPE_MISMATCH);
        }
        if (! caseOnlyRename && ! allowOverwriteEffective)
        {
            return HRESULT_FROM_WIN32(destinationIsLink ? ERROR_REPARSE_POINT_ENCOUNTERED : ERROR_ALREADY_EXISTS);
        }

        if (! caseOnlyRename && (destinationAttributes & FILE_ATTRIBUTE_READONLY) != 0)
        {
            if (! HasReplaceReadonlyGrant(context))
            {
                return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
            }
        }
    }

    if (destinationExistedBeforeCopy && ! caseOnlyRename && allowOverwriteEffective)
    {
        if (! context.objectBinding || ! context.options)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        wil::com_ptr<IFileSystemBoundObject> expectedDestination = context.oneShotExpectedDestination;
        if (! expectedDestination)
        {
            constexpr FileSystemBindFlags destinationFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA |
                                                                                              FILESYSTEM_BIND_RENAME | FILESYSTEM_BIND_PUBLICATION);
            hr = context.objectBinding->BindObject(destination.display.c_str(), destinationFlags, expectedDestination.put());
            if (FAILED(hr) || ! expectedDestination)
            {
                return FAILED(hr) ? hr : E_UNEXPECTED;
            }
        }

        constexpr FileSystemBindFlags sourceFlags =
            static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_RENAME);
        wil::com_ptr<IFileSystemBoundObject> sourceAuthority;
        hr = context.objectBinding->BindObject(source.display.c_str(), sourceFlags, sourceAuthority.put());
        if (FAILED(hr) || ! sourceAuthority)
        {
            return FAILED(hr) ? hr : E_UNEXPECTED;
        }

        BOOL sameObject = FALSE;
        hr              = sourceAuthority->IsSameObject(expectedDestination.get(), &sameObject);
        if (FAILED(hr))
        {
            return hr;
        }
        if (sameObject != FALSE)
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
        }

        FileSystemFlags renameFlags = FILESYSTEM_FLAG_ALLOW_OVERWRITE;
        if (HasReplaceReadonlyGrant(context))
        {
            renameFlags = static_cast<FileSystemFlags>(renameFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);
        }
        if (HasReplaceLinkGrant(context))
        {
            renameFlags = static_cast<FileSystemFlags>(renameFlags | FILESYSTEM_FLAG_ALLOW_REPLACE_LINK);
        }
        FileSystemConditionalMutationResult mutationResult{};
        mutationResult.sizeBytes = sizeof(mutationResult);
        wil::com_ptr<IFileSystemBoundObject> renamedAuthority;
        hr = sourceAuthority->RenameIfUnchanged(
            destination.display.c_str(), expectedDestination.get(), renameFlags, context.options, &mutationResult, renamedAuthority.put());
        if (mutationResult.outcomeKnown == FALSE)
        {
            context.trackedNativeMove.store(TrackedPublicationTruth::Unknown, std::memory_order_release);
            return HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
        }
        if (mutationResult.mutationCommitted == FALSE)
        {
            if (hr == HRESULT_FROM_WIN32(ERROR_FILE_INVALID) || hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
            {
                Debug::Perf::EmitCounter(L"FileOps.Conflict.ExpectedDestinationMismatchCount");
            }
            return FAILED(hr) ? hr : E_UNEXPECTED;
        }
        context.trackedNativeMove.store(TrackedPublicationTruth::Published, std::memory_order_release);
        if (! renamedAuthority || mutationResult.originalStillPresent != FALSE)
        {
            return E_UNEXPECTED;
        }
        Debug::Perf::EmitCounter(L"FileOps.Conflict.ExpectedDestinationBoundCount");
        return hr;
    }

    DWORD moveFlags = 0;
    if (allowOverwriteEffective)
    {
        moveFlags |= MOVEFILE_REPLACE_EXISTING;
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

    bool forceCaseOnlyTempForSelfTest = false;
#ifdef ENABLE_TESTS
    if (caseOnlyRename)
    {
        const std::wstring entropyFailurePath = GetSelfTestEnvironmentString(kCaseRenameEntropyFailurePathEnvVar);
        const std::wstring collisionPath      = GetSelfTestEnvironmentString(kCaseRenameTempCollisionPathEnvVar);
        forceCaseOnlyTempForSelfTest =
            (! entropyFailurePath.empty() && OrdinalString::EqualsNoCase(source.extended, MakePathInfo(entropyFailurePath).extended)) ||
            (! collisionPath.empty() && OrdinalString::EqualsNoCase(source.extended, MakePathInfo(collisionPath).extended));
    }
#endif
    if (! forceCaseOnlyTempForSelfTest &&
        MoveFileWithProgressW(source.extended.c_str(), destination.extended.c_str(), CopyProgressRoutine, &progress, renameFlags))
    {
        // The native rename committed; a later progress-callback failure must not hide that.
        context.trackedNativeMove.store(TrackedPublicationTruth::Published, std::memory_order_release);
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

    const DWORD error = forceCaseOnlyTempForSelfTest ? ERROR_ACCESS_DENIED : GetLastError();
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

    return HRESULT_FROM_WIN32(error);
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
        _mutationResult.outcomeKnown         = TRUE;
        _mutationResult.mutationCommitted    = SUCCEEDED(hrDelete) ? TRUE : FALSE;
        _mutationResult.originalStillPresent = SUCCEEDED(hrDelete) ? FALSE : TRUE;

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

    [[nodiscard]] const FileSystemItemMutationResult& GetMutationResult() const noexcept
    {
        return _mutationResult;
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
    FileSystemItemMutationResult _mutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
};

struct RecycleBinBatchEntry
{
    unsigned long itemIndex = 0;
    PathInfo path{};
    HRESULT result = S_OK;
    bool observed  = false;
    FileSystemItemMutationResult mutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
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
            entry->result                              = hrDelete;
            entry->observed                            = true;
            entry->mutationResult.outcomeKnown         = TRUE;
            entry->mutationResult.mutationCommitted    = SUCCEEDED(hrDelete) ? TRUE : FALSE;
            entry->mutationResult.originalStillPresent = SUCCEEDED(hrDelete) ? FALSE : TRUE;
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

HRESULT DeleteToRecycleBinCore(OperationContext& context, const PathInfo& path, FileSystemItemMutationResult* mutationResult) noexcept
{
    if (mutationResult != nullptr)
    {
        *mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
    }
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
    if (mutationResult != nullptr)
    {
        *mutationResult = progressSinkImpl->GetMutationResult();
    }
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

HRESULT DeleteToRecycleBin(OperationContext& context, const PathInfo& path, FileSystemItemMutationResult* mutationResult) noexcept
{
    if (path.display.empty())
    {
        return E_INVALIDARG;
    }

#ifdef ENABLE_TESTS
    if (TryInjectRecycleFailureForSelfTest(path, mutationResult))
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
#endif

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
        return RunShellOperationOnDedicatedStaThread([&]() noexcept { return DeleteToRecycleBinCore(context, path, mutationResult); });
    }
    if (FAILED(coInitHr))
    {
        return coInitHr;
    }

    return DeleteToRecycleBinCore(context, path, mutationResult);
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

#if defined(ENABLE_TESTS)
    g_recycleBinBatchTestCalls.fetch_add(1u, std::memory_order_relaxed);
    g_recycleBinBatchTestRequestedItems.fetch_add(items.size(), std::memory_order_relaxed);
    RecordRecycleBinBatchSizeForTest(items.size());
#endif

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
#if defined(ENABLE_TESTS)
    g_recycleBinBatchTestObservedItems.fetch_add(observedCount, std::memory_order_relaxed);
    g_recycleBinBatchTestFailedItems.fetch_add(failedCount, std::memory_order_relaxed);
#endif

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

HRESULT DeletePathInternal(OperationContext& context,
                           const PathInfo& path,
                           bool discoveryAlreadyReported,
                           FileSystemItemMutationResult* mutationResult) noexcept
{
    if (mutationResult != nullptr)
    {
        *mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
    }
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
        if (! discoveryAlreadyReported)
        {
            DeleteHandleSnapshot discoverySnapshot{};
            const HRESULT discoveryOpenHr = OpenPathForDeleteNoFollow(path.extended, discoverySnapshot);
            if (SUCCEEDED(discoveryOpenHr))
            {
                const HRESULT discoveryHr = ReportDeleteDiscoveryObject(context, discoverySnapshot.attributes, discoverySnapshot.fileBytes, 1u);
                if (FAILED(discoveryHr))
                {
                    return discoveryHr;
                }
            }
        }
        return DeleteToRecycleBin(context, path, mutationResult);
    }

    // Permanent Local deletion is synchronous and handle-bound. Until the selected root itself is
    // removed, every return path is a known non-commit for that root: cancellation, admission/open
    // failure, a locked leaf, and a recursive partial delete all leave the selected object present.
    // Recycle deliberately keeps the unknown default above because IFileOperation can lose terminal
    // per-item truth; its progress sink is the only authority allowed to replace that default.
    if (mutationResult != nullptr)
    {
        *mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
    }

    DeleteHandleSnapshot snapshot{};
    hr = OpenPathForDeleteNoFollow(path.extended, snapshot);
    if (FAILED(hr))
    {
        return hr;
    }

    const auto reportKnownDeletion = [mutationResult](HRESULT deleteHr) noexcept
    {
        if (deleteHr == S_OK && mutationResult != nullptr)
        {
            *mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, TRUE, FALSE};
        }
        return deleteHr;
    };

    if (! discoveryAlreadyReported)
    {
        hr = ReportDeleteDiscoveryObject(context, snapshot.attributes, snapshot.fileBytes, 1u);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    if ((snapshot.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        // Never traverse directory reparse points during delete recursion (junction/symlink safety).
        if (IsReparsePoint(snapshot.attributes))
        {
            return reportKnownDeletion(DeleteOpenedPathNoFollow(context, snapshot, true));
        }

        if (! context.recursive)
        {
            return reportKnownDeletion(DeleteOpenedPathNoFollow(context, snapshot, true));
        }

        if (snapshot.handle)
        {
            snapshot.handle.reset();
        }
        return reportKnownDeletion(DeleteDirectoryRecursive(context, path));
    }

    return reportKnownDeletion(DeleteOpenedPathNoFollow(context, snapshot, true));
}

constexpr size_t kDeleteTraversalBatchEntries          = 256u;
constexpr uint64_t kDeleteTraversalMaxTerminalFailures = 4'096u;
constexpr uint64_t kDeleteTraversalMaxFailurePathBytes = 16ull * 1024ull * 1024ull;

struct DeleteTraversalBatchEntry final
{
    PathInfo path;
    DWORD attributes   = 0;
    uint64_t sizeBytes = 0;
};

// R4-T2: one directory level of the permanent recursive Delete walk. The walk keeps these frames on
// an explicit stack instead of recursing through DeletePathInternal for every child directory, so
// tree depth is not a ceiling. A frame owns exactly the state the former recursion level owned: its
// terminal-failure set (released on pop), the current batch and cursor, and the results of that batch.
struct DeleteWalkFrame final
{
    PathInfo path;
    bool deleteRoot = true;
    std::unordered_set<std::wstring> terminalFailures;
    uint64_t localFailurePathBytes = 0u;
    std::vector<DeleteTraversalBatchEntry> batch;
    std::vector<HRESULT> results;
    size_t index             = 0u;
    bool batchLoaded         = false;
    unsigned int concurrency = 1u; // the root keeps its caller's request; a child uses the context budget, as its recursion did
};

[[nodiscard]] HRESULT DeleteDirectoryRecursiveBatched(OperationContext& context,
                                                      const PathInfo& rootPath,
                                                      unsigned int requestedConcurrency,
                                                      bool deleteRoot = true) noexcept
{
    std::vector<std::unique_ptr<DeleteWalkFrame>> frames;

    const auto pushFrame = [&](const PathInfo& path, bool frameDeletesRoot, unsigned int frameConcurrency) noexcept -> HRESULT
    {
        std::unique_ptr<DeleteWalkFrame> frame(new (std::nothrow) DeleteWalkFrame{});
        if (! frame)
        {
            return E_OUTOFMEMORY;
        }
        // Only std::bad_alloc can follow the nothrow allocation above, and std::bad_alloc is fatal
        // by policy (AGENTS.md), so the copies below are not wrapped.
        frame->path        = path;
        frame->deleteRoot  = frameDeletesRoot;
        frame->concurrency = std::clamp(frameConcurrency, 1u, 8u);
        frames.push_back(std::move(frame));
        ++context.deleteTraversalDepth;
        context.deleteTraversalMaxDepth = (std::max)(context.deleteTraversalMaxDepth, context.deleteTraversalDepth);
        return S_OK;
    };
    // The former recursion level released its terminal-failure record on return; a frame does so on pop.
    const auto popFrame = [&]() noexcept
    {
        DeleteWalkFrame& frame = *frames.back();
        context.deleteTraversalRetainedFailureCount -= frame.terminalFailures.size();
        context.deleteTraversalRetainedFailurePathBytes -= frame.localFailurePathBytes;
        --context.deleteTraversalDepth;
        frames.pop_back();
    };
    const auto unwind = [&]() noexcept
    {
        while (! frames.empty())
        {
            popFrame();
        }
    };

    const auto retainTerminalFailure = [&](DeleteWalkFrame& frame, const PathInfo& failedPath) noexcept -> HRESULT
    {
        if (frame.terminalFailures.contains(failedPath.extended))
        {
            return S_OK;
        }

        const uint64_t pathBytes = static_cast<uint64_t>(failedPath.extended.size()) * sizeof(wchar_t);
        if (context.deleteTraversalRetainedFailureCount >= kDeleteTraversalMaxTerminalFailures || pathBytes > kDeleteTraversalMaxFailurePathBytes ||
            context.deleteTraversalRetainedFailurePathBytes > kDeleteTraversalMaxFailurePathBytes - pathBytes)
        {
            Debug::Warning(L"FileSystem: recursive delete stopped at '{}' after reaching the bounded terminal-failure record limit.", failedPath.display);
            return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
        }

        frame.terminalFailures.insert(failedPath.extended);
        ++context.deleteTraversalRetainedFailureCount;
        frame.localFailurePathBytes += pathBytes;
        context.deleteTraversalRetainedFailurePathBytes += pathBytes;
        context.deleteTraversalMaxRetainedFailureCount =
            (std::max)(context.deleteTraversalMaxRetainedFailureCount, context.deleteTraversalRetainedFailureCount);
        context.deleteTraversalMaxRetainedFailurePathBytes =
            (std::max)(context.deleteTraversalMaxRetainedFailurePathBytes, context.deleteTraversalRetainedFailurePathBytes);
        return S_OK;
    };

    // (Re)enumerates one batch of the frame's directory. S_OK: entries follow in frame.batch;
    // S_FALSE: the directory came back empty and was finalized (deleted, or kept when the frame
    // does not own its root); a failure is the frame's result.
    const auto loadBatch = [&](DeleteWalkFrame& frame) noexcept -> HRESULT
    {
        const PathInfo& path = frame.path;
        HRESULT hr           = CheckCancel(context);
        if (FAILED(hr))
        {
            return hr;
        }

        DirectoryEnumHandle enumeration{};
        hr = OpenDirectoryForEnumerationNoFollow(path.extended, enumeration);
        if (FAILED(hr))
        {
            if (HRESULT_CODE(hr) == ERROR_FILE_NOT_FOUND || HRESULT_CODE(hr) == ERROR_PATH_NOT_FOUND)
            {
                return frame.terminalFailures.empty() ? S_FALSE : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }
            return hr;
        }

        if (! IsDirectory(enumeration.attributes) || IsReparsePoint(enumeration.attributes))
        {
            if (! frame.deleteRoot)
            {
                return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
            }
            enumeration.handle.reset();
            DeleteHandleSnapshot swappedSnapshot{};
            hr = OpenPathForDeleteNoFollow(path.extended, swappedSnapshot);
            if (FAILED(hr))
            {
                return hr;
            }
            hr = DeleteOpenedPathNoFollow(context, swappedSnapshot, true);
            return SUCCEEDED(hr) ? S_FALSE : hr;
        }

        frame.batch.clear();
        frame.batch.reserve(kDeleteTraversalBatchEntries);
        hr = EnumerateDirectoryByHandle(enumeration.handle.get(),
                                        [&](std::wstring_view name, DWORD attributes, uint64_t sizeBytes) noexcept -> HRESULT
        {
            const std::wstring childName(name);
            DeleteTraversalBatchEntry entry{};
            entry.path.display  = AppendPath(path.display, childName.c_str());
            entry.path.extended = AppendPath(path.extended, childName.c_str());
            entry.attributes    = attributes;
            entry.sizeBytes     = sizeBytes;
#if defined(ENABLE_TESTS)
            MaybeInjectDeleteToctouSwapForSelfTest(entry.path.extended);
#endif
            if (frame.terminalFailures.contains(entry.path.extended))
            {
                return S_OK;
            }

            const HRESULT discoveryHr = ReportDeleteDiscoveryObject(context, attributes, sizeBytes, static_cast<uint32_t>(frame.batch.size() + 1u));
            if (FAILED(discoveryHr))
            {
                return discoveryHr;
            }
            frame.batch.push_back(std::move(entry));
            return frame.batch.size() >= kDeleteTraversalBatchEntries ? S_FALSE : S_OK;
        });
        enumeration.handle.reset();
        if (FAILED(hr))
        {
            return hr;
        }

        context.deleteTraversalMaxBatchEntries = (std::max)(context.deleteTraversalMaxBatchEntries, static_cast<uint64_t>(frame.batch.size()));
        if (frame.batch.empty())
        {
            if (! frame.terminalFailures.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
            }

            if (! frame.deleteRoot)
            {
                return S_FALSE;
            }

            DeleteHandleSnapshot snapshot{};
            hr = OpenPathForDeleteNoFollow(path.extended, snapshot);
            if (FAILED(hr))
            {
                return hr;
            }
            hr = DeleteOpenedPathNoFollow(context, snapshot, true);
            return SUCCEEDED(hr) ? S_FALSE : hr;
        }

        frame.results.assign(frame.batch.size(), E_UNEXPECTED);
        frame.index       = 0u;
        frame.batchLoaded = true;
        return S_OK;
    };

    // The per-entry result handling the former recursion applied after a batch ran.
    const auto settleBatch = [&](DeleteWalkFrame& frame) noexcept -> HRESULT
    {
        for (size_t index = 0u; index < frame.batch.size(); ++index)
        {
            const HRESULT itemHr = frame.results[index];
            if (SUCCEEDED(itemHr))
            {
                continue;
            }
            if (IsCancellationHr(itemHr))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            if (! context.continueOnError)
            {
                return itemHr;
            }
            const HRESULT hr = retainTerminalFailure(frame, frame.batch[index].path);
            if (FAILED(hr))
            {
                return hr;
            }
        }
        frame.batchLoaded = false; // enumerate again until the directory comes back empty
        return S_OK;
    };

    // A serial child directory: the pre-steps DeletePathInternal performs for a directory, without
    // its recursion. S_OK with pushed == true means a child frame is now on top of the stack;
    // otherwise the returned value is the entry's result.
    const auto beginChildDirectory = [&](const DeleteTraversalBatchEntry& entry, bool& pushed) noexcept -> HRESULT
    {
        pushed     = false;
        HRESULT hr = SetProgressPaths(context, entry.path.display.c_str(), nullptr);
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

        DeleteHandleSnapshot snapshot{};
        hr = OpenPathForDeleteNoFollow(entry.path.extended, snapshot);
        if (FAILED(hr))
        {
            return hr;
        }
        // Never traverse directory reparse points during delete recursion (junction/symlink safety);
        // an entry that is no longer a real directory is deleted in place, as DeletePathInternal does.
        if ((snapshot.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0 || IsReparsePoint(snapshot.attributes) || ! context.recursive)
        {
            return DeleteOpenedPathNoFollow(context, snapshot, true);
        }
        snapshot.handle.reset();
        hr = pushFrame(entry.path, true, context.deleteConcurrencyBudget);
        if (FAILED(hr))
        {
            return hr;
        }
        pushed = true;
        return S_OK;
    };

    HRESULT rootHr = pushFrame(rootPath, deleteRoot, requestedConcurrency);
    if (FAILED(rootHr))
    {
        return rootHr;
    }

    HRESULT result = S_OK;
    while (! frames.empty())
    {
        DeleteWalkFrame& frame = *frames.back();
        HRESULT frameResult    = S_OK;
        bool frameDone         = false;

        if (! frame.batchLoaded)
        {
            const HRESULT loadHr = loadBatch(frame);
            if (loadHr == S_FALSE)
            {
                frameResult = frame.terminalFailures.empty() ? S_OK : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                frameDone   = true;
            }
            else if (FAILED(loadHr))
            {
                frameResult = loadHr;
                frameDone   = true;
            }
        }

        if (! frameDone)
        {
            const bool runParallelBatch =
                frame.concurrency > 1u && context.parallel == nullptr && frame.batch.size() > 1u && GetSharedFileOpsJobScheduler().EnsureWorkersAvailable();
            if (runParallelBatch)
            {
                // Workers call DeletePathInternal on their entry (one nesting each) and walk the
                // directory below it through this same loop with a serial budget.
                const auto processEntry = [&](size_t index, uint64_t schedulerStreamId) noexcept
                {
                    if (index >= frame.batch.size())
                    {
                        return;
                    }
                    OperationContext worker{};
                    InitializeOperationContext(
                        worker,
                        FILESYSTEM_DELETE,
                        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE |
                                                     (context.continueOnError ? FILESYSTEM_FLAG_CONTINUE_ON_ERROR : FILESYSTEM_FLAG_NONE) |
                                                     (context.allowReplaceReadonly ? FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY : FILESYSTEM_FLAG_NONE)),
                        context.options,
                        context.callback,
                        context.callbackCookie,
                        0,
                        context.reparsePointPolicy);
                    worker.parallel                = context.parallel;
                    worker.progressStreamId        = schedulerStreamId;
                    worker.deleteConcurrencyBudget = 1u;
                    worker.deleteDiscovery         = context.deleteDiscovery;
                    worker.deleteTraversalDepth    = context.deleteTraversalDepth;
                    frame.results[index]           = DeletePathInternal(worker, frame.batch[index].path, true);
                };

                ParallelOperationState parallel{};
                parallel.startTick = GetTickCount64();
                parallel.bandwidthLimitBytesPerSecond.store(context.options ? context.options->bandwidthLimitBytesPerSecond : 0ull, std::memory_order_release);
                context.parallel     = &parallel;
                auto restoreParallel = wil::scope_exit([&]() noexcept { context.parallel = nullptr; });
                auto job             = GetSharedFileOpsJobScheduler().StartJob(context.options, frame.concurrency, frame.batch.size(), processEntry);
                GetSharedFileOpsJobScheduler().WaitJob(job);
                context.parallel = nullptr;
                restoreParallel.release();
                AddCompletedItems(context, parallel.completedItems.load(std::memory_order_acquire));
                AddCompletedBytes(context, parallel.completedBytes.load(std::memory_order_acquire));
                if (parallel.cancelRequested.load(std::memory_order_acquire))
                {
                    unwind();
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                {
                    const HRESULT first = parallel.firstError.load(std::memory_order_acquire);
                    unwind();
                    return FAILED(first) ? first : HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                frame.index = frame.batch.size();
            }
            else if (frame.index < frame.batch.size())
            {
                // Serial: one entry per iteration; a child directory becomes the top frame and its
                // result lands in this batch's results when it pops.
                const size_t index                     = frame.index;
                const DeleteTraversalBatchEntry& entry = frame.batch[index];
                HRESULT entryHr                        = S_OK;
                if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && ! IsReparsePoint(entry.attributes))
                {
                    bool pushed = false;
                    entryHr     = beginChildDirectory(entry, pushed);
                    if (pushed)
                    {
                        continue;
                    }
                }
                else
                {
                    entryHr = DeletePathInternal(context, entry.path, true);
                }
                frame.results[index] = entryHr;
                ++frame.index;
                if (IsCancellationHr(entryHr))
                {
                    unwind();
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                continue;
            }

            // Every entry of the batch ran: settle the results and enumerate again.
            const HRESULT settleHr = settleBatch(frame);
            if (FAILED(settleHr))
            {
                frameResult = settleHr;
                frameDone   = true;
            }
            else
            {
                continue;
            }
        }

        // The frame's result: the root returns it; a child delivers it as its entry's result.
        const bool rootFrame = frames.size() == 1u;
        popFrame();
        if (rootFrame)
        {
            result = frameResult;
            break;
        }
        DeleteWalkFrame& parent      = *frames.back();
        parent.results[parent.index] = frameResult;
        ++parent.index;
        if (IsCancellationHr(frameResult))
        {
            unwind();
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }
    unwind();
    return result;
}

HRESULT DeleteDirectoryRecursive(OperationContext& context, const PathInfo& path) noexcept
{
    const unsigned int requestedConcurrency = std::clamp(context.deleteConcurrencyBudget, 1u, 8u);
    // Every caller uses the same no-follow, bounded discovery/execution walk. When no operation
    // control is supplied, discovery reporting becomes a no-op without changing traversal shape.
    return DeleteDirectoryRecursiveBatched(context, path, requestedConcurrency);
}
} // namespace

#if defined(ENABLE_TESTS)
// Test-enabled builds expose a process-local snapshot so the host selftest can hard-gate the
// scheduler route without treating live Shell/Defender/indexer wall time as deterministic.
// The selftest resets and reads this only while no recycle-bin task is active.
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderFileSystemTestGetRecycleBinBatchSnapshot(BOOL reset,
                                                                                                         uint64_t* batchCalls,
                                                                                                         uint64_t* requestedItems,
                                                                                                         uint64_t* observedItems,
                                                                                                         uint64_t* failedItems,
                                                                                                         uint64_t* fallbackCount,
                                                                                                         uint64_t* maxBatchSize) noexcept
{
    if (batchCalls == nullptr || requestedItems == nullptr || observedItems == nullptr || failedItems == nullptr || fallbackCount == nullptr ||
        maxBatchSize == nullptr)
    {
        return E_POINTER;
    }

    const auto snapshot = [reset](std::atomic<uint64_t>& value) noexcept -> uint64_t
    { return reset == TRUE ? value.exchange(0u, std::memory_order_acq_rel) : value.load(std::memory_order_acquire); };

    *batchCalls     = snapshot(g_recycleBinBatchTestCalls);
    *requestedItems = snapshot(g_recycleBinBatchTestRequestedItems);
    *observedItems  = snapshot(g_recycleBinBatchTestObservedItems);
    *failedItems    = snapshot(g_recycleBinBatchTestFailedItems);
    *fallbackCount  = snapshot(g_recycleBinBatchTestFallbacks);
    *maxBatchSize   = snapshot(g_recycleBinBatchTestMaxBatchSize);
    return S_OK;
}
#endif

HRESULT FileSystemInternal::DeleteBoundLocalDirectoryContents(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options) noexcept
{
    if (path == nullptr || path[0] == L'\0' || ! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }

    OperationContext context{};
    InitializeOperationContext(
        context,
        FILESYSTEM_DELETE,
        static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE |
                                     (HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR) ? FILESYSTEM_FLAG_CONTINUE_ON_ERROR : FILESYSTEM_FLAG_NONE)),
        options,
        nullptr,
        nullptr,
        0u,
        FileSystemReparsePointPolicy::Preserve);
    DeleteDiscoveryState discovery{};
    context.deleteDiscovery         = &discovery;
    context.deleteConcurrencyBudget = 1u;
    const PathInfo target           = MakePathInfo(path);
    const HRESULT hr                = DeleteDirectoryRecursiveBatched(context, target, 1u, false);
    Debug::Perf::Emit(L"FileOps.DeleteTraversal.MaxDepth",
                      L"bound-directory-contents",
                      0u,
                      context.deleteTraversalMaxDepth,
                      0u, // R4-T2: no walk-depth ceiling; depth is reported only
                      hr);
    return hr;
}

HRESULT FileSystemInternal::ReadBoundLocalLink(HANDLE boundHandle, const FileSystemLinkTransform& transform, FileSystemLinkInformation& information) noexcept
{
    information.kind                            = static_cast<FileSystemLinkKind>(0u);
    information.targetIsRelative                = FALSE;
    information.targetMapping                   = static_cast<FileSystemLinkTargetMapping>(0u);
    information.targetLengthUtf16               = 0u;
    information.sourceRelativeTargetLengthUtf16 = 0u;
    if (transform.sizeBytes != sizeof(FileSystemLinkTransform) || information.sizeBytes != sizeof(FileSystemLinkInformation) ||
        transform.sourceLinkPath == nullptr || transform.destinationLinkPath == nullptr || transform.sourceRootPath == nullptr ||
        transform.destinationRootPath == nullptr || transform.sourceLinkPath[0] == L'\0' || transform.destinationLinkPath[0] == L'\0' ||
        transform.sourceRootPath[0] == L'\0' || transform.destinationRootPath[0] == L'\0' ||
        transform.componentMappingCount > Common::FileOperations::kTraversalMaxQueuedEntries ||
        (transform.componentMappingCount != 0u && transform.componentMappings == nullptr))
    {
        return E_INVALIDARG;
    }

    ReparsePointData reparse{};
    HRESULT hr = ReadReparsePointDataHandle(boundHandle, reparse);
    if (FAILED(hr))
    {
        return hr;
    }
    if (reparse.tag != IO_REPARSE_TAG_SYMLINK && reparse.tag != IO_REPARSE_TAG_MOUNT_POINT)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    ParsedDirectoryReparsePoint parsed{};
    if (! ParseDirectoryReparsePoint(reparse, parsed))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    FILE_ATTRIBUTE_TAG_INFO attributes{};
    if (GetFileInformationByHandleEx(boundHandle, FileAttributeTagInfo, &attributes, sizeof(attributes)) == FALSE)
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    if (parsed.tag == IO_REPARSE_TAG_MOUNT_POINT)
    {
        information.kind = FILESYSTEM_LINK_KIND_JUNCTION;
    }
    else
    {
        information.kind =
            (attributes.FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u ? FILESYSTEM_LINK_KIND_SYMBOLIC_DIRECTORY : FILESYSTEM_LINK_KIND_SYMBOLIC_FILE;
    }

    // Literal Preserve: the payload is the stored target text and relative flag of the source link.
    // No target is resolved, opened, or rewritten, so a link into the copied tree keeps naming the
    // source location. The transform's roots and component mappings were validated above for ABI
    // hygiene only; the mapping enum's other values are reserved for an optional transform.
    const std::wstring encodedTarget = ! parsed.printPath.empty() ? parsed.printPath : NtPathToWin32Path(parsed.substitutePath);
    if (encodedTarget.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (encodedTarget.size() > 32'767u || encodedTarget.size() > (std::numeric_limits<uint32_t>::max)())
    {
        return HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
    }

    information.targetMapping                   = FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT;
    information.targetIsRelative                = parsed.tag == IO_REPARSE_TAG_SYMLINK && parsed.isRelative ? TRUE : FALSE;
    information.targetLengthUtf16               = static_cast<uint32_t>(encodedTarget.size());
    information.sourceRelativeTargetLengthUtf16 = 0u;
    if (information.targetBuffer == nullptr ||
        static_cast<uint64_t>(information.targetCapacityUtf16) < static_cast<uint64_t>(information.targetLengthUtf16) + 1u)
    {
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    }
    std::memcpy(information.targetBuffer, encodedTarget.data(), encodedTarget.size() * sizeof(wchar_t));
    information.targetBuffer[encodedTarget.size()] = L'\0';
    return S_OK;
}

HRESULT FileSystemInternal::CreateExclusiveLocalLink(const wchar_t* stagePath,
                                                     const FileSystemLinkInformation& information,
                                                     wil::unique_handle& ownedHandle) noexcept
{
    ownedHandle.reset();
    if (stagePath == nullptr || stagePath[0] == L'\0' || information.sizeBytes != sizeof(FileSystemLinkInformation) || information.targetBuffer == nullptr ||
        information.targetLengthUtf16 == 0u || information.targetLengthUtf16 >= information.targetCapacityUtf16 || information.targetLengthUtf16 > 32'767u ||
        information.targetBuffer[information.targetLengthUtf16] != L'\0' ||
        std::wmemchr(information.targetBuffer, L'\0', information.targetLengthUtf16) != nullptr)
    {
        return E_INVALIDARG;
    }
    if (information.targetMapping != FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT &&
        information.targetMapping != FILESYSTEM_LINK_TARGET_MAPPED_INSIDE_SOURCE_ROOT)
    {
        return E_INVALIDARG;
    }
    if (information.kind != FILESYSTEM_LINK_KIND_SYMBOLIC_FILE && information.kind != FILESYSTEM_LINK_KIND_SYMBOLIC_DIRECTORY &&
        information.kind != FILESYSTEM_LINK_KIND_JUNCTION)
    {
        return E_INVALIDARG;
    }
    if (information.kind == FILESYSTEM_LINK_KIND_JUNCTION && information.targetIsRelative != FALSE)
    {
        return E_INVALIDARG;
    }

    std::wstring target(information.targetBuffer, information.targetLengthUtf16);
    ReparsePointData reparse{};
    HRESULT hr = information.kind == FILESYSTEM_LINK_KIND_JUNCTION ? BuildMountPointReparseData(std::move(target), reparse)
                                                                   : BuildSymlinkReparseData(std::move(target), information.targetIsRelative != FALSE, reparse);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring extendedStage = ToExtendedPath(stagePath);
    const bool directoryLink         = information.kind != FILESYSTEM_LINK_KIND_SYMBOLIC_FILE;
    wil::unique_handle handle;
    if (directoryLink)
    {
        if (CreateDirectoryW(extendedStage.c_str(), nullptr) == FALSE)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
        handle.reset(CreateFileW(extendedStage.c_str(),
                                 FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES | DELETE,
                                 FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                 nullptr,
                                 OPEN_EXISTING,
                                 FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                 nullptr));
        if (! handle)
        {
            const DWORD error = GetLastError();
            static_cast<void>(RemoveDirectoryW(extendedStage.c_str()));
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
    }
    else
    {
        handle.reset(CreateFileW(extendedStage.c_str(),
                                 FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES | DELETE,
                                 FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                 nullptr,
                                 CREATE_NEW,
                                 FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_OPEN_REPARSE_POINT,
                                 nullptr));
        if (! handle)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
        }
    }

    hr = WriteReparsePointDataHandle(handle.get(), reparse);
    if (FAILED(hr))
    {
        FILE_DISPOSITION_INFO_EX disposition{};
        disposition.Flags = FILE_DISPOSITION_FLAG_DELETE | FILE_DISPOSITION_FLAG_POSIX_SEMANTICS | FILE_DISPOSITION_FLAG_IGNORE_READONLY_ATTRIBUTE;
        static_cast<void>(SetFileInformationByHandle(handle.get(), FileDispositionInfoEx, &disposition, sizeof(disposition)));
        return hr;
    }

    ownedHandle = std::move(handle);
    return S_OK;
}

#if defined(_DEBUG)
void FileSystemInternal::RunDebugObjectBindingSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept -> bool
    {
        if (condition)
        {
            ++passed;
            return true;
        }
        ++failed;
        Debug::Error(L"FileSystem object-binding selftest failed: {}", message);
        return false;
    };

    if (! check(IsNameSurrogateReparseTag(IO_REPARSE_TAG_SYMLINK) && IsNameSurrogateReparseTag(IO_REPARSE_TAG_MOUNT_POINT) &&
                    ! IsNameSurrogateReparseTag(IO_REPARSE_TAG_WOF) && ! IsNameSurrogateReparseTag(IO_REPARSE_TAG_CLOUD) &&
                    ! IsNameSurrogateReparseTag(IO_REPARSE_TAG_DEDUP),
                L"Move conflict classification must distinguish link name-surrogates from WOF/cloud-style reparses"))
    {
        return;
    }
    if (! check(LocalCopyPathKindFromFacts(FILE_ATTRIBUTE_REPARSE_POINT, IsNameSurrogateReparseTag(IO_REPARSE_TAG_WOF)) == LocalCopyPathKind::RegularFile &&
                    LocalCopyPathKindFromFacts(FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT, IsNameSurrogateReparseTag(IO_REPARSE_TAG_CLOUD)) ==
                        LocalCopyPathKind::RegularDirectory &&
                    LocalCopyPathKindFromFacts(FILE_ATTRIBUTE_REPARSE_POINT, IsNameSurrogateReparseTag(IO_REPARSE_TAG_SYMLINK)) ==
                        LocalCopyPathKind::SemanticLink &&
                    LocalCopyPathKindFromFacts(FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT,
                                               IsNameSurrogateReparseTag(IO_REPARSE_TAG_MOUNT_POINT)) == LocalCopyPathKind::SemanticLink,
                L"Local Copy must route only name-surrogate reparses through semantic-link Preserve/Skip"))
    {
        return;
    }

    std::array<wchar_t, MAX_PATH + 1u> tempPathBuffer{};
    const DWORD tempPathLength = GetTempPathW(static_cast<DWORD>(tempPathBuffer.size()), tempPathBuffer.data());
    if (! check(tempPathLength > 0u && tempPathLength < tempPathBuffer.size(), L"temporary path should be available"))
    {
        return;
    }

    const std::filesystem::path root =
        std::filesystem::path(tempPathBuffer.data()) / std::format(L"RedSalamander-ObjectBinding-{}-{}", GetCurrentProcessId(), GetTickCount64());
    const std::filesystem::path original              = root / L"original.bin";
    const std::filesystem::path alias                 = root / L"alias.bin";
    const std::filesystem::path missing               = root / L"missing.bin";
    const std::filesystem::path target                = root / L"target";
    const std::filesystem::path targetSentinel        = target / L"sentinel.bin";
    const std::filesystem::path junction              = root / L"junction";
    const std::filesystem::path relativeFileLink      = root / L"relative-file-link";
    const std::filesystem::path relativeDirectoryLink = root / L"relative-directory-link";
    const std::filesystem::path absoluteDirectoryLink = root / L"absolute-directory-link";
    const std::filesystem::path rootRelativeFileLink  = root / L"root-relative-file-link";
    const std::filesystem::path outsideTarget =
        root.parent_path() / std::format(L"RedSalamander-ObjectBinding-Outside-{}-{}.bin", GetCurrentProcessId(), GetTickCount64());
    const std::filesystem::path outsideFileLink          = root / L"outside-file-link";
    const std::filesystem::path stagePath                = root / L"stage.bin";
    const std::filesystem::path publishedPath            = root / L"published.bin";
    const std::filesystem::path abortStagePath           = root / L"abort-stage.bin";
    const std::filesystem::path replaceStagePath         = root / L"replace-stage.bin";
    const std::filesystem::path replaceDestinationPath   = root / L"replace-destination.bin";
    const std::filesystem::path raceStagePath            = root / L"race-stage.bin";
    const std::filesystem::path raceDestinationPath      = root / L"race-destination.bin";
    const std::filesystem::path raceMovedPath            = root / L"race-original-moved.bin";
    const std::filesystem::path retargetRoot             = root / L"retargeted";
    const std::filesystem::path linkStagePath            = root / L"link-stage";
    const std::filesystem::path directoryStagePath       = root / L"directory-stage";
    const std::filesystem::path directoryPublishedPath   = root / L"directory-published";
    const std::filesystem::path linkReplacementStagePath = root / L"link-replacement-stage";
    const std::filesystem::path retainedStagePath        = root / L"retained-stage.bin";
    const std::filesystem::path retainedPublishedPath    = root / L"retained-published.bin";
    const std::filesystem::path uncertainDeleteStagePath = root / L"uncertain-delete-stage.bin";

    if (! check(CreateDirectoryW(root.c_str(), nullptr) != FALSE, L"fixture root should be created"))
    {
        return;
    }
    auto cleanup = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(RemoveDirectoryW(junction.c_str()));
        static_cast<void>(DeleteFileW(targetSentinel.c_str()));
        static_cast<void>(DeleteFileW(relativeFileLink.c_str()));
        static_cast<void>(RemoveDirectoryW(relativeDirectoryLink.c_str()));
        static_cast<void>(RemoveDirectoryW(absoluteDirectoryLink.c_str()));
        static_cast<void>(DeleteFileW(rootRelativeFileLink.c_str()));
        static_cast<void>(DeleteFileW(outsideFileLink.c_str()));
        static_cast<void>(DeleteFileW(outsideTarget.c_str()));
        static_cast<void>(DeleteFileW(original.c_str()));
        static_cast<void>(DeleteFileW(alias.c_str()));
        static_cast<void>(DeleteFileW(stagePath.c_str()));
        static_cast<void>(DeleteFileW(publishedPath.c_str()));
        static_cast<void>(DeleteFileW(abortStagePath.c_str()));
        static_cast<void>(DeleteFileW(replaceStagePath.c_str()));
        static_cast<void>(DeleteFileW(replaceDestinationPath.c_str()));
        static_cast<void>(DeleteFileW(raceStagePath.c_str()));
        static_cast<void>(DeleteFileW(raceDestinationPath.c_str()));
        static_cast<void>(DeleteFileW(raceMovedPath.c_str()));
        static_cast<void>(RemoveDirectoryW(linkStagePath.c_str()));
        static_cast<void>(RemoveDirectoryW(directoryStagePath.c_str()));
        static_cast<void>(RemoveDirectoryW(directoryPublishedPath.c_str()));
        static_cast<void>(RemoveDirectoryW(linkReplacementStagePath.c_str()));
        static_cast<void>(DeleteFileW(retainedStagePath.c_str()));
        static_cast<void>(DeleteFileW(retainedPublishedPath.c_str()));
        static_cast<void>(DeleteFileW(uncertainDeleteStagePath.c_str()));
        static_cast<void>(RemoveDirectoryW(target.c_str()));
        static_cast<void>(RemoveDirectoryW(root.c_str()));
    });

    {
        wil::unique_hfile seed(CreateFileW(
            original.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
        constexpr std::array<std::byte, 4> payload{{std::byte{0x11}, std::byte{0x22}, std::byte{0x33}, std::byte{0x44}}};
        DWORD written = 0u;
        if (! check(seed && WriteFile(seed.get(), payload.data(), static_cast<DWORD>(payload.size()), &written, nullptr) != FALSE && written == payload.size(),
                    L"fixture file should be written"))
        {
            return;
        }
    }
    if (! check(CreateHardLinkW(alias.c_str(), original.c_str(), nullptr) != FALSE, L"hard-link alias should be created"))
    {
        return;
    }
    {
        wil::unique_hfile outside(CreateFileW(
            outsideTarget.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! check(static_cast<bool>(outside), L"outside-link target should be created"))
        {
            return;
        }
    }

    wil::com_ptr<IFileSystem> fileSystem;
    fileSystem.attach(new (std::nothrow) FileSystem());
    if (! check(static_cast<bool>(fileSystem), L"local provider should allocate"))
    {
        return;
    }
    wil::com_ptr<IFileSystemObjectBinding> binding;
    if (! check(SUCCEEDED(fileSystem->QueryInterface(__uuidof(IFileSystemObjectBinding), binding.put_void())) && binding,
                L"local provider should expose optional object binding"))
    {
        return;
    }

    constexpr FileSystemBindFlags readFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_CONTENT | FILESYSTEM_BIND_READ_METADATA);
    wil::com_ptr<IFileSystemBoundObject> originalBound;
    wil::com_ptr<IFileSystemBoundObject> aliasBound;
    if (! check(SUCCEEDED(binding->BindObject(original.c_str(), readFlags, originalBound.put())) && originalBound, L"regular file should bind") ||
        ! check(SUCCEEDED(binding->BindObject(alias.c_str(), readFlags, aliasBound.put())) && aliasBound, L"hard-link alias should bind"))
    {
        return;
    }

    BOOL same = FALSE;
    if (! check(SUCCEEDED(originalBound->IsSameObject(aliasBound.get(), &same)) && same != FALSE, L"hard-link names should compare as the same exact object"))
    {
        return;
    }
    FileSystemBoundObjectSnapshot snapshot{};
    snapshot.sizeBytes = sizeof(snapshot);
    if (! check(SUCCEEDED(originalBound->GetSnapshot(&snapshot)) && snapshot.kind == FILESYSTEM_BOUND_REGULAR_FILE && snapshot.objectId != nullptr &&
                    snapshot.objectIdBytes == 32u && snapshot.revisionId == nullptr && snapshot.revisionIdBytes == 0u && snapshot.committedSizeBytes == 4u,
                L"regular snapshot should expose FILE_ID_INFO authority and no invented revision"))
    {
        return;
    }

    wil::com_ptr<IFileReader> reader;
    if (! check(SUCCEEDED(originalBound->OpenReader(nullptr, reader.put())) && reader, L"read-authorized binding should reopen by handle"))
    {
        return;
    }
    std::array<std::byte, 4> readPayload{};
    unsigned long bytesRead = 0u;
    if (! check(SUCCEEDED(reader->Read(readPayload.data(), static_cast<unsigned long>(readPayload.size()), &bytesRead)) && bytesRead == readPayload.size() &&
                    readPayload[0] == std::byte{0x11} && readPayload[3] == std::byte{0x44},
                L"bound reader should read the retained regular object"))
    {
        return;
    }
    reader.reset();

    wil::com_ptr<IFileSystemBoundObject> absent;
    const HRESULT missingHr = binding->BindObject(missing.c_str(), readFlags, absent.put());
    if (! check((missingHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || missingHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND)) && ! absent,
                L"disappeared path should report missing without a bound object"))
    {
        return;
    }

    if (! check(DeleteFileW(original.c_str()) != FALSE, L"original hard-link name should be removed"))
    {
        return;
    }
    {
        wil::unique_hfile replacement(CreateFileW(
            original.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! check(static_cast<bool>(replacement), L"replacement object should be created at the same path"))
        {
            return;
        }
    }
    wil::com_ptr<IFileSystemBoundObject> replacementBound;
    if (! check(SUCCEEDED(binding->BindObject(original.c_str(), readFlags, replacementBound.put())) && replacementBound, L"replacement path should bind"))
    {
        return;
    }
    same = TRUE;
    if (! check(SUCCEEDED(originalBound->IsSameObject(replacementBound.get(), &same)) && same == FALSE,
                L"same pathname after replacement must compare as a changed object"))
    {
        return;
    }

    if (! check(CreateDirectoryW(target.c_str(), nullptr) != FALSE && CreateDirectoryW(junction.c_str(), nullptr) != FALSE,
                L"junction fixture directories should be created"))
    {
        return;
    }
    ReparsePointData junctionData{};
    if (! check(SUCCEEDED(BuildMountPointReparseData(target.wstring(), junctionData)) && SUCCEEDED(WriteReparsePointData(junction.wstring(), junctionData)),
                L"junction fixture should be materialized"))
    {
        return;
    }

    const auto createSymbolicLink = [](const std::filesystem::path& linkPath, std::wstring_view targetText, bool directory, bool relative) noexcept -> bool
    {
        bool placeholderCreated = false;
        if (directory)
        {
            placeholderCreated = CreateDirectoryW(linkPath.c_str(), nullptr) != FALSE;
        }
        else
        {
            wil::unique_hfile placeholder(CreateFileW(linkPath.c_str(),
                                                      FILE_WRITE_ATTRIBUTES,
                                                      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                      nullptr,
                                                      CREATE_NEW,
                                                      FILE_ATTRIBUTE_NORMAL,
                                                      nullptr));
            placeholderCreated = static_cast<bool>(placeholder);
        }
        if (! placeholderCreated)
        {
            return false;
        }
        auto removePlaceholder = wil::scope_exit([&]() noexcept
        {
            if (directory)
            {
                static_cast<void>(RemoveDirectoryW(linkPath.c_str()));
            }
            else
            {
                static_cast<void>(DeleteFileW(linkPath.c_str()));
            }
        });
        ReparsePointData data{};
        if (FAILED(BuildSymlinkReparseData(std::wstring(targetText), relative, data)) || FAILED(WriteReparsePointData(linkPath.wstring(), data)))
        {
            return false;
        }
        removePlaceholder.release();
        return true;
    };
    const std::wstring sourceRootName = original.root_name().wstring();
    const std::wstring rootRelativeTarget =
        sourceRootName.empty() || original.native().size() <= sourceRootName.size() ? std::wstring{} : original.native().substr(sourceRootName.size());
    if (! check(createSymbolicLink(relativeFileLink, original.filename().wstring(), false, true), L"relative file-symlink fixture should be created") ||
        ! check(createSymbolicLink(relativeDirectoryLink, target.filename().wstring(), true, true), L"relative directory-symlink fixture should be created") ||
        ! check(createSymbolicLink(absoluteDirectoryLink, target.native(), true, false), L"absolute directory-symlink fixture should be created") ||
        ! check(! rootRelativeTarget.empty() && createSymbolicLink(rootRelativeFileLink, rootRelativeTarget, false, false),
                L"root-relative file-symlink fixture should be created") ||
        ! check(createSymbolicLink(outsideFileLink, outsideTarget.native(), false, false), L"outside-root file-symlink fixture should be created"))
    {
        return;
    }

    wil::com_ptr<IFileSystemBoundObject> junctionBound;
    wil::com_ptr<IFileSystemBoundObject> targetBound;
    constexpr FileSystemBindFlags metadataFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    if (! check(SUCCEEDED(binding->BindObject(junction.c_str(), metadataFlags, junctionBound.put())) && junctionBound,
                L"junction object should bind no-follow") ||
        ! check(SUCCEEDED(binding->BindObject(target.c_str(), metadataFlags, targetBound.put())) && targetBound,
                L"junction target directory should bind separately"))
    {
        return;
    }
    FileSystemBoundObjectSnapshot junctionSnapshot{};
    junctionSnapshot.sizeBytes = sizeof(junctionSnapshot);
    same                       = TRUE;
    if (! check(SUCCEEDED(junctionBound->GetSnapshot(&junctionSnapshot)) && junctionSnapshot.kind == FILESYSTEM_BOUND_LINK &&
                    SUCCEEDED(junctionBound->IsSameObject(targetBound.get(), &same)) && same == FALSE,
                L"no-follow junction binding must identify the link object, never its target"))
    {
        return;
    }

    FileSystemOptions options{};
    options.sizeBytes                        = sizeof(options);
    const auto validateSymbolicLinkTransform = [&](const std::filesystem::path& sourceLink,
                                                   FileSystemLinkKind expectedKind,
                                                   bool expectedRelative,
                                                   FileSystemLinkTargetMapping expectedMapping,
                                                   const std::filesystem::path& expectedAbsoluteTarget) noexcept -> bool
    {
        wil::com_ptr<IFileSystemBoundObject> boundLink;
        if (FAILED(binding->BindObject(sourceLink.c_str(), metadataFlags, boundLink.put())) || ! boundLink)
        {
            return false;
        }

        const std::filesystem::path destination = retargetRoot / sourceLink.filename();
        FileSystemLinkTransform transform{};
        transform.sizeBytes           = sizeof(transform);
        transform.sourceLinkPath      = sourceLink.c_str();
        transform.destinationLinkPath = destination.c_str();
        transform.sourceRootPath      = root.c_str();
        transform.destinationRootPath = retargetRoot.c_str();
        FileSystemLinkInformation information{};
        information.sizeBytes = sizeof(information);
        HRESULT hr            = binding->ReadBoundLink(boundLink.get(), &transform, &options, &information);
        if (hr != HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER) || information.kind != expectedKind ||
            (information.targetIsRelative != FALSE) != expectedRelative || information.targetMapping != expectedMapping || information.targetLengthUtf16 == 0u)
        {
            return false;
        }

        std::vector<wchar_t> targetText(static_cast<size_t>(information.targetLengthUtf16) + 1u, L'\0');
        std::vector<wchar_t> sourceRelativeText;
        information.targetBuffer        = targetText.data();
        information.targetCapacityUtf16 = static_cast<uint32_t>(targetText.size());
        if (information.sourceRelativeTargetLengthUtf16 != 0u)
        {
            sourceRelativeText.assign(static_cast<size_t>(information.sourceRelativeTargetLengthUtf16) + 1u, L'\0');
            information.sourceRelativeTargetBuffer        = sourceRelativeText.data();
            information.sourceRelativeTargetCapacityUtf16 = static_cast<uint32_t>(sourceRelativeText.size());
        }
        hr = binding->ReadBoundLink(boundLink.get(), &transform, &options, &information);
        if (FAILED(hr))
        {
            return false;
        }

        std::filesystem::path resolvedTarget(std::wstring_view(targetText.data(), information.targetLengthUtf16));
        if (information.targetIsRelative != FALSE)
        {
            resolvedTarget = (destination.parent_path() / resolvedTarget).lexically_normal();
        }
        const std::wstring normalizedActual = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(resolvedTarget.wstring())));
        const std::wstring normalizedExpected =
            TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(expectedAbsoluteTarget.wstring())));
        return OrdinalString::EqualsNoCase(normalizedActual, normalizedExpected);
    };
    // C3: literal Preserve. Every payload reports the stored target text with the outside-root
    // mapping and no source-relative component; roots and component mappings never change it.
    if (! check(validateSymbolicLinkTransform(
                    relativeFileLink, FILESYSTEM_LINK_KIND_SYMBOLIC_FILE, true, FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT, retargetRoot / original.filename()),
                L"relative file symlink should keep its kind and relative target text") ||
        ! check(validateSymbolicLinkTransform(relativeDirectoryLink,
                                              FILESYSTEM_LINK_KIND_SYMBOLIC_DIRECTORY,
                                              true,
                                              FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT,
                                              retargetRoot / target.filename()),
                L"relative directory symlink should keep its kind and relative target text") ||
        ! check(validateSymbolicLinkTransform(
                    absoluteDirectoryLink, FILESYSTEM_LINK_KIND_SYMBOLIC_DIRECTORY, false, FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT, target),
                L"absolute directory symlink should keep naming its source-side target") ||
        ! check(validateSymbolicLinkTransform(rootRelativeFileLink,
                                              FILESYSTEM_LINK_KIND_SYMBOLIC_FILE,
                                              false,
                                              FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT,
                                              std::filesystem::path(rootRelativeTarget)),
                L"root-relative file symlink should keep its root-relative target text") ||
        ! check(validateSymbolicLinkTransform(
                    outsideFileLink, FILESYSTEM_LINK_KIND_SYMBOLIC_FILE, false, FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT, outsideTarget),
                L"outside-root absolute file symlink should preserve its target spelling and meaning"))
    {
        return;
    }

    const std::filesystem::path destinationLink = retargetRoot / junction.filename();
    FileSystemLinkTransform linkTransform{};
    linkTransform.sizeBytes           = sizeof(linkTransform);
    linkTransform.sourceLinkPath      = junction.c_str();
    linkTransform.destinationLinkPath = destinationLink.c_str();
    linkTransform.sourceRootPath      = root.c_str();
    linkTransform.destinationRootPath = retargetRoot.c_str();
    FileSystemLinkInformation linkInformation{};
    linkInformation.sizeBytes = sizeof(linkInformation);
    HRESULT linkReadHr        = binding->ReadBoundLink(junctionBound.get(), &linkTransform, &options, &linkInformation);
    if (! check(linkReadHr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER) && linkInformation.kind == FILESYSTEM_LINK_KIND_JUNCTION &&
                    linkInformation.targetLengthUtf16 > 0u && linkInformation.targetMapping == FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT &&
                    linkInformation.sourceRelativeTargetLengthUtf16 == 0u,
                L"bound junction read should report its literal payload size and no source-relative component"))
    {
        return;
    }
    std::vector<wchar_t> literalTarget(static_cast<size_t>(linkInformation.targetLengthUtf16) + 1u, L'\0');
    linkInformation.targetBuffer        = literalTarget.data();
    linkInformation.targetCapacityUtf16 = static_cast<uint32_t>(literalTarget.size());
    linkReadHr                          = binding->ReadBoundLink(junctionBound.get(), &linkTransform, &options, &linkInformation);
    const auto linkTargetsEqual         = [](std::wstring_view left, std::wstring_view right) noexcept
    {
        const std::wstring normalizedLeft  = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(std::wstring(left))));
        const std::wstring normalizedRight = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(std::wstring(right))));
        return OrdinalString::EqualsNoCase(normalizedLeft, normalizedRight);
    };
    if (! check(SUCCEEDED(linkReadHr) && linkInformation.kind == FILESYSTEM_LINK_KIND_JUNCTION && linkInformation.targetIsRelative == FALSE &&
                    linkTargetsEqual(std::wstring_view(literalTarget.data(), linkInformation.targetLengthUtf16), target.native()),
                L"bound junction read should keep the stored absolute target text"))
    {
        return;
    }

    const std::wstring sourceMappedComponent       = target.filename().wstring();
    constexpr wchar_t destinationMappedComponent[] = L"target (2)";
    const FileSystemLinkComponentMapping componentMapping{
        .sourceRelativeComponentPath = sourceMappedComponent.c_str(),
        .destinationComponentName    = destinationMappedComponent,
    };
    linkTransform.componentMappings     = &componentMapping;
    linkTransform.componentMappingCount = 1u;
    linkReadHr                          = binding->ReadBoundLink(junctionBound.get(), &linkTransform, &options, &linkInformation);
    if (! check(SUCCEEDED(linkReadHr) && linkInformation.targetMapping == FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT &&
                    linkInformation.sourceRelativeTargetLengthUtf16 == 0u &&
                    linkTargetsEqual(std::wstring_view(literalTarget.data(), linkInformation.targetLengthUtf16), target.native()),
                L"component mappings must never change a literal link payload"))
    {
        return;
    }

    wil::com_ptr<IFileSystemBoundObject> ownedLinkStage;
    if (! check(SUCCEEDED(binding->CreateExclusiveLink(linkStagePath.c_str(), &linkInformation, &options, ownedLinkStage.put())) && ownedLinkStage &&
                    (GetFileAttributesW(linkStagePath.c_str()) & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) ==
                        (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT),
                L"exclusive link stage should return exact ownership of a visible link object"))
    {
        return;
    }
    wil::com_ptr<IFileSystemBoundObject> collidingLinkStage;
    const HRESULT linkCollisionHr = binding->CreateExclusiveLink(linkStagePath.c_str(), &linkInformation, &options, collidingLinkStage.put());
    if (! check((linkCollisionHr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) || linkCollisionHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS)) &&
                    ! collidingLinkStage,
                L"exclusive link-stage collision should fail without adopting the existing link"))
    {
        return;
    }
    FileSystemConditionalMutationResult linkAbortResult{};
    linkAbortResult.sizeBytes = sizeof(linkAbortResult);
    if (! check(SUCCEEDED(ownedLinkStage->AbortOwnedObject(&options, &linkAbortResult)) && linkAbortResult.outcomeKnown != FALSE &&
                    linkAbortResult.mutationCommitted != FALSE && linkAbortResult.originalStillPresent == FALSE &&
                    GetFileAttributesW(linkStagePath.c_str()) == INVALID_FILE_ATTRIBUTES,
                L"exclusive link stage abort should remove only the exact owned link"))
    {
        return;
    }

    const auto writeFile = [&](const std::filesystem::path& path, std::byte firstByte) noexcept -> bool
    {
        wil::unique_hfile file(CreateFileW(
            path.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
        const std::array<std::byte, 4u> bytes{{firstByte, std::byte{0x42}, std::byte{0x43}, std::byte{0x44}}};
        DWORD written = 0u;
        return file && WriteFile(file.get(), bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr) != FALSE && written == bytes.size();
    };
    const auto readFirstByte = [&](const std::filesystem::path& path, std::byte& firstByte) noexcept -> bool
    {
        wil::unique_hfile file(CreateFileW(
            path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        DWORD read = 0u;
        return file && ReadFile(file.get(), &firstByte, sizeof(firstByte), &read, nullptr) != FALSE && read == sizeof(firstByte);
    };
    const auto writeStage = [&](IFileWriter* writer, std::byte firstByte) noexcept -> bool
    {
        if (writer == nullptr)
        {
            return false;
        }
        const std::array<std::byte, 4u> bytes{{firstByte, std::byte{0x52}, std::byte{0x53}, std::byte{0x54}}};
        unsigned long written = 0u;
        return SUCCEEDED(writer->Write(bytes.data(), static_cast<unsigned long>(bytes.size()), &written)) && written == bytes.size() &&
               SUCCEEDED(writer->Commit());
    };

    constexpr FileSystemBindFlags publicationFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_PUBLICATION);
    wil::com_ptr<IFileSystemBoundObject> directoryStage;
    if (! check(SUCCEEDED(binding->CreateExclusiveDirectory(directoryStagePath.c_str(), &options, directoryStage.put())) && directoryStage,
                L"exclusive directory creation should return exact owned authority"))
    {
        return;
    }
    FileSystemBoundObjectSnapshot directorySnapshot{};
    directorySnapshot.sizeBytes = sizeof(directorySnapshot);
    wil::com_ptr<IFileSystemBoundObject> collidingDirectoryStage;
    const HRESULT directoryCollisionHr = binding->CreateExclusiveDirectory(directoryStagePath.c_str(), &options, collidingDirectoryStage.put());
    if (! check(SUCCEEDED(directoryStage->GetSnapshot(&directorySnapshot)) && directorySnapshot.kind == FILESYSTEM_BOUND_DIRECTORY,
                L"exclusive directory authority should report directory kind") ||
        ! check((directoryCollisionHr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) || directoryCollisionHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS)) &&
                    ! collidingDirectoryStage,
                L"exclusive directory collision must not adopt the existing object"))
    {
        return;
    }

    FileSystemConditionalMutationResult directoryPublishResult{};
    directoryPublishResult.sizeBytes = sizeof(directoryPublishResult);
    wil::com_ptr<IFileSystemBoundObject> publishedDirectory;
    if (! check(SUCCEEDED(directoryStage->PublishAs(
                    directoryPublishedPath.c_str(), nullptr, FILESYSTEM_FLAG_NONE, &options, &directoryPublishResult, publishedDirectory.put())) &&
                    directoryPublishResult.outcomeKnown != FALSE && directoryPublishResult.mutationCommitted != FALSE && publishedDirectory &&
                    GetFileAttributesW(directoryStagePath.c_str()) == INVALID_FILE_ATTRIBUTES,
                L"exclusive directory stage should publish only-if-absent through its exact handle"))
    {
        return;
    }

    if (! check(writeFile(targetSentinel, std::byte{0x5a}), L"link target sentinel should be created before typed replacement"))
    {
        return;
    }
    wil::com_ptr<IFileSystemBoundObject> expectedLink;
    wil::com_ptr<IFileSystemBoundObject> linkReplacementStage;
    if (! check(SUCCEEDED(binding->BindObject(junction.c_str(), publicationFlags, expectedLink.put())) && expectedLink,
                L"destination junction should bind no-follow for typed replacement") ||
        ! check(SUCCEEDED(binding->CreateExclusiveDirectory(linkReplacementStagePath.c_str(), &options, linkReplacementStage.put())) && linkReplacementStage,
                L"typed link replacement should own an exclusive real-directory stage"))
    {
        return;
    }

    FileSystemConditionalMutationResult linkReplaceResult{};
    linkReplaceResult.sizeBytes = sizeof(linkReplaceResult);
    wil::com_ptr<IFileSystemBoundObject> replacedLink;
    const HRESULT untypedLinkReplaceHr = linkReplacementStage->PublishAs(
        junction.c_str(), expectedLink.get(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, &options, &linkReplaceResult, replacedLink.put());
    const DWORD junctionAttributesBeforeTypedReplace = GetFileAttributesW(junction.c_str());
    if (! check(untypedLinkReplaceHr == HRESULT_FROM_WIN32(ERROR_REPARSE_POINT_ENCOUNTERED) && linkReplaceResult.outcomeKnown != FALSE &&
                    linkReplaceResult.mutationCommitted == FALSE && ! replacedLink && junctionAttributesBeforeTypedReplace != INVALID_FILE_ATTRIBUTES &&
                    (junctionAttributesBeforeTypedReplace & FILE_ATTRIBUTE_REPARSE_POINT) != 0u,
                L"ordinary Overwrite must not replace an exact destination link"))
    {
        return;
    }

    linkReplaceResult                           = {};
    linkReplaceResult.sizeBytes                 = sizeof(linkReplaceResult);
    const FileSystemFlags typedLinkReplaceFlags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_LINK);
    if (! check(SUCCEEDED(linkReplacementStage->PublishAs(
                    junction.c_str(), expectedLink.get(), typedLinkReplaceFlags, &options, &linkReplaceResult, replacedLink.put())) &&
                    linkReplaceResult.outcomeKnown != FALSE && linkReplaceResult.mutationCommitted != FALSE && replacedLink,
                L"typed Replace Link should publish the exact owned directory over the exact junction object"))
    {
        return;
    }
    const DWORD replacedLinkAttributes = GetFileAttributesW(junction.c_str());
    if (! check(replacedLinkAttributes != INVALID_FILE_ATTRIBUTES && (replacedLinkAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u &&
                    (replacedLinkAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0u && GetFileAttributesW(targetSentinel.c_str()) != INVALID_FILE_ATTRIBUTES,
                L"typed Replace Link should replace only the link object and preserve its target tree"))
    {
        return;
    }

    wil::com_ptr<IFileWriter> stageWriter;
    wil::com_ptr<IFileSystemBoundObject> ownedStage;
    if (! check(SUCCEEDED(binding->CreateExclusiveWriter(stagePath.c_str(), &options, stageWriter.put(), ownedStage.put())) && stageWriter && ownedStage,
                L"exclusive stage creation should return writer and exact ownership") ||
        ! check((GetFileAttributesW(stagePath.c_str()) & FILE_ATTRIBUTE_HIDDEN) == 0u, L"owned stages remain visible to the file manager"))
    {
        return;
    }

    wil::com_ptr<IFileWriter> collidingWriter;
    wil::com_ptr<IFileSystemBoundObject> collidingStage;
    const HRESULT collisionHr = binding->CreateExclusiveWriter(stagePath.c_str(), &options, collidingWriter.put(), collidingStage.put());
    if (! check((collisionHr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS) || collisionHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS)) && ! collidingWriter &&
                    ! collidingStage,
                L"exclusive stage collision must fail without adopting the existing stage") ||
        ! check(writeStage(stageWriter.get(), std::byte{0x61}), L"owned stage should accept and commit exact bytes"))
    {
        return;
    }
    stageWriter.reset();

    FileSystemConditionalMutationResult publishResult{};
    publishResult.sizeBytes = sizeof(publishResult);
    wil::com_ptr<IFileSystemBoundObject> published;
    if (! check(SUCCEEDED(ownedStage->PublishAs(publishedPath.c_str(), nullptr, FILESYSTEM_FLAG_NONE, &options, &publishResult, published.put())) &&
                    publishResult.outcomeKnown != FALSE && publishResult.mutationCommitted != FALSE && publishResult.originalStillPresent == FALSE && published,
                L"owned stage should publish only-if-absent through its exact handle") ||
        ! check(GetFileAttributesW(stagePath.c_str()) == INVALID_FILE_ATTRIBUTES, L"publication should remove the owned stage name") ||
        ! check(GetFileAttributesW(publishedPath.c_str()) != INVALID_FILE_ATTRIBUTES, L"publication should expose the final visible name"))
    {
        return;
    }

    same = FALSE;
    if (! check(SUCCEEDED(ownedStage->IsSameObject(published.get(), &same)) && same != FALSE, L"published authority should retain the stage object identity"))
    {
        return;
    }

    wil::com_ptr<IFileWriter> abortWriter;
    wil::com_ptr<IFileSystemBoundObject> abortStage;
    if (! check(SUCCEEDED(binding->CreateExclusiveWriter(abortStagePath.c_str(), &options, abortWriter.put(), abortStage.put())) && abortWriter && abortStage &&
                    writeStage(abortWriter.get(), std::byte{0x62}),
                L"abort fixture should create and commit an owned stage"))
    {
        return;
    }
    abortWriter.reset();
    FileSystemConditionalMutationResult abortResult{};
    abortResult.sizeBytes = sizeof(abortResult);
    if (! check(SUCCEEDED(abortStage->AbortOwnedObject(&options, &abortResult)) && abortResult.outcomeKnown != FALSE &&
                    abortResult.mutationCommitted != FALSE && abortResult.originalStillPresent == FALSE &&
                    GetFileAttributesW(abortStagePath.c_str()) == INVALID_FILE_ATTRIBUTES,
                L"AbortOwnedObject should delete only the exact retained stage"))
    {
        return;
    }

    if (! check(writeFile(replaceDestinationPath, std::byte{0x6f}), L"replacement destination should be created"))
    {
        return;
    }
    wil::com_ptr<IFileSystemBoundObject> expectedDestination;
    wil::com_ptr<IFileWriter> replaceWriter;
    wil::com_ptr<IFileSystemBoundObject> replaceStage;
    if (! check(SUCCEEDED(binding->BindObject(replaceDestinationPath.c_str(), publicationFlags, expectedDestination.put())) && expectedDestination,
                L"existing destination should bind for exact replacement") ||
        ! check(SUCCEEDED(binding->CreateExclusiveWriter(replaceStagePath.c_str(), &options, replaceWriter.put(), replaceStage.put())) && replaceWriter &&
                    replaceStage && writeStage(replaceWriter.get(), std::byte{0x6e}),
                L"replacement stage should commit before publication"))
    {
        return;
    }
    replaceWriter.reset();
    publishResult           = {};
    publishResult.sizeBytes = sizeof(publishResult);
    published.reset();
    if (! check(SUCCEEDED(replaceStage->PublishAs(
                    replaceDestinationPath.c_str(), expectedDestination.get(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, &options, &publishResult, published.put())) &&
                    publishResult.mutationCommitted != FALSE && published,
                L"replacement should publish the stage and remove only the exact expected destination"))
    {
        return;
    }
    std::byte firstByte{};
    if (! check(readFirstByte(replaceDestinationPath, firstByte) && firstByte == std::byte{0x6e}, L"replacement destination should contain the staged bytes"))
    {
        return;
    }

    if (! check(writeFile(raceDestinationPath, std::byte{0x70}), L"race destination should be created"))
    {
        return;
    }
    wil::com_ptr<IFileSystemBoundObject> staleExpectedDestination;
    wil::com_ptr<IFileWriter> raceWriter;
    wil::com_ptr<IFileSystemBoundObject> raceStage;
    if (! check(SUCCEEDED(binding->BindObject(raceDestinationPath.c_str(), publicationFlags, staleExpectedDestination.put())) && staleExpectedDestination,
                L"race destination should bind before replacement") ||
        ! check(SUCCEEDED(binding->CreateExclusiveWriter(raceStagePath.c_str(), &options, raceWriter.put(), raceStage.put())) && raceWriter && raceStage &&
                    writeStage(raceWriter.get(), std::byte{0x71}),
                L"race stage should commit before publication") ||
        ! check(MoveFileExW(raceDestinationPath.c_str(), raceMovedPath.c_str(), 0u) != FALSE && writeFile(raceDestinationPath, std::byte{0x72}),
                L"race should replace the destination pathname with a different object"))
    {
        return;
    }
    FileSystemBasicInformation staleDestinationInfo{};
    staleDestinationInfo.sizeBytes = sizeof(staleDestinationInfo);
    if (! check(SUCCEEDED(staleExpectedDestination->GetBasicInformation(&staleDestinationInfo)),
                L"stale destination authority should read metadata through its exact handle"))
    {
        return;
    }
    staleDestinationInfo.attributes |= FILE_ATTRIBUTE_HIDDEN;
    if (! check(SUCCEEDED(staleExpectedDestination->SetBasicInformation(&staleDestinationInfo)) &&
                    (GetFileAttributesW(raceMovedPath.c_str()) & FILE_ATTRIBUTE_HIDDEN) != 0u &&
                    (GetFileAttributesW(raceDestinationPath.c_str()) & FILE_ATTRIBUTE_HIDDEN) == 0u,
                L"bound metadata mutation after a pathname swap must touch the original object, never the substitute"))
    {
        return;
    }
    raceWriter.reset();
    publishResult           = {};
    publishResult.sizeBytes = sizeof(publishResult);
    published.reset();
    const HRESULT racePublishHr = raceStage->PublishAs(
        raceDestinationPath.c_str(), staleExpectedDestination.get(), FILESYSTEM_FLAG_ALLOW_OVERWRITE, &options, &publishResult, published.put());
    firstByte = {};
    if (! check(FAILED(racePublishHr) && publishResult.outcomeKnown != FALSE && publishResult.mutationCommitted == FALSE && ! published &&
                    readFirstByte(raceDestinationPath, firstByte) && firstByte == std::byte{0x72} &&
                    GetFileAttributesW(raceStagePath.c_str()) != INVALID_FILE_ATTRIBUTES,
                L"replacement race must preserve the racer and the exact owned stage"))
    {
        return;
    }
    abortResult           = {};
    abortResult.sizeBytes = sizeof(abortResult);
    if (! check(SUCCEEDED(raceStage->AbortOwnedObject(&options, &abortResult)) && GetFileAttributesW(raceStagePath.c_str()) == INVALID_FILE_ATTRIBUTES,
                L"failed publication should clean up only through the retained stage token"))
    {
        return;
    }

    constexpr wchar_t kStageIdentityPathEnv[]           = L"REDSALAMANDER_FILEOPS_STAGE_IDENTITY_UNSUPPORTED_PATH";
    constexpr wchar_t kStageIdentityFiredEnv[]          = L"REDSALAMANDER_FILEOPS_STAGE_IDENTITY_UNSUPPORTED_FIRED";
    constexpr wchar_t kRenameCommittedFailurePathEnv[]  = L"REDSALAMANDER_FILEOPS_RENAME_COMMITTED_FAILURE_PATH";
    constexpr wchar_t kRenameCommittedFailureFiredEnv[] = L"REDSALAMANDER_FILEOPS_RENAME_COMMITTED_FAILURE_FIRED";
    constexpr wchar_t kDeleteCommittedFailurePathEnv[]  = L"REDSALAMANDER_FILEOPS_DELETE_COMMITTED_FAILURE_PATH";
    constexpr wchar_t kDeleteCommittedFailureFiredEnv[] = L"REDSALAMANDER_FILEOPS_DELETE_COMMITTED_FAILURE_FIRED";
    const auto clearMutationTruthInjections             = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(kStageIdentityPathEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kStageIdentityFiredEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kRenameCommittedFailurePathEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kRenameCommittedFailureFiredEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kDeleteCommittedFailurePathEnv, nullptr));
        static_cast<void>(SetEnvironmentVariableW(kDeleteCommittedFailureFiredEnv, nullptr));
    });

    static_cast<void>(SetEnvironmentVariableW(kStageIdentityPathEnv, retainedStagePath.c_str()));
    wil::com_ptr<IFileWriter> retainedWriter;
    wil::com_ptr<IFileSystemBoundObject> retainedStage;
    if (! check(SUCCEEDED(binding->CreateExclusiveWriter(retainedStagePath.c_str(), &options, retainedWriter.put(), retainedStage.put())) && retainedWriter &&
                    retainedStage && writeStage(retainedWriter.get(), std::byte{0x73}) && GetEnvironmentVariableW(kStageIdentityFiredEnv, nullptr, 0u) != 0u,
                L"a CREATE_NEW stage should retain exact handle authority when FileIdInfo is unsupported"))
    {
        return;
    }
    retainedWriter.reset();

    static_cast<void>(SetEnvironmentVariableW(kRenameCommittedFailurePathEnv, retainedPublishedPath.c_str()));
    publishResult           = {};
    publishResult.sizeBytes = sizeof(publishResult);
    published.reset();
    const HRESULT uncertainPublishHr =
        retainedStage->PublishAs(retainedPublishedPath.c_str(), nullptr, FILESYSTEM_FLAG_NONE, &options, &publishResult, published.put());
    if (! check(FAILED(uncertainPublishHr) && publishResult.outcomeKnown == FALSE && publishResult.mutationCommitted == FALSE && ! published &&
                    GetEnvironmentVariableW(kRenameCommittedFailureFiredEnv, nullptr, 0u) != 0u &&
                    GetFileAttributesW(retainedStagePath.c_str()) == INVALID_FILE_ATTRIBUTES &&
                    GetFileAttributesW(retainedPublishedPath.c_str()) != INVALID_FILE_ATTRIBUTES,
                L"an SMB-style committed rename without persistent identity must report Unknown, never a known non-commit"))
    {
        return;
    }
    abortResult           = {};
    abortResult.sizeBytes = sizeof(abortResult);
    if (! check(SUCCEEDED(retainedStage->AbortOwnedObject(&options, &abortResult)) && abortResult.outcomeKnown != FALSE &&
                    abortResult.mutationCommitted != FALSE && GetFileAttributesW(retainedPublishedPath.c_str()) == INVALID_FILE_ATTRIBUTES,
                L"Unknown publication should remain cleanable through the retained exact handle"))
    {
        return;
    }

    wil::com_ptr<IFileWriter> uncertainDeleteWriter;
    wil::com_ptr<IFileSystemBoundObject> uncertainDeleteStage;
    if (! check(
            SUCCEEDED(binding->CreateExclusiveWriter(uncertainDeleteStagePath.c_str(), &options, uncertainDeleteWriter.put(), uncertainDeleteStage.put())) &&
                uncertainDeleteWriter && uncertainDeleteStage && writeStage(uncertainDeleteWriter.get(), std::byte{0x74}),
            L"delete-outcome fixture should create an exact owned stage"))
    {
        return;
    }
    uncertainDeleteWriter.reset();
    static_cast<void>(SetEnvironmentVariableW(kDeleteCommittedFailurePathEnv, uncertainDeleteStagePath.c_str()));
    abortResult                    = {};
    abortResult.sizeBytes          = sizeof(abortResult);
    const HRESULT uncertainAbortHr = uncertainDeleteStage->AbortOwnedObject(&options, &abortResult);
    if (! check(FAILED(uncertainAbortHr) && abortResult.outcomeKnown == FALSE && GetEnvironmentVariableW(kDeleteCommittedFailureFiredEnv, nullptr, 0u) != 0u,
                L"a failed delete report after mutation must be Unknown, never a known retained stage"))
    {
        return;
    }
    abortResult           = {};
    abortResult.sizeBytes = sizeof(abortResult);
    if (! check(SUCCEEDED(uncertainDeleteStage->AbortOwnedObject(&options, &abortResult)) && abortResult.outcomeKnown != FALSE &&
                    abortResult.mutationCommitted != FALSE && GetFileAttributesW(uncertainDeleteStagePath.c_str()) == INVALID_FILE_ATTRIBUTES,
                L"an uncertain delete should remain retryable through the exact retained handle"))
    {
        return;
    }
}
#endif

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

    const HRESULT optionsHr = ValidateFileSystemOptions(options, FILESYSTEM_COPY);
    if (FAILED(optionsHr))
    {
        return optionsHr;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::Preserve;
    unsigned int copyMoveMaxConcurrency             = 1u;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_COPY, flags, options, callback, cookie, 1, reparsePointPolicy);
    context.objectBinding = static_cast<IFileSystemObjectBinding*>(this);

    const PathInfo source                   = MakePathInfo(sourcePath);
    const PathInfo destination              = MakePathInfo(destinationPath);
    const DWORD rootSourceAttributes        = GetFileAttributesW(source.extended.c_str());
    context.trackTopLevelOwnedStageMutation = rootSourceAttributes != INVALID_FILE_ATTRIBUTES;

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

    const unsigned int maxConcurrency = ResolveCopyMoveConcurrencyLimit(copyMoveMaxConcurrency, options, kMaxCopyMoveMaxConcurrency);
    itemHr                            = ReportTopLevelDiscovery(context, source.extended, context.recursive);
    if (SUCCEEDED(itemHr))
    {
        PathInfo selectedDestination = destination;
        for (;;)
        {
            context.reparseRootDestinationPath = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(selectedDestination.display)));
            bytesCopied                        = 0u;
            itemHr =
                CopyPathInternalWithDirectoryParallelism(context, source, selectedDestination, flags, context.reparsePointPolicy, maxConcurrency, &bytesCopied);
            ClearOneShotGrants(context);
            if (SUCCEEDED(itemHr))
            {
                break;
            }
            if (! callback || IsCancellationHr(itemHr))
            {
                break;
            }

            FileSystemItemMutationResult attemptMutation{};
            if (SnapshotTrackedItemMutation(context, attemptMutation) && FileSystemRouteContract::ClassifyFailedMutation(true, itemHr, &attemptMutation) !=
                                                                             FileSystemRouteContract::MutationClassification::RetryableNoCommit)
            {
                // Publication/cleanup truth wins over a later attributes or transport error.
                // A new conflict decision must not replay an already committed/uncertain copy.
                break;
            }

            const DWORD error = HRESULT_FACILITY(itemHr) == FACILITY_WIN32 ? static_cast<DWORD>(HRESULT_CODE(itemHr)) : ERROR_SUCCESS;
            if (error == ERROR_SUCCESS || (error != ERROR_ALREADY_EXISTS && error != ERROR_FILE_EXISTS && error != ERROR_REPARSE_POINT_ENCOUNTERED &&
                                           error != ERROR_ACCESS_DENIED && error != ERROR_DATATYPE_MISMATCH))
            {
                break;
            }

            FileSystemIssueAction action = FileSystemIssueAction::Cancel;
            const HRESULT issueHr        = ReportIssue(context, itemHr, &action);
            if (FAILED(issueHr))
            {
                itemHr = issueHr;
                break;
            }
            switch (action)
            {
                case FileSystemIssueAction::Overwrite: context.oneShotAllowOverwrite = true; continue;
                case FileSystemIssueAction::ReplaceLink:
                    context.oneShotAllowOverwrite   = true;
                    context.oneShotAllowReplaceLink = true;
                    continue;
                case FileSystemIssueAction::ReplaceReadOnly:
                    context.oneShotAllowOverwrite       = true;
                    context.oneShotAllowReplaceReadonly = true;
                    continue;
                case FileSystemIssueAction::KeepBoth:
                {
                    const DWORD sourceAttributes = GetFileAttributesW(source.extended.c_str());
                    if (sourceAttributes == INVALID_FILE_ATTRIBUTES)
                    {
                        itemHr = HRESULT_FROM_WIN32(GetLastError());
                        break;
                    }
                    const HRESULT keepBothHr = SelectKeepBothPath(destination, IsDirectory(sourceAttributes), selectedDestination);
                    if (FAILED(keepBothHr))
                    {
                        itemHr = keepBothHr;
                        break;
                    }
                    continue;
                }
                case FileSystemIssueAction::Retry: continue;
                case FileSystemIssueAction::Skip: itemHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY); break;
                case FileSystemIssueAction::PermanentDelete:
                case FileSystemIssueAction::Cancel:
                case FileSystemIssueAction::None:
                default: itemHr = HRESULT_FROM_WIN32(ERROR_CANCELLED); break;
            }
            break;
        }
    }
    if (FAILED(itemHr))
    {
        Debug::Warning(L"FileSystem: CopyItem failed for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(itemHr));
    }

    FileSystemItemMutationResult mutationResult{};
    const FileSystemItemMutationResult* reportedMutation = SnapshotTrackedItemMutation(context, mutationResult) ? &mutationResult : nullptr;
    hr                                                   = ReportItemCompleted(context, 0, itemHr, reportedMutation);
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

    const HRESULT optionsHr = ValidateFileSystemOptions(options, FILESYSTEM_MOVE);
    if (FAILED(optionsHr))
    {
        return optionsHr;
    }
    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::Preserve;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy = _reparsePointPolicy;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_MOVE, flags, options, callback, cookie, 1, reparsePointPolicy);
    context.objectBinding = static_cast<IFileSystemObjectBinding*>(this);

    const PathInfo source      = MakePathInfo(sourcePath);
    const PathInfo destination = MakePathInfo(destinationPath);

    HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        Debug::Warning(L"FileSystem: MoveItem failed to set paths for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(hr));
        return hr;
    }

    // Intra-tree links in a moved tree must be retargeted into the destination before the
    // source is deleted, exactly as CopyItem does; without these roots a moved tree ends with
    // links dangling into the deleted source.
    context.reparseRootSourcePath      = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(source.display)));
    context.reparseRootDestinationPath = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(destination.display)));

    // Local Move is native-only for every moveMode. Cross-volume and semantic-copy routes belong
    // to the host Managed engine, which owns publication and exact conditional cleanup.
    HRESULT itemHr = ReportTopLevelDiscovery(context, source.extended, false);
    if (SUCCEEDED(itemHr))
    {
        PathInfo selectedDestination = destination;
        for (;;)
        {
            itemHr = MovePathInternal(context, source, selectedDestination);
            ClearOneShotGrants(context);
            if (SUCCEEDED(itemHr) || ! callback || IsCancellationHr(itemHr))
            {
                break;
            }
            if (context.trackedNativeMove.load(std::memory_order_acquire) != TrackedPublicationTruth::NotPublished)
            {
                // A committed or unproved rename must not be offered as a fresh conflict decision.
                break;
            }

            const DWORD error = HRESULT_FACILITY(itemHr) == FACILITY_WIN32 ? static_cast<DWORD>(HRESULT_CODE(itemHr)) : ERROR_SUCCESS;
            if (error == ERROR_SUCCESS || (error != ERROR_ALREADY_EXISTS && error != ERROR_FILE_EXISTS && error != ERROR_REPARSE_POINT_ENCOUNTERED &&
                                           error != ERROR_ACCESS_DENIED && error != ERROR_DATATYPE_MISMATCH))
            {
                break;
            }

            const DWORD currentSourceAttributes      = GetFileAttributesW(source.extended.c_str());
            const DWORD currentDestinationAttributes = GetFileAttributesW(selectedDestination.extended.c_str());
            const bool regularDirectoryMergeShape    = currentSourceAttributes != INVALID_FILE_ATTRIBUTES &&
                                                       currentDestinationAttributes != INVALID_FILE_ATTRIBUTES && IsDirectory(currentSourceAttributes) &&
                                                       ! IsReparsePoint(currentSourceAttributes) && IsDirectory(currentDestinationAttributes) &&
                                                       ! IsReparsePoint(currentDestinationAttributes);
            if (regularDirectoryMergeShape)
            {
                // The host owns folder merge and same-task Native->Managed race requalification.
                // Bubble the provider's known non-commit instead of opening an item conflict from
                // inside Native Move; otherwise a direct test/task can wait forever for a decision
                // that should never be offered for the selected root.
                break;
            }

            FileSystemIssueAction action = FileSystemIssueAction::Cancel;
            const HRESULT issueHr        = ReportIssue(context, itemHr, &action);
            if (FAILED(issueHr))
            {
                itemHr = issueHr;
                break;
            }
            switch (action)
            {
                case FileSystemIssueAction::Overwrite: context.oneShotAllowOverwrite = true; continue;
                case FileSystemIssueAction::ReplaceLink:
                    context.oneShotAllowOverwrite   = true;
                    context.oneShotAllowReplaceLink = true;
                    continue;
                case FileSystemIssueAction::ReplaceReadOnly:
                    context.oneShotAllowOverwrite       = true;
                    context.oneShotAllowReplaceReadonly = true;
                    continue;
                case FileSystemIssueAction::KeepBoth:
                {
                    const DWORD sourceAttributes = GetFileAttributesW(source.extended.c_str());
                    if (sourceAttributes == INVALID_FILE_ATTRIBUTES)
                    {
                        itemHr = HRESULT_FROM_WIN32(GetLastError());
                        break;
                    }
                    const HRESULT keepBothHr = SelectKeepBothPath(destination, IsDirectory(sourceAttributes), selectedDestination);
                    if (FAILED(keepBothHr))
                    {
                        itemHr = keepBothHr;
                        break;
                    }
                    continue;
                }
                case FileSystemIssueAction::Retry: continue;
                case FileSystemIssueAction::Skip: itemHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY); break;
                case FileSystemIssueAction::PermanentDelete:
                case FileSystemIssueAction::Cancel:
                case FileSystemIssueAction::None:
                default: itemHr = HRESULT_FROM_WIN32(ERROR_CANCELLED); break;
            }
            break;
        }
    }
    if (FAILED(itemHr))
    {
        Debug::Warning(L"FileSystem: MoveItem failed for '{}' -> '{}' (hr={:#x})", source.display, destination.display, static_cast<uint32_t>(itemHr));
    }

    // MovePathInternal returns success only after the one native provider mutation committed.
    // Report that destructive truth explicitly; a null result would force the host to treat an
    // otherwise successful Native Move as indeterminate and could not authorize source removal.
    // Failures carry the same axis: a refused or failed atomic rename is a proved no-commit
    // (Retry-eligible), a committed rename followed by a later error stays committed, and an
    // unproved rename/revert outcome stays unknown so the host never replays it.
    FileSystemItemMutationResult mutationResult{};
    switch (SUCCEEDED(itemHr) ? TrackedPublicationTruth::Published : context.trackedNativeMove.load(std::memory_order_acquire))
    {
        case TrackedPublicationTruth::Published: mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, TRUE, FALSE}; break;
        case TrackedPublicationTruth::Unknown: mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, FALSE}; break;
        case TrackedPublicationTruth::NotPublished:
        case TrackedPublicationTruth::NotObserved: // Failed before any rename was issued.
        default: mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE}; break;
    }
    hr = ReportItemCompleted(context, 0, itemHr, &mutationResult);
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

    const HRESULT optionsHr = ValidateFileSystemOptions(options, FILESYSTEM_DELETE);
    if (FAILED(optionsHr))
    {
        return optionsHr;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::Preserve;
    unsigned int deleteMaxConcurrency               = 1;
    unsigned int deleteRecycleBinMaxConcurrency     = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy             = _reparsePointPolicy;
        deleteMaxConcurrency           = _deleteMaxConcurrency;
        deleteRecycleBinMaxConcurrency = _deleteRecycleBinMaxConcurrency;
    }

    OperationContext context{};
    // totalItems is 0 because recursive totals remain unknown until the operation traversal closes.
    InitializeOperationContext(context, FILESYSTEM_DELETE, flags, options, callback, cookie, 0, reparsePointPolicy);
    DeleteDiscoveryState discovery{};
    context.deleteDiscovery                     = &discovery;
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

    FileSystemItemMutationResult mutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
    HRESULT itemHr = DeletePathInternal(context, target, false, &mutationResult);
    if (FAILED(itemHr))
    {
        Debug::Warning(L"FileSystem: DeleteItem failed for '{}' (hr={:#x})", target.display, static_cast<uint32_t>(itemHr));
    }

    hr = ReportItemCompleted(context, 0, itemHr, &mutationResult);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ReportProgressForced(context, 0, 0);
    if (FAILED(hr))
    {
        return hr;
    }
    Debug::Perf::Emit(L"FileOps.DeleteTraversal.MaxDepth",
                      L"",
                      0u,
                      context.deleteTraversalMaxDepth,
                      0u, // R4-T2: no walk-depth ceiling; depth is reported only
                      itemHr);
    Debug::Perf::Emit(L"FileOps.DeleteTraversal.MaxBatchEntries", L"", 0u, context.deleteTraversalMaxBatchEntries, kDeleteTraversalBatchEntries, itemHr);
    Debug::Perf::Emit(
        L"FileOps.DeleteTraversal.MaxTerminalFailures", L"", 0u, context.deleteTraversalMaxRetainedFailureCount, kDeleteTraversalMaxTerminalFailures, itemHr);
    Debug::Perf::Emit(L"FileOps.DeleteTraversal.MaxFailurePathBytes",
                      L"",
                      0u,
                      context.deleteTraversalMaxRetainedFailurePathBytes,
                      kDeleteTraversalMaxFailurePathBytes,
                      itemHr);
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

    const HRESULT optionsHr = ValidateFileSystemOptions(options, FILESYSTEM_RENAME);
    if (FAILED(optionsHr))
    {
        return optionsHr;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::Preserve;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy = _reparsePointPolicy;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_RENAME, flags, options, callback, cookie, 1, reparsePointPolicy);
    context.objectBinding = static_cast<IFileSystemObjectBinding*>(this);

    const PathInfo source      = MakePathInfo(sourcePath);
    const PathInfo destination = MakePathInfo(destinationPath);

    HRESULT hr = SetItemPaths(context, source.display.c_str(), destination.display.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    HRESULT itemHr = MovePathInternal(context, source, destination);
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

    const HRESULT optionsHr = ValidateFileSystemOptions(options, FILESYSTEM_COPY);
    if (FAILED(optionsHr))
    {
        return optionsHr;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::Preserve;
    unsigned int copyMoveMaxConcurrency             = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    const PathInfo destinationRoot    = MakePathInfo(destinationFolder);
    const unsigned int maxConcurrency = ResolveCopyMoveConcurrencyLimit(copyMoveMaxConcurrency, options, kMaxCopyMoveMaxConcurrency);
    const unsigned int concurrency    = std::min<unsigned int>(maxConcurrency, count);

    if (concurrency <= 1u)
    {
        OperationContext context{};
        InitializeOperationContext(context, FILESYSTEM_COPY, flags, options, callback, cookie, count, reparsePointPolicy);
        context.objectBinding = static_cast<IFileSystemObjectBinding*>(this);

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
            HRESULT itemHr =
                CopyPathInternalWithDirectoryParallelism(context, source, destination, flags, context.reparsePointPolicy, maxConcurrency, &bytesCopied);

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

    auto job = GetSharedFileOpsJobScheduler().StartJob(options,
                                                       concurrency,
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
        context.objectBinding    = static_cast<IFileSystemObjectBinding*>(this);
        context.options          = &sharedOptionsState;
        context.parallel         = &parallel;
        context.totalBytes       = 0; // discovered totals are reported by this traversal
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
        HRESULT itemHr =
            CopyPathInternalWithDirectoryParallelism(context, source, destination, flags, context.reparsePointPolicy, nestedConcurrency, &bytesCopied);

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

    const HRESULT optionsHr = ValidateFileSystemOptions(options, FILESYSTEM_MOVE);
    if (FAILED(optionsHr))
    {
        return optionsHr;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::Preserve;
    unsigned int copyMoveMaxConcurrency             = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    const PathInfo destinationRoot    = MakePathInfo(destinationFolder);
    const unsigned int maxConcurrency = ResolveCopyMoveConcurrencyLimit(copyMoveMaxConcurrency, options, kMaxCopyMoveMaxConcurrency);
    const unsigned int concurrency    = std::min<unsigned int>(maxConcurrency, count);

    if (concurrency <= 1u)
    {
        OperationContext context{};
        InitializeOperationContext(context, FILESYSTEM_MOVE, flags, options, callback, cookie, count, reparsePointPolicy);
        context.objectBinding = static_cast<IFileSystemObjectBinding*>(this);

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

            HRESULT itemHr = ReportTopLevelDiscovery(context, source.extended, false);
            if (SUCCEEDED(itemHr))
            {
                itemHr = MovePathInternal(context, source, destination);
            }
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
    Debug::Perf::Emit(L"FileOps.MoveItems.NestedConcurrencyBudget", L"", nestedConcurrency, maxConcurrency, concurrency, S_OK);

    auto job = GetSharedFileOpsJobScheduler().StartJob(options,
                                                       concurrency,
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
        context.objectBinding    = static_cast<IFileSystemObjectBinding*>(this);
        context.options          = &sharedOptionsState;
        context.parallel         = &parallel;
        context.totalBytes       = 0; // discovered totals are reported by this traversal
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

        HRESULT itemHr = ReportTopLevelDiscovery(context, source.extended, false);
        if (SUCCEEDED(itemHr))
        {
            itemHr = MovePathInternal(context, source, destination);
        }

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

    const HRESULT optionsHr = ValidateFileSystemOptions(options, FILESYSTEM_DELETE);
    if (FAILED(optionsHr))
    {
        return optionsHr;
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::Preserve;
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
    DeleteDiscoveryState discovery{};

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
        // Mirror ready.size() as an atomic so the scheduler can gate dynamic-job dispatch
        // without taking scheduleMutex (kept in sync under scheduleMutex below).
        std::atomic<size_t> readyAtomic{ready.size()};

        // Per scheduler-worker-stream OperationContext, reused across the items a stream runs
        // (stable progress-stream ids); a stream never runs two steps at once, so the mutex only
        // serialises lazy init.
        constexpr size_t kDeleteStreamContextSlots = 16;
        std::array<std::optional<OperationContext>, kDeleteStreamContextSlots> deleteStreamContexts{};
        std::mutex deleteStreamContextsMutex;
        const auto getDeleteContext = [&](uint64_t streamId) -> OperationContext&
        {
            const size_t slot = streamId < deleteStreamContexts.size() ? static_cast<size_t>(streamId) : 0;
            std::scoped_lock lock(deleteStreamContextsMutex);
            if (! deleteStreamContexts[slot].has_value())
            {
                deleteStreamContexts[slot].emplace();
                OperationContext& fresh = *deleteStreamContexts[slot];
                // totalItems is 0 because recursive totals remain unknown until the operation traversal closes.
                InitializeOperationContext(fresh, FILESYSTEM_DELETE, flags, &sharedOptionsState, callback, cookie, 0, reparsePointPolicy);
                fresh.options             = &sharedOptionsState;
                fresh.parallel            = &parallel;
                fresh.totalBytes          = 0; // discovered totals are reported by this traversal
                fresh.progressStreamId    = streamId;
                fresh.recycleBinBatchSize = configuredRecycleBinBatchSize;
                fresh.deleteDiscovery     = &discovery;
            }
            return *deleteStreamContexts[slot];
        };

        // Dynamic, self-feeding job: each dispatch deletes one ready item (or one recycle-bin
        // batch) and returns the worker to the pool, so a multi-item delete no longer pins every
        // scheduler worker for its whole duration (concurrent operations keep making progress).
        auto job = GetSharedFileOpsJobScheduler().StartDynamicJob(options,
                                                                  concurrency,
                                                                  &readyAtomic,
                                                                  [&](uint64_t streamId) noexcept -> DynamicStep
        {
            if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
            {
                return DynamicStep::Finished;
            }

            OperationContext& context = getDeleteContext(streamId);

            // A step that drives remainingWork to 0 must report Finished so the job winds down;
            // otherwise (work remains) it reports Processed and the scheduler re-dispatches.
            // The error paths further down return DynamicStep::Finished DIRECTLY (not via this
            // helper): every finalizeItem()-false path first sets cancelRequested or
            // stopOnErrorRequested, so the whole operation is aborting and the job must wind down
            // at once. Claimed-but-unfinalized items are then intentionally left unprocessed —
            // exactly as the old consumer loop did when it hit a fatal/cancelled `return` — and
            // WaitJob still returns because Finished latches dynamicFinished regardless of
            // remainingWork.
            const auto finishOrContinue = [&]() noexcept -> DynamicStep
            {
                if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                {
                    return DynamicStep::Finished;
                }
                std::scoped_lock lock(scheduleMutex);
                return remainingWork == 0 ? DynamicStep::Finished : DynamicStep::Processed;
            };

            auto signalFatal = [&](HRESULT fatalHr) noexcept
            {
                parallel.stopOnErrorRequested.store(true, std::memory_order_release);
                HRESULT expected = S_OK;
                static_cast<void>(parallel.firstError.compare_exchange_strong(expected, fatalHr, std::memory_order_acq_rel));
                scheduleCv.notify_all();
            };

            auto finalizeItem = [&](unsigned long completedIndex,
                                    const PathInfo& completedTarget,
                                    HRESULT itemHr,
                                    const FileSystemItemMutationResult* mutationResult) noexcept -> bool
            {
                HRESULT hr = SetItemPaths(context, completedTarget.display.c_str(), nullptr);
                if (FAILED(hr))
                {
                    signalFatal(hr);
                    return false;
                }

                hr = ReportItemCompleted(context, completedIndex, itemHr, mutationResult);
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
                                readyAtomic.fetch_add(1, std::memory_order_release);
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

            std::vector<unsigned long> claimedIndices;
            const size_t recycleBinBatchLimit = useRecycleBin ? static_cast<size_t>(std::max(context.recycleBinBatchSize, 1u)) : 1u;
            claimedIndices.reserve(useRecycleBin ? (std::min)(static_cast<size_t>(count), recycleBinBatchLimit) : 1u);
            {
                std::unique_lock lock(scheduleMutex);
                if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                {
                    return DynamicStep::Finished;
                }

                if (remainingWork == 0)
                {
                    return DynamicStep::Finished;
                }

                if (ready.empty())
                {
                    // Transiently empty (another worker took the last ready item). Park; an
                    // in-flight worker will promote dependents or finish the last item and wake us.
                    return DynamicStep::Idle;
                }

                const unsigned long index = ready.front();
                ready.pop_front();
                readyAtomic.fetch_sub(1, std::memory_order_release);
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
                                readyAtomic.fetch_sub(1, std::memory_order_release);
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
                return DynamicStep::Idle;
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
                        return DynamicStep::Finished;
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
#if defined(ENABLE_TESTS)
                    g_recycleBinBatchTestFallbacks.fetch_add(1u, std::memory_order_relaxed);
#endif
                    for (const RecycleBinBatchEntry& entry : batchEntries)
                    {
                        FileSystemItemMutationResult mutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
                        const HRESULT itemHr = DeletePathInternal(context, entry.path, false, &mutationResult);
                        if (! finalizeItem(entry.itemIndex, entry.path, itemHr, &mutationResult))
                        {
                            return DynamicStep::Finished;
                        }
                    }

                    return finishOrContinue();
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
                    if (! finalizeItem(entry.itemIndex, entry.path, entry.result, &entry.mutationResult))
                    {
                        return DynamicStep::Finished;
                    }
                }

                if (observedCount != batchEntries.size())
                {
                    if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                    {
                        return DynamicStep::Finished;
                    }

                    if (FAILED(batchHr) && IsCancellationHr(batchHr))
                    {
                        parallel.cancelRequested.store(true, std::memory_order_release);
                        scheduleCv.notify_all();
                        return DynamicStep::Finished;
                    }

                    signalFatal(FAILED(batchHr) ? batchHr : E_UNEXPECTED);
                    return DynamicStep::Finished;
                }

                if (FAILED(batchHr))
                {
                    if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                    {
                        return DynamicStep::Finished;
                    }

                    if (IsCancellationHr(batchHr))
                    {
                        parallel.cancelRequested.store(true, std::memory_order_release);
                        scheduleCv.notify_all();
                        return DynamicStep::Finished;
                    }

                    if (! observedFailure)
                    {
                        signalFatal(batchHr);
                        return DynamicStep::Finished;
                    }
                }

                return finishOrContinue();
            }

            const unsigned long index = claimedIndices.front();
            const wchar_t* path       = paths[index];
            if (! path || path[0] == L'\0')
            {
                signalFatal(path ? E_INVALIDARG : E_POINTER);
                return DynamicStep::Finished;
            }

            const PathInfo target = MakePathInfo(path);
            FileSystemItemMutationResult mutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
            const HRESULT itemHr = DeletePathInternal(context, target, false, &mutationResult);
            if (! finalizeItem(index, target, itemHr, &mutationResult))
            {
                return DynamicStep::Finished;
            }

            return finishOrContinue();
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
    // totalItems is 0 because recursive totals remain unknown until the operation traversal closes.
    InitializeOperationContext(context, FILESYSTEM_DELETE, flags, options, callback, cookie, 0, reparsePointPolicy);
    context.deleteConcurrencyBudget = maxConcurrency;
    context.deleteDiscovery         = &discovery;

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

        FileSystemItemMutationResult mutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
        HRESULT itemHr = DeletePathInternal(context, target, false, &mutationResult);
        hr             = ReportItemCompleted(context, index, itemHr, &mutationResult);
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

    const HRESULT optionsHr = ValidateFileSystemOptions(options, FILESYSTEM_RENAME);
    if (FAILED(optionsHr))
    {
        return optionsHr;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        if (items[index].sizeBytes != sizeof(FileSystemRenamePair))
        {
            return E_INVALIDARG;
        }
    }

    FileSystemReparsePointPolicy reparsePointPolicy = FileSystemReparsePointPolicy::Preserve;
    unsigned int copyMoveMaxConcurrency             = 1;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
    }

    const unsigned int maxConcurrency = ResolveCopyMoveConcurrencyLimit(copyMoveMaxConcurrency, options, kMaxCopyMoveMaxConcurrency);
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

        auto job = GetSharedFileOpsJobScheduler().StartJob(options,
                                                           concurrency,
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
            context.objectBinding    = static_cast<IFileSystemObjectBinding*>(this);
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
                            itemHr = MovePathInternal(context, source, destination);
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
    context.objectBinding = static_cast<IFileSystemObjectBinding*>(this);

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

        HRESULT itemHr = MovePathInternal(context, source, destination);
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
