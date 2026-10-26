#include "DirectoryRecordCursor.h"
#include "FileOperationTraversalPolicy.h"
#include "FileSystem.Internal.h"
#include "FileSystemRouteContract.h"
#include "PathUtils.h"
#include "SynchronousIoCancelWatch.h"

#if defined(ENABLE_TESTS)
#include "ContentDigest.h"
#endif

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

#if defined(_DEBUG) || defined(ENABLE_TESTS)
// The object-binding debug self-test keeps its links in the marked sandbox; test builds journal the
// Recycle Bin items they create for originals inside one.
#include "Common.Tests.SandboxPath.h"
#endif

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
    Idle,      // No unit ready right now, but the job is not finished. The job is parked until woken: by one of
               // its own steps that ran a unit, or by NotifyDynamicWorkAvailable(job) after a change made outside
               // its steps.
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
        // Dynamic jobs: every wake (one of the job's own steps that ran a unit, NotifyDynamicWorkAvailable(job))
        // advances wakeGeneration; a step that found nothing to run, with no wake since its dispatch, parks the job
        // until the next one.
        uint64_t wakeGeneration   = 0;
        bool parked               = false;
        uint64_t parkedGeneration = 0;

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
    // A step that returns DynamicStep::Idle parks the job until it is woken. A step that
    // ran a unit wakes it; any change the step depends on that is made outside the job's
    // steps (a unit queued, a limit raised) must be followed by NotifyDynamicWorkAvailable(job).
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
                size_t index        = 0;
                uint64_t generation = 0;
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

                    if (! tryDequeueWorkLocked(dequeued, index, generation))
                    {
                        continue;
                    }
                }

                static_cast<void>(executeWorkItem(std::move(dequeued), index, tls_workerStreamId, generation));
            }

            return;
        }

        std::unique_lock lock(job->doneMutex);
        job->doneCv.wait(lock, [&]() noexcept { return job->done.load(std::memory_order_acquire); });
    }

    // Runs one step of a dynamic job on the calling scheduler worker when none of the job's steps is in
    // flight and it has a runnable unit. A walk's producer that waits for queue room on a scheduler
    // worker calls this: every worker of the pool can be such a producer (each root of a bulk CopyItems
    // walks on one), and the steps that would make room then have no worker left. Only the job's own
    // steps run here (one file or link each), never another job's units, so the producer's stack grows
    // by one step at most. It does not wait for a wake: a parked job's step runs here too, and unparks it.
    // Returns the step's result; Idle when the caller is not a scheduler worker, a step of the job is in
    // flight, the job is finished or nothing is ready, or the step found no unit.
    [[nodiscard]] DynamicStep TryRunUnservedDynamicStep(const JobPtr& job) noexcept
    {
        if (! job || ! job->isDynamic || ! IsWorkerThread())
        {
            return DynamicStep::Idle;
        }

        uint64_t generation = 0;
        {
            std::scoped_lock lock(_mutex);
            if (job->dynamicFinished.load(std::memory_order_acquire) || job->inFlight != 0u || ! job->readyCountPtr ||
                job->readyCountPtr->load(std::memory_order_acquire) == 0u)
            {
                return DynamicStep::Idle;
            }
            job->inFlight += 1u;
            job->parked = false;
            generation  = job->wakeGeneration;
        }

        return executeWorkItem(job, 0u, tls_workerStreamId, generation);
    }

    // Called after a change `job`'s step depends on that is made outside its steps (a unit queued, a walk's
    // listing ended, a limit raised); a step that ran a unit wakes its job anyway, and may call this earlier so
    // other workers start sooner. The job's wake generation advances under the scheduler lock, which a worker
    // holds from its check to its wait, so no worker sleeps through the change (a job's ready count changes
    // outside this lock), and the job is dispatchable again if it was parked. Other jobs stay parked.
    void NotifyDynamicWorkAvailable(const JobPtr& job) noexcept
    {
        {
            std::scoped_lock lock(_mutex);
            if (job)
            {
                ++job->wakeGeneration;
            }
        }
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
                {
                    // Under the scheduler lock, which a worker holds from its stop check to its wait: no worker
                    // sleeps through the stop.
                    std::scoped_lock workLock(_mutex);
                    for (std::jthread& worker : _workers)
                    {
                        worker.request_stop();
                    }
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
                if (job->dynamicFinished.load(std::memory_order_acquire) || dynamicJobParkedLocked(*job))
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

    // A parked dynamic job (its last step found nothing to run) waits for its next wake.
    [[nodiscard]] bool dynamicJobParkedLocked(const Job& job) const noexcept
    {
        return job.parked && job.parkedGeneration == job.wakeGeneration;
    }

    // outGeneration: the job's wake generation this dispatch saw (a dynamic step that finds nothing parks its
    // job only if no wake came since).
    [[nodiscard]] bool tryDequeueWorkLocked(JobPtr& outJob, size_t& outIndex, uint64_t& outGeneration) noexcept
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
                if (job->dynamicFinished.load(std::memory_order_acquire) || dynamicJobParkedLocked(*job))
                {
                    continue;
                }
                if (! job->readyCountPtr || job->readyCountPtr->load(std::memory_order_acquire) == 0)
                {
                    continue;
                }

                outJob        = job;
                outIndex      = 0; // Dynamic steps pull their own unit; the index is unused.
                outGeneration = job->wakeGeneration;
                job->parked   = false;
                job->inFlight += 1;

                _rrCursor = (idx + 1u) % jobCount;
                return true;
            }

            if (job->nextIndex >= job->totalItems)
            {
                continue;
            }

            outJob        = job;
            outIndex      = job->nextIndex;
            outGeneration = 0; // static units never park
            job->nextIndex += 1;
            job->inFlight += 1;

            _rrCursor = (idx + 1u) % jobCount;
            return true;
        }

        return false;
    }

    // Returns what a dynamic step reported (Processed for a static unit, Idle when no unit ran). A dynamic step
    // that ran a unit wakes its job (the unit may have made others ready, as a Delete promoting a parent does);
    // one that found nothing to run parks it when no wake came since the dispatch (`dispatchGeneration`), so
    // the job is not dispatched again before something it depends on changes.
    DynamicStep executeWorkItem(JobPtr job, size_t index, uint64_t streamId, uint64_t dispatchGeneration) noexcept
    {
        DynamicStep result = DynamicStep::Idle;
        if (job)
        {
            const Common::SynchronousIoCancelWatch::Scope cancelWatch(job->operationOptions != nullptr ? &OperationOptionsRequestCancel : nullptr,
                                                                      const_cast<FileSystemOptions*>(job->operationOptions));
            if (job->isDynamic)
            {
                if (job->processDynamic)
                {
                    result = job->processDynamic(streamId);
                }
            }
            else if (job->processIndex)
            {
                job->processIndex(index, streamId);
                result = DynamicStep::Processed;
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
                if (job->isDynamic && result == DynamicStep::Finished)
                {
                    job->dynamicFinished.store(true, std::memory_order_release);
                }
                else if (job->isDynamic && result == DynamicStep::Processed)
                {
                    ++job->wakeGeneration;
                }
                else if (job->isDynamic && result == DynamicStep::Idle && job->wakeGeneration == dispatchGeneration)
                {
                    job->parked           = true;
                    job->parkedGeneration = dispatchGeneration;
                }
            }
            cleanupJobsLocked();
        }

        _cv.notify_all();
        return result;
    }

    void workerMain(std::stop_token stopToken, uint64_t streamId) noexcept
    {
        tls_scheduler      = this;
        tls_workerStreamId = streamId;

        [[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast();

        while (! stopToken.stop_requested())
        {
            JobPtr job;
            size_t index        = 0;
            uint64_t generation = 0;

            {
                std::unique_lock lock(_mutex);
                cleanupJobsLocked();
                if (! hasSchedulableWorkLocked() && ! stopToken.stop_requested())
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
                if (! tryDequeueWorkLocked(job, index, generation))
                {
                    continue;
                }
            }

            static_cast<void>(executeWorkItem(std::move(job), index, streamId, generation));
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

void RunOnSharedFileOpsScheduler(const FileSystemOptions* options, unsigned int concurrency, size_t count, const std::function<void(size_t)>& run) noexcept
{
    if (count == 0u || ! run)
    {
        return;
    }
    SharedFileOpsJobScheduler& scheduler = GetSharedFileOpsJobScheduler();
    if (concurrency <= 1u || count == 1u || ! scheduler.EnsureWorkersAvailable())
    {
        for (size_t index = 0u; index < count; ++index)
        {
            run(index);
        }
        return;
    }
    auto job = scheduler.StartJob(options, concurrency, count, [&run](size_t index, uint64_t) { run(index); });
    scheduler.WaitJob(job);
}

[[nodiscard]] HRESULT NormalizeReparseCopyFailure(HRESULT failure, bool allowOverwrite) noexcept
{
    if (! allowOverwrite && failure == HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    return failure;
}

// Recycle must never let the Shell destroy an item. IFileOperation decides recyclability for the
// whole operation and offers each item to PreDeleteItem with TSF_DELETE_RECYCLE_IF_POSSIBLE only
// when it will move it to the Recycle Bin. Without that flag (UNC/SMB share, volume without a bin,
// item over the bin quota) it is about to delete permanently, and FOF_NOCONFIRMATION would answer
// its own "delete permanently?" question. The sink refuses such an item before any mutation.
constexpr HRESULT kRecycleUnavailableStatus = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
// The Shell reported a delete without a Recycle Bin item: the original is gone but was not recycled.
constexpr HRESULT kRecycleDestroyedStatus = COPYENGINE_E_RECYCLE_FORCE_NUKE;

struct RecycleItemVerdict
{
    HRESULT status = S_OK;
    FileSystemItemMutationResult receipt{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
};

[[nodiscard]] bool RecyclePreDeleteWouldDestroy(DWORD transferFlags) noexcept
{
    return (transferFlags & TSF_DELETE_RECYCLE_IF_POSSIBLE) == 0;
}

// Terminal truth for one Recycle item, reduced from its PostDeleteItem observation:
// - refused by PreDeleteItem: known non-commit, original present (the host's exact-item
//   RecycleEscalation may then ask for permanent-delete consent);
// - Shell failure: known non-commit, original present;
// - success with a Recycle Bin item: committed, original removed;
// - success without a Recycle Bin item (the Shell destroyed it anyway): committed, original
//   removed, but reported as a failure so it is never counted as recycled.
[[nodiscard]] RecycleItemVerdict ReduceRecyclePostDelete(bool refusedBeforeDelete, HRESULT hrDelete, bool recycleBinItemCreated) noexcept
{
    constexpr FileSystemItemMutationResult kKnownNonCommit{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
    constexpr FileSystemItemMutationResult kCommittedRemoved{sizeof(FileSystemItemMutationResult), TRUE, TRUE, FALSE};
    if (refusedBeforeDelete)
    {
        return {kRecycleUnavailableStatus, kKnownNonCommit};
    }
    if (FAILED(hrDelete))
    {
        return {hrDelete, kKnownNonCommit};
    }
    if (! recycleBinItemCreated)
    {
        return {kRecycleDestroyedStatus, kCommittedRemoved};
    }
    return {S_OK, kCommittedRemoved};
}

// A Recycle canceled because another item's failure or a fatal error stopped the call, not by the
// caller, is not a cancellation of this item. Untouched (a known non-commit), it is left unreported,
// as never reached (false). A commit the Shell observed is reported as the success it is: the stop
// reached the Shell only after the item was recycled. Any other outcome is unknown and is reported as
// such, never as a cancellation a caller would take for the user's.
[[nodiscard]] bool ReduceStopTimeRecycleCancellation(HRESULT& itemHr, const FileSystemItemMutationResult*& mutationResult) noexcept
{
    static constexpr FileSystemItemMutationResult kUnknownOutcome{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
    if (mutationResult != nullptr && mutationResult->outcomeKnown == TRUE)
    {
        if (mutationResult->mutationCommitted == FALSE && mutationResult->originalStillPresent == TRUE)
        {
            return false;
        }
        if (mutationResult->mutationCommitted == TRUE && mutationResult->originalStillPresent == FALSE)
        {
            itemHr = S_OK;
            return true;
        }
    }
    itemHr         = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
    mutationResult = &kUnknownOutcome;
    return true;
}

#if defined(_DEBUG)
// Every check in this plugin's debug self-tests goes through the shared Check so a failure is
// mirrored to stderr. The plugin's Debug::Error stream is invisible to PluginContractTests and to
// the FileOps case that runs these self-tests in-process unless an ETW listener is attached, and a
// failed contract that nobody can name cost a whole Fresh Full on 2026-09-15.
constexpr Common::DebugSelfTest::Check kFileSystemDebugCheck{L"FileSystem"};

void RunDebugReparseCopyErrorMappingSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept -> bool { return kFileSystemDebugCheck(condition, message, passed, failed); };

    constexpr HRESULT invalidParameter = HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER);
    check(NormalizeReparseCopyFailure(invalidParameter, true) == invalidParameter,
          L"staged symlink overwrite must preserve ERROR_INVALID_PARAMETER from the failed promotion");
    check(NormalizeReparseCopyFailure(invalidParameter, false) == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
          L"non-overwrite symlink copy may map ERROR_INVALID_PARAMETER to ERROR_NOT_SUPPORTED");
}

void RunDebugRecycleItemVerdictSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check     = [&](bool condition, const wchar_t* message) noexcept -> bool { return kFileSystemDebugCheck(condition, message, passed, failed); };
    const auto isReceipt = [](const FileSystemItemMutationResult& receipt, BOOL committed, BOOL originalPresent) noexcept
    {
        return receipt.sizeBytes == sizeof(FileSystemItemMutationResult) && receipt.outcomeKnown == TRUE && receipt.mutationCommitted == committed &&
               receipt.originalStillPresent == originalPresent;
    };

    // Flags observed from IFileOperation: 0x282 for a local file it recycles, 0x202 for a
    // \\localhost\C$ path it would delete permanently.
    check(! RecyclePreDeleteWouldDestroy(TSF_DELETE_RECYCLE_IF_POSSIBLE | 0x202u), L"Recycle must let the Shell move an item it can recycle");
    check(RecyclePreDeleteWouldDestroy(0x202u), L"Recycle must refuse an item the Shell would delete permanently");
    check(RecyclePreDeleteWouldDestroy(0u), L"Recycle must refuse an item offered without TSF_DELETE_RECYCLE_IF_POSSIBLE");

    RecycleItemVerdict verdict = ReduceRecyclePostDelete(true, kRecycleUnavailableStatus, false);
    check(verdict.status == kRecycleUnavailableStatus && isReceipt(verdict.receipt, FALSE, TRUE),
          L"a refused Recycle item must be a known non-commit with the original present");
    verdict = ReduceRecyclePostDelete(true, S_OK, false);
    check(verdict.status == kRecycleUnavailableStatus && isReceipt(verdict.receipt, FALSE, TRUE),
          L"a refused Recycle item stays a known non-commit whatever the Shell reports afterwards");
    verdict = ReduceRecyclePostDelete(false, E_ACCESSDENIED, false);
    check(verdict.status == E_ACCESSDENIED && isReceipt(verdict.receipt, FALSE, TRUE), L"a failed Recycle item must keep its status as a known non-commit");
    verdict = ReduceRecyclePostDelete(false, COPYENGINE_S_DONT_PROCESS_CHILDREN, true);
    check(verdict.status == S_OK && isReceipt(verdict.receipt, TRUE, FALSE), L"a recycled item must be committed with the original removed");
    verdict = ReduceRecyclePostDelete(false, S_OK, false);
    check(verdict.status == kRecycleDestroyedStatus && isReceipt(verdict.receipt, TRUE, FALSE),
          L"an item deleted without a Recycle Bin item must fail as destroyed, never count as recycled");

    // A Recycle canceled by another item's stop rather than by the caller.
    constexpr FileSystemItemMutationResult kKnownNonCommit{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
    constexpr FileSystemItemMutationResult kCommittedRemoved{sizeof(FileSystemItemMutationResult), TRUE, TRUE, FALSE};
    constexpr FileSystemItemMutationResult kUnknown{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
    HRESULT stoppedHr                                  = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    const FileSystemItemMutationResult* stoppedReceipt = &kKnownNonCommit;
    check(! ReduceStopTimeRecycleCancellation(stoppedHr, stoppedReceipt), L"a stop-time Recycle left untouched must stay unreported, as never reached");
    stoppedHr      = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    stoppedReceipt = &kCommittedRemoved;
    check(ReduceStopTimeRecycleCancellation(stoppedHr, stoppedReceipt) && stoppedHr == S_OK && stoppedReceipt == &kCommittedRemoved,
          L"a stop-time Recycle whose commit the Shell observed must be reported as the success it is, never as an unknown outcome");
    stoppedHr      = E_ABORT;
    stoppedReceipt = &kUnknown;
    check(ReduceStopTimeRecycleCancellation(stoppedHr, stoppedReceipt) && stoppedHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) && stoppedReceipt != nullptr &&
              stoppedReceipt->outcomeKnown == FALSE,
          L"a stop-time Recycle the Shell may have acted on must be reported with an unknown outcome");
    stoppedHr      = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    stoppedReceipt = nullptr;
    check(ReduceStopTimeRecycleCancellation(stoppedHr, stoppedReceipt) && stoppedHr == HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE) && stoppedReceipt != nullptr &&
              stoppedReceipt->outcomeKnown == FALSE,
          L"a stop-time Recycle without a receipt must be reported with an unknown outcome");
}

void RunDebugSharedFileOpsSchedulerShutdownSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept -> bool { return kFileSystemDebugCheck(condition, message, passed, failed); };

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

// Owner decision Q12: no test file may outlive its marked sandbox, yet a Recycle moves it into the
// user's Recycle Bin. Test-enabled builds journal each Recycle Bin item they create for an original
// inside a marked X:\RedSalamander.Perf, so the self-test that recycled it can purge exactly that
// item (RedSalamanderFileSystemTestTakeRecycledItems). Recycles of any other path are not recorded.
struct RecycledItemForTest
{
    std::wstring originalPath;
    std::wstring recycledPath; // the $R entry inside $Recycle.Bin\<SID>; empty when the Shell did not name it
};

constexpr size_t kRecycledItemJournalForTestLimit = 65'536u;
std::mutex g_recycledItemJournalForTestMutex;
std::vector<RecycledItemForTest> g_recycledItemJournalForTest;
uint64_t g_recycledItemJournalForTestDropped = 0u;

void JournalRecycledItemForTest(IShellItem* item, IShellItem* newlyCreated) noexcept
{
    if (item == nullptr || newlyCreated == nullptr)
    {
        return;
    }

    wil::unique_cotaskmem_string originalPath;
    if (FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, originalPath.put())) || ! originalPath || originalPath.get()[0] == L'\0')
    {
        return;
    }
    if (! Common::Testing::GetDedicatedExternalTestSandboxBase(std::filesystem::path(originalPath.get()), Common::Testing::kTestSandboxDirectoryName))
    {
        return;
    }

    RecycledItemForTest record{};
    record.originalPath.assign(originalPath.get());
    wil::unique_cotaskmem_string recycledPath;
    if (SUCCEEDED(newlyCreated->GetDisplayName(SIGDN_FILESYSPATH, recycledPath.put())) && recycledPath)
    {
        record.recycledPath.assign(recycledPath.get());
    }

    std::scoped_lock lock(g_recycledItemJournalForTestMutex);
    if (g_recycledItemJournalForTest.size() >= kRecycledItemJournalForTestLimit)
    {
        ++g_recycledItemJournalForTestDropped;
        return;
    }
    g_recycledItemJournalForTest.push_back(std::move(record));
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

// I35 S2 (R6.7): fails one source file's copy once, before any destination mutation, with the status
// REDSALAMANDER_FILEOPS_COPY_FAIL_STATUS names (E_FAIL when absent). The source path is consumed.
[[nodiscard]] HRESULT MaybeFailCopyBeforeMutationForSelfTest(std::wstring_view sourcePath) noexcept
{
    constexpr std::wstring_view kSource = L"REDSALAMANDER_FILEOPS_COPY_FAIL_SOURCE";
    constexpr std::wstring_view kStatus = L"REDSALAMANDER_FILEOPS_COPY_FAIL_STATUS";
    if (GetEnvironmentVariableW(kSource.data(), nullptr, 0u) == 0u)
    {
        return S_OK; // not armed: no lock on the copy path
    }
    static std::mutex consumeMutex;
    std::scoped_lock lock(consumeMutex);
    const std::wstring configured = ReadDirectFinalSelfTestPath(kSource);
    if (configured.empty() || ! OrdinalString::EqualsNoCase(configured, sourcePath))
    {
        return S_OK;
    }
    const std::wstring statusText = ReadDirectFinalSelfTestPath(kStatus);
    wchar_t* end                  = nullptr;
    const unsigned long parsed    = statusText.empty() ? 0ul : std::wcstoul(statusText.c_str(), &end, 0);
    static_cast<void>(SetEnvironmentVariableW(kSource.data(), nullptr));
    const HRESULT status = static_cast<HRESULT>(parsed);
    return FAILED(status) && end != nullptr && *end == L'\0' ? status : E_FAIL;
}

// I35 S2 (R6.3): fails a directory's listing once, after a number of its entries were handled, with
// ERROR_SHARING_VIOLATION (a failure the user can retry or skip). REDSALAMANDER_FILEOPS_LIST_FAIL_DIRECTORY
// names one or more directories separated by '|' (a character no Windows path contains) and
// REDSALAMANDER_FILEOPS_LIST_FAIL_AFTER the matching counts; a directory leaves both lists when it fails.
// A walk asks IsListingFailureArmedForSelfTest once per directory and calls this per entry only when it
// is armed.
constexpr std::wstring_view kListFailDirectoryEnvVar = L"REDSALAMANDER_FILEOPS_LIST_FAIL_DIRECTORY";

[[nodiscard]] bool IsListingFailureArmedForSelfTest() noexcept
{
    return GetEnvironmentVariableW(kListFailDirectoryEnvVar.data(), nullptr, 0u) != 0u;
}

[[nodiscard]] HRESULT MaybeFailListingForSelfTest(std::wstring_view directoryPath, uint64_t handledEntries) noexcept
{
    constexpr std::wstring_view kDirectory = kListFailDirectoryEnvVar;
    constexpr std::wstring_view kAfter     = L"REDSALAMANDER_FILEOPS_LIST_FAIL_AFTER";
    static std::mutex consumeMutex;
    std::scoped_lock lock(consumeMutex);
    const auto split = [](std::wstring_view text) noexcept
    {
        std::vector<std::wstring> parts;
        size_t start = 0u;
        for (;;)
        {
            const size_t end = text.find(L'|', start);
            parts.emplace_back(text.substr(start, end == std::wstring_view::npos ? std::wstring_view::npos : end - start));
            if (end == std::wstring_view::npos)
            {
                return parts;
            }
            start = end + 1u;
        }
    };
    std::vector<std::wstring> directories = split(ReadDirectFinalSelfTestPath(kDirectory));
    std::vector<std::wstring> afters      = split(ReadDirectFinalSelfTestPath(kAfter));
    for (size_t index = 0u; index < directories.size(); ++index)
    {
        if (directories[index].empty() || ! OrdinalString::EqualsNoCase(directories[index], directoryPath))
        {
            continue;
        }
        const unsigned long after = index < afters.size() && ! afters[index].empty() ? std::wcstoul(afters[index].c_str(), nullptr, 10) : 0ul;
        if (handledEntries < after)
        {
            return S_OK;
        }
        directories.erase(directories.begin() + static_cast<std::ptrdiff_t>(index));
        if (index < afters.size())
        {
            afters.erase(afters.begin() + static_cast<std::ptrdiff_t>(index));
        }
        const auto join = [](const std::vector<std::wstring>& parts) noexcept
        {
            std::wstring text;
            for (size_t part = 0u; part < parts.size(); ++part)
            {
                if (part != 0u)
                {
                    text += L'|';
                }
                text += parts[part];
            }
            return text;
        };
        static_cast<void>(SetEnvironmentVariableW(kDirectory.data(), directories.empty() ? nullptr : join(directories).c_str()));
        static_cast<void>(SetEnvironmentVariableW(kAfter.data(), afters.empty() ? nullptr : join(afters).c_str()));
        return HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION);
    }
    return S_OK;
}

[[nodiscard]] bool ShouldFailDirectFinalCopyForSelfTest(std::wstring_view destinationPath) noexcept
{
    const std::wstring swapPath  = ReadDirectFinalSelfTestPath(kDirectFinalRollbackSwapPathEnvVar);
    const std::wstring abortPath = ReadDirectFinalSelfTestPath(kDirectFinalAbortFailPathEnvVar);
    return (! swapPath.empty() && OrdinalString::EqualsNoCase(swapPath, destinationPath)) ||
           (! abortPath.empty() && OrdinalString::EqualsNoCase(abortPath, destinationPath));
}

// This test-only fault is admitted by one exact-case source path. It runs only
// after a successful real stage Write, with bytes still remaining, and reads
// back the retained owned stage rather than treating the source buffer as proof.
// Environment transport follows existing Local test hooks; the mutex makes
// consuming this one-shot arm atomic across provider workers. No callback or
// borrowed test-stack state is retained by the DLL.
[[nodiscard]] HRESULT MaybeFailStagedCopyAfterWriteForSelfTest(std::wstring_view sourcePath,
                                                               const std::wstring& stagePath,
                                                               IFileSystemBoundObject* ownedStage,
                                                               const FileSystemOptions* options,
                                                               uint64_t actualWrittenBytes) noexcept
{
    constexpr std::wstring_view kSource = L"REDSALAMANDER_FILEOPS_STAGE_WRITE_FAIL_SOURCE";
    constexpr std::wstring_view kFired  = L"REDSALAMANDER_FILEOPS_STAGE_WRITE_FAIL_FIRED";
    constexpr std::wstring_view kStage  = L"REDSALAMANDER_FILEOPS_STAGE_WRITE_FAIL_STAGE";
    constexpr std::wstring_view kBytes  = L"REDSALAMANDER_FILEOPS_STAGE_WRITE_FAIL_BYTES";
    constexpr std::wstring_view kDigest = L"REDSALAMANDER_FILEOPS_STAGE_WRITE_FAIL_BLAKE3";
    static std::mutex consumeMutex;
    {
        std::scoped_lock lock(consumeMutex);
        const std::wstring configured = ReadDirectFinalSelfTestPath(kSource);
        if (configured.empty() || configured != sourcePath)
        {
            return S_OK;
        }
        if (! SetEnvironmentVariableW(kSource.data(), nullptr))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
    }

    if (! ownedStage || actualWrittenBytes == 0u)
    {
        return E_UNEXPECTED;
    }
    FileSystemContentProof proof{};
    proof.sizeBytes = sizeof(proof);
    wil::com_ptr<IFileSystemBoundContentProof> contentProof;
    const HRESULT queryHr = ownedStage->QueryInterface(IID_PPV_ARGS(contentProof.addressof()));
    if (FAILED(queryHr) || ! contentProof)
    {
        return FAILED(queryHr) ? queryHr : E_UNEXPECTED;
    }
    const HRESULT proofHr = contentProof->GetContentProof(options, &proof);
    if (FAILED(proofHr))
    {
        return proofHr;
    }
    if (proof.algorithm != FILESYSTEM_CONTENT_PROOF_BLAKE3_256 || proof.contentSizeBytes != actualWrittenBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    const std::wstring bytes = std::format(L"{}", proof.contentSizeBytes);
    const std::string digest = Common::Crypto::EncodeHexDigest(std::as_bytes(std::span(proof.digest)));
    const std::wstring digestText(digest.begin(), digest.end());
    // Publish FIRED last. A partial/failed observation never passes the fixture.
    if (! SetEnvironmentVariableW(kStage.data(), stagePath.c_str()) || ! SetEnvironmentVariableW(kBytes.data(), bytes.c_str()) ||
        ! SetEnvironmentVariableW(kDigest.data(), digestText.c_str()) || ! SetEnvironmentVariableW(kFired.data(), L"1"))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    const HRESULT failure = HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
    Debug::Perf::Emit(L"FileOps.Local.StageWriteFailureInjected", stagePath, 0u, actualWrittenBytes, proof.contentSizeBytes, failure);
    return failure;
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

// One governed bulk call is one discovery scope, because the caller supplies one operation-control
// cookie for the whole call and has no way to tell one selected root's scope from another's. Each
// root's traversal still produces its own cumulative stream starting at zero, so the call adds those
// streams together here and closes exactly once, after the last root. Without this a bulk call
// reports several closures and several restarting streams, and the first root would declare the
// caller's totals final while later roots are still producing work.
struct CallDiscoveryAggregator final
{
    CallDiscoveryAggregator()                                          = default;
    CallDiscoveryAggregator(const CallDiscoveryAggregator&)            = delete;
    CallDiscoveryAggregator(CallDiscoveryAggregator&&)                 = delete;
    CallDiscoveryAggregator& operator=(const CallDiscoveryAggregator&) = delete;
    CallDiscoveryAggregator& operator=(CallDiscoveryAggregator&&)      = delete;
    ~CallDiscoveryAggregator()                                         = default;

    std::mutex mutex;
    uint64_t discoveredBytes       = 0;
    uint64_t discoveredFiles       = 0;
    uint64_t discoveredDirectories = 0;
    uint32_t queuedItems           = 0;
};

// The last cumulative values one selected root reported, so the aggregator receives deltas. A retry
// or a restarted root lowers its own cumulative value; that is never new discovery. One root owns
// one record, so the aggregator's mutex covers both.
struct RootDiscoveryTranslation final
{
    CallDiscoveryAggregator* aggregator = nullptr;
    uint64_t lastBytes                  = 0;
    uint64_t lastFiles                  = 0;
    uint64_t lastDirectories            = 0;

    void BeginRoot() noexcept
    {
        lastBytes       = 0;
        lastFiles       = 0;
        lastDirectories = 0;
    }
};

enum class TrackedPublicationTruth : uint8_t
{
    NotObserved,
    NotPublished,
    Published,
    Unknown,
};

struct DestinationMutationFacts final
{
    DestinationMutationFacts() noexcept                                  = default;
    DestinationMutationFacts(const DestinationMutationFacts&)            = delete;
    DestinationMutationFacts(DestinationMutationFacts&&)                 = delete;
    DestinationMutationFacts& operator=(const DestinationMutationFacts&) = delete;
    DestinationMutationFacts& operator=(DestinationMutationFacts&&)      = delete;

    std::atomic<bool> stageCreated{false};
    std::atomic<bool> stageRetained{false};
    std::atomic<bool> stageUnknown{false};
    std::atomic<bool> incompleteRetained{false};
    std::atomic<bool> publicationUnknown{false};
    std::atomic<bool> published{false};
    // I35 S1b: metadata outcomes of every object this item published natively, reported to the host
    // with the item's receipt (FileSystemItemMutationResult V3 suffix).
    std::atomic<uint32_t> metadataLostFeatures{0u};
    std::atomic<uint32_t> metadataChangedFeatures{0u};
    std::atomic<uint32_t> metadataInformationalFeatures{0u};
    std::atomic<uint32_t> metadataInformationalObjects{0u};
    std::atomic<HRESULT> metadataFirstFailure{S_OK};

    // Called only between selected roots, after every recursive borrower has joined.
    void Reset() noexcept
    {
        stageCreated.store(false, std::memory_order_relaxed);
        stageRetained.store(false, std::memory_order_relaxed);
        stageUnknown.store(false, std::memory_order_relaxed);
        incompleteRetained.store(false, std::memory_order_relaxed);
        publicationUnknown.store(false, std::memory_order_relaxed);
        published.store(false, std::memory_order_relaxed);
        metadataLostFeatures.store(0u, std::memory_order_relaxed);
        metadataChangedFeatures.store(0u, std::memory_order_relaxed);
        metadataInformationalFeatures.store(0u, std::memory_order_relaxed);
        metadataInformationalObjects.store(0u, std::memory_order_relaxed);
        metadataFirstFailure.store(S_OK, std::memory_order_relaxed);
    }
};

// Provider-local receipt policy: only this execution layer knows when a queued leaf's retry
// loop ends and when a retained Local object has passed its compensating-abort boundary.
struct NativeFileWorkCallState final
{
    NativeFileWorkCallState()                                          = default;
    NativeFileWorkCallState(const NativeFileWorkCallState&)            = delete;
    NativeFileWorkCallState(NativeFileWorkCallState&&)                 = delete;
    NativeFileWorkCallState& operator=(const NativeFileWorkCallState&) = delete;
    NativeFileWorkCallState& operator=(NativeFileWorkCallState&&)      = delete;
    ~NativeFileWorkCallState()                                         = default;

    std::atomic<uint64_t> nextLogicalLeafId{1u};
    std::atomic<bool> published{false};
    std::atomic<bool> countsAvailable{true};
};

struct NativeFileWorkLeafState final
{
    const PathInfo* source      = nullptr;
    const PathInfo* destination = nullptr;
    uint64_t logicalLeafId      = 0u;
    uint64_t progressStreamId   = 0u;
    FileSystemFileWorkKind kind = FileSystemFileWorkKind::None;
    uint64_t totalBytes         = 0u;
    uint64_t completedBytes     = 0u;
    bool sizeKnown              = false;
    bool began                  = false;
    bool publicationEmitted     = false;
    bool skipped                = false;
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
    unsigned int deleteConcurrencyBudget  = 1;
    unsigned int recycleBinBatchSize      = 1;
    DeleteDiscoveryState* deleteDiscovery = nullptr;
    // Set only by a bulk entry point, where several selected roots share one governed call scope.
    RootDiscoveryTranslation* discoveryTranslation = nullptr;
    NativeFileWorkCallState ownedFileWorkCall;
    NativeFileWorkCallState* fileWorkCall               = &ownedFileWorkCall;
    NativeFileWorkLeafState* fileWorkLeaf               = nullptr; // borrowed only for the outer logical-leaf retry scope
    std::mutex* fileWorkCallbackMutex                   = nullptr; // a producer borrows its walk's callback lock without changing its byte-accounting context
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
    // The exact top-level destination the tracked fields describe when a Keep Both retry moved
    // it off itemDestination; null tracks itemDestination. Points into the CopyItem loop's path.
    const wchar_t* trackedItemDestination = nullptr;
    std::atomic<TrackedPublicationTruth> trackedPublication{TrackedPublicationTruth::NotObserved};
    std::atomic<FileSystemOwnedStageDisposition> trackedOwnedStageDisposition{FileSystemOwnedStageDisposition::NotApplicable};
    // Whole-item destination truth for the receipt when no exact top-level stage was
    // observed: a merge into an existing directory, a failure before any stage, or a
    // subtree whose children own their own stages. Every owned-stage site records here
    // unconditionally; only the exact top-level destination also updates the fields above.
    DestinationMutationFacts ownedDestinationMutation;
    // Recursive copy workers borrow this item's facts. The scheduler joins all children
    // before the owning root context is destroyed or its terminal receipt is read.
    DestinationMutationFacts* destinationMutation = &ownedDestinationMutation;
    // Native Move truth (MovePathInternal only): NotObserved until the first attempt,
    // NotPublished while every issued rename is proved not committed, Published after the
    // one native mutation committed, Unknown when a rename/revert outcome is unproved. A source that
    // is missing before any rename stays NotPublished: this item issued nothing, so it is a known failure.
    // Monotonic: Unknown/Published survive a later attempt so a Retry can never erase them.
    std::atomic<TrackedPublicationTruth> trackedNativeMove{TrackedPublicationTruth::NotObserved};
};
#pragma warning(pop)

bool IsCancellationHr(HRESULT hr) noexcept;
[[nodiscard]] bool IsReparsePoint(DWORD attributes) noexcept;
[[nodiscard]] bool IsDirectory(DWORD attributes) noexcept;

void ReportNativeFileWork(OperationContext& context,
                          FileSystemFileWorkPhase phase,
                          FileSystemFileWorkDisposition disposition = FileSystemFileWorkDisposition::None,
                          HRESULT status                            = S_OK) noexcept
{
    if ((context.type != FILESYSTEM_COPY && context.type != FILESYSTEM_DELETE) || context.options == nullptr || context.options->fileWorkCallback == nullptr)
    {
        return;
    }
    if (context.type == FILESYSTEM_DELETE && context.useRecycleBin)
    {
        // Shell Recycle has its own progress and terminal receipts, but no exact per-object
        // completion observer. Never certify a known zero while the Shell removes items.
        if (phase == FileSystemFileWorkPhase::ScopeSupported)
            phase = FileSystemFileWorkPhase::CountsUnavailable;
        else if (phase != FileSystemFileWorkPhase::CountsUnavailable)
            return;
    }

    FileSystemFileWorkProgress report{};
    report.sizeBytes       = sizeof(report);
    report.phase           = phase;
    report.disposition     = disposition;
    report.status          = status;
    report.scopeBytesKnown = TRUE; // Win32 listings supply logical file sizes; literal links contribute zero payload bytes.
    if (phase == FileSystemFileWorkPhase::ScopeSupported)
        report.kind = FileSystemFileWorkKind::Directory; // this Local traversal also reports completed ordinary directories
    if (phase != FileSystemFileWorkPhase::ScopeSupported && phase != FileSystemFileWorkPhase::CountsUnavailable)
    {
        const NativeFileWorkLeafState* leaf = context.fileWorkLeaf;
        if (leaf == nullptr || ! leaf->began)
        {
            return;
        }
        report.logicalLeafId    = leaf->logicalLeafId;
        report.progressStreamId = leaf->progressStreamId;
        report.sourcePath       = leaf->source->display.c_str();
        report.destinationPath  = leaf->destination->display.c_str();
        report.kind             = leaf->kind;
        report.sizeKnown        = leaf->sizeKnown ? TRUE : FALSE;
        report.totalBytes       = leaf->totalBytes;
        report.completedBytes   = leaf->completedBytes;
    }

    const auto emit = [&]() noexcept
    {
        // This observer cannot pause/cancel work or revoke a committed publication.
        static_cast<void>(context.options->fileWorkCallback->FileSystemReportFileWork(&report, context.options->fileWorkCookie));
    };
    std::mutex* callbackMutex = context.fileWorkCallbackMutex;
    if (callbackMutex == nullptr && context.parallel != nullptr)
    {
        callbackMutex = &context.parallel->callbackMutex;
    }
    if (callbackMutex != nullptr)
    {
        std::scoped_lock lock(*callbackMutex);
        emit();
    }
    else
    {
        emit();
    }
}

void ReportCompletedObject(OperationContext& context, const wchar_t* sourcePath, DWORD attributes, uint64_t bytes) noexcept
{
    if (context.options == nullptr || context.options->fileWorkCallback == nullptr)
        return;
    FileSystemFileWorkProgress report{};
    report.sizeBytes   = sizeof(report);
    report.phase       = FileSystemFileWorkPhase::ObjectCompleted;
    report.disposition = FileSystemFileWorkDisposition::Completed;
    report.kind        = IsDirectory(attributes) ? FileSystemFileWorkKind::Directory : FileSystemFileWorkKind::RegularFile;
    if (IsReparsePoint(attributes))
        report.kind = IsDirectory(attributes) ? FileSystemFileWorkKind::DirectoryLink : FileSystemFileWorkKind::FileLink;
    report.sourcePath       = sourcePath;
    report.progressStreamId = context.progressStreamId;
    report.sizeKnown        = TRUE;
    report.totalBytes       = IsDirectory(attributes) || IsReparsePoint(attributes) ? 0u : bytes;
    report.completedBytes   = report.totalBytes;
    if (context.type == FILESYSTEM_COPY)
        context.fileWorkCall->published.store(true, std::memory_order_release); // a replay must also withdraw completed-folder certainty
    const auto emit = [&]() noexcept
    {
        // Reporting a committed object is observational; cancellation stays in operationControl.
        static_cast<void>(context.options->fileWorkCallback->FileSystemReportFileWork(&report, context.options->fileWorkCookie));
    };
    std::mutex* callbackMutex = context.fileWorkCallbackMutex;
    if (callbackMutex == nullptr && context.parallel != nullptr)
        callbackMutex = &context.parallel->callbackMutex;
    if (callbackMutex != nullptr)
    {
        std::scoped_lock lock(*callbackMutex);
        emit();
    }
    else
        emit();
}

void WithdrawNativeFileWorkCounts(OperationContext& context) noexcept
{
    if (context.fileWorkCall->countsAvailable.exchange(false, std::memory_order_acq_rel))
    {
        ReportNativeFileWork(context, FileSystemFileWorkPhase::CountsUnavailable);
    }
}

#pragma warning(push)
#pragma warning(disable : 4820) // ABI observation receipt keeps enum, byte counters and flags together.
class NativeFileWorkScope final
{
public:
    NativeFileWorkScope(OperationContext& context, const PathInfo& source, const PathInfo& destination, bool bind = true) noexcept
        : _context(context),
          _previous(context.fileWorkLeaf)
    {
        if (! bind || context.type != FILESYSTEM_COPY || context.options == nullptr || context.options->fileWorkCallback == nullptr)
        {
            return;
        }
        _leaf.source           = &source;
        _leaf.destination      = &destination;
        _leaf.logicalLeafId    = context.fileWorkCall->nextLogicalLeafId.fetch_add(1u, std::memory_order_relaxed);
        _leaf.progressStreamId = context.progressStreamId;
        _context.fileWorkLeaf  = &_leaf;
        _bound                 = true;
    }
    NativeFileWorkScope(const NativeFileWorkScope&)            = delete;
    NativeFileWorkScope(NativeFileWorkScope&&)                 = delete;
    NativeFileWorkScope& operator=(const NativeFileWorkScope&) = delete;
    NativeFileWorkScope& operator=(NativeFileWorkScope&&)      = delete;

    ~NativeFileWorkScope() noexcept
    {
        Finish();
    }

    void Finish() noexcept
    {
        if (! _bound)
        {
            return;
        }
        const FileSystemFileWorkDisposition disposition = _leaf.skipped               ? FileSystemFileWorkDisposition::Skipped
                                                          : IsCancellationHr(_status) ? FileSystemFileWorkDisposition::Canceled
                                                          : FAILED(_status)           ? FileSystemFileWorkDisposition::Failed
                                                                                      : FileSystemFileWorkDisposition::Completed;
        ReportNativeFileWork(_context, FileSystemFileWorkPhase::Terminal, disposition, _status);
        _context.fileWorkLeaf = _previous;
        _bound                = false;
    }

    void SetResult(HRESULT status) noexcept
    {
        _status = status;
    }
    void SetSkipped() noexcept
    {
        _leaf.skipped = true;
    }
    void SetProducerStream() noexcept
    {
        _leaf.progressStreamId = FILESYSTEM_PRODUCER_PROGRESS_STREAM_ID;
    }

private:
    OperationContext& _context;
    NativeFileWorkLeafState* _previous = nullptr;
    NativeFileWorkLeafState _leaf;
    HRESULT _status = E_UNEXPECTED;
    bool _bound     = false;
};
#pragma warning(pop)

void BeginNativeFileWork(OperationContext& context, FileSystemFileWorkKind kind, bool sizeKnown, uint64_t totalBytes) noexcept
{
    NativeFileWorkLeafState* leaf = context.fileWorkLeaf;
    if (leaf == nullptr)
    {
        return;
    }
    leaf->kind           = kind;
    leaf->sizeKnown      = sizeKnown;
    leaf->totalBytes     = totalBytes;
    leaf->completedBytes = 0u;
    leaf->began          = true;
    ReportNativeFileWork(context, FileSystemFileWorkPhase::Begin);
}

void MarkNativeMoveSourceUnproven(OperationContext& context) noexcept
{
    TrackedPublicationTruth observed = context.trackedNativeMove.load(std::memory_order_acquire);
    while (observed != TrackedPublicationTruth::Published && observed != TrackedPublicationTruth::Unknown &&
           ! context.trackedNativeMove.compare_exchange_weak(observed, TrackedPublicationTruth::Unknown, std::memory_order_acq_rel))
    {
    }
}

void BeginCopyItemMutation(OperationContext& context, const PathInfo& source) noexcept
{
    context.ownedDestinationMutation.Reset();
    context.destinationMutation = &context.ownedDestinationMutation;
    context.trackedPublication.store(TrackedPublicationTruth::NotObserved, std::memory_order_relaxed);
    context.trackedOwnedStageDisposition.store(FileSystemOwnedStageDisposition::NotApplicable, std::memory_order_relaxed);
    context.trackedItemDestination          = nullptr;
    context.trackTopLevelOwnedStageMutation = GetFileAttributesW(source.extended.c_str()) != INVALID_FILE_ATTRIBUTES;
}

[[nodiscard]] FileSystemItemMutationResult SnapshotNativeMoveMutation(const OperationContext& context, HRESULT status) noexcept
{
    // Success is returned only after the native rename committed. On failure, keep
    // the strongest recorded rename/revert evidence. A failed source probe or an
    // atomic rename refusal (a vanished source, a missing destination parent) issued
    // no mutation, so it stays a known no-commit with the real error.
    switch (SUCCEEDED(status) ? TrackedPublicationTruth::Published : context.trackedNativeMove.load(std::memory_order_acquire))
    {
        case TrackedPublicationTruth::Published: return FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, TRUE, FALSE};
        case TrackedPublicationTruth::Unknown: return FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, FALSE};
        case TrackedPublicationTruth::NotPublished:
        case TrackedPublicationTruth::NotObserved:
        default: return FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
    }
}

// I35 S1b: one published object's metadata outcome (both transfer phases combined), folded into the
// item's receipt so the host reports native Copy metadata exactly as it reports the bridge's.
void RecordNativeMetadataOutcome(OperationContext& context, const FileSystemMetadataTransferResult& transfer) noexcept
{
    DestinationMutationFacts& facts = *context.destinationMutation;
    if (transfer.lostFeatures != 0u)
    {
        facts.metadataLostFeatures.fetch_or(transfer.lostFeatures, std::memory_order_acq_rel);
    }
    if (transfer.changedFeatures != 0u)
    {
        facts.metadataChangedFeatures.fetch_or(transfer.changedFeatures, std::memory_order_acq_rel);
    }
    if (transfer.informationalFeatures != 0u)
    {
        facts.metadataInformationalFeatures.fetch_or(transfer.informationalFeatures, std::memory_order_acq_rel);
        facts.metadataInformationalObjects.fetch_add(1u, std::memory_order_acq_rel);
    }
    if (FAILED(transfer.firstFailure))
    {
        HRESULT expected = S_OK;
        static_cast<void>(facts.metadataFirstFailure.compare_exchange_strong(expected, transfer.firstFailure, std::memory_order_acq_rel));
    }
}

[[nodiscard]] FileSystemMetadataTransferResult CombineMetadataTransferPhases(const FileSystemMetadataTransferResult& prepare,
                                                                             const FileSystemMetadataTransferResult& finalize) noexcept
{
    FileSystemMetadataTransferResult combined = finalize;
    combined.attemptedFeatures |= prepare.attemptedFeatures;
    combined.preservedFeatures |= prepare.preservedFeatures;
    combined.changedFeatures |= prepare.changedFeatures;
    combined.lostFeatures |= prepare.lostFeatures;
    combined.informationalFeatures |= prepare.informationalFeatures;
    if (FAILED(prepare.firstFailure))
    {
        combined.firstFailure = prepare.firstFailure;
    }
    return combined;
}

void AttachNativeMetadataOutcome(const DestinationMutationFacts& facts, FileSystemItemMutationResult& result) noexcept
{
    result.metadataLostFeatures             = facts.metadataLostFeatures.load(std::memory_order_acquire);
    result.metadataChangedFeatures          = facts.metadataChangedFeatures.load(std::memory_order_acquire);
    result.metadataInformationalFeatures    = facts.metadataInformationalFeatures.load(std::memory_order_acquire);
    result.metadataInformationalObjectCount = facts.metadataInformationalObjects.load(std::memory_order_acquire);
    result.metadataFirstFailure             = facts.metadataFirstFailure.load(std::memory_order_acquire);
}

[[nodiscard]] bool SnapshotTrackedItemMutationAxes(const OperationContext& context, FileSystemItemMutationResult& result) noexcept;

// The receipt for a Copy item: its mutation axes plus the metadata its published objects lost,
// changed or inherited.
[[nodiscard]] bool SnapshotTrackedItemMutation(const OperationContext& context, FileSystemItemMutationResult& result) noexcept
{
    if (! SnapshotTrackedItemMutationAxes(context, result))
    {
        return false;
    }
    AttachNativeMetadataOutcome(*context.destinationMutation, result);
    return true;
}

// The receipt axes for a Copy item. An exact tracked top-level stage wins; otherwise the
// whole-item aggregates decide. They are recorded at every owned-stage site regardless
// of top-level tracking, so a failure that never reached a stage is a proved no-commit.
[[nodiscard]] bool SnapshotTrackedItemMutationAxes(const OperationContext& context, FileSystemItemMutationResult& result) noexcept
{
    const DestinationMutationFacts& facts = *context.destinationMutation;
    const bool publicationUnknown         = facts.publicationUnknown.load(std::memory_order_acquire);
    const FileSystemOwnedStageDisposition cleanupDebt =
        facts.stageUnknown.load(std::memory_order_acquire) || publicationUnknown ? FileSystemOwnedStageDisposition::Unknown
        : facts.incompleteRetained.load(std::memory_order_acquire)               ? FileSystemOwnedStageDisposition::RetainedIncomplete
        : facts.stageRetained.load(std::memory_order_acquire)                    ? FileSystemOwnedStageDisposition::Retained
                                                                                 : FileSystemOwnedStageDisposition::NotApplicable;
    const TrackedPublicationTruth publication = context.trackedPublication.load(std::memory_order_acquire);
    if (publication != TrackedPublicationTruth::NotObserved)
    {
        result = FileSystemItemMutationResult{
            sizeof(FileSystemItemMutationResult),
            publication == TrackedPublicationTruth::Unknown ? FALSE : TRUE,
            publication == TrackedPublicationTruth::Published ? TRUE : FALSE,
            TRUE,
            cleanupDebt != FileSystemOwnedStageDisposition::NotApplicable ? cleanupDebt : context.trackedOwnedStageDisposition.load(std::memory_order_acquire),
        };
        return true;
    }
    // Copy never removes its source, so originalStillPresent is TRUE on every axis below.
    const bool published = facts.published.load(std::memory_order_acquire);
    if (publicationUnknown && ! published)
    {
        result = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE, FileSystemOwnedStageDisposition::Unknown};
        return true;
    }
    if (published)
    {
        // At least one destination object under this item was published: the destination
        // changed, even when a later child failed or the user canceled.
        // A proven child publication remains a committed destination even if another child
        // has an uncertain publication or surviving stage. Cleanup debt is a separate axis.
        result = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, TRUE, TRUE, cleanupDebt};
        return true;
    }
    if (cleanupDebt != FileSystemOwnedStageDisposition::NotApplicable)
    {
        result = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE, cleanupDebt};
        return true;
    }
    // No destination object was published and every created stage was removed (or none
    // was ever created): the attempt is a proved no-commit and may be retried.
    result = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult),
                                          TRUE,
                                          FALSE,
                                          TRUE,
                                          facts.stageCreated.load(std::memory_order_acquire) ? FileSystemOwnedStageDisposition::Removed
                                                                                             : FileSystemOwnedStageDisposition::NotCreated};
    return true;
}

[[nodiscard]] bool ShouldTrackOwnedStageMutation(const OperationContext& context, const PathInfo& destination) noexcept
{
    const wchar_t* const tracked = context.trackedItemDestination != nullptr ? context.trackedItemDestination : context.itemDestination;
    return context.trackTopLevelOwnedStageMutation && tracked != nullptr && OrdinalString::EqualsNoCase(tracked, destination.display);
}

// Keep Both moves the top-level Copy to a new sibling. The earlier attempt's tracked stage
// described the original name, so the receipt must describe the destination actually written.
void RetargetTrackedCopyDestination(OperationContext& context, const PathInfo& selectedDestination) noexcept
{
    context.trackedPublication.store(TrackedPublicationTruth::NotObserved, std::memory_order_release);
    context.trackedOwnedStageDisposition.store(FileSystemOwnedStageDisposition::NotApplicable, std::memory_order_release);
    context.trackedItemDestination = selectedDestination.display.c_str();
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
        case TrackedPublicationTruth::Published: context.destinationMutation->published.store(true, std::memory_order_release); break;
        case TrackedPublicationTruth::Unknown: context.destinationMutation->publicationUnknown.store(true, std::memory_order_release); break;
        case TrackedPublicationTruth::NotPublished: context.destinationMutation->stageCreated.store(true, std::memory_order_release); break;
        case TrackedPublicationTruth::NotObserved:
        default: break;
    }
    if (ShouldTrackOwnedStageMutation(context, destination))
    {
        context.trackedPublication.store(publication, std::memory_order_release);
        context.trackedOwnedStageDisposition.store(disposition, std::memory_order_release);
    }
    NativeFileWorkLeafState* leaf = context.fileWorkLeaf;
    if (publication == TrackedPublicationTruth::Published && leaf != nullptr && leaf->began && ! leaf->publicationEmitted)
    {
        leaf->publicationEmitted = true;
        if (leaf->kind != FileSystemFileWorkKind::DirectoryLink)
        {
            context.fileWorkCall->published.store(true, std::memory_order_release);
        }
        ReportNativeFileWork(context, FileSystemFileWorkPhase::Published);
    }
}

// Records the final disposition of an owned stage after its abort/cleanup ran.
void RecordOwnedStageDisposition(OperationContext& context, const PathInfo& destination, FileSystemOwnedStageDisposition disposition) noexcept
{
    switch (disposition)
    {
        case FileSystemOwnedStageDisposition::Retained: context.destinationMutation->stageRetained.store(true, std::memory_order_release); break;
        case FileSystemOwnedStageDisposition::Unknown: context.destinationMutation->stageUnknown.store(true, std::memory_order_release); break;
        case FileSystemOwnedStageDisposition::RetainedIncomplete: context.destinationMutation->incompleteRetained.store(true, std::memory_order_release); break;
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

unsigned int ResolveCopyMoveConcurrencyLimit(unsigned int configuredConcurrency,
                                             const FileSystemOptions* options,
                                             unsigned int maxConcurrency,
                                             FileSystemConcurrencyMode concurrencyMode = FileSystemConcurrencyMode::Manual) noexcept
{
    unsigned int effective = std::clamp(configuredConcurrency, 1u, maxConcurrency);
    if (options != nullptr && options->copyMoveMaxConcurrency != 0u)
    {
        const unsigned int requested = std::clamp(options->copyMoveMaxConcurrency, 1u, maxConcurrency);
        // Auto is resolved by the host from both endpoints. The inactive Manual value must not
        // re-clamp recursive contents while the equivalent selected leaves use the resolved budget.
        effective = concurrencyMode == FileSystemConcurrencyMode::Auto ? requested : std::min(effective, requested);
    }
    return effective;
}

constexpr uint64_t kBandwidthThrottleBurstWindowMs    = 250ull;
constexpr uint64_t kBandwidthThrottleMinCapacityBytes = 64ull * 1024ull;
constexpr uint64_t kBandwidthThrottleMaxCapacityBytes = 4ull * 1024ull * 1024ull;
constexpr uint64_t kBandwidthThrottleBoundaryGuardMs  = 10ull;
constexpr DWORD kBandwidthThrottleCancelPollMs        = 10u;
#ifdef ENABLE_TESTS
constexpr std::wstring_view kBandwidthThrottleWorkerModeEnvVar    = L"REDSALAMANDER_FILEOPS_BW_WORKER_MODE";
constexpr std::wstring_view kRecycleFailurePathEnvVar             = L"REDSALAMANDER_FILEOPS_RECYCLE_FAIL_PATH";
constexpr std::wstring_view kRecycleFailureDirectoryEnvVar        = L"REDSALAMANDER_FILEOPS_RECYCLE_FAIL_DIRECTORY";
constexpr std::wstring_view kRecycleBatchUnknownPathEnvVar        = L"REDSALAMANDER_FILEOPS_RECYCLE_BATCH_UNKNOWN_PATH";
constexpr std::wstring_view kRecycleBatchCommitPathEnvVar         = L"REDSALAMANDER_FILEOPS_RECYCLE_BATCH_COMMIT_PATH";
constexpr std::wstring_view kRecycleSiblingHoldPathEnvVar         = L"REDSALAMANDER_FILEOPS_RECYCLE_SIBLING_HOLD_PATH";
constexpr std::wstring_view kRecycleSiblingStopSeenEnvVar         = L"REDSALAMANDER_FILEOPS_RECYCLE_SIBLING_STOP_SEEN";
constexpr std::wstring_view kRecycleShellStopPathEnvVar           = L"REDSALAMANDER_FILEOPS_RECYCLE_SHELL_STOP_PATH";
constexpr std::wstring_view kRecycleShellCommitStopPathEnvVar     = L"REDSALAMANDER_FILEOPS_RECYCLE_SHELL_COMMIT_STOP_PATH";
constexpr std::wstring_view kRecycleUnknownPathEnvVar             = L"REDSALAMANDER_FILEOPS_RECYCLE_UNKNOWN_PATH";
constexpr std::wstring_view kRecycleUnknownHoldEnvVar             = L"REDSALAMANDER_FILEOPS_RECYCLE_UNKNOWN_HOLD";
constexpr std::wstring_view kRecycleDestroyPathEnvVar             = L"REDSALAMANDER_FILEOPS_RECYCLE_DESTROY_PATH";
constexpr std::wstring_view kCaseRenameEntropyFailurePathEnvVar   = L"REDSALAMANDER_FILEOPS_CASE_RENAME_ENTROPY_FAIL_PATH";
constexpr std::wstring_view kCaseRenameTempCollisionPathEnvVar    = L"REDSALAMANDER_FILEOPS_CASE_RENAME_TEMP_COLLISION_PATH";
constexpr std::wstring_view kCaseRenameTempCollisionFiredEnvVar   = L"REDSALAMANDER_FILEOPS_CASE_RENAME_TEMP_COLLISION_FIRED";
constexpr std::wstring_view kCaseRenameSecondLegFailurePathEnvVar = L"REDSALAMANDER_FILEOPS_CASE_RENAME_SECOND_LEG_FAIL_PATH";
constexpr std::wstring_view kCaseRenameFirstLegPathEnvVar         = L"REDSALAMANDER_FILEOPS_CASE_RENAME_FIRST_LEG_PATH";
constexpr std::wstring_view kCaseRenameRevertedEnvVar             = L"REDSALAMANDER_FILEOPS_CASE_RENAME_REVERTED";
constexpr std::wstring_view kMoveDestinationRacePathEnvVar        = L"REDSALAMANDER_FILEOPS_MOVE_DESTINATION_RACE_PATH";
constexpr std::wstring_view kMoveDestinationRaceFiredEnvVar       = L"REDSALAMANDER_FILEOPS_MOVE_DESTINATION_RACE_FIRED";
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

// A Recycle seam variable holds one path or a '|'-separated list of paths. Each listed path fires
// once: a consuming match removes it (and clears the variable with the last entry). Concurrent
// Recycle workers share one lock, so two items never consume or rewrite the list at once.
[[nodiscard]] bool MatchRecycleSeamPathForSelfTest(std::wstring_view envVarName, const PathInfo& path, bool consume) noexcept
{
    static std::mutex seamMutex;
    std::scoped_lock lock(seamMutex);
    const std::wstring list = GetSelfTestEnvironmentString(envVarName);
    if (list.empty())
    {
        return false;
    }

    bool matched = false;
    std::wstring remaining;
    for (size_t start = 0u; start <= list.size();)
    {
        const size_t end              = std::min(list.find(L'|', start), list.size());
        const std::wstring_view entry = std::wstring_view(list).substr(start, end - start);
        if (! matched && ! entry.empty() && OrdinalString::EqualsNoCase(path.extended, MakePathInfo(std::wstring(entry)).extended))
        {
            matched = true;
        }
        else if (! entry.empty())
        {
            if (! remaining.empty())
            {
                remaining.push_back(L'|');
            }
            remaining.append(entry);
        }
        start = end + 1u;
    }
    if (matched && consume)
    {
        static_cast<void>(::SetEnvironmentVariableW(envVarName.data(), remaining.empty() ? nullptr : remaining.c_str()));
    }
    return matched;
}

// The directory form arms every item directly inside one directory, for as long as it is set, so
// a scale test can fail thousands of Recycles without a per-path list.
[[nodiscard]] bool MatchRecycleFailureDirectoryForSelfTest(const PathInfo& path) noexcept
{
    const std::wstring directory = GetSelfTestEnvironmentString(kRecycleFailureDirectoryEnvVar);
    return ! directory.empty() && OrdinalString::EqualsNoCase(GetPathDirectory(path.extended), MakePathInfo(directory).extended);
}

[[nodiscard]] bool TryInjectRecycleFailureForSelfTest(const PathInfo& path, FileSystemItemMutationResult* mutationResult) noexcept
{
    if (! MatchRecycleFailureDirectoryForSelfTest(path) && ! MatchRecycleSeamPathForSelfTest(kRecycleFailurePathEnvVar, path, true))
    {
        return false;
    }

    if (mutationResult != nullptr)
    {
        *mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
    }
    return true;
}

// Stands in for a Shell that loses an item's terminal truth: the item is untouched and the Recycle
// route reports an unknown outcome. While the hold variable is set the call waits (bounded), so a
// test can let another item of the same task reach its escalation first.
[[nodiscard]] bool TryInjectRecycleUnknownForSelfTest(const PathInfo& path, HRESULT& status, FileSystemItemMutationResult* mutationResult) noexcept
{
    if (! MatchRecycleSeamPathForSelfTest(kRecycleUnknownPathEnvVar, path, true))
    {
        return false;
    }

    const ULONGLONG holdDeadline = ::GetTickCount64() + 30'000ull;
    while (! GetSelfTestEnvironmentString(kRecycleUnknownHoldEnvVar).empty() && ::GetTickCount64() < holdDeadline)
    {
        ::Sleep(10);
    }
    status = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
    if (mutationResult != nullptr)
    {
        *mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
    }
    return true;
}

// Waits (bounded) until another item's stop or a cancellation ends a parallel provider call, and
// records that the stop arrived while this item was in flight.
void AwaitParallelStopForSelfTest(const ParallelOperationState* parallel) noexcept
{
    if (parallel == nullptr)
    {
        return;
    }
    const ULONGLONG deadline = ::GetTickCount64() + 10'000ull;
    while (! parallel->stopOnErrorRequested.load(std::memory_order_acquire) && ! parallel->cancelRequested.load(std::memory_order_acquire) &&
           ::GetTickCount64() < deadline)
    {
        ::Sleep(1);
    }
    if (parallel->stopOnErrorRequested.load(std::memory_order_acquire))
    {
        static_cast<void>(::SetEnvironmentVariableW(kRecycleSiblingStopSeenEnvVar.data(), L"1"));
    }
}

// Puts a sibling worker's Recycle in flight when another item's failure stops the call: the held
// item (fires once) waits for that stop before its own cancellation check, and while it, a
// committing batch (DeleteToRecycleBinBatched) or an item armed to see the stop inside the Shell
// (DeleteToRecycleBinCore) is still armed, every other Recycle item waits (bounded) so its failure
// cannot arrive first. An item armed for the Shell runs on to its own seam there.
void HoldRecycleForSiblingStopForSelfTest(OperationContext& context, const PathInfo& path) noexcept
{
    const auto anyArmed = []() noexcept
    {
        return ! GetSelfTestEnvironmentString(kRecycleSiblingHoldPathEnvVar).empty() || ! GetSelfTestEnvironmentString(kRecycleBatchCommitPathEnvVar).empty() ||
               ! GetSelfTestEnvironmentString(kRecycleShellStopPathEnvVar).empty() || ! GetSelfTestEnvironmentString(kRecycleShellCommitStopPathEnvVar).empty();
    };
    if (context.parallel == nullptr || ! anyArmed())
    {
        return;
    }
    if (MatchRecycleSeamPathForSelfTest(kRecycleSiblingHoldPathEnvVar, path, true))
    {
        AwaitParallelStopForSelfTest(context.parallel);
        return;
    }
    if (MatchRecycleSeamPathForSelfTest(kRecycleShellStopPathEnvVar, path, false) ||
        MatchRecycleSeamPathForSelfTest(kRecycleShellCommitStopPathEnvVar, path, false))
    {
        return;
    }
    const ULONGLONG deadline = ::GetTickCount64() + 10'000ull;
    while (anyArmed() && ::GetTickCount64() < deadline)
    {
        ::Sleep(1);
    }
}

// Stands in for a Shell that deletes a regular file permanently while reporting success without a
// Recycle Bin item: the file is really removed and the Recycle route reports the destroyed backstop.
[[nodiscard]] bool TryInjectRecycleDestroyForSelfTest(const PathInfo& path, HRESULT& status, FileSystemItemMutationResult* mutationResult) noexcept
{
    const std::wstring destroyPath = GetSelfTestEnvironmentString(kRecycleDestroyPathEnvVar);
    if (destroyPath.empty())
    {
        return false;
    }

    const PathInfo destroy = MakePathInfo(destroyPath);
    if (! OrdinalString::EqualsNoCase(path.extended, destroy.extended))
    {
        return false;
    }

    static_cast<void>(::SetEnvironmentVariableW(kRecycleDestroyPathEnvVar.data(), nullptr));
    if (::DeleteFileW(path.extended.c_str()) == FALSE)
    {
        status = HRESULT_FROM_WIN32(::GetLastError());
        if (mutationResult != nullptr)
        {
            *mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
        }
        return true;
    }

    const RecycleItemVerdict verdict = ReduceRecyclePostDelete(false, S_OK, false);
    status                           = verdict.status;
    if (mutationResult != nullptr)
    {
        *mutationResult = verdict.receipt;
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

// Polls the caller's own cancellation (operation control, else the legacy callback), whatever the
// call's internal stop state.
HRESULT PollCallerCancellationLocked(OperationContext& context) noexcept
{
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

    return PollCallerCancellationLocked(context);
}

// For work that must finish although another item's failure stopped the call (a Recycle the Shell
// already holds): CheckCancel answers ERROR_CANCELLED for that stop without asking the caller, so
// this keeps polling the caller, on the same cadence, and a user Cancel issued after the stop is
// still seen and recorded.
HRESULT CheckCallerCancel(OperationContext& context) noexcept
{
    if (! context.parallel)
    {
        return PollCallerCancellationLocked(context);
    }
    if (context.parallel->cancelRequested.load(std::memory_order_acquire))
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
    const HRESULT hr = PollCallerCancellationLocked(context);
    context.parallel->lastCancelCheckTick.store(nowTick, std::memory_order_release);
    if (IsCancellationHr(hr))
    {
        context.parallel->cancelRequested.store(true, std::memory_order_release);
    }
    return hr;
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

// delivered, when given, says whether the callback accepted the report: a failure after delivery
// is the cancellation check that follows it, and the item still has its status.
HRESULT ReportItemCompleted(OperationContext& context,
                            unsigned long itemIndex,
                            HRESULT status,
                            const FileSystemItemMutationResult* mutationResult = nullptr,
                            bool* delivered                                    = nullptr) noexcept
{
    if (delivered != nullptr)
    {
        *delivered = false;
    }
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
        if (delivered != nullptr)
        {
            *delivered = true;
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
    if (delivered != nullptr)
    {
        *delivered = true;
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

// I35 R10.5: times the silent phases of one Local native file copy (open, commit, metadata finalize,
// publish), each emitted as its own metric when perf capture is on. `route` names a literal.
class LocalFilePhaseTimer final
{
public:
    LocalFilePhaseTimer(std::wstring_view route, uint64_t fileBytes) noexcept
        : _enabled(Debug::Perf::IsCaptureEnabled()),
          _route(route),
          _fileBytes(fileBytes),
          _start(std::chrono::steady_clock::now())
    {
    }

    LocalFilePhaseTimer(const LocalFilePhaseTimer&)            = delete;
    LocalFilePhaseTimer(LocalFilePhaseTimer&&)                 = delete;
    LocalFilePhaseTimer& operator=(const LocalFilePhaseTimer&) = delete;
    LocalFilePhaseTimer& operator=(LocalFilePhaseTimer&&)      = delete;

    void Restart() noexcept
    {
        _start = std::chrono::steady_clock::now();
    }

    // Emits the phase that ends now and starts timing the next one.
    void Lap(std::wstring_view metric, HRESULT hr) noexcept
    {
        const auto now = std::chrono::steady_clock::now();
        if (_enabled)
        {
            const uint64_t elapsedUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(now - _start).count());
            Debug::Perf::Emit(metric, _route, elapsedUs, _fileBytes, 0u, hr);
        }
        _start = now;
    }

private:
    bool _enabled = false;
    std::wstring_view _route;
    uint64_t _fileBytes = 0;
    std::chrono::steady_clock::time_point _start;
};

#ifdef ENABLE_TESTS
// I35 S6 self-test knobs of the phase reports, read from the environment at most every 100 ms so a
// stream of small files pays no environment lookup per file. A test changes them between operations
// and waits past the refresh before it starts the next one.
// REDSALAMANDER_FILEOPS_SLOW_COMMIT_MS: every Local writer's commit lasts at least this long.
// REDSALAMANDER_FILEOPS_SUPPRESS_PHASE_BOUNDARIES=1: no phase boundary report (a negative control).
struct PhaseReportSelfTestKnobs final
{
    DWORD slowCommitMs      = 0u;
    bool suppressBoundaries = false;
};

[[nodiscard]] PhaseReportSelfTestKnobs ReadPhaseReportSelfTestKnobs() noexcept
{
    static std::atomic<ULONGLONG> refreshedTick{0u};
    static std::atomic<DWORD> slowCommitMs{0u};
    static std::atomic<bool> suppressBoundaries{false};
    const ULONGLONG nowTick = GetTickCount64();
    const ULONGLONG last    = refreshedTick.load(std::memory_order_acquire);
    if (last == 0u || nowTick < last || nowTick - last >= 100u)
    {
        wchar_t text[32]{};
        const DWORD length = GetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_SLOW_COMMIT_MS", text, static_cast<DWORD>(std::size(text)));
        slowCommitMs.store(length > 0u && length < std::size(text) ? static_cast<DWORD>(std::min(std::wcstoul(text, nullptr, 10), 60'000ul)) : 0u,
                           std::memory_order_release);
        const DWORD suppressLength = GetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_SUPPRESS_PHASE_BOUNDARIES", text, static_cast<DWORD>(std::size(text)));
        suppressBoundaries.store(suppressLength == 1u && text[0] == L'1', std::memory_order_release);
        refreshedTick.store(std::max<ULONGLONG>(nowTick, 1u), std::memory_order_release);
    }
    return {slowCommitMs.load(std::memory_order_acquire), suppressBoundaries.load(std::memory_order_acquire)};
}

// The commit of a self-test slowed by REDSALAMANDER_FILEOPS_SLOW_COMMIT_MS (after the real one).
void MaybeSlowCommitForSelfTest(OperationContext& context, uint64_t fileBytes) noexcept
{
    if (const DWORD slowMs = ReadPhaseReportSelfTestKnobs().slowCommitMs; slowMs != 0u)
        Sleep(slowMs);
    // The live-progress capture owns an optional bounded observer lease. Empty files must not
    // block a serial copy before the sampled nonempty file can start. Primary cancel stays active.
    if (fileBytes == 0u)
        return;
    const ULONGLONG deadline = GetTickCount64() + 30'000u;
    for (;;)
    {
        wchar_t held[2]{};
        if (GetEnvironmentVariableW(L"REDSALAMANDER_FILEOPS_HOLD_NONEMPTY_COMMIT", held, static_cast<DWORD>(std::size(held))) != 1u || held[0] != L'1' ||
            GetTickCount64() >= deadline || FAILED(FileSystemCheckOperationControl(context.options)))
            return;
        Sleep(10u);
    }
}
#endif

// I35 R10.1: a report at a phase boundary inside one file (every byte written, before the commit
// flush, the metadata and the publication). It moves no bytes; the host reads the phase from the item
// bytes. An observation point only: a cancellation is honored by the item's next check, so a boundary
// never changes what a cancel leaves behind. A boundary right after another report adds nothing and is
// skipped, which keeps a stream of small files at the ordinary report cadence.
void ReportFilePhaseBoundary(OperationContext& context, uint64_t itemBytes) noexcept
{
#ifdef ENABLE_TESTS
    if (ReadPhaseReportSelfTestKnobs().suppressBoundaries)
        return;
#endif
    constexpr ULONGLONG kMinimumPhaseBoundaryGapMs = 50ull;
    const ULONGLONG nowTick                        = GetTickCount64();
    if (context.lastProgressReportTick != 0u && nowTick >= context.lastProgressReportTick &&
        nowTick - context.lastProgressReportTick < kMinimumPhaseBoundaryGapMs)
    {
        return;
    }
    static_cast<void>(ReportProgressForced(context, itemBytes, itemBytes));
}

// I35 R10.3 candidate (b): reserves the destination's allocation for a file larger than one copy
// buffer when the host asks for it and the destination is not sparse. Best effort: a refusal or a
// failure leaves the copy exactly as it was without it (a cancellation is met by the pump's next check).
void MaybeReserveDestinationAllocation(const OperationContext& context,
                                       IFileWriter& writer,
                                       uint64_t sizeBytes,
                                       uint64_t minimumBytes,
                                       const FileSystemMetadataTransferResult& prepareMetadata) noexcept
{
    if (context.options == nullptr || (context.options->transferTuning & FILESYSTEM_TRANSFER_TUNING_PREALLOCATE) == 0u || sizeBytes <= minimumBytes ||
        (prepareMetadata.attemptedFeatures & FILESYSTEM_METADATA_SPARSE) != 0u)
    {
        return;
    }
    wil::com_ptr<ILocalFileWriterAllocation> allocation;
    if (FAILED(writer.QueryInterface(IID_PPV_ARGS(allocation.addressof()))) || ! allocation)
    {
        return;
    }
    const auto started = std::chrono::steady_clock::now();
    const HRESULT hr   = allocation->ReserveAllocation(sizeBytes);
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Local.FilePreallocateUs", L"", Debug::Perf::ElapsedUs(started), sizeBytes, 0u, hr);
    }
}

// The Local pumps' copy buffer: 1 MiB, or 8 MiB toward a network share whose writer is written through
// (candidate w, FILESYSTEM_TRANSFER_TUNING_WRITE_THROUGH_REMOTE). Each such write waits for the server,
// so it writes the provider's own transfer size for a high-latency destination, as the host bridge does
// there: in 1 MiB writes several concurrent files cost the registered NAS about 13 % of its rate.
[[nodiscard]] unsigned long LocalPumpBufferBytes(const OperationContext& context, std::wstring_view destinationPath) noexcept
{
    constexpr unsigned long kLocalBufferBytes        = 1024u * 1024u;
    constexpr unsigned long kWriteThroughBufferBytes = 8u * 1024u * 1024u;
    const bool writeThrough = context.options != nullptr && (context.options->transferTuning & FILESYSTEM_TRANSFER_TUNING_WRITE_THROUGH_REMOTE) != 0u &&
                              FileSystemInternal::IsRemoteStoragePath(destinationPath);
    return writeThrough ? kWriteThroughBufferBytes : kLocalBufferBytes;
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

    // I35 R10.5: the silent per-file phases around the content pump, timed one by one.
    LocalFilePhaseTimer phaseTimer(L"route=staged-overwrite", expectedFileBytes);
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

    // I35 S1b: the replacement carries the source's metadata exactly like the direct-final route
    // (sparse and EFS before content; streams, MOTW, EA, ordinary metadata and the security
    // comparison after commit, on the exact stage before publication). Without it an Overwrite
    // silently dropped named streams, MOTW and EA even on NTFS.
    wil::com_ptr<IFileSystemBoundMetadata> sourceMetadata;
    static_cast<void>(sourceAuthority->QueryInterface(IID_PPV_ARGS(sourceMetadata.addressof())));
    FileSystemMetadataTransferResult prepareMetadata{};
    prepareMetadata.sizeBytes    = sizeof(prepareMetadata);
    prepareMetadata.firstFailure = S_OK;
    if (sourceMetadata)
    {
        hr = sourceMetadata->TransferMetadataTo(ownedStage.get(), FILESYSTEM_METADATA_TRANSFER_PREPARE_CONTENT, context.options, &prepareMetadata);
        if (FAILED(hr) || (prepareMetadata.lostFeatures & FILESYSTEM_METADATA_EFS) != 0u)
        {
            // Never publish plaintext for an encrypted source.
            return finishWithOwnedStageCleanup(FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_ENCRYPTION_FAILED));
        }
    }
    phaseTimer.Lap(L"FileOps.Local.FileOpenUs", hr);

    const unsigned long bufferBytes = LocalPumpBufferBytes(context, destination.display);
    MaybeReserveDestinationAllocation(context, *writer, readerSize, bufferBytes, prepareMetadata);
    auto buffer = std::unique_ptr<std::byte[]>(new (std::nothrow) std::byte[bufferBytes]);
    if (! buffer)
    {
        return finishWithOwnedStageCleanup(E_OUTOFMEMORY);
    }

    ReportNativeFileWork(context, FileSystemFileWorkPhase::Transferring);
    uint64_t copiedBytes = 0u;
    while (copiedBytes < readerSize)
    {
        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return finishWithOwnedStageCleanup(hr);
        }

        const unsigned long toRead = static_cast<unsigned long>((std::min)(static_cast<uint64_t>(bufferBytes), readerSize - copiedBytes));
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
#if defined(ENABLE_TESTS)
            // Never inject before Write or once the full file has been copied.
            // The normal failure return owns AbortOwnedObject and result axes.
            if (copiedBytes == 0u && consumed == bytesWritten && consumed < readerSize)
            {
                const HRESULT faultHr =
                    MaybeFailStagedCopyAfterWriteForSelfTest(source.display, stagePath, ownedStage.get(), context.options, copiedBytes + consumed);
                if (FAILED(faultHr))
                {
                    return finishWithOwnedStageCleanup(faultHr);
                }
            }
#endif
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

    ReportNativeFileWork(context, FileSystemFileWorkPhase::Finalizing);
    phaseTimer.Restart();
    hr = writer->Commit();
#ifdef ENABLE_TESTS
    if (SUCCEEDED(hr))
        MaybeSlowCommitForSelfTest(context, readerSize);
#endif
    phaseTimer.Lap(L"FileOps.Local.FileCommitUs", hr);
    if (FAILED(hr))
    {
        return finishWithOwnedStageCleanup(hr);
    }
    writer.reset();
    ReportFilePhaseBoundary(context, readerSize);
    phaseTimer.Restart();

    FileSystemMetadataTransferResult finalMetadata{};
    finalMetadata.sizeBytes    = sizeof(finalMetadata);
    finalMetadata.firstFailure = S_OK;
    if (sourceMetadata)
    {
        hr = sourceMetadata->TransferMetadataTo(ownedStage.get(), FILESYSTEM_METADATA_TRANSFER_FINALIZE, context.options, &finalMetadata);
        phaseTimer.Lap(L"FileOps.Local.FileFinalizeMetadataUs", hr);
        if (FAILED(hr))
        {
            return finishWithOwnedStageCleanup(hr);
        }
    }
    else
    {
        FileSystemBasicInformation sourceInfo{};
        sourceInfo.sizeBytes = sizeof(sourceInfo);
        hr                   = sourceAuthority->GetBasicInformation(&sourceInfo);
        if (SUCCEEDED(hr))
        {
            hr = ownedStage->SetBasicInformation(&sourceInfo);
        }
        phaseTimer.Lap(L"FileOps.Local.FileFinalizeMetadataUs", hr);
        if (FAILED(hr) && hr != E_NOTIMPL && hr != HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        {
            return finishWithOwnedStageCleanup(hr);
        }
    }
    ReportFilePhaseBoundary(context, readerSize);
    phaseTimer.Restart();

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
    ReportNativeFileWork(context, FileSystemFileWorkPhase::Publishing);
    hr = ownedStage->PublishAs(
        destination.display.c_str(), expectedDestination.get(), publicationFlags, context.options, &mutationResult, publishedAuthority.put());
    phaseTimer.Lap(L"FileOps.Local.FilePublishUs", hr);
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
    if (sourceMetadata)
    {
        RecordNativeMetadataOutcome(context, CombineMetadataTransferPhases(prepareMetadata, finalMetadata));
    }
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

HRESULT OpenPathForDeleteNoFollow(const std::wstring& pathExtended,
                                  DeleteHandleSnapshot& snapshot,
                                  DWORD shareMode = FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE) noexcept
{
    snapshot = {};

    wil::unique_handle handle(::CreateFileW(pathExtended.c_str(),
                                            DELETE | FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES,
                                            shareMode,
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

#if defined(ENABLE_TESTS)
// This DLL-local seam observes an actual successful child removal. The armed parent is
// compared exactly, not as a prefix; no handle/path lookup or test mutation is performed.
// The host owns the named events until the admitted operation reaches its quiet point.
HRESULT PauseAfterPermanentDeleteChildForSelfTest(std::wstring_view actualPath) noexcept
{
    constexpr std::wstring_view parentVariable  = L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_PARENT";
    constexpr std::wstring_view enteredVariable = L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_ENTERED_EVENT";
    constexpr std::wstring_view releaseVariable = L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_RELEASE_EVENT";
    constexpr std::wstring_view removedVariable = L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_REMOVED_PATH";
    const std::wstring parent                   = GetSelfTestEnvironmentString(parentVariable);
    const size_t separator                      = actualPath.find_last_of(L'\\');
    if (parent.empty() || separator == std::wstring_view::npos || actualPath.substr(0u, separator) != parent)
    {
        return S_OK;
    }

    const std::wstring enteredName = GetSelfTestEnvironmentString(enteredVariable);
    const std::wstring releaseName = GetSelfTestEnvironmentString(releaseVariable);
    wil::unique_handle entered(OpenEventW(EVENT_MODIFY_STATE, FALSE, enteredName.c_str()));
    if (! entered)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    wil::unique_handle release(OpenEventW(SYNCHRONIZE, FALSE, releaseName.c_str()));
    if (! release)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const std::wstring removedPath(actualPath);
    if (! SetEnvironmentVariableW(parentVariable.data(), nullptr) || ! SetEnvironmentVariableW(removedVariable.data(), removedPath.c_str()) ||
        ! SetEvent(entered.get()))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    // No test-only cancellation check: release resumes the unchanged production walk.
    // A lost host/test process cannot leave the provider worker parked indefinitely.
    const DWORD wait = WaitForSingleObject(release.get(), 30'000u);
    return wait == WAIT_OBJECT_0 ? S_OK : HRESULT_FROM_WIN32(wait == WAIT_TIMEOUT ? ERROR_TIMEOUT : GetLastError());
}
#endif

HRESULT DeleteOpenedPathNoFollow(OperationContext& context,
                                 DeleteHandleSnapshot& snapshot,
                                 bool countCompletion,
                                 [[maybe_unused]] std::wstring_view actualPath) noexcept
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
            if (context.options != nullptr && context.options->fileWorkCallback != nullptr)
            {
                FileSystemFileWorkProgress blocked{};
                blocked.sizeBytes   = sizeof(blocked);
                blocked.phase       = FileSystemFileWorkPhase::ReadOnlyDeleteBlocked;
                blocked.kind        = IsDirectory(snapshot.attributes) ? FileSystemFileWorkKind::Directory : FileSystemFileWorkKind::RegularFile;
                blocked.sourcePath  = actualPath.data();
                blocked.status      = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
                blocked.disposition = FileSystemFileWorkDisposition::Failed;
                const auto emit     = [&]() noexcept
                { static_cast<void>(context.options->fileWorkCallback->FileSystemReportFileWork(&blocked, context.options->fileWorkCookie)); };
                if (context.parallel != nullptr)
                {
                    std::scoped_lock lock(context.parallel->callbackMutex);
                    emit();
                }
                else
                    emit();
            }
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
        ReportCompletedObject(context, actualPath.data(), snapshot.attributes, snapshot.fileBytes);
    }

#if defined(ENABLE_TESTS)
    if (countCompletion && context.type == FILESYSTEM_DELETE && ! context.useRecycleBin && ! IsDirectory(snapshot.attributes) &&
        ! IsReparsePoint(snapshot.attributes))
    {
        return PauseAfterPermanentDeleteChildForSelfTest(actualPath);
    }
#endif
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

        size_t offset = 0u;
        for (;;)
        {
            Common::DirectoryRecords::RecordView<FILE_FULL_DIR_INFO> record{};
            const HRESULT parseHr = Common::DirectoryRecords::Read<FILE_FULL_DIR_INFO>(buffer, offset, record);
            if (FAILED(parseHr))
                return parseHr;
            if (record.name != L"." && record.name != L"..")
            {
                const HRESULT hr = onEntry(record.name, record.information->FileAttributes, static_cast<uint64_t>(record.information->EndOfFile.QuadPart));
                if (hr != S_OK)
                    return hr;
            }
            if (record.nextOffset == 0u)
                break;
            offset = record.nextOffset;
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
        return DeleteOpenedPathNoFollow(context, swappedSnapshot, false, directoryExtended);
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
    return DeleteOpenedPathNoFollow(context, snapshot, false, directoryExtended);
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
            return DeleteOpenedPathNoFollow(context, snapshot, false, pathExtended);
        }

        if (snapshot.handle)
        {
            snapshot.handle.reset();
        }
        return RemoveDirectoryRecursiveNoFollow(context, pathExtended);
    }

    return DeleteOpenedPathNoFollow(context, snapshot, false, pathExtended);
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
    if (opContext.fileWorkLeaf != nullptr)
    {
        // Keep boundary observations truthful without adding another callback for each chunk.
        opContext.fileWorkLeaf->sizeKnown      = true;
        opContext.fileWorkLeaf->totalBytes     = itemTotal;
        opContext.fileWorkLeaf->completedBytes = itemCompleted;
    }

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
    const unsigned long bufferBytes = LocalPumpBufferBytes(context, destination.display);
    const auto metricStart          = std::chrono::steady_clock::now();
    HRESULT metricStatus            = E_UNEXPECTED;
    const auto emitMetric           = wil::scope_exit([&]() noexcept
    {
        const uint64_t elapsedUs =
            static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - metricStart).count());
        Debug::Perf::Emit(L"FileOps.Local.DirectFinalCopyUs", L"shape=new-name-local-regular-file", elapsedUs, expectedFileBytes, bufferBytes, metricStatus);
    });
    const auto finish               = [&](HRESULT hr) noexcept -> HRESULT
    {
        metricStatus = hr;
        return hr;
    };

    if (! context.objectBinding || ! context.options)
    {
        return finish(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    }

    // I35 R10.5: the silent per-file phases around the content pump, timed one by one.
    LocalFilePhaseTimer phaseTimer(L"route=direct-final", expectedFileBytes);
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
    phaseTimer.Lap(L"FileOps.Local.FileOpenUs", hr);

    MaybeReserveDestinationAllocation(context, *writer, readerSize, bufferBytes, prepareMetadata);
    auto buffer = std::unique_ptr<std::byte[]>(new (std::nothrow) std::byte[bufferBytes]);
    if (! buffer)
    {
        return finish(finishWithExactAbort(E_OUTOFMEMORY));
    }

    ReportNativeFileWork(context, FileSystemFileWorkPhase::Transferring);
    uint64_t copiedBytes = 0u;
    while (copiedBytes < readerSize)
    {
        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            return finish(finishWithExactAbort(hr));
        }

        const unsigned long toRead = static_cast<unsigned long>((std::min)(static_cast<uint64_t>(bufferBytes), readerSize - copiedBytes));
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

    ReportNativeFileWork(context, FileSystemFileWorkPhase::Finalizing);
    phaseTimer.Restart();
    hr = writer->Commit();
#ifdef ENABLE_TESTS
    if (SUCCEEDED(hr))
        MaybeSlowCommitForSelfTest(context, readerSize);
#endif
    phaseTimer.Lap(L"FileOps.Local.FileCommitUs", hr);
    if (FAILED(hr))
    {
        return finish(finishWithExactAbort(hr));
    }
    writer.reset();
    // The metadata that follows ends at the item's final report (CopyFileInternal), so the commit is
    // the one boundary this route adds.
    ReportFilePhaseBoundary(context, readerSize);
    phaseTimer.Restart();

    FileSystemMetadataTransferResult finalMetadata{};
    finalMetadata.sizeBytes    = sizeof(finalMetadata);
    finalMetadata.firstFailure = S_OK;
    hr                         = sourceMetadata->TransferMetadataTo(ownedFinal.get(), FILESYSTEM_METADATA_TRANSFER_FINALIZE, context.options, &finalMetadata);
    phaseTimer.Lap(L"FileOps.Local.FileFinalizeMetadataUs", hr);
    if (FAILED(hr))
    {
        return finish(finishWithExactAbort(hr));
    }
    if (finalMetadata.lostFeatures != 0u)
    {
        // Named streams, MOTW, extended attributes, or ordinary attributes the destination cannot
        // hold (for example FAT/exFAT) are dropped exactly as CopyFileExW drops them. The complete
        // content is published; the loss is never fatal, and it reaches the host with the item's
        // receipt (I35 S1b) so the card reports it exactly as the bridge does.
        Debug::Perf::EmitValue(L"FileOps.Local.DirectFinalMetadataLost", finalMetadata.lostFeatures, finalMetadata.firstFailure);
    }

    ReportNativeFileWork(context, FileSystemFileWorkPhase::Publishing);
    finalCommitted = true;
    RecordOwnedStagePublication(context, destination, TrackedPublicationTruth::Published, FileSystemOwnedStageDisposition::Published);
    RecordNativeMetadataOutcome(context, CombineMetadataTransferPhases(prepareMetadata, finalMetadata));
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

    BeginNativeFileWork(context, FileSystemFileWorkKind::RegularFile, false, 0u);

    hr = CheckCancel(context);
    if (FAILED(hr))
    {
        return hr;
    }
#ifdef ENABLE_TESTS
    hr = MaybeFailCopyBeforeMutationForSelfTest(source.display);
    if (FAILED(hr))
    {
        return hr;
    }
#endif

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
    if (context.fileWorkLeaf != nullptr)
    {
        context.fileWorkLeaf->sizeKnown  = true;
        context.fileWorkLeaf->totalBytes = fileBytes;
    }
    if (fileBytes > 0u)
    {
        // I35 R10.1: the file is announced at zero of its bytes before the transfer slot, the open,
        // the create and the metadata preparation, so that time belongs to this file's Opening phase
        // rather than to the end of the previous file. An empty file has no transfer to announce.
        hr = ReportProgressForced(context, fileBytes, 0u);
        if (FAILED(hr))
        {
            return hr;
        }
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
    ReportNativeFileWork(context, FileSystemFileWorkPhase::Publishing);
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

    BeginNativeFileWork(context, IsDirectory(sourceAttributes) ? FileSystemFileWorkKind::DirectoryLink : FileSystemFileWorkKind::FileLink, true, 0u);

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

[[nodiscard]] uint64_t AddSaturatingDiscoveryTotal(uint64_t left, uint64_t right) noexcept
{
    return left > std::numeric_limits<uint64_t>::max() - right ? std::numeric_limits<uint64_t>::max() : left + right;
}

HRESULT ReportRawDiscoveryProgress(OperationContext& context,
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

// Every traversal reports through here. A single-root call passes its own cumulative totals straight
// through and owns its closure. A root inside a bulk call is translated into the call's aggregate
// instead, and its closure is withheld: only the bulk entry point knows when the last root is done.
HRESULT ReportDiscoveryProgress(OperationContext& context,
                                uint64_t discoveredBytes,
                                uint64_t discoveredFiles,
                                uint64_t discoveredDirectories,
                                uint32_t queuedItems,
                                bool traversalClosed) noexcept
{
    RootDiscoveryTranslation* const translation = context.discoveryTranslation;
    if (translation == nullptr || translation->aggregator == nullptr)
    {
        return ReportRawDiscoveryProgress(context, discoveredBytes, discoveredFiles, discoveredDirectories, queuedItems, traversalClosed);
    }

    CallDiscoveryAggregator& aggregate = *translation->aggregator;
    // The emission happens under the same lock that advances the aggregate. Roots run concurrently,
    // so computing a total under the lock and emitting it afterwards lets two workers deliver their
    // snapshots out of order and hand the caller a cumulative stream that goes backwards. The lock
    // is provider-local and is never taken while holding a host lock, so serializing here cannot
    // invert against the host's own discovery lock.
    std::scoped_lock lock(aggregate.mutex);
    const uint64_t bytesDelta       = discoveredBytes > translation->lastBytes ? discoveredBytes - translation->lastBytes : 0u;
    const uint64_t filesDelta       = discoveredFiles > translation->lastFiles ? discoveredFiles - translation->lastFiles : 0u;
    const uint64_t directoriesDelta = discoveredDirectories > translation->lastDirectories ? discoveredDirectories - translation->lastDirectories : 0u;
    translation->lastBytes          = (std::max)(translation->lastBytes, discoveredBytes);
    translation->lastFiles          = (std::max)(translation->lastFiles, discoveredFiles);
    translation->lastDirectories    = (std::max)(translation->lastDirectories, discoveredDirectories);

    aggregate.discoveredBytes       = AddSaturatingDiscoveryTotal(aggregate.discoveredBytes, bytesDelta);
    aggregate.discoveredFiles       = AddSaturatingDiscoveryTotal(aggregate.discoveredFiles, filesDelta);
    aggregate.discoveredDirectories = AddSaturatingDiscoveryTotal(aggregate.discoveredDirectories, directoriesDelta);
    aggregate.queuedItems           = (std::max)(aggregate.queuedItems, queuedItems);

    return ReportRawDiscoveryProgress(
        context, aggregate.discoveredBytes, aggregate.discoveredFiles, aggregate.discoveredDirectories, aggregate.queuedItems, false);
}

// The one closure a bulk call emits, after its last root has resolved. It runs on every exit,
// including error, partial and cancellation, because a scope the caller cannot close is worse than
// a scope that closes on totals the call actually reached.
void ReportCallDiscoveryClosed(OperationContext& context, CallDiscoveryAggregator& aggregate) noexcept
{
    std::scoped_lock lock(aggregate.mutex);
    static_cast<void>(ReportRawDiscoveryProgress(
        context, aggregate.discoveredBytes, aggregate.discoveredFiles, aggregate.discoveredDirectories, aggregate.queuedItems, true));
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
        // Nothing was mutated: a Move whose source is gone or unreadable here is a known failure.
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const bool directory          = IsDirectory(data.dwFileAttributes);
    const bool recursiveDirectory = recursiveDirectoryOwnsDiscovery && directory && ! IsReparsePoint(data.dwFileAttributes);
    if (recursiveDirectory)
    {
        // The recursive copy walker owns the root and descendant cumulative totals.
        return S_OK;
    }

    bool literalCopyLink = false;
    if (context.type == FILESYSTEM_COPY && IsReparsePoint(data.dwFileAttributes))
    {
        LocalCopyPathKind kind = LocalCopyPathKind::RegularFile;
        const HRESULT kindHr   = ClassifyLocalCopyPathKind(sourcePath, data.dwFileAttributes, kind);
        if (FAILED(kindHr))
        {
            return kindHr;
        }
        literalCopyLink = kind == LocalCopyPathKind::SemanticLink;
    }
    const uint64_t fileBytes = directory || literalCopyLink ? 0u : (static_cast<uint64_t>(data.nFileSizeHigh) << 32u) | data.nFileSizeLow;
    return ReportDiscoveryProgress(context, fileBytes, directory ? 0u : 1u, directory ? 1u : 0u, 1u, true);
}

// Delete already accumulates into one call-level state, so only its closure was missing. The
// governed call owns that scope and closes it exactly once on every exit, including error, partial
// and cancellation: a scope the caller can never close is worse than one that closes on the totals
// the call actually reached. The emission stays under the lock, matching every other delete report,
// so the closure can never overtake a concurrent worker's report.
void ReportDeleteDiscoveryClosed(OperationContext& context, DeleteDiscoveryState& discovery) noexcept
{
    std::scoped_lock lock(discovery.mutex);
    static_cast<void>(ReportRawDiscoveryProgress(context, discovery.discoveredBytes, discovery.discoveredFiles, discovery.discoveredDirectories, 0u, true));
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

// Several discovered objects at once (a listed folder's contents), in one report.
[[nodiscard]] HRESULT ReportDeleteDiscoveryBatch(OperationContext& context, uint64_t files, uint64_t directories, uint64_t bytes, uint32_t queuedItems) noexcept
{
    DeleteDiscoveryState* const discovery = context.deleteDiscovery;
    if (discovery == nullptr)
    {
        return S_OK;
    }
    const auto saturatingAdd = [](uint64_t left, uint64_t right) noexcept
    { return left > std::numeric_limits<uint64_t>::max() - right ? std::numeric_limits<uint64_t>::max() : left + right; };
    std::scoped_lock lock(discovery->mutex);
    discovery->discoveredFiles       = saturatingAdd(discovery->discoveredFiles, files);
    discovery->discoveredDirectories = saturatingAdd(discovery->discoveredDirectories, directories);
    discovery->discoveredBytes       = saturatingAdd(discovery->discoveredBytes, bytes);
    return ReportDiscoveryProgress(context, discovery->discoveredBytes, discovery->discoveredFiles, discovery->discoveredDirectories, queuedItems, false);
}

// The recursive Copy walkers take one stack level per directory level, on whatever thread runs them
// (a host worker, a shared scheduler worker, or a caller's thread). Before a level starts they make sure
// this much of the thread's stack is left, and otherwise stop the walk with ERROR_STACK_OVERFLOW, as at
// the depth ceiling: a deep tree, a deep caller or larger frames (Debug builds) end in a reported failure,
// never in a stack overflow.
constexpr ULONG_PTR kRecursiveCopyMinimumStackBytes = 128u * 1024u;

[[nodiscard]] bool RecursiveCopyStackIsLow() noexcept
{
    ULONG_PTR low  = 0;
    ULONG_PTR high = 0;
    GetCurrentThreadStackLimits(&low, &high);
    const ULONG_PTR here = reinterpret_cast<ULONG_PTR>(&low);
    return here < low || here - low < kRecursiveCopyMinimumStackBytes;
}

// User plus kernel time the process has used, in microseconds (0 when it cannot be read).
[[nodiscard]] uint64_t ProcessCpuTimeUs() noexcept
{
    FILETIME creation{};
    FILETIME exit{};
    FILETIME kernel{};
    FILETIME user{};
    if (GetProcessTimes(GetCurrentProcess(), &creation, &exit, &kernel, &user) == FALSE)
    {
        return 0u;
    }
    const auto toUs = [](const FILETIME& time) noexcept
    { return ((static_cast<uint64_t>(time.dwHighDateTime) << 32u) | static_cast<uint64_t>(time.dwLowDateTime)) / 10u; };
    return toUs(kernel) + toUs(user);
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

// I35 NAS nested Copy: what the Local Copy walker keeps about a folder it created until the folder's children have
// finished: its path and identity, never a handle. Some SMB servers refuse every rename inside a folder another
// handle holds with DELETE access, which is every child's publication, and a handle per folder for the whole walk
// buys nothing: the restore opens the path again and checks the identity first
// (FileSystemInternal::RestoreLocalDirectoryBasicInformation).
struct LocalCreatedDirectory final
{
    std::wstring path;
    LocalDirectoryIdentity identity;
};

[[nodiscard]] HRESULT PublishLocalCopyDirectory(OperationContext& context,
                                                const PathInfo& destination,
                                                IFileSystemBoundObject* expectedDestination,
                                                LocalCreatedDirectory& created) noexcept
{
    created = {};
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
    wil::com_ptr<IFileSystemBoundObject> publishedAuthority;
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
    // The published object is the stage itself: its handle, which has DELETE access, closes when this function
    // returns. The walker keeps the path and the identity recorded while that handle is still open.
    created.path             = destination.display;
    const HRESULT identityHr = CaptureLocalDirectoryIdentity(publishedAuthority.get(), destination.display, created.identity);
    if (FAILED(identityHr))
    {
        created.identity = {};
        Debug::Perf::Emit(L"fileops.local.created_directory.identity_unavailable", L"", 0u, 1u, 0u, identityHr);
    }
    return hr;
}

[[nodiscard]] HRESULT EnsureLocalCopyDestinationDirectory(OperationContext& context,
                                                          const PathInfo& destination,
                                                          bool& createdDestination,
                                                          LocalCreatedDirectory& createdDirectory) noexcept
{
    createdDestination     = false;
    createdDirectory       = {};
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
        const HRESULT publishHr = PublishLocalCopyDirectory(context, destination, expectedDestination.get(), createdDirectory);
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
    const HRESULT publishHr = PublishLocalCopyDirectory(context, destination, nullptr, createdDirectory);
    if (SUCCEEDED(publishHr))
    {
        createdDestination = true;
    }
    return publishHr;
}

// A created folder's restore, once its children have finished: the source's attributes and times go to the folder at
// the recorded path only when the path still names the folder the walker created. A folder whose identity was not
// recorded keeps the times the copy gave it, and the item reports basic metadata the destination did not keep, which
// the host shows as the I35 R5.4 note, as for a destination without object binding. A folder the path no longer names
// (renamed away, replaced, removed), or that cannot be opened and checked again, is not touched either: the item
// reports its times lost with that refusal, which the host warns about (metadata.basic.lost), as the host bridge does.
void RestoreLocalCreatedDirectoryMetadata(OperationContext& context, const PathInfo& source, const LocalCreatedDirectory& created) noexcept
{
    if (! context.objectBinding || created.path.empty())
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
    if (FAILED(sourceAuthority->GetBasicInformation(&info)))
    {
        return;
    }
    sourceAuthority.reset();

    const LocalDirectoryIdentity::Proof proof = created.identity.proof;
    const auto startedAt                      = std::chrono::steady_clock::now();
    const HRESULT restoreHr  = proof == LocalDirectoryIdentity::Proof::None ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)
                                                                            : RestoreLocalDirectoryBasicInformation(created.path, created.identity, info);
    const uint64_t elapsedUs = Debug::Perf::ElapsedUs(startedAt);
    const std::wstring_view proofDetail = proof == LocalDirectoryIdentity::Proof::FileId      ? L"proof=file-id"
                                          : proof == LocalDirectoryIdentity::Proof::FileIndex ? L"proof=file-index"
                                                                                              : L"proof=none";
    if (SUCCEEDED(restoreHr))
    {
        Debug::Perf::Emit(L"fileops.local.created_directory.restored", proofDetail, elapsedUs, 1u, 0u, restoreHr);
        return;
    }
    Debug::Perf::Emit(L"fileops.local.created_directory.restore_refused", proofDetail, elapsedUs, 1u, 0u, restoreHr);
    if (proof != LocalDirectoryIdentity::Proof::None)
    {
        Debug::Warning(L"FileSystem: the created folder '{}' keeps its times: {} (hr={:#x}).",
                       created.path,
                       restoreHr == HRESULT_FROM_WIN32(ERROR_FILE_INVALID) ? L"its path names another object now" : L"it could not be opened and checked",
                       static_cast<uint32_t>(restoreHr));
    }
    FileSystemMetadataTransferResult transfer{};
    transfer.sizeBytes         = sizeof(transfer);
    transfer.attemptedFeatures = FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES;
    transfer.lostFeatures      = FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES;
    // No recorded identity, or a volume that cannot check it: ERROR_NOT_SUPPORTED makes the host count a destination that
    // keeps no timestamps, one Information note per task. Any other refusal is a write the destination could have kept and
    // did not get: the host warns (I35 P1).
    const bool unsupported = restoreHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || restoreHr == E_NOTIMPL;
    transfer.firstFailure  = unsupported ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) : restoreHr;
    RecordNativeMetadataOutcome(context, transfer);
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
    NativeFileWorkLeafState* previousFileWorkLeaf = context.fileWorkLeaf;
    context.fileWorkLeaf                          = nullptr; // directory publication is never a copied-file event
    const auto restoreFileWorkLeaf                = wil::scope_exit([&]() noexcept { context.fileWorkLeaf = previousFileWorkLeaf; });
    const bool ownsBudget                         = sharedBudget == nullptr;
    RecursiveCopySerialBudget& budget             = sharedBudget != nullptr ? *sharedBudget : localBudget;
    const auto emitBudgetOnExit                   = wil::scope_exit([&]() noexcept
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
    if (RecursiveCopyStackIsLow())
    {
        Debug::Warning(L"FileSystem: serial recursive copy stopped at '{}' at depth {}: the thread's stack is nearly exhausted.", source.display, depth);
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
    LocalCreatedDirectory createdDirectory;
    hr = EnsureLocalCopyDestinationDirectory(context, destination, createdDestination, createdDirectory);
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
            RestoreLocalCreatedDirectoryMetadata(context, source, createdDirectory);
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
            ReportCompletedObject(context, source.display.c_str(), FILE_ATTRIBUTE_DIRECTORY, 0u);
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
            const uint64_t fileBytes =
                childKind == LocalCopyPathKind::SemanticLink ? 0u : (static_cast<uint64_t>(data.nFileSizeHigh) << 32u) | data.nFileSizeLow;
            budget.discoveredBytes = AddRecursiveCopyRetainedBytes(budget.discoveredBytes, fileBytes);
        }
        hr = ReportDiscoveryProgress(context, budget.discoveredBytes, budget.discoveredFiles, budget.discoveredDirectories, 0u, false);
        if (FAILED(hr))
        {
            return hr;
        }

        NativeFileWorkScope fileWorkScope(context, childSource, selectedChildDestination, childKind != LocalCopyPathKind::RegularDirectory);
        const auto prepareChildRetry = [&]() noexcept
        {
            if (childKind == LocalCopyPathKind::RegularDirectory && context.fileWorkCall->published.load(std::memory_order_acquire))
            {
                // The subtree is re-enumerated, so old published descendants no longer have bounded leaf receipts.
                WithdrawNativeFileWorkCounts(context);
            }
        };
        for (;;)
        {
            childBytes = 0;
            childHr    = S_OK;

            if (childKind == LocalCopyPathKind::SemanticLink)
            {
                if (context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
                {
                    BeginNativeFileWork(context, childHasDirectoryShape ? FileSystemFileWorkKind::DirectoryLink : FileSystemFileWorkKind::FileLink, true, 0u);
                    fileWorkScope.SetSkipped();
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

            fileWorkScope.SetResult(childHr);
            if (childHr == S_FALSE)
            {
                fileWorkScope.SetSkipped();
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
                fileWorkScope.SetResult(issueHr);
                return issueHr;
            }

            switch (issueAction)
            {
                case FileSystemIssueAction::Overwrite:
                    context.oneShotAllowOverwrite = true;
                    prepareChildRetry();
                    continue;
                case FileSystemIssueAction::ReplaceLink:
                    context.oneShotAllowOverwrite   = true;
                    context.oneShotAllowReplaceLink = true;
                    prepareChildRetry();
                    continue;
                case FileSystemIssueAction::ReplaceReadOnly:
                    context.oneShotAllowOverwrite       = true;
                    context.oneShotAllowReplaceReadonly = true;
                    prepareChildRetry();
                    continue;
                case FileSystemIssueAction::PermanentDelete:
                case FileSystemIssueAction::Cancel: fileWorkScope.SetResult(HRESULT_FROM_WIN32(ERROR_CANCELLED)); return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                case FileSystemIssueAction::Retry: prepareChildRetry(); continue;
                case FileSystemIssueAction::KeepBoth:
                {
                    const HRESULT keepBothHr = SelectKeepBothPath(childDestination, childHasDirectoryShape, selectedChildDestination);
                    if (FAILED(keepBothHr))
                    {
                        fileWorkScope.SetResult(keepBothHr);
                        return keepBothHr;
                    }
                    prepareChildRetry();
                    continue;
                }
                case FileSystemIssueAction::Skip:
                    fileWorkScope.SetSkipped();
                    hadFailure = true;
                    childHr    = S_OK;
                    break;
                case FileSystemIssueAction::None:
                default: fileWorkScope.SetResult(HRESULT_FROM_WIN32(ERROR_CANCELLED)); return HRESULT_FROM_WIN32(ERROR_CANCELLED);
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

    if (! hadIdenticalFileSkip)
        ReportCompletedObject(context, source.display.c_str(), FILE_ATTRIBUTE_DIRECTORY, 0u);
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
            BeginNativeFileWork(context, IsDirectory(attributes) ? FileSystemFileWorkKind::DirectoryLink : FileSystemFileWorkKind::FileLink, true, 0u);
            if (context.fileWorkLeaf != nullptr)
            {
                context.fileWorkLeaf->skipped = true;
            }
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

// I35 S2 (R6.1/R6.2): a directory's completion outlives the producer frame that enumerated it. Every
// queued child item and every child directory holds one count, and the producer holds one until its
// enumeration of the directory ends. Whoever releases the last count (a worker or the producer)
// restores the created directory's metadata, so children always finish first (deepest-first), and
// then releases the parent's count. Enumeration never waits for a directory's copies to drain.
struct RecursiveCopyDirectoryCompletion final
{
    RecursiveCopyDirectoryCompletion() noexcept                                          = default;
    RecursiveCopyDirectoryCompletion(const RecursiveCopyDirectoryCompletion&)            = delete;
    RecursiveCopyDirectoryCompletion& operator=(const RecursiveCopyDirectoryCompletion&) = delete;
    RecursiveCopyDirectoryCompletion(RecursiveCopyDirectoryCompletion&&)                 = delete;
    RecursiveCopyDirectoryCompletion& operator=(RecursiveCopyDirectoryCompletion&&)      = delete;

    std::atomic<uint64_t> outstanding{1u};
    // Written by the producer before the first child count is taken; read by the last releaser.
    PathInfo source;
    bool createdDestination = false;
    LocalCreatedDirectory createdDirectory;
    std::atomic<bool> restored{false};
    std::atomic<bool> complete{true};
    // A directory whose enumeration ended while children are still pending is charged to the queue's
    // retained budget, so parked directories stay bounded by the same back-pressure as queued items.
    std::atomic<bool> parkedCharge{false};
    uint64_t parkedPathBytes = 0u;
    std::shared_ptr<RecursiveCopyDirectoryCompletion> parent;
};

struct RecursiveCopyWorkItem
{
    RecursiveCopyWorkKind kind = RecursiveCopyWorkKind::File;
    PathInfo source;
    PathInfo destination;
    DWORD attributes           = 0;
    uint64_t retainedPathBytes = 0;
    std::shared_ptr<RecursiveCopyDirectoryCompletion> directory; // the parent directory's completion (files and links)
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
    NativeFileWorkLeafState* previousFileWorkLeaf = rootContext.fileWorkLeaf;
    rootContext.fileWorkLeaf                      = nullptr;
    const auto restoreFileWorkLeaf                = wil::scope_exit([&]() noexcept { rootContext.fileWorkLeaf = previousFileWorkLeaf; });

    constexpr unsigned int kMaxWorkers = 16u;
    // I35 S2 (R6.6): at concurrency 1 the single worker copies in listing order while the producer
    // lists ahead (bounded by the queue), so the directory's discovery closes when its listing ends.
    const unsigned int concurrency = std::max(1u, std::min<unsigned int>(requestedConcurrency, kMaxWorkers));
    if (! GetSharedFileOpsJobScheduler().EnsureWorkersAvailable())
    {
        Debug::Perf::EmitCounter(L"FileOps.CopyRecursiveParallel.SerialFallback.SchedulerUnavailable");
        return CopyDirectoryInternal(rootContext, source, destination, bytesCopied);
    }

    FileSystemOptions* sharedOptionsState = rootContext.options;

    ParallelOperationState parallelStorage{};
    ParallelOperationState& parallel          = rootContext.parallel ? *rootContext.parallel : parallelStorage;
    std::mutex* previousFileWorkCallbackMutex = rootContext.fileWorkCallbackMutex;
    rootContext.fileWorkCallbackMutex         = &parallel.callbackMutex;
    const auto restoreFileWorkCallbackMutex   = wil::scope_exit([&]() noexcept { rootContext.fileWorkCallbackMutex = previousFileWorkCallbackMutex; });
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
    // R6.7: this walk's own first failure. A bulk call shares `parallel` across its roots, so a root that a
    // sibling's failure stopped reports a cancellation, never the sibling's failure.
    std::atomic<HRESULT> walkFailure{S_OK};

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
    // Dispatches whose step found nothing to run: at the walk's active limit (while its listing runs ahead)
    // or with an empty queue. The process CPU time the walk took is measured against its wall time.
    std::atomic<uint64_t> idleAtLimitDispatches{0};
    std::atomic<uint64_t> idleEmptyDispatches{0};
    const auto walkStart                 = std::chrono::steady_clock::now();
    const uint64_t walkStartProcessCpuUs = ProcessCpuTimeUs();

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
        bool listingDone              = false; // this walk's discovery closed (R6.5: its reservation ends)
        bool done                     = false;
        size_t maxActiveItems         = 0;
        // Count of items immediately available in `items`, mirrored as an atomic so the
        // scheduler can gate dynamic-job dispatch without taking `mutex` (avoiding a lock-
        // order inversion against the scheduler lock). Kept in sync under `mutex`.
        std::atomic<size_t> ready{0};
    };

    RecursiveCopyQueue queue{};
    // The walk's dynamic job (started below, before the producer runs) and whether the producer runs on a
    // shared scheduler worker, as each root of a bulk CopyItems does. The producer remembers whether it last saw
    // discovery ahead, so a Skip counting that lifts the reservation wakes the walk (producerDiscoveryAhead).
    SharedFileOpsJobScheduler::JobPtr job;
    const bool producerOnSchedulerWorker      = GetSharedFileOpsJobScheduler().IsWorkerThread();
    uint64_t producerInlineSteps              = 0u;
    bool producerSawDiscoveryAhead            = true;
    uint64_t maxTraversalDepth                = 0;
    uint64_t retainedAncestorMetadataBytes    = 0;
    uint64_t maxRetainedAncestorMetadataBytes = 0;

    const auto initializeChildContext = [&](OperationContext& context, uint64_t progressStreamId) noexcept
    {
        InitializeOperationContext(
            context, rootContext.type, flags, sharedOptionsState, rootContext.callback, rootContext.callbackCookie, 1, reparsePointPolicy);
        context.options                    = sharedOptionsState;
        context.parallel                   = &parallel;
        context.fileWorkCall               = rootContext.fileWorkCall;
        context.totalBytes                 = 0; // discovery totals are reported by this same traversal
        context.progressStreamId           = progressStreamId;
        context.reparseRootSourcePath      = rootSource;
        context.reparseRootDestinationPath = rootDestination;
        context.objectBinding              = rootContext.objectBinding;
        context.destinationMutation        = rootContext.destinationMutation;
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
        GetSharedFileOpsJobScheduler().NotifyDynamicWorkAvailable(job);
    };

    // I35 S2 (R6.7): a worker's failure stops the producer with that failure; only a real cancel is a
    // cancellation. With no per-directory drain the producer is usually inside enqueueWork when a
    // worker fails, and a cancellation here would mask the failure as Canceled.
    // R6.7: a recorded worker failure is the walk's result. Once it stops the task the host answers
    // cancel to every other worker, so a cancellation seen after it is that stop, not the user's.
    const auto stoppedResult = [&]() noexcept -> HRESULT
    {
        if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            const HRESULT failure = walkFailure.load(std::memory_order_acquire);
            if (FAILED(failure))
            {
                return failure;
            }
        }
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    };

    // The producer's view of the discovery mode, read as it lists and queues entries. Skip counting lifts the
    // walk's discovery-ahead reservation (DiscoveryWorkerLimit): the first time the producer sees it, it wakes
    // the walk, which may be parked at that limit. Producer thread only.
    const auto producerDiscoveryAhead = [&]() noexcept -> bool
    {
        const bool discoveryAhead = GetDiscoveryMode(rootContext) == FILESYSTEM_DISCOVERY_AHEAD;
        if (producerSawDiscoveryAhead && ! discoveryAhead)
        {
            producerSawDiscoveryAhead = false;
            GetSharedFileOpsJobScheduler().NotifyDynamicWorkAvailable(job);
        }
        return discoveryAhead;
    };

    const auto enqueueWork = [&](RecursiveCopyWorkItem item) noexcept -> HRESULT
    {
        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            return stoppedResult();
        }

        item.retainedPathBytes = RecursiveCopyRetainedPathBytes(item);
        if (item.retainedPathBytes > Common::FileOperations::kTraversalMaxQueuedPathBytes)
        {
            Debug::Warning(L"FileSystem: recursive copy stopped because one queued item requires {} retained path bytes (limit={}).",
                           item.retainedPathBytes,
                           Common::FileOperations::kTraversalMaxQueuedPathBytes);
            return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
        }

        const auto stoppedLocked = [&]() noexcept
        { return queue.done || parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire); };
        for (;;)
        {
            // The target follows the discovery mode (Skip counting shrinks it), asked outside the queue lock.
            const size_t queueTarget = Common::FileOperations::DiscoveryQueueTarget(static_cast<size_t>(concurrency), producerDiscoveryAhead());
            const auto hasRoomLocked = [&]() noexcept
            {
                return queue.retainedEntries < queueTarget && queue.retainedEntries < Common::FileOperations::kTraversalMaxQueuedEntries &&
                       queue.retainedPathBytes <= Common::FileOperations::kTraversalMaxQueuedPathBytes - item.retainedPathBytes;
            };
            {
                std::scoped_lock lock(queue.mutex);
                if (stoppedLocked())
                {
                    return stoppedResult();
                }
                if (hasRoomLocked())
                {
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
                    break;
                }
            }

            // No room. On a scheduler worker the producer runs one of the walk's own items when none of them
            // is in flight: every worker may be such a producer, waiting for steps none is left to run.
            if (producerOnSchedulerWorker && GetSharedFileOpsJobScheduler().TryRunUnservedDynamicStep(job) != DynamicStep::Idle)
            {
                ++producerInlineSteps;
                continue;
            }

            bool roomOrStop = false;
            {
                std::unique_lock lock(queue.mutex);
                roomOrStop = queue.cv.wait_for(lock, std::chrono::milliseconds(50), [&]() noexcept { return stoppedLocked() || hasRoomLocked(); });
            }
            if (! roomOrStop)
            {
                // The caller's cancel must reach a producer that waits: no worker may be running the walk's
                // items to see it.
                const HRESULT cancelHr = CheckCancel(rootContext);
                if (FAILED(cancelHr))
                {
                    return parallel.stopOnErrorRequested.load(std::memory_order_acquire) ? stoppedResult() : cancelHr;
                }
            }
        }

        queue.cv.notify_one();
        GetSharedFileOpsJobScheduler().NotifyDynamicWorkAvailable(job);
        return S_OK;
    };

    const auto completeWithError = [&](HRESULT failure) noexcept
    {
        failure = NormalizeCancellation(failure);
        if (IsCancellationHr(failure))
        {
            // R6.7: a cancellation after a recorded failure is that failure's stop, not the user's cancel.
            if (! parallel.stopOnErrorRequested.load(std::memory_order_acquire))
            {
                parallel.cancelRequested.store(true, std::memory_order_release);
            }
        }
        else
        {
            HRESULT expected = S_OK;
            static_cast<void>(walkFailure.compare_exchange_strong(expected, failure, std::memory_order_acq_rel));
            expected = S_OK;
            static_cast<void>(parallel.firstError.compare_exchange_strong(expected, failure, std::memory_order_acq_rel));
            parallel.stopOnErrorRequested.store(true, std::memory_order_release);
        }
        {
            std::scoped_lock lock(queue.mutex);
            queue.done = true;
        }
        queue.cv.notify_all();
    };

    std::atomic<uint64_t> parkedDirectoriesMax{0u};
    std::atomic<uint64_t> parkedDirectories{0u};

    // I35 S2 (D-6.2): bounded name-only count-ahead. While discovery runs ahead and the queue has no
    // room, the producer keeps listing: an entry it cannot queue yet waits in its directory's stash as
    // a name record (leaf, attributes, kind), charged here against a 16 MiB name budget and never
    // against the queue's retained work entries. Once every directory of the walk is listed the scope
    // closes, and the stashes then drain into the queue in listing order. Beyond the budget the
    // producer waits for the queue (D-6.1). Producer-thread state only: workers never list.
    constexpr uint64_t kCountAheadNameBudgetBytes = 16ull * 1024ull * 1024ull;
    uint64_t countAheadRetainedBytes              = 0u;
    uint64_t countAheadMaxRetainedBytes           = 0u;
    uint64_t countAheadPendingDirectories         = 0u; // held directory entries not listed yet
    uint64_t listingsInProgress                   = 0u; // directories whose FindNextFile loop is running
    bool rootListingDone                          = false;
    bool discoveryClosureReported                 = false;

    // A non-blocking view of enqueueWork's back-pressure: room for one more item of `pathBytes`.
    const auto queueHasRoom = [&](uint64_t pathBytes) noexcept -> bool
    {
        // The mode is asked outside the queue lock (a wake from it must not run under that lock).
        const size_t target = Common::FileOperations::DiscoveryQueueTarget(static_cast<size_t>(concurrency), producerDiscoveryAhead());
        std::scoped_lock lock(queue.mutex);
        return queue.retainedEntries < target && queue.retainedEntries < Common::FileOperations::kTraversalMaxQueuedEntries &&
               pathBytes <= Common::FileOperations::kTraversalMaxQueuedPathBytes &&
               queue.retainedPathBytes <= Common::FileOperations::kTraversalMaxQueuedPathBytes - pathBytes;
    };

    // Entries the queue holds for work, excluding parked directories (reported as queued items).
    const auto queuedWorkEntries = [&]() noexcept -> uint32_t
    {
        std::scoped_lock lock(queue.mutex);
        const uint64_t parked = parkedDirectories.load(std::memory_order_acquire);
        const size_t work     = queue.retainedEntries >= parked ? queue.retainedEntries - static_cast<size_t>(parked) : 0u;
        return static_cast<uint32_t>((std::min)(work, static_cast<size_t>(std::numeric_limits<uint32_t>::max())));
    };

    // R6.1: the scope closes when the walk's last listing ends, before held entries are queued.
    const auto maybeCloseDiscoveryAtListingEnd = [&](OperationContext& context) noexcept -> HRESULT
    {
        if (discoveryClosureReported || ! rootListingDone || countAheadPendingDirectories != 0u || listingsInProgress != 0u)
        {
            return S_OK;
        }
        discoveryClosureReported = true;
        {
            std::scoped_lock lock(queue.mutex);
            queue.listingDone = true;
        }
        // The discovery-ahead reservation ends here: a walk parked at its active limit may run more items.
        GetSharedFileOpsJobScheduler().NotifyDynamicWorkAvailable(job);
        return ReportDiscoveryProgress(context, discoveredBytes, discoveredFiles, discoveredDirectories, queuedWorkEntries(), true);
    };

    const auto uncharge = [&](RecursiveCopyDirectoryCompletion& directory) noexcept
    {
        if (! directory.parkedCharge.exchange(false, std::memory_order_acq_rel))
        {
            return;
        }
        {
            std::scoped_lock lock(queue.mutex);
            parkedDirectories.fetch_sub(1u, std::memory_order_acq_rel);
            if (queue.retainedEntries > 0u)
            {
                --queue.retainedEntries;
            }
            queue.retainedPathBytes = queue.retainedPathBytes >= directory.parkedPathBytes ? queue.retainedPathBytes - directory.parkedPathBytes : 0u;
        }
        queue.cv.notify_all();
    };

    // Releases one count; the last releaser restores the directory's metadata and walks up.
    const auto releaseDirectory = [&](std::shared_ptr<RecursiveCopyDirectoryCompletion> directory, OperationContext& context) noexcept
    {
        while (directory)
        {
            if (directory->outstanding.fetch_sub(1u, std::memory_order_acq_rel) != 1u)
            {
                return;
            }
            if (directory->createdDestination && ! directory->restored.exchange(true, std::memory_order_acq_rel))
            {
                RestoreLocalCreatedDirectoryMetadata(context, directory->source, directory->createdDirectory);
            }
            const bool complete = directory->complete.load(std::memory_order_acquire);
            if (complete)
                ReportCompletedObject(context, directory->source.display.c_str(), FILE_ATTRIBUTE_DIRECTORY, 0u);
            else if (directory->parent)
                directory->parent->complete.store(false, std::memory_order_release);
            directory->createdDirectory = {};
            uncharge(*directory);
            std::shared_ptr<RecursiveCopyDirectoryCompletion> parent = std::move(directory->parent);
            directory                                                = std::move(parent);
        }
    };

    // The producer's enumeration of a directory ended: park it (charged to the queue budget, so parked
    // directories stay under the same back-pressure as queued items) when children are still pending,
    // then release the producer's count.
    const auto finishDirectoryEnumeration = [&](std::shared_ptr<RecursiveCopyDirectoryCompletion> directory, OperationContext& context) noexcept
    {
        if (directory->outstanding.load(std::memory_order_acquire) > 1u)
        {
            directory->parkedPathBytes = RecursiveCopyRetainedPathBytes(directory->source, directory->source);
            directory->parkedCharge.store(true, std::memory_order_release);
            std::scoped_lock lock(queue.mutex);
            const uint64_t parked = parkedDirectories.fetch_add(1u, std::memory_order_acq_rel) + 1u;
            if (parked > parkedDirectoriesMax.load(std::memory_order_acquire))
            {
                parkedDirectoriesMax.store(parked, std::memory_order_release);
            }
            ++queue.retainedEntries;
            queue.retainedPathBytes += directory->parkedPathBytes;
            queue.maxRetainedEntries   = (std::max)(queue.maxRetainedEntries, static_cast<uint64_t>(queue.retainedEntries));
            queue.maxRetainedPathBytes = (std::max)(queue.maxRetainedPathBytes, queue.retainedPathBytes);
        }
        releaseDirectory(std::move(directory), context);
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

        uint64_t ancestorBytes = 0;
        std::wstring searchPattern;
        WIN32_FIND_DATAW findData{};
        wil::unique_hfind findHandle;
        uint64_t cancelCheckCounter = 0;
        RecursiveCopyWorkItem childItem;

        // D-6.2: entries listed ahead of the queue, in listing order.
        struct HeldChild final
        {
            std::wstring leaf;
            DWORD attributes       = 0;
            LocalCopyPathKind kind = LocalCopyPathKind::RegularFile;
            uint64_t retainedBytes = 0;
            bool countSubtree      = true; // false while a retried directory replays counted entries
        };
        std::deque<HeldChild> held;

        // The listed entry examineEntry prepared (in childItem) and what to do with it, and the held entry
        // drainHeld is queuing, so neither occupies the recursive producer's stack.
        LocalCopyPathKind childKind = LocalCopyPathKind::RegularFile;
        bool childCountSubtree      = true;
        HeldChild drainEntry;
        RecursiveCopyWorkItem drainItem;
    };

    using DirectoryCompletionPtr = std::shared_ptr<RecursiveCopyDirectoryCompletion>;
    // The producer recurses once per directory level, on the calling thread or a shared scheduler worker
    // (default 1 MiB stack, Debug builds included), up to kTraversalMaxDepth levels. Only runDirectory and
    // processDirectory are on that recursion; per-entry work (examineEntry), failure decisions
    // (resolveDirectoryFailure) and file queuing (queueChild) run in their own frames, which return before
    // the next level starts, and only runDirectory goes through a std::function.
    std::function<HRESULT(OperationContext&, const RecursiveCopyWorkItem&, uint64_t, const DirectoryCompletionPtr&, bool)> runDirectory;

    enum class ListedEntryStep : uint8_t
    {
        Skip,            // a dot entry or a link the policy skips: nothing to queue, no cancel check
        Held,            // held by name (D-6.2): it goes to the queue when the listing ends
        Handle,          // queue the entry now (or walk it, for a directory)
        DrainThenHandle, // the name budget is spent: queue the held entries first, then this one
    };

    // One listed entry up to the decision what to do with it: its paths and kind (in frame.childItem and
    // frame.childKind), its discovery count (R6.3: a retried directory replays the entries an earlier
    // attempt counted, and the subtrees below them, without counting them again; a subdirectory that
    // attempt held by name but never walked still owes its subtree), and the count-ahead hold.
    const auto examineEntry = [&](OperationContext& context,
                                  const RecursiveCopyWorkItem& item,
                                  RecursiveCopyDirectoryFrame& frame,
                                  uint64_t& enumeratedEntries,
                                  uint64_t& countedEntries,
                                  std::vector<std::wstring>& owedSubtrees,
                                  const std::vector<std::wstring>& incompleteChildren,
                                  bool countDiscovery,
                                  ListedEntryStep& step) noexcept -> HRESULT
    {
        step = ListedEntryStep::Skip;
        if (IsDotOrDotDot(frame.findData.cFileName))
        {
            return S_OK;
        }
        const bool replayedEntry = enumeratedEntries++ < countedEntries;
        const bool countEntry    = countDiscovery && ! replayedEntry;

        frame.childItem                      = RecursiveCopyWorkItem{};
        frame.childItem.source.display       = AppendPath(item.source.display, frame.findData.cFileName);
        frame.childItem.source.extended      = AppendPath(item.source.extended, frame.findData.cFileName);
        frame.childItem.destination.display  = AppendPath(item.destination.display, frame.findData.cFileName);
        frame.childItem.destination.extended = AppendPath(item.destination.extended, frame.findData.cFileName);

        const DWORD childAttributes       = frame.findData.dwFileAttributes;
        const bool childHasDirectoryShape = IsDirectory(childAttributes);
        frame.childKind                   = LocalCopyPathKind::RegularFile;
        HRESULT hr                        = ClassifyLocalCopyPathKind(frame.childItem.source.extended, childAttributes, frame.childKind);
        if (FAILED(hr))
        {
            return hr;
        }
        bool owedSubtree = false;
        if (replayedEntry && frame.childKind == LocalCopyPathKind::RegularDirectory)
        {
            // A subdirectory an earlier attempt walked but left incomplete (skipped, or failed under
            // continue-on-error) keeps that outcome: a retry of its parent does not walk it again.
            if (std::find(incompleteChildren.begin(), incompleteChildren.end(), std::wstring_view(frame.findData.cFileName)) != incompleteChildren.end())
            {
                return S_OK; // step stays Skip
            }
            const auto owed = std::find(owedSubtrees.begin(), owedSubtrees.end(), std::wstring_view(frame.findData.cFileName));
            if (owed != owedSubtrees.end())
            {
                owedSubtrees.erase(owed);
                owedSubtree = true;
            }
        }
        frame.childCountSubtree = countDiscovery && (! replayedEntry || owedSubtree);

        if (countEntry)
        {
            ++countedEntries;
            if (childHasDirectoryShape)
            {
                ++discoveredDirectories;
            }
            else
            {
                ++discoveredFiles;
                const uint64_t fileBytes = frame.childKind == LocalCopyPathKind::SemanticLink
                                               ? 0u
                                               : (static_cast<uint64_t>(frame.findData.nFileSizeHigh) << 32u) | frame.findData.nFileSizeLow;
                discoveredBytes          = AddRecursiveCopyRetainedBytes(discoveredBytes, fileBytes);
            }
            hr = ReportDiscoveryProgress(context, discoveredBytes, discoveredFiles, discoveredDirectories, queuedWorkEntries(), false);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        frame.childItem.attributes = childAttributes;

        if (frame.childKind == LocalCopyPathKind::SemanticLink && context.reparsePointPolicy == FileSystemReparsePointPolicy::Skip)
        {
            NativeFileWorkScope fileWorkScope(context, frame.childItem.source, frame.childItem.destination);
            fileWorkScope.SetProducerStream();
            BeginNativeFileWork(context, childHasDirectoryShape ? FileSystemFileWorkKind::DirectoryLink : FileSystemFileWorkKind::FileLink, true, 0u);
            fileWorkScope.SetResult(S_OK);
            fileWorkScope.SetSkipped();
            hadSkipped.store(true, std::memory_order_release);
            return S_OK;
        }
        if (frame.childKind == LocalCopyPathKind::RegularDirectory && ! context.recursive)
        {
            return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
        }

        // D-6.2: once this directory holds entries, or the queue has no room, keep listing and hold the
        // entry by name while the budget allows; otherwise the held entries go first.
        step = ListedEntryStep::Handle;
        if (! discoveryClosureReported && (! frame.held.empty() || ! queueHasRoom(RecursiveCopyRetainedPathBytes(frame.childItem))) && producerDiscoveryAhead())
        {
            const uint64_t leafBytes     = static_cast<uint64_t>(wcslen(frame.findData.cFileName)) * sizeof(wchar_t);
            const uint64_t retainedBytes = AddRecursiveCopyRetainedBytes(leafBytes, Common::FileOperations::kTraversalRecordOverheadBytes);
            if (retainedBytes <= kCountAheadNameBudgetBytes && countAheadRetainedBytes <= kCountAheadNameBudgetBytes - retainedBytes)
            {
                RecursiveCopyDirectoryFrame::HeldChild& heldEntry = frame.held.emplace_back();
                heldEntry.leaf.assign(frame.findData.cFileName);
                heldEntry.attributes    = childAttributes;
                heldEntry.kind          = frame.childKind;
                heldEntry.retainedBytes = retainedBytes;
                heldEntry.countSubtree  = frame.childCountSubtree;
                countAheadRetainedBytes += retainedBytes;
                countAheadMaxRetainedBytes = (std::max)(countAheadMaxRetainedBytes, countAheadRetainedBytes);
                if (frame.childKind == LocalCopyPathKind::RegularDirectory)
                {
                    ++countAheadPendingDirectories;
                }
                step = ListedEntryStep::Held;
            }
            else
            {
                step = ListedEntryStep::DrainThenHandle;
            }
        }
        return S_OK;
    };

    // Queues one listed file or link of `directory`; directories are walked by the caller.
    const auto queueChild = [&](RecursiveCopyWorkItem& child, LocalCopyPathKind kind, const DirectoryCompletionPtr& directory) noexcept -> HRESULT
    {
        child.kind = kind == LocalCopyPathKind::SemanticLink ? RecursiveCopyWorkKind::ReparsePoint : RecursiveCopyWorkKind::File;
        directory->outstanding.fetch_add(1u, std::memory_order_acq_rel);
        child.directory  = directory;
        const HRESULT hr = enqueueWork(std::move(child));
        if (FAILED(hr))
        {
            // The item was not queued: drop its count (the producer's count keeps this nonzero).
            directory->outstanding.fetch_sub(1u, std::memory_order_acq_rel);
        }
        return hr;
    };

    const auto releaseHeld = [&](const RecursiveCopyDirectoryFrame::HeldChild& entry) noexcept
    {
        countAheadRetainedBytes = countAheadRetainedBytes >= entry.retainedBytes ? countAheadRetainedBytes - entry.retainedBytes : 0u;
        if (entry.kind == LocalCopyPathKind::RegularDirectory && countAheadPendingDirectories > 0u)
        {
            --countAheadPendingDirectories;
        }
    };

    // Takes the oldest held entry of `frame` into frame.drainEntry / frame.drainItem (D-6.2: held entries
    // go to the queue in listing order).
    const auto takeHeld = [&](const RecursiveCopyWorkItem& item, RecursiveCopyDirectoryFrame& frame) noexcept
    {
        frame.drainEntry = std::move(frame.held.front());
        frame.held.pop_front();
        releaseHeld(frame.drainEntry);
        frame.drainItem                      = RecursiveCopyWorkItem{};
        frame.drainItem.source.display       = AppendPath(item.source.display, frame.drainEntry.leaf.c_str());
        frame.drainItem.source.extended      = AppendPath(item.source.extended, frame.drainEntry.leaf.c_str());
        frame.drainItem.destination.display  = AppendPath(item.destination.display, frame.drainEntry.leaf.c_str());
        frame.drainItem.destination.extended = AppendPath(item.destination.extended, frame.drainEntry.leaf.c_str());
        frame.drainItem.attributes           = frame.drainEntry.attributes;
    };

    // countedEntries / owedSubtrees / incompleteChildren: what earlier attempts of this directory (a
    // retried directory issue) already reported to discovery or left incomplete (R6.3).
    const auto processDirectory = [&](OperationContext& context,
                                      const RecursiveCopyWorkItem& item,
                                      uint64_t depth,
                                      const DirectoryCompletionPtr& parentDirectory,
                                      uint64_t& countedEntries,
                                      std::vector<std::wstring>& owedSubtrees,
                                      std::vector<std::wstring>& incompleteChildren,
                                      bool countDiscovery) noexcept -> HRESULT
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
        if (RecursiveCopyStackIsLow())
        {
            Debug::Warning(L"FileSystem: recursive copy stopped at '{}' at depth {}: the thread's stack is nearly exhausted.", item.source.display, depth);
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

        // This attempt's completion: the producer's count is released when the attempt ends, on every
        // path; queued children and child directories hold their own counts.
        DirectoryCompletionPtr directory = std::make_shared<RecursiveCopyDirectoryCompletion>();
        directory->source                = item.source;
        bool enumerationComplete         = false;
        if (parentDirectory)
        {
            parentDirectory->outstanding.fetch_add(1u, std::memory_order_acq_rel);
            directory->parent = parentDirectory;
        }
        const auto endEnumeration = wil::scope_exit([&]() noexcept
        {
            if (! enumerationComplete)
                directory->complete.store(false, std::memory_order_release);
            finishDirectoryEnumeration(std::move(directory), context);
        });

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

        directoryHr = EnsureLocalCopyDestinationDirectory(context, item.destination, directory->createdDestination, directory->createdDirectory);
        if (FAILED(directoryHr))
        {
            return directoryHr;
        }

        // A grant that authorized this directory's own destination conflict must not leak into
        // children enqueued below; they conflict and prompt on their own.
        ClearOneShotGrants(context);

        // Queues the held entries in listing order (blocking on the queue's back-pressure); a held
        // directory is walked when its turn comes.
        const auto drainHeld = [&]() noexcept -> HRESULT
        {
            while (! frame->held.empty())
            {
                takeHeld(item, *frame);
                HRESULT hr = S_OK;
                if (frame->drainEntry.kind == LocalCopyPathKind::RegularDirectory)
                {
                    queuedDirectories.fetch_add(1, std::memory_order_relaxed);
                    hr = runDirectory(context, frame->drainItem, depth + 1u, directory, frame->drainEntry.countSubtree);
                    if (hr == S_FALSE)
                    {
                        incompleteChildren.push_back(frame->drainEntry.leaf);
                    }
                    if (SUCCEEDED(hr))
                    {
                        // R6.1: a held directory that ended without finishing its listing (skipped, or failed
                        // under continue-on-error) may have been the last one pending.
                        hr = maybeCloseDiscoveryAtListingEnd(context);
                    }
                }
                else
                {
                    hr = queueChild(frame->drainItem, frame->drainEntry.kind, directory);
                }
                if (FAILED(hr))
                {
                    return hr;
                }
            }
            return S_OK;
        };
        // Every exit, failure included, returns what this directory still holds to the budget. A child
        // directory this attempt counted but only held by name was never walked: a retry still owes
        // discovery its subtree (R6.3). Only std::bad_alloc can follow, and it is fatal by policy.
        const auto releaseHeldOnExit = wil::scope_exit([&]() noexcept
        {
            for (const RecursiveCopyDirectoryFrame::HeldChild& entry : frame->held)
            {
                releaseHeld(entry);
                if (entry.kind == LocalCopyPathKind::RegularDirectory && entry.countSubtree)
                {
                    owedSubtrees.push_back(entry.leaf);
                }
            }
            frame->held.clear();
        });

        frame->searchPattern = AppendPath(item.source.extended, L"*");
        frame->findHandle.reset(
            FindFirstFileExW(frame->searchPattern.c_str(), FindExInfoBasic, &frame->findData, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
        if (! frame->findHandle)
        {
            const DWORD error = GetLastError();
            if (error == ERROR_FILE_NOT_FOUND)
            {
                // Empty: the producer's count is the last one, so its metadata is restored on return.
                if (depth == 0u)
                {
                    rootListingDone = true;
                }
                directoryHr = maybeCloseDiscoveryAtListingEnd(context);
                if (FAILED(directoryHr))
                {
                    return directoryHr;
                }
                processedDirectories.fetch_add(1, std::memory_order_relaxed);
                enumerationComplete = true;
                return S_OK;
            }
            return HRESULT_FROM_WIN32(error);
        }
        // A listing in progress keeps the scope open: a nested directory that ends its listing first
        // must not close it while an ancestor is still listing.
        ++listingsInProgress;
        bool listingOpen         = true;
        const auto endListingRun = wil::scope_exit([&]() noexcept
        {
            if (listingOpen)
            {
                --listingsInProgress;
            }
        });

        uint64_t enumeratedEntries = 0u;
#ifdef ENABLE_TESTS
        const bool listFailureArmed = IsListingFailureArmedForSelfTest();
#endif
        do
        {
            ListedEntryStep step = ListedEntryStep::Skip;
            directoryHr = examineEntry(context, item, *frame, enumeratedEntries, countedEntries, owedSubtrees, incompleteChildren, countDiscovery, step);
            if (FAILED(directoryHr))
            {
                return directoryHr;
            }
            if (step == ListedEntryStep::Skip)
            {
                continue;
            }
            if (step == ListedEntryStep::DrainThenHandle)
            {
                directoryHr = drainHeld();
                if (FAILED(directoryHr))
                {
                    return directoryHr;
                }
            }
            if (step != ListedEntryStep::Held)
            {
                if (frame->childKind == LocalCopyPathKind::RegularDirectory)
                {
                    queuedDirectories.fetch_add(1, std::memory_order_relaxed);
                    directoryHr = runDirectory(context, frame->childItem, depth + 1u, directory, frame->childCountSubtree);
                    if (directoryHr == S_FALSE)
                    {
                        incompleteChildren.emplace_back(frame->findData.cFileName);
                    }
                }
                else
                {
                    directoryHr = queueChild(frame->childItem, frame->childKind, directory);
                }
                if (FAILED(directoryHr))
                {
                    return directoryHr;
                }
            }

            if ((++frame->cancelCheckCounter % 64u) == 0u)
            {
                directoryHr = CheckCancel(context);
                if (FAILED(directoryHr))
                {
                    return directoryHr;
                }
            }
#ifdef ENABLE_TESTS
            if (listFailureArmed)
            {
                directoryHr = MaybeFailListingForSelfTest(item.source.display, enumeratedEntries);
                if (FAILED(directoryHr))
                {
                    return directoryHr;
                }
            }
#endif
        } while (FindNextFileW(frame->findHandle.get(), &frame->findData));

        const DWORD enumError = GetLastError();
        if (enumError != ERROR_NO_MORE_FILES)
        {
            return HRESULT_FROM_WIN32(enumError);
        }

        frame->findHandle.reset();

        // R6.1: the directory's enumeration is complete; its children finish on the workers and the
        // last of them restores its metadata. The producer moves on without waiting. When this was the
        // walk's last listing the scope closes now, before the entries held by name are queued.
        listingOpen = false;
        --listingsInProgress;
        if (depth == 0u)
        {
            rootListingDone = true;
        }
        directoryHr = maybeCloseDiscoveryAtListingEnd(context);
        if (FAILED(directoryHr))
        {
            return directoryHr;
        }
        directoryHr = drainHeld();
        if (FAILED(directoryHr))
        {
            return directoryHr;
        }
        processedDirectories.fetch_add(1, std::memory_order_relaxed);
        enumerationComplete = true;
        return S_OK;
    };

    // Decides what a failed attempt of one directory does next: true retries it (itemHr is unchanged),
    // false returns itemHr from runDirectory (S_FALSE: the directory was left incomplete, skipped or failed
    // under continue-on-error).
    const auto retryFailedDirectory =
        [&](OperationContext& context, const RecursiveCopyWorkItem& item, RecursiveCopyWorkItem& selectedItem, HRESULT& itemHr) noexcept -> bool
    {
        if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            // R6.7: a worker's failure already stopped the walk; never ask again or turn it into a
            // cancel. The producer may have seen the stop as a cancellation (the stop answers every other
            // cancel check with ERROR_CANCELLED): it reports the worker's failure.
            itemHr = stoppedResult();
            return false;
        }

        itemHr = NormalizeCancellation(itemHr);
        if (IsCancellationHr(itemHr))
        {
            completeWithError(itemHr);
            return false;
        }
        if (itemHr == HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW) || itemHr == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY))
        {
            completeWithError(itemHr);
            return false;
        }
        if (itemHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
        {
            hadSkipped.store(true, std::memory_order_release);
            itemHr = S_FALSE;
            return false;
        }
        if (context.continueOnError)
        {
            hadFailure.store(true, std::memory_order_release);
            itemHr = S_FALSE;
            return false;
        }

        FileSystemIssueAction issueAction = FileSystemIssueAction::Cancel;
        const HRESULT issueHr             = ReportIssue(context, itemHr, &issueAction);
        if (FAILED(issueHr))
        {
            completeWithError(issueHr);
            itemHr = issueHr;
            return false;
        }

        switch (issueAction)
        {
            case FileSystemIssueAction::Overwrite: context.oneShotAllowOverwrite = true; return true;
            case FileSystemIssueAction::ReplaceLink:
                context.oneShotAllowOverwrite   = true;
                context.oneShotAllowReplaceLink = true;
                return true;
            case FileSystemIssueAction::ReplaceReadOnly:
                context.oneShotAllowOverwrite       = true;
                context.oneShotAllowReplaceReadonly = true;
                return true;
            case FileSystemIssueAction::PermanentDelete:
            case FileSystemIssueAction::Cancel: itemHr = HRESULT_FROM_WIN32(ERROR_CANCELLED); return false;
            case FileSystemIssueAction::Retry: return true;
            case FileSystemIssueAction::KeepBoth:
            {
                const HRESULT keepBothHr = SelectKeepBothPath(item.destination, true, selectedItem.destination);
                if (FAILED(keepBothHr))
                {
                    completeWithError(keepBothHr);
                    itemHr = keepBothHr;
                    return false;
                }
                return true;
            }
            case FileSystemIssueAction::Skip:
                hadFailure.store(true, std::memory_order_release);
                itemHr = S_FALSE;
                return false;
            case FileSystemIssueAction::None:
            default:
                completeWithError(HRESULT_FROM_WIN32(ERROR_CANCELLED));
                itemHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                return false;
        }
    };

    runDirectory = [&](OperationContext& context,
                       const RecursiveCopyWorkItem& item,
                       uint64_t depth,
                       const DirectoryCompletionPtr& parentDirectory,
                       bool countDiscovery) noexcept -> HRESULT
    {
        auto selectedItem       = std::make_unique<RecursiveCopyWorkItem>(item);
        uint64_t countedEntries = 0u;
        std::vector<std::wstring> owedSubtrees;
        std::vector<std::wstring> incompleteChildren;
        for (;;)
        {
            HRESULT itemHr = processDirectory(context, *selectedItem, depth, parentDirectory, countedEntries, owedSubtrees, incompleteChildren, countDiscovery);
            if (SUCCEEDED(itemHr) || ! retryFailedDirectory(context, item, *selectedItem, itemHr))
            {
                return itemHr;
            }
            if (countedEntries != 0u)
            {
                // Previously queued leaves may already be published or still in flight when this listing replays.
                WithdrawNativeFileWorkCounts(context);
            }
        }
    };

    const auto processWork = [&](OperationContext& context, const RecursiveCopyWorkItem& item) noexcept
    {
        bool completed             = false;
        const auto settleDirectory = wil::scope_exit([&]() noexcept
        {
            if (! completed && item.directory)
                item.directory->complete.store(false, std::memory_order_release);
        });
        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
        {
            return;
        }

        // Worker contexts are reused across queue items; a per-conflict grant must never
        // survive into an unrelated item.
        auto clearGrantsOnExit = wil::scope_exit([&]() noexcept { ClearOneShotGrants(context); });

        PathInfo selectedDestination = item.destination;
        NativeFileWorkScope fileWorkScope(context, item.source, selectedDestination);
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

            fileWorkScope.SetResult(itemHr);
            if (itemHr == S_FALSE)
            {
                fileWorkScope.SetSkipped();
                hadIdenticalFileSkip.store(true, std::memory_order_release);
            }
            if (SUCCEEDED(itemHr))
            {
                completed = itemHr == S_OK;
                return;
            }

            itemHr = NormalizeCancellation(itemHr);
            if (IsCancellationHr(itemHr))
            {
                if (! parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                {
                    parallel.cancelRequested.store(true, std::memory_order_release); // R6.7: not a recorded stop's echo
                }
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
                fileWorkScope.SetResult(issueHr);
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
                        fileWorkScope.SetResult(keepBothHr);
                        completeWithError(keepBothHr);
                        return;
                    }
                    continue;
                }
                case FileSystemIssueAction::Skip:
                    fileWorkScope.SetSkipped();
                    hadFailure.store(true, std::memory_order_release);
                    return;
                case FileSystemIssueAction::PermanentDelete:
                case FileSystemIssueAction::Cancel:
                case FileSystemIssueAction::None:
                default:
                    fileWorkScope.SetResult(HRESULT_FROM_WIN32(ERROR_CANCELLED));
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
    job = GetSharedFileOpsJobScheduler().StartDynamicJob(rootContext.options,
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
                idleEmptyDispatches.fetch_add(1u, std::memory_order_relaxed);
                return DynamicStep::Idle;
            }

            const bool discoveryAhead = ! queue.producerDone && ! queue.listingDone && GetDiscoveryMode(rootContext) == FILESYSTEM_DISCOVERY_AHEAD;
            const size_t activeLimit  = Common::FileOperations::DiscoveryWorkerLimit(static_cast<size_t>(concurrency), queue.items.size(), discoveryAhead);
            if (queue.activeItems >= activeLimit)
            {
                idleAtLimitDispatches.fetch_add(1u, std::memory_order_relaxed);
                return DynamicStep::Idle;
            }

            item = std::move(queue.items.front());
            queue.items.pop_front();
            queue.ready.fetch_sub(1, std::memory_order_release);
            ++queue.activeItems;
            queue.maxActiveItems = (std::max)(queue.maxActiveItems, queue.activeItems);
        }

        workItemDispatches.fetch_add(1, std::memory_order_relaxed);
        queue.cv.notify_all();
        if (item.kind == RecursiveCopyWorkKind::Finish)
        {
            finishActiveItem(item.retainedPathBytes);
            return DynamicStep::Finished;
        }
        // One worker copies in listing order on its root's stable stream (zero for a standalone
        // call). A bulk root keeps its owning scheduler stream so shared-cookie roots cannot collide.
        const uint64_t progressStreamId = concurrency > 1u ? schedulerStreamId : rootContext.progressStreamId;
        OperationContext& streamContext = getStreamContext(progressStreamId);
        processWork(streamContext, item);
        // Release before finishing the item: a quiescent queue then means every completion this
        // item could trigger (metadata restore, parent release) already ran.
        releaseDirectory(std::move(item.directory), streamContext);
        finishActiveItem(item.retainedPathBytes);

        std::scoped_lock lock(queue.mutex);
        return queue.done ? DynamicStep::Finished : DynamicStep::Processed;
    });

    const HRESULT producerHr = runDirectory(rootContext, rootItem, 0u, nullptr, true);
    bool finishQueued        = false;
    if (SUCCEEDED(producerHr))
    {
        RecursiveCopyWorkItem finishItem{};
        finishItem.kind        = RecursiveCopyWorkKind::Finish;
        const HRESULT finishHr = enqueueWork(std::move(finishItem));
        if (FAILED(finishHr))
        {
            completeWithError(finishHr);
        }
        else
        {
            finishQueued = true;
        }
    }
    else
    {
        completeWithError(producerHr);
    }
    if (! finishQueued)
    {
        // The job dispatches a step only while queue.ready is nonzero and finishes only when a step
        // reports Finished: a walk that stopped with nothing queued wakes one step, which sees the stop
        // (queue.done is set) and finishes the job. Without it, WaitJob below would never return.
        queue.ready.fetch_add(1, std::memory_order_release);
    }
    {
        std::scoped_lock lock(queue.mutex);
        queue.producerDone = true;
        if (queue.items.empty() && queue.activeItems == 0u)
        {
            queue.done = true;
        }
    }
    if (! discoveryClosureReported)
    {
        // A walk that failed or stopped before its last listing closes its scope here, once.
        discoveryClosureReported        = true;
        const uint32_t finalQueuedItems = queuedWorkEntries();
        const HRESULT discoveryCloseHr  = ReportDiscoveryProgress(rootContext, discoveredBytes, discoveredFiles, discoveredDirectories, finalQueuedItems, true);
        if (FAILED(discoveryCloseHr))
        {
            completeWithError(discoveryCloseHr);
        }
    }
    queue.cv.notify_all();
    GetSharedFileOpsJobScheduler().NotifyDynamicWorkAvailable(job);

    GetSharedFileOpsJobScheduler().WaitJob(job);

    // A stop abandons the items still queued. No worker runs now: releasing their directory counts lets
    // every created directory complete as usual, children first, and restore its metadata.
    {
        std::deque<RecursiveCopyWorkItem> abandoned;
        {
            std::scoped_lock lock(queue.mutex);
            abandoned.swap(queue.items);
        }
        for (RecursiveCopyWorkItem& abandonedItem : abandoned)
        {
            if (abandonedItem.directory)
                abandonedItem.directory->complete.store(false, std::memory_order_release);
            releaseDirectory(std::move(abandonedItem.directory), rootContext);
        }
    }

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
    // I35 S2: worker high-water (DISC-3: a flat folder reaches the configured concurrency) and the
    // number of directories that waited for their children after their enumeration ended.
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.ActiveWorkersMax", L"", 0u, queue.maxActiveItems, concurrency, producerHr);
    // Items the producer copied itself because none of the walk's items was in flight on a worker.
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.ProducerInlineSteps", L"", 0u, producerInlineSteps, concurrency, producerHr);
    // Dispatches that found nothing to run (at the active limit; with an empty queue) and the process CPU the
    // walk took, both against the walk's wall time.
    const uint64_t walkWallUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - walkStart).count());
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.IdleDispatches",
                      L"",
                      walkWallUs,
                      idleAtLimitDispatches.load(std::memory_order_relaxed),
                      idleEmptyDispatches.load(std::memory_order_relaxed),
                      producerHr);
    const uint64_t walkEndProcessCpuUs = ProcessCpuTimeUs();
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.ProcessCpuUs",
                      L"",
                      walkWallUs,
                      walkEndProcessCpuUs >= walkStartProcessCpuUs ? walkEndProcessCpuUs - walkStartProcessCpuUs : 0u,
                      concurrency,
                      producerHr);
    Debug::Perf::Emit(
        L"FileOps.CopyRecursiveParallel.ParkedDirectoriesMax", L"", 0u, parkedDirectoriesMax.load(std::memory_order_acquire), concurrency, producerHr);
    Debug::Perf::Emit(L"FileOps.CopyRecursiveParallel.CountAheadMaxRetainedBytes", L"", 0u, countAheadMaxRetainedBytes, kCountAheadNameBudgetBytes, producerHr);
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

    if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
    {
        const HRESULT failure = walkFailure.load(std::memory_order_acquire);
        // R6.7: the walk's own failure, before any cancellation its stop caused; a walk that a bulk
        // sibling's failure stopped was canceled.
        return FAILED(failure) ? returnFailure(failure) : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
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
    // I35 S2 (R6.6): a recursive directory always takes the producer/worker walker, one worker
    // included, so its discovery closes at its listing's end; CopyDirectoryInternal stays the
    // fallback when no scheduler workers are available.
    const bool walkDirectory = sourceKind == LocalCopyPathKind::RegularDirectory && context.recursive;

    // A root of a bulk call at concurrency 1 keeps the serial walker on the thread that runs it: a bulk
    // call closes its discovery once at its end anyway, and a sequential bulk call keeps every callback
    // on its caller's thread.
    if (walkDirectory && effectiveConcurrency <= 1u && (context.parallel != nullptr || context.discoveryTranslation != nullptr))
    {
        Debug::Perf::EmitCounter(L"FileOps.CopyRecursiveParallel.SerialFallback.BulkConcurrencyOne");
        return CopyDirectoryInternal(context, source, destination, bytesCopied);
    }

    if (walkDirectory)
    {
        return CopyDirectoryChildrenParallel(context, source, destination, flags, reparsePointPolicy, effectiveConcurrency, bytesCopied);
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

#ifdef ENABLE_TESTS
// Simulates another writer creating a different object at a Move destination name after the
// provider's absence probe and before its unbound rename. It fires only once the name is free,
// so for a case-only rename it lands on the temp route's vacated second-hop name.
void MaybeInjectMoveDestinationRaceForSelfTest(const std::wstring& destinationExtended) noexcept
{
    const std::wstring racePath = GetSelfTestEnvironmentString(kMoveDestinationRacePathEnvVar);
    if (racePath.empty() || ! OrdinalString::EqualsNoCase(destinationExtended, MakePathInfo(racePath).extended))
    {
        return;
    }

    wil::unique_hfile raced(::CreateFileW(destinationExtended.c_str(),
                                          GENERIC_WRITE,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          CREATE_NEW,
                                          FILE_ATTRIBUTE_NORMAL,
                                          nullptr));
    if (! raced)
    {
        return;
    }

    constexpr char racePayload[] = "race";
    DWORD written                = 0u;
    static_cast<void>(::WriteFile(raced.get(), racePayload, static_cast<DWORD>(sizeof(racePayload) - 1u), &written, nullptr));
    raced.reset();
    static_cast<void>(::SetEnvironmentVariableW(kMoveDestinationRacePathEnvVar.data(), nullptr));
    static_cast<void>(::SetEnvironmentVariableW(kMoveDestinationRaceFiredEnvVar.data(), L"1"));
}
#endif

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

#if defined(ENABLE_TESTS)
        // Only the existing pathname-provider fallback reaches this seam. The host's
        // identity-bound RenameIfUnchanged route is separate and never redirected here.
        const std::wstring secondLegFailurePath = GetSelfTestEnvironmentString(kCaseRenameSecondLegFailurePathEnvVar);
        const bool failSecondLegForSelfTest =
            context.type == FILESYSTEM_RENAME && ! secondLegFailurePath.empty() && sourceExtended == MakePathInfo(secondLegFailurePath).extended;
#endif
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
#if defined(ENABLE_TESTS)
            else if (failSecondLegForSelfTest)
            {
                static_cast<void>(SetEnvironmentVariableW(kCaseRenameRevertedEnvVar.data(), L"1"));
            }
#endif
        };

        hr = CheckCancel(context);
        if (FAILED(hr))
        {
            revertToSource();
            return hr;
        }

        bool renamedToDestination = false;
        DWORD secondLegError      = ERROR_SUCCESS;
#if defined(ENABLE_TESTS)
        if (failSecondLegForSelfTest)
        {
            // The first MoveFileExW above has actually succeeded. Record its real
            // temporary name, inject failure at leg two, then use ordinary rollback.
            static_cast<void>(SetEnvironmentVariableW(kCaseRenameSecondLegFailurePathEnvVar.data(), nullptr));
            static_cast<void>(SetEnvironmentVariableW(kCaseRenameFirstLegPathEnvVar.data(), tempPath.c_str()));
            secondLegError = ERROR_WRITE_FAULT;
        }
        else
#endif
        {
#ifdef ENABLE_TESTS
            MaybeInjectMoveDestinationRaceForSelfTest(destinationExtended);
#endif
            renamedToDestination = ::MoveFileExW(tempPath.c_str(), destinationExtended.c_str(), renameFlags) != FALSE;
            secondLegError       = renamedToDestination ? ERROR_SUCCESS : GetLastError();
        }
        if (! renamedToDestination)
        {
            revertToSource();
            return HRESULT_FROM_WIN32(secondLegError);
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
        // No rename was issued, so a vanished source is a known no-commit with its real error.
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
    // An overwrite grant is not a replacement receipt when the name is currently absent; only
    // an existing destination reaches the identity-bound RenameIfUnchanged route below.
    const bool allowOverwriteEffective = destinationExistedBeforeCopy && HasOverwriteGrant(context);
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
            if (mutationResult.originalStillPresent == FALSE)
                MarkNativeMoveSourceUnproven(context);
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

    // This route is unbound: the destination was absent at the probe, or this is a case-only
    // rename of the source itself (whose temp second hop targets the name it just vacated).
    // Never replace here, so a destination that races into existence fails the rename atomically
    // with ERROR_ALREADY_EXISTS and returns to the conflict prompt.
    constexpr DWORD renameFlags = 0u;

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
        const std::wstring entropyFailurePath   = GetSelfTestEnvironmentString(kCaseRenameEntropyFailurePathEnvVar);
        const std::wstring collisionPath        = GetSelfTestEnvironmentString(kCaseRenameTempCollisionPathEnvVar);
        const std::wstring secondLegFailurePath = GetSelfTestEnvironmentString(kCaseRenameSecondLegFailurePathEnvVar);
        const std::wstring destinationRacePath  = GetSelfTestEnvironmentString(kMoveDestinationRacePathEnvVar);
        forceCaseOnlyTempForSelfTest =
            (! entropyFailurePath.empty() && OrdinalString::EqualsNoCase(source.extended, MakePathInfo(entropyFailurePath).extended)) ||
            (! collisionPath.empty() && OrdinalString::EqualsNoCase(source.extended, MakePathInfo(collisionPath).extended)) ||
            (context.type == FILESYSTEM_RENAME && ! secondLegFailurePath.empty() && source.extended == MakePathInfo(secondLegFailurePath).extended) ||
            (! destinationRacePath.empty() && OrdinalString::EqualsNoCase(destination.extended, MakePathInfo(destinationRacePath).extended));
    }
    MaybeInjectMoveDestinationRaceForSelfTest(destination.extended);
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

    // A failed rename without MOVEFILE_COPY_ALLOWED is atomic, so every error here, including a
    // vanished source or a missing destination parent (ERROR_PATH_NOT_FOUND), is a known no-commit.
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

    HRESULT STDMETHODCALLTYPE PreDeleteItem(DWORD flags, IShellItem* item) noexcept override
    {
        const HRESULT hr = ReportItemPath(item, false);
        if (IsCancellationToHonour(hr))
        {
            // A canceled PreDeleteItem stops the Shell before it touches the item, as a refusal does.
            MarkUntouchedUnlessOffered();
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (RecyclePreDeleteWouldDestroy(flags))
        {
            // Failing PreDeleteItem stops the Shell before it touches the item; PostDeleteItem then
            // reports this refusal, and the known non-commit stands even if it never arrives.
            _recycleRefused = true;
            _mutationResult = ReduceRecyclePostDelete(true, kRecycleUnavailableStatus, false).receipt;
            return kRecycleUnavailableStatus;
        }
        _deleteOffered = true;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostDeleteItem([[maybe_unused]] DWORD flags, IShellItem* item, HRESULT hrDelete, IShellItem* newlyCreated) noexcept override
    {
#if defined(ENABLE_TESTS)
        JournalRecycledItemForTest(item, newlyCreated);
#endif
        const RecycleItemVerdict verdict = ReduceRecyclePostDelete(_recycleRefused, hrDelete, newlyCreated != nullptr);
        _mutationResult                  = verdict.receipt;
        if (verdict.status == kRecycleDestroyedStatus)
        {
            Debug::Error(L"FileSystem: Recycle Bin delete removed an item without a Recycle Bin item (hr={:#x})", static_cast<unsigned long>(hrDelete));
        }
        hrDelete = verdict.status;
#if defined(ENABLE_TESTS)
        if (_stopAfterDeletePathForSelfTest != nullptr && _context != nullptr &&
            MatchRecycleSeamPathForSelfTest(kRecycleShellCommitStopPathEnvVar, *_stopAfterDeletePathForSelfTest, true))
        {
            // Another item's failure stops the call once this item's commit has been observed.
            AwaitParallelStopForSelfTest(_context->parallel);
        }
#endif

        if (SUCCEEDED(hrDelete) && _context != nullptr)
        {
            if (! _workProgressAvailable)
            {
                AddCompletedItems(*_context, 1);
            }
            const HRESULT hr = ReportItemPath(item, false);
            if (IsCancellationToHonour(hr))
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
        if (IsCancellationToHonour(hr))
        {
            // Canceled before PreDeleteItem let the Shell act: the item is untouched.
            MarkUntouchedUnlessOffered();
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

    [[nodiscard]] bool WasRecycleRefused() const noexcept
    {
        return _recycleRefused;
    }

#if defined(ENABLE_TESTS)
    void ArmStopAfterDeleteForSelfTest(const PathInfo& path) noexcept
    {
        _stopAfterDeletePathForSelfTest = &path;
    }
#endif

private:
    // Only the caller's cancellation stops the Shell. Once the Shell holds this one item, another
    // item's failure stopping the provider call lets it run to its terminal observation, which is
    // then reported: an item left untouched there would be resumed after its discovery was already
    // counted, and a commit already observed would turn into an unknown outcome. The progress calls
    // answer that stop without asking the caller, so the caller is polled here, and a user Cancel
    // issued after the stop still stops the Shell (a large folder's walk, for example).
    [[nodiscard]] bool IsCancellationToHonour(HRESULT hr) const noexcept
    {
        if (hr != HRESULT_FROM_WIN32(ERROR_CANCELLED) && hr != E_ABORT)
        {
            return false;
        }
        const ParallelOperationState* const parallel = _context != nullptr ? _context->parallel : nullptr;
        if (parallel == nullptr || ! parallel->stopOnErrorRequested.load(std::memory_order_acquire) ||
            parallel->cancelRequested.load(std::memory_order_acquire))
        {
            return true;
        }
        return IsCancellationHr(CheckCallerCancel(*_context));
    }

    // Until PreDeleteItem has let the Shell act on the item, a cancellation leaves it untouched: a
    // known non-commit rather than the unknown default, so a caller may reach it again.
    void MarkUntouchedUnlessOffered() noexcept
    {
        if (! _deleteOffered)
        {
            _mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
        }
    }

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
    bool _recycleRefused              = false;
    bool _deleteOffered               = false; // PreDeleteItem let the Shell act on the item
    HRESULT _firstError               = S_OK;
    std::wstring _firstErrorPath;
    FileSystemItemMutationResult _mutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
#if defined(ENABLE_TESTS)
    const PathInfo* _stopAfterDeletePathForSelfTest = nullptr; // PostDeleteItem waits for another item's stop
#endif
};

struct RecycleBinBatchEntry
{
    unsigned long itemIndex = 0;
    PathInfo path{};
    HRESULT result      = S_OK;
    bool observed       = false;
    bool deleteOffered  = false; // PreDeleteItem saw this entry
    bool recycleRefused = false; // PreDeleteItem refused it because the Shell would delete it permanently
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

    HRESULT STDMETHODCALLTYPE PreDeleteItem(DWORD flags, IShellItem* item) noexcept override
    {
        RecycleBinBatchEntry* entry = FindEntry(item);
        if (entry != nullptr)
        {
            entry->deleteOffered = true;
        }

        const HRESULT hr = ReportItemPath(item, false);
        if (FAILED(hr))
        {
            _callbackFailure = hr;
            return IsCancellationHr(hr) ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : hr;
        }
        if (RecyclePreDeleteWouldDestroy(flags))
        {
            // The refusal ends the whole IFileOperation before it touches any item; the caller
            // replays the refused entry and every entry the Shell never offered one by one. The
            // sink-level flag survives an item path that matches no entry.
            _recycleRefused = true;
            if (entry != nullptr)
            {
                entry->recycleRefused = true;
                entry->result         = kRecycleUnavailableStatus;
                entry->mutationResult = ReduceRecyclePostDelete(true, kRecycleUnavailableStatus, false).receipt;
            }
            return kRecycleUnavailableStatus;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PostDeleteItem([[maybe_unused]] DWORD flags, IShellItem* item, HRESULT hrDelete, IShellItem* newlyCreated) noexcept override
    {
#if defined(ENABLE_TESTS)
        JournalRecycledItemForTest(item, newlyCreated);
#endif
        RecycleBinBatchEntry* entry      = FindEntry(item);
        const RecycleItemVerdict verdict = ReduceRecyclePostDelete(entry != nullptr && entry->recycleRefused, hrDelete, newlyCreated != nullptr);
        if (verdict.status == kRecycleDestroyedStatus)
        {
            Debug::Error(L"FileSystem: Recycle Bin batch delete removed an item without a Recycle Bin item (hr={:#x})", static_cast<unsigned long>(hrDelete));
        }
        hrDelete = verdict.status;
        if (entry != nullptr)
        {
            entry->result         = hrDelete;
            entry->observed       = true;
            entry->mutationResult = verdict.receipt;
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

    [[nodiscard]] bool WasRecycleRefused() const noexcept
    {
        return _recycleRefused;
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
    bool _recycleRefused     = false;
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
#ifdef ENABLE_TESTS
    if (MatchRecycleSeamPathForSelfTest(kRecycleShellCommitStopPathEnvVar, path, false))
    {
        progressSinkImpl->ArmStopAfterDeleteForSelfTest(path);
    }
#endif

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

#ifdef ENABLE_TESTS
    if (MatchRecycleSeamPathForSelfTest(kRecycleShellStopPathEnvVar, path, true))
    {
        // Another item's failure stops the call once the Shell holds this item, before it acts.
        AwaitParallelStopForSelfTest(context.parallel);
    }
#endif
    hr = fileOperation->PerformOperations();
    if (mutationResult != nullptr)
    {
        *mutationResult = progressSinkImpl->GetMutationResult();
    }
    if (progressSinkImpl->WasRecycleRefused())
    {
        // The Shell could not recycle the item and was stopped before deleting it permanently.
        Debug::Warning(L"FileSystem: Recycle Bin is unavailable for '{}'; the item was kept", path.display);
        return kRecycleUnavailableStatus;
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
        // An item the Shell destroyed before it was stopped stays a failure: its committed receipt
        // must never read as a Recycle once the cancellation is resolved.
        return progressSinkImpl->GetFirstError() == kRecycleDestroyedStatus ? kRecycleDestroyedStatus : HRESULT_FROM_WIN32(ERROR_CANCELLED);
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
    if (HRESULT injectedStatus = S_OK; TryInjectRecycleUnknownForSelfTest(path, injectedStatus, mutationResult))
    {
        return injectedStatus;
    }
    if (HRESULT injectedStatus = S_OK; TryInjectRecycleDestroyForSelfTest(path, injectedStatus, mutationResult))
    {
        return injectedStatus;
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

HRESULT DeleteToRecycleBinBatchedCore(OperationContext& context, std::span<RecycleBinBatchEntry> items, bool* batchStarted, bool* recycleRefused) noexcept
{
    if (batchStarted != nullptr)
    {
        *batchStarted = false;
    }
    if (recycleRefused != nullptr)
    {
        *recycleRefused = false;
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
    if (recycleRefused != nullptr)
    {
        *recycleRefused = progressSinkImpl->WasRecycleRefused();
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

HRESULT DeleteToRecycleBinBatched(OperationContext& context, std::span<RecycleBinBatchEntry> items, bool* batchStarted, bool* recycleRefused) noexcept
{
    if (batchStarted != nullptr)
    {
        *batchStarted = false;
    }
    if (recycleRefused != nullptr)
    {
        *recycleRefused = false;
    }

    if (items.size() < 2)
    {
        return E_INVALIDARG;
    }

#ifdef ENABLE_TESTS
    // Stands in for a Shell batch that recycles its armed entries while another item's failure stops
    // the call: each armed entry is removed and observed committed, the others are never offered,
    // and the batch then holds (bounded) until the call stops, so its observed entries are reported
    // after that stop.
    if (std::ranges::any_of(items, [](const RecycleBinBatchEntry& entry) noexcept {
        return MatchRecycleSeamPathForSelfTest(kRecycleBatchCommitPathEnvVar, entry.path, false);
    }))
    {
        bool allObserved = true;
        for (RecycleBinBatchEntry& entry : items)
        {
            if (! MatchRecycleSeamPathForSelfTest(kRecycleBatchCommitPathEnvVar, entry.path, true))
            {
                allObserved = false;
                continue;
            }
            const HRESULT removeHr           = ::DeleteFileW(entry.path.extended.c_str()) ? S_OK : HRESULT_FROM_WIN32(::GetLastError());
            const RecycleItemVerdict verdict = ReduceRecyclePostDelete(false, removeHr, SUCCEEDED(removeHr));
            entry.deleteOffered              = true;
            entry.observed                   = true;
            entry.result                     = verdict.status;
            entry.mutationResult             = verdict.receipt;
        }
        if (batchStarted != nullptr)
        {
            *batchStarted = true;
        }
        AwaitParallelStopForSelfTest(context.parallel);
        return allObserved ? S_OK : E_FAIL;
    }

    // Stands in for a Shell batch that loses one entry's terminal truth next to a known failure: the
    // armed entry is offered without a terminal observation, each entry armed to fail is observed as
    // a known non-commit, the others are never offered, and the batch fails. Nothing is touched.
    if (std::ranges::any_of(items, [](const RecycleBinBatchEntry& entry) noexcept {
        return MatchRecycleSeamPathForSelfTest(kRecycleBatchUnknownPathEnvVar, entry.path, false);
    }))
    {
        for (RecycleBinBatchEntry& entry : items)
        {
            FileSystemItemMutationResult failure{};
            if (MatchRecycleSeamPathForSelfTest(kRecycleBatchUnknownPathEnvVar, entry.path, true))
            {
                entry.deleteOffered = true;
            }
            else if (TryInjectRecycleFailureForSelfTest(entry.path, &failure))
            {
                entry.deleteOffered  = true;
                entry.observed       = true;
                entry.result         = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
                entry.mutationResult = failure;
            }
        }
        if (batchStarted != nullptr)
        {
            *batchStarted = true;
        }
        return E_FAIL;
    }

    // A batched IFileOperation would bypass the per-item Recycle seams, so a batch holding an armed
    // path is not started and every entry takes the single-item route instead.
    if (std::ranges::any_of(items,
                            [](const RecycleBinBatchEntry& entry) noexcept
    {
        return MatchRecycleFailureDirectoryForSelfTest(entry.path) || MatchRecycleSeamPathForSelfTest(kRecycleFailurePathEnvVar, entry.path, false) ||
               MatchRecycleSeamPathForSelfTest(kRecycleUnknownPathEnvVar, entry.path, false);
    }))
    {
        return E_NOTIMPL;
    }
#endif

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
        return RunShellOperationOnDedicatedStaThread([&]() noexcept { return DeleteToRecycleBinBatchedCore(context, items, batchStarted, recycleRefused); });
    }
    if (FAILED(coInitHr))
    {
        return coInitHr;
    }

    return DeleteToRecycleBinBatchedCore(context, items, batchStarted, recycleRefused);
}

HRESULT DeleteDirectoryRecursive(OperationContext& context, const PathInfo& path) noexcept;

HRESULT DeletePathInternal(OperationContext& context,
                           const PathInfo& path,
                           bool discoveryAlreadyReported,
                           FileSystemItemMutationResult* mutationResult) noexcept
{
    // Nothing is touched before the Recycle route or the permanent delete below starts, so every
    // early return, a cancellation by another item's stop included, is a known non-commit.
    if (mutationResult != nullptr)
    {
        *mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
    }
    HRESULT hr = SetProgressPaths(context, path.display.c_str(), nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

#ifdef ENABLE_TESTS
    if (context.useRecycleBin)
    {
        HoldRecycleForSiblingStopForSelfTest(context, path);
    }
#endif

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
    // The Recycle route resets the receipt to the unknown default once the Shell may act, because
    // IFileOperation can lose terminal per-item truth; its progress sink is the only authority
    // allowed to replace that default.
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
            return reportKnownDeletion(DeleteOpenedPathNoFollow(context, snapshot, true, path.extended));
        }

        if (! context.recursive)
        {
            return reportKnownDeletion(DeleteOpenedPathNoFollow(context, snapshot, true, path.extended));
        }

        if (snapshot.handle)
        {
            snapshot.handle.reset();
        }
        return reportKnownDeletion(DeleteDirectoryRecursive(context, path));
    }

    return reportKnownDeletion(DeleteOpenedPathNoFollow(context, snapshot, true, path.extended));
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
    DeleteWalkFrame()                                  = default;
    DeleteWalkFrame(const DeleteWalkFrame&)            = delete;
    DeleteWalkFrame& operator=(const DeleteWalkFrame&) = delete;
    DeleteWalkFrame(DeleteWalkFrame&&)                 = delete;
    DeleteWalkFrame& operator=(DeleteWalkFrame&&)      = delete;

    // A separate short-lived listing may close, but ancestors stay pinned against rename,
    // write/reparse replacement and deletion until every child callback has joined. Dispose
    // this exact handle at completion; identity-less SMB needs no close/reopen authority gap.
    DeleteHandleSnapshot directory;
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
        if (frameDeletesRoot)
        {
            const HRESULT pinHr = OpenPathForDeleteNoFollow(path.extended, frame->directory, FILE_SHARE_READ);
            if (FAILED(pinHr))
                return pinHr;
        }
        // A bound contents-only root already has the caller's DELETE/share-read pin.
        // Opening a second exclusive deletion pin on that same object would conflict.
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

        if (frame.deleteRoot && (! IsDirectory(frame.directory.attributes) || IsReparsePoint(frame.directory.attributes)))
        {
            hr = DeleteOpenedPathNoFollow(context, frame.directory, true, path.extended);
            return SUCCEEDED(hr) ? S_FALSE : hr;
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
            hr = DeleteOpenedPathNoFollow(context, frame.directory, true, path.extended);
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

            hr = DeleteOpenedPathNoFollow(context, frame.directory, true, path.extended);
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
            return DeleteOpenedPathNoFollow(context, snapshot, true, entry.path.extended);
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

#if defined(ENABLE_TESTS)
// I35 S4 round 2: parks a delete by name of the armed root after the check of its listing and before its
// first deletion, so a self-test can write into the tree there. The host self-test arms it through the
// process environment; it fires once, and a lost test cannot park the worker for more than 30 seconds.
void PauseDeleteByNameBeforeDeletionsForSelfTest(std::wstring_view providerPath) noexcept
{
    constexpr std::wstring_view rootVariable    = L"REDSALAMANDER_FILEOPS_BYNAME_DELETE_PAUSE_ROOT";
    constexpr std::wstring_view enteredVariable = L"REDSALAMANDER_FILEOPS_BYNAME_DELETE_PAUSE_ENTERED";
    constexpr std::wstring_view releaseVariable = L"REDSALAMANDER_FILEOPS_BYNAME_DELETE_PAUSE_RELEASE";
    const std::wstring root                     = GetSelfTestEnvironmentString(rootVariable);
    if (root.empty() || ! OrdinalString::EqualsNoCase(root, providerPath))
    {
        return;
    }
    if (! ::SetEnvironmentVariableW(rootVariable.data(), nullptr) || ! ::SetEnvironmentVariableW(enteredVariable.data(), L"1"))
    {
        return;
    }
    // The host self-test scales the deadline with its timeout multiplier.
    constexpr std::wstring_view timeoutVariable = L"REDSALAMANDER_FILEOPS_BYNAME_DELETE_PAUSE_TIMEOUT_MS";
    const std::wstring timeoutText              = GetSelfTestEnvironmentString(timeoutVariable);
    wchar_t* timeoutEnd                         = nullptr;
    const unsigned long configuredMs            = timeoutText.empty() ? 0ul : std::wcstoul(timeoutText.c_str(), &timeoutEnd, 10);
    const ULONGLONG timeoutMs                   = configuredMs != 0ul && timeoutEnd != nullptr && *timeoutEnd == L'\0' ? configuredMs : 30'000u;
    const ULONGLONG deadline                    = ::GetTickCount64() + timeoutMs;
    while (GetSelfTestEnvironmentString(releaseVariable) != L"1" && ::GetTickCount64() < deadline)
    {
        ::Sleep(10u);
    }
}
#endif

// I35 R7.3 (pinned tree): a listed folder's contents, reported as discovered when the listing reaches them.
void ReportDeleteByNameListedFolder(void* cookie, uint64_t files, uint64_t directories, uint64_t bytes) noexcept
{
    if (cookie != nullptr)
    {
        static_cast<void>(ReportDeleteDiscoveryBatch(*static_cast<OperationContext*>(cookie), files, directories, bytes, 0u));
    }
}

// At most this many folders of a pinned tree are held open at once (plus the chain above them): a larger tree
// is deleted in chunks of sibling subtrees that fit, each held whole.
constexpr size_t kDeleteByNameHeldFolderBudget = 1024u;

[[nodiscard]] size_t DeleteByNameHeldFolderBudget() noexcept
{
#if defined(ENABLE_TESTS)
    // A self-test makes the budget small, so a small tree goes in chunks.
    const std::wstring configured = GetSelfTestEnvironmentString(L"REDSALAMANDER_FILEOPS_BYNAME_HOLD_BUDGET");
    wchar_t* end                  = nullptr;
    const unsigned long budget    = configured.empty() ? 0ul : std::wcstoul(configured.c_str(), &end, 10);
    if (budget != 0ul && end != nullptr && *end == L'\0')
    {
        return static_cast<size_t>(budget);
    }
#endif
    return kDeleteByNameHeldFolderBudget;
}

// One folder of the pinned tree of a delete by name.
struct ByNameFolderNode final
{
    ByNameFolderNode()                                   = default;
    ByNameFolderNode(const ByNameFolderNode&)            = delete;
    ByNameFolderNode& operator=(const ByNameFolderNode&) = delete;
    ByNameFolderNode(ByNameFolderNode&&)                 = default;
    ByNameFolderNode& operator=(ByNameFolderNode&&)      = default;

    size_t entryIndex = std::numeric_limits<size_t>::max(); // its manifest entry (max: the selected folder)
    size_t parent     = 0u;
    std::vector<size_t> files;    // manifest indices of the files and links directly inside
    std::vector<size_t> folders;  // node indices of the folders directly inside
    uint64_t subtreeFolders = 0u; // folders below this one
    DeleteHandleSnapshot hold;    // held, sharing only reads, from its check to its own deletion
};

// The deletions of one pinned tree: its folders, the concurrency, the first stop and the entries deleted.
struct ByNameTreeWalk final
{
    ByNameTreeWalk(OperationContext& operationContext, const PathInfo& rootPath, const FileSystemInternal::LocalDeleteManifest& listed) noexcept
        : context(operationContext),
          root(rootPath),
          manifest(listed)
    {
    }
    ByNameTreeWalk(const ByNameTreeWalk&)            = delete;
    ByNameTreeWalk& operator=(const ByNameTreeWalk&) = delete;

    OperationContext& context;
    const PathInfo& root;
    const FileSystemInternal::LocalDeleteManifest& manifest;
    std::vector<ByNameFolderNode> nodes;
    unsigned int concurrency = 1u;
    size_t holdBudget        = kDeleteByNameHeldFolderBudget;
    std::atomic<uint64_t> deleted{0u};
    std::atomic<bool> stopped{false};
    std::mutex stopMutex;
    HRESULT stopHr = S_OK;
    std::wstring_view stopStep;
    std::wstring stopAt;

    [[nodiscard]] bool Stopped() const noexcept
    {
        return stopped.load(std::memory_order_acquire);
    }

    // The first stop wins; the walk starts nothing after it.
    void Stop(const HRESULT hr, const std::wstring_view step, const std::wstring_view at) noexcept
    {
        std::scoped_lock lock(stopMutex);
        if (stopped.load(std::memory_order_acquire))
        {
            return;
        }
        stopHr   = hr;
        stopStep = step;
        stopAt.assign(at);
        stopped.store(true, std::memory_order_release);
    }
};

// The folders of the manifest as a tree. The canonical order puts every folder before what it holds, so each
// entry's folder is already known when the entry is reached.
[[nodiscard]] HRESULT BuildByNameFolderTree(ByNameTreeWalk& walk) noexcept
{
    const std::vector<FileSystemInternal::LocalDeleteManifestEntry>& entries = walk.manifest.entries;
    walk.nodes.clear();
    walk.nodes.reserve(static_cast<size_t>(walk.manifest.directories) + 1u);
    walk.nodes.emplace_back();
    std::unordered_map<std::wstring_view, size_t> folderByPath;
    folderByPath.reserve(static_cast<size_t>(walk.manifest.directories) + 1u);
    folderByPath.emplace(std::wstring_view(), 0u);
    for (size_t index = 0u; index < entries.size(); ++index)
    {
        const std::wstring_view path(entries[index].relativePath);
        const size_t separator             = path.find_last_of(L'\\');
        const std::wstring_view parentPath = separator == std::wstring_view::npos ? std::wstring_view() : path.substr(0u, separator);
        const auto parent                  = folderByPath.find(parentPath);
        if (parent == folderByPath.end())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        const size_t parentNode = parent->second;
        if (IsDirectory(entries[index].attributes) && ! IsReparsePoint(entries[index].attributes))
        {
            const size_t node = walk.nodes.size();
            walk.nodes.emplace_back();
            walk.nodes[node].entryIndex = index;
            walk.nodes[node].parent     = parentNode;
            walk.nodes[parentNode].folders.push_back(node);
            folderByPath.emplace(path, node);
        }
        else
        {
            walk.nodes[parentNode].files.push_back(index);
        }
    }
    // A folder's node comes after its parent's, so one backward pass adds every subtree up.
    for (size_t node = walk.nodes.size(); node-- > 1u;)
    {
        walk.nodes[walk.nodes[node].parent].subtreeFolders += 1u + walk.nodes[node].subtreeFolders;
    }
    return S_OK;
}

// Opens a listed entry no-follow for DELETE, sharing only reads, and keeps the handle only while its facts still
// equal the listed ones (a folder's last-write time excepted). A listed entry that is gone or changed is
// ERROR_REVISION_MISMATCH. Every folder above it is held already, so its path cannot be redirected.
[[nodiscard]] HRESULT OpenListedEntryForDeleteByName(const std::wstring& extendedPath,
                                                     const FileSystemInternal::LocalDeleteManifestEntry& listed,
                                                     DeleteHandleSnapshot& snapshot) noexcept
{
    snapshot = {};
    wil::unique_handle handle(::CreateFileW(extendedPath.c_str(),
                                            DELETE | FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES,
                                            FILE_SHARE_READ,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                            nullptr));
    if (! handle)
    {
        const DWORD error = ::GetLastError();
        if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_OPEN_FAILED);
    }
    BY_HANDLE_FILE_INFORMATION information{};
    if (! ::GetFileInformationByHandle(handle.get(), &information))
    {
        const DWORD error = ::GetLastError();
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_GEN_FAILURE);
    }
    if (! FileSystemInternal::LocalDeleteManifestEntryMatches(listed, information))
    {
        Debug::Perf::EmitCounter(L"FileOps.Local.DeleteByNameRefused");
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
    snapshot.handle     = std::move(handle);
    snapshot.attributes = information.dwFileAttributes;
    snapshot.fileBytes  = IsDirectory(information.dwFileAttributes)
                              ? 0u
                              : (static_cast<uint64_t>(information.nFileSizeHigh) << 32u) | static_cast<uint64_t>(information.nFileSizeLow);
    return S_OK;
}

// The progress, cancellation and path of one step on `worker`; a failure stops the walk.
[[nodiscard]] bool BeginByNameStep(ByNameTreeWalk& walk, OperationContext& worker, const std::wstring& display, const std::wstring_view step) noexcept
{
    HRESULT hr = SetProgressPaths(worker, display.c_str(), nullptr);
    if (SUCCEEDED(hr))
    {
        hr = CheckCancel(worker);
    }
    if (SUCCEEDED(hr))
    {
        hr = ReportProgress(worker, 0, 0);
    }
    if (FAILED(hr))
    {
        walk.Stop(hr, step, display);
        return false;
    }
    return true;
}

// A file or link of the pinned tree: checked on its own handle, then deleted through it.
void DeleteByNameListedEntry(ByNameTreeWalk& walk, OperationContext& worker, const size_t entryIndex) noexcept
{
    const FileSystemInternal::LocalDeleteManifestEntry& entry = walk.manifest.entries[entryIndex];
    const std::wstring display                                = AppendPath(walk.root.display, entry.relativePath);
    const std::wstring extended                               = AppendPath(walk.root.extended, entry.relativePath);
    if (! BeginByNameStep(walk, worker, display, L"entry"))
    {
        return;
    }
    DeleteHandleSnapshot child{};
    HRESULT hr = OpenListedEntryForDeleteByName(extended, entry, child);
    if (FAILED(hr))
    {
        walk.Stop(hr, hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) ? L"compare-entry" : L"entry", display);
        return;
    }
    hr = DeleteOpenedPathNoFollow(worker, child, true, extended);
    if (FAILED(hr))
    {
        walk.Stop(hr, L"entry", display);
        return;
    }
    walk.deleted.fetch_add(1u, std::memory_order_acq_rel);
}

// A folder of the pinned tree, held from its check until its own deletion: no other program can rename it, turn
// it into a link or delete it meanwhile, so nothing below it can be reached through another object.
void HoldByNameListedFolder(ByNameTreeWalk& walk, OperationContext& worker, const size_t node) noexcept
{
    ByNameFolderNode& folder                                  = walk.nodes[node];
    const FileSystemInternal::LocalDeleteManifestEntry& entry = walk.manifest.entries[folder.entryIndex];
    const std::wstring display                                = AppendPath(walk.root.display, entry.relativePath);
    const std::wstring extended                               = AppendPath(walk.root.extended, entry.relativePath);
    if (! BeginByNameStep(walk, worker, display, L"folder"))
    {
        return;
    }
    const HRESULT hr = OpenListedEntryForDeleteByName(extended, entry, folder.hold);
    if (FAILED(hr))
    {
        walk.Stop(hr, hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) ? L"compare-folder" : L"folder", display);
    }
}

// A held folder once everything listed below it is gone, deleted through the handle that holds it: its removal
// fails while it holds anything the listing never saw.
void DeleteByNameHeldFolder(ByNameTreeWalk& walk, OperationContext& worker, const size_t node) noexcept
{
    ByNameFolderNode& folder   = walk.nodes[node];
    const std::wstring& path   = walk.manifest.entries[folder.entryIndex].relativePath;
    const std::wstring display = AppendPath(walk.root.display, path);
    if (! BeginByNameStep(walk, worker, display, L"folder"))
    {
        return;
    }
    const HRESULT hr = DeleteOpenedPathNoFollow(worker, folder.hold, true, AppendPath(walk.root.extended, path));
    if (FAILED(hr))
    {
        walk.Stop(hr, L"folder", display);
        return;
    }
    walk.deleted.fetch_add(1u, std::memory_order_acq_rel);
}

// Runs `step` for every item of one pass on the shared scheduler with the walk's concurrency (inline for one
// item, or when that budget is 1); nothing new starts after the first stop. Each worker has its own context
// sharing the call's progress, cancellation and discovery, as the recursive Delete walk's workers do. A pass the
// scheduler ended before every item was reached (its shutdown) stops the walk.
template <typename Step> void RunByNamePass(ByNameTreeWalk& walk, std::span<const size_t> items, const Step& step) noexcept
{
    if (items.empty() || walk.Stopped())
    {
        return;
    }
    OperationContext& context = walk.context;
    const bool parallelPass =
        walk.concurrency > 1u && items.size() > 1u && context.parallel == nullptr && GetSharedFileOpsJobScheduler().EnsureWorkersAvailable();
    if (! parallelPass)
    {
        for (const size_t item : items)
        {
            if (walk.Stopped())
            {
                return;
            }
            step(context, item);
        }
        return;
    }

    ParallelOperationState parallel{};
    parallel.startTick = GetTickCount64();
    parallel.bandwidthLimitBytesPerSecond.store(context.options ? context.options->bandwidthLimitBytesPerSecond : 0ull, std::memory_order_release);
    const FileSystemFlags workerFlags = static_cast<FileSystemFlags>(
        FILESYSTEM_FLAG_RECURSIVE | (context.allowReplaceReadonly ? FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY : FILESYSTEM_FLAG_NONE));
    std::atomic<size_t> reached{0u};
    context.parallel     = &parallel;
    auto restoreParallel = wil::scope_exit([&]() noexcept { context.parallel = nullptr; });
    auto job             = GetSharedFileOpsJobScheduler().StartJob(context.options,
                                                                   walk.concurrency,
                                                                   items.size(),
                                                                   [&](size_t index, uint64_t schedulerStreamId) noexcept
    {
        reached.fetch_add(1u, std::memory_order_acq_rel);
        if (index >= items.size() || walk.Stopped())
        {
            return;
        }
        OperationContext worker{};
        InitializeOperationContext(
            worker, FILESYSTEM_DELETE, workerFlags, context.options, context.callback, context.callbackCookie, 0, context.reparsePointPolicy);
        worker.parallel                = &parallel;
        worker.progressStreamId        = schedulerStreamId;
        worker.deleteConcurrencyBudget = 1u;
        worker.deleteDiscovery         = context.deleteDiscovery;
        step(worker, items[index]);
    });
    GetSharedFileOpsJobScheduler().WaitJob(job);
    context.parallel = nullptr;
    restoreParallel.release();
    AddCompletedItems(context, parallel.completedItems.load(std::memory_order_acquire));
    AddCompletedBytes(context, parallel.completedBytes.load(std::memory_order_acquire));
    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        walk.Stop(HRESULT_FROM_WIN32(ERROR_CANCELLED), L"cancel", walk.root.display);
    }
    if (reached.load(std::memory_order_acquire) != items.size())
    {
        walk.Stop(HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED), L"pass", walk.root.display);
    }
}

// Sibling subtrees that fit the hold budget together: every folder held level by level from the top, then every
// file and link of the chunk, then the folders deepest level first (each level at the same time).
void DeleteByNameChunk(ByNameTreeWalk& walk, std::span<const size_t> chunkRoots) noexcept
{
    std::vector<std::vector<size_t>> levels;
    levels.emplace_back(chunkRoots.begin(), chunkRoots.end());
    for (;;)
    {
        std::vector<size_t> next;
        for (const size_t node : levels.back())
        {
            next.insert(next.end(), walk.nodes[node].folders.begin(), walk.nodes[node].folders.end());
        }
        if (next.empty())
        {
            break;
        }
        levels.push_back(std::move(next));
    }
    const auto hold = [&walk](OperationContext& worker, const size_t node) noexcept { HoldByNameListedFolder(walk, worker, node); };
    for (const std::vector<size_t>& level : levels)
    {
        RunByNamePass(walk, level, hold);
        if (walk.Stopped())
        {
            return;
        }
    }
    // The chunk's folders in turn (the first file of each, then the second...): deletions in one folder wait on
    // each other, so the entries deleted at the same time rarely share a folder.
    std::vector<const std::vector<size_t>*> folderFiles;
    size_t chunkFiles   = 0u;
    size_t mostInFolder = 0u;
    for (const std::vector<size_t>& level : levels)
    {
        for (const size_t node : level)
        {
            const std::vector<size_t>& inFolder = walk.nodes[node].files;
            if (! inFolder.empty())
            {
                folderFiles.push_back(&inFolder);
                chunkFiles += inFolder.size();
                mostInFolder = (std::max)(mostInFolder, inFolder.size());
            }
        }
    }
    std::vector<size_t> files;
    files.reserve(chunkFiles);
    for (size_t position = 0u; position < mostInFolder; ++position)
    {
        for (const std::vector<size_t>* inFolder : folderFiles)
        {
            if (position < inFolder->size())
            {
                files.push_back((*inFolder)[position]);
            }
        }
    }
    RunByNamePass(walk, files, [&walk](OperationContext& worker, const size_t entry) noexcept { DeleteByNameListedEntry(walk, worker, entry); });
    if (walk.Stopped())
    {
        return;
    }
    const auto removeFolder = [&walk](OperationContext& worker, const size_t node) noexcept { DeleteByNameHeldFolder(walk, worker, node); };
    for (auto level = levels.rbegin(); level != levels.rend(); ++level)
    {
        RunByNamePass(walk, *level, removeFolder);
        if (walk.Stopped())
        {
            return;
        }
    }
}

// Everything listed below the held folder `node` (never the folder itself): its own files and links first, then
// its sub-folders in chunks that fit the hold budget; a sub-folder whose subtree alone does not fit is held and
// gone through the same way, one stack level each. Depth is bounded by the listing (kTraversalMaxDepth), and a
// nearly exhausted stack (a deep caller, Debug frames) stops the delete, as the recursive Copy walkers do.
void DeleteByNameFolderContents(ByNameTreeWalk& walk, const size_t node) noexcept
{
    RunByNamePass(
        walk, walk.nodes[node].files, [&walk](OperationContext& worker, const size_t entry) noexcept { DeleteByNameListedEntry(walk, worker, entry); });
    std::vector<size_t> group;
    uint64_t groupFolders = 0u;
    const auto flush      = [&]() noexcept
    {
        if (! group.empty() && ! walk.Stopped())
        {
            DeleteByNameChunk(walk, group);
        }
        group.clear();
        groupFolders = 0u;
    };
    for (const size_t child : walk.nodes[node].folders)
    {
        if (walk.Stopped())
        {
            return;
        }
        const uint64_t size = 1u + walk.nodes[child].subtreeFolders;
        if (size > walk.holdBudget)
        {
            flush();
            if (walk.Stopped())
            {
                return;
            }
            if (RecursiveCopyStackIsLow())
            {
                walk.Stop(HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW),
                          L"folder",
                          AppendPath(walk.root.display, walk.manifest.entries[walk.nodes[child].entryIndex].relativePath));
                return;
            }
            const std::array<size_t, 1u> single{child};
            RunByNamePass(walk, single, [&walk](OperationContext& worker, const size_t folder) noexcept { HoldByNameListedFolder(walk, worker, folder); });
            if (walk.Stopped())
            {
                return;
            }
            DeleteByNameFolderContents(walk, child);
            RunByNamePass(walk, single, [&walk](OperationContext& worker, const size_t folder) noexcept { DeleteByNameHeldFolder(walk, worker, folder); });
            continue;
        }
        if (groupFolders + size > walk.holdBudget)
        {
            flush();
        }
        group.push_back(child);
        groupFolders += size;
    }
    flush();
}

// I35 R7.3 (D-7.3, pinned tree): the selected root of a delete by name deletes exactly the tree that was
// pinned. The root is opened once, no-follow, for DELETE (and, for a folder, listing) while sharing only reads,
// so no other program can write, rename or delete the item itself after its facts are compared on that very
// handle. A folder is then listed again through that handle, its whole subtree no-follow (a link or junction is
// a leaf), its folders at the same time level by level, and reported as discovered as it is listed; the digest
// of that manifest must equal the one pinned before the card and re-checked after the answer. Any difference (or
// an exact identity the location now reports) is ERROR_REVISION_MISMATCH with nothing deleted. Only the listed
// entries are then deleted, each through its own no-follow handle that shares only reads, once that handle's
// facts equal the listed ones; every folder is held, sharing only reads, from its check until its own deletion,
// so no path below it can be redirected. Files and links go first, then folders deepest level first, each pass
// on the shared scheduler with the delete's concurrency (a tree with more folders than the hold budget goes in
// chunks of sibling subtrees), and the root last, through the compared handle. A listed entry that changed or is
// gone stops the delete with ERROR_REVISION_MISMATCH; a folder that holds anything the listing never saw cannot
// be removed and stops it too; nothing new starts after a stop, and nothing is listed again. Until the first
// deletion every return is a known non-commit; a stop after it reports the partial receipt (committed, original
// still present). The trace names the step with the deleted and listed counts; the root's discovery is reported
// only once its facts matched, which tells the host that the check passed.
HRESULT DeleteRootByNameInternal(OperationContext& context,
                                 const PathInfo& path,
                                 const wchar_t* providerPath,
                                 const FileSystemDeleteIdentity& expected,
                                 FileSystemItemMutationResult& mutationResult) noexcept
{
    mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
    HRESULT hr     = SetProgressPaths(context, path.display.c_str(), nullptr);
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

    const bool pinnedFolder = expected.isDirectory != FALSE;
    DeleteHandleSnapshot snapshot{};
    snapshot.handle.reset(::CreateFileW(path.extended.c_str(),
                                        DELETE | FILE_READ_ATTRIBUTES | FILE_WRITE_ATTRIBUTES | (pinnedFolder ? FILE_LIST_DIRECTORY : 0u),
                                        FILE_SHARE_READ,
                                        nullptr,
                                        OPEN_EXISTING,
                                        FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                        nullptr));
    if (! snapshot.handle)
    {
        // A sharing violation, access denial or pending delete: the facts were never compared and
        // nothing is deleted; the caller offers Retry, never another route.
        const DWORD error    = ::GetLastError();
        const HRESULT openHr = HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_FILE_NOT_FOUND);
        Debug::Warning(L"FileSystem: delete by name of '{}' could not open it: step=open openHr=0x{:08X} identityHr=not compared.",
                       path.display,
                       static_cast<unsigned long>(openHr));
        return openHr;
    }

    const std::wstring_view expectedText(expected.identity, ::wcsnlen(expected.identity, std::size(expected.identity)));
    const size_t treeFacts = expectedText.find(L"|m=");
    const auto refuse      = [&](const HRESULT corroborationHr, const size_t listed) noexcept -> HRESULT
    {
        if (IsCancellationHr(corroborationHr))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        // ERROR_NOT_SUPPORTED here means the location now reports an exact identity: a delete by name
        // must never replace exact authority, so it is refused like any other difference.
        const HRESULT refusal = SUCCEEDED(corroborationHr) || corroborationHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ||
                                        corroborationHr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH)
                                    ? HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH)
                                    : corroborationHr;
        Debug::Warning(L"FileSystem: delete by name of '{}' refused: step=compare openHr=0x00000000 identityHr=0x{:08X} listed={}.",
                       path.display,
                       static_cast<unsigned long>(refusal),
                       listed);
        Debug::Perf::EmitCounter(L"FileOps.Local.DeleteByNameRefused");
        return refusal;
    };

    // The item itself first: its own facts, without its tree.
    FileSystemDeleteIdentity live{};
    live.sizeBytes          = sizeof(live);
    HRESULT corroborationHr = FileSystemInternal::ResolveLocalDeleteCorroboration(snapshot.handle.get(), providerPath, live);
    if (FAILED(corroborationHr) || live.isDirectory != expected.isDirectory ||
        std::wstring_view(live.identity) != expectedText.substr(0u, treeFacts == std::wstring_view::npos ? expectedText.size() : treeFacts))
    {
        return refuse(corroborationHr, 0u);
    }
    BY_HANDLE_FILE_INFORMATION information{};
    if (! ::GetFileInformationByHandle(snapshot.handle.get(), &information))
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }
    snapshot.attributes = information.dwFileAttributes;
    if (! IsDirectory(snapshot.attributes))
    {
        snapshot.fileBytes = (static_cast<uint64_t>(information.nFileSizeHigh) << 32u) | static_cast<uint64_t>(information.nFileSizeLow);
    }
    hr = ReportDeleteDiscoveryObject(context, snapshot.attributes, snapshot.fileBytes, 1u);
    if (FAILED(hr))
    {
        return hr;
    }

    // Then the tree, listed again through this very handle; a tree over the listing's bounds now is not the
    // tree that was pinned within them.
    FileSystemInternal::LocalDeleteManifest manifest;
    if (pinnedFolder)
    {
        FileSystemInternal::LocalDeleteManifestListing listing{};
        listing.options        = context.options;
        listing.concurrency    = std::clamp(context.deleteConcurrencyBudget, 1u, 8u);
        listing.onFolderListed = &ReportDeleteByNameListedFolder;
        listing.cookie         = &context;
        corroborationHr        = FileSystemInternal::ListLocalDeleteManifest(snapshot.handle.get(), path.extended, listing, manifest);
        if (corroborationHr == HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA))
        {
            corroborationHr = HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        if (SUCCEEDED(corroborationHr))
        {
            corroborationHr = FileSystemInternal::ResolveLocalDeleteCorroboration(snapshot.handle.get(), providerPath, live, &manifest);
        }
        if (FAILED(corroborationHr) || live.isDirectory != expected.isDirectory || std::wstring_view(live.identity) != expectedText)
        {
            return refuse(corroborationHr, manifest.entries.size());
        }
    }
    const size_t listed = manifest.entries.size();
    // The call's items: every listed entry and the item itself (the host names "N of M listed items" from them).
    context.totalItems = static_cast<unsigned long>((std::min<uint64_t>)(static_cast<uint64_t>(listed) + 1u, std::numeric_limits<unsigned long>::max()));

#if defined(ENABLE_TESTS)
    PauseDeleteByNameBeforeDeletionsForSelfTest(providerPath != nullptr ? std::wstring_view(providerPath) : std::wstring_view());
#endif

    ByNameTreeWalk walk(context, path, manifest);
    walk.concurrency    = std::clamp(context.deleteConcurrencyBudget, 1u, 8u);
    walk.holdBudget     = DeleteByNameHeldFolderBudget();
    const auto stopHere = [&](const HRESULT stopHr, const std::wstring_view step, const std::wstring_view at) noexcept -> HRESULT
    {
        const uint64_t deleted = walk.deleted.load(std::memory_order_acquire);
        if (deleted != 0u)
        {
            // Part of the pinned tree is gone and the selected folder is still there: the partial-tree receipt.
            mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, TRUE, TRUE};
        }
        Debug::Warning(L"FileSystem: delete by name of '{}' stopped: step={} deleted={} of {} listed at='{}' hr=0x{:08X}.",
                       path.display,
                       step,
                       deleted,
                       listed,
                       at,
                       static_cast<unsigned long>(stopHr));
        Debug::Perf::Emit(L"FileOps.Local.DeleteByNameStopped", step, 0u, deleted, static_cast<uint64_t>(listed), stopHr);
        return stopHr;
    };

    const auto deletionsStarted = std::chrono::steady_clock::now();
    if (pinnedFolder && context.recursive && listed != 0u)
    {
        hr = BuildByNameFolderTree(walk);
        if (FAILED(hr))
        {
            return stopHere(hr, L"tree", path.display);
        }
        DeleteByNameFolderContents(walk, 0u);
        if (walk.Stopped())
        {
            return stopHere(walk.stopHr, walk.stopStep, walk.stopAt);
        }
    }

    // The consented root is removed even when read-only, as the exact route's disposition does.
    const bool allowReplaceReadonly = context.allowReplaceReadonly;
    context.allowReplaceReadonly    = true;
    hr                              = DeleteOpenedPathNoFollow(context, snapshot, true, path.extended);
    context.allowReplaceReadonly    = allowReplaceReadonly;
    if (FAILED(hr))
    {
        return stopHere(hr, L"disposition", path.display);
    }
    mutationResult = FileSystemItemMutationResult{sizeof(FileSystemItemMutationResult), TRUE, TRUE, FALSE};
    Debug::Perf::EmitCounter(L"FileOps.Local.DeleteByNameRemoved");
    Debug::Perf::Emit(L"FileOps.Local.DeleteByNameTree",
                      walk.concurrency > 1u ? L"parallel" : L"serial",
                      static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - deletionsStarted).count()),
                      walk.deleted.load(std::memory_order_acquire),
                      static_cast<uint64_t>(listed),
                      S_OK);
    return hr;
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

using RedSalamanderFileSystemTestRecycledItemVisitor = void(__stdcall*)(void* cookie, const wchar_t* originalPath, const wchar_t* recycledPath) noexcept;

// Hands every journaled Recycle Bin item (see JournalRecycledItemForTest) to the visitor, outside the
// journal lock, and empties the journal. droppedCount receives, and resets, the records the bounded
// journal had to drop; the caller treats any as a leak it cannot purge.
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderFileSystemTestTakeRecycledItems(RedSalamanderFileSystemTestRecycledItemVisitor visitor,
                                                                                                void* cookie,
                                                                                                uint64_t* droppedCount) noexcept
{
    if (visitor == nullptr || droppedCount == nullptr)
    {
        return E_POINTER;
    }

    std::vector<RecycledItemForTest> taken;
    {
        std::scoped_lock lock(g_recycledItemJournalForTestMutex);
        taken.swap(g_recycledItemJournalForTest);
        *droppedCount                       = g_recycledItemJournalForTestDropped;
        g_recycledItemJournalForTestDropped = 0u;
    }
    for (const RecycledItemForTest& record : taken)
    {
        visitor(cookie, record.originalPath.c_str(), record.recycledPath.c_str());
    }
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
                                     (HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR) ? FILESYSTEM_FLAG_CONTINUE_ON_ERROR : FILESYSTEM_FLAG_NONE) |
                                     (HasFlag(flags, FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY) ? FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY : FILESYSTEM_FLAG_NONE)),
        options,
        nullptr,
        nullptr,
        0u,
        FileSystemReparsePointPolicy::Preserve);
    DeleteDiscoveryState discovery{};
    context.deleteDiscovery         = &discovery;
    context.deleteConcurrencyBudget = 1u;
    ReportNativeFileWork(context, FileSystemFileWorkPhase::ScopeSupported);
    const PathInfo target = MakePathInfo(path);
    const HRESULT hr      = DeleteDirectoryRecursiveBatched(context, target, 1u, false);
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
    const auto check = [&](bool condition, const wchar_t* message) noexcept -> bool { return kFileSystemDebugCheck(condition, message, passed, failed); };

    // Recursive workers borrow one selected item's facts. Exercise the production join with
    // explicit storage outcomes, including a published root with uncertain child cleanup.
    {
        const PathInfo rootDestination{L"receipt-root", L"receipt-root"};
        const PathInfo childDestination{L"receipt-root/child", L"receipt-root/child"};
        OperationContext root{};
        OperationContext child{};
        child.destinationMutation = root.destinationMutation;
        FileSystemItemMutationResult receipt{};
        if (! check(SnapshotTrackedItemMutation(root, receipt) && receipt.outcomeKnown == TRUE && receipt.mutationCommitted == FALSE &&
                        receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::NotCreated,
                    L"an untouched recursive item must report known no-commit"))
            return;
        RecordOwnedStagePublication(child, childDestination, TrackedPublicationTruth::Published, FileSystemOwnedStageDisposition::Published);
        if (! check(SnapshotTrackedItemMutation(root, receipt) && receipt.outcomeKnown == TRUE && receipt.mutationCommitted == TRUE &&
                        receipt.originalStillPresent == TRUE && receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::NotApplicable,
                    L"a recursive worker publication must reach its selected root receipt"))
            return;
        RecordOwnedStageDisposition(child, childDestination, FileSystemOwnedStageDisposition::RetainedIncomplete);
        if (! check(SnapshotTrackedItemMutation(root, receipt) && receipt.outcomeKnown == TRUE && receipt.mutationCommitted == TRUE &&
                        receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::RetainedIncomplete,
                    L"a published sibling must not erase an incomplete retained child stage"))
            return;
        RecordOwnedStageDisposition(child, childDestination, FileSystemOwnedStageDisposition::Unknown);
        if (! check(SnapshotTrackedItemMutation(root, receipt) && receipt.outcomeKnown == TRUE && receipt.mutationCommitted == TRUE &&
                        receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::Unknown,
                    L"a published sibling must not turn uncertain child cleanup into known retention"))
            return;
        root.trackTopLevelOwnedStageMutation = true;
        root.itemDestination                 = rootDestination.display.c_str();
        RecordOwnedStagePublication(root, rootDestination, TrackedPublicationTruth::Published, FileSystemOwnedStageDisposition::Published);
        if (! check(SnapshotTrackedItemMutation(root, receipt) && receipt.outcomeKnown == TRUE && receipt.mutationCommitted == TRUE &&
                        receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::Unknown,
                    L"exact root publication must preserve uncertain descendant stage cleanup"))
            return;
        OperationContext otherRoot{};
        if (! check(SnapshotTrackedItemMutation(otherRoot, receipt) && receipt.outcomeKnown == TRUE && receipt.mutationCommitted == FALSE &&
                        receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::NotCreated,
                    L"separate selected roots must never share publication or cleanup facts"))
            return;
        OperationContext uncertainRoot{};
        RecordOwnedStagePublication(uncertainRoot, childDestination, TrackedPublicationTruth::Unknown, FileSystemOwnedStageDisposition::Unknown);
        if (! check(SnapshotTrackedItemMutation(uncertainRoot, receipt) && receipt.outcomeKnown == FALSE && receipt.mutationCommitted == FALSE &&
                        receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::Unknown,
                    L"uncertain publication without a proven commit must remain unknown"))
            return;
        RecordOwnedStagePublication(uncertainRoot, rootDestination, TrackedPublicationTruth::Published, FileSystemOwnedStageDisposition::Published);
        if (! check(SnapshotTrackedItemMutation(uncertainRoot, receipt) && receipt.outcomeKnown == TRUE && receipt.mutationCommitted == TRUE &&
                        receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::Unknown,
                    L"a proven partial commit must survive another uncertain publication"))
            return;
    }

    // Keep Both after a failed tracked attempt: the first stage at the original name was removed,
    // then the copy published the collision-safe sibling. The receipt describes that sibling.
    {
        const PathInfo originalDestination{L"keep-both", L"keep-both"};
        const PathInfo siblingDestination{L"keep-both (2)", L"keep-both (2)"};
        OperationContext copy{};
        copy.trackTopLevelOwnedStageMutation = true;
        copy.itemDestination                 = originalDestination.display.c_str();
        RecordOwnedStagePublication(copy, originalDestination, TrackedPublicationTruth::NotPublished, FileSystemOwnedStageDisposition::Retained);
        RecordOwnedStageDisposition(copy, originalDestination, FileSystemOwnedStageDisposition::Removed);
        FileSystemItemMutationResult receipt{};
        if (! check(SnapshotTrackedItemMutation(copy, receipt) && receipt.outcomeKnown == TRUE && receipt.mutationCommitted == FALSE &&
                        receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::Removed,
                    L"a removed stage at the original name must be a known retryable no-commit"))
            return;
        RetargetTrackedCopyDestination(copy, siblingDestination);
        RecordOwnedStagePublication(copy, siblingDestination, TrackedPublicationTruth::NotPublished, FileSystemOwnedStageDisposition::Retained);
        RecordOwnedStagePublication(copy, siblingDestination, TrackedPublicationTruth::Published, FileSystemOwnedStageDisposition::Published);
        if (! check(SnapshotTrackedItemMutation(copy, receipt) && receipt.outcomeKnown == TRUE && receipt.mutationCommitted == TRUE &&
                        receipt.originalStillPresent == TRUE && receipt.ownedStageDisposition == FileSystemOwnedStageDisposition::Published,
                    L"a Keep Both copy must report the sibling it published, not the original name's removed stage"))
            return;
    }

    {
        OperationContext move{};
        move.type = FILESYSTEM_MOVE;
        move.trackedNativeMove.store(TrackedPublicationTruth::NotPublished, std::memory_order_release);
        const FileSystemItemMutationResult missingSource = SnapshotNativeMoveMutation(move, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
        if (! check(missingSource.outcomeKnown == TRUE && missingSource.mutationCommitted == FALSE && missingSource.originalStillPresent == TRUE,
                    L"a source missing before any native rename is a known no-commit, not an unknown outcome"))
            return;
        MarkNativeMoveSourceUnproven(move);
        const FileSystemItemMutationResult unprovedSource = SnapshotNativeMoveMutation(move, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
        if (! check(unprovedSource.outcomeKnown == FALSE && unprovedSource.mutationCommitted == FALSE && unprovedSource.originalStillPresent == FALSE,
                    L"conditional source uncertainty cannot prove retention or authorize a native Move retry"))
            return;
        move.trackedNativeMove.store(TrackedPublicationTruth::Published, std::memory_order_release);
        MarkNativeMoveSourceUnproven(move);
        const FileSystemItemMutationResult committed = SnapshotNativeMoveMutation(move, HRESULT_FROM_WIN32(ERROR_CANCELLED));
        if (! check(committed.outcomeKnown == TRUE && committed.mutationCommitted == TRUE && committed.originalStillPresent == FALSE,
                    L"later source uncertainty must never erase an already proven native Move commit"))
            return;
    }

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

    // Every link below must live in, and point into, one marked sandbox; TEMP is there only under the runner.
    std::error_code sandboxEc;
    if (! check(Common::Testing::RequireSandboxedExposureTarget(root, sandboxEc),
                L"object-binding fixture root must be inside a marked X:\\RedSalamander.Perf (Tools\\Run-AllTests.ps1 places TEMP there)") ||
        ! check(CreateDirectoryW(root.c_str(), nullptr) != FALSE, L"fixture root should be created"))
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
    if (! check(Common::Testing::CreateSandboxedHardLink(alias, original, sandboxEc), L"hard-link alias should be created"))
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
    if (! check(Common::Testing::RequireSandboxedLinkTarget(junction, target, sandboxEc) &&
                    SUCCEEDED(BuildMountPointReparseData(target.wstring(), junctionData)) && SUCCEEDED(WriteReparsePointData(junction.wstring(), junctionData)),
                L"junction fixture should be materialized"))
    {
        return;
    }

    HRESULT lastSymlinkFixtureHr = S_OK;
    const auto createSymbolicLink =
        [&lastSymlinkFixtureHr](const std::filesystem::path& linkPath, std::wstring_view targetText, bool directory, bool relative) noexcept -> bool
    {
        // A root-relative target ("\dir\file") resolves on the link's own volume.
        const bool rootRelative = targetText.starts_with(L'\\') && ! targetText.starts_with(L"\\\\");
        std::error_code guardEc;
        if (! Common::Testing::RequireSandboxedLinkTarget(
                linkPath, rootRelative ? linkPath.root_name() / std::filesystem::path(targetText) : std::filesystem::path(targetText), guardEc))
        {
            lastSymlinkFixtureHr = HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
            return false;
        }
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
        HRESULT hr = BuildSymlinkReparseData(std::wstring(targetText), relative, data);
        if (SUCCEEDED(hr))
        {
            hr = WriteReparsePointData(linkPath.wstring(), data);
        }
        if (FAILED(hr))
        {
            lastSymlinkFixtureHr = hr;
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
        // The fixtures need SeCreateSymbolicLinkPrivilege, which an unelevated session holds only
        // with Developer Mode on. The failed check above names what did not happen; this names
        // why, so a run on such a machine is not left to rediscover it. The contract is still
        // reported as failed: a machine that cannot create the fixtures has not proved it, and
        // this export has no channel for a declared skip.
        if (lastSymlinkFixtureHr == HRESULT_FROM_WIN32(ERROR_PRIVILEGE_NOT_HELD))
        {
            std::fwprintf(stderr,
                          L"[FileSystem] object-binding selftest is blocked: SeCreateSymbolicLinkPrivilege is not held. "
                          L"Run from an elevated shell or enable Developer Mode; the symlink fixture checks cannot pass without it.\n");
            std::fflush(stderr);
        }
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

// A single-item Recycle the Shell holds runs on when another item's failure stops the call, but a
// user Cancel issued after that stop still stops the Shell (a large folder's walk reports through
// UpdateProgress), although the progress calls answer the stop without asking the caller.
void FileSystemInternal::RunDebugRecycleHeldCancellationSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    const auto check = [&](bool condition, const wchar_t* message) noexcept -> bool { return kFileSystemDebugCheck(condition, message, passed, failed); };

    class AbortControl final : public IFileSystemOperationControl
    {
    public:
        HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void*) noexcept override
        {
            if (abort == nullptr)
            {
                return E_POINTER;
            }
            *abort = abortRequested ? TRUE : FALSE;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void*) noexcept override
        {
            if (mode == nullptr)
            {
                return E_POINTER;
            }
            *mode = FILESYSTEM_DISCOVERY_AHEAD;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress*, void*) noexcept override
        {
            return S_OK;
        }

        bool abortRequested = false;
    };

    AbortControl control;
    FileSystemOptions options{};
    options.sizeBytes        = sizeof(options);
    options.operationControl = &control;
    ParallelOperationState parallel;
    OperationContext context;
    InitializeOperationContext(
        context, FILESYSTEM_DELETE, FILESYSTEM_FLAG_USE_RECYCLE_BIN, &options, nullptr, nullptr, 1u, FileSystemReparsePointPolicy::Preserve);
    context.parallel = &parallel;

    wil::com_ptr<IFileOperationProgressSink> sink;
    sink.attach(new (std::nothrow) RecycleBinDeleteProgressSink(context));
    if (! check(sink != nullptr, L"the held-Recycle cancellation self-test must create its progress sink"))
    {
        return;
    }
    check(sink->UpdateProgress(0u, 0u) == S_OK, L"a live Recycle call must let the Shell go on");

    parallel.stopOnErrorRequested.store(true, std::memory_order_release);
    parallel.lastCancelCheckTick.store(0u, std::memory_order_release);
    check(sink->UpdateProgress(0u, 0u) == S_OK && ! parallel.cancelRequested.load(std::memory_order_acquire),
          L"another item's stop alone must not stop a Recycle the Shell already holds");

    control.abortRequested = true;
    parallel.lastCancelCheckTick.store(0u, std::memory_order_release);
    check(sink->UpdateProgress(0u, 0u) == HRESULT_FROM_WIN32(ERROR_CANCELLED) && parallel.cancelRequested.load(std::memory_order_acquire),
          L"a user Cancel issued after another item's stop must still stop a Recycle the Shell holds, as the caller's cancellation");
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
    FileSystemConcurrencyMode concurrencyMode       = FileSystemConcurrencyMode::Manual;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
        concurrencyMode        = _concurrencyMode;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_COPY, flags, options, callback, cookie, 1, reparsePointPolicy);
    context.objectBinding = static_cast<IFileSystemObjectBinding*>(this);
    ReportNativeFileWork(context, FileSystemFileWorkPhase::ScopeSupported);

    const PathInfo source      = MakePathInfo(sourcePath);
    const PathInfo destination = MakePathInfo(destinationPath);
    BeginCopyItemMutation(context, source);

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

    const unsigned int maxConcurrency = ResolveCopyMoveConcurrencyLimit(copyMoveMaxConcurrency, options, kMaxCopyMoveMaxConcurrency, concurrencyMode);
    itemHr                            = ReportTopLevelDiscovery(context, source.extended, context.recursive);
    if (SUCCEEDED(itemHr))
    {
        PathInfo selectedDestination = destination;
        NativeFileWorkScope fileWorkScope(context, source, selectedDestination);
        for (;;)
        {
            context.reparseRootDestinationPath = TrimTrailingSeparatorsPreserveRoot(StripWin32ExtendedPrefix(MakeAbsolutePath(selectedDestination.display)));
            bytesCopied                        = 0u;
            itemHr =
                CopyPathInternalWithDirectoryParallelism(context, source, selectedDestination, flags, context.reparsePointPolicy, maxConcurrency, &bytesCopied);
            fileWorkScope.SetResult(itemHr);
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
                fileWorkScope.SetResult(issueHr);
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
                        fileWorkScope.SetResult(itemHr);
                        break;
                    }
                    const HRESULT keepBothHr = SelectKeepBothPath(destination, IsDirectory(sourceAttributes), selectedDestination);
                    if (FAILED(keepBothHr))
                    {
                        fileWorkScope.SetResult(keepBothHr);
                        itemHr = keepBothHr;
                        break;
                    }
                    RetargetTrackedCopyDestination(context, selectedDestination);
                    continue;
                }
                case FileSystemIssueAction::Retry: continue;
                case FileSystemIssueAction::Skip:
                    fileWorkScope.SetSkipped();
                    itemHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    break;
                case FileSystemIssueAction::PermanentDelete:
                case FileSystemIssueAction::Cancel:
                case FileSystemIssueAction::None:
                default: itemHr = HRESULT_FROM_WIN32(ERROR_CANCELLED); break;
            }
            fileWorkScope.SetResult(itemHr);
            break;
        }
        // selectedDestination ends with this scope; the tracked fields already hold its truth.
        context.trackedItemDestination = nullptr;
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
                        // The source vanished at the prompt; no rename was issued for it.
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
    const FileSystemItemMutationResult mutationResult = SnapshotNativeMoveMutation(context, itemHr);
    hr                                                = ReportItemCompleted(context, 0, itemHr, &mutationResult);
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
    ReportNativeFileWork(context, FileSystemFileWorkPhase::ScopeSupported);
    DeleteDiscoveryState discovery{};
    context.deleteDiscovery                     = &discovery;
    const auto closeDeleteDiscovery             = wil::scope_exit([&]() noexcept { ReportDeleteDiscoveryClosed(context, discovery); });
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

// I35 R7.3 (D-7.3): a permanent delete by name of an object this location opens but cannot identify.
// It behaves as DeleteItem (discovery, progress and one per-item receipt through the callback) except
// that the selected root is deleted only while the corroboration facts the host pinned still match on
// the handle that deletes it; any difference is ERROR_REVISION_MISMATCH and nothing is deleted.
HRESULT STDMETHODCALLTYPE FileSystem::DeleteIfIdentity(const wchar_t* path,
                                                       const FileSystemDeleteIdentity* identity,
                                                       FileSystemFlags flags,
                                                       const FileSystemOptions* options,
                                                       IFileSystemCallback* callback,
                                                       void* cookie) noexcept
{
    if (path == nullptr || identity == nullptr)
    {
        return E_POINTER;
    }
    if (path[0] == L'\0' || identity->sizeBytes != sizeof(FileSystemDeleteIdentity) || identity->identity[0] == L'\0' ||
        HasFlag(flags, FILESYSTEM_FLAG_USE_RECYCLE_BIN))
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
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy   = _reparsePointPolicy;
        deleteMaxConcurrency = _deleteMaxConcurrency;
    }

    OperationContext context{};
    InitializeOperationContext(context, FILESYSTEM_DELETE, flags, options, callback, cookie, 0, reparsePointPolicy);
    ReportNativeFileWork(context, FileSystemFileWorkPhase::ScopeSupported);
    DeleteDiscoveryState discovery{};
    context.deleteDiscovery         = &discovery;
    const auto closeDeleteDiscovery = wil::scope_exit([&]() noexcept { ReportDeleteDiscoveryClosed(context, discovery); });
    context.deleteConcurrencyBudget = std::clamp(deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency);

    const PathInfo target = MakePathInfo(path);
    HRESULT hr            = SetItemPaths(context, target.display.c_str(), nullptr);
    if (FAILED(hr))
    {
        return hr;
    }

    FileSystemItemMutationResult mutationResult{sizeof(FileSystemItemMutationResult), TRUE, FALSE, TRUE};
    const HRESULT itemHr = DeleteRootByNameInternal(context, target, path, *identity, mutationResult);
    hr                   = ReportItemCompleted(context, 0, itemHr, &mutationResult);
    if (FAILED(hr))
    {
        return hr;
    }
    hr = ReportProgressForced(context, 0, 0);
    if (FAILED(hr))
    {
        return hr;
    }
    Debug::Perf::Emit(L"FileOps.DeleteTraversal.MaxDepth", L"by-name", 0u, context.deleteTraversalMaxDepth, 0u, itemHr);
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
    FileSystemConcurrencyMode concurrencyMode       = FileSystemConcurrencyMode::Manual;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
        concurrencyMode        = _concurrencyMode;
    }

    const PathInfo destinationRoot    = MakePathInfo(destinationFolder);
    const unsigned int maxConcurrency = ResolveCopyMoveConcurrencyLimit(copyMoveMaxConcurrency, options, kMaxCopyMoveMaxConcurrency, concurrencyMode);
    const unsigned int concurrency    = std::min<unsigned int>(maxConcurrency, count);
    NativeFileWorkCallState fileWorkCall;

    if (concurrency <= 1u)
    {
        OperationContext context{};
        InitializeOperationContext(context, FILESYSTEM_COPY, flags, options, callback, cookie, count, reparsePointPolicy);
        context.objectBinding = static_cast<IFileSystemObjectBinding*>(this);
        context.fileWorkCall  = &fileWorkCall;
        ReportNativeFileWork(context, FileSystemFileWorkPhase::ScopeSupported);

        CallDiscoveryAggregator callDiscovery{};
        RootDiscoveryTranslation rootDiscovery{};
        rootDiscovery.aggregator      = &callDiscovery;
        context.discoveryTranslation  = &rootDiscovery;
        const auto closeCallDiscovery = wil::scope_exit([&]() noexcept
        {
            context.discoveryTranslation = nullptr;
            ReportCallDiscoveryClosed(context, callDiscovery);
        });

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

            // Each root restarts its own cumulative stream at zero; the call keeps the running sum.
            rootDiscovery.BeginRoot();
            BeginCopyItemMutation(context, source);
            NativeFileWorkScope fileWorkScope(context, source, destination);
            uint64_t bytesCopied = 0;
            HRESULT itemHr       = ReportTopLevelDiscovery(context, source.extended, context.recursive);
            if (SUCCEEDED(itemHr))
            {
                itemHr =
                    CopyPathInternalWithDirectoryParallelism(context, source, destination, flags, context.reparsePointPolicy, maxConcurrency, &bytesCopied);
            }
            fileWorkScope.SetResult(itemHr);

            FileSystemItemMutationResult mutationResult{};
            static_cast<void>(SnapshotTrackedItemMutation(context, mutationResult));
            fileWorkScope.Finish();
            hr = ReportItemCompleted(context, index, itemHr, &mutationResult);
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

    // Workers run roots concurrently, so each worker context owns its own translation record while
    // the call owns the one aggregate and the one closure. The closing context carries no
    // translation, so it emits the call's totals directly.
    CallDiscoveryAggregator callDiscovery{};
    OperationContext callDiscoveryContext{};
    InitializeOperationContext(callDiscoveryContext, FILESYSTEM_COPY, flags, &sharedOptionsState, callback, cookie, count, reparsePointPolicy);
    callDiscoveryContext.options      = &sharedOptionsState;
    callDiscoveryContext.parallel     = &parallel;
    callDiscoveryContext.fileWorkCall = &fileWorkCall;
    ReportNativeFileWork(callDiscoveryContext, FileSystemFileWorkPhase::ScopeSupported);
    const auto closeCallDiscovery = wil::scope_exit([&]() noexcept { ReportCallDiscoveryClosed(callDiscoveryContext, callDiscovery); });

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
        context.fileWorkCall     = &fileWorkCall;
        context.totalBytes       = 0; // discovered totals are reported by this traversal
        context.progressStreamId = schedulerStreamId;
        RootDiscoveryTranslation rootDiscovery{};
        rootDiscovery.aggregator     = &callDiscovery;
        context.discoveryTranslation = &rootDiscovery;

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

        BeginCopyItemMutation(context, source);
        NativeFileWorkScope fileWorkScope(context, source, destination);
        uint64_t bytesCopied = 0;
        HRESULT itemHr       = ReportTopLevelDiscovery(context, source.extended, context.recursive);
        if (SUCCEEDED(itemHr))
        {
            itemHr = CopyPathInternalWithDirectoryParallelism(context, source, destination, flags, context.reparsePointPolicy, nestedConcurrency, &bytesCopied);
        }
        fileWorkScope.SetResult(itemHr);

        FileSystemItemMutationResult mutationResult{};
        static_cast<void>(SnapshotTrackedItemMutation(context, mutationResult));
        fileWorkScope.Finish();
        hr = ReportItemCompleted(context, itemIndex, itemHr, &mutationResult);
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

    // R6.7: a recorded failure stops the other roots' walks, which then see cancellations; the failure
    // stays the call's result.
    if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
    {
        const HRESULT hr = parallel.firstError.load(std::memory_order_acquire);
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
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
    FileSystemConcurrencyMode concurrencyMode       = FileSystemConcurrencyMode::Manual;
    {
        std::lock_guard lock(_stateMutex);
        reparsePointPolicy     = _reparsePointPolicy;
        copyMoveMaxConcurrency = _copyMoveMaxConcurrency;
        concurrencyMode        = _concurrencyMode;
    }

    const PathInfo destinationRoot    = MakePathInfo(destinationFolder);
    const unsigned int maxConcurrency = ResolveCopyMoveConcurrencyLimit(copyMoveMaxConcurrency, options, kMaxCopyMoveMaxConcurrency, concurrencyMode);
    const unsigned int concurrency    = std::min<unsigned int>(maxConcurrency, count);

    if (concurrency <= 1u)
    {
        OperationContext context{};
        InitializeOperationContext(context, FILESYSTEM_MOVE, flags, options, callback, cookie, count, reparsePointPolicy);
        context.objectBinding = static_cast<IFileSystemObjectBinding*>(this);

        CallDiscoveryAggregator callDiscovery{};
        RootDiscoveryTranslation rootDiscovery{};
        rootDiscovery.aggregator      = &callDiscovery;
        context.discoveryTranslation  = &rootDiscovery;
        const auto closeCallDiscovery = wil::scope_exit([&]() noexcept
        {
            context.discoveryTranslation = nullptr;
            ReportCallDiscoveryClosed(context, callDiscovery);
        });

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

            // Each root restarts its own cumulative stream at zero; the call keeps the running sum
            // and closes once, so an early root can no longer declare the call's totals final.
            rootDiscovery.BeginRoot();
            context.trackedNativeMove.store(TrackedPublicationTruth::NotObserved, std::memory_order_relaxed);
            HRESULT itemHr = ReportTopLevelDiscovery(context, source.extended, false);
            if (SUCCEEDED(itemHr))
            {
                itemHr = MovePathInternal(context, source, destination);
            }
            const FileSystemItemMutationResult mutationResult = SnapshotNativeMoveMutation(context, itemHr);
            hr                                                = ReportItemCompleted(context, index, itemHr, &mutationResult);
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

    // Workers run roots concurrently, so each worker context owns its own translation record while
    // the call owns the one aggregate and the one closure. The closing context carries no
    // translation, so it emits the call's totals directly.
    CallDiscoveryAggregator callDiscovery{};
    OperationContext callDiscoveryContext{};
    InitializeOperationContext(callDiscoveryContext, FILESYSTEM_MOVE, flags, &sharedOptionsState, callback, cookie, count, reparsePointPolicy);
    callDiscoveryContext.options  = &sharedOptionsState;
    const auto closeCallDiscovery = wil::scope_exit([&]() noexcept { ReportCallDiscoveryClosed(callDiscoveryContext, callDiscovery); });

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
        RootDiscoveryTranslation rootDiscovery{};
        rootDiscovery.aggregator     = &callDiscovery;
        context.discoveryTranslation = &rootDiscovery;

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

        const FileSystemItemMutationResult mutationResult = SnapshotNativeMoveMutation(context, itemHr);
        hr                                                = ReportItemCompleted(context, itemIndex, itemHr, &mutationResult);
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

    // R6.7: a recorded failure stops the other roots' walks, which then see cancellations; the failure
    // stays the call's result.
    if (parallel.stopOnErrorRequested.load(std::memory_order_acquire))
    {
        const HRESULT hr = parallel.firstError.load(std::memory_order_acquire);
        return FAILED(hr) ? hr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (parallel.cancelRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
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
    // Both branches below share this one call-level discovery state, so the call owns the single
    // closure. The closing context exists only to carry the caller's operation control.
    OperationContext callDiscoveryContext{};
    InitializeOperationContext(callDiscoveryContext, FILESYSTEM_DELETE, flags, options, callback, cookie, count, reparsePointPolicy);
    ReportNativeFileWork(callDiscoveryContext, FileSystemFileWorkPhase::ScopeSupported);
    const auto closeDeleteDiscovery = wil::scope_exit([&]() noexcept { ReportDeleteDiscoveryClosed(callDiscoveryContext, discovery); });

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

            // Set when the last finalizeItem or report could not deliver its item (the callback refused
            // it). Any other false return comes after delivery: the call is stopping, but the batch
            // entries already observed are facts and are still reported.
            bool reportRefused = false;
            const FileSystemItemMutationResult unknownOutcome{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
            auto finalizeItem = [&](unsigned long completedIndex,
                                    const PathInfo& completedTarget,
                                    HRESULT itemHr,
                                    const FileSystemItemMutationResult* mutationResult) noexcept -> bool
            {
                reportRefused = false;
                // A Recycle canceled because another item's failure or a fatal error stopped this call,
                // not by the caller (see ReduceStopTimeRecycleCancellation).
                if (useRecycleBin && IsCancellationHr(itemHr) && parallel.stopOnErrorRequested.load(std::memory_order_acquire) &&
                    ! parallel.cancelRequested.load(std::memory_order_acquire) && ! ReduceStopTimeRecycleCancellation(itemHr, mutationResult))
                {
                    return false;
                }

                HRESULT hr = SetItemPaths(context, completedTarget.display.c_str(), nullptr);
                if (FAILED(hr))
                {
                    reportRefused = true;
                    signalFatal(hr);
                    return false;
                }

                bool delivered = false;
                hr             = ReportItemCompleted(context, completedIndex, itemHr, mutationResult, &delivered);
                if (FAILED(hr))
                {
                    reportRefused = ! delivered;
                    if (delivered)
                    {
                        parallel.completedItems.fetch_add(1u, std::memory_order_acq_rel);
                    }
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
                    // Transiently empty (another worker took the last ready item). The job parks; the
                    // in-flight worker's step promotes dependents or finishes the last item, and a step
                    // that ran a unit wakes the job.
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
                bool batchRefused = false;
                HRESULT batchHr   = DeleteToRecycleBinBatched(context, batchEntries, &batchStarted, &batchRefused);
                if (FAILED(batchHr) && ! batchStarted)
                {
                    Debug::Perf::EmitCounter(L"FileOps.RecycleBin.BatchFallbackCount");
#if defined(ENABLE_TESTS)
                    g_recycleBinBatchTestFallbacks.fetch_add(1u, std::memory_order_relaxed);
#endif
                    for (const RecycleBinBatchEntry& entry : batchEntries)
                    {
                        // Once another item stopped the call or it was canceled, an entry not yet
                        // attempted is left unreported, as never reached, rather than reported canceled.
                        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                        {
                            return DynamicStep::Finished;
                        }
                        FileSystemItemMutationResult mutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
                        const HRESULT itemHr = DeletePathInternal(context, entry.path, false, &mutationResult);
                        if (! finalizeItem(entry.itemIndex, entry.path, itemHr, &mutationResult))
                        {
                            return DynamicStep::Finished;
                        }
                    }

                    return finishOrContinue();
                }

                // IFileOperation decides recyclability for the whole batch, and a PreDeleteItem refusal
                // ends it before any item is touched. The refused entry and every entry the Shell never
                // offered are therefore untouched: each goes through the single-item route, which
                // recycles it or refuses it on its own. The sink reports the refusal even when the
                // refused item matched no entry. An entry offered without a terminal observation is
                // not replayed; it keeps the unknown-outcome path below.
                const bool refusalReplayable = batchRefused && std::ranges::none_of(batchEntries, [](const RecycleBinBatchEntry& entry) noexcept {
                    return ! entry.observed && entry.deleteOffered && ! entry.recycleRefused;
                });
                if (refusalReplayable)
                {
                    Debug::Perf::EmitCounter(L"FileOps.RecycleBin.BatchRefusalReplayCount");
                    // Every observed entry is reported even once the call stops; only a report the
                    // callback refuses ends the loop.
                    for (const RecycleBinBatchEntry& entry : batchEntries)
                    {
                        if (entry.observed && ! entry.recycleRefused)
                        {
                            if (! finalizeItem(entry.itemIndex, entry.path, entry.result, &entry.mutationResult) && reportRefused)
                            {
                                return DynamicStep::Finished;
                            }
                            continue;
                        }
                        if (parallel.cancelRequested.load(std::memory_order_acquire) || parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                        {
                            // Never reached: left unreported, while the observed entries are still reported.
                            continue;
                        }

                        FileSystemItemMutationResult mutationResult{sizeof(FileSystemItemMutationResult), FALSE, FALSE, TRUE};
                        const HRESULT itemHr = DeletePathInternal(context, entry.path, false, &mutationResult);
                        if (! finalizeItem(entry.itemIndex, entry.path, itemHr, &mutationResult) && reportRefused)
                        {
                            return DynamicStep::Finished;
                        }
                    }

                    return finishOrContinue();
                }

                // An entry the Shell offered without a terminal observation is reported first, with the
                // unknown outcome the single-item route reports, so an observed failure that stops this
                // call can never leave it looking unreached: a caller records it and never replays it.
                // Every such entry and every observed entry below is reported even once the call
                // stops; only a report the callback refuses ends these loops.
                for (const RecycleBinBatchEntry& entry : batchEntries)
                {
                    if (entry.observed || ! entry.deleteOffered || entry.recycleRefused)
                    {
                        continue;
                    }

                    HRESULT reportHr = SetItemPaths(context, entry.path.display.c_str(), nullptr);
                    if (FAILED(reportHr))
                    {
                        signalFatal(reportHr);
                        return DynamicStep::Finished;
                    }
                    bool delivered = false;
                    reportHr       = ReportItemCompleted(context, entry.itemIndex, HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE), &unknownOutcome, &delivered);
                    if (delivered)
                    {
                        parallel.completedItems.fetch_add(1u, std::memory_order_acq_rel);
                    }
                    if (FAILED(reportHr))
                    {
                        if (! parallel.stopOnErrorRequested.load(std::memory_order_acquire))
                        {
                            parallel.cancelRequested.store(true, std::memory_order_release);
                        }
                        scheduleCv.notify_all();
                        if (! delivered)
                        {
                            return DynamicStep::Finished;
                        }
                    }
                }

                // An observed failure stops this call before any entry is reported, so an entry the
                // Shell canceled after it is never mistaken for the caller's cancellation, whatever
                // the entry order.
                if (! context.continueOnError)
                {
                    const auto firstObservedFailure = std::ranges::find_if(batchEntries, [](const RecycleBinBatchEntry& entry) noexcept {
                        return entry.observed && FAILED(entry.result) && ! IsCancellationHr(entry.result);
                    });
                    if (firstObservedFailure != batchEntries.end())
                    {
                        parallel.hadFailure.store(true, std::memory_order_release);
                        signalFatal(firstObservedFailure->result);
                    }
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
                    if (! finalizeItem(entry.itemIndex, entry.path, entry.result, &entry.mutationResult) && reportRefused)
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
