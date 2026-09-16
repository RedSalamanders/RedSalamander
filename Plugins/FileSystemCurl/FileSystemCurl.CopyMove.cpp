#include "FileSystemCurl.Internal.h"
#include "FileOperationTraversalPolicy.h"

#include <array>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <deque>
#include <functional>
#include <memory>
#include <mutex>
#include <span>
#include <thread>
#include <unordered_map>
#include <unordered_set>

using namespace FileSystemCurlInternal;

namespace FileSystemCurlInternal
{
// Module anchor for AcquireModuleReferenceFromAddress — keeps the DLL loaded while worker threads are running.
const int kFileSystemCurlModuleAnchor = 0;

class SharedCopyMoveJobScheduler final
{
public:
    SharedCopyMoveJobScheduler() = default;
    ~SharedCopyMoveJobScheduler() noexcept
    {
        ShutdownAndJoin();
    }

    SharedCopyMoveJobScheduler(const SharedCopyMoveJobScheduler&)            = delete;
    SharedCopyMoveJobScheduler(SharedCopyMoveJobScheduler&&)                 = delete;
    SharedCopyMoveJobScheduler& operator=(const SharedCopyMoveJobScheduler&) = delete;
    SharedCopyMoveJobScheduler& operator=(SharedCopyMoveJobScheduler&&)      = delete;

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
        // The owning File Operations call's options, re-established on every worker that runs a
        // unit of this job (R0f-Curl cancel/deadline polling).
        const FileSystemOptions* operationOptions = nullptr;

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
        auto job              = std::make_shared<Job>();
        job->operationOptions = CurlCurrentOperationOptions();
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

    [[nodiscard]] bool EnsureWorkersAvailable() noexcept
    {
        ensureWorkers();

        std::scoped_lock lock(_initMutex);
        return ! _workers.empty();
    }

    void WaitJob(const JobPtr& job) noexcept
    {
        if (! job)
        {
            return;
        }

        std::unique_lock lock(job->doneMutex);
        job->doneCv.wait(lock, [&]() noexcept { return job->done.load(std::memory_order_acquire); });
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

private:
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

        constexpr unsigned int kMaxWorkers = 8u;
        workerCount                        = std::max(1u, std::min(workerCount, kMaxWorkers));

        _workers.reserve(workerCount);
        for (unsigned int i = 0; i < workerCount; ++i)
        {
            // Pin the module so the DLL cannot be unloaded while worker threads are running.
            wil::unique_hmodule modulePin = AcquireModuleReferenceFromAddress(&kFileSystemCurlModuleAnchor);
            if (! modulePin)
            {
                Debug::Error(L"FileSystemCurl: Failed to pin module for job scheduler worker thread {}", i);
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
            const CurlOperationOptionsScope operationOptionsScope(job->operationOptions);
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
        for (;;)
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

SharedCopyMoveJobScheduler& GetSharedCopyMoveJobScheduler() noexcept
{
    static SharedCopyMoveJobScheduler scheduler;
    return scheduler;
}

void ShutdownSharedCopyMoveJobScheduler() noexcept
{
    GetSharedCopyMoveJobScheduler().ShutdownAndJoin();
}
} // namespace FileSystemCurlInternal

namespace
{
[[nodiscard]] HRESULT ValidateConditionalMutationAdmission(FileSystemFlags flags, const FileSystemOptions* options) noexcept
{
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }

    // R0f-Curl: FTP RNTO and libcurl's SFTP/SCP rename cannot publish only when absent. A
    // no-overwrite mutation therefore probes its destination first (ERROR_FILE_EXISTS when it is
    // present) and then publishes; a creator that appears between the probe and the rename is
    // overwritten. That residual window is documented in the FTP/SFTP/SCP spec; refusing every
    // no-overwrite mutation made these destinations unusable for a file manager.
    static_cast<void>(flags);
    return S_OK;
}

class ConnectionConcurrencyLimiter final
{
public:
    ConnectionConcurrencyLimiter()  = default;
    ~ConnectionConcurrencyLimiter() = default;

    ConnectionConcurrencyLimiter(const ConnectionConcurrencyLimiter&)            = delete;
    ConnectionConcurrencyLimiter& operator=(const ConnectionConcurrencyLimiter&) = delete;
    ConnectionConcurrencyLimiter(ConnectionConcurrencyLimiter&&)                 = delete;
    ConnectionConcurrencyLimiter& operator=(ConnectionConcurrencyLimiter&&)      = delete;

    enum class Kind : uint8_t
    {
        CopyMove,
        Delete,
    };

    class Permit final
    {
    public:
        Permit() = default;

        Permit(ConnectionConcurrencyLimiter* limiter, std::wstring key, Kind kind) noexcept : _limiter(limiter), _key(std::move(key)), _kind(kind)
        {
        }

        Permit(const Permit&)            = delete;
        Permit& operator=(const Permit&) = delete;

        Permit(Permit&& other) noexcept : _limiter(std::exchange(other._limiter, nullptr)), _key(std::move(other._key)), _kind(other._kind)
        {
        }

        Permit& operator=(Permit&& other) noexcept
        {
            if (this == &other)
            {
                return *this;
            }

            Release();
            _limiter = std::exchange(other._limiter, nullptr);
            _key     = std::move(other._key);
            _kind    = other._kind;
            return *this;
        }

        ~Permit()
        {
            Release();
        }

        explicit operator bool() const noexcept
        {
            return _limiter != nullptr;
        }

    private:
        void Release() noexcept
        {
            if (! _limiter)
            {
                return;
            }

            _limiter->Release(_key, _kind);
            _limiter = nullptr;
        }

        ConnectionConcurrencyLimiter* _limiter = nullptr;
        std::wstring _key;
        Kind _kind = Kind::CopyMove;
    };

    template <typename CancelPredicate> [[nodiscard]] Permit AcquireCopyMove(std::wstring_view key, uint32_t max, CancelPredicate&& shouldCancel) noexcept
    {
        return Acquire(key, Kind::CopyMove, max, std::forward<CancelPredicate>(shouldCancel));
    }

    template <typename CancelPredicate> [[nodiscard]] Permit AcquireDelete(std::wstring_view key, uint32_t max, CancelPredicate&& shouldCancel) noexcept
    {
        return Acquire(key, Kind::Delete, max, std::forward<CancelPredicate>(shouldCancel));
    }

private:
    struct Entry final
    {
        uint32_t maxCopyMove      = 1;
        uint32_t inFlightCopyMove = 0;
        uint32_t maxDelete        = 1;
        uint32_t inFlightDelete   = 0;
    };

    void Release(const std::wstring& key, Kind kind) noexcept
    {
        std::lock_guard lock(_mutex);

        const auto it = _entries.find(key);
        if (it == _entries.end())
        {
            return;
        }

        Entry& entry = it->second;
        if (kind == Kind::CopyMove)
        {
            if (entry.inFlightCopyMove > 0)
            {
                --entry.inFlightCopyMove;
            }
        }
        else
        {
            if (entry.inFlightDelete > 0)
            {
                --entry.inFlightDelete;
            }
        }

        _cv.notify_all();
    }

    template <typename CancelPredicate>
    [[nodiscard]] Permit Acquire(std::wstring_view keyView, Kind kind, uint32_t max, CancelPredicate&& shouldCancel) noexcept
    {
        if (keyView.empty())
        {
            return {};
        }

        std::wstring key(keyView);
        const uint32_t maxEffective = (std::max)(1u, max);

        std::unique_lock lock(_mutex);
        for (;;)
        {
            lock.unlock();
            if (shouldCancel())
            {
                return {};
            }
            lock.lock();

            Entry& entry = _entries[key];
            if (kind == Kind::CopyMove)
            {
                entry.maxCopyMove = maxEffective;
                if (entry.inFlightCopyMove < entry.maxCopyMove)
                {
                    ++entry.inFlightCopyMove;
                    return Permit(this, std::move(key), kind);
                }
            }
            else
            {
                entry.maxDelete = maxEffective;
                if (entry.inFlightDelete < entry.maxDelete)
                {
                    ++entry.inFlightDelete;
                    return Permit(this, std::move(key), kind);
                }
            }

            _cv.wait_for(lock, std::chrono::milliseconds(100));
        }
    }

    std::mutex _mutex;
    std::condition_variable _cv;
    std::unordered_map<std::wstring, Entry> _entries;
};

ConnectionConcurrencyLimiter& GetConnectionConcurrencyLimiter() noexcept
{
    static ConnectionConcurrencyLimiter limiter;
    return limiter;
}

using SourceSizeCommitmentMap = SourceSizeCommitmentDigestMap;
using SourceTreePathSet       = SourceTreeDigestSet;

[[nodiscard]] HRESULT DeleteDirectoryTree(const ConnectionInfo& conn,
                                          std::wstring_view directoryRemotePath,
                                          std::wstring_view directoryFullPath,
                                          FileSystemFlags flags,
                                          ConnectionConcurrencyLimiter::Kind kind,
                                          unsigned int requestedConcurrency,
                                          FileOperationProgress& progress,
                                          std::atomic_bool* mutationAttempted       = nullptr,
                                          const SourceTreePathSet* allowedMembers = nullptr) noexcept;

} // namespace

namespace FileSystemCurlInternal
{

uint64_t DigestPluginPath(std::wstring_view normalizedPath) noexcept
{
    // FNV-1a 64 over the UTF-16 code units of the already normalized plugin path. One trailing
    // slash is dropped first (never the root "/"): the preflight walks a directory spelled with
    // it while the delete pass names the same directory without it, and both must digest alike.
    if (normalizedPath.size() > 1u && normalizedPath.back() == L'/')
    {
        normalizedPath.remove_suffix(1u);
    }
    uint64_t hash = 14695981039346656037ull;
    for (const wchar_t unit : normalizedPath)
    {
        hash ^= static_cast<uint64_t>(static_cast<uint16_t>(unit));
        hash *= 1099511628211ull;
    }
    return hash;
}

void SourceTreeDigestSet::Insert(std::wstring_view normalizedPath)
{
    _finalized = false;
    _digests.push_back(DigestPluginPath(normalizedPath));
}

void SourceTreeDigestSet::Finalize() noexcept
{
    std::ranges::sort(_digests);
    const auto duplicates = std::ranges::unique(_digests);
    _digests.erase(duplicates.begin(), duplicates.end());
    _finalized = true;
}

bool SourceTreeDigestSet::Contains(std::wstring_view normalizedPath) const noexcept
{
    if (! _finalized)
    {
        return false;
    }
    return std::ranges::binary_search(_digests, DigestPluginPath(normalizedPath));
}

void SourceSizeCommitmentDigestMap::Insert(std::wstring_view normalizedPath, uint64_t sizeBytes)
{
    _finalized = false;
    _entries.emplace_back(DigestPluginPath(normalizedPath), sizeBytes);
}

void SourceSizeCommitmentDigestMap::Finalize() noexcept
{
    // A later insert for the same path wins (insert_or_assign semantics): stable sort, keep the last.
    std::ranges::stable_sort(_entries, {}, &std::pair<uint64_t, uint64_t>::first);
    std::vector<std::pair<uint64_t, uint64_t>> compact;
    compact.reserve(_entries.size());
    for (const auto& entry : _entries)
    {
        if (! compact.empty() && compact.back().first == entry.first)
        {
            compact.back().second = entry.second;
            continue;
        }
        compact.push_back(entry);
    }
    _entries.swap(compact);
    _finalized = true;
}

std::optional<uint64_t> SourceSizeCommitmentDigestMap::Find(std::wstring_view normalizedPath) const noexcept
{
    if (! _finalized)
    {
        return std::nullopt;
    }
    const uint64_t digest = DigestPluginPath(normalizedPath);
    const auto found      = std::ranges::lower_bound(_entries, digest, {}, &std::pair<uint64_t, uint64_t>::first);
    if (found == _entries.end() || found->first != digest)
    {
        return std::nullopt;
    }
    return found->second;
}

} // namespace FileSystemCurlInternal

namespace
{

[[nodiscard]] HRESULT RemoteDeleteFileWithPermit(const ConnectionInfo& conn,
                                                 std::wstring_view remotePath,
                                                 FileOperationProgress& progress,
                                                 ConnectionConcurrencyLimiter::Kind kind,
                                                 std::atomic_bool* mutationAttempted = nullptr) noexcept
{
    if (conn.limiterKey.empty())
    {
        if (mutationAttempted)
        {
            mutationAttempted->store(true, std::memory_order_release);
        }
        return RemoteDeleteFile(conn, remotePath);
    }

    auto shouldCancel = [&]() noexcept { return FAILED(progress.CheckCancel()); };

    ConnectionConcurrencyLimiter& limiter = GetConnectionConcurrencyLimiter();
    const uint32_t max = kind == ConnectionConcurrencyLimiter::Kind::CopyMove ? conn.effectiveCopyMoveMaxConcurrency : conn.effectiveDeleteMaxConcurrency;

    ConnectionConcurrencyLimiter::Permit permit = kind == ConnectionConcurrencyLimiter::Kind::CopyMove
                                                      ? limiter.AcquireCopyMove(conn.limiterKey, max, shouldCancel)
                                                      : limiter.AcquireDelete(conn.limiterKey, max, shouldCancel);
    if (! permit)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (mutationAttempted)
    {
        mutationAttempted->store(true, std::memory_order_release);
    }
    return RemoteDeleteFile(conn, remotePath);
}

[[nodiscard]] HRESULT RemoteRemoveDirectoryWithPermit(const ConnectionInfo& conn,
                                                      std::wstring_view remotePath,
                                                      FileOperationProgress& progress,
                                                      ConnectionConcurrencyLimiter::Kind kind,
                                                      std::atomic_bool* mutationAttempted = nullptr) noexcept
{
    if (conn.limiterKey.empty())
    {
        if (mutationAttempted)
        {
            mutationAttempted->store(true, std::memory_order_release);
        }
        return RemoteRemoveDirectory(conn, remotePath);
    }

    auto shouldCancel = [&]() noexcept { return FAILED(progress.CheckCancel()); };

    ConnectionConcurrencyLimiter& limiter = GetConnectionConcurrencyLimiter();
    const uint32_t max = kind == ConnectionConcurrencyLimiter::Kind::CopyMove ? conn.effectiveCopyMoveMaxConcurrency : conn.effectiveDeleteMaxConcurrency;

    ConnectionConcurrencyLimiter::Permit permit = kind == ConnectionConcurrencyLimiter::Kind::CopyMove
                                                      ? limiter.AcquireCopyMove(conn.limiterKey, max, shouldCancel)
                                                      : limiter.AcquireDelete(conn.limiterKey, max, shouldCancel);
    if (! permit)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (mutationAttempted)
    {
        mutationAttempted->store(true, std::memory_order_release);
    }
    return RemoteRemoveDirectory(conn, remotePath);
}


[[nodiscard]] bool CanServerSideRename(const ConnectionInfo& sourceConn, const ConnectionInfo& destinationConn) noexcept
{
    if (sourceConn.protocol != destinationConn.protocol)
    {
        return false;
    }

    if (sourceConn.host != destinationConn.host)
    {
        return false;
    }

    if (sourceConn.port != destinationConn.port)
    {
        return false;
    }

    if (sourceConn.user != destinationConn.user)
    {
        return false;
    }

    if (sourceConn.password != destinationConn.password)
    {
        return false;
    }

    if (sourceConn.basePath != destinationConn.basePath)
    {
        return false;
    }

    if (sourceConn.sshPrivateKey != destinationConn.sshPrivateKey)
    {
        return false;
    }

    if (sourceConn.sshPublicKey != destinationConn.sshPublicKey)
    {
        return false;
    }

    if (sourceConn.sshKeyPassphrase != destinationConn.sshKeyPassphrase)
    {
        return false;
    }

    if (sourceConn.sshKnownHosts != destinationConn.sshKnownHosts)
    {
        return false;
    }

    return true;
}

[[nodiscard]] HRESULT BuildRemoteSiblingLeaf(std::wstring_view purposeTag, std::wstring& leafOut) noexcept
{
    std::wstring marker = L".redsalamander-";
    marker.append(purposeTag);
    marker.push_back(L'-');
    return Common::Paths::BuildUniqueSiblingName(
        std::wstring_view{}, std::wstring_view(marker), std::wstring_view{}, (std::numeric_limits<size_t>::max)(), leafOut);
}

[[nodiscard]] HRESULT GenerateRemoteSiblingPath(const ConnectionInfo& conn,
                                                std::wstring_view destinationPath,
                                                std::wstring_view purposeTag,
                                                std::wstring& siblingPathOut) noexcept
{
    siblingPathOut.clear();

    const std::wstring parent = ParentPath(destinationPath);

    for (unsigned int attempt = 0; attempt < 32u; ++attempt)
    {
        std::wstring leaf;
        HRESULT hr = BuildRemoteSiblingLeaf(purposeTag, leaf);
        if (FAILED(hr))
        {
            return hr;
        }

        const std::wstring candidate = JoinPluginPath(parent, leaf);

        FilesInformationCurl::Entry ignored{};
        const HRESULT existsHr = GetEntryInfo(conn, candidate, ignored);
        if (existsHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            siblingPathOut = candidate;
            return S_OK;
        }
        if (FAILED(existsHr))
        {
            return existsHr;
        }
    }

    return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
}

// FTP, SFTP, and SCP advertise ordinalCaseSensitive path identity. Only an exact text match is a
// self-rename for those providers; a case-only destination must reach the server rename operation.
[[nodiscard]] bool IsCaseSensitiveSelfRename(std::wstring_view sourcePath, std::wstring_view destinationPath) noexcept
{
    return sourcePath == destinationPath;
}

#if defined(ENABLE_TESTS)
constexpr Common::DebugSelfTest::Check DebugCheck{L"FileSystemCurl"};

[[nodiscard]] bool IsHexToken(std::wstring_view token) noexcept
{
    for (const wchar_t ch : token)
    {
        if (! ((ch >= L'0' && ch <= L'9') || (ch >= L'A' && ch <= L'F') || (ch >= L'a' && ch <= L'f')))
        {
            return false;
        }
    }
    return true;
}

void RunDebugRemoteSiblingLeafEntropySelfTest(unsigned int& passed, unsigned int& failed)
{
    constexpr std::wstring_view prefix = L".redsalamander-upload-";

    std::vector<std::wstring> leaves;
    leaves.reserve(8u);

    for (unsigned int i = 0; i < 8u; ++i)
    {
        std::wstring leaf;
        const HRESULT hr = BuildRemoteSiblingLeaf(L"upload", leaf);
        DebugCheck(SUCCEEDED(hr), L"remote sibling leaf generation should succeed", passed, failed);
        if (SUCCEEDED(hr))
        {
            leaves.push_back(std::move(leaf));
        }
    }

    const std::wstring processIdText = std::to_wstring(GetCurrentProcessId());
    for (const std::wstring& leaf : leaves)
    {
        DebugCheck(leaf.starts_with(prefix), L"remote sibling leaf should keep the staging prefix", passed, failed);
        DebugCheck(leaf.find(processIdText) == std::wstring::npos, L"remote sibling leaf should not contain the process id", passed, failed);

        if (leaf.starts_with(prefix) && leaf.size() >= prefix.size())
        {
            const std::wstring_view token(leaf.data() + prefix.size(), leaf.size() - prefix.size());
            DebugCheck(token.size() == 32u, L"remote sibling leaf should contain a 128-bit hex entropy token", passed, failed);
            DebugCheck(IsHexToken(token), L"remote sibling leaf entropy token should be hex only", passed, failed);
        }
    }

    for (size_t i = 0; i < leaves.size(); ++i)
    {
        for (size_t j = i + 1u; j < leaves.size(); ++j)
        {
            DebugCheck(leaves[i] != leaves[j], L"remote sibling leaves should be unique across immediate generations", passed, failed);
        }
    }
}

void RunDebugOverwriteCleanupContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    CurlPublicationResult committedCleanupDebt{};
    committedCleanupDebt.primaryCommitted = true;
    committedCleanupDebt.RecordCleanupDebt(CurlCleanupDebtKind::RetainedRollbackSibling, HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
    DebugCheck(committedCleanupDebt.OperationResult() == S_OK && committedCleanupDebt.primaryMutationHr == S_OK &&
                   committedCleanupDebt.cleanupHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) &&
                   committedCleanupDebt.sourceDeletionHr == S_OK && committedCleanupDebt.cleanupDebtCount == 1u,
               L"structured publication result should preserve committed success while exposing cleanup debt",
               passed,
               failed);

    CurlPublicationResult failedStagingCleanup{};
    failedStagingCleanup.RecordPrimaryFailure(HRESULT_FROM_WIN32(ERROR_WRITE_FAULT));
    failedStagingCleanup.RecordCleanupDebt(CurlCleanupDebtKind::RetainedStagingSibling, HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
    DebugCheck(failedStagingCleanup.OperationResult() == HRESULT_FROM_WIN32(ERROR_WRITE_FAULT) && ! failedStagingCleanup.primaryCommitted &&
                   failedStagingCleanup.cleanupDebtCount == 1u,
               L"pre-publication staging cleanup debt should retain the primary mutation failure",
               passed,
               failed);

    CurlPublicationResult failedSourceDeletion{};
    failedSourceDeletion.primaryCommitted = true;
    failedSourceDeletion.RecordSourceDeletionFailure(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
    DebugCheck(failedSourceDeletion.OperationResult() == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) &&
                   failedSourceDeletion.sourceDeletionHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
               L"post-publication source deletion failure should remain a partial operation",
               passed,
               failed);
}

void RunDebugCaseOnlyRenameContractSelfTest(unsigned int& passed, unsigned int& failed) noexcept
{
    DebugCheck(IsCaseSensitiveSelfRename(L"/folder/name.txt", L"/folder/name.txt"), L"an exact Curl path match should remain a self-rename", passed, failed);
    DebugCheck(! IsCaseSensitiveSelfRename(L"/folder/name.txt", L"/folder/Name.txt"),
               L"a Curl case-only path change must execute the provider rename operation",
               passed,
               failed);
}
#endif

[[nodiscard]] HRESULT PrepareOverwriteTargetForRename(const ConnectionInfo& conn,
                                                      std::wstring_view destinationPath,
                                                      bool allowOverwrite,
                                                      std::wstring& backupPathOut,
                                                      bool* sourceObservedWithoutMutation = nullptr) noexcept
{
    backupPathOut.clear();

    FilesInformationCurl::Entry existing{};
    const HRESULT existsHr = GetEntryInfo(conn, destinationPath, existing);
    if (FAILED(existsHr))
    {
        return existsHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) ? S_OK : existsHr;
    }

    if (! allowOverwrite)
    {
        // R0f-Curl no-overwrite publication: the destination exists, so the caller decides
        // (the host raises its overwrite conflict). A creator that appears after this probe and
        // before the rename is overwritten; FTP/SFTP expose no atomic no-replace primitive.
        return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }

    if ((existing.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }

    HRESULT hr = GenerateRemoteSiblingPath(conn, destinationPath, L"rollback", backupPathOut);
    if (FAILED(hr))
    {
        return hr;
    }

    if (sourceObservedWithoutMutation)
    {
        *sourceObservedWithoutMutation = false;
    }
    hr = RemoteRename(conn, destinationPath, backupPathOut);
    if (FAILED(hr))
    {
        backupPathOut.clear();
        return hr;
    }

    return S_OK;
}

[[nodiscard]] CurlPublicationResult PreserveOverwriteArtifactsAfterFailure(HRESULT primaryMutationHr, std::wstring_view backupPath) noexcept
{
    CurlPublicationResult result{};
    result.RecordPrimaryFailure(primaryMutationHr);
    if (backupPath.empty())
    {
        return result;
    }

    // Neither the final pathname nor the randomized backup name carries an
    // immutable remote object identity. Restoring/deleting by path could remove
    // a concurrent owner, so preserve both names and report partial state.
    result.RecordCleanupDebt(CurlCleanupDebtKind::RetainedRollbackSibling, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    return result;
}

[[nodiscard]] CurlPublicationResult FinalizeOverwriteTarget(const ConnectionInfo& conn, std::wstring_view backupPath) noexcept
{
    static_cast<void>(conn);
    CurlPublicationResult result{};
    result.primaryCommitted = true;
    if (backupPath.empty())
    {
        return result;
    }

    // The transport exposes no conditional delete by immutable identity. The
    // backup may have been replaced after publication; retain it as cleanup debt
    // instead of deleting a path that may now belong to another actor.
    result.RecordCleanupDebt(CurlCleanupDebtKind::RetainedRollbackSibling, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    return result;
}

[[nodiscard]] CurlPublicationResult RenameWithOverwriteRollback(const ConnectionInfo& conn,
                                                                std::wstring_view sourcePath,
                                                                std::wstring_view destinationPath,
                                                                bool allowOverwrite,
                                                                bool& sourceObservedWithoutMutation) noexcept
{
    std::wstring backupPath;
    HRESULT hr = PrepareOverwriteTargetForRename(conn, destinationPath, allowOverwrite, backupPath, &sourceObservedWithoutMutation);
    if (FAILED(hr))
    {
        CurlPublicationResult result{};
        result.RecordPrimaryFailure(hr);
        return result;
    }

    sourceObservedWithoutMutation = false;
    hr                            = RemoteRename(conn, sourcePath, destinationPath);
    if (FAILED(hr))
    {
        return PreserveOverwriteArtifactsAfterFailure(hr, backupPath);
    }

    return FinalizeOverwriteTarget(conn, backupPath);
}

[[nodiscard]] CurlPublicationResult PreserveMovedFileDestinationAfterSourceDeleteFailure(HRESULT sourceDeleteHr,
                                                                                          std::wstring_view backupPath) noexcept
{
    CurlPublicationResult result{};
    result.primaryCommitted = true;
    result.RecordSourceDeletionFailure(sourceDeleteHr);
    if (! backupPath.empty())
    {
        result.RecordCleanupDebt(CurlCleanupDebtKind::RetainedRollbackSibling, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    }
    return result;
}

[[nodiscard]] CurlPublicationResult PreserveMovedDirectoryAfterSourceDeleteFailure(HRESULT sourceDeleteHr) noexcept
{
    CurlPublicationResult result{};
    result.primaryCommitted = true;
    result.RecordSourceDeletionFailure(sourceDeleteHr);
    result.RecordCleanupDebt(CurlCleanupDebtKind::PreservedDestinationTree, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    return result;
}

[[nodiscard]] HRESULT QueryPathExists(const ConnectionInfo& conn, std::wstring_view path, bool& existsOut) noexcept
{
    existsOut = false;

    FilesInformationCurl::Entry existing{};
    const HRESULT hr = GetEntryInfo(conn, path, existing);
    if (SUCCEEDED(hr))
    {
        existsOut = true;
        return S_OK;
    }

    return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) ? S_OK : hr;
}

[[nodiscard]] CurlPublicationResult PreserveCopiedDirectoryAfterFailure(HRESULT primaryMutationHr) noexcept
{
    CurlPublicationResult result{};
    result.RecordPrimaryFailure(primaryMutationHr);
    result.RecordCleanupDebt(CurlCleanupDebtKind::PreservedDestinationTree, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    return result;
}

[[nodiscard]] CurlPublicationResult PromoteStagedFileToDestination(const ConnectionInfo& destinationConn,
                                                                    std::wstring_view stagedRemotePath,
                                                                    std::wstring_view destinationRemotePath,
                                                                    bool allowOverwrite,
                                                                    std::wstring* backupPathOut) noexcept
{
    std::wstring backupPath;
    HRESULT hr = PrepareOverwriteTargetForRename(destinationConn, destinationRemotePath, allowOverwrite, backupPath);
    if (FAILED(hr))
    {
        CurlPublicationResult result{};
        result.RecordPrimaryFailure(hr);
        const HRESULT cleanupHr = RemoteDeleteFile(destinationConn, stagedRemotePath);
        if (FAILED(cleanupHr))
        {
            result.RecordCleanupDebt(CurlCleanupDebtKind::RetainedStagingSibling, cleanupHr);
        }
        return result;
    }

    hr = RemoteRename(destinationConn, stagedRemotePath, destinationRemotePath);
    if (FAILED(hr))
    {
        CurlPublicationResult result = PreserveOverwriteArtifactsAfterFailure(hr, backupPath);
        const HRESULT cleanupHr       = RemoteDeleteFile(destinationConn, stagedRemotePath);
        if (FAILED(cleanupHr))
        {
            result.RecordCleanupDebt(CurlCleanupDebtKind::RetainedStagingSibling, cleanupHr);
        }
        return result;
    }

    if (backupPathOut)
    {
        *backupPathOut = std::move(backupPath);
        CurlPublicationResult result{};
        result.primaryCommitted = true;
        return result;
    }

    return FinalizeOverwriteTarget(destinationConn, backupPath);
}

[[nodiscard]] HRESULT VerifyStagedUploadExactSize(const ConnectionInfo& destinationConn,
                                                  std::wstring_view stagedRemotePath,
                                                  uint64_t expectedSizeBytes,
                                                  uint64_t& requestCount) noexcept
{
    uint64_t stagedProbeSize  = 0u;
    bool stagedProbeSizeKnown = false;
    ++requestCount;
    const HRESULT probeHr = CurlProbeRemoteFileSize(destinationConn, stagedRemotePath, stagedProbeSize, stagedProbeSizeKnown);
    if (SUCCEEDED(probeHr) && stagedProbeSizeKnown)
    {
        return stagedProbeSize == expectedSizeBytes ? S_OK : HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    FilesInformationCurl::Entry stagedInfo{};
    ++requestCount;
    const HRESULT infoHr = GetEntryInfo(destinationConn, stagedRemotePath, stagedInfo);
    if (FAILED(infoHr))
    {
        return infoHr;
    }
    if ((stagedInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 || ! stagedInfo.sizeKnown || stagedInfo.sizeBytes != expectedSizeBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }
    return S_OK;
}

[[nodiscard]] HRESULT PreflightDirectorySourceSizes(const ConnectionInfo& sourceConn,
                                                    std::wstring_view sourceRemoteDir,
                                                    FileOperationProgress& progress,
                                                    SourceSizeCommitmentMap& commitments,
                                                    SourceTreePathSet& members) noexcept
{
    using namespace Common::FileOperations;
    const auto started         = std::chrono::steady_clock::now();
    HRESULT result             = E_PENDING;
    uint64_t peakFrames        = 0u;
    uint64_t peakPathBytes     = 0u;
    uint64_t peakMetadataBytes = 0u;
    const auto reportMetrics   = wil::scope_exit([&]() noexcept
    {
        Debug::Perf::Emit(L"FileOps.Curl.MovePreflight.Traversal", L"read-only", Debug::Perf::ElapsedUs(started), peakFrames, commitments.size(), result);
        Debug::Perf::Emit(L"FileOps.Curl.MovePreflight.Retained", L"cursor-frames-only", 0u, peakPathBytes, peakMetadataBytes, result);
        // Whole-operation retention: the digests the copy and delete walkers keep for the Move.
        Debug::Perf::Emit(L"FileOps.Curl.MovePreflight.RetainedDigestBytes",
                          L"whole-operation",
                          0u,
                          commitments.RetainedBytes() + members.RetainedBytes(),
                          static_cast<uint64_t>(commitments.size() + members.size()),
                          result);
    });
    const auto finish          = [&](HRESULT hr) noexcept
    {
        result = NormalizeCancellation(hr);
        if (SUCCEEDED(result))
        {
            commitments.Finalize();
            members.Finalize();
        }
        return result;
    };

    // This pass must finish before Copy can mutate anything. Unlike Delete's
    // work batches, every discovery/probe failure is global; there is no partial
    // preflight success. Reuse the strict transport cursor and quantitative policy.
    // The caller's commitment map still retains whole-tree per-file proof and is
    // deliberately not included in the cursor-frame reservation metrics.
    struct Frame final
    {
        Frame()                        = default;
        Frame(const Frame&)            = delete;
        Frame& operator=(const Frame&) = delete;
        Frame(Frame&&)                 = delete;
        Frame& operator=(Frame&&)      = delete;
        std::wstring remotePath;
        CurlDirectoryCursor cursor;
        std::vector<std::wstring> pendingDirectories;
        uint64_t pathBytes = 0u;
    };
    std::vector<std::unique_ptr<Frame>> frames;
    frames.reserve(static_cast<size_t>(kTraversalMaxDepth + 1u));
    uint64_t retainedPathBytes     = 0u;
    const uint64_t frameSlotsBytes = static_cast<uint64_t>(frames.capacity()) * sizeof(std::unique_ptr<Frame>);
    const auto pushFrame           = [&](std::wstring path) noexcept -> HRESULT
    {
        const uint64_t metadataBytes = frameSlotsBytes + (frames.size() + 1u) * (sizeof(Frame) + CurlDirectoryCursor::kMetadataReservationBytes);
        if (frames.size() >= kTraversalMaxDepth + 1u || metadataBytes > kTraversalMaxMetadataBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        auto frame        = std::make_unique<Frame>();
        frame->remotePath = std::move(path);
        frame->pathBytes  = static_cast<uint64_t>(frame->remotePath.capacity() + 1u) * sizeof(wchar_t);
        if (frame->pathBytes > kTraversalMaxQueuedPathBytes - retainedPathBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        const HRESULT openHr =
            frame->cursor.Open(sourceConn, EnsureTrailingSlash(frame->remotePath), [&progress]() noexcept { return progress.CheckCancel(); });
        if (FAILED(openHr))
        {
            return openHr;
        }
        frame->pathBytes += frame->cursor.RetainedPathBytes();
        if (frame->pathBytes > kTraversalMaxQueuedPathBytes - retainedPathBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        retainedPathBytes += frame->pathBytes;
        frames.push_back(std::move(frame));
        peakFrames        = (std::max)(peakFrames, static_cast<uint64_t>(frames.size()));
        peakPathBytes     = (std::max)(peakPathBytes, retainedPathBytes);
        peakMetadataBytes = (std::max)(peakMetadataBytes, metadataBytes);
        return S_OK;
    };

    HRESULT hr = pushFrame(NormalizePluginPath(sourceRemoteDir));
    if (FAILED(hr))
    {
        return finish(hr);
    }
    members.Insert(frames.back()->remotePath);
    while (! frames.empty())
    {
        Frame& frame = *frames.back();
        FilesInformationCurl::Entry entry{};
        hr = frame.cursor.Next(entry);
        if (FAILED(hr))
        {
            return finish(hr);
        }
        if (hr == S_FALSE)
        {
            if (! frame.pendingDirectories.empty())
            {
                std::wstring child = std::move(frame.pendingDirectories.front());
                const uint64_t pendingBytes = static_cast<uint64_t>(child.capacity() + 1u) * sizeof(wchar_t);
                frame.pendingDirectories.erase(frame.pendingDirectories.begin());
                if (pendingBytes > frame.pathBytes || pendingBytes > retainedPathBytes)
                {
                    return finish(E_UNEXPECTED);
                }
                frame.pathBytes -= pendingBytes;
                retainedPathBytes -= pendingBytes;
                hr = pushFrame(std::move(child));
                if (FAILED(hr))
                {
                    return finish(hr);
                }
                continue;
            }
            retainedPathBytes -= frame.pathBytes;
            frames.pop_back();
            continue;
        }
        const std::wstring sourceChildRemote = JoinPluginPath(frame.remotePath, entry.name);
        members.Insert(sourceChildRemote);
        if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            const uint64_t pendingBytes = static_cast<uint64_t>(sourceChildRemote.capacity() + 1u) * sizeof(wchar_t);
            if (pendingBytes > kTraversalMaxQueuedPathBytes - retainedPathBytes)
            {
                return finish(HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW));
            }
            frame.pendingDirectories.push_back(sourceChildRemote);
            frame.pathBytes += pendingBytes;
            retainedPathBytes += pendingBytes;
            continue;
        }
        if ((entry.attributes & FILE_ATTRIBUTE_DEVICE) != 0)
        {
            return finish(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
        }

        CurlSourceSizeCommitment commitment{};
        hr = ResolveCurlSourceSizeCommitment(sourceConn, sourceChildRemote, entry.sizeBytes, entry.sizeKnown, commitment);
        if (FAILED(hr))
        {
            return finish(hr);
        }
        if (! commitment.known)
        {
            return finish(HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY));
        }

        commitments.Insert(sourceChildRemote, commitment.sizeBytes);
    }

    return finish(S_OK);
}

enum class CopyFilePhase : size_t
{
    SourceCommitment,
    LocalTemp,
    SourceDownload,
    LocalSourceValidation,
    ReserveStagingName,
    Upload,
    VerifyStaging,
    Promote,
    Completion,
    Count,
};

constexpr std::array<std::wstring_view, static_cast<size_t>(CopyFilePhase::Count)> kCopyFilePhaseNames{{
    L"source-commitment",
    L"local-temp",
    L"source-download",
    L"local-source-validation",
    L"reserve-staging-name",
    L"upload",
    L"verify-staging",
    L"promote",
    L"completion",
}};

// One bounded accumulator per native tree; scheduler workers add file-phase wall
// time and visits, then the owner emits once after drain. Parallel durations may
// overlap. These are not CPU timings or per-file/path-bearing telemetry.
struct CopyFilePhaseTimings final
{
    CopyFilePhaseTimings()                                       = default;
    CopyFilePhaseTimings(const CopyFilePhaseTimings&)            = delete;
    CopyFilePhaseTimings(CopyFilePhaseTimings&&)                 = delete;
    CopyFilePhaseTimings& operator=(const CopyFilePhaseTimings&) = delete;
    CopyFilePhaseTimings& operator=(CopyFilePhaseTimings&&)      = delete;

    std::array<std::atomic<uint64_t>, kCopyFilePhaseNames.size()> elapsedUs{};
    std::array<std::atomic<uint64_t>, kCopyFilePhaseNames.size()> visits{};
};

[[nodiscard]] HRESULT CopyFileViaTemp(const ConnectionInfo& sourceConn,
                                      std::wstring_view sourceRemotePath,
                                      std::wstring_view sourceFullPath,
                                      const ConnectionInfo& destinationConn,
                                      std::wstring_view destinationRemotePath,
                                      std::wstring_view destinationFullPath,
                                      FileSystemFlags flags,
                                      FileOperationProgress& progress,
                                      uint64_t expectedSizeBytes,
                                      bool expectedSizeKnown,
                                      bool requireAuthoritativeSourceSize,
                                      bool sourceSizeAlreadyResolved,
                                      std::atomic<uint64_t>* concurrentOverallBytes,
                                      std::wstring* backupPathOut                        = nullptr,
                                      CurlPublicationAccumulator* publicationAccumulator = nullptr,
                                      CopyFilePhaseTimings* phaseTimings                 = nullptr) noexcept
{
    CurlPublicationResult publicationResult{};
    CopyFilePhase phase    = CopyFilePhase::SourceCommitment;
    auto phaseStarted      = std::chrono::steady_clock::now();
    const auto recordPhase = [&]() noexcept
    {
        if (phaseTimings)
        {
            const size_t index = static_cast<size_t>(phase);
            phaseTimings->elapsedUs[index].fetch_add(Debug::Perf::ElapsedUs(phaseStarted), std::memory_order_relaxed);
            phaseTimings->visits[index].fetch_add(1u, std::memory_order_relaxed);
        }
    };
    const auto enterPhase = [&](CopyFilePhase next) noexcept
    {
        recordPhase();
        phase        = next;
        phaseStarted = std::chrono::steady_clock::now();
    };
    auto mergePublicationResult      = wil::scope_exit([&]() noexcept
    {
        recordPhase();
        if (FAILED(publicationResult.OperationResult()))
        {
            Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.FileFailure",
                              kCopyFilePhaseNames[static_cast<size_t>(phase)].data(),
                              0u,
                              0u,
                              0u,
                              publicationResult.OperationResult());
        }
        if (publicationAccumulator)
        {
            publicationAccumulator->Merge(publicationResult);
        }
    });
    const auto failBeforePublication = [&](HRESULT failureHr) noexcept
    {
        publicationResult.RecordPrimaryFailure(failureHr);
        return failureHr;
    };

    CurlSourceSizeCommitment sourceSize{.sizeBytes = expectedSizeBytes, .known = expectedSizeKnown};
    HRESULT hr = S_OK;
    if (! sourceSizeAlreadyResolved)
    {
        hr = ResolveCurlSourceSizeCommitment(sourceConn, sourceRemotePath, expectedSizeBytes, expectedSizeKnown, sourceSize);
        if (FAILED(hr))
        {
            return failBeforePublication(hr);
        }
    }

    if (requireAuthoritativeSourceSize && ! sourceSize.known)
    {
        return failBeforePublication(HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY));
    }

    hr = progress.ReportProgress(sourceSize.known ? sourceSize.sizeBytes : 0u, 0, sourceFullPath, destinationFullPath);
    if (FAILED(hr))
    {
        return failBeforePublication(hr);
    }

    const bool allowOverwrite = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);

    hr = progress.CheckCancel();
    if (FAILED(hr))
    {
        return failBeforePublication(hr);
    }

    enterPhase(CopyFilePhase::LocalTemp);
    wil::unique_hfile tempFile;
    if (const HRESULT tempHr = Common::Files::CreateDeleteOnCloseTemporaryFile(kCurlTemporaryFileOptions, tempFile); FAILED(tempHr))
    {
        return failBeforePublication(tempHr);
    }

    const uint64_t baseCompleted = concurrentOverallBytes ? 0 : progress.completedBytes;

    TransferProgressContext downloadCtx{};
    downloadCtx.progress               = &progress;
    downloadCtx.sourcePath             = sourceFullPath;
    downloadCtx.destinationPath        = destinationFullPath;
    downloadCtx.baseCompletedBytes     = baseCompleted;
    downloadCtx.concurrentOverallBytes = concurrentOverallBytes;
    downloadCtx.itemTotalBytes         = sourceSize.known ? sourceSize.sizeBytes : 0u;
    downloadCtx.isUpload               = false;
    downloadCtx.scaleForCopy           = true;
    downloadCtx.scaleForCopySecond     = false;

    ConnectionConcurrencyLimiter::Permit downloadPermit;
    if (! sourceConn.limiterKey.empty())
    {
        auto shouldCancel = [&]() noexcept { return FAILED(progress.CheckCancel()); };
        downloadPermit    = GetConnectionConcurrencyLimiter().AcquireCopyMove(sourceConn.limiterKey, sourceConn.effectiveCopyMoveMaxConcurrency, shouldCancel);
        if (! downloadPermit)
        {
            return failBeforePublication(HRESULT_FROM_WIN32(ERROR_CANCELLED));
        }
    }

    enterPhase(CopyFilePhase::SourceDownload);
    hr = CurlDownloadToFile(
        sourceConn, sourceRemotePath, tempFile.get(), nullptr, &downloadCtx, sourceSize.known ? std::optional<uint64_t>{sourceSize.sizeBytes} : std::nullopt);
    if (FAILED(hr))
    {
        return failBeforePublication(hr);
    }

    // Same-connection copies acquire download then upload permits sequentially.
    // Release the download permit once the temp file is materialized so a
    // single-slot limiter does not deadlock when source and destination share
    // the same remote connection.
    if (! sourceConn.limiterKey.empty() && sourceConn.limiterKey == destinationConn.limiterKey)
    {
        downloadPermit = {};
    }

    enterPhase(CopyFilePhase::LocalSourceValidation);
    uint64_t fileSize = 0;
    hr                = GetFileSizeBytes(tempFile.get(), fileSize);
    if (FAILED(hr))
    {
        return failBeforePublication(hr);
    }
    if (sourceSize.known && fileSize != sourceSize.sizeBytes)
    {
        return failBeforePublication(HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY));
    }

    hr = ResetFilePointerToStart(tempFile.get());
    if (FAILED(hr))
    {
        return failBeforePublication(hr);
    }

    hr = progress.CheckCancel();
    if (FAILED(hr))
    {
        return failBeforePublication(hr);
    }

    TransferProgressContext uploadCtx{};
    uploadCtx.progress               = &progress;
    uploadCtx.sourcePath             = sourceFullPath;
    uploadCtx.destinationPath        = destinationFullPath;
    uploadCtx.baseCompletedBytes     = baseCompleted;
    uploadCtx.concurrentOverallBytes = concurrentOverallBytes;
    uploadCtx.lastConcurrentWireDone = concurrentOverallBytes ? fileSize : 0;
    uploadCtx.itemTotalBytes         = fileSize;
    uploadCtx.isUpload               = true;
    uploadCtx.scaleForCopy           = true;
    uploadCtx.scaleForCopySecond     = true;

    ConnectionConcurrencyLimiter::Permit uploadPermit;
    if (! destinationConn.limiterKey.empty())
    {
        auto shouldCancel = [&]() noexcept { return FAILED(progress.CheckCancel()); };
        uploadPermit =
            GetConnectionConcurrencyLimiter().AcquireCopyMove(destinationConn.limiterKey, destinationConn.effectiveCopyMoveMaxConcurrency, shouldCancel);
        if (! uploadPermit)
        {
            return failBeforePublication(HRESULT_FROM_WIN32(ERROR_CANCELLED));
        }
    }

    std::wstring stagedRemotePath;
    enterPhase(CopyFilePhase::ReserveStagingName);
    hr = GenerateRemoteSiblingPath(destinationConn, destinationRemotePath, L"upload", stagedRemotePath);
    if (FAILED(hr))
    {
        return failBeforePublication(hr);
    }

    enterPhase(CopyFilePhase::Upload);
    hr = CurlUploadFromFile(destinationConn, stagedRemotePath, tempFile.get(), fileSize, nullptr, &uploadCtx);
    if (FAILED(hr))
    {
        publicationResult.RecordPrimaryFailure(hr);
        const HRESULT cleanupHr = RemoteDeleteFile(destinationConn, stagedRemotePath);
        if (FAILED(cleanupHr))
        {
            publicationResult.RecordCleanupDebt(CurlCleanupDebtKind::RetainedStagingSibling, cleanupHr);
        }
        return hr;
    }

    // Verify the staged upload landed at the expected size before promoting it. Prefer a targeted SIZE/stat
    // probe over GetEntryInfo, which lists the whole destination directory. If SIZE cannot prove the exact
    // byte count, use the listing only when it also carries an exact size. Transport success proves what the
    // client sent, not what the server durably retained, so an unknown remote size must fail closed.
    uint64_t verificationRequestCount = 0u;
    enterPhase(CopyFilePhase::VerifyStaging);
    hr = VerifyStagedUploadExactSize(destinationConn, stagedRemotePath, fileSize, verificationRequestCount);
    if (FAILED(hr))
    {
        publicationResult.RecordPrimaryFailure(hr);
        const HRESULT cleanupHr = RemoteDeleteFile(destinationConn, stagedRemotePath);
        if (FAILED(cleanupHr))
        {
            publicationResult.RecordCleanupDebt(CurlCleanupDebtKind::RetainedStagingSibling, cleanupHr);
        }
        return hr;
    }

    enterPhase(CopyFilePhase::Promote);
    const CurlPublicationResult promotionResult =
        PromoteStagedFileToDestination(destinationConn, stagedRemotePath, destinationRemotePath, allowOverwrite, backupPathOut);
    publicationResult.Merge(promotionResult);
    hr = promotionResult.OperationResult();
    if (FAILED(hr))
    {
        return hr;
    }

    enterPhase(CopyFilePhase::Completion);
    if (! concurrentOverallBytes)
    {
        uint64_t wireTotalBytes = fileSize;
        if (wireTotalBytes > (std::numeric_limits<uint64_t>::max)() - fileSize)
        {
            wireTotalBytes = (std::numeric_limits<uint64_t>::max)();
        }
        else
        {
            wireTotalBytes += fileSize;
        }

        progress.completedBytes = (baseCompleted > (std::numeric_limits<uint64_t>::max)() - wireTotalBytes) ? (std::numeric_limits<uint64_t>::max)()
                                                                                                            : (baseCompleted + wireTotalBytes);

        hr = progress.ReportProgress(fileSize, fileSize, sourceFullPath, destinationFullPath);
        if (FAILED(hr))
        {
            return hr;
        }
    }
    else
    {
        hr = progress.ReportProgressWithCompletedBytes(
            concurrentOverallBytes->load(std::memory_order_acquire), fileSize, fileSize, sourceFullPath, destinationFullPath);
        if (FAILED(hr))
        {
            return hr;
        }
    }
    return S_OK;
}

[[nodiscard]] HRESULT CopyDirectoryTree(const ConnectionInfo& sourceConn,
                                        std::wstring_view sourceRemoteDir,
                                        std::wstring_view sourceFullDir,
                                        const ConnectionInfo& destinationConn,
                                        std::wstring_view destinationRemoteDir,
                                        std::wstring_view destinationFullDir,
                                        FileSystemFlags flags,
                                        unsigned int maxConcurrency,
                                        FileOperationProgress& progress,
                                        std::atomic<uint64_t>* concurrentOverallBytes,
                                        bool requireAuthoritativeSourceSize,
                                        const SourceSizeCommitmentMap* preflightCommitments,
                                        CurlPublicationAccumulator* publicationAccumulator = nullptr) noexcept
{
    using namespace Common::FileOperations;
    const auto started         = std::chrono::steady_clock::now();
    HRESULT result             = E_PENDING;
    uint64_t peakFrames        = 0u;
    uint64_t peakQueued        = 0u;
    uint64_t peakPathBytes     = 0u;
    uint64_t peakMetadataBytes = 0u;
    CopyFilePhaseTimings phaseTimings;
    const auto reportMetrics = wil::scope_exit([&]() noexcept
    {
        const std::wstring detail = std::format(L"concurrency={};move={}", maxConcurrency, requireAuthoritativeSourceSize);
        Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.Traversal", detail.c_str(), Debug::Perf::ElapsedUs(started), peakFrames, peakQueued, result);
        Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.Retained", detail.c_str(), 0u, peakPathBytes, peakMetadataBytes, result);
        for (size_t index = 0u; index < kCopyFilePhaseNames.size(); ++index)
        {
            const std::wstring phaseDetail = std::format(L"{};phase={}", detail, kCopyFilePhaseNames[index]);
            Debug::Perf::Emit(L"FileOps.Curl.NativeCopy.FilePhase",
                              phaseDetail.c_str(),
                              phaseTimings.elapsedUs[index].load(std::memory_order_relaxed),
                              phaseTimings.visits[index].load(std::memory_order_relaxed),
                              0u,
                              result);
        }
    });
    const auto finish        = [&](HRESULT hr) noexcept
    {
        result = NormalizeCancellation(hr);
        return result;
    };
    struct CopyFileWorkItem final
    {
        std::wstring sourceRemotePath;
        std::wstring sourceDisplayPath;
        std::wstring destinationRemotePath;
        std::wstring destinationDisplayPath;
        uint64_t expectedSizeBytes = 0u;
        bool expectedSizeKnown     = false;
        bool sourceSizeResolved    = false;

        [[nodiscard]] uint64_t PathBytes() const noexcept
        {
            return static_cast<uint64_t>(sourceRemotePath.capacity() + sourceDisplayPath.capacity() + destinationRemotePath.capacity() +
                                         destinationDisplayPath.capacity() + 4u) *
                   sizeof(wchar_t);
        }
    };
    struct Frame final
    {
        Frame()                        = default;
        Frame(const Frame&)            = delete;
        Frame& operator=(const Frame&) = delete;
        Frame(Frame&&)                 = delete;
        Frame& operator=(Frame&&)      = delete;
        CopyFileWorkItem paths;
        CurlDirectoryCursor cursor;
        std::vector<CopyFileWorkItem> pendingDirectories;
        uint64_t pathBytes    = 0u;
        bool destinationReady = false;
    };
    // The same iterative discovery loop feeds serial or scheduler-owned batches.
    // A batch drains before child discovery; no recursive producer, full-directory
    // buffer, or worker waiting for producer-owned work survives this call.
    const bool continueOnError     = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    const unsigned int concurrency = progress.callback ? std::clamp(maxConcurrency, 1u, 8u) : 1u;
    const size_t batchLimit        = concurrency == 1u ? 1u : DiscoveryQueueTarget(concurrency, true);
    std::vector<std::unique_ptr<Frame>> frames;
    frames.reserve(static_cast<size_t>(kTraversalMaxDepth + 1u));
    std::vector<CopyFileWorkItem> batch;
    batch.reserve(batchLimit);
    const uint64_t slotsBytes = sizeof(phaseTimings) + static_cast<uint64_t>(frames.capacity()) * sizeof(std::unique_ptr<Frame>) +
                                static_cast<uint64_t>(batch.capacity()) * (sizeof(CopyFileWorkItem) + kTraversalRecordOverheadBytes);
    uint64_t framePathBytes   = 0u;
    uint64_t batchPathBytes   = 0u;
    bool hadFailure           = false;
    std::atomic<uint64_t> localOverallBytes{0u};
    if (concurrency > 1u && ! concurrentOverallBytes)
    {
        concurrentOverallBytes = &localOverallBytes;
    }
    const auto sample = [&]() noexcept
    {
        peakFrames    = (std::max)(peakFrames, static_cast<uint64_t>(frames.size()));
        peakQueued    = (std::max)(peakQueued, static_cast<uint64_t>(batch.size()));
        peakPathBytes = (std::max)(peakPathBytes, framePathBytes + batchPathBytes);
        peakMetadataBytes =
            (std::max)(peakMetadataBytes, slotsBytes + static_cast<uint64_t>(frames.size()) * (sizeof(Frame) + CurlDirectoryCursor::kMetadataReservationBytes));
    };
    const auto pushFrame = [&](CopyFileWorkItem paths) noexcept -> HRESULT
    {
        if (frames.size() >= kTraversalMaxDepth + 1u ||
            slotsBytes + (frames.size() + 1u) * (sizeof(Frame) + CurlDirectoryCursor::kMetadataReservationBytes) > kTraversalMaxMetadataBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        auto frame       = std::make_unique<Frame>();
        frame->paths     = std::move(paths);
        frame->pathBytes = frame->paths.PathBytes();
        if (frame->pathBytes > kTraversalMaxQueuedPathBytes - framePathBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        const HRESULT hr =
            frame->cursor.Open(sourceConn, EnsureTrailingSlash(frame->paths.sourceRemotePath), [&progress]() noexcept { return progress.CheckCancel(); });
        if (FAILED(hr))
        {
            return hr;
        }
        frame->pathBytes += frame->cursor.RetainedPathBytes();
        if (frame->pathBytes > kTraversalMaxQueuedPathBytes - framePathBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        framePathBytes += frame->pathBytes;
        frames.push_back(std::move(frame));
        sample();
        return S_OK;
    };
    const auto popFrame = [&]() noexcept
    {
        framePathBytes -= frames.back()->pathBytes;
        frames.pop_back();
    };
    const auto mustStop = [&](HRESULT hr) noexcept
    {
        return ! continueOnError || NormalizeCancellation(hr) == HRESULT_FROM_WIN32(ERROR_CANCELLED) || IsAuthenticationFailureHr(hr) ||
               hr == HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    };
    const auto drainBatch = [&]() noexcept -> HRESULT
    {
        if (batch.empty())
        {
            return S_OK;
        }
        std::atomic<HRESULT> fatalFailure{S_OK};
        std::atomic_bool anyFailure{false};
        const auto processFile = [&](size_t index, uint64_t streamId) noexcept
        {
            if (FAILED(fatalFailure.load(std::memory_order_acquire)))
            {
                return;
            }
            FileOperationProgress::ProgressStreamScope streamScope(streamId);
            const CopyFileWorkItem& item = batch[index];
            HRESULT hr                   = progress.CheckCancel();
            if (SUCCEEDED(hr))
            {
                hr = CopyFileViaTemp(sourceConn,
                                     item.sourceRemotePath,
                                     item.sourceDisplayPath,
                                     destinationConn,
                                     item.destinationRemotePath,
                                     item.destinationDisplayPath,
                                     flags,
                                     progress,
                                     item.expectedSizeBytes,
                                     item.expectedSizeKnown,
                                     requireAuthoritativeSourceSize,
                                     item.sourceSizeResolved,
                                     concurrentOverallBytes,
                                     nullptr,
                                     publicationAccumulator,
                                     &phaseTimings);
            }
            hr = NormalizeCancellation(hr);
            if (FAILED(hr))
            {
                anyFailure.store(true, std::memory_order_release);
                if (mustStop(hr))
                {
                    HRESULT expected = S_OK;
                    static_cast<void>(fatalFailure.compare_exchange_strong(expected, hr, std::memory_order_acq_rel));
                }
            }
        };
        if (concurrency == 1u || batch.size() == 1u)
        {
            for (size_t index = 0u; index < batch.size(); ++index)
            {
                processFile(index, 0u);
            }
        }
        else
        {
            auto job = GetSharedCopyMoveJobScheduler().StartJob(concurrency, batch.size(), processFile);
            GetSharedCopyMoveJobScheduler().WaitJob(job);
        }
        batch.clear();
        batchPathBytes = 0u;
        hadFailure     = hadFailure || anyFailure.load(std::memory_order_acquire);
        return fatalFailure.load(std::memory_order_acquire);
    };
    HRESULT hr = pushFrame(
        {NormalizePluginPath(sourceRemoteDir), std::wstring(sourceFullDir), NormalizePluginPath(destinationRemoteDir), std::wstring(destinationFullDir)});
    if (FAILED(hr))
    {
        return finish(hr);
    }
    while (! frames.empty())
    {
        Frame& frame = *frames.back();
        FilesInformationCurl::Entry entry{};
        hr = frame.cursor.Next(entry);
        if (FAILED(hr))
        {
            // Undispatched rows are discarded after discovery failure. Already
            // published siblings remain published; the caller preserves that truth.
            batch.clear();
            batchPathBytes = 0u;
            if (frames.size() == 1u || mustStop(hr))
            {
                return finish(hr);
            }
            hadFailure = true;
            popFrame();
            continue;
        }
        const bool complete = hr == S_FALSE;
        if (! frame.destinationReady)
        {
            // Validate the first entry (or a complete empty listing) before creating
            // this destination. A malformed first row is not an empty directory.
            hr = progress.CheckCancel();
            if (SUCCEEDED(hr))
            {
                hr = EnsureDirectoryExists(destinationConn, frame.paths.destinationRemotePath);
            }
            if (FAILED(hr))
            {
                if (frames.size() == 1u || mustStop(hr))
                {
                    return finish(hr);
                }
                hadFailure = true;
                popFrame();
                continue;
            }
            frame.destinationReady = true;
        }
        if (complete)
        {
            hr = drainBatch();
            if (FAILED(hr))
            {
                return finish(hr);
            }
            if (! frame.pendingDirectories.empty())
            {
                CopyFileWorkItem child         = std::move(frame.pendingDirectories.front());
                const uint64_t pendingBytes    = child.PathBytes();
                frame.pendingDirectories.erase(frame.pendingDirectories.begin());
                if (pendingBytes > frame.pathBytes || pendingBytes > framePathBytes)
                {
                    return finish(E_UNEXPECTED);
                }
                frame.pathBytes -= pendingBytes;
                framePathBytes -= pendingBytes;
                hr = pushFrame(std::move(child));
                if (FAILED(hr))
                {
                    if (mustStop(hr))
                    {
                        return finish(hr);
                    }
                    hadFailure = true;
                }
                continue;
            }
            popFrame();
            continue;
        }
        // Discovery happens here, on the single-threaded enumeration, not on the worker that later
        // transfers the item: this is the moment the call learns the work exists.
        progress.NoteDiscoveredEntry(entry);
        CopyFileWorkItem child{JoinPluginPath(frame.paths.sourceRemotePath, entry.name),
                               JoinDisplayPath(frame.paths.sourceDisplayPath, entry.name),
                               JoinPluginPath(frame.paths.destinationRemotePath, entry.name),
                               JoinDisplayPath(frame.paths.destinationDisplayPath, entry.name),
                               entry.sizeBytes,
                               entry.sizeKnown,
                               false};
        if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
        {
            if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
            {
                if (! continueOnError)
                {
                    return finish(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
                }
                hadFailure = true;
                continue;
            }
            const uint64_t pendingBytes = child.PathBytes();
            if (pendingBytes > kTraversalMaxQueuedPathBytes - framePathBytes - batchPathBytes)
            {
                return finish(HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW));
            }
            frame.pendingDirectories.push_back(std::move(child));
            frame.pathBytes += pendingBytes;
            framePathBytes += pendingBytes;
            sample();
            continue;
        }
        if ((entry.attributes & FILE_ATTRIBUTE_DEVICE) != 0u)
        {
            if (! continueOnError)
            {
                return finish(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
            }
            hadFailure = true;
            continue;
        }
        if (preflightCommitments)
        {
            const std::optional<uint64_t> commitment = preflightCommitments->Find(child.sourceRemotePath);
            if (! commitment.has_value())
            {
                return finish(HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY));
            }
            child.expectedSizeBytes  = commitment.value();
            child.expectedSizeKnown  = true;
            child.sourceSizeResolved = true;
        }
        const uint64_t itemBytes = child.PathBytes();
        if (itemBytes > kTraversalMaxQueuedPathBytes - framePathBytes)
        {
            return finish(HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW));
        }
        if (itemBytes > kTraversalMaxQueuedPathBytes - framePathBytes - batchPathBytes)
        {
            hr = drainBatch(); // Backpressure, not an aggregate tree-size ceiling.
            if (FAILED(hr))
            {
                return finish(hr);
            }
        }
        batchPathBytes += itemBytes;
        batch.push_back(std::move(child));
        sample();
        if (batch.size() >= batchLimit)
        {
            hr = drainBatch();
            if (FAILED(hr))
            {
                return finish(hr);
            }
        }
    }
    return finish(hadFailure ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK);
}

struct DeleteTreeWorkItem final
{
    std::wstring remotePath;
    std::wstring displayPath;
};

[[nodiscard]] HRESULT DeleteDirectoryTree(const ConnectionInfo& conn,
                                          std::wstring_view directoryRemotePath,
                                          std::wstring_view directoryFullPath,
                                          FileSystemFlags flags,
                                          ConnectionConcurrencyLimiter::Kind kind,
                                          unsigned int requestedConcurrency,
                                          FileOperationProgress& progress,
                                          std::atomic_bool* mutationAttempted,
                                          const SourceTreePathSet* allowedMembers) noexcept
{
    using namespace Common::FileOperations;
    const auto started         = std::chrono::steady_clock::now();
    HRESULT result             = E_PENDING;
    uint64_t peakFrames        = 0u;
    uint64_t peakQueued        = 0u;
    uint64_t peakPathBytes     = 0u;
    uint64_t peakMetadataBytes = 0u;
    const auto reportMetrics   = wil::scope_exit([&]() noexcept
    {
        const std::wstring detail = std::format(L"concurrency={}", requestedConcurrency);
        Debug::Perf::Emit(L"FileOps.Curl.NativeDelete.Traversal", detail.c_str(), Debug::Perf::ElapsedUs(started), peakFrames, peakQueued, result);
        Debug::Perf::Emit(L"FileOps.Curl.NativeDelete.Retained", detail.c_str(), 0u, peakPathBytes, peakMetadataBytes, result);
    });
    const auto finish          = [&](HRESULT hr) noexcept
    {
        result = NormalizeCancellation(hr);
        return result;
    };
    if (NormalizePluginPath(directoryRemotePath) == L"/")
    {
        return finish(HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED));
    }

    // One cursor per active ancestry frame, not one vector per directory or tree.
    // The only work queue is a batch that drains before entering another child.
    // Thus a late LIST never depends on undispatched work in an earlier subtree.
    struct Frame final
    {
        Frame()                        = default;
        Frame(const Frame&)            = delete;
        Frame& operator=(const Frame&) = delete;
        Frame(Frame&&)                 = delete;
        Frame& operator=(Frame&&)      = delete;
        std::wstring remotePath;
        std::wstring displayPath;
        CurlDirectoryCursor cursor;
        std::vector<DeleteTreeWorkItem> pendingDirectories;
        uint64_t pathBytes  = 0u;
        bool failedChildren = false;
    };
    std::vector<std::unique_ptr<Frame>> frames;
    std::vector<DeleteTreeWorkItem> batch;
    uint64_t framePathBytes        = 0u;
    uint64_t batchPathBytes        = 0u;
    bool hadFailure                = false;
    const bool continueOnError     = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    const unsigned int concurrency = std::clamp(requestedConcurrency, 1u, 8u);
    const size_t batchLimit        = concurrency == 1u ? 1u : DiscoveryQueueTarget(concurrency, true);
    const auto sample              = [&]() noexcept
    {
        peakFrames        = (std::max)(peakFrames, static_cast<uint64_t>(frames.size()));
        peakQueued        = (std::max)(peakQueued, static_cast<uint64_t>(batch.size()));
        peakPathBytes     = (std::max)(peakPathBytes, framePathBytes + batchPathBytes);
        peakMetadataBytes = (std::max)(peakMetadataBytes,
                                       static_cast<uint64_t>(frames.size()) * (CurlDirectoryCursor::kMetadataReservationBytes + sizeof(Frame)) +
                                           static_cast<uint64_t>(frames.capacity()) * sizeof(std::unique_ptr<Frame>) +
                                           static_cast<uint64_t>(batch.capacity()) * sizeof(DeleteTreeWorkItem) +
                                           static_cast<uint64_t>(batch.size()) * kTraversalRecordOverheadBytes);
    };
    const auto pushFrame = [&](std::wstring remotePath, std::wstring displayPath) noexcept -> HRESULT
    {
        if (frames.size() >= kTraversalMaxDepth + 1u ||
            (frames.size() + 1u) * (CurlDirectoryCursor::kMetadataReservationBytes + sizeof(Frame)) > kTraversalMaxMetadataBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        auto frame         = std::make_unique<Frame>();
        frame->remotePath  = NormalizePluginPath(remotePath);
        frame->displayPath = std::move(displayPath);
        frame->pathBytes   = static_cast<uint64_t>(frame->remotePath.capacity() + frame->displayPath.capacity() + 2u) * sizeof(wchar_t);
        if (frame->pathBytes > kTraversalMaxQueuedPathBytes - framePathBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        HRESULT hr = frame->cursor.Open(conn, EnsureTrailingSlash(frame->remotePath), [&progress]() noexcept { return progress.CheckCancel(); });
        if (FAILED(hr))
        {
            return hr;
        }
        frame->pathBytes += frame->cursor.RetainedPathBytes();
        if (frame->pathBytes > kTraversalMaxQueuedPathBytes - framePathBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
        }
        framePathBytes += frame->pathBytes;
        frames.push_back(std::move(frame));
        sample();
        return S_OK;
    };
    const auto popFrame = [&]() noexcept
    {
        framePathBytes -= frames.back()->pathBytes;
        frames.pop_back();
    };
    const auto drainBatch = [&]() noexcept -> HRESULT
    {
        if (batch.empty())
        {
            return S_OK;
        }
        std::atomic<HRESULT> fatalFailure{S_OK};
        std::atomic_bool anyFailure{false};
        const auto processFile = [&](size_t index, uint64_t streamId) noexcept
        {
            if (FAILED(fatalFailure.load(std::memory_order_acquire)))
            {
                return;
            }
            FileOperationProgress::ProgressStreamScope streamScope(streamId);
            const DeleteTreeWorkItem& item = batch[index];
            HRESULT hr                     = progress.CheckCancel();
            if (SUCCEEDED(hr))
            {
                hr = progress.ReportProgress(0u, 0u, item.displayPath, {});
            }
            if (SUCCEEDED(hr))
            {
                hr = RemoteDeleteFileWithPermit(conn, item.remotePath, progress, kind, mutationAttempted);
            }
            hr = NormalizeCancellation(hr);
            if (FAILED(hr))
            {
                anyFailure.store(true, std::memory_order_release);
                if (! continueOnError || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || IsAuthenticationFailureHr(hr))
                {
                    HRESULT expected = S_OK;
                    static_cast<void>(fatalFailure.compare_exchange_strong(expected, hr, std::memory_order_acq_rel));
                }
            }
        };
        if (concurrency == 1u || batch.size() == 1u)
        {
            for (size_t index = 0u; index < batch.size(); ++index)
            {
                processFile(index, 0u);
            }
        }
        else
        {
            auto job = GetSharedCopyMoveJobScheduler().StartJob(concurrency, batch.size(), processFile);
            GetSharedCopyMoveJobScheduler().WaitJob(job);
        }
        batch.clear();
        batchPathBytes      = 0u;
        const HRESULT fatal = fatalFailure.load(std::memory_order_acquire);
        if (FAILED(fatal))
        {
            return fatal;
        }
        return anyFailure.load(std::memory_order_acquire) ? S_FALSE : S_OK;
    };
    const auto rememberPartial = [&]() noexcept
    {
        hadFailure = true;
        if (! frames.empty())
        {
            frames.back()->failedChildren = true;
        }
    };

    HRESULT hr = pushFrame(std::wstring(directoryRemotePath), std::wstring(directoryFullPath));
    if (FAILED(hr))
    {
        return finish(hr);
    }
    if (allowedMembers != nullptr && (! allowedMembers->IsFinalized() || ! allowedMembers->Contains(frames.back()->remotePath)))
    {
        // The preflight and the delete walker must spell the same member paths. If the root itself
        // is not a member, nothing below could match either; stop before any deletion rather than
        // silently keeping the whole source as "late writers".
        Debug::Perf::EmitCounter(L"FileOps.Curl.Move.MemberParityRefused");
        return finish(HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
    }
    while (! frames.empty())
    {
        Frame& frame = *frames.back();
        FilesInformationCurl::Entry entry{};
        hr = frame.cursor.Next(entry);
        if (FAILED(hr))
        {
            // No new work is dispatched after a failed listing. Its already drained
            // batches keep their mutation truth; remaining buffered rows are discarded.
            batch.clear();
            batchPathBytes = 0u;
            // A selected-root listing failure keeps its exact cause. Continuing
            // other selected items belongs to DeleteItems, not this tree walker.
            if (frames.size() == 1u || ! continueOnError || NormalizeCancellation(hr) == HRESULT_FROM_WIN32(ERROR_CANCELLED) || IsAuthenticationFailureHr(hr))
            {
                return finish(hr);
            }
            popFrame();
            rememberPartial();
            continue;
        }
        if (hr == S_FALSE)
        {
            hr = drainBatch();
            if (FAILED(hr))
            {
                return finish(hr);
            }
            if (hr == S_FALSE)
            {
                rememberPartial();
            }
            if (! frame.pendingDirectories.empty())
            {
                DeleteTreeWorkItem child     = std::move(frame.pendingDirectories.front());
                const uint64_t pendingBytes  = static_cast<uint64_t>(child.remotePath.capacity() + child.displayPath.capacity() + 2u) * sizeof(wchar_t);
                frame.pendingDirectories.erase(frame.pendingDirectories.begin());
                if (pendingBytes > frame.pathBytes || pendingBytes > framePathBytes)
                {
                    return finish(E_UNEXPECTED);
                }
                frame.pathBytes -= pendingBytes;
                framePathBytes -= pendingBytes;
                hr = pushFrame(std::move(child.remotePath), std::move(child.displayPath));
                if (FAILED(hr))
                {
                    if (! continueOnError || NormalizeCancellation(hr) == HRESULT_FROM_WIN32(ERROR_CANCELLED) || IsAuthenticationFailureHr(hr))
                    {
                        return finish(hr);
                    }
                    rememberPartial();
                }
                continue;
            }
            if (! frame.failedChildren)
            {
                hr = progress.CheckCancel();
                if (SUCCEEDED(hr))
                {
                    hr = progress.ReportProgress(0u, 0u, frame.displayPath, {});
                }
                if (SUCCEEDED(hr))
                {
                    hr = RemoteRemoveDirectoryWithPermit(conn, frame.remotePath, progress, kind, mutationAttempted);
                }
                if (FAILED(hr))
                {
                    if (! continueOnError || NormalizeCancellation(hr) == HRESULT_FROM_WIN32(ERROR_CANCELLED) || IsAuthenticationFailureHr(hr))
                    {
                        return finish(hr);
                    }
                    rememberPartial();
                }
            }
            const bool incomplete = frame.failedChildren;
            popFrame();
            if (incomplete)
            {
                rememberPartial();
            }
            continue;
        }
        std::wstring childRemote  = JoinPluginPath(frame.remotePath, entry.name);
        std::wstring childDisplay = JoinDisplayPath(frame.displayPath, entry.name);
        if (allowedMembers != nullptr && ! allowedMembers->Contains(childRemote))
        {
            // A late writer created this name after Move preflight. Do not delete
            // an object this operation never copied. It is not this call's work, so it is not
            // this call's discovery either.
            rememberPartial();
            continue;
        }
        // A delete discovers what it is about to remove, on this single-threaded enumeration.
        progress.NoteDiscoveredEntry(entry);
        if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
        {
            if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
            {
                return finish(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
            }
            const uint64_t pendingBytes = static_cast<uint64_t>(childRemote.capacity() + childDisplay.capacity() + 2u) * sizeof(wchar_t);
            if (pendingBytes > kTraversalMaxQueuedPathBytes - framePathBytes - batchPathBytes)
            {
                return finish(HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW));
            }
            frame.pendingDirectories.push_back({std::move(childRemote), std::move(childDisplay)});
            frame.pathBytes += pendingBytes;
            framePathBytes += pendingBytes;
            sample();
            continue;
        }
        const uint64_t itemBytes = static_cast<uint64_t>(childRemote.capacity() + childDisplay.capacity() + 2u) * sizeof(wchar_t);
        if (itemBytes > kTraversalMaxQueuedPathBytes - framePathBytes)
        {
            return finish(HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW));
        }
        if (itemBytes > kTraversalMaxQueuedPathBytes - framePathBytes - batchPathBytes)
        {
            hr = drainBatch(); // Backpressure, never an aggregate-work limit.
            if (FAILED(hr))
            {
                return finish(hr);
            }
            if (hr == S_FALSE)
            {
                rememberPartial();
            }
        }
        batchPathBytes += itemBytes;
        batch.push_back({std::move(childRemote), std::move(childDisplay)});
        sample();
        if (batch.size() >= batchLimit)
        {
            hr = drainBatch();
            if (FAILED(hr))
            {
                return finish(hr);
            }
            if (hr == S_FALSE)
            {
                rememberPartial();
            }
        }
    }
    return finish(hadFailure ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK);
}
} // namespace

namespace
{
// R3-1: a replacement was granted for one occupant; it must still be there, be a file, and be
// unchanged since the user saw it (same listing dialect). Nothing is uploaded otherwise.
[[nodiscard]] HRESULT ValidateCurlReplaceOccupant(const ConnectionInfo& conn,
                                                  std::wstring_view destinationPath,
                                                  const CurlReplaceExpectation& expectation,
                                                  uint64_t& requestCount) noexcept
{
    FilesInformationCurl::Entry occupant{};
    const HRESULT occupantHr = GetEntryInfo(conn, destinationPath, occupant);
    ++requestCount;
    if (occupantHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
    {
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
    if (FAILED(occupantHr))
    {
        return occupantHr;
    }
    if ((occupant.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
    }
    if (expectation.lastWriteTime == 0 || occupant.lastWriteTime == 0)
    {
        // R0-RC3: the last-write time is the only token this transport has; without it the
        // occupant cannot be revalidated and the replacement is refused rather than fail-open.
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    if (occupant.lastWriteTime != expectation.lastWriteTime)
    {
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
    return S_OK;
}
} // namespace

CurlPublicationResult FileSystemCurlInternal::PublishCurlWriterTransaction(const ConnectionInfo& conn,
                                                                            std::wstring_view destinationPath,
                                                                            HANDLE file,
                                                                            uint64_t sizeBytes,
                                                                            bool allowOverwrite,
                                                                            const CurlReplaceExpectation* replaceExpectation,
                                                                            CurlWriterPublicationMetrics& metrics) noexcept
{
    metrics                   = {};
    const auto started        = std::chrono::steady_clock::now();
    const auto recordDuration = wil::scope_exit([&]() noexcept
    { metrics.commitUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count()); });

    if (conn.protocol == Protocol::Imap)
    {
        CurlPublicationResult result{};
        result.RecordPrimaryFailure(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
        return result;
    }

    if (replaceExpectation != nullptr)
    {
        const HRESULT occupantHr = ValidateCurlReplaceOccupant(conn, destinationPath, *replaceExpectation, metrics.requestCount);
        if (FAILED(occupantHr))
        {
            CurlPublicationResult result{};
            result.RecordPrimaryFailure(occupantHr);
            return result;
        }
    }

    std::wstring stagedRemotePath;
    HRESULT hr = GenerateRemoteSiblingPath(conn, destinationPath, L"writer", stagedRemotePath);
    ++metrics.requestCount;
    if (FAILED(hr))
    {
        CurlPublicationResult result{};
        result.RecordPrimaryFailure(hr);
        return result;
    }

    hr = CurlUploadFromFile(conn, stagedRemotePath, file, sizeBytes, nullptr, nullptr);
    ++metrics.requestCount;
    if (FAILED(hr))
    {
        ++metrics.cleanupAttemptCount;
        CurlPublicationResult result{};
        result.RecordPrimaryFailure(hr);
        const HRESULT cleanupHr = RemoteDeleteFile(conn, stagedRemotePath);
        if (FAILED(cleanupHr))
        {
            result.RecordCleanupDebt(CurlCleanupDebtKind::RetainedStagingSibling, cleanupHr);
        }
        return result;
    }
    metrics.stagedBytes = sizeBytes;

    hr = VerifyStagedUploadExactSize(conn, stagedRemotePath, sizeBytes, metrics.requestCount);
    if (FAILED(hr))
    {
        ++metrics.cleanupAttemptCount;
        CurlPublicationResult result{};
        result.RecordPrimaryFailure(hr);
        const HRESULT cleanupHr = RemoteDeleteFile(conn, stagedRemotePath);
        if (FAILED(cleanupHr))
        {
            result.RecordCleanupDebt(CurlCleanupDebtKind::RetainedStagingSibling, cleanupHr);
        }
        return result;
    }

    ++metrics.requestCount;
    CurlPublicationResult result = PromoteStagedFileToDestination(conn, stagedRemotePath, destinationPath, allowOverwrite, nullptr);
    if (FAILED(result.primaryMutationHr))
    {
        ++metrics.cleanupAttemptCount;
    }
    return result;
}

#if defined(ENABLE_TESTS)
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlDebugSelfTests(unsigned int* passed, unsigned int* failed)
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }

    *passed = 0;
    *failed = 0;

    RunDebugRemoteSiblingLeafEntropySelfTest(*passed, *failed);
    RunDebugOverwriteCleanupContractSelfTest(*passed, *failed);
    RunDebugCaseOnlyRenameContractSelfTest(*passed, *failed);
    RunDebugCurlStreamingReaderContractSelfTests(*passed, *failed);
    RunDebugCurlWriterOwnershipContractSelfTests(*passed, *failed);
    RunCurlShortSuccessfulUploadSelfTests(*passed, *failed);
    RunCurlStalledControlCommandCancelSelfTests(*passed, *failed);
    RunCurlStalledReaderCancelSelfTests(*passed, *failed);
    RunCurlParallelWritersSelfTests(*passed, *failed);

    return *failed == 0u ? S_OK : E_FAIL;
}
#endif

HRESULT STDMETHODCALLTYPE FileSystemCurl::CopyItem(const wchar_t* sourcePath,
                                                   const wchar_t* destinationPath,
                                                   FileSystemFlags flags,
                                                   const FileSystemOptions* options,
                                                   IFileSystemCallback* callback,
                                                   void* cookie) noexcept
{
    const CurlOperationOptionsScope operationOptionsScope(options);
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (const HRESULT admissionHr = ValidateConditionalMutationAdmission(flags, options); FAILED(admissionHr))
    {
        return admissionHr;
    }

    CurlPublicationAccumulator publicationAccumulator;
    auto observeCleanupDebt = wil::scope_exit([&]() noexcept { ObserveCurlCleanupDebt(publicationAccumulator.Snapshot()); });

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FileOperationProgress progress{};
    HRESULT hr = progress.Initialize(FILESYSTEM_COPY, 1, options, callback, cookie);
    if (FAILED(hr))
    {
        return hr;
    }

    const unsigned int requestedConcurrency = std::clamp(settings.copyMoveMaxConcurrency, 1u, 8u);

    const std::wstring sourceDisplay      = BuildDisplayPath(_protocol, sourcePath);
    const std::wstring destinationDisplay = BuildDisplayPath(_protocol, destinationPath);

    hr = progress.ReportProgress(0, 0, sourceDisplay, destinationDisplay);
    if (FAILED(hr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr));
        return hr;
    }

    ResolvedLocation sourceResolved{};
    const HRESULT resolveSourceHr = ResolveLocation(_protocol, settings, sourcePath, _hostConnections.get(), true, sourceResolved);
    if (FAILED(resolveSourceHr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, resolveSourceHr));
        return resolveSourceHr;
    }

    ResolvedLocation destinationResolved{};
    const HRESULT resolveDestinationHr = ResolveLocation(_protocol, settings, destinationPath, _hostConnections.get(), true, destinationResolved);
    if (FAILED(resolveDestinationHr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, resolveDestinationHr));
        return resolveDestinationHr;
    }

    FilesInformationCurl::Entry sourceInfo{};
    hr = GetEntryInfo(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
    if (FAILED(hr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr));
        return hr;
    }

    if ((sourceInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        // The selected root itself; the walk below reports each descendant as it finds it.
        progress.NoteDiscoveredEntry(sourceInfo);
        if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
        {
            hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        else
        {
            bool destinationExisted = false;
            hr                      = QueryPathExists(destinationResolved.connection, destinationResolved.remotePath, destinationExisted);
            if (SUCCEEDED(hr))
            {
                hr = CopyDirectoryTree(sourceResolved.connection,
                                            EnsureTrailingSlash(sourceResolved.remotePath),
                                            EnsureTrailingSlashDisplay(sourceDisplay),
                                            destinationResolved.connection,
                                            EnsureTrailingSlash(destinationResolved.remotePath),
                                            EnsureTrailingSlashDisplay(destinationDisplay),
                                            flags,
                                            requestedConcurrency,
                                            progress,
                                            nullptr,
                                            false,
                                            nullptr,
                                            &publicationAccumulator);
                if (FAILED(hr) && ! destinationExisted)
                {
                    const CurlPublicationResult recoveryResult = PreserveCopiedDirectoryAfterFailure(hr);
                    publicationAccumulator.Merge(recoveryResult);
                    hr = recoveryResult.OperationResult();
                }
            }
        }
    }
    else
    {
        // A known leaf: its exact size is already in hand, so the call's total is complete here.
        progress.NoteDiscoveredEntry(sourceInfo);
        progress.CloseDiscovery();
        hr = CopyFileViaTemp(sourceResolved.connection,
                             sourceResolved.remotePath,
                             sourceDisplay,
                             destinationResolved.connection,
                             destinationResolved.remotePath,
                             destinationDisplay,
                             flags,
                             progress,
                             sourceInfo.sizeBytes,
                             sourceInfo.sizeKnown,
                             false,
                             false,
                             nullptr,
                             nullptr,
                             &publicationAccumulator);
    }

    progress.completedItems = 1;
    const HRESULT cbHr      = progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr);
    const HRESULT resultHr  = FAILED(cbHr) ? cbHr : hr;
    if (SUCCEEDED(resultHr))
    {
        NotifySyntheticPathCreated(destinationPath);
    }
    return resultHr;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::MoveItem(const wchar_t* sourcePath,
                                                   const wchar_t* destinationPath,
                                                   FileSystemFlags flags,
                                                   const FileSystemOptions* options,
                                                   IFileSystemCallback* callback,
                                                   void* cookie) noexcept
{
    const CurlOperationOptionsScope operationOptionsScope(options);
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (const HRESULT admissionHr = ValidateConditionalMutationAdmission(flags, options); FAILED(admissionHr))
    {
        return admissionHr;
    }

    CurlPublicationAccumulator publicationAccumulator;
    auto observeCleanupDebt = wil::scope_exit([&]() noexcept { ObserveCurlCleanupDebt(publicationAccumulator.Snapshot()); });

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FileOperationProgress progress{};
    HRESULT hr = progress.Initialize(FILESYSTEM_MOVE, 1, options, callback, cookie);
    if (FAILED(hr))
    {
        return hr;
    }
    // Progress callbacks may adjust live options, not widen the admitted Move route.
    const bool nativeMoveOnly = progress.options.moveMode == FILESYSTEM_MOVE_NATIVE_ONLY;

    const unsigned int requestedConcurrency = std::clamp(settings.copyMoveMaxConcurrency, 1u, 8u);

    const std::wstring sourceDisplay      = BuildDisplayPath(_protocol, sourcePath);
    const std::wstring destinationDisplay = BuildDisplayPath(_protocol, destinationPath);

    hr = progress.ReportProgress(0, 0, sourceDisplay, destinationDisplay);
    if (FAILED(hr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr));
        return hr;
    }

    ResolvedLocation sourceResolved{};
    const HRESULT resolveSourceHr = ResolveLocation(_protocol, settings, sourcePath, _hostConnections.get(), true, sourceResolved);
    if (FAILED(resolveSourceHr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, resolveSourceHr));
        return resolveSourceHr;
    }

    ResolvedLocation destinationResolved{};
    const HRESULT resolveDestinationHr = ResolveLocation(_protocol, settings, destinationPath, _hostConnections.get(), true, destinationResolved);
    if (FAILED(resolveDestinationHr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, resolveDestinationHr));
        return resolveDestinationHr;
    }

    const bool allowOverwrite = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);

    bool sourceObservedWithoutMutation = false;
    if (CanServerSideRename(sourceResolved.connection, destinationResolved.connection))
    {
        FilesInformationCurl::Entry sourceInfo{};
        hr                            = GetEntryInfo(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
        sourceObservedWithoutMutation = SUCCEEDED(hr);
        const bool isSelfRename       = IsCaseSensitiveSelfRename(sourceResolved.remotePath, destinationResolved.remotePath);
        if (SUCCEEDED(hr))
        {
            // A server-side rename relocates the object itself, so a directory discovers one
            // directory and none of its descendants: none of them are visited. When the destination
            // is an existing directory the host continues the item as a rename merge and stages the
            // closure itself; the provider still reports its own truth here.
            progress.NoteDiscoveredEntry(sourceInfo);
            progress.CloseDiscovery();
        }
        if (SUCCEEDED(hr) && ! isSelfRename)
        {
            const CurlPublicationResult renameResult = RenameWithOverwriteRollback(
                sourceResolved.connection, sourceResolved.remotePath, destinationResolved.remotePath, allowOverwrite, sourceObservedWithoutMutation);
            publicationAccumulator.Merge(renameResult);
            hr = renameResult.OperationResult();
        }
        else if (SUCCEEDED(hr))
        {
            hr = S_OK;
        }
    }
    else
    {
        FilesInformationCurl::Entry sourceInfo{};
        hr                            = GetEntryInfo(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
        sourceObservedWithoutMutation = SUCCEEDED(hr);
        if (SUCCEEDED(hr))
        {
            // The copy-then-delete fallback: the selected root, with the walk below adding each
            // descendant it finds. A leaf is the whole total already.
            progress.NoteDiscoveredEntry(sourceInfo);
            if ((sourceInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
            {
                progress.CloseDiscovery();
            }
            if (nativeMoveOnly)
            {
                hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
            else if ((sourceInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
            {
                if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
                {
                    hr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }
                else
                {
                    SourceSizeCommitmentMap sourceSizeCommitments;
                    SourceTreePathSet sourceTreeMembers;
                    bool destinationExisted = false;
                    hr                      = PreflightDirectorySourceSizes(sourceResolved.connection,
                                                                           EnsureTrailingSlash(sourceResolved.remotePath),
                                                                           progress,
                                                                           sourceSizeCommitments,
                                                                           sourceTreeMembers);
                    if (SUCCEEDED(hr))
                    {
                        hr = QueryPathExists(destinationResolved.connection, destinationResolved.remotePath, destinationExisted);
                    }
                    if (SUCCEEDED(hr))
                    {
                        sourceObservedWithoutMutation = false; // Copy may create staging or destination objects before failing.
                        hr                            = CopyDirectoryTree(sourceResolved.connection,
                                                                               EnsureTrailingSlash(sourceResolved.remotePath),
                                                                               EnsureTrailingSlashDisplay(sourceDisplay),
                                                                               destinationResolved.connection,
                                                                               EnsureTrailingSlash(destinationResolved.remotePath),
                                                                               EnsureTrailingSlashDisplay(destinationDisplay),
                                                                               flags,
                                                                               requestedConcurrency,
                                                                               progress,
                                                                               nullptr,
                                                                               true,
                                                                               &sourceSizeCommitments,
                                                                               &publicationAccumulator);
                        if (FAILED(hr) && ! destinationExisted)
                        {
                            const CurlPublicationResult recoveryResult = PreserveCopiedDirectoryAfterFailure(hr);
                            publicationAccumulator.Merge(recoveryResult);
                            hr = recoveryResult.OperationResult();
                        }
                    }
                    if (SUCCEEDED(hr))
                    {
                        const HRESULT deleteSourceHr = DeleteDirectoryTree(sourceResolved.connection,
                                                                           sourceResolved.remotePath,
                                                                           sourceDisplay,
                                                                           FILESYSTEM_FLAG_RECURSIVE,
                                                                           ConnectionConcurrencyLimiter::Kind::CopyMove,
                                                                           1u,
                                                                           progress,
                                                                           nullptr,
                                                                           &sourceTreeMembers);
                        if (FAILED(deleteSourceHr))
                        {
                            const CurlPublicationResult sourceDeleteResult = PreserveMovedDirectoryAfterSourceDeleteFailure(deleteSourceHr);
                            publicationAccumulator.Merge(sourceDeleteResult);
                            hr = sourceDeleteResult.OperationResult();
                        }
                    }
                }
            }
            else
            {
                std::wstring destinationBackupPath;
                sourceObservedWithoutMutation = false;
                hr                            = CopyFileViaTemp(sourceResolved.connection,
                                                                sourceResolved.remotePath,
                                                                sourceDisplay,
                                                                destinationResolved.connection,
                                                                destinationResolved.remotePath,
                                                                destinationDisplay,
                                                                flags,
                                                                progress,
                                                                sourceInfo.sizeBytes,
                                                                sourceInfo.sizeKnown,
                                                                true,
                                                                false,
                                                                nullptr,
                                                                &destinationBackupPath,
                                                                &publicationAccumulator);
                if (SUCCEEDED(hr))
                {
                    const HRESULT deleteSourceHr = RemoteDeleteFileWithPermit(
                        sourceResolved.connection, sourceResolved.remotePath, progress, ConnectionConcurrencyLimiter::Kind::CopyMove);
                    if (FAILED(deleteSourceHr))
                    {
                        const CurlPublicationResult sourceDeleteResult =
                            PreserveMovedFileDestinationAfterSourceDeleteFailure(deleteSourceHr, destinationBackupPath);
                        publicationAccumulator.Merge(sourceDeleteResult);
                        hr = sourceDeleteResult.OperationResult();
                    }
                    else
                    {
                        const CurlPublicationResult cleanupResult = FinalizeOverwriteTarget(destinationResolved.connection, destinationBackupPath);
                        publicationAccumulator.Merge(cleanupResult);
                        hr = cleanupResult.OperationResult();
                    }
                }
            }
        }
    }

    progress.completedItems = 1;
    const HRESULT cbHr      = progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr, sourceObservedWithoutMutation);
    const HRESULT resultHr  = FAILED(cbHr) ? cbHr : hr;
    if (SUCCEEDED(resultHr))
    {
        NotifySyntheticPathMoved(sourcePath, destinationPath);
    }
    return resultHr;
}

HRESULT STDMETHODCALLTYPE
FileSystemCurl::DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    const CurlOperationOptionsScope operationOptionsScope(options);
    if (! path)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FileOperationProgress progress{};
    HRESULT hr = progress.Initialize(FILESYSTEM_DELETE, 1, options, callback, cookie);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring displayPath = BuildDisplayPath(_protocol, path);
    hr                             = progress.ReportProgress(0, 0, displayPath, {});
    if (FAILED(hr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, displayPath, {}, hr));
        return hr;
    }

    ResolvedLocation resolved{};
    const HRESULT resolveHr = ResolveLocation(_protocol, settings, path, _hostConnections.get(), true, resolved);
    if (FAILED(resolveHr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, displayPath, {}, resolveHr));
        return resolveHr;
    }

    FilesInformationCurl::Entry info{};
    // One selected item's flag spans all child workers and outlives their join. A sibling
    // selected item must not lend or invalidate this item's no-mutation proof.
    std::atomic_bool mutationAttempted{false};
    // Cover nonrecursive removal too; no root probe or mutation may precede refusal.
    hr = resolved.remotePath == L"/" ? HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) : GetEntryInfo(resolved.connection, resolved.remotePath, info);
    const bool sourceObserved = SUCCEEDED(hr);
    if (SUCCEEDED(hr))
    {
        // The selected root; a recursive walk below adds whatever it finds underneath, while a
        // leaf is the whole total already.
        progress.NoteDiscoveredEntry(info);
        if ((info.attributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
        {
            progress.CloseDiscovery();
        }
        if ((info.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            if (HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
            {
                const unsigned int concurrency = std::clamp(resolved.connection.effectiveDeleteMaxConcurrency, 1u, 8u);
                hr                             = DeleteDirectoryTree(resolved.connection,
                                                                     resolved.remotePath,
                                                                     displayPath,
                                                                     flags,
                                                                     ConnectionConcurrencyLimiter::Kind::Delete,
                                                                     concurrency,
                                                                     progress,
                                                                     &mutationAttempted);
            }
            else
            {
                hr = RemoteRemoveDirectoryWithPermit(
                    resolved.connection, resolved.remotePath, progress, ConnectionConcurrencyLimiter::Kind::Delete, &mutationAttempted);
            }
        }
        else
        {
            hr = RemoteDeleteFileWithPermit(resolved.connection, resolved.remotePath, progress, ConnectionConcurrencyLimiter::Kind::Delete, &mutationAttempted);
        }
    }

    progress.completedItems = 1;
    const HRESULT cbHr      = progress.ReportItemCompleted(0, displayPath, {}, hr, sourceObserved && ! mutationAttempted.load(std::memory_order_acquire));
    const HRESULT resultHr  = FAILED(cbHr) ? cbHr : hr;
    if (SUCCEEDED(resultHr))
    {
        NotifySyntheticPathDeleted(path);
    }
    return resultHr;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::RenameItem(const wchar_t* sourcePath,
                                                     const wchar_t* destinationPath,
                                                     FileSystemFlags flags,
                                                     const FileSystemOptions* options,
                                                     IFileSystemCallback* callback,
                                                     void* cookie) noexcept
{
    const CurlOperationOptionsScope operationOptionsScope(options);
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (const HRESULT admissionHr = ValidateConditionalMutationAdmission(flags, options); FAILED(admissionHr))
    {
        return admissionHr;
    }

    CurlPublicationAccumulator publicationAccumulator;
    auto observeCleanupDebt = wil::scope_exit([&]() noexcept { ObserveCurlCleanupDebt(publicationAccumulator.Snapshot()); });

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FileOperationProgress progress{};
    HRESULT hr = progress.Initialize(FILESYSTEM_RENAME, 1, options, callback, cookie);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring sourceDisplay      = BuildDisplayPath(_protocol, sourcePath);
    const std::wstring destinationDisplay = BuildDisplayPath(_protocol, destinationPath);

    hr = progress.ReportProgress(0, 0, sourceDisplay, destinationDisplay);
    if (FAILED(hr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr));
        return hr;
    }

    ResolvedLocation sourceResolved{};
    const HRESULT resolveSourceHr = ResolveLocation(_protocol, settings, sourcePath, _hostConnections.get(), true, sourceResolved);
    if (FAILED(resolveSourceHr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, resolveSourceHr));
        return resolveSourceHr;
    }

    ResolvedLocation destinationResolved{};
    const HRESULT resolveDestinationHr = ResolveLocation(_protocol, settings, destinationPath, _hostConnections.get(), true, destinationResolved);
    if (FAILED(resolveDestinationHr))
    {
        static_cast<void>(progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, resolveDestinationHr));
        return resolveDestinationHr;
    }

    const bool allowOverwrite = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);

    bool sourceObservedWithoutMutation = false;
    if (! CanServerSideRename(sourceResolved.connection, destinationResolved.connection))
    {
        hr = HRESULT_FROM_WIN32(ERROR_NOT_SAME_DEVICE);
    }
    else
    {
        FilesInformationCurl::Entry sourceInfo{};
        hr                            = GetEntryInfo(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
        sourceObservedWithoutMutation = SUCCEEDED(hr);
        const bool isSelfRename       = IsCaseSensitiveSelfRename(sourceResolved.remotePath, destinationResolved.remotePath);
        if (SUCCEEDED(hr))
        {
            // A server-side rename relocates the object itself, so a directory discovers one
            // directory and none of its descendants: none of them are visited. When the destination
            // is an existing directory the host continues the item as a rename merge and stages the
            // closure itself; the provider still reports its own truth here.
            progress.NoteDiscoveredEntry(sourceInfo);
            progress.CloseDiscovery();
        }
        if (SUCCEEDED(hr) && ! isSelfRename)
        {
            const CurlPublicationResult renameResult = RenameWithOverwriteRollback(
                sourceResolved.connection, sourceResolved.remotePath, destinationResolved.remotePath, allowOverwrite, sourceObservedWithoutMutation);
            publicationAccumulator.Merge(renameResult);
            hr = renameResult.OperationResult();
        }
        else if (SUCCEEDED(hr))
        {
            hr = S_OK;
        }
    }

    progress.completedItems = 1;
    const HRESULT cbHr      = progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr, sourceObservedWithoutMutation);
    const HRESULT resultHr  = FAILED(cbHr) ? cbHr : hr;
    if (SUCCEEDED(resultHr))
    {
        NotifySyntheticPathMoved(sourcePath, destinationPath);
    }
    return resultHr;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::CopyItems(const wchar_t* const* sourcePaths,
                                                    unsigned long count,
                                                    const wchar_t* destinationFolder,
                                                    FileSystemFlags flags,
                                                    const FileSystemOptions* options,
                                                    IFileSystemCallback* callback,
                                                    void* cookie) noexcept
{
    const CurlOperationOptionsScope operationOptionsScope(options);
    if (! sourcePaths || ! destinationFolder)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (destinationFolder[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (const HRESULT admissionHr = ValidateConditionalMutationAdmission(flags, options); FAILED(admissionHr))
    {
        return admissionHr;
    }

    CurlPublicationAccumulator publicationAccumulator;
    auto observeCleanupDebt = wil::scope_exit([&]() noexcept { ObserveCurlCleanupDebt(publicationAccumulator.Snapshot()); });

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FileOperationProgress progress{};
    HRESULT hr = progress.Initialize(FILESYSTEM_COPY, count, options, callback, cookie);
    if (FAILED(hr))
    {
        return hr;
    }

    ResolvedLocation destinationResolved{};
    const HRESULT resolveDestinationHr = ResolveLocation(_protocol, settings, destinationFolder, _hostConnections.get(), true, destinationResolved);
    if (FAILED(resolveDestinationHr))
    {
        return resolveDestinationHr;
    }

    const std::wstring destinationRemoteRoot  = EnsureTrailingSlash(destinationResolved.remotePath);
    const std::wstring destinationDisplayRoot = EnsureTrailingSlashDisplay(BuildDisplayPath(_protocol, destinationFolder));

    const unsigned int requestedConcurrency = std::clamp(settings.copyMoveMaxConcurrency, 1u, 8u);

    struct CopyTask
    {
        unsigned long index = 0;
        ConnectionInfo sourceConn{};
        std::wstring sourceRemotePath;
        std::wstring sourceDisplayPath;
        std::wstring destinationRemotePath;
        std::wstring destinationDisplayPath;
        uint64_t expectedSizeBytes = 0;
        bool expectedSizeKnown     = false;
        bool isDirectory           = false;
    };

    std::vector<CopyTask> tasks;
    tasks.reserve(count);

    const bool continueOnError = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    std::atomic<unsigned long> completedCount{0};
    std::atomic<long> firstFailure{S_OK};
    std::atomic<bool> hadItemFailure{false};

    const auto recordFailure = [&](HRESULT failureHr) noexcept
    {
        long expected = S_OK;
        static_cast<void>(firstFailure.compare_exchange_strong(expected, static_cast<long>(failureHr), std::memory_order_acq_rel));
    };

    for (unsigned long index = 0; index < count; ++index)
    {
        if (! sourcePaths[index] || sourcePaths[index][0] == L'\0')
        {
            hadItemFailure.store(true, std::memory_order_release);
            recordFailure(E_INVALIDARG);
            if (! continueOnError)
            {
                return E_INVALIDARG;
            }
            continue;
        }

        const HRESULT cancelHr = progress.CheckCancel();
        if (FAILED(cancelHr))
        {
            progress.internalCancel.store(true, std::memory_order_release);
            return cancelHr;
        }

        const std::wstring source = NormalizePluginPath(sourcePaths[index]);
        const std::wstring leaf(LeafName(source));

        const std::wstring sourceDisplay     = BuildDisplayPath(_protocol, source);
        const std::wstring destDisplay       = JoinDisplayPath(destinationDisplayRoot, leaf);
        const std::wstring destinationRemote = JoinPluginPath(destinationRemoteRoot, leaf);

        hr = progress.ReportProgress(0, 0, sourceDisplay, destDisplay);
        if (FAILED(hr))
        {
            return hr;
        }

        ResolvedLocation sourceResolved{};
        HRESULT itemHr = ResolveLocation(_protocol, settings, source, _hostConnections.get(), true, sourceResolved);
        FilesInformationCurl::Entry sourceInfo{};
        if (SUCCEEDED(itemHr))
        {
            itemHr = GetEntryInfo(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
        }

        if (FAILED(itemHr))
        {
            hadItemFailure.store(true, std::memory_order_release);
            recordFailure(itemHr);

            const unsigned long done = completedCount.fetch_add(1u, std::memory_order_acq_rel) + 1u;
            progress.SetCompletedItems(done);

            const HRESULT cbHr = progress.ReportItemCompleted(index, sourceDisplay, destDisplay, itemHr);
            if (FAILED(cbHr))
            {
                progress.internalCancel.store(true, std::memory_order_release);
                return cbHr;
            }

            if (! continueOnError)
            {
                progress.internalCancel.store(true, std::memory_order_release);
                return itemHr;
            }

            continue;
        }

        // The selected root, discovered on this sequential pre-pass; a directory root's descendants
        // are added by the walk that later transfers it.
        progress.NoteDiscoveredEntry(sourceInfo);

        CopyTask task{};
        task.index                  = index;
        task.sourceConn             = std::move(sourceResolved.connection);
        task.sourceRemotePath       = std::move(sourceResolved.remotePath);
        task.sourceDisplayPath      = sourceDisplay;
        task.destinationRemotePath  = destinationRemote;
        task.destinationDisplayPath = destDisplay;
        task.expectedSizeBytes      = sourceInfo.sizeBytes;
        task.expectedSizeKnown      = sourceInfo.sizeKnown;
        task.isDirectory            = (sourceInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        tasks.push_back(std::move(task));
    }

    if (tasks.empty())
    {
        const HRESULT failureHr = static_cast<HRESULT>(firstFailure.load(std::memory_order_acquire));
        return FAILED(failureHr) ? failureHr : S_OK;
    }

    std::atomic<uint64_t> overallBytes{0};

    const unsigned long maxWorkers = std::clamp<unsigned long>(static_cast<unsigned long>(requestedConcurrency), 1u, 8u);
    const unsigned long desiredParallelism =
        (std::min)(maxWorkers, static_cast<unsigned long>((std::min)(tasks.size(), static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))));

    const unsigned int concurrency = std::max(1u, static_cast<unsigned int>(desiredParallelism));

    const auto processTask = [&](size_t taskIndex, uint64_t schedulerStreamId) noexcept
    {
        if (taskIndex >= tasks.size())
        {
            return;
        }

        if (progress.internalCancel.load(std::memory_order_acquire))
        {
            return;
        }

        const uint64_t progressStreamId = schedulerStreamId;
        FileOperationProgress::ProgressStreamScope streamScope(progressStreamId);

        const CopyTask& task = tasks[taskIndex];

        HRESULT itemHr = progress.CheckCancel();
        if (SUCCEEDED(itemHr))
        {
            if (task.isDirectory)
            {
                if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
                {
                    itemHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }
                else
                {
                    bool destinationExisted = false;
                    itemHr                  = QueryPathExists(destinationResolved.connection, task.destinationRemotePath, destinationExisted);
                    if (FAILED(itemHr))
                    {
                        // Keep the original failure for this item; no directory copy has started yet.
                    }
                    const unsigned int directoryConcurrency = (concurrency <= 1u) ? requestedConcurrency : 1u;
                    if (SUCCEEDED(itemHr))
                    {
                        itemHr = CopyDirectoryTree(task.sourceConn,
                                                        EnsureTrailingSlash(task.sourceRemotePath),
                                                        EnsureTrailingSlashDisplay(task.sourceDisplayPath),
                                                        destinationResolved.connection,
                                                        EnsureTrailingSlash(task.destinationRemotePath),
                                                        EnsureTrailingSlashDisplay(task.destinationDisplayPath),
                                                        flags,
                                                        directoryConcurrency,
                                                        progress,
                                                        &overallBytes,
                                                        false,
                                                        nullptr,
                                                        &publicationAccumulator);
                        if (FAILED(itemHr) && ! destinationExisted)
                        {
                            const CurlPublicationResult recoveryResult = PreserveCopiedDirectoryAfterFailure(itemHr);
                            publicationAccumulator.Merge(recoveryResult);
                            itemHr = recoveryResult.OperationResult();
                        }
                    }
                }
            }
            else
            {
                itemHr = CopyFileViaTemp(task.sourceConn,
                                         task.sourceRemotePath,
                                         task.sourceDisplayPath,
                                         destinationResolved.connection,
                                         task.destinationRemotePath,
                                         task.destinationDisplayPath,
                                         flags,
                                         progress,
                                         task.expectedSizeBytes,
                                         task.expectedSizeKnown,
                                         false,
                                         false,
                                         &overallBytes,
                                         nullptr,
                                         &publicationAccumulator);
            }
        }

        if (FAILED(itemHr))
        {
            hadItemFailure.store(true, std::memory_order_release);
            recordFailure(itemHr);
            if (! continueOnError || NormalizeCancellation(itemHr) == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                progress.internalCancel.store(true, std::memory_order_release);
            }
        }

        const unsigned long done = completedCount.fetch_add(1u, std::memory_order_acq_rel) + 1u;
        progress.SetCompletedItems(done);

        const HRESULT cbHr = progress.ReportItemCompleted(task.index, task.sourceDisplayPath, task.destinationDisplayPath, itemHr);
        if (FAILED(cbHr))
        {
            recordFailure(cbHr);
            progress.internalCancel.store(true, std::memory_order_release);
            return;
        }

        if (FAILED(itemHr) && ! continueOnError)
        {
            return;
        }
    };

    if (concurrency <= 1u)
    {
        for (size_t i = 0; i < tasks.size(); ++i)
        {
            processTask(i, 0);
            if (progress.internalCancel.load(std::memory_order_acquire))
            {
                break;
            }
        }
    }
    else
    {
        auto job = GetSharedCopyMoveJobScheduler().StartJob(concurrency, tasks.size(), processTask);
        GetSharedCopyMoveJobScheduler().WaitJob(job);
    }

    const HRESULT failureHr = static_cast<HRESULT>(firstFailure.load(std::memory_order_acquire));
    if (progress.internalCancel.load(std::memory_order_acquire))
    {
        return FAILED(failureHr) ? failureHr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (continueOnError && hadItemFailure.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    const HRESULT resultHr = FAILED(failureHr) ? failureHr : S_OK;
    if (SUCCEEDED(resultHr))
    {
        for (unsigned long index = 0; index < count; ++index)
        {
            if (! sourcePaths[index] || sourcePaths[index][0] == L'\0')
            {
                continue;
            }

            const std::wstring source = NormalizePluginPath(sourcePaths[index]);
            NotifySyntheticPathCreated(JoinPluginPath(destinationFolder, LeafName(source)));
        }
    }
    return resultHr;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::MoveItems(const wchar_t* const* sourcePaths,
                                                    unsigned long count,
                                                    const wchar_t* destinationFolder,
                                                    FileSystemFlags flags,
                                                    const FileSystemOptions* options,
                                                    IFileSystemCallback* callback,
                                                    void* cookie) noexcept
{
    const CurlOperationOptionsScope operationOptionsScope(options);
    if (! sourcePaths || ! destinationFolder)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    if (destinationFolder[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (const HRESULT admissionHr = ValidateConditionalMutationAdmission(flags, options); FAILED(admissionHr))
    {
        return admissionHr;
    }

    CurlPublicationAccumulator publicationAccumulator;
    auto observeCleanupDebt = wil::scope_exit([&]() noexcept { ObserveCurlCleanupDebt(publicationAccumulator.Snapshot()); });

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FileOperationProgress progress{};
    HRESULT hr = progress.Initialize(FILESYSTEM_MOVE, count, options, callback, cookie);
    if (FAILED(hr))
    {
        return hr;
    }
    // Keep route admission immutable while workers/callbacks update progress options.
    const bool nativeMoveOnly = progress.options.moveMode == FILESYSTEM_MOVE_NATIVE_ONLY;

    ResolvedLocation destinationResolved{};
    const HRESULT resolveDestinationHr = ResolveLocation(_protocol, settings, destinationFolder, _hostConnections.get(), true, destinationResolved);
    if (FAILED(resolveDestinationHr))
    {
        return resolveDestinationHr;
    }

    const std::wstring destinationRemoteRoot  = EnsureTrailingSlash(destinationResolved.remotePath);
    const std::wstring destinationDisplayRoot = EnsureTrailingSlashDisplay(BuildDisplayPath(_protocol, destinationFolder));
    const bool allowOverwrite                 = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);

    const unsigned int requestedConcurrency = std::clamp(settings.copyMoveMaxConcurrency, 1u, 8u);

    struct MoveTask
    {
        unsigned long index = 0;
        ConnectionInfo sourceConn{};
        std::wstring sourceRemotePath;
        std::wstring sourceDisplayPath;
        std::wstring destinationRemotePath;
        std::wstring destinationDisplayPath;
        uint64_t expectedSizeBytes = 0;
        bool expectedSizeKnown     = false;
        bool isDirectory           = false;
        bool canServerSideRename   = false;
    };

    std::vector<MoveTask> tasks;
    tasks.reserve(count);

    const bool continueOnError = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    std::atomic<unsigned long> completedCount{0};
    std::atomic<long> firstFailure{S_OK};
    std::atomic<bool> hadItemFailure{false};

    const auto recordFailure = [&](HRESULT failureHr) noexcept
    {
        long expected = S_OK;
        static_cast<void>(firstFailure.compare_exchange_strong(expected, static_cast<long>(failureHr), std::memory_order_acq_rel));
    };


    for (unsigned long index = 0; index < count; ++index)
    {
        if (! sourcePaths[index] || sourcePaths[index][0] == L'\0')
        {
            hadItemFailure.store(true, std::memory_order_release);
            recordFailure(E_INVALIDARG);
            if (! continueOnError)
            {
                return E_INVALIDARG;
            }
            continue;
        }

        const std::wstring source = NormalizePluginPath(sourcePaths[index]);
        const std::wstring leaf(LeafName(source));

        const std::wstring sourceDisplay     = BuildDisplayPath(_protocol, source);
        const std::wstring destDisplay       = JoinDisplayPath(destinationDisplayRoot, leaf);
        const std::wstring destinationRemote = JoinPluginPath(destinationRemoteRoot, leaf);

        const HRESULT cancelHr = progress.CheckCancel();
        if (FAILED(cancelHr))
        {
            progress.internalCancel.store(true, std::memory_order_release);
            return cancelHr;
        }

        hr = progress.ReportProgress(0, 0, sourceDisplay, destDisplay);
        if (FAILED(hr))
        {
            return hr;
        }

        ResolvedLocation sourceResolved{};
        HRESULT itemHr = ResolveLocation(_protocol, settings, source, _hostConnections.get(), true, sourceResolved);
        FilesInformationCurl::Entry sourceInfo{};
        bool canServerSideRename = false;
        if (SUCCEEDED(itemHr))
        {
            canServerSideRename = CanServerSideRename(sourceResolved.connection, destinationResolved.connection);
            itemHr              = GetEntryInfo(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
        }

        if (FAILED(itemHr))
        {
            hadItemFailure.store(true, std::memory_order_release);
            recordFailure(itemHr);

            const unsigned long done = completedCount.fetch_add(1u, std::memory_order_acq_rel) + 1u;
            progress.SetCompletedItems(done);

            const HRESULT cbHr = progress.ReportItemCompleted(index, sourceDisplay, destDisplay, itemHr);
            if (FAILED(cbHr))
            {
                progress.internalCancel.store(true, std::memory_order_release);
                return cbHr;
            }

            if (! continueOnError)
            {
                progress.internalCancel.store(true, std::memory_order_release);
                return itemHr;
            }

            continue;
        }

        // The selected root, discovered on this sequential pre-pass. A server-side rename relocates
        // a directory without visiting its descendants; the copy-then-delete fallback's walk adds
        // each descendant it finds.
        progress.NoteDiscoveredEntry(sourceInfo);

        MoveTask task{};
        task.index                  = index;
        task.sourceConn             = std::move(sourceResolved.connection);
        task.sourceRemotePath       = std::move(sourceResolved.remotePath);
        task.sourceDisplayPath      = sourceDisplay;
        task.destinationRemotePath  = destinationRemote;
        task.destinationDisplayPath = destDisplay;
        task.expectedSizeBytes      = sourceInfo.sizeBytes;
        task.expectedSizeKnown      = sourceInfo.sizeKnown;
        task.isDirectory            = (sourceInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        task.canServerSideRename    = canServerSideRename;
        tasks.push_back(std::move(task));
    }

    if (tasks.empty())
    {
        const HRESULT failureHr = static_cast<HRESULT>(firstFailure.load(std::memory_order_acquire));
        return FAILED(failureHr) ? failureHr : S_OK;
    }

    std::atomic<uint64_t> overallBytes{0};

    const unsigned long maxWorkers = std::clamp<unsigned long>(static_cast<unsigned long>(requestedConcurrency), 1u, 8u);
    const unsigned long desiredParallelism =
        (std::min)(maxWorkers, static_cast<unsigned long>((std::min)(tasks.size(), static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))));

    const unsigned int concurrency = std::max(1u, static_cast<unsigned int>(desiredParallelism));

    const auto processTask = [&](size_t taskIndex, uint64_t schedulerStreamId) noexcept
    {
        if (taskIndex >= tasks.size())
        {
            return;
        }

        if (progress.internalCancel.load(std::memory_order_acquire))
        {
            return;
        }

        const uint64_t progressStreamId = schedulerStreamId;
        FileOperationProgress::ProgressStreamScope streamScope(progressStreamId);

        const MoveTask& task               = tasks[taskIndex];
        bool sourceObservedWithoutMutation = true; // Task admission successfully observed this source.

        HRESULT itemHr = progress.CheckCancel();
        if (SUCCEEDED(itemHr))
        {
            if (task.canServerSideRename)
            {
                const bool isSelfRename = IsCaseSensitiveSelfRename(task.sourceRemotePath, task.destinationRemotePath);
                if (! isSelfRename)
                {
                    const CurlPublicationResult renameResult = RenameWithOverwriteRollback(
                        destinationResolved.connection, task.sourceRemotePath, task.destinationRemotePath, allowOverwrite, sourceObservedWithoutMutation);
                    publicationAccumulator.Merge(renameResult);
                    itemHr = renameResult.OperationResult();
                }
                else
                {
                    itemHr = S_OK;
                }
            }
            else if (nativeMoveOnly)
            {
                itemHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
            else if (task.isDirectory)
            {
                if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
                {
                    itemHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }
                else
                {
                    SourceSizeCommitmentMap sourceSizeCommitments;
                    SourceTreePathSet sourceTreeMembers;
                    bool destinationExisted = false;
                    itemHr                  = PreflightDirectorySourceSizes(task.sourceConn,
                                                                            EnsureTrailingSlash(task.sourceRemotePath),
                                                                            progress,
                                                                            sourceSizeCommitments,
                                                                            sourceTreeMembers);
                    if (SUCCEEDED(itemHr))
                    {
                        itemHr = QueryPathExists(destinationResolved.connection, task.destinationRemotePath, destinationExisted);
                    }
                    const unsigned int directoryConcurrency = (concurrency <= 1u) ? requestedConcurrency : 1u;
                    if (SUCCEEDED(itemHr))
                    {
                        sourceObservedWithoutMutation = false;
                        itemHr                        = CopyDirectoryTree(task.sourceConn,
                                                                               EnsureTrailingSlash(task.sourceRemotePath),
                                                                               EnsureTrailingSlashDisplay(task.sourceDisplayPath),
                                                                               destinationResolved.connection,
                                                                               EnsureTrailingSlash(task.destinationRemotePath),
                                                                               EnsureTrailingSlashDisplay(task.destinationDisplayPath),
                                                                               flags,
                                                                               directoryConcurrency,
                                                                               progress,
                                                                               &overallBytes,
                                                                               true,
                                                                               &sourceSizeCommitments,
                                                                               &publicationAccumulator);
                        if (FAILED(itemHr) && ! destinationExisted)
                        {
                            const CurlPublicationResult recoveryResult = PreserveCopiedDirectoryAfterFailure(itemHr);
                            publicationAccumulator.Merge(recoveryResult);
                            itemHr = recoveryResult.OperationResult();
                        }
                    }
                    if (SUCCEEDED(itemHr))
                    {
                        const HRESULT deleteSourceHr = DeleteDirectoryTree(task.sourceConn,
                                                                           task.sourceRemotePath,
                                                                           task.sourceDisplayPath,
                                                                           FILESYSTEM_FLAG_RECURSIVE,
                                                                           ConnectionConcurrencyLimiter::Kind::CopyMove,
                                                                           1u,
                                                                           progress,
                                                                           nullptr,
                                                                           &sourceTreeMembers);
                        if (FAILED(deleteSourceHr))
                        {
                            const CurlPublicationResult sourceDeleteResult = PreserveMovedDirectoryAfterSourceDeleteFailure(deleteSourceHr);
                            publicationAccumulator.Merge(sourceDeleteResult);
                            itemHr = sourceDeleteResult.OperationResult();
                        }
                    }
                }
            }
            else
            {
                std::wstring destinationBackupPath;
                sourceObservedWithoutMutation = false;
                itemHr                        = CopyFileViaTemp(task.sourceConn,
                                                                task.sourceRemotePath,
                                                                task.sourceDisplayPath,
                                                                destinationResolved.connection,
                                                                task.destinationRemotePath,
                                                                task.destinationDisplayPath,
                                                                flags,
                                                                progress,
                                                                task.expectedSizeBytes,
                                                                task.expectedSizeKnown,
                                                                true,
                                                                false,
                                                                &overallBytes,
                                                                &destinationBackupPath,
                                                                &publicationAccumulator);
                if (SUCCEEDED(itemHr))
                {
                    const HRESULT deleteSourceHr =
                        RemoteDeleteFileWithPermit(task.sourceConn, task.sourceRemotePath, progress, ConnectionConcurrencyLimiter::Kind::CopyMove);
                    if (FAILED(deleteSourceHr))
                    {
                        const CurlPublicationResult sourceDeleteResult =
                            PreserveMovedFileDestinationAfterSourceDeleteFailure(deleteSourceHr, destinationBackupPath);
                        publicationAccumulator.Merge(sourceDeleteResult);
                        itemHr = sourceDeleteResult.OperationResult();
                    }
                    else
                    {
                        const CurlPublicationResult cleanupResult = FinalizeOverwriteTarget(destinationResolved.connection, destinationBackupPath);
                        publicationAccumulator.Merge(cleanupResult);
                        itemHr = cleanupResult.OperationResult();
                    }
                }
            }
        }

        if (FAILED(itemHr))
        {
            hadItemFailure.store(true, std::memory_order_release);
            recordFailure(itemHr);
            if (! continueOnError || NormalizeCancellation(itemHr) == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                progress.internalCancel.store(true, std::memory_order_release);
            }
        }

        const unsigned long done = completedCount.fetch_add(1u, std::memory_order_acq_rel) + 1u;
        progress.SetCompletedItems(done);

        const HRESULT cbHr =
            progress.ReportItemCompleted(task.index, task.sourceDisplayPath, task.destinationDisplayPath, itemHr, sourceObservedWithoutMutation);
        if (FAILED(cbHr))
        {
            recordFailure(cbHr);
            progress.internalCancel.store(true, std::memory_order_release);
            return;
        }

        if (FAILED(itemHr) && ! continueOnError)
        {
            return;
        }
    };

    if (concurrency <= 1u)
    {
        for (size_t i = 0; i < tasks.size(); ++i)
        {
            processTask(i, 0);
            if (progress.internalCancel.load(std::memory_order_acquire))
            {
                break;
            }
        }
    }
    else
    {
        auto job = GetSharedCopyMoveJobScheduler().StartJob(concurrency, tasks.size(), processTask);
        GetSharedCopyMoveJobScheduler().WaitJob(job);
    }

    const HRESULT failureHr = static_cast<HRESULT>(firstFailure.load(std::memory_order_acquire));
    if (progress.internalCancel.load(std::memory_order_acquire))
    {
        return FAILED(failureHr) ? failureHr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (continueOnError && hadItemFailure.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    const HRESULT resultHr = FAILED(failureHr) ? failureHr : S_OK;
    if (SUCCEEDED(resultHr))
    {
        for (unsigned long index = 0; index < count; ++index)
        {
            if (! sourcePaths[index] || sourcePaths[index][0] == L'\0')
            {
                continue;
            }

            const std::wstring source = NormalizePluginPath(sourcePaths[index]);
            NotifySyntheticPathMoved(source, JoinPluginPath(destinationFolder, LeafName(source)));
        }
    }
    return resultHr;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::DeleteItems(const wchar_t* const* paths,
                                                      unsigned long count,
                                                      FileSystemFlags flags,
                                                      const FileSystemOptions* options,
                                                      IFileSystemCallback* callback,
                                                      void* cookie) noexcept
{
    const CurlOperationOptionsScope operationOptionsScope(options);
    if (! paths)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    const bool continueOnError = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FileOperationProgress progress{};
    HRESULT hr = progress.Initialize(FILESYSTEM_DELETE, count, options, callback, cookie);
    if (FAILED(hr))
    {
        return hr;
    }


    std::atomic<unsigned long> completedCount{0};
    std::atomic_bool hadFailure{false};
    std::atomic<HRESULT> firstFailure{S_OK};

    auto recordFailure = [&](HRESULT itemHr) noexcept
    {
        itemHr = NormalizeCancellation(itemHr);
        if (FAILED(itemHr) && itemHr != HRESULT_FROM_WIN32(ERROR_CANCELLED))
        {
            hadFailure.store(true, std::memory_order_release);
            HRESULT expected = S_OK;
            static_cast<void>(firstFailure.compare_exchange_strong(expected, itemHr, std::memory_order_acq_rel));
        }
    };

    struct DeleteTask final
    {
        unsigned long index = 0;
        ConnectionInfo connection;
        std::wstring remotePath;
        std::wstring displayPath;
        bool isDirectory = false;
    };

    std::vector<DeleteTask> tasks;
    tasks.reserve(count);

    unsigned int requestedConcurrency = std::clamp(settings.deleteMaxConcurrency, 1u, 8u);

    for (unsigned long index = 0; index < count; ++index)
    {
        const wchar_t* const rawPath = paths[index];
        if (! rawPath || rawPath[0] == L'\0')
        {
            recordFailure(E_INVALIDARG);

            const unsigned long done = completedCount.fetch_add(1u, std::memory_order_acq_rel) + 1u;
            progress.SetCompletedItems(done);

            static_cast<void>(progress.ReportItemCompleted(index, {}, {}, E_INVALIDARG));

            if (! continueOnError)
            {
                return E_INVALIDARG;
            }
            continue;
        }

        const std::wstring displayPath = BuildDisplayPath(_protocol, rawPath);
        hr                             = progress.ReportProgress(0, 0, displayPath, {});
        if (FAILED(hr))
        {
            return hr;
        }

        ResolvedLocation resolved{};
        HRESULT itemHr = ResolveLocation(_protocol, settings, rawPath, _hostConnections.get(), true, resolved);
        if (SUCCEEDED(itemHr))
        {
            FilesInformationCurl::Entry info{};
            itemHr = resolved.remotePath == L"/" ? HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)
                                                 : GetEntryInfo(resolved.connection, resolved.remotePath, info);
            if (SUCCEEDED(itemHr))
            {
                // The selected root; a recursive walk later adds whatever it finds underneath.
                progress.NoteDiscoveredEntry(info);

                DeleteTask task{};
                task.index      = index;
                task.connection = std::move(resolved.connection);
                task.remotePath = std::move(resolved.remotePath);
                task.displayPath.assign(displayPath);
                task.isDirectory = (info.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                tasks.push_back(std::move(task));
            }
        }

        if (FAILED(itemHr))
        {
            recordFailure(itemHr);

            const unsigned long done = completedCount.fetch_add(1u, std::memory_order_acq_rel) + 1u;
            progress.SetCompletedItems(done);

            const HRESULT cbHr = progress.ReportItemCompleted(index, displayPath, {}, itemHr);
            if (FAILED(cbHr))
            {
                return cbHr;
            }

            if (! continueOnError)
            {
                return itemHr;
            }
        }
    }

    if (tasks.empty())
    {
        const HRESULT failureHr = static_cast<HRESULT>(firstFailure.load(std::memory_order_acquire));
        return FAILED(failureHr) ? failureHr : S_OK;
    }

    // Avoid nested parallelism when already parallel across top-level items.
    const unsigned long maxWorkers = std::clamp<unsigned long>(static_cast<unsigned long>(requestedConcurrency), 1u, 8u);
    const unsigned long desiredParallelism =
        (std::min)(maxWorkers, static_cast<unsigned long>((std::min)(tasks.size(), static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))));

    const unsigned int concurrency = std::max(1u, static_cast<unsigned int>(desiredParallelism));

    const unsigned int directoryConcurrency = (concurrency <= 1u) ? requestedConcurrency : 1u;

    const auto processTask = [&](size_t taskIndex, uint64_t schedulerStreamId) noexcept
    {
        if (taskIndex >= tasks.size())
        {
            return;
        }

        if (progress.internalCancel.load(std::memory_order_acquire))
        {
            return;
        }

        const uint64_t progressStreamId = schedulerStreamId;
        FileOperationProgress::ProgressStreamScope streamScope(progressStreamId);

        const DeleteTask& task = tasks[taskIndex];
        std::atomic_bool mutationAttempted{false};

        HRESULT itemHr = progress.CheckCancel();
        if (SUCCEEDED(itemHr))
        {
            itemHr = progress.ReportProgress(0, 0, task.displayPath, {});
        }
        if (SUCCEEDED(itemHr))
        {
            if (task.isDirectory)
            {
                if (HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
                {
                    itemHr = DeleteDirectoryTree(task.connection,
                                                 task.remotePath,
                                                 task.displayPath,
                                                 flags,
                                                 ConnectionConcurrencyLimiter::Kind::Delete,
                                                 directoryConcurrency,
                                                 progress,
                                                 &mutationAttempted);
                }
                else
                {
                    itemHr = RemoteRemoveDirectoryWithPermit(
                        task.connection, task.remotePath, progress, ConnectionConcurrencyLimiter::Kind::Delete, &mutationAttempted);
                }
            }
            else
            {
                itemHr = RemoteDeleteFileWithPermit(task.connection, task.remotePath, progress, ConnectionConcurrencyLimiter::Kind::Delete, &mutationAttempted);
            }
        }

        if (FAILED(itemHr))
        {
            recordFailure(itemHr);
            if (! continueOnError || NormalizeCancellation(itemHr) == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                progress.internalCancel.store(true, std::memory_order_release);
            }
        }

        const unsigned long done = completedCount.fetch_add(1u, std::memory_order_acq_rel) + 1u;
        progress.SetCompletedItems(done);

        const HRESULT cbHr = progress.ReportItemCompleted(task.index, task.displayPath, {}, itemHr, ! mutationAttempted.load(std::memory_order_acquire));
        if (FAILED(cbHr))
        {
            recordFailure(cbHr);
            progress.internalCancel.store(true, std::memory_order_release);
            return;
        }

        if (FAILED(itemHr) && ! continueOnError)
        {
            return;
        }
    };

    if (concurrency <= 1u)
    {
        for (size_t i = 0; i < tasks.size(); ++i)
        {
            processTask(i, 0);
            if (progress.internalCancel.load(std::memory_order_acquire))
            {
                break;
            }
        }
    }
    else
    {
        auto job = GetSharedCopyMoveJobScheduler().StartJob(concurrency, tasks.size(), processTask);
        GetSharedCopyMoveJobScheduler().WaitJob(job);
    }

    if (progress.internalCancel.load(std::memory_order_acquire))
    {
        const HRESULT failureHr = static_cast<HRESULT>(firstFailure.load(std::memory_order_acquire));
        return FAILED(failureHr) ? failureHr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (hadFailure.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    const HRESULT failureHr = static_cast<HRESULT>(firstFailure.load(std::memory_order_acquire));
    const HRESULT resultHr  = FAILED(failureHr) ? failureHr : S_OK;
    if (SUCCEEDED(resultHr))
    {
        for (unsigned long index = 0; index < count; ++index)
        {
            if (paths[index] && paths[index][0] != L'\0')
            {
                NotifySyntheticPathDeleted(paths[index]);
            }
        }
    }
    return resultHr;
}

HRESULT STDMETHODCALLTYPE FileSystemCurl::RenameItems(const FileSystemRenamePair* items,
                                                      unsigned long count,
                                                      FileSystemFlags flags,
                                                      const FileSystemOptions* options,
                                                      IFileSystemCallback* callback,
                                                      void* cookie) noexcept
{
    const CurlOperationOptionsScope operationOptionsScope(options);
    if (! items)
    {
        return E_POINTER;
    }

    if (count == 0)
    {
        return S_OK;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        if (items[index].sizeBytes != sizeof(FileSystemRenamePair))
        {
            return E_INVALIDARG;
        }
    }

    if (const HRESULT admissionHr = ValidateConditionalMutationAdmission(flags, options); FAILED(admissionHr))
    {
        return admissionHr;
    }

    CurlPublicationAccumulator publicationAccumulator;
    auto observeCleanupDebt = wil::scope_exit([&]() noexcept { ObserveCurlCleanupDebt(publicationAccumulator.Snapshot()); });

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    FileOperationProgress progress{};
    HRESULT hr = progress.Initialize(FILESYSTEM_RENAME, count, options, callback, cookie);
    if (FAILED(hr))
    {
        return hr;
    }

    const bool allowOverwrite  = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);
    const bool continueOnError = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    std::atomic<HRESULT> firstFailure{S_OK};
    std::atomic<unsigned long> completedCount{0};
    std::atomic<bool> hadItemFailure{false};

    const auto recordFailure = [&](HRESULT failure) noexcept
    {
        if (SUCCEEDED(failure))
        {
            return;
        }

        HRESULT expected = S_OK;
        static_cast<void>(firstFailure.compare_exchange_strong(expected, failure, std::memory_order_acq_rel));
    };

    const unsigned long maxWorkers         = 4u;
    const unsigned long desiredParallelism = std::min(maxWorkers, count);
    const unsigned int concurrency         = std::max(1u, static_cast<unsigned int>(desiredParallelism));

    const auto processTask = [&](size_t taskIndex, uint64_t schedulerStreamId) noexcept
    {
        if (taskIndex >= count)
        {
            return;
        }

        if (progress.internalCancel.load(std::memory_order_acquire))
        {
            return;
        }

        const unsigned long index = static_cast<unsigned long>(taskIndex);

        const uint64_t progressStreamId = schedulerStreamId;
        FileOperationProgress::ProgressStreamScope streamScope(progressStreamId);

        std::wstring sourceDisplay;
        std::wstring destDisplay;
        bool sourceObservedWithoutMutation = false;

        HRESULT itemHr = progress.CheckCancel();

        const FileSystemRenamePair& pair = items[index];
        if (SUCCEEDED(itemHr))
        {
            if (! pair.sourcePath || ! pair.newName || pair.sourcePath[0] == L'\0' || pair.newName[0] == L'\0')
            {
                itemHr = E_INVALIDARG;
            }
            else
            {
                const std::wstring_view newName = pair.newName;
                if (newName.find_first_of(L"\\/") != std::wstring_view::npos)
                {
                    itemHr = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
                }
                else
                {
                    const std::wstring source = NormalizePluginPath(pair.sourcePath);
                    const std::wstring dest   = JoinPluginPath(ParentPath(source), newName);

                    sourceDisplay = BuildDisplayPath(_protocol, source);
                    destDisplay   = BuildDisplayPath(_protocol, dest);

                    const HRESULT progressHr = progress.ReportProgress(0, 0, sourceDisplay, destDisplay);
                    if (FAILED(progressHr))
                    {
                        recordFailure(progressHr);
                        progress.internalCancel.store(true, std::memory_order_release);
                        return;
                    }

                    ResolvedLocation sourceResolved{};
                    itemHr = ResolveLocation(_protocol, settings, source, _hostConnections.get(), true, sourceResolved);
                    if (SUCCEEDED(itemHr))
                    {
                        FilesInformationCurl::Entry sourceInfo{};
                        itemHr                        = GetEntryInfo(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
                        sourceObservedWithoutMutation = SUCCEEDED(itemHr);
                        if (SUCCEEDED(itemHr))
                        {
                            // A rename relocates the object itself, so a directory discovers one
                            // directory and none of its descendants. Emission is serialized inside
                            // the scope, so the rename workers may report concurrently.
                            progress.NoteDiscoveredEntry(sourceInfo);
                        }
                    }
                    if (SUCCEEDED(itemHr))
                    {
                        ResolvedLocation destinationResolved{};
                        itemHr = ResolveLocation(_protocol, settings, dest, _hostConnections.get(), true, destinationResolved);
                        if (SUCCEEDED(itemHr))
                        {
                            if (! CanServerSideRename(sourceResolved.connection, destinationResolved.connection))
                            {
                                itemHr = HRESULT_FROM_WIN32(ERROR_NOT_SAME_DEVICE);
                            }
                            else
                            {
                                const bool isSelfRename = IsCaseSensitiveSelfRename(sourceResolved.remotePath, destinationResolved.remotePath);
                                if (! isSelfRename)
                                {
                                    const CurlPublicationResult renameResult = RenameWithOverwriteRollback(destinationResolved.connection,
                                                                                                           sourceResolved.remotePath,
                                                                                                           destinationResolved.remotePath,
                                                                                                           allowOverwrite,
                                                                                                           sourceObservedWithoutMutation);
                                    publicationAccumulator.Merge(renameResult);
                                    itemHr = renameResult.OperationResult();
                                }
                                else
                                {
                                    itemHr = S_OK;
                                }
                            }
                        }
                    }
                }
            }
        }

        if (FAILED(itemHr))
        {
            hadItemFailure.store(true, std::memory_order_release);
            recordFailure(itemHr);
            if (! continueOnError || NormalizeCancellation(itemHr) == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                progress.internalCancel.store(true, std::memory_order_release);
            }
        }

        const unsigned long done = completedCount.fetch_add(1u, std::memory_order_acq_rel) + 1u;
        progress.SetCompletedItems(done);

        const HRESULT cbHr = progress.ReportItemCompleted(index, sourceDisplay, destDisplay, itemHr, sourceObservedWithoutMutation);
        if (FAILED(cbHr))
        {
            recordFailure(cbHr);
            progress.internalCancel.store(true, std::memory_order_release);
            return;
        }
    };

    auto job = GetSharedCopyMoveJobScheduler().StartJob(concurrency, count, processTask);
    GetSharedCopyMoveJobScheduler().WaitJob(job);

    const HRESULT failureHr = static_cast<HRESULT>(firstFailure.load(std::memory_order_acquire));
    if (progress.internalCancel.load(std::memory_order_acquire))
    {
        return FAILED(failureHr) ? failureHr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (continueOnError && hadItemFailure.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    const HRESULT resultHr = FAILED(failureHr) ? failureHr : S_OK;
    if (SUCCEEDED(resultHr))
    {
        for (unsigned long index = 0; index < count; ++index)
        {
            if (! items[index].sourcePath || ! items[index].newName)
            {
                continue;
            }

            const std::wstring sourcePath      = NormalizePluginPath(items[index].sourcePath);
            const std::wstring destinationPath = JoinPluginPath(ParentPath(sourcePath), items[index].newName);
            NotifySyntheticPathMoved(sourcePath, destinationPath);
        }
    }
    return resultHr;
}
