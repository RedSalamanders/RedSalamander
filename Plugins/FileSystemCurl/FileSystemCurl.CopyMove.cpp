#include "FileSystemCurl.Internal.h"

#include <bcrypt.h>

#include <atomic>
#include <condition_variable>
#include <deque>
#include <functional>
#include <memory>
#include <mutex>
#include <span>
#include <thread>
#include <unordered_map>

#pragma comment(lib, "bcrypt.lib")

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
[[nodiscard]] bool EqualsInsensitive(std::wstring_view left, std::wstring_view right) noexcept
{
    if (left.size() != right.size())
    {
        return false;
    }

    if (left.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }

    return OrdinalString::EqualsNoCase(left, right);
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

[[nodiscard]] HRESULT DeleteDirectoryRecursive(const ConnectionInfo& conn,
                                               std::wstring_view directoryRemotePath,
                                               std::wstring_view directoryFullPath,
                                               FileSystemFlags flags,
                                               ConnectionConcurrencyLimiter::Kind kind,
                                               FileOperationProgress& progress) noexcept;

[[nodiscard]] HRESULT RemoteDeleteFileWithPermit(const ConnectionInfo& conn,
                                                 std::wstring_view remotePath,
                                                 FileOperationProgress& progress,
                                                 ConnectionConcurrencyLimiter::Kind kind) noexcept
{
    if (conn.limiterKey.empty())
    {
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

    return RemoteDeleteFile(conn, remotePath);
}

[[nodiscard]] HRESULT RemoteRemoveDirectoryWithPermit(const ConnectionInfo& conn,
                                                      std::wstring_view remotePath,
                                                      FileOperationProgress& progress,
                                                      ConnectionConcurrencyLimiter::Kind kind) noexcept
{
    if (conn.limiterKey.empty())
    {
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

    return RemoteRemoveDirectory(conn, remotePath);
}

constexpr size_t kHashMixConstant = 0x9e3779b97f4a7c15ull;

inline void HashCombine(size_t& seed, size_t value) noexcept
{
    seed ^= value + kHashMixConstant + (seed << 6) + (seed >> 2);
}

struct ConnectionCacheKey final
{
    Protocol protocol = Protocol::Sftp;

    std::string host;
    unsigned int port = 0;
    std::string user;
    std::string password;
    std::string basePath;

    bool ftpUseEpsv                  = true;
    unsigned long connectTimeoutMs   = 0;
    unsigned long operationTimeoutMs = 0;
    bool ignoreSslTrust              = false;

    std::string sshPrivateKey;
    std::string sshPublicKey;
    std::string sshKeyPassphrase;
    std::string sshKnownHosts;

    explicit ConnectionCacheKey(const ConnectionInfo& conn)
        : protocol(conn.protocol),
          host(conn.host),
          port(conn.port.value_or(0u)),
          user(conn.user),
          password(conn.password),
          basePath(conn.basePath),
          ftpUseEpsv(conn.ftpUseEpsv),
          connectTimeoutMs(conn.connectTimeoutMs),
          operationTimeoutMs(conn.operationTimeoutMs),
          ignoreSslTrust(conn.ignoreSslTrust),
          sshPrivateKey(conn.sshPrivateKey),
          sshPublicKey(conn.sshPublicKey),
          sshKeyPassphrase(conn.sshKeyPassphrase),
          sshKnownHosts(conn.sshKnownHosts)
    {
    }

    bool operator==(const ConnectionCacheKey&) const noexcept = default;
};

struct ConnectionCacheKeyHash final
{
    size_t operator()(const ConnectionCacheKey& key) const noexcept
    {
        size_t h = 0;
        HashCombine(h, std::hash<int>{}(static_cast<int>(key.protocol)));
        HashCombine(h, std::hash<std::string>{}(key.host));
        HashCombine(h, std::hash<unsigned int>{}(key.port));
        HashCombine(h, std::hash<std::string>{}(key.user));
        HashCombine(h, std::hash<std::string>{}(key.password));
        HashCombine(h, std::hash<std::string>{}(key.basePath));
        HashCombine(h, std::hash<bool>{}(key.ftpUseEpsv));
        HashCombine(h, std::hash<unsigned long>{}(key.connectTimeoutMs));
        HashCombine(h, std::hash<unsigned long>{}(key.operationTimeoutMs));
        HashCombine(h, std::hash<bool>{}(key.ignoreSslTrust));
        HashCombine(h, std::hash<std::string>{}(key.sshPrivateKey));
        HashCombine(h, std::hash<std::string>{}(key.sshPublicKey));
        HashCombine(h, std::hash<std::string>{}(key.sshKeyPassphrase));
        HashCombine(h, std::hash<std::string>{}(key.sshKnownHosts));
        return h;
    }
};

class DirectoryEntryCache final
{
public:
    [[nodiscard]] HRESULT GetEntryInfoCached(const ConnectionInfo& conn, std::wstring_view path, FilesInformationCurl::Entry& out) noexcept
    {
        const std::wstring normalized = NormalizePluginPath(path);
        if (normalized == L"/")
        {
            out            = {};
            out.attributes = FILE_ATTRIBUTE_DIRECTORY;
            out.name       = L"/";
            return S_OK;
        }

        const std::wstring parent    = ParentPath(normalized);
        const std::wstring_view leaf = LeafName(normalized);

        auto& byDirectory = _cache[ConnectionCacheKey(conn)];
        auto foundDir     = byDirectory.find(parent);
        if (foundDir == byDirectory.end())
        {
            std::vector<FilesInformationCurl::Entry> entries;
            const HRESULT hr = ReadDirectoryEntries(conn, parent, entries);
            if (FAILED(hr))
            {
                return hr;
            }
            foundDir = byDirectory.emplace(parent, std::move(entries)).first;
        }

        const auto found = FindEntryByName(foundDir->second, leaf);
        if (! found.has_value())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        out = found.value();
        return S_OK;
    }

private:
    using DirectoryMap = std::unordered_map<std::wstring, std::vector<FilesInformationCurl::Entry>>;
    std::unordered_map<ConnectionCacheKey, DirectoryMap, ConnectionCacheKeyHash> _cache;
};

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

[[nodiscard]] HRESULT GenerateRandomBytes(std::span<std::byte> bytes) noexcept
{
    if (bytes.empty())
    {
        return S_OK;
    }

    const NTSTATUS status = BCryptGenRandom(nullptr, reinterpret_cast<PUCHAR>(bytes.data()), static_cast<ULONG>(bytes.size()), BCRYPT_USE_SYSTEM_PREFERRED_RNG);
    return BCRYPT_SUCCESS(status) ? S_OK : HRESULT_FROM_NT(status);
}

void AppendHexToken(std::wstring& value, std::span<const std::byte> bytes)
{
    constexpr wchar_t kHex[] = L"0123456789abcdef";
    value.reserve(value.size() + (bytes.size() * 2u));
    for (const std::byte byte : bytes)
    {
        const unsigned int v = std::to_integer<unsigned int>(byte);
        value.push_back(kHex[(v >> 4u) & 0x0Fu]);
        value.push_back(kHex[v & 0x0Fu]);
    }
}

[[nodiscard]] HRESULT BuildRemoteSiblingLeaf(std::wstring_view purposeTag, std::wstring& leafOut) noexcept
{
    leafOut.clear();

    std::array<std::byte, 16> entropy{};
    HRESULT hr = GenerateRandomBytes(entropy);
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring leaf;
    leaf.reserve(std::wstring_view(L".redsalamander-").size() + purposeTag.size() + 1u + (entropy.size() * 2u));
    leaf.append(L".redsalamander-");
    leaf.append(purposeTag);
    leaf.push_back(L'-');
    AppendHexToken(leaf, entropy);

    leafOut = std::move(leaf);
    return S_OK;
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

#if defined(_DEBUG)
bool DebugCheck(bool condition, const wchar_t* message, unsigned int& passed, unsigned int& failed) noexcept
{
    if (condition)
    {
        ++passed;
        return true;
    }

    ++failed;
    Debug::Error(L"FileSystemCurl debug selftest failed: {}", message);
    return false;
}

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
#endif

[[nodiscard]] HRESULT PrepareOverwriteTargetForRename(const ConnectionInfo& conn,
                                                      std::wstring_view destinationPath,
                                                      bool allowOverwrite,
                                                      std::wstring& backupPathOut) noexcept
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

    hr = RemoteRename(conn, destinationPath, backupPathOut);
    if (FAILED(hr))
    {
        backupPathOut.clear();
        return hr;
    }

    return S_OK;
}

[[nodiscard]] HRESULT RestoreOverwriteTargetAfterFailure(const ConnectionInfo& conn, std::wstring_view destinationPath, std::wstring_view backupPath) noexcept
{
    if (backupPath.empty())
    {
        return S_OK;
    }

    FilesInformationCurl::Entry existing{};
    const HRESULT existsHr = GetEntryInfo(conn, destinationPath, existing);
    if (SUCCEEDED(existsHr))
    {
        if ((existing.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        const HRESULT deleteHr = RemoteDeleteFile(conn, destinationPath);
        if (FAILED(deleteHr))
        {
            return deleteHr;
        }
    }
    else if (existsHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
    {
        return existsHr;
    }

    return RemoteRename(conn, backupPath, destinationPath);
}

[[nodiscard]] HRESULT FinalizeOverwriteTarget(const ConnectionInfo& conn, std::wstring_view backupPath) noexcept
{
    if (backupPath.empty())
    {
        return S_OK;
    }

    return RemoteDeleteFile(conn, backupPath);
}

[[nodiscard]] HRESULT RenameWithOverwriteRollback(const ConnectionInfo& conn,
                                                  std::wstring_view sourcePath,
                                                  std::wstring_view destinationPath,
                                                  bool allowOverwrite) noexcept
{
    std::wstring backupPath;
    HRESULT hr = PrepareOverwriteTargetForRename(conn, destinationPath, allowOverwrite, backupPath);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = RemoteRename(conn, sourcePath, destinationPath);
    if (FAILED(hr))
    {
        const HRESULT restoreHr = RestoreOverwriteTargetAfterFailure(conn, destinationPath, backupPath);
        return FAILED(restoreHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : hr;
    }

    return FinalizeOverwriteTarget(conn, backupPath);
}

[[nodiscard]] HRESULT RollbackMovedFileDestination(const ConnectionInfo& conn,
                                                   std::wstring_view destinationPath,
                                                   std::wstring_view backupPath,
                                                   FileOperationProgress& progress,
                                                   ConnectionConcurrencyLimiter::Kind kind) noexcept
{
    const HRESULT deleteHr = RemoteDeleteFileWithPermit(conn, destinationPath, progress, kind);
    if (FAILED(deleteHr) && deleteHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
    {
        return deleteHr;
    }

    if (backupPath.empty())
    {
        return S_OK;
    }

    return RemoteRename(conn, backupPath, destinationPath);
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

[[nodiscard]] HRESULT TryRollbackCopiedDirectory(const ConnectionInfo& destinationConn,
                                                 std::wstring_view destinationRemotePath,
                                                 std::wstring_view destinationDisplayPath,
                                                 FileOperationProgress& progress) noexcept
{
    return DeleteDirectoryRecursive(
        destinationConn, destinationRemotePath, destinationDisplayPath, FILESYSTEM_FLAG_RECURSIVE, ConnectionConcurrencyLimiter::Kind::CopyMove, progress);
}

[[nodiscard]] HRESULT PromoteStagedFileToDestination(const ConnectionInfo& destinationConn,
                                                     std::wstring_view stagedRemotePath,
                                                     std::wstring_view destinationRemotePath,
                                                     bool allowOverwrite,
                                                     std::wstring* backupPathOut) noexcept
{
    std::wstring backupPath;
    HRESULT hr = PrepareOverwriteTargetForRename(destinationConn, destinationRemotePath, allowOverwrite, backupPath);
    if (FAILED(hr))
    {
        static_cast<void>(RemoteDeleteFile(destinationConn, stagedRemotePath));
        return hr;
    }

    hr = RemoteRename(destinationConn, stagedRemotePath, destinationRemotePath);
    if (FAILED(hr))
    {
        static_cast<void>(RemoteDeleteFile(destinationConn, stagedRemotePath));
        const HRESULT restoreHr = RestoreOverwriteTargetAfterFailure(destinationConn, destinationRemotePath, backupPath);
        return FAILED(restoreHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : hr;
    }

    if (backupPathOut)
    {
        *backupPathOut = std::move(backupPath);
        return S_OK;
    }

    return FinalizeOverwriteTarget(destinationConn, backupPath);
}

[[nodiscard]] HRESULT CopyFileViaTemp(const ConnectionInfo& sourceConn,
                                      std::wstring_view sourceRemotePath,
                                      std::wstring_view sourceFullPath,
                                      const ConnectionInfo& destinationConn,
                                      std::wstring_view destinationRemotePath,
                                      std::wstring_view destinationFullPath,
                                      FileSystemFlags flags,
                                      FileOperationProgress& progress,
                                      uint64_t expectedSizeBytes,
                                      std::atomic<uint64_t>* concurrentOverallBytes,
                                      std::wstring* backupPathOut = nullptr) noexcept
{
    HRESULT hr = progress.ReportProgress(expectedSizeBytes, 0, sourceFullPath, destinationFullPath);
    if (FAILED(hr))
    {
        return hr;
    }

    const bool allowOverwrite = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);

    hr = progress.CheckCancel();
    if (FAILED(hr))
    {
        return hr;
    }

    wil::unique_hfile tempFile = CreateTemporaryDeleteOnCloseFile();
    if (! tempFile)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const uint64_t baseCompleted = concurrentOverallBytes ? 0 : progress.completedBytes;

    TransferProgressContext downloadCtx{};
    downloadCtx.progress               = &progress;
    downloadCtx.sourcePath             = sourceFullPath;
    downloadCtx.destinationPath        = destinationFullPath;
    downloadCtx.baseCompletedBytes     = baseCompleted;
    downloadCtx.concurrentOverallBytes = concurrentOverallBytes;
    downloadCtx.itemTotalBytes         = expectedSizeBytes;
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
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    hr = CurlDownloadToFile(sourceConn, sourceRemotePath, tempFile.get(), nullptr, &downloadCtx);
    if (FAILED(hr))
    {
        return hr;
    }

    // Same-connection copies acquire download then upload permits sequentially.
    // Release the download permit once the temp file is materialized so a
    // single-slot limiter does not deadlock when source and destination share
    // the same remote connection.
    if (! sourceConn.limiterKey.empty() && sourceConn.limiterKey == destinationConn.limiterKey)
    {
        downloadPermit = {};
    }

    uint64_t fileSize = 0;
    hr                = GetFileSizeBytes(tempFile.get(), fileSize);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = ResetFilePointerToStart(tempFile.get());
    if (FAILED(hr))
    {
        return hr;
    }

    hr = progress.CheckCancel();
    if (FAILED(hr))
    {
        return hr;
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
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }

    std::wstring stagedRemotePath;
    hr = GenerateRemoteSiblingPath(destinationConn, destinationRemotePath, L"upload", stagedRemotePath);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CurlUploadFromFile(destinationConn, stagedRemotePath, tempFile.get(), fileSize, nullptr, &uploadCtx);
    if (FAILED(hr))
    {
        static_cast<void>(RemoteDeleteFile(destinationConn, stagedRemotePath));
        return hr;
    }

    FilesInformationCurl::Entry stagedInfo{};
    hr = GetEntryInfo(destinationConn, stagedRemotePath, stagedInfo);
    if (FAILED(hr))
    {
        static_cast<void>(RemoteDeleteFile(destinationConn, stagedRemotePath));
        return hr;
    }
    // Verify the staged upload landed as a non-directory of the expected size. Some FTP listing
    // dialects cannot report a file size (the parser leaves sizeKnown == false); for those we fall
    // back to existence verification instead of failing every non-empty upload. This stays data-safe:
    // CurlUploadFromFile only returns success after libcurl transmits all fileSize bytes.
    if ((stagedInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0 || (stagedInfo.sizeKnown && stagedInfo.sizeBytes != fileSize))
    {
        static_cast<void>(RemoteDeleteFile(destinationConn, stagedRemotePath));
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    hr = PromoteStagedFileToDestination(destinationConn, stagedRemotePath, destinationRemotePath, allowOverwrite, backupPathOut);
    if (FAILED(hr))
    {
        return hr;
    }

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

[[nodiscard]] HRESULT CopyDirectoryRecursive(const ConnectionInfo& sourceConn,
                                             std::wstring_view sourceRemoteDir,
                                             std::wstring_view sourceFullDir,
                                             const ConnectionInfo& destinationConn,
                                             std::wstring_view destinationRemoteDir,
                                             std::wstring_view destinationFullDir,
                                             FileSystemFlags flags,
                                             unsigned int maxConcurrency,
                                             FileOperationProgress& progress,
                                             std::atomic<uint64_t>* concurrentOverallBytes) noexcept
{
    const bool continueOnError = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);
    bool hadFailure            = false;

    HRESULT hr = EnsureDirectoryExists(destinationConn, destinationRemoteDir);
    if (FAILED(hr))
    {
        return hr;
    }

    const unsigned int requestedConcurrency = std::clamp(maxConcurrency, 1u, 8u);
    if (requestedConcurrency > 1u && progress.callback != nullptr && GetSharedCopyMoveJobScheduler().EnsureWorkersAvailable())
    {
        struct CopyFileWorkItem final
        {
            std::wstring sourceRemotePath;
            std::wstring sourceDisplayPath;
            std::wstring destinationRemotePath;
            std::wstring destinationDisplayPath;
            uint64_t expectedSizeBytes = 0;
        };

        struct CopyFileQueue final
        {
            CopyFileQueue()                                = default;
            CopyFileQueue(const CopyFileQueue&)            = delete;
            CopyFileQueue(CopyFileQueue&&)                 = delete;
            CopyFileQueue& operator=(const CopyFileQueue&) = delete;
            CopyFileQueue& operator=(CopyFileQueue&&)      = delete;

            std::mutex mutex;
            std::condition_variable notEmptyCv;
            std::condition_variable notFullCv;
            std::deque<CopyFileWorkItem> items;
            bool enumerationDone = false;
        };

        CopyFileQueue queue{};
        const size_t maxQueuedItems = std::max<size_t>(256u, static_cast<size_t>(requestedConcurrency) * 32u);

        std::atomic<bool> cancelRequested{false};
        std::atomic<bool> stopRequested{false};
        std::atomic<long> firstFailure{S_OK};
        std::atomic<bool> hadAnyFailure{false};

        // Ensure progress.completedBytes is only touched under the callback lock by using the concurrentOverallBytes path.
        std::atomic<uint64_t> localOverallBytes{0};
        if (! concurrentOverallBytes)
        {
            concurrentOverallBytes = &localOverallBytes;
        }

        const auto recordFailure = [&](HRESULT failureHr) noexcept
        {
            if (SUCCEEDED(failureHr))
            {
                return;
            }

            long expected = S_OK;
            static_cast<void>(firstFailure.compare_exchange_strong(expected, static_cast<long>(failureHr), std::memory_order_acq_rel));
        };

        const auto requestCancel = [&]() noexcept
        {
            cancelRequested.store(true, std::memory_order_release);
            queue.notEmptyCv.notify_all();
            queue.notFullCv.notify_all();
        };

        const auto requestStop = [&](HRESULT failureHr) noexcept
        {
            recordFailure(failureHr);
            stopRequested.store(true, std::memory_order_release);
            queue.notEmptyCv.notify_all();
            queue.notFullCv.notify_all();
        };

        const auto enqueueFile = [&](CopyFileWorkItem item) noexcept -> bool
        {
            for (;;)
            {
                if (cancelRequested.load(std::memory_order_acquire) || stopRequested.load(std::memory_order_acquire))
                {
                    return false;
                }

                std::unique_lock lock(queue.mutex);
                if (queue.items.size() < maxQueuedItems)
                {
                    queue.items.push_back(std::move(item));
                    lock.unlock();
                    queue.notEmptyCv.notify_one();
                    return true;
                }

                queue.notFullCv.wait(lock, [&]() noexcept {
                    return cancelRequested.load(std::memory_order_acquire) || stopRequested.load(std::memory_order_acquire) ||
                           queue.items.size() < maxQueuedItems;
                });
            }
        };

        const unsigned int concurrency = std::max(2u, requestedConcurrency);

        auto job = GetSharedCopyMoveJobScheduler().StartJob(concurrency,
                                                            concurrency,
                                                            [&](size_t /*index*/, uint64_t schedulerStreamId) noexcept
        {
            const uint64_t progressStreamId = schedulerStreamId;
            FileOperationProgress::ProgressStreamScope streamScope(progressStreamId);

            for (;;)
            {
                if (cancelRequested.load(std::memory_order_acquire) || stopRequested.load(std::memory_order_acquire))
                {
                    return;
                }

                CopyFileWorkItem item{};
                {
                    std::unique_lock lock(queue.mutex);
                    queue.notEmptyCv.wait(lock, [&]() noexcept {
                        return cancelRequested.load(std::memory_order_acquire) || stopRequested.load(std::memory_order_acquire) || ! queue.items.empty() ||
                               queue.enumerationDone;
                    });

                    if (cancelRequested.load(std::memory_order_acquire) || stopRequested.load(std::memory_order_acquire))
                    {
                        return;
                    }

                    if (queue.items.empty())
                    {
                        return;
                    }

                    item = std::move(queue.items.front());
                    queue.items.pop_front();
                }

                queue.notFullCv.notify_one();

                HRESULT itemHr = CopyFileViaTemp(sourceConn,
                                                 item.sourceRemotePath,
                                                 item.sourceDisplayPath,
                                                 destinationConn,
                                                 item.destinationRemotePath,
                                                 item.destinationDisplayPath,
                                                 flags,
                                                 progress,
                                                 item.expectedSizeBytes,
                                                 concurrentOverallBytes);
                if (FAILED(itemHr))
                {
                    itemHr = NormalizeCancellation(itemHr);
                    if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
                    {
                        requestCancel();
                    }
                    else if (IsAuthenticationFailureHr(itemHr))
                    {
                        requestStop(itemHr);
                    }
                    else if (! continueOnError)
                    {
                        requestStop(itemHr);
                    }
                    else
                    {
                        hadAnyFailure.store(true, std::memory_order_release);
                    }
                }

                queue.notFullCv.notify_all();
            }
        });

        const auto produceDirectory = [&](auto&& self,
                                          std::wstring_view currentSourceRemoteDir,
                                          std::wstring_view currentSourceFullDir,
                                          std::wstring_view currentDestinationRemoteDir,
                                          std::wstring_view currentDestinationFullDir,
                                          bool isRootDir) noexcept -> HRESULT
        {
            if (cancelRequested.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }

            if (stopRequested.load(std::memory_order_acquire))
            {
                return S_OK;
            }

            if (! isRootDir)
            {
                const HRESULT ensureHr = EnsureDirectoryExists(destinationConn, currentDestinationRemoteDir);
                if (FAILED(ensureHr))
                {
                    if (! continueOnError)
                    {
                        requestStop(ensureHr);
                        return ensureHr;
                    }

                    hadAnyFailure.store(true, std::memory_order_release);
                    return S_OK;
                }
            }

            std::vector<FilesInformationCurl::Entry> entries;
            HRESULT hr = ReadDirectoryEntries(sourceConn, currentSourceRemoteDir, entries);
            if (FAILED(hr))
            {
                if (! continueOnError || isRootDir)
                {
                    requestStop(hr);
                    return hr;
                }

                hadAnyFailure.store(true, std::memory_order_release);
                return S_OK;
            }

            uint64_t cancelCheckCounter = 0;
            for (const auto& entry : entries)
            {
                if (cancelRequested.load(std::memory_order_acquire))
                {
                    return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                }
                if (stopRequested.load(std::memory_order_acquire))
                {
                    return S_OK;
                }

                if (IsDotOrDotDotName(entry.name))
                {
                    continue;
                }

                if ((++cancelCheckCounter % 64u) == 0u)
                {
                    const HRESULT cancelHr = progress.CheckCancel();
                    if (FAILED(cancelHr))
                    {
                        hr = NormalizeCancellation(cancelHr);
                        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
                        {
                            requestCancel();
                        }
                        else
                        {
                            requestStop(hr);
                        }
                        return hr;
                    }
                }

                const std::wstring sourceChildRemote      = JoinPluginPath(currentSourceRemoteDir, entry.name);
                const std::wstring destinationChildRemote = JoinPluginPath(currentDestinationRemoteDir, entry.name);
                const std::wstring sourceChildFull        = JoinDisplayPath(currentSourceFullDir, entry.name);
                const std::wstring destinationChildFull   = JoinDisplayPath(currentDestinationFullDir, entry.name);

                if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
                {
                    if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
                    {
                        if (continueOnError)
                        {
                            hadAnyFailure.store(true, std::memory_order_release);
                            continue;
                        }

                        const HRESULT notSupported = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                        requestStop(notSupported);
                        return notSupported;
                    }

                    const std::wstring sourceSubRemote      = EnsureTrailingSlash(sourceChildRemote);
                    const std::wstring destinationSubRemote = EnsureTrailingSlash(destinationChildRemote);
                    const std::wstring sourceSubFull        = EnsureTrailingSlashDisplay(sourceChildFull);
                    const std::wstring destinationSubFull   = EnsureTrailingSlashDisplay(destinationChildFull);

                    const HRESULT childHr = self(self, sourceSubRemote, sourceSubFull, destinationSubRemote, destinationSubFull, false);
                    if (FAILED(childHr))
                    {
                        hr = NormalizeCancellation(childHr);
                        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || IsAuthenticationFailureHr(hr) || ! continueOnError)
                        {
                            return hr;
                        }

                        hadAnyFailure.store(true, std::memory_order_release);
                        continue;
                    }
                }
                else
                {
                    CopyFileWorkItem task{};
                    task.sourceRemotePath       = sourceChildRemote;
                    task.destinationRemotePath  = destinationChildRemote;
                    task.sourceDisplayPath      = sourceChildFull;
                    task.destinationDisplayPath = destinationChildFull;
                    task.expectedSizeBytes      = entry.sizeBytes;

                    if (! enqueueFile(std::move(task)))
                    {
                        break;
                    }
                }
            }

            return S_OK;
        };

        const HRESULT rootHr = produceDirectory(produceDirectory, sourceRemoteDir, sourceFullDir, destinationRemoteDir, destinationFullDir, true);
        if (FAILED(rootHr))
        {
            const HRESULT normalized = NormalizeCancellation(rootHr);
            if (normalized == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                requestCancel();
            }
            else if (IsAuthenticationFailureHr(normalized) || ! continueOnError)
            {
                requestStop(normalized);
            }
            else
            {
                hadAnyFailure.store(true, std::memory_order_release);
            }
        }

        {
            std::scoped_lock lock(queue.mutex);
            queue.enumerationDone = true;
        }
        queue.notEmptyCv.notify_all();
        queue.notFullCv.notify_all();

        GetSharedCopyMoveJobScheduler().WaitJob(job);

        if (cancelRequested.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (stopRequested.load(std::memory_order_acquire))
        {
            const HRESULT failureHr = static_cast<HRESULT>(firstFailure.load(std::memory_order_acquire));
            return FAILED(failureHr) ? failureHr : HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
        }

        if (hadAnyFailure.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        return S_OK;
    }

    std::vector<FilesInformationCurl::Entry> entries;
    hr = ReadDirectoryEntries(sourceConn, sourceRemoteDir, entries);
    if (FAILED(hr))
    {
        return hr;
    }

    for (const auto& entry : entries)
    {
        if (IsDotOrDotDotName(entry.name))
        {
            continue;
        }

        hr = progress.CheckCancel();
        if (FAILED(hr))
        {
            return hr;
        }

        const std::wstring sourceChildRemote      = JoinPluginPath(sourceRemoteDir, entry.name);
        const std::wstring destinationChildRemote = JoinPluginPath(destinationRemoteDir, entry.name);
        const std::wstring sourceChildFull        = JoinDisplayPath(sourceFullDir, entry.name);
        const std::wstring destinationChildFull   = JoinDisplayPath(destinationFullDir, entry.name);

        if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
            {
                if (! continueOnError)
                {
                    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }

                hadFailure = true;
                continue;
            }

            const std::wstring sourceSubRemote      = EnsureTrailingSlash(sourceChildRemote);
            const std::wstring destinationSubRemote = EnsureTrailingSlash(destinationChildRemote);
            const std::wstring sourceSubFull        = EnsureTrailingSlashDisplay(sourceChildFull);
            const std::wstring destinationSubFull   = EnsureTrailingSlashDisplay(destinationChildFull);

            hr = CopyDirectoryRecursive(sourceConn,
                                        sourceSubRemote,
                                        sourceSubFull,
                                        destinationConn,
                                        destinationSubRemote,
                                        destinationSubFull,
                                        flags,
                                        1u,
                                        progress,
                                        concurrentOverallBytes);
            if (FAILED(hr))
            {
                hr = NormalizeCancellation(hr);
                if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || IsAuthenticationFailureHr(hr) || ! continueOnError)
                {
                    return hr;
                }

                hadFailure = true;
                continue;
            }
        }
        else
        {
            hr = CopyFileViaTemp(sourceConn,
                                 sourceChildRemote,
                                 sourceChildFull,
                                 destinationConn,
                                 destinationChildRemote,
                                 destinationChildFull,
                                 flags,
                                 progress,
                                 entry.sizeBytes,
                                 concurrentOverallBytes);
            if (FAILED(hr))
            {
                hr = NormalizeCancellation(hr);
                if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || IsAuthenticationFailureHr(hr) || ! continueOnError)
                {
                    return hr;
                }

                hadFailure = true;
                continue;
            }
        }
    }

    return hadFailure ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK;
}

struct DeleteTreeWorkItem final
{
    std::wstring remotePath;
    std::wstring displayPath;
};

[[nodiscard]] HRESULT CollectDeleteDirectoryWork(const ConnectionInfo& conn,
                                                 std::wstring_view directoryRemotePath,
                                                 std::wstring_view directoryFullPath,
                                                 FileSystemFlags flags,
                                                 FileOperationProgress& progress,
                                                 std::vector<DeleteTreeWorkItem>& outFiles,
                                                 std::vector<DeleteTreeWorkItem>& outDirectoriesPostOrder) noexcept
{
    const std::wstring directoryRemote = EnsureTrailingSlash(directoryRemotePath);
    const std::wstring directoryFull   = EnsureTrailingSlashDisplay(directoryFullPath);

    std::vector<FilesInformationCurl::Entry> entries;
    HRESULT hr = ReadDirectoryEntries(conn, directoryRemote, entries);
    if (FAILED(hr))
    {
        return hr;
    }

    for (const auto& entry : entries)
    {
        if (IsDotOrDotDotName(entry.name))
        {
            continue;
        }

        hr = progress.CheckCancel();
        if (FAILED(hr))
        {
            return hr;
        }

        const std::wstring childRemote = JoinPluginPath(directoryRemote, entry.name);
        const std::wstring childFull   = JoinDisplayPath(directoryFull, entry.name);

        if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }

            hr = CollectDeleteDirectoryWork(conn, childRemote, childFull, flags, progress, outFiles, outDirectoriesPostOrder);
            if (FAILED(hr))
            {
                return hr;
            }
        }
        else
        {
            outFiles.push_back(DeleteTreeWorkItem{childRemote, childFull});
        }
    }

    const std::wstring normalizedRemote = NormalizePluginPath(directoryRemotePath);
    if (normalizedRemote == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    outDirectoriesPostOrder.push_back(DeleteTreeWorkItem{normalizedRemote, std::wstring(directoryFullPath)});
    return S_OK;
}

[[nodiscard]] HRESULT DeleteDirectoryRecursiveParallel(const ConnectionInfo& conn,
                                                       std::wstring_view directoryRemotePath,
                                                       std::wstring_view directoryFullPath,
                                                       FileSystemFlags flags,
                                                       ConnectionConcurrencyLimiter::Kind kind,
                                                       unsigned int requestedConcurrency,
                                                       FileOperationProgress& progress) noexcept
{
    const bool continueOnError = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    std::vector<DeleteTreeWorkItem> files;
    std::vector<DeleteTreeWorkItem> directoriesPostOrder;

    HRESULT hr = CollectDeleteDirectoryWork(conn, directoryRemotePath, directoryFullPath, flags, progress, files, directoriesPostOrder);
    if (FAILED(hr))
    {
        return hr;
    }

    std::atomic_bool internalCancel{false};
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

    const unsigned long maxWorkers = std::clamp<unsigned long>(static_cast<unsigned long>(requestedConcurrency), 1u, 8u);
    const unsigned long desiredParallelism =
        (std::min)(maxWorkers, static_cast<unsigned long>((std::min)(files.size(), static_cast<size_t>((std::numeric_limits<unsigned long>::max)()))));

    const unsigned int concurrency = std::max(1u, static_cast<unsigned int>(desiredParallelism));

    const auto processFile = [&](size_t fileIndex, uint64_t schedulerStreamId) noexcept
    {
        if (fileIndex >= files.size())
        {
            return;
        }

        if (internalCancel.load(std::memory_order_acquire))
        {
            return;
        }

        const uint64_t progressStreamId = schedulerStreamId;
        FileOperationProgress::ProgressStreamScope streamScope(progressStreamId);

        const DeleteTreeWorkItem& item = files[fileIndex];

        HRESULT itemHr = progress.CheckCancel();
        if (SUCCEEDED(itemHr))
        {
            itemHr = progress.ReportProgress(0, 0, item.displayPath, {});
        }
        if (SUCCEEDED(itemHr))
        {
            itemHr = RemoteDeleteFileWithPermit(conn, item.remotePath, progress, kind);
        }

        if (FAILED(itemHr))
        {
            recordFailure(itemHr);
            if (! continueOnError || NormalizeCancellation(itemHr) == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                internalCancel.store(true, std::memory_order_release);
                progress.internalCancel.store(true, std::memory_order_release);
            }
        }
    };

    if (concurrency <= 1u)
    {
        for (size_t i = 0; i < files.size(); ++i)
        {
            processFile(i, 0);
            if (internalCancel.load(std::memory_order_acquire))
            {
                break;
            }
        }
    }
    else
    {
        auto job = GetSharedCopyMoveJobScheduler().StartJob(concurrency, files.size(), processFile);
        GetSharedCopyMoveJobScheduler().WaitJob(job);
    }

    if (internalCancel.load(std::memory_order_acquire))
    {
        const HRESULT failureHr = firstFailure.load(std::memory_order_acquire);
        return FAILED(failureHr) ? failureHr : HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    for (const DeleteTreeWorkItem& directory : directoriesPostOrder)
    {
        hr = progress.CheckCancel();
        if (FAILED(hr))
        {
            return hr;
        }

        hr = progress.ReportProgress(0, 0, directory.displayPath, {});
        if (FAILED(hr))
        {
            return hr;
        }

        hr = RemoteRemoveDirectoryWithPermit(conn, directory.remotePath, progress, kind);
        if (FAILED(hr))
        {
            recordFailure(hr);
            if (! continueOnError || NormalizeCancellation(hr) == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                return NormalizeCancellation(hr);
            }
        }
    }

    if (hadFailure.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    }

    const HRESULT failureHr = firstFailure.load(std::memory_order_acquire);
    return FAILED(failureHr) ? failureHr : S_OK;
}

[[nodiscard]] HRESULT DeleteDirectoryRecursive(const ConnectionInfo& conn,
                                               std::wstring_view directoryRemotePath,
                                               std::wstring_view directoryFullPath,
                                               FileSystemFlags flags,
                                               ConnectionConcurrencyLimiter::Kind kind,
                                               FileOperationProgress& progress) noexcept
{
    const bool continueOnError = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    const std::wstring directoryRemote = EnsureTrailingSlash(directoryRemotePath);
    const std::wstring directoryFull   = EnsureTrailingSlashDisplay(directoryFullPath);

    std::vector<FilesInformationCurl::Entry> entries;
    HRESULT hr = ReadDirectoryEntries(conn, directoryRemote, entries);
    if (FAILED(hr))
    {
        return hr;
    }

    bool hadFailure = false;

    for (const auto& entry : entries)
    {
        if (IsDotOrDotDotName(entry.name))
        {
            continue;
        }

        hr = progress.CheckCancel();
        if (FAILED(hr))
        {
            return hr;
        }

        const std::wstring childRemote = JoinPluginPath(directoryRemote, entry.name);
        const std::wstring childFull   = JoinDisplayPath(directoryFull, entry.name);
        hr                             = progress.ReportProgress(0, 0, childFull, {});
        if (FAILED(hr))
        {
            return hr;
        }

        if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }

            hr = DeleteDirectoryRecursive(conn, childRemote, childFull, flags, kind, progress);
        }
        else
        {
            hr = RemoteDeleteFileWithPermit(conn, childRemote, progress, kind);
        }

        if (FAILED(hr))
        {
            hr = NormalizeCancellation(hr);
            if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || IsAuthenticationFailureHr(hr) || ! continueOnError)
            {
                return hr;
            }

            hadFailure = true;
            continue;
        }
    }

    const std::wstring normalizedRemote = NormalizePluginPath(directoryRemotePath);
    if (normalizedRemote == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    hr = RemoteRemoveDirectoryWithPermit(conn, normalizedRemote, progress, kind);
    if (FAILED(hr))
    {
        hr = NormalizeCancellation(hr);
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || IsAuthenticationFailureHr(hr) || ! continueOnError)
        {
            return hr;
        }

        hadFailure = true;
    }

    return hadFailure ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK;
}
} // namespace

#if defined(_DEBUG)
extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderCurlDebugSelfTests(unsigned int* passed, unsigned int* failed)
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }

    *passed = 0;
    *failed = 0;

    RunDebugRemoteSiblingLeafEntropySelfTest(*passed, *failed);

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
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

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
                hr = CopyDirectoryRecursive(sourceResolved.connection,
                                            EnsureTrailingSlash(sourceResolved.remotePath),
                                            EnsureTrailingSlashDisplay(sourceDisplay),
                                            destinationResolved.connection,
                                            EnsureTrailingSlash(destinationResolved.remotePath),
                                            EnsureTrailingSlashDisplay(destinationDisplay),
                                            flags,
                                            requestedConcurrency,
                                            progress,
                                            nullptr);
                if (FAILED(hr) && ! destinationExisted)
                {
                    const HRESULT rollbackHr =
                        TryRollbackCopiedDirectory(destinationResolved.connection, destinationResolved.remotePath, destinationDisplay, progress);
                    if (FAILED(rollbackHr))
                    {
                        hr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    }
                }
            }
        }
    }
    else
    {
        hr = CopyFileViaTemp(sourceResolved.connection,
                             sourceResolved.remotePath,
                             sourceDisplay,
                             destinationResolved.connection,
                             destinationResolved.remotePath,
                             destinationDisplay,
                             flags,
                             progress,
                             sourceInfo.sizeBytes,
                             nullptr);
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
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

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

    if (CanServerSideRename(sourceResolved.connection, destinationResolved.connection))
    {
        FilesInformationCurl::Entry sourceInfo{};
        hr                      = GetEntryInfo(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
        const bool isSelfRename = EqualsInsensitive(sourceResolved.remotePath, destinationResolved.remotePath);
        if (SUCCEEDED(hr) && ! isSelfRename)
        {
            hr = RenameWithOverwriteRollback(sourceResolved.connection, sourceResolved.remotePath, destinationResolved.remotePath, allowOverwrite);
        }
        else if (SUCCEEDED(hr))
        {
            hr = S_OK;
        }
    }
    else
    {
        FilesInformationCurl::Entry sourceInfo{};
        hr = GetEntryInfo(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
        if (SUCCEEDED(hr))
        {
            if ((sourceInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
            {
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
                        hr = CopyDirectoryRecursive(sourceResolved.connection,
                                                    EnsureTrailingSlash(sourceResolved.remotePath),
                                                    EnsureTrailingSlashDisplay(sourceDisplay),
                                                    destinationResolved.connection,
                                                    EnsureTrailingSlash(destinationResolved.remotePath),
                                                    EnsureTrailingSlashDisplay(destinationDisplay),
                                                    flags,
                                                    requestedConcurrency,
                                                    progress,
                                                    nullptr);
                        if (FAILED(hr) && ! destinationExisted)
                        {
                            const HRESULT rollbackHr =
                                TryRollbackCopiedDirectory(destinationResolved.connection, destinationResolved.remotePath, destinationDisplay, progress);
                            if (FAILED(rollbackHr))
                            {
                                hr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                            }
                        }
                    }
                    if (SUCCEEDED(hr))
                    {
                        const HRESULT deleteSourceHr = DeleteDirectoryRecursive(sourceResolved.connection,
                                                                                sourceResolved.remotePath,
                                                                                sourceDisplay,
                                                                                FILESYSTEM_FLAG_RECURSIVE,
                                                                                ConnectionConcurrencyLimiter::Kind::CopyMove,
                                                                                progress);
                        if (FAILED(deleteSourceHr))
                        {
                            if (! destinationExisted)
                            {
                                const HRESULT rollbackHr =
                                    TryRollbackCopiedDirectory(destinationResolved.connection, destinationResolved.remotePath, destinationDisplay, progress);
                                hr = FAILED(rollbackHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : deleteSourceHr;
                            }
                            else
                            {
                                hr = deleteSourceHr;
                            }
                        }
                    }
                }
            }
            else
            {
                std::wstring destinationBackupPath;
                hr = CopyFileViaTemp(sourceResolved.connection,
                                     sourceResolved.remotePath,
                                     sourceDisplay,
                                     destinationResolved.connection,
                                     destinationResolved.remotePath,
                                     destinationDisplay,
                                     flags,
                                     progress,
                                     sourceInfo.sizeBytes,
                                     nullptr,
                                     &destinationBackupPath);
                if (SUCCEEDED(hr))
                {
                    const HRESULT deleteSourceHr = RemoteDeleteFileWithPermit(
                        sourceResolved.connection, sourceResolved.remotePath, progress, ConnectionConcurrencyLimiter::Kind::CopyMove);
                    if (FAILED(deleteSourceHr))
                    {
                        const HRESULT rollbackHr = RollbackMovedFileDestination(destinationResolved.connection,
                                                                                destinationResolved.remotePath,
                                                                                destinationBackupPath,
                                                                                progress,
                                                                                ConnectionConcurrencyLimiter::Kind::CopyMove);
                        hr                       = FAILED(rollbackHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : deleteSourceHr;
                    }
                    else
                    {
                        hr = FinalizeOverwriteTarget(destinationResolved.connection, destinationBackupPath);
                    }
                }
            }
        }
    }

    progress.completedItems = 1;
    const HRESULT cbHr      = progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr);
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
    hr = GetEntryInfo(resolved.connection, resolved.remotePath, info);
    if (SUCCEEDED(hr))
    {
        if ((info.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            if (HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
            {
                const unsigned int concurrency = std::clamp(resolved.connection.effectiveDeleteMaxConcurrency, 1u, 8u);
                hr = (concurrency > 1u)
                         ? DeleteDirectoryRecursiveParallel(
                               resolved.connection, resolved.remotePath, displayPath, flags, ConnectionConcurrencyLimiter::Kind::Delete, concurrency, progress)
                         : DeleteDirectoryRecursive(
                               resolved.connection, resolved.remotePath, displayPath, flags, ConnectionConcurrencyLimiter::Kind::Delete, progress);
            }
            else
            {
                hr = RemoteRemoveDirectoryWithPermit(resolved.connection, resolved.remotePath, progress, ConnectionConcurrencyLimiter::Kind::Delete);
            }
        }
        else
        {
            hr = RemoteDeleteFileWithPermit(resolved.connection, resolved.remotePath, progress, ConnectionConcurrencyLimiter::Kind::Delete);
        }
    }

    progress.completedItems = 1;
    const HRESULT cbHr      = progress.ReportItemCompleted(0, displayPath, {}, hr);
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
    if (! sourcePath || ! destinationPath)
    {
        return E_POINTER;
    }

    if (sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

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

    if (! CanServerSideRename(sourceResolved.connection, destinationResolved.connection))
    {
        hr = HRESULT_FROM_WIN32(ERROR_NOT_SAME_DEVICE);
    }
    else
    {
        FilesInformationCurl::Entry sourceInfo{};
        hr                      = GetEntryInfo(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
        const bool isSelfRename = EqualsInsensitive(sourceResolved.remotePath, destinationResolved.remotePath);
        if (SUCCEEDED(hr) && ! isSelfRename)
        {
            hr = RenameWithOverwriteRollback(sourceResolved.connection, sourceResolved.remotePath, destinationResolved.remotePath, allowOverwrite);
        }
        else if (SUCCEEDED(hr))
        {
            hr = S_OK;
        }
    }

    progress.completedItems = 1;
    const HRESULT cbHr      = progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr);
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

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    ResolvedLocation destinationResolved{};
    const HRESULT resolveDestinationHr = ResolveLocation(_protocol, settings, destinationFolder, _hostConnections.get(), true, destinationResolved);
    if (FAILED(resolveDestinationHr))
    {
        return resolveDestinationHr;
    }

    FileOperationProgress progress{};
    HRESULT hr = progress.Initialize(FILESYSTEM_COPY, count, options, callback, cookie);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring destinationRemoteRoot  = EnsureTrailingSlash(destinationResolved.remotePath);
    const std::wstring destinationDisplayRoot = EnsureTrailingSlashDisplay(BuildDisplayPath(_protocol, destinationFolder));
    DirectoryEntryCache entryCache;

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
            itemHr = entryCache.GetEntryInfoCached(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
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

        CopyTask task{};
        task.index                  = index;
        task.sourceConn             = std::move(sourceResolved.connection);
        task.sourceRemotePath       = std::move(sourceResolved.remotePath);
        task.sourceDisplayPath      = sourceDisplay;
        task.destinationRemotePath  = destinationRemote;
        task.destinationDisplayPath = destDisplay;
        task.expectedSizeBytes      = sourceInfo.sizeBytes;
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
                        itemHr = CopyDirectoryRecursive(task.sourceConn,
                                                        EnsureTrailingSlash(task.sourceRemotePath),
                                                        EnsureTrailingSlashDisplay(task.sourceDisplayPath),
                                                        destinationResolved.connection,
                                                        EnsureTrailingSlash(task.destinationRemotePath),
                                                        EnsureTrailingSlashDisplay(task.destinationDisplayPath),
                                                        flags,
                                                        directoryConcurrency,
                                                        progress,
                                                        &overallBytes);
                        if (FAILED(itemHr) && ! destinationExisted)
                        {
                            const HRESULT rollbackHr =
                                TryRollbackCopiedDirectory(destinationResolved.connection, task.destinationRemotePath, task.destinationDisplayPath, progress);
                            if (FAILED(rollbackHr))
                            {
                                itemHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                            }
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
                                         &overallBytes);
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

    Settings settings;
    {
        std::lock_guard lock(_stateMutex);
        settings = _settings;
    }

    ResolvedLocation destinationResolved{};
    const HRESULT resolveDestinationHr = ResolveLocation(_protocol, settings, destinationFolder, _hostConnections.get(), true, destinationResolved);
    if (FAILED(resolveDestinationHr))
    {
        return resolveDestinationHr;
    }

    FileOperationProgress progress{};
    HRESULT hr = progress.Initialize(FILESYSTEM_MOVE, count, options, callback, cookie);
    if (FAILED(hr))
    {
        return hr;
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

    DirectoryEntryCache entryCache;

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
            itemHr              = entryCache.GetEntryInfoCached(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
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

        MoveTask task{};
        task.index                  = index;
        task.sourceConn             = std::move(sourceResolved.connection);
        task.sourceRemotePath       = std::move(sourceResolved.remotePath);
        task.sourceDisplayPath      = sourceDisplay;
        task.destinationRemotePath  = destinationRemote;
        task.destinationDisplayPath = destDisplay;
        task.expectedSizeBytes      = sourceInfo.sizeBytes;
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

        const MoveTask& task = tasks[taskIndex];

        HRESULT itemHr = progress.CheckCancel();
        if (SUCCEEDED(itemHr))
        {
            if (task.canServerSideRename)
            {
                const bool isSelfRename = EqualsInsensitive(task.sourceRemotePath, task.destinationRemotePath);
                if (! isSelfRename)
                {
                    itemHr = RenameWithOverwriteRollback(destinationResolved.connection, task.sourceRemotePath, task.destinationRemotePath, allowOverwrite);
                }
                else
                {
                    itemHr = S_OK;
                }
            }
            else if (task.isDirectory)
            {
                if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
                {
                    itemHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                }
                else
                {
                    bool destinationExisted                 = false;
                    itemHr                                  = QueryPathExists(destinationResolved.connection, task.destinationRemotePath, destinationExisted);
                    const unsigned int directoryConcurrency = (concurrency <= 1u) ? requestedConcurrency : 1u;
                    if (SUCCEEDED(itemHr))
                    {
                        itemHr = CopyDirectoryRecursive(task.sourceConn,
                                                        EnsureTrailingSlash(task.sourceRemotePath),
                                                        EnsureTrailingSlashDisplay(task.sourceDisplayPath),
                                                        destinationResolved.connection,
                                                        EnsureTrailingSlash(task.destinationRemotePath),
                                                        EnsureTrailingSlashDisplay(task.destinationDisplayPath),
                                                        flags,
                                                        directoryConcurrency,
                                                        progress,
                                                        &overallBytes);
                        if (FAILED(itemHr) && ! destinationExisted)
                        {
                            const HRESULT rollbackHr =
                                TryRollbackCopiedDirectory(destinationResolved.connection, task.destinationRemotePath, task.destinationDisplayPath, progress);
                            if (FAILED(rollbackHr))
                            {
                                itemHr = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                            }
                        }
                    }
                    if (SUCCEEDED(itemHr))
                    {
                        const HRESULT deleteSourceHr = DeleteDirectoryRecursive(task.sourceConn,
                                                                                task.sourceRemotePath,
                                                                                task.sourceDisplayPath,
                                                                                FILESYSTEM_FLAG_RECURSIVE,
                                                                                ConnectionConcurrencyLimiter::Kind::CopyMove,
                                                                                progress);
                        if (FAILED(deleteSourceHr))
                        {
                            if (! destinationExisted)
                            {
                                const HRESULT rollbackHr = TryRollbackCopiedDirectory(
                                    destinationResolved.connection, task.destinationRemotePath, task.destinationDisplayPath, progress);
                                itemHr = FAILED(rollbackHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : deleteSourceHr;
                            }
                            else
                            {
                                itemHr = deleteSourceHr;
                            }
                        }
                    }
                }
            }
            else
            {
                std::wstring destinationBackupPath;
                itemHr = CopyFileViaTemp(task.sourceConn,
                                         task.sourceRemotePath,
                                         task.sourceDisplayPath,
                                         destinationResolved.connection,
                                         task.destinationRemotePath,
                                         task.destinationDisplayPath,
                                         flags,
                                         progress,
                                         task.expectedSizeBytes,
                                         &overallBytes,
                                         &destinationBackupPath);
                if (SUCCEEDED(itemHr))
                {
                    const HRESULT deleteSourceHr =
                        RemoteDeleteFileWithPermit(task.sourceConn, task.sourceRemotePath, progress, ConnectionConcurrencyLimiter::Kind::CopyMove);
                    if (FAILED(deleteSourceHr))
                    {
                        const HRESULT rollbackHr = RollbackMovedFileDestination(destinationResolved.connection,
                                                                                task.destinationRemotePath,
                                                                                destinationBackupPath,
                                                                                progress,
                                                                                ConnectionConcurrencyLimiter::Kind::CopyMove);
                        itemHr                   = FAILED(rollbackHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : deleteSourceHr;
                    }
                    else
                    {
                        itemHr = FinalizeOverwriteTarget(destinationResolved.connection, destinationBackupPath);
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

    DirectoryEntryCache entryCache;

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
            itemHr = entryCache.GetEntryInfoCached(resolved.connection, resolved.remotePath, info);
            if (SUCCEEDED(itemHr))
            {
                requestedConcurrency = std::min(requestedConcurrency, std::clamp(resolved.connection.effectiveDeleteMaxConcurrency, 1u, 8u));

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
                    itemHr = directoryConcurrency > 1u
                                 ? DeleteDirectoryRecursiveParallel(task.connection,
                                                                    task.remotePath,
                                                                    task.displayPath,
                                                                    flags,
                                                                    ConnectionConcurrencyLimiter::Kind::Delete,
                                                                    directoryConcurrency,
                                                                    progress)
                                 : DeleteDirectoryRecursive(
                                       task.connection, task.remotePath, task.displayPath, flags, ConnectionConcurrencyLimiter::Kind::Delete, progress);
                }
                else
                {
                    itemHr = RemoteRemoveDirectoryWithPermit(task.connection, task.remotePath, progress, ConnectionConcurrencyLimiter::Kind::Delete);
                }
            }
            else
            {
                itemHr = RemoteDeleteFileWithPermit(task.connection, task.remotePath, progress, ConnectionConcurrencyLimiter::Kind::Delete);
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

        const HRESULT cbHr = progress.ReportItemCompleted(task.index, task.displayPath, {}, itemHr);
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
    DirectoryEntryCache entryCache;

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
                        itemHr = entryCache.GetEntryInfoCached(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
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
                                const bool isSelfRename = EqualsInsensitive(sourceResolved.remotePath, destinationResolved.remotePath);
                                if (! isSelfRename)
                                {
                                    itemHr = RenameWithOverwriteRollback(
                                        destinationResolved.connection, sourceResolved.remotePath, destinationResolved.remotePath, allowOverwrite);
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

        const HRESULT cbHr = progress.ReportItemCompleted(index, sourceDisplay, destDisplay, itemHr);
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
