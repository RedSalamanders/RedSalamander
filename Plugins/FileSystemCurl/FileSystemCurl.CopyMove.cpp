#include "FileSystemCurl.Internal.h"

#include <atomic>
#include <condition_variable>
#include <functional>
#include <memory>
#include <mutex>
#include <thread>
#include <unordered_map>

using namespace FileSystemCurlInternal;

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

    const int len = static_cast<int>(left.size());
    return CompareStringOrdinal(left.data(), len, right.data(), len, TRUE) == CSTR_EQUAL;
}

class SharedCopyMoveJobScheduler final
{
public:
    SharedCopyMoveJobScheduler() = default;
    ~SharedCopyMoveJobScheduler() noexcept { Shutdown(); }

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
        size_t totalItems          = 0;
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

        constexpr unsigned int kMaxWorkers = 4u;
        workerCount                        = std::max(1u, std::min(workerCount, kMaxWorkers));

        _workers.reserve(workerCount);
        for (unsigned int i = 0; i < workerCount; ++i)
        {
            try
            {
                _workers.emplace_back([this, i](std::stop_token stopToken) noexcept { workerMain(stopToken, static_cast<uint64_t>(i)); });
            }
            catch (const std::system_error&)
            {
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
                    return;
                }

                cleanupJobsLocked();
                if (! tryDequeueWorkLocked(job, index))
                {
                    continue;
                }
            }

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

[[nodiscard]] HRESULT EnsureOverwriteTargetForRename(const ConnectionInfo& conn, std::wstring_view destinationPath, bool allowOverwrite) noexcept
{
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

    return RemoteDeleteFile(conn, destinationPath);
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
                                      std::atomic<uint64_t>* concurrentOverallBytes) noexcept
{
    HRESULT hr = progress.ReportProgress(expectedSizeBytes, 0, sourceFullPath, destinationFullPath);
    if (FAILED(hr))
    {
        return hr;
    }

    const bool allowOverwrite = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);

    hr = EnsureOverwriteTargetFile(destinationConn, destinationRemotePath, allowOverwrite);
    if (FAILED(hr))
    {
        return hr;
    }

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

    hr = CurlDownloadToFile(sourceConn, sourceRemotePath, tempFile.get(), nullptr, &downloadCtx);
    if (FAILED(hr))
    {
        return hr;
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

    hr = CurlUploadFromFile(destinationConn, destinationRemotePath, tempFile.get(), fileSize, nullptr, &uploadCtx);
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
                                             FileOperationProgress& progress,
                                             std::atomic<uint64_t>* concurrentOverallBytes) noexcept
{
    HRESULT hr = EnsureDirectoryExists(destinationConn, destinationRemoteDir);
    if (FAILED(hr))
    {
        return hr;
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
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }

            const std::wstring sourceSubRemote      = EnsureTrailingSlash(sourceChildRemote);
            const std::wstring destinationSubRemote = EnsureTrailingSlash(destinationChildRemote);
            const std::wstring sourceSubFull        = EnsureTrailingSlashDisplay(sourceChildFull);
            const std::wstring destinationSubFull   = EnsureTrailingSlashDisplay(destinationChildFull);

            hr = CopyDirectoryRecursive(
                sourceConn, sourceSubRemote, sourceSubFull, destinationConn, destinationSubRemote, destinationSubFull, flags, progress, concurrentOverallBytes);
            if (FAILED(hr))
            {
                return hr;
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
                return hr;
            }
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT DeleteDirectoryRecursive(const ConnectionInfo& conn,
                                               std::wstring_view directoryRemotePath,
                                               std::wstring_view directoryFullPath,
                                               FileSystemFlags flags,
                                               FileOperationProgress& progress) noexcept
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

            hr = DeleteDirectoryRecursive(conn, childRemote, childFull, flags, progress);
        }
        else
        {
            hr = RemoteDeleteFile(conn, childRemote);
        }

        if (FAILED(hr))
        {
            return hr;
        }
    }

    const std::wstring normalizedRemote = NormalizePluginPath(directoryRemotePath);
    if (normalizedRemote == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }

    return RemoteRemoveDirectory(conn, normalizedRemote);
}
} // namespace

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
            hr = EnsureDirectoryExists(destinationResolved.connection, destinationResolved.remotePath);
            if (SUCCEEDED(hr))
            {
                hr = CopyDirectoryRecursive(sourceResolved.connection,
                                            EnsureTrailingSlash(sourceResolved.remotePath),
                                            EnsureTrailingSlashDisplay(sourceDisplay),
                                            destinationResolved.connection,
                                            EnsureTrailingSlash(destinationResolved.remotePath),
                                            EnsureTrailingSlashDisplay(destinationDisplay),
                                            flags,
                                            progress,
                                            nullptr);
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
    return FAILED(cbHr) ? cbHr : hr;
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
        const bool isSelfRename = EqualsInsensitive(sourceResolved.remotePath, destinationResolved.remotePath);
        if (! isSelfRename)
        {
            hr = EnsureOverwriteTargetForRename(sourceResolved.connection, destinationResolved.remotePath, allowOverwrite);
        }
        if (SUCCEEDED(hr))
        {
            hr = RemoteRename(sourceResolved.connection, sourceResolved.remotePath, destinationResolved.remotePath);
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
                    hr = CopyDirectoryRecursive(sourceResolved.connection,
                                                EnsureTrailingSlash(sourceResolved.remotePath),
                                                EnsureTrailingSlashDisplay(sourceDisplay),
                                                destinationResolved.connection,
                                                EnsureTrailingSlash(destinationResolved.remotePath),
                                                EnsureTrailingSlashDisplay(destinationDisplay),
                                                flags,
                                                progress,
                                                nullptr);
                    if (SUCCEEDED(hr))
                    {
                        hr = DeleteDirectoryRecursive(sourceResolved.connection, sourceResolved.remotePath, sourceDisplay, FILESYSTEM_FLAG_RECURSIVE, progress);
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
                if (SUCCEEDED(hr))
                {
                    hr = RemoteDeleteFile(sourceResolved.connection, sourceResolved.remotePath);
                }
            }
        }
    }

    progress.completedItems = 1;
    const HRESULT cbHr      = progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr);
    return FAILED(cbHr) ? cbHr : hr;
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
                hr = DeleteDirectoryRecursive(resolved.connection, resolved.remotePath, displayPath, flags, progress);
            }
            else
            {
                hr = RemoteRemoveDirectory(resolved.connection, resolved.remotePath);
            }
        }
        else
        {
            hr = RemoteDeleteFile(resolved.connection, resolved.remotePath);
        }
    }

    progress.completedItems = 1;
    const HRESULT cbHr      = progress.ReportItemCompleted(0, displayPath, {}, hr);
    return FAILED(cbHr) ? cbHr : hr;
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
        const bool isSelfRename = EqualsInsensitive(sourceResolved.remotePath, destinationResolved.remotePath);
        if (! isSelfRename)
        {
            hr = EnsureOverwriteTargetForRename(sourceResolved.connection, destinationResolved.remotePath, allowOverwrite);
        }
        if (SUCCEEDED(hr))
        {
            hr = RemoteRename(sourceResolved.connection, sourceResolved.remotePath, destinationResolved.remotePath);
        }
    }

    progress.completedItems = 1;
    const HRESULT cbHr      = progress.ReportItemCompleted(0, sourceDisplay, destinationDisplay, hr);
    return FAILED(cbHr) ? cbHr : hr;
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

    const auto recordFailure = [&](HRESULT failureHr) noexcept
    {
        long expected = S_OK;
        static_cast<void>(firstFailure.compare_exchange_strong(expected, static_cast<long>(failureHr), std::memory_order_acq_rel));
    };

    for (unsigned long index = 0; index < count; ++index)
    {
        if (! sourcePaths[index] || sourcePaths[index][0] == L'\0')
        {
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

    const unsigned long maxWorkers = 4u;
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

        const uint64_t progressStreamId = (concurrency > 0) ? (schedulerStreamId % static_cast<uint64_t>(concurrency)) : 0;
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
                    itemHr = CopyDirectoryRecursive(task.sourceConn,
                                                    EnsureTrailingSlash(task.sourceRemotePath),
                                                    EnsureTrailingSlashDisplay(task.sourceDisplayPath),
                                                    destinationResolved.connection,
                                                    EnsureTrailingSlash(task.destinationRemotePath),
                                                    EnsureTrailingSlashDisplay(task.destinationDisplayPath),
                                                    flags,
                                                    progress,
                                                    &overallBytes);
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
    return FAILED(failureHr) ? failureHr : S_OK;
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

    HRESULT firstFailure = S_OK;
    DirectoryEntryCache entryCache;

    for (unsigned long index = 0; index < count; ++index)
    {
        if (! sourcePaths[index] || sourcePaths[index][0] == L'\0')
        {
            firstFailure = FAILED(firstFailure) ? firstFailure : E_INVALIDARG;
            if (! HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR))
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

        hr = progress.ReportProgress(0, 0, sourceDisplay, destDisplay);
        if (FAILED(hr))
        {
            return hr;
        }

        ResolvedLocation sourceResolved{};
        HRESULT itemHr = ResolveLocation(_protocol, settings, source, _hostConnections.get(), true, sourceResolved);
        if (SUCCEEDED(itemHr))
        {
            if (CanServerSideRename(sourceResolved.connection, destinationResolved.connection))
            {
                const bool isSelfRename = EqualsInsensitive(sourceResolved.remotePath, destinationRemote);
                if (! isSelfRename)
                {
                    itemHr = EnsureOverwriteTargetForRename(destinationResolved.connection, destinationRemote, allowOverwrite);
                }
                if (SUCCEEDED(itemHr))
                {
                    itemHr = RemoteRename(destinationResolved.connection, sourceResolved.remotePath, destinationRemote);
                }
            }
            else
            {
                FilesInformationCurl::Entry sourceInfo{};
                itemHr = entryCache.GetEntryInfoCached(sourceResolved.connection, sourceResolved.remotePath, sourceInfo);
                if (SUCCEEDED(itemHr))
                {
                    if ((sourceInfo.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
                    {
                        if (! HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
                        {
                            itemHr = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                        }
                        else
                        {
                            itemHr = CopyDirectoryRecursive(sourceResolved.connection,
                                                            EnsureTrailingSlash(sourceResolved.remotePath),
                                                            EnsureTrailingSlashDisplay(sourceDisplay),
                                                            destinationResolved.connection,
                                                            EnsureTrailingSlash(destinationRemote),
                                                            EnsureTrailingSlashDisplay(destDisplay),
                                                            flags,
                                                            progress,
                                                            nullptr);
                            if (SUCCEEDED(itemHr))
                            {
                                itemHr = DeleteDirectoryRecursive(
                                    sourceResolved.connection, sourceResolved.remotePath, sourceDisplay, FILESYSTEM_FLAG_RECURSIVE, progress);
                            }
                        }
                    }
                    else
                    {
                        itemHr = CopyFileViaTemp(sourceResolved.connection,
                                                 sourceResolved.remotePath,
                                                 sourceDisplay,
                                                 destinationResolved.connection,
                                                 destinationRemote,
                                                 destDisplay,
                                                 flags,
                                                 progress,
                                                 sourceInfo.sizeBytes,
                                                 nullptr);
                        if (SUCCEEDED(itemHr))
                        {
                            itemHr = RemoteDeleteFile(sourceResolved.connection, sourceResolved.remotePath);
                        }
                    }
                }
            }
        }

        progress.completedItems = index + 1u;
        const HRESULT cbHr      = progress.ReportItemCompleted(index, sourceDisplay, destDisplay, itemHr);
        if (FAILED(cbHr))
        {
            return cbHr;
        }

        if (FAILED(itemHr))
        {
            firstFailure = FAILED(firstFailure) ? firstFailure : itemHr;
            if (! HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR))
            {
                return itemHr;
            }
        }
    }

    return FAILED(firstFailure) ? firstFailure : S_OK;
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

    HRESULT firstFailure = S_OK;
    DirectoryEntryCache entryCache;

    for (unsigned long index = 0; index < count; ++index)
    {
        if (! paths[index] || paths[index][0] == L'\0')
        {
            firstFailure = FAILED(firstFailure) ? firstFailure : E_INVALIDARG;
            if (! HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR))
            {
                return E_INVALIDARG;
            }
            continue;
        }

        const std::wstring displayPath = BuildDisplayPath(_protocol, paths[index]);
        hr                             = progress.ReportProgress(0, 0, displayPath, {});
        if (FAILED(hr))
        {
            return hr;
        }

        ResolvedLocation resolved{};
        HRESULT itemHr = ResolveLocation(_protocol, settings, paths[index], _hostConnections.get(), true, resolved);
        if (SUCCEEDED(itemHr))
        {
            FilesInformationCurl::Entry info{};
            itemHr = entryCache.GetEntryInfoCached(resolved.connection, resolved.remotePath, info);
            if (SUCCEEDED(itemHr))
            {
                if ((info.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
                {
                    if (HasFlag(flags, FILESYSTEM_FLAG_RECURSIVE))
                    {
                        itemHr = DeleteDirectoryRecursive(resolved.connection, resolved.remotePath, displayPath, flags, progress);
                    }
                    else
                    {
                        itemHr = RemoteRemoveDirectory(resolved.connection, resolved.remotePath);
                    }
                }
                else
                {
                    itemHr = RemoteDeleteFile(resolved.connection, resolved.remotePath);
                }
            }
        }

        progress.completedItems = index + 1u;
        const HRESULT cbHr      = progress.ReportItemCompleted(index, displayPath, {}, itemHr);
        if (FAILED(cbHr))
        {
            return cbHr;
        }

        if (FAILED(itemHr))
        {
            firstFailure = FAILED(firstFailure) ? firstFailure : itemHr;
            if (! HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR))
            {
                return itemHr;
            }
        }
    }

    return FAILED(firstFailure) ? firstFailure : S_OK;
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

    const bool allowOverwrite = HasFlag(flags, FILESYSTEM_FLAG_ALLOW_OVERWRITE);
    const bool continueOnError = HasFlag(flags, FILESYSTEM_FLAG_CONTINUE_ON_ERROR);

    std::atomic<HRESULT> firstFailure{S_OK};
    std::atomic<unsigned long> completedCount{0};

    const auto recordFailure = [&](HRESULT failure) noexcept
    {
        if (SUCCEEDED(failure))
        {
            return;
        }

        HRESULT expected = S_OK;
        static_cast<void>(firstFailure.compare_exchange_strong(expected, failure, std::memory_order_acq_rel));
    };

    const unsigned long maxWorkers = 4u;
    const unsigned long desiredParallelism = std::min(maxWorkers, count);
    const unsigned int concurrency          = std::max(1u, static_cast<unsigned int>(desiredParallelism));

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

        const uint64_t progressStreamId = (concurrency > 0) ? (schedulerStreamId % static_cast<uint64_t>(concurrency)) : 0;
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
                                    itemHr = EnsureOverwriteTargetForRename(destinationResolved.connection, destinationResolved.remotePath, allowOverwrite);
                                }
                                if (SUCCEEDED(itemHr))
                                {
                                    itemHr = RemoteRename(destinationResolved.connection, sourceResolved.remotePath, destinationResolved.remotePath);
                                }
                            }
                        }
                    }
                }
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
    return FAILED(failureHr) ? failureHr : S_OK;
}
