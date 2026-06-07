#include "FileSystem.Internal.h"
#include "LocalSearchIndexCore.h"
#include "SearchServiceBroker.h"
#include "SearchTextHelpers.h"

#include <algorithm>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <deque>
#include <limits>
#include <list>
#include <mutex>
#include <regex>
#include <string_view>
#include <thread>
#include <unordered_set>
#include <vector>

using namespace FileSystemInternal;

namespace
{
constexpr uint64_t kDefaultSearchContentBytesPerFile         = SearchTextHelpers::kDefaultContentBytesPerFile;
constexpr uint32_t kDefaultSearchSnippetChars                = SearchTextHelpers::kDefaultSnippetCharacters;
constexpr uint64_t kProgressIntervalItems                    = 128u;
constexpr ULONGLONG kProgressIntervalMs                      = 200u;
constexpr ULONGLONG kSearchServiceUnavailableRetryCooldownMs = 5000u;
constexpr HRESULT kCancelledHr                               = HRESULT_FROM_WIN32(ERROR_CANCELLED);
constexpr HRESULT kFileTooLargeHr                            = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
constexpr HRESULT kAccessDeniedHr                            = HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
constexpr size_t kMaxRegexPatternLength                      = 1000u;
constexpr size_t kMaxRegexGroupDepth                         = 20u;

// Parallel scan: directories with at least this many entries are evaluated in parallel.
// Below this threshold, single-threaded evaluation is used to avoid threadpool overhead.
constexpr unsigned long kParallelScanThreshold = 500u;
// Number of entries per chunk dispatched to a threadpool worker.
constexpr unsigned long kParallelScanChunkSize        = 128u;
constexpr size_t kParallelDirectoryResultFlushMatches = 32u;

// ---------------------------------------------------------------------------
// Regex compilation cache (bounded LRU)
// ---------------------------------------------------------------------------
// Caches compiled std::wregex objects to avoid recompilation on repeated
// search patterns. Thread-safe via mutex. The cache is global to the module
// (shared across FileSystem instances) for wider hit rate.
class CompiledRegexCache final
{
public:
    static constexpr size_t kMaxEntries = 10;

    CompiledRegexCache()                                     = default;
    CompiledRegexCache(const CompiledRegexCache&)            = delete;
    CompiledRegexCache(CompiledRegexCache&&)                 = delete;
    CompiledRegexCache& operator=(const CompiledRegexCache&) = delete;
    CompiledRegexCache& operator=(CompiledRegexCache&&)      = delete;

    // Returns a cached regex or compiles and caches a new one.
    // Throws std::regex_error if the pattern is invalid (caller must handle).
    [[nodiscard]] std::shared_ptr<const std::wregex> GetOrCompile(const std::wstring& pattern, std::regex_constants::syntax_option_type flags)
    {
        std::lock_guard lock(_mutex);

        // Search for existing entry (front = most recently used).
        for (auto it = _entries.begin(); it != _entries.end(); ++it)
        {
            if (it->flags == flags && it->pattern == pattern)
            {
                // Move to front (MRU).
                if (it != _entries.begin())
                {
                    _entries.splice(_entries.begin(), _entries, it);
                }
                return it->compiled;
            }
        }

        // Compile new regex (may throw std::regex_error).
        auto compiled = std::make_shared<const std::wregex>(pattern, flags);

        // Insert at front, evict LRU if at capacity.
        if (_entries.size() >= kMaxEntries)
        {
            _entries.pop_back();
        }
        _entries.push_front({pattern, flags, compiled});
        return compiled;
    }

private:
    struct Entry
    {
        std::wstring pattern;
        std::regex_constants::syntax_option_type flags{};
        std::shared_ptr<const std::wregex> compiled;
    };

    std::mutex _mutex;
    std::list<Entry> _entries;
};

CompiledRegexCache g_regexCache;

struct SearchEntryMetadata final
{
    std::wstring fullPath;
    std::wstring relativePath;
    std::wstring displayName;
    unsigned long fileAttributes = 0;
    __int64 creationTime         = 0;
    __int64 lastAccessTime       = 0;
    __int64 lastWriteTime        = 0;
    __int64 changeTime           = 0;
    __int64 endOfFile            = 0;
    __int64 allocationSize       = 0;
};

struct DirectoryVisitIdentity final
{
    DWORD volumeSerialNumber = 0u;
    uint64_t fileIndex       = 0u;

    [[nodiscard]] bool operator==(const DirectoryVisitIdentity& other) const noexcept = default;
};

struct DirectoryVisitIdentityHasher final
{
    [[nodiscard]] size_t operator()(const DirectoryVisitIdentity& identity) const noexcept
    {
        const uint64_t mixed = (static_cast<uint64_t>(identity.volumeSerialNumber) << 32u) ^ identity.fileIndex;
        return static_cast<size_t>(mixed ^ (mixed >> 32u));
    }
};

struct SearchContentResult final
{
    bool matched        = false;
    uint64_t byteOffset = 0;
    uint32_t byteLength = 0;
    std::wstring previewText;
};

struct SearchBackendSelection final
{
    FileSystemSearchBackend backend = FILESYSTEM_SEARCH_BACKEND_SCAN;
    uint32_t warningFlags           = FILESYSTEM_SEARCH_WARNING_NONE;
};

struct SearchRuntime final
{
    SearchRuntime()                                = default;
    SearchRuntime(const SearchRuntime&)            = delete;
    SearchRuntime(SearchRuntime&&)                 = delete;
    SearchRuntime& operator=(const SearchRuntime&) = delete;
    SearchRuntime& operator=(SearchRuntime&&)      = delete;

    FileSystem* fileSystem                               = nullptr;
    const FileSystemSearchQuery* query                   = nullptr;
    IFileSystemSearchCallback* callback                  = nullptr;
    const FileSystemSearchHostExtensions* hostExtensions = nullptr;
    void* cookie                                         = nullptr;
    std::wstring rootPath;
    std::wstring namePattern;
    std::wstring contentPattern;
    std::shared_ptr<const std::wregex> nameRegex;
    std::shared_ptr<const std::wregex> contentRegex;
    std::unordered_set<std::wstring> queuedDirectories;
    std::unordered_set<DirectoryVisitIdentity, DirectoryVisitIdentityHasher> queuedDirectoryIdentities;
    std::unordered_set<std::wstring> emittedMatchPaths;
    std::mutex queuedDirectoriesMutex;
    bool includeFiles                         = false;
    bool includeDirectories                   = false;
    bool recursive                            = false;
    bool followSymlinks                       = false;
    bool wantSnippets                         = false;
    bool matchCaseName                        = false;
    bool matchCaseContent                     = false;
    uint64_t maxResults                       = 0;
    uint64_t maxContentBytesPerFile           = kDefaultSearchContentBytesPerFile;
    uint32_t maxSnippetCharacters             = kDefaultSearchSnippetChars;
    uint64_t scannedDirectories               = 0;
    uint64_t scannedFiles                     = 0;
    uint64_t candidateFiles                   = 0;
    uint64_t indexedCandidatesConsumed        = 0;
    uint64_t matchedEntries                   = 0;
    FileSystemSearchBackend backend           = FILESYSTEM_SEARCH_BACKEND_SCAN;
    uint32_t warningFlags                     = FILESYSTEM_SEARCH_WARNING_NONE;
    unsigned int parallelDirectoryWalkWorkers = 4u;
    ULONGLONG lastProgressTick                = 0;
    uint64_t lastProgressItems                = 0;
    FileSystemSearchPhase lastProgressPhase   = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    bool hasReportedProgress                  = false;
    bool stopRequested                        = false;
    bool usingIndexedEnumeration              = false;
};

[[nodiscard]] const wchar_t* SearchBackendToString(FileSystemSearchBackend backend) noexcept
{
    switch (backend)
    {
        case FILESYSTEM_SEARCH_BACKEND_SCAN: return L"scan";
        case FILESYSTEM_SEARCH_BACKEND_INDEX: return L"local-index";
        case FILESYSTEM_SEARCH_BACKEND_SERVICE: return L"service";
        case FILESYSTEM_SEARCH_BACKEND_UNKNOWN:
        default: return L"unknown";
    }
}

[[nodiscard]] const wchar_t* BackendPreferenceToString(FileSystemSearchBackendPreference preference) noexcept
{
    switch (preference)
    {
        case FileSystemSearchBackendPreference::Auto: return L"auto";
        case FileSystemSearchBackendPreference::Service: return L"service";
        case FileSystemSearchBackendPreference::LocalIndex: return L"local-index";
        case FileSystemSearchBackendPreference::Scan:
        default: return L"scan";
    }
}

[[nodiscard]] unsigned long ByteCountOfString(std::wstring_view text) noexcept
{
    const size_t bytes = text.size() * sizeof(wchar_t);
    if (bytes > (std::numeric_limits<unsigned long>::max)())
    {
        return (std::numeric_limits<unsigned long>::max)();
    }

    return static_cast<unsigned long>(bytes);
}

[[nodiscard]] std::wstring ToCaseFolded(std::wstring_view text) noexcept
{
    return OrdinalString::FoldCaseInvariant(text);
}

[[nodiscard]] std::wstring NormalizeSearchPath(std::wstring_view path) noexcept
{
    std::wstring normalized(path);
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');

    while (normalized.size() > 1u && (normalized.back() == L'\\' || normalized.back() == L'/'))
    {
        const bool keepDriveRoot = normalized.size() == 3u && normalized[1] == L':' && (normalized[2] == L'\\' || normalized[2] == L'/');
        const bool keepShareRoot = normalized.size() == 2u && normalized[0] == L'\\' && normalized[1] == L'\\';
        const bool keepExtendedDriveRoot =
            normalized.size() == 7u && normalized.rfind(L"\\\\?\\", 0) == 0 && normalized[5] == L':' && (normalized[6] == L'\\' || normalized[6] == L'/');
        if (keepDriveRoot || keepShareRoot || keepExtendedDriveRoot)
        {
            break;
        }

        normalized.pop_back();
    }

    return normalized;
}

[[nodiscard]] std::wstring NormalizeVisitKey(std::wstring_view path) noexcept
{
    std::wstring key = NormalizeSearchPath(path);
    return OrdinalString::FoldCaseInvariant(key);
}

[[nodiscard]] HRESULT TryGetDirectoryVisitIdentity(std::wstring_view path, DirectoryVisitIdentity& outIdentity) noexcept
{
    outIdentity = {};
    if (path.empty())
    {
        return E_INVALIDARG;
    }

    const PathInfo pathInfo = MakePathInfo(std::wstring(path));
    wil::unique_handle handle(::CreateFileW(pathInfo.extended.c_str(),
                                            FILE_READ_ATTRIBUTES,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_BACKUP_SEMANTICS,
                                            nullptr));
    if (! handle)
    {
        const DWORD lastError = ::GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0u ? lastError : ERROR_FILE_NOT_FOUND);
    }

    BY_HANDLE_FILE_INFORMATION info{};
    if (::GetFileInformationByHandle(handle.get(), &info) == 0)
    {
        const DWORD lastError = ::GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0u ? lastError : ERROR_GEN_FAILURE);
    }

    outIdentity.volumeSerialNumber = info.dwVolumeSerialNumber;
    outIdentity.fileIndex          = (static_cast<uint64_t>(info.nFileIndexHigh) << 32u) | static_cast<uint64_t>(info.nFileIndexLow);
    return S_OK;
}

[[nodiscard]] bool MarkQueuedDirectory(SearchRuntime& runtime, std::wstring_view fullPath) noexcept
{
    const std::wstring visitKey = NormalizeVisitKey(fullPath);
    {
        std::lock_guard lock(runtime.queuedDirectoriesMutex);
        if (! runtime.queuedDirectories.insert(visitKey).second)
        {
            return false;
        }
    }

    if (! runtime.followSymlinks)
    {
        return true;
    }

    DirectoryVisitIdentity identity{};
    if (FAILED(TryGetDirectoryVisitIdentity(fullPath, identity)))
    {
        return true;
    }

    std::lock_guard lock(runtime.queuedDirectoriesMutex);
    return runtime.queuedDirectoryIdentities.insert(identity).second;
}

[[nodiscard]] bool ShouldUseParallelDirectoryWalk(const SearchRuntime& runtime) noexcept
{
    return runtime.recursive && runtime.query != nullptr && runtime.query->contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED;
}

[[nodiscard]] unsigned int ResolveParallelDirectoryWalkWorkerCount(const SearchRuntime& runtime) noexcept
{
    constexpr unsigned int kMaxDirectoryWalkWorkers = 8u;
    return (std::max)(1u, (std::min)(runtime.parallelDirectoryWalkWorkers, kMaxDirectoryWalkWorkers));
}

HRESULT SearchDirectoryTree(SearchRuntime& runtime) noexcept;
HRESULT CheckSearchCancelled(SearchRuntime& runtime) noexcept;
HRESULT ReportSearchProgress(SearchRuntime& runtime, FileSystemSearchPhase phase, const std::wstring* currentPath, HRESULT statusHint, bool force) noexcept;
HRESULT EmitSearchMatch(SearchRuntime& runtime, const SearchEntryMetadata& entry, uint32_t matchedBy, const SearchContentResult& contentResult) noexcept;
[[nodiscard]] bool WorkerMatchNamePattern(const SearchRuntime& runtime, const std::wstring& displayName) noexcept;

HRESULT RejectRegexSearch(SearchRuntime& runtime, std::wstring_view regexKind, std::wstring_view rejectReason) noexcept
{
    Debug::Warning(L"FileSystem::Search: {} regex rejected: {}", regexKind, rejectReason);
    runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED;
    static_cast<void>(ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_COMPLETED, nullptr, E_INVALIDARG, true));
    return E_INVALIDARG;
}

// ---------------------------------------------------------------------------
// Parallel directory-walk infrastructure
// ---------------------------------------------------------------------------

struct ParallelDirectoryFrame final
{
    std::wstring fullPath;
    std::wstring relativeBase;
};

struct ParallelDirectoryResult final
{
    std::wstring fullPath;
    struct Match final
    {
        SearchEntryMetadata metadata;
        SearchContentResult contentResult;
        uint32_t matchedBy = FILESYSTEM_SEARCH_MATCH_SOURCE_NONE;
    };
    std::vector<Match> matches;
    uint64_t filesScanned = 0u;
    uint32_t warningFlags = FILESYSTEM_SEARCH_WARNING_NONE;
    HRESULT status        = S_OK;
};

struct ParallelDirectoryWalkState final
{
    ParallelDirectoryWalkState()                                             = default;
    ParallelDirectoryWalkState(const ParallelDirectoryWalkState&)            = delete;
    ParallelDirectoryWalkState& operator=(const ParallelDirectoryWalkState&) = delete;
    ParallelDirectoryWalkState(ParallelDirectoryWalkState&&)                 = delete;
    ParallelDirectoryWalkState& operator=(ParallelDirectoryWalkState&&)      = delete;

    std::mutex mutex;
    std::condition_variable cv;
    std::deque<ParallelDirectoryFrame> pendingDirectories;
    std::deque<ParallelDirectoryResult> completedResults;
    std::wstring latestPath;
    std::atomic<bool> cancelFlag{false};
    std::atomic<uint64_t> scannedDirectories{0u};
    std::atomic<uint64_t> scannedFiles{0u};
    std::atomic<uint32_t> warningFlags{FILESYSTEM_SEARCH_WARNING_NONE};
    uint64_t queuedDirectories    = 1u;
    uint64_t completedDirectories = 0u;
    unsigned int activeWorkers    = 0u;
    unsigned int remainingWorkers = 0u;
    unsigned int maxActiveWorkers = 0u;
    HRESULT terminalHr            = S_OK;
};

struct ParallelDirectoryWalkWorkerContext final
{
    ParallelDirectoryWalkWorkerContext()                                                     = default;
    ParallelDirectoryWalkWorkerContext(const ParallelDirectoryWalkWorkerContext&)            = delete;
    ParallelDirectoryWalkWorkerContext& operator=(const ParallelDirectoryWalkWorkerContext&) = delete;
    ParallelDirectoryWalkWorkerContext(ParallelDirectoryWalkWorkerContext&&)                 = delete;
    ParallelDirectoryWalkWorkerContext& operator=(ParallelDirectoryWalkWorkerContext&&)      = delete;

    SearchRuntime* runtime            = nullptr;
    ParallelDirectoryWalkState* state = nullptr;
    wil::unique_hmodule modulePin;
};

void RequestParallelDirectoryWalkStop(ParallelDirectoryWalkState& state, HRESULT hr) noexcept
{
    {
        std::lock_guard lock(state.mutex);
        state.cancelFlag.store(true, std::memory_order_release);
        if (FAILED(hr) && SUCCEEDED(state.terminalHr))
        {
            state.terminalHr = hr;
        }
    }

    state.cv.notify_all();
}

void EnqueueParallelDirectory(ParallelDirectoryWalkState& state, SearchRuntime& runtime, std::wstring fullPath, std::wstring relativeBase) noexcept
{
    if (! MarkQueuedDirectory(runtime, fullPath))
    {
        return;
    }

    {
        std::lock_guard lock(state.mutex);
        if (state.cancelFlag.load(std::memory_order_acquire))
        {
            return;
        }

        ParallelDirectoryFrame frame{};
        frame.fullPath     = std::move(fullPath);
        frame.relativeBase = std::move(relativeBase);
        state.pendingDirectories.push_back(std::move(frame));
        ++state.queuedDirectories;
    }

    state.cv.notify_one();
}

void AccumulateParallelNameOnlyEntry(const SearchRuntime& runtime, SearchEntryMetadata metadata, ParallelDirectoryResult& result) noexcept
{
    const bool isDirectory = (metadata.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    if ((isDirectory && ! runtime.includeDirectories) || (! isDirectory && ! runtime.includeFiles))
    {
        return;
    }

    if (! isDirectory && runtime.includeFiles)
    {
        ++result.filesScanned;
    }

    const bool nameMatched = WorkerMatchNamePattern(runtime, metadata.displayName);
    if (runtime.query->nameMode != FILESYSTEM_SEARCH_NAME_DISABLED && ! nameMatched)
    {
        return;
    }

    ParallelDirectoryResult::Match match{};
    match.metadata  = std::move(metadata);
    match.matchedBy = runtime.query->nameMode != FILESYSTEM_SEARCH_NAME_DISABLED ? FILESYSTEM_SEARCH_MATCH_SOURCE_NAME : FILESYSTEM_SEARCH_MATCH_SOURCE_NONE;
    result.matches.push_back(std::move(match));
}

[[nodiscard]] bool HasQueuedParallelDirectoryResultWork(const ParallelDirectoryResult& result) noexcept
{
    return SUCCEEDED(result.status) && ! result.matches.empty();
}

void QueueParallelDirectoryResultChunk(ParallelDirectoryWalkState& state, ParallelDirectoryResult& result) noexcept
{
    if (result.filesScanned != 0u)
    {
        state.scannedFiles.fetch_add(result.filesScanned, std::memory_order_acq_rel);
        result.filesScanned = 0u;
    }
    if (result.warningFlags != FILESYSTEM_SEARCH_WARNING_NONE)
    {
        state.warningFlags.fetch_or(result.warningFlags, std::memory_order_acq_rel);
    }
    if (! HasQueuedParallelDirectoryResultWork(result))
    {
        return;
    }

    ParallelDirectoryResult queued{};
    queued.fullPath     = result.fullPath;
    queued.matches      = std::move(result.matches);
    queued.warningFlags = result.warningFlags;
    queued.status       = result.status;
    result.matches.clear();
    result.warningFlags = FILESYSTEM_SEARCH_WARNING_NONE;
    result.status       = S_OK;

    {
        std::lock_guard lock(state.mutex);
        state.completedResults.push_back(std::move(queued));
    }
    state.cv.notify_all();
}

void BuildParallelDirectoryResult(SearchRuntime& runtime,
                                  ParallelDirectoryWalkState& state,
                                  const ParallelDirectoryFrame& frame,
                                  ParallelDirectoryResult& result) noexcept
{
    result.fullPath = frame.fullPath;
    state.scannedDirectories.fetch_add(1u, std::memory_order_acq_rel);

    if (state.cancelFlag.load(std::memory_order_acquire))
    {
        result.status = kCancelledHr;
        return;
    }

    wil::com_ptr<IFilesInformation> information;
    HRESULT hr = runtime.fileSystem->ReadDirectoryInfo(frame.fullPath.c_str(), information.put());
    if (FAILED(hr))
    {
        if (hr == kAccessDeniedHr && frame.fullPath != runtime.rootPath)
        {
            result.warningFlags |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
            state.warningFlags.fetch_or(FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED, std::memory_order_acq_rel);
            return;
        }

        result.status = hr;
        return;
    }

    unsigned long count = 0u;
    hr                  = information->GetCount(&count);
    if (FAILED(hr))
    {
        result.status = hr;
        return;
    }

    for (unsigned long index = 0; index < count; ++index)
    {
        if (state.cancelFlag.load(std::memory_order_acquire))
        {
            result.status = kCancelledHr;
            break;
        }

        FileInfo* entry = nullptr;
        hr              = information->Get(index, &entry);
        if (FAILED(hr) || entry == nullptr)
        {
            result.status = FAILED(hr) ? hr : E_FAIL;
            break;
        }

        const std::wstring_view name(entry->FileName, entry->FileNameSize / sizeof(wchar_t));
        if (IsDotOrDotDot(name))
        {
            continue;
        }

        SearchEntryMetadata metadata{};
        metadata.displayName    = std::wstring(name);
        metadata.relativePath   = frame.relativeBase.empty() ? metadata.displayName : NormalizeSearchPath(AppendPath(frame.relativeBase, metadata.displayName));
        metadata.fullPath       = NormalizeSearchPath(AppendPath(frame.fullPath, metadata.displayName));
        metadata.fileAttributes = entry->FileAttributes;
        metadata.creationTime   = entry->CreationTime;
        metadata.lastAccessTime = entry->LastAccessTime;
        metadata.lastWriteTime  = entry->LastWriteTime;
        metadata.changeTime     = entry->ChangeTime;
        metadata.endOfFile      = entry->EndOfFile;
        metadata.allocationSize = entry->AllocationSize;

        const bool isDirectory = (metadata.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const bool isReparse   = (metadata.fileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
        if (runtime.recursive && isDirectory && (! isReparse || runtime.followSymlinks))
        {
            EnqueueParallelDirectory(state, runtime, std::wstring(metadata.fullPath), std::wstring(metadata.relativePath));
        }

        AccumulateParallelNameOnlyEntry(runtime, std::move(metadata), result);
        if (result.matches.size() >= kParallelDirectoryResultFlushMatches)
        {
            QueueParallelDirectoryResultChunk(state, result);
        }
    }

    QueueParallelDirectoryResultChunk(state, result);
}

void CALLBACK ParallelDirectoryWalkWorkerCallback(PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
{
    auto* worker         = static_cast<ParallelDirectoryWalkWorkerContext*>(context);
    auto completionGuard = wil::scope_exit([&]
    {
        if (worker != nullptr && worker->state != nullptr)
        {
            {
                std::lock_guard lock(worker->state->mutex);
                if (worker->state->remainingWorkers > 0u)
                {
                    --worker->state->remainingWorkers;
                }
            }

            worker->state->cv.notify_all();
        }
    });

    if (worker == nullptr || worker->runtime == nullptr || worker->state == nullptr)
    {
        return;
    }

    static_cast<void>(worker->modulePin); // Keep DLL loaded for callback lifetime.
    [[maybe_unused]] auto coInit = wil::CoInitializeEx(COINIT_MULTITHREADED);

    SearchRuntime& runtime            = *worker->runtime;
    ParallelDirectoryWalkState& state = *worker->state;

    while (true)
    {
        ParallelDirectoryFrame frame{};
        {
            std::unique_lock lock(state.mutex);
            state.cv.wait(lock,
                          [&]
            {
                return state.cancelFlag.load(std::memory_order_acquire) || ! state.pendingDirectories.empty() ||
                       (state.activeWorkers == 0u && state.completedDirectories == state.queuedDirectories);
            });

            if (state.cancelFlag.load(std::memory_order_acquire))
            {
                break;
            }

            if (state.pendingDirectories.empty())
            {
                if (state.activeWorkers == 0u && state.completedDirectories == state.queuedDirectories)
                {
                    break;
                }

                continue;
            }

            frame = std::move(state.pendingDirectories.front());
            state.pendingDirectories.pop_front();
            ++state.activeWorkers;
            state.maxActiveWorkers = (std::max)(state.maxActiveWorkers, state.activeWorkers);
            state.latestPath       = frame.fullPath;
        }

        ParallelDirectoryResult result{};
        BuildParallelDirectoryResult(runtime, state, frame, result);

        {
            std::lock_guard lock(state.mutex);
            if (state.activeWorkers > 0u)
            {
                --state.activeWorkers;
            }

            ++state.completedDirectories;
            if (FAILED(result.status) && result.status != kCancelledHr && SUCCEEDED(state.terminalHr))
            {
                state.terminalHr = result.status;
                state.cancelFlag.store(true, std::memory_order_release);
            }

            if (HasQueuedParallelDirectoryResultWork(result))
            {
                state.completedResults.push_back(std::move(result));
            }
        }

        state.cv.notify_all();
    }
}

HRESULT SearchDirectoryTreeParallelNameOnly(SearchRuntime& runtime) noexcept
{
    Debug::Perf::Scope treePerf(L"FileSystem.Search.ScanTree");
    treePerf.SetDetail(runtime.rootPath);
    Debug::Perf::Scope parallelWalkPerf(L"FileSystem.Search.ParallelDirectoryWalk");
    parallelWalkPerf.SetDetail(runtime.rootPath);

    constexpr auto kCoordinatorPollInterval = std::chrono::milliseconds(50);
    const unsigned int workerCount          = ResolveParallelDirectoryWalkWorkerCount(runtime);

    ParallelDirectoryWalkState state{};
    static_cast<void>(MarkQueuedDirectory(runtime, runtime.rootPath));
    {
        std::lock_guard lock(state.mutex);
        ParallelDirectoryFrame rootFrame{};
        rootFrame.fullPath = runtime.rootPath;
        state.pendingDirectories.push_back(std::move(rootFrame));
        state.latestPath = runtime.rootPath;
    }

    std::vector<ParallelDirectoryWalkWorkerContext> workers(workerCount);
    bool workerSubmitted = false;
    for (unsigned int index = 0u; index < workerCount; ++index)
    {
        workers[index].runtime   = &runtime;
        workers[index].state     = &state;
        workers[index].modulePin = AcquireModuleReferenceFromAddress(&kFileSystemModuleAnchor);

        {
            std::lock_guard lock(state.mutex);
            ++state.remainingWorkers;
        }

        const BOOL queued = ::TrySubmitThreadpoolCallback(&ParallelDirectoryWalkWorkerCallback, &workers[index], nullptr);
        if (queued == 0)
        {
            bool noWorkersRemaining = false;
            {
                std::lock_guard lock(state.mutex);
                if (state.remainingWorkers > 0u)
                {
                    --state.remainingWorkers;
                }
                noWorkersRemaining = state.remainingWorkers == 0u;
            }

            if (! workerSubmitted && noWorkersRemaining)
            {
                return SearchDirectoryTree(runtime);
            }

            continue;
        }

        workerSubmitted = true;
    }

    if (! workerSubmitted)
    {
        return SearchDirectoryTree(runtime);
    }

    HRESULT finalHr = S_OK;
    while (true)
    {
        ParallelDirectoryResult readyResult{};
        bool haveReadyResult = false;
        bool workersDone     = false;
        std::wstring progressPath;
        HRESULT terminalHr = S_OK;

        {
            std::unique_lock lock(state.mutex);
            const auto ready = [&]
            {
                if (FAILED(state.terminalHr))
                {
                    return state.remainingWorkers == 0u;
                }

                return state.cancelFlag.load(std::memory_order_acquire) || ! state.completedResults.empty() || state.remainingWorkers == 0u;
            };
            if (! ready())
            {
                state.cv.wait_for(lock, kCoordinatorPollInterval, ready);
            }

            if (SUCCEEDED(state.terminalHr) && ! state.completedResults.empty())
            {
                readyResult = std::move(state.completedResults.front());
                state.completedResults.pop_front();
                haveReadyResult = true;
            }

            progressPath = state.latestPath;
            terminalHr   = state.terminalHr;
            workersDone  = FAILED(terminalHr) ? (state.remainingWorkers == 0u) : (state.remainingWorkers == 0u && state.completedResults.empty());
        }

        runtime.scannedDirectories = state.scannedDirectories.load(std::memory_order_acquire);
        runtime.scannedFiles       = state.scannedFiles.load(std::memory_order_acquire);
        runtime.warningFlags |= state.warningFlags.load(std::memory_order_acquire);

        if (FAILED(terminalHr))
        {
            if (workersDone)
            {
                finalHr = terminalHr;
                break;
            }

            continue;
        }

        if (haveReadyResult)
        {
            runtime.warningFlags |= readyResult.warningFlags;
            for (auto& match : readyResult.matches)
            {
                if (runtime.stopRequested)
                {
                    break;
                }

                const HRESULT emitHr = EmitSearchMatch(runtime, match.metadata, match.matchedBy, match.contentResult);
                if (FAILED(emitHr))
                {
                    RequestParallelDirectoryWalkStop(state, emitHr);
                    break;
                }

                const HRESULT cancelHr = CheckSearchCancelled(runtime);
                if (FAILED(cancelHr))
                {
                    RequestParallelDirectoryWalkStop(state, cancelHr);
                    break;
                }
            }

            if (runtime.stopRequested)
            {
                RequestParallelDirectoryWalkStop(state, S_OK);
            }
        }
        else
        {
            const std::wstring* currentPath = progressPath.empty() ? nullptr : &progressPath;
            const HRESULT progressHr        = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_ENUMERATING, currentPath, S_OK, false);
            if (FAILED(progressHr))
            {
                RequestParallelDirectoryWalkStop(state, progressHr);
            }
            else
            {
                const HRESULT cancelHr = CheckSearchCancelled(runtime);
                if (FAILED(cancelHr))
                {
                    RequestParallelDirectoryWalkStop(state, cancelHr);
                }
            }
        }

        if (workersDone)
        {
            break;
        }
    }

    runtime.scannedDirectories = state.scannedDirectories.load(std::memory_order_acquire);
    runtime.scannedFiles       = state.scannedFiles.load(std::memory_order_acquire);
    runtime.warningFlags |= state.warningFlags.load(std::memory_order_acquire);

    unsigned int maxActiveWorkers = 0u;
    {
        std::lock_guard lock(state.mutex);
        maxActiveWorkers = state.maxActiveWorkers;
    }

    treePerf.SetValue0(runtime.scannedDirectories);
    treePerf.SetValue1(maxActiveWorkers);
    parallelWalkPerf.SetValue0(runtime.scannedDirectories);
    parallelWalkPerf.SetValue1(maxActiveWorkers);

    return finalHr;
}

[[nodiscard]] std::wstring BuildRelativeSearchPath(std::wstring_view rootPath, std::wstring_view fullPath) noexcept
{
    const std::wstring normalizedRoot = NormalizeSearchPath(rootPath);
    const std::wstring normalizedFull = NormalizeSearchPath(fullPath);
    const std::wstring foldedRoot     = ToCaseFolded(normalizedRoot);
    const std::wstring foldedFull     = ToCaseFolded(normalizedFull);

    if (foldedFull.size() >= foldedRoot.size() && foldedFull.compare(0u, foldedRoot.size(), foldedRoot) == 0)
    {
        std::wstring_view remainder(normalizedFull);
        remainder.remove_prefix(normalizedRoot.size());
        while (! remainder.empty() && (remainder.front() == L'\\' || remainder.front() == L'/'))
        {
            remainder.remove_prefix(1u);
        }

        if (! remainder.empty())
        {
            return std::wstring(remainder);
        }
    }

    const std::wstring leaf = std::wstring(GetPathLeaf(normalizedFull));
    return leaf.empty() ? normalizedFull : leaf;
}

[[nodiscard]] SearchBackendSelection SelectSearchBackend(FileSystemSearchBackendPreference preference,
                                                         FileSystemSearchFlags flags,
                                                         const LocalSearchIndexCore::SupportInfo& support,
                                                         bool preferServiceBackend) noexcept
{
    SearchBackendSelection selection{};
    if ((flags & FILESYSTEM_SEARCH_FORCE_SCAN) != 0)
    {
        selection.backend = FILESYSTEM_SEARCH_BACKEND_SCAN;
        return selection;
    }

    const bool preferIndexHint = (flags & FILESYSTEM_SEARCH_PREFER_INDEX) != 0;

    switch (preference)
    {
        case FileSystemSearchBackendPreference::Auto:
            if (preferServiceBackend)
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_SERVICE;
            }
            else if (support.indexable)
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_INDEX;
            }
            else if (preferIndexHint)
            {
                selection.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
            }
            break;

        case FileSystemSearchBackendPreference::Service:
            if (preferServiceBackend)
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_SERVICE;
            }
            else if (support.indexable)
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_INDEX;
            }
            else
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_SCAN;
                selection.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
            }
            break;

        case FileSystemSearchBackendPreference::LocalIndex:
            if (support.indexable)
            {
                selection.backend = FILESYSTEM_SEARCH_BACKEND_INDEX;
            }
            else
            {
                selection.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
            }
            break;

        case FileSystemSearchBackendPreference::Scan: selection.backend = FILESYSTEM_SEARCH_BACKEND_SCAN; break;
    }

    return selection;
}

[[nodiscard]] bool WildcardMatchCaseSensitive(std::wstring_view text, std::wstring_view pattern) noexcept
{
    size_t textPos    = 0;
    size_t patternPos = 0;
    size_t starPos    = std::wstring_view::npos;
    size_t matchPos   = 0;

    while (textPos < text.size())
    {
        if (patternPos < pattern.size() && (pattern[patternPos] == L'?' || pattern[patternPos] == text[textPos]))
        {
            ++patternPos;
            ++textPos;
            continue;
        }

        if (patternPos < pattern.size() && pattern[patternPos] == L'*')
        {
            starPos  = patternPos++;
            matchPos = textPos;
            continue;
        }

        if (starPos != std::wstring_view::npos)
        {
            patternPos = starPos + 1;
            textPos    = ++matchPos;
            continue;
        }

        return false;
    }

    while (patternPos < pattern.size() && pattern[patternPos] == L'*')
    {
        ++patternPos;
    }

    return patternPos == pattern.size();
}

[[nodiscard]] bool WildcardMatch(std::wstring_view text, std::wstring_view pattern, bool caseSensitive) noexcept
{
    if (caseSensitive)
    {
        return WildcardMatchCaseSensitive(text, pattern);
    }

    return WildcardMatchCaseSensitive(ToCaseFolded(text), ToCaseFolded(pattern));
}

[[nodiscard]] bool FindLiteral(std::wstring_view haystack, std::wstring_view needle, bool caseSensitive, size_t& outPosition) noexcept
{
    outPosition = std::wstring_view::npos;

    if (needle.empty())
    {
        outPosition = 0;
        return true;
    }

    if (caseSensitive)
    {
        outPosition = haystack.find(needle);
        return outPosition != std::wstring_view::npos;
    }

    const std::wstring foldedHaystack = ToCaseFolded(haystack);
    const std::wstring foldedNeedle   = ToCaseFolded(needle);
    outPosition                       = foldedHaystack.find(foldedNeedle);
    return outPosition != std::wstring_view::npos;
}

HRESULT CheckSearchCancelled(SearchRuntime& runtime) noexcept
{
    BOOL cancel            = FALSE;
    const HRESULT cancelHr = runtime.callback->FileSystemSearchShouldCancel(&cancel, runtime.cookie);
    if (FAILED(cancelHr))
    {
        return cancelHr;
    }

    return cancel ? kCancelledHr : S_OK;
}

HRESULT STDMETHODCALLTYPE SearchReadCancelThunk(void* cookie) noexcept
{
    if (cookie == nullptr)
    {
        return E_POINTER;
    }

    return CheckSearchCancelled(*static_cast<SearchRuntime*>(cookie));
}

[[nodiscard]] bool IsCallbackCancellationResult(HRESULT hr) noexcept
{
    return hr == E_ABORT || hr == kCancelledHr;
}

// Validates that a regex pattern doesn't contain known pathological constructs
// (nested quantifiers) that could cause catastrophic backtracking (ReDoS).
[[nodiscard]] bool ValidateRegexPatternSafety(std::wstring_view pattern, std::wstring& outReason) noexcept
{
    if (pattern.size() > kMaxRegexPatternLength)
    {
        outReason = L"Regex pattern exceeds maximum length.";
        return false;
    }

    struct GroupState
    {
        bool containsQuantifier = false;
    };

    GroupState groups[kMaxRegexGroupDepth + 1]{};
    size_t depth                      = 0;
    bool inCharClass                  = false;
    bool lastWasEscape                = false;
    bool prevWasOpenParen             = false;
    bool lastWasGroupClose            = false;
    bool lastGroupContainedQuantifier = false;

    for (size_t i = 0; i < pattern.size(); ++i)
    {
        const wchar_t ch = pattern[i];

        if (lastWasEscape)
        {
            lastWasEscape     = false;
            lastWasGroupClose = false;
            prevWasOpenParen  = false;
            continue;
        }

        if (ch == L'\\')
        {
            lastWasEscape     = true;
            lastWasGroupClose = false;
            prevWasOpenParen  = false;
            continue;
        }

        if (inCharClass)
        {
            if (ch == L']')
            {
                inCharClass = false;
            }
            continue;
        }

        if (ch == L'[')
        {
            inCharClass       = true;
            lastWasGroupClose = false;
            prevWasOpenParen  = false;
            continue;
        }

        if (ch == L'(')
        {
            ++depth;
            if (depth > kMaxRegexGroupDepth)
            {
                outReason = L"Regex group nesting too deep.";
                return false;
            }
            groups[depth]     = {};
            lastWasGroupClose = false;
            prevWasOpenParen  = true;
            continue;
        }

        if (ch == L')')
        {
            prevWasOpenParen = false;
            if (depth == 0)
            {
                lastWasGroupClose = false;
                continue;
            }
            lastGroupContainedQuantifier = groups[depth].containsQuantifier;
            // Propagate: if the closed group contained a quantifier, parent does too.
            if (lastGroupContainedQuantifier && depth > 0)
            {
                groups[depth - 1].containsQuantifier = true;
            }
            --depth;
            lastWasGroupClose = true;
            continue;
        }

        // Detect unbounded quantifiers: +, *, {n,...}
        bool isUnboundedQuantifier = (ch == L'+' || ch == L'*');
        if (! isUnboundedQuantifier && ch == L'{')
        {
            size_t j = i + 1;
            while (j < pattern.size() && pattern[j] >= L'0' && pattern[j] <= L'9')
            {
                ++j;
            }
            if (j > i + 1 && j < pattern.size() && (pattern[j] == L'}' || pattern[j] == L','))
            {
                isUnboundedQuantifier = true;
            }
        }

        if (isUnboundedQuantifier)
        {
            if (lastWasGroupClose && lastGroupContainedQuantifier)
            {
                outReason = L"Nested repetition in regex pattern (potential ReDoS).";
                return false;
            }
            groups[depth].containsQuantifier = true;
            lastWasGroupClose                = false;
            prevWasOpenParen                 = false;
            continue;
        }

        // ? immediately after ( is group syntax (?:, (?=, etc.), not a quantifier.
        if (ch == L'?')
        {
            if (prevWasOpenParen)
            {
                prevWasOpenParen = false;
                continue;
            }
            // Standalone ? is a quantifier (0 or 1). It's bounded so it can't cause
            // ReDoS as the outer quantifier, but it marks the group for inner detection
            // (e.g. (a?)+ IS dangerous when the outer quantifier is unbounded).
            groups[depth].containsQuantifier = true;
            lastWasGroupClose                = false;
            prevWasOpenParen                 = false;
            continue;
        }

        lastWasGroupClose = false;
        prevWasOpenParen  = false;
    }

    return true;
}

[[nodiscard]] LocalSearchIndexCore::StoreState ResolveFallbackStoreState(LocalSearchIndexCore::FallbackReason reason) noexcept
{
    switch (reason)
    {
        case LocalSearchIndexCore::FallbackReason::StoreMissing:
        case LocalSearchIndexCore::FallbackReason::StoreInvalid: return LocalSearchIndexCore::StoreState::Invalid;
        case LocalSearchIndexCore::FallbackReason::SqliteFailure: return LocalSearchIndexCore::StoreState::Recovering;
        case LocalSearchIndexCore::FallbackReason::StoreStale:
        case LocalSearchIndexCore::FallbackReason::CutoverBlocked:
        case LocalSearchIndexCore::FallbackReason::WarmupRunning: return LocalSearchIndexCore::StoreState::Syncing;
        case LocalSearchIndexCore::FallbackReason::None:
        default: return LocalSearchIndexCore::StoreState::Unknown;
    }
}

HRESULT ReportServiceStatusSnapshot(SearchRuntime& runtime,
                                    LocalSearchIndexCore::StoreState storeState,
                                    LocalSearchIndexCore::SyncPhase syncPhase,
                                    LocalSearchIndexCore::QueryExecutionMode queryExecutionMode,
                                    LocalSearchIndexCore::FallbackReason fallbackReason,
                                    uint64_t completedRoots,
                                    uint64_t totalRoots,
                                    std::wstring_view activeRoot) noexcept
{
    const FileSystemSearchHostExtensions* const hostExtensions = runtime.hostExtensions;
    if (! hostExtensions || hostExtensions->sizeBytes != sizeof(FileSystemSearchHostExtensions) ||
        hostExtensions->version != FILESYSTEM_SEARCH_HOST_EXTENSIONS_V1 || ! hostExtensions->serviceStatusCallback)
    {
        return S_OK;
    }

    FileSystemSearchServiceStatus status{};
    status.sizeBytes          = sizeof(FileSystemSearchServiceStatus);
    status.storeState         = static_cast<uint32_t>(storeState);
    status.syncPhase          = static_cast<uint32_t>(syncPhase);
    status.queryExecutionMode = static_cast<uint32_t>(queryExecutionMode);
    status.fallbackReason     = static_cast<uint32_t>(fallbackReason);
    status.completedRoots     = completedRoots;
    status.totalRoots         = totalRoots;
    status.activeRoot         = activeRoot.empty() ? nullptr : activeRoot.data();
    status.activeRootSize     = activeRoot.empty() ? 0u : ByteCountOfString(activeRoot);
    return hostExtensions->serviceStatusCallback(&status, hostExtensions->serviceStatusCookie);
}

HRESULT ReportServiceStatus(SearchRuntime& runtime, const SearchServiceBroker::QueryProgress& progress) noexcept
{
    return ReportServiceStatusSnapshot(runtime,
                                       progress.storeState,
                                       progress.syncPhase,
                                       progress.queryExecutionMode,
                                       progress.fallbackReason,
                                       progress.completedRoots,
                                       progress.totalRoots,
                                       progress.activeRoot);
}

HRESULT ReportSearchProgress(SearchRuntime& runtime, FileSystemSearchPhase phase, const std::wstring* currentPath, HRESULT statusHint, bool force) noexcept
{
    const uint64_t processedItems = runtime.scannedDirectories + runtime.scannedFiles;
    const ULONGLONG now           = ::GetTickCount64();

    if (! force && runtime.hasReportedProgress && phase == runtime.lastProgressPhase && (processedItems - runtime.lastProgressItems) < kProgressIntervalItems &&
        (now - runtime.lastProgressTick) < kProgressIntervalMs)
    {
        return S_OK;
    }

    FileSystemSearchProgress progress{};
    progress.sizeBytes          = sizeof(FileSystemSearchProgress);
    progress.phase              = phase;
    progress.backend            = runtime.backend;
    progress.warningFlags       = runtime.warningFlags;
    progress.statusHint         = statusHint;
    progress.scannedDirectories = runtime.scannedDirectories;
    progress.scannedFiles       = runtime.scannedFiles;
    progress.candidateFiles     = runtime.candidateFiles;
    progress.matchedEntries     = runtime.matchedEntries;
    progress.currentPath        = currentPath ? currentPath->c_str() : nullptr;
    progress.currentPathSize    = currentPath ? ByteCountOfString(*currentPath) : 0u;

    const HRESULT hr = runtime.callback->FileSystemSearchProgress(&progress, runtime.cookie);
    if (FAILED(hr))
    {
        return IsCallbackCancellationResult(hr) ? kCancelledHr : hr;
    }

    runtime.hasReportedProgress = true;
    runtime.lastProgressPhase   = phase;
    runtime.lastProgressTick    = now;
    runtime.lastProgressItems   = processedItems;
    return S_OK;
}

[[nodiscard]] bool MatchNamePattern(const SearchRuntime& runtime, const std::wstring& displayName) noexcept
{
    switch (runtime.query->nameMode)
    {
        case FILESYSTEM_SEARCH_NAME_DISABLED: return true;
        case FILESYSTEM_SEARCH_NAME_WILDCARD: return WildcardMatch(displayName, runtime.namePattern, runtime.matchCaseName);
        case FILESYSTEM_SEARCH_NAME_LITERAL:
        {
            size_t position = std::wstring::npos;
            return FindLiteral(displayName, runtime.namePattern, runtime.matchCaseName, position);
        }
        case FILESYSTEM_SEARCH_NAME_REGEX:
        {
            if (! runtime.nameRegex)
            {
                return false;
            }
            // noexcept boundary: std::regex_search may throw regex_error on
            // implementation-defined complexity/stack limits.
            try
            {
                return std::regex_search(displayName, *runtime.nameRegex);
            }
            catch (const std::regex_error&)
            {
                return false;
            }
        }
    }

    return false;
}

HRESULT MatchFileContent(SearchRuntime& runtime, const std::wstring& fullPath, SearchContentResult& result) noexcept
{
    result = {};

    PathInfo pathInfo = MakePathInfo(fullPath);
    wil::com_ptr<IFileReader> reader;
    HRESULT hr = runtime.fileSystem->CreateFileReader(pathInfo.extended.c_str(), reader.put());
    if (FAILED(hr))
    {
        if (hr == kAccessDeniedHr)
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
            return S_OK;
        }

        return hr;
    }

    if (! runtime.usingIndexedEnumeration)
    {
        ++runtime.candidateFiles;
    }

    hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_CONTENT_SCAN, &fullPath, S_OK, false);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = CheckSearchCancelled(runtime);
    if (FAILED(hr))
    {
        return hr;
    }

    SearchTextHelpers::TextSearchPattern pattern{};
    pattern.mode                   = runtime.query->contentMode;
    pattern.pattern                = runtime.contentPattern;
    pattern.compiledRegex          = runtime.contentRegex.get();
    pattern.caseSensitive          = runtime.matchCaseContent;
    pattern.literalChunkCharacters = SearchTextHelpers::kDefaultLiteralChunkChars;

    SearchTextHelpers::TextSearchResult helperResult{};
    hr = SearchTextHelpers::SearchFileReaderText(reader.get(),
                                                 pattern,
                                                 runtime.maxContentBytesPerFile,
                                                 0u,
                                                 runtime.maxSnippetCharacters,
                                                 runtime.wantSnippets,
                                                 &SearchReadCancelThunk,
                                                 &runtime,
                                                 helperResult);
    if (hr == kFileTooLargeHr)
    {
        return S_OK;
    }
    if (FAILED(hr))
    {
        if (hr == kAccessDeniedHr)
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
            return S_OK;
        }

        return hr;
    }

    result.matched     = helperResult.matched;
    result.byteOffset  = helperResult.matchOffset;
    result.byteLength  = helperResult.matchLength;
    result.previewText = std::move(helperResult.previewText);
    return S_OK;
}

HRESULT EmitSearchMatch(SearchRuntime& runtime, const SearchEntryMetadata& entry, uint32_t matchedBy, const SearchContentResult& contentResult) noexcept
{
    const auto [_, inserted] = runtime.emittedMatchPaths.emplace(entry.fullPath);
    if (! inserted)
    {
        return S_OK;
    }

    FileSystemSearchMatch match{};
    match.sizeBytes              = sizeof(FileSystemSearchMatch);
    match.fullPath               = entry.fullPath.c_str();
    match.fullPathSize           = ByteCountOfString(entry.fullPath);
    match.relativePath           = entry.relativePath.c_str();
    match.relativePathSize       = ByteCountOfString(entry.relativePath);
    match.displayName            = entry.displayName.c_str();
    match.displayNameSize        = ByteCountOfString(entry.displayName);
    match.previewText            = contentResult.previewText.empty() ? nullptr : contentResult.previewText.c_str();
    match.previewTextSize        = ByteCountOfString(contentResult.previewText);
    match.fileAttributes         = entry.fileAttributes;
    match.creationTime           = entry.creationTime;
    match.lastAccessTime         = entry.lastAccessTime;
    match.lastWriteTime          = entry.lastWriteTime;
    match.changeTime             = entry.changeTime;
    match.endOfFile              = entry.endOfFile;
    match.allocationSize         = entry.allocationSize;
    match.matchedBy              = matchedBy;
    match.contentMatchByteOffset = contentResult.byteOffset;
    match.contentMatchByteLength = contentResult.byteLength;

    const HRESULT hr = runtime.callback->FileSystemSearchMatch(&match, runtime.cookie);
    if (FAILED(hr))
    {
        return IsCallbackCancellationResult(hr) ? kCancelledHr : hr;
    }

    ++runtime.matchedEntries;
    if (runtime.maxResults != 0 && runtime.matchedEntries >= runtime.maxResults)
    {
        runtime.stopRequested = true;
    }

    return S_OK;
}

HRESULT EvaluateEntry(SearchRuntime& runtime, const SearchEntryMetadata& entry) noexcept
{
    const bool isDirectory = (entry.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    if ((isDirectory && ! runtime.includeDirectories) || (! isDirectory && ! runtime.includeFiles))
    {
        return S_OK;
    }

    if (! isDirectory && ! runtime.usingIndexedEnumeration)
    {
        ++runtime.scannedFiles;
    }

    const bool nameMatched = MatchNamePattern(runtime, entry.displayName);
    if (runtime.query->nameMode != FILESYSTEM_SEARCH_NAME_DISABLED && ! nameMatched)
    {
        return S_OK;
    }

    SearchContentResult contentResult{};
    bool contentMatched = (runtime.query->contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED);
    if (! contentMatched)
    {
        if (isDirectory)
        {
            return S_OK;
        }

        const HRESULT hr = MatchFileContent(runtime, entry.fullPath, contentResult);
        if (FAILED(hr))
        {
            return hr;
        }

        contentMatched = contentResult.matched;
    }

    if (! contentMatched)
    {
        return S_OK;
    }

    uint32_t matchedBy = FILESYSTEM_SEARCH_MATCH_SOURCE_NONE;
    if (runtime.query->nameMode != FILESYSTEM_SEARCH_NAME_DISABLED)
    {
        matchedBy |= FILESYSTEM_SEARCH_MATCH_SOURCE_NAME;
    }
    if (runtime.query->contentMode != FILESYSTEM_SEARCH_CONTENT_DISABLED)
    {
        matchedBy |= FILESYSTEM_SEARCH_MATCH_SOURCE_CONTENT;
    }

    return EmitSearchMatch(runtime, entry, matchedBy, contentResult);
}

void PopulateSinglePathMetadata(std::wstring_view fullPath, unsigned long fileAttributes, SearchEntryMetadata& outMetadata) noexcept
{
    outMetadata              = {};
    outMetadata.fullPath     = NormalizeSearchPath(fullPath);
    outMetadata.relativePath = std::wstring(GetPathLeaf(outMetadata.fullPath));
    if (outMetadata.relativePath.empty())
    {
        outMetadata.relativePath = outMetadata.fullPath;
    }
    outMetadata.displayName    = outMetadata.relativePath;
    outMetadata.fileAttributes = fileAttributes;

    const PathInfo pathInfo = MakePathInfo(outMetadata.fullPath);
    wil::unique_handle handle(::CreateFileW(pathInfo.extended.c_str(),
                                            FILE_READ_ATTRIBUTES,
                                            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                            nullptr,
                                            OPEN_EXISTING,
                                            FILE_FLAG_BACKUP_SEMANTICS,
                                            nullptr));
    if (! handle)
    {
        return;
    }

    FILE_BASIC_INFO basic{};
    if (::GetFileInformationByHandleEx(handle.get(), FileBasicInfo, &basic, sizeof(basic)) != 0)
    {
        outMetadata.creationTime   = basic.CreationTime.QuadPart;
        outMetadata.lastAccessTime = basic.LastAccessTime.QuadPart;
        outMetadata.lastWriteTime  = basic.LastWriteTime.QuadPart;
        outMetadata.changeTime     = basic.ChangeTime.QuadPart;
        if (basic.FileAttributes != 0u)
        {
            outMetadata.fileAttributes = basic.FileAttributes;
        }
    }

    FILE_STANDARD_INFO standard{};
    if (::GetFileInformationByHandleEx(handle.get(), FileStandardInfo, &standard, sizeof(standard)) != 0)
    {
        outMetadata.endOfFile      = standard.EndOfFile.QuadPart;
        outMetadata.allocationSize = standard.AllocationSize.QuadPart;
    }
}

void PopulateIndexedCandidateMetadata(SearchRuntime& runtime, const LocalSearchIndexCore::Candidate& candidate, SearchEntryMetadata& outMetadata) noexcept
{
    const bool hasPersistedEndOfFile      = (candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_END_OF_FILE) != 0u;
    const bool hasPersistedLastWriteTime  = (candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_LAST_WRITE_TIME) != 0u;
    const bool hasPersistedCreationTime   = (candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_CREATION_TIME) != 0u;
    const bool hasPersistedLastAccessTime = (candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_LAST_ACCESS_TIME) != 0u;
    const bool hasPersistedChangeTime     = (candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_CHANGE_TIME) != 0u;
    const bool hasPersistedAllocationSize = (candidate.metadataFlags & LocalSearchIndexCore::CANDIDATE_METADATA_ALLOCATION_SIZE) != 0u;

    if (hasPersistedEndOfFile && hasPersistedLastWriteTime && hasPersistedCreationTime && hasPersistedLastAccessTime && hasPersistedChangeTime &&
        hasPersistedAllocationSize)
    {
        outMetadata              = {};
        outMetadata.fullPath     = NormalizeSearchPath(candidate.fullPath);
        outMetadata.relativePath = BuildRelativeSearchPath(runtime.rootPath, outMetadata.fullPath);
        outMetadata.displayName  = candidate.displayName.empty() ? std::wstring(GetPathLeaf(outMetadata.fullPath)) : candidate.displayName;
        if (outMetadata.displayName.empty())
        {
            outMetadata.displayName = outMetadata.relativePath;
        }
        outMetadata.fileAttributes = candidate.fileAttributes;
        outMetadata.creationTime   = candidate.creationTime100ns;
        outMetadata.lastAccessTime = candidate.lastAccessTime100ns;
        outMetadata.lastWriteTime  = candidate.lastWriteTime100ns;
        outMetadata.changeTime     = candidate.changeTime100ns;
        outMetadata.endOfFile      = candidate.endOfFile;
        outMetadata.allocationSize = candidate.allocationSize;
        return;
    }

    PopulateSinglePathMetadata(candidate.fullPath, candidate.fileAttributes, outMetadata);
    outMetadata.relativePath = BuildRelativeSearchPath(runtime.rootPath, outMetadata.fullPath);
    outMetadata.displayName  = candidate.displayName.empty() ? std::wstring(GetPathLeaf(outMetadata.fullPath)) : candidate.displayName;
    if (outMetadata.displayName.empty())
    {
        outMetadata.displayName = outMetadata.relativePath;
    }
}

void ApplyIndexedQueryStats(SearchRuntime& runtime, const LocalSearchIndexCore::QueryStats& stats) noexcept
{
    runtime.scannedDirectories = (std::max)(runtime.scannedDirectories, stats.directoryCount);
    runtime.scannedFiles       = (std::max)(runtime.scannedFiles, stats.fileCount);
    if (runtime.stopRequested && runtime.indexedCandidatesConsumed != 0u)
    {
        runtime.candidateFiles = runtime.indexedCandidatesConsumed;
        return;
    }

    runtime.candidateFiles = (std::max)(runtime.candidateFiles, stats.candidateCount);
}

HRESULT HandleIndexedCandidate(SearchRuntime& runtime, const LocalSearchIndexCore::Candidate& candidate) noexcept
{
    if (runtime.stopRequested)
    {
        return S_FALSE;
    }

    HRESULT hr = CheckSearchCancelled(runtime);
    if (FAILED(hr))
    {
        return hr;
    }

    ++runtime.indexedCandidatesConsumed;
    SearchEntryMetadata metadata{};
    PopulateIndexedCandidateMetadata(runtime, candidate, metadata);

    hr = EvaluateEntry(runtime, metadata);
    if (FAILED(hr))
    {
        return hr;
    }

    return runtime.stopRequested ? S_FALSE : S_OK;
}

HRESULT STDMETHODCALLTYPE SearchIndexedCandidateThunk(LocalSearchIndexCore::Candidate* candidate, void* cookie) noexcept
{
    if (candidate == nullptr || cookie == nullptr)
    {
        return E_POINTER;
    }

    return HandleIndexedCandidate(*static_cast<SearchRuntime*>(cookie), *candidate);
}

HRESULT ApplyIndexedProgressUpdate(SearchRuntime& runtime,
                                   FileSystemSearchPhase phase,
                                   uint64_t scannedDirectories,
                                   uint64_t scannedFiles,
                                   uint64_t candidateFiles,
                                   uint64_t matchedEntries,
                                   const std::wstring* currentPath,
                                   HRESULT statusHint) noexcept
{
    runtime.scannedDirectories = scannedDirectories;
    runtime.scannedFiles       = scannedFiles;
    runtime.candidateFiles     = candidateFiles;
    runtime.matchedEntries     = matchedEntries;
    return ReportSearchProgress(runtime, phase, currentPath, statusHint, true);
}

HRESULT STDMETHODCALLTYPE SearchIndexedProgressThunk(const LocalSearchIndexCore::ProgressUpdate* progress, void* cookie) noexcept
{
    if (progress == nullptr || cookie == nullptr)
    {
        return E_POINTER;
    }

    SearchRuntime& runtime          = *static_cast<SearchRuntime*>(cookie);
    const std::wstring* currentPath = progress->currentPath.empty() ? nullptr : &progress->currentPath;
    return ApplyIndexedProgressUpdate(runtime,
                                      progress->phase,
                                      progress->scannedDirectories,
                                      progress->scannedFiles,
                                      progress->candidateFiles,
                                      progress->matchedEntries,
                                      currentPath,
                                      progress->statusHint);
}

HRESULT STDMETHODCALLTYPE SearchServiceCandidateBatchThunk(LocalSearchIndexCore::Candidate* candidates,
                                                           size_t count,
                                                           size_t* consumedCount,
                                                           void* cookie) noexcept
{
    if ((candidates == nullptr && count != 0u) || consumedCount == nullptr || cookie == nullptr)
    {
        return E_POINTER;
    }

    *consumedCount = 0u;
    auto& runtime  = *static_cast<SearchRuntime*>(cookie);
    for (size_t index = 0u; index < count; ++index)
    {
        const HRESULT hr = HandleIndexedCandidate(runtime, candidates[index]);
        if (hr == S_OK || hr == S_FALSE)
        {
            ++(*consumedCount);
        }
        if (hr != S_OK)
        {
            return hr;
        }
    }

    return runtime.stopRequested ? S_FALSE : S_OK;
}

HRESULT SearchIndexedTree(SearchRuntime& runtime, LocalSearchIndexCore::Repository& repository) noexcept
{
    Debug::Perf::Scope indexPerf(L"FileSystem.Search.IndexedTree");
    indexPerf.SetDetail(runtime.rootPath);

    LocalSearchIndexCore::QueryPlan plan{};
    plan.rootPath           = runtime.rootPath;
    plan.namePattern        = runtime.namePattern;
    plan.nameMode           = runtime.query->nameMode;
    plan.compiledNameRegex  = runtime.nameRegex.get();
    plan.matchCaseName      = runtime.matchCaseName;
    plan.recursive          = runtime.recursive;
    plan.includeFiles       = runtime.includeFiles;
    plan.includeDirectories = runtime.includeDirectories;
    plan.maxResults         = runtime.query->contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED ? runtime.maxResults : 0u;

    LocalSearchIndexCore::QueryStats stats{};

    HRESULT hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, &runtime.rootPath, S_OK, true);
    if (FAILED(hr))
    {
        return hr;
    }

    runtime.usingIndexedEnumeration   = true;
    runtime.indexedCandidatesConsumed = 0u;
    hr = repository.Enumerate(plan, &SearchReadCancelThunk, &runtime, &SearchIndexedCandidateThunk, &runtime, &stats, &SearchIndexedProgressThunk, &runtime);
    if (FAILED(hr))
    {
        return hr;
    }

    ApplyIndexedQueryStats(runtime, stats);

    hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, &runtime.rootPath, S_OK, true);
    if (FAILED(hr))
    {
        return hr;
    }

    indexPerf.SetValue0(runtime.matchedEntries);
    indexPerf.SetValue1(runtime.candidateFiles);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SearchServiceProgressThunk(const SearchServiceBroker::QueryProgress* progress, void* cookie) noexcept
{
    if (progress == nullptr || cookie == nullptr)
    {
        return E_POINTER;
    }

    SearchRuntime& runtime = *static_cast<SearchRuntime*>(cookie);
    const HRESULT statusHr = ReportServiceStatus(runtime, *progress);
    if (FAILED(statusHr))
    {
        return IsCallbackCancellationResult(statusHr) ? kCancelledHr : statusHr;
    }

    runtime.warningFlags |= progress->warningFlags;
    const std::wstring* currentPath = progress->currentPath.empty() ? nullptr : &progress->currentPath;
    return ApplyIndexedProgressUpdate(runtime,
                                      progress->phase,
                                      progress->scannedDirectories,
                                      progress->scannedFiles,
                                      progress->candidateFiles,
                                      progress->matchedEntries,
                                      currentPath,
                                      progress->statusHint);
}

HRESULT SearchServiceTree(SearchRuntime& runtime) noexcept
{
    Debug::Perf::Scope servicePerf(L"FileSystem.Search.ServiceTree");
    servicePerf.SetDetail(runtime.rootPath);

    SearchServiceBroker::QueryRequest request{};
    request.rootPath           = runtime.rootPath;
    request.namePattern        = runtime.namePattern;
    request.nameMode           = runtime.query->nameMode;
    request.flags              = runtime.query->flags;
    request.recursive          = runtime.recursive;
    request.includeFiles       = runtime.includeFiles;
    request.includeDirectories = runtime.includeDirectories;
    request.matchCaseName      = runtime.matchCaseName;
    request.maxResults         = runtime.query->contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED ? runtime.maxResults : 0u;

    std::vector<LocalSearchIndexCore::Candidate> candidates;
    LocalSearchIndexCore::QueryStats stats{};
    runtime.usingIndexedEnumeration   = true;
    runtime.indexedCandidatesConsumed = 0u;

    SearchServiceBroker::ServiceStatus preQueryStatus{};
    const HRESULT preQueryStatusHr = SearchServiceBroker::GetStatus(preQueryStatus);
    if (SUCCEEDED(preQueryStatusHr))
    {
        LocalSearchIndexCore::QueryExecutionMode preQueryMode       = preQueryStatus.queryExecutionMode;
        LocalSearchIndexCore::FallbackReason preQueryFallbackReason = preQueryStatus.fallbackReason;
        const uint64_t preQueryCompletedRoots =
            preQueryStatus.completedRoots != 0u || preQueryStatus.totalRoots != 0u ? preQueryStatus.completedRoots : preQueryStatus.startupWarmupCompletedRoots;
        const uint64_t preQueryTotalRoots =
            preQueryStatus.completedRoots != 0u || preQueryStatus.totalRoots != 0u ? preQueryStatus.totalRoots : preQueryStatus.startupWarmupTotalRoots;
        const std::wstring_view preQueryActiveRoot =
            ! preQueryStatus.activeRoot.empty() ? std::wstring_view(preQueryStatus.activeRoot) : std::wstring_view(preQueryStatus.startupWarmupCurrentRoot);
        if (preQueryMode == LocalSearchIndexCore::QueryExecutionMode::Unknown && preQueryFallbackReason != LocalSearchIndexCore::FallbackReason::None)
        {
            preQueryMode = LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback;
        }
        else if (preQueryMode == LocalSearchIndexCore::QueryExecutionMode::Unknown &&
                 (preQueryStatus.startupWarmupRunning || (preQueryTotalRoots != 0u && preQueryCompletedRoots < preQueryTotalRoots)))
        {
            preQueryMode           = LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback;
            preQueryFallbackReason = LocalSearchIndexCore::FallbackReason::WarmupRunning;
        }
        else if (preQueryMode == LocalSearchIndexCore::QueryExecutionMode::Unknown && ! preQueryStatus.readyForQueryCutover)
        {
            preQueryMode           = LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback;
            preQueryFallbackReason = LocalSearchIndexCore::FallbackReason::CutoverBlocked;
        }

        if (preQueryMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback)
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
        }

        const HRESULT statusHr = ReportServiceStatusSnapshot(runtime,
                                                             preQueryStatus.storeState,
                                                             preQueryStatus.syncPhase,
                                                             preQueryMode,
                                                             preQueryFallbackReason,
                                                             preQueryCompletedRoots,
                                                             preQueryTotalRoots,
                                                             preQueryActiveRoot);
        if (FAILED(statusHr))
        {
            return IsCallbackCancellationResult(statusHr) ? kCancelledHr : statusHr;
        }

        const HRESULT progressHr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, &runtime.rootPath, S_OK, true);
        if (FAILED(progressHr))
        {
            return progressHr;
        }
    }

#if defined(_DEBUG) || defined(__SANITIZE_ADDRESS__)
    Debug::Info(L"FileSystem::Search: dispatching service query root='{}' pattern='{}' mode={} flags=0x{:08X} maxResults={}",
                runtime.rootPath,
                runtime.namePattern,
                static_cast<uint32_t>(runtime.query->nameMode),
                static_cast<unsigned long>(runtime.query->flags),
                request.maxResults);
#endif
    HRESULT hr = SearchServiceBroker::Query(
        request, &SearchServiceProgressThunk, &runtime, &SearchReadCancelThunk, &runtime, candidates, &stats, &SearchServiceCandidateBatchThunk, &runtime);
    if (FAILED(hr))
    {
        Debug::Warning(L"FileSystem::Search: service query failed root='{}' pattern='{}' mode={} flags=0x{:08X} hr=0x{:08X}",
                       runtime.rootPath,
                       runtime.namePattern,
                       static_cast<uint32_t>(runtime.query->nameMode),
                       static_cast<unsigned long>(runtime.query->flags),
                       static_cast<unsigned long>(hr));
        return hr;
    }

    if (stats.usedLiveScanFallback)
    {
        runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
    }

    const LocalSearchIndexCore::StoreState runtimeStoreState =
        stats.usedLiveScanFallback ? ResolveFallbackStoreState(stats.fallbackReason) : LocalSearchIndexCore::StoreState::Ready;
    const LocalSearchIndexCore::SyncPhase runtimeSyncPhase =
        stats.usedLiveScanFallback ? LocalSearchIndexCore::SyncPhase::Idle : LocalSearchIndexCore::SyncPhase::Watching;
    hr = ReportServiceStatusSnapshot(runtime, runtimeStoreState, runtimeSyncPhase, stats.queryExecutionMode, stats.fallbackReason, 0u, 0u, runtime.rootPath);
    if (FAILED(hr))
    {
        return IsCallbackCancellationResult(hr) ? kCancelledHr : hr;
    }

    ApplyIndexedQueryStats(runtime, stats);

#if defined(_DEBUG) || defined(__SANITIZE_ADDRESS__)
    Debug::Info(L"FileSystem::Search: service query completed root='{}' pattern='{}' dirs={} files={} candidates={} matches={} warnings=0x{:08X}",
                runtime.rootPath,
                runtime.namePattern,
                runtime.scannedDirectories,
                runtime.scannedFiles,
                runtime.candidateFiles,
                runtime.matchedEntries,
                runtime.warningFlags);
#endif

    servicePerf.SetValue0(runtime.matchedEntries);
    servicePerf.SetValue1(runtime.candidateFiles);
    return S_OK;
}

[[nodiscard]] bool IsServiceFallbackCandidate(HRESULT hr) noexcept
{
    switch (hr)
    {
        case kAccessDeniedHr:
        case HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED):
        case HRESULT_FROM_WIN32(ERROR_INVALID_FUNCTION):
        case HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND):
        case HRESULT_FROM_WIN32(ERROR_NO_DATA):
        case HRESULT_FROM_WIN32(ERROR_PIPE_NOT_CONNECTED):
        case HRESULT_FROM_WIN32(ERROR_PIPE_BUSY):
        case HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT):
        case HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE):
        case RPC_S_PROTOCOL_ERROR: return true;
    }

    return false;
}

// ---------------------------------------------------------------------------
// Parallel scan infrastructure
// ---------------------------------------------------------------------------

// Cancel thunk for parallel workers — checks a shared atomic flag
// instead of invoking the host callback (which is not thread-safe).
HRESULT STDMETHODCALLTYPE ParallelWorkerCancelThunk(void* cookie) noexcept
{
    if (cookie == nullptr)
    {
        return E_POINTER;
    }

    const auto* cancelFlag = static_cast<const std::atomic<bool>*>(cookie);
    return cancelFlag->load(std::memory_order_acquire) ? kCancelledHr : S_OK;
}

struct ParallelMatchEntry final
{
    SearchEntryMetadata metadata;
    SearchContentResult contentResult;
    uint32_t matchedBy = FILESYSTEM_SEARCH_MATCH_SOURCE_NONE;
};

struct ParallelChunkResult final
{
    std::vector<ParallelMatchEntry> matches;
    uint64_t filesScanned   = 0;
    uint64_t candidateFiles = 0;
    uint32_t warningFlags   = FILESYSTEM_SEARCH_WARNING_NONE;
    HRESULT status          = S_OK;
};

struct ParallelChunkWork final
{
    ParallelChunkWork()                                    = default;
    ParallelChunkWork(const ParallelChunkWork&)            = delete;
    ParallelChunkWork(ParallelChunkWork&&)                 = default;
    ParallelChunkWork& operator=(const ParallelChunkWork&) = delete;
    ParallelChunkWork& operator=(ParallelChunkWork&&)      = default;

    std::vector<SearchEntryMetadata> entries;
    // Read-only access to immutable search parameters in SearchRuntime.
    const SearchRuntime* runtime  = nullptr;
    std::atomic<bool>* cancelFlag = nullptr;
    wil::unique_hmodule modulePin;
    std::mutex* completionMutex           = nullptr;
    std::condition_variable* completionCv = nullptr;
    unsigned long* remainingChunks        = nullptr;
    ParallelChunkResult result;
};

// Worker-safe name matching — reads only immutable fields from runtime.
[[nodiscard]] bool WorkerMatchNamePattern(const SearchRuntime& runtime, const std::wstring& displayName) noexcept
{
    return MatchNamePattern(runtime, displayName);
}

// Worker-safe content matching — opens file independently, uses atomic cancel flag,
// does not invoke any host callbacks, does not modify runtime counters.
[[nodiscard]] HRESULT WorkerMatchFileContent(
    const SearchRuntime& runtime, std::atomic<bool>& cancelFlag, const std::wstring& fullPath, SearchContentResult& result, uint32_t& outWarnings) noexcept
{
    result = {};

    PathInfo pathInfo = MakePathInfo(fullPath);
    wil::com_ptr<IFileReader> reader;
    HRESULT hr = runtime.fileSystem->CreateFileReader(pathInfo.extended.c_str(), reader.put());
    if (FAILED(hr))
    {
        if (hr == kAccessDeniedHr)
        {
            outWarnings |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
            return S_OK;
        }

        return hr;
    }

    if (cancelFlag.load(std::memory_order_acquire))
    {
        return kCancelledHr;
    }

    SearchTextHelpers::TextSearchPattern pattern{};
    pattern.mode                   = runtime.query->contentMode;
    pattern.pattern                = runtime.contentPattern;
    pattern.compiledRegex          = runtime.contentRegex.get();
    pattern.caseSensitive          = runtime.matchCaseContent;
    pattern.literalChunkCharacters = SearchTextHelpers::kDefaultLiteralChunkChars;

    SearchTextHelpers::TextSearchResult helperResult{};
    hr = SearchTextHelpers::SearchFileReaderText(reader.get(),
                                                 pattern,
                                                 runtime.maxContentBytesPerFile,
                                                 0u,
                                                 runtime.maxSnippetCharacters,
                                                 runtime.wantSnippets,
                                                 &ParallelWorkerCancelThunk,
                                                 &cancelFlag,
                                                 helperResult);
    if (hr == kFileTooLargeHr)
    {
        return S_OK;
    }
    if (FAILED(hr))
    {
        if (hr == kAccessDeniedHr)
        {
            outWarnings |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
            return S_OK;
        }

        return hr;
    }

    result.matched     = helperResult.matched;
    result.byteOffset  = helperResult.matchOffset;
    result.byteLength  = helperResult.matchLength;
    result.previewText = std::move(helperResult.previewText);
    return S_OK;
}

void CALLBACK ParallelScanWorkerCallback(PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
{
    auto* chunk = static_cast<ParallelChunkWork*>(context);
    // Ensure completion is signaled even on early exit.
    auto completionGuard = wil::scope_exit([&]
    {
        if (chunk->completionMutex != nullptr && chunk->remainingChunks != nullptr)
        {
            {
                std::lock_guard lock(*chunk->completionMutex);
                if (*chunk->remainingChunks > 0u)
                {
                    --(*chunk->remainingChunks);
                }
            }

            if (chunk->completionCv != nullptr)
            {
                chunk->completionCv->notify_all();
            }
        }
    });
    static_cast<void>(chunk->modulePin); // Keep DLL loaded for callback lifetime.

    [[maybe_unused]] auto coInit = wil::CoInitializeEx(COINIT_MULTITHREADED);

    const SearchRuntime& runtime = *chunk->runtime;
    ParallelChunkResult& result  = chunk->result;

    for (auto& entry : chunk->entries)
    {
        if (chunk->cancelFlag->load(std::memory_order_acquire))
        {
            result.status = kCancelledHr;
            return;
        }

        const bool isDirectory = (entry.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        if ((isDirectory && ! runtime.includeDirectories) || (! isDirectory && ! runtime.includeFiles))
        {
            continue;
        }

        if (! isDirectory)
        {
            ++result.filesScanned;
        }

        const bool nameMatched = WorkerMatchNamePattern(runtime, entry.displayName);
        if (runtime.query->nameMode != FILESYSTEM_SEARCH_NAME_DISABLED && ! nameMatched)
        {
            continue;
        }

        SearchContentResult contentResult{};
        bool contentMatched = (runtime.query->contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED);
        if (! contentMatched)
        {
            if (isDirectory)
            {
                continue;
            }

            ++result.candidateFiles;
            const HRESULT hr = WorkerMatchFileContent(runtime, *chunk->cancelFlag, entry.fullPath, contentResult, result.warningFlags);
            if (FAILED(hr))
            {
                if (hr == kAccessDeniedHr)
                {
                    result.warningFlags |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
                    continue;
                }

                if (IsCallbackCancellationResult(hr))
                {
                    result.status = kCancelledHr;
                    return;
                }

                result.status = hr;
                return;
            }

            contentMatched = contentResult.matched;
        }

        if (! contentMatched)
        {
            continue;
        }

        uint32_t matchedBy = FILESYSTEM_SEARCH_MATCH_SOURCE_NONE;
        if (runtime.query->nameMode != FILESYSTEM_SEARCH_NAME_DISABLED)
        {
            matchedBy |= FILESYSTEM_SEARCH_MATCH_SOURCE_NAME;
        }
        if (runtime.query->contentMode != FILESYSTEM_SEARCH_CONTENT_DISABLED)
        {
            matchedBy |= FILESYSTEM_SEARCH_MATCH_SOURCE_CONTENT;
        }

        result.matches.push_back({std::move(entry), std::move(contentResult), matchedBy});
    }
}

// Dispatches entry evaluation to the threadpool in parallel chunks, then drains
// matched results back through the host callback on the calling thread.
// Returns S_OK on success, or the first failure HRESULT from a worker/callback.
HRESULT ParallelEvaluateEntries(SearchRuntime& runtime, std::vector<SearchEntryMetadata>& entries) noexcept
{
    Debug::Perf::Scope parallelPerf(L"FileSystem.Search.ParallelEvaluate");

    const auto entryCount          = static_cast<unsigned long>(entries.size());
    const unsigned long chunkCount = (entryCount + kParallelScanChunkSize - 1u) / kParallelScanChunkSize;

    std::atomic<bool> cancelFlag{false};

    // Check host cancellation once before dispatching workers.
    {
        const HRESULT cancelHr = CheckSearchCancelled(runtime);
        if (FAILED(cancelHr))
        {
            return cancelHr;
        }
    }

    // Build chunk work items. Main thread owns these objects; workers get raw pointers.
    std::vector<ParallelChunkWork> chunks(chunkCount);
    std::mutex completionMutex;
    std::condition_variable completionCv;
    unsigned long remainingChunks = chunkCount;

    unsigned long offset = 0;
    for (unsigned long i = 0; i < chunkCount; ++i)
    {
        const unsigned long remaining     = entryCount - offset;
        const unsigned long thisChunkSize = (std::min)(kParallelScanChunkSize, remaining);

        chunks[i].entries.reserve(thisChunkSize);
        for (unsigned long j = 0; j < thisChunkSize; ++j)
        {
            chunks[i].entries.push_back(std::move(entries[offset + j]));
        }
        offset += thisChunkSize;

        chunks[i].runtime         = &runtime;
        chunks[i].cancelFlag      = &cancelFlag;
        chunks[i].completionMutex = &completionMutex;
        chunks[i].completionCv    = &completionCv;
        chunks[i].remainingChunks = &remainingChunks;
        chunks[i].modulePin       = AcquireModuleReferenceFromAddress(&kFileSystemModuleAnchor);
    }

    // Dispatch chunks to the threadpool.
    for (unsigned long i = 0; i < chunkCount; ++i)
    {
        const BOOL queued = ::TrySubmitThreadpoolCallback(&ParallelScanWorkerCallback, &chunks[i], nullptr);
        if (queued == 0)
        {
            // Threadpool submission failed — evaluate this chunk sequentially on the calling thread.
            ParallelScanWorkerCallback(nullptr, &chunks[i]);
        }
    }

    // Wait for all workers to complete while continuing to honor host cancellation.
    HRESULT waitHr = S_OK;
    {
        std::unique_lock lock(completionMutex);
        while (remainingChunks > 0u)
        {
            const bool completed = completionCv.wait_for(lock, std::chrono::milliseconds{50}, [&] { return remainingChunks == 0u; });
            if (completed || FAILED(waitHr))
            {
                continue;
            }

            lock.unlock();
            const HRESULT cancelHr = CheckSearchCancelled(runtime);
            lock.lock();
            if (FAILED(cancelHr))
            {
                waitHr = cancelHr;
                cancelFlag.store(true, std::memory_order_release);
            }
        }
    }

    // Drain results: merge statistics, emit matches through the host callback.
    HRESULT firstError = S_OK;
    for (auto& chunk : chunks)
    {
        runtime.warningFlags |= chunk.result.warningFlags;
        runtime.scannedFiles += chunk.result.filesScanned;
        runtime.candidateFiles += chunk.result.candidateFiles;

        if (FAILED(chunk.result.status) && SUCCEEDED(firstError))
        {
            firstError = chunk.result.status;
        }
    }

    if (FAILED(waitHr) && SUCCEEDED(firstError))
    {
        firstError = waitHr;
    }

    // If any worker hit a fatal error (not cancellation), propagate it.
    if (FAILED(firstError) && ! IsCallbackCancellationResult(firstError))
    {
        return firstError;
    }

    // Stream matched results to the host callback (single-threaded, safe).
    for (auto& chunk : chunks)
    {
        for (auto& match : chunk.result.matches)
        {
            if (runtime.stopRequested)
            {
                return S_OK;
            }

            const HRESULT emitHr = EmitSearchMatch(runtime, match.metadata, match.matchedBy, match.contentResult);
            if (FAILED(emitHr))
            {
                return emitHr;
            }
        }
    }

    // If workers were cancelled, propagate cancellation.
    if (FAILED(firstError))
    {
        return firstError;
    }

    parallelPerf.SetValue0(static_cast<uint64_t>(entryCount));
    parallelPerf.SetValue1(static_cast<uint64_t>(chunkCount));
    return S_OK;
}

HRESULT SearchDirectoryTree(SearchRuntime& runtime) noexcept
{
    Debug::Perf::Scope treePerf(L"FileSystem.Search.ScanTree");
    treePerf.SetDetail(runtime.rootPath);

    struct DirectoryFrame final
    {
        std::wstring fullPath;
        std::wstring relativeBase;
    };

    std::vector<DirectoryFrame> stack;
    stack.push_back({runtime.rootPath, std::wstring()});
    static_cast<void>(MarkQueuedDirectory(runtime, runtime.rootPath));

    while (! stack.empty() && ! runtime.stopRequested)
    {
        const DirectoryFrame frame = std::move(stack.back());
        stack.pop_back();

        ++runtime.scannedDirectories;

        HRESULT hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_ENUMERATING, &frame.fullPath, S_OK, false);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = CheckSearchCancelled(runtime);
        if (FAILED(hr))
        {
            return hr;
        }

        wil::com_ptr<IFilesInformation> information;
        hr = runtime.fileSystem->ReadDirectoryInfo(frame.fullPath.c_str(), information.put());
        if (FAILED(hr))
        {
            if (hr == kAccessDeniedHr && frame.fullPath != runtime.rootPath)
            {
                runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
                continue;
            }

            return hr;
        }

        unsigned long count = 0;
        hr                  = information->GetCount(&count);
        if (FAILED(hr))
        {
            return hr;
        }

        // Collect metadata for all entries in this directory. This is needed regardless of
        // the parallel path because we must separate directory entries (for DFS stack) from
        // file entries (for evaluation). Metadata construction is cheap — no file I/O.
        std::vector<SearchEntryMetadata> allEntries;
        allEntries.reserve(count);

        for (unsigned long index = 0; index < count; ++index)
        {
            FileInfo* entry = nullptr;
            hr              = information->Get(index, &entry);
            if (FAILED(hr) || entry == nullptr)
            {
                return FAILED(hr) ? hr : E_FAIL;
            }

            const std::wstring_view name(entry->FileName, entry->FileNameSize / sizeof(wchar_t));
            if (IsDotOrDotDot(name))
            {
                continue;
            }

            SearchEntryMetadata metadata{};
            metadata.displayName = std::wstring(name);
            metadata.relativePath =
                frame.relativeBase.empty() ? metadata.displayName : NormalizeSearchPath(AppendPath(frame.relativeBase, metadata.displayName));
            metadata.fullPath       = NormalizeSearchPath(AppendPath(frame.fullPath, metadata.displayName));
            metadata.fileAttributes = entry->FileAttributes;
            metadata.creationTime   = entry->CreationTime;
            metadata.lastAccessTime = entry->LastAccessTime;
            metadata.lastWriteTime  = entry->LastWriteTime;
            metadata.changeTime     = entry->ChangeTime;
            metadata.endOfFile      = entry->EndOfFile;
            metadata.allocationSize = entry->AllocationSize;

            allEntries.push_back(std::move(metadata));
        }

        // Release the enumeration buffer before evaluation — entries are fully materialized.
        information.reset();

        // Push subdirectories onto the DFS stack (must happen on the main thread
        // because MarkQueuedDirectory mutates runtime.queuedDirectories).
        for (const auto& metadata : allEntries)
        {
            const bool isDirectory = (metadata.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            const bool isReparse   = (metadata.fileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
            if (runtime.recursive && isDirectory && (! isReparse || runtime.followSymlinks))
            {
                if (MarkQueuedDirectory(runtime, metadata.fullPath))
                {
                    stack.push_back({metadata.fullPath, metadata.relativePath});
                }
            }
        }

        // Parallel path: dispatch evaluation to the threadpool when the directory
        // is large enough that the per-chunk overhead is worth the parallelism.
        const auto entryCount           = static_cast<unsigned long>(allEntries.size());
        const bool contentSearchEnabled = runtime.query->contentMode != FILESYSTEM_SEARCH_CONTENT_DISABLED;
        if (contentSearchEnabled && entryCount >= kParallelScanThreshold)
        {
            hr = ParallelEvaluateEntries(runtime, allEntries);
            if (FAILED(hr))
            {
                return hr;
            }
        }
        else
        {
            // Sequential path: evaluate entries directly on the calling thread.
            for (auto& metadata : allEntries)
            {
                if (runtime.stopRequested)
                {
                    break;
                }

                hr = EvaluateEntry(runtime, metadata);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
        }
    }

    treePerf.SetValue0(runtime.scannedDirectories);
    treePerf.SetValue1(runtime.scannedFiles);
    return S_OK;
}
} // namespace

HRESULT STDMETHODCALLTYPE FileSystem::Search(const FileSystemSearchQuery* query, IFileSystemSearchCallback* callback, void* cookie) noexcept
{
    if (query == nullptr || callback == nullptr)
    {
        return E_POINTER;
    }

    if (query->sizeBytes != sizeof(FileSystemSearchQuery))
    {
        return E_INVALIDARG;
    }

    if (query->rootPath == nullptr || query->rootPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const FileSystemSearchHostExtensions* hostExtensions = nullptr;
    void* callbackCookie                                 = cookie;
    if (query->reserved != 0u)
    {
        if (query->reserved != FILESYSTEM_SEARCH_HOST_EXTENSIONS_V1 || cookie == nullptr)
        {
            return E_INVALIDARG;
        }

        hostExtensions = static_cast<const FileSystemSearchHostExtensions*>(cookie);
        if (hostExtensions->sizeBytes != sizeof(FileSystemSearchHostExtensions) || hostExtensions->version != FILESYSTEM_SEARCH_HOST_EXTENSIONS_V1)
        {
            return E_INVALIDARG;
        }

        callbackCookie = hostExtensions->callbackCookie;
    }

    const bool includeFiles       = (query->flags & FILESYSTEM_SEARCH_INCLUDE_FILES) != 0;
    const bool includeDirectories = (query->flags & FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES) != 0;
    if (! includeFiles && ! includeDirectories)
    {
        return E_INVALIDARG;
    }

    if (query->nameMode == FILESYSTEM_SEARCH_NAME_DISABLED && query->contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED)
    {
        return E_INVALIDARG;
    }

    if (query->nameMode != FILESYSTEM_SEARCH_NAME_DISABLED && query->namePattern == nullptr)
    {
        return E_INVALIDARG;
    }

    if (query->contentMode != FILESYSTEM_SEARCH_CONTENT_DISABLED)
    {
        if (query->contentPattern == nullptr || ! includeFiles)
        {
            return E_INVALIDARG;
        }
    }

    FileSystemSearchBackendPreference backendPreference = kDefaultSearchBackendPreference;
    unsigned int searchMaxDirectoryWalkers              = 4u;
    std::shared_ptr<LocalSearchIndexCore::Repository> searchIndexRepository;
    const std::wstring configuredServicePipeName = SearchServiceBroker::GetConfiguredPipeName();
    bool serviceTemporarilyUnavailable           = false;
    {
        std::lock_guard lock(_stateMutex);
        backendPreference             = _searchBackendPreference;
        searchMaxDirectoryWalkers     = _searchMaxDirectoryWalkers;
        searchIndexRepository         = _searchIndexRepository;
        serviceTemporarilyUnavailable = ! _searchServiceUnavailablePipeName.empty() && _searchServiceUnavailablePipeName == configuredServicePipeName &&
                                        ::GetTickCount64() < _searchServiceUnavailableUntilTick;
    }

    try
    {
        Debug::Perf::Scope searchPerf(L"FileSystem.Search");

        SearchRuntime runtime{};
        runtime.fileSystem                   = this;
        runtime.query                        = query;
        runtime.callback                     = callback;
        runtime.hostExtensions               = hostExtensions;
        runtime.cookie                       = callbackCookie;
        runtime.rootPath                     = NormalizeSearchPath(MakeAbsolutePath(query->rootPath));
        runtime.namePattern                  = query->namePattern ? std::wstring(query->namePattern) : std::wstring();
        runtime.contentPattern               = query->contentPattern ? std::wstring(query->contentPattern) : std::wstring();
        runtime.includeFiles                 = includeFiles;
        runtime.includeDirectories           = includeDirectories;
        runtime.recursive                    = (query->flags & FILESYSTEM_SEARCH_RECURSIVE) != 0;
        runtime.followSymlinks               = (query->flags & FILESYSTEM_SEARCH_FOLLOW_SYMLINKS) != 0;
        runtime.wantSnippets                 = (query->flags & FILESYSTEM_SEARCH_WANT_SNIPPETS) != 0;
        runtime.matchCaseName                = (query->flags & FILESYSTEM_SEARCH_MATCH_CASE_NAME) != 0;
        runtime.matchCaseContent             = (query->flags & FILESYSTEM_SEARCH_MATCH_CASE_CONTENT) != 0;
        runtime.maxResults                   = query->maxResults;
        runtime.maxContentBytesPerFile       = query->maxContentBytesPerFile != 0 ? query->maxContentBytesPerFile : kDefaultSearchContentBytesPerFile;
        runtime.maxSnippetCharacters         = query->maxSnippetCharacters != 0 ? query->maxSnippetCharacters : kDefaultSearchSnippetChars;
        runtime.parallelDirectoryWalkWorkers = searchMaxDirectoryWalkers;

        if (runtime.rootPath.empty())
        {
            runtime.rootPath = NormalizeSearchPath(query->rootPath);
        }

        LocalSearchIndexCore::SupportInfo indexSupport{};
        if (searchIndexRepository)
        {
            const HRESULT probeHr = searchIndexRepository->ProbePath(runtime.rootPath, indexSupport);
            if (FAILED(probeHr))
            {
                indexSupport = {};
            }
        }

        const bool autoBackendPreference              = backendPreference == FileSystemSearchBackendPreference::Auto;
        const bool preferServiceBackend               = indexSupport.indexable &&
                                                        (autoBackendPreference || backendPreference == FileSystemSearchBackendPreference::Service) &&
                                                        ! serviceTemporarilyUnavailable;
        const SearchBackendSelection backendSelection = SelectSearchBackend(backendPreference, query->flags, indexSupport, preferServiceBackend);
        runtime.backend                               = backendSelection.backend;
        runtime.warningFlags |= backendSelection.warningFlags;
        if (serviceTemporarilyUnavailable && indexSupport.indexable &&
            (autoBackendPreference || backendPreference == FileSystemSearchBackendPreference::Service))
        {
            runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE;
        }

        Debug::Info(
            L"FileSystem::Search: root='{}' preference='{}' selected='{}' preferService={} indexable={} fsKind={} flags=0x{:08X} warnings=0x{:08X} walkers={}",
            runtime.rootPath,
            BackendPreferenceToString(backendPreference),
            SearchBackendToString(runtime.backend),
            preferServiceBackend,
            indexSupport.indexable,
            static_cast<uint32_t>(indexSupport.fileSystemKind),
            static_cast<uint32_t>(query->flags),
            runtime.warningFlags,
            runtime.parallelDirectoryWalkWorkers);

        if (query->nameMode == FILESYSTEM_SEARCH_NAME_REGEX)
        {
            std::wstring rejectReason;
            if (! ValidateRegexPatternSafety(runtime.namePattern, rejectReason))
            {
                return RejectRegexSearch(runtime, L"name", rejectReason);
            }

            const auto flags = runtime.matchCaseName ? std::regex_constants::ECMAScript : (std::regex_constants::ECMAScript | std::regex_constants::icase);
            try
            {
                runtime.nameRegex = g_regexCache.GetOrCompile(runtime.namePattern, flags);
            }
            catch (const std::regex_error&)
            {
                return RejectRegexSearch(runtime, L"name", L"Invalid regex syntax.");
            }
        }

        if (query->contentMode == FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX)
        {
            std::wstring rejectReason;
            if (! ValidateRegexPatternSafety(runtime.contentPattern, rejectReason))
            {
                return RejectRegexSearch(runtime, L"content", rejectReason);
            }

            const auto flags = runtime.matchCaseContent ? std::regex_constants::ECMAScript : (std::regex_constants::ECMAScript | std::regex_constants::icase);
            try
            {
                runtime.contentRegex = g_regexCache.GetOrCompile(runtime.contentPattern, flags);
            }
            catch (const std::regex_error&)
            {
                return RejectRegexSearch(runtime, L"content", L"Invalid regex syntax.");
            }
        }

        HRESULT hr = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_INITIALIZING, &runtime.rootPath, S_OK, true);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = CheckSearchCancelled(runtime);
        if (FAILED(hr))
        {
            return hr;
        }

        unsigned long rootAttributes = 0;
        hr                           = GetAttributes(runtime.rootPath.c_str(), &rootAttributes);
        if (FAILED(hr))
        {
            return hr;
        }

        if (runtime.backend == FILESYSTEM_SEARCH_BACKEND_SERVICE)
        {
            hr = SearchServiceTree(runtime);
            if (SUCCEEDED(hr))
            {
                std::lock_guard lock(_stateMutex);
                if (_searchServiceUnavailablePipeName == configuredServicePipeName)
                {
                    _searchServiceUnavailablePipeName.clear();
                    _searchServiceUnavailableUntilTick = 0u;
                }
            }
            else if (IsServiceFallbackCandidate(hr))
            {
                Debug::Warning(
                    L"FileSystem::Search: service backend failed root='{}' hr=0x{:08X}; falling back.", runtime.rootPath, static_cast<unsigned long>(hr));
                runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE;
                {
                    std::lock_guard lock(_stateMutex);
                    _searchServiceUnavailablePipeName  = configuredServicePipeName;
                    _searchServiceUnavailableUntilTick = ::GetTickCount64() + kSearchServiceUnavailableRetryCooldownMs;
                }
                if (searchIndexRepository && indexSupport.indexable)
                {
                    runtime.backend = FILESYSTEM_SEARCH_BACKEND_INDEX;
                }
                else
                {
                    runtime.backend = FILESYSTEM_SEARCH_BACKEND_SCAN;
                    runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
                }

                runtime.scannedDirectories        = 0u;
                runtime.scannedFiles              = 0u;
                runtime.candidateFiles            = 0u;
                runtime.indexedCandidatesConsumed = 0u;
                runtime.matchedEntries            = 0u;
                runtime.usingIndexedEnumeration   = false;
                hr                                = S_OK;
            }
        }

        if (runtime.backend == FILESYSTEM_SEARCH_BACKEND_INDEX && searchIndexRepository)
        {
            hr = SearchIndexedTree(runtime, *searchIndexRepository);
            if (FAILED(hr) && (hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) ||
                               hr == HRESULT_FROM_WIN32(ERROR_INVALID_FUNCTION)))
            {
                Debug::Warning(
                    L"FileSystem::Search: indexed backend failed root='{}' hr=0x{:08X}; degrading to scan.", runtime.rootPath, static_cast<unsigned long>(hr));
                runtime.backend = FILESYSTEM_SEARCH_BACKEND_SCAN;
                runtime.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
                runtime.scannedDirectories        = 0u;
                runtime.scannedFiles              = 0u;
                runtime.candidateFiles            = 0u;
                runtime.indexedCandidatesConsumed = 0u;
                runtime.usingIndexedEnumeration   = false;
                hr                                = S_OK;
            }
        }

        if (SUCCEEDED(hr))
        {
            if ((rootAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 && runtime.backend == FILESYSTEM_SEARCH_BACKEND_SCAN)
            {
                hr = ShouldUseParallelDirectoryWalk(runtime) ? SearchDirectoryTreeParallelNameOnly(runtime) : SearchDirectoryTree(runtime);
            }
            else if (runtime.backend == FILESYSTEM_SEARCH_BACKEND_SCAN)
            {
                SearchEntryMetadata entry{};
                PopulateSinglePathMetadata(runtime.rootPath, rootAttributes, entry);
                hr = EvaluateEntry(runtime, entry);
            }
        }

        const HRESULT finalStatus = FAILED(hr) ? hr : (runtime.stopRequested ? S_FALSE : S_OK);
        const HRESULT progressHr  = ReportSearchProgress(runtime, FILESYSTEM_SEARCH_PHASE_COMPLETED, nullptr, finalStatus, true);
        if (FAILED(progressHr))
        {
            return progressHr;
        }

        searchPerf.SetDetail(runtime.rootPath);
        searchPerf.SetValue0(runtime.matchedEntries);
        searchPerf.SetValue1(runtime.scannedFiles + runtime.scannedDirectories);
        searchPerf.SetHr(FAILED(hr) ? hr : finalStatus);

        Debug::Info(
            L"FileSystem::Search: completed root='{}' backend='{}' status=0x{:08X} matched={} candidates={} scannedFiles={} scannedDirs={} warnings=0x{:08X}",
            runtime.rootPath,
            SearchBackendToString(runtime.backend),
            static_cast<unsigned long>(FAILED(hr) ? hr : finalStatus),
            runtime.matchedEntries,
            runtime.candidateFiles,
            runtime.scannedFiles,
            runtime.scannedDirectories,
            runtime.warningFlags);

        if (runtime.stopRequested && SUCCEEDED(hr))
        {
            return S_OK;
        }

        return hr;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::regex_error&)
    {
        Debug::Warning(L"FileSystem::Search: regex compilation failed (invalid pattern).");
        return E_INVALIDARG;
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FileSystem: Search failed with an unexpected std::exception.");
        return E_FAIL;
    }
}
