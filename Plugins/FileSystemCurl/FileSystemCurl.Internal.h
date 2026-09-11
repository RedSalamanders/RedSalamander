#pragma once

#include "FileSystemCurl.ImapHelpers.h"
#include "FileSystemCurl.h"
#include "PlugInterfaces/Host.h"

#include "DeleteOnCloseTemporaryFile.h"
#include "Helpers.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <charconv>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <format>
#include <functional>
#include <limits>
#include <memory>
#include <mutex>
#include <new>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include <curl/curl.h>

namespace FileSystemCurlInternal
{
using Protocol = FileSystemCurlProtocol;

struct CurlSlistDeleter
{
    void operator()(curl_slist* list) const noexcept
    {
        if (list)
        {
            curl_slist_free_all(list);
        }
    }
};

using unique_curl_slist = std::unique_ptr<curl_slist, CurlSlistDeleter>;

struct CurlEasyDeleter
{
    void operator()(CURL* handle) const noexcept
    {
        if (handle)
        {
            curl_easy_cleanup(handle);
        }
    }
};

using unique_curl_easy = std::unique_ptr<CURL, CurlEasyDeleter>;

class CurlEasyPool
{
    struct MultiDeleter final
    {
        void operator()(CURLM* value) const noexcept
        {
            static_cast<void>(curl_multi_cleanup(value));
        }
    };

    struct Handles final
    {
        Handles()                              = default;
        Handles(const Handles&)                = delete;
        Handles& operator=(const Handles&)     = delete;
        Handles(Handles&&) noexcept            = default;
        Handles& operator=(Handles&&) noexcept = default;
        // Idle cursor handles are detached. Destroy both easy handles before the
        // retained multi cache; all owners move together under one exclusive borrow.
        std::unique_ptr<CURLM, MultiDeleter> multi;
        unique_curl_easy easy;
        unique_curl_easy cursorEasy;
    };

    CurlEasyPool(const CurlEasyPool&)            = delete;
    CurlEasyPool(CurlEasyPool&&)                 = delete;
    CurlEasyPool& operator=(const CurlEasyPool&) = delete;
    CurlEasyPool& operator=(CurlEasyPool&&)      = delete;

public:
    CurlEasyPool() = default;

    class BorrowedHandle
    {
    public:
        BorrowedHandle() noexcept = default;
        BorrowedHandle(CurlEasyPool* pool, std::wstring key, Handles handles) noexcept : _pool(pool), _key(std::move(key)), _handles(std::move(handles))
        {
        }

        ~BorrowedHandle()
        {
            if (_pool && _handles.easy)
            {
                _pool->ReturnHandle(std::move(_key), std::move(_handles));
            }
        }

        BorrowedHandle(const BorrowedHandle&)            = delete;
        BorrowedHandle& operator=(const BorrowedHandle&) = delete;

        BorrowedHandle(BorrowedHandle&& other) noexcept
            : _pool(std::exchange(other._pool, nullptr)),
              _key(std::move(other._key)),
              _handles(std::move(other._handles))
        {
        }

        BorrowedHandle& operator=(BorrowedHandle&& other) noexcept
        {
            if (this != &other)
            {
                if (_pool && _handles.easy)
                {
                    _pool->ReturnHandle(std::move(_key), std::move(_handles));
                }
                _pool    = std::exchange(other._pool, nullptr);
                _key     = std::move(other._key);
                _handles = std::move(other._handles);
            }
            return *this;
        }

        [[nodiscard]] CURL* get() const noexcept
        {
            return _handles.easy.get();
        }

        // Worker-exclusive, lazily allocated. The caller must detach its easy
        // handle before returning this borrow; no multi callbacks escape it.
        [[nodiscard]] CURLM* GetOrCreateMulti() noexcept
        {
            if (! _handles.multi)
            {
                _handles.multi.reset(curl_multi_init());
            }
            return _handles.multi.get();
        }

        // Never attach get() to a caller-owned multi: libcurl then destroys its
        // private curl_easy_perform cache. Keep cursor and blocking identities apart.
        [[nodiscard]] CURL* GetOrCreateCursorEasy() noexcept
        {
            if (! _handles.cursorEasy)
            {
                _handles.cursorEasy.reset(curl_easy_init());
            }
            return _handles.cursorEasy.get();
        }

        explicit operator bool() const noexcept
        {
            return _handles.easy != nullptr;
        }

    private:
        CurlEasyPool* _pool = nullptr;
        std::wstring _key;
        Handles _handles;
    };

    [[nodiscard]] BorrowedHandle Borrow(std::wstring_view limiterKey) noexcept;
    void BeginShutdown() noexcept;
    [[nodiscard]] bool CanUnloadNow() noexcept;

private:
    struct IdleEntry
    {
        IdleEntry(Handles h, uint64_t ts) noexcept : handles(std::move(h)), returnedAtMs(ts)
        {
        }

        IdleEntry(const IdleEntry&)            = delete;
        IdleEntry& operator=(const IdleEntry&) = delete;
        IdleEntry(IdleEntry&&)                 = default;
        IdleEntry& operator=(IdleEntry&&)      = default;

        Handles handles;
        uint64_t returnedAtMs = 0;
    };

    static constexpr size_t kMaxIdlePerConnection = 4;
    static constexpr uint64_t kIdleExpiryMs       = 60000;

    void ReturnHandle(std::wstring key, Handles handles) noexcept;
    void CancelBorrow() noexcept;
    void EvictExpired(uint64_t now, std::vector<Handles>& cleanup) noexcept;

    std::mutex _mutex;
    std::unordered_map<std::wstring, std::vector<IdleEntry>> _idle;
    size_t _activeBorrowCount = 0u;
    bool _shutdownRequested   = false;
};

[[nodiscard]] CurlEasyPool& GetCurlEasyPool() noexcept;

struct ArenaOwner
{
    ArenaOwner() = default;

    ~ArenaOwner()
    {
        DestroyFileSystemArena(&_arena);
    }

    ArenaOwner(const ArenaOwner&)            = delete;
    ArenaOwner(ArenaOwner&&)                 = delete;
    ArenaOwner& operator=(const ArenaOwner&) = delete;
    ArenaOwner& operator=(ArenaOwner&&)      = delete;

    HRESULT Initialize(unsigned long capacityBytes) noexcept
    {
        DestroyFileSystemArena(&_arena);
        _arena = {};
        return InitializeFileSystemArena(&_arena, capacityBytes);
    }

    FileSystemArena* Get() noexcept
    {
        return &_arena;
    }

    FileSystemArena _arena{};
};

struct ConnectionInfo
{
    Protocol protocol = Protocol::Sftp;

    bool fromConnectionManagerProfile = false;
    std::wstring connectionName;
    std::wstring connectionId;
    std::wstring connectionAuthMode;
    bool connectionSavePassword = false;
    bool connectionRequireHello = false;

    // Effective per-connection caps after applying ConnectionProfile.extra overrides.
    // Used both for scheduling and for the global per-connection limiter.
    unsigned int effectiveCopyMoveMaxConcurrency = 1;
    unsigned int effectiveDeleteMaxConcurrency   = 1;
    std::wstring limiterKey;

    std::string host;
    std::optional<unsigned int> port;
    std::string user;
    std::string password;
    std::string basePath;
    std::wstring basePathWide;

    bool ftpUseEpsv                  = true;
    unsigned long connectTimeoutMs   = 10000;
    unsigned long operationTimeoutMs = 0;
    bool ignoreSslTrust              = false;

    std::string sshPrivateKey;
    std::string sshPublicKey;
    std::string sshKeyPassphrase;
    std::string sshKnownHosts;

#ifdef ENABLE_TESTS
    // Synchronous IMAP orchestration fixture, not a transport/cancellation substitute.
    // The caller owns the context for this ConnectionInfo's entire call lifetime.
    HRESULT (*imapRequestForSelfTest)(void*, std::wstring_view, std::string_view, std::string&) noexcept = nullptr;
    void* imapRequestContextForSelfTest = nullptr;
#endif
};

struct ResolvedLocation
{
    ConnectionInfo connection;
    std::wstring remotePath;
};

#ifdef ENABLE_TESTS
struct CurlFailedConnectPorts final
{
    int remote = 0;
    int local  = 0;
};
#endif

// Native FTP/SSH traversal cursor. Unlike the buffered folder-view facade, this
// owns a bounded transport chunk and one line, and never replays emitted entries.
// Open/Next/destruction stay on the calling worker; the callback is borrowed only
// for that cursor lifetime. S_FALSE means a complete, successfully ended listing.
// Native Move preflight memory (plan C0 item 1): the whole-tree membership and per-file size proof are
// kept as 64-bit digests of the normalized member paths instead of one std::wstring per entry, so a
// tree of N files retains 16 * N bytes (plus 8 per directory) regardless of path length. Both
// containers are filled during preflight, finalized once, and then read concurrently by the copy and
// delete walkers (no mutation after Finalize). A late writer whose new name collides on 64 bits and,
// for a file, also matches the committed size would pass as a member; that probability is ~n^2 / 2^65.
[[nodiscard]] uint64_t DigestPluginPath(std::wstring_view normalizedPath) noexcept;

class SourceTreeDigestSet final
{
public:
    void Insert(std::wstring_view normalizedPath);
    void Finalize() noexcept;
    [[nodiscard]] bool Contains(std::wstring_view normalizedPath) const noexcept;
    [[nodiscard]] size_t size() const noexcept { return _digests.size(); }
    [[nodiscard]] uint64_t RetainedBytes() const noexcept { return static_cast<uint64_t>(_digests.capacity()) * sizeof(uint64_t); }
    [[nodiscard]] bool IsFinalized() const noexcept { return _finalized; }

private:
    std::vector<uint64_t> _digests;
    bool _finalized = false;
};

class SourceSizeCommitmentDigestMap final
{
public:
    void Insert(std::wstring_view normalizedPath, uint64_t sizeBytes);
    void Finalize() noexcept;
    [[nodiscard]] std::optional<uint64_t> Find(std::wstring_view normalizedPath) const noexcept;
    [[nodiscard]] size_t size() const noexcept { return _entries.size(); }
    [[nodiscard]] uint64_t RetainedBytes() const noexcept { return static_cast<uint64_t>(_entries.capacity()) * sizeof(std::pair<uint64_t, uint64_t>); }
    [[nodiscard]] bool IsFinalized() const noexcept { return _finalized; }

private:
    std::vector<std::pair<uint64_t, uint64_t>> _entries;
    bool _finalized = false;
};

class CurlDirectoryCursor final
{
public:
    CurlDirectoryCursor() noexcept;
    ~CurlDirectoryCursor();
    CurlDirectoryCursor(const CurlDirectoryCursor&)            = delete;
    CurlDirectoryCursor& operator=(const CurlDirectoryCursor&) = delete;
    CurlDirectoryCursor(CurlDirectoryCursor&&)                 = delete;
    CurlDirectoryCursor& operator=(CurlDirectoryCursor&&)      = delete;

    static constexpr uint64_t kMetadataReservationBytes = 3u * CURL_MAX_WRITE_SIZE + 2048u;
    [[nodiscard]] HRESULT Open(const ConnectionInfo& conn, std::wstring_view path, std::function<HRESULT()> checkpoint) noexcept;
    // Empty filter materializes all timestamps (ordinary traversal). A lookup
    // still validates every row, but only needs the requested occupant's time.
    [[nodiscard]] HRESULT Next(FilesInformationCurl::Entry& entry, std::wstring_view timestampLeaf = {}) noexcept;
    [[nodiscard]] uint64_t RetainedPathBytes() const noexcept;

#ifdef ENABLE_TESTS
    // Feeds exact callback chunks through the production consumer, without I/O.
    // A lookup leaf retains only matching rows and counts every validated row,
    // matching GetEntryInfo's consumer without sockets or fixture LIST generation.
    [[nodiscard]] static HRESULT ParseChunksForSelfTest(std::span<const std::string_view> chunks,
                                                        std::vector<FilesInformationCurl::Entry>& entries,
                                                        std::wstring_view lookupLeaf = {},
                                                        uint64_t* inspectedRows = nullptr) noexcept;
    [[nodiscard]] static std::optional<CurlFailedConnectPorts> ParseFailedConnectTraceForSelfTest(curl_infotype type, std::string_view text) noexcept;
    [[nodiscard]] std::optional<CurlFailedConnectPorts> FailedConnectPortsForSelfTest() const noexcept;
    [[nodiscard]] bool HasActiveTransferForSelfTest() const noexcept;
#endif

private:
    struct Impl;
    std::unique_ptr<Impl> _impl;
};

[[nodiscard]] bool HasFlag(FileSystemFlags flags, FileSystemFlags flag) noexcept;
[[nodiscard]] HRESULT NormalizeCancellation(HRESULT hr) noexcept;

// Module anchor for AcquireModuleReferenceFromAddress — keeps the DLL loaded while worker threads are running.
extern const int kFileSystemCurlModuleAnchor;

// Stops and joins the shared background copy/move worker threads.
// Intended to be invoked at a host "quiet point" when the last FileSystemCurl instance is being destroyed.
void ShutdownSharedCopyMoveJobScheduler() noexcept;

[[nodiscard]] inline bool IsAuthenticationFailureHr(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD) || hr == HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
}

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept;
[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept;

[[nodiscard]] std::wstring ProtocolToDisplay(Protocol protocol);

[[nodiscard]] std::wstring NormalizePluginPath(std::wstring_view rawPath) noexcept;
[[nodiscard]] std::wstring_view TrimTrailingSlash(std::wstring_view path) noexcept;
[[nodiscard]] std::wstring_view LeafName(std::wstring_view path) noexcept;
[[nodiscard]] std::wstring ParentPath(std::wstring_view path) noexcept;
[[nodiscard]] std::wstring EnsureTrailingSlash(std::wstring_view path) noexcept;
[[nodiscard]] std::wstring EnsureTrailingSlashDisplay(std::wstring_view path) noexcept;
[[nodiscard]] std::wstring JoinPluginPath(std::wstring_view folder, std::wstring_view leaf) noexcept;
[[nodiscard]] std::wstring JoinPluginPathWide(std::wstring_view basePath, std::wstring_view pluginPath) noexcept;
[[nodiscard]] std::wstring JoinDisplayPath(std::wstring_view folder, std::wstring_view leaf) noexcept;
[[nodiscard]] std::wstring BuildDisplayPath(Protocol protocol, std::wstring_view pluginPath) noexcept;

[[nodiscard]] std::string EscapeUrlPath(std::wstring_view path) noexcept;
[[nodiscard]] std::string TrimAscii(std::string_view text) noexcept;
[[nodiscard]] bool EqualsAsciiIgnoreCase(std::string_view left, std::string_view right) noexcept;

[[nodiscard]] inline bool IsDotOrDotDotName(std::wstring_view name) noexcept
{
    while (! name.empty() && (name.front() == L' ' || name.front() == L'\t' || name.front() == L'\r' || name.front() == L'\n'))
    {
        name.remove_prefix(1);
    }
    while (! name.empty() && (name.back() == L' ' || name.back() == L'\t' || name.back() == L'\r' || name.back() == L'\n'))
    {
        name.remove_suffix(1);
    }
    while (! name.empty() && (name.back() == L'/' || name.back() == L'\\'))
    {
        name.remove_suffix(1);
    }
    return name == L"." || name == L"..";
}

[[nodiscard]] inline bool IsDotOrDotDotName(std::string_view name) noexcept
{
    while (! name.empty() && (name.front() == ' ' || name.front() == '\t' || name.front() == '\r' || name.front() == '\n'))
    {
        name.remove_prefix(1);
    }
    while (! name.empty() && (name.back() == ' ' || name.back() == '\t' || name.back() == '\r' || name.back() == '\n'))
    {
        name.remove_suffix(1);
    }
    while (! name.empty() && (name.back() == '/' || name.back() == '\\'))
    {
        name.remove_suffix(1);
    }
    return name == "." || name == "..";
}

[[nodiscard]] HRESULT ParseDirectoryListing(std::string_view listing, std::vector<FilesInformationCurl::Entry>& out) noexcept;
[[nodiscard]] std::optional<FilesInformationCurl::Entry> FindEntryByName(const std::vector<FilesInformationCurl::Entry>& entries,
                                                                         std::wstring_view leaf) noexcept;

[[nodiscard]] HRESULT ResolveLocation(Protocol protocol,
                                      const FileSystemCurl::Settings& settings,
                                      std::wstring_view pluginPath,
                                      IHostConnections* hostConnections,
                                      bool acquireSecrets,
                                      ResolvedLocation& out) noexcept;

// Resolves a location and runs an operation. On authentication failures for Connection Manager profiles, the helper:
// - for FTP anonymous rejection: asks the host to upgrade the profile to password auth and retries once,
// - for session-only secrets: clears the cached secret and retries once (this triggers a reprompt on the next resolve).
template <typename Func>
[[nodiscard]] HRESULT ResolveLocationWithAuthRetry(Protocol protocol,
                                                   const FileSystemCurl::Settings& settings,
                                                   std::wstring_view pluginPath,
                                                   IHostConnections* hostConnections,
                                                   bool acquireSecrets,
                                                   Func&& operation) noexcept
{
    ResolvedLocation resolved{};
    HRESULT hr = ResolveLocation(protocol, settings, pluginPath, hostConnections, acquireSecrets, resolved);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = operation(resolved);
    if (! IsAuthenticationFailureHr(hr) || ! hostConnections || ! resolved.connection.fromConnectionManagerProfile ||
        resolved.connection.connectionName.empty())
    {
        return hr;
    }

    if (protocol == Protocol::Ftp && resolved.connection.connectionAuthMode == L"anonymous")
    {
        const HRESULT upgradeHr = hostConnections->UpgradeFtpAnonymousToPassword(resolved.connection.connectionName.c_str(), nullptr);
        if (upgradeHr == S_FALSE)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (FAILED(upgradeHr))
        {
            return upgradeHr;
        }

        ResolvedLocation retryResolved{};
        const HRESULT resolveRetryHr = ResolveLocation(protocol, settings, pluginPath, hostConnections, acquireSecrets, retryResolved);
        if (FAILED(resolveRetryHr))
        {
            return resolveRetryHr;
        }
        return operation(retryResolved);
    }

    if (resolved.connection.connectionSavePassword)
    {
        return hr;
    }

    const HostConnectionSecretKind secretKind =
        (resolved.connection.connectionAuthMode == L"sshKey") ? HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE : HOST_CONNECTION_SECRET_PASSWORD;
    static_cast<void>(hostConnections->ClearCachedConnectionSecret(resolved.connection.connectionName.c_str(), secretKind));

    ResolvedLocation retryResolved{};
    const HRESULT resolveRetryHr = ResolveLocation(protocol, settings, pluginPath, hostConnections, acquireSecrets, retryResolved);
    if (FAILED(resolveRetryHr))
    {
        return resolveRetryHr;
    }
    return operation(retryResolved);
}

HRESULT EnsureCurlInitialized() noexcept;
void BeginFileSystemCurlShutdown() noexcept;
[[nodiscard]] bool CanUnloadFileSystemCurlNow() noexcept;
[[nodiscard]] bool CanCreateFileSystemCurl() noexcept;
#if defined(_DEBUG)
[[nodiscard]] HRESULT RunDebugCurlRuntimeProbe() noexcept;
#endif
[[nodiscard]] HRESULT HResultFromCurl(CURLcode code) noexcept;

#ifdef ENABLE_TESTS
// Diagnostic wildcard/ephemeral bind only; never connects or changes socket policy.
[[nodiscard]] HRESULT ProbeCurlEphemeralBindForSelfTest(unsigned short& portOut) noexcept;
// One nonblocking connect to a fresh owned loopback listener; no protocol traffic.
[[nodiscard]] HRESULT ProbeCurlLoopbackConnectForSelfTest(bool bindFirst, unsigned short& localPortOut, const wchar_t*& stageOut) noexcept;
#endif
void ApplyCommonCurlOptions(CURL* curl, const ConnectionInfo& conn, const FileSystemOptions* options, bool forUpload) noexcept;

size_t CurlWriteToString(void* buffer, size_t size, size_t nitems, void* outstream) noexcept;
[[nodiscard]] std::string BuildUrl(const ConnectionInfo& conn, std::wstring_view pluginPath, bool forDirectory, bool forCommand) noexcept;

[[nodiscard]] std::string RemotePathForCommand(const ConnectionInfo& conn, std::wstring_view pluginPath) noexcept;
[[nodiscard]] HRESULT CurlPerformList(const ConnectionInfo& conn, std::wstring_view pluginPath, std::string& outListing) noexcept;
// Targeted existence/size probe (FTP SIZE, SFTP stat) that avoids listing the parent directory. Returns
// S_OK when the file exists; sizeKnownOut reports whether the server gave a byte count. Explicit endpoint
// unavailability returns ERROR_NOT_SUPPORTED. Other failures remain hard errors.
[[nodiscard]] HRESULT CurlProbeRemoteFileSize(const ConnectionInfo& conn, std::wstring_view pluginPath, uint64_t& sizeOut, bool& sizeKnownOut) noexcept;

struct CurlSourceSizeCommitment final
{
    uint64_t sizeBytes = 0u;
    bool known         = false;
};

// Prefer the targeted probe over listing metadata. Only explicit primitive
// unavailability may fall back to the listed value/unknown-size policy.
[[nodiscard]] HRESULT ResolveCurlSourceSizeCommitment(const ConnectionInfo& conn,
                                                      std::wstring_view pluginPath,
                                                      uint64_t listedSizeBytes,
                                                      bool listedSizeKnown,
                                                      CurlSourceSizeCommitment& commitmentOut) noexcept;
[[nodiscard]] HRESULT CurlPerformListAndParse(const ConnectionInfo& conn,
                                              std::wstring_view pluginPath,
                                              std::vector<FilesInformationCurl::Entry>& outEntries) noexcept;
[[nodiscard]] HRESULT CurlPerformQuote(const ConnectionInfo& conn, const std::vector<std::string>& commands) noexcept;

constexpr unsigned long kCallbackArenaBytes = 64u * 1024u;

[[nodiscard]] const wchar_t* CopyArenaString(FileSystemArena* arena, std::wstring_view text) noexcept;

struct FileOperationProgress
{
    FileOperationProgress() = default;

    FileOperationProgress(const FileOperationProgress&)            = delete;
    FileOperationProgress(FileOperationProgress&&)                 = delete;
    FileOperationProgress& operator=(const FileOperationProgress&) = delete;
    FileOperationProgress& operator=(FileOperationProgress&&)      = delete;

    static inline thread_local uint64_t tlsProgressStreamId = 0;

    struct ProgressStreamScope final
    {
        explicit ProgressStreamScope(uint64_t streamId) noexcept : _previous(std::exchange(FileOperationProgress::tlsProgressStreamId, streamId))
        {
        }

        ProgressStreamScope(const ProgressStreamScope&)            = delete;
        ProgressStreamScope& operator=(const ProgressStreamScope&) = delete;
        ProgressStreamScope(ProgressStreamScope&&)                 = delete;
        ProgressStreamScope& operator=(ProgressStreamScope&&)      = delete;

        ~ProgressStreamScope()
        {
            FileOperationProgress::tlsProgressStreamId = _previous;
        }

    private:
        uint64_t _previous = 0;
    };

    FileSystemOperation operation = FILESYSTEM_COPY;
    unsigned long totalItems      = 0;
    unsigned long completedItems  = 0;

    uint64_t completedBytes = 0;

    FileSystemOptions options{};
    IFileSystemCallback* callback = nullptr;
    void* cookie                  = nullptr;

    std::atomic_bool internalCancel{false};
    std::atomic<uint64_t> bandwidthLimitBytesPerSecond{0};

    ArenaOwner arenaOwner{};

    std::mutex callbackMutex{};

    HRESULT Initialize(FileSystemOperation op, unsigned long total, const FileSystemOptions* initialOptions, IFileSystemCallback* cb, void* ck) noexcept
    {
        operation  = op;
        totalItems = total;
        callback   = cb;
        cookie     = ck;

        options = {};
        if (! FileSystemOptionsHaveValidHeader(initialOptions))
        {
            return E_INVALIDARG;
        }
        if (initialOptions && initialOptions->moveMode != FILESYSTEM_MOVE_DEFAULT &&
            (op != FILESYSTEM_MOVE || initialOptions->moveMode != FILESYSTEM_MOVE_NATIVE_ONLY))
        {
            return E_INVALIDARG;
        }
        if (initialOptions)
        {
            options = *initialOptions;
        }
        options.sizeBytes = sizeof(FileSystemOptions);
        bandwidthLimitBytesPerSecond.store(options.bandwidthLimitBytesPerSecond, std::memory_order_release);

        if (callback)
        {
            return arenaOwner.Initialize(kCallbackArenaBytes);
        }

        return S_OK;
    }

    [[nodiscard]] HRESULT CheckCancel() noexcept
    {
        if (internalCancel.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        const HRESULT operationControlHr = FileSystemCheckOperationControl(&options);
        if (FAILED(operationControlHr))
        {
            if (operationControlHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                internalCancel.store(true, std::memory_order_release);
            }
            return operationControlHr;
        }

        if (! callback)
        {
            return S_OK;
        }

        std::scoped_lock lock(callbackMutex);
        if (internalCancel.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        BOOL cancel      = FALSE;
        const HRESULT hr = callback->FileSystemShouldCancel(&cancel, cookie);
        if (FAILED(hr))
        {
            return hr;
        }

        if (cancel)
        {
            internalCancel.store(true, std::memory_order_release);
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        return S_OK;
    }

    [[nodiscard]] HRESULT ReportProgress(uint64_t currentItemTotalBytes,
                                         uint64_t currentItemCompletedBytes,
                                         std::wstring_view currentSourcePath,
                                         std::wstring_view currentDestinationPath) noexcept
    {
        if (! callback)
        {
            return S_OK;
        }

        std::scoped_lock lock(callbackMutex);

        FileSystemArena* arena = arenaOwner.Get();
        arena->usedBytes       = 0;

        const wchar_t* source = CopyArenaString(arena, currentSourcePath);
        const wchar_t* dest   = CopyArenaString(arena, currentDestinationPath);
        if (! source && ! currentSourcePath.empty())
        {
            return E_OUTOFMEMORY;
        }
        if (! dest && ! currentDestinationPath.empty())
        {
            return E_OUTOFMEMORY;
        }

        const HRESULT hr = callback->FileSystemProgress(operation,
                                                        totalItems,
                                                        completedItems,
                                                        0,
                                                        completedBytes,
                                                        source,
                                                        dest,
                                                        currentItemTotalBytes,
                                                        currentItemCompletedBytes,
                                                        &options,
                                                        tlsProgressStreamId,
                                                        cookie);
        bandwidthLimitBytesPerSecond.store(options.bandwidthLimitBytesPerSecond, std::memory_order_release);
        return NormalizeCancellation(hr);
    }

    [[nodiscard]] HRESULT ReportProgressWithCompletedBytes(uint64_t overallCompletedBytes,
                                                           uint64_t currentItemTotalBytes,
                                                           uint64_t currentItemCompletedBytes,
                                                           std::wstring_view currentSourcePath,
                                                           std::wstring_view currentDestinationPath) noexcept
    {
        if (! callback)
        {
            completedBytes = overallCompletedBytes;
            return S_OK;
        }

        std::scoped_lock lock(callbackMutex);
        completedBytes = overallCompletedBytes;

        FileSystemArena* arena = arenaOwner.Get();
        arena->usedBytes       = 0;

        const wchar_t* source = CopyArenaString(arena, currentSourcePath);
        const wchar_t* dest   = CopyArenaString(arena, currentDestinationPath);
        if (! source && ! currentSourcePath.empty())
        {
            return E_OUTOFMEMORY;
        }
        if (! dest && ! currentDestinationPath.empty())
        {
            return E_OUTOFMEMORY;
        }

        const HRESULT hr = callback->FileSystemProgress(operation,
                                                        totalItems,
                                                        completedItems,
                                                        0,
                                                        completedBytes,
                                                        source,
                                                        dest,
                                                        currentItemTotalBytes,
                                                        currentItemCompletedBytes,
                                                        &options,
                                                        tlsProgressStreamId,
                                                        cookie);
        bandwidthLimitBytesPerSecond.store(options.bandwidthLimitBytesPerSecond, std::memory_order_release);
        return NormalizeCancellation(hr);
    }

    void SetCompletedItems(unsigned long value) noexcept
    {
        if (! callback)
        {
            completedItems = value;
            return;
        }

        std::scoped_lock lock(callbackMutex);
        completedItems = value;
    }

    [[nodiscard]] HRESULT ReportItemCompleted(unsigned long itemIndex,
                                              std::wstring_view sourcePath,
                                              std::wstring_view destinationPath,
                                              HRESULT status,
                                              bool sourceObservedWithoutMutation = false) noexcept
    {
        if (! callback)
        {
            return S_OK;
        }

        std::scoped_lock lock(callbackMutex);

        FileSystemArena* arena = arenaOwner.Get();
        arena->usedBytes       = 0;

        const wchar_t* source = CopyArenaString(arena, sourcePath);
        const wchar_t* dest   = CopyArenaString(arena, destinationPath);
        if (! source && ! sourcePath.empty())
        {
            return E_OUTOFMEMORY;
        }
        if (! dest && ! destinationPath.empty())
        {
            return E_OUTOFMEMORY;
        }

        // The final HRESULT cannot prove no commit: an earlier child or rollback rename may
        // already have changed the namespace. Only the selected item's explicit pre-mutation
        // evidence can report a retained original. Callers must clear it before any mutation
        // attempt, including owned staging/overwrite work; absent evidence stays unknown.
        FileSystemItemMutationResult mutation{};
        const FileSystemItemMutationResult* mutationResult = nullptr;
        if (operation == FILESYSTEM_DELETE || operation == FILESYSTEM_MOVE || operation == FILESYSTEM_RENAME)
        {
            mutation.sizeBytes = sizeof(mutation);
            if (SUCCEEDED(status))
            {
                mutation.outcomeKnown         = TRUE;
                mutation.mutationCommitted    = TRUE;
                mutation.originalStillPresent = FALSE;
                mutationResult                = &mutation;
            }
            else if (sourceObservedWithoutMutation)
            {
                mutation.outcomeKnown         = TRUE;
                mutation.mutationCommitted    = FALSE;
                mutation.originalStillPresent = TRUE;
                mutationResult                = &mutation;
            }
        }
        const HRESULT hr = callback->FileSystemItemCompleted(operation, itemIndex, source, dest, status, mutationResult, &options, cookie);
        return NormalizeCancellation(hr);
    }
};

struct TransferProgressContext
{
    FileOperationProgress* progress = nullptr;
    std::wstring_view sourcePath;
    std::wstring_view destinationPath;

    uint64_t baseCompletedBytes                   = 0;
    std::atomic<uint64_t>* concurrentOverallBytes = nullptr;
    uint64_t lastConcurrentWireDone               = 0;

    uint64_t itemTotalBytes = 0;
    bool isUpload           = false;

    bool scaleForCopy       = false;
    bool scaleForCopySecond = false; // upload phase

    unsigned long reportIntervalMs = 100;
    unsigned long cancelIntervalMs = 250;

    uint64_t lastReportedItemDone = 0;
    uint64_t lastReportedOverall  = 0;

    uint64_t lastThrottleBytes = 0;
    uint64_t throttleStartTick = 0;

    uint64_t lastCancelTick = 0;
    uint64_t lastReportTick = 0;

    HRESULT abortHr = S_OK;

    void Begin() noexcept
    {
        const uint64_t now   = GetTickCount64();
        throttleStartTick    = now;
        lastReportTick       = 0;
        lastCancelTick       = 0;
        lastReportedItemDone = 0;
        lastReportedOverall  = 0;
        lastThrottleBytes    = 0;
        abortHr              = S_OK;
    }
};

[[nodiscard]] HRESULT CurlDownloadToFile(
    const ConnectionInfo& conn,
    std::wstring_view pluginPath,
    HANDLE file,
    const FileSystemOptions* options,
    TransferProgressContext* progressCtx,
    std::optional<uint64_t> expectedSizeBytes = std::nullopt) noexcept;

[[nodiscard]] HRESULT CurlUploadFromFile(const ConnectionInfo& conn,
                                         std::wstring_view pluginPath,
                                         HANDLE file,
                                         uint64_t sizeBytes,
                                         const FileSystemOptions* options,
                                         TransferProgressContext* progressCtx) noexcept;

struct CurlWriterPublicationMetrics final
{
    uint64_t requestCount        = 0u;
    uint64_t stagedBytes         = 0u;
    uint64_t cleanupAttemptCount = 0u;
    uint64_t commitUs            = 0u;
};

enum class CurlCleanupDebtKind : uint32_t
{
    None                     = 0u,
    RetainedRollbackSibling  = 1u << 0u,
    RetainedStagingSibling   = 1u << 1u,
    PreservedDestinationTree = 1u << 2u,
};

[[nodiscard]] constexpr uint32_t CurlCleanupDebtMask(CurlCleanupDebtKind kind) noexcept
{
    return static_cast<uint32_t>(kind);
}

struct CurlPublicationResult final
{
    HRESULT primaryMutationHr = S_OK;
    HRESULT cleanupHr         = S_OK;
    HRESULT sourceDeletionHr  = S_OK;
    bool primaryCommitted          = false;
    bool immutableRollbackEligible = false;
    bool artifactsPreserved        = false;
    uint32_t cleanupDebtMask        = CurlCleanupDebtMask(CurlCleanupDebtKind::None);
    uint64_t cleanupDebtCount       = 0u;

    void RecordPrimaryFailure(HRESULT failureHr) noexcept
    {
        if (SUCCEEDED(primaryMutationHr) && FAILED(failureHr))
        {
            primaryMutationHr = failureHr;
        }
    }

    void RecordSourceDeletionFailure(HRESULT failureHr) noexcept
    {
        if (SUCCEEDED(sourceDeletionHr) && FAILED(failureHr))
        {
            sourceDeletionHr = failureHr;
        }
        if (FAILED(failureHr))
        {
            artifactsPreserved = true;
        }
    }

    void RecordCleanupDebt(CurlCleanupDebtKind kind, HRESULT failureHr, uint64_t count = 1u) noexcept
    {
        if (count == 0u)
        {
            return;
        }

        cleanupDebtMask |= CurlCleanupDebtMask(kind);
        cleanupDebtCount = cleanupDebtCount > (std::numeric_limits<uint64_t>::max)() - count
                               ? (std::numeric_limits<uint64_t>::max)()
                               : cleanupDebtCount + count;
        artifactsPreserved = true;
        if (SUCCEEDED(cleanupHr) && FAILED(failureHr))
        {
            cleanupHr = failureHr;
        }
    }

    void Merge(const CurlPublicationResult& other) noexcept
    {
        RecordPrimaryFailure(other.primaryMutationHr);
        RecordSourceDeletionFailure(other.sourceDeletionHr);
        if (other.cleanupDebtCount != 0u)
        {
            cleanupDebtMask |= other.cleanupDebtMask;
            cleanupDebtCount = cleanupDebtCount > (std::numeric_limits<uint64_t>::max)() - other.cleanupDebtCount
                                   ? (std::numeric_limits<uint64_t>::max)()
                                   : cleanupDebtCount + other.cleanupDebtCount;
            if (SUCCEEDED(cleanupHr) && FAILED(other.cleanupHr))
            {
                cleanupHr = other.cleanupHr;
            }
        }
        primaryCommitted          = primaryCommitted || other.primaryCommitted;
        immutableRollbackEligible = immutableRollbackEligible || other.immutableRollbackEligible;
        artifactsPreserved        = artifactsPreserved || other.artifactsPreserved;
    }

    [[nodiscard]] HRESULT OperationResult() const noexcept
    {
        if (FAILED(primaryMutationHr))
        {
            constexpr uint32_t kPartialMutationDebtMask = CurlCleanupDebtMask(CurlCleanupDebtKind::RetainedRollbackSibling) |
                                                           CurlCleanupDebtMask(CurlCleanupDebtKind::PreservedDestinationTree);
            return (cleanupDebtMask & kPartialMutationDebtMask) != 0u ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : primaryMutationHr;
        }
        if (FAILED(sourceDeletionHr))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }
        return S_OK;
    }
};

class CurlPublicationAccumulator final
{
public:
    CurlPublicationAccumulator()                                            = default;
    CurlPublicationAccumulator(const CurlPublicationAccumulator&)           = delete;
    CurlPublicationAccumulator(CurlPublicationAccumulator&&)                = delete;
    CurlPublicationAccumulator& operator=(const CurlPublicationAccumulator&) = delete;
    CurlPublicationAccumulator& operator=(CurlPublicationAccumulator&&)      = delete;

    void Merge(const CurlPublicationResult& result) noexcept
    {
        RecordFirstFailure(_primaryMutationHr, result.primaryMutationHr);
        RecordFirstFailure(_cleanupHr, result.cleanupHr);
        RecordFirstFailure(_sourceDeletionHr, result.sourceDeletionHr);
        _cleanupDebtMask.fetch_or(result.cleanupDebtMask, std::memory_order_acq_rel);
        SaturatingAdd(_cleanupDebtCount, result.cleanupDebtCount);
        if (result.primaryCommitted)
        {
            _primaryCommitted.store(true, std::memory_order_release);
        }
        if (result.immutableRollbackEligible)
        {
            _immutableRollbackEligible.store(true, std::memory_order_release);
        }
        if (result.artifactsPreserved)
        {
            _artifactsPreserved.store(true, std::memory_order_release);
        }
    }

    [[nodiscard]] CurlPublicationResult Snapshot() const noexcept
    {
        return CurlPublicationResult{
            .primaryMutationHr          = static_cast<HRESULT>(_primaryMutationHr.load(std::memory_order_acquire)),
            .cleanupHr                  = static_cast<HRESULT>(_cleanupHr.load(std::memory_order_acquire)),
            .sourceDeletionHr           = static_cast<HRESULT>(_sourceDeletionHr.load(std::memory_order_acquire)),
            .primaryCommitted           = _primaryCommitted.load(std::memory_order_acquire),
            .immutableRollbackEligible  = _immutableRollbackEligible.load(std::memory_order_acquire),
            .artifactsPreserved         = _artifactsPreserved.load(std::memory_order_acquire),
            .cleanupDebtMask            = _cleanupDebtMask.load(std::memory_order_acquire),
            .cleanupDebtCount           = _cleanupDebtCount.load(std::memory_order_acquire),
        };
    }

private:
    static void RecordFirstFailure(std::atomic<long>& target, HRESULT failureHr) noexcept
    {
        if (SUCCEEDED(failureHr))
        {
            return;
        }

        long expected = S_OK;
        static_cast<void>(target.compare_exchange_strong(expected, static_cast<long>(failureHr), std::memory_order_acq_rel));
    }

    static void SaturatingAdd(std::atomic<uint64_t>& target, uint64_t value) noexcept
    {
        uint64_t current = target.load(std::memory_order_acquire);
        for (;;)
        {
            const uint64_t next = current > (std::numeric_limits<uint64_t>::max)() - value ? (std::numeric_limits<uint64_t>::max)() : current + value;
            if (target.compare_exchange_weak(current, next, std::memory_order_acq_rel))
            {
                return;
            }
        }
    }

    std::atomic<long> _primaryMutationHr{S_OK};
    std::atomic<long> _cleanupHr{S_OK};
    std::atomic<long> _sourceDeletionHr{S_OK};
    std::atomic<bool> _primaryCommitted{false};
    std::atomic<bool> _immutableRollbackEligible{false};
    std::atomic<bool> _artifactsPreserved{false};
    std::atomic<uint32_t> _cleanupDebtMask{CurlCleanupDebtMask(CurlCleanupDebtKind::None)};
    std::atomic<uint64_t> _cleanupDebtCount{0u};
};

// R3-1: the occupant the host showed the user before it granted a replacement (zero = unknown).
// FTP/SFTP/SCP expose no object identity, so the last-write time reported by the same listing
// dialect is the strongest revalidation available before the server-side rename.
struct CurlReplaceExpectation final
{
    __int64 lastWriteTime = 0;
};

[[nodiscard]] CurlPublicationResult PublishCurlWriterTransaction(const ConnectionInfo& conn,
                                                                 std::wstring_view destinationPath,
                                                                 HANDLE file,
                                                                 uint64_t sizeBytes,
                                                                 bool allowOverwrite,
                                                                 const CurlReplaceExpectation* replaceExpectation,
                                                                 CurlWriterPublicationMetrics& metrics) noexcept;

inline constexpr Common::Files::DeleteOnCloseTemporaryFileOptions kCurlTemporaryFileOptions{
    .prefix             = L"rsc",
    .flagsAndAttributes = FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_SEQUENTIAL_SCAN,
};
[[nodiscard]] HRESULT ResetFilePointerToStart(HANDLE file) noexcept;
[[nodiscard]] HRESULT GetFileSizeBytes(HANDLE file, uint64_t& out) noexcept;

[[nodiscard]] HRESULT ImapDownloadMessageToFile(const ConnectionInfo& conn, std::wstring_view pluginPath, HANDLE file) noexcept;

// Which libcurl channel carries the server text an IMAP custom request needs.
//
// libcurl routes every IMAP response line through its "header" (CLIENTWRITE_INFO)
// channel. Its body channel receives LIST/SEARCH/STATUS rows and single-message FETCH
// data, but a custom `FETCH`/`UID FETCH` whose set contains ',' or ':' is classified
// as a listing (lib/imap.c `is_custom_fetch_listing`, libcurl >= 8.18; the pinned
// 8.21.0 behaves this way) and its `* N FETCH` rows are never written as body. A
// bare single UID instead enters the literal-download path, which truncates a row
// after its first `{N}` literal. Summary fetches therefore capture WireLines and
// always send listing-shaped sets.
enum class ImapResponseCapture : uint8_t
{
    Body,
    WireLines,
};

// Runs one IMAP custom command against `mailboxPath` ("/" means no SELECT) and returns the
// captured server text. `stopReasonOut` receives the latched control verdict when a
// cancellation/deadline callback aborted the transfer.
[[nodiscard]] HRESULT CurlPerformImapCustomRequest(const ConnectionInfo& conn,
                                                   std::wstring_view mailboxPath,
                                                   std::string_view request,
                                                   std::string& outResponse,
                                                   HRESULT* stopReasonOut      = nullptr,
                                                   ImapResponseCapture capture = ImapResponseCapture::Body) noexcept;

[[nodiscard]] HRESULT ReadDirectoryEntries(const ConnectionInfo& conn, std::wstring_view path, std::vector<FilesInformationCurl::Entry>& entries) noexcept;
struct CurlEntryLookupMetrics final
{
    uint64_t rows = 0u;
    uint64_t pathBytes = 0u;
    uint64_t metadataBytes = 0u;
};

[[nodiscard]] HRESULT GetEntryInfo(const ConnectionInfo& conn,
                                  std::wstring_view path,
                                  FilesInformationCurl::Entry& out,
                                  CurlEntryLookupMetrics* metrics = nullptr,
                                  std::function<HRESULT()> checkpoint = {}) noexcept;

// R0f-Curl operation-control plumbing. The mutation entry points and the shared scheduler keep the
// owning File Operations call's options on the current thread; ApplyCommonCurlOptions installs a
// progress callback that polls them, so every transfer started underneath (control commands
// included) returns within about a second of Cancel or a passed deadline.
[[nodiscard]] const FileSystemOptions* CurlCurrentOperationOptions() noexcept;

class CurlOperationOptionsScope final
{
public:
    explicit CurlOperationOptionsScope(const FileSystemOptions* options) noexcept;
    ~CurlOperationOptionsScope();

    CurlOperationOptionsScope(const CurlOperationOptionsScope&)            = delete;
    CurlOperationOptionsScope& operator=(const CurlOperationOptionsScope&) = delete;
    CurlOperationOptionsScope(CurlOperationOptionsScope&&)                 = delete;
    CurlOperationOptionsScope& operator=(CurlOperationOptionsScope&&)      = delete;

private:
    const FileSystemOptions* _previous = nullptr;
};

[[nodiscard]] long CurlLowSpeedTimeSeconds(unsigned long operationTimeoutMs) noexcept;
[[nodiscard]] uint32_t CurlProviderWatchdogTimeoutMs(unsigned long connectTimeoutMs, unsigned long operationTimeoutMs) noexcept;

[[nodiscard]] HRESULT RemoteMkdir(const ConnectionInfo& conn, std::wstring_view path) noexcept;
[[nodiscard]] HRESULT RemoteDeleteFile(const ConnectionInfo& conn, std::wstring_view path) noexcept;
[[nodiscard]] HRESULT RemoteRemoveDirectory(const ConnectionInfo& conn, std::wstring_view path) noexcept;
[[nodiscard]] HRESULT RemoteRename(const ConnectionInfo& conn, std::wstring_view sourcePath, std::wstring_view destinationPath) noexcept;

[[nodiscard]] HRESULT EnsureDirectoryExists(const ConnectionInfo& conn, std::wstring_view directoryPath) noexcept;
[[nodiscard]] HRESULT EnsureOverwriteTargetFile(const ConnectionInfo& conn, std::wstring_view destinationPath, bool allowOverwrite) noexcept;

#if defined(ENABLE_TESTS)
[[nodiscard]] HRESULT CreateFtpSelfTestInstance(IFileSystem** result, IHost* host = nullptr) noexcept;
void RunCurlShortSuccessfulUploadSelfTests(unsigned int& passed, unsigned int& failed) noexcept;
void RunCurlStalledControlCommandCancelSelfTests(unsigned int& passed, unsigned int& failed) noexcept;
void RunCurlStalledReaderCancelSelfTests(unsigned int& passed, unsigned int& failed) noexcept; // R0f-Curl-OR1
void RunCurlParallelWritersSelfTests(unsigned int& passed, unsigned int& failed) noexcept;     // R0f-Curl-OR2
void RunDebugCurlStreamingReaderContractSelfTests(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugCurlWriterOwnershipContractSelfTests(unsigned int& passed, unsigned int& failed) noexcept;
#endif
} // namespace FileSystemCurlInternal
