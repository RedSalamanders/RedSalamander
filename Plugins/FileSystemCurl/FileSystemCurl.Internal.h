#pragma once

#include "FileSystemCurl.h"
#include "PlugInterfaces/Host.h"

#include "Helpers.h"

#include <algorithm>
#include <array>
#include <atomic>
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
#include <string>
#include <string_view>
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
};

struct ResolvedLocation
{
    ConnectionInfo connection;
    std::wstring remotePath;
};

[[nodiscard]] bool HasFlag(FileSystemFlags flags, FileSystemFlags flag) noexcept;
[[nodiscard]] HRESULT NormalizeCancellation(HRESULT hr) noexcept;

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
[[nodiscard]] HRESULT HResultFromCurl(CURLcode code) noexcept;
void ApplyCommonCurlOptions(CURL* curl, const ConnectionInfo& conn, const FileSystemOptions* options, bool forUpload) noexcept;

size_t CurlWriteToString(void* buffer, size_t size, size_t nitems, void* outstream) noexcept;
[[nodiscard]] std::string BuildUrl(const ConnectionInfo& conn, std::wstring_view pluginPath, bool forDirectory, bool forCommand) noexcept;

[[nodiscard]] std::string RemotePathForCommand(const ConnectionInfo& conn, std::wstring_view pluginPath) noexcept;
[[nodiscard]] HRESULT CurlPerformList(const ConnectionInfo& conn, std::wstring_view pluginPath, std::string& outListing) noexcept;
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

    ArenaOwner arenaOwner{};

    std::mutex callbackMutex{};

    HRESULT Initialize(FileSystemOperation op, unsigned long total, const FileSystemOptions* initialOptions, IFileSystemCallback* cb, void* ck) noexcept
    {
        operation  = op;
        totalItems = total;
        callback   = cb;
        cookie     = ck;

        options = {};
        if (initialOptions)
        {
            options = *initialOptions;
        }

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

    [[nodiscard]] HRESULT ReportItemCompleted(unsigned long itemIndex, std::wstring_view sourcePath, std::wstring_view destinationPath, HRESULT status) noexcept
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

        const HRESULT hr = callback->FileSystemItemCompleted(operation, itemIndex, source, dest, status, &options, cookie);
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
    const ConnectionInfo& conn, std::wstring_view pluginPath, HANDLE file, const FileSystemOptions* options, TransferProgressContext* progressCtx) noexcept;

[[nodiscard]] HRESULT CurlUploadFromFile(const ConnectionInfo& conn,
                                         std::wstring_view pluginPath,
                                         HANDLE file,
                                         uint64_t sizeBytes,
                                         const FileSystemOptions* options,
                                         TransferProgressContext* progressCtx) noexcept;

[[nodiscard]] wil::unique_hfile CreateTemporaryDeleteOnCloseFile() noexcept;
[[nodiscard]] HRESULT ResetFilePointerToStart(HANDLE file) noexcept;
[[nodiscard]] HRESULT GetFileSizeBytes(HANDLE file, uint64_t& out) noexcept;

[[nodiscard]] HRESULT ImapDownloadMessageToFile(const ConnectionInfo& conn, std::wstring_view pluginPath, HANDLE file) noexcept;

[[nodiscard]] HRESULT ReadDirectoryEntries(const ConnectionInfo& conn, std::wstring_view path, std::vector<FilesInformationCurl::Entry>& entries) noexcept;
[[nodiscard]] HRESULT GetEntryInfo(const ConnectionInfo& conn, std::wstring_view path, FilesInformationCurl::Entry& out) noexcept;

[[nodiscard]] HRESULT RemoteMkdir(const ConnectionInfo& conn, std::wstring_view path) noexcept;
[[nodiscard]] HRESULT RemoteDeleteFile(const ConnectionInfo& conn, std::wstring_view path) noexcept;
[[nodiscard]] HRESULT RemoteRemoveDirectory(const ConnectionInfo& conn, std::wstring_view path) noexcept;
[[nodiscard]] HRESULT RemoteRename(const ConnectionInfo& conn, std::wstring_view sourcePath, std::wstring_view destinationPath) noexcept;

[[nodiscard]] HRESULT EnsureDirectoryExists(const ConnectionInfo& conn, std::wstring_view directoryPath) noexcept;
[[nodiscard]] HRESULT EnsureOverwriteTargetFile(const ConnectionInfo& conn, std::wstring_view destinationPath, bool allowOverwrite) noexcept;
} // namespace FileSystemCurlInternal
